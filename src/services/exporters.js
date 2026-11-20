const fs = require('node:fs');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');
const XLSX = require('xlsx');
const { serializeCsv } = require('../utils/csv');

/**
 * Derive a safe base filename from an uploaded filename.
 */
function sanitizeFileBase(originalFilename, suffix = 'kota_metro_filtered') {
  const rawBase = path.basename(originalFilename || 'data', path.extname(originalFilename || ''));
  const cleaned = rawBase
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^[._-]+|[._-]+$/g, '')
    .slice(0, 80);
  const base = cleaned || 'data';
  return `${base}_${suffix}`;
}

/**
 * Derive a safe SQLite identifier from a worksheet name.
 */
function sanitizeTableName(sheetName) {
  let name = String(sheetName || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_]+/g, '_')
    .replace(/^_+|_+$/g, '');
  if (/^[0-9]/.test(name)) {
    name = `t_${name}`;
  }
  return name || 'filtered_data';
}

/**
 * Quote a SQLite identifier safely using double quotes.
 */
function quoteIdent(identifier) {
  return `"${String(identifier).replace(/"/g, '""')}"`;
}

/**
 * Sanitize an Excel worksheet name to meet Excel's 31-char / disallowed-char constraints.
 */
function sanitizeSheetName(sheetName) {
  const cleaned = String(sheetName || 'Filtered')
    .replace(/[:\\/?*[\]]/g, '_')
    .trim()
    .slice(0, 31);
  return cleaned || 'Filtered';
}

/**
 * Export filtered data as an .xlsx workbook with sequential iteration number Column A ("NO": 1..N)
 * and explicit string cells so long numeric identifiers (e.g. 18-digit NIP_BARU) are never altered.
 */
function exportXlsx({ filePath, sheetName, headers, rows, rawRows }) {
  const ws = {};
  const dataColCount = headers.length;
  const noWidth = Math.max(6, String(rows.length).length + 2);
  const dataColWidths = headers.map((h) => Math.max(String(h).length + 2, 10));

  ws[XLSX.utils.encode_cell({ r: 0, c: 0 })] = { t: 's', v: 'NO' };
  for (let c = 0; c < dataColCount; c += 1) {
    ws[XLSX.utils.encode_cell({ r: 0, c: c + 1 })] = { t: 's', v: headers[c] };
  }

  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r];
    const rawRow = rawRows && rawRows[r] ? rawRows[r] : null;

    ws[XLSX.utils.encode_cell({ r: r + 1, c: 0 })] = { t: 'n', v: r + 1 };

    for (let c = 0; c < dataColCount; c += 1) {
      const val = row[c] === undefined || row[c] === null ? '' : String(row[c]);
      const rawCell = rawRow ? rawRow[c] : null;
      const addr = XLSX.utils.encode_cell({ r: r + 1, c: c + 1 });

      if (rawCell && (rawCell.t === 'n' || rawCell.t === 'b')) {
        ws[addr] = { ...rawCell };
      } else {
        ws[addr] = { t: 's', v: val };
      }

      if (val.length + 2 > dataColWidths[c]) {
        dataColWidths[c] = Math.min(48, val.length + 2);
      }
    }
  }

  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: rows.length, c: dataColCount },
  });
  ws['!cols'] = [{ wch: noWidth }, ...dataColWidths.map((wch) => ({ wch }))];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(sheetName));
  XLSX.writeFile(wb, filePath, { bookType: 'xlsx', compression: true });
}

/**
 * Export filtered data as an RFC 4180 UTF-8 CSV file with BOM and sequential Column A ("NO": 1..N).
 */
function exportCsv({ filePath, headers, rows }) {
  const csvHeaders = ['NO', ...headers];
  const csvRows = rows.map((row, idx) => [String(idx + 1), ...row]);
  const content = serializeCsv(csvHeaders, csvRows);
  fs.writeFileSync(filePath, content, 'utf8');
}

/**
 * Export filtered data into a single self-contained SQLite database file.
 * All columns are stored as TEXT NOT NULL to prevent numeric precision loss or leading-zero stripping.
 */
function exportSqlite({ filePath, tableName, headers, rows }) {
  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }

  const safeTable = quoteIdent(tableName);
  const colDefs = headers.map((h) => `${quoteIdent(h)} TEXT NOT NULL`);
  const quotedCols = headers.map((h) => quoteIdent(h));
  const placeholders = headers.map(() => '?').join(', ');

  const db = new DatabaseSync(filePath);
  try {
    db.exec('PRAGMA journal_mode = DELETE;');
    db.exec(`CREATE TABLE ${safeTable} (${colDefs.join(', ')});`);

    const insertSql = `INSERT INTO ${safeTable} (${quotedCols.join(', ')}) VALUES (${placeholders});`;
    const stmt = db.prepare(insertSql);

    db.exec('BEGIN TRANSACTION;');
    try {
      for (let r = 0; r < rows.length; r += 1) {
        stmt.run(...rows[r]);
      }
      db.exec('COMMIT;');
    } catch (err) {
      db.exec('ROLLBACK;');
      throw err;
    }
  } finally {
    db.close();
  }
}

module.exports = {
  sanitizeFileBase,
  sanitizeTableName,
  sanitizeSheetName,
  quoteIdent,
  exportXlsx,
  exportCsv,
  exportSqlite,
};
