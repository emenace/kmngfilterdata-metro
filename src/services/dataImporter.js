const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');
const XLSX = require('xlsx');
const { ValidationError, extractCell } = require('./excelProcessor');
const { quoteIdent } = require('./exporters');
const { parseCsv } = require('../utils/csv');

/**
 * Strip an optional leading iteration column ("NO") and normalize column names and row matrices.
 */
function normalizeColumnsAndRows(rawHeaders, rawRows) {
  if (!Array.isArray(rawHeaders) || rawHeaders.length === 0) {
    throw new ValidationError('No header columns were found in the imported file.');
  }

  const hasLeadingNoCol =
    rawHeaders.length > 1 && String(rawHeaders[0] || '').trim().toUpperCase() === 'NO';
  const offset = hasLeadingNoCol ? 1 : 0;
  const candidateHeaders = rawHeaders.slice(offset);

  if (candidateHeaders.length === 0) {
    throw new ValidationError('No data columns were found in the imported file.');
  }

  const seen = new Map();
  const columns = candidateHeaders.map((rawName, idx) => {
    let baseName = String(rawName === undefined || rawName === null ? '' : rawName).trim();
    if (!baseName) {
      baseName = `COLUMN_${idx + 1}`;
    }
    if (baseName.toLowerCase() === '_row_id') {
      baseName = 'ROW_ID_COL';
    }

    const key = baseName.toLowerCase();
    const count = seen.get(key) || 0;
    seen.set(key, count + 1);
    return count === 0 ? baseName : `${baseName}_${count + 1}`;
  });

  const rows = [];
  for (let r = 0; r < rawRows.length; r += 1) {
    const srcRow = rawRows[r] || [];
    const normalized = new Array(columns.length);
    let hasContent = false;

    for (let c = 0; c < columns.length; c += 1) {
      const val = srcRow[c + offset];
      const strVal = val === undefined || val === null ? '' : String(val);
      normalized[c] = strVal;
      if (strVal !== '') {
        hasContent = true;
      }
    }

    if (hasContent) {
      rows.push(normalized);
    }
  }

  return { columns, rows };
}

/**
 * Parse an .xls or .xlsx buffer into normalized columns and rows without mutating cell values.
 */
function parseExcelImportBuffer(buffer, originalFilename) {
  let workbook;
  try {
    workbook = XLSX.read(buffer, {
      type: 'buffer',
      raw: true,
      cellDates: false,
      cellNF: true,
      cellStyles: false,
      cellFormula: false,
    });
  } catch (err) {
    throw new ValidationError(`Failed to read Excel file "${originalFilename}": ${err.message}`);
  }

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new ValidationError('The uploaded Excel file does not contain any worksheets.');
  }

  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws['!ref']) {
    throw new ValidationError(`Worksheet "${sheetName}" is empty.`);
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  const headerCols = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    headerCols.push({
      colIndex: c,
      name: extractCell(ws, range.s.r, c).canonical,
    });
  }

  while (headerCols.length > 0 && headerCols[headerCols.length - 1].name === '') {
    const lastCol = headerCols[headerCols.length - 1].colIndex;
    let hasAnyValue = false;
    for (let r = range.s.r + 1; r <= range.e.r; r += 1) {
      if (extractCell(ws, r, lastCol).canonical !== '') {
        hasAnyValue = true;
        break;
      }
    }
    if (hasAnyValue) {
      break;
    }
    headerCols.pop();
  }

  const rawHeaders = headerCols.map((col) => col.name);
  const rawRows = [];
  for (let r = range.s.r + 1; r <= range.e.r; r += 1) {
    const row = new Array(headerCols.length);
    for (let c = 0; c < headerCols.length; c += 1) {
      row[c] = extractCell(ws, r, headerCols[c].colIndex).canonical;
    }
    rawRows.push(row);
  }

  return normalizeColumnsAndRows(rawHeaders, rawRows);
}

/**
 * Parse a UTF-8 CSV buffer (with optional BOM) into normalized columns and rows.
 */
function parseCsvImportBuffer(buffer, originalFilename) {
  let parsed;
  try {
    parsed = parseCsv(buffer.toString('utf8'));
  } catch (err) {
    throw new ValidationError(`Failed to parse CSV file "${originalFilename}": ${err.message}`);
  }

  if (parsed.length === 0) {
    throw new ValidationError('The uploaded CSV file is empty.');
  }

  return normalizeColumnsAndRows(parsed[0], parsed.slice(1));
}

/**
 * Parse an uploaded SQLite 3 database buffer into normalized columns and rows.
 */
function parseSqliteImportBuffer(buffer, originalFilename) {
  if (
    buffer.length < 16 ||
    buffer.subarray(0, 16).toString('ascii') !== 'SQLite format 3\0'
  ) {
    throw new ValidationError(`File "${originalFilename}" is not a valid SQLite 3 database.`);
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'kmng-import-sqlite-'));
  const tempDbPath = path.join(tempDir, 'import.sqlite');

  let db = null;
  try {
    fs.writeFileSync(tempDbPath, buffer);
    db = new DatabaseSync(tempDbPath, { readOnly: true });

    const tableRows = db
      .prepare(
        `SELECT name FROM sqlite_master
         WHERE type = 'table'
           AND name NOT LIKE 'sqlite_%'
           AND name NOT LIKE '\\_%' ESCAPE '\\'
         ORDER BY rowid ASC;`
      )
      .all();

    if (tableRows.length === 0) {
      throw new ValidationError('The uploaded SQLite file does not contain any data tables.');
    }

    const tableNames = tableRows.map((row) => row.name);
    const preferredTable =
      tableNames.find((name) => name.toLowerCase() === 'pegawai') ||
      tableNames.find((name) => name.toLowerCase() === 'pegawai_data') ||
      tableNames[0];

    const safeTable = quoteIdent(preferredTable);
    const colInfo = db.prepare(`PRAGMA table_info(${safeTable});`).all();
    const rawHeaders = colInfo
      .map((col) => col.name)
      .filter((name) => name.toLowerCase() !== '_row_id');

    if (rawHeaders.length === 0) {
      throw new ValidationError(`SQLite table "${preferredTable}" contains no readable columns.`);
    }

    const selectCols = rawHeaders.map((col) => quoteIdent(col)).join(', ');
    const dbRecords = db
      .prepare(`SELECT ${selectCols} FROM ${safeTable} ORDER BY rowid ASC;`)
      .all();

    const rawRows = dbRecords.map((rec) =>
      rawHeaders.map((col) => (rec[col] === undefined || rec[col] === null ? '' : String(rec[col])))
    );

    return normalizeColumnsAndRows(rawHeaders, rawRows);
  } catch (err) {
    if (err instanceof ValidationError) {
      throw err;
    }
    throw new ValidationError(`Failed to read SQLite file "${originalFilename}": ${err.message}`);
  } finally {
    if (db) {
      try {
        db.close();
      } catch {
        // Ignore close errors during cleanup.
      }
    }
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

/**
 * Parse an uploaded .xls, .xlsx, .csv, or .sqlite/.db buffer into normalized { sourceFormat, columns, rows }.
 */
function parseImportBuffer(buffer, originalFilename) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new ValidationError('Uploaded import file is empty.');
  }

  const ext = path.extname(originalFilename || '').toLowerCase();
  if (ext === '.xls' || ext === '.xlsx') {
    const parsed = parseExcelImportBuffer(buffer, originalFilename);
    return { sourceFormat: ext.slice(1), ...parsed };
  }

  if (ext === '.csv') {
    const parsed = parseCsvImportBuffer(buffer, originalFilename);
    return { sourceFormat: 'csv', ...parsed };
  }

  if (ext === '.sqlite' || ext === '.db') {
    const parsed = parseSqliteImportBuffer(buffer, originalFilename);
    return { sourceFormat: 'sqlite', ...parsed };
  }

  throw new ValidationError(
    `Unsupported import file extension "${ext || '(none)'}". Allowed formats: .xls, .xlsx, .csv, .sqlite, .db.`
  );
}

module.exports = {
  normalizeColumnsAndRows,
  parseExcelImportBuffer,
  parseCsvImportBuffer,
  parseSqliteImportBuffer,
  parseImportBuffer,
};

