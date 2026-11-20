const crypto = require('node:crypto');
const fs = require('node:fs');
const { DatabaseSync } = require('node:sqlite');
const XLSX = require('xlsx');
const { parseCsv } = require('../utils/csv');
const { extractCell } = require('./excelProcessor');
const { quoteIdent } = require('./exporters');

/**
 * Compute a deterministic SHA-256 digest over exact column headers and cell strings.
 * Uses explicit length framing so cell boundaries can never collide.
 */
function computeCanonicalHash(headers, rows) {
  const hash = crypto.createHash('sha256');
  hash.update(`HDR:${headers.length}\0`);
  for (let c = 0; c < headers.length; c += 1) {
    const h = String(headers[c]);
    hash.update(`${Buffer.byteLength(h)}:${h}\0`);
  }

  hash.update(`ROWS:${rows.length}\0`);
  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r];
    hash.update(`${row.length}\0`);
    for (let c = 0; c < row.length; c += 1) {
      const val = String(row[c]);
      hash.update(`${Buffer.byteLength(val)}:${val}\0`);
    }
  }

  return hash.digest('hex');
}

/**
 * Read back headers and canonical cell strings from an exported .xlsx file.
 */
function readBackXlsx(filePath) {
  const wb = XLSX.readFile(filePath, {
    raw: true,
    cellDates: false,
    cellNF: true,
    cellStyles: false,
    cellFormula: false,
  });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws['!ref']) {
    return { headers: [], rows: [] };
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    headers.push(extractCell(ws, range.s.r, c).canonical);
  }

  const rows = [];
  for (let r = range.s.r + 1; r <= range.e.r; r += 1) {
    const row = new Array(headers.length);
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      row[c - range.s.c] = extractCell(ws, r, c).canonical;
    }
    rows.push(row);
  }

  return { headers, rows };
}

/**
 * Read back headers and cell strings from an exported .csv file.
 */
function readBackCsv(filePath) {
  const text = fs.readFileSync(filePath, 'utf8');
  const parsed = parseCsv(text);
  if (parsed.length === 0) {
    return { headers: [], rows: [] };
  }
  return {
    headers: parsed[0],
    rows: parsed.slice(1),
  };
}

/**
 * Read back headers and cell strings from an exported .sqlite database.
 */
function readBackSqlite(filePath, tableName) {
  const db = new DatabaseSync(filePath, { readOnly: true });
  try {
    const safeTable = quoteIdent(tableName);
    const colInfo = db.prepare(`PRAGMA table_info(${safeTable});`).all();
    const headers = colInfo.map((col) => col.name);
    const selectCols = headers.map((h) => quoteIdent(h)).join(', ');
    const dbRows = db.prepare(`SELECT ${selectCols} FROM ${safeTable} ORDER BY rowid ASC;`).all();
    const rows = dbRows.map((record) =>
      headers.map((h) => (record[h] === null || record[h] === undefined ? '' : String(record[h])))
    );
    return { headers, rows };
  } finally {
    db.close();
  }
}

/**
 * Verify a single exported dataset matches expected headers, filter invariant,
 * optional sequential Column A ("NO": 1..N), and every cell value.
 */
function verifyFormatDataset({
  formatName,
  actual,
  expectedHeaders,
  expectedRows,
  expectedHash,
  filterColumnName,
  targetSatker,
  removeColumnPattern,
  expectRowNumberCol = false,
}) {
  let dataHeaders = actual.headers;
  const expectedColCount = expectRowNumberCol ? expectedHeaders.length + 1 : expectedHeaders.length;

  if (actual.headers.length !== expectedColCount) {
    throw new Error(
      `Integrity check failed for ${formatName}: header count mismatch (expected ${expectedColCount}, got ${actual.headers.length}).`
    );
  }

  if (expectRowNumberCol) {
    if (actual.headers[0] !== 'NO') {
      throw new Error(
        `Integrity check failed for ${formatName}: Column A must be "NO", got "${actual.headers[0]}".`
      );
    }
    dataHeaders = actual.headers.slice(1);
  }

  for (let c = 0; c < expectedHeaders.length; c += 1) {
    if (dataHeaders[c] !== expectedHeaders[c]) {
      throw new Error(
        `Integrity check failed for ${formatName}: header mismatch at column ${c} (expected "${expectedHeaders[c]}", got "${dataHeaders[c]}").`
      );
    }
    if (removeColumnPattern.test(dataHeaders[c].trim())) {
      throw new Error(
        `Integrity check failed for ${formatName}: forbidden column "${dataHeaders[c]}" remained in output.`
      );
    }
  }

  if (actual.rows.length !== expectedRows.length) {
    throw new Error(
      `Integrity check failed for ${formatName}: row count mismatch (expected ${expectedRows.length}, got ${actual.rows.length}).`
    );
  }

  const satkerColIdx = expectedHeaders.indexOf(filterColumnName);
  const needle = targetSatker.toLowerCase();
  const dataRowsForHash = new Array(actual.rows.length);

  for (let r = 0; r < expectedRows.length; r += 1) {
    const expectedRow = expectedRows[r];
    const actualRow = actual.rows[r];

    if (actualRow.length !== expectedColCount) {
      throw new Error(
        `Integrity check failed for ${formatName}: row ${r + 1} column count mismatch.`
      );
    }

    let dataRow = actualRow;
    if (expectRowNumberCol) {
      if (actualRow[0] !== String(r + 1)) {
        throw new Error(
          `Integrity check failed for ${formatName}: iteration number mismatch at row ${r + 1} (expected "${r + 1}", got "${actualRow[0]}").`
        );
      }
      dataRow = actualRow.slice(1);
    }

    if (satkerColIdx !== -1 && !dataRow[satkerColIdx].toLowerCase().includes(needle)) {
      throw new Error(
        `Integrity check failed for ${formatName}: row ${r + 1} does not match "${targetSatker}".`
      );
    }

    for (let c = 0; c < expectedHeaders.length; c += 1) {
      if (dataRow[c] !== expectedRow[c]) {
        throw new Error(
          `Integrity check failed for ${formatName}: cell mismatch at row ${r + 1}, column "${expectedHeaders[c]}".`
        );
      }
    }

    dataRowsForHash[r] = dataRow;
  }

  const actualHash = computeCanonicalHash(dataHeaders, dataRowsForHash);
  if (actualHash !== expectedHash) {
    throw new Error(
      `Integrity check failed for ${formatName}: SHA-256 checksum mismatch.`
    );
  }

  return actualHash;
}

/**
 * Verify all 3 generated output files (.xlsx, .csv, .sqlite) by reading them back from disk
 * and proving zero cells were altered compared to the filtered/combined source rows.
 */
function verifyOutputs({
  expectedHeaders,
  expectedRows,
  filterColumnName,
  targetSatker,
  removeColumnPattern,
  files,
  tableName,
}) {
  const expectedHash = computeCanonicalHash(expectedHeaders, expectedRows);

  const xlsxActual = readBackXlsx(files.xlsx);
  const xlsxHash = verifyFormatDataset({
    formatName: 'XLSX',
    actual: xlsxActual,
    expectedHeaders,
    expectedRows,
    expectedHash,
    filterColumnName,
    targetSatker,
    removeColumnPattern,
    expectRowNumberCol: true,
  });

  const csvActual = readBackCsv(files.csv);
  const csvHash = verifyFormatDataset({
    formatName: 'CSV',
    actual: csvActual,
    expectedHeaders,
    expectedRows,
    expectedHash,
    filterColumnName,
    targetSatker,
    removeColumnPattern,
    expectRowNumberCol: true,
  });

  const sqliteActual = readBackSqlite(files.sqlite, tableName);
  const sqliteHash = verifyFormatDataset({
    formatName: 'SQLite',
    actual: sqliteActual,
    expectedHeaders,
    expectedRows,
    expectedHash,
    filterColumnName,
    targetSatker,
    removeColumnPattern,
    expectRowNumberCol: false,
  });

  return {
    verified: true,
    algorithm: 'SHA-256',
    checksum: expectedHash,
    checkedRows: expectedRows.length,
    checkedColumns: expectedHeaders.length,
    alteredCells: 0,
    rowNumberFormats: ['xlsx', 'csv'],
    formats: {
      xlsx: {
        verified: true,
        checksum: xlsxHash,
        rowCount: xlsxActual.rows.length,
        columnCount: xlsxActual.headers.length,
        hasRowNumberColumn: true,
        bytes: fs.statSync(files.xlsx).size,
      },
      csv: {
        verified: true,
        checksum: csvHash,
        rowCount: csvActual.rows.length,
        columnCount: csvActual.headers.length,
        hasRowNumberColumn: true,
        bytes: fs.statSync(files.csv).size,
      },
      sqlite: {
        verified: true,
        checksum: sqliteHash,
        rowCount: sqliteActual.rows.length,
        columnCount: sqliteActual.headers.length,
        hasRowNumberColumn: false,
        bytes: fs.statSync(files.sqlite).size,
        tableName,
      },
    },
  };
}

module.exports = {
  computeCanonicalHash,
  readBackXlsx,
  readBackCsv,
  readBackSqlite,
  verifyOutputs,
};
