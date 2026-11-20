const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const { ValidationError, processExcelBuffer } = require('../src/services/excelProcessor');
const { exportXlsx, exportCsv, exportSqlite } = require('../src/services/exporters');
const { parseImportBuffer } = require('../src/services/dataImporter');
const { LocalDatabase } = require('../src/services/localDb');

const SAMPLE_XLS_PATH = path.resolve(__dirname, '..', '02090000000000-3.xls');

function makeTempDir(prefix) {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

test('LocalDatabase stores sessions and persists dataset across close and reopen with SQL search', () => {
  const tmpDir = makeTempDir('kmng-localdb-test-');
  const dbPath = path.join(tmpDir, 'data', 'kmng_metro.sqlite');

  try {
    const db1 = new LocalDatabase({ dbPath, sessionTtlMs: 60 * 60 * 1000 });

    assert.equal(db1.getDataset().hasData, false);

    // Session lifecycle
    const { token } = db1.createSession();
    assert.equal(db1.validateSession(token), true);
    assert.equal(db1.validateSession('invalid-token'), false);

    // Save dataset
    const columns = ['NIP', 'NIP_BARU', 'NAMA_LENGKAP', 'SATKER_3'];
    const rows = [
      ['250108080', '199708312025212042', 'SHOLEHA, S.Pd.', 'Kantor Kementerian Agama Kota Metro'],
      ['150000001', '198001012005011001', 'BUDI_100% TEST', 'Kantor Kementerian Agama Kota Metro'],
    ];
    const saved = db1.replaceDataset({
      sourceFilename: 'sample.xlsx',
      sourceFormat: 'xlsx',
      columns,
      rows,
    });
    assert.equal(saved.rowCount, 2);

    // Reopen DB to verify persistence on disk
    db1.close();
    const db2 = new LocalDatabase({ dbPath, sessionTtlMs: 60 * 60 * 1000 });

    assert.equal(db2.validateSession(token), true);
    const loaded = db2.getDataset();
    assert.equal(loaded.hasData, true);
    assert.equal(loaded.totalRows, 2);
    assert.deepEqual(loaded.columns, columns);
    assert.deepEqual(loaded.rows, rows);

    // Literal search escaping for '_' and '%'
    const searchUnderscore = db2.getDataset({ search: 'BUDI_100%' });
    assert.equal(searchUnderscore.filteredRows, 1);
    assert.equal(searchUnderscore.rows[0][1], '198001012005011001');

    const searchNip = db2.getDataset({ search: '199708312025212042' });
    assert.equal(searchNip.filteredRows, 1);
    assert.equal(searchNip.rows[0][2], 'SHOLEHA, S.Pd.');

    db2.deleteSession(token);
    assert.equal(db2.validateSession(token), false);
    db2.close();
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('parseImportBuffer imports raw .xls and generated .xlsx, .csv, and .sqlite files identically without mutating data', () => {
  const xlsBuf = fs.readFileSync(SAMPLE_XLS_PATH);
  const filtered = processExcelBuffer(xlsBuf, '02090000000000-3.xls');

  const tmpDir = makeTempDir('kmng-importer-test-');
  try {
    const xlsxPath = path.join(tmpDir, 'filtered.xlsx');
    const csvPath = path.join(tmpDir, 'filtered.csv');
    const sqlitePath = path.join(tmpDir, 'filtered.sqlite');

    exportXlsx({
      filePath: xlsxPath,
      sheetName: filtered.sheetName,
      headers: filtered.keptColumns,
      rows: filtered.rows,
      rawRows: filtered.rawRows,
    });
    exportCsv({
      filePath: csvPath,
      headers: filtered.keptColumns,
      rows: filtered.rows,
    });
    exportSqlite({
      filePath: sqlitePath,
      tableName: 'pegawai',
      headers: filtered.keptColumns,
      rows: filtered.rows,
    });

    const fromXlsx = parseImportBuffer(fs.readFileSync(xlsxPath), 'filtered.xlsx');
    assert.equal(fromXlsx.sourceFormat, 'xlsx');
    assert.deepEqual(fromXlsx.columns, filtered.keptColumns);
    assert.deepEqual(fromXlsx.rows, filtered.rows);

    const fromCsv = parseImportBuffer(fs.readFileSync(csvPath), 'filtered.csv');
    assert.equal(fromCsv.sourceFormat, 'csv');
    assert.deepEqual(fromCsv.columns, filtered.keptColumns);
    assert.deepEqual(fromCsv.rows, filtered.rows);

    const fromSqlite = parseImportBuffer(fs.readFileSync(sqlitePath), 'filtered.sqlite');
    assert.equal(fromSqlite.sourceFormat, 'sqlite');
    assert.deepEqual(fromSqlite.columns, filtered.keptColumns);
    assert.deepEqual(fromSqlite.rows, filtered.rows);

    // Raw .xls import preserves all 8,066 rows and 36 columns
    const fromRawXls = parseImportBuffer(xlsBuf, '02090000000000-3.xls');
    assert.equal(fromRawXls.sourceFormat, 'xls');
    assert.equal(fromRawXls.columns.length, 36);
    assert.equal(fromRawXls.rows.length, 8066);

    // Validation errors
    assert.throws(
      () => parseImportBuffer(Buffer.from('not-sqlite'), 'broken.sqlite'),
      (err) => err instanceof ValidationError && /SQLite 3/i.test(err.message)
    );
    assert.throws(
      () => parseImportBuffer(Buffer.from('hello'), 'notes.txt'),
      (err) => err instanceof ValidationError && /Unsupported import/i.test(err.message)
    );
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

