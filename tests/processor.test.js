const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const XLSX = require('xlsx');

const { DEFAULT_CONFIG } = require('../src/config');
const { ValidationError, processExcelBuffer } = require('../src/services/excelProcessor');
const {
  parseDukHtmlBuffer,
  mergeDukIntoExcelResult,
} = require('../src/services/dukHtmlProcessor');
const {
  sanitizeFileBase,
  sanitizeTableName,
  exportXlsx,
  exportCsv,
  exportSqlite,
} = require('../src/services/exporters');
const {
  readBackXlsx,
  readBackCsv,
  verifyOutputs,
} = require('../src/services/integrityVerifier');

const SAMPLE_XLS_PATH = path.resolve(__dirname, '..', '02090000000000-3.xls');
const SAMPLE_DUK_PATH = path.resolve(__dirname, '..', 'November DUK-SISDM Kementerian Agama.html');

const FALLBACK_DUK_HTML = `<!DOCTYPE html>
<html>
<head><title>DUK-SISDM Kementerian Agama</title></head>
<body>
  <span id="lblSatker">KANTOR KEMENTERIAN AGAMA KOTA METRO</span>
  <table>
    <tr>
      <td rowspan="2">NO</td>
      <td rowspan="2">NAMA</td>
      <td colspan="2">PANGKAT</td>
      <td colspan="2">JABATAN</td>
      <td colspan="2">MASA KERJA</td>
      <td colspan="3">PENDIDIKAN TERAKHIR</td>
      <td rowspan="2">TGL<br/>LAHIR</td>
      <td rowspan="2">TMT PENSIUN</td>
      <td rowspan="2">KET.</td>
    </tr>
    <tr>
      <td>GOL/ RUANG</td><td>T M T</td><td>NAMA</td><td>T M T</td><td>THN</td><td>BLN</td><td>NAMA</td><td>THN</td><td>JML JAM</td>
    </tr>
    <tr><td>1</td><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td><td>8</td><td>9</td><td>10</td><td>11</td><td>12</td><td>13</td><td>14</td></tr>
    <tr>
      <td>1</td><td>Dra. ERLINA HARNIATI <br/>150261975 / 196703251993032001</td><td>IV/c</td><td>01-11-2026</td>
      <td>Guru Ahli Madya pada MAN 1 Metro Kota Metro Provinsi Lampung</td><td>01-04-2019</td><td>33</td><td>8</td>
      <td>Univ. Muhammadiyah</td><td>1991</td><td>S-1/Sarjana</td><td>25-03-1967</td><td>01-04-2027</td><td></td>
    </tr>
    <tr>
      <td>2</td><td>SHOLEHA, S.Pd.<br/>250108080 / 199708312025212042</td><td>IX</td><td>01-10-2025</td>
      <td>Penata Layanan Operasional pada MAN 1 Metro Kota Metro Provinsi Lampung</td><td>01-10-2025</td><td>1</td><td>1</td>
      <td>Univ</td><td>2021</td><td>S-1/Sarjana</td><td>31-08-1997</td><td>01-09-2055</td><td></td>
    </tr>
    <tr>
      <td>3</td><td>PRAKAS SETIAWAN S.Kom.<br/>PW0006265 / 199809162025211045</td><td>LAIN-LAIN</td><td>01-10-2025</td>
      <td>Penata Layanan Operasional (Jabatan Pelaksana) pada Kantor Kementerian Agama Kota Metro Provinsi Lampung (PPPK PARUH WAKTU)</td>
      <td>01-10-2025</td><td>1</td><td>1</td><td>Sekolah Tinggi Manajemen dan Informatika</td><td>2021</td><td>S-1/Sarjana</td><td>16-09-1998</td><td>01-10-2056</td><td></td>
    </tr>
    <tr>
      <td>4</td><td>MUCHAMAD ATOQ ILLAH <br/>PW0001210 / 197611182025211029</td><td>LAIN-LAIN</td><td>01-10-2025</td>
      <td>Operator Layanan Operasional (Jabatan Pelaksana) pada Kantor Kementerian Agama Kota Metro Provinsi Lampung (PPPK PARUH WAKTU)</td>
      <td>01-10-2025</td><td>1</td><td>1</td><td>-</td><td>1995</td><td>SLTA/SMA SEDERAJAT</td><td>18-11-1976</td><td>01-12-2034</td><td></td>
    </tr>
  </table>
</body>
</html>`;

function makeTempDir(prefix) {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

test('processes real provincial sample file 02090000000000-3.xls with NO column (1..360) in XLSX/CSV and 100% integrity', () => {
  assert.equal(fs.existsSync(SAMPLE_XLS_PATH), true, 'Sample .xls file must exist in workspace');

  const buffer = fs.readFileSync(SAMPLE_XLS_PATH);
  const result = processExcelBuffer(buffer, '02090000000000-3.xls');

  assert.equal(result.sheetName, 'Pegawai');
  assert.equal(result.filterColumn.name, 'SATKER_3');
  assert.equal(result.filterColumn.letter, 'X');
  assert.equal(result.totalSourceRows, 8066);
  assert.equal(result.filteredRowCount, 360);
  assert.equal(result.originalColumnCount, 36);
  assert.equal(result.removedColumns.length, 9);
  assert.equal(result.keptColumns.length, 27);

  for (const col of result.removedColumns) {
    assert.match(col, /^KODE_/i);
  }
  for (const col of result.keptColumns) {
    assert.doesNotMatch(col, /^KODE_/i);
  }

  const satkerIdx = result.keptColumns.indexOf('SATKER_3');
  assert.notEqual(satkerIdx, -1);
  for (const row of result.rows) {
    assert.equal(
      row[satkerIdx].toLowerCase().includes('kantor kementerian agama kota metro'),
      true
    );
  }

  // Spot-check critical data integrity cases from the real dataset.
  const nipBaruIdx = result.keptColumns.indexOf('NIP_BARU');
  const namaIdx = result.keptColumns.indexOf('NAMA_LENGKAP');
  const pangkatIdx = result.keptColumns.indexOf('PANGKAT');
  const tmtCpnsIdx = result.keptColumns.indexOf('TMT_CPNS');
  const tmtJabatanIdx = result.keptColumns.indexOf('TMT_JABATAN');
  const mkTahunIdx = result.keptColumns.indexOf('MASA_KERJA_TAHUN');
  const mkBulanIdx = result.keptColumns.indexOf('MASA_KERJA_BULAN');

  const sholehaRow = result.rows.find((r) => r[nipBaruIdx] === '199708312025212042');
  assert.ok(sholehaRow, 'Expected 18-digit NIP_BARU 199708312025212042 in Kota Metro slice');
  assert.equal(sholehaRow[namaIdx], 'SHOLEHA, S.Pd.');
  assert.equal(sholehaRow[pangkatIdx], '');
  assert.equal(sholehaRow[tmtCpnsIdx], '01-10-2025');
  assert.equal(sholehaRow[tmtJabatanIdx], '2025-10-01');

  assert.equal(result.rows.some((r) => r[mkTahunIdx] === '-2'), true);
  assert.equal(result.rows.some((r) => r[mkBulanIdx] === '-3'), true);

  const tmpDir = makeTempDir('kmng-real-test-');
  try {
    const fileBase = sanitizeFileBase('02090000000000-3.xls');
    const tableName = sanitizeTableName(result.sheetName);
    assert.equal(tableName, 'pegawai');

    const files = {
      xlsx: path.join(tmpDir, `${fileBase}.xlsx`),
      csv: path.join(tmpDir, `${fileBase}.csv`),
      sqlite: path.join(tmpDir, `${fileBase}.sqlite`),
    };

    exportXlsx({
      filePath: files.xlsx,
      sheetName: result.sheetName,
      headers: result.keptColumns,
      rows: result.rows,
      rawRows: result.rawRows,
    });
    exportCsv({
      filePath: files.csv,
      headers: result.keptColumns,
      rows: result.rows,
    });
    exportSqlite({
      filePath: files.sqlite,
      tableName,
      headers: result.keptColumns,
      rows: result.rows,
    });

    const report = verifyOutputs({
      expectedHeaders: result.keptColumns,
      expectedRows: result.rows,
      filterColumnName: result.filterColumn.name,
      targetSatker: result.targetSatker,
      removeColumnPattern: DEFAULT_CONFIG.REMOVE_COLUMN_PATTERN,
      files,
      tableName,
    });

    assert.equal(report.verified, true);
    assert.equal(report.alteredCells, 0);
    assert.equal(report.checkedRows, 360);
    assert.equal(report.checkedColumns, 27);
    assert.equal(report.formats.xlsx.columnCount, 28);
    assert.equal(report.formats.xlsx.hasRowNumberColumn, true);
    assert.equal(report.formats.csv.columnCount, 28);
    assert.equal(report.formats.csv.hasRowNumberColumn, true);
    assert.equal(report.formats.sqlite.columnCount, 27);
    assert.equal(report.formats.sqlite.hasRowNumberColumn, false);
    assert.equal(report.formats.xlsx.checksum, report.checksum);
    assert.equal(report.formats.csv.checksum, report.checksum);
    assert.equal(report.formats.sqlite.checksum, report.checksum);

    const xlsxData = readBackXlsx(files.xlsx);
    assert.equal(xlsxData.headers[0], 'NO');
    assert.equal(xlsxData.rows[0][0], '1');
    assert.equal(xlsxData.rows[359][0], '360');

    const csvData = readBackCsv(files.csv);
    assert.equal(csvData.headers[0], 'NO');
    assert.equal(csvData.rows[0][0], '1');
    assert.equal(csvData.rows[359][0], '360');
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('merges missing personnel from DUK-SISDM HTML with base Excel and preserves 100% integrity', () => {
  const hasRealDuk = fs.existsSync(SAMPLE_DUK_PATH);
  const dukBuffer = hasRealDuk
    ? fs.readFileSync(SAMPLE_DUK_PATH)
    : Buffer.from(FALLBACK_DUK_HTML, 'utf8');
  const dukFilename = hasRealDuk ? 'November DUK-SISDM Kementerian Agama.html' : 'duk-sample.html';

  const baseResult = processExcelBuffer(
    fs.readFileSync(SAMPLE_XLS_PATH),
    '02090000000000-3.xls'
  );
  const dukResult = parseDukHtmlBuffer(dukBuffer, dukFilename);

  const expectedDukTotal = hasRealDuk ? 373 : 4;
  const expectedAlready = hasRealDuk ? 360 : 2;
  const expectedAdded = hasRealDuk ? 13 : 2;
  const expectedCombinedCount = 360 + expectedAdded;

  assert.equal(dukResult.satkerTitle, 'KANTOR KEMENTERIAN AGAMA KOTA METRO');
  assert.equal(dukResult.totalDukRows, expectedDukTotal);

  const combined = mergeDukIntoExcelResult(baseResult, dukResult, { dukFilename });

  assert.equal(combined.baseExcelFilteredRows, 360);
  assert.equal(combined.filteredRowCount, expectedCombinedCount);
  assert.equal(combined.duk.enabled, true);
  assert.equal(combined.duk.alreadyInExcelCount, expectedAlready);
  assert.equal(combined.duk.addedFromDukCount, expectedAdded);
  assert.equal(combined.duk.addedPeople.length, expectedAdded);

  // First 360 rows must remain completely untouched compared to base Excel.
  assert.deepEqual(combined.rows.slice(0, 360), baseResult.rows);

  const cols = combined.keptColumns;
  const idx = (name) => cols.indexOf(name);

  // Spot-check first appended DUK person (#361 -> index 360) and last appended row.
  const firstAdded = combined.rows[360];
  assert.equal(firstAdded[idx('NIP')], 'PW0006265');
  assert.equal(firstAdded[idx('NIP_BARU')], '199809162025211045');
  assert.equal(firstAdded[idx('NAMA_LENGKAP')], 'PRAKAS SETIAWAN S.Kom.');
  assert.equal(firstAdded[idx('STATUS_PEGAWAI')], 'PPPK PARUH WAKTU');
  assert.equal(firstAdded[idx('TAMPIL_JABATAN')], 'Penata Layanan Operasional');
  assert.equal(firstAdded[idx('TIPE_JABATAN')], 'Jabatan Pelaksana');
  assert.equal(firstAdded[idx('TMT_JABATAN')], '2025-10-01');
  assert.equal(firstAdded[idx('TMT_PENSIUN')], '2056-10-01');
  assert.equal(firstAdded[idx('USIA_PENSIUN')], '58');
  assert.equal(firstAdded[idx('SATKER_3')], 'Kantor Kementerian Agama Kota Metro');
  assert.equal(firstAdded[idx('SATKER_4')], 'Kanwil Kementerian Agama Provinsi Lampung');

  const lastAdded = combined.rows[expectedCombinedCount - 1];
  assert.equal(lastAdded[idx('NIP')], 'PW0001210');
  assert.equal(lastAdded[idx('NIP_BARU')], '197611182025211029');
  assert.equal(lastAdded[idx('NAMA_LENGKAP')], 'MUCHAMAD ATOQ ILLAH');

  const tmpDir = makeTempDir('kmng-combined-test-');
  try {
    const fileBase = sanitizeFileBase('02090000000000-3.xls', 'kota_metro_combined');
    const tableName = sanitizeTableName(combined.sheetName);

    const files = {
      xlsx: path.join(tmpDir, `${fileBase}.xlsx`),
      csv: path.join(tmpDir, `${fileBase}.csv`),
      sqlite: path.join(tmpDir, `${fileBase}.sqlite`),
    };

    exportXlsx({
      filePath: files.xlsx,
      sheetName: combined.sheetName,
      headers: combined.keptColumns,
      rows: combined.rows,
      rawRows: combined.rawRows,
    });
    exportCsv({
      filePath: files.csv,
      headers: combined.keptColumns,
      rows: combined.rows,
    });
    exportSqlite({
      filePath: files.sqlite,
      tableName,
      headers: combined.keptColumns,
      rows: combined.rows,
    });

    const report = verifyOutputs({
      expectedHeaders: combined.keptColumns,
      expectedRows: combined.rows,
      filterColumnName: combined.filterColumn.name,
      targetSatker: combined.targetSatker,
      removeColumnPattern: DEFAULT_CONFIG.REMOVE_COLUMN_PATTERN,
      files,
      tableName,
    });

    assert.equal(report.verified, true);
    assert.equal(report.alteredCells, 0);
    assert.equal(report.checkedRows, expectedCombinedCount);
    assert.equal(report.checkedColumns, 27);
    assert.equal(report.formats.xlsx.columnCount, 28);
    assert.equal(report.formats.csv.columnCount, 28);
    assert.equal(report.formats.sqlite.columnCount, 27);

    const lastIdx = expectedCombinedCount - 1;
    const xlsxData = readBackXlsx(files.xlsx);
    assert.equal(xlsxData.rows[lastIdx][0], String(expectedCombinedCount));
    assert.equal(xlsxData.rows[lastIdx][2], '197611182025211029');

    const csvData = readBackCsv(files.csv);
    assert.equal(csvData.rows[lastIdx][0], String(expectedCombinedCount));
    assert.equal(csvData.rows[lastIdx][2], '197611182025211029');
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('DUK HTML validation rejects malformed HTML and deduplicates intra-DUK duplicates & maps PNS pangkat', () => {
  assert.throws(
    () => parseDukHtmlBuffer(Buffer.from('<html><body><p>Not a DUK table</p></body></html>'), 'bad.html'),
    (err) => err instanceof ValidationError && /DUK-SISDM/i.test(err.message)
  );

  const syntheticHtml = `
    <html><body>
      <table>
        <tr><td>NO</td><td>NAMA</td><td colspan="2">PANGKAT</td><td colspan="2">JABATAN</td><td colspan="2">MASA KERJA</td><td colspan="3">PENDIDIKAN TERAKHIR</td><td>TGL LAHIR</td><td>TMT PENSIUN</td><td>KET.</td></tr>
        <tr><td>1</td><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td><td>8</td><td>9</td><td>10</td><td>11</td><td>12</td><td>13</td><td>14</td></tr>
        <tr><td>1</td><td>BUDI SANTOSO<br/>150000001 / 198001012005011001</td><td>IV/c</td><td>01-10-2024</td><td>Analis SDM pada Kantor Kementerian Agama Kota Metro Provinsi Lampung</td><td>01-01-2020</td><td>19</td><td>0</td><td>UIN</td><td>2004</td><td>S-1/Sarjana</td><td>01-01-1980</td><td>01-02-2038</td><td></td></tr>
        <tr><td>2</td><td>BUDI DUPLICATE<br/>150000001 / 198001012005011001</td><td>IV/c</td><td>01-10-2024</td><td>Analis SDM pada Kantor Kementerian Agama Kota Metro Provinsi Lampung</td><td>01-01-2020</td><td>19</td><td>0</td><td>UIN</td><td>2004</td><td>S-1/Sarjana</td><td>01-01-1980</td><td>01-02-2038</td><td></td></tr>
      </table>
    </body></html>
  `;

  const ws = XLSX.utils.aoa_to_sheet([
    ['NIP', 'NIP_BARU', 'NAMA_LENGKAP', 'PANGKAT', 'GOL_RUANG', 'SATKER_3', 'STATUS_PEGAWAI'],
    ['150000002', '198501012010011002', 'SITI', 'Penata', 'III/c', 'Kantor Kementerian Agama Kota Metro', 'PNS'],
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Pegawai');
  const baseResult = processExcelBuffer(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), 'base.xlsx');

  const dukParsed = parseDukHtmlBuffer(Buffer.from(syntheticHtml), 'synthetic.html');
  const combined = mergeDukIntoExcelResult(baseResult, dukParsed, { dukFilename: 'synthetic.html' });

  assert.equal(combined.filteredRowCount, 2);
  assert.equal(combined.duk.addedFromDukCount, 1);
  assert.equal(combined.rows[1][combined.keptColumns.indexOf('PANGKAT')], 'Pembina Utama Muda');
  assert.equal(combined.rows[1][combined.keptColumns.indexOf('STATUS_PEGAWAI')], 'PNS');
});

test('handles synthetic .xlsx edge cases: leading zeros, quotes, newlines, trailing spaces, numbers, booleans, and case-insensitive matching', () => {
  const aoa = [
    ['NIP', 'kode_internal', 'NAMA', 'CATATAN', 'SATKER_3', 'ANGKA', 'AKTIF'],
    ['0012345', 'K01', 'Budi, S.Pd.', 'Say "Hello"\nLine2', 'Sub Kantor Kementerian Agama Kota Metro Timur', 42, true],
    ['0099999', 'K02', 'Andi ', 'Trailing space ', 'Kantor Kementerian Agama Kota Bandar Lampung', 10, false],
    ['0077777', 'K03', 'Siti', '', 'KANTOR KEMENTERIAN AGAMA KOTA METRO', 99, false],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data Pegawai');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  const result = processExcelBuffer(buf, 'edge_cases.xlsx');
  assert.equal(result.totalSourceRows, 3);
  assert.equal(result.filteredRowCount, 2);
  assert.deepEqual(result.removedColumns, ['kode_internal']);
  assert.deepEqual(result.keptColumns, ['NIP', 'NAMA', 'CATATAN', 'SATKER_3', 'ANGKA', 'AKTIF']);
  assert.equal(result.rows[0][0], '0012345');
  assert.equal(result.rows[0][2], 'Say "Hello"\nLine2');
  assert.equal(result.rows[1][2], '');

  const tmpDir = makeTempDir('kmng-edge-test-');
  try {
    const tableName = sanitizeTableName(result.sheetName);
    const files = {
      xlsx: path.join(tmpDir, 'out.xlsx'),
      csv: path.join(tmpDir, 'out.csv'),
      sqlite: path.join(tmpDir, 'out.sqlite'),
    };
    exportXlsx({ filePath: files.xlsx, sheetName: result.sheetName, headers: result.keptColumns, rows: result.rows, rawRows: result.rawRows });
    exportCsv({ filePath: files.csv, headers: result.keptColumns, rows: result.rows });
    exportSqlite({ filePath: files.sqlite, tableName, headers: result.keptColumns, rows: result.rows });

    const report = verifyOutputs({
      expectedHeaders: result.keptColumns,
      expectedRows: result.rows,
      filterColumnName: result.filterColumn.name,
      targetSatker: result.targetSatker,
      removeColumnPattern: DEFAULT_CONFIG.REMOVE_COLUMN_PATTERN,
      files,
      tableName,
    });
    assert.equal(report.verified, true);
    assert.equal(report.alteredCells, 0);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('handles zero matching rows cleanly and creates valid empty exports', () => {
  const ws = XLSX.utils.aoa_to_sheet([
    ['NIP', 'KODE_UNIT', 'NAMA_LENGKAP', 'SATKER_3'],
    ['123', '01', 'Test User', 'Kantor Kementerian Agama Kabupaten Lampung Selatan'],
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Pegawai');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  const result = processExcelBuffer(buf, 'zero_matches.xlsx');
  assert.equal(result.totalSourceRows, 1);
  assert.equal(result.filteredRowCount, 0);
  assert.deepEqual(result.keptColumns, ['NIP', 'NAMA_LENGKAP', 'SATKER_3']);

  const tmpDir = makeTempDir('kmng-empty-test-');
  try {
    const files = {
      xlsx: path.join(tmpDir, 'empty.xlsx'),
      csv: path.join(tmpDir, 'empty.csv'),
      sqlite: path.join(tmpDir, 'empty.sqlite'),
    };
    exportXlsx({ filePath: files.xlsx, sheetName: 'Pegawai', headers: result.keptColumns, rows: result.rows, rawRows: result.rawRows });
    exportCsv({ filePath: files.csv, headers: result.keptColumns, rows: result.rows });
    exportSqlite({ filePath: files.sqlite, tableName: 'pegawai', headers: result.keptColumns, rows: result.rows });

    const report = verifyOutputs({
      expectedHeaders: result.keptColumns,
      expectedRows: result.rows,
      filterColumnName: result.filterColumn.name,
      targetSatker: result.targetSatker,
      removeColumnPattern: DEFAULT_CONFIG.REMOVE_COLUMN_PATTERN,
      files,
      tableName: 'pegawai',
    });
    assert.equal(report.verified, true);
    assert.equal(report.checkedRows, 0);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('throws ValidationError when SATKER_3 column is missing', () => {
  const ws = XLSX.utils.aoa_to_sheet([
    ['NIP', 'NAMA_LENGKAP', 'UNIT'],
    ['123', 'Test', 'Kantor Kementerian Agama Kota Metro'],
  ]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  assert.throws(
    () => processExcelBuffer(buf, 'missing_col.xlsx'),
    (err) => err instanceof ValidationError && /SATKER_3/i.test(err.message)
  );
});
