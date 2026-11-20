const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');
const XLSX = require('xlsx');

const { createApp } = require('../src/app');

const SAMPLE_XLS_PATH = path.resolve(__dirname, '..', '02090000000000-3.xls');
const SAMPLE_DUK_PATH = path.resolve(__dirname, '..', 'November DUK-SISDM Kementerian Agama.html');

test('HTTP API end-to-end upload, optional DUK HTML merge, verification, 3-format download, and error handling', async (t) => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'kmng-http-test-'));
  const app = createApp({ tmpDir });
  const server = http.createServer(app);

  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  t.after(async () => {
    app.locals.jobStore.stopCleanupScheduler();
    await new Promise((resolve) => server.close(resolve));
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  // 1. Static UI & Health check
  const homeRes = await fetch(`${baseUrl}/`);
  assert.equal(homeRes.status, 200);
  const homeHtml = await homeRes.text();
  assert.match(homeHtml, /Add DUK HTML support file/);

  const healthRes = await fetch(`${baseUrl}/api/health`);
  assert.equal(healthRes.status, 200);
  const healthJson = await healthRes.json();
  assert.equal(healthJson.status, 'ok');

  // 2. Upload base Excel file only
  const fileBuf = fs.readFileSync(SAMPLE_XLS_PATH);
  const form = new FormData();
  form.append('file', new Blob([fileBuf]), '02090000000000-3.xls');

  const procRes = await fetch(`${baseUrl}/api/process`, {
    method: 'POST',
    body: form,
  });
  assert.equal(procRes.status, 200);
  const data = await procRes.json();

  assert.equal(data.sourceFile, '02090000000000-3.xls');
  assert.equal(data.sheetName, 'Pegawai');
  assert.equal(data.tableName, 'pegawai');
  assert.equal(data.duk.enabled, false);
  assert.equal(data.stats.totalSourceRows, 8066);
  assert.equal(data.stats.baseExcelFilteredRows, 360);
  assert.equal(data.stats.addedFromDukCount, 0);
  assert.equal(data.stats.filteredRowCount, 360);
  assert.equal(data.stats.removedColumnCount, 9);
  assert.equal(data.stats.outputColumnCount, 27);
  assert.equal(data.integrity.verified, true);
  assert.equal(data.integrity.alteredCells, 0);
  assert.equal(data.preview.rows.length, 360);
  assert.equal(data.preview.totalRows, 360);

  // Download all 3 base outputs & check Column A ("NO": 1..360) in XLSX and CSV
  const xlsxRes = await fetch(`${baseUrl}${data.downloads.xlsx.url}`);
  assert.equal(xlsxRes.status, 200);
  assert.match(xlsxRes.headers.get('content-disposition') || '', /^attachment; filename=".*\.xlsx"$/);
  const xlsxBuf = Buffer.from(await xlsxRes.arrayBuffer());
  assert.equal(xlsxBuf.length, data.downloads.xlsx.bytes);
  assert.equal(xlsxBuf.subarray(0, 2).toString('ascii'), 'PK');
  const xlsxWb = XLSX.read(xlsxBuf, { type: 'buffer', raw: true });
  const xlsxWs = xlsxWb.Sheets[xlsxWb.SheetNames[0]];
  assert.equal(xlsxWs.A1.v, 'NO');
  assert.equal(xlsxWs.B1.v, 'NIP');
  assert.equal(xlsxWs.A2.v, 1);
  assert.equal(xlsxWs.A361.v, 360);

  const csvRes = await fetch(`${baseUrl}${data.downloads.csv.url}`);
  assert.equal(csvRes.status, 200);
  assert.match(csvRes.headers.get('content-disposition') || '', /^attachment; filename=".*\.csv"$/);
  const csvBuf = Buffer.from(await csvRes.arrayBuffer());
  assert.equal(csvBuf.length, data.downloads.csv.bytes);
  const csvText = csvBuf.toString('utf8');
  assert.equal(csvText.startsWith('\uFEFFNO,NIP,NIP_BARU,NAMA_LENGKAP,'), true);

  const sqliteRes = await fetch(`${baseUrl}${data.downloads.sqlite.url}`);
  assert.equal(sqliteRes.status, 200);
  assert.match(sqliteRes.headers.get('content-disposition') || '', /^attachment; filename=".*\.sqlite"$/);
  const sqliteBuf = Buffer.from(await sqliteRes.arrayBuffer());
  assert.equal(sqliteBuf.length, data.downloads.sqlite.bytes);
  assert.equal(sqliteBuf.subarray(0, 16).toString('ascii'), 'SQLite format 3\0');

  // 3. Upload Base Excel + DUK HTML support file combined
  const dukBuf = fs.readFileSync(SAMPLE_DUK_PATH);
  const combinedForm = new FormData();
  combinedForm.append('file', new Blob([fileBuf]), '02090000000000-3.xls');
  combinedForm.append('includeDuk', 'true');
  combinedForm.append('dukFile', new Blob([dukBuf]), 'November DUK-SISDM Kementerian Agama.html');

  const combRes = await fetch(`${baseUrl}/api/process`, {
    method: 'POST',
    body: combinedForm,
  });
  assert.equal(combRes.status, 200);
  const comb = await combRes.json();

  assert.equal(comb.stats.baseExcelFilteredRows, 360);
  assert.equal(comb.stats.addedFromDukCount, 13);
  assert.equal(comb.stats.filteredRowCount, 373);
  assert.equal(comb.duk.enabled, true);
  assert.equal(comb.duk.totalDukRows, 373);
  assert.equal(comb.duk.alreadyInExcelCount, 360);
  assert.equal(comb.duk.addedFromDukCount, 13);
  assert.equal(comb.duk.addedPeople.length, 13);
  assert.equal(comb.integrity.verified, true);
  assert.equal(comb.integrity.alteredCells, 0);
  assert.equal(comb.preview.rows.length, 373);
  assert.equal(comb.preview.totalRows, 373);

  const combXlsxRes = await fetch(`${baseUrl}${comb.downloads.xlsx.url}`);
  assert.equal(combXlsxRes.status, 200);
  const combXlsxBuf = Buffer.from(await combXlsxRes.arrayBuffer());
  const combWb = XLSX.read(combXlsxBuf, { type: 'buffer', raw: true });
  const combWs = combWb.Sheets[combWb.SheetNames[0]];
  assert.equal(combWs.A1.v, 'NO');
  assert.equal(combWs.AB1.v, 'MASA_KERJA_BULAN');
  assert.equal(combWs.A374.v, 373);
  assert.equal(combWs.C374.v, '197611182025211029');

  const combCsvRes = await fetch(`${baseUrl}${comb.downloads.csv.url}`);
  assert.equal(combCsvRes.status, 200);
  const combCsvText = Buffer.from(await combCsvRes.arrayBuffer()).toString('utf8');
  assert.equal(combCsvText.startsWith('\uFEFFNO,NIP,NIP_BARU,NAMA_LENGKAP,'), true);
  assert.equal(
    combCsvText.includes('\r\n373,PW0001210,197611182025211029,MUCHAMAD ATOQ ILLAH,'),
    true
  );

  const combSqliteRes = await fetch(`${baseUrl}${comb.downloads.sqlite.url}`);
  assert.equal(combSqliteRes.status, 200);
  const combSqliteBuf = Buffer.from(await combSqliteRes.arrayBuffer());
  assert.equal(combSqliteBuf.length, comb.downloads.sqlite.bytes);
  assert.equal(combSqliteBuf.subarray(0, 16).toString('ascii'), 'SQLite format 3\0');

  // 4. Error handling paths
  const emptyPostRes = await fetch(`${baseUrl}/api/process`, { method: 'POST' });
  assert.equal(emptyPostRes.status, 400);

  const badExtForm = new FormData();
  badExtForm.append('file', new Blob(['hello']), 'notes.txt');
  const badExtRes = await fetch(`${baseUrl}/api/process`, { method: 'POST', body: badExtForm });
  assert.equal(badExtRes.status, 400);

  const missingDukForm = new FormData();
  missingDukForm.append('file', new Blob([fileBuf]), '02090000000000-3.xls');
  missingDukForm.append('includeDuk', 'true');
  const missingDukRes = await fetch(`${baseUrl}/api/process`, { method: 'POST', body: missingDukForm });
  assert.equal(missingDukRes.status, 400);

  const badDukExtForm = new FormData();
  badDukExtForm.append('file', new Blob([fileBuf]), '02090000000000-3.xls');
  badDukExtForm.append('includeDuk', 'true');
  badDukExtForm.append('dukFile', new Blob(['hello']), 'bad.txt');
  const badDukExtRes = await fetch(`${baseUrl}/api/process`, { method: 'POST', body: badDukExtForm });
  assert.equal(badDukExtRes.status, 400);

  const invalidDukHtmlForm = new FormData();
  invalidDukHtmlForm.append('file', new Blob([fileBuf]), '02090000000000-3.xls');
  invalidDukHtmlForm.append('includeDuk', 'true');
  invalidDukHtmlForm.append('dukFile', new Blob(['<html><body>hello</body></html>']), 'notduk.html');
  const invalidDukHtmlRes = await fetch(`${baseUrl}/api/process`, { method: 'POST', body: invalidDukHtmlForm });
  assert.equal(invalidDukHtmlRes.status, 400);

  const badJobRes = await fetch(`${baseUrl}/api/download/not-a-uuid/csv`);
  assert.equal(badJobRes.status, 400);

  const missingJobRes = await fetch(`${baseUrl}/api/download/00000000-0000-4000-8000-000000000000/csv`);
  assert.equal(missingJobRes.status, 404);

  const badFormatRes = await fetch(`${baseUrl}/api/download/${data.jobId}/pdf`);
  assert.equal(badFormatRes.status, 400);
});
