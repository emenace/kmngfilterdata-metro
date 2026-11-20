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

const FALLBACK_DUK_HTML = `<!DOCTYPE html>
<html>
<head><title>DUK-SISDM Kementerian Agama</title></head>
<body>
  <span id="lblSatker">KANTOR KEMENTERIAN AGAMA KOTA METRO</span>
  <table>
    <tr>
      <td rowspan="2">NO</td><td rowspan="2">NAMA</td><td colspan="2">PANGKAT</td>
      <td colspan="2">JABATAN</td><td colspan="2">MASA KERJA</td><td colspan="3">PENDIDIKAN TERAKHIR</td>
      <td rowspan="2">TGL<br/>LAHIR</td><td rowspan="2">TMT PENSIUN</td><td rowspan="2">KET.</td>
    </tr>
    <tr>
      <td>GOL/ RUANG</td><td>T M T</td><td>NAMA</td><td>T M T</td><td>THN</td><td>BLN</td><td>NAMA</td><td>THN</td><td>JML JAM</td>
    </tr>
    <tr><td>1</td><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td><td>8</td><td>9</td><td>10</td><td>11</td><td>12</td><td>13</td><td>14</td></tr>
    <tr>
      <td>1</td><td>SHOLEHA, S.Pd.<br/>250108080 / 199708312025212042</td><td>IX</td><td>01-10-2025</td>
      <td>Penata Layanan Operasional pada MAN 1 Metro Kota Metro Provinsi Lampung</td><td>01-10-2025</td><td>1</td><td>1</td>
      <td>Univ</td><td>2021</td><td>S-1/Sarjana</td><td>31-08-1997</td><td>01-09-2055</td><td></td>
    </tr>
    <tr>
      <td>2</td><td>PRAKAS SETIAWAN S.Kom.<br/>PW0006265 / 199809162025211045</td><td>LAIN-LAIN</td><td>01-10-2025</td>
      <td>Penata Layanan Operasional (Jabatan Pelaksana) pada Kantor Kementerian Agama Kota Metro Provinsi Lampung (PPPK PARUH WAKTU)</td>
      <td>01-10-2025</td><td>1</td><td>1</td><td>STMIK</td><td>2021</td><td>S-1/Sarjana</td><td>16-09-1998</td><td>01-10-2056</td><td></td>
    </tr>
    <tr>
      <td>3</td><td>MUCHAMAD ATOQ ILLAH <br/>PW0001210 / 197611182025211029</td><td>LAIN-LAIN</td><td>01-10-2025</td>
      <td>Operator Layanan Operasional (Jabatan Pelaksana) pada Kantor Kementerian Agama Kota Metro Provinsi Lampung (PPPK PARUH WAKTU)</td>
      <td>01-10-2025</td><td>1</td><td>1</td><td>-</td><td>1995</td><td>SLTA/SMA SEDERAJAT</td><td>18-11-1976</td><td>01-12-2034</td><td></td>
    </tr>
  </table>
</body>
</html>`;

async function startServer(app) {
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  return { server, baseUrl: `http://127.0.0.1:${port}` };
}

async function stopServer(app, server) {
  app.locals.jobStore.stopCleanupScheduler();
  await new Promise((resolve) => server.close(resolve));
  app.locals.localDb.close();
}

test('HTTP auth, dashboard, multi-format import into persistent SQLite, generator, and restart persistence', async (t) => {
  const rootTemp = fs.mkdtempSync(path.join(os.tmpdir(), 'kmng-http-test-'));
  const tmpDir = path.join(rootTemp, 'jobs');
  const dbPath = path.join(rootTemp, 'data', 'kmng_metro.sqlite');
  const appPassword = 'test-metro-password';

  let app1 = createApp({ tmpDir, dbPath, APP_PASSWORD: appPassword });
  let { server: server1, baseUrl } = await startServer(app1);
  let activeServer = server1;
  let activeApp = app1;

  t.after(async () => {
    await stopServer(activeApp, activeServer);
    fs.rmSync(rootTemp, { recursive: true, force: true });
  });

  // 1. Public vs protected endpoints before login
  const healthRes = await fetch(`${baseUrl}/api/health`);
  assert.equal(healthRes.status, 200);
  assert.equal((await healthRes.json()).status, 'ok');

  const loginPageRes = await fetch(`${baseUrl}/login`);
  assert.equal(loginPageRes.status, 200);
  assert.match(await loginPageRes.text(), /Unlock App/);

  const unauthHomeRes = await fetch(`${baseUrl}/`, { redirect: 'manual' });
  assert.equal(unauthHomeRes.status, 302);
  assert.equal(unauthHomeRes.headers.get('location'), '/login');

  const unauthGenRes = await fetch(`${baseUrl}/generate`, { redirect: 'manual' });
  assert.equal(unauthGenRes.status, 302);
  assert.equal(unauthGenRes.headers.get('location'), '/login');

  assert.equal((await fetch(`${baseUrl}/api/data`)).status, 401);
  assert.equal((await fetch(`${baseUrl}/api/process`, { method: 'POST' })).status, 401);
  assert.equal((await fetch(`${baseUrl}/api/import`, { method: 'POST' })).status, 401);

  // 2. Password-only login flow
  const missingPwRes = await fetch(`${baseUrl}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({}),
  });
  assert.equal(missingPwRes.status, 400);

  const wrongPwRes = await fetch(`${baseUrl}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ password: 'wrong-password' }),
  });
  assert.equal(wrongPwRes.status, 401);

  const loginRes = await fetch(`${baseUrl}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ password: appPassword }),
  });
  assert.equal(loginRes.status, 200);
  const setCookie = loginRes.headers.get('set-cookie') || '';
  assert.match(setCookie, /kmng_session=/);
  assert.match(setCookie, /HttpOnly/i);
  assert.match(setCookie, /Max-Age=18000/i);
  assert.match(setCookie, /Expires=/i);
  const sessionCookie = setCookie.split(';')[0];

  const authHeaders = { Cookie: sessionCookie };

  const statusRes = await fetch(`${baseUrl}/api/auth/status`, { headers: authHeaders });
  assert.equal((await statusRes.json()).authenticated, true);

  // Logged-in visit to /login redirects to /
  const loggedInLoginPage = await fetch(`${baseUrl}/login`, { headers: authHeaders, redirect: 'manual' });
  assert.equal(loggedInLoginPage.status, 302);
  assert.equal(loggedInLoginPage.headers.get('location'), '/');

  // Authenticated Dashboard & Generate pages
  const dashHtmlRes = await fetch(`${baseUrl}/`, { headers: authHeaders });
  assert.equal(dashHtmlRes.status, 200);
  const dashHtml = await dashHtmlRes.text();
  assert.match(dashHtml, /Import/);
  assert.match(dashHtml, /Generate/);

  const genHtmlRes = await fetch(`${baseUrl}/generate`, { headers: authHeaders });
  assert.equal(genHtmlRes.status, 200);
  assert.match(await genHtmlRes.text(), /Add DUK HTML support file/);

  // Initial empty SQLite dashboard state
  const emptyDataRes = await fetch(`${baseUrl}/api/data`, { headers: authHeaders });
  assert.equal(emptyDataRes.status, 200);
  const emptyData = await emptyDataRes.json();
  assert.equal(emptyData.hasData, false);
  assert.equal(emptyData.totalRows, 0);

  // 3. Generate base-only filtered outputs
  const fileBuf = fs.readFileSync(SAMPLE_XLS_PATH);
  const form = new FormData();
  form.append('file', new Blob([fileBuf]), '02090000000000-3.xls');

  const procRes = await fetch(`${baseUrl}/api/process`, {
    method: 'POST',
    headers: authHeaders,
    body: form,
  });
  assert.equal(procRes.status, 200);
  const data = await procRes.json();

  assert.equal(data.stats.filteredRowCount, 360);
  assert.equal(data.integrity.verified, true);

  const xlsxRes = await fetch(`${baseUrl}${data.downloads.xlsx.url}`, { headers: authHeaders });
  assert.equal(xlsxRes.status, 200);
  const xlsxBuf = Buffer.from(await xlsxRes.arrayBuffer());
  const xlsxWb = XLSX.read(xlsxBuf, { type: 'buffer', raw: true });
  assert.equal(xlsxWb.Sheets[xlsxWb.SheetNames[0]].A1.v, 'NO');
  assert.equal(xlsxWb.Sheets[xlsxWb.SheetNames[0]].A361.v, 360);

  const csvRes = await fetch(`${baseUrl}${data.downloads.csv.url}`, { headers: authHeaders });
  assert.equal(csvRes.status, 200);
  const csvBuf = Buffer.from(await csvRes.arrayBuffer());
  assert.equal(csvBuf.toString('utf8').startsWith('\uFEFFNO,NIP,NIP_BARU,NAMA_LENGKAP,'), true);

  // 4. Generate Base Excel + DUK HTML combined outputs
  const hasRealDuk = fs.existsSync(SAMPLE_DUK_PATH);
  const dukBuf = hasRealDuk
    ? fs.readFileSync(SAMPLE_DUK_PATH)
    : Buffer.from(FALLBACK_DUK_HTML, 'utf8');
  const expectedAdded = hasRealDuk ? 13 : 2;
  const expectedCombinedCount = 360 + expectedAdded;

  const combinedForm = new FormData();
  combinedForm.append('file', new Blob([fileBuf]), '02090000000000-3.xls');
  combinedForm.append('includeDuk', 'true');
  combinedForm.append('dukFile', new Blob([dukBuf]), 'duk.html');

  const combRes = await fetch(`${baseUrl}/api/process`, {
    method: 'POST',
    headers: authHeaders,
    body: combinedForm,
  });
  assert.equal(combRes.status, 200);
  const comb = await combRes.json();
  assert.equal(comb.stats.filteredRowCount, expectedCombinedCount);
  assert.equal(comb.integrity.verified, true);

  const combSqliteRes = await fetch(`${baseUrl}${comb.downloads.sqlite.url}`, { headers: authHeaders });
  assert.equal(combSqliteRes.status, 200);
  const combSqliteBuf = Buffer.from(await combSqliteRes.arrayBuffer());
  assert.equal(combSqliteBuf.subarray(0, 16).toString('ascii'), 'SQLite format 3\0');

  // 5. Import XLSX -> CSV -> SQLite into persistent local SQLite database
  const importXlsxForm = new FormData();
  importXlsxForm.append('file', new Blob([xlsxBuf]), data.downloads.xlsx.filename);
  const impXlsxRes = await fetch(`${baseUrl}/api/import`, {
    method: 'POST',
    headers: authHeaders,
    body: importXlsxForm,
  });
  assert.equal(impXlsxRes.status, 200);
  assert.equal((await impXlsxRes.json()).totalRows, 360);

  const importCsvForm = new FormData();
  importCsvForm.append('file', new Blob([csvBuf]), data.downloads.csv.filename);
  const impCsvRes = await fetch(`${baseUrl}/api/import`, {
    method: 'POST',
    headers: authHeaders,
    body: importCsvForm,
  });
  assert.equal(impCsvRes.status, 200);
  assert.equal((await impCsvRes.json()).totalRows, 360);

  const importSqliteForm = new FormData();
  importSqliteForm.append('file', new Blob([combSqliteBuf]), comb.downloads.sqlite.filename);
  const impSqliteRes = await fetch(`${baseUrl}/api/import`, {
    method: 'POST',
    headers: authHeaders,
    body: importSqliteForm,
  });
  assert.equal(impSqliteRes.status, 200);
  const impSqliteJson = await impSqliteRes.json();
  assert.equal(impSqliteJson.totalRows, expectedCombinedCount);
  assert.equal(impSqliteJson.columns.length, 27);

  // Verify dashboard query & search against persisted SQLite
  const searchRes = await fetch(`${baseUrl}/api/data?q=SHOLEHA`, { headers: authHeaders });
  assert.equal(searchRes.status, 200);
  const searchData = await searchRes.json();
  assert.equal(searchData.filteredRows, 2);

  const searchNipRes = await fetch(`${baseUrl}/api/data?q=199708312025212042`, { headers: authHeaders });
  assert.equal(searchNipRes.status, 200);
  const searchNipData = await searchNipRes.json();
  assert.equal(searchNipData.filteredRows, 1);
  assert.equal(searchNipData.rows[0][1], '199708312025212042');

  // 6. Verify BOTH session and saved SQLite data survive server restart!
  await stopServer(app1, server1);
  const app2 = createApp({ tmpDir, dbPath, APP_PASSWORD: appPassword });
  const started2 = await startServer(app2);
  activeApp = app2;
  activeServer = started2.server;
  baseUrl = started2.baseUrl;

  const afterRestartDataRes = await fetch(`${baseUrl}/api/data`, { headers: authHeaders });
  assert.equal(afterRestartDataRes.status, 200);
  const afterRestartData = await afterRestartDataRes.json();
  assert.equal(afterRestartData.hasData, true);
  assert.equal(afterRestartData.totalRows, expectedCombinedCount);
  assert.equal(afterRestartData.meta.sourceFormat, 'sqlite');

  // 7. Validation errors & Logout
  const badImportForm = new FormData();
  badImportForm.append('file', new Blob(['hello']), 'invalid.txt');
  const badImportRes = await fetch(`${baseUrl}/api/import`, {
    method: 'POST',
    headers: authHeaders,
    body: badImportForm,
  });
  assert.equal(badImportRes.status, 400);

  const badFormatRes = await fetch(`${baseUrl}/api/download/${data.jobId}/pdf`, { headers: authHeaders });
  assert.equal(badFormatRes.status, 400);

  const logoutRes = await fetch(`${baseUrl}/api/auth/logout`, {
    method: 'POST',
    headers: authHeaders,
  });
  assert.equal(logoutRes.status, 200);

  const afterLogoutDataRes = await fetch(`${baseUrl}/api/data`, { headers: authHeaders });
  assert.equal(afterLogoutDataRes.status, 401);
});
