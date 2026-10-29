// 企業主檔「境外」標記檢查：單筆新增、set-foreign、匯入匯出、匯入預覽警示、TA 篩選、權限、稽核 ＋ 變異（CFOR_MUT=1）。埠 B2_PORT 預設 3469。
process.env.B2_PORT = process.env.B2_PORT || '3469';
const H = require('./_b2-harness');
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');
const XLSX = require('xlsx');

async function importXlsx(cookie, rows, dry) {
  const ws = XLSX.utils.aoa_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'S');
  const fd = new FormData(); fd.append('file', new Blob([XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })]), 'a.xlsx');
  const r = await fetch(H.BASE + '/api/companies/import' + (dry ? '?dryRun=1' : ''), { method: 'POST', headers: { Cookie: cookie }, body: fd });
  return { status: r.status, json: await r.json().catch(() => ({})) };
}

async function api(ok) {
  await H.fresh([H.mkUser('_s_adm', 'admin'), H.mkUser('_s_mkt', 'marketing'), H.mkUser('_s_u1', 'user', { bu: ['ERP'] })],
    { companies: [
      { id: 'T1', name: '國內有統編公司', taxId: '11111111', matchKey: 'tax:11111111', createdAt: '2026-05-01T00:00:00Z' },
      { id: 'T2', name: 'Existing Overseas Ltd', taxId: '', matchKey: 'name:existingoverseasltd', createdAt: '2026-05-01T00:00:00Z' },
    ], contacts: [
      { id: 'k1', companyId: 'T1', owner: '_s_u1', name: 'a', email: 'a@x.com' },
      { id: 'k2', companyId: 'T2', owner: '_s_u1', name: 'b', email: 'b@y.com' },
    ], opportunities: [], industries: ['製造業'] }, { env: {} });
  await H.loginAll(['_s_adm', '_s_mkt', '_s_u1']);
  const post = (u, url, b) => H.http(u, 'POST', url, b);

  let r = await post('_s_mkt', '/api/companies', { name: 'Foo Corp', taxId: '12345678', foreign: true });
  ok(r.status === 400, 'B1 境外同時填統編 → 400（' + r.status + '）');
  r = await post('_s_mkt', '/api/companies', { name: 'Foo Corp', foreign: true, region: '美國' });
  ok(r.status === 200, 'B2 行銷建立境外公司（' + r.status + '）');
  const fooId = r.json.id;
  r = await H.http('_s_mkt', 'GET', '/api/companies');
  ok((r.json.find(x => x.id === fooId) || {}).foreign === true && (r.json.find(x => x.id === 'T1') || {}).foreign !== true, 'B3 列表 foreign 旗標正確');
  r = await H.http('_s_mkt', 'GET', '/api/companies/' + fooId);
  ok(r.status === 200 && r.json.company && r.json.company.foreign === true, 'B4 明細帶 foreign');

  r = await post('_s_u1', '/api/companies/T2/set-foreign', { foreign: true });
  ok(r.status === 403, 'B5 業務不能設境外（' + r.status + '）');
  r = await post('_s_mkt', '/api/companies/T1/set-foreign', { foreign: true });
  ok(r.status === 400, 'B6 已有統編不能標境外（' + r.status + '）');
  r = await post('_s_mkt', '/api/companies/T2/set-foreign', { foreign: true });
  ok(r.status === 200 && r.json.foreign === true, 'B7 無統編公司可標境外');
  r = await post('_s_mkt', '/api/companies/nope/set-foreign', { foreign: true });
  ok(r.status === 404, 'B8 找不到 404');

  const lr = await H.http(null, 'POST', '/api/login', { username: '_s_mkt', password: H.PW });
  const cookie = lr.setCookie.map(c => c.split(';')[0]).join('; ');
  let im = await importXlsx(cookie, [['公司名稱', '統一編號', '境外'], ['Bar GmbH', '', '是'], ['無統編國內行號', '', '']], true);
  ok(im.status === 200 && im.json.noTaxId === 1, 'B9 匯入預覽：境外不列入無統編警示，只剩 1 筆（' + im.json.noTaxId + '）');
  im = await importXlsx(cookie, [['公司名稱', '統一編號', '境外'], ['Bar GmbH', '', '是'], ['Existing Overseas Ltd', '', '否'], ['Foo Corp', '', '']], false);
  ok(im.status === 200 && im.json.success, 'B10 匯入寫入');
  r = await H.http('_s_mkt', 'GET', '/api/companies');
  const by = n => r.json.find(x => x.name === n) || {};
  ok(by('Bar GmbH').foreign === true && !by('Existing Overseas Ltd').foreign && by('Foo Corp').foreign === true, 'B11 匯入：是→境外、否→取消、空白→不變');

  r = await H.http('_s_mkt', 'GET', '/api/companies/export');
  const wb = XLSX.read(r.buf, { type: 'buffer' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  const hd = rows[0] || [], ci = hd.indexOf('境外');
  ok(ci >= 0 && (rows.find(x => x[0] === 'Bar GmbH') || [])[ci] === '是' && (rows.find(x => x[0] === '國內有統編公司') || [])[ci] === '', 'B12 匯出含「境外」欄');
  r = await H.http('_s_mkt', 'GET', '/api/companies/import-template');
  const tw = XLSX.read(r.buf, { type: 'buffer' });
  const th = (XLSX.utils.sheet_to_json(tw.Sheets[tw.SheetNames[0]], { header: 1, defval: '' })[0]) || [];
  ok(th.includes('境外'), 'B13 匯入範本含「境外」欄');

  r = await post('_s_adm', '/api/companies/ta/query', { filters: { foreign: 'yes' }, opts: { requireEmail: false } });
  const nYes = r.status === 200 ? r.json.stats && r.json.stats.companies : -1;
  r = await post('_s_adm', '/api/companies/ta/query', { filters: { foreign: 'no' }, opts: { requireEmail: false } });
  const nNo = r.status === 200 ? r.json.stats && r.json.stats.companies : -1;
  ok(nYes === 2 && nNo === 2, 'B14 TA 篩選：僅境外 ' + nYes + '（應 2：Bar GmbH、Foo Corp）、僅國內 ' + nNo + '（應 2：T1、T2）');

  const logs = await H.getLogs('_s_adm');
  ok(logs.some(l => l.action === 'SET_COMPANY_FOREIGN') && logs.some(l => l.action === 'CREATE_COMPANY' && /境外/.test(l.detail || '')), 'B15 稽核：SET_COMPANY_FOREIGN、CREATE_COMPANY（含境外）');
  await H.stopServer();
}

async function run() {
  let fail = 0; const ok = (c, m) => { console.log((c ? 'OK  ' : 'FAIL') + ' ' + m); if (!c) fail++; };
  try { await api(ok); } catch (e) { console.log('FAIL 例外 ' + (e && e.stack || e)); fail++; try { await H.stopServer(); } catch { /* */ } }
  return fail;
}

const MUTS = [
  ['M1 境外仍列入無統編警示', 'server.js', 'if (!taxId && foreignSet !== true) {', 'if (!taxId) {'],
  ['M2 已有統編也可標境外', 'server.js', "if (foreign && /^\\d{8}$/.test(String(m.taxId || ''))) return res.status(400)", "if (false) return res.status(400)"],
  ['M3 TA 忽略境外條件', 'lib/companyTa.js', "if (f.foreign) { if ((f.foreign === 'yes') !== !!m.foreign) return false; }", ""],
  ['M4 set-foreign 不驗權限', 'server.js', "app.post('/api/companies/:id/set-foreign', requireAuth, (req, res) => {\n  if (!isAdminOrMarketing(req)) return", "app.post('/api/companies/:id/set-foreign', requireAuth, (req, res) => {\n  if (false) return"],
  ['M5 匯入「否」不取消', 'server.js', "if (foreignSet !== null) m.foreign = foreignSet;", "if (foreignSet === true) m.foreign = true;"],
];

(async () => {
  let fail = await run();
  if (process.env.CFOR_MUT === '1') {
    for (const [name, rel, from, to] of MUTS) {
      const file = path.join(ROOT, rel), src = fs.readFileSync(file, 'utf8');
      if (!src.includes(from)) { console.log('FAIL 變異無法套用：' + name); fail++; continue; }
      fs.writeFileSync(file, src.replace(from, () => to));
      let mf;
      try { console.log('--- 變異 ' + name + '（必須有失敗）---'); mf = await run(); } finally { fs.writeFileSync(file, src); }
      console.log((mf > 0 ? 'OK  ' : 'FAIL') + ' 變異被抓到：' + name + '（失敗 ' + mf + ' 項）'); if (mf === 0) fail++;
    }
  }
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('例外', e && e.message); process.exit(1); });
