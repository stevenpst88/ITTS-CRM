// 企業主檔「單筆新增」檢查：權限、驗證、重複擋下（EXISTS／SIMILAR＋force）、客戶池、稽核 ＋ 變異（CCR_MUT=1）。埠 B2_PORT 預設 3467。
process.env.B2_PORT = process.env.B2_PORT || '3467';
const H = require('./_b2-harness');
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');

async function api(ok) {
  await H.fresh([H.mkUser('_s_adm', 'admin'), H.mkUser('_s_mkt', 'marketing'), H.mkUser('_s_u1', 'user', { bu: ['ERP'] })],
    { companies: [{ id: 'X1', name: '既有科技股份有限公司', taxId: '', matchKey: 'name:既有科技股份有限公司', createdAt: '2026-05-01T00:00:00Z' }], contacts: [], opportunities: [], industries: ['製造業', '半導體'] }, { env: {} });
  await H.loginAll(['_s_adm', '_s_mkt', '_s_u1']);
  const post = (u, b) => H.http(u, 'POST', '/api/companies', b);
  let r = await post('_s_u1', { name: '業務想建', taxId: '12345678' });
  ok(r.status === 403, 'A1 業務 403（' + r.status + '）');
  r = await post('_s_mkt', { name: '' });
  ok(r.status === 400, 'A2 缺公司名稱 400');
  r = await post('_s_mkt', { name: '甲公司', taxId: '123' });
  ok(r.status === 400, 'A3 統編非 8 碼 400');
  r = await post('_s_mkt', { name: '甲公司', taxId: '12345678', industry: '不存在的產業' });
  ok(r.status === 400, 'A4 產業不在清單 400');
  r = await post('_s_mkt', { name: '甲公司', taxId: '12345678', industry: '製造業', region: '台北', customerCode: 'C001', toPool: true });
  ok(r.status === 200 && r.json.success && r.json.poolCreated === 1, 'A5 行銷建立成功並放入客戶池（' + r.status + '）');
  const id = r.json.id;
  r = await H.http('_s_mkt', 'GET', '/api/companies');
  const c = (r.json || []).find(x => x.id === id);
  ok(c && c.name === '甲公司' && c.taxId === '12345678' && c.industry === '製造業' && c.region === '台北' && c.customerCode === 'C001', 'A6 列表帶出欄位（含區域）');
  r = await post('_s_mkt', { name: '另一個名字', taxId: '12345678' });
  ok(r.status === 409 && r.json.code === 'EXISTS', 'A7 同統編 → 409 EXISTS（' + r.status + '）');
  r = await post('_s_mkt', { name: '既有科技有限公司', taxId: '' });
  ok(r.status === 409 && r.json.code === 'SIMILAR' && r.json.similar.length === 1, 'A8 同名變體 → 409 SIMILAR');
  r = await post('_s_mkt', { name: '既有科技有限公司', taxId: '', force: true });
  ok(r.status === 200, 'A9 force 後可建立（' + r.status + '）');
  r = await post('_s_adm', { name: '海外客戶 Inc', taxId: '' });
  ok(r.status === 200 && r.json.poolCreated === 0, 'A10 管理員建立無統編、不放客戶池');
  const logs = await H.getLogs('_s_adm');
  ok(logs.filter(l => l.action === 'CREATE_COMPANY').length === 3, 'A11 稽核 CREATE_COMPANY 3 筆');
  const raw = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
  ok(raw.companies.length === 4 && raw.contacts.filter(x => x.isPlaceholder && x.owner === '_pool').length === 1, 'A12 主檔 4 筆、客戶池待補名片 1 張');
  // 匯出：含區域、客戶代號，且與匯入欄名一致（可來回）
  const XLSX = require('xlsx');
  r = await H.http('_s_mkt', 'GET', '/api/companies/export');
  const wb = XLSX.read(r.buf, { type: 'buffer' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  const hd = rows[0] || [];
  const row = rows.find(x => x[0] === '甲公司') || [];
  ok(hd.includes('區域') && hd.includes('客戶代號') && row[hd.indexOf('區域')] === '台北' && row[hd.indexOf('客戶代號')] === 'C001', 'A13 匯出含區域與客戶代號（值正確）');
  // 匯入預覽：境外公司寫法差異（Inc / Inc.）要被列為疑似重複
  const lr = await H.http(null, 'POST', '/api/login', { username: '_s_mkt', password: H.PW });
  const cookie = lr.setCookie.map(c => c.split(';')[0]).join('; ');
  const ws2 = XLSX.utils.aoa_to_sheet([['公司名稱', '統一編號'], ['Apex Global Inc.', '']]);
  const wb2 = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb2, ws2, 'S');
  const fd = new FormData(); fd.append('file', new Blob([XLSX.write(wb2, { type: 'buffer', bookType: 'xlsx' })]), 'a.xlsx');
  r = await H.http('_s_adm', 'POST', '/api/companies', { name: 'Apex Global Inc', taxId: '' });
  ok(r.status === 200, 'A14 建立境外公司（無統編）');
  const imp = await fetch(H.BASE + '/api/companies/import?dryRun=1', { method: 'POST', headers: { Cookie: cookie }, body: fd });
  const ij = await imp.json();
  ok(imp.status === 200 && ij.duplicates === 1 && ij.created === 0, 'A15 匯入預覽：Inc 與 Inc. 判為疑似重複（' + imp.status + ' dup=' + ij.duplicates + '）');
  await H.stopServer();
}

async function run() {
  let fail = 0; const ok = (c, m) => { console.log((c ? 'OK  ' : 'FAIL') + ' ' + m); if (!c) fail++; };
  try { await api(ok); } catch (e) { console.log('FAIL 例外 ' + (e && e.stack || e)); fail++; try { await H.stopServer(); } catch { /* */ } }
  return fail;
}

const MUTS = [
  ['M1 不驗權限', 'server.js', "app.post('/api/companies', requireAuth, (req, res) => {\n  if (!isAdminOrMarketing(req)) return", "app.post('/api/companies', requireAuth, (req, res) => {\n  if (false) return"],
  ['M2 不擋同鍵重複', 'server.js', 'if (exist) return res.status(409)', 'if (false) return res.status(409)'],
  ['M3 不擋同名變體', 'server.js', "if (twins.length && b.force !== true)", "if (false)"],
  ['M5 匯入預覽仍用嚴格名稱', 'server.js', "const nn = normalizeCompanyLoose(name);\n      const twins = nn.length >= 2 ?", "const nn = normalizeCompanyName(name);\n      const twins = nn.length >= 2 ?"],
  ['M4 不驗統編格式', 'server.js', "if (taxId && !/^\\d{8}$/.test(taxId)) return res.status(400).json({ error: '統一編號須為", "if (false) return res.status(400).json({ error: '統一編號須為"],
];

(async () => {
  let fail = await run();
  if (process.env.CCR_MUT === '1') {
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
