// 企業主檔「疑似重複」檢查：單元（lib/companyDuplicates.js）＋ API（權限、清單、標記非重複、還原、稽核）＋ 變異（CDUP_MUT=1）。埠 B2_PORT 預設 3465。
process.env.B2_PORT = process.env.B2_PORT || '3465';
const H = require('./_b2-harness');
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');
const loose = s => String(s || '').replace(/\s+/g, '').replace(/股份有限公司|有限公司/g, '').toLowerCase();

const mk = (id, name, taxId, extra) => Object.assign({ id, name, taxId: taxId || '', createdAt: '2026-05-01T00:00:00Z' }, extra || {});
const sample = () => ({
  companies: [
    mk('A1', '甲科技', ''), mk('A2', '甲科技股份有限公司', '11111111'),           // 有統編 vs 無統編，同名 → 高
    mk('B1', '乙電子', ''), mk('B2', '乙電子有限公司', ''),                         // 都無統編、寫法不同 → 高
    mk('C1', '丙實業', '33333333'), mk('C2', '丙實業(分公司)', '33333333'),         // 同統編 → 高
    mk('D1', '丁資訊', '44444444'), mk('D2', '丁資訊股份有限公司', '55555555'),      // 同名不同統編 → 低
    mk('E1', '戊獨家', '66666666'),                                                  // 單獨 → 不列
    mk('F1', '已刪', '', { deleted: true }), mk('F2', '已刪股份有限公司', ''),       // 刪除者不算
  ],
  contacts: [
    { id: 'k1', companyId: 'A1', owner: 'u', name: 'x' }, { id: 'k2', companyId: 'A2', owner: 'u', name: 'y' }, { id: 'k3', companyId: 'A2', owner: 'u', isPlaceholder: true },
    { id: 'k4', companyId: 'A2', owner: 'u', deleted: true },
  ],
});

function unit(ok) {
  const D = require('../lib/companyDuplicates');
  const s = sample(); const r = D.run({ companies: s.companies, contacts: s.contacts, loose });
  const g = id => r.groups.find(x => x.members.some(m => m.id === id));
  ok(r.groups.length === 4, 'U1 共 4 組（甲、乙、丙、丁；單獨與已刪不列）→ ' + r.groups.length);
  ok(g('A1').level === 'high' && g('A1').reasons.includes('taxVsNoTax'), 'U2 一有統編一無 → 高信心 taxVsNoTax');
  ok(g('B1').level === 'high' && g('B1').reasons.includes('nameVariant'), 'U3 都無統編名稱變體 → 高信心 nameVariant');
  ok(g('C1').level === 'high' && g('C1').reasons.includes('sameTaxId') && g('C1').members.length === 2, 'U4 同統編 → 高信心 sameTaxId（名稱不同也連起來）');
  ok(g('D1').level === 'low' && g('D1').reasons.includes('diffTaxId'), 'U5 同名不同統編 → 低信心');
  ok(!g('E1') && !g('F1') && !g('F2'), 'U6 單獨主檔與已刪除主檔不列');
  const a = g('A1').members.find(m => m.id === 'A2');
  ok(a.contactCount === 2 && a.realContactCount === 1, 'U7 名片數：含待補 2、真實 1、已刪不算');
  ok(r.groups[0].level === 'high' && r.groups[r.groups.length - 1].level === 'low', 'U8 高信心排前、低信心排後');
  const key = g('A1').key;
  ok(D.run({ companies: s.companies, contacts: s.contacts, loose, ignores: { [key]: { by: 'x' } } }).groups.length === 3, 'U9 ignores 命中的組不列');
  ok(D.keyOf(['b', 'a', ' ']) === 'a|b', 'U10 keyOf 排序且去空');
  ok(D.run({ companies: [mk('Z1', '甲', '')], contacts: [], loose }).groups.length === 0 && D.run({ companies: [mk('Z1', 'a', ''), mk('Z2', 'a ', '')], contacts: [], loose }).groups.length === 0, 'U11 名稱過短（<2 字）不比對');
}

async function api(ok) {
  const s = sample();
  await H.fresh([H.mkUser('_s_adm', 'admin'), H.mkUser('_s_mkt', 'marketing'), H.mkUser('_s_u1', 'user', { bu: ['ERP'] })],
    { companies: s.companies, contacts: s.contacts, opportunities: [] }, { env: {} });
  await H.loginAll(['_s_adm', '_s_mkt', '_s_u1']);
  let r = await H.http('_s_mkt', 'GET', '/api/companies/duplicates');
  ok(r.status === 200 && r.json.groups.length === 4 && r.json.summary.open === 4 && r.json.summary.high === 3 && r.json.summary.low === 1, 'A1 行銷取得清單：4 組（高 3、低 1）（' + r.status + '）');
  r = await H.http('_s_adm', 'GET', '/api/companies/duplicates');
  ok(r.status === 200, 'A2 管理員可看');
  r = await H.http('_s_u1', 'GET', '/api/companies/duplicates');
  ok(r.status === 403, 'A3 業務 403（' + r.status + '）');
  r = await H.http('_s_u1', 'POST', '/api/companies/duplicates/dismiss', { ids: ['A1', 'A2'] });
  ok(r.status === 403, 'A4 業務不能標記（' + r.status + '）');
  r = await H.http('_s_mkt', 'POST', '/api/companies/duplicates/dismiss', { ids: ['A1', 'A2'] });
  ok(r.status === 200, 'A5 行銷標記非重複（' + r.status + '）');
  r = await H.http('_s_mkt', 'GET', '/api/companies/duplicates');
  ok(r.json.summary.open === 3 && r.json.summary.dismissed === 1 && r.json.groups.find(g => g.dismissed).members.length === 2, 'A6 標記後 open 3、dismissed 1');
  r = await H.http('_s_mkt', 'POST', '/api/companies/duplicates/dismiss', { ids: ['A1', 'E1'] });
  ok(r.status === 404, 'A7 不是現有疑似組的 id 組合 404，不能亂寫（' + r.status + '）');
  r = await H.http('_s_mkt', 'POST', '/api/companies/duplicates/dismiss', { ids: ['A1', 'A2'], undo: true });
  r = await H.http('_s_mkt', 'GET', '/api/companies/duplicates');
  ok(r.json.summary.open === 4 && r.json.summary.dismissed === 0, 'A8 還原後 open 4');
  const logs = await H.getLogs('_s_adm');
  ok(logs.some(l => l.action === 'DISMISS_COMPANY_DUP') && logs.some(l => l.action === 'RESTORE_COMPANY_DUP'), 'A9 稽核有標記與還原紀錄');
  const raw = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
  ok(raw.companies.length === s.companies.length && raw.contacts.length === s.contacts.length, 'A10 只讀偵測：主檔與名片數量不變');
  await H.stopServer();
}

async function run() {
  let fail = 0; const ok = (c, m) => { console.log((c ? 'OK  ' : 'FAIL') + ' ' + m); if (!c) fail++; };
  try { unit(ok); await api(ok); } catch (e) { console.log('FAIL 例外 ' + (e && e.stack || e)); fail++; try { await H.stopServer(); } catch { /* */ } }
  return fail;
}

const MUTS = [
  ['M1 不比對統編', 'lib/companyDuplicates.js', "[...byName.values(), ...byTax.values()]", "[...byName.values()]"],
  ['M2 清單不驗權限', 'server.js', "app.get('/api/companies/duplicates', requireAuth, (req, res) => {\n  if (!isAdminOrMarketing(req)) return", "app.get('/api/companies/duplicates', requireAuth, (req, res) => {\n  if (false) return"],
  ['M3 已標記仍列出', 'server.js', "r.groups.forEach(g => { g.dismissed = !!ign[g.key];", "r.groups.forEach(g => { g.dismissed = false;"],
  ['M4 同名不同統編也判高信心', 'lib/companyDuplicates.js', "reasons.includes('nameVariant') ? 'high' : 'low'", "reasons.includes('nameVariant') || reasons.includes('diffTaxId') ? 'high' : 'low'"],
];

(async () => {
  let fail = await run();
  if (process.env.CDUP_MUT === '1') {
    for (const [name, rel, from, to] of MUTS) {
      const file = path.join(ROOT, rel), src = fs.readFileSync(file, 'utf8');
      if (!src.includes(from)) { console.log('FAIL 變異無法套用：' + name); fail++; continue; }
      fs.writeFileSync(file, src.replace(from, to));
      delete require.cache[require.resolve('../lib/companyDuplicates')];
      let mf;
      try { console.log('--- 變異 ' + name + '（必須有失敗）---'); mf = await run(); } finally { fs.writeFileSync(file, src); delete require.cache[require.resolve('../lib/companyDuplicates')]; }
      console.log((mf > 0 ? 'OK  ' : 'FAIL') + ' 變異被抓到：' + name + '（失敗 ' + mf + ' 項）'); if (mf === 0) fail++;
    }
  }
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('例外', e && e.message); process.exit(1); });
