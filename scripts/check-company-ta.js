// 企業主檔「篩選 TA」檢查：單元（lib/companyTa.js）＋ API（範圍、權限、下載、稽核）＋ 變異（CTA_MUT=1）。埠 B2_PORT 預設 3463。
process.env.B2_PORT = process.env.B2_PORT || '3463';
const H = require('./_b2-harness');
const fs = require('fs'), path = require('path');
const XLSX = require('xlsx');
const ROOT = path.join(__dirname, '..');
const NONE = '__none__';

const iso = '2026-10-01T00:00:00.000Z';
const companies = () => [
  { id: 'A', name: '甲製造', taxId: '11111111', industry: '製造業', region: '北區', capital: 50000000, createdAt: '2026-01-10T00:00:00Z', gcisEnriched: true },
  { id: 'B', name: '乙金融', taxId: '22222222', industry: '金融業', region: '南區', capital: 2000000, createdAt: '2026-03-10T00:00:00Z' },
  { id: 'C', name: '丙新匯入', taxId: '', industry: '製造業', region: '', capital: null, createdAt: '2026-05-10T00:00:00Z' },
  { id: 'D', name: '丁待補', taxId: '44444444', industry: '', region: '北區', capital: 100000000, createdAt: '2026-06-10T00:00:00Z' },
];
const contacts = () => [
  { id: 'a1', owner: '_s_uerp', name: '王主要', title: '資訊經理', email: 'a@x.com', companyId: 'A', company: '甲製造', isPrimary: true, createdAt: iso, bu: 'ERP' },
  { id: 'a2', owner: '_s_uerp', name: '李重複', title: '採購', email: 'A@X.COM', companyId: 'A', company: '甲製造', createdAt: iso, bu: 'ERP' },
  { id: 'a3', owner: '_s_uerp', name: '陳離職', title: 'IT', email: 'gone@x.com', companyId: 'A', company: '甲製造', employmentStatus: 'resigned', createdAt: iso, bu: 'ERP' },
  { id: 'a4', owner: '_s_uerp', name: '趙無信', title: '會計', email: '', companyId: 'A', company: '甲製造', createdAt: iso, bu: 'ERP' },
  { id: 'b1', owner: '_s_uits', name: '孫金融', title: '協理', email: 'b@y.com', companyId: 'B', company: '乙金融', createdAt: iso, bu: 'ITS' },
  { id: 'p1', owner: '_pool', name: '（待補名片）', title: '', email: '', companyId: 'D', company: '丁待補', isPlaceholder: true, createdAt: iso },
];
const opps = () => [
  { id: 'o1', owner: '_s_uerp', contactId: 'a1', company: '甲製造', stage: 'Won', amount: 100 },
  { id: 'o2', owner: '_s_uits', contactId: 'b1', company: '乙金融', stage: 'B', amount: 50 },
];
const buUnion = (names, imp) => [...new Set([...(imp || []), ...names.map(n => ({ _s_uerp: 'ERP', _s_uits: 'ITS' }[n])).filter(Boolean)])];

function unit(ok) {
  const T = require('../lib/companyTa');
  const real = contacts().filter(c => c.owner !== '_pool');
  const go = (filters, opts) => T.run({ companies: companies(), contacts: real, opps: opps(), buUnion, filters, opts });
  const ids = r => r.companyIds.slice().sort().join('');
  ok(ids(go({})) === 'ABCD', 'U1 無條件＝全部公司');
  ok(ids(go({ industries: ['製造業'] })) === 'AC', 'U2 產業多選');
  ok(ids(go({ industries: ['金融業', NONE] })) === 'BD', 'U3 產業＋未分類（OR）');
  ok(ids(go({ capMinWan: 1000 })) === 'AD', 'U4 資本額下限 1000 萬（未填排除）');
  ok(ids(go({ capMinWan: 1000, capMaxWan: 6000 })) === 'A', 'U5 資本額區間');
  ok(ids(go({ capMinWan: 1000, capNone: 'include' })) === 'ACD', 'U6 區間＋納入未填');
  ok(ids(go({ capNone: 'only' })) === 'C', 'U7 只要未填');
  ok(ids(go({ regions: [NONE] })) === 'C', 'U8 區域未填');
  ok(ids(go({ industries: ['製造業'], regions: ['北區'] })) === 'A', 'U9 不同欄位 AND');
  ok(ids(go({ bus: ['ITS'] })) === 'B' && ids(go({ bus: ['ERP', 'ITS'] })) === 'AB', 'U10 BU');
  ok(ids(go({ statuses: ['won'] })) === 'A' && ids(go({ statuses: ['open'] })) === 'B', 'U11 客戶狀態 已成交／已開商機');
  ok(ids(go({ statuses: ['newImport'] })) === 'CD' && ids(go({ statuses: ['noOpp'] })) === 'CD', 'U12 新匯入／尚無商機');
  ok(ids(go({ hasTaxId: 'no' })) === 'C' && ids(go({ gcis: 'yes' })) === 'A', 'U13 統編／GCIS');
  ok(ids(go({ createdFrom: '2026-03-01', createdTo: '2026-05-31' })) === 'BC', 'U14 建立時間區間');
  ok(ids(go({ keyword: '金融' })) === 'B' && ids(go({ keyword: '1111' })) === 'A', 'U15 關鍵字（名稱／統編）');
  ok(ids(go({ responsible: ['x'] })) === '', 'U16 責任業務不符');
  const r = go({ industries: ['製造業'] });
  ok(r.stats.contacts === 1 && r.rows[0].contact.id === 'a1', 'U17 預設：必有 Email＋在職＋Email 去重（大小寫）→ 只剩王主要');
  ok(go({ industries: ['製造業'] }, { dedupeEmail: false }).stats.contacts === 2, 'U18 關閉去重 → 2 位');
  ok(go({ industries: ['製造業'] }, { activeOnly: false, dedupeEmail: false }).stats.contacts === 3, 'U19 含離職 → 3 位');
  ok(go({ industries: ['製造業'] }, { requireEmail: false, activeOnly: false, dedupeEmail: false }).stats.contacts === 4, 'U20 全放寬 → 4 位');
  ok(go({ industries: ['製造業'] }, { requireEmail: false, activeOnly: false, dedupeEmail: false, onePerCompany: true }).rows[0].contact.id === 'a1', 'U21 每家一位優先主要聯絡人');
  ok(go({ industries: ['製造業'] }, { titleKeyword: '採購', dedupeEmail: false }).rows.map(x => x.contact.id).join() === 'a2', 'U22 職稱關鍵字');
  const g = go({ industries: 'x', capMinWan: 'abc', statuses: ['zzz'], bus: null, hasTaxId: 'maybe' }, { requireEmail: 'x' });
  ok(ids(g) === 'ABCD', 'U23 畸形輸入被正規化（不拋例外、不誤篩）');
}

async function api(ok) {
  await H.fresh([
    H.mkUser('_s_adm', 'admin'), H.mkUser('_s_mkt', 'marketing', { canDownloadContacts: true }), H.mkUser('_s_mkt2', 'marketing', { canDownloadContacts: false }),
    H.mkUser('_s_uerp', 'user', { bu: ['ERP'], canDownloadContacts: true }), H.mkUser('_s_uits', 'user', { bu: ['ITS'], canDownloadContacts: false }),
  ], { companies: companies(), contacts: contacts(), opportunities: opps() }, { env: {} });
  await H.loginAll(['_s_adm', '_s_mkt', '_s_mkt2', '_s_uerp', '_s_uits']);
  const q = (who, filters, opts) => H.http(who, 'POST', '/api/companies/ta/query', { filters, opts });

  let r = await H.http('_s_mkt', 'GET', '/api/companies/ta/options');
  ok(r.status === 200 && r.json.total === 4 && r.json.industries.find(x => x.name === '製造業').n === 2 && r.json.canDownload === true, 'A1 options：行銷看全部 4 家、產業計數、可下載');
  r = await q('_s_mkt', { industries: ['製造業'] });
  ok(r.status === 200 && r.json.stats.companies === 2 && r.json.stats.contacts === 1 && r.json.preview.length === 1, 'A2 行銷查產業：2 家公司、1 位聯絡人、有預覽');
  r = await q('_s_mkt', { statuses: ['newImport'] }, { requireEmail: false });
  ok(r.json.stats.companies === 2 && r.json.stats.contacts === 0, 'A3 客戶池待補名片不算聯絡人（新匯入 2 家、聯絡人 0）');
  r = await q('_s_mkt2', {});
  ok(r.status === 200 && r.json.canDownload === false && r.json.preview.length === 0 && r.json.stats.companies === 4, 'A4 無下載權限：看得到數量、預覽為空');
  r = await H.http('_s_mkt2', 'POST', '/api/companies/ta/export', { filters: {} });
  ok(r.status === 403, 'A5 無下載權限：匯出 403（' + r.status + '）');
  r = await q('_s_uerp', {});
  ok(r.json.stats.companies === 1 && r.json.stats.contacts === 1, 'A6 業務只看自己名片的公司（ERP 業務：1 家）');
  r = await q('_s_uits', {});
  ok(r.json.stats.companies === 1 && r.json.preview.length === 0, 'A7 另一位業務只看 ITS 的 1 家，且無下載權限不給預覽');
  r = await H.http('_s_uerp', 'POST', '/api/companies/ta/export', { filters: { industries: ['金融業'] } });
  ok(r.status === 200 && XLSX.read(r.buf).Sheets['聯絡人'] && XLSX.utils.sheet_to_json(XLSX.read(r.buf).Sheets['聯絡人']).length === 0, 'A8 業務匯出別人的公司＝空（不能越權取得）');

  r = await H.http('_s_mkt', 'POST', '/api/companies/ta/export', { filters: { industries: ['製造業'] } });
  const wb = r.status === 200 ? XLSX.read(r.buf) : null;
  const rows = wb ? XLSX.utils.sheet_to_json(wb.Sheets['聯絡人']) : [];
  ok(r.status === 200 && /spreadsheetml/.test(r.headers.get('content-type')) && rows.length === 1 && rows[0]['Email'] === 'a@x.com' && rows[0]['公司名稱'] === '甲製造' && rows[0]['BU'] === 'ERP', 'A9 匯出聯絡人名單：1 筆、欄位正確');
  ok(!JSON.stringify(rows).includes('待補名片') && !JSON.stringify(rows).includes('gone@x.com'), 'A10 匯出不含客戶池待補名片與離職者');
  r = await H.http('_s_mkt', 'POST', '/api/companies/ta/export', { filters: { capMinWan: 1000 }, mode: 'companies' });
  const crow = r.status === 200 ? XLSX.utils.sheet_to_json(XLSX.read(r.buf).Sheets['公司']) : [];
  ok(crow.length === 2 && crow.map(x => x['公司名稱']).sort().join() === '丁待補,甲製造', 'A11 匯出公司清單（資本額 ≥1000 萬：甲、丁）');
  const logs = await H.getLogs('_s_adm');
  const lg = logs.find(l => l.action === 'EXPORT_TA');
  ok(!!lg && /資料匯出|篩選TA/.test(JSON.stringify(lg)) && !/a@x\.com/.test(JSON.stringify(lg)), 'A12 稽核 EXPORT_TA 有紀錄且不含 Email');
  r = await q('_s_mkt', { industries: 'x', capMinWan: 'abc', statuses: ['zzz'] }, { requireEmail: 'x' });
  ok(r.status === 200 && r.json.stats.companies === 4, 'A13 畸形輸入不當機');
  r = await H.http(null, 'POST', '/api/companies/ta/query', {});
  ok(r.status === 401 || r.status === 403, 'A14 未登入被擋（' + r.status + '）');
  await H.stopServer();
}

async function run() {
  let fail = 0; const ok = (c, m) => { console.log((c ? 'OK  ' : 'FAIL') + ' ' + m); if (!c) fail++; };
  try { unit(ok); await api(ok); } catch (e) { console.log('FAIL 例外 ' + (e && e.stack || e)); fail++; try { await H.stopServer(); } catch { /* */ } }
  return fail;
}

const MUTS = [
  ['M1 資本額下限算錯', 'lib/companyTa.js', 'cap < f.capMinWan * 10000', 'cap < f.capMinWan * 1'],
  ['M2 不過濾在職', 'lib/companyTa.js', 'if (o.activeOnly) list', 'if (false) list'],
  ['M3 公司可視範圍放寬', 'server.js', 'const companies = (data.companies || []).filter(m => seeAll || countById[m.id]);\n  const oOwners', 'const companies = (data.companies || []).filter(m => true);\n  const oOwners'],
  ['M4 下載權限形同虛設', 'server.js', "return !!(u && (u.canDownloadContacts || u.role === 'admin'));\n}\nfunction taScope", "return true;\n}\nfunction taScope"],
  ['M5 Email 不去重', 'lib/companyTa.js', 'if (o.dedupeEmail) {', 'if (false) {'],
];

(async () => {
  let fail = await run();
  if (process.env.CTA_MUT === '1') {
    for (const [name, rel, from, to] of MUTS) {
      const file = path.join(ROOT, rel), src = fs.readFileSync(file, 'utf8');
      if (!src.includes(from)) { console.log('FAIL 變異無法套用：' + name); fail++; continue; }
      fs.writeFileSync(file, src.replace(from, to));
      delete require.cache[require.resolve('../lib/companyTa')];
      let mf;
      try { console.log('--- 變異 ' + name + '（必須有失敗）---'); mf = await run(); } finally { fs.writeFileSync(file, src); delete require.cache[require.resolve('../lib/companyTa')]; }
      console.log((mf > 0 ? 'OK  ' : 'FAIL') + ' 變異被抓到：' + name + '（失敗 ' + mf + ' 項）'); if (mf === 0) fail++;
    }
  }
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('例外', e && e.message); process.exit(1); });
