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
  ok(ids(go({ statuses: ['customer'] })) === 'A' && ids(go({ statuses: ['prospect'] })) === 'BCD', 'U11 客戶狀態 我的客戶＝有 Won 商機（本組資料無 customerType）；潛在客戶＝其餘三種（已開商機未成交／尚無商機／新匯入）');
  ok(ids(go({ statuses: ['customer', 'prospect'] })) === 'ABCD' && ids(go({ statuses: [] })) === 'ABCD', 'U12 兩類都勾＝不限，與不勾等價');
  ok(ids(go({ statuses: ['won'] })) === 'A' && ids(go({ statuses: ['open'] })) === 'BCD' && ids(go({ statuses: ['noOpp'] })) === 'BCD' && ids(go({ statuses: ['newImport'] })) === 'BCD' && ids(go({ statuses: ['won', 'open'] })) === 'ABCD' && ids(go({ statuses: ['bogus'] })) === 'ABCD', 'U12b 舊四值安全對應（won→我的客戶；其餘→潛在客戶；未知值忽略）');
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

// ───────── 客戶狀態＝與「我的客戶」頁同一套規則（名片 customerType／同名公司／Won 商機的聯集）─────────
const CT = {
  companies: () => [
    { id: 'K1', name: 'ctA甲', taxId: '' }, { id: 'K2', name: 'ctB乙', taxId: '' }, { id: 'K2b', name: 'ctB乙2', taxId: '' },
    { id: 'K3', name: 'ctC丙', taxId: '' }, { id: 'K4', name: 'ctD丁', taxId: '' }, { id: 'K5', name: 'ctE戊', taxId: '' },
    { id: 'K6', name: 'ctF己', taxId: '' }, { id: 'K6X', name: 'ctF己X', taxId: '' }, { id: 'K7', name: 'ctG庚', taxId: '' },
    { id: 'K8', name: 'ctH辛', taxId: '' }, { id: 'K9', name: 'ctI壬', taxId: '' }, { id: 'K10', name: 'ctJ癸', taxId: '' },
  ],
  contacts: () => {
    const c = (id, owner, companyId, company, customerType, extra) => Object.assign({ id, owner, name: 'N' + id, title: 'T', email: id + '@ct.example.com', companyId, company, customerType, createdAt: iso, bu: owner === '_s_uerp' ? 'ERP' : 'ITS' }, extra || {});
    return [
      c('c1', '_s_uerp', 'K1', 'ctA甲（舊稱）', 'customer'),                 // 只有規則 2：卡片字串與主檔名稱不同，不會靠同名規則撈到
      c('c2', '_s_uerp', 'K2', 'ctB乙', ''),                                 // Won 商機（contactId）
      c('c2b', '_s_uerp', 'K2b', 'ctB乙2', ''),                              // Won 商機只掛公司名稱（contactId 空）＝規則 1 專屬
      c('c3', '_s_uerp', 'K3', 'ctC丙', ''),                                 // 兩者皆無
      c('c4a', '_s_uerp', 'K4', 'ctD丁', 'customer'), c('c4b', '_s_uerp', 'K4', 'ctD丁', 'prospect'),   // 混合
      c('c5', '_s_uerp', 'K5', 'ctE戊', ''), c('p5', '_pool', 'K5', 'ctE戊', 'customer', { isPlaceholder: true }),   // placeholder 設 customer 不算
      c('c6', '_s_uerp', 'K6', 'ctF己', 'prospect'),                         // 同名字串、不同 companyId
      c('c6x', '_s_uits', 'K6X', 'ctF己', 'customer'),                       //   ↳ 另一個主檔的 customer 名片，字串與 K6 相同
      c('c7', '_s_uerp', '', 'ctG庚', 'customer'),                           // 未掛 companyId，靠名稱字串（規則 3 專屬）
      c('c8', '_s_uerp', '', 'ctH辛 ', 'customer'),                          // 名稱尾端空白（trim 後相同）
      c('c9', '_s_uerp', 'K9', 'ctI壬', 'prospect'),                         // 歸 prospect 但有 Won 商機
      c('c10', '_s_uerp', 'K10', 'ctJ癸', 'prospect'), c('c10b', '_s_uerp', '', 'ctJ癸分公司', 'customer'),   // 名稱只是前綴相同不算
    ];
  },
  opps: () => [
    { id: 'q1', owner: '_s_uerp', contactId: 'c2', company: 'ctB乙', stage: 'Won' },
    { id: 'q2', owner: '_s_uerp', contactId: null, company: 'ctB乙2', stage: 'Won' },
    { id: 'q3', owner: '_s_uerp', contactId: 'c3', company: 'ctC丙', stage: 'B' },
    { id: 'q4', owner: '_s_uerp', contactId: 'c9', company: 'ctI壬', stage: 'Won' },
  ],
  CUSTOMER: ['ctA甲', 'ctB乙', 'ctB乙2', 'ctD丁', 'ctF己', 'ctF己X', 'ctG庚', 'ctH辛', 'ctI壬'].sort(),
  PROSPECT: ['ctC丙', 'ctE戊', 'ctJ癸'].sort(),
};

function unitCtype(ok) {
  const T = require('../lib/companyTa');
  const cos = CT.companies(), cs = CT.contacts(), ops = CT.opps();
  const name = r => r.matched.map(x => x.company.name).sort();
  const run = st => T.run({ companies: cos, contacts: cs, opps: ops, buUnion: () => [], filters: { statuses: [st] }, opts: { requireEmail: false } });
  const cu = name(run('customer')), pr = name(run('prospect'));
  ok(JSON.stringify(cu) === JSON.stringify(CT.CUSTOMER), 'C1 我的客戶集合＝名片歸類／Won／同名字串聯集（' + cu.join('、') + '）');
  ok(JSON.stringify(pr) === JSON.stringify(CT.PROSPECT), 'C2 潛在客戶集合＝兩者皆無／placeholder／名稱僅前綴相同（' + pr.join('、') + '）');
  const one = (nm, st) => name(run(st)).includes(nm);
  ok(one('ctA甲', 'customer') && !one('ctA甲', 'prospect'), 'C3 (a) 只有 customerType=customer、無 Won 商機（卡片字串與主檔名稱不同）→ 我的客戶');
  ok(one('ctB乙', 'customer') && one('ctB乙2', 'customer'), 'C4 (b) 只有 Won 商機、customerType 空（含 contactId 與只掛公司名稱兩種）→ 我的客戶');
  ok(one('ctC丙', 'prospect') && !one('ctC丙', 'customer'), 'C5 (c) 兩者皆無（只有未成交商機）→ 潛在客戶');
  ok(one('ctD丁', 'customer'), 'C6 (d) 公司下 customer 與 prospect 混合名片 → 我的客戶');
  ok(one('ctE戊', 'prospect'), 'C7 (e) placeholder 名片設 customer 不算 → 潛在客戶');
  ok(one('ctF己', 'customer') && one('ctF己X', 'customer'), 'C8 (f) 同名字串、不同 companyId：另一主檔的 customer 名片讓同名公司也算（與前端規則 c 一致）');
  ok(one('ctG庚', 'customer') && one('ctH辛', 'customer'), 'C9 (f) 未掛 companyId 的 customer 名片、名稱 trim 後相同 → 對應主檔算我的客戶');
  ok(one('ctJ癸', 'prospect'), 'C10 (f) 名稱只是前綴相同（分公司）不算同公司');
  ok(one('ctI壬', 'customer'), 'C11 (g) 名片歸 prospect 但有 Won 商機 → 我的客戶（與「我的客戶」頁一致）');
  // 範圍：呼叫端只傳自己看得到的名片 → 看不到別人的 customer 名片（不擴大可見性）
  const mine = cs.filter(c => c.owner === '_s_uerp');
  const rs = T.run({ companies: cos.filter(m => m.id === 'K6'), contacts: mine, opps: ops, buUnion: () => [], filters: { statuses: ['customer'] }, opts: {} });
  ok(rs.companyIds.length === 0, 'C12 可視範圍內沒有 customer 名片時，不會用範圍外的資料判成我的客戶');
  // 對拍：直接執行 _client/app.js 的 getWonContactIds 原始碼，與 TA 結果逐家比對
  const appSrc = fs.readFileSync(path.join(ROOT, '_client', 'app.js'), 'utf8');
  const m = appSrc.match(/function getWonContactIds\(\) \{[\s\S]*?\n\}/);
  ok(!!m, 'C13 找得到前端 getWonContactIds 原始碼');
  if (m) {
    const real = cs.filter(c => !c.isPlaceholder);
    const fw = new Function('allContacts', 'allOpportunities', m[0] + '; return getWonContactIds();')(real, ops);
    const trim = v => String(v == null ? '' : v).trim();
    const expect = cos.filter(co => {
      const linked = real.filter(c => c.companyId === co.id);
      const byName = real.some(c => fw.has(c.id) && trim(c.company) === trim(co.name));
      const wonByName = ops.some(o => o.stage === 'Won' && trim(o.company) === trim(co.name));
      return linked.some(c => fw.has(c.id)) || byName || wonByName;
    }).map(co => co.name).sort();
    ok(JSON.stringify(expect) === JSON.stringify(cu), '對拍 C14 TA「我的客戶」＝依前端 getWonContactIds 原始碼推得的公司集合（' + expect.length + ' 家）');
    const W = require('../lib/secretaryCustomers').wonContactIds(real, ops);
    const fwv = [...fw].filter(Boolean);
    ok(fwv.length === W.size && fwv.every(id => W.has(id)), '對拍 C15 前端 getWonContactIds＝後端 wonContactIds（同一份資料）');
  }
}

async function apiCtype(ok) {
  await H.fresh([
    H.mkUser('_s_adm', 'admin'), H.mkUser('_s_mkt', 'marketing', { canDownloadContacts: true }),
    H.mkUser('_s_uerp', 'user', { bu: ['ERP'], canDownloadContacts: true }), H.mkUser('_s_uits', 'user', { bu: ['ITS'], canDownloadContacts: true }),
  ], { companies: CT.companies(), contacts: CT.contacts(), opportunities: CT.opps() }, { env: {} });
  await H.loginAll(['_s_adm', '_s_mkt', '_s_uerp', '_s_uits']);
  const q = (who, st) => H.http(who, 'POST', '/api/companies/ta/query', { filters: { statuses: [st] }, opts: { requireEmail: false } });
  const exp = async (who, st, mode) => { const x = await H.http(who, 'POST', '/api/companies/ta/export', { filters: { statuses: [st] }, opts: { requireEmail: false }, mode }); return x.status === 200 ? XLSX.utils.sheet_to_json(XLSX.read(x.buf).Sheets[mode === 'companies' ? '公司' : '聯絡人']) : null; };
  // (h) 行銷：畫面、聯絡人名單、公司清單三者一致
  let r = await q('_s_mkt', 'customer');
  const coList = await exp('_s_mkt', 'customer', 'companies'), ctList = await exp('_s_mkt', 'customer', 'contacts');
  const coNames = (coList || []).map(x => x['公司名稱']).sort();
  ok(r.json.stats.companies === CT.CUSTOMER.length && JSON.stringify(coNames) === JSON.stringify(CT.CUSTOMER), 'CA1 (h) 行銷：我的客戶 畫面 ' + r.json.stats.companies + ' 家＝公司清單（' + coNames.length + ' 家）');
  const ctCos = [...new Set((ctList || []).map(x => x['公司名稱']))].sort();
  const withCards = CT.CUSTOMER.filter(n => !['ctG庚', 'ctH辛'].includes(n));   // 庚、辛沒有以 companyId 掛勾的名片 → 公司算數但沒有可下載的聯絡人
  ok(JSON.stringify(ctCos) === JSON.stringify(withCards) && r.json.stats.contacts === (ctList || []).length, 'CA2 (h) 聯絡人名單的公司＝我的客戶公司中有掛勾名片者（' + ctCos.length + ' 家／' + (ctList || []).length + ' 位），與畫面聯絡人數一致');
  ok(!(ctList || []).some(x => /c5@|p5@|c3@|c10@/.test(x['Email'] || '')), 'CA3 潛在客戶（丙／戊／癸）的名片不在我的客戶名單');
  r = await q('_s_mkt', 'prospect');
  const prList = await exp('_s_mkt', 'prospect', 'companies');
  ok(r.json.stats.companies === CT.PROSPECT.length && JSON.stringify((prList || []).map(x => x['公司名稱']).sort()) === JSON.stringify(CT.PROSPECT), 'CA4 行銷：潛在客戶 畫面＝下載公司清單＝' + CT.PROSPECT.length + ' 家');
  // 可視範圍差異：ERP 業務只看得到自己的名片；ITS 業務那張 customer 名片看不到 → 「ctF己」對 ERP 業務是潛在客戶
  r = await q('_s_uerp', 'customer');
  const erpCo = (await exp('_s_uerp', 'customer', 'companies') || []).map(x => x['公司名稱']).sort();
  const erpExpect = ['ctA甲', 'ctB乙', 'ctB乙2', 'ctD丁', 'ctI壬'].sort();
  ok(r.json.stats.companies === erpExpect.length && JSON.stringify(erpExpect) === JSON.stringify(erpCo), 'CA5 ERP 業務：我的客戶只算自己看得到的名片／商機（' + erpCo.join('、') + '）；庚／辛無掛勾名片看不到、己的 customer 名片在別人名下不算');
  r = await q('_s_uerp', 'prospect');
  ok(r.json.stats.companies === 4, 'CA6 ERP 業務：潛在客戶 4 家（丙／戊／己／癸）（' + r.json.stats.companies + '）');
  r = await q('_s_uits', 'customer');
  ok(r.json.stats.companies === 1, 'CA7 ITS 業務：只看得到己X 一家且為我的客戶（' + r.json.stats.companies + '）');
  await H.stopServer();
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
  r = await q('_s_mkt', { statuses: ['prospect'] }, { requireEmail: false });
  ok(r.json.stats.companies === 3 && r.json.stats.contacts === 1, 'A3 潛在客戶 3 家（乙已開商機、丙新匯入、丁待補）；客戶池待補名片不算聯絡人（只有乙的 1 位）');
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
  // 客戶狀態兩類：畫面（/ta）與下載（/ta/export）同一組條件、同一結果；舊值不報錯
  const coNames = async (st) => { const x = await H.http('_s_mkt', 'POST', '/api/companies/ta/export', { filters: { statuses: st }, opts: { requireEmail: false }, mode: 'companies' }); return x.status === 200 ? XLSX.utils.sheet_to_json(XLSX.read(x.buf).Sheets['公司']).map(y => y['公司名稱']).sort().join() : 'HTTP' + x.status; };
  r = await q('_s_mkt', { statuses: ['customer'] }, { requireEmail: false });
  ok(r.status === 200 && r.json.stats.companies === 1 && await coNames(['customer']) === '甲製造', 'A15 我的客戶＝只有有 Won 商機的甲（畫面 1 家、下載清單一致）');
  r = await q('_s_mkt', { statuses: ['prospect'] }, { requireEmail: false });
  ok(r.json.stats.companies === 3 && await coNames(['prospect']) === [ '丁待補', '丙新匯入', '乙金融' ].sort().join(), 'A16 潛在客戶＝乙／丙／丁（畫面 3 家、下載清單一致）');
  r = await q('_s_mkt', { statuses: ['customer', 'prospect'] }, { requireEmail: false });
  ok(r.json.stats.companies === 4 && (await coNames(['customer', 'prospect'])).split(',').length === 4, 'A17 兩類都勾＝全部 4 家（畫面與下載一致）');
  r = await q('_s_mkt', { statuses: ['won', 'noOpp', 'zzz'] }, { requireEmail: false });
  ok(r.status === 200 && r.json.stats.companies === 4 && (await coNames(['open'])).split(',').length === 3, 'A18 舊值不報錯：won+noOpp＝兩類皆有＝全部；open→潛在客戶 3 家');
  const lg2 = (await H.getLogs('_s_adm')).filter(l => l.action === 'EXPORT_TA').map(l => JSON.stringify(l)).join('');
  ok(/狀態=我的客戶/.test(lg2) && /狀態=潛在客戶/.test(lg2), 'A19 稽核摘要寫中文狀態名');
  r = await q('_s_mkt', { industries: 'x', capMinWan: 'abc', statuses: ['zzz'] }, { requireEmail: 'x' });
  ok(r.status === 200 && r.json.stats.companies === 4, 'A13 畸形輸入不當機');
  r = await H.http(null, 'POST', '/api/companies/ta/query', {});
  ok(r.status === 401 || r.status === 403, 'A14 未登入被擋（' + r.status + '）');
  await H.stopServer();
}

async function run() {
  let fail = 0; const ok = (c, m) => { console.log((c ? 'OK  ' : 'FAIL') + ' ' + m); if (!c) fail++; };
  try { unit(ok); unitCtype(ok); await api(ok); await apiCtype(ok); } catch (e) { console.log('FAIL 例外 ' + (e && e.stack || e)); fail++; try { await H.stopServer(); } catch { /* */ } }
  return fail;
}

const MUTS = [
  ['M1 資本額下限算錯', 'lib/companyTa.js', 'cap < f.capMinWan * 10000', 'cap < f.capMinWan * 1'],
  ['M2 不過濾在職', 'lib/companyTa.js', 'if (o.activeOnly) list', 'if (false) list'],
  ['M3 公司可視範圍放寬', 'server.js', 'const companies = (data.companies || []).filter(m => seeAll || countById[m.id]);\n  const oOwners', 'const companies = (data.companies || []).filter(m => true);\n  const oOwners'],
  ['M4 下載權限形同虛設', 'server.js', "return !!(u && (u.canDownloadContacts || u.role === 'admin'));\n}\nfunction taScope", "return true;\n}\nfunction taScope"],
  ['M5 Email 不去重', 'lib/companyTa.js', 'if (o.dedupeEmail) {', 'if (false) {'],
  ['M6 拿掉規則 1（Won 商機）', 'lib/companyTa.js', 'const isCustomer = hasWon ||', 'const isCustomer = false ||'],
  ['M7 拿掉規則 2（掛勾名片 customerType）', 'lib/companyTa.js', '|| real.some(c => won.has(c.id)) ||', '|| false ||'],
  ['M8 拿掉規則 3（同名字串）', 'lib/companyTa.js', '|| customerNames.has(str(m.name));', '|| false;'],
  ['M9 placeholder 也算', 'lib/companyTa.js', 'c && !c.isPlaceholder', 'c'],
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
