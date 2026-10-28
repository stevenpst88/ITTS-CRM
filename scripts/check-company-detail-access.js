// 企業主檔明細權限：行銷／管理員打得開「尚無名片」「只有客戶池待補名片」的主檔；業務仍只能開自己看得到名片的公司（埠 B2_PORT 預設 3462）
// 變異：CDA_MUT=1 會把修正還原後重跑，必須變紅。
process.env.B2_PORT = process.env.B2_PORT || '3462';
const H = require('./_b2-harness');
const fs = require('fs'), path = require('path');
const SERVER = path.join(__dirname, '..', 'server.js');
const FIX = "if (contacts.length === 0 && !isAdminOrMarketing(req)) return res.status(403)";
const OLD = "if (contacts.length === 0) return res.status(403)";

async function run() {
  let fail = 0; const ok = (c, m) => { console.log((c ? 'OK  ' : 'FAIL') + ' ' + m); if (!c) fail++; };
  const now = new Date().toISOString();
  const companies = [
    { id: 'co-new', name: '新匯入無名片公司', taxId: '11111111', createdAt: now },
    { id: 'co-pool', name: '只有待補名片公司', taxId: '22222222', createdAt: now },
    { id: 'co-real', name: '有業務名片公司', taxId: '33333333', createdAt: now },
  ];
  const contacts = [
    { id: 'p1', owner: '_pool', name: '（待補名片）', company: '只有待補名片公司', companyId: 'co-pool', isPlaceholder: true, bu: 'ERP' },
    { id: 'c1', owner: '_s_u1', name: '王小明', company: '有業務名片公司', companyId: 'co-real', bu: 'ERP' },
  ];
  await H.fresh([H.mkUser('_s_adm', 'admin'), H.mkUser('_s_mkt', 'marketing'), H.mkUser('_s_u1', 'user', { bu: ['ERP'] }), H.mkUser('_s_u2', 'user', { bu: ['MDM'] })],
    { companies, contacts, opportunities: [] }, { env: {} });
  await H.loginAll(['_s_adm', '_s_mkt', '_s_u1', '_s_u2']);
  const get = (who, id) => H.http(who, 'GET', '/api/companies/' + id);

  let r = await get('_s_mkt', 'co-new');  ok(r.status === 200 && r.json.company.id === 'co-new' && r.json.contacts.length === 0, '行銷開「尚無名片」主檔 200，聯絡人 0（' + r.status + '）');
  r = await get('_s_mkt', 'co-pool');     ok(r.status === 200 && r.json.contacts.length === 0, '行銷開「只有待補名片」主檔 200，不外洩客戶池名片（' + r.status + '）');
  r = await get('_s_mkt', 'co-real');     ok(r.status === 200 && r.json.contacts.length === 1, '行銷開有業務名片的主檔 200，聯絡人 1（' + r.status + '）');
  r = await get('_s_adm', 'co-new');      ok(r.status === 200, '管理員開「尚無名片」主檔 200（' + r.status + '）');
  r = await get('_s_u1', 'co-new');       ok(r.status === 403, '業務開別人沒名片的主檔仍 403（' + r.status + '）');
  r = await get('_s_u1', 'co-pool');      ok(r.status === 403, '業務開只有待補名片的主檔仍 403（' + r.status + '）');
  r = await get('_s_u1', 'co-real');      ok(r.status === 200 && r.json.contacts.length === 1, '業務開自己名片的公司 200（' + r.status + '）');
  r = await get('_s_u2', 'co-real');      ok(r.status === 403, '其他 BU 業務開該公司仍 403（' + r.status + '）');
  r = await get('_s_mkt', 'co-none');     ok(r.status === 404, '不存在的主檔 404（' + r.status + '）');
  await H.stopServer();
  return fail;
}

(async () => {
  let fail = await run();
  if (process.env.CDA_MUT === '1') {
    const src = fs.readFileSync(SERVER, 'utf8');
    if (!src.includes(FIX)) { console.log('FAIL 找不到修正內容，無法跑變異'); process.exit(1); }
    fs.writeFileSync(SERVER, src.replace(FIX, OLD));
    let mutFail;
    try { console.log('--- 變異：還原修正（必須有失敗）---'); mutFail = await run(); } finally { fs.writeFileSync(SERVER, src); }
    console.log((mutFail > 0 ? 'OK  ' : 'FAIL') + ' 變異被抓到（失敗 ' + mutFail + ' 項）');
    if (mutFail === 0) fail++;
  }
  process.exit(fail ? 1 : 0);
})().catch(e => { console.log('例外', e.message); process.exit(1); });
