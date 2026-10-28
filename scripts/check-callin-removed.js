// Call-in Pass 移除後的煙霧檢查：/api/callins 404、poll-bundle／權限無 callin、舊 callins 資料原封不動（埠 B2_PORT 預設 3461）
process.env.B2_PORT = process.env.B2_PORT || '3461';
const H = require('./_b2-harness');
(async () => {
  let fail = 0; const ok = (c, m) => { console.log((c ? 'OK  ' : 'FAIL') + ' ' + m); if (!c) fail++; };
  const legacy = [{ id: 'ci1', company: 'LEGACY-CO', contactName: 'x', assignedTo: '_s_u1', createdBy: '_s_adm', status: 'pending' }];
  await H.fresh([H.mkUser('_s_adm', 'admin'), H.mkUser('_s_u1', 'user', { bu: ['ERP'] })], { callins: legacy, contacts: [], opportunities: [] }, { env: {} });
  await H.loginAll(['_s_adm', '_s_u1']);
  for (const p of ['/api/callins']) { const r = await H.http('_s_adm', 'GET', p); ok(r.status === 404, 'GET ' + p + ' => ' + r.status); }
  const r2 = await H.http('_s_u1', 'POST', '/api/callins', { company: 'a' }); ok(r2.status === 404, 'POST /api/callins => ' + r2.status);
  const pb = await H.http('_s_u1', 'GET', '/api/poll-bundle'); ok(pb.status === 200 && !/callin/i.test(JSON.stringify(pb.json)), 'poll-bundle 200 且無 callin 欄位');
  const pm = await H.http('_s_u1', 'GET', '/api/me/permissions'); ok(pm.status === 200 && !(pm.json.features || []).includes('callin'), 'permissions 無 callin');
  const ad = await H.http('_s_adm', 'GET', '/api/admin/role-permissions'); ok(ad.status === 200 && !/callin/i.test(JSON.stringify(ad.json)), '功能權限頁無 callin');
  const idx = await H.http('_s_u1', 'GET', '/index.html'); const mm = idx.text.match(/.{40}(Call-in|callinView|navCallin).{40}/); ok(idx.status === 200 && !mm, '首頁 HTML 無 Call-in ' + idx.status + ' ' + (mm ? mm[0] : ''));
  const raw = JSON.parse(require('fs').readFileSync('data.json', 'utf8'));
  ok(Array.isArray(raw.callins) && raw.callins.length === 1 && raw.callins[0].company === 'LEGACY-CO', '舊 callins 資料原封不動');
  await H.stopServer(); process.exit(fail ? 1 : 0);
})().catch(e => { console.log('例外', e.message); process.exit(1); });
