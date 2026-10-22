#!/usr/bin/env node
'use strict';
/**
 * 管理員「版本資訊」頁檢查：部署環境／資料庫型態由真實環境推導、稽核儲存狀態顯示、XSS 跳脫。
 * 不含帳密：測試帳號 _s_vp_admin 與隨機密碼只在記憶體；臨時 auth.json/data.json 於結束時刪除（PID 關閉伺服器）。
 * 守門：工作樹已有 auth.json/data.json/audit.log.json 時拒絕執行。預設埠 3141（VP_PORT 可改；禁用 3000/3020/3030/3121/3131）。
 *   node scripts/check-version-page.js
 */
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const ROOT = path.join(__dirname, '..');
const { deriveRuntimeInfo } = require(path.join(ROOT, 'lib', 'runtimeInfo.js'));

let pass = 0, fail = 0;
function ok(c, n) { if (c) pass++; else { fail++; console.log('  FAIL:', n); } }
function eq(a, b, n) { ok(JSON.stringify(a) === JSON.stringify(b), `${n}（得到 ${JSON.stringify(a)}）`); }

// ── 1. 純函式 deriveRuntimeInfo ──
{
  const j = deriveRuntimeInfo({});
  eq(j.deploy, { kind: 'local', vercelEnv: '', label: '本機' }, '預設＝本機');
  eq([j.dbBackend, j.storageBackend, j.sessionBackend], ['json', 'local', 'jwt'], '預設後端');
  eq(j.database.type, 'JSON 檔案（本機開發）', 'JSON 型態');
  eq(j.database.file, 'data.json', 'JSON 檔名');
  const p = deriveRuntimeInfo({ DB_BACKEND: 'postgres', VERCEL: '1', VERCEL_ENV: 'production', STORAGE_BACKEND: 'supabase' });
  eq(p.deploy.label, 'Vercel（production）', 'Vercel production');
  eq([p.dbBackend, p.storageBackend], ['postgres', 'supabase'], 'postgres 後端');
  eq(p.database.type, 'PostgreSQL（Supabase）', 'PG 型態');
  ok(!/data\.json|auth\.json/.test(JSON.stringify(p.database)), 'PG 模式不提 data.json/auth.json');
  eq(deriveRuntimeInfo({ VERCEL: '1', VERCEL_ENV: 'preview' }).deploy.label, 'Vercel（preview）', 'Vercel preview');
  eq(deriveRuntimeInfo({ VERCEL: '1' }).deploy.label, 'Vercel（unknown）', 'Vercel 無 VERCEL_ENV');
  eq(deriveRuntimeInfo({ VERCEL: '1', VERCEL_ENV: '<img src=x onerror=1>' }).deploy.vercelEnv, 'unknown', 'VERCEL_ENV 非列舉值不回顯');
  eq(deriveRuntimeInfo({ SESSION_BACKEND: 'memory', STORAGE_BACKEND: 'weird' }).sessionBackend, 'memory', 'session memory');
  eq(deriveRuntimeInfo({ STORAGE_BACKEND: 'weird' }).storageBackend, 'local', '未知 storage 回 local');
  eq(deriveRuntimeInfo({ DB_BACKEND: 'POSTGRES' }).dbBackend, 'json', 'DB_BACKEND 大小寫嚴格');
  // 不洩漏機敏
  const secret = { DB_BACKEND: 'postgres', DATABASE_URL: 'postgres://u:SECRETPW@db.example-host.supabase.co:5432/x', SUPABASE_SERVICE_KEY: 'SECRETKEY', JWT_SECRET: 'SECRETJWT', VERCEL: '1', VERCEL_ENV: 'production', VERCEL_URL: 'host-leak.vercel.app', STORAGE_BACKEND: 'supabase' };
  const out = JSON.stringify(deriveRuntimeInfo(secret));
  ok(!/SECRET|example-host|host-leak|postgres:\/\//.test(out), '輸出不含連線字串／主機／金鑰');
  ok(deriveRuntimeInfo(undefined).dbBackend === 'json', 'undefined env 不丟錯');
}

// ── 2. admin.html 渲染函式（DOM 層：取出內嵌函式、以 new Function 執行）──
const html = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
function extractFn(name) {
  const i = html.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('找不到 ' + name);
  let j = html.indexOf('{', i), depth = 0, k = j;
  for (; k < html.length; k++) { if (html[k] === '{') depth++; else if (html[k] === '}') { depth--; if (!depth) break; } }
  return html.slice(i, k + 1);
}
const R = new Function([extractFn('escH'), extractFn('verEnvLabel'), extractFn('verBackendsLabel'), extractFn('verDbHtml'), extractFn('verAuditHtml'),
  'return { escH, verEnvLabel, verBackendsLabel, verDbHtml, verAuditHtml };'].join('\n'))();
const W0 = { inserted: 0, failed: 0, retried: 0, legacyWrites: 0, fellBack: false, unavailable: '', lastError: '', lastFailedAt: null };
{
  const legacy = R.verAuditHtml({ auditStore: 'legacy', auditWrites: Object.assign({}, W0, { inserted: 12, legacyWrites: 12 }) });
  ok(legacy.includes('舊路徑（整包 JSON，5000 筆上限）'), 'legacy 文案');
  ok(legacy.includes('AUDIT_EVENTS_TABLE=1'), 'legacy 顯示啟用提示');
  ok(legacy.includes('#137333') && legacy.includes('>正常<'), 'legacy 正常＝綠色 chip');
  ok(legacy.includes('>12<'), '寫入成功次數');
  ok((legacy.match(/>0</g) || []).length >= 2, '數值 0 要顯示「0」（escH 對 0 會變空字串）');
  const table = R.verAuditHtml({ auditStore: 'table', auditWrites: Object.assign({}, W0, { inserted: 3 }) });
  ok(table.includes('新資料表 audit_events（只增不改）'), 'table 文案');
  ok(!table.includes('AUDIT_EVENTS_TABLE=1'), 'table 模式不顯示啟用提示');
  const fb = R.verAuditHtml({ auditStore: 'legacy', auditWrites: Object.assign({}, W0, { fellBack: true, unavailable: 'relation "audit_events" does not exist' }) });
  ok(fb.includes('#b06000') && fb.includes('已退回舊路徑'), 'fellBack＝琥珀色 chip');
  ok(fb.includes('自動退回舊路徑') && fb.includes('does not exist'.replace(/"/g, '&quot;').replace('does not exist', 'does not exist')), 'fellBack 原因顯示');
  const un = R.verAuditHtml({ auditStore: 'legacy', auditWrites: Object.assign({}, W0, { unavailable: 'x' }) });
  ok(un.includes('#b06000'), 'unavailable＝琥珀');
  const bad = R.verAuditHtml({ auditStore: 'table', auditWrites: Object.assign({}, W0, { failed: 2, retried: 2, fellBack: true, lastError: 'boom', lastFailedAt: '2026-10-13T01:02:03.000Z' }) });
  ok(bad.includes('#c5221f') && bad.includes('寫入失敗 2 筆'), 'failed>0＝紅色（優先於琥珀）');
  ok(bad.includes('2026-10-13 01:02:03'), 'lastFailedAt 顯示');
  // XSS
  const evil = '<img src=x onerror=alert(1)>"\'&<script>x</script>' + 'A'.repeat(400);
  const x = R.verAuditHtml({ auditStore: 'legacy', auditWrites: Object.assign({}, W0, { failed: 1, lastError: evil, unavailable: evil, lastFailedAt: '<b>t</b>' }) });
  ok(!/<img|<script|<b>/.test(x), 'XSS：lastError／unavailable／lastFailedAt 皆已跳脫');
  ok(x.includes('&lt;img src=x onerror=alert(1)&gt;'), 'XSS：以文字顯示');
  ok(!x.includes('A'.repeat(250)), 'lastError 顯示截斷');
  const xd = R.verDbHtml({ database: { type: '<i>t</i>', engine: '<script>', note: '"><svg onload=1>', file: '<f>', authFile: '<a>' } });
  ok(!/<i>|<script|<svg|<f>|<a>/.test(xd), 'XSS：資料庫區塊已跳脫');
  ok(!/<script|<i>/.test(R.verEnvLabel({ deploy: { label: 'ok' } }).replace(/ok/, '')), '環境標籤（呼叫端 escH）');
  // 舊版 API（缺欄位）不丟錯
  ok(R.verAuditHtml({}).includes('未提供稽核儲存資訊'), '舊 API：無 auditStore');
  ok(R.verAuditHtml({ auditStore: 'legacy' }).includes('舊路徑'), '舊 API：有 auditStore 無 auditWrites 不丟錯');
  ok(R.verAuditHtml({ auditStore: 'legacy', auditWrites: { inserted: 'NaN-x', failed: null } }).includes('正常'), '壞數值視為 0');
  ok(R.verEnvLabel({}).includes('—'), '舊 API：無 deploy');
  ok(R.verBackendsLabel({}) === '', '舊 API：無 backends');
  ok(R.verDbHtml({}).includes('—'), '舊 API：無 database');
  ok(R.verBackendsLabel({ backends: { db: 'postgres', storage: 'supabase', session: 'jwt' } }).includes('postgres'), 'backends 標籤');
}
// admin.html 已不含寫死字串，且內嵌 script 語法正確
ok(!html.includes("'On-Premise 本地部署'"), '已移除寫死的 On-Premise');
{
  const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g; let m, n = 0, bad = 0;
  while ((m = re.exec(html))) { n++; try { new Function(m[1]); } catch (e) { bad++; console.log('  inline script syntax:', e.message); } }
  ok(n > 0 && bad === 0, `admin.html 內嵌 script 語法（${n} 段）`);
}

// ── 3. 真實伺服器（JSON 模式沙盒）──
const PORT = parseInt(process.env.VP_PORT || '3141', 10);
const BASE = `http://localhost:${PORT}`;
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json', 'mail-outbox.json'].map(f => path.join(ROOT, f));
async function serverPart() {
  if ([3000, 3020, 3030, 3121, 3131].includes(PORT)) { console.log('拒絕：保留埠'); process.exit(2); }
  const g = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (g.length) { console.log('拒絕：工作樹已有', g.map(f => path.basename(f)).join('/')); process.exit(2); }
  const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
  const PW = 'Vp' + crypto.randomBytes(9).toString('hex') + '!Z', now = new Date().toISOString();
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users: [{ username: '_s_vp_admin', displayName: '_s_vp_admin', role: 'admin', active: true, password: bcrypt.hashSync(PW, 6), createdAt: now, passwordChangedAt: now, mustChangePassword: false, bu: [] }] }));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify({ contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [] }));
  let child = null, pid = null;
  try {
    const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
    delete env.DATABASE_URL; delete env.VERCEL; delete env.VERCEL_ENV; delete env.STORAGE_BACKEND; delete env.AUDIT_EVENTS_TABLE; delete env.MAIL_MODE;
    await assertPortFree(PORT);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
    child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: 'ignore' }); pid = child.pid;
    let up = false;
    for (let i = 0; i < 80 && !up; i++) { await new Promise(r => setTimeout(r, 250)); try { up = (await fetch(BASE + '/login.html')).ok; } catch { /* 尚未起來 */ } }
    ok(up, '沙盒伺服器啟動');
    if (!up) return;
    const lr = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: '_s_vp_admin', password: PW }) });
    const cookie = lr.headers.getSetCookie().map(c => c.split(';')[0]).join('; ');
    const r = await fetch(BASE + '/api/admin/version', { headers: { Cookie: cookie } });
    const d = await r.json();
    eq(r.status, 200, 'version 200');
    eq(d.auditStore, 'legacy', 'API auditStore＝legacy（預設）');
    ok(d.auditWrites && typeof d.auditWrites.inserted === 'number' && 'failed' in d.auditWrites && 'retried' in d.auditWrites && 'fellBack' in d.auditWrites && 'unavailable' in d.auditWrites, 'API auditWrites 欄位');
    eq(Object.assign({}, d.deploy, { git: undefined }), { kind: 'local', vercelEnv: '', label: '本機', git: undefined }, 'API deploy＝本機（git 欄位另由 check-version-deploy.js 驗證）');
    eq(d.backends, { db: 'json', storage: 'local', session: 'jwt' }, 'API backends');
    eq(d.database.type, 'JSON 檔案（本機開發）', 'API database.type');
    ok(!JSON.stringify(d).includes('On-Premise'), 'API 不含 On-Premise');
    const anon = await fetch(BASE + '/api/admin/version');
    ok(anon.status === 401 || anon.status === 403, '未登入不可讀 version（' + anon.status + '）');
    const rendered = R.verAuditHtml(d) + R.verDbHtml(d) + R.verEnvLabel(d);
    ok(rendered.includes('舊路徑') && rendered.includes('data.json') && rendered.includes('本機'), '真實 API 回應 → 渲染');
  } finally {
    if (pid) { try { process.kill(pid); } catch { /* 已結束 */ } await new Promise(r => setTimeout(r, 600)); }
    FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  }
}
serverPart().catch(e => { fail++; console.log('  FAIL: 伺服器段例外', e && e.message); }).then(() => {
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  process.exit(fail ? 1 : 0);
});
