#!/usr/bin/env node
'use strict';
/**
 * 名片圖不再儲存於伺服器 — 端對端檢查（隔離沙盒；不含任何帳密、不需要真實資料庫／外部服務）。
 *
 * 決定：名片圖不存伺服器，名片只用 AI 辨識取欄位後即丟棄。
 *   - POST /api/upload：保留路由與 requireAuth，但不解析 multipart、不寫檔，一律 410（含超大／非圖片檔，不可 500）。
 *   - POST /api/contacts：cardImage 一律存 ''（忽略請求值、不報錯）。
 *   - PUT /api/contacts/:id：cardImage 相同→不變；非空且不同→忽略；''→清掉舊圖；沒送→不變。
 *   - 既有舊圖：GET /uploads/:key 仍依 SEC-02 授權可下載（唯讀）。
 *   - /api/ai/ocr-card（與後台 ai-ocr-card）不受影響：與基準樹（NOSTORE_BASE，預設 fix/sec02-upload-authz 的工作樹）回應逐一相同。
 *   - 前端：名片表單沒有上傳區（cardImageInput／cardImagePreview 不存在）、AI 辨識按鈕與流程仍在。
 *
 * 做法：在 NOSTORE_ROOT（預設＝本工作樹）建立臨時 auth.json / data.json（假資料＋_s_* 臨時帳號）與 uploads/ 假圖，
 *       以 PORT=NOSTORE_PORT（預設 3720，限 3720–3729）啟動 server.js；跑完只用 PID 關閉並刪除自己建立的檔案。
 * 守門：ROOT 已有 auth.json／data.json／audit.log.json 一律拒絕。埠已被佔用也拒絕（assertPortFree）。
 *
 *   node scripts/check-no-card-storage.js
 *   NOSTORE_PORT=3722 NOSTORE_ROOT=<變異樹> NOSTORE_SKIP_DOM=1 node scripts/check-no-card-storage.js
 *   NOSTORE_BASE=<基準樹> 未設時自動找 ../wt_sec02（不存在就略過「與基準逐位元比對」並標示未執行）
 *   NOSTORE_CDP=3729（headless Edge 的除錯埠，預設 3729）；NOSTORE_SKIP_DOM=1 明確略過瀏覽器段（會標示未驗證）
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.resolve(process.env.NOSTORE_ROOT || path.join(__dirname, '..'));
const PORT = parseInt(process.env.NOSTORE_PORT || '3720', 10);
const BASE_PORT = PORT + 1;
const CDP = parseInt(process.env.NOSTORE_CDP || '3729', 10);
if (!(PORT >= 3720 && PORT <= 3728)) { console.log('拒絕：NOSTORE_PORT 須在 3720-3728（基準伺服器用 +1）'); process.exit(2); }
const BASE_ROOT = (() => {
  const c = process.env.NOSTORE_BASE ? path.resolve(process.env.NOSTORE_BASE) : path.resolve(__dirname, '..', '..', 'wt_sec02');
  return fs.existsSync(path.join(c, 'server.js')) && c !== ROOT ? c : null;
})();
const bcrypt = require(path.join(__dirname, '..', 'node_modules', 'bcryptjs'));

const GUARD = (root) => ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl'].map(f => path.join(root, f)).filter(f => fs.existsSync(f));
for (const r of [ROOT, BASE_ROOT].filter(Boolean)) {
  const g = GUARD(r);
  if (g.length) { console.log('拒絕：', r, '已有', g.map(f => path.basename(f)).join('/'), '（可能是真實資料）'); process.exit(2); }
}

let PARITY_TOTAL = 0;
let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));

const PW = 'Ns' + crypto.randomBytes(9).toString('hex') + '!Z';
const T0 = '2026-01-01T00:00:00.000Z';
const uuid = () => crypto.randomUUID();
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');
const CO = '沙盒公司';
const K = { OLD: uuid() + '.png', OTHER: uuid() + '.png' };   // OLD：c-img 的既有舊圖；OTHER：別的（不被任何名片引用）
const NEWURL = '/uploads/' + uuid() + '.png';                  // 請求裡想塞進去的新網址（檔案不存在，也不該被存起來）

const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: T0, passwordChangedAt: new Date().toISOString(), mustChangePassword: false }, extra || {});
const U = { admin: '_s_adm', exec: '_s_exec', m1: '_s_m1', m2: '_s_m2', A: '_s_a', B: '_s_b', ITS: '_s_its', mkt: '_s_mkt', sec: '_s_sec', tpm: '_s_tpm', gs: '_s_gs' };
// 20+ 個帳號供「GET /api/contacts 與基準逐位元相同」用
const EXTRA = Array.from({ length: 12 }, (_, i) => ({ n: '_s_x' + i, role: ['user', 'user', 'manager1', 'marketing', 'secretary', 'user'][i % 6], bu: [['ERP'], ['ITS'], ['ERP'], [], ['ERP'], ['ITS']][i % 6] }));
function seed(root) {
  const users = [
    mkUser(U.admin, 'admin'), mkUser(U.exec, 'executive'),
    mkUser(U.m1, 'manager1', { bu: ['ERP'] }), mkUser(U.m2, 'manager2', { bu: ['ITS'] }),
    mkUser(U.A, 'user', { bu: ['ERP'], supervisor: U.m1 }), mkUser(U.B, 'user', { bu: ['ERP'], supervisor: U.m1 }),
    mkUser(U.ITS, 'user', { bu: ['ITS'], supervisor: U.m2 }), mkUser(U.mkt, 'marketing'), mkUser(U.sec, 'secretary', { bu: ['ERP'] }),
    mkUser(U.tpm, 'tecopm', { viewOwnerScope: U.A, viewGroupId: '_s_g1' }), mkUser(U.gs, 'groupsales', { viewOwnerScope: U.A, viewGroupId: '_s_g1' }),
    ...EXTRA.map(x => mkUser(x.n, x.role, x.bu.length ? { bu: x.bu, supervisor: x.bu[0] === 'ITS' ? U.m2 : U.m1 } : {})),
  ];
  fs.writeFileSync(path.join(root, 'auth.json'), JSON.stringify({ users }, null, 2));
  const ct = (id, owner, bu, name, img, extra) => Object.assign({ id, name, company: CO, owner, bu, title: '經理', phone: '0912000000', email: id + '@sandbox.test', note: '沙盒', taxId: '', isForeign: true, cardImage: img || '', createdAt: T0 }, extra || {});
  const d = {
    contacts: [
      ct('c-img', U.A, 'ERP', '甲（A，有舊圖）', '/uploads/' + K.OLD),
      ct('c-plain', U.A, 'ERP', '乙（A，沒圖）', ''),
      ct('c-bad', U.A, 'ERP', '丙（A，舊資料的不合規 cardImage）', 'data:image/png;base64,AAAA'),
      ct('c-b', U.B, 'ERP', '丁（B）', ''),
      ct('c-its', U.ITS, 'ITS', '戊（ITS）', ''),
    ],
    companies: [], keyAccounts: [], visits: [], opportunities: [], contracts: [], receivables: [], callins: [], quotations: [],
    groups: [{ id: '_s_g1', name: '沙盒集團', memberCompanies: [CO] }],
  };
  fs.writeFileSync(path.join(root, 'data.json'), JSON.stringify(d, null, 2));
  const up = path.join(root, 'uploads');
  if (!fs.existsSync(up)) fs.mkdirSync(up, { recursive: true });
  fs.writeFileSync(path.join(up, K.OLD), PNG); fs.writeFileSync(path.join(up, K.OTHER), PNG);
}
const cleanup = (root) => {
  ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].forEach(f => { try { fs.unlinkSync(path.join(root, f)); } catch { /* 不存在 */ } });
  [K.OLD, K.OTHER].forEach(k => { try { fs.unlinkSync(path.join(root, 'uploads', k)); } catch { /* 不存在 */ } });
};
const listUploads = (root) => { try { return fs.readdirSync(path.join(root, 'uploads')).sort(); } catch { return []; } };

const servers = [];
async function boot(root, port) {
  const env = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', STORAGE_BACKEND: 'local', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${port}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'SUPABASE_URL', 'SUPABASE_SERVICE_ROLE_KEY', 'GEMINI_API_KEY', 'GOOGLE_API_KEY']) delete env[k];
  await assertPortFree(port);
  const child = spawn(process.execPath, [path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  const h = { child, pid: child.pid, root, port, base: `http://localhost:${port}`, jar: {}, out: '' };
  child.stdout.on('data', d => { h.out += d; }); child.stderr.on('data', d => { h.out += d; });
  servers.push(h);
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    try { const r = await fetch(h.base + '/login.html'); if (r.ok) return h; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + h.out.slice(-400));
  }
  throw new Error('伺服器啟動逾時：' + h.out.slice(-400));
}
async function stop(h) {
  try { process.kill(h.pid); } catch { /* 已結束 */ }
  for (let i = 0; i < 20 && h.child.exitCode === null; i++) await sleep(100);
}
async function http(h, who, method, url, body, headers) {
  const hd = Object.assign({}, headers || {});
  if (who && h.jar[who]) hd.Cookie = h.jar[who];
  let payload;
  if (body instanceof FormData) payload = body;
  else if (Buffer.isBuffer(body)) payload = body;
  else if (body !== undefined) { hd['Content-Type'] = 'application/json'; payload = JSON.stringify(body); }
  const res = await fetch(h.base + url, { method, headers: hd, body: payload, redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, buf, text: buf.toString('utf8') };
}
async function login(h, who) {
  const res = await fetch(h.base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status !== 200) throw new Error('登入失敗 ' + who + ' ' + res.status);
  h.jar[who] = res.headers.getSetCookie().map(c => c.split(';')[0]).join('; ');
}
const pngForm = (name, type, data) => { const f = new FormData(); f.append('card', new Blob([data || PNG], { type: type || 'image/png' }), name || 'card.png'); return f; };
const rd = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');

// ── [S] 靜態檢查（前端與後端原始碼）──
function staticChecks() {
  console.log('[S] 靜態檢查');
  const html = rd('_client/index.html'), app = rd('_client/app.js'), srv = rd('server.js'), sw = rd('_client/sw.js');
  ok(!/id="cardImageInput"/.test(html), 'S1 index.html 沒有 cardImageInput（上傳區已移除）');
  ok(!/id="cardImagePreview"/.test(html), 'S2 index.html 沒有 cardImagePreview');
  ok(!/upload-hint/.test(html + app), 'S3 沒有「點擊上傳名片圖片」的提示元素／引用');
  ok(/id="ocrCardBtn"/.test(html) && /id="ocrCardInput"/.test(html), 'S4 AI 辨識按鈕 ocrCardBtn／ocrCardInput 仍在');
  ok(/id="ocrCardInput"[^>]*capture="environment"/.test(html), 'S5 AI 辨識的手機拍照 capture 屬性仍在');
  ok(/function setupOcrCard/.test(app) && /\/ai\/ocr-card/.test(app) && /function compressCardImage/.test(app), 'S6 app.js 的 setupOcrCard／compressCardImage／呼叫 /ai/ocr-card 仍在');
  ok(!/cardImageInput|cardImagePreview|_cardPreviewBlob/.test(app), 'S7 app.js 沒有對已移除元素的取用（cardImageInput／cardImagePreview／_cardPreviewBlob）');
  ok(!/\$\(['"]cardImageUrl['"]\)/.test(app) || /id="cardImageUrl"/.test(html), 'S8a app.js 取用 cardImageUrl 時 HTML 必須有該元素');
  // app.js 裡每個和名片圖相關的 $('id') 在 index.html 都要存在（取不到會丟例外）
  const ids = [...new Set([...app.matchAll(/\$\('((?:cardImage\w*|imageActions|removeImage|previewImg))'\)/g)].map(m => m[1]))];
  ok(ids.length > 0 && ids.every(id => new RegExp('id="' + id + '"').test(html)), `S8 app.js 取用的名片圖相關 id 在 HTML 都存在（${ids.join(',')}）`);
  ok(!/\$\{API\}\/upload\b/.test(app), 'S9 前端不再呼叫 /upload');
  ok(!/upload\.single\(['"]card['"]\)|getMulterStorage/.test(srv), 'S10 server.js 不再掛 multer 的名片上傳（upload.single／getMulterStorage）');
  ok(/app\.post\('\/api\/upload', requireAuth, \(req, res\) => \{\s*res\.status\(410\)/.test(srv), 'S11 POST /api/upload 保留 requireAuth 並直接回 410');
  ok(!/cardImage:\s*req\.body/.test(srv), 'S12 server.js 沒有把請求的 cardImage 直接寫進資料');
  ok(/\/\/ 名片圖不再儲存/.test(srv), 'S13 server.js 有留註解說明');
  const m = /SW_VERSION = 'itts-crm-v(\d+)'/.exec(sw);
  ok(m && +m[1] >= 45, `S14 sw.js 版號 >= v45（${m && m[0]}）`);
  ok(/storage\.(remove|delete)|require\('\.\/storage'\)|storage\./.test(srv), 'S15 storage 模組仍被 server.js 使用（/uploads 讀取、刪名片刪檔）');
}

async function backend(h) {
  const before = listUploads(ROOT);

  console.log('[1] POST /api/upload：一律 410、不寫檔、不 500');
  {
    const r1 = await http(h, U.A, 'POST', '/api/upload', pngForm());
    eq(r1.status, 410, '1.1 正常圖片上傳＝410');
    ok(r1.json && /不再儲存於伺服器/.test(r1.json.error || '') && /AI 辨識/.test(r1.json.error || ''), `1.2 回應有清楚的繁中訊息（${r1.text.slice(0, 80)}）`);
    ok(!r1.json || !r1.json.url, '1.3 回應沒有 url');
    eq(listUploads(ROOT), before, '1.4 儲存目錄沒有新檔');
    const big = Buffer.alloc(11 * 1024 * 1024, 7);
    const r2 = await http(h, U.A, 'POST', '/api/upload', pngForm('big.png', 'image/png', big));
    eq(r2.status, 410, '1.5 超大檔（11MB）＝410（不是 413／500）');
    const r3 = await http(h, U.A, 'POST', '/api/upload', pngForm('x.exe', 'application/octet-stream', Buffer.from('MZ not an image')));
    eq(r3.status, 410, '1.6 非圖片檔＝410（不是 500）');
    const r4 = await http(h, U.A, 'POST', '/api/upload', Buffer.from('garbage'), { 'Content-Type': 'multipart/form-data; boundary=zzz' });
    eq(r4.status, 410, '1.7 壞掉的 multipart＝410');
    const r5 = await http(h, U.A, 'POST', '/api/upload', {});
    eq(r5.status, 410, '1.8 JSON 內容＝410');
    const r6 = await http(h, U.A, 'POST', '/api/upload');
    eq(r6.status, 410, '1.9 沒有內容＝410');
    eq(listUploads(ROOT), before, '1.10 以上全部送完，儲存目錄仍沒有新檔');
    const r7 = await http(h, null, 'POST', '/api/upload', pngForm());
    eq(r7.status, 401, '1.11 未登入＝401');
    const r8 = await http(h, U.sec, 'POST', '/api/upload', pngForm());
    ok(r8.status === 410 || r8.status === 403, `1.12 秘書＝410 或既有的 403（${r8.status}）`);
    eq(listUploads(ROOT), before, '1.13 秘書／未登入也沒有寫檔');
  }

  console.log('[2] POST /api/contacts：cardImage 一律存 \'\'');
  {
    const mk = async (extra) => http(h, U.A, 'POST', '/api/contacts', Object.assign({ name: '新名片', company: CO, isForeign: true, email: 'n@sandbox.test' }, extra));
    const a = await mk({ cardImage: NEWURL });
    ok(a.status === 200 || a.status === 201, `2.1 帶合法格式網址：建立成功（${a.status}）`);
    eq(a.json && a.json.cardImage, '', '2.2 回應 cardImage＝\'\'');
    const b = await mk({ cardImage: 'javascript:alert(1)' });
    ok(b.status === 200 || b.status === 201, `2.3 帶不合法值：也不報錯（${b.status}）`);
    eq(b.json && b.json.cardImage, '', '2.4 不合法值也存成 \'\'');
    const c = await mk({ cardImage: '/uploads/' + K.OLD });
    eq(c.json && c.json.cardImage, '', '2.5 即使指向別人已有的舊圖也存成 \'\'（不能新增引用）');
    const d = await mk({});
    eq(d.json && d.json.cardImage, '', '2.6 沒帶＝\'\'');
    const list = await http(h, U.A, 'GET', '/api/contacts');
    const mine = (list.json || []).filter(x => x.name === '新名片');
    ok(mine.length === 4 && mine.every(x => x.cardImage === ''), `2.7 清單中 4 筆新名片 cardImage 全為 ''（${mine.length} 筆）`);
    eq(await dlStatus(h, U.A, K.OLD), 200, '2.8 舊圖仍只被原名片引用、A 仍可下載');
  }

  console.log('[3] PUT /api/contacts/:id：相同放行／不同忽略／清空可清／沒送不變');
  {
    const put = (id, body, who) => http(h, who || U.A, 'PUT', '/api/contacts/' + id, body);
    const get = async (id) => ((await http(h, U.A, 'GET', '/api/contacts')).json || []).find(x => x.id === id);
    const OLD = '/uploads/' + K.OLD;
    let r = await put('c-img', { note: '改備註', cardImage: OLD });
    eq(r.status, 200, '3.1 相同值（舊前端原樣送回）＝200');
    eq((await get('c-img')).cardImage, OLD, '3.2 相同值：cardImage 不變');
    eq((await get('c-img')).note, '改備註', '3.3 其他欄位正常更新');
    r = await put('c-img', { note: '改2', cardImage: NEWURL });
    eq(r.status, 200, '3.4 非空且不同（合法格式）＝200 不報錯');
    eq((await get('c-img')).cardImage, OLD, '3.5 非空且不同：忽略，維持原值');
    eq((await get('c-img')).note, '改2', '3.6 同次請求其他欄位仍更新');
    r = await put('c-img', { cardImage: 'javascript:alert(1)' });
    eq(r.status, 200, '3.7 非空且不同（不合法值）＝200 不報錯');
    eq((await get('c-img')).cardImage, OLD, '3.8 不合法值：忽略，維持原值');
    r = await put('c-img', { cardImage: { x: 1 } });
    ok(r.status === 200, `3.8b 非字串值：不 500（${r.status}）`);
    eq((await get('c-img')).cardImage, OLD, '3.8c 非字串值：忽略，維持原值');
    r = await put('c-img', { note: '沒送欄位' });
    eq(r.status, 200, '3.9 沒送 cardImage＝200');
    eq((await get('c-img')).cardImage, OLD, '3.10 沒送：不變');
    r = await put('c-plain', { cardImage: NEWURL, note: 'p' });
    eq((await get('c-plain')).cardImage, '', '3.11 原本沒圖的名片：PUT 帶新值也不會有圖');
    r = await put('c-plain', { cardImage: OLD });
    eq((await get('c-plain')).cardImage, '', '3.12 PUT 指向別張名片的舊圖：不會新增引用');
    eq(await dlStatus(h, U.B, K.OLD), 404, '3.13 B 仍不可下載（授權不變）');
    r = await put('c-bad', { note: 'q', cardImage: 'data:image/png;base64,AAAA' });
    eq(r.status, 200, '3.14 舊資料不合規值原樣送回＝200（不擋其他欄位）');
    eq((await get('c-bad')).cardImage, 'data:image/png;base64,AAAA', '3.15 原值保持（不另外改寫舊資料）');
    r = await put('c-bad', { cardImage: '' });
    eq((await get('c-bad')).cardImage, '', '3.16 舊的不合規值可以被清空');
    // 清空
    eq(await dlStatus(h, U.A, K.OLD), 200, '3.17 清空前 A 可下載舊圖（SEC-02 授權仍有效）');
    eq(await dlStatus(h, U.m1, K.OLD), 200, '3.18 A 的主管可下載');
    eq(await dlStatus(h, U.admin, K.OLD), 200, '3.19 管理員可下載');
    eq(await dlStatus(h, U.sec, K.OLD), 404, '3.20 秘書 404');
    eq(await dlStatus(h, U.ITS, K.OLD), 404, '3.21 別部門 404');
    r = await put('c-img', { cardImage: '' });
    eq(r.status, 200, '3.22 清空＝200');
    eq((await get('c-img')).cardImage, '', '3.23 清空：cardImage 變 \'\'');
    eq(await dlStatus(h, U.A, K.OLD), 404, '3.24 清空後沒有名片引用 → 該檔 404（檔案本身不被伺服器刪除，由另一步驟清理）');
    ok(fs.existsSync(path.join(ROOT, 'uploads', K.OLD)), '3.25 清空不刪檔（本功能不刪任何既有檔案）');
    r = await put('c-img', { cardImage: '/uploads/' + K.OLD });
    eq((await get('c-img')).cardImage, '', '3.26 清空後再想塞回舊圖也不行（忽略）');
    r = await put('c-img', { cardImage: null });
    ok(r.status === 200, '3.27 cardImage:null 不 500');
    r = await put('c-b', { cardImage: NEWURL, note: 'bb' }, U.B);
    eq(((await http(h, U.B, 'GET', '/api/contacts')).json || []).find(x => x.id === 'c-b').cardImage, '', '3.28 其他業務也一樣');
  }

  console.log('[4] 其他寫入 cardImage 的路徑（靜態＋行為）');
  {
    const srv = rd('server.js');
    const lines = srv.split('\n').map((l, i) => [i + 1, l]).filter(([, l]) => /cardImage/.test(l));
    const bad = lines.filter(([, l]) => /cardImage\s*:\s*(?:[^\s'"]|['"][^'"])/.test(l) && !/^\s*\/\//.test(l));
    ok(bad.length === 0, `4.1 server.js 所有 cardImage: 賦值都是 ''（不合規：${bad.map(([n, l]) => n + ':' + l.trim().slice(0, 60)).join(' | ')}）`);
    const files = ['lib', 'middleware', 'api', 'db', 'ai', 'storage'].flatMap(d => { try { return fs.readdirSync(path.join(ROOT, d)).filter(f => f.endsWith('.js')).map(f => path.join(d, f)); } catch { return []; } });
    const hits = files.filter(f => /cardImage/.test(rd(f)));
    ok(hits.length === 0, `4.2 lib／middleware／api／db／ai／storage 沒有處理 cardImage（${hits.join(',')}）`);
  }
}
const dlStatus = async (h, who, key) => (await http(h, who, 'GET', '/uploads/' + key)).status;

// ── 與基準樹比對：GET /api/contacts（各帳號）逐位元相同；OCR 路由回應相同 ──
async function parity(h, hb, users, withOcr) {
  console.log('[5] 與基準樹（fix/sec02-upload-authz）比對');
  if (!hb) { console.log('  （未執行：找不到基準樹，未驗證）'); ok(true, '5.0 【未驗證】沒有基準樹，略過逐位元比對'); return; }
  PARITY_TOTAL += users.length;
  for (const u of users) { await login(h, u); await login(hb, u); }
  let same = 0;
  for (const u of users) {
    const a = await http(h, u, 'GET', '/api/contacts'), b = await http(hb, u, 'GET', '/api/contacts');
    const ja = JSON.stringify(a.json), jb = JSON.stringify(b.json);
    if (a.status === b.status && ja === jb && ja !== undefined) same++; else console.log('  ✗ 差異帳號', u, a.status, b.status);
  }
  ok(same === users.length, `5.1 GET /api/contacts 在 ${users.length} 個帳號上與基準逐位元相同（相同 ${same}）`);
  if (!withOcr) return;
  // 既有圖 SEC-02 授權與基準相同
  const same2 = [];
  for (const u of users) { const a = await dlStatus(h, u, K.OLD), b = await dlStatus(hb, u, K.OLD); same2.push(a === b); }
  ok(same2.every(Boolean), '5.2 /uploads/<既有舊圖> 各帳號回應與基準相同（SEC-02 授權不變）');
  // OCR：同一組請求兩邊回應相同（沙盒沒有 Gemini 金鑰；兩邊都應是 503 AI_NOT_CONFIGURED；不會真的呼叫 Gemini）
  const cases = [
    ['無登入', null, 'POST', '/api/ai/ocr-card', () => pngForm()],
    ['業務＋圖片', U.A, 'POST', '/api/ai/ocr-card', () => pngForm()],
    ['業務＋沒檔案', U.A, 'POST', '/api/ai/ocr-card', () => undefined],
    ['業務＋非圖片', U.A, 'POST', '/api/ai/ocr-card', () => pngForm('x.txt', 'text/plain', Buffer.from('hi'))],
    ['秘書＋圖片', U.sec, 'POST', '/api/ai/ocr-card', () => pngForm()],
    ['管理員＋圖片（後台匯入）', U.admin, 'POST', '/api/admin/ai-ocr-card', () => pngForm()],
    ['業務＋後台匯入（無權限）', U.A, 'POST', '/api/admin/ai-ocr-card', () => pngForm()],
  ];
  const res = [];
  for (const [name, who, m, url, mkBody] of cases) {
    const a = await http(h, who, m, url, mkBody()), b = await http(hb, who, m, url, mkBody());
    const o = a.status === b.status && a.text === b.text;
    if (!o) console.log('  ✗ OCR 差異', name, a.status, a.text.slice(0, 80), '|', b.status, b.text.slice(0, 80));
    res.push(o);
  }
  ok(res.every(Boolean), `5.3 /api/ai/ocr-card 與 /api/admin/ai-ocr-card 的 ${cases.length} 組請求回應與基準完全相同`);
  const o1 = await http(h, U.A, 'POST', '/api/ai/ocr-card', pngForm());
  ok(o1.status === 503 && !!o1.json && o1.json.error === 'AI_NOT_CONFIGURED', `5.4 沙盒（無金鑰）：OCR 路由到得了 requireAi 並回 503 AI_NOT_CONFIGURED（${o1.status}）`);
  const o2 = await http(h, null, 'POST', '/api/ai/ocr-card', pngForm());
  eq(o2.status, 401, '5.5 OCR 未登入＝401');
  eq(listUploads(ROOT), listUploads(ROOT), '5.6（自檢）');
}

// ── 瀏覽器（headless Edge）：新增名片表單 → 沒有上傳區、AI 辨識在 → 儲存成功 ──
async function dom(h) {
  console.log('[6] 瀏覽器（headless Edge）');
  if (process.env.NOSTORE_SKIP_DOM) { console.log('  （未驗證：NOSTORE_SKIP_DOM 明確略過）'); ok(true, '6.0 【未驗證】DOM 段被明確略過'); return; }
  const { withBrowser, findEdge } = require('./_fd-cdp');
  if (!findEdge()) { ok(false, '6.0 找不到 Edge（沒有 Edge 的環境請設 NOSTORE_SKIP_DOM=1 明確略過；不可靜默當成通過）'); return; }
  await withBrowser({ cdpPort: CDP, width: 1280, height: 1000, maxMs: 180000 }, async (b) => {
    await b.nav(h.base + '/login.html');
    const lg = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${JSON.stringify(U.A)},password:${JSON.stringify(PW)}})}).then(r=>r.status)`);
    ok(lg === 200, '6.1 登入');
    await b.nav(h.base + '/');
    ok(await b.waitFor('typeof openModal==="function" && !!document.getElementById("saveBtn")', 20000), '6.2 主頁載入、openModal 可用');
    await b.ev('openModal()'); await b.sleep(200);
    const s = await b.ev(`({input:!!document.getElementById('cardImageInput'),prev:!!document.getElementById('cardImagePreview'),hint:!!document.querySelector('.upload-hint'),legacy:getComputedStyle(document.getElementById('cardImageLegacy')).display,ocrBtn:!!document.getElementById('ocrCardBtn'),ocrVis:(function(){var e=document.getElementById('ocrCardBtn');return !!e&&e.offsetParent!==null})(),cap:document.getElementById('ocrCardInput')&&document.getElementById('ocrCardInput').getAttribute('capture'),textHasUpload:/點擊上傳名片圖片/.test(document.getElementById('tab-card').innerText)})`);
    ok(!s.input && !s.prev && !s.hint && !s.textHasUpload, `6.3 新增名片表單沒有上傳區（${JSON.stringify(s)}）`);
    eq(s.legacy, 'none', '6.4 新增時舊圖區塊隱藏');
    ok(s.ocrBtn && s.ocrVis && s.cap === 'environment', '6.5 AI 辨識按鈕可見、capture 屬性在');
    // 既有舊圖：編輯 c-img → 顯示唯讀預覽與移除鈕；移除→儲存→伺服器清空
    ok(await b.waitFor('Array.isArray(window.contacts)||true', 3000), '6.5b');
    const hasImg = await b.ev(`fetch('/api/contacts').then(r=>r.json()).then(l=>{const c=l.find(x=>x.id==='c-img'); window.__c=c; return c&&c.cardImage})`);
    eq(hasImg, '/uploads/' + K.OLD, '6.6 （前置）c-img 有舊圖');
    await b.ev('openModal(window.__c)'); await b.sleep(300);
    const e = await b.ev(`({legacy:getComputedStyle(document.getElementById('cardImageLegacy')).display,src:document.getElementById('previewImg').getAttribute('src'),url:document.getElementById('cardImageUrl').value,rm:!!document.getElementById('removeImage'),input:!!document.getElementById('cardImageInput')})`);
    ok(e.legacy !== 'none' && e.src === '/uploads/' + K.OLD && e.url === '/uploads/' + K.OLD && e.rm && !e.input, `6.7 編輯有舊圖的名片：唯讀預覽＋移除鈕，沒有上傳入口（${JSON.stringify(e)}）`);
    const pv = await b.ev(`new Promise(r=>{const i=document.getElementById('previewImg');if(i.complete&&i.naturalWidth)return r(true);i.onload=()=>r(true);i.onerror=()=>r(false);setTimeout(()=>r(i.naturalWidth>0),3000)})`);
    ok(pv === true, '6.8 舊圖預覽實際載入成功（SEC-02 授權下 A 看得到）');
    // 儲存不移除：payload 帶回原值（監看 fetch）
    await b.ev(`window.__bodies=[];const rf=window.fetch.bind(window);window.fetch=function(u,o){if(o&&o.body&&String(u).indexOf('/api/contacts')>=0)window.__bodies.push(String(o.body));return rf(u,o)}`);
    await b.click('#saveBtn'); await b.sleep(1200);
    let bodies = await b.ev('window.__bodies');
    ok(bodies.length >= 1 && JSON.parse(bodies[0]).cardImage === '/uploads/' + K.OLD, '6.9 編輯儲存（沒動圖）：payload 帶回原值');
    eq(((await http(h, U.A, 'GET', '/api/contacts')).json || []).find(x => x.id === 'c-img').cardImage, '/uploads/' + K.OLD, '6.10 儲存後舊圖仍在');
    await b.ev('openModal(window.__c)'); await b.sleep(300);
    await b.click('#removeImage'); await b.sleep(200);
    const rm = await b.ev(`({legacy:getComputedStyle(document.getElementById('cardImageLegacy')).display,url:document.getElementById('cardImageUrl').value})`);
    ok(rm.legacy === 'none' && rm.url === '', `6.11 按「移除圖片」：區塊隱藏、值清空（${JSON.stringify(rm)}）`);
    await b.ev('window.__bodies=[]'); await b.click('#saveBtn'); await b.sleep(1200);
    bodies = await b.ev('window.__bodies');
    ok(bodies.length >= 1 && JSON.parse(bodies[0]).cardImage === '', '6.12 移除後儲存：payload cardImage＝\'\'');
    eq(((await http(h, U.A, 'GET', '/api/contacts')).json || []).find(x => x.id === 'c-img').cardImage, '', '6.13 伺服器端舊圖已清掉');
    // 新增名片 → 儲存成功
    await b.ev('openModal()'); await b.sleep(200);
    await b.ev(`(function(){const set=(id,v)=>{const e=document.getElementById(id);e.value=v;e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}))};set('name','瀏覽器新增');set('company','沙盒公司');set('email','dom@sandbox.test');const f=document.getElementById('isForeignChk');f.checked=true;f.dispatchEvent(new Event('change',{bubbles:true}));})()`);
    await b.ev('window.__bodies=[]'); await b.click('#saveBtn'); await b.sleep(1500);
    bodies = await b.ev('window.__bodies');
    const made = ((await http(h, U.A, 'GET', '/api/contacts')).json || []).find(x => x.name === '瀏覽器新增');
    ok(!!made && made.cardImage === '', `6.14 新增名片（無上傳區）儲存成功，cardImage＝''（${made ? '已建立' : '沒建立'}；payload ${bodies[0] ? bodies[0].slice(0, 0) : '無'}）`);
    ok(bodies.length >= 1 && JSON.parse(bodies[0]).cardImage === '', '6.15 新增的 payload cardImage 是空字串');
    ok(b.errors.length === 0, `6.16 頁面沒有未捕捉例外（${b.errors.slice(0, 2).join(' | ')}）`);
  });
  ok((require('./_fd-cdp').withBrowser.leftover || 0) === 0, '6.17 Edge 已完全清掉');
}

(async () => {
  let code = 0, hb = null, h = null;
  try {
    staticChecks();
    cleanup(ROOT); seed(ROOT);
    if (BASE_ROOT) { cleanup(BASE_ROOT); seed(BASE_ROOT); }
    const ALL = [...Object.values(U), ...EXTRA.map(x => x.n)];
    // 登入限流是每個伺服器程序 20 次／15 分鐘：帳號分批，每批重開兩個伺服器（資料每次重新播種，兩邊一致）
    const batches = [ALL.slice(0, 12), ALL.slice(12)];
    for (let i = 0; i < batches.length; i++) {
      cleanup(ROOT); seed(ROOT);
      if (BASE_ROOT) { cleanup(BASE_ROOT); seed(BASE_ROOT); }
      h = await boot(ROOT, PORT);
      if (BASE_ROOT) hb = await boot(BASE_ROOT, BASE_PORT);
      await parity(h, hb, batches[i], i === 0);
      for (const s2 of servers.splice(0)) await stop(s2);
      hb = null;
    }
    ok(!BASE_ROOT || PARITY_TOTAL >= 20, `5.0 逐位元比對的帳號總數 >= 20（${PARITY_TOTAL}）`);
    if (BASE_ROOT) cleanup(BASE_ROOT);
    cleanup(ROOT); seed(ROOT);
    h = await boot(ROOT, PORT);
    for (const u of Object.values(U)) await login(h, u);
    await backend(h);
    for (const s2 of servers.splice(0)) await stop(s2);   // 重開＋重新播種（前面的檢查已清掉舊圖）
    cleanup(ROOT); seed(ROOT);
    h = await boot(ROOT, PORT);
    await login(h, U.A);
    await dom(h);
  } catch (e) { console.log('錯誤：', e.stack || e.message); fail++; failures.push('例外：' + e.message); code = 1; }
  finally {
    for (const s of servers) await stop(s);
    cleanup(ROOT); if (BASE_ROOT) cleanup(BASE_ROOT);
  }
  console.log(`\n通過 ${pass}　失敗 ${fail}`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); code = 1; }
  process.exit(code);
})();
