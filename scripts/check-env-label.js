#!/usr/bin/env node
'use strict';
/**
 * 測試環境標示（APP_ENV_LABEL）檢查：單元（lib/envLabel.js）＋ 真實 server.js 行為 ＋ 未設定時與「修改前」位元級相同
 * ＋ 變異測試 ＋ 瀏覽器（headless Edge + CDP）。不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體／本次執行的資料檔。
 *
 *   node scripts/check-env-label.js                    單元＋伺服器＋位元級比對（不含瀏覽器）
 *   ENVLABEL_ONLY=unit,http,off,ui node ...            只跑指定群組（ui 需要 Edge）
 *   ENVLABEL_MUT=1 node ...                            另外跑變異測試（每個變異都必須讓對應群組變紅）
 *   ENVLABEL_SHOTS=<資料夾> node ...                   ui 群組輸出截圖
 *   ENVLABEL_PORT（預設 3490，基準用 +1）／ENVLABEL_CDP（預設 9690）／ENVLABEL_BASE_REF（基準提交，預設固定 8220199＝功能加入前的 main，不用會前進的 origin/main）
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行；跑完會刪掉自己建立的這些檔。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const Module = require('module');
const { spawn, spawnSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.ENVLABEL_PORT || '3490', 10);
const CDP = parseInt(process.env.ENVLABEL_CDP || '9690', 10);
if (PORT < 3490 || PORT > 3498) { console.log('拒絕：測試埠限 3490–3498'); process.exit(2); }
const BASE_REF = process.env.ENVLABEL_BASE_REF || '8220199';
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'];
{
  const g = ['auth.json', 'data.json', 'audit.log.json'].filter((f) => fs.existsSync(path.join(ROOT, f)));
  if (g.length) { console.log('拒絕：此工作樹已有', g.join('/'), '（可能是真實資料）'); process.exit(2); }
}
const ONLY = (process.env.ENVLABEL_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length ? k !== 'ui' : ONLY.includes(k);
const SHOTS = process.env.ENVLABEL_SHOTS || '';
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── 單元 ─────────────────
function loadLib(file) { return file ? require(file) : require(path.join(ROOT, 'lib', 'envLabel.js')); }
function groupUnit(L) {
  console.log('[單元] getEnvLabel / injectEnvLabel / 用戶端腳本');
  const g = (v) => L.getEnvLabel({ APP_ENV_LABEL: v });
  eq(L.getEnvLabel(undefined), '', 'env 為 undefined → 空');
  eq(L.getEnvLabel({}), '', '未設定 → 空');
  eq(g(''), '', '空字串 → 空');
  eq(g('   '), '', '只有空白 → 空');
  eq(g('DEV'), 'DEV', 'DEV → DEV');
  eq(g('  DEV  '), 'DEV', '前後空白被修掉');
  eq(g('UAT-1_a b'), 'UAT-1_a b', '允許 - _ 與中間空白');
  eq(g('A    B'), 'A B', '連續空白合併');
  eq(g('<script>alert(1)</script>'), 'scriptalert1', 'XSS 字串：剝除 < > ( ) /，截斷到 12');
  eq(g('"><img src=x onerror=1>'), 'img srcx one', '屬性跳脫嘗試：剝除引號與等號後截斷');
  ok(!/[<>"'&=()\/\\]/.test(g('"><svg/onload=alert(1)>')), '輸出不含任何危險字元');
  eq(g('DEV'.repeat(50)).length, 12, '超長 → 12 字');
  eq(g('中文測試'), '', '全中文 → 空（視為未設定）');
  eq(g('測試DEV環境'), 'DEV', '混合中文：只留英數');
  eq(g('DEV😀'), 'DEV', 'emoji 被剝除');
  eq(g('ＤＥＶ'), '', '全形英文被剝除 → 空');
  eq(g('---'), '', '只剩符號 → 空');
  eq(g('___'), '', '只有底線 → 空');
  eq(g('DEV‮\u0000\n'), 'DEV', '控制字元／方向覆寫字元被剝除');
  eq(g('A'.repeat(11) + ' B'), 'AAAAAAAAAAA', '截斷後尾端空白被修掉');
  eq(g(123), '', '非字串 → 空');
  eq(g(null), '', 'null → 空');
  eq(g({ toString() { return 'DEV'; } }), '', '物件 → 空（不呼叫 toString）');
  eq(g(['DEV']), '', '陣列 → 空');
  const html = '<html><body><p>x</p></body></html>';
  eq(L.injectEnvLabel(html, ''), html, '空標示：HTML 原樣回傳');
  ok(L.injectEnvLabel(html, '') === html, '空標示：同一個字串');
  const out = L.injectEnvLabel(html, 'DEV');
  ok(out.indexOf('<script src="/env-label.js" data-env="DEV"></script>') > 0 && out.indexOf('</body>') > out.indexOf('data-env'), '注入在 </body> 之前');
  eq(out.split('data-env').length - 1, 1, '只注入一次');
  eq(L.injectEnvLabel(html, '"><script>x</script>'), html, '不合格式的標示不注入（防禦性檢查）');
  eq(L.injectEnvLabel('<p>no body</p>', 'DEV'), '<p>no body</p><script src="/env-label.js" data-env="DEV"></script>\n', '沒有 </body> 時附在最後');
  ok(L.injectEnvLabel('a$&b</body>', 'DEV').startsWith('a$&b<script'), '$ 與 & 不被當成取代樣式');
  ok(typeof L.CLIENT_JS === 'string' && !/innerHTML|outerHTML|insertAdjacentHTML|document\.write|eval\(/.test(L.CLIENT_JS), '用戶端腳本不含 HTML sink');
  ok(/textContent/.test(L.CLIENT_JS) && /createElement/.test(L.CLIENT_JS), '用戶端腳本用 createElement＋textContent');
  try { new Function(L.CLIENT_JS); ok(true, '用戶端腳本語法正確'); } catch (e) { ok(false, '用戶端腳本語法正確', e.message); }
  // 對比度：#1f1300 在 #f59e0b 上（WCAG）
  const lum = (h) => { const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); }; const n = parseInt(h, 16); return 0.2126 * f(n >> 16) + 0.7152 * f((n >> 8) & 255) + 0.0722 * f(n & 255); };
  const cr = (lum('f59e0b') + 0.05) / (lum('1f1300') + 0.05);
  ok(cr >= 7, '標示文字對比度 ≥ 7:1（實際 ' + cr.toFixed(1) + '）');
}

// ───────────────── 伺服器 ─────────────────
const PW = 'El' + crypto.randomBytes(9).toString('hex') + '!Z';
const nowIso = new Date().toISOString();
const PRELOAD = path.join(os.tmpdir(), 'envlabel_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.ENVLABEL_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.ENVLABEL_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');

function seedFiles(dir) {
  const u = (username, role) => ({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] });
  fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users: [u('_s_adm', 'admin'), u('_s_u1', 'sales')] }, null, 2));
  fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify({ contacts: [], companies: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [], opportunities: [] }, null, 2));
}
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(path.join(dir, f)); } catch { /* 不存在 */ } }); }

let child = null, childPid = null;
async function startServer(root, port, label, mut) {
  const env = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${port}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE', 'APP_ENV_LABEL', 'ENVLABEL_MUT_SPEC']) delete env[k];
  if (label !== undefined) env.APP_ENV_LABEL = label;
  if (mut && mut.length) { env.ENVLABEL_MUT_SPEC = JSON.stringify(mut); env.ENVLABEL_ROOT = root; }
  await assertPortFree(port);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 100; i++) {
    await sleep(250);
    try { const r = await fetch(`http://localhost:${port}/login.html`); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-600));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-600));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 30 && child.exitCode === null; i++) await sleep(100);
  child = null; await sleep(300);
}
async function fetchAt(base, cookie, url) {
  const res = await fetch(base + url, { headers: cookie ? { Cookie: cookie } : {}, redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  const headers = {}; res.headers.forEach((v, k) => { if (!['date', 'set-cookie', 'etag', 'last-modified', 'content-length'].includes(k)) headers[k] = v; });
  return { status: res.status, headers, text: buf.toString('utf8') };
}
async function loginCookie(base, who) {
  const res = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status !== 200) throw new Error('登入失敗 ' + who + ' ' + res.status);
  return (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
}
// 建置版號（?v=毫秒）與 /api/me 內的登入時間戳每次啟動都不同，比對前先正規化
const norm = (s) => s.replace(/\?v=\d{10,}/g, '?v=<build>').replace(/\d{4}-\d{2}-\d{2}T[\d:.]+Z/g, '<ts>').replace(/("(?:iat|exp|loginAt|lastActive|at)":)\d{10,}/g, '$1<n>');
const URLS_ANON = ['/login.html', '/env-label.js', '/login-boot.js', '/api/env-label', '/deep-link.js'];
const URLS_AUTH = ['/', '/index.html', '/admin.html', '/help.html', '/env-label.js', '/api/me', '/style.css'];
async function snapshot(base) {
  const out = {};
  for (const u of URLS_ANON) { const r = await fetchAt(base, '', u); out['anon ' + u] = { status: r.status, headers: r.headers, text: norm(r.text) }; }
  const adm = await loginCookie(base, '_s_adm');
  for (const u of URLS_AUTH) { const r = await fetchAt(base, adm, u); out['adm ' + u] = { status: r.status, headers: r.headers, text: norm(r.text) }; }
  const r = await fetchAt(base, '', '/sw.js'); out['sw'] = { status: r.status, text: r.text };
  return out;
}
const TAG = (l) => `<script src="/env-label.js" data-env="${l}"></script>`;

async function groupHttp(root, port, mut) {
  console.log('[伺服器] 設 DEV／設惡意值');
  cleanupFiles(root); seedFiles(root);
  try {
    await startServer(root, port, 'DEV', mut);
    const base = `http://localhost:${port}`;
    const adm = await loginCookie(base, '_s_adm');
    const lg = await fetchAt(base, '', '/login.html');
    ok(lg.status === 200 && lg.text.split(TAG('DEV')).length === 2, 'DEV：登入頁含一個標示腳本標籤', lg.status);
    ok(lg.text.indexOf(TAG('DEV')) < lg.text.lastIndexOf('</body>'), 'DEV：登入頁標籤在 </body> 之前');
    ok(/no-store|max-age=0/.test(lg.headers['cache-control'] || ''), 'DEV：登入頁不被快取');
    for (const u of ['/', '/index.html', '/admin.html']) {
      const r = await fetchAt(base, adm, u);
      ok(r.status === 200 && r.text.split(TAG('DEV')).length === 2, `DEV：${u} 含一個標示腳本標籤`, r.status);
    }
    const js = await fetchAt(base, '', '/env-label.js');
    ok(js.status === 200 && /javascript/.test(js.headers['content-type'] || ''), 'DEV：/env-label.js 未登入可取得（登入頁需要）', js.status);
    ok(js.text === loadLib().CLIENT_JS && !/DEV/.test(js.text.replace(/data-env/g, '')), 'DEV：腳本內容固定、不含標示值（值只在 HTML 屬性）');
    ok((await fetchAt(base, '', '/api/env-label')).status !== 200, 'DEV：不存在公開 JSON 端點（/api/env-label 非 200）');
    const anonIdx = await fetchAt(base, '', '/index.html');
    ok(anonIdx.status === 302 && !/data-env/.test(anonIdx.text), 'DEV：未登入的主畫面仍導向登入、不外洩內容', anonIdx.status);
    await stopServer();
    await startServer(root, port, '"><script>alert(1)</script>', mut);
    const lg2 = await fetchAt(base, '', '/login.html');
    const m = lg2.text.match(/<script src="\/env-label\.js" data-env="([^"]*)"><\/script>/);
    ok(!!m && /^[A-Za-z0-9_\- ]{1,12}$/.test(m[1]) && m[1] === 'scriptalert1', '惡意值：標籤存在且屬性值已清理為 scriptalert1', m && m[1]);
    ok(lg2.text.indexOf('alert(1)') < 0, '惡意值：登入頁不含 alert(1) 字樣');
    const idx2 = await fetchAt(base, await loginCookie(base, '_s_adm'), '/');
    const m2 = idx2.text.match(/data-env="([^"]*)"/g) || [];
    ok(m2.length === 1 && m2[0] === 'data-env="scriptalert1"' && idx2.text.indexOf('alert(1)') < 0, '惡意值：主畫面同樣只有清理後的單一標籤', m2.join(','));
    await stopServer();
    await startServer(root, port, '中文', mut);   // 清理後為空 → 等同未設定
    const lg3 = await fetchAt(base, '', '/login.html');
    ok(lg3.status === 200 && lg3.text.indexOf('data-env') < 0 && (await fetchAt(base, '', '/env-label.js')).status !== 200, '清理後為空：視為未設定（無標籤、無腳本路由）');
  } finally { await stopServer(); cleanupFiles(root); }
}

function exportBase() {
  if (process.env.ENVLABEL_BASE_DIR) return { dir: process.env.ENVLABEL_BASE_DIR, cleanup() {} };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'envlabel-base-'));
  const a = spawnSync('git', ['archive', BASE_REF], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + BASE_REF + ' 失敗：' + String(a.stderr));
  const t = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗：' + String(t.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function groupOff(root, port, mut) {
  console.log('[位元級] 未設定 APP_ENV_LABEL ＝ 與基準提交 ' + BASE_REF + ' 的回應相同');
  const base = exportBase();
  let a, b;
  try {
    cleanupFiles(base.dir); seedFiles(base.dir);
    await startServer(base.dir, port + 1, undefined, null);
    a = await snapshot(`http://localhost:${port + 1}`);
    await stopServer(); cleanupFiles(base.dir);
    cleanupFiles(root); seedFiles(root);
    await startServer(root, port, undefined, mut);
    b = await snapshot(`http://localhost:${port}`);
  } finally { await stopServer(); cleanupFiles(root); cleanupFiles(base.dir); base.cleanup(); }
  ok(Object.keys(a).length === Object.keys(b).length && Object.keys(a).length >= 12, '比對項目數 ' + Object.keys(a).length);
  for (const k of Object.keys(a)) {
    if (k === 'sw') {
      const strip = (s) => s.replace(/const SW_VERSION = '[^\n]*\n/, '');
      ok(strip(a.sw.text) === strip(b.sw.text) && a.sw.text !== b.sw.text, 'sw.js：除 SW_VERSION 那一行外相同，且版本確實有變');
      continue;
    }
    ok(a[k].status === b[k].status, k + ' 狀態碼相同', a[k].status + ' vs ' + b[k].status);
    ok(JSON.stringify(a[k].headers) === JSON.stringify(b[k].headers), k + ' 標頭相同', JSON.stringify(a[k].headers) + ' vs ' + JSON.stringify(b[k].headers));
    ok(a[k].text === b[k].text, k + ' 內容逐位元相同', a[k].text.length + ' vs ' + b[k].text.length);
  }
  ok(!/env-label|data-env/.test(b['anon /login.html'].text + b['adm /index.html'].text + b['adm /admin.html'].text), '未設定：三個頁面都沒有 env-label 痕跡');
  ok(b['anon /env-label.js'].status === a['anon /env-label.js'].status && b['adm /env-label.js'].status === a['adm /env-label.js'].status, '未設定：/env-label.js 行為與基準相同（未登入與登入）');
}

// ───────────────── 變異 ─────────────────
const MUTANTS = [
  { name: '拿掉字元清理', group: 'unit+http', spec: [{ file: 'lib/envLabel.js', edits: [{ find: ".replace(/[^A-Za-z0-9_\\- ]/g, '')", rep: '' }] }] },
  { name: '拿掉「清理後為空＝未設定」判斷', group: 'unit', spec: [{ file: 'lib/envLabel.js', edits: [{ find: "return /[A-Za-z0-9]/.test(s) ? s : '';", rep: 'return s;' }] }] },
  { name: '未設定時仍改寫 HTML（寫死 DEV）', group: 'off', spec: [{ file: 'server.js', edits: [{ find: 'res.send(ENV_LABEL ? envLabel.injectEnvLabel(versioned, ENV_LABEL) : versioned);', rep: "res.send(envLabel.injectEnvLabel(versioned, 'DEV'));" }] }] },
  { name: '未設定時仍註冊公開腳本路由／登入頁改寫', group: 'off', spec: [{ file: 'server.js', edits: [{ find: 'if (ENV_LABEL) {\n  app.get(\'/env-label.js\'', rep: 'if (true) {\n  app.get(\'/env-label.js\'' }] }] },
];
async function runMutants(root, port) {
  console.log('[變異] 每個變異都必須讓對應群組變紅');
  for (const m of MUTANTS) {
    const snap = { pass, fail, n: failures.length };
    quiet = true;
    let threw = '';
    try {
      if (m.group.includes('unit')) {
        // 單元：把變異後的原始碼寫成暫存模組再跑
        let src = fs.readFileSync(path.join(root, 'lib', 'envLabel.js'), 'utf8');
        for (const sp of m.spec) if (sp.file === 'lib/envLabel.js') for (const e of sp.edits) { if (!src.includes(e.find)) throw new Error('錨點不存在：' + e.find); src = src.replace(e.find, () => e.rep); }
        const tmp = path.join(os.tmpdir(), 'envlabel_mut_' + process.pid + '.js');
        fs.writeFileSync(tmp, src); delete require.cache[tmp];
        try { groupUnit(require(tmp)); } finally { try { fs.unlinkSync(tmp); } catch { /* 忽略 */ } }
      }
      if (m.group.includes('http')) await groupHttp(root, port, m.spec);
      if (m.group.includes('off')) await groupOff(root, port, m.spec);
    } catch (e) { threw = String(e && e.message || e).slice(0, 120); }
    const newFails = failures.length - snap.n;
    quiet = false;
    const red = newFails > 0 || /ANCHOR|伺服器/.test(threw) && newFails > 0;
    pass = snap.pass; fail = snap.fail; failures.length = snap.n;
    if (threw && newFails === 0) console.log('  （變異執行例外：' + threw + '）');
    ok(red, `變異「${m.name}」→ 測試變紅（${newFails} 項失敗）`);
  }
}

// ───────────────── 瀏覽器 ─────────────────
async function groupUi(root, port) {
  const { withBrowser } = require('./_fd-cdp');
  console.log('[瀏覽器] 未設／DEV × 三個頁面 × 深色／手機');
  const J = JSON.stringify;
  const state = (b) => b.ev(`(()=>{const e=document.getElementById('itts-env-label');const r=e&&e.lastChild?e.lastChild.getBoundingClientRect():null;return {has:!!e,text:e?e.textContent:'',title:document.title,tags:document.querySelectorAll('[data-env]').length,sw:document.documentElement.scrollWidth,iw:window.innerWidth,pill:r?{x:r.left,r:r.right,b:r.bottom}:null}})()`);
  const shot = async (b, name) => { if (SHOTS) await b.shot(path.join(SHOTS, name)); };
  for (const label of [undefined, 'DEV']) {
    const tag = label ? 'DEV' : '未設';
    cleanupFiles(root); seedFiles(root);
    try {
      await startServer(root, port, label, null);
      const base = `http://localhost:${port}`;
      await withBrowser({ cdpPort: CDP, width: 1440, height: 900 }, async (b) => {
        // 登入頁
        await b.nav(base + '/login.html');
        await sleep(1500);
        let s = await state(b);
        if (label) { ok(s.has && s.text === 'DEV' && s.tags === 1, `${tag}：登入頁顯示標示`, J(s)); ok(s.title.startsWith('[DEV] '), `${tag}：登入頁分頁標題有 [DEV] 前綴`, s.title); }
        else { ok(!s.has && s.tags === 0 && !s.title.startsWith('['), `${tag}：登入頁無任何標示節點、標題不變`, J(s)); }
        await shot(b, `${tag}-login-desktop.png`);
        await b.viewport(375, 800, true); await sleep(500);
        s = await state(b);
        ok(s.sw <= s.iw, `${tag}：登入頁手機寬度 375 不橫向溢出`, `${s.sw}>${s.iw}`);
        if (label) ok(s.pill && s.pill.x >= 0 && s.pill.r <= 375, `${tag}：登入頁手機版標示在畫面內`, J(s.pill));
        await shot(b, `${tag}-login-mobile.png`);
        await b.viewport(1440, 900, false);
        // 登入後
        const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:'_s_adm',password:${J(PW)}})}).then(r=>r.status)`);
        ok(st === 200, `${tag}：瀏覽器登入`);
        for (const [pg, url, wait] of [['main', '/', 'document.readyState==="complete"'], ['admin', '/admin.html', 'document.readyState==="complete"']]) {
          await b.ev(`sessionStorage.setItem('justLoggedIn','1'); 0`);
          await b.nav(base + url); await sleep(2500);
          s = await state(b);
          if (label) { ok(s.has && s.text === 'DEV' && s.tags === 1, `${tag}：${pg} 顯示標示`, J(s)); ok(s.title.startsWith('[DEV] ') && s.title.indexOf('[DEV] [DEV]') < 0, `${tag}：${pg} 標題前綴且不重複`, s.title); }
          else ok(!s.has && s.tags === 0 && !s.title.startsWith('['), `${tag}：${pg} 無任何標示節點、標題不變`, J(s));
          await shot(b, `${tag}-${pg}-desktop.png`);
          if (label) {
            await b.ev(`document.body.classList.add('dark'); 0`); await sleep(300);
            await shot(b, `${tag}-${pg}-dark.png`);
            await b.ev(`document.body.classList.remove('dark'); 0`);
          }
          const base375 = await (async () => { await b.viewport(375, 800, true); await sleep(600); return state(b); })();
          if (label) ok(base375.pill && base375.pill.x >= 0 && base375.pill.r <= 375, `${tag}：${pg} 手機版標示在畫面內`, J(base375.pill));
          await shot(b, `${tag}-${pg}-mobile.png`);
          // 溢出：與「未設」基準比較（標示不可造成新增溢出）→ 這裡先記錄，兩輪比對
          globalThis.__sw = globalThis.__sw || {}; globalThis.__sw[(label ? 'D' : 'U') + pg] = base375.sw - base375.iw;
          await b.viewport(1440, 900, false);
        }
        if (label) {  // 標題被頁面改寫後仍保持前綴
          await b.ev(`document.title='別的標題'; 0`); await sleep(300);
          ok((await state(b)).title === '[DEV] 別的標題', 'DEV：頁面改寫標題後前綴自動補回');
        }
        ok(b.errors.filter((e) => /env-label|itts-env/.test(e)).length === 0, `${tag}：無標示相關的 JS 錯誤`, b.errors.join(' | '));
      });
    } finally { await stopServer(); cleanupFiles(root); }
  }
  for (const pg of ['main', 'admin']) ok((globalThis.__sw['D' + pg]) <= (globalThis.__sw['U' + pg]), `手機版 ${pg}：設 DEV 後橫向溢出量不大於未設（${globalThis.__sw['D' + pg]} vs ${globalThis.__sw['U' + pg]}）`);
}

(async () => {
  try {
    if (want('unit')) groupUnit(loadLib());
    if (want('http')) await groupHttp(ROOT, PORT, null);
    if (want('off')) await groupOff(ROOT, PORT, null);
    if (want('ui')) await groupUi(ROOT, PORT);
    if (process.env.ENVLABEL_MUT === '1') await runMutants(ROOT, PORT);
  } catch (e) {
    fail++; failures.push('例外：' + (e && e.stack || e)); console.log('例外：', e && e.stack || e);
  } finally {
    await stopServer(); cleanupFiles(ROOT);
    try { fs.unlinkSync(PRELOAD); } catch { /* 不存在 */ }
  }
  console.log(`\n測試環境標示檢查：PASS ${pass} / FAIL ${fail}`);
  if (fail) console.log('失敗項目：\n  ' + failures.join('\n  '));
  process.exit(fail ? 1 : 0);
})();
