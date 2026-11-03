#!/usr/bin/env node
'use strict';
/**
 * 「靜態檔請求拿過期的帳號快取誤踢登入」修正檢查 —— 分支 fix/session-static-stale。
 *
 * 背景：server.js 的 db.ready()（stale check）只套 /api/；session-refresh 中介層卻對所有請求用記憶體快取判斷
 * 「帳號不存在／停用／密碼已變更」並強制登出。Vercel 多實例下，靜態檔請求（載入畫面約 40 個）會拿舊快取誤踢人。
 * 修法：只對「GET/HEAD 且對應 _client 內實際存在的檔案」豁免（lib/staticAssets.js）；/api/*、/uploads、/admin.html、其他動態路由不變。
 *
 *   A 單元：lib/staticAssets.js 判斷（含路徑穿越／編碼／大小寫／方法）
 *   B 真實 server.js（JSON 模式，埠 CS_PORT 預設 3520，臨時 _s_* 帳號、隨機密碼與 JWT 金鑰）
 *     B1 重現：用「大小與 mtime 都不變」的 auth.json 改寫讓 5 秒 TTL 的帳號快取真的過期（缺帳號／舊 passwordChangedAt／舊 active），
 *        靜態檔請求不得被登出。同一組探針也對 origin/main 的 server.js 跑一次，必須「重現 bug」（確認測試有效）
 *     B2 安全不退步：停用／刪除／改密碼（舊 token）→ /api/* 仍 401 SESSION_INVALID＋X-Session-Expired＋SESSION_REJECTED_INACTIVE 稽核；
 *        /uploads、不存在的 .js、/admin.html 仍照舊被拒（302）
 *     B3 未登入者的靜態檔／API 回應與 origin/main 逐項相同
 *   M 變異：每個都必須讓檢查變紅
 *
 *   node scripts/check-session-static-stale.js          （CS_SKIP_MUT=1 略過變異；CS_ONLY=A,B 選區）
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 一律拒絕；跑完會刪除它們。禁用埠 3000／3020／3030／3101-3161。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn, execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.CS_PORT || '3520', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3161)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const BASE = `http://localhost:${PORT}`;
const ONLY = process.env.CS_ONLY ? new Set(process.env.CS_ONLY.split(',').map(s => s.trim().toUpperCase())) : null;
const want = (k) => !ONLY || ONLY.has(k);
const DATA_FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const sleep = (n) => new Promise(r => setTimeout(r, n));

function mkT() {
  const T = { n: 0, f: 0, fl: [],
    ok(c, name) { if (c) T.n++; else { T.f++; T.fl.push(name); } },
    eq(a, b, name) { T.ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); } };
  return T;
}
const G = mkT();
const bind = (T) => ({ ok: (c, n) => { T.ok(c, n); if (!c) console.log('  ✗ FAIL:', n); }, eq: (a, b, n) => { T.eq(a, b, n); if (JSON.stringify(a) !== JSON.stringify(b)) console.log('  ✗ FAIL:', n, '得到', JSON.stringify(a), '預期', JSON.stringify(b)); } });

// ═════════ A：lib/staticAssets.js 單元 ═════════
function probeA(T, libFile) {
  const SA = require(libFile || path.join(ROOT, 'lib', 'staticAssets.js'));
  const sa = SA.create(path.join(ROOT, '_client'));
  const rq = (method, p) => ({ method, path: p });
  T.ok(sa.size() > 20, '[A] _client 檔名集合已建立（>20 個）');
  T.eq(sa.isStaticAsset(rq('GET', '/quote-coststeps.js')), true, '[A] GET 實際存在的 .js → 靜態');
  T.eq(sa.isStaticAsset(rq('HEAD', '/style.css')), true, '[A] HEAD 實際存在的 .css → 靜態');
  T.eq(sa.isStaticAsset(rq('GET', '/vendor/' + (fs.readdirSync(path.join(ROOT, '_client', 'vendor'))[0]))), true, '[A] 子目錄內的真實檔案 → 靜態');
  T.eq(sa.isStaticAsset(rq('POST', '/quote-coststeps.js')), false, '[A] POST → 不豁免');
  T.eq(sa.isStaticAsset(rq('PUT', '/quote-coststeps.js')), false, '[A] PUT → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/nonexistent-zzz.js')), false, '[A] 不存在的 .js（只看副檔名不行）→ 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/uploads/abc.js')), false, '[A] /uploads/* → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/api/me')), false, '[A] /api/* → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/admin.html')), false, '[A] /admin.html（有自己的角色檢查路由）→ 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/')), false, '[A] 根目錄 → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/vendor')), false, '[A] 目錄 → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/STYLE.CSS')), false, '[A] 大小寫不同 → 不放寬');
  T.eq(sa.isStaticAsset(rq('GET', '/../server.js')), false, '[A] .. 穿越 → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/vendor/../style.css')), false, '[A] 中間 .. → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/%2e%2e/server.js')), false, '[A] %2e%2e → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/vendor/%2E%2E/style.css')), false, '[A] 大寫 %2E → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/vendor%2fx.js')), false, '[A] %2f → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/vendor\\x.js')), false, '[A] 反斜線 → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/vendor%5cx.js')), false, '[A] %5c → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/style.css%00.js')), false, '[A] %00 → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/style.css\0')), false, '[A] 空位元組 → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '//style.css')), false, '[A] 空區段 → 不豁免');
  T.eq(sa.isStaticAsset(rq('GET', '/%E0%A4%A')), false, '[A] 壞掉的百分比編碼 → 不豁免（不丟例外）');
  T.eq(sa.isStaticAsset({ method: 'GET' }), false, '[A] 沒有 path → 不豁免');
}

// ═════════ B：真實伺服器 ═════════
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const jwtLib = require(path.join(ROOT, 'node_modules', 'jsonwebtoken'));
const rnd = (p) => p + crypto.randomBytes(9).toString('hex') + '!Z';
const PW = rnd('Cs');
const HASH = bcrypt.hashSync(PW, 6);
const P0 = '2026-01-01T00:00:00.000Z', P1 = '2026-02-02T00:00:00.000Z', P2 = '2026-03-03T00:00:00.000Z';
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: HASH, createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: ['ERP'] }, extra || {});
const AUTH_FILE = path.join(ROOT, 'auth.json');
let child = null, childPid = null, srvSecret = '';
function guardClean() {
  const g = DATA_FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map(f => path.basename(f)).join('/'), '，請在乾淨的 worktree 執行'); process.exit(2); }
}
function cleanupFiles() { DATA_FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }
function writeData() {
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify({ contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], groups: [], quotations: [] }, null, 2));
}
async function startServer(serverFile) {
  srvSecret = crypto.randomBytes(24).toString('hex');
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: srvSecret, SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'AUDIT_EVENTS_TABLE', 'LOGIN_LOCKOUT', 'SESSION_PWD_BINDING', 'LOGIN_LOCK_SCOPE']) delete env[k];
  child = spawn(process.execPath, [serverFile], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.CS_DEBUG) process.stdout.write('[srv] ' + d); };
  child.stdout.on('data', onData); child.stderr.on('data', onData);
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-500));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-500));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100);
  child = null;
}
const mkTok = (user) => { const s = Math.floor(Date.now() / 1000); return 'itts_auth=' + encodeURIComponent(jwtLib.sign({ user, lgn: s, lat: s, exp: s + 3600 }, srvSecret)); };
const tokOf = (username, role, pwa) => mkTok({ username, displayName: username, role: role || 'user', bu: ['ERP'], loginAt: Date.now(), pwa });
async function req(method, url, cookie) {
  const r = await fetch(BASE + url, { method, headers: cookie ? { Cookie: cookie } : {}, redirect: 'manual' });
  const buf = Buffer.from(await r.arrayBuffer());
  const sc = r.headers.getSetCookie ? r.headers.getSetCookie() : [];
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: r.status, location: r.headers.get('location') || '', expired: r.headers.get('x-session-expired') || '', cleared: sc.some(c => /^itts_auth=;|itts_auth=; /i.test(c) || /Expires=Thu, 01 Jan 1970/i.test(c)), sha: crypto.createHash('sha1').update(buf).digest('hex'), json, ct: r.headers.get('content-type') || '' };
}
const writeAuthStr = (s) => fs.writeFileSync(AUTH_FILE, s, 'utf8');

/** 讓帳號快取「真的過期」：先以舊檔填快取，再改寫成大小、mtime 都相同的新內容（5 秒 TTL 內）。回傳是否成功。 */
async function makeStale(initialUsers, finalUsers, warmCookie) {
  let fin = JSON.stringify({ users: finalUsers });
  let ini = JSON.stringify({ users: initialUsers });
  const L = Math.max(Buffer.byteLength(ini), Buffer.byteLength(fin));
  ini += ' '.repeat(L - Buffer.byteLength(ini)); fin += ' '.repeat(L - Buffer.byteLength(fin));   // 兩份補成等長
  const T0 = Math.floor(Date.now() / 1000) - 60;   // 整數秒的 mtime 可被精確還原（小數 mtime 無法）
  writeAuthStr(ini); fs.utimesSync(AUTH_FILE, T0, T0);
  const st0 = fs.statSync(AUTH_FILE);
  await sleep(50);
  const w = await req('GET', '/api/me', warmCookie);        // 以 adm1 填快取
  if (w.status !== 200) throw new Error('暖機失敗 ' + w.status);
  writeAuthStr(fin);
  fs.utimesSync(AUTH_FILE, T0, T0);
  const st1 = fs.statSync(AUTH_FILE);
  return st1.mtimeMs === st0.mtimeMs && st1.size === st0.size;
}

/** 探針（可對 origin/main 的 server.js 或修好的 server.js 跑）。回傳量測結果。 */
async function probeServer(serverFile) {
  const R = {};
  cleanupFiles(); writeData();
  const adm = mkUser('_s_adm1', 'admin', { passwordChangedAt: P0 });
  // 階段 1：快取過期（缺帳號／舊 passwordChangedAt／舊 active）
  const iniUsers = [adm, mkUser('_s_pwd', 'user', { passwordChangedAt: P0 }), mkUser('_s_act', 'user', { active: false })];
  const finUsers = [adm, mkUser('_s_pwd', 'user', { passwordChangedAt: P1 }), mkUser('_s_act', 'user', { active: true }), mkUser('_s_new', 'user')];
  // 讓 final 的 active:true 與 initial 的 active:false 等長：序列化後 'true' 補一個空白不影響 JSON，但 JSON.stringify 不會補；改由 makeStale 的 padding 處理總長
  await startServer(serverFile);
  try {
    R.sameKey = await makeStale(iniUsers, finUsers, tokOf('_s_adm1', 'admin', P0));
    const probes = {
      missing: tokOf('_s_new', 'user', undefined),
      pwd: tokOf('_s_pwd', 'user', P1),
      inactive: tokOf('_s_act', 'user', undefined),
    };
    R.stale = {};
    for (const [k, ck] of Object.entries(probes)) {
      R.stale[k] = {};
      for (const p of ['/quote-coststeps.js', '/company-dups.js']) R.stale[k][p] = await req('GET', p, ck);
      R.stale[k].head = await req('HEAD', '/style.css', ck);
    }
    // 階段 2：安全不退步（auth.json 換成真正的狀態：停用／刪除／改密碼）
    const sec = [adm, mkUser('_s_dis', 'user', { active: false }), mkUser('_s_pw2', 'user', { passwordChangedAt: P2 })];   // _s_del 不存在
    writeAuthStr(JSON.stringify({ users: sec }, null, 2));
    const st = fs.statSync(AUTH_FILE); fs.utimesSync(AUTH_FILE, new Date(), new Date(st.mtimeMs + 3000));
    await sleep(100);
    const ck = { dis: tokOf('_s_dis', 'user', undefined), del: tokOf('_s_del', 'user', undefined), pw2: tokOf('_s_pw2', 'user', P1) };
    R.api = {};
    for (const k of Object.keys(ck)) R.api[k] = await req('GET', '/api/me', ck[k]);
    R.dyn = {};
    for (const k of Object.keys(ck)) {
      R.dyn[k] = {
        uploads: await req('GET', '/uploads/zzz', ck[k]),
        uploadsJs: await req('GET', '/uploads/zzz.js', ck[k]),
        ghostJs: await req('GET', '/nonexistent-zzz.js', ck[k]),
        admin: await req('GET', '/admin.html', ck[k]),
        asset: await req('GET', '/quote-coststeps.js', ck[k]),
        post: await req('POST', '/quote-coststeps.js', ck[k]),
      };
    }
    await sleep(400);
    const lg = await req('GET', '/api/admin/logs?limit=500', tokOf('_s_adm1', 'admin', P0));
    const arr = Array.isArray(lg.json) ? lg.json : ((lg.json && (lg.json.logs || lg.json.rows)) || []);
    R.audit = arr.filter(l => l.action === 'SESSION_REJECTED_INACTIVE').map(l => l.target).sort();
    // 階段 3：未登入者
    R.unauth = {};
    for (const p of ['/quote-coststeps.js', '/company-dups.js', '/style.css', '/index.html', '/app.js', '/uploads/zzz', '/admin.html', '/nonexistent-zzz.js', '/vendor/x', '/', '/login.html', '/api/me', '/api/quotations']) {
      const r = await req('GET', p, null);
      R.unauth[p] = { s: r.status, loc: r.location, exp: r.expired, sha: r.sha, ct: r.ct.split(';')[0] };
    }
    const hp = await req('HEAD', '/quote-coststeps.js', null); R.unauth.head = { s: hp.status, loc: hp.location };
    const po = await req('POST', '/quote-coststeps.js', null); R.unauth.post = { s: po.status, loc: po.location, sha: po.sha };
  } finally { await stopServer(); cleanupFiles(); }
  return R;
}

function assertFixed(T, R, tag) {
  T.ok(R.sameKey, `${tag}前提：auth.json 改寫後 mtime／大小與快取鍵相同（快取確實過期）`);
  for (const k of ['missing', 'pwd', 'inactive']) {
    for (const p of ['/quote-coststeps.js', '/company-dups.js']) {
      const r = R.stale[k][p];
      T.eq([r.status, r.cleared], [200, false], `${tag}[B1] 快取${{ missing: '缺帳號', pwd: '舊密碼版本', inactive: '舊 active=false' }[k]}：GET ${p} 不被登出（200、不清 cookie）`);
    }
    T.eq(R.stale[k].head.status, 200, `${tag}[B1] 快取過期：HEAD /style.css → 200`);
  }
  for (const k of ['dis', 'del', 'pw2']) {
    const a = R.api[k];
    T.eq([a.status, a.json && a.json.code, a.expired], [401, 'SESSION_INVALID', '1'], `${tag}[B2] ${k}：/api/me → 401 SESSION_INVALID＋X-Session-Expired`);
    const d = R.dyn[k];
    T.eq(d.uploads.status, 302, `${tag}[B2] ${k}：/uploads/* 仍照舊被拒（302）`);
    T.eq(d.uploadsJs.status, 302, `${tag}[B2] ${k}：/uploads/*.js 仍照舊被拒（302，不能只看副檔名）`);
    T.eq(d.ghostJs.status, 302, `${tag}[B2] ${k}：不存在的 .js 仍照舊被拒（302）`);
    T.eq(d.admin.status, 302, `${tag}[B2] ${k}：/admin.html 仍照舊被拒（302）`);
    T.eq([d.post.status, d.post.cleared], [302, true], `${tag}[B2] ${k}：POST 靜態路徑 不豁免（照舊 302 並清除登入 cookie）`);
    T.eq(d.asset.status, 200, `${tag}[B2] ${k}：真實靜態檔（GET）依新規格放行 200（只是 JS 檔，不含資料；API 仍擋）`);
  }
  T.eq(R.audit, ['_s_del', '_s_dis', '_s_pw2'], `${tag}[B2] 三個被擋的帳號都寫了 SESSION_REJECTED_INACTIVE 稽核`);
}

async function probeB(T, serverFile, origFile) {
  const R = await probeServer(serverFile);
  assertFixed(T, R, '');
  const O = await probeServer(origFile);
  // 重現 bug：origin/main 版對過期快取的靜態檔請求會登出（302 或清 cookie）
  for (const k of ['missing', 'pwd', 'inactive']) {
    const r = O.stale[k]['/quote-coststeps.js'];
    T.ok(O.sameKey && (r.status !== 200 || r.cleared), `[B1 原版] 快取過期（${k}）→ origin/main 確實把靜態檔請求登出（重現 bug；得到 ${r.status}）`);
  }
  // 原版的安全行為也量一次，作為「不退步」的對照
  for (const k of ['dis', 'del', 'pw2']) {
    T.eq([R.api[k].status, R.api[k].json && R.api[k].json.code], [O.api[k].status, O.api[k].json && O.api[k].json.code], `[B2 對照] ${k}：/api/me 與 origin/main 相同`);
    for (const f of ['uploads', 'uploadsJs', 'ghostJs', 'admin', 'post']) T.eq(R.dyn[k][f].status, O.dyn[k][f].status, `[B2 對照] ${k}：${f} 與 origin/main 相同`);
  }
  T.ok(['_s_del', '_s_dis', '_s_pw2'].every(u => O.audit.includes(u)), '[B2 對照] origin/main 對這三個帳號同樣寫了稽核（原版另外還有 3 筆誤踢紀錄＝bug 的副作用）');
  T.eq(R.unauth, O.unauth, '[B3] 未登入者的靜態檔／API 回應（狀態、Location、標頭、內容雜湊）與 origin/main 逐項相同');
  return R;
}

// ═════════ M：變異 ═════════
const EX_LINE = 'if (staticAssets.isStaticAsset(req)) return next();';
async function runMut(name, edits, origFile) {
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  let out = src;
  for (const [a, b] of edits) { if (!out.includes(a)) { console.log('  ✗ 變異字串找不到：', name, a.slice(0, 60)); G.ok(false, `[M] ${name}：變異字串存在`); return; } out = out.replace(a, () => b); }
  const mf = path.join(ROOT, '_s_mut_server.js');
  fs.writeFileSync(mf, out);
  const T = mkT();
  try { const R = await probeServer(mf); assertFixed(T, R, ''); }
  catch (e) { T.ok(false, '例外：' + e.message.slice(0, 120)); }
  finally { try { fs.unlinkSync(mf); } catch { /* */ } }
  const red = T.f > 0;
  console.log(`  ${red ? '✓' : '✗'} 變異「${name}」${red ? '被抓到（' + T.f + ' 項失敗）' : '沒有被抓到！'}`);
  G.ok(red, `[M] 變異「${name}」必須讓檢查變紅`);
}

(async () => {
  guardClean();
  const t0 = Date.now();
  const origFile = path.join(ROOT, '_s_orig_server.js');
  try {
    if (want('A')) { console.log('[A] 靜態資產判斷（單元）'); probeA(bind(G), null); }
    if (want('B') || want('M')) {
      fs.writeFileSync(origFile, execFileSync('git', ['show', 'origin/main:server.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024 }));
    }
    if (want('B')) { console.log(`[B] 真實伺服器（埠 ${PORT}）`); await probeB(bind(G), path.join(ROOT, 'server.js'), origFile); }
    if (want('M') && !process.env.CS_SKIP_MUT) {
      console.log('[M] 變異檢查');
      await runMut('拿掉豁免', [[EX_LINE, '']]);
      await runMut('豁免擴大到非 /api/', [[EX_LINE, "if (!req.path.startsWith('/api/')) return next();"]]);
      await runMut('豁免擴大到 /api/（含 API 全放行）', [[EX_LINE, "if (req.path.startsWith('/api/') || staticAssets.isStaticAsset(req)) return next();"]]);
      await runMut('只看副檔名 .js、不驗檔案存在', [[EX_LINE, "if (/\\.js$/.test(req.path)) return next();"]]);
      await runMut('/uploads 也被豁免', [[EX_LINE, "if (staticAssets.isStaticAsset(req) || req.path.startsWith('/uploads/')) return next();"]]);
      await runMut('不限方法（POST 也豁免）', [[EX_LINE, "if (staticAssets.isStaticAsset({ method: 'GET', path: req.path })) return next();"]]);
      await runMut('拿掉 /api/ 的拒絕邏輯', [["return res.status(401).json({ error: '帳號已停用或不存在，請重新登入', code: 'SESSION_INVALID' });", 'return next();']]);
      await runMut('拿掉稽核寫入', [["    _logSessionRejected(req, su.username, bad);\n", '']]);
    }
  } catch (e) { console.log('例外：', e && e.stack || e); G.ok(false, '腳本例外：' + (e && e.message)); }
  finally { await stopServer(); cleanupFiles(); try { fs.unlinkSync(origFile); } catch { /* */ } }
  console.log(`\ncheck-session-static-stale 結果：${G.n} 通過，${G.f} 失敗（${((Date.now() - t0) / 1000).toFixed(1)} 秒）`);
  if (G.f) { G.fl.forEach(x => console.log('  ✗', x)); process.exit(1); }
  console.log('全部通過');
})();
