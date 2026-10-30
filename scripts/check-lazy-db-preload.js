#!/usr/bin/env node
'use strict';
/**
 * Postgres 懶載入預熱檢查：冷啟動不連資料庫、第一個 /api/ 請求才 ready 並只跑一次遷移。
 * 用「假 pg Pool」（記錄連線／查詢次數）在子行程裡啟動真實 server.js（DB_BACKEND=postgres），不連任何遠端資料庫。
 *
 *   node scripts/check-lazy-db-preload.js            全跑：基準重現＋5 個情境＋JSON 模式逐位元比對＋變異測試
 *   LAZY_ONLY=base,scen,json,mut node ...            只跑指定群組
 *   LAZY_PORT（預設 3530）／LAZY_BASE_REF（JSON 比對與重現用的固定基準提交，預設 f6659c4）
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行；跑完會清掉自己產生的檔。
 * 全部是假資料；不含任何帳密（SESSION_SECRET／密碼只在本次行程隨機產生）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const Module = require('module');
const { spawnSync, spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.LAZY_PORT || '3530', 10);
const BASE_REF = process.env.LAZY_BASE_REF || 'f6659c4';
const SECRET = 's' + crypto.randomBytes(16).toString('hex');

// ═════════════════════════ 子行程：單一情境 ═════════════════════════
const sleep = (n) => new Promise((r) => setTimeout(r, n));

async function childMain(scen) {
  const stats = { pools: 0, connects: 0, queries: 0, mainLoads: 0, mainSaves: 0, statsLoads: 0 };
  globalThis.__PGFAIL = false;
  const iso = '2026-01-01T00:00:00.000Z';
  const content = {
    contacts: [{ id: 'k1', name: '假人甲' }],                 // 沒有 owner → migrateOwner 會改
    opportunities: [{ id: 'o1', stage: '成交', owner: 'x' }], // → migrateStage成交ToWon 會改
    groups: [], monthlyBudgets: [],
    _auth: { users: [{ username: '_s_adm', displayName: '_s_adm', role: 'admin', active: true, passwordChangedAt: iso, mustChangePassword: false, bu: [] }] },
  };
  let updatedAt = new Date('2026-01-01T00:00:00Z');
  const fakeQuery = async (sql) => {
    stats.queries++;
    if (globalThis.__PGFAIL) throw new Error('fake pg: connection timeout');
    const q = String(sql);
    if (/SELECT content, updated_at FROM app_data WHERE id = 'main'/.test(q)) { stats.mainLoads++; return { rows: [{ content: JSON.parse(JSON.stringify(content)), updated_at: updatedAt }] }; }
    if (/SELECT updated_at FROM app_data WHERE id = 'main'/.test(q)) return { rows: [{ updated_at: updatedAt }] };
    if (/INSERT INTO app_data[\s\S]*'main'/.test(q)) { stats.mainSaves++; updatedAt = new Date(updatedAt.getTime() + 1000); return { rows: [{ updated_at: updatedAt }] }; }
    if (/api-stats/.test(q) && /SELECT/.test(q)) { stats.statsLoads++; return { rows: [] }; }
    return { rows: [] };
  };
  class FakePool {
    constructor() { stats.pools++; }
    on() { return this; }
    query(sql, p) { return fakeQuery(sql, p); }
    async connect() { stats.connects++; return { query: fakeQuery, release() {} }; }
    async end() {}
  }
  const origLoad = Module._load;
  Module._load = function (request, ...rest) { return request === 'pg' ? { Pool: FakePool } : origLoad.call(this, request, ...rest); };

  const logs = { overdue: 0, owner: 0, won: 0, sessionRefresh: 0, postReadyFail: 0 };
  const oLog = console.log, oErr = console.error;
  console.log = (...a) => { const s = a.join(' '); if (s.includes('[checkOverdueReminders] 執行完成')) logs.overdue++; if (s.includes('資料遷移完成')) logs.owner++; if (s.includes('[migrate] 成交→Won')) logs.won++; };
  console.error = (...a) => { const s = a.join(' '); if (s.includes('[session-refresh]')) logs.sessionRefresh++; if (s.includes('post-ready tasks failed')) logs.postReadyFail++; };

  const serverFile = path.resolve(ROOT, process.env.LAZY_SERVER || 'api/index.js');
  const app = require(serverFile);   // import 路徑（等同 Vercel 的 api/index.js）：此時 module-load 階段的行為由這裡觸發
  const jwtSession = require(path.join(ROOT, 'middleware', 'jwtSession.js'));
  const cookie = jwtSession.COOKIE_NAME + '=' + jwtSession._issueToken({ username: '_s_adm', displayName: '_s_adm', role: 'admin', bu: [] }, { lgn: Math.floor(Date.now() / 1000) });
  const srv = await new Promise((res) => { const s = app.listen(PORT, '127.0.0.1', () => res(s)); });
  const get = async (p, withCookie) => {
    const r = await fetch(`http://127.0.0.1:${PORT}${p}`, { redirect: 'manual', headers: withCookie ? { Cookie: cookie } : {} });
    await r.arrayBuffer();
    return r.status;
  };
  const snap = () => JSON.parse(JSON.stringify({ stats, logs }));
  const out = { scen, checks: {} };
  const ck = (name, cond, extra) => { out.checks[name] = !!cond; if (!cond) (out.fail = out.fail || []).push(name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : '')); };

  if (scen === 'static') {
    const paths = ['/login.html', '/login-boot.js', '/login-cinematic.css', '/itts-logo.png', '/sw.js', '/manifest.webmanifest', '/app.js', '/style.css', '/deep-link.js', '/q/abc', '/nonexistent.js'];
    const st = await Promise.all([...paths, ...paths, ...paths].map((p) => get(p, false)));
    const stAuthed = await Promise.all(['/app.js', '/style.css', '/quote.js'].map((p) => get(p, true)));
    await sleep(500);
    out.statuses = { login: st[0], authedStatic: stAuthed };
    ck('login.html 200', st[0] === 200, st[0]);
    ck('已登入靜態檔不 5xx', stAuthed.every((s) => s < 500), stAuthed);
    ck('靜態檔 0 個 Pool', stats.pools === 0, stats);
    ck('靜態檔 0 次連線', stats.connects === 0, stats);
    ck('靜態檔 0 次查詢', stats.queries === 0, stats);
    ck('靜態檔 0 次遷移', logs.overdue + logs.owner + logs.won === 0, logs);
  } else if (scen === 'api') {
    const burst = async () => Promise.all(Array.from({ length: 10 }, () => get('/api/me', true)));
    const s1 = await burst(); await sleep(700);
    ck('第一批 10 個 API 全 200', s1.every((s) => s === 200), s1);
    const a = snap();
    ck('overdue 只跑 1 次', a.logs.overdue === 1, a.logs);
    ck('migrateOwner 只跑 1 次', a.logs.owner === 1, a.logs);
    ck('migrateStage 只跑 1 次', a.logs.won === 1, a.logs);
    ck('apiMonitor 只載 1 次', a.stats.statsLoads === 1, a.stats);
    const s2 = await burst(); await sleep(400);
    const b = snap();
    ck('第二批仍 200', s2.every((s) => s === 200), s2);
    ck('第二批後仍各 1 次', b.logs.overdue === 1 && b.logs.owner === 1 && b.logs.won === 1 && b.stats.statsLoads === 1, b);
    out.snap = b;
  } else if (scen === 'nonapi') {
    const r = {};
    for (const p of ['/admin.html', '/uploads/none.png', '/', '/help.html', '/app.js']) r[p] = await get(p, true);
    await sleep(500);
    out.statuses = r;
    ck('/admin.html 冷實例 200（不 500）', r['/admin.html'] === 200, r);
    ck('/uploads 冷實例不 5xx', r['/uploads/none.png'] < 500, r);
    ck('/ 與 /help.html 不 5xx', r['/'] < 500 && r['/help.html'] < 500, r);
    ck('/app.js 已登入 200', r['/app.js'] === 200, r);
    ck('admin.html 觸發 ready 後遷移跑 1 次', logs.overdue === 1, logs);
  } else if (scen === 'fail') {
    globalThis.__PGFAIL = true;
    const f = await Promise.all([...Array.from({ length: 5 }, () => get('/api/me', true)), get('/admin.html', true)]);
    await sleep(400);
    ck('ready 失敗：API 全 503', f.slice(0, 5).every((s) => s === 503), f);
    ck('ready 失敗：admin.html 503', f[5] === 503, f);
    ck('ready 失敗：遷移 0 次', logs.overdue + logs.owner + logs.won === 0, logs);
    globalThis.__PGFAIL = false;
    const ok1 = await get('/api/me', true); await sleep(600);
    ck('恢復後 API 200', ok1 === 200, ok1);
    ck('恢復後遷移重試且各 1 次', logs.overdue === 1 && logs.owner === 1 && logs.won === 1, logs);
    await Promise.all(Array.from({ length: 5 }, () => get('/api/me', true))); await sleep(400);
    ck('之後不再重跑', logs.overdue === 1 && logs.owner === 1 && logs.won === 1, logs);
  } else throw new Error('unknown scen ' + scen);

  out.stats = stats; out.logs = logs;
  console.log = oLog; console.error = oErr;
  oLog('RESULT:' + JSON.stringify(out));
  srv.close(); process.exit(0);
}

// ═════════════════════════ 母行程 ═════════════════════════
const GUARD = ['auth.json', 'data.json', 'audit.log.json'].map((f) => path.join(ROOT, f));
const TEMPS = [];
function cleanup() {
  for (const f of TEMPS) { try { fs.rmSync(f, { recursive: true, force: true }); } catch { /* 忽略 */ } }
  for (const f of ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'mail-outbox.json', 'contact_audit.json']) { try { if (fs.existsSync(path.join(ROOT, f)) && ownFiles.has(f)) fs.rmSync(path.join(ROOT, f)); } catch { /* 忽略 */ } }
}
const ownFiles = new Set();
let pass = 0, fail = 0;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 400)); } }

function runChild(scen, serverFile) {
  const env = Object.assign({}, process.env, {
    DB_BACKEND: 'postgres', DATABASE_URL: 'postgres://fake:fake@127.0.0.1:1/fake', SESSION_SECRET: SECRET, JWT_SECRET: SECRET,
    LAZY_SERVER: serverFile || '', LAZY_PORT: String(PORT), NODE_ENV: 'test',
  });
  delete env.APP_ENV_LABEL;
  const r = spawnSync(process.execPath, [__filename, '--child', scen], { cwd: ROOT, env, encoding: 'utf8', timeout: 120000 });
  const line = String(r.stdout || '').split(/\r?\n/).filter((l) => l.startsWith('RESULT:')).pop();
  if (!line) return { crash: 'no RESULT; status=' + r.status + ' stderr=' + String(r.stderr).slice(-600) + ' stdout=' + String(r.stdout).slice(-300) };
  return JSON.parse(line.slice(7));
}
const SCENS = ['static', 'api', 'nonapi', 'fail'];
function scenGreen(res) { return !res.crash && !(res.fail && res.fail.length); }

function writeVariant(name, transform) {
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const out = transform(src);
  if (out === src) throw new Error('變異沒有改到任何東西：' + name);
  const f = path.join(ROOT, 'server.__lazy_' + name + '.js');
  fs.writeFileSync(f, out); TEMPS.push(f);
  return path.basename(f);
}

async function parentMain() {
  const ONLY = (process.env.LAZY_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
  const want = (k) => !ONLY.length || ONLY.includes(k);
  const exist = GUARD.filter((f) => fs.existsSync(f));
  if (exist.length) { console.log('拒絕：此工作樹已有', exist.map((f) => path.basename(f)).join('/'), '（可能是真實資料）'); process.exit(2); }
  const { isPortInUse } = require('./_port-free');
  if (typeof isPortInUse === 'function' && await isPortInUse(PORT)) { console.log('拒絕：埠', PORT, '已被佔用'); process.exit(2); }

  if (want('base')) {
    console.log('[基準重現] ' + BASE_REF + ' 的 server.js（修改前）：冷啟動只送靜態檔就該連資料庫');
    const g = spawnSync('git', ['show', BASE_REF + ':server.js'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
    if (g.status !== 0) throw new Error('git show 失敗');
    const f = path.join(ROOT, 'server.__lazy_base.js'); fs.writeFileSync(f, g.stdout); TEMPS.push(f);
    const r = runChild('static', path.basename(f));
    ok(!r.crash && r.fail && r.fail.length > 0, '修改前 static 情境必須紅（重現問題）', r.crash || JSON.stringify(r.stats));
    console.log('   修改前：', JSON.stringify(r.stats), JSON.stringify(r.logs));
  }
  if (want('scen')) {
    console.log('[修改後] 4 個情境（假 Pool）');
    for (const s of SCENS) {
      const r = runChild(s);
      ok(scenGreen(r), '情境 ' + s, r.crash || (r.fail || []).join(' | '));
      console.log('   ' + s + '：', r.crash ? 'CRASH' : JSON.stringify(r.checks), r.stats ? JSON.stringify(r.stats) : '');
    }
  }
  if (want('json')) await jsonCompare();
  if (want('mut')) {
    console.log('[變異測試] 每個變異都必須讓對應情境變紅');
    const muts = [
      ['eager_preload', ['static'], (s) => s.replace(/  \/\/ Postgres 模式：import 時不預熱/, "  if (process.env.DB_BACKEND === 'postgres') { db.ready().then(() => { migrateOwner(); migrateStage成交ToWon(); checkOverdueReminders(); }).catch(() => {}); apiMonitor.ready().catch(() => {}); }\n  // Postgres 模式：import 時不預熱")],
      ['flag_reset', ['api'], (s) => s.replace('function startInstanceTasksOnce() {', 'function startInstanceTasksOnce() { _instanceTasksStarted = false;')],
      ['nonapi_no_ready', ['nonapi'], (s) => s.replace(/\.\.\.\(process\.env\.DB_BACKEND === 'postgres' \? \[dbReadyMiddleware\] : \[\]\), /, '')],
      ['static_ready', ['static'], (s) => s.replace(/if \(!req\.path\.startsWith\('\/api\/'\)\) return next\(\);\s*return dbReadyMiddleware/, 'return dbReadyMiddleware')],
      ['flag_on_failure', ['fail'], (s) => s.replace('function dbReadyMiddleware(req, res, next) {', 'function dbReadyMiddleware(req, res, next) { _instanceTasksStarted = true;')],
    ];
    for (const [name, scens, tf] of muts) {
      const file = writeVariant(name, tf);
      let red = false; const detail = [];
      for (const s of scens) { const r = runChild(s, file); const bad = !scenGreen(r); red = red || bad; detail.push(s + (bad ? ' 紅' : ' 綠')); }
      ok(red, '變異 ' + name + ' 必須變紅', detail.join(','));
      console.log('   ' + name + '：' + detail.join(','));
    }
  }
  cleanup();
  console.log(`\n結果：${pass} 通過，${fail} 失敗` + (fail ? '：' + failures.join('；') : ''));
  process.exit(fail ? 1 : 0);
}

// ═════════════ JSON 模式：與固定基準提交逐位元相同 ═════════════
async function jsonCompare() {
  console.log('[JSON 模式] 與 ' + BASE_REF + ' 的主要端點輸出逐位元比對');
  const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
  const PW = 'Lz' + crypto.randomBytes(9).toString('hex') + '!Q';
  const now = new Date().toISOString();
  const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lazy-base-')); TEMPS.push(baseDir);
  const a = spawnSync('git', ['archive', BASE_REF], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive 失敗');
  const t = spawnSync('tar', ['-x', '-C', baseDir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗');
  fs.symlinkSync(path.join(ROOT, 'node_modules'), path.join(baseDir, 'node_modules'), 'junction');
  const seed = (dir) => {
    fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users: [{ username: '_s_adm', displayName: '_s_adm', role: 'admin', active: true, password: bcrypt.hashSync(PW, 6), createdAt: now, passwordChangedAt: now, mustChangePassword: false, canSetTargets: true, canDownloadContacts: true, bu: [] }] }));
    fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify({ contacts: [{ id: 'k1', name: '假人甲', owner: '_s_adm', company: '假公司', createdAt: now }], groups: [], opportunities: [{ id: 'o1', owner: '_s_adm', name: '假商機', stage: 'Won', amount: 1 }], visits: [], contracts: [], monthlyBudgets: [] }));
  };
  const norm = (s) => String(s).replace(/"loginAt":\d+/g, '"loginAt":N').replace(/\?v=\d+/g, '?v=N').replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>').replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<uuid>');
  const paths = ['/login.html', '/app.js', '/admin.html', '/', '/help.html', '/api/me', '/api/contacts', '/api/opportunities', '/api/groups', '/api/usermap', '/api/nonexistent', '/nonexistent', '/uploads/none.png', '/q/abc', '/sw.js', '/manifest.webmanifest'];
  const run = async (dir) => {
    seed(dir); ownFiles.add('auth.json'); ownFiles.add('data.json'); ownFiles.add('audit.log.json');
    const env = Object.assign({}, process.env, { SESSION_SECRET: SECRET, JWT_SECRET: SECRET, PORT: String(PORT), NODE_ENV: 'test' });
    delete env.DB_BACKEND; delete env.DATABASE_URL;
    const child = spawn(process.execPath, ['server.js'], { cwd: dir, env, stdio: 'ignore' });
    const outp = [];
    try {
      let up = false;
      for (let i = 0; i < 80 && !up; i++) { await sleep(250); try { up = (await fetch(`http://127.0.0.1:${PORT}/login.html`)).ok; } catch { /* 尚未起來 */ } }
      if (!up) throw new Error('伺服器沒起來：' + dir);
      const lr = await fetch(`http://127.0.0.1:${PORT}/api/login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: '_s_adm', password: PW }) });
      const ck = (lr.headers.getSetCookie ? lr.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
      outp.push('login ' + lr.status);
      for (const p of paths) {
        const r = await fetch(`http://127.0.0.1:${PORT}${p}`, { redirect: 'manual', headers: { Cookie: ck } });
        const body = Buffer.from(await r.arrayBuffer());
        outp.push([p, r.status, r.headers.get('content-type'), r.headers.get('location'), norm(body.toString('latin1'))].join(' | '));
      }
    } finally { try { child.kill(); } catch { /* 已結束 */ } await sleep(400); for (const f of ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'mail-outbox.json']) { try { fs.rmSync(path.join(dir, f), { force: true }); } catch { /* 忽略 */ } } }
    return outp;
  };
  const before = await run(baseDir);
  const after = await run(ROOT);
  let diff = 0;
  before.forEach((l, i) => { if (l !== after[i]) { diff++; console.log('   差異：', l.slice(0, 600), '\n      ≠', String(after[i]).slice(0, 600)); } });
  ok(before.length === after.length && before.length === paths.length + 1, 'JSON 比對筆數', before.length + '/' + after.length);
  ok(before[0] === 'login 200', 'JSON 基準登入成功', before[0]);
  ok(diff === 0, 'JSON 模式 ' + paths.length + ' 個端點輸出逐位元相同', diff + ' 處不同');
  console.log('   比對 ' + paths.length + ' 個端點，差異 ' + diff);
}

if (process.argv[2] === '--child') {
  childMain(process.argv[3]).catch((e) => { console.log('RESULT:' + JSON.stringify({ crash: String(e && e.stack || e) })); process.exit(1); });
} else {
  parentMain().catch((e) => { console.error(e); cleanup(); process.exit(1); });
}
