'use strict';
/**
 * 稽核第三波第二批三支檢查腳本共用的沙盒測試骨架（不含任何帳密：測試帳號與隨機密碼只存在記憶體）。
 *
 * 做法：在「本工作樹」建立臨時 auth.json / data.json（假資料＋_s_* 臨時帳號），自行以 PORT=B2_PORT（預設 3211）
 * 啟動 server.js（JSON 模式、JWT），跑完只用 PID 關閉並刪除臨時檔。絕不連任何真實資料庫。
 * 守門：工作樹若已有 auth.json / data.json / audit.log.json*（可能是真實資料），一律拒絕執行。
 * 禁用埠：3000 / 3020 / 3030 / 3101–3201。
 */
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.B2_PORT || '3211', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3201)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));

function guardClean() {
  const g = FILES.filter(f => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}

const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const PW = 'B2' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();

const R = { pass: 0, fail: 0, failures: [] };
function ok(cond, name) { if (cond) R.pass++; else { R.fail++; R.failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));

const JWT_SECRET_FIXED = crypto.randomBytes(24).toString('hex');   // 整個測試行程固定：重啟伺服器後舊 cookie 仍有效（驗證持久化用）
const SESSION_SECRET_FIXED = crypto.randomBytes(24).toString('hex');
const HASH6 = bcrypt.hashSync(PW, 6);
const HASH12 = bcrypt.hashSync(PW, 12);
function mkUser(username, role, extra) {
  return Object.assign({ username, displayName: username, role, active: true, password: HASH6, createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
}
function writeSandbox(users, data) {
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(Object.assign({ contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [], groups: [] }, data || {}), null, 2));
}
function cleanupFiles() { FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }

// ───────── 伺服器 ─────────
let child = null, childPid = null;
const pids = [];
async function startServer(opts) {
  opts = opts || {};
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: JWT_SECRET_FIXED, SESSION_SECRET: SESSION_SECRET_FIXED, SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE }, opts.env || {});
  delete env.DATABASE_URL; delete env.MAIL_MODE;
  for (const k of ['LOGIN_LOCKOUT', 'LOGIN_MAX_FAILS', 'LOGIN_LOCK_MINUTES', 'LOGIN_LOCK_MAX_MINUTES', 'LOGIN_FAIL_WINDOW_MINUTES', 'SESSION_IDLE_MINUTES', 'FEATURE_GUARD_MODE']) if (!(opts.env && k in opts.env)) delete env[k];
  await assertPortFree(PORT);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
  child = spawn(process.execPath, [opts.file || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid; pids.push(childPid);
  let out = '';
  const onData = d => { out += d; if (process.env.B2_DEBUG) process.stdout.write('[srv] ' + d); };
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
  try { process.kill(childPid); } catch { /* 已結束 */ }   // 只殺自己記下的 PID
  for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100);
  child = null;
}
/** 重新播種並啟動（每個區段一台新伺服器：登入限流 15 分鐘 20 次是行程內記憶體，重啟歸零） */
async function fresh(users, data, opts) { await stopServer(); cleanupFiles(); writeSandbox(users, data); jar = {}; await startServer(opts); }
/** 不重新播種，只重啟（保留 data.json / auth.json，驗證持久化） */
async function restartKeep(opts) { await stopServer(); await startServer(opts); }   // 保留 jar：測試行程用固定密鑰，舊 cookie 仍有效

// ───────── HTTP ─────────
let jar = {};
async function http(who, method, url, body, opt) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  if (opt && opt.cookie) h.Cookie = opt.cookie;
  const t0 = Date.now();
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const ms = Date.now() - t0;
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fr = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fr && !/^itts_auth=;/.test(fr)) jar[who] = fr.split(';')[0];
  return { status: res.status, json, buf, text: buf.toString('utf8'), headers: res.headers, setCookie: sc, ms };
}
async function login(who, user, pw) {
  const r = await http(null, 'POST', '/api/login', { username: user || who, password: pw || PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
async function loginAll(names) { for (const n of names) { const r = await login(n); if (r.status !== 200) throw new Error('登入失敗 ' + n + ' ' + r.status); } }
async function getLogs(who) { const r = await http(who || '_s_admin', 'GET', '/api/admin/logs'); return Array.isArray(r.json) ? r.json : []; }
const byAction = (logs, a) => logs.filter(l => l.action === a);
function jwtPayload(cookie) { try { return JSON.parse(Buffer.from(String(cookie).split('=')[1].split('.')[1], 'base64url').toString('utf8')); } catch { return null; } }

function report(title) {
  console.log(`\n${title}結果：${R.pass} 通過，${R.fail} 失敗`);
  if (R.fail) console.log(R.failures.map(f => ' - ' + f).join('\n'));
  return R.fail ? 1 : 0;
}
async function finish() { await stopServer(); cleanupFiles(); }

module.exports = {
  ROOT, PORT, BASE, PW, FILES, R, ok, eq, sleep, mkUser, writeSandbox, cleanupFiles, guardClean, startServer, stopServer, fresh, restartKeep,
  http, login, loginAll, getLogs, byAction, jwtPayload, report, finish, bcrypt, HASH12, get jar() { return jar; }, set jar(v) { jar = v; }, pids,
};
