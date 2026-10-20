'use strict';
/**
 * 簽核流程設計器 UI 檢查用的沙盒：臨時 _s_* 帳號＋資料檔複本＋自行啟動的 server.js（JSON 模式、埠 FD_PORT 預設 3130）。
 * 不含任何帳密（密碼隨機只存在記憶體）；工作樹已有 auth.json／data.json／稽核檔時拒絕；結束以 PID 關閉並刪檔，絕不連遠端資料庫。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.FD_PORT || '3130', 10);
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
function guard() {
  if ([3000, 3020, 3030].includes(PORT)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const PW = 'Fu' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false }, extra || {});
function seed(extraData) {
  const users = [
    mkUser('_s_admin', 'admin'), mkUser('_s_own', 'user', { bu: ['ITS'], supervisor: '_s_m2' }), mkUser('_s_m2', 'manager2', { bu: ['ITS'], supervisor: '_s_m1' }),
    mkUser('_s_m1', 'manager1', { bu: ['ITS', 'ERP'] }), mkUser('_s_gm', 'executive'), mkUser('_s_ch', 'executive'), mkUser('_s_sec', 'secretary'),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [],
    quoteApproval: { roster: { gm: ['_s_gm'], chairman: ['_s_ch'], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: {} },
    approvalFlowConfig: { version: 1, buSettings: { ITS: { hasMgr2Step: true }, MDM: { hasMgr2Step: true } }, exceptions: [{ username: '_s_m2', action: 'skip_mgr1' }] },
  };
  Object.assign(d, extraData || {});
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
const cleanup = () => FILES.forEach((f) => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
const sleep = (n) => new Promise((r) => setTimeout(r, n));
let child = null, childPid = null;
async function startServer(env) {
  const e = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE']) delete e[k];
  Object.assign(e, env || {});
  child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env: e, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 100; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-600));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-600));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 30 && child.exitCode === null; i++) await sleep(100);
  child = null;
  await sleep(300);
}
module.exports = { ROOT, PORT, BASE, PW, guard, seed, cleanup, readData, startServer, stopServer, sleep, mkUser, FILES };
