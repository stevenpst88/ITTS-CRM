#!/usr/bin/env node
'use strict';
/**
 * 簽核設定變更「兩人覆核」（內控稽核第三波第 2 項）端對端＋單元檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 * 做法：
 *  - 在「本工作樹」建立臨時 auth.json / data.json（假資料 + _s_* 臨時帳號），自行以 PORT=DR_PORT（預設 3191）啟動 server.js，
 *    依情境用環境變數開關 APPROVAL_CONFIG_DUAL_REVIEW 重啟；跑完只用 PID 關閉並刪除臨時檔。
 *  - 守門：工作樹若已有 auth.json / data.json / audit.log.json*（可能是真實資料）一律拒絕執行。禁用埠：3000/3020/3030/3101~3185。
 *  - 「關閉＝逐位元相同」：把基底 commit（DR_BASE_REF，預設釘死 397abc5）用 git archive 解到暫存資料夾（不是真實資料），
 *    以同一組請求打「舊程式」與「新程式（開關關）」，逐一比對回應、儲存結果與稽核紀錄。
 *  - 單元：lib/configReview.js 以假的 db / app 直接驅動（原子性、重新驗證失敗、逾期、擴充點），並做變異檢查。
 *  - 前端：_client/config-review.js 以 vm 載入（無 DOM），檢查晶片／橫幅／覆核卡片文字與跳脫（XSS）。
 *
 *   node scripts/check-config-dual-review.js            # 全部
 *   DR_ONLY=off,on node scripts/check-config-dual-review.js
 *   DR_SKIP_MUT=1 ...                                    # 略過變異檢查
 *   DR_SERVE=1 ...                                       # 截圖用：起一台已塞好待覆核資料的伺服器並留著（印出 PID 與測試密碼，Ctrl+C 或依 PID 關閉）
 */
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const path = require('path');
const vm = require('vm');
const Module = require('module');
const crypto = require('crypto');
const { spawn, execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.DR_PORT || '3191', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3185)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILE_NAMES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'];
const FILES = FILE_NAMES.map(f => path.join(ROOT, f));
const guardFiles = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
let failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));
const sect = (s) => console.log(s);

// ───────── 沙盒資料 ─────────
const PW = 'Dr' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
function seedUsers(opt) {
  return [
    mkUser('_s_admin', 'admin'), mkUser('_s_admin2', 'admin', opt && opt.admin2Off ? { active: false } : {}),
    mkUser('_s_user', 'user', { bu: ['ERP'] }),
    mkUser('_s_m1', 'manager1', { bu: ['ERP'] }), mkUser('_s_m2', 'manager2', { bu: ['ERP'] }),
    mkUser('_s_gm1', 'executive'), mkUser('_s_gm2', 'executive'),
    mkUser('_s_cp1', 'user', { bu: ['ERP'] }), mkUser('_s_cp2', 'user', { bu: ['ERP'] }),
  ];
}
function seedInto(dir, opt) {
  fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users: seedUsers(opt) }, null, 2));
  const d = { contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [], groups: [] };
  fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify(d, null, 2));
}
function cleanupIn(dir) { FILE_NAMES.forEach(f => { try { fs.unlinkSync(path.join(dir, f)); } catch { /* 不存在 */ } }); }
const readData = async (dir) => { await sleep(250); return JSON.parse(fs.readFileSync(path.join(dir || ROOT, 'data.json'), 'utf8')); };
const writeData = (d, dir) => fs.writeFileSync(path.join(dir || ROOT, 'data.json'), JSON.stringify(d, null, 2));

// ───────── 伺服器 ─────────
let child = null, childPid = null;
const pids = [];   // 記錄所有起過的 PID（只用 PID 關）
async function startServer(dir, env) {
  const e = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE }, env || {});
  delete e.DATABASE_URL; delete e.MAIL_MODE;
  if (!env || env.APPROVAL_CONFIG_DUAL_REVIEW === undefined) delete e.APPROVAL_CONFIG_DUAL_REVIEW;
  await assertPortFree(PORT);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
  child = spawn(process.execPath, [path.join(dir || ROOT, 'server.js')], { cwd: dir || ROOT, env: e, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid; pids.push(childPid);
  let out = '';
  const onData = d => { out += d; if (process.env.DR_DEBUG) process.stdout.write('[srv] ' + d); };
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
/** 清檔→種資料→以指定環境啟動（每個情境一台新伺服器：登入速率限制是行程內記憶體） */
async function fresh(env, opt, dir) { await stopServer(); cleanupIn(dir || ROOT); seedInto(dir || ROOT, opt); jar = {}; await startServer(dir, env); }
async function restartKeep(env, dir) { await stopServer(); jar = {}; await startServer(dir, env); }
const ON = { APPROVAL_CONFIG_DUAL_REVIEW: '1' };

// ───────── HTTP ─────────
let jar = {};
async function http(who, method, url, body) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fresh2 = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fresh2) jar[who] = fresh2.split(';')[0];
  return { status: res.status, json, text: buf.toString('utf8'), setCookie: sc };
}
async function login(who) {
  const r = await http(null, 'POST', '/api/login', { username: who, password: PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
async function loginAll(names) { for (const n of names) { const r = await login(n); if (r.status !== 200) throw new Error('登入失敗 ' + n + ' ' + r.status); } }
async function getLogs() { const r = await http('_s_admin', 'GET', '/api/admin/logs'); return Array.isArray(r.json) ? r.json : []; }
const byAction = (logs, a) => logs.filter(l => l.action === a);
const CFG = '/api/admin/quote-approval/config';
const FLOW = '/api/admin/quote-approval/flow-config';
const REQ = '/api/admin/approval-config-requests';
const cfgBody = (o) => Object.assign({ roster: { gm: [], chairman: [], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: {} }, o || {});

// ═════════════ 第 OFF 項：開關關閉＝行為與基底 逐位元相同 ═════════════
const normStore = (d) => ({
  quoteApproval: d.quoteApproval ? Object.assign({}, d.quoteApproval, { updatedAt: '<t>' }) : null,
  approvalFlowConfig: d.approvalFlowConfig ? Object.assign({}, d.approvalFlowConfig, { updatedAt: '<t>' }) : null,
  hasRequests: d.approvalChangeRequests !== undefined,
});
async function offScenario(dir, quick) {
  const out = [];
  const rec = async (name, r) => { out.push({ name, status: r.status, body: r.json }); return r; };
  await loginAll(['_s_admin']);
  const steps = quick ? [['c1']] : [['c1'], ['c2'], ['c3'], ['c4'], ['c5'], ['f1'], ['f2'], ['c6']];
  const bodies = {
    c1: cfgBody({ roster: { gm: ['_s_gm1'], chairman: [], boardProxy: [], costProviders: ['_s_cp1', '_s_cp1'], sealManagers: ['_s_admin'] }, productClasses: { 'ERP顧問': { cls: 'consult' }, '授權': { cls: 'software', costBySales: true } } }),
    c2: cfgBody({ roster: { gm: ['_s_admin'] } }),
    c3: cfgBody({ productClasses: { 'X': { cls: 'nope' } } }),
    c4: cfgBody({ roster: { gm: ['_s_nobody'] } }),
    c5: cfgBody({ roster: { gm: ['_s_gm2'], chairman: ['_s_gm1'], boardProxy: ['_s_cp2'], costProviders: ['_s_cp1'], sealManagers: [] }, productClasses: { 'ERP顧問': { cls: 'hardware' }, 'MDM': { cls: 'mdm' } } }),
    c6: {},
    f1: { buSettings: { ERP: { hasMgr2Step: true }, ITS: { hasMgr2Step: true } }, exceptions: [{ id: 'ex-1', username: '_s_m2', bu: 'ERP', role: 'manager2', action: 'skip_mgr1', label: '測試' }, { username: 'nobody', action: 'skip_mgr1' }, { id: 'ex-2', username: '_s_m1', action: 'other' }] },
    f2: { buSettings: { ERP: { hasMgr2Step: false } }, exceptions: [] },
  };
  for (const [k] of steps) {
    const r = await http('_s_admin', 'PUT', k.startsWith('f') ? FLOW : CFG, bodies[k]);
    await rec(k, r);
  }
  const g = await http('_s_admin', 'GET', FLOW); await rec('getFlow', g);
  const logs = (await getLogs()).filter(l => /^(UPDATE_QUOTE_APPROVAL_CONFIG|SAVE_APPROVAL_FLOW_CONFIG)$/.test(l.action)).map(l => ({ action: l.action, operator: l.operator, target: l.target, detail: l.detail }));
  return { out, store: normStore(await readData(dir)), logs };
}
async function secOff() {
  // 基底釘死為 commit（原本是 origin/main：兩人覆核合併進 main 之後變成自己比自己，檢查形同空殼）。
  // 397abc5＝兩人覆核功能（698c2e8）的父提交。可用 DR_BASE_REF 覆寫。
  const ref = process.env.DR_BASE_REF || '397abc5fbb4299976cb8889be3054f8dfa5c741f';
  sect('[OFF] 開關關閉：與修補前基底（' + ref.slice(0, 10) + '）逐位元相同');
  const tmp = path.join(process.env.DR_TMP || path.join(require('os').tmpdir(), 'dr_base_' + crypto.randomBytes(4).toString('hex')));
  fs.mkdirSync(tmp, { recursive: true });
  let baseRes = null;
  try {
    const tarFile = path.join(tmp, 'base.tar');
    // 排除根目錄的中文檔名啟動檔（.bat／.vbs）：部分 Windows tar／終端機編碼組合解不開非 ASCII 檔名，且伺服器用不到它們
    execFileSync('git', ['-C', ROOT, 'archive', '--format=tar', '-o', tarFile, ref, '--', '.', ':(exclude,glob)*.bat', ':(exclude,glob)*.vbs']);
    const dir = path.join(tmp, 'src'); fs.mkdirSync(dir);
    execFileSync('tar', ['-xf', 'base.tar', '-C', 'src'], { cwd: tmp });   // 用相對路徑：Windows 的 tar 會把 "C:" 當成遠端主機
    fs.symlinkSync(path.join(ROOT, 'node_modules'), path.join(dir, 'node_modules'), 'junction');
    await fresh(undefined, undefined, dir);
    baseRes = await offScenario(dir);
    await stopServer();
    cleanupIn(dir);
  } finally {
    await stopServer();
    // 先單獨移除 node_modules 連結（只刪連結本身）；確認連結已消失才整個刪暫存資料夾，絕不讓遞迴刪除跟進到真正的 node_modules
    const jl = path.join(tmp, 'src', 'node_modules');
    let linkGone = true;
    try { if (fs.existsSync(jl)) { fs.rmdirSync(jl); } } catch (e) { linkGone = false; console.log('  (略過暫存資料夾清理：連結無法移除)', e.message); }
    if (linkGone && !fs.existsSync(jl)) { try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* 已刪 */ } }
  }
  await fresh(undefined);
  const mine = await offScenario(ROOT);
  eq(mine.out.length, baseRes.out.length, 'OFF：步驟數相同');
  mine.out.forEach((m, i) => eq(m, baseRes.out[i], `OFF：步驟 ${m.name} 的 HTTP 狀態與回應與基底 逐位元相同`));
  eq(mine.store, baseRes.store, 'OFF：儲存結果（quoteApproval／approvalFlowConfig）與基底 相同，且沒有建立 approvalChangeRequests');
  eq(mine.logs, baseRes.logs, 'OFF：稽核紀錄（action／operator／target／detail）與基底 逐字相同');
  ok(mine.logs.length >= 5 && mine.out.filter(o => o.status === 400).length === 3, 'OFF：情境有效（至少 5 筆稽核、3 個 400 驗證失敗）');
  // 具體回應形狀
  eq(mine.out[0], { name: 'c1', status: 200, body: { success: true } }, 'OFF：PUT config 回 200 {success:true}（不是 202）');
  const v = await http('_s_admin', 'GET', '/api/admin/version');
  eq(v.json && v.json.approvalConfigDualReview, false, 'OFF：/api/admin/version 的 approvalConfigDualReview=false');
  const l = await http('_s_admin', 'GET', REQ);
  eq([l.status, l.json && l.json.enabled, l.json && l.json.requests.length], [200, false, 0], 'OFF：列表端點回 enabled=false、沒有申請');
  return mine;
}

// ═════════════ 第 ON 項：啟用 ═════════════
async function propose(who, kind, body) { return http(who, 'PUT', kind === 'flow' ? FLOW : CFG, body); }
const CLS_A = cfgBody({ roster: { gm: ['_s_gm1'], chairman: [], boardProxy: [], costProviders: ['_s_cp1'], sealManagers: [] }, productClasses: { 'ERP顧問': { cls: 'consult', costBySales: false }, '授權': { cls: 'software', costBySales: true } } });
const CLS_B = cfgBody({ roster: { gm: ['_s_gm1'], chairman: ['_s_gm2'], boardProxy: [], costProviders: ['_s_cp1'], sealManagers: [] }, productClasses: { 'ERP顧問': { cls: 'software', costBySales: false }, '授權': { cls: 'software', costBySales: true } } });

async function secOn() {
  sect('[ON-1] 提案 → 不生效；非管理員；基本欄位');
  await fresh(ON);
  await loginAll(['_s_admin', '_s_admin2', '_s_user']);
  let r = await http('_s_admin', 'GET', '/api/admin/version');
  eq(r.json && r.json.approvalConfigDualReview, true, 'ON：/api/admin/version 的 approvalConfigDualReview=true');
  const before = await readData();
  r = await propose('_s_admin', 'config', CLS_A);
  eq(r.status, 202, 'ON：PUT config → 202');
  eq(r.json && r.json.status, 'pending', '202 body.status=pending');
  eq(r.json && r.json.message, '已提出，等待另一位管理員覆核', '202 body.message');
  ok(r.json && /^[0-9a-f-]{36}$/.test(r.json.requestId), '202 帶 requestId（uuid）');
  const rid = r.json.requestId;
  let d = await readData();
  eq(d.quoteApproval, before.quoteApproval, 'ON：提案後儲存的 quoteApproval 完全沒變');
  eq((d.approvalChangeRequests || []).length, 1, 'ON：建立一筆申請');
  const rec = d.approvalChangeRequests[0];
  ok(rec.id === rid && rec.kind === 'config' && rec.proposedBy === '_s_admin' && rec.status === 'pending', '申請欄位：id／kind／proposedBy／status');
  ok(/^[0-9a-f]{64}$/.test(rec.baseHash), 'baseHash 是 sha256');
  eq(Date.parse(rec.expiresAt) - Date.parse(rec.proposedAt), 3 * 24 * 3600 * 1000, 'expiresAt = proposedAt + 3 天');
  eq(rec.history.length, 1, '申請 history 有一筆 propose');
  eq(rec.after, { roster: CLS_A.roster, productClasses: CLS_A.productClasses }, 'after 子樹＝提案內容（正規化後）');
  ok(rec.before && rec.before.roster && Array.isArray(rec.before.roster.gm), 'before 子樹存在');
  r = await propose('_s_user', 'config', CLS_A);
  eq(r.status, 403, '非管理員 PUT config → 403');
  r = await http('_s_user', 'GET', REQ);
  eq(r.status, 403, '非管理員 GET 列表 → 403');
  r = await http('_s_user', 'POST', REQ + '/' + rid + '/approve', {});
  eq(r.status, 403, '非管理員核可 → 403');
  r = await http(null, 'GET', REQ);
  eq(r.status, 401, '未登入 GET 列表 → 401');
  r = await propose('_s_admin', 'config', cfgBody({ roster: { gm: ['_s_admin2'] } }));
  eq([r.status, r.json && r.json.code], [400, 'ADMIN_IN_ROSTER'], '提案階段就跑同一套驗證（管理員不得入名單 → 400）');
  d = await readData();
  eq(d.approvalChangeRequests.length, 1, '驗證失敗不建立申請');
  // 沒變更
  r = await propose('_s_admin', 'config', cfgBody({ roster: before.quoteApproval ? before.quoteApproval.roster : undefined }));
  eq([r.status, r.json && r.json.noChange], [200, true], '與現況相同 → 200 noChange、不建立申請');
  d = await readData();
  eq(d.approvalChangeRequests.length, 1, 'noChange 不建立申請');
  // 列表
  r = await http('_s_admin', 'GET', REQ + '?status=pending');
  eq([r.status, r.json.enabled, r.json.pendingCount, r.json.requests.length, r.json.requests[0].isMine], [200, true, 1, 1, true], '列表：enabled／pendingCount／isMine（提案人看到 isMine=true）');
  r = await http('_s_admin2', 'GET', REQ + '/' + rid);
  eq([r.status, r.json.isMine, r.json.status], [200, false, 'pending'], '另一位管理員讀單筆：isMine=false');
  ok(r.json.before && r.json.after, '單筆含 before／after');
  r = await http('_s_admin', 'GET', REQ + '?status=approved');
  eq(r.json.requests.length, 0, '列表 status 篩選');

  sect('[ON-2] 差異摘要內容');
  const lines = rec.diffSummary;
  ok(lines.some(l => l === '總經理名單：新增 _s_gm1'), 'diffSummary：名冊 + 帳號');
  ok(lines.some(l => l === '成本填寫人名單：新增 _s_cp1'), 'diffSummary：成本填寫人名單');
  ok(lines.some(l => l.includes('新增「ERP顧問」→顧問服務') && l.includes('毛利≥25%一級主管、15%～25%總經理、<15%董事長')), 'diffSummary：新增商品歸類顯示結果門檻（顧問服務 25/15）');
  ok(lines.some(l => l.includes('新增「授權」→軟體規劃服務(業務填成本)') && l.includes('毛利≥10%一級主管、5%～10%總經理、<5%董事長')), 'diffSummary：軟體 10/5 且標示業務填成本');

  sect('[ON-3] 分權：提案人不能核可／退回；可撤回；退回需意見');
  r = await http('_s_admin', 'POST', REQ + '/' + rid + '/approve', {});
  eq([r.status, r.json && r.json.code], [403, 'SELF_REVIEW_FORBIDDEN'], '提案人核可自己的 → 403 SELF_REVIEW_FORBIDDEN');
  r = await http('_s_admin', 'POST', REQ + '/' + rid + '/reject', { comment: '我自己退回' });
  eq([r.status, r.json && r.json.code], [403, 'SELF_REVIEW_FORBIDDEN'], '提案人退回自己的 → 403 SELF_REVIEW_FORBIDDEN');
  r = await http('_s_admin2', 'POST', REQ + '/' + rid + '/cancel', {});
  eq([r.status, r.json && r.json.code], [403, 'NOT_PROPOSER'], '別人撤回 → 403 NOT_PROPOSER');
  d = await readData();
  eq(d.approvalChangeRequests[0].status, 'pending', '被拒絕的操作不改變狀態');
  eq(d.quoteApproval, before.quoteApproval, '被拒絕的操作不改變設定');
  r = await http('_s_admin2', 'POST', REQ + '/' + rid + '/reject', {});
  eq([r.status, r.json && r.json.code], [400, 'COMMENT_REQUIRED'], '退回沒意見 → 400 COMMENT_REQUIRED');
  r = await http('_s_admin2', 'POST', REQ + '/' + rid + '/reject', { comment: '啊' });
  eq(r.status, 400, '退回意見只有 1 字 → 400');
  r = await http('_s_admin2', 'POST', REQ + '/' + rid + '/reject', { comment: '名單要再確認' });
  eq([r.status, r.json && r.json.status], [200, 'rejected'], '另一位管理員退回（有意見）→ rejected');
  d = await readData();
  eq(d.quoteApproval, before.quoteApproval, '退回後設定不變');
  eq(d.approvalChangeRequests[0].history.map(h => h.action), ['propose', 'reject'], '歷史：propose → reject');
  eq(d.approvalChangeRequests[0].history[1].comment, '名單要再確認', '歷史記錄退回意見');
  r = await http('_s_admin2', 'POST', REQ + '/' + rid + '/approve', {});
  eq([r.status, r.json && r.json.code], [409, 'NOT_PENDING'], '已退回的不能再核可 → 409 NOT_PENDING');
  // 撤回
  r = await propose('_s_admin', 'config', CLS_A);
  const rid2 = r.json.requestId;
  r = await http('_s_admin', 'POST', REQ + '/' + rid2 + '/cancel', {});
  eq([r.status, r.json && r.json.status], [200, 'cancelled'], '提案人撤回自己的 pending → cancelled');
  r = await http('_s_admin2', 'POST', REQ + '/' + rid2 + '/approve', {});
  eq([r.status, r.json && r.json.code], [409, 'NOT_PENDING'], '已撤回的不能核可');

  sect('[ON-4] 核可 → 恰好套用提案內容；稽核；通知');
  r = await propose('_s_admin', 'config', CLS_A);
  const rid3 = r.json.requestId;
  r = await http('_s_admin2', 'POST', REQ + '/' + rid3 + '/approve', { comment: '同意' });
  eq([r.status, r.json && r.json.status], [200, 'approved'], '另一位管理員核可 → approved');
  d = await readData();
  eq({ roster: d.quoteApproval.roster, productClasses: d.quoteApproval.productClasses }, { roster: CLS_A.roster, productClasses: CLS_A.productClasses }, '核可後儲存的設定＝提案內容');
  eq([d.quoteApproval.updatedBy, d.quoteApproval.reviewedBy, d.quoteApproval.changeRequestId], ['_s_admin', '_s_admin2', rid3], 'updatedBy=提案人、reviewedBy=覆核人、changeRequestId');
  eq(d.approvalChangeRequests.find(x => x.id === rid3).history.map(h => h.action), ['propose', 'approve'], '歷史：propose → approve');
  // 一般設定讀取端也看到新設定
  r = await http('_s_admin', 'GET', '/api/quote-approval/config');
  eq(r.json && r.json.roster && r.json.roster.gm, ['_s_gm1'], 'GET 簽核設定讀到已生效的名冊');
  const logs = await getLogs();
  for (const a of ['CONFIG_CHANGE_PROPOSED', 'CONFIG_CHANGE_REJECTED', 'CONFIG_CHANGE_CANCELLED', 'CONFIG_CHANGE_APPROVED']) ok(byAction(logs, a).length >= 1, `稽核有 ${a}`);
  const up = byAction(logs, 'UPDATE_QUOTE_APPROVAL_CONFIG');
  eq(up.length, 1, 'UPDATE_QUOTE_APPROVAL_CONFIG 只有「真的套用」那一次');
  ok(up[0] && up[0].operator === '_s_admin2' && up[0].detail.includes('提案人 _s_admin') && up[0].detail.includes('覆核人 _s_admin2') && up[0].detail.includes('單號 ' + rid3), '套用的稽核含提案人／覆核人／單號');
  ok(up[0] && up[0].detail.includes('名單異動：總經理 +_s_gm1') && up[0].detail.includes('新增「ERP顧問」→顧問服務'), '套用的稽核仍有 before→after（名冊／商品歸類）');
  const all = JSON.stringify(logs);
  ok(!all.includes(PW) && !/\$2[aby]\$\d\d\$/.test(all), '稽核紀錄不含密碼／雜湊');
  const dr = JSON.stringify(d.approvalChangeRequests);
  ok(!dr.includes(PW) && !/\$2[aby]\$\d\d\$/.test(dr), '申請紀錄不含密碼／雜湊');
  // 通知
  r = await http('_s_admin2', 'GET', '/api/notifications');
  ok(Array.isArray(r.json) && r.json.some(n => n.type === 'config_review_requested' && n.refId === rid3), '另一位管理員收到 config_review_requested');
  r = await http('_s_admin', 'GET', '/api/notifications');
  ok(r.json.every(n => n.type !== 'config_review_requested'), '提案人本人不會收到「待覆核」通知');
  ok(r.json.some(n => n.type === 'config_review_approved' && n.refId === rid3), '提案人收到 config_review_approved');
  ok(r.json.some(n => n.type === 'config_review_rejected' && n.refId === rid), '提案人收到 config_review_rejected');
  ok(!JSON.stringify(r.json).includes('名單要再確認') && !JSON.stringify(r.json).includes(PW), '通知內文不含退回意見原文／密碼');

  sect('[ON-5] 流程設定 flow');
  const FL = { buSettings: { ERP: { hasMgr2Step: true } }, exceptions: [{ id: 'ex-1', username: '_s_m2', bu: 'ERP', role: 'manager2', action: 'skip_mgr1', label: '測試' }] };
  r = await propose('_s_admin2', 'flow', FL);
  eq(r.status, 202, 'ON：PUT flow-config → 202');
  const frid = r.json.requestId;
  d = await readData();
  eq(d.approvalFlowConfig, undefined, 'flow 提案後現況不變（沒有 approvalFlowConfig）');
  const frec = d.approvalChangeRequests.find(x => x.id === frid);
  eq(frec.kind, 'flow', 'kind=flow');
  ok(frec.diffSummary.includes('ERP 二級主管關 停用→啟用') && frec.diffSummary.includes('新增例外 skip_mgr1：_s_m2(ERP)'), 'flow 差異摘要：BU 二級主管關＋例外');
  r = await http('_s_admin2', 'POST', REQ + '/' + frid + '/approve', {});
  eq(r.status, 403, 'flow：提案人不能核可');
  r = await http('_s_admin', 'POST', REQ + '/' + frid + '/approve', {});
  eq(r.status, 200, 'flow：另一位管理員核可');
  d = await readData();
  eq([d.approvalFlowConfig.buSettings.ERP, d.approvalFlowConfig.exceptions], [{ hasMgr2Step: true }, FL.exceptions], 'flow：核可後生效且等於提案');
  const fl = byAction(await getLogs(), 'SAVE_APPROVAL_FLOW_CONFIG');
  ok(fl.length === 1 && fl[0].operator === '_s_admin' && fl[0].detail.includes('單號 ' + frid) && fl[0].detail.includes('ERP 二級主管關 停用→啟用'), 'flow：SAVE_APPROVAL_FLOW_CONFIG 稽核含 before→after 與單號');
  r = await http('_s_admin', 'GET', FLOW);
  eq(r.json.buSettings.ERP.hasMgr2Step, true, 'GET flow-config 讀到已生效設定');
  r = await propose('_s_admin', 'flow', { buSettings: { ERP: { hasMgr2Step: true } }, exceptions: FL.exceptions });
  eq([r.status, r.json && r.json.noChange], [200, true], 'flow：與現況相同 → noChange');
}

async function secStale() {
  sect('[STALE] 基準變動');
  await fresh(ON);
  await loginAll(['_s_admin', '_s_admin2']);
  let r = await propose('_s_admin', 'config', CLS_A); const p1 = r.json.requestId;
  r = await propose('_s_admin2', 'config', CLS_B); const p2 = r.json.requestId;
  r = await http('_s_admin2', 'POST', REQ + '/' + p1 + '/approve', {});
  eq(r.status, 200, 'P1 核可（基準仍相同）');
  r = await http('_s_admin', 'POST', REQ + '/' + p2 + '/approve', {});
  eq([r.status, r.json && r.json.code, r.json && r.json.status], [409, 'STALE_BASE', 'stale'], 'P2 在 P1 生效後核可 → 409 STALE_BASE');
  ok(r.json && /重新提出/.test(r.json.error), 'STALE_BASE 訊息提示提案人重新提出');
  const d = await readData();
  eq(d.approvalChangeRequests.find(x => x.id === p2).status, 'stale', 'P2 狀態標成 stale');
  eq({ roster: d.quoteApproval.roster, productClasses: d.quoteApproval.productClasses }, { roster: CLS_A.roster, productClasses: CLS_A.productClasses }, '設定仍是 P1，P2 沒有套用');
  ok(byAction(await getLogs(), 'CONFIG_CHANGE_STALE').length === 1, '稽核有 CONFIG_CHANGE_STALE');
  r = await http('_s_admin2', 'GET', '/api/notifications');
  ok(r.json.some(n => n.type === 'config_review_stale' && n.refId === p2), '提案人 P2 收到 stale 通知');
  r = await http('_s_admin', 'POST', REQ + '/' + p2 + '/approve', {});
  eq([r.status, r.json && r.json.code], [409, 'NOT_PENDING'], 'stale 後不能再核可');
}

async function secExpiry() {
  sect('[EXPIRY] 3 天逾期（惰性）');
  await fresh(ON);
  await loginAll(['_s_admin', '_s_admin2']);
  let r = await propose('_s_admin', 'config', CLS_A); const e1 = r.json.requestId;
  r = await propose('_s_admin', 'flow', { buSettings: { ITS: { hasMgr2Step: true } }, exceptions: [] }); const e2 = r.json.requestId;
  r = await propose('_s_admin', 'config', CLS_B); const e3 = r.json.requestId;
  await stopServer();
  const d0 = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
  const back = (id, ms) => { const x = d0.approvalChangeRequests.find(q => q.id === id); x.proposedAt = new Date(Date.now() - ms).toISOString(); x.expiresAt = new Date(Date.parse(x.proposedAt) + 3 * 24 * 3600 * 1000).toISOString(); };
  back(e1, 3 * 24 * 3600 * 1000 + 60 * 1000);   // 逾期 1 分鐘
  back(e2, 3 * 24 * 3600 * 1000 - 3600 * 1000); // 還剩 1 小時
  back(e3, 4 * 24 * 3600 * 1000);               // 逾期
  writeData(d0);
  await restartKeep(ON);
  await loginAll(['_s_admin', '_s_admin2']);
  // 先直接核可 e1（沒先列表）：必須在核可當下結算為 expired
  r = await http('_s_admin2', 'POST', REQ + '/' + e1 + '/approve', {});
  eq([r.status, r.json && r.json.code], [409, 'EXPIRED'], '逾期的不能核可 → 409 EXPIRED');
  let d = await readData();
  eq(d.approvalChangeRequests.find(x => x.id === e1).status, 'expired', '核可當下把申請結算成 expired（已存檔）');
  eq(d.quoteApproval, undefined, '逾期核可沒有套用');
  // 列表結算 e3，未逾期的 e2 維持 pending
  r = await http('_s_admin2', 'GET', REQ);
  const st = Object.fromEntries(r.json.requests.map(x => [x.id, x.status]));
  eq([st[e1], st[e2], st[e3]], ['expired', 'pending', 'expired'], '列表惰性結算：逾期者 expired、未逾期者 pending');
  eq(r.json.pendingCount, 1, 'pendingCount 只算 pending');
  ok(byAction(await getLogs(), 'CONFIG_CHANGE_EXPIRED').length === 2, '稽核有 2 筆 CONFIG_CHANGE_EXPIRED（不重複）');
  await http('_s_admin2', 'GET', REQ);
  ok(byAction(await getLogs(), 'CONFIG_CHANGE_EXPIRED').length === 2, '再讀一次不會重複記 EXPIRED');
  r = await http('_s_admin', 'GET', '/api/notifications');
  ok(r.json.filter(n => n.type === 'config_review_expired').length === 2, '提案人收到逾期通知');
  r = await http('_s_admin', 'POST', REQ + '/' + e3 + '/cancel', {});
  eq([r.status, r.json && r.json.code], [409, 'EXPIRED'], '逾期的不能撤回 → 409 EXPIRED');
  r = await http('_s_admin2', 'POST', REQ + '/' + e2 + '/approve', {});
  eq(r.status, 200, '還沒逾期的仍可核可');
}

async function secOneAdmin() {
  sect('[FAILSAFE] 在職管理員不足 2 人');
  await fresh(ON, { admin2Off: true });
  await loginAll(['_s_admin']);
  let r = await propose('_s_admin', 'config', CLS_A);
  eq([r.status, r.json && r.json.code], [409, 'DUAL_REVIEW_NEEDS_TWO_ADMINS'], 'ON＋只有 1 位在職管理員 → 409 DUAL_REVIEW_NEEDS_TWO_ADMINS');
  ok(r.json && /不足 2 人/.test(r.json.error), '訊息說明原因');
  r = await propose('_s_admin', 'flow', { buSettings: {}, exceptions: [] });
  eq([r.status, r.json && r.json.code], [409, 'DUAL_REVIEW_NEEDS_TWO_ADMINS'], 'flow 同樣擋下');
  const d = await readData();
  eq([d.approvalChangeRequests, d.quoteApproval], [undefined, undefined], '沒有建立申請、也沒有直接生效（無單人繞過）');
  r = await http('_s_admin', 'GET', REQ);
  eq([r.json.enabled, r.json.activeAdmins], [true, 1], '列表回報 activeAdmins=1（畫面據此顯示紅色晶片）');
  // 開關關閉時，同樣只有 1 位管理員仍可直接儲存
  await fresh(undefined, { admin2Off: true });
  await loginAll(['_s_admin']);
  r = await propose('_s_admin', 'config', CLS_A);
  eq(r.status, 200, 'OFF＋1 位管理員 → 照舊直接儲存');
}

async function secStatic() {
  sect('[STATIC] 前端檔案與接線');
  await fresh(undefined);
  await loginAll(['_s_admin']);
  const r = await http('_s_admin', 'GET', '/config-review.js');
  ok(r.status === 200 && /CfgReview/.test(r.text), '/config-review.js 可取得');
  const i = await http('_s_admin', 'GET', '/index.html');
  ok(/config-review\.js\?v=/.test(i.text) && i.text.indexOf('config-review.js') < i.text.indexOf('quote-approval.js'), 'index.html 載入 config-review.js（有版本戳、在 quote-approval.js 之前）');
  const a = await http('_s_admin', 'GET', '/admin.html');
  ok(/config-review\.js/.test(a.text) && a.text.indexOf('id="afCrHost"') > 0, 'admin.html 載入 config-review.js 且有 afCrHost');
}

// ═════════════ 單元：lib/configReview.js（假 db／假 app）═════════════
function loadCR(mutate) {
  const file = path.join(ROOT, 'lib', 'configReview.js');
  let src = fs.readFileSync(file, 'utf8');
  if (mutate) { const [a, b] = mutate; if (!src.includes(a)) throw new Error('變異錨點不存在：' + a.slice(0, 50)); src = src.replace(a, b); }
  const m = new Module(file, module); m.filename = file; m.paths = Module._nodeModulePaths(path.dirname(file));
  m._compile(src, file);
  return m.exports;
}
function fakeApp() {
  const routes = {};
  const mk = (method) => (p, guards, handler) => { routes[method + ' ' + p] = handler; };
  return { routes, get: mk('GET'), post: mk('POST') };
}
function fakeRes() { const r = { code: 200, body: null, status(c) { r.code = c; return r; }, json(b) { r.body = b; return r; } }; return r; }
async function unitRules(CRmod, opt) {
  opt = opt || {};
  const res = {};
  const mkEnv = () => {
    const data = { quoteApproval: { roster: { gm: [] }, productClasses: {} }, approvalFlowConfig: undefined };
    const logs = [], notes = [];
    let clock = Date.parse('2026-10-14T00:00:00Z');
    const db = { data, saved: 0, failSave: false, load() { return data; }, save() { if (db.failSave) throw new Error('boom'); db.saved++; }, flush() { return Promise.resolve(); } };
    const auth = { users: [{ username: 'a1', role: 'admin', active: true }, { username: 'a2', role: 'admin', active: true }, { username: 'u1', role: 'user', active: true }] };
    const cr = CRmod.create({ db, writeLog: (a, o, t, d) => logs.push({ a, o, t, d }), pushNotification: (to, type, title, body, ref) => notes.push({ to, type, title, body, ref }), loadAuth: () => auth, now: () => clock, env: { APPROVAL_CONFIG_DUAL_REVIEW: '1' } });
    let applyThrows = false, validateFails = false;
    cr.registerKind('config', {
      label: '測試設定', rawKeys: ['quoteApproval'],
      read: (d) => JSON.parse(JSON.stringify(d.quoteApproval)),
      validate: (after) => (validateFails ? { ok: false, status: 400, code: 'BAD_CONFIG', error: '壞了' } : { ok: true, value: after }),
      diff: (b, a) => ['改了 ' + JSON.stringify(a)],
      apply: (d, v, m) => { d.quoteApproval = Object.assign({}, v, { by: m.by }); if (applyThrows) throw new Error('apply boom'); return { action: 'TEST_APPLY', target: 't', detail: 'd' }; },
    });
    const app = fakeApp(); cr.registerRoutes(app, [], (req) => ({ me: req.me, role: req.role || 'admin', users: {} }));
    const call = async (route, me, id, body, role) => { const rs = fakeRes(); await app.routes[route]({ me, role, params: { id }, body: body || {}, query: {} }, rs); return rs; };
    const propose = (me, value) => { const rs = fakeRes(); cr.propose(rs, 'config', { data, value, proposer: me, users: {}, note: '', req: {} }); return rs; };
    return { data, logs, notes, db, cr, call, propose, tick: (ms) => { clock += ms; }, set: (k, v) => { if (k === 'apply') applyThrows = v; else validateFails = v; } };
  };
  // 自我覆核
  { const e = mkEnv(); const p = e.propose('a1', { roster: { gm: ['x'] }, productClasses: {} }); const id = p.body.requestId;
    const a = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a1', id);
    const rj = await e.call('POST /api/admin/approval-config-requests/:id/reject', 'a1', id, { comment: '自己退回' });
    res.selfApproveBlocked = a.code === 403 && a.body.code === 'SELF_REVIEW_FORBIDDEN';
    res.selfRejectBlocked = rj.code === 403 && rj.body.code === 'SELF_REVIEW_FORBIDDEN';
    res.selfStillPending = e.data.approvalChangeRequests[0].status === 'pending' && JSON.stringify(e.data.quoteApproval.roster.gm) === '[]'; }
  // 退回需意見
  { const e = mkEnv(); const id = e.propose('a1', { roster: { gm: ['x'] } }).body.requestId;
    const r1 = await e.call('POST /api/admin/approval-config-requests/:id/reject', 'a2', id, { comment: ' ' });
    const r2 = await e.call('POST /api/admin/approval-config-requests/:id/reject', 'a2', id, { comment: '一' });
    res.rejectNeedsComment = r1.code === 400 && r2.code === 400 && e.data.approvalChangeRequests[0].status === 'pending'; }
  // 基準變動
  { const e = mkEnv(); const id = e.propose('a1', { roster: { gm: ['x'] } }).body.requestId;
    e.data.quoteApproval = { roster: { gm: ['z'] }, productClasses: {} };
    const a = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a2', id);
    res.staleBlocked = a.code === 409 && a.body.code === 'STALE_BASE' && e.data.approvalChangeRequests[0].status === 'stale' && JSON.stringify(e.data.quoteApproval.roster.gm) === '["z"]'; }
  // 逾期（邊界：剛好等於 expiresAt 也算逾期）
  { const e = mkEnv(); const id = e.propose('a1', { roster: { gm: ['x'] } }).body.requestId;
    e.tick(3 * 24 * 3600 * 1000 - 1);
    const okA = e.cr.sweep(e.data).length === 0;
    e.tick(1);
    const a = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a2', id);
    res.expiryBlocked = okA && a.code === 409 && a.body.code === 'EXPIRED' && e.data.approvalChangeRequests[0].status === 'expired' && JSON.stringify(e.data.quoteApproval.roster.gm) === '[]'; }
  // 原子性：apply 丟例外 → 還原、維持 pending、設定不變
  { const e = mkEnv(); const id = e.propose('a1', { roster: { gm: ['x'] } }).body.requestId;
    const orig = e.data.quoteApproval; e.set('apply', true);
    const a = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a2', id);
    const rec = e.data.approvalChangeRequests[0];
    res.applyFailRollback = a.code === 500 && a.body.code === 'APPLY_FAILED' && rec.status === 'pending' && rec.history.length === 1 && e.data.quoteApproval === orig;
    e.set('apply', false);
    const b = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a2', id);
    res.retryAfterFailWorks = b.code === 200 && e.data.quoteApproval.roster.gm[0] === 'x' && rec.status === 'approved'; }
  // 原子性：db.save 丟例外 → 還原
  { const e = mkEnv(); const id = e.propose('a1', { roster: { gm: ['x'] } }).body.requestId;
    const orig = e.data.quoteApproval; e.db.failSave = true;
    const a = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a2', id);
    const rec = e.data.approvalChangeRequests[0];
    res.saveFailRollback = a.code === 500 && rec.status === 'pending' && rec.history.length === 1 && e.data.quoteApproval === orig && !e.logs.some(l => l.a === 'TEST_APPLY' || l.a === 'CONFIG_CHANGE_APPROVED'); }
  // 核可當下重新驗證失敗 → 400、維持 pending、不套用
  { const e = mkEnv(); const id = e.propose('a1', { roster: { gm: ['x'] } }).body.requestId; e.set('validate', true);
    const a = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a2', id);
    res.revalidateBlocks = a.code === 400 && a.body.code === 'BAD_CONFIG' && e.data.approvalChangeRequests[0].status === 'pending' && JSON.stringify(e.data.quoteApproval.roster.gm) === '[]'; }
  // 擴充點：新登錄的 kind 也走同一套（例如日後的 quoteMode）
  { const e = mkEnv(); e.data.quoteMode = { mode: 'old' };
    e.cr.registerKind('quoteMode', { label: '報價模式', rawKeys: ['quoteMode'], read: (d) => d.quoteMode, validate: (a) => ({ ok: true, value: a }), diff: (b, a) => ['模式 ' + b.mode + '→' + a.mode], apply: (d, v) => { d.quoteMode = v; return { action: 'SAVE_MODE', target: 'q', detail: 'x' }; } });
    const rs = fakeRes(); e.cr.propose(rs, 'quoteMode', { data: e.data, value: { mode: 'new' }, proposer: 'a1', users: {}, req: {} });
    const a = await e.call('POST /api/admin/approval-config-requests/:id/approve', 'a2', rs.body.requestId);
    res.extensionKind = rs.code === 202 && a.code === 200 && e.data.quoteMode.mode === 'new'; }
  // 雜湊穩定（鍵順序無關）
  res.hashStable = CRmod.hashOf({ a: 1, b: { c: [1, 2], d: 2 } }) === CRmod.hashOf({ b: { d: 2, c: [1, 2] }, a: 1 }) && CRmod.hashOf({ a: 1 }) !== CRmod.hashOf({ a: 2 });
  // 開關解析
  res.switchParse = CRmod.isEnabled({ APPROVAL_CONFIG_DUAL_REVIEW: '1' }) && CRmod.isEnabled({ APPROVAL_CONFIG_DUAL_REVIEW: 'true' }) && !CRmod.isEnabled({}) && !CRmod.isEnabled({ APPROVAL_CONFIG_DUAL_REVIEW: '0' }) && !CRmod.isEnabled({ APPROVAL_CONFIG_DUAL_REVIEW: '' }) && !CRmod.isEnabled({ APPROVAL_CONFIG_DUAL_REVIEW: 'no' });
  // 只有 1 位在職管理員
  { const e = mkEnv(); e.cr.activeAdmins; const rs = fakeRes();
    const auth1 = CRmod.create({ db: e.db, writeLog() {}, pushNotification() {}, loadAuth: () => ({ users: [{ username: 'a1', role: 'admin', active: true }, { username: 'a2', role: 'admin', active: false }] }), env: {} });
    auth1.registerKind('config', { label: 'x', rawKeys: [], read: () => ({}), validate: (a) => ({ ok: true, value: a }), diff: () => [], apply: () => ({}) });
    auth1.propose(rs, 'config', { data: e.data, value: { z: 1 }, proposer: 'a1', users: {}, req: {} });
    res.needsTwoAdmins = rs.code === 409 && rs.body.code === 'DUAL_REVIEW_NEEDS_TWO_ADMINS' && !e.data.approvalChangeRequests; }
  return res;
}
const UNIT_NAMES = {
  selfApproveBlocked: '提案人核可自己的被擋', selfRejectBlocked: '提案人退回自己的被擋', selfStillPending: '被擋後申請仍 pending 且設定不變', rejectNeedsComment: '退回必須有 ≥2 字意見',
  staleBlocked: '基準變動 → STALE_BASE 並標 stale', expiryBlocked: '逾期（含邊界）不能核可且標 expired', applyFailRollback: '套用失敗 → 還原，申請維持 pending、設定不變', retryAfterFailWorks: '失敗後重試可成功',
  saveFailRollback: 'db.save 失敗 → 還原、不記套用稽核', revalidateBlocks: '核可當下重新驗證失敗 → 不套用', extensionKind: '擴充點：新登錄的 kind（quoteMode 之類）走同一套流程', hashStable: '雜湊與鍵順序無關', switchParse: '開關解析（僅 1／true 啟用）', needsTwoAdmins: '在職管理員不足 2 人 → 409 且不建立申請',
};
async function secUnit() {
  sect('[UNIT] lib/configReview.js 單元');
  const res = await unitRules(loadCR());
  Object.keys(UNIT_NAMES).forEach(k => ok(res[k] === true, `單元：${UNIT_NAMES[k]}`));
}

// ═════════════ 前端（vm）═════════════
function secUi() {
  sect('[UI] _client/config-review.js（vm）與接線靜態檢查');
  const src = fs.readFileSync(path.join(ROOT, '_client', 'config-review.js'), 'utf8');
  const win = {}; const ctx = { window: win, console }; vm.createContext(ctx); vm.runInContext(src, ctx);
  const C = win.CfgReview;
  ok(C && typeof C.chipHtml === 'function', 'CfgReview 載入');
  ok(/兩人覆核：未啟用/.test(C.chipHtml({ enabled: false })), '晶片：未啟用');
  ok(/兩人覆核：啟用</.test(C.chipHtml({ enabled: true, activeAdmins: 2 })) && /data-cr-chip="on"/.test(C.chipHtml({ enabled: true, activeAdmins: 2 })), '晶片：啟用（綠）');
  const red = C.chipHtml({ enabled: true, activeAdmins: 1 });
  ok(/data-cr-chip="blocked"/.test(red) && /#fee2e2/.test(red) && /不足 2 人/.test(red), '晶片：啟用但管理員不足 2 人 → 紅色');
  eq(C.saveLabel(true, '儲存設定'), '提出變更（待覆核）', '啟用時儲存鈕＝「提出變更（待覆核）」');
  eq(C.saveLabel(false, '儲存設定'), '儲存設定', '未啟用時儲存鈕文字不變');
  ok(C.proposalMessage({ requestId: 'abcdef12-0000', message: '已提出，等待另一位管理員覆核' }).includes('abcdef12'), '202 訊息含單號前 8 碼');
  eq(C.bannerHtml({ enabled: false, pendingCount: 3 }), '', '未啟用：沒有橫幅');
  ok(/待覆核變更 3 筆/.test(C.bannerHtml({ enabled: true, pendingCount: 3 })) && /data-cr="open"/.test(C.bannerHtml({ enabled: true, pendingCount: 3 })), '啟用：橫幅「待覆核變更 N 筆」＋覆核鈕');
  const XSS = '<img src=x onerror=alert(1)>';
  const mine = { id: 'id"><script>1</script>', kind: 'config', kindLabel: XSS, status: 'pending', proposedBy: XSS, proposedAt: nowIso, expiresAt: nowIso, diffSummary: [XSS, '「' + XSS + '」類別 A→B'], history: [{ at: nowIso, by: XSS, action: 'propose', comment: XSS }], isMine: true, baseChanged: false };
  const theirs = Object.assign({}, mine, { isMine: false, baseChanged: true });
  const done = Object.assign({}, mine, { status: 'rejected', decidedBy: XSS, decidedAt: nowIso, history: mine.history.concat([{ at: nowIso, by: XSS, action: 'reject', comment: XSS }]) });
  const html = C.modalInnerHtml([mine, theirs, done]);
  ok(!/<img|<script|<svg|<iframe/i.test(html), 'XSS：提案人／摘要／意見／單號中的標籤全被跳脫');
  ok(!/data-id="[^"]*"[^>]*onerror/i.test(html) && !/="id">/.test(html) && (html.match(/data-id="/g) || []).length > 0 && !/data-id="id"/.test(html), 'XSS：單號中的引號不會跳出屬性');
  ok(/&lt;img src=x onerror=alert\(1\)&gt;/.test(html), 'XSS：內容以文字顯示');
  const cardMine = C.requestCardHtml(mine);
  ok(/這是你自己提出的變更/.test(cardMine) && /data-cr="cancel"/.test(cardMine) && !/data-cr="approve"/.test(cardMine) && !/data-cr="reject"/.test(cardMine) && (cardMine.match(/disabled/g) || []).length === 2, '提案人視角：核可／退回停用並說明、只有撤回可按');
  const cardOther = C.requestCardHtml(theirs);
  ok(/data-cr="approve"/.test(cardOther) && /data-cr="reject"/.test(cardOther) && !/data-cr="cancel"/.test(cardOther) && /至少 2 字/.test(cardOther), '覆核人視角：核可／退回、意見欄（退回必填）、沒有撤回');
  ok(/提案之後設定已被其他變更改動/.test(cardOther), 'baseChanged 顯示提醒');
  const hist = C.historyHtml([mine, done]);
  ok(/已退回/.test(hist) && !/待覆核/.test(hist.replace(/待覆核變更/g, '')), '歷史只列已處理的（含狀態文字）');
  // create(): paint / refresh（假 host）
  const host = { innerHTML: '', _l: [], addEventListener(t, f) { this._l.push(f); } };
  let calls = [];
  const cr = C.create({ api: async (m, p) => { calls.push(m + ' ' + p); return { ok: true, status: 200, data: { enabled: true, activeAdmins: 2, pendingCount: 2, me: 'a', requests: [theirs, mine] } }; } });
  cr.paint(host);
  ok(/未啟用/.test(host.innerHTML), 'paint（尚未載入）預設顯示未啟用');
  return cr.refresh().then(() => {
    ok(/啟用/.test(host.innerHTML) && /待覆核變更 2 筆/.test(host.innerHTML) && calls[0] === 'GET /admin/approval-config-requests', 'refresh 後重畫晶片與橫幅，並呼叫列表端點');
    eq(cr.saveLabel('儲存'), '提出變更（待覆核）', 'create().saveLabel 隨伺服器狀態');
    // 靜態接線
    const qa = fs.readFileSync(path.join(ROOT, '_client', 'quote-approval.js'), 'utf8');
    const ah = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
    ok(/CfgReview\.create/.test(qa) && /qapCrHost/.test(qa) && /saveLabel\('儲存名冊與歸類'\)/.test(qa) && /res\.status === 202/.test(qa), 'quote-approval.js 接線：晶片／儲存鈕文字／202 訊息');
    ok(/_afCrApi/.test(ah) && /afCrHost/.test(ah) && /r\.status === 202/.test(ah) && /async function _afCrApi[\s\S]{0,600}fetch\(API \+ path/.test(ah), 'admin.html 接線：不走 adminFetch（403 不會被當 session 過期）、202 訊息');
    // 內嵌 script 語法
    const scripts = [...ah.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]).filter(s => s.trim());
    let bad = 0; scripts.forEach(s => { try { new Function(s); } catch (e) { bad++; console.log('   admin.html inline script 語法錯誤：', e.message); } });
    ok(scripts.length > 0 && bad === 0, `admin.html 內嵌 script ${scripts.length} 段語法正確`);
    ['_client/config-review.js', '_client/quote-approval.js'].forEach(f => { try { new Function(fs.readFileSync(path.join(ROOT, f), 'utf8')); ok(true, f + ' 語法'); } catch (e) { ok(false, f + ' 語法錯誤 ' + e.message); } });
  });
}

// ═════════════ 變異檢查 ═════════════
async function secMut() {
  sect('[MUT] 變異檢查（拿掉規則 → 對應斷言必須失敗）');
  const mutations = [
    ['自我覆核護欄', ["if (rec.proposedBy === ctx.me) return fail(res, 403, 'SELF_REVIEW_FORBIDDEN', '不能覆核自己提出的變更，請由另一位管理員核可');", ''], ['selfApproveBlocked', 'selfStillPending']],
    ['基準雜湊比對', ['if (hashOf(spec.read(data)) !== rec.baseHash) {', 'if (false) {'], ['staleBlocked']],
    ['逾期結算', ['Date.parse(r.expiresAt) <= t', 'Date.parse(r.expiresAt) < 0'], ['expiryBlocked']],
    ['退回須有意見', ['if (Array.from(comment).length < 2) return fail(res, 400, \'COMMENT_REQUIRED\'', 'if (false) return fail(res, 400, \'COMMENT_REQUIRED\''], ['rejectNeedsComment']],
    ['套用失敗還原', ["backups.forEach(([k, val]) => { data[k] = val; });", ''], ['applyFailRollback']],
    ['管理員不足 2 人', ['if (activeAdmins().length < 2) {', 'if (false) {'], ['needsTwoAdmins']],
  ];
  for (const [name, mut, expectFail] of mutations) {
    const res = await unitRules(loadCR(mut));
    ok(expectFail.some(k => res[k] !== true), `變異「${name}」→ 對應單元斷言失敗（抓得到）：${expectFail.join('/')}`);
  }
  // 伺服器層：開關關閉＝逐位元相同（把 PUT 改成永遠走提案）
  const f = path.join(ROOT, 'lib', 'quoteRoutes.js');
  const orig = fs.readFileSync(f, 'utf8');
  const a = "if (CR.enabled()) return CR.propose(res, 'config'";
  if (!orig.includes(a)) throw new Error('變異錨點不存在（quoteRoutes）');
  try {
    fs.writeFileSync(f, orig.replace(a, "if (true) return CR.propose(res, 'config'"));
    await fresh(undefined);
    await loginAll(['_s_admin']);
    const r = await http('_s_admin', 'PUT', CFG, CLS_A);
    ok(r.status !== 200, `變異「開關關閉仍走覆核」→ OFF 的 PUT 回 ${r.status}（≠200）：OFF 逐位元相同斷言會失敗（抓得到）`);
  } finally {
    fs.writeFileSync(f, orig);
    await stopServer();
  }
  ok(fs.readFileSync(f, 'utf8') === orig, '變異後 lib/quoteRoutes.js 已還原');
}

// ═════════════ 截圖用伺服器 ═════════════
async function serveMode() {
  await fresh(ON);
  await loginAll(['_s_admin', '_s_admin2']);
  await propose('_s_admin', 'config', CLS_A);
  await propose('_s_admin2', 'flow', { buSettings: { ERP: { hasMgr2Step: true } }, exceptions: [{ id: 'ex-1', username: '_s_m2', bu: 'ERP', role: 'manager2', action: 'skip_mgr1', label: '' }] });
  let r = await propose('_s_admin2', 'config', CLS_B);
  const x = await propose('_s_admin', 'config', cfgBody({ roster: { gm: ['_s_gm2'] }, productClasses: { 'ERP顧問': { cls: 'ot' } } }));
  await http('_s_admin2', 'POST', REQ + '/' + x.json.requestId + '/reject', { comment: '名單與歸類需一併確認' });
  const y = await propose('_s_admin', 'config', cfgBody({ roster: { gm: ['_s_gm2'] } }));
  await http('_s_admin', 'POST', REQ + '/' + y.json.requestId + '/cancel', {});
  console.log('SERVE_READY pid=' + childPid + ' port=' + PORT + ' pw=' + (process.env.DR_SHOW_PW ? PW : '(設 DR_SHOW_PW=1 才印；預設不把密碼寫進輸出)'));
  await new Promise(() => {});
}

// ═════════════ 主程式 ═════════════
const SECTIONS = { off: secOff, on: secOn, stale: secStale, expiry: secExpiry, oneadmin: secOneAdmin, static: secStatic, unit: secUnit, ui: secUi, mut: secMut };
(async () => {
  const only = (process.env.DR_ONLY || '').split(',').map(s => s.trim()).filter(Boolean);
  try {
    if (process.env.DR_SERVE) { await serveMode(); return; }
    for (const k of Object.keys(SECTIONS)) {
      if (k === 'mut' && process.env.DR_SKIP_MUT) continue;
      if (only.length && !only.includes(k)) continue;
      try { await SECTIONS[k](); } catch (e) { fail++; failures.push(`${k} 例外：${e && e.stack || e}`); console.log('  ✗ 例外', e && e.stack || e); }
    }
  } finally {
    await stopServer();
    cleanupIn(ROOT);
  }
  console.log(`\n結果：${pass} 通過，${fail} 失敗；啟動過的伺服器 PID：${pids.join(',')}`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); process.exit(1); }
})().catch(async e => { console.error(e); await stopServer(); cleanupIn(ROOT); process.exit(1); });
