#!/usr/bin/env node
'use strict';
/**
 * 秘書跨 BU 唯讀查詢報價單（2026-10-14 業主決定）端對端檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 * 做法：在「本工作樹」建立臨時 auth.json / data.json（假資料 + _s_* 臨時帳號），自行以 PORT=XBU_PORT（預設 3171）
 * 啟動 server.js，跑完後只用 PID 關閉並刪除臨時檔。
 * 守門：工作樹若已有 auth.json / data.json / audit.log.json（可能是真實資料），一律拒絕執行。
 * 禁用埠：3000 / 3020 / 3030 / 3101-3161。
 * 稽核節流（10 分鐘）用「預載腳本偏移 Date.now()」模擬時間經過，不必真的等。
 *
 *   node scripts/check-secretary-cross-bu.js
 *   XBU_SKIP_MUT=1 node scripts/check-secretary-cross-bu.js   # 略過變異檢查
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.XBU_PORT || '3171', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3161)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const guardFiles = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));

// ───────── 沙盒資料 ─────────
const PW = 'Xbu' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
function mkQuote(tag, owner, extra) {
  return Object.assign({
    id: 'q-' + tag, quoteNo: 'SB-' + tag, company: '沙盒公司' + tag, projectName: '沙盒專案' + tag, owner, status: 'draft', createdAt: nowIso, updatedAt: nowIso,
    quoteDate: nowIso.slice(0, 10), products: [], contactName: '測試聯絡人',
    items: [{ lid: 'l-' + tag, desc: '沙盒品項', unit: '式', qty: 1, unitPrice: 100000, cost: 60000 }],
    discountType: 'none', discountValue: 0,
  }, extra || {});
}
const OWNERS = { E: '_s_u_erp', E2: '_s_u_erp2', I: '_s_u_its', M: '_s_u_mdm', C: '_s_u_crm', S: '_s_u_sec', N: '_s_u_nobu' };
const ALL_TAGS = ['QE', 'QE2', 'QI', 'QM', 'QC', 'QS', 'QN', 'QP'];
const ownerOfTag = { QE: OWNERS.E, QE2: OWNERS.E2, QI: OWNERS.I, QM: OWNERS.M, QC: OWNERS.C, QS: OWNERS.S, QN: OWNERS.N, QP: OWNERS.I };
function seedFiles() {
  const users = [
    mkUser('_s_adm', 'admin'),
    mkUser('_s_m1', 'manager1', { bu: ['ERP'] }),
    mkUser('_s_u_erp', 'user', { bu: ['ERP'], supervisor: '_s_m1' }),
    mkUser('_s_u_erp2', 'user', { bu: ['ERP'], supervisor: '_s_m1' }),
    mkUser('_s_u_its', 'user', { bu: ['ITS'] }),
    mkUser('_s_u_mdm', 'user', { bu: ['MDM'] }),
    mkUser('_s_u_crm', 'user', { bu: ['CRM'] }),
    mkUser('_s_u_sec', 'user', { bu: ['資安'] }),
    mkUser('_s_u_nobu', 'user', { bu: [] }),
    mkUser('_s_sec_erp', 'secretary', { bu: ['ERP'] }),
    mkUser('_s_sec_none', 'secretary', { bu: [] }),
    mkUser('_s_sec_all4', 'secretary', { bu: ['ERP', 'ITS', 'MDM', 'CRM'] }),
    mkUser('_s_sec_off', 'secretary', { bu: ['ERP'], active: false }),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const pend = (assignee) => ({ state: 'pending', rulesVersion: 1, hash: 'h', submittedAt: nowIso, submittedBy: OWNERS.I, steps: [{ tier: 'mgr1', label: '一級主管', assignee, status: 'pending' }, { tier: 'gm', label: '總經理', assignee: null, status: 'waiting' }], cur: 0, board: null, derived: null, history: [{ at: nowIso, by: OWNERS.I, action: 'SUBMIT', comment: '', tier: '' }] });
  const d = {
    contacts: [
      { id: 'c-erp', name: '聯絡甲', company: '沙盒公司QE', owner: OWNERS.E, bu: 'ERP', email: 'a@sandbox.test', phone: '0200000001', createdAt: nowIso },
      { id: 'c-its', name: '聯絡乙', company: '沙盒公司QI', owner: OWNERS.I, bu: 'ITS', email: 'b@sandbox.test', phone: '0200000002', createdAt: nowIso },
    ],
    companies: [],
    opportunities: [
      { id: 'o-erp', title: 'ERP 商機', company: '沙盒公司QE', owner: OWNERS.E, category: 'ERP', stage: 'B', amount: 100, expectedDate: new Date().getFullYear() + '-12-01', createdAt: nowIso },
      { id: 'o-its', title: 'ITS 商機', company: '沙盒公司QI', owner: OWNERS.I, category: 'ITS', stage: 'B', amount: 100, expectedDate: new Date().getFullYear() + '-12-01', createdAt: nowIso },
    ],
    visits: [
      { id: 'v-erp', visitDate: '2026-01-01', owner: OWNERS.E, contactName: '聯絡甲', topic: '沙盒', createdAt: nowIso },
      { id: 'v-its', visitDate: '2026-01-01', owner: OWNERS.I, contactName: '聯絡乙', topic: '沙盒', createdAt: nowIso },
    ],
    contracts: [], receivables: [], callins: [], keyAccounts: [], groups: [],
    quotations: [
      mkQuote('QE', OWNERS.E), mkQuote('QE2', OWNERS.E2), mkQuote('QI', OWNERS.I), mkQuote('QM', OWNERS.M),
      mkQuote('QC', OWNERS.C), mkQuote('QS', OWNERS.S), mkQuote('QN', OWNERS.N),
      mkQuote('QP', OWNERS.I, { approval: pend('_s_m1') }),   // 簽核中（ITS 業務；一級主管＝_s_m1）
    ],
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
const CLOCK_FILE = path.join(os.tmpdir(), 'xbu_clock_' + process.pid + '.txt');
const PRELOAD = path.join(ROOT, '__xbu_clock_preload.js');
function cleanupFiles() {
  FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  try { fs.unlinkSync(PRELOAD); } catch { /* 不存在 */ }
  try { fs.unlinkSync(CLOCK_FILE); } catch { /* 不存在 */ }
}
function setClockOffsetMs(ms) { fs.writeFileSync(CLOCK_FILE, String(ms)); }

// ───────── 伺服器 ─────────
let child = null, childPid = null;
async function startServer(serverFile) {
  fs.writeFileSync(PRELOAD, "const fs=require('fs');const f=process.env.XBU_CLOCK_FILE;const real=Date.now.bind(Date);let cache=0,at=0;Date.now=function(){const r=real();if(r-at>200){at=r;try{cache=parseInt(fs.readFileSync(f,'utf8'),10)||0;}catch(e){cache=0;}}return r+cache;};\n");
  setClockOffsetMs(0);
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE, XBU_CLOCK_FILE: CLOCK_FILE });
  delete env.DATABASE_URL; delete env.MAIL_MODE;
  child = spawn(process.execPath, ['-r', PRELOAD, serverFile || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.XBU_DEBUG) process.stdout.write('[srv] ' + d); };
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
async function fresh(serverFile) { await stopServer(); cleanupFiles(); seedFiles(); await startServer(serverFile); jar = {}; }

// ───────── HTTP ─────────
let jar = {};
async function http(who, method, url, body) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, buf, text: buf.toString('utf8'), headers: res.headers, setCookie: res.headers.getSetCookie ? res.headers.getSetCookie() : [] };
}
async function login(who) {
  const r = await http(null, 'POST', '/api/login', { username: who, password: PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
async function loginAll(names) { for (const n of names) { const r = await login(n); if (r.status !== 200) throw new Error('登入失敗 ' + n + ' ' + r.status); } }
const idsOf = (r) => (Array.isArray(r.json) ? r.json.map(q => q.id.replace(/^q-/, '')) : null);
const sorted = (a) => (a || []).slice().sort();
async function logsOf(action) {
  await sleep(300);
  const r = await http('_s_adm', 'GET', '/api/admin/logs');
  return (Array.isArray(r.json) ? r.json : []).filter(l => l.action === action);
}
const ALL_SORTED = sorted(ALL_TAGS);

// ═════════════ 1) 清單／單筆可視範圍 ═════════════
async function sec1() {
  console.log('[1] 可視範圍（清單與單筆）');
  await loginAll(['_s_adm', '_s_sec_erp', '_s_sec_none', '_s_sec_all4', '_s_m1', '_s_u_erp', '_s_u_erp2', '_s_u_its']);
  let r = await http('_s_sec_erp', 'GET', '/api/quotations');
  eq(r.status, 200, '1.1 ERP 秘書 GET /api/quotations 200');
  eq(sorted(idsOf(r)), ALL_SORTED, '1.2 掛 ERP 的秘書在清單看到全部 BU（ERP/ITS/MDM/CRM/資安/無 BU）的報價單');
  for (const tag of ['QI', 'QM', 'QC', 'QS', 'QN']) {
    r = await http('_s_sec_erp', 'GET', '/api/quotations/q-' + tag);
    eq(r.status, 200, `1.3 ERP 秘書以 id 開啟 ${tag}（非自己 BU）200`);
  }
  r = await http('_s_sec_none', 'GET', '/api/quotations');
  eq(sorted(idsOf(r)), ALL_SORTED, '1.4 沒有任何 BU 的秘書也看到全部');
  r = await http('_s_sec_none', 'GET', '/api/quotations/q-QS');
  eq(r.status, 200, '1.5 沒有 BU 的秘書以 id 開啟 資安 單 200');
  r = await http('_s_sec_all4', 'GET', '/api/quotations');
  eq(sorted(idsOf(r)), ALL_SORTED, '1.6 掛滿舊四個 BU 的秘書：行為不變（仍看全部，含 資安）');
  r = await login('_s_sec_off');
  ok(r.status !== 200, '1.7 停用的秘書無法登入 (' + r.status + ')');
  r = await http('_s_u_erp', 'GET', '/api/quotations');
  eq(sorted(idsOf(r)), ['QE'], '1.8 一般業務仍只看到自己的單');
  r = await http('_s_u_erp', 'GET', '/api/quotations/q-QI');
  eq(r.status, 403, '1.9 一般業務以 id 開別人的單 403');
  r = await http('_s_u_its', 'GET', '/api/quotations');
  eq(sorted(idsOf(r)), ['QI', 'QP'], '1.10 ITS 業務只看到自己的兩張單');
  r = await http('_s_m1', 'GET', '/api/quotations');
  eq(sorted(idsOf(r)), ['QE', 'QE2', 'QP'], '1.11 manager1 範圍不變：自己下屬（ERP 兩位業務）的單，加上「輪到他簽」的 QP（既有的簽核可視規則）；看不到 MDM/CRM/資安/無 BU 的單');
  r = await http('_s_m1', 'GET', '/api/quotations/q-QM');
  eq(r.status, 403, '1.12 manager1 開下屬以外的單仍 403');
  r = await http('_s_adm', 'GET', '/api/quotations');
  eq(sorted(idsOf(r)), ALL_SORTED, '1.13 管理員看全部（不變）');

  // BU 篩選用欄位
  r = await http('_s_sec_erp', 'GET', '/api/quotations');
  const by = Object.fromEntries(r.json.map(q => [q.id.replace(/^q-/, ''), q.ownerBu]));
  eq(by, { QE: ['ERP'], QE2: ['ERP'], QI: ['ITS'], QM: ['MDM'], QC: ['CRM'], QS: ['資安'], QN: [], QP: ['ITS'] }, '1.14 清單每筆帶 ownerBu（陣列；無 BU 的業務為空陣列）');
  r = await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
  ok(r.json && !('ownerBu' in r.json), '1.15 單筆 GET 回應不帶 ownerBu（單筆回應位元級不變，只有清單加欄位）');
  r = await http('_s_sec_erp', 'GET', '/api/quote-approval/config');
  eq(r.json && r.json.me && r.json.me.isSecretary, true, '1.16 /api/quote-approval/config 的 me.isSecretary=true（前端顯示 BU 徽章用）');
  r = await http('_s_u_erp', 'GET', '/api/quote-approval/config');
  eq(r.json && r.json.me && r.json.me.isSecretary, false, '1.17 一般業務 me.isSecretary=false');

  // 商機／聯絡人／拜訪的可視範圍沒有被放寬（秘書看不到 ITS 業務的資料）
  for (const [url, field, itsId] of [['/api/contacts', 'id', 'c-its'], ['/api/opportunities', 'id', 'o-its'], ['/api/visits', 'id', 'v-its']]) {
    r = await http('_s_sec_erp', 'GET', url);
    const arr = Array.isArray(r.json) ? r.json : (r.json && (r.json.items || r.json.data)) || [];
    ok(r.status === 200 && !arr.some(x => x[field] === itsId), `1.18 ERP 秘書 GET ${url} 看不到 ITS 業務的資料（未放寬）(${r.status})`);
  }
  // 商機連結的報價摘要端點同樣走同一個範圍判斷
  r = await http('_s_sec_erp', 'GET', '/api/opportunities/o-its/quotes');
  ok(r.status === 200, '1.19 /api/opportunities/:id/quotes 對秘書可用 (' + r.status + ')');
}

// ═════════════ 2) 秘書唯讀：寫入一律被擋 ═════════════
async function sec2() {
  console.log('[2] 秘書唯讀');
  await loginAll(['_s_adm', '_s_sec_erp', '_s_sec_none', '_s_m1']);
  const before = JSON.stringify((await http('_s_adm', 'GET', '/api/quotations/q-QP')).json);
  const beforeI = JSON.stringify((await http('_s_adm', 'GET', '/api/quotations/q-QI')).json);
  let r = await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
  const perm = r.json && r.json.perm;
  ok(perm && perm.canEdit === false && perm.canSubmit === false && perm.canWithdraw === false && perm.canApprove === false && perm.canReturn === false && perm.canReassign === false && perm.canEditCost === false && perm.isOwner === false,
    '2.1 其他 BU 單的 perm：不可編輯／送簽／撤回／簽核／退回／改派／填成本');
  r = await http('_s_sec_erp', 'POST', '/api/quotations', { company: '沙盒新單', projectName: 'x', items: [{ desc: 'a', qty: 1, unitPrice: 1 }], products: [] });
  eq(r.status, 403, '2.2 秘書 POST 建立報價單 403');
  eq(r.json && r.json.code, 'NO_CREATE_ROLE', '2.2b code=NO_CREATE_ROLE');
  r = await http('_s_sec_erp', 'PUT', '/api/quotations/q-QI', { company: '改名' });
  eq(r.status, 403, '2.3 秘書 PUT 修改其他 BU 的單 403');
  r = await http('_s_sec_none', 'PUT', '/api/quotations/q-QE', { company: '改名' });
  eq(r.status, 403, '2.3b 沒有 BU 的秘書 PUT 修改 403');
  r = await http('_s_sec_erp', 'PUT', '/api/quotations/q-QI/item-notes', { items: [{ lid: 'l-QI', note: 'x' }] });
  eq(r.status, 403, '2.4 秘書 PUT item-notes 403');
  r = await http('_s_sec_erp', 'PUT', '/api/quotations/q-QI/costs', { items: [{ lid: 'l-QI', cost: 1 }] });
  eq(r.status, 403, '2.5 秘書 PUT costs 403');
  r = await http('_s_sec_erp', 'DELETE', '/api/quotations/q-QI');
  eq(r.status, 403, '2.6 秘書 DELETE 403');
  r = await http('_s_sec_erp', 'POST', '/api/quotations/q-QI/submit', {});
  ok(r.status >= 400 && r.status < 500, '2.7 秘書 submit 被拒 (' + r.status + ')');
  r = await http('_s_sec_erp', 'POST', '/api/quotations/q-QP/withdraw', {});
  ok(r.status >= 400 && r.status < 500, '2.8 秘書 withdraw 被拒 (' + r.status + ')');
  r = await http('_s_sec_erp', 'POST', '/api/quotations/q-QP/approve', { comment: 'ok' });
  eq(r.status, 403, '2.9 秘書 approve 一級主管關的單 403（簽核規則不因可視範圍改變）');
  r = await http('_s_sec_erp', 'POST', '/api/quotations/q-QP/return', { comment: '退回' });
  eq(r.status, 403, '2.10 秘書 return 403');
  r = await http('_s_sec_erp', 'POST', '/api/quotations/q-QP/reassign', { username: '_s_m1', reason: 'x' });
  ok(r.status >= 400 && r.status < 500, '2.11 秘書 reassign 被拒 (' + r.status + ')');
  const after = JSON.stringify((await http('_s_adm', 'GET', '/api/quotations/q-QP')).json);
  const afterI = JSON.stringify((await http('_s_adm', 'GET', '/api/quotations/q-QI')).json);
  ok(before === after && beforeI === afterI, '2.12 所有寫入嘗試後，報價單內容（含簽核）位元級不變');
  // 簽核 inbox：非輪到她的單不會出現在待簽名單
  r = await http('_s_sec_erp', 'GET', '/api/quotations/inbox');
  eq(r.json && r.json.counts, { approvals: 0, costs: 0 }, '2.13 秘書的「待我處理」不因可視範圍變寬而多出項目');
  // 一級主管可簽（對照：簽核機制本身沒壞）
  r = await http('_s_m1', 'GET', '/api/quotations/inbox');
  eq(r.json && r.json.counts && r.json.counts.approvals, 1, '2.14 對照：輪到 m1 簽的 QP 仍在他的待簽名單（簽核機制未受影響；秘書只是唯讀）');
}

// ═════════════ 3) 匯出／預覽端點跟隨同一個範圍判斷 ═════════════
async function sec3() {
  console.log('[3] 匯出與預覽');
  await loginAll(['_s_adm', '_s_sec_erp', '_s_sec_none', '_s_u_erp']);
  const eps = [
    ['GET', '/api/quotations/q-QI/issue-info'],
    ['GET', '/api/quotations/q-QI/export'],
    ['GET', '/api/quotations/q-QI/pnl-preview'],
    ['GET', '/api/quotations/q-QI/export-pnl'],
    ['POST', '/api/quotations/q-QI/export-log', { format: 'pdf' }],
  ];
  for (const [m, u, b] of eps) {
    const a = await http('_s_adm', m, u, b);
    const s = await http('_s_sec_erp', m, u, b);
    const n = await http('_s_sec_none', m, u, b);
    const sales = await http('_s_u_erp', m, u, b);
    ok(a.status === 200, `3.x admin ${m} ${u} 200（基準）(${a.status})`);
    eq(s.status, a.status, `3.x ERP 秘書 ${m} ${u} 與管理員一致`);
    eq(n.status, a.status, `3.x 無 BU 秘書 ${m} ${u} 與管理員一致`);
    eq(sales.status, 403, `3.x 無關業務 ${m} ${u} 403（範圍仍擋住）`);
  }
  let r = await http('_s_sec_erp', 'GET', '/api/quotations/q-QI/export');
  ok(r.buf.slice(0, 2).toString() === 'PK', '3.1 秘書跨 BU 下載 Excel 為有效 xlsx（PK 檔頭）');
  r = await http('_s_sec_erp', 'GET', '/api/quotations/q-QI/seal?stampToken=x');   // 整併 audit3-document-issuance 後 seal 端點先要求憑證（無憑證一律 403），帶任意憑證才會走到範圍／核准判斷
  ok(r.status === 404 && r.json && r.json.code === 'NO_SEAL', '3.2 seal 端點對秘書走同一判斷：到達「未核准無章」404 而非 403');
  r = await http('_s_u_erp', 'GET', '/api/quotations/q-QI/seal?stampToken=x');   // 整併 audit3-document-issuance 後 seal 端點先要求憑證（無憑證一律 403），帶任意憑證才會走到範圍／核准判斷
  eq(r.status, 403, '3.3 無關業務 seal 403');
  // 稽核：匯出與 PDF 記錄
  const lg = await logsOf('EXPORT_QUOTATION');
  ok(lg.some(l => l.operator === '_s_sec_erp' && l.target === 'SB-QI'), '3.4 秘書跨 BU 匯出仍寫 EXPORT_QUOTATION 稽核');
  const lp = await logsOf('EXPORT_QUOTATION_PDF');
  ok(lp.some(l => l.operator === '_s_sec_erp' && l.target === 'SB-QI'), '3.5 秘書跨 BU PDF 記錄仍寫 EXPORT_QUOTATION_PDF 稽核');
}

// ═════════════ 4) 欄位可見度（成本／毛利）═════════════
async function sec4() {
  console.log('[4] 成本與毛利欄位');
  await loginAll(['_s_adm', '_s_sec_erp', '_s_m1']);
  const sec = (await http('_s_sec_erp', 'GET', '/api/quotations/q-QI')).json;
  const mgr = (await http('_s_m1', 'GET', '/api/quotations/q-QE')).json;   // 對照：同樣屬 COST_VIEW_ROLES 的 manager1 在自己範圍內的單
  const adm = (await http('_s_adm', 'GET', '/api/quotations/q-QI')).json;
  eq(sec.perm.canSeeCost, true, '4.1 秘書 perm.canSeeCost=true（COST_VIEW_ROLES 原本就含 secretary，與 BU 內相同）');
  eq(sec.perm.canSeePrice, true, '4.1b 秘書 perm.canSeePrice=true');
  ok(sec.items[0].cost === 60000 && sec.items[0].unitPrice === 100000, '4.2 品項逐列 cost／unitPrice 可見（BU 內原本就可見，未新增）');
  const secKeys = Object.keys(sec).sort(), mgrKeys = Object.keys(mgr).sort(), admKeys = Object.keys(adm).sort();
  eq(secKeys, mgrKeys, '4.3 秘書看他 BU 單的頂層欄位集合＝manager1 看自己範圍單的欄位集合（同為 COST_VIEW_ROLES，沒有多出欄位）');
  ok(secKeys.every(k => admKeys.includes(k)), '4.4 秘書的欄位是管理員欄位集合的子集（沒有秘書專屬的新欄位；ownerBu 兩者都有）');
  eq(Object.keys(sec.items[0]).sort(), Object.keys(adm.items[0]).sort(), '4.5 品項欄位集合與管理員一致');
  ok(!('contentHash' in sec), '4.6 秘書不拿到 contentHash（非可簽核人）');
  ok(sec.perm.isCostProvider === false && sec.perm.canEditCost === false, '4.7 秘書不能填成本');
  // 同一位秘書看 BU 內與 BU 外的單：欄位完全相同（可視範圍擴大沒有改變欄位集合）
  const inBu = (await http('_s_sec_erp', 'GET', '/api/quotations/q-QE')).json;
  eq(Object.keys(inBu).sort(), secKeys, '4.8 同一位秘書看自己 BU 與其他 BU 的單：欄位集合相同');
  console.log('  [資訊] 秘書看其他 BU 單的頂層欄位：' + secKeys.join(', '));
  console.log('  [資訊] 品項欄位：' + Object.keys(sec.items[0]).sort().join(', '));
  console.log('  [資訊] perm 欄位：' + Object.keys(sec.perm).sort().join(', '));
}

// ═════════════ 5) 稽核 VIEW_QUOTE_CROSS_BU（節流）═════════════
async function sec5() {
  console.log('[5] 跨 BU 稽核');
  await loginAll(['_s_adm', '_s_sec_erp', '_s_sec_none', '_s_sec_all4', '_s_u_erp']);
  // 清單與搜尋不記
  await http('_s_sec_erp', 'GET', '/api/quotations');
  await http('_s_sec_erp', 'GET', '/api/quotations/inbox');
  let lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.length, 0, '5.1 只看清單／收件匣：不記 VIEW_QUOTE_CROSS_BU');
  // 自己 BU 的單不記
  await http('_s_sec_erp', 'GET', '/api/quotations/q-QE');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.length, 0, '5.2 開啟自己 BU（ERP）的單：不記');
  // 開別的 BU：記一筆；重複開不再記
  await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
  await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
  await http('_s_sec_erp', 'GET', '/api/quotations/q-QI/issue-info');
  await http('_s_sec_erp', 'GET', '/api/quotations/q-QI/export');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.length, 1, '5.3 同一秘書同一張單，GET／預覽／匯出多次只記 1 筆（10 分鐘節流）');
  ok(lg[0] && lg[0].operator === '_s_sec_erp' && lg[0].target === 'SB-QI' && /SB-QI/.test(lg[0].detail) && /ITS/.test(lg[0].detail), '5.4 稽核內容含單號與業務 BU（ITS）(' + (lg[0] && lg[0].detail) + ')');
  // 不同單 → 另記
  await http('_s_sec_erp', 'GET', '/api/quotations/q-QS');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.length, 2, '5.5 開另一張跨 BU 的單 → 另記一筆');
  ok(lg.some(l => l.target === 'SB-QS' && /資安/.test(l.detail)), '5.6 資安 單的稽核記業務 BU＝資安');
  // 不同秘書各自記
  await http('_s_sec_none', 'GET', '/api/quotations/q-QI');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.filter(l => l.operator === '_s_sec_none').length, 1, '5.7 另一位秘書開同一張單 → 各自記一筆');
  await http('_s_sec_none', 'GET', '/api/quotations/q-QN');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  ok(lg.some(l => l.target === 'SB-QN' && /未設定/.test(l.detail)), '5.8 業務沒有 BU 的單 → 記為「未設定」');
  // 掛滿四個 BU 的秘書：開 資安 單仍是 BU 外
  await http('_s_sec_all4', 'GET', '/api/quotations/q-QE');
  await http('_s_sec_all4', 'GET', '/api/quotations/q-QS');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.filter(l => l.operator === '_s_sec_all4').map(l => l.target), ['SB-QS'], '5.9 掛滿四個 BU 的秘書：只有 資安 單算跨 BU');
  // 非秘書不記
  await http('_s_u_erp', 'GET', '/api/quotations/q-QE');
  await http('_s_adm', 'GET', '/api/quotations/q-QI');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.filter(l => l.operator === '_s_u_erp' || l.operator === '_s_adm').length, 0, '5.10 業務／管理員不記');
  // 被擋的請求（403）不記
  const n0 = lg.length;
  await http('_s_u_erp', 'GET', '/api/quotations/q-QI');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  eq(lg.length, n0, '5.11 被拒絕的請求不記');
  // 毛利預覽也走同一支
  await http('_s_sec_none', 'GET', '/api/quotations/q-QM/pnl-preview');
  lg = await logsOf('VIEW_QUOTE_CROSS_BU');
  ok(lg.some(l => l.operator === '_s_sec_none' && l.target === 'SB-QM'), '5.12 毛利預覽（pnl-preview）也記跨 BU 稽核');
  // 時間過了 10 分鐘 → 再記一筆
  setClockOffsetMs(11 * 60 * 1000);
  await sleep(450);   // 預載腳本每 200ms 重讀偏移檔
  const nBefore = (await logsOf('VIEW_QUOTE_CROSS_BU')).filter(l => l.operator === '_s_sec_erp' && l.target === 'SB-QI').length;
  await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
  const nAfter = (await logsOf('VIEW_QUOTE_CROSS_BU')).filter(l => l.operator === '_s_sec_erp' && l.target === 'SB-QI').length;
  eq([nBefore, nAfter], [1, 2], '5.13 節流視窗（10 分鐘）過後再開同一張單 → 再記一筆');
  setClockOffsetMs(0);
}

// ═════════════ 6) 前端：BU 篩選／徽章（靜態＋執行）═════════════
async function sec6() {
  console.log('[6] 前端 BU 篩選與徽章');
  const html = fs.readFileSync(path.join(ROOT, '_client', 'index.html'), 'utf8');
  const js = fs.readFileSync(path.join(ROOT, '_client', 'quote.js'), 'utf8');
  ok(/<select id="quoteBuFilter"[^>]*style="display:none"/.test(html), '6.1 index.html 有 #quoteBuFilter 且預設隱藏');
  ok(/全部 BU/.test(html), '6.2 篩選第一項為「全部 BU」');
  // 取出函式本體在 vm 裡執行（不依賴瀏覽器）
  const grab = (re) => { const m = js.match(re); if (!m) throw new Error('找不到：' + re); return m[0]; };
  const src = [
    grab(/const QUOTE_BU_ORDER = [^\n]*\n/),
    grab(/function _qBuSet\(list\) \{[\s\S]*?\n\}\n/),
    grab(/function quoteBuBadges\(q\) \{[\s\S]*?\n\}\n/),
  ].join('\n');
  const sandbox = { escapeHtml: (s) => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])) };
  vm.createContext(sandbox);
  vm.runInContext(src + '\nthis._qBuSet=_qBuSet;this.quoteBuBadges=quoteBuBadges;', sandbox);
  eq(sandbox._qBuSet([{ ownerBu: ['ITS'] }, { ownerBu: ['ERP', '資安'] }, { ownerBu: [] }, {}]), ['ERP', 'ITS', '資安'], '6.3 _qBuSet：去重並依固定順序（ERP／ITS／MDM／CRM／資安）');
  eq(sandbox._qBuSet([{ ownerBu: ['ERP'] }, { ownerBu: ['ERP'] }]), ['ERP'], '6.4 只有單一 BU → 只有 1 個（篩選下拉不顯示）');
  const evil = sandbox.quoteBuBadges({ ownerBu: ['<img src=x onerror=alert(1)>'] });
  ok(!/<img/.test(evil) && /&lt;img/.test(evil), '6.5 徽章輸出跳脫 HTML（ownerBu 含標籤時不會變成元素）');
  ok(/未設定 BU/.test(sandbox.quoteBuBadges({ ownerBu: [] })), '6.6 無 BU 的業務顯示「未設定 BU」');
  ok(/_qSyncBuFilter\(\)/.test(js) && /sel\.style\.display = multi \? '' : 'none'/.test(js), '6.7 下拉只在涵蓋 2 個以上 BU 時顯示');
  ok(/if \(buFilter\)\s+list = list\.filter\(q => Array\.isArray\(q\.ownerBu\) && q\.ownerBu\.includes\(buFilter\)\)/.test(js), '6.8 篩選依 ownerBu 過濾');
  ok(/\(q\.ownerBu\)/.test(js) || /Array\.isArray\(q\.ownerBu\) \? q\.ownerBu : \[\]\)\.some\(b => String\(b\)\.toLowerCase\(\)\.includes\(search\)\)/.test(js), '6.9 搜尋框也比對 BU 字樣');
  // 語法
  try { new Function(js); ok(true, ''); } catch (e) { ok(false, '6.10 quote.js 語法 ' + e.message); }
}

// ═════════════ 7) 變異檢查 ═════════════
async function secM() {
  console.log('[M] 變異檢查');
  const tmpFiles = [];
  const mkMut = (name, libEdit) => {
    let srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
    let lib = fs.readFileSync(path.join(ROOT, 'lib', 'quoteRoutes.js'), 'utf8');
    if (!lib.includes(libEdit[0])) throw new Error('變異樣式找不到：' + name);
    lib = lib.replace(libEdit[0], () => libEdit[1]);
    const libFile = path.join(ROOT, 'lib', 'quoteRoutes.__mut_' + name + '.js');
    fs.writeFileSync(libFile, lib); tmpFiles.push(libFile);
    if (!srv.includes("require('./lib/quoteRoutes')")) throw new Error('server.js 找不到 quoteRoutes require');
    srv = srv.replace("require('./lib/quoteRoutes')", () => "require('./lib/quoteRoutes.__mut_" + name + ".js')");
    const f = path.join(ROOT, 'server.__mut_' + name + '.js');
    fs.writeFileSync(f, srv); tmpFiles.push(f);
    return f;
  };
  try {
    // 變異 1：把秘書範圍改回「BU 交集」
    {
      const f = mkMut('m1', ["if (ctx.role === 'secretary') return true; // 秘書：全部 BU（唯讀）", "if (ctx.role === 'secretary') { const mine = normalizeBu((ctx.users[ctx.me] || {}).bu); return normalizeBu((ctx.users[q.owner] || {}).bu).some(b => mine.includes(b)); }"]);
      await fresh(f);
      await loginAll(['_s_sec_erp']);
      const r = await http('_s_sec_erp', 'GET', '/api/quotations');
      ok(JSON.stringify(sorted(idsOf(r))) !== JSON.stringify(ALL_SORTED), '變異1（秘書範圍改回 BU 交集）→ ERP 秘書看不到全部：sec1 的 1.2 斷言會失敗（抓得到）');
    }
    // 變異 2：拿掉稽核節流
    {
      const f = mkMut('m2', ['if (crossBuSeen.get(key) > now - CROSS_BU_AUDIT_MS) return;', '']);
      await fresh(f);
      await loginAll(['_s_adm', '_s_sec_erp']);
      await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
      await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
      const n = (await logsOf('VIEW_QUOTE_CROSS_BU')).length;
      ok(n === 2, '變異2（拿掉節流）→ 同一張單重複記了 ' + n + ' 筆：sec5 的 5.3 斷言會失敗（抓得到）');
    }
    // 變異 3：讓秘書可編輯（permOf 的 canEdit 納入秘書）
    {
      const f = mkMut('m3', ["canEdit: (rel.isOwner || rel.isAdmin) && st !== 'pending',", "canEdit: (rel.isOwner || rel.isAdmin || rel.role === 'secretary') && st !== 'pending',"]);
      await fresh(f);
      await loginAll(['_s_sec_erp']);
      const r = await http('_s_sec_erp', 'GET', '/api/quotations/q-QI');
      ok(r.json && r.json.perm.canEdit === true, '變異3（秘書 canEdit）→ perm.canEdit 變 true：sec2 的 2.1 斷言會失敗（抓得到）');
    }
  } finally {
    await stopServer();
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } });
  }
}

// ═════════════ 主程式 ═════════════
const SECTIONS = { 1: sec1, 2: sec2, 3: sec3, 4: sec4, 5: sec5, 6: sec6, M: secM };
(async () => {
  cleanupFiles(); seedFiles();
  try {
    for (const k of ['1', '2', '3', '4', '5', '6', 'M']) {
      if (k === 'M' && process.env.XBU_SKIP_MUT) continue;
      if (k !== 'M') await fresh();
      try { await SECTIONS[k](); } catch (e) { fail++; failures.push(`第 ${k} 項例外：${e && e.stack || e}`); console.log('  ✗ 例外', e && e.stack || e); }
    }
  } finally {
    await stopServer();
    cleanupFiles();
  }
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); process.exit(1); }
})().catch(async e => { console.error(e); await stopServer(); cleanupFiles(); process.exit(1); });
