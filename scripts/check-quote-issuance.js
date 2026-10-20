#!/usr/bin/env node
'use strict';
/**
 * 正式文件發出機制（內控稽核第三波 ①）檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 *  A 模組單元（lib/quoteIssuance.js）：憑證簽章／竄改／過期／單次使用、短雜湊、序號、official/draft 判定、ALLOW 開關、verify
 *  B 端對端（真的 server.js，本機 JSON 模式，絕不連任何資料庫）：發出 → 蓋章憑證 → seal 端點 → 草稿 → 作廢後 → 無效單 → 權限 → 稽核 → 驗證端點 → 序號競爭
 *  C Excel：發出標記（序號／短雜湊／草稿標記）、與舊版模組輸出逐檔比對（未帶 issuance＝位元級相同；帶了＝只差 G4／G5 兩格與追加樣式）
 *  D 前端 PDF 流程（vm 假 DOM）：先取序號、失敗中止、頁尾、草稿浮水印且不蓋章、用憑證換章、畫布本身不變
 *  E 靜態接線：server.js 唯讀帳號豁免、sw 版號、核准面板發出紀錄區塊、預覽不再直接載章
 *  M 變異檢查：憑證單次使用、核准有效判定、先記錄後串流、seal 必須憑證、草稿標記、前端失敗中止
 *
 * 守門：工作樹若已有 auth.json / data.json / audit.log.json(l) （可能是真實資料），一律拒絕執行。
 * 做法：在本工作樹建臨時 auth.json / data.json（假資料 + _s_* 臨時帳號），以 PORT=QISS_PORT（預設 3181）啟動 server.js，
 * 跑完只用 PID 關閉並刪除臨時檔。禁用埠：3000 / 3020 / 3030 / 3101~3175。
 *
 *   node scripts/check-quote-issuance.js
 *   QISS_ONLY=A,C node scripts/check-quote-issuance.js
 *   QISS_SKIP_MUT=1 node scripts/check-quote-issuance.js
 */
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const os = require('os');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const zlib = require('zlib');
const { spawn, execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.QISS_PORT || '3181', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3175)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const DATA_FILES = ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json', 'mail-outbox.json'].map(f => path.join(ROOT, f));
const guard = DATA_FILES.filter(f => ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'contact_audit.json'].includes(path.basename(f)) && fs.existsSync(f));
if (guard.length) { console.log('拒絕：此工作樹已有', guard.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const JSZip = require(path.join(ROOT, 'node_modules', 'jszip'));
const QA = require(path.join(ROOT, 'lib', 'quoteApproval.js'));

const ONLY = process.env.QISS_ONLY ? new Set(process.env.QISS_ONLY.split(',')) : null;
const want = (k) => !ONLY || ONLY.has(k);
let pass = 0, fail = 0, muting = false, mutFail = 0;
const failures = [];
function ok(cond, name) {
  if (cond) { if (!muting) pass++; return; }
  if (muting) { mutFail++; if (process.env.QISS_MUTDBG) console.log('    caught:', name); return; }
  fail++; failures.push(name); console.log('  ✗ FAIL:', name);
}
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));
const b64u = (buf) => Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const unb64u = (s) => Buffer.from(String(s).replace(/-/g, '+').replace(/_/g, '/'), 'base64');

// ═════════════ 共用：假資料 ═════════════
const PW = 'Qi' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
function mkQuote(id, no, extra) {
  const q = Object.assign({
    id, quoteNo: no, owner: '_s_owner', company: '沙盒公司甲', contactName: '測試聯絡人', phone: '0200000001', address: '沙盒路 1 號', projectName: '沙盒專案',
    products: [], status: 'draft', oppId: 'O1', quoteDate: '2026-10-08', discountType: 'none', discountValue: 0, createdAt: nowIso, updatedAt: nowIso,
    items: [{ lid: 'L1-' + id, desc: '顧問服務', unit: '式', qty: 2, unitPrice: 5000, cost: 0 }, { lid: 'L2-' + id, desc: '軟體授權', unit: '套', qty: 1, unitPrice: 12000, cost: 0 }],
  }, extra || {});
  return q;
}
function approve(q) {
  q.approval = { state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: nowIso, submittedBy: q.owner, derived: null, steps: [], cur: 0, board: null, history: [] };
  return q;
}
/** 1x1 PNG（含 IEND，通過 sniffImage） */
function makePng() {
  const crcT = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; crcT[n] = c >>> 0; }
  const crc = (buf) => { let c = 0xffffffff; for (const b of buf) c = crcT[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const cr = Buffer.alloc(4); cr.writeUInt32BE(crc(td)); return Buffer.concat([len, td, cr]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(8, 0); ihdr.writeUInt32BE(8, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc(8 * (1 + 8 * 3), 0); for (let y = 0; y < 8; y++) { raw[y * 25] = 0; for (let x = 0; x < 8; x++) { raw[y * 25 + 1 + x * 3] = 200; } }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
const PNG = makePng();

// ═════════════ A：模組單元 ═════════════
async function secA(QIS) {
  console.log('[A] lib/quoteIssuance.js 單元');
  const SECRET = 'unit-test-secret-' + crypto.randomBytes(6).toString('hex');
  const saved = { j: process.env.JWT_SECRET, s: process.env.SESSION_SECRET, i: process.env.QUOTE_ISSUE_SECRET, a: process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE };
  process.env.JWT_SECRET = SECRET; delete process.env.QUOTE_ISSUE_SECRET; delete process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE;
  try {
    const mk = () => ({ quotations: [], quoteIssuances: [] });
    const data = mk();
    const q1 = approve(mkQuote('a1', 'QU-1')); data.quotations.push(q1);
    const q2 = mkQuote('a2', 'QU-2'); data.quotations.push(q2);
    const T0 = Date.parse('2026-10-14T01:00:00Z');
    let r = QIS.issue(data, { q: q1, format: 'pdf', me: 'u1', role: 'user', ip: '1.2.3.4', nowMs: T0 });
    ok(r.ok && r.rec.kind === 'official' && r.rec.approvedValid === true, '[A1] 核准有效 → official');
    eq(r.rec.issueNo, 'ISS-QU-1-001', '[A1] 發出序號格式 ISS-<單號>-<3 位序號>');
    eq(r.rec.contentHash, QA.contentHash(q1), '[A1] 紀錄存當下 contentHash');
    eq(r.rec.approvalHash, q1.approval.hash, '[A1] official 帶 approvalHash');
    eq(r.shortHash, crypto.createHash('sha256').update(r.rec.issueNo + ':' + r.rec.contentHash).digest('hex').slice(0, 12), '[A1] shortHash＝sha256(issueNo:contentHash) 前 12 碼');
    ok(/^[0-9a-f]{12}$/.test(r.shortHash), '[A1] shortHash 為 12 碼十六進位');
    ok(typeof r.stampToken === 'string' && r.stampToken.includes('.'), '[A1] official/pdf 有 stampToken');
    eq(r.rec.issuedAt, new Date(T0).toISOString(), '[A1] issuedAt 為伺服器時間');
    eq([q1.issuanceCount, q1.issuances.length], [1, 1], '[A1] 單上摘要快取');
    const r2 = QIS.issue(data, { q: q1, format: 'xlsx', me: 'u1', nowMs: T0 });
    eq([r2.rec.issueNo, r2.stampToken], ['ISS-QU-1-002', null], '[A2] 序號遞增；xlsx 不核發蓋章憑證');
    const r3 = QIS.issue(data, { q: q2, format: 'pdf', me: 'u1', nowMs: T0 });
    eq([r3.ok, r3.rec.kind, r3.rec.approvalHash, r3.rec.approvedValid, r3.stampToken], [true, 'draft', null, false, null], '[A3] 未核准 → draft、無 approvalHash、無憑證');
    eq(r3.rec.issueNo, 'ISS-QU-2-001', '[A3] 序號每張單各自從 001 起');
    const pn = QIS.issue(data, { q: q1, format: 'pnl-xlsx', me: 'u1', nowMs: T0 });
    eq([pn.rec.kind, pn.stampToken], ['internal', null], '[A4] 毛利分析 → internal、無憑證');
    eq(QIS.issue(data, { q: q1, format: 'docx', me: 'u1' }).code, 'BAD_FORMAT', '[A4] 不認得的格式');
    // 核准有效判定：改內容 → 雜湊不同 → draft
    const q1x = JSON.parse(JSON.stringify(q1)); q1x.items[0].unitPrice = 1;
    eq(QIS.issue(data, { q: q1x, format: 'pdf', me: 'u1' }).rec.kind, 'draft', '[A5] 核准後內容被改（雜湊不符）→ draft');
    const q1y = JSON.parse(JSON.stringify(q1)); q1y.approval.state = 'returned';
    eq(QIS.issue(data, { q: q1y, format: 'pdf', me: 'u1' }).rec.kind, 'draft', '[A5] 非 approved 狀態 → draft');
    // 憑證
    const t = r.stampToken;
    let c = QIS.consumeStamp(data, q1, t, 'u1', T0 + 1000);
    ok(c.ok, '[A6] 正確憑證通過');
    c = QIS.consumeStamp(data, q1, t, 'u1', T0 + 2000);
    ok(!c.ok && c.reason === 'used' && c.status === 403, '[A6] 第二次使用 → used');
    const r4 = QIS.issue(data, { q: q1, format: 'pdf', me: 'u1', nowMs: T0 });
    ok(!QIS.consumeStamp(data, q1, r4.stampToken, 'u1', T0 + 5 * 60 * 1000 + 1).ok && QIS.consumeStamp(data, q1, r4.stampToken, 'u1', T0 + 5 * 60 * 1000 + 1).reason === 'expired', '[A7] 超過 5 分鐘 → expired');
    ok(QIS.consumeStamp(data, q1, r4.stampToken, 'u1', T0 + 5 * 60 * 1000 - 1).ok, '[A7] 5 分鐘內仍有效（邊界）');
    const r5 = QIS.issue(data, { q: q1, format: 'pdf', me: 'u1', nowMs: T0 });
    eq(QIS.consumeStamp(data, q1, r5.stampToken, 'u2', T0).reason, 'wrong_user', '[A8] 別人拿去用 → wrong_user');
    eq(QIS.consumeStamp(data, { id: 'other', quoteNo: 'X', approval: null, items: [] }, r5.stampToken, 'u1', T0).reason, 'wrong_quote', '[A8] 拿去別張單 → wrong_quote');
    const parts = r5.stampToken.split('.');
    const flip = (s) => s.slice(0, -2) + (s.slice(-2) === 'AA' ? 'BB' : 'AA');
    eq(QIS.consumeStamp(data, q1, parts[0] + '.' + flip(parts[1]), 'u1', T0).reason, 'bad_signature', '[A9] 簽章被改 → bad_signature');
    const pl = JSON.parse(unb64u(parts[0]).toString()); pl.e += 9e9;
    eq(QIS.consumeStamp(data, q1, b64u(JSON.stringify(pl)) + '.' + parts[1], 'u1', T0).reason, 'bad_signature', '[A9] 內容被改（延長效期）→ bad_signature');
    eq(QIS.consumeStamp(data, q1, 'garbage', 'u1', T0).reason, 'malformed', '[A9] 亂碼 → malformed');
    eq(QIS.consumeStamp(data, q1, '', 'u1', T0).reason, 'malformed', '[A9] 空字串 → malformed');
    const q1z = JSON.parse(JSON.stringify(q1)); q1z.items[0].qty = 9;
    eq(QIS.consumeStamp(data, q1z, r5.stampToken, 'u1', T0).reason, 'hash_mismatch', '[A10] 單被改過（雜湊不符）→ hash_mismatch');
    // 金鑰獨立：換 JWT_SECRET 後舊憑證失效
    process.env.JWT_SECRET = SECRET + 'x';
    eq(QIS.consumeStamp(data, q1, r5.stampToken, 'u1', T0).reason, 'bad_signature', '[A11] 換金鑰後舊憑證失效');
    process.env.JWT_SECRET = SECRET;
    process.env.QUOTE_ISSUE_SECRET = 'dedicated-' + SECRET;
    eq(QIS.consumeStamp(data, q1, r5.stampToken, 'u1', T0).reason, 'bad_signature', '[A11] 設了專用金鑰後以專用金鑰簽發（與 JWT 金鑰分離）');
    delete process.env.QUOTE_ISSUE_SECRET;
    // 使用紀錄修剪
    data.quoteStampUses.push({ jti: 'old', issueNo: 'x', quoteId: 'a1', by: 'u', at: new Date(T0 - 3600e3).toISOString() });
    QIS.consumeStamp(data, q1, QIS.issue(data, { q: q1, format: 'pdf', me: 'u1', nowMs: T0 }).stampToken, 'u1', T0);
    ok(!data.quoteStampUses.some(u => u.jti === 'old'), '[A12] 超過效期的使用紀錄被修剪');
    // ALLOW 開關
    process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE = '0';
    const nb = data.quoteIssuances.length;
    const denied = QIS.issue(data, { q: q2, format: 'pdf', me: 'u1' });
    eq([denied.ok, denied.status, denied.code, data.quoteIssuances.length], [false, 409, 'NOT_APPROVED', nb], '[A13] ALLOW=false：未核准 → 409 NOT_APPROVED 且不留紀錄');
    ok(QIS.issue(data, { q: q1, format: 'pdf', me: 'u1' }).ok, '[A13] ALLOW=false：核准有效照常 official');
    ok(QIS.issue(data, { q: q1, format: 'pnl-xlsx', me: 'u1' }).ok, '[A13] ALLOW=false：內部文件不受影響');
    for (const v of ['false', 'no', 'OFF']) { process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE = v; ok(!QIS.allowUnapprovedIssue(), `[A13] 環境變數 ${v} → 禁止`); }
    for (const v of ['1', 'true', 'yes']) { process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE = v; ok(QIS.allowUnapprovedIssue(), `[A13] 環境變數 ${v} 不誤判為禁止`); }
    delete process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE;
    ok(QIS.allowUnapprovedIssue(), '[A13] 預設允許（與現況一致）');
    delete process.env.QUOTE_DRAFT_WATERMARK;
    ok(!QIS.draftWatermarkEnabled(), '[A13b] 草稿浮水印預設關');
    for (const v of ['1', 'true', 'ON']) { process.env.QUOTE_DRAFT_WATERMARK = v; ok(QIS.draftWatermarkEnabled(), `[A13b] QUOTE_DRAFT_WATERMARK=${v} → 啟用`); }
    for (const v of ['0', 'false', '', 'x']) { process.env.QUOTE_DRAFT_WATERMARK = v; ok(!QIS.draftWatermarkEnabled(), `[A13b] QUOTE_DRAFT_WATERMARK=${v || '(空)'} → 不啟用`); }
    delete process.env.QUOTE_DRAFT_WATERMARK;
    // stillValid／verify
    const recO = data.quoteIssuances.find(x => x.issueNo === 'ISS-QU-1-001');
    ok(QIS.stillValidOf(recO, q1), '[A14] official 內容未變且核准有效 → stillValid');
    ok(!QIS.stillValidOf(recO, q1z), '[A14] 內容改了 → 失效');
    const q1v = JSON.parse(JSON.stringify(q1)); q1v.approval = { state: 'none', hash: null };
    ok(!QIS.stillValidOf(recO, q1v), '[A14] 內容沒變但核准不在了 → official 失效');
    const recD = data.quoteIssuances.find(x => x.issueNo === 'ISS-QU-2-001');
    ok(QIS.stillValidOf(recD, q2), '[A14] draft 內容未變 → 有效');
    let v = QIS.verify(data, 'ISS-QU-1-001', q1, { shortHash: QIS.shortHashOf('ISS-QU-1-001', recO.contentHash) });
    ok(v.found && v.stillValid && v.hashMatches === true && v.mismatches.length === 0, '[A15] verify 吻合');
    v = QIS.verify(data, 'ISS-QU-1-001', q1, { shortHash: 'deadbeef0000' });
    ok(v.hashMatches === false && v.mismatches.includes('SHORT_HASH_DIFFERENT'), '[A15] verify：短雜湊不符被偵測');
    v = QIS.verify(data, 'ISS-QU-1-001', q1z, {});
    ok(!v.stillValid && v.mismatches.includes('CONTENT_CHANGED'), '[A15] verify：內容已變');
    v = QIS.verify(data, 'ISS-QU-1-001', q1v, {});
    ok(v.mismatches.includes('OFFICIAL_NOW_NOT_APPROVED'), '[A15] verify：official 但現在不是核准有效');
    const q2ap = approve(JSON.parse(JSON.stringify(q2)));
    v = QIS.verify(data, 'ISS-QU-2-001', Object.assign(JSON.parse(JSON.stringify(q2)), { approval: q2ap.approval }), {});
    ok(v.contentMatches && v.mismatches.includes('DRAFT_NOW_APPROVED') && v.record.kind === 'draft', '[A15] verify：草稿紀錄、現在已核准 → 標示 DRAFT_NOW_APPROVED（草稿不會變正式）');
    // 偽造紀錄（official 但 approvedValid=false）
    data.quoteIssuances.push({ issueNo: 'ISS-QU-9-001', quoteId: 'a1', quoteNo: 'QU-9', kind: 'official', format: 'pdf', contentHash: 'x', approvalHash: null, approvedValid: false, issuedBy: 'u', issuedAt: nowIso, ver: 1 });
    v = QIS.verify(data, 'ISS-QU-9-001', q1, {});
    ok(v.mismatches.includes('OFFICIAL_RECORD_NOT_APPROVED'), '[A15] verify：紀錄自相矛盾（official 卻無核准證據）被偵測');
    eq(QIS.verify(data, 'ISS-NOPE', q1, {}).found, false, '[A15] 查無序號');
    ok(!('ip' in QIS.pub(recO)) && !('tokenJti' in QIS.pub(recO)), '[A16] pub() 不含 ip／jti');
    // 競爭：同步連發 50 次，序號不重複
    const d2 = mk(); const q9 = mkQuote('a9', 'QU-RACE'); d2.quotations.push(q9);
    const nos = []; for (let i = 0; i < 50; i++) nos.push(QIS.issue(d2, { q: q9, format: 'pdf', me: 'u' }).rec.issueNo);
    eq(new Set(nos).size, 50, '[A17] 連發 50 次序號全不重複');
    eq(nos[49], 'ISS-QU-RACE-050', '[A17] 序號單調遞增');
  } finally {
    for (const [k, v] of [['JWT_SECRET', saved.j], ['SESSION_SECRET', saved.s], ['QUOTE_ISSUE_SECRET', saved.i], ['QUOTE_ALLOW_UNAPPROVED_ISSUE', saved.a], ['QUOTE_DRAFT_WATERMARK', undefined]]) { if (v === undefined) delete process.env[k]; else process.env[k] = v; }
  }
}

// ═════════════ 伺服器管理 ═════════════
let child = null, childPid = null, extraEnv = {}, srvSecret = '';
const tmpFiles = [];
function cleanupFiles() { DATA_FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }
function seedFiles() {
  const users = [
    mkUser('_s_admin', 'admin'),
    mkUser('_s_owner', 'user', { bu: ['ERP'], canDownloadContacts: true }),
    mkUser('_s_other', 'user', { bu: ['ERP'] }),
    mkUser('_s_secro', 'secretary', { bu: ['ERP', 'ITS', 'MDM', 'CRM'], accessMode: 'view' }),
    mkUser('_s_sec', 'secretary', { bu: ['ERP', 'ITS', 'MDM', 'CRM'] }),
  ];
  const quotations = [
    approve(mkQuote('Q1', 'QU-202610-001')),
    mkQuote('Q2', 'QU-202610-002'),
    approve(mkQuote('Q3', 'QU-202610-003')),
    approve(mkQuote('Q4', 'QU-202610-004', { invalid: true })),
    mkQuote('Q5', 'QU-202610-005'),
    approve(mkQuote('Q6', 'QU-202610-006')),
  ];
  const d = {
    contacts: [], companies: [], opportunities: [{ id: 'O1', owner: '_s_owner', stage: 'C', company: '沙盒公司甲', product: 'P', amount: '100', bu: 'ERP' }],
    visits: [], contracts: [], quotations, receivables: [], callins: [], keyAccounts: [], groups: [],
  };
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
  return quotations;
}
async function startServer(serverFile) {
  const base = Object.assign({}, process.env); delete base.AUDIT_EVENTS_TABLE; delete base.QUOTE_ALLOW_UNAPPROVED_ISSUE; delete base.QUOTE_ISSUE_SECRET;
  srvSecret = crypto.randomBytes(24).toString('hex');
  const env = Object.assign(base, extraEnv, { PORT: String(PORT), JWT_SECRET: srvSecret, SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE;
  await assertPortFree(PORT);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
  child = spawn(process.execPath, [serverFile || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.QISS_DEBUG) process.stdout.write('[srv] ' + d); };
  child.stdout.on('data', onData); child.stderr.on('data', onData);
  child.serverOut = () => out;
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
let jar = {};
async function http(who, method, url, body, opt) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fresh = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fresh) jar[who] = fresh.split(';')[0];
  return { status: res.status, json, buf, text: buf.toString('utf8'), headers: res.headers, setCookie: sc };
}
async function login(who) {
  const r = await http(null, 'POST', '/api/login', { username: who, password: PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
async function freshServer(serverFile, env) {
  await stopServer(); cleanupFiles(); extraEnv = env || {};
  const qs = seedFiles();
  await startServer(serverFile); jar = {};
  for (const n of ['_s_admin', '_s_owner', '_s_other', '_s_secro', '_s_sec']) { const r = await login(n); if (r.status !== 200) throw new Error('登入失敗 ' + n + ' ' + r.status + ' ' + r.text.slice(0, 120)); }
  const up = await http('_s_admin', 'PUT', '/api/quote-approval/seal', { dataUrl: 'data:image/png;base64,' + PNG.toString('base64') });
  if (up.status !== 200) throw new Error('上傳測試章失敗 ' + up.status + ' ' + up.text.slice(0, 200));
  return qs;
}
const issue = (who, id, format) => http(who, 'POST', `/api/quotations/${id}/issue`, { format: format || 'pdf' });
const seal = (who, id, token, viaHeader) => viaHeader
  ? fetch(BASE + `/api/quotations/${id}/seal`, { headers: { Cookie: jar[who], 'X-Stamp-Token': token } }).then(async r => ({ status: r.status, buf: Buffer.from(await r.arrayBuffer()), ct: r.headers.get('content-type') }))
  : fetch(BASE + `/api/quotations/${id}/seal` + (token === undefined ? '' : '?stampToken=' + encodeURIComponent(token)), { headers: { Cookie: jar[who] } }).then(async r => { const buf = Buffer.from(await r.arrayBuffer()); let j = null; try { j = JSON.parse(buf.toString('utf8')); } catch { /* 圖 */ } return { status: r.status, buf, json: j, ct: r.headers.get('content-type') }; });
function mintToken(payload) {
  const key = crypto.createHash('sha256').update('quote-issuance-stamp-token:v1:' + srvSecret).digest();
  const p = b64u(JSON.stringify(payload));
  return p + '.' + b64u(crypto.createHmac('sha256', key).update(p).digest());
}
const decodeToken = (t) => JSON.parse(unb64u(t.split('.')[0]).toString('utf8'));
const hashOf = (id) => { const d = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8')); return QA.contentHash(d.quotations.find(q => q.id === id)); };
const dataNow = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
async function getLogs() { const r = await http('_s_admin', 'GET', '/api/admin/logs?limit=500'); return Array.isArray(r.json) ? r.json : ((r.json && (r.json.logs || r.json.rows)) || []); }

// ═════════════ B：端對端 ═════════════
async function secB() {
  console.log('[B] 端對端（server.js，3181）');
  await freshServer();
  // B1 official + 憑證
  let r = await issue('_s_owner', 'Q1');
  ok(r.status === 200, `[B1] 核准有效單 POST /issue 200（${r.status} ${r.text.slice(0, 120)}）`);
  const j = r.json || {};
  eq([j.kind, j.issueNo, j.format], ['official', 'ISS-QU-202610-001-001', 'pdf'], '[B1] kind=official、序號、格式');
  ok(/^[0-9a-f]{12}$/.test(j.shortHash || '') && j.shortHash === crypto.createHash('sha256').update(j.issueNo + ':' + hashOf('Q1')).digest('hex').slice(0, 12), '[B1] shortHash＝sha256(issueNo:contentHash) 前 12 碼');
  ok(typeof j.stampToken === 'string' && j.stampToken.length > 40, '[B1] 有 stampToken');
  eq(j.watermark, false, '[B1] 浮水印開關預設關：POST /issue 回 watermark:false（由伺服器決定）');
  { const vr = await http('_s_admin', 'GET', '/api/admin/version'); eq(vr.json && vr.json.quoteIssuance, { draftIssueAllowed: true, draftWatermark: false }, '[B1] /api/admin/version 回報開關狀態（只有布林）：預設可發草稿、浮水印關'); }
  ok(!isNaN(Date.parse(j.issuedAt)) && Math.abs(Date.now() - Date.parse(j.issuedAt)) < 60000, '[B1] issuedAt 為伺服器時間');
  eq(r.headers.get('cache-control'), 'no-store', '[B1] 回應 no-store');
  // B2 seal：沒有憑證 → 403；有憑證一次；第二次 403
  let s = await seal('_s_owner', 'Q1');
  eq([s.status, s.json && s.json.code], [403, 'STAMP_TOKEN_REQUIRED'], '[B2] seal 沒有憑證 → 403 STAMP_TOKEN_REQUIRED');
  s = await seal('_s_owner', 'Q1', '');
  eq([s.status, s.json && s.json.code], [403, 'STAMP_TOKEN_REQUIRED'], '[B2] seal 空憑證 → 403 STAMP_TOKEN_REQUIRED');
  s = await seal('_s_owner', 'Q1', j.stampToken);
  ok(s.status === 200 && /image\/png/.test(s.ct || '') && s.buf.equals(PNG), '[B2] seal 帶有效憑證 → 200 且就是上傳的章圖');
  s = await seal('_s_owner', 'Q1', j.stampToken);
  eq([s.status, s.json && s.json.code, s.json && s.json.reason], [403, 'STAMP_TOKEN_INVALID', 'used'], '[B2] 同一憑證第二次 → 403（used）');
  ok(!JSON.stringify(s.json).includes(j.stampToken), '[B2] 錯誤回應不回傳憑證');
  // header 方式
  let t2 = (await issue('_s_owner', 'Q1')).json;
  s = await seal('_s_owner', 'Q1', t2.stampToken, true);
  ok(s.status === 200 && s.buf.equals(PNG), '[B2] 用 X-Stamp-Token 標頭也可');
  // B3 過期／換單／竄改／換人／雜湊
  const t3 = (await issue('_s_owner', 'Q1')).json;
  const pl = decodeToken(t3.stampToken);
  s = await seal('_s_owner', 'Q1', mintToken(Object.assign({}, pl, { e: Date.now() - 1000 })));
  eq([s.status, s.json && s.json.reason], [403, 'expired'], '[B3] 簽章正確但已過期 → 403 expired');
  s = await seal('_s_owner', 'Q3', t3.stampToken);
  eq([s.status, s.json && s.json.reason], [403, 'wrong_quote'], '[B3] Q1 的憑證拿去取 Q3 的章 → 403 wrong_quote');
  const sig = t3.stampToken.split('.');
  s = await seal('_s_owner', 'Q1', sig[0] + '.' + sig[1].slice(0, -3) + (sig[1].endsWith('AAA') ? 'BBB' : 'AAA'));
  eq([s.status, s.json && s.json.reason], [403, 'bad_signature'], '[B3] 竄改簽章 → 403 bad_signature');
  s = await seal('_s_owner', 'Q1', mintToken(Object.assign({}, pl, { h: 'f'.repeat(64) })));
  eq([s.status, s.json && s.json.reason], [403, 'hash_mismatch'], '[B3] 簽章正確但雜湊不是目前版本 → 403 hash_mismatch');
  s = await seal('_s_sec', 'Q1', t3.stampToken);
  eq([s.status, s.json && s.json.reason], [403, 'wrong_user'], '[B3] 別人拿去用 → 403 wrong_user');
  s = await seal('_s_owner', 'Q1', mintToken(Object.assign({}, pl, { j: 'nope' + pl.j })));
  eq([s.status, s.json && s.json.reason], [403, 'no_issuance'], '[B3] 簽章正確但 jti 對不上發出紀錄 → no_issuance');
  s = await seal('_s_owner', 'Q1', t3.stampToken);
  eq(s.status, 200, '[B3] 上述失敗嘗試沒有消耗合法憑證（之後仍可使用一次）');
  s = await seal('_s_other', 'Q1', t3.stampToken);
  eq(s.status, 403, '[B3] 無權限者（非擁有者業務）→ 403');
  // B4 未核准：draft、沒有憑證
  r = await issue('_s_owner', 'Q2');
  eq([r.status, r.json.kind, r.json.stampToken, r.json.issueNo], [200, 'draft', undefined, 'ISS-QU-202610-002-001'], '[B4] 未核准 → 200 draft、無 stampToken');
  s = await seal('_s_owner', 'Q2', t3.stampToken);
  ok(s.status === 403 || s.status === 404, `[B4] 未核准單用別張單的憑證取章被拒（${s.status}）`);
  s = await seal('_s_owner', 'Q2');
  eq([s.status, s.json && s.json.code], [403, 'STAMP_TOKEN_REQUIRED'], '[B4] 未核准單 seal 沒憑證 → 403 STAMP_TOKEN_REQUIRED（不洩漏核准狀態）');
  // B5 序號：同單遞增、異單各自
  const q1b = await issue('_s_owner', 'Q1', 'xlsx');
  ok(q1b.status === 200 && /ISS-QU-202610-001-00[4-9]|ISS-QU-202610-001-01\d/.test(q1b.json.issueNo) && q1b.json.stampToken === undefined, '[B5] xlsx 發出：序號接續、不核發憑證');
  const lst = (await http('_s_owner', 'GET', '/api/quotations/Q1/issuances')).json;
  const seqs = lst.issuances.map(x => parseInt(x.issueNo.slice(-3), 10));
  ok(seqs.every((v, i) => i === 0 || v === seqs[i - 1] - 1), '[B5] 清單新→舊、序號連號無重複');
  ok(new Set(lst.issuances.map(x => x.issueNo)).size === lst.issuances.length, '[B5] 序號不重複');
  // B6 作廢後發出
  let r3 = await issue('_s_owner', 'Q3');
  eq(r3.json.kind, 'official', '[B6] Q3 核准有效 → official');
  const edit = await http('_s_owner', 'PUT', '/api/quotations/Q3', { items: [{ lid: 'L1-Q3', desc: '顧問服務', unit: '式', qty: 2, unitPrice: 6000, cost: 0 }, { lid: 'L2-Q3', desc: '軟體授權', unit: '套', qty: 1, unitPrice: 12000, cost: 0 }], confirmVoid: true });
  ok(edit.status === 200, `[B6] 業務改價（confirmVoid）成功（${edit.status} ${edit.text.slice(0, 160)}）`);
  eq(edit.json && edit.json.approval && edit.json.approval.state, 'none', '[B6] 核准因改價作廢');
  let stale = await seal('_s_owner', 'Q3', r3.json.stampToken);
  ok(stale.status === 403 || stale.status === 404, `[B6] 作廢後舊憑證換不到章（${stale.status}）`);
  const r3b = await issue('_s_owner', 'Q3');
  eq([r3b.json.kind, r3b.json.stampToken, r3b.json.issueNo], ['draft', undefined, 'ISS-QU-202610-003-002'], '[B6] 作廢後再發出 → draft、序號 002、無憑證');
  const l3 = (await http('_s_owner', 'GET', '/api/quotations/Q3/issuances')).json.issuances;
  eq(l3.map(x => [x.issueNo, x.kind, x.stillValid]), [['ISS-QU-202610-003-002', 'draft', true], ['ISS-QU-202610-003-001', 'official', false]], '[B6] 發出紀錄 stillValid：舊正式版已失效、新草稿有效');
  // B7 無效單／不存在／格式
  r = await issue('_s_owner', 'Q4');
  eq([r.status, r.json.code], [409, 'QUOTE_INVALID'], '[B7] 無效單 /issue → 409 QUOTE_INVALID');
  r = await http('_s_owner', 'GET', '/api/quotations/Q4/export');
  eq([r.status, r.json && r.json.code], [409, 'QUOTE_INVALID'], '[B7] 無效單 export → 409（與以前相同）');
  ok(!(dataNow().quoteIssuances || []).some(x => x.quoteId === 'Q4'), '[B7] 無效單沒有留下發出紀錄');
  eq((await issue('_s_owner', 'NOPE')).status, 404, '[B7] 不存在 → 404');
  eq((await issue('_s_owner', 'Q1', 'docx')).status, 400, '[B7] 不支援格式 → 400');
  eq((await http('_s_owner', 'POST', '/api/quotations/Q1/issue', { format: 'pnl-xlsx' })).status, 400, '[B7] 不可由 /issue 直接發 pnl-xlsx（內部文件只走毛利分析匯出）');
  eq((await http('_s_owner', 'POST', '/api/quotations/Q1/issue', { format: 'pdf', kind: 'official', approvedValid: true })).json.kind, 'official', '[B7] （核准有效單）前端傳 kind 無作用');
  r = await http('_s_owner', 'POST', '/api/quotations/Q5/issue', { format: 'pdf', kind: 'official', approvedValid: true, valid: true });
  eq([r.json.kind, r.json.stampToken], ['draft', undefined], '[B7] 不信任前端：未核准單即使傳 kind=official 仍是 draft');
  // B8 權限
  r = await issue('_s_other', 'Q1');
  eq(r.status, 403, '[B8] 非擁有者業務 → 403');
  r = await issue(null, 'Q1');
  ok(r.status === 401 || r.status === 403 || r.status === 302, `[B8] 未登入被擋（${r.status}）`);
  r = await issue('_s_secro', 'Q1');
  eq([r.status, r.json.kind], [200, 'official'], '[B8] 唯讀秘書可發出（與今日可下載一致；伺服器 middleware 豁免 /issue）');
  r = await issue('_s_sec', 'Q2');
  eq([r.status, r.json.kind], [200, 'draft'], '[B8] 一般秘書可發出草稿');
  r = await issue('_s_admin', 'Q1');
  eq(r.status, 200, '[B8] 管理員可發出');
  r = await http('_s_secro', 'POST', '/api/quotations/Q1/withdraw', {});
  eq(r.status, 403, '[B8] 唯讀秘書的其他寫入仍被擋（豁免只限 /issue）');
  r = await http('_s_other', 'GET', '/api/quotations/Q1/issuances');
  eq(r.status, 403, '[B8] 非相關者看不到發出紀錄');
  r = await http('_s_secro', 'GET', '/api/quotations/Q1/issuances');
  ok(r.status === 200 && r.json.issuances.length > 0, '[B8] 唯讀秘書可看發出紀錄');
  // B9 章管理人預覽不受影響
  r = await fetch(BASE + '/api/quote-approval/seal', { headers: { Cookie: jar._s_admin } });
  eq([r.status, Buffer.from(await r.arrayBuffer()).equals(PNG)], [200, true], '[B9] 章管理人（admin）預覽 /quote-approval/seal 不需憑證、行為不變');
  r = await fetch(BASE + '/api/quote-approval/seal', { headers: { Cookie: jar._s_sec } });
  eq(r.status, 200, '[B9] 秘書（章管理人）預覽照常');
  r = await fetch(BASE + '/api/quote-approval/seal', { headers: { Cookie: jar._s_owner } });
  eq(r.status, 403, '[B9] 一般業務仍不可直接取章');
  // B10 序列化相容
  const g1 = (await http('_s_owner', 'GET', '/api/quotations/Q1')).json;
  ok(g1.issuanceCount > 0, '[B10] 發出過的單多一個 issuanceCount');
  ok(!('issuances' in g1) && !('quoteIssuances' in g1), '[B10] 摘要快取不外洩到回應');
  const g6 = (await http('_s_owner', 'GET', '/api/quotations/Q6')).json;
  ok(!('issuanceCount' in g6), '[B10] 沒發出過的單完全沒有 issuanceCount 鍵（與以前相同）');
  const lstAll = (await http('_s_owner', 'GET', '/api/quotations')).json;
  const row6 = (Array.isArray(lstAll) ? lstAll : lstAll.quotations || []).find(x => x.id === 'Q6');
  ok(row6 && !('issuanceCount' in row6), '[B10] 清單裡未發出的單也沒有');
  // B11 稽核
  const logs = await getLogs();
  const iq = logs.filter(l => l.action === 'ISSUE_QUOTE');
  ok(iq.length >= 10, `[B11] 每次發出都有 ISSUE_QUOTE 稽核（${iq.length} 筆）`);
  ok(iq.some(l => /ISS-QU-202610-001-001/.test(l.detail) && /official/.test(l.detail) && new RegExp(j.shortHash).test(l.detail)), '[B11] 稽核含序號／類型／格式／短雜湊');
  ok(iq.some(l => /draft/.test(l.detail) && /草稿/.test(l.detail)), '[B11] 草稿發出有標註');
  ok(logs.some(l => l.action === 'ISSUE_STAMP' && /ISS-QU-202610-001-001/.test(l.detail)), '[B11] 蓋章憑證被使用有 ISSUE_STAMP 稽核');
  const everything = JSON.stringify(logs);
  ok(!everything.includes(j.stampToken) && !everything.includes(t3.stampToken) && !everything.includes(t2.stampToken), '[B11] 稽核中沒有任何憑證字串');
  ok(!/stampToken/i.test(everything), '[B11] 稽核中沒有 stampToken 字樣');
  // B12 驗證端點
  const rec1 = dataNow().quoteIssuances.find(x => x.issueNo === 'ISS-QU-202610-001-001');
  let v = await http('_s_admin', 'GET', '/api/admin/issuances/verify?issueNo=ISS-QU-202610-001-001&shortHash=' + j.shortHash);
  ok(v.status === 200 && v.json.found && v.json.stillValid && v.json.hashMatches === true && v.json.mismatches.length === 0 && v.json.record.kind === 'official', `[B12] verify：吻合（${v.status} ${v.text.slice(0, 200)}）`);
  ok(!('tokenJti' in v.json.record) && !('ip' in v.json.record) && !v.text.includes(rec1.tokenJti), '[B12] verify 不回傳 jti／ip');
  v = await http('_s_admin', 'GET', '/api/admin/issuances/verify?issueNo=ISS-QU-202610-003-001');
  ok(v.json.found && !v.json.stillValid && v.json.mismatches.includes('CONTENT_CHANGED') && v.json.mismatches.includes('OFFICIAL_BUT_CONTENT_CHANGED'), '[B12] verify：作廢後的舊正式版 → stillValid=false、CONTENT_CHANGED');
  v = await http('_s_admin', 'GET', '/api/admin/issuances/verify?issueNo=ISS-QU-202610-001-001&shortHash=000000000000');
  ok(v.json.hashMatches === false && v.json.mismatches.includes('SHORT_HASH_DIFFERENT'), '[B12] verify：文件上的短雜湊不符被抓到');
  v = await http('_s_admin', 'GET', '/api/admin/issuances/verify?issueNo=ISS-QU-202610-002-001');
  ok(v.json.found && v.json.record.kind === 'draft' && v.json.stillValid, '[B12] verify：草稿紀錄顯示 kind=draft');
  eq((await http('_s_admin', 'GET', '/api/admin/issuances/verify?issueNo=ISS-NOPE-000')).status, 404, '[B12] verify：查無序號 404');
  eq((await http('_s_admin', 'GET', '/api/admin/issuances/verify')).status, 400, '[B12] verify：缺 issueNo 400');
  ok([401, 403].includes((await http('_s_owner', 'GET', '/api/admin/issuances/verify?issueNo=ISS-QU-202610-001-001')).status), '[B12] verify：非管理員被擋');
  // B13 競爭：並行 12 次，序號不重複且連號
  const par = await Promise.all(Array.from({ length: 12 }, () => issue('_s_owner', 'Q5')));
  const nos = par.map(x => x.json.issueNo).sort();
  ok(par.every(x => x.status === 200) && new Set(nos).size === 12, `[B13] 12 個並行發出：全部成功且序號不重複（${nos.join(',')}）`);
  const l5 = dataNow().quoteIssuances.filter(x => x.quoteId === 'Q5').map(x => parseInt(x.issueNo.slice(-3), 10)).sort((a, b) => a - b);
  ok(l5.every((v2, i) => v2 === i + 1), '[B13] Q5 的序號 1..N 連號、無缺號');
  // B14 持久化：重啟後紀錄仍在、序號延續
  const before = dataNow().quoteIssuances.length;
  await stopServer(); await startServer(); jar = {}; await login('_s_owner'); await login('_s_admin');
  const afterR = await issue('_s_owner', 'Q2');
  ok(dataNow().quoteIssuances.length === before + 1 && /ISS-QU-202610-002-00[2-9]/.test(afterR.json.issueNo), '[B14] 重啟後紀錄仍在，序號延續');
  // B15 PNL 與 Excel 的紀錄（細節在 C）
  r = await http('_s_admin', 'GET', '/api/quotations/Q1/export-pnl');
  ok(r.status === 200 && /^ISS-QU-202610-001-\d{3}$/.test(r.headers.get('x-issue-no') || ''), `[B15] 毛利分析匯出：X-Issue-No（${r.status}）`);
  const pr = dataNow().quoteIssuances.find(x => x.issueNo === r.headers.get('x-issue-no'));
  ok(pr && pr.kind === 'internal' && pr.format === 'pnl-xlsx' && pr.tokenJti === null, '[B15] 毛利分析的紀錄 kind=internal、format=pnl-xlsx、無憑證');
}

// ═════════════ C：Excel ═════════════
async function readZip(buf) { const z = await JSZip.loadAsync(buf); const o = {}; for (const n of Object.keys(z.files)) if (!z.files[n].dir) o[n] = await z.file(n).async('nodebuffer'); return o; }
async function secC() {
  console.log('[C] Excel 發出標記');
  await freshServer();
  const TPL = path.join(ROOT, 'templates', 'quotation_template.xlsx');
  // 舊版模組（釘死 commit 397abc5＝發出機制 518f461 的父提交；原本用 origin/main，合併後變成自己比自己）。可用 QISS_BASE_REF 覆寫
  const oldSrc = execFileSync('git', ['show', (process.env.QISS_BASE_REF || '397abc5fbb4299976cb8889be3054f8dfa5c741f') + ':lib/quoteExcel.js'], { cwd: ROOT, encoding: 'utf8' });
  const oldFile = path.join(ROOT, 'lib', 'quoteExcel.__old.js'); fs.writeFileSync(oldFile, oldSrc); tmpFiles.push(oldFile);
  const OLD = require(oldFile), NEW = require(path.join(ROOT, 'lib', 'quoteExcel.js'));
  const q = approve(mkQuote('QX', 'QU-X-001', { contactName: '王小明', company: '甲公司' }));
  const base = { issueDate: '2026-10-14', issuer: { name: '業務甲', phone: '02-1234', ext: '12', mobile: '0900' }, approved: true, seal: null, report: {} };
  const oldBuf = await OLD.buildQuoteWorkbook(q, TPL, Object.assign({}, base, { report: {} }));
  const newBuf = await NEW.buildQuoteWorkbook(q, TPL, Object.assign({}, base, { report: {} }));
  const zo = await readZip(oldBuf), zn = await readZip(newBuf);
  eq(Object.keys(zn).sort(), Object.keys(zo).sort(), '[C1] 未帶 issuance：部件清單與舊版相同');
  ok(Object.keys(zo).every(n => zo[n].equals(zn[n])), '[C1] 未帶 issuance：每個部件內容與舊版位元級相同');
  const issO = { issueNo: 'ISS-QU-X-001-007', shortHash: 'abcdef012345', kind: 'official', watermark: false };
  const issD = { issueNo: 'ISS-QU-X-001-008', shortHash: '0123456789ab', kind: 'draft', watermark: false };
  const issDW = Object.assign({}, issD, { watermark: true });
  const bo = await readZip(await NEW.buildQuoteWorkbook(q, TPL, Object.assign({}, base, { issuance: issO, report: {} })));
  const bd = await readZip(await NEW.buildQuoteWorkbook(q, TPL, Object.assign({}, base, { approved: false, issuance: issD, report: {} })));
  const bdw = await readZip(await NEW.buildQuoteWorkbook(q, TPL, Object.assign({}, base, { approved: false, issuance: issDW, report: {} })));
  const sheet = (z) => z['xl/worksheets/sheet1.xml'].toString('utf8');
  ok(sheet(bo).includes('發出序號 ISS-QU-X-001-007｜版本雜湊 abcdef012345（正式）'), '[C2] official：G5 含發出序號、版本雜湊與（正式）');
  ok(!sheet(bo).includes('草稿'), '[C2] official：沒有草稿標記');
  ok(sheet(bd).includes('發出序號 ISS-QU-X-001-008｜版本雜湊 0123456789ab（草稿）') && !sheet(bd).includes('草稿／未經核准'), '[C3] draft（浮水印關）：G5 有序號、短雜湊與（草稿），沒有 G4 大標記');
  ok(sheet(bdw).includes('發出序號 ISS-QU-X-001-008｜版本雜湊 0123456789ab（草稿）') && sheet(bdw).includes('草稿／未經核准'), '[C3] draft（浮水印開）：G5＋G4「草稿／未經核准」');
  ok(!sheet(bo).includes('草稿'), '[C3] official：永遠沒有草稿字樣');
  const stripCells = (x) => x.replace(/<c r="G4"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/, '').replace(/<c r="G5"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/, '');
  const names = Object.keys(zo).filter(n => n !== 'xl/worksheets/sheet1.xml' && n !== 'xl/styles.xml');
  ok(names.every(n => zo[n].equals(bo[n]) && zo[n].equals(bd[n])), '[C4] 其餘部件（圖片、sharedStrings、workbook、列印設定…）與舊版位元級相同');
  ok(stripCells(sheet(bo)) === stripCells(sheet(zo)), '[C4] official：工作表除 G4／G5 兩格外與舊版逐字相同');
  const zoSheetNoG = stripCells(sheet(zo));
  ok(stripCells(sheet(bd)) === zoSheetNoG && stripCells(sheet(bdw)) === zoSheetNoG, '[C4] draft（浮水印開／關）：工作表除 G4／G5 兩格外與舊版逐字相同');
  const stylesO = zo['xl/styles.xml'].toString('utf8'), stylesN = bd['xl/styles.xml'].toString('utf8');
  const nOld = +/<cellXfs count="(\d+)"/.exec(stylesO)[1], nNew = +/<cellXfs count="(\d+)"/.exec(stylesN)[1];
  eq(nNew, nOld + 2, '[C5] styles.xml 只追加 2 個樣式（發出標記用）');
  ok(stylesN.startsWith(stylesO.slice(0, 200)) && stylesN.replace(/<fonts count="\d+"/, '').length > stylesO.length, '[C5] 樣式檔只是追加');
  ok(/G4"[^>]*s="\d+"/.test(sheet(bd)) && /G5"[^>]*s="\d+"/.test(sheet(bd)), '[C5] G4／G5 套用新樣式');
  // 真實伺服器端
  let r = await http('_s_owner', 'GET', '/api/quotations/Q1/export');
  ok(r.status === 200 && r.headers.get('x-quote-approval') === 'approved' && r.headers.get('x-issue-kind') === 'official' && /^ISS-QU-202610-001-\d{3}$/.test(r.headers.get('x-issue-no') || ''), `[C6] 核准單 export：200、X-Quote-Approval=approved、X-Issue-Kind=official、X-Issue-No（${r.status}）`);
  const no1 = r.headers.get('x-issue-no');
  let z = await readZip(r.buf);
  ok(sheet(z).includes('發出序號 ' + no1) && !sheet(z).includes('草稿'), '[C6] 實際下載檔含該序號、無草稿標記');
  const rec = dataNow().quoteIssuances.find(x => x.issueNo === no1);
  ok(rec && rec.kind === 'official' && rec.format === 'xlsx' && rec.issuedBy === '_s_owner' && sheet(z).includes(crypto.createHash('sha256').update(no1 + ':' + rec.contentHash).digest('hex').slice(0, 12)), '[C6] 紀錄與檔案內的短雜湊一致');
  eq(r.headers.get('x-quote-seal'), 'none', '[C6] X-Quote-Seal 仍為 none（Excel 不蓋章）');
  r = await http('_s_owner', 'GET', '/api/quotations/Q2/export');
  z = await readZip(r.buf);
  ok(r.status === 200 && r.headers.get('x-quote-approval') === 'unapproved' && r.headers.get('x-issue-kind') === 'draft' && !sheet(z).includes('草稿／未經核准') && sheet(z).includes('發出序號 ' + r.headers.get('x-issue-no')) && sheet(z).includes('（草稿）'), '[C7] 未核准單 export：仍可下載、標頭 unapproved、檔內有草稿標記與序號');
  const logs = await getLogs();
  ok(logs.some(l => l.action === 'EXPORT_QUOTATION' && l.detail.includes('發出序號 ' + no1)), '[C8] EXPORT_QUOTATION 稽核保留且帶 issueNo');
  ok(logs.some(l => l.action === 'ISSUE_QUOTE' && l.detail.includes(no1)), '[C8] 同時有 ISSUE_QUOTE');
  // export-log 不重複計數
  const nIss = dataNow().quoteIssuances.length;
  const iss = (await issue('_s_owner', 'Q1')).json;
  r = await http('_s_owner', 'POST', '/api/quotations/Q1/export-log', { format: 'pdf', issueNo: iss.issueNo });
  ok(r.status === 200 && dataNow().quoteIssuances.length === nIss + 1, '[C9] export-log 不新增發出紀錄（沒有重複計數）');
  const lg2 = await getLogs();
  ok(lg2.some(l => l.action === 'EXPORT_QUOTATION_PDF' && l.detail.includes('發出序號 ' + iss.issueNo)), '[C9] export-log 稽核帶發出序號');
  r = await http('_s_owner', 'POST', '/api/quotations/Q1/export-log', { format: 'pdf', issueNo: 'ISS-FAKE-1' });
  const lg3 = await getLogs();
  ok(r.status === 200 && lg3.some(l => l.action === 'EXPORT_QUOTATION_PDF' && l.detail.includes('發出序號不符')), '[C9] export-log 帶別人的／偽造序號 → 不採信，只註記');
  r = await http('_s_owner', 'POST', '/api/quotations/Q1/export-log', { format: 'pdf' });
  ok(r.status === 200 && (await getLogs()).some(l => l.action === 'EXPORT_QUOTATION_PDF' && l.detail.includes('未帶發出序號')), '[C9] 舊版前端（不帶序號）仍可回報，註記未帶序號');
  // 浮水印開關開啟
  await freshServer(undefined, { QUOTE_DRAFT_WATERMARK: '1' });
  r = await http('_s_owner', 'GET', '/api/quotations/Q2/export');
  z = await readZip(r.buf);
  ok(sheet(z).includes('草稿／未經核准') && sheet(z).includes('（草稿）'), '[C7b] 浮水印開：未核准 Excel 有 G4「草稿／未經核准」且 G5 有（草稿）');
  r = await http('_s_owner', 'GET', '/api/quotations/Q1/export');
  z = await readZip(r.buf);
  ok(!sheet(z).includes('草稿') && sheet(z).includes('（正式）'), '[C7b] 浮水印開：核准有效的 Excel 沒有草稿標記、G5 有（正式）');
  { const i2 = await issue('_s_owner', 'Q2'); eq([i2.json.kind, i2.json.watermark, i2.json.stampToken], ['draft', true, undefined], '[C7b] 浮水印開：/issue 回 watermark:true，草稿仍是 draft、仍無憑證'); const i3 = await issue('_s_owner', 'Q1'); eq([i3.json.kind, i3.json.watermark], ['official', true], '[C7b] 浮水印開：official 的回應也帶開關狀態'); }
  { const vr = await http('_s_admin', 'GET', '/api/admin/version'); eq(vr.json && vr.json.quoteIssuance, { draftIssueAllowed: true, draftWatermark: true }, '[C7b] /api/admin/version：浮水印開'); }
  // 禁止草稿模式
  await freshServer(undefined, { QUOTE_ALLOW_UNAPPROVED_ISSUE: '0' });
  const n0 = (dataNow().quoteIssuances || []).length;
  { const vr = await http('_s_admin', 'GET', '/api/admin/version'); eq(vr.json && vr.json.quoteIssuance, { draftIssueAllowed: false, draftWatermark: false }, '[C10] /api/admin/version：禁止草稿發出、浮水印關'); }
  r = await http('_s_owner', 'GET', '/api/quotations/Q2/export');
  eq([r.status, r.json && r.json.code], [409, 'NOT_APPROVED'], '[C10] ALLOW=false：未核准 export → 409 NOT_APPROVED');
  r = await issue('_s_owner', 'Q2');
  eq([r.status, r.json.code], [409, 'NOT_APPROVED'], '[C10] ALLOW=false：未核准 /issue → 409 NOT_APPROVED');
  ok((dataNow().quoteIssuances || []).length === n0, '[C10] ALLOW=false：被拒絕時不留發出紀錄');
  r = await http('_s_owner', 'GET', '/api/quotations/Q1/export');
  eq([r.status, r.headers.get('x-issue-kind')], [200, 'official'], '[C10] ALLOW=false：核准有效單照常匯出');
  r = await issue('_s_owner', 'Q1');
  ok(r.status === 200 && r.json.kind === 'official', '[C10] ALLOW=false：核准有效單 /issue 照常');
  await freshServer();
}

// ═════════════ D：前端 PDF 流程（vm）═════════════
function makeDom(state) {
  const canvases = [];
  const mkCanvas = (w, h) => {
    const c = { id: canvases.length, width: w || 0, height: h || 0, texts: [], fonts: [], ops: [], args: [], rotated: 0 };
    c.getContext = () => ({
      fillStyle: '', font: '', textAlign: '', textBaseline: '', globalCompositeOperation: '',
      save() {}, restore() {}, translate() {}, rotate() { c.rotated++; }, setTransform() {}, fillRect() { c.ops.push('fillRect'); },
      drawImage(...a) { c.ops.push('drawImage'); c.args.push(a); }, fillText(t) { c.texts.push(String(t)); },
    });
    c.toDataURL = () => 'data:image/png;base64,CANVAS' + c.id;
    canvases.push(c);
    return c;
  };
  const paper = {
    style: {}, children: [],
    querySelectorAll(sel) { if (sel === '.qpv-sealnote') return state.sealnotes; return []; },
    querySelector(sel) {
      if (sel === '.qpv-seal' && /class="qpv-seal"/.test(state.html || '')) return { complete: true, naturalWidth: 8, naturalHeight: 8, style: {}, getBoundingClientRect: () => ({ left: 700, top: state.sealTop || 800, width: 84, height: 84 }) };
      return null;
    },
    getBoundingClientRect: () => ({ top: 0, left: 0, height: state.paperH || 1000, width: 900, bottom: state.paperH || 1000 }),
  };
  const host = { style: {}, setAttribute() {}, set innerHTML(v) { state.html = v; state.htmlAtBuildSealSrc = state.getSealSrc && state.getSealSrc(); }, get firstElementChild() { return paper; }, remove() { state.hostRemoved = true; } };
  const document = {
    createElement(tag) { if (tag === 'canvas') return mkCanvas(); if (tag === 'div') return host; return { style: {}, setAttribute() {}, appendChild() {} }; },
    getElementById: () => null, body: { appendChild() {} }, head: { appendChild() {} },
  };
  return { document, canvases, mkCanvas, paper };
}
function loadClient(opts) {
  opts = opts || {};
  const state = { html: '', calls: [], toasts: [], saves: [], addImages: [], pages: 0, sealnotes: [{ textContent: 'note' }], docCount: 0 };
  const dom = makeDom(state);
  const pdfSrc = fs.readFileSync(opts.pdfFile || path.join(ROOT, '_client', 'quote-pdf.js'), 'utf8');
  const prevSrc = fs.readFileSync(path.join(ROOT, '_client', 'quote-preview.js'), 'utf8');
  const sandbox = {
    console, API: '/api', Date, Math, JSON, Promise, setTimeout, clearTimeout, encodeURIComponent, String, Number, Array, Object, parseFloat, RegExp, Error, isNaN, Intl, URL,
    document: dom.document,
    window: {},
    getComputedStyle: () => ({ paddingBottom: '42px' }),
    escapeHtml: (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])),
    showToast: (m) => state.toasts.push(String(m)),
    _qFetchQuote: async () => opts.quote, _qFind: () => opts.quote, _qConfirmUnapproved: async () => true,
    FileReader: class { readAsDataURL() { this.result = 'data:image/png;base64,SEALBYTES'; setTimeout(() => this.onload && this.onload(), 0); } },
    fetch: async (url, init) => {
      const u = String(url), m = (init && init.method) || 'GET';
      state.calls.push(m + ' ' + u + (init && init.body ? ' ' + init.body : ''));
      if (/\/issue$/.test(u) && m === 'POST') {
        if (opts.issueThrows) throw new Error('net down');
        return { ok: opts.issueStatus === undefined, status: opts.issueStatus || 200, json: async () => opts.issueStatus ? { error: opts.issueError || '無權限', code: 'X' } : opts.issueResp };
      }
      if (/\/seal\?stampToken=/.test(u)) return { ok: !opts.sealFails, status: opts.sealFails ? 403 : 200, blob: async () => ({}) };
      if (/export-log$/.test(u)) return { ok: true, status: 200, json: async () => ({ success: true }) };
      if (/issue-info$/.test(u)) return { ok: true, status: 200, json: async () => ({ issueDate: '2026-10-14', issuer: {}, remarks: ['r1'] }) };
      return { ok: false, status: 404, json: async () => ({}) };
    },
  };
  sandbox.window.html2canvas = async (el, o) => { state.h2c = (state.h2c || 0) + 1; state.h2cOpts = o; const c = dom.mkCanvas(1800, 2000); state.mainCanvas = c; return c; };
  sandbox.window.jspdf = { jsPDF: class {
    constructor() { state.docCount++; state.props = null; }
    addPage() { state.pages++; }
    addImage(url, fmt, x, y, w, h) { state.addImages.push({ url, x, y, w, h }); }
    setProperties(p) { state.props = p; }
    save(n) { state.saves.push(n); }
  } };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(prevSrc, ctx);
  vm.runInContext(pdfSrc, ctx);
  vm.runInContext('state_get_seal_src = function () { return _qpvSealSrc; };', ctx);
  state.getSealSrc = () => ctx.state_get_seal_src();
  return { ctx, state, dom };
}
async function secD() {
  console.log('[D] 前端 PDF 流程（vm 假 DOM）');
  const q = { id: 'Q1', quoteNo: 'QU-202610-001', company: '甲公司', owner: 'u', approval: { state: 'approved', valid: true }, perm: { canSeePrice: true },
    items: [{ lid: 'a', desc: '品項', qty: 1, unit: '式', unitPrice: 100 }], quoteDate: '2026-10-14', discountType: 'none', products: [] };
  const issOfficial = { issueNo: 'ISS-QU-202610-001-005', kind: 'official', shortHash: 'abc123def456', issuedAt: nowIso, stampToken: 'TOKEN.SIG', watermark: false };
  const issDraft = { issueNo: 'ISS-QU-202610-002-001', kind: 'draft', shortHash: '0a1b2c3d4e5f', issuedAt: nowIso, watermark: false };
  const issDraftW = Object.assign({}, issDraft, { watermark: true });
  const overlays = (S) => S.addImages.filter(a => a.w === 210 && a.h === 297).map(a => S.dom.canvases[parseInt(/CANVAS(\d+)/.exec(a.url)[1], 10)]);
  const run = async (opts) => { const C = loadClient(opts); C.state.dom = C.dom; await C.ctx.exportQuotePdf('Q1', 'QU-202610-001'); return C; };

  // D1 official
  let C = await run({ quote: q, issueResp: issOfficial }); if (process.env.QISS_DEBUG) console.log(C.state.toasts, C.state.calls);
  let S = C.state;
  ok(/POST \/api\/quotations\/Q1\/issue \{"format":"pdf"\}/.test(S.calls[0]), '[D1] 第一個請求就是 POST /issue（先取序號）');
  ok(/GET \/api\/quotations\/Q1\/seal\?stampToken=TOKEN\.SIG$/.test(S.calls.find(c => /seal/.test(c)) || ''), '[D1] 用 stampToken 取章');
  ok(S.calls.indexOf(S.calls.find(c => /issue$|\/issue /.test(c))) < S.calls.findIndex(c => /seal/.test(c)), '[D1] 取序號在取章之前');
  ok(/<img class="qpv-seal" src="data:image\/png;base64,SEALBYTES"/.test(S.html), '[D1] 預覽 HTML 的章圖用換到的 data URL（不再由 img 直接打 /seal）');
  ok(!/\/seal/.test(S.html), '[D1] 預覽 HTML 沒有直接指向 /seal 的 URL');
  eq(S.saves.length, 1, '[D1] 有存檔');
  ok(/蓋報價專用章/.test(S.toasts[S.toasts.length - 1]) && S.toasts[S.toasts.length - 1].includes('ISS-QU-202610-001-005'), '[D1] 完成提示含「已蓋」與序號');
  let ovs = overlays({ addImages: S.addImages, dom: C.dom });
  ok(ovs.length >= 1 && ovs.every(o => o.texts.includes('發出序號 ISS-QU-202610-001-005｜版本雜湊 abc123def456（正式）')), '[D2] 每一頁頁尾都有「發出序號 …｜版本雜湊 …」');
  ok(ovs.every(o => !o.texts.includes('草稿・未經核准') && o.rotated === 0), '[D2] official 沒有浮水印');
  ok(ovs.every((o, i) => o.texts.includes((i + 1) + ' / ' + ovs.length)), '[D2] 頁碼 i / n');
  eq(S.addImages.length, ovs.length * 2, '[D2] 每頁 = 內容圖 + 一張透明圖層（內容圖不被改）');
  ok(S.mainCanvas.texts.length === 0 && S.mainCanvas.ops.filter(x => x === 'drawImage').length >= 1, '[D3] 紙張畫布本身沒有被畫頁尾／浮水印（只疊章圖，與以前相同）');
  ok(/export-log/.test(S.calls[S.calls.length - 1]) && /"issueNo":"ISS-QU-202610-001-005"/.test(S.calls[S.calls.length - 1]), '[D4] export-log 作為次要回報，帶 issueNo');
  ok(S.props && /ISS-QU-202610-001-005/.test(S.props.subject || ''), '[D4] PDF 文件屬性 subject 記序號');
  // D5 草稿
  const qUn = Object.assign({}, q, { approval: { state: 'none', valid: false } });
  C = await run({ quote: qUn, issueResp: issDraft }); S = C.state;
  ok(!S.calls.some(c => /seal/.test(c)), '[D5] draft：完全不去取章');
  ok(!/qpv-seal"/.test(S.html), '[D5] draft：預覽 HTML 沒有章圖');
  ovs = overlays({ addImages: S.addImages, dom: C.dom });
  ok(ovs.length >= 1 && ovs.every(o => !o.texts.includes('草稿・未經核准') && o.rotated === 0), '[D5] draft（浮水印開關關）：沒有浮水印');
  ok(ovs.every(o => o.texts.includes('發出序號 ISS-QU-202610-002-001｜版本雜湊 0a1b2c3d4e5f（草稿）')), '[D5] draft（浮水印關）：頁尾仍有序號、雜湊與（草稿）');
  ok(/草稿/.test(S.saves[0]) && /草稿・未經核准/.test(S.toasts[S.toasts.length - 1]), '[D5] draft：檔名與提示標示草稿');
  { const CW = await run({ quote: qUn, issueResp: issDraftW }); const ow = overlays({ addImages: CW.state.addImages, dom: CW.dom });
    ok(ow.length >= 1 && ow.every(o => o.texts.includes('草稿・未經核准') && o.rotated >= 1), '[D5w] draft（浮水印開）：每頁斜向「草稿・未經核准」浮水印');
    ok(ow.every(o => o.texts.includes('發出序號 ISS-QU-202610-002-001｜版本雜湊 0a1b2c3d4e5f（草稿）')) && !CW.state.calls.some(c => /seal/.test(c)) && CW.state.mainCanvas.ops.filter(x => x === 'drawImage').length === 0, '[D5w] draft（浮水印開）：頁尾（草稿）、仍不取章、不蓋章');
    const CO = await run({ quote: q, issueResp: Object.assign({}, issOfficial, { watermark: true }) }); const oo = overlays({ addImages: CO.state.addImages, dom: CO.dom });
    ok(oo.every(o => o.rotated === 0 && !o.texts.includes('草稿・未經核准')), '[D5w] official：開關開也沒有浮水印'); }
  eq(S.mainCanvas.ops.filter(x => x === 'drawImage').length, 0, '[D5] draft：紙張畫布沒有疊任何章');
  // D11 蓋章位置（壓線）：章畫在『預覽版面位置 + STAMP_LINE_OFFSET』
  { const cst = (n) => vm.runInContext(n, C.ctx); const OY = cst('STAMP_LINE_OFFSET_Y'), OX = cst('STAMP_LINE_OFFSET_X');
    const sd = S0 => S0.mainCanvas.args[0];
    C = await run({ quote: q, issueResp: issOfficial }); const a0 = sd(C.state);
    eq([OX, OY], [0, 40], '[D11] 常數 STAMP_LINE_OFFSET_X/Y＝0/40');
    ok(a0 && a0[1] === (700 + OX) * 2 && a0[2] === (800 + OY) * 2 && OY > 0, `[D11] 章座標＝(預覽位置＋偏移)×倍率（${a0 && a0.slice(1, 3)}）`);
    ok(a0[2] + a0[4] <= 2000 && a0[1] + a0[3] <= 1800, '[D11] 章仍在紙張畫布內');
    // 多頁：章只畫一次，且落在「有簽名區」那一頁的範圍內
    const C4 = loadClient({ quote: q, issueResp: issOfficial }); C4.state.paperH = 3000; C4.state.sealTop = 2850;
    const b4 = await C4.ctx._qpdfBuildDoc(q, { issueDate: '2026-10-14', issuer: {}, remarks: ['r1'] }, { sealSrc: 'data:image/png;base64,SEAL', issuance: { issueNo: 'X', shortHash: 'y', kind: 'official' } });
    const drawn = C4.state.mainCanvas.args; const scale = C4.state.mainCanvas.width / 900;
    const yTop = drawn[0][2] / scale, hits = b4.pages.filter(p => yTop + 42 >= p.y0 && yTop + 42 < p.y1);
    ok(b4.pageCount >= 2 && drawn.length === 1 && hits.length === 1 && hits[0] === b4.pages[b4.pages.length - 1], `[D11] 多頁（${b4.pageCount} 頁）：章只畫一次，且中心落在最後一頁（簽名區所在頁）`);
    // 舊位置對照：偏移 0 ＝ 預覽原位
    const C5 = loadClient({ quote: q, issueResp: issOfficial }); vm.runInContext('_qpdfStampOffset.y = 0', C5.ctx);
    await C5.ctx.exportQuotePdf('Q1', 'N'); eq(sd(C5.state)[2], 800 * 2, '[D11] 偏移 0＝舊位置（供截圖對照）'); }
  // D6 伺服器說 draft 但前端以為已核准 → 以伺服器為準，不蓋章
  C = await run({ quote: q, issueResp: issDraft }); S = C.state;
  ok(!S.calls.some(c => /seal/.test(c)) && !/qpv-seal"/.test(S.html) && S.mainCanvas.ops.filter(x => x === 'drawImage').length === 0, '[D6] 前端以為已核准、伺服器說 draft → 不取章、不蓋章（不信任前端狀態）');
  // D7 issue 失敗 → 中止
  for (const [label, o] of [['HTTP 409', { issueStatus: 409, issueError: '這張報價單尚未核准，目前不允許發出文件' }], ['HTTP 403', { issueStatus: 403, issueError: '無權限' }], ['網路中斷', { issueThrows: true }]]) {
    C = await run(Object.assign({ quote: q }, o)); S = C.state;
    ok(S.saves.length === 0 && !S.h2c && S.docCount === 0, `[D7] ${label}：中止下載（沒有產圖、沒有存檔）`);
    ok(!S.calls.some(c => /seal|export-log|issue-info/.test(c)), `[D7] ${label}：沒有取章、沒有回報下載`);
    ok(/PDF 未產生|無權限|尚未核准/.test(S.toasts[S.toasts.length - 1]) && S.toasts[S.toasts.length - 1].startsWith('⚠'), `[D7] ${label}：清楚的錯誤訊息（${S.toasts[S.toasts.length - 1]}）`);
  }
  C = await run({ quote: q, issueResp: {} }); S = C.state;
  ok(S.saves.length === 0, '[D7] 伺服器回應沒有 issueNo → 中止');
  // D8 official 但取章失敗 → 照樣出檔（有序號）、不宣稱蓋章
  C = await run({ quote: q, issueResp: issOfficial, sealFails: true }); S = C.state;
  ok(S.saves.length === 1 && /沒有蓋上/.test(S.toasts[S.toasts.length - 1]) && S.mainCanvas.ops.filter(x => x === 'drawImage').length === 0, '[D8] 取章失敗：PDF 仍帶序號出檔但明講沒蓋章');
  ok(!/報價專用章於下載 PDF 時蓋上/.test(S.sealnotes.map(x => x.textContent).join('')), '[D8] 預覽上的「章於下載 PDF 時蓋上」提示不會進 PDF');
  // D9 預覽預設無章
  const C2 = loadClient({ quote: q });
  const html = C2.ctx.buildQuotePreviewHtml(q, { issueDate: '2026-10-14', issuer: {}, remarks: ['r1'] });
  ok(!/class="qpv-seal"/.test(html) && /qpv-sealnote/.test(html), '[D9] 網頁預覽（不是發出流程）不載章，只有文字提示');
  ok(!/\/seal/.test(html), '[D9] 預覽 HTML 沒有指向 /seal 的 URL');
  // D10 沒有 opts 的 _qpdfBuildDoc（相容舊呼叫）＝沒有圖層
  const C3 = loadClient({ quote: q });
  const built = await C3.ctx._qpdfBuildDoc(q, { issueDate: '2026-10-14', issuer: {}, remarks: ['r1'] });
  eq(C3.state.addImages.length, built.pageCount, '[D10] 不帶 opts：與以前相同，每頁只有內容圖（無頁尾、無浮水印）');
}

// ═════════════ E：靜態接線 ═════════════
function secE() {
  console.log('[E] 靜態接線');
  const rd = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
  const srv = rd('server.js'), sw = rd('_client/sw.js'), appr = rd('_client/quote-approval.js'), prev = rd('_client/quote-preview.js'), routes = rd('lib/quoteRoutes.js'), pdf = rd('_client/quote-pdf.js');
  ok(/export-log\|issue/.test(srv), '[E1] server.js：唯讀帳號 middleware 豁免 /issue');
  const mv = /SW_VERSION = 'itts-crm-v(\d+)'/.exec(sw);
  ok(mv && +mv[1] >= 7, `[E2] sw.js SW_VERSION 至少 v7（${mv && mv[0]}）`);
  ok(/buildIssuancesSection\(s\)/.test(appr) && /發出紀錄/.test(appr) && /已失效（內容已變更）/.test(appr), '[E3] 簽核面板有「發出紀錄」區塊與「已失效（內容已變更）」');
  const blk = appr.slice(appr.indexOf('function buildIssuancesSection'), appr.indexOf('function buildIssuancesSection') + 1400);
  const exprs = [...blk.matchAll(/\$\{([^}]+)\}/g)].map(m => m[1]).filter(x => !/^e\(|^rows$|^list\.length|^\(r\.stillValid/.test(x.trim()) && !/^r\.stillValid/.test(x.trim()));
  ok(exprs.length === 0, `[E4] 發出紀錄區塊的插值全部經 e() 跳脫（未跳脫：${exprs.join(' | ')}）`);
  ok(/<details class="qap-sec qap-ref" id="qapIssuances">/.test(blk), '[E5] 可摺疊（details，預設收合）且沿用既有 class');
  ok(!/src="\$\{API\}\/quotations\/\$\{encodeURIComponent\(q\.id\)\}\/seal/.test(prev), '[E6] quote-preview.js 不再直接以 <img src> 載入 /seal');
  ok(/await _qpdfIssue\(id\)/.test(pdf) && pdf.indexOf('iss = await _qpdfIssue(id)') < pdf.indexOf('await _qpdfBuildDoc(q, info'), '[E7] quote-pdf.js 先取序號再建 PDF');
  ok(/app\.get\('\/api\/quotations\/:id\/seal'/.test(routes) && /STAMP_TOKEN_REQUIRED/.test(routes), '[E8] 路由有 STAMP_TOKEN_REQUIRED');
  ok(/lib\/quoteIssuance\.js|quoteIssuance/.test(routes) && fs.existsSync(path.join(ROOT, 'lib', 'quoteIssuance.js')), '[E9] 新模組存在並被引用（漏 add 會讓部署 MODULE_NOT_FOUND）');
  ok(!/(JWT_SECRET|SESSION_SECRET)[^\n]*stampToken|console\.\w+\([^)]*stampToken/.test(routes + rd('lib/quoteIssuance.js')), '[E10] 程式裡沒有把憑證寫進 log 的呼叫');
}

// ═════════════ M：變異檢查 ═════════════
function copyMut(rels, edits) {
  // 把 lib/quoteIssuance.js 或 lib/quoteRoutes.js 等複製成 __mut_ 版本；routes 版指向 issuance 變異版；server 版指向 routes 變異版
  const tag = '__mut_' + crypto.randomBytes(3).toString('hex');
  const out = {};
  const read = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
  let issSrc = read('lib/quoteIssuance.js'), routesSrc = read('lib/quoteRoutes.js'), excelSrc = read('lib/quoteExcel.js'), srvSrc = read('server.js');
  const apply = (src, from, to, name) => { if (!src.includes(from)) throw new Error('變異樣式找不到：' + name); return src.replace(from, () => to); };
  for (const e of edits) {
    if (e.file === 'issuance') issSrc = apply(issSrc, e.from, e.to, e.name);
    else if (e.file === 'routes') routesSrc = apply(routesSrc, e.from, e.to, e.name);
    else if (e.file === 'excel') excelSrc = apply(excelSrc, e.from, e.to, e.name);
  }
  const f = (n) => path.join(ROOT, 'lib', n + tag + '.js');
  fs.writeFileSync(f('quoteIssuance'), issSrc); fs.writeFileSync(f('quoteExcel'), excelSrc);
  routesSrc = routesSrc.replace("require('./quoteIssuance')", () => `require('./quoteIssuance${tag}')`).replace("require('./quoteExcel')", () => `require('./quoteExcel${tag}')`);
  fs.writeFileSync(f('quoteRoutes'), routesSrc);
  srvSrc = srvSrc.replace("require('./lib/quoteRoutes')", () => `require('./lib/quoteRoutes${tag}')`);
  // server.js 也直接 require quoteExcel（buildQuoteWorkbook）：變異 excel 時一併替換
  if (edits.some(e => e.file === 'excel')) srvSrc = srvSrc.replace(/require\('\.\/lib\/quoteExcel'\)/g, () => `require('./lib/quoteExcel${tag}')`);
  const sf = path.join(ROOT, 'server' + tag + '.js'); fs.writeFileSync(sf, srvSrc);
  [f('quoteIssuance'), f('quoteExcel'), f('quoteRoutes'), sf].forEach(x => tmpFiles.push(x));
  out.server = sf; out.files = [f('quoteIssuance'), f('quoteExcel'), f('quoteRoutes'), sf];
  return out;
}
async function mutRun(name, edits, probe, env) {
  const m = copyMut(null, edits);
  try {
    muting = true; mutFail = 0;
    try { await freshServer(m.server, env); await probe(); } catch (e) { mutFail++; if (process.env.QISS_MUTDBG) console.log('    exception:', e.message); }
    muting = false;
    ok(mutFail > 0, `變異 ${name} → 對應檢查必須失敗，實際失敗 ${mutFail} 項`);
  } finally { muting = false; await stopServer(); m.files.forEach(x => { try { fs.unlinkSync(x); } catch { /* 已刪 */ } }); }
}
async function secM() {
  console.log('[M] 變異檢查');
  // M1 憑證單次使用
  await mutRun('憑證可重複使用', [{ file: 'issuance', name: 'used', from: "if (data.quoteStampUses.some(u => u.jti === pl.j)) return bad('used');", to: '' }], async () => {
    const j = (await issue('_s_owner', 'Q1')).json;
    await seal('_s_owner', 'Q1', j.stampToken);
    const s2 = await seal('_s_owner', 'Q1', j.stampToken);
    ok(s2.status === 403, '（變異探針）第二次使用必須 403');
  });
  // M2 核准有效判定：用「核准還在但雜湊不同」的單（seed 時就不符）驗證
  await mutRunSeeded('核准有效不比對內容雜湊（雜湊不符的核准單）', [{ file: 'issuance', name: 'valid2', from: "return !!ap && ap.state === 'approved' && QA.contentHash(q) === ap.hash;", to: "return !!ap && ap.state === 'approved';" }]);
  // M3 先記錄後串流：匯出不留紀錄
  await mutRun('匯出沒有先留紀錄', [{ file: 'routes', name: 'norec', from: "const iss = recordIssuance(ctx, req, live, q, 'xlsx');\n    if (!iss.ok)", to: "const iss = { ok: true, rec: { issueNo: 'ISS-FAKE-000', kind: valid ? 'official' : 'draft' }, shortHash: '000000000000' };\n    if (!iss.ok)" }], async () => {
    const r = await http('_s_owner', 'GET', '/api/quotations/Q1/export');
    const recs = (dataNow().quoteIssuances || []).filter(x => x.quoteId === 'Q1' && x.format === 'xlsx');
    ok(r.status === 200 && recs.length === 1 && recs[0].issueNo === r.headers.get('x-issue-no'), '（變異探針）回應前資料庫必須已有該序號的紀錄');
  });
  // M4 seal 必須憑證
  await mutRun('seal 不再要求憑證', [{ file: 'routes', name: 'seal', from: "if (!tok) return fail(res, 403, 'STAMP_TOKEN_REQUIRED'", to: "if (false) return fail(res, 403, 'STAMP_TOKEN_REQUIRED'" }, { file: 'routes', name: 'seal2', from: "const used = QIS.consumeStamp(ctx.data, q, tok, ctx.me);\n    if (!used.ok)", to: "const used = { ok: true, rec: { issueNo: 'x' } };\n    if (!used.ok)" }], async () => {
    const s = await seal('_s_owner', 'Q1');
    ok(s.status === 403, '（變異探針）沒憑證必須 403');
  });
  // M5 草稿標記
  await mutRun('草稿 Excel 沒有標記', [{ file: 'excel', name: 'draftmark', from: "if (opts.issuance.kind === 'draft' && opts.issuance.watermark === true) sh.setText", to: "if (false) sh.setText" }], async () => {
    const r = await http('_s_owner', 'GET', '/api/quotations/Q2/export');
    const z = await readZip(r.buf);
    ok(z['xl/worksheets/sheet1.xml'].toString('utf8').includes('草稿／未經核准'), '（變異探針）浮水印開時草稿 Excel 必須有標記');
  }, { QUOTE_DRAFT_WATERMARK: '1' });
  // M5b 開關關閉時 Excel 仍標記
  await mutRun('浮水印關閉時 Excel 仍標草稿', [{ file: 'excel', name: 'wmoff', from: "opts.issuance.kind === 'draft' && opts.issuance.watermark === true", to: "opts.issuance.kind === 'draft'" }], async () => {
    const r = await http('_s_owner', 'GET', '/api/quotations/Q2/export');
    const z = await readZip(r.buf);
    ok(!z['xl/worksheets/sheet1.xml'].toString('utf8').includes('草稿／未經核准'), '（變異探針）開關關閉時不得有 G4 標記');
  });
  // M5c 伺服器端開關恆為開
  await mutRun('伺服器浮水印開關恆開', [{ file: 'issuance', name: 'wmon', from: "return /^(1|true|yes|on)$/i.test(String(process.env.QUOTE_DRAFT_WATERMARK || '').trim());", to: 'return true;' }], async () => {
    const r = await issue('_s_owner', 'Q2');
    ok(r.json.watermark === false, '（變異探針）預設 /issue 必須回 watermark:false');
  });
  // M6 kind 信任前端
  await mutRun('kind 信任前端', [{ file: 'routes', name: 'trustkind', from: "const iss = recordIssuance(ctx, req, live, q, format);\n    if (!iss.ok) return fail(res, iss.status, iss.code, iss.message);\n    await flushAuditWrites();", to: "if (req.body && req.body.kind === 'official') { q.approval = live.approval = Object.assign({}, live.approval || {}, { state: 'approved', hash: QA.contentHash(q) }); }\n    const iss = recordIssuance(ctx, req, live, q, format);\n    if (!iss.ok) return fail(res, iss.status, iss.code, iss.message);\n    await flushAuditWrites();" }], async () => {
    const r = await http('_s_owner', 'POST', '/api/quotations/Q5/issue', { format: 'pdf', kind: 'official' });
    ok(r.json.kind === 'draft', '（變異探針）未核准單傳 kind=official 必須仍是 draft');
  });
  // M7 前端：issue 失敗不中止
  const pdfFile = path.join(ROOT, '_client', 'quote-pdf.__mut.js');
  const pdfSrc = fs.readFileSync(path.join(ROOT, '_client', 'quote-pdf.js'), 'utf8');
  const from = "catch (e) { return showToast('⚠ ' + ((e && e.message) || '無法取得發出序號，PDF 未產生')); }";
  if (!pdfSrc.includes(from)) throw new Error('變異樣式找不到：前端中止');
  fs.writeFileSync(pdfFile, pdfSrc.replace(from, () => "catch (e) { iss = { issueNo: 'X', kind: 'official', shortHash: '0' }; }")); tmpFiles.push(pdfFile);
  try {
    muting = true; mutFail = 0;
    try {
      const C = loadClient({ quote: { id: 'Q1', quoteNo: 'N', approval: { state: 'approved', valid: true }, perm: {}, items: [], products: [] }, issueStatus: 409, pdfFile });
      await C.ctx.exportQuotePdf('Q1', 'N');
      ok(C.state.saves.length === 0, '（變異探針）issue 失敗必須不存檔');
    } catch (e) { mutFail++; if (process.env.QISS_MUTDBG) console.log('    exception:', e.message); }
    muting = false;
    ok(mutFail > 0, `變異 前端 issue 失敗仍繼續下載 → 對應檢查必須失敗，實際失敗 ${mutFail} 項`);
  } finally { muting = false; try { fs.unlinkSync(pdfFile); } catch { /* 已刪 */ } }
  // M9 前端：浮水印不看伺服器開關
  { const pf = path.join(ROOT, '_client', 'quote-pdf.__mut3.js'); const from3 = "iss.kind === 'draft' && iss.watermark === true";
    if (!pdfSrc.includes(from3)) throw new Error('變異樣式找不到：前端浮水印');
    fs.writeFileSync(pf, pdfSrc.replace(from3, () => "iss.kind === 'draft'")); tmpFiles.push(pf);
    try { muting = true; mutFail = 0;
      try { const C = loadClient({ quote: { id: 'Q1', quoteNo: 'N', approval: { state: 'none', valid: false }, perm: {}, items: [], products: [] }, issueResp: { issueNo: 'ISS-N-001', kind: 'draft', shortHash: 'a', watermark: false }, pdfFile: pf }); await C.ctx.exportQuotePdf('Q1', 'N');
        ok(C.dom.canvases.filter(c => c.width === 1240).every(c => c.rotated === 0), '（變異探針）開關關閉不得畫浮水印'); } catch (e) { mutFail++; }
      muting = false; ok(mutFail > 0, `變異 前端浮水印不看伺服器開關 → 對應檢查必須失敗，實際失敗 ${mutFail} 項`);
    } finally { muting = false; try { fs.unlinkSync(pf); } catch { /* 已刪 */ } } }
  // M10 前端：章不壓線（偏移 0）
  { const pf = path.join(ROOT, '_client', 'quote-pdf.__mut4.js'); const from4 = 'const STAMP_LINE_OFFSET_Y = 40;';
    if (!pdfSrc.includes(from4)) throw new Error('變異樣式找不到：章偏移');
    fs.writeFileSync(pf, pdfSrc.replace(from4, () => 'const STAMP_LINE_OFFSET_Y = 0;')); tmpFiles.push(pf);
    try { muting = true; mutFail = 0;
      try { const C = loadClient({ quote: { id: 'Q1', quoteNo: 'N', approval: { state: 'approved', valid: true }, perm: {}, items: [], products: [] }, issueResp: { issueNo: 'ISS-N-001', kind: 'official', shortHash: 'a', stampToken: 'T', watermark: false }, pdfFile: pf }); await C.ctx.exportQuotePdf('Q1', 'N');
        ok(C.state.mainCanvas.args[0][2] > 800 * 2, '（變異探針）章必須往下移到簽名線'); } catch (e) { mutFail++; }
      muting = false; ok(mutFail > 0, `變異 章位置偏移為 0 → 對應檢查必須失敗，實際失敗 ${mutFail} 項`);
    } finally { muting = false; try { fs.unlinkSync(pf); } catch { /* 已刪 */ } } }
  // M8 前端：draft 仍蓋章（取章不看 kind）
  const pdfFile2 = path.join(ROOT, '_client', 'quote-pdf.__mut2.js');
  const from2 = "const official = iss.kind === 'official';";
  fs.writeFileSync(pdfFile2, pdfSrc.replace(from2, () => "const official = true;")); tmpFiles.push(pdfFile2);
  try {
    muting = true; mutFail = 0;
    try {
      const C = loadClient({ quote: { id: 'Q1', quoteNo: 'N', approval: { state: 'none', valid: false }, perm: {}, items: [], products: [] }, issueResp: { issueNo: 'ISS-N-001', kind: 'draft', shortHash: 'a', stampToken: 'T' }, pdfFile: pdfFile2 });
      await C.ctx.exportQuotePdf('Q1', 'N');
      ok(!C.state.calls.some(c => /seal/.test(c)), '（變異探針）draft 不得取章');
    } catch (e) { mutFail++; if (process.env.QISS_MUTDBG) console.log('    exception:', e.message); }
    muting = false;
    ok(mutFail > 0, `變異 前端 draft 也取章 → 對應檢查必須失敗，實際失敗 ${mutFail} 項`);
  } finally { muting = false; try { fs.unlinkSync(pdfFile2); } catch { /* 已刪 */ } }
}
/** 變異 M2b：種一張「approval.state=approved 但雜湊不符」的單，變異版必須把它誤判成 official */
async function mutRunSeeded(name, edits) {
  const m = copyMut(null, edits);
  try {
    muting = true; mutFail = 0;
    try {
      await stopServer(); cleanupFiles(); extraEnv = {};
      const qs = seedFiles();
      const d = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
      d.quotations.find(x => x.id === 'Q6').items[0].unitPrice = 99999;   // 核准後內容被改（雜湊不符），但 approval.state 仍是 approved
      fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d));
      await startServer(m.server); jar = {}; await login('_s_owner');
      const r = await issue('_s_owner', 'Q6');
      ok(r.json.kind === 'draft' && !r.json.stampToken, '（變異探針）核准後被改過的單必須是 draft、不得有憑證');
    } catch (e) { mutFail++; if (process.env.QISS_MUTDBG) console.log('    exception:', e.message); }
    muting = false;
    ok(mutFail > 0, `變異 ${name} → 對應檢查必須失敗，實際失敗 ${mutFail} 項`);
  } finally { muting = false; await stopServer(); m.files.forEach(x => { try { fs.unlinkSync(x); } catch { /* 已刪 */ } }); }
}

(async () => {
  const t0 = Date.now();
  try {
    if (want('A')) await secA(require(path.join(ROOT, 'lib', 'quoteIssuance.js')));
    if (want('B')) await secB();
    if (want('C')) await secC();
    if (want('D')) await secD();
    if (want('E')) secE();
    if (want('M') && !process.env.QISS_SKIP_MUT) await secM();
  } catch (e) {
    fail++; failures.push('例外：' + (e && e.stack || e)); console.log('✗ 例外', e && e.stack || e);
    if (child && child.serverOut) console.log('[server out tail]', child.serverOut().slice(-800));
  } finally {
    await stopServer();
    cleanupFiles();
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } });
    for (const f of DATA_FILES) { try { fs.unlinkSync(f); } catch { /* 無 */ } }
  }
  console.log(`\n結果：通過 ${pass}，失敗 ${fail}（${((Date.now() - t0) / 1000).toFixed(1)} 秒）`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); process.exit(1); }
  process.exit(0);
})();
