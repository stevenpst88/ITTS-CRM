#!/usr/bin/env node
'use strict';
/**
 * 信件派送「端到端行程內模擬」：用合成使用者與合成單據，把 P1-P3 的全部新模組串起來，
 * 證明整合階段可以直接接上現有簽核流程。
 * 用法：node scripts/check-mail-integration.js（不開伺服器、不連網、不碰 data.json／auth.json／audit.log.json；只寫系統暫存資料夾）
 *
 * 管線（每一環都是真的模組，只有傳輸層是假的）：
 *   模擬簽核狀態機 → 整合膠水（本檔的 stand-in，等價於整合階段落在 lib/quoteRoutes.js notify() 旁的程式）
 *   → events.validateEvent／dedupeKey → recipients.resolveRecipients → render.renderMail（真樣板）
 *   → dispatcher（真）→ outbox（真，JSON 檔案 adapter 寫到暫存目錄）→ transports.selectTransport（真模式護欄）→ 假傳輸
 *
 * 涵蓋（章節）：
 *   1 E1／E2／E3（一般、需總經理、需董事長、董事會關）／E4／E5／E6 的收件人集合與「信件內文」可見性
 *   2 去重與重複觸發（含同時兩次）  3 缺 email、停用、網域不合法、操作者與重複帳號  4 off／log／redirect／live 與模式護欄
 *   5 失敗注入（TIMEOUT／AUTH／THROTTLED／5xx／拋例外／亂回傳／儲存體爆炸）與恢復  6 isStillValid／rebuild／同級已簽
 *   7 熔斷  8 Hobby 方案重試路徑（當次嘗試、機會式清理、每日 Cron、後台重送）  9 稽核呼叫內容與 outbox 儲存內容衛生
 *   10 P1 到 P2：批次補齊 email 後補寄  11 連結鏈：信、跳板頁、深層連結  12 與真實 lib/quoteApproval 銜接（若可載入）
 *   13 惡意內容穿過整條管線  14 dispatch／drainDue 永不 throw 總帳
 *
 * 「整合膠水」stand-in 與真實程式的對照（整合階段要改成用真的）：
 *   findManager1／boardProxySet／stepRecipients／isActive／canAct／dispName ＝ lib/quoteRoutes.js 同名函式的等價複本；
 *   noticeFilter ＝ notify() 的「排除操作者、去重、排除非在職」；buildEvent／numbersOf ＝ 要新增的「單據 → 信件事件」轉換。
 * 環境變數 MAIL_CORE_ROOT：要檢查的專案根目錄（變異測試用）。撰寫備註：本檔不寫四位數的 \uXXXX，一律用 cp()／\x..。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const util = require('util');
const assert = require('assert');

const ROOT = process.env.MAIL_CORE_ROOT || path.join(__dirname, '..');
const load = (rel) => require(path.join(ROOT, rel));

// ── 迷你測試框架（章節為 async）──────────────────────────────────────────────
const results = [];
const sections = [];
let currentSection = '';
function record(name, ok, extra) { results.push({ section: currentSection, name, ok: !!ok, extra: extra === undefined ? '' : String(extra) }); }
const t = (name, ok, extra) => record(name, ok, extra);
function short(v) {
  let s;
  try { s = typeof v === 'string' ? JSON.stringify(v) : util.inspect(v, { depth: 4, breakLength: Infinity }); } catch (e) { s = String(v); }
  return s.length > 240 ? s.slice(0, 240) + '…' : s;
}
function eq(name, actual, expected) {
  let ok = true;
  try { assert.deepStrictEqual(actual, expected); } catch (e) { ok = false; }
  record(name, ok, ok ? '' : 'actual=' + short(actual) + ' expected=' + short(expected));
}
function section(name, fn) { sections.push({ name, fn }); }
function finish() {
  const failed = results.filter((r) => !r.ok);
  const bySec = {};
  results.forEach((r) => { const s = bySec[r.section] || (bySec[r.section] = { p: 0, f: 0 }); if (r.ok) s.p++; else s.f++; });
  Object.keys(bySec).forEach((k) => console.log((bySec[k].f ? 'FAIL ' : 'ok   ') + k + '  通過 ' + bySec[k].p + (bySec[k].f ? '，失敗 ' + bySec[k].f : '')));
  failed.forEach((r) => console.log('  ✗ [' + r.section + '] ' + r.name + (r.extra ? '  → ' + r.extra : '')));
  console.log('dispatch／drainDue 呼叫 ' + GUARD.calls + ' 次，throw ' + GUARD.throws + ' 次，卡住 ' + GUARD.hangs + ' 次');
  console.log((failed.length ? 'FAILED' : 'PASSED') + '：' + (results.length - failed.length) + ' / ' + results.length);
  process.exit(failed.length ? 1 : 0);
}
async function main() {
  const watchdog = setTimeout(() => { console.log('WATCHDOG：測試超過 240 秒未結束'); process.exit(2); }, 240000);
  for (const s of sections) {
    currentSection = s.name;
    try { await s.fn(); } catch (e) { record('章節執行中例外', false, (e && e.stack) || e); }
  }
  clearTimeout(watchdog);
  finish();
}

// ── 載入被測模組（全部是真的）────────────────────────────────────────────────
const { getMailConfig } = load('lib/mail/config.js');
const EVT = load('lib/mail/events.js');
const RND = load('lib/mail/render.js');
const TRN = load('lib/mail/transports.js');
const SAF = load('lib/mail/safety.js');
const UE = load('lib/mail/userEmail.js');
const LNK = load('lib/mail/link.js');
const VIS = load('lib/mail/visibility.js');
const { createDispatcher } = load('lib/mail/dispatcher.js');
const { createOutbox, jsonFileAdapter, memoryAdapter } = load('lib/mail/outbox.js');
const DL = load('_client/deep-link.js');
let QA = null;
let QA_WHY = '';
try { QA = load('lib/quoteApproval.js'); } catch (e) { QA = null; QA_WHY = (e && e.code) || (e && e.name) || 'ERR'; }

const cp = (...n) => String.fromCodePoint(...n);
const T0 = Date.UTC(2026, 9, 8, 3, 0, 0);          // 2026-10-08 11:00（台北）
const SEC = 1000;
const iso = (ms) => new Date(ms).toISOString();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const tmpDirs = [];
function tmpDir() {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'mail-integ-test-'));
  tmpDirs.push(d);
  return d;
}
process.on('exit', () => { tmpDirs.forEach((d) => { try { fs.rmSync(d, { recursive: true, force: true }); } catch (e) { /* ignore */ } }); });

// ── 「永不 throw」總帳：所有 dispatch／drainDue 都經過這裡 ──────────────────────
const GUARD = { calls: 0, throws: 0, hangs: 0, bad: [] };
const emptySummary = () => ({ queued: 0, sent: 0, skipped: [], failed: [], cancelled: 0, errors: 0, guardFailed: true });
async function guard(label, fn) {
  GUARD.calls += 1;
  let timer;
  const hang = new Promise((res) => { timer = setTimeout(() => res({ hung: true }), 20000); });
  let r;
  try {
    r = await Promise.race([Promise.resolve().then(fn).then((value) => ({ value }), (error) => ({ error })), hang]);
  } finally { clearTimeout(timer); }
  if (r.hung) { GUARD.hangs += 1; GUARD.bad.push(label + ' 卡住'); return emptySummary(); }
  if (r.error) { GUARD.throws += 1; GUARD.bad.push(label + ' throw：' + (r.error && r.error.message)); return emptySummary(); }
  return r.value;
}

// ═════════════════════════════════════════════════════════════════════════
// 合成資料
// ═════════════════════════════════════════════════════════════════════════
const BASE_URL = 'https://crm.example.test';
const CUSTOMER = 'Synthetic Customer Co';
const PROJECT = 'Synthetic Project Alpha';
const OWNER_LABEL = 'Sales Nick';
const FULL_EMAIL_RE = /[A-Za-z0-9._+-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+/;
const qid = (n) => 'b3f1c2d4-1111-4222-8333-' + String(n).padStart(12, '0');

function mkUsers() {
  const mk = (username, role, extra) => Object.assign({ username, role, displayName: username.toUpperCase() + ' Name', email: username + '@itts.com.tw' }, extra || {});
  return [
    mk('admin1', 'admin'),
    mk('sales1', 'user', { nickname: OWNER_LABEL, supervisor: 'lead1' }),
    mk('sales2', 'user', { supervisor: 'lead1' }),
    mk('lead1', 'manager2', { supervisor: 'mgr1old' }),
    mk('mgr1old', 'manager1', { active: false, supervisor: 'mgr1a' }),    // 停用的一級主管：沿 supervisor 鏈會被跳過
    mk('mgr1a', 'manager1', { nickname: 'Mgr Nick' }),
    mk('gm1', 'executive', { nickname: 'GM One' }),
    mk('gm2', 'executive', { nickname: 'GM Two' }),
    mk('chair1', 'executive', { nickname: 'Chair One' }),
    mk('sec1', 'secretary', { nickname: 'Sec One' }),
    mk('sec2', 'secretary', { nickname: 'Sec Two' }),
    mk('sec3', 'secretary', { nickname: 'Sec Three' }),
    mk('proxy1', 'user', { nickname: 'Proxy One' }),
    mk('cons1', 'consult_manager_south', { nickname: 'Cons One' }),
  ];
}
const mkRoster = () => ({ gm: ['gm1', 'gm2'], chairman: ['chair1'], boardProxy: ['proxy1'], costProviders: ['cons1'], sealManagers: [] });

const STEP_LABEL = { mgr1: '一級主管', gm: '總經理', chairman: '董事長', board: '董事會決議（秘書代核）' };   // 與 lib/quoteApproval.js TIERS 相同
const STEP_LEVEL = { mgr1: 1, gm: 2, chairman: 3, board: 'board' };
const TIER_SHORT = { 1: '一級主管', 2: '總經理', 3: '董事長' };
// 送簽當下凍結的 approval.derived 的形狀（只列信件用得到的欄位）。金額單位：分
const DERIVED = {
  L1: { level: 1, board: false, tiers: ['mgr1'], revenueCents: 200000000, gpCents: 80000000, marginText: '40.00' },
  L2: { level: 2, board: false, tiers: ['mgr1', 'gm'], revenueCents: 200000000, gpCents: 40000000, marginText: '20.00' },
  L3: { level: 3, board: false, tiers: ['mgr1', 'gm', 'chairman'], revenueCents: 200000000, gpCents: 10000000, marginText: '5.00' },
  BOARD: { level: 3, board: true, tiers: ['mgr1', 'gm', 'board'], revenueCents: 6000000000, gpCents: 2400000000, marginText: '40.00' },
  LOSS: { level: 3, board: false, tiers: ['mgr1', 'gm', 'chairman'], revenueCents: 100000000, gpCents: -3210000, marginText: '-3.21' },
};
// 這些價格與成本數字絕不可出現在顧問信
const PRICE_NEEDLES = ['98765', '98,765', '12345', '12,345', '54321', '54,321', '6789', '6,789'];

function mkQuote(over) {
  return Object.assign({
    id: qid(1), quoteNo: 'QU-261008-001', owner: 'sales1', company: CUSTOMER, projectName: PROJECT, costBy: 'cons1',
    costFlow: { state: 'na' },
    items: [
      { lid: 'i1', desc: 'On-site installation service', qty: 2, unit: 'set', unitPrice: 98765, cost: 54321 },
      { lid: 'i2', kind: 'title', desc: 'Section header' },
      { lid: 'i3', desc: 'Training course', qty: 1.5, unit: 'day', unitPrice: 12345, cost: 6789 },
      { lid: 'i4', kind: 'subtotal', desc: 'Subtotal' },
    ],
    approval: { state: 'none', history: [] },
  }, over || {});
}

// ═════════════════════════════════════════════════════════════════════════
// 整合膠水 stand-in（等價於整合階段要寫的程式；對照見檔頭）
// ═════════════════════════════════════════════════════════════════════════
const isActive = (u) => !!u && u.active !== false && u.role !== 'pool';
const canAct = (u) => isActive(u) && u.role !== 'admin' && u.accessMode !== 'view';
const dispName = (w, un) => { const u = w.userMap[un]; return (u && (u.nickname || u.displayName || u.username)) || un || ''; };

function findManager1(w, ownerUsername) {
  const seen = new Set();
  let cur = w.userMap[ownerUsername];
  while (cur && cur.supervisor && !seen.has(cur.username)) {
    seen.add(cur.username);
    const up = w.userMap[cur.supervisor];
    if (!up) break;
    if (up.role === 'manager1' && canAct(up)) return up.username;
    cur = up;
  }
  return null;
}
function boardProxySet(w) {
  const s = new Set(w.roster.boardProxy);
  Object.values(w.userMap).forEach((u) => { if (u.role === 'secretary') s.add(u.username); });
  return new Set([...s].filter((un) => canAct(w.userMap[un])));
}
function stepRecipients(w, step) {
  if (!step) return [];
  if (step.tier === 'mgr1') return [step.assignee];
  if (step.tier === 'gm') return w.roster.gm.slice();
  if (step.tier === 'chairman') return w.roster.chairman.slice();
  if (step.tier === 'board') return [...boardProxySet(w)];
  return [];
}
/** notify() 的前置過濾：排除操作者、去重、排除非在職。w.glueFilter=false 時原樣交給信件層（用來驗證信件層自己的檢查）。 */
function noticeFilter(w, names, actor) {
  if (!w.glueFilter) return names.slice();
  const seen = new Set();
  const out = [];
  names.forEach((un) => {
    if (!un || un === actor || seen.has(un) || !isActive(w.userMap[un])) return;
    seen.add(un);
    out.push(un);
  });
  return out;
}
function kindOf(tier, user) {
  if (tier === 'mgr1') return 'mgr1';
  if (tier === 'gm') return 'gm';
  if (tier === 'chairman') return 'chairman';
  if (tier === 'owner') return 'owner';
  if (tier === 'board') return user && user.role === 'secretary' ? 'secretary' : 'boardProxy';
  return 'mgr1';
}
function numbersOf(d) {
  if (!d || typeof d.marginText !== 'string') return null;
  const pct = Number(d.marginText);
  if (!Number.isFinite(pct)) return null;
  return { revenueCents: d.revenueCents, gpCents: d.gpCents, marginText: d.marginText, marginPct: pct, tierLevel: d.level, tierLabel: d.board ? '董事會' : TIER_SHORT[d.level] };
}
const itemsOf = (q) => (q.items || []).filter((it) => !it.kind || it.kind === 'item').map((it) => ({ desc: it.desc, qty: it.qty, unit: it.unit }));
const stepOfStep = (s) => ({ level: STEP_LEVEL[s.tier], label: s.label });

const keyStep = (q, idx) => q.approval.submittedAt + '#' + idx;
const keyResult = (q, kind, idx) => q.approval.submittedAt + '#r:' + kind + ':' + idx;
const keyE6 = (q, kind) => q.approval.submittedAt + '#' + kind;
const keyCost = (q) => q.costFlow.requestedAt + '#' + q.costBy;
const keyCostDone = (q) => q.costFlow.filledAt + '#done';

function buildEvent(w, q, type, o) {
  const ap = q.approval || {};
  const ev = { type, quoteId: q.id, quoteNo: q.quoteNo, projectName: q.projectName, company: q.company, ownerLabel: dispName(w, q.owner), at: o.at || iso(w.clock.t), stepKey: o.stepKey };
  if (o.actor) ev.actor = { label: dispName(w, o.actor) };
  if (o.step) ev.step = o.step;
  if (type === 'E1_SUBMIT' || type === 'E3_NEXT_STEP') ev.numbers = numbersOf(ap.derived);
  if (type === 'E4_RESULT' || type === 'E6_WITHDRAWN') ev.result = o.reason ? { kind: o.resultKind, reason: o.reason } : { kind: o.resultKind };
  if (type === 'E2_COST_REQUEST') ev.items = itemsOf(q);
  return ev;
}

// ── 狀態機 + 通知（等價於 quoteRoutes.js 各路由裡 notify() 的呼叫點）──
async function sendStep(w, q, type, idx, actor) {
  const ap = q.approval;
  const step = ap.steps[idx];
  const ev = buildEvent(w, q, type, { stepKey: keyStep(q, idx), step: stepOfStep(step), actor });
  const rc = noticeFilter(w, stepRecipients(w, step), actor).map((un) => ({ username: un, kind: kindOf(step.tier, w.userMap[un]) }));
  return w.dispatch(ev, rc, { actorUsername: actor, operatorLabel: dispName(w, actor) });
}
async function sendResult(w, q, resultKind, idx, actor, reason) {
  const ap = q.approval;
  const ev = buildEvent(w, q, 'E4_RESULT', { stepKey: keyResult(q, resultKind, idx), step: stepOfStep(ap.steps[idx]), actor, resultKind, reason });
  const rc = noticeFilter(w, [q.owner], actor).map((un) => ({ username: un, kind: 'owner' }));
  return w.dispatch(ev, rc, { actorUsername: actor, operatorLabel: dispName(w, actor) });
}
async function submitQuote(w, q, actor, derived) {
  const d = typeof derived === 'string' ? DERIVED[derived] : derived;
  const ap = q.approval;
  ap.state = 'pending'; ap.submittedAt = iso(w.clock.t); ap.submittedBy = actor; ap.derived = d; ap.cur = 0; ap.history = [];
  ap.steps = d.tiers.map((tier, i) => ({ tier, label: (QA && QA.TIERS && QA.TIERS[tier]) || STEP_LABEL[tier], assignee: tier === 'mgr1' ? findManager1(w, q.owner) : null, status: i === 0 ? 'pending' : 'waiting' }));
  return { e1: await sendStep(w, q, 'E1_SUBMIT', 0, actor) };
}
async function approveStep(w, q, actor) {
  const ap = q.approval;
  const idx = ap.cur;
  const step = ap.steps[idx];
  step.status = 'approved'; step.by = actor; step.at = iso(w.clock.t);
  const out = {};
  if (idx === ap.steps.length - 1) {
    ap.state = 'approved'; ap.cur = ap.steps.length;
    out.e4 = await sendResult(w, q, 'final_approved', idx, actor);
  } else {
    ap.cur = idx + 1; ap.steps[idx + 1].status = 'pending';
    out.e4 = await sendResult(w, q, 'approved', idx, actor);
    out.e3 = await sendStep(w, q, 'E3_NEXT_STEP', idx + 1, actor);
  }
  return out;
}
async function rejectStep(w, q, actor, reason) {
  const ap = q.approval;
  const idx = ap.cur;
  ap.steps[idx].status = 'returned'; ap.steps[idx].comment = reason; ap.state = 'returned';
  return { e4: await sendResult(w, q, 'rejected', idx, actor, reason) };
}
function e6Recipients(w, q, actor, includeOwner) {
  const ap = q.approval;
  const list = [];
  if (ap.state === 'pending' && ap.steps[ap.cur]) stepRecipients(w, ap.steps[ap.cur]).forEach((un) => list.push({ un, tier: ap.steps[ap.cur].tier }));
  (ap.steps || []).forEach((s) => { if (s.status === 'approved' && s.by) list.push({ un: s.by, tier: s.tier }); });
  if (includeOwner) list.push({ un: q.owner, tier: 'owner' });
  const keep = new Set(noticeFilter(w, list.map((x) => x.un), actor));
  const seen = new Set();
  const out = [];
  list.forEach((x) => { if (keep.has(x.un) && !seen.has(x.un)) { seen.add(x.un); out.push({ username: x.un, kind: kindOf(x.tier, w.userMap[x.un]) }); } });
  if (!w.glueFilter) { out.length = 0; list.forEach((x) => out.push({ username: x.un, kind: kindOf(x.tier, w.userMap[x.un]) })); }
  return out;
}
async function endApproval(w, q, actor, resultKind, includeOwner) {
  const ap = q.approval;
  const originStep = ap.steps && ap.steps[Math.min(ap.cur, ap.steps.length - 1)] ? stepOfStep(ap.steps[Math.min(ap.cur, ap.steps.length - 1)]) : undefined;
  const rc = e6Recipients(w, q, actor, includeOwner);
  const ev = buildEvent(w, q, 'E6_WITHDRAWN', { stepKey: keyE6(q, resultKind), step: originStep, actor, resultKind });
  ap.state = 'none'; ap.steps = []; ap.derived = null;                       // 與 quoteRoutes：撤回／作廢後清空 steps 與 derived
  return { e6: await w.dispatch(ev, rc, { actorUsername: actor, operatorLabel: dispName(w, actor) }) };
}
const withdrawQuote = (w, q, actor) => endApproval(w, q, actor, 'withdrawn', false);
const voidQuote = (w, q, actor) => endApproval(w, q, actor, 'voided', true);
async function requestCost(w, q, actor) {
  q.costFlow = { state: 'requested', requestedAt: iso(w.clock.t) };
  const ev = buildEvent(w, q, 'E2_COST_REQUEST', { stepKey: keyCost(q), actor });
  const rc = noticeFilter(w, [q.costBy], actor).map((un) => ({ username: un, kind: 'consultant' }));
  return { e2: await w.dispatch(ev, rc, { actorUsername: actor, operatorLabel: dispName(w, actor) }) };
}
async function completeCost(w, q, actor) {
  q.costFlow = { state: 'filled', requestedAt: q.costFlow.requestedAt, filledAt: iso(w.clock.t) };
  const ev = buildEvent(w, q, 'E5_COST_DONE', { stepKey: keyCostDone(q), actor });
  const rc = noticeFilter(w, [q.owner], actor).map((un) => ({ username: un, kind: 'owner' }));
  return { e5: await w.dispatch(ev, rc, { actorUsername: actor, operatorLabel: dispName(w, actor) }) };
}

// ── drainDue／isStillValid 需要的兩個回呼 ──
const stepKeyOf = (job) => job.dedupeKey.split(':').slice(3).join(':');          // type:quoteId:user:stepKey（user 的 ':' 已跳脫成 %3A，所以前三段不含 ':'）
function isStillValid(w, job) {
  const q = w.quotes.get(job.quoteId);
  if (!q) return false;
  const ap = q.approval || {};
  const sk = stepKeyOf(job);
  if (job.type === 'E1_SUBMIT' || job.type === 'E3_NEXT_STEP') {
    const idx = Number(sk.split('#')[1]);
    if (ap.state !== 'pending' || sk !== ap.submittedAt + '#' + idx || ap.cur !== idx) return false;
    return stepRecipients(w, ap.steps[idx]).indexOf(job.toUser) >= 0;
  }
  if (job.type === 'E2_COST_REQUEST') return !!q.costFlow && q.costFlow.state === 'requested' && q.costBy === job.toUser;
  return true;
}
function rebuild(w, job) {
  const q = w.quotes.get(job.quoteId);
  if (!q) return null;
  const ap = q.approval || {};
  const sk = stepKeyOf(job);
  const base = { stepKey: sk, at: job.createdAt };                                // 重建的事件用工作建立時間，所以重試信與首次嘗試逐字相同
  if (job.type === 'E1_SUBMIT' || job.type === 'E3_NEXT_STEP') {
    const idx = Number(sk.split('#')[1]);
    if (!ap.steps || !ap.steps[idx]) return null;                                 // 撤回／作廢後 steps 已清空：沒有東西可寄
    const actor = job.type === 'E1_SUBMIT' ? ap.submittedBy : (ap.steps[idx - 1] && ap.steps[idx - 1].by);
    return { ev: buildEvent(w, q, job.type, Object.assign(base, { step: stepOfStep(ap.steps[idx]), actor })) };
  }
  if (job.type === 'E2_COST_REQUEST') return { ev: buildEvent(w, q, job.type, Object.assign(base, { actor: q.owner })) };
  if (job.type === 'E4_RESULT') {
    const m = /#r:([a-z_]+):([0-9]+)$/.exec(sk);
    if (!m) return null;
    const step = ap.steps && ap.steps[Number(m[2])];
    return { ev: buildEvent(w, q, job.type, Object.assign(base, { resultKind: m[1], step: step ? stepOfStep(step) : undefined, reason: step ? step.comment : undefined })) };
  }
  if (job.type === 'E5_COST_DONE') return { ev: buildEvent(w, q, job.type, Object.assign(base, { actor: q.costBy })) };
  if (job.type === 'E6_WITHDRAWN') {
    const m = /#(withdrawn|voided)$/.exec(sk);
    if (!m) return null;
    return { ev: buildEvent(w, q, job.type, Object.assign(base, { resultKind: m[1], actor: q.owner })) };
  }
  return null;
}

// ═════════════════════════════════════════════════════════════════════════
// 世界（一個獨立的 CRM 模擬：假時鐘、JSON 檔案 outbox、可替換行為的假傳輸、稽核收集器）
// ═════════════════════════════════════════════════════════════════════════
const WORLDS = [];
let QSEQ = 0;
function newQuote(w, over) {
  QSEQ += 1;
  const q = mkQuote(Object.assign({ id: qid(QSEQ), quoteNo: 'QU-261008-' + String(QSEQ).padStart(3, '0') }, over || {}));
  w.quotes.set(q.id, q);
  return q;
}
function mkWorld(o) {
  const opt = o || {};
  const dir = tmpDir();
  const file = path.join(dir, 'mail-outbox.json');
  const clock = { t: T0, advance(ms) { this.t += ms; } };
  const env = Object.assign({ MAIL_MODE: 'live', MAIL_PREVIEW_DIR: path.join(dir, 'prev'), MAIL_OUTBOX_FILE: file, APP_BASE_URL: BASE_URL }, opt.env || {});
  const config = getMailConfig(env);
  config.timeouts = opt.timeouts || { connectMs: 20, totalMs: 120 };
  if (opt.config) Object.assign(config, opt.config);
  const adapter = opt.adapter || (opt.memory ? memoryAdapter() : jsonFileAdapter({ file }));
  const outbox = createOutbox(adapter, { now: () => clock.t, config });
  const w = {
    dir, file, clock, config, adapter, outbox, env, sent: [], logs: [], signals: [], quotes: new Map(),
    renderCount: 0, getUsersCalls: 0, behavior: { fn: null }, glueFilter: opt.glueFilter !== false, roster: mkRoster(),
  };
  w.setUsers = (arr) => { w.users = arr; w.userMap = {}; arr.forEach((u) => { w.userMap[u.username] = u; }); };
  w.setUsers(opt.users || mkUsers());
  if (opt.roster) w.roster = opt.roster;
  w.transport = {
    kind: 'FAKE',
    async send(msg, sendOpts) {
      w.sent.push(msg);
      w.signals.push(sendOpts && sendOpts.signal);
      return w.behavior.fn ? await w.behavior.fn(msg, sendOpts, w.sent.length) : { ok: true, providerId: 'fake:' + w.sent.length };
    },
  };
  const disp = createDispatcher({
    config, outbox, transport: w.transport,
    getUsers: opt.getUsers || (() => { w.getUsersCalls += 1; return opt.usersForm === 'array' ? w.users : w.userMap; }),
    isStillValid: opt.isStillValid === undefined ? (job) => isStillValid(w, job) : opt.isStillValid,
    writeLog: opt.writeLog || ((...a) => { w.logs.push(a); }),
    render: (ev, viewer, ctx) => { w.renderCount += 1; return RND.renderMail(ev, viewer, ctx); },
    rebuild: (job) => rebuild(w, job),
    now: () => clock.t, rootDir: dir, auxTimeoutMs: 2000,
  });
  w.dispatch = (ev, rcs, ops) => guard('dispatch', () => disp.dispatch(ev, rcs, ops));
  w.drainDue = (ops) => guard('drainDue', () => disp.drainDue(ops));
  w.records = async () => (await outbox.list({ limit: 500 })).rows;
  w.emails = () => w.sent.map((m) => m.to[0]).sort();
  w.link = (q, cost) => LNK.buildQuoteLink(w.config, q.id, cost ? { cost: true } : undefined);
  WORLDS.push(w);
  return w;
}

// ═════════════════════════════════════════════════════════════════════════
// 信件內容斷言（用實際送到傳輸層的主旨／html／text，不是看 meta）
// ═════════════════════════════════════════════════════════════════════════
// 獨立抄錄的可見性期望（不是從 visibility.js 推導，避免循環驗證）：null＝不在乎（業務本人的稱呼可能就是業務名）
// project：專案名稱（業主 2026-10-08 決定：秘書與董事會代核人的信只放單號，不放專案名稱）
const EXPECT = {
  mgr1: { amount: true, customer: true, owner: true, project: true },
  gm: { amount: true, customer: true, owner: true, project: true },
  chairman: { amount: true, customer: true, owner: true, project: true },
  secretary: { amount: true, customer: false, owner: false, project: false },
  boardProxy: { amount: true, customer: false, owner: false, project: false },
  consultant: { amount: false, customer: false, owner: true, project: true },
  owner: { amount: false, customer: true, owner: null, project: true },
};
const blobOf = (m) => m.subject + '\n' + m.html + '\n' + m.text;
function checkMail(name, m, kind, type, ctx) {
  const c = Object.assign({ amount: 'NT$ 2,000,000', amountDigits: '2,000,000', margin: '40.00%', marginDigits: '40.00', tag: null, customer: CUSTOMER, owner: OWNER_LABEL, project: PROJECT, link: null }, ctx || {});
  const E = EXPECT[kind];
  const decision = type === 'E1_SUBMIT' || type === 'E3_NEXT_STEP';
  const showMoney = !!(E && E.amount && decision);
  const b = blobOf(m);
  const tag = name + '｜' + kind + '｜' + type;
  t(tag + '：主旨以事件對應的類別開頭、不含金額／毛利率／客戶名／業務名', /^【(簽核通知|請填寫成本|簽核結果|成本已填寫|簽核撤回)】/.test(m.subject) && !/NT\$|[0-9],[0-9]{3}|[0-9]+\.[0-9]{2}%?/.test(m.subject) && m.subject.indexOf(c.customer) < 0 && m.subject.indexOf(c.owner) < 0, m.subject);
  if (showMoney) {
    t(tag + '：html 與純文字都有金額與毛利率', m.html.indexOf(c.amountDigits) >= 0 && m.text.indexOf(c.amountDigits) >= 0 && m.html.indexOf(c.margin) >= 0 && m.text.indexOf(c.margin) >= 0, short(m.text.slice(0, 200)));
    t(tag + '：金額標示「折扣後未稅」', m.html.indexOf('折扣後未稅') >= 0 && m.text.indexOf('折扣後未稅') >= 0);
    if (c.tag) t(tag + '：核決層級文字標籤「' + c.tag + '」（html 與純文字）', m.html.indexOf(c.tag) >= 0 && m.text.indexOf(c.tag) >= 0);
  } else {
    t(tag + '：整封信沒有任何金額與毛利率（NT$／金額數字／毛利率數字／「毛利」）', !/NT\$/.test(b) && b.indexOf(c.amountDigits) < 0 && b.indexOf(c.marginDigits) < 0 && b.indexOf('毛利') < 0 && !/[0-9]\s*%/.test(m.text));
  }
  if (E && E.customer !== null) {
    const both = m.html.indexOf(c.customer) >= 0 && m.text.indexOf(c.customer) >= 0;
    const none = b.indexOf(c.customer) < 0;
    t(tag + '：客戶名' + (E.customer ? '出現在 html 與純文字' : '完全不出現（主旨、html、純文字）'), E.customer ? both : none);
  }
  if (E && E.owner !== null) {
    const both = m.html.indexOf(c.owner) >= 0 && m.text.indexOf(c.owner) >= 0;
    const none = b.indexOf(c.owner) < 0 && !/sales1/i.test(b);
    t(tag + '：業務名' + (E.owner ? '出現在 html 與純文字' : '完全不出現（主旨、html、純文字，也沒有業務帳號）'), E.owner ? both : none);
  }
  if (E && E.project === false) {
    t(tag + '：專案名稱完全不出現（主旨、html、純文字，只放單號）', b.indexOf(c.project) < 0 && b.indexOf('專案名稱') < 0, m.subject);
    t(tag + '：主旨只有類別與單號（單號之後最多只有「請勿簽核」）', /^【[^】]+】\S+( 請勿簽核)?$/.test(m.subject), m.subject);
  } else {
    t(tag + '：專案名稱在 html 與純文字', m.html.indexOf(c.project) >= 0 && m.text.indexOf(c.project) >= 0);
  }
  if (c.link) {
    const hrefs = Array.from(new Set((m.html.match(/href="[^"]*"/g) || []).map((x) => x.slice(6, -1))));
    t(tag + '：連結 ' + c.link.replace(BASE_URL, '') + '（html 只有這一個 href、純文字也有、無 token）', hrefs.length === 1 && hrefs[0] === c.link && m.text.indexOf(c.link) >= 0 && !/token|[?&]t=/i.test(b), short(hrefs));
  }
  t(tag + '：頁尾「機密，請勿轉寄」（html 與純文字）', m.html.indexOf('機密，請勿轉寄') >= 0 && m.text.indexOf('機密，請勿轉寄') >= 0);
  t(tag + '：通過傳輸層訊息驗證、單封 <100KB、只有一位收件人', TRN.validateMessage(m).ok === true && Buffer.byteLength(m.html + m.text, 'utf8') < 100 * 1024 && m.to.length === 1);
}

// ═════════════════════════════════════════════════════════════════════════
section('1 收件人集合與內容可見性（live、JSON outbox、真 render）', async () => {
  const w = mkWorld();
  const slice = (n) => w.sent.slice(n);
  const L = (q, cost) => w.link(q, cost);

  // ── E1：一般單（一級主管可核）──
  const q1 = newQuote(w);
  let r = await submitQuote(w, q1, 'sales1', 'L1');
  eq('E1：Summary（寄出 1 封，其餘為空）', [r.e1.sent, r.e1.queued, r.e1.skipped, r.e1.failed, r.e1.errors], [1, 0, [], [], 0]);
  eq('E1：只寄給沿 supervisor 鏈找到的在職一級主管（停用的 mgr1old、manager2 的 lead1 都沒收到）', w.emails(), ['mgr1a@itts.com.tw']);
  eq('E1：主旨＝【簽核通知】單號 專案名稱', w.sent[0].subject, '【簽核通知】' + q1.quoteNo + ' ' + PROJECT);
  eq('E1：傳輸標籤＝事件型別', w.sent[0].tag, 'E1_SUBMIT');
  checkMail('E1 一般', w.sent[0], 'mgr1', 'E1_SUBMIT', { tag: '一級主管可核', link: L(q1) });
  t('E1：稱呼用暱稱（Mgr Nick）', w.sent[0].html.indexOf('Mgr Nick') >= 0 && w.sent[0].text.indexOf('Mgr Nick') >= 0);
  t('E1：時間顯示台北時間 2026-10-08 11:00', w.sent[0].text.indexOf('2026-10-08 11:00') >= 0);
  let recs = await w.records();
  eq('E1：outbox 一筆 sent（遮罩位址、嘗試 1 次、關卡與 kind）', recs.map((x) => [x.type, x.toUser, x.toMasked, x.status, x.attempts, x.meta]), [['E1_SUBMIT', 'mgr1a', 'm***@itts.com.tw', 'sent', 1, { level: 1, kind: 'mgr1' }]]);
  eq('E1：成功不寫稽核', w.logs, []);

  // E4：一級主管直接最終核准 → 通知業務（沒有金額）
  let n = w.sent.length;
  r = await approveStep(w, q1, 'mgr1a');
  eq('E4（最終核准）：寄給業務 1 封', [r.e4.sent, w.emails().slice(0).filter((e) => e === 'sales1@itts.com.tw').length, w.sent.length - n], [1, 1, 1]);
  checkMail('E4 最終核准', w.sent[n], 'owner', 'E4_RESULT', { link: L(q1) });
  t('E4：結果文字「已完成核准」', w.sent[n].html.indexOf('已完成核准') >= 0 && w.sent[n].text.indexOf('已完成核准') >= 0);
  t('E4：沒有再寄 E3（路徑只有一關）', r.e3 === undefined && w.sent.length === n + 1);

  // ── E1→E3：需總經理（L2）──
  const q2 = newQuote(w);
  n = w.sent.length;
  await submitQuote(w, q2, 'sales1', 'L2');
  r = await approveStep(w, q2, 'mgr1a');
  const m2 = slice(n);
  eq('L2：E1 給一級主管、E4（本關通過）給業務、E3 給名冊內兩位總經理', m2.map((m) => m.tag + '>' + m.to[0]).sort(), ['E1_SUBMIT>mgr1a@itts.com.tw', 'E3_NEXT_STEP>gm1@itts.com.tw', 'E3_NEXT_STEP>gm2@itts.com.tw', 'E4_RESULT>sales1@itts.com.tw']);
  eq('L2：E3 收件人＝總經理名冊全員（不含操作者 mgr1a）', m2.filter((m) => m.tag === 'E3_NEXT_STEP').map((m) => m.to[0]).sort(), ['gm1@itts.com.tw', 'gm2@itts.com.tw']);
  m2.filter((m) => m.tag === 'E3_NEXT_STEP').forEach((m) => checkMail('L2 總經理關', m, 'gm', 'E3_NEXT_STEP', { margin: '20.00%', marginDigits: '20.00', tag: '需總經理核准', link: L(q2) }));
  const e4a = m2.find((m) => m.tag === 'E4_RESULT');
  checkMail('L2 本關通過', e4a, 'owner', 'E4_RESULT', { marginDigits: '20.00', link: L(q2) });
  t('L2：E4 本關通過的結果文字「本關已核准」', e4a.html.indexOf('本關已核准') >= 0);
  n = w.sent.length;
  r = await approveStep(w, q2, 'gm2');
  eq('L2：總經理核准＝最終核准，只通知業務', [r.e4.sent, r.e3, w.sent.length - n, w.sent[n].to[0], w.sent[n].tag], [1, undefined, 1, 'sales1@itts.com.tw', 'E4_RESULT']);

  // ── 需董事長（L3）──
  const q3 = newQuote(w);
  await submitQuote(w, q3, 'sales1', 'L3');
  await approveStep(w, q3, 'mgr1a');
  n = w.sent.length;
  r = await approveStep(w, q3, 'gm1');
  const m3 = slice(n).filter((m) => m.tag === 'E3_NEXT_STEP');
  eq('L3：董事長關 E3 只寄給董事長名冊', m3.map((m) => m.to[0]), ['chair1@itts.com.tw']);
  checkMail('L3 董事長關', m3[0], 'chairman', 'E3_NEXT_STEP', { margin: '5.00%', marginDigits: '5.00', tag: '需董事長核准', link: L(q3) });
  t('L3：總經理關核准後的通知只有「通知業務」與「董事長關」，沒有再寄給另一位總經理（gm2）或操作者（gm1）', slice(n).length === 2 && slice(n).every((m) => ['sales1@itts.com.tw', 'chair1@itts.com.tw'].indexOf(m.to[0]) >= 0), short(slice(n).map((m) => m.to[0])));

  // ── 董事會關（金額 > 5,000 萬）：秘書與代核人信不帶客戶名與業務名 ──
  const q4 = newQuote(w);
  await submitQuote(w, q4, 'sales1', 'BOARD');
  await approveStep(w, q4, 'mgr1a');
  n = w.sent.length;
  r = await approveStep(w, q4, 'gm1');
  const m4 = slice(n);
  const e3b = m4.filter((m) => m.tag === 'E3_NEXT_STEP');
  eq('董事會關：收件人＝所有在職秘書 ∪ 代核名冊（共 4 位），不含總經理／一級主管', e3b.map((m) => m.to[0]).sort(), ['proxy1@itts.com.tw', 'sec1@itts.com.tw', 'sec2@itts.com.tw', 'sec3@itts.com.tw']);
  const boardCtx = { amount: 'NT$ 60,000,000', amountDigits: '60,000,000', margin: '40.00%', tag: '需董事會決議', link: L(q4) };
  e3b.forEach((m) => checkMail('董事會關', m, m.to[0] === 'proxy1@itts.com.tw' ? 'boardProxy' : 'secretary', 'E3_NEXT_STEP', boardCtx));
  t('董事會關：4 封信的主旨、html、純文字都沒有客戶名、業務名、業務帳號', e3b.every((m) => blobOf(m).indexOf(CUSTOMER) < 0 && blobOf(m).indexOf(OWNER_LABEL) < 0 && !/sales1|SALES1/.test(blobOf(m))));
  t('董事會關：目前關卡顯示「董事會決議（秘書代核）」', e3b.every((m) => m.text.indexOf('董事會決議（秘書代核）') >= 0));
  const recB = (await w.records()).filter((x) => x.quoteNo === q4.quoteNo && x.type === 'E3_NEXT_STEP' && x.meta.level === 'board');
  eq('董事會關：outbox 記 kind（3 位 secretary、1 位 boardProxy）與 level=board', recB.map((x) => x.meta.kind).sort(), ['boardProxy', 'secretary', 'secretary', 'secretary']);
  n = w.sent.length;
  r = await approveStep(w, q4, 'sec2');
  eq('董事會關：秘書代核＝最終核准，通知業務', [r.e4.sent, w.sent[n].to[0], w.sent[n].tag], [1, 'sales1@itts.com.tw', 'E4_RESULT']);

  // ── 駁回（E4 rejected）──
  const q5 = newQuote(w);
  await submitQuote(w, q5, 'sales1', 'L2');
  n = w.sent.length;
  r = await rejectStep(w, q5, 'mgr1a', '資料不齊請補件');
  checkMail('駁回', w.sent[n], 'owner', 'E4_RESULT', { link: L(q5) });
  t('駁回：結果文字「已駁回」與原因都在信裡', w.sent[n].html.indexOf('已駁回') >= 0 && w.sent[n].html.indexOf('資料不齊請補件') >= 0 && w.sent[n].text.indexOf('資料不齊請補件') >= 0);

  // ── E2：請顧問填成本（不含任何金額）──
  const q6 = newQuote(w);
  n = w.sent.length;
  r = await requestCost(w, q6, 'sales1');
  eq('E2：只寄給被指定的顧問', [r.e2.sent, w.sent.length - n, w.sent[n].to[0], w.sent[n].tag], [1, 1, 'cons1@itts.com.tw', 'E2_COST_REQUEST']);
  const e2m = w.sent[n];
  checkMail('E2', e2m, 'consultant', 'E2_COST_REQUEST', { link: L(q6, true) });
  eq('E2：主旨＝【請填寫成本】單號 專案名稱', e2m.subject, '【請填寫成本】' + q6.quoteNo + ' ' + PROJECT);
  t('E2：連結帶 ?cost=1', e2m.html.indexOf('/q/' + q6.id + '?cost=1') >= 0 && e2m.text.indexOf('/q/' + q6.id + '?cost=1') >= 0);
  t('E2：品項說明、數量、單位在信裡（標題列與小計列不列入）', e2m.text.indexOf('On-site installation service') >= 0 && e2m.text.indexOf('Training course') >= 0 && e2m.text.indexOf('Section header') < 0 && e2m.text.indexOf('Subtotal') < 0 && e2m.html.indexOf('set') >= 0);
  t('E2：沒有任何價格或成本數字（單價、成本）', PRICE_NEEDLES.every((s) => blobOf(e2m).indexOf(s) < 0));
  t('E2：沒有「折扣」「單價」「未稅」「毛利」字樣', !/折扣|單價|未稅|毛利/.test(blobOf(e2m)));
  // 粗心的膠水：把 numbers 與逐列價格／成本整包塞進事件，信件層仍然不可外洩
  n = w.sent.length;
  const careless = Object.assign(buildEvent(w, q6, 'E2_COST_REQUEST', { stepKey: 'careless#1', actor: 'sales1' }), {
    numbers: numbersOf(DERIVED.L1),
    items: q6.items.filter((it) => !it.kind).map((it) => Object.assign({}, it, { price: 12345, cost: 54321 })),
  });
  await w.dispatch(careless, [{ username: 'cons1', kind: 'consultant' }], { actorUsername: 'sales1' });
  const cm = w.sent[n];
  t('E2（粗心膠水：事件帶 numbers 與逐列 unitPrice／cost／price）：顧問信仍然沒有任何金額、毛利、價格', !/NT\$/.test(blobOf(cm)) && blobOf(cm).indexOf('2,000,000') < 0 && blobOf(cm).indexOf('40.00') < 0 && PRICE_NEEDLES.every((s) => blobOf(cm).indexOf(s) < 0) && !/折扣|單價|未稅|毛利/.test(blobOf(cm)));

  // 粗心的膠水：在不需要金額的事件（E4／E5／E6）也夾帶了 numbers，信件層仍然不可放決策條、金額或毛利率
  {
    const cases = [
      ['E4_RESULT', 'owner', { resultKind: 'final_approved' }, 'sales1'],
      ['E5_COST_DONE', 'owner', {}, 'sales1'],
      ['E6_WITHDRAWN', 'mgr1', { resultKind: 'withdrawn' }, 'mgr1a'],
      ['E6_WITHDRAWN', 'gm', { resultKind: 'voided' }, 'gm1'],
      ['E6_WITHDRAWN', 'secretary', { resultKind: 'voided' }, 'sec1'],
      ['E6_WITHDRAWN', 'owner', { resultKind: 'voided' }, 'sales1'],
    ];
    for (const [type, kind, extra, un] of cases) {
      const nb = w.sent.length;
      const evx = Object.assign(buildEvent(w, q6, type, Object.assign({ stepKey: 'careless#' + type + '#' + kind, actor: 'admin1' }, extra)), { numbers: numbersOf(DERIVED.L1) });
      await w.dispatch(evx, [{ username: un, kind }], { actorUsername: 'admin1' });
      const mx = w.sent[nb];
      t('粗心膠水（' + type + ' 夾帶 numbers）→ ' + kind + ' 的信仍沒有決策條、金額、毛利率', !!mx && !/NT\$/.test(blobOf(mx)) && blobOf(mx).indexOf('2,000,000') < 0 && blobOf(mx).indexOf('40.00') < 0 && blobOf(mx).indexOf('毛利') < 0 && blobOf(mx).indexOf('一級主管可核') < 0, mx ? '' : '沒有寄出');
    }
  }

  // ── E5：顧問完成成本 → 業務 ──
  n = w.sent.length;
  w.clock.advance(30 * 60 * SEC);
  r = await completeCost(w, q6, 'cons1');
  eq('E5：只寄給業務', [r.e5.sent, w.sent[n].to[0], w.sent[n].tag], [1, 'sales1@itts.com.tw', 'E5_COST_DONE']);
  checkMail('E5', w.sent[n], 'owner', 'E5_COST_DONE', { link: L(q6) });
  eq('E5：主旨＝【成本已填寫】單號 專案名稱', w.sent[n].subject, '【成本已填寫】' + q6.quoteNo + ' ' + PROJECT);

  // ── E6：撤回（目前關卡收件人＋已簽過的人）──
  const q7 = newQuote(w);
  await submitQuote(w, q7, 'sales1', 'L2');
  await approveStep(w, q7, 'mgr1a');
  n = w.sent.length;
  r = await withdrawQuote(w, q7, 'sales1');
  const m7 = slice(n);
  eq('E6 撤回（總經理關）：收件人＝目前關卡的兩位總經理＋已簽過的一級主管', m7.map((m) => m.to[0]).sort(), ['gm1@itts.com.tw', 'gm2@itts.com.tw', 'mgr1a@itts.com.tw']);
  m7.forEach((m) => {
    const kind = m.to[0] === 'mgr1a@itts.com.tw' ? 'mgr1' : 'gm';
    checkMail('E6 撤回', m, kind, 'E6_WITHDRAWN', { link: L(q7) });
    t('E6 撤回：主旨結尾「請勿簽核」、結果文字「已撤回」、不放決策條', /【簽核撤回】.* 請勿簽核$/.test(m.subject) && m.html.indexOf('已撤回') >= 0 && m.text.indexOf('已撤回') >= 0);
  });

  // 撤回（董事會關）：秘書與代核人收到的撤回信也不帶客戶名與業務名
  const q8 = newQuote(w);
  await submitQuote(w, q8, 'sales1', 'BOARD');
  await approveStep(w, q8, 'mgr1a');
  await approveStep(w, q8, 'gm1');
  n = w.sent.length;
  r = await withdrawQuote(w, q8, 'sales1');
  const m8 = slice(n);
  eq('E6 撤回（董事會關）：收件人＝4 位秘書／代核人＋已簽過的 mgr1a、gm1', m8.map((m) => m.to[0]).sort(), ['gm1@itts.com.tw', 'mgr1a@itts.com.tw', 'proxy1@itts.com.tw', 'sec1@itts.com.tw', 'sec2@itts.com.tw', 'sec3@itts.com.tw']);
  m8.forEach((m) => {
    const un = m.to[0].split('@')[0];
    const kind = un === 'mgr1a' ? 'mgr1' : (un === 'gm1' ? 'gm' : (un === 'proxy1' ? 'boardProxy' : 'secretary'));
    checkMail('E6 撤回（董事會）', m, kind, 'E6_WITHDRAWN', { link: L(q8) });
  });

  // 作廢（核准後修改）：已簽過的人＋業務（由管理員觸發時業務才會收到；業務自己改則是操作者，不寄）
  const q9 = newQuote(w);
  await submitQuote(w, q9, 'sales1', 'L1');
  await approveStep(w, q9, 'mgr1a');
  n = w.sent.length;
  r = await voidQuote(w, q9, 'admin1');
  const m9 = slice(n);
  eq('E6 作廢（管理員觸發）：收件人＝已簽過的一級主管＋業務', m9.map((m) => m.to[0]).sort(), ['mgr1a@itts.com.tw', 'sales1@itts.com.tw']);
  m9.forEach((m) => {
    checkMail('E6 作廢', m, m.to[0] === 'sales1@itts.com.tw' ? 'owner' : 'mgr1', 'E6_WITHDRAWN', { link: L(q9) });
    t('E6 作廢：結果文字「核准已作廢」', m.html.indexOf('核准已作廢') >= 0 && m.text.indexOf('核准已作廢') >= 0);
  });
  const q10 = newQuote(w);
  await submitQuote(w, q10, 'sales1', 'L1');
  await approveStep(w, q10, 'mgr1a');
  n = w.sent.length;
  await voidQuote(w, q10, 'sales1');
  eq('E6 作廢（業務自己改，業務是操作者）：業務不會收到自己觸發的信', w.sent.slice(n).map((m) => m.to[0]), ['mgr1a@itts.com.tw']);

  // ── 全局檢查 ──
  t('整段流程：每一封信都是單一收件人、通過傳輸層驗證、<100KB', w.sent.length > 30 && w.sent.every((m) => m.to.length === 1 && TRN.validateMessage(m).ok && Buffer.byteLength(m.html + m.text, 'utf8') < 100 * 1024), w.sent.length);
  t('整段流程：沒有寄給停用帳號 mgr1old、manager2 的 lead1、其他業務 sales2、管理員 admin1（只有作廢信的管理員是操作者）', w.sent.every((m) => !/^(mgr1old|lead1|sales2|admin1)@/.test(m.to[0])));
  recs = await w.records();
  t('outbox：所有紀錄 sent、attempts=1、無錯誤碼', recs.length === w.sent.length && recs.every((x) => x.status === 'sent' && x.attempts === 1 && x.lastErrorCode === null), recs.length + '/' + w.sent.length);
  eq('整段流程：沒有任何稽核（成功與正常流程不寫稽核）', w.logs, []);
  eq('整段流程：render 呼叫次數＝寄出封數', w.renderCount, w.sent.length);
});

// ═════════════════════════════════════════════════════════════════════════
section('2 去重與重複觸發', async () => {
  const w = mkWorld();
  const q = newQuote(w);
  await submitQuote(w, q, 'sales1', 'L2');
  const renders = w.renderCount;
  let r = await sendStep(w, q, 'E1_SUBMIT', 0, 'sales1');
  eq('同一事件重複觸發（例如重複存檔）：不重寄、記 DUPLICATE', [r.sent, r.skipped], [0, [{ username: 'mgr1a', reason: 'DUPLICATE' }]]);
  eq('重複觸發：沒有再渲染、傳輸只被呼叫 1 次、outbox 仍 1 筆、沒有稽核', [w.renderCount - renders, w.sent.length, (await w.records()).length, w.logs.length], [0, 1, 1, 0]);

  // 同時兩次（雙擊送出／兩個請求競爭）
  const q2 = newQuote(w);
  q2.approval = { state: 'pending', submittedAt: iso(T0), submittedBy: 'sales1', derived: DERIVED.L1, cur: 0, history: [], steps: [{ tier: 'mgr1', label: STEP_LABEL.mgr1, assignee: 'mgr1a', status: 'pending' }] };
  const [a, b] = await Promise.all([sendStep(w, q2, 'E1_SUBMIT', 0, 'sales1'), sendStep(w, q2, 'E1_SUBMIT', 0, 'sales1')]);
  eq('同時兩次相同事件：恰好寄出 1 封（另一個是 DUPLICATE）', [a.sent + b.sent, a.skipped.length + b.skipped.length, w.sent.filter((m) => m.subject.indexOf(q2.quoteNo) >= 0).length], [1, 1, 1]);

  // 駁回後重新送簽：新的 submittedAt → 新事件 → 要再寄
  const q3 = newQuote(w);
  await submitQuote(w, q3, 'sales1', 'L1');
  await rejectStep(w, q3, 'mgr1a', '請修正');
  const before = w.sent.length;
  w.clock.advance(3600 * SEC);
  q3.approval.state = 'none';
  r = await submitQuote(w, q3, 'sales1', 'L1');
  eq('駁回後重新送簽（stepKey 含新的送簽時間）：一級主管會再收到一封', [r.e1.sent, w.sent.length - before, w.sent[before].to[0]], [1, 1, 'mgr1a@itts.com.tw']);
  eq('同一單的兩次 E1 都有紀錄（不同去重鍵）', (await w.records()).filter((x) => x.quoteNo === q3.quoteNo && x.type === 'E1_SUBMIT').length, 2);

  // 不同事件型別／不同收件人／不同單據互不影響
  const q4 = newQuote(w);
  await submitQuote(w, q4, 'sales1', 'L2');
  await approveStep(w, q4, 'mgr1a');
  const cnt = w.sent.length;
  r = await sendStep(w, q4, 'E3_NEXT_STEP', 1, 'mgr1a');
  eq('E3 對兩位總經理重複觸發：兩位都是 DUPLICATE，不重寄', [r.sent, r.skipped.map((x) => x.reason), w.sent.length - cnt], [0, ['DUPLICATE', 'DUPLICATE'], 0]);
  t('去重鍵格式：type:quoteId:username:stepKey，帳號含冒號會跳脫', EVT.dedupeKey({ type: 'E1_SUBMIT', quoteId: q4.id, stepKey: 'k#0' }, 'a:b') === 'E1_SUBMIT:' + q4.id + ':a%3Ab:k#0');
});

// ═════════════════════════════════════════════════════════════════════════
function degradedWorld(extra) {
  const users = mkUsers();
  const U = (nm) => users.find((u) => u.username === nm);
  delete U('gm2').email;                                                                                     // 缺 email
  users.push({ username: 'gm3', role: 'executive', displayName: 'GM3 Name', email: 'gm3@itts.com.tw', active: false });   // 停用
  users.push({ username: 'gm4', role: 'executive', displayName: 'GM4 Name', email: 'gm4@example.test' });    // 網域不在白名單
  users.push({ username: 'gm5', role: 'executive', displayName: 'GM5 Name', email: 'not-an-email' });        // 格式非法
  U('chair1').email = 'chair1@itts.com.tw.evil.test';                                                        // 後綴攻擊的網域
  U('proxy1').email = 'proxy1@example.test';
  const w = mkWorld(Object.assign({ users, glueFilter: false }, extra || {}));
  w.roster.gm = ['gm1', 'gm2', 'gm3', 'gm4', 'gm5', 'ghost', 'gm1'];                                         // ghost＝名冊裡的帳號已不存在；gm1 重複
  return w;
}
const reasonMap = (summary) => { const o = {}; summary.skipped.forEach((x) => { o[x.username] = x.reason; }); return o; };

section('3 缺 email、停用、網域不合法、操作者與重複帳號', async () => {
  const w = degradedWorld();
  const q = newQuote(w);
  await submitQuote(w, q, 'sales1', 'L3');
  eq('E1 一級主管正常寄出', w.emails(), ['mgr1a@itts.com.tw']);
  const r = await approveStep(w, q, 'mgr1a');
  eq('E3 總經理關：只有 gm1 寄出（gm1 在名冊重複兩次只寄一封）', [r.e3.sent, w.emails().filter((e) => e.startsWith('gm'))], [1, ['gm1@itts.com.tw']]);
  eq('E3 總經理關：各帳號的略過原因', reasonMap(r.e3), { gm2: 'NO_EMAIL', gm3: 'INACTIVE', gm4: 'DOMAIN_NOT_ALLOWED', gm5: 'BAD_EMAIL', ghost: 'UNKNOWN_USER', gm1: 'DUP' });
  const recs = await w.records();
  eq('outbox：略過的帳號都有 skipped 紀錄與原因（DUP 不入紀錄）', recs.filter((x) => x.status === 'skipped').map((x) => x.toUser + ':' + x.skipReason).sort(), ['ghost:UNKNOWN_USER', 'gm2:NO_EMAIL', 'gm3:INACTIVE', 'gm4:DOMAIN_NOT_ALLOWED', 'gm5:BAD_EMAIL']);
  eq('outbox：略過的紀錄沒有 email（遮罩欄位為空）', recs.filter((x) => x.status === 'skipped').map((x) => x.toMasked), ['', '', '', '', '']);
  eq('稽核：5 筆 QUOTE_MAIL_SKIPPED，內容固定格式、不含位址', w.logs.map((l) => l.slice(0, 1).concat(l.slice(3))).sort(), [
    ['QUOTE_MAIL_SKIPPED', 'type=E3_NEXT_STEP to=ghost reason=UNKNOWN_USER'], ['QUOTE_MAIL_SKIPPED', 'type=E3_NEXT_STEP to=gm2 reason=NO_EMAIL'],
    ['QUOTE_MAIL_SKIPPED', 'type=E3_NEXT_STEP to=gm3 reason=INACTIVE'], ['QUOTE_MAIL_SKIPPED', 'type=E3_NEXT_STEP to=gm4 reason=DOMAIN_NOT_ALLOWED'],
    ['QUOTE_MAIL_SKIPPED', 'type=E3_NEXT_STEP to=gm5 reason=BAD_EMAIL'],
  ].sort());
  eq('稽核：操作者標籤與單號正確傳入 writeLog', w.logs.map((l) => [l[1], l[2]]).filter((x, i, a) => a.findIndex((y) => y[0] === x[0] && y[1] === x[1]) === i), [['Mgr Nick', q.quoteNo]]);
  t('簽核本身不受影響：狀態已進到總經理關（業務動作照常成功）', q.approval.state === 'pending' && q.approval.cur === 1);

  // 董事長關：email 的網域是後綴攻擊 → 沒有任何人收到信，但簽核照常往前
  await approveStep(w, q, 'gm1');
  const cr = (await w.records()).find((x) => x.toUser === 'chair1' && x.type === 'E3_NEXT_STEP');
  eq('董事長關：chair1 的 email 是 itts.com.tw.evil.test → DOMAIN_NOT_ALLOWED，沒有寄出', [cr && cr.status, cr && cr.skipReason, w.emails().filter((e) => e.startsWith('chair'))], ['skipped', 'DOMAIN_NOT_ALLOWED', []]);
  const logsBefore = w.logs.length;
  const fin = await approveStep(w, q, 'chair1');
  const lastMail = w.sent[w.sent.length - 1];
  eq('董事長（即使自己收不到信）照常核准後，業務仍收到最終結果信', [fin.e4.sent, lastMail.to[0], lastMail.tag, lastMail.html.indexOf('已完成核准') >= 0], [1, 'sales1@itts.com.tw', 'E4_RESULT', true]);
  t('缺 email 造成的略過不阻擋後續事件：最終核准沒有新增稽核', w.logs.length === logsBefore);

  // 董事會關：proxy1 網域不合法，3 位秘書正常
  const q2 = newQuote(w);
  await submitQuote(w, q2, 'sales1', 'BOARD');
  await approveStep(w, q2, 'mgr1a');
  const n = w.sent.length;
  // 總經理關的 E3：只有 gm1 有效，由 gm1 簽
  const rb = await approveStep(w, q2, 'gm1');
  eq('董事會關：3 位秘書收到、proxy1 因網域不合法略過（4 位中的 3 位）', [rb.e3.sent, reasonMap(rb.e3), w.sent.slice(n).filter((m) => m.tag === 'E3_NEXT_STEP').map((m) => m.to[0]).sort()], [3, { proxy1: 'DOMAIN_NOT_ALLOWED' }, ['sec1@itts.com.tw', 'sec2@itts.com.tw', 'sec3@itts.com.tw']]);
  t('董事會關：沒有寄給網域不合法的位址', w.sent.every((m) => !/example\.test|evil\.test/.test(m.to[0])));

  // 管理員的「缺 email」清單與派送時的略過原因一致（P1 與 P2 的跨模組對拍）
  const rep = UE.missingEmailReport({ users: w.users, roster: w.roster, config: w.config });
  eq('missingEmailReport：停用的 gm3、不存在的 ghost 不列；其餘缺口依帳號排序', rep.map((x) => [x.username, x.reason, x.roles]), [
    ['chair1', 'DOMAIN_NOT_ALLOWED', ['chairman']], ['gm2', 'NO_EMAIL', ['gm']], ['gm4', 'DOMAIN_NOT_ALLOWED', ['gm']], ['gm5', 'BAD_EMAIL', ['gm']], ['proxy1', 'DOMAIN_NOT_ALLOWED', ['boardProxy']],
  ]);
  const skipByUser = {};
  (await w.records()).filter((x) => x.status === 'skipped').forEach((x) => { skipByUser[x.toUser] = x.skipReason; });
  t('跨模組一致：缺口清單的每一筆原因＝派送當時該帳號被略過的原因', rep.every((x) => skipByUser[x.username] === x.reason), short(skipByUser));
});

section('3b 操作者、停用、無 email 的其他情境', async () => {
  // 膠水有過濾（notify 同款）：停用與不存在的帳號根本到不了信件層
  {
    const w = degradedWorld();
    w.glueFilter = true;
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L2');
    const r = await approveStep(w, q, 'mgr1a');
    eq('膠水先過濾：停用的 gm3、不存在的 ghost、重複的 gm1 都不會進信件層（只剩缺 email／網域／格式問題）', reasonMap(r.e3), { gm2: 'NO_EMAIL', gm4: 'DOMAIN_NOT_ALLOWED', gm5: 'BAD_EMAIL' });
    t('膠水先過濾：outbox 沒有 gm3、ghost 的紀錄', (await w.records()).every((x) => x.toUser !== 'gm3' && x.toUser !== 'ghost'));
  }
  // 膠水沒有過濾：操作者本人在名單內 → ACTOR 略過且不入 outbox、不寫稽核
  {
    const w = mkWorld({ glueFilter: false });
    w.roster.gm = ['gm1', 'gm2'];
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L2');
    await approveStep(w, q, 'mgr1a');
    const r = await sendStep(w, q, 'E3_NEXT_STEP', 1, 'gm1');
    eq('操作者 gm1 自己在收件名單：ACTOR 略過、其餘照常（gm2 已寄過所以是 DUPLICATE）', reasonMap(r), { gm1: 'ACTOR', gm2: 'DUPLICATE' });
    t('ACTOR 不入 outbox（gm1 在 E3 只有第一次的那一筆）、不寫稽核', (await w.records()).filter((x) => x.toUser === 'gm1' && x.type === 'E3_NEXT_STEP').length === 1 && w.logs.length === 0);
  }
  // 顧問沒有 email：E2 略過、簽核與存檔不受影響
  {
    const users = mkUsers();
    delete users.find((u) => u.username === 'cons1').email;
    const w = mkWorld({ users });
    const q = newQuote(w);
    const r = await requestCost(w, q, 'sales1');
    eq('E2 顧問缺 email：略過 NO_EMAIL、沒寄、寫稽核', [r.e2.sent, reasonMap(r.e2), w.sent.length, w.logs.map((l) => l[0] + ' ' + l[3])], [0, { cons1: 'NO_EMAIL' }, 0, ['QUOTE_MAIL_SKIPPED type=E2_COST_REQUEST to=cons1 reason=NO_EMAIL']]);
    t('E2 顧問缺 email：成本流程狀態仍是 requested（業務動作成功）', q.costFlow.state === 'requested');
    const rep = UE.missingEmailReport({ users: w.users, roster: w.roster, config: w.config });
    eq('missingEmailReport：顧問列為 costProvider 缺口', rep.map((x) => [x.username, x.roles, x.reason]), [['cons1', ['costProvider'], 'NO_EMAIL']]);
  }
  // 寄送當下再檢查：排隊後帳號被停用／移除 email／改成外部網域 → 不寄
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT', message: 'slow' });
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L2');
    await approveStep(w, q, 'mgr1a');                                                 // gm1、gm2 的 E3 都排入重試
    const pend = (await w.records()).filter((x) => x.status === 'pending');
    eq('前置：一級主管的 E1 與兩位總經理的 E3、業務的 E4 全都因傳輸逾時而排入重試', pend.map((x) => x.toUser).sort(), ['gm1', 'gm2', 'mgr1a', 'sales1']);
    w.behavior.fn = null;
    w.users.find((u) => u.username === 'gm1').active = false;
    delete w.users.find((u) => u.username === 'gm2').email;
    w.users.find((u) => u.username === 'sales1').email = 'sales1@example.test';
    w.clock.advance(61 * SEC);
    const sentBefore = w.sent.length;
    const d = await w.drainDue({ limit: 10 });
    eq('重試當下重新檢查：gm1 停用、gm2 無 email、sales1 網域改了 → 都不寄；mgr1a（E1）因已進到別關而被取消', [d.skipped.map((x) => x.username + ':' + x.reason).sort(), d.sent, d.cancelled], [['gm1:INACTIVE', 'gm2:NO_EMAIL', 'sales1:DOMAIN_NOT_ALLOWED'], 0, 1]);
    t('重試當下略過：傳輸層沒有收到任何新訊息', w.sent.length === sentBefore);
    eq('重試當下略過：紀錄改標 skipped（原因與稽核一致）', (await w.records()).filter((x) => x.status === 'skipped').map((x) => x.toUser + ':' + x.skipReason).sort(), ['gm1:INACTIVE', 'gm2:NO_EMAIL', 'sales1:DOMAIN_NOT_ALLOWED']);
    eq('重試當下略過：寫 3 筆 QUOTE_MAIL_SKIPPED', w.logs.filter((l) => l[0] === 'QUOTE_MAIL_SKIPPED').length, 3);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('4 寄信模式 off／log／redirect／live', async () => {
  const REAL = ['mgr1a', 'sales1', 'gm1', 'gm2', 'sales1', 'sec1', 'sec2', 'sec3', 'proxy1'];     // 完整流程會寄的 9 位收件人（業務出現兩次：兩個 E4）
  async function fullFlow(w) {
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'BOARD');
    await approveStep(w, q, 'mgr1a');
    await approveStep(w, q, 'gm1');
    return q;
  }
  const statusCount = (recs) => { const o = {}; recs.forEach((x) => { const k = x.status + ':' + (x.skipReason || ''); o[k] = (o[k] || 0) + 1; }); return o; };
  const previewFiles = (w) => (fs.existsSync(path.join(w.dir, 'prev')) ? fs.readdirSync(path.join(w.dir, 'prev')) : []);

  // off
  {
    const w = mkWorld({ env: { MAIL_MODE: 'off' } });
    await fullFlow(w);
    eq('off：傳輸層沒收到任何信、沒有渲染、沒有查帳號資料', [w.sent.length, w.renderCount, w.getUsersCalls], [0, 0, 0]);
    eq('off：outbox 9 筆 skipped/MODE_OFF（簽核流程照常進行）', statusCount(await w.records()), { 'skipped:MODE_OFF': 9 });
    eq('off：沒有預覽檔、沒有稽核', [previewFiles(w), w.logs], [[], []]);
    const d = await w.drainDue({});
    eq('off：drainDue 什麼都不做', [d.sent, d.errors, w.sent.length], [0, 0, 0]);
  }
  // 沒設 MAIL_MODE＝預設 off
  {
    const w = mkWorld({ env: { MAIL_MODE: undefined } });
    t('未設 MAIL_MODE：有效模式是 off', w.config.mode === 'off');
    await fullFlow(w);
    eq('未設 MAIL_MODE：一封都不寄', [w.sent.length, w.renderCount], [0, 0]);
  }
  // log
  {
    const w = mkWorld({ env: { MAIL_MODE: 'log' } });
    await fullFlow(w);
    eq('log：真傳輸一次都沒被呼叫', w.sent.length, 0);
    eq('log：outbox 9 筆 skipped/MODE_LOG（可由後台 requeue）', statusCount(await w.records()), { 'skipped:MODE_LOG': 9 });
    const files = previewFiles(w);
    t('log：預覽目錄寫出 9 封信×3 個檔（暫存目錄內）', files.length === 27 && files.every((f) => /^[A-Za-z0-9_-]+\.(json|html|txt)$/.test(f)), files.length);
    const jsons = files.filter((f) => /\.json$/.test(f)).map((f) => JSON.parse(fs.readFileSync(path.join(w.dir, 'prev', f), 'utf8')));
    eq('log：預覽 json 記錄的是真實收件位址（本機對照用，目錄已被 .gitignore）', jsons.map((j) => j.to[0]).sort(), REAL.map((u) => u + '@itts.com.tw').sort());
    const secHtml = files.filter((f) => /E3_NEXT_STEP.*\.html$/.test(f)).map((f) => fs.readFileSync(path.join(w.dir, 'prev', f), 'utf8'));
    t('log：預覽的 E3 信件仍依角色過濾（4 封董事會信沒有客戶名）', secHtml.length === 6 && secHtml.filter((h) => h.indexOf(CUSTOMER) < 0).length === 4, secHtml.length);
    eq('log：沒有稽核', w.logs, []);
    // 之後切到 redirect 並由後台重送。單據已經走到董事會關，所以挑「目前這一關」的紀錄（舊關卡的紀錄重送會被判定過期而取消，見第 6 章）
    const rec = (await w.records()).find((x) => x.toUser === 'sec1' && x.type === 'E3_NEXT_STEP');
    w.config.mode = 'redirect'; w.config.redirectTo = 'redirect.box@example.test';
    const rq = await w.outbox.requeue(rec.id);
    const d = await w.drainDue({});
    eq('log 紀錄切到 redirect 後可由後台重送（requeue→drainDue），寄到測試信箱', [rq.ok, d.sent, w.sent.map((m) => m.to)], [true, 1, [['redirect.box@example.test']]]);
    const stale = (await w.records()).find((x) => x.toUser === 'mgr1a' && x.type === 'E1_SUBMIT');
    await w.outbox.requeue(stale.id);
    const d2 = await w.drainDue({});
    eq('log 紀錄重送時單據已走到後面的關卡：舊關卡的「請簽核」信被取消（STALE），不會誤叫一級主管再簽', [d2.sent, d2.cancelled, (await w.records()).find((x) => x.id === stale.id).skipReason], [0, 1, 'STALE']);
  }
  // redirect
  {
    const w = mkWorld({ env: { MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'Redirect.Box@Example.TEST' } });
    t('redirect：config 有效模式 redirect、redirectTo 已正規化小寫', w.config.mode === 'redirect' && w.config.redirectTo === 'redirect.box@example.test');
    await fullFlow(w);
    eq('redirect：9 封全部寄出', [w.sent.length, w.renderCount], [9, 9]);
    t('redirect：每一封的收件人只有 MAIL_REDIRECT_TO', w.sent.every((m) => m.to.length === 1 && m.to[0] === 'redirect.box@example.test'), short(w.emails()));
    t('redirect：主旨前綴 [測試轉送]', w.sent.every((m) => m.subject.indexOf('[測試轉送] 【') === 0));
    const masks = w.sent.map((m) => {
      const h = /原收件人：([^（<]+)（redirect 模式/.exec(m.html);
      const x = /^【測試轉送】原收件人：([^（]+)（redirect 模式/.exec(m.text);
      return h && x && h[1] === x[1] ? h[1] : null;
    });
    eq('redirect：html 與純文字最上方都加註遮罩後的原收件人，且與預定收件人一一對應', masks.slice().sort(), REAL.map((u) => SAF.maskEmail(u + '@itts.com.tw')).sort());
    t('redirect：html 橫幅位於 <body> 之後第一個元素', w.sent.every((m) => /<body[^>]*><div [^>]*>【測試轉送】/.test(m.html)));
    t('redirect：傳輸層收到的任何內容都不含預定收件人的完整位址', REAL.every((u) => JSON.stringify(w.sent).indexOf(u + '@itts.com.tw') < 0));
    const boardMails = w.sent.filter((m) => m.tag === 'E3_NEXT_STEP' && m.text.indexOf('董事會決議（秘書代核）') >= 0);
    t('redirect：信件內容仍依「真實收件人」的角色過濾（4 封董事會信沒有客戶名與業務名；E1／總經理關的信有）', boardMails.length === 4 && boardMails.every((m) => blobOf(m).indexOf(CUSTOMER) < 0 && blobOf(m).indexOf(OWNER_LABEL) < 0) && w.sent.filter((m) => m.tag === 'E1_SUBMIT' || (m.tag === 'E3_NEXT_STEP' && boardMails.indexOf(m) < 0)).every((m) => blobOf(m).indexOf(CUSTOMER) >= 0), boardMails.length);
    const rec = await w.records();
    t('redirect：outbox 記的是預定收件人的遮罩位址，不含測試信箱', rec.every((x) => /^[a-z]\*\*\*@itts\.com\.tw$/.test(x.toMasked)) && fs.readFileSync(w.file, 'utf8').indexOf('redirect.box') < 0);
    t('redirect：稽核沒有出現測試信箱', w.logs.every((l) => JSON.stringify(l).indexOf('redirect.box') < 0));
  }
  // redirect 沒設目標 → 降為 log
  {
    const w = mkWorld({ env: { MAIL_MODE: 'redirect' } });
    t('redirect 沒設 MAIL_REDIRECT_TO：有效模式降為 log、warnings 有說明', w.config.mode === 'log' && w.config.warnings.some((x) => /MAIL_REDIRECT_TO/.test(x)));
    await fullFlow(w);
    eq('redirect 沒目標：不會寄給任何人（也不會退回寄給原收件人）', [w.sent.length, statusCount(await w.records())], [0, { 'skipped:MODE_LOG': 9 }]);
  }
  // 手工破壞 config：mode=redirect 但 redirectTo 被清空 → NOT_CONFIGURED，且不呼叫真傳輸
  {
    const w = mkWorld({ env: { MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'redirect.box@example.test' } });
    w.config.redirectTo = '';
    const q = newQuote(w);
    const r = await submitQuote(w, q, 'sales1', 'L1');
    eq('config 被破壞（redirect 且無目標）：失敗 NOT_CONFIGURED、真傳輸未被呼叫', [r.e1.failed, w.sent.length], [[{ username: 'mgr1a', code: 'NOT_CONFIGURED' }], 0]);
  }
  // live
  {
    const w = mkWorld({ env: { MAIL_MODE: 'live' } });
    await fullFlow(w);
    eq('live：9 封寄到真實位址', w.emails(), REAL.map((u) => u + '@itts.com.tw').sort());
    t('live：沒有測試轉送橫幅或前綴', w.sent.every((m) => m.subject.indexOf('測試轉送') < 0 && blobOf(m).indexOf('測試轉送') < 0));
    eq('live：outbox 9 筆 sent', statusCount(await w.records()), { 'sent:': 9 });
  }
  // MAIL_MODE 變體：只有去空白、不分大小寫後恰為 live 才會真寄
  {
    const cases = [['LIVE ', 'live'], [' live', 'live'], ['Live', 'live'], ['LiVe\t', 'live'], ['true', 'off'], ['1', 'off'], ['yes', 'off'], ['livee', 'off'], ['live!', 'off'], ['', 'off'], ['on', 'off'], ['Live' + cp(0xa0), 'off'], [cp(0xff2c, 0xff29, 0xff36, 0xff25), 'off'], ['liv' + cp(0x65, 0x301), 'off']];
    for (const [raw, want] of cases) {
      const w = mkWorld({ env: { MAIL_MODE: raw } });
      const q = newQuote(w);
      await submitQuote(w, q, 'sales1', 'L1');
      const gotLive = w.sent.length === 1 && w.sent[0].to[0] === 'mgr1a@itts.com.tw';
      t('MAIL_MODE=' + short(raw) + ' → ' + want + '（' + (want === 'live' ? '真的寄出' : '一封都不寄') + '）', want === 'live' ? gotLive : w.sent.length === 0, w.config.mode + '/' + w.sent.length);
    }
  }
  // 執行中切換模式
  {
    const w = mkWorld({ env: { MAIL_MODE: 'live' } });
    const q1 = newQuote(w);
    await submitQuote(w, q1, 'sales1', 'L1');
    w.config.mode = 'off';
    const q2 = newQuote(w);
    await submitQuote(w, q2, 'sales1', 'L1');
    w.config.mode = 'redirect'; w.config.redirectTo = 'redirect.box@example.test';
    const q3 = newQuote(w);
    await submitQuote(w, q3, 'sales1', 'L1');
    w.config.mode = 'LIVE ';
    const q4 = newQuote(w);
    await submitQuote(w, q4, 'sales1', 'L1');
    w.config.mode = 'garbage';
    const q5 = newQuote(w);
    await submitQuote(w, q5, 'sales1', 'L1');
    eq('執行中切換 live→off→redirect→「LIVE 」→亂值：寄到 [真實, 不寄, 測試信箱, 真實, 不寄]', w.sent.map((m) => m.to[0]), ['mgr1a@itts.com.tw', 'redirect.box@example.test', 'mgr1a@itts.com.tw']);
  }
  // 選傳輸的護欄（selectTransport）：log／redirect 模式即使傳入真傳輸，行為也受模式約束
  {
    const real = { kind: 'REAL', async send() { return { ok: true }; } };
    t('selectTransport：live 才回傳真傳輸', TRN.selectTransport({ mode: 'live' }, { realTransport: real }) === real && ['off', 'log', 'redirect', 'x', undefined, null, 5].every((m) => TRN.selectTransport({ mode: m, redirectTo: 'a@example.test' }, { realTransport: real }) !== real));
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('5 失敗注入與恢復（業務動作必須照常成功）', async () => {
  const oneMail = async (w, over) => {
    const q = newQuote(w, over);
    const r = await submitQuote(w, q, 'sales1', 'L1');
    return { q, r: r.e1 };
  };
  const recOf = async (w, user) => (await w.records()).find((x) => x.toUser === user);

  // TIMEOUT：傳輸永不 resolve
  {
    const w = mkWorld({ timeouts: { connectMs: 20, totalMs: 100 } });
    w.behavior.fn = () => new Promise(() => { /* never */ });
    const t0 = Date.now();
    const { q, r } = await oneMail(w);
    const elapsed = Date.now() - t0;
    eq('TIMEOUT（永不 resolve）：dispatch 返回 queued=1，沒有 failed', [r.queued, r.sent, r.failed, r.errors], [1, 0, [], 0]);
    t('TIMEOUT：在總逾時內返回（不被傳輸拖住）', elapsed < 3000, elapsed + 'ms');
    t('TIMEOUT：已對傳輸送出 abort 訊號', w.signals.length === 1 && w.signals[0] && w.signals[0].aborted === true);
    t('TIMEOUT：簽核狀態已進入 pending（業務動作成功）', q.approval.state === 'pending');
    const rec = await recOf(w, 'mgr1a');
    eq('TIMEOUT：紀錄 pending、已嘗試 1 次、錯誤碼 TIMEOUT、60 秒後重試', [rec.status, rec.attempts, rec.lastErrorCode, rec.nextAttemptAt], ['pending', 1, 'TIMEOUT', iso(T0 + 60 * SEC)]);
    eq('TIMEOUT：還沒到期，drainDue 不會領取', [(await w.drainDue({})).sent, w.sent.length], [0, 1]);
    w.behavior.fn = null;
    w.clock.advance(60 * SEC);
    const d = await w.drainDue({});
    const rec2 = await recOf(w, 'mgr1a');
    eq('TIMEOUT 恢復：到期後 drainDue 重寄成功，紀錄 sent、attempts=2', [d.sent, rec2.status, rec2.attempts], [1, 'sent', 2]);
    t('TIMEOUT 恢復：重試的信與首次嘗試逐字相同（主旨、html、純文字）', w.sent.length === 2 && w.sent[0].subject === w.sent[1].subject && w.sent[0].html === w.sent[1].html && w.sent[0].text === w.sent[1].text);
    eq('TIMEOUT 全程不寫稽核（中途失敗不算最終失敗）', w.logs, []);
  }
  // 5xx 連續失敗：退避 60／300／900 → 第 4 次失敗＝最終失敗 → 稽核 → 後台重送
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER', message: 'upstream 503' });
    const { q } = await oneMail(w);
    const delays = [];
    let rec = await recOf(w, 'mgr1a');
    delays.push((Date.parse(rec.nextAttemptAt) - w.clock.t) / SEC);
    for (let i = 0; i < 3; i++) {
      w.clock.advance(Math.max(1, Date.parse(rec.nextAttemptAt) - w.clock.t));
      const d = await w.drainDue({});
      rec = await recOf(w, 'mgr1a');
      if (i < 2) { delays.push((Date.parse(rec.nextAttemptAt) - w.clock.t) / SEC); eq('5xx 第 ' + (i + 2) + ' 次嘗試仍失敗：還在重試', [d.queued, d.failed.length, rec.status], [1, 0, 'pending']); }
      else eq('5xx 第 4 次嘗試失敗＝本輪用盡：最終失敗', [d.failed, rec.status, rec.attempts, rec.lastErrorCode], [[{ username: 'mgr1a', code: 'SERVER' }], 'failed', 4, 'SERVER']);
    }
    eq('5xx：退避依序 60 → 300 → 900 秒（規格 config.retry.delaysSec）', delays, [60, 300, 900]);
    eq('5xx：最終失敗寫 QUOTE_MAIL_FAILED 稽核（型別、帳號、錯誤碼；不含位址）', w.logs.map((l) => [l[0], l[3]]), [['QUOTE_MAIL_FAILED', 'type=E1_SUBMIT to=mgr1a code=SERVER']]);
    t('5xx：業務流程不受影響（簽核仍 pending）', q.approval.state === 'pending');
    w.clock.advance(3600 * SEC);
    eq('最終失敗後不會自動再領取', [(await w.drainDue({})).sent, w.sent.length], [0, 4]);
    w.behavior.fn = null;
    const rq = await w.outbox.requeue(rec.id);
    const d2 = await w.drainDue({});
    rec = await recOf(w, 'mgr1a');
    eq('後台「立即重送」：requeue 後寄出，紀錄 sent、attempts 累計 5', [rq.ok, d2.sent, rec.status, rec.attempts], [true, 1, 'sent', 5]);
  }
  // 第 3 次重試的退避是 900 秒
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'NETWORK', message: 'reset' });
    await oneMail(w);
    const steps = [];
    for (let i = 0; i < 3; i++) {
      const rec = await recOf(w, 'mgr1a');
      w.clock.advance(Date.parse(rec.nextAttemptAt) - w.clock.t);
      await w.drainDue({});
      const r2 = await recOf(w, 'mgr1a');
      if (r2.nextAttemptAt) steps.push((Date.parse(r2.nextAttemptAt) - w.clock.t) / SEC);
    }
    eq('NETWORK 連續失敗：退避依序 60（首次）→ 300 → 900 秒，之後最終失敗', [steps, (await recOf(w, 'mgr1a')).status], [[300, 900], 'failed']);
  }
  // THROTTLED + Retry-After
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'THROTTLED', retryAfterSec: 120, message: '429' });
    const { r } = await oneMail(w);
    const rec = await recOf(w, 'mgr1a');
    eq('THROTTLED（Retry-After 120）：排在 120 秒後（優先於預設 60 秒）', [r.queued, rec.lastErrorCode, rec.nextAttemptAt], [1, 'THROTTLED', iso(T0 + 120 * SEC)]);
    w.behavior.fn = null;
    w.clock.advance(119 * SEC);
    eq('THROTTLED：119 秒時還不會領取', (await w.drainDue({})).sent, 0);
    w.clock.advance(1 * SEC);
    eq('THROTTLED：120 秒時寄出', (await w.drainDue({})).sent, 1);
    t('THROTTLED：熔斷未開（權重 3 < 門檻 5），成功後重置', (await w.outbox.breaker.state()).open === false);
  }
  // 傳輸拋例外／reject／亂回傳：包成可重試失敗，dispatch 不 throw
  {
    const kinds = {
      '同步拋例外': () => { throw new Error('socket hang up'); },
      'reject': () => Promise.reject(new Error('ECONNRESET')),
      '回傳 undefined': () => undefined,
      '回傳字串': () => 'ok',
      '回傳 ok:"yes"': () => ({ ok: 'yes' }),
      '未知錯誤碼': () => ({ ok: false, code: 'WAT' }),
    };
    for (const k of Object.keys(kinds)) {
      const w = mkWorld();
      w.behavior.fn = kinds[k];
      const { q, r } = await oneMail(w);
      const rec = await recOf(w, 'mgr1a');
      t('傳輸' + k + '：dispatch 返回 queued=1、紀錄 pending 且可重試、業務動作成功', r.queued === 1 && r.failed.length === 0 && rec.status === 'pending' && q.approval.state === 'pending' && ['NETWORK', 'SERVER'].indexOf(rec.lastErrorCode) >= 0, short(rec));
    }
  }
  // permanent：REJECTED、BAD_MESSAGE → 一次就最終失敗，不影響熔斷
  {
    for (const code of ['REJECTED', 'BAD_MESSAGE']) {
      const w = mkWorld();
      w.behavior.fn = async () => ({ ok: false, code, message: 'no' });
      const { r } = await oneMail(w);
      const rec = await recOf(w, 'mgr1a');
      eq(code + '：permanent，直接最終失敗＋稽核；熔斷不受影響', [r.failed, rec.status, rec.attempts, w.logs.map((l) => l[3]), (await w.outbox.breaker.state()).open], [[{ username: 'mgr1a', code }], 'failed', 1, ['type=E1_SUBMIT to=mgr1a code=' + code], false]);
    }
  }
  // 錯誤訊息含 email／權杖時，outbox 與稽核都看不到
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER', message: 'mailbox mgr1a@itts.com.tw unavailable Bearer abcdefghijklmnopqrstuvwxyz0123456789 client_secret=supersecretvalue' });
    await oneMail(w);
    const rec = await recOf(w, 'mgr1a');
    t('錯誤訊息清理：lastErrorMsg 不含 email 全址、Bearer 權杖、client_secret 值', !/mgr1a@itts/.test(rec.lastErrorMsg) && !/abcdefghijklmnopqrstuvwxyz/.test(rec.lastErrorMsg) && !/supersecretvalue/.test(rec.lastErrorMsg), rec.lastErrorMsg);
  }
  // render 失敗（膠水把事件寄給不該收的角色）：只影響該收件人。
  // 這裡 isStillValid 一律放行；若用真的 isStillValid，非當前關卡的收件人會先被判定過期（STALE）而取消，根本走不到 render
  {
    const w = mkWorld({ isStillValid: () => true });
    const q = newQuote(w);
    q.approval = { state: 'pending', submittedAt: iso(T0), submittedBy: 'sales1', derived: DERIVED.L1, cur: 0, history: [], steps: [{ tier: 'mgr1', label: '一級主管', assignee: 'mgr1a', status: 'pending' }] };
    const ev = buildEvent(w, q, 'E1_SUBMIT', { stepKey: keyStep(q, 0), step: stepOfStep(q.approval.steps[0]), actor: 'sales1' });
    const r = await w.dispatch(ev, [{ username: 'cons1', kind: 'consultant' }, { username: 'mgr1a', kind: 'mgr1' }, { username: 'sec1', kind: 'nonsense' }], { actorUsername: 'sales1' });
    eq('膠水給錯 kind（E1 寄給顧問、未知 kind）：只有這兩位 RENDER 失敗，簽核人照常寄出', [r.sent, r.failed.map((x) => x.username + ':' + x.code).sort(), w.emails()], [1, ['cons1:RENDER', 'sec1:RENDER'], ['mgr1a@itts.com.tw']]);
    eq('膠水給錯 kind：稽核記 QUOTE_MAIL_FAILED code=RENDER（不含內容）', w.logs.map((l) => l[3]).sort(), ['type=E1_SUBMIT to=cons1 code=RENDER', 'type=E1_SUBMIT to=sec1 code=RENDER']);
    t('膠水給錯 kind：RENDER 失敗不計入熔斷', (await w.outbox.breaker.state()).failures === 0);
  }
  // 事件不合法（例如 stepKey 空白）：不寄、不 throw、記 BAD_EVENT
  {
    const w = mkWorld();
    const q = newQuote(w);
    q.approval = { state: 'pending', submittedAt: iso(T0), derived: DERIVED.L1, cur: 0, steps: [{ tier: 'mgr1', label: '一級主管', assignee: 'mgr1a' }] };
    const ev = buildEvent(w, q, 'E1_SUBMIT', { stepKey: '', step: { level: 1, label: '一級主管' }, actor: 'sales1' });
    const r = await w.dispatch(ev, [{ username: 'mgr1a', kind: 'mgr1' }], { actorUsername: 'sales1' });
    eq('stepKey 空白（膠水漏帶）：不寄、BAD_EVENT 失敗、寫稽核', [r.sent, r.failed, w.sent.length, w.logs.map((l) => l[3])], [0, [{ username: 'mgr1a', code: 'BAD_EVENT' }], 0, ['type=E1_SUBMIT code=BAD_EVENT']]);
    // 金額不是整數分（膠水誤傳元）
    const ev2 = buildEvent(w, q, 'E1_SUBMIT', { stepKey: 'k#0', step: { level: 1, label: '一級主管' }, actor: 'sales1' });
    ev2.numbers.revenueCents = 2000000.5;
    const r2 = await w.dispatch(ev2, [{ username: 'mgr1a', kind: 'mgr1' }], { actorUsername: 'sales1' });
    eq('revenueCents 不是整數（膠水誤傳元／小數）：不寄、BAD_EVENT', [r2.sent, r2.failed], [0, [{ username: 'mgr1a', code: 'BAD_EVENT' }]]);
  }
  // 相依爆炸：稽核函式丟例外、帳號資料讀取失敗
  {
    const w = mkWorld({ writeLog: () => { throw new Error('audit store down'); } });
    w.behavior.fn = async () => ({ ok: false, code: 'REJECTED' });
    const { q, r } = await oneMail(w);
    eq('稽核函式丟例外：不影響派送結果（仍回報 failed）、簽核照常', [r.failed.length, q.approval.state], [1, 'pending']);
    const w2 = mkWorld({ getUsers: () => { throw new Error('user store down'); } });
    const { q: q2, r: r2 } = await oneMail(w2);
    const rec2 = await recOf(w2, 'mgr1a');
    eq('帳號資料暫時讀不到：事件先入列（email 留空）等 drainDue 補寄，不丟事件、簽核照常', [r2.errors > 0, rec2 && rec2.status, rec2 && rec2.toMasked, w2.sent.length, q2.approval.state], [true, 'pending', '', 0, 'pending']);
  }
  // 儲存體爆炸（唯讀檔案系統）：dispatch 不 throw，errors>0，業務照常
  {
    const dir = tmpDir();
    const badFs = Object.assign({}, fs, { writeFileSync() { const e = new Error('EROFS'); e.code = 'EROFS'; throw e; }, mkdirSync() { const e = new Error('EROFS'); e.code = 'EROFS'; throw e; } });
    const w = mkWorld({ adapter: jsonFileAdapter({ file: path.join(dir, 'ro', 'mail-outbox.json'), fs: badFs }) });
    const q = newQuote(w);
    const r = await submitQuote(w, q, 'sales1', 'L1');
    const r2 = await approveStep(w, q, 'mgr1a');
    eq('outbox 儲存體唯讀：dispatch 不 throw、回報 errors>0、沒有寄出、簽核流程照常走完', [r.e1.errors > 0, r.e1.sent, r2.e4.errors > 0, q.approval.state, w.sent.length], [true, 0, true, 'approved', 0]);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('6 isStillValid、rebuild 與「同級已簽」', async () => {
  // 撤回：排隊中的 E1 在重試時被取消（rebuild 發現 steps 已清空）
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L1');
    w.behavior.fn = null;
    await withdrawQuote(w, q, 'sales1');                                      // mgr1a 的 E1 還在排隊；撤回的 E6 也寄給 mgr1a
    const pend = (await w.records()).filter((x) => x.status === 'pending');
    const sentBefore = w.sent.length;
    w.clock.advance(61 * SEC);
    const d = await w.drainDue({ limit: 10 });
    const e1 = (await w.records()).find((x) => x.type === 'E1_SUBMIT');
    eq('撤回後重試：排隊中的 E1 取消（GONE：steps 已清空），不寄過期的「請簽核」信', [pend.map((x) => x.type), d.cancelled, e1.status, e1.skipReason], [['E1_SUBMIT'], 1, 'cancelled', 'GONE']);
    t('撤回後重試：沒有再寄出任何信（撤回信 E6 早已寄出）', w.sent.length === sentBefore);
  }
  // 單據被刪除
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER' });
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L1');
    w.behavior.fn = null;
    w.quotes.delete(q.id);
    w.clock.advance(61 * SEC);
    const d = await w.drainDue({});
    eq('單據已刪除：重試時取消（GONE）', [d.cancelled, d.sent, (await w.records())[0].skipReason], [1, 0, 'GONE']);
  }
  // 同級另一人先簽：gm2 的 E3 在排隊，gm1 先簽了 → gm2 的信取消（STALE）
  {
    const w = mkWorld();
    w.behavior.fn = async (m) => (m.to[0] === 'gm2@itts.com.tw' ? { ok: false, code: 'TIMEOUT' } : { ok: true });
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L3');
    await approveStep(w, q, 'mgr1a');
    const g2 = (await w.records()).find((x) => x.toUser === 'gm2' && x.type === 'E3_NEXT_STEP');
    eq('前置：gm1 的 E3 已寄出、gm2 的 E3 因逾時排隊', [w.emails().indexOf('gm1@itts.com.tw') >= 0, g2.status], [true, 'pending']);
    await approveStep(w, q, 'gm1');                                            // gm1 簽了 → 進到董事長關（E3 給 chair1）
    w.behavior.fn = null;
    const sentBefore = w.sent.length;
    w.clock.advance(61 * SEC);
    const d = await w.drainDue({ limit: 10 });
    const g2b = (await w.records()).find((x) => x.id === g2.id);
    eq('同級已簽：gm2 的 E3 取消（STALE），不再叫他簽已經簽過的關', [d.cancelled, g2b.status, g2b.skipReason, w.sent.slice(sentBefore).map((m) => m.to[0])], [1, 'cancelled', 'STALE', []]);
  }
  // 已被駁回：排隊中的 E3 取消
  {
    const w = mkWorld();
    w.behavior.fn = async (m) => (m.to[0] === 'gm1@itts.com.tw' ? { ok: false, code: 'SERVER' } : { ok: true });
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L2');
    await approveStep(w, q, 'mgr1a');
    w.behavior.fn = null;
    q.approval.state = 'returned';                                              // gm2 駁回
    w.clock.advance(61 * SEC);
    const d = await w.drainDue({ limit: 10 });
    eq('單據已駁回：排隊中的 E3 取消（STALE）', [d.cancelled, d.sent], [1, 0]);
  }
  // E2：換了顧問或成本已填 → 舊顧問的請求信取消
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    const q = newQuote(w);
    await requestCost(w, q, 'sales1');
    w.behavior.fn = null;
    q.costBy = 'someone-else';
    w.clock.advance(61 * SEC);
    const d = await w.drainDue({});
    eq('E2：業務已換顧問 → 原顧問的請求信取消（STALE）', [d.cancelled, d.sent, w.sent.length], [1, 0, 1]);
  }
  // isStillValid 丟例外：不寄也不取消，稍後重試
  {
    let boom = true;
    const w = mkWorld({ isStillValid: () => { if (boom) throw new Error('lookup failed'); return true; } });
    const q = newQuote(w);
    const r = await submitQuote(w, q, 'sales1', 'L1');
    eq('isStillValid 丟例外：不寄、排入重試（STALE_CHECK），不取消', [r.e1.sent, r.e1.queued, r.e1.cancelled, w.sent.length, (await w.records())[0].lastErrorCode], [0, 1, 0, 0, 'STALE_CHECK']);
    boom = false;
    w.clock.advance(61 * SEC);
    eq('isStillValid 恢復後重試即寄出', [(await w.drainDue({})).sent, w.sent.length], [1, 1]);
  }
  // isStillValid 回傳 false：當次就取消（還沒寄）
  {
    const w = mkWorld({ isStillValid: () => false });
    const q = newQuote(w);
    const r = await submitQuote(w, q, 'sales1', 'L1');
    eq('isStillValid=false：當次取消、不渲染以外的動作、不寄', [r.e1.cancelled, r.e1.sent, w.sent.length, (await w.records())[0].skipReason], [1, 0, 0, 'STALE']);
  }
  // 只有回傳 true 才寄（"yes"、1 都不算）
  {
    for (const v of ['yes', 1, {}, undefined, null]) {
      const w = mkWorld({ isStillValid: () => v });
      const q = newQuote(w);
      await submitQuote(w, q, 'sales1', 'L1');
      t('isStillValid 回傳 ' + short(v) + '（不是 true）：不寄', w.sent.length === 0);
    }
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('7 熔斷', async () => {
  // AUTH：一次就開
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'AUTH', message: '401' });
    const q1 = newQuote(w);
    const r1 = await submitQuote(w, q1, 'sales1', 'L2');
    eq('AUTH（401／403）：permanent 最終失敗＋稽核', [r1.e1.failed, w.logs.map((l) => l[3])], [[{ username: 'mgr1a', code: 'AUTH' }], ['type=E1_SUBMIT to=mgr1a code=AUTH']]);
    const st = await w.outbox.breaker.state();
    eq('AUTH：熔斷立刻開啟 10 分鐘', [st.open, st.until], [true, iso(T0 + 600 * SEC)]);
    // 熔斷期間：下一個事件（E3 兩位總經理）只入列、不嘗試寄送
    const callsBefore = w.sent.length;
    w.behavior.fn = async () => ({ ok: true });
    const ap = q1.approval;
    ap.steps[0].status = 'approved'; ap.steps[0].by = 'mgr1a'; ap.cur = 1; ap.steps[1].status = 'pending';
    const r2 = await sendStep(w, q1, 'E3_NEXT_STEP', 1, 'mgr1a');
    eq('熔斷期間：E3 兩位收件人都只入列（queued=2），傳輸層沒被呼叫', [r2.queued, r2.sent, w.sent.length - callsBefore], [2, 0, 0]);
    const held = (await w.records()).filter((x) => x.type === 'E3_NEXT_STEP');
    eq('熔斷期間：排隊工作的下次嘗試時間＝熔斷結束時間', held.map((x) => [x.status, x.nextAttemptAt]), [['pending', iso(T0 + 600 * SEC)], ['pending', iso(T0 + 600 * SEC)]]);
    const d0 = await w.drainDue({});
    eq('熔斷期間：drainDue 直接返回 breakerOpen、不領取', [d0.breakerOpen, d0.sent, w.sent.length - callsBefore], [true, 0, 0]);
    w.clock.advance(601 * SEC);
    const d1 = await w.drainDue({});
    eq('熔斷結束後：drainDue 把排隊的兩封寄出', [d1.sent, w.sent.slice(callsBefore).map((m) => m.to[0]).sort()], [2, ['gm1@itts.com.tw', 'gm2@itts.com.tw']]);
    const st2 = await w.outbox.breaker.state();
    t('成功後熔斷重置（open=false、failures=0）', st2.open === false && st2.failures === 0, short(st2));
    t('AUTH 造成的最終失敗仍留在紀錄上，等後台處理（不會被自動重試）', (await w.records()).find((x) => x.type === 'E1_SUBMIT' && x.toUser === 'mgr1a').status === 'failed');
  }
  // 連續 5 次可重試失敗（NETWORK）→ 熔斷；半開：失敗立刻重開、成功重置
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'NETWORK', message: 'down' });
    for (let i = 0; i < 4; i++) { const q = newQuote(w); await submitQuote(w, q, 'sales1', 'L1'); }
    eq('連續 4 次失敗：熔斷尚未開', (await w.outbox.breaker.state()).open, false);
    const q5 = newQuote(w);
    await submitQuote(w, q5, 'sales1', 'L1');
    t('第 5 次失敗：熔斷開啟', (await w.outbox.breaker.state()).open === true);
    const calls = w.sent.length;
    const q6 = newQuote(w);
    const r6 = await submitQuote(w, q6, 'sales1', 'L1');
    eq('熔斷開著的新事件：只入列，不嘗試寄送（不吃逾時）', [r6.e1.queued, w.sent.length - calls], [1, 0]);
    w.clock.advance(601 * SEC);
    const dh = await w.drainDue({ limit: 1 });                                // 半開：只放一封試探，仍失敗
    t('半開狀態下試探失敗：立刻重新開啟', dh.sent === 0 && (await w.outbox.breaker.state()).open === true);
    w.clock.advance(601 * SEC);
    w.behavior.fn = null;
    const ds = await w.drainDue({ limit: 10 });
    t('傳輸恢復且熔斷結束：排隊的信全部寄出、熔斷重置', ds.sent >= 5 && (await w.outbox.breaker.state()).open === false, ds.sent);
  }
  // 429 的權重：2 次就開（權重 3×2=6 ≥ 5）
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'THROTTLED', retryAfterSec: 5 });
    for (let i = 0; i < 2; i++) { const q = newQuote(w); await submitQuote(w, q, 'sales1', 'L1'); }
    t('THROTTLED 權重較高：2 次就開熔斷', (await w.outbox.breaker.state()).open === true);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('8 Hobby 方案重試路徑（當次嘗試 → 機會式清理 → 每日 Cron → 後台重送）', async () => {
  // ① 當次請求內短逾時嘗試，失敗留在 outbox；② 之後的請求順手 drainDue({limit:3}) 領取到期項目
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER' });
    const qs = [];
    for (let i = 0; i < 3; i++) { const q = newQuote(w); qs.push(q); await submitQuote(w, q, 'sales1', 'L1'); }
    eq('① 當次嘗試：3 封都失敗但沒影響簽核（3 張單都是 pending），3 筆留在 outbox', [qs.map((q) => q.approval.state), (await w.records()).filter((x) => x.status === 'pending').length], [['pending', 'pending', 'pending'], 3]);
    w.behavior.fn = null;
    w.clock.advance(30 * SEC);
    eq('② 機會式清理（30 秒後的下一個請求）：還沒到期，不處理', (await w.drainDue({ limit: 3 })).sent, 0);
    w.clock.advance(31 * SEC);
    const d = await w.drainDue({ limit: 3 });
    eq('② 機會式清理（61 秒後的請求順手呼叫 drainDue({limit:3})）：3 封到期的一次寄出', [d.sent, (await w.records()).filter((x) => x.status === 'sent').length], [3, 3]);
  }
  // limit 保護：每次請求最多順手處理 limit 筆
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER' });
    for (let i = 0; i < 4; i++) { const q = newQuote(w); await submitQuote(w, q, 'sales1', 'L1'); }
    w.behavior.fn = null;
    w.clock.advance(61 * SEC);
    const d1 = await w.drainDue({ limit: 3 });
    const d2 = await w.drainDue({ limit: 3 });
    eq('機會式清理 limit=3：第一次處理 3 筆，下一個請求處理剩下的 1 筆', [d1.sent, d2.sent], [3, 1]);
  }
  // ③ 整天沒有流量：隔天每日 Cron（limit 放大）一次清空
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    for (let i = 0; i < 3; i++) { const q = newQuote(w); await submitQuote(w, q, 'sales1', 'L1'); }
    w.behavior.fn = null;
    w.clock.advance(24 * 3600 * SEC);
    const d = await w.drainDue({ limit: 50 });
    eq('③ 每日 Cron（隔天 drainDue({limit:50})）：排隊整天的 3 封全部寄出', [d.sent, d.failed, (await w.records()).map((x) => x.status)], [3, [], ['sent', 'sent', 'sent']]);
  }
  // ④ 整天傳輸壞掉：Cron 後仍失敗 → 本輪用盡 → 最終失敗 → 後台重送
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER' });
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'L1');
    for (let day = 0; day < 4; day++) { w.clock.advance(24 * 3600 * SEC); await w.drainDue({ limit: 50 }); }
    const rec = (await w.records())[0];
    eq('④ 連續 4 天每天一次 Cron 都失敗：第 4 次嘗試後最終失敗（共 4 次嘗試）＋稽核', [rec.status, rec.attempts, w.logs.length], ['failed', 4, 1]);
    w.behavior.fn = null;
    await w.outbox.requeue(rec.id);
    eq('④ 後台「立即重送」：requeue 後下一次 drainDue 寄出', [(await w.drainDue({ limit: 5 })).sent, (await w.records())[0].status], [1, 'sent']);
  }
  // 兩個請求同時順手清理：不會重複寄
  {
    const w = mkWorld({ config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER' });
    for (let i = 0; i < 6; i++) { const q = newQuote(w); await submitQuote(w, q, 'sales1', 'L1'); }
    w.behavior.fn = async () => { await sleep(3); return { ok: true }; };
    w.clock.advance(61 * SEC);
    const before = w.sent.length;
    const ds = await Promise.all([w.drainDue({ limit: 3 }), w.drainDue({ limit: 3 }), w.drainDue({ limit: 3 }), w.drainDue({ limit: 3 })]);
    const mine = w.sent.slice(before).map((m) => m.subject);
    eq('4 個請求同時順手清理 6 筆到期工作：每筆恰好寄一次', [mine.length, new Set(mine).size, ds.reduce((a, s) => a + s.sent, 0)], [6, 6, 6]);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('9 稽核呼叫內容與 outbox 儲存內容衛生（全部世界）', async () => {
  // 稽核呼叫簽名：writeLog(action, operatorLabel, quoteNo, detail)。operatorLabel 是操作者的顯示名稱（既有稽核慣例），
  // outbox 紀錄的 actorLabel 欄位同理；除此之外，名稱、專案、客戶、金額、主旨、信件內容都不可出現在稽核與儲存體裡。
  const ACTIONS = new Set(['QUOTE_MAIL_FAILED', 'QUOTE_MAIL_SKIPPED']);
  const DETAIL_RE = /^type=E[1-6]_[A-Z0-9_]+( to=[A-Za-z0-9._-]+)? (code|reason)=[A-Z0-9_]+$/;
  const MONEY = ['NT$', '2,000,000', '60,000,000', '40.00', '20.00', '5.00%', '200000000', '6000000000'];
  const NAMES = [CUSTOMER, PROJECT, OWNER_LABEL, 'SALES1 Name', 'Mgr Nick', 'GM One', 'Sec One', '【簽核', '機密'];
  const hasAny = (s, list) => list.filter((x) => s.indexOf(x) >= 0);
  let logCount = 0;
  let worldsWithFile = 0;
  let jobCount = 0;
  WORLDS.forEach((w, i) => {
    w.logs.forEach((l) => {
      logCount += 1;
      const detail = String(l[3]);
      const why = [];
      if (!ACTIONS.has(l[0])) why.push('動作');
      if (!DETAIL_RE.test(detail)) why.push('detail 格式');
      if (/@/.test(JSON.stringify(l))) why.push('含 @');
      if (hasAny(detail, MONEY.concat(NAMES)).length) why.push('detail 含敏感內容');
      if (hasAny(String(l[1]) + String(l[2]), MONEY).length) why.push('operator／單號含金額');
      if (!/^(QU-[0-9]{6}-[0-9]{3})?$/.test(String(l[2]))) why.push('單號格式');
      if (why.length) record('稽核呼叫格式（世界 ' + i + '）：' + short(l), false, why.join(','));
    });
    let text = '';
    try { text = fs.existsSync(w.file) ? fs.readFileSync(w.file, 'utf8') : (w.adapter.dump ? w.adapter.dump() : ''); } catch (e) { text = ''; }
    if (text) {
      worldsWithFile += 1;
      const bad = [];
      if (FULL_EMAIL_RE.test(text)) bad.push('完整 email');
      let parsed = null;
      try { parsed = JSON.parse(text); } catch (e) { parsed = null; }
      if (!parsed || !Array.isArray(parsed.jobs)) bad.push('JSON 結構');
      else {
        jobCount += parsed.jobs.length;
        const noActor = JSON.stringify(parsed.jobs.map((j) => Object.assign({}, j, { actorLabel: '' })));
        hasAny(noActor, MONEY.concat(NAMES)).forEach((s) => bad.push(s));
        parsed.jobs.forEach((j) => { Object.keys(j).forEach((k) => { if (['html', 'text', 'subject', 'body', 'content', 'email', 'to', 'cc', 'bcc'].indexOf(k) >= 0) bad.push('欄位 ' + k); }); });
      }
      if (bad.length) record('outbox 儲存內容衛生（世界 ' + i + '）', false, bad.join(','));
    }
    const leftovers = fs.existsSync(w.dir) ? fs.readdirSync(w.dir).filter((f) => /\.(tmp|bak)$/.test(f) || /\.corrupt-/.test(f)) : [];
    if (leftovers.length) record('暫存目錄殘留（世界 ' + i + '）', false, leftovers.join(','));
  });
  t('全部 ' + WORLDS.length + ' 個世界、共 ' + logCount + ' 筆稽核呼叫：動作只有 QUOTE_MAIL_FAILED／QUOTE_MAIL_SKIPPED，detail 固定格式、不含 @、金額、專案／客戶／業務名、主旨', logCount >= 20, 'worlds=' + WORLDS.length + ' logs=' + logCount);
  t('全部世界的 outbox 儲存內容（JSON 檔案，共 ' + jobCount + ' 筆紀錄）不含完整 email、金額、專案／客戶名、主旨、信件內容欄位（actorLabel 是規格內的欄位，例外）；沒有 .tmp／.bak 殘留', worldsWithFile >= 20 && jobCount >= 100, 'files=' + worldsWithFile + ' jobs=' + jobCount);
});

// ═════════════════════════════════════════════════════════════════════════
section('10 P1 到 P2：批次補齊 email 之後的補寄', async () => {
  // 目前實況：帳號資料沒有 email 欄位
  const users = mkUsers();
  users.forEach((u) => { delete u.email; });
  const w = mkWorld({ users, usersForm: 'array' });
  const rep0 = UE.missingEmailReport({ users: w.users, roster: w.roster, config: w.config });
  eq('補齊前：缺口清單＝所有簽核相關角色（停用的 mgr1old 不列）', rep0.map((x) => x.username), ['chair1', 'cons1', 'gm1', 'gm2', 'mgr1a', 'proxy1', 'sec1', 'sec2', 'sec3']);
  t('補齊前：每一筆原因都是 NO_EMAIL', rep0.every((x) => x.reason === 'NO_EMAIL'));
  const q = newQuote(w);
  const r = await submitQuote(w, q, 'sales1', 'L2');
  eq('補齊前送簽：E1 略過（NO_EMAIL）、沒寄、簽核照常進行', [r.e1.sent, reasonMap(r.e1), q.approval.state], [0, { mgr1a: 'NO_EMAIL' }, 'pending']);
  eq('補齊前送簽：寫稽核 QUOTE_MAIL_SKIPPED', w.logs.map((l) => l[0] + ' ' + l[3]), ['QUOTE_MAIL_SKIPPED type=E1_SUBMIT to=mgr1a reason=NO_EMAIL']);

  // 管理員貼上批次清單 → 預覽 → 確認套用
  const text = [
    '# username, email（範例）', '',
    'mgr1a, mgr1a@itts.com.tw',
    'gm1\tgm1@itts.com.tw',
    'gm2, GM2@Itts.com.tw',
    'chair1, chair1@example.test',
    'ghost, ghost@itts.com.tw',
    'sec1, dup@itts.com.tw',
    'sec2, DUP@itts.com.tw',
    'cons1, cons1@itts.com.tw',
  ].join('\n');
  const plan = UE.parseBulkEmailText(text, { config: w.config, users: w.users });
  eq('批次預覽：各行狀態（空行與註解忽略；同批重複 email 兩行都標錯）', plan.rows.map((x) => x.username + ':' + x.status + (x.code ? ':' + x.code : '')), [
    'mgr1a:ok', 'gm1:ok', 'gm2:ok', 'chair1:error:DOMAIN_NOT_ALLOWED', 'ghost:error:UNKNOWN_USER', 'sec1:error:DUP_EMAIL_IN_BATCH', 'sec2:error:DUP_EMAIL_IN_BATCH', 'cons1:ok',
  ]);
  eq('批次預覽：統計', plan.summary, { ok: 4, unchanged: 0, error: 4 });
  const applied = UE.applyBulkPlan(plan.rows, w.users, { config: w.config });
  eq('套用：更新 4 人，稽核文字固定、不含位址', [applied.updated.map((x) => x.username), applied.updated.map((x) => x.detail)], [['mgr1a', 'gm1', 'gm2', 'cons1'], ['email 未設定→已設定', 'email 未設定→已設定', 'email 未設定→已設定', 'email 未設定→已設定']]);
  t('套用：回傳新陣列、原陣列沒被修改、gm2 的 email 已正規化成小寫', applied.users !== w.users && w.users.every((u) => u.email === undefined) && applied.users.find((u) => u.username === 'gm2').email === 'gm2@itts.com.tw');
  w.setUsers(applied.users);
  const rep1 = UE.missingEmailReport({ users: w.users, roster: w.roster, config: w.config });
  eq('補齊後：缺口清單縮短為尚未匯入的人', rep1.map((x) => x.username), ['chair1', 'proxy1', 'sec1', 'sec2', 'sec3']);

  // 已經略過的 E1：補上 email 後，後台對該筆「重送」即可（同一事件的去重鍵已被占用，所以不能靠重新觸發）
  const again = await sendStep(w, q, 'E1_SUBMIT', 0, 'sales1');
  eq('補 email 後同一事件重新觸發：視為重複（DUPLICATE），不會自動補寄', [again.sent, reasonMap(again)], [0, { mgr1a: 'DUPLICATE' }]);
  const skippedRec = (await w.records()).find((x) => x.toUser === 'mgr1a' && x.type === 'E1_SUBMIT');
  eq('略過的紀錄可由後台重送（requeue skipped→pending）', [skippedRec.status, skippedRec.skipReason, (await w.outbox.requeue(skippedRec.id)).ok], ['skipped', 'NO_EMAIL', true]);
  const d = await w.drainDue({});
  eq('重送後 drainDue 重新解析 email 並寄出（內容用現況重建）', [d.sent, w.sent.map((m) => m.to[0]), (await w.records()).find((x) => x.id === skippedRec.id).status], [1, ['mgr1a@itts.com.tw'], 'sent']);
  checkMail('補寄的 E1', w.sent[0], 'mgr1', 'E1_SUBMIT', { margin: '20.00%', marginDigits: '20.00', tag: '需總經理核准', link: w.link(q) });

  // 補齊後的新事件直接寄出；尚未補齊的人仍被略過並記錄
  const rA = await approveStep(w, q, 'mgr1a');
  eq('補齊後的下一關 E3：gm1、gm2 直接收到', [rA.e3.sent, w.sent.slice(1).filter((m) => m.tag === 'E3_NEXT_STEP').map((m) => m.to[0]).sort()], [2, ['gm1@itts.com.tw', 'gm2@itts.com.tw']]);
  const q2 = newQuote(w);
  await submitQuote(w, q2, 'sales1', 'BOARD');
  await approveStep(w, q2, 'mgr1a');
  const rB = await approveStep(w, q2, 'gm1');
  eq('尚未補 email 的秘書與代核人：董事會關 4 位全部 NO_EMAIL，沒有人收到、簽核不受阻', [rB.e3.sent, reasonMap(rB.e3), q2.approval.cur], [0, { sec1: 'NO_EMAIL', sec2: 'NO_EMAIL', sec3: 'NO_EMAIL', proxy1: 'NO_EMAIL' }, 2]);
  // 唯一性：不分大小寫、已被使用的位址不能再給別人
  const vu = UE.validateUserEmail('MGR1A@itts.com.tw', { config: w.config, users: w.users, selfUsername: 'sec3' });
  t('email 全站唯一（不分大小寫）：別人已使用的位址被拒絕', vu.ok === false && vu.code === 'DUPLICATE', short(vu));
  t('稽核用文字不含位址', UE.maskedAuditDetail('', 'a@itts.com.tw') === 'email 未設定→已設定' && UE.maskedAuditDetail('a@itts.com.tw', 'b@itts.com.tw') === 'email 已變更' && UE.maskedAuditDetail('a@itts.com.tw', '') === 'email 已清除');
});

// ═════════════════════════════════════════════════════════════════════════
section('11 連結鏈：信 → /q/:id 跳板頁 → 深層連結', async () => {
  const w = mkWorld();
  const q = newQuote(w);
  await submitQuote(w, q, 'sales1', 'L1');
  await requestCost(w, newQuote(w), 'sales1');
  const e1 = w.sent[0];
  const e2 = w.sent[1];
  const hrefOf = (m) => (/href="([^"]+)"/.exec(m.html) || [])[1];
  eq('E1 信內連結＝link.buildQuoteLink（render 與 link 兩個模組沒有漂移）', hrefOf(e1), LNK.buildQuoteLink(w.config, q.id));
  const q2id = qid(QSEQ);
  eq('E2 信內連結＝buildQuoteLink(…, {cost:true})', hrefOf(e2), LNK.buildQuoteLink(w.config, q2id, { cost: true }));
  eq('連結格式：APP_BASE_URL + /q/ + 單據 id（E1），E2 加 ?cost=1', [hrefOf(e1), hrefOf(e2)], [BASE_URL + '/q/' + q.id, BASE_URL + '/q/' + q2id + '?cost=1']);

  // 跳板頁
  const jp = LNK.renderJumpPage({ quoteId: q.id, cost: undefined });
  const jc = LNK.renderJumpPage({ quoteId: q2id, cost: '1' });
  eq('跳板頁（E1）：200、導向 /index.html#quote:<id>', [jp.status, jp.body.indexOf('url=/index.html#quote:' + q.id + '"') >= 0], [200, true]);
  eq('跳板頁（E2）：導向 #quote:<id>:cost', [jc.status, jc.body.indexOf('url=/index.html#quote:' + q2id + ':cost"') >= 0], [200, true]);
  t('跳板頁：no-store、noindex、no-referrer、自帶嚴格 CSP、不含 script', jp.headers['Cache-Control'] === 'no-store' && /noindex/.test(jp.headers['X-Robots-Tag']) && jp.headers['Referrer-Policy'] === 'no-referrer' && /default-src 'none'/.test(jp.headers['Content-Security-Policy']) && !/<script/i.test(jp.body));
  const bad = ['', '../etc/passwd', 'a/b', 'a b', 'x'.repeat(65), '<script>', null, undefined, 5, [], {}];
  const pages = bad.map((id) => LNK.renderJumpPage({ quoteId: id }));
  t('跳板頁：無效 id 一律 404 且內容與標頭完全相同（不洩漏單據是否存在、不回顯輸入）', pages.every((p) => p.status === 404) && new Set(pages.map((p) => p.body + JSON.stringify(p.headers))).size === 1);

  // 瀏覽器端：登入前記住、登入後取回（fake window，普通物件）
  const target = /url=([^"]+)"/.exec(jp.body)[1];                      // '/index.html#quote:<id>'
  const hash = target.slice(target.indexOf('#'));
  const store = {};
  const prevWin = global.window;
  global.window = { sessionStorage: { getItem: (k) => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); }, removeItem: (k) => { delete store[k]; } }, location: { hash } };
  try {
    t('深層連結：跳板頁導向的片段通過 ITTSDeepLink.isValidHash', DL.isValidHash(hash) && DL.isValidHash(hash + ':cost'));
    t('深層連結：登入頁載入時 remember(location.hash) 成功', DL.remember(global.window.location) === true);
    eq('深層連結：登入成功後 hashForRedirect() 取回同一個片段，且只能取一次', [DL.hashForRedirect(), DL.hashForRedirect()], [hash, '']);
    t('深層連結：被竄改的儲存內容（含外部網址）取不出來', (() => { store['itts.deepLink'] = '#quote:x/../../evil'; return DL.consume() === ''; })());
  } finally {
    if (prevWin === undefined) delete global.window; else global.window = prevWin;
  }
  // 與 _client/app.js 現行的深層連結解析（複本）相容：單據 id 是 uuid 時兩邊一致
  const APP_RE = /^#quote:([0-9a-fA-F-]{8,64})(:cost)?$/;       // 複本自 _client/app.js _handleQuoteDeepLink（2026-10-08 工作樹）
  t('與 app.js 現行 #quote: 解析相容（uuid 單據 id；非 uuid 的 id 需放寬 app.js 才能被處理）', APP_RE.test(hash) && APP_RE.test(hash + ':cost') && !APP_RE.test('#quote:not_a_uuid-ZZ'));
  t('連結不帶 token、不是一鍵核准（沒有 approve／action 之類參數）', !/token|approve|action|hash=/i.test(hrefOf(e1) + hrefOf(e2)));
});

// ═════════════════════════════════════════════════════════════════════════
section('12 與真實 lib/quoteApproval.js 銜接（用真的 derived 與核決路徑）', async () => {
  if (!QA || typeof QA.buildDerived !== 'function' || typeof QA.requiredPath !== 'function') {
    t('【略過】lib/quoteApproval.js 無法載入（' + (QA_WHY || 'N/A') + '）：未跑真實銜接測試（整合階段補跑）', true);
    return;
  }
  const PCS = { 'Product A': { cls: 'consult', costBySales: true } };
  const realQuote = (qty, unitPrice, cost) => ({ products: ['Product A'], discountType: 'none', items: [{ lid: 'a', desc: 'Item A', qty, unit: 'set', unitPrice, cost }] });

  // ① 預設 derived 與真實規則一致（防止本檔的合成數字與真規則漂移）
  const chk = (key, rowKey, rev, gp) => {
    const p = QA.requiredPath(rowKey, gp, rev);
    const d = DERIVED[key];
    eq('合成 derived ' + key + ' 與 QA.requiredPath 一致（level／board／tiers／毛利率文字）', [p.level, p.board, p.tiers, QA.marginText(gp, rev)], [d.level, d.board, d.tiers, d.marginText]);
  };
  chk('L1', 'consult', 200000000, 80000000);
  chk('L2', 'consult', 200000000, 40000000);
  chk('L3', 'consult', 200000000, 10000000);
  chk('BOARD', 'consult', 6000000000, 2400000000);
  chk('LOSS', 'consult', 100000000, -3210000);

  // ② 真實 buildDerived → 事件 → 信件：金額與毛利率逐字對得上
  {
    const w = mkWorld();
    const q = newQuote(w);
    const d = QA.buildDerived(realQuote(2, 1000000, 600000), PCS);
    t('真實 buildDerived 產出 derived（營收 2,000,000、毛利率 40.00、一級主管可核）', !!d && d.revenueCents === 200000000 && d.gpCents === 80000000 && d.marginText === '40.00' && d.level === 1 && d.board === false, short(d));
    await submitQuote(w, q, 'sales1', d);
    const m = w.sent[0];
    t('真實 derived：信裡的金額＝RND.formatNtd(derived.revenueCents)、毛利率＝derived.marginText + %', m.text.indexOf(RND.formatNtd(d.revenueCents)) >= 0 && m.text.indexOf(d.marginText + '%') >= 0, m.text.slice(0, 160));
    t('真實 derived：目前關卡文字是系統的 TIERS 標籤（QA.TIERS.mgr1）', m.text.indexOf(QA.TIERS.mgr1) >= 0);
    const dBoard = QA.buildDerived(realQuote(1, 60000000, 54000000), PCS);
    t('真實 buildDerived（金額 6,000 萬、毛利 10%）→ 需董事會（board=true、level=3、路徑 mgr1→gm→board）', !!dBoard && dBoard.board === true && dBoard.level === 3 && dBoard.tiers.join() === 'mgr1,gm,board', short(dBoard));
    const q2 = newQuote(w);
    await submitQuote(w, q2, 'sales1', dBoard);
    await approveStep(w, q2, 'mgr1a');
    const n = w.sent.length;
    const rr = await approveStep(w, q2, 'gm1');
    const boardMails = w.sent.slice(n).filter((x) => x.tag === 'E3_NEXT_STEP');
    t('真實 derived 走完整條路徑到董事會關：4 封信有金額、毛利率、「需董事會決議」，沒有客戶名與業務名', rr.e3.sent === 4 && boardMails.every((x) => x.text.indexOf(RND.formatNtd(dBoard.revenueCents)) >= 0 && x.text.indexOf(dBoard.marginText + '%') >= 0 && x.html.indexOf('需董事會決議') >= 0 && blobOf(x).indexOf(CUSTOMER) < 0 && blobOf(x).indexOf(OWNER_LABEL) < 0), short(rr.e3));
    t('真實 TIERS 標籤「董事會決議（秘書代核）」當作關卡文字顯示；色塊標籤用短標籤「董事會」（未映射會變成「需董事會決議（秘書代核）核准」）', boardMails.every((x) => x.text.indexOf(QA.TIERS.board) >= 0 && x.text.indexOf('需董事會決議（秘書代核）核准') < 0));
  }

  // ③ 規則表 × 毛利 × 金額的格子：每種 derived 都能組出合法事件並渲染給所有簽核人角色
  {
    const rows = QA.ROWS.map((r) => r.key);
    const revs = [100000, 99999999, 1000000000, 1000000001, 5000000000, 5000000001];
    const pcts = [-5, 0, 3, 7, 12, 18, 30, 60];
    let combos = 0;
    let rendered = 0;
    let bad = '';
    const w = mkWorld();
    const q = newQuote(w);
    for (const rowKey of rows) {
      for (const rev of revs) {
        for (const pct of pcts) {
          const gp = Math.round(rev * pct / 100);
          const p = QA.requiredPath(rowKey, gp, rev);
          const mt = QA.marginText(gp, rev);
          const d = { level: p.level, board: p.board, tiers: p.tiers, revenueCents: rev, gpCents: gp, marginText: mt };
          const num = numbersOf(d);
          combos += 1;
          const wantTag = d.board ? '需董事會決議' : (d.level === 1 ? '一級主管可核' : (d.level === 2 ? '需總經理核准' : '需董事長核准'));
          const kinds = ['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy'];
          for (const kind of kinds) {
            const ev = { type: 'E3_NEXT_STEP', quoteId: q.id, quoteNo: q.quoteNo, projectName: PROJECT, company: CUSTOMER, ownerLabel: OWNER_LABEL, step: { level: 2, label: '總經理' }, numbers: num, at: iso(T0), stepKey: 'grid#1' };
            const v = EVT.validateEvent(ev);
            if (!v.ok) { bad = 'validate ' + rowKey + '/' + rev + '/' + pct + ' ' + v.error; break; }
            let m;
            try { m = RND.renderMail(ev, { username: 'x', label: 'X', kind }, { config: w.config }); } catch (e) { bad = 'render ' + rowKey + '/' + rev + '/' + pct + '/' + kind + ' ' + e.message; break; }
            const okAmt = m.text.indexOf(RND.formatNtd(rev)) >= 0 && m.text.indexOf((mt.replace(/%$/, '')) + '%') >= 0 && m.text.indexOf(wantTag) >= 0;
            const okVis = (kind === 'secretary' || kind === 'boardProxy')
              ? (m.text.indexOf(CUSTOMER) < 0 && m.text.indexOf(OWNER_LABEL) < 0 && m.text.indexOf(PROJECT) < 0 && m.subject.indexOf(PROJECT) < 0 && m.html.indexOf(PROJECT) < 0)
              : (m.text.indexOf(CUSTOMER) >= 0 && m.text.indexOf(OWNER_LABEL) >= 0 && m.text.indexOf(PROJECT) >= 0);
            if (!okAmt || !okVis) { bad = 'content ' + rowKey + '/' + rev + '/' + pct + '/' + kind + ' tag=' + wantTag + ' amt=' + okAmt + ' vis=' + okVis; break; }
            rendered += 1;
          }
          if (bad) break;
        }
        if (bad) break;
      }
      if (bad) break;
    }
    t('規則表 ' + rows.length + ' 列 × 金額 ' + revs.length + ' 檔 × 毛利 ' + pcts.length + ' 檔＝' + combos + ' 種 derived：全部能組出合法事件、渲染 ' + rendered + ' 封（金額、毛利率、層級標籤、角色可見性都正確）', !bad && combos === rows.length * revs.length * pcts.length && rendered === combos * 5, bad || rendered);
  }

  // ④ 負毛利
  {
    const w = mkWorld();
    const q = newQuote(w);
    await submitQuote(w, q, 'sales1', 'LOSS');
    const m = w.sent[0];
    t('虧損單（毛利率 -3.21%）：信裡顯示 -3.21% 與「毛利為負」，色塊標籤仍有文字', m.text.indexOf('-3.21%') >= 0 && m.text.indexOf('毛利為負') >= 0 && m.html.indexOf('-3.21%') >= 0, m.text.slice(0, 200));
  }

  // ④b 極端虧損單：成本是營收的 1 萬倍以上，真實 marginText 的整數部分超過 6 位（-99999900.00）。
  //    以前 validateEvent 退件（BAD_EVENT）→ 簽核人完全收不到信；現在照常寄出，毛利率顯示 <-999999%
  {
    const w = mkWorld();
    const q = newQuote(w);
    const d = QA.buildDerived(realQuote(1, 1, 1000000), PCS);        // 營收 NT$1、成本 NT$1,000,000
    t('真實 buildDerived 的極端虧損單：marginText 整數部分超過 6 位', !!d && /^-[0-9]{7,}\./.test(d.marginText), short(d));
    const r = await submitQuote(w, q, 'sales1', d);
    t('極端虧損單送簽：沒有 failed／BAD_EVENT，簽核人收到信', r.e1.failed.length === 0 && r.e1.errors === 0 && r.e1.sent >= 1 && w.sent.length >= 1, short(r.e1));
    const m = w.sent[0];
    t('極端虧損單：信裡毛利率顯示 <-999999% 與「毛利為負」，不含 8 位數的原始毛利率', m.text.indexOf('<-999999%') >= 0 && m.text.indexOf('毛利為負') >= 0 && m.html.indexOf('&lt;-999999%') >= 0 && m.text.indexOf(d.marginText.replace(/^-/, '').split('.')[0]) < 0, m.text.slice(0, 220));
    t('極端虧損單：outbox 沒有留下失敗紀錄、稽核沒有 QUOTE_MAIL_FAILED', (await w.records()).every((x) => x.status === 'sent') && w.logs.every((l) => l[0] !== 'QUOTE_MAIL_FAILED'), short(w.logs));
  }

  // ⑤ 真實單據的品項（item／title／subtotal）轉 E2 品項，價格與成本不會外洩
  {
    const QI = (() => { try { return load('lib/quoteItems.js'); } catch (e) { return null; } })();
    if (!QI || typeof QI.itemRows !== 'function') {
      t('【略過】lib/quoteItems.js 無法載入：未驗證 itemRows 對應', true);
    } else {
      const w = mkWorld();
      const q = newQuote(w);
      const viaQI = QI.itemRows(q.items).map((it) => ({ desc: it.desc, qty: it.qty, unit: it.unit }));
      eq('QI.itemRows(q.items)（真實品項列過濾）＝本檔 itemsOf：標題列與小計列不列入', viaQI, itemsOf(q));
    }
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('13 惡意內容穿過整條管線', async () => {
  const evil = {
    project: '<script>alert(1)</script>\r\nBcc: x@evil.test',
    company: '"><img src=x onerror=alert(1)>',
    reason: 'javascript:alert(1) ' + cp(0x202e) + 'txt.exe <b>x</b> & "q"',
    desc: '<iframe src=//evil.test></iframe>',
  };
  const w = mkWorld();
  const q = newQuote(w, { projectName: evil.project, company: evil.company, items: [{ lid: 'z', desc: evil.desc, qty: 1, unit: '<u>' }] });
  await submitQuote(w, q, 'sales1', 'L2');
  await approveStep(w, q, 'mgr1a');                                           // E4 approved ＋ E3 給兩位總經理
  await rejectStep(w, q, 'gm1', evil.reason);                                 // E4 rejected（原因含惡意字串）
  await requestCost(w, q, 'sales1');                                          // E2（品項含惡意字串）
  eq('惡意內容：每一封信都寄出（E1、E4 本關通過、E3×2、E4 駁回、E2 共 6 封）', w.sent.map((m) => m.tag).sort(), ['E1_SUBMIT', 'E2_COST_REQUEST', 'E3_NEXT_STEP', 'E3_NEXT_STEP', 'E4_RESULT', 'E4_RESULT']);
  t('惡意內容：主旨沒有 CR／LF（標頭注入失敗），且通過傳輸層驗證', w.sent.every((m) => !/[\r\n]/.test(m.subject) && TRN.validateMessage(m).ok));
  t('惡意內容：每封信只有 1 位已知收件人（專案名稱裡的 Bcc 字樣沒有變成收件人）', w.sent.every((m) => m.to.length === 1 && ['mgr1a', 'gm1', 'gm2', 'sales1', 'cons1'].map((u) => u + '@itts.com.tw').indexOf(m.to[0]) >= 0));
  // 標籤集合必須和一封正常信的標籤集合相同（沒有任何新標籤），且任何標籤內都沒有 on* 屬性
  const tagSet = (h) => new Set((h.match(/<\/?[A-Za-z][A-Za-z0-9]*/g) || []).map((x) => x.replace('/', '').toLowerCase()));
  const benign = (type, kind, extra) => RND.renderMail(Object.assign({ type, quoteId: qid(999), quoteNo: 'QU-000000-000', projectName: 'ok', company: 'ok', ownerLabel: 'ok', at: iso(T0), stepKey: 'b#0' }, extra), { username: 'x', label: 'x', kind }, { config: w.config }).html;
  const baseTags = new Set();
  [benign('E1_SUBMIT', 'mgr1', { step: { level: 1, label: '一級主管' }, numbers: numbersOf(DERIVED.L1), actor: { label: 'ok' } }),
    benign('E2_COST_REQUEST', 'consultant', { items: [{ desc: 'ok', qty: 1, unit: 'set' }], actor: { label: 'ok' } }),
    benign('E4_RESULT', 'owner', { step: { level: 1, label: '一級主管' }, result: { kind: 'rejected', reason: 'ok' }, actor: { label: 'ok' } }),
    benign('E6_WITHDRAWN', 'gm', { result: { kind: 'withdrawn' }, actor: { label: 'ok' } })].forEach((h) => tagSet(h).forEach((x) => baseTags.add(x)));
  t('惡意內容：html 的標籤集合沒有任何基準信件以外的標籤（沒有 script／iframe／img／b／u）', w.sent.every((m) => Array.from(tagSet(m.html)).every((x) => baseTags.has(x))) && !baseTags.has('script') && !baseTags.has('iframe') && !baseTags.has('img'), short(w.sent.map((m) => Array.from(tagSet(m.html)).filter((x) => !baseTags.has(x)))));
  t('惡意內容：任何 HTML 標籤內都沒有 on* 事件屬性、javascript: 網址', w.sent.every((m) => (m.html.match(/<[^>]*>/g) || []).every((tg) => !/\son[a-z]+\s*=/i.test(tg) && !/javascript:/i.test(tg))));
  t('惡意內容：跳脫後的文字仍可在 html 看到（顯示為純文字，不是被吃掉）', w.sent.filter((m) => m.tag === 'E1_SUBMIT').every((m) => m.html.indexOf('&lt;script&gt;') >= 0));
  t('惡意內容：href 只有站內連結', w.sent.every((m) => (m.html.match(/href="[^"]*"/g) || []).every((h) => h.indexOf('href="' + BASE_URL + '/q/') === 0)));
  t('惡意內容：方向控制字元（RLO）不會進信（主旨與內文）', w.sent.every((m) => blobOf(m).indexOf(cp(0x202e)) < 0));
  t('惡意內容：純文字版保留原字串（純文字不需跳脫）但沒有 HTML 標籤被當成結構', w.sent.every((m) => m.text.indexOf('<html') < 0));
  // 超長欄位（在事件允許的上限內）
  const long = newQuote(w, { projectName: 'P'.repeat(200), company: 'C'.repeat(100) });
  const r = await submitQuote(w, long, 'sales1', 'L1');
  t('超長專案名稱（200 字，系統上限）：主旨截斷到 120 字內、信件仍寄出且 <100KB', r.e1.sent === 1 && w.sent[w.sent.length - 1].subject.length <= 120 && Buffer.byteLength(w.sent[w.sent.length - 1].html, 'utf8') < 100 * 1024, w.sent[w.sent.length - 1].subject.length);
  // 單號含特殊字元的防呆：quoteId 不合法時整個事件被拒（連結不可能被注入）
  const badId = newQuote(w);
  badId.id = 'x"><script>alert(1)</script>';
  const ev = buildEvent(w, badId, 'E1_SUBMIT', { stepKey: 'k#0', step: { level: 1, label: '一級主管' }, actor: 'sales1' });
  ev.numbers = numbersOf(DERIVED.L1);
  const rb = await w.dispatch(ev, [{ username: 'mgr1a', kind: 'mgr1' }], { actorUsername: 'sales1' });
  eq('單據 id 含引號與標籤：事件被拒（BAD_EVENT），不會產生連結', [rb.sent, rb.failed.map((x) => x.code)], [0, ['BAD_EVENT']]);
});

// ═════════════════════════════════════════════════════════════════════════
section('14 dispatch／drainDue 永不 throw 總帳', async () => {
  t('整個測試過程中 dispatch／drainDue 共被呼叫 ' + GUARD.calls + ' 次（> 150，確認不是空轉）', GUARD.calls > 150, GUARD.calls);
  eq('整個測試過程中 dispatch／drainDue throw 或 reject 的次數', GUARD.throws, 0);
  eq('整個測試過程中 dispatch／drainDue 卡住（>20 秒）的次數', GUARD.hangs, 0);
  eq('異常明細', GUARD.bad, []);
  // 隨機亂輸入（與膠水無關的壓力）：真 render + JSON outbox + 隨機模式與傳輸行為
  let seed = 424242;
  const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
  const pick = (a) => a[Math.floor(rnd() * a.length)];
  const behaviors = [null, () => { throw new Error('boom'); }, () => Promise.reject(new Error('boom')), () => ({ ok: false, code: 'TIMEOUT' }), () => ({ ok: false, code: 'AUTH' }), () => 5, () => ({ ok: true })];
  let bad = 0;
  let total = 0;
  for (let i = 0; i < 120; i++) {
    const w = mkWorld({ env: { MAIL_MODE: pick(['live', 'live', 'redirect', 'log', 'off', 'garbage']), MAIL_REDIRECT_TO: 'redirect.box@example.test' }, glueFilter: pick([true, false]), timeouts: { connectMs: 5, totalMs: 40 } });
    w.behavior.fn = pick(behaviors);
    const q = newQuote(w);
    const derived = pick(['L1', 'L2', 'L3', 'BOARD', 'LOSS']);
    try {
      total += 1;
      await submitQuote(w, q, pick(['sales1', 'sales2', 'ghost']), derived);
      for (let k = 0; k < 2 && q.approval.state === 'pending'; k++) { await approveStep(w, q, pick(['mgr1a', 'gm1', 'chair1', 'sec1'])); await w.drainDue({ limit: pick([1, 5, 50]) }); w.clock.advance(pick([1, 61, 400, 1000]) * SEC); }
      await withdrawQuote(w, q, 'sales1');
      await w.drainDue({});
    } catch (e) { bad += 1; GUARD.bad.push('隨機流程 throw：' + (e && e.message)); }
  }
  t('隨機 ' + total + ' 組（模式×傳輸行為×流程）完整流程：全部正常結束、沒有未預期例外（異常 ' + bad + '）', bad === 0, GUARD.bad.slice(-3).join(' | '));
  eq('隨機壓力之後 dispatch／drainDue 仍然 0 次 throw、0 次卡住', [GUARD.throws, GUARD.hangs], [0, 0]);
});

main();
