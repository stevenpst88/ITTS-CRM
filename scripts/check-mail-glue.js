#!/usr/bin/env node
'use strict';
/**
 * 信件派送「整合膠水」單元測試：lib/mail/quoteMail.js 與 lib/mail/routes.js
 * 用法：node scripts/check-mail-glue.js（不開伺服器、不連網、不碰 data.json／auth.json／audit.log.json；只用記憶體 outbox 與假傳輸）
 *
 * 與其他測試的分工：
 *   check-mail-integration.js  用「膠水 stand-in」驗證 events／recipients／render／dispatcher／outbox 的銜接
 *   本檔                       驗證「真的膠水」quoteMail.js（事件建構、收件人類型、E6 快照、isStillValid／rebuild、機會式清理的節流與逾時）
 *                              與 routes.js（Cron 驗證、寄件匣後台、Email 批次匯入）的行為；quoteRoutes 的 helper（stepRecipients 等）在這裡是假的小複本，
 *                              真的 helper 與 HTTP 層由 API 層端對端測試（e2e_mail.js，不在 repo 內）覆蓋。
 *
 * 環境變數 MAIL_CORE_ROOT：要檢查的專案根目錄（變異測試用）。
 * 撰寫備註：本檔不寫四位數的 \uXXXX，一律用 cp()／\x..。
 */
const path = require('path');
const util = require('util');
const assert = require('assert');

const ROOT = process.env.MAIL_CORE_ROOT || path.join(__dirname, '..');
const load = (rel) => require(path.join(ROOT, rel));

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
const realWarn = console.warn;
const realError = console.error;
function finish() {
  console.warn = realWarn; console.error = realError;
  const failed = results.filter((r) => !r.ok);
  const bySec = {};
  results.forEach((r) => { const s = bySec[r.section] || (bySec[r.section] = { p: 0, f: 0 }); if (r.ok) s.p++; else s.f++; });
  Object.keys(bySec).forEach((k) => console.log((bySec[k].f ? 'FAIL ' : 'ok   ') + k + '  通過 ' + bySec[k].p + (bySec[k].f ? '，失敗 ' + bySec[k].f : '')));
  failed.forEach((r) => console.log('  ✗ [' + r.section + '] ' + r.name + (r.extra ? '  → ' + r.extra : '')));
  console.log((failed.length ? 'FAILED' : 'PASSED') + '：' + (results.length - failed.length) + ' / ' + results.length);
  process.exit(failed.length ? 1 : 0);
}
async function main() {
  console.warn = () => {}; console.error = () => {};      // 錯誤路徑的測試會讓膠水印出 [quote mail]／[mail cron] 警告，這裡不需要看到
  const watchdog = setTimeout(() => { console.log('WATCHDOG：測試超過 120 秒未結束'); process.exit(2); }, 120000);
  process.on('unhandledRejection', (e) => { record('未處理的 Promise rejection', false, (e && e.message) || String(e)); });
  for (const s of sections) {
    currentSection = s.name;
    try { await s.fn(); } catch (e) { record('章節執行中例外', false, (e && e.stack) || e); }
  }
  clearTimeout(watchdog);
  finish();
}

const { getMailConfig } = load('lib/mail/config.js');
const { createOutbox, memoryAdapter } = load('lib/mail/outbox.js');
const QM = load('lib/mail/quoteMail.js');
const RT = load('lib/mail/routes.js');
const { createQuoteMail } = QM;

const T0 = Date.UTC(2026, 9, 8, 3, 0, 0);
const iso = (ms) => new Date(ms).toISOString();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const FULL_EMAIL_RE = /[A-Za-z0-9._+-]{2,}@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+/;
const QID = 'b3f1c2d4-1111-4222-8333-000000000001';
const COMPANY = 'Synthetic Customer Co';
const PROJECT = 'Synthetic Project Alpha';
const OWNER_NICK = 'Nick S';

// ═════════════════════════════════════════════════════════════════════════
// 假世界：帳號、名冊、單據、假傳輸、稽核收集器、假時鐘
// ═════════════════════════════════════════════════════════════════════════
const isActive = (u) => !!u && u.active !== false && u.role !== 'pool';
const canAct = (u) => isActive(u) && u.role !== 'admin' && u.accessMode !== 'view';
const dispName = (users, un) => (users[un] && (users[un].nickname || users[un].displayName || users[un].username)) || un || '';
// quoteRoutes.js 同名函式的小複本（這裡只為了讓膠水有東西可以呼叫；真的函式由端對端測試覆蓋）
function stepRecipients(step, ctx) {
  if (!step) return [];
  if (step.tier === 'mgr1') return [step.assignee];
  if (step.tier === 'gm') return ctx.cfg.roster.gm;
  if (step.tier === 'chairman') return ctx.cfg.roster.chairman;
  if (step.tier === 'board') {
    const s = new Set(ctx.cfg.roster.boardProxy);
    Object.values(ctx.users).forEach((u) => { if (u.role === 'secretary') s.add(u.username); });
    return [...s].filter((un) => canAct(ctx.users[un]));
  }
  return [];
}
const getCfg = (data) => ({ roster: Object.assign({ gm: [], chairman: [], boardProxy: [], costProviders: [], sealManagers: [] }, (data.quoteApproval || {}).roster) });
const HELPERS = { getCfg, stepRecipients, isActive, dispName };

function mkUsers() {
  const mk = (username, role, extra) => Object.assign({ username, role, displayName: username.toUpperCase() + ' Name', email: username + '@itts.test' }, extra || {});
  return [
    mk('sales1', 'user', { nickname: OWNER_NICK, supervisor: 'm1' }), mk('m1', 'manager1'), mk('gm1', 'executive'), mk('gm2', 'executive'), mk('ch1', 'executive'),
    mk('sec1', 'secretary'), mk('prx1', 'user'), mk('cons1', 'user'), mk('cons2', 'user'),
    { username: 'noemail', role: 'secretary', displayName: 'No Email' },
    mk('off1', 'user', { active: false }), mk('adm1', 'admin'),
  ];
}
const DERIVED = {
  L2: { level: 2, board: false, tiers: ['mgr1', 'gm'], revenueCents: 200000000, gpCents: 20000000, marginText: '10.00' },
  BOARD: { level: 3, board: true, tiers: ['mgr1', 'gm', 'board'], revenueCents: 6000000000, gpCents: 2400000000, marginText: '40.00' },
};
const STEP_LABEL = { mgr1: '一級主管', gm: '總經理', chairman: '董事長', board: '董事會決議（秘書代核）' };
function mkQuote(derivedKey, over) {
  const d = DERIVED[derivedKey || 'L2'];
  return Object.assign({
    id: QID, quoteNo: 'QU-261008-001', owner: 'sales1', company: COMPANY, projectName: PROJECT, costBy: 'cons1',
    costFlow: { state: 'requested', requestedAt: iso(T0 - 5000), filledAt: null },
    items: [
      { lid: 'i1', desc: 'Install service', qty: 2, unit: 'set', unitPrice: 98765, cost: 54321 },
      { lid: 'i2', kind: 'title', desc: 'Section header' },
      { lid: 'i3', desc: 'Training', qty: 1.5, unit: 'day', unitPrice: 12345, cost: 6789 },
      { lid: 'i4', kind: 'subtotal', desc: 'Subtotal' },
    ],
    approval: {
      state: 'pending', submittedAt: iso(T0), submittedBy: 'sales1', cur: 0, derived: d,
      steps: d.tiers.map((tier, i) => ({ tier, label: STEP_LABEL[tier], assignee: tier === 'mgr1' ? 'm1' : null, status: i === 0 ? 'pending' : 'waiting', by: null, comment: '' })),
      history: [],
    },
  }, over || {});
}

function mkWorld(opt) {
  const o = opt || {};
  const clock = { t: T0 };
  const env = Object.assign({ MAIL_MODE: 'live', MAIL_ALLOWED_DOMAINS: 'itts.test', APP_BASE_URL: 'https://crm.example.test', MAIL_PREVIEW_DIR: path.join(require('os').tmpdir(), 'mail-glue-test-unused') }, o.env || {});
  const config = getMailConfig(env);
  config.timeouts = o.timeouts || { connectMs: 20, totalMs: 500 };
  const adapter = o.adapter || memoryAdapter();
  const outbox = o.noOutbox ? null : createOutbox(adapter, { now: () => clock.t, config });
  const w = { clock, config, outbox, sent: [], logs: [], quote: o.quote || mkQuote(), users: mkUsers(), roster: { gm: ['gm1', 'gm2'], chairman: ['ch1'], boardProxy: ['prx1'], costProviders: ['cons1', 'cons2'], sealManagers: [] }, behavior: null };
  w.data = { quotations: [w.quote], quoteApproval: { roster: w.roster } };
  w.userMap = () => { const m = {}; w.users.forEach((u) => { m[u.username] = u; }); return m; };
  w.transport = {
    kind: 'FAKE',
    async send(msg) {
      w.sent.push(msg);
      if (w.behavior) return w.behavior(msg);
      return { ok: true, providerId: 'fake:' + w.sent.length };
    },
  };
  w.qm = createQuoteMail({
    config, outbox, transport: w.transport, now: () => clock.t, limits: o.limits,
    db: { load: () => w.data, flush: async () => { w.flushes = (w.flushes || 0) + 1; } },
    loadAuth: () => ({ users: w.users }),
    writeLog: (action, operator, target, detail) => { w.logs.push({ action, operator, target, detail }); },
  });
  if (!o.noBind) w.qm.bind(HELPERS);
  w.ctx = (me) => ({ me: me || 'sales1', users: w.userMap(), cfg: getCfg(w.data), data: w.data });
  return w;
}
/** 呼叫 notifyMail 並等它做完 */
async function fire(w, spec, me, q) {
  const req = {};
  w.qm.notifyMail(req, w.ctx(me), q || w.quote, spec);
  await Promise.all(req._mailPending || []);
  return req;
}
const toOf = (m) => m.to[0];
const textOf = (m) => m.subject + '\n' + m.html + '\n' + m.text;

// ═════════════════════════════════════════════════════════════════════════
section('1 純函式：kindOf／numbersOf／stepOf／itemsOf', () => {
  eq('kindOf：各關卡', ['mgr1', 'gm', 'chairman', 'owner'].map((x) => QM.kindOf(x, null)), ['mgr1', 'gm', 'chairman', 'owner']);
  eq('kindOf：董事會關 — 秘書角色 → secretary', QM.kindOf('board', { role: 'secretary' }), 'secretary');
  eq('kindOf：董事會關 — 其他代核人 → boardProxy', QM.kindOf('board', { role: 'user' }), 'boardProxy');
  eq('kindOf：董事會關 — 帳號不存在 → boardProxy（不會誤給 secretary 以外的資料）', QM.kindOf('board', undefined), 'boardProxy');
  eq('kindOf：未知關卡 → mgr1（最小風險的簽核人類型）', QM.kindOf('zzz', null), 'mgr1');

  eq('numbersOf：一般', QM.numbersOf(DERIVED.L2), { revenueCents: 200000000, gpCents: 20000000, marginText: '10.00', marginPct: 10, tierLevel: 2, tierLabel: '總經理' });
  eq('numbersOf：董事會關 → tierLevel 3＋短標籤「董事會」（不是 TIERS.board 的完整文字）', QM.numbersOf(DERIVED.BOARD).tierLabel + '/' + QM.numbersOf(DERIVED.BOARD).tierLevel, '董事會/3');
  eq('numbersOf：虧損（負毛利率）', QM.numbersOf({ level: 3, board: false, revenueCents: 100, gpCents: -50, marginText: '-50.00' }).marginPct, -50);
  [['null', null], ['沒有 marginText', { level: 1, revenueCents: 1, gpCents: 1 }], ['marginText 不是數字', { level: 1, revenueCents: 1, gpCents: 1, marginText: 'abc' }],
    ['revenueCents 是小數', { level: 1, revenueCents: 1.5, gpCents: 1, marginText: '1' }], ['revenueCents 為負', { level: 1, revenueCents: -1, gpCents: 1, marginText: '1' }],
    ['gpCents 是字串', { level: 1, revenueCents: 1, gpCents: '1', marginText: '1' }], ['level 不是 1／2／3 且不是董事會', { level: 9, revenueCents: 1, gpCents: 1, marginText: '1' }],
    ['revenueCents 超過安全整數', { level: 1, revenueCents: 2 ** 60, gpCents: 1, marginText: '1' }]].forEach(([name, d]) => {
    eq('numbersOf：' + name + ' → null（信上沒有決策條，而不是整封寄不出去）', QM.numbersOf(d), null);
  });

  eq('stepOf：mgr1／gm／chairman／board', ['mgr1', 'gm', 'chairman', 'board'].map((x) => QM.stepOf({ tier: x, label: STEP_LABEL[x] }).level), [1, 2, 3, 'board']);
  eq('stepOf：未知關卡 level=null、label 保留', QM.stepOf({ tier: 'zzz', label: 'X' }), { level: null, label: 'X' });
  eq('stepOf：label 為空 → 有預設文字（驗證要求非空）', QM.stepOf({ tier: 'gm', label: '' }).label.length > 0, true);
  eq('stepOf：__proto__ 之類的 tier 不會命中原型屬性', QM.stepOf({ tier: '__proto__', label: 'X' }).level, null);

  const items = QM.itemsOf(mkQuote());
  eq('itemsOf：標題與小計列被濾掉，只剩 2 個品項', items.length, 2);
  eq('itemsOf：只有 desc／unit／qty 三個欄位（價格與成本在這裡就被丟掉）', items.map((x) => Object.keys(x).sort().join()), ['desc,qty,unit', 'desc,qty,unit']);
  t('itemsOf：輸出裡沒有任何價格數字', !/98765|12345|54321|6789/.test(JSON.stringify(items)), JSON.stringify(items));
  eq('itemsOf：空說明 → 預設文字（驗證要求非空）', QM.itemsOf({ items: [{ desc: '  ', qty: 1, unit: 'x' }] })[0].desc, '(未命名品項)');
  eq('itemsOf：qty 為字串 → 保留（截 20 字）；負數或 NaN → 不帶 qty', QM.itemsOf({ items: [{ desc: 'a', qty: '3 式', unit: '' }, { desc: 'b', qty: -1 }, { desc: 'c', qty: NaN }] }).map((x) => x.qty), ['3 式', undefined, undefined]);
  eq('itemsOf：不是陣列 → 空', QM.itemsOf({ items: 'x' }), []);
  eq('itemsOf：最多 200 列', QM.itemsOf({ items: Array.from({ length: 300 }, (_, i) => ({ desc: 'i' + i })) }).length, 200);
});

// ═════════════════════════════════════════════════════════════════════════
section('2 停用狀態與零成本', async () => {
  // 沒有 outbox
  const w0 = mkWorld({ noOutbox: true });
  t('沒有 outbox：enabled=false', w0.qm.enabled() === false);
  const req0 = {};
  w0.qm.notifyMail(req0, w0.ctx(), w0.quote, { type: 'E1_SUBMIT', idx: 0 });
  t('沒有 outbox：notifyMail 什麼都不做（沒有 _mailPending）', req0._mailPending === undefined);
  const d0 = await w0.qm.drainDue({});
  t('沒有 outbox：drainDue 回 disabled 摘要', d0.disabled === true && d0.sent === 0);
  t('沒有 outbox：pollDrain 回 false', (await w0.qm.pollDrain()) === false);
  eq('沒有 outbox：diagnostics', w0.qm.diagnostics().available, false);

  // off：儲存體完全不被碰
  const calls = [];
  const spy = new Proxy(memoryAdapter(), { get(target, prop) { const v = target[prop]; return typeof v === 'function' ? (...a) => { calls.push(String(prop)); return v.apply(target, a); } : v; } });
  const wo = mkWorld({ env: { MAIL_MODE: 'off' }, adapter: spy });
  calls.length = 0;
  const reqO = {};
  wo.qm.notifyMail(reqO, wo.ctx(), wo.quote, { type: 'E1_SUBMIT', idx: 0 });
  t('off：enabled=false', wo.qm.enabled() === false);
  t('off：notifyMail 不呼叫 dispatch（沒有 _mailPending、沒有傳輸、沒有儲存體操作）', reqO._mailPending === undefined && wo.sent.length === 0 && calls.length === 0, calls.join());
  t('off：pollDrain 回 false 且 runs 不增加', (await wo.qm.pollDrain()) === false && wo.qm.diagnostics().poll.runs === 0 && calls.length === 0);
  const dO = await wo.qm.drainDue({});
  t('off：drainDue 立刻返回（儲存體只可能被 breaker 讀取，不會有寫入）', dO.sent === 0 && !calls.some((c) => /insert|claim|requeue|purge|breakerSave/.test(c)), calls.join());
  t('off：沒有稽核', wo.logs.length === 0);
  // MAIL_MODE 未設／非法值 → off
  for (const raw of [undefined, '', 'true', '1', 'yes', 'LIVE!', 'ｌｉｖｅ', 'liv']) {
    const wx = mkWorld({ env: { MAIL_MODE: raw } });
    t('MAIL_MODE=' + short(raw) + ' → enabled=false', wx.qm.enabled() === false && wx.config.mode === 'off');
  }
  // 未 bind：不 throw、不寄
  const wn = mkWorld({ noBind: true });
  const reqN = {};
  wn.qm.notifyMail(reqN, wn.ctx(), wn.quote, { type: 'E1_SUBMIT', idx: 0 });
  t('未 bind：notifyMail 不 throw、不寄', reqN._mailPending === undefined && wn.sent.length === 0);
  t('未 bind：snapApproval 回 null、missingEmails 回空、pollDrain 回 false', wn.qm.snapApproval(wn.quote, wn.ctx()) === null && wn.qm.missingEmails().length === 0 && (await wn.qm.pollDrain()) === false);
  // bind 驗證
  let threw = 0;
  [undefined, null, {}, { getCfg, stepRecipients, isActive }, { getCfg, stepRecipients, isActive, dispName: 'x' }].forEach((h) => { try { w0.qm.bind(h); } catch (e) { threw++; } });
  eq('bind：缺任何一個輔助函式（或型別不對）都 throw', threw, 5);
});

// ═════════════════════════════════════════════════════════════════════════
section('3 notifyMail：E1／E3／E4／E2／E5／E6 的事件、收件人與信件內容', async () => {
  // ── E1：一般單，第一關 ──
  let w = mkWorld();
  let req = await fire(w, { type: 'E1_SUBMIT', idx: 0 });
  t('E1：寄給一級主管 1 封；操作者（業務）沒有', w.sent.length === 1 && toOf(w.sent[0]) === 'm1@itts.test', JSON.stringify(w.sent.map(toOf)));
  const e1 = w.sent[0];
  t('E1：主旨有單號與專案名稱，沒有金額／客戶名', e1.subject.includes('QU-261008-001') && e1.subject.includes(PROJECT) && !/NT\$|%|\d,\d{3}/.test(e1.subject) && !e1.subject.includes(COMPANY), e1.subject);
  t('E1：信內有金額（折扣後未稅）、毛利率、客戶名、業務暱稱、連結', textOf(e1).includes('NT$ 2,000,000') && textOf(e1).includes('10.00%') && textOf(e1).includes('折扣後未稅') && textOf(e1).includes(COMPANY) && textOf(e1).includes(OWNER_NICK) && textOf(e1).includes('https://crm.example.test/q/' + QID), '');
  let rows = (await w.outbox.list({})).rows;
  t('E1：outbox 一筆，dedupeKey＝E1_SUBMIT:<id>:m1:<submittedAt>#0，kind=mgr1，遮罩位址', rows.length === 1 && rows[0].dedupeKey === `E1_SUBMIT:${QID}:m1:${w.quote.approval.submittedAt}#0` && rows[0].meta.kind === 'mgr1' && rows[0].toMasked === 'm***@itts.test' && rows[0].status === 'sent', JSON.stringify(rows[0]));
  t('E1：同一個請求重複觸發 → 不重寄（去重）', (await fire(w, { type: 'E1_SUBMIT', idx: 0 }), w.sent.length === 1));
  t('E1：注入的 _mailPending 是 promise 陣列且已全部完成（永不 reject）', Array.isArray(req._mailPending) && req._mailPending.length === 1);
  t('E1：dispatch 後做了一次路由尾端清理（drainDue），沒有多寄', w.sent.length === 1);

  // ── E3：需總經理（2 位）、需董事長、董事會關 ──
  w = mkWorld();
  w.quote.approval.cur = 1; w.quote.approval.steps[0].status = 'approved'; w.quote.approval.steps[0].by = 'm1'; w.quote.approval.steps[1].status = 'pending';
  await fire(w, { type: 'E3_NEXT_STEP', idx: 1 }, 'm1');
  eq('E3 總經理關：名冊兩位都收到（gm1、gm2）', w.sent.map(toOf).sort(), ['gm1@itts.test', 'gm2@itts.test']);
  t('E3：信內有金額與毛利率、需總經理核准、客戶名', w.sent.every((m) => textOf(m).includes('NT$ 2,000,000') && textOf(m).includes('需總經理核准') && textOf(m).includes(COMPANY)));
  rows = (await w.outbox.list({})).rows;
  t('E3：stepKey＝submittedAt#1，kind=gm，level=2', rows.every((r) => /#1$/.test(r.dedupeKey) && r.meta.kind === 'gm' && r.meta.level === 2), JSON.stringify(rows.map((r) => r.meta)));

  const wb = mkWorld({ quote: mkQuote('BOARD') });
  wb.quote.approval.cur = 2; wb.quote.approval.steps[2].status = 'pending';
  await fire(wb, { type: 'E3_NEXT_STEP', idx: 2 }, 'gm1');
  eq('E3 董事會關：代核人（名冊）與秘書角色各一封；沒有 email 的秘書略過；管理員／停用不收', wb.sent.map(toOf).sort(), ['prx1@itts.test', 'sec1@itts.test']);
  const secM = wb.sent.find((m) => toOf(m) === 'sec1@itts.test');
  const prxM = wb.sent.find((m) => toOf(m) === 'prx1@itts.test');
  t('E3 董事會：秘書與代核人的信有金額與毛利率、需董事會決議，且不是長標籤', [secM, prxM].every((m) => textOf(m).includes('NT$ 60,000,000') && textOf(m).includes('40.00%') && textOf(m).includes('需董事會決議') && !textOf(m).includes('秘書代核）核准')));
  t('E3 董事會：秘書與代核人的信不帶客戶名、業務名（暱稱／顯示名稱／帳號）', [secM, prxM].every((m) => !textOf(m).includes(COMPANY) && !textOf(m).includes(OWNER_NICK) && !textOf(m).includes('SALES1 Name') && !textOf(m).includes('sales1')));
  t('E3 董事會：秘書與代核人的信不帶專案名稱（主旨只有類別與單號；html、純文字、preheader 都沒有）', [secM, prxM].every((m) => !textOf(m).includes(PROJECT) && !textOf(m).includes('專案名稱') && m.subject === '【簽核通知】QU-261008-001'), secM.subject);
  rows = (await wb.outbox.list({})).rows;
  eq('E3 董事會：outbox kind（秘書＝secretary、代核人＝boardProxy；無 email 的秘書 skipped／NO_EMAIL）', rows.map((r) => [r.toUser, r.meta.kind, r.status, r.skipReason]).sort(), [['noemail', 'secretary', 'skipped', 'NO_EMAIL'], ['prx1', 'boardProxy', 'sent', null], ['sec1', 'secretary', 'sent', null]]);
  const sk = wb.logs.filter((l) => l.action === 'QUOTE_MAIL_SKIPPED');
  t('E3 董事會：稽核 QUOTE_MAIL_SKIPPED 一筆（NO_EMAIL），沒有位址／金額／客戶名', sk.length === 1 && /reason=NO_EMAIL/.test(sk[0].detail) && /to=noemail/.test(sk[0].detail) && !FULL_EMAIL_RE.test(sk[0].detail) && !/NT\$|%/.test(sk[0].detail) && !sk[0].detail.includes(COMPANY), JSON.stringify(sk));

  // ── E4 ──
  w = mkWorld();
  w.quote.approval.cur = 1; w.quote.approval.steps[0].status = 'approved'; w.quote.approval.steps[0].by = 'm1'; w.quote.approval.steps[1].status = 'pending';
  await fire(w, { type: 'E4_RESULT', resultKind: 'approved', idx: 0 }, 'm1');
  t('E4 approved：只寄給業務；信上沒有金額與毛利', w.sent.length === 1 && toOf(w.sent[0]) === 'sales1@itts.test' && !/NT\$|毛利/.test(textOf(w.sent[0])), JSON.stringify(w.sent.map(toOf)));
  rows = (await w.outbox.list({})).rows;
  t('E4：stepKey＝submittedAt#r:approved:0，kind=owner', rows.length === 1 && rows[0].dedupeKey.endsWith(`${w.quote.approval.submittedAt}#r:approved:0`) && rows[0].meta.kind === 'owner', rows[0] && rows[0].dedupeKey);
  w = mkWorld();
  w.quote.approval.steps[0].status = 'returned'; w.quote.approval.steps[0].by = 'm1'; w.quote.approval.steps[0].comment = 'bad <b>scope</b>'; w.quote.approval.state = 'returned';
  await fire(w, { type: 'E4_RESULT', resultKind: 'rejected', idx: 0, reason: 'bad <b>scope</b>' }, 'm1');
  t('E4 rejected：原因進信件且 HTML 被跳脫', w.sent.length === 1 && w.sent[0].html.includes('bad &lt;b&gt;scope&lt;/b&gt;') && !w.sent[0].html.includes('<b>scope</b>') && w.sent[0].text.includes('bad <b>scope</b>'), '');
  w = mkWorld();
  w.quote.approval.cur = 1; w.quote.approval.steps[0].status = 'approved'; w.quote.approval.steps[0].by = 'm1'; w.quote.approval.steps[1].status = 'pending';
  await fire(w, { type: 'E4_RESULT', resultKind: 'approved', idx: 0, reason: 'secret approval note' }, 'm1');
  t('E4 approved：就算傳了 reason 也不會放進信（只有 rejected 才帶原因）', w.sent.length === 1 && !textOf(w.sent[0]).includes('secret approval note'));
  w = mkWorld({ quote: mkQuote('BOARD') });
  w.quote.approval.cur = 2; w.quote.approval.steps[0].status = 'approved'; w.quote.approval.steps[0].by = 'm1'; w.quote.approval.steps[1].status = 'approved'; w.quote.approval.steps[1].by = 'gm1'; w.quote.approval.steps[2].status = 'pending';
  await fire(w, { type: 'E4_RESULT', resultKind: 'approved', idx: 0 }, 'm1');
  await fire(w, { type: 'E4_RESULT', resultKind: 'approved', idx: 1 }, 'gm1');
  t('E4：同一次送簽、兩個關卡各簽一次 → 業務收到兩封（去重鍵含關卡序號）', w.sent.length === 2 && (await w.outbox.list({})).rows.length === 2, w.sent.length);
  {
    let drains = 0;
    const spy = new Proxy(memoryAdapter(), { get(target, prop) { const v = target[prop]; return typeof v === 'function' ? (...a) => { if (prop === 'expireExhausted') drains++; return v.apply(target, a); } : v; } });
    const wd = mkWorld({ adapter: spy });
    const rq = {};
    wd.qm.notifyMail(rq, wd.ctx('m1'), wd.quote, { type: 'E4_RESULT', resultKind: 'approved', idx: 0 });
    wd.qm.notifyMail(rq, wd.ctx('m1'), wd.quote, { type: 'E3_NEXT_STEP', idx: 1 });
    await Promise.all(rq._mailPending);
    t('同一個請求內兩次 notifyMail（核准同時寄 E4 與 E3）→ 路由尾端只清理一次', rq._mailPending.length === 2 && drains === 1, 'pending=' + rq._mailPending.length + ' drains=' + drains);
  }
  w = mkWorld();
  await fire(w, { type: 'E4_RESULT', resultKind: 'bogus', idx: 0 }, 'm1');
  t('E4 未知 resultKind → 事件不合法 → 不寄（dispatcher 記 BAD_EVENT，永不 throw）', w.sent.length === 0);

  // ── E2 ──
  w = mkWorld();
  await fire(w, { type: 'E2_COST_REQUEST' }, 'sales1');
  t('E2：只寄給顧問；品項說明與數量與單位在信內', w.sent.length === 1 && toOf(w.sent[0]) === 'cons1@itts.test' && textOf(w.sent[0]).includes('Install service') && textOf(w.sent[0]).includes('Training') && textOf(w.sent[0]).includes('set'), '');
  t('E2：沒有任何價格／成本／金額／毛利／折扣、沒有客戶名', !/98765|98,765|12345|12,345|54321|54,321|6789|6,789|NT\$|毛利|折扣/.test(textOf(w.sent[0])) && !textOf(w.sent[0]).includes(COMPANY), '');
  t('E2：連結帶 ?cost=1', textOf(w.sent[0]).includes('https://crm.example.test/q/' + QID + '?cost=1'));
  rows = (await w.outbox.list({})).rows;
  t('E2：stepKey＝requestedAt#<顧問>，kind=consultant', rows.length === 1 && rows[0].dedupeKey.endsWith(`${w.quote.costFlow.requestedAt}#cons1`) && rows[0].meta.kind === 'consultant', rows[0] && rows[0].dedupeKey);
  const q2 = w.quote;
  q2.costBy = 'cons2'; q2.costFlow.requestedAt = iso(T0 + 1000);
  await fire(w, { type: 'E2_COST_REQUEST' }, 'sales1');
  t('E2：換顧問（新的 requestedAt）→ 新顧問再收到一封', w.sent.length === 2 && toOf(w.sent[1]) === 'cons2@itts.test');
  q2.costBy = 'cons1'; q2.costFlow.requestedAt = iso(T0 + 2000);
  await fire(w, { type: 'E2_COST_REQUEST' }, 'sales1');
  t('E2：換回原顧問（新的 requestedAt）→ 再收到一封，不會被第一次的紀錄擋掉', w.sent.length === 3 && toOf(w.sent[2]) === 'cons1@itts.test', JSON.stringify(w.sent.map(toOf)));
  q2.costFlow.requestedAt = null;
  await fire(w, { type: 'E2_COST_REQUEST' }, 'sales1');
  t('E2：沒有 requestedAt → 不寄（不產生壞的去重鍵）', w.sent.length === 3);
  q2.costBy = null;
  await fire(w, { type: 'E2_COST_REQUEST' }, 'sales1');
  t('E2：沒有顧問 → 不寄', w.sent.length === 3);

  // ── E5 ──
  w = mkWorld();
  w.quote.costFlow = { state: 'filled', requestedAt: iso(T0 - 5000), filledAt: iso(T0 + 2000) };
  await fire(w, { type: 'E5_COST_DONE' }, 'cons1');
  t('E5：只寄給業務；沒有金額與毛利', w.sent.length === 1 && toOf(w.sent[0]) === 'sales1@itts.test' && !/NT\$|毛利/.test(textOf(w.sent[0])));
  rows = (await w.outbox.list({})).rows;
  t('E5：stepKey＝filledAt#done，kind=owner', rows.length === 1 && rows[0].dedupeKey.endsWith(`${iso(T0 + 2000)}#done`) && rows[0].meta.kind === 'owner');
  w.quote.costFlow.filledAt = null;
  await fire(w, { type: 'E5_COST_DONE' }, 'cons1');
  t('E5：沒有 filledAt → 不寄', w.sent.length === 1);

  // ── E6：撤回（快照在清空 steps 之前拍）──
  w = mkWorld();
  w.quote.approval.cur = 1; w.quote.approval.steps[0].status = 'approved'; w.quote.approval.steps[0].by = 'm1'; w.quote.approval.steps[1].status = 'pending';
  const snap = w.qm.snapApproval(w.quote, w.ctx());
  eq('snapApproval：submittedAt、原關卡、目前關卡收件人、已簽過的人', snap, {
    submittedAt: w.quote.approval.submittedAt, originStep: { level: 2, label: '總經理' },
    current: [{ username: 'gm1', tier: 'gm' }, { username: 'gm2', tier: 'gm' }], signers: [{ username: 'm1', tier: 'mgr1' }],
  });
  const savedAt = w.quote.approval.submittedAt;
  w.quote.approval.state = 'none'; w.quote.approval.steps = []; w.quote.approval.derived = null;      // 撤回：steps 與 derived 清空
  await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'withdrawn', snap }, 'sales1');
  eq('E6 撤回：目前關卡（總經理兩位）＋已簽過的一級主管；業務（操作者）沒有', w.sent.map(toOf).sort(), ['gm1@itts.test', 'gm2@itts.test', 'm1@itts.test']);
  t('E6：主旨含「請勿簽核」，內文說明已撤回，沒有金額', w.sent.every((m) => m.subject.includes('請勿簽核') && /已撤回/.test(textOf(m)) && !/NT\$|毛利/.test(textOf(m))));
  t('E6：帶原關卡（總經理）', w.sent.every((m) => textOf(m).includes('總經理')));
  rows = (await w.outbox.list({})).rows;
  t('E6：stepKey＝submittedAt#withdrawn；kind 依各人的關卡（gm／gm／mgr1）', rows.length === 3 && rows.every((r) => r.dedupeKey.endsWith(`${savedAt}#withdrawn`)) && rows.filter((r) => r.meta.kind === 'gm').length === 2 && rows.filter((r) => r.meta.kind === 'mgr1').length === 1, JSON.stringify(rows.map((r) => r.meta)));
  // E6 作廢：核准完成後修改；包含業務本人（操作者是別人時才會收到）
  w = mkWorld();
  w.quote.approval.state = 'approved'; w.quote.approval.cur = 2; w.quote.approval.steps.forEach((s, i) => { s.status = 'approved'; s.by = i === 0 ? 'm1' : 'gm1'; });
  const snapV = w.qm.snapApproval(w.quote, w.ctx());
  eq('snapApproval（已核准）：沒有目前關卡收件人；原關卡＝最後一關', [snapV.current.length, snapV.originStep.label, snapV.signers.length], [0, '總經理', 2]);
  const voidAt = w.quote.approval.submittedAt;
  w.quote.approval = { state: 'none', submittedAt: null, submittedBy: null, derived: null, steps: [], cur: 0, history: [] };      // quoteRoutes 的作廢：整個 approval 換掉，submittedAt 變 null
  await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'voided', snap: snapV }, 'sales1');
  eq('E6 作廢：已簽過的人（操作者業務自己沒有）', w.sent.map(toOf).sort(), ['gm1@itts.test', 'm1@itts.test']);
  t('E6 作廢：去重鍵用作廢前的 submittedAt（approval 已被清空仍然正確）', (await w.outbox.list({})).rows.every((r) => r.dedupeKey.endsWith(voidAt + '#voided')), JSON.stringify((await w.outbox.list({})).rows.map((r) => r.dedupeKey)));
  t('E6 作廢：內文說明核准已作廢', w.sent.every((m) => /已作廢/.test(textOf(m))));
  w = mkWorld();
  w.quote.approval.state = 'approved'; w.quote.approval.cur = 2; w.quote.approval.steps.forEach((s, i) => { s.status = 'approved'; s.by = i === 0 ? 'm1' : 'gm1'; });
  const snapV2 = w.qm.snapApproval(w.quote, w.ctx('adm1'));
  w.quote.approval = { state: 'none', submittedAt: null, submittedBy: null, derived: null, steps: [], cur: 0, history: [] };      // 真實路由：作廢＝approval 整個重置
  await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'voided', snap: snapV2 }, 'adm1');
  eq('E6 作廢（別人操作）：業務本人也會收到', w.sent.map(toOf).sort(), ['gm1@itts.test', 'm1@itts.test', 'sales1@itts.test']);
  // E6（董事會關）撤回：秘書與董事會代核人不帶專案名稱，其餘簽核人照舊；E6 事件沒有 reason（自由文字不是專案名稱的來源）
  {
    const wq = mkWorld({ quote: mkQuote('BOARD') });
    const ap6 = wq.quote.approval;
    ap6.cur = 2; ap6.steps[0].status = 'approved'; ap6.steps[0].by = 'm1'; ap6.steps[1].status = 'approved'; ap6.steps[1].by = 'gm1'; ap6.steps[2].status = 'pending';
    const snap6 = wq.qm.snapApproval(wq.quote, wq.ctx());
    ap6.state = 'none'; ap6.steps = []; ap6.derived = null;
    await fire(wq, { type: 'E6_WITHDRAWN', resultKind: 'withdrawn', snap: snap6 }, 'sales1');
    const hid = wq.sent.filter((m) => ['sec1@itts.test', 'prx1@itts.test'].indexOf(toOf(m)) >= 0);
    const shown = wq.sent.filter((m) => ['m1@itts.test', 'gm1@itts.test'].indexOf(toOf(m)) >= 0);
    eq('E6 董事會關撤回：秘書、代核人、已簽過的一級主管與總經理各一封', wq.sent.map(toOf).sort(), ['gm1@itts.test', 'm1@itts.test', 'prx1@itts.test', 'sec1@itts.test']);
    t('E6 董事會關撤回：秘書與代核人的信不帶專案名稱（主旨＝【簽核撤回】單號 請勿簽核）', hid.length === 2 && hid.every((m) => !textOf(m).includes(PROJECT) && m.subject === '【簽核撤回】QU-261008-001 請勿簽核'), JSON.stringify(hid.map((m) => m.subject)));
    t('E6 董事會關撤回：一級主管與總經理的信照舊含專案名稱', shown.length === 2 && shown.every((m) => m.subject.includes(PROJECT) && textOf(m).includes(PROJECT)), JSON.stringify(shown.map((m) => m.subject)));
    t('E6 撤回信沒有「原因」欄（事件不帶 reason）', wq.sent.every((m) => !/原因/.test(textOf(m))));
    const rowsQ = (await wq.outbox.list({})).rows;
    t('E6：outbox 紀錄與稽核沒有專案名稱', JSON.stringify(rowsQ).indexOf(PROJECT) < 0 && JSON.stringify(wq.logs).indexOf(PROJECT) < 0);
  }
  w = mkWorld();
  await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'withdrawn', snap: null }, 'sales1');
  await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'withdrawn', snap: { submittedAt: '' , current: [], signers: [] } }, 'sales1');
  t('E6：沒有快照（或快照沒有 submittedAt）→ 不寄', w.sent.length === 0);

  // ── 收件人過濾 ──
  w = mkWorld();
  w.quote.approval.steps[0].assignee = 'off1';
  await fire(w, { type: 'E1_SUBMIT', idx: 0 }, 'sales1');
  t('收件人是停用帳號 → 不入列、不寄（與 notify() 的過濾相同）', w.sent.length === 0 && (await w.outbox.list({})).rows.length === 0);
  w = mkWorld();
  w.quote.approval.steps[0].assignee = 'sales1';
  await fire(w, { type: 'E1_SUBMIT', idx: 0 }, 'sales1');
  t('收件人就是操作者 → 排除', w.sent.length === 0);
  w = mkWorld();
  w.roster.gm = ['gm1', 'gm1', 'gm2', '', null];
  w.quote.approval.cur = 1; w.quote.approval.steps[1].status = 'pending';
  await fire(w, { type: 'E3_NEXT_STEP', idx: 1 }, 'm1');
  eq('名冊有重複與空值 → 去重、略過空值', w.sent.map(toOf).sort(), ['gm1@itts.test', 'gm2@itts.test']);
  w = mkWorld();
  w.users.find((x) => x.username === 'm1').email = 'm1@example.com';
  await fire(w, { type: 'E1_SUBMIT', idx: 0 });
  t('收件人網域不在白名單 → skipped／DOMAIN_NOT_ALLOWED，沒有寄', w.sent.length === 0 && (await w.outbox.list({})).rows[0].skipReason === 'DOMAIN_NOT_ALLOWED');
});

// ═════════════════════════════════════════════════════════════════════════
section('4 永不 throw、不影響簽核', async () => {
  const w = mkWorld();
  const bad = [undefined, null, 0, 'x', [], {}, { type: 'NOPE' }, { type: 'E1_SUBMIT' }, { type: 'E1_SUBMIT', idx: 99 }, { type: 'E1_SUBMIT', idx: -1 }, { type: 'E1_SUBMIT', idx: 1.5 }, { type: 'E4_RESULT' }, { type: 'E6_WITHDRAWN' }];
  let threw = 0;
  for (const spec of bad) { try { w.qm.notifyMail({}, w.ctx(), w.quote, spec); } catch (e) { threw++; } }
  [[undefined, undefined, undefined], [{}, null, w.quote], [{}, w.ctx(), null], [{}, w.ctx(), {}], [null, w.ctx(), w.quote]].forEach(([r, c, q]) => { try { w.qm.notifyMail(r, c, q, { type: 'E1_SUBMIT', idx: 0 }); } catch (e) { threw++; } });
  eq('各種亂七八糟的參數：notifyMail 都不 throw', threw, 0);
  await sleep(50);
  t('亂參數沒有造成任何寄信（除了合法的 E1 兩次：同一事件去重）', w.sent.length <= 1, w.sent.length);

  // 輔助函式爆炸
  const w2 = mkWorld();
  w2.qm.bind({ getCfg, isActive, dispName, stepRecipients: () => { throw new Error('boom'); } });
  const r2 = {};
  let th = false;
  try { w2.qm.notifyMail(r2, w2.ctx(), w2.quote, { type: 'E1_SUBMIT', idx: 0 }); } catch (e) { th = true; }
  t('stepRecipients 丟例外：notifyMail 不 throw、不寄', !th && w2.sent.length === 0);
  eq('stepRecipients 丟例外：snapApproval 回 null（不 throw）', w2.qm.snapApproval(w2.quote, w2.ctx()), null);

  // 傳輸失敗不影響 notifyMail 的 promise（永不 reject）
  const w3 = mkWorld();
  w3.behavior = () => { throw new Error('transport exploded'); };
  const r3 = await fire(w3, { type: 'E1_SUBMIT', idx: 0 });
  const settled = await Promise.allSettled(r3._mailPending);
  t('傳輸丟例外：_mailPending 全部 fulfilled（server.js 的 res.end 包裝不會被卡住）', settled.every((s) => s.status === 'fulfilled'), JSON.stringify(settled.map((s) => s.status)));
  const rows3 = (await w3.outbox.list({})).rows;
  t('傳輸丟例外：工作留在 outbox 等重試（pending，有 nextAttemptAt）', rows3.length === 1 && rows3[0].status === 'pending' && !!rows3[0].nextAttemptAt, JSON.stringify(rows3[0]));
  // 儲存體壞掉
  const bad4 = { insert() { throw new Error('disk gone'); }, claimDue() { throw new Error('disk gone'); }, expireExhausted() { throw new Error('disk gone'); }, get() { throw new Error('x'); }, list() { throw new Error('x'); }, stats() { throw new Error('x'); }, breakerLoad() { throw new Error('x'); }, breakerSave() { throw new Error('x'); } };
  const w4 = mkWorld({ adapter: bad4 });
  const r4 = await fire(w4, { type: 'E1_SUBMIT', idx: 0 });
  const s4 = await Promise.allSettled(r4._mailPending);
  t('儲存體壞掉：不 throw、_mailPending 全部 fulfilled、沒寄', s4.every((s) => s.status === 'fulfilled') && w4.sent.length === 0);
});

// ═════════════════════════════════════════════════════════════════════════
section('5 isStillValid／rebuild（drainDue 的契約）', async () => {
  const w = mkWorld();
  await fire(w, { type: 'E1_SUBMIT', idx: 0 });
  const job1 = (await w.outbox.list({})).rows[0];
  t('E1：單據仍在這一關 → 有效', w.qm.isStillValid(job1) === true);
  const ap = w.quote.approval;
  ap.cur = 1; ap.steps[0].status = 'approved'; ap.steps[1].status = 'pending';
  t('E1：單據已走到下一關 → 失效', w.qm.isStillValid(job1) === false);
  ap.cur = 0; ap.steps[1].status = 'waiting';
  ap.steps[0].assignee = 'gm1';
  t('E1：這一關的收件人已被改派 → 失效', w.qm.isStillValid(job1) === false);
  ap.steps[0].assignee = 'm1';
  ap.submittedAt = iso(T0 + 99999);
  t('E1：重新送簽（submittedAt 不同）→ 失效', w.qm.isStillValid(job1) === false);
  ap.submittedAt = iso(T0);
  ap.state = 'none';
  t('E1：已撤回／作廢（state≠pending）→ 失效', w.qm.isStillValid(job1) === false);
  ap.state = 'pending';
  w.data.quotations = [];
  t('單據不存在 → 失效', w.qm.isStillValid(job1) === false);
  w.data.quotations = [w.quote];
  t('E1：恢復後又有效', w.qm.isStillValid(job1) === true);
  t('壞掉的 stepKey（沒有 #序號）→ 失效', w.qm.isStillValid(Object.assign({}, job1, { dedupeKey: job1.dedupeKey.replace(/#0$/, '#x') })) === false);

  // E2
  const w2 = mkWorld();
  await fire(w2, { type: 'E2_COST_REQUEST' });
  const job2 = (await w2.outbox.list({})).rows[0];
  t('E2：成本仍是 requested、顧問沒換 → 有效', w2.qm.isStillValid(job2) === true);
  w2.quote.costFlow.state = 'filled';
  t('E2：顧問已完成 → 失效', w2.qm.isStillValid(job2) === false);
  w2.quote.costFlow.state = 'requested'; w2.quote.costBy = 'cons2';
  t('E2：換了顧問 → 失效', w2.qm.isStillValid(job2) === false);
  w2.quote.costBy = 'cons1'; w2.quote.costFlow.requestedAt = iso(T0 + 7777);
  t('E2：重新請求（requestedAt 不同）→ 舊的失效', w2.qm.isStillValid(job2) === false);
  // E4／E5／E6 的過期判斷改成檢查單據現況（FIX-2；逐情境的窮舉在 5b）。這裡只留「壞掉的 stepKey 一律失效」
  for (const type of ['E4_RESULT', 'E5_COST_DONE', 'E6_WITHDRAWN']) t(type + '：壞掉的 stepKey → 失效（不寄）', w2.qm.isStillValid({ type, quoteId: QID, dedupeKey: 'x', toUser: 'sales1' }) === false);

  // rebuild：與首次寄送逐字相同（at＝工作建立時間）
  const cases = [
    ['E1_SUBMIT', { type: 'E1_SUBMIT', idx: 0 }, 'sales1', (x) => x],
    ['E3_NEXT_STEP', { type: 'E3_NEXT_STEP', idx: 1 }, 'm1', (x) => { x.approval.cur = 1; x.approval.steps[0].status = 'approved'; x.approval.steps[0].by = 'm1'; x.approval.steps[1].status = 'pending'; return x; }],
    ['E4_RESULT', { type: 'E4_RESULT', resultKind: 'rejected', idx: 0, reason: 'no <good>' }, 'm1', (x) => { x.approval.steps[0].status = 'returned'; x.approval.steps[0].by = 'm1'; x.approval.steps[0].comment = 'no <good>'; x.approval.state = 'returned'; return x; }],
    ['E4_RESULT(approved)', { type: 'E4_RESULT', resultKind: 'approved', idx: 0 }, 'm1', (x) => { x.approval.steps[0].status = 'approved'; x.approval.steps[0].by = 'm1'; x.approval.steps[0].comment = 'internal approval note'; x.approval.cur = 1; x.approval.steps[1].status = 'pending'; return x; }],
    ['E2_COST_REQUEST', { type: 'E2_COST_REQUEST' }, 'sales1', (x) => x],
    ['E5_COST_DONE', { type: 'E5_COST_DONE' }, 'cons1', (x) => { x.costFlow = { state: 'filled', requestedAt: iso(T0 - 5000), filledAt: iso(T0 + 3000) }; return x; }],
  ];
  for (const [type, spec, me, prep] of cases) {
    const x = mkWorld({ quote: prep(mkQuote()) });
    x.clock.t = T0 + 12345;                                   // 事件發生時間＝工作建立時間（outbox 用同一個時鐘）
    await fire(x, spec, me);
    t(type + '：首次寄送成功（' + x.sent.length + ' 封）', x.sent.length >= 1);
    const job = (await x.outbox.list({})).rows[0];
    x.clock.t += 120000;                                      // 兩分鐘後重試：內容不可因為時間而不同
    const rb = x.qm.rebuild(job);
    t(type + '：rebuild 回 {ev}，stepKey 與工作相同、at＝工作建立時間', rb && rb.ev && rb.ev.stepKey === job.dedupeKey.split(':').slice(3).join(':') && rb.ev.at === job.createdAt, rb && JSON.stringify(rb.ev).slice(0, 160));
    const RND = load('lib/mail/render.js');
    const viewer = { username: job.toUser, label: dispNameFor(x, job.toUser), kind: rb.kind || job.meta.kind };
    const re = RND.renderMail(rb.ev, viewer, { config: x.config, now: x.clock.t });
    const reKey = x.sent.find((m) => m.to[0] === job.toUser + '@itts.test');
    t(type + '：重建的信與首次寄出的逐字相同（主旨、HTML、純文字）', !!reKey && re.subject === reKey.subject && re.html === reKey.html && re.text === reKey.text, '');
  }
  // E6：重建沒有「原關卡」那一列，其餘相同
  const w6 = mkWorld();
  const snap6 = w6.qm.snapApproval(w6.quote, w6.ctx());
  const at6 = w6.quote.approval.submittedAt;
  w6.quote.approval.state = 'none'; w6.quote.approval.steps = []; w6.quote.approval.derived = null;
  await fire(w6, { type: 'E6_WITHDRAWN', resultKind: 'withdrawn', snap: snap6 }, 'sales1');
  const job6 = (await w6.outbox.list({})).rows[0];
  const rb6 = w6.qm.rebuild(job6);
  t('E6：撤回後 rebuild 仍可重建（steps 已清空，沒有原關卡）', rb6 && rb6.ev && rb6.ev.result.kind === 'withdrawn' && rb6.ev.step === undefined && rb6.ev.stepKey === at6 + '#withdrawn', rb6 && JSON.stringify(rb6.ev).slice(0, 200));
  // 各種「該取消」的情況
  const wg = mkWorld();
  await fire(wg, { type: 'E1_SUBMIT', idx: 0 });
  const jg = (await wg.outbox.list({})).rows[0];
  wg.quote.approval.steps = [];
  t('E1：該關資料已清空 → rebuild 回 null（取消 GONE）', wg.qm.rebuild(jg) === null);
  wg.data.quotations = [];
  t('單據不存在 → rebuild 回 null', wg.qm.rebuild(jg) === null);
  t('未知的工作型別 → rebuild 回 null', wg.qm.rebuild({ type: 'E9_X', quoteId: QID, dedupeKey: 'a:b:c:d', toUser: 'x', createdAt: iso(T0) }) === null);
  t('E4 壞掉的 stepKey → rebuild 回 null', mkWorld().qm.rebuild({ type: 'E4_RESULT', quoteId: QID, dedupeKey: 'E4_RESULT:' + QID + ':sales1:garbage', toUser: 'sales1', createdAt: iso(T0) }) === null);
  t('E6 壞掉的 stepKey → rebuild 回 null', mkWorld().qm.rebuild({ type: 'E6_WITHDRAWN', quoteId: QID, dedupeKey: 'E6_WITHDRAWN:' + QID + ':sales1:garbage', toUser: 'sales1', createdAt: iso(T0) }) === null);
  function dispNameFor(x, un) { return dispName(x.userMap(), un); }
});


// ═════════════════════════════════════════════════════════════════════════
// FIX-2：過期判斷（lib/mail/validity.js 的 checkValidity，經 quoteMail.isStillValid 走真的 liveCtx）
// 以下 helper 模擬 lib/quoteRoutes.js 各路由「寫進單據」的狀態變化（欄位與真實路由一致：撤回保留 submittedAt、作廢把 approval 整個重置、駁回 state=returned…）
// ═════════════════════════════════════════════════════════════════════════
const histPush = (w, action, atMs, by, meta) => {
  const ap = w.quote.approval;
  if (!Array.isArray(ap.history)) ap.history = [];
  const h = { at: iso(atMs), by: by || 'sales1', action, comment: '', tier: '' };
  if (meta) h.meta = meta;
  ap.history.push(h);
};
/** 管理員改派一級主管（POST /reassign）：assignee 換人、歷史多一筆 REASSIGN；路由把改派前後的承辦人記在 meta {from, to}（validity 靠它辨識「改派給同一人」） */
function reassignTo(w, to, atMs) {
  const st = w.quote.approval.steps[0];
  const from = st.assignee;
  st.assignee = to;
  histPush(w, 'REASSIGN', atMs, 'adm1', { from, to });
}
/** 業務（重新）送簽：新的 approval 物件（pending、cur 0、新的 submittedAt、history 保留）— 與 POST /submit 相同 */
function newRound(w, atMs, derivedKey) {
  const dd = DERIVED[derivedKey || 'L2'];
  const hist = ((w.quote.approval || {}).history || []).slice();
  hist.push({ at: iso(atMs), by: 'sales1', action: 'SUBMIT', comment: '', tier: '' });
  w.quote.approval = {
    state: 'pending', submittedAt: iso(atMs), submittedBy: 'sales1', cur: 0, derived: dd, board: null, history: hist,
    steps: dd.tiers.map((tier, i) => ({ tier, label: STEP_LABEL[tier], assignee: tier === 'mgr1' ? 'm1' : null, status: i === 0 ? 'pending' : 'waiting', by: null, comment: '' })),
  };
}
/** 撤回：steps 清空、state=none，但 submittedAt 保留（POST /withdraw 的行為） */
function withdrawQuote(w, atMs) {
  const ap = w.quote.approval;
  ap.state = 'none'; ap.steps = []; ap.cur = 0; ap.hash = null; ap.derived = null; ap.board = null;
  histPush(w, 'WITHDRAW', atMs);
}
/** 核准後修改使核准作廢：approval 整個重置，submittedAt=null（PUT /:id 的 confirmVoid） */
function voidQuote(w, atMs) {
  const hist = (w.quote.approval.history || []).slice();
  w.quote.approval = { state: 'none', submittedAt: null, submittedBy: null, derived: null, steps: [], cur: 0, board: null, history: hist };
  histPush(w, 'INVALIDATE', atMs);
}
/** 第 i 關核准（POST /approve）：最後一關 → state=approved、cur=steps.length；否則 cur 前進、下一關 pending */
function approveAt(w, i, by) {
  const ap = w.quote.approval;
  ap.steps[i].status = 'approved'; ap.steps[i].by = by;
  if (i >= ap.steps.length - 1) { ap.state = 'approved'; ap.cur = ap.steps.length; } else { ap.cur = i + 1; ap.steps[ap.cur].status = 'pending'; }
}
/** 第 i 關駁回（POST /return）：state=returned */
function rejectAt(w, i, by) {
  const ap = w.quote.approval;
  ap.steps[i].status = 'returned'; ap.steps[i].by = by; ap.steps[i].comment = 'no good'; ap.state = 'returned';
}
const lastRow = async (w, type, toUser) => { const rows = (await w.outbox.list({ type, toUser })).rows; return rows[rows.length - 1]; };

section('5b 過期判斷（FIX-2）：E1～E6 各情境（重送簽後、再核准後、作廢後、改派後…）', async () => {
  const VC = load('lib/mail/validity.js');
  t('validity.js 匯出 checkValidity／stepKeyOf／stepIdxOf／reassignEpoch／e1StepKey／VALIDITY_CODES', ['checkValidity', 'stepKeyOf', 'stepIdxOf', 'reassignEpoch', 'e1StepKey'].every((k) => typeof VC[k] === 'function') && Object.isFrozen(VC.VALIDITY_CODES));
  let n = 0;
  /** variants: [名稱, 變化(w), 預期是否仍有效]；每個變化都在全新的世界上做 */
  async function matrix(title, build, variants) {
    for (const [name, mutate, want] of variants) {
      const { w, job } = await build();
      mutate(w);
      let got;
      try { got = w.qm.isStillValid(job); } catch (e) { got = 'throw:' + e.message; }
      n += 1;
      t(title + '｜' + name + ' → ' + (want ? '仍有效（會寄）' : '過期（不寄）'), got === want, 'got=' + got + ' code=' + (() => { try { return w.qm.checkJob(job).code; } catch (e) { return '?'; } })());
    }
  }

  // ── E4 本關通過（BOARD 三關：m1 → gm1 → 董事會）──
  await matrix('E4 本關通過(idx0)', async () => {
    const w = mkWorld({ quote: mkQuote('BOARD') });
    approveAt(w, 0, 'm1');
    await fire(w, { type: 'E4_RESULT', resultKind: 'approved', idx: 0 }, 'm1');
    return { w, job: await lastRow(w, 'E4_RESULT', 'sales1') };
  }, [
    ['沒有變化', () => {}, true],
    ['下一關也通過了（單據仍在簽核中、cur 越過該關）', (w) => approveAt(w, 1, 'gm1'), true],
    ['全部關卡簽完（已 final approved，這封「將送往下一關」已被取代）', (w) => { approveAt(w, 1, 'gm1'); approveAt(w, 2, 'sec1'); }, false],
    ['下一關駁回', (w) => rejectAt(w, 1, 'gm1'), false],
    ['業務撤回', (w) => withdrawQuote(w, T0 + 5000), false],
    ['業務撤回後重新送簽', (w) => { withdrawQuote(w, T0 + 5000); newRound(w, T0 + 9000, 'BOARD'); }, false],
    ['全部簽完後作廢', (w) => { approveAt(w, 1, 'gm1'); approveAt(w, 2, 'sec1'); voidQuote(w, T0 + 5000); }, false],
    ['作廢後重新送簽（新的一輪、第 0 關 pending）', (w) => { approveAt(w, 1, 'gm1'); approveAt(w, 2, 'sec1'); voidQuote(w, T0 + 5000); newRound(w, T0 + 9000, 'BOARD'); }, false],
    ['該關的狀態被改回 pending（資料異常）', (w) => { w.quote.approval.steps[0].status = 'pending'; }, false],
    ['單據的業務換人', (w) => { w.quote.owner = 'cons2'; }, false],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);

  // ── E4 最終核准（BOARD：三關全簽）──
  await matrix('E4 最終核准', async () => {
    const w = mkWorld({ quote: mkQuote('BOARD') });
    approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); approveAt(w, 2, 'sec1');
    await fire(w, { type: 'E4_RESULT', resultKind: 'final_approved', idx: 2 }, 'sec1');
    return { w, job: await lastRow(w, 'E4_RESULT', 'sales1') };
  }, [
    ['沒有變化', () => {}, true],
    ['同一輪但 state 被改回 pending（資料異常；關卡仍是 approved）', (w) => { w.quote.approval.state = 'pending'; }, false],
    ['核准後修改 → 作廢（approval 重置、submittedAt=null）', (w) => voidQuote(w, T0 + 5000), false],
    ['作廢後重新送簽', (w) => { voidQuote(w, T0 + 5000); newRound(w, T0 + 9000, 'BOARD'); }, false],
    ['作廢、重新送簽、又全部簽完（新的一輪 approved）', (w) => { voidQuote(w, T0 + 5000); newRound(w, T0 + 9000, 'BOARD'); approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); approveAt(w, 2, 'sec1'); }, false],
    ['業務換人', (w) => { w.quote.owner = 'cons2'; }, false],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);

  // ── E4 駁回 ──
  await matrix('E4 駁回', async () => {
    const w = mkWorld();
    rejectAt(w, 0, 'm1');
    await fire(w, { type: 'E4_RESULT', resultKind: 'rejected', idx: 0, reason: 'no good' }, 'm1');
    return { w, job: await lastRow(w, 'E4_RESULT', 'sales1') };
  }, [
    ['沒有變化', () => {}, true],
    ['業務修改內容但還沒重新送簽（state 仍是 returned）', (w) => { w.quote.projectName = 'edited'; }, true],
    ['同一輪但 state 被改成 pending（資料異常；關卡仍是 returned）', (w) => { w.quote.approval.state = 'pending'; }, false],
    ['同一輪但 state 被改成 approved（資料異常）', (w) => { w.quote.approval.state = 'approved'; }, false],
    ['業務重新送簽（X1：重試時業務已重送，不可再收到「已被駁回」）', (w) => newRound(w, T0 + 9000), false],
    ['重新送簽後又撤回', (w) => { newRound(w, T0 + 9000); withdrawQuote(w, T0 + 12000); }, false],
    ['重新送簽後第一關核准', (w) => { newRound(w, T0 + 9000); approveAt(w, 0, 'm1'); }, false],
    ['重新送簽後全部核准', (w) => { newRound(w, T0 + 9000); approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); }, false],
    ['重新送簽後又被駁回（新的一輪）', (w) => { newRound(w, T0 + 9000); rejectAt(w, 0, 'm1'); }, false],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);

  // ── E5 顧問完成成本 ──
  await matrix('E5 成本完成', async () => {
    const w = mkWorld();
    w.quote.costFlow = { state: 'filled', requestedAt: iso(T0 - 5000), filledAt: iso(T0 + 2000) };
    await fire(w, { type: 'E5_COST_DONE' }, 'cons1');
    return { w, job: await lastRow(w, 'E5_COST_DONE', 'sales1') };
  }, [
    ['沒有變化', () => {}, true],
    ['顧問改回「未完成」（state=requested、filledAt=null）', (w) => { w.quote.costFlow = { state: 'requested', requestedAt: iso(T0 - 5000), filledAt: null }; }, false],
    ['業務改品項使成本退回 requested（X4）', (w) => { w.quote.costFlow = Object.assign({}, w.quote.costFlow, { state: 'requested', filledAt: null }); }, false],
    ['顧問又完成一次（新的 filledAt）→ 舊的這封過期（新的另有自己的去重鍵）', (w) => { w.quote.costFlow.filledAt = iso(T0 + 9000); }, false],
    ['業務換人', (w) => { w.quote.owner = 'cons2'; }, false],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);

  // ── E6 撤回 ──
  await matrix('E6 撤回', async () => {
    const w = mkWorld();
    approveAt(w, 0, 'm1');
    const snap = w.qm.snapApproval(w.quote, w.ctx());
    withdrawQuote(w, T0 + 3000);
    await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'withdrawn', snap }, 'sales1');
    return { w, job: await lastRow(w, 'E6_WITHDRAWN', 'm1') };
  }, [
    ['沒有變化（仍是撤回狀態）', () => {}, true],
    ['業務修改內容但沒有重新送簽', (w) => { w.quote.projectName = 'edited'; }, true],
    ['業務重新送簽（X3：「請勿簽核」已過期）', (w) => newRound(w, T0 + 9000), false],
    ['重新送簽且一級主管再次核准（目前輪到總經理）', (w) => { newRound(w, T0 + 9000); approveAt(w, 0, 'm1'); }, false],
    ['重新送簽後又撤回一次（新的 submittedAt）', (w) => { newRound(w, T0 + 9000); withdrawQuote(w, T0 + 12000); }, false],
    ['重新送簽後全部核准', (w) => { newRound(w, T0 + 9000); approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); }, false],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);

  // ── E6 作廢 ──
  await matrix('E6 作廢', async () => {
    const w = mkWorld();
    approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1');
    const snap = w.qm.snapApproval(w.quote, w.ctx('adm1'));
    voidQuote(w, T0 + 3000);
    await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'voided', snap }, 'adm1');
    return { w, job: await lastRow(w, 'E6_WITHDRAWN', 'm1') };
  }, [
    ['沒有變化（仍是作廢狀態）', () => {}, true],
    ['業務修改內容但沒有重新送簽', (w) => { w.quote.projectName = 'edited'; }, true],
    ['業務重新送簽（作廢的通知已過期）', (w) => newRound(w, T0 + 9000), false],
    ['重新送簽後撤回（state=none 但 submittedAt 有值）', (w) => { newRound(w, T0 + 9000); withdrawQuote(w, T0 + 12000); }, false],
    ['重新送簽後全部核准', (w) => { newRound(w, T0 + 9000); approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); }, false],
    ['已知限制：重新送簽、核准、又作廢第二次 → 較早那次作廢通知仍通過（單據目前確實處於「核准已作廢」，內容為真）', (w) => { newRound(w, T0 + 9000); approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); voidQuote(w, T0 + 15000); }, true],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);
  // 作廢通知的業務本人（別人操作時才會收到）也走同一個判斷
  {
    const w = mkWorld();
    approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1');
    const snap = w.qm.snapApproval(w.quote, w.ctx('adm1'));
    voidQuote(w, T0 + 3000);
    await fire(w, { type: 'E6_WITHDRAWN', resultKind: 'voided', snap }, 'adm1');
    const job = await lastRow(w, 'E6_WITHDRAWN', 'sales1');
    t('E6 作廢｜業務本人收到的那封：沒有變化 → 有效', !!job && w.qm.isStillValid(job) === true);
    newRound(w, T0 + 9000);
    t('E6 作廢｜業務本人收到的那封：重新送簽後 → 過期', w.qm.isStillValid(job) === false);
  }

  // ── E1／E3：既有規則仍然成立，加上改派 ──
  await matrix('E1 送簽', async () => {
    const w = mkWorld();
    await fire(w, { type: 'E1_SUBMIT', idx: 0 });
    return { w, job: await lastRow(w, 'E1_SUBMIT', 'm1') };
  }, [
    ['沒有變化', () => {}, true],
    ['第一關已核准（單據走到下一關）', (w) => approveAt(w, 0, 'gm1'), false],
    ['這一關改派給別人（assignee 變了、歷史有 REASSIGN）', (w) => { w.quote.approval.steps[0].assignee = 'gm1'; histPush(w, 'REASSIGN', T0 + 5000, 'adm1'); }, false],
    ['收件人被改派但歷史沒有 REASSIGN（資料異常）→ 收件人不符', (w) => { w.quote.approval.steps[0].assignee = 'gm1'; }, false],
    ['重新送簽', (w) => newRound(w, T0 + 9000), false],
    ['撤回', (w) => withdrawQuote(w, T0 + 5000), false],
    ['駁回', (w) => rejectAt(w, 0, 'm1'), false],
    ['舊紀錄：REASSIGN 沒有 meta 而 assignee 沒變（無法辨識是同一人）→ 照舊當成真的改派，標記與 stepKey 不符', (w) => { histPush(w, 'REASSIGN', T0 + 5000, 'adm1'); }, false],
    ['管理員改派給同一人（REASSIGN 的 meta.from === meta.to）→ 改派標記不變，原本的 E1 仍有效', (w) => { reassignTo(w, 'm1', T0 + 5000); }, true],
    ['連續兩次改派給同一人 → 仍有效', (w) => { reassignTo(w, 'm1', T0 + 5000); reassignTo(w, 'm1', T0 + 7000); }, true],
    ['A→B 之後 B→B（同一人）→ A 的 E1 因 A→B 過期', (w) => { reassignTo(w, 'gm1', T0 + 5000); reassignTo(w, 'gm1', T0 + 7000); }, false],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);
  await matrix('E3 下一關', async () => {
    const w = mkWorld({ quote: mkQuote('BOARD') });
    approveAt(w, 0, 'm1');
    await fire(w, { type: 'E3_NEXT_STEP', idx: 1 }, 'm1');
    return { w, job: await lastRow(w, 'E3_NEXT_STEP', 'gm1') };
  }, [
    ['沒有變化', () => {}, true],
    ['總經理已核准（再核准後輪到董事會）', (w) => approveAt(w, 1, 'gm2'), false],
    ['業務撤回', (w) => withdrawQuote(w, T0 + 5000), false],
    ['撤回後重新送簽', (w) => { withdrawQuote(w, T0 + 5000); newRound(w, T0 + 9000, 'BOARD'); }, false],
    ['名冊移除了這位總經理', (w) => { w.roster.gm = ['gm2']; }, false],
    ['單據不存在', (w) => { w.data.quotations = []; }, false],
  ]);
  await matrix('E2 請填成本', async () => {
    const w = mkWorld();
    await fire(w, { type: 'E2_COST_REQUEST' });
    return { w, job: await lastRow(w, 'E2_COST_REQUEST', 'cons1') };
  }, [
    ['沒有變化', () => {}, true],
    ['顧問已完成', (w) => { w.quote.costFlow.state = 'filled'; }, false],
    ['換了顧問', (w) => { w.quote.costBy = 'cons2'; }, false],
    ['重新請求（requestedAt 不同）', (w) => { w.quote.costFlow.requestedAt = iso(T0 + 7777); }, false],
    ['requestedAt 被清空', (w) => { w.quote.costFlow.requestedAt = null; }, false],
  ]);
  t('5b 共檢查 ' + n + ' 個情境', n >= 60, n);

  // ── 純函式：code 與壞輸入 ──
  const base = () => mkQuote();
  const helpers = { stepRecipients: (step) => (step.tier === 'mgr1' ? [step.assignee] : ['gm1']) };
  const jobOf = (type, user, sk) => ({ type, toUser: user, dedupeKey: type + ':' + QID + ':' + user + ':' + sk });
  const S = iso(T0);
  eq('checkValidity：E1 有效', VC.checkValidity(jobOf('E1_SUBMIT', 'm1', S + '#0'), base(), helpers), { valid: true, code: 'OK' });
  eq('checkValidity：q 是 null → NO_QUOTE', VC.checkValidity(jobOf('E1_SUBMIT', 'm1', S + '#0'), null, helpers), { valid: false, code: 'NO_QUOTE' });
  eq('checkValidity：job 不是物件 → BAD_KEY', VC.checkValidity(null, base(), helpers).code, 'BAD_KEY');
  eq('checkValidity：未知事件類型 → UNKNOWN_TYPE（不寄）', VC.checkValidity(jobOf('E9_X', 'm1', S + '#0'), base(), helpers), { valid: false, code: 'UNKNOWN_TYPE' });
  eq('checkValidity：沒有 helpers 的 E1 → 無法確認收件人 → RECIPIENT_CHANGED', VC.checkValidity(jobOf('E1_SUBMIT', 'm1', S + '#0'), base()).code, 'RECIPIENT_CHANGED');
  for (const [type, sk] of [['E1_SUBMIT', ''], ['E1_SUBMIT', 'x'], ['E1_SUBMIT', S + '#'], ['E1_SUBMIT', S + '#a'], ['E3_NEXT_STEP', S + '#0@2026-01-01T00:00:00.000Z'], ['E4_RESULT', S + '#r:bogus:0'], ['E4_RESULT', S + '#approved:0'], ['E5_COST_DONE', S], ['E6_WITHDRAWN', S + '#other']]) {
    eq('checkValidity：壞的 stepKey ' + type + ' ' + short(sk) + ' → BAD_KEY', VC.checkValidity(jobOf(type, type === 'E1_SUBMIT' || type === 'E3_NEXT_STEP' ? 'm1' : 'sales1', sk), base(), helpers).code, 'BAD_KEY');
  }
  t('checkValidity：helpers.stepRecipients 丟例外 → 例外往外傳（呼叫端 dispatcher 視為「無法確認」，不寄也不取消，稍後重試）', (() => { try { VC.checkValidity(jobOf('E1_SUBMIT', 'm1', S + '#0'), base(), { stepRecipients: () => { throw new Error('boom'); } }); return false; } catch (e) { return e.message === 'boom'; } })());
  eq('stepKeyOf：取第三個 : 之後（ISO 時間內的 : 保留）', VC.stepKeyOf({ dedupeKey: 'E1_SUBMIT:q1:m1:2026-10-08T03:00:00.000Z#0@2026-10-08T04:00:00.000Z' }), '2026-10-08T03:00:00.000Z#0@2026-10-08T04:00:00.000Z');
  eq('stepKeyOf：帳號含 : 時已跳脫成 %3A，仍取到 stepKey', VC.stepKeyOf({ dedupeKey: 'E1_SUBMIT:q1:a%3Ab:k#0' }), 'k#0');
  eq('stepKeyOf：沒有 dedupeKey → 空字串', VC.stepKeyOf({}), '');
  eq('stepIdxOf：一般／帶改派標記／壞格式', [VC.stepIdxOf(S + '#2'), VC.stepIdxOf(S + '#0@' + iso(T0 + 1)), VC.stepIdxOf('x'), VC.stepIdxOf(S + '#a')], [2, 0, -1, -1]);
  eq('reassignEpoch：沒有歷史／不是物件 → 空字串', [VC.reassignEpoch({}), VC.reassignEpoch(null), VC.reassignEpoch({ history: 'x' })], ['', '', '']);
  eq('reassignEpoch：只有 SUBMIT → 空字串', VC.reassignEpoch({ history: [{ action: 'SUBMIT', at: iso(T0) }] }), '');
  eq('reassignEpoch：SUBMIT 之後的最近一次 REASSIGN', VC.reassignEpoch({ history: [{ action: 'SUBMIT', at: iso(T0) }, { action: 'REASSIGN', at: iso(T0 + 1) }, { action: 'REASSIGN', at: iso(T0 + 2) }] }), iso(T0 + 2));
  eq('reassignEpoch：上一輪的改派不算（更新的 SUBMIT 在後）', VC.reassignEpoch({ history: [{ action: 'SUBMIT', at: iso(T0) }, { action: 'REASSIGN', at: iso(T0 + 1) }, { action: 'WITHDRAW', at: iso(T0 + 2) }, { action: 'SUBMIT', at: iso(T0 + 3) }] }), '');
  eq('reassignEpoch：REASSIGN 缺 at／at 含 #@ 這類會破壞 stepKey 的字元 → 忽略（空字串）', [VC.reassignEpoch({ history: [{ action: 'REASSIGN' }] }), VC.reassignEpoch({ history: [{ action: 'REASSIGN', at: 'a#b' }] })], ['', '']);
  eq('e1StepKey：一般送簽＝<submittedAt>#0（與改派功能加入前完全相同）', VC.e1StepKey({ submittedAt: S, history: [{ action: 'SUBMIT', at: S }] }, 0), S + '#0');
  eq('e1StepKey：改派過＝<submittedAt>#0@<改派時間>', VC.e1StepKey({ submittedAt: S, history: [{ action: 'SUBMIT', at: S }, { action: 'REASSIGN', at: iso(T0 + 5) }] }, 0), S + '#0@' + iso(T0 + 5));

  // ── 同一人再改派（REASSIGN 的 meta.from === meta.to）不是新的改派 ──
  const SUB = { action: 'SUBMIT', at: S };
  const RE = (atMs, meta) => (meta === undefined ? { action: 'REASSIGN', at: iso(atMs) } : { action: 'REASSIGN', at: iso(atMs), meta });
  const epochOf = (hist) => VC.reassignEpoch({ history: hist });
  eq('reassignEpoch：同一人再改派（meta.from===meta.to）→ 略過，沒有更早的改派就是空字串', epochOf([SUB, RE(T0 + 1, { from: 'a', to: 'a' })]), '');
  eq('reassignEpoch：A→B 之後又 B→B → 取 A→B 的時間（B→B 不算）', epochOf([SUB, RE(T0 + 1, { from: 'a', to: 'b' }), RE(T0 + 2, { from: 'b', to: 'b' })]), iso(T0 + 1));
  eq('reassignEpoch：A→B、B→B、B→A → 取最後的 B→A', epochOf([SUB, RE(T0 + 1, { from: 'a', to: 'b' }), RE(T0 + 2, { from: 'b', to: 'b' }), RE(T0 + 3, { from: 'b', to: 'a' })]), iso(T0 + 3));
  eq('reassignEpoch：A→A 後 A→B → 取 A→B', epochOf([SUB, RE(T0 + 1, { from: 'a', to: 'a' }), RE(T0 + 2, { from: 'a', to: 'b' })]), iso(T0 + 2));
  eq('reassignEpoch：meta 不完整／from 為空或不是字串／from≠to／meta 不是物件／沒有 meta → 一律當成真的改派（舊紀錄行為不變）',
    [{ to: 'a' }, { from: '', to: '' }, { from: 1, to: 1 }, { from: 'a', to: 'b' }, 'a→a', null, undefined].map((m) => epochOf([SUB, RE(T0 + 1, m)])),
    [iso(T0 + 1), iso(T0 + 1), iso(T0 + 1), iso(T0 + 1), iso(T0 + 1), iso(T0 + 1), iso(T0 + 1)]);
  eq('reassignEpoch：上一輪的同人改派與新一輪無關（更新的 SUBMIT 在後）', epochOf([SUB, RE(T0 + 1, { from: 'a', to: 'a' }), { action: 'SUBMIT', at: iso(T0 + 9) }]), '');
  eq('e1StepKey：同一人再改派 → stepKey 與沒有改派時完全相同', VC.e1StepKey({ submittedAt: S, history: [SUB, RE(T0 + 1, { from: 'a', to: 'a' })] }, 0), S + '#0');
  eq('e1StepKey：A→B→B → 帶 A→B 的改派時間', VC.e1StepKey({ submittedAt: S, history: [SUB, RE(T0 + 1, { from: 'a', to: 'b' }), RE(T0 + 2, { from: 'b', to: 'b' })] }, 0), S + '#0@' + iso(T0 + 1));
});

// ═════════════════════════════════════════════════════════════════════════
section('5c 重試路徑不寄過期信（FIX-2）：第一次寄送失敗 → 單據狀態改變 → drainDue 取消（STALE），沒改變就照寄', async () => {
  const live = () => mkWorld({ env: { MAIL_MODE: 'live' } });
  const failing = () => ({ ok: false, code: 'SERVER', permanent: false, message: '5xx' });
  /** 先讓第一次寄送失敗（工作留在 pending），再做 change，過了退避時間後 drainDue；回傳 {w, sentBefore, r, rows} */
  async function scenario(setup, spec, me, change) {
    const w = live();
    setup(w);
    w.behavior = failing;
    const snapshot = spec.type === 'E6_WITHDRAWN' ? spec.makeSnap(w) : null;
    if (spec.beforeFire) spec.beforeFire(w);
    await fire(w, Object.assign({}, spec.spec, snapshot ? { snap: snapshot } : {}), me);
    const sentBefore = w.sent.length;
    w.behavior = null;
    change(w);
    w.clock.t += 61000;
    const r = await w.qm.drainDue({ limit: 10 });
    const rows = (await w.outbox.list({})).rows;
    return { w, sentBefore, r, rows };
  }
  const X = [
    // [名稱, setup, spec, me, change, 預期：照寄(true)/取消(false)]
    ['X1 駁回信失敗 → 業務重新送簽', (w) => rejectAt(w, 0, 'm1'), { spec: { type: 'E4_RESULT', resultKind: 'rejected', idx: 0, reason: 'no good' } }, 'm1', (w) => newRound(w, T0 + 9000), false],
    ['X1 對照：駁回信失敗，單據沒變', (w) => rejectAt(w, 0, 'm1'), { spec: { type: 'E4_RESULT', resultKind: 'rejected', idx: 0, reason: 'no good' } }, 'm1', () => {}, true],
    ['X2 最終核准信失敗 → 核准後修改使核准作廢', (w) => { approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); }, { spec: { type: 'E4_RESULT', resultKind: 'final_approved', idx: 1 } }, 'gm1', (w) => voidQuote(w, T0 + 5000), false],
    ['X2 對照：最終核准信失敗，單據沒變', (w) => { approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); }, { spec: { type: 'E4_RESULT', resultKind: 'final_approved', idx: 1 } }, 'gm1', () => {}, true],
    ['X3 撤回信失敗 → 重新送簽且一級主管再次核准', (w) => approveAt(w, 0, 'm1'), { type: 'E6_WITHDRAWN', spec: { type: 'E6_WITHDRAWN', resultKind: 'withdrawn' }, makeSnap: (w) => { const sn = w.qm.snapApproval(w.quote, w.ctx()); withdrawQuote(w, T0 + 3000); return sn; } }, 'sales1', (w) => { newRound(w, T0 + 9000); approveAt(w, 0, 'm1'); }, false],
    ['X3 對照：撤回信失敗，單據沒變', (w) => approveAt(w, 0, 'm1'), { type: 'E6_WITHDRAWN', spec: { type: 'E6_WITHDRAWN', resultKind: 'withdrawn' }, makeSnap: (w) => { const sn = w.qm.snapApproval(w.quote, w.ctx()); withdrawQuote(w, T0 + 3000); return sn; } }, 'sales1', () => {}, true],
    ['X4 成本完成信失敗 → 業務改品項使成本退回', (w) => { w.quote.costFlow = { state: 'filled', requestedAt: iso(T0 - 5000), filledAt: iso(T0 + 2000) }; }, { spec: { type: 'E5_COST_DONE' } }, 'cons1', (w) => { w.quote.costFlow = { state: 'requested', requestedAt: iso(T0 + 9000), filledAt: null }; }, false],
    ['X4 對照：成本完成信失敗，單據沒變', (w) => { w.quote.costFlow = { state: 'filled', requestedAt: iso(T0 - 5000), filledAt: iso(T0 + 2000) }; }, { spec: { type: 'E5_COST_DONE' } }, 'cons1', () => {}, true],
    ['X5 作廢信失敗 → 業務重新送簽', (w) => { approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); }, { type: 'E6_WITHDRAWN', spec: { type: 'E6_WITHDRAWN', resultKind: 'voided' }, makeSnap: (w) => { const sn = w.qm.snapApproval(w.quote, w.ctx('adm1')); voidQuote(w, T0 + 3000); return sn; } }, 'adm1', (w) => newRound(w, T0 + 9000), false],
    ['X5 對照：作廢信失敗，單據沒變', (w) => { approveAt(w, 0, 'm1'); approveAt(w, 1, 'gm1'); }, { type: 'E6_WITHDRAWN', spec: { type: 'E6_WITHDRAWN', resultKind: 'voided' }, makeSnap: (w) => { const sn = w.qm.snapApproval(w.quote, w.ctx('adm1')); voidQuote(w, T0 + 3000); return sn; } }, 'adm1', () => {}, true],
    ['X6 一級主管送簽信失敗 → 一級主管已核准、輪到總經理（再核准後）', () => {}, { spec: { type: 'E1_SUBMIT', idx: 0 } }, 'sales1', (w) => approveAt(w, 0, 'm1'), false],
    ['X6 對照：送簽信失敗，單據沒變', () => {}, { spec: { type: 'E1_SUBMIT', idx: 0 } }, 'sales1', () => {}, true],
  ];
  for (const [name, setup, spec, me, change, wantSent] of X) {
    const { w, sentBefore, r, rows } = await scenario(setup, spec, me, change);
    const statuses = rows.map((x) => x.status + (x.skipReason ? '/' + x.skipReason : ''));
    if (wantSent) {
      t(name + '：照常補寄（sent ≥1、沒有取消）', r.sent >= 1 && r.cancelled === 0 && w.sent.length > sentBefore && rows.every((x) => x.status === 'sent'), JSON.stringify({ sent: r.sent, cancelled: r.cancelled, statuses }));
    } else {
      t(name + '：過期信不寄（沒有新的寄送、全部 cancelled/STALE）', r.sent === 0 && w.sent.length === sentBefore && r.cancelled >= 1 && rows.every((x) => x.status === 'cancelled' && x.skipReason === 'STALE'), JSON.stringify({ sent: r.sent, cancelled: r.cancelled, statuses, newMails: w.sent.slice(sentBefore).map((m) => m.subject) }));
    }
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('5d 改派去重（FIX-6）：A→B→A 時 A 會再收到一封 E1；同一次改派重複觸發仍去重', async () => {
  const w = mkWorld();
  w.users.push({ username: 'm1b', role: 'manager1', displayName: 'M1B Name', email: 'm1b@itts.test' });
  const reassign = (to, atMs) => reassignTo(w, to, atMs);       // 與 POST /reassign 相同：換 assignee、歷史多一筆帶 meta {from,to} 的 REASSIGN
  histPush(w, 'SUBMIT', T0);                                      // 送簽當下的歷史（POST /submit 會寫）
  await fire(w, { type: 'E1_SUBMIT', idx: 0 });                  // ① 送簽 → m1
  t('① 送簽 → 一級主管 m1 收到 1 封', w.sent.map(toOf).join() === 'm1@itts.test', JSON.stringify(w.sent.map(toOf)));
  reassign('m1b', T0 + 5000);
  await fire(w, { type: 'E1_SUBMIT', idx: 0 }, 'adm1');          // ② 改派 A→B
  t('② 改派給 m1b → m1b 收到 1 封', w.sent.map(toOf).join() === 'm1@itts.test,m1b@itts.test', JSON.stringify(w.sent.map(toOf)));
  await fire(w, { type: 'E1_SUBMIT', idx: 0 }, 'adm1');          // 同一次改派被重複觸發
  t('同一次改派重複觸發 → 去重（沒有新信）', w.sent.length === 2, w.sent.length);
  reassign('m1', T0 + 9000);
  await fire(w, { type: 'E1_SUBMIT', idx: 0 }, 'adm1');          // ③ 改派 B→A
  t('③ 改回 m1（A→B→A）→ m1 再收到第二封 E1', w.sent.map(toOf).join() === 'm1@itts.test,m1b@itts.test,m1@itts.test', JSON.stringify(w.sent.map(toOf)));
  await fire(w, { type: 'E1_SUBMIT', idx: 0 }, 'adm1');
  t('③ 同一次改派（B→A）重複觸發 → 去重', w.sent.length === 3, w.sent.length);
  const rows = (await w.outbox.list({ type: 'E1_SUBMIT' })).rows;
  const keys = rows.map((x) => x.dedupeKey.split(':').slice(3).join(':')).sort();
  eq('三筆 E1 的 stepKey：一般送簽沒有標記；兩次改派各帶自己的改派時間', keys, [iso(T0) + '#0', iso(T0) + '#0@' + iso(T0 + 5000), iso(T0) + '#0@' + iso(T0 + 9000)].sort());
  t('三筆 dedupeKey 彼此不同', new Set(rows.map((x) => x.dedupeKey)).size === 3);
  const jobA1 = rows.find((x) => x.toUser === 'm1' && !/@/.test(x.dedupeKey.split(':').slice(3).join(':')));
  const jobB = rows.find((x) => x.toUser === 'm1b');
  const jobA2 = rows.find((x) => x.toUser === 'm1' && /@/.test(x.dedupeKey.split(':').slice(3).join(':')));
  t('isStillValid：A 第一次送簽的 E1 已過期（之後又改派過）、B 的過期（已改回 A）、A 第二封有效', w.qm.isStillValid(jobA1) === false && w.qm.isStillValid(jobB) === false && w.qm.isStillValid(jobA2) === true, JSON.stringify([w.qm.checkJob(jobA1).code, w.qm.checkJob(jobB).code, w.qm.checkJob(jobA2).code]));
  const rb = w.qm.rebuild(jobA2);
  t('rebuild 認得帶改派標記的 stepKey，事件 stepKey 與工作相同、信件內容與首次相同', !!rb && rb.ev.stepKey === jobA2.dedupeKey.split(':').slice(3).join(':') && rb.ev.at === jobA2.createdAt && !!rb.ev.step, rb && rb.ev && rb.ev.stepKey);
  const RND = load('lib/mail/render.js');
  const re = RND.renderMail(rb.ev, { username: 'm1', label: 'M1 Name', kind: rb.kind || jobA2.meta.kind }, { config: w.config, now: w.clock.t });
  const first = w.sent[2];
  t('重建的信與第一次寄出的逐字相同（主旨、純文字）', re.subject === first.subject && re.text === first.text);

  // 重新送簽後（新的一輪）改派標記歸零：新一輪的第一封 E1 是 <新 submittedAt>#0（沒有 @）
  newRound(w, T0 + 20000);
  await fire(w, { type: 'E1_SUBMIT', idx: 0 });
  const rows2 = (await w.outbox.list({ type: 'E1_SUBMIT' })).rows;
  const last = rows2.find((x) => x.dedupeKey.indexOf(iso(T0 + 20000)) >= 0);
  t('新的一輪送簽：stepKey 回到 <submittedAt>#0（上一輪的改派不算）', !!last && last.dedupeKey.split(':').slice(3).join(':') === iso(T0 + 20000) + '#0', last && last.dedupeKey);

  // 失敗後重試：帶改派標記的 E1 第一次寄送失敗，之後由 drainDue 補寄（rebuild 認得、isStillValid 通過）
  const wr = mkWorld({ env: { MAIL_MODE: 'live' } });
  wr.users.push({ username: 'm1b', role: 'manager1', displayName: 'M1B Name', email: 'm1b@itts.test' });
  histPush(wr, 'SUBMIT', T0);
  wr.quote.approval.steps[0].assignee = 'm1b'; histPush(wr, 'REASSIGN', T0 + 5000, 'adm1');
  wr.behavior = () => ({ ok: false, code: 'SERVER', permanent: false, message: '5xx' });
  await fire(wr, { type: 'E1_SUBMIT', idx: 0 }, 'adm1');
  wr.behavior = null; wr.clock.t += 61000;
  const dr = await wr.qm.drainDue({ limit: 5 });
  t('改派後的 E1 第一次失敗 → drainDue 補寄成功（sent 1、沒有被當成過期）', dr.sent === 1 && dr.cancelled === 0 && wr.sent.length === 2 && wr.sent[1].to[0] === 'm1b@itts.test', JSON.stringify(dr));
});

// ═════════════════════════════════════════════════════════════════════════
section('5f 同人改派不重複寄（T1）：只有承辦人真的改變才發新的 E1；A→B→A 仍再寄；同一次改派重複觸發去重', async () => {
  const E1 = { type: 'E1_SUBMIT', idx: 0 };
  const mk = (env) => {
    const w = mkWorld(env ? { env } : undefined);
    w.users.push({ username: 'm1b', role: 'manager1', displayName: 'M1B Name', email: 'm1b@itts.test' });
    histPush(w, 'SUBMIT', T0);
    return w;
  };
  const skOf = (r) => r.dedupeKey.split(':').slice(3).join(':');
  const e1Rows = async (w) => (await w.outbox.list({ type: 'E1_SUBMIT' })).rows;
  /** 管理員改派（換 assignee＋歷史 REASSIGN）後觸發 E1，與 POST /reassign 的順序相同 */
  const adminReassign = async (w, to, atMs) => { reassignTo(w, to, atMs); await fire(w, E1, 'adm1'); };

  // ① A→A：管理員把一級主管「改派」給目前的承辦人本人（含連續按）→ 沒有新信
  {
    const w = mk();
    await fire(w, E1);
    eq('① 送簽 → 一級主管 m1 收到 1 封', w.sent.map(toOf), ['m1@itts.test']);
    await adminReassign(w, 'm1', T0 + 1000);
    t('① 改派給目前的承辦人本人（A→A）→ 沒有新信（仍是 1 封）', w.sent.length === 1, JSON.stringify(w.sent.map(toOf)));
    await adminReassign(w, 'm1', T0 + 2000);
    await adminReassign(w, 'm1', T0 + 3000);
    t('① 連續再按兩次 → 仍是 1 封', w.sent.length === 1, JSON.stringify(w.sent.map(toOf)));
    const rows = await e1Rows(w);
    t('① 寄件匣只有 1 筆 E1，stepKey 沒有改派標記（同人改派不產生新的去重鍵）', rows.length === 1 && skOf(rows[0]) === iso(T0) + '#0', rows.map(skOf).join(' | '));
    t('① 原本那封 E1 仍然有效（沒有被同人改派當成過期）', rows.length === 1 && w.qm.isStillValid(rows[0]) === true, rows[0] && w.qm.checkJob(rows[0]).code);
  }

  // ② A→A 時，上一封 E1 還在重試中（暫時寄不出去）：不能被取消，也不能因為「不寄新信」而漏信
  {
    const w = mk({ MAIL_MODE: 'live' });
    w.behavior = () => ({ ok: false, code: 'SERVER', permanent: false, message: '5xx' });
    await fire(w, E1);
    const pend = await e1Rows(w);
    t('② （準備）第一封 E1 暫時寄不出去 → 留在寄件匣 pending', pend.length === 1 && pend[0].status === 'pending', pend.map((x) => x.status).join());
    await adminReassign(w, 'm1', T0 + 1000);
    w.behavior = null;
    w.clock.t += 61000;
    const r = await w.qm.drainDue({ limit: 5 });
    const rows = await e1Rows(w);
    t('② 同人改派後重試：補寄成功（sent 1、沒有取消），m1 收到那封 E1', r.sent === 1 && r.cancelled === 0 && rows.length === 1 && rows[0].status === 'sent' && w.sent.length === 2 && w.sent[1].to[0] === 'm1@itts.test', JSON.stringify({ r, st: rows.map((x) => x.status + (x.skipReason ? '/' + x.skipReason : '')), n: w.sent.length }));
  }

  // ③ A→B→B→B→A→A：每次「真的換人」才寄；同人再改派都不寄
  {
    const w = mk();
    await fire(w, E1);
    await adminReassign(w, 'm1b', T0 + 1000);          // A→B：B 收到
    await adminReassign(w, 'm1b', T0 + 2000);          // B→B
    await adminReassign(w, 'm1b', T0 + 2500);          // B→B
    await adminReassign(w, 'm1', T0 + 3000);           // B→A：A 再收到第二封（FIX-6 的行為保留）
    await adminReassign(w, 'm1', T0 + 4000);           // A→A
    eq('③ 共 3 封：m1（送簽）、m1b（A→B）、m1（B→A）；兩次 B→B、一次 A→A 都沒有新信', w.sent.map(toOf), ['m1@itts.test', 'm1b@itts.test', 'm1@itts.test']);
    await fire(w, E1, 'adm1');                          // 同一次改派（A→A 之後）重複觸發
    await fire(w, E1, 'adm1');
    t('③ 同一次改派被重複觸發 → 去重（仍是 3 封）', w.sent.length === 3, w.sent.length);
    const rows = await e1Rows(w);
    eq('③ 三筆 E1 的 stepKey：送簽沒有標記；A→B 帶 T0+1000；B→A 帶 T0+3000（同人改派的時間 T0+2000／2500／4000 都不出現）',
      rows.map(skOf).sort(), [iso(T0) + '#0', iso(T0) + '#0@' + iso(T0 + 1000), iso(T0) + '#0@' + iso(T0 + 3000)].sort());
    const aFirst = rows.find((x) => x.toUser === 'm1' && !/@/.test(skOf(x)));
    const bRow = rows.find((x) => x.toUser === 'm1b');
    const aSecond = rows.find((x) => x.toUser === 'm1' && /@/.test(skOf(x)));
    t('③ isStillValid：A 第一封（之後被改派過）過期、B 的過期（已改回 A）、A 第二封有效（最後的 A→A 不影響）',
      !!aFirst && !!bRow && !!aSecond && w.qm.isStillValid(aFirst) === false && w.qm.isStillValid(bRow) === false && w.qm.isStillValid(aSecond) === true,
      [aFirst, bRow, aSecond].map((x) => (x ? w.qm.checkJob(x).code : '?')).join());
  }

  // ④ 先 A→A 再 A→B：B 照常收到（同人改派不會吃掉後面真的改派）
  {
    const w = mk();
    await fire(w, E1);
    await adminReassign(w, 'm1', T0 + 1000);
    await adminReassign(w, 'm1b', T0 + 2000);
    eq('④ A→A 不寄；之後 A→B → m1b 收到', w.sent.map(toOf), ['m1@itts.test', 'm1b@itts.test']);
    const rows = await e1Rows(w);
    t('④ B 那封的 stepKey 帶 A→B 的時間（T0+2000）', rows.some((x) => x.toUser === 'm1b' && skOf(x) === iso(T0) + '#0@' + iso(T0 + 2000)), rows.map(skOf).join(' | '));
  }

  // ⑤ 舊紀錄（REASSIGN 沒有 meta，例如上線前就存在的改派紀錄）：無法辨識是不是同一人，維持原本的行為（每次改派都是新的 E1）
  {
    const w = mk();
    await fire(w, E1);
    histPush(w, 'REASSIGN', T0 + 1000, 'adm1');          // 沒有 meta
    await fire(w, E1, 'adm1');
    t('⑤ 沒有 meta 的 REASSIGN → 視為真的改派，新的 E1 照寄（維持改版前的行為）', w.sent.length === 2 && w.sent.every((m) => toOf(m) === 'm1@itts.test'), JSON.stringify(w.sent.map(toOf)));
  }
});

// ═════════════════════════════════════════════════════════════════════════
// FIX-3：時限。注入「永不 resolve」「慢」「之後才 reject」的儲存體，驗證回應在時限內返回、業務動作（呼叫端）不受影響、沒有 unhandled rejection
// ═════════════════════════════════════════════════════════════════════════
const hangMethods = (inner, names, mode) => new Proxy(inner, {
  get(target, prop) {
    const v = target[prop];
    if (typeof v !== 'function') return v;
    if (names.indexOf(String(prop)) < 0) return v.bind(target);
    if (mode === 'hang') return () => new Promise(() => {});
    if (mode === 'slow') return (...a) => sleep(150).then(() => v.apply(target, a));
    if (mode === 'late-reject') return () => sleep(900).then(() => { throw new Error('late boom'); });
    return v.bind(target);
  },
});
/** 等寄信 promise 全部 settle，但最多 capMs。優先用 quoteMail.waitPending（server.js 的 res.end 包裝用的那個）；舊版沒有就退回 allSettled（舊行為）。 */
async function settleOrTimeout(w, req, capMs) {
  const t0 = Date.now();
  const p = typeof w.qm.waitPending === 'function' ? w.qm.waitPending(req) : Promise.allSettled(req._mailPending || []);
  const r = await Promise.race([p.then(() => 'settled'), sleep(capMs).then(() => 'STUCK')]);
  return { r, ms: Date.now() - t0 };
}

section('5e 時限（FIX-3）：outbox／資料庫卡住時，寄信流程在時限內放手', async () => {
  const unhandled = [];
  const onUnhandled = (e) => { unhandled.push(e && e.message ? e.message : String(e)); };
  process.on('unhandledRejection', onUnhandled);
  try {
    const LIM = { notifyMs: 300, waitMs: 500, drainCapMs: 600, flushMs: 200 };
    // 各個 adapter 方法卡住（insert＝入列；breakerLoad＝熔斷查詢；claimById＝當次領取；claimDue／expireExhausted＝清理）
    // 送出之前的操作（insert 入列、breakerLoad 熔斷查詢、claimById 當次領取）卡住 → 沒寄；送出之後的操作（markSent、claimDue／expireExhausted 路由尾端清理）卡住 → 信已寄出，回應照樣不被拖住
    const PRE_SEND = ['insert', 'breakerLoad', 'claimById'];
    for (const method of PRE_SEND.concat(['markSent', 'claimDue', 'expireExhausted'])) {
      const w = mkWorld({ adapter: hangMethods(memoryAdapter(), [method], 'hang'), limits: LIM });
      const req = {};
      w.qm.notifyMail(req, w.ctx(), w.quote, { type: 'E1_SUBMIT', idx: 0 });
      const { r, ms } = await settleOrTimeout(w, req, 2500);
      t('adapter.' + method + ' 永不 resolve：寄信 promise 在時限內 settle（≤1.5 秒；notifyMs=300）', r === 'settled' && ms < 1500, r + ' ' + ms + 'ms');
      const st = await Promise.race([Promise.allSettled(req._mailPending), sleep(1500).then(() => 'STUCK')]);
      t('adapter.' + method + ' 卡住：_mailPending 裡的 promise 本身已 fulfilled（永不 reject）', st !== 'STUCK' && st.every((x) => x.status === 'fulfilled'), JSON.stringify(st));
      t('adapter.' + method + ' 卡住：' + (PRE_SEND.indexOf(method) >= 0 ? '沒有寄出任何信' : '信已經寄出（1 封）') + '、呼叫端沒有例外', w.sent.length === (PRE_SEND.indexOf(method) >= 0 ? 0 : 1), String(w.sent.length));
      if (typeof w.qm.waitPending === 'function') {
        const dg = w.qm.diagnostics().deadlines;
        t('adapter.' + method + ' 卡住：diagnostics 記到一次 notify 時限（只有計數）', dg && dg.notify === 1, JSON.stringify(dg));
        t('adapter.' + method + ' 卡住：稽核 QUOTE_MAIL_FAILED code=DEADLINE，無位址／金額／專案名', w.logs.some((l) => l.action === 'QUOTE_MAIL_FAILED' && /code=DEADLINE/.test(l.detail)) && w.logs.every((l) => !FULL_EMAIL_RE.test(l.detail) && !l.detail.includes(PROJECT)), JSON.stringify(w.logs));
      }
    }
    // 同一請求多次 notifyMail（核准同時寄 E4 與 E3）：全部卡住也只等一個時限
    {
      const w = mkWorld({ adapter: hangMethods(memoryAdapter(), ['insert'], 'hang'), limits: LIM });
      w.quote.approval.cur = 1; w.quote.approval.steps[0].status = 'approved'; w.quote.approval.steps[0].by = 'm1'; w.quote.approval.steps[1].status = 'pending';
      const req = {};
      w.qm.notifyMail(req, w.ctx('m1'), w.quote, { type: 'E4_RESULT', resultKind: 'approved', idx: 0 });
      w.qm.notifyMail(req, w.ctx('m1'), w.quote, { type: 'E3_NEXT_STEP', idx: 1 });
      const { r, ms } = await settleOrTimeout(w, req, 2500);
      t('一個請求兩次 notifyMail 都卡住：合計仍在一個時限內返回（≤1.5 秒）', r === 'settled' && ms < 1500 && req._mailPending.length === 2, r + ' ' + ms + 'ms');
    }
    // 慢但沒卡死：每個儲存體操作多 150ms，仍在時限內 → 信寄出
    {
      const w = mkWorld({ adapter: hangMethods(memoryAdapter(), ['insert', 'claimById', 'breakerLoad'], 'slow'), limits: { notifyMs: 3000, waitMs: 4000, drainCapMs: 6000, flushMs: 1000 } });
      const req = {};
      w.qm.notifyMail(req, w.ctx(), w.quote, { type: 'E1_SUBMIT', idx: 0 });
      const { r, ms } = await settleOrTimeout(w, req, 5000);
      t('慢 adapter（每個操作 +150ms）：時限內完成、信照常寄出', r === 'settled' && w.sent.length === 1 && ms < 3000, r + ' ' + ms + 'ms sent=' + w.sent.length);
      if (typeof w.qm.waitPending === 'function') t('慢 adapter：沒有觸發任何時限', w.qm.diagnostics().deadlines.notify === 0);
    }
    // 被丟下的 promise 之後才 reject：不得有 unhandled rejection
    {
      const w = mkWorld({ adapter: hangMethods(memoryAdapter(), ['insert'], 'late-reject'), limits: LIM });
      const req = {};
      w.qm.notifyMail(req, w.ctx(), w.quote, { type: 'E1_SUBMIT', idx: 0 });
      await settleOrTimeout(w, req, 2500);
      await sleep(1300);                                        // 等被丟下的 adapter promise（900ms 後 reject）真的 reject 了
      t('被丟下的工作之後才 reject：沒有 unhandled rejection', unhandled.length === 0, unhandled.join(' | '));
    }
    // waitPending 的外層保險：就算有人把永不 settle 的 promise 放進 _mailPending 也不會卡住回應
    {
      const w = mkWorld({ limits: LIM });
      const req = { _mailPending: [new Promise(() => {})] };
      if (typeof w.qm.waitPending === 'function') {
        const t0 = Date.now();
        const r = await Promise.race([w.qm.waitPending(req), sleep(2000).then(() => 'STUCK')]);
        t('waitPending：_mailPending 裡有永不 settle 的 promise → 在 waitMs（500ms）後返回 {timedOut:true}', r !== 'STUCK' && r.timedOut === true && Date.now() - t0 < 1500, JSON.stringify(r));
        t('waitPending：沒有 _mailPending → 立刻 {timedOut:false,pending:0}', JSON.stringify(await w.qm.waitPending({})) === JSON.stringify({ timedOut: false, pending: 0 }));
        t('waitPending：參數亂七八糟不 throw', (await Promise.all([undefined, null, 5, 'x', []].map((x) => w.qm.waitPending(x)))).every((x) => x.timedOut === false));
      } else {
        t('waitPending 存在（server.js 的 res.end 包裝用它）', false, '尚未實作');
      }
    }
    // deadline 輔助函式
    {
      const QMmod = load('lib/mail/quoteMail.js');
      if (typeof QMmod.deadline === 'function') {
        const dl = QMmod.deadline;
        eq('deadline：正常完成', await dl(Promise.resolve(7), 100), { timedOut: false, value: 7 });
        const e1 = await dl(Promise.reject(new Error('x')), 100);
        t('deadline：reject → 回 {error}，不 throw', e1.timedOut === false && e1.error && e1.error.message === 'x');
        eq('deadline：超時', await dl(new Promise(() => {}), 30), { timedOut: true });
        eq('deadline：同步值也可以', await dl(5, 30), { timedOut: false, value: 5 });
        const timers = () => process.getActiveResourcesInfo().filter((x) => x === 'Timeout').length;
        const tb = timers();
        await dl(Promise.resolve(1), 5000);
        await dl(Promise.reject(new Error('x')), 5000);
        t('deadline：完成或 reject 之後計時器已清除（不留 pending 的 Timeout）', timers() <= tb, tb + ' → ' + timers());
        const before = unhandled.length;
        await dl(new Promise((resolve, reject) => setTimeout(() => reject(new Error('late')), 120)), 20);
        await sleep(300);
        t('deadline：被丟下的 promise 之後才 reject → 沒有 unhandled rejection', unhandled.length === before, unhandled.join(' | '));
      } else {
        t('deadline 輔助函式存在', false, '尚未實作');
      }
    }
    // drainDue（Cron／後台重送）有界
    {
      const w = mkWorld({ adapter: hangMethods(memoryAdapter(), ['claimDue'], 'hang'), limits: LIM });
      const t0 = Date.now();
      const r = await Promise.race([w.qm.drainDue({ limit: 2, budgetMs: 100 }), sleep(3000).then(() => 'STUCK')]);
      t('drainDue：claimDue 卡住 → 在整體時限內返回（≤1.5 秒；drainCapMs=600）且標記 timedOut', r !== 'STUCK' && Date.now() - t0 < 1500 && r.timedOut === true && r.errors >= 1, JSON.stringify(r));
      if (r !== 'STUCK') t('drainDue 逾時的摘要只有計數（沒有收件人）', Array.isArray(r.skipped) && Array.isArray(r.failed) && r.sent === 0);
    }
    // pollDrain：清理之後的 db.flush 卡住也有界
    {
      const w = mkWorld({ limits: LIM });
      const qm2 = createQuoteMail({
        config: w.config, outbox: w.outbox, transport: w.transport, now: () => w.clock.t, limits: LIM,
        db: { load: () => w.data, flush: () => new Promise(() => {}) }, loadAuth: () => ({ users: w.users }), writeLog: () => {},
      });
      qm2.bind(HELPERS);
      const t0 = Date.now();
      const r = await Promise.race([qm2.pollDrain(), sleep(3000).then(() => 'STUCK')]);
      t('pollDrain：db.flush 永不 resolve → 在 flushMs（200ms）內放手，pollDrain 照常返回 true', r === true && Date.now() - t0 < 1500, r + ' ' + (Date.now() - t0) + 'ms');
    }
  } finally {
    process.off('unhandledRejection', onUnhandled);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('6 drainDue 走膠水（到期工作被領取、重建、重寄、過期取消）', async () => {
  const w = mkWorld({ env: { MAIL_MODE: 'live' } });
  w.behavior = () => ({ ok: false, code: 'TIMEOUT', permanent: false, message: 'slow' });
  await fire(w, { type: 'E1_SUBMIT', idx: 0 });
  let job = (await w.outbox.list({})).rows[0];
  t('傳輸逾時 → 工作留 pending（等重試）', job.status === 'pending' && job.attempts === 1 && w.sent.length === 1, JSON.stringify(job));
  w.behavior = null;
  let r = await w.qm.drainDue({ limit: 5 });
  t('還沒到期的工作不會被領取', r.sent === 0 && w.sent.length === 1);
  w.clock.t += 61000;
  r = await w.qm.drainDue({ limit: 5 });
  job = (await w.outbox.list({})).rows[0];
  t('到期後 drainDue 重建並寄出（sent）', r.sent === 1 && job.status === 'sent' && w.sent.length === 2 && w.sent[1].to[0] === 'm1@itts.test', JSON.stringify(r));
  t('重試的信與第一次逐字相同', w.sent[0].subject === w.sent[1].subject && w.sent[0].html === w.sent[1].html && w.sent[0].text === w.sent[1].text);
  // 到期但單據已過期 → 取消
  const w2 = mkWorld();
  w2.behavior = () => ({ ok: false, code: 'SERVER', permanent: false, message: '5xx' });
  await fire(w2, { type: 'E1_SUBMIT', idx: 0 });
  w2.behavior = null;
  w2.quote.approval.cur = 1; w2.quote.approval.steps[0].status = 'approved'; w2.quote.approval.steps[1].status = 'pending';
  w2.clock.t += 61000;
  const r2 = await w2.qm.drainDue({ limit: 5 });
  const j2 = (await w2.outbox.list({})).rows[0];
  t('過期的舊關卡信 → cancelled／STALE，沒有誤寄', r2.cancelled === 1 && j2.status === 'cancelled' && j2.skipReason === 'STALE' && w2.sent.length === 1, JSON.stringify(j2));
  // 單據被刪 → GONE
  const w3 = mkWorld();
  w3.behavior = () => ({ ok: false, code: 'SERVER', permanent: false, message: '5xx' });
  await fire(w3, { type: 'E1_SUBMIT', idx: 0 });
  w3.behavior = null; w3.data.quotations = []; w3.clock.t += 61000;
  const r3 = await w3.qm.drainDue({ limit: 5 });
  t('單據不存在 → cancelled／GONE', r3.cancelled === 1 && (await w3.outbox.list({})).rows[0].skipReason === 'GONE');
  // 帳號補了 email 之後重送（requeue → drain）
  const w4 = mkWorld();
  w4.users.find((x) => x.username === 'm1').email = '';
  await fire(w4, { type: 'E1_SUBMIT', idx: 0 });
  const j4 = (await w4.outbox.list({})).rows[0];
  t('沒有 email → skipped／NO_EMAIL', j4.status === 'skipped' && j4.skipReason === 'NO_EMAIL' && w4.sent.length === 0);
  w4.users.find((x) => x.username === 'm1').email = 'm1@itts.test';
  await fire(w4, { type: 'E1_SUBMIT', idx: 0 });
  t('補了 email 後同一事件重新觸發仍是 DUPLICATE（不自動補寄）', w4.sent.length === 0);
  const rq = await w4.outbox.requeue(j4.id);
  const r4 = await w4.qm.drainDue({ limit: 5 });
  t('後台 requeue 後 drainDue 補寄', rq.ok && r4.sent === 1 && w4.sent.length === 1 && w4.sent[0].to[0] === 'm1@itts.test');
});

// ═════════════════════════════════════════════════════════════════════════
section('7 pollDrain：節流（每實例 60 秒）、逾時（4 秒）、零成本', async () => {
  const w = mkWorld();
  t('第一次 → 跑（runs=1）', (await w.qm.pollDrain()) === true && w.qm.diagnostics().poll.runs === 1);
  w.clock.t += 30000;
  t('30 秒後 → 節流，不跑', (await w.qm.pollDrain()) === false && w.qm.diagnostics().poll.runs === 1);
  w.clock.t += 29999;
  t('59.999 秒後 → 還是節流', (await w.qm.pollDrain()) === false && w.qm.diagnostics().poll.runs === 1);
  w.clock.t += 2;
  t('超過 60 秒 → 再跑（runs=2）', (await w.qm.pollDrain()) === true && w.qm.diagnostics().poll.runs === 2);
  const d = w.qm.diagnostics().poll;
  t('診斷只有計數（sent／failed／skipped／cancelled／queued／errors），沒有收件人', d.last && typeof d.last.sent === 'number' && !/sales1|m1|gm1/.test(JSON.stringify(d)), JSON.stringify(d));
  t('清理後有 flush（稽核寫入）', w.flushes >= 2, w.flushes);
  // 同時兩個 poll 只有一個跑
  const w2 = mkWorld();
  const [a, b] = await Promise.all([w2.qm.pollDrain(), w2.qm.pollDrain()]);
  t('同時進來的兩個 poll 只有一個會跑', [a, b].filter(Boolean).length === 1, JSON.stringify([a, b]));
  // 領取到期工作
  const w3 = mkWorld();
  w3.behavior = () => ({ ok: false, code: 'SERVER', permanent: false, message: '5xx' });
  await fire(w3, { type: 'E1_SUBMIT', idx: 0 });
  w3.behavior = null; w3.clock.t += 61000;
  t('poll 領走到期工作並寄出', (await w3.qm.pollDrain()) === true && w3.sent.length === 2 && (await w3.outbox.list({})).rows[0].status === 'sent');
  // 逾時：傳輸永不回應 → pollDrain 在 4 秒後放手（工作靠租約到期後由下一個請求接手）
  const w4 = mkWorld({ timeouts: { connectMs: 20, totalMs: 60000 } });
  w4.behavior = () => ({ ok: false, code: 'SERVER', permanent: false, message: '5xx' });
  await fire(w4, { type: 'E1_SUBMIT', idx: 0 });
  w4.behavior = () => new Promise(() => {});                    // 永不 resolve
  w4.clock.t += 61000;
  const t0 = Date.now();
  const ran = await w4.qm.pollDrain();
  const dt = Date.now() - t0;
  const dg = w4.qm.diagnostics().poll;
  t('傳輸永不回應：pollDrain 約 4 秒後放手（3.5～5 秒）', ran === true && dt >= 3500 && dt < 5000, dt + 'ms');
  t('逾時被記錄（last.timedOut）且 runs 增加', dg.last && dg.last.timedOut === true && dg.runs === 1, JSON.stringify(dg));
  const jobs4 = (await w4.outbox.list({})).rows;
  t('被丟下的工作在 sending（租約中），沒有被重複領取', jobs4.length === 1 && jobs4[0].status === 'sending' && w4.sent.length === 2, JSON.stringify(jobs4[0]));
  w4.clock.t += 61000;
  w4.behavior = null;
  t('節流窗口過後可以再跑（不會被前一次卡住的清理擋住）', (await w4.qm.pollDrain()) === true && w4.qm.diagnostics().poll.runs === 2);
});

// ═════════════════════════════════════════════════════════════════════════
section('8 emailStatusMap／missingEmails', () => {
  const w = mkWorld();
  const st = w.qm.emailStatusMap(w.users);
  t('有 email＝OK、沒有＝NO_EMAIL；停用帳號一律 OK', st.m1 === 'OK' && st.noemail === 'NO_EMAIL' && st.off1 === 'OK', JSON.stringify(st));
  w.users.find((x) => x.username === 'm1').email = 'm1@example.com';
  w.users.find((x) => x.username === 'gm1').email = 'not an email';
  const st2 = w.qm.emailStatusMap(w.users);
  eq('網域不在白名單／格式壞掉', [st2.m1, st2.gm1], ['DOMAIN_NOT_ALLOWED', 'BAD_EMAIL']);
  t('結果裡沒有任何位址', !FULL_EMAIL_RE.test(JSON.stringify(st2)));
  eq('輸入不是陣列 → 空物件', w.qm.emailStatusMap(null), {});
  eq('陣列裡有 null／沒有帳號名稱的項目 → 略過，不 throw', w.qm.emailStatusMap([null, 5, {}, { username: '' }, { username: 'x1', role: 'user' }]), { x1: 'NO_EMAIL' });
  const miss = w.qm.missingEmails();
  const m = (un) => miss.find((x) => x.username === un);
  t('missingEmails：用名冊（gm／chairman／boardProxy／costProviders）與角色（manager1／secretary）', m('noemail') && m('noemail').roles.join() === 'secretary' && m('m1') && m('m1').roles.includes('manager1') && m('gm1') && m('gm1').reason === 'BAD_EMAIL' && !m('gm2') && !m('cons1'), JSON.stringify(miss.map((x) => [x.username, x.reason])));
  t('missingEmails：沒有任何位址', !FULL_EMAIL_RE.test(JSON.stringify(miss)));
});

// ═════════════════════════════════════════════════════════════════════════
// routes.js：用假 app 收集 handler，直接呼叫
// ═════════════════════════════════════════════════════════════════════════
function fakeApp() {
  const routes = [];
  const mk = (method) => (p, ...hs) => { routes.push({ method, path: p, handlers: hs }); };
  return { routes, get: mk('GET'), post: mk('POST') };
}
function fakeRes() {
  return {
    statusCode: 200, headers: {}, body: undefined,
    set(k, v) { if (typeof k === 'object') Object.assign(this.headers, k); else this.headers[k] = v; return this; },
    status(c) { this.statusCode = c; return this; },
    json(b) { this.body = b; return this; },
  };
}
const REQUIRE_ADMIN = function requireAdmin() {};
function mkRoutes(over) {
  const o = over || {};
  const w = o.world || mkWorld();
  const app = fakeApp();
  const auth = o.auth || { users: [{ username: 'u1', role: 'user' }, { username: 'u2', role: 'user', email: 'u2@itts.test' }, { username: 'adm', role: 'admin' }], _subscription: { plan: 'x' } };
  const saves = [];
  const logs = [];
  let flushes = 0;
  RT.registerMailRoutes(app, {
    requireAdmin: REQUIRE_ADMIN, quoteMail: o.quoteMail || w.qm, backend: 'json', env: o.env || { CRON_SECRET: 'unit-test-secret' },
    loadAuth: () => auth, saveAuth: (a) => { saves.push(a); },
    writeLog: (action, operator, target, detail) => { logs.push({ action, operator, target, detail }); },
    db: { flush: o.flush || (async () => { flushes++; }) },
  });
  const find = (method, p) => app.routes.find((r) => r.method === method && r.path === p);
  const call = async (method, p, req) => {
    const route = find(method, p);
    const res = fakeRes();
    await route.handlers[route.handlers.length - 1](Object.assign({ headers: {}, query: {}, body: {}, params: {}, session: { user: { username: 'adm' } } }, req || {}), res);
    return res;
  };
  return { w, app, auth, saves, logs, find, call, flushes: () => flushes };
}

section('9 routes.js：註冊與權限鏈', () => {
  const r = mkRoutes();
  eq('註冊的路由', r.app.routes.map((x) => x.method + ' ' + x.path).sort(), [
    'GET /api/admin/mail/config', 'GET /api/admin/mail/outbox', 'GET /api/cron/mail-outbox',
    'POST /api/admin/mail/outbox/:id/requeue', 'POST /api/admin/users/email-import/apply', 'POST /api/admin/users/email-import/preview',
  ]);
  t('後台路由都有 requireAdmin 擋在前面', r.app.routes.filter((x) => x.path.startsWith('/api/admin/')).every((x) => x.handlers[0] === REQUIRE_ADMIN && x.handlers.length === 2));
  t('Cron 路由沒有 requireAdmin／requireAuth（自己驗 secret）', r.find('GET', '/api/cron/mail-outbox').handlers.length === 1);
});

section('10 routes.js：Cron 驗證', async () => {
  const ok = (h, secret) => RT.cronAuthorized({ headers: h }, secret === undefined ? { CRON_SECRET: 's3cret-value' } : secret);
  t('正確 → true', ok({ authorization: 'Bearer s3cret-value' }) === true);
  t('Bearer 大小寫不拘', ok({ authorization: 'bearer s3cret-value' }) === true);
  [['沒有標頭', {}], ['沒有 Bearer 前綴', { authorization: 's3cret-value' }], ['Basic', { authorization: 'Basic s3cret-value' }], ['錯的 secret', { authorization: 'Bearer nope' }],
    ['少一個字', { authorization: 'Bearer s3cret-valu' }], ['多一個字', { authorization: 'Bearer s3cret-valuex' }], ['空 Bearer', { authorization: 'Bearer ' }],
    ['Bearer 之後兩段', { authorization: 'Bearer s3cret-value extra' }], ['標頭不是字串', { authorization: ['Bearer s3cret-value'] }], ['空白前綴', { authorization: ' Bearer s3cret-value' }]].forEach(([name, h]) => {
    t('拒絕：' + name, ok(h) === false);
  });
  [['未設', {}], ['空字串', { CRON_SECRET: '' }], ['不是字串', { CRON_SECRET: 123 }], ['null', { CRON_SECRET: null }]].forEach(([name, env]) => {
    t('CRON_SECRET ' + name + ' → 任何標頭都拒絕（含 Bearer 空、Bearer undefined）', ['Bearer ', 'Bearer undefined', 'Bearer null', 'Bearer ', 'Bearer 123'].every((a) => ok({ authorization: a }, env) === false));
  });
  t('不存在的 req.headers → false（不 throw）', RT.cronAuthorized({}, { CRON_SECRET: 'x' }) === false);

  // handler：未授權 401、授權後只有計數
  const fq = { enabled: () => true, config: { mode: 'log' }, outbox: { purge: async () => 3 }, drainDue: async () => ({ sent: 2, queued: 1, cancelled: 1, skipped: [{ username: 'who1', reason: 'NO_EMAIL' }], failed: [{ username: 'who2', code: 'TIMEOUT' }], errors: 0, breakerOpen: true, budgetExhausted: true }), diagnostics: () => ({}), missingEmails: () => [] };
  const r = mkRoutes({ quoteMail: fq });
  let res = await r.call('GET', '/api/cron/mail-outbox', { headers: {} });
  t('Cron handler：沒有標頭 → 401 {error}，不呼叫 drain', res.statusCode === 401 && res.body.error && res.headers['Cache-Control'] === 'no-store');
  res = await r.call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } });
  eq('Cron handler：授權後 200，只有計數與代碼', res.body, { ok: true, enabled: true, mode: 'log', purged: 3, sent: 2, queued: 1, cancelled: 1, skipped: 1, failed: 1, errors: 0, failedCodes: { TIMEOUT: 1 }, skippedReasons: { NO_EMAIL: 1 }, breakerOpen: true, budgetExhausted: true });
  t('Cron 回應不含收件人帳號', !/who1|who2/.test(JSON.stringify(res.body)));
  t('Cron 之後有 flush（稽核寫入）', r.flushes() >= 1);
  // 預算參數
  let seen = null;
  const fq2 = Object.assign({}, fq, { drainDue: async (o) => { seen = o; return { sent: 0, queued: 0, cancelled: 0, skipped: [], failed: [], errors: 0 }; } });
  await mkRoutes({ quoteMail: fq2 }).call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } });
  eq('Cron 用 drainDue({limit:20, budgetMs:16000})（預算 + 單封逾時 8 秒 < maxDuration 30 秒）', seen, { limit: 20, budgetMs: 16000 });
  // off
  const fq3 = Object.assign({}, fq, { enabled: () => false, drainDue: async () => { throw new Error('不該被呼叫'); }, outbox: { purge: async () => { throw new Error('不該被呼叫'); } } });
  res = await mkRoutes({ quoteMail: fq3 }).call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } });
  t('Cron（off）：200 enabled:false、全 0，不 drain 也不 purge', res.statusCode === 200 && res.body.enabled === false && res.body.sent === 0 && res.body.purged === 0);
  // purge 壞掉不影響回應；drain 爆炸 → 500 乾淨訊息
  const fq4 = Object.assign({}, fq, { outbox: { purge: async () => { throw new Error('purge boom'); } } });
  res = await mkRoutes({ quoteMail: fq4 }).call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } });
  t('Cron：purge 失敗不影響 drain 結果（200，purged 0）', res.statusCode === 200 && res.body.purged === 0 && res.body.sent === 2);
  const fq5 = Object.assign({}, fq, { drainDue: async () => { throw new Error('drain boom with secret-ish text'); } });
  res = await mkRoutes({ quoteMail: fq5 }).call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } });
  t('Cron：drain 爆炸 → 500，訊息不回顯例外內容', res.statusCode === 500 && res.body.ok === false && !/boom|secret-ish/.test(JSON.stringify(res.body)));
  // FIX-3：Cron 全程有界。purge 永不 resolve／db.flush 永不 resolve／drainDue 逾時（timedOut）都不會讓端點無限期等下去
  {
    const fqHang = Object.assign({}, fq, { outbox: { purge: () => new Promise(() => {}) } });
    const t0 = Date.now();
    res = await Promise.race([mkRoutes({ quoteMail: fqHang }).call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } }), sleep(7000).then(() => ({ statusCode: 'STUCK', body: {} }))]);
    const dt = Date.now() - t0;
    t('Cron：purge 永不 resolve → 約 3 秒（2.5～4.5 秒）後放手，仍 200、purged 0、drain 結果還在', res.statusCode === 200 && res.body.purged === 0 && res.body.sent === 2 && dt >= 2500 && dt < 4500, res.statusCode + ' ' + dt + 'ms ' + JSON.stringify(res.body));
    const t1 = Date.now();
    res = await Promise.race([mkRoutes({ quoteMail: fq, flush: () => new Promise(() => {}) }).call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } }), sleep(7000).then(() => ({ statusCode: 'STUCK', body: {} }))]);
    const dt1 = Date.now() - t1;
    t('Cron：db.flush 永不 resolve → 約 3 秒後放手，仍 200 並回傳 drain 結果', res.statusCode === 200 && res.body.sent === 2 && dt1 >= 2500 && dt1 < 4500, res.statusCode + ' ' + dt1 + 'ms');
    const fqTo = Object.assign({}, fq, { drainDue: async () => ({ queued: 0, sent: 0, skipped: [], failed: [], cancelled: 0, errors: 1, timedOut: true }) });
    res = await mkRoutes({ quoteMail: fqTo }).call('GET', '/api/cron/mail-outbox', { headers: { authorization: 'Bearer unit-test-secret' } });
    t('Cron：drainDue 逾時 → 回應標 timedOut:true（只有計數）', res.statusCode === 200 && res.body.timedOut === true && res.body.errors === 1, JSON.stringify(res.body));
  }
  // summaryCounts
  eq('summaryCounts：非物件 → 全 0', RT.summaryCounts(null), { sent: 0, queued: 0, cancelled: 0, skipped: 0, failed: 0, errors: 0, failedCodes: {}, skippedReasons: {} });
  t('summaryCounts：未知代碼歸 UNKNOWN', RT.summaryCounts({ failed: [{ username: 'x' }, { username: 'y', code: 5 }] }).failedCodes.UNKNOWN === 2);
});

section('11 routes.js：後台寄件匣與設定頁', async () => {
  const w = mkWorld();
  await fire(w, { type: 'E1_SUBMIT', idx: 0 });
  await fire(w, { type: 'E2_COST_REQUEST' });
  const r = mkRoutes({ world: w });
  let res = await r.call('GET', '/api/admin/mail/outbox', { query: {} });
  t('列表：rows／total／limit／offset／stats／breaker／mode', res.statusCode === 200 && res.body.rows.length === 2 && res.body.total === 2 && res.body.limit === 50 && res.body.offset === 0 && res.body.stats.total === 2 && res.body.breaker.open === false && res.body.mode === 'live' && res.body.available === true, JSON.stringify(res.body).slice(0, 200));
  t('列表：沒有完整位址', !FULL_EMAIL_RE.test(JSON.stringify(res.body)));
  res = await r.call('GET', '/api/admin/mail/outbox', { query: { type: 'E2_COST_REQUEST', limit: '1', offset: '0' } });
  t('列表：篩選 type＋limit', res.body.rows.length === 1 && res.body.rows[0].type === 'E2_COST_REQUEST' && res.body.total === 1 && res.body.limit === 1);
  res = await r.call('GET', '/api/admin/mail/outbox', { query: { limit: 'abc', offset: '-5', since: 'not a date', status: ['x'], toUser: 5 } });
  t('列表：非法的 limit／offset／since／status／toUser 被忽略（不 500）', res.statusCode === 200 && res.body.rows.length === 2);
  res = await r.call('GET', '/api/admin/mail/outbox', { query: { toUser: 'cons1' } });
  t('列表：篩選 toUser', res.body.rows.length === 1 && res.body.rows[0].toUser === 'cons1');
  // 沒有 outbox → 503
  const rn = mkRoutes({ world: mkWorld({ noOutbox: true }) });
  res = await rn.call('GET', '/api/admin/mail/outbox', { query: {} });
  t('沒有 outbox → 503 MAIL_UNAVAILABLE', res.statusCode === 503 && res.body.code === 'MAIL_UNAVAILABLE');
  res = await rn.call('POST', '/api/admin/mail/outbox/:id/requeue', { params: { id: 'mo_x' } });
  t('沒有 outbox：requeue → 503', res.statusCode === 503);
  // 儲存體壞掉 → 500 乾淨
  const bad = mkWorld({ adapter: { insert() { throw new Error('x'); }, claimDue() { throw new Error('x'); }, expireExhausted() { throw new Error('x'); }, get() { throw new Error('x'); }, list() { throw new Error('disk path C:/secret/path'); }, stats() { throw new Error('x'); }, breakerLoad() { throw new Error('x'); }, breakerSave() {} } });
  res = await mkRoutes({ world: bad }).call('GET', '/api/admin/mail/outbox', { query: {} });
  t('儲存體壞掉 → 500，不回顯例外文字（路徑等）', res.statusCode === 500 && typeof res.body.error === 'string' && !/secret|path|disk/.test(JSON.stringify(res.body)), JSON.stringify(res.body));

  // 設定頁
  w.config.redirectTo = 'qa@itts.test';
  res = await r.call('GET', '/api/admin/mail/config');
  const body = JSON.stringify(res.body);
  t('設定頁：mode／enabled／backend／cronSecretConfigured／runtime／missingEmails', res.body.config.mode === 'live' && res.body.enabled === true && res.body.backend === 'json' && res.body.cronSecretConfigured === true && res.body.runtime.bound === true && Array.isArray(res.body.missingEmails));
  t('設定頁：redirectConfigured 是布林，沒有 redirectTo 原值、密鑰、位址', res.body.config.redirectConfigured === true && !body.includes('qa@itts.test') && !body.includes('unit-test-secret') && !FULL_EMAIL_RE.test(body), body.slice(0, 200));
  t('設定頁：沒有 HTML 標籤', !/[<>]/.test(body));
  const r0 = mkRoutes({ world: w, env: {} });
  res = await r0.call('GET', '/api/admin/mail/config');
  t('設定頁：CRON_SECRET 未設 → cronSecretConfigured=false', res.body.cronSecretConfigured === false);

  // requeue：用「永久失敗」的紀錄（failed）測，單據之後走到下一關 → 重送被 STALE 取消；仍有效的 → 重送成功
  const wr = mkWorld();
  wr.behavior = () => ({ ok: false, code: 'REJECTED', permanent: true, message: 'no' });
  await fire(wr, { type: 'E1_SUBMIT', idx: 0 });
  await fire(wr, { type: 'E2_COST_REQUEST' });
  const rr = mkRoutes({ world: wr });
  const rrows = (await wr.outbox.list({})).rows;
  t('前置：兩筆都是 failed（permanent）', rrows.length === 2 && rrows.every((x) => x.status === 'failed'), JSON.stringify(rrows.map((x) => x.status)));
  const e1 = rrows.find((x) => x.type === 'E1_SUBMIT');
  const e2 = rrows.find((x) => x.type === 'E2_COST_REQUEST');
  res = await rr.call('POST', '/api/admin/mail/outbox/:id/requeue', { params: { id: 'bad id!' } });
  t('requeue：id 格式不合法 → 400 BAD_ID', res.statusCode === 400 && res.body.code === 'BAD_ID');
  res = await rr.call('POST', '/api/admin/mail/outbox/:id/requeue', { params: { id: 'mo_nope' } });
  t('requeue：不存在 → 404', res.statusCode === 404 && res.body.code === 'NOT_FOUND');
  const pend = await wr.outbox.enqueue({ type: 'E5_COST_DONE', quoteId: QID, quoteNo: 'QU-1', toUser: 'sales1', toMasked: '', dedupeKey: 'E5_COST_DONE:' + QID + ':sales1:k#done', meta: { level: null, kind: 'owner' }, nextAttemptAt: iso(T0 + 3600000) });
  res = await rr.call('POST', '/api/admin/mail/outbox/:id/requeue', { params: { id: pend.id } });
  t('requeue：pending 的紀錄 → 409 BAD_STATE（只有 failed／skipped／cancelled 可重送）', res.statusCode === 409 && res.body.code === 'BAD_STATE' && res.body.status === 'pending', JSON.stringify(res.body));
  wr.quote.approval.cur = 1; wr.quote.approval.steps[0].status = 'approved'; wr.quote.approval.steps[1].status = 'pending';
  wr.behavior = null;
  res = await rr.call('POST', '/api/admin/mail/outbox/:id/requeue', { params: { id: e1.id } });
  t('requeue：舊關卡的信 → 200，紀錄變 cancelled／STALE，沒有誤寄', res.statusCode === 200 && res.body.success === true && res.body.previousStatus === 'failed' && res.body.record.status === 'cancelled' && res.body.record.skipReason === 'STALE' && wr.sent.length === 2, JSON.stringify(res.body).slice(0, 240));
  t('requeue：回應有 drain 計數（沒有收件人）', res.body.drain && res.body.drain.cancelled === 1 && !/m1|sales1/.test(JSON.stringify(res.body.drain)));
  const lg = rr.logs.filter((x) => x.action === 'REQUEUE_MAIL');
  t('requeue：稽核 REQUEUE_MAIL（id／type／帳號／前狀態），沒有位址', lg.length === 1 && /id=mo_/.test(lg[0].detail) && /type=E1_SUBMIT/.test(lg[0].detail) && /to=m1/.test(lg[0].detail) && /prev=failed/.test(lg[0].detail) && !FULL_EMAIL_RE.test(lg[0].detail) && lg[0].operator === 'adm', JSON.stringify(lg));
  const before = wr.sent.length;
  res = await rr.call('POST', '/api/admin/mail/outbox/:id/requeue', { params: { id: e2.id } });
  t('requeue：仍有效的事件 → 重新渲染並寄出（sent）', res.statusCode === 200 && res.body.record.status === 'sent' && wr.sent.length === before + 1 && wr.sent[before].to[0] === 'cons1@itts.test', JSON.stringify(res.body.record));

  // off：requeue 後不清理，工作留 pending
  const woff = mkWorld({ env: { MAIL_MODE: 'off' } });
  const e = await woff.outbox.enqueue({ type: 'E2_COST_REQUEST', quoteId: QID, quoteNo: 'QU-1', toUser: 'cons1', toMasked: '', dedupeKey: 'E2_COST_REQUEST:' + QID + ':cons1:k#cons1', status: 'skipped', skipReason: 'MODE_OFF', meta: { level: null, kind: 'consultant' } });
  const roff = mkRoutes({ world: woff });
  res = await roff.call('POST', '/api/admin/mail/outbox/:id/requeue', { params: { id: e.id } });
  t('requeue（off）：200，drain=null，工作留在 pending', res.statusCode === 200 && res.body.drain === null && res.body.record.status === 'pending' && woff.sent.length === 0, JSON.stringify(res.body.record));
});

section('12 routes.js：Email 批次匯入', async () => {
  const r = mkRoutes();
  const text = ['# c', '', 'u1, u1@itts.test', 'u2\tu2@itts.test', 'ghost,ghost@itts.test', 'adm,adm@example.com'].join('\n');
  let res = await r.call('POST', '/api/admin/users/email-import/preview', { body: { text } });
  eq('預覽：summary', res.body.summary, { ok: 1, unchanged: 1, error: 2 });
  t('預覽：ok 的行回傳完整位址（唯一會回的地方）', res.body.rows.find((x) => x.username === 'u1').email === 'u1@itts.test' && res.body.fatal === false && res.body.limits.maxRows === 500);
  t('預覽不寫入：auth 沒變、沒有 saveAuth、沒有稽核', r.auth.users[0].email === undefined && r.saves.length === 0 && r.logs.length === 0);
  res = await r.call('POST', '/api/admin/users/email-import/preview', { body: { text: 123 } });
  t('預覽：text 不是字串 → 400', res.statusCode === 400 && res.body.code === 'BAD_REQUEST');
  res = await r.call('POST', '/api/admin/users/email-import/preview', { body: undefined });
  t('預覽：沒有 body → 400（不 throw）', res.statusCode === 400);
  res = await r.call('POST', '/api/admin/users/email-import/preview', { body: { text: Array.from({ length: 501 }, (_, i) => 'u' + i + ',u' + i + '@itts.test').join('\n') } });
  t('預覽：超過 500 行 → fatal', res.body.fatal === true && res.body.rows.length === 1 && res.body.rows[0].code === 'TOO_MANY_ROWS');

  const u1Ref = r.auth.users[0];
  res = await r.call('POST', '/api/admin/users/email-import/apply', { body: { text } });
  t('套用：更新 1 人、無變更 1、錯誤 2', res.statusCode === 200 && res.body.success === true && res.body.updated === 1 && res.body.unchanged === 1 && res.body.errors === 2, JSON.stringify(res.body));
  t('套用：就地改在原本的 user 物件上（參照不變）、其他欄位不動、saveAuth 一次', r.auth.users[0] === u1Ref && u1Ref.email === 'u1@itts.test' && u1Ref.role === 'user' && r.saves.length === 1 && r.saves[0] === r.auth && r.auth._subscription.plan === 'x');
  t('套用：錯誤行的帳號沒被改', r.auth.users[2].email === undefined);
  t('套用回應沒有完整位址', !FULL_EMAIL_RE.test(JSON.stringify(res.body)));
  const lg = r.logs.filter((x) => x.action === 'IMPORT_USER_EMAILS');
  t('套用稽核：只有摘要與帳號名稱，沒有位址', lg.length === 1 && /更新 1 人（u1）/.test(lg[0].detail) && !FULL_EMAIL_RE.test(lg[0].detail) && lg[0].operator === 'adm' && lg[0].target === 'bulk', JSON.stringify(lg));
  res = await r.call('POST', '/api/admin/users/email-import/apply', { body: { text } });
  t('再套用一次：更新 0、不 saveAuth（冪等）', res.body.updated === 0 && r.saves.length === 1);
  res = await r.call('POST', '/api/admin/users/email-import/apply', { body: { text: 'x'.repeat(200001) } });
  t('套用：超過上限 → 400 fatal', res.statusCode === 400 && res.body.fatal === true && res.body.code === 'TOO_LARGE');
  res = await r.call('POST', '/api/admin/users/email-import/apply', { body: {} });
  t('套用：沒有 text → 400', res.statusCode === 400);
  // 預覽與確認之間帳號資料變了：重新解析
  const r2 = mkRoutes();
  const txt2 = 'u1,shared@itts.test';
  r2.auth.users[1].email = 'shared@itts.test';           // 預覽時還沒人用，套用前別人先用了
  res = await r2.call('POST', '/api/admin/users/email-import/apply', { body: { text: txt2 } });
  t('套用時重新驗證：位址已被別人使用 → 不套用', res.body.updated === 0 && res.body.errors === 1 && r2.auth.users[0].email === undefined);
});

section('13 原始碼衛生（新檔案）', () => {
  const fs = require('fs');
  for (const f of ['lib/mail/quoteMail.js', 'lib/mail/routes.js']) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    t(f + '：沒有完整 email 位址、沒有 console.log', !FULL_EMAIL_RE.test(src.replace(/user\d*@example\.test/g, '')) && !/console\.log\(/.test(src), (src.match(FULL_EMAIL_RE) || [''])[0]);
    t(f + '：沒有 eval／new Function／child_process', !/\beval\(|new Function\(|child_process/.test(src));
    t(f + '：只 require 相對路徑或 Node 內建模組', (src.match(/require\(['"][^'"]+['"]\)/g) || []).every((m) => /require\(['"](\.|crypto|os|path|fs|util|assert)/.test(m)), (src.match(/require\(['"][^'"]+['"]\)/g) || []).join(','));
  }
});

main();
