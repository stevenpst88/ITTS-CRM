'use strict';
/**
 * lib/mail/quoteMail.js — 簽核信件的「整合膠水」：把 lib/quoteRoutes.js 的簽核事件接到 lib/mail 派送模組（E1–E6）
 *
 * 匯出：createQuoteMail(deps) → quoteMail
 *   deps.config     getMailConfig() 的結果（mode、timeouts、…）
 *   deps.outbox     createOutbox() 的結果；缺或建立失敗時傳 null → 整個寄信功能停用（notifyMail 什麼都不做，後台端點回「未啟用」）
 *   deps.db         { load() }（取得單據與簽核設定；與 server.js 的 db 相同）
 *   deps.loadAuth   () → { users:[…] }（帳號資料；雲端是 _auth、本機是 auth.json，由 server.js 決定）
 *   deps.writeLog   (action, operator, target, detail) → void   server.js 的 writeLog（req 參數省略）
 *   deps.transport  真傳輸（P4 的 Graph；本階段不傳 → live 模式是 NOT_CONFIGURED stub，絕不會真寄）
 *   deps.now        時鐘（預設 Date.now；測試用）
 *   deps.limits     時限（毫秒；測試用，預設值見 DEFAULT_LIMITS）：notifyMs 單次 notifyMail 的整體時限、waitMs res.end 包裝等待的外層時限、
 *                   drainCapMs drainDue（Cron／後台重送）的整體時限上限、flushMs 清理之後 db.flush 的時限
 *
 * quoteMail 的方法：
 *   bind(helpers)                   lib/quoteRoutes.js 註冊路由時呼叫一次，把「同一份」stepRecipients／getCfg／isActive／dispName 交給本檔，
 *                                   本檔不複製第二份（stepRecipients 內部已經用到 quoteRoutes 的 boardProxySet／findManager1 的結果）
 *   notifyMail(req, ctx, q, spec)   在每個 notify(...) 之後呼叫一行。永不 throw；MAIL_MODE=off 時直接返回（零成本：不讀不寫任何儲存體）；
 *                                   把 dispatch 的 promise 放進 req._mailPending，server.js 的 res.end 包裝會等它（Vercel 回應送出後實例可能被凍結）。
 *                                   整體有時限（notifyMs，預設 10 秒）：寄件匣／資料庫卡住時放手並記錄（console.warn＋稽核 QUOTE_MAIL_FAILED code=DEADLINE＋diagnostics），
 *                                   業務回應照送；被丟下的工作若已入列，之後由租約／清理接手（至少一次）。放進 _mailPending 的 promise 永不 reject、也永不長於時限
 *   waitPending(req)                server.js 的 res.end 包裝呼叫：等這個請求累積的寄信 promise，但最多 waitMs（預設 12 秒，外層保險）；永不 reject
 *   snapApproval(q, ctx)            撤回／作廢「清空 steps 之前」拍一張快照（誰收過信、原關卡），交給 notifyMail 的 E6 用
 *   pollDrain()                     GET /api/poll-bundle 前的機會式清理：每個實例至少間隔 60 秒、逾時 4 秒（之後的 db.flush 另有 3 秒時限）；永不 throw
 *   drainDue(opts)                  dispatcher.drainDue 的轉接（cron、後台重送用）；整體有時限（min(drainCapMs, 預算＋單封逾時＋1 秒)），超時回 {timedOut:true, errors:1}
 *   isStillValid(job, ev) / rebuild(job)   README §7 的契約（drainDue 與當次嘗試前一刻的檢查用）。有效性判斷是 lib/mail/validity.js 的純函式 checkValidity；
 *                                   checkJob(job) 回傳 {valid, code} 供診斷與測試
 *   missingEmails() / emailStatusMap(users)  缺 Email 清單（只含帳號、顯示名稱、角色、原因，不含位址）
 *   diagnostics()                   後台設定頁用：{ enabled, mode, bound, poll:{runs,lastAt,last} }（只有計數，沒有收件人）
 *
 * spec（notifyMail 的第 4 個參數；type 決定其他欄位）：
 *   { type:'E1_SUBMIT', idx }                       送簽（idx=0）或管理員改派一級主管；收件人＝該關 stepRecipients。改派給目前的承辦人本人時不會多寄（stepKey 不變，被去重擋掉）
 *   { type:'E3_NEXT_STEP', idx }                    核准後輪到的下一關（idx＝新的 ap.cur）
 *   { type:'E4_RESULT', idx, resultKind, reason? }  結果通知業務；resultKind＝approved｜final_approved｜rejected；idx＝剛簽的那一關；reason 只用於 rejected
 *   { type:'E2_COST_REQUEST' }                      請顧問填成本；收件人＝q.costBy
 *   { type:'E5_COST_DONE' }                         顧問完成成本；收件人＝q.owner
 *   { type:'E6_WITHDRAWN', resultKind, snap }       撤回（withdrawn）或核准後修改作廢（voided）；snap＝snapApproval 的結果
 *
 * 事件內容只由「單據現況」決定（makeEvent 是純函式），所以 rebuild(job) 重建出來的信與首次寄送逐字相同（at 用工作建立時間）；
 * 唯一例外：E6 的「原關卡」列——撤回／作廢後 steps 已清空，重試信沒有這一列（版面差異，內容不影響判斷）。
 * 金額一律是整數分，來源是送簽時凍結的 approval.derived；tierLabel 用短標籤（董事會關給 tierLevel 3＋「董事會」）。
 *
 * 限制：
 *  - 收件人過濾（排除操作者、去重、排除非在職）與 quoteRoutes.notify() 同規則，是等價小複本（notify 不能改）。
 *  - 不處理「管理員改派後，原一級主管已收過的 E1」：舊主管會在 isStillValid 重試時被 STALE 取消，但已寄出的那封不會撤回。
 *  - 本檔只被 lib/quoteRoutes.js 與 server.js 使用；不可 require npm 套件（lib/mail 規則），只 require 相對路徑與 Node 內建模組。
 */

const { createDispatcher } = require('./dispatcher');
const { normalizeMode } = require('./config');
const { missingEmailReport } = require('./userEmail');
const { safeText } = require('./safety');
const { checkValidity, stepIdxOf, e1StepKey } = require('./validity');
const QI = require('../quoteItems');

const STEP_LEVEL = Object.freeze({ mgr1: 1, gm: 2, chairman: 3, board: 'board' });
const TIER_SHORT = Object.freeze({ 1: '一級主管', 2: '總經理', 3: '董事長' });
const POLL_MIN_INTERVAL_MS = 60 * 1000;     // GET /api/poll-bundle 的機會式清理：每個實例至少間隔 60 秒
const POLL_TIMEOUT_MS = 4000;               // 清理最久等 4 秒（被丟下的工作靠租約到期後由下一個請求接手）
const POLL_DRAIN = Object.freeze({ limit: 2, budgetMs: 3000 });
const ROUTE_DRAIN = Object.freeze({ limit: 2, budgetMs: 3000 });
// 時限（FIX-3）：寄信流程不可讓簽核回應無限期等下去。傳輸本來就有 8 秒總逾時；這裡補上「儲存體操作」（outbox／資料庫）卡住時的整體上限。
const DEFAULT_LIMITS = Object.freeze({
  notifyMs: 10000,      // 單次 notifyMail（入列＋當次寄送＋路由尾端清理）
  waitMs: 12000,        // res.end 包裝等待的外層保險（略大於 notifyMs）
  drainCapMs: 22000,    // drainDue 整體時限上限（Cron 另有 purge／flush 各 3 秒，加起來仍低於 vercel.json 的 maxDuration 30 秒）
  flushMs: 3000,        // 清理之後 db.flush 的時限
});

function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
function clip(s, n) { return typeof s === 'string' ? safeText(s, n) : ''; }
function int(v, dflt) { return Number.isSafeInteger(v) && v >= 0 ? v : dflt; }

/**
 * 等 promise 但不超過 ms。永遠 resolve：{timedOut:true}｜{timedOut:false, value}｜{timedOut:false, error}。
 * timer 一定清除（不留 pending 的計時器）。p 之後才 reject 也不會變成 unhandled rejection（rejection 在這裡就被接住了）。
 */
function deadline(p, ms) {
  return new Promise((resolve) => {
    let done = false;
    let timer = null;
    const finish = (r) => { if (done) return; done = true; if (timer) clearTimeout(timer); resolve(r); };
    timer = setTimeout(() => finish({ timedOut: true }), ms);
    Promise.resolve(p).then((value) => finish({ timedOut: false, value }), (error) => finish({ timedOut: false, error }));
  });
}

/** 收件人類型：由「關卡」與收件人帳號決定（董事會關：秘書角色是 secretary，其餘代核人是 boardProxy） */
function kindOf(tier, user) {
  if (tier === 'mgr1') return 'mgr1';
  if (tier === 'gm') return 'gm';
  if (tier === 'chairman') return 'chairman';
  if (tier === 'owner') return 'owner';
  if (tier === 'board') return user && user.role === 'secretary' ? 'secretary' : 'boardProxy';
  return 'mgr1';
}

/** approval.derived → 事件的 numbers。derived 缺或不合法 → null（信上就沒有決策條，而不是整封寄不出去） */
function numbersOf(d) {
  if (!isObj(d) || typeof d.marginText !== 'string') return null;
  const pct = Number(d.marginText);
  if (!Number.isFinite(pct)) return null;
  if (!Number.isSafeInteger(d.revenueCents) || d.revenueCents < 0 || !Number.isSafeInteger(d.gpCents)) return null;
  const lv = d.level === 1 || d.level === 2 || d.level === 3 ? d.level : null;
  const tierLabel = d.board ? '董事會' : (lv ? TIER_SHORT[lv] : '');
  if (!tierLabel) return null;
  return { revenueCents: d.revenueCents, gpCents: d.gpCents, marginText: d.marginText, marginPct: pct, tierLevel: d.board ? 3 : lv, tierLabel };
}

/** 簽核關卡 → 事件的 step（label 用 step.label，也就是 QA.TIERS 的完整文字） */
function stepOf(step) {
  const lv = Object.prototype.hasOwnProperty.call(STEP_LEVEL, step.tier) ? STEP_LEVEL[step.tier] : null;
  return { level: lv, label: clip(step.label, 40) || (lv === 'board' ? '董事會' : '簽核') };
}

/** E2 的品項（只放說明／數量／單位；價格與成本在這裡就被丟掉，信件層也不會讀） */
function itemsOf(q) {
  return QI.itemRows(q.items).slice(0, 200).map((it) => {
    const row = { desc: clip(it && it.desc, 500) || '(未命名品項)', unit: clip(it && it.unit, 20) };
    const qty = it && it.qty;
    if (typeof qty === 'number' && Number.isFinite(qty) && qty >= 0) row.qty = qty;
    else if (typeof qty === 'string' && qty.trim() !== '') row.qty = clip(qty, 20);
    return row;
  });
}

function createQuoteMail(deps) {
  const d = isObj(deps) ? deps : {};
  const config = isObj(d.config) ? d.config : {};
  const outbox = d.outbox && typeof d.outbox.enqueue === 'function' ? d.outbox : null;
  const nowMs = () => { const v = typeof d.now === 'function' ? Number(d.now()) : Date.now(); return Number.isFinite(v) ? v : Date.now(); };
  const warn = (what, e) => { try { console.warn('[quote mail] ' + what + ':', e && e.message ? e.message : e); } catch (_) { /* ignore */ } };

  let H = null;                                  // 來自 lib/quoteRoutes.js 的輔助函式（bind）
  const poll = { runs: 0, lastAt: null, last: null, running: false, lastStartMs: 0 };
  const lim = Object.assign({}, DEFAULT_LIMITS);
  if (isObj(d.limits)) Object.keys(DEFAULT_LIMITS).forEach((k) => { const v = d.limits[k]; if (typeof v === 'number' && Number.isFinite(v) && v > 0) lim[k] = v; });
  const deadlines = { notify: 0, wait: 0, drain: 0, flush: 0 };      // 發生過幾次時限（診斷用；只有計數）

  const enabled = () => !!outbox && normalizeMode(config.mode) !== 'off';

  // ── 即時脈絡（drainDue、isStillValid、rebuild 沒有 HTTP 請求可用）─────────
  function liveCtx() {
    if (!H) throw new Error('quoteMail 尚未 bind');
    const data = d.db.load();
    const auth = d.loadAuth();
    const users = {};
    ((auth && auth.users) || []).forEach((u) => { if (u && typeof u.username === 'string') users[u.username] = u; });
    return { data, users, cfg: H.getCfg(data) };
  }
  const findQuote = (data, id) => ((data && data.quotations) || []).find((x) => x && x.id === id) || null;
  const labelOf = (users, un) => clip(H.dispName(users, un), 100);

  // ── 事件 ────────────────────────────────────────────────────────────────
  /** 單據現況 → 事件。純函式（同樣的單據＋同樣的 o 永遠得到同樣的事件），retry 重建靠這個保證與首次寄送相同。 */
  function makeEvent(ctx, q, type, o) {
    const ap = isObj(q.approval) ? q.approval : {};
    const ev = {
      type, quoteId: q.id, quoteNo: q.quoteNo,
      projectName: clip(q.projectName, 300), company: clip(q.company, 300), ownerLabel: labelOf(ctx.users, q.owner),
      at: o.at || new Date(nowMs()).toISOString(), stepKey: o.stepKey,
    };
    const setActor = (un) => { const l = un ? labelOf(ctx.users, un) : ''; if (l) ev.actor = { label: l }; };
    const steps = Array.isArray(ap.steps) ? ap.steps : [];
    switch (type) {
      case 'E1_SUBMIT':
      case 'E3_NEXT_STEP': {
        const step = steps[o.idx];
        if (!step) return null;
        ev.step = stepOf(step);
        ev.numbers = numbersOf(ap.derived);
        setActor(type === 'E1_SUBMIT' ? ap.submittedBy : (steps[o.idx - 1] && steps[o.idx - 1].by));
        return ev;
      }
      case 'E4_RESULT': {
        const step = steps[o.idx];
        if (step) { ev.step = stepOf(step); setActor(step.by); }
        ev.result = o.resultKind === 'rejected' && o.reason ? { kind: o.resultKind, reason: clip(o.reason, 2000) } : { kind: o.resultKind };
        return ev;
      }
      case 'E2_COST_REQUEST':
        ev.items = itemsOf(q);
        setActor(q.owner);
        return ev;
      case 'E5_COST_DONE':
        setActor(q.costBy);
        return ev;
      case 'E6_WITHDRAWN':
        if (o.originStep) ev.step = o.originStep;
        ev.result = { kind: o.resultKind };
        setActor(q.owner);
        return ev;
      default:
        return null;
    }
  }

  /** notify() 的前置過濾（排除操作者、去重、排除非在職）＋ 對應收件人類型 */
  function toRecipients(ctx, actor, names, kindFn) {
    const seen = new Set();
    const out = [];
    (Array.isArray(names) ? names : []).forEach((un) => {
      if (!un || typeof un !== 'string' || un === actor || seen.has(un) || !H.isActive(ctx.users[un])) return;
      seen.add(un);
      out.push({ username: un, kind: kindFn(un) });
    });
    return out;
  }

  /** spec → { ev, recipients }；不該寄（資料不足）回 null */
  function plan(ctx, q, spec) {
    const actor = ctx.me;
    const ap = isObj(q.approval) ? q.approval : {};
    switch (spec.type) {
      case 'E1_SUBMIT':
      case 'E3_NEXT_STEP': {
        const idx = int(spec.idx, 0);
        const step = Array.isArray(ap.steps) ? ap.steps[idx] : null;
        if (!step || typeof ap.submittedAt !== 'string') return null;
        // E1 的 stepKey 帶改派標記（管理員改派過就多 @<改派時間>）：A→B→A 時 A 會再收到一封；同一次改派重複觸發仍是同一把去重鍵。
        // 改派給目前的承辦人本人（歷史 REASSIGN 的 meta.from===meta.to）不算新的改派：標記不變 → 同一把去重鍵 → 不重複寄（validity.reassignEpoch 負責辨識）。E3 沒有改派
        const ev = makeEvent(ctx, q, spec.type, { idx, stepKey: spec.type === 'E1_SUBMIT' ? e1StepKey(ap, idx) : ap.submittedAt + '#' + idx });
        return ev && { ev, recipients: toRecipients(ctx, actor, H.stepRecipients(step, ctx), (un) => kindOf(step.tier, ctx.users[un])) };
      }
      case 'E4_RESULT': {
        const idx = int(spec.idx, 0);
        if (typeof ap.submittedAt !== 'string') return null;
        const ev = makeEvent(ctx, q, 'E4_RESULT', { idx, resultKind: spec.resultKind, reason: spec.reason, stepKey: ap.submittedAt + '#r:' + spec.resultKind + ':' + idx });
        return ev && { ev, recipients: toRecipients(ctx, actor, [q.owner], () => 'owner') };
      }
      case 'E2_COST_REQUEST': {
        const cf = isObj(q.costFlow) ? q.costFlow : {};
        if (!q.costBy || !cf.requestedAt) return null;
        const ev = makeEvent(ctx, q, 'E2_COST_REQUEST', { stepKey: cf.requestedAt + '#' + q.costBy });
        return ev && { ev, recipients: toRecipients(ctx, actor, [q.costBy], () => 'consultant') };
      }
      case 'E5_COST_DONE': {
        const cf = isObj(q.costFlow) ? q.costFlow : {};
        if (!cf.filledAt) return null;
        const ev = makeEvent(ctx, q, 'E5_COST_DONE', { stepKey: cf.filledAt + '#done' });
        return ev && { ev, recipients: toRecipients(ctx, actor, [q.owner], () => 'owner') };
      }
      case 'E6_WITHDRAWN': {
        const snap = spec.snap;
        if (!isObj(snap) || typeof snap.submittedAt !== 'string' || !snap.submittedAt) return null;
        const voided = spec.resultKind === 'voided';
        const list = (voided ? [] : snap.current).concat(snap.signers).concat(voided ? [{ username: q.owner, tier: 'owner' }] : []);
        const tierBy = new Map();
        list.forEach((x) => { if (x && typeof x.username === 'string' && !tierBy.has(x.username)) tierBy.set(x.username, x.tier); });
        const ev = makeEvent(ctx, q, 'E6_WITHDRAWN', { resultKind: spec.resultKind, originStep: snap.originStep || null, stepKey: snap.submittedAt + '#' + spec.resultKind });
        return ev && { ev, recipients: toRecipients(ctx, actor, Array.from(tierBy.keys()), (un) => kindOf(tierBy.get(un), ctx.users[un])) };
      }
      default:
        return null;
    }
  }

  // ── 派送器 ──────────────────────────────────────────────────────────────
  const stepKeyOf = (job) => String(job.dedupeKey || '').split(':').slice(3).join(':');   // type:quoteId:user:stepKey（user 的 ':' 已跳脫成 %3A）

  /**
   * README §7：只有回傳 true 才寄；false → 取消（STALE）；丟例外 → 稍後重試。
   * 判斷全部在 lib/mail/validity.js 的純函式 checkValidity（E1～E6 各有自己的過期條件；本檔只負責讀單據現況、把 stepRecipients 接進去）。
   */
  function checkJob(job) {
    const live = liveCtx();
    const q = findQuote(live.data, job.quoteId);
    return checkValidity(job, q, { stepRecipients: (step) => H.stepRecipients(step, live) });
  }
  function isStillValid(job) {
    return checkJob(job).valid === true;
  }

  /** README §7：依工作重建事件（單據不存在或該關資料已清空 → null → 取消 GONE） */
  function rebuild(job) {
    const live = liveCtx();
    const q = findQuote(live.data, job.quoteId);
    if (!q) return null;
    const ap = isObj(q.approval) ? q.approval : {};
    const sk = stepKeyOf(job);
    const base = { stepKey: sk, at: job.createdAt };           // 用工作建立時間：重試信與首次嘗試逐字相同
    const ctx = live;
    if (job.type === 'E1_SUBMIT' || job.type === 'E3_NEXT_STEP') {
      const idx = stepIdxOf(sk);                                  // <submittedAt>#<idx>[@<改派時間>]
      if (idx < 0 || !Array.isArray(ap.steps) || !ap.steps[idx]) return null;
      const ev = makeEvent(ctx, q, job.type, Object.assign(base, { idx }));
      return ev && { ev, kind: kindOf(ap.steps[idx].tier, live.users[job.toUser]) };
    }
    if (job.type === 'E2_COST_REQUEST') {
      const ev = makeEvent(ctx, q, job.type, base);
      return ev && { ev, kind: 'consultant' };
    }
    if (job.type === 'E4_RESULT') {
      const m = /#r:([a-z_]+):([0-9]+)$/.exec(sk);
      if (!m) return null;
      const idx = Number(m[2]);
      const step = Array.isArray(ap.steps) ? ap.steps[idx] : null;
      const ev = makeEvent(ctx, q, job.type, Object.assign(base, { idx, resultKind: m[1], reason: step && m[1] === 'rejected' ? step.comment : undefined }));
      return ev && { ev, kind: 'owner' };
    }
    if (job.type === 'E5_COST_DONE') {
      const ev = makeEvent(ctx, q, job.type, base);
      return ev && { ev, kind: 'owner' };
    }
    if (job.type === 'E6_WITHDRAWN') {
      const m = /#(withdrawn|voided)$/.exec(sk);
      if (!m) return null;
      const ev = makeEvent(ctx, q, job.type, Object.assign(base, { resultKind: m[1] }));
      return ev && { ev };                                        // kind 沿用工作建立時記下的 meta.kind
    }
    return null;
  }

  const dispatcher = outbox ? createDispatcher({
    config, outbox, transport: d.transport,
    getUsers: () => d.loadAuth().users,
    isStillValid: (job) => isStillValid(job),
    rebuild: (job) => rebuild(job),
    writeLog: (action, operator, target, detail) => { if (typeof d.writeLog === 'function') d.writeLog(action, operator, target, detail); },
    now: d.now,
  }) : null;

  // ── 對外 ────────────────────────────────────────────────────────────────
  function bind(helpers) {
    if (!isObj(helpers) || ['getCfg', 'stepRecipients', 'isActive', 'dispName'].some((k) => typeof helpers[k] !== 'function')) {
      throw new Error('quoteMail.bind 需要 getCfg／stepRecipients／isActive／dispName');
    }
    H = helpers;
  }

  function track(req, p) {
    if (!req || typeof req !== 'object') return;
    if (!Array.isArray(req._mailPending)) req._mailPending = [];
    req._mailPending.push(p);
  }
  /** 同一個請求內多次 notifyMail（核准會同時寄 E4 與 E3）只清理一次 */
  function drainOnce(req) {
    if (!req || typeof req !== 'object') return dispatcher.drainDue(ROUTE_DRAIN);
    if (!req._mailDrain) req._mailDrain = dispatcher.drainDue(ROUTE_DRAIN);
    return req._mailDrain;
  }

  /** 時限到了：只記錄（console.warn、diagnostics 計數、稽核 QUOTE_MAIL_FAILED code=DEADLINE），不影響業務回應。稽核寫入失敗也不往外丟 */
  function noteDeadline(kind, ev, operator) {
    deadlines[kind] += 1;
    warn('deadline ' + kind, '寄信流程超過時限，放手不等（已入列的工作之後由租約／清理接手）');
    try {
      if (kind === 'notify' && typeof d.writeLog === 'function') {
        d.writeLog('QUOTE_MAIL_FAILED', safeText(operator, 60) || 'system', safeText(ev && ev.quoteNo, 40), 'type=' + (ev && ev.type) + ' code=DEADLINE');
      }
    } catch (_) { /* 稽核失敗不影響回應 */ }
  }

  function notifyMail(req, ctx, q, spec) {
    try {
      if (!enabled() || !H || !isObj(spec) || !q || !ctx) return;
      const job = plan(ctx, q, spec);
      if (!job || !job.recipients.length) return;
      const operator = labelOf(ctx.users, ctx.me);
      const opts = { actorUsername: ctx.me, operatorLabel: operator };
      // 整體時限：dispatcher 對傳輸有 8 秒總逾時、對 getUsers／isStillValid／render 有 5 秒輔助逾時，但對 outbox／資料庫操作沒有——
      // 這裡補上。放進 _mailPending 的 promise 一定在 notifyMs 內 settle、而且永不 reject（deadline() 內就接住 rejection）
      const work = dispatcher.dispatch(job.ev, job.recipients, opts).then(() => drainOnce(req));
      const p = deadline(work, lim.notifyMs).then((r) => {
        try {
          if (r.timedOut) noteDeadline('notify', job.ev, operator);
          else if (r.error) warn('dispatch', r.error);
        } catch (_) { /* 記錄失敗不影響回應 */ }
      });
      track(req, p);
    } catch (e) {
      warn('notifyMail', e);
    }
  }

  /** server.js 的 res.end 包裝：等這個請求累積的寄信 promise，最多 waitMs（外層保險）。永不 reject，沒有寄信時幾乎零成本 */
  function waitPending(req) {
    const list = req && typeof req === 'object' && Array.isArray(req._mailPending) ? req._mailPending.slice() : [];
    if (!list.length) return Promise.resolve({ timedOut: false, pending: 0 });
    return deadline(Promise.allSettled(list), lim.waitMs).then((r) => {
      if (r.timedOut) { deadlines.wait += 1; warn('deadline wait', '等待寄信超過 ' + lim.waitMs + 'ms，先送出回應'); }
      return { timedOut: !!r.timedOut, pending: list.length };
    });
  }

  function snapApproval(q, ctx) {
    try {
      if (!H || !q) return null;
      const ap = isObj(q.approval) ? q.approval : {};
      const steps = Array.isArray(ap.steps) ? ap.steps : [];
      const signers = steps.filter((s) => s && s.status === 'approved' && s.by).map((s) => ({ username: s.by, tier: s.tier }));
      const curStep = ap.state === 'pending' ? steps[ap.cur] : null;
      const current = curStep ? H.stepRecipients(curStep, ctx).filter(Boolean).map((un) => ({ username: un, tier: curStep.tier })) : [];
      const origin = curStep || steps[steps.length - 1] || null;
      return { submittedAt: typeof ap.submittedAt === 'string' ? ap.submittedAt : null, originStep: origin ? stepOf(origin) : null, current, signers };
    } catch (e) {
      warn('snapApproval', e);
      return null;
    }
  }

  /** 整體時限 = min(drainCapMs, 預算 + 單封總逾時 + 1 秒)：預算只管「還要不要領下一筆」，已開始的寄送最多再 8 秒；儲存體卡住時也不會超過這個上限 */
  async function drainDue(opts) {
    if (!dispatcher) return { queued: 0, sent: 0, skipped: [], failed: [], cancelled: 0, errors: 0, disabled: true };
    const o = isObj(opts) ? opts : {};
    const totalMs = (config.timeouts && Number(config.timeouts.totalMs)) || 8000;
    const budget = typeof o.budgetMs === 'number' && o.budgetMs > 0 ? o.budgetMs : Math.min(20000, totalMs * 2);
    const ms = Math.min(lim.drainCapMs, budget + totalMs + 1000);
    const r = await deadline(dispatcher.drainDue(o), ms);
    if (r.timedOut || r.error) {
      if (r.timedOut) { deadlines.drain += 1; warn('deadline drain', '清理超過 ' + ms + 'ms，放手不等'); }
      return { queued: 0, sent: 0, skipped: [], failed: [], cancelled: 0, errors: 1, timedOut: !!r.timedOut };
    }
    return r.value;
  }

  /** GET /api/poll-bundle 的機會式清理：關閉時零成本；每個實例至少間隔 60 秒；最久等 4 秒；永不 throw。回傳是否真的跑了清理 */
  async function pollDrain() {
    try {
      if (!enabled() || !H) return false;
      const t = nowMs();
      if (poll.running || (poll.lastStartMs && t - poll.lastStartMs < POLL_MIN_INTERVAL_MS)) return false;
      poll.running = true;
      poll.lastStartMs = t;
      let timer = null;
      try {
        const timeout = new Promise((resolve) => { timer = setTimeout(() => resolve('TIMEOUT'), POLL_TIMEOUT_MS); });
        const r = await Promise.race([dispatcher.drainDue(POLL_DRAIN), timeout]);
        poll.runs += 1;
        poll.lastAt = new Date(nowMs()).toISOString();
        poll.last = r === 'TIMEOUT' ? { timedOut: true } : { sent: r.sent, failed: r.failed.length, skipped: r.skipped.length, cancelled: r.cancelled, queued: r.queued, errors: r.errors };
        if (r !== 'TIMEOUT' && d.db && typeof d.db.flush === 'function') {
          const fr = await deadline(Promise.resolve().then(() => d.db.flush()), lim.flushMs);      // 稽核寫入失敗或卡住都不影響輪詢
          if (fr.timedOut) { deadlines.flush += 1; warn('deadline flush', 'db.flush 超過 ' + lim.flushMs + 'ms，先放手'); }
        }
      } finally {
        if (timer) clearTimeout(timer);
        poll.running = false;
      }
      return true;
    } catch (e) {
      warn('pollDrain', e);
      return false;
    }
  }

  function missingEmails() {
    if (!H) return [];
    const live = liveCtx();
    return missingEmailReport({ users: d.loadAuth().users, roster: live.cfg.roster, config });
  }

  /** 每個帳號的 Email 狀態：OK｜NO_EMAIL｜BAD_EMAIL｜DOMAIN_NOT_ALLOWED（用 missingEmailReport 的同一套規則；不回傳位址；停用帳號一律 OK） */
  function emailStatusMap(users) {
    const list = Array.isArray(users) ? users.filter((u) => isObj(u) && typeof u.username === 'string' && u.username !== '') : [];
    const rep = missingEmailReport({ users: list, roster: { gm: list.map((u) => u.username) }, config });
    const out = {};
    list.forEach((u) => { out[u.username] = 'OK'; });
    rep.forEach((r) => { out[r.username] = r.reason; });
    return out;
  }

  function diagnostics() {
    return { enabled: enabled(), available: !!outbox, mode: normalizeMode(config.mode), bound: !!H, poll: { runs: poll.runs, lastAt: poll.lastAt, last: poll.last }, deadlines: Object.assign({}, deadlines) };
  }

  return {
    config, outbox, dispatcher,
    enabled, bind, notifyMail, waitPending, snapApproval, pollDrain, drainDue, isStillValid, checkJob, rebuild,
    missingEmails, emailStatusMap, diagnostics,
  };
}

module.exports = { createQuoteMail, kindOf, numbersOf, stepOf, itemsOf, deadline, DEFAULT_LIMITS, POLL_MIN_INTERVAL_MS, POLL_TIMEOUT_MS };
