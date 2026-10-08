'use strict';
/**
 * lib/mail/dispatcher.js — 信件派送器：事件 + 收件人 → 入列 → 渲染 → 寄送 → 回報
 *
 * 匯出與簽名：
 *   createDispatcher(deps) → { dispatch, drainDue }
 *     deps.config        getMailConfig() 的結果（mode、timeouts、retry、allowedDomains…）
 *     deps.outbox        createOutbox() 的結果
 *     deps.transport     「真的會寄信」的傳輸（P4 的 Graph；測試用假傳輸）。不是 selectTransport 的輸出！
 *                        實際使用哪個傳輸由 selectTransport(config, {realTransport}) 依有效模式決定（見 transports.js）：
 *                        off 不寄、log 只寫預覽、redirect 一定改寫收件人、live 才用這個傳輸
 *     deps.getUsers      () → users（陣列或 username 為鍵的物件；可為 Promise）
 *     deps.isStillValid  (job, ev) → boolean（可為 Promise）。寄送前一刻呼叫；只有回傳 true 才寄，其他（含丟例外／逾時）都不寄
 *                        回傳 false → cancel(id,'STALE')；丟例外／逾時 → 稍後重試（STALE_CHECK）。未提供則視為有效
 *     deps.writeLog      (action, operatorLabel, quoteNo, detail) → void|Promise。只用於失敗／略過
 *     deps.render        (ev, viewer, ctx) → {subject, html, text, meta}（可為 Promise）。預設用 ./render 的 renderMail
 *     deps.rebuild       (job) → null | {ev, kind?}（drainDue 用，可在呼叫時覆寫）
 *     deps.now           時鐘（預設 Date.now）
 *     deps.concurrency   同時寄送數（預設 3；Graph 對單一信箱限 4 個並行）
 *     deps.auxTimeoutMs  getUsers／isStillValid／render／rebuild／writeLog 各自的逾時（預設 5 秒）
 *     deps.budgetMs      單次 dispatch 的時間預算（預設 15 秒：最壞情況＝15 秒內才開始的最後一封再加單封逾時 8 秒＝23 秒，低於專案 maxDuration 30 秒；超過預算就不再當次寄送，留 pending 給 drainDue）
 *
 *   dispatch(ev, recipients, opts) → Promise<Summary>      永不 throw、永不 reject
 *     recipients  [{username, kind}]（kind 來自呼叫端，見 visibility.js）；opts：{actorUsername, operatorLabel}
 *   drainDue({limit=5, rebuild?, budgetMs?}) → Promise<Summary>   領取到期工作並重寄；永不 throw、永不 reject
 *     limit     最多處理幾筆（1–50）。budgetMs  時間預算（毫秒），預設 min(20 秒, 2 × config.timeouts.totalMs)＝16 秒。
 *     領取方式：先清掃「租約過期且本輪次數用盡」的 sending（改標 failed/LEASE_EXPIRED，記進 Summary.failed 並寫 QUOTE_MAIL_FAILED），
 *     之後每個工作者「預算還有且熔斷沒開 → 才領一筆 → 處理完再領下一筆」，所以沒輪到的工作不會被預先領走、不會白扣嘗試次數；
 *     到預算或熔斷中途開啟就停止領取（Summary 帶 budgetExhausted:true／breakerOpen:true），剩下的原封不動留在 pending 給下一次。
 *     已經開始的寄送會做完（各自受 totalMs 約束），所以最壞耗時＝預算 + 一封寄送逾時；呼叫端的平台逾時（Vercel maxDuration）要大於此值。
 *   Summary = {queued, sent, skipped:[{username,reason}], failed:[{username,code}], cancelled, errors}
 *     errors 是「內部例外／環境問題」的計數（不是收件人層級的失敗）；drainDue 在熔斷開啟時另帶 breakerOpen:true、
 *     因預算用完而停止時另帶 budgetExhausted:true（只在發生時才有這些欄位）
 *
 * 流程（每位收件人獨立、互不影響）：
 *   mode off → 全部記 skipped/MODE_OFF（outbox 記一筆，不渲染、不查帳號）
 *   否則  resolveRecipients（排除操作者、停用、無 email、網域不在白名單）
 *         → enqueue（以 dedupeKey 冪等；重複觸發記 DUPLICATE，不重寄）
 *         → 熔斷開著／超過時間預算 → 留 pending（熔斷時 nextAttemptAt＝熔斷結束）
 *         → 否則 claim（當次嘗試）→ isStillValid → render → 寄送（總逾時 config.timeouts.totalMs）→ markSent／markFailed
 *   與規格文字的差異（結果相同）：規格寫「render 在 enqueue 之前」；這裡先 enqueue 再 render——事件先落地，程序在渲染中途死掉也不會漏信，
 *   重複觸發也不會白白渲染。render 失敗記 failed/RENDER（permanent）。
 *
 * 稽核（writeLog）：只在「最終失敗」與「值得注意的略過」時寫：
 *   QUOTE_MAIL_FAILED（本輪次數用盡、permanent、渲染失敗、事件不合法…）、QUOTE_MAIL_SKIPPED（NO_EMAIL／BAD_EMAIL／DOMAIN_NOT_ALLOWED／UNKNOWN_USER／INACTIVE）。
 *   不寫：成功、MODE_OFF、MODE_LOG、重複（DUP／DUPLICATE）、操作者本人（ACTOR）、尚未用盡的重試。
 *   detail 格式 `type=E1_SUBMIT to=<帳號> code=TIMEOUT`，只含事件型別、收件帳號（含 @ 的帳號會遮罩）、錯誤碼／略過原因；
 *   不含 email、金額、專案／客戶名、信件內容。writeLog 本身丟例外或卡住都不會影響派送。
 *
 * 模式與紀錄：log 模式（以及 redirect 模式但內層只是 log 傳輸）寄「成功」時，outbox 記為 skipped/MODE_LOG（沒有真的送出），
 * 之後切到 redirect／live 可由 admin 用 requeue 重送，不會被去重擋掉。
 *
 * 語意是「至少一次」：寄出成功但 markSent 沒寫成功（或租約過期）時，之後可能重寄一次（極端 race 下重複一封；不會靜默丟失）。
 * 租約過期且本輪次數用盡的工作改標 failed/LEASE_EXPIRED 時，drainDue 會把它列入 Summary.failed 並寫 QUOTE_MAIL_FAILED（code=LEASE_EXPIRED），
 * 後台列表／requeue 也看得到；它不是靜默狀態轉換。
 * 錯誤訊息寫進 outbox 前會先經 scrubMessage，並把主旨／專案名稱／客戶名／業務名／收件位址這些已知字串遮掉。
 */

const { validateEvent, dedupeKey, EVENT_TYPES } = require('./events');
const { resolveRecipients } = require('./recipients');
const { normalizeMode } = require('./config');
const { selectTransport, createLogTransport, normalizeResult, failure, validateMessage } = require('./transports');
const { HEALTH_CODES } = require('./outbox');
const { scrubMessage } = require('./scrub');
const { safeText, maskEmail } = require('./safety');

const AUDIT_SKIP = Object.freeze(['NO_EMAIL', 'BAD_EMAIL', 'DOMAIN_NOT_ALLOWED', 'UNKNOWN_USER', 'INACTIVE']);
const MAX_RECIPIENTS = 200;
const DEFAULT_AUX_TIMEOUT_MS = 5000;

function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
function clampInt(v, min, max, dflt) { return (typeof v === 'number' && isFinite(v)) ? Math.max(min, Math.min(max, Math.floor(v))) : dflt; }

/** 等 promise 但不超過 ms；永遠 resolve，結果是 {ok:true,value}|{ok:false,error}|{ok:false,timedOut:true}。timer 一定清除。 */
function withTimeout(p, ms) {
  return new Promise((resolve) => {
    let done = false;
    const timer = setTimeout(() => { if (!done) { done = true; resolve({ ok: false, timedOut: true }); } }, ms);
    p.then(
      (value) => { if (!done) { done = true; clearTimeout(timer); resolve({ ok: true, value }); } },
      (error) => { if (!done) { done = true; clearTimeout(timer); resolve({ ok: false, error }); } }
    );
  });
}
/** 呼叫可能同步丟例外、可能回傳 Promise、可能永不 resolve 的函式。 */
function guarded(fn, ms) {
  let p;
  try { p = Promise.resolve(fn()); } catch (e) { return Promise.resolve({ ok: false, error: e }); }
  return withTimeout(p, ms);
}

async function pool(items, n, worker) {
  let i = 0;
  const runners = [];
  for (let k = 0; k < Math.min(n, items.length); k++) {
    runners.push((async () => {
      for (;;) {
        const idx = i++;
        if (idx >= items.length) return;
        await worker(items[idx], idx);
      }
    })());
  }
  await Promise.all(runners);
}

function newSummary() { return { queued: 0, sent: 0, skipped: [], failed: [], cancelled: 0, errors: 0 }; }
function typeOf(ev) { try { return ev && EVENT_TYPES.indexOf(ev.type) >= 0 ? ev.type : 'UNKNOWN'; } catch (e) { return 'UNKNOWN'; } }
function quoteNoOf(ev) { try { return ev && typeof ev.quoteNo === 'string' ? safeText(ev.quoteNo, 40) : ''; } catch (e) { return ''; } }
function userInDetail(u) { return (typeof u === 'string' && u.indexOf('@') >= 0) ? maskEmail(u) : safeText(u, 64); }

function createDispatcher(deps) {
  const d = isObj(deps) ? deps : {};
  const config = isObj(d.config) ? d.config : {};
  const outbox = d.outbox && typeof d.outbox.enqueue === 'function' ? d.outbox : null;
  const nowFn = typeof d.now === 'function' ? d.now : Date.now;
  const concurrency = clampInt(d.concurrency, 1, 8, 3);
  const budgetMs = clampInt(d.budgetMs, 1, 600000, 15000);
  const logT = createLogTransport({ config, rootDir: d.rootDir, fs: d.fs, now: nowFn });

  const nowMs = () => { const v = Number(nowFn()); return Number.isFinite(v) ? v : Date.now(); };
  const auxMs = () => clampInt(d.auxTimeoutMs, 1, 60000, DEFAULT_AUX_TIMEOUT_MS);   // getUsers／isStillValid／render／rebuild／writeLog 各自的逾時
  const totalMs = () => clampInt(config.timeouts && config.timeouts.totalMs, 1, 120000, 8000);
  const leaseSec = () => Math.max(60, Math.ceil(totalMs() / 1000) * 3 + 30);
  // drainDue 的預設時間預算：min(20 秒, 2 × 單封總逾時)＝預設 16 秒。到點就不再領取新工作；
  // 最壞情況總耗時＝預算 + 最後一批已開始的寄送（各自受 totalMs 約束，預設 8 秒）≈ 24 秒，低於專案 maxDuration 30 秒（假設輔助呼叫很快，見檔頭）
  const defaultDrainBudgetMs = () => Math.min(20000, totalMs() * 2);
  const currentMode = () => normalizeMode(config.mode);
  const transportNow = () => selectTransport(config, { realTransport: d.transport, logTransport: logT, fetchImpl: d.fetchImpl, now: nowFn });

  function render(ev, viewer) {
    const fn = typeof d.render === 'function' ? d.render : (e, v, c) => require('./render').renderMail(e, v, c);
    return fn(ev, viewer, { config, now: nowMs() });
  }

  // ── 稽核 ────────────────────────────────────────────────────────────────
  async function audit(action, operator, quoteNo, type, username, what) {
    if (typeof d.writeLog !== 'function') return;
    try {
      const detail = 'type=' + type + (username ? ' to=' + userInDetail(username) : '') + ' ' + what;
      await guarded(() => d.writeLog(action, operator, quoteNo, detail), auxMs());
    } catch (e) { /* 稽核失敗不影響派送 */ }
  }

  async function breakerState() {
    try { const s = await outbox.breaker.state(); return { open: !!s.open, until: s.until || null }; } catch (e) { return { open: false, until: null }; }
  }
  async function breakerRecord(ok, code) {
    try { await outbox.breaker.record(ok, code); } catch (e) { /* 熔斷狀態寫不進去不影響這封信 */ }
  }

  /** 讓一個「已領取（sending）」的工作失敗：依 permanent／次數決定重試或最終失敗；最終失敗才算 failed 並寫稽核。 */
  async function failJob(job, c, summary, code, message, extra) {
    const x = extra || {};
    let res;
    try {
      res = await outbox.markFailed(job.id, { code, msg: scrubMessage(message, c.known), permanent: x.permanent === true, retryAfterSec: x.retryAfterSec, attempts: job.attempts });
    } catch (e) { summary.errors += 1; return; }
    if (!res || !res.ok) return;                 // 租約已被別人接手或狀態已變：由對方負責，這邊不重複計數
    if (res.final) {
      summary.failed.push({ username: job.toUser, code });
      await audit('QUOTE_MAIL_FAILED', c.operator, c.quoteNo, c.type, job.toUser, 'code=' + code);
    } else {
      summary.queued += 1;
    }
  }

  /**
   * 對一個「已領取」的工作完成：確認仍有效 → 渲染 → 寄送 → 回報。
   * c = {ev, kind, label, email, operator, type, quoteNo, known:[…]}
   */
  async function sendClaimed(job, c, summary) {
    if (typeof d.isStillValid === 'function') {
      const r = await guarded(() => d.isStillValid(job, c.ev), auxMs());
      if (!r.ok) { await failJob(job, c, summary, 'STALE_CHECK', '無法確認單據狀態', { permanent: false }); return; }
      if (r.value !== true) {
        const cr = await outbox.cancel(job.id, 'STALE');
        if (cr && cr.ok) summary.cancelled += 1;
        return;
      }
    }

    const rr = await guarded(() => render(c.ev, { username: job.toUser, label: c.label, kind: c.kind }), auxMs());
    const mail = rr.ok ? rr.value : null;
    if (!rr.ok || !isObj(mail) || typeof mail.subject !== 'string' || (typeof mail.html !== 'string' && typeof mail.text !== 'string')) {
      const why = rr.ok ? '渲染結果格式不合法' : (rr.timedOut ? '渲染逾時' : '渲染失敗：' + (rr.error && rr.error.code ? rr.error.code : 'ERR'));
      await failJob(job, c, summary, 'RENDER', why, { permanent: true });
      return;
    }
    // 派送器自己也驗證一次訊息（主旨換行注入、收件位址格式、大小），不依賴傳輸層；之後送出去的是驗證後的「乾淨副本」
    const vm = validateMessage({ to: [c.email], subject: mail.subject, html: typeof mail.html === 'string' ? mail.html : '', text: typeof mail.text === 'string' ? mail.text : '', tag: c.type });
    c.known = c.known.concat([mail.subject, c.email]);
    if (!vm.ok) { await failJob(job, c, summary, 'BAD_MESSAGE', vm.error, { permanent: true }); return; }
    const msg = vm.msg;

    // 寄送（總逾時）。逾時會 abort 並視為 TIMEOUT；傳輸丟例外／回傳亂格式都包成失敗結果
    const tr = transportNow();
    const ac = new AbortController();
    let timer = null;
    const timeoutP = new Promise((resolve) => {
      timer = setTimeout(() => { try { ac.abort(); } catch (e) { /* ignore */ } resolve(failure('TIMEOUT', '寄送逾時（' + totalMs() + 'ms）')); }, totalMs());
    });
    let result;
    try {
      let sp;
      try { sp = Promise.resolve(tr.send(msg, { timeouts: config.timeouts, signal: ac.signal })); } catch (e) { sp = Promise.resolve(failure('NETWORK', '傳輸拋出例外')); }
      result = await Promise.race([sp.then(normalizeResult, () => failure('NETWORK', '傳輸拒絕')), timeoutP]);
    } finally {
      clearTimeout(timer);
    }

    if (result.ok) {
      await breakerRecord(true);
      const logOnly = typeof result.providerId === 'string' && result.providerId.indexOf('log:') === 0;
      if (logOnly) {
        const sr = await outbox.markSkipped(job.id, 'MODE_LOG');
        summary.skipped.push({ username: job.toUser, reason: 'MODE_LOG' });
        if (!sr || !sr.ok) summary.errors += 1;
        return;
      }
      let marked = null;
      for (let i = 0; i < 2 && !(marked && marked.ok); i++) {
        try { marked = await outbox.markSent(job.id); } catch (e) { marked = null; }
      }
      if (!marked || !marked.ok) summary.errors += 1;          // 信已寄出但沒記到：租約過期後可能重寄一次（至少一次語意）
      summary.sent += 1;
      return;
    }
    if (HEALTH_CODES.indexOf(result.code) >= 0) await breakerRecord(false, result.code);
    await failJob(job, c, summary, result.code, result.message, { permanent: result.permanent === true, retryAfterSec: result.retryAfterSec });
  }

  function auditCtx(ev, operator) {
    const type = typeOf(ev);
    const quoteNo = quoteNoOf(ev);
    let known = [];
    try { known = [ev.projectName, ev.company, ev.ownerLabel].filter((x) => typeof x === 'string'); } catch (e) { known = []; }
    return { ev, operator, type, quoteNo, known };
  }

  function metaOf(ev, kind) {
    let level = null;
    try { level = ev.step ? ev.step.level : null; } catch (e) { level = null; }
    return { level, kind: typeof kind === 'string' ? kind : null };
  }
  function actorLabelOf(ev) { try { return ev.actor && typeof ev.actor.label === 'string' ? ev.actor.label : ''; } catch (e) { return ''; } }

  /** 記一筆 skipped。回 'created'（新紀錄）、'exists'（同一事件已記過，視為重複）、'bad_key'（帳號或事件無法產生去重鍵）。 */
  async function recordSkipped(ev, username, kind, reason) {
    let key;
    try { key = dedupeKey(ev, username); } catch (e) { return 'bad_key'; }
    const r = await outbox.enqueue({
      type: ev.type, quoteId: ev.quoteId, quoteNo: ev.quoteNo, toUser: username, toMasked: '', dedupeKey: key,
      actorLabel: actorLabelOf(ev), meta: metaOf(ev, kind), status: 'skipped', skipReason: reason,
    });
    return r.created ? 'created' : 'exists';
  }

  // ═══════════════════════════════════════════════════════════════════════
  async function dispatchInner(ev, recipients, opts, summary) {
    const o = isObj(opts) ? opts : {};
    const actor = typeof o.actorUsername === 'string' ? o.actorUsername : '';
    let operator = 'system';
    try { operator = safeText(o.operatorLabel || (ev && ev.actor && ev.actor.label) || 'system', 60) || 'system'; } catch (e) { operator = 'system'; }

    const list = [];
    const kindBy = new Map();
    (Array.isArray(recipients) ? recipients : []).slice(0, MAX_RECIPIENTS).forEach((r) => {
      if (isObj(r) && typeof r.username === 'string' && r.username !== '') {
        list.push(r.username);
        if (!kindBy.has(r.username)) kindBy.set(r.username, r.kind);
      } else {
        summary.skipped.push({ username: '', reason: 'UNKNOWN_USER' });
      }
    });

    const c = auditCtx(ev, operator);
    const v = validateEvent(ev);
    if (!v.ok) {
      summary.errors += 1;
      list.forEach((u) => summary.failed.push({ username: safeText(u, 64), code: 'BAD_EVENT' }));
      await audit('QUOTE_MAIL_FAILED', operator, c.quoteNo, c.type, '', 'code=BAD_EVENT');
      return;
    }
    if (!outbox) { summary.errors += 1; return; }

    const mode = currentMode();
    const startMs = nowMs();

    // ── off：只記錄，不渲染、不查帳號 ──
    if (mode === 'off') {
      const seen = new Set();
      for (let i = 0; i < list.length; i++) {
        const u = list[i];
        if (seen.has(u)) { summary.skipped.push({ username: safeText(u, 64), reason: 'DUP' }); continue; }
        seen.add(u);
        try {
          const out = await recordSkipped(ev, u, kindBy.get(u), 'MODE_OFF');
          if (out === 'bad_key') summary.errors += 1;
        } catch (e) { summary.errors += 1; }
        summary.skipped.push({ username: safeText(u, 64), reason: 'MODE_OFF' });
      }
      return;
    }

    // ── 解析收件人 ──
    let resolved = null;
    const ur = await guarded(() => (typeof d.getUsers === 'function' ? d.getUsers() : null), auxMs());
    if (ur.ok && ur.value) {
      resolved = resolveRecipients(list, { users: ur.value, actorUsername: actor, config });
    } else {
      // 帳號資料暫時讀不到：事件先入列（email 留空），等 drainDue 時再重新解析。不丟事件。
      summary.errors += 1;
      const seen = new Set();
      resolved = { deliver: [], skipped: [] };
      list.forEach((u) => {
        if (seen.has(u)) { resolved.skipped.push({ username: safeText(u, 64), reason: 'DUP' }); return; }
        seen.add(u);
        if (actor !== '' && u === actor) { resolved.skipped.push({ username: safeText(u, 64), reason: 'ACTOR' }); return; }
        resolved.deliver.push({ username: u, email: '', label: '' });
      });
    }

    for (let i = 0; i < resolved.skipped.length; i++) {
      const s = resolved.skipped[i];
      if (s.reason === 'DUP' || s.reason === 'ACTOR') { summary.skipped.push({ username: s.username, reason: s.reason }); continue; }
      let outcome = 'error';
      try { outcome = await recordSkipped(ev, s.username, kindBy.get(s.username), s.reason); } catch (e) { summary.errors += 1; }
      if (outcome === 'exists') { summary.skipped.push({ username: s.username, reason: 'DUPLICATE' }); continue; }   // 同一事件重複觸發：不重複記、不重複寫稽核
      summary.skipped.push({ username: s.username, reason: s.reason });
      if (AUDIT_SKIP.indexOf(s.reason) >= 0) await audit('QUOTE_MAIL_SKIPPED', operator, c.quoteNo, c.type, s.username, 'reason=' + s.reason);
    }

    // ── 逐位收件人：入列 → （可以的話）當次寄送 ──
    await pool(resolved.deliver, concurrency, async (r) => {
      try {
        let key;
        try { key = dedupeKey(ev, r.username); } catch (e) {
          summary.errors += 1;
          summary.failed.push({ username: safeText(r.username, 64), code: 'BAD_KEY' });
          await audit('QUOTE_MAIL_FAILED', operator, c.quoteNo, c.type, r.username, 'code=BAD_KEY');
          return;
        }
        const kind = kindBy.get(r.username);
        const br = await breakerState();
        const hold = br.open || (nowMs() - startMs > budgetMs) || r.email === '';
        const enq = await outbox.enqueue({
          type: ev.type, quoteId: ev.quoteId, quoteNo: ev.quoteNo, toUser: r.username,
          toMasked: r.email ? maskEmail(r.email) : '', dedupeKey: key, actorLabel: actorLabelOf(ev),
          meta: metaOf(ev, kind), nextAttemptAt: br.open && br.until ? br.until : undefined,
        });
        if (!enq.created) { summary.skipped.push({ username: r.username, reason: 'DUPLICATE' }); return; }
        if (hold) { summary.queued += 1; return; }
        const job = await outbox.claim(enq.id, { leaseSec: leaseSec() });
        if (!job) { summary.queued += 1; return; }            // 已被別的工作者領走
        await sendClaimed(job, Object.assign({}, c, { kind, label: r.label, email: r.email, known: c.known.slice() }), summary);
      } catch (e) {
        summary.errors += 1;
        await audit('QUOTE_MAIL_FAILED', operator, c.quoteNo, c.type, r && r.username, 'code=INTERNAL');
      }
    });
  }

  async function dispatch(ev, recipients, opts) {
    const summary = newSummary();
    try {
      await dispatchInner(ev, recipients, opts, summary);
    } catch (e) {
      summary.errors += 1;
      await audit('QUOTE_MAIL_FAILED', 'system', quoteNoOf(ev), typeOf(ev), '', 'code=INTERNAL');
    }
    return summary;
  }

  // ═══════════════════════════════════════════════════════════════════════
  async function drainInner(opts, summary) {
    const o = isObj(opts) ? opts : {};
    if (!outbox) { summary.errors += 1; return; }
    if (currentMode() === 'off') return;
    const rebuild = typeof o.rebuild === 'function' ? o.rebuild : d.rebuild;
    if (typeof rebuild !== 'function') { summary.errors += 1; return; }       // 沒有 rebuild 就不領取，免得把工作領走卻沒辦法處理
    const br = await breakerState();
    if (br.open) { summary.breakerOpen = true; return; }

    const startMs = nowMs();
    const limit = clampInt(o.limit, 1, 50, 5);                                 // 與 outbox.claimDue 的夾法一致（非數字→5、<1→1、>50→50）
    const budget = clampInt(o.budgetMs, 1, 600000, defaultDrainBudgetMs());

    // 先清掃「租約過期且本輪次數用盡」的 sending：它們在儲存層被改成 failed/LEASE_EXPIRED，這裡把結果記進 Summary 與稽核，不再靜默
    const swept = await outbox.expireStale();
    for (let i = 0; i < swept.length; i++) {
      const rec = swept[i];
      summary.failed.push({ username: rec.toUser, code: 'LEASE_EXPIRED' });
      await audit('QUOTE_MAIL_FAILED', 'system', safeText(rec.quoteNo, 40), EVENT_TYPES.indexOf(rec.type) >= 0 ? rec.type : 'UNKNOWN', rec.toUser, 'code=LEASE_EXPIRED');
    }

    let usersP = null;
    const getUsersOnce = () => {
      if (!usersP) usersP = guarded(() => (typeof d.getUsers === 'function' ? d.getUsers() : null), auxMs());
      return usersP;
    };

    /**
     * 一筆一筆領、處理前才領：每個工作者「確認預算還有、熔斷沒開 → 領一筆 → 處理完再領下一筆」。
     * 不一次領走 limit 筆——否則傳輸卡住時，後面沒輪到的工作已被扣一次嘗試並佔著租約，函式被平台殺掉就停在 sending，
     * 租約一到又被重領再扣，退避（60/300/900 秒）形同虛設。到預算或熔斷開啟就停止領取，沒領的工作原封不動留在 pending。
     */
    let taken = 0;
    let halted = false;
    const worker = async () => {
      for (;;) {
        if (halted || taken >= limit) return;
        if (nowMs() - startMs >= budget) { halted = true; summary.budgetExhausted = true; return; }
        const b = await breakerState();                                  // 每筆之間重查熔斷（別的請求或本批前面的失敗可能已把它打開）
        if (halted || taken >= limit) return;
        if (b.open) { halted = true; summary.breakerOpen = true; return; }
        taken += 1;
        let got;
        // 領取失敗（儲存體暫時不可用）：全體停手並計入 errors。不能讓例外往外丟——Promise.all 會提早 reject，其他工作者卻還在背景繼續跑
        try { got = await outbox.claimDue({ limit: 1, leaseSec: leaseSec(), sweep: false }); } catch (e) { summary.errors += 1; halted = true; return; }
        if (!Array.isArray(got) || !got.length) { halted = true; return; }   // 沒有到期的了
        await processJob(got[0]);
      }
    };

    async function processJob(job) {
      const base = { ev: null, operator: 'system', type: EVENT_TYPES.indexOf(job.type) >= 0 ? job.type : 'UNKNOWN', quoteNo: safeText(job.quoteNo, 40), known: [] };
      try {
        const rb = await guarded(() => rebuild(job), auxMs());
        if (!rb.ok) { await failJob(job, base, summary, 'REBUILD', '重建事件失敗', { permanent: false }); return; }
        if (rb.value === null || rb.value === undefined) {
          const cr = await outbox.cancel(job.id, 'GONE');
          if (cr && cr.ok) summary.cancelled += 1;
          return;
        }
        const ev = isObj(rb.value) ? (rb.value.ev || rb.value.event) : null;
        const kind = isObj(rb.value) && typeof rb.value.kind === 'string' ? rb.value.kind : job.meta.kind;
        const c = auditCtx(ev, 'system');
        c.type = base.type;
        c.quoteNo = base.quoteNo;
        if (!validateEvent(ev).ok) { await failJob(job, c, summary, 'BAD_EVENT', '重建的事件不合法', { permanent: true }); return; }
        let key;
        try { key = dedupeKey(ev, job.toUser); } catch (e) { await failJob(job, c, summary, 'BAD_KEY', '無法產生去重鍵', { permanent: true }); return; }
        if (key !== job.dedupeKey) {                      // 單據已進到別的關卡／別次送簽：這封舊信不寄
          const cr = await outbox.cancel(job.id, 'STALE');
          if (cr && cr.ok) summary.cancelled += 1;
          return;
        }
        const ur = await getUsersOnce();
        if (!ur.ok || !ur.value) { await failJob(job, c, summary, 'USERS_UNAVAILABLE', '帳號資料暫時無法讀取', { permanent: false }); return; }
        const rs = resolveRecipients([job.toUser], { users: ur.value, actorUsername: '', config });
        if (!rs.deliver.length) {
          const reason = rs.skipped.length ? rs.skipped[0].reason : 'UNKNOWN_USER';
          await outbox.markSkipped(job.id, reason);
          summary.skipped.push({ username: job.toUser, reason });
          if (AUDIT_SKIP.indexOf(reason) >= 0) await audit('QUOTE_MAIL_SKIPPED', 'system', c.quoteNo, c.type, job.toUser, 'reason=' + reason);
          return;
        }
        const r = rs.deliver[0];
        await sendClaimed(job, Object.assign(c, { kind, label: r.label, email: r.email }), summary);
      } catch (e) {
        summary.errors += 1;
        try { await outbox.markFailed(job.id, { code: 'INTERNAL', msg: '派送內部錯誤', permanent: false }); } catch (e2) { /* 租約到期後會被重新領取 */ }
      }
    }

    await Promise.all(Array.from({ length: Math.min(concurrency, limit) }, () => worker()));
  }

  async function drainDue(opts) {
    const summary = newSummary();
    try {
      await drainInner(opts, summary);
    } catch (e) {
      summary.errors += 1;
      await audit('QUOTE_MAIL_FAILED', 'system', '', 'UNKNOWN', '', 'code=INTERNAL');
    }
    return summary;
  }

  return { dispatch, drainDue };
}

module.exports = {
  createDispatcher,
  AUDIT_SKIP,
};
