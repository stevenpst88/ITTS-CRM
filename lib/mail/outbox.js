'use strict';
/**
 * lib/mail/outbox.js — 寄件匣（outbox）：只存「引用與狀態」，不存信件內容、金額、完整 email
 *
 * 為什麼需要它：Vercel 是無伺服器環境，沒有常駐程序可以重試。先把「要寄的信」記下來，再寄、失敗再重寄，
 * 寄信出問題才不會連累簽核。內容與收件位址在「寄送當下」才由 dispatcher 重建（見 dispatcher.js）。
 *
 * 匯出與簽名（除非另註，全部回傳 Promise）：
 *   createOutbox(adapter, {now?, config?, genId?}) → outbox
 *   jsonFileAdapter / memoryAdapter / postgresAdapter（定義在 outboxAdapters.js，這裡重新匯出）
 *   scrubMessage(v, known?) → string       錯誤訊息清理（≤200 字、去密鑰樣式字串、去 email／GUID／權杖）
 *   STATUSES, MailOutboxError, DEFAULT_RETRY, DEFAULT_BREAKER, HEALTH_CODES
 *
 * outbox 方法：
 *   enqueue(job)                  → {created, id, record?, existing?}  以 dedupeKey 冪等；撞鍵一律回既有紀錄（不新建、不改狀態）
 *   claim(id, {leaseSec?})        → record | null                      領取「剛入列、已到期」的那一筆（當次請求內嘗試寄送用）
 *   claimDue({limit=5, now?, leaseSec=60, sweep?}) → record[]          領取到期工作；兩個並行領取絕不會拿到同一筆。領取前預設先 expireStale()（靜默）；sweep:false 跳過
 *   expireStale({now?})           → record[]                           租約過期且本輪次數用盡的 sending → failed/LEASE_EXPIRED，回傳被改掉的紀錄
 *                                                                      （drainDue 用它把這類狀態轉換寫進 Summary 與稽核，不再靜默）
 *   markSent(id)                  → {ok, record?, reason?}
 *   markFailed(id, {code, msg, retryAfterSec?, permanent?, attempts?}) → {ok, record?, final?, reason?}   attempts＝領取當時的 attempts（樂觀鎖）
 *   markSkipped(id, reason)       → {ok, record?, reason?}             pending/sending → skipped
 *   cancel(id, reason)            → {ok, record?, reason?}             pending/sending/failed → cancelled
 *   requeue(id)                   → {ok, record?, reason?}             admin 重送：failed/cancelled/skipped → pending（再給一輪）
 *   get(id) / list(filter) / purge(olderThanDays) / stats()
 *   breaker.state() → {open, until, failures, halfOpen}；breaker.record(ok, code?) → state
 *
 * record（camelCase；時間都是 ISO 字串或 null）：
 *   id, type, quoteId, quoteNo, toUser, toMasked, dedupeKey, status('pending'|'sending'|'sent'|'failed'|'skipped'|'cancelled'),
 *   attempts（累計寄送次數）, attemptsInRound（本輪次數；requeue 歸零）, requeues, nextAttemptAt, leaseUntil,
 *   lastErrorCode, lastErrorMsg(≤200), skipReason（skipped 與 cancelled 的原因）, createdAt, updatedAt, sentAt, actorLabel, meta:{level, kind}
 *   （attemptsInRound、requeues 是規格欄位之外的補充，用來實作「requeue：attempts 不歸零但允許再一輪」。）
 *
 * 規則：
 *  - enqueue 只從白名單欄位組紀錄；傳入的 html／text／subject／金額／email 等多餘欄位一律丟棄（有測試）。toMasked 若傳入完整位址會被改成遮罩。
 *  - 重試：第 1 次嘗試失敗後，依 config.retry.delaysSec（預設 60/300/900 秒）排下一次；Graph 429 的 retryAfterSec 優先（上限 24 小時）。
 *    【規格解讀，請審】config.retry.maxAttempts（預設 3）解讀為「首次嘗試之後的重試次數」，所以一輪最多嘗試 1+3＝4 次，三個延遲都會用到
 *    （規格註解「3 次用盡標記 failed」與 plan §9「重試 3 次（1/5/15 分鐘）」）。若原意是「總共 3 次」，改下面的 roundLimit 一行即可。
 *  - permanent 錯誤、或本輪次數用盡 → failed。requeue 讓 attemptsInRound 歸零再給一輪（attempts 累計不歸零）。
 *  - 領取會把 attempts／attemptsInRound 加一（領取＝開始一次嘗試）。工作者當掉（租約過期）會被重新領取；
 *    若租約過期時本輪次數已用盡，改標 failed/LEASE_EXPIRED，避免無限重寄；這個轉換由 expireStale() 回報，drainDue 會把它記進 Summary.failed
 *    並寫 QUOTE_MAIL_FAILED（code=LEASE_EXPIRED）。語意是「至少一次」：租約過期代表上一次是否已寄出不確定，極端 race 下可能重複寄出一封；不會靜默丟失。
 *  - finishAttempt 有樂觀鎖（attempts 必須相同）：租約被別人搶走後，舊工作者的回報會被忽略。
 *  - 熔斷（breaker）：只有「傳輸健康度」類錯誤計數（TIMEOUT／NETWORK／SERVER／THROTTLED／AUTH）；REJECTED／BAD_MESSAGE／NOT_CONFIGURED
 *    是單封信或設定的問題，不算。AUTH 直接開；THROTTLED 權重 3；其餘權重 1；視窗內累計達 failures（預設 5）就開 openSec（預設 600 秒）。
 *    任何成功都完全重置。開啟期間結束後進入「半開」：下一次失敗立刻重新開啟，成功則重置。狀態存在 adapter（Postgres 為 mail_breaker），
 *    多實例以最後寫入為準（近似值即可，不求精確）。
 */

const crypto = require('crypto');
const { EVENT_TYPES, STEP_LEVELS, isValidQuoteId } = require('./events');
const { KINDS } = require('./visibility');
const { normalizeEmail, maskEmail, safeText } = require('./safety');
const adapters = require('./outboxAdapters');
const { scrubMessage } = require('./scrub');

const { STATUSES, TERMINAL_STATUSES, MailOutboxError, msOf, isoOf } = adapters;

const DEFAULT_RETRY = Object.freeze({ delaysSec: Object.freeze([60, 300, 900]), maxAttempts: 3 });
const DEFAULT_BREAKER = Object.freeze({ failures: 5, windowSec: 600, openSec: 600 });
const DEFAULT_RETENTION_DAYS = 90;
const MAX_RETRY_AFTER_SEC = 24 * 3600;
const HEALTH_CODES = Object.freeze(['TIMEOUT', 'NETWORK', 'SERVER', 'THROTTLED', 'AUTH']);
const BREAKER_WEIGHT = Object.freeze({ THROTTLED: 3 });

const RE_CTRL = /[\x00-\x1f\x7f]/;
const RE_REASON = /^[A-Z][A-Z0-9_]{0,39}$/;
const RE_MASKED = /^[^@\s*]\*\*\*@[a-z0-9.-]+$/;

function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }

function cleanCode(c, dflt) {
  return (typeof c === 'string' && RE_REASON.test(c)) ? c : dflt;
}

/** toMasked 只接受遮罩格式；若傳入完整位址就改成遮罩；其他一律 ''。 */
function sanitizeMasked(v) {
  if (typeof v !== 'string' || v.length > 300) return '';
  if (RE_MASKED.test(v)) return v;
  const n = normalizeEmail(v);
  return n.ok ? maskEmail(n.value) : '';
}

function sanitizeRetry(cfg) {
  const r = (cfg && isObj(cfg.retry)) ? cfg.retry : DEFAULT_RETRY;
  let delays = Array.isArray(r.delaysSec) ? r.delaysSec.filter((x) => typeof x === 'number' && isFinite(x) && x >= 1 && x <= MAX_RETRY_AFTER_SEC).map((x) => Math.floor(x)) : [];
  if (!delays.length) delays = DEFAULT_RETRY.delaysSec.slice();
  const retries = (Number.isSafeInteger(r.maxAttempts) && r.maxAttempts >= 0 && r.maxAttempts <= 20) ? r.maxAttempts : DEFAULT_RETRY.maxAttempts;
  return { delays, retries };
}

function sanitizeBreakerCfg(cfg) {
  const b = (cfg && isObj(cfg.breaker)) ? cfg.breaker : DEFAULT_BREAKER;
  const pos = (x, d) => (typeof x === 'number' && isFinite(x) && x >= 1 ? Math.floor(x) : d);
  return { failures: pos(b.failures, DEFAULT_BREAKER.failures), windowSec: pos(b.windowSec, DEFAULT_BREAKER.windowSec), openSec: pos(b.openSec, DEFAULT_BREAKER.openSec) };
}

function newId() { return 'mo_' + crypto.randomBytes(10).toString('hex'); }

/**
 * @param {object} adapter  jsonFileAdapter／memoryAdapter／postgresAdapter 的回傳值
 * @param {{now?: function(): (number|Date|string), config?: object, genId?: function(): string}} deps
 *   now      時鐘（可回毫秒、Date 或 ISO 字串）；預設 Date.now。測試用假時鐘。
 *   config   getMailConfig() 的結果；只讀 retry、breaker、retentionDays
 */
function createOutbox(adapter, deps) {
  if (!adapter || typeof adapter.insert !== 'function' || typeof adapter.claimDue !== 'function' || typeof adapter.expireExhausted !== 'function') {
    throw new MailOutboxError('BAD_ADAPTER', 'createOutbox 需要 outbox adapter');
  }
  const d = isObj(deps) ? deps : {};
  const nowFn = typeof d.now === 'function' ? d.now : Date.now;
  const genId = typeof d.genId === 'function' ? d.genId : newId;
  const cfg = isObj(d.config) ? d.config : {};
  const retry = sanitizeRetry(cfg);
  const bcfg = sanitizeBreakerCfg(cfg);
  const retentionDays = (typeof cfg.retentionDays === 'number' && cfg.retentionDays >= 0) ? cfg.retentionDays : DEFAULT_RETENTION_DAYS;
  const roundLimit = 1 + retry.retries;            // 一輪最多嘗試幾次（首次 + 重試）。見檔頭「規格解讀」

  function nowMs() {
    const v = msOf(nowFn());
    return Number.isFinite(v) ? v : Date.now();
  }
  const iso = (ms) => new Date(ms).toISOString();

  // ── 入列 ───────────────────────────────────────────────────────────────
  function buildRecord(job) {
    if (!isObj(job)) throw new MailOutboxError('BAD_JOB', 'job 必須是物件');
    if (EVENT_TYPES.indexOf(job.type) < 0) throw new MailOutboxError('BAD_JOB', 'job.type 不在事件列舉內');
    if (!isValidQuoteId(job.quoteId)) throw new MailOutboxError('BAD_JOB', 'job.quoteId 格式不合法');
    const quoteNo = typeof job.quoteNo === 'string' ? safeText(job.quoteNo, 40) : '';
    if (!quoteNo) throw new MailOutboxError('BAD_JOB', 'job.quoteNo 必填');
    if (typeof job.toUser !== 'string' || job.toUser === '' || job.toUser.length > 200 || RE_CTRL.test(job.toUser)) {
      throw new MailOutboxError('BAD_JOB', 'job.toUser 不合法');
    }
    if (typeof job.dedupeKey !== 'string' || job.dedupeKey === '' || job.dedupeKey.length > 700 || RE_CTRL.test(job.dedupeKey)) {
      throw new MailOutboxError('BAD_JOB', 'job.dedupeKey 不合法');
    }
    const t = nowMs();
    const skipped = job.status === 'skipped';
    let nextMs = job.nextAttemptAt === undefined || job.nextAttemptAt === null ? t : msOf(job.nextAttemptAt);
    if (!Number.isFinite(nextMs)) nextMs = t;
    const m = isObj(job.meta) ? job.meta : {};
    return {
      id: genId(),
      type: job.type,
      quoteId: job.quoteId,
      quoteNo,
      toUser: job.toUser,
      toMasked: sanitizeMasked(job.toMasked),
      dedupeKey: job.dedupeKey,
      status: skipped ? 'skipped' : 'pending',
      attempts: 0,
      attemptsInRound: 0,
      requeues: 0,
      nextAttemptAt: skipped ? null : iso(nextMs),
      leaseUntil: null,
      lastErrorCode: null,
      lastErrorMsg: null,
      skipReason: skipped ? cleanCode(job.skipReason, 'UNSPECIFIED') : null,
      createdAt: iso(t),
      updatedAt: iso(t),
      sentAt: null,
      actorLabel: typeof job.actorLabel === 'string' ? safeText(job.actorLabel, 100) : '',
      meta: {
        level: STEP_LEVELS.indexOf(m.level) >= 0 ? m.level : null,
        kind: typeof m.kind === 'string' && KINDS.indexOf(m.kind) >= 0 ? m.kind : null,
      },
    };
  }

  async function enqueue(job) {
    const rec = buildRecord(job);
    const r = await adapter.insert(rec);
    if (r.inserted) return { created: true, id: r.record.id, record: r.record };
    return { created: false, id: r.record.id, existing: r.record };
  }

  // ── 領取 ───────────────────────────────────────────────────────────────
  function leaseOf(t, leaseSec) {
    const s = (typeof leaseSec === 'number' && isFinite(leaseSec) && leaseSec >= 1) ? Math.min(Math.floor(leaseSec), 3600) : 60;
    return iso(t + s * 1000);
  }

  async function claim(id, opts) {
    if (typeof id !== 'string' || id === '') return null;
    const t = nowMs();
    return adapter.claimById({ id, now: iso(t), leaseUntil: leaseOf(t, opts && opts.leaseSec) });
  }

  /**
   * 清掃「租約過期且本輪次數已用盡」的 sending → failed/LEASE_EXPIRED，回傳被改掉的紀錄（陣列，可能為空）。
   * 這個轉換原本藏在 claimDue 裡、靜默發生；dispatcher.drainDue 現在先呼叫本函式，把結果記進 Summary 與稽核（QUOTE_MAIL_FAILED）。
   */
  async function expireStale(opts) {
    const o = isObj(opts) ? opts : {};
    const t = o.now === undefined ? nowMs() : (Number.isFinite(msOf(o.now)) ? msOf(o.now) : nowMs());
    return adapter.expireExhausted({ now: iso(t), maxRoundAttempts: roundLimit });
  }

  async function claimDue(opts) {
    const o = isObj(opts) ? opts : {};
    const t = o.now === undefined ? nowMs() : (Number.isFinite(msOf(o.now)) ? msOf(o.now) : nowMs());
    const limit = (typeof o.limit === 'number' && isFinite(o.limit)) ? Math.max(1, Math.min(50, Math.floor(o.limit))) : 5;
    // 預設領取前先清掃（結果丟棄＝靜默，維持既有行為）。要看到被清掃的紀錄：自己先呼叫 expireStale()，再以 sweep:false 領取
    if (o.sweep !== false) await adapter.expireExhausted({ now: iso(t), maxRoundAttempts: roundLimit });
    return adapter.claimDue({ now: iso(t), leaseUntil: leaseOf(t, o.leaseSec), limit, maxRoundAttempts: roundLimit });
  }

  // ── 結果回報 ───────────────────────────────────────────────────────────
  async function markSent(id) {
    const rec = await adapter.get(id);
    if (!rec) return { ok: false, reason: 'NOT_FOUND' };
    if (rec.status === 'sent') return { ok: true, record: rec, already: true };
    const upd = await adapter.markSent({ id, now: iso(nowMs()) });
    if (upd) return { ok: true, record: upd };
    const cur = await adapter.get(id);
    return { ok: false, reason: 'BAD_STATE', status: cur ? cur.status : null };
  }

  async function markFailed(id, info) {
    const o = isObj(info) ? info : {};
    const rec = await adapter.get(id);
    if (!rec) return { ok: false, reason: 'NOT_FOUND' };
    if (rec.status !== 'sending') return { ok: false, reason: 'NOT_SENDING', status: rec.status };
    // o.attempts＝呼叫端領取當時的 attempts。租約過期被別的工作者接手後 attempts 會變大，舊工作者的回報就在這裡被擋掉
    if (Number.isSafeInteger(o.attempts) && o.attempts !== rec.attempts) return { ok: false, reason: 'LOST_LEASE' };
    const t = nowMs();
    const code = cleanCode(o.code, 'UNKNOWN');
    const msg = scrubMessage(o.msg);
    const final = o.permanent === true || rec.attemptsInRound >= roundLimit;
    let status = 'failed';
    let next = null;
    if (!final) {
      status = 'pending';
      let delay = retry.delays[Math.min(Math.max(rec.attemptsInRound, 1) - 1, retry.delays.length - 1)];
      const ra = o.retryAfterSec;
      if (typeof ra === 'number' && isFinite(ra) && ra > 0) delay = Math.min(Math.max(ra, 1), MAX_RETRY_AFTER_SEC);   // Retry-After 優先
      next = iso(t + Math.ceil(delay) * 1000);
    }
    const upd = await adapter.finishAttempt({ id, expectAttempts: Number.isSafeInteger(o.attempts) ? o.attempts : rec.attempts, now: iso(t), status, nextAttemptAt: next, code, msg });
    if (!upd) return { ok: false, reason: 'LOST_LEASE' };          // 租約已被別人接手，這次回報作廢
    return { ok: true, record: upd, final: status === 'failed' };
  }

  async function terminal(id, status, reason, from) {
    const rec = await adapter.get(id);
    if (!rec) return { ok: false, reason: 'NOT_FOUND' };
    if (from.indexOf(rec.status) < 0) return { ok: false, reason: 'BAD_STATE', status: rec.status };
    const upd = await adapter.setTerminal({ id, now: iso(nowMs()), status, reason: cleanCode(reason, 'UNSPECIFIED'), from });
    if (!upd) {
      const cur = await adapter.get(id);
      return { ok: false, reason: 'BAD_STATE', status: cur ? cur.status : null };
    }
    return { ok: true, record: upd };
  }
  const markSkipped = (id, reason) => terminal(id, 'skipped', reason, ['pending', 'sending']);
  const cancel = (id, reason) => terminal(id, 'cancelled', reason, ['pending', 'sending', 'failed']);

  async function requeue(id) {
    const rec = await adapter.get(id);
    if (!rec) return { ok: false, reason: 'NOT_FOUND' };
    const upd = await adapter.requeue({ id, now: iso(nowMs()) });
    if (!upd) return { ok: false, reason: 'BAD_STATE', status: rec.status };
    return { ok: true, record: upd };
  }

  // ── 查詢與維護 ─────────────────────────────────────────────────────────
  const get = (id) => adapter.get(id);

  function listFilter(f) {
    const o = isObj(f) ? f : {};
    const s = (v, max) => (typeof v === 'string' && v !== '' && v.length <= max ? v : null);
    const status = STATUSES.indexOf(o.status) >= 0 ? o.status : null;
    const sinceIso = o.since === undefined || o.since === null ? null : isoOf(o.since);
    const limit = (typeof o.limit === 'number' && isFinite(o.limit)) ? Math.max(1, Math.min(500, Math.floor(o.limit))) : 50;
    const offset = (typeof o.offset === 'number' && isFinite(o.offset) && o.offset > 0) ? Math.min(Math.floor(o.offset), 1000000) : 0;
    return { status, type: s(o.type, 40), quoteNo: s(o.quoteNo, 40), toUser: s(o.toUser, 200), since: sinceIso, limit, offset };
  }
  async function list(f) {
    const filt = listFilter(f);
    const r = await adapter.list(filt);
    return { rows: r.rows, total: r.total, limit: filt.limit, offset: filt.offset };
  }

  async function purge(olderThanDays) {
    const days = (typeof olderThanDays === 'number' && isFinite(olderThanDays) && olderThanDays >= 0) ? olderThanDays : retentionDays;
    return adapter.purge({ cutoff: iso(nowMs() - days * 86400000) });
  }

  const stats = () => adapter.stats({ now: iso(nowMs()) });

  // ── 熔斷 ───────────────────────────────────────────────────────────────
  function viewOf(st, t) {
    const until = st && Number.isFinite(st.openUntil) && st.openUntil > 0 ? st.openUntil : 0;
    const open = until > t;
    const failures = st && Array.isArray(st.failures) ? st.failures.reduce((a, f) => a + (Number(f.w) || 0), 0) : 0;
    return { open, until: until ? iso(until) : null, failures, halfOpen: until > 0 && !open };
  }
  async function loadBreaker() {
    const raw = await adapter.breakerLoad();
    const st = isObj(raw) ? raw : {};
    return {
      failures: Array.isArray(st.failures) ? st.failures.filter((f) => isObj(f) && Number.isFinite(f.t) && Number.isFinite(f.w)) : [],
      openUntil: Number.isFinite(st.openUntil) ? st.openUntil : 0,
    };
  }
  const breaker = {
    async state() { return viewOf(await loadBreaker(), nowMs()); },
    async record(ok, code) {
      const t = nowMs();
      let st = await loadBreaker();
      if (ok) {
        if (st.failures.length === 0 && st.openUntil === 0) return viewOf(st, t);
        st = { failures: [], openUntil: 0 };
      } else {
        const c = typeof code === 'string' ? code : '';
        if (HEALTH_CODES.indexOf(c) < 0) return viewOf(st, t);                    // 與傳輸健康無關的錯誤不計入
        if (st.openUntil > t) return viewOf(st, t);                                // 已開啟：不延長
        const halfOpen = st.openUntil > 0;
        st.failures = st.failures.filter((f) => f.t > t - bcfg.windowSec * 1000);
        const w = c === 'AUTH' ? bcfg.failures : (BREAKER_WEIGHT[c] || 1);
        st.failures.push({ t, w });
        const total = st.failures.reduce((a, f) => a + f.w, 0);
        if (halfOpen || total >= bcfg.failures) {
          st.openUntil = t + bcfg.openSec * 1000;
          if (halfOpen) st.failures = [{ t, w: bcfg.failures }];
        }
      }
      await adapter.breakerSave(st);
      return viewOf(st, t);
    },
  };

  return {
    adapter,
    roundLimit,
    enqueue, claim, claimDue, expireStale, markSent, markFailed, markSkipped, cancel, requeue,
    get, list, purge, stats, breaker,
  };
}

module.exports = {
  createOutbox,
  jsonFileAdapter: adapters.jsonFileAdapter,
  memoryAdapter: adapters.memoryAdapter,
  postgresAdapter: adapters.postgresAdapter,
  scrubMessage,
  sanitizeMasked,
  STATUSES,
  TERMINAL_STATUSES,
  MailOutboxError,
  DEFAULT_RETRY,
  DEFAULT_BREAKER,
  HEALTH_CODES,
};
