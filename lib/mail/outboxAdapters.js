'use strict';
/**
 * lib/mail/outboxAdapters.js — outbox 的儲存後端（由 lib/mail/outbox.js 重新匯出；呼叫端一律 require('./outbox')）
 *
 * 分工：策略（退避、重試用盡、熔斷、欄位白名單）全部在 outbox.js；本檔只負責「保證原子的狀態轉換」。
 * 所有 adapter 都回傳 Promise，且對同一份儲存體的每一個操作都是原子的：
 *   - JSON／記憶體：每個操作的本體是「同步」的讀—改—寫（中間沒有 await），再加一條 Promise 鏈當互斥鎖，
 *     所以同一個 Node 程序內 50 個並行 claim 也不會領到同一筆。
 *   - Postgres：每個操作都是「單一條 SQL」（UPDATE … WHERE … RETURNING；領取用 FOR UPDATE SKIP LOCKED），
 *     由資料庫保證原子；不使用多語句交易，因此 Supabase 交易模式連線池（pgbouncer）也能用。
 *
 * 匯出：
 *   STATUSES, TERMINAL_STATUSES, MailOutboxError
 *   jsonFileAdapter({file, rootDir?, fs?, allowOnVercel?})   本機單一程序用。整檔讀寫；寫入＝暫存檔＋rename 原子替換；
 *                                               檔案損壞→備份成 <file>.corrupt-<ts>.bak 並從空開始（讀取絕不 throw）；
 *                                               只壞幾筆→保留合格的、複製一份備份。同一個損壞狀態（內容雜湊相同）只備份一次，
 *                                               且每個檔案最多留 MAX_CORRUPT_BACKUPS 份（超過刪最舊的），唯讀操作不會讓備份無限增長
 *   memoryAdapter({allowOnVercel?})              測試／預覽用（資料只在記憶體，重啟即消失）
 *   （jsonFileAdapter 與 memoryAdapter 在 Vercel 上（環境變數 VERCEL 有值）建立時會直接丟 NOT_FOR_VERCEL：沒有可寫的持久檔案系統，
 *    誤接線會變成「寄信紀錄默默掉光」；正式環境一律用 postgresAdapter。測試用 allowOnVercel:true 才能繞過）
 *   postgresAdapter({query})                     query(sql, params) → Promise<{rows}>；表 mail_outbox、mail_breaker
 *   postgresAdapter.SQL / postgresAdapter.DDL    全部 SQL 都是下面的常數（執行期 query 前會檢查「sql 必須是這些常數之一」）
 *
 * Adapter 介面（outbox.js 只依賴這些；所有時間參數都是 ISO 字串）：
 *   init()                                                    → void
 *   insert(rec)                                               → {inserted, record}     以 dedupeKey 唯一；撞鍵回既有紀錄
 *   get(id) / getByKey(dedupeKey)                             → record | null
 *   claimById({id, now, leaseUntil})                          → record | null          pending 且到期 → sending
 *   claimDue({now, leaseUntil, limit, maxRoundAttempts})      → record[]               pending 到期，或 sending 租約過期且本輪次數未用盡
 *                                                                                      （只領取，不清掃；租約過期且次數已用盡者不會被領取）
 *   expireExhausted({now, maxRoundAttempts})                  → record[]               sending、租約已過期、本輪次數已用盡 → 改標 failed/LEASE_EXPIRED，
 *                                                                                      並「回傳被改掉的紀錄」，讓呼叫端（dispatcher.drainDue）能記進 Summary 與稽核，
 *                                                                                      不再是靜默轉換。outbox.claimDue 領取前一定先呼叫它
 *   markSent({id, now})                                       → record | null          pending/sending/failed/cancelled → sent
 *   finishAttempt({id, expectAttempts, now, status, nextAttemptAt, code, msg}) → record | null
 *                                                              只有「仍是 sending 且 attempts 相同」才會套用（樂觀鎖：租約被別人搶走後，舊工作者的回報無效）
 *   setTerminal({id, now, status, reason, from})              → record | null         status 為 skipped/cancelled；from 為允許的現況狀態
 *   requeue({id, now})                                        → record | null         failed/cancelled/skipped → pending
 *   list(filter) / purge({cutoff}) / stats({now})
 *   breakerLoad() / breakerSave(state)
 *
 * 紀錄（record）只有引用與狀態：不含信件內容、金額、完整 email。欄位見 outbox.js。
 *
 * 關於 Postgres adapter（重要）：本機沒有 Postgres，這個 adapter 沒有對真實資料庫執行過。
 * 已做的是：SQL 靜態審查（見各 SQL 上方的註解）、用假 query 函式驗證「參數順序／參數化／領取語意」、
 * 以及用記憶體模擬器跑與 JSON adapter 相同的行為測試。沒做的是 SQL 語法／型別／鎖行為在真 Postgres 上的驗證，
 * 請先在 Demo 環境（獨立 Supabase）依 mail_delivery_report.md 的步驟驗證，再上正式環境。
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const STATUSES = Object.freeze(['pending', 'sending', 'sent', 'failed', 'skipped', 'cancelled']);
const TERMINAL_STATUSES = Object.freeze(['sent', 'failed', 'skipped', 'cancelled']);
const MAX_JOBS_PER_FILE = 200000;       // JSON 檔案的保險上限（超過代表 purge 沒在跑；拒絕新增而不是讓檔案無限長大）
const MAX_CORRUPT_BACKUPS = 3;          // JSON 檔案損壞備份（<file>.corrupt-*.bak）最多保留幾份；超過刪最舊的
const MAX_RECOVERY_NOTES = 50;          // diagnostics().recoveries 最多留幾筆（保留最新的）
const MAX_SEEN_CORRUPT = 32;            // 記住「已備份過的損壞內容雜湊」的上限（超過忘掉最舊的）
const LEASE_EXPIRED_MSG = '寄送工作者逾時未回報，且已用盡本輪嘗試次數';

class MailOutboxError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'MailOutboxError';
    this.code = code;
  }
}

// ── 小工具 ──────────────────────────────────────────────────────────────────
function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
function strOr(v, dflt) { return typeof v === 'string' ? v : dflt; }
function intOr(v, dflt) { return Number.isSafeInteger(v) && v >= 0 ? v : dflt; }

/** 任意時間表示（ISO 字串／毫秒／Date）→ 毫秒；無法解析回 NaN。 */
function msOf(v) {
  if (v === null || v === undefined || v === '') return NaN;
  if (v instanceof Date) return v.getTime();
  if (typeof v === 'number') return v;
  if (typeof v === 'string') return Date.parse(v);
  return NaN;
}
/** 任意時間表示 → ISO 字串；無法解析回 null。 */
function isoOf(v) {
  const ms = msOf(v);
  if (!Number.isFinite(ms)) return null;
  try { return new Date(ms).toISOString(); } catch (e) { return null; }
}

/**
 * 把「從儲存體讀出來的任何東西」整理成標準 record；不合格（缺 id／dedupeKey、狀態不認得）回 null。
 * 順便維持兩個不變條件（Postgres 版以 CHECK 約束保證）：pending 一定有 nextAttemptAt；sending 一定有 leaseUntil。
 */
function normalizeRecord(r) {
  if (!isObj(r)) return null;
  if (typeof r.id !== 'string' || r.id === '') return null;
  if (typeof r.dedupeKey !== 'string' || r.dedupeKey === '') return null;
  if (STATUSES.indexOf(r.status) < 0) return null;
  const createdAt = isoOf(r.createdAt) || new Date(0).toISOString();
  const updatedAt = isoOf(r.updatedAt) || createdAt;
  const m = isObj(r.meta) ? r.meta : {};
  const rec = {
    id: r.id,
    type: strOr(r.type, ''),
    quoteId: strOr(r.quoteId, ''),
    quoteNo: strOr(r.quoteNo, ''),
    toUser: strOr(r.toUser, ''),
    toMasked: strOr(r.toMasked, ''),
    dedupeKey: r.dedupeKey,
    status: r.status,
    attempts: intOr(r.attempts, 0),
    attemptsInRound: intOr(r.attemptsInRound, 0),
    requeues: intOr(r.requeues, 0),
    nextAttemptAt: isoOf(r.nextAttemptAt),
    leaseUntil: isoOf(r.leaseUntil),
    lastErrorCode: typeof r.lastErrorCode === 'string' ? r.lastErrorCode : null,
    lastErrorMsg: typeof r.lastErrorMsg === 'string' ? r.lastErrorMsg : null,
    skipReason: typeof r.skipReason === 'string' ? r.skipReason : null,
    createdAt,
    updatedAt,
    sentAt: isoOf(r.sentAt),
    actorLabel: strOr(r.actorLabel, ''),
    meta: {
      level: (typeof m.level === 'number' || typeof m.level === 'string') ? m.level : null,
      kind: typeof m.kind === 'string' ? m.kind : null,
    },
  };
  if (rec.status === 'pending' && !rec.nextAttemptAt) rec.nextAttemptAt = createdAt;
  if (rec.status === 'sending' && !rec.leaseUntil) rec.leaseUntil = updatedAt;
  return rec;
}

function rowToRecord(row) {
  if (!isObj(row)) return null;
  let meta = row.meta;
  if (typeof meta === 'string') { try { meta = JSON.parse(meta); } catch (e) { meta = {}; } }
  return normalizeRecord({
    id: row.id, type: row.type, quoteId: row.quote_id, quoteNo: row.quote_no, toUser: row.to_user,
    toMasked: row.to_masked, dedupeKey: row.dedupe_key, status: row.status,
    attempts: row.attempts, attemptsInRound: row.attempts_in_round, requeues: row.requeues,
    nextAttemptAt: row.next_attempt_at, leaseUntil: row.lease_until,
    lastErrorCode: row.last_error_code, lastErrorMsg: row.last_error_msg, skipReason: row.skip_reason,
    createdAt: row.created_at, updatedAt: row.updated_at, sentAt: row.sent_at, actorLabel: row.actor_label, meta,
  });
}

function emptyCounts() {
  const c = {};
  STATUSES.forEach((s) => { c[s] = 0; });
  return c;
}

function cmpStr(a, b) { return a < b ? -1 : (a > b ? 1 : 0); }

// ═════════════════════════════════════════════════════════════════════════
// 共用的「狀態儲存」實作：JSON 檔案與記憶體只差在 io.load()／io.save()
// ═════════════════════════════════════════════════════════════════════════
/**
 * @param {{load: function():{jobs:object[], breaker:(object|null)}, save: function(object):void}} io 兩個函式都必須是同步的
 * @param {string} kind 'json' | 'memory'
 */
function createStateAdapter(io, kind) {
  // 互斥鎖：每個操作排隊執行。操作本體本來就是同步的，這條鏈是第二層保險（日後有人在操作中加 await 也不會出事）。
  let chain = Promise.resolve();
  function run(fn) {
    const p = chain.then(() => fn());
    chain = p.then(() => undefined, () => undefined);
    return p;
  }
  function find(st, id) {
    for (let i = 0; i < st.jobs.length; i++) if (st.jobs[i].id === id) return st.jobs[i];
    return null;
  }
  function isDuePending(j, nowMs) {
    if (j.status !== 'pending') return false;
    const t = msOf(j.nextAttemptAt);
    return !Number.isFinite(t) || t <= nowMs;
  }
  function isExpiredSending(j, nowMs) {
    if (j.status !== 'sending') return false;
    const t = msOf(j.leaseUntil);
    return !Number.isFinite(t) || t <= nowMs;
  }

  return {
    kind,
    init: () => run(() => { io.save(io.load()); }),

    insert: (rec) => run(() => {
      const st = io.load();
      for (let i = 0; i < st.jobs.length; i++) {
        if (st.jobs[i].dedupeKey === rec.dedupeKey) return { inserted: false, record: st.jobs[i] };
        if (st.jobs[i].id === rec.id) throw new MailOutboxError('ID_CONFLICT', 'outbox id 重複');
      }
      if (st.jobs.length >= MAX_JOBS_PER_FILE) throw new MailOutboxError('STORE_FULL', 'outbox 檔案筆數已達上限，請先 purge');
      const stored = normalizeRecord(rec);
      if (!stored) throw new MailOutboxError('BAD_RECORD', 'outbox 紀錄格式不合法');
      st.jobs.push(stored);
      io.save(st);
      return { inserted: true, record: stored };
    }),

    get: (id) => run(() => find(io.load(), id)),
    getByKey: (key) => run(() => {
      const st = io.load();
      for (let i = 0; i < st.jobs.length; i++) if (st.jobs[i].dedupeKey === key) return st.jobs[i];
      return null;
    }),

    claimById: ({ id, now, leaseUntil }) => run(() => {
      const st = io.load();
      const j = find(st, id);
      const nowMs = msOf(now);
      if (!j || !isDuePending(j, nowMs)) return null;
      j.status = 'sending';
      j.leaseUntil = leaseUntil;
      j.attempts += 1;
      j.attemptsInRound += 1;
      j.updatedAt = now;
      io.save(st);
      return j;
    }),

    // 租約過期、本輪次數已用盡的 sending：不再領取，直接標 failed（避免當掉的工作者造成無限重寄）。
    // 回傳被改掉的紀錄：這個轉換發生在儲存層，若不回報，派送器的 Summary／稽核就看不到（曾經是靜默的）。
    expireExhausted: ({ now, maxRoundAttempts }) => run(() => {
      const st = io.load();
      const nowMs = msOf(now);
      const flipped = [];
      st.jobs.forEach((j) => {
        if (isExpiredSending(j, nowMs) && j.attemptsInRound >= maxRoundAttempts) {
          j.status = 'failed';
          j.lastErrorCode = 'LEASE_EXPIRED';
          j.lastErrorMsg = LEASE_EXPIRED_MSG;
          j.leaseUntil = null;
          j.updatedAt = now;
          flipped.push(j);
        }
      });
      if (flipped.length) io.save(st);
      return flipped;
    }),

    claimDue: ({ now, leaseUntil, limit, maxRoundAttempts }) => run(() => {
      const st = io.load();
      const nowMs = msOf(now);
      // 領取到期的（租約過期且次數已用盡者不領；它們由 expireExhausted 處理）
      const due = st.jobs.filter((j) => isDuePending(j, nowMs) || (isExpiredSending(j, nowMs) && j.attemptsInRound < maxRoundAttempts));
      due.sort((a, b) => cmpStr(a.nextAttemptAt || a.leaseUntil || '', b.nextAttemptAt || b.leaseUntil || '') || cmpStr(a.createdAt, b.createdAt) || cmpStr(a.id, b.id));
      const picked = due.slice(0, limit);
      picked.forEach((j) => {
        j.status = 'sending';
        j.leaseUntil = leaseUntil;
        j.attempts += 1;
        j.attemptsInRound += 1;
        j.updatedAt = now;
      });
      if (picked.length) io.save(st);
      return picked;
    }),

    markSent: ({ id, now }) => run(() => {
      const st = io.load();
      const j = find(st, id);
      if (!j || ['pending', 'sending', 'failed', 'cancelled'].indexOf(j.status) < 0) return null;
      j.status = 'sent';
      j.sentAt = now;
      j.leaseUntil = null;
      j.nextAttemptAt = null;
      j.updatedAt = now;
      io.save(st);
      return j;
    }),

    finishAttempt: ({ id, expectAttempts, now, status, nextAttemptAt, code, msg }) => run(() => {
      const st = io.load();
      const j = find(st, id);
      if (!j || j.status !== 'sending' || j.attempts !== expectAttempts) return null;
      j.status = status;
      j.nextAttemptAt = nextAttemptAt;
      j.leaseUntil = null;
      j.lastErrorCode = code;
      j.lastErrorMsg = msg;
      j.updatedAt = now;
      io.save(st);
      return j;
    }),

    setTerminal: ({ id, now, status, reason, from }) => run(() => {
      const st = io.load();
      const j = find(st, id);
      if (!j || from.indexOf(j.status) < 0) return null;
      j.status = status;
      j.skipReason = reason;
      j.leaseUntil = null;
      j.nextAttemptAt = null;
      j.updatedAt = now;
      io.save(st);
      return j;
    }),

    requeue: ({ id, now }) => run(() => {
      const st = io.load();
      const j = find(st, id);
      if (!j || ['failed', 'cancelled', 'skipped'].indexOf(j.status) < 0) return null;
      j.status = 'pending';
      j.nextAttemptAt = now;
      j.leaseUntil = null;
      j.attemptsInRound = 0;
      j.requeues += 1;
      j.skipReason = null;
      j.updatedAt = now;
      io.save(st);
      return j;
    }),

    list: (f) => run(() => {
      const st = io.load();
      const sinceMs = f.since ? msOf(f.since) : NaN;
      const rows = st.jobs.filter((j) => (!f.status || j.status === f.status)
        && (!f.type || j.type === f.type)
        && (!f.quoteNo || j.quoteNo === f.quoteNo)
        && (!f.toUser || j.toUser === f.toUser)
        && (!Number.isFinite(sinceMs) || msOf(j.createdAt) >= sinceMs));
      rows.sort((a, b) => cmpStr(b.createdAt, a.createdAt) || cmpStr(b.id, a.id));
      return { rows: rows.slice(f.offset, f.offset + f.limit), total: rows.length };
    }),

    purge: ({ cutoff }) => run(() => {
      const st = io.load();
      const cutMs = msOf(cutoff);
      const keep = st.jobs.filter((j) => !(TERMINAL_STATUSES.indexOf(j.status) >= 0 && msOf(j.updatedAt) < cutMs));
      const removed = st.jobs.length - keep.length;
      if (removed) { st.jobs = keep; io.save(st); }
      return removed;
    }),

    stats: ({ now }) => run(() => {
      const st = io.load();
      const counts = emptyCounts();
      let oldestPending = null;
      let failed24h = 0;
      const since = msOf(now) - 24 * 3600 * 1000;
      st.jobs.forEach((j) => {
        counts[j.status] += 1;
        if (j.status === 'pending' && (oldestPending === null || j.createdAt < oldestPending)) oldestPending = j.createdAt;
        if (j.status === 'failed' && msOf(j.updatedAt) >= since) failed24h += 1;
      });
      return { counts, oldestPendingAt: oldestPending, failed24h, total: st.jobs.length };
    }),

    breakerLoad: () => run(() => { const st = io.load(); return st.breaker ? JSON.parse(JSON.stringify(st.breaker)) : null; }),
    breakerSave: (state) => run(() => { const st = io.load(); st.breaker = JSON.parse(JSON.stringify(state)); io.save(st); }),
  };
}

/** 在 Vercel 上拒絕使用非資料庫的 adapter（見檔頭）。 */
function refuseOnVercel(kind, opts) {
  if (process.env.VERCEL && !(isObj(opts) && opts.allowOnVercel === true)) {
    throw new MailOutboxError('NOT_FOR_VERCEL', kind + ' adapter 不能用在 Vercel（沒有可寫的持久檔案系統／重啟即消失）；請改用 postgresAdapter');
  }
}

function freshState() { return { version: 1, jobs: [], breaker: null }; }

// ═════════════════════════════════════════════════════════════════════════
// 記憶體 adapter
// ═════════════════════════════════════════════════════════════════════════
function memoryAdapter(opts) {
  refuseOnVercel('memory', opts);
  let text = JSON.stringify(freshState());
  const io = {
    load() { const o = JSON.parse(text); return { version: 1, jobs: (o.jobs || []).map(normalizeRecord).filter(Boolean), breaker: o.breaker || null }; },
    save(st) { text = JSON.stringify(st); },
  };
  const a = createStateAdapter(io, 'memory');
  a.dump = () => text;                    // 測試用：看實際序列化進儲存體的內容
  return a;
}

// ═════════════════════════════════════════════════════════════════════════
// JSON 檔案 adapter
// ═════════════════════════════════════════════════════════════════════════
function sleepSync(ms) {
  try { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms); } catch (e) {
    const end = Date.now() + ms;
    while (Date.now() < end) { /* 沒有 SharedArrayBuffer 時才會走到這裡 */ }
  }
}

/**
 * @param {{file?:string, rootDir?:string, fs?:object}} opts
 *   file    相對路徑以 rootDir（預設 repo 根）為基準；絕對路徑原樣使用。預設 'mail-outbox.json'
 *   fs      可注入（測試用：模擬 rename 失敗等）
 *
 * 限制：只保證「同一個 Node 程序內」的原子性。多個程序同時寫同一個檔案會有後寫者覆蓋的風險（本機開發用途，正式環境用 Postgres）。
 * 暫存檔命名 <file>.<pid>.<rand>.tmp、損壞備份命名 <file>.corrupt-<ts>.bak：兩者都落在 .gitignore 既有的 *.tmp／*.bak 規則內。
 */
function jsonFileAdapter(opts) {
  refuseOnVercel('json', opts);
  const o = isObj(opts) ? opts : {};
  const fsx = o.fs || fs;
  const rootDir = typeof o.rootDir === 'string' && o.rootDir ? o.rootDir : path.join(__dirname, '..', '..');
  const rel = typeof o.file === 'string' && o.file ? o.file : 'mail-outbox.json';
  const file = path.isAbsolute(rel) ? rel : path.join(rootDir, rel);
  const diag = { recoveries: [], seen: new Set() };

  function note(entry) {
    diag.recoveries.push(entry);
    if (diag.recoveries.length > MAX_RECOVERY_NOTES) diag.recoveries.splice(0, diag.recoveries.length - MAX_RECOVERY_NOTES);
  }

  /** 備份檔只留最新的 MAX_CORRUPT_BACKUPS 份（檔名內的時間戳等長，字典序＝時間序）。任何檔案系統錯誤都忽略：修剪失敗不能影響讀取。 */
  function pruneBackups() {
    try {
      const dir = path.dirname(file);
      const prefix = path.basename(file) + '.corrupt-';
      const names = fsx.readdirSync(dir).filter((n) => n.indexOf(prefix) === 0 && /\.bak$/.test(n)).sort();
      for (let i = 0; i < names.length - MAX_CORRUPT_BACKUPS; i++) {
        try { fsx.unlinkSync(path.join(dir, names[i])); } catch (e) { /* 刪不掉就留著，下次再修剪 */ }
      }
    } catch (e) { /* 目錄讀不到（或注入的 fs 沒有 readdirSync）：略過 */ }
  }

  /**
   * 備份損壞的檔案。同一份損壞內容（sha256 相同）只備份一次——唯讀操作（list／stats／get）不會寫檔，
   * 損壞紀錄會一直留在原檔裡，若每次讀取都備份，備份檔會無上限增長（等於整個 outbox 大小 × 讀取次數）。
   * mode 'move'＝整檔壞掉，搬走原檔（搬不動就複製）；'copy'＝只壞幾筆，原檔其餘部分仍在使用，只複製。
   * 備份成功後修剪舊備份。備份失敗不記為「已備份」，下次讀取會再試一次。
   * 限制：「已備份過」只記在這個 adapter 實例的記憶體；程序重啟後第一次讀取會再備份一次，但仍受份數上限約束。
   */
  function backupCorrupt(raw, mode, reason) {
    const sig = crypto.createHash('sha256').update(raw).digest('hex');
    if (diag.seen.has(sig)) return null;
    const stamp = new Date().toISOString().replace(/[-:.]/g, '');
    const dest = file + '.corrupt-' + stamp + '.bak';
    let backup = null;
    if (mode === 'move') {
      try { fsx.renameSync(file, dest); backup = dest; } catch (e) { backup = null; }
    }
    if (!backup) {
      try { fsx.copyFileSync(file, dest); backup = dest; } catch (e2) { backup = null; }
    }
    if (backup) {
      diag.seen.add(sig);
      if (diag.seen.size > MAX_SEEN_CORRUPT) diag.seen.delete(diag.seen.values().next().value);
      pruneBackups();
    }
    note({ at: new Date().toISOString(), reason, backup: backup ? path.basename(backup) : null });
    return backup;
  }

  function load() {
    let raw;
    try {
      raw = fsx.readFileSync(file, 'utf8');
    } catch (e) {
      if (e && e.code === 'ENOENT') return freshState();
      throw new MailOutboxError('STORE_IO', 'outbox 檔案無法讀取：' + (e && e.code ? e.code : 'ERR'));
    }
    if (raw.length > 0 && raw.charCodeAt(0) === 0xfeff) raw = raw.slice(1);       // 去 BOM
    if (raw.trim() === '') return freshState();
    let obj;
    try { obj = JSON.parse(raw); } catch (e) { obj = undefined; }
    if (!isObj(obj) || !Array.isArray(obj.jobs)) {
      backupCorrupt(raw, 'move', 'PARSE_OR_SHAPE');
      return freshState();
    }
    const jobs = [];
    let dropped = 0;
    obj.jobs.forEach((r) => { const n = normalizeRecord(r); if (n) jobs.push(n); else dropped += 1; });
    if (dropped > 0) {
      // 有部分紀錄壞掉：保留合格的，並把原檔複製一份備份（不搬走，因為檔案其餘部分仍在使用）。同一狀態只備份一次
      backupCorrupt(raw, 'copy', 'DROPPED_' + dropped);
    }
    return { version: 1, jobs, breaker: isObj(obj.breaker) ? obj.breaker : null };
  }

  function renameWithRetry(from, to) {
    let lastErr = null;
    for (let i = 0; i < 6; i++) {
      try { fsx.renameSync(from, to); return; } catch (e) {
        lastErr = e;
        // Windows 上防毒軟體／索引服務偶爾會短暫鎖住檔案
        if (!e || (e.code !== 'EPERM' && e.code !== 'EBUSY' && e.code !== 'EACCES')) break;
        sleepSync(15 * (i + 1));
      }
    }
    throw lastErr;
  }

  function save(st) {
    const text = JSON.stringify({ version: 1, jobs: st.jobs, breaker: st.breaker || null }, null, 1) + '\n';
    const tmp = file + '.' + process.pid + '.' + crypto.randomBytes(4).toString('hex') + '.tmp';
    try {
      fsx.mkdirSync(path.dirname(file), { recursive: true });
      fsx.writeFileSync(tmp, text, { encoding: 'utf8', flag: 'wx' });
      renameWithRetry(tmp, file);
    } catch (e) {
      try { fsx.unlinkSync(tmp); } catch (e2) { /* 暫存檔可能根本沒建立 */ }
      throw new MailOutboxError('STORE_IO', 'outbox 檔案無法寫入：' + (e && e.code ? e.code : 'ERR'));
    }
  }

  const a = createStateAdapter({ load, save }, 'json');
  a.file = file;
  a.diagnostics = () => ({ file, recoveries: diag.recoveries.slice() });
  return a;
}

// ═════════════════════════════════════════════════════════════════════════
// Postgres adapter
// ═════════════════════════════════════════════════════════════════════════
/*
 * SQL 靜態審查備註（給審查者；本機沒有 Postgres，以下沒有對真實資料庫執行過）：
 *  1. 所有 SQL 都是下面 SQL／DDL 物件裡的常數字串；沒有任何字串拼接或模板插值。執行前 q() 會檢查「sql 必須是這些常數之一」，
 *     並檢查每個參數只能是 string／number／boolean／null 或字串陣列（防止物件被轉成 [object Object]）。
 *  2. 值一律走 $n 參數；欄位名稱、狀態字面值（'pending' 等）是常數寫死，不來自輸入。
 *  3. 時間一律用 $n::timestamptz，參數是 ISO 字串，時間來源是應用程式的注入時鐘（不用資料庫 now()），測試才能固定時間。
 *  4. 去重：mail_outbox.dedupe_key 有唯一索引；insert 用 ON CONFLICT (dedupe_key) DO NOTHING RETURNING *，沒回列就再 SELECT 既有那筆。
 *  5. 領取（claimDue）：單一條 UPDATE … WHERE id IN (SELECT … FOR UPDATE SKIP LOCKED) RETURNING *；外層 WHERE 再重複一次到期條件
 *     （READ COMMITTED 下被別人先更新的列會被跳過／重新檢查）。兩個並行領取不會拿到同一列。
 *  6. 狀態轉換都帶「現況狀態」條件（WHERE status = …），finishAttempt 另外要求 attempts 相同（樂觀鎖），所以過期的工作者不能覆蓋新的結果。
 *  7. CHECK 約束維持不變條件：status 只能是六種；pending 一定有 next_attempt_at；sending 一定有 lease_until。
 *  8. 啟用 Row Level Security（不建任何 policy）：Supabase 會把 public schema 的表經 PostgREST 開放給 anon 金鑰，沒開 RLS 就等於可被匿名讀取。
 *     應用程式用 DATABASE_URL（資料表擁有者）直連，不受 RLS 影響。【未驗證】請在 Demo 用 get_advisors 確認沒有 RLS 警告。
 *  9. 並行建表：CREATE TABLE IF NOT EXISTS 在多個冷啟動同時執行時，偶爾會因競態拋出「重複鍵」錯誤；ensure() 失敗會重試一次。
 */
const DDL = Object.freeze({
  createOutbox: `CREATE TABLE IF NOT EXISTS mail_outbox (
  id                TEXT PRIMARY KEY,
  type              TEXT NOT NULL,
  quote_id          TEXT NOT NULL,
  quote_no          TEXT NOT NULL DEFAULT '',
  to_user           TEXT NOT NULL,
  to_masked         TEXT NOT NULL DEFAULT '',
  dedupe_key        TEXT NOT NULL,
  status            TEXT NOT NULL,
  attempts          INTEGER NOT NULL DEFAULT 0,
  attempts_in_round INTEGER NOT NULL DEFAULT 0,
  requeues          INTEGER NOT NULL DEFAULT 0,
  next_attempt_at   TIMESTAMPTZ,
  lease_until       TIMESTAMPTZ,
  last_error_code   TEXT,
  last_error_msg    TEXT,
  skip_reason       TEXT,
  actor_label       TEXT NOT NULL DEFAULT '',
  meta              JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at        TIMESTAMPTZ NOT NULL,
  updated_at        TIMESTAMPTZ NOT NULL,
  sent_at           TIMESTAMPTZ,
  CONSTRAINT mail_outbox_status_chk  CHECK (status IN ('pending','sending','sent','failed','skipped','cancelled')),
  CONSTRAINT mail_outbox_pending_chk CHECK (status <> 'pending' OR next_attempt_at IS NOT NULL),
  CONSTRAINT mail_outbox_sending_chk CHECK (status <> 'sending' OR lease_until IS NOT NULL)
)`,
  uniqueKey: 'CREATE UNIQUE INDEX IF NOT EXISTS mail_outbox_dedupe_key_uq ON mail_outbox (dedupe_key)',
  dueIdx: "CREATE INDEX IF NOT EXISTS mail_outbox_due_idx ON mail_outbox (next_attempt_at) WHERE status = 'pending'",
  leaseIdx: "CREATE INDEX IF NOT EXISTS mail_outbox_lease_idx ON mail_outbox (lease_until) WHERE status = 'sending'",
  createdIdx: 'CREATE INDEX IF NOT EXISTS mail_outbox_created_idx ON mail_outbox (created_at DESC)',
  rlsOutbox: 'ALTER TABLE mail_outbox ENABLE ROW LEVEL SECURITY',
  createBreaker: `CREATE TABLE IF NOT EXISTS mail_breaker (
  id         TEXT PRIMARY KEY,
  state      JSONB NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL
)`,
  rlsBreaker: 'ALTER TABLE mail_breaker ENABLE ROW LEVEL SECURITY',
});
const DDL_ORDER = Object.freeze(['createOutbox', 'uniqueKey', 'dueIdx', 'leaseIdx', 'createdIdx', 'rlsOutbox', 'createBreaker', 'rlsBreaker']);

const DUE_COND = `((status = 'pending' AND next_attempt_at <= $1::timestamptz)
      OR (status = 'sending' AND lease_until <= $1::timestamptz AND attempts_in_round < $4))`;

const SQL = Object.freeze({
  // params: [id, type, quote_id, quote_no, to_user, to_masked, dedupe_key, status, attempts, attempts_in_round, requeues,
  //          next_attempt_at, lease_until, last_error_code, last_error_msg, skip_reason, actor_label, meta(json 字串),
  //          created_at, updated_at, sent_at]   共 21 個
  insert: `INSERT INTO mail_outbox
  (id, type, quote_id, quote_no, to_user, to_masked, dedupe_key, status, attempts, attempts_in_round, requeues,
   next_attempt_at, lease_until, last_error_code, last_error_msg, skip_reason, actor_label, meta, created_at, updated_at, sent_at)
VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11,
        $12::timestamptz, $13::timestamptz, $14, $15, $16, $17, $18::jsonb, $19::timestamptz, $20::timestamptz, $21::timestamptz)
ON CONFLICT (dedupe_key) DO NOTHING
RETURNING *`,
  // params: [id]
  getById: 'SELECT * FROM mail_outbox WHERE id = $1',
  // params: [dedupe_key]
  getByKey: 'SELECT * FROM mail_outbox WHERE dedupe_key = $1',
  // params: [id, now, leaseUntil]
  claimById: `UPDATE mail_outbox
   SET status = 'sending', lease_until = $3::timestamptz, attempts = attempts + 1,
       attempts_in_round = attempts_in_round + 1, updated_at = $2::timestamptz
 WHERE id = $1 AND status = 'pending' AND next_attempt_at <= $2::timestamptz
RETURNING *`,
  // params: [now, reasonMsg, maxRoundAttempts]   RETURNING *：把被改成 failed 的列回傳，派送器才能記 Summary／稽核（不再靜默）
  expireExhausted: `UPDATE mail_outbox
   SET status = 'failed', last_error_code = 'LEASE_EXPIRED', last_error_msg = $2, lease_until = NULL, updated_at = $1::timestamptz
 WHERE status = 'sending' AND lease_until <= $1::timestamptz AND attempts_in_round >= $3
RETURNING *`,
  // params: [now, leaseUntil, limit, maxRoundAttempts]
  claimDue: `UPDATE mail_outbox
   SET status = 'sending', lease_until = $2::timestamptz, attempts = attempts + 1,
       attempts_in_round = attempts_in_round + 1, updated_at = $1::timestamptz
 WHERE id IN (
         SELECT id FROM mail_outbox
          WHERE ${DUE_COND}
          ORDER BY next_attempt_at ASC, created_at ASC
          LIMIT $3
          FOR UPDATE SKIP LOCKED)
   AND ${DUE_COND}
RETURNING *`,
  // params: [id, now]
  markSent: `UPDATE mail_outbox
   SET status = 'sent', sent_at = $2::timestamptz, lease_until = NULL, next_attempt_at = NULL, updated_at = $2::timestamptz
 WHERE id = $1 AND status IN ('pending', 'sending', 'failed', 'cancelled')
RETURNING *`,
  // params: [id, expectAttempts, now, status('pending'|'failed'), nextAttemptAt|null, code, msg]
  finishAttempt: `UPDATE mail_outbox
   SET status = $4, next_attempt_at = $5::timestamptz, lease_until = NULL, last_error_code = $6, last_error_msg = $7,
       updated_at = $3::timestamptz
 WHERE id = $1 AND status = 'sending' AND attempts = $2
RETURNING *`,
  // params: [id, now, status('skipped'|'cancelled'), reason, fromStatuses(text[])]
  setTerminal: `UPDATE mail_outbox
   SET status = $3, skip_reason = $4, lease_until = NULL, next_attempt_at = NULL, updated_at = $2::timestamptz
 WHERE id = $1 AND status = ANY($5::text[])
RETURNING *`,
  // params: [id, now]
  requeue: `UPDATE mail_outbox
   SET status = 'pending', next_attempt_at = $2::timestamptz, lease_until = NULL, attempts_in_round = 0,
       requeues = requeues + 1, skip_reason = NULL, updated_at = $2::timestamptz
 WHERE id = $1 AND status IN ('failed', 'cancelled', 'skipped')
RETURNING *`,
  // params: [status|null, type|null, quoteNo|null, toUser|null, since|null, limit, offset]
  list: `SELECT * FROM mail_outbox
 WHERE ($1::text IS NULL OR status = $1)
   AND ($2::text IS NULL OR type = $2)
   AND ($3::text IS NULL OR quote_no = $3)
   AND ($4::text IS NULL OR to_user = $4)
   AND ($5::timestamptz IS NULL OR created_at >= $5::timestamptz)
 ORDER BY created_at DESC, id DESC
 LIMIT $6 OFFSET $7`,
  // params: [status|null, type|null, quoteNo|null, toUser|null, since|null]
  listCount: `SELECT count(*)::int AS n FROM mail_outbox
 WHERE ($1::text IS NULL OR status = $1)
   AND ($2::text IS NULL OR type = $2)
   AND ($3::text IS NULL OR quote_no = $3)
   AND ($4::text IS NULL OR to_user = $4)
   AND ($5::timestamptz IS NULL OR created_at >= $5::timestamptz)`,
  // params: [cutoff]
  purge: `WITH d AS (
  DELETE FROM mail_outbox
   WHERE status IN ('sent', 'failed', 'skipped', 'cancelled') AND updated_at < $1::timestamptz
  RETURNING 1)
SELECT count(*)::int AS n FROM d`,
  // params: []
  statsByStatus: 'SELECT status, count(*)::int AS n, min(created_at) AS oldest FROM mail_outbox GROUP BY status',
  // params: [since]
  statsFailedSince: "SELECT count(*)::int AS n FROM mail_outbox WHERE status = 'failed' AND updated_at >= $1::timestamptz",
  // params: []
  breakerGet: "SELECT state FROM mail_breaker WHERE id = 'main'",
  // params: [stateJson, now]
  breakerSet: `INSERT INTO mail_breaker (id, state, updated_at) VALUES ('main', $1::jsonb, $2::timestamptz)
ON CONFLICT (id) DO UPDATE SET state = EXCLUDED.state, updated_at = EXCLUDED.updated_at`,
});

const SQL_SET = new Set(Object.keys(SQL).map((k) => SQL[k]).concat(Object.keys(DDL).map((k) => DDL[k])));

function isParamOk(p) {
  if (p === null || typeof p === 'string' || typeof p === 'boolean') return true;
  if (typeof p === 'number') return Number.isFinite(p);
  if (Array.isArray(p)) return p.every((x) => typeof x === 'string');
  return false;
}

function insertParams(r) {
  return [
    r.id, r.type, r.quoteId, r.quoteNo, r.toUser, r.toMasked, r.dedupeKey, r.status,
    r.attempts, r.attemptsInRound, r.requeues,
    r.nextAttemptAt, r.leaseUntil, r.lastErrorCode, r.lastErrorMsg, r.skipReason, r.actorLabel,
    JSON.stringify(r.meta || {}),
    r.createdAt, r.updatedAt, r.sentAt,
  ];
}

/**
 * @param {{query: function(string, any[]): Promise<{rows: object[]}>}} opts
 *   query  例如 (sql, params) => pool.query(sql, params)（pg 的 Pool.query 就符合）。連線池設定沿用 db/postgres.js 的慣例
 *          （serverless 建議 max 2、ssl、connectionTimeoutMillis），由接線的人建立，本檔不自己連線。
 */
function postgresAdapter(opts) {
  const query = opts && opts.query;
  if (typeof query !== 'function') throw new MailOutboxError('BAD_ADAPTER', 'postgresAdapter 需要 query(sql, params) 函式');

  async function q(sql, params) {
    if (!SQL_SET.has(sql)) throw new MailOutboxError('DYNAMIC_SQL', '只允許執行 outboxAdapters.js 內定義的 SQL 常數');
    const p = params || [];
    if (!Array.isArray(p) || !p.every(isParamOk)) throw new MailOutboxError('BAD_PARAM', 'SQL 參數只能是字串、數字、布林、null 或字串陣列');
    const res = await query(sql, p);
    return res && Array.isArray(res.rows) ? res.rows : [];
  }

  let schemaP = null;
  function ensure() {
    if (!schemaP) {
      schemaP = (async () => {
        for (let i = 0; i < DDL_ORDER.length; i++) {
          const sql = DDL[DDL_ORDER[i]];
          try { await q(sql, []); } catch (e) { await q(sql, []); }      // 並行建表競態：重試一次
        }
      })().catch((e) => { schemaP = null; throw e; });
    }
    return schemaP;
  }
  const one = (rows) => (rows.length ? rowToRecord(rows[0]) : null);

  return {
    kind: 'postgres',
    SQL,
    DDL,
    init: () => ensure(),
    /** 只供測試／診斷：走與內部完全相同的 q() 守衛（非 SQL 常數 → DYNAMIC_SQL；參數型別不合 → BAD_PARAM），讓守衛本身可以被測試。 */
    execConstant: (sql, params) => q(sql, params),

    async insert(rec) {
      await ensure();
      const rows = await q(SQL.insert, insertParams(rec));
      if (rows.length) return { inserted: true, record: rowToRecord(rows[0]) };
      const existing = await q(SQL.getByKey, [rec.dedupeKey]);
      if (!existing.length) throw new MailOutboxError('INSERT_CONFLICT', 'dedupe_key 衝突但找不到既有紀錄（可能剛被 purge）');
      return { inserted: false, record: rowToRecord(existing[0]) };
    },
    async get(id) { await ensure(); return one(await q(SQL.getById, [id])); },
    async getByKey(key) { await ensure(); return one(await q(SQL.getByKey, [key])); },

    async claimById({ id, now, leaseUntil }) {
      await ensure();
      return one(await q(SQL.claimById, [id, now, leaseUntil]));
    },
    async expireExhausted({ now, maxRoundAttempts }) {
      await ensure();
      const rows = await q(SQL.expireExhausted, [now, LEASE_EXPIRED_MSG, maxRoundAttempts]);
      return rows.map(rowToRecord).filter(Boolean);
    },
    async claimDue({ now, leaseUntil, limit, maxRoundAttempts }) {
      await ensure();
      const rows = await q(SQL.claimDue, [now, leaseUntil, limit, maxRoundAttempts]);
      return rows.map(rowToRecord).filter(Boolean);
    },
    async markSent({ id, now }) {
      await ensure();
      return one(await q(SQL.markSent, [id, now]));
    },
    async finishAttempt({ id, expectAttempts, now, status, nextAttemptAt, code, msg }) {
      await ensure();
      return one(await q(SQL.finishAttempt, [id, expectAttempts, now, status, nextAttemptAt, code, msg]));
    },
    async setTerminal({ id, now, status, reason, from }) {
      await ensure();
      return one(await q(SQL.setTerminal, [id, now, status, reason, from]));
    },
    async requeue({ id, now }) {
      await ensure();
      return one(await q(SQL.requeue, [id, now]));
    },
    async list(f) {
      await ensure();
      const filt = [f.status || null, f.type || null, f.quoteNo || null, f.toUser || null, f.since || null];
      const rows = await q(SQL.list, filt.concat([f.limit, f.offset]));
      const cnt = await q(SQL.listCount, filt);
      return { rows: rows.map(rowToRecord).filter(Boolean), total: cnt.length ? Number(cnt[0].n) || 0 : 0 };
    },
    async purge({ cutoff }) {
      await ensure();
      const rows = await q(SQL.purge, [cutoff]);
      return rows.length ? Number(rows[0].n) || 0 : 0;
    },
    async stats({ now }) {
      await ensure();
      const rows = await q(SQL.statsByStatus, []);
      const counts = emptyCounts();
      let oldest = null;
      let total = 0;
      rows.forEach((r) => {
        if (STATUSES.indexOf(r.status) < 0) return;
        const n = Number(r.n) || 0;
        counts[r.status] = n;
        total += n;
        if (r.status === 'pending' && r.oldest) oldest = isoOf(r.oldest);
      });
      const since = new Date(msOf(now) - 24 * 3600 * 1000).toISOString();
      const f = await q(SQL.statsFailedSince, [since]);
      return { counts, oldestPendingAt: oldest, failed24h: f.length ? Number(f[0].n) || 0 : 0, total };
    },
    async breakerLoad() {
      await ensure();
      const rows = await q(SQL.breakerGet, []);
      if (!rows.length) return null;
      let s = rows[0].state;
      if (typeof s === 'string') { try { s = JSON.parse(s); } catch (e) { s = null; } }
      return isObj(s) ? s : null;
    },
    async breakerSave(state) {
      await ensure();
      await q(SQL.breakerSet, [JSON.stringify(state), new Date().toISOString()]);
    },
  };
}
postgresAdapter.SQL = SQL;
postgresAdapter.DDL = DDL;
postgresAdapter.DDL_ORDER = DDL_ORDER;

module.exports = {
  STATUSES,
  TERMINAL_STATUSES,
  MailOutboxError,
  MAX_CORRUPT_BACKUPS,
  jsonFileAdapter,
  memoryAdapter,
  postgresAdapter,
  // 以下給 outbox.js 與測試使用
  normalizeRecord,
  isoOf,
  msOf,
};
