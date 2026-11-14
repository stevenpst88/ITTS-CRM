'use strict';
/**
 * 請求層的「全域寫入序列化」（DB_WRITE_SERIALIZE=1 才啟用；預設完全不介入）。
 * 流程：已驗證的寫入型請求 → 行程內 FIFO 互斥 → （Postgres）DB advisory 鎖 → 重新 db.ready() 取最新 blob →
 *       handler（同步 load／改／save）→ res.end 前 flush → 本請求自己的存檔失敗改回 503 → 釋放鎖 → 再往下走（寄信等待、稽核 flush、真正送出）。
 * 設計與取捨見 DEPLOY.md「寫入競爭防護」。
 *
 * 請求歸屬：開啟時每個 /api 請求都在 AsyncLocalStorage（lib/requestScope.js）裡帶一個 scope 物件，db.save 透過 onSave hook 把
 * 「這次存檔的 promise」記到發起它的 scope——所以「本請求的存檔有沒有失敗」「這次存檔有沒有在本請求持有的鎖內」都是以請求為單位判斷，
 * 不受其他並行請求影響。（multer 之類由 socket 事件驅動的回呼會丟失 context：那種情況歸給目前唯一的持鎖者；沒有持鎖者就視為鎖外寫入。）
 */

const RS = require('./requestScope');

const WRITE_METHODS = new Set(['POST', 'PUT', 'PATCH', 'DELETE']);

/**
 * 內建豁免＝不拿鎖（會呼叫外部服務，或根本不寫主 blob）。字串：以 / 結尾＝只比對其底下的路徑（不含該路徑本身）；不以 / 結尾＝該路徑本身及其所有子路徑；RegExp＝整條路徑比對。
 * 一律小寫比對（路徑會先正規化成小寫、去尾斜線）。內建清單盡量用完整路徑而不是寬鬆前綴：純 DB 的同步寫入路由（如 /api/push/subscribe）必須拿鎖。
 * 注意「不以 / 結尾」的項目同時豁免子路徑（DB_WRITE_LOCK_EXEMPT 同規則，見 DEPLOY.md）。
 */
const DEFAULT_EXEMPT = [
  '/api/login', '/api/logout',
  '/api/ai/', '/api/admin/ai-ocr-card',             // Gemini
  '/api/admin/integrations/test',                    // SAP 連線測試（外部呼叫）
  '/api/admin/integrations/push/batch',              // SAP 批次推送（長時間外部呼叫）
  '/api/admin/companies/enrich-gcis',                // GCIS
  '/api/admin/bulk-fill-website',                    // 外部查詢
  '/api/company-lookup',                             // GCIS／證交所
  '/api/push/test',                                  // Web Push 測試
  '/api/upload',                                     // 名片圖片上傳：只存檔（含外部儲存 PUT），不碰 blob
  '/api/admin/mail/outbox/',                         // 寄件匣（獨立資料表／檔案，不寫主 blob）；requeue 會立刻實際寄信（外部呼叫）
  /^\/api\/quotations\/[^/]+\/export-log$/,          // 只寫稽核
  /^\/api\/quotations\/[^/]+\/cost-draft\//,         // 草稿試算，不寫入
  '/api/admin/quote-pricebook/import-preview',       // 只解析 xlsx
];
/** 會寫入 blob 的 GET（惰性遷移／已讀回條）。其餘 GET 不拿鎖 */
const DEFAULT_GET_WRITERS = ['/api/callins', '/api/announcements/active', '/api/admin/quote-ledger'];

function csv(s) { return String(s || '').split(',').map((x) => x.trim().toLowerCase()).filter(Boolean); }
function posInt(v, d) { const n = Math.floor(Number(v)); return Number.isFinite(n) && n > 0 ? n : d; }
/** 持有上限的硬上限：DB 端 idle_in_transaction_session_timeout 保底是 25s，持有上限超過它就沒有意義（DB 會先把鎖收回）；Vercel maxDuration 是 30s */
const MAX_HOLD_MS = 22000;

function readConfig(env) {
  env = env || process.env;
  return {
    serialize: env.DB_WRITE_SERIALIZE === '1',
    waitMs: posInt(env.DB_WRITE_LOCK_WAIT_MS, 8000),
    holdMs: Math.min(posInt(env.DB_WRITE_LOCK_HOLD_MS, 20000), MAX_HOLD_MS),
    exempt: DEFAULT_EXEMPT.concat(csv(env.DB_WRITE_LOCK_EXEMPT)),
    getWriters: DEFAULT_GET_WRITERS.concat(csv(env.DB_WRITE_LOCK_EXTRA_PATHS)),
  };
}

/**
 * 路徑正規化：Express 預設路由不分大小寫、不嚴格比對尾斜線（/api/Callins、/api/callins/ 都會進同一個 handler；
 * /api/ 前綴本身另有 server.js 最上方的中介層強制小寫）。比對一律用小寫、去尾斜線的路徑。
 */
function normPath(p) {
  const s = String(p == null ? '' : p).toLowerCase().replace(/\/+$/, '');
  return s === '' ? '/' : s;
}
function isExempt(cfg, p) {
  const n = normPath(p);
  return cfg.exempt.some((e) => (e instanceof RegExp ? e.test(n) : (e.endsWith('/') ? n.startsWith(e) : n === e || n.startsWith(e + '/'))));
}
/** 這個請求需不需要拿寫入鎖（不含「是否已登入」的判斷）。HEAD 會落到 GET handler，視同 GET */
function needsLock(cfg, method, p) {
  const m = String(method || '').toUpperCase();
  const n = normPath(p);
  if (!n.startsWith('/api/')) return false;
  if (isExempt(cfg, n)) return false;
  if (WRITE_METHODS.has(m)) return true;
  if (m === 'GET' || m === 'HEAD') return cfg.getWriters.includes(n);
  return false;
}
function isMultipart(req) {
  const ct = req && req.headers && (req.headers['content-type'] || '');
  return /^multipart\/form-data/i.test(String(ct));
}

class GuardError extends Error {
  constructor(code, message, status) { super(message); this.name = 'GuardError'; this.code = code; this.status = status || 503; }
}
const BUSY_BODY = {
  code: 'DB_BUSY', retryable: true,
  error: '系統目前較忙，您的操作尚未儲存。請稍候幾秒後再按一次（資料不會遺失，也不會重複）。',
};
const WRITE_FAILED_BODY = {
  code: 'DB_WRITE_FAILED', retryable: true,
  error: '資料暫時無法儲存，這次操作可能沒有生效。請稍候再試一次，並確認清單中是否已有這筆資料，避免重複建立。',
};
const sleep = (ms) => new Promise((r) => { const t = setTimeout(r, ms); if (t.unref) t.unref(); });

/**
 * @param {object} o
 * @param {object} o.db        { ready, flush, [acquireDbWriteLock], [flushMain], [invalidateCache], [setSaveHook] }（db/index.js 或測試用 fake）
 * @param {object} [o.config]  readConfig() 的結果
 * @param {object} [o.log]     console 相容
 * @param {function} [o.preDeny] (req, fullPath) => true 表示這個請求一定會被後面的權限中介層拒絕（例如非管理員打 /api/admin/*），不必排隊拿鎖
 */
function createWriteGuard(o) {
  const db = o.db;
  const cfg = o.config || readConfig();
  const log = o.log || console;
  const preDeny = typeof o.preDeny === 'function' ? o.preDeny : null;

  // ── 行程內 FIFO 互斥（帶逾時）──
  let held = false;
  const waiters = [];
  function acquireLocal(waitMs) {
    return new Promise((resolve, reject) => {
      const grant = () => {
        held = true;
        let done = false;
        resolve(() => { if (done) return; done = true; held = false; const n = waiters.shift(); if (n) { clearTimeout(n.timer); n.grant(); } });
      };
      if (!held) return grant();
      const w = { grant, timer: null };
      w.timer = setTimeout(() => {
        const i = waiters.indexOf(w);
        if (i >= 0) waiters.splice(i, 1);
        reject(new GuardError('DB_BUSY', '等待寫入鎖逾時（行程內排隊）', 503));
      }, waitMs);
      waiters.push(w);
    });
  }

  // 「確實在路由內 multer 之後掛了 lateLock(路由路徑)」的路徑集合（小寫、去尾斜線）。只有這些路徑的 multipart 請求才延後取鎖；
  // 其他路由即使偽造 Content-Type: multipart/form-data 也照一般請求在全域中介層取鎖（審查 R2-#6）
  const lateRoutes = new Set();

  let _holders = 0;
  let _holderScope = null;   // 目前持鎖者的 scope（同行程同時最多一個：行程內互斥保證）
  const newScope = (label) => ({ label, locked: false, released: false, saves: [] });

  /** 取得寫入鎖，回傳 { release():Promise, forced, waitedMs, scope }。逾時丟 GuardError(DB_BUSY)；DB 連不上丟 GuardError(DB_UNAVAILABLE) */
  async function acquire(label, opts) {
    opts = opts || {};
    const waitMs = posInt(opts.waitMs, cfg.waitMs);
    const holdMs = Math.min(posInt(opts.holdMs, cfg.holdMs), MAX_HOLD_MS);
    const t0 = Date.now();
    const releaseLocal = await acquireLocal(waitMs);
    let dbh = null;
    try {
      if (typeof db.acquireDbWriteLock === 'function') {
        dbh = await db.acquireDbWriteLock({ waitMs: Math.max(1, waitMs - (Date.now() - t0)), holdMs: holdMs + 2000 });
      }
    } catch (e) {
      releaseLocal();
      if (e && e.code === 'LOCK_TIMEOUT') throw new GuardError('DB_BUSY', '等待資料庫寫入鎖逾時', 503);
      throw new GuardError('DB_UNAVAILABLE', '取得資料庫寫入鎖失敗：' + (e && e.message), 503);
    }
    _holders++;
    const scope = newScope(label);
    scope.locked = true;
    _holderScope = scope;
    let released = false;
    let timer = null;
    const handle = {
      label, forced: false, waitedMs: Date.now() - t0, heldSince: Date.now(), holdMs, scope,
      async release() {
        if (released) return;
        released = true;
        _holders--;
        scope.released = true;           // 這之後從本請求發出的 db.save 就是「鎖外寫入」
        if (_holderScope === scope) _holderScope = null;
        if (timer) { clearTimeout(timer); timer = null; }
        releaseLocal();                  // 先同步放行程內互斥，再等 DB 鎖的 COMMIT（COMMIT 有用戶端逾時，不會永遠卡住）
        try { if (dbh) await dbh.release(); } catch (e) { log.error('[writeGuard] DB 鎖釋放失敗：', e && e.message); }
      },
    };
    timer = setTimeout(() => {
      if (released) return;
      handle.forced = true;
      log.error('[writeGuard] 寫入鎖持有超過 ' + holdMs + 'ms，強制釋放（label=' + label + '）。該請求之後的存檔將不受保護。');
      handle.release().catch(() => {});
    }, holdMs);
    if (timer.unref) timer.unref();
    return handle;
  }

  /**
   * 非請求路徑（冷啟動遷移等）：開關開啟時在鎖內執行 fn；關閉時直接執行。opts { waitMs, holdMs } 可縮短（冷啟動工作用短逾時，取不到就丟 DB_BUSY 由呼叫端略過）
   */
  async function withLock(label, fn, opts) {
    if (!cfg.serialize) return fn();
    const h = await acquire(label, opts);
    try { return await RS.run(h.scope, () => fn(h)); } finally { await h.release(); }
  }

  function sendBusy(res, err) {
    log.warn('[writeGuard] ' + (err && err.code) + '：' + (err && err.message));
    if (res.headersSent) return;
    res.set('Retry-After', '2');
    res.status(err && err.status ? err.status : 503).json(BUSY_BODY);
  }

  // ── db.save 的 hook：把這次存檔歸屬到發起的請求；偵測「鎖外寫入」 ──
  function warnUnlocked(scope) {
    const where = String(new Error().stack || '').split('\n').slice(3, 6).map((x) => x.trim().replace(/^at /, '')).join(' <- ');
    log.warn('[db] UNLOCKED_WRITE：在寫入鎖之外呼叫 db.save（' + ((scope && scope.label) || '無請求脈絡') + '）。請檢查該路由是否需要列入鎖清單。呼叫處：' + where);
  }
  function onSave(p) {
    const st = RS.current();
    if (st) {
      if (st.locked && !st.released) { st.saves.push(p); return; }
      warnUnlocked(st);
      return;
    }
    // 沒有請求脈絡（multer 回呼等丟失了 ALS context、或背景計時器）：有持鎖者就歸給它，否則視為鎖外寫入
    if (_holderScope) { _holderScope.saves.push(p); return; }
    warnUnlocked(null);
  }
  if (cfg.serialize && typeof db.setSaveHook === 'function') db.setSaveHook(onSave);

  /** 取得鎖 → 重新 ready() → 往下走（共用於一般請求與 multipart 延後取鎖） */
  function lockAndRun(req, res, next, full) {
    acquire(req.method + ' ' + full).then(async (handle) => {
      req._writeLock = handle;
      installEnd(req, res, handle);
      try {
        await db.ready();   // 持鎖後重新確認／抓取最新 blob；這一步才是防止丟失更新的關鍵
      } catch (e) {
        log.error('[writeGuard] 持鎖後 ready 失敗：', e && e.message);
        await handle.release();
        if (!res.headersSent) res.status(503).json({ code: 'DB_UNAVAILABLE', retryable: true, error: '資料庫暫時無法使用，請稍後再試' });
        return;
      }
      RS.run(handle.scope, next);
    }).catch((err) => sendBusy(res, err));
  }
  const fullPathOf = (req) => (req.baseUrl || '') + req.path;   // 掛在 app.use('/api/', …) 之下時 req.path 已去掉 /api 前綴，要用 baseUrl 還原完整路徑
  function shouldLock(req, full) {
    if (!needsLock(cfg, req.method, full)) return false;
    if (!(req.session && req.session.user)) return false;       // 未登入：不拿鎖（匿名洪水不能拖慢寫入者）
    if (preDeny && preDeny(req, full)) return false;            // 一定會被權限中介層拒絕：不排隊
    return true;
  }

  /** Express middleware：必須放在驗證／權限／rate-limit 之後，所以匿名請求與被拒絕的請求不會佔鎖 */
  function middleware() {
    return function writeLockMw(req, res, next) {
      if (!cfg.serialize) return next();
      const full = fullPathOf(req);
      // 不拿鎖的請求也要掛一個 scope：它們若自己 db.save（豁免路由、惰性寫入的 GET）才能被精確標成 UNLOCKED_WRITE
      if (!shouldLock(req, full)) return RS.run(newScope(req.method + ' ' + full), next);
      // multipart：要等路由內的 multer 把 body 讀完才取鎖（上傳時間不佔鎖）→ 路由要在 multer 之後掛 lateLock('<路由路徑>')。
      // 只認「已註冊的 multer 路由」：任何人給其他寫入路由加上 multipart 標頭都不能跳過鎖
      if (isMultipart(req) && lateRoutes.has(normPath(full))) return RS.run(newScope(req.method + ' ' + full + ' (multipart，待 lateLock)'), next);
      return lockAndRun(req, res, next, full);
    };
  }

  /**
   * 路由層中介：放在 multer（讀完 body）之後。已經持鎖或不需要鎖就直接放行。
   * routePath＝這條路由的完整路徑（例如 '/api/admin/import-contacts'）；同時把它註冊成「multipart 延後取鎖」的路由。
   */
  function lateLock(routePath) {
    if (routePath) lateRoutes.add(normPath(routePath));
    return function lateLockMw(req, res, next) {
      if (!cfg.serialize || req._writeLock) return next();
      const full = fullPathOf(req);
      if (!shouldLock(req, full)) return next();
      return lockAndRun(req, res, next, full);
    };
  }

  /**
   * 這個 scope 的存檔有沒有失敗：只看「最後一次 save」。每次 db.save 都是整份 blob 覆寫，最後一次成功＝先前（含失敗那次）的變更都已經在裡面；
   * 最後一次失敗＝這次請求的變更沒有落地。（審查 R2-#4：舊版用 allSettled().some(rejected)，前一次暫時失敗、後一次成功也會誤回 503，
   * 使用者依訊息重試會重複建立資料。）
   */
  async function lastSaveFailed(scope) {
    const saves = scope.saves;
    if (!saves.length) return false;
    const r = await Promise.allSettled([saves[saves.length - 1]]);
    return r[0].status === 'rejected';
  }

  /** 包 res.end：先 flush 本次存檔 → 偵測「本請求自己的」存檔失敗 → 釋放鎖 → 才交給原本的 end（它會等寄信、再 flush 稽核） */
  function installEnd(req, res, handle) {
    const origEnd = res.end.bind(res);
    let called = false;
    res.end = function guardedEnd(...args) {
      if (called) return origEnd(...args);
      called = true;
      Promise.resolve()
        .then(async () => {
          // 只等主資料；稽核由後面原本的 end 包裝去 flush（不佔著鎖）。有上限：剩餘的持有時間，超過就視為存檔失敗，不讓使用者一直等到 Vercel 強殺
          const remain = Math.max(1000, handle.holdMs - (Date.now() - handle.heldSince));
          const flushed = typeof db.flushMain === 'function' ? db.flushMain() : db.flush();
          const r = await Promise.race([Promise.resolve(flushed).then(() => 'ok', () => 'ok'), sleep(remain).then(() => 'timeout')]);
          if (r === 'timeout') { log.error('[writeGuard] 等待存檔逾時 ' + remain + 'ms（' + handle.label + '）'); return true; }
          return await lastSaveFailed(handle.scope);   // 以本請求自己發起的「最後一次」存檔結果為準，不看全行程計數器
        })
        .catch((e) => { log.error('[writeGuard] flush 失敗：', e && e.message); return true; })
        .then(async (failed) => {
          // 存檔失敗：必須在「放鎖之前」同步讓快取失效。放鎖（COMMIT）一旦送出，排隊的下一個寫入者隨時可能拿到鎖並走進 ready()；
          // 若此時快取還是含本次失敗變更的那一份，它會沿用並把這次回 503 的變更一起存進 DB（審查 R2-#3）。
          if (failed) { try { if (typeof db.invalidateCache === 'function') db.invalidateCache(); } catch (_) { /* ignore */ } }
          await handle.release();
          if (failed) {
            // 已經在背景寄出的簽核信不會取消（信件與主資料存檔無關）：在 log 留下數量，使用者重試成功後可能多收一封
            const mailN = Array.isArray(req._mailPending) ? req._mailPending.length : 0;
            log.error('[writeGuard] 本次請求的資料庫存檔失敗，改回 503（' + handle.label + '）' + (mailN ? '；此請求已啟動 ' + mailN + ' 封通知信，信件不會取消，使用者重試後可能重複收到' : ''));
            if (!res.headersSent) {
              res.statusCode = 503;
              res.removeHeader('Content-Length'); res.removeHeader('ETag'); res.removeHeader('Content-Encoding');
              res.setHeader('Content-Type', 'application/json; charset=utf-8');
              res.setHeader('Retry-After', '2');
              return origEnd(JSON.stringify(WRITE_FAILED_BODY));
            }
          }
          return origEnd(...args);
        });
      return res;
    };
  }

  /**
   * 「外部呼叫（無鎖）→ 短鎖內重讀最新資料、只套用自己改動的欄位」。給被豁免、await 外部服務之後才要存檔的路由用。
   * 開關關閉＝完全維持舊行為（直接修改 await 前讀到的 legacyData 並整份存回，回傳 true）。
   * 開啟＝取短鎖 → ready() 取最新 blob → mutate(最新資料)（回傳 false 表示沒有東西要存）→ save → flush → 檢查這次 save 有沒有真的成功。
   * 取不到鎖、DB 不通、存檔失敗都「不丟例外」，回傳 false（外部呼叫本身已經成功，不能因此讓整個請求失敗），由呼叫端決定要不要警告使用者。
   * opts: { waitMs=3000, holdMs=8000, attempts=1, retryDelayMs=300 }。mutate 必須可重複套用（冪等）——重試時會在新抓回的資料上再套一次。
   * 存檔失敗時會在放鎖之前讓快取失效（同 guardedEnd 的理由），下一次嘗試／下個請求會整份重抓。
   * @returns {Promise<boolean>} true＝已存入（或沒有東西要存）；false＝沒有存成功
   */
  async function applyOnLatest(label, legacyData, mutate, opts) {
    if (!cfg.serialize) {
      if (mutate(legacyData) !== false) db.save(legacyData);
      return true;
    }
    opts = opts || {};
    const attempts = posInt(opts.attempts, 1);
    const retryDelayMs = opts.retryDelayMs == null ? 300 : Math.max(0, Number(opts.retryDelayMs) || 0);
    const lockOpts = { waitMs: posInt(opts.waitMs, 3000), holdMs: posInt(opts.holdMs, 8000) };
    let lastErr = null;
    for (let i = 1; i <= attempts; i++) {
      try {
        await withLock(label, async (h) => {
          await db.ready();
          const d = db.load();
          if (mutate(d) === false) return;
          db.save(d);
          await (typeof db.flushMain === 'function' ? db.flushMain() : db.flush());
          // flushMain／flush 會吞掉存檔錯誤：要看這次 save 自己的 promise（同樣只看最後一次）
          if (await lastSaveFailed(h.scope)) {
            try { if (typeof db.invalidateCache === 'function') db.invalidateCache(); } catch (_) { /* ignore */ }
            throw new GuardError('DB_WRITE_FAILED', '資料庫存檔失敗', 503);
          }
        }, lockOpts);
        return true;
      } catch (e) {
        lastErr = e;
        log.warn('[applyOnLatest] ' + label + ' 第 ' + i + '/' + attempts + ' 次沒有存成功：' + (e && e.message));
        if (i < attempts && retryDelayMs) await sleep(retryDelayMs);
      }
    }
    log.warn('[applyOnLatest] ' + label + ' 沒有存成功（外部呼叫的結果不受影響）：' + (lastErr && lastErr.message));
    return false;
  }

  return { acquire, withLock, applyOnLatest, middleware, lateLock, isLateRoute: (p) => lateRoutes.has(normPath(p)), onSave, isHeld: () => _holders > 0, config: cfg };
}

module.exports = { createWriteGuard, readConfig, needsLock, isExempt, normPath, isMultipart, GuardError, BUSY_BODY, WRITE_FAILED_BODY, DEFAULT_EXEMPT, DEFAULT_GET_WRITERS, MAX_HOLD_MS };
