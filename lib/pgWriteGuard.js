'use strict';
/**
 * Postgres 層的寫入競爭防護（SQL 層，不碰 Express）。
 *  - acquireLock：請求範圍的全域寫入鎖（專用連線上的交易 + pg_advisory_xact_lock；交易結束或連線斷就自動釋放）
 *  - nextQuoteSeq：報價單號原子計數器（單一 INSERT … ON CONFLICT DO UPDATE … RETURNING）＋已發號紀錄
 * 設計與取捨見 DEPLOY.md「寫入競爭防護」。表名／schema／鎖鍵可注入：正式環境用預設值，
 * scripts/verify-write-race-pg.js 用獨立 schema 與隨機鎖鍵，絕不碰 app_data。
 */

const IDENT_RE = /^[a-z_][a-z0-9_]*$/;
/** 預設鎖鍵（任意固定 bigint）。同一個資料庫的所有實例必須用同一把才會互斥 */
const DEFAULT_LOCK_KEY = '7301420260101';
/** 建表專用 advisory 鎖鍵（與寫入鎖不同把；只在第一次建 quote_no_* 表時用） */
const SCHEMA_LOCK_KEY = '7301420260102';

class WriteLockError extends Error {
  /** code: LOCK_TIMEOUT（等太久）｜LOCK_UNAVAILABLE（連不上／其他 DB 錯誤） */
  constructor(code, message, cause) {
    super(message);
    this.name = 'WriteLockError';
    this.code = code;
    if (cause) this.cause = cause;
  }
}

function ident(name) {
  if (!IDENT_RE.test(name)) throw new Error('不合法的識別字：' + name);
  return '"' + name + '"';
}
function posInt(n, dflt) {
  const v = Math.floor(Number(n));
  return Number.isFinite(v) && v > 0 ? v : dflt;
}
/**
 * 用戶端逾時：p 在 ms 內沒結束就以 CLIENT_TIMEOUT 拒絕。
 * 黑洞連線（Vercel 凍結後解凍的死連線、NAT 靜默丟包、半開 TCP）上的 query 永不 settle，沒有它，COMMIT／BEGIN 就會永遠卡住、行程內互斥永不釋放。
 */
function withTimeout(p, ms) {
  let t;
  const guard = new Promise((_, reject) => { t = setTimeout(() => reject(Object.assign(new Error('client timeout after ' + ms + 'ms'), { code: 'CLIENT_TIMEOUT' })), ms); });
  Promise.resolve(p).catch(() => {});   // 原本的 promise 之後才失敗也不要變成 unhandled rejection
  return Promise.race([p, guard]).finally(() => clearTimeout(t));
}

/** SQL 文字單一來源：lib 內使用，也匯出給測試／驗證腳本比對（不要在別處手寫重複的字串） */
const SQL = {
  /** 取鎖交易的開頭（一次往返送出）。waitMs／idleTxMs 必須是整數，lockKey 必須是整數字串 */
  lockBegin(waitMs, idleTxMs, lockKey) {
    return `BEGIN; SET LOCAL lock_timeout = ${waitMs}; SET LOCAL idle_in_transaction_session_timeout = ${idleTxMs}; SET LOCAL statement_timeout = ${waitMs + 3000}; SELECT pg_advisory_xact_lock(${lockKey})`;
  },
  /** 原子取號。T＝取（含 schema 的）表名函式，seqIdent＝計數器表的識別字。參數：$1 年月、$2 下界、$3 發給誰 */
  nextQuoteSeq(T, seqIdent) {
    return `WITH c AS (
         INSERT INTO ${T('seq')} (ym, last_serial) VALUES ($1::text, $2::int + 1)
         ON CONFLICT (ym) DO UPDATE
           SET last_serial = GREATEST(${seqIdent}.last_serial, $2::int) + 1, updated_at = now()
         RETURNING last_serial
       )
       INSERT INTO ${T('issued')} (quote_no, ym, serial, issued_to)
       SELECT 'QU-' || $1::text || '-' || CASE WHEN c.last_serial < 1000 THEN lpad(c.last_serial::text, 3, '0') ELSE c.last_serial::text END,
              $1::text, c.last_serial, $3::text
       FROM c
       RETURNING serial, quote_no`;
  },
};

/**
 * @param {object} o
 * @param {function():{query:function,connect:function}} o.getPool 回傳 pg.Pool（或相容物件）
 * @param {string} [o.schema] 例：'wrg_verify_ab12'（不給＝預設 search_path）
 * @param {{seq?:string, issued?:string}} [o.tables]
 * @param {string|number} [o.lockKey]
 * @param {number} [o.idleTxMs] 伺服器端保底：交易閒置超過這麼久，DB 端自動終止（預設 25000，須小於 Vercel maxDuration）
 * @param {number} [o.commitTimeoutMs] COMMIT／ROLLBACK 的用戶端逾時（預設 2000）：逾時就銷毀連線，讓伺服器端因連線關閉而釋放 advisory 鎖
 * @param {number} [o.beginSlackMs] 取鎖 BEGIN 在 lock_timeout 之外多給的用戶端緩衝（預設 3000）
 * @param {string|number} [o.schemaLockKey] 建表專用 advisory 鎖鍵（預設 SCHEMA_LOCK_KEY；驗證腳本用隨機值，不與正式環境互相干擾）
 */
function createPgWriteGuard(o) {
  const getPool = o.getPool;
  const schema = o.schema || null;
  const tables = Object.assign({ seq: 'quote_no_seq', issued: 'quote_no_issued' }, o.tables || {});
  const lockKey = String(o.lockKey != null ? o.lockKey : DEFAULT_LOCK_KEY);
  if (!/^-?\d{1,18}$/.test(lockKey)) throw new Error('lockKey 必須是整數');
  const idleTxMs = posInt(o.idleTxMs, 25000);
  const commitTimeoutMs = posInt(o.commitTimeoutMs, 2000);
  const beginSlackMs = posInt(o.beginSlackMs, 3000);
  const schemaLockKey = String(o.schemaLockKey != null ? o.schemaLockKey : SCHEMA_LOCK_KEY);
  if (!/^-?\d{1,18}$/.test(schemaLockKey)) throw new Error('schemaLockKey 必須是整數');
  const T = (k) => (schema ? ident(schema) + '.' : '') + ident(tables[k]);

  let _schemaReady = null;
  /**
   * 建表。兩張表都已存在就完全跳過（to_regclass 檢查，不執行任何 DDL、不拿 ACCESS EXCLUSIVE 鎖）；
   * 要建時用專屬連線的單一交易：先取一把「建表專用」advisory 鎖（並發冷啟動串行，避免 CREATE TABLE IF NOT EXISTS 並發時的 23505），
   * CREATE 與 ENABLE ROW LEVEL SECURITY 在同一交易提交（不會有「已建表、未開 RLS」的空窗）。
   * 新表一律開 RLS（與 app_data 不同；無 policy＝只有表擁有者連線可存取，anon／PostgREST 看不到）。
   */
  async function ensureSchema() {
    if (_schemaReady) return _schemaReady;
    _schemaReady = (async () => {
      const pool = getPool();
      const exists = async (q) => {
        const { rows } = await q.query('SELECT to_regclass($1::text) IS NOT NULL AS a, to_regclass($2::text) IS NOT NULL AS b', [T('seq'), T('issued')]);
        return !!(rows[0] && rows[0].a && rows[0].b);
      };
      if (await exists(pool)) return;
      const client = await pool.connect();
      let broken = false;
      try {
        await client.query(`BEGIN; SET LOCAL lock_timeout = 10000; SELECT pg_advisory_xact_lock(${schemaLockKey})`);
        if (!(await exists(client))) {
          await client.query(`CREATE TABLE IF NOT EXISTS ${T('seq')} (
            ym          TEXT PRIMARY KEY,
            last_serial INTEGER NOT NULL,
            updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
          )`);
          await client.query(`CREATE TABLE IF NOT EXISTS ${T('issued')} (
            quote_no  TEXT PRIMARY KEY,
            ym        TEXT NOT NULL,
            serial    INTEGER NOT NULL,
            issued_to TEXT,
            issued_at TIMESTAMPTZ NOT NULL DEFAULT now()
          )`);
          await client.query(`ALTER TABLE ${T('seq')} ENABLE ROW LEVEL SECURITY`);
          await client.query(`ALTER TABLE ${T('issued')} ENABLE ROW LEVEL SECURITY`);
        }
        await client.query('COMMIT');
      } catch (e) {
        broken = true;
        try { await withTimeout(client.query('ROLLBACK'), commitTimeoutMs); broken = false; } catch (_) { /* 連線可能已壞 */ }
        throw e;
      } finally {
        try { client.release(broken ? true : undefined); } catch (_) { /* ignore */ }
      }
    })().catch((e) => { _schemaReady = null; throw e; });
    return _schemaReady;
  }

  /**
   * 取得全域寫入鎖。成功回傳 { release():Promise<void>, forced:boolean, waitedMs }；
   * 等太久丟 WriteLockError('LOCK_TIMEOUT')。release 可重複呼叫。
   * holdMs>0：持有上限看門狗，到期強制釋放（forced=true），避免單一請求拖死所有寫入者。
   */
  async function acquireDb(opts) {
    opts = opts || {};
    const waitMs = posInt(opts.waitMs, 8000);
    const holdMs = posInt(opts.holdMs, 0);
    const t0 = Date.now();
    const pool = getPool();
    let client;
    try {
      client = await pool.connect();
    } catch (e) {
      throw new WriteLockError('LOCK_UNAVAILABLE', '無法取得資料庫連線', e);
    }
    let lost = false;
    const onErr = () => { lost = true; };
    if (client.on) client.on('error', onErr);
    const dispose = (destroy) => {
      try { if (client.removeListener) client.removeListener('error', onErr); } catch (_) { /* ignore */ }
      if (destroy) {
        // 黑洞連線：pg 的 client.end() 是優雅關閉，對半開 TCP 無效 → 同時直接 destroy 底層 socket，讓伺服器端立刻看到連線斷線並釋放 advisory 鎖
        try { if (client.connection && client.connection.stream && typeof client.connection.stream.destroy === 'function') client.connection.stream.destroy(); } catch (_) { /* ignore */ }
      }
      try { client.release(destroy ? true : undefined); } catch (_) { /* ignore */ }
    };
    // 連線池等待時間也算進等鎖預算：lock_timeout 只給剩下的時間
    const lockMs = Math.max(1, waitMs - (Date.now() - t0));
    // 鎖交易的閒置保底：不超過「持有上限＋3 秒」，也不超過 idleTxMs（持有上限比 DB 保底長時，DB 會先把鎖收回）
    const idleMs = holdMs > 0 ? Math.min(idleTxMs, holdMs + 3000) : idleTxMs;
    try {
      // 一次往返送出（簡單查詢協定、無參數）：少 3 次網路往返＝鎖的持有時間更短。
      // SET LOCAL 不支援參數化；lockMs／idleMs 已強制為整數、lockKey 已驗證為整數字串，所以內嵌安全。
      // 用戶端另有逾時（lock_timeout＋緩衝）：連線變黑洞時 BEGIN 不會永遠卡住，行程內互斥才有機會釋放
      await withTimeout(client.query(SQL.lockBegin(lockMs, idleMs, lockKey)), lockMs + beginSlackMs);
    } catch (e) {
      let healthy = false;
      if (e && e.code !== 'CLIENT_TIMEOUT') {
        try { await withTimeout(client.query('ROLLBACK'), commitTimeoutMs); healthy = true; } catch (_) { /* 連線可能已壞 */ }
      }
      dispose(!healthy);   // ROLLBACK 成功（例如 55P03 等鎖逾時）＝連線仍健康，歸還；否則銷毀
      if (e && (e.code === '55P03' || e.code === 'CLIENT_TIMEOUT')) {
        throw new WriteLockError('LOCK_TIMEOUT', '等待寫入鎖逾時 ' + waitMs + 'ms' + (e.code === 'CLIENT_TIMEOUT' ? '（用戶端逾時）' : ''), e);
      }
      throw new WriteLockError('LOCK_UNAVAILABLE', '取得寫入鎖失敗：' + (e && e.message), e);
    }
    let released = false;
    let timer = null;
    const handle = {
      forced: false,
      waitedMs: Date.now() - t0,
      heldSince: Date.now(),
      async release() {
        if (released) return;
        released = true;
        if (timer) { clearTimeout(timer); timer = null; }
        let broken = lost;
        // COMMIT 有用戶端逾時：連線變黑洞時不會永遠卡住；逾時就銷毀連線（伺服器端連線一斷，交易結束、advisory 鎖自動釋放）
        try { await withTimeout(client.query('COMMIT'), commitTimeoutMs); } catch (_) { broken = true; }
        dispose(broken);
      },
    };
    if (holdMs > 0) {
      timer = setTimeout(() => {
        if (released) return;
        handle.forced = true;
        if (typeof o.onForcedRelease === 'function') { try { o.onForcedRelease(Date.now() - handle.heldSince); } catch (_) { /* ignore */ } }
        handle.release().catch(() => {});
      }, holdMs);
      if (timer.unref) timer.unref();
    }
    return handle;
  }

  // 同一個 guard（＝同一個行程／連線池）一次只讓一個人去等 DB 鎖：否則一群等鎖者會占滿本池的連線，
  // 持有者自己做「讀-改-寫」要查詢時拿不到連線 → 分散式死結（驗證腳本在假 PG 上實際撞到過）。
  let _localHeld = false;
  const _localQ = [];
  function localAcquire(waitMs) {
    return new Promise((resolve, reject) => {
      const grant = () => {
        _localHeld = true;
        let done = false;
        resolve(() => { if (done) return; done = true; _localHeld = false; const n = _localQ.shift(); if (n) { clearTimeout(n.timer); n.grant(); } });
      };
      if (!_localHeld) return grant();
      const w = { grant, timer: null };
      w.timer = setTimeout(() => {
        const i = _localQ.indexOf(w);
        if (i >= 0) _localQ.splice(i, 1);
        reject(new WriteLockError('LOCK_TIMEOUT', '等待寫入鎖逾時 ' + waitMs + 'ms（本行程內排隊）'));
      }, waitMs);
      _localQ.push(w);
    });
  }
  async function acquireLock(opts) {
    opts = opts || {};
    const waitMs = posInt(opts.waitMs, 8000);
    const t0 = Date.now();
    const releaseLocal = await localAcquire(waitMs);
    let handle;
    try { handle = await acquireDb(Object.assign({}, opts, { waitMs: Math.max(1, waitMs - (Date.now() - t0)) })); }
    catch (e) { releaseLocal(); throw e; }
    const dbRelease = handle.release;
    // 先同步放掉行程內互斥、再等 COMMIT：就算 COMMIT 變慢或卡住，下一位也不必被行程內互斥擋住（它仍要等 DB 鎖真的放掉才拿得到）
    handle.release = async function () { releaseLocal(); await dbRelease(); };
    handle.waitedMs = Date.now() - t0;
    return handle;
  }

  /**
   * 原子取號。floor＝呼叫端從 blob 算出的「已用過的最大序號」下界（見 lib/quoteNo.js）。
   * 回傳 { serial, quoteNo }。不依賴 blob 快取、不吃鎖；號碼一旦回傳就永不再發。
   */
  async function nextQuoteSeq(ym, floor, issuedTo) {
    if (!/^\d{6}$/.test(String(ym))) throw new Error('ym 必須是 YYYYMM');
    const fl = Math.max(0, Math.floor(Number(floor)) || 0);
    await ensureSchema();
    const { rows } = await getPool().query(
      SQL.nextQuoteSeq(T, ident(tables.seq)),
      [String(ym), fl, issuedTo == null ? null : String(issuedTo)]
    );
    if (!rows.length) throw new Error('取號沒有回傳結果');
    return { serial: Number(rows[0].serial), quoteNo: rows[0].quote_no };
  }

  /** 已發號紀錄（可選 ym 篩選），供管理員對帳「已發出但不在報價單／流水帳」的號碼 */
  async function listIssued(ym) {
    await ensureSchema();
    const params = [];
    let where = '';
    if (ym) { params.push(String(ym)); where = ' WHERE ym = $1::text'; }
    const { rows } = await getPool().query(
      `SELECT quote_no, ym, serial, issued_to, issued_at FROM ${T('issued')}${where} ORDER BY ym, serial`, params);
    return rows.map((r) => ({ quoteNo: r.quote_no, ym: r.ym, serial: Number(r.serial), issuedTo: r.issued_to, issuedAt: r.issued_at }));
  }

  return { ensureSchema, acquireLock, nextQuoteSeq, listIssued, lockKey, tableName: T };
}

module.exports = { createPgWriteGuard, WriteLockError, DEFAULT_LOCK_KEY, SCHEMA_LOCK_KEY, SQL, withTimeout };
