// ════════════════════════════════════════════════════════════
//  Postgres（Supabase）實作
//
//  策略：主資料 (`id='main'`) 與稽核日誌 (`id='audit_log'`) 拆成兩 row。
//  原本一坨 JSONB 約 337 KB，其中 _auditLog 佔 55%，每次請求都被
//  load() 拖回 Vercel 形成大量 Supabase egress。
//
//  拆出後：
//   - 主資料 row 約 150 KB（聯絡人/商機/拜訪等）
//   - audit log row 獨立，只有 admin 看 /api/admin/logs 才會讀
//   - 99% 的請求只搬 150 KB（省 55%）
//
//  v2 進一步優化（egress 戰術第二輪）：
//   - 主資料 cache 過期時，先做 updated_at stale check（<100 bytes），
//     沒變動就跳過完整 fetch（省 150 KB / 次）
//   - audit log 改 lazy load，不再每次 db.ready() 預載
//     （只有 admin 查 logs 或實際 writeLog 觸發時才拉）
//
//  搭配 REFRESH_TTL 30s → 300s（5 分鐘）：再省 80-90% 重抓次數。
// ════════════════════════════════════════════════════════════
const { Pool } = require('pg');
const { createAuditStore, pgAdapter, legacyReadAdapter, createAuditFacade } = require('../lib/auditStore');
const AuditLegacy = require('../lib/auditLegacy');
const { createPgWriteGuard } = require('../lib/pgWriteGuard');
const path = require('path');

// 寫入競爭防護開關（預設關閉＝與先前行為相同；細節見 lib/pgWriteGuard.js、lib/writeGuard.js）
//   DB_WRITE_SERIALIZE=1  寫入型請求以 advisory 鎖序列化（鎖本身由 lib/writeGuard.js 的 middleware 取得）
//   QUOTE_NO_ATOMIC=1     報價單號改由 DB 原子計數器發放
const SERIALIZE_ON = process.env.DB_WRITE_SERIALIZE === '1';
const QUOTE_NO_ATOMIC_ON = process.env.QUOTE_NO_ATOMIC === '1';

const DATABASE_URL = process.env.DATABASE_URL;
if (!DATABASE_URL) {
  throw new Error('[db/postgres] 缺少 DATABASE_URL 環境變數');
}

// ── 單例 Pool（serverless 中會跨 invocation 重用連線）──
let _pool = null;
function getPool() {
  if (_pool) return _pool;
  _pool = new Pool({
    connectionString:        DATABASE_URL,
    ssl:                     { rejectUnauthorized: false },
    max:                     SERIALIZE_ON ? 3 : 2,   // serverless 建議 2；開寫入鎖時多 1 條（鎖持有者獨占一條連線）
    idleTimeoutMillis:       10000,
    connectionTimeoutMillis: 5000,
    // 開寫入鎖時才加用戶端查詢逾時（黑洞連線不會讓存檔／flush 永遠卡住）；關閉時與原本完全相同
    ...(SERIALIZE_ON ? { query_timeout: Number(process.env.DB_QUERY_TIMEOUT_MS) > 0 ? Math.floor(Number(process.env.DB_QUERY_TIMEOUT_MS)) : 15000 } : {}),
  });
  _pool.on('error', (err) => console.error('[db/postgres] pool error:', err));
  return _pool;
}

// ── 初始化 schema（第一次呼叫自動建表）──
let _schemaReady = false;
async function ensureSchema() {
  if (_schemaReady) return;
  const pool = getPool();
  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_data (
      id         TEXT PRIMARY KEY,
      content    JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);
  _schemaReady = true;
}

// ── 讀取主資料（含 updated_at 用於 stale check）──
// fetchMain：純讀取，不動任何模組狀態，回傳 { content, ms, snap }。snap＝發出 SELECT 當下的存檔世代／進行中存檔數，
// 供 ready()（DB_WRITE_SERIALIZE=1）判斷「這份抓回來的內容有沒有可能比快取舊」。
async function fetchMain() {
  await ensureSchema();
  const pool = getPool();
  const snap = { gen: _saveGen, pending: _pendingSaves };
  const { rows } = await pool.query(`SELECT content, updated_at FROM app_data WHERE id = 'main'`);
  if (rows.length === 0) {
    const empty = { contacts: [], groups: [], monthlyBudgets: [] };
    const { rows: ins } = await pool.query(
      `INSERT INTO app_data (id, content) VALUES ('main', $1) RETURNING updated_at`,
      [JSON.stringify(empty)]
    );
    return { content: empty, ms: ins.length ? new Date(ins[0].updated_at).getTime() : Date.now(), snap };
  }
  return { content: rows[0].content, ms: new Date(rows[0].updated_at).getTime(), snap };
}
async function loadAsync() {
  const r = await fetchMain();
  _lastUpdatedAtMs = r.ms;
  return r.content;
}

// ── 寫入主資料（更新本地 _lastUpdatedAtMs，避免馬上又被 stale check 偵測為過期）──
async function saveAsync(data) {
  await ensureSchema();
  const pool = getPool();
  const { rows } = await pool.query(
    `INSERT INTO app_data (id, content, updated_at)
     VALUES ('main', $1, now())
     ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now()
     RETURNING updated_at`,
    [JSON.stringify(data)]
  );
  if (rows.length) {
    const ms = new Date(rows[0].updated_at).getTime();
    _lastUpdatedAtMs = SERIALIZE_ON ? Math.max(_lastUpdatedAtMs, ms) : ms;   // 開關開啟時只增不減（見 _refetchGuarded）
  }
}

// ── 主資料 stale check：只 SELECT updated_at（<100 bytes），
//    回 true 表示主 row 已被其他 Lambda 寫過，需要重抓 ──
async function _checkIfMainStale() {
  if (!_lastUpdatedAtMs) return true; // 從未抓過
  try {
    const pool = getPool();
    const { rows } = await pool.query(`SELECT updated_at FROM app_data WHERE id = 'main'`);
    if (rows.length === 0) return true;
    const newMs = new Date(rows[0].updated_at).getTime();
    return newMs > _lastUpdatedAtMs;
  } catch (e) {
    // 查詢失敗時保守視為過期 → 走完整 fetch
    console.error('[db/postgres] stale check failed, will refetch full:', e.message);
    return true;
  }
}

// ── 讀取 audit log row ──
async function loadAuditLogAsync() {
  await ensureSchema();
  const pool = getPool();
  const { rows } = await pool.query(`SELECT content FROM app_data WHERE id = 'audit_log'`);
  if (rows.length === 0) return [];
  // content 可能存成 { logs: [...] } 或直接 [...]，做相容處理
  const c = rows[0].content;
  if (Array.isArray(c)) return c;
  if (c && Array.isArray(c.logs)) return c.logs;
  return [];
}

// ── 寫入 audit log row ──
async function saveAuditLogAsync(logs) {
  await ensureSchema();
  const pool = getPool();
  const payload = { logs: Array.isArray(logs) ? logs : [] };
  await pool.query(
    `INSERT INTO app_data (id, content, updated_at)
     VALUES ('audit_log', $1, now())
     ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now()`,
    [JSON.stringify(payload)]
  );
}

// ════════════════════════════════════════════════════════════
//  相容層：把 async 包成同步介面
//
//  - 第一次 load()：啟動時 preload（見 server.js bootstrap）
//  - 後續 load()：回傳 in-memory 快照
//  - save()：更新 in-memory + 背景 flush 到 DB
//
//  Audit log 走獨立 cache（_auditCache），lazy load 避免拖累 egress。
// ════════════════════════════════════════════════════════════

let _cache             = null;
let _writeQueue        = Promise.resolve();
let _lastFetch         = 0;
let _lastUpdatedAtMs   = 0;   // 主 row 的 updated_at（毫秒）
let _saveGen           = 0;   // save() 被呼叫的次數（僅 DB_WRITE_SERIALIZE=1 使用）
let _pendingSaves      = 0;   // 已排入佇列、尚未完成的主資料存檔數（僅 DB_WRITE_SERIALIZE=1 使用）

let _auditCache        = null;
let _auditWriteQueue   = Promise.resolve();
let _auditLoadPromise  = null; // 同時間只有一次 lazy load 進行

// REFRESH_TTL=0：每次 API 請求都先做輕量 stale check（只抓 updated_at，~80 bytes），
// 確認 DB 沒被其他 Vercel 實例改過才用快取，有改才完整重抓。
// 這修正「多實例快取分歧」造成的存了又跳回 / 一下有一下沒有 / 互相蓋資料。
// egress 仍低：stale check 只 80 bytes，150KB 完整重抓僅在資料真有變動時發生。
const REFRESH_TTL = 0;

function load() {
  if (_cache === null) {
    throw new Error(
      '[db/postgres] 尚未初始化。請在啟動時先呼叫 await require("./db").ready()'
    );
  }
  return _cache;
}

let _writeFailures     = 0;     // 主資料存檔失敗累計次數（診斷用；writeGuard 不再用它判斷單一請求的成敗，改看該請求自己的存檔 promise）
let _saveHook          = null;  // writeGuard 注入：每次 save() 把「這次存檔的 promise」交給它（歸屬到發起的請求、偵測鎖外寫入）

function save(data) {
  _cache     = data;
  _lastFetch = Date.now(); // 剛寫完，視同剛 fetch
  _writeQueue = _writeQueue
    .catch(() => null)
    .then(() => saveAsync(data));
  _writeQueue.catch((err) => { _writeFailures++; console.error('[db/postgres] save failed:', err); });
  if (SERIALIZE_ON) {
    // 世代號／進行中計數：ready() 據此丟棄「在存檔期間抓回來的舊內容」（審查 #1）
    _saveGen++; _pendingSaves++;
    const dec = () => { _pendingSaves--; };
    _writeQueue.then(dec, dec);
    if (_saveHook) { try { _saveHook(_writeQueue); } catch (e) { console.error('[db] saveHook error:', e && e.message); } }
  }
  return _writeQueue;
}

// ── Audit log lazy load（避免每次 db.ready() 都把 ~190 KB 拉回）──
function _ensureAuditCacheLoaded() {
  if (_auditCache !== null) return Promise.resolve(_auditCache);
  if (!_auditLoadPromise) {
    _auditLoadPromise = loadAuditLogAsync()
      .then(d => {
        _auditCache = Array.isArray(d) ? d : [];
        _auditLoadPromise = null;
        return _auditCache;
      })
      .catch(err => {
        _auditLoadPromise = null;
        console.error('[db/postgres] audit lazy load failed:', err);
        // 失敗也回空陣列，避免阻塞後續流程
        _auditCache = [];
        return _auditCache;
      });
  }
  return _auditLoadPromise;
}

// 讀 audit log（async：第一次呼叫才從 DB 拉，後續用 in-memory）
async function loadAuditLog() {
  await _ensureAuditCacheLoaded();
  return Array.isArray(_auditCache) ? _auditCache : [];
}

// append 一筆 log（用於 writeLog）
// 整個操作（lazy load + 修改 + 寫回）串在 _auditWriteQueue 序列化執行，
// 保證即使多筆連續寫入也不會打亂順序或丟資料
function appendAuditLog(entry) {
  _auditWriteQueue = _auditWriteQueue
    .catch(() => null)
    .then(async () => {
      await _ensureAuditCacheLoaded();
      _auditCache.unshift(entry);
      if (_auditCache.length > 5000) _auditCache.length = 5000;
      await saveAuditLogAsync(_auditCache);
    });
  _auditWriteQueue.catch((err) => console.error('[db/postgres] audit save failed:', err));
  return _auditWriteQueue;
}

// ── 稽核事件（只增不改）：表 audit_events，單列 INSERT，不經快取、不讀-改-寫 ──
// 舊的整包 JSONB（app_data.id='audit_log'）保留為唯讀歷史；appendAuditLog 仍保留（相容）但 server.js 已不再呼叫。
// 安全開關 AUDIT_EVENTS_TABLE：沒設成 '1'（預設）＝第二波之前的舊行為（整包 JSONB／contact_audit.json），完全不碰 audit_events；
// 設成 '1'＝新表（先自我檢查，失敗自動退回舊路徑）。見 lib/auditStore.js createAuditFacade。
const AUDIT_TABLE_ON = process.env.AUDIT_EVENTS_TABLE === '1';
const auditStore = createAuditFacade({
  enabled: AUDIT_TABLE_ON,
  tableStore: createAuditStore(pgAdapter({ query: (sql, params) => getPool().query(sql, params) })),
  legacyStore: createAuditStore(legacyReadAdapter({ readLogs: () => loadAuditLog(), legacyContactFile: path.join(__dirname, '..', 'contact_audit.json') })),
  legacyWrite: (ev) => AuditLegacy.writePgMode(ev, appendAuditLog, path.join(__dirname, '..')),
});
function insertAuditEvent(ev) { return auditStore.insert(ev); }

// ── 等待所有背景寫入完成（含主資料 + audit log）──
async function flush() {
  try { await _writeQueue; }      catch (e) { /* 已 log */ }
  try { await _auditWriteQueue; } catch (e) { /* 已 log */ }
  try { await auditStore.flush(); } catch (e) { /* 已 log */ }
}

// ── 重抓主資料（只在 DB_WRITE_SERIALIZE=1 時使用）──
// 問題（審查 #1）：同實例上「無鎖」請求的 ready() 發出 SELECT 後，持鎖寫入者 W 才 save()；SELECT 在 W 的 UPDATE 提交前就在資料庫端執行完，
// 回應卻在 W 的 save() 之後才到 → 直接指派會把 W 剛改的快取換回舊物件，而 W 的 saveAsync 之後又把時間戳設成最新，
// 之後 stale check 永遠判定「沒過期」，下一個寫入者就在舊內容上改、整份蓋掉 W。
// 解法：SELECT 發出時記下 { 存檔世代, 進行中存檔數 }；回來時若世代變了、發出時或回來時有存檔在進行，就丟棄這份內容（快取已是較新的世代），
// 並先等佇列清空再重抓（最多 3 次）。另外，兩個並行的重抓若「後發的先到、先發的後到」，舊內容不可蓋掉較新的快取（比較時間戳）。
// _lastUpdatedAtMs 只增不減。回傳 true＝快取已是這份內容或更新，false＝放棄（快取維持現狀）。
async function _refetchGuarded() {
  for (let attempt = 0; attempt < 3; attempt++) {
    if (_pendingSaves > 0) { try { await _writeQueue; } catch (e) { /* 失敗由 save() 的 catch 記錄 */ } }
    const r = await fetchMain();
    if (r.snap.pending === 0 && _pendingSaves === 0 && r.snap.gen === _saveGen) {
      // 同實例上兩個並行的重抓可能「後執行的先回來」：舊的那份最後才到，不可把較新的快取換回舊的。
      // 只採用比目前快取時間戳更新的內容；一樣新或更舊＝快取已經是這份（或更新），不動。
      if (r.ms > _lastUpdatedAtMs) { _cache = r.content; _lastUpdatedAtMs = r.ms; }
      return true;
    }
  }
  console.warn('[db/postgres] 重抓主資料時一直有存檔在進行，放棄這次重抓（沿用目前快取）');
  return false;
}

// ── 啟動時呼叫：preload 資料、跑一次性 migration ──
async function ready() {
  await flush();
  if (AUDIT_TABLE_ON) await auditStore.init();   // 只第一次真的檢查（記憶化）；失敗只會退回舊路徑，不丟例外
  const now = Date.now();
  if (_cache !== null && now - _lastFetch < REFRESH_TTL) {
    // 主 cache 還在 TTL 內 → 直接用，不做 audit 預載（lazy）
    return;
  }

  // 主 cache 過期或從未載入
  if (_cache !== null) {
    // 已有 cache，先 stale check（只回 timestamp，~80 bytes 比 150 KB 划算很多）
    const stale = await _checkIfMainStale();
    if (!stale) {
      _lastFetch = now;
      return;
    }
  }

  // 確認需要重抓 → 完整 fetch
  if (SERIALIZE_ON) {
    // 開關開啟：抓回來的內容若可能比快取舊（抓取期間有人 save()），就丟棄，不動 _cache／_lastUpdatedAtMs
    if (!(await _refetchGuarded())) return;
  } else {
    _cache = await loadAsync();
  }
  _lastFetch = now;

  // ── 一次性 migration：若主資料還有 _auditLog → 搬到獨立 row ──
  if (Array.isArray(_cache._auditLog) && _cache._auditLog.length > 0) {
    try {
      const oldLogs = _cache._auditLog;
      const existingNewLogs = await loadAuditLogAsync();
      const seen = new Set(existingNewLogs.map(l => l.id).filter(Boolean));
      const merged = [...existingNewLogs];
      for (const l of oldLogs) {
        if (!l.id || !seen.has(l.id)) merged.push(l);
      }
      merged.sort((a, b) => {
        const ta = a.timestamp ? Date.parse(a.timestamp) : 0;
        const tb = b.timestamp ? Date.parse(b.timestamp) : 0;
        return tb - ta;
      });
      if (merged.length > 5000) merged.length = 5000;

      await saveAuditLogAsync(merged);
      _auditCache = merged;

      delete _cache._auditLog;
      _cache._auditLogMigratedAt = new Date().toISOString();
      await saveAsync(_cache);
      _lastFetch = Date.now();
      console.log('[db/postgres] migration: _auditLog 已移到獨立 row（' + merged.length + ' 筆）');
    } catch (e) {
      console.error('[db/postgres] migration failed (將保留原狀，下次啟動再試):', e);
    }
  }
  // ↑ 移除原本「else 分支自動預載 audit」的邏輯，
  //   改為 lazy load（第一次 admin 看 logs 或 writeLog 時才拉）
}

// ── 寫入競爭防護（lib/pgWriteGuard.js）：表／鎖鍵用預設值；只有開關開啟時才會被呼叫（才會建表）──
const pgGuard = createPgWriteGuard({
  getPool,
  onForcedRelease: (ms) => console.error('[db/postgres] 寫入鎖持有 ' + ms + 'ms 後被強制釋放'),
});
/** 取得 DB 層全域寫入鎖（專用連線上的交易）。waitMs 等鎖上限；holdMs 持有上限看門狗 */
function acquireDbWriteLock(opts) { return pgGuard.acquireLock(opts); }
/** 原子取號：floor＝blob 現況已用最大序號。QUOTE_NO_ATOMIC 沒開時直接拒絕，避免誤用 */
function nextQuoteSeq(ym, floor, issuedTo) {
  if (!QUOTE_NO_ATOMIC_ON) return Promise.reject(new Error('QUOTE_NO_ATOMIC 未啟用'));
  return pgGuard.nextQuoteSeq(ym, floor, issuedTo);
}
function listIssuedQuoteNos(ym) { return pgGuard.listIssued(ym); }
function writeFailureCount() { return _writeFailures; }
/** 只等主資料那一條寫入佇列（不等稽核）：寫入鎖在主資料提交後就可以放，稽核另一列、另一條佇列 */
async function flushMain() { try { await _writeQueue; } catch (e) { /* 已 log；失敗由 writeFailureCount 偵測 */ } }
/** 存檔失敗後讓下個請求整份重抓（stale check 在 _lastUpdatedAtMs=0 時一律回 true） */
function invalidateCache() { _lastUpdatedAtMs = 0; }
function setSaveHook(fn) { _saveHook = typeof fn === 'function' ? fn : null; }

module.exports = {
  load, save, flush, ready,
  loadAuditLog, appendAuditLog,
  auditStore, insertAuditEvent,
  acquireDbWriteLock, nextQuoteSeq, listIssuedQuoteNos, writeFailureCount, flushMain, invalidateCache, setSaveHook,
  _setPoolForTest: (p) => { _pool = p; },
  // 給 admin 工具 / 測試用
  _loadAsync: loadAsync, _saveAsync: saveAsync,
  _loadAuditLogAsync: loadAuditLogAsync, _saveAuditLogAsync: saveAuditLogAsync,
};
