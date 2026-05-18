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
    max:                     2,       // serverless 建議 2
    idleTimeoutMillis:       10000,
    connectionTimeoutMillis: 5000,
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
async function loadAsync() {
  await ensureSchema();
  const pool = getPool();
  const { rows } = await pool.query(`SELECT content, updated_at FROM app_data WHERE id = 'main'`);
  if (rows.length === 0) {
    const empty = { contacts: [], groups: [], monthlyBudgets: [] };
    const { rows: ins } = await pool.query(
      `INSERT INTO app_data (id, content) VALUES ('main', $1) RETURNING updated_at`,
      [JSON.stringify(empty)]
    );
    _lastUpdatedAtMs = ins.length ? new Date(ins[0].updated_at).getTime() : Date.now();
    return empty;
  }
  _lastUpdatedAtMs = new Date(rows[0].updated_at).getTime();
  return rows[0].content;
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
  if (rows.length) _lastUpdatedAtMs = new Date(rows[0].updated_at).getTime();
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

let _auditCache        = null;
let _auditWriteQueue   = Promise.resolve();
let _auditLoadPromise  = null; // 同時間只有一次 lazy load 進行

// 30s → 300s：cache 命中率大幅提升，省 80-90% Supabase egress
const REFRESH_TTL = 5 * 60 * 1000;

function load() {
  if (_cache === null) {
    throw new Error(
      '[db/postgres] 尚未初始化。請在啟動時先呼叫 await require("./db").ready()'
    );
  }
  return _cache;
}

function save(data) {
  _cache     = data;
  _lastFetch = Date.now(); // 剛寫完，視同剛 fetch
  _writeQueue = _writeQueue
    .catch(() => null)
    .then(() => saveAsync(data));
  _writeQueue.catch((err) => console.error('[db/postgres] save failed:', err));
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

// ── 等待所有背景寫入完成（含主資料 + audit log）──
async function flush() {
  try { await _writeQueue; }      catch (e) { /* 已 log */ }
  try { await _auditWriteQueue; } catch (e) { /* 已 log */ }
}

// ── 啟動時呼叫：preload 資料、跑一次性 migration ──
async function ready() {
  await flush();
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
  _cache = await loadAsync();
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

module.exports = {
  load, save, flush, ready,
  loadAuditLog, appendAuditLog,
  // 給 admin 工具 / 測試用
  _loadAsync: loadAsync, _saveAsync: saveAsync,
  _loadAuditLogAsync: loadAuditLogAsync, _saveAuditLogAsync: saveAuditLogAsync,
};
