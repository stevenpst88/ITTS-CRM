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

// ── 讀取主資料 ──
async function loadAsync() {
  await ensureSchema();
  const pool = getPool();
  const { rows } = await pool.query(`SELECT content FROM app_data WHERE id = 'main'`);
  if (rows.length === 0) {
    const empty = { contacts: [], groups: [], monthlyBudgets: [] };
    await pool.query(
      `INSERT INTO app_data (id, content) VALUES ('main', $1)`,
      [JSON.stringify(empty)]
    );
    return empty;
  }
  return rows[0].content;
}

// ── 寫入主資料 ──
async function saveAsync(data) {
  await ensureSchema();
  const pool = getPool();
  await pool.query(
    `INSERT INTO app_data (id, content, updated_at)
     VALUES ('main', $1, now())
     ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now()`,
    [JSON.stringify(data)]
  );
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
//  Audit log 走獨立 cache（_auditCache），避免拖累主資料 egress。
// ════════════════════════════════════════════════════════════

let _cache        = null;
let _writeQueue   = Promise.resolve();
let _lastFetch    = 0;

let _auditCache       = null;
let _auditWriteQueue  = Promise.resolve();
let _auditLastFetch   = 0;

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

// ── Audit log 同步介面 ──
async function _ensureAuditLoaded() {
  const now = Date.now();
  if (_auditCache !== null && now - _auditLastFetch < REFRESH_TTL) return;
  _auditCache = await loadAuditLogAsync();
  _auditLastFetch = now;
}

// 同步讀 audit log（必須先呼叫過一次 ensureAuditLoaded，或 ready() 內預載）
function loadAuditLog() {
  return Array.isArray(_auditCache) ? _auditCache : [];
}

// append 一筆 log（用於 writeLog）
// 注意：呼叫者已透過 ready() 中介層觸發 _ensureAuditLoaded，這裡只做寫入
function appendAuditLog(entry) {
  if (!Array.isArray(_auditCache)) _auditCache = [];
  _auditCache.unshift(entry);
  if (_auditCache.length > 5000) _auditCache.length = 5000;
  _auditLastFetch = Date.now();
  _auditWriteQueue = _auditWriteQueue
    .catch(() => null)
    .then(() => saveAuditLogAsync(_auditCache));
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
    // 主 cache 還新鮮，audit 也預載（若過期）
    await _ensureAuditLoaded();
    return;
  }
  _cache     = await loadAsync();
  _lastFetch = now;

  // ── 一次性 migration：若主資料還有 _auditLog → 搬到獨立 row ──
  if (Array.isArray(_cache._auditLog) && _cache._auditLog.length > 0) {
    try {
      const oldLogs = _cache._auditLog;
      const existingNewLogs = await loadAuditLogAsync();
      // 合併：用 id 去重（若已 migrate 過一次又有殘留則不會重複）
      const seen = new Set(existingNewLogs.map(l => l.id).filter(Boolean));
      const merged = [...existingNewLogs];
      for (const l of oldLogs) {
        if (!l.id || !seen.has(l.id)) merged.push(l);
      }
      // 依 timestamp 倒序（新的在前），與原本 unshift 邏輯一致
      merged.sort((a, b) => {
        const ta = a.timestamp ? Date.parse(a.timestamp) : 0;
        const tb = b.timestamp ? Date.parse(b.timestamp) : 0;
        return tb - ta;
      });
      if (merged.length > 5000) merged.length = 5000;

      // 先寫入新 row（確保資料安全）
      await saveAuditLogAsync(merged);
      _auditCache     = merged;
      _auditLastFetch = Date.now();

      // 再從主資料移除 _auditLog
      delete _cache._auditLog;
      _cache._auditLogMigratedAt = new Date().toISOString();
      await saveAsync(_cache);
      _lastFetch = Date.now();
      console.log('[db/postgres] migration: _auditLog 已移到獨立 row（' + merged.length + ' 筆）');
    } catch (e) {
      console.error('[db/postgres] migration failed (將保留原狀，下次啟動再試):', e);
    }
  } else {
    // 主資料沒 _auditLog → 直接預載 audit row
    await _ensureAuditLoaded();
  }
}

module.exports = {
  load, save, flush, ready,
  loadAuditLog, appendAuditLog,
  // 給 admin 工具 / 測試用
  _loadAsync: loadAsync, _saveAsync: saveAsync,
  _loadAuditLogAsync: loadAuditLogAsync, _saveAuditLogAsync: saveAuditLogAsync,
};
