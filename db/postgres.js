// ════════════════════════════════════════════════════════════
//  Postgres（Supabase）實作
//
//  策略：保留 load()/save() 介面完全不變，把整個資料當 JSONB blob
//  存在 app_data 表裡（id='main'）。這樣 server.js 完全不用改。
//
//  優點：最低侵入、業務邏輯完全保留
//  缺點：每次 save 會覆寫整包（小量資料 OK，<10MB 沒問題）
//
//  若未來資料量大需要改為逐表查詢，只需改這個檔的內部實作。
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

// ── 讀取 ──
async function loadAsync() {
  await ensureSchema();
  const pool = getPool();
  const { rows } = await pool.query(`SELECT content FROM app_data WHERE id = 'main'`);
  if (rows.length === 0) {
    // 第一次：建立空資料結構
    const empty = { contacts: [], groups: [] };
    await pool.query(
      `INSERT INTO app_data (id, content) VALUES ('main', $1)`,
      [JSON.stringify(empty)]
    );
    return empty;
  }
  return rows[0].content;
}

// ── 寫入 ──
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

// ════════════════════════════════════════════════════════════
//  相容層：把 async 包成同步介面
//
//  server.js 所有呼叫都是 db.load() / db.save()，不是 await。
//  因此這裡用「同步快取 + 背景寫回」策略：
//
//  - 第一次 load()：啟動時 preload（見 bootstrap.js）
//  - 後續 load()：回傳 in-memory 快照
//  - save()：更新 in-memory + 背景 flush 到 DB
//
//  這樣 server.js 所有業務邏輯完全不用改。
// ════════════════════════════════════════════════════════════

let _cache        = null;
let _writeQueue   = Promise.resolve();
let _lastFetch    = 0;
const REFRESH_TTL = 30 * 1000; // 30 秒內外部寫入（seed / 多 instance）也能被感知

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
  // 序列化寫入（避免併發寫衝突）
  // 用 .catch(()=>null) 確保前次失敗不會卡住下一次寫入
  _writeQueue = _writeQueue
    .catch(() => null)
    .then(() => saveAsync(data));
  // 額外掛一層 log，但不影響 _writeQueue 本身（讓呼叫者能 await 拿到錯誤）
  _writeQueue.catch((err) => console.error('[db/postgres] save failed:', err));
  return _writeQueue;
}

// 等待所有排程中的寫入完成（Vercel serverless 必備：避免 function 在背景寫入完成前被回收）
async function flush() {
  try { await _writeQueue; } catch (e) { /* 已 log，不再拋出 */ }
}

// 啟動時呼叫：preload 資料到快取
// 若快取存在但 TTL 到期，重新拉一次（支援外部工具寫入後自動感知）
// 重要：若有 pending 寫入，先等寫入完成再決定是否 refetch（避免覆蓋本地新資料）
async function ready() {
  await flush(); // 確保任何背景寫入完成，避免 refetch 拉回舊資料
  const now = Date.now();
  if (_cache !== null && now - _lastFetch < REFRESH_TTL) return;
  _cache     = await loadAsync();
  _lastFetch = now;
}

module.exports = { load, save, flush, ready, _loadAsync: loadAsync, _saveAsync: saveAsync };
