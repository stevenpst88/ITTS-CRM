'use strict';
const fs   = require('fs');
const path = require('path');

// ── 儲存後端選擇 ──────────────────────────────────────────
// Postgres（Supabase）：DB_BACKEND=postgres，使用 app_data 表 id='api-stats'
// 本地 JSON：fallback，寫入 data/api-stats.json
const USE_PG = process.env.DB_BACKEND === 'postgres' && !!process.env.DATABASE_URL;
const STATS_PATH = path.join(__dirname, '..', 'data', 'api-stats.json');

const GEMINI_FEATURES = [
  'admin-ocr-card', 'ocr-card', 'visit-suggest',
  'opp-win-rate', 'contact-summary', 'follow-up-email', 'company-insight'
];

// ── 共用工具 ──────────────────────────────────────────────
function today() {
  return new Date().toISOString().slice(0, 10);
}

function thisMonthPrefix() {
  return new Date().toISOString().slice(0, 7);
}

// ── 本地 JSON 後端 ────────────────────────────────────────
function localLoad() {
  try {
    if (fs.existsSync(STATS_PATH)) {
      return JSON.parse(fs.readFileSync(STATS_PATH, 'utf8'));
    }
  } catch { /* 讀取失敗則重建 */ }
  return { gemini: {}, companyLookup: {}, rateLimitHits: {} };
}

function localSave(stats) {
  try {
    const dir = path.dirname(STATS_PATH);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(STATS_PATH, JSON.stringify(stats, null, 2), 'utf8');
  } catch (e) {
    console.error('[apiMonitor] localSave error:', e.message);
  }
}

// ── Postgres（Supabase）後端 ──────────────────────────────
let _pool = null;
function getPgPool() {
  if (_pool) return _pool;
  const { Pool } = require('pg');
  _pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false },
    max: 2,
    idleTimeoutMillis: 10000,
    connectionTimeoutMillis: 5000,
  });
  _pool.on('error', (err) => console.error('[apiMonitor] pool error:', err));
  return _pool;
}

async function pgLoad() {
  const pool = getPgPool();
  // 確保 app_data 表存在（與 db/postgres.js 共用同一張表）
  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_data (
      id         TEXT PRIMARY KEY,
      content    JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);
  const { rows } = await pool.query(`SELECT content FROM app_data WHERE id = 'api-stats'`);
  return rows.length > 0 ? rows[0].content : { gemini: {}, companyLookup: {}, rateLimitHits: {} };
}

async function pgSave(stats) {
  const pool = getPgPool();
  await pool.query(
    `INSERT INTO app_data (id, content, updated_at)
     VALUES ('api-stats', $1, now())
     ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now()`,
    [JSON.stringify(stats)]
  );
}

// ── In-memory 快取 + 背景寫入佇列 ───────────────────────
let _cache      = null;
let _writeQueue = Promise.resolve();

function getCache() {
  if (_cache === null) return USE_PG ? null : localLoad();
  return _cache;
}

function scheduleWrite(stats) {
  _cache = stats;
  if (USE_PG) {
    _writeQueue = _writeQueue.catch(() => null).then(() => pgSave(stats));
    _writeQueue.catch((e) => console.error('[apiMonitor] pgSave failed:', e.message));
  } else {
    localSave(stats);
  }
}

// ── 初始化（Postgres 模式需要在啟動時呼叫）──────────────
async function ready() {
  if (USE_PG) {
    _cache = await pgLoad();
  } else {
    _cache = localLoad();
  }
}

// ── 確保欄位存在 ──────────────────────────────────────────
function ensureGeminiDay(stats, date) {
  if (!stats.gemini[date]) stats.gemini[date] = {};
  for (const f of GEMINI_FEATURES) {
    if (!stats.gemini[date][f])
      stats.gemini[date][f] = { calls: 0, promptTokens: 0, outputTokens: 0, errors: 0 };
  }
  return stats.gemini[date];
}

function ensureCompanyLookupDay(stats, date) {
  if (!stats.companyLookup[date])
    stats.companyLookup[date] = { calls: 0, gcisSuccess: 0, gcisError: 0, twseTpexSuccess: 0, ddgSuccess: 0, ddgError: 0 };
  return stats.companyLookup[date];
}

function ensureRateLimitDay(stats, date) {
  if (!stats.rateLimitHits[date])
    stats.rateLimitHits[date] = { api: 0, login: 0 };
  return stats.rateLimitHits[date];
}

// ── 公開記錄 API ──────────────────────────────────────────
function recordGemini(feature, meta) {
  const stats = getCache();
  if (!stats) return; // 尚未初始化，跳過（Postgres 冷啟動極短暫期）
  const date   = today();
  const day    = ensureGeminiDay(stats, date);
  const bucket = day[feature] || (day[feature] = { calls: 0, promptTokens: 0, outputTokens: 0, errors: 0 });
  bucket.calls++;
  if (meta === null || meta === undefined) {
    bucket.errors++;
  } else {
    bucket.promptTokens  += meta?.promptTokenCount     ?? 0;
    bucket.outputTokens  += meta?.candidatesTokenCount ?? 0;
  }
  scheduleWrite(stats);
}

function recordCompanyLookup({ gcisSuccess = 0, gcisError = 0, twseTpexSuccess = 0, ddgSuccess = 0, ddgError = 0 } = {}) {
  const stats = getCache();
  if (!stats) return;
  const date = today();
  const day  = ensureCompanyLookupDay(stats, date);
  day.calls++;
  day.gcisSuccess     += gcisSuccess;
  day.gcisError       += gcisError;
  day.twseTpexSuccess += twseTpexSuccess;
  day.ddgSuccess      += ddgSuccess;
  day.ddgError        += ddgError;
  scheduleWrite(stats);
}

function recordRateLimit(type) {
  const stats = getCache();
  if (!stats) return;
  const date = today();
  const day  = ensureRateLimitDay(stats, date);
  if (type === 'login') day.login++;
  else day.api++;
  scheduleWrite(stats);
}

// ── 彙整摘要 ─────────────────────────────────────────────
function sumDays(obj, prefix) {
  return Object.entries(obj)
    .filter(([k]) => k.startsWith(prefix))
    .map(([, v]) => v);
}

function sumGeminiDays(days) {
  const totals = { totalCalls: 0, totalPromptTokens: 0, totalOutputTokens: 0, errors: 0 };
  for (const day of days) {
    for (const f of GEMINI_FEATURES) {
      const b = day[f];
      if (!b) continue;
      totals.totalCalls        += b.calls;
      totals.totalPromptTokens += b.promptTokens;
      totals.totalOutputTokens += b.outputTokens;
      totals.errors            += b.errors;
    }
  }
  return totals;
}

function sumCompanyLookupDays(days) {
  const t = { calls: 0, gcisSuccess: 0, gcisError: 0, twseTpexSuccess: 0, ddgSuccess: 0, ddgError: 0 };
  for (const d of days) { for (const k of Object.keys(t)) t[k] += (d[k] || 0); }
  return t;
}

function sumRateLimitDays(days) {
  const t = { api: 0, login: 0 };
  for (const d of days) { t.api += (d.api || 0); t.login += (d.login || 0); }
  return t;
}

function getSummary() {
  const stats  = getCache() || { gemini: {}, companyLookup: {}, rateLimitHits: {} };
  const date   = today();
  const prefix = thisMonthPrefix();

  const todayGemini = stats.gemini[date] || {};
  const todayGeminiFull = {};
  for (const f of GEMINI_FEATURES) {
    todayGeminiFull[f] = todayGemini[f] || { calls: 0, promptTokens: 0, outputTokens: 0, errors: 0 };
  }

  const monthGeminiDays = sumDays(stats.gemini, prefix);
  const todayGeminiDays = stats.gemini[date] ? [stats.gemini[date]] : [];

  const clToday     = stats.companyLookup[date] || { calls: 0, gcisSuccess: 0, gcisError: 0, twseTpexSuccess: 0, ddgSuccess: 0, ddgError: 0 };
  const clMonthDays = sumDays(stats.companyLookup, prefix);

  const rlToday     = stats.rateLimitHits[date] || { api: 0, login: 0 };
  const rlMonthDays = sumDays(stats.rateLimitHits, prefix);

  return {
    gemini: {
      today:     { ...sumGeminiDays(todayGeminiDays), perFeature: todayGeminiFull },
      thisMonth: sumGeminiDays(monthGeminiDays),
    },
    companyLookup: {
      today:     clToday,
      thisMonth: sumCompanyLookupDays(clMonthDays),
    },
    rateLimitHits: {
      today:     rlToday,
      thisMonth: sumRateLimitDays(rlMonthDays),
    },
  };
}

module.exports = { ready, recordGemini, recordCompanyLookup, recordRateLimit, getSummary };
