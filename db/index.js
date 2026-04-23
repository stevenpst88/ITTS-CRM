// ════════════════════════════════════════════════════════════
//  DB Router — 依環境變數決定用哪個後端
//
//  DB_BACKEND=json      → ./json.js（本地開發用）
//  DB_BACKEND=postgres  → ./postgres.js（Supabase 正式用）
//
//  預設 json，保證舊行為不變。
// ════════════════════════════════════════════════════════════
const backend = (process.env.DB_BACKEND || 'json').toLowerCase();

let impl;
if (backend === 'postgres') {
  impl = require('./postgres');
  console.log('[db] 使用 Postgres 後端（Supabase）');
} else {
  impl = require('./json');
  console.log('[db] 使用 JSON 檔案後端（本地）');
}

module.exports = {
  load: impl.load,
  save: impl.save,
  // Postgres 特有：啟動時 preload。若是 JSON 則為 no-op
  ready: impl.ready || (async () => {}),
};
