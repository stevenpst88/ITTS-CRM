'use strict';
// 版本資訊頁用：由「真實執行環境」推導部署環境／資料庫型態，取代寫死的「On-Premise／JSON File Database」。
// 純函式（只讀傳入的 env 物件）。輸出只含列舉值／布林／固定文案——絕不含連線字串、主機名、金鑰。

const pick = (v, allowed, dflt) => (allowed.indexOf(v) >= 0 ? v : dflt);

function deriveRuntimeInfo(env) {
  env = env || {};
  const dbBackend      = env.DB_BACKEND === 'postgres' ? 'postgres' : 'json';
  const storageBackend = pick(env.STORAGE_BACKEND, ['local', 'supabase'], 'local');
  const sessionBackend = pick(env.SESSION_BACKEND, ['jwt', 'memory'], 'jwt');
  const onVercel = !!env.VERCEL;
  // VERCEL_ENV 只會是 production / preview / development；其他值一律不回顯（避免把任意字串帶進畫面）
  const vercelEnv = onVercel ? pick(env.VERCEL_ENV, ['production', 'preview', 'development'], 'unknown') : '';

  const deploy = onVercel
    ? { kind: 'vercel', vercelEnv, label: 'Vercel（' + vercelEnv + '）' }
    : { kind: 'local', vercelEnv: '', label: '本機' };

  const database = dbBackend === 'postgres'
    ? {
        type: 'PostgreSQL（Supabase）',
        engine: 'PostgreSQL，資料以 JSONB 存於 app_data 表',
        file: 'app_data 表（JSONB）',
        authFile: '_auth（雲端帳號，同一資料庫）',
        note: '雲端資料庫，由 Supabase 託管；伺服器本身不存放資料檔',
      }
    : {
        type: 'JSON 檔案（本機開發）',
        engine: '自製輕量 JSON 資料庫（db.js）',
        file: 'data.json',
        authFile: 'auth.json',
        note: '無需額外資料庫伺服器，資料以 JSON 格式儲存於本機磁碟',
      };

  return { deploy, dbBackend, storageBackend, sessionBackend, database };
}

module.exports = { deriveRuntimeInfo };
