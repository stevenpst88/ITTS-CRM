// ════════════════════════════════════════════════════════════
//  Storage Router — 檔案上傳後端切換
//
//  STORAGE_BACKEND=local     → ./local.js（開發用，寫本地 uploads/）
//  STORAGE_BACKEND=supabase  → ./supabase.js（正式用，Supabase Storage）
//
//  兩邊都提供相同介面：
//    getMulterStorage()        → 給 multer() 用的 storage engine
//    serveFile(req, res, key)  → 讀檔回應（供 /uploads/:key 路由用）
//    getPublicUrl(key)         → 傳回可被瀏覽器存取的 URL
// ════════════════════════════════════════════════════════════
const backend = (process.env.STORAGE_BACKEND || 'local').toLowerCase();

let impl;
if (backend === 'supabase') {
  impl = require('./supabase');
  console.log('[storage] 使用 Supabase Storage');
} else {
  impl = require('./local');
  console.log('[storage] 使用本地檔案系統');
}

module.exports = impl;
