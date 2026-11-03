'use strict';
/**
 * 判斷一個請求是不是「純靜態資產」（GET/HEAD，且路徑對應到 _client 內實際存在的檔案）。
 * 用途：server.js 的 session-refresh 中介層對這類請求不做「帳號不存在／停用／改密碼 → 強制登出」判斷。
 * 原因：靜態請求不會經過 db.ready()（只套 /api/），跨實例時記憶體快取可能是舊的，會把剛建立／剛改密碼的帳號誤踢。
 * 靜態檔本身不含資料；存取仍由後面的 requireAuth 把關。/uploads、/admin.html、其他動態路由、/api/* 都不在此集合內，行為不變。
 * 比對採精確大小寫（比檔案系統寬鬆的情況一律當成「不是靜態檔」＝維持舊行為）。
 */
const fs = require('fs');
const path = require('path');

// 不納入豁免：admin.html 有自己的角色／啟用檢查路由，維持原行為
const EXCLUDE = new Set(['admin.html']);

function listFiles(root) {
  const out = new Set();
  (function walk(dir, rel) {
    let ents;
    try { ents = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of ents) {
      if (e.name.startsWith('.')) continue;
      const r = rel ? rel + '/' + e.name : e.name;
      if (e.isDirectory()) walk(path.join(dir, e.name), r);
      else if (e.isFile() && !EXCLUDE.has(r)) out.add(r);
    }
  })(root, '');
  return out;
}

function create(clientDir) {
  const files = listFiles(clientDir);   // 啟動時建一次
  function isStaticAsset(req) {
    const m = req && req.method;
    if (m !== 'GET' && m !== 'HEAD') return false;
    const raw = String(req.path || '');
    if (raw.charAt(0) !== '/') return false;
    // 編碼過的點／斜線／反斜線／空位元組一律不信任
    if (/%2e|%2f|%5c|%00/i.test(raw)) return false;
    if (/[\\\0]/.test(raw)) return false;
    let dec;
    try { dec = decodeURIComponent(raw); } catch { return false; }
    if (/[\\\0]/.test(dec)) return false;
    const rel = dec.slice(1);
    if (!rel) return false;
    const segs = rel.split('/');
    if (segs.some(s => s === '' || s === '.' || s === '..')) return false;
    return files.has(rel);
  }
  return { isStaticAsset, size: () => files.size };
}

module.exports = { create, listFiles };
