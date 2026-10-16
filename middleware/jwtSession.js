// ════════════════════════════════════════════════════════════
//  JWT Session Middleware
//
//  目的：在 Vercel Serverless 上取代 express-session
//  策略：用 JWT 存在 httpOnly cookie，stateless
//
//  對 server.js 完全透明 —— 提供相同的 req.session API：
//    req.session.user             （讀）
//    req.session.user = {...}     （寫 → 自動簽新 JWT 存 cookie）
//    req.session.destroy(cb)      （清除 cookie）
//
//  靠 Proxy 攔截屬性賦值，讓既有業務邏輯完全不用改。
// ════════════════════════════════════════════════════════════
const jwt = require('jsonwebtoken');

const COOKIE_NAME = 'itts_auth';
const MAX_AGE_MS  = 8 * 60 * 60 * 1000; // 登入後的絕對上限 8 小時（不管有沒有一直在用）

// 閒置逾時（2026-10-16 起）：最後一次「使用者操作」後超過這段時間沒有操作，登入失效。
// 預設 30 分鐘；SESSION_IDLE_MS 只給測試縮短用（下限 5 秒，避免設成 0 把全站鎖死）。
function resolveIdleMs() {
  const n = Number(process.env.SESSION_IDLE_MS);
  return Number.isFinite(n) && n >= 5000 ? Math.floor(n) : 30 * 60 * 1000;
}
const IDLE_MS = resolveIdleMs();
// 背景輪詢不算「使用者操作」：畫面開著沒人碰，鈴鐺每隔一陣子還是會打 poll-bundle，
// 若算活動，閒置登出就永遠不會發生。其他請求（頁面載入、API 操作、/api/session/touch）都算。
const PASSIVE_PATHS = new Set(['/api/poll-bundle']);
// Express 預設不分大小寫、容許結尾斜線：/API/Poll-Bundle/ 一樣會進同一支路由，所以比對前先正規化
function isPassivePath(p) { return PASSIVE_PATHS.has(String(p || '').replace(/\/+$/, '').toLowerCase()); }
// cookie 比閒置視窗多活一小段：讓伺服器能「看到」過期的 token 並回報是閒置（訊息才會寫「閒置逾時」，
// 而不是 cookie 被瀏覽器悄悄刪掉後變成含糊的「請先登入」）。逾時與否仍以 token 內的最後活動時間為準，與這段寬限無關。
const COOKIE_GRACE_MS = 5 * 60 * 1000;
// 有活動時，距離上次簽發超過這段時間才重簽 cookie（省掉每個請求都簽 JWT 與寫 Set-Cookie）
const REFRESH_AFTER_MS = Math.min(30 * 1000, Math.floor(IDLE_MS / 4));
const SECRET      = process.env.JWT_SECRET || process.env.SESSION_SECRET;

if (!SECRET) {
  throw new Error('[jwtSession] 缺少 JWT_SECRET 或 SESSION_SECRET 環境變數');
}
// 安全強度檢查：production 必須 ≥32 字元；development 給警告即可
if (SECRET.length < 32) {
  if (process.env.NODE_ENV === 'production') {
    throw new Error('[jwtSession] production 環境的 JWT_SECRET 必須至少 32 字元（目前長度 ' + SECRET.length + '），可用 `openssl rand -base64 48` 產生');
  } else {
    console.warn('[jwtSession] ⚠️  JWT_SECRET 過短（' + SECRET.length + ' 字元），production 部署前請改長到 32+ 字元');
  }
}

// ── 手動解析 cookie（不依賴 cookie-parser）──
function parseCookies(cookieHeader) {
  const jar = {};
  if (!cookieHeader) return jar;
  cookieHeader.split(/;\s*/).forEach((pair) => {
    const idx = pair.indexOf('=');
    if (idx < 0) return;
    const k = pair.slice(0, idx).trim();
    const v = pair.slice(idx + 1).trim();
    try { jar[k] = decodeURIComponent(v); } catch { jar[k] = v; }
  });
  return jar;
}

// meta.lgn＝登入時間（秒，絕對上限的起點；換發時沿用）、meta.lat＝最後活動時間（秒）。省略＝現在。
// exp 固定是 lgn+8h，所以一直有活動也最多撐 8 小時。
function issueToken(userPayload, meta) {
  const nowS = Math.floor(Date.now() / 1000);
  const lgn = (meta && meta.lgn) || nowS;
  const lat = (meta && meta.lat) || nowS;
  return jwt.sign({ user: userPayload, lgn, lat, exp: lgn + Math.floor(MAX_AGE_MS / 1000) }, SECRET);
}

// cookie 存活時間＝min(閒置視窗＋寬限, 離 8 小時上限還剩多少)
function setCookieOnRes(res, token, lgnS) {
  const remain = lgnS ? (lgnS * 1000 + MAX_AGE_MS - Date.now()) : MAX_AGE_MS;
  res.cookie(COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: 'strict',
    secure:   process.env.NODE_ENV === 'production',
    maxAge:   Math.max(1000, Math.min(IDLE_MS + COOKIE_GRACE_MS, remain)),
    path:     '/',
  });
}

function clearCookieOnRes(res) {
  res.clearCookie(COOKIE_NAME, { path: '/' });
}

// ── 輔助：手動寫 cookie（如果 express.response.cookie 不可用）──
// 保留原生 res.cookie 的情況優先，這裡純做後備

module.exports = function jwtSession(req, res, next) {
  // 1. 解析現有 JWT
  const cookies = parseCookies(req.headers.cookie);
  const token   = cookies[COOKIE_NAME];
  let currentUser = null;
  let lgn = 0;   // 登入時間（秒）；舊版 token 沒有這欄 → 用簽發時間
  if (token) {
    try {
      const decoded = jwt.verify(token, SECRET);
      lgn = decoded.lgn || decoded.iat || 0;
      const lat = decoded.lat || decoded.iat || 0;   // 舊版 token 沒有 lat → 以簽發時間當最後活動
      if (decoded.user && lat && (Date.now() - lat * 1000) > IDLE_MS) {
        // 閒置逾時：視為未登入，清掉 cookie；標記起來讓 requireAuth 回應「是因為閒置」
        req._sessionIdleExpired = true;
        lgn = 0;   // 之後重新登入要以新的時間當 8 小時上限的起點
        clearCookieOnRes(res);
      } else {
        currentUser = decoded.user || null;
        // 使用者操作（非背景輪詢）→ 滑動延長閒置視窗
        if (currentUser && !isPassivePath(req.path) && (Date.now() - lat * 1000) >= REFRESH_AFTER_MS) {
          setCookieOnRes(res, issueToken(currentUser, { lgn, lat: Math.floor(Date.now() / 1000) }), lgn);
        }
      }
    } catch (e) {
      // 無效或過期：忽略，視為未登入
    }
  }

  // 2. 內部狀態物件
  const state = { user: currentUser, lgn };

  // 3. 用 Proxy 模擬 req.session
  req.session = new Proxy(state, {
    get(target, prop) {
      if (prop === 'destroy') {
        return function destroy(cb) {
          target.user = null;
          target.lgn = 0;
          clearCookieOnRes(res);
          if (typeof cb === 'function') cb();
        };
      }
      return target[prop];
    },
    set(target, prop, value) {
      const prevUser = target.user;
      target[prop] = value;
      if (prop === 'user') {
        if (value) {
          // 同一個人改個人資料重簽：沿用原本的登入時間（8 小時上限不能被重置）；全新登入或換成別人 → 以現在為起點
          const sameUser = prevUser && value && prevUser.username === value.username;
          if (!target.lgn || !sameUser) target.lgn = Math.floor(Date.now() / 1000);
          setCookieOnRes(res, issueToken(value, { lgn: target.lgn }), target.lgn);
        } else {
          clearCookieOnRes(res);
        }
      }
      return true;
    },
  });

  next();
};

// 導出 helper 供測試
module.exports.COOKIE_NAME = COOKIE_NAME;
module.exports._issueToken = issueToken;
module.exports.IDLE_MS = IDLE_MS;
module.exports.COOKIE_GRACE_MS = COOKIE_GRACE_MS;
module.exports.PASSIVE_PATHS = PASSIVE_PATHS;
module.exports.isPassivePath = isPassivePath;
