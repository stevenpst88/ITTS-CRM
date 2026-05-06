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
const MAX_AGE_MS  = 8 * 60 * 60 * 1000; // 8 小時
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

function issueToken(userPayload) {
  return jwt.sign({ user: userPayload }, SECRET, { expiresIn: '8h' });
}

function setCookieOnRes(res, token) {
  res.cookie(COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: 'strict',
    secure:   process.env.NODE_ENV === 'production',
    maxAge:   MAX_AGE_MS,
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
  if (token) {
    try {
      const decoded = jwt.verify(token, SECRET);
      currentUser = decoded.user || null;
    } catch (e) {
      // 無效或過期：忽略，視為未登入
    }
  }

  // 2. 內部狀態物件
  const state = { user: currentUser };

  // 3. 用 Proxy 模擬 req.session
  req.session = new Proxy(state, {
    get(target, prop) {
      if (prop === 'destroy') {
        return function destroy(cb) {
          target.user = null;
          clearCookieOnRes(res);
          if (typeof cb === 'function') cb();
        };
      }
      return target[prop];
    },
    set(target, prop, value) {
      target[prop] = value;
      if (prop === 'user') {
        if (value) {
          setCookieOnRes(res, issueToken(value));
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
