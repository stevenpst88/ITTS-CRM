require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');
const XLSX = require('xlsx');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const db = require('./db');
const jwtSession = require('./middleware/jwtSession');
const storage = require('./storage');
const gemini = require('./ai/gemini');
const apiMonitor = require('./lib/apiMonitor');

const app = express();
const PORT = process.env.PORT || 3000;

// ── 信任 Vercel 代理（取得正確的 req.ip / HTTPS 判斷）──
if (process.env.NODE_ENV === 'production') {
  app.set('trust proxy', 1);
}

// ── Postgres backend：每次請求都呼叫 ready()（TTL 內會直接命中快取）──
if (process.env.DB_BACKEND === 'postgres') {
  app.use(async (req, res, next) => {
    try {
      await db.ready();
      next();
    } catch (err) {
      console.error('[db] ready middleware error:', err);
      res.status(503).json({ error: '資料庫初始化失敗，請稍後再試' });
    }
  });
}

// ── Session 設定 ────────────────────────────────────────
// Vercel Serverless 無持久記憶體：改用 JWT + httpOnly cookie
// 本地開發可用舊 express-session（若 SESSION_BACKEND=memory）
const SESSION_BACKEND = process.env.SESSION_BACKEND || 'jwt';

if (SESSION_BACKEND === 'memory') {
  // 本地/傳統伺服器用
  app.use(session({
    secret: process.env.SESSION_SECRET || (() => { throw new Error('SESSION_SECRET 未設定，請建立 .env 檔案'); })(),
    resave: false,
    saveUninitialized: false,
    cookie: {
      maxAge: 8 * 60 * 60 * 1000,
      httpOnly: true,
      sameSite: 'strict',
      secure: process.env.NODE_ENV === 'production'
    }
  }));
} else {
  // Serverless / Vercel 用（推薦）
  app.use(jwtSession);
}

// ── 安全 HTTP Headers (helmet) ────────────────────────────
app.use(helmet({
  // Content Security Policy：useDefaults:false 完全自訂，避免 Helmet v8 預設
  // 加入 script-src-attr 'none' 封鎖 onclick= 等 HTML 內嵌事件處理器
  contentSecurityPolicy: {
    useDefaults: false,
    directives: {
      "default-src":      ["'self'"],
      "script-src":       ["'self'", "'unsafe-inline'", "https://cdn.jsdelivr.net"],
      "script-src-attr":  ["'unsafe-inline'"],   // 允許 onclick/onchange 等行內事件
      "style-src":        ["'self'", "'unsafe-inline'"],
      "img-src":          ["'self'", "data:", "blob:"],
      "connect-src":      ["'self'"],
      "font-src":         ["'self'"],
      "object-src":       ["'none'"],
      "frame-src":        ["'none'"],
      "frame-ancestors":  ["'none'"],
      "base-uri":         ["'self'"],
      "form-action":      ["'self'"],
      "upgrade-insecure-requests": []
    }
  },
  crossOriginEmbedderPolicy: false,  // 避免影響現有圖片/資源載入
  xFrameOptions:      { action: 'DENY' },          // 雙重防止 Clickjacking
  xContentTypeOptions: true,                        // 防止 MIME sniffing
  referrerPolicy:     { policy: 'strict-origin-when-cross-origin' }
}));

// ── CORS：限制只允許本機與同網域 ────────────────────────────
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || 'http://localhost:3000')
  .split(',').map(s => s.trim());

app.use(cors({
  origin: (origin, cb) => {
    // 允許無 Origin（server-to-server 或同頁導覽）
    if (!origin) return cb(null, true);
    // 允許明確設定的來源
    if (ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
    // 允許 Vercel 自動注入的 VERCEL_URL（形如 itts-crm.vercel.app）
    if (process.env.VERCEL_URL && origin === `https://${process.env.VERCEL_URL}`) return cb(null, true);
    // 允許 VERCEL_BRANCH_URL / VERCEL_PROJECT_PRODUCTION_URL（多 alias 部署）
    const vercelProd = process.env.VERCEL_PROJECT_PRODUCTION_URL;
    if (vercelProd && origin === `https://${vercelProd}`) return cb(null, true);
    cb(new Error('CORS 不允許此來源'));
  },
  credentials: true
}));

// ── 登入速率限制 ──────────────────────────────────────────
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  handler: (req, res) => {
    apiMonitor.recordRateLimit('login');
    res.status(429).json({ success: false, message: '嘗試次數過多，請 15 分鐘後再試' });
  }
});

// ── API 全域速率限制 ──────────────────────────────────────
const apiLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 300,
  standardHeaders: true,
  legacyHeaders: false,
  skip: req => /\.(css|js|png|svg|ico|woff2?)$/.test(req.path),
  handler: (req, res) => {
    apiMonitor.recordRateLimit('api');
    res.status(429).json({ error: '請求過於頻繁，請稍後再試' });
  }
});

app.use(express.json({ limit: '2mb' }));        // 限制 request body 大小
app.use(express.urlencoded({ extended: false, limit: '2mb' }));

// ── DB 寫入完成保證 middleware ─────────────────────────────
// 修復 bug：在 Vercel serverless，db.save() 是非同步背景寫入，
// 若 function 在寫入完成前被回收，資料會遺失，過幾分鐘 cold start 從 DB 拉回舊資料 → 用戶看到「資料還原」
// 解法：在 POST/PUT/DELETE/PATCH 的 response 送出前，自動 await db.flush()
app.use((req, res, next) => {
  if (!['POST', 'PUT', 'DELETE', 'PATCH'].includes(req.method)) return next();
  const origEnd = res.end.bind(res);
  res.end = function (...args) {
    Promise.resolve(db.flush ? db.flush() : null)
      .catch((e) => console.error('[db.flush] error before response:', e))
      .finally(() => origEnd(...args));
    return res;
  };
  next();
});

// ── API 全域速率限制套用 ──────────────────────────────────
app.use('/api/', apiLimiter);

// ── 輸入清理工具函式 ──────────────────────────────────────
/**
 * 從 body 取出字串，限制最大長度，去除首尾空白
 * @param {any} val
 * @param {number} maxLen
 * @returns {string}
 */
function sanitizeStr(val, maxLen = 500) {
  if (val === null || val === undefined) return '';
  return String(val).trim().slice(0, maxLen);
}
/** 驗證整數範圍，回傳 NaN 代表失敗 */
function sanitizeInt(val, min = -Infinity, max = Infinity) {
  const n = parseInt(val, 10);
  if (isNaN(n) || n < min || n > max) return NaN;
  return n;
}
/** 驗證正浮點數 */
function sanitizePositiveFloat(val) {
  const n = parseFloat(val);
  if (isNaN(n) || n < 0) return NaN;
  return n;
}
/** 簡易 Email 格式驗證 */
function isValidEmail(val) {
  return /^[^\s@]{1,100}@[^\s@]{1,254}$/.test(String(val));
}

// ── 驗證 Middleware ──────────────────────────────────────
function requireAuth(req, res, next) {
  if (req.session && req.session.user) return next();
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: '請先登入' });
  res.redirect('/login.html');
}

// ── Auth helpers ────────────────────────────────────────
// Serverless 相容：當 DB_BACKEND=postgres 時，auth / audit 也存在 app_data.content 裡
const _USE_DB_FOR_META = (process.env.DB_BACKEND === 'postgres');

function requireAdmin(req, res, next) {
  if (!req.session || !req.session.user) return res.status(401).json({ error: '請先登入' });
  const authData = loadAuth();
  const user = authData.users.find(u => u.username === req.session.user.username);
  if (!user || user.role !== 'admin' || user.active === false)
    return res.status(403).json({ error: '無管理者權限' });
  next();
}

function loadAuth() {
  if (_USE_DB_FOR_META) {
    const d = db.load();
    return d._auth || { users: [] };
  }
  return JSON.parse(fs.readFileSync(path.join(__dirname, 'auth.json'), 'utf8'));
}

function saveAuth(data) {
  if (_USE_DB_FOR_META) {
    const d = db.load();
    d._auth = data;
    db.save(d);
    return;
  }
  fs.writeFileSync(path.join(__dirname, 'auth.json'), JSON.stringify(data, null, 2), 'utf8');
}

function writeLog(action, operator, target, detail, req) {
  if (_USE_DB_FOR_META) {
    const d = db.load();
    if (!Array.isArray(d._auditLog)) d._auditLog = [];
    d._auditLog.unshift({
      id: uuidv4(),
      action, operator, target, detail,
      ip:        (req && req.ip) || '',
      userAgent: (req && req.headers && req.headers['user-agent']) || '',
      timestamp: new Date().toISOString(),
    });
    // 保留最近 5000 筆
    if (d._auditLog.length > 5000) d._auditLog.length = 5000;
    db.save(d);
    return;
  }
  const logFile = path.join(__dirname, 'audit.log.json');
  let logs = [];
  try { logs = JSON.parse(fs.readFileSync(logFile, 'utf8')); } catch {}
  logs.unshift({
    id: uuidv4(),
    timestamp: new Date().toISOString(),
    action,      // 'CREATE_USER' | 'UPDATE_USER' | 'DELETE_USER' | 'RESET_PASSWORD' | 'UPDATE_PERMISSION'
    operator,    // who did it
    target,      // affected username
    detail,      // description string
    ip: req ? (req.ip || '') : '',
    userAgent: req ? (req.get('user-agent') || '') : ''
  });
  if (logs.length > 500) logs = logs.slice(0, 500); // keep latest 500
  fs.writeFileSync(logFile, JSON.stringify(logs, null, 2), 'utf8');
}

// ── 聯絡人欄位中文標籤 ───────────────────────────────────
const FIELD_LABELS = {
  name:'姓名', nameEn:'英文名', company:'公司', title:'職稱',
  phone:'電話', mobile:'手機', ext:'分機', email:'Email',
  address:'地址', website:'網站', taxId:'統一編號',
  industry:'產業屬性', opportunityStage:'商機分類',
  isPrimary:'主要聯繫窗口', systemVendor:'系統廠商',
  systemProduct:'系統產品', note:'備註', jobFunction:'職能分類'
};

// ── 聯絡人稽核日誌（寫入 contact_audit.json）────────────
const CONTACT_AUDIT_FILE = path.join(__dirname, 'contact_audit.json');

function writeContactAudit(action, req, target, changes) {
  let logs = [];
  try { logs = JSON.parse(fs.readFileSync(CONTACT_AUDIT_FILE, 'utf8')); } catch {}
  logs.unshift({
    id: uuidv4(),
    timestamp: new Date().toISOString(),
    userId: req.session.user.username,
    userName: req.session.user.displayName || req.session.user.username,
    action,          // CREATE | UPDATE | DELETE | RESTORE | PERMANENT_DELETE
    targetId: target.id,
    targetName: target.name || '',
    targetCompany: target.company || '',
    changes: changes || []
  });
  if (logs.length > 3000) logs = logs.slice(0, 3000);
  try { fs.writeFileSync(CONTACT_AUDIT_FILE, JSON.stringify(logs, null, 2), 'utf8'); } catch {}
}

// ── 欄位白名單（Mass Assignment 防護）────────────────────
function pickFields(obj, fields) {
  return fields.reduce((acc, f) => { if (f in obj) acc[f] = obj[f]; return acc; }, {});
}
const CONTACT_FIELDS   = ['name','nameEn','company','title','phone','mobile','ext','email','address','website','taxId','industry','opportunityStage','isPrimary','isResigned','systemVendor','systemProduct','note','cardImage','jobFunction','customerType','productLine','personalDrink','personalHobbies','personalDiet','personalBirthday','personalMemo'];
const VISIT_FIELDS     = ['contactId','contactName','visitDate','visitType','topic','content','nextAction'];
const OPP_FIELDS       = ['contactId','contactName','company','category','product','amount','expectedDate','description','stage','visitId','achievedDate','grossMarginRate'];
const CONTRACT_FIELDS  = ['contractNo','company','contactName','product','startDate','endDate','renewDate','amount','yearAmounts','tcv','salesPerson','note','type'];
const RECEIVABLE_FIELDS= ['company','contactName','invoiceNo','invoiceDate','dueDate','amount','paidAmount','currency','note','status'];

// ── 網址安全驗證 ──────────────────────────────────────────
function sanitizeUrl(url) {
  if (!url) return '';
  const trimmed = url.trim();
  // 只接受 http:// 或 https:// 開頭，防止 javascript: 注入
  return /^https?:\/\//i.test(trimmed) ? trimmed : '';
}

// ── 公司層級欄位同步（新增/更新名片時自動補齊同公司空白欄位）──
const COMPANY_SYNC_FIELDS = ['phone', 'industry', 'address', 'website', 'systemVendor'];
function syncCompanyFields(data, contact) {
  if (!contact.company) return;
  const siblings = data.contacts.filter(c =>
    !c.deleted && c.owner === contact.owner &&
    c.company === contact.company && c.id !== contact.id
  );
  if (siblings.length === 0) return;
  COMPANY_SYNC_FIELDS.forEach(field => {
    const val = contact[field];
    if (!val) return;          // 來源欄位是空的，無需同步
    siblings.forEach(c => {
      if (!c[field]) c[field] = val;  // 只補空白，不蓋掉已有值
    });
  });
}

// ── 公開路由（不需驗證，必須在 requireAuth 之前）─────────
// Static assets：圖片快取 1 天；HTML/JS/CSS 不快取（確保部署後立即生效）
const STATIC_CACHE      = { maxAge: '1d' };
const STATIC_NO_CACHE   = { maxAge: 0, etag: false, lastModified: false };
app.use('/login.html', express.static(path.join(__dirname, '_client', 'login.html'), STATIC_NO_CACHE));
app.use('/itts-logo.png', express.static(path.join(__dirname, '_client', 'itts-logo.png'), STATIC_CACHE));
app.use('/itts-logo.svg', express.static(path.join(__dirname, '_client', 'itts-logo.svg'), STATIC_CACHE));
// admin.html 已移至受保護路由（需登入 + admin 角色）

app.post('/api/login', loginLimiter, async (req, res) => {
  try {
    const { username, password } = req.body;
    if (!username || !password) return res.status(400).json({ success: false, message: '請輸入帳號與密碼' });
    const authData = loadAuth();
    const user = authData.users.find(u => u.username === username);
    if (!user) return res.status(401).json({ success: false, message: '帳號或密碼錯誤' });
    if (user.active === false) return res.status(403).json({ success: false, message: '帳號已停用，請聯繫管理者' });

    let isValid = false;
    if (user.password && user.password.startsWith('$2b$')) {
      // 已 hash 的密碼：bcrypt 比對
      isValid = await bcrypt.compare(password, user.password);
    } else {
      // 舊明文密碼：比對後自動升級為 hash（首次登入自動遷移）
      isValid = (user.password === password);
      if (isValid) {
        user.password = await bcrypt.hash(password, 12);
        saveAuth(authData);
      }
    }

    if (!isValid) return res.status(401).json({ success: false, message: '帳號或密碼錯誤' });
    req.session.user = { username: user.username, displayName: user.displayName, role: user.role || 'user' };
    writeLog('LOGIN', username, username, `登入成功`, req);
    res.json({ success: true, displayName: user.displayName, role: user.role || 'user' });
  } catch (e) {
    console.error('[LOGIN ERROR]', e);
    res.status(500).json({ success: false, message: '登入服務暫時無法使用，請稍後再試' });
  }
});

app.post('/api/logout', (req, res) => {
  req.session.destroy(() => {
    res.clearCookie('connect.sid');
    res.json({ success: true });
  });
});

// ── 受保護路由（需驗證）─────────────────────────────────
// 伺服器啟動時間戳，用於前端資源版本控制（每次重啟強制瀏覽器重抓）
const BUILD_VERSION = Date.now().toString();

function serveHtmlWithVersion(htmlPath, res) {
  fs.readFile(htmlPath, 'utf8', (err, html) => {
    if (err) { res.status(500).send('Internal Server Error'); return; }
    const versioned = html
      .replace(/href="style\.css"/g,    `href="style.css?v=${BUILD_VERSION}"`)
      .replace(/src="app\.js"/g,        `src="app.js?v=${BUILD_VERSION}"`)
      .replace(/src="quote\.js"/g,      `src="quote.js?v=${BUILD_VERSION}"`)
      .replace(/src="admin\.js"/g,      `src="admin.js?v=${BUILD_VERSION}"`);
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.send(versioned);
  });
}

// 根路徑與 /index.html：動態注入版本號到 JS/CSS 連結
app.get(['/', '/index.html'], requireAuth, (req, res) => {
  serveHtmlWithVersion(path.join(__dirname, '_client', 'index.html'), res);
});

// 操作手冊（登入後可存取）
app.get('/help', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, '_client', 'help.html'));
});
app.get('/help.html', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, '_client', 'help.html'));
});
app.use('/help-img', requireAuth, express.static(path.join(__dirname, '_client', 'help-img'), STATIC_NO_CACHE));

// admin.html：需登入且需 admin 角色，且帳號必須為啟用狀態
app.get('/admin.html', requireAuth, (req, res) => {
  const auth = loadAuth();
  const user = auth.users.find(u => u.username === req.session.user.username);
  if (!user || user.role !== 'admin' || user.active === false) return res.redirect('/');
  serveHtmlWithVersion(path.join(__dirname, '_client', 'admin.html'), res);
});
app.use(requireAuth, express.static(path.join(__dirname, '_client'), STATIC_NO_CACHE));
// ── 上傳檔案路由：改由 storage 模組代理（支援本地/Supabase）──
app.get('/uploads/:key', requireAuth, (req, res, next) => {
  Promise.resolve(storage.serveFile(req, res, req.params.key)).catch(next);
});

app.get('/api/me', requireAuth, (req, res) => {
  res.json(req.session.user);
});

// ── 自助更改密碼（登入者本人）────────────────────────────
app.put('/api/user/password', requireAuth, async (req, res) => {
  try {
    const { currentPassword, newPassword } = req.body;
    if (!currentPassword || !newPassword)
      return res.status(400).json({ error: '請填寫舊密碼與新密碼' });
    if (newPassword.length < 6)
      return res.status(400).json({ error: '新密碼至少需要 6 個字元' });

    const auth = loadAuth();
    const user = auth.users.find(u => u.username === req.session.user.username);
    if (!user) return res.status(404).json({ error: '找不到帳號' });

    const valid = await bcrypt.compare(currentPassword, user.password);
    if (!valid) return res.status(401).json({ error: '舊密碼不正確' });

    user.password = await bcrypt.hash(newPassword, 12);
    saveAuth(auth);
    writeLog('CHANGE_PASSWORD', req.session.user.username, req.session.user.username, '自行更改密碼', req);
    res.json({ success: true });
  } catch (e) {
    console.error('[CHANGE_PASSWORD ERROR]', e);
    res.status(500).json({ error: '更改密碼失敗，請稍後再試' });
  }
});

// ════════════════════════════════════════════════════════
// ── AI 功能（Google Gemini）──────────────────────────────
// ════════════════════════════════════════════════════════

// 共用：未設定 API Key 時的回應
function requireAi(req, res, next) {
  if (!gemini.isConfigured()) return res.status(503).json({ error: 'AI_NOT_CONFIGURED' });
  next();
}

// ── Feature 1：名片 AI 辨識 ─────────────────────────────
const uploadOcr = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /^image\//i.test(file.mimetype);
    ok ? cb(null, true) : cb(new Error('只支援圖片格式'));
  }
});

app.post('/api/admin/ai-ocr-card', requireAdmin, requireAi,
  (req, res, next) => uploadOcr.single('card')(req, res, next),
  async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: '未收到圖片' });
      const base64 = req.file.buffer.toString('base64');
      const mimeType = req.file.mimetype;

      const model = gemini.getModel();
      const result = await model.generateContent([
        {
          inlineData: { mimeType, data: base64 }
        },
        `你是名片 OCR 助手。請從這張名片圖片中擷取所有聯絡資訊。
回傳一個 JSON 陣列，每位聯絡人一個物件，使用以下欄位名稱（沒有的填空字串 ""）：
name（中文姓名）, nameEn（英文姓名）, company（公司名稱）, title（職稱）,
phone（市話，含區碼）, mobile（手機）, ext（分機號碼）, email（Email）,
address（地址）, website（網址，需含 http/https）, taxId（統一編號 8 碼）, industry（產業）

只回傳 JSON 陣列，不要任何說明或 markdown。`
      ]);

      apiMonitor.recordGemini('admin-ocr-card', result.response.usageMetadata);
      const text = result.response.text();
      let contacts;
      try {
        contacts = gemini.parseJson(text);
        if (!Array.isArray(contacts)) contacts = [contacts];
      } catch {
        return res.status(500).json({ error: 'AI 回應格式錯誤，請重試', raw: text.slice(0, 200) });
      }

      res.json({ contacts });
    } catch (e) {
      console.error('[ai-ocr-card admin]', e.message);
      apiMonitor.recordGemini('admin-ocr-card', null);
      const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('quota') || String(e.message).includes('RESOURCE_EXHAUSTED');
      res.status(is429 ? 429 : 500).json({ error: is429 ? 'AI 服務暫時忙碌，請稍後再試' : 'AI 辨識失敗：' + e.message });
    }
  }
);

// ── 業務名片拍照辨識（業務可用，單張填入）────────────────
const uploadOcrUser = multer({ storage: multer.memoryStorage(), limits: { fileSize: 8 * 1024 * 1024 } });
app.post('/api/ai/ocr-card', requireAuth, requireAi,
  (req, res, next) => uploadOcrUser.single('card')(req, res, next),
  async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: '未收到圖片' });
      const base64 = req.file.buffer.toString('base64');
      const mimeType = req.file.mimetype;

      const model = gemini.getModel();
      const result = await model.generateContent([
        { inlineData: { mimeType, data: base64 } },
        `你是名片 OCR 助手。請從這張名片圖片中擷取聯絡資訊。
回傳一個 JSON 物件（單一聯絡人），使用以下欄位名稱（沒有的填空字串 ""）：
name（中文姓名）, nameEn（英文姓名）, company（公司名稱）, title（職稱）,
phone（市話，含區碼）, mobile（手機）, ext（分機號碼）, email（Email）,
address（地址）, website（網址，需含 http/https）, taxId（統一編號 8 碼）, industry（產業）

只回傳 JSON 物件，不要任何說明或 markdown。`
      ]);

      apiMonitor.recordGemini('ocr-card', result.response.usageMetadata);
      const text = result.response.text();
      let contact;
      try {
        const parsed = gemini.parseJson(text);
        contact = Array.isArray(parsed) ? parsed[0] : parsed;
      } catch {
        return res.status(500).json({ error: 'AI 回應格式錯誤，請重試' });
      }

      res.json({ contact });
    } catch (e) {
      console.error('[ai/ocr-card]', e.message);
      apiMonitor.recordGemini('ocr-card', null);
      const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('quota') || String(e.message).includes('RESOURCE_EXHAUSTED');
      res.status(is429 ? 429 : 500).json({ error: is429 ? 'AI 服務暫時忙碌，請稍後再試' : 'AI 辨識失敗：' + e.message });
    }
  }
);

// ── Feature 2：拜訪記錄 AI 建議 ──────────────────────────
app.post('/api/ai/visit-suggest', requireAuth, requireAi, async (req, res) => {
  try {
    const { topic, content, visitType, contactName, company } = req.body;
    if (!content || content.trim().length < 10)
      return res.status(400).json({ error: '請先填寫會談內容（至少 10 字）' });

    const model = gemini.getModel();
    const result = await model.generateContent(
      `你是一位 B2B 業務顧問。以下是一筆拜訪記錄：
客戶：${company || '（未填）'} / ${contactName || '（未填）'}
拜訪方式：${visitType || ''}
主題：${topic || '（未填）'}
內容：${content}

請根據內容：
1. 建議一個具體的「下一步行動」（30 字以內，繁體中文）
2. 列出 2–3 個關鍵重點（每點 20 字以內，繁體中文）

只回傳 JSON，不要其他說明：
{"nextAction":"...","keyTakeaways":["...","..."]}`
    );

    apiMonitor.recordGemini('visit-suggest', result.response.usageMetadata);
    const text = result.response.text();
    try {
      const data = gemini.parseJson(text);
      res.json(data);
    } catch {
      res.status(500).json({ error: 'AI 回應格式錯誤，請重試' });
    }
  } catch (e) {
    apiMonitor.recordGemini('visit-suggest', null);
    const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('RESOURCE_EXHAUSTED'); const status = is429 ? 429 : 500;
    const msg = is429 ? 'AI 服務暫時忙碌，請稍後再試' : ('AI 發生錯誤：' + e.message);
    res.status(status).json({ error: msg });
  }
});

// ── Feature 3：商機贏率預測 ──────────────────────────────
app.post('/api/ai/opp-win-rate', requireAuth, requireAi, async (req, res) => {
  try {
    const { oppId } = req.body;
    if (!oppId) return res.status(400).json({ error: '缺少 oppId' });

    const data   = db.load();
    const viewable = getViewableOwners(req, 'opportunities');
    const opp = (data.opportunities || []).find(o => o.id === oppId && viewable.includes(o.owner));
    if (!opp) return res.status(404).json({ error: '找不到此商機' });

    // 統計拜訪活躍度
    const now = Date.now();
    const visits = (data.visits || []).filter(v => {
      if (opp.contactId && v.contactId === opp.contactId) return true;
      if (opp.company && v.contactName && (data.contacts || []).some(c => c.id === v.contactId && c.company === opp.company)) return true;
      return false;
    });
    const visits30 = visits.filter(v => v.visitDate && (now - new Date(v.visitDate)) < 30*86400000).length;
    const visits60 = visits.filter(v => v.visitDate && (now - new Date(v.visitDate)) < 60*86400000).length;
    const hist = opp.stageHistory || [];
    const promotions = hist.filter(h => ['D','C','B','A'].indexOf(h.to) > ['D','C','B','A'].indexOf(h.from)).length;
    const demotions  = hist.filter(h => ['D','C','B','A'].indexOf(h.to) < ['D','C','B','A'].indexOf(h.from)).length;
    const daysToClose = opp.expectedDate
      ? Math.round((new Date(opp.expectedDate) - new Date()) / 86400000)
      : null;
    const STAGE_LABEL_AI = { D:'D（靜止）', C:'C（Pipeline）', B:'B（Upside）', A:'A（Commit）', Won:'Won' };

    const model = gemini.getModel();
    const result = await model.generateContent(
      `你是 B2B 銷售預測分析師。請根據以下資料預測商機贏率（0–100 整數）：

商機：${opp.company} / ${opp.product || opp.description || '未填'}
現在階段：${STAGE_LABEL_AI[opp.stage] || opp.stage}
金額：${opp.amount || '未填'} 萬元
距預計成交：${daysToClose !== null ? daysToClose + ' 天' : '未設定'}
最近 30 天拜訪：${visits30} 次，60 天：${visits60} 次
歷史晉升：${promotions} 次，退後：${demotions} 次
說明：${opp.description || '（無）'}

只回傳 JSON：{"winRate":72,"reasoning":"20字內說明","factors":{"stage":35,"activity":10,"timeline":20,"amount":7}}
winRate 是 0–100 整數，factors 各項加總約等於 winRate。`
    );

    apiMonitor.recordGemini('opp-win-rate', result.response.usageMetadata);
    const text = result.response.text();
    try {
      const aiData = gemini.parseJson(text);
      // 快取到商機物件
      const idx = data.opportunities.findIndex(o => o.id === oppId);
      if (idx !== -1) {
        data.opportunities[idx].aiWinRate   = aiData.winRate;
        data.opportunities[idx].aiWinRateAt = new Date().toISOString();
        db.save(data);
      }
      res.json(aiData);
    } catch {
      res.status(500).json({ error: 'AI 回應格式錯誤，請重試' });
    }
  } catch (e) {
    apiMonitor.recordGemini('opp-win-rate', null);
    const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('RESOURCE_EXHAUSTED'); const status = is429 ? 429 : 500;
    const msg = is429 ? 'AI 服務暫時忙碌，請稍後再試' : ('AI 發生錯誤：' + e.message);
    res.status(status).json({ error: msg });
  }
});

// ── Feature 4：客戶輪廓 AI 摘要 ──────────────────────────
app.post('/api/ai/contact-summary', requireAuth, requireAi, async (req, res) => {
  try {
    const { contactId } = req.body;
    if (!contactId) return res.status(400).json({ error: '缺少 contactId' });

    const data  = db.load();
    const username = req.session.user.username;
    const viewable = getViewableOwners(req, 'contacts');
    const contact = (data.contacts || []).find(c => c.id === contactId && viewable.includes(c.owner) && !c.deleted);
    if (!contact) return res.status(404).json({ error: '找不到此聯絡人' });

    // 最近 5 筆拜訪
    const visits = (data.visits || [])
      .filter(v => v.contactId === contactId)
      .sort((a, b) => (b.visitDate || '').localeCompare(a.visitDate || ''))
      .slice(0, 5);

    const daysSinceVisit = visits[0]?.visitDate
      ? Math.round((Date.now() - new Date(visits[0].visitDate)) / 86400000) : null;

    const visitSummary = visits.length
      ? visits.map(v => `・${v.visitDate} ${v.visitType}：${v.topic || ''} ${v.content ? '— ' + v.content.slice(0, 40) : ''}`).join('\n')
      : '（尚無拜訪記錄）';

    // 進行中商機
    const opps = (data.opportunities || []).filter(o =>
      o.contactId === contactId && o.stage !== 'D' && o.stage !== 'Won'
    );
    const oppSummary = opps.length
      ? opps.map(o => `${o.company} ${o.product || ''} ${o.stage} $${o.amount || '?'}萬`).join('；')
      : '（無進行中商機）';

    const model = gemini.getModel();
    const result = await model.generateContent(
      `你是一位 B2B 業務關係分析師。請根據以下資料，用繁體中文生成一段 100–150 字的「客戶輪廓摘要」：
1. 先判斷關係健康度（良好 / 普通 / 需關注）
2. 近期拜訪重點摘要
3. 商機現況
4. 一個具體的關係維護建議

客戶：${contact.name || ''}，${contact.title || ''}，${contact.company || ''}
產業：${contact.industry || '（未填）'}
距上次拜訪：${daysSinceVisit !== null ? daysSinceVisit + ' 天' : '未知'}
最近拜訪：
${visitSummary}
進行中商機：${oppSummary}

只回傳 JSON：{"summary":"...","health":"良好"}`
    );

    apiMonitor.recordGemini('contact-summary', result.response.usageMetadata);
    const text = result.response.text();
    try {
      const aiData = gemini.parseJson(text);
      // 快取到聯絡人
      const idx = data.contacts.findIndex(c => c.id === contactId);
      if (idx !== -1) {
        data.contacts[idx].aiSummary       = aiData.summary;
        data.contacts[idx].aiSummaryHealth = aiData.health;
        data.contacts[idx].aiSummaryAt     = new Date().toISOString();
        db.save(data);
      }
      res.json(aiData);
    } catch {
      res.status(500).json({ error: 'AI 回應格式錯誤，請重試' });
    }
  } catch (e) {
    apiMonitor.recordGemini('contact-summary', null);
    const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('RESOURCE_EXHAUSTED'); const status = is429 ? 429 : 500;
    const msg = is429 ? 'AI 服務暫時忙碌，請稍後再試' : ('AI 發生錯誤：' + e.message);
    res.status(status).json({ error: msg });
  }
});

// ── Feature 5：個人化跟進信件草稿 ──────────────────────
app.post('/api/ai/follow-up-email', requireAuth, requireAi, async (req, res) => {
  try {
    const { contactName, company, title, visitType, topic, content, nextAction } = req.body;
    if (!content && !topic)
      return res.status(400).json({ error: '請先填寫拜訪主題或會談內容' });

    const model = gemini.getModel();
    const result = await model.generateContent(
      `你是 B2B 業務助理。根據以下拜訪記錄，以繁體中文撰寫一封專業的後續跟進信件（Email）草稿。
客戶：${company || '（未填）'} / ${contactName || '（未填）'}${title ? `（${title}）` : ''}
拜訪方式：${visitType || ''}，主題：${topic || '（未填）'}
會談內容：${content || '（未填）'}
下一步行動：${nextAction || '（未填）'}

要求：
- 主旨簡短（20 字以內）
- 信件內容 150–250 字，語氣專業且友善
- 開頭問候，摘要本次討論重點，提及後續行動，結尾署名「[您的姓名]」

只回傳 JSON，不要任何說明：{"subject":"...","body":"..."}`
    );

    apiMonitor.recordGemini('follow-up-email', result.response.usageMetadata);
    const text = result.response.text();
    try {
      const data = gemini.parseJson(text);
      res.json(data);
    } catch {
      res.status(500).json({ error: 'AI 回應格式錯誤，請重試' });
    }
  } catch (e) {
    apiMonitor.recordGemini('follow-up-email', null);
    const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('RESOURCE_EXHAUSTED');
    res.status(is429 ? 429 : 500).json({ error: is429 ? 'AI 服務暫時忙碌，請稍後再試' : ('AI 發生錯誤：' + e.message) });
  }
});

// ── Feature 6：AI 公司背景分析（網頁 fetch + Gemini）───
app.post('/api/ai/company-insight', requireAuth, requireAi, async (req, res) => {
  try {
    const { url } = req.body;
    if (!url || !/^https?:\/\//i.test(url))
      return res.status(400).json({ error: '請提供有效的網址（需包含 https://）' });

    // 抓取網頁內容（10 秒逾時）
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 10000);
    let pageText;
    try {
      const resp = await fetch(url, {
        signal: controller.signal,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; CRM-AI-Bot/1.0)' }
      });
      clearTimeout(timer);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const html = await resp.text();
      // 移除 script / style 區塊及 HTML tags
      pageText = html
        .replace(/<script[\s\S]*?<\/script>/gi, '')
        .replace(/<style[\s\S]*?<\/style>/gi, '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim()
        .slice(0, 6000);
    } catch (fetchErr) {
      clearTimeout(timer);
      return res.status(400).json({ error: '無法存取此網址，請確認網址是否正確或網站是否允許存取' });
    }

    if (!pageText || pageText.length < 30)
      return res.status(400).json({ error: '網頁內容過少，無法分析' });

    const model = gemini.getModel();
    const result = await model.generateContent(
      `你是一位頂尖的 B2B Key Account 業務顧問，熟悉 ERP/IT 解決方案銷售。
以下是客戶公司官網的文字內容（已移除 HTML）：
${pageText}

請從 KA 業務視角，對這家公司進行五構面分析。對於無法從官網確認的數據，請誠實標注「⚠️ 資料有限」並給出補充建議，不要捏造數字。
只回傳 JSON，繁體中文，不要任何說明：
{"companyName":"公司全名","analysisBase":"本次分析依據（如：官網首頁、產品頁、徵才頁等，20字內）","strategic":{"signal":"green|yellow|red","marketPosition":"市場定位與競爭態勢觀察（40字內）","industryTrend":"行業趨勢與政策風險推斷（40字內）","growthDriver":"增長動能與創新信號（40字內）","salesHook":"業務切入話題建議（30字內）"},"financial":{"signal":"green|yellow|red","profitability":"獲利能力推斷（40字內，資料不足請標注⚠️）","cashFlow":"現金流與投資傾向推斷（40字內）","capexSignal":"資本支出需求信號（40字內）","salesHook":"財務面切入話題（30字內）"},"operational":{"signal":"green|yellow|red","efficiency":"營運效率與IT成熟度觀察（40字內）","riskExposure":"合規與供應鏈風險點（40字內）","itNeed":"IT/ERP需求痛點推斷（40字內）","salesHook":"營運面切入話題（30字內）"},"humanCapital":{"signal":"green|yellow|red","talentStrategy":"人才策略與組織信號（40字內）","cultureSignal":"企業文化與轉型準備度（40字內）","leadershipSignal":"領導層穩定性觀察（40字內）","salesHook":"人力面切入話題（30字內）"},"customerBrand":{"signal":"green|yellow|red","brandStrength":"品牌影響力與客戶黏性觀察（40字內）","esgSignal":"ESG與社會責任信號（40字內）","loyaltySignal":"客戶忠誠度與口碑推斷（40字內）","salesHook":"品牌面切入話題（30字內）"},"executiveSummary":"給KA業務的整體建議與優先行動（120字內）","topOpportunities":["機會點1（25字內）","機會點2（25字內）","機會點3（25字內）"]}`
    );

    apiMonitor.recordGemini('company-insight', result.response.usageMetadata);
    const text = result.response.text();
    try {
      const data = gemini.parseJson(text);
      res.json(data);
    } catch {
      res.status(500).json({ error: 'AI 回應格式錯誤，請重試' });
    }
  } catch (e) {
    apiMonitor.recordGemini('company-insight', null);
    const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('RESOURCE_EXHAUSTED');
    res.status(is429 ? 429 : 500).json({ error: is429 ? 'AI 服務暫時忙碌，請稍後再試' : ('AI 發生錯誤：' + e.message) });
  }
});

// ── Admin: 批次填入官網（分批分頁，避免 Vercel timeout）──
app.post('/api/admin/bulk-fill-website', requireAdmin, async (req, res) => {
  const BATCH = 15; // 每批處理幾家公司（並行）
  const offset = parseInt(req.body?.offset ?? req.query?.offset ?? 0);

  const data     = db.load();
  const contacts = (data.contacts || []).filter(c => !c.deleted);
  const targets  = contacts.filter(c => !c.website);

  // 聚合：同統編只查一次，同公司名只查一次
  const taxIdGroups = {};
  const nameGroups  = {};
  for (const c of targets) {
    const tid = (c.taxId || '').trim();
    if (tid && /^\d{8}$/.test(tid)) {
      (taxIdGroups[tid] = taxIdGroups[tid] || []).push(c);
    } else {
      const key = (c.company || '').trim();
      if (key) (nameGroups[key] = nameGroups[key] || []).push(c);
    }
  }

  // 合併成任務清單 [ { key, group, taxId, companyName } ]
  const allTasks = [
    ...Object.entries(taxIdGroups).map(([tid, g]) => ({ key: tid,  group: g, taxId: tid,  companyName: g[0].company || '' })),
    ...Object.entries(nameGroups) .map(([nm,  g]) => ({ key: '_'+nm, group: g, taxId: '',   companyName: nm })),
  ];

  const totalCompanies = allTasks.length;
  const batch = allTasks.slice(offset, offset + BATCH);
  const hasMore = offset + BATCH < totalCompanies;

  if (batch.length === 0) {
    return res.json({ success: true, updated: 0, notFound: 0, skipped: 0,
                      totalContacts: targets.length, totalCompanies, offset, hasMore: false });
  }

  // 只在第一批時載入上市/上櫃清單（避免重複下載）
  let lists = { twse: [], tpex: [] };
  try { lists = await getCompanyLists(); } catch {}

  // 並行查詢這批公司的官網
  const results = await Promise.allSettled(
    batch.map(async task => {
      let stockCode = '', listedType = '';
      if (task.taxId) {
        const twseM = (lists.twse || []).find(c => c['營利事業統一編號'] === task.taxId);
        const tpexM = !twseM && (lists.tpex || []).find(c => c['UnifiedBusinessNo.'] === task.taxId);
        if (twseM) { stockCode = twseM['公司代號'] || ''; listedType = '上市'; }
        else if (tpexM) { stockCode = tpexM.SecuritiesCompanyCode || ''; listedType = '上櫃'; }
      }
      const { website, debug } = await fetchWebsiteOnly(task.taxId, task.companyName, stockCode, listedType);
      return { task, website: website || '', debug, stockCode, listedType };
    })
  );

  // 回填資料
  let updated = 0, notFound = 0;
  const debugSamples = [];
  for (const r of results) {
    if (r.status !== 'fulfilled') {
      notFound += r.reason?.task?.group?.length || 1;
      debugSamples.push({ company: '?', err: String(r.reason?.message || r.reason) });
      continue;
    }
    const { task, website, debug, stockCode, listedType } = r.value;
    if (website) {
      task.group.forEach(c => { c.website = website; });
      updated += task.group.length;
    } else {
      notFound += task.group.length;
    }
    if (debugSamples.length < 5) {
      debugSamples.push({
        company: task.companyName, taxId: task.taxId,
        stockCode, listedType, website, debug
      });
    }
  }

  db.save(data);
  res.json({
    success: true, updated, notFound, skipped: 0,
    totalContacts: targets.length, totalCompanies,
    offset, nextOffset: offset + BATCH, hasMore,
    debugSamples
  });
});

// ── Admin: get all users ─────────────────────────────────
app.get('/api/admin/users', requireAdmin, (req, res) => {
  const auth = loadAuth();
  const users = auth.users.map(u => ({
    username: u.username,
    displayName: u.displayName,
    role: u.role || 'user',
    canDownloadContacts: u.canDownloadContacts || false,
    canSetTargets: u.canSetTargets || false,
    active: u.active !== false
  }));
  res.json(users);
});

// ── Admin: create user ───────────────────────────────────
app.post('/api/admin/users', requireAdmin, async (req, res) => {
  const { username, password, displayName, role, canDownloadContacts, canSetTargets } = req.body;
  if (!username || !password) return res.status(400).json({ error: '帳號與密碼為必填' });
  if (password.length < 6) return res.status(400).json({ error: '密碼至少需要 6 個字元' });
  const auth = loadAuth();
  if (auth.users.find(u => u.username === username)) return res.status(400).json({ error: '帳號已存在' });
  const hashedPassword = await bcrypt.hash(password, 12);
  const newUser = {
    username: username.trim(),
    password: hashedPassword,
    displayName: displayName || username,
    role: role || 'user',
    canDownloadContacts: !!canDownloadContacts,
    canSetTargets: !!canSetTargets,
    active: true,
    createdAt: new Date().toISOString()
  };
  auth.users.push(newUser);
  saveAuth(auth);
  writeLog('CREATE_USER', req.session.user.username, username, `新增帳號 ${username}（${newUser.displayName}）`, req);
  res.json({ success: true });
});

// ── Admin: update user ───────────────────────────────────
app.put('/api/admin/users/:username', requireAdmin, (req, res) => {
  const auth = loadAuth();
  const idx = auth.users.findIndex(u => u.username === req.params.username);
  if (idx === -1) return res.status(404).json({ error: '找不到此帳號' });
  const { displayName, role, canDownloadContacts, canSetTargets, active } = req.body;
  if (displayName !== undefined)        auth.users[idx].displayName = displayName;
  if (role !== undefined)               auth.users[idx].role = role;
  if (canDownloadContacts !== undefined) auth.users[idx].canDownloadContacts = !!canDownloadContacts;
  if (canSetTargets !== undefined)       auth.users[idx].canSetTargets = !!canSetTargets;
  if (active !== undefined)             auth.users[idx].active = !!active;
  saveAuth(auth);
  writeLog('UPDATE_USER', req.session.user.username, req.params.username, `更新帳號設定`, req);
  res.json({ success: true });
});

// ── Admin: reset password ────────────────────────────────
app.put('/api/admin/users/:username/password', requireAdmin, async (req, res) => {
  const { password } = req.body;
  if (!password) return res.status(400).json({ error: '新密碼不能為空' });
  if (password.length < 6) return res.status(400).json({ error: '密碼至少需要 6 個字元' });
  const auth = loadAuth();
  const user = auth.users.find(u => u.username === req.params.username);
  if (!user) return res.status(404).json({ error: '找不到此帳號' });
  user.password = await bcrypt.hash(password, 12);
  saveAuth(auth);
  writeLog('RESET_PASSWORD', req.session.user.username, req.params.username, `重設密碼`, req);
  res.json({ success: true });
});

// ── Admin: delete user ───────────────────────────────────
app.delete('/api/admin/users/:username', requireAdmin, (req, res) => {
  const auth = loadAuth();
  const idx = auth.users.findIndex(u => u.username === req.params.username);
  if (idx === -1) return res.status(404).json({ error: '找不到此帳號' });
  if (auth.users[idx].role === 'admin' && auth.users.filter(u => u.role === 'admin').length <= 1)
    return res.status(400).json({ error: '無法刪除最後一位管理者' });
  const deleted = auth.users[idx];
  auth.users.splice(idx, 1);
  saveAuth(auth);
  writeLog('DELETE_USER', req.session.user.username, deleted.username, `刪除帳號 ${deleted.username}`, req);
  res.json({ success: true });
});

// ── Admin: get audit logs ────────────────────────────────
app.get('/api/admin/logs', requireAdmin, (req, res) => {
  if (_USE_DB_FOR_META) {
    const d = db.load();
    return res.json(Array.isArray(d._auditLog) ? d._auditLog : []);
  }
  const logFile = path.join(__dirname, 'audit.log.json');
  try {
    const logs = JSON.parse(fs.readFileSync(logFile, 'utf8'));
    res.json(logs);
  } catch { res.json([]); }
});

// ── Admin: 名片資料普及率 ──────────────────────────────────
app.get('/api/admin/contact-completeness', requireAdmin, (req, res) => {
  const data = db.load();
  const auth = loadAuth();
  const contacts = (data.contacts || []).filter(c => !c.deleted);

  const CORE_FIELDS = [
    { key: 'company',      label: '公司',     check: c => !!c.company },
    { key: 'title',        label: '職稱',     check: c => !!c.title },
    { key: 'phone',        label: '電話',     check: c => !!(c.phone || c.mobile) },
    { key: 'email',        label: 'Email',    check: c => !!c.email },
    { key: 'industry',     label: '產業別',   check: c => !!c.industry },
    { key: 'taxId',        label: '統一編號', check: c => !!c.taxId },
    { key: 'address',      label: '地址',     check: c => !!c.address },
    { key: 'website',      label: '網站',     check: c => !!c.website },
    { key: 'systemVendor', label: '系統廠商', check: c => !!c.systemVendor },
  ];

  const total = contacts.length;
  const fields = CORE_FIELDS.map(f => {
    const filled = contacts.filter(c => f.check(c)).length;
    return { key: f.key, label: f.label, filled, total,
             pct: total > 0 ? Math.round(filled / total * 100) : 0 };
  });

  const byOwnerMap = {};
  contacts.forEach(c => {
    const o = c.owner || 'unknown';
    if (!byOwnerMap[o]) byOwnerMap[o] = [];
    byOwnerMap[o].push(c);
  });

  const userMap = {};
  (auth.users || []).forEach(u => { userMap[u.username] = u.displayName || u.username; });

  const byOwner = Object.entries(byOwnerMap).map(([username, cs]) => {
    const ownerFields = CORE_FIELDS.map(f => {
      const filled = cs.filter(c => f.check(c)).length;
      return { key: f.key, label: f.label, filled, total: cs.length,
               pct: cs.length > 0 ? Math.round(filled / cs.length * 100) : 0 };
    });
    const totalFilled   = ownerFields.reduce((s, f) => s + f.filled, 0);
    const totalPossible = CORE_FIELDS.length * cs.length;
    return {
      username, displayName: userMap[username] || username,
      total: cs.length,
      pct: totalPossible > 0 ? Math.round(totalFilled / totalPossible * 100) : 0,
      fields: ownerFields
    };
  }).sort((a, b) => a.pct - b.pct);

  res.json({ fields, byOwner, total, updatedAt: new Date().toISOString() });
});

// ── Admin: check current user permissions ────────────────
app.get('/api/me/permissions', requireAuth, (req, res) => {
  const auth = loadAuth();
  const user = auth.users.find(u => u.username === req.session.user.username);
  if (!user) return res.status(404).json({ error: '找不到帳號' });
  res.json({
    role: user.role || 'user',
    canDownloadContacts: user.canDownloadContacts || false,
    canSetTargets: user.canSetTargets || false,
    active: user.active !== false
  });
});

// ── 名片圖片上傳設定（使用 storage 模組的 engine）─────────
const upload = multer({
  storage: storage.getMulterStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowedExt  = /jpeg|jpg|png|gif|webp/;
    const allowedMime = /^image\/(jpeg|png|gif|webp)$/;
    const extOk  = allowedExt.test(path.extname(file.originalname).toLowerCase());
    const mimeOk = allowedMime.test(file.mimetype);
    if (extOk && mimeOk) {
      cb(null, true);
    } else {
      cb(new Error('只允許上傳圖片檔案（jpg/png/gif/webp）'));
    }
  }
});

// ── 角色可視範圍：取得此用戶可見的 owner 清單 ──────────────
function getViewableOwners(req, dataType) {
  const { username, role } = req.session.user;
  if (role === 'manager1') {
    const auth = loadAuth();
    return auth.users
      .filter(u => u.role === 'user' || u.role === 'manager2' || u.username === username)
      .map(u => u.username);
  }
  if (role === 'manager2') {
    const auth = loadAuth();
    return auth.users
      .filter(u => u.role === 'user' || u.username === username)
      .map(u => u.username);
  }
  if (role === 'secretary') {
    // 秘書可看商機預測及應收帳款（業務 + 二級 + 一級主管）
    if (dataType === 'opportunities' || dataType === 'receivables') {
      const auth = loadAuth();
      return auth.users
        .filter(u => u.role === 'user' || u.role === 'manager2' || u.role === 'manager1')
        .map(u => u.username);
    }
    return []; // 聯絡人、拜訪記錄、合約、年度目標秘書看不到
  }
  if (role === 'marketing') {
    // 行銷只看自己的活動與 Lead，不看業務資料
    if (dataType === 'campaigns' || dataType === 'leads') return [username];
    return [];
  }
  return [username]; // 一般業務只看自己
}

// ── 生日提醒（N 天內）──────────────────────────────────
app.get('/api/birthday-reminders', requireAuth, (req, res) => {
  const days = parseInt(req.query.days) || 3;   // 預設提前 3 天
  const data = db.load();
  const auth = loadAuth();
  const role = req.session.user.role;

  // secretary 不看個人資訊
  if (role === 'secretary') return res.json([]);

  const owners = getViewableOwners(req, 'contacts');
  const contacts = (data.contacts || []).filter(c =>
    !c.deleted && owners.includes(c.owner) && c.personalBirthday
  );

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const upcoming = [];
  contacts.forEach(c => {
    const parts = c.personalBirthday.split('/').map(Number);
    if (parts.length < 2) return;
    const [mm, dd] = parts;
    if (!mm || !dd || mm < 1 || mm > 12 || dd < 1 || dd > 31) return;

    // 計算今年或明年的下一個生日
    let bDate = new Date(today.getFullYear(), mm - 1, dd);
    if (bDate < today) bDate = new Date(today.getFullYear() + 1, mm - 1, dd);

    const diffDays = Math.round((bDate - today) / 86400000);
    if (diffDays < 0 || diffDays > days) return;

    const ownerUser = auth.users.find(u => u.username === c.owner);
    upcoming.push({
      id:               c.id,
      name:             c.name,
      nameEn:           c.nameEn || '',
      company:          c.company || '',
      title:            c.title  || '',
      personalBirthday: c.personalBirthday,
      owner:            c.owner,
      ownerName:        ownerUser?.displayName || c.owner,
      daysLeft:         diffDays,
      birthdayFull:     `${bDate.getFullYear()}/${String(mm).padStart(2,'0')}/${String(dd).padStart(2,'0')}`
    });
  });

  upcoming.sort((a, b) => a.daysLeft - b.daysLeft);
  res.json(upcoming);
});

// ── 取得所有聯絡人 ──────────────────────────────────────
app.get('/api/contacts', requireAuth, (req, res) => {
  const { search } = req.query;
  const data = db.load();
  const role = req.session.user.role;
  if (role === 'secretary') return res.json([]);
  const owners = getViewableOwners(req, 'contacts');
  let contacts = (data.contacts || []).filter(c => owners.includes(c.owner) && !c.deleted);
  if (search) {
    const kw = search.toLowerCase();
    contacts = contacts.filter(c =>
      (c.name || '').toLowerCase().includes(kw) ||
      (c.company || '').toLowerCase().includes(kw) ||
      (c.title || '').toLowerCase().includes(kw) ||
      (c.phone || '').includes(kw) ||
      (c.email || '').toLowerCase().includes(kw) ||
      (c.note || '').toLowerCase().includes(kw)
    );
  }
  contacts.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'zh-TW'));
  res.json(contacts);
});

// ── 新增聯絡人 ──────────────────────────────────────────
app.post('/api/contacts', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  const contact = {
    id: uuidv4(),
    owner,
    name: req.body.name || '',
    nameEn: req.body.nameEn || '',
    company: req.body.company || '',
    title: req.body.title || '',
    phone: req.body.phone || '',
    mobile: req.body.mobile || '',
    ext: req.body.ext || '',
    email: req.body.email || '',
    address: req.body.address || '',
    website: sanitizeUrl(req.body.website),
    taxId: req.body.taxId || '',
    industry: req.body.industry || '',
    opportunityStage: req.body.opportunityStage || '',
    isPrimary: req.body.isPrimary === true || req.body.isPrimary === 'true',
    systemVendor: req.body.systemVendor || '',
    systemProduct: req.body.systemProduct || '',
    note: req.body.note || '',
    cardImage: req.body.cardImage || '',
    createdAt: new Date().toISOString()
  };
  // 若設為主要聯繫窗口，取消同公司（同擁有者）其他人的主要狀態
  if (contact.isPrimary && contact.company) {
    data.contacts.forEach(c => {
      if (!c.deleted && c.owner === owner && c.company === contact.company) c.isPrimary = false;
    });
  }
  data.contacts.push(contact);
  syncCompanyFields(data, contact);  // 同步公司層級欄位到同公司其他空白名片
  db.save(data);
  writeContactAudit('CREATE', req, contact, []);
  res.status(201).json(contact);
});

// ── 更新聯絡人 ──────────────────────────────────────────
app.put('/api/contacts/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  const idx = data.contacts.findIndex(c => c.id === req.params.id && c.owner === owner && !c.deleted);
  if (idx === -1) return res.status(404).json({ error: '找不到此聯絡人' });
  const old = data.contacts[idx];
  const safeBody = pickFields(req.body, CONTACT_FIELDS);
  if (safeBody.website !== undefined) safeBody.website = sanitizeUrl(safeBody.website);
  const updated = { ...old, ...safeBody, id: req.params.id, owner };
  updated.isPrimary = req.body.isPrimary === true || req.body.isPrimary === 'true';
  // 若設為主要聯繫窗口，取消同公司（同擁有者）其他人的主要狀態
  if (updated.isPrimary && updated.company) {
    data.contacts.forEach((c, i) => {
      if (i !== idx && !c.deleted && c.owner === owner && c.company === updated.company) c.isPrimary = false;
    });
  }
  // 計算變更 diff
  const changes = CONTACT_FIELDS
    .filter(f => String(old[f] ?? '') !== String(updated[f] ?? ''))
    .map(f => ({ field: f, fieldLabel: FIELD_LABELS[f] || f, oldValue: old[f] ?? '', newValue: updated[f] ?? '' }));
  data.contacts[idx] = updated;
  syncCompanyFields(data, updated);  // 同步公司層級欄位到同公司其他空白名片
  db.save(data);
  if (changes.length > 0) writeContactAudit('UPDATE', req, updated, changes);
  res.json(data.contacts[idx]);
});

// ── 刪除聯絡人（軟刪除）────────────────────────────────
app.delete('/api/contacts/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  const idx = data.contacts.findIndex(c => c.id === req.params.id && c.owner === owner && !c.deleted);
  if (idx === -1) return res.status(404).json({ error: '找不到此聯絡人' });
  const contact = data.contacts[idx];
  // 軟刪除：標記而不移除（圖片保留，等永久刪除時才移除）
  contact.deleted = true;
  contact.deletedAt = new Date().toISOString();
  contact.deletedBy = owner;
  contact.deletedByName = req.session.user.displayName || owner;
  db.save(data);
  writeContactAudit('DELETE', req, contact, []);
  res.json({ success: true });
});

// ── 集團 CRUD ──────────────────────────────────────────────
app.get('/api/groups', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.groups) data.groups = [];
  res.json(data.groups.filter(g => g.owner === owner));
});

app.post('/api/groups', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const { name, taxId, industry, website, address, note, memberCompanies } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: '集團名稱為必填' });
  const data = db.load();
  if (!data.groups) data.groups = [];
  const group = {
    id: uuidv4(), owner, name: name.trim(), taxId: taxId || '', industry: industry || '',
    website: website || '', address: address || '', note: note || '',
    memberCompanies: Array.isArray(memberCompanies) ? memberCompanies : [],
    createdAt: new Date().toISOString()
  };
  data.groups.push(group);
  db.save(data);
  writeLog('CREATE_GROUP', owner, group.name, `集團：${group.name}`, req);
  res.json(group);
});

app.put('/api/groups/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.groups) data.groups = [];
  const idx = data.groups.findIndex(g => g.id === req.params.id && g.owner === owner);
  if (idx === -1) return res.status(404).json({ error: '找不到此集團' });
  const { name, taxId, industry, website, address, note, memberCompanies } = req.body;
  if (!name || !name.trim()) return res.status(400).json({ error: '集團名稱為必填' });
  data.groups[idx] = {
    ...data.groups[idx], name: name.trim(), taxId: taxId || '',
    industry: industry || '', website: website || '', address: address || '', note: note || '',
    memberCompanies: Array.isArray(memberCompanies) ? memberCompanies : [],
    updatedAt: new Date().toISOString()
  };
  db.save(data);
  writeLog('UPDATE_GROUP', owner, name, `集團：${name}`, req);
  res.json(data.groups[idx]);
});

app.delete('/api/groups/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.groups) data.groups = [];
  const idx = data.groups.findIndex(g => g.id === req.params.id && g.owner === owner);
  if (idx === -1) return res.status(404).json({ error: '找不到此集團' });
  const [deleted] = data.groups.splice(idx, 1);
  db.save(data);
  writeLog('DELETE_GROUP', owner, deleted.name, `集團：${deleted.name}`, req);
  res.json({ success: true });
});

// ── 上傳名片圖片 ────────────────────────────────────────
app.post('/api/upload', requireAuth, upload.single('card'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '未收到圖片' });
  res.json({ url: `/uploads/${req.file.filename}` });
});

// ── 匯出 Excel ──────────────────────────────────────────
app.get('/api/export', requireAuth, (req, res, next) => {
  const auth = loadAuth();
  const user = auth.users.find(u => u.username === req.session.user.username);
  if (!user || (!user.canDownloadContacts && user.role !== 'admin')) {
    return res.status(403).json({ error: '您沒有下載客戶名單的權限' });
  }
  next();
}, (req, res) => {
  const data = db.load();
  const exportOwners = getViewableOwners(req, 'contacts');
  const rows = (data.contacts || []).filter(c => exportOwners.includes(c.owner) && !c.deleted).map(c => ({
    '姓名': c.name,
    '英文名稱': c.nameEn,
    '公司': c.company,
    '職稱': c.title,
    '電話': c.phone,
    '分機': c.ext,
    '手機': c.mobile,
    'Email': c.email,
    '地址': c.address,
    '網站': c.website,
    '統一編號': c.taxId,
    '產業屬性': c.industry,
    '商機分類': c.opportunityStage || '',
    '使用中系統': c.systemVendor,
    '系統產品': c.systemProduct,
    '備註': c.note,
    '建立時間': c.createdAt ? new Date(c.createdAt).toLocaleString('zh-TW') : ''
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [14,20,12,14,14,24,30,20,16,12,14,14,20,16].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '聯絡人');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="contacts.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── 拜訪記錄 CRUD ────────────────────────────────────────
app.get('/api/visits', requireAuth, (req, res) => {
  const data = db.load();
  const role = req.session.user.role;
  if (role === 'secretary') return res.json([]);
  const owners = getViewableOwners(req, 'visits');
  res.json((data.visits || []).filter(v => owners.includes(v.owner)));
});

app.post('/api/visits', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.visits) data.visits = [];
  const visit = {
    id: uuidv4(),
    owner,
    contactId:   req.body.contactId   || '',
    contactName: req.body.contactName || '',
    visitDate:   req.body.visitDate   || '',
    visitType:   req.body.visitType   || '親訪',
    topic:       req.body.topic       || '',
    content:     req.body.content     || '',
    nextAction:  req.body.nextAction  || '',
    createdAt:   new Date().toISOString()
  };
  data.visits.push(visit);
  db.save(data);
  writeLog('CREATE_VISIT', owner, visit.contactName || visit.id,
    `${visit.visitType} ${visit.visitDate} 主題:${visit.topic}`, req);
  res.status(201).json(visit);
});

app.put('/api/visits/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.visits) data.visits = [];
  const idx = data.visits.findIndex(v => v.id === req.params.id && v.owner === owner);
  if (idx === -1) return res.status(404).json({ error: '找不到此記錄' });
  data.visits[idx] = { ...data.visits[idx], ...pickFields(req.body, VISIT_FIELDS), id: req.params.id, owner };
  db.save(data);
  writeLog('UPDATE_VISIT', owner, data.visits[idx].contactName || req.params.id,
    `${data.visits[idx].visitType} ${data.visits[idx].visitDate}`, req);
  res.json(data.visits[idx]);
});

app.delete('/api/visits/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.visits) return res.json({ success: true });
  const v = data.visits.find(v => v.id === req.params.id && v.owner === owner);
  data.visits = data.visits.filter(v => !(v.id === req.params.id && v.owner === owner));
  db.save(data);
  if (v) writeLog('DELETE_VISIT', owner, v.contactName || req.params.id,
    `${v.visitType} ${v.visitDate}`, req);
  res.json({ success: true });
});

// ── 主管業績達成率總覽 ───────────────────────────────────
app.get('/api/manager/achievement', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  if (!['manager1', 'manager2', 'admin'].includes(role)) {
    return res.status(403).json({ error: '權限不足' });
  }
  const year = parseInt(req.query.year) || new Date().getFullYear();
  const auth = loadAuth();
  const data = db.load();
  const viewableUsernames = getViewableOwners(req, 'opportunities');

  // 只列業務 & 二級主管（依角色過濾顯示範圍）
  const salesUsers = auth.users.filter(u => viewableUsernames.includes(u.username));

  const yearStart = new Date(year, 0, 1);
  const yearEnd   = new Date(year, 11, 31, 23, 59, 59);

  const rows = salesUsers.map(u => {
    const target = (data.targets || []).find(t => t.owner === u.username && t.year === year);

    // 主管彙總轄下所有人的商機；一般業務只看自己
    let rowOwners;
    if (u.role === 'manager1') {
      // 一級主管：彙總所有可見成員（含 manager2 + user + 自己）
      rowOwners = viewableUsernames;
    } else if (u.role === 'manager2') {
      // 二級主管：彙總 user 角色 + 自己（不含其他 manager2）
      rowOwners = auth.users
        .filter(x => x.role === 'user' || x.username === u.username)
        .map(x => x.username);
    } else {
      rowOwners = [u.username];
    }
    const myOpps = (data.opportunities || []).filter(o => rowOwners.includes(o.owner));

    // 成交：以 achievedDate 為準，無則 createdAt
    const achieved = myOpps
      .filter(o => o.stage === 'Won')
      .filter(o => {
        const d = new Date(o.achievedDate || o.updatedAt || o.createdAt);
        return d >= yearStart && d <= yearEnd;
      })
      .reduce((s, o) => s + (parseFloat(o.amount) || 0), 0);

    // 在手商機（排除 Won、D 停止中）
    const pipeline = myOpps
      .filter(o => !['Won','D'].includes(o.stage))
      .reduce((s, o) => s + (parseFloat(o.amount) || 0), 0);

    const wonCount = myOpps.filter(o => o.stage === 'Won' && (() => {
      const d = new Date(o.achievedDate || o.updatedAt || o.createdAt);
      return d >= yearStart && d <= yearEnd;
    })()).length;

    const targetAmt = target ? (parseFloat(target.amount) || 0) : 0;
    const rate = targetAmt > 0 ? Math.round(achieved / targetAmt * 100) : null;

    return {
      username:    u.username,
      displayName: u.displayName || u.username,
      role:        u.role,
      target:      targetAmt,
      achieved,
      pipeline,
      wonCount,
      rate,  // null = 未設目標
    };
  });

  // 按達成率降冪，未設目標排最後
  rows.sort((a, b) => {
    if (a.rate === null && b.rate === null) return b.achieved - a.achieved;
    if (a.rate === null) return 1;
    if (b.rate === null) return -1;
    return b.rate - a.rate;
  });

  res.json({ year, rows });
});

// ── 主管幫特定業務設定年度目標 ──────────────────────────
app.put('/api/manager/target/:username', requireAuth, (req, res) => {
  const { role } = req.session.user;
  if (!['manager1', 'manager2', 'admin'].includes(role)) {
    return res.status(403).json({ error: '權限不足' });
  }
  const targetUsername = req.params.username;
  const viewable = getViewableOwners(req, 'opportunities');
  if (!viewable.includes(targetUsername)) {
    return res.status(403).json({ error: '無權限編輯此業務目標' });
  }
  const year   = parseInt(req.body.year);
  const amount = parseFloat(req.body.amount);
  if (!year || isNaN(amount) || amount < 0) {
    return res.status(400).json({ error: '請輸入正確的年度與金額' });
  }
  const data = db.load();
  if (!data.targets) data.targets = [];
  const existing = data.targets.find(t => t.year === year && t.owner === targetUsername);
  if (existing) {
    existing.amount = amount;
    existing.updatedAt = new Date().toISOString();
    existing.setByManager = req.session.user.username;
  } else {
    data.targets.push({
      id: require('crypto').randomUUID(),
      owner: targetUsername, year, amount,
      createdAt: new Date().toISOString(),
      setByManager: req.session.user.username,
    });
  }
  db.save(data);
  res.json({ success: true, username: targetUsername, year, amount });
});

// ── 年度目標 CRUD ────────────────────────────────────────
app.get('/api/targets', requireAuth, (req, res) => {
  const data = db.load();
  const role = req.session.user.role;
  if (role === 'secretary') return res.json([]);
  const owners = getViewableOwners(req, 'targets');
  res.json((data.targets || []).filter(t => owners.includes(t.owner)));
});

app.post('/api/targets', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.targets) data.targets = [];
  const year = parseInt(req.body.year);
  const amount = parseFloat(req.body.amount) || 0;
  const existing = data.targets.find(t => t.year === year && t.owner === owner);
  if (existing) {
    existing.amount = amount;
    existing.updatedAt = new Date().toISOString();
    db.save(data);
    writeLog('SET_TARGET', owner, owner, `${year} 年目標更新為 ${amount}`, req);
    return res.json(existing);
  }
  const target = { id: uuidv4(), owner, year, amount, createdAt: new Date().toISOString() };
  data.targets.push(target);
  db.save(data);
  writeLog('SET_TARGET', owner, owner, `${year} 年目標 ${amount}`, req);
  res.status(201).json(target);
});

// ── 季度配比設定 ─────────────────────────────────────────
// GET  /api/settings/quarter-ratios          → { "2026":[20,30,30,20], ... }
// PUT  /api/settings/quarter-ratios          → body { year, ratios:[q1,q2,q3,q4] }
app.get('/api/settings/quarter-ratios', requireAuth, (req, res) => {
  const data = db.load();
  res.json((data.settings && data.settings.quarterRatios) || {});
});

app.put('/api/settings/quarter-ratios', requireAdmin, (req, res) => {
  const { year, ratios } = req.body;
  const y = parseInt(year);
  if (!y || !Array.isArray(ratios) || ratios.length !== 4) {
    return res.status(400).json({ error: '格式錯誤：需提供 year 與 ratios[4]' });
  }
  const sum = ratios.reduce((a, v) => a + (parseFloat(v) || 0), 0);
  if (Math.round(sum) !== 100) {
    return res.status(400).json({ error: `配比合計需為 100，目前為 ${sum}` });
  }
  const data = db.load();
  if (!data.settings) data.settings = {};
  if (!data.settings.quarterRatios) data.settings.quarterRatios = {};
  data.settings.quarterRatios[y] = ratios.map(v => parseFloat(v) || 0);
  db.save(data);
  writeLog('SET_QUARTER_RATIO', req.session.user.username, String(y),
    `Q1~Q4配比: ${ratios.join('/')}`, req);
  res.json({ success: true, year: y, ratios: data.settings.quarterRatios[y] });
});

// ── 個人季度配比 ─────────────────────────────────────────
// GET /api/settings/user-quarter-ratios
// PUT /api/settings/user-quarter-ratios  body { username, year, ratios:[4] }
app.get('/api/settings/user-quarter-ratios', requireAuth, (req, res) => {
  const data = db.load();
  const role = req.session.user.role;
  const username = req.session.user.username;
  const all = (data.settings && data.settings.userQuarterRatios) || {};
  if (role === 'admin' || role === 'manager1') return res.json(all);
  return res.json(all[username] ? { [username]: all[username] } : {});
});

app.put('/api/settings/user-quarter-ratios', requireAdmin, (req, res) => {
  const { username, year, ratios, clearPersonal } = req.body;
  const y = parseInt(year);
  if (!username || !y) {
    return res.status(400).json({ error: '格式錯誤：需提供 username, year' });
  }
  const data = db.load();
  if (!data.settings) data.settings = {};
  if (!data.settings.userQuarterRatios) data.settings.userQuarterRatios = {};

  // 清除個人配比（回到全域預設）
  if (clearPersonal) {
    if (data.settings.userQuarterRatios[username]) {
      delete data.settings.userQuarterRatios[username][y];
      if (Object.keys(data.settings.userQuarterRatios[username]).length === 0) {
        delete data.settings.userQuarterRatios[username];
      }
    }
    db.save(data);
    writeLog('CLEAR_USER_QR', req.session.user.username, username, `清除 ${y}年個人配比`, req);
    return res.json({ success: true, cleared: true });
  }

  if (!Array.isArray(ratios) || ratios.length !== 4) {
    return res.status(400).json({ error: '格式錯誤：需提供 ratios[4]' });
  }
  const sum = ratios.reduce((a, v) => a + (parseFloat(v) || 0), 0);
  if (Math.round(sum) !== 100) {
    return res.status(400).json({ error: `配比合計需為 100，目前為 ${sum.toFixed(1)}` });
  }
  if (!data.settings.userQuarterRatios[username]) data.settings.userQuarterRatios[username] = {};
  data.settings.userQuarterRatios[username][y] = ratios.map(v => parseFloat(v) || 0);
  db.save(data);
  writeLog('SET_USER_QR', req.session.user.username, username,
    `${y}年個人配比: ${ratios.join('/')}`, req);
  res.json({ success: true });
});

// GET  /api/settings/opr-range → { min:1.1, max:1.2 }
// PUT  /api/settings/opr-range → body { min, max }
app.get('/api/settings/opr-range', requireAuth, (req, res) => {
  const data = db.load();
  res.json((data.settings && data.settings.oprRange) || { min: 1.1, max: 1.2 });
});

app.put('/api/settings/opr-range', requireAdmin, (req, res) => {
  const min = parseFloat(req.body.min);
  const max = parseFloat(req.body.max);
  if (isNaN(min) || isNaN(max) || min <= 0 || max <= min) {
    return res.status(400).json({ error: '格式錯誤：min/max 需為正數且 min < max' });
  }
  const data = db.load();
  if (!data.settings) data.settings = {};
  data.settings.oprRange = { min, max };
  db.save(data);
  writeLog('SET_OPR_RANGE', req.session.user.username, '-', `OPR範圍: ${min}~${max}`, req);
  res.json({ success: true, min, max });
});

// ── 商機 CRUD ────────────────────────────────────────────
app.get('/api/opportunities', requireAuth, (req, res) => {
  const data = db.load();
  const owners = getViewableOwners(req, 'opportunities');
  res.json((data.opportunities || []).filter(o => owners.includes(o.owner)));
});

app.post('/api/opportunities', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.opportunities) data.opportunities = [];
  const opp = {
    id: uuidv4(),
    owner,
    contactId:       req.body.contactId       || '',
    contactName:     req.body.contactName     || '',
    company:         req.body.company         || '',
    category:        req.body.category        || '',
    product:         req.body.product         || '',
    amount:          req.body.amount          || '',
    expectedDate:    req.body.expectedDate    || '',
    description:     req.body.description     || '',
    grossMarginRate: req.body.grossMarginRate || '',
    stage:           'C',
    visitId:         req.body.visitId         || '',
    createdAt:       new Date().toISOString()
  };
  data.opportunities.push(opp);
  if (opp.expectedDate) {
    if (!data.opportunityDateChanges) data.opportunityDateChanges = [];
    data.opportunityDateChanges.push({
      dealId: opp.id, dealValue: parseFloat(opp.amount) || 0,
      oldDate: null, newDate: opp.expectedDate,
      changedAt: new Date().toISOString(), owner: opp.owner
    });
  }
  db.save(data);
  writeLog('CREATE_OPP', owner, opp.company || opp.contactName,
    `${opp.product} 階段:C 金額:${opp.amount}`, req);
  res.status(201).json(opp);
});

app.put('/api/opportunities/:id', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  const data = db.load();
  if (!data.opportunities) data.opportunities = [];
  // 一般業務只能改自己的；主管/admin 可改其可查看範圍內的商機（owner 不變）
  const viewable = getViewableOwners(req, 'opportunities');
  const idx = data.opportunities.findIndex(o =>
    o.id === req.params.id && viewable.includes(o.owner)
  );
  if (idx === -1) return res.status(404).json({ error: '找不到此商機' });
  const owner = data.opportunities[idx].owner; // 保留原始 owner
  const oldStage = data.opportunities[idx].stage;
  const oldExpectedDate = data.opportunities[idx].expectedDate || null;
  data.opportunities[idx] = { ...data.opportunities[idx], ...pickFields(req.body, OPP_FIELDS), id: req.params.id, owner };
  const newStage = data.opportunities[idx].stage;
  // 記錄預計簽約日變動歷史
  const newExpectedDate = data.opportunities[idx].expectedDate || null;
  if (newExpectedDate && oldExpectedDate !== newExpectedDate) {
    if (!data.opportunityDateChanges) data.opportunityDateChanges = [];
    data.opportunityDateChanges.push({
      dealId: req.params.id, dealValue: parseFloat(data.opportunities[idx].amount) || 0,
      oldDate: oldExpectedDate, newDate: newExpectedDate,
      changedAt: new Date().toISOString(), owner
    });
  }
  // 記錄階段變動歷史
  if (oldStage && newStage && oldStage !== newStage) {
    if (!data.opportunities[idx].stageHistory) data.opportunities[idx].stageHistory = [];
    data.opportunities[idx].stageHistory.push({
      from: oldStage, to: newStage,
      date: new Date().toISOString(),
      changedBy: owner
    });
  }
  db.save(data);
  const stageDetail = oldStage !== newStage ? ` 階段:${oldStage}→${newStage}` : ` 階段:${newStage}`;
  writeLog('UPDATE_OPP', username, data.opportunities[idx].company || data.opportunities[idx].contactName,
    `${data.opportunities[idx].product}${stageDetail}`, req);
  res.json(data.opportunities[idx]);
});

// ── Pipeline 月度變動報表 ─────────────────────────────────
app.get('/api/pipeline-date-changes', requireAuth, (req, res) => {
  const { role } = req.session.user;
  if (!['admin', 'manager1', 'secretary'].includes(role)) {
    return res.status(403).json({ error: '無此頁面權限' });
  }
  const year = parseInt(req.query.year) || new Date().getFullYear();
  const data = db.load();
  const changes = data.opportunityDateChanges || [];

  const months = Array.from({ length: 12 }, (_, i) => {
    const m = i + 1;
    const pfx = `${year}-${String(m).padStart(2, '0')}`;

    // 補進：newDate 落在此月，且 oldDate 不在此月（或 oldDate 為 null）
    const inflow = changes
      .filter(c => c.newDate && c.newDate.startsWith(pfx) &&
                   (c.oldDate === null || !c.oldDate.startsWith(pfx)))
      .reduce((s, c) => s + (c.dealValue || 0), 0);

    // 退後：oldDate 落在此月，且 newDate 比 oldDate 更晚（且不在同月）
    const outflow = changes
      .filter(c => c.oldDate && c.oldDate.startsWith(pfx) &&
                   c.newDate && c.newDate > c.oldDate &&
                   !c.newDate.startsWith(pfx))
      .reduce((s, c) => s + (c.dealValue || 0), 0);

    return { month: m, label: `${m}月`, inflow, outflow, diff: inflow - outflow };
  });

  res.json({ year, months });
});

// ── 商機動態報表 ──────────────────────────────────────────
// 從低到高排列：D(靜止) → C(Pipeline) → B(Upside) → A(Commit) → Won
const STAGE_ORDER = ['D','C','B','A','Won'];
const STAGE_LABEL = { D:'D｜靜止中', C:'C｜Pipeline', B:'B｜Upside', A:'A｜Commit', Won:'🏆 Won' };
app.get('/api/pipeline-report', requireAuth, (req, res) => {
  const { from, to, owner: ownerFilter } = req.query;
  const now = new Date();
  const fromDate = from ? new Date(from) : new Date(now.getFullYear(), now.getMonth(), 1);
  const toDate   = to   ? new Date(to)   : new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  const allOwners = getViewableOwners(req, 'opportunities');
  // 若前端指定了特定業務且在可視範圍內，就篩選該業務
  const owners = (ownerFilter && allOwners.includes(ownerFilter)) ? [ownerFilter] : allOwners;
  const data   = db.load();
  const auth   = loadAuth();
  // 建立業務人員選單（顯示名稱），只有多人可選時才有意義
  const ownerOptions = allOwners.length > 1
    ? allOwners.map(u => {
        const usr = auth.users.find(x => x.username === u);
        return { username: u, displayName: usr ? (usr.displayName || u) : u };
      })
    : [];
  const opps   = (data.opportunities || []).filter(o => owners.includes(o.owner));
  const lostAll= (data.lostOpportunities || []).filter(o => owners.includes(o.owner));

  // 當前各階段漏斗（排除成交，成交單獨計算）
  const funnel = STAGE_ORDER.map(stage => {
    const list = opps.filter(o => o.stage === stage);
    return {
      stage,
      label: STAGE_LABEL[stage] || stage,
      count: list.length,
      amount: list.reduce((s,o) => s+(parseFloat(o.amount)||0), 0),
      deals: list.map(o=>({id:o.id,company:o.company,product:o.product,amount:parseFloat(o.amount)||0,owner:o.owner}))
    };
  });

  // 期間新增
  const newDeals = opps.filter(o => { const d=new Date(o.createdAt); return d>=fromDate&&d<=toDate; });

  // 期間流失
  const lostDeals = lostAll.filter(o => { const d=new Date(o.deletedAt); return d>=fromDate&&d<=toDate; });

  // 期間階段晉升 / 退後
  // 同一商機同一天若多次調整，只保留「淨移動」（當天第一筆 from → 最後一筆 to）
  const promoted=[], demoted=[];
  opps.forEach(o => {
    // 篩出期間內的歷史，按時間升冪排序
    const inRange = (o.stageHistory||[])
      .filter(h => { const d=new Date(h.date); return d>=fromDate&&d<=toDate; })
      .sort((a,b)=>new Date(a.date)-new Date(b.date));

    if (!inRange.length) return;

    // 以「日曆天 YYYY-MM-DD」為 key 分組
    const byDay = {};
    inRange.forEach(h => {
      const dayKey = h.date.slice(0,10);
      if (!byDay[dayKey]) byDay[dayKey] = [];
      byDay[dayKey].push(h);
    });

    // 每天取淨移動：第一筆的 from → 最後一筆的 to
    Object.entries(byDay).forEach(([day, entries]) => {
      const netFrom = entries[0].from;
      const netTo   = entries[entries.length-1].to;
      if (netFrom === netTo) return;   // 來回抵銷，略過
      const fi = STAGE_ORDER.indexOf(netFrom);
      const ti = STAGE_ORDER.indexOf(netTo);
      if (fi === -1 || ti === -1) return;
      const item = {
        id:      o.id, company: o.company, product: o.product,
        amount:  parseFloat(o.amount)||0,
        from:    netFrom, to: netTo,
        date:    entries[entries.length-1].date,   // 使用當天最後操作時間
        owner:   o.owner
      };
      if (ti > fi) promoted.push(item);
      else          demoted.push(item);
    });
  });

  const sum = arr => arr.reduce((s,o)=>s+(parseFloat(o.amount)||0),0);
  res.json({
    period: { from: fromDate.toISOString(), to: toDate.toISOString() },
    ownerOptions,
    funnel,
    newDeals:   newDeals.map(o=>({id:o.id,company:o.company,product:o.product,amount:parseFloat(o.amount)||0,stage:o.stage,owner:o.owner,createdAt:o.createdAt})),
    lostDeals:  lostDeals.map(o=>({id:o.id,company:o.company,product:o.product,amount:parseFloat(o.amount)||0,stage:o.stage,deleteReason:o.deleteReason,deletedAt:o.deletedAt,owner:o.owner})),
    promoted, demoted,
    summary: {
      totalPipeline: sum(opps.filter(o=>o.stage!=='Won'&&o.stage!=='D')),
      totalCount:    opps.filter(o=>o.stage!=='Won'&&o.stage!=='D').length,
      newAmount:  sum(newDeals), newCount: newDeals.length,
      lostAmount: sum(lostDeals), lostCount: lostDeals.length,
      promotedAmount: sum(promoted), promotedCount: promoted.length,
      demotedAmount:  sum(demoted),  demotedCount: demoted.length,
    }
  });
});

app.delete('/api/opportunities/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const { deleteReason } = req.body || {};
  const data = db.load();
  if (!data.opportunities) return res.json({ success: true });
  const opp = data.opportunities.find(o => o.id === req.params.id && o.owner === owner);
  if (!opp) return res.status(404).json({ error: '找不到此商機' });
  // 保存到 lostOpportunities 供事後分析
  if (!data.lostOpportunities) data.lostOpportunities = [];
  data.lostOpportunities.push({
    ...opp,
    deleteReason: deleteReason || '',
    deletedAt: new Date().toISOString(),
    deletedBy: owner,
    deletedByName: req.session.user.displayName || owner
  });
  data.opportunities = data.opportunities.filter(o => !(o.id === req.params.id && o.owner === owner));
  db.save(data);
  writeLog('DELETE_OPP', owner, opp.company || opp.contactName,
    `${opp.product} 原因:${deleteReason || '未說明'}`, req);
  res.json({ success: true });
});

// ── 取得流失商機（主管 / 管理員）─────────────────────────────
app.get('/api/lost-opportunities', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  const data = db.load();
  let list = data.lostOpportunities || [];
  if (role === 'user') {
    list = list.filter(o => o.owner === username);
  } else if (role === 'manager2') {
    const auth = loadAuth();
    const subs = auth.users.filter(u => u.role === 'user').map(u => u.username);
    list = list.filter(o => subs.includes(o.owner) || o.owner === username);
  }
  // 依刪除時間排序（最新在前）
  list = list.sort((a, b) => (b.deletedAt || '').localeCompare(a.deletedAt || ''));
  res.json(list);
});

// ── 殭屍商機偵測 ──────────────────────────────────────────
// 規則：
//   C 級：最後聯繫距今 > 14 天（或從未聯繫）
//   B 級：最後親訪/視訊/展覽 > 24 天 OR 最後電話 > 14 天
//   A 級：任何聯繫形式距今 > 7 天
app.get('/api/zombie-opportunities', requireAuth, (req, res) => {
  const data  = db.load();
  const auth  = loadAuth();
  const owners = getViewableOwners(req, 'opportunities');
  const today  = new Date(); today.setHours(0,0,0,0);

  // contactId → company 對照表
  const contactCo = {};
  (data.contacts || []).forEach(c => { contactCo[c.id] = (c.company || '').trim(); });

  // 建立拜訪索引（依 contactId 與公司名）
  const byContact = {};   // contactId  → visit[]
  const byCompany = {};   // company    → visit[]
  (data.visits || []).filter(v => owners.includes(v.owner) && v.visitDate).forEach(v => {
    if (v.contactId) {
      (byContact[v.contactId] = byContact[v.contactId] || []).push(v);
    }
    const co = contactCo[v.contactId] || '';
    if (co) (byCompany[co] = byCompany[co] || []).push(v);
  });

  const daysSince = d => {
    if (!d) return Infinity;
    const t = new Date(d); t.setHours(0,0,0,0);
    return Math.floor((today - t) / 86400000);
  };

  const FACE_TYPES  = new Set(['親訪','視訊','展覽']);
  const PHONE_TYPES = new Set(['電話']);

  const zombies = [];

  (data.opportunities || [])
    .filter(o => owners.includes(o.owner) && o.stage !== 'Won')
    .forEach(o => {
      // 彙整此商機相關的所有拜訪
      const seen = new Set();
      const visits = [];
      const addV = v => { if (!seen.has(v.id)) { seen.add(v.id); visits.push(v); } };
      if (o.contactId && byContact[o.contactId]) byContact[o.contactId].forEach(addV);
      const co = (o.company || '').trim();
      if (co && byCompany[co]) byCompany[co].forEach(addV);

      // 最後各類型聯繫日期
      const lastAny   = visits.map(v=>v.visitDate).sort().reverse()[0] || null;
      const lastFace  = visits.filter(v=>FACE_TYPES.has(v.visitType)) .map(v=>v.visitDate).sort().reverse()[0] || null;
      const lastPhone = visits.filter(v=>PHONE_TYPES.has(v.visitType)).map(v=>v.visitDate).sort().reverse()[0] || null;

      const dAny   = daysSince(lastAny);
      const dFace  = daysSince(lastFace);
      const dPhone = daysSince(lastPhone);

      let isZombie = false, reasons = [], severity = 'warn';

      if (o.stage === 'C') {
        if (dAny > 14) {
          isZombie = true;
          reasons.push(lastAny ? `最後聯繫距今 ${dAny} 天（超過 14 天）` : '從未有日報記錄');
          if (dAny > 30) severity = 'danger';
        }
      } else if (o.stage === 'B') {
        if (dFace > 24) {
          isZombie = true;
          reasons.push(lastFace ? `最後拜訪距今 ${dFace} 天（超過 24 天）` : '從未有拜訪記錄');
        }
        if (dPhone > 14) {
          isZombie = true;
          reasons.push(lastPhone ? `最後電話距今 ${dPhone} 天（超過 14 天）` : '從未有電話記錄');
        }
        if (isZombie && (dFace > 48 || dPhone > 28)) severity = 'danger';
      } else if (o.stage === 'A') {
        if (dAny > 7) {
          isZombie = true;
          reasons.push(lastAny ? `最後聯繫距今 ${dAny} 天（超過 7 天）` : '從未有聯繫記錄');
          severity = 'danger';
        }
      }

      if (!isZombie) return;

      const u = auth.users.find(u => u.username === o.owner);
      zombies.push({
        id:          o.id,
        company:     o.company     || '',
        contactName: o.contactName || '',
        product:     o.product     || '',
        amount:      o.amount      || 0,
        stage:       o.stage,
        expectedDate:o.expectedDate|| '',
        createdAt:   o.createdAt   || '',
        owner:       o.owner,
        ownerName:   u ? (u.displayName || u.username) : o.owner,
        lastVisit:   lastAny,
        lastFace,
        lastPhone,
        daysSinceAny:   dAny   === Infinity ? null : dAny,
        daysSinceFace:  dFace  === Infinity ? null : dFace,
        daysSincePhone: dPhone === Infinity ? null : dPhone,
        reasons,
        severity
      });
    });

  // 排序：danger 優先 → daysSinceAny 由多到少
  zombies.sort((a,b) => {
    if (a.severity !== b.severity) return a.severity === 'danger' ? -1 : 1;
    return (b.daysSinceAny||0) - (a.daysSinceAny||0);
  });

  res.json(zombies);
});

// ── 還原誤刪商機 ──────────────────────────────────────────
app.post('/api/opportunities/restore/:id', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  const data = db.load();
  if (!data.lostOpportunities) return res.status(404).json({ error: '找不到此流失商機' });

  const idx = data.lostOpportunities.findIndex(o => o.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '找不到此流失商機' });

  const opp = data.lostOpportunities[idx];
  // 權限：user 只能還原自己的，manager 可還原轄下
  if (role === 'user' && opp.owner !== username) {
    return res.status(403).json({ error: '無權限還原此商機' });
  }

  // 還原：移除 deleted 相關欄位，設為 C 階段（或原始 stage）
  const restored = { ...opp };
  delete restored.deletedAt;
  delete restored.deleteReason;
  delete restored.deletedBy;
  delete restored.deletedByName;
  if (!restored.stage || restored.stage === '流失') restored.stage = 'C';

  if (!data.opportunities) data.opportunities = [];
  data.opportunities.push(restored);
  data.lostOpportunities.splice(idx, 1);
  db.save(data);
  res.json({ success: true, opportunity: restored });
});

// ── 銷售預測 Excel 匯出 ──────────────────────────────────
const STAGE_CONF       = { D: 10, C: 25, B: 50, A: 90, Won: 100 };
const STAGE_LABEL_EXPORT = { A: 'Commit', B: 'Upside', C: 'Pipeline', Won: 'Won' };

app.get('/api/forecast/export', requireAuth, (req, res) => {
  const yr   = parseInt(req.query.year) || new Date().getFullYear();
  const data = db.load();
  const user = req.session.user;
  const salesPerson = user ? (user.displayName || user.username) : '';

  // 篩選當年度商機（依預定簽約日，依角色可視範圍）
  const auth = loadAuth();
  const forecastOwners = getViewableOwners(req, 'opportunities');
  const opps = (data.opportunities || [])
    .filter(o => forecastOwners.includes(o.owner) && o.expectedDate && new Date(o.expectedDate).getFullYear() === yr && o.stage !== 'D')
    .sort((a, b) => (a.expectedDate || '').localeCompare(b.expectedDate || ''));

  // 標題列
  const titleRow = ['未來商機預測'];

  // 表頭列
  const headers = [
    '客戶名稱', '銷售案名', 'BU', '預定簽約日',
    '業務人員', '把握度', '預估毛利率',
    '合約金額(NT$K)', '毛利金額(NT$K)'
  ];

  // 資料列
  const rows = opps.map(o => {
    const amt    = (parseFloat(o.amount) || 0) * 10;
    const gm     = parseFloat(o.grossMarginRate) || 0;
    const profit = Math.round(amt * gm / 100);
    return [
      o.company      || '',
      o.product      || o.description || '',
      o.category     || '',
      o.expectedDate || '',
      auth.users.find(u => u.username === o.owner)?.displayName || o.owner || salesPerson,
      STAGE_LABEL_EXPORT[o.stage] || '',
      gm   ? gm   + '%' : '',
      amt  || '',
      profit || ''
    ];
  });

  // 合計列
  const totAmt    = opps.reduce((s, o) => s + (parseFloat(o.amount) || 0) * 10, 0);
  const totProfit = Math.round(opps.reduce((s, o) => {
    const a = (parseFloat(o.amount) || 0) * 10;
    const g = parseFloat(o.grossMarginRate) || 0;
    return s + a * g / 100;
  }, 0));
  const totalRow = [
    `合計（${opps.length}筆）`, '', '', '', '',
    '', '',
    totAmt || '', totProfit || ''
  ];

  const wsData = [titleRow, headers, ...rows, totalRow];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // 欄寬
  ws['!cols'] = [
    { wch: 28 }, { wch: 30 }, { wch: 8 }, { wch: 14 },
    { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 16 }, { wch: 16 }
  ];

  // 合併標題列 A1:I1
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }];

  // 儲存格樣式（標題粗體大字）
  const titleCell = ws['A1'];
  if (titleCell) {
    titleCell.s = {
      font: { bold: true, sz: 16 },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
  }

  const sheetName = `${yr}年`;
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const today = new Date();
  const dateStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
  const fname = `${yr}%E5%B9%B4SalesPipeline_${encodeURIComponent(salesPerson || 'export')}_${dateStr}.xlsx`;
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fname}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── 商機全量匯出（Admin）─────────────────────────────────
app.get('/api/admin/opportunities/export', requireAdmin, (req, res) => {
  const data = db.load();
  const auth = loadAuth();
  const userMap = {};
  (auth.users || []).forEach(u => { userMap[u.username] = u.displayName || u.username; });

  const STAGE_LABELS = { A: 'Commit', B: 'Upside', C: 'Pipeline', C2: 'Pipeline', D: 'D', Won: 'Won' };

  const headers = [
    '客戶名稱', '銷售案名', 'BU(category)', '預定簽約日',
    '業務帳號(owner)', '業務姓名', '把握度階段(A/B/C/Won)',
    '合約金額(萬元)', '預估毛利率(%)', '備註(description)',
    '建立時間', '商機ID'
  ];

  const rows = (data.opportunities || []).map(o => [
    o.company      || '',
    o.product      || '',
    o.category     || '',
    o.expectedDate || '',
    o.owner        || '',
    userMap[o.owner] || o.owner || '',
    o.stage        || '',
    o.amount       || '',
    o.grossMarginRate || '',
    o.description  || '',
    o.createdAt    ? o.createdAt.slice(0, 10) : '',
    o.id           || ''
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = [
    { wch: 28 }, { wch: 30 }, { wch: 10 }, { wch: 14 },
    { wch: 14 }, { wch: 12 }, { wch: 18 },
    { wch: 14 }, { wch: 14 }, { wch: 30 },
    { wch: 12 }, { wch: 36 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, '商機資料');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const today = new Date();
  const dateStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%E5%95%86%E6%A9%9F%E8%B3%87%E6%96%99_${dateStr}.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── 客戶資料匯出（Admin）────────────────────────────────
app.get('/api/admin/contacts/export', requireAdmin, (req, res) => {
  const data = db.load();
  const auth = loadAuth();
  const userMap = {};
  (auth.users || []).forEach(u => { userMap[u.username] = u.displayName || u.username; });

  const headers = [
    '姓名', '英文名稱', '公司', '職稱', '電話', '分機', '手機', 'Email',
    '地址', '網站', '統一編號', '產業屬性', '商機分類', '使用中系統', '系統產品',
    '備註', '業務帳號(owner)', '業務姓名', '建立時間', '聯絡人ID'
  ];

  const rows = (data.contacts || []).filter(c => !c.deleted).map(c => [
    c.name        || '',
    c.nameEn      || '',
    c.company     || '',
    c.title       || '',
    c.phone       || '',
    c.ext         || '',
    c.mobile      || '',
    c.email       || '',
    c.address     || '',
    c.website     || '',
    c.taxId       || '',
    c.industry    || '',
    c.opportunityStage || '',
    c.systemVendor || '',
    c.systemProduct || '',
    c.note        || '',
    c.owner       || '',
    userMap[c.owner] || c.owner || '',
    c.createdAt   ? c.createdAt.slice(0, 10) : '',
    c.id          || ''
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = [
    { wch: 14 }, { wch: 18 }, { wch: 26 }, { wch: 14 },
    { wch: 14 }, { wch: 8  }, { wch: 14 }, { wch: 28 },
    { wch: 30 }, { wch: 24 }, { wch: 12 }, { wch: 14 },
    { wch: 12 }, { wch: 16 }, { wch: 16 },
    { wch: 30 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 36 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, '客戶資料');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const today = new Date();
  const dateStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%E5%AE%A2%E6%88%B6%E8%B3%87%E6%96%99_${dateStr}.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── 客戶資料批次匯入（Admin）────────────────────────────
app.post('/api/admin/contacts/import', requireAdmin,
  (req, res, next) => uploadImport.single('file')(req, res, next),
  async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: '請上傳 Excel 檔案' });

      const { defaultOwner, skipDuplicates } = req.body;

      const buf  = req.file.buffer;
      const wb   = XLSX.read(buf, { type: 'buffer' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      if (rows.length < 2) return res.status(400).json({ error: '檔案無資料列' });

      const header = rows[0].map(h => String(h).trim());

      // 欄位 index 對應（容錯：取包含關鍵字的欄位）
      const col = k => header.findIndex(h => h.includes(k));
      const COL = {
        name:           col('姓名'),
        nameEn:         col('英文名稱'),
        company:        col('公司'),
        title:          col('職稱'),
        phone:          col('電話'),
        ext:            col('分機'),
        mobile:         col('手機'),
        email:          col('Email'),
        address:        col('地址'),
        website:        col('網站'),
        taxId:          col('統一編號'),
        industry:       col('產業屬性'),
        opportunityStage: col('商機分類'),
        systemVendor:   col('使用中系統'),
        systemProduct:  col('系統產品'),
        note:           col('備註'),
        owner:          col('業務帳號'),
      };

      const auth = loadAuth();
      const usernames = new Set((auth.users || []).map(u => u.username));
      const displayToUser = {};
      (auth.users || []).forEach(u => { displayToUser[u.displayName || u.username] = u.username; });

      const data = db.load();
      if (!data.contacts) data.contacts = [];

      let imported = 0, skipped = 0;
      const errors = [];

      rows.slice(1).forEach((row, i) => {
        const rowNum = i + 2;
        if (!row.some(c => String(c).trim())) return; // 空列

        const get = idx => idx >= 0 ? String(row[idx] ?? '').trim() : '';

        const name    = get(COL.name);
        const company = get(COL.company);
        if (!name && !company) { skipped++; return; }

        // 解析 owner
        let owner = get(COL.owner);
        if (owner && !usernames.has(owner)) {
          owner = displayToUser[owner] || owner;
        }
        if (!owner || !usernames.has(owner)) {
          if (defaultOwner && usernames.has(defaultOwner)) {
            owner = defaultOwner;
          } else {
            errors.push(`第${rowNum}列：業務帳號「${owner}」不存在，請填寫或設定預設業務`);
            return;
          }
        }

        // 重複檢查（同 owner、同公司、同姓名）
        if (skipDuplicates === 'true') {
          const dup = data.contacts.find(c =>
            !c.deleted && c.owner === owner &&
            c.name === name && c.company === company
          );
          if (dup) { skipped++; return; }
        }

        let website = get(COL.website);
        if (website && !/^https?:\/\//i.test(website)) website = '';

        data.contacts.push({
          id: uuidv4(), owner, createdAt: new Date().toISOString(),
          name, nameEn: get(COL.nameEn), company, title: get(COL.title),
          phone: get(COL.phone), ext: get(COL.ext), mobile: get(COL.mobile),
          email: get(COL.email), address: get(COL.address), website,
          taxId: get(COL.taxId), industry: get(COL.industry),
          opportunityStage: get(COL.opportunityStage),
          systemVendor: get(COL.systemVendor), systemProduct: get(COL.systemProduct),
          note: get(COL.note), isPrimary: false, cardImage: '',
          deleted: false
        });
        imported++;
      });

      db.save(data);
      writeLog('IMPORT_CONTACTS_ADMIN', req.session.user.username, 'admin',
        `Admin 批次匯入客戶：${imported} 筆成功，${skipped} 筆略過，${errors.length} 筆錯誤`, req);

      res.json({ success: true, imported, skipped, errors });
    } catch (e) {
      console.error('[import-contacts]', e);
      res.status(500).json({ error: '匯入失敗：' + e.message });
    }
  }
);

// ── 商機批次匯入（Admin）─────────────────────────────────
app.post('/api/admin/opportunities/import', requireAdmin, (req, res, next) => uploadImport.single('file')(req, res, next), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: '請上傳 Excel 檔案' });

    const buf  = req.file.buffer || require('fs').readFileSync(req.file.path);
    const wb   = XLSX.read(buf, { type: 'buffer' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    if (rows.length < 2) return res.status(400).json({ error: '檔案無資料列' });

    // 找標頭列（允許第一列是標頭）
    const header = rows[0].map(h => String(h).trim());
    const COL = {
      company:      header.findIndex(h => h.includes('客戶名稱')),
      product:      header.findIndex(h => h.includes('銷售案名')),
      category:     header.findIndex(h => h.includes('BU')),
      expectedDate: header.findIndex(h => h.includes('預定簽約日')),
      owner:        header.findIndex(h => h.includes('業務帳號')),
      stage:        header.findIndex(h => h.includes('把握度')),
      amount:       header.findIndex(h => h.includes('合約金額')),
      grossMarginRate: header.findIndex(h => h.includes('毛利率')),
      description:  header.findIndex(h => h.includes('備註')),
    };

    const auth = loadAuth();
    const usernames = new Set((auth.users || []).map(u => u.username));
    // displayName → username 反查
    const displayToUser = {};
    (auth.users || []).forEach(u => { displayToUser[u.displayName || u.username] = u.username; });

    const VALID_STAGES = new Set(['A', 'B', 'C', 'D', 'Won']);

    const data = db.load();
    if (!data.opportunities) data.opportunities = [];

    let created = 0;
    const errors = [];

    rows.slice(1).forEach((row, i) => {
      const rowNum = i + 2;
      if (!row.some(c => String(c).trim())) return; // 空列跳過

      const company = String(row[COL.company] ?? '').trim();
      if (!company) { errors.push(`第${rowNum}列：客戶名稱不可空白`); return; }

      // owner 解析：優先用帳號欄，找不到再用姓名欄
      let owner = String(row[COL.owner] ?? '').trim();
      if (!usernames.has(owner)) {
        const byName = displayToUser[owner];
        if (byName) owner = byName;
        else { errors.push(`第${rowNum}列：找不到業務帳號「${owner}」`); return; }
      }

      const stage = String(row[COL.stage] ?? '').trim();
      // 允許中文別名
      const stageMap = { 'Commit': 'A', 'commit': 'A', 'Upside': 'B', 'upside': 'B', 'Pipeline': 'C', 'pipeline': 'C', 'won': 'Won', '成交': 'Won' };
      const resolvedStage = VALID_STAGES.has(stage) ? stage : (stageMap[stage] || 'C');

      const opp = {
        id:             uuidv4(),
        owner,
        company,
        product:        String(row[COL.product]      ?? '').trim(),
        category:       String(row[COL.category]     ?? '').trim(),
        expectedDate:   String(row[COL.expectedDate] ?? '').trim(),
        stage:          resolvedStage,
        amount:         String(row[COL.amount]       ?? '').trim(),
        grossMarginRate:String(row[COL.grossMarginRate] ?? '').trim(),
        description:    String(row[COL.description]  ?? '').trim(),
        contactId:      '',
        contactName:    '',
        visitId:        '',
        createdAt:      new Date().toISOString(),
        importedAt:     new Date().toISOString(),
      };
      data.opportunities.push(opp);
      created++;
    });

    if (created > 0) db.save(data);

    res.json({ success: true, created, errors });
  } catch (err) {
    console.error('[import opp]', err);
    res.status(500).json({ error: '匯入失敗：' + err.message });
  }
});

// ── 日報（拜訪記錄）匯出（Admin）────────────────────────
app.get('/api/admin/visits/export', requireAdmin, (req, res) => {
  const data = db.load();
  const auth = loadAuth();
  const userMap = {};
  (auth.users || []).forEach(u => { userMap[u.username] = u.displayName || u.username; });

  const headers = [
    '拜訪日期', '拜訪方式', '客戶姓名', '拜訪主題', '會談內容', '下一步行動',
    '業務帳號(owner)', '業務姓名', '建立時間', '記錄ID'
  ];

  const rows = (data.visits || []).map(v => [
    v.visitDate    || '',
    v.visitType    || '',
    v.contactName  || '',
    v.topic        || '',
    v.content      || '',
    v.nextAction   || '',
    v.owner        || '',
    userMap[v.owner] || v.owner || '',
    v.createdAt    ? v.createdAt.slice(0, 10) : '',
    v.id           || ''
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = [
    { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 24 }, { wch: 40 }, { wch: 30 },
    { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 36 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, '日報記錄');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const today = new Date();
  const dateStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%E6%97%A5%E5%A0%B1%E8%A8%98%E9%8C%84_${dateStr}.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── 日報（拜訪記錄）批次匯入（Admin）──────────────────────
app.post('/api/admin/visits/import', requireAdmin,
  (req, res, next) => uploadImport.single('file')(req, res, next),
  async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: '請上傳 Excel 檔案' });

      const buf  = req.file.buffer;
      const wb   = XLSX.read(buf, { type: 'buffer' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) return res.status(400).json({ error: '檔案無資料列' });

      const header = rows[0].map(h => String(h).trim());
      const COL = {
        visitDate:   header.findIndex(h => h.includes('拜訪日期')),
        visitType:   header.findIndex(h => h.includes('拜訪方式')),
        contactName: header.findIndex(h => h.includes('客戶姓名')),
        topic:       header.findIndex(h => h.includes('拜訪主題')),
        content:     header.findIndex(h => h.includes('會談內容')),
        nextAction:  header.findIndex(h => h.includes('下一步行動')),
        owner:       header.findIndex(h => h.includes('業務帳號')),
      };

      const auth = loadAuth();
      const usernames = new Set((auth.users || []).map(u => u.username));
      const displayToUser = {};
      (auth.users || []).forEach(u => { displayToUser[u.displayName || u.username] = u.username; });

      const data = db.load();
      if (!data.visits) data.visits = [];

      let created = 0;
      const errors = [];

      rows.slice(1).forEach((row, i) => {
        const rowNum = i + 2;
        if (!row.some(c => String(c).trim())) return;

        // owner 解析
        let owner = String(row[COL.owner] ?? '').trim();
        if (!usernames.has(owner)) {
          const byName = displayToUser[owner];
          if (byName) owner = byName;
          else { errors.push(`第${rowNum}列：找不到業務帳號「${owner}」`); return; }
        }

        const visitDate = String(row[COL.visitDate] ?? '').trim();
        if (!visitDate) { errors.push(`第${rowNum}列：拜訪日期不可空白`); return; }

        const visit = {
          id:          uuidv4(),
          owner,
          contactId:   '',
          contactName: String(row[COL.contactName] ?? '').trim(),
          visitDate,
          visitType:   String(row[COL.visitType]   ?? '').trim() || '其他',
          topic:       String(row[COL.topic]        ?? '').trim(),
          content:     String(row[COL.content]      ?? '').trim(),
          nextAction:  String(row[COL.nextAction]   ?? '').trim(),
          createdAt:   new Date().toISOString(),
          importedAt:  new Date().toISOString(),
        };
        data.visits.push(visit);
        created++;
      });

      if (created > 0) db.save(data);
      res.json({ success: true, created, errors });
    } catch (err) {
      console.error('[import visits]', err);
      res.status(500).json({ error: '匯入失敗：' + err.message });
    }
  }
);

// ── 合約管理匯出（Admin）─────────────────────────────────
app.get('/api/admin/contracts/export', requireAdmin, (req, res) => {
  const data = db.load();
  const auth = loadAuth();
  const userMap = {};
  (auth.users || []).forEach(u => { userMap[u.username] = u.displayName || u.username; });

  const headers = [
    '合約編號', '客戶名稱', '聯絡人', '產品/服務', '合約開始日', '合約結束日',
    '合約金額(萬元)', '業務人員', '類型', '備註',
    '業務帳號(owner)', '業務姓名', '建立時間', '合約ID'
  ];

  const rows = (data.contracts || []).map(c => [
    c.contractNo   || '',
    c.company      || '',
    c.contactName  || '',
    c.product      || '',
    c.startDate    || '',
    c.endDate      || '',
    c.amount       || '',
    c.salesPerson  || '',
    c.type         || '',
    c.note         || '',
    c.owner        || '',
    userMap[c.owner] || c.owner || '',
    c.createdAt    ? c.createdAt.slice(0, 10) : '',
    c.id           || ''
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = [
    { wch: 16 }, { wch: 28 }, { wch: 14 }, { wch: 24 },
    { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 12 },
    { wch: 10 }, { wch: 30 }, { wch: 14 }, { wch: 12 },
    { wch: 12 }, { wch: 36 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, '合約資料');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const today = new Date();
  const dateStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''%E5%90%88%E7%B4%84%E8%B3%87%E6%96%99_${dateStr}.xlsx`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── 合約管理批次匯入（Admin）─────────────────────────────
app.post('/api/admin/contracts/import', requireAdmin,
  (req, res, next) => uploadImport.single('file')(req, res, next),
  async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: '請上傳 Excel 檔案' });

      const buf  = req.file.buffer;
      const wb   = XLSX.read(buf, { type: 'buffer' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) return res.status(400).json({ error: '檔案無資料列' });

      const header = rows[0].map(h => String(h).trim());
      const COL = {
        contractNo:   header.findIndex(h => h.includes('合約編號')),
        company:      header.findIndex(h => h.includes('客戶名稱')),
        contactName:  header.findIndex(h => h.includes('聯絡人')),
        product:      header.findIndex(h => h.includes('產品')),
        startDate:    header.findIndex(h => h.includes('開始日')),
        endDate:      header.findIndex(h => h.includes('結束日')),
        amount:       header.findIndex(h => h.includes('合約金額')),
        salesPerson:  header.findIndex(h => h.includes('業務人員')),
        type:         header.findIndex(h => h.includes('類型')),
        note:         header.findIndex(h => h.includes('備註')),
        owner:        header.findIndex(h => h.includes('業務帳號')),
      };

      const auth = loadAuth();
      const usernames = new Set((auth.users || []).map(u => u.username));
      const displayToUser = {};
      (auth.users || []).forEach(u => { displayToUser[u.displayName || u.username] = u.username; });

      const data = db.load();
      if (!data.contracts) data.contracts = [];

      let created = 0;
      const errors = [];

      rows.slice(1).forEach((row, i) => {
        const rowNum = i + 2;
        if (!row.some(c => String(c).trim())) return;

        const company = String(row[COL.company] ?? '').trim();
        if (!company) { errors.push(`第${rowNum}列：客戶名稱不可空白`); return; }

        let owner = String(row[COL.owner] ?? '').trim();
        if (!usernames.has(owner)) {
          const byName = displayToUser[owner];
          if (byName) owner = byName;
          else { errors.push(`第${rowNum}列：找不到業務帳號「${owner}」`); return; }
        }

        const contract = {
          id:          uuidv4(),
          owner,
          contractNo:  String(row[COL.contractNo]  ?? '').trim(),
          company,
          contactName: String(row[COL.contactName] ?? '').trim(),
          product:     String(row[COL.product]      ?? '').trim(),
          startDate:   String(row[COL.startDate]    ?? '').trim(),
          endDate:     String(row[COL.endDate]      ?? '').trim(),
          amount:      String(row[COL.amount]        ?? '').trim(),
          salesPerson: String(row[COL.salesPerson]  ?? '').trim(),
          type:        String(row[COL.type]          ?? '').trim(),
          note:        String(row[COL.note]          ?? '').trim(),
          createdAt:   new Date().toISOString(),
          importedAt:  new Date().toISOString(),
        };
        data.contracts.push(contract);
        created++;
      });

      if (created > 0) db.save(data);
      res.json({ success: true, created, errors });
    } catch (err) {
      console.error('[import contracts]', err);
      res.status(500).json({ error: '匯入失敗：' + err.message });
    }
  }
);

// ── 合約管理 CRUD ─────────────────────────────────────────
app.get('/api/contracts', requireAuth, (req, res) => {
  const data = db.load();
  const role = req.session.user.role;
  if (role === 'secretary') return res.json([]);
  const owners = getViewableOwners(req, 'contracts');
  res.json((data.contracts || []).filter(c => owners.includes(c.owner)));
});

app.post('/api/contracts', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.contracts) data.contracts = [];
  const c = {
    id:           uuidv4(),
    owner,
    contractNo:   req.body.contractNo   || '',
    company:      req.body.company      || '',
    contactName:  req.body.contactName  || '',
    product:      req.body.product      || '',
    startDate:    req.body.startDate    || '',
    endDate:      req.body.endDate      || '',
    amount:       req.body.amount       || '',
    salesPerson:  req.body.salesPerson  || '',
    note:         req.body.note         || '',
    type:         req.body.type         || 'ERP_MA',
    createdAt:    new Date().toISOString()
  };
  data.contracts.push(c);
  db.save(data);
  writeLog('CREATE_CONTRACT', owner, c.company,
    `合約No:${c.contractNo} 產品:${c.product} 金額:${c.amount}`, req);
  res.status(201).json(c);
});

app.put('/api/contracts/:id', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  const data = db.load();
  if (!data.contracts) data.contracts = [];
  // admin / manager1 可編輯任何人的合約；一般業務只能編自己的
  const canEditAll = ['admin', 'manager1', 'manager2'].includes(role);
  const idx = data.contracts.findIndex(c =>
    c.id === req.params.id && (canEditAll || c.owner === username)
  );
  if (idx === -1) return res.status(404).json({ error: '找不到此合約' });
  const owner = data.contracts[idx].owner; // 保留原 owner，不讓前端改
  data.contracts[idx] = { ...data.contracts[idx], ...pickFields(req.body, CONTRACT_FIELDS), id: req.params.id, owner };
  db.save(data);
  writeLog('UPDATE_CONTRACT', username, data.contracts[idx].company,
    `合約No:${data.contracts[idx].contractNo}`, req);
  res.json(data.contracts[idx]);
});

app.delete('/api/contracts/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.contracts) return res.json({ success: true });
  const c = data.contracts.find(c => c.id === req.params.id && c.owner === owner);
  data.contracts = data.contracts.filter(c => !(c.id === req.params.id && c.owner === owner));
  db.save(data);
  if (c) writeLog('DELETE_CONTRACT', owner, c.company, `合約No:${c.contractNo}`, req);
  res.json({ success: true });
});

// ── 公司查詢快取 ────────────────────────────────────────
let companyCache = { twse: null, tpex: null, ts: 0 };

// ── 財務報表快取（損益表）────────────────────────────────
let finCache = { tse: {}, tpex: {}, dataYear: null, ts: 0 };

// ── Yahoo Finance crumb 快取 ─────────────────────────────
let yahooCache = { cookies: '', crumb: '', ts: 0 };

// 用 Node https 模組抓取，繞開 fetch 的潛在問題
function fetchJsonWithHttps(url, timeoutMs = 20000) {
  return new Promise((resolve, reject) => {
    const https = require('https');
    const http  = require('http');
    const lib   = url.startsWith('https') ? https : http;
    const urlObj = new URL(url);
    let raw = '';
    const req = lib.get({
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }
    }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchJsonWithHttps(res.headers.location, timeoutMs).then(resolve).catch(reject);
      }
      res.setEncoding('utf8');
      res.on('data', chunk => { raw += chunk; });
      res.on('end', () => {
        try {
          const data = JSON.parse(raw);
          resolve(data);
        } catch (e) {
          reject(new Error('JSON parse error: ' + e.message + ' (first 100 chars: ' + raw.slice(0,100) + ')'));
        }
      });
    });
    req.setTimeout(timeoutMs, () => { req.destroy(); reject(new Error('timeout after ' + timeoutMs + 'ms')); });
    req.on('error', reject);
  });
}

async function getCompanyLists() {
  // 只在有效快取（且確實有資料）時才使用快取
  if (companyCache.twse && companyCache.twse.length > 0 && (Date.now() - companyCache.ts) < 86400000) {
    return companyCache;
  }

  const results = await Promise.allSettled([
    fetchJsonWithHttps('https://openapi.twse.com.tw/v1/opendata/t187ap03_L', 20000),
    fetchJsonWithHttps('https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O', 15000)
  ]);

  const twseResult = results[0];
  const tpexResult = results[1];

  if (twseResult.status === 'fulfilled' && Array.isArray(twseResult.value) && twseResult.value.length > 0) {
    companyCache.twse = twseResult.value;
    console.log('[TWSE] 載入成功，', companyCache.twse.length, '家上市公司');
  } else {
    // 失敗時不更新 twse，保留舊快取（若有），只記錄錯誤
    console.log('[TWSE] 載入失敗:', twseResult.reason?.message || '未知錯誤');
    if (!companyCache.twse) companyCache.twse = [];
  }

  if (tpexResult.status === 'fulfilled' && Array.isArray(tpexResult.value) && tpexResult.value.length > 0) {
    companyCache.tpex = tpexResult.value;
    console.log('[TPEX] 載入成功，', companyCache.tpex.length, '家上櫃公司');
  } else {
    console.log('[TPEX] 載入失敗（可能為正常，將改用名稱比對）:', tpexResult.reason?.message || '未知錯誤');
    if (!companyCache.tpex) companyCache.tpex = [];
  }

  // 只在 TWSE 成功時才更新快取時間戳（避免失敗快取）
  if (companyCache.twse.length > 0) {
    companyCache.ts = Date.now();
  }

  return companyCache;
}

// ── 上市/上櫃公司詳細資料快取（含公司網址欄位）──────────
let _listedDetailsCache = null;
let _listedDetailsCachedAt = 0;
const LISTED_DETAILS_TTL = 24 * 3600 * 1000;

async function getListedCompanyDetails() {
  const now = Date.now();
  if (_listedDetailsCache && now - _listedDetailsCachedAt < LISTED_DETAILS_TTL) {
    return _listedDetailsCache;
  }
  const cache = { byStock: {} };

  // TWSE 上市
  try {
    const r = await fetch('https://openapi.twse.com.tw/v1/opendata/t187ap03_L', {
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' },
      signal: AbortSignal.timeout(15000)
    });
    if (r.ok) {
      const arr = await r.json();
      for (const it of arr) {
        const code = (it['公司代號'] || '').toString().trim();
        const web  = (it['公司網址'] || '').trim();
        if (code) cache.byStock[code] = web;
      }
    }
  } catch (e) { console.warn('[TWSE OpenAPI]', e.message); }

  // TPEX 上櫃
  try {
    const r = await fetch('https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O', {
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' },
      signal: AbortSignal.timeout(15000)
    });
    if (r.ok) {
      const arr = await r.json();
      for (const it of arr) {
        const code = (it['公司代號'] || it.SecuritiesCompanyCode || '').toString().trim();
        const web  = (it['公司網址'] || it.CompanyWebsite || '').trim();
        if (code) cache.byStock[code] = web;
      }
    }
  } catch (e) { console.warn('[TPEX OpenAPI]', e.message); }

  _listedDetailsCache = cache;
  _listedDetailsCachedAt = now;
  return cache;
}

// ── 共用：只查官網（給 bulk 及 company-lookup 使用）──────
async function fetchWebsiteOnly(taxId, companyName, stockCode, listedType) {
  let website = '';
  let emailDomain = '';
  const debug = [];

  // 1. TWSE/TPEX OpenAPI（上市/上櫃，公司網址欄位）
  if (!website && stockCode) {
    try {
      const det = await getListedCompanyDetails();
      const raw = (det.byStock[stockCode] || '').trim();
      if (raw && raw.length > 4) {
        website = (/^https?:\/\//i.test(raw) ? raw : 'https://' + raw).replace(/\/$/, '');
        debug.push('TWSE/TPEX OpenAPI ✓');
      } else { debug.push('TWSE/TPEX 無資料'); }
    } catch (e) { debug.push('TWSE/TPEX err: ' + e.message); }
  }

  // 2. GCIS 通訊資料（有統編才查）
  if (!website && taxId) {
    try {
      const r = await fetch(
        `https://data.gcis.nat.gov.tw/od/data/api/9A6764F8-C567-4B97-985A-B2FFA47A7B4F?$format=json&$filter=Business_Accounting_NO eq ${taxId}&$skip=0&$top=1`,
        { headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }, signal: AbortSignal.timeout(6000) }
      );
      if (r.ok) {
        const d = await r.json();
        if (d?.[0]) {
          const raw = d[0].Company_Website || d[0].Website || d[0].website || '';
          if (raw && raw.length > 4) {
            website = (/^https?:\/\//i.test(raw) ? raw : 'https://' + raw).replace(/\/$/, '');
            debug.push('GCIS ✓');
          }
          const email = d[0].Company_Email || d[0].E_Mail || '';
          if (!website && email && email.includes('@')) {
            const dom = email.split('@')[1].toLowerCase().trim();
            const PUBLIC = ['gmail.com','yahoo.com','yahoo.com.tw','hotmail.com','outlook.com'];
            if (dom && !PUBLIC.includes(dom)) emailDomain = 'https://www.' + dom;
          }
        } else { debug.push('GCIS 查無資料'); }
      } else { debug.push('GCIS HTTP ' + r.status); }
    } catch (e) { debug.push('GCIS err: ' + e.message); }
  }

  // 3. DuckDuckGo Instant Answer
  if (!website && companyName) {
    try {
      const q = encodeURIComponent(companyName + ' 官方網站');
      const r = await fetch(
        `https://api.duckduckgo.com/?q=${q}&format=json&no_redirect=1&no_html=1&skip_disambig=1`,
        { headers: { 'User-Agent': 'Mozilla/5.0' }, signal: AbortSignal.timeout(5000) }
      );
      if (r.ok) {
        const ddg = await r.json();
        const webItem = (ddg.Infobox?.content || []).find(
          item => /^(website|official website|官網|網址|homepage)/i.test(item.label || '')
        );
        if (webItem?.value && /^https?:\/\//i.test(webItem.value)) {
          website = webItem.value.replace(/\/+$/, '');
          debug.push('DDG Infobox ✓');
        } else if (ddg.AbstractURL && !/wikipedia|wikimedia/i.test(ddg.AbstractURL)) {
          website = ddg.AbstractURL;
          debug.push('DDG Abstract ✓');
        } else { debug.push('DDG 無相符結果'); }
      } else { debug.push('DDG HTTP ' + r.status); }
    } catch (e) { debug.push('DDG err: ' + e.message); }
  }

  // 4. Email domain 備援
  if (!website && emailDomain) {
    website = emailDomain;
    debug.push('Email domain ✓');
  }

  return { website, debug: debug.join(' / ') };
}

function formatCapital(amount) {
  const num = parseInt(amount);
  if (!num || isNaN(num)) return '';
  if (num >= 100000000) return `NT$ ${(num / 100000000).toFixed(1)} 億`;
  if (num >= 10000) return `NT$ ${(num / 10000).toFixed(0)} 萬`;
  return `NT$ ${num.toLocaleString()}`;
}

// ── 取得 TWSE + TPEX 損益表快取 ─────────────────────────────
async function getFinancialLists() {
  const now = Date.now();
  if (finCache.ts > 0 && (now - finCache.ts) < 86400000 && Object.keys(finCache.tse).length > 0) {
    return finCache;
  }

  const [tseResult, tpexResult] = await Promise.allSettled([
    fetchJsonWithHttps('https://openapi.twse.com.tw/v1/opendata/t187ap06_L_ci', 30000),
    fetchJsonWithHttps('https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_ci', 25000)
  ]);

  function indexByCode(rows, codeField, revField, gpFields, yearField, yearIsROC) {
    const map = {};
    (rows || []).forEach(row => {
      const code = (row[codeField] || '').replace(/\*/g, '').trim();
      if (!code) return;
      const rev = parseFloat(row[revField]) || 0;
      const gp = parseFloat(gpFields.map(f => row[f]).find(v => v && parseFloat(v) !== 0) || '0') || 0;
      const rawYear = parseInt(row[yearField]) || 0;
      const westYear = yearIsROC ? rawYear + 1911 : rawYear;
      const epsRaw = parseFloat(row['基本每股盈餘（元）']) || null;
      map[code] = { year: westYear, revRaw: rev, gpRaw: gp,
        grossMargin: rev > 0 ? ((gp / rev) * 100).toFixed(1) + '%' : 'N/A',
        eps: epsRaw !== null ? epsRaw.toFixed(2) : 'N/A' };
    });
    return map;
  }

  if (tseResult.status === 'fulfilled' && Array.isArray(tseResult.value) && tseResult.value.length > 0) {
    finCache.tse = indexByCode(tseResult.value, '公司代號', '營業收入',
      ['營業毛利（毛損）淨額', '營業毛利（毛損）'], '年度', true);
    const sample = Object.values(finCache.tse)[0];
    if (sample) finCache.dataYear = sample.year;
    console.log('[FinCache] TSE 損益表載入成功，', Object.keys(finCache.tse).length, '家，年度:', finCache.dataYear);
  } else {
    console.log('[FinCache] TSE 損益表載入失敗:', tseResult.reason?.message || '未知');
  }

  if (tpexResult.status === 'fulfilled' && Array.isArray(tpexResult.value) && tpexResult.value.length > 0) {
    finCache.tpex = indexByCode(tpexResult.value, 'SecuritiesCompanyCode', '營業收入',
      ['營業毛利（毛損）淨額', '營業毛利（毛損）'], 'Year', true);
    console.log('[FinCache] TPEX 損益表載入成功，', Object.keys(finCache.tpex).length, '家');
  } else {
    console.log('[FinCache] TPEX 損益表載入失敗:', tpexResult.reason?.message || '未知');
  }

  if (Object.keys(finCache.tse).length > 0) finCache.ts = now;
  return finCache;
}

// ── 取得 Yahoo Finance crumb（用於查詢前一年營收）──────────────
async function getYahooCrumb() {
  const now = Date.now();
  if (yahooCache.crumb && (now - yahooCache.ts) < 3600000) return yahooCache;

  return new Promise(resolve => {
    const https = require('https');
    let cookieStr = '';
    const req1 = https.get({
      hostname: 'finance.yahoo.com',
      path: '/quote/2330.TW',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120',
        'Accept': 'text/html', 'Accept-Language': 'en-US,en;q=0.9'
      }
    }, r1 => {
      cookieStr = (r1.headers['set-cookie'] || []).map(c => c.split(';')[0]).join('; ');
      r1.resume(); // drain
      r1.on('end', () => {
        let raw2 = '';
        const req2 = https.get({
          hostname: 'query2.finance.yahoo.com', path: '/v1/test/getcrumb',
          headers: { 'User-Agent': 'Mozilla/5.0', 'Cookie': cookieStr }
        }, r2 => {
          r2.on('data', c => raw2 += c);
          r2.on('end', () => {
            if (raw2 && raw2.length < 50) {
              yahooCache = { cookies: cookieStr, crumb: raw2.trim(), ts: Date.now() };
              console.log('[Yahoo] crumb 更新成功');
            }
            resolve(yahooCache);
          });
        });
        req2.setTimeout(8000, () => { req2.destroy(); resolve(yahooCache); });
        req2.on('error', () => resolve(yahooCache));
      });
    });
    req1.setTimeout(10000, () => { req1.destroy(); resolve(yahooCache); });
    req1.on('error', () => resolve(yahooCache));
  });
}

// ── 查詢財務數據：年度一用 TWSE/TPEX 官方 API，年度二用 Yahoo Finance ──
async function fetchFinancialData(stockCode, year, exchange) {
  const result = { revenue: '無法取得', grossMargin: '無法取得', eps: 'N/A' };
  try {
    const fins = await getFinancialLists();
    const lookup = exchange === 'OTC' ? fins.tpex[stockCode] : fins.tse[stockCode];

    if (lookup && lookup.year === year) {
      // 官方 API 有此年度資料（完整：含毛利率＋EPS）
      result.revenue = formatCapital(lookup.revRaw * 1000);
      result.grossMargin = lookup.grossMargin;
      result.eps = lookup.eps;
      return result;
    }

    // 前一年度：嘗試 Yahoo Finance（只有營收，無毛利率）
    const suffix = exchange === 'OTC' ? '.TWO' : '.TW';
    const ticker = stockCode + suffix;
    const yahoo = await getYahooCrumb();
    if (!yahoo.crumb) return result;

    const url = `https://query1.finance.yahoo.com/v10/finance/quoteSummary/${ticker}?modules=incomeStatementHistory&crumb=${encodeURIComponent(yahoo.crumb)}`;
    const data = await fetchJsonWithHttps(url, 10000);
    const stmts = data?.quoteSummary?.result?.[0]?.incomeStatementHistory?.incomeStatementHistory;
    if (stmts) {
      const match = stmts.find(s => s.endDate?.fmt?.startsWith(String(year)));
      if (match?.totalRevenue?.raw) {
        result.revenue = formatCapital(match.totalRevenue.raw);
        result.grossMargin = 'N/A';
      }
    }
  } catch (e) { /* ignore */ }
  return result;
}

// ── 公司查詢 API ─────────────────────────────────────────
app.get('/api/company-lookup', requireAuth, async (req, res) => {
  const { taxId } = req.query;
  if (!taxId || !/^\d{8}$/.test(taxId)) return res.status(400).json({ error: '請輸入正確的統一編號（8碼數字）' });

  const result = {
    companyName: '', representative: '', capital: '', address: '', companyStatus: '',
    listedType: '未上市櫃', stockCode: '', exchange: '', website: '',
    revenue2025: 'N/A', grossMargin2025: 'N/A',
    revenue2024: 'N/A', grossMargin2024: 'N/A',
    eps: 'N/A', epsYear: '',
  };

  // 1. 經濟部商業司（基本資料 + 通訊資料）
  try {
    const r = await fetch(
      `https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq ${taxId}&$skip=0&$top=1`,
      { headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }, signal: AbortSignal.timeout(8000) }
    );
    if (r.ok) {
      const data = await r.json();
      if (data?.[0]) {
        result.companyName   = data[0].Company_Name || '';
        result.representative = data[0].Responsible_Name || '';
        result.capital       = formatCapital(data[0].Capital_Stock_Amount);
        result.address       = data[0].Company_Location || '';
        result.companyStatus = data[0].Company_Status_Desc || '';
        // GCIS 有時含 Email 欄位，可推算官網 domain
        const gcisEmail = data[0].Company_Email || data[0].E_Mail || '';
        if (gcisEmail && gcisEmail.includes('@')) {
          const domain = gcisEmail.split('@')[1].toLowerCase().trim();
          if (domain && !domain.includes('gmail') && !domain.includes('yahoo') &&
              !domain.includes('hotmail') && !domain.includes('outlook')) {
            result._emailDomain = 'https://www.' + domain; // 暫存，稍後用作備援
          }
        }
      }
    }
  } catch (e) { /* ignore */ }

  // 1.5 GCIS 通訊資料（含公司網址欄位）
  try {
    const r2 = await fetch(
      `https://data.gcis.nat.gov.tw/od/data/api/9A6764F8-C567-4B97-985A-B2FFA47A7B4F?$format=json&$filter=Business_Accounting_NO eq ${taxId}&$skip=0&$top=1`,
      { headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }, signal: AbortSignal.timeout(6000) }
    );
    if (r2.ok) {
      const d2 = await r2.json();
      if (d2?.[0]) {
        const rawWeb = d2[0].Company_Website || d2[0].Website || d2[0].website || '';
        if (rawWeb && rawWeb.length > 4) {
          result.website = /^https?:\/\//i.test(rawWeb) ? rawWeb : 'https://' + rawWeb;
          result.website = result.website.replace(/\/$/, '');
        }
        // 備援 email domain
        if (!result._emailDomain) {
          const email2 = d2[0].Company_Email || d2[0].E_Mail || '';
          if (email2 && email2.includes('@')) {
            const domain = email2.split('@')[1].toLowerCase().trim();
            if (domain && !['gmail.com','yahoo.com','yahoo.com.tw','hotmail.com','outlook.com'].includes(domain)) {
              result._emailDomain = 'https://www.' + domain;
            }
          }
        }
      }
    }
  } catch { /* ignore */ }

  // 2. 上市/上櫃判斷
  try {
    const lists = await getCompanyLists();

    // 優先：TWSE（上市）- 直接用統一編號比對（欄位名稱：營利事業統一編號）
    console.log('[Lookup] taxId:', taxId, '| TWSE筆數:', (lists.twse||[]).length, '| TPEX筆數:', (lists.tpex||[]).length);
    const twseMatch = (lists.twse || []).find(c => c['營利事業統一編號'] === taxId);
    console.log('[Lookup] TWSE比對結果:', twseMatch ? twseMatch['公司代號'] + ' ' + twseMatch['公司名稱'] : '未找到');
    if (twseMatch) {
      result.listedType = '上市';
      result.stockCode = twseMatch['公司代號'] || '';
      result.exchange = 'TSE';
      // 補強公司名稱（若 GCIS 未取到）
      if (!result.companyName) result.companyName = twseMatch['公司名稱'] || '';
      if (!result.representative) result.representative = twseMatch['董事長'] || '';
      if (!result.address) result.address = twseMatch['住址'] || '';
    }

    // 其次：TPEX（上櫃）- 直接用統一編號比對
    if (!result.stockCode) {
      const tpexMatch = (lists.tpex || []).find(c => c['UnifiedBusinessNo.'] === taxId);
      if (tpexMatch) {
        result.listedType = '上櫃';
        result.stockCode = tpexMatch.SecuritiesCompanyCode || '';
        result.exchange = 'OTC';
        if (!result.companyName) result.companyName = tpexMatch.CompanyName || '';
        if (!result.representative) result.representative = tpexMatch.Chairman || '';
        if (!result.address) result.address = tpexMatch.Address || '';
      }
    }

    // 最後備援：TWSE 名稱比對（處理統一編號不一致的特殊情況）
    if (!result.stockCode && result.companyName) {
      const twseFallback = (lists.twse || []).find(c => {
        const full = (c['公司名稱'] || '').trim();
        const abbr = (c['公司簡稱'] || '').replace(/[*＊]/g, '').trim();
        return full === result.companyName ||
          (abbr.length >= 2 && result.companyName.includes(abbr));
      });
      if (twseFallback) {
        result.listedType = '上市';
        result.stockCode = twseFallback['公司代號'] || '';
        result.exchange = 'TSE';
      }
    }
  } catch (e) { /* ignore */ }

  // 2.5 抓取公司官網（TWSE/TPEX OpenAPI - 含公司網址欄位）
  if (!result.website && result.stockCode) {
    try {
      const det = await getListedCompanyDetails();
      const raw = (det.byStock[result.stockCode] || '').trim();
      if (raw && raw.length > 4) {
        result.website = (/^https?:\/\//i.test(raw) ? raw : 'https://' + raw).replace(/\/$/, '');
      }
    } catch { /* ignore */ }
  }

  // 2.6 未上市/上櫃備援：DuckDuckGo Instant Answer（免費，無需 API Key）
  if (!result.website && result.companyName) {
    try {
      const query = encodeURIComponent(result.companyName + ' 官方網站');
      const ddgR = await fetch(
        `https://api.duckduckgo.com/?q=${query}&format=json&no_redirect=1&no_html=1&skip_disambig=1`,
        { headers: { 'User-Agent': 'Mozilla/5.0' }, signal: AbortSignal.timeout(5000) }
      );
      if (ddgR.ok) {
        const ddg = await ddgR.json();
        // 從 Infobox 找 Website 欄位（知名企業通常有）
        const infoItems = ddg.Infobox?.content || [];
        const webItem = infoItems.find(item =>
          /^(website|official website|官網|網址|home ?page)/i.test(item.label || '')
        );
        if (webItem?.value && /^https?:\/\//i.test(webItem.value)) {
          result.website = webItem.value.replace(/\/+$/, '');
        }
        // 備援：AbstractURL 非 wikipedia 則使用
        if (!result.website && ddg.AbstractURL &&
            !/wikipedia|wikimedia/i.test(ddg.AbstractURL)) {
          result.website = ddg.AbstractURL;
        }
      }
    } catch { /* ignore */ }
  }

  // 2.7 最終備援：使用 GCIS Email 推算 domain（已知不是公版信箱）
  if (!result.website && result._emailDomain) {
    result.website = result._emailDomain;
  }
  delete result._emailDomain; // 清除暫存欄位

  // 3. 財務數據（僅上市/上櫃）
  if (result.stockCode) {
    const fins = await getFinancialLists();
    const currentYear = fins.dataYear || (new Date().getFullYear() - 1);
    const prevYear = currentYear - 1;
    result.dataYear1 = currentYear;
    result.dataYear2 = prevYear;

    const [fin1, fin2] = await Promise.allSettled([
      fetchFinancialData(result.stockCode, currentYear, result.exchange),
      fetchFinancialData(result.stockCode, prevYear, result.exchange)
    ]);
    if (fin1.status === 'fulfilled') {
      result.revenue2025 = fin1.value.revenue;
      result.grossMargin2025 = fin1.value.grossMargin;
      if (fin1.value.eps && fin1.value.eps !== 'N/A') {
        result.eps = fin1.value.eps;
        result.epsYear = String(currentYear);
      }
    }
    if (fin2.status === 'fulfilled') { result.revenue2024 = fin2.value.revenue; result.grossMargin2024 = fin2.value.grossMargin; }
  }

  apiMonitor.recordCompanyLookup({
    gcisSuccess:     result.companyName ? 1 : 0,
    gcisError:       result.companyName ? 0 : 1,
    twseTpexSuccess: result.listedType !== '未上市櫃' ? 1 : 0,
    ddgSuccess:      0,
    ddgError:        0,
  });
  res.json(result);
});

// ── 可視用戶名稱對應表（供前端顯示業務人員姓名）──────────────
app.get('/api/usermap', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  const auth = loadAuth();
  let visibleRoles;
  if (role === 'manager1') {
    visibleRoles = ['user', 'manager2', 'manager1'];
  } else if (role === 'manager2') {
    visibleRoles = ['user', 'manager2'];
  } else if (role === 'secretary') {
    visibleRoles = ['user', 'manager2', 'manager1'];
  } else {
    visibleRoles = [];
  }
  const map = {};
  // 自己一定包含
  const me = auth.users.find(u => u.username === username);
  if (me) map[me.username] = me.displayName || me.username;
  // 加入可見角色
  auth.users.filter(u => visibleRoles.includes(u.role)).forEach(u => {
    map[u.username] = u.displayName || u.username;
  });
  res.json(map);
});

// ── 通知輔助 ─────────────────────────────────────────────
function pushNotification(toUsername, type, title, body, refId) {
  const data = db.load();
  if (!data.notifications) data.notifications = [];
  data.notifications.unshift({
    id: uuidv4(), to: toUsername, type, title, body, refId,
    read: false, createdAt: new Date().toISOString()
  });
  if (data.notifications.length > 500) data.notifications = data.notifications.slice(0, 500);
  db.save(data);
}

// ── 合約到期提醒 API ────────────────────────────────────
app.get('/api/contract-reminders', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  if (role === 'secretary') return res.json([]);   // 秘書不看合約

  const data = db.load();
  const auth = loadAuth();

  // 可視合約範圍：manager1 看全部，manager2 看自己+業務，user 看自己
  let owners;
  if (role === 'manager1') {
    owners = auth.users.map(u => u.username);
  } else if (role === 'manager2') {
    owners = auth.users.filter(u => u.role === 'user' || u.username === username).map(u => u.username);
  } else {
    owners = [username];
  }

  const contracts = (data.contracts || []).filter(c => owners.includes(c.owner));

  const today = new Date(); today.setHours(0, 0, 0, 0);

  // ── 套用與前端相同的 contractStatus 邏輯 ──
  const calcStatus = (c) => {
    if (!c.endDate) return null;
    const end     = new Date(c.endDate);
    const endDiff = Math.ceil((end - today) / 86400000);
    let effectiveEnd, isRenewed = false;

    if (endDiff < 0 && c.renewDate) {
      effectiveEnd = new Date(c.renewDate);
      isRenewed    = true;
    } else if (endDiff < 0) {
      return { key: 'expired', days: Math.abs(endDiff), isRenewed: false };
    } else {
      effectiveEnd = end;
    }

    const diff = Math.ceil((effectiveEnd - today) / 86400000);
    if (diff <= 25)  return { key: 'urgent',   days: diff, isRenewed };
    if (diff <= 90)  return { key: 'expiring', days: diff, isRenewed };
    return null;   // 有效，不需要提醒
  };

  const reminders = [];
  contracts.forEach(c => {
    const st = calcStatus(c);
    if (!st) return;

    const ownerUser = auth.users.find(u => u.username === c.owner);
    const ownerName = ownerUser?.displayName || c.owner;

    let title, body, icon;
    if (st.key === 'expired') {
      icon  = '🔴';
      title = `合約逾期：${c.company}`;
      body  = `${c.product || '合約'} 已逾期 ${st.days} 天，請盡速處理${role !== 'user' ? `（業務：${ownerName}）` : ''}`;
    } else if (st.key === 'urgent') {
      icon  = '🟠';
      title = `合約即將到期：${c.company}`;
      body  = `${c.product || '合約'} 剩餘 ${st.days} 天${st.isRenewed ? '（已續約）' : ''}到期${role !== 'user' ? `，業務：${ownerName}` : ''}`;
    } else {
      icon  = '🟡';
      title = `合約 90 天內到期：${c.company}`;
      body  = `${c.product || '合約'} 剩餘 ${st.days} 天${st.isRenewed ? '（已續約）' : ''}到期${role !== 'user' ? `，業務：${ownerName}` : ''}`;
    }

    reminders.push({
      id:         `contract_${c.id}`,
      type:       `contract_${st.key}`,
      title:      `${icon} ${title}`,
      body,
      days:       st.days,
      isRenewed:  st.isRenewed,
      contractId: c.id,
      company:    c.company,
      product:    c.product || '',
      endDate:    c.endDate,
      renewDate:  c.renewDate || null,
      ownerName,
    });
  });

  // 排序：逾期 → urgent → expiring，同類型依天數升冪
  const ORDER = { expired: 0, contract_expired: 0, contract_urgent: 1, contract_expiring: 2 };
  reminders.sort((a, b) => (ORDER[a.type] ?? 9) - (ORDER[b.type] ?? 9) || a.days - b.days);

  res.json(reminders);
});

// ── 通知 API ─────────────────────────────────────────────
app.get('/api/notifications', requireAuth, (req, res) => {
  const username = req.session.user.username;
  const data = db.load();
  const list = (data.notifications || []).filter(n => n.to === username);
  res.json(list);
});

app.put('/api/notifications/:id/read', requireAuth, (req, res) => {
  const username = req.session.user.username;
  const data = db.load();
  if (!data.notifications) return res.json({ success: true });
  const n = data.notifications.find(n => n.id === req.params.id && n.to === username);
  if (n) { n.read = true; db.save(data); }
  res.json({ success: true });
});

app.put('/api/notifications/read-all', requireAuth, (req, res) => {
  const username = req.session.user.username;
  const data = db.load();
  if (!data.notifications) return res.json({ success: true });
  data.notifications.filter(n => n.to === username).forEach(n => n.read = true);
  db.save(data);
  res.json({ success: true });
});

// ── 帳務管理：應收帳款 CRUD ───────────────────────────────
app.get('/api/receivables', requireAuth, (req, res) => {
  const role = req.session.user.role;
  if (role === 'secretary') return res.json([]);
  const owners = getViewableOwners(req, 'receivables');
  const data = db.load();
  res.json((data.receivables || []).filter(r => owners.includes(r.owner)));
});

app.post('/api/receivables', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.receivables) data.receivables = [];
  const item = {
    id: uuidv4(), owner,
    company:     req.body.company     || '',
    contactName: req.body.contactName || '',
    invoiceNo:   req.body.invoiceNo   || '',
    invoiceDate: req.body.invoiceDate || '',
    dueDate:     req.body.dueDate     || '',
    amount:      parseFloat(req.body.amount) || 0,
    paidAmount:  parseFloat(req.body.paidAmount) || 0,
    currency:    req.body.currency    || 'NTD',
    note:        req.body.note        || '',
    status:      'pending',
    createdAt:   new Date().toISOString()
  };
  data.receivables.push(item);
  db.save(data);
  writeLog('CREATE_RECEIVABLE', owner, item.company,
    `發票:${item.invoiceNo} 金額:${item.amount}`, req);
  res.status(201).json(item);
});

app.put('/api/receivables/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  const idx = (data.receivables || []).findIndex(r => r.id === req.params.id && r.owner === owner);
  if (idx === -1) return res.status(404).json({ error: '找不到此帳款' });
  data.receivables[idx] = { ...data.receivables[idx], ...pickFields(req.body, RECEIVABLE_FIELDS), id: req.params.id, owner };
  if (req.body.amount !== undefined) data.receivables[idx].amount = parseFloat(req.body.amount) || 0;
  if (req.body.paidAmount !== undefined) data.receivables[idx].paidAmount = parseFloat(req.body.paidAmount) || 0;
  db.save(data);
  writeLog('UPDATE_RECEIVABLE', owner, data.receivables[idx].company,
    `發票:${data.receivables[idx].invoiceNo} 狀態:${data.receivables[idx].status}`, req);
  res.json(data.receivables[idx]);
});

app.delete('/api/receivables/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  const r = (data.receivables || []).find(r => r.id === req.params.id && r.owner === owner);
  data.receivables = (data.receivables || []).filter(r => !(r.id === req.params.id && r.owner === owner));
  db.save(data);
  if (r) writeLog('DELETE_RECEIVABLE', owner, r.company, `發票:${r.invoiceNo}`, req);
  res.json({ success: true });
});

// ── Call-in Pass CRUD ─────────────────────────────────────
// 取得 Call-in 列表（角色可視）
app.get('/api/callins', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  const data = db.load();
  if (!data.callins) data.callins = [];

  // 逾期自動更新
  const now = new Date();
  let changed = false;
  data.callins.forEach(c => {
    if (c.status === 'assigned' && c.deadline && new Date(c.deadline) < now) {
      c.status = 'overdue';
      changed = true;
      // 推播通知一二級主管
      const auth = loadAuth();
      auth.users.filter(u => u.role === 'manager1' || u.role === 'manager2').forEach(u => {
        pushNotification(u.username, 'callin_overdue',
          '⏰ Call-in 逾時未處理',
          `${c.company || c.contactName} 的 Call-in 已逾時，業務：${c.assignedTo}`, c.id);
      });
    }
  });
  if (changed) db.save(data);

  let list;
  if (role === 'secretary' || role === 'manager1' || role === 'admin') {
    list = data.callins; // 全部
  } else if (role === 'manager2') {
    // 看全部 user 及 manager2 的
    const auth = loadAuth();
    const visibleOwners = auth.users.filter(u => u.role === 'user' || u.role === 'manager2').map(u => u.username);
    list = data.callins.filter(c => visibleOwners.includes(c.createdBy) || visibleOwners.includes(c.assignedTo) || c.createdBy === username);
  } else {
    // user：只看指派給自己的
    list = data.callins.filter(c => c.assignedTo === username || c.createdBy === username);
  }
  res.json(list.sort((a, b) => b.createdAt.localeCompare(a.createdAt)));
});

// 秘書建立新 Call-in
app.post('/api/callins', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  if (role !== 'secretary' && role !== 'admin' && role !== 'manager1' && role !== 'manager2') {
    return res.status(403).json({ error: '只有秘書或主管可以建立 Call-in' });
  }
  const data = db.load();
  if (!data.callins) data.callins = [];
  const item = {
    id: uuidv4(), createdBy: username,
    company:     req.body.company     || '',
    contactName: req.body.contactName || '',
    phone:       req.body.phone       || '',
    topic:       req.body.topic       || '',
    source:      req.body.source      || '',
    note:        req.body.note        || '',
    status:      'pending',
    assignedTo:  null, assignedBy: null, assignedAt: null,
    deadline:    null, contactedAt: null, opportunityId: null,
    createdAt:   new Date().toISOString()
  };
  data.callins.push(item);
  db.save(data);
  // 通知二級主管
  const auth = loadAuth();
  auth.users.filter(u => u.role === 'manager2' || u.role === 'manager1').forEach(u => {
    pushNotification(u.username, 'callin_new', '📞 新 Call-in Pass',
      `${item.company || item.contactName} 來電，請指派業務`, item.id);
  });
  res.status(201).json(item);
});

// 主管指派 Call-in 給業務
app.put('/api/callins/:id/assign', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  if (role !== 'manager1' && role !== 'manager2' && role !== 'admin') {
    return res.status(403).json({ error: '無指派權限' });
  }
  const { assignedTo } = req.body;
  if (!assignedTo) return res.status(400).json({ error: '請指定業務人員' });
  const data = db.load();
  const item = (data.callins || []).find(c => c.id === req.params.id);
  if (!item) return res.status(404).json({ error: '找不到此 Call-in' });

  const now = new Date();
  const deadline = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59).toISOString();
  item.assignedTo = assignedTo;
  item.assignedBy = username;
  item.assignedAt = now.toISOString();
  item.deadline   = deadline;
  item.status     = 'assigned';
  db.save(data);

  // 通知被指派的業務
  pushNotification(assignedTo, 'callin_assigned', '📞 您有新的 Call-in 指派',
    `${item.company || item.contactName} 來電，請於今日完成聯繫`, item.id);

  // 也同步加入潛在客戶（若公司/聯絡人有填）
  if (item.company || item.contactName) {
    const auth = loadAuth();
    const assignedUser = auth.users.find(u => u.username === assignedTo);
    if (assignedUser) {
      if (!data.contacts) data.contacts = [];
      // 避免重複
      const exists = data.contacts.find(c => c.owner === assignedTo &&
        ((item.company && c.company === item.company) || (item.contactName && c.name === item.contactName)));
      if (!exists) {
        data.contacts.push({
          id: uuidv4(), owner: assignedTo,
          name: item.contactName || '', company: item.company || '',
          phone: item.phone || '', email: '', title: '', mobile: '', ext: '',
          address: '', website: '', taxId: '', industry: '', note: `[Call-in] ${item.topic}`,
          opportunityStage: 'prospect', isPrimary: false,
          systemVendor: '', systemProduct: '', cardImage: '',
          createdAt: new Date().toISOString(), fromCallIn: item.id
        });
        db.save(data);
      }
    }
  }
  res.json(item);
});

// 業務回應 Call-in（聯繫完成 / 建立商機 / 不合格）
app.put('/api/callins/:id/respond', requireAuth, (req, res) => {
  const { username } = req.session.user;
  const data = db.load();
  const item = (data.callins || []).find(c => c.id === req.params.id && c.assignedTo === username);
  if (!item) return res.status(404).json({ error: '找不到此 Call-in 或無權限' });

  const { action, opportunityName, opportunityStage, note } = req.body;
  item.contactedAt = new Date().toISOString();
  if (note) item.responseNote = note;

  if (action === 'qualified') {
    item.status = 'qualified';
    // 建立商機
    if (!data.opportunities) data.opportunities = [];
    const opp = {
      id: uuidv4(), owner: username,
      contactId: '', contactName: item.contactName, company: item.company,
      category: 'Call-in', product: opportunityName || item.topic,
      amount: '', expectedDate: '', description: item.topic,
      stage: opportunityStage || 'C', visitId: '',
      createdAt: new Date().toISOString()
    };
    data.opportunities.push(opp);
    item.opportunityId = opp.id;
  } else if (action === 'unqualified') {
    item.status = 'unqualified';
  } else {
    item.status = 'contacted';
  }
  db.save(data);

  // 通知指派主管
  if (item.assignedBy) {
    const statusLabel = { qualified: '已建立商機 ✅', unqualified: '非合格商機 ❌', contacted: '已完成聯繫' };
    pushNotification(item.assignedBy, 'callin_responded', '📞 Call-in 已回覆',
      `${item.company || item.contactName}：${statusLabel[item.status] || ''}`, item.id);
  }
  res.json(item);
});

// ════════════════════════════════════════════════════════
//  行銷管理：活動（campaigns）& 線索（leads）
// ════════════════════════════════════════════════════════

const CAMPAIGN_FIELDS = ['name','type','startDate','endDate','description','budget','targetCount','status'];
const LEAD_FIELDS     = ['campaignId','campaignName','company','contactName','title','phone','email','interest','note','status'];

// ── 行銷活動 CRUD ─────────────────────────────────────

app.get('/api/campaigns', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  const data = db.load();
  let list = data.campaigns || [];
  if (role === 'marketing') {
    list = list.filter(c => c.owner === username);
  } else if (!['admin','manager1','manager2'].includes(role)) {
    return res.json([]);
  }
  // 附加每個活動的 lead 統計
  const leads = data.leads || [];
  list = list.map(c => ({
    ...c,
    leadCount:     leads.filter(l => l.campaignId === c.id).length,
    convertedCount: leads.filter(l => l.campaignId === c.id && l.status === 'converted').length,
  }));
  res.json(list.sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || '')));
});

app.post('/api/campaigns', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  if (!['marketing','admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '無權限' });
  const data = db.load();
  if (!data.campaigns) data.campaigns = [];
  const c = { id: uuidv4(), owner: username, ...pickFields(req.body, CAMPAIGN_FIELDS), createdAt: new Date().toISOString() };
  if (!c.name) return res.status(400).json({ error: '請填入活動名稱' });
  c.status = c.status || 'planned';
  data.campaigns.push(c);
  db.save(data);
  res.status(201).json(c);
});

app.put('/api/campaigns/:id', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  const data = db.load();
  const idx = (data.campaigns || []).findIndex(c => c.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '找不到活動' });
  const c = data.campaigns[idx];
  if (c.owner !== username && !['admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '無權限' });
  data.campaigns[idx] = { ...c, ...pickFields(req.body, CAMPAIGN_FIELDS), id: c.id, owner: c.owner };
  db.save(data);
  res.json(data.campaigns[idx]);
});

app.delete('/api/campaigns/:id', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  const data = db.load();
  const idx = (data.campaigns || []).findIndex(c => c.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '找不到活動' });
  if (data.campaigns[idx].owner !== username && !['admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '無權限' });
  data.campaigns.splice(idx, 1);
  db.save(data);
  res.json({ success: true });
});

// ── Lead CRUD ─────────────────────────────────────────

app.get('/api/leads', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  const data = db.load();
  let list = data.leads || [];
  if (role === 'marketing') {
    list = list.filter(l => l.owner === username);
  } else if (role === 'user') {
    list = list.filter(l => l.assignedTo === username);
  } else if (!['admin','manager1','manager2'].includes(role)) {
    return res.json([]);
  }
  // manager2 只看自己可視範圍的業務分配
  if (role === 'manager2') {
    const auth = loadAuth();
    const visUsers = auth.users.filter(u => u.role === 'user' || u.username === username).map(u => u.username);
    const allMarketing = auth.users.filter(u => u.role === 'marketing').map(u => u.username);
    list = list.filter(l => allMarketing.includes(l.owner) || visUsers.includes(l.assignedTo) || l.assignedTo === username);
  }
  res.json(list.sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || '')));
});

app.post('/api/leads', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  if (!['marketing','admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '無權限' });
  const data = db.load();
  if (!data.leads) data.leads = [];
  const l = {
    id: uuidv4(), owner: username,
    ...pickFields(req.body, LEAD_FIELDS),
    status: 'new',
    assignedTo: null, assignedBy: null, assignedAt: null,
    opportunityId: null, convertedAt: null,
    createdAt: new Date().toISOString()
  };
  if (!l.company && !l.contactName) return res.status(400).json({ error: '請填入公司或聯絡人' });
  data.leads.push(l);
  db.save(data);
  res.status(201).json(l);
});

app.put('/api/leads/:id', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  const data = db.load();
  const idx = (data.leads || []).findIndex(l => l.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '找不到 Lead' });
  const l = data.leads[idx];
  if (l.owner !== username && !['admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '無權限' });
  data.leads[idx] = { ...l, ...pickFields(req.body, LEAD_FIELDS), id: l.id, owner: l.owner };
  db.save(data);
  res.json(data.leads[idx]);
});

// 指派 Lead 給業務
app.post('/api/leads/:id/assign', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  if (!['admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '僅主管可指派 Lead' });
  const { assignedTo } = req.body;
  if (!assignedTo) return res.status(400).json({ error: '請選擇指派業務' });
  const data = db.load();
  const auth = loadAuth();
  const targetUser = auth.users.find(u => u.username === assignedTo && u.role === 'user');
  if (!targetUser) return res.status(400).json({ error: '找不到此業務帳號' });
  const l = (data.leads || []).find(l => l.id === req.params.id);
  if (!l) return res.status(404).json({ error: '找不到 Lead' });
  if (l.status === 'converted') return res.status(400).json({ error: 'Lead 已轉換，無法重新指派' });
  l.assignedTo  = assignedTo;
  l.assignedBy  = username;
  l.assignedAt  = new Date().toISOString();
  l.status      = 'assigned';
  db.save(data);
  // 通知業務
  pushNotification(assignedTo, 'lead_assigned', '🎯 新 Lead 指派',
    `${l.company || l.contactName} 已指派給您`, l.id);
  res.json(l);
});

// 轉換 Lead → Contact + Opportunity
app.post('/api/leads/:id/convert', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  if (!['admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '僅主管可轉換 Lead' });
  const data = db.load();
  const l = (data.leads || []).find(l => l.id === req.params.id);
  if (!l) return res.status(404).json({ error: '找不到 Lead' });
  if (l.status === 'converted') return res.status(400).json({ error: '此 Lead 已轉換' });

  const salesPerson = req.body.salesPerson || l.assignedTo;
  if (!salesPerson) return res.status(400).json({ error: '請指定負責業務' });
  const { product, category, stage, oppName } = req.body;

  if (!data.contacts) data.contacts = [];
  if (!data.opportunities) data.opportunities = [];

  // 建立聯絡人（不重複）
  const exists = data.contacts.find(c =>
    c.owner === salesPerson &&
    ((l.company && c.company === l.company) || (l.contactName && c.name === l.contactName))
  );
  let contactId = exists ? exists.id : null;
  if (!exists) {
    const newContact = {
      id: uuidv4(), owner: salesPerson,
      name: l.contactName || '', company: l.company || '',
      title: l.title || '', phone: l.phone || '', email: l.email || '',
      note: `[Lead] ${l.campaignName || ''} - ${l.interest || ''}`,
      fromLeadId: l.id, createdAt: new Date().toISOString()
    };
    data.contacts.push(newContact);
    contactId = newContact.id;
  }

  // 建立商機
  const opp = {
    id: uuidv4(), owner: salesPerson,
    contactId: contactId || '',
    contactName: l.contactName || '', company: l.company || '',
    product: oppName || product || l.interest || '',
    category: category || '', stage: stage || 'C',
    amount: '', expectedDate: '', grossMarginRate: '',
    description: `來源活動：${l.campaignName || ''}`,
    fromLeadId: l.id, createdAt: new Date().toISOString()
  };
  data.opportunities.push(opp);

  l.status = 'converted';
  l.opportunityId = opp.id;
  l.convertedAt = new Date().toISOString();
  db.save(data);

  res.json({ lead: l, opportunity: opp, contactId });
});

// 標記不合格
app.post('/api/leads/:id/disqualify', requireAuth, (req, res) => {
  const { role, username } = req.session.user;
  if (!['admin','manager1','manager2'].includes(role))
    return res.status(403).json({ error: '僅主管可標記不合格' });
  const data = db.load();
  const l = (data.leads || []).find(l => l.id === req.params.id);
  if (!l) return res.status(404).json({ error: '找不到 Lead' });
  l.status = 'disqualified';
  l.disqualifyReason = req.body.reason || '';
  l.disqualifiedAt = new Date().toISOString();
  db.save(data);
  res.json(l);
});

// ── 取得指定業務的客戶清單（供移轉功能使用）──────────────────
app.get('/api/contacts-by-owner', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  if (!['admin', 'manager1', 'manager2'].includes(role)) {
    return res.status(403).json({ error: '無查詢權限' });
  }
  const { owner } = req.query;
  if (!owner) return res.status(400).json({ error: '請指定 owner' });

  // 確認查詢對象在可視範圍內
  const auth = loadAuth();
  const targetUser = auth.users.find(u => u.username === owner);
  if (!targetUser) return res.status(404).json({ error: '找不到此使用者' });

  if (role === 'manager2') {
    if (targetUser.role !== 'user' && targetUser.username !== username)
      return res.status(403).json({ error: '超出可視範圍' });
  } else if (role === 'manager1') {
    if (!['user','manager2'].includes(targetUser.role) && targetUser.username !== username)
      return res.status(403).json({ error: '超出可視範圍' });
  }

  const data = db.load();
  const contacts = (data.contacts || []).filter(c => c.owner === owner && !c.deleted);
  res.json(contacts);
});

// ── 責任業務名單移轉 ──────────────────────────────────────────
app.post('/api/transfer-contacts', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  if (!['admin', 'manager1', 'manager2'].includes(role)) {
    return res.status(403).json({ error: '無移轉權限' });
  }

  const { fromOwner, toOwner, contactIds } = req.body;
  if (!fromOwner || !toOwner) return res.status(400).json({ error: '請指定來源與目標業務' });
  if (fromOwner === toOwner) return res.status(400).json({ error: '來源與目標業務不能相同' });

  const auth = loadAuth();
  const fromUser = auth.users.find(u => u.username === fromOwner);
  const toUser   = auth.users.find(u => u.username === toOwner);
  if (!fromUser || !toUser) return res.status(400).json({ error: '指定的使用者不存在' });

  // 權限：manager2 只能移轉 user，manager1 可移轉 user/manager2
  const transferableRoles = role === 'admin'    ? ['user','manager1','manager2','secretary']
                          : role === 'manager1' ? ['user','manager2']
                          : ['user'];
  if (!transferableRoles.includes(fromUser.role))
    return res.status(403).json({ error: '您無權移轉此業務的客戶名單' });

  const data = db.load();

  // 取得要移轉的聯絡人
  let toTransfer;
  if (!contactIds || contactIds === 'all') {
    toTransfer = (data.contacts || []).filter(c => c.owner === fromOwner);
  } else {
    const idSet = new Set(Array.isArray(contactIds) ? contactIds : [contactIds]);
    toTransfer = (data.contacts || []).filter(c => c.owner === fromOwner && idSet.has(c.id));
  }

  if (!toTransfer.length) return res.status(400).json({ error: '沒有可移轉的客戶資料' });

  const transferredIds      = new Set(toTransfer.map(c => c.id));
  const transferredCompanies = new Set(toTransfer.map(c => c.company).filter(Boolean));

  // 移轉各類資料
  let contactCount = 0, visitCount = 0, oppCount = 0, recvCount = 0;

  (data.contacts || []).forEach(c => {
    if (transferredIds.has(c.id)) { c.owner = toOwner; contactCount++; }
  });
  (data.visits || []).forEach(v => {
    if (v.owner === fromOwner && transferredIds.has(v.contactId)) { v.owner = toOwner; visitCount++; }
  });
  (data.opportunities || []).forEach(o => {
    if (o.owner === fromOwner && transferredIds.has(o.contactId)) { o.owner = toOwner; oppCount++; }
  });
  (data.receivables || []).forEach(r => {
    if (r.owner === fromOwner && transferredCompanies.has(r.company)) { r.owner = toOwner; recvCount++; }
  });

  db.save(data);
  writeLog('TRANSFER_CONTACTS', username, `${fromOwner}→${toOwner}`,
    `移轉 ${contactCount} 位客戶（拜訪 ${visitCount} 筆、商機 ${oppCount} 筆、帳款 ${recvCount} 筆）`, req);

  res.json({ success: true, contactCount, visitCount, oppCount, recvCount });
});

// ── 客戶名單匯入：下載範本 ────────────────────────────────────
app.get('/api/admin/import-template', requireAdmin, (req, res) => {
  const headers = ['姓名','英文名稱','公司','職稱','電話','分機','手機','Email','地址','網站','統一編號','產業屬性','備註'];
  const sample  = ['王大明','David Wang','東捷資訊','業務經理','02-12345678','123','0912-345678','david@itts.com.tw','台北市信義區','https://www.itts.com.tw','12345678','資訊服務','重要客戶'];
  const ws = XLSX.utils.aoa_to_sheet([headers, sample]);
  ws['!cols'] = [10,12,18,12,14,6,14,24,24,28,10,12,16].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '客戶名單');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename*=UTF-8\'\'%E5%AE%A2%E6%88%B6%E5%90%8D%E5%96%AE%E5%8C%AF%E5%85%A5%E7%AF%84%E6%9C%AC.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// ── 客戶名單匯入：上傳並匯入 ──────────────────────────────────
const uploadImport = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /xlsx|xls|csv/.test(path.extname(file.originalname).toLowerCase())
             || /spreadsheet|csv/.test(file.mimetype);
    ok ? cb(null, true) : cb(new Error('只支援 .xlsx / .xls / .csv 格式'));
  }
});

app.post('/api/admin/import-contacts', requireAdmin, uploadImport.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '未收到檔案' });
  const { targetOwner, skipDuplicates } = req.body;
  if (!targetOwner) return res.status(400).json({ error: '請指定匯入對象' });

  const auth = loadAuth();
  if (!auth.users.find(u => u.username === targetOwner))
    return res.status(400).json({ error: '指定的使用者不存在' });

  // 解析 Excel
  let rows;
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  } catch (e) {
    return res.status(400).json({ error: '檔案解析失敗，請確認格式正確' });
  }

  if (!rows.length) return res.status(400).json({ error: '檔案內無資料列' });

  const COL_MAP = {
    '姓名': 'name', '英文名稱': 'nameEn', '公司': 'company', '職稱': 'title',
    '電話': 'phone', '分機': 'ext', '手機': 'mobile', 'Email': 'email',
    '地址': 'address', '網站': 'website', '統一編號': 'taxId',
    '產業屬性': 'industry', '備註': 'note'
  };

  const data = db.load();
  if (!data.contacts) data.contacts = [];

  let imported = 0, skipped = 0, errors = 0;
  const errorDetails = [];

  rows.forEach((row, i) => {
    try {
      const contact = { id: uuidv4(), owner: targetOwner, createdAt: new Date().toISOString(),
        name:'', nameEn:'', company:'', title:'', phone:'', ext:'', mobile:'', email:'',
        address:'', website:'', taxId:'', industry:'', note:'',
        opportunityStage:'', isPrimary: false, systemVendor:'', systemProduct:'', cardImage:'' };

      Object.entries(COL_MAP).forEach(([col, field]) => {
        if (row[col] !== undefined && row[col] !== '') contact[field] = String(row[col]).trim();
      });

      if (!contact.name && !contact.company) { skipped++; return; }

      // 驗證網址
      if (contact.website && !/^https?:\/\//i.test(contact.website)) contact.website = '';

      // 重複檢查（同 owner、同公司、同姓名，已刪除不算重複）
      if (skipDuplicates === 'true') {
        const dup = data.contacts.find(c => !c.deleted && c.owner === targetOwner
          && c.name === contact.name && c.company === contact.company);
        if (dup) { skipped++; return; }
      }

      data.contacts.push(contact);
      imported++;
    } catch (e) {
      errors++;
      errorDetails.push(`第 ${i + 2} 列：${e.message}`);
    }
  });

  db.save(data);
  writeLog('IMPORT_CONTACTS', req.session.user.username, targetOwner,
    `批次匯入客戶名單：${imported} 筆成功，${skipped} 筆略過，${errors} 筆錯誤`, req);

  res.json({ success: true, imported, skipped, errors, errorDetails });
});

// ── 名片辨識 JSON 批次匯入 ────────────────────────────────────
app.post('/api/admin/import-contacts-json', requireAdmin, (req, res) => {
  const { targetOwner, contacts: rows, skipDuplicates } = req.body;
  if (!targetOwner) return res.status(400).json({ error: '請指定匯入對象' });
  if (!Array.isArray(rows) || !rows.length) return res.status(400).json({ error: '無資料' });

  const auth = loadAuth();
  if (!auth.users.find(u => u.username === targetOwner))
    return res.status(400).json({ error: '指定的使用者不存在' });

  const data = db.load();
  if (!data.contacts) data.contacts = [];

  let imported = 0, skipped = 0;
  rows.forEach(row => {
    if (!row.name && !row.company) { skipped++; return; }
    if (skipDuplicates) {
      const dup = data.contacts.find(c => !c.deleted && c.owner === targetOwner
        && c.name === (row.name||'') && c.company === (row.company||''));
      if (dup) { skipped++; return; }
    }
    data.contacts.push({
      id: uuidv4(), owner: targetOwner, createdAt: new Date().toISOString(),
      name: row.name||'', nameEn: row.nameEn||'', company: row.company||'',
      title: row.title||'', phone: row.phone||'', ext: row.ext||'',
      mobile: row.mobile||'', email: row.email||'', address: row.address||'',
      website: row.website||'', taxId: row.taxId||'', industry: row.industry||'',
      note: row.note||'', isPrimary: false, systemVendor:'', systemProduct:'',
      cardImage:'', customerType:'', productLine:'', isResigned: false
    });
    imported++;
  });

  db.save(data);
  writeLog('IMPORT_JSON', req.session.user.username, targetOwner,
    `名片辨識匯入：${imported} 筆成功，${skipped} 筆略過`, req);
  res.json({ success: true, imported, skipped });
});

// ── 資料遷移：為舊資料加上 owner ────────────────────────────
// 一次性遷移：將資料庫中所有 stage==='成交' 改為 'Won'
function migrateStage成交ToWon() {
  const data = db.load();
  let changed = 0;
  ['opportunities', 'lostOpportunities'].forEach(key => {
    if (!Array.isArray(data[key])) return;
    data[key].forEach(o => {
      if (o.stage === '成交') { o.stage = 'Won'; changed++; }
      // 也修正 stageHistory 中的 from/to 欄位
      if (Array.isArray(o.stageHistory)) {
        o.stageHistory.forEach(h => {
          if (h.from === '成交') { h.from = 'Won'; changed++; }
          if (h.to   === '成交') { h.to   = 'Won'; changed++; }
        });
      }
    });
  });
  if (changed > 0) {
    db.save(data);
    console.log(`[migrate] 成交→Won: 更新了 ${changed} 筆資料`);
  }
}

function migrateOwner() {
  const data = db.load();
  let changed = false;
  const DEFAULT_OWNER = 'Stevenlee';
  ['contacts','visits','opportunities','contracts','targets'].forEach(key => {
    if (Array.isArray(data[key])) {
      data[key].forEach(item => {
        if (!item.owner) { item.owner = DEFAULT_OWNER; changed = true; }
      });
    }
  });
  if (changed) {
    db.save(data);
    console.log(`✅ 資料遷移完成：舊資料已歸屬至 ${DEFAULT_OWNER}`);
  }
}

// ── 職能分類 Auto-Mapping（伺服器端）────────────────────
const JOB_FUNCTION_CATEGORIES_SVR = [
  { key: 'management', keywords: ['董事長','副董事長','總經理','副總','協理','總監','CEO','COO','CFO','CTO','CMO','VP','Vice President','President','Director','執行長','執行副總'] },
  { key: 'operations', keywords: ['廠長','生管','物管','資材','MC','課長','組長','班長','Supervisor','作業員','技術員','OP','Technician','生產','製造','倉管','物料','採購主管','廠務','現場'] },
  { key: 'engineering', keywords: ['研發工程師','R&D','研發','製程工程師','製程','PE','設備工程師','設備','EE','工業工程師','工業工程','IE','產品經理','PM','Product Manager','軟體工程師','系統工程師','MIS','IT工程','架構師','開發工程師','韌體','硬體','機械工程師'] },
  { key: 'quality', keywords: ['品保','品管','IQC','IPQC','OQC','測試工程師','QA','QC','TE','品質','品控','稽核','驗證','認證','可靠度'] },
  { key: 'admin', keywords: ['業務','Sales','採購','Buyer','財務','會計','人力資源','HR','環安衛','ESH','行政','秘書','助理','公關','行銷','Marketing','法務','企劃','客服','業務專員','業務經理','業務副理'] }
];

function autoMapJobFunctionSvr(title) {
  if (!title) return '';
  const t = title.toLowerCase();
  for (const cat of JOB_FUNCTION_CATEGORIES_SVR) {
    if (cat.keywords.some(kw => t.includes(kw.toLowerCase()))) return cat.key;
  }
  return '';
}

// ── 管理員：取得聯絡人稽核日誌 ──────────────────────────────
app.get('/api/admin/contact-audit', requireAdmin, (req, res) => {
  let logs = [];
  try { logs = JSON.parse(fs.readFileSync(CONTACT_AUDIT_FILE, 'utf8')); } catch {}
  const { userId, action, dateFrom, dateTo } = req.query;
  let result = logs;
  if (userId) result = result.filter(l => l.userId === userId);
  if (action) result = result.filter(l => l.action === action);
  if (dateFrom) result = result.filter(l => l.timestamp >= dateFrom);
  if (dateTo)   result = result.filter(l => l.timestamp <= dateTo + 'T23:59:59Z');
  res.json(result.slice(0, 500)); // 最多回傳 500 筆
});

// ── 管理員：取得已軟刪除的聯絡人 ───────────────────────────
app.get('/api/admin/deleted-contacts', requireAdmin, (req, res) => {
  const data = db.load();
  const deleted = (data.contacts || [])
    .filter(c => c.deleted)
    .sort((a, b) => (b.deletedAt || '').localeCompare(a.deletedAt || ''));
  res.json(deleted);
});

// ── 管理員：還原已刪除的聯絡人 ─────────────────────────────
app.post('/api/admin/contacts/:id/restore', requireAdmin, (req, res) => {
  const data = db.load();
  const idx = data.contacts.findIndex(c => c.id === req.params.id && c.deleted);
  if (idx === -1) return res.status(404).json({ error: '找不到此已刪除聯絡人' });
  const contact = data.contacts[idx];
  delete contact.deleted;
  delete contact.deletedAt;
  delete contact.deletedBy;
  delete contact.deletedByName;
  db.save(data);
  writeContactAudit('RESTORE', req, contact, []);
  res.json({ success: true });
});

// ── 管理員：永久刪除聯絡人（含圖片）───────────────────────
app.delete('/api/admin/contacts/:id/permanent', requireAdmin, (req, res) => {
  const data = db.load();
  const contact = data.contacts.find(c => c.id === req.params.id && c.deleted);
  if (!contact) return res.status(404).json({ error: '找不到此已刪除聯絡人' });
  // 這時才真正刪除圖片
  if (contact.cardImage) {
    const basename = path.basename(contact.cardImage);
    if (/^[0-9a-f-]{36}\.(jpg|jpeg|png|gif|webp)$/i.test(basename)) {
      const imgPath = path.join(__dirname, 'uploads', basename);
      if (fs.existsSync(imgPath)) try { fs.unlinkSync(imgPath); } catch {}
    }
  }
  writeContactAudit('PERMANENT_DELETE', req, contact, []);
  data.contacts = data.contacts.filter(c => c.id !== req.params.id);
  db.save(data);
  res.json({ success: true });
});

// ── 管理員：批次自動 Mapping 職能分類 ──────────────────────
app.post('/api/admin/migrate-job-function', requireAdmin, (req, res) => {
  const data = db.load();
  let mapped = 0, skipped = 0;
  (data.contacts || []).forEach(c => {
    if (c.jobFunction) { skipped++; return; } // 已有值不覆蓋
    const key = autoMapJobFunctionSvr(c.title || '');
    if (key) { c.jobFunction = key; mapped++; }
  });
  db.save(data);
  writeLog('MIGRATE_JOB_FUNCTION', 'admin', 'all', `職能分類批次 Mapping：${mapped} 筆已更新，${skipped} 筆已略過`, req);
  res.json({ success: true, mapped, skipped });
});

// ── 版本資訊 ────────────────────────────────────────────
const os = require('os');
const SERVER_START_TIME = new Date();

app.get('/api/admin/version', requireAdmin, (req, res) => {
  // 讀取 package.json
  let pkg = {};
  try { pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8')); } catch {}

  // 取各套件實際安裝版本
  const getVer = (name) => {
    try {
      const p = JSON.parse(fs.readFileSync(path.join(__dirname, 'node_modules', name, 'package.json'), 'utf8'));
      return p.version;
    } catch { return '—'; }
  };

  const uptimeSec = Math.floor((Date.now() - SERVER_START_TIME.getTime()) / 1000);
  const d = Math.floor(uptimeSec / 86400);
  const h = Math.floor((uptimeSec % 86400) / 3600);
  const m = Math.floor((uptimeSec % 3600) / 60);
  const s = uptimeSec % 60;
  const uptimeStr = `${d > 0 ? d + ' 天 ' : ''}${h > 0 ? h + ' 時 ' : ''}${m} 分 ${s} 秒`;

  res.json({
    app: {
      name:        pkg.name        || 'business-card-crm',
      version:     pkg.version     || '1.0.0',
      description: pkg.description || '名片管理 CRM 系統',
      startTime:   SERVER_START_TIME.toISOString(),
      uptime:      uptimeStr,
    },
    runtime: {
      node:      process.version,
      platform:  os.platform() === 'win32' ? 'Windows' : os.platform() === 'darwin' ? 'macOS' : 'Linux',
      arch:      os.arch(),
      hostname:  os.hostname(),
      cpus:      os.cpus().length,
      cpuModel:  os.cpus()[0]?.model || '—',
      totalMem:  os.totalmem(),
      freeMem:   os.freemem(),
      osUptime:  Math.floor(os.uptime()),
    },
    backend: [
      { name: 'Express',            ver: getVer('express'),            desc: 'Web 框架',              icon: '🚀' },
      { name: 'express-session',    ver: getVer('express-session'),    desc: 'Session 管理',          icon: '🔐' },
      { name: 'bcryptjs',           ver: getVer('bcryptjs'),           desc: '密碼雜湊加密',          icon: '🔒' },
      { name: 'helmet',             ver: getVer('helmet'),             desc: 'HTTP 安全標頭',         icon: '🛡️' },
      { name: 'express-rate-limit', ver: getVer('express-rate-limit'), desc: '登入速率限制',          icon: '⏱️' },
      { name: 'multer',             ver: getVer('multer'),             desc: '檔案上傳處理',          icon: '📤' },
      { name: 'xlsx',               ver: getVer('xlsx'),               desc: 'Excel 匯出',            icon: '📊' },
      { name: 'uuid',               ver: getVer('uuid'),               desc: '唯一識別碼產生',        icon: '🔑' },
      { name: 'dotenv',             ver: getVer('dotenv'),             desc: '環境變數管理',          icon: '⚙️' },
      { name: 'cors',               ver: getVer('cors'),               desc: 'CORS 跨域控制',         icon: '🌐' },
    ],
    frontend: [
      { name: 'Vanilla JS (ES2022)',  ver: '—',        desc: '前端互動邏輯',     icon: '🟨' },
      { name: 'Chart.js',             ver: '4.4.0',    desc: '圖表視覺化',       icon: '📈' },
      { name: 'HTML5 / CSS3',         ver: '—',        desc: '頁面結構與樣式',   icon: '🎨' },
    ],
    database: {
      type:     'JSON File Database',
      engine:   '自製輕量 JSON 資料庫（db.js）',
      file:     'data.json',
      authFile: 'auth.json',
      note:     '無需額外資料庫伺服器，資料以 JSON 格式儲存於本機磁碟',
    },
    generatedAt: new Date().toISOString(),
  });
});

// ── 容量空間監控 ──────────────────────────────────────────
app.get('/api/admin/storage', requireAdmin, (req, res) => {
  try {
    const data = db.load();
    const auth = loadAuth();
    const dbPath = path.join(__dirname, 'data.json');
    const uploadsPath = path.join(__dirname, 'uploads');

    // data.json 大小
    const dbStat = fs.existsSync(dbPath) ? fs.statSync(dbPath) : null;
    const dbSize = dbStat ? dbStat.size : 0;

    // uploads 資料夾
    let uploadsSize = 0, uploadsCount = 0;
    const uploadsFiles = [];
    if (fs.existsSync(uploadsPath)) {
      const walkDir = (dir) => {
        const files = fs.readdirSync(dir);
        files.forEach(f => {
          const fp = path.join(dir, f);
          const st = fs.statSync(fp);
          if (st.isDirectory()) { walkDir(fp); }
          else {
            uploadsSize += st.size;
            uploadsCount++;
            uploadsFiles.push({ name: f, size: st.size, mtime: st.mtime });
          }
        });
      };
      walkDir(uploadsPath);
    }
    // 最近 10 個檔案（依修改時間排序）
    uploadsFiles.sort((a, b) => new Date(b.mtime) - new Date(a.mtime));

    // 各集合筆數
    const collections = [
      { key: 'contacts',           label: '聯絡人名片',   icon: '👤' },
      { key: 'opportunities',      label: '商機記錄',     icon: '💡' },
      { key: 'lostOpportunities',  label: '流失商機',     icon: '💔' },
      { key: 'visits',             label: '業務日報',     icon: '📋' },
      { key: 'contracts',          label: '合約管理',     icon: '📄' },
      { key: 'receivables',        label: '應收帳款',     icon: '💰' },
      { key: 'callin',             label: 'Call-in Pass', icon: '📞' },
      { key: 'targets',            label: '業績目標',     icon: '🎯' },
    ].map(c => ({
      ...c,
      count: Array.isArray(data[c.key]) ? data[c.key].length : 0,
    }));

    const totalRecords = collections.reduce((s, c) => s + c.count, 0);

    res.json({
      dbSize,
      uploadsSize,
      uploadsCount,
      recentFiles: uploadsFiles.slice(0, 10),
      collections,
      totalRecords,
      users: (auth.users || []).length,
      generatedAt: new Date().toISOString()
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── 雲端基礎設施監控 ──────────────────────────────────────
app.get('/api/admin/infra-stats', requireAdmin, async (req, res) => {
  const result = {
    db:      null,
    storage: null,
    runtime: null,
    vercel:  null,
  };

  // ── 1. Supabase PostgreSQL 統計（需要 postgres 後端）──
  if (process.env.DB_BACKEND === 'postgres' && process.env.DATABASE_URL) {
    try {
      const { Pool } = require('pg');
      const pool = new Pool({
        connectionString: process.env.DATABASE_URL,
        ssl: { rejectUnauthorized: false },
        max: 1, connectionTimeoutMillis: 5000,
      });
      const [sizeRow, connRow, tablesRow, storageRow] = await Promise.all([
        pool.query(`SELECT pg_database_size(current_database()) AS db_bytes,
                          pg_size_pretty(pg_database_size(current_database())) AS db_pretty`),
        pool.query(`SELECT count(*) FILTER (WHERE state = 'active') AS active,
                          count(*) AS total
                   FROM pg_stat_activity`),
        pool.query(`SELECT relname AS name,
                          n_live_tup AS rows,
                          pg_size_pretty(pg_total_relation_size(relid)) AS size,
                          pg_total_relation_size(relid) AS size_bytes
                   FROM pg_stat_user_tables
                   ORDER BY pg_total_relation_size(relid) DESC
                   LIMIT 10`),
        pool.query(`SELECT COUNT(*) AS files,
                          COALESCE(SUM((metadata->>'size')::bigint), 0) AS total_bytes
                   FROM storage.objects
                   WHERE bucket_id = $1`, [process.env.SUPABASE_BUCKET || 'uploads']),
      ]);
      await pool.end();

      result.db = {
        sizeBytes:   parseInt(sizeRow.rows[0].db_bytes),
        sizePretty:  sizeRow.rows[0].db_pretty,
        connections: { active: parseInt(connRow.rows[0].active), total: parseInt(connRow.rows[0].total) },
        tables:      tablesRow.rows.map(r => ({ name: r.name, rows: parseInt(r.rows), size: r.size, sizeBytes: parseInt(r.size_bytes) })),
      };
      result.storage = {
        files:      parseInt(storageRow.rows[0].files),
        totalBytes: parseInt(storageRow.rows[0].total_bytes),
        bucket:     process.env.SUPABASE_BUCKET || 'uploads',
      };
    } catch (e) {
      result.db = { error: e.message };
    }
  }

  // ── 2. 執行時資訊（本地 + Vercel 都能取得）──
  const mem = process.memoryUsage();
  result.runtime = {
    uptimeSeconds: Math.floor(process.uptime()),
    nodeVersion:   process.version,
    env:           process.env.NODE_ENV || 'development',
    heapUsed:      mem.heapUsed,
    heapTotal:     mem.heapTotal,
    rss:           mem.rss,
    backend:       process.env.DB_BACKEND || 'json',
    storageBackend: process.env.STORAGE_BACKEND || 'local',
  };

  // ── 3. Vercel 環境（只在 Vercel 上有值）──
  result.vercel = {
    url:        process.env.VERCEL_URL || null,
    productionUrl: process.env.VERCEL_PROJECT_PRODUCTION_URL || null,
    region:     process.env.VERCEL_REGION || (process.env.VERCEL_URL ? 'hnd1' : null),
    isVercel:   !!process.env.VERCEL_URL,
  };

  result.generatedAt = new Date().toISOString();
  res.json(result);
});

// ── API 使用量監控 ────────────────────────────────────────
app.get('/api/admin/api-stats', requireAdmin, (req, res) => {
  try {
    const summary = apiMonitor.getSummary();
    summary.generatedAt = new Date().toISOString();
    res.json(summary);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── 報價單功能 ─────────────────────────────────────────────
const QUOTE_TEMPLATE = path.join(__dirname, 'templates', 'quotation_template.xlsx');

function genQuoteNo(data) {
  const d = new Date();
  const yyyymm = d.getFullYear().toString() + String(d.getMonth() + 1).padStart(2, '0');
  const prefix = `QU-${yyyymm}-`;
  const count = (data.quotations || []).filter(q => q.quoteNo && q.quoteNo.startsWith(prefix)).length;
  return prefix + String(count + 1).padStart(3, '0');
}

/**
 * 寫入儲存格值（字串/數字/公式）
 * addr: Excel 位址字串，如 'B9'
 */
function _wc(ws, addr, val) {
  if (!ws[addr]) ws[addr] = {};
  if (typeof val === 'string' && val.startsWith('=')) {
    ws[addr].t = 'n';
    ws[addr].f = val.slice(1);
    delete ws[addr].v;
  } else if (typeof val === 'number') {
    ws[addr].t = 'n';
    ws[addr].v = val;
    delete ws[addr].f;
  } else {
    ws[addr].t = 's';
    ws[addr].v = val == null ? '' : String(val);
    delete ws[addr].f;
  }
}

/**
 * 將 fromRow1（含，1-based）以下所有列往下移 count 列
 * 並同步更新 !merges 與 !ref
 */
function _shiftRowsDown(ws, fromRow1, count) {
  if (count <= 0) return;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const fromR0 = fromRow1 - 1; // 0-based

  // 從底部往上移動，避免覆蓋
  for (let r = range.e.r; r >= fromR0; r--) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const srcAddr = XLSX.utils.encode_cell({ r, c });
      const dstAddr = XLSX.utils.encode_cell({ r: r + count, c });
      if (ws[srcAddr]) {
        ws[dstAddr] = { ...ws[srcAddr] };
        delete ws[srcAddr];
      } else {
        delete ws[dstAddr];
      }
    }
  }

  // 更新 merges
  if (ws['!merges']) {
    ws['!merges'] = ws['!merges'].map(m => {
      if (m.s.r >= fromR0) {
        return { s: { r: m.s.r + count, c: m.s.c }, e: { r: m.e.r + count, c: m.e.c } };
      }
      return m;
    });
  }

  // 更新 !ref 範圍
  range.e.r += count;
  ws['!ref'] = XLSX.utils.encode_range(range);
}

/**
 * 以報價單範本產生 Excel Buffer
 */
function buildQuoteExcel(q) {
  const wb = XLSX.readFile(QUOTE_TEMPLATE);
  const sheetName = wb.SheetNames[0]; // "報價單 "（含尾端空格）
  const ws = wb.Sheets[sheetName];

  const items = Array.isArray(q.items) && q.items.length > 0
    ? q.items
    : [{ desc: '', unit: '式', qty: 1, unitPrice: 0 }];
  const n = items.length;

  // 範本有 2 個項目列（row 17, 18），超出時往下插
  const ITEM_START  = 17; // 1-based
  const ORIG_ROWS   = 2;
  const SUMMARY_ROW = ITEM_START + ORIG_ROWS; // 19（1-based）：第一個小計列
  const extraRows   = Math.max(0, n - ORIG_ROWS);
  const ss          = SUMMARY_ROW + extraRows; // 移位後小計列（1-based）

  if (extraRows > 0) {
    _shiftRowsDown(ws, SUMMARY_ROW, extraRows);
  }

  // ── 表頭資訊 ──
  const dateStr = (q.quoteDate || new Date().toISOString().slice(0, 10)).replace(/-/g, '/');
  _wc(ws, 'G6',  `表單編號：${q.quoteNo || ''}`);
  _wc(ws, 'B9',  q.company     || '');
  _wc(ws, 'F9',  dateStr);
  _wc(ws, 'B10', q.contactName || '');
  _wc(ws, 'F10', q.contactName || '');
  _wc(ws, 'B11', q.address     || '');
  _wc(ws, 'F11', q.mobile      || '');
  _wc(ws, 'B12', q.phone       || '');

  // ── 項目列 ──
  const lastItemRow = ITEM_START + n - 1; // 1-based
  for (let i = 0; i < n; i++) {
    const row1 = ITEM_START + i;
    const rs   = String(row1);
    const item = items[i];
    _wc(ws, `B${rs}`, i + 1);
    _wc(ws, `C${rs}`, item.desc      || '');
    _wc(ws, `F${rs}`, item.unit      || '式');
    _wc(ws, `G${rs}`, parseFloat(item.qty)       || 1);
    _wc(ws, `H${rs}`, parseFloat(item.unitPrice) || 0);
    _wc(ws, `J${rs}`, `=H${rs}*G${rs}`);

    // 第 2 列以後需補 C:E merge（第 1 列範本已有）
    if (i >= ORIG_ROWS) {
      if (!ws['!merges']) ws['!merges'] = [];
      ws['!merges'].push({ s: { r: row1 - 1, c: 2 }, e: { r: row1 - 1, c: 4 } });
    }
  }

  // ── 小計/優惠/稅/合計 公式（參照移位後的正確列號）──
  const ssStr        = String(ss);
  const discType     = q.discountType  || 'none';
  const discValue    = parseFloat(q.discountValue) || 0;

  _wc(ws, `J${ss}`, `=SUM(J${ITEM_START}:J${lastItemRow})`); // 小計

  // 優惠價：依折扣類型決定公式或數值
  if (discType === 'percent' && discValue > 0 && discValue < 100) {
    // 例：90 → 九折 → J(ss)*90/100
    _wc(ws, `J${ss + 1}`, `=J${ssStr}*${discValue}/100`);
    // 在 G(ss+1) 補上折扣說明（同列已有 "優惠價：" 標籤的欄）
    _wc(ws, `G${ss + 1}`, `優惠 ${discValue}%（${(discValue / 10).toFixed(1).replace(/\.0$/, '')} 折）`);
  } else if (discType === 'amount' && discValue > 0) {
    // 業務直接輸入議價金額
    _wc(ws, `J${ss + 1}`, discValue);
    _wc(ws, `G${ss + 1}`, '議價金額');
  } else {
    // 無折扣：優惠價 = 小計
    _wc(ws, `J${ss + 1}`, `=J${ssStr}`);
  }

  _wc(ws, `J${ss + 3}`, `=J${ss + 1}*0.05`);          // 稅 5%
  _wc(ws, `J${ss + 4}`, `=J${ss + 1}+J${ss + 3}`);    // 含稅合計

  // ── 專案名稱 / 專案號碼（原本在 B28, B31，隨 extraRows 移位）──
  const projNameRow = 28 + extraRows;
  const projNoRow   = 31 + extraRows;
  _wc(ws, `B${projNameRow}`, q.projectName || '');
  _wc(ws, `B${projNoRow}`,   q.projectNo   || '');

  // ════════════════════════════════════════════
  // ── PNL 毛利分析工作表 ───────────────────────
  // ════════════════════════════════════════════
  const pnlWs = {};
  pnlWs['!ref'] = 'A1:J' + (items.length + 10);

  const PNL_BLUE   = { rgb: '1A4E8C' };
  const PNL_WHITE  = { rgb: 'FFFFFF' };
  const PNL_LBLUE  = { rgb: 'D6E4F7' };
  const PNL_GREEN  = { rgb: '1B5E20' };
  const PNL_RED    = { rgb: 'B71C1C' };
  const PNL_YELLOW = { rgb: 'FFF8E1' };

  const hdrFont  = { bold: true, color: PNL_WHITE, sz: 11 };
  const hdrFill  = { fgColor: PNL_BLUE,  patternType: 'solid' };
  const subFill  = { fgColor: PNL_LBLUE, patternType: 'solid' };
  const bodyFont = { sz: 10 };
  const ctrAlign = { horizontal: 'center', vertical: 'center' };
  const rgtAlign = { horizontal: 'right',  vertical: 'center' };

  function _pc(ws2, addr, v, style) {
    ws2[addr] = { v, t: typeof v === 'number' ? 'n' : 's', s: style || {} };
  }

  // 標題列
  _pc(pnlWs, 'A1', q.quoteNo + '  毛利分析（PNL）', { font: { bold: true, sz: 13, color: PNL_BLUE } });
  _pc(pnlWs, 'A2', '報價日期：' + (q.quoteDate || ''), { font: { sz: 10, color: { rgb: '555555' } } });
  _pc(pnlWs, 'E2', '客戶：' + (q.company || ''),       { font: { sz: 10, color: { rgb: '555555' } } });

  // 欄標題（row 4）
  const pnlCols = ['#', '品項說明', '單位', '數量', '報價單價', '報價小計', '成本單價', '成本小計', '毛利', '毛利率'];
  const colLetters = ['A','B','C','D','E','F','G','H','I','J'];
  pnlCols.forEach((label, ci) => {
    _pc(pnlWs, colLetters[ci] + '4', label, { font: hdrFont, fill: hdrFill, alignment: ctrAlign,
      border: { bottom: { style: 'thin', color: { rgb: '888888' } } } });
  });

  // 資料列
  let totalRevenue = 0;
  let totalCost    = 0;

  items.forEach((it, i) => {
    const row    = String(i + 5);
    const qty    = parseFloat(it.qty)       || 0;
    const price  = parseFloat(it.unitPrice) || 0;
    const cost   = parseFloat(it.cost)      || 0;
    const revSub = qty * price;
    const cstSub = qty * cost;
    const gp     = revSub - cstSub;
    const gpPct  = revSub > 0 ? gp / revSub : 0;

    totalRevenue += revSub;
    totalCost    += cstSub;

    const rowFill = i % 2 === 0 ? {} : { fgColor: { rgb: 'F5F5F5' }, patternType: 'solid' };
    const baseStyle = { font: bodyFont, fill: rowFill };

    _pc(pnlWs, 'A' + row, i + 1,            { ...baseStyle, alignment: ctrAlign });
    _pc(pnlWs, 'B' + row, it.desc || '',     { ...baseStyle });
    _pc(pnlWs, 'C' + row, it.unit || '式',   { ...baseStyle, alignment: ctrAlign });
    _pc(pnlWs, 'D' + row, qty,               { ...baseStyle, alignment: rgtAlign });
    _pc(pnlWs, 'E' + row, price,             { ...baseStyle, alignment: rgtAlign, numFmt: '#,##0' });
    _pc(pnlWs, 'F' + row, revSub,            { ...baseStyle, alignment: rgtAlign, numFmt: '#,##0' });
    _pc(pnlWs, 'G' + row, cost,              { ...baseStyle, alignment: rgtAlign, numFmt: '#,##0',
      fill: { fgColor: PNL_YELLOW, patternType: 'solid' }, font: { sz: 10, italic: true } });
    _pc(pnlWs, 'H' + row, cstSub,           { ...baseStyle, alignment: rgtAlign, numFmt: '#,##0' });
    _pc(pnlWs, 'I' + row, gp, { ...baseStyle, alignment: rgtAlign, numFmt: '#,##0',
      font: { sz: 10, bold: true, color: gp >= 0 ? PNL_GREEN : PNL_RED } });
    _pc(pnlWs, 'J' + row, gpPct, { ...baseStyle, alignment: rgtAlign, numFmt: '0.0%',
      font: { sz: 10, bold: true, color: gpPct >= 0.3 ? PNL_GREEN : gpPct >= 0.15 ? { rgb: 'E65100' } : PNL_RED } });
  });

  // 折扣調整（如有）
  const discTypeP  = q.discountType  || 'none';
  const discValueP = parseFloat(q.discountValue) || 0;
  let adjustedRevenue = totalRevenue;
  if (discTypeP === 'percent' && discValueP > 0 && discValueP < 100) {
    adjustedRevenue = totalRevenue * discValueP / 100;
  } else if (discTypeP === 'amount' && discValueP > 0) {
    adjustedRevenue = discValueP;
  }

  const sumRow  = String(items.length + 6);
  const gpTotal = adjustedRevenue - totalCost;
  const gpPctTotal = adjustedRevenue > 0 ? gpTotal / adjustedRevenue : 0;
  const sumFill  = { fgColor: PNL_BLUE, patternType: 'solid' };
  const sumFont  = { bold: true, color: PNL_WHITE, sz: 11 };

  _pc(pnlWs, 'A' + sumRow, '合計', { font: sumFont, fill: sumFill, alignment: ctrAlign });
  _pc(pnlWs, 'B' + sumRow, discTypeP !== 'none' ? '（含折扣調整）' : '',
    { font: { ...sumFont, italic: true, sz: 10 }, fill: sumFill });
  _pc(pnlWs, 'F' + sumRow, adjustedRevenue, { font: sumFont, fill: sumFill, alignment: rgtAlign, numFmt: '#,##0' });
  _pc(pnlWs, 'H' + sumRow, totalCost,       { font: sumFont, fill: sumFill, alignment: rgtAlign, numFmt: '#,##0' });
  _pc(pnlWs, 'I' + sumRow, gpTotal, { font: { ...sumFont, color: gpTotal >= 0 ? { rgb: '69F0AE' } : { rgb: 'FF5252' } },
    fill: sumFill, alignment: rgtAlign, numFmt: '#,##0' });
  _pc(pnlWs, 'J' + sumRow, gpPctTotal, { font: { ...sumFont, color: gpPctTotal >= 0.3 ? { rgb: '69F0AE' } : gpPctTotal >= 0.15 ? { rgb: 'FFCC02' } : { rgb: 'FF5252' } },
    fill: sumFill, alignment: rgtAlign, numFmt: '0.0%' });

  // 欄寬
  pnlWs['!cols'] = [
    { wch: 5 }, { wch: 30 }, { wch: 8 }, { wch: 8 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }
  ];

  XLSX.utils.book_append_sheet(wb, pnlWs, '毛利分析PNL');

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ── 報價單 CRUD ───────────────────────────────────────────
app.get('/api/quotations', requireAuth, (req, res) => {
  const data  = db.load();
  const owners = getViewableOwners(req, 'quotations');
  const list   = (data.quotations || [])
    .filter(q => owners.includes(q.owner))
    .sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
  res.json(list);
});

app.get('/api/quotations/:id', requireAuth, (req, res) => {
  const data = db.load();
  const q    = (data.quotations || []).find(q => q.id === req.params.id);
  if (!q) return res.status(404).json({ error: '找不到此報價單' });
  const owners = getViewableOwners(req, 'quotations');
  if (!owners.includes(q.owner)) return res.status(403).json({ error: '無權限' });
  res.json(q);
});

app.post('/api/quotations', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data  = db.load();
  if (!data.quotations) data.quotations = [];

  const items = Array.isArray(req.body.items)
    ? req.body.items.slice(0, 50).map(it => ({
        desc:      sanitizeStr(it.desc,      200),
        unit:      sanitizeStr(it.unit,       20),
        qty:       Math.max(0.001, sanitizePositiveFloat(it.qty)       || 1),
        unitPrice: sanitizePositiveFloat(it.unitPrice) || 0,
        cost:      sanitizePositiveFloat(it.cost)      || 0,
      }))
    : [];

  const q = {
    id:          uuidv4(),
    owner,
    quoteNo:     genQuoteNo(data),
    contactId:   sanitizeStr(req.body.contactId,    36),
    company:     sanitizeStr(req.body.company,     100),
    contactName: sanitizeStr(req.body.contactName, 100),
    phone:       sanitizeStr(req.body.phone,        50),
    mobile:      sanitizeStr(req.body.mobile,       50),
    address:     sanitizeStr(req.body.address,     200),
    quoteDate:   sanitizeStr(req.body.quoteDate,    10),
    projectName: sanitizeStr(req.body.projectName, 200),
    projectNo:   sanitizeStr(req.body.projectNo,    50),
    items,
    note:          sanitizeStr(req.body.note,        500),
    discountType:  ['none','percent','amount'].includes(req.body.discountType) ? req.body.discountType : 'none',
    discountValue: sanitizePositiveFloat(req.body.discountValue) || 0,
    status:        'draft',
    createdAt:     new Date().toISOString(),
    updatedAt:     new Date().toISOString(),
  };
  data.quotations.push(q);
  db.save(data);
  res.status(201).json(q);
});

app.put('/api/quotations/:id', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  const data = db.load();
  const idx  = (data.quotations || []).findIndex(q => q.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: '找不到此報價單' });
  const q = data.quotations[idx];
  if (role === 'user' && q.owner !== username) return res.status(403).json({ error: '無權限' });

  const items = Array.isArray(req.body.items)
    ? req.body.items.slice(0, 50).map(it => ({
        desc:      sanitizeStr(it.desc,      200),
        unit:      sanitizeStr(it.unit,       20),
        qty:       Math.max(0.001, sanitizePositiveFloat(it.qty)       || 1),
        unitPrice: sanitizePositiveFloat(it.unitPrice) || 0,
        cost:      sanitizePositiveFloat(it.cost)      || 0,
      }))
    : q.items;

  data.quotations[idx] = {
    ...q,
    company:     sanitizeStr(req.body.company,     100) || q.company,
    contactName: sanitizeStr(req.body.contactName, 100),
    phone:       sanitizeStr(req.body.phone,        50),
    mobile:      sanitizeStr(req.body.mobile,       50),
    address:     sanitizeStr(req.body.address,     200),
    quoteDate:   sanitizeStr(req.body.quoteDate,    10) || q.quoteDate,
    projectName: sanitizeStr(req.body.projectName, 200),
    projectNo:   sanitizeStr(req.body.projectNo,    50),
    items,
    note:          sanitizeStr(req.body.note,        500),
    discountType:  ['none','percent','amount'].includes(req.body.discountType) ? req.body.discountType : (q.discountType || 'none'),
    discountValue: req.body.discountValue !== undefined ? (sanitizePositiveFloat(req.body.discountValue) || 0) : (q.discountValue || 0),
    status:        sanitizeStr(req.body.status,       20) || q.status,
    updatedAt:     new Date().toISOString(),
  };
  db.save(data);
  res.json(data.quotations[idx]);
});

app.delete('/api/quotations/:id', requireAuth, (req, res) => {
  const { username, role } = req.session.user;
  const data = db.load();
  const q    = (data.quotations || []).find(q => q.id === req.params.id);
  if (!q) return res.status(404).json({ error: '找不到此報價單' });
  if (role === 'user' && q.owner !== username) return res.status(403).json({ error: '無權限' });
  data.quotations = data.quotations.filter(q => q.id !== req.params.id);
  db.save(data);
  res.json({ success: true });
});

app.get('/api/quotations/:id/export', requireAuth, (req, res) => {
  const data   = db.load();
  const q      = (data.quotations || []).find(q => q.id === req.params.id);
  if (!q) return res.status(404).json({ error: '找不到此報價單' });
  const owners = getViewableOwners(req, 'quotations');
  if (!owners.includes(q.owner)) return res.status(403).json({ error: '無權限' });

  if (!fs.existsSync(QUOTE_TEMPLATE)) {
    return res.status(500).json({ error: '報價單範本不存在，請聯繫管理員' });
  }
  try {
    const buf   = buildQuoteExcel(q);
    const fname = encodeURIComponent(`${q.quoteNo}_${q.company || '報價單'}.xlsx`);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fname}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) {
    console.error('[QuoteExport]', e.message, e.stack);
    res.status(500).json({ error: '報價單產生失敗：' + e.message });
  }
});

// ── 全域錯誤處理（必須在所有路由之後）─────────────────────
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  // 記錄到 console，不輸出 stack 給用戶端
  console.error('[ERROR]', new Date().toISOString(), req.method, req.path,
    err.status || err.statusCode || 500, err.message);
  const statusCode = err.status || err.statusCode || 500;
  // 生產環境不暴露錯誤細節
  const message = process.env.NODE_ENV === 'production' ? '伺服器內部錯誤' : err.message;
  if (!res.headersSent) res.status(statusCode).json({ error: message });
});

// ── 管理儀表板 API ────────────────────────────────────────

function getOwnerOptions(req) {
  const auth = loadAuth();
  const allOwners = getViewableOwners(req, 'opportunities');
  return allOwners.length > 1
    ? allOwners.map(u => {
        const usr = auth.users.find(x => x.username === u);
        return { username: u, displayName: usr ? (usr.displayName || u) : u };
      })
    : [];
}

// ── 主管首頁 Dashboard ─────────────────────────────────
// 達成儀表盤 / 本月可成交 / 商機 Aging / 客戶 TOP 10
app.get('/api/manager-home', requireAuth, (req, res) => {
  try {
    const role = req.session.user.role;
    if (!['manager1','manager2','admin'].includes(role)) {
      return res.status(403).json({ error: '權限不足' });
    }
    const yearNum = parseInt(req.query.year) || new Date().getFullYear();
    const ownerFilter = req.query.owner || '';

    const data = db.load();
    const allOwners = getViewableOwners(req, 'opportunities');
    const owners = (ownerFilter && allOwners.includes(ownerFilter)) ? [ownerFilter] : allOwners;

    const opps = (data.opportunities || []).filter(o => owners.includes(o.owner));
    // manager1 的目標由部屬加總（排除自身），避免手動設定的殘留值影響計算
    const targetOwners = role === 'manager1'
      ? owners.filter(u => u !== req.session.user.username)
      : owners;
    const targets = (data.targets || []).filter(t => targetOwners.includes(t.owner) && t.year === yearNum);

    // ── 1. 業績達成度 ──
    const totalTarget = targets.reduce((s, t) => s + (parseFloat(t.amount) || 0), 0);
    const achieved = opps
      .filter(o => o.stage === 'Won')
      .filter(o => {
        const d = new Date(o.achievedDate || o.updatedAt || o.createdAt);
        return d.getFullYear() === yearNum;
      })
      .reduce((s, o) => s + (parseFloat(o.amount) || 0), 0);
    const achievementPct = totalTarget > 0 ? Math.round((achieved / totalTarget) * 100) : null;

    // ── 2. 本月可望成交（expectedDate 落在當月、stage 非 Won/D）──
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
    const monthDeals = opps.filter(o => {
      if (!o.expectedDate || ['Won','D'].includes(o.stage)) return false;
      const ed = new Date(o.expectedDate);
      return ed >= monthStart && ed <= monthEnd;
    });
    const daysSince = (createdAt) => Math.floor((Date.now() - new Date(createdAt).getTime()) / 86400000);
    const slim = (o) => ({
      id: o.id, company: o.company, product: o.product || o.category || '',
      amount: parseFloat(o.amount) || 0, stage: o.stage,
      expectedDate: o.expectedDate, days: daysSince(o.createdAt), owner: o.owner
    });
    const confirmed = monthDeals.filter(o => o.stage === 'A' && daysSince(o.createdAt) <= 60);
    const atRisk    = monthDeals.filter(o => ['A','B'].includes(o.stage) && daysSince(o.createdAt) > 60);
    const uncertain = monthDeals.filter(o => o.stage === 'B' && daysSince(o.createdAt) <= 60);
    const sumAmt = arr => arr.reduce((s, o) => s + (parseFloat(o.amount) || 0), 0);

    // ── 3. 商機 Aging（依 stage × 天數區間，計件數 + 細項）──
    const stages = ['A','B','C','D'];
    const buckets = ['0-7','8-30','31-60','61-90','90+'];
    const agingOwner = req.query.agingOwner || '';
    const agingOwners = (agingOwner && allOwners.includes(agingOwner)) ? [agingOwner] : allOwners;
    const agingOpps = (data.opportunities || []).filter(o => agingOwners.includes(o.owner));
    const aging = {};
    const agingItems = {};
    stages.forEach(s => {
      aging[s] = {};
      agingItems[s] = {};
      buckets.forEach(b => { aging[s][b] = 0; agingItems[s][b] = []; });
    });
    agingOpps.filter(o => stages.includes(o.stage)).forEach(o => {
      const d = daysSince(o.createdAt);
      const b = d <= 7 ? '0-7' : d <= 30 ? '8-30' : d <= 60 ? '31-60' : d <= 90 ? '61-90' : '90+';
      aging[o.stage][b]++;
      agingItems[o.stage][b].push({ company: o.company || '—', product: o.product || o.category || '' });
    });
    // 算出「需介入」案件數（停滯 >60 天且階段 >=C）
    const stalledCount = agingOpps.filter(o => ['C','B','A'].includes(o.stage) && daysSince(o.createdAt) > 60).length;

    // ── 4. 客戶 TOP 10（歷史成交 + 在手商機 累計金額）──
    const byCompany = {};
    opps.forEach(o => {
      const c = o.company || '（未填公司）';
      if (!byCompany[c]) byCompany[c] = { company: c, total: 0, won: 0, active: 0, count: 0 };
      const amt = parseFloat(o.amount) || 0;
      byCompany[c].total += amt;
      byCompany[c].count++;
      if (o.stage === 'Won') byCompany[c].won += amt;
      else if (['C','B','A'].includes(o.stage)) byCompany[c].active += amt;
    });
    const topCustomers = Object.values(byCompany)
      .sort((a, b) => b.total - a.total)
      .slice(0, 10);

    // ── 業務篩選器選項（同 exec dashboard 模式）──
    const auth = loadAuth();
    const ownerOptions = (auth.users || [])
      .filter(u => allOwners.includes(u.username))
      .map(u => ({ username: u.username, displayName: u.displayName || u.username }));

    res.json({
      year: yearNum,
      achievement: { target: totalTarget, achieved, pct: achievementPct },
      thisMonthCommit: {
        confirmed: { count: confirmed.length, amount: sumAmt(confirmed), items: confirmed.map(slim) },
        atRisk:    { count: atRisk.length,    amount: sumAmt(atRisk),    items: atRisk.map(slim) },
        uncertain: { count: uncertain.length, amount: sumAmt(uncertain), items: uncertain.map(slim) },
      },
      aging: { stages, buckets, data: aging, items: agingItems, stalledCount },
      topCustomers,
      ownerOptions,
    });
  } catch (e) {
    console.error('[manager-home]', e);
    res.status(500).json({ error: '載入失敗：' + e.message });
  }
});

// 轉換率漏斗
app.get('/api/exec/conversion', requireAuth, (req, res) => {
  const { year, owner } = req.query;
  const data = db.load();
  const allOwners = getViewableOwners(req, 'opportunities');
  const owners = (owner && allOwners.includes(owner)) ? [owner] : allOwners;
  const yearNum = year ? parseInt(year) : new Date().getFullYear();

  const opps = (data.opportunities || []).filter(o => owners.includes(o.owner));
  const lost = (data.lostOpportunities || []).filter(o => {
    if (!owners.includes(o.owner)) return false;
    const d = o.deletedAt || o.createdAt;
    return d && new Date(d).getFullYear() === yearNum;
  });

  const wonOpps = opps.filter(o => {
    if (o.stage !== 'Won') return false;
    const d = o.achievedDate || o.updatedAt || o.createdAt;
    return d && new Date(d).getFullYear() === yearNum;
  });

  const totalWon = wonOpps.length;
  const totalLost = lost.length;
  const totalClosed = totalWon + totalLost;
  const winRate = totalClosed > 0 ? Math.round(totalWon / totalClosed * 1000) / 10 : null;

  const cycleDays = wonOpps
    .filter(o => o.achievedDate && o.createdAt)
    .map(o => Math.round((new Date(o.achievedDate) - new Date(o.createdAt)) / 86400000))
    .filter(d => d >= 0);
  const avgCycleDays = cycleDays.length > 0
    ? Math.round(cycleDays.reduce((a, b) => a + b, 0) / cycleDays.length) : null;

  const TRANSITIONS = [
    { from: 'C', to: 'B' },
    { from: 'B', to: 'A' },
    { from: 'A', to: 'Won' },
  ];

  const stages = TRANSITIONS.map(({ from, to }) => {
    const atFrom = new Set();
    const movedOn = new Set();
    const daysArr = [];

    opps.forEach(o => {
      const hist = (o.stageHistory || []).slice()
        .sort((a, b) => (a.date || '').localeCompare(b.date || ''));
      if (o.stage === from) atFrom.add(o.id);
      hist.forEach((h, i) => {
        if (h.from === from || h.to === from) atFrom.add(o.id);
        if (h.from === from && h.to === to) {
          movedOn.add(o.id);
          // 找進入 from 的時間
          let entryDate = null;
          for (let j = i - 1; j >= 0; j--) {
            if (hist[j].to === from) { entryDate = hist[j].date; break; }
          }
          if (!entryDate && from === 'C') entryDate = o.createdAt;
          if (entryDate) {
            const days = Math.round((new Date(h.date) - new Date(entryDate)) / 86400000);
            if (days >= 0 && days < 1000) daysArr.push(days);
          }
        }
      });
    });

    const rate = atFrom.size > 0 ? Math.round(movedOn.size / atFrom.size * 1000) / 10 : null;
    const avgDays = daysArr.length > 0
      ? Math.round(daysArr.reduce((a, b) => a + b, 0) / daysArr.length) : null;
    return { from, to, count: movedOn.size, total: atFrom.size, rate, avgDays };
  });

  const stageCounts = ['C', 'B', 'A', 'Won'].reduce((acc, s) => {
    acc[s] = opps.filter(o => o.stage === s).length;
    return acc;
  }, {});

  res.json({ winRate, avgCycleDays, stages, totalClosed, totalWon, totalLost, stageCounts, ownerOptions: getOwnerOptions(req) });
});

// 月度業績趨勢（近 24 個月）
app.get('/api/exec/trend', requireAuth, (req, res) => {
  const { owner } = req.query;
  const data = db.load();
  const allOwners = getViewableOwners(req, 'opportunities');
  const owners = (owner && allOwners.includes(owner)) ? [owner] : allOwners;

  const wonOpps = (data.opportunities || []).filter(o =>
    o.stage === 'Won' && owners.includes(o.owner)
  );

  const now = new Date();
  const months = [];
  for (let i = 23; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  const map = {};
  months.forEach(m => { map[m] = { month: m, amount: 0, count: 0 }; });
  wonOpps.forEach(o => {
    const month = (o.achievedDate || '').slice(0, 7);
    if (map[month]) {
      map[month].amount += parseFloat(o.amount) || 0;
      map[month].count++;
    }
  });
  res.json(months.map(m => map[m]));
});

// 產品 / BU 分析
app.get('/api/exec/product-analysis', requireAuth, (req, res) => {
  const { year, owner } = req.query;
  const data = db.load();
  const allOwners = getViewableOwners(req, 'opportunities');
  const owners = (owner && allOwners.includes(owner)) ? [owner] : allOwners;
  const yearNum = year ? parseInt(year) : new Date().getFullYear();

  const inYear = (o, isLost) => {
    const d = isLost
      ? (o.deletedAt || o.createdAt)
      : (o.stage === 'Won' ? (o.achievedDate || o.updatedAt || o.createdAt) : (o.expectedDate || o.createdAt));
    return !yearNum || (d && new Date(d).getFullYear() === yearNum);
  };

  const opps = (data.opportunities || []).filter(o => owners.includes(o.owner) && inYear(o, false));
  const lost = (data.lostOpportunities || []).filter(o => owners.includes(o.owner) && inYear(o, true));

  const grouped = {};
  const add = (o, isLost) => {
    const cat = (o.category || '（未分類）').trim();
    if (!grouped[cat]) grouped[cat] = { category: cat, count: 0, wonCount: 0, lostCount: 0, pipelineAmount: 0, wonAmount: 0, grossTotal: 0, grossCount: 0 };
    const g = grouped[cat];
    const amt = parseFloat(o.amount) || 0;
    if (isLost) { g.lostCount++; g.count++; return; }
    g.count++;
    if (o.stage === 'Won') { g.wonCount++; g.wonAmount += amt; }
    else { g.pipelineAmount += amt; }
    if (o.grossMarginRate) { g.grossTotal += parseFloat(o.grossMarginRate); g.grossCount++; }
  };
  opps.forEach(o => add(o, false));
  lost.forEach(o => add(o, true));

  const result = Object.values(grouped).map(g => ({
    category: g.category,
    count: g.count,
    wonCount: g.wonCount,
    lostCount: g.lostCount,
    pipelineAmount: Math.round(g.pipelineAmount * 10) / 10,
    wonAmount: Math.round(g.wonAmount * 10) / 10,
    winRate: (g.wonCount + g.lostCount) > 0
      ? Math.round(g.wonCount / (g.wonCount + g.lostCount) * 1000) / 10 : null,
    avgGrossMargin: g.grossCount > 0
      ? Math.round(g.grossTotal / g.grossCount * 10) / 10 : null,
  })).sort((a, b) => (b.wonAmount + b.pipelineAmount) - (a.wonAmount + a.pipelineAmount));

  res.json(result);
});

// ── 404 fallback ─────────────────────────────────────────
app.use((req, res) => {
  if (req.path.startsWith('/api/')) return res.status(404).json({ error: '找不到此 API 端點' });
  res.redirect('/login.html');
});

// ════════════════════════════════════════════════════════════
//  雙模啟動：本地直接 listen；Vercel serverless 只 export app
// ════════════════════════════════════════════════════════════
if (require.main === module) {
  // 本地執行：node server.js
  (async () => {
    try { await db.ready(); } catch (e) { console.error('[db] ready failed:', e); }
    try { await apiMonitor.ready(); } catch (e) { console.error('[apiMonitor] ready failed:', e); }
    app.listen(PORT, () => {
      migrateOwner();
      migrateStage成交ToWon();
      console.log(`\n✅ 業務名片管理系統已啟動`);
      console.log(`👉 請開啟瀏覽器，前往 http://localhost:${PORT}\n`);
    });
  })();
} else {
  // 被 import（如 api/index.js）：只 export app，由 serverless 呼叫
  // Postgres 模式需要先 preload data 才能同步 load()
  if (process.env.DB_BACKEND === 'postgres') {
    // 背景預熱 + 遷移（不阻塞 cold start；首個 request 會等 ready）
    db.ready()
      .then(() => {
        migrateOwner();
        migrateStage成交ToWon();
      })
      .catch((e) => console.error('[db] preload failed:', e));
    apiMonitor.ready().catch((e) => console.error('[apiMonitor] preload failed:', e));
  }
}

module.exports = app;
