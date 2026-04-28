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
  windowMs: 15 * 60 * 1000, // 15 分鐘
  max: 20,                   // 最多嘗試 20 次
  message: { success: false, message: '嘗試次數過多，請 15 分鐘後再試' },
  standardHeaders: true,
  legacyHeaders: false
});

// ── API 全域速率限制 ──────────────────────────────────────
const apiLimiter = rateLimit({
  windowMs: 60 * 1000,  // 1 分鐘
  max: 300,             // 每 IP 每分鐘最多 300 次 API 呼叫
  message: { error: '請求過於頻繁，請稍後再試' },
  standardHeaders: true,
  legacyHeaders: false,
  skip: req => /\.(css|js|png|svg|ico|woff2?)$/.test(req.path) // 靜態資源不計入
});

app.use(express.json({ limit: '2mb' }));        // 限制 request body 大小
app.use(express.urlencoded({ extended: false, limit: '2mb' }));

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
const CONTRACT_FIELDS  = ['contractNo','company','contactName','product','startDate','endDate','amount','salesPerson','note','type'];
const RECEIVABLE_FIELDS= ['company','contactName','invoiceNo','invoiceDate','dueDate','amount','paidAmount','currency','note','status'];

// ── 網址安全驗證 ──────────────────────────────────────────
function sanitizeUrl(url) {
  if (!url) return '';
  const trimmed = url.trim();
  // 只接受 http:// 或 https:// 開頭，防止 javascript: 注入
  return /^https?:\/\//i.test(trimmed) ? trimmed : '';
}

// ── 公開路由（不需驗證，必須在 requireAuth 之前）─────────
app.use('/login.html', express.static(path.join(__dirname, '_client', 'login.html')));
app.use('/itts-logo.png', express.static(path.join(__dirname, '_client', 'itts-logo.png')));
app.use('/itts-logo.svg', express.static(path.join(__dirname, '_client', 'itts-logo.svg')));
// admin.html 已移至受保護路由（需登入 + admin 角色）

app.post('/api/login', loginLimiter, async (req, res) => {
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
});

app.post('/api/logout', (req, res) => {
  req.session.destroy(() => {
    res.clearCookie('connect.sid');
    res.json({ success: true });
  });
});

// ── 受保護路由（需驗證）─────────────────────────────────
// admin.html：需登入且需 admin 角色，且帳號必須為啟用狀態
app.get('/admin.html', requireAuth, (req, res) => {
  const auth = loadAuth();
  const user = auth.users.find(u => u.username === req.session.user.username);
  if (!user || user.role !== 'admin' || user.active === false) return res.redirect('/');
  res.sendFile(path.join(__dirname, '_client', 'admin.html'));
});
app.use(requireAuth, express.static(path.join(__dirname, '_client')));
// ── 上傳檔案路由：改由 storage 模組代理（支援本地/Supabase）──
app.get('/uploads/:key', requireAuth, (req, res, next) => {
  Promise.resolve(storage.serveFile(req, res, req.params.key)).catch(next);
});

app.get('/api/me', requireAuth, (req, res) => {
  res.json(req.session.user);
});

// ── 自助更改密碼（登入者本人）────────────────────────────
app.put('/api/user/password', requireAuth, async (req, res) => {
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

    const text = result.response.text();
    try {
      const data = gemini.parseJson(text);
      res.json(data);
    } catch {
      res.status(500).json({ error: 'AI 回應格式錯誤，請重試' });
    }
  } catch (e) {
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
    const STAGE_LABEL_AI = { D:'D（靜止）', C:'C（Pipeline）', B:'B（Upside）', A:'A（Commit）', '成交':'成交', Won:'Won' };

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
      o.contactId === contactId && o.stage !== 'D' && o.stage !== '成交' && o.stage !== 'Won'
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
    const is429 = e.status === 429 || String(e.message).includes('429') || String(e.message).includes('RESOURCE_EXHAUSTED'); const status = is429 ? 429 : 500;
    const msg = is429 ? 'AI 服務暫時忙碌，請稍後再試' : ('AI 發生錯誤：' + e.message);
    res.status(status).json({ error: msg });
  }
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
  res.json(data.visits[idx]);
});

app.delete('/api/visits/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.visits) return res.json({ success: true });
  data.visits = data.visits.filter(v => !(v.id === req.params.id && v.owner === owner));
  db.save(data);
  res.json({ success: true });
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
    return res.json(existing);
  }
  const target = { id: uuidv4(), owner, year, amount, createdAt: new Date().toISOString() };
  data.targets.push(target);
  db.save(data);
  res.status(201).json(target);
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
    contactId:    req.body.contactId    || '',
    contactName:  req.body.contactName  || '',
    company:      req.body.company      || '',
    category:     req.body.category     || '',
    product:      req.body.product      || '',
    amount:       req.body.amount       || '',
    expectedDate: req.body.expectedDate || '',
    description:  req.body.description  || '',
    stage:        'C',
    visitId:      req.body.visitId      || '',
    createdAt:    new Date().toISOString()
  };
  data.opportunities.push(opp);
  db.save(data);
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
  data.opportunities[idx] = { ...data.opportunities[idx], ...pickFields(req.body, OPP_FIELDS), id: req.params.id, owner };
  const newStage = data.opportunities[idx].stage;
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
  res.json(data.opportunities[idx]);
});

// ── 商機動態報表 ──────────────────────────────────────────
// 從低到高排列：D(靜止) → C(Pipeline) → B(Upside) → A(Commit) → 成交
const STAGE_ORDER = ['D','C','B','A','成交','Won'];
const STAGE_LABEL = { D:'D｜靜止中', C:'C｜Pipeline', B:'B｜Upside', A:'A｜Commit', '成交':'✅ 成交', Won:'🏆 Won' };
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
      totalPipeline: sum(opps.filter(o=>o.stage!=='成交'&&o.stage!=='Won'&&o.stage!=='D')),
      totalCount:    opps.filter(o=>o.stage!=='成交'&&o.stage!=='Won'&&o.stage!=='D').length,
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
    .filter(o => owners.includes(o.owner) && o.stage !== '成交' && o.stage !== 'Won')
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
const STAGE_CONF       = { D: 10, C: 25, B: 50, A: 90, '成交': 100, Won: 100 };
const STAGE_LABEL_EXPORT = { A: 'Commit', B: 'Upside', C: 'Pipeline', '成交': '成交', Won: 'Won' };

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

  const STAGE_LABELS = { A: 'Commit', B: 'Upside', C: 'Pipeline', C2: 'Pipeline', D: 'D', '成交': '成交', Won: 'Won' };

  const headers = [
    '客戶名稱', '銷售案名', 'BU(category)', '預定簽約日',
    '業務帳號(owner)', '業務姓名', '把握度階段(A/B/C/成交)',
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

    const VALID_STAGES = new Set(['A', 'B', 'C', '成交', 'D', 'Won']);

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
      const stageMap = { 'Commit': 'A', 'commit': 'A', 'Upside': 'B', 'upside': 'B', 'Pipeline': 'C', 'pipeline': 'C', 'won': 'Won', '成交': '成交' };
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
  res.status(201).json(c);
});

app.put('/api/contracts/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.contracts) data.contracts = [];
  const idx = data.contracts.findIndex(c => c.id === req.params.id && c.owner === owner);
  if (idx === -1) return res.status(404).json({ error: '找不到此合約' });
  data.contracts[idx] = { ...data.contracts[idx], ...pickFields(req.body, CONTRACT_FIELDS), id: req.params.id, owner };
  db.save(data);
  res.json(data.contracts[idx]);
});

app.delete('/api/contracts/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  if (!data.contracts) return res.json({ success: true });
  data.contracts = data.contracts.filter(c => !(c.id === req.params.id && c.owner === owner));
  db.save(data);
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
    listedType: '未上市櫃', stockCode: '', exchange: '',
    revenue2025: 'N/A', grossMargin2025: 'N/A',
    revenue2024: 'N/A', grossMargin2024: 'N/A',
    eps: 'N/A', epsYear: '',
  };

  // 1. 經濟部商業司
  try {
    const r = await fetch(
      `https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq ${taxId}&$skip=0&$top=1`,
      { headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }, signal: AbortSignal.timeout(8000) }
    );
    if (r.ok) {
      const data = await r.json();
      if (data?.[0]) {
        result.companyName  = data[0].Company_Name || '';
        result.representative = data[0].Responsible_Name || '';
        result.capital      = formatCapital(data[0].Capital_Stock_Amount);
        result.address      = data[0].Company_Location || '';
        result.companyStatus = data[0].Company_Status_Desc || '';
      }
    }
  } catch (e) { /* ignore */ }

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
  res.json(data.receivables[idx]);
});

app.delete('/api/receivables/:id', requireAuth, (req, res) => {
  const owner = req.session.user.username;
  const data = db.load();
  data.receivables = (data.receivables || []).filter(r => !(r.id === req.params.id && r.owner === owner));
  db.save(data);
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
    app.listen(PORT, () => {
      migrateOwner();
      console.log(`\n✅ 業務名片管理系統已啟動`);
      console.log(`👉 請開啟瀏覽器，前往 http://localhost:${PORT}\n`);
    });
  })();
} else {
  // 被 import（如 api/index.js）：只 export app，由 serverless 呼叫
  // Postgres 模式需要先 preload data 才能同步 load()
  if (process.env.DB_BACKEND === 'postgres') {
    // 背景預熱（不阻塞 cold start；首個 request 會等）
    db.ready().catch((e) => console.error('[db] preload failed:', e));
  }
}

module.exports = app;
