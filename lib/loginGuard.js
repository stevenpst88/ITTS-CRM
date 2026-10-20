'use strict';
/**
 * 帳號鎖定（LOGIN_LOCKOUT）與閒置逾時（SESSION_IDLE_MINUTES）——內控稽核第三波第二批 ②。
 *
 * 兩個開關預設都是關閉：環境變數未設定 / 0 / 空白 → 所有函式立即回傳「不介入」，登入與連線行為與先前位元組相同。
 *
 *   LOGIN_LOCKOUT=1                 開啟帳號鎖定
 *   LOGIN_MAX_FAILS=5               連續失敗幾次鎖定（預設 5）
 *   LOGIN_FAIL_WINDOW_MINUTES=15    失敗計數的視窗（從該段連續失敗的第一次起算，預設 15 分鐘）
 *   LOGIN_LOCK_MINUTES=15           第一次鎖定時間（預設 15 分鐘；可用小數，例如 0.05＝3 秒，僅供測試）
 *   LOGIN_LOCK_MAX_MINUTES=240      重複鎖定時「加倍」的上限（預設 240 分鐘）
 *   LOGIN_LOCK_SCOPE=ip             鎖定鍵範圍（資安審查 M3）：ip（預設）＝「帳號＋來源 IP」，攻擊者只能鎖住自己的來源；name＝舊行為，只看帳號名稱
 *   SESSION_IDLE_MINUTES=60         閒置逾時分鐘數（>0 才啟用；未設定或 0＝關閉）
 *
 * 鎖定狀態（次數與時間戳，不存密碼／原因）放在主資料 blob 的 `loginFailures`，
 * 鍵為「u:」＋正規化帳號（含不存在的帳號，這樣存在與否的行為完全一致；加前綴也避免 __proto__ 之類的鍵）；
 * LOGIN_LOCK_SCOPE=ip（預設）時再接「控制字元 0x01＋來源 IP 的雜湊（不存明文 IP）」——呼叫端沒傳 IP 時退回只看帳號名稱（單元測試與舊資料）。
 * 資安審查 M3 的三項強化：① 鎖定鍵含來源 IP（匿名者無法再從自己的位址把別人的管理員鎖死）；
 * ② 共用快取後端（Postgres）下，未達鎖定的失敗計數只改記憶體中的快取，寫庫節流為每 10 秒一次（達鎖定門檻那次一定寫）——不再每次失敗登入都 UPSERT 整包 JSONB；
 * ③ 表滿時淘汰最舊的「非鎖定」條目，而不是拒收新名稱（過去 1000 個假名稱可讓真實帳號的失敗不再被追蹤）。
 *
 * 取捨（DoS-by-lockout）：任何人只要知道帳號就能讓對方被鎖。緩解：
 *   ① 鎖定只影響該帳號名稱，不影響其他帳號、也不擴及 IP；
 *   ② 已鎖定期間的嘗試不再累計、不延長鎖定；
 *   ③ 重複鎖定加倍但有上限（預設 4 小時），安靜 24 小時後倍率歸零；
 *   ④ 管理員可在後台隨時解鎖（POST /api/admin/users/:u/unlock）；緊急時把 LOGIN_LOCKOUT 關掉立即全面失效；
 *   ⑤ 既有的每 IP 登入限流（20 次／15 分鐘）仍在，攻擊者單一 IP 的速率有限。
 * 建議至少保留 2 位管理員帳號，避免唯一管理員被鎖。
 *
 * 時間來源可注入（opts.clock），單元測試用；出貨的伺服器不提供任何測試掛鉤。
 */

const crypto = require('crypto');
const PREFIX = 'u:';
const IP_SEP = String.fromCharCode(1);          // 帳號與 IP 的分隔字元：normName 會把控制字元換成空白，帳號名稱裡不可能出現
const PERSIST_MIN_MS = 10 * 1000;               // 共用快取後端下，未鎖定的失敗計數最多每 10 秒寫庫一次
const MAX_ENTRIES = 1000;
const ABSOLUTE_SESSION_MS = 8 * 60 * 60 * 1000;   // 與 middleware/jwtSession.js 的 JWT 壽命一致（8 小時絕對上限不因續期而延長）
const DECAY_MS = 24 * 60 * 60 * 1000;             // 這麼久沒有失敗，「重複鎖定加倍」的倍率歸零

function posNum(v, d) { const n = parseFloat(v); return Number.isFinite(n) && n > 0 ? n : d; }
function truthy(v) { return /^(1|true|on|yes)$/i.test(String(v == null ? '' : v).trim()); }

function readConfig(env) {
  env = env || process.env;
  const idleMin = posNum(env.SESSION_IDLE_MINUTES, 0);
  return {
    lockout: truthy(env.LOGIN_LOCKOUT),
    maxFails: Math.max(1, Math.floor(posNum(env.LOGIN_MAX_FAILS, 5))),
    windowMs: posNum(env.LOGIN_FAIL_WINDOW_MINUTES, 15) * 60000,
    lockMs: posNum(env.LOGIN_LOCK_MINUTES, 15) * 60000,
    lockCapMs: posNum(env.LOGIN_LOCK_MAX_MINUTES, 240) * 60000,
    scope: /^name$/i.test(String(env.LOGIN_LOCK_SCOPE || '').trim()) ? 'name' : 'ip',
    idleMinutes: idleMin,
    idleMs: idleMin * 60000,
  };
}

/** 帳號正規化：去控制字元、trim、小寫、限長（與實際登入比對無關，只用來當鎖定鍵） */
function normName(name) {
  const s = String(name == null ? '' : name).split('').map((ch) => { const c = ch.charCodeAt(0); return (c < 32 || c === 127) ? ' ' : ch; }).join('');
  return s.trim().toLowerCase().slice(0, 64);
}
const keyOf = (name) => PREFIX + normName(name);
/** 來源 IP → 鎖定鍵的一部分：不把 IP 明文放進主資料 blob，只存 sha256 前 16 碼；空值＝沒有 IP */
function normIp(ip) {
  const s = String(ip == null ? '' : ip).trim().toLowerCase().slice(0, 64);
  return s ? 'h' + crypto.createHash('sha256').update(s).digest('hex').slice(0, 16) : '';
}
/** 某帳號名稱底下的所有鍵：舊的純帳號鍵與各 IP 的鍵 */
function belongsTo(k, name) { const b = keyOf(name); return k === b || k.startsWith(b + IP_SEP); }
function nameOfKey(k) { const body = k.slice(PREFIX.length); const i = body.indexOf(IP_SEP); return i < 0 ? body : body.slice(0, i); }

/** 輪詢型端點：背景自動呼叫，不代表使用者還在操作，不可刷新閒置計時（GET 才算；其餘方法一律算活動） */
const IDLE_POLL_PATHS = ['/api/poll-bundle', '/api/notifications', '/api/announcements/active'];
function isPollingRequest(method, path) {
  if (String(method).toUpperCase() !== 'GET' && String(method).toUpperCase() !== 'HEAD') return false;
  const p = String(path || '').toLowerCase().replace(/\/+$/, '');
  return IDLE_POLL_PATHS.includes(p);
}
function isPageRequest(method, path) {
  if (String(method).toUpperCase() !== 'GET') return false;
  const p = String(path || '');
  return p === '/' || /\.html$/i.test(p);
}

function create(opts) {
  opts = opts || {};
  const cfg = opts.config || readConfig(opts.env);
  const clock = typeof opts.clock === 'function' ? opts.clock : Date.now;
  const store = opts.store || { load: () => ({}), save: () => {} };
  const shared = !!opts.sharedStore;   // true＝store.load() 回共用快取（Postgres）：就地修改即生效，可節流寫庫；false＝每次回副本（JSON），每次都要存
  let lastPersistAt = 0;
  const kOf = (name, ip) => { const b = keyOf(name); const p = cfg.scope === 'ip' ? normIp(ip) : ''; return p ? b + IP_SEP + p : b; };

  // 注意：JSON 模式的 db.load() 每次都回新的副本（Postgres 模式回共用快取）。所以「載入 → 修改 → 把同一個物件存回」必須用同一份 d。
  function ctx(createIfMissing) {
    const d = store.load();
    if (!d || typeof d !== 'object') return null;
    if (!d.loginFailures || typeof d.loginFailures !== 'object' || Array.isArray(d.loginFailures)) {
      if (!createIfMissing) return null;
      d.loginFailures = {};
    }
    return { d, m: d.loginFailures };
  }
  function persist(c, force) {
    if (shared && !force) { const t = clock(); if (t - lastPersistAt < PERSIST_MIN_MS) return; lastPersistAt = t; }
    try { store.save(c.d); } catch (e) { console.error('[loginGuard save]', e && e.message); }
  }

  function prune(m, t) {
    for (const k of Object.keys(m)) {
      const e = m[k];
      if (!e || typeof e !== 'object') { delete m[k]; continue; }
      const quiet = t - (e.last || 0);
      const keep = (e.lockUntil || 0) > t || quiet <= (e.lockCount > 0 ? DECAY_MS : cfg.windowMs);
      if (!keep) delete m[k];
    }
  }

  const api = {
    config: cfg,
    lockoutEnabled: cfg.lockout,
    idleEnabled: cfg.idleMs > 0,
    normName,

    /** 此帳號名稱目前是否鎖定中（不寫入） */
    checkLocked(name, ip) {
      if (!cfg.lockout) return { locked: false };
      const c = ctx(false); if (!c) return { locked: false };
      const e = c.m[kOf(name, ip)]; const t = clock();
      if (e && (e.lockUntil || 0) > t) return { locked: true, until: e.lockUntil, lockCount: e.lockCount || 0 };
      return { locked: false };
    },

    /** 記一次登入失敗（任何原因；帳號不存在也一樣）。回傳 { locked, justLocked, until, lockCount, failCount } */
    recordFailure(name, ip) {
      if (!cfg.lockout) return { locked: false, justLocked: false };
      const t = clock(); const k = kOf(name, ip);
      const c = ctx(true); if (!c) return { locked: false, justLocked: false };
      const m = c.m;
      let e = m[k];
      if (e && (e.lockUntil || 0) > t) return { locked: true, justLocked: false, until: e.lockUntil, lockCount: e.lockCount || 0 };   // 鎖定中不累計、不延長
      if (!e) {
        if (Object.keys(m).length >= MAX_ENTRIES) {
          prune(m, t);
          if (Object.keys(m).length >= MAX_ENTRIES) {
            // 資安審查 M3：表滿 → 淘汰最舊的「非鎖定」條目騰出位置（過去直接拒收新名稱，匿名者灌滿 1000 個假名稱後真實帳號的失敗就不再被追蹤）
            let oldest = null, oldestAt = Infinity;
            for (const kk of Object.keys(m)) { const x = m[kk]; if ((x.lockUntil || 0) > t) continue; if ((x.last || 0) < oldestAt) { oldestAt = x.last || 0; oldest = kk; } }
            if (oldest === null) return { locked: false, justLocked: false, dropped: true };   // 極端：全部是「鎖定中」的條目，新名稱只靠 IP 限流
            delete m[oldest];
          }
        }
        e = m[k] = { n: 0, first: t, last: t, lockUntil: 0, lockCount: 0 };
      }
      if (t - (e.last || 0) > DECAY_MS) e.lockCount = 0;
      if (!e.n || t - (e.first || 0) > cfg.windowMs) { e.n = 0; e.first = t; }
      e.n++; e.last = t;
      let justLocked = false;
      if (e.n >= cfg.maxFails) {
        e.lockCount = (e.lockCount || 0) + 1;
        const dur = Math.min(cfg.lockMs * Math.pow(2, e.lockCount - 1), Math.max(cfg.lockCapMs, cfg.lockMs));
        e.lockUntil = t + dur; e.n = 0; justLocked = true;
      }
      persist(c, justLocked);   // 剛達鎖定門檻那次一定寫庫（其他實例才看得到鎖）；其餘在共用快取後端下節流
      return { locked: justLocked, justLocked, until: e.lockUntil || 0, lockCount: e.lockCount || 0, failCount: e.n };
    },

    /** 登入成功：清掉該帳號的失敗紀錄（沒有紀錄就不寫入，正常登入零額外 DB 寫入） */
    recordSuccess(name, ip) {
      if (!cfg.lockout) return false;
      const c = ctx(false); if (!c) return false;
      const k = kOf(name, ip);
      if (!c.m[k]) return false;
      delete c.m[k]; persist(c, true); return true;
    },

    /** 管理員解鎖（連同加倍倍率一併清除）；回傳是否有東西被清掉 */
    unlock(name) {
      const c = ctx(false); if (!c) return false;
      const ks = Object.keys(c.m).filter((k) => belongsTo(k, name));   // 舊的純帳號鍵＋各來源 IP 的鍵一次全清
      if (!ks.length) return false;
      ks.forEach((k) => { delete c.m[k]; }); persist(c, true); return true;
    },

    /** 目前鎖定中的清單（給管理員看；不含密碼／IP） */
    listLocks() {
      if (!cfg.lockout) return [];
      const c = ctx(false); if (!c) return [];
      const m = c.m; const t = clock();
      const by = new Map();   // 同一帳號名稱可能同時有多個來源 IP 的鎖：合併成一列（取最晚到期者；回傳形狀與舊版相同）
      Object.keys(m).filter(k => k.startsWith(PREFIX) && m[k] && (m[k].lockUntil || 0) > t).forEach((k) => {
        const u = nameOfKey(k); const cur = by.get(u);
        if (!cur) by.set(u, { username: u, lockUntil: m[k].lockUntil, lockCount: m[k].lockCount || 0, last: m[k].last || 0 });
        else { if (m[k].lockUntil > cur.lockUntil) { cur.lockUntil = m[k].lockUntil; cur.lockCount = m[k].lockCount || 0; } if ((m[k].last || 0) > cur.last) cur.last = m[k].last || 0; }
      });
      return [...by.values()].sort((a, b) => b.lockUntil - a.lockUntil);
    },
    /** 單一帳號的鎖定資訊（管理員帳號列表的「已鎖定」徽章用）；未鎖定回 null */
    getLock(name) {
      if (!cfg.lockout) return null;
      const c = ctx(false); if (!c) return null;
      const t = clock(); let best = null;
      Object.keys(c.m).forEach((k) => { const e = c.m[k]; if (belongsTo(k, name) && e && (e.lockUntil || 0) > t && (!best || e.lockUntil > best.lockUntil)) best = { lockUntil: e.lockUntil, lockCount: e.lockCount || 0 }; });
      return best;
    },

    // ───────── 閒置逾時 ─────────
    /**
     * 判斷一個 session 使用者在 now 時的狀態。
     *  expired: 'idle'（閒置超過上限）｜'absolute'（登入已滿 8 小時）｜false
     *  refresh: 是否要把 lat 更新為 now（並由呼叫端重簽 cookie）
     * 沒有 lat 也沒有 loginAt 的舊 session：視為剛活動，並刷新。
     */
    idleDecision(su, now, counts) {
      if (!cfg.idleMs) return { expired: false, refresh: false };
      const loginAt = Number(su && su.loginAt) || 0;
      if (loginAt && now - loginAt > ABSOLUTE_SESSION_MS) return { expired: 'absolute', refresh: false };
      const base = Number(su && su.lat) || loginAt;
      if (base && base <= now && now - base > cfg.idleMs) return { expired: 'idle', refresh: false };
      if (!counts) return { expired: false, refresh: false };
      const every = Math.min(60000, cfg.idleMs / 4);
      const stale = !base || (now - base) >= every;
      return { expired: false, refresh: stale };
    },

    /**
     * Express 中介層（放在「以 DB 現值覆寫 session」之後）。
     * o.isJwt：true＝JWT cookie（指派 req.session.user 會重簽 cookie，無 DB 寫入）；false＝express-session（就地改物件即由 store 保存）
     * o.onExpire(req, username, kind)：記稽核
     */
    idleMiddleware(o) {
      o = o || {};
      return function idleGuard(req, res, next) {
        if (!cfg.idleMs) return next();
        const su = req.session && req.session.user;
        if (!su || !su.username) return next();
        const p = req.path || '';
        const isApi = p.startsWith('/api/');
        if (isApi) { if (p === '/api/login' || p === '/api/logout') return next(); }
        else if (!isPageRequest(req.method, p)) return next();   // 靜態資產不算活動、也不在這裡擋
        const t = clock();
        const d = api.idleDecision(su, t, !isPollingRequest(req.method, p));
        if (d.expired) {
          try { if (o.onExpire) o.onExpire(req, su.username, d.expired); } catch (e) { console.error('[idle onExpire]', e && e.message); }
          const finish = () => {
            if (isApi) {
              return res.status(401).json({
                error: d.expired === 'idle' ? '閒置過久，已自動登出，請重新登入' : '登入已逾時，請重新登入',
                code: d.expired === 'idle' ? 'SESSION_IDLE_TIMEOUT' : 'SESSION_EXPIRED'
              });
            }
            next();
          };
          try { req.session.destroy(() => { try { res.clearCookie('connect.sid'); } catch (e) { /* ignore */ } finish(); }); }
          catch (e) { finish(); }
          return;
        }
        if (d.refresh) {
          su.lat = t;
          // JWT：重簽 cookie；但沒有 loginAt 的舊 token 不重簽（否則 8 小時絕對上限會被續掉）
          if (o.isJwt && su.loginAt) { try { req.session.user = su; } catch (e) { console.error('[idle resign]', e && e.message); } }
        }
        next();
      };
    },
  };
  return api;
}

module.exports = { create, readConfig, normName, isPollingRequest, isPageRequest, IDLE_POLL_PATHS, ABSOLUTE_SESSION_MS, MAX_ENTRIES };
