/**
 * Web Push 推播輔助
 *
 * 訂閱資料結構（存在 auth.json 的每個 user 物件下）：
 *   user.pushSubscriptions: [
 *     {
 *       endpoint: 'https://fcm.googleapis.com/fcm/send/...',
 *       keys: { p256dh: '...', auth: '...' },
 *       ua: 'Mozilla/5.0 ...',
 *       addedAt: '2026-05-09T...'
 *     },
 *     ...
 *   ]
 *
 * 同一個 user 可以有多個訂閱（桌機 Chrome、手機 PWA 各一）。
 *
 * 環境變數：
 *   VAPID_PUBLIC_KEY   公鑰（前端 subscribe 時用）
 *   VAPID_PRIVATE_KEY  私鑰（伺服器送推播時用）
 *   VAPID_SUBJECT      mailto: 或 https:// URL
 *
 * 若上述任一缺失，模組會把所有送推播動作當成 noop（不丟錯）— 方便本機開發或未設定環境直接跑。
 */

const webpush = require('web-push');
const dns = require('dns');
const https = require('https');
const SU = require('./safeUrl');

const PUBLIC_KEY  = process.env.VAPID_PUBLIC_KEY;
const PRIVATE_KEY = process.env.VAPID_PRIVATE_KEY;
const SUBJECT     = process.env.VAPID_SUBJECT || 'mailto:admin@itts.com.tw';

const enabled = !!(PUBLIC_KEY && PRIVATE_KEY);

if (enabled) {
  webpush.setVapidDetails(SUBJECT, PUBLIC_KEY, PRIVATE_KEY);
} else {
  console.warn('[pushNotify] 未設定 VAPID_PUBLIC_KEY / VAPID_PRIVATE_KEY，Web Push 已停用');
}

// ───────── 推播端點的 SSRF 防護（SEC-08b）─────────
// 訂閱的 endpoint 是使用者（瀏覽器）提交的網址，伺服器之後會對它送出請求。沒驗證的話，任何登入者都能讓伺服器去連內網／metadata（盲 SSRF）。
// 做法：不寫死推播服務白名單（FCM／Mozilla autopush／WNS／APNs 的網域會變，寫死會誤擋）；改用 lib/safeUrl.js 的通用規則：
//   ① 訂閱時：只允許 https、不可帶帳密、只允許 443、網域解析後每個位址都必須是公開位址（validateEndpoint）
//   ② 送出時：再驗一次（防止舊的／被改過的訂閱），且 web-push 的連線走「連線時才解析並檢查」的 agent（makeSafeAgent）：
//      實際連線用的就是剛檢查過的位址（沒有驗證後被換掉的空窗），web-push 本身不跟轉址，所以不需逐跳檢查。
const PUSH_PORTS = ['443'];
const PUSH_TIMEOUT_MS = 10000;

/** 驗證訂閱 endpoint。回傳 { ok:true } 或 { ok:false, reason, message }（中文）。o.lookup 供測試注入。 */
async function validateEndpoint(endpoint, o) {
  if (typeof endpoint !== 'string' || !endpoint || endpoint.length > 2048) return { ok: false, reason: 'BAD_URL', message: SU.REASON_ZH.BAD_URL };
  const v = await SU.validateUrl(endpoint, Object.assign({}, o, { ports: PUSH_PORTS }));
  return v.ok ? { ok: true } : { ok: false, reason: v.reason, message: v.message };
}

/** 連線時用的 DNS lookup：解析 → 任一位址不是公開位址就拒絕（錯誤碼 ESSRF）。resolver 供測試注入。 */
function makeSafeLookup(resolver) {
  return function safeLookup(hostname, options, cb) {
    if (typeof options === 'function') { cb = options; options = {}; }
    const finish = (err, addrs) => {
      if (err) return cb(err);
      if (!Array.isArray(addrs) || !addrs.length || addrs.some((a) => SU.isBlockedIp(a.address)) || SU.isLocalName(hostname)) {
        return cb(Object.assign(new Error('推播端點解析到不允許的位址'), { code: 'ESSRF' }));
      }
      if (options && options.all) return cb(null, addrs);
      cb(null, addrs[0].address, addrs[0].family || (addrs[0].address.includes(':') ? 6 : 4));
    };
    if (resolver) Promise.resolve(resolver(hostname)).then((a) => finish(null, a), (e) => finish(e));
    else dns.lookup(hostname, { all: true }, finish);
  };
}
function makeSafeAgent(resolver) {
  return new https.Agent({ keepAlive: false, lookup: makeSafeLookup(resolver) });
}

function isEnabled() {
  return enabled;
}

function getPublicKey() {
  return PUBLIC_KEY || null;
}

/**
 * 送推播給單一訂閱。
 * 回傳：{ ok: true } 成功；{ ok: false, gone: true, error } 表示 endpoint 已失效，呼叫端應移除；
 *       { ok: false, gone: false, error } 為其他暫時性錯誤。
 */
async function sendToSubscription(subscription, payload, o) {
  if (!enabled) return { ok: false, gone: false, error: 'web push disabled' };
  o = o || {};
  const wp = o.webpush || webpush;
  try {
    // 送出前再驗一次：網址形狀／內網位址不合格的訂閱永遠不會合法 → 視為失效（gone），呼叫端會把它移除
    const chk = await validateEndpoint(subscription && subscription.endpoint, { lookup: o.lookup });
    if (!chk.ok) {
      if (chk.reason === 'DNS_FAILED') return { ok: false, gone: false, error: chk.message };
      return { ok: false, gone: true, error: chk.message };
    }
    const body = typeof payload === 'string' ? payload : JSON.stringify(payload);
    await wp.sendNotification(subscription, body, { TTL: 60 * 60 * 24, timeout: PUSH_TIMEOUT_MS, agent: makeSafeAgent(o.lookup) });
    return { ok: true };
  } catch (e) {
    if (e && e.code === 'ESSRF') return { ok: false, gone: true, error: e.message };
    // 410 Gone / 404 Not Found = subscription 已被瀏覽器撤銷，永遠不會再送達
    const gone = e && (e.statusCode === 410 || e.statusCode === 404);
    return { ok: false, gone, error: e.message || String(e) };
  }
}

/**
 * 對指定 user 送推播（會並行送所有他的訂閱），並自動把 410/404 的訂閱從 auth.json 拿掉。
 * 回傳：{ sent: N, removed: M, total: T }
 *
 * 參數：
 *   loadAuth, saveAuth：auth.json 讀寫函式（從 server.js 注入，避免循環依賴）
 *   username：對哪個帳號送
 *   payload：物件，會 JSON.stringify 後送出，前端 SW 收到後解析
 *     {
 *       title: string,
 *       body:  string,
 *       icon?: string  (預設 /icon-192.png)
 *       badge?: string (預設 /icon-192.png)
 *       tag?:  string  (相同 tag 的通知會合併)
 *       url?:  string  (點擊通知後要開的 URL；預設 /)
 *       data?: object  (任意附加資料)
 *     }
 */
async function sendToUser(loadAuth, saveAuth, username, payload) {
  if (!enabled) return { sent: 0, removed: 0, total: 0, skipped: 'disabled' };

  const auth = loadAuth();
  const user = auth.users.find(u => u.username === username);
  if (!user || !Array.isArray(user.pushSubscriptions) || user.pushSubscriptions.length === 0) {
    return { sent: 0, removed: 0, total: 0 };
  }

  const subs = user.pushSubscriptions.slice();
  const results = await Promise.all(subs.map(sub => sendToSubscription(sub, payload)));

  let sent = 0;
  const survivors = [];
  let removed = 0;

  for (let i = 0; i < subs.length; i++) {
    if (results[i].ok) {
      sent++;
      survivors.push(subs[i]);
    } else if (results[i].gone) {
      removed++;
    } else {
      // 暫時性錯誤：保留訂閱下次重試
      survivors.push(subs[i]);
    }
  }

  if (removed > 0) {
    user.pushSubscriptions = survivors;
    saveAuth(auth);
  }

  return { sent, removed, total: subs.length };
}

module.exports = {
  isEnabled,
  validateEndpoint,
  makeSafeLookup,
  makeSafeAgent,
  getPublicKey,
  sendToSubscription,
  sendToUser,
};
