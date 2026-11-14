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

const PUBLIC_KEY  = process.env.VAPID_PUBLIC_KEY;
const PRIVATE_KEY = process.env.VAPID_PRIVATE_KEY;
const SUBJECT     = process.env.VAPID_SUBJECT || 'mailto:admin@itts.com.tw';

const enabled = !!(PUBLIC_KEY && PRIVATE_KEY);

if (enabled) {
  webpush.setVapidDetails(SUBJECT, PUBLIC_KEY, PRIVATE_KEY);
} else {
  console.warn('[pushNotify] 未設定 VAPID_PUBLIC_KEY / VAPID_PRIVATE_KEY，Web Push 已停用');
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
async function sendToSubscription(subscription, payload) {
  if (!enabled) return { ok: false, gone: false, error: 'web push disabled' };
  try {
    const body = typeof payload === 'string' ? payload : JSON.stringify(payload);
    await webpush.sendNotification(subscription, body, { TTL: 60 * 60 * 24 });
    return { ok: true };
  } catch (e) {
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
 *   opts（可省略）：{ applyOnLatestAuth(mutate) }——清除失效訂閱時不要拿「等待網路之前」取得的 auth 整份存回（那會把這段時間別人改的帳號資料還原），
 *     改由呼叫端在鎖內重讀最新 auth 後執行 mutate(最新 auth)（回傳 false＝沒有變更、不用存）。不給＝舊行為（直接改手上的 auth 並 saveAuth）。
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
async function sendToUser(loadAuth, saveAuth, username, payload, opts) {
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
    if (opts && typeof opts.applyOnLatestAuth === 'function') {
      const goneEndpoints = new Set(subs.filter((_, i) => results[i].gone).map((x) => x.endpoint));
      await opts.applyOnLatestAuth((latestAuth) => {
        const u = ((latestAuth && latestAuth.users) || []).find((x) => x.username === username);
        if (!u || !Array.isArray(u.pushSubscriptions)) return false;
        const before = u.pushSubscriptions.length;
        u.pushSubscriptions = u.pushSubscriptions.filter((x) => !goneEndpoints.has(x.endpoint));
        return u.pushSubscriptions.length !== before;
      });
    } else {
      user.pushSubscriptions = survivors;
      saveAuth(auth);
    }
  }

  return { sent, removed, total: subs.length };
}

module.exports = {
  isEnabled,
  getPublicKey,
  sendToSubscription,
  sendToUser,
};
