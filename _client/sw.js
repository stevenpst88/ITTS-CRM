/* ITTS CRM Service Worker
 *
 * 策略：
 *   - /api/*                → 永遠 network-only（live data + auth，絕不 cache）
 *   - /uploads/*            → network-only（簽署 URL / auth-gated）
 *   - HTML（含 /, /admin.html, /login.html, /help, /help.html）
 *                          → network-first，失敗時回 cache，最後 fallback offline
 *   - /sw.js, /manifest.webmanifest
 *                          → network-only（避免 cache 自身造成升級問題）
 *   - 其他靜態（CSS / JS / 圖片 / 字型）
 *                          → stale-while-revalidate
 *
 * 升級流程：
 *   - 改 SW_VERSION 觸發新版本
 *   - install 立即 skipWaiting()，但不主動 controllerchange，靠頁面收到 message 後自行 reload
 *   - 客戶端可 postMessage({type:'SKIP_WAITING'}) 強制接管
 */

const SW_VERSION = 'itts-crm-v2';
const STATIC_CACHE = `${SW_VERSION}-static`;
const HTML_CACHE = `${SW_VERSION}-html`;

const PRECACHE_URLS = [
  '/icon.svg',
  '/icon-192.png',
  '/icon-512.png',
  '/itts-logo.png',
  '/itts-logo.svg',
  '/manifest.webmanifest',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(STATIC_CACHE).then((cache) =>
      cache.addAll(PRECACHE_URLS).catch(() => {}),
    ),
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(
        keys
          .filter((k) => !k.startsWith(SW_VERSION))
          .map((k) => caches.delete(k)),
      );
      await self.clients.claim();
    })(),
  );
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

function isNetworkOnly(url) {
  return (
    url.pathname.startsWith('/api/') ||
    url.pathname.startsWith('/uploads/') ||
    url.pathname === '/sw.js' ||
    url.pathname === '/manifest.webmanifest'
  );
}

function isHtmlRequest(request, url) {
  if (request.mode === 'navigate') return true;
  if (request.destination === 'document') return true;
  const accept = request.headers.get('accept') || '';
  if (accept.includes('text/html')) return true;
  return /\.html$/.test(url.pathname);
}

self.addEventListener('fetch', (event) => {
  const { request } = event;

  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  if (url.origin !== self.location.origin) return;

  if (isNetworkOnly(url)) return;

  if (isHtmlRequest(request, url)) {
    event.respondWith(networkFirstHtml(request));
    return;
  }

  event.respondWith(staleWhileRevalidate(request));
});

async function networkFirstHtml(request) {
  const cache = await caches.open(HTML_CACHE);
  try {
    const fresh = await fetch(request);
    if (fresh && fresh.ok) cache.put(request, fresh.clone());
    return fresh;
  } catch (err) {
    const cached = await cache.match(request);
    if (cached) return cached;
    return new Response(
      '<!doctype html><meta charset="utf-8"><title>離線</title>' +
        '<body style="font-family:-apple-system,sans-serif;text-align:center;padding:60px 20px;color:#1a3c7a">' +
        '<h2>📡 網路連線中斷</h2><p>請確認連線後重新整理。</p></body>',
      { status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' } },
    );
  }
}

async function staleWhileRevalidate(request) {
  const cache = await caches.open(STATIC_CACHE);
  const cached = await cache.match(request);
  const fetchPromise = fetch(request)
    .then((res) => {
      if (res && res.ok) cache.put(request, res.clone());
      return res;
    })
    .catch(() => cached);
  return cached || fetchPromise;
}

// ── Web Push 推播 ─────────────────────────────────────────
self.addEventListener('push', (event) => {
  let payload = {};
  try {
    payload = event.data ? event.data.json() : {};
  } catch (_) {
    payload = { title: 'ITTS CRM', body: event.data ? event.data.text() : '' };
  }

  const title = payload.title || 'ITTS CRM';
  const options = {
    body: payload.body || '',
    icon: payload.icon || '/icon-192.png',
    badge: payload.badge || '/icon-192.png',
    tag: payload.tag || 'itts-crm',
    renotify: !!payload.renotify,
    requireInteraction: !!payload.requireInteraction,
    data: payload.data || {},
  };

  event.waitUntil(self.registration.showNotification(title, options));
});

// 點擊通知 → 帶到對應頁面（已開的 tab 優先 focus，沒有就開新的）
self.addEventListener('notificationclick', (event) => {
  event.notification.close();

  const targetUrl = (event.notification.data && event.notification.data.url) || '/';

  event.waitUntil((async () => {
    const all = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    // 先找已開的同來源 tab
    for (const client of all) {
      const u = new URL(client.url);
      if (u.origin === self.location.origin) {
        await client.focus();
        // 用 message 通知前端內部 navigate（避免整頁 reload 失去未存資料）
        try {
          client.postMessage({ type: 'NAVIGATE', url: targetUrl });
        } catch (_) {}
        return;
      }
    }
    // 沒有 tab → 開新的
    if (self.clients.openWindow) {
      await self.clients.openWindow(targetUrl);
    }
  })());
});

// 訂閱失效（瀏覽器主動撤銷）→ 嘗試重新訂閱（最佳努力，失敗就算了，下次使用者開頁面會重新跑訂閱流程）
self.addEventListener('pushsubscriptionchange', (event) => {
  event.waitUntil((async () => {
    try {
      const r = await fetch('/api/push/vapid-public-key');
      const { key } = await r.json();
      if (!key) return;
      const sub = await self.registration.pushManager.subscribe({
        userVisibleOnly: true,
        applicationServerKey: urlBase64ToUint8Array(key),
      });
      await fetch('/api/push/subscribe', {
        method: 'POST',
        credentials: 'include',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(sub),
      });
    } catch (_) {}
  })());
});

function urlBase64ToUint8Array(base64String) {
  const padding = '='.repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
  const raw = atob(base64);
  const out = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i);
  return out;
}
