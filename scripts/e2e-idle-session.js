// 閒置自動登出：真實伺服器端對端檢查（隔離沙盒）。
// 用法（在乾淨的工作資料夾，會建立並結束後刪除 data.json／auth.json 測試複本）：
//   node scripts/e2e-idle-session.js
// 會用 3961 埠、SESSION_IDLE_MS=8000（8 秒）啟動一個全新資料的伺服器，臨時帳號 _s_idle。不碰真實資料檔。
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const bcrypt = require('bcryptjs');

const ROOT = path.join(__dirname, '..');
const PORT = 3961;
const BASE = 'http://localhost:' + PORT;
const IDLE = 8000;
const authFile = path.join(ROOT, 'auth.json');
const dataFile = path.join(ROOT, 'data.json');
if (fs.existsSync(authFile) || fs.existsSync(dataFile)) {
  console.error('這個資料夾已有 auth.json／data.json，為避免覆蓋真實資料，請在乾淨的 worktree 執行。');
  process.exit(2);
}
const PW = 'T3st-' + Math.random().toString(36).slice(2, 10) + '!A1';   // 臨時密碼：隨機產生，只活在這個程序

let pass = 0, fail = 0;
const ok = (c, n, x) => { if (c) pass++; else { fail++; console.log('FAIL ' + n + (x ? '  ← ' + x : '')); } };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function cookieOf(res, jar) {
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  for (const c of sc) {
    const m = /^itts_auth=([^;]*)/.exec(c);
    if (m) { jar.v = m[1]; jar.raw = c; if (!m[1]) jar.v = ''; }
  }
}
async function call(jar, method, p, body) {
  const res = await fetch(BASE + p, {
    method,
    headers: Object.assign({ Origin: BASE }, jar.v ? { Cookie: 'itts_auth=' + jar.v } : {}, body ? { 'Content-Type': 'application/json' } : {}),
    body: body ? JSON.stringify(body) : undefined,
    redirect: 'manual',
  });
  cookieOf(res, jar);
  res._nSet = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).filter((c) => /^itts_auth=/.test(c)).length;   // M-01：每請求的 itts_auth Set-Cookie 數
  return res;
}
async function login(jar) {
  const r = await call(jar, 'POST', '/api/login', { username: '_s_idle', password: PW });
  return r.status;
}

(async () => {
  const hash = await bcrypt.hash(PW, 4);
  fs.writeFileSync(authFile, JSON.stringify({ users: [{
    username: '_s_idle', password: hash, displayName: '閒置測試', role: 'user', active: true,
    passwordChangedAt: new Date().toISOString(),
  }] }));
  const srv = spawn(process.execPath, ['server.js'], {
    cwd: ROOT,
    env: Object.assign({}, process.env, { SESSION_SECRET: 'e2e-idle-secret-not-real-0123456789abcdef', DB_BACKEND: 'json', PORT: String(PORT), SESSION_IDLE_MS: String(IDLE), MAIL_MODE: 'off', ALLOWED_ORIGINS: BASE }),
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  let log = '';
  srv.stdout.on('data', (d) => { log += d; });
  srv.stderr.on('data', (d) => { log += d; });
  const cleanup = () => { try { srv.kill(); } catch (e) { /* */ } try { fs.unlinkSync(authFile); } catch (e) { /* */ } try { fs.unlinkSync(dataFile); } catch (e) { /* */ } };
  try {
    for (let i = 0; i < 60; i++) { try { const r = await fetch(BASE + '/login.html'); if (r.ok) break; } catch (e) { /* 等啟動 */ } await sleep(250); }

    // 1. 沒登入
    let jar = {};
    let r = await call(jar, 'GET', '/api/me');
    ok(r.status === 401 && r.headers.get('X-Session-Expired') === '1', '1a 未登入 401 帶 X-Session-Expired: 1', r.status + ' ' + r.headers.get('X-Session-Expired'));
    r = await call(jar, 'POST', '/api/login', { username: '_s_idle', password: 'wrong-' + PW });
    ok(r.status === 401 && !r.headers.get('X-Session-Expired'), '1b 密碼錯誤的 401 不帶 X-Session-Expired（不會被當成登入過期）', r.status + ' ' + r.headers.get('X-Session-Expired'));

    // 2. 登入、cookie 存活不超過閒置視窗、touch
    ok(await login(jar) === 200, '2a 登入成功');
    ok(/Max-Age=(\d+)/.test(jar.raw) && Number(/Max-Age=(\d+)/.exec(jar.raw)[1]) <= IDLE / 1000 + 300 && Number(/Max-Age=(\d+)/.exec(jar.raw)[1]) > IDLE / 1000, '2b cookie Max-Age = 閒置視窗＋5 分鐘寬限', jar.raw);
    r = await call(jar, 'GET', '/api/me');
    ok(r.status === 200, '2c 登入後 /api/me 200');
    r = await call(jar, 'GET', '/api/session/touch');
    const tj = await r.json().catch(() => ({}));
    ok(r.status === 200 && tj.ok === true && tj.idleMs === IDLE, '2d /api/session/touch 200 且回報閒置視窗', JSON.stringify(tj));

    // 3. 只有背景輪詢：不續命，閒置視窗到了就登出（且看得出是 idle）
    const t0 = Date.now();
    let expiredAt = 0;
    let sawHeader = '';
    while (Date.now() - t0 < IDLE + 5000) {
      await sleep(2000);
      r = await call(jar, 'GET', '/api/poll-bundle?days=3');
      if (r.status === 401) { expiredAt = Date.now() - t0; sawHeader = r.headers.get('X-Session-Expired'); break; }
    }
    ok(expiredAt > 0 && expiredAt >= IDLE - 1500 && expiredAt <= IDLE + 4500, '3a 只輪詢 → 約在閒置視窗到時登出', 'expiredAt=' + expiredAt);
    ok(sawHeader === 'idle', '3b 逾時的 401 標頭是 idle', sawHeader);
    const body3 = await r.json().catch(() => ({}));
    ok(/閒置/.test(body3.error || ''), '3c 逾時回應訊息提到閒置', JSON.stringify(body3));
    // 逾時後 cookie 已被清掉；再打任何 API 是一般 401（不是 idle）
    r = await call(jar, 'GET', '/api/me');
    ok(r.status === 401 && r.headers.get('X-Session-Expired') === '1', '3d 逾時後 cookie 已清除，之後是一般未登入 401', r.status + ' ' + r.headers.get('X-Session-Expired'));

    // 4. 有操作：滑動延長，超過閒置視窗總時間仍有效
    jar = {};
    ok(await login(jar) === 200, '4a 重新登入');
    const t1 = Date.now();
    let alive = true;
    while (Date.now() - t1 < IDLE * 2 + 2000) {
      await sleep(3000);
      r = await call(jar, 'GET', '/api/session/touch');
      if (r.status !== 200) { alive = false; break; }
    }
    ok(alive, '4b 每 3 秒有操作 → 經過 2 倍閒置視窗仍有效（滑動延長）');
    r = await call(jar, 'GET', '/api/me');
    ok(r.status === 200, '4c 仍可使用');

    // 5. 操作停止 → 逾時；輪詢夾雜其中不續命
    await sleep(IDLE + 1500);
    r = await call(jar, 'GET', '/api/me');
    ok(r.status === 401 && r.headers.get('X-Session-Expired') === 'idle', '5a 停止操作超過視窗 → 下一個請求是 idle 逾時', r.status + ' ' + r.headers.get('X-Session-Expired'));

    // 5b. 登入狀態下同一帳號再登入一次：登入時間要重置（拿到全新的 8 小時），不能沿用舊的
    jar = {};
    ok(await login(jar) === 200, '5b-1 登入');
    const first = JSON.parse(Buffer.from(decodeURIComponent(jar.v).split('.')[1], 'base64url').toString());
    await sleep(3000);
    ok(await login(jar) === 200, '5b-2 帶著有效 cookie 再登入一次（同帳號）');
    const second = JSON.parse(Buffer.from(decodeURIComponent(jar.v).split('.')[1], 'base64url').toString());
    ok(second.lgn >= first.lgn + 2 && second.exp >= first.exp + 2, '5b-3 重新登入 → 登入時間與 8 小時上限都重置', JSON.stringify({ first: [first.lgn, first.exp], second: [second.lgn, second.exp] }));

    // 5c. M-01：單一機制——/api/notifications 與 /api/announcements/active 的 GET 不續命；每個請求最多一個 itts_auth Set-Cookie
    jar = {};
    ok(await login(jar) === 200, '5c-1 登入');
    let maxSet = 0;
    for (let i = 0; i < 6; i++) { await sleep(700); r = await call(jar, 'GET', '/api/me'); maxSet = Math.max(maxSet, r._nSet); }
    ok(maxSet <= 1, '5c-2 活動中每個請求最多 1 個 itts_auth Set-Cookie（不重複簽發）', String(maxSet));
    const t2 = Date.now(); let exp2 = 0, hdr2 = '';
    while (Date.now() - t2 < IDLE + 5000) {
      await sleep(2000);
      r = await call(jar, 'GET', Math.random() < 0.5 ? '/api/notifications' : '/api/announcements/active');
      if (r.status === 401) { exp2 = Date.now() - t2; hdr2 = r.headers.get('X-Session-Expired'); break; }
    }
    ok(exp2 > 0 && exp2 <= IDLE + 4500, '5c-3 只打 notifications／announcements/active（背景輪詢）不續命 → 逾時登出', 'expiredAt=' + exp2);
    ok(hdr2 === 'idle', '5c-4 該 401 帶 X-Session-Expired: idle', hdr2);

    // 6. 前端檔案可取得、頁面有載入
    for (const f of ['/idle-timeout.js', '/login.html?reason=idle']) {
      r = await fetch(BASE + f);
      ok(r.status === 200, '6 取得 ' + f, String(r.status));
    }
    const idx = await (await fetch(BASE + '/login.html')).text();
    const lcSrc = await (await fetch(BASE + '/login-cinematic.js')).text();   // 登入頁無 inline script：reason 由 login-cinematic.js 處理
    ok(/reason/.test(idx + lcSrc), '6x 登入頁有處理 reason 參數');
    // idle-timeout.js 要等登入才看得到 index.html；以靜態檔確認被 index／admin 引用
    const idxSrc = fs.readFileSync(path.join(ROOT, '_client', 'index.html'), 'utf8');
    const admSrc = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
    ok(/src="idle-timeout\.js"/.test(idxSrc) && /src="idle-timeout\.js"/.test(admSrc), '6y index.html 與 admin.html 都有載入 idle-timeout.js');
  } catch (e) {
    fail++; console.log('例外：' + (e && e.stack || e));
  } finally {
    cleanup();
  }
  console.log('閒置登出端對端檢查：PASS ' + pass + ' / FAIL ' + fail);
  if (fail) console.log('--- 伺服器輸出末段 ---\n' + log.split('\n').slice(-15).join('\n'));
  process.exit(fail ? 1 : 0);
})();
