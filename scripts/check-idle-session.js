// 閒置自動登出：jwtSession 的單元檢查（不啟動伺服器）。
// 用法：SESSION_SECRET=... node scripts/check-idle-session.js   （未設定時用測試用假值）
process.env.SESSION_SECRET = process.env.SESSION_SECRET || 'unit-test-secret-not-for-production-use-0123456789';
process.env.SESSION_IDLE_MS = '600000';   // 10 分鐘：用來驗證「可覆寫」，預設值另外在最後一組測
const jwt = require('jsonwebtoken');
const jwtSession = require('../middleware/jwtSession');

let pass = 0, fail = 0;
function ok(cond, name, extra) {
  if (cond) pass++; else { fail++; console.log('FAIL ' + name + (extra ? '  ← ' + extra : '')); }
}
const SECRET = process.env.SESSION_SECRET;
const nowS = () => Math.floor(Date.now() / 1000);
const HOUR = 3600;

function run(cookieToken, path, method) {
  const req = { headers: { cookie: cookieToken ? 'itts_auth=' + encodeURIComponent(cookieToken) : '' }, path: path || '/api/contacts', method: method || 'GET' };
  const cookies = [];
  const cleared = [];
  const res = {
    cookie(name, val, opt) { cookies.push({ name, val, opt }); },
    clearCookie(name) { cleared.push(name); },
  };
  let nexted = false;
  jwtSession(req, res, () => { nexted = true; });
  return { req, res, cookies, cleared, nexted };
}
const mk = (payload) => jwt.sign(payload, SECRET);
const U = { username: 'u1', role: 'user' };

// ── 1. 剛簽發的 token：有登入、活動時間沒過期 ──────────────
{
  const t = mk({ user: U, lgn: nowS() - 60, lat: nowS() - 5, exp: nowS() + 7 * HOUR });
  const r = run(t);
  ok(r.nexted, '1a 有呼叫 next');
  ok(r.req.session.user && r.req.session.user.username === 'u1', '1b 登入狀態有效');
  ok(!r.req._sessionIdleExpired, '1c 沒有閒置標記');
  ok(r.cookies.length === 0, '1d 5 秒前才活動過 → 不重簽（省寫入）', JSON.stringify(r.cookies.length));
}

// ── 2. 使用者操作 → 滑動延長（重簽，lat 更新，lgn 與 exp 不變）─────
{
  const lgn = nowS() - 2 * HOUR;
  const t = mk({ user: U, lgn, lat: nowS() - 120, exp: lgn + 8 * HOUR });
  const r = run(t, '/api/contacts');
  ok(r.cookies.length === 1, '2a 活動超過重簽門檻 → 重簽一次', String(r.cookies.length));
  if (r.cookies.length) {
    const d = jwt.verify(decodeURIComponent(r.cookies[0].val), SECRET);
    ok(d.lgn === lgn, '2b 登入時間沿用');
    ok(d.exp === lgn + 8 * HOUR, '2c 絕對上限 exp = lgn+8h 不變');
    ok(Math.abs(d.lat - nowS()) <= 2, '2d lat 更新為現在');
    ok(r.cookies[0].opt.maxAge <= 600000 + jwtSession.COOKIE_GRACE_MS && r.cookies[0].opt.maxAge >= 600000, '2e cookie 存活 = 閒置視窗＋寬限', String(r.cookies[0].opt.maxAge));
    ok(r.cookies[0].opt.httpOnly === true && r.cookies[0].opt.sameSite === 'strict', '2f cookie 屬性 httpOnly／strict 不變');
  }
}

// ── 3. 背景輪詢不算活動：不續期 ─────────────────────────
{
  const lgn = nowS() - HOUR;
  const t = mk({ user: U, lgn, lat: nowS() - 300, exp: lgn + 8 * HOUR });
  const r = run(t, '/api/poll-bundle');
  ok(r.req.session.user && r.req.session.user.username === 'u1', '3a 輪詢時登入仍有效');
  ok(r.cookies.length === 0, '3b poll-bundle 不重簽（不續期）');
}

// ── 3x. 路徑正規化：結尾斜線、大小寫都還是背景輪詢 ─────────
{
  const lgn = nowS() - HOUR;
  const t = mk({ user: U, lgn, lat: nowS() - 300, exp: lgn + 8 * HOUR });
  for (const p of ['/api/poll-bundle/', '/API/Poll-Bundle', '/api/poll-bundle//']) {
    ok(run(t, p).cookies.length === 0, '3x 輪詢路徑變體不算活動：' + p);
  }
  ok(run(t, '/api/poll-bundle-other').cookies.length === 1, '3y 只是前綴相同的其他路徑仍算活動');
}

// ── 4. 閒置超過視窗 → 視為未登入、清 cookie、標記 idle ─────────
{
  const lgn = nowS() - HOUR;
  const t = mk({ user: U, lgn, lat: nowS() - 601, exp: lgn + 8 * HOUR });
  const r = run(t, '/api/contacts');
  ok(!r.req.session.user, '4a 閒置超時 → 未登入');
  ok(r.req._sessionIdleExpired === true, '4b 標記 _sessionIdleExpired');
  ok(r.cleared.includes('itts_auth'), '4c 清掉 cookie');
  // 輪詢也一樣會被判逾時
  const r2 = run(t, '/api/poll-bundle');
  ok(!r2.req.session.user && r2.req._sessionIdleExpired === true, '4d 輪詢請求遇到逾時一樣登出（才看得到逾時）');
  // 剛好在邊界內
  const t2 = mk({ user: U, lgn, lat: nowS() - 590, exp: lgn + 8 * HOUR });
  ok(run(t2, '/api/poll-bundle').req.session.user, '4e 590 秒前活動（視窗 600 秒）仍有效');
}

// ── 5. 舊版 token（沒有 lat／lgn）相容 ─────────────────
{
  const recent = jwt.sign({ user: U }, SECRET, { expiresIn: '8h' });   // iat = 現在
  ok(run(recent).req.session.user, '5a 舊版 token 剛簽發 → 有效');
  const old = jwt.sign({ user: U, iat: nowS() - 3600 }, SECRET, { expiresIn: '8h' });   // 1 小時前簽的
  const r = run(old);
  ok(!r.req.session.user && r.req._sessionIdleExpired === true, '5b 舊版 token 簽發超過閒置視窗 → 逾時（部署後一次性重新登入）');
}

// ── 6. 8 小時絕對上限 ──────────────────────────────────
{
  const t = mk({ user: U, lgn: nowS() - 9 * HOUR, lat: nowS() - 5, exp: nowS() - HOUR });
  const r = run(t);
  ok(!r.req.session.user, '6a 超過 8 小時上限（即使剛有活動）→ 未登入');
  ok(!r.req._sessionIdleExpired, '6b 這種情況不是「閒置」（不誤導成閒置訊息）');
  // 接近上限：cookie 存活不得超過剩餘時間
  const lgn = nowS() - 8 * HOUR + 120;   // 還剩 120 秒
  const t2 = mk({ user: U, lgn, lat: nowS() - 300, exp: lgn + 8 * HOUR });
  const r2 = run(t2);
  ok(r2.cookies.length === 1 && r2.cookies[0].opt.maxAge <= 125000, '6c 快到上限時 cookie 存活 ≤ 剩餘時間', r2.cookies[0] && String(r2.cookies[0].opt.maxAge));
}

// ── 7. 寫入 req.session.user 與 destroy ────────────────
{
  const lgn = nowS() - 3 * HOUR;
  const t = mk({ user: U, lgn, lat: nowS() - 10, exp: lgn + 8 * HOUR });
  // 同一人改資料重簽：lgn 不能被重置
  const r = run(t);
  r.req.session.user = { username: 'u1', role: 'user', displayName: 'x' };
  const c = r.cookies[r.cookies.length - 1];
  const d = jwt.verify(decodeURIComponent(c.val), SECRET);
  ok(d.lgn === lgn, '7a 同一人改個人資料重簽 → 登入時間沿用（上限不被重置）');
  // 換成別人（重新登入）：lgn 重置
  const r2 = run(t);
  r2.req.session.user = { username: 'someone-else', role: 'user' };
  const d2 = jwt.verify(decodeURIComponent(r2.cookies[r2.cookies.length - 1].val), SECRET);
  ok(Math.abs(d2.lgn - nowS()) <= 2, '7b 換成別的帳號登入 → 登入時間重置為現在');
  // 全新登入（沒有 cookie）
  const r3 = run('');
  r3.req.session.user = { username: 'new', role: 'user' };
  const d3 = jwt.verify(decodeURIComponent(r3.cookies[0].val), SECRET);
  ok(Math.abs(d3.lgn - nowS()) <= 2 && Math.abs(d3.lat - nowS()) <= 2, '7c 全新登入 lgn／lat 都是現在');
  ok(d3.exp - d3.lgn === 8 * HOUR, '7d 新登入的上限仍是 8 小時');
  // 閒置逾時後在同一請求重新登入 → 以現在為起點，不沿用舊 lgn
  const old = mk({ user: U, lgn: nowS() - 5 * HOUR, lat: nowS() - 4000, exp: nowS() + 3 * HOUR });
  const r4 = run(old);
  r4.req.session.user = { username: 'u1', role: 'user' };
  const d4 = jwt.verify(decodeURIComponent(r4.cookies[r4.cookies.length - 1].val), SECRET);
  ok(Math.abs(d4.lgn - nowS()) <= 2, '7e 逾時後重新登入 → 登入時間重置（不繼承舊的 8 小時上限）');
  // destroy
  const r5 = run(t);
  r5.req.session.destroy(() => {});
  ok(!r5.req.session.user && r5.cleared.includes('itts_auth'), '7f destroy 清除登入與 cookie');
}

// ── 8. 無效／竄改 token ───────────────────────────────
{
  const bad = jwt.sign({ user: U, lgn: nowS(), lat: nowS() }, 'another-secret-another-secret-another-0123');
  ok(!run(bad).req.session.user, '8a 簽章不符 → 未登入');
  ok(!run('not.a.jwt').req.session.user, '8b 亂碼 → 未登入');
  ok(!run('').req.session.user, '8c 沒有 cookie → 未登入');
  ok(run('').nexted, '8d 未登入也要呼叫 next（交給 requireAuth 處理）');
}

// ── 9. 設定值 ────────────────────────────────────────
ok(jwtSession.IDLE_MS === 600000, '9a SESSION_IDLE_MS 可覆寫（測試用）', String(jwtSession.IDLE_MS));
ok(jwtSession.PASSIVE_PATHS.has('/api/poll-bundle'), '9b poll-bundle 在背景輪詢清單');
{
  const { spawnSync } = require('child_process');
  const probe = (val) => spawnSync(process.execPath, ['-e', "process.env.SESSION_SECRET='unit-test-secret-not-for-production-use-0123456789';" + (val === undefined ? '' : "process.env.SESSION_IDLE_MS='" + val + "';") + "console.log(require('./middleware/jwtSession').IDLE_MS)"], { cwd: require('path').join(__dirname, '..'), encoding: 'utf8', env: Object.assign({}, process.env, { SESSION_IDLE_MS: val === undefined ? '' : val }) }).stdout.trim();
  ok(probe(undefined) === String(30 * 60 * 1000), '9c 預設閒置視窗 = 30 分鐘', probe(undefined));
  ok(probe('0') === String(30 * 60 * 1000), '9d 設成 0 不會把全站鎖死（退回預設）', probe('0'));
  ok(probe('100') === String(30 * 60 * 1000), '9e 小於 5 秒的值退回預設', probe('100'));
  ok(probe('abc') === String(30 * 60 * 1000), '9f 非數字退回預設', probe('abc'));
  ok(probe('8000') === '8000', '9g 合理的測試值生效', probe('8000'));
}

// ── 10. M-01：單一機制——輪詢清單共用、SESSION_IDLE_MINUTES 由 jwtSession 讀取 ─────────────
{
  const lgn = nowS() - HOUR;
  const t = mk({ user: U, lgn, lat: nowS() - 300, exp: lgn + 8 * HOUR });
  for (const p of ['/api/notifications', '/api/announcements/active', '/API/Notifications/']) {
    ok(run(t, p, 'GET').cookies.length === 0, '10a 背景輪詢（GET）不續命：' + p);
    ok(run(t, p, 'HEAD').cookies.length === 0, '10b 背景輪詢（HEAD）不續命：' + p);
  }
  ok(run(t, '/api/notifications', 'POST').cookies.length === 1, '10c 同一路徑的 POST 是使用者操作 → 續命');
  ok(run(t, '/api/notifications/read-all', 'POST').cookies.length === 1, '10d 其他 notifications 子路徑算活動');
  ok(['/api/poll-bundle', '/api/notifications', '/api/announcements/active'].every((x) => jwtSession.PASSIVE_PATHS.has(x)), '10e PASSIVE_PATHS 與 loginGuard 共用同一份清單');
  const LG = require('../lib/loginGuard');
  ok(JSON.stringify(LG.IDLE_POLL_PATHS) === JSON.stringify(Array.from(jwtSession.PASSIVE_PATHS)), '10f loginGuard.IDLE_POLL_PATHS 與 jwtSession.PASSIVE_PATHS 相同');
  // 逾時時帶出帳號，server.js 用它寫 SESSION_IDLE_TIMEOUT 稽核
  const tOld = mk({ user: U, lgn, lat: nowS() - 601, exp: lgn + 8 * HOUR });
  ok(run(tOld).req._sessionIdleUser === 'u1', '10g 逾時時 req._sessionIdleUser = 帳號');
  // 一個請求最多一個 Set-Cookie（活動重簽只簽一次）
  const rr = run(t, '/api/contacts', 'GET');
  ok(rr.cookies.length <= 1, '10h 一個請求最多簽發一次 cookie', String(rr.cookies.length));
  const { spawnSync } = require('child_process');
  const probe2 = (env) => spawnSync(process.execPath, ['-e', "require('./middleware/jwtSession');console.log(require('./middleware/jwtSession').IDLE_MS)"], { cwd: require('path').join(__dirname, '..'), encoding: 'utf8', env: Object.assign({}, process.env, { SESSION_SECRET: 'unit-test-secret-not-for-production-use-0123456789', SESSION_IDLE_MS: '', SESSION_IDLE_MINUTES: '' }, env) }).stdout.trim();
  ok(probe2({ SESSION_IDLE_MINUTES: '60' }) === String(3600000), '10i SESSION_IDLE_MINUTES=60 生效（jwtSession 的 IDLE_MS = 60 分鐘）', probe2({ SESSION_IDLE_MINUTES: '60' }));
  ok(probe2({ SESSION_IDLE_MINUTES: '60', SESSION_IDLE_MS: '8000' }) === '8000', '10j SESSION_IDLE_MS（測試）優先於 SESSION_IDLE_MINUTES');
  ok(probe2({ SESSION_IDLE_MINUTES: '0' }) === String(1800000) && probe2({ SESSION_IDLE_MINUTES: 'abc' }) === String(1800000) && probe2({ SESSION_IDLE_MINUTES: '-5' }) === String(1800000), '10k SESSION_IDLE_MINUTES 為 0／非數字／負數 → 預設 30 分鐘');
  ok(probe2({ SESSION_IDLE_MINUTES: '0.1' }) === '6000', '10l 小數分鐘（0.1＝6 秒）可用於測試');
}

console.log('閒置登出 jwtSession 單元檢查：PASS ' + pass + ' / FAIL ' + fail);
process.exit(fail ? 1 : 0);
