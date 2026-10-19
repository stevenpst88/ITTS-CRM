#!/usr/bin/env node
'use strict';
/**
 * 帳號鎖定／閒置逾時（lib/loginGuard.js）檢查：
 *   Part A 單元測試（注入時鐘、記憶體資料庫，不開伺服器）＋ 變異檢查（把規則改壞，測試必須抓得到）
 *   Part B 沙盒伺服器端對端（預設埠 3211；JWT 與 express-session 兩種後端；開關關閉＝與先前相同）
 *
 *   node scripts/check-login-guard.js            # 全部
 *   LG_SKIP_E2E=1 node scripts/check-login-guard.js   # 只跑單元＋變異
 * 守門與埠限制見 scripts/_b2-harness.js。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const H = require('./_b2-harness');
const { ROOT, ok, eq, sleep } = H;
const LG_PATH = path.join(ROOT, 'lib', 'loginGuard.js');

// ═══════════════════════ Part A：單元測試 ═══════════════════════
function mkClock(t0) { const c = { t: t0 || 1_700_000_000_000, now() { return c.t; }, adv(ms) { c.t += ms; } }; return c; }
// 仿 JSON 後端：load() 每次回新的副本，必須 save(同一份) 才會生效（抓得到「改了副本卻存到別份」這類錯誤）
function mkStore() { const s = { data: {}, saves: 0, load: () => JSON.parse(JSON.stringify(s.data)), save: (d) => { s.saves++; s.data = JSON.parse(JSON.stringify(d)); } }; return s; }
const MIN = 60000;

function unitSuite(mod, T) {
  // T：{ ok, eq }，變異檢查時換成「記錄失敗次數」版本
  const { ok, eq } = T;
  const env = { LOGIN_LOCKOUT: '1', LOGIN_MAX_FAILS: '5', LOGIN_LOCK_MINUTES: '15', LOGIN_LOCK_MAX_MINUTES: '60' };
  const mk = (e, c, s) => { const clock = c || mkClock(), store = s || mkStore(); return { g: mod.create({ env: e || env, clock: clock.now, store }), clock, store }; };

  // ── 預設關閉：完全不介入、不寫入 ──
  {
    const { g, store } = mk({});
    eq([g.lockoutEnabled, g.idleEnabled], [false, false], '預設兩個開關皆關');
    for (let i = 0; i < 20; i++) g.recordFailure('x');
    eq(g.checkLocked('x'), { locked: false }, '關閉時連續失敗不鎖');
    eq(store.saves, 0, '關閉時不寫入資料庫');
    ok(!('loginFailures' in store.data), '關閉時主資料 blob 沒有 loginFailures');
    eq(g.recordSuccess('x'), false, '關閉時 recordSuccess 無動作');
    eq(g.idleDecision({ loginAt: 1, lat: 1 }, 9e15, true), { expired: false, refresh: false }, '關閉時閒置判斷永不過期');
    let called = 0; const req = { path: '/api/x', method: 'GET', session: { user: { username: 'u', loginAt: 1 } } };
    g.idleMiddleware({})(req, {}, () => called++);
    eq([called, req.session.user.lat], [1, undefined], '關閉時 idle 中介層直接放行且不加 lat');
  }
  for (const v of ['0', '', 'false', 'off', undefined]) { const { g } = mk({ LOGIN_LOCKOUT: v, SESSION_IDLE_MINUTES: v }); ok(!g.lockoutEnabled && !g.idleEnabled, `LOGIN_LOCKOUT=${JSON.stringify(v)} 視為關閉`); }

  // ── 計數與鎖定 ──
  {
    const { g, clock, store } = mk();
    for (let i = 1; i <= 4; i++) { const r = g.recordFailure('Alice'); eq([r.locked, r.failCount], [false, i], `第 ${i} 次失敗尚未鎖定`); }
    eq(g.checkLocked('alice').locked, false, '4 次失敗未鎖');
    let r = g.recordFailure('ALICE ');
    eq([r.locked, r.justLocked, r.lockCount], [true, true, 1], '第 5 次失敗（大小寫／空白視為同名）鎖定');
    eq(r.until - clock.now(), 15 * MIN, '第一次鎖定 15 分鐘');
    eq(g.checkLocked('alice').locked, true, 'checkLocked 回報鎖定');
    // 鎖定中不累計、不延長
    const until0 = g.checkLocked('alice').until;
    clock.adv(5 * MIN); r = g.recordFailure('alice');
    eq([r.locked, r.justLocked, r.until], [true, false, until0], '鎖定中再失敗：不延長、不重複觸發');
    clock.adv(9 * MIN + 59000); eq(g.checkLocked('alice').locked, true, '到期前 1 秒仍鎖定');
    clock.adv(2000); eq(g.checkLocked('alice').locked, false, '到期後解除');
    // 第二次鎖定加倍
    for (let i = 0; i < 5; i++) r = g.recordFailure('alice');
    eq([r.justLocked, r.lockCount, r.until - clock.now()], [true, 2, 30 * MIN], '第二次鎖定加倍＝30 分鐘');
    clock.adv(31 * MIN);
    for (let i = 0; i < 5; i++) r = g.recordFailure('alice');
    eq(r.until - clock.now(), 60 * MIN, '第三次 60 分鐘（=上限）');
    clock.adv(61 * MIN);
    for (let i = 0; i < 5; i++) r = g.recordFailure('alice');
    eq([r.lockCount, r.until - clock.now()], [4, 60 * MIN], '第四次仍被上限 60 分鐘封頂');
    // 持久化內容：只有次數與時間戳
    const e = store.data.loginFailures['u:alice'];
    eq(Object.keys(e).sort(), ['first', 'last', 'lockCount', 'lockUntil', 'n'], '儲存欄位只有 n/first/last/lockUntil/lockCount');
    ok(Object.values(e).every(v => typeof v === 'number'), '儲存值全是數字（無密碼／IP／原因文字）');
    // 成功登入重置（需先解鎖才可能成功；這裡直接模擬鎖定到期後成功）
    clock.adv(61 * MIN);
    g.recordFailure('alice'); g.recordFailure('alice');
    eq(g.recordSuccess('alice'), true, '成功登入清掉紀錄');
    ok(!('u:alice' in store.data.loginFailures), '紀錄已刪除');
    for (let i = 0; i < 4; i++) g.recordFailure('alice');
    eq(g.checkLocked('alice').locked, false, '成功登入後計數歸零（再 4 次仍未鎖）');
    const saves = store.saves; eq(g.recordSuccess('nobody'), false, '沒有紀錄的成功登入回 false'); eq(store.saves, saves, '沒有紀錄的成功登入不寫入（一般登入零額外寫入）');
  }
  // 視窗過期：舊失敗不計
  {
    const { g, clock } = mk();
    for (let i = 0; i < 4; i++) g.recordFailure('bob');
    clock.adv(16 * MIN);
    for (let i = 0; i < 4; i++) g.recordFailure('bob');
    eq(g.checkLocked('bob').locked, false, '視窗(15 分鐘)外的失敗不累計：前 4 次過期、後 4 次不到門檻');
    g.recordFailure('bob');
    eq(g.checkLocked('bob').locked, true, '視窗內累積第 5 次才鎖');
  }
  // 倍率衰減：安靜 24 小時後重新從 15 分鐘起算
  {
    const { g, clock } = mk();
    let r; for (let i = 0; i < 5; i++) r = g.recordFailure('carl');
    clock.adv(16 * MIN); for (let i = 0; i < 5; i++) r = g.recordFailure('carl');
    eq(r.until - clock.now(), 30 * MIN, '衰減前第二次鎖為 30 分鐘');
    clock.adv(25 * 3600e3);
    for (let i = 0; i < 5; i++) r = g.recordFailure('carl');
    eq([r.lockCount, r.until - clock.now()], [1, 15 * MIN], '安靜 24 小時後倍率歸零，回到 15 分鐘');
  }
  // 不存在的帳號行為一致（單元層：同一份邏輯、沒有「存在與否」參數）；__proto__ 等名稱安全；大小
  {
    const { g, store } = mk();
    for (const n of ['__proto__', 'constructor', 'toString', '', '   ', 'a'.repeat(500)]) { for (let i = 0; i < 5; i++) g.recordFailure(n); }
    ok(Object.getPrototypeOf(store.data.loginFailures) === Object.prototype && ({}).polluted === undefined, '__proto__ 當帳號名不會污染原型');
    eq(g.checkLocked('__proto__').locked, true, '__proto__ 名稱可正常鎖定（鍵有前綴）');
    ok(Object.keys(store.data.loginFailures).every(k => k.startsWith('u:') && k.length <= 66), '鍵都有 u: 前綴且限長 64');
  }
  // 表滿時：先 prune 過期；全是鎖定中則新名稱不記（不崩潰）
  {
    const { g, clock, store } = mk();
    for (let i = 0; i < mod.MAX_ENTRIES; i++) g.recordFailure('n' + i);
    eq(Object.keys(store.data.loginFailures).length, mod.MAX_ENTRIES, '填滿 MAX_ENTRIES');
    let r = g.recordFailure('overflow-new');
    eq(r.dropped, true, '全部未過期且表滿 → 新名稱 dropped（不崩潰）');
    clock.adv(20 * MIN);
    r = g.recordFailure('overflow-new');
    ok(!r.dropped && ('u:overflow-new' in store.data.loginFailures), '過期後 prune 騰出空間，新名稱可記');
    ok(Object.keys(store.data.loginFailures).length < mod.MAX_ENTRIES, 'prune 後條目數下降');
  }
  // 解鎖／清單
  {
    const { g, clock } = mk();
    for (let i = 0; i < 5; i++) g.recordFailure('dave');
    for (let i = 0; i < 5; i++) g.recordFailure('ghost');
    g.recordFailure('partial');
    eq(g.listLocks().map(l => l.username).sort(), ['dave', 'ghost'], 'listLocks 只列鎖定中（含不存在的名稱）');
    eq(g.getLock('dave').lockCount, 1, 'getLock 回傳鎖定資訊');
    eq(g.unlock('dave'), true, '解鎖回 true');
    eq(g.checkLocked('dave').locked, false, '解鎖後可登入');
    eq(g.unlock('dave'), false, '重複解鎖回 false');
    for (let i = 0; i < 5; i++) g.recordFailure('dave');
    eq(g.getLock('dave').lockCount, 1, '解鎖會連倍率一起清掉（再鎖又是第 1 次）');
    clock.adv(16 * MIN); eq(g.listLocks().length, 0, '到期後 listLocks 為空');
  }
  // 設定值解析
  {
    const c = mod.readConfig({ LOGIN_LOCKOUT: 'true', LOGIN_MAX_FAILS: '3', LOGIN_LOCK_MINUTES: '0.05', SESSION_IDLE_MINUTES: '30' });
    eq([c.lockout, c.maxFails, c.lockMs, c.idleMs], [true, 3, 3000, 1800000], '環境變數解析');
    const d = mod.readConfig({ LOGIN_LOCKOUT: '1', LOGIN_MAX_FAILS: 'abc', LOGIN_LOCK_MINUTES: '-5' });
    eq([d.maxFails, d.lockMs, d.windowMs, d.lockCapMs], [5, 15 * MIN, 15 * MIN, 240 * MIN], '無效值退回預設（5 次／15 分鐘／窗 15 分鐘／上限 240 分鐘）');
  }

  // ── 閒置逾時判斷 ──
  {
    const { g } = mk({ SESSION_IDLE_MINUTES: '30' });
    const T0 = 1_700_000_000_000;
    const su = { username: 'u', loginAt: T0, lat: T0 };
    eq(g.idleDecision(su, T0 + 29 * MIN, true), { expired: false, refresh: true }, '29 分鐘＋有活動 → 刷新');
    eq(g.idleDecision(su, T0 + 30 * MIN + 1, true).expired, 'idle', '超過 30 分鐘 → idle 過期');
    eq(g.idleDecision(su, T0 + 30 * MIN, true).expired, false, '剛好 30 分鐘不算過期（> 才過期）');
    eq(g.idleDecision(su, T0 + 30 * MIN + 1, false).expired, 'idle', '輪詢請求也會因閒置過期被擋');
    eq(g.idleDecision(su, T0 + 20 * MIN, false), { expired: false, refresh: false }, '輪詢不刷新 lat');
    eq(g.idleDecision(su, T0 + 30000, true), { expired: false, refresh: false }, '一分鐘內不重簽（刷新間隔 60 秒）');
    eq(g.idleDecision(su, T0 + 60000, true).refresh, true, '滿 60 秒才刷新');
    eq(g.idleDecision({ username: 'u', loginAt: T0 }, T0 + 5 * MIN, true), { expired: false, refresh: true }, '沒有 lat 以 loginAt 為基準');
    eq(g.idleDecision({ username: 'u', loginAt: T0 }, T0 + 31 * MIN, true).expired, 'idle', '沒有 lat：loginAt 起算超過上限也過期');
    eq(g.idleDecision({ username: 'u' }, T0, true), { expired: false, refresh: true }, '沒有 lat 也沒有 loginAt：視為剛活動並刷新');
    eq(g.idleDecision({ username: 'u', loginAt: T0, lat: T0 + 8 * 3600e3 - 1000 }, T0 + 8 * 3600e3 + 1000, true).expired, 'absolute', '登入滿 8 小時（即使剛活動）→ absolute 過期，續期不會延長絕對壽命');
    eq(g.idleDecision({ username: 'u', loginAt: T0, lat: T0 + 100 * MIN }, T0 + 50 * MIN, true).expired, false, 'lat 在未來（時鐘偏差）不誤判過期');
    const g2 = mk({ SESSION_IDLE_MINUTES: '0.1' }).g;
    eq(g2.idleDecision({ loginAt: T0, lat: T0 }, T0 + 1400, true).refresh, false, '小數分鐘：刷新間隔 = min(60s, idle/4)=1.5s（1.4s 不刷）');
    eq(g2.idleDecision({ loginAt: T0, lat: T0 }, T0 + 1500, true).refresh, true, '1.5s 刷新');
  }
  // 輪詢判斷
  {
    const P = mod.isPollingRequest;
    eq([P('GET', '/api/poll-bundle'), P('GET', '/api/notifications'), P('GET', '/api/announcements/active'), P('GET', '/API/POLL-BUNDLE/')], [true, true, true, true], '三個輪詢端點（GET）不算活動');
    eq([P('PUT', '/api/notifications/read-all'), P('POST', '/api/announcements/x/ack'), P('GET', '/api/contacts'), P('GET', '/api/me'), P('POST', '/api/poll-bundle')], [false, false, false, false, false], '寫入與一般端點都算活動');
    eq([mod.isPageRequest('GET', '/'), mod.isPageRequest('GET', '/admin.html'), mod.isPageRequest('GET', '/app.js'), mod.isPageRequest('POST', '/')], [true, true, false, false], '頁面請求判斷');
  }
  // 中介層行為（假 req/res）
  const mkReq = (method, p, user, jwt) => {
    const calls = { destroy: 0, resign: 0 };
    const state = { user };
    const session = new Proxy(state, {
      get(t, k) { if (k === 'destroy') return (cb) => { calls.destroy++; t.user = null; cb && cb(); }; return t[k]; },
      set(t, k, v) { t[k] = v; if (k === 'user') calls.resign++; return true; }
    });
    const res = { code: null, body: null, cleared: 0, status(c) { this.code = c; return this; }, json(b) { this.body = b; return this; }, clearCookie() { this.cleared++; } };
    return { req: { method, path: p, session }, res, calls };
  };
  {
    const T0 = 1_700_000_000_000;
    const clock = mkClock(T0 + 2 * MIN);
    const g = mod.create({ env: { SESSION_IDLE_MINUTES: '5' }, clock: clock.now, store: mkStore() });
    const expired = []; const mw = g.idleMiddleware({ isJwt: true, onExpire: (rq, u, k) => expired.push([u, k]) });
    // 活動：重簽 JWT
    let { req, res, calls } = mkReq('GET', '/api/contacts', { username: 'u', loginAt: T0, lat: T0 });
    let n = 0; mw(req, res, () => n++);
    eq([n, calls.resign, req.session.user.lat], [1, 1, clock.now()], 'JWT：有活動且距上次 > 刷新間隔 → 放行並重簽（lat=現在）');
    // 輪詢：不重簽
    ({ req, res, calls } = mkReq('GET', '/api/poll-bundle', { username: 'u', loginAt: T0, lat: T0 })); n = 0; mw(req, res, () => n++);
    eq([n, calls.resign, req.session.user.lat], [1, 0, T0], '輪詢請求不重簽、不更新 lat');
    // 沒有 loginAt 的舊 JWT：不重簽（避免續掉 8 小時上限），但仍放行
    ({ req, res, calls } = mkReq('GET', '/api/contacts', { username: 'u' })); n = 0; mw(req, res, () => n++);
    eq([n, calls.resign], [1, 0], '沒有 loginAt 的舊 token：不重簽');
    // express-session 路徑：就地改 lat，不指派 session.user
    const mwS = g.idleMiddleware({ isJwt: false });
    ({ req, res, calls } = mkReq('POST', '/api/visits', { username: 'u', loginAt: T0, lat: T0 })); n = 0; mwS(req, res, () => n++);
    eq([n, calls.resign, req.session.user.lat], [1, 0, clock.now()], 'express-session：就地更新 lat（由 store 保存），不重簽');
    // 過期：API → 401 + code + destroy
    clock.adv(10 * MIN);
    ({ req, res, calls } = mkReq('GET', '/api/contacts', { username: 'u', loginAt: T0, lat: T0 })); n = 0; mw(req, res, () => n++);
    eq([n, res.code, res.body && res.body.code, calls.destroy, res.cleared], [0, 401, 'SESSION_IDLE_TIMEOUT', 1, 1], '閒置過期：401 SESSION_IDLE_TIMEOUT、session 銷毀、cookie 清除、不放行');
    eq(expired[expired.length - 1], ['u', 'idle'], 'onExpire 收到 (user, idle)');
    // 輪詢也會被過期擋
    ({ req, res } = mkReq('GET', '/api/poll-bundle', { username: 'u', loginAt: T0, lat: T0 })); n = 0; mw(req, res, () => n++);
    eq([n, res.code], [0, 401], '過期後輪詢請求同樣 401');
    // 頁面：銷毀但放行（交給 requireAuth 導向登入）；資產：完全不管
    ({ req, res, calls } = mkReq('GET', '/', { username: 'u', loginAt: T0, lat: T0 })); n = 0; mw(req, res, () => n++);
    eq([n, res.code, calls.destroy], [1, null, 1], '頁面請求過期：銷毀 session 後放行（requireAuth 會導向登入）');
    ({ req, res, calls } = mkReq('GET', '/style.css', { username: 'u', loginAt: T0, lat: T0 })); n = 0; mw(req, res, () => n++);
    eq([n, calls.destroy, calls.resign], [1, 0, 0], '靜態資產不受管（不擋、不算活動）');
    ({ req, res, calls } = mkReq('POST', '/api/login', { username: 'u', loginAt: T0, lat: T0 })); n = 0; mw(req, res, () => n++);
    eq([n, calls.destroy], [1, 0], '/api/login 不受閒置檢查');
    ({ req, res } = mkReq('GET', '/api/me', null)); n = 0; mw(req, res, () => n++);
    eq(n, 1, '未登入請求放行');
    // 絕對 8 小時
    clock.t = T0 + 9 * 3600e3;
    ({ req, res } = mkReq('GET', '/api/contacts', { username: 'u', loginAt: T0, lat: clock.t - 1000 })); n = 0; mw(req, res, () => n++);
    eq([n, res.code, res.body && res.body.code], [0, 401, 'SESSION_EXPIRED'], '登入滿 8 小時：401 SESSION_EXPIRED（續期不延長）');
  }
}

// 變異：把規則改壞，單元測試必須失敗
const MUTANTS = [
  ['鎖定門檻 >= 改成 >', "if (e.n >= cfg.maxFails)", "if (e.n > cfg.maxFails)"],
  ['鎖定中仍累計/延長（拿掉鎖定中提早返回）', "if (e && (e.lockUntil || 0) > t) return { locked: true, justLocked: false, until: e.lockUntil, lockCount: e.lockCount || 0 };   // 鎖定中不累計、不延長", "/* mut */"],
  ['不加倍', "Math.pow(2, e.lockCount - 1)", "1"],
  ['拿掉加倍上限', "Math.max(cfg.lockCapMs, cfg.lockMs))", "1e15)"],
  ['成功登入不重置', "delete c.m[k]; persist(c); return true;\n    },\n\n    /** 管理員解鎖", "return false;\n    },\n\n    /** 管理員解鎖"],
  ['存到別份副本（JSON 後端會失效）', "try { store.save(c.d); }", "try { store.save(store.load()); }"],
  ['解鎖不持久化', "delete c.m[k]; persist(c); return true;\n    },\n\n    /** 目前鎖定中", "delete c.m[k]; return true;\n    },\n\n    /** 目前鎖定中"],
  ['視窗過期不重置', "if (!e.n || t - (e.first || 0) > cfg.windowMs) { e.n = 0; e.first = t; }", "if (!e.n) { e.n = 0; e.first = t; }"],
  ['不做大小寫正規化', ".trim().toLowerCase().slice(0, 64)", ".trim().slice(0, 64)"],
  ['鍵拿掉前綴', "const PREFIX = 'u:';", "const PREFIX = '';"],
  ['閒置判斷 > 改 >=', "now - base > cfg.idleMs", "now - base >= cfg.idleMs"],
  ['輪詢也刷新', "if (!counts) return { expired: false, refresh: false };", ""],
  ['拿掉 8 小時絕對上限', "if (loginAt && now - loginAt > ABSOLUTE_SESSION_MS) return { expired: 'absolute', refresh: false };", ""],
  ['關閉時仍計數', "if (!cfg.lockout) return { locked: false, justLocked: false };", ""],
  ['衰減拿掉', "if (t - (e.last || 0) > DECAY_MS) e.lockCount = 0;", ""],
  ['無 loginAt 的 JWT 也重簽', "if (o.isJwt && su.loginAt)", "if (o.isJwt)"],
  ['輪詢清單拿掉 poll-bundle', "['/api/poll-bundle', '/api/notifications', '/api/announcements/active']", "['/api/notifications', '/api/announcements/active']"],
];

async function partA() {
  console.log('[A] 單元測試');
  const real = require(LG_PATH);
  unitSuite(real, { ok, eq });
  console.log('[A-M] 變異檢查（每個變異都必須讓單元測試失敗）');
  const src = fs.readFileSync(LG_PATH, 'utf8');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'lgmut-'));
  try {
    MUTANTS.forEach(([name, from, to], i) => {
      if (src.split(from).length !== 2) { ok(false, `變異樣式找不到或重複：${name}`); return; }
      const f = path.join(tmp, `m${i}.js`);
      fs.writeFileSync(f, src.replace(from, () => to));
      let bad = 0; let threw = false;
      try { unitSuite(require(f), { ok: (c) => { if (!c) bad++; }, eq: (a, b) => { if (JSON.stringify(a) !== JSON.stringify(b)) bad++; } }); } catch (e) { threw = true; }
      ok(bad > 0 || threw, `變異「${name}」被單元測試抓到（失敗數 ${bad}${threw ? '，且拋例外' : ''}）`);
    });
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
}

// ═══════════════════════ Part B：端對端 ═══════════════════════
const { mkUser, http, login, loginAll, getLogs, byAction, PW, jwtPayload } = H;
const USERS = () => [
  mkUser('_s_admin', 'admin'), mkUser('_s_admin2', 'admin'),
  mkUser('_s_u1', 'user', { bu: ['ERP'] }), mkUser('_s_u2', 'user', { bu: ['ERP'] }), mkUser('_s_u3', 'user', { bu: ['ERP'] }),
  mkUser('_s_slow', 'user', { bu: ['ERP'], password: H.HASH12 }),
];
const bad = async (user, pw) => http(null, 'POST', '/api/login', { username: user, password: pw || 'wrong-' + Math.random() });
const LOCK_ENV = { LOGIN_LOCKOUT: '1', LOGIN_MAX_FAILS: '3', LOGIN_LOCK_MINUTES: '0.05', LOGIN_LOCK_MAX_MINUTES: '0.2' };   // 鎖 3 秒、上限 12 秒

async function e2eDefaults() {
  console.log('[B1] 開關關閉（預設）：與先前行為相同');
  await H.fresh(USERS());
  for (let i = 0; i < 8; i++) { const r = await bad('_s_u1'); if (i === 0) eq([r.status, r.json], [401, { success: false, message: '帳號或密碼錯誤' }], '錯誤密碼的回應不變（401＋通用訊息）'); }
  let r = await login('_s_u1');
  eq(r.status, 200, '連續 8 次失敗後正確密碼仍可登入（鎖定關閉）');
  const keys = Object.keys(jwtPayload(H.jar._s_u1)).sort();
  eq(keys, ['exp', 'iat', 'lat', 'lgn', 'user'], 'JWT 外層欄位（整併 main 的閒置登出 PR #21 後多了 lat／lgn，其餘不變）');
  const uk = Object.keys(jwtPayload(H.jar._s_u1).user).sort();
  ok(!uk.includes('lat'), '關閉時 session 內沒有 lat 欄位（cookie 內容與先前相同）');
  await loginAll(['_s_admin']);
  const ul = await http('_s_admin', 'GET', '/api/admin/users');
  ok(ul.json.every(u => u.locked === false && u.lockUntil === null), '帳號列表 locked 全為 false');
  const data = JSON.parse(fs.readFileSync(path.join(H.ROOT, 'data.json'), 'utf8'));
  ok(!('loginFailures' in data), '關閉時 data.json 沒有 loginFailures');
  const v = await http('_s_admin', 'GET', '/api/admin/version');
  eq([v.json.loginGuard.lockout, v.json.loginGuard.idleMinutes], [false, 0], '版本頁顯示兩個開關皆關');
  await sleep(100);
  // 活動後 cookie 不重簽（Set-Cookie 不出現）
  r = await http('_s_u1', 'GET', '/api/contacts'); r = await http('_s_u1', 'GET', '/api/contacts');
  ok(!r.setCookie.some(c => /^itts_auth=/.test(c)), '關閉時一般請求不重簽 cookie');
}

async function e2eLockout() {
  console.log('[B2] 帳號鎖定（JWT 後端）');
  await H.fresh(USERS(), {}, { env: LOCK_ENV });
  const GEN = { success: false, message: '帳號或密碼錯誤' };
  // 基準：未鎖定時，錯誤密碼的回應與耗時（_s_slow 是 cost 12，與正式環境一致）
  const baseline = []; for (let i = 0; i < 2; i++) { const r = await bad('_s_slow'); baseline.push(r.ms); eq([r.status, r.json], [401, GEN], '第 ' + (i + 1) + ' 次錯誤：通用 401'); }
  let r = await bad('_s_slow'); eq([r.status, r.json], [401, GEN], '第 3 次錯誤（觸發鎖定）仍是通用 401');
  // 鎖定中：正確密碼也被拒絕、回應完全相同
  const lockedMs = [];
  r = await login('_s_slow'); lockedMs.push(r.ms);
  eq([r.status, r.json], [401, GEN], '鎖定期間「正確密碼」也 401 且訊息與密碼錯誤完全相同（不洩漏已鎖定）');
  ok(!r.setCookie.some(c => /^itts_auth=[^;]/.test(c)), '鎖定期間不發 cookie');
  r = await bad('_s_slow'); lockedMs.push(r.ms);
  const med = a => a.slice().sort((x, y) => x - y)[Math.floor(a.length / 2)];
  ok(Math.min(...lockedMs) > 0.4 * Math.min(...baseline), `鎖定回應耗時與錯誤密碼同量級（鎖定 ${lockedMs.join('/')}ms，基準 ${baseline.join('/')}ms；不會快速短路）`);
  // 不存在的帳號：同樣被鎖、行為一致
  for (let i = 0; i < 3; i++) { r = await bad('_s_ghost'); eq([r.status, r.json], [401, GEN], '不存在帳號第 ' + (i + 1) + ' 次：通用 401'); }
  r = await login('_s_ghost'); eq([r.status, r.json], [401, GEN], '不存在帳號鎖定後的回應與存在者相同');
  // 其他帳號不受影響
  r = await login('_s_u1'); eq(r.status, 200, '鎖定 _s_slow 不影響 _s_u1');
  await loginAll(['_s_admin']);
  // 管理員檢視
  let ul = await http('_s_admin', 'GET', '/api/admin/users');
  const slow = ul.json.find(u => u.username === '_s_slow');
  eq([slow.locked, typeof slow.lockUntil], [true, 'number'], '後台帳號列表：_s_slow locked=true 且有 lockUntil');
  eq(ul.json.filter(u => u.locked).map(u => u.username), ['_s_slow'], '只有 _s_slow 顯示鎖定（不存在的名稱不會出現在帳號列表）');
  const lk = await http('_s_admin', 'GET', '/api/admin/login-locks');
  eq(lk.json.locks.map(l => [l.username, l.exists]).sort(), [['_s_ghost', false], ['_s_slow', true]], '鎖定清單（管理員專用）列出兩者，並標示是否為既有帳號');
  eq(Object.keys(lk.json.locks[0]).sort(), ['exists', 'last', 'lockCount', 'lockUntil', 'username'], '鎖定清單只含 username/exists/lockUntil/lockCount/last');
  r = await http('_s_u1', 'GET', '/api/admin/login-locks'); eq(r.status, 403, '非管理員不能看鎖定清單');
  r = await http('_s_u1', 'POST', '/api/admin/users/_s_slow/unlock'); eq(r.status, 403, '非管理員不能解鎖');
  // 稽核
  let logs = await getLogs();
  const lockedEv = byAction(logs, 'LOGIN_LOCKED');
  eq(lockedEv.map(l => l.operator).sort(), ['_s_ghost', '_s_slow'], 'LOGIN_LOCKED 各一筆（含不存在的名稱）');
  ok(lockedEv.every(l => /第 1 次鎖定/.test(l.detail)), 'LOGIN_LOCKED 內容含鎖定次數');
  const lf = byAction(logs, 'LOGIN_FAIL').filter(l => l.operator === '_s_slow');
  ok(lf.length >= 1 && lf.length <= 3, '波一 LOGIN_FAIL 仍記錄且 30 秒節流仍生效（_s_slow 共 5 次失敗＋1 次鎖定期間 → 合併成 ' + lf.length + ' 筆）');
  ok(lf.some(l => /原因=LOCKED/.test(l.detail)), '鎖定期間的嘗試以原因=LOCKED 記錄');
  ok(!lf.some(l => /\$2[aby]\$/.test(l.detail) || l.detail.includes(PW)), '稽核不含密碼或雜湊');
  // 管理員解鎖
  r = await http('_s_admin', 'POST', '/api/admin/users/_s_slow/unlock');
  eq([r.status, r.json && r.json.wasLocked], [200, true], '管理員解鎖');
  r = await login('_s_slow'); eq(r.status, 200, '解鎖後立即可用正確密碼登入');
  await H.restartKeep({ env: LOCK_ENV });   // 重啟歸零行程內的登入限流（15 分鐘 20 次）；舊 cookie 仍有效（測試行程用固定密鑰）
  logs = await getLogs();
  ok(byAction(logs, 'LOGIN_UNLOCK').some(l => l.target === '_s_slow' && l.operator === '_s_admin'), 'LOGIN_UNLOCK 有稽核（操作者、目標）');
  ul = await http('_s_admin', 'GET', '/api/admin/users'); ok(!ul.json.find(u => u.username === '_s_slow').locked, '解鎖後徽章消失');
  // 計數歸零：2 次失敗 → 成功 → 2 次失敗，不鎖
  await bad('_s_u2'); await bad('_s_u2'); r = await login('_s_u2'); eq(r.status, 200, '2 次失敗後成功登入');
  await bad('_s_u2'); await bad('_s_u2'); r = await login('_s_u2'); eq(r.status, 200, '成功登入重置計數：再 2 次失敗仍未鎖');
  // 持久化：重啟後仍鎖定；blob 內只有數字
  await bad('_s_u3'); await bad('_s_u3'); await bad('_s_u3');
  const raw = fs.readFileSync(path.join(H.ROOT, 'data.json'), 'utf8');
  const lf2 = JSON.parse(raw).loginFailures;
  ok(lf2 && lf2['u:_s_u3'] && Object.values(lf2['u:_s_u3']).every(v => typeof v === 'number'), 'data.json 的 loginFailures 只存數字');
  ok(!/\$2[aby]\$/.test(JSON.stringify(lf2)), 'loginFailures 不含雜湊');
  await H.restartKeep({ env: LOCK_ENV });
  r = await login('_s_u3'); eq([r.status, r.json], [401, GEN], '重啟伺服器後仍在鎖定中（狀態已持久化）');
  // 到期自動解除 + 加倍
  await sleep(3300);
  r = await login('_s_u3'); eq(r.status, 200, '3 秒鎖定到期後正確密碼可登入（成功後歸零）');
  await bad('_s_u1'); await bad('_s_u1'); await bad('_s_u1');   // 第 1 次鎖
  await sleep(3300);                                              // 到期
  await bad('_s_u1'); await bad('_s_u1'); await bad('_s_u1');   // 第 2 次鎖：6 秒
  await loginAll(['_s_admin']);
  const lk2 = await http('_s_admin', 'GET', '/api/admin/login-locks');
  const l1 = lk2.json.locks.find(l => l.username === '_s_u1');
  ok(l1 && l1.lockCount === 2, '第二次鎖定 lockCount=2');
  const remain = l1 ? l1.lockUntil - Date.now() : 0;
  ok(remain > 3500 && remain <= 6200, `第二次鎖定約 6 秒（剩餘 ${remain}ms，第一次是 3 秒）`);
  // 管理員重設密碼同時解鎖
  r = await http('_s_admin', 'PUT', '/api/admin/users/_s_u1/password', { password: PW + 'x1' });
  eq(r.status, 200, '管理員重設密碼');
  r = await login('_s_u1', '_s_u1', PW + 'x1'); eq(r.status, 200, '重設密碼後（解除鎖定）可用新密碼登入');
  // 管理員自己被鎖：另一位管理員可解
  await bad('_s_admin2'); await bad('_s_admin2'); await bad('_s_admin2');
  r = await login('_s_admin2'); eq(r.status, 401, '管理員帳號也會被鎖');
  r = await http('_s_admin', 'POST', '/api/admin/users/_s_admin2/unlock'); eq(r.status, 200, '另一位管理員可解鎖');
  r = await login('_s_admin2'); eq(r.status, 200, '被鎖的管理員解鎖後可登入');
  // 版本頁
  const v = await http('_s_admin', 'GET', '/api/admin/version');
  eq([v.json.loginGuard.lockout, v.json.loginGuard.maxFails], [true, 3], '版本頁顯示鎖定已開、門檻 3');
}

async function e2eIdle(backend) {
  console.log(`[B3] 閒置逾時（${backend}）`);
  await H.fresh(USERS(), {}, { env: { SESSION_IDLE_MINUTES: '0.1', SESSION_BACKEND: backend } });   // 6 秒；刷新間隔 1.5 秒
  const isJwt = backend === 'jwt';
  await loginAll(['_s_u1', '_s_u2', '_s_admin']);
  let r = await http('_s_u1', 'GET', '/api/contacts'); eq(r.status, 200, '剛登入：200');
  if (isJwt) {
    const p0 = jwtPayload(H.jar._s_u1).user;
    ok(!p0.lat || p0.lat >= p0.loginAt, 'JWT 內 lat 欄位（簽章保護）');
  }
  // 持續活動：每 4 秒一次，共 12 秒 > 6 秒上限，仍然有效（滑動視窗）
  let lastLat = isJwt ? (jwtPayload(H.jar._s_u1).user.lat || 0) : 0;
  for (let i = 0; i < 3; i++) {
    await sleep(4000);
    r = await http('_s_u1', 'GET', '/api/contacts');
    eq(r.status, 200, `活動第 ${i + 1} 輪（距上次 4 秒 < 上限 6 秒）仍有效`);
    if (isJwt) { const lat = jwtPayload(H.jar._s_u1).user.lat; ok(lat > lastLat, `第 ${i + 1} 輪 cookie 被重簽且 lat 前進`); lastLat = lat; }
  }
  // 只有輪詢：不算活動 → 6 秒後過期
  const pollStatuses = [];
  for (let i = 0; i < 5; i++) { await sleep(2000); const p = await http('_s_u1', 'GET', '/api/poll-bundle?days=3'); pollStatuses.push(p.status); if (p.status === 401) { r = p; break; } }
  ok(pollStatuses.includes(401), '只有輪詢（poll-bundle）不延長 session：最終 401（' + pollStatuses.join(',') + '）');
  eq(r.json && r.json.code, 'SESSION_IDLE_TIMEOUT', '回應 code=SESSION_IDLE_TIMEOUT');
  if (isJwt) ok(r.setCookie.some(c => /^itts_auth=;/.test(c) || /Expires=Thu, 01 Jan 1970/.test(c)), '回應清除 cookie');
  r = await http('_s_u1', 'GET', '/api/contacts', undefined, isJwt ? { cookie: H.jar._s_u1 } : undefined);
  eq(r.status, 401, '被登出後再用任何 API 仍 401');
  // 頁面導向
  r = await http('_s_u2', 'GET', '/api/me'); // u2 登入後閒置超過 → 看頁面請求
  await sleep(100);
  await login('_s_admin');   // 管理員自己也閒置過期了，重新登入才能讀稽核
  const logs = await getLogs('_s_admin');
  ok(byAction(logs, 'SESSION_IDLE_TIMEOUT').some(l => l.operator === '_s_u1'), '閒置登出有稽核 SESSION_IDLE_TIMEOUT');
  // u2：閒置 7 秒後直接開頁面 → 導向登入（session 被銷毀）
  await sleep(7000);
  r = await http('_s_u2', 'GET', '/');
  ok([302].includes(r.status) && /login\.html/.test(r.headers.get('location') || ''), '閒置過久後開頁面 → 302 登入頁（' + r.status + '）');
  // 活動不被輪詢以外端點誤判：/api/me 算活動
  await login('_s_u3'); await sleep(4000); r = await http('_s_u3', 'GET', '/api/me'); eq(r.status, 200, '/api/me 算活動');
  await sleep(4000); r = await http('_s_u3', 'GET', '/api/me'); eq(r.status, 200, '/api/me 持續活動 → 仍有效（共 8 秒 > 6 秒上限）');
  // 一般 API 寫入也算活動：_s_u1 重新登入
  await login('_s_u1');
  await sleep(4000); r = await http('_s_u1', 'POST', '/api/key-accounts', { company: 'X' }); ok(r.status !== 401, 'POST 算活動（' + r.status + '）');
  await sleep(4000); r = await http('_s_u1', 'GET', '/api/contacts'); eq(r.status, 200, 'POST 刷新後 4 秒仍有效');
}

async function e2eMutation() {
  console.log('[B-M] 伺服器層變異檢查');
  const fsx = fs; const tmp = [];
  const mk = (name, edit) => {
    let s = fsx.readFileSync(path.join(H.ROOT, 'server.js'), 'utf8');
    if (!s.includes(edit[0])) throw new Error('變異樣式找不到：' + name);
    s = s.replace(edit[0], () => edit[1]);
    const f = path.join(H.ROOT, 'server.__mutlg_' + name + '.js'); fsx.writeFileSync(f, s); tmp.push(f); return f;
  };
  try {
    // 1. 登入端不檢查鎖定 → 鎖定期間正確密碼會成功
    let f = mk('m1', ["if (loginGuard.checkLocked(username).locked) {", 'if (false) {']);
    await H.fresh(USERS(), {}, { env: LOCK_ENV, file: f });
    for (let i = 0; i < 3; i++) await bad('_s_u1');
    let r = await login('_s_u1');
    ok(r.status === 200, '變異1（登入不檢查鎖定）→ 鎖定期間正確密碼竟然成功(' + r.status + ')：B2 的 401 斷言會失敗（抓得到）');
    // 2. 失敗不計數（loginFailed 不呼叫 recordFailure）
    f = mk('m2', ["const r = loginGuard.recordFailure(attempted);", 'const r = null;']);
    await H.fresh(USERS(), {}, { env: LOCK_ENV, file: f });
    for (let i = 0; i < 4; i++) await bad('_s_u1');
    r = await login('_s_u1');
    ok(r.status === 200, '變異2（失敗不計數）→ 永遠鎖不住(' + r.status + ')：B2 斷言會失敗（抓得到）');
    // 3. 閒置中介層沒掛
    f = mk('m3', ["app.use(loginGuard.idleMiddleware({", "app.use((function () { return (q, s, n) => n(); })({"]);
    await H.fresh(USERS(), {}, { env: { SESSION_IDLE_MINUTES: '0.05' }, file: f });
    await login('_s_u1'); await sleep(4000);
    r = await http('_s_u1', 'GET', '/api/contacts');
    ok(r.status === 200, '變異3（沒掛閒置中介層）→ 閒置後仍 200：B3 的 401 斷言會失敗（抓得到）');
    // 4. 管理員解鎖路由不清除
    f = mk('m4', ["const had = loginGuard.unlock(target);", 'const had = false;']);
    await H.fresh(USERS(), {}, { env: LOCK_ENV, file: f });
    await loginAll(['_s_admin']);
    for (let i = 0; i < 3; i++) await bad('_s_u1');
    await http('_s_admin', 'POST', '/api/admin/users/_s_u1/unlock');
    r = await login('_s_u1');
    ok(r.status === 401, '變異4（解鎖不生效）→ 解鎖後仍無法登入：B2 的解鎖斷言會失敗（抓得到）');
  } finally { await H.stopServer(); tmp.forEach(f => { try { fsx.unlinkSync(f); } catch { /* 已刪 */ } }); }
}

(async () => {
  H.guardClean();
  try {
    await partA();
    if (!process.env.LG_SKIP_E2E) {
      await e2eDefaults();
      await e2eLockout();
      await e2eIdle('jwt');
      await e2eIdle('memory');
      if (!process.env.LG_SKIP_MUT) await e2eMutation();
    }
  } catch (e) { H.R.fail++; H.R.failures.push('例外：' + (e && e.stack || e)); console.log('  ✗ 例外', e && e.stack || e); }
  finally { await H.finish(); }
  process.exit(H.report('check-login-guard '));
})();
