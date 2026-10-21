#!/usr/bin/env node
'use strict';
/**
 * 公告功能檢查（單元 + 端對端）。不含任何帳密：全部讀環境變數。
 *
 *   單元（永遠執行）：lib/announcements.js 的驗證／時區／適用判斷／統計，
 *                    以及 _client/announcements.js 在最小 DOM 替身下的行為（XSS 惰性、失敗隔離、每次登入判斷）。
 *   端對端（設了 ANN_BASE 才執行）：打一台「隔離沙盒」伺服器（不可是正式站！）。
 *
 * 用法：
 *   node scripts/check-announcements.js                       # 只跑單元
 *   ANN_BASE=http://localhost:3041 ANN_PW=<沙盒測試密碼> node scripts/check-announcements.js
 *
 * 端對端需要沙盒內已存在這些臨時帳號（密碼都是 ANN_PW；可用環境變數改名）：
 *   ANN_ADMIN(_s_admin, role admin) / ANN_USER(_s_user, role user) / ANN_MGR(_s_mgr, manager1)
 *   ANN_TPM(_s_tpm, tecopm) / ANN_GS(_s_gs, groupsales) / ANN_VIEW(_s_view, role user + accessMode view)
 * 端對端會建立並在結束時刪除自己建立的公告（只動 title 以 "[ANNCHK]" 開頭者）。
 * ANN_LIB 可指定另一份 lib/announcements.js（變異測試用）。
 */
const path = require('path');
const LIB = process.env.ANN_LIB || path.join(__dirname, '..', 'lib', 'announcements.js');
const lib = require(LIB);
const fs = require('fs');

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) {
  if (cond) { pass++; } else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); }
}
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const ROLES = ['admin','executive','manager1','manager2','secretary','tecopm','groupsales','marketing','user','accounting_manager','finance_manager','consult_manager_south','consult_manager_north'];
const XSS = '<img src=x onerror=alert(1)>';

// ───────────────────────── 單元：lib ─────────────────────────
function unitLib() {
  console.log('[unit] lib/announcements.js');
  const v = (b, partial) => lib.validateInput(b, ROLES, !!partial);
  ok(v({ title: 'T', body: 'B' }).fields, '最小輸入通過');
  eq(v({ title: 'T', body: 'B' }).fields.mode, 'once', '預設 mode=once');
  eq(v({ title: 'T', body: 'B' }).fields.roles, [], '預設 roles=全部（空陣列）');
  eq(v({ title: 'T', body: 'B' }).fields.priority, 100, '預設 priority=100');
  ok(v({ title: '', body: 'B' }).error, '空標題拒絕');
  ok(v({ title: 'x'.repeat(101), body: 'B' }).error, '標題 101 字拒絕');
  ok(v({ title: 'x'.repeat(100), body: 'B' }).fields, '標題 100 字通過');
  ok(v({ title: 'T', body: 'x'.repeat(4001) }).error, '內容 4001 字拒絕');
  ok(v({ title: 'T', body: 'x'.repeat(4000) }).fields, '內容 4000 字通過');
  ok(v({ title: 'T', body: 'B', mode: 'always' }).error, '未知 mode 拒絕');
  ok(v({ title: 'T', body: 'B', severity: 'x' }).error, '未知 severity 拒絕');
  ok(v({ title: 'T', body: 'B', roles: ['nope'] }).error, '未知角色拒絕');
  ok(v({ title: 'T', body: 'B', roles: 'user' }).error, 'roles 非陣列拒絕');
  eq(v({ title: 'T', body: 'B', roles: ROLES }).fields.roles, [], '全選角色＝全部（空陣列）');
  eq(v({ title: 'T', body: 'B', roles: ['user', 'user'] }).fields.roles, ['user'], '角色去重');
  ok(v({ title: 'T', body: 'B', priority: 10000 }).error, 'priority 超界拒絕');
  ok(v({ title: 'T', body: 'B', startAt: 'garbage' }).error, '壞時間格式拒絕');
  ok(v({ title: 'T', body: 'B', startAt: '2026-02-31T10:00' }).error, '不存在的日期拒絕');
  eq(v({ title: 'T', body: 'a\r\nb\u0000c' }).fields.body, 'a\nbc', '換行統一、控制字元移除');
  eq(v({ title: 'a\nb', body: 'B' }).fields.title, 'a b', '標題換行轉空白');
  eq(v({ title: XSS, body: XSS }).fields.title, XSS, 'XSS 字串原樣保存（由顯示端跳脫，不在伺服器改寫）');
  eq(Object.keys(v({ enabled: false }, true).fields), ['enabled'], 'partial 只驗證有帶的欄位');
  ok(v({ title: '' }, true).error, 'partial 帶空標題仍拒絕');
  // 時區
  eq(lib.parseTaipeiInput('2026-10-12T09:00').value, '2026-10-12T01:00:00.000Z', 'datetime-local 以台北時間解讀（+8）');
  eq(lib.toTaipeiLocal('2026-10-12T01:00:00.000Z'), '2026-10-12T09:00', 'ISO → 台北 datetime-local');
  eq(lib.toTaipeiLocal('2026-10-12T17:30:00.000Z'), '2026-10-13T01:30', '跨日換算');
  eq(lib.parseTaipeiInput('2026-10-12T09:00:00+08:00').value, '2026-10-12T01:00:00.000Z', '帶時區字串照字面');
  eq(lib.parseTaipeiInput('').value, null, '空字串＝不限制');
  ok(lib.checkWindow({ startAt: '2026-10-12T02:00:00Z', endAt: '2026-10-12T01:00:00Z' }), '結束早於開始拒絕');
  ok(!lib.checkWindow({ startAt: '2026-10-12T01:00:00Z', endAt: '2026-10-12T02:00:00Z' }), '正常區間通過');
  // 適用判斷
  const now = Date.parse('2026-10-12T04:00:00Z');
  const mk = (o) => ({ id: 'a', version: 1, title: 't', body: 'b', mode: 'once', severity: 'normal', roles: [], startAt: null, endAt: null, enabled: true, priority: 100, readBy: {}, createdAt: '2026-01-01T00:00:00Z', ...o });
  const U = { username: 'u1', role: 'user' };
  eq(lib.activeFor([mk({})], U, now).length, 1, '基本公告出現');
  eq(lib.activeFor([mk({ enabled: false })], U, now).length, 0, '停用不出現');
  eq(lib.activeFor([mk({ startAt: '2026-10-12T05:00:00Z' })], U, now).length, 0, '未到開始時間不出現');
  eq(lib.activeFor([mk({ endAt: '2026-10-12T04:00:00Z' })], U, now).length, 0, '結束時間（含邊界）不出現');
  eq(lib.activeFor([mk({ startAt: '2026-10-12T03:00:00Z', endAt: '2026-10-12T05:00:00Z' })], U, now).length, 1, '區間內出現');
  eq(lib.activeFor([mk({ roles: ['manager1'] })], U, now).length, 0, '角色不符不出現');
  eq(lib.activeFor([mk({ roles: ['user'] })], U, now).length, 1, '角色符合出現');
  eq(lib.activeFor([mk({ readBy: { u1: { version: 1, at: 'x' } } })], U, now).length, 0, 'once 已確認不出現');
  eq(lib.activeFor([mk({ version: 2, readBy: { u1: { version: 1, at: 'x' } } })], U, now).length, 1, 'once 版本升級後再出現');
  eq(lib.activeFor([mk({ mode: 'everyLogin', readBy: { u1: { version: 1, at: 'x' } } })], U, now).length, 1, 'everyLogin 已確認仍回傳（每次登入由前端判斷）');
  eq(lib.activeFor([mk({ readBy: { other: { version: 1, at: 'x' } } })], U, now).length, 1, '別人確認不影響我');
  const two = lib.activeFor([mk({ id: 'b', priority: 50 }), mk({ id: 'a', priority: 10 }), mk({ id: 'c', priority: 50, createdAt: '2025-01-01T00:00:00Z' })], U, now);
  eq(two.map(x => x.id), ['a', 'c', 'b'], '依 priority（小先）再依建立時間排序');
  ok(!JSON.stringify(lib.activeFor([mk({ readBy: { u1: { version: 1, at: 'x' } }, mode: 'everyLogin' })], U, now)).includes('readBy'), 'active 回傳不含 readBy');
  eq(Object.keys(lib.activeFor([mk({})], U, now)[0]).sort(), ['body', 'id', 'mode', 'severity', 'title', 'version'], 'active 欄位最小化');
  // 統計
  const users = [{ username: 'u1', role: 'user' }, { username: 'u2', role: 'user' }, { username: 'm', role: 'manager1' }, { username: 'off', role: 'user', active: false }, { username: '_pool', role: 'pool' }];
  const st = lib.statsFor(mk({ readBy: { u1: { version: 1, at: 'x' }, off: { version: 1, at: 'x' } } }), users);
  eq(st, { readCount: 1, targetCount: 3 }, '統計：排除停用帳號與客戶池');
  eq(lib.statsFor(mk({ roles: ['user'], readBy: { u1: { version: 1, at: 'x' } } }), users), { readCount: 1, targetCount: 2 }, '統計：依目標角色');
  eq(lib.statsFor(mk({ version: 2, readBy: { u1: { version: 1, at: 'x' } } }), users).readCount, 0, '統計：升版後舊版已讀不算');
  ok(!('readBy' in lib.adminView(mk({ readBy: { u1: {} } }), users, now)), 'adminView 不含 readBy 內容');
  ok(!('shown' in lib.adminView(mk({ shown: { u1: { version: 1, at: 1 } } }), users, now)), 'adminView 不含 shown');
  { const a = mk({}); eq(lib.ackTooEarly(a, 'u1', now), true, '從未送達 → ack 太早');
    eq(lib.markShown(a, 'u1', now), true, 'markShown 首次記錄');
    eq(lib.ackTooEarly(a, 'u1', now + 3999), true, '送達後 3.999 秒 ack → 太早');
    eq(lib.ackTooEarly(a, 'u1', now + 4000), false, '送達後 4 秒 ack → 放行');
    eq(lib.markShown(a, 'u1', now + 1000), false, '4 秒內重複送達不重設起點');
    eq(lib.markShown(a, 'u1', now + 4000), true, '超過 4 秒的再次送達會更新起點');
    eq(lib.ackTooEarly(a, 'u1', now + 4001), true, '更新後倒數重新計算');
    eq(lib.ackTooEarly(a, 'u2', now + 99999), true, '別人的送達紀錄不能替我放行'); }
  eq(lib.status(mk({ enabled: false }), now), 'disabled', 'status disabled');
  eq(lib.status(mk({ startAt: '2026-10-12T05:00:00Z' }), now), 'scheduled', 'status scheduled');
  eq(lib.status(mk({ endAt: '2026-10-12T03:00:00Z' }), now), 'expired', 'status expired');
  eq(lib.status(mk({}), now), 'live', 'status live');
}

// ───────────────────── 單元：前端元件（最小 DOM 替身）─────────────────────
function makeDom() {
  class El {
    constructor(tag) { this.tag = tag; this.children = []; this.attrs = {}; this.listeners = {}; this.parentNode = null; this._text = ''; this.className = ''; this.id = ''; this.disabled = false; this.innerHTMLSet = false; }
    setAttribute(k, v) { this.attrs[k] = String(v); }
    appendChild(c) { c.parentNode = this; this.children.push(c); return c; }
    removeChild(c) { this.children = this.children.filter(x => x !== c); c.parentNode = null; }
    addEventListener(t, f) { this.listeners[t] = f; }
    focus() {}
    set textContent(v) { this._text = String(v); }
    get textContent() { return this._text; }
    set innerHTML(v) { this.innerHTMLSet = true; }
    all(fn, out = []) { if (fn(this)) out.push(this); this.children.forEach(c => c.all(fn, out)); return out; }
  }
  const body = new El('body'), head = new El('head');
  const doc = { body, head, documentElement: head, createElement: t => new El(t), getElementById: id => head.all(e => e.id === id)[0] || null };
  return { doc, body, head, El };
}
async function unitClient() {
  console.log('[unit] _client/announcements.js（DOM 替身）');
  const src = fs.readFileSync(path.join(__dirname, '..', '_client', 'announcements.js'), 'utf8');
  async function run(fetchImpl, { username = 'u', ls } = {}) {
    const { doc, body } = makeDom();
    const store = ls || {};
    const win = { _myUsername: username };
    const localStorage = { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); } };
    const calls = [];
    const timers = new Map(); let tid = 0;
    const fakeSetInterval = (fn) => { timers.set(++tid, fn); return tid; };
    const fakeClearInterval = (id) => { timers.delete(id); };
    const fetchWrap = (u, o) => { calls.push({ u, o }); return fetchImpl(u, o); };
    new Function('window', 'document', 'localStorage', 'fetch', 'setInterval', 'clearInterval', src)(win, doc, localStorage, fetchWrap, fakeSetInterval, fakeClearInterval);
    // 在點擊前記錄看到的彈窗
    const seen = [];
    const origAppend = body.appendChild.bind(body);
    body.appendChild = c => { origAppend(c); seen.push(c); };
    const p = win.ITTSAnnounce.start({ username });
    // 輪流自動點「我知道了」
    let guard = 0;
    while (guard++ < 50) {
      await new Promise(r => setImmediate(r));
      const btn = body.all(e => e.tag === 'button')[0];
      if (btn && btn.disabled && timers.size) [...timers.values()].forEach(f => f()); // 倒數中：推進 1 秒
      else if (btn && btn.listeners.click && !btn.disabled) btn.listeners.click();
      else if (!btn) { const done = await Promise.race([p.then(() => true), new Promise(r => setImmediate(() => r(false)))]); if (done) break; }
    }
    await p;
    return { seen, calls, store, body, win, timers };
  }
  const json = (o, status = 200) => Promise.resolve({ ok: status < 400, status, json: () => Promise.resolve(o) });
  const item = (o) => ({ id: 'i1', version: 1, title: 'T', body: 'B', mode: 'once', severity: 'normal', ...o });

  // 強制閱讀倒數（DOM 層）
  {
    const { doc, body } = makeDom(); const win = { _myUsername: 'u' }; const tm = new Map(); let id = 0;
    new Function('window', 'document', 'localStorage', 'fetch', 'setInterval', 'clearInterval', src)(win, doc, {}, (u) => u.includes('/active') ? json({ loginAt: 1, items: [item({})] }) : json({ ok: true }), (fn) => { tm.set(++id, fn); return id; }, (i) => tm.delete(i));
    eq(win.ITTSAnnounce.READ_SECONDS, 5, '倒數秒數常數為 5');
    const p = win.ITTSAnnounce.start({ username: 'u' });
    await new Promise(r => setImmediate(r)); await new Promise(r => setImmediate(r));
    const b = body.all(e => e.tag === 'button')[0];
    eq([b.disabled, b.textContent], [true, '請先閱讀（5）'], '倒數：一開始 disabled 且顯示（5）');
    ok(body.all(e => e.tag === 'button').length === 1, '倒數：沒有關閉(X)按鈕');
    ok(Object.keys(body.children[0].listeners).length === 0, '倒數：背景無點擊/關閉處理器（點背景不會關）');
    b.listeners.click(); b.disabled = false; b.listeners.click(); // 即使 DOM 被竄改成 enabled，倒數未完仍不得送出
    await new Promise(r => setImmediate(r));
    ok(body.children.length === 1, '倒數：未完成時點擊（含竄改 disabled）不會關閉');
    const seq = [];
    b.disabled = true;
    for (let i = 0; i < 4; i++) { [...tm.values()].forEach(f => f()); seq.push(b.textContent); }
    eq(seq, ['請先閱讀（4）', '請先閱讀（3）', '請先閱讀（2）', '請先閱讀（1）'], '倒數：逐秒遞減 4→1');
    eq(b.disabled, true, '倒數：第 4 秒仍 disabled');
    [...tm.values()].forEach(f => f());
    eq([b.disabled, b.textContent, tm.size], [false, '我知道了', 0], '倒數：第 5 秒啟用、文字變「我知道了」、計時器清除');
    b.listeners.click(); await p;
    ok(body.children.length === 0, '倒數結束後可按下並關閉');
  }
  // XSS 惰性
  let r = await run((u, o) => u.includes('/active') ? json({ loginAt: 111, items: [item({ title: XSS, body: XSS })] }) : json({ ok: true }));
  eq(r.seen.length, 1, 'XSS：彈出 1 則');
  const all = []; r.seen[0].all(() => true, all);
  ok(all.every(e => !e.innerHTMLSet), 'XSS：元件從未使用 innerHTML');
  ok(all.every(e => e.tag !== 'img'), 'XSS：DOM 中沒有 <img> 元素');
  ok(all.some(e => e.textContent === XSS), 'XSS：內容以純文字 textContent 呈現（原字串）');
  ok(!/\.innerHTML\s*=|insertAdjacentHTML|document\.write|\beval\(/.test(src), 'XSS：原始碼不含 innerHTML=/insertAdjacentHTML/document.write/eval');
  eq(r.calls.filter(c => /\/ack$/.test(c.u)).length, 1, '按「我知道了」送出 1 次 ack');
  eq(JSON.parse(r.calls.find(c => /\/ack$/.test(c.u)).o.body).version, 1, 'ack 帶版本');

  // 失敗隔離
  r = await run(() => Promise.reject(new Error('net')));
  eq(r.seen.length, 0, '隔離：fetch 失敗 → 不彈、不拋錯');
  r = await run(() => json({}, 500));
  eq(r.seen.length, 0, '隔離：HTTP 500 → 不彈、不拋錯');
  r = await run(() => Promise.resolve({ ok: true, status: 200, json: () => Promise.reject(new Error('bad json')) }));
  eq(r.seen.length, 0, '隔離：回應非 JSON → 不彈、不拋錯');
  r = await run(() => json({ items: [null, { id: 'x' }, 5] }));
  eq(r.seen.length >= 0, true, '隔離：畸形項目不拋錯');
  // ack 失敗仍放行
  r = await run((u) => u.includes('/active') ? json({ loginAt: 1, items: [item({})] }) : Promise.reject(new Error('ack down')));
  eq(r.seen.length, 1, '隔離：ack 失敗仍關閉彈窗、流程不卡');

  // 順序：伺服器已排序，前端依序逐則
  r = await run((u) => u.includes('/active') ? json({ loginAt: 1, items: [item({ id: 'a', title: 'first' }), item({ id: 'b', title: 'second' })] }) : json({ ok: true }));
  eq(r.seen.length, 2, '多則：依序彈出 2 則');
  eq(r.seen.map(o => o.attrs['data-ann-id']), ['a', 'b'], '多則：順序與伺服器一致');

  // 每次登入：同 loginAt 不重彈（F5）；loginAt 變了（重新登入）重彈
  const ls = {};
  const ev = (la) => (u) => u.includes('/active') ? json({ loginAt: la, items: [item({ mode: 'everyLogin' })] }) : json({ ok: true });
  r = await run(ev(1000), { ls });
  eq(r.seen.length, 1, 'everyLogin：該次登入第一次載入 → 彈出');
  r = await run(ev(1000), { ls });
  eq(r.seen.length, 0, 'everyLogin：同一次登入再載入（F5）→ 不重彈');
  r = await run(ev(2000), { ls });
  eq(r.seen.length, 1, 'everyLogin：重新登入（loginAt 改變）→ 再彈');
  const ls2 = {};
  r = await run((u) => u.includes('/active') ? json({ loginAt: 5, items: [item({ mode: 'everyLogin', version: 1 })] }) : json({ ok: true }), { ls: ls2 });
  r = await run((u) => u.includes('/active') ? json({ loginAt: 5, items: [item({ mode: 'everyLogin', version: 2 })] }) : json({ ok: true }), { ls: ls2 });
  eq(r.seen.length, 1, 'everyLogin：同登入內版本升級 → 再彈');
  // once：前端不自行過濾（伺服器負責）；同一登入內 ack 後重載不重彈
  const ls3 = {};
  r = await run((u) => u.includes('/active') ? json({ loginAt: 9, items: [item({})] }) : json({ ok: true }), { ls: ls3 });
  r = await run((u) => u.includes('/active') ? json({ loginAt: 9, items: [item({})] }) : json({ ok: true }), { ls: ls3 });
  eq(r.seen.length, 0, 'once：伺服器回條遺失時，同次登入內也不重彈');
  // localStorage 不可用
  {
    const { doc, body } = makeDom();
    const win = { _myUsername: 'u' };
    const lt = new Map(); let lid = 0;
    const ls = { getItem() { throw new Error('blocked'); }, setItem() { throw new Error('blocked'); } };
    new Function('window', 'document', 'localStorage', 'fetch', 'setInterval', 'clearInterval', src)(win, doc, ls, (u) => u.includes('/active') ? json({ loginAt: 1, items: [item({})] }) : json({ ok: true }), (fn) => { lt.set(++lid, fn); return lid; }, (i) => lt.delete(i));
    const p = win.ITTSAnnounce.start({ username: 'u' });
    for (let i = 0; i < 40; i++) { await new Promise(r2 => setImmediate(r2)); const b = body.all(e => e.tag === 'button')[0]; if (b && b.disabled) [...lt.values()].forEach(f => f()); else if (b) b.listeners.click(); }
    await p;
    ok(true, 'localStorage 全面拋錯時不當機');
  }
  // 預覽不送 ack
  {
    const { doc, body } = makeDom();
    const win = {}; const calls = [];
    const pt = new Map(); let pid = 0;
    new Function('window', 'document', 'localStorage', 'fetch', 'setInterval', 'clearInterval', src)(win, doc, {}, (...a) => { calls.push(a); return json({}); }, (fn) => { pt.set(++pid, fn); return pid; }, (i) => pt.delete(i));
    const p = win.ITTSAnnounce.preview({ title: XSS, body: XSS, severity: 'important' });
    const b = body.all(e => e.tag === 'button')[0];
    eq([b.disabled, b.textContent], [true, '請先閱讀（5）'], '預覽：同樣有 5 秒倒數');
    b.listeners.click(); ok(body.children.length === 1, '預覽：倒數中點擊無效');
    for (let i = 0; i < 5; i++) [...pt.values()].forEach(f => f());
    eq([b.disabled, b.textContent], [false, '關閉預覽'], '預覽：倒數完成後可關閉');
    b.listeners.click(); await p;
    eq(calls.length, 0, '預覽：不送任何請求（不寫已讀）');
  }
}

// ───────────────────────── 端對端 ─────────────────────────
async function e2e() {
  const BASE = process.env.ANN_BASE.replace(/\/$/, '');
  const PW = process.env.ANN_PW;
  if (!PW) { console.log('ANN_PW 未設定，略過端對端'); return; }
  if (/vercel\.app|itts/i.test(BASE)) { console.log('拒絕：ANN_BASE 看起來不是本機沙盒'); process.exit(2); }
  console.log('[e2e]', BASE);
  const N = { admin: process.env.ANN_ADMIN || '_s_admin', user: process.env.ANN_USER || '_s_user', mgr: process.env.ANN_MGR || '_s_mgr', tpm: process.env.ANN_TPM || '_s_tpm', gs: process.env.ANN_GS || '_s_gs', view: process.env.ANN_VIEW || '_s_view' };
  const jar = {};
  async function req(who, method, url, body) {
    const h = { 'Content-Type': 'application/json' };
    if (who && jar[who]) h.Cookie = jar[who];
    for (let attempt = 0; attempt < 2; attempt++) {
      const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
      if (res.status === 429 && attempt === 0) { console.log('  (429，等 61 秒)'); await new Promise(r => setTimeout(r, 61000)); continue; }
      let j = null; const t = await res.text(); try { j = JSON.parse(t); } catch { /* not json */ }
      return { status: res.status, json: j, text: t, res };
    }
  }
  async function login(key) {
    const r = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: N[key], password: PW }) });
    if (!r.ok) throw new Error('登入失敗 ' + key + ' ' + r.status);
    jar[key] = r.headers.getSetCookie().map(c => c.split(';')[0]).join('; ');
    return r.json();
  }
  for (const k of Object.keys(N)) await login(k);

  const sleepMs = (n) => new Promise(r => setTimeout(r, n));
  // 合法 ack：先取得 active（記 shownAt），等 4.3 秒（強制閱讀），再送 ack
  const ackOk = async (who, id, v) => { await req(who, 'GET', '/api/announcements/active'); await sleepMs(4300); return req(who, 'POST', `/api/announcements/${id}/ack`, { version: v }); };
  const created = [];
  const adm = async (method, url, body) => req('admin', method, '/api/admin/announcements' + url, body);
  const mkAnn = async (o) => {
    const r = await adm('POST', '', { title: '[ANNCHK] ' + (o.t || 'x'), body: o.body || 'body', ...o.f });
    if (r.status === 200) created.push(r.json.id);
    return r;
  };
  const active = async (who) => (await req(who, 'GET', '/api/announcements/active')).json;
  const ids = (a) => (a.items || []).map(x => x.id);
  const tz = (ms) => lib.toTaipeiLocal(new Date(ms).toISOString());
  const H = 3600e3;

  try {
    // ── 權限 ──
    let r = await req(null, 'GET', '/api/announcements/active');
    eq(r.status, 401, '未登入 GET active → 401');
    r = await req(null, 'POST', '/api/announcements/x/ack', {});
    eq(r.status, 401, '未登入 ack → 401');
    for (const who of ['user', 'mgr', 'tpm', 'gs', 'view']) {
      r = await req(who, 'GET', '/api/admin/announcements');
      eq(r.status, 403, `${who} 讀管理列表 → 403`);
      r = await req(who, 'POST', '/api/admin/announcements', { title: 'x', body: 'y' });
      eq(r.status, 403, `${who} 新增公告 → 403`);
    }
    r = await adm('GET', '');
    eq(r.status, 200, 'admin 讀列表 200');
    ok(Array.isArray(r.json.roles) && r.json.roles.includes('groupsales'), '列表附帶角色清單');

    // ── CRUD + 驗證 ──
    r = await mkAnn({ t: 'basic', f: { mode: 'once' } });
    eq(r.status, 200, '新增公告 200');
    const basic = r.json.id;
    r = await adm('GET', '');
    const row = r.json.items.find(x => x.id === basic);
    ok(row && row.version === 1 && row.mode === 'once' && row.status === 'live', '列表可見、v1、live');
    ok(!('readBy' in row), '列表不含 readBy');
    eq([row.readCount, row.targetCount > 0], [0, true], '統計初始 0 / 目標>0');
    r = await adm('PUT', '/' + basic, { title: '[ANNCHK] basic2' });
    eq(r.status, 200, '更新 200');
    r = await adm('PUT', '/' + basic, { title: '' });
    eq(r.status, 400, '更新為空標題 → 400');
    r = await adm('PUT', '/nonexistent', { title: 'x' });
    eq(r.status, 404, '更新不存在 → 404');
    r = await adm('POST', '', { title: 'x'.repeat(101), body: 'b' }); eq(r.status, 400, '標題過長 400');
    r = await adm('POST', '', { title: 't', body: 'x'.repeat(4001) }); eq(r.status, 400, '內容過長 400');
    r = await adm('POST', '', { title: 't', body: 'b', roles: ['hacker'] }); eq(r.status, 400, '未知角色 400');
    r = await adm('POST', '', { title: 't', body: 'b', mode: 'x' }); eq(r.status, 400, '未知模式 400');
    r = await adm('POST', '', { title: 't', body: 'b', startAt: tz(Date.now() + 2 * H), endAt: tz(Date.now() + H) }); eq(r.status, 400, '結束早於開始 400');
    r = await adm('PUT', '/' + basic, { endAt: tz(Date.now() - 5 * H) }); // 開始為空→允許；結束在過去 → 已結束
    eq(r.status, 200, '只設結束時間（過去）接受');
    r = await adm('PUT', '/' + basic, { endAt: null }); eq(r.status, 200, '清除結束時間');
    r = await adm('DELETE', '/nonexistent'); eq(r.status, 404, '刪除不存在 404');

    // ── 時間窗 / 啟用 ──
    const fut = (await mkAnn({ t: 'future', f: { startAt: tz(Date.now() + 2 * H) } })).json.id;
    const past = (await mkAnn({ t: 'past', f: { startAt: tz(Date.now() - 3 * H), endAt: tz(Date.now() - 2 * H) } })).json.id;
    const win = (await mkAnn({ t: 'window', f: { startAt: tz(Date.now() - H), endAt: tz(Date.now() + H) } })).json.id;
    const dis = (await mkAnn({ t: 'disabled', f: { enabled: false } })).json.id;
    let a = await active('user');
    ok(ids(a).includes(win), '時間窗內出現');
    ok(!ids(a).includes(fut), '未開始不出現');
    ok(!ids(a).includes(past), '已結束不出現');
    ok(!ids(a).includes(dis), '停用不出現');
    ok(ids(a).includes(basic), '一般公告出現');
    r = await adm('GET', '');
    const stOf = id => r.json.items.find(x => x.id === id).status;
    eq([stOf(fut), stOf(past), stOf(win), stOf(dis)], ['scheduled', 'expired', 'live', 'disabled'], '管理列表狀態標示');
    eq(r.json.items.find(x => x.id === win).startLocal, tz(Date.now() - H).slice(0, 15) + r.json.items.find(x => x.id === win).startLocal.slice(15), '台北時間往返一致（至小時十位分）');
    r = await adm('PUT', '/' + dis, { enabled: true }); eq(r.status, 200, '啟用');
    ok(ids(await active('user')).includes(dis), '啟用後出現');
    await adm('PUT', '/' + dis, { enabled: false });
    ok(!ids(await active('user')).includes(dis), '停用後消失');
    eq((await adm('PUT', '/' + dis, { enabled: false })).json.unchanged, true, '無變更回報 unchanged');

    // ── 角色鎖定 ──
    const mgrOnly = (await mkAnn({ t: 'mgronly', f: { roles: ['manager1'] } })).json.id;
    ok(!ids(await active('user')).includes(mgrOnly), 'user 看不到 manager1 專屬');
    ok(ids(await active('mgr')).includes(mgrOnly), 'manager1 看得到');
    r = await req('user', 'POST', `/api/announcements/${mgrOnly}/ack`, {});
    eq(r.status, 404, 'user 對非目標公告 ack → 404');
    r = await req('user', 'POST', `/api/announcements/nonexistent/ack`, {});
    eq(r.status, 404, 'ack 不存在 → 404');

    // ── once 語意 ──
    const once = (await mkAnn({ t: 'once', f: { mode: 'once', priority: 1 } })).json.id;
    ok(ids(await active('user')).includes(once), 'once：未確認出現');
    // 強制閱讀：伺服器端防繞過（送達後不足 4 秒的 ack 一律拒絕）
    r = await req('user', 'POST', `/api/announcements/${once}/ack`, { version: 1 });
    eq([r.status, (r.json || {}).code], [409, 'TOO_EARLY'], '強制閱讀：送達後立即 ack → 409 TOO_EARLY');
    ok(ids(await active('user')).includes(once), '強制閱讀：被拒的 ack 不記已讀（公告仍會出現）');
    { const late = (await mkAnn({ t: 'neverShown' })).json.id; // 建立後從未被 GET active 送達
      r = await req('user', 'POST', `/api/announcements/${late}/ack`, { version: 1 });
      eq(r.status, 409, '強制閱讀：從未送達（無 shownAt）的 ack → 409');
      await adm('DELETE', '/' + late); created.splice(created.indexOf(late), 1); }
    await sleepMs(4300);
    r = await req('user', 'POST', `/api/announcements/${once}/ack`, { version: 1 });
    eq(r.status, 200, 'once：ack 200（送達 4.3 秒後放行）');
    ok(!ids(await active('user')).includes(once), 'once：確認後不再出現');
    ok(ids(await active('mgr')).includes(once), 'once：他人不受影響');
    r = await req('user', 'POST', `/api/announcements/${once}/ack`, { version: 1 });
    eq(r.status, 200, 'once：重複 ack 冪等');
    r = await adm('GET', '');
    let st = r.json.items.find(x => x.id === once);
    eq(st.readCount, 1, 'once：已讀人數 1');
    // 版本升級
    r = await adm('PUT', '/' + once, { body: '內容更新' });
    eq(r.json.version, 1, '改內文不升版');
    ok(!ids(await active('user')).includes(once), '改內文不重新公告');
    r = await adm('PUT', '/' + once, { body: '重大更新', renotify: true });
    eq(r.json.version, 2, '重新公告 → 版本 2');
    a = await active('user');
    ok(ids(a).includes(once), '重新公告後已讀者再看到');
    eq(a.items.find(x => x.id === once).version, 2, '回傳新版本號');
    r = await adm('GET', ''); eq(r.json.items.find(x => x.id === once).readCount, 0, '升版後已讀歸零');
    r = await ackOk("user", once, 1); // 舊版彈窗的確認
    ok(ids(await active('user')).includes(once), '舊版 ack 不吃掉新版');
    r = await ackOk("user", once, 2);
    ok(!ids(await active('user')).includes(once), '新版 ack 後不再出現');
    r = await req('user', 'POST', `/api/announcements/${once}/ack`, { version: 99 });
    ok(!ids(await active('user')).includes(once), '超大版本號被夾到目前版本');

    // ── everyLogin 語意 ──
    const el = (await mkAnn({ t: 'every', f: { mode: 'everyLogin', priority: 2 } })).json.id;
    await ackOk("user", el, 1);
    ok(ids(await active('user')).includes(el), 'everyLogin：確認後伺服器仍回傳（由前端依 loginAt 判斷）');
    r = await adm('GET', ''); eq(r.json.items.find(x => x.id === el).readCount, 1, 'everyLogin：統計計入確認過的人');
    const la1 = (await active('user')).loginAt;
    ok(Number.isFinite(la1) && la1 > 1e12, 'active 回傳 loginAt（毫秒戳）');
    const la1b = (await active('user')).loginAt;
    eq(la1, la1b, '同一 session 重複請求 loginAt 不變（F5 不算新登入）');
    await new Promise(rr => setTimeout(rr, 20));
    await login('user');
    const la2 = (await active('user')).loginAt;
    ok(la2 > la1, '重新登入 → loginAt 變大');
    const me = (await req('user', 'GET', '/api/me')).json;
    eq(me.loginAt, la2, '/api/me 也帶同一個 loginAt（前端可用）');

    // ── 排序 ──
    a = await active('mgr');
    const order = ids(a).filter(i => [once, el].includes(i));
    eq(order.indexOf(once) < order.indexOf(el) || !order.includes(once), true, '依 priority 排序（小者先）');

    // ── 隱私：readBy 不外洩 ──
    const raw = (await req('user', 'GET', '/api/announcements/active')).text;
    ok(!raw.includes('readBy') && !raw.includes(N.mgr), 'active 回應不含 readBy / 他人帳號');
    ok(!raw.includes('createdBy'), 'active 回應不含管理欄位');

    // ── XSS：存入與回傳皆為原字串（JSON 字串），顯示端用 textContent ──
    const x = await mkAnn({ t: XSS, body: XSS + '\n<script>alert(2)</script>', f: { priority: 0 } });
    eq(x.status, 200, 'XSS 公告可儲存（純文字）');
    a = await active('user');
    const xi = a.items.find(i => i.id === x.json.id);
    eq(xi.title, '[ANNCHK] ' + XSS, 'XSS：active 回傳為原字串');
    ok(/application\/json/.test((await req('user', 'GET', '/api/announcements/active')).res.headers.get('content-type')), 'active 以 application/json 回傳（不會被當 HTML 解析）');
    ok(((await req('user', 'GET', '/api/announcements/active')).res.headers.get('cache-control') || '').includes('no-store'), 'active 不可快取（no-store）');

    // ── 集團角色白名單 / 唯讀模式 ──
    const allRole = (await mkAnn({ t: 'forall', f: { priority: 3 } })).json.id;
    for (const who of ['tpm', 'gs']) {
      r = await req(who, 'GET', '/api/announcements/active');
      eq(r.status, 200, `${who} 可讀 active`);
      ok(ids(r.json).includes(allRole), `${who} 看得到全員公告`);
      r = await ackOk(who, allRole, 1);
      eq(r.status, 200, `${who} 可 ack`);
      ok(!ids(await active(who)).includes(allRole), `${who} ack 後不再出現`);
      r = await req(who, 'GET', '/api/admin/announcements'); eq(r.status, 403, `${who} 仍不可進後台端點`);
    }
    r = await ackOk("view", allRole, 1);
    eq(r.status, 200, '唯讀模式帳號可 ack');
    ok(!ids(await active('view')).includes(allRole), '唯讀模式 ack 後不再出現');
    r = await req('view', 'POST', '/api/admin/announcements', { title: 'x', body: 'y' });
    eq(r.status, 403, '唯讀帳號仍不可建立');

    // ── 刪除 ──
    r = await adm('DELETE', '/' + allRole); eq(r.status, 200, '刪除 200');
    created.splice(created.indexOf(allRole), 1);
    ok(!ids(await active('user')).includes(allRole), '刪除後消失');
    r = await adm('GET', ''); ok(!r.json.items.some(i => i.id === allRole), '刪除後列表無');

    // ── 稽核 ──
    // （writeLog 的檢查：伺服器 audit.log.json 在沙盒目錄；此處用 /api/admin/logs 若可用）
    const logs = await req('admin', 'GET', '/api/admin/logs');
    if (logs.status === 200 && Array.isArray(logs.json.logs || logs.json)) {
      const arr = logs.json.logs || logs.json;
      const acts = new Set(arr.map(l => l.action));
      ok(acts.has('CREATE_ANNOUNCEMENT') && acts.has('UPDATE_ANNOUNCEMENT') && acts.has('DELETE_ANNOUNCEMENT'), '稽核：CREATE/UPDATE/DELETE_ANNOUNCEMENT 均有 writeLog');
    } else console.log('  (略過稽核檢查：/api/admin/logs 格式不符，改由手動檢查 audit.log.json)');

    // ── 上限 200 ──
    r = await adm('GET', '');
    const have = r.json.items.length;
    let made = 0;
    for (let i = have; i < 200; i++) { const q = await mkAnn({ t: 'bulk' + i }); if (q.status !== 200) break; made++; }
    r = await mkAnn({ t: 'over' });
    eq(r.status, 400, '第 201 則被拒絕（上限 200）');
    ok(/200/.test((r.json || {}).error || ''), '上限訊息含數字');
  } finally {
    // 只刪自己建的 [ANNCHK]
    const list = await adm('GET', '');
    const mine = ((list.json && list.json.items) || []).filter(i => String(i.title).startsWith('[ANNCHK]'));
    for (const i of mine) await adm('DELETE', '/' + i.id);
    const after = await adm('GET', '');
    ok(!((after.json && after.json.items) || []).some(i => String(i.title).startsWith('[ANNCHK]')), '清理：測試公告已全數刪除');
  }
}

(async () => {
  unitLib();
  await unitClient();
  if (process.env.ANN_BASE) await e2e(); else console.log('⚠ 未驗證：未設 ANN_BASE，端對端（路由／權限／稽核）沒有執行；請自行起伺服器並設 ANN_BASE、ANN_PW');
  console.log(`\n結果：通過 ${pass}，失敗 ${fail}` + (process.env.ANN_BASE ? '' : '（端對端未執行＝未驗證）'));
  if (fail) { console.log('失敗項目：\n - ' + failures.join('\n - ')); process.exit(1); }
})().catch(e => { console.error('腳本錯誤：', e); process.exit(1); });
