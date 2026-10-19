#!/usr/bin/env node
/**
 * 商機動態報表「預計簽約日調整」反覆延後警示檢查。用法：node scripts/check-pipeline-date-warning.js
 * 動 lib/pipelineDateWarning.js、server.js 的 /api/pipeline-report、_client/app.js 的 getPrDateRange／renderPrMoves、
 * index.html 的 prDateChangeSection、style.css／glass.css 的 .pr-dc-* 之後必跑。
 *
 *  1) 單元：累計次數／天數（跨期、忽略建立紀錄、來回、多次）、停留天數、repeated 閾值邊界（次數 2、天數 30、停留 45 的 ±1）、
 *     Won／D／A 不標、環境變數覆寫、排序、台北月界
 *  2) API（會啟動真的 server.js，埠 3095，臨時 _s_* 帳號與資料檔；測完刪除）：同一組案例走完整路由；
 *     與基底（merge-base 的 server.js）逐欄比對既有欄位不變；各角色可視範圍不變；預設區間；環境變數覆寫
 *  3) 前端：getPrDateRange 自訂區間起點（台北 0–8 點）、renderPrMoves 輸出文字／徽章／展開鈕、
 *     日期區塊不得有寫死 inline color、CSS class 都有定義、區塊順序
 *  4) 變異測試：故意改壞實作（>= 改 >、少算建立紀錄、改用本期而非全歷史……），對應測試必須變紅
 *
 * 安全：API 段需要在 repo 根目錄寫 data.json／auth.json——若這兩個檔已存在（真實資料）就拒絕執行該段，絕不覆蓋。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const vm = require('vm');
const { spawn, execFileSync, spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = 3095;
const DAY = 86400000;
const LIB_FILE = path.join(ROOT, 'lib', 'pipelineDateWarning.js');
const libSrc = fs.readFileSync(LIB_FILE, 'utf8');
const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
const appSrc = fs.readFileSync(path.join(ROOT, '_client', 'app.js'), 'utf8');

let pass = 0, fail = 0;
function mkAsserter(sink) {
  return (cond, name, extra) => {
    if (cond) { if (!sink) { pass++; } }
    else { if (sink) sink.push(name); else { fail++; console.log('  FAIL ' + name + (extra !== undefined ? ' -> ' + JSON.stringify(extra) : '')); } }
  };
}
function section(t) { console.log('\n== ' + t); }
const sleep = ms => new Promise(r => setTimeout(r, ms));

function loadLib(src) {
  const m = { exports: {} };
  new Function('module', 'exports', 'process', src)(m, m.exports, process);
  return m.exports;
}

// ── 共用案例表（單元與 API 共用）────────────────────────────────
const addDays = (ymd, n) => new Date(Date.parse(ymd) + n * DAY).toISOString().slice(0, 10);
const B0 = '2027-03-01';
const rec = (old, nw, agoDays) => ({ oldDate: old, newDate: nw, agoDays });
// since = 停在目前階段的起點（距今天數）；hist 用 stageHistory，created 用 createdAt（無 hist 時的後備）
const CASES = [
  { id: 'r1', stage: 'B', since: 60, recs: [rec(B0, addDays(B0, 10), 100), rec(addDays(B0, 10), addDays(B0, 20), 70), rec(addDays(B0, 20), addDays(B0, 30), 5)],
    exp: { count: 3, days: 30, stay: 60, rep: true }, note: '跨期累計（2 筆在期外）' },
  { id: 'r2', stage: 'B', since: 60, recs: [rec(null, B0, 90), rec(B0, addDays(B0, 20), 3)],
    exp: { count: 1, days: 20, stay: 60, rep: false }, note: '建立紀錄(oldDate=null)被忽略' },
  { id: 'r3', stage: 'B', since: 60, recs: [rec(B0, addDays(B0, 10), 100), rec(addDays(B0, 10), B0, 90), rec(B0, addDays(B0, 5), 3)],
    exp: { count: 2, days: 5, stay: 60, rep: true }, note: '來回：+10、-10、+5 → 次數 2、淨 5 天' },
  { id: 't1', stage: 'B', since: 46, recs: [rec(B0, addDays(B0, 30), 3)], exp: { count: 1, days: 30, stay: 46, rep: true }, note: '天數=30 且停留 46' },
  { id: 't2', stage: 'B', since: 46, recs: [rec(B0, addDays(B0, 29), 3)], exp: { count: 1, days: 29, stay: 46, rep: false }, note: '天數=29 不標' },
  { id: 't3', stage: 'C', since: 46, recs: [rec(B0, addDays(B0, 1), 4), rec(addDays(B0, 1), addDays(B0, 2), 2)], exp: { count: 2, days: 2, stay: 46, rep: true }, note: '次數=2 且停留 46（C 階）' },
  { id: 't4', stage: 'B', since: 45, recs: [rec(B0, addDays(B0, 1), 4), rec(addDays(B0, 1), addDays(B0, 2), 2)], exp: { count: 2, days: 2, stay: 45, rep: false }, note: '停留=45 不標（需 >45）' },
  { id: 't5', stage: 'B', since: 45, recs: [rec(B0, addDays(B0, 30), 3)], exp: { count: 1, days: 30, stay: 45, rep: false }, note: '天數 30 但停留 45 不標' },
  { id: 't6', stage: 'B', since: 60, recs: [rec(B0, addDays(B0, 5), 3)], exp: { count: 1, days: 5, stay: 60, rep: false }, note: '次數 1、天數 5 不標' },
  { id: 'a1', stage: 'A', since: 90, recs: [rec(B0, addDays(B0, 20), 100), rec(addDays(B0, 20), addDays(B0, 50), 3)], exp: { count: 2, days: 50, stay: 90, rep: false }, note: 'A 階不標' },
  { id: 'w1', stage: 'Won', since: 90, recs: [rec(B0, addDays(B0, 20), 100), rec(addDays(B0, 20), addDays(B0, 50), 3)], exp: { count: 2, days: 50, stay: 90, rep: false }, note: 'Won 不標' },
  { id: 'd1', stage: 'D', since: 90, recs: [rec(B0, addDays(B0, 20), 100), rec(addDays(B0, 20), addDays(B0, 50), 3)], exp: { count: 2, days: 50, stay: 90, rep: false }, note: 'D 不標' },
  { id: 'c1', stage: 'C', created: 50, recs: [rec(B0, addDays(B0, 10), 100), rec(addDays(B0, 10), addDays(B0, 20), 3)], exp: { count: 2, days: 20, stay: 50, rep: true }, note: '無 stageHistory → 用 createdAt' },
  { id: 'h1', stage: 'B', hist: [['B', 100], ['C', 60], ['B', 10]], created: 200, recs: [rec(B0, addDays(B0, 10), 100), rec(addDays(B0, 10), addDays(B0, 40), 3)], exp: { count: 2, days: 40, stay: 10, rep: false }, note: '用最後一次 to=B（10 天），不是 createdAt 或第一次' },
  { id: 'h2', stage: 'B', hist: [['B', 80], ['C', 70]], created: 200, since: 0, recs: [rec(B0, addDays(B0, 10), 3)], exp: { count: 1, days: 10, stay: 80, rep: false }, note: '歷史中最後一次 to=B 是 80 天前（後來又退到別的階段再退回不算）' },
];
function oppFor(c, nowMs) {
  const o = { id: c.id, stage: c.stage };
  if (c.hist) o.stageHistory = c.hist.map(([to, ago]) => ({ from: 'X', to, date: new Date(nowMs - ago * DAY).toISOString() }));
  else if (c.since != null) o.stageHistory = [{ from: 'X', to: c.stage, date: new Date(nowMs - c.since * DAY).toISOString() }];
  if (c.created != null) o.createdAt = new Date(nowMs - c.created * DAY).toISOString();
  else if (!o.stageHistory) o.createdAt = new Date(nowMs - 30 * DAY).toISOString();
  return o;
}
const recsFor = (c, nowMs) => c.recs.map(r => ({ dealId: c.id, oldDate: r.oldDate, newDate: r.newDate, changedAt: new Date(nowMs - r.agoDays * DAY).toISOString(), owner: '_s_sales' }));
// h2 的 hist 不含目前階段以外…（h2：最後一次 to=B 為 80 天前）

// ═════════ 1) 單元 ═════════
function unitSuite(lib, assert) {
  const nowMs = Date.now(), now = new Date(nowMs);
  const saved = {};
  ['PIPELINE_REPEAT_MIN_COUNT', 'PIPELINE_REPEAT_MIN_DAYS', 'PIPELINE_REPEAT_STALL_DAYS'].forEach(k => { saved[k] = process.env[k]; delete process.env[k]; });
  try {
    for (const c of CASES) {
      const r = lib.evaluate(oppFor(c, nowMs), recsFor(c, nowMs), now);
      assert(r.postponeCount === c.exp.count, `[${c.id}] 次數 ${c.note}`, r);
      assert(r.postponeDays === c.exp.days, `[${c.id}] 天數 ${c.note}`, r);
      assert(r.stageDays === c.exp.stay, `[${c.id}] 停留 ${c.note}`, r);
      assert(r.repeated === c.exp.rep, `[${c.id}] repeated ${c.note}`, r);
    }
    // 空紀錄／亂序輸入／無法解析
    const e = lib.evaluate({ stage: 'B', createdAt: new Date(nowMs - 90 * DAY).toISOString() }, [], now);
    assert(e.postponeCount === 0 && e.postponeDays === 0 && e.repeated === false, '無紀錄 → 0/0/false', e);
    const shuffled = recsFor(CASES[0], nowMs).reverse();
    const s = lib.evaluate(oppFor(CASES[0], nowMs), shuffled, now);
    assert(s.postponeCount === 3 && s.postponeDays === 30, '輸入亂序仍依 changedAt 升冪計算', s);
    const noDate = lib.evaluate({ stage: 'B' }, recsFor(CASES[0], nowMs), now);
    assert(noDate.stageDays === null && noDate.repeated === false, '無階段起點資訊 → stageDays=null 且不標', noDate);
    // 提前不扣次數、最終淨延後為負 → 天數 0
    const adv = lib.cumulativePostpone([{ oldDate: B0, newDate: addDays(B0, 10), changedAt: '2026-01-01T00:00:00Z' }, { oldDate: addDays(B0, 10), newDate: addDays(B0, -5), changedAt: '2026-01-02T00:00:00Z' }]);
    assert(adv.count === 1 && adv.days === 0, '先延後再大幅提前：次數 1、淨天數夾 0', adv);
    // 環境變數覆寫
    process.env.PIPELINE_REPEAT_MIN_COUNT = '3';
    const t3 = CASES.find(c => c.id === 't3');
    assert(lib.evaluate(oppFor(t3, nowMs), recsFor(t3, nowMs), now).repeated === false, 'env MIN_COUNT=3 → 次數 2 不再標', lib.getRules());
    delete process.env.PIPELINE_REPEAT_MIN_COUNT;
    process.env.PIPELINE_REPEAT_MIN_DAYS = '50';
    const t1 = CASES.find(c => c.id === 't1');
    assert(lib.evaluate(oppFor(t1, nowMs), recsFor(t1, nowMs), now).repeated === false, 'env MIN_DAYS=50 → 天數 30 不再標', lib.getRules());
    delete process.env.PIPELINE_REPEAT_MIN_DAYS;
    process.env.PIPELINE_REPEAT_STALL_DAYS = '44';
    const t4 = CASES.find(c => c.id === 't4');
    assert(lib.evaluate(oppFor(t4, nowMs), recsFor(t4, nowMs), now).repeated === true, 'env STALL_DAYS=44 → 停留 45 變成標記', lib.getRules());
    process.env.PIPELINE_REPEAT_STALL_DAYS = 'abc';
    assert(lib.getRules().stallDays === 45, 'env 非數字 → 回預設 45', lib.getRules());
    delete process.env.PIPELINE_REPEAT_STALL_DAYS;
    const rules = lib.getRules();
    assert(rules.minCount === 2 && rules.minDays === 30 && rules.stallDays === 45, '預設值 2／30／45', rules);
    // 排序
    const list = [
      { id: 'n1', repeated: false, postponeDays: 90, daysDiff: 90, changedAt: '2026-01-01' },
      { id: 'r_small', repeated: true, postponeDays: 5, daysDiff: 5, changedAt: '2026-01-01' },
      { id: 'r_big', repeated: true, postponeDays: 40, daysDiff: 1, changedAt: '2026-01-01' },
      { id: 'n2', repeated: false, postponeDays: 90, daysDiff: 10, changedAt: '2026-01-01' },
      { id: 'adv', repeated: false, postponeDays: 0, daysDiff: -3, changedAt: '2026-01-02' },
    ].sort(lib.compareEntries).map(x => x.id).join(',');
    assert(list === 'r_big,r_small,n1,n2,adv', '排序：repeated → 累計天數 → 本期天數', list);
    // 台北月界（不受 server 時區影響）
    const iso = d => d.toISOString();
    let tm = lib.taipeiMonthRange(new Date('2026-10-31T17:00:00Z')); // 台北 11/1 01:00
    assert(iso(tm.from) === '2026-10-31T16:00:00.000Z' && iso(tm.to) === '2026-11-30T15:59:59.000Z', '台北 11/1 01:00 → 11 月', [iso(tm.from), iso(tm.to)]);
    tm = lib.taipeiMonthRange(new Date('2026-10-31T15:00:00Z')); // 台北 10/31 23:00
    assert(iso(tm.from) === '2026-09-30T16:00:00.000Z' && iso(tm.to) === '2026-10-31T15:59:59.000Z', '台北 10/31 23:00 → 10 月', [iso(tm.from), iso(tm.to)]);
    tm = lib.taipeiMonthRange(new Date('2028-02-10T00:00:00Z'));
    assert(iso(tm.to) === '2028-02-29T15:59:59.000Z', '閏年 2 月底', iso(tm.to));
    tm = lib.taipeiMonthRange(new Date('2026-12-31T20:00:00Z')); // 台北 2027/1/1 04:00
    assert(iso(tm.from) === '2026-12-31T16:00:00.000Z' && iso(tm.to) === '2027-01-31T15:59:59.000Z', '跨年：台北 1/1 → 1 月', [iso(tm.from), iso(tm.to)]);
  } finally {
    Object.entries(saved).forEach(([k, v]) => { if (v === undefined) delete process.env[k]; else process.env[k] = v; });
  }
}

// ═════════ 2) API ═════════
const DATA_FILE = path.join(ROOT, 'data.json');
const AUTH_FILE = path.join(ROOT, 'auth.json');
// server 啟動後會順手寫出稽核檔（audit.log.json 等）；只清掉「本腳本啟動前不存在」的，絕不動既有檔案
const SIDE_FILES = ['audit.log.json', 'contact_audit.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'server.log'].map(f => path.join(ROOT, f));
const PRE_EXISTING = new Set(SIDE_FILES.filter(f => fs.existsSync(f)));
function cleanFiles() {
  [DATA_FILE, AUTH_FILE].forEach(f => { try { fs.unlinkSync(f); } catch {} });
  SIDE_FILES.forEach(f => { if (!PRE_EXISTING.has(f)) { try { fs.unlinkSync(f); } catch {} } });
}
const PW = 'Sbx-' + crypto.randomBytes(9).toString('hex') + 'Aa1!';

function seed(nowMs) {
  const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
  const pw = bcrypt.hashSync(PW, 10);
  const iso = new Date(nowMs).toISOString();
  const mk = (username, role, bu, sup) => ({ username, displayName: username, role, bu: [bu], supervisor: sup, active: true, password: pw, passwordChangedAt: iso, createdAt: iso, accessMode: 'edit' });
  fs.writeFileSync(AUTH_FILE, JSON.stringify({ users: [
    mk('_s_admin', 'admin', 'ERP'), mk('_s_mgr', 'manager1', 'ERP'), mk('_s_sales', 'user', 'ERP', '_s_mgr'),
    mk('_s_sec', 'secretary', 'ERP'), mk('_s_mgr2', 'manager1', 'MDM'), mk('_s_sales2', 'user', 'MDM', '_s_mgr2'),
  ] }, null, 2));
  const opps = [], changes = [];
  const base = o => ({ owner: '_s_sales', bu: 'ERP', company: 'Co-' + o.id, contactName: 'x', category: 'ERP', product: 'P-' + o.id, amount: '1000', expectedDate: B0, description: '' });
  for (const c of CASES) {
    const o = Object.assign(base(c), oppFor(c, nowMs));
    opps.push(o); changes.push(...recsFor(c, nowMs));
  }
  const x = Object.assign(base({ id: 'x_other_bu' }), { id: 'x_other_bu', owner: '_s_sales2', bu: 'MDM', stage: 'B', createdAt: new Date(nowMs - 90 * DAY).toISOString() });
  opps.push(x);
  changes.push({ dealId: 'x_other_bu', oldDate: B0, newDate: addDays(B0, 40), changedAt: new Date(nowMs - 2 * DAY).toISOString(), owner: '_s_sales2' });
  fs.writeFileSync(DATA_FILE, JSON.stringify({ contacts: [], visits: [], opportunities: opps, lostOpportunities: [], opportunityDateChanges: changes }, null, 2));
}

async function startServer(entry, env) {
  const child = spawn(process.execPath, [entry], { cwd: ROOT, env: Object.assign({}, process.env, { PORT: String(PORT), SESSION_SECRET: crypto.randomBytes(24).toString('hex') }, env || {}), stdio: ['ignore', 'pipe', 'pipe'] });
  let log = '';
  child.stdout.on('data', d => log += d); child.stderr.on('data', d => log += d);
  for (let i = 0; i < 60; i++) {
    await sleep(500);
    try { const r = await fetch(`http://localhost:${PORT}/api/login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' }); if (r.status) break; } catch {}
    if (child.exitCode != null) throw new Error('server 提早結束：' + log.slice(-400));
  }
  return { child, stop: async () => {
    const pid = child.pid;
    try { if (process.platform === 'win32') spawnSync('taskkill', ['/PID', String(pid), '/T', '/F']); else child.kill('SIGKILL'); } catch {}
    await sleep(500);
  } };
}
async function login(u) {
  const r = await fetch(`http://localhost:${PORT}/api/login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: u, password: PW }) });
  const j = await r.json();
  if (!j.success) throw new Error('login ' + u + ' ' + JSON.stringify(j));
  return r.headers.getSetCookie().map(x => x.split(';')[0]).join('; ');
}
async function report(cookie, qs) {
  const r = await fetch(`http://localhost:${PORT}/api/pipeline-report${qs ? '?' + qs : ''}`, { headers: { Cookie: cookie } });
  return { status: r.status, j: await r.json() };
}
const USERS = ['_s_admin', '_s_mgr', '_s_sales', '_s_sec', '_s_mgr2', '_s_sales2'];
async function collect(entry, env, nowMs) {
  seed(nowMs);
  const srv = await startServer(entry, env);
  try {
    const from = new Date(nowMs - 60 * DAY).toISOString(), to = new Date(nowMs + 1 * DAY).toISOString();
    const out = {};
    for (const u of USERS) out[u] = (await report(await login(u), `from=${from}&to=${to}`)).j;
    out.defaultWindow = (await report(await login('_s_admin'), '')).j;
    return out;
  } finally { await srv.stop(); cleanFiles(); }
}

function apiAssertions(res, assert, nowMs) {
  const adm = res._s_admin;
  const byId = Object.fromEntries((adm.dateChanges || []).map(d => [d.id, d]));
  for (const c of CASES) {
    const d = byId[c.id];
    assert(!!d, `[API ${c.id}] 出現在清單`);
    if (!d) continue;
    assert(d.postponeCount === c.exp.count, `[API ${c.id}] postponeCount ${c.note}`, d.postponeCount);
    assert(d.postponeDays === c.exp.days, `[API ${c.id}] postponeDays ${c.note}`, d.postponeDays);
    assert(d.stageDays === c.exp.stay, `[API ${c.id}] stageDays ${c.note}`, d.stageDays);
    assert(d.repeated === c.exp.rep, `[API ${c.id}] repeated ${c.note}`, d);
  }
  assert(adm.summary.repeatedCount === CASES.filter(c => c.exp.rep).length + 1, '[API] summary.repeatedCount（案例 + x_other_bu）', adm.summary.repeatedCount);
  assert(adm.repeatRule && adm.repeatRule.minCount === 2 && adm.repeatRule.minDays === 30 && adm.repeatRule.stallDays === 45, '[API] repeatRule 回傳預設值', adm.repeatRule);
  // 排序（用獨立寫的比較子驗）
  const order = adm.dateChanges;
  let sorted = true;
  for (let i = 1; i < order.length; i++) {
    const a = order[i - 1], b = order[i];
    if (!!a.repeated !== !!b.repeated) { if (!a.repeated) sorted = false; continue; }
    if (a.postponeDays !== b.postponeDays) { if (a.postponeDays < b.postponeDays) sorted = false; continue; }
    if ((a.daysDiff ?? -1) < (b.daysDiff ?? -1)) sorted = false;
  }
  assert(sorted && order.length && order[0].repeated === true, '[API] 清單排序：repeated 優先、再累計天數、再本期天數', order.map(o => [o.id, o.repeated, o.postponeDays, o.daysDiff]));
  const w = new Date(res.defaultWindow.period.from).getTime();
  const tm = loadLib(libSrc).taipeiMonthRange(new Date(nowMs));
  assert(Math.abs(w - tm.from.getTime()) < 1 && res.defaultWindow.period.to === tm.to.toISOString(), '[API] 未帶 from/to → 台北當月', res.defaultWindow.period);
  // 可視範圍：其他 BU 業務只看到自己的
  const ids = u => (res[u].dateChanges || []).map(d => d.id).sort().join(',');
  assert(ids('_s_sales2') === 'x_other_bu', '[API] 其他 BU 業務只看到自己的', ids('_s_sales2'));
  assert(!(res._s_sales.dateChanges || []).some(d => d.id === 'x_other_bu'), '[API] 業務看不到別 BU');
}
const NEW_KEYS = ['postponeCount', 'postponeDays', 'stageDays', 'repeated'];
function compareToBase(res, base, assert) {
  for (const u of USERS) {
    const a = res[u], b = base[u];
    const strip = d => { const x = Object.assign({}, d); NEW_KEYS.forEach(k => delete x[k]); return x; };
    const sortId = arr => arr.slice().sort((p, q) => p.id < q.id ? -1 : 1);
    assert(JSON.stringify(sortId(a.dateChanges.map(strip))) === JSON.stringify(sortId(b.dateChanges)), `[基底比對 ${u}] dateChanges 既有欄位逐欄相同（不含排序）`);
    assert(a.dateChanges.map(d => d.id).sort().join() === b.dateChanges.map(d => d.id).sort().join(), `[基底比對 ${u}] 可視範圍／列出的商機相同`);
    for (const k of ['funnel', 'newDeals', 'lostDeals', 'promoted', 'demoted', 'amountChanges', 'period', 'ownerOptions']) {
      assert(JSON.stringify(a[k]) === JSON.stringify(b[k]), `[基底比對 ${u}] ${k} 不變`);
    }
    const sa = Object.assign({}, a.summary); delete sa.repeatedCount;
    assert(JSON.stringify(sa) === JSON.stringify(b.summary), `[基底比對 ${u}] summary 既有欄位不變`, [sa, b.summary]);
  }
}
function baseServerSrc() {
  try {
    const mb = execFileSync('git', ['merge-base', 'HEAD', 'origin/main'], { cwd: ROOT, encoding: 'utf8' }).trim();
    return execFileSync('git', ['show', mb + ':server.js'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  } catch (e) { return null; }
}

// ═════════ 3) 前端 ═════════
function sliceBetween(src, a, b) {
  const i = src.indexOf(a), j = src.indexOf(b, i + 1);
  if (i < 0 || j < 0) throw new Error('找不到區段 ' + a);
  return src.slice(i, j);
}
function runRange(src, tz, f, t) {
  const fn = sliceBetween(src, 'function getPrDateRange', 'function fmtPeriodLabel');
  const code = `${fn}\nconst $ = id => ({ value: id === 'prFromDate' ? ${JSON.stringify(f)} : ${JSON.stringify(t)} });\nconst r = getPrDateRange('custom');\nconsole.log(JSON.stringify({ from: r.from.toISOString(), to: r.to.toISOString(), fh: r.from.getHours(), fd: r.from.getDate() }));`;
  const out = execFileSync(process.execPath, ['-e', code], { env: Object.assign({}, process.env, { TZ: tz }), encoding: 'utf8' });
  return JSON.parse(out.trim());
}
function frontendSuite(src, assert) {
  const tp = runRange(src, 'Asia/Taipei', '2026-10-01', '2026-10-31');
  assert(tp.from === '2026-09-30T16:00:00.000Z', '[前端] 自訂區間起點 = 台北 10/1 00:00（不是 08:00）', tp);
  assert(tp.to === '2026-10-31T15:59:59.000Z', '[前端] 自訂區間終點 = 台北 10/31 23:59:59', tp);
  const la = runRange(src, 'America/Los_Angeles', '2026-10-01', '2026-10-31');
  assert(la.fh === 0 && la.fd === 1, '[前端] 其他時區：起點為本地 0 點', la);
  // 「台北 10/1 03:00」的異動必須落在自訂 10/1 起的區間
  assert(Date.parse('2026-09-30T19:00:00Z') >= Date.parse(tp.from), '[前端] 台北 10/1 03:00 的異動在區間內');
}
function renderSuite(src, assert) {
  const fn = sliceBetween(src, 'function renderPrMoves', '// ── 合約資料載入');
  const els = {};
  const $ = id => els[id] || (els[id] = { id, innerHTML: '', textContent: '', style: {}, classList: { _s: new Set(), remove(c) { this._s.delete(c); }, toggle(c) { if (this._s.has(c)) { this._s.delete(c); return false; } this._s.add(c); return true; } }, setAttribute() {} });
  const ctx = { $, kaCompanyMark: () => '', escapeHtml: s => String(s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch])), STAGE_COLOR: { B: '#f57c00', C: '#1a73e8' }, STAGE_LBL: { B: 'B｜Upside', C: 'C｜Pipeline' } };
  vm.createContext(ctx);
  vm.runInContext(fn, ctx);
  const mk = (i, extra) => Object.assign({ id: 'd' + i, company: 'Co' + i, product: 'P', amount: 1000, stage: 'B', oldDate: '2027-03-01', newDate: '2027-03-11', daysDiff: 10, changedAt: '2026-10-01', owner: 'u', postponeCount: 0, postponeDays: 0, stageDays: 10, repeated: false }, extra);
  const data = { promoted: [], demoted: [], newDeals: [], lostDeals: [], amountChanges: [], dateChanges: [
    mk(1, { repeated: true, postponeCount: 3, postponeDays: 30, stageDays: 60 }), mk(2, { postponeCount: 1, postponeDays: 10, stageDays: 5 }), mk(3), mk(4), mk(5), mk(6),
  ] };
  vm.runInContext('renderPrMoves', ctx)(data);
  const html = els.prDateChangeList.innerHTML;
  assert(html.includes('第 3 次延後｜累計 +30 天｜停在 B 60 天'), '[渲染] 顯示「第 N 次延後｜累計 +N 天｜停在 B N 天」', html.slice(0, 300));
  assert((html.match(/⚠ 反覆延後/g) || []).length === 1, '[渲染] 只有 repeated 的列有徽章');
  assert(html.includes('第 1 次延後｜累計 +10 天｜停在 B 5 天'), '[渲染] 非 repeated 但有延後也顯示累計');
  assert(!/style="[^"]*color:\s*#/.test(html.replace(/<span class="pr-stage-tag"[^>]*>/g, '')), '[渲染] 日期區塊除階段標籤外沒有 inline color', html.match(/style="[^"]*color[^"]*"/g));
  assert(html.includes('class="pr-dc-new"'), '[渲染] 新日期用 .pr-dc-new class');
  assert(els.prDateChangeToggle === undefined || els.prDateChangeToggle.style.display === '', '[渲染] 超過 5 筆顯示展開鈕');
}
function staticSuite(assert) {
  const html = fs.readFileSync(path.join(ROOT, '_client', 'index.html'), 'utf8');
  const css = fs.readFileSync(path.join(ROOT, '_client', 'style.css'), 'utf8');
  const glass = fs.readFileSync(path.join(ROOT, '_client', 'glass.css'), 'utf8');
  const iDem = html.indexOf('id="prDemotedSection"'), iDc = html.indexOf('id="prDateChangeSection"'), iNew = html.indexOf('id="prNewSection"');
  assert(iDem > 0 && iDc > iDem && iDc < iNew, '[靜態] 日期調整區塊在「階段退後」之後、「本期新增」之前');
  assert(html.includes('id="prDateChangeToggle"'), '[靜態] 有展開鈕');
  for (const c of ['pr-dc-new', 'pr-dc-old', 'pr-dc-late', 'pr-dc-early', 'pr-dc-badge', 'pr-dc-meta', 'pr-dc-muted', 'pr-dc-arrow']) {
    assert(css.includes('.' + c), `[靜態] style.css 定義 .${c}`);
  }
  for (const c of ['pr-dc-new', 'pr-dc-late', 'pr-dc-early']) assert(glass.includes('.' + c), `[靜態] glass.css 覆寫 .${c}`);
  assert(/\.pr-move-list-date\s*\{[^}]*max-height:\s*(4\d\d)px/.test(css) && /\.pr-move-list-date\.expanded\s*\{[^}]*max-height:\s*none/.test(css), '[靜態] 日期清單高度放寬且可展開');
  const block = sliceBetween(appSrc, '// 預計簽約日調整（顏色一律走', '// ── 合約資料載入');
  assert(!/color:\s*#1a3c7a/.test(block), '[靜態] app.js 不再寫死 #1a3c7a');
  assert(/body\.dark \.pr-summary-card\[data-kpi="date"\] \.pr-summary-value\s*\{[^}]*#c4b5fd/.test(css), '[靜態] 日期延後統計卡深色數字色覆寫');
  assert(/label:'日期延後'/.test(appSrc) && /其中反覆延後/.test(appSrc), '[靜態] 頂部統計卡有「日期延後（其中反覆延後）」');
}

// ═════════ 4) 變異 ═════════
function mutate(src, from, to) {
  if (!src.includes(from)) throw new Error('變異點不存在（實作被改過？）：' + from);
  return src.replace(from, () => to);
}

(async () => {
  const nowMs = Date.now();
  section('1) 單元（lib/pipelineDateWarning.js）');
  unitSuite(loadLib(libSrc), mkAsserter(null));

  section('3) 前端（getPrDateRange／renderPrMoves／靜態）');
  frontendSuite(appSrc, mkAsserter(null));
  renderSuite(appSrc, mkAsserter(null));
  staticSuite(mkAsserter(null));

  let apiRan = false;
  section('2) API（真實 server，埠 ' + PORT + '）');
  if (fs.existsSync(DATA_FILE) || fs.existsSync(AUTH_FILE)) {
    console.log('  略過：根目錄已有 data.json／auth.json（可能是真實資料，不覆蓋）——未驗證');
  } else {
    try {
      const a = mkAsserter(null);
      const cur = await collect('server.js', {}, nowMs);
      apiAssertions(cur, a, nowMs);
      apiRan = true;
      const bsrc = baseServerSrc();
      if (bsrc) {
        const tmp = path.join(ROOT, '_s_base_server.js');
        fs.writeFileSync(tmp, bsrc);
        try { compareToBase(cur, await collect('_s_base_server.js', {}, nowMs), a); } finally { try { fs.unlinkSync(tmp); } catch {} }
      } else console.log('  略過基底逐欄比對：取不到 merge-base 的 server.js——未驗證');
      // 環境變數覆寫（經由真實 server）
      const ov = await collect('server.js', { PIPELINE_REPEAT_MIN_COUNT: '3', PIPELINE_REPEAT_MIN_DAYS: '1000' }, nowMs);
      const d3 = ov._s_admin.dateChanges.find(d => d.id === 't3'), d1 = ov._s_admin.dateChanges.find(d => d.id === 'r1');
      a(d3 && d3.repeated === false, '[API env] MIN_COUNT=3／MIN_DAYS=1000 → t3(次數2) 不標', d3);
      a(d1 && d1.repeated === true, '[API env] r1(次數3) 仍標', d1);
      a(ov._s_admin.repeatRule.minCount === 3 && ov._s_admin.repeatRule.minDays === 1000, '[API env] repeatRule 反映覆寫值', ov._s_admin.repeatRule);
    } catch (e) { fail++; console.log('  FAIL API 段例外：' + (e && e.stack || e)); cleanFiles(); }
  }

  section('4) 變異測試（改壞實作，對應測試必須變紅）');
  const muts = [
    { name: 'M1 次數門檻 >= 改 >', file: 'lib', from: 'cum.count >= rules.minCount', to: 'cum.count > rules.minCount', suite: 'unit' },
    { name: 'M2 天數門檻 >= 改 >', file: 'lib', from: 'cum.days >= rules.minDays', to: 'cum.days > rules.minDays', suite: 'unit' },
    { name: 'M3 停留 > 45 改 >=', file: 'lib', from: 'stageDays > rules.stallDays', to: 'stageDays >= rules.stallDays', suite: 'unit' },
    { name: 'M4 少處理建立紀錄(oldDate=null 不再被忽略)', file: 'lib', from: "c && c.oldDate && c.newDate && !Number.isNaN(dayNum(c.oldDate)) && !Number.isNaN(dayNum(c.newDate))", to: 'c && c.newDate', suite: 'unit' },
    { name: 'M5 Won／D 也標(取消階段限制)', file: 'lib', from: 'rules.stages.includes(opp && opp.stage) &&', to: 'true &&', suite: 'unit' },
    { name: 'M6 天數基準改成第一筆 newDate', file: 'lib', from: 'dayNum(list[0].oldDate)', to: 'dayNum(list[0].newDate)', suite: 'unit' },
    { name: 'M7 停留天數忽略 stageHistory 只用 createdAt', file: 'lib', from: 'if (h && h.to === opp.stage && h.date', to: 'if (false && h.to === opp.stage && h.date', suite: 'unit' },
    { name: 'M8 自訂區間起點退回 new Date(f)', file: 'app', from: "new Date(f + 'T00:00:00')", to: 'new Date(f)', suite: 'frontend' },
    { name: 'M9 日期區塊改回寫死 #1a3c7a', file: 'app', from: '<span class="pr-dc-new">${escapeHtml(c.newDate)}</span>\n              <span class="${diffCls}">', to: '<span style="color:#1a3c7a">${escapeHtml(c.newDate)}</span>\n              <span class="${diffCls}">', suite: 'render' },
    { name: 'M10 server 改用本期而非全歷史', file: 'server', from: 'evaluate(opp, fullDCByDeal[dealId], now)', to: 'evaluate(opp, dcByDeal[dealId], now)', suite: 'api' },
    { name: 'M11 server 排序退回舊規則(只看本期天數)', file: 'server', from: 'dateChanges.sort(pipelineDateWarning.compareEntries);', to: 'dateChanges.sort((a, b) => (b.daysDiff ?? -1) - (a.daysDiff ?? -1));', suite: 'api' },
    { name: 'M12 server 預設區間退回 server 本地月界', file: 'server', from: 'const fromDate = from ? new Date(from) : _tpMonth.from;', to: 'const fromDate = from ? new Date(from) : new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1));', suite: 'api' },
  ];
  let mutRed = 0;
  for (const m of muts) {
    if (m.suite === 'api' && !apiRan) { console.log(`  略過 ${m.name}：API 段未執行——未驗證`); continue; }
    const sink = [];
    const assert = mkAsserter(sink);
    try {
      if (m.suite === 'unit') unitSuite(loadLib(mutate(libSrc, m.from, m.to)), assert);
      else if (m.suite === 'frontend') frontendSuite(mutate(appSrc, m.from, m.to), assert);
      else if (m.suite === 'render') renderSuite(mutate(appSrc, m.from, m.to), assert);
      else {
        const tmp = path.join(ROOT, '_s_mut_server.js');
        fs.writeFileSync(tmp, mutate(serverSrc, m.from, m.to));
        try { apiAssertions(await collect('_s_mut_server.js', {}, nowMs), assert, nowMs); } finally { try { fs.unlinkSync(tmp); } catch {} }
      }
    } catch (e) { sink.push('例外：' + e.message); }
    if (sink.length) { mutRed++; pass++; console.log(`  OK   ${m.name} → 紅（${sink.length} 項失敗，例：${sink[0]}）`); }
    else { fail++; console.log(`  FAIL ${m.name} → 測試仍是綠的（沒抓到）`); }
  }

  cleanFiles();
  console.log(`\n結果：${pass} 通過，${fail} 失敗` + (apiRan ? '' : '（API 段未執行）'));
  process.exit(fail ? 1 : 0);
})().catch(e => { cleanFiles(); console.error(e); process.exit(1); });
