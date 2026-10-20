#!/usr/bin/env node
'use strict';
/**
 * 程式碼審查（2026-10）PR-2 報價／簽核 ＋ PR-3 稽核 ＋ M-01 閒置登出合併 的檢查。不含任何帳密：測試密碼隨機產生、只活在記憶體。
 *
 *   D1  D-01／D-15  稽核事件 well-formed（孤立代理字元、截斷不切代理對、seal→往返→雜湊相等、字面 \u0000 不被吃掉）
 *   D2  D-02        補欄位／建索引：先探測、缺才跑、交易內 SET LOCAL lock_timeout（假 pool 驗 SQL 序列；真 PG 未驗證）
 *   D3  D-07/08/14  錨點排程：鏈設定了卻不可用 → 503；冪等；verify 預設比對已存錨點；breakRisk 帶進回應
 *   D9  D-09        兩人覆核：同提案人同種類 pending 上限、180 天前已結案申請壓縮
 *   D5  D-05        登入鎖定：bcrypt 比對完成後再檢查一次（真的 server.js，並發請求）
 *   M1  M-01        閒置登出單一機制：SESSION_IDLE_MINUTES 生效、輪詢清單共用、loginGuard 在 JWT 模式不重複、401 標頭
 *   C1  C-01/M-04   設計器：sweep 抽樣上限 3,000、floorProven（取樣被截斷＝不能證明）
 *   C2  C-02/C-06   發出機制：存檔失敗回 503 並回滾、簽章失敗不留幽靈紀錄、/seal 存檔失敗不消耗憑證
 *   C4  C-04        引擎模式紀錄：非 production 部署不寫、10 分鐘節流、非同步存檔失敗還原
 *   C5  C-05        防拆單：凍結快照不消失、撤回清快照、hiddenCount 只計入「計入累計」的
 *   C7  C-07        flow-config 例外清單：筆數上限、重複去重
 *   C9  C-09        已發出文件的草稿不可由業務刪除；stillValidOf 看 q.invalid
 *   C11 C-11        GET /api/quote-approval/config：要有功能或實際關係
 *   MUT 變異        把修法改壞，對應區段必須變紅（C-02、C-06、D-01、D-05、M-01 必做；其餘盡量）
 *
 *   node scripts/check-codrev-quote-audit.js
 *   CRQA_ONLY=D1,C2 node scripts/check-codrev-quote-audit.js      只跑指定區
 *   CRQA_SKIP_MUT=1 node scripts/check-codrev-quote-audit.js      略過變異
 *
 * 守門：D5 用 scripts/_b2-harness.js（拒絕在已有 auth.json／data.json 的資料夾執行；預設埠 3211，用 B2_PORT 改）。
 * 變異是「就地改檔→跑子行程→還原」，備份檔 *.crqa_bak；中途被打斷時下次啟動會先還原。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const L = (f) => path.join(ROOT, f);
const ONLY = process.env.CRQA_ONLY ? new Set(process.env.CRQA_ONLY.split(',').map(s => s.trim().toUpperCase())) : null;
const want = (k) => !ONLY || ONLY.has(k);
const CHILD = process.env.CRQA_CHILD === '1';
const BAK = '.crqa_bak';
const MUT_FILES = ['lib/auditStore.js', 'lib/quoteIssuance.js', 'lib/quoteRoutes.js', 'lib/quoteSplitGuard.js', 'lib/configReview.js', 'lib/auditChainRoutes.js', 'lib/sessionIdle.js', 'lib/loginGuard.js', 'middleware/jwtSession.js', 'server.js', 'lib/quoteFlowDesigner.js'];
if (!CHILD) for (const f of MUT_FILES) { if (fs.existsSync(L(f) + BAK)) { fs.copyFileSync(L(f) + BAK, L(f)); fs.unlinkSync(L(f) + BAK); console.log('（已還原上次被打斷的變異：' + f + '）'); } }

let pass = 0, fail = 0; const failures = [];
const ok = (c, n) => { if (c) pass++; else { fail++; failures.push(n); console.log('  ✗ FAIL:', n); } };
const eq = (a, b, n) => ok(JSON.stringify(a) === JSON.stringify(b), `${n}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`);
const sleep = (n) => new Promise(r => setTimeout(r, n));
async function quietly(fn) {   // 暫時吃掉 console.error／warn，回傳收到的訊息
  const e0 = console.error, w0 = console.warn, msgs = [];
  console.error = (...a) => msgs.push(a.map(String).join(' ')); console.warn = (...a) => msgs.push(a.map(String).join(' '));
  try { await fn(); } finally { console.error = e0; console.warn = w0; }
  return msgs;
}
const withEnv = async (vars, fn) => {
  const saved = {}; for (const k of Object.keys(vars)) { saved[k] = process.env[k]; if (vars[k] === undefined) delete process.env[k]; else process.env[k] = vars[k]; }
  try { return await fn(); } finally { for (const k of Object.keys(vars)) { if (saved[k] === undefined) delete process.env[k]; else process.env[k] = saved[k]; } }
};

// ═════════════════════════ D1：D-01／D-15 ═════════════════════════
async function secD1() {
  console.log('[D1] D-01／D-15 稽核事件 well-formed 與雜湊往返');
  const AS = require(L('lib/auditStore.js')), AC = require(L('lib/auditChain.js'));
  const cfg = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: crypto.randomBytes(24).toString('hex') });
  const lone = '\ud800';
  const ev = { ts: '2026-10-20T00:00:00.000Z', kind: 'log', action: 'LOGIN_FAIL', operator: 'a' + lone + 'b', target: 't\udc00', detail: 'x'.repeat(19999) + '😀', ip: '1.2.3.4', ua: 'u'.repeat(499) + '😀', extra: { k: lone, ['key' + lone]: 'v', n: { arr: ['ok', '\udc01'] } } };
  const row = AS.normalizeEvent(ev);
  for (const f of ['action', 'operator', 'target', 'detail', 'ip', 'ua']) ok(row[f].isWellFormed(), `[D1] normalizeEvent.${f} 是 well-formed 字串`);
  ok(!/\\ud[89a-f]/i.test(JSON.stringify(row.extra)) && JSON.stringify(row.extra).includes('"k"') , '[D1] extra 值與鍵都沒有孤立代理字元');
  eq(row.detail, 'x'.repeat(19999), '[D1] detail 在 20000 邊界截到 emoji：不切代理對（也不留 U+FFFD）');
  eq(row.ua, 'u'.repeat(499), '[D1] ua 在 500 邊界截到 emoji：不切代理對');
  // 往返：node-pg 以 UTF-8 寫出，孤立代理字元會變 U+FFFD；正規化後往返前後的雜湊必須相同
  const rt = (s) => Buffer.from(s, 'utf8').toString('utf8');
  const rtRow = Object.assign({}, row, { action: rt(row.action), operator: rt(row.operator), target: rt(row.target), detail: rt(row.detail), ip: rt(row.ip), ua: rt(row.ua), extra: JSON.parse(rt(JSON.stringify(row.extra))) });
  eq(AC.seal(cfg, null, rtRow).row_hash, AC.seal(cfg, null, row).row_hash, '[D1] seal → UTF-8 往返 → 雜湊相等');
  // 對照：未正規化的原始事件，往返後雜湊就不同（證明這個測試有鑑別力）
  const raw = Object.assign({}, ev, { extra: {} });
  const rawRt = Object.assign({}, raw, { operator: rt(raw.operator) });
  ok(AC.seal(cfg, null, raw).row_hash !== AC.seal(cfg, null, rawRt).row_hash, '[D1] 對照：未正規化的孤立代理字元往返後雜湊會變（問題確實存在）');
  // 走真正的 pgAdapter 寫入路徑（假 pool）：送進資料庫的參數與雜湊一致
  const pg = mkFakePg({ cols: true, indexes: ['audit_events_kind_ts_idx', 'audit_events_action_idx', 'audit_events_operator_idx', 'audit_events_seq_uidx'] });
  const ad = AS.pgAdapter({ query: pg.query, connect: pg.connect, chain: cfg });
  await ad.insert(row);
  const ins = pg.log.find(x => /^INSERT INTO audit_events .*seq, prev_hash/.test(x.sql));
  ok(!!ins && ins.params.length === 13, '[D1] 鏈寫入有送出 13 個參數的 INSERT');
  if (ins) {
    const p = ins.params;
    ok(p.slice(1, 8).every(v => typeof v !== 'string' || v.isWellFormed()) && !/\\ud[89a-f]/i.test(p[8]), '[D1] 送進資料庫的欄位與 extra JSON 都是 well-formed（jsonb 不會拒收）');
    const sealedAgain = AC.seal(cfg, null, { ts: row.ts, kind: p[1], action: p[2], operator: rt(p[3]), target: rt(p[4]), detail: rt(p[5]), ip: p[6], ua: p[7], extra: JSON.parse(rt(p[8])) });
    eq(sealedAgain.row_hash, p[11], '[D1] 資料庫讀回的值重算雜湊 = 寫入時的雜湊（不會永久 HASH_MISMATCH）');
  }
  // D-15：字面上的「反斜線 u0000」文字不被吃掉；真正的 NUL 被拿掉
  const e2 = AS.normalizeEvent({ extra: { t: 'x\\u0000y', n: 'a\u0000b', ['k\u0000']: 1 } });
  eq([e2.extra.t, e2.extra.n, Object.keys(e2.extra).includes('k')], ['x\\u0000y', 'ab', true], '[D15] 字面 \\u0000 文字保留、真實 NUL 與 NUL 鍵被清掉');
  eq(AS.normalizeEvent({ detail: 'a\u0000b' }).detail, 'ab', '[D1] detail 的 NUL 仍被拿掉（回歸）');
}

// ═════════════════════════ 假 Postgres（D1／D2 共用）═════════════════════════
function mkFakePg(o) {
  o = o || {};
  const log = [];
  const exec = async (sql, params) => {
    const s = String(sql).replace(/\s+/g, ' ').trim();
    log.push({ sql: s, params });
    if (/^SELECT indexname FROM pg_indexes/.test(s)) { if (o.probeFail) throw new Error('probe failed'); return { rows: (o.indexes || []).map(n => ({ indexname: n })) }; }
    if (/^SELECT count\(\*\)::int AS n FROM information_schema\.columns/.test(s)) { if (o.probeFail) throw new Error('probe failed'); return { rows: [{ n: o.cols ? 4 : 0 }] }; }
    if (/^SELECT seq, row_hash FROM audit_events WHERE seq IS NOT NULL/.test(s)) return { rows: [] };
    if (o.ddlLockTimeout && /^(ALTER TABLE|CREATE UNIQUE INDEX)/.test(s)) throw Object.assign(new Error('canceling statement due to lock timeout'), { code: '55P03' });
    return { rows: [] };
  };
  return { log, query: exec, connect: async () => ({ query: exec, release() { } }) };
}

// ═════════════════════════ D2：D-02 ═════════════════════════
async function secD2() {
  console.log('[D2] D-02 DDL 先探測、缺才跑、交易內設 lock_timeout（假 pool；真 Postgres 未驗證）');
  const AS = require(L('lib/auditStore.js')), AC = require(L('lib/auditChain.js'));
  const cfg = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: crypto.randomBytes(24).toString('hex') });
  const ALL = ['audit_events_kind_ts_idx', 'audit_events_action_idx', 'audit_events_operator_idx', 'audit_events_seq_uidx'];
  const isDdl = (s) => /^(ALTER TABLE|CREATE INDEX|CREATE UNIQUE INDEX)/.test(s);
  const run = async (o, chain) => { const pg = mkFakePg(o); const ad = AS.pgAdapter({ query: pg.query, connect: pg.connect, chain }); const msgs = await quietly(() => ad.ensure()); return { pg, ad, msgs, sqls: pg.log.map(x => x.sql) }; };
  // a. 全都在 → 不送任何 DDL、不開交易
  let r = await run({ cols: true, indexes: ALL }, cfg);
  eq(r.sqls.filter(isDdl), [], '[D2a] 欄位與索引都已存在 → 不送 ALTER／CREATE INDEX');
  ok(!r.sqls.includes('BEGIN'), '[D2a] 也不開交易');
  ok(r.sqls.some(s => /^SELECT indexname FROM pg_indexes/.test(s)) && r.sqls.some(s => /information_schema\.columns/.test(s)), '[D2a] 有先探測 pg_indexes 與 information_schema');
  // b. 什麼都沒有 → 每個 DDL 都在 BEGIN 之後、SET LOCAL lock_timeout 之後、COMMIT 之前
  r = await run({ cols: false, indexes: [] }, cfg);
  { let inTx = false, locked = false, bad = 0, n = 0;
    for (const s of r.sqls) { if (s === 'BEGIN') { inTx = true; locked = false; } else if (/^SET LOCAL lock_timeout = \d+$/.test(s)) locked = true; else if (s === 'COMMIT' || s === 'ROLLBACK') { inTx = false; locked = false; } else if (isDdl(s)) { n++; if (!(inTx && locked)) bad++; } }
    ok(n === 5, '[D2b] 全缺時補 3 個一般索引＋補欄位＋seq 唯一索引（共 ' + n + ' 句 DDL）');
    eq(bad, 0, '[D2b] 每一句 DDL 都在交易內且已 SET LOCAL lock_timeout'); }
  ok(r.sqls.some(s => /^SET LOCAL lock_timeout = 5000$/.test(s)), '[D2b] DDL 的 lock_timeout 預設 5000 毫秒');
  ok(r.sqls.findIndex(s => /^ALTER TABLE/.test(s)) < r.sqls.findIndex(s => /^CREATE UNIQUE INDEX/.test(s)), '[D2b] 先補欄位、再建 seq 唯一索引');
  // c. 欄位在、seq 索引缺 → 只建唯一索引
  r = await run({ cols: true, indexes: ALL.slice(0, 3) }, cfg);
  eq(r.sqls.filter(isDdl).map(s => s.slice(0, 26)), ['CREATE UNIQUE INDEX IF NOT'], '[D2c] 只缺 seq 唯一索引 → 只送這一句（不 ALTER）');
  // d. 探測失敗 → 照舊全跑
  r = await run({ probeFail: true }, cfg);
  eq(r.sqls.filter(isDdl).length, 5, '[D2d] 探測失敗 → 退回舊行為（全部 DDL 都送）');
  // e. 鏈未啟用 → 與第二波位元級相同：TABLE_SQL＋3 個 CREATE INDEX，沒有探測、沒有交易
  r = await run({}, null);
  eq(r.sqls.length, 4, '[D2e] 鏈未啟用 → 只有 1 個 CREATE TABLE＋3 個 CREATE INDEX');
  ok(r.sqls.every(s => !/pg_indexes|information_schema|BEGIN|lock_timeout/.test(s)), '[D2e] 鏈未啟用 → 不探測、不開交易（預設行為不變）');
  // f. 拿不到鎖（55P03）：不丟例外、不永久標記 DDL 失敗，之後寫入照常（不上鏈）
  r = await run({ cols: false, indexes: ALL.slice(0, 3), ddlLockTimeout: true }, cfg);
  eq(r.ad.chainStats().ddlFailed, false, '[D2f] 55P03 只是拿不到鎖 → 不永久標記 ddlFailed（30 秒後可重跑）');
  ok(r.msgs.some(m => /chain DDL failed|BREAK_RISK/.test(m)), '[D2f] 有記錄 AUDIT_CHAIN_BREAK_RISK');
  const pg2 = mkFakePg({ cols: false, indexes: ALL.slice(0, 3), ddlLockTimeout: true });
  const ad2 = AS.pgAdapter({ query: pg2.query, connect: pg2.connect, chain: cfg });
  await quietly(() => ad2.insert(AS.normalizeEvent({ action: 'X', operator: 'o' })));
  ok(pg2.log.some(x => /^INSERT INTO audit_events \(ts, kind, action, operator, target, detail, ip, ua, extra\) VALUES/.test(x.sql)), '[D2f] 拿不到鎖時事件仍寫入（不上鏈的 INSERT），不丟失');
  // g. 環境變數可調 lock_timeout
  await withEnv({ AUDIT_DDL_LOCK_TIMEOUT_MS: '2500' }, async () => {
    const pg = mkFakePg({ cols: false, indexes: [] });
    const ad = AS.pgAdapter({ query: pg.query, connect: pg.connect, chain: cfg });
    await quietly(() => ad.ensure());
    ok(pg.log.some(x => x.sql === 'SET LOCAL lock_timeout = 2500'), '[D2g] AUDIT_DDL_LOCK_TIMEOUT_MS=2500 生效');
  });
}

// ═════════════════════════ D3：D-07／D-08／D-14 ═════════════════════════
function mkChainApp(o) {
  const routes = {}; const app = {};
  ['get', 'post'].forEach(m => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = [].concat(...h); }; });
  const data = o.data || {}; const logs = [];
  const db = { load: () => data, save: () => { }, flush: async () => { }, auditStore: { chain: o.cfg, chainSource: async () => o.src, mode: () => 'table', stats: () => ({ chain: o.stats || null }) } };
  require(L('lib/auditChainRoutes.js')).register(app, { requireAdmin: (q, s, n) => n(), writeLog: (a, op, t, d) => logs.push({ a, d }), db, env: o.env || {} });
  const call = (key, req) => new Promise((resolve, reject) => {
    const h = routes[key];
    const rq = Object.assign({ headers: {}, query: {}, body: {}, session: { user: { username: 'adm' } } }, req || {});
    const rs = { _s: 200, status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s, j }); return this; }, set() { return this; } };
    let i = 0; const next = () => { const f = h[i++]; if (f) { try { const x = f(rq, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return { call, data, logs };
}
async function secD3() {
  console.log('[D3] D-07／D-08／D-14 錨點排程與驗證');
  const AC = require(L('lib/auditChain.js'));
  const CRON = 'c'.repeat(24);
  const cronReq = { headers: { authorization: 'Bearer ' + CRON } };
  // D-07：設定了要啟用卻不可用 → 503＋console.error；沒設定 → 200 enabled:false
  let app = mkChainApp({ cfg: AC.getConfig({ AUDIT_CHAIN: '1' }), src: null, env: { CRON_SECRET: CRON } });
  let r; let msgs = await quietly(async () => { r = await app.call('GET /api/cron/audit-anchor', cronReq); });
  ok(r.s === 503 && r.j.ok === false && r.j.configured === true, '[D3/D-07] AUDIT_CHAIN=1 但鏈不可用 → 503（Vercel 看得到）（' + r.s + '）');
  ok(msgs.some(m => /AUDIT_CHAIN_DISABLED/.test(m)), '[D3/D-07] 有 console.error AUDIT_CHAIN_DISABLED');
  app = mkChainApp({ cfg: AC.getConfig({}), src: null, env: { CRON_SECRET: CRON } });
  r = await app.call('GET /api/cron/audit-anchor', cronReq);
  ok(r.s === 200 && r.j.ok === true && r.j.enabled === false, '[D3/D-07] 沒設定啟用 → 仍是 200 enabled:false（行為不變）');
  r = await app.call('GET /api/cron/audit-anchor', { headers: {} });
  eq(r.s, 401, '[D3] 沒帶 CRON_SECRET → 401（回歸）');
  // 可用的鏈（假資料來源）
  const cfg = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: crypto.randomBytes(24).toString('hex') });
  const H = 'ab'.repeat(32);
  const src = { chainHead: async () => ({ seq: 5, row_hash: H }), chainCount: async () => 5, chainReadable: async () => true, chainUncovered: async () => ({ legacyUncovered: 0, unchainedAfterStart: 2 }), chainRowAt: async (s) => (src.truncated ? null : (s === 5 ? { row_hash: H } : { row_hash: 'cd'.repeat(32) })), async *chainRows() { } };
  app = mkChainApp({ cfg, src, env: { CRON_SECRET: CRON }, stats: { breakRisk: 3 } });
  msgs = await quietly(async () => { r = await app.call('GET /api/cron/audit-anchor', cronReq); });
  ok(r.s === 200 && r.j.enabled === true && r.j.stored === 1 && !r.j.deduped, '[D3/D-14] 第一次：產生並存入錨點');
  eq(r.j.risk, { breakRisk: 3, unchainedAfterStart: 2 }, '[D3/D-08] 回應帶出 breakRisk 與鏈後未上鏈筆數');
  ok(msgs.some(m => /AUDIT_CHAIN_BREAK_RISK/.test(m)), '[D3/D-08] 有 console.error AUDIT_CHAIN_BREAK_RISK');
  ok(app.logs.some(l => l.a === 'AUDIT_CHAIN_ANCHOR_CRON' && /breakRisk=3/.test(l.d)), '[D3/D-08] 稽核事件的說明帶 breakRisk');
  r = await app.call('GET /api/cron/audit-anchor', cronReq);
  ok(r.s === 200 && r.j.deduped === true && app.data.auditAnchors.length === 1, '[D3/D-14] 重複送達（同 seq、20 小時內）→ 不再新增（仍 200）');
  app.data.auditAnchors[0].ts = new Date(Date.now() - 21 * 3600e3).toISOString();
  r = await app.call('GET /api/cron/audit-anchor', cronReq);
  ok(r.s === 200 && !r.j.deduped && app.data.auditAnchors.length === 2, '[D3/D-14] 距上一份超過 20 小時 → 照常新增');
  src.chainHead = async () => ({ seq: 6, row_hash: H });
  r = await app.call('GET /api/cron/audit-anchor', cronReq);
  ok(app.data.auditAnchors.length === 3, '[D3/D-14] 鏈頭有前進（seq 不同）→ 一定新增');
  // D-14：verify 預設比對最新一份已存錨點
  src.chainHead = async () => ({ seq: 5, row_hash: H });
  app = mkChainApp({ cfg, src, env: { CRON_SECRET: CRON } });
  await app.call('GET /api/cron/audit-anchor', cronReq);
  r = await app.call('GET /api/admin/audit-chain/verify', {});
  ok(r.s === 200 && r.j.ok === true && r.j.anchorSource === 'stored' && r.j.anchor && r.j.anchor.checked === true && r.j.anchor.ok === true, '[D3/D-14] 沒帶錨點的 verify → 預設比對已存錨點（通過）');
  src.truncated = true;
  r = await app.call('GET /api/admin/audit-chain/verify', {});
  ok(r.j.ok === false && r.j.anchor && r.j.anchor.reason === 'TRUNCATED' && r.j.anchorSource === 'stored', '[D3/D-14] 尾端被截（錨點那一筆不在了）→ 預設 verify 就抓得到 TRUNCATED');
  r = await app.call('GET /api/admin/audit-chain/verify', { query: { storedAnchor: '0' } });
  ok(r.j.ok === true && !r.j.anchor && !r.j.anchorSource, '[D3/D-14] ?storedAnchor=0 可關閉自動比對');
  src.truncated = false;
  // head：帶出 unchainedAfterStart
  r = await app.call('GET /api/admin/audit-chain/head', {});
  eq(r.j.unchainedAfterStart, 2, '[D3/D-08] head 帶出 unchainedAfterStart');
}

// ═════════════════════════ D9：D-09 ═════════════════════════
async function secD9() {
  console.log('[D9] D-09 兩人覆核申請上限與壓縮');
  const CR = require(L('lib/configReview.js'));
  const mk = (env) => {
    const T0 = Date.parse('2026-10-20T00:00:00Z'); const st = { now: T0 }; const data = { approvalChangeRequests: [] };
    const cr = CR.create({ db: { load: () => data, save() { }, flush: async () => { } }, writeLog() { }, pushNotification() { }, loadAuth: () => ({ users: [{ username: 'a1', role: 'admin', active: true }, { username: 'a2', role: 'admin', active: true }] }), now: () => st.now, env: Object.assign({ APPROVAL_CONFIG_DUAL_REVIEW: '1' }, env || {}) });
    cr.registerKind('k', { label: '測試', rawKeys: [], read: (d) => d.val || { n: 0 }, validate: (a) => ({ ok: true, value: a }), diff: (b, a) => ['n ' + b.n + '→' + a.n], apply() { return {}; } });
    const prop = (n, by) => { const r = { code: 0, body: null, status(c) { this.code = c; return this; }, json(b) { this.body = b; return this; } }; cr.propose(r, 'k', { data, value: { n }, proposer: by || 'a1', users: {}, note: '', req: {} }); return r; };
    return { data, st, prop };
  };
  let w = mk();
  const codes = [1, 2, 3, 4].map(n => w.prop(n).code);
  eq(codes, [202, 202, 202, 429], '[D9] 同提案人同種類 pending 第 4 筆被擋（429 TOO_MANY_PENDING）');
  eq(w.prop(9, 'a2').code, 202, '[D9] 另一位管理員不受影響');
  w.st.now += 4 * 24 * 3600e3;   // 逾期（3 天）的 pending 不再算
  eq(w.prop(10).code, 202, '[D9] 已逾期的 pending 不佔名額');
  w = mk({ APPROVAL_CONFIG_MAX_PENDING: '0' });
  eq([1, 2, 3, 4, 5].map(n => w.prop(n).code), [202, 202, 202, 202, 202], '[D9] APPROVAL_CONFIG_MAX_PENDING=0 → 不限（可還原）');
  // 壓縮
  w = mk();
  const old = (days, status, extra) => Object.assign({ id: 'r' + days + status, kind: 'k', proposedBy: 'a1', proposedAt: new Date(w.st.now - days * 86400e3).toISOString(), status, before: { n: 1 }, after: { n: 2 }, diffSummary: ['x'], history: [{ at: new Date(w.st.now - days * 86400e3).toISOString(), by: 'a2', action: status }], expiresAt: new Date(w.st.now).toISOString() }, extra || {});
  w.data.approvalChangeRequests.push(old(200, 'approved'), old(100, 'approved'), old(200, 'rejected'));
  w.prop(1);
  const byId = (id) => w.data.approvalChangeRequests.find(r => r.id === id);
  ok(!('before' in byId('r200approved')) && !('after' in byId('r200approved')) && byId('r200approved').compacted === true, '[D9] 180 天前已結案的申請：去掉 before／after 全文、標 compacted');
  ok(!('before' in byId('r200rejected')), '[D9] 退回的同樣壓縮');
  ok('before' in byId('r100approved') && !byId('r100approved').compacted, '[D9] 100 天前的不動');
  ok(byId('r200approved').diffSummary.length === 1 && byId('r200approved').history.length === 1, '[D9] 壓縮後保留摘要與歷程');
  const live = w.data.approvalChangeRequests[w.data.approvalChangeRequests.length - 1];
  ok(live.status === 'pending' && live.before && live.after, '[D9] 新的 pending 申請完整保留');
}

// ═════════════════════════ M1：M-01 ═════════════════════════
async function secM1() {
  console.log('[M1] M-01 閒置登出單一機制');
  process.env.SESSION_SECRET = process.env.SESSION_SECRET || 'unit-test-secret-not-for-production-use-0123456789';
  const SI = require(L('lib/sessionIdle.js')), LG = require(L('lib/loginGuard.js'));
  eq(SI.resolveIdleMs({}), 1800000, '[M1] 預設 30 分鐘');
  eq(SI.resolveIdleMs({ SESSION_IDLE_MINUTES: '60' }), 3600000, '[M1] SESSION_IDLE_MINUTES=60 → 60 分鐘');
  eq(SI.resolveIdleMs({ SESSION_IDLE_MINUTES: '60', SESSION_IDLE_MS: '8000' }), 8000, '[M1] SESSION_IDLE_MS（測試用）優先');
  eq(SI.resolveIdleMs({ SESSION_IDLE_MS: '100' }), 1800000, '[M1] SESSION_IDLE_MS 小於 5 秒 → 忽略（不會把全站鎖死）');
  eq([SI.resolveIdleMs({ SESSION_IDLE_MINUTES: '0' }), SI.resolveIdleMs({ SESSION_IDLE_MINUTES: 'abc' }), SI.resolveIdleMs({ SESSION_IDLE_MINUTES: '-5' })], [1800000, 1800000, 1800000], '[M1] 0／非數字／負數 → 預設');
  ok(['/api/poll-bundle', '/api/notifications', '/API/Announcements/Active/'].every(p => SI.isPollingRequest('GET', p)) && SI.isPollingRequest('HEAD', '/api/notifications'), '[M1] 三個背景輪詢路徑（GET／HEAD，不分大小寫與結尾斜線）');
  ok(!SI.isPollingRequest('POST', '/api/notifications') && !SI.isPollingRequest('GET', '/api/contacts'), '[M1] POST 或其他路徑算使用者操作');
  eq(LG.IDLE_POLL_PATHS, SI.IDLE_POLL_PATHS, '[M1] loginGuard 與 sessionIdle 共用同一份輪詢清單');
  // jwtSession：SESSION_IDLE_MINUTES 由子行程驗證（IDLE_MS 在載入時決定）
  const probe = (env) => spawnSync(process.execPath, ['-e', "const j=require('./middleware/jwtSession');console.log(j.IDLE_MS+','+Array.from(j.PASSIVE_PATHS).length)"], { cwd: ROOT, encoding: 'utf8', env: Object.assign({}, process.env, { SESSION_SECRET: 'unit-test-secret-not-for-production-use-0123456789', SESSION_IDLE_MS: '', SESSION_IDLE_MINUTES: '' }, env) }).stdout.trim();
  eq(probe({ SESSION_IDLE_MINUTES: '60' }), '3600000,3', '[M1] jwtSession.IDLE_MS 讀 SESSION_IDLE_MINUTES（60 分鐘）、輪詢清單 3 個');
  eq(probe({ SESSION_IDLE_MINUTES: '60', SESSION_IDLE_MS: '8000' }), '8000,3', '[M1] jwtSession：SESSION_IDLE_MS 優先');
  eq(probe({}), '1800000,3', '[M1] jwtSession 預設 30 分鐘');
  // jwtSession 行為（預設 30 分鐘視窗）
  const jwt = require(L('node_modules/jsonwebtoken')), jwtSession = require(L('middleware/jwtSession.js'));
  const nowS = () => Math.floor(Date.now() / 1000);
  const run = (tok, p, method) => { const req = { headers: { cookie: tok ? 'itts_auth=' + encodeURIComponent(tok) : '' }, path: p, method: method || 'GET' }; const cookies = [], cleared = []; jwtSession(req, { cookie(n, v, o) { cookies.push(o); }, clearCookie(n) { cleared.push(n); } }, () => { }); return { req, cookies, cleared }; };
  const lgn = nowS() - 3600, tok = jwt.sign({ user: { username: 'u1', role: 'user' }, lgn, lat: nowS() - 300, exp: lgn + 8 * 3600 }, process.env.SESSION_SECRET);
  ok(run(tok, '/api/notifications', 'GET').cookies.length === 0 && run(tok, '/api/announcements/active', 'GET').cookies.length === 0, '[M1] GET /api/notifications、/api/announcements/active 不續命（與 poll-bundle 相同）');
  eq([run(tok, '/api/notifications', 'POST').cookies.length, run(tok, '/api/contacts', 'GET').cookies.length], [1, 1], '[M1] POST 或一般 API 續命，且一個請求最多簽發一次 cookie');
  const old = jwt.sign({ user: { username: 'u1', role: 'user' }, lgn, lat: nowS() - 31 * 60, exp: lgn + 8 * 3600 }, process.env.SESSION_SECRET);
  const ro = run(old, '/api/contacts');
  ok(ro.req._sessionIdleExpired === true && ro.req._sessionIdleUser === 'u1' && !ro.req.session.user, '[M1] 閒置逾時：標記 _sessionIdleExpired 並帶出帳號（server.js 據此寫 SESSION_IDLE_TIMEOUT 稽核）');
  // loginGuard.idleMiddleware：JWT 模式不重複判斷、不重簽；memory 模式帶 X-Session-Expired
  const T0 = 1_700_000_000_000;
  const mkReq = (user) => { const calls = { destroy: 0, resign: 0 }; const state = { user }; const session = new Proxy(state, { get(t, k) { if (k === 'destroy') return (cb) => { calls.destroy++; t.user = null; cb && cb(); }; return t[k]; }, set(t, k, v) { t[k] = v; if (k === 'user') calls.resign++; return true; } }); const res = { code: null, headers: {}, status(c) { this.code = c; return this; }, json() { return this; }, set(k, v) { this.headers[k] = v; return this; }, clearCookie() { } }; return { req: { method: 'GET', path: '/api/contacts', session }, res, calls }; };
  const g = LG.create({ env: { SESSION_IDLE_MINUTES: '5' }, clock: () => T0 + 60 * 60000, store: { load: () => ({}), save() { } } });
  let x = mkReq({ username: 'u', loginAt: T0 - 20 * 3600e3, lat: T0 }); let n = 0;
  g.idleMiddleware({ isJwt: true })(x.req, x.res, () => n++);
  eq([n, x.res.code, x.calls.destroy, x.calls.resign], [1, null, 0, 0], '[M1] JWT 模式：loginGuard 直接放行（不判閒置、不判 8 小時、不重簽 cookie）');
  x = mkReq({ username: 'u', loginAt: T0, lat: T0 }); n = 0;
  g.idleMiddleware({ isJwt: false })(x.req, x.res, () => n++);
  eq([n, x.res.code, x.res.headers['X-Session-Expired']], [0, 401, 'idle'], '[M1] memory 模式：閒置的 401 帶 X-Session-Expired: idle');
}

// ═════════════════════════ D5：D-05（真的 server.js）═════════════════════════
async function secD5() {
  console.log('[D5] D-05 登入鎖定：bcrypt 比對完成後再檢查一次（並發）');
  const H = require('./_b2-harness');
  H.guardClean();
  const users = [H.mkUser('_s_admin', 'admin'), H.mkUser('_s_slow', 'user', { bu: ['ERP'], password: H.HASH12 })];
  try {
    await H.fresh(users, {}, { env: { LOGIN_LOCKOUT: '1', LOGIN_MAX_FAILS: '3', LOGIN_LOCK_MINUTES: '5', LOGIN_LOCK_MAX_MINUTES: '10' } });
    await H.loginAll(['_s_admin']);
    const bad = () => H.http(null, 'POST', '/api/login', { username: '_s_slow', password: 'wrong-' + Math.random() });
    const good = () => H.http(null, 'POST', '/api/login', { username: '_s_slow', password: H.PW });
    const goodStatuses = [];
    for (let round = 0; round < 2; round++) {
      const ps = [];
      for (let i = 0; i < 5; i++) { ps.push(bad()); await sleep(15); }   // 5 個錯誤密碼先出發（門檻 3）
      ps.push(good());                                                  // 正確密碼最後出發：它的 bcrypt 比對一定在門檻被推過之後才完成
      const rs = await Promise.all(ps);
      goodStatuses.push(rs[5].status);
      const lk = await H.http('_s_admin', 'GET', '/api/admin/login-locks');
      ok(lk.json && (lk.json.locks || []).some(l => l.username === '_s_slow'), `[D5] 第 ${round + 1} 輪：帳號已鎖定`);
      await H.http('_s_admin', 'POST', '/api/admin/users/_s_slow/unlock');
    }
    eq(goodStatuses, [401, 401], '[D5] 並發下「鎖定後才完成」的正確密碼請求一律 401（不得成功登入）');
    const r = await H.http(null, 'POST', '/api/login', { username: '_s_slow', password: H.PW });
    eq(r.status, 200, '[D5] 解鎖後正確密碼可登入（沒有誤傷）');
    const logs = await H.getLogs('_s_admin');
    ok(H.byAction(logs, 'LOGIN_FAIL').some(l => /LOCKED/.test(l.detail || '')), '[D5] 被拒的並發請求有稽核 LOGIN_FAIL（LOCKED）');
  } finally { await H.finish(); }
}

// ═════════════════════════ 報價單路由（假 app，Postgres 風格）═════════════════════════
const QA = require(L('lib/quoteApproval.js'));
const NOW = '2026-10-20T00:00:00Z';
function mkEnv(opt) {
  const reg = opt.reg || require(L('lib/quoteRoutes.js'));
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = { logs: [], flushes: 0, seq: 0, saves: 0, saveMode: 'ok' };
  env.data = Object.assign({ quotations: JSON.parse(JSON.stringify(opt.quotations || [])), quoteApproval: { roster: { gm: ['gm1'], chairman: [], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: { PS: { cls: 'software', costBySales: true } } } }, opt.data || {});
  env.auth = { users: JSON.parse(JSON.stringify(opt.users)) };
  const deps = {
    db: { load: () => env.data, save: () => { env.saves++; if (env.saveMode === 'reject') { const p = Promise.reject(new Error('simulated db failure')); p.catch(() => { }); return p; } if (env.saveMode === 'throw') throw new Error('simulated sync failure'); return Promise.resolve(); }, flush: async () => { } },
    loadAuth: () => env.auth, saveAuth: () => { },
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (a, op, t, d) => env.logs.push({ a, op, t, d }), pushNotification: () => { },
    flushAuditBeforeEnd: () => { env.flushes++; },
    getViewableOwners: (req) => (opt.scope || {})[req.session.user.username] || [req.session.user.username],
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-20', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from('xlsx'), buildQuotePnlExcel: async () => Buffer.from(''), QUOTE_TEMPLATE: __filename,
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: opt.features || (() => ['quotations']),
  };
  const out = reg(app, deps);
  env.internal = out && out._internal;
  env.call = (user, method, p, params, body, extra) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const hdr = Object.assign({}, (extra && extra.headers) || {});
    const req = { method: method.toUpperCase(), ip: '127.0.0.1', headers: hdr, get: (k) => hdr[String(k).toLowerCase()] || '', session: { user: { username: user, role } }, params: params || {}, query: (extra && extra.query) || {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
    const rs = { _h: {}, status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j, h: this._h }); return this; }, send() { resolve({ s: this._s || 200, j: null, h: this._h }); return this; }, setHeader(k, v) { this._h[k] = v; }, set(k, v) { this._h[k] = v; return this; } };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}
const U = (username, role, extra) => Object.assign({ username, role, active: true, displayName: username, bu: [] }, extra || {});
const mkQuote = (id, owner, extra) => Object.assign({
  id, quoteNo: 'QN-' + id, owner, company: 'TestCo', projectName: 'P', quoteDate: '2026-10-20', status: 'draft', createdAt: NOW, updatedAt: NOW, validUntil: '2026-10-30',
  products: ['PS'], approval: null, discountType: 'none', discountValue: 0,
  items: [{ lid: 'i1-' + id, desc: '軟體授權', unit: '套', qty: 1, unitPrice: 100000, cost: 60000 }],
}, extra || {});
const approved = (q) => { q.approval = { state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: NOW, submittedBy: q.owner, derived: null, steps: [], cur: 0, board: null, history: [] }; return q; };

// ─── C2：C-02／C-06 ───
async function secC2() {
  console.log('[C2] C-02／C-06 發出機制：存檔失敗回 503＋回滾；簽章失敗不留幽靈紀錄');
  const QIS = require(L('lib/quoteIssuance.js'));
  const SECRET = 'crqa-secret-' + crypto.randomBytes(8).toString('hex');
  await withEnv({ JWT_SECRET: SECRET, SESSION_SECRET: SECRET, QUOTE_ISSUE_SECRET: undefined, QUOTE_ISSUE_REUSE_SEC: '0', QUOTE_ISSUE_RATE_PER_MIN: '0' }, async () => {
    const users = [U('own1', 'user', { bu: ['ITS'] }), U('adm', 'admin')];
    const mk = () => mkEnv({ users, quotations: [approved(mkQuote('Q1', 'own1'))], data: { quoteSeal: { mime: 'image/png', base64: 'iVBORw0KGgo=' } } });
    // 基準
    let env = mk();
    let r = await env.call('own1', 'POST', '/api/quotations/:id/issue', { id: 'Q1' }, { format: 'pdf' });
    ok(r.s === 200 && /^ISS-QN-Q1-001$/.test(r.j.issueNo) && typeof r.j.stampToken === 'string', '[C2] 基準：核准有效 → 200、發出序號、蓋章憑證');
    eq([env.data.quoteIssuances.length, env.data.quotations[0].issuanceCount, env.saves >= 1], [1, 1, true], '[C2] 基準：紀錄與摘要已寫入並存檔');
    // C-02：/issue 存檔被拒（Promise rejected）
    env = mk(); env.saveMode = 'reject';
    r = await env.call('own1', 'POST', '/api/quotations/:id/issue', { id: 'Q1' }, { format: 'pdf' });
    ok(r.s === 503 && r.j.code === 'ISSUE_SAVE_FAILED', '[C2/C-02] /issue 存檔失敗 → 503 ISSUE_SAVE_FAILED（得到 ' + r.s + '）');
    ok(!('stampToken' in (r.j || {})) && !('issueNo' in (r.j || {})), '[C2/C-02] 失敗的回應不含序號與蓋章憑證');
    eq([(env.data.quoteIssuances || []).length, env.data.quotations[0].issuanceCount, (env.data.quotations[0].issuances || []).length], [0, undefined, 0], '[C2/C-02] 記憶體內的發出紀錄與摘要已回滾（共用快取不會把幽靈紀錄帶進下次存檔）');
    ok(!env.logs.some(l => l.a === 'ISSUE_QUOTE'), '[C2/C-02] 失敗時不寫 ISSUE_QUOTE 稽核（沒有發出）');
    // 同步丟例外（JSON 模式）一樣 503
    env = mk(); env.saveMode = 'throw';
    r = await env.call('own1', 'POST', '/api/quotations/:id/issue', { id: 'Q1' }, { format: 'pdf' });
    eq([r.s, (env.data.quoteIssuances || []).length], [503, 0], '[C2/C-02] 存檔同步丟例外 → 同樣 503 並回滾');
    // 失敗後恢復：序號不會被跳號（因為已回滾）
    env.saveMode = 'ok';
    r = await env.call('own1', 'POST', '/api/quotations/:id/issue', { id: 'Q1' }, { format: 'pdf' });
    ok(r.s === 200 && r.j.issueNo === 'ISS-QN-Q1-001', '[C2/C-02] 恢復後發出的第一個序號仍是 001（失敗那次沒有佔號）');
    // GET /export（xlsx）
    env = mk(); env.saveMode = 'reject';
    r = await env.call('own1', 'GET', '/api/quotations/:id/export', { id: 'Q1' });
    ok(r.s === 503 && !(r.h && r.h['X-Issue-No']), '[C2/C-02] /export 存檔失敗 → 503、沒有 X-Issue-No 標頭');
    env = mk();
    r = await env.call('own1', 'GET', '/api/quotations/:id/export', { id: 'Q1' });
    ok(r.s === 200 && /^ISS-/.test((r.h || {})['X-Issue-No'] || ''), '[C2/C-02] /export 正常 → 200＋X-Issue-No（基準）');
    // /seal：存檔失敗不消耗憑證，使用者可重試
    env = mk();
    const iss = await env.call('own1', 'POST', '/api/quotations/:id/issue', { id: 'Q1' }, { format: 'pdf' });
    env.saveMode = 'reject';
    r = await env.call('own1', 'GET', '/api/quotations/:id/seal', { id: 'Q1' }, undefined, { headers: { 'x-stamp-token': iss.j.stampToken } });
    ok(r.s === 503, '[C2/C-02] /seal 存檔失敗 → 503，不送章（得到 ' + r.s + '）');
    eq((env.data.quoteStampUses || []).length, 0, '[C2/C-02] /seal 失敗後憑證使用紀錄已回滾（憑證沒被消耗）');
    env.saveMode = 'ok';
    r = await env.call('own1', 'GET', '/api/quotations/:id/seal', { id: 'Q1' }, undefined, { headers: { 'x-stamp-token': iss.j.stampToken } });
    ok(r.s === 200, '[C2/C-02] 重試同一張憑證 → 200 送章');
    r = await env.call('own1', 'GET', '/api/quotations/:id/seal', { id: 'Q1' }, undefined, { headers: { 'x-stamp-token': iss.j.stampToken } });
    eq(r.s, 403, '[C2] 同一張憑證只能用一次（回歸）');
    // C-06：簽章失敗（缺所有金鑰）→ 不留紀錄
    env = mk();
    await withEnv({ JWT_SECRET: undefined, SESSION_SECRET: undefined, QUOTE_ISSUE_SECRET: undefined }, async () => {
      const q = approved(mkQuote('QX', 'own1')); const data = { quotations: [q], quoteIssuances: [] };
      let threw = false; try { QIS.issue(data, { q, format: 'pdf', me: 'own1', role: 'user', ip: '1.1.1.1' }); } catch (e) { threw = true; }
      ok(threw, '[C2/C-06] 沒有任何簽章金鑰 → issue() 丟例外（前提）');
      eq([data.quoteIssuances.length, q.issuanceCount, q.issuances], [0, undefined, undefined], '[C2/C-06] 簽章失敗時沒有留下 official 紀錄與單上摘要（先簽後改）');
      const msgs = await quietly(async () => { r = await env.call('own1', 'POST', '/api/quotations/:id/issue', { id: 'Q1' }, { format: 'pdf' }); });
      ok(r.s === 500 && r.j.code === 'ISSUE_FAILED' && (env.data.quoteIssuances || []).length === 0 && !('stampToken' in r.j), '[C2/C-06] 路由：簽章失敗 → 500 ISSUE_FAILED、沒有紀錄、沒有憑證（' + r.s + '）');
      ok(msgs.some(m => /issue failed/.test(m)), '[C2/C-06] 有 console.error');
    });
    // rollback／unconsume 單元
    const q = approved(mkQuote('QR', 'own1')); const data = { quotations: [q], quoteIssuances: [] };
    const a = QIS.issue(data, { q, format: 'xlsx', me: 'own1' }), b = QIS.issue(data, { q, format: 'xlsx', me: 'own1' });
    QIS.rollbackIssue(data, b.rec);
    eq([data.quoteIssuances.map(x => x.issueNo), q.issuanceCount, q.issuances.map(x => x.issueNo)], [[a.rec.issueNo], 1, [a.rec.issueNo]], '[C2] rollbackIssue 只移除指定那筆、計數與摘要同步回退');
    data.quoteStampUses = [{ jti: 'j1' }, { jti: 'j2' }]; QIS.unconsumeStamp(data, 'j1');
    eq(data.quoteStampUses.map(x => x.jti), ['j2'], '[C2] unconsumeStamp 只移除指定憑證');
  });
}

// ─── C4 ───
async function secC4() {
  console.log('[C4] C-04 引擎模式紀錄：非 production 不寫、節流、非同步失敗還原');
  const users = [U('adm', 'admin')];
  const mk = (saveMode) => { const e = mkEnv({ users, quotations: [] }); if (saveMode) e.saveMode = saveMode; return e; };
  const note = (e) => e.internal.noteEngineMode({ data: e.data, me: 'adm' }, { method: 'GET' });
  const modeLogs = (e) => e.logs.filter(l => l.a === 'QUOTE_FLOW_ENGINE_MODE').length;
  await withEnv({ QUOTE_FLOW_ENGINE: 'shadow', QUOTE_FLOW_MODE_THROTTLE_MS: undefined, VERCEL_ENV: 'preview', QUOTE_FLOW_MODE_AUDIT_NON_PROD: undefined }, async () => {
    let e = mk(); note(e);
    eq([e.saves, 'quoteFlowEngineMode' in e.data, modeLogs(e)], [0, false, 0], '[C4] VERCEL_ENV=preview：不存檔、不記錄、不寫稽核（不改共用資料庫）');
    await withEnv({ QUOTE_FLOW_MODE_AUDIT_NON_PROD: '1' }, async () => { e = mk(); note(e); eq([e.saves, modeLogs(e)], [1, 1], '[C4] QUOTE_FLOW_MODE_AUDIT_NON_PROD=1 可還原舊行為'); });
    await withEnv({ VERCEL_ENV: 'production' }, async () => {
      e = mk(); note(e); note(e);
      eq([e.saves, modeLogs(e), e.data.quoteFlowEngineMode.mode], [1, 1, 'shadow'], '[C4] production：寫一次（同模式不重複）');
      await withEnv({ QUOTE_FLOW_ENGINE: 'on' }, async () => { note(e); eq(modeLogs(e), 1, '[C4] 10 分鐘節流：剛寫過，模式又變也先不記'); });
      await withEnv({ QUOTE_FLOW_ENGINE: 'on', QUOTE_FLOW_MODE_THROTTLE_MS: '0' }, async () => { note(e); eq(modeLogs(e), 2, '[C4] QUOTE_FLOW_MODE_THROTTLE_MS=0 → 不節流（測試／還原用）'); });
    });
    await withEnv({ VERCEL_ENV: undefined }, async () => { e = mk(); note(e); eq([e.saves, modeLogs(e)], [1, 1], '[C4] 沒有 VERCEL_ENV（本機／舊部署）→ 照舊會寫（向後相容）'); });
    await withEnv({ VERCEL_ENV: 'production' }, async () => {
      e = mk('reject'); await quietly(async () => { note(e); await sleep(20); });
      ok(!('quoteFlowEngineMode' in e.data), '[C4] Postgres 非同步存檔失敗 → 記憶體內的紀錄還原（過去只有同步例外有還原）');
      await withEnv({ QUOTE_FLOW_MODE_THROTTLE_MS: '0' }, async () => { e.saveMode = 'ok'; note(e); });
    });
  });
}

// ─── C5 ───
async function secC5() {
  console.log('[C5] C-05 防拆單：快照不消失、撤回清除、hiddenCount');
  const QSG = require(L('lib/quoteSplitGuard.js'));
  // ③ hiddenCount：看不到的姊妹單，只有「計入累計」的才算
  const res = { applicable: true, suspect: true, ownCents: 1, cumCents: 3, ownLevel: 0, cumLevel: 1, thresholdYuan: 10, windowDays: 90, countedCount: 1, oppId: 'O',
    siblings: [{ id: 'a', q: {}, counted: true, quoteNo: 'A', revenueCents: 1, approvalState: 'none' }, { id: 'b', q: {}, counted: false, why: 'OUT_OF_WINDOW', quoteNo: 'B', revenueCents: 1, approvalState: 'none' }, { id: 'c', q: {}, counted: false, why: 'INVALID', quoteNo: 'C', revenueCents: 1, approvalState: 'none' }] };
  eq(QSG.toPublic(res, () => false).hiddenCount, 1, '[C5/③] 3 張都看不到、只有 1 張計入 → hiddenCount 1（過去誤報 3）');
  eq(QSG.toPublic(res, (x) => true).hiddenCount, 0, '[C5/③] 全看得到 → 0');
  // ① 凍結快照不消失；② 撤回清快照
  const users = [U('own1', 'user', { bu: ['ITS'] }), U('adm', 'admin')];
  const snap = { v: 1, at: NOW, by: 'own1', trigger: 'SUBMIT', mode: 'warn', ownCents: 5e8, cumCents: 2e9, ownLevel: 0, cumLevel: 1, thresholdYuan: 10000000, windowDays: 90, siblings: [{ id: 'gone', quoteNo: 'QN-GONE', revenueCents: 1.5e9, approvalState: 'none' }] };
  await withEnv({ QUOTE_SPLIT_GUARD: 'warn' }, async () => {
    const q = approved(mkQuote('S1', 'own1', { oppId: 'O1', splitGuard: snap }));
    const env = mkEnv({ users, quotations: [q] });
    const r = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'S1' });
    ok(r.s === 200 && r.j.split && r.j.split.atSubmit && r.j.split.atSubmit.cumCents === 2e9, '[C5/①] 核准中／已核准的單：即使同商機的草稿姊妹單已被刪光，送簽時的凍結快照仍然顯示（過去整塊消失）（' + JSON.stringify(r.j && r.j.split).slice(0, 120) + '）');
    eq(r.j.split && r.j.split.atSubmit && r.j.split.atSubmit.hiddenCount, 1, '[C5/①] 快照裡已不存在的姊妹單依「看不到」計入 hiddenCount');
    const q2 = mkQuote('S2', 'own1', { oppId: 'O1', splitGuard: snap });   // 草稿狀態：快照不給
    const env2 = mkEnv({ users, quotations: [q2] });
    const r2 = await env2.call('own1', 'GET', '/api/quotations/:id', { id: 'S2' });
    ok(!r2.j.split, '[C5/①] 草稿（非簽核中／已核准）沒有姊妹單時不多出 split 區塊（回歸）');
  });
  // ② 撤回：清掉 q.splitGuard
  {
    const q = mkQuote('S3', 'own1', { oppId: 'O1', splitGuard: snap });
    q.approval = { state: 'pending', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: NOW, submittedBy: 'own1', derived: { tiers: ['mgr1'] }, steps: [{ tier: 'mgr1', assignee: 'adm', status: 'pending' }], cur: 0, board: null, history: [] };
    const env = mkEnv({ users, quotations: [q] });
    const r = await env.call('own1', 'POST', '/api/quotations/:id/withdraw', { id: 'S3' }, {});
    ok(r.s === 200 && !('splitGuard' in env.data.quotations[0]), '[C5/②] 撤回後 q.splitGuard 被清掉（快照只屬於那一輪送簽）（' + r.s + '）');
  }
}

// ─── C7 ───
async function secC7() {
  console.log('[C7] C-07 flow-config 例外清單');
  const users = [U('adm', 'admin'), U('mgr1', 'manager1', { bu: ['ITS'] }), U('mgr2', 'manager1', { bu: ['ITS'] })];
  const env = mkEnv({ users, quotations: [] });
  const put = (ex) => env.call('adm', 'PUT', '/api/admin/quote-approval/flow-config', null, { buSettings: {}, exceptions: ex });
  let r = await put(Array.from({ length: 201 }, (_, i) => ({ username: 'mgr1', action: 'skip_mgr1', id: 'e' + i })));
  ok(r.s === 400 && r.j.code === 'TOO_MANY_EXCEPTIONS', '[C7] 例外超過 200 條 → 400 TOO_MANY_EXCEPTIONS（' + r.s + '）');
  r = await put([{ username: 'mgr1', action: 'skip_mgr1', id: 'a' }, { username: 'mgr1', action: 'skip_mgr1', id: 'b' }, { username: 'mgr2', action: 'skip_mgr1', id: 'c' }]);
  ok(r.s === 200 && env.data.approvalFlowConfig.exceptions.length === 2 && env.data.approvalFlowConfig.exceptions[0].id === 'a', '[C7] 同帳號同動作重複 → 只留第一條（共 2 條）（' + r.s + '）');
  r = await put([null]); eq(r.s, 400, '[C7] [null] 仍是 400（快修回歸）');
  r = await put(Array.from({ length: 200 }, (_, i) => ({ username: 'mgr1', action: 'skip_mgr1', id: 'e' + i })));
  eq(r.s, 200, '[C7] 恰好 200 條 → 通過（邊界）');
}

// ─── C9 ───
async function secC9() {
  console.log('[C9] C-09 已發出文件的草稿不可由業務刪除；stillValidOf 看 q.invalid');
  const QIS = require(L('lib/quoteIssuance.js'));
  const users = [U('own1', 'user', { bu: ['ITS'] }), U('adm', 'admin')];
  const rec = (qid) => ({ issueNo: 'ISS-QN-' + qid + '-001', quoteId: qid, quoteNo: 'QN-' + qid, kind: 'draft', format: 'xlsx', contentHash: 'h', approvalHash: null, approvedValid: false, issuedBy: 'own1', issuedAt: NOW });
  let env = mkEnv({ users, quotations: [mkQuote('D1', 'own1'), mkQuote('D2', 'own1')], data: { quoteIssuances: [rec('D1')] } });
  let r = await env.call('own1', 'DELETE', '/api/quotations/:id', { id: 'D1' }, {});
  ok(r.s === 409 && r.j.code === 'HAS_ISSUANCE' && env.data.quotations.length === 2, '[C9] 業務刪除「已發出過文件」的草稿 → 409 HAS_ISSUANCE、單還在（' + r.s + '）');
  r = await env.call('own1', 'DELETE', '/api/quotations/:id', { id: 'D2' }, {});
  ok(r.s === 200 && env.data.quotations.length === 1, '[C9] 沒發出過文件的草稿照常可刪（回歸）');
  r = await env.call('adm', 'DELETE', '/api/quotations/:id', { id: 'D1' }, {});
  ok(r.s === 200 && env.logs.some(l => l.a === 'DELETE_QUOTATION' && /已發出 1 筆文件/.test(l.d)), '[C9] 管理員仍可刪，稽核註記「已發出 1 筆文件」（' + r.s + '）');
  // 摘要快取不準時，仍以權威的 quoteIssuances 為準；反過來摘要說有、紀錄被截掉也擋
  env = mkEnv({ users, quotations: [mkQuote('D3', 'own1', { issuanceCount: 2 })] });
  r = await env.call('own1', 'DELETE', '/api/quotations/:id', { id: 'D3' }, {});
  eq(r.s, 409, '[C9] 單上 issuanceCount>0 也擋（保守）');
  // stillValidOf
  const q = approved(mkQuote('V1', 'own1')); const data = { quotations: [q], quoteIssuances: [] };
  const sec = 'crqa-' + crypto.randomBytes(6).toString('hex');
  const got = await withEnv({ JWT_SECRET: sec }, async () => QIS.issue(data, { q, format: 'xlsx', me: 'own1' }));
  ok(QIS.stillValidOf(got.rec, q) === true, '[C9] 基準：有效單的發出紀錄 stillValid');
  q.invalid = true;
  ok(QIS.stillValidOf(got.rec, q) === false, '[C9] 商機作廢（q.invalid）後 stillValid 變 false');
}

// ─── C11 ───
async function secC11() {
  console.log('[C11] C-11 GET /api/quote-approval/config 的存取');
  const users = [U('own1', 'user', { bu: ['ITS'] }), U('mkt1', 'marketing'), U('gm1', 'marketing'), U('cons1', 'marketing'), U('rel1', 'manager1', { bu: ['ITS'] }), U('sec1', 'secretary', { bu: ['ITS'] }), U('adm', 'admin')];
  const features = (role) => (role === 'marketing' || role === 'manager1' ? [] : ['quotations']);
  const pending = mkQuote('P1', 'own1'); pending.approval = { state: 'pending', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(pending), submittedAt: NOW, submittedBy: 'own1', derived: { tiers: ['mgr1'] }, steps: [{ tier: 'mgr1', assignee: 'rel1', status: 'pending' }], cur: 0, board: null, history: [] };
  const env = mkEnv({ users, quotations: [pending], features });
  const code = async (u) => (await env.call(u, 'GET', '/api/quote-approval/config')).s;
  const res = {};
  for (const u of ['own1', 'adm', 'mkt1', 'gm1', 'cons1', 'rel1', 'sec1']) res[u] = await code(u);
  eq(res, { own1: 200, adm: 200, mkt1: 403, gm1: 200, cons1: 200, rel1: 200, sec1: 200 }, '[C11] 有報價單功能、管理員、名冊成員（總經理／成本填寫人）、被指派簽核人、秘書可讀；沒有功能也沒有關係的角色 403');
  await withEnv({ QUOTE_CONFIG_OPEN: '1' }, async () => { eq(await code('mkt1'), 200, '[C11] QUOTE_CONFIG_OPEN=1 → 還原成只驗登入'); });
}

// ─── C1 ───
async function secC1() {
  console.log('[C1] C-01／M-04 設計器 sweep');
  const FD = require(L('lib/quoteFlowDesigner.js')), QF = require(L('lib/quoteFlow.js'));
  const CFG = { buSettings: {}, exceptions: [] };
  const env = { roster: { gm: ['g'], chairman: ['c'], boardProxy: ['b'] }, signerState: () => 'ok', chain: { owners: 1, noMgr1: 0, noMgr2: 0 }, activeDef: QF.legacyDefault(CFG) };
  const base = { amount: '1', marginLevel: 1, category: 'consult', ownerBu: 'ERP' };
  const legacy = () => JSON.parse(JSON.stringify(QF.buildLegacyDefault(CFG)));
  let s = FD.sweep(legacy(), base, env, QF.legacyDefault(CFG), CFG);
  ok(s.ok && s.summary.floorProven === true && s.summary.floorViolations === 0, '[C1] 預設流程：floorProven=true、違反底線 0');
  // 取樣被灌爆（60,000 組）→ floorProven=false，且抽樣上限 3,000
  const d = legacy(); const owners = Array.from({ length: 60 }, (_, i) => 'u' + i);
  d.routes.unshift({ id: 'r_evil', name: '捷徑', when: { cmp: ['owner', 'in', owners] }, path: ['n_mgr1'] });
  d.routes.unshift({ id: 'r_pad', name: '墊', when: { all: [{ cmp: ['amount', '<', '0.01'] }, { cmp: ['marginPct', '>', '1'] }, { cmp: ['marginPct', '>', '2'] }, { cmp: ['ownerBuPrimary', 'in', ['X1', 'X2', 'X3']] }, { cmp: ['classKeys', 'contains', 'consult'] }] }, path: ['n_mgr1', 'n_gm', 'n_chairman'] });
  s = FD.sweep(d, base, env, QF.legacyDefault(CFG), CFG);
  ok(s.ok && s.summary.truncated === true && s.summary.floorProven === false, '[C1] 取樣被截斷 → floorProven=false（不能把 floorViolations=0 當成通過）');
  eq(s.summary.sampledEvery, Math.ceil(s.summary.samples / 3000), '[C1] 逐筆 plan 的抽樣上限是 3,000 組（M-04；過去 12,000）');
  ok(s.summary.sampledEvery > 1, '[C1] 大取樣確實有抽樣（sampledEvery=' + s.summary.sampledEvery + '）');
  const flowUi = fs.readFileSync(L('_client/flow-designer.js'), 'utf8');
  ok(/floorProven === false/.test(flowUi), '[C1] 畫面在 floorProven=false 時加註「取樣被截斷，無法證明底線成立」');
}

// ═════════════════════════ 變異 ═════════════════════════
const MUTANTS = [
  // [名稱, 檔案, from, to, 該跑哪些區段]
  ['D-01 well-formed 失效', 'lib/auditStore.js', 'function wellFormed(s) {', 'function wellFormed(s) { return s;', 'D1'],
  ['D-01 截斷切在代理對中間', 'lib/auditStore.js', 'if (last >= 0xD800 && last <= 0xDBFF) end--;', '', 'D1'],
  ['D-15 又對輸出文字做 regex', 'lib/auditStore.js', "if (typeof v === 'string') return wellFormed(v.replace(/\\u0000/g, ''));", "if (typeof v === 'string') return wellFormed(v.replace(/\\u0000/g, '').replace(/\\\\u0000/g, ''));", 'D1'],
  ['D-02 一律全跑 DDL', 'lib/auditStore.js', 'const missingIdx = INDEX_SQL.filter((s) => !have || !have.has(idxNameOf(s)));', 'const missingIdx = INDEX_SQL.slice();', 'D2'],
  ['D-02 DDL 不設 lock_timeout', 'lib/auditStore.js', "await client.query(`SET LOCAL lock_timeout = ${DDL_LOCK_TIMEOUT_MS}`);", '', 'D2'],
  ['D-07 鏈不可用仍回 200', 'lib/auditChainRoutes.js', 'if (c.cfg && c.cfg.flag) {', 'if (false) {', 'D3'],
  ['D-14 錨點不冪等', 'lib/auditChainRoutes.js', 'if (prevA && Number(prevA.seq) === Number(a.seq)', 'if (false && prevA && Number(prevA.seq) === Number(a.seq)', 'D3'],
  ['D-14 verify 不比對已存錨點', 'lib/auditChainRoutes.js', "if (!anchor && String(q.storedAnchor) !== '0') {", 'if (false) {', 'D3'],
  ['D-09 pending 無上限', 'lib/configReview.js', 'if (maxPending > 0 && store(data)', 'if (false && store(data)', 'D9'],
  ['D-09 不壓縮', 'lib/configReview.js', '    compactOld(data);\n', '', 'D9'],
  ['D-05 登入比對後不再檢查鎖定', 'server.js', '回應與鎖定時完全相同\n    if (loginGuard.checkLocked(username, req.ip).locked) {', '回應與鎖定時完全相同\n    if (false) {', 'D5'],
  ['M-01 SESSION_IDLE_MINUTES 不生效', 'lib/sessionIdle.js', 'if (Number.isFinite(m) && m > 0) return', 'if (false) return', 'M1'],
  ['M-01 輪詢清單少一個', 'lib/sessionIdle.js', "const IDLE_POLL_PATHS = ['/api/poll-bundle', '/api/notifications', '/api/announcements/active'];", "const IDLE_POLL_PATHS = ['/api/poll-bundle'];", 'M1'],
  ['M-01 jwtSession 讓所有請求續命', 'middleware/jwtSession.js', '!SI.isPollingRequest(req.method, req.path) && (Date.now()', 'true && (Date.now()', 'M1'],
  ['M-01 loginGuard 在 JWT 模式重複判斷', 'lib/loginGuard.js', 'if (o.isJwt) return next();', '', 'M1'],
  ['M-01 閒置 401 不帶標頭', 'lib/loginGuard.js', "res.set('X-Session-Expired', d.expired === 'idle' ? 'idle' : '1');", '', 'M1'],
  ['C-02 /issue 不等存檔', 'lib/quoteRoutes.js', 'try { await db.save(ctx.data); }\n    catch (e) {\n      QIS.rollbackIssue(ctx.data, r.rec);', 'try { db.save(ctx.data); }\n    catch (e) {\n      QIS.rollbackIssue(ctx.data, r.rec);', 'C2'],
  ['C-02 存檔失敗不回滾', 'lib/quoteRoutes.js', '      QIS.rollbackIssue(ctx.data, r.rec);\n', '', 'C2'],
  ['C-02 /seal 失敗不還原憑證使用紀錄', 'lib/quoteRoutes.js', '      QIS.unconsumeStamp(ctx.data, used.rec.tokenJti);\n', '', 'C2'],
  ['C-02 /seal 不等存檔', 'lib/quoteRoutes.js', 'try { await db.save(ctx.data); }\n    catch (e) {\n      QIS.unconsumeStamp', 'try { db.save(ctx.data); }\n    catch (e) {\n      QIS.unconsumeStamp', 'C2'],
  ['C-06 先改資料後簽憑證', 'lib/quoteIssuance.js', "const stampToken = rec.tokenJti ? signStampToken(rec, p.me, nowMs) : null;\n  data.quoteIssuances.push(rec);", "data.quoteIssuances.push(rec);\n  const stampToken = rec.tokenJti ? signStampToken(rec, p.me, nowMs) : null;", 'C2'],
  ['C-09 stillValid 不看 invalid', 'lib/quoteIssuance.js', 'if (q.invalid) return false;', '', 'C9'],
  ['C-09 刪除不擋已發出', 'lib/quoteRoutes.js', "if (issuedN > 0 && !rel.isAdmin) return", "if (false) return", 'C9'],
  ['C-04 預覽部署也寫', 'lib/quoteRoutes.js', 'if (modeNonProd()) return;', '', 'C4'],
  ['C-04 不節流', 'lib/quoteRoutes.js', 'if (Date.now() - modeWroteAt < modeThrottleMs()) return;', '', 'C4'],
  ['C-05 快照在無姊妹單時消失', 'lib/quoteRoutes.js', 'if ((!sg || !sg.res.siblings.length) && !frozenKept) return null;', 'if (!sg || !sg.res.siblings.length) return null;', 'C5'],
  ['C-05 撤回不清快照', 'lib/quoteRoutes.js', '    if (q.splitGuard) delete q.splitGuard;   // C-05②：撤回也清掉', '    // mut   // C-05②：撤回也清掉', 'C5'],
  ['C-05 hidden 全算', 'lib/quoteSplitGuard.js', 'if (!canView(s.q)) { if (s.counted) hidden++; return; }', 'if (!canView(s.q)) { hidden++; return; }', 'C5'],
  ['C-07 不限筆數', 'lib/quoteRoutes.js', 'if (b.exceptions.length > MAX_FLOW_EXCEPTIONS)', 'if (false)', 'C7'],
  ['C-07 不去重', 'lib/quoteRoutes.js', 'if (seenEx.has(exKey)) continue;', '', 'C7'],
  ['C-11 config 不檢查功能', 'lib/quoteRoutes.js', "if (!configReadAllowed(ctx)) return fail", "if (false) return fail", 'C11'],
  ['C-01 floorProven 恆為 true', 'lib/quoteFlowDesigner.js', 'floorProven: !an.truncated,', 'floorProven: true,', 'C1'],
  ['C-01 sweep 抽樣回到 12,000', 'lib/quoteFlowDesigner.js', 'sweepPlanCap: 3000,', 'sweepPlanCap: 12000,', 'C1'],
];

function runMutants() {
  console.log('[MUT] 變異檢查：把修法改壞，對應區段必須變紅');
  let caught = 0;
  for (const [name, file, from, to, sec] of MUTANTS) {
    const p = L(file); const src = fs.readFileSync(p, 'utf8');
    const cnt = src.split(from).length - 1;
    if (cnt !== 1) { ok(false, `[MUT] 「${name}」變異點必須唯一（在 ${file} 找到 ${cnt} 處）`); continue; }
    fs.writeFileSync(p + BAK, src);
    let exit = -1, tail = '';
    try {
      fs.writeFileSync(p, src.replace(from, () => to));
      const r = spawnSync(process.execPath, [__filename], { cwd: ROOT, encoding: 'utf8', timeout: 300000, env: Object.assign({}, process.env, { CRQA_CHILD: '1', CRQA_ONLY: sec, CRQA_SKIP_MUT: '1' }) });
      exit = r.status; tail = (r.stdout || '').split('\n').filter(l => /FAIL|例外/.test(l)).slice(0, 2).join(' | ').slice(0, 160);
    } finally { fs.copyFileSync(p + BAK, p); fs.unlinkSync(p + BAK); }
    const red = exit !== 0;
    if (red) caught++;
    ok(red, `[MUT] 變異「${name}」→ ${sec} 區段變紅（exit=${exit}${red ? '；' + tail : ''}）`);
  }
  console.log(`  （${caught}/${MUTANTS.length} 個變異被抓到）`);
}

(async () => {
  try {
    if (want('D1')) await secD1();
    if (want('D2')) await secD2();
    if (want('D3')) await secD3();
    if (want('D9')) await secD9();
    if (want('M1')) await secM1();
    if (want('C1')) await secC1();
    if (want('C2')) await secC2();
    if (want('C4')) await secC4();
    if (want('C5')) await secC5();
    if (want('C7')) await secC7();
    if (want('C9')) await secC9();
    if (want('C11')) await secC11();
    if (want('D5')) await secD5();
    if (!CHILD && !process.env.CRQA_SKIP_MUT && want('MUT')) runMutants();
  } catch (e) { fail++; failures.push('例外 ' + (e && e.stack || e)); console.log('檢查腳本例外', e && e.stack || e); }
  console.log(`\ncodrev-quote-audit 結果：${pass} 通過，${fail} 失敗`);
  if (fail) console.log(failures.map(f => ' - ' + f).join('\n'));
  process.exit(fail ? 1 : 0);
})();
