#!/usr/bin/env node
/**
 * 寫入競爭防護（DB_WRITE_SERIALIZE／QUOTE_NO_ATOMIC）的邏輯檢查。用法：node scripts/check-write-guard.js
 * 使用 scripts/_fake-pg.js 的「假 Postgres」（模擬鎖互斥、lock_timeout、交易、原子 upsert、連線池、網路延遲）。
 * ★ 這驗證的是邏輯；不等於真實 Postgres 行為。真實 PG 請跑 scripts/verify-write-race-pg.js（由使用者執行）。
 *
 *  A. 分類器：哪些請求要鎖、豁免、會寫入的 GET、未登入不鎖
 *  B. pgWriteGuard：互斥、FIFO、lock_timeout→LOCK_TIMEOUT、伺服器 idle 逾時釋放、看門狗強制釋放、釋放可重複呼叫
 *  C. 原子取號：200 並發唯一且連續、floor 銜接舊資料、月份重置、≥1000 的格式、已發號紀錄
 *  D. 重現：多實例並發建立報價單 —— 開關全關必定掉單／重號（證明問題存在）；(b) 開：不掉不重；(a) 開：不重號且掉單可被對帳；兩個都開
 *  E. 中介層：鎖逾時回 503 DB_BUSY（友善訊息、Retry-After）、handler 不執行；存檔失敗回 503 DB_WRITE_FAILED 且不遺留鎖、快取失效；
 *     豁免路徑／未登入不拿鎖；鎖內 handler 看到最新資料
 *  F. 開關關閉：不建新表、不多任何 SQL、nextQuoteSeq 拒絕、middleware 直接放行
 *  G. quoteLedger.review：不傳 issued → 輸出與舊版相同；傳了 → 已發未用不算 GAP_UNEXPLAINED
 */
const path = require('path');
const { FakeDb, freshPostgres, sleep } = require('./_fake-pg');
const WG = require('../lib/writeGuard');
const { createPgWriteGuard, SQL } = require('../lib/pgWriteGuard');
const fsx = require('fs');
const RS = require('../lib/requestScope');
const QN = require('../lib/quoteNo');
const QL = require('../lib/quoteLedger');

const results = [];
// 防「假綠燈」：若事件迴圈清空、程式在印出摘要前就結束（某個 Promise 永不 settle），要印出已有結果並以非 0 結束
let summaryPrinted = false;
process.on('exit', (code) => {
  if (summaryPrinted) return;
  for (const [n, p, e] of results) console.log((p ? 'PASS ' : 'FAIL ') + n + (e && !p ? '  ← ' + e : ''));
  console.log('\n【未完成】測試在印出摘要前就結束了（有 Promise 永不結束＝中介層／鎖卡住），已通過 ' + results.filter((r) => r[1]).length + ' 條，視為失敗。');
  process.exitCode = code || 3;
});
const T_START = Date.now();
const ok = (name, cond, extra) => { results.push([name, !!cond, extra === undefined ? '' : String(extra)]); if (process.env.WG_TIMING) console.error('[' + (Date.now() - T_START) + 'ms] ' + name.slice(0, 40)); };
const quiet = { log() {}, warn() {}, error() {}, info() {} };

// ── 測試用小工具 ─────────────────────────────────────────
function mkRes() {
  let resolve; const done = new Promise((r) => { resolve = r; });
  const res = {
    statusCode: 200, headersSent: false, headers: {}, body: null, done,
    end(b) { this.body = b; this.headersSent = true; resolve(); return this; },
    set(k, v) { this.headers[String(k).toLowerCase()] = v; return this; },
    setHeader(k, v) { this.headers[String(k).toLowerCase()] = v; },
    removeHeader(k) { delete this.headers[String(k).toLowerCase()]; },
    status(c) { this.statusCode = c; return this; },
    json(o) { this.setHeader('content-type', 'application/json'); return this.end(JSON.stringify(o)); },
  };
  return res;
}
function mkInstance(fake, envFlags, cfgExtra, logObj, guardExtra) {
  const env = Object.assign({}, envFlags || {});
  const db = freshPostgres(fake, env);
  const cfg = Object.assign(WG.readConfig(env), cfgExtra || {});
  const guard = WG.createWriteGuard(Object.assign({ db, config: cfg, log: logObj || quiet }, guardExtra || {}));
  return { db, guard, mw: guard.middleware(), late: guard.lateLock('/api/admin/import-contacts'), cfg };   // 測試用 multipart 路由：/api/admin/import-contacts（與 server.js 的註冊方式相同）
}
/** 模擬一個請求：全域 ready 中介層 → writeGuard 中介層 → handler → res.end。回傳 { res, ran } */
async function request(inst, method, p, handler, opts) {
  opts = opts || {};
  await inst.db.ready();
  // 與 server.js 相同的掛載方式 app.use('/api/', …)：Express 會把 req.path 去掉 /api 前綴、放進 req.baseUrl（曾因此整個中介層沒生效，被 HTTP 整合測試抓到）
  const mounted = p.startsWith('/api/');
  const req = { method, baseUrl: mounted ? '/api' : '', path: mounted ? p.slice(4) : p, session: opts.anon ? {} : { user: { username: opts.user || 'u' } }, headers: opts.headers || {} };
  const res = mkRes();
  let ran = false;
  let store = null;   // 中介層 next() 當下的 ALS store：真實 Express 的 handler 就在這個脈絡裡同步執行，測試要還原它
  const nexted = new Promise((r) => inst.mw(req, res, () => { store = RS.current(); r('next'); }));
  const T = opts.timeoutMs || 60000;
  const first = await Promise.race([nexted, res.done.then(() => 'responded'), sleep(T).then(() => 'timeout')]);
  if (first === 'timeout') { ok('【逾時】request ' + method + ' ' + p + ' 在 ' + T + 'ms 內沒有放行也沒有回應（中介層卡住）', false); return { res, ran: false, timedOut: true }; }
  if (first === 'next') {
    if (opts.bodyMs != null) {   // multipart 路由：全域中介層放行後，路由內的 multer 讀 body（不佔鎖），讀完才進 lateLock
      if (opts.onBodyReading) opts.onBodyReading(inst);
      await sleep(opts.bodyMs);
      const lated = new Promise((r) => inst.late(req, res, () => { store = RS.current(); r('next'); }));
      const f2 = await Promise.race([lated, res.done.then(() => 'responded'), sleep(T).then(() => 'timeout')]);
      if (f2 !== 'next') return { res, ran: false, timedOut: f2 === 'timeout' };
    }
    ran = true;
    const fin = await Promise.race([Promise.resolve(RS.run(store, () => handler(req, res, inst.db))).then(() => res.done).then(() => 'done'), sleep(T).then(() => 'timeout')]);
    if (fin === 'timeout') { ok('【逾時】request ' + method + ' ' + p + ' 的 handler／回應在 ' + T + 'ms 內沒有結束', false); return { res, ran, timedOut: true }; }
  }
  return { res, ran };
}
/** p 在 ms 內沒結束就回傳 'TIMEOUT'（測試「不會永遠卡住」時用，避免測試本身卡死） */
const within = (p, ms) => Promise.race([Promise.resolve(p).then((v) => v, (e) => { throw e; }), sleep(ms).then(() => 'TIMEOUT')]);
const YM = '202610';
let uid = 0;
function createHandler(atomic, who) {
  return async (req, res, db) => {
    let data = db.load();
    if (!data.quotations) data.quotations = [];
    let no;
    if (atomic) {
      const got = await db.nextQuoteSeq(YM, QN.floorOf(data, YM), who);
      data = db.load();
      if (!data.quotations) data.quotations = [];
      QN.recordSeq(data, YM, got.serial);
      no = got.quoteNo;
    } else {
      const seq = QN.floorOf(data, YM) + 1;     // 與 server.js genQuoteNo 同邏輯
      QN.recordSeq(data, YM, seq);
      no = QN.format(YM, seq);
    }
    data.quotations.push({ id: 'q' + (++uid), quoteNo: no, owner: who });
    db.save(data);
    res.status(201).json({ quoteNo: no });
  };
}
async function runCreates(insts, atomic, n) {
  const jobs = [];
  for (let i = 0; i < n; i++) {
    const inst = insts[i % insts.length];
    jobs.push(request(inst, 'POST', '/api/quotations', createHandler(atomic, 'w' + i)));
  }
  return Promise.all(jobs);
}
const summarize = (fake, results_) => {
  const final = (fake.main() || {}).quotations || [];
  const nos = final.map((q) => q.quoteNo);
  const dup = nos.length - new Set(nos).size;
  const okRes = results_.filter((r) => r.res.statusCode === 201);
  return { created: okRes.length, finalCount: final.length, lost: okRes.length - final.length, dup, nos };
};

(async () => {
  // ═══ A. 分類器 ═══
  {
    const cfg = WG.readConfig({ DB_WRITE_SERIALIZE: '1' });
    const need = (m, p) => WG.needsLock(cfg, m, p);
    ok('A1 POST /api/quotations 要鎖', need('POST', '/api/quotations'));
    ok('A2 PUT/DELETE/PATCH 要鎖', need('PUT', '/api/contacts/1') && need('DELETE', '/api/contacts/1') && need('PATCH', '/api/x'));
    ok('A3 一般 GET 不鎖', !need('GET', '/api/quotations') && !need('GET', '/api/contacts'));
    ok('A4 會寫入的 GET 要鎖', need('GET', '/api/callins') && need('GET', '/api/announcements/active') && need('GET', '/api/admin/quote-ledger'));
    ok('A5 外部呼叫路由豁免', !need('POST', '/api/ai/ocr-card') && !need('POST', '/api/admin/integrations/push/batch') && !need('POST', '/api/admin/companies/enrich-gcis') && !need('POST', '/api/admin/ai-ocr-card') && !need('POST', '/api/admin/bulk-fill-website') && !need('GET', '/api/company-lookup'));
    ok('A6 登入／export-log／草稿試算豁免', !need('POST', '/api/login') && !need('POST', '/api/quotations/abc/export-log') && !need('POST', '/api/quotations/abc/cost-draft/summary'));
    ok('A7 非 /api/ 不鎖', !need('POST', '/login.html'));
    ok('A8 前綴比對不誤殺相似路徑', need('POST', '/api/aix') && need('POST', '/api/admin/integrationsX'));
    const cfg2 = WG.readConfig({ DB_WRITE_SERIALIZE: '1', DB_WRITE_LOCK_EXEMPT: '/api/foo', DB_WRITE_LOCK_EXTRA_PATHS: '/api/bar' });
    ok('A9 環境變數可追加豁免與 GET 寫入清單', !WG.needsLock(cfg2, 'POST', '/api/foo') && WG.needsLock(cfg2, 'GET', '/api/bar'));
    ok('A10 預設關閉', WG.readConfig({}).serialize === false && WG.readConfig({ DB_WRITE_SERIALIZE: '0' }).serialize === false);
    ok('A11 預設逾時 8s／持有 20s（< Vercel 30s）', cfg.waitMs === 8000 && cfg.holdMs === 20000);
  }

  // ═══ B. pgWriteGuard（鎖） ═══
  {
    const fake = new FakeDb({ latency: [0, 1] });
    const pool = fake.pool({ max: 4 });
    const mkG = () => createPgWriteGuard({ getPool: () => pool, lockKey: '42', idleTxMs: 400 });   // 每個 guard＝一個獨立的「實例」（同 guard 內另有行程內互斥）
    const g = mkG(), gB = mkG();
    const order = [];
    let inCrit = 0, maxCrit = 0;
    const worker = async (i) => {
      const h = await mkG().acquireLock({ waitMs: 5000 });
      inCrit++; maxCrit = Math.max(maxCrit, inCrit); order.push(i);
      await sleep(3);
      inCrit--;
      await h.release();
    };
    // 池 max=4：同時 4 個等鎖者，其餘在池外排隊，所以分批送出
    const ws = []; for (let i = 0; i < 12; i++) ws.push(worker(i));
    await Promise.all(ws);
    ok('B1 12 個並發取鎖：臨界區同時最多 1 人', maxCrit === 1, 'max=' + maxCrit);
    ok('B2 12 個都拿到鎖', order.length === 12);
    ok('B3 鎖都釋放（無殘留持有者）', fake.locks.size === 0);

    // 逾時
    const holder = await g.acquireLock({ waitMs: 1000 });
    const t0 = Date.now();
    let err;
    try { await gB.acquireLock({ waitMs: 120 }); } catch (e) { err = e; }
    ok('B4 等鎖逾時 → WriteLockError LOCK_TIMEOUT', err && err.name === 'WriteLockError' && err.code === 'LOCK_TIMEOUT', err && err.code);
    ok('B5 逾時約在設定時間（不會久等）', Date.now() - t0 < 600, Date.now() - t0);
    ok('B6 逾時的等待者已從佇列移除、連線已歸還', fake.locks.get('42').queue.length === 0);
    await holder.release();
    await holder.release();   // 重複釋放不應丟例外
    ok('B7 release 可重複呼叫', true);
    const h2 = await g.acquireLock({ waitMs: 500 });
    ok('B8 釋放後別人立刻取得', !!h2); await h2.release();

    // 伺服器端 idle 逾時：holder 一直不放（模擬實例當掉），DB 端終止連線 → 鎖自動釋放
    const stuck = await g.acquireLock({ waitMs: 500 });
    const t1 = Date.now();
    const next = await gB.acquireLock({ waitMs: 3000 });
    ok('B9 持有者閒置超過 idle_in_transaction_session_timeout → DB 終止其連線、鎖交給下一位', fake.stats.idleKills >= 1 && Date.now() - t1 < 2000, 'kills=' + fake.stats.idleKills);
    await stuck.release();   // 連線已被終止：不可丟例外
    ok('B10 被終止的持有者之後 release 不丟例外', true);
    await next.release();

    // 看門狗
    let forced = 0;
    const g2 = createPgWriteGuard({ getPool: () => pool, lockKey: '43', idleTxMs: 60000, onForcedRelease: () => { forced++; } });
    const wd = await g2.acquireLock({ waitMs: 500, holdMs: 100 });
    await sleep(250);
    ok('B11 看門狗：持有超過 holdMs 強制釋放', wd.forced === true && forced === 1 && fake.locks.get('43') === undefined);
    const after = await g2.acquireLock({ waitMs: 300 }); await after.release();
    ok('B12 強制釋放後別人可取得', true);

    // 連線池耗盡 → LOCK_UNAVAILABLE（不是掛住）
    const tiny = fake.pool({ max: 1, connectionTimeoutMillis: 80 });
    const g3 = createPgWriteGuard({ getPool: () => tiny, lockKey: '44' });
    const hold1 = await g3.acquireLock({ waitMs: 500 });
    let e3; try { await createPgWriteGuard({ getPool: () => tiny, lockKey: '44' }).acquireLock({ waitMs: 500 }); } catch (e) { e3 = e; }
    ok('B13 取不到連線 → LOCK_UNAVAILABLE（快速失敗）', e3 && e3.code === 'LOCK_UNAVAILABLE', e3 && e3.code);
    await hold1.release();
    // 同一個 guard（同行程）第二位：在行程內排隊，不佔用連線、不打到 DB（避免等鎖者占滿連線池造成持有者拿不到連線的死結）
    const roomy = fake.pool({ max: 3 });
    const gl = createPgWriteGuard({ getPool: () => roomy, lockKey: '45' });
    const waitsBefore = fake.stats.lockWaits;
    const lh = await gl.acquireLock({ waitMs: 500 });
    let e6; try { await gl.acquireLock({ waitMs: 150 }); } catch (e) { e6 = e; }
    ok('B16 同 guard 第二位：行程內排隊逾時 LOCK_TIMEOUT，沒有打到 DB 等鎖', e6 && e6.code === 'LOCK_TIMEOUT' && fake.stats.lockWaits === waitsBefore, e6 && e6.code);
    await lh.release();
    const lh2 = await gl.acquireLock({ waitMs: 500 }); await lh2.release();
    ok('B17 持有者放鎖後同 guard 可再取得', true);
    // 非法鎖鍵
    let e4; try { createPgWriteGuard({ getPool: () => pool, lockKey: '1; DROP TABLE x' }); } catch (e) { e4 = e; }
    ok('B14 lockKey 非整數被拒（避免注入）', !!e4);
    // 非法表名／schema
    let e5; try { createPgWriteGuard({ getPool: () => pool, schema: 'a"b' }).tableName('seq'); } catch (e) { e5 = e; }
    ok('B15 schema 識別字驗證', !!e5);
  }

  // ═══ C. 原子取號 ═══
  {
    const fake = new FakeDb({ latency: [0, 3] });
    const inst = [0, 1, 2, 3].map(() => freshPostgres(fake, { QUOTE_NO_ATOMIC: '1' }, { max: 3 }));
    const jobs = [];
    for (let i = 0; i < 200; i++) jobs.push(inst[i % 4].nextQuoteSeq(YM, 0, 'u' + i));
    const got = await Promise.all(jobs);
    const serials = got.map((g) => g.serial).sort((a, b) => a - b);
    ok('C1 200 並發（4 實例）取號：全部唯一', new Set(serials).size === 200);
    ok('C2 號碼連續 1..200（沒有跳號）', serials[0] === 1 && serials[199] === 200 && serials.every((s, i) => s === i + 1));
    ok('C3 單號格式 QU-YYYYMM-NNN', got.every((g) => /^QU-202610-\d{3}$/.test(g.quoteNo)) && got.find((g) => g.serial === 7).quoteNo === 'QU-202610-007');
    const issued = await inst[0].listIssuedQuoteNos(YM);
    ok('C4 已發號紀錄 200 筆、與取號結果一致', issued.length === 200 && issued.every((r, i) => r.serial === i + 1));
    // floor 銜接舊資料
    const f2 = await inst[1].nextQuoteSeq(YM, 250, 'legacy');
    ok('C5 blob 內已有更大號（250）→ 接續 251', f2.serial === 251 && f2.quoteNo === 'QU-202610-251');
    const f3 = await inst[2].nextQuoteSeq(YM, 10, 'stale');
    ok('C6 傳入較小的下界（舊快照）不會倒退 → 252', f3.serial === 252);
    const m2 = await inst[3].nextQuoteSeq('202611', 0, 'u');
    ok('C7 新月份從 1 開始（年月重置規則不變）', m2.quoteNo === 'QU-202611-001');
    const big = await inst[0].nextQuoteSeq('202612', 998, 'u');
    const big2 = await inst[0].nextQuoteSeq('202612', 0, 'u');
    const big3 = await inst[0].nextQuoteSeq('202612', 0, 'u');
    ok('C8 序號 ≥1000 不被截斷（999→1000→1001）', big.quoteNo === 'QU-202612-999' && big2.quoteNo === 'QU-202612-1000' && big3.quoteNo === 'QU-202612-1001', [big.quoteNo, big2.quoteNo, big3.quoteNo].join());
    ok('C9 格式與 QN.format 一致', QN.format('202612', 1000) === 'QU-202612-1000' && QN.format('202610', 7) === 'QU-202610-007');
    let off; try { await freshPostgres(fake, {}).nextQuoteSeq(YM, 0, 'x'); } catch (e) { off = e; }
    ok('C10 QUOTE_NO_ATOMIC 未開 → nextQuoteSeq 拒絕（不會誤用）', !!off);
    // floorOf 與 server.js genQuoteNo 定義一致
    const d = { quotations: [{ quoteNo: 'QU-202610-004' }, { quoteNo: 'QU-202609-099' }, { quoteNo: '' }], quoteSeq: { '202610': 2 } };
    ok('C11 floorOf＝max(quoteSeq, 現有最大號)', QN.floorOf(d, '202610') === 4 && QN.floorOf({ quotations: [], quoteSeq: { '202610': 9 } }, '202610') === 9 && QN.floorOf({}, '202610') === 0);
  }

  // ═══ D. 重現：多實例並發建立報價單 ═══
  const N = 60, M = 4;
  let reproSeed = null;
  for (const seed of [1, 2, 3]) {   // 問題是機率性的；只要有任何一輪重現就算證明（實測每輪都會）
    const fake = new FakeDb({ latency: [0, 3], seed });
    fake.seedMain({ contacts: [], quotations: [] });
    const insts = Array.from({ length: M }, () => mkInstance(fake, {}));
    const r = await runCreates(insts, false, N);
    await Promise.all(insts.map((i) => i.db.flush()));
    const s = summarize(fake, r);
    if (s.lost > 0 || s.dup > 0) { reproSeed = { seed, ...s }; break; }
  }
  ok('D1 【重現】開關全關：多實例並發建單 → 掉單或重號（問題存在）', !!reproSeed, reproSeed ? `seed=${reproSeed.seed} 回報成功 ${reproSeed.created} 張，DB 只剩 ${reproSeed.finalCount} 張（掉 ${reproSeed.lost}），重號 ${reproSeed.dup}` : '沒重現');

  for (const [label, flags, atomic] of [
    ['D2 (b) DB_WRITE_SERIALIZE=1', { DB_WRITE_SERIALIZE: '1' }, false],
    ['D3 (a)+(b) 兩個都開', { DB_WRITE_SERIALIZE: '1', QUOTE_NO_ATOMIC: '1' }, true],
  ]) {
    for (const [n, m] of [[50, 4], [200, 5]]) {
      const fake = new FakeDb({ latency: [0, 2], seed: 7 });
      fake.seedMain({ contacts: [], quotations: [] });
      const insts = Array.from({ length: m }, () => mkInstance(fake, flags, { waitMs: 60000, holdMs: 60000 }));
      const r = await runCreates(insts, atomic, n);
      await Promise.all(insts.map((i) => i.db.flush()));
      const s = summarize(fake, r);
      const sorted = s.nos.map((x) => parseInt(x.slice(-3), 10)).sort((a, b) => a - b);
      ok(`${label} N=${n}/${m}實例：全數成功、不掉單、不重號、號碼連續`, s.created === n && s.lost === 0 && s.dup === 0 && s.finalCount === n && sorted.every((x, i) => x === i + 1),
        `created=${s.created} final=${s.finalCount} dup=${s.dup} lockWaits=${fake.stats.lockWaits}`);
      ok(`${label} N=${n}：鎖都已釋放`, fake.locks.size === 0 && insts.every((i) => !i.guard.isHeld()));
    }
  }
  {   // (a) 單獨：不重號；掉單則可由「已發號紀錄」對帳
    let found = null;
    for (const seed of [1, 2, 3, 4]) {
      const fake = new FakeDb({ latency: [0, 3], seed });
      fake.seedMain({ contacts: [], quotations: [] });
      const insts = Array.from({ length: M }, () => mkInstance(fake, { QUOTE_NO_ATOMIC: '1' }));
      const r = await runCreates(insts, true, N);
      await Promise.all(insts.map((i) => i.db.flush()));
      const s = summarize(fake, r);
      const issued = await insts[0].db.listIssuedQuoteNos(YM);
      const rv = QL.review(JSON.parse(JSON.stringify(fake.main())), new Date().toISOString(), issued);
      const m = rv.months.find((x) => x.ym === YM);
      found = { s, issued: issued.length, unused: m.issuedUnused.length, gaps: m.gaps.length, seed };
      if (s.lost > 0) break;
    }
    ok('D4 (a) 單獨：單號 0 重複（即使仍可能掉單）', found.s.dup === 0 && found.issued === N, JSON.stringify(found.s.dup));
    ok('D5 (a) 單獨：被覆蓋掉的單 = 已發未用，全部對得上（無 GAP_UNEXPLAINED）', found.unused === found.s.lost && found.gaps === 0, `lost=${found.s.lost} issuedUnused=${found.unused} gaps=${found.gaps}`);
  }

  // ═══ E. 中介層 ═══
  {
    const fake = new FakeDb({ latency: [0, 1] });
    fake.seedMain({ n: 0 });
    const A = mkInstance(fake, { DB_WRITE_SERIALIZE: '1' }, { waitMs: 150 });
    const B = mkInstance(fake, { DB_WRITE_SERIALIZE: '1' }, { waitMs: 150 });
    // E1 鎖被別的實例佔住 → 逾時 503
    const blocker = await A.guard.acquire('blocker', { waitMs: 500 });
    let ran = false;
    const t0 = Date.now();
    const r1 = await request(B, 'POST', '/api/contacts', async () => { ran = true; });
    ok('E1 鎖被佔住 → 回 503 DB_BUSY', r1.res.statusCode === 503 && JSON.parse(r1.res.body).code === 'DB_BUSY', r1.res.statusCode);
    ok('E2 友善訊息＋可重試＋Retry-After', /尚未儲存/.test(JSON.parse(r1.res.body).error) && JSON.parse(r1.res.body).retryable === true && r1.res.headers['retry-after'] === '2');
    ok('E3 handler 沒有執行（沒有半套寫入）', ran === false && !r1.ran);
    ok('E4 等待時間受限（遠小於 Vercel 30s）', Date.now() - t0 < 1500, Date.now() - t0);
    // E5 同實例行程內排隊逾時
    const holdLocal = await B.guard.acquire('local', { waitMs: 500 }).catch((e) => e);
    // blocker 在 A 持有 DB 鎖，B.acquire 會卡在 DB；等它逾時
    ok('E5 行程內／DB 鎖逾時都丟 GuardError(DB_BUSY)', holdLocal instanceof WG.GuardError && holdLocal.code === 'DB_BUSY', holdLocal && holdLocal.code);
    await blocker.release();
    // E6 釋放後正常
    const r2 = await request(B, 'POST', '/api/contacts', async (req, res, db) => { const d = db.load(); d.n = (d.n || 0) + 1; db.save(d); res.status(200).json({ ok: 1 }); });
    ok('E6 鎖釋放後請求正常完成並寫入', r2.res.statusCode === 200 && fake.main().n === 1);
    // E7 豁免／未登入／純 GET 不拿鎖（鎖被佔時照樣通過）
    const blocker2 = await A.guard.acquire('blocker2', { waitMs: 500 });
    const ex = await request(B, 'POST', '/api/ai/visit-suggest', async (req, res) => { res.json({ ok: 1 }); });
    const an = await request(B, 'POST', '/api/contacts', async (req, res) => { res.json({ ok: 1 }); }, { anon: true });
    const gt = await request(B, 'GET', '/api/contacts', async (req, res) => { res.json({ ok: 1 }); });
    ok('E7 豁免路由／未登入／純 GET 不受鎖影響', ex.ran && an.ran && gt.ran && ex.res.statusCode === 200 && an.res.statusCode === 200 && gt.res.statusCode === 200);
    const gw = await request(B, 'GET', '/api/callins', async () => {});
    ok('E8 會寫入的 GET（callins）要鎖 → 鎖被佔時 503', gw.res.statusCode === 503);
    await blocker2.release();
    // E9 鎖內 handler 看到最新資料（別的實例剛寫入的）
    const w = await request(A, 'POST', '/api/contacts', async (req, res, db) => { const d = db.load(); d.fromA = 'x'; db.save(d); res.json({ ok: 1 }); });
    let seen;
    await request(B, 'POST', '/api/contacts', async (req, res, db) => { seen = db.load().fromA; res.json({ ok: 1 }); });
    ok('E9 持鎖後 handler 讀到另一實例剛提交的資料', seen === 'x');
    // E10 存檔失敗 → 503 DB_WRITE_FAILED、釋放鎖、快取失效
    fake.failNext(/^INSERT INTO app_data \(id, content, updated_at\) VALUES \('main'/, 1);
    const bad = await request(B, 'POST', '/api/contacts', async (req, res, db) => { const d = db.load(); d.lost = 'zzz'; db.save(d); res.status(201).json({ ok: 'created' }); });
    ok('E10 存檔失敗 → 不回 2xx，改回 503 DB_WRITE_FAILED（友善、可重試）', bad.res.statusCode === 503 && JSON.parse(bad.res.body).code === 'DB_WRITE_FAILED' && /可能沒有生效/.test(JSON.parse(bad.res.body).error), bad.res.statusCode + ' ' + bad.res.body);
    ok('E11 失敗後鎖已釋放', !B.guard.isHeld() && fake.locks.size === 0);
    let after;
    await request(A, 'POST', '/api/contacts', async (req, res, db) => { after = db.load(); res.json({ ok: 1 }); });
    ok('E12 失敗的變更沒有進 DB；別的實例看不到', fake.main().lost === undefined && after.lost === undefined);
    let afterB;
    await request(B, 'POST', '/api/contacts', async (req, res, db) => { afterB = db.load(); res.json({ ok: 1 }); });
    ok('E13 失敗實例的記憶體快取已失效並重抓（不殘留未存成功的變更）', afterB.lost === undefined, JSON.stringify(afterB));
    // E14 持鎖後 ready 失敗
    const C = mkInstance(fake, { DB_WRITE_SERIALIZE: '1' }, { waitMs: 300 });
    await C.db.ready();
    const origReady = C.db.ready;
    const realWrapped = WG.createWriteGuard({ db: Object.assign({}, C.db, { ready: async () => { throw new Error('db down'); } }), config: C.cfg, log: quiet });
    const req = { method: 'POST', baseUrl: '/api', path: '/contacts', session: { user: { username: 'u' } } }; const res = mkRes();
    realWrapped.middleware()(req, res, () => {});
    await res.done;
    ok('E14 持鎖後 ready 失敗 → 503 且釋放鎖', res.statusCode === 503 && !realWrapped.isHeld());
  }

  // ═══ F. 開關關閉 ═══
  {
    const fake = new FakeDb({ latency: [0, 0] });
    fake.seedMain({ n: 0 });
    const off = mkInstance(fake, {});
    ok('F1 預設 serialize=false', off.cfg.serialize === false);
    const before = fake.log.length;
    const r = await request(off, 'POST', '/api/contacts', async (req, res, db) => { const d = db.load(); d.n++; db.save(d); res.json({ ok: 1 }); });
    const sqls = fake.log.slice(before);
    ok('F2 關閉時沒有任何鎖／交易／新表 SQL', !sqls.some((q) => /advisory|BEGIN|SET LOCAL|quote_no_/.test(q)), sqls.join(' | '));
    ok('F3 關閉時 middleware 直接放行、不包 res.end', r.ran && r.res.statusCode === 200 && !off.guard.isHeld() && fake.stats.lockWaits === 0);
    ok('F5 關閉時沒有建立 quote_no_* 表（fake 內沒有資料）', fake.seqTable.size === 0 && fake.issued.size === 0);
    let ex; try { await off.db.nextQuoteSeq(YM, 0, 'x'); } catch (e) { ex = e; }
    ok('F6 關閉時 nextQuoteSeq 拒絕', !!ex);
    const wl = await off.guard.withLock('x', async () => 'ran');
    ok('F7 關閉時 withLock 直接執行 fn', wl === 'ran');
    // F8 平價測試：開關全關時，db/postgres.js 對一段固定操作送出的 SQL 序列必須與 origin/main（本功能加入前）完全相同
    {
      const GOLDEN = [
        "CREATE TABLE IF NOT EXISTS app_data ( id TEXT PRIMARY KEY, content JSONB NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT now() );",
        "SELECT content, updated_at FROM app_data WHERE id = 'main'",
        "INSERT INTO app_data (id, content, updated_at) VALUES ('main', $1, now()) ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now() RETURNING updated_at",
        "SELECT updated_at FROM app_data WHERE id = 'main'",
        "INSERT INTO app_data (id, content, updated_at) VALUES ('main', $1, now()) ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now() RETURNING updated_at",
        "SELECT updated_at FROM app_data WHERE id = 'main'",
        "SELECT content, updated_at FROM app_data WHERE id = 'main'",
        "INSERT INTO app_data (id, content, updated_at) VALUES ('main', $1, now()) ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now() RETURNING updated_at",
        "SELECT content FROM app_data WHERE id = 'audit_log'",
        "INSERT INTO app_data (id, content, updated_at) VALUES ('audit_log', $1, now()) ON CONFLICT (id) DO UPDATE SET content = EXCLUDED.content, updated_at = now()"
      ];
      const f = new FakeDb({ latency: [0, 0] });
      f.seedMain({ contacts: [], groups: [], monthlyBudgets: [] });
      const m = freshPostgres(f, {});
      await m.ready();
      const d = m.load(); d.contacts = [{ id: 1 }]; m.save(d); await m.flush();
      await m.ready();
      const d2 = m.load(); d2.contacts.push({ id: 2 }); m.save(d2); await m.flush();
      f.appData.set('main', { content: JSON.stringify({ contacts: [{ id: 9 }] }), updated_at: f.now() });
      await m.ready();
      const d3 = m.load(); d3.x = 1; m.save(d3);
      await m.appendAuditLog({ id: 'a1', timestamp: new Date(0).toISOString(), action: 'T' });
      await m.flush();
      ok('F8 【平價】開關全關：SQL 序列與 origin/main 逐條相同（共 ' + GOLDEN.length + ' 條）', JSON.stringify(f.log) === JSON.stringify(GOLDEN), f.log.length + ' 條');
      ok('F9 開關全關：沒有任何世代號／鎖／交易／新表 SQL', !f.log.some((q) => /advisory|BEGIN|SET LOCAL|quote_no_|to_regclass/.test(q)));
    }
  }

  // ═══ G. quoteLedger.review 相容 ═══
  {
    const mk = () => ({ quotations: [{ id: '1', quoteNo: 'QU-202610-001', owner: 'a', createdAt: 't' }, { id: '3', quoteNo: 'QU-202610-003', owner: 'a', createdAt: 't' }], quoteSeq: { '202610': 3 } });
    const base = QL.review(mk(), 'now');
    const baseExplicit = QL.review(mk(), 'now', undefined);
    ok('G1 不傳 issued：輸出與舊版相同（無 issuedUnused 欄位）', JSON.stringify(base) === JSON.stringify(baseExplicit) && !('issuedUnused' in base.summary) && !('issuedUnused' in base.months[0]));
    ok('G2 舊行為：002 是 GAP_UNEXPLAINED', base.months[0].gaps.length === 1 && base.months[0].gaps[0].no === 'QU-202610-002');
    const withIssued = QL.review(mk(), 'now', [{ quoteNo: 'QU-202610-001' }, { quoteNo: 'QU-202610-002', issuedTo: 'u', issuedAt: 'x' }, { quoteNo: 'QU-202610-003' }]);
    ok('G3 傳 issued：002 變成已發未用、不再是 GAP', withIssued.months[0].gaps.length === 0 && withIssued.months[0].issuedUnused.length === 1 && withIssued.months[0].issuedUnused[0].no === 'QU-202610-002' && withIssued.summary.issuedUnused === 1 && withIssued.summary.gapUnexplained === 0);
    const beyond = QL.review(mk(), 'now', [{ quoteNo: 'QU-202610-005' }]);
    ok('G4 已發號大於現有最大號：004 仍是缺號、005 已發未用', beyond.months[0].gaps.map((g) => g.no).join() === 'QU-202610-002,QU-202610-004' && beyond.months[0].issuedUnused.length === 1);
  }

  // ═══ H. ready() 重抓 vs save() 競態（審查 #1）：無鎖請求的並行 ready() 不可把鎖內寫入者剛存的內容蓋回舊的 ═══
  {
    const SER = { DB_WRITE_SERIALIZE: '1' };
    const writer = (key) => async (req, res, db) => { const d = db.load(); d[key] = 1; db.save(d); res.status(200).json({ ok: key }); };
    // H1 確定性重現：R1 的 SELECT 在 W 的 save() 之前就在資料庫端執行完（看到舊資料），回應拖到 W 的 save() 之後、W 的 UPDATE 回應之前才到
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1 });
      const X = mkInstance(fake, SER, { waitMs: 5000, holdMs: 20000 });
      const Y = mkInstance(fake, SER, { waitMs: 5000, holdMs: 20000 });
      await X.db.ready();
      await request(Y, 'POST', '/api/contacts', writer('fromY'));
      const gate1 = fake.holdNext(/^SELECT content, updated_at FROM app_data WHERE id = 'main'/);
      const r1 = X.db.ready();                       // R1：無鎖的一般請求，快取過期 → 重抓；回應被扣住
      await gate1.reached;
      const gate2 = fake.holdNext(/^INSERT INTO app_data \(id, content, updated_at\) VALUES \('main'/);   // 扣住 W 的 UPDATE 回應（資料庫端已提交）
      const w = await request(X, 'POST', '/api/contacts', async (req, res, db) => {
        const d = db.load(); d.fromW = 1; db.save(d);
        gate1.release();                             // R1 的（舊）內容此刻才回到實例
        await sleep(40);
        gate2.release();                             // W 的 saveAsync 才完成（把 _lastUpdatedAtMs 設成最新）
        res.status(200).json({ ok: 1 });
      });
      await r1;
      await request(X, 'POST', '/api/contacts', writer('fromW2'));
      await Promise.all([X.db.flush(), Y.db.flush()]);
      const fin = fake.main();
      ok('H1 【確定性重現】無鎖 ready() 的舊抓取在 W 存檔後才回來 → W 與 W2 的寫入都不遺失', w.res.statusCode === 200 && fin.fromW === 1 && fin.fromW2 === 1 && fin.fromY === 1, JSON.stringify(fin));
    }
    // H3 世代號的獨立證明：抓回的時間戳比快取新（鎖外的豁免路由先提交了 v2），但那份內容沒有本實例剛 save() 的變更 → 必須丟棄
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1 });
      const X = mkInstance(fake, SER, { waitMs: 5000, holdMs: 20000 });
      await X.db.ready();                                   // X 快取＝v1
      fake.appData.set('main', { content: JSON.stringify({ base: 1, other: 1 }), updated_at: fake.now() });   // 另一個實例（鎖外）提交了 v2 → X 的快取過期
      const gate = fake.holdNext(/^SELECT content, updated_at FROM app_data WHERE id = 'main'/);
      const r1 = X.db.ready();                              // 無鎖 ready：過期 → 重抓；SELECT 在資料庫端看到 v2（沒有 W），回應被扣住
      await gate.reached;
      const d = X.db.load(); d.fromW = 1; X.db.save(d);     // 此刻本實例 save()（存檔世代 +1）
      gate.release();
      await r1;
      await X.db.flush();
      ok('H3 抓取期間本實例有 save() → 即使抓回的時間戳較新也丟棄，快取仍含 W 的變更', X.db.load().fromW === 1, JSON.stringify(X.db.load()));
    }
    // H2 隨機交錯壓力：200 回合，每回合 Y 先寫、再讓 X 上的 R（無鎖 ready）／W／W2 與 Y 上的 W3 以 0~100ms 隨機延遲起跑；所有回報成功的寫入都必須在最終 DB 裡
    {
      let bad = 0, firstBad = null, timeouts = 0;
      const one = async (seed) => {
        const fake = new FakeDb({ latency: [0, 6], seed, slow: { re: /^SELECT content, updated_at/, ms: [0, 80] } });   // 大 SELECT 的傳輸時間比小查詢長（資料庫端看到執行當下的資料）
        fake.seedMain({ base: 1 });
        const X = mkInstance(fake, SER, { waitMs: 20000, holdMs: 60000 });
        const Y = mkInstance(fake, SER, { waitMs: 20000, holdMs: 60000 });
        await X.db.ready(); await Y.db.ready();
        const rnd = (() => { let a = seed * 7919 + 13; return () => { a = (a * 1103515245 + 12345) & 0x7fffffff; return a / 0x7fffffff; }; })();
        const jit = () => sleep(Math.floor(rnd() * 100));
        const written = [];
        const w = (inst, key) => (async () => { await jit(); const r = await request(inst, 'POST', '/api/contacts', writer(key), { timeoutMs: 20000 }); if (r.timedOut) timeouts++; if (r.res.statusCode === 200) written.push(key); })();
        await request(Y, 'POST', '/api/contacts', writer('y0'), { timeoutMs: 20000 }); written.push('y0');
        const chain = (inst, k1, k2) => (async () => { await w(inst, k1); await w(inst, k2); })();   // 先後兩次寫入：第二次依賴第一次留下的快取／時間戳
        await Promise.all([
          (async () => { await jit(); await X.db.ready(); })(),
          (async () => { await jit(); await X.db.ready(); })(),
          (async () => { await jit(); await Y.db.ready(); })(),
          chain(X, 'w1', 'w2'), chain(Y, 'w3', 'w4'), chain(X, 'w5', 'w6'),
        ]);
        await Promise.all([X.db.flush(), Y.db.flush()]);
        const fin = fake.main();
        const missing = written.filter((k) => fin[k] !== 1);
        if (missing.length) { bad++; if (!firstBad) firstBad = 'seed=' + seed + ' 遺失 ' + missing.join(','); }
      };
      const seeds = Array.from({ length: 200 }, (_, i) => i + 1);
      for (let i = 0; i < seeds.length; i += 25) await Promise.all(seeds.slice(i, i + 25).map(one));
      ok('H2 200 回合隨機交錯（0~100ms）：回報成功的寫入全部留在最終 DB（0 遺失）', bad === 0 && timeouts === 0, bad + ' 回合遺失；' + (firstBad || '') + '；逾時 ' + timeouts);
    }
  }

  // ═══ I. COMMIT／BEGIN 用戶端逾時（審查 #2）＋建表（審查 #16／#29） ═══
  const SER = { DB_WRITE_SERIALIZE: '1' };
  const MAIN_UPDATE_RE = /^INSERT INTO app_data \(id, content, updated_at\) VALUES \('main'/;
  const writerOf = (key) => async (req, res, db) => { const d = db.load(); d[key] = 1; db.save(d); res.status(200).json({ ok: key }); };
  {
    const fake = new FakeDb({ latency: [0, 1] });
    const pool = fake.pool({ max: 4 });
    const mk = (key, extra) => createPgWriteGuard(Object.assign({ getPool: () => pool, lockKey: key, idleTxMs: 60000, commitTimeoutMs: 150, beginSlackMs: 150 }, extra || {}));
    const g = mk('51');
    const h = await g.acquireLock({ waitMs: 2000 });
    fake.hangNext(/^COMMIT$/, 1);
    const t0 = Date.now();
    const relP = h.release();
    const h2P = g.acquireLock({ waitMs: 3000 });   // release() 先同步放行程內互斥 → 第二位立刻開始取 DB 鎖
    const relRes = await within(relP, 3000);
    const relMs = Date.now() - t0;
    ok('I1 COMMIT 永不回應：release() 在用戶端逾時內結束（不永遠卡住）', relRes !== 'TIMEOUT' && relMs < 1000, relRes === 'TIMEOUT' ? '卡住' : relMs);
    const h2 = await within(h2P, 4000);
    ok('I2 逾時後連線被銷毀 → 伺服器端鎖釋放 → 下一位拿到鎖', !!h2 && h2 !== 'TIMEOUT' && Date.now() - t0 < 2000, h2 === 'TIMEOUT' ? '卡住' : Date.now() - t0);
    if (h2 && h2 !== 'TIMEOUT') await h2.release();
    ok('I3 沒有殘留的鎖', fake.locks.size === 0);

    const g4 = mk('52');
    fake.hangNext(/^BEGIN;/, 1);
    const t4 = Date.now();
    let e4;
    let r4; try { r4 = await within(g4.acquireLock({ waitMs: 200 }), 4000); } catch (e) { e4 = e; }
    ok('I4 取鎖的 BEGIN 永不回應：用戶端逾時 → LOCK_TIMEOUT，不永遠卡住', r4 !== 'TIMEOUT' && e4 && e4.code === 'LOCK_TIMEOUT' && Date.now() - t4 < 1500, r4 === 'TIMEOUT' ? '卡住' : (e4 && e4.code) + ' ' + (Date.now() - t4));
    const h5 = await within(g4.acquireLock({ waitMs: 1000 }), 4000);
    ok('I5 BEGIN 逾時後行程內互斥已釋放，下一位可取得鎖', !!h5 && h5 !== 'TIMEOUT');
    if (h5 && h5 !== 'TIMEOUT') await h5.release();

    const g6 = mk('53');
    const h6 = await g6.acquireLock({ waitMs: 500, holdMs: 2000 });
    await h6.release();
    const begins = fake.log.filter((q) => q.startsWith('BEGIN; SET LOCAL lock_timeout'));
    ok('I6 取鎖 SQL 帶 lock_timeout／idle_in_transaction_session_timeout／statement_timeout／advisory 鎖（來源＝lib 匯出的 SQL.lockBegin）',
      begins.length > 0 && begins.every((q) => /^BEGIN; SET LOCAL lock_timeout = [0-9]+; SET LOCAL idle_in_transaction_session_timeout = [0-9]+; SET LOCAL statement_timeout = [0-9]+; SELECT pg_advisory_xact_lock\([0-9]+\)$/.test(q)) && typeof SQL.lockBegin === 'function');
    ok('I7 鎖交易的閒置保底不超過「持有上限＋3 秒」（holdMs=2000 → 5000）', begins.some((q) => q.includes('idle_in_transaction_session_timeout = 5000')), begins[begins.length - 1]);

    // 端到端：持鎖請求的 COMMIT 卡住 → 回應仍會送出；另一個實例之後拿得到鎖
    const f2 = new FakeDb({ latency: [0, 1] });
    f2.seedMain({ n: 0 });
    const A = mkInstance(f2, SER, { waitMs: 4000, holdMs: 20000 });
    const B = mkInstance(f2, SER, { waitMs: 8000, holdMs: 20000 });
    f2.hangNext(/^COMMIT$/, 1);
    const t8 = Date.now();
    const ra = await request(A, 'POST', '/api/contacts', writerOf('fromA'), { timeoutMs: 15000 });
    ok('I8 持鎖請求的 COMMIT 卡住：回應仍在用戶端逾時（約 2 秒）後送出，不卡到平台強殺', ra.res.statusCode === 200 && Date.now() - t8 < 8000, ra.res.statusCode + ' ' + (Date.now() - t8) + 'ms');
    const rb = await request(B, 'POST', '/api/contacts', writerOf('fromB'), { timeoutMs: 15000 });
    ok('I9 之後另一個實例能取得鎖並正常寫入，A 的寫入沒有遺失', rb.res.statusCode === 200 && f2.main().fromA === 1 && f2.main().fromB === 1);

    // 建表：兩個實例並發第一次取號
    const f3 = new FakeDb({ latency: [0, 2] });
    const i1 = freshPostgres(f3, { QUOTE_NO_ATOMIC: '1' }), i2 = freshPostgres(f3, { QUOTE_NO_ATOMIC: '1' });
    const [n1, n2] = await Promise.all([i1.nextQuoteSeq(YM, 0, 'a'), i2.nextQuoteSeq(YM, 0, 'b')]);
    ok('I10 兩個實例並發第一次取號：表只建一次（DDL 共 4 條）、兩表都開了 RLS、號碼不重複', f3.stats.ddl === 4 && f3.defaultRls.size === 2 && n1.serial !== n2.serial, 'ddl=' + f3.stats.ddl);
    const ddlBefore = f3.stats.ddl;
    const i3 = freshPostgres(f3, { QUOTE_NO_ATOMIC: '1' });
    await i3.nextQuoteSeq(YM, 0, 'c');
    ok('I11 表已存在：之後的冷啟動實例不再執行任何 DDL', f3.stats.ddl === ddlBefore);
    const nq = f3.log.filter((q) => q.startsWith('WITH c AS'));
    ok('I12 取號 SQL 含 GREATEST／ON CONFLICT (ym) DO UPDATE／RETURNING serial, quote_no／lpad（來源＝lib 匯出的 SQL.nextQuoteSeq）',
      nq.length >= 3 && nq.every((q) => q.includes('GREATEST(') && q.includes('ON CONFLICT (ym) DO UPDATE') && q.includes('RETURNING serial, quote_no') && q.includes("lpad(c.last_serial::text, 3, '0')")) && typeof SQL.nextQuoteSeq === 'function');
  }

  // ═══ J. 存檔失敗只算「本請求自己的」；UNLOCKED_WRITE 以請求為單位（審查 #11／#13／#24／#30／#36） ═══
  {
    const warns = [];
    const logCap = { log() {}, info() {}, error() {}, warn(...a) { warns.push(a.join(' ')); } };
    const fake = new FakeDb({ latency: [0, 1] });
    fake.seedMain({ n: 0 });
    const A = mkInstance(fake, SER, { waitMs: 3000 }, logCap);
    const r = await request(A, 'POST', '/api/contacts', async (req, res, db) => {
      fake.failNext(MAIN_UPDATE_RE, 1);   // 下一次主資料存檔會失敗——由下面這個「豁免路由」的存檔吃掉
      await request(A, 'POST', '/api/ai/visit-suggest', async (rq, rs, d) => { const x = d.load(); x.exempt = 1; d.save(x); rs.json({ ok: 1 }); });
      await sleep(40);
      const d = db.load(); d.a = 1; db.save(d);
      res.status(200).json({ ok: 1 });
    });
    ok('J1 持鎖期間「無關的豁免路由」存檔失敗 → 本請求仍回 200（不誤判為 DB_WRITE_FAILED）', r.res.statusCode === 200, r.res.statusCode + ' ' + r.res.body);
    ok('J2 本請求自己的變更確實存進 DB', fake.main().a === 1);
    ok('J3 豁免路由的鎖外存檔留下 UNLOCKED_WRITE 警告（即使當時有別的請求持鎖也不被抑制）', warns.some((w) => w.includes('UNLOCKED_WRITE') && w.includes('/api/ai/visit-suggest')), warns.join(' | ').slice(0, 300));
    ok('J4 持鎖者自己的存檔不產生 UNLOCKED_WRITE', !warns.some((w) => w.includes('UNLOCKED_WRITE') && w.includes('/api/contacts')));
    // J5 本請求自己的存檔失敗 → 503
    fake.failNext(MAIN_UPDATE_RE, 1);
    const bad = await request(A, 'POST', '/api/contacts', writerOf('lostB'));
    ok('J5 本請求自己的存檔失敗 → 503 DB_WRITE_FAILED', bad.res.statusCode === 503 && JSON.parse(bad.res.body).code === 'DB_WRITE_FAILED');
    // J6 鎖釋放後才呼叫的 db.save（例如背景推播清理）→ UNLOCKED_WRITE，且不影響任何請求的回應
    warns.length = 0;
    let bgSave;
    const okReq = await request(A, 'POST', '/api/contacts', async (req, res, db) => {
      bgSave = (async () => { await sleep(80); const d = db.load(); d.bg = 1; db.save(d); })();   // 在請求結束、鎖放掉之後才存檔
      res.status(200).json({ ok: 1 });
    });
    await bgSave;
    ok('J6 請求結束後才發出的背景存檔 → UNLOCKED_WRITE 警告，請求本身回 200', okReq.res.statusCode === 200 && warns.some((w) => w.includes('UNLOCKED_WRITE')), warns.join(' | ').slice(0, 200));
    await A.db.flush();
  }

  // ═══ K. 分類器：HEAD／大小寫／尾斜線／精確豁免（審查 #18／#21） ═══
  {
    const cfg = WG.readConfig({ DB_WRITE_SERIALIZE: '1' });
    const need = (m, p) => WG.needsLock(cfg, m, p);
    ok('K1 HEAD 視同 GET：HEAD /api/callins 要鎖', need('HEAD', '/api/callins') && !need('HEAD', '/api/contacts'));
    ok('K2 大小寫不敏感：GET /api/Callins、/api/Admin/Quote-Ledger 要鎖', need('GET', '/api/Callins') && need('GET', '/api/Admin/Quote-Ledger'));
    ok('K3 尾斜線：GET /api/callins/ 要鎖', need('GET', '/api/callins/') && need('GET', '/api/callins//'));
    ok('K4 豁免比對同樣正規化：POST /api/AI/Opp-Win-Rate、/api/ai/x/ 不拿鎖', !need('POST', '/api/AI/Opp-Win-Rate') && !need('POST', '/api/ai/x/'));
    ok('K5 push 只豁免 test：subscribe／unsubscribe 要鎖（純 DB 的同步寫入）', !need('POST', '/api/push/test') && need('POST', '/api/push/subscribe') && need('POST', '/api/push/unsubscribe'));
    ok('K6 integrations 只豁免 test／push/batch：mappings 要鎖', !need('POST', '/api/admin/integrations/test') && !need('POST', '/api/admin/integrations/push/batch') && need('POST', '/api/admin/integrations/mappings'));
    ok('K7 /api/upload 豁免（只存檔不寫 blob）；寄件匣 requeue 豁免；email-import/apply 要鎖', !need('POST', '/api/upload') && !need('POST', '/api/admin/mail/outbox/abc/requeue') && need('POST', '/api/admin/users/email-import/apply'));
    ok('K8 不誤殺相似路徑', need('POST', '/api/uploads') && need('POST', '/api/pushx/test') && need('POST', '/api/admin/mail/outboxes/1'));
    ok('K9 持有上限有硬上限（不超過 DB 保底）', WG.readConfig({ DB_WRITE_SERIALIZE: '1', DB_WRITE_LOCK_HOLD_MS: '99999' }).holdMs === WG.MAX_HOLD_MS);
  }

  // ═══ L. 鎖範圍：multipart 讀完 body 才取鎖、被拒絕的請求不排隊、requeue 不占鎖（審查 #7／#9／#20／#25） ═══
  {
    const fake = new FakeDb({ latency: [0, 1] });
    fake.seedMain({ n: 0 });
    const A = mkInstance(fake, SER, { waitMs: 200, holdMs: 20000 });
    const B = mkInstance(fake, SER, { waitMs: 200, holdMs: 20000 }, undefined, { preDeny: (req, full) => full.startsWith('/api/admin/') && req.session.user.username !== 'root' });
    const MP = { 'content-type': 'multipart/form-data; boundary=xx' };
    // L1 上傳 body 讀取期間不持鎖（isHeld 為 false）；讀完後才取鎖並寫入
    let heldDuringUpload = null;
    const up = await request(A, 'POST', '/api/admin/import-contacts', writerOf('imported'), { headers: MP, bodyMs: 150, onBodyReading: () => { setTimeout(() => { heldDuringUpload = A.guard.isHeld(); }, 50); } });
    ok('L1 multipart：body 讀取期間不持有寫入鎖；讀完後取鎖、寫入成功', heldDuringUpload === false && up.res.statusCode === 200 && fake.main().imported === 1, 'held=' + heldDuringUpload);
    // L2 別人持鎖時，慢速上傳不受影響（讀 body 不排隊），只有進 lateLock 才等
    const blocker = await B.guard.acquire('blocker', { waitMs: 500 });
    const t0 = Date.now();
    const up2 = await request(A, 'POST', '/api/admin/import-contacts', writerOf('x'), { headers: MP, bodyMs: 100 });
    ok('L2 鎖被占住：multipart 讀完 body 後取鎖逾時 → 503 DB_BUSY（沒有半套寫入）', up2.res.statusCode === 503 && JSON.parse(up2.res.body).code === 'DB_BUSY' && fake.main().x === undefined);
    // L3 /api/upload 豁免：鎖被占住也照樣完成（不占鎖）
    const upl = await request(A, 'POST', '/api/upload', async (req, res) => { res.json({ url: 'u' }); }, { headers: MP, bodyMs: 50 });
    ok('L3 /api/upload（豁免）：鎖被占住時照常完成', upl.res.statusCode === 200);
    // L4 預先排除：非管理員打 /api/admin/*（一定被 requireAdmin 拒絕）不排隊拿鎖 → 直接放行到 handler（handler 內才被拒）
    let ranHandler = false;
    const den = await request(B, 'POST', '/api/admin/users', async (req, res) => { ranHandler = true; res.status(403).json({ error: '無管理者權限' }); });
    ok('L4 preDeny：一定會被拒絕的請求不排隊拿鎖（鎖被占住也立刻到 handler 被拒）', ranHandler && den.res.statusCode === 403 && Date.now() - t0 < 2000);
    await blocker.release();
    // L5 持鎖者的 res 結束後鎖已放
    ok('L5 全部結束後沒有殘留的鎖', !A.guard.isHeld() && !B.guard.isHeld() && fake.locks.size === 0);
  }

  // ═══ N. 靜態接線檢查（讀原始碼；測試程式碼無法替代「真的有掛上」）（審查 #25／#34） ═══
  {
    const root = path.join(__dirname, '..');
    const src = fsx.readFileSync(path.join(root, 'server.js'), 'utf8');
    const iMw = src.indexOf("app.use('/api/', writeGuard.middleware());");
    const after = (needle) => { const i = src.indexOf(needle); return i; };
    ok('N1 server.js 掛了 app.use(\'/api/\', writeGuard.middleware())（只有一處）', iMw > 0 && src.indexOf("app.use('/api/', writeGuard.middleware());", iMw + 1) < 0);
    ok('N2 中介層掛在 apiLimiter／planGuard／強制改密碼／集團角色白名單之後、登入路由之前（被拒絕的請求不排隊拿鎖）',
      iMw > after("app.use('/api/', apiLimiter);") && iMw > after("app.use('/api/', planGuard);") && iMw > after('PWCHG_ALLOWED_PATHS.has(req.path)') && iMw > after('GROUPSALES_ALLOWED_PREFIXES.some') && iMw < after("app.post('/api/login', loginLimiter"));
    // 每個 uploadImport.single(...) 之後（同一條路由）都要有 writeLateLock
    let missingLate = 0, nImport = 0, from = 0; const badPaths = [];
    for (;;) {
      const i = src.indexOf('uploadImport.single(', from);
      if (i < 0) break;
      from = i + 1; nImport++;
      // 同一條路由裡要有 writeLateLock('<路徑>')，且路徑必須與 app.post('<路徑>' 完全一致（註冊的路徑錯了＝multipart 請求不會被延後取鎖，上傳會在 body 讀完前就排隊占鎖）
      const routeHead = src.lastIndexOf('app.post(', i);
      const rp = /app\.post\('([^']+)'/.exec(src.slice(routeHead, routeHead + 120));
      const lm = /writeLateLock\('([^']+)'\)/.exec(src.slice(i, i + 200));
      if (!lm || !rp || lm[1] !== rp[1]) { missingLate++; badPaths.push(rp ? rp[1] : '?'); }
    }
    ok('N3 所有 uploadImport.single(...) 路由（' + nImport + ' 條）都接了 writeLateLock(<與路由相同的路徑>)（multipart 讀完 body 才取鎖）', nImport >= 7 && missingLate === 0, 'missing=' + missingLate + ' ' + badPaths.join());
    const count = (text, needle) => text.split(needle).length - 1;
    const qr = fsx.readFileSync(path.join(root, 'lib', 'quoteRoutes.js'), 'utf8');
    ok('N4 multer 實例數量固定（server.js 4、quoteRoutes 1）：新增上傳路由時這條會失敗，提醒你決定「豁免」或「接 lateLock」', count(src, ' = multer(') === 4 && count(qr, ' = multer(') === 1, count(src, ' = multer(') + '/' + count(qr, ' = multer('));
    const cfg = WG.readConfig({ DB_WRITE_SERIALIZE: '1' });
    ok('N5 其餘 multer 路由（AI OCR、牌價簿預覽、名片圖上傳）都在豁免清單', ['/api/admin/ai-ocr-card', '/api/ai/ocr-card', '/api/admin/quote-pricebook/import-preview', '/api/upload'].every((p) => !WG.needsLock(cfg, 'POST', p)));
    const pg = fsx.readFileSync(path.join(root, 'db', 'postgres.js'), 'utf8');
    ok('N6 連線池的 query_timeout 只在 DB_WRITE_SERIALIZE=1 時加（關閉時連線池設定與原版相同）', pg.includes('...(SERIALIZE_ON ? { query_timeout:'));
    ok('N7 重抓／世代號邏輯只在 SERIALIZE_ON 分支（關閉時 ready() 仍是 _cache = await loadAsync()）', pg.includes('if (SERIALIZE_ON) {') && pg.includes('_cache = await loadAsync();') && pg.includes('if (!(await _refetchGuarded())) return;'));
    const dbIndex = fsx.readFileSync(path.join(root, 'db', 'index.js'), 'utf8');
    ok('N8 db/index.js 匯出 setSaveHook／nextQuoteSeq／acquireDbWriteLock／flushMain', ['setSaveHook', 'nextQuoteSeq', 'acquireDbWriteLock', 'flushMain', 'invalidateCache'].every((k) => dbIndex.includes(k)));
    ok('N10 SAP push/batch：開始時併入上次沒寫成功的對照（mergePendingInto）；寫回失敗時回應帶 saved:false＋warning（不是默默回成功）', src.includes('sapWriteback.mergePendingInto(data)') && src.includes('res.json(sapSaved') && src.includes('saved: false, warning: SAP_WRITEBACK_WARNING'));
    ok('N9 啟動遷移在開關開啟時用短鎖（等 1.5 秒、持有 5 秒）並可略過', src.includes("{ waitMs: 1500, holdMs: 5000 }") && src.includes('略過（取不到寫入鎖'));
  }

  // ═══ O. 外部呼叫之後的寫回：applyOnLatest／AI 點數短鎖／推播清理（審查 #3／#4／#5／#6／#19） ═══
  {
    const vm = require('vm');
    const root = path.join(__dirname, '..');
    const src = fsx.readFileSync(path.join(root, 'server.js'), 'utf8');
    const cut = (startMarker) => {
      const a = src.indexOf(startMarker);
      if (a < 0) throw new Error('server.js 找不到 ' + startMarker);
      const b = src.indexOf('\n}\n', a);
      return src.slice(a, b + 3);
    };
    // applyOnLatest 的實作在 lib/writeGuard.js（server.js 只是委派）：用真的 guard＋記錄呼叫順序的 stub db 測
    const mkStub = (serialize) => {
      const calls = [];
      const latest = { opportunities: [{ id: 'o1', aiWinRate: 1 }] };
      const stale = { opportunities: [{ id: 'o1', aiWinRate: 0 }] };
      const db = {
        ready: async () => { calls.push('ready'); },
        load: () => latest,
        save: (d) => { calls.push(d === latest ? 'save(latest)' : d === stale ? 'save(stale)' : 'save(?)'); },
        flushMain: async () => { calls.push('flushMain'); },
        flush: async () => { calls.push('flush'); },
      };
      const warns = [];
      const guard = WG.createWriteGuard({ db, config: WG.readConfig(serialize ? { DB_WRITE_SERIALIZE: '1' } : {}), log: { log() {}, info() {}, error() {}, warn: (...x) => warns.push(x.join(' ')) } });
      return { guard, calls, latest, stale, warns };
    };
    // O1 開關關閉：維持舊行為——直接改 await 前的物件並整份存回，不取鎖、不 ready，回傳 true
    {
      const c = mkStub(false);
      const r = await c.guard.applyOnLatest('x', c.stale, (d) => { d.opportunities[0].aiWinRate = 9; });
      ok('O1 開關關閉：直接改舊物件並 save(舊物件)，沒有取鎖／ready，回傳 true', r === true && c.stale.opportunities[0].aiWinRate === 9 && c.calls.join() === 'save(stale)', c.calls.join());
    }
    // O2 開關開啟：短鎖內 ready() 重讀、只改最新物件、存最新物件
    {
      const c = mkStub(true);
      let heldInside = null;
      const r = await c.guard.applyOnLatest('ai x', c.stale, (d) => { heldInside = c.guard.isHeld(); d.opportunities[0].aiWinRate = 9; });
      ok('O2 開關開啟：鎖內 ready → 改最新物件 → save(最新) → flushMain；舊物件完全沒被改；回傳 true；結束後放鎖', r === true && heldInside === true && !c.guard.isHeld() && c.latest.opportunities[0].aiWinRate === 9 && c.stale.opportunities[0].aiWinRate === 0 && c.calls.join('|') === 'ready|save(latest)|flushMain', c.calls.join('|'));
    }
    // O3 mutate 回傳 false＝沒有東西要存
    {
      const c = mkStub(true);
      const r = await c.guard.applyOnLatest('ai x', c.stale, () => false);
      ok('O3 mutate 回傳 false → 不存檔（回傳 true＝沒有東西要存也不算失敗）', r === true && !c.calls.includes('save(latest)'), c.calls.join('|'));
    }
    // O4 取不到鎖：只記 log、不丟例外（外部呼叫已成功，不能讓整個請求失敗），回傳 false
    {
      const c = mkStub(true);
      const blocker = await c.guard.acquire('blocker', { waitMs: 100 });
      let thrown = null, r = null;
      try { r = await c.guard.applyOnLatest('ai x', c.stale, (d) => { d.opportunities[0].aiWinRate = 9; }, { waitMs: 100 }); } catch (e) { thrown = e; }
      await blocker.release();
      ok('O4 取不到鎖：不丟例外、記警告、不寫入、回傳 false', !thrown && r === false && c.warns.some((w) => w.includes('沒有存成功')) && c.latest.opportunities[0].aiWinRate === 1);
    }
    // O5 這三條路由確實改用 applyOnLatest（靜態）；AI 點數扣款包了短鎖
    ok('O5 opp-win-rate／contact-summary／SAP push/batch 都改用 applyOnLatest（不再拿 await 前的舊物件整份存回）',
      src.includes("applyOnLatest('ai opp-win-rate cache'") && src.includes("applyOnLatest('ai contact-summary cache'") && src.includes('sapWriteback.commit(data, logEntry)') && fsx.readFileSync(path.join(root, 'lib', 'sapLinkWriteback.js'), 'utf8').includes("applyOnLatest('sap push/batch result'"));
    // O6 AI 點數扣款：開關開啟時只把「扣點」放進短鎖，放鎖後才 next
    {
      const guardSrc = cut('function aiCreditsGuard(featureKey) {');
      const mk = (serialize, inner, busy, sub) => {
        const calls = [];
        const res = { headersSent: false, set() {}, status(c) { this.code = c; return this; }, json(j) { this.body = j; this.headersSent = true; return this; } };
        const ctx = vm.createContext({
          db: { ready: async () => { calls.push('ready'); }, flushMain: async () => { calls.push('flushMain'); } },
          writeGuard: { config: { serialize }, withLock: async (label, fn, opts) => { calls.push('lock'); if (busy) throw new Error('busy'); const r = await fn({}); calls.push('unlock'); return r; } },
          _aiCreditsBase: () => inner, console: { warn() {}, error() {} },
          loadAuth: () => ({ _subscription: sub === undefined ? { plan: 'pro', aiCredits: 100 } : sub }),
          createAiCreditsMiddleware: { CREDIT_COSTS: { 'opp-win-rate': 8 } },
        });
        vm.runInContext(guardSrc, ctx);
        return { calls, res, mw: vm.runInContext("aiCreditsGuard('opp-win-rate')", ctx) };
      };
      let nexted = false;
      const a = mk(true, (req, res, next) => { a.calls.push('inner'); next(); });
      await new Promise((r) => { a.mw({}, a.res, () => { nexted = true; a.calls.push('next'); r(); }); });
      ok('O6 開關開啟：鎖內 ready → 扣點 → flush → 放鎖，之後才呼叫 next（Gemini 在鎖外）', nexted && a.calls.join() === 'lock,ready,inner,flushMain,unlock,next', a.calls.join());
      const b = mk(true, (req, res) => { res.status(402).json({ code: 'INSUFFICIENT_CREDITS' }); });
      let nextedB = false;
      await new Promise((r) => { b.mw({}, b.res, () => { nextedB = true; }); setTimeout(r, 30); });
      ok('O7 點數不足（inner 自己回 402）：不呼叫 next', !nextedB && b.res.code === 402);
      const c2 = mk(true, () => {}, true);
      await new Promise((r) => { c2.mw({}, c2.res, () => {}); setTimeout(r, 30); });
      ok('O8 取不到鎖：回 503 DB_BUSY（可重試），AI 請求尚未執行', c2.res.code === 503 && c2.res.body && c2.res.body.retryable === true);
      let direct = false;
      const d0 = mk(false, (req, res, next) => { d0.calls.push('inner'); next(); });
      await new Promise((r) => { d0.mw({}, d0.res, () => { direct = true; r(); }); });
      ok('O9 開關關閉：完全沿用原本的 middleware（不取鎖）', direct && d0.calls.join() === 'inner');
      // O12~O14（審查 R2-#7）：不需要扣點的情況不搶鎖（搶鎖會讓這類請求在別人占鎖時無謂地吃 503，且多一次 ready＋flush 往返）
      for (const [nm, sub, expect] of [['沒有訂閱設定（舊安裝）', null, true], ['企業版無上限', { plan: 'enterprise', aiCredits: null }, true], ['點數不足（inner 自己回 402）', { plan: 'pro', aiCredits: 3 }, false]]) {
        let passed = false;
        const x = mk(true, (req, res, next) => { x.calls.push('inner'); if (expect) next(); else res.status(402).json({ code: 'INSUFFICIENT_CREDITS' }); }, false, sub);
        await new Promise((r) => { x.mw({}, x.res, () => { passed = true; r(); }); setTimeout(r, 30); });
        ok('O12 ' + nm + '：不取鎖，直接交給原本的 middleware（' + (expect ? '放行' : '回 402') + '）', x.calls.join() === 'inner' && passed === expect && (expect || x.res.code === 402), x.calls.join());
      }
      let passedPay = false;
      const pay = mk(true, (req, res, next) => { pay.calls.push('inner'); next(); }, false, { plan: 'pro', aiCredits: 100 });
      await new Promise((r) => { pay.mw({}, pay.res, () => { passedPay = true; r(); }); });
      ok('O13 真的要扣點（點數足夠）才取短鎖', passedPay && pay.calls.join() === 'lock,ready,inner,flushMain,unlock', pay.calls.join());
      // 預先判斷用的是快取；鎖內 inner 會用 ready() 之後的最新資料重新判斷（inner 本身就是每次讀 loadAuth）
    }
    // O10 Web Push 清除失效訂閱：網路等待之後在鎖內重讀最新 auth，只移除該 endpoint，不拿等待前的舊 auth 整份存回
    {
      const crypto = require('crypto');
      process.env.VAPID_PUBLIC_KEY = 'BNb1vCfdq3mE8Zx2QzA1p6m8X7z8wYk4m0Kk5H3m1wq5mP5v3dN2kq7yq0V8mZ4m3k1c2Zk6y7t8u9w0xAbCdEfGhI';
      process.env.VAPID_PRIVATE_KEY = crypto.randomBytes(32).toString('base64').replace(/[+]/g, '-').replace(/[/]/g, '_').replace(/=+$/, '');
      process.env.VAPID_SUBJECT = 'mailto:test@example.com';
      let PN = null, wp = null;
      try { wp = require('web-push'); } catch (_) { /* 沒有 web-push 套件：略過 */ }
      if (wp) {
        const origSend = wp.sendNotification, origSet = wp.setVapidDetails;
        wp.setVapidDetails = () => {};
        wp.sendNotification = async (sub) => { if (sub.endpoint === 'gone') throw Object.assign(new Error('gone'), { statusCode: 410 }); return {}; };
        delete require.cache[require.resolve('../lib/pushNotify')];
        PN = require('../lib/pushNotify');
        if (PN.isEnabled()) {
          const startAuth = { users: [{ username: 'u1', pushSubscriptions: [{ endpoint: 'gone', keys: {} }, { endpoint: 'live', keys: {} }] }] };
          const latestAuth = { users: [{ username: 'u1', pushSubscriptions: [{ endpoint: 'gone', keys: {} }, { endpoint: 'live', keys: {} }, { endpoint: 'new-device', keys: {} }], passwordHash: 'newer' }] };
          let saved = null, applied = false;
          const r = await PN.sendToUser(() => startAuth, (a) => { saved = a; }, 'u1', { title: 't' }, { applyOnLatestAuth: async (mutate) => { applied = true; mutate(latestAuth); } });
          ok('O10 推播 410：改在「最新 auth」上只移除失效的 endpoint；等待期間新增的裝置與其他欄位保留；沒有用舊 auth 呼叫 saveAuth',
            applied && saved === null && r.removed === 1 && latestAuth.users[0].pushSubscriptions.map((s) => s.endpoint).join() === 'live,new-device' && latestAuth.users[0].passwordHash === 'newer', JSON.stringify(latestAuth.users[0].pushSubscriptions));
          const startAuth2 = { users: [{ username: 'u1', pushSubscriptions: [{ endpoint: 'gone', keys: {} }, { endpoint: 'live', keys: {} }] }] };
          let saved2 = null;
          await PN.sendToUser(() => startAuth2, (a) => { saved2 = a; }, 'u1', { title: 't' });
          ok('O11 不給 opts（開關關閉）：沿用舊行為——改手上的 auth 並 saveAuth', saved2 === startAuth2 && startAuth2.users[0].pushSubscriptions.length === 1);
        } else {
          ok('O10 推播 410 清理（略過：測試環境無法啟用 Web Push）', true, '略過');
        }
        wp.sendNotification = origSend; wp.setVapidDetails = origSet;
        delete process.env.VAPID_PUBLIC_KEY; delete process.env.VAPID_PRIVATE_KEY; delete process.env.VAPID_SUBJECT;
        delete require.cache[require.resolve('../lib/pushNotify')];
      } else {
        ok('O10 推播 410 清理（略過：沒有 web-push 套件）', true, '略過');
      }
    }
  }

  // ═══ P. 第二輪審查（R2）修正的行為測試 ═══
  {
    const SER = { DB_WRITE_SERIALIZE: '1' };
    const MAIN_UPDATE_RE = /^INSERT INTO app_data \(id, content, updated_at\) VALUES \('main'/;
    const writer = (key) => async (req, res, db) => { const d = db.load(); d[key] = 1; db.save(d); res.status(200).json({ ok: key }); };
    const { createSapLinkWriteback, WRITEBACK_OPTS } = require('../lib/sapLinkWriteback');
    const iso = () => new Date().toISOString();
    const mkLog = (id, extra) => Object.assign({ id, integrationId: 'i1', startedAt: iso(), endedAt: null, operator: 'adm', pushed: 0, failed: 0, records: [] }, extra || {});
    const linkOf = (id, sap) => ({ crmType: 'company', crmId: id, integrationId: 'i1', sapId: sap, syncedAt: new Date(Date.now() + 1000).toISOString() });
    const countLinks = (main, id) => ((main && main.integrationLinks) || []).filter((l) => l.crmId === id).length;

    // ── P1 (#1) SAP 結果寫回：鎖被占 4 秒也要等到並寫入（舊版只等 3 秒、單次嘗試 → 默默丟掉對照） ──
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1 });
      const A = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      const B = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      const Y = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      await A.db.ready();
      const data = A.db.load();                               // 推送期間使用的快取
      await request(Y, 'POST', '/api/contacts', writer('fromY'));   // 別的實例寫入 → A 的快取過期，寫回時 d !== data（走合併路徑）
      const logEntry = mkLog('LG1', { endedAt: iso(), pushed: 2, records: [{ id: 'r1' }, { id: 'r2' }] });
      data.integrationLinks = [linkOf('c1', 'S1')]; data.integrationLogs = [logEntry];
      const wb = createSapLinkWriteback({ applyOnLatest: (l, d, m, o) => A.guard.applyOnLatest(l, d, m, o), log: quiet });
      const blocker = await B.guard.acquire('blocker', { waitMs: 500 });
      setTimeout(() => { blocker.release(); }, 4000);
      const t0 = Date.now();
      const okSaved = await wb.commit(data, logEntry);
      const main = fake.main();
      ok('P1 SAP 寫回：鎖被別的實例占住 4 秒 → 等到鎖並寫入成功（對照、紀錄、他人的寫入都在）', okSaved === true && countLinks(main, 'c1') === 1 && main.integrationLogs && main.integrationLogs[0].id === 'LG1' && main.fromY === 1 && Date.now() - t0 >= 3500, JSON.stringify(main).slice(0, 200) + ' ' + (Date.now() - t0) + 'ms');
      ok('P1b 寫回完成後沒有殘留的鎖／待補寫', !A.guard.isHeld() && wb.pendingCount() === 0 && fake.locks.size === 0);
      // 對照組：舊設定（等 3 秒、單次）在同樣情境一定失敗 → 證明這個測試抓得到原本的缺陷
      const A2 = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      const blocker2 = await B.guard.acquire('blocker2', { waitMs: 500 });
      setTimeout(() => { blocker2.release(); }, 4000);
      const d2 = (await A2.db.ready(), A2.db.load()); d2.integrationLinks = [linkOf('c9', 'S9')];
      const lgOld = mkLog('LGOLD', { endedAt: iso() }); d2.integrationLogs = [lgOld];
      const wbOld = createSapLinkWriteback({ applyOnLatest: (l, d, m) => A2.guard.applyOnLatest(l, d, m, { waitMs: 3000, holdMs: 8000 }), log: quiet });
      const oldSaved = await wbOld.commit(d2, lgOld);
      ok('P1c 【對照】舊設定（等 3 秒、單次）在同樣情境寫不進去（commit=false、DB 沒有對照）；現在會回報失敗而不是默默吞掉', oldSaved === false && countLinks(fake.main(), 'c9') === 0 && wbOld.pendingCount() === 1);
      await A2.db.flush();
      await sleep(1200);   // 等 blocker2 的計時器放鎖，免得影響後面的測試
    }

    // ── P2 (#1) 鎖被占超過所有重試 → commit=false、待補寫留在記憶體；下一次補寫並併入 ──
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1 });
      const A = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      const B = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      await A.db.ready();
      const wb = createSapLinkWriteback({ applyOnLatest: (l, d, m, o) => A.guard.applyOnLatest(l, d, m, o), opts: { waitMs: 150, attempts: 2, retryDelayMs: 10 }, log: quiet });
      const blocker = await B.guard.acquire('blocker', { waitMs: 500 });
      const data1 = A.db.load(); data1.integrationLinks = [linkOf('c1', 'S1')];
      const lg1 = mkLog('LG1', { endedAt: iso() }); data1.integrationLogs = [lg1];
      const r1 = await wb.commit(data1, lg1);
      ok('P2 鎖一直被占：commit 回 false（呼叫端據此回 saved:false＋warning），待補寫留在記憶體，DB 沒有對照', r1 === false && wb.pendingCount() === 1 && countLinks(fake.main(), 'c1') === 0);
      await blocker.release();
      // 下一次請求：先把待補寫併進這次使用的資料（避免重複建立），結尾一併補寫
      await A.db.ready();
      const data2 = A.db.load();
      data2.integrationLinks = [];   // 模擬快取被換掉、對照不在了
      const merged = wb.mergePendingInto(data2);
      ok('P2b 下一次推送開始時，上次沒寫成功的對照併回使用的資料（這次找得到 sapId，不會重複建立）', merged === 1 && data2.integrationLinks.some((l) => l.crmId === 'c1' && l.sapId === 'S1'));
      data2.integrationLinks.push(linkOf('c2', 'S2'));
      const lg2 = mkLog('LG2', { endedAt: iso() }); data2.integrationLogs = [lg2];
      const r2 = await wb.commit(data2, lg2);
      const main = fake.main();
      ok('P2c 下一次成功：上次待補寫的對照與紀錄一併寫入（c1、c2 各一筆、LG1／LG2 都在），待補寫清空', r2 === true && countLinks(main, 'c1') === 1 && countLinks(main, 'c2') === 1 && main.integrationLogs.map((x) => x.id).sort().join() === 'LG1,LG2' && wb.pendingCount() === 0, JSON.stringify(main.integrationLinks));
    }

    // ── P3 (#1) 存檔失敗：flushMain 會吞掉錯誤，必須看 save 自己的 promise ──
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1 });
      const A = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      await A.db.ready();
      const wb = createSapLinkWriteback({ applyOnLatest: (l, d, m, o) => A.guard.applyOnLatest(l, d, m, o), opts: { waitMs: 500, attempts: 3, retryDelayMs: 10 }, log: quiet });
      // 每次嘗試都存檔失敗（3 次）→ false
      fake.failNext(MAIN_UPDATE_RE, 3);
      const data = A.db.load(); data.integrationLinks = [linkOf('c1', 'S1')];
      const lg = mkLog('LG1', { endedAt: iso() }); data.integrationLogs = [lg];
      const r = await wb.commit(data, lg);
      ok('P3 存檔連續失敗 3 次 → commit 回 false（不是默默回成功）、DB 沒有對照、待補寫保留', r === false && countLinks(fake.main(), 'c1') === 0 && wb.pendingCount() === 1);
      // 失敗一次、重試成功 → true，且對照只有一筆（重試是冪等的）
      fake.failNext(MAIN_UPDATE_RE, 1);
      await A.db.ready();
      const data2 = A.db.load(); data2.integrationLinks = [linkOf('c2', 'S2')];
      const lg2 = mkLog('LG2', { endedAt: iso() }); data2.integrationLogs = [lg2];
      const r2 = await wb.commit(data2, lg2);
      const main = fake.main();
      ok('P3b 存檔失敗一次後重試成功 → true；先前待補寫的 c1 也一併補上；每筆對照只有一份', r2 === true && countLinks(main, 'c1') === 1 && countLinks(main, 'c2') === 1 && wb.pendingCount() === 0, JSON.stringify(main.integrationLinks));
    }

    // ── P4 (#5) 最新資料已有同 id 的半成品執行紀錄 → 整筆取代，不是略過 ──
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1, integrationLogs: [mkLog('LG1', { pushed: 1, records: [{ id: 'partial' }] }), mkLog('OLD', { endedAt: iso() })], integrationLinks: [] });
      const A = mkInstance(fake, SER, { waitMs: 8000, holdMs: 20000 });
      await A.db.ready();
      const wb = createSapLinkWriteback({ applyOnLatest: (l, d, m, o) => A.guard.applyOnLatest(l, d, m, o), log: quiet });
      const full = mkLog('LG1', { endedAt: iso(), pushed: 5, records: [{ id: 'a' }, { id: 'b' }, { id: 'c' }] });
      const data = { integrationLinks: [], integrationLogs: [full] };   // 不是快取物件 → 合併路徑
      const r = await wb.commit(data, full);
      const lgs = fake.main().integrationLogs;
      const one = lgs.filter((x) => x.id === 'LG1');
      ok('P4 DB 內已有同 id 的半成品紀錄 → 被完整紀錄取代（endedAt 有值、records 3 筆），不重複、其他紀錄保留', r === true && one.length === 1 && one[0].endedAt && one[0].records.length === 3 && lgs.some((x) => x.id === 'OLD'), JSON.stringify(lgs).slice(0, 200));
    }

    // ── P5 (#3) 存檔失敗後快取失效要在放鎖之前（確定性：扣住放鎖 COMMIT 的回應） ──
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1 });
      const X = mkInstance(fake, SER, { waitMs: 5000, holdMs: 20000 });
      await request(X, 'POST', '/api/contacts', writer('warm'));        // 先把建表等一次性的 COMMIT 用掉
      await X.db.flush();
      let r1Handled = null; const r1HandlerDone = new Promise((r) => { r1Handled = r; });
      let seenByR2 = null;
      fake.failNext(MAIN_UPDATE_RE, 1);                                  // R1 的存檔會失敗
      const hold = fake.holdNext(/^COMMIT/);                             // R1 放鎖的 COMMIT：資料庫端照常提交（鎖已放），回應被扣住
      const p1 = request(X, 'POST', '/api/contacts', async (req, res, db) => {
        const d = db.load(); d.failedItem = 1; db.save(d);
        r1Handled();
        res.status(200).json({ ok: 1 });
      });
      await r1HandlerDone;
      const p2 = request(X, 'POST', '/api/contacts', async (req, res, db) => {
        seenByR2 = JSON.parse(JSON.stringify(db.load()));
        const d = db.load(); d.r2 = 1; db.save(d);
        res.status(200).json({ ok: 2 });
      });
      await hold.reached;
      await sleep(60);                                                   // 讓 R2 在 R1 的 COMMIT 回應還被扣住時跑完它的 ready()／handler
      hold.release();
      const [a1, a2] = await Promise.all([p1, p2]);
      await X.db.flush();
      const fin = fake.main();
      ok('P5 存檔失敗的請求回 503', a1.res.statusCode === 503 && JSON.parse(a1.res.body).code === 'DB_WRITE_FAILED', a1.res.statusCode);
      ok('P5b 【確定性】R1 放鎖的 COMMIT 回應還沒到時，R2 看到的是資料庫內容（沒有 R1 失敗的變更）；最終 DB 也沒有 failedItem', seenByR2 && seenByR2.failedItem === undefined && fin.failedItem === undefined && fin.r2 === 1 && a2.res.statusCode === 200, JSON.stringify(seenByR2) + ' / ' + JSON.stringify(fin));
    }

    // ── P6 (#4) 同一請求多次 save：只看最後一次 ──
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ base: 1 });
      const X = mkInstance(fake, SER, { waitMs: 5000, holdMs: 20000 });
      await X.db.ready();
      // 前一次暫時失敗、最後一次成功 → 200，變更確實落地
      fake.failNext(MAIN_UPDATE_RE, 1);
      const a = await request(X, 'POST', '/api/contacts', async (req, res, db) => {
        const d = db.load(); d.item = 1; db.save(d);
        await sleep(30);
        const d2 = db.load(); d2.item2 = 1; db.save(d2);
        res.status(200).json({ ok: 1 });
      });
      ok('P6 前一次 save 暫時失敗、最後一次成功 → 回 200（變更已整份落地，不誤報失敗）', a.res.statusCode === 200 && fake.main().item === 1 && fake.main().item2 === 1, a.res.statusCode + ' ' + JSON.stringify(fake.main()));
      // 最後一次失敗 → 503
      await X.db.ready();
      fake.failNext(MAIN_UPDATE_RE, 1);
      const b = await request(X, 'POST', '/api/contacts', async (req, res, db) => {
        const d = db.load(); d.itemB = 1; db.save(d);
        await sleep(30);
        fake.failNext(MAIN_UPDATE_RE, 1);
        const d2 = db.load(); d2.itemB2 = 1; db.save(d2);
        res.status(200).json({ ok: 1 });
      });
      ok('P6b 最後一次 save 失敗 → 仍回 503 DB_WRITE_FAILED', b.res.statusCode === 503 && JSON.parse(b.res.body).code === 'DB_WRITE_FAILED', b.res.statusCode);
      // 前一次成功、最後一次失敗 → 503
      await X.db.ready();
      const c = await request(X, 'POST', '/api/contacts', async (req, res, db) => {
        const d = db.load(); d.itemC = 1; db.save(d);
        await sleep(30);
        fake.failNext(MAIN_UPDATE_RE, 1);
        const d2 = db.load(); d2.itemC2 = 1; db.save(d2);
        res.status(200).json({ ok: 1 });
      });
      ok('P6c 前一次成功、最後一次失敗 → 503', c.res.statusCode === 503, c.res.statusCode);
      await X.db.flush();
    }

    // ── P7 (#6) multipart 只對「註冊過的 multer 路由」延後取鎖；其他路由偽造標頭仍走全域鎖 ──
    {
      const fake = new FakeDb({ latency: [0, 1] });
      fake.seedMain({ n: 0 });
      const A = mkInstance(fake, SER, { waitMs: 300, holdMs: 20000 });
      const MP = { 'content-type': 'multipart/form-data; boundary=xx' };
      let heldInHandler = null;
      const forged = await request(A, 'DELETE', '/api/contacts/1', async (req, res, db) => { heldInHandler = A.guard.isHeld(); const d = db.load(); d.deleted1 = 1; db.save(d); res.status(200).json({ ok: 1 }); }, { headers: MP });
      ok('P7 偽造 multipart 標頭的一般寫入路由（未註冊）→ 照樣在全域中介層取鎖（handler 執行時持有鎖）', forged.res.statusCode === 200 && heldInHandler === true && fake.main().deleted1 === 1, 'held=' + heldInHandler);
      const B = mkInstance(fake, SER, { waitMs: 300, holdMs: 20000 });
      const blocker = await B.guard.acquire('blocker', { waitMs: 300 });
      const forged2 = await request(A, 'DELETE', '/api/contacts/2', writer('deleted2'), { headers: MP });
      ok('P7b 鎖被占住時，偽造標頭的未註冊路由 → 503 DB_BUSY（不能藉此繞過鎖）', forged2.res.statusCode === 503 && fake.main().deleted2 === undefined, forged2.res.statusCode);
      const reg = await request(A, 'POST', '/api/admin/import-contacts', writer('imp'), { headers: MP, bodyMs: 50 });
      ok('P7c 已註冊的 multer 路由：全域中介層放行、讀完 body 後才取鎖（鎖被占住 → 503）', reg.res.statusCode === 503 && A.guard.isLateRoute('/api/admin/import-contacts') && !A.guard.isLateRoute('/api/contacts/1'));
      await blocker.release();
      ok('P7d 結束後沒有殘留的鎖', !A.guard.isHeld() && !B.guard.isHeld() && fake.locks.size === 0);
    }

    // ── P8 (#2) preDeny：依實際路由的角色白名單；掃描所有 /api/admin/* 路由，非 admin 專用的必須登記 ──
    {
      const vm = require('vm');
      const root = path.join(__dirname, '..');
      const src = fsx.readFileSync(path.join(root, 'server.js'), 'utf8');
      const a = src.indexOf('const ADMIN_PATH_OPEN_ROLES = [');
      const fnA = src.indexOf('function writeGuardPreDeny(');
      const fnB = src.indexOf('\n}\n', fnA);
      if (a < 0 || fnA < 0 || fnB < 0) throw new Error('server.js 找不到 writeGuardPreDeny');
      const users = [{ username: 'adm', role: 'admin', active: true }, { username: 'mgr', role: 'manager1', active: true }, { username: 'sal', role: 'user', active: true }, { username: 'off', role: 'admin', active: false }];
      const ctx = vm.createContext({ loadAuth: () => ({ users }) });
      vm.runInContext(src.slice(a, fnB + 3), ctx);
      const pd = (full, username, role) => vm.runInContext('writeGuardPreDeny', ctx)({ session: { user: { username, role } } }, full);
      ok('P8 客戶池 /api/admin/pool/assign、/pool/contacts/:id/delete：manager1／executive／admin 不預先排除（要拿鎖），一般業務才排除', pd('/api/admin/pool/assign', 'mgr', 'manager1') === false && pd('/api/admin/pool/contacts/x/delete', 'mgr', 'manager2') === false && pd('/api/admin/pool/assign', 'e', 'executive') === false && pd('/api/admin/pool/assign', 'adm', 'admin') === false && pd('/api/admin/pool/assign', 'sal', 'user') === true);
      ok('P8b 其他 /api/admin/*：非 admin、停用帳號預先排除；admin 不排除；非 /api/admin 路徑不處理', pd('/api/admin/users', 'mgr', 'manager1') === true && pd('/api/admin/users', 'off', 'admin') === true && pd('/api/admin/users', 'adm', 'admin') === false && pd('/api/contacts', 'sal', 'user') === false && pd('/API/ADMIN/Users', 'sal', 'user') === true);
      // 掃描：每條 /api/admin/* 路由要嘛是 requireAdmin／qAdmin，要嘛在 ADMIN_PATH_OPEN_ROLES 登記
      const open = vm.runInContext('ADMIN_PATH_OPEN_ROLES', ctx).map((o) => o.prefix);
      const files = [path.join(root, 'server.js'), path.join(root, 'lib', 'quoteRoutes.js'), path.join(root, 'lib', 'mail', 'routes.js')];
      let scanned = 0; const unregistered = [];
      for (const f of files) {
        const L = fsx.readFileSync(f, 'utf8').split('\n');
        for (let i = 0; i < L.length; i++) {
          const m = /\b(?:app|router)\.(get|post|put|patch|delete)\(\s*['"`](\/api\/admin\/[^'"`]*)['"`]\s*,/.exec(L[i]);
          if (!m) continue;
          scanned++;
          const head = L.slice(i, i + 3).join(' ');
          if (/\brequireAdmin\b|\bqAdmin\b/.test(head)) continue;
          if (open.some((pre) => m[2].toLowerCase().startsWith(pre))) continue;
          unregistered.push(path.basename(f) + ':' + (i + 1) + ' ' + m[1].toUpperCase() + ' ' + m[2]);
        }
      }
      ok('P8c 所有 /api/admin/* 路由（共掃描 ' + scanned + ' 條）：不是 requireAdmin／qAdmin 的都已登記在 ADMIN_PATH_OPEN_ROLES（否則會被 preDeny 當成一定被拒絕而逃出寫入鎖）', scanned > 60 && unregistered.length === 0, unregistered.join(' ; '));
      // 登記的角色必須與 handler 內的角色檢查一致
      const poolRoles = vm.runInContext('ADMIN_PATH_OPEN_ROLES', ctx).find((o) => o.prefix === '/api/admin/pool/').roles.slice().sort().join();
      const handlerRoles = [...src.matchAll(/app\.(?:post|get)\('\/api\/admin\/pool\/[^']*', requireAuth, \(req, res\) => \{\s*const \{ username, role \} = req\.session\.user;\s*if \(!\[([^\]]+)\]\.includes\(role\)\)/g)].map((x) => x[1].replace(/['\s]/g, '').split(',').sort().join());
      ok('P8d 客戶池路由 handler 內的角色名單（' + handlerRoles.length + ' 條）與 ADMIN_PATH_OPEN_ROLES 一致', handlerRoles.length >= 2 && handlerRoles.every((r) => r === poolRoles), poolRoles + ' vs ' + handlerRoles.join(' | '));
    }

    // ── P9 (#10) DB_WRITE_LOCK_EXEMPT 的比對規則：與 DEPLOY.md／.env.example 的說明一致 ──
    {
      const cfg = WG.readConfig({ DB_WRITE_SERIALIZE: '1', DB_WRITE_LOCK_EXEMPT: '/api/foo,/api/bar/' });
      const need = (m, p) => WG.needsLock(cfg, m, p);
      ok('P9 不以 / 結尾（/api/foo）＝該路徑本身及其所有子路徑都豁免；字面相似的兄弟路徑（/api/foox）仍要鎖', !need('POST', '/api/foo') && !need('DELETE', '/api/foo/1/x') && need('POST', '/api/foox'));
      ok('P9b 以 / 結尾（/api/bar/）＝只豁免其底下的路徑；/api/bar 本身仍要鎖', !need('POST', '/api/bar/1') && need('POST', '/api/bar'));
      const doc = fsx.readFileSync(path.join(__dirname, '..', 'DEPLOY.md'), 'utf8');
      const envx = fsx.readFileSync(path.join(__dirname, '..', '.env.example'), 'utf8');
      ok('P9c 文件與程式一致：.env.example 與 DEPLOY.md 都寫明「不以 / 結尾＝含子路徑」，且不再有「否則為精確路徑」的錯誤說法', /不以 \/ 結尾＝該路徑及其所有子路徑/.test(envx) && /不以 `\/` 結尾[^\n]*及其所有子路徑/.test(doc) && !/否則為精確路徑/.test(envx + doc));
    }
  }

  // ── 輸出 ──
  let bad = 0;
  for (const [n, p, e] of results) { if (!p) bad++; console.log((p ? 'PASS ' : 'FAIL ') + n + (e && !p ? '  ← ' + e : (e && /^D|^C|^B9|^E10/.test(n) ? '  [' + e + ']' : ''))); }
  console.log(`\n${results.length - bad}/${results.length} 通過` + (bad ? `，${bad} 失敗` : ''));
  summaryPrinted = true;
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error('測試腳本例外：', e); process.exit(2); });
