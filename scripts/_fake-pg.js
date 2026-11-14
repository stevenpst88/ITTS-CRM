'use strict';
/**
 * 測試用「假 Postgres」：只模擬本專案寫入競爭防護與 db/postgres.js 用到的 SQL 子集，並忠實模擬以下語意：
 *   - 連線池上限（max）與排隊；每次 query 有隨機網路延遲（製造跨實例交錯）
 *   - 交易：BEGIN/COMMIT/ROLLBACK；SET LOCAL lock_timeout／idle_in_transaction_session_timeout
 *   - pg_advisory_xact_lock：同一把鎖互斥、FIFO 等待、lock_timeout 到期丟 55P03、交易結束或連線被終止即釋放
 *   - 單一敘述的原子性（INSERT … ON CONFLICT DO UPDATE … RETURNING、WITH … INSERT … RETURNING）
 *   - 故障注入：failNext(正規表示式, 次數)
 * 重要限制：這是「邏輯」驗證用的模擬，不等於真實 Postgres 的行為（交易隔離、MVCC、pooler 相容性都沒有模擬）。
 * 真實 PG 的驗證要靠 scripts/verify-write-race-pg.js（由使用者親自執行）。
 */

function norm(sql) { return String(sql).replace(/\s+/g, ' ').trim(); }
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

class FakeDb {
  /** opts: { latency: [minMs,maxMs] 每次 query 前後各抽一次延遲, seed } */
  constructor(opts) {
    opts = opts || {};
    this.latency = opts.latency || [0, 2];
    this.slow = opts.slow || null;       // { re, ms:[a,b] }：符合的敘述「執行完之後」回應額外延遲（模擬大 SELECT 的傳輸時間；資料庫端看到的是執行當下的資料）
    this.appData = new Map();            // id -> { content:string(JSON), updated_at:number(ms) }
    this.seqTable = new Map();           // ym -> last_serial
    this.issued = new Map();             // quote_no -> row
    this.locks = new Map();              // key -> { holder:tx, queue:[{tx,resolve,reject,timer}] }
    this.faults = [];                    // { re, n, error | hang }
    this.holds = [];                     // { re, used, reached, gate }：命中後先在「資料庫端」執行完（看到當時的資料），回應被扣住直到 release()
    this.log = [];                       // 執行過的敘述（除錯用）
    this.stats = { queries: 0, lockWaits: 0, lockTimeouts: 0, idleKills: 0, ddl: 0 };
    this.defaultTables = new Set();      // 預設 schema 內已建立的表（供 to_regclass 檢查）
    this.defaultRls = new Set();
    this.schemas = new Map();            // schema -> { tables:Set, rls:Set, blob:Map }
    this.clients = new Map();            // pid -> client
    this._pid = 1000;
    this._lastTs = 0;
    this._rnd = mulberry32(opts.seed || 12345);
  }
  now() { this._lastTs = Math.max(Date.now(), this._lastTs + 1); return this._lastTs; }   // 單調遞增（真實 PG 為微秒精度）
  rand() { return this._rnd(); }
  delay() { const [a, b] = this.latency; return b <= 0 ? 0 : a + Math.floor(this.rand() * (b - a + 1)); }
  failNext(re, n, error) { this.faults.push({ re, n: n || 1, error: error || Object.assign(new Error('injected failure'), { code: 'XX000' }) }); }
  /** 下一次符合 re 的敘述：永不回應（模擬黑洞連線／半開 TCP）。n＝次數 */
  hangNext(re, n) { this.faults.push({ re, n: n || 1, hang: true }); }
  /** 下一次符合 re 的敘述：資料庫端照常執行（交易照常提交），但回應被扣住到 release()。回傳 { reached:Promise, release() } */
  holdNext(re) {
    let reachedRes, gateRes;
    const h = { re, used: false, reached: new Promise((r) => { reachedRes = r; }), gate: new Promise((r) => { gateRes = r; }), _reached: () => reachedRes(), release: () => gateRes() };
    this.holds.push(h);
    return h;
  }
  seedMain(obj) { this.appData.set('main', { content: JSON.stringify(obj), updated_at: this.now() }); }
  main() { const r = this.appData.get('main'); return r ? JSON.parse(r.content) : null; }
  pool(opts) { return new FakePool(this, opts); }
}

function mulberry32(a) { return function () { a |= 0; a = (a + 0x6D2B79F5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }

class FakePool {
  constructor(db, opts) {
    this.db = db; this.max = (opts && opts.max) || 2; this.connectTimeout = (opts && opts.connectionTimeoutMillis) || 5000;
    this.idle = []; this.total = 0; this.waiters = []; this._handlers = {};
  }
  on(ev, fn) { this._handlers[ev] = fn; return this; }
  async connect() {
    if (this.idle.length) return this.idle.pop();
    if (this.total < this.max) { this.total++; return new FakeClient(this); }
    return new Promise((resolve, reject) => {
      const w = { resolve, timer: setTimeout(() => { const i = this.waiters.indexOf(w); if (i >= 0) this.waiters.splice(i, 1); reject(new Error('timeout exceeded when trying to connect')); }, this.connectTimeout) };
      this.waiters.push(w);
    });
  }
  _return(client, destroy) {
    if (destroy || client.broken) { this.total--; if (this.waiters.length && this.total < this.max) { this.total++; const w = this.waiters.shift(); clearTimeout(w.timer); w.resolve(new FakeClient(this)); } return; }
    if (this.waiters.length) { const w = this.waiters.shift(); clearTimeout(w.timer); w.resolve(client); } else this.idle.push(client);
  }
  async query(sql, params) { const c = await this.connect(); try { return await c.query(sql, params); } finally { c.release(); } }
  async end() { /* noop */ }
}

class FakeClient {
  constructor(pool) { this.pool = pool; this.db = pool.db; this.pid = ++pool.db._pid; pool.db.clients.set(this.pid, this); this.tx = null; this.broken = false; this.released = false; this._errH = null; this._idleTimer = null; }
  on(ev, fn) { if (ev === 'error') this._errH = fn; return this; }
  removeListener() { this._errH = null; return this; }
  release(destroy) { if (this.tx) this._endTx(); this.pool._return(this, destroy === true); }
  _terminate() {   // 伺服器端終止連線（idle_in_transaction_session_timeout）
    this.db.stats.idleKills++;
    this.broken = true;
    this._endTx();
    if (this._errH) { try { this._errH(Object.assign(new Error('terminating connection due to idle-in-transaction timeout'), { code: '25P03' })); } catch (_) { /* ignore */ } }
  }
  _armIdle() {
    if (this._idleTimer) { clearTimeout(this._idleTimer); this._idleTimer = null; }
    if (this.tx && this.tx.idleMs > 0) { this._idleTimer = setTimeout(() => this._terminate(), this.tx.idleMs); if (this._idleTimer.unref) this._idleTimer.unref(); }
  }
  _endTx() {
    if (this._idleTimer) { clearTimeout(this._idleTimer); this._idleTimer = null; }
    const tx = this.tx; this.tx = null;
    if (!tx) return;
    for (const key of tx.locks) {
      const L = this.db.locks.get(key);
      if (!L || L.holder !== tx) continue;
      L.holder = null;
      const next = L.queue.shift();
      if (next) { clearTimeout(next.timer); L.holder = next.tx; next.tx.locks.add(key); next.resolve(); }
      else this.db.locks.delete(key);
    }
  }
  async query(sql, params) {
    params = params || [];
    const db = this.db;
    if (this.broken) throw Object.assign(new Error('Client has encountered a connection error and is not queryable'), { code: '57P01' });
    const s = norm(sql);
    if (this._idleTimer) { clearTimeout(this._idleTimer); this._idleTimer = null; }
    await sleep(db.delay());
    if (this.broken) throw Object.assign(new Error('terminating connection'), { code: '57P01' });
    db.stats.queries++; db.log.push(s);
    for (const f of db.faults) { if (f.n > 0 && f.re.test(s)) { f.n--; if (f.hang) await new Promise(() => {}); throw f.error; } }
    let holdGate = null;
    for (const h of db.holds) { if (!h.used && h.re.test(s)) { h.used = true; holdGate = h; break; } }
    let res;
    if (params.length === 0 && s.includes(';')) { for (const part of s.split(';').map((x) => x.trim()).filter(Boolean)) res = await this._exec(part, params); }
    else res = await this._exec(s, params);
    if (holdGate) { holdGate._reached(); await holdGate.gate; }
    if (db.slow && db.slow.re.test(s)) await sleep(db.slow.ms[0] + Math.floor(db.rand() * (db.slow.ms[1] - db.slow.ms[0] + 1)));
    await sleep(db.delay());
    if (this.tx && !this.broken) this._armIdle();
    return res;
  }
  async _exec(s, p) {
    const db = this.db;
    let m;
    if (s === 'BEGIN') { this.tx = { locks: new Set(), lockMs: 0, idleMs: 0, client: this }; return { rows: [] }; }
    if (s === 'BEGIN READ ONLY') { this.tx = { locks: new Set(), lockMs: 0, idleMs: 0, client: this }; return { rows: [] }; }
    if (s === 'COMMIT' || s === 'ROLLBACK') { this._endTx(); return { rows: [] }; }
    if ((m = /^SET LOCAL lock_timeout = (\d+)$/.exec(s))) { if (!this.tx) throw new Error('SET LOCAL outside tx'); this.tx.lockMs = +m[1]; return { rows: [] }; }
    if ((m = /^SET LOCAL statement_timeout = (\d+)$/.exec(s))) { if (!this.tx) throw new Error('SET LOCAL outside tx'); this.tx.stmtMs = +m[1]; return { rows: [] }; }
    if ((m = /^SET LOCAL idle_in_transaction_session_timeout = (\d+)$/.exec(s))) { if (!this.tx) throw new Error('SET LOCAL outside tx'); this.tx.idleMs = +m[1]; return { rows: [] }; }
    if ((m = /^SELECT pg_advisory_xact_lock\((-?\d+|\$1::bigint)\)$/.exec(s))) return this._advisory(m[1] === '$1::bigint' ? String(p[0]) : m[1]);
    if (s.startsWith('SELECT to_regclass($1::text) IS NOT NULL AS a, to_regclass($2::text) IS NOT NULL AS b')) {
      const has = (q) => { const mm = /^(?:"([a-z0-9_]+)"\.)?"([a-z0-9_]+)"$/.exec(String(q)); if (!mm) return false; if (mm[1]) { const sc = db.schemas.get(mm[1]); return !!(sc && sc.tables.has(mm[2])); } return db.defaultTables.has(mm[2]); };
      return { rows: [{ a: has(p[0]), b: has(p[1]) }] };
    }
    if ((m = /^CREATE TABLE IF NOT EXISTS (app_data|"quote_no_seq"|"quote_no_issued")/.exec(s))) { db.stats.ddl++; db.defaultTables.add(m[1].replace(/"/g, '')); return { rows: [] }; }
    if ((m = /^ALTER TABLE "quote_no_(seq|issued)" ENABLE ROW LEVEL SECURITY/.exec(s))) { db.stats.ddl++; db.defaultRls.add('quote_no_' + m[1]); return { rows: [] }; }   // 預設 schema（正式用）；帶 schema 的由下方處理

    // ── app_data ──
    if (s === "SELECT content, updated_at FROM app_data WHERE id = 'main'") {
      const r = db.appData.get('main'); return { rows: r ? [{ content: JSON.parse(r.content), updated_at: new Date(r.updated_at) }] : [] };
    }
    if (s === "SELECT updated_at FROM app_data WHERE id = 'main'") {
      const r = db.appData.get('main'); return { rows: r ? [{ updated_at: new Date(r.updated_at) }] : [] };
    }
    if (s.startsWith("INSERT INTO app_data (id, content) VALUES ('main', $1) RETURNING updated_at")) {
      const t = db.now(); db.appData.set('main', { content: p[0], updated_at: t }); return { rows: [{ updated_at: new Date(t) }] };
    }
    if (s.startsWith("INSERT INTO app_data (id, content, updated_at) VALUES ('main', $1, now()) ON CONFLICT")) {
      const t = db.now(); db.appData.set('main', { content: p[0], updated_at: t }); return { rows: [{ updated_at: new Date(t) }] };
    }
    if ((m = /^INSERT INTO app_data \(id, content, updated_at\) VALUES \('([a-z_-]+)', \$1, now\(\)\) ON CONFLICT/.exec(s))) {
      db.appData.set(m[1], { content: p[0], updated_at: db.now() }); return { rows: [] };
    }
    if ((m = /^SELECT content FROM app_data WHERE id = '([a-z_-]+)'$/.exec(s))) {
      const r = db.appData.get(m[1]); return { rows: r ? [{ content: JSON.parse(r.content) }] : [] };
    }

    // ── 原子取號（單一敘述，JS 同步段內完成＝原子）──
    if (s.startsWith('WITH c AS ( INSERT INTO "quote_no_seq"') || /^WITH c AS \( INSERT INTO (?:"[a-z0-9_]+"\.)?"quote_no_seq"/.test(s)) {
      const ym = String(p[0]), floor = Number(p[1]);
      const cur = db.seqTable.get(ym);
      const last = cur === undefined ? floor + 1 : Math.max(cur, floor) + 1;
      db.seqTable.set(ym, last);
      const quoteNo = 'QU-' + ym + '-' + (last < 1000 ? String(last).padStart(3, '0') : String(last));
      if (db.issued.has(quoteNo)) throw Object.assign(new Error('duplicate key value violates unique constraint'), { code: '23505' });
      db.issued.set(quoteNo, { quote_no: quoteNo, ym, serial: last, issued_to: p[2] == null ? null : String(p[2]), issued_at: new Date(db.now()) });
      return { rows: [{ serial: last, quote_no: quoteNo }] };
    }
    if (/^SELECT quote_no, ym, serial, issued_to, issued_at FROM (?:"[a-z0-9_]+"\.)?"quote_no_issued"/.test(s)) {
      let rows = [...db.issued.values()];
      if (/WHERE ym = \$1::text/.test(s)) rows = rows.filter((r) => r.ym === String(p[0]));
      rows.sort((a, b) => (a.ym < b.ym ? -1 : a.ym > b.ym ? 1 : a.serial - b.serial));
      return { rows };
    }

    // ── verify-write-race-pg.js 用的 SQL ──
    if (s.startsWith('SELECT version() AS v')) return { rows: [{ v: 'PostgreSQL 15.0 (FakePG)' }] };
    if (s === "SELECT to_regclass('public.app_data') IS NOT NULL AS ok") return { rows: [{ ok: db.appData.size > 0 }] };
    if ((m = /^CREATE SCHEMA "([a-z0-9_]+)"$/.exec(s))) { if (db.schemas.has(m[1])) throw Object.assign(new Error('schema exists'), { code: '42P06' }); db.schemas.set(m[1], { tables: new Set(), rls: new Set(), blob: new Map() }); return { rows: [] }; }
    if ((m = /^DROP SCHEMA IF EXISTS "([a-z0-9_]+)" CASCADE$/.exec(s))) { db.schemas.delete(m[1]); return { rows: [] }; }
    if (s.startsWith('SELECT to_regnamespace($1::text) IS NULL AS gone')) return { rows: [{ gone: !db.schemas.has(String(p[0])) }] };
    if ((m = /^CREATE TABLE IF NOT EXISTS "([a-z0-9_]+)"."([a-z0-9_]+)"/.exec(s))) { const sc = db.schemas.get(m[1]); if (!sc) throw Object.assign(new Error('schema does not exist'), { code: '3F000' }); sc.tables.add(m[2]); return { rows: [] }; }
    if ((m = /^ALTER TABLE "([a-z0-9_]+)"."([a-z0-9_]+)" ENABLE ROW LEVEL SECURITY$/.exec(s))) { db.schemas.get(m[1]).rls.add(m[2]); return { rows: [] }; }
    if ((m = /^CREATE TABLE "([a-z0-9_]+)"."wrg_blob"/.exec(s))) { const sc = db.schemas.get(m[1]); sc.tables.add('wrg_blob'); return { rows: [] }; }
    if (s.startsWith('SELECT table_name FROM information_schema.tables WHERE table_schema = $1')) { const sc = db.schemas.get(String(p[0])); return { rows: sc ? [...sc.tables].filter((t) => t !== 'wrg_blob').sort().map((t) => ({ table_name: t })) : [] }; }
    if (s.startsWith('SELECT c.relname, c.relrowsecurity FROM pg_class')) { const sc = db.schemas.get(String(p[0])); return { rows: sc ? [...sc.tables].filter((t) => t !== 'wrg_blob').sort().map((t) => ({ relname: t, relrowsecurity: sc.rls.has(t) })) : [] }; }
    if ((m = /^DELETE FROM "([a-z0-9_]+)"."wrg_blob"$/.exec(s))) { db.schemas.get(m[1]).blob.clear(); return { rows: [] }; }
    if ((m = /^INSERT INTO "([a-z0-9_]+)"\."wrg_blob" \(id, content\) VALUES \('main', '(.+)'::jsonb\)$/.exec(s))) { db.schemas.get(m[1]).blob.set('main', m[2]); return { rows: [] }; }
    if ((m = /^SELECT content FROM "([a-z0-9_]+)"."wrg_blob" WHERE id = 'main'$/.exec(s))) { const v = db.schemas.get(m[1]).blob.get('main'); return { rows: [{ content: JSON.parse(v) }] }; }
    if ((m = /^UPDATE "([a-z0-9_]+)"\."wrg_blob" SET content = \$1::jsonb WHERE id = 'main'$/.exec(s))) { db.schemas.get(m[1]).blob.set('main', p[0]); return { rows: [] }; }
    if (s.startsWith('SELECT count(*)::int AS n FROM pg_locks')) { const L = db.locks.get(String(p[0])); return { rows: [{ n: L && L.holder ? 1 : 0 }] }; }
    if (s.startsWith('SELECT pid FROM pg_locks')) { const L = db.locks.get(String(p[0])); return { rows: L && L.holder && L.holder.client ? [{ pid: L.holder.client.pid }] : [] }; }
    if (s.startsWith('SELECT pg_terminate_backend(')) { const c = db.clients.get(Number(p[0])); if (c) c._terminate(); return { rows: [{ pg_terminate_backend: !!c }] }; }
    throw new Error('[fake-pg] 不支援的 SQL：' + s.slice(0, 160));
  }
  _advisory(key) {
    const db = this.db;
    const tx = this.tx || { locks: new Set(), lockMs: 0, idleMs: 0, auto: true, client: this };
    let L = db.locks.get(key);
    if (!L) { L = { holder: null, queue: [] }; db.locks.set(key, L); }
    if (!L.holder || L.holder === tx) { L.holder = tx; tx.locks.add(key); return Promise.resolve({ rows: [{ pg_advisory_xact_lock: '' }] }); }
    db.stats.lockWaits++;
    return new Promise((resolve, reject) => {
      const w = { tx, resolve: () => resolve({ rows: [{ pg_advisory_xact_lock: '' }] }), reject, timer: null };
      if (tx.lockMs > 0) {
        w.timer = setTimeout(() => {
          const i = L.queue.indexOf(w); if (i >= 0) L.queue.splice(i, 1);
          db.stats.lockTimeouts++;
          reject(Object.assign(new Error('canceling statement due to lock timeout'), { code: '55P03' }));
        }, tx.lockMs);
      }
      L.queue.push(w);
    });
  }
}

/** 載入「全新的」db/postgres.js（獨立模組狀態＝一個獨立的 Vercel 實例），連到共用的 FakeDb */
function freshPostgres(fakeDb, env, poolOpts) {
  const path = require('path');
  const root = path.join(__dirname, '..');
  const keep = {};
  const keys = ['DATABASE_URL', 'DB_WRITE_SERIALIZE', 'QUOTE_NO_ATOMIC', 'AUDIT_EVENTS_TABLE'];
  keys.forEach((k) => { keep[k] = process.env[k]; delete process.env[k]; });
  process.env.DATABASE_URL = 'postgres://fake:fake@127.0.0.1:1/fake';
  Object.keys(env || {}).forEach((k) => { process.env[k] = env[k]; });
  for (const k of Object.keys(require.cache)) {
    if (k.startsWith(path.join(root, 'db') + path.sep) || k === path.join(root, 'lib', 'pgWriteGuard.js')) delete require.cache[k];
  }
  const mod = require(path.join(root, 'db', 'postgres.js'));
  mod._setPoolForTest(fakeDb.pool(poolOpts || { max: (env && env.DB_WRITE_SERIALIZE === '1') ? 3 : 2 }));
  keys.forEach((k) => { if (keep[k] === undefined) delete process.env[k]; else process.env[k] = keep[k]; });
  return mod;
}

module.exports = { FakeDb, FakePool, freshPostgres, sleep };
