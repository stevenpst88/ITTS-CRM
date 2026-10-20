#!/usr/bin/env node
'use strict';
/**
 * 內控稽核第三波第 3 項「稽核日誌防竄改（hash chain）＋外部錨點＋簽核歷程鏈」檢查。不含任何帳密：金鑰與測試帳號密碼都是隨機值、只在記憶體／子行程環境變數。
 *
 *  U  單元（不開伺服器）
 *     U1 canonical／getConfig（旗標關、無金鑰、金鑰太短、備援金鑰、SECRETS_JSON、金鑰不可序列化）
 *     U2 假 Postgres 的交易 SQL 順序與 advisory lock、多實例交錯（序號無缺號）、失敗處理（連線失敗／交易中途失敗／DDL 失敗不丟事件）
 *     U3 竄改偵測（Postgres 假庫＋本機 JSONL 兩種後端各跑一遍）：改欄位、刪中間列、刪首列、插入偽造列、換順序、截尾（有／無錨點）、回溯重寫、錯誤金鑰、kid 被換、舊資料涵蓋報告、換金鑰、區間驗證、分批
 *     U4 旗標關閉＝與第二波（397abc5）的寫入位元級相同（SQL 與參數、JSONL 位元組、stats 欄位）
 *     U5 錨點簽章、Webhook SSRF 防護（http／內網／metadata／轉址／逾時）、Cron 驗證
 *     U6 簽核歷程鏈（真實 lib/quoteRoutes.js：送簽→核准全流程、序列化不含 h、history-check 權限、竄改偵測、旗標關閉不變）
 *     U7 後台版本資訊頁「稽核防竄改」卡（vm 渲染、XSS 跳脫、各種狀態）
 *     U9 scripts/verify-audit-pg.js --chain 的 [13] 區段在假庫上自測（使用者在真庫執行前先確認腳本本身沒有手誤；腳本本身絕不由本檢查對真庫執行）
 *     U8 變異檢查（刻意弄壞：prev 銜接／HMAC 金鑰／advisory lock／序號連續／錨點比對／旗標／錨點簽章／canonical）——每個變異都必須被抓到
 *  E  端對端：自己的伺服器（JSONL 本機模式，絕不連任何資料庫），旗標開／關／金鑰太短 三輪
 *
 * 守門：工作樹若已有 auth.json／data.json／audit.log.json(l)／contact_audit.json（可能是真實資料），一律拒絕執行。
 * 做法：臨時 auth.json／data.json（假資料＋_s_* 臨時帳號），以 PORT=AC_PORT（預設 3201）啟動 server.js，跑完只用 PID 關閉並刪除臨時檔。
 * 禁用埠：3000／3020／3030／3101–3195。
 *
 *   node scripts/check-audit-chain.js              # 全部
 *   AC_ONLY=U,E node scripts/check-audit-chain.js  # 只跑指定區段（U／E）
 *   AC_SKIP_MUT=1 ...                              # 略過變異檢查
 */
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const vm = require('vm');
const { spawn, execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.AC_PORT || '3201', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3195)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const GUARD = ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f)).filter(f => fs.existsSync(f));
if (GUARD.length) { console.log('拒絕：此工作樹已有', GUARD.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json', 'mail-outbox.json'].map(f => path.join(ROOT, f));
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const AC = require(path.join(ROOT, 'lib', 'auditChain.js'));
const Store = require(path.join(ROOT, 'lib', 'auditStore.js'));
const HC = require(path.join(ROOT, 'lib', 'historyChain.js'));

let pass = 0, fail = 0, muting = false, mutFail = 0;
const failures = [];
const UNVERIFIED = [];   // 被略過、沒有真的驗證到的項目。有項目就 exit 3，不能被當成通過（B-10）
function ok(cond, name) {
  if (cond) { if (!muting) pass++; return; }
  if (muting) { mutFail++; return; }
  fail++; failures.push(name); console.log('  ✗ FAIL:', name);
}
const J = (x) => JSON.stringify(x);
function eq(a, b, name) { ok(J(a) === J(b), `${name}（得到 ${J(a)}，預期 ${J(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));
const rnd = (n) => crypto.randomBytes(n || 24).toString('hex');
const SECRET = rnd(32);                       // 測試用鏈金鑰（只在記憶體）
const CFG = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: SECRET });
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'ac-chain-'));
const cleanupTmp = () => { try { for (const f of fs.readdirSync(TMP)) fs.rmSync(path.join(TMP, f), { force: true, recursive: true }); fs.rmdirSync(TMP); } catch { /* 已清 */ } };

/** 暫時蓋掉 console.error，回傳收集到的訊息 */
async function captureErr(fn) {
  const orig = console.error; const msgs = [];
  console.error = (...a) => { msgs.push(a.map(String).join(' ')); };
  try { await fn(); } finally { console.error = orig; }
  return msgs;
}

// ═════════════ 假 Postgres（有交易、advisory lock、唯一索引）═════════════
function makeChainPg(shared, opt) {
  opt = opt || {};
  const S = shared || { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] };
  const log = [];
  const norm = (s) => s.replace(/\s+/g, ' ').trim();
  const tick = async () => { if (opt.jitter) await sleep(Math.floor(Math.random() * opt.jitter)); else await Promise.resolve(); };
  function acquire(c) { return new Promise((res) => { if (!S.lockHolder) { S.lockHolder = c; res(); } else S.waiters.push({ c, res }); }); }
  function release(c) {
    if (S.lockHolder === c) { S.lockHolder = null; const w = S.waiters.shift(); if (w) { S.lockHolder = w.c; w.res(); } }
    else S.waiters = S.waiters.filter((w) => w.c !== c);
  }
  const dup = () => Object.assign(new Error('duplicate key value violates unique constraint "audit_events_seq_uidx"'), { code: '23505' });
  const mk = (p, seq) => ({ id: ++S.id, ts: new Date(p[0]), kind: p[1], action: p[2], operator: p[3], target: p[4], detail: p[5], ip: p[6], ua: p[7], extra: JSON.parse(p[8]), seq, prev_hash: seq == null ? null : p[10], row_hash: seq == null ? null : p[11], kid: seq == null ? null : p[12] });
  const out = (r) => ({ id: String(r.id), ts: r.ts, kind: r.kind, action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip, ua: r.ua, extra: r.extra, seq: r.seq == null ? null : String(r.seq), prev_hash: r.prev_hash, row_hash: r.row_hash, kid: r.kid });
  async function exec(c, sql, params) {
    params = params || [];
    const s = norm(sql);
    log.push(s);
    await tick();
    if (/^BEGIN$/i.test(s)) { if (opt.failBegin) throw new Error('simulated begin failure'); c.inTx = true; c.pend = []; return { rows: [] }; }
    if (/^SET LOCAL lock_timeout = \d+$/i.test(s)) return { rows: [] };
    if (/^SELECT pg_advisory_xact_lock\(\$1::bigint\)$/i.test(s)) { await acquire(c); c.locked = true; return { rows: [] }; }
    if (/^COMMIT$/i.test(s)) {
      try { for (const r of c.pend) { if (S.rows.some((x) => x.seq != null && x.seq === r.seq)) throw dup(); } S.rows.push(...c.pend); }
      finally { c.inTx = false; c.pend = []; release(c); }
      return { rows: [] };
    }
    if (/^ROLLBACK$/i.test(s)) { c.inTx = false; c.pend = []; release(c); return { rows: [] }; }
    if (/^CREATE TABLE IF NOT EXISTS audit_events/i.test(s) || /^CREATE INDEX IF NOT EXISTS audit_events_/i.test(s)) { S.ddl.push(s); const mi = /INDEX IF NOT EXISTS (w+)/i.exec(s); if (mi) (S.idx = S.idx || new Set()).add(mi[1]); return { rows: [] }; }
    if (/^ALTER TABLE audit_events ADD COLUMN IF NOT EXISTS seq BIGINT/i.test(s)) { if (opt.failAlter) throw new Error('must be owner of table audit_events'); S.chainCols = true; S.ddl.push(s); return { rows: [] }; }
    if (/^CREATE UNIQUE INDEX IF NOT EXISTS audit_events_seq_uidx/i.test(s)) { S.ddl.push(s); (S.idx = S.idx || new Set()).add('audit_events_seq_uidx'); return { rows: [] }; }
    if (/^SELECT indexname FROM pg_indexes WHERE tablename = 'audit_events'$/i.test(s)) return { rows: Array.from(S.idx || []).map((n) => ({ indexname: n })) };
    if (/^SELECT 1 FROM audit_events LIMIT 1$/i.test(s)) return { rows: [] };
    if (/^SELECT count\(\*\)::int AS n FROM information_schema\.columns/i.test(s)) return { rows: [{ n: S.chainCols ? 4 : 0 }] };
    if (/^SELECT seq, row_hash FROM audit_events WHERE seq IS NOT NULL ORDER BY seq DESC LIMIT 1$/i.test(s) || /^SELECT seq, row_hash, kid FROM audit_events WHERE seq IS NOT NULL ORDER BY seq DESC LIMIT 1$/i.test(s)) {
      const r = S.rows.filter((x) => x.seq != null).sort((a, b) => b.seq - a.seq)[0];
      return { rows: r ? [{ seq: String(r.seq), row_hash: r.row_hash, kid: r.kid }] : [] };
    }
    if (/^INSERT INTO audit_events \(ts, kind, action, operator, target, detail, ip, ua, extra, seq,/i.test(s)) {
      if (params.length !== 13) throw new Error('chain INSERT 參數數量不對');
      if (opt.failChainInsert > 0) { opt.failChainInsert--; throw new Error('simulated insert failure'); }
      const seq = Number(params[9]);
      if (S.rows.some((x) => x.seq === seq) || (c.pend || []).some((x) => x.seq === seq)) throw dup();
      const r = mk(params, seq);
      if (c.inTx) c.pend.push(r); else S.rows.push(r);
      return { rows: [] };
    }
    if (/^INSERT INTO audit_events \(ts, kind, action, operator, target, detail, ip, ua, extra\) VALUES/i.test(s)) {
      if (params.length !== 9) throw new Error('INSERT 參數數量不對');
      S.rows.push(mk(params, null));
      return { rows: [] };
    }
    let m = /^SELECT count\(\*\)::int AS n FROM audit_events WHERE seq IS NOT NULL$/i.exec(s);
    if (m) return { rows: [{ n: S.rows.filter((x) => x.seq != null).length }] };
    if (/^SELECT id, ts, kind, action, operator, target, detail, ip, ua, extra, seq, prev_hash, row_hash, kid FROM audit_events WHERE seq = \$1::bigint$/i.test(s)) {
      return { rows: S.rows.filter((x) => x.seq === Number(params[0])).map(out) };
    }
    m = /^SELECT id, ts, kind, action, operator, target, detail, ip, ua, extra, seq, prev_hash, row_hash, kid FROM audit_events WHERE seq IS NOT NULL AND seq >= \$1::bigint( AND seq <= \$2::bigint)? ORDER BY seq ASC LIMIT \$(\d+)$/i.exec(s);
    if (m) {
      const lim = params[Number(m[2]) - 1];
      if (lim > 1000) throw new Error('LIMIT 超過 1000：' + lim);
      return { rows: S.rows.filter((x) => x.seq != null && x.seq >= Number(params[0]) && (!m[1] || x.seq <= Number(params[1]))).sort((a, b) => a.seq - b.seq).slice(0, lim).map(out) };
    }
    if (/^WITH f AS \(SELECT min\(id\) AS m FROM audit_events WHERE seq IS NOT NULL\)/i.test(s)) {
      const first = S.rows.filter((x) => x.seq != null).reduce((a, x) => (a == null || x.id < a ? x.id : a), null);
      const un = S.rows.filter((x) => x.seq == null);
      return { rows: [{ legacy: un.filter((x) => first == null || x.id < first).length, after: un.filter((x) => first != null && x.id > first).length }] };
    }
    if (opt.extra) { const r = opt.extra(s, params, S); if (r) return r; }
    throw new Error('假 pg 不認得的語句：' + s.slice(0, 140));
  }
  const api = {
    S, log, opt,
    async query(sql, params) { return exec({ pool: true, inTx: false, pend: [] }, sql, params); },
    async connect() {
      if (opt.failConnect) throw new Error('simulated connect failure');
      const c = { inTx: false, pend: [] };
      return { query: (sql, params) => exec(c, sql, params), release() { release(c); } };
    },
  };
  return api;
}

// ═════════════ 兩種後端的統一介面（讓同一組竄改情境跑兩遍）═════════════
const EVENTS = [];
for (let i = 1; i <= 12; i++) EVENTS.push({ ts: new Date(Date.UTC(2026, 9, 1, 0, i)).toISOString(), kind: i % 5 === 0 ? 'contact' : 'log', action: 'ACT_' + i, operator: 'op' + (i % 3), target: '目標 ' + i, detail: '內容 ' + i + ' 😀 \u0001', ip: '10.0.0.' + i, ua: 'UA/' + i, extra: { n: i, nested: { z: [1, 'x', null], a: { k: 'v' } } } });

function pgBackend(cfg, shared, opt) {
  const fake = makeChainPg(shared, opt);
  const ad = Store.pgAdapter({ query: fake.query, connect: fake.connect, chain: cfg === undefined ? CFG : cfg, legacyTtlMs: 0 });
  return {
    kind: 'pg', fake, ad,
    edit(fn) { const recs = fake.S.rows; const res = fn(recs); if (res) fake.S.rows = res; },
    rows: () => fake.S.rows,
    mkRaw(r) { return Object.assign({ id: ++fake.S.id, ts: new Date(r.ts || Date.now()), kind: 'log', action: 'FORGED', operator: 'x', target: '', detail: '', ip: '', ua: '', extra: {}, seq: null, prev_hash: null, row_hash: null, kid: null }, r); },
  };
}
function jsonlBackend(cfg, fileArg) {
  const file = fileArg || path.join(TMP, 'jsonl-' + rnd(4) + '.jsonl');
  const ad = Store.jsonlAdapter({ file, chain: cfg === undefined ? CFG : cfg });
  const load = () => fs.readFileSync(file, 'utf8').split('\n').filter(Boolean).map((l) => { const o = JSON.parse(l); if ('cseq' in o) { o.seq = o.cseq; delete o.cseq; } return o; });
  const save = (recs) => fs.writeFileSync(file, recs.map((o) => { o = Object.assign({}, o); if (o.seq != null) { o.cseq = o.seq; } delete o.seq; return JSON.stringify(o); }).join('\n') + (recs.length ? '\n' : ''));
  return {
    kind: 'jsonl', file, ad,
    edit(fn) { const recs = load(); const res = fn(recs); save(res || recs); },
    rows: () => load(),
    mkRaw(r) { return Object.assign({ id: crypto.randomUUID(), ts: r.ts || new Date().toISOString(), kind: 'log', action: 'FORGED', operator: 'x', target: '', detail: '', ip: '', ua: '', extra: {} }, r); },
  };
}
const BACKENDS = { pg: pgBackend, jsonl: jsonlBackend };

async function fill(be, n, evs) {
  const list = evs || EVENTS.slice(0, n || 12);
  for (const e of list) await be.ad.insert(Store.normalizeEvent(e));
}
const V = (be, opts, cfg) => AC.verifyChain(be.ad, cfg || CFG, opts);

// ═════════════ U1 設定 ═════════════
async function secU1() {
  console.log('[U1] canonical／設定');
  eq(AC.canonical({ b: 1, a: [2, { d: 1, c: undefined }], s: '中' }), '{"a":[2,{"d":1}],"b":1,"s":"中"}', 'canonical：鍵排序、略過 undefined');
  ok(AC.canonical({ x: { b: 1, a: 2 } }) === AC.canonical({ x: { a: 2, b: 1 } }), 'canonical：與 jsonb 重排鍵順序無關');
  ok(AC.canonical(NaN) === 'null' && AC.canonical(undefined) === 'null', 'canonical：NaN／undefined → null');
  const errs = [];
  let c;
  errs.push(...await captureErr(async () => { c = AC.getConfig({}); }));
  ok(!c.enabled && !c.canVerify && c.disabledReason === 'OFF', '旗標未設 → 停用且沒有金鑰也不吵（不印警告）');
  ok(errs.length === 0, '旗標未設時不印任何訊息');
  c = AC.getConfig({ AUDIT_CHAIN: '0', AUDIT_CHAIN_SECRET: SECRET });
  ok(!c.enabled && c.canVerify, 'AUDIT_CHAIN=0 → 不上鏈，但有金鑰仍可驗證既有的鏈');
  c = AC.getConfig({ AUDIT_CHAIN: 'true', AUDIT_CHAIN_SECRET: SECRET });
  ok(!c.enabled, '只認 AUDIT_CHAIN=1（"true" 不算）');
  let m = await captureErr(async () => { c = AC.getConfig({ AUDIT_CHAIN: '1' }); });
  ok(!c.enabled && c.disabledReason === 'NO_KEY' && m.some(x => x.includes('AUDIT_CHAIN_DISABLED_NO_KEY')), '旗標開但無任何金鑰 → 拒絕啟用並記 AUDIT_CHAIN_DISABLED_NO_KEY（不丟例外）');
  m = await captureErr(async () => { c = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: 'short', JWT_SECRET: rnd(20) }); });
  ok(!c.enabled && c.disabledReason === 'KEY_TOO_SHORT', 'AUDIT_CHAIN_SECRET 太短 → 停用（不悄悄改用 JWT_SECRET）');
  c = AC.getConfig({ AUDIT_CHAIN: '1', JWT_SECRET: rnd(20) });
  ok(c.enabled && c.secretSource === 'JWT_SECRET', '沒有專用金鑰 → 備援 JWT_SECRET');
  c = AC.getConfig({ AUDIT_CHAIN: '1', SESSION_SECRET: rnd(20) });
  ok(c.enabled && c.secretSource === 'SESSION_SECRET', '再備援 SESSION_SECRET');
  const sA = rnd(20);
  const k1 = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sA }).keyFor('k1');
  const k2 = AC.getConfig({ AUDIT_CHAIN: '1', JWT_SECRET: sA }).keyFor('k1');
  ok(k1.equals(k2), '同一把 secret 衍生出相同金鑰（來源不同不影響）');
  ok(!k1.equals(crypto.createHmac('sha256', sA).update('x').digest()) && k1.length === 32, '衍生金鑰 32 位元組，且不是 secret 本身的直接雜湊（網域分隔）');
  c = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sA, AUDIT_CHAIN_KID: 'k2', AUDIT_CHAIN_SECRETS_JSON: J({ k1: rnd(20), bad: 'x', 'we ird': rnd(20) }) });
  ok(c.kid === 'k2' && c.kids.sort().join() === 'k1,k2' && c.keyFor('k1') && c.keyFor('k2') && !c.keyFor('bad') && !c.keyFor('we ird'), 'SECRETS_JSON：合法 kid 才收，太短／非法 kid 忽略');
  m = await captureErr(async () => { c = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sA, AUDIT_CHAIN_SECRETS_JSON: '{not json' }); });
  ok(c.enabled && m.some(x => x.includes('AUDIT_CHAIN_SECRETS_JSON')), 'SECRETS_JSON 壞掉 → 忽略並提示，不影響啟用');
  ok(!JSON.stringify(c).includes(SECRET) && !/"keyFor"/.test(JSON.stringify(c)) && Object.keys(c).indexOf('keyFor') < 0, '設定物件序列化後不含金鑰、也看不到 keyFor');
  ok(!J(c).includes(k1.toString('hex')), '序列化輸出不含衍生金鑰');
  ok(AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sA }) === AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sA }), '設定以環境指紋記憶化');
  const kid = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sA, AUDIT_CHAIN_KID: '<bad kid>' }).kid;
  ok(kid === 'k1', '非法 kid 字元 → 回到預設 k1');
}

// ═════════════ U2 交易、鎖、交錯、失敗 ═════════════
async function secU2() {
  console.log('[U2] 假 Postgres：交易 SQL 順序／advisory lock／多實例交錯／失敗處理');
  // 2.1 SQL 順序
  {
    const be = pgBackend();
    await be.ad.insert(Store.normalizeEvent(EVENTS[0]));
    const log = be.fake.log;
    const i = log.findIndex((s, k) => /^BEGIN$/.test(s) && /^SET LOCAL lock_timeout = 4000$/.test(log[k + 1] || ''));   // 跳過補欄位 DDL 那個交易（D-02：DDL 交易的 lock_timeout 是 AUDIT_DDL_LOCK_TIMEOUT_MS）
    const seqSql = log.slice(i, i + 6).map(s => s.replace(/\(.*$/, '').slice(0, 40));
    eq(seqSql[0], 'BEGIN', '[2.1] 第 1 句 BEGIN');
    ok(/^SET LOCAL lock_timeout = 4000$/.test(log[i + 1]), '[2.1] 第 2 句 SET LOCAL lock_timeout（鎖等待有上限）');
    ok(/^SELECT pg_advisory_xact_lock\(\$1::bigint\)$/.test(log[i + 2]), '[2.1] 第 3 句 pg_advisory_xact_lock');
    ok(/^SELECT seq, row_hash FROM audit_events WHERE seq IS NOT NULL ORDER BY seq DESC LIMIT 1$/.test(log[i + 3]), '[2.1] 第 4 句 讀鏈頭');
    ok(/^INSERT INTO audit_events \(ts, kind, action, operator, target, detail, ip, ua, extra, seq, prev_hash, row_hash, kid\)/.test(log[i + 4]), '[2.1] 第 5 句 帶鏈欄位的 INSERT');
    ok(log[i + 5] === 'COMMIT', '[2.1] 第 6 句 COMMIT');
    ok(be.fake.S.ddl.some(s => /ADD COLUMN IF NOT EXISTS seq BIGINT.*prev_hash TEXT.*row_hash TEXT.*kid TEXT/.test(s)) && be.fake.S.ddl.some(s => /CREATE UNIQUE INDEX IF NOT EXISTS audit_events_seq_uidx ON audit_events \(seq\) WHERE seq IS NOT NULL/.test(s)), '[2.1] ensureSchema 補欄位＋seq 唯一索引（冪等語法）');
    ok(log.every(s => !/\b(UPDATE|DELETE|TRUNCATE|DROP)\b/i.test(s)), '[2.1] 仍然沒有 UPDATE／DELETE／TRUNCATE／DROP');
    const r = be.rows()[0];
    ok(r.seq === 1 && r.prev_hash === AC.GENESIS && /^[0-9a-f]{64}$/.test(r.row_hash) && r.kid === 'k1', '[2.1] 第一筆：seq=1、prev_hash=64 個 0、row_hash 64 位十六進位、kid=k1');
  }
  // 2.2 三個實例並行（含隨機延遲）
  {
    const shared = { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] };
    const bes = [pgBackend(CFG, shared, { jitter: 3 }), pgBackend(CFG, shared, { jitter: 3 }), pgBackend(CFG, shared, { jitter: 3 })];
    const ps = [];
    for (let i = 0; i < 90; i++) ps.push(bes[i % 3].ad.insert(Store.normalizeEvent(Object.assign({}, EVENTS[i % 12], { detail: 'c' + i }))));
    await Promise.all(ps);
    const seqs = shared.rows.map(r => r.seq).sort((a, b) => a - b);
    eq(seqs, Array.from({ length: 90 }, (_, i) => i + 1), '[2.2] 3 實例交錯寫入 90 筆：序號 1..90 無缺號、無重複');
    ok(bes.every(b => b.ad.chainStats().breakRisk === 0), '[2.2] 沒有任何一筆因衝突而降級（鎖讓寫入序列化）');
    const byId = shared.rows.slice().sort((a, b) => a.seq - b.seq);
    ok(byId.every((r, i) => r.prev_hash === (i ? byId[i - 1].row_hash : AC.GENESIS)), '[2.2] 每一筆 prev_hash 都接上一筆 row_hash');
    const v = await AC.verifyChain(bes[0].ad, CFG, {});
    ok(v.ok && v.checked === 90 && v.headSeq === 90, '[2.2] 驗證通過：90 筆、鏈頭 90');
    // D-02：補欄位 DDL 也在交易內（lock_timeout 非 4000、沒有 advisory lock），不計入「每個鏈寫入交易都有鎖」
    const perTx = bes.every(b => {
      const l = b.fake.log;
      const ddlTx = l.filter(s => /^SET LOCAL lock_timeout = \d+$/.test(s) && !/= 4000$/.test(s)).length;
      return l.filter(s => /^BEGIN$/.test(s)).length === l.filter(s => /^COMMIT$/.test(s)).length && l.filter(s => /pg_advisory_xact_lock/.test(s)).length === l.filter(s => /^BEGIN$/.test(s)).length - ddlTx;
    });
    ok(perTx, '[2.2] 每個交易都有 advisory lock，BEGIN／COMMIT 成對');
  }
  // 2.3 兩實例強制交錯（無延遲，純靠鎖排序）
  {
    const shared = { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] };
    const a = pgBackend(CFG, shared), b = pgBackend(CFG, shared);
    await Promise.all([a.ad.insert(Store.normalizeEvent(EVENTS[0])), b.ad.insert(Store.normalizeEvent(EVENTS[1])), a.ad.insert(Store.normalizeEvent(EVENTS[2])), b.ad.insert(Store.normalizeEvent(EVENTS[3]))]);
    eq(shared.rows.map(r => r.seq).sort(), [1, 2, 3, 4], '[2.3] 兩實例同時寫入 → 序號 1..4');
    ok((await AC.verifyChain(a.ad, CFG, {})).ok, '[2.3] 驗證通過');
  }
  // 2.4 連線失敗 → 仍寫入（不上鏈）＋ AUDIT_CHAIN_BREAK_RISK
  {
    const shared = { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] };
    const good = pgBackend(CFG, shared);
    await good.ad.insert(Store.normalizeEvent(EVENTS[0]));
    const bad = pgBackend(CFG, shared, { failConnect: true });
    let msgs = await captureErr(async () => { await bad.ad.insert(Store.normalizeEvent(EVENTS[1])); });
    ok(shared.rows.length === 2 && shared.rows[1].seq == null, '[2.4] 取不到交易連線 → 事件仍寫入（未上鏈），沒有遺失');
    ok(msgs.some(x => x.includes('AUDIT_CHAIN_BREAK_RISK')) && bad.ad.chainStats().breakRisk === 1, '[2.4] 記 AUDIT_CHAIN_BREAK_RISK 並計數');
    await good.ad.insert(Store.normalizeEvent(EVENTS[2]));
    const v = await AC.verifyChain(good.ad, CFG, {});
    ok(v.ok && v.checked === 2 && v.unchainedAfterStart === 1 && v.legacyUncovered === 0, '[2.4] 鏈繼續接上去；未上鏈那一筆被報告為 unchainedAfterStart=1（警示，不算斷鏈）');
  }
  // 2.5 交易中途失敗 → ROLLBACK、釋放鎖、降級寫入、後續正常
  {
    const shared = { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] };
    const be = pgBackend(CFG, shared, { failChainInsert: 1 });
    await captureErr(async () => { await be.ad.insert(Store.normalizeEvent(EVENTS[0])); });
    ok(shared.rows.length === 1 && shared.rows[0].seq == null, '[2.5] 鏈 INSERT 失敗 → 該筆改為不上鏈寫入');
    ok(be.fake.log.includes('ROLLBACK') && shared.lockHolder == null && shared.waiters.length === 0, '[2.5] 有送 ROLLBACK，advisory lock 已釋放（不會卡住後面的寫入）');
    await be.ad.insert(Store.normalizeEvent(EVENTS[1]));
    ok(shared.rows[1].seq === 1, '[2.5] 下一筆正常上鏈（seq=1）');
  }
  // 2.6 DDL 失敗（非表擁有者）→ 整個新表不退回舊路徑，只是不上鏈
  {
    const be = pgBackend(CFG, undefined, { failAlter: true });
    let msgs = await captureErr(async () => { await be.ad.insert(Store.normalizeEvent(EVENTS[0])); await be.ad.insert(Store.normalizeEvent(EVENTS[1])); });
    ok(be.rows().length === 2 && be.rows().every(r => r.seq == null), '[2.6] 補欄位失敗 → 事件照常寫入（不上鏈）');
    ok(msgs.some(x => x.includes('AUDIT_CHAIN_BREAK_RISK')) && be.ad.chainStats().ddlFailed === true, '[2.6] 記 AUDIT_CHAIN_BREAK_RISK、stats.ddlFailed=true');
  }
  // 2.7 金鑰遺失／序號異常 → 降級不丟事件
  {
    const bad = { enabled: true, kid: 'k1', keyFor: () => null, canVerify: true };
    const be = pgBackend(bad);
    const m = await captureErr(async () => { await be.ad.insert(Store.normalizeEvent(EVENTS[0])); });
    ok(be.rows().length === 1 && be.rows()[0].seq == null && m.some(x => x.includes('CHAIN_KEY_MISSING')), '[2.7] 金鑰不可用 → 不上鏈寫入，不丟例外');
    const be2 = pgBackend(); await be2.ad.insert(Store.normalizeEvent(EVENTS[0]));
    be2.edit((rows) => { rows[0].row_hash = 'zz'; });
    const m2 = await captureErr(async () => { await be2.ad.insert(Store.normalizeEvent(EVENTS[1])); });
    ok(be2.rows().length === 2 && m2.some(x => x.includes('CHAIN_HEAD_CORRUPT')), '[2.7] 鏈頭被破壞（row_hash 非 64 位十六進位）→ 新事件降級寫入並告警');
  }
  // 2.8 完整 store 層：insert() 回傳 true；facade stats 才有 chain
  {
    const be = pgBackend();
    const st = Store.createAuditStore(be.ad, { retryDelayMs: 1, fallbackFile: path.join(TMP, 'fb.jsonl') });
    ok((await st.insert(EVENTS[0])) === true, '[2.8] createAuditStore.insert 回傳 true');
    const fac = Store.createAuditFacade({ enabled: true, chain: CFG, tableStore: st, legacyStore: st, legacyWrite: () => true });
    await fac.init();
    ok(fac.stats().chain && fac.stats().chain.enabled && fac.stats().chain.sealed === 1, '[2.8] facade.stats().chain 顯示 sealed=1');
    ok((await fac.chainSource()) === be.ad, '[2.8] facade.chainSource() 回傳轉接器（新表模式）');
    const fac2 = Store.createAuditFacade({ enabled: false, chain: CFG, tableStore: st, legacyStore: st, legacyWrite: () => true });
    ok((await fac2.chainSource()) === null && !('chain' in fac2.stats()), '[2.8] 旗標關（legacy 模式）：chainSource=null、stats 沒有 chain 欄位');
  }
  // 2.9 分批 ≤1000
  {
    const be = pgBackend();
    const evs = []; for (let i = 0; i < 2300; i++) evs.push({ ts: new Date(Date.UTC(2026, 9, 1) + i * 1000).toISOString(), kind: 'log', action: 'BULK', operator: 'o', target: 't', detail: 'd' + i, extra: { i } });
    for (const e of evs) await be.ad.insert(Store.normalizeEvent(e));
    be.fake.log.length = 0;
    const v = await AC.verifyChain(be.ad, CFG, {});
    const pages = be.fake.log.filter(s => /ORDER BY seq ASC LIMIT/.test(s));
    ok(v.ok && v.checked === 2300 && pages.length === 3, `[2.9] 2300 筆分 3 批讀取（實際 ${pages.length} 批）`);
    const v2 = await AC.verifyChain(be.ad, CFG, { chunk: 99999 });
    ok(v2.ok, '[2.9] chunk 參數超過 1000 也被限縮（fake 對 LIMIT>1000 會丟錯）');
  }
}

// ═════════════ U3 竄改偵測（兩種後端）═════════════
async function secU3() {
  for (const name of Object.keys(BACKENDS)) {
    const mk = BACKENDS[name];
    console.log('[U3] 竄改偵測：' + name);
    const T = (s) => `[${name}] ${s}`;
    const fresh = async (n) => { const be = mk(); await fill(be, n); return be; };
    let be = await fresh();
    let v = await V(be);
    ok(v.ok && v.checked === 12 && v.headSeq === 12 && v.firstBreak === null && v.complete && v.legacyUncovered === 0 && v.unchainedAfterStart === 0, T('乾淨的鏈：通過、12 筆、鏈頭 12'));
    ok(be.rows().every(r => /^[0-9a-f]{64}$/.test(r.row_hash)), T('每列都有 64 位十六進位 row_hash'));
    // 改欄位
    for (const [f, val] of [['detail', '被改過'], ['ts', pgTs(be, '2026-01-01T00:00:00.000Z')], ['operator', 'someone_else'], ['action', 'OTHER'], ['target', 'x'], ['ip', '6.6.6.6'], ['ua', 'evil'], ['kind', 'quote-edit'], ['extra', { n: 99 }]]) {
      be = await fresh();
      be.edit((rows) => { const r = rows.find(x => x.seq === 5); r[f] = f === 'ts' && name === 'pg' ? new Date(val) : val; });
      v = await V(be);
      ok(!v.ok && v.firstBreak && v.firstBreak.seq === 5 && v.firstBreak.reason === 'HASH_MISMATCH' && v.checked === 4, T(`改 ${f} → 在 seq 5 斷鏈（HASH_MISMATCH），前 4 筆通過`));
    }
    // extra 微小改動（巢狀）
    be = await fresh();
    be.edit((rows) => { const r = rows.find(x => x.seq === 3); r.extra = JSON.parse(JSON.stringify(r.extra)); r.extra.nested.z[1] = 'y'; });
    v = await V(be);
    ok(!v.ok && v.firstBreak.seq === 3, T('改 extra 內巢狀陣列的一個字 → 斷在 seq 3'));
    // 只是 jsonb 重排鍵順序 → 仍然通過
    be = await fresh();
    be.edit((rows) => { rows.forEach(r => { const e = r.extra; r.extra = JSON.parse(J({ nested: { a: e.nested.a, z: e.nested.z }, n: e.n })); }); });
    v = await V(be);
    ok(v.ok, T('extra 鍵順序被資料庫重排（jsonb）不影響驗證'));
    // 刪中間列
    be = await fresh();
    be.edit((rows) => rows.filter(x => x.seq !== 6));
    v = await V(be);
    ok(!v.ok && v.firstBreak.reason === 'SEQ_GAP' && v.firstBreak.seq === 6 && v.checked === 5, T('刪掉 seq 6 → SEQ_GAP 在 6'));
    // 刪首列
    be = await fresh();
    be.edit((rows) => rows.filter(x => x.seq !== 1));
    v = await V(be);
    ok(!v.ok && v.firstBreak.reason === 'SEQ_GAP' && v.firstBreak.seq === 1, T('刪掉第一列 → SEQ_GAP 在 1'));
    // 插入偽造列（接在鏈尾，雜湊亂填）
    be = await fresh();
    be.edit((rows) => { rows.push(be.mkRaw({ seq: 13, prev_hash: rows[rows.length - 1].row_hash, row_hash: rnd(32), kid: 'k1', detail: '偽造' })); });
    v = await V(be);
    ok(!v.ok && v.firstBreak.seq === 13 && v.firstBreak.reason === 'HASH_MISMATCH' && v.checked === 12, T('在尾端插入偽造列（亂填雜湊）→ 斷在 13'));
    // 插入沒有鏈欄位的偽造列：不算斷鏈但被計入警示
    be = await fresh();
    be.edit((rows) => { rows.splice(6, 0, be.mkRaw({ detail: '直接寫入、沒有鏈欄位' })); });
    v = await V(be);
    ok(v.ok && v.unchainedAfterStart === 1 && v.legacyUncovered === 0, T('鏈中間被直接塞入一列沒有鏈欄位 → 通過，但 unchainedAfterStart=1 警示'));
    // 換順序（交換 seq 4、5 兩列）
    be = await fresh();
    be.edit((rows) => { const a = rows.find(x => x.seq === 4), b = rows.find(x => x.seq === 5); if (name === 'jsonl') { const ia = rows.indexOf(a), ib = rows.indexOf(b); rows[ia] = b; rows[ib] = a; } else { a.seq = 5; b.seq = 4; } });
    v = await V(be);
    ok(!v.ok && v.firstBreak.seq <= 5, T('交換兩列順序 → 斷在 ≤5'));
    // 截尾：無錨點偵測不到（已知限制）；有錨點偵測得到
    be = await fresh();
    const anchor = AC.makeAnchor(CFG, await be.ad.chainHead(), await be.ad.chainCount());
    be.edit((rows) => rows.filter(x => x.seq <= 9));
    v = await V(be);
    ok(v.ok && v.headSeq === 9, T('截掉尾端 3 筆且不帶錨點 → 驗證仍通過（已知限制：必須靠錨點）'));
    v = await V(be, { anchor });
    ok(!v.ok && v.anchor && v.anchor.reason === 'TRUNCATED' && v.firstBreak.reason === 'TRUNCATED', T('帶錨點驗證 → 抓到截尾（TRUNCATED）'));
    // 截尾後補寫新事件（回溯重寫）
    await fill(be, 3, EVENTS.slice(0, 3).map(e => Object.assign({}, e, { detail: '回溯後重寫' + e.detail })));
    v = await V(be, { anchor });
    ok(v.checked === 12 && !v.ok && v.anchor.reason === 'ANCHOR_MISMATCH', T('截尾後再由系統正常補寫到同一個序號 → 鏈本身通過，但與舊錨點比對不符（ANCHOR_MISMATCH）'));
    // 錨點有效且資料完整 → 通過
    be = await fresh();
    const anc2 = AC.makeAnchor(CFG, await be.ad.chainHead(), await be.ad.chainCount());
    await fill(be, 2, EVENTS.slice(0, 2));
    v = await V(be, { anchor: anc2 });
    ok(v.ok && v.anchor && v.anchor.ok && v.checked === 14, T('舊錨點之後又新增事件 → 比對通過（鏈只增不改）'));
    // 錨點被改過
    const forged = Object.assign({}, anc2, { headHash: rnd(32) });
    v = await V(be, { anchor: forged });
    ok(!v.ok && v.anchor.reason === 'ANCHOR_HMAC_INVALID', T('錨點 headHash 被改 → ANCHOR_HMAC_INVALID'));
    const otherKeyAnchor = AC.makeAnchor(AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: rnd(32) }), await be.ad.chainHead(), 14);
    v = await V(be, { anchor: otherKeyAnchor });
    ok(!v.ok && v.anchor.reason === 'ANCHOR_HMAC_INVALID', T('用別的金鑰簽的錨點 → ANCHOR_HMAC_INVALID'));
    // 錯誤金鑰
    be = await fresh();
    v = await V(be, {}, AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: rnd(32) }));
    ok(!v.ok && v.firstBreak.seq === 1 && v.firstBreak.reason === 'HASH_MISMATCH', T('用錯誤金鑰驗證 → 第 1 筆就失敗'));
    // kid 被換
    be = await fresh();
    be.edit((rows) => { rows.find(x => x.seq === 4).kid = 'k9'; });
    v = await V(be);
    ok(!v.ok && v.firstBreak.seq === 4 && v.firstBreak.reason === 'UNKNOWN_KID', T('kid 被改成不存在的代號 → UNKNOWN_KID'));
    // prev_hash 被改
    be = await fresh();
    be.edit((rows) => { rows.find(x => x.seq === 7).prev_hash = rnd(32); });
    v = await V(be);
    ok(!v.ok && v.firstBreak.seq === 7 && v.firstBreak.reason === 'PREV_MISMATCH', T('prev_hash 被改 → PREV_MISMATCH'));
    // 區間驗證與分批
    be = await fresh();
    v = await V(be, { from: 5, to: 8 });
    ok(v.ok && v.checked === 4 && v.complete === true, T('區間 5..8 只驗 4 筆'));
    be.edit((rows) => { rows.find(x => x.seq === 2).detail = '區間外被改'; });
    v = await V(be, { from: 5, to: 8 });
    ok(v.ok, T('區間外的竄改不影響區間驗證（所以預設要驗整條）'));
    v = await V(be, {});
    ok(!v.ok && v.firstBreak.seq === 2, T('整條驗證抓得到'));
    be = await fresh();
    be.edit((rows) => rows.filter(x => x.seq !== 4));
    v = await V(be, { from: 5, to: 8 });
    ok(!v.ok && v.firstBreak.reason === 'SEQ_GAP', T('區間起點前一筆不存在 → SEQ_GAP'));
    be = await fresh();
    v = await V(be, { maxRows: 5 });
    ok(v.ok && v.checked === 5 && v.complete === false && v.nextFrom === 6, T('maxRows=5 → 只驗 5 筆、complete=false、nextFrom=6'));
    v = await V(be, { from: v.nextFrom });
    ok(v.ok && v.complete && v.checked === 7, T('從 nextFrom 接續驗完'));
    // 空鏈
    const empty = mk();
    v = await V(empty);
    ok(v.ok && v.checked === 0 && v.headSeq === 0 && v.headHash === AC.GENESIS, T('空鏈：通過、headSeq 0'));
    const ea = AC.makeAnchor(CFG, null, 0);
    v = await V(empty, { anchor: ea });
    ok(v.ok && v.anchor.ok, T('空鏈的錨點（seq 0）可比對'));
    // 舊資料涵蓋報告：鏈啟用前已有未上鏈資料
    {
      const legacyBe = mk(null);                      // 沒有鏈設定的轉接器＝啟用前
      await fill(legacyBe, 3);
      const be2 = name === 'pg' ? Object.assign(pgBackend(CFG, legacyBe.fake.S), {}) : jsonlBackend(CFG, legacyBe.file);
      await fill(be2, 4);
      v = await V(be2);
      ok(v.ok && v.checked === 4 && v.legacyUncovered === 3 && v.unchainedAfterStart === 0, T('啟用前的 3 筆舊資料 → legacyUncovered=3（未納入鏈），其餘通過'));
      ok(be2.rows().slice(0, 3).every(r => r.seq == null), T('舊資料列沒有被補上鏈欄位（不改動舊資料）'));
    }
    // 換金鑰（kid 輪替）
    {
      const sOld = rnd(32), sNew = rnd(32);
      const cOld = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sOld, AUDIT_CHAIN_KID: 'k1' });
      const first = mk(cOld); await fill(first, 5);
      const cNew = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sNew, AUDIT_CHAIN_KID: 'k2', AUDIT_CHAIN_SECRETS_JSON: J({ k1: sOld }) });
      const second = name === 'pg' ? pgBackend(cNew, first.fake.S) : jsonlBackend(cNew, first.file);
      await fill(second, 4, EVENTS.slice(5, 9));
      const kids = second.rows().map(r => r.kid);
      eq(kids, ['k1', 'k1', 'k1', 'k1', 'k1', 'k2', 'k2', 'k2', 'k2'], T('換金鑰後：前 5 列 kid=k1，之後 kid=k2，序號接續'));
      v = await V(second, {}, cNew);
      ok(v.ok && v.checked === 9, T('帶著舊 secret（SECRETS_JSON）驗證 9 筆全部通過'));
      const cLost = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: sNew, AUDIT_CHAIN_KID: 'k2' });
      v = await V(second, {}, cLost);
      ok(!v.ok && v.firstBreak.reason === 'UNKNOWN_KID' && v.firstBreak.seq === 1, T('遺失舊 secret → 舊列 UNKNOWN_KID'));
      const anc = AC.makeAnchor(cNew, await second.ad.chainHead(), 9);
      ok(anc.kid === 'k2' && AC.verifyAnchor(cNew, anc) && !AC.verifyAnchor(cLost, Object.assign({}, anc, { kid: 'k1' })), T('錨點用現行 kid 簽，驗證時依 kid 取金鑰'));
    }
  }
  function pgTs(be, iso) { return iso; }
}

// ═════════════ U4 旗標關閉＝與第二波位元級相同 ═════════════
async function secU4() {
  console.log('[U4] 旗標關閉：寫入與第二波（397abc5）位元級相同');
  let oldSrc = null;
  try { oldSrc = execFileSync('git', ['show', '397abc5:lib/auditStore.js'], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }); } catch { oldSrc = null; }
  if (!oldSrc) { console.log('  ⚠ 未驗證：取不到 git 物件 397abc5（淺層 clone？請 git fetch --unshallow）——「旗標關閉與第二波位元級相同」整段沒有執行'); UNVERIFIED.push('U4 旗標關閉對第二波基底（git 物件 397abc5）'); return; }
  const oldFile = path.join(TMP, 'old-auditStore.js');
  fs.writeFileSync(oldFile, oldSrc);
  const Old = require(oldFile);
  ok(!/auditChain|cseq/.test(oldSrc), '[4.0] 取到的是第二波的舊模組（不含 auditChain）');
  const evs = EVENTS.slice(0, 6).concat([{ ts: '2026-10-02T00:00:00.000Z', kind: 'quote-edit', action: 'DRAFT_SAVE', operator: 'o', target: 'Q', detail: 'x\u0000y', extra: { quoteNo: 'Q', n: { a: 1 } }, id: 'fixed-id-1' }]);
  const offCfgs = { 'chain 未提供': undefined, 'getConfig 旗標未設': AC.getConfig({ AUDIT_CHAIN_SECRET: SECRET }), 'AUDIT_CHAIN=0': AC.getConfig({ AUDIT_CHAIN: '0', AUDIT_CHAIN_SECRET: SECRET }) };
  for (const [label, cfg] of Object.entries(offCfgs)) {
    // PG：SQL 與參數
    const fa = makeOldPg(), fb = makeOldPg();
    const a = Old.createAuditStore(Old.pgAdapter({ query: fa.query }), { retryDelayMs: 1 });
    const b = Store.createAuditStore(Store.pgAdapter({ query: fb.query, connect: fb.connect, chain: cfg }), { retryDelayMs: 1 });
    for (const e of evs) { await a.insert(Object.assign({}, e)); await b.insert(Object.assign({}, e)); }
    ok(J(fa.log) === J(fb.log), `[4.1] (${label}) PG：送出的 SQL 與參數逐筆完全相同（${fa.log.length} 句）`);
    ok(fb.log.every(x => !/ALTER|BEGIN|advisory|seq/i.test(x.sql)), `[4.1] (${label}) PG：沒有 ALTER／交易／鎖`);
    ok(J(Object.keys(a.stats()).sort()) === J(Object.keys(b.stats()).sort()), `[4.1] (${label}) store.stats() 欄位相同`);
    // JSONL：檔案位元組
    const f1 = path.join(TMP, 'old.jsonl'), f2 = path.join(TMP, 'new.jsonl');
    fs.rmSync(f1, { force: true }); fs.rmSync(f2, { force: true });
    const ja = Old.createAuditStore(Old.jsonlAdapter({ file: f1 }), {}), jb = Store.createAuditStore(Store.jsonlAdapter({ file: f2, chain: cfg }), {});
    for (const e of evs) { await ja.insert(Object.assign({ id: 'id-' + e.action }, e)); await jb.insert(Object.assign({ id: 'id-' + e.action }, e)); }
    ok(fs.readFileSync(f1).equals(fs.readFileSync(f2)), `[4.2] (${label}) JSONL：檔案位元組完全相同`);
    ok(!/cseq|row_hash/.test(fs.readFileSync(f2, 'utf8')), `[4.2] (${label}) JSONL：沒有鏈欄位`);
  }
  // facade stats 欄位
  const fac = (S, cfg) => S.createAuditFacade({ enabled: true, chain: cfg, tableStore: S.createAuditStore(S.jsonlAdapter({ file: path.join(TMP, 'fs.jsonl'), chain: cfg })), legacyStore: null, legacyWrite: () => true });
  const f_old = fac(Old), f_new = fac(Store, offCfgs['AUDIT_CHAIN=0']);
  ok(J(Object.keys(f_old.stats()).sort()) === J(Object.keys(f_new.stats()).sort()), '[4.3] facade.stats() 欄位相同（沒有 chain 欄位）');
  // 開啟後才有鏈欄位（對照）
  const f3 = path.join(TMP, 'on.jsonl'); fs.rmSync(f3, { force: true });
  await Store.createAuditStore(Store.jsonlAdapter({ file: f3, chain: CFG }), {}).insert(evs[0]);
  ok(/"cseq":1/.test(fs.readFileSync(f3, 'utf8')), '[4.4] 對照：旗標開啟後才出現 cseq／row_hash');
  function makeOldPg() {
    const log = []; let id = 0;
    return {
      log,
      query: async (sql, params) => { log.push({ sql: sql.replace(/\s+/g, ' ').trim(), params: params || [] }); return { rows: [] }; },
      connect: async () => { throw new Error('旗標關閉時不應該取得交易連線'); },
    };
  }
}

// ═════════════ U5 錨點、Webhook SSRF、Cron ═════════════
async function secU5() {
  console.log('[U5] 錨點／Webhook SSRF／Cron');
  const head = { seq: 42, row_hash: rnd(32) };
  const a = AC.makeAnchor(CFG, head, 42, new Date('2026-10-14T00:10:00Z'));
  ok(a.seq === 42 && a.headHash === head.row_hash && a.count === 42 && a.ts === '2026-10-14T00:10:00.000Z' && a.kid === 'k1' && /^[0-9a-f]{64}$/.test(a.anchorHmac), '[5.1] 錨點欄位：seq／headHash／count／ts／kid／anchorHmac');
  const expectHmac = crypto.createHmac('sha256', CFG.keyFor('k1')).update('anchor|42|' + head.row_hash + '|2026-10-14T00:10:00.000Z').digest('hex');
  ok(a.anchorHmac === expectHmac, '[5.1] anchorHmac = HMAC(鏈金鑰, "anchor|seq|headHash|ts")');
  ok(AC.verifyAnchor(CFG, a), '[5.2] 驗證通過');
  for (const [k, v] of [['seq', 43], ['headHash', rnd(32)], ['ts', '2026-10-15T00:00:00.000Z'], ['anchorHmac', rnd(32)], ['kid', 'k7']]) ok(!AC.verifyAnchor(CFG, Object.assign({}, a, { [k]: v })), `[5.2] 改 ${k} → 驗證失敗`);
  ok(!AC.verifyAnchor(AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: rnd(32) }), a), '[5.2] 不同金鑰 → 驗證失敗');
  ok(!AC.verifyAnchor(CFG, null) && !AC.verifyAnchor(CFG, {}) && !AC.verifyAnchor(CFG, 'x') && !AC.verifyAnchor(CFG, Object.assign({}, a, { seq: '42' })), '[5.2] 形狀不對 → false（不丟例外）');
  ok(!J(a).includes(SECRET), '[5.3] 錨點不含金鑰');

  // SSRF：IP 判斷
  const blocked = ['127.0.0.1', '10.1.2.3', '172.16.0.1', '172.31.255.255', '192.168.1.1', '169.254.169.254', '169.254.0.1', '0.0.0.0', '100.64.0.1', '224.0.0.1', '255.255.255.255', '198.18.0.1', '::1', '::', 'fe80::1', 'fc00::1', 'fd12:3456::1', '::ffff:127.0.0.1', '::ffff:10.0.0.1', '::ffff:7f00:1', '::ffff:a9fe:a9fe', 'ff02::1', '64:ff9b::1', '2001:db8::1', 'not-an-ip'];
  const allowed = ['8.8.8.8', '1.1.1.1', '172.15.0.1', '172.32.0.1', '203.0.114.5', '2606:4700:4700::1111', '::ffff:8.8.8.8'];
  for (const ip of blocked) ok(AC.isBlockedIp(ip), `[5.4] 擋下 ${ip}`);
  for (const ip of allowed) ok(!AC.isBlockedIp(ip), `[5.4] 放行公開位址 ${ip}`);
  // 網址形狀
  for (const [u, why] of [['http://hooks.example.com/x', 'http'], ['ftp://x.example.com', 'ftp'], ['https://user:pw@hooks.example.com/x', '帶帳密'], ['https://hooks.example.com:8443/x', '非 443'], ['javascript:alert(1)', 'js'], ['not a url', '亂字串'], ['', '空']]) {
    ok(!AC.checkWebhookUrlShape(u).ok, `[5.5] 拒絕網址：${why}`);
  }
  ok(AC.checkWebhookUrlShape('https://hooks.example.com/path?token=abc').ok && AC.checkWebhookUrlShape('https://hooks.example.com:443/x').ok, '[5.5] 接受 https（443）');
  // pushAnchor
  const mkReq = (status, calls) => (opts, cb) => {
    calls.push(opts);
    const handlers = {};
    const req = { on(ev, fn) { handlers[ev] = fn; return req; }, end(body) { calls[calls.length - 1].body = body.toString('utf8'); setImmediate(() => { if (status === 'timeout') handlers.timeout && handlers.timeout(); else if (status === 'error') handlers.error && handlers.error(new Error('ECONNRESET secret-host')); else cb({ statusCode: status, resume() {} }); }); }, destroy() { calls[calls.length - 1].destroyed = true; } };
    return req;
  };
  const lookupPub = async () => [{ address: '93.184.216.34' }];
  let calls = [];
  let r = await AC.pushAnchor('http://hooks.example.com/x', a, { lookup: lookupPub, request: mkReq(200, calls) });
  ok(!r.ok && r.reason === 'SCHEME_NOT_HTTPS' && calls.length === 0, '[5.6] http → 拒絕，完全沒發出請求');
  r = await AC.pushAnchor('https://internal.example.com/x', a, { lookup: async () => [{ address: '10.0.0.5' }], request: mkReq(200, calls) });
  ok(!r.ok && r.reason === 'BLOCKED_ADDRESS' && calls.length === 0, '[5.6] 主機名解析到內網 → 拒絕、沒發請求');
  r = await AC.pushAnchor('https://meta.example.com/x', a, { lookup: async () => [{ address: '169.254.169.254' }], request: mkReq(200, calls) });
  ok(!r.ok && r.reason === 'BLOCKED_ADDRESS' && calls.length === 0, '[5.6] 解析到雲端 metadata 位址 → 拒絕');
  r = await AC.pushAnchor('https://mixed.example.com/x', a, { lookup: async () => [{ address: '93.184.216.34' }, { address: '127.0.0.1' }], request: mkReq(200, calls) });
  ok(!r.ok && r.reason === 'BLOCKED_ADDRESS' && calls.length === 0, '[5.6] 多筆位址中只要有一筆是內網 → 整個拒絕（防 DNS 輪替）');
  r = await AC.pushAnchor('https://[::1]/x', a, { lookup: lookupPub, request: mkReq(200, calls) });
  ok(!r.ok && r.reason === 'BLOCKED_ADDRESS' && calls.length === 0, '[5.6] IP 直接寫在網址（::1）→ 拒絕');
  r = await AC.pushAnchor('https://127.0.0.1/x', a, { lookup: lookupPub, request: mkReq(200, calls) });
  ok(!r.ok && r.reason === 'BLOCKED_ADDRESS' && calls.length === 0, '[5.6] IP 直接寫在網址（127.0.0.1）→ 拒絕');
  r = await AC.pushAnchor('https://nx.example.com/x', a, { lookup: async () => { throw new Error('ENOTFOUND'); }, request: mkReq(200, calls) });
  ok(!r.ok && r.reason === 'DNS_FAILED' && calls.length === 0, '[5.6] DNS 失敗 → 回報失敗不丟例外');
  calls = [];
  r = await AC.pushAnchor('https://hooks.example.com/path?token=SECRETTOKEN', a, { lookup: lookupPub, request: mkReq(204, calls) });
  ok(r.ok && calls.length === 1, '[5.7] 公開位址 → 成功送出');
  ok(calls[0].host === '93.184.216.34' && calls[0].servername === 'hooks.example.com' && calls[0].headers.Host === 'hooks.example.com' && calls[0].port === 443 && calls[0].method === 'POST' && calls[0].timeout <= 5000, '[5.7] 直接連到已驗證的 IP（防 rebinding），SNI／Host 用原主機名，443，POST，逾時 ≤5 秒');
  ok(calls[0].path === '/path?token=SECRETTOKEN' && JSON.parse(calls[0].body).anchorHmac === a.anchorHmac && !calls[0].body.includes(SECRET), '[5.7] 送出的內容是錨點 JSON（含簽章，不含金鑰）');
  calls = [];
  r = await AC.pushAnchor('https://hooks.example.com/x', a, { lookup: lookupPub, request: mkReq(302, calls) });
  ok(!r.ok && r.reason === 'REDIRECT_NOT_FOLLOWED' && calls.length === 1, '[5.8] 回 302 → 不跟隨、回報失敗（只發出 1 次請求）');
  r = await AC.pushAnchor('https://hooks.example.com/x', a, { lookup: lookupPub, request: mkReq(500, calls) });
  ok(!r.ok && r.reason === 'HTTP_500', '[5.8] 回 500 → 失敗');
  r = await AC.pushAnchor('https://hooks.example.com/x', a, { lookup: lookupPub, request: mkReq('timeout', calls) });
  ok(!r.ok && r.reason === 'TIMEOUT' && calls[calls.length - 1].destroyed === true, '[5.8] 逾時 → 失敗並中止連線');
  r = await AC.pushAnchor('https://hooks.example.com/x', a, { lookup: lookupPub, request: mkReq('error', calls) });
  ok(!r.ok && r.reason === 'REQUEST_FAILED' && !J(r).includes('secret-host'), '[5.8] 連線錯誤 → 失敗，不外洩底層錯誤訊息');
  r = await AC.pushAnchor('https://hooks.example.com/x', a, { lookup: lookupPub, request: () => { throw new Error('boom'); } });
  ok(!r.ok && r.reason === 'REQUEST_FAILED', '[5.8] request 丟例外 → 失敗不 throw');
  // Cron 驗證
  const env = { CRON_SECRET: 'cron-' + rnd(10) };
  const rq = (h) => ({ headers: h });
  ok(AC.cronAuthorized(rq({ authorization: 'Bearer ' + env.CRON_SECRET }), env), '[5.9] 正確 Bearer → 通過');
  ok(AC.cronAuthorized(rq({ authorization: 'bearer ' + env.CRON_SECRET }), env), '[5.9] Bearer 大小寫不拘');
  for (const [label, h] of [['沒有標頭', {}], ['錯誤密鑰', { authorization: 'Bearer wrong' }], ['缺 Bearer', { authorization: env.CRON_SECRET }], ['Basic', { authorization: 'Basic ' + env.CRON_SECRET }], ['多餘內容', { authorization: 'Bearer ' + env.CRON_SECRET + ' x' }], ['空 token', { authorization: 'Bearer ' }]]) ok(!AC.cronAuthorized(rq(h), env), `[5.9] ${label} → 拒絕`);
  ok(!AC.cronAuthorized(rq({ authorization: 'Bearer anything' }), {}) && !AC.cronAuthorized(rq({ authorization: 'Bearer ' }), { CRON_SECRET: '' }) && !AC.cronAuthorized(rq({ authorization: 'Bearer undefined' }), { CRON_SECRET: undefined }), '[5.9] 沒設 CRON_SECRET → 一律拒絕');
  // 與寄信 cron 同規則
  const MR = require(path.join(ROOT, 'lib', 'mail', 'routes.js'));
  let same = true;
  for (const h of [{ authorization: 'Bearer ' + env.CRON_SECRET }, { authorization: 'Bearer x' }, {}, { authorization: 'Basic abc' }]) if (MR.cronAuthorized(rq(h), env) !== AC.cronAuthorized(rq(h), env)) same = false;
  ok(same, '[5.9] 與 lib/mail/routes.js 的 cronAuthorized 判定一致');
  const vj = JSON.parse(fs.readFileSync(path.join(ROOT, 'vercel.json'), 'utf8'));
  ok(vj.crons.some(c => c.path === '/api/cron/audit-anchor' && /^\S+ \S+ \* \* \*$/.test(c.schedule)) && vj.crons.some(c => c.path === '/api/cron/mail-outbox'), '[5.10] vercel.json 有每日 audit-anchor cron，且原本的 mail-outbox 還在');
}

// ═════════════ U6 簽核歷程鏈（真實 quoteRoutes）═════════════
const QA = require(path.join(ROOT, 'lib', 'quoteApproval.js'));
const registerQuoteRoutes = require(path.join(ROOT, 'lib', 'quoteRoutes.js'));
const OL = require(path.join(ROOT, 'lib', 'quoteOppLink.js'));
const QUSERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'other1', role: 'user', active: true, displayName: 'Other1', bu: ['ERP'] },
  { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
  { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
];
function mkQEnv(quotations) {
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = { data: { opportunities: [{ id: 'O1', owner: 'own1', stage: 'C', company: 'TestCo', product: 'Prod', amount: '100', bu: 'ITS' }], quotations, quoteApproval: { roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } } } }, seq: 0, saves: 0, logs: [], notes: [] };
  env.auth = { users: JSON.parse(JSON.stringify(QUSERS)) };
  const deps = {
    db: { load: () => env.data, save: () => { env.saves++; }, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (...a) => env.logs.push(a), pushNotification: (...a) => env.notes.push(a),
    getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1' || rl === 'secretary') return ['own1']; return [req.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: async () => Buffer.from(''), QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
  };
  registerQuoteRoutes(app, deps);
  env.call = (user, method, p, params, body, query) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: params || {}, query: query || {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}
const mkQuote = (over) => Object.assign({
  id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00Z', updatedAt: '2026-10-08T00:00:00Z',
  validUntil: '2026-10-30', products: ['PS'], costBy: null, approval: null, discountType: 'none', discountValue: 0,
  items: [{ lid: 'i1', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 600000, cost: 590000 }, { lid: 'i2', desc: '教育訓練', unit: '式', qty: 1, unitPrice: 100000, cost: 95000 }],   // 毛利很低 → 要走到董事長關，歷程才夠長
  costFlow: { state: 'na' },
}, over || {});
async function approveAll(env) {
  const S = () => env.data.quotations[0];
  await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' }, {});
  await env.call('mgr1', 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { comment: 'ok', hash: QA.contentHash(S()) });
  for (const u of ['gm1', 'ch1', 'sec1']) if (S().approval.state === 'pending') await env.call(u, 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { hash: QA.contentHash(S()), comment: 'ok', resolutionDate: '2026-10-10', resolutionNo: 'T-2026-001' });
}
async function withEnv(vars, fn) {
  const saved = {};
  for (const k of Object.keys(vars)) { saved[k] = process.env[k]; if (vars[k] === undefined) delete process.env[k]; else process.env[k] = vars[k]; }
  try { return await fn(); } finally { for (const k of Object.keys(saved)) { if (saved[k] === undefined) delete process.env[k]; else process.env[k] = saved[k]; } }
}
async function secU6() {
  console.log('[U6] 簽核歷程鏈');
  const ON = { AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: SECRET, AUDIT_CHAIN_KID: undefined, AUDIT_CHAIN_SECRETS_JSON: undefined };
  const OFF = { AUDIT_CHAIN: undefined, AUDIT_CHAIN_SECRET: SECRET };
  let histOn = null, histOff = null;
  // 旗標關：與以前完全一樣
  await withEnv(OFF, async () => {
    const env = mkQEnv([mkQuote()]);
    await approveAll(env);
    const q = env.data.quotations[0];
    histOff = JSON.parse(JSON.stringify(q.approval.history));
    ok(q.approval.state === 'approved' && histOff.length >= 4, `[6.1] 旗標關：流程跑完（核准），歷程 ${histOff.length} 筆`);
    ok(histOff.every(e => !('h' in e)) && !('historyBase' in q.approval), '[6.1] 旗標關：歷程沒有 h、approval 沒有 historyBase');
    ok(histOff.every(e => J(Object.keys(e).filter(k => k !== 'meta')) === J(['at', 'by', 'action', 'comment', 'tier'])), '[6.1] 旗標關：每筆欄位集合與順序＝以前（at,by,action,comment,tier[,meta]）');
    const r = await env.call('own1', 'GET', '/api/quotations/:id/history-check', { id: 'Q1' });
    ok(r.s === 200 && r.j.ok === true && r.j.enabled === true && r.j.chainWriting === false && r.j.chained === 0 && r.j.legacy === histOff.length, '[6.1] 旗標關但有金鑰：history-check 可用，全部算舊資料（chained 0、legacy ' + histOff.length + '）、chainWriting=false ' + J(r.j));
  });
  await withEnv({ AUDIT_CHAIN: undefined, AUDIT_CHAIN_SECRET: undefined, JWT_SECRET: undefined, SESSION_SECRET: undefined }, async () => {
    const env = mkQEnv([mkQuote()]);
    await approveAll(env);
    const r = await env.call('own1', 'GET', '/api/quotations/:id/history-check', { id: 'Q1' });
    ok(r.s === 200 && r.j.ok === null && r.j.enabled === false, '[6.1b] 完全沒有金鑰：history-check 回未啟用');
  });
  // 旗標開
  let env;
  await withEnv(ON, async () => {
    env = mkQEnv([mkQuote()]);
    await approveAll(env);
    const q = env.data.quotations[0];
    histOn = JSON.parse(JSON.stringify(q.approval.history));
    ok(q.approval.state === 'approved', '[6.2] 旗標開：流程照常跑完（核准）');
    ok(histOn.length === histOff.length && histOn.every(e => /^k1:[0-9a-f]{64}$/.test(e.h)), `[6.2] 每筆歷程都有 h（格式 k1:<64 hex>），筆數與旗標關相同（${histOn.length}）`);
    ok(histOn.every((e, i) => J(Object.keys(e).filter(k => k !== 'h')) === J(Object.keys(histOff[i]))), '[6.2] 除了多一個 h，欄位集合與順序和旗標關完全相同');
    ok(new Set(histOn.map(e => e.h)).size === histOn.length, '[6.2] 每筆 h 都不同（含前一筆）');
    const ser = (await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' })).j;
    ok(!/"h":"k1:/.test(J(ser)) && !/historyBase/.test(J(ser)) && ser.approval.history.every(x => !('h' in x)), '[6.3] 一般序列化輸出完全不含 h（golden 不受影響）');
    const clean = JSON.parse(J(ser.approval.history)).map(x => Object.keys(x).join());
    ok(clean.every(k => k.startsWith('at,byName,action,comment,tier')), '[6.3] 序列化後歷程鍵＝at,byName,action,comment,tier(,items…)');
    // history-check 權限
    let r = await env.call('own1', 'GET', '/api/quotations/:id/history-check', { id: 'Q1' });
    ok(r.s === 200 && r.j.ok === true && r.j.entries === histOn.length && r.j.chained === histOn.length && r.j.firstBreak === null, '[6.4] 業務本人：200，ok，全部納入鏈');
    ok(!/"comment"|"meta"|"by"/.test(J(r.j)), '[6.4] history-check 回應不含歷程內容');
    r = await env.call('mgr1', 'GET', '/api/quotations/:id/history-check', { id: 'Q1' });
    ok(r.s === 200 && r.j.ok === true, '[6.4] 主管（能看這張單）：200');
    r = await env.call('other1', 'GET', '/api/quotations/:id/history-check', { id: 'Q1' });
    ok(r.s === 403, '[6.4] 看不到這張單的人 → 403');
    r = await env.call('own1', 'GET', '/api/quotations/:id/history-check', { id: 'NOPE' });
    ok(r.s === 404, '[6.4] 單不存在 → 404');
    // 竄改
    const orig = JSON.parse(J(q.approval.history));
    const reset = () => { q.approval.history = JSON.parse(J(orig)); };
    const chk = async () => (await env.call('own1', 'GET', '/api/quotations/:id/history-check', { id: 'Q1' })).j;
    q.approval.history[2].comment = '被改過的意見';
    let c = await chk();
    ok(c.ok === false && c.firstBreak.index === 2 && c.firstBreak.reason === 'HASH_MISMATCH' && c.firstBreak.action === orig[2].action, '[6.5] 改某筆意見 → 在該筆斷鏈（index 2）');
    reset(); q.approval.history[1].by = 'someone_else';
    c = await chk(); ok(!c.ok && c.firstBreak.index === 1, '[6.5] 改簽核人 → 在該筆斷鏈');
    reset(); q.approval.history[3].action = 'RETURN';
    c = await chk(); ok(!c.ok && c.firstBreak.index === 3, '[6.5] 改動作（APPROVE→RETURN）→ 斷鏈');
    reset(); q.approval.history[1].meta = Object.assign({}, q.approval.history[1].meta, { forged: 1 });
    c = await chk(); ok(!c.ok && c.firstBreak.index === 1, '[6.5] 改 meta 快照 → 斷鏈');
    reset(); q.approval.history.splice(2, 1);
    c = await chk(); ok(!c.ok && c.firstBreak.index === 2, '[6.5] 刪掉中間一筆 → 斷鏈');
    reset(); { const t = q.approval.history[1]; q.approval.history[1] = q.approval.history[2]; q.approval.history[2] = t; }
    c = await chk(); ok(!c.ok && c.firstBreak.index === 1, '[6.5] 調換兩筆順序 → 斷鏈');
    reset(); q.approval.history.push({ at: new Date().toISOString(), by: 'x', action: 'APPROVE', comment: '偽造', tier: 'gm' });
    c = await chk(); ok(c.ok === true && c.unchained === 1, '[6.5] 尾端塞入沒有 h 的偽造筆 → 不算斷鏈但 unchained=1 警示');
    reset(); q.approval.history.push({ at: 'x', by: 'x', action: 'APPROVE', comment: '偽造', tier: 'gm', h: 'k1:' + rnd(32) });
    c = await chk(); ok(!c.ok && c.firstBreak.index === orig.length, '[6.5] 尾端塞入亂填 h 的偽造筆 → 斷鏈');
    reset(); q.approval.history.pop();
    c = await chk(); ok(c.ok === true, '[6.5] 刪掉最後一筆 → 偵測不到（已知限制：靠稽核日誌交叉比對）');
    reset(); q.approval.history[1].h = 'k9:' + q.approval.history[1].h.slice(3);
    c = await chk(); ok(!c.ok && c.firstBreak.reason === 'UNKNOWN_KID', '[6.5] h 的 kid 被換 → UNKNOWN_KID');
    reset(); q.approval.history.forEach(e => { delete e.h; });
    c = await chk(); ok(c.ok === true && c.chained === 0 && c.legacy === orig.length, '[6.5] 所有 h 都被拿掉 → 視為舊資料（已知限制）');
    reset();
    c = await chk(); ok(c.ok === true, '[6.5] 還原後恢復通過');
    // 錯誤金鑰
    const bad = HC.checkQuote(q, AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: rnd(32) }));
    ok(!bad.ok && bad.firstBreak.index === 0, '[6.5] 用錯誤金鑰檢查 → 第 0 筆就失敗');
    // 舊資料 + 新資料混合
    reset();
    const mixed = JSON.parse(J(orig)); mixed.slice(0, 2).forEach(e => { delete e.h; });
    q.approval.history = mixed;
    c = await chk(); ok(c.ok === false && c.firstBreak.index === 2, '[6.6] 前 2 筆是舊資料（沒 h）但第 3 筆的 h 是接在原本第 2 筆之後 → 斷鏈（舊筆被拿掉 h 也會牽動後面）');
    reset();
  });
  // 混合：舊資料在前，鏈啟用後新增
  await withEnv(OFF, async () => {
    const e2 = mkQEnv([mkQuote()]);
    await e2.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' }, {});     // 旗標關：SUBMIT 沒有 h
    await withEnv(ON, async () => {
      await e2.call('mgr1', 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { comment: 'ok', hash: QA.contentHash(e2.data.quotations[0]) });   // 旗標開：這筆有 h
      const c = (await e2.call('own1', 'GET', '/api/quotations/:id/history-check', { id: 'Q1' })).j;
      ok(c.ok === true && c.legacy === 1 && c.chained === 1 && c.unchained === 0, '[6.7] 鏈啟用前的歷程算舊資料（legacy=1），啟用後的第一筆接在空字串之後（chained=1）' + J(c));
    });
  });
  // quoteOppLink 的歷程也簽名
  await withEnv(ON, async () => {
    const q = mkQuote({ oppId: 'O1', approval: { state: 'approved', history: [] } });
    const data = { quotations: [q], opportunities: [] };
    const hmod = require(path.join(ROOT, 'lib', 'quoteOppLink.js'));
    HC.append(q.approval, { at: 't', by: 'u', action: 'X', comment: '', tier: '' });
    const before = q.approval.history.length;
    // 呼叫模組內部同等路徑：HC.append 無 max
    HC.append(q.approval, { at: 't2', by: 'u', action: 'OPP_VOID_INVALIDATE', comment: 'c', tier: '' });
    ok(q.approval.history.length === before + 1 && q.approval.history.every(e => /^k1:/.test(e.h)) && HC.checkQuote(q).ok, '[6.8] HC.append 無上限時照樣簽名且可驗證');
    ok(fs.readFileSync(path.join(ROOT, 'lib', 'quoteOppLink.js'), 'utf8').includes('HC.append(ap,'), '[6.8] quoteOppLink 的 pushHist 已改走 HC.append（OPP_VOID／OPP_RESTORE 歷程也簽名）');
    ok(typeof hmod.invalidateForOpp === 'function', '[6.8] quoteOppLink 仍可載入');
  });
  // 超過上限截斷：historyBase 接續
  await withEnv(ON, async () => {
    const ap = { history: [] };
    for (let i = 0; i < 12; i++) HC.append(ap, { at: 't' + i, by: 'u', action: 'A' + i, comment: '', tier: '' }, 5);
    ok(ap.history.length === 5 && /^k1:/.test(ap.historyBase), '[6.9] 超過上限截斷最舊的筆後，記下被截掉那筆的 h 為 historyBase');
    ok(HC.verifyHistory(ap.history, CFG, ap.historyBase).ok, '[6.9] 帶 historyBase 驗證通過');
    ok(!HC.verifyHistory(ap.history, CFG, '').ok, '[6.9] 不帶 historyBase → 第 0 筆失敗（所以 base 不能漏）');
    const c = HC.checkQuote({ approval: ap });
    ok(c.ok, '[6.9] checkQuote 會自動帶 historyBase');
    ap.history[0].comment = 'x';
    ok(!HC.checkQuote({ approval: ap }).ok, '[6.9] 截斷後的第一筆被改 → 仍抓得到（沒有留下漏洞）');
  });
  await withEnv(OFF, async () => {
    const ap = { history: [] };
    for (let i = 0; i < 12; i++) HC.append(ap, { at: 't' + i, by: 'u', action: 'A' + i, comment: '', tier: '' }, 5);
    ok(ap.history.length === 5 && !('historyBase' in ap) && ap.history.every(e => !('h' in e)) && ap.history[0].action === 'A7', '[6.10] 旗標關：截斷行為與原本相同（保留最新 5 筆、無額外欄位）');
  });
  // 全部報價單檢查
  await withEnv(ON, async () => {
    const e3 = mkQEnv([mkQuote(), mkQuote({ id: 'Q2', quoteNo: 'QU-2' }), mkQuote({ id: 'Q3', quoteNo: 'QU-3', approval: null })]);
    await approveAll(e3);
    const q2 = e3.data.quotations[1];
    await e3.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q2' }, {});
    let r = await HC.checkAll(e3.data, CFG);
    ok(r.ok && r.quotes === 3 && r.entries > 0 && r.breaks.length === 0, '[6.11] checkAll：3 張單全部通過');
    q2.approval.history[0].comment = '改';
    r = await HC.checkAll(e3.data, CFG);
    ok(!r.ok && r.breaks.length === 1 && r.firstBreak.quoteNo === 'QU-2' && r.firstBreak.index === 0, '[6.11] checkAll：抓到 QU-2 第 0 筆');
    ok(!J(r).includes('改') || !/"comment"/.test(J(r)), '[6.11] checkAll 結果不含歷程內容');
    const big = []; for (let i = 0; i < 650; i++) big.push(mkQuote({ id: 'B' + i, quoteNo: 'B' + i, approval: { state: 'pending', history: [] } }));
    big.forEach(q => HC.append(q.approval, { at: 't', by: 'u', action: 'SUBMIT', comment: '', tier: '' }));
    r = await HC.checkAll({ quotations: big }, CFG, { yieldEvery: 100 });
    ok(r.ok && r.quotes === 650, '[6.11] checkAll 分批處理 650 張單');
  });
}

// ═════════════ U7 UI 渲染 ═════════════
function secU7() {
  console.log('[U7] 版本資訊頁「稽核防竄改」卡');
  const html = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
  const extractFn = (name) => {
    const i = html.indexOf('function ' + name + '(');
    if (i < 0) throw new Error('找不到 ' + name);
    let j = html.indexOf('{', i), depth = 0, k = j;
    for (; k < html.length; k++) { if (html[k] === '{') depth++; else if (html[k] === '}') { depth--; if (!depth) break; } }
    return html.slice(i, k + 1);
  };
  const ci = html.indexOf('const CHAIN_REASON = {'); const ce = html.indexOf('};', ci) + 2;
  ok(ci > 0 && /id="vChainInfo"/.test(html), '[7.0] admin.html 有 vChainInfo 卡與 CHAIN_REASON');
  const R = new Function(extractFn('escH') + '\n' + html.slice(ci, ce) + '\n' + extractFn('chainReasonText') + '\n' + extractFn('chainFmtTs') + '\n' + extractFn('verChainHtml') + '\nreturn { verChainHtml, escH, chainReasonText };')();
  const X = '<img src=x onerror=alert(1)>', Q = '"><script>alert(2)</script>';
  const base = { enabled: true, configured: true, canVerify: true, reason: '', mode: 'table', kid: 'k1', kids: ['k1'], secretSource: 'AUDIT_CHAIN_SECRET', headSeq: 321, headHash: 'a'.repeat(64), writes: { enabled: true, sealed: 5, breakRisk: 0 }, lastVerify: null, lastAnchor: null, anchorsStored: 0, webhook: { configured: false, host: '' }, cronSecretConfigured: true };
  let h = R.verChainHtml(base);
  ok(/運作中/.test(h) && />321</.test(h) && /立即驗證/.test(h) && /下載錨點/.test(h) && /尚未驗證過/.test(h), '[7.1] 運作中：顯示鏈頭 321、兩個按鈕、尚未驗證過');
  h = R.verChainHtml(base, { ok: true, checked: 321, complete: true, firstBreak: null, ts: '2026-10-14T01:02:03.000Z', headSeq: 321, legacyUncovered: 7, unchainedAfterStart: 0, anchor: null });
  ok(/驗證通過/.test(h) && /2026-10-14 01:02:03 UTC/.test(h) && /已檢查 321 筆/.test(h) && /未納入鏈（舊資料）：7 筆/.test(h) && /#137333/.test(h), '[7.2] 驗證通過：時間、筆數、舊資料 7 筆未納入鏈');
  h = R.verChainHtml(base, { ok: false, checked: 41, complete: true, firstBreak: { seq: 42, reason: 'HASH_MISMATCH', detail: '這一筆的內容與簽章不符' }, ts: '2026-10-14T01:02:03.000Z', headSeq: 321, legacyUncovered: 0, unchainedAfterStart: 2, anchor: { checked: true, ok: false } });
  ok(/驗證失敗/.test(h) && /序號 42/.test(h) && /內容與簽章不符/.test(h) && /含錨點比對/.test(h) && /#c5221f/.test(h) && />驗證失敗</.test(h), '[7.3] 驗證失敗：斷點序號 42、原因、錨點比對、紅色晶片');
  ok(/鏈啟用後仍有 2 筆沒有簽章/.test(h), '[7.3] 顯示 unchainedAfterStart 警示');
  h = R.verChainHtml(Object.assign({}, base, { lastVerify: { ts: '2026-10-13T00:00:00Z', ok: false, checked: 9, firstBreak: { seq: 3, reason: 'SEQ_GAP' } } }));
  ok(/驗證失敗/.test(h) && /序號 3/.test(h) && /序號不連續/.test(h), '[7.4] 沿用伺服器記錄的上次驗證結果（失敗）');
  h = R.verChainHtml(base, { error: '連線失敗' });
  ok(/驗證沒有完成：連線失敗/.test(h), '[7.5] 驗證請求失敗 → 顯示錯誤');
  h = R.verChainHtml({ enabled: false, configured: false, canVerify: false, reason: 'OFF', mode: 'legacy', kid: 'k1', headSeq: 0, lastVerify: null, lastAnchor: null, anchorsStored: 0, webhook: { configured: false }, cronSecretConfigured: false });
  ok(/未啟用/.test(h) && /AUDIT_CHAIN=1/.test(h) && !/立即驗證/.test(h) && /未設定（排程不會執行）/.test(h), '[7.6] 未啟用：顯示啟用步驟、沒有驗證按鈕');
  h = R.verChainHtml(Object.assign({}, base, { enabled: false, configured: true, canVerify: false, reason: 'KEY_TOO_SHORT' }));
  ok(/無法啟用/.test(h) && /金鑰太短/.test(h), '[7.7] 旗標開但金鑰太短：顯示「無法啟用」與原因');
  h = R.verChainHtml(Object.assign({}, base, { writes: { breakRisk: 3 } }));
  ok(/有事件未上鏈/.test(h) && /#c5221f/.test(h), '[7.8] breakRisk>0 → 警示晶片與紅字');
  h = R.verChainHtml(Object.assign({}, base, { lastAnchor: { ts: '2026-10-14T00:10:00Z', seq: 300 }, anchorsStored: 12, webhook: { configured: true, valid: true, host: 'hooks.example.com' } }));
  ok(/2026-10-14 00:10:00 UTC/.test(h) && /序號 300/.test(h) && /已保存 12 份/.test(h) && /已設定（hooks\.example\.com）/.test(h), '[7.9] 最近錨點時間／序號／份數／外部推送主機');
  h = R.verChainHtml(Object.assign({}, base, { webhook: { configured: true, valid: false, reason: 'SCHEME_NOT_HTTPS' } }));
  ok(/設定無效/.test(h), '[7.9] webhook 設定無效時提示');
  // XSS
  h = R.verChainHtml(Object.assign({}, base, { kid: X, secretSource: Q, webhook: { configured: true, valid: true, host: X }, lastAnchor: { ts: X, seq: X }, writes: { breakRisk: X }, headSeq: X }),
    { ok: false, checked: X, complete: false, ts: X, firstBreak: { seq: X, reason: X, detail: Q }, legacyUncovered: X, unchainedAfterStart: X, anchor: {} });
  ok(!/<img|<script/i.test(h) && !/<[a-z]+[^>]*onerror/i.test(h), '[7.10] 所有欄位帶 <img onerror>／<script> → 輸出沒有未跳脫標籤');
  ok(h.includes('&lt;img src=x onerror=alert(1)&gt;') || /&lt;img/.test(h), '[7.10] 以跳脫後的文字顯示');
  h = R.verChainHtml({ configured: true, canVerify: true, enabled: true, kid: 'k1', reason: X, webhook: { configured: true, valid: false, reason: X } });
  ok(!/<img/i.test(h), '[7.11] reason／webhook.reason 帶標籤 → 跳脫');
  ok(R.verChainHtml(null).includes('未提供'), '[7.12] 沒有資料 → 友善提示');
  ok(!/onclick="[^"]*(\$|\+|\{)/.test(R.verChainHtml(base)), '[7.13] 按鈕 onclick 不帶任何動態資料');
  // 載入點接線
  ok(/loadChainCard\(\);\s*\n/.test(html) && /id="vChainInfo"/.test(html), '[7.14] loadVersionInfo 會載入卡片');
  // 內嵌 script 語法
  const re = /<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/g; let m, bad = 0;
  while ((m = re.exec(html))) { try { new Function(m[1]); } catch (e) { bad++; } }
  ok(bad === 0, '[7.15] admin.html 所有內嵌 <script> 語法正確（new Function）');
}

// ═════════════ U8 變異檢查 ═════════════
async function coreChecks(ACm, Stm) {
  const res = {};
  const cfg = ACm.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: SECRET });
  const mkPg = (shared, c) => { const f = makeChainPg(shared); return { f, ad: Stm.pgAdapter({ query: f.query, connect: f.connect, chain: c === undefined ? cfg : c }) }; };
  const ins = async (ad, n, tag) => { for (let i = 0; i < n; i++) await ad.insert(Stm.normalizeEvent(Object.assign({}, EVENTS[i % 12], { detail: (tag || 'd') + i }))); };
  // c1 乾淨
  { const b = mkPg(); await ins(b.ad, 6); const v = await ACm.verifyChain(b.ad, cfg, {}); res.c1_clean = v.ok && v.checked === 6; }
  // c2 改欄位
  { const b = mkPg(); await ins(b.ad, 6); b.f.S.rows[2].detail = 'x'; const v = await ACm.verifyChain(b.ad, cfg, {}); res.c2_edit = !v.ok && v.firstBreak.seq === 3; }
  // c3 錯誤金鑰
  { const b = mkPg(); await ins(b.ad, 4); const v = await ACm.verifyChain(b.ad, ACm.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: rnd(32) }), {}); res.c3_wrongKey = !v.ok; }
  // c4 prev 銜接：第 4 列的 prev_hash 換掉並用真金鑰重算出「自己合法」的 row_hash
  { const b = mkPg(); await ins(b.ad, 6); const r = b.f.S.rows[3]; r.prev_hash = 'a'.repeat(64);
    r.row_hash = ACm.computeRowHash(cfg.keyFor('k1'), r.prev_hash, ACm.hashFields({ kind: r.kind, action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip, ua: r.ua, ts: r.ts.toISOString(), extra: r.extra }, r.seq, r.kid));
    const v = await ACm.verifyChain(b.ad, cfg, {}); res.c4_prev = !v.ok; }
  // c5 序號連續：第 4 列換成 seq=9，但雜湊、prev 都合法
  { const b = mkPg(); await ins(b.ad, 4); const rows = b.f.S.rows; const r = rows[3]; const prev = rows[2].row_hash; r.seq = 9; r.prev_hash = prev;
    r.row_hash = ACm.computeRowHash(cfg.keyFor('k1'), prev, ACm.hashFields({ kind: r.kind, action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip, ua: r.ua, ts: r.ts.toISOString(), extra: r.extra }, 9, r.kid));
    const v = await ACm.verifyChain(b.ad, cfg, {}); res.c5_seq = !v.ok && v.firstBreak && v.firstBreak.seq === 4; }
  // c6 錨點：截尾 → TRUNCATED；截尾後補寫 → ANCHOR_MISMATCH
  { const b = mkPg(); await ins(b.ad, 6); const anc = ACm.makeAnchor(cfg, await b.ad.chainHead(), 6);
    b.f.S.rows = b.f.S.rows.filter(r => r.seq <= 4); const v1 = await ACm.verifyChain(b.ad, cfg, { anchor: anc });
    await ins(b.ad, 2, 'new'); const v2 = await ACm.verifyChain(b.ad, cfg, { anchor: anc });
    res.c6_truncated = !v1.ok && v1.anchor.reason === 'TRUNCATED'; res.c6_mismatch = !v2.ok && v2.anchor.reason === 'ANCHOR_MISMATCH'; }
  // c7 並行
  { const shared = { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] }; const bs = [1, 2, 3].map(() => { const f = makeChainPg(shared, { jitter: 2 }); return Stm.pgAdapter({ query: f.query, connect: f.connect, chain: cfg }); });
    const ps = []; for (let i = 0; i < 30; i++) ps.push(bs[i % 3].insert(Stm.normalizeEvent(Object.assign({}, EVENTS[i % 12], { detail: 'p' + i }))));
    await Promise.all(ps); const seqs = shared.rows.map(r => r.seq).sort((a, b) => a - b);
    res.c7_concurrent = J(seqs) === J(Array.from({ length: 30 }, (_, i) => i + 1)) && bs.every(a => a.chainStats().breakRisk === 0); }
  // c8 旗標
  res.c8_flag = ACm.getConfig({ AUDIT_CHAIN_SECRET: SECRET }).enabled === false && ACm.getConfig({ AUDIT_CHAIN: '0', AUDIT_CHAIN_SECRET: SECRET }).enabled === false && ACm.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: SECRET }).enabled === true;
  // c9 錨點簽章
  { const a = ACm.makeAnchor(cfg, { seq: 3, row_hash: rnd(32) }, 3); res.c9_anchorSig = ACm.verifyAnchor(cfg, a) && !ACm.verifyAnchor(cfg, Object.assign({}, a, { headHash: rnd(32) })); }
  // c10 canonical
  res.c10_canon = ACm.canonical({ b: 1, a: 2 }) === ACm.canonical({ a: 2, b: 1 });
  return res;
}
async function secU8() {
  console.log('[U8] 變異檢查（刻意弄壞，必須被抓到）');
  const base = await coreChecks(AC, Store);
  const allOk = Object.values(base).every(Boolean);
  ok(allOk, '[8.0] 基準：未變異的程式通過全部 ' + Object.keys(base).length + ' 項核心檢查 ' + J(base));
  const muts = [
    { name: 'prev_hash 銜接檢查拿掉', file: 'auditChain.js', from: "if (r.prev_hash !== prevHash) {", to: 'if (false) {', expect: 'c4_prev' },
    { name: 'HMAC 不使用金鑰（用固定金鑰）', file: 'auditChain.js', from: 'function computeRowHash(key, prevHash, fields) { return hmacHex(key,', to: 'function computeRowHash(key, prevHash, fields) { return hmacHex(Buffer.alloc(32),', expect: 'c3_wrongKey' },
    { name: 'advisory lock 拿掉', file: 'auditStore.js', from: "await client.query('SELECT pg_advisory_xact_lock($1::bigint)', [AC.LOCK_KEY]);", to: '', expect: 'c7_concurrent' },
    { name: '序號連續檢查拿掉', file: 'auditChain.js', from: 'if (seq !== expected) {', to: 'if (false) {', expect: 'c5_seq' },
    { name: '錨點雜湊比對拿掉', file: 'auditChain.js', from: 'else if (!safeEq(String(row.row_hash), String(a.headHash))) {', to: 'else if (false) {', expect: 'c6_mismatch' },
    { name: '錨點截尾判斷拿掉', file: 'auditChain.js', from: "if (!row) { ar.ok = false; ar.reason = 'TRUNCATED'; }", to: 'if (!row) { }', expect: 'c6_truncated' },
    { name: '旗標被忽略（只要有金鑰就上鏈）', file: 'auditChain.js', from: 'const enabled = flag && !!secret;', to: 'const enabled = !!secret;', expect: 'c8_flag' },
    { name: '錨點簽章不驗證', file: 'auditChain.js', from: 'return safeEq(hmacHex(key, anchorText(a.seq, a.headHash, a.ts)), a.anchorHmac);', to: 'return true;', expect: 'c9_anchorSig' },
    { name: 'canonical 不排序鍵', file: 'auditChain.js', from: ".filter((k) => v[k] !== undefined).sort().map(", to: ".filter((k) => v[k] !== undefined).map(", expect: 'c10_canon' },
    { name: '雜湊內容漏掉 detail', file: 'auditChain.js', from: 'action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip,', to: 'action: r.action, operator: r.operator, target: r.target, ip: r.ip,', expect: 'c2_edit' },
  ];
  let n = 0;
  for (const mt of muts) {
    const dir = path.join(TMP, 'mut' + (++n));
    fs.mkdirSync(dir);
    for (const f of ['auditChain.js', 'auditStore.js']) {
      let src = fs.readFileSync(path.join(ROOT, 'lib', f), 'utf8');
      if (f === mt.file) { if (!src.includes(mt.from)) { ok(false, `[8] 變異「${mt.name}」找不到要替換的原始碼（測試過期）`); continue; } src = src.replace(mt.from, () => mt.to); }
      fs.writeFileSync(path.join(dir, f), src);
    }
    let r;
    try { r = await coreChecks(require(path.join(dir, 'auditChain.js')), require(path.join(dir, 'auditStore.js'))); } catch (e) { r = { [mt.expect]: false, _threw: e.message }; }
    ok(r[mt.expect] === false, `[8] 變異「${mt.name}」被抓到（${mt.expect} 失敗）`);
  }
}

// ═════════════ U9 verify-audit-pg.js --chain 區段（用假庫自測，確保使用者在真庫執行時不會因手誤而白跑）═════════════
async function secU9() {
  console.log('[U9] scripts/verify-audit-pg.js 的 [13] 區段（假庫自測；真庫由使用者自行執行）');
  const VP = require(path.join(ROOT, 'scripts', 'verify-audit-pg.js'));
  ok(typeof VP.chainSection === 'function', '[9.0] verify-audit-pg.js 匯出 chainSection');
  const norm = (s) => s.replace(/\s+/g, ' ').trim();
  const shared = { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] };
  const marked = (r) => r.extra && r.extra.verifyChain === '1';
  const clone = (rows) => rows.map((r) => Object.assign({}, r, { extra: JSON.parse(JSON.stringify(r.extra)), ts: new Date(r.ts) }));
  const handler = (s, p, S) => {
    let m;
    if (/^SELECT count\(\*\)::int AS n FROM audit_events WHERE seq IS NOT NULL AND coalesce\(extra->>'verifyChain', ''\) <> '1'$/.test(s)) return { rows: [{ n: S.rows.filter((r) => r.seq != null && !marked(r)).length }] };
    if (/^SELECT column_name, data_type FROM information_schema\.columns/.test(s)) return { rows: S.chainCols ? [['seq', 'bigint'], ['prev_hash', 'text'], ['row_hash', 'text'], ['kid', 'text']].map(([c, d]) => ({ column_name: c, data_type: d })) : [] };
    if (/^SELECT indexdef FROM pg_indexes/.test(s)) return { rows: S.chainCols ? [{ indexdef: 'CREATE UNIQUE INDEX audit_events_seq_uidx ON public.audit_events USING btree (seq) WHERE (seq IS NOT NULL)' }] : [] };
    if (/^SELECT seq, prev_hash, row_hash, kid FROM audit_events WHERE seq IS NOT NULL ORDER BY seq$/.test(s)) return { rows: S.rows.filter((r) => r.seq != null).sort((a, b) => a.seq - b.seq).map((r) => ({ seq: String(r.seq), prev_hash: r.prev_hash, row_hash: r.row_hash, kid: r.kid })) };
    if (/^SELECT kid FROM audit_events WHERE seq > \$1::bigint ORDER BY seq$/.test(s)) return { rows: S.rows.filter((r) => r.seq != null && r.seq > Number(p[0])).sort((a, b) => a.seq - b.seq).map((r) => ({ kid: r.kid })) };
    if (/^DELETE\sFROM audit_events WHERE extra->>'verifyChain' = '1' RETURNING id$/.test(s)) { const d = S.rows.filter(marked); S.rows = S.rows.filter((r) => !marked(r)); return { rows: d.map((r) => ({ id: r.id })) }; }
    if (/^SELECT count\(\*\)::int AS n FROM audit_events WHERE extra->>'verifyChain' = '1' OR operator = \$1$/.test(s)) return { rows: [{ n: S.rows.filter((r) => marked(r) || r.operator === p[0]).length }] };
    if ((m = /^DELETE\sFROM audit_events WHERE seq = (\d+) AND operator = \$1$/.exec(s))) { S.rows = S.rows.filter((r) => r.seq !== Number(m[1])); return { rows: [] }; }
    if ((m = /^DELETE\sFROM audit_events WHERE seq > \$1::bigint AND extra->>'verifyChain' = '1'$/.exec(s))) { S.rows = S.rows.filter((r) => !(r.seq != null && r.seq > Number(p[0]) && marked(r))); return { rows: [] }; }
    if ((m = /^UPDATE\saudit_events SET (.+) WHERE seq = (-?\d+)(?: AND operator = \$1)?$/.exec(s))) {
      const r = S.rows.find((x) => x.seq === Number(m[2])); if (!r) return { rows: [] };
      const set = m[1];
      if (/^detail = '被竄改'$/.test(set)) r.detail = '被竄改';
      else if (/^ts = ts \+ interval '1 second'$/.test(set)) r.ts = new Date(r.ts.getTime() + 1000);
      else if (/^operator = operator \|\| 'x'$/.test(set)) r.operator += 'x';
      else if (/^extra = jsonb_set\(extra, '\{i\}', '999'::jsonb\)$/.test(set)) r.extra = Object.assign({}, r.extra, { i: 999 });
      else if (/^prev_hash = repeat\('a', 64\)$/.test(set)) r.prev_hash = 'a'.repeat(64);
      else if (/^kid = 'nope'$/.test(set)) r.kid = 'nope';
      else if (/^seq = (-?\d+)$/.test(set)) r.seq = Number(/^seq = (-?\d+)$/.exec(set)[1]);
      else throw new Error('假庫不認得的 UPDATE：' + set);
      return { rows: [] };
    }
    if (/^INSERT INTO audit_events \(kind, action, operator, extra, seq, prev_hash, row_hash, kid\) VALUES/.test(s)) {
      S.rows.push({ id: ++S.id, ts: new Date(), kind: 'log', action: 'CH_FORGED', operator: p[0], target: '', detail: '', ip: '', ua: '', extra: JSON.parse(p[1]), seq: Number(p[2]), prev_hash: 'b'.repeat(64), row_hash: 'c'.repeat(64), kid: 'vfy1' });
      return { rows: [] };
    }
    if (/^INSERT INTO audit_events \(kind, action, operator, extra\) VALUES/.test(s)) { S.rows.push({ id: ++S.id, ts: new Date(), kind: 'log', action: 'CH_DIRECT', operator: p[0], target: '', detail: '', ip: '', ua: '', extra: JSON.parse(p[1]), seq: null, prev_hash: null, row_hash: null, kid: null }); return { rows: [] }; }
    return undefined;
  };
  const fake = makeChainPg(shared, { extra: handler });
  const snaps = [];
  const mkCx = (q) => ({
    async query(sql, params) {
      const s = norm(sql); let m;
      if (s === 'BEGIN') { snaps.push({ name: '__tx', rows: clone(shared.rows) }); return { rows: [] }; }
      if ((m = /^SAVEPOINT (\w+)$/.exec(s))) { snaps.push({ name: m[1], rows: clone(shared.rows) }); return { rows: [] }; }
      if ((m = /^ROLLBACK TO SAVEPOINT (\w+)$/.exec(s))) { const i = snaps.map((x) => x.name).lastIndexOf(m[1]); shared.rows = clone(snaps[i].rows); snaps.length = i; return { rows: [] }; }
      if (s === 'ROLLBACK') { const i = snaps.map((x) => x.name).indexOf('__tx'); if (i >= 0) { shared.rows = clone(snaps[i].rows); snaps.length = i; } return { rows: [] }; }
      return q(sql, params);
    },
    release() {},
  });
  const cx = mkCx(fake.query);
  const lines = [];
  let skipped = '';
  const mkLoad = (S, h, enabledAlways) => (flag, extraEnv) => {
    const cfg = AC.getConfig(Object.assign({}, extraEnv || {}));
    const f = makeChainPg(S, { extra: h });
    const ts = Store.createAuditStore(Store.pgAdapter({ query: f.query, connect: f.connect, chain: cfg }), { retryDelayMs: 1, fallbackFile: path.join(TMP, 'u9-fb.jsonl') });
    const fac = Store.createAuditFacade({ enabled: flag === '1', chain: cfg, tableStore: ts, legacyStore: ts, legacyWrite: () => true });
    return { auditStore: fac, insertAuditEvent: (ev) => fac.insert(ev), flush: () => fac.flush() };
  };
  const ctx = {
    pool: { connect: async () => cx, query: fake.query }, q: async (sql, p) => (await fake.query(sql, p)).rows,
    say: (l) => lines.push(l), skip: (n, why) => { skipped = n + why; },
    check: (name, cond, detail) => ok(!!cond, '[9.1 假庫] ' + name + (cond ? '' : ' ← ' + J(detail).slice(0, 200))),
    loadRepo: mkLoad(shared, handler), runId: 'vTEST', op: 'vfy_vTEST', AS: Store, AC,
  };
  const before = pass;
  await VP.chainSection(ctx);
  ok(!skipped, '[9.1] 沒有被略過 ' + skipped);
  ok(pass - before >= 30, `[9.1] [13] 區段在假庫上跑了 ${pass - before} 項檢查（≥30）`);
  ok(shared.rows.length === 0, '[9.2] 跑完後假庫沒有任何殘留資料列');
  ok(lines.some((l) => /100 筆併發寫入共 \d+ ms/.test(l)), '[9.3] 有輸出併發耗時（效能資訊）');
  // 已有真實的鏈 → 整段略過且不動資料
  const shared2 = { rows: [], id: 0, ddl: [], chainCols: true, lockHolder: null, waiters: [] };
  const f2 = makeChainPg(shared2, { extra: handler });
  shared2.rows.push({ id: 1, ts: new Date(), kind: 'log', action: 'REAL', operator: 'someone', target: '', detail: '', ip: '', ua: '', extra: {}, seq: 1, prev_hash: AC.GENESIS, row_hash: 'd'.repeat(64), kid: 'k1' });
  let sk = '';
  await VP.chainSection(Object.assign({}, ctx, { pool: { connect: async () => cx, query: f2.query }, q: async (s2, p) => (await f2.query(s2, p)).rows, skip: (n, why) => { sk = n + why; }, check: () => { throw new Error('不該執行任何檢查'); }, loadRepo: () => { throw new Error('不該載入程式'); } }));
  ok(/非本腳本產生的鏈資料/.test(sk) && shared2.rows.length === 1, '[9.4] 資料庫已有非本腳本產生的鏈資料 → 整段略過、完全不寫入、不刪除');
  // 保護 SQL 已套用（UPDATE 被拒）→ 以被拒絕為證
  const shared3 = { rows: [], id: 0, ddl: [], chainCols: false, lockHolder: null, waiters: [] };
  const handler3 = (s, p, S) => { if (/^UPDATE\saudit_events SET detail/.test(s)) throw new Error('audit_events 是只增不改的稽核表'); return handler(s, p, S); };
  const f3 = makeChainPg(shared3, { extra: handler3 });
  const cx3 = mkCx(f3.query);
  const lines3 = [];
  let c3 = 0;
  await VP.chainSection(Object.assign({}, ctx, { pool: { connect: async () => cx3, query: f3.query }, q: async (s2, p) => (await f3.query(s2, p)).rows, say: (l) => lines3.push(l), check: (n, cond) => { c3++; ok(!!cond, '[9.5 保護庫] ' + n); }, loadRepo: mkLoad(shared3, handler3) }));
  ok(lines3.some((l) => /已套用保護 SQL/.test(l)) && c3 > 10, '[9.5] 資料表已套用保護：UPDATE 被拒 → 改以被拒絕為證、其餘檢查照常（' + c3 + ' 項）');
}

// ═════════════ E 端對端 ═════════════
const PW = 'Ac' + rnd(9) + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
let child = null, childPid = null, serverOut = '';
const SERVER_SECRETS = { jwt: rnd(24), chain: rnd(32), cron: rnd(16) };
function seedFiles(opts) {
  opts = opts || {};
  const users = [mkUser('_s_admin', 'admin'), mkUser('_s_u1', 'user', { bu: ['ERP'] })];
  const d = { contacts: [], companies: [], opportunities: [], visits: [], contracts: [], quotations: opts.quotations || [], receivables: [], callins: [], keyAccounts: [], groups: [] };
  if (opts.anchors) d.auditAnchors = opts.anchors;
  fs.writeFileSync(path.join(ROOT, 'auth.json'), J({ users }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), J(d));
  if (opts.preJsonl) fs.writeFileSync(path.join(ROOT, 'audit.log.jsonl'), opts.preJsonl);
}
function cleanupFiles() { FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }
async function startServer(extraEnv) {
  const base = Object.assign({}, process.env);
  for (const k of ['AUDIT_EVENTS_TABLE', 'AUDIT_CHAIN', 'AUDIT_CHAIN_SECRET', 'AUDIT_CHAIN_KID', 'AUDIT_CHAIN_SECRETS_JSON', 'AUDIT_ANCHOR_WEBHOOK_URL', 'CRON_SECRET', 'DATABASE_URL', 'MAIL_MODE', 'SESSION_SECRET']) delete base[k];
  const env = Object.assign(base, extraEnv || {}, { PORT: String(PORT), JWT_SECRET: SERVER_SECRETS.jwt, SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  await assertPortFree(PORT);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
  child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid; serverOut = '';
  child.stdout.on('data', d => { serverOut += d; }); child.stderr.on('data', d => { serverOut += d; });
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + serverOut.slice(-500));
  }
  throw new Error('伺服器啟動逾時：' + serverOut.slice(-500));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100);
  child = null;
}
let jar = {};
const bodies = [];      // 所有回應內容（金鑰洩漏掃描用）
async function http(who, method, url, body, hdr) {
  const h = Object.assign({ 'Content-Type': 'application/json' }, hdr || {});
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : J(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  const text = buf.toString('utf8');
  bodies.push(text);
  let json = null; try { json = JSON.parse(text); } catch { /* 非 JSON */ }
  return { status: res.status, json, text, headers: res.headers, setCookie: res.headers.getSetCookie ? res.headers.getSetCookie() : [] };
}
async function login(who) {
  const r = await http(null, 'POST', '/api/login', { username: who, password: PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
const jsonlLines = () => { try { return fs.readFileSync(path.join(ROOT, 'audit.log.jsonl'), 'utf8').split('\n').filter(Boolean).map(l => JSON.parse(l)); } catch { return []; } };
const writeJsonl = (recs) => fs.writeFileSync(path.join(ROOT, 'audit.log.jsonl'), recs.map(o => J(o)).join('\n') + '\n');
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));

async function secE() {
  const CHAIN_ENV = { AUDIT_EVENTS_TABLE: '1', AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: SERVER_SECRETS.chain, CRON_SECRET: SERVER_SECRETS.cron };
  const cfgSrv = AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: SERVER_SECRETS.chain });
  // 預先存在的舊（未上鏈）事件 3 筆
  const pre = [1, 2, 3].map(i => J({ id: 'pre-' + i, ts: '2026-01-0' + i + 'T00:00:00.000Z', kind: 'log', action: 'PRE_CHAIN', operator: 'old', target: '', detail: '啟用前 ' + i, ip: '', ua: '', extra: {} })).join('\n') + '\n';
  // 簽核歷程（供 admin chain-check）：用同一把金鑰簽
  const qa = await withEnv({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: SERVER_SECRETS.chain, AUDIT_CHAIN_KID: undefined, AUDIT_CHAIN_SECRETS_JSON: undefined }, async () => {
    const mkq = (id, no) => { const q = mkQuote({ id, quoteNo: no, approval: { state: 'pending', history: [] } }); ['SUBMIT', 'APPROVE', 'APPROVE'].forEach((a, i) => HC.append(q.approval, { at: '2026-10-0' + (i + 1) + 'T00:00:00Z', by: 'u' + i, action: a, comment: 'c' + i, tier: '' })); return q; };
    return [mkq('QA', 'QE-1'), mkq('QB', 'QE-2')];
  });
  const dummy = Array.from({ length: 95 }, (_, i) => ({ v: 1, seq: i, headHash: 'a'.repeat(64), count: i, ts: '2026-01-01T00:00:00.000Z', kid: 'k1', anchorHmac: 'b'.repeat(64) }));
  seedFiles({ preJsonl: pre, quotations: qa, anchors: dummy });
  jar = {}; bodies.length = 0;
  await startServer(CHAIN_ENV);
  console.log('[E1] 端對端：AUDIT_EVENTS_TABLE=1 AUDIT_CHAIN=1（port ' + PORT + '）');
  for (const u of ['_s_admin', '_s_u1']) { const r = await login(u); if (r.status !== 200) throw new Error('登入失敗 ' + u + ' ' + r.status); }
  for (const nm of ['=HYPERLINK("http://x","y")', 'plain_user', '<img src=x onerror=alert(1)>']) await http(null, 'POST', '/api/login', { username: nm, password: 'x' });
  let r = await http('_s_u1', 'POST', '/api/contacts', { name: '稽核乙', company: '沙盒公司甲', taxId: '12345675', bu: 'ERP', email: 'z@sandbox.test' });
  ok([200, 201].includes(r.status), '[E1.0] 建立聯絡人產生更多事件 (' + r.status + ')');
  await sleep(500);
  // 存取控制
  for (const [m, u] of [['GET', '/api/admin/audit-chain/head'], ['GET', '/api/admin/audit-chain/verify'], ['POST', '/api/admin/audit-chain/verify'], ['GET', '/api/admin/audit-chain/anchor'], ['GET', '/api/admin/quotations/chain-check']]) {
    r = await http('_s_u1', m, u, m === 'POST' ? {} : undefined);
    ok(r.status === 403, `[E1.1] 一般使用者 ${m} ${u} → 403（得到 ${r.status}）`);
    r = await http(null, m, u, m === 'POST' ? {} : undefined);
    ok(r.status === 401, `[E1.1] 未登入 ${m} ${u} → 401（得到 ${r.status}）`);
  }
  // head
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/head');
  const head0 = r.json || {};
  ok(r.status === 200 && head0.enabled === true && head0.canVerify === true && head0.mode === 'table' && head0.kid === 'k1' && head0.headSeq > 0 && head0.secretSource === 'AUDIT_CHAIN_SECRET', '[E1.2] head：鏈已啟用、kid=k1、鏈頭>0、金鑰來源名稱（不含金鑰）' + J(head0).slice(0, 200));
  ok(head0.cronSecretConfigured === true && head0.webhook && head0.webhook.configured === false && head0.lastVerify === null && head0.anchorsStored === 95, '[E1.2] head：排程密鑰已設、沒有 webhook、尚無驗證、已存 95 份舊錨點');
  const lines0 = jsonlLines();
  ok(lines0.slice(0, 3).every(l => l.action === 'PRE_CHAIN' && l.cseq === undefined), '[E1.3] 啟用前的 3 筆舊事件維持原樣（沒有鏈欄位）');
  ok(lines0.slice(3).every(l => Number.isInteger(l.cseq)) && lines0.slice(3).map(l => l.cseq).join() === lines0.slice(3).map((_, i) => i + 1).join(), '[E1.3] 啟用後的事件 cseq 為 1..N 連續');
  // verify
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify?storedAnchor=0');   // 95 份是假造的舊錨點（簽章無效），本案例不比對已存錨點（D-14 預設會比對）
  const v1 = r.json || {};
  ok(r.status === 200 && v1.ok === true && v1.checked >= lines0.length - 3 && v1.legacyUncovered === 3 && v1.unchainedAfterStart === 0 && v1.firstBreak === null && /^[0-9a-f]{64}$/.test(v1.headHash), '[E1.4] verify：通過、舊資料 3 筆報告為未納入鏈' + J(v1).slice(0, 250));
  await sleep(400);
  const logs = (await http('_s_admin', 'GET', '/api/admin/logs?action=AUDIT_CHAIN_VERIFY')).json || [];
  ok(Array.isArray(logs) && logs.some(l => l.action === 'AUDIT_CHAIN_VERIFY' && /通過/.test(l.detail)), '[E1.5] 驗證動作本身寫了稽核事件 AUDIT_CHAIN_VERIFY');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/head');
  ok(r.json.lastVerify && r.json.lastVerify.ok === true && r.json.lastVerify.by === '_s_admin', '[E1.5] head 之後帶出上次驗證結果');
  // 區間
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify?storedAnchor=0&from=2&to=4');
  ok(r.json.ok && r.json.checked === 3, '[E1.6] 區間 2..4 驗 3 筆');
  // anchor
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/anchor');
  const anchor1 = r.json;
  ok(r.status === 200 && AC.verifyAnchor(cfgSrv, anchor1) && anchor1.seq > 0 && anchor1.kid === 'k1', '[E1.7] anchor：簽章用測試端獨立算的金鑰驗證通過');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/anchor?download=1');
  ok(/attachment; filename="audit-anchor-\d+-\d{4}-\d{2}-\d{2}\.json"/.test(r.headers.get('content-disposition') || ''), '[E1.7] ?download=1 → Content-Disposition 附件');
  await sleep(300);
  const logs2 = (await http('_s_admin', 'GET', '/api/admin/logs?action=AUDIT_CHAIN_ANCHOR')).json || [];
  ok(logs2.length >= 2, '[E1.7] 取得錨點被稽核（AUDIT_CHAIN_ANCHOR）');
  // 帶錨點驗證（POST／query）
  r = await http('_s_admin', 'POST', '/api/admin/audit-chain/verify', { anchor: anchor1 });
  ok(r.status === 200 && r.json.ok && r.json.anchor && r.json.anchor.ok, '[E1.8] POST body 帶錨點驗證 → 通過');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify?anchor=' + encodeURIComponent(J(anchor1)));
  ok(r.status === 200 && r.json.ok && r.json.anchor.ok, '[E1.8] query 帶錨點驗證 → 通過');
  r = await http('_s_admin', 'POST', '/api/admin/audit-chain/verify', { anchor: Object.assign({}, anchor1, { headHash: 'c'.repeat(64) }) });
  ok(r.status === 200 && r.json.ok === false && r.json.anchor.reason === 'ANCHOR_HMAC_INVALID', '[E1.8] 偽造的錨點 → ANCHOR_HMAC_INVALID');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify?anchor=' + encodeURIComponent('{bad json'));
  ok(r.status === 400 && r.json.code === 'BAD_ANCHOR', '[E1.8] 錨點不是合法 JSON → 400');
  // 竄改檔案 → 驗證抓到
  const orig = fs.readFileSync(path.join(ROOT, 'audit.log.jsonl'), 'utf8');
  let recs = jsonlLines();
  const idx = recs.findIndex(l => l.cseq === 3);
  recs[idx].detail = '被竄改';
  writeJsonl(recs);
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify');
  ok(r.json.ok === false && r.json.firstBreak.seq === 3 && r.json.firstBreak.reason === 'HASH_MISMATCH', '[E1.9] 改檔案內一筆 detail → verify 在 seq 3 失敗');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/head');
  ok(r.json.lastVerify && r.json.lastVerify.ok === false && r.json.lastVerify.firstBreak.seq === 3, '[E1.9] head 反映最近一次失敗');
  fs.writeFileSync(path.join(ROOT, 'audit.log.jsonl'), orig);
  recs = jsonlLines(); writeJsonl(recs.filter(l => l.cseq !== 4));
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify');
  ok(r.json.ok === false && r.json.firstBreak.reason === 'SEQ_GAP' && r.json.firstBreak.seq === 4, '[E1.10] 刪掉一行 → SEQ_GAP');
  fs.writeFileSync(path.join(ROOT, 'audit.log.jsonl'), orig);
  // 截尾 + 錨點：先取一份新錨點（其後只會多一筆「取得錨點」事件），截掉最後 2 行就會少掉錨點那一筆
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/anchor');
  const anchorT = r.json;
  await sleep(400);
  const orig2 = fs.readFileSync(path.join(ROOT, 'audit.log.jsonl'), 'utf8');
  recs = jsonlLines(); writeJsonl(recs.slice(0, recs.length - 2));
  r = await http('_s_admin', 'POST', '/api/admin/audit-chain/verify', { anchor: anchorT });
  ok(r.json.ok === false && r.json.anchor.reason === 'TRUNCATED', '[E1.11] 截掉最後 2 行後用先前下載的錨點驗證 → TRUNCATED');
  fs.writeFileSync(path.join(ROOT, 'audit.log.jsonl'), orig2);
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify?storedAnchor=0');
  ok(r.json.ok === true, '[E1.12] 還原檔案後再次通過');
  // Cron
  r = await http(null, 'GET', '/api/cron/audit-anchor');
  ok(r.status === 401, '[E1.13] cron 沒帶標頭 → 401');
  r = await http(null, 'GET', '/api/cron/audit-anchor', undefined, { Authorization: 'Bearer wrong' });
  ok(r.status === 401, '[E1.13] cron 密鑰錯誤 → 401');
  r = await http('_s_admin', 'GET', '/api/cron/audit-anchor');
  ok(r.status === 401, '[E1.13] 管理員 cookie 不能代替 cron 密鑰 → 401');
  r = await http(null, 'GET', '/api/cron/audit-anchor', undefined, { Authorization: 'Bearer ' + SERVER_SECRETS.cron });
  ok(r.status === 200 && r.json.ok && r.json.enabled && r.json.seq > 0 && r.json.stored === 90 && r.json.pushed === null, '[E1.14] cron 密鑰正確 → 產生錨點；95 份舊錨點被裁成最近 90 份（含新的）；沒設 webhook 不推送' + J(r.json));
  const stored = readData().auditAnchors;
  ok(stored.length === 90 && AC.verifyAnchor(cfgSrv, stored[stored.length - 1]) && stored[stored.length - 1].seq === r.json.seq, '[E1.14] data.json 的 auditAnchors 最後一份是新錨點，簽章有效');
  r = await http(null, 'GET', '/api/cron/audit-anchor', undefined, { Authorization: 'Bearer ' + SERVER_SECRETS.cron });
  ok(r.status === 200 && readData().auditAnchors.length === 90, '[E1.14] 再跑一次仍然 90 份');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/head');
  ok(r.json.anchorsStored === 90 && r.json.lastAnchor && r.json.lastAnchor.seq > 0, '[E1.14] head 帶出最近錨點');
  await sleep(300);
  ok(((await http('_s_admin', 'GET', '/api/admin/logs?action=AUDIT_CHAIN_ANCHOR_CRON')).json || []).length >= 2, '[E1.14] 排程錨點被稽核');
  // 簽核歷程鏈（admin 全站檢查）
  r = await http('_s_admin', 'GET', '/api/admin/quotations/chain-check');
  ok(r.status === 200 && r.json.ok === true && r.json.quotes === 2 && r.json.chained === 6, '[E1.15] 全部報價單歷程鏈檢查：2 張單 6 筆全部通過');
  const d = readData(); d.quotations[1].approval.history[1].comment = '被改過'; fs.writeFileSync(path.join(ROOT, 'data.json'), J(d));
  r = await http('_s_admin', 'GET', '/api/admin/quotations/chain-check');
  ok(r.json.ok === false && r.json.firstBreak.quoteNo === 'QE-2' && r.json.firstBreak.index === 1, '[E1.16] 改了 QE-2 第 1 筆 → 抓到');
  ok(!/被改過/.test(r.text), '[E1.16] 回應不含歷程內容');
  // 整頁 UI 靜態檢查
  r = await http('_s_admin', 'GET', '/admin.html');
  ok(r.status === 200 && r.text.includes('id="vChainInfo"') && r.text.includes('verChainHtml'), '[E1.17] admin.html 由伺服器送出且含防竄改卡');
  // 金鑰洩漏掃描
  await sleep(200);
  const everything = bodies.join('\n') + '\n' + serverOut + '\n' + fs.readFileSync(path.join(ROOT, 'audit.log.jsonl'), 'utf8') + fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8');
  const derived = cfgSrv.keyFor('k1').toString('hex');
  ok(![SERVER_SECRETS.chain, SERVER_SECRETS.jwt, SERVER_SECRETS.cron, derived].some(s => everything.includes(s)), '[E1.18] 所有回應、伺服器輸出、稽核檔、資料檔都不含鏈金鑰／衍生金鑰／JWT／CRON 密鑰');
  ok(!/AUDIT_CHAIN_BREAK_RISK|AUDIT_CHAIN_DISABLED/.test(serverOut), '[E1.19] 伺服器輸出沒有 BREAK_RISK／DISABLED 警示');
  await stopServer();

  // ── E2：旗標關（預設）──
  cleanupFiles(); seedFiles({ preJsonl: pre });
  jar = {}; bodies.length = 0;
  await startServer({ AUDIT_EVENTS_TABLE: '1', CRON_SECRET: SERVER_SECRETS.cron });
  console.log('[E2] 端對端：AUDIT_CHAIN 未設（預設關閉）');
  await login('_s_admin');
  await http(null, 'POST', '/api/login', { username: 'plain_user', password: 'x' });
  await sleep(400);
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/head');
  ok(r.status === 200 && r.json.enabled === false && r.json.configured === false && r.json.canVerify === false && r.json.reason === 'OFF', '[E2.1] head：未啟用（OFF）');
  const l2 = jsonlLines();
  ok(l2.length > 3 && l2.every(l => !('cseq' in l) && !('row_hash' in l) && !('prev_hash' in l) && !('kid' in l)), '[E2.2] 旗標關：audit.log.jsonl 沒有任何鏈欄位');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify');
  ok(r.status === 409 && r.json.code === 'CHAIN_NOT_ENABLED', '[E2.3] verify → 409 CHAIN_NOT_ENABLED');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/anchor');
  ok(r.status === 409, '[E2.3] anchor → 409');
  r = await http(null, 'GET', '/api/cron/audit-anchor', undefined, { Authorization: 'Bearer ' + SERVER_SECRETS.cron });
  ok(r.status === 200 && r.json.ok && r.json.enabled === false, '[E2.4] cron 密鑰正確但鏈未啟用 → 200 enabled:false（不寫錨點）');
  ok(!readData().auditAnchors, '[E2.4] data.json 沒有 auditAnchors');
  r = await http('_s_admin', 'GET', '/api/admin/quotations/chain-check');
  ok(r.status === 200 && r.json.ok === true && r.json.chainWriting === false && r.json.chained === 0, '[E2.5] 歷程鏈全站檢查：旗標關 → 不寫 h（chainWriting=false、chained=0），有金鑰仍可檢查舊資料');
  r = await http('_s_admin', 'GET', '/api/admin/version');
  ok(r.status === 200 && r.json.auditStore === 'table' && !('chain' in (r.json.auditWrites || {})), '[E2.6] /api/admin/version 的 auditWrites 沒有多出 chain 欄位（輸出同以前）');
  await stopServer();

  // ── E3：旗標開但金鑰太短 ──
  cleanupFiles(); seedFiles({});
  jar = {}; bodies.length = 0;
  await startServer({ AUDIT_EVENTS_TABLE: '1', AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: 'sh0rt-k3y', CRON_SECRET: SERVER_SECRETS.cron });
  console.log('[E3] 端對端：AUDIT_CHAIN=1 但金鑰太短（拒絕啟用，不崩潰）');
  await login('_s_admin');
  await sleep(400);
  ok(/AUDIT_CHAIN_DISABLED_NO_KEY/.test(serverOut) && !serverOut.includes('sh0rt-k3y'), '[E3.1] 伺服器記 AUDIT_CHAIN_DISABLED_NO_KEY（且不印出金鑰值）');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/head');
  ok(r.status === 200 && r.json.enabled === false && r.json.configured === true && r.json.reason === 'KEY_TOO_SHORT', '[E3.2] head：configured=true、enabled=false、KEY_TOO_SHORT');
  ok(jsonlLines().length > 0 && jsonlLines().every(l => !('cseq' in l)), '[E3.3] 事件照常寫入、未上鏈（沒有遺失）');
  await stopServer();

  // ── E4：用 JWT_SECRET 備援金鑰啟用 ──
  cleanupFiles(); seedFiles({});
  jar = {}; bodies.length = 0;
  await startServer({ AUDIT_EVENTS_TABLE: '1', AUDIT_CHAIN: '1' });
  console.log('[E4] 端對端：AUDIT_CHAIN=1、沒有專用金鑰 → 備援 JWT_SECRET');
  await login('_s_admin'); await sleep(400);
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/head');
  ok(r.json.enabled === true && r.json.secretSource === 'JWT_SECRET', '[E4.1] 以 JWT_SECRET 備援啟用（head 顯示金鑰來源名稱）');
  r = await http('_s_admin', 'GET', '/api/admin/audit-chain/verify');
  ok(r.json.ok === true && r.json.checked > 0, '[E4.2] 驗證通過');
  r = await http(null, 'GET', '/api/cron/audit-anchor');
  ok(r.status === 401, '[E4.3] 沒設 CRON_SECRET → cron 一律 401');
  await stopServer();
}

// ═════════════ 主流程 ═════════════
(async () => {
  const only = (process.env.AC_ONLY || 'U,E').split(',');
  let phase = '';
  try {
    if (only.includes('U')) {
      await secU1(); await secU2(); await secU3(); await secU4(); await secU5(); await secU6(); secU7(); await secU9();
      if (!process.env.AC_SKIP_MUT) { const oe = console.error; console.error = () => {}; try { await secU8(); } finally { console.error = oe; } }
    }
    if (only.includes('E')) await secE();
  } catch (e) {
    fail++; failures.push('例外：' + e.message); console.log('例外', e && e.stack);
  } finally {
    await stopServer();
    cleanupFiles();
    cleanupTmp();
    const left = FILES.filter(f => fs.existsSync(f));
    if (left.length) { fail++; console.log('  ✗ 殘留檔案：', left.map(f => path.basename(f)).join(',')); }
  }
  console.log(`\n稽核防竄改檢查：PASS ${pass} / FAIL ${fail}`);
  if (fail) { console.log('失敗項目：'); failures.forEach(f => console.log(' -', f)); }
  if (UNVERIFIED.length) console.log('未驗證項目（' + UNVERIFIED.length + '，結束碼 3）：' + UNVERIFIED.join('；'));
  process.exit(fail ? 1 : UNVERIFIED.length ? 3 : 0);
})();
