#!/usr/bin/env node
'use strict';
/**
 * outbox 檢查。用法：node scripts/check-mail-outbox.js（不開伺服器、不連網、不碰 data.json／auth.json；檔案測試只寫系統暫存資料夾）
 *
 * 涵蓋：
 *   1) 同一組「行為測試」依序跑在三種 adapter 上：記憶體、JSON 檔案、Postgres（用本檔的記憶體模擬器代替真資料庫）
 *      入列冪等與欄位白名單、領取（含 50 個並行領取不重複）、租約過期、退避 60/300/900、Retry-After、重試用盡、
 *      requeue 再一輪、樂觀鎖（租約被搶走後舊工作者回報無效）、skipped／cancelled、list／purge／stats
 *   2) 熔斷（breaker）
 *   3) JSON 檔案 adapter：原子寫入、rename 失敗不弄壞原檔、損壞復原（備份＋從空開始、不 throw）、跨 adapter 實例的並行
 *   4) Postgres adapter 的靜態審查：SQL 全是常數、全部參數化、參數個數與 $n 相符、注入字串只出現在參數、領取語意（SKIP LOCKED）
 *   5) scrubMessage（密鑰樣式字串、email、GUID、ReDoS）
 *
 * 重要限制：Postgres 這一節跑的是「假 query 函式」＋記憶體模擬器。模擬器是照 SQL 文字的語意寫的，能驗證參數順序、參數型別、
 * 各個狀態轉換的行為與兩種 adapter 的一致性，但「SQL 語法／型別推斷／鎖行為在真 Postgres 上是否如預期」沒有被驗證。
 *
 * 環境變數 MAIL_CORE_ROOT：要檢查的專案根目錄（變異測試用）。撰寫備註：本檔不寫四位數的 \uXXXX，一律用 cp()／\x..。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const util = require('util');
const assert = require('assert');

const ROOT = process.env.MAIL_CORE_ROOT || path.join(__dirname, '..');
const load = (rel) => require(path.join(ROOT, rel));

// ── 迷你測試框架（章節可為 async）────────────────────────────────────────────
const results = [];
const perf = [];
const sections = [];
let currentSection = '';
function record(name, ok, extra) { results.push({ section: currentSection, name, ok: !!ok, extra: extra === undefined ? '' : String(extra) }); }
const t = (name, ok, extra) => record(name, ok, extra);
function short(v) {
  let s;
  try { s = typeof v === 'string' ? JSON.stringify(v) : util.inspect(v, { depth: 4, breakLength: Infinity }); } catch (e) { s = String(v); }
  return s.length > 200 ? s.slice(0, 200) + '…' : s;
}
function eq(name, actual, expected) {
  let ok = true;
  try { assert.deepStrictEqual(actual, expected); } catch (e) { ok = false; }
  record(name, ok, ok ? '' : 'actual=' + short(actual) + ' expected=' + short(expected));
}
async function rejectsCode(name, p, code) {
  try { await p; record(name, false, 'did not reject'); } catch (e) { record(name, !!e && e.code === code, 'code=' + (e && e.code) + ' msg=' + (e && e.message)); }
}
function minMs(fn, runs) {
  let best = Infinity;
  for (let i = 0; i < (runs || 3); i++) { const s = process.hrtime.bigint(); fn(); best = Math.min(best, Number(process.hrtime.bigint() - s) / 1e6); }
  return best;
}
function fast(name, fn, limitMs) {
  const ms = minMs(fn);
  perf.push(name + ' ' + ms.toFixed(1) + 'ms');
  record('效能 ' + name + ' < ' + (limitMs || 50) + 'ms', ms < (limitMs || 50), ms.toFixed(1) + 'ms');
}
function section(name, fn) { sections.push({ name, fn }); }
function finish() {
  const failed = results.filter((r) => !r.ok);
  const bySec = {};
  results.forEach((r) => { const s = bySec[r.section] || (bySec[r.section] = { p: 0, f: 0 }); if (r.ok) s.p++; else s.f++; });
  Object.keys(bySec).forEach((k) => console.log((bySec[k].f ? 'FAIL ' : 'ok   ') + k + '  通過 ' + bySec[k].p + (bySec[k].f ? '，失敗 ' + bySec[k].f : '')));
  failed.forEach((r) => console.log('  ✗ [' + r.section + '] ' + r.name + (r.extra ? '  → ' + r.extra : '')));
  if (perf.length) console.log('效能：' + perf.join('；'));
  console.log((failed.length ? 'FAILED' : 'PASSED') + '：' + (results.length - failed.length) + ' / ' + results.length);
  process.exit(failed.length ? 1 : 0);
}
async function main() {
  const watchdog = setTimeout(() => { console.log('WATCHDOG：測試超過 120 秒未結束'); process.exit(2); }, 120000);
  for (const s of sections) {
    currentSection = s.name;
    try { await s.fn(); } catch (e) { record('章節執行中例外', false, (e && e.stack) || e); }
  }
  clearTimeout(watchdog);
  finish();
}

// ── 載入被測模組 ────────────────────────────────────────────────────────────
const OB = load('lib/mail/outbox.js');
const { createOutbox, jsonFileAdapter, memoryAdapter, postgresAdapter, scrubMessage, MailOutboxError } = OB;
const { getMailConfig } = load('lib/mail/config.js');
const cp = (...n) => String.fromCodePoint(...n);
const T0 = Date.UTC(2026, 9, 8, 3, 0, 0);
const iso = (ms) => new Date(ms).toISOString();
const SEC = 1000;
const DAY = 86400 * 1000;

const tmpDirs = [];
function tmpDir() {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'mail-outbox-test-'));
  tmpDirs.push(d);
  return d;
}
process.on('exit', () => { tmpDirs.forEach((d) => { try { fs.rmSync(d, { recursive: true, force: true }); } catch (e) { /* ignore */ } }); });

function mkJob(i, over) {
  return Object.assign({
    type: 'E1_SUBMIT', quoteId: 'q' + (i % 3), quoteNo: 'QU-' + i, toUser: 'user' + i, toMasked: 'u***@itts.com.tw',
    dedupeKey: 'E1_SUBMIT:q' + (i % 3) + ':user' + i + ':k' + i, actorLabel: 'Actor', meta: { level: 1, kind: 'mgr1' },
  }, over || {});
}

// ═════════════════════════════════════════════════════════════════════════
// Postgres 模擬器：照 SQL 文字的語意寫，並檢查參數個數與型別（從 SQL 文字推導，不是寫死）
// ═════════════════════════════════════════════════════════════════════════
function makeFakePg(opts) {
  const o = opts || {};
  const SQL = postgresAdapter.SQL;
  const nameOf = new Map(Object.keys(SQL).map((k) => [SQL[k], k]));
  const ddlTexts = new Set(Object.keys(postgresAdapter.DDL).map((k) => postgresAdapter.DDL[k]));
  const rows = [];
  let breakerRow = null;
  const calls = [];
  const ddlRuns = [];
  const STATUS = ['pending', 'sending', 'sent', 'failed', 'skipped', 'cancelled'];
  const D = (v) => (v === null || v === undefined ? null : new Date(v));
  const ms = (d) => (d === null ? NaN : d.getTime());
  const copy = (r) => Object.assign({}, r, { meta: JSON.parse(JSON.stringify(r.meta)) });

  function checkSql(sql, params) {
    const nums = Array.from(sql.matchAll(/\$(\d+)/g)).map((m) => Number(m[1]));
    const max = nums.length ? Math.max.apply(null, nums) : 0;
    if (params.length !== max) throw new Error('參數個數 ' + params.length + ' 與 SQL 最大 $n=' + max + ' 不符');
    for (let i = 1; i <= max; i++) if (nums.indexOf(i) < 0) throw new Error('$' + i + ' 沒被使用');
    for (const m of sql.matchAll(/\$(\d+)::timestamptz/g)) {
      const v = params[Number(m[1]) - 1];
      if (v !== null && !(typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v) && !isNaN(Date.parse(v)))) throw new Error('$' + m[1] + ' 不是合法的 timestamptz 字串：' + short(v));
    }
    for (const m of sql.matchAll(/\$(\d+)::jsonb/g)) {
      const v = params[Number(m[1]) - 1];
      if (typeof v !== 'string') throw new Error('$' + m[1] + ' jsonb 參數必須是 JSON 字串');
      JSON.parse(v);
    }
    for (const m of sql.matchAll(/\$(\d+)::text\[\]/g)) {
      const v = params[Number(m[1]) - 1];
      if (!Array.isArray(v) || !v.every((x) => typeof x === 'string')) throw new Error('$' + m[1] + ' 必須是字串陣列');
    }
    for (const m of sql.matchAll(/(?:LIMIT|OFFSET) \$(\d+)/g)) {
      if (!Number.isInteger(params[Number(m[1]) - 1])) throw new Error('LIMIT/OFFSET 參數必須是整數');
    }
    for (const m of sql.matchAll(/(?:attempts|attempts_in_round)\s*(?:=|<|>=)\s*\$(\d+)/g)) {
      if (!Number.isInteger(params[Number(m[1]) - 1])) throw new Error('$' + m[1] + ' 必須是整數');
    }
  }
  function checkRow(r) {
    if (STATUS.indexOf(r.status) < 0) throw new Error('CHECK mail_outbox_status_chk 違反');
    if (r.status === 'pending' && r.next_attempt_at === null) throw new Error('CHECK mail_outbox_pending_chk 違反');
    if (r.status === 'sending' && r.lease_until === null) throw new Error('CHECK mail_outbox_sending_chk 違反');
  }
  function touch(r, patch) { Object.assign(r, patch); checkRow(r); return r; }
  const find = (id) => rows.find((r) => r.id === id);

  async function query(sql, params) {
    // 這個函式本體是同步的（沒有 await），所以每次呼叫對資料來說是原子的——模擬資料庫的單一語句原子性
    calls.push({ sql, params: params ? params.slice() : params });
    if (o.failNext && o.failNext > 0) { o.failNext -= 1; throw new Error(o.failMessage || 'simulated failure'); }
    if (ddlTexts.has(sql)) { ddlRuns.push(sql); if (o.ddlHook) o.ddlHook(sql, ddlRuns.length); return { rows: [] }; }
    const name = nameOf.get(sql);
    if (!name) throw new Error('未知的 SQL（不是 outboxAdapters 的常數）');
    checkSql(sql, params);
    const p = params;
    switch (name) {
      case 'insert': {
        if (rows.some((r) => r.dedupe_key === p[6])) return { rows: [] };
        const r = {
          id: p[0], type: p[1], quote_id: p[2], quote_no: p[3], to_user: p[4], to_masked: p[5], dedupe_key: p[6], status: p[7],
          attempts: p[8], attempts_in_round: p[9], requeues: p[10], next_attempt_at: D(p[11]), lease_until: D(p[12]),
          last_error_code: p[13], last_error_msg: p[14], skip_reason: p[15], actor_label: p[16], meta: JSON.parse(p[17]),
          created_at: D(p[18]), updated_at: D(p[19]), sent_at: D(p[20]),
        };
        checkRow(r);
        rows.push(r);
        return { rows: [copy(r)] };
      }
      case 'getById': return { rows: rows.filter((r) => r.id === p[0]).map(copy) };
      case 'getByKey': return { rows: rows.filter((r) => r.dedupe_key === p[0]).map(copy) };
      case 'claimById': {
        const r = find(p[0]);
        if (!r || r.status !== 'pending' || !(ms(r.next_attempt_at) <= Date.parse(p[1]))) return { rows: [] };
        touch(r, { status: 'sending', lease_until: D(p[2]), attempts: r.attempts + 1, attempts_in_round: r.attempts_in_round + 1, updated_at: D(p[1]) });
        return { rows: [copy(r)] };
      }
      case 'expireExhausted': {
        const flipped = [];                                  // SQL 有 RETURNING *：回傳被改成 failed 的列
        rows.forEach((r) => {
          if (r.status === 'sending' && ms(r.lease_until) <= Date.parse(p[0]) && r.attempts_in_round >= p[2]) {
            touch(r, { status: 'failed', last_error_code: 'LEASE_EXPIRED', last_error_msg: p[1], lease_until: null, updated_at: D(p[0]) });
            flipped.push(copy(r));
          }
        });
        return { rows: flipped };
      }
      case 'claimDue': {
        const now = Date.parse(p[0]);
        const due = rows.filter((r) => (r.status === 'pending' && ms(r.next_attempt_at) <= now) || (r.status === 'sending' && ms(r.lease_until) <= now && r.attempts_in_round < p[3]));
        due.sort((a, b) => (ms(a.next_attempt_at) - ms(b.next_attempt_at)) || (ms(a.created_at) - ms(b.created_at)));
        const picked = due.slice(0, p[2]);
        picked.forEach((r) => touch(r, { status: 'sending', lease_until: D(p[1]), attempts: r.attempts + 1, attempts_in_round: r.attempts_in_round + 1, updated_at: D(p[0]) }));
        return { rows: picked.map(copy) };
      }
      case 'markSent': {
        const r = find(p[0]);
        if (!r || ['pending', 'sending', 'failed', 'cancelled'].indexOf(r.status) < 0) return { rows: [] };
        touch(r, { status: 'sent', sent_at: D(p[1]), lease_until: null, next_attempt_at: null, updated_at: D(p[1]) });
        return { rows: [copy(r)] };
      }
      case 'finishAttempt': {
        const r = find(p[0]);
        if (!r || r.status !== 'sending' || r.attempts !== p[1]) return { rows: [] };
        touch(r, { status: p[3], next_attempt_at: D(p[4]), lease_until: null, last_error_code: p[5], last_error_msg: p[6], updated_at: D(p[2]) });
        return { rows: [copy(r)] };
      }
      case 'setTerminal': {
        const r = find(p[0]);
        if (!r || p[4].indexOf(r.status) < 0) return { rows: [] };
        touch(r, { status: p[2], skip_reason: p[3], lease_until: null, next_attempt_at: null, updated_at: D(p[1]) });
        return { rows: [copy(r)] };
      }
      case 'requeue': {
        const r = find(p[0]);
        if (!r || ['failed', 'cancelled', 'skipped'].indexOf(r.status) < 0) return { rows: [] };
        touch(r, { status: 'pending', next_attempt_at: D(p[1]), lease_until: null, attempts_in_round: 0, requeues: r.requeues + 1, skip_reason: null, updated_at: D(p[1]) });
        return { rows: [copy(r)] };
      }
      case 'list':
      case 'listCount': {
        let out = rows.filter((r) => (p[0] === null || r.status === p[0]) && (p[1] === null || r.type === p[1]) && (p[2] === null || r.quote_no === p[2])
          && (p[3] === null || r.to_user === p[3]) && (p[4] === null || ms(r.created_at) >= Date.parse(p[4])));
        if (name === 'listCount') return { rows: [{ n: out.length }] };
        out = out.slice().sort((a, b) => (ms(b.created_at) - ms(a.created_at)) || (a.id < b.id ? 1 : (a.id > b.id ? -1 : 0)));
        return { rows: out.slice(p[6], p[6] + p[5]).map(copy) };
      }
      case 'purge': {
        const cut = Date.parse(p[0]);
        let n = 0;
        for (let i = rows.length - 1; i >= 0; i--) {
          if (['sent', 'failed', 'skipped', 'cancelled'].indexOf(rows[i].status) >= 0 && ms(rows[i].updated_at) < cut) { rows.splice(i, 1); n += 1; }
        }
        return { rows: [{ n }] };
      }
      case 'statsByStatus': {
        const g = {};
        rows.forEach((r) => { const x = g[r.status] || (g[r.status] = { status: r.status, n: 0, oldest: r.created_at }); x.n += 1; if (ms(r.created_at) < ms(x.oldest)) x.oldest = r.created_at; });
        return { rows: Object.keys(g).map((k) => g[k]) };
      }
      case 'statsFailedSince': return { rows: [{ n: rows.filter((r) => r.status === 'failed' && ms(r.updated_at) >= Date.parse(p[0])).length }] };
      case 'breakerGet': return { rows: breakerRow ? [{ state: JSON.parse(breakerRow.state) }] : [] };
      case 'breakerSet': breakerRow = { state: p[0] }; return { rows: [] };
      default: throw new Error('模擬器未實作 ' + name);
    }
  }
  return { query, rows, calls, ddlRuns };
}

// ── 三種 adapter 的工廠 ─────────────────────────────────────────────────────
const FACTORIES = {
  memory() { const adapter = memoryAdapter(); return { adapter, dump: () => adapter.dump() }; },
  json() {
    const file = path.join(tmpDir(), 'mail-outbox.json');
    const adapter = jsonFileAdapter({ file });
    return { adapter, file, dump: () => (fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '') };
  },
  pg() {
    const fake = makeFakePg();
    const adapter = postgresAdapter({ query: fake.query });
    return { adapter, fake, dump: () => JSON.stringify(fake.calls.map((c) => c.params)) };
  },
};

function mkEnv(kind, configOver) {
  const clock = { t: T0, now() { return this.t; }, advance(ms) { this.t += ms; } };
  const f = FACTORIES[kind]();
  const config = Object.assign(getMailConfig({}), configOver || {});
  const outbox = createOutbox(f.adapter, { now: () => clock.t, config });
  return Object.assign({ outbox, clock, config }, f);
}

// ═════════════════════════════════════════════════════════════════════════
// 1) 共用行為測試
// ═════════════════════════════════════════════════════════════════════════
async function behaviorSuite(kind) {
  const L = '[' + kind + '] ';

  // ── 入列：白名單與冪等 ──
  {
    const e = mkEnv(kind);
    const leaky = Object.assign(mkJob(1), {
      html: '<p>SECRET_HTML_BODY_77</p>', text: 'SECRET_TEXT_BODY_77', subject: 'SECRET_SUBJECT_77', revenueCents: 987654321012,
      amount: 'SECRET_AMOUNT_77', marginText: '41.04%', email: 'full.address77@itts.com.tw', to: ['to.address77@itts.com.tw'],
      projectName: 'SECRET_PROJECT_77', company: 'SECRET_COMPANY_77', numbers: { revenueCents: 987654321012 }, cc: 'cc77@itts.com.tw',
    });
    const a = await e.outbox.enqueue(leaky);
    t(L + 'enqueue 建立新紀錄', a.created === true && typeof a.id === 'string' && a.id.length > 3, short(a));
    const rec = await e.outbox.get(a.id);
    eq(L + '紀錄只有白名單欄位', Object.keys(rec).sort(), ['actorLabel', 'attempts', 'attemptsInRound', 'createdAt', 'dedupeKey', 'id', 'lastErrorCode', 'lastErrorMsg', 'leaseUntil', 'meta', 'nextAttemptAt', 'quoteId', 'quoteNo', 'requeues', 'sentAt', 'skipReason', 'status', 'toMasked', 'toUser', 'type', 'updatedAt']);
    eq(L + '新紀錄初始狀態', [rec.status, rec.attempts, rec.attemptsInRound, rec.requeues, rec.nextAttemptAt, rec.leaseUntil, rec.sentAt], ['pending', 0, 0, 0, iso(T0), null, null]);
    const dump = e.dump();
    ['SECRET_HTML_BODY_77', 'SECRET_TEXT_BODY_77', 'SECRET_SUBJECT_77', '987654321012', 'SECRET_AMOUNT_77', '41.04', 'full.address77', 'to.address77', 'SECRET_PROJECT_77', 'SECRET_COMPANY_77', 'cc77@'].forEach((needle) => {
      t(L + '儲存體內容不含「' + needle + '」', dump.indexOf(needle) < 0);
    });
    t(L + '儲存體內容不含任何完整 email', !/[A-Za-z0-9._+-]+@[A-Za-z0-9-]+\.[A-Za-z]{2,}/.test(dump), (dump.match(/[A-Za-z0-9._+-]+@[A-Za-z0-9-]+\.[A-Za-z]{2,}/) || [''])[0]);

    const b = await e.outbox.enqueue(mkJob(1));
    t(L + '同 dedupeKey 再入列：不新建、回既有紀錄', b.created === false && b.id === a.id && b.existing && b.existing.id === a.id, short(b));
    eq(L + '重複入列後仍只有一筆', (await e.outbox.list({})).total, 1);

    // toMasked 傳入完整位址 → 改成遮罩，不會存原值
    const c = await e.outbox.enqueue(mkJob(2, { toMasked: 'full.person@itts.com.tw' }));
    eq(L + '完整位址的 toMasked 被改成遮罩', (await e.outbox.get(c.id)).toMasked, 'f***@itts.com.tw');
    const c2 = await e.outbox.enqueue(mkJob(3, { toMasked: 'not an email at all' }));
    eq(L + '亂字串的 toMasked 被清成空字串', (await e.outbox.get(c2.id)).toMasked, '');

    // 狀態欄位不能由呼叫端偽造
    const forged = await e.outbox.enqueue(mkJob(4, { status: 'sent', attempts: 99, sentAt: iso(T0), leaseUntil: iso(T0 + 1), id: 'forged-id', requeues: 7, lastErrorCode: 'X' }));
    const fr = await e.outbox.get(forged.id);
    t(L + '呼叫端無法偽造 id／狀態／次數', fr.id !== 'forged-id' && fr.status === 'pending' && fr.attempts === 0 && fr.sentAt === null && fr.requeues === 0 && fr.lastErrorCode === null, short(fr));

    // 非法 job
    const bads = {
      'null': null, '非物件': 'x', '未知 type': mkJob(5, { type: 'E9_X' }), 'quoteId 含斜線': mkJob(5, { quoteId: '../x' }), 'quoteId 空': mkJob(5, { quoteId: '' }),
      'quoteNo 空': mkJob(5, { quoteNo: '' }), 'toUser 空': mkJob(5, { toUser: '' }), 'toUser 含換行': mkJob(5, { toUser: 'a\nb' }),
      'dedupeKey 空': mkJob(5, { dedupeKey: '' }), 'dedupeKey 含 NUL': mkJob(5, { dedupeKey: 'a\x00b' }), 'dedupeKey 過長': mkJob(5, { dedupeKey: 'k'.repeat(701) }),
    };
    for (const k of Object.keys(bads)) await rejectsCode(L + '非法 job 被拒：' + k, e.outbox.enqueue(bads[k]), 'BAD_JOB');
    eq(L + '非法 job 沒有被存進去', (await e.outbox.list({})).total, 4);

    // meta 只收白名單值
    const m = await e.outbox.enqueue(mkJob(6, { meta: { level: 'hacker', kind: 'superadmin', extra: 'x' } }));
    eq(L + 'meta 不認得的 level／kind 變 null', (await e.outbox.get(m.id)).meta, { level: null, kind: null });
    const m2 = await e.outbox.enqueue(mkJob(7, { meta: { level: 'board', kind: 'secretary' } }));
    eq(L + 'meta 合法值保留', (await e.outbox.get(m2.id)).meta, { level: 'board', kind: 'secretary' });

    // 50 個並行入列同一把鍵 → 只建立一次
    const e2 = mkEnv(kind);
    const rs = await Promise.all(Array.from({ length: 50 }, () => e2.outbox.enqueue(mkJob(9))));
    eq(L + '50 個並行入列同一把 dedupeKey：恰好 1 個 created', rs.filter((r) => r.created).length, 1);
    eq(L + '50 個並行入列後只有 1 筆', (await e2.outbox.list({})).total, 1);
    t(L + '50 個並行入列回傳的 id 全部相同', new Set(rs.map((r) => r.id)).size === 1);

    // skipped 直接入列
    const sk = await e2.outbox.enqueue(mkJob(10, { status: 'skipped', skipReason: 'MODE_OFF' }));
    const skr = await e2.outbox.get(sk.id);
    eq(L + '以 skipped 入列：狀態與原因', [skr.status, skr.skipReason, skr.nextAttemptAt], ['skipped', 'MODE_OFF', null]);
    const sk2 = await e2.outbox.enqueue(mkJob(11, { status: 'skipped', skipReason: 'bad reason!' }));
    eq(L + 'skipReason 格式不合法會被改成 UNSPECIFIED', (await e2.outbox.get(sk2.id)).skipReason, 'UNSPECIFIED');
    eq(L + 'skipped 紀錄不會被領取', (await e2.outbox.claimDue({ limit: 50 })).filter((j) => j.id === sk.id).length, 0);
  }

  // ── claim／claimDue ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    const c1 = await e.outbox.claim(a.id, { leaseSec: 60 });
    eq(L + 'claim：pending → sending，attempts=1，租約 60 秒', [c1.status, c1.attempts, c1.attemptsInRound, c1.leaseUntil], ['sending', 1, 1, iso(T0 + 60 * SEC)]);
    t(L + 'claim 同一筆第二次 → null', (await e.outbox.claim(a.id)) === null);
    const future = await e.outbox.enqueue(mkJob(2, { nextAttemptAt: iso(T0 + 10 * SEC) }));
    t(L + '未到期的 claim → null', (await e.outbox.claim(future.id)) === null);
    t(L + 'claim 不存在的 id → null', (await e.outbox.claim('nope')) === null);
    eq(L + 'claimDue 不領未到期的', (await e.outbox.claimDue({ limit: 5 })).length, 0);
    e.clock.advance(10 * SEC);
    const due = await e.outbox.claimDue({ limit: 5 });
    eq(L + 'claimDue 到期後領到那一筆', due.map((j) => j.id), [future.id]);

    const e2 = mkEnv(kind);
    for (let i = 0; i < 7; i++) { await e2.outbox.enqueue(mkJob(i, { nextAttemptAt: iso(T0 - (10 - i) * SEC) })); }
    const got = await e2.outbox.claimDue({ limit: 3 });
    eq(L + 'claimDue 依到期時間由舊到新、遵守 limit', got.map((j) => j.toUser), ['user0', 'user1', 'user2']);
    eq(L + 'claimDue limit 超界會被夾住（0→1）', (await e2.outbox.claimDue({ limit: 0 })).length, 1);
    eq(L + 'claimDue limit 缺省＝5', (await e2.outbox.claimDue({})).length, 3);   // 剩 3 筆
  }

  // ── 並行領取：每一筆只會被領到一次 ──
  {
    const e = mkEnv(kind);
    for (let i = 0; i < 20; i++) await e.outbox.enqueue(mkJob(i));
    const batches = await Promise.all(Array.from({ length: 50 }, () => e.outbox.claimDue({ limit: 1 })));
    const ids = batches.reduce((a, b) => a.concat(b.map((j) => j.id)), []);
    eq(L + '50 個並行 claimDue(limit 1) 搶 20 筆：總共領到 20 筆', ids.length, 20);
    eq(L + '50 個並行 claimDue：沒有任何一筆被領到兩次', new Set(ids).size, 20);
    eq(L + '50 個並行 claimDue 後全部都是 sending 且 attempts=1', (await e.outbox.list({ status: 'sending', limit: 100 })).rows.every((r) => r.attempts === 1), true);

    const e2 = mkEnv(kind);
    for (let i = 0; i < 30; i++) await e2.outbox.enqueue(mkJob(i));
    const b2 = await Promise.all(Array.from({ length: 50 }, () => e2.outbox.claimDue({ limit: 5 })));
    const ids2 = b2.reduce((a, b) => a.concat(b.map((j) => j.id)), []);
    eq(L + '50 個並行 claimDue(limit 5) 搶 30 筆：領到 30 筆且不重複', [ids2.length, new Set(ids2).size], [30, 30]);

    const e3 = mkEnv(kind);
    const one = await e3.outbox.enqueue(mkJob(1));
    const w = await Promise.all(Array.from({ length: 50 }, () => e3.outbox.claim(one.id)));
    eq(L + '50 個並行 claim 同一筆：恰好 1 個贏家', w.filter(Boolean).length, 1);
    eq(L + '贏家之後 attempts=1', (await e3.outbox.get(one.id)).attempts, 1);
  }

  // ── 租約 ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    await e.outbox.claim(a.id, { leaseSec: 60 });
    e.clock.advance(59 * SEC);
    eq(L + '租約未到期：sending 不會被重新領取', (await e.outbox.claimDue({ limit: 5 })).length, 0);
    e.clock.advance(1 * SEC);
    const re = await e.outbox.claimDue({ limit: 5 });
    eq(L + '租約到期：被重新領取（attempts=2）', re.map((j) => [j.id, j.attempts, j.attemptsInRound]), [[a.id, 2, 2]]);
    // 次數用盡的過期租約 → failed/LEASE_EXPIRED，不再領取
    for (let i = 0; i < 2; i++) { e.clock.advance(61 * SEC); await e.outbox.claimDue({ limit: 5 }); }   // 第 3、4 次
    eq(L + '已領取 4 次（本輪上限）', (await e.outbox.get(a.id)).attempts, 4);
    e.clock.advance(61 * SEC);
    const none = await e.outbox.claimDue({ limit: 5 });
    const fin = await e.outbox.get(a.id);
    eq(L + '第 5 次不再領取', none.length, 0);
    eq(L + '次數用盡且租約過期：標 failed／LEASE_EXPIRED', [fin.status, fin.lastErrorCode], ['failed', 'LEASE_EXPIRED']);
  }

  // ── 租約過期且次數用盡：狀態轉換要「看得見」（expireStale 回傳被改掉的紀錄；以前是 claimDue 裡的靜默轉換）──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    await e.outbox.claim(a.id, { leaseSec: 60 });
    for (let i = 0; i < 3; i++) { e.clock.advance(61 * SEC); await e.outbox.claimDue({ limit: 5, sweep: false }); }   // a 被領到第 4 次（本輪上限）
    eq(L + '前置：a 已領取 4 次、仍是 sending', [(await e.outbox.get(a.id)).attempts, (await e.outbox.get(a.id)).status], [4, 'sending']);
    e.clock.advance(1 * SEC);
    const b = await e.outbox.enqueue(mkJob(2));            // b 之後才入列（排序不受隨機 id 影響），只領 1 次，租約跟著過期
    await e.outbox.claim(b.id, { leaseSec: 60 });
    e.clock.advance(61 * SEC);
    const noSweep = await e.outbox.claimDue({ limit: 5, sweep: false });
    eq(L + 'sweep:false：不清掃（a 仍是 sending），只領得到 b（b 次數未用盡）', [noSweep.map((j) => j.id), (await e.outbox.get(a.id)).status], [[b.id], 'sending']);
    const swept = await e.outbox.expireStale();
    eq(L + 'expireStale：回傳被改掉的紀錄（a），status=failed、LEASE_EXPIRED', swept.map((j) => [j.id, j.status, j.lastErrorCode, j.toUser, j.quoteNo]), [[a.id, 'failed', 'LEASE_EXPIRED', 'user1', 'QU-1']]);
    t(L + 'expireStale 回傳的紀錄不含信件內容、金額、完整 email', !/[A-Za-z0-9._+-]+@[A-Za-z0-9-]+\.[A-Za-z]{2,}/.test(JSON.stringify(swept)) && !('html' in swept[0]) && !('text' in swept[0]), short(swept));
    eq(L + 'expireStale 第二次呼叫：沒有東西可清掃（回傳空陣列）', await e.outbox.expireStale(), []);
    eq(L + 'expireStale：b（次數未用盡）不受影響', (await e.outbox.get(b.id)).status, 'sending');
    eq(L + 'stats().failed24h 看得到這筆 LEASE_EXPIRED（後台統計也查得到）', (await e.outbox.stats()).failed24h, 1);
    // 預設 claimDue（不帶 sweep:false）仍會順手清掃（維持既有行為，結果丟棄）
    const e2 = mkEnv(kind);
    const c = await e2.outbox.enqueue(mkJob(3));
    await e2.outbox.claim(c.id, { leaseSec: 60 });
    for (let i = 0; i < 3; i++) { e2.clock.advance(61 * SEC); await e2.outbox.claimDue({ limit: 1 }); }
    e2.clock.advance(61 * SEC);
    await e2.outbox.claimDue({ limit: 5 });
    eq(L + '預設 claimDue 仍會清掃（既有行為不變）', (await e2.outbox.get(c.id)).status, 'failed');
  }

  // ── actorLabel 清理：控制字元、方向控制字元、零寬字元不入庫，長度上限 100 ──
  {
    const e = mkEnv(kind);
    const nasty = 'Act' + cp(0) + 'or\n' + cp(0x1b) + '[31m' + cp(0x202e) + 'evil' + cp(0x2066) + cp(0x200b) + cp(0xfeff) + '<b>x</b>';
    const a = await e.outbox.enqueue(mkJob(1, { actorLabel: nasty }));
    const lab = (await e.outbox.get(a.id)).actorLabel;
    t(L + 'actorLabel：控制字元（NUL／換行／ESC）不入庫', !/[\x00-\x1f\x7f-\x9f]/.test(lab), short(lab));
    t(L + 'actorLabel：方向控制／零寬／BOM 字元不入庫', !/[\u{202a}-\u{202e}\u{2066}-\u{2069}\u{200b}\u{200e}\u{200f}\u{2060}\u{feff}]/u.test(lab), short(lab));
    t(L + 'actorLabel：可見文字保留（只拿掉不可見字元）', lab.indexOf('Act') === 0 && lab.indexOf('evil') > 0, short(lab));
    const long = await e.outbox.enqueue(mkJob(2, { actorLabel: 'Z'.repeat(500) }));
    const longLab = (await e.outbox.get(long.id)).actorLabel;
    t(L + 'actorLabel：超長字串被截斷（最多 100 字加省略號）', longLab.length <= 101 && longLab.length < 500 && longLab.indexOf('Z'.repeat(100)) === 0, longLab.length);
    const dump = e.dump();
    const BS = String.fromCharCode(92);
    t(L + 'actorLabel：儲存體內容不含 NUL／ESC／RLO（原字元或 JSON 跳脫形式）', dump.indexOf(BS + 'u0000') < 0 && dump.indexOf(BS + 'u001b') < 0 && dump.indexOf(cp(0x202e)) < 0 && dump.indexOf(cp(0)) < 0 && dump.indexOf(cp(0x1b)) < 0);
    const odd = await e.outbox.enqueue(mkJob(3, { actorLabel: { toString: () => 'x' } }));
    eq(L + 'actorLabel：非字串輸入變空字串', (await e.outbox.get(odd.id)).actorLabel, '');
  }

  // ── markSent ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    await e.outbox.claim(a.id);
    e.clock.advance(5 * SEC);
    const r = await e.outbox.markSent(a.id);
    eq(L + 'markSent：sending → sent，sentAt 與租約', [r.ok, r.record.status, r.record.sentAt, r.record.leaseUntil, r.record.nextAttemptAt], [true, 'sent', iso(T0 + 5 * SEC), null, null]);
    const r2 = await e.outbox.markSent(a.id);
    t(L + 'markSent 重複呼叫是冪等的', r2.ok === true && r2.already === true && r2.record.sentAt === iso(T0 + 5 * SEC), short(r2));
    eq(L + 'markSent 不存在的 id', (await e.outbox.markSent('nope')).reason, 'NOT_FOUND');
    const s = await e.outbox.enqueue(mkJob(2, { status: 'skipped', skipReason: 'NO_EMAIL' }));
    eq(L + 'skipped 不能 markSent', (await e.outbox.markSent(s.id)).reason, 'BAD_STATE');
    eq(L + 'sent 之後不會再被領取', (await e.outbox.claimDue({ limit: 5 })).length, 0);
  }

  // ── markFailed：退避 60/300/900、用盡 ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    const expectDelay = [60, 300, 900];
    for (let n = 1; n <= 3; n++) {
      const job = n === 1 ? await e.outbox.claim(a.id) : (await e.outbox.claimDue({ limit: 5 }))[0];
      eq(L + '第 ' + n + ' 次嘗試的 attempts', job.attempts, n);
      const f = await e.outbox.markFailed(a.id, { code: 'TIMEOUT', msg: '逾時', attempts: job.attempts });
      eq(L + '第 ' + n + ' 次失敗：回 pending、排 ' + expectDelay[n - 1] + ' 秒後', [f.ok, f.final, f.record.status, f.record.nextAttemptAt, f.record.lastErrorCode], [true, false, 'pending', iso(e.clock.t + expectDelay[n - 1] * SEC), 'TIMEOUT']);
      e.clock.advance(expectDelay[n - 1] * SEC - 1);
      eq(L + '差 1 毫秒不會被領取', (await e.outbox.claimDue({ limit: 5 })).length, 0);
      e.clock.advance(1);
    }
    const job4 = (await e.outbox.claimDue({ limit: 5 }))[0];
    eq(L + '第 4 次嘗試（首次 + 3 次重試）', job4.attempts, 4);
    const f4 = await e.outbox.markFailed(a.id, { code: 'TIMEOUT', msg: '逾時', attempts: 4 });
    eq(L + '第 4 次失敗：重試用盡 → failed（final）', [f4.ok, f4.final, f4.record.status, f4.record.nextAttemptAt, f4.record.leaseUntil], [true, true, 'failed', null, null]);
    e.clock.advance(100 * DAY);
    eq(L + 'failed 不會再被領取', (await e.outbox.claimDue({ limit: 5 })).length, 0);
  }

  // ── markFailed：Retry-After、permanent、錯誤訊息清理、狀態檢查 ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    await e.outbox.claim(a.id);
    const f = await e.outbox.markFailed(a.id, { code: 'THROTTLED', msg: 'slow down', retryAfterSec: 7 });
    eq(L + 'Retry-After 7 秒優先於 60 秒排程', f.record.nextAttemptAt, iso(T0 + 7 * SEC));
    e.clock.advance(7 * SEC);
    await e.outbox.claimDue({ limit: 5 });
    const f2 = await e.outbox.markFailed(a.id, { code: 'THROTTLED', retryAfterSec: 10 * DAY / SEC });
    eq(L + 'Retry-After 過大被夾在 24 小時', f2.record.nextAttemptAt, iso(e.clock.t + DAY));
    e.clock.advance(DAY);
    await e.outbox.claimDue({ limit: 5 });
    const f3 = await e.outbox.markFailed(a.id, { code: 'SERVER', retryAfterSec: -5 });
    eq(L + 'Retry-After 為負數：忽略，改用排程（第 3 次＝900 秒）', f3.record.nextAttemptAt, iso(e.clock.t + 900 * SEC));
    e.clock.advance(900 * SEC);
    await e.outbox.claimDue({ limit: 5 });
    const f4 = await e.outbox.markFailed(a.id, { code: 'SERVER', retryAfterSec: NaN });
    t(L + 'Retry-After 為 NaN：照常用盡 → failed', f4.final === true && f4.record.status === 'failed', short(f4));

    const p = await e.outbox.enqueue(mkJob(2));
    await e.outbox.claim(p.id);
    const pf = await e.outbox.markFailed(p.id, { code: 'AUTH', msg: '401', permanent: true });
    eq(L + 'permanent：第一次就 failed', [pf.final, pf.record.status, pf.record.lastErrorCode, pf.record.attempts], [true, 'failed', 'AUTH', 1]);

    const q = await e.outbox.enqueue(mkJob(3));
    await e.outbox.claim(q.id);
    const secretMsg = 'Bearer ZZZtoken0123456789abcdefghij client_secret=FAKE~secret~0123456789abcdefghijklmnop tenant 00000000-0000-4000-8000-000000000001 to someone@itts.com.tw';
    const qf = await e.outbox.markFailed(q.id, { code: 'bad code!', msg: secretMsg });
    const qm = qf.record.lastErrorMsg;
    t(L + 'lastErrorMsg 不含密鑰樣式字串／GUID／email', !/ZZZtoken|FAKE~secret|00000000-0000-4000|someone@/.test(qm) && qm.length <= 200, qm);
    eq(L + '不合法的錯誤碼變 UNKNOWN', qf.record.lastErrorCode, 'UNKNOWN');
    t(L + '儲存體內容不含密鑰', e.dump().indexOf('FAKE~secret') < 0 && e.dump().indexOf('someone@') < 0);

    eq(L + 'markFailed 不在 sending 狀態', (await e.outbox.markFailed(a.id, { code: 'TIMEOUT' })).reason, 'NOT_SENDING');
    eq(L + 'markFailed 不存在的 id', (await e.outbox.markFailed('nope', { code: 'TIMEOUT' })).reason, 'NOT_FOUND');
  }

  // ── 樂觀鎖：租約被搶走後，舊工作者的回報無效 ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    const A = await e.outbox.claim(a.id, { leaseSec: 60 });
    e.clock.advance(61 * SEC);
    const B = (await e.outbox.claimDue({ limit: 5 }))[0];
    eq(L + '工作者 B 接手（attempts=2）', [B.id, B.attempts], [a.id, 2]);
    const stale = await e.outbox.markFailed(a.id, { code: 'TIMEOUT', attempts: A.attempts });
    eq(L + '舊工作者 A 的 markFailed 被拒（LOST_LEASE）', stale.reason, 'LOST_LEASE');
    eq(L + 'A 被拒後紀錄仍是 B 的 sending', [(await e.outbox.get(a.id)).status, (await e.outbox.get(a.id)).attempts], ['sending', 2]);
    const ok = await e.outbox.markFailed(a.id, { code: 'TIMEOUT', attempts: B.attempts });
    t(L + '新工作者 B 的 markFailed 成功', ok.ok === true && ok.record.status === 'pending', short(ok));
  }

  // ── markSkipped／cancel ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    const s = await e.outbox.markSkipped(a.id, 'NO_EMAIL');
    eq(L + 'markSkipped：pending → skipped（含原因）', [s.ok, s.record.status, s.record.skipReason, s.record.nextAttemptAt], [true, 'skipped', 'NO_EMAIL', null]);
    eq(L + 'skipped 再 markSkipped → BAD_STATE', (await e.outbox.markSkipped(a.id, 'NO_EMAIL')).reason, 'BAD_STATE');
    const b = await e.outbox.enqueue(mkJob(2));
    await e.outbox.claim(b.id);
    eq(L + 'sending 可 markSkipped', (await e.outbox.markSkipped(b.id, 'bad reason')).record.skipReason, 'UNSPECIFIED');

    const c = await e.outbox.enqueue(mkJob(3));
    const cc = await e.outbox.cancel(c.id, 'STALE');
    eq(L + 'cancel：pending → cancelled（原因存在 skipReason）', [cc.ok, cc.record.status, cc.record.skipReason, cc.record.nextAttemptAt, cc.record.leaseUntil], [true, 'cancelled', 'STALE', null, null]);
    eq(L + 'cancelled 不會被領取', (await e.outbox.claimDue({ limit: 50 })).filter((j) => j.id === c.id).length, 0);
    const d = await e.outbox.enqueue(mkJob(4));
    await e.outbox.claim(d.id);
    eq(L + 'sending 可 cancel', (await e.outbox.cancel(d.id, 'GONE')).record.status, 'cancelled');
    const f = await e.outbox.enqueue(mkJob(5));
    await e.outbox.claim(f.id);
    await e.outbox.markFailed(f.id, { code: 'AUTH', permanent: true });
    eq(L + 'failed 可 cancel', (await e.outbox.cancel(f.id, 'GONE')).record.status, 'cancelled');
    const g = await e.outbox.enqueue(mkJob(6));
    await e.outbox.claim(g.id);
    await e.outbox.markSent(g.id);
    eq(L + 'sent 不能 cancel', (await e.outbox.cancel(g.id, 'STALE')).reason, 'BAD_STATE');
    eq(L + 'cancel 不存在的 id', (await e.outbox.cancel('nope', 'STALE')).reason, 'NOT_FOUND');
  }

  // ── requeue：再給一輪 ──
  {
    const e = mkEnv(kind);
    const a = await e.outbox.enqueue(mkJob(1));
    await e.outbox.claim(a.id);
    await e.outbox.markFailed(a.id, { code: 'AUTH', permanent: true });
    eq(L + 'pending 不能 requeue', (await e.outbox.requeue((await e.outbox.enqueue(mkJob(2))).id)).reason, 'BAD_STATE');
    e.clock.advance(5 * SEC);
    const r = await e.outbox.requeue(a.id);
    eq(L + 'requeue：failed → pending，立即到期，attempts 保留、本輪歸零、requeues=1', [r.ok, r.record.status, r.record.nextAttemptAt, r.record.attempts, r.record.attemptsInRound, r.record.requeues], [true, 'pending', iso(T0 + 5 * SEC), 1, 0, 1]);
    // 之後再給完整一輪：4 次嘗試
    let finalAt = 0;
    for (let n = 1; n <= 4; n++) {
      const job = (await e.outbox.claimDue({ limit: 50 })).filter((j) => j.id === a.id)[0];
      if (!job) { finalAt = -n; break; }
      const f = await e.outbox.markFailed(a.id, { code: 'TIMEOUT', attempts: job.attempts });
      if (f.final) { finalAt = n; break; }
      e.clock.advance(1000 * SEC);
    }
    eq(L + 'requeue 後再給 4 次嘗試才 failed（第 4 次 final）', finalAt, 4);
    eq(L + 'requeue 後累計 attempts＝1 + 4', (await e.outbox.get(a.id)).attempts, 5);
    const sk = await e.outbox.enqueue(mkJob(3, { status: 'skipped', skipReason: 'MODE_OFF' }));
    const rs = await e.outbox.requeue(sk.id);
    eq(L + 'skipped 可 requeue（清掉 skipReason）', [rs.ok, rs.record.status, rs.record.skipReason], [true, 'pending', null]);
    const cn = await e.outbox.enqueue(mkJob(4));
    await e.outbox.cancel(cn.id, 'STALE');
    eq(L + 'cancelled 可 requeue', (await e.outbox.requeue(cn.id)).record.status, 'pending');
    eq(L + 'requeue 不存在的 id', (await e.outbox.requeue('nope')).reason, 'NOT_FOUND');
  }

  // ── list／get／purge／stats ──
  {
    const e = mkEnv(kind);
    const ids = [];
    for (let i = 0; i < 6; i++) {
      ids.push((await e.outbox.enqueue(mkJob(i, { type: i % 2 ? 'E3_NEXT_STEP' : 'E1_SUBMIT' }))).id);
      e.clock.advance(1000 * SEC);
    }
    await e.outbox.claim(ids[0]); await e.outbox.markSent(ids[0]);
    await e.outbox.claim(ids[1]); await e.outbox.markFailed(ids[1], { code: 'AUTH', permanent: true });
    await e.outbox.markSkipped(ids[2], 'NO_EMAIL');
    const all = await e.outbox.list({});
    eq(L + 'list：總數與預設排序（新→舊）', [all.total, all.rows.map((r) => r.toUser)], [6, ['user5', 'user4', 'user3', 'user2', 'user1', 'user0']]);
    eq(L + 'list 依 status 篩', (await e.outbox.list({ status: 'sent' })).rows.map((r) => r.toUser), ['user0']);
    eq(L + 'list 依 type 篩', (await e.outbox.list({ type: 'E3_NEXT_STEP' })).total, 3);
    eq(L + 'list 依 quoteNo 篩', (await e.outbox.list({ quoteNo: 'QU-3' })).rows.map((r) => r.toUser), ['user3']);
    eq(L + 'list 依 toUser 篩', (await e.outbox.list({ toUser: 'user4' })).total, 1);
    eq(L + 'list 依 since 篩（含）', (await e.outbox.list({ since: iso(T0 + 3000 * SEC) })).rows.map((r) => r.toUser), ['user5', 'user4', 'user3']);
    const pg = await e.outbox.list({ limit: 2, offset: 2 });
    eq(L + 'list 分頁', [pg.total, pg.rows.map((r) => r.toUser), pg.limit, pg.offset], [6, ['user3', 'user2'], 2, 2]);
    eq(L + 'list 不認得的 status 被忽略（不會變成 SQL 或空結果）', (await e.outbox.list({ status: "x' OR '1'='1" })).total, 6);
    eq(L + 'list limit 夾在 1..500', [(await e.outbox.list({ limit: 0 })).limit, (await e.outbox.list({ limit: 99999 })).limit], [1, 500]);
    const stt = await e.outbox.stats();
    eq(L + 'stats：各狀態筆數', stt.counts, { pending: 3, sending: 0, sent: 1, failed: 1, skipped: 1, cancelled: 0 });
    eq(L + 'stats：總數、最舊 pending、近 24 小時失敗數', [stt.total, stt.oldestPendingAt, stt.failed24h], [6, iso(T0 + 3000 * SEC), 1]);
    e.clock.advance(25 * 3600 * SEC);
    eq(L + 'stats：失敗滿 24 小時後不再計入 failed24h', (await e.outbox.stats()).failed24h, 0);

    // purge：只清「終態且夠舊」的
    const e2 = mkEnv(kind);
    const a = await e2.outbox.enqueue(mkJob(1)); await e2.outbox.claim(a.id); await e2.outbox.markSent(a.id);       // T0 sent
    const b = await e2.outbox.enqueue(mkJob(2));                                                                     // 一直 pending
    const c = await e2.outbox.enqueue(mkJob(3)); await e2.outbox.claim(c.id);                                        // 一直 sending
    const f = await e2.outbox.enqueue(mkJob(4)); await e2.outbox.markSkipped(f.id, 'NO_EMAIL');                      // T0 skipped
    e2.clock.advance(80 * DAY);
    const g = await e2.outbox.enqueue(mkJob(5)); await e2.outbox.claim(g.id); await e2.outbox.markSent(g.id);       // T0+80d sent
    e2.clock.advance(11 * DAY);                                                                                      // 現在 T0+91d
    eq(L + 'purge(90)：清掉 91 天前的終態紀錄（2 筆）', await e2.outbox.purge(90), 2);
    const left = (await e2.outbox.list({ limit: 100 })).rows.map((r) => r.toUser).sort();
    eq(L + 'purge 後剩下：pending、sending 與較新的 sent', left, ['user2', 'user3', 'user5']);
    eq(L + 'purge 再跑一次＝0', await e2.outbox.purge(90), 0);
    eq(L + 'purge 不合法的天數：用預設保留天數（90）', await e2.outbox.purge('abc'), 0);
    e2.clock.advance(100 * DAY);
    eq(L + 'purge 絕不清 pending／sending（即使很舊）', [await e2.outbox.purge(1), (await e2.outbox.list({ limit: 100 })).rows.map((r) => r.status).sort()], [1, ['pending', 'sending']]);
  }
}

['memory', 'json', 'pg'].forEach((kind) => section('1 outbox 行為 [' + kind + ']', () => behaviorSuite(kind)));

// ═════════════════════════════════════════════════════════════════════════
// 2) 熔斷
// ═════════════════════════════════════════════════════════════════════════
section('2 熔斷 breaker', async () => {
  function mk(over, kind) {
    const e = mkEnv(kind || 'memory', over);
    return e;
  }
  {
    const e = mk();
    eq('初始狀態：關閉', await e.outbox.breaker.state(), { open: false, until: null, failures: 0, halfOpen: false });
    for (let i = 0; i < 4; i++) await e.outbox.breaker.record(false, 'TIMEOUT');
    eq('4 次 TIMEOUT：仍關閉，failures=4', [(await e.outbox.breaker.state()).open, (await e.outbox.breaker.state()).failures], [false, 4]);
    const s = await e.outbox.breaker.record(false, 'TIMEOUT');
    eq('第 5 次 TIMEOUT：開啟 600 秒', [s.open, s.until], [true, iso(T0 + 600 * SEC)]);
    e.clock.advance(599 * SEC);
    eq('開啟後 599 秒仍開著', (await e.outbox.breaker.state()).open, true);
    e.clock.advance(1 * SEC);
    const half = await e.outbox.breaker.state();
    eq('600 秒後：關閉並進入半開', [half.open, half.halfOpen], [false, true]);
    const again = await e.outbox.breaker.record(false, 'NETWORK');
    eq('半開時一次失敗就立刻重新開啟', [again.open, again.until], [true, iso(e.clock.t + 600 * SEC)]);
    e.clock.advance(601 * SEC);
    const ok = await e.outbox.breaker.record(true);
    eq('半開時成功：完全重置', ok, { open: false, until: null, failures: 0, halfOpen: false });
  }
  {
    const e = mk();
    for (let i = 0; i < 20; i++) { await e.outbox.breaker.record(false, ['BAD_MESSAGE', 'REJECTED', 'NOT_CONFIGURED', 'RENDER', '', undefined, 123][i % 7]); }
    eq('與傳輸健康無關的錯誤碼（BAD_MESSAGE／REJECTED／NOT_CONFIGURED／未知）不計入熔斷', await e.outbox.breaker.state(), { open: false, until: null, failures: 0, halfOpen: false });
  }
  {
    const e = mk();
    for (let i = 0; i < 4; i++) await e.outbox.breaker.record(false, 'SERVER');
    await e.outbox.breaker.record(true);
    for (let i = 0; i < 4; i++) await e.outbox.breaker.record(false, 'SERVER');
    eq('任何成功都會重置計數（4+成功+4 不會開）', (await e.outbox.breaker.state()).open, false);
  }
  {
    const e = mk();
    const s = await e.outbox.breaker.record(false, 'AUTH');
    eq('AUTH 一次就開啟', [s.open, s.until], [true, iso(T0 + 600 * SEC)]);
    const e2 = mk();
    await e2.outbox.breaker.record(false, 'THROTTLED');
    eq('THROTTLED 權重 3：一次還沒開', [(await e2.outbox.breaker.state()).open, (await e2.outbox.breaker.state()).failures], [false, 3]);
    eq('THROTTLED + 2 次 TIMEOUT（3+1+1=5）：開啟', (await (async () => { await e2.outbox.breaker.record(false, 'TIMEOUT'); return e2.outbox.breaker.record(false, 'TIMEOUT'); })()).open, true);
    const e3 = mk();
    await e3.outbox.breaker.record(false, 'THROTTLED');
    eq('兩次 THROTTLED（3+3）：開啟', (await e3.outbox.breaker.record(false, 'THROTTLED')).open, true);
  }
  {
    const e = mk();
    for (let i = 0; i < 4; i++) await e.outbox.breaker.record(false, 'TIMEOUT');
    e.clock.advance(601 * SEC);
    const s = await e.outbox.breaker.record(false, 'TIMEOUT');
    eq('視窗（600 秒）外的舊失敗不累計', [s.open, s.failures], [false, 1]);
  }
  {
    const e = mk();
    await e.outbox.breaker.record(false, 'AUTH');
    const until0 = (await e.outbox.breaker.state()).until;
    e.clock.advance(100 * SEC);
    await e.outbox.breaker.record(false, 'TIMEOUT');
    await e.outbox.breaker.record(false, 'AUTH');
    eq('開啟期間的失敗不會延長開啟時間', (await e.outbox.breaker.state()).until, until0);
  }
  {
    const e = mk({ breaker: { failures: 2, windowSec: 30, openSec: 5 } });
    await e.outbox.breaker.record(false, 'TIMEOUT');
    const s = await e.outbox.breaker.record(false, 'TIMEOUT');
    eq('config.breaker 可調整（2 次、開 5 秒）', [s.open, s.until], [true, iso(T0 + 5 * SEC)]);
    const bad = mk({ breaker: { failures: -1, windowSec: 'x', openSec: 0 } });
    for (let i = 0; i < 5; i++) await bad.outbox.breaker.record(false, 'TIMEOUT');
    eq('config.breaker 不合法時用預設（5 次／600 秒）', (await bad.outbox.breaker.state()).until, iso(T0 + 600 * SEC));
  }
  // 持久化：兩個 outbox 實例共用同一個 adapter → 看到同一個熔斷狀態
  for (const kind of ['memory', 'json', 'pg']) {
    const e = mkEnv(kind);
    const other = createOutbox(e.adapter, { now: () => e.clock.t, config: e.config });
    await e.outbox.breaker.record(false, 'AUTH');
    eq('[' + kind + '] 熔斷狀態存在 adapter，另一個實例看得到', (await other.breaker.state()).open, true);
    await other.breaker.record(true);
    eq('[' + kind + '] 另一個實例成功後重置，這邊看得到', (await e.outbox.breaker.state()).open, false);
  }
  // 損壞的熔斷狀態不會讓 outbox 壞掉
  {
    const e = mkEnv('memory');
    await e.adapter.breakerSave({ failures: 'x', openUntil: 'y' });
    eq('熔斷狀態格式損壞：當成關閉', (await e.outbox.breaker.state()).open, false);
  }
});

// ═════════════════════════════════════════════════════════════════════════
// 3) JSON 檔案 adapter
// ═════════════════════════════════════════════════════════════════════════
section('3 JSON 檔案 adapter', async () => {
  const mkOb = (adapter, clock) => createOutbox(adapter, { now: () => clock.t, config: getMailConfig({}) });
  const clock = { t: T0 };

  // 基本：持久化、無殘留暫存檔
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob = mkOb(jsonFileAdapter({ file }), clock);
    const a = await ob.enqueue(mkJob(1));
    t('寫入後檔案存在且是合法 JSON', JSON.parse(fs.readFileSync(file, 'utf8')).jobs.length === 1);
    eq('目錄中只有主檔（沒有殘留 .tmp）', fs.readdirSync(dir), ['mail-outbox.json']);
    const ob2 = mkOb(jsonFileAdapter({ file }), clock);
    eq('新的 adapter 實例讀得到先前的資料（持久化）', (await ob2.get(a.id)).dedupeKey, mkJob(1).dedupeKey);
  }
  // 相對路徑＋rootDir；子目錄自動建立
  {
    const dir = tmpDir();
    const ob = mkOb(jsonFileAdapter({ file: 'sub/dir/box.json', rootDir: dir }), clock);
    await ob.enqueue(mkJob(1));
    t('相對路徑以 rootDir 為基準，子目錄自動建立', fs.existsSync(path.join(dir, 'sub', 'dir', 'box.json')));
    eq('adapter.file 是解析後的完整路徑', jsonFileAdapter({ file: 'x.json', rootDir: dir }).file, path.join(dir, 'x.json'));
    const dflt = jsonFileAdapter({ rootDir: dir });
    eq('沒指定 file：預設 mail-outbox.json', path.basename(dflt.file), 'mail-outbox.json');
  }
  // 暫存檔命名（落在 .gitignore 既有的 *.tmp 規則內）與原子替換
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const written = [];
    const spyFs = Object.assign({}, fs, { writeFileSync(p, ...rest) { written.push(p); return fs.writeFileSync(p, ...rest); } });
    const ob = mkOb(jsonFileAdapter({ file, fs: spyFs }), clock);
    await ob.enqueue(mkJob(1));
    t('寫入是先寫暫存檔（*.tmp）再 rename', written.length === 1 && /\.tmp$/.test(written[0]) && written[0] !== file && path.dirname(written[0]) === dir, short(written));
  }
  // rename 失敗：原檔不變、暫存檔清掉、錯誤以 STORE_IO 回報；修好之後恢復正常
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    let failRename = false;
    const flakyFs = Object.assign({}, fs, { renameSync(a, b) { if (failRename) { const e = new Error('boom'); e.code = 'EXDEV'; throw e; } return fs.renameSync(a, b); } });
    const ob = mkOb(jsonFileAdapter({ file, fs: flakyFs }), clock);
    await ob.enqueue(mkJob(1));
    const before = fs.readFileSync(file, 'utf8');
    failRename = true;
    await rejectsCode('rename 失敗：入列 reject（STORE_IO）', ob.enqueue(mkJob(2)), 'STORE_IO');
    eq('rename 失敗：主檔內容一個位元都沒變', fs.readFileSync(file, 'utf8'), before);
    eq('rename 失敗：暫存檔已清掉', fs.readdirSync(dir), ['mail-outbox.json']);
    failRename = false;
    const ok = await ob.enqueue(mkJob(2));
    t('修好之後可繼續使用，且先前失敗的那筆沒有殘留', ok.created === true && (await ob.list({})).total === 2);
  }
  // Windows 上短暫鎖檔：EPERM 重試後成功
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    let n = 0;
    const lockedFs = Object.assign({}, fs, { renameSync(a, b) { n += 1; if (n <= 2) { const e = new Error('locked'); e.code = 'EPERM'; throw e; } return fs.renameSync(a, b); } });
    const ob = mkOb(jsonFileAdapter({ file, fs: lockedFs }), clock);
    const r = await ob.enqueue(mkJob(1));
    t('rename 遇到 EPERM（防毒鎖檔）會重試', r.created === true && n === 3, 'rename 次數=' + n);
  }
  // 讀取失敗（非 ENOENT）：不當成空檔覆蓋
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob0 = mkOb(jsonFileAdapter({ file }), clock);
    await ob0.enqueue(mkJob(1));
    const before = fs.readFileSync(file, 'utf8');
    const denyFs = Object.assign({}, fs, { readFileSync(p, ...r) { if (p === file) { const e = new Error('denied'); e.code = 'EACCES'; throw e; } return fs.readFileSync(p, ...r); } });
    const ob = mkOb(jsonFileAdapter({ file, fs: denyFs }), clock);
    await rejectsCode('讀取遇到 EACCES：reject STORE_IO（不是當成空檔）', ob.enqueue(mkJob(2)), 'STORE_IO');
    eq('讀取失敗時原檔未被覆蓋', fs.readFileSync(file, 'utf8'), before);
  }
  // 損壞復原
  const corrupt = {
    '亂碼': 'this is not json {{{',
    '被截斷的 JSON': '{"version":1,"jobs":[{"id":"a","dedupeKey":"k","status":"pen',
    '根是陣列': '[]',
    'jobs 不是陣列': '{"version":1,"jobs":"oops"}',
    '根是字串': '"hello"',
    '根是 null': 'null',
    '二進位亂碼': '\x00\x01\x02\xff\xfe garbage',
  };
  for (const name of Object.keys(corrupt)) {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    fs.writeFileSync(file, corrupt[name], 'utf8');
    const ad = jsonFileAdapter({ file });
    const ob = mkOb(ad, clock);
    let threw = null;
    let list = null;
    try { list = await ob.list({}); } catch (e) { threw = e; }
    t('損壞（' + name + '）：讀取不 throw、視為空', !threw && list && list.total === 0, threw && threw.message);
    const baks = fs.readdirSync(dir).filter((n) => /\.corrupt-.*\.bak$/.test(n));
    t('損壞（' + name + '）：備份成 *.corrupt-<ts>.bak（落在 .gitignore 的 *.bak 規則內）', baks.length === 1, baks.join(','));
    if (baks.length === 1) eq('損壞（' + name + '）：備份內容與原檔相同', fs.readFileSync(path.join(dir, baks[0]), 'utf8'), corrupt[name]);
    const r = await ob.enqueue(mkJob(1));
    t('損壞（' + name + '）：之後可正常寫入', r.created === true && JSON.parse(fs.readFileSync(file, 'utf8')).jobs.length === 1);
    eq('損壞（' + name + '）：diagnostics 記錄了復原', ad.diagnostics().recoveries.length, 1);
  }
  // 空檔／全空白：視為空，不產生備份
  for (const content of ['', '   \n\t ']) {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    fs.writeFileSync(file, content, 'utf8');
    const ob = mkOb(jsonFileAdapter({ file }), clock);
    await ob.enqueue(mkJob(1));
    t('空檔案（' + short(content) + '）：視為空且不產生備份', fs.readdirSync(dir).filter((n) => /corrupt/.test(n)).length === 0);
  }
  // BOM
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob0 = mkOb(jsonFileAdapter({ file }), clock);
    const a = await ob0.enqueue(mkJob(1));
    fs.writeFileSync(file, cp(0xfeff) + fs.readFileSync(file, 'utf8'), 'utf8');
    const ob = mkOb(jsonFileAdapter({ file }), clock);
    t('檔案開頭有 BOM 仍可讀取（不算損壞）', (await ob.get(a.id)) !== null && fs.readdirSync(dir).filter((n) => /corrupt/.test(n)).length === 0);
  }
  // 部分紀錄損壞：保留合格的，原檔複製一份備份
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob0 = mkOb(jsonFileAdapter({ file }), clock);
    const a = await ob0.enqueue(mkJob(1));
    const st = JSON.parse(fs.readFileSync(file, 'utf8'));
    st.jobs.push({ id: 'bad1' }, 'junk', null, { id: 'bad2', dedupeKey: 'k', status: 'weird' });
    fs.writeFileSync(file, JSON.stringify(st), 'utf8');
    const ad = jsonFileAdapter({ file });
    const ob = mkOb(ad, clock);
    const l = await ob.list({});
    t('部分紀錄損壞：保留合格的 1 筆', l.total === 1 && l.rows[0].id === a.id, short(l));
    t('部分紀錄損壞：留了備份，主檔仍在使用', fs.readdirSync(dir).some((n) => /\.corrupt-.*\.bak$/.test(n)) && fs.existsSync(file));
    t('部分紀錄損壞：diagnostics 記錄丟棄筆數', /DROPPED_4/.test(JSON.stringify(ad.diagnostics())), short(ad.diagnostics()));
  }
  // 部分紀錄損壞＋反覆唯讀操作：同一個損壞狀態只備份一次（以前 list／stats／get／getByKey 每呼叫一次就複製一份，備份檔無上限增長）
  {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));        // 讓舊版每次備份的時間戳不同（否則同毫秒會互相覆蓋而看不出增長）
    const bakNames = (dir) => fs.readdirSync(dir).filter((n) => /\.corrupt-.*\.bak$/.test(n)).sort();
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob0 = mkOb(jsonFileAdapter({ file }), clock);
    const a = await ob0.enqueue(mkJob(1));
    const st = JSON.parse(fs.readFileSync(file, 'utf8'));
    st.jobs.push({ bogus: true });
    const corruptText = JSON.stringify(st);
    fs.writeFileSync(file, corruptText, 'utf8');
    const ad = jsonFileAdapter({ file });
    const ob = mkOb(ad, clock);
    for (let i = 0; i < 8; i++) {
      await ob.list({});
      await ob.stats();
      await ob.get(a.id);
      await ad.getByKey(mkJob(1).dedupeKey);
      await sleep(3);
    }
    eq('反覆唯讀操作 8 輪×4 種（32 次讀取）：只有 1 份備份', bakNames(dir).length, 1);
    eq('備份內容＝損壞當時的原檔', fs.readFileSync(path.join(dir, bakNames(dir)[0]), 'utf8'), corruptText);
    eq('diagnostics 只記一次復原（反覆讀取不重複記）', ad.diagnostics().recoveries.length, 1);
    eq('唯讀操作沒有改動主檔（損壞紀錄仍在原檔）', fs.readFileSync(file, 'utf8'), corruptText);
    eq('損壞狀態下合格的紀錄仍可讀', (await ob.get(a.id)).id, a.id);
    // 一次寫入操作會把乾淨版本存回去；之後讀取不再產生備份
    await ob.enqueue(mkJob(2));
    await ob.list({});
    eq('寫入後主檔已是乾淨版本、不再備份', [bakNames(dir).length, JSON.parse(fs.readFileSync(file, 'utf8')).jobs.every((j) => !j.bogus)], [1, true]);
    eq('份數上限常數是 3', require(path.join(ROOT, 'lib/mail/outboxAdapters.js')).MAX_CORRUPT_BACKUPS, 3);
  }
  // 不同的損壞狀態各備份一次，但備份檔總數有上限（最新的 3 份）；diagnostics 也有上限
  {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const bakNames = (dir) => fs.readdirSync(dir).filter((n) => /\.corrupt-.*\.bak$/.test(n)).sort();
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob0 = mkOb(jsonFileAdapter({ file }), clock);
    await ob0.enqueue(mkJob(1));
    const good = JSON.parse(fs.readFileSync(file, 'utf8'));
    const ad = jsonFileAdapter({ file });
    const ob = mkOb(ad, clock);
    let lastText = '';
    for (let i = 0; i < 60; i++) {
      const st = JSON.parse(JSON.stringify(good));
      st.jobs.push({ bogus: i });
      lastText = JSON.stringify(st);
      fs.writeFileSync(file, lastText, 'utf8');                    // 每一輪是不同的損壞內容
      await ob.list({});
      await ob.list({});                                           // 同一狀態讀兩次：第二次不再備份
      await sleep(1);
    }
    const names = bakNames(dir);
    eq('60 種不同損壞狀態：備份檔最多留 3 份', names.length, 3);
    eq('留下的是最新的那份（內容＝最後一次損壞）', fs.readFileSync(path.join(dir, names[names.length - 1]), 'utf8'), lastText);
    t('diagnostics.recoveries 有上限（50）', ad.diagnostics().recoveries.length <= 50 && ad.diagnostics().recoveries.length > 0, ad.diagnostics().recoveries.length);
    eq('目錄裡除了主檔與備份沒有別的東西（沒有殘留暫存檔）', fs.readdirSync(dir).filter((n) => !/\.corrupt-.*\.bak$/.test(n)), ['mail-outbox.json']);
  }
  // 整檔損壞（無法解析）反覆發生：每種內容搬走一次，總數同樣受上限約束
  {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ad = jsonFileAdapter({ file });
    const ob = mkOb(ad, clock);
    for (let i = 0; i < 7; i++) {
      fs.writeFileSync(file, '{"version":1,"jobs":[{"id":"half' + i, 'utf8');
      await ob.list({});
      await sleep(2);
    }
    eq('整檔損壞 7 次：備份最多 3 份', fs.readdirSync(dir).filter((n) => /\.corrupt-.*\.bak$/.test(n)).length, 3);
  }
  // 備份失敗或修剪失敗都不影響讀取（注入的 fs 沒有 readdirSync／copyFileSync 會丟例外）
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob0 = mkOb(jsonFileAdapter({ file }), clock);
    const a = await ob0.enqueue(mkJob(1));
    const st = JSON.parse(fs.readFileSync(file, 'utf8'));
    st.jobs.push({ bogus: true });
    fs.writeFileSync(file, JSON.stringify(st), 'utf8');
    const noReaddir = Object.assign({}, fs, { readdirSync() { throw new Error('no readdir'); } });
    const ob1 = mkOb(jsonFileAdapter({ file, fs: noReaddir }), clock);
    eq('修剪（readdirSync）失敗：讀取照常、備份仍建立', [(await ob1.get(a.id)).id, fs.readdirSync(dir).filter((n) => /\.corrupt-.*\.bak$/.test(n)).length], [a.id, 1]);
    const noCopy = Object.assign({}, fs, { copyFileSync() { throw new Error('no copy'); } });
    const ad2 = jsonFileAdapter({ file, fs: noCopy });
    const ob2 = mkOb(ad2, clock);
    await ob2.list({});
    await ob2.list({});
    eq('複製失敗：讀取照常；失敗的備份不算「已備份」，下次讀取會再試（診斷記下 backup:null）', ad2.diagnostics().recoveries.map((r) => r.backup), [null, null]);
  }
  // 殘留的半寫暫存檔不影響主檔
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const ob0 = mkOb(jsonFileAdapter({ file }), clock);
    await ob0.enqueue(mkJob(1));
    fs.writeFileSync(file + '.9999.deadbeef.tmp', '{"version":1,"jobs":[{"id":"half', 'utf8');
    const ob = mkOb(jsonFileAdapter({ file }), clock);
    eq('殘留半寫暫存檔：主檔照常讀取', (await ob.list({})).total, 1);
  }
  // 兩個 adapter 實例同時搶同一個檔案（同程序）：每次操作都重新讀檔，不會重複領取
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const obA = mkOb(jsonFileAdapter({ file }), clock);
    const obB = mkOb(jsonFileAdapter({ file }), clock);
    for (let i = 0; i < 20; i++) await obA.enqueue(mkJob(i));
    const claims = await Promise.all(Array.from({ length: 50 }, (_, i) => (i % 2 ? obA : obB).claimDue({ limit: 1 })));
    const ids = claims.reduce((a, b) => a.concat(b.map((j) => j.id)), []);
    eq('兩個 adapter 實例共用同一檔、50 個並行領取：20 筆各領一次', [ids.length, new Set(ids).size], [20, 20]);
  }
  // 在 Vercel 上拒絕使用檔案／記憶體 adapter
  {
    const old = process.env.VERCEL;
    try {
      process.env.VERCEL = '1';
      const codeOf = (mk) => { try { mk(); return null; } catch (e) { return e && e.code; } };
      eq('Vercel 上建立 jsonFileAdapter：丟 NOT_FOR_VERCEL', codeOf(() => jsonFileAdapter({ file: path.join(tmpDir(), 'x.json') })), 'NOT_FOR_VERCEL');
      eq('Vercel 上建立 memoryAdapter：丟 NOT_FOR_VERCEL', codeOf(() => memoryAdapter()), 'NOT_FOR_VERCEL');
      eq('Vercel 上以 allowOnVercel:true 明確放行才能建立', [codeOf(() => memoryAdapter({ allowOnVercel: true })), codeOf(() => jsonFileAdapter({ file: path.join(tmpDir(), 'y.json'), allowOnVercel: true }))], [null, null]);
      eq('Vercel 上 postgresAdapter 不受影響', codeOf(() => postgresAdapter({ query: async () => ({ rows: [] }) })), null);
      delete process.env.VERCEL;
      eq('非 Vercel 環境照常可建立', [codeOf(() => memoryAdapter()), codeOf(() => jsonFileAdapter({ file: path.join(tmpDir(), 'z.json') }))], [null, null]);
    } finally {
      if (old === undefined) delete process.env.VERCEL; else process.env.VERCEL = old;
    }
  }
  // 筆數上限的保險
  {
    const dir = tmpDir();
    const ad = jsonFileAdapter({ file: path.join(dir, 'x.json') });
    t('jsonFileAdapter 有 diagnostics', typeof ad.diagnostics === 'function' && Array.isArray(ad.diagnostics().recoveries));
  }
});

// ═════════════════════════════════════════════════════════════════════════
// 4) Postgres adapter 靜態審查與假 query
// ═════════════════════════════════════════════════════════════════════════
section('4 Postgres adapter（假 query）', async () => {
  const SQL = postgresAdapter.SQL;
  const DDL = postgresAdapter.DDL;
  const all = Object.keys(SQL).map((k) => [k, SQL[k]]).concat(Object.keys(DDL).map((k) => ['DDL.' + k, DDL[k]]));

  // 靜態：SQL 常數本身
  t('SQL 常數不是空的', all.length >= 20, all.length);
  all.forEach(([name, sql]) => {
    t('SQL ' + name + '：是字串且沒有未展開的 ${}', typeof sql === 'string' && sql.indexOf('${') < 0);
    t('SQL ' + name + '：沒有 DROP／TRUNCATE', !/\b(DROP|TRUNCATE)\b/i.test(sql));
    const nums = Array.from(sql.matchAll(/\$(\d+)/g)).map((m) => Number(m[1]));
    const max = nums.length ? Math.max.apply(null, nums) : 0;
    let contiguous = true;
    for (let i = 1; i <= max; i++) if (nums.indexOf(i) < 0) contiguous = false;
    t('SQL ' + name + '：$n 編號連續（最大 $' + max + '）', contiguous);
    if (/^\s*(UPDATE|DELETE)/i.test(sql) || /\bDELETE FROM\b/i.test(sql)) t('SQL ' + name + '：UPDATE／DELETE 都有 WHERE', /\bWHERE\b/i.test(sql));
  });
  t('SQL 裡的單引號字面值都是常數（任何一個字面值內都沒有 $n 佔位符）', all.every(([n, s]) => (s.match(/'[^']*'/g) || []).every((lit) => !/\$\d/.test(lit))));

  // 靜態：原始碼裡每個 q(...) 呼叫都傳 SQL 常數
  const src = fs.readFileSync(path.join(ROOT, 'lib/mail/outboxAdapters.js'), 'utf8');
  const callSites = src.match(/await q\([^)]*/g) || [];
  t('原始碼中所有 await q( 呼叫都傳 SQL.xxx 或 sql（共 ' + callSites.length + ' 處）', callSites.length > 10 && callSites.every((s) => /^await q\((SQL\.[A-Za-z]+|sql),/.test(s)), callSites.filter((s) => !/^await q\((SQL\.[A-Za-z]+|sql),/.test(s)).join(' | '));
  t('原始碼沒有用 + 或模板字串組 SQL 呼叫', !/q\([^)]*\+/.test(src) && !/query\([^)]*`/.test(src));

  // 語意指紋
  t('claimDue 使用 FOR UPDATE SKIP LOCKED', /FOR UPDATE SKIP LOCKED/.test(SQL.claimDue));
  t('claimDue 外層重複檢查到期條件且包含租約過期與本輪次數', /status = 'pending' AND next_attempt_at <= \$1::timestamptz/.test(SQL.claimDue) && /status = 'sending' AND lease_until <= \$1::timestamptz AND attempts_in_round < \$4/.test(SQL.claimDue) && (SQL.claimDue.match(/attempts_in_round < \$4/g) || []).length === 2);
  t('claimDue 領取時 attempts 與 attempts_in_round 都加一', /attempts = attempts \+ 1/.test(SQL.claimDue) && /attempts_in_round = attempts_in_round \+ 1/.test(SQL.claimDue));
  t('claimById 只領 pending 且到期', /WHERE id = \$1 AND status = 'pending' AND next_attempt_at <= \$2::timestamptz/.test(SQL.claimById));
  t('finishAttempt 有狀態與 attempts 樂觀鎖', /WHERE id = \$1 AND status = 'sending' AND attempts = \$2/.test(SQL.finishAttempt));
  t('insert 用 ON CONFLICT (dedupe_key) DO NOTHING', /ON CONFLICT \(dedupe_key\) DO NOTHING/.test(SQL.insert));
  t('dedupe_key 有唯一索引', /CREATE UNIQUE INDEX IF NOT EXISTS \S+ ON mail_outbox \(dedupe_key\)/.test(DDL.uniqueKey));
  t('建表使用 IF NOT EXISTS 且有三個 CHECK 約束', /CREATE TABLE IF NOT EXISTS mail_outbox/.test(DDL.createOutbox) && (DDL.createOutbox.match(/CHECK \(/g) || []).length === 3);
  t('兩張表都啟用 RLS', /ENABLE ROW LEVEL SECURITY/.test(DDL.rlsOutbox) && /ENABLE ROW LEVEL SECURITY/.test(DDL.rlsBreaker));
  t('expireExhausted 帶 RETURNING *（被清掃的列要回傳，派送器才能記 Summary／稽核而不是靜默轉換）', /RETURNING \*/.test(SQL.expireExhausted) && /last_error_code = 'LEASE_EXPIRED'/.test(SQL.expireExhausted) && /attempts_in_round >= \$3/.test(SQL.expireExhausted));
  t('purge 只刪終態且夠舊的', /status IN \('sent', 'failed', 'skipped', 'cancelled'\) AND updated_at < \$1::timestamptz/.test(SQL.purge));
  t('markSent 不會把 skipped 改成 sent', !/skipped/.test(SQL.markSent));
  t('requeue 只處理 failed／cancelled／skipped', /status IN \('failed', 'cancelled', 'skipped'\)/.test(SQL.requeue));
  t('時間參數一律 ::timestamptz、不用資料庫 now()', !/\bnow\(\)/i.test(all.map((x) => x[1]).join('\n')));

  // 假 query：DDL 只跑一次（20 個並行操作共用）
  {
    const fake = makeFakePg();
    const ad = postgresAdapter({ query: fake.query });
    const ob = createOutbox(ad, { now: () => T0, config: getMailConfig({}) });
    await Promise.all(Array.from({ length: 20 }, (_, i) => ob.enqueue(mkJob(i))));
    eq('20 個並行操作：每個 DDL 只執行一次', fake.ddlRuns.length, postgresAdapter.DDL_ORDER.length);
    eq('DDL 執行順序：表在索引與 RLS 之前', fake.ddlRuns.map((s) => s.split(/\s+/).slice(0, 3).join(' ')), postgresAdapter.DDL_ORDER.map((k) => DDL[k].split(/\s+/).slice(0, 3).join(' ')));
    await ob.list({});
    eq('建表之後的操作不會再跑 DDL', fake.ddlRuns.length, postgresAdapter.DDL_ORDER.length);
  }
  // 並行建表競態：第一次失敗（模擬 pg_type 重複鍵）→ 重試一次成功
  {
    let failed = false;
    const fake = makeFakePg({ ddlHook(sql, n) { if (n === 1 && !failed) { failed = true; throw new Error('duplicate key value violates unique constraint "pg_type_typname_nsp_index"'); } } });
    const ad = postgresAdapter({ query: fake.query });
    const ob = createOutbox(ad, { now: () => T0, config: getMailConfig({}) });
    const r = await ob.enqueue(mkJob(1));
    t('建表競態：重試一次後成功', failed && r.created === true);
  }
  // 持續失敗：reject，且下一次操作會重新嘗試建表
  {
    const fake = makeFakePg({ failNext: 2, failMessage: 'connection refused' });
    const ad = postgresAdapter({ query: fake.query });
    const ob = createOutbox(ad, { now: () => T0, config: getMailConfig({}) });
    let threw = false;
    try { await ob.enqueue(mkJob(1)); } catch (e) { threw = true; }
    t('建表持續失敗：操作 reject', threw);
    const r = await ob.enqueue(mkJob(1));
    t('資料庫恢復後下一次操作重新建表並成功', r.created === true);
  }
  // RLS 必須「真的被執行」：只檢查 DDL 文字存在不夠（有人把 'rlsOutbox' 從執行清單拿掉，表就會被 Supabase PostgREST 匿名金鑰讀到）。
  // 下面的檢查刻意不引用 DDL_ORDER 的內容，而是看假資料庫「實際收到」哪些語句。
  {
    const fake = makeFakePg();
    const ad = postgresAdapter({ query: fake.query });
    await ad.init();
    const ran = fake.ddlRuns;
    const idxOf = (re) => ran.findIndex((s) => re.test(s));
    t('RLS：ALTER TABLE mail_outbox ENABLE ROW LEVEL SECURITY 真的被送到資料庫', ran.some((s) => s === 'ALTER TABLE mail_outbox ENABLE ROW LEVEL SECURITY'), short(ran.map((s) => s.slice(0, 40))));
    t('RLS：ALTER TABLE mail_breaker ENABLE ROW LEVEL SECURITY 真的被送到資料庫', ran.some((s) => s === 'ALTER TABLE mail_breaker ENABLE ROW LEVEL SECURITY'), short(ran.map((s) => s.slice(0, 40))));
    t('RLS 在對應的表建好之後才執行', idxOf(/^CREATE TABLE IF NOT EXISTS mail_outbox/) >= 0 && idxOf(/^CREATE TABLE IF NOT EXISTS mail_outbox/) < idxOf(/^ALTER TABLE mail_outbox ENABLE/)
      && idxOf(/^CREATE TABLE IF NOT EXISTS mail_breaker/) >= 0 && idxOf(/^CREATE TABLE IF NOT EXISTS mail_breaker/) < idxOf(/^ALTER TABLE mail_breaker ENABLE/));
    const tables = ran.map((s) => (s.match(/^CREATE TABLE IF NOT EXISTS (\w+)/) || [])[1]).filter(Boolean);
    t('實際建出的每一張表（' + tables.join('、') + '）都有一條 ENABLE ROW LEVEL SECURITY 被執行', tables.length === 2 && tables.every((tn) => ran.some((s) => s === 'ALTER TABLE ' + tn + ' ENABLE ROW LEVEL SECURITY')));
    eq('DDL_ORDER 涵蓋 DDL 物件的每一個語句（沒有「定義了卻不執行」的 DDL）', postgresAdapter.DDL_ORDER.slice().sort(), Object.keys(DDL).sort());
    eq('實際執行的 DDL 語句＝DDL 物件的全部語句（各一次）', ran.slice().sort(), Object.keys(DDL).map((k) => DDL[k]).sort());
  }
  // 動態 SQL 守衛（q() 的 SQL_SET 檢查）：非常數 SQL 必須被擋下、且完全不會送到資料庫。execConstant 走與內部相同的 q()
  {
    const fake = makeFakePg();
    const ad = postgresAdapter({ query: fake.query });
    const n0 = fake.calls.length;
    await rejectsCode('任意 SQL 文字被拒（DYNAMIC_SQL）', ad.execConstant('SELECT 1', []), 'DYNAMIC_SQL');
    await rejectsCode('SQL 常數後面接注入字串也被拒（DYNAMIC_SQL）', ad.execConstant(SQL.getById + " OR '1'='1'", ['x']), 'DYNAMIC_SQL');
    await rejectsCode('DROP TABLE 被拒（DYNAMIC_SQL）', ad.execConstant('DROP TABLE mail_outbox', []), 'DYNAMIC_SQL');
    await rejectsCode('空字串被拒（DYNAMIC_SQL）', ad.execConstant('', []), 'DYNAMIC_SQL');
    await rejectsCode('非字串（toString 回傳常數的物件）被拒（DYNAMIC_SQL）', ad.execConstant({ toString: () => SQL.getById }, ['x']), 'DYNAMIC_SQL');
    eq('被拒的 SQL 一個都沒有送到資料庫', fake.calls.length, n0);
    const ok = await ad.execConstant(SQL.getById, ['nope']);
    t('SQL 常數可以執行（會送到 query）', Array.isArray(ok) && fake.calls.length === n0 + 1);
    await rejectsCode('常數＋物件參數：BAD_PARAM', ad.execConstant(SQL.getById, [{ a: 1 }]), 'BAD_PARAM');
    await rejectsCode('常數＋非陣列參數：BAD_PARAM', ad.execConstant(SQL.getById, 'x'), 'BAD_PARAM');
    eq('BAD_PARAM 也不會送到資料庫', fake.calls.length, n0 + 1);
  }
  // 注入字串：只出現在參數，不出現在 SQL
  {
    const fake = makeFakePg();
    const ad = postgresAdapter({ query: fake.query });
    const ob = createOutbox(ad, { now: () => T0, config: getMailConfig({}) });
    const evil = "u'; DROP TABLE mail_outbox; --";
    const evil2 = "x' OR '1'='1";
    await ob.enqueue(mkJob(1, { toUser: evil, dedupeKey: 'k:' + evil, quoteNo: 'QU-9' }));
    await ob.list({ toUser: evil, quoteNo: evil2, type: evil2 });
    await ob.markSkipped((await ob.list({})).rows[0].id, 'NO_EMAIL');
    await ob.get(evil);
    await ob.claim(evil2);
    const sqlText = fake.calls.map((c) => c.sql).join('\n');
    t('注入字串（DROP TABLE／OR 1=1）沒有出現在任何 SQL 文字', sqlText.indexOf('DROP TABLE mail_outbox; --') < 0 && sqlText.indexOf("OR '1'='1") < 0);
    t('注入字串確實是以參數傳遞', fake.calls.some((c) => c.params.indexOf(evil) >= 0));
    t('每一次 query 呼叫的 sql 都是 SQL／DDL 常數之一', fake.calls.every((c) => all.some((x) => x[1] === c.sql)));
    t('每一次 query 呼叫的 params 都是陣列', fake.calls.every((c) => Array.isArray(c.params)));
    eq('注入字串當帳號存進去後可原樣查回', (await ob.list({ toUser: evil })).total, 1);
  }
  // 執行期護欄：動態 SQL、壞參數
  {
    const seen = [];
    const ad = postgresAdapter({ query: async (s, p) => { seen.push(s); return { rows: [] }; } });
    await ad.init();                                       // DDL 本身是常數 → 通過
    t('DDL 常數可以執行', seen.length === postgresAdapter.DDL_ORDER.length);
  }
  {
    const fake = makeFakePg();
    const ad = postgresAdapter({ query: fake.query });
    // 取出內部 q 的行為：用會傳出物件參數的呼叫來驗證
    await rejectsCode('參數含物件：被執行期護欄擋下（BAD_PARAM）', ad.finishAttempt({ id: 'x', expectAttempts: 1, now: iso(T0), status: 'pending', nextAttemptAt: iso(T0), code: { evil: 1 }, msg: 'm' }), 'BAD_PARAM');
    await rejectsCode('參數含 NaN：被擋下（BAD_PARAM）', ad.claimDue({ now: iso(T0), leaseUntil: iso(T0), limit: NaN, maxRoundAttempts: 4 }), 'BAD_PARAM');
    t('postgresAdapter 沒給 query 函式：建立時就丟 BAD_ADAPTER', (() => { try { postgresAdapter({}); return false; } catch (e) { return e.code === 'BAD_ADAPTER'; } })());
    t('createOutbox 沒給 adapter：丟 BAD_ADAPTER', (() => { try { createOutbox(null); return false; } catch (e) { return e.code === 'BAD_ADAPTER'; } })());
    t('createOutbox 的 adapter 缺 expireExhausted：丟 BAD_ADAPTER（不能悄悄少了清掃步驟）', (() => { try { createOutbox({ insert() {}, claimDue() {} }); return false; } catch (e) { return e.code === 'BAD_ADAPTER'; } })());
  }
  // stats／list 的型別對映：pg 回 Date 物件與數字
  {
    const e = mkEnv('pg');
    const a = await e.outbox.enqueue(mkJob(1));
    const rec = await e.outbox.get(a.id);
    t('Date 物件被轉成 ISO 字串', typeof rec.createdAt === 'string' && rec.createdAt === iso(T0) && rec.nextAttemptAt === iso(T0), short(rec));
    t('insert 的參數順序：21 個，第 7 個是 dedupe_key，第 8 個是 status', (() => {
      const c = e.fake.calls.find((x) => x.sql === SQL.insert);
      return c.params.length === 21 && c.params[6] === mkJob(1).dedupeKey && c.params[7] === 'pending' && c.params[0] === a.id;
    })());
    t('claimDue 的參數順序：[now, leaseUntil, limit, maxRoundAttempts]', await (async () => {
      await e.outbox.claimDue({ limit: 7 });
      const c = e.fake.calls.filter((x) => x.sql === SQL.claimDue).pop();
      return c.params[0] === iso(T0) && c.params[1] === iso(T0 + 60 * SEC) && c.params[2] === 7 && c.params[3] === 4;
    })());
    t('claimDue 前一定先跑 expireExhausted', (() => {
      const idx = e.fake.calls.map((x) => x.sql);
      const i = idx.lastIndexOf(SQL.claimDue);
      return idx[i - 1] === SQL.expireExhausted;
    })());
  }
});

// ═════════════════════════════════════════════════════════════════════════
// 5) scrubMessage
// ═════════════════════════════════════════════════════════════════════════
section('5 scrubMessage', () => {
  const cases = [
    ['Bearer 權杖', 'Authorization failed: Bearer abcDEF0123456789xyz.token', /abcDEF0123456789/],
    ['JWT', 'got eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJ4In0.c2lnbmF0dXJl here', /eyJhbGci/],
    ['client_secret=…', 'invalid client_secret=Zq8Q~abcdefghijklmnop', /Zq8Q|abcdefghij/],
    ['password: 值', 'password: hunter2hunter2', /hunter2/],
    ['JSON 風格 "token":"…"', 'body {"access_token": "AbCdEf1234567890"}', /AbCdEf123/],
    ['authorization 標頭', 'authorization=Basic dXNlcjpwYXNz', /dXNlcjpw/],
    ['email', 'recipient first.last@itts.com.tw not found', /first\.last@/],
    ['GUID（租戶 ID）', 'tenant 00000000-0000-4000-8000-0000000000aa is disabled', /00000000-0000-4000/],
    ['32+ 字元不透明字串', 'key=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcd end', /ABCDEFGHIJKLMNOPQRSTUVWX/],
    ['Azure 風格密鑰（含 ~ 與 .）', 'secret Abc8Q~dEfGhIjKlMnOpQrStUvWxYz.0123456789 x', /Abc8Q~dEfG/],
  ];
  cases.forEach(([name, input, bad]) => {
    const out = scrubMessage(input);
    t('scrubMessage：' + name + ' 被遮蔽', !bad.test(out) && out.length <= 200, out);
  });
  eq('scrubMessage：一般訊息原樣保留', scrubMessage('Graph 回應 503 Service Unavailable'), 'Graph 回應 503 Service Unavailable');
  eq('scrubMessage：GUID 換成 [id] 標記', scrubMessage('tenant 00000000-0000-4000-8000-0000000000aa is disabled'), 'tenant [id] is disabled');
  eq('scrubMessage：email 換成 [email] 標記', scrubMessage('no such user first.last@itts.com.tw here'), 'no such user [email] here');
  eq('scrubMessage：保留 AADSTS 這類夾在字母間的錯誤代碼，只遮獨立的大數字', scrubMessage('AADSTS7000215 amount 1234567800 and 12,345,678'), 'AADSTS7000215 amount [數字] and [數字]');
  eq('scrubMessage：Bearer 權杖換成標記', scrubMessage('denied Bearer abcdefgh12345678 now'), 'denied Bearer [已遮蔽] now');
  eq('scrubMessage：非字串輸入', [scrubMessage(null), scrubMessage(undefined), scrubMessage(123), scrubMessage({})], ['', '', '123', '[object Object]']);
  t('scrubMessage：單行（換行變空白）', scrubMessage('a\r\nb\nc').indexOf('\n') < 0 && scrubMessage('a\r\nb') === 'a b');
  t('scrubMessage：結果最多 200 字', scrubMessage('word '.repeat(500)).length <= 201);
  eq('scrubMessage：known 清單中的字串被遮蔽', scrubMessage('failed for 專案甲乙丙 and Subject-XYZ', ['專案甲乙丙', 'Subject-XYZ']), 'failed for [已遮蔽] and [已遮蔽]');
  t('scrubMessage：known 太短（<3）不處理，避免誤傷', scrubMessage('ab ab', ['ab']) === 'ab ab');
  t('scrubMessage：重複套用結果不變（冪等）', (() => { const o = scrubMessage('x@itts.com.tw Bearer abcdefgh12345678'); return scrubMessage(o) === o; })());
  fast('scrubMessage 200 萬字元不透明字串', () => scrubMessage('A'.repeat(2000000)), 50);
  fast('scrubMessage 200 萬個 @', () => scrubMessage('@'.repeat(2000000)), 50);
  fast('scrubMessage 200 萬字元 a.a.a.', () => scrubMessage('a.'.repeat(1000000)), 50);
  fast('scrubMessage 200 萬字元 password: 重複', () => scrubMessage('password:'.repeat(222222)), 50);
});

main();
