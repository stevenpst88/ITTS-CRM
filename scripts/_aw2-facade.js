'use strict';
// check-audit-wave2.js 的 F 項：安全開關 AUDIT_EVENTS_TABLE（預設關＝舊路徑；開＝新表；表不能用時自動退回舊路徑）。
// 全部用假 pg／暫存資料夾；不連任何資料庫、不碰專案根目錄的資料檔。
const fs = require('fs');
const os = require('os');
const path = require('path');
const ROOT = path.join(__dirname, '..');

/** 第二波之前的舊程式碼（逐字保留，當作「位元級相同」的對照組） */
function refJsonWriteLog(dir, entry) {   // 原 server.js writeLog 的本機 JSON 分支
  const logFile = path.join(dir, 'audit.log.json');
  let logs = [];
  try { logs = JSON.parse(fs.readFileSync(logFile, 'utf8')); } catch {}
  logs.unshift(entry);
  if (logs.length > 500) logs = logs.slice(0, 500);
  fs.writeFileSync(logFile, JSON.stringify(logs, null, 2), 'utf8');
}
function refContactAudit(dir, entry) {   // 原 server.js writeContactAudit
  const f = path.join(dir, 'contact_audit.json');
  let logs = [];
  try { logs = JSON.parse(fs.readFileSync(f, 'utf8')); } catch {}
  logs.unshift(entry);
  if (logs.length > 3000) logs = logs.slice(0, 3000);
  try { fs.writeFileSync(f, JSON.stringify(logs, null, 2), 'utf8'); } catch {}
}
const norm = (t) => t.replace(/"id": "[^"]*"/g, '"id": "X"').replace(/"timestamp": "[^"]*"/g, '"timestamp": "T"');

async function runF(h, opt) {
  const { ok, eq, makeFakePg } = h;
  opt = opt || {};
  const S = opt.store || require(path.join(ROOT, 'lib/auditStore.js'));
  const AL = opt.legacy || require(path.join(ROOT, 'lib/auditLegacy.js'));
  const errs = [];
  const origErr = console.error;
  const quiet = (fn) => async () => { console.error = (...a) => { errs.push(a.join(' ')); }; try { return await fn(); } finally { console.error = origErr; } };
  const mkFacade = (fake, enabled, legacyCalls, legacyRows) => S.createAuditFacade({
    enabled,
    tableStore: S.createAuditStore(S.pgAdapter({ query: fake.query, legacyTtlMs: 0 }), { retryDelayMs: 1, fallbackFile: path.join(os.tmpdir(), 'aw2-facade-fb-' + process.pid + '.jsonl') }),
    legacyStore: S.createAuditStore(S.legacyReadAdapter({ readLogs: async () => legacyRows || [] })),
    legacyWrite: async (ev) => { legacyCalls.push(ev); },
  });
  const ev = (n) => ({ kind: 'log', action: 'F_' + n, operator: 'u', target: 't', detail: 'd' + n });

  // ── F1 旗標關（預設）：完全不碰資料庫表；讀寫走舊路徑 ──
  {
    const fake = makeFakePg(); const calls = [];
    const fa = mkFacade(fake, false, calls, [{ id: 'L1', timestamp: '2026-01-01T00:00:00.000Z', action: 'OLD', operator: 'o', target: '', detail: 'legacy row' }]);
    eq(fa.mode(), 'legacy', '[F1] 預設模式 legacy');
    eq(await fa.insert(ev(1)), true, '[F1] 寫入成功');
    await fa.insert({ kind: 'contact', action: 'UPDATE', operator: 'u', target: 'c', detail: 'x', extra: { userName: 'U', targetId: 'c', targetName: 'n', targetCompany: 'co', changes: [] } });
    await fa.insert({ kind: 'quote-edit', action: 'DRAFT_SAVE', operator: 'u', target: 'QU-1', detail: 'd' });
    eq(calls.length, 3, '[F1] 三種事件都交給舊路徑');
    const l = await fa.list({ kind: 'log' }, {});
    eq([l.total, l.rows[0].detail], [1, 'legacy row'], '[F1] 讀取只看舊資料');
    eq((await fa.total({ kind: 'all' })), 1, '[F1] total 只算舊資料');
    const it = []; for await (const r of fa.iterate({ kind: 'all' })) it.push(r.id); eq(it, ['L1'], '[F1] 匯出迭代器只走舊資料');
    eq(fake.log.length, 0, '[F1] 假 pg 完全沒收到任何語句（沒有 CREATE TABLE、沒有 SELECT）');
    eq(fa.stats().auditStore, 'legacy', '[F1] stats.auditStore=legacy');
    await fa.init(); eq(fake.log.length, 0, '[F1] init() 在旗標關時也不碰資料庫');
  }
  // ── F2 旗標開、表正常：自我檢查順序與寫入 ──
  {
    const fake = makeFakePg(); const calls = [];
    const fa = mkFacade(fake, true, calls);
    await fa.insert(ev(1));
    const seq = fake.log.map((q) => q.split(' ').slice(0, 5).join(' '));
    ok(/^CREATE TABLE IF NOT EXISTS audit_events/.test(fake.log[0]), '[F2] 第一句是 CREATE TABLE IF NOT EXISTS audit_events');
    const iPing = fake.log.findIndex((q) => q === 'SELECT 1 FROM audit_events LIMIT 1');
    const iIns = fake.log.findIndex((q) => /^INSERT INTO audit_events/.test(q));
    ok(iPing > 0 && iIns > iPing, '[F2] 先自我檢查（SELECT 1 FROM audit_events LIMIT 1）再寫入：' + seq.join(' | '));
    eq([fake.db.rows.length, calls.length, fa.mode()], [1, 0, 'table'], '[F2] 只進新表、舊路徑沒被叫');
    // 並行：自我檢查只做一次
    const f2 = makeFakePg(); const fb = mkFacade(f2, true, []);
    await Promise.all(Array.from({ length: 20 }, (_, i) => fb.insert(ev(i))));
    eq([f2.db.rows.length, f2.log.filter((q) => /^CREATE TABLE/.test(q)).length, f2.log.filter((q) => q === 'SELECT 1 FROM audit_events LIMIT 1').length], [20, 1, 1], '[F2] 20 筆並行寫入：建表與自我檢查各只做 1 次');
  }
  // ── F3 表不可用（建表失敗）→ 自動退回舊路徑 ──
  await quiet(async () => {
    const fake = makeFakePg(); fake.failTable(true); const calls = [];
    const fa = mkFacade(fake, true, calls, [{ id: 'L1', timestamp: '2026-01-01T00:00:00.000Z', action: 'OLD', operator: 'o', target: '', detail: 'legacy row' }]);
    const r = await fa.insert(ev(1));
    eq([r, fa.mode(), calls.length], [true, 'legacy', 1], '[F3] 表不可用：寫入仍成功、已退回舊路徑、事件在舊路徑');
    ok(errs.some((e) => e.includes('AUDIT_TABLE_UNAVAILABLE')), '[F3] 記了 AUDIT_TABLE_UNAVAILABLE');
    const n0 = fake.log.length;
    await fa.insert(ev(2)); await fa.insert(ev(3));
    eq([calls.length, fake.log.length], [3, n0], '[F3] 之後不再碰資料庫表（語句數不變），事件全進舊路徑');
    eq((await fa.list({ kind: 'log' }, {})).rows.map((x) => x.detail), ['legacy row'], '[F3] 讀取改看舊資料');
    const st = fa.stats();
    eq([st.auditStore, st.fellBack], ['legacy', true], '[F3] stats 標示 legacy／fellBack');
  })();
  // ── F4 建表成功但 SELECT 被拒（權限）→ 同樣退回 ──
  await quiet(async () => {
    const fake = makeFakePg(); fake.failPing(true); const calls = [];
    const fa = mkFacade(fake, true, calls);
    const r = await fa.insert(ev(1));
    eq([r, fa.mode(), calls.length, fake.db.rows.length], [true, 'legacy', 1, 0], '[F4] 自我檢查查詢失敗 → 退回舊路徑，表內沒有寫入');
  })();
  // ── F5 開啟後單筆寫入失敗（重試後）→ 記 AUDIT_WRITE_FAILED，同時寫舊路徑 ──
  await quiet(async () => {
    const fake = makeFakePg(); const calls = [];
    const fa = mkFacade(fake, true, calls);
    await fa.insert(ev(0));
    fake.failAlways(true);
    const r = await fa.insert(ev(1));
    fake.failAlways(false);
    eq([r, calls.length, calls[0] && calls[0].action], [false, 1, 'F_1'], '[F5] 新表寫入失敗：回報 false，但事件同時寫進舊路徑');
    ok(errs.some((e) => e.includes('AUDIT_WRITE_FAILED')), '[F5] 有 AUDIT_WRITE_FAILED 標記');
    eq(fa.mode(), 'table', '[F5] 單筆失敗不改模式（表本身可用）');
  })();
  // ── F6 舊路徑寫入丟例外也不影響請求 ──
  await quiet(async () => {
    const fa = S.createAuditFacade({ enabled: false, tableStore: S.createAuditStore(S.pgAdapter({ query: makeFakePg().query })), legacyStore: S.createAuditStore(S.legacyReadAdapter({ readLogs: async () => [] })), legacyWrite: () => { throw new Error('disk full'); } });
    eq(await fa.insert(ev(1)), false, '[F6] 舊路徑寫入例外 → insert 回 false 而不是拋出');
  })();

  // ── F7 舊路徑檔案：與第二波之前位元級相同 ──
  {
    const dirA = fs.mkdtempSync(path.join(os.tmpdir(), 'aw2-legA-')), dirB = fs.mkdtempSync(path.join(os.tmpdir(), 'aw2-legB-'));
    try {
      const evs = [];
      for (let i = 0; i < 3; i++) evs.push({ kind: 'log', action: 'ACT' + i, operator: 'op', target: 'tg', detail: i === 1 ? { nested: 1 } : '說明 ' + i, ip: '1.2.3.4', ua: 'UA/1' });
      for (const e of evs) {
        AL.writeJsonMode(e, dirA);
        refJsonWriteLog(dirB, { id: 'i', timestamp: 't', action: e.action, operator: e.operator, target: e.target, detail: e.detail, ip: e.ip, userAgent: e.ua });
      }
      eq(norm(fs.readFileSync(path.join(dirA, 'audit.log.json'), 'utf8')), norm(fs.readFileSync(path.join(dirB, 'audit.log.json'), 'utf8')), '[F7] audit.log.json 內容（欄位順序、縮排、排序）與舊程式位元級相同');
      // 500 筆上限
      const pre = Array.from({ length: 520 }, (_, i) => ({ id: 'p' + i, timestamp: 't' + i, action: 'P', operator: 'o', target: '', detail: '', ip: '', userAgent: '' }));
      fs.writeFileSync(path.join(dirA, 'audit.log.json'), JSON.stringify(pre, null, 2)); fs.writeFileSync(path.join(dirB, 'audit.log.json'), JSON.stringify(pre, null, 2));
      AL.writeJsonMode(evs[0], dirA); refJsonWriteLog(dirB, { id: 'i', timestamp: 't', action: evs[0].action, operator: evs[0].operator, target: evs[0].target, detail: evs[0].detail, ip: evs[0].ip, userAgent: evs[0].ua });
      const a = JSON.parse(fs.readFileSync(path.join(dirA, 'audit.log.json'), 'utf8'));
      eq(a.length, 500, '[F7] 本機舊路徑上限 500 筆');
      eq(norm(fs.readFileSync(path.join(dirA, 'audit.log.json'), 'utf8')), norm(fs.readFileSync(path.join(dirB, 'audit.log.json'), 'utf8')), '[F7] 超過上限時與舊程式位元級相同');
      // 名片稽核檔
      const ce = { kind: 'contact', action: 'UPDATE', operator: 'u1', target: 'c1', detail: 'n / co', extra: { userName: 'U One', targetId: 'c1', targetName: 'n', targetCompany: '', changes: [{ field: 'title', fieldLabel: '職稱', oldValue: 'a', newValue: 'b' }] } };
      AL.writeJsonMode(ce, dirA);
      refContactAudit(dirB, { id: 'i', timestamp: 't', userId: 'u1', userName: 'U One', action: 'UPDATE', targetId: 'c1', targetName: 'n', targetCompany: '', changes: ce.extra.changes });
      eq(norm(fs.readFileSync(path.join(dirA, 'contact_audit.json'), 'utf8')), norm(fs.readFileSync(path.join(dirB, 'contact_audit.json'), 'utf8')), '[F7] contact_audit.json 與舊程式位元級相同');
      const big = Array.from({ length: 3005 }, (_, i) => ({ id: 'c' + i, timestamp: 't', userId: 'u', userName: 'U', action: 'C', targetId: 'x', targetName: '', targetCompany: '', changes: [] }));
      fs.writeFileSync(path.join(dirA, 'contact_audit.json'), JSON.stringify(big)); AL.writeJsonMode(ce, dirA);
      eq(JSON.parse(fs.readFileSync(path.join(dirA, 'contact_audit.json'), 'utf8')).length, 3000, '[F7] 名片稽核舊路徑上限 3000 筆');
      // Postgres 舊項目欄位順序
      const pe = AL.pgLogEntry({ action: 'A', operator: 'o', target: 't', detail: 'd', ip: 'i', ua: 'u' });
      eq(Object.keys(pe), ['id', 'action', 'operator', 'target', 'detail', 'ip', 'userAgent', 'timestamp'], '[F7] Postgres 舊項目欄位順序與舊程式相同');
      eq(Object.keys(AL.jsonLogEntry({ action: 'A' })), ['id', 'timestamp', 'action', 'operator', 'target', 'detail', 'ip', 'userAgent'], '[F7] 本機舊項目欄位順序與舊程式相同');
      // quote-edit 轉一般日誌
      AL.writeJsonMode({ kind: 'quote-edit', action: 'DRAFT_SAVE', operator: 'u', target: 'QU-1', detail: '改備註' }, dirA);
      const top = JSON.parse(fs.readFileSync(path.join(dirA, 'audit.log.json'), 'utf8'))[0];
      eq([top.action, top.target, top.detail], ['QUOTE_EDIT', 'QU-1', 'DRAFT_SAVE｜改備註'], '[F7] 旗標關時報價單原始編輯事件改寫成舊路徑的一般日誌 QUOTE_EDIT');
      ok(!fs.existsSync(path.join(dirA, 'audit.log.jsonl')), '[F7] 舊路徑不產生 audit.log.jsonl');
    } finally { for (const d of [dirA, dirB]) try { fs.rmSync(d, { recursive: true, force: true }); } catch { /* ignore */ } }
  }

  // ── F8 真的 db/postgres.js（假 pg Pool）：旗標關＝舊整包 JSONB 路徑，不碰 audit_events ──
  if (!opt.skipPg) await runPgModule(h, ev, quiet, errs, opt);
}

async function runPgModule(h, ev, quiet, errs, opt) {
  const { ok, eq, makeFakePg } = h;
  const pgPath = require.resolve(path.join(ROOT, 'node_modules', 'pg'));
  const dbFile = (opt && opt.dbFile) || path.join(ROOT, 'db', 'postgres.js');
  const storeFile = path.join(ROOT, 'lib', 'auditStore.js');
  const saveEnv = { u: process.env.DATABASE_URL, f: process.env.AUDIT_EVENTS_TABLE };
  const savePg = require.cache[pgPath];
  const load = (flag, fake) => {
    class Pool { constructor() {} query(sql, p) { return fake.query(sql, p); } on() {} }
    require.cache[pgPath] = { id: pgPath, filename: pgPath, loaded: true, exports: { Pool } };
    process.env.DATABASE_URL = 'postgres://fake.invalid/none';
    if (flag === undefined) delete process.env.AUDIT_EVENTS_TABLE; else process.env.AUDIT_EVENTS_TABLE = flag;
    delete require.cache[dbFile]; delete require.cache[storeFile];
    return require(dbFile);
  };
  try {
    // 旗標關（預設）
    {
      const fake = makeFakePg();
      fake.db.legacy = { logs: Array.from({ length: 5000 }, (_, i) => ({ id: 'b' + i, action: 'B', operator: 'o', target: '', detail: 'b' + i, ip: '', userAgent: '', timestamp: new Date(Date.UTC(2026, 0, 1) + i * 1000).toISOString() })).reverse() };
      const pg = load(undefined, fake);
      await pg.ready();
      await pg.insertAuditEvent({ kind: 'log', action: 'NEW1', operator: 'u', target: 't', detail: 'd', ip: '9.9.9.9', ua: 'UA' });
      await pg.insertAuditEvent({ kind: 'log', action: 'NEW2', operator: 'u', target: 't', detail: 'd2' });
      await pg.flush();
      eq(pg.auditStore.mode(), 'legacy', '[F8] 預設（未設 AUDIT_EVENTS_TABLE）模式 legacy');
      ok(fake.log.every((q) => !/audit_events/i.test(q)), '[F8] 預設：任何語句都不含 audit_events（沒建表、沒查表）');
      const blob = fake.db.legacy.logs;
      eq([blob.length, blob[0].action, blob[1].action], [5000, 'NEW2', 'NEW1'], '[F8] 整包 JSONB 新的在最前、維持 5000 筆上限');
      eq(Object.keys(blob[1]), ['id', 'action', 'operator', 'target', 'detail', 'ip', 'userAgent', 'timestamp'], '[F8] 整包項目欄位順序與舊程式相同');
      eq([blob[1].ip, blob[1].userAgent], ['9.9.9.9', 'UA'], '[F8] ip／userAgent 照舊');
      const l = await pg.auditStore.list({ kind: 'log' }, { limit: 3 });
      eq([l.total, l.rows.map((r) => r.action)], [5000, ['NEW2', 'NEW1', 'B']], '[F8] 讀取只看整包 JSONB（分頁在記憶體）');
      ok(fake.log.every((q) => !/audit_events/i.test(q)), '[F8] 讀取後仍不碰 audit_events');
      ok((await pg.auditStore.list({ kind: 'log', q: 'b4999' }, {})).total === 1, '[F8] 記憶體文字篩選可用');
    }
    // 旗標開、表正常
    {
      const fake = makeFakePg(); const pg = load('1', fake);
      await pg.ready();
      await pg.insertAuditEvent({ kind: 'log', action: 'T1', operator: 'u', target: 't', detail: 'd' });
      await pg.flush();
      eq([pg.auditStore.mode(), fake.db.rows.length, fake.db.legacy], ['table', 1, null], '[F8] 旗標=1：進新表、模式 table、整包 JSONB 沒被寫');
      ok(fake.log.includes('SELECT 1 FROM audit_events LIMIT 1'), '[F8] 旗標=1：做了自我檢查');
    }
    // 旗標開、表不可用 → ready() 不丟例外、退回舊路徑、事件進整包 JSONB
    await quiet(async () => {
      const fake = makeFakePg(); fake.failTable(true); const pg = load('1', fake);
      let threw = false; try { await pg.ready(); } catch { threw = true; }
      await pg.insertAuditEvent({ kind: 'log', action: 'FB1', operator: 'u', target: 't', detail: 'd' });
      await pg.flush();
      eq([threw, pg.auditStore.mode()], [false, 'legacy'], '[F8] 旗標=1 但表不可用：ready() 不拋例外、自動退回 legacy');
      eq(fake.db.legacy && fake.db.legacy.logs[0].action, 'FB1', '[F8] 退回後事件寫進整包 JSONB（沒丟）');
      ok(errs.some((e) => e.includes('AUDIT_TABLE_UNAVAILABLE')), '[F8] 記了 AUDIT_TABLE_UNAVAILABLE');
    })();
    // 旗標值不是 '1'（例如 'true'）→ 仍是 legacy
    {
      const fake = makeFakePg(); const pg = load('true', fake);
      await pg.insertAuditEvent(ev(1)); await pg.flush();
      eq([pg.auditStore.mode(), fake.log.some((q) => /audit_events/i.test(q))], ['legacy', false], '[F8] AUDIT_EVENTS_TABLE=true（非 1）→ 仍 legacy，不碰表');
    }
  } finally {
    if (savePg) require.cache[pgPath] = savePg; else delete require.cache[pgPath];
    if (saveEnv.u === undefined) delete process.env.DATABASE_URL; else process.env.DATABASE_URL = saveEnv.u;
    if (saveEnv.f === undefined) delete process.env.AUDIT_EVENTS_TABLE; else process.env.AUDIT_EVENTS_TABLE = saveEnv.f;
    delete require.cache[dbFile]; delete require.cache[storeFile];
  }
}

module.exports = { runF };
