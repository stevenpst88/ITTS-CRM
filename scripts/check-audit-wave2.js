#!/usr/bin/env node
'use strict';
/**
 * 內控稽核第二波（audit wave 2）檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 *  A 稽核事件只增不改（lib/auditStore.js）：假 pg client 單元測試（INSERT-only、並行、分頁、篩選、舊資料合併、CSV 公式注入、失敗處理）
 *    ＋本機 JSONL 模式的 HTTP 端對端（分頁／篩選／匯出／名片稽核／舊檔唯讀）
 *  B 報價單刪除快照＋單號流水帳
 *  C 規則快照＋版本釘選
 *  D 編輯事件完整性
 *  E 董事會代核證據
 *  M 變異檢查
 *
 * 守門：工作樹若已有 auth.json / data.json / audit.log.json(l) （可能是真實資料），一律拒絕執行。
 * 做法：在本工作樹建臨時 auth.json / data.json（假資料 + _s_* 臨時帳號），以 PORT=AW2_PORT（預設 3131）啟動 server.js（JSON 模式，絕不連任何資料庫），
 * 跑完只用 PID 關閉並刪除臨時檔。禁用埠：3000 / 3020 / 3030 / 3101 / 3111 / 3121。
 *
 *   node scripts/check-audit-wave2.js                # 全部
 *   AW2_ONLY=A,B node scripts/check-audit-wave2.js   # 只跑指定項目
 *   AW2_SKIP_MUT=1 ...                               # 略過變異檢查
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.AW2_PORT || '3131', 10);
if ([3000, 3020, 3030, 3101, 3111, 3121].includes(PORT)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json', 'mail-outbox.json'].map(f => path.join(ROOT, f));
const guardFiles = FILES.filter(f => ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'contact_audit.json'].includes(path.basename(f)) && fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0, muting = false, mutFail = 0;
const failures = [];
function ok(cond, name) {
  if (cond) { if (!muting) pass++; return; }
  if (muting) { mutFail++; return; }
  fail++; failures.push(name); console.log('  ✗ FAIL:', name);
}
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));
const sha = (f) => { try { return crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex'); } catch { return 'none'; } };

// ═════════════ A-1：假 pg client ═════════════
// 只實作 lib/auditStore.js 會送出的語句：CREATE TABLE/INDEX、INSERT、SELECT count、SELECT rows(WHERE/ORDER BY id DESC/LIMIT)、讀 app_data.audit_log。
// 不認得的語句一律丟錯（所以任何 UPDATE／DELETE／TRUNCATE 都會在測試中現形）。
function likeToRe(p) {
  let re = '';
  const esc = (c) => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  for (let i = 0; i < p.length; i++) {
    const ch = p[i];
    if (ch === '\\' && i + 1 < p.length) re += esc(p[++i]);
    else if (ch === '%') re += '[\\s\\S]*';
    else if (ch === '_') re += '[\\s\\S]';
    else re += esc(ch);
  }
  return new RegExp('^' + re + '$', 'i');
}
function makeFakePg(shared) {
  const db = shared || { rows: [], seq: 0, legacy: null, ddl: [] };
  const log = [];
  let failNext = 0, failAlways = false, delayMs = 0;
  const api = {
    db, log,
    failNext(n) { failNext = n; }, failAlways(b) { failAlways = b; }, delay(ms) { delayMs = ms; },
    async query(sql, params) {
      params = params || [];
      log.push(sql.replace(/\s+/g, ' ').trim());
      if (delayMs) await sleep(delayMs);
      const s = sql.replace(/\s+/g, ' ').trim();
      if (/^CREATE TABLE IF NOT EXISTS audit_events/i.test(s) || /^CREATE INDEX IF NOT EXISTS audit_events_/i.test(s)) { db.ddl.push(s); return { rows: [] }; }
      if (/^INSERT INTO audit_events \(/i.test(s)) {
        if (failAlways || failNext > 0) { if (failNext > 0) failNext--; throw new Error('simulated connection reset'); }
        if (params.length !== 9) throw new Error('INSERT 參數數量不對');
        const [ts, kind, action, operator, target, detail, ip, ua, extra] = params;
        if (/\u0000/.test(String([ts, kind, action, operator, target, detail, ip, ua, extra].join('')))) throw new Error('invalid byte sequence 0x00');
        const ex = JSON.parse(extra);   // 必須是合法 JSON 字串
        db.rows.push({ id: ++db.seq, ts: new Date(ts), kind, action, operator, target, detail, ip, ua, extra: ex });
        return { rows: [] };
      }
      if (/^SELECT content FROM app_data WHERE id = 'audit_log'$/i.test(s)) return { rows: db.legacy ? [{ content: db.legacy }] : [] };
      let m = /^SELECT count\(\*\)::int AS n FROM audit_events(?: WHERE (.*))?$/i.exec(s);
      if (m) return { rows: [{ n: db.rows.filter(r => matchWhere(r, m[1], params)).length }] };
      m = /^SELECT id, ts, kind, action, operator, target, detail, ip, ua, extra FROM audit_events(?: WHERE (.*?))? ORDER BY id DESC LIMIT \$(\d+)$/i.exec(s);
      if (m) {
        const lim = params[parseInt(m[2], 10) - 1];
        const rows = db.rows.filter(r => matchWhere(r, m[1], params)).sort((a, b) => b.id - a.id).slice(0, lim);
        return { rows: rows.map(r => Object.assign({}, r)) };
      }
      throw new Error('假 pg 不認得的語句：' + s.slice(0, 120));
    },
  };
  return api;
}
function matchWhere(r, where, params) {
  if (!where) return true;
  return where.split(' AND ').every(c => {
    let m;
    if ((m = /^kind = \$(\d+)$/.exec(c))) return r.kind === params[m[1] - 1];
    if ((m = /^action = \$(\d+)$/.exec(c))) return r.action === params[m[1] - 1];
    if ((m = /^operator = \$(\d+)$/.exec(c))) return r.operator === params[m[1] - 1];
    if ((m = /^ts >= \$(\d+)::timestamptz$/.exec(c))) return r.ts.getTime() >= Date.parse(params[m[1] - 1]);
    if ((m = /^ts < \$(\d+)::timestamptz$/.exec(c))) return r.ts.getTime() < Date.parse(params[m[1] - 1]);
    if ((m = /^id < \$(\d+)$/.exec(c))) return r.id < params[m[1] - 1];
    if ((m = /^\(action ILIKE \$(\d+) OR operator ILIKE \$(\d+) OR target ILIKE \$(\d+) OR detail ILIKE \$(\d+)\)$/.exec(c))) {
      const re = likeToRe(params[m[1] - 1]);
      return re.test(r.action) || re.test(r.operator) || re.test(r.target) || re.test(r.detail);
    }
    throw new Error('假 pg 不認得的條件：' + c);
  });
}

// ═════════════ A-2：單元測試（對任一版本的 auditStore 模組）═════════════
async function unitStore(Mod, tag) {
  const S = Mod;
  const mk = (fake, opt) => S.createAuditStore(S.pgAdapter({ query: fake.query, legacyTtlMs: 0 }), Object.assign({ retryDelayMs: 1, fallbackFile: path.join(os.tmpdir(), 'aw2-fallback-' + process.pid + '.jsonl') }, opt || {}));
  const T = (day, h) => `${day}T${String(h).padStart(2, '0')}:00:00+08:00`;
  const origErr = console.error;
  const errs = [];

  // ── 1 INSERT-only／冪等建表 ──
  {
    const f = makeFakePg(); const st = mk(f);
    await st.insert({ kind: 'log', action: 'A1', operator: 'u', target: 't', detail: 'd' });
    await st.insert({ kind: 'log', action: 'A2', operator: 'u', target: 't', detail: 'd' });
    await st.list({ kind: 'log' }, {}); await st.total({ kind: 'all' });
    ok(f.log.every(q => !/\b(UPDATE|DELETE|TRUNCATE|DROP|ALTER)\b/i.test(q)), '[A1] 所有送出的 SQL 不含 UPDATE／DELETE／TRUNCATE／DROP／ALTER');
    eq(f.log.filter(q => /^INSERT INTO audit_events/i.test(q)).length, 2, '[A1] 兩次寫入＝兩次單列 INSERT');
    ok(f.log.filter(q => /^CREATE TABLE IF NOT EXISTS audit_events/i.test(q)).length === 1, '[A1] 建表 IF NOT EXISTS 只送一次（記憶化）');
    eq(f.db.rows.length, 2, '[A1] 表內兩列');
  }
  // ── 2 並行：兩個「實例」交錯寫入，全部保留 ──
  {
    const shared = { rows: [], seq: 0, legacy: null, ddl: [] };
    const fa = makeFakePg(shared), fb = makeFakePg(shared);
    const sa = mk(fa), sb = mk(fb);
    const ps = [];
    for (let i = 0; i < 60; i++) { ps.push(sa.insert({ action: 'INST_A', operator: 'a', detail: 'a' + i })); ps.push(sb.insert({ action: 'INST_B', operator: 'b', detail: 'b' + i })); }
    await Promise.all(ps);
    eq(shared.rows.length, 120, '[A2] 兩實例交錯寫入 → 120 筆全部在');
    eq(shared.rows.filter(r => r.action === 'INST_A').length, 60, '[A2] 實例 A 的 60 筆沒被覆蓋');
    eq(shared.rows.filter(r => r.action === 'INST_B').length, 60, '[A2] 實例 B 的 60 筆沒被覆蓋');
    eq(new Set(shared.rows.map(r => r.id)).size, 120, '[A2] id 不重複');
    // 舊實作對照：整包讀改寫會丟筆（確認測試的意義：這裡每個實例只有自己的快取）
  }
  // ── 3 無筆數上限 ──
  {
    const f = makeFakePg(); const st = mk(f);
    for (let i = 0; i < 5200; i++) f.db.rows.push({ id: ++f.db.seq, ts: new Date(Date.parse('2026-03-01T00:00:00Z') + i * 1000), kind: 'log', action: 'BULK', operator: 'u', target: '', detail: 'n' + i, ip: '', ua: '', extra: {} });
    eq(await st.total({ kind: 'log' }), 5200, '[A3] 超過 5000 筆仍全數計入（無截斷）');
    const last = await st.list({ kind: 'log' }, { limit: 100, offset: 5100 });
    eq(last.rows.length, 100, '[A3] 第 5100 筆之後仍可翻頁讀到');
    eq(last.rows[last.rows.length - 1].detail, 'n0', '[A3] 最舊的一筆還在');
  }
  // ── 4 分頁正確性（跨 iterate 的 500 筆分塊）──
  {
    const f = makeFakePg(); const st = mk(f);
    for (let i = 0; i < 1203; i++) f.db.rows.push({ id: ++f.db.seq, ts: new Date(Date.parse('2026-04-01T00:00:00Z') + i * 1000), kind: 'log', action: i % 2 ? 'ODD' : 'EVEN', operator: 'u', target: '', detail: 'p' + i, ip: '', ua: '', extra: {} });
    const seen = [];
    for (const off of [0, 400, 800, 1200]) { const r = await st.list({ kind: 'log' }, { limit: 400, offset: off }); seen.push(...r.rows.map(x => x.detail)); eq(r.total, 1203, '[A4] total 固定 1203（offset ' + off + '）'); }
    eq(seen.length, 1203, '[A4] 各頁相加＝總數');
    eq(new Set(seen).size, 1203, '[A4] 各頁無重複');
    eq(seen[0], 'p1202', '[A4] 第一筆是最新');
    eq(seen[1202], 'p0', '[A4] 最後一筆是最舊');
    let sorted = true; for (let i = 1; i < seen.length; i++) if (parseInt(seen[i - 1].slice(1)) < parseInt(seen[i].slice(1))) sorted = false;
    ok(sorted, '[A4] 全程由新到舊');
    const ev = await st.list({ kind: 'log', action: 'EVEN' }, { limit: 1000 });
    eq(ev.total, 602, '[A4] 動作篩選的 total');
    ok(ev.rows.every(r => r.action === 'EVEN'), '[A4] 動作篩選只回該動作');
    const ofs = await st.list({ kind: 'log' }, { limit: 5, offset: 3 });
    eq(ofs.rows.map(x => x.detail), ['p1199', 'p1198', 'p1197', 'p1196', 'p1195'], '[A4] offset 生效');
  }
  // ── 5 篩選：類型／操作者／日期（台北日）／文字（含 LIKE 萬用字元）──
  {
    const f = makeFakePg(); const st = mk(f);
    await st.insert({ ts: T('2026-05-01', 1), kind: 'log', action: 'X', operator: 'alice', target: 'T1', detail: '折扣 100% 特價' });
    await st.insert({ ts: T('2026-05-01', 23), kind: 'log', action: 'X', operator: 'bob', target: 'T2', detail: '折扣 1000 一般' });
    await st.insert({ ts: T('2026-05-02', 0), kind: 'log', action: 'Y', operator: 'alice', target: 'T3', detail: 'a_b' });
    await st.insert({ ts: T('2026-05-03', 12), kind: 'contact', action: 'UPDATE', operator: 'alice', target: 'c1', detail: '王小明 / 甲公司', extra: { userName: 'A', targetId: 'c1', targetName: '王小明', targetCompany: '甲公司', changes: [{ field: 'title' }] } });
    await st.insert({ ts: T('2026-05-03', 13), kind: 'quote-edit', action: 'EDIT', operator: 'alice', target: 'Q-1', detail: 'x' });
    eq((await st.list({ kind: 'log' }, {})).total, 3, '[A5] 預設類型 log 不含 contact／quote-edit');
    eq((await st.list({ kind: 'all' }, {})).total, 5, '[A5] kind=all 全含');
    eq((await st.list({ kind: 'contact' }, {})).rows.length, 1, '[A5] kind=contact');
    eq((await st.list({ kind: 'log', operator: 'alice' }, {})).total, 2, '[A5] 操作者篩選');
    eq((await st.list({ kind: 'log', from: '2026-05-01', to: '2026-05-01' }, {})).total, 2, '[A5] 單日範圍含該日 00:00–24:00（台北）');
    eq((await st.list({ kind: 'log', from: '2026-05-02' }, {})).total, 1, '[A5] 只有 from');
    eq((await st.list({ kind: 'log', to: '2026-05-01' }, {})).total, 2, '[A5] 只有 to（含當日）');
    eq((await st.list({ kind: 'log', q: '100%' }, {})).rows.map(r => r.target), ['T1'], '[A5] 文字搜尋把 % 當字面（不是萬用字元）');
    eq((await st.list({ kind: 'log', q: 'a_b' }, {})).rows.map(r => r.target), ['T3'], '[A5] 文字搜尋把 _ 當字面');
    eq((await st.list({ kind: 'log', q: 't2' }, {})).rows.map(r => r.target), ['T2'], '[A5] 文字搜尋不分大小寫、含 target');
    eq((await st.list({ kind: 'log', action: 'X', operator: 'alice', from: '2026-05-01', q: '折扣' }, {})).total, 1, '[A5] 多條件 AND');
    const c = S.toApiContact((await st.list({ kind: 'contact' }, {})).rows[0]);
    eq([c.userId, c.targetName, c.targetCompany, c.changes.length, c.action], ['alice', '王小明', '甲公司', 1, 'UPDATE'], '[A5] 名片稽核輸出形狀（舊前端相容欄位）');
    const l = S.toApiLog((await st.list({ kind: 'log' }, {})).rows[0]);
    ok('userAgent' in l && 'timestamp' in l && 'detail' in l, '[A5] 一般日誌輸出形狀（userAgent／timestamp 相容欄位）');
  }
  // ── 6 舊整包 JSONB 合併 ──
  {
    const f = makeFakePg(); const st = mk(f);
    // 舊資料故意以「由舊到新」排列（真實 blob 是新→舊；合併端不得依賴順序）
    f.db.legacy = { logs: [
      { id: 'L1', action: 'OLD', operator: 'o', target: '', detail: 'l1', timestamp: '2026-01-01T00:00:00.000Z', ip: '1.1.1.1', userAgent: 'ua' },
      { id: 'L2', action: 'OLD', operator: 'o', target: '', detail: 'l2', timestamp: '2026-01-03T00:00:00.000Z' },
      { id: 'L3', action: 'OLD', operator: 'o', target: '', detail: 'l3', timestamp: '2026-01-05T00:00:00.000Z' },
    ] };
    await st.insert({ ts: '2026-01-02T00:00:00.000Z', action: 'NEW', operator: 'n', detail: 'n1' });
    await st.insert({ ts: '2026-01-04T00:00:00.000Z', action: 'NEW', operator: 'n', detail: 'n2' });
    await st.insert({ ts: '2026-01-06T00:00:00.000Z', action: 'NEW', operator: 'n', detail: 'n3' });
    const all = await st.list({ kind: 'log' }, {});
    eq(all.total, 6, '[A6] total＝新表＋舊資料');
    eq(all.rows.map(r => r.detail), ['n3', 'l3', 'n2', 'l2', 'n1', 'l1'], '[A6] 新舊依時間交錯由新到舊');
    eq((await st.list({ kind: 'log' }, { limit: 2, offset: 2 })).rows.map(r => r.detail), ['n2', 'l2'], '[A6] 跨新舊邊界分頁');
    eq((await st.list({ kind: 'log', action: 'OLD' }, {})).total, 3, '[A6] 舊資料也吃動作篩選');
    eq((await st.list({ kind: 'log', from: '2026-01-04', to: '2026-01-05' }, {})).rows.map(r => r.detail), ['l3', 'n2'], '[A6] 日期篩選套用到舊資料（台北日）');
    ok(f.log.every(q => !/\b(UPDATE|DELETE)\b/i.test(q)), '[A6] 讀舊資料不寫回、不改');
    const o = S.toApiLog(all.rows[5]);
    eq([o.id, o.ip, o.userAgent], ['L1', '1.1.1.1', 'ua'], '[A6] 舊資料欄位對應（id／ip／userAgent）');
  }
  // ── 7 CSV 儲存格與公式注入 ──
  {
    eq(S.csvCell('=1+1'), "'=1+1", '[A7] = 開頭加單引號');
    eq(S.csvCell('+1'), "'+1", '[A7] + 開頭加單引號');
    eq(S.csvCell('-2'), "'-2", '[A7] - 開頭加單引號');
    eq(S.csvCell('@SUM(A1)'), "'@SUM(A1)", '[A7] @ 開頭加單引號');
    eq(S.csvCell('\tcmd'), "'\tcmd", '[A7] Tab 開頭加單引號');
    eq(S.csvCell('a=b'), 'a=b', '[A7] 中間的 = 不動');
    eq(S.csvCell('他說"好",行二\n行三'), '"他說""好"",行二\n行三"', '[A7] 逗號／引號／換行以 RFC4180 包起來');
    eq(S.csvCell('=HYPERLINK("http://x","y")'), '"\'=HYPERLINK(""http://x"",""y"")"', '[A7] 公式＋引號：先加單引號再包引號');
    eq(S.csvCell(null), '', '[A7] null → 空字串');
    const line = S.csvLine({ ts: '2026-01-01T00:00:00.000Z', kind: 'log', action: 'A', operator: '=evil', target: '', detail: 'x', ip: '', ua: '', extra: {} });
    ok(line.startsWith('2026-01-01T00:00:00.000Z,log,A,\'=evil,') && line.endsWith('\r\n'), '[A7] csvLine 套用防護');
  }
  // ── 8 失敗處理 ──
  {
    const f = makeFakePg();
    const fb = path.join(os.tmpdir(), 'aw2-fb-' + process.pid + '-' + (tag || 'x') + '.jsonl');
    try { fs.unlinkSync(fb); } catch { /* 無 */ }
    const st = mk(f, { fallbackFile: fb });
    f.failNext(1);
    const r1 = await st.insert({ action: 'RETRY_OK', operator: 'u' });
    eq(r1, true, '[A8] 第一次失敗、重試成功 → 寫入成功');
    eq(f.db.rows.filter(r => r.action === 'RETRY_OK').length, 1, '[A8] 重試後恰一筆（不重複）');
    eq([st.stats().retried, st.stats().failed], [1, 0], '[A8] 統計：重試 1、失敗 0');
    console.error = (...a) => { errs.push(a.join(' ')); };
    f.failAlways(true);
    const r2 = await st.insert({ action: 'ALWAYS_FAIL', operator: 'u', detail: 'secret-free' });
    console.error = origErr;
    eq(r2, false, '[A8] 兩次都失敗 → insert 仍 resolve(false)，不拋例外');
    ok(errs.some(e => e.includes('AUDIT_WRITE_FAILED') && e.includes('ALWAYS_FAIL')), '[A8] console 有 AUDIT_WRITE_FAILED 標記與事件內容');
    ok(fs.existsSync(fb) && fs.readFileSync(fb, 'utf8').includes('AUDIT_WRITE_FAILED') && fs.readFileSync(fb, 'utf8').includes('ALWAYS_FAIL'), '[A8] 備援檔留下完整事件');
    eq(st.stats().failed, 1, '[A8] 失敗計數 = 1');
    ok(st.stats().lastError.includes('simulated'), '[A8] 記錄最後錯誤');
    f.failAlways(false);
    eq(await st.insert({ action: 'AFTER' }), true, '[A8] 之後恢復正常寫入');
    try { fs.unlinkSync(fb); } catch { /* 已無 */ }
    // flush 會等進行中的寫入
    const g = makeFakePg(); const st2 = mk(g); g.delay(40);
    st2.insert({ action: 'SLOW' });
    eq(g.db.rows.length, 0, '[A8] 慢寫入尚未完成');
    await st2.flush();
    eq(g.db.rows.length, 1, '[A8] flush() 等到寫入完成（wave-1 的回應前 flush 仍有效）');
  }
  // ── 9 NUL 與過長欄位 ──
  {
    const f = makeFakePg(); const st = mk(f);
    eq(await st.insert({ action: 'NUL', operator: 'u', detail: 'a\u0000b', extra: { k: 'x\u0000y' } }), true, '[A9] 含 NUL 的內容可寫入（先剝除）');
    eq(f.db.rows[0].detail, 'ab', '[A9] NUL 被移除');
    await st.insert({ action: 'BIG', detail: 'x'.repeat(50000) });
    eq(f.db.rows[1].detail.length, 20000, '[A9] 單欄上限 20000 字');
  }
  // ── 10 本機 JSONL 轉接器：無 500 上限、舊檔唯讀合併、壞行不影響 ──
  {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'aw2-jsonl-'));
    try {
      const file = path.join(dir, 'a.jsonl');
      const oldLog = path.join(dir, 'old.json'); const oldC = path.join(dir, 'oldc.json');
      fs.writeFileSync(oldLog, JSON.stringify([{ id: 'ol1', timestamp: '2020-01-02T00:00:00.000Z', action: 'LEG', operator: 'o', target: '', detail: 'legacy log' }]));
      fs.writeFileSync(oldC, JSON.stringify([{ id: 'oc1', timestamp: '2020-01-03T00:00:00.000Z', userId: 'u1', userName: 'U', action: 'UPDATE', targetId: 'c9', targetName: '舊名片', targetCompany: '舊公司', changes: [{ field: 'name' }] }]));
      const oldSha = [sha(oldLog), sha(oldC)];
      const st = S.createAuditStore(S.jsonlAdapter({ file, legacyLogFile: oldLog, legacyContactFile: oldC }), { retryDelayMs: 1, fallbackFile: path.join(dir, 'fb.jsonl') });
      for (let i = 0; i < 700; i++) await st.insert({ action: 'J', operator: 'u', detail: 'j' + i });
      fs.appendFileSync(file, '{壞掉的一行\n');
      await st.insert({ kind: 'contact', action: 'CREATE', operator: 'u2', target: 'cc', detail: '新 / 公司', extra: { userName: 'U2', targetId: 'cc', targetName: '新', targetCompany: '公司', changes: [] } });
      eq((await st.list({ kind: 'log' }, {})).total, 701, '[A10] JSONL：700 筆 + 1 筆舊 → 701（超過舊上限 500）');
      eq((await st.list({ kind: 'log' }, { limit: 1, offset: 700 })).rows[0].detail, 'legacy log', '[A10] JSONL：最舊的是舊檔的紀錄');
      eq((await st.list({ kind: 'contact' }, {})).total, 2, '[A10] JSONL：名片稽核＝新 1 ＋ 舊 1');
      eq(fs.readFileSync(file, 'utf8').split('\n').filter(l => l.startsWith('{"id"')).length, 701, '[A10] 檔案是只附加：701 行有效 JSON（含 1 筆名片）＋1 壞行保留原樣');
      eq([sha(oldLog), sha(oldC)], oldSha, '[A10] 舊檔未被改動（唯讀）');
    } finally { try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignore */ } }
  }
}

// 全專案不得出現對 audit_events 的 UPDATE／DELETE／TRUNCATE（保護 SQL 檔與本檢查腳本自身除外）
function scanNoMutation() {
  const hits = [];
  const skipDir = new Set(['node_modules', '.git', 'scratchpad', 'uploads', 'storage']);
  (function walk(dir) {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      if (e.isDirectory()) { if (!skipDir.has(e.name)) walk(path.join(dir, e.name)); continue; }
      if (!/\.(js|mjs|cjs|html|sql|ps1|json)$/.test(e.name)) continue;
      const p = path.join(dir, e.name);
      if (p.endsWith('audit_events_protect.sql') || p.endsWith('check-audit-wave2.js')) continue;
      let t = ''; try { t = fs.readFileSync(p, 'utf8'); } catch { continue; }
      if (/\b(UPDATE\s+(public\.)?audit_events|DELETE\s+FROM\s+(public\.)?audit_events|TRUNCATE\s+(TABLE\s+)?(public\.)?audit_events)\b/i.test(t)) hits.push(path.relative(ROOT, p));
    }
  })(ROOT);
  return hits;
}

// ═════════════ A-3：端對端（本機 JSONL 模式）═════════════
const PW = 'Aw2' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
let seedExtra = null;   // 各項目追加資料的鉤子
function seedFiles() {
  const users = [
    mkUser('_s_admin', 'admin'), mkUser('_s_admin2', 'admin'),
    mkUser('_s_u1', 'user', { bu: ['ERP'], canDownloadContacts: true }),
    mkUser('_s_u2', 'user', { bu: ['ERP'] }),
  ];
  const d = {
    contacts: [{ id: 'c1', name: '測試甲', company: '沙盒公司甲', owner: '_s_u1', bu: 'ERP', email: 'a@sandbox.test', phone: '0200000001', createdAt: nowIso }],
    companies: [{ id: 'co1', name: '沙盒公司甲', taxId: '12345675', industry: '製造業' }],
    opportunities: [], visits: [], contracts: [], quotations: [], receivables: [], callins: [], keyAccounts: [], groups: [],
  };
  const files = { users, d };
  if (seedExtra) seedExtra(files);
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users: files.users }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(files.d, null, 2));
  if (files.legacyLog) fs.writeFileSync(path.join(ROOT, 'audit.log.json'), JSON.stringify(files.legacyLog, null, 2));
  if (files.legacyContact) fs.writeFileSync(path.join(ROOT, 'contact_audit.json'), JSON.stringify(files.legacyContact, null, 2));
}
function cleanupFiles() { FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }

let child = null, childPid = null, extraEnv = {};
async function startServer(serverFile) {
  const env = Object.assign({}, process.env, extraEnv, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE;
  child = spawn(process.execPath, [serverFile || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.AW2_DEBUG) process.stdout.write('[srv] ' + d); };
  child.stdout.on('data', onData); child.stderr.on('data', onData);
  child.serverOut = () => out;
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-500));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-500));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100);
  child = null;
}
async function restart(serverFile, env) { await stopServer(); extraEnv = env || {}; await startServer(serverFile); jar = {}; }

let jar = {};
async function http(who, method, url, body, opt) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  if (opt && opt.ua) h['User-Agent'] = opt.ua;
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fresh = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fresh) jar[who] = fresh.split(';')[0];
  return { status: res.status, json, buf, text: buf.toString('utf8'), headers: res.headers, setCookie: sc };
}
async function login(who, user, pw) {
  const r = await http(null, 'POST', '/api/login', { username: user || who, password: pw || PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
async function loginAll(names) { for (const n of names) { const r = await login(n); if (r.status !== 200) throw new Error('登入失敗 ' + n + ' ' + r.status); } }
async function getLogs(qs) { const r = await http('_s_admin', 'GET', '/api/admin/logs' + (qs || '')); return r.json; }
const arr = (j) => (Array.isArray(j) ? j : []);
const byAction = (logs, a) => arr(logs).filter(l => l.action === a);

async function secA() {
  console.log('[A] 稽核事件只增不改');
  // —— 單元（正式模組）——
  const Store = require(path.join(ROOT, 'lib', 'auditStore.js'));
  await unitStore(Store, 'real');
  eq(scanNoMutation(), [], '[A] 全專案掃描：沒有任何對 audit_events 的 UPDATE／DELETE／TRUNCATE');
  const protect = fs.readFileSync(path.join(ROOT, 'supabase', 'audit_events_protect.sql'), 'utf8');
  ok(/BEFORE UPDATE OR DELETE ON public\.audit_events/.test(protect) && /BEFORE TRUNCATE ON public\.audit_events/.test(protect) && /REVOKE UPDATE, DELETE, TRUNCATE/.test(protect), '[A] 保護 SQL：含 UPDATE/DELETE trigger、TRUNCATE trigger、REVOKE');
  ok(/^\s*--.*audit_events/m.test(protect) && !/DROP TABLE|DELETE FROM/i.test(protect.replace(/--.*$/gm, '')), '[A] 保護 SQL：註解齊全、沒有破壞性語句');
  const pgSrc = fs.readFileSync(path.join(ROOT, 'db', 'postgres.js'), 'utf8');
  ok(/CREATE TABLE IF NOT EXISTS audit_events/.test(Store.TABLE_SQL) && /ensure/.test(fs.readFileSync(path.join(ROOT, 'lib', 'auditStore.js'), 'utf8')) && /insertAuditEvent/.test(pgSrc), '[A] db/postgres.js 掛上 insertAuditEvent，ensureSchema 於首次寫入建表');
  // 表欄位與規格一致
  for (const col of ['id        BIGSERIAL', 'ts        TIMESTAMPTZ', 'kind      TEXT', 'action', 'operator', 'target', 'detail', 'ip', 'ua', 'extra     JSONB']) ok(Store.TABLE_SQL.includes(col), '[A] 表欄位 ' + col.trim());

  // —— 端對端 ——
  const logFile = path.join(ROOT, 'audit.log.json'), contactFile = path.join(ROOT, 'contact_audit.json'), jsonl = path.join(ROOT, 'audit.log.jsonl');
  const legacyShas = [sha(logFile), sha(contactFile)];
  await loginAll(['_s_admin', '_s_u1']);
  // 產生事件：不同帳號的 LOGIN_FAIL（含公式注入字串）
  for (const nm of ['=HYPERLINK("http://x","y")', '+cmd', '-cmd', '@sum', 'plain_user']) await http(null, 'POST', '/api/login', { username: nm, password: 'x' });
  // 名片：建立 → 修改 → 刪除（writeContactAudit 路徑）
  let r = await http('_s_u1', 'POST', '/api/contacts', { name: '稽核乙', company: '沙盒公司甲', taxId: '12345675', bu: 'ERP', email: 'z@sandbox.test' });
  ok([200, 201].includes(r.status), '[A] 新增聯絡人 (' + r.status + ')');
  const cid = r.json && (r.json.id || (r.json.contact && r.json.contact.id));
  r = await http('_s_u1', 'PUT', '/api/contacts/' + cid, { name: '稽核乙改', company: '沙盒公司甲', taxId: '12345675', bu: 'ERP', email: 'z@sandbox.test', title: '經理' });
  ok(r.status === 200, '[A] 修改聯絡人 (' + r.status + ')');
  r = await http('_s_u1', 'DELETE', '/api/contacts/' + cid);
  ok(r.status === 200, '[A] 刪除聯絡人 (' + r.status + ')');

  // 相容：不帶參數＝純陣列、含舊資料
  let logs = await getLogs();
  ok(Array.isArray(logs), '[A] 不帶分頁參數 → 純陣列（舊前端／wave-1 檢查相容）');
  ok(arr(logs).some(l => l.id === 'LEG-OLD-1'), '[A] 舊 audit.log.json 的紀錄併入顯示');
  ok(arr(logs).every((l, i, a) => i === 0 || Date.parse(a[i - 1].timestamp) >= Date.parse(l.timestamp)), '[A] 新→舊排序');
  ok(!arr(logs).some(l => l.action === 'CREATE' && l.operator === '_s_u1' && l.target === cid), '[A] kind 預設 log：名片欄位異動不混進一般日誌');
  // 分頁
  const p1 = await getLogs('?paged=1&limit=5&offset=0');
  eq([p1.limit, p1.offset, p1.rows.length], [5, 0, 5], '[A] 分頁回應形狀');
  ok(p1.total >= 8, '[A] total 為全部筆數 (' + p1.total + ')');
  const p2 = await getLogs('?paged=1&limit=5&offset=5');
  ok(p2.rows.length > 0 && !p2.rows.some(x => p1.rows.some(y => y.id === x.id)), '[A] 第二頁與第一頁不重複');
  const all = (await getLogs('?paged=1&limit=1000')).rows;
  eq(all.length, p1.total, '[A] limit=1000 一次取完＝total');
  eq(all[all.length - 1].id, 'LEG-OLD-1', '[A] 最舊一筆是舊檔紀錄');
  // 篩選
  eq((await getLogs('?paged=1&action=LEGACY_ACT')).total, 2, '[A] 動作篩選（含舊資料）');
  ok((await getLogs('?paged=1&action=LOGIN_FAIL')).total >= 5, '[A] LOGIN_FAIL 篩選');
  eq((await getLogs('?paged=1&operator=legacy_op')).total, 1, '[A] 操作者篩選（舊資料）');
  eq((await getLogs('?paged=1&from=2020-01-01&to=2020-01-31')).total, 2, '[A] 日期範圍（舊資料 2020-01）');
  eq((await getLogs('?paged=1&q=' + encodeURIComponent('舊紀錄二'))).total, 1, '[A] 文字搜尋');
  ok((await getLogs('?paged=1&from=' + new Date(Date.now() + 86400000).toISOString().slice(0, 10))).total === 0, '[A] 未來日期範圍為空');
  // 名片稽核
  const ca = await http('_s_admin', 'GET', '/api/admin/contact-audit?paged=1&limit=50');
  ok(ca.status === 200 && Array.isArray(ca.json.rows), '[A] contact-audit 分頁形狀');
  const mine = ca.json.rows.filter(x => x.targetId === cid);
  eq(mine.map(x => x.action).sort(), ['CREATE', 'DELETE', 'UPDATE'], '[A] 名片 新增／修改／刪除 都進稽核（JSON 模式也進事件檔）');
  const upd = mine.find(x => x.action === 'UPDATE');
  ok(upd && upd.changes.some(c => c.field === 'name' && c.newValue === '稽核乙改') && upd.userId === '_s_u1', '[A] 修改的欄位級 before→after 與操作者');
  ok(ca.json.rows.some(x => x.id === 'LEG-C-1'), '[A] 舊 contact_audit.json 併入顯示');
  eq((await http('_s_admin', 'GET', '/api/admin/contact-audit?paged=1&userId=_s_u1&action=UPDATE')).json.total, 1, '[A] contact-audit 篩選 userId＋action');
  ok(Array.isArray((await http('_s_admin', 'GET', '/api/admin/contact-audit')).json), '[A] contact-audit 不帶參數＝舊的純陣列');
  // 檔案層：事件檔只增、舊檔不變
  eq([sha(logFile), sha(contactFile)], legacyShas, '[A] 舊 audit.log.json／contact_audit.json 位元級不變（唯讀歷史）');
  const lines = fs.readFileSync(jsonl, 'utf8').split('\n').filter(Boolean);
  ok(lines.length >= 8 && lines.every(l => { try { JSON.parse(l); return true; } catch { return false; } }), '[A] audit.log.jsonl 每行一筆合法 JSON (' + lines.length + ' 行)');
  const before = fs.readFileSync(jsonl, 'utf8');
  await http('_s_admin', 'GET', '/api/admin/logs?paged=1');
  ok(fs.readFileSync(jsonl, 'utf8').startsWith(before), '[A] 新事件只附加在尾端（既有內容未動）');

  // 匯出
  const exCsv = await http('_s_admin', 'GET', '/api/admin/logs/export?format=csv&kind=all');
  eq(exCsv.status, 200, '[A] 匯出 CSV 200');
  ok(/text\/csv/.test(exCsv.headers.get('content-type')) && /attachment; filename="audit-log-/.test(exCsv.headers.get('content-disposition')), '[A] CSV 標頭與檔名');
  const csv = exCsv.text;
  ok(csv.startsWith('﻿時間,類型,動作'), '[A] CSV 含 BOM 與表頭');
  ok(csv.includes(`'=HYPERLINK`) && csv.includes(`'+cmd`) && csv.includes(`'-cmd`) && csv.includes(`'@sum`), '[A] CSV 公式注入：= + - @ 開頭的儲存格都加單引號');
  ok(!/(^|,)=HYPERLINK/m.test(csv) && !/(^|,)[+@]cmd/m.test(csv), '[A] CSV 沒有任何以 = + @ 直接開頭的儲存格');
  ok(csv.includes('LEGACY_ACT'), '[A] CSV 含舊資料');
  const exJson = await http('_s_admin', 'GET', '/api/admin/logs/export?format=json&kind=all');
  ok(Array.isArray(exJson.json) && exJson.json.length > 8, '[A] 匯出 JSON 為合法陣列 (' + (exJson.json && exJson.json.length) + ')');
  const logsAfter = await getLogs();
  const exLog = byAction(logsAfter, 'EXPORT_AUDIT_LOG');
  eq(exLog.length, 2, '[A] 兩次匯出各寫一筆 EXPORT_AUDIT_LOG');
  ok(exLog[0] && /匯出 \d+ 筆/.test(exLog[0].detail) && /格式：json/.test(exLog[0].detail) && exLog[0].operator === '_s_admin', '[A] 稽核內容含筆數／格式／操作者');
  const exRange = await http('_s_admin', 'GET', '/api/admin/logs/export?format=json&from=2020-01-01&to=2020-01-31');
  eq(exRange.json.length, 2, '[A] 匯出支援日期範圍（舊資料 2 筆）');
  // 權限
  eq((await http('_s_u1', 'GET', '/api/admin/logs/export')).status, 403, '[A] 非管理員匯出 403');
  eq((await http('_s_u1', 'GET', '/api/admin/logs?paged=1')).status, 403, '[A] 非管理員讀日誌 403');
  eq((await http('_s_u1', 'GET', '/api/admin/contact-audit')).status, 403, '[A] 非管理員讀名片稽核 403');
  // 失敗計數曝光
  const inf = await http('_s_admin', 'GET', '/api/admin/infra-stats');
  ok(inf.json && inf.json.auditWrites && inf.json.auditWrites.failed === 0 && inf.json.auditWrites.inserted > 0, '[A] infra-stats 曝光 auditWrites（failed=0）');
  const ver = await http('_s_admin', 'GET', '/api/admin/version');
  ok(ver.json && ver.json.auditWrites && typeof ver.json.auditWrites.failed === 'number', '[A] version 曝光 auditWrites');
  // 管理後台頁面含新 UI
  const html = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
  ok(/id="logPager"/.test(html) && /exportLogs\('csv'\)/.test(html) && /logFilterQ/.test(html), '[A] admin.html：分頁列／篩選／匯出按鈕');
  ok(!/const logs = await adminFetch\(`\$\{API\}\/admin\/contact-audit/.test(html) && !/const list = await adminFetch\(`\$\{API\}\/admin\/deleted-contacts/.test(html), '[A] admin.html：不再把 Response 當陣列用');
}

// ═════════════ 其他項目（後續 commit 補上）═════════════
const EXTRA = require('./_aw2-sections.js');

// ═════════════ M：變異檢查 ═════════════
function loadMutated(file, from, to, tag) {
  let src = fs.readFileSync(file, 'utf8');
  if (!src.includes(from)) throw new Error('變異樣式找不到：' + tag);
  src = src.replace(from, () => to);
  const out = path.join(os.tmpdir(), `aw2-mut-${process.pid}-${tag}.js`);
  fs.writeFileSync(out, src);
  return out;
}
async function secM() {
  console.log('[M] 變異檢查');
  const storeFile = path.join(ROOT, 'lib', 'auditStore.js');
  const muts = [
    ['csv', "if (/^[=+\\-@\\t\\r]/.test(s)) s = \"'\" + s;", '', 'CSV 不加公式防護'],
    ['retry', 'attempt < 2', 'attempt < 1', '寫入失敗不重試'],
    ['offset', 'if (i++ < offset) continue;', 'if (false) continue;', '分頁 offset 失效'],
    ['sort', ".sort((a, b) => Date.parse(b.ts) - Date.parse(a.ts));\n    const it", ".sort(() => 0);\n    const it", '舊資料不依時間合併'],
    ['action', "if (f.action) add('action = $#', f.action);", '', 'SQL 動作篩選失效'],
    ['like', "function likeEscape(s) { return s.replace(/[\\\\%_]/g, (m) => '\\\\' + m); }", 'function likeEscape(s) { return s; }', '文字搜尋不跳脫 % _'],
    ['legacytotal', 'return (await adapter.count(f)) + legacy;', 'return (await adapter.count(f));', 'total 漏算舊資料'],
    ['nul', "s = s.replace(/\\u0000/g, '');", '', '不剝除 NUL'],
  ];
  for (const [tag, from, to, desc] of muts) {
    const file = loadMutated(storeFile, from, to, tag);
    try {
      muting = true; mutFail = 0;
      try { await unitStore(require(file), 'mut-' + tag); } catch (e) { mutFail++; /* 變異造成例外也算抓到 */ }
      muting = false;
      ok(mutFail > 0, `變異 ${tag}（${desc}）→ 單元測試必須失敗，實際失敗 ${mutFail} 項`);
    } finally { muting = false; try { fs.unlinkSync(file); } catch { /* 已刪 */ } }
  }
  await EXTRA.mutations({ ok, eq, ROOT, fs, path, os, loadMutated, setMuting: (v) => { muting = v; mutFail = 0; return () => mutFail; } });

  // 伺服器層變異：複製 server.js 改壞一條規則，起一台變異伺服器，用 A 項的同一組斷言檢查（必須出現「不對勁」）
  const tmpFiles = [];
  const mkMut = (name, from, to) => {
    const srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
    if (!srv.includes(from)) throw new Error('變異樣式找不到：' + name);
    const f = path.join(ROOT, 'server.__mut_' + name + '.js');
    fs.writeFileSync(f, srv.replace(from, () => to)); tmpFiles.push(f);
    return f;
  };
  const fresh = async (file) => { await stopServer(); cleanupFiles(); seedExtra = null; seedFiles(); await restart(file); };
  try {
    {   // 匯出不寫稽核
      await fresh(mkMut('xa', "auditExport(req, 'EXPORT_AUDIT_LOG', n,", "void (req, 'EXPORT_AUDIT_LOG', n,"));
      await loginAll(['_s_admin']);
      await http('_s_admin', 'GET', '/api/admin/logs/export?format=json');
      eq(byAction(await getLogs(), 'EXPORT_AUDIT_LOG').length, 0, '變異（匯出稽核日誌不留痕）→ 沒有 EXPORT_AUDIT_LOG：A 項「兩次匯出各寫一筆」斷言會失敗（抓得到）');
    }
    {   // 名片稽核不寫
      await fresh(mkMut('ca', 'function writeContactAudit(action, req, target, changes) {', 'function writeContactAudit(action, req, target, changes) {\n  return;'));
      await loginAll(['_s_admin', '_s_u1']);
      await http('_s_u1', 'POST', '/api/contacts', { name: '變異乙', company: '沙盒公司甲', taxId: '12345675', bu: 'ERP', email: 'm@sandbox.test' });
      const ca = await http('_s_admin', 'GET', '/api/admin/contact-audit?paged=1');
      eq(ca.json.total, 0, '變異（名片欄位稽核不寫）→ contact-audit 為空：A 項「新增／修改／刪除都進稽核」斷言會失敗（抓得到）');
    }
    {   // 分頁 offset 失效
      await fresh(mkMut('pg', "const r = await db.auditStore.list(filters, { limit: q.limit || 200, offset: q.offset || 0 });", "const r = await db.auditStore.list(filters, { limit: q.limit || 200, offset: 0 });"));
      await loginAll(['_s_admin']);
      for (let i = 0; i < 8; i++) await http(null, 'POST', '/api/login', { username: 'mut_user_' + i, password: 'x' });
      const p1 = await getLogs('?paged=1&limit=3&offset=0'), p2 = await getLogs('?paged=1&limit=3&offset=3');
      ok(p2.rows.some(x => p1.rows.some(y => y.id === x.id)), '變異（分頁 offset 失效）→ 第二頁與第一頁重複：A 項「兩頁不重複」斷言會失敗（抓得到）');
    }
    {   // 日誌 API 不擋非管理員
      await fresh(mkMut('adm', "app.get('/api/admin/logs/export', requireAdmin,", "app.get('/api/admin/logs/export', (req, res, next) => next(),"));
      await loginAll(['_s_u1']);
      const r = await http('_s_u1', 'GET', '/api/admin/logs/export?format=json');
      ok(r.status !== 403, '變異（匯出端點不擋非管理員）→ 一般帳號拿到 ' + r.status + '：A 項「非管理員匯出 403」斷言會失敗（抓得到）');
    }
  } finally {
    await stopServer();
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } });
  }
}

// ═════════════ 主程式 ═════════════
const ctx = () => ({ ok, eq, sleep, http, login, loginAll, getLogs, byAction, arr, jar: () => jar, setSeed: (fn) => { seedExtra = fn; }, mkUser, nowIso, PW, ROOT, fs, path, os, crypto, sha, restartServer: restart, FILES, BASE });
const ORDER = ['A', 'B', 'C', 'D', 'E', 'M'];
(async () => {
  const only = (process.env.AW2_ONLY || '').split(',').map(s => s.trim()).filter(Boolean);
  const SECTIONS = Object.assign({ A: secA, M: secM }, EXTRA.sections(ctx()));
  const seedFor = { A: (f) => {
    f.legacyLog = [
      { id: 'LEG-OLD-1', timestamp: '2020-01-02T03:04:05.000Z', action: 'LEGACY_ACT', operator: 'legacy_op', target: 'x', detail: '舊紀錄一', ip: '9.9.9.9', userAgent: 'old-ua' },
      { id: 'LEG-OLD-2', timestamp: '2020-01-20T03:04:05.000Z', action: 'LEGACY_ACT', operator: 'other_op', target: 'y', detail: '舊紀錄二' },
    ];
    f.legacyContact = [{ id: 'LEG-C-1', timestamp: '2020-02-01T00:00:00.000Z', userId: 'old_u', userName: '舊使用者', action: 'UPDATE', targetId: 'oc', targetName: '舊名片', targetCompany: '舊公司', changes: [{ field: 'name', fieldLabel: '姓名', oldValue: 'a', newValue: 'b' }] }];
  } };
  try {
    for (const k of ORDER) {
      if (!SECTIONS[k] || (k === 'M' && process.env.AW2_SKIP_MUT)) continue;
      if (only.length && !only.includes(k)) continue;
      if (k === 'M' || (EXTRA.noServer || []).includes(k)) {   // 不需要伺服器的項目（假 app＋記憶體資料庫）
        try { await SECTIONS[k](); } catch (e) { fail++; failures.push(`第 ${k} 項例外：${e && e.stack || e}`); console.log('  ✗ 例外', e && e.stack || e); }
        continue;
      }
      await stopServer(); cleanupFiles();
      seedExtra = (EXTRA.seed && EXTRA.seed[k]) || seedFor[k] || null;
      seedFiles(); await restart();
      try { await SECTIONS[k](); } catch (e) { fail++; failures.push(`第 ${k} 項例外：${e && e.stack || e}`); console.log('  ✗ 例外', e && e.stack || e); }
    }
  } finally {
    await stopServer();
    cleanupFiles();
    try { fs.unlinkSync(path.join(os.tmpdir(), 'aw2-fallback-' + process.pid + '.jsonl')); } catch { /* 無 */ }
  }
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); process.exit(1); }
})().catch(async e => { console.error(e); await stopServer(); cleanupFiles(); process.exit(1); });
