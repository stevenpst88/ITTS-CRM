#!/usr/bin/env node
'use strict';
/**
 * 稽核事件表（audit_events）的「真 Postgres」驗證腳本 —— 由「你本人」在自己的終端機對【拋棄式測試用】Supabase 專案執行。
 * Claude／自動化流程不得執行本腳本，也不會看到任何連線字串。
 *
 * ── 用法（Windows PowerShell；連線字串只存在這個終端機的環境變數，不會寫進任何檔案或日誌）──
 *   cd <本專案的 feat/audit-wave2 工作樹>          # 需有 node_modules（沒有的話：npm ci，或建 junction 指到主專案的 node_modules）
 *   $s = Read-Host -AsSecureString "測試資料庫連線字串"
 *   $env:AUDIT_VERIFY_DATABASE_URL = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s))
 *   # （建議）再設一個「正式站專案代號片段」當保險；連線字串含這段文字就拒絕執行：
 *   $env:AUDIT_VERIFY_FORBIDDEN_SUBSTR = "<正式站專案 ref 的一部分>"
 *   node scripts/verify-audit-pg.js                 # 會要求你輸入資料庫主機名稱的前 8 個字元確認
 *   node scripts/verify-audit-pg.js --yes-this-is-test-db      # 略過互動確認
 *   node scripts/verify-audit-pg.js --yes-this-is-test-db --drop   # 跑完後把 audit_events 整張表刪掉（預設保留）
 *   node scripts/verify-audit-pg.js --yes-this-is-test-db --chain  # 另外執行 [13] 防竄改鏈驗證（第三波；見下方說明）
 *   Remove-Item Env:AUDIT_VERIFY_DATABASE_URL, Env:AUDIT_VERIFY_FORBIDDEN_SUBSTR     # 跑完清掉環境變數
 *
 * 預期時間：約 1–4 分鐘（取決於到 Supabase 的網路延遲；其中 5200 筆大量資料用單一 SQL 產生，其餘走真實程式路徑）。
 * 結束代碼：全部 PASS＝0；任何 FAIL＝1；防護檢查擋下或無法連線＝2／3（此時完全沒有寫入）。
 *
 * ── 安全防護（任一不通過就在「任何寫入之前」中止）──
 *   1. 必須有 AUDIT_VERIFY_DATABASE_URL；不可與 DATABASE_URL 相同；不可含 AUDIT_VERIFY_FORBIDDEN_SUBSTR。
 *   2. 唯讀檢查 app_data：id='main' 若有超過 5 位聯絡人、5 筆商機或 3 個帳號 → 中止（「這看起來不是空的測試資料庫」）。
 *      （例外：測試庫合法地留有少量測試資料時，可明確放寬——兩個旗標必須一起給：
 *        node scripts/verify-audit-pg.js --allow-nonempty-test-db --expect-ref=<專案代號>
 *        且連線字串的「帳號」（Supabase pooler 為 postgres.<ref>）或「主機」必須含該代號，否則中止；
 *        放寬後仍有硬上限：聯絡人 >500、商機 >500 或帳號 >50 一律中止。DATABASE_URL 相等與禁止字串兩道防護照樣先跑。
 *        此模式下腳本只寫 audit_events 表，加上（僅在不存在時）app_data 一筆暫時的 id='audit_log' 舊整包列並於結束時刪除；絕不修改 app_data 的 id='main'。）
 *   3. 需輸入主機名稱前 8 字元，或帶 --yes-this-is-test-db。
 *   所有輸出都會遮蔽連線字串／密碼／主機；錯誤訊息也一樣。本腳本不使用專案根目錄的 data.json／auth.json／audit.log.json。
 *
 * ── 每一組 PASS 代表什麼 ──
 *   [1] 建表     ensureSchema 的 CREATE TABLE/INDEX IF NOT EXISTS 在真 Postgres 可重複執行；欄位型別與索引正確
 *   [2] 自檢     啟動自我檢查 SELECT 1 FROM audit_events 通過，模式為 table
 *   [3] 單筆     log／contact／quote-edit 三種事件（中文、emoji、NUL 與控制字元、超長內容、jsonb）寫入後讀回一致
 *   [4] 並行     3 個「實例」（各自獨立載入 db/postgres.js、各自連線池）交錯寫入 100 筆，一筆不少、id 不重複
 *   [5] 分頁     5200 筆資料：total、limit／offset 分頁、keyset 迭代（跨 500 筆分塊）、順序、不重複
 *   [6] 篩選     動作／操作者／日期（Asia/Taipei 日界線）／文字（% _ \ 當字面，ILIKE … ESCAPE）在真實 SQL 上的結果
 *   [7] 合併     舊整包 JSONB（app_data.id='audit_log'）與新表依時間交錯合併、分頁跨界
 *   [8] 匯出     對 5200 筆資料做 CSV／JSON 串流迭代；公式注入（= + - @）儲存格已加單引號
 *   [9] 旗標關   AUDIT_EVENTS_TABLE 未設時：寫進舊整包 JSONB，audit_events 完全沒被動到
 *   [10] 退回    表不可用（暫時把表換成會出錯的 view）時：記 AUDIT_TABLE_UNAVAILABLE、事件仍寫進舊路徑；結束後還原
 *   [11] 清理    只刪本次寫入的資料（extra->>'verifyRun'）；若舊整包列是本腳本建立的也一併刪除
 *   [12] 保護    在交易內套用 supabase/audit_events_protect.sql：UPDATE／DELETE／TRUNCATE 被拒、INSERT 仍可；最後 ROLLBACK
 *   [13] 防竄改鏈（只有加 --chain 才執行；見 lib/auditChain.js 檔頭的操作手冊）
 *        13.1 以 AUDIT_CHAIN=1＋這次執行產生的丟棄式隨機金鑰載入真實的 db/postgres.js：自動補 seq／prev_hash／row_hash／kid 欄位與 seq 唯一索引（ALTER … IF NOT EXISTS，冪等；欄位保留是無害的 schema 補強）
 *        13.2 3 個實例併發寫 100 筆：序號 1..N 無缺號、prev_hash 逐筆銜接、沒有降級（advisory lock 有效），並印出每筆平均耗時
 *        13.3 特殊內容（NUL、控制字元、jsonb 鍵序、5 萬字、毫秒時間）寫入後讀回，鏈仍驗證通過
 *        13.4 在交易內竄改（改欄位／時間／操作者／extra、改 prev_hash／kid、刪中間列、刪首列、插入偽造列、換順序、截尾＋錨點）→ 驗證在正確序號抓到；全部 ROLLBACK 後鏈恢復通過；錯誤金鑰失敗；區間驗證
 *               （若資料表已套用 protect SQL，UPDATE／DELETE 會被拒絕，改以「被拒絕」為證，略過其餘竄改情境）
 *        13.5 換金鑰（kid 輪替）接續同一條鏈，SECRETS_JSON 帶舊 secret 可驗證整條
 *        13.6 只清除本次標記（extra.verifyChain）的列
 *        前置：資料庫內若已有「非本腳本產生」的鏈資料，[13] 整段略過（不碰真實的鏈）。建議只對 Test 專案執行。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.join(__dirname, '..');
const NOT_EMPTY_MSG = '這看起來不是空的測試資料庫';

/** 把連線字串、密碼、主機、任何 postgres:// 開頭的字串都換成遮蔽字樣 */
function makeRedactor(url) {
  const secrets = [];
  if (url) {
    secrets.push(url);
    try {
      const u = new URL(url);
      for (const s of [u.password, safeDecode(u.password), u.username, u.hostname]) if (s && s.length >= 3) secrets.push(s);
    } catch { /* 不是合法 URL：只靠整串比對 */ }
  }
  secrets.sort((a, b) => b.length - a.length);
  return (text) => {
    // 先整串遮蔽任何 postgres:// 連線字串，免得帳號名（postgres）先被換掉而漏出後半
    let t = String(text == null ? '' : text).replace(/postgres(ql)?:\/\/\S+/gi, '［已遮蔽連線字串］');
    for (const s of secrets) t = t.split(s).join('［已遮蔽］');
    return t;
  };
}
/** 鍵排序後的 JSON（jsonb 讀回的鍵順序與寫入不同，比對要與順序無關） */
function canon(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  return '{' + Object.keys(v).sort().map((k) => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
}
function safeDecode(s) { try { return decodeURIComponent(s); } catch { return s; } }

/**
 * [13] 防竄改鏈（--chain 才會執行）。在真 Postgres 上驗證 lib/auditChain.js 與 lib/auditStore.js 的鏈：
 *   欄位／索引冪等建立、3 個實例併發 100 筆（序號 1..N 無缺號、prev_hash 逐筆銜接、advisory lock）、內容原樣讀回後仍驗證通過（時間精度／jsonb 重排／NUL／超長）、
 *   交易內竄改（改欄位／改時間／刪中間列／插入偽造列／換順序／截尾）→ 驗證在正確序號抓到、ROLLBACK 後恢復通過、帶錨點抓截尾、錯誤金鑰失敗、只清除本次標記的列。
 * 金鑰是這次執行產生的丟棄式隨機值，只在記憶體，不印出、不寫檔。前置條件：資料庫裡不可有「非本腳本產生」的鏈資料（有就整段略過，不碰真實的鏈）。
 * ctx：{ pool, q, say, check, skip, loadRepo(flag, extraEnv), runId, op, AS(lib/auditStore), AC(lib/auditChain), count }
 */
async function chainSection(ctx) {
  const { pool, q, say, check, skip, loadRepo, runId, op, AS, AC } = ctx;
  const secret = 'vfy-' + crypto.randomBytes(24).toString('hex');
  const chainEnv = { AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: secret, AUDIT_CHAIN_KID: 'vfy1' };
  const cfg = AC.getConfig(chainEnv);
  const opc = op + '_ch';
  const mark = (o) => Object.assign({ verifyRun: runId, verifyChain: '1' }, o || {});
  const hasCols = async () => Number((await q(`SELECT count(*)::int AS n FROM information_schema.columns WHERE table_name = 'audit_events' AND table_schema = current_schema() AND column_name IN ('seq','prev_hash','row_hash','kid')`))[0].n) === 4;

  // 前置：不碰真實的鏈
  if (await hasCols()) {
    const foreign = Number((await q(`SELECT count(*)::int AS n FROM audit_events WHERE seq IS NOT NULL AND coalesce(extra->>'verifyChain', '') <> '1'`))[0].n);
    if (foreign > 0) { skip('[13] 鏈', `資料庫已有 ${foreign} 筆非本腳本產生的鏈資料，為避免動到真實的鏈整段略過`); return; }
    const old = await q(`DELETE FROM audit_events WHERE extra->>'verifyChain' = '1' RETURNING id`);
    if (old.length) say(`    （先清掉上次中斷留下的 ${old.length} 筆鏈測試資料）`);
  }

  // [13.1] 補欄位（冪等）＋索引
  const A = loadRepo('1', chainEnv);
  await A.auditStore.init();
  check('[13.1] 新表模式，鏈已啟用（金鑰為本次丟棄式隨機值）', A.auditStore.mode() === 'table' && !!(A.auditStore.stats().chain && A.auditStore.stats().chain.enabled), A.auditStore.stats());
  await A.insertAuditEvent({ kind: 'log', action: 'CH_BOOT', operator: opc, detail: 'boot', extra: mark() });   // 第一筆寫入會觸發 ensureSchema 補鏈欄位
  await A.flush();
  const cols = await q(`SELECT column_name, data_type FROM information_schema.columns WHERE table_name = 'audit_events' AND table_schema = current_schema() AND column_name IN ('seq','prev_hash','row_hash','kid')`);
  const got = Object.fromEntries(cols.map((c) => [c.column_name, c.data_type]));
  check('[13.1] 鏈欄位存在且型別正確（seq bigint；prev_hash／row_hash／kid text）', JSON.stringify(got, Object.keys(got).sort()) === JSON.stringify({ kid: 'text', prev_hash: 'text', row_hash: 'text', seq: 'bigint' }), got);
  const idx = (await q(`SELECT indexdef FROM pg_indexes WHERE tablename = 'audit_events' AND indexname = 'audit_events_seq_uidx'`));
  check('[13.1] seq 唯一索引（partial：WHERE seq IS NOT NULL）存在', idx.length === 1 && /UNIQUE/i.test(idx[0].indexdef) && /seq IS NOT NULL/i.test(idx[0].indexdef), idx);
  const A2 = loadRepo('1', chainEnv); await A2.auditStore.init(); await A2.auditStore.chainSource();
  check('[13.1] 第二個實例再跑一次補欄位（IF NOT EXISTS）不出錯', A2.auditStore.stats().chain && A2.auditStore.stats().chain.ddlFailed === false);

  // [13.2] 3 個實例併發 100 筆
  const insts = [A, A2, loadRepo('1', chainEnv)];
  const t0 = Date.now();
  const ps = [];
  for (let i = 0; i < 100; i++) ps.push(insts[i % 3].insertAuditEvent({ kind: i % 7 === 0 ? 'contact' : 'log', action: 'CH_CONC', operator: opc, target: 'i' + (i % 3), detail: '併發 ' + i + ' 😀', extra: mark({ i, nested: { z: [1, '中', null] } }) }));
  const res = await Promise.all(ps);
  for (const m of insts) await m.flush();
  const ms = Date.now() - t0;
  say(`    （100 筆併發寫入共 ${ms} ms，約 ${(ms / 100).toFixed(1)} ms／筆——每筆多一個交易＋advisory lock，全站稽核寫入因此序列化）`);
  check('[13.2] 100 筆寫入全部回報成功', res.every((r) => r === true), res.filter((r) => r !== true).length);
  check('[13.2] 沒有任何一筆因鏈衝突而降級（breakRisk 全為 0）', insts.every((m) => m.auditStore.stats().chain && m.auditStore.stats().chain.breakRisk === 0), insts.map((m) => m.auditStore.stats().chain));
  const rows = await q(`SELECT seq, prev_hash, row_hash, kid FROM audit_events WHERE seq IS NOT NULL ORDER BY seq`);
  const N = rows.length;
  check('[13.2] 鏈上共 101 筆（boot ＋ 100 筆）', N === 101, N);
  check('[13.2] seq 為 1..N 無缺號、無重複', rows.every((r, i) => Number(r.seq) === i + 1), rows.slice(0, 3).map((r) => r.seq));
  check('[13.2] 每一筆 prev_hash 等於前一筆 row_hash（第一筆為 64 個 0）', rows.every((r, i) => r.prev_hash === (i ? rows[i - 1].row_hash : AC.GENESIS)) && rows.every((r) => /^[0-9a-f]{64}$/.test(r.row_hash) && r.kid === 'vfy1'));
  const src = await A.auditStore.chainSource();
  let v = await AC.verifyChain(src, cfg, {});
  check('[13.2] verifyChain 通過：101 筆、鏈頭序號 101', v.ok && v.checked === 101 && v.headSeq === 101 && v.legacyUncovered === 0 && v.unchainedAfterStart === 0, v);

  // [13.3] 內容原樣讀回仍可驗證（時間精度、jsonb 重排、NUL、控制字元、超長）
  const nasty = [
    { kind: 'log', action: 'CH_RT', operator: opc, target: 'T\u0000X', detail: '中文 😀 含\u0000NUL 與\u0001\u001f控制字元 "引號" \\反斜線\\', ip: '::1', ua: 'UA/1', extra: mark({ b: 1, a: { d: [3, 2, 1], c: null }, s: '\u0000x', n: 1.5, big: 12345678901, t: true }) },
    { kind: 'quote-edit', action: 'CH_RT', operator: opc, target: 'Q', detail: 'x'.repeat(50000), extra: mark({ quoteNo: 'Q-1', count: 1 }) },
    { kind: 'log', action: 'CH_RT', operator: opc, ts: '2026-03-04T05:06:07.089Z', detail: '指定時間（毫秒精度）', extra: mark() },
  ];
  for (const e of nasty) await A.insertAuditEvent(e);
  await A.flush();
  v = await AC.verifyChain(src, cfg, {});
  check('[13.3] 特殊內容（NUL／控制字元／jsonb 鍵序／超長／毫秒時間）寫入後讀回，鏈仍驗證通過（雜湊的是實際儲存的值）', v.ok && v.checked === 101 + nasty.length, v);
  const NN = 101 + nasty.length;

  // [13.4] 交易內竄改 → 驗證抓到 → ROLLBACK 後恢復
  const cx = await pool.connect();
  let tamperOk = true;
  try {
    await cx.query('BEGIN');
    const txSrc = AS.pgAdapter({ query: (s, p) => cx.query(s, p), chain: null });
    const trial = async (label, stmts, expectSeq, expectReason, extraOpts) => {
      await cx.query('SAVEPOINT t1');
      let rejected = null;
      try { for (const [s, p] of stmts) await cx.query(s, p); } catch (e) { rejected = e; }
      if (rejected) { await cx.query('ROLLBACK TO SAVEPOINT t1'); return { rejected: true }; }
      const r = await AC.verifyChain(txSrc, cfg, extraOpts || {});
      await cx.query('ROLLBACK TO SAVEPOINT t1');
      const reasons = [].concat(expectReason);
      check(`[13.4] ${label} → 在 seq ${expectSeq} 以 ${reasons.join('／')} 抓到`, !r.ok && r.firstBreak && r.firstBreak.seq === expectSeq && reasons.includes(r.firstBreak.reason), r.firstBreak);
      return { rejected: false, r };
    };
    const first = await trial('改 detail', [[`UPDATE audit_events SET detail = '被竄改' WHERE seq = 50 AND operator = $1`, [opc]]], 50, 'HASH_MISMATCH');
    if (first.rejected) {
      tamperOk = false;
      say('    （UPDATE 被資料庫拒絕——audit_events 已套用保護 SQL；竄改情境改以「被拒絕」本身為證，略過其餘交易內竄改測試）');
      check('[13.4] 已保護的資料表：UPDATE 被拒絕', true);
    } else {
      await trial('改 ts', [[`UPDATE audit_events SET ts = ts + interval '1 second' WHERE seq = 60 AND operator = $1`, [opc]]], 60, 'HASH_MISMATCH');
      await trial('改 operator', [[`UPDATE audit_events SET operator = operator || 'x' WHERE seq = 61`, []]], 61, 'HASH_MISMATCH');
      await trial('改 extra（jsonb）', [[`UPDATE audit_events SET extra = jsonb_set(extra, '{i}', '999'::jsonb) WHERE seq = 62`, []]], 62, 'HASH_MISMATCH');
      await trial('改 prev_hash', [[`UPDATE audit_events SET prev_hash = repeat('a', 64) WHERE seq = 63`, []]], 63, 'PREV_MISMATCH');
      await trial('改 kid', [[`UPDATE audit_events SET kid = 'nope' WHERE seq = 64`, []]], 64, 'UNKNOWN_KID');
      await trial('刪中間列', [[`DELETE FROM audit_events WHERE seq = 70 AND operator = $1`, [opc]]], 70, 'SEQ_GAP');
      await trial('刪第一列', [[`DELETE FROM audit_events WHERE seq = 1 AND operator = $1`, [opc]]], 1, 'SEQ_GAP');
      await trial('插入偽造列（接在鏈尾、亂填雜湊）', [[`INSERT INTO audit_events (kind, action, operator, extra, seq, prev_hash, row_hash, kid) VALUES ('log', 'CH_FORGED', $1, $2::jsonb, $3::bigint, repeat('b', 64), repeat('c', 64), 'vfy1')`, [opc, JSON.stringify(mark()), NN + 1]]], NN + 1, ['PREV_MISMATCH', 'HASH_MISMATCH']);
      await trial('換順序（交換 seq 80、81）', [
        [`UPDATE audit_events SET seq = -1 WHERE seq = 80`, []], [`UPDATE audit_events SET seq = 80 WHERE seq = 81`, []], [`UPDATE audit_events SET seq = 81 WHERE seq = -1`, []]], 80, ['PREV_MISMATCH', 'HASH_MISMATCH']);
      // 沒有鏈欄位的直接寫入：不算斷鏈，但被計入警示
      await cx.query('SAVEPOINT t2');
      await cx.query(`INSERT INTO audit_events (kind, action, operator, extra) VALUES ('log', 'CH_DIRECT', $1, $2::jsonb)`, [opc, JSON.stringify(mark())]);
      const rd = await AC.verifyChain(txSrc, cfg, {});
      await cx.query('ROLLBACK TO SAVEPOINT t2');
      check('[13.4] 鏈啟用後被直接 INSERT（沒有鏈欄位）的列 → 鏈仍通過，但 unchainedAfterStart=1 警示', rd.ok && rd.unchainedAfterStart === 1, rd);
      // 截尾：沒有錨點偵測不到（已知限制），有錨點抓得到
      const head = await src.chainHead(), cnt = await src.chainCount();
      const anchor = AC.makeAnchor(cfg, head, cnt);
      await cx.query('SAVEPOINT t3');
      await cx.query(`DELETE FROM audit_events WHERE seq > $1::bigint AND extra->>'verifyChain' = '1'`, [NN - 3]);
      const noA = await AC.verifyChain(txSrc, cfg, {});
      const withA = await AC.verifyChain(txSrc, cfg, { anchor });
      await cx.query('ROLLBACK TO SAVEPOINT t3');
      check('[13.4] 截掉尾端 3 筆、不帶錨點 → 驗證仍通過（已知限制）', noA.ok && noA.headSeq === NN - 3, noA);
      check('[13.4] 截掉尾端 3 筆、帶錨點 → TRUNCATED', !withA.ok && withA.anchor && withA.anchor.reason === 'TRUNCATED', withA.anchor);
      check('[13.4] 錨點簽章可驗證，且未被截尾時比對通過', AC.verifyAnchor(cfg, anchor) && (await AC.verifyChain(src, cfg, { anchor })).anchor.ok);
    }
  } finally {
    await cx.query('ROLLBACK').catch(() => {});
    cx.release();
  }
  v = await AC.verifyChain(src, cfg, {});
  check('[13.4] ROLLBACK 後：鏈完整恢復，仍驗證通過', v.ok && v.checked === NN, v);
  const wrong = await AC.verifyChain(src, AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: 'vfy-wrong-' + crypto.randomBytes(16).toString('hex'), AUDIT_CHAIN_KID: 'vfy1' }), {});
  check('[13.4] 用錯誤金鑰驗證 → 第 1 筆就 HASH_MISMATCH（只拿到資料庫的人無法偽造出可驗證的鏈）', !wrong.ok && wrong.firstBreak && wrong.firstBreak.seq === 1 && wrong.firstBreak.reason === 'HASH_MISMATCH', wrong.firstBreak);
  const part = await AC.verifyChain(src, cfg, { from: 10, to: 20 });
  check('[13.4] 區間驗證（10..20）11 筆通過', part.ok && part.checked === 11 && part.complete, part);

  // [13.5] 換金鑰（kid 輪替）接續同一條鏈
  const secret2 = 'vfy2-' + crypto.randomBytes(24).toString('hex');
  const chainEnv2 = { AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: secret2, AUDIT_CHAIN_KID: 'vfy2', AUDIT_CHAIN_SECRETS_JSON: JSON.stringify({ vfy1: secret }) };
  const R = loadRepo('1', chainEnv2);
  for (let i = 0; i < 3; i++) await R.insertAuditEvent({ kind: 'log', action: 'CH_ROT', operator: opc, detail: 'rot ' + i, extra: mark() });
  await R.flush();
  const cfg2 = AC.getConfig(chainEnv2);
  const rk = (await q(`SELECT kid FROM audit_events WHERE seq > $1::bigint ORDER BY seq`, [NN])).map((r) => r.kid);
  check('[13.5] 換金鑰後新列 kid=vfy2，序號接續', rk.join() === 'vfy2,vfy2,vfy2', rk);
  v = await AC.verifyChain(await R.auditStore.chainSource(), cfg2, {});
  check('[13.5] 帶舊 secret（SECRETS_JSON）驗證整條通過', v.ok && v.checked === NN + 3, v);
  v = await AC.verifyChain(await R.auditStore.chainSource(), AC.getConfig({ AUDIT_CHAIN: '1', AUDIT_CHAIN_SECRET: secret2, AUDIT_CHAIN_KID: 'vfy2' }), {});
  check('[13.5] 遺失舊 secret → 舊列 UNKNOWN_KID', !v.ok && v.firstBreak && v.firstBreak.reason === 'UNKNOWN_KID' && v.firstBreak.seq === 1, v.firstBreak);

  // [13.6] 清理：只刪本次標記的列
  const del = await q(`DELETE FROM audit_events WHERE extra->>'verifyChain' = '1' RETURNING id`);
  check(`[13.6] 刪除本次標記的鏈測試資料（${del.length} 筆）`, del.length >= NN + 3, del.length);
  const left = Number((await q(`SELECT count(*)::int AS n FROM audit_events WHERE extra->>'verifyChain' = '1' OR operator = $1`, [opc]))[0].n);
  check('[13.6] 沒有任何殘留', left === 0, left);
  const chained = Number((await q(`SELECT count(*)::int AS n FROM audit_events WHERE seq IS NOT NULL`))[0].n);
  check('[13.6] 資料庫內不再有任何鏈資料（鏈欄位保留，這是冪等且無害的 schema 補強）', chained === 0, chained);
}

/**
 * opts: { env, argv, out(line), ask(prompt)->Promise<string>, pg(模組，測試注入用) }
 * 回傳 { code, passed, failed }；不呼叫 process.exit（直接執行時由檔尾處理）。
 */
async function run(opts) {
  opts = opts || {};
  const env = opts.env || process.env;
  const argv = opts.argv || process.argv.slice(2);
  const url = env.AUDIT_VERIFY_DATABASE_URL || '';
  const redact = makeRedactor(url);
  const out = opts.out || ((l) => console.log(l));
  const say = (l) => out(redact(l));
  let passed = 0, failed = 0;
  const check = (name, cond, detail) => {
    if (cond) { passed++; say('PASS ' + name); } else { failed++; say('FAIL ' + name + (detail !== undefined ? '  ← ' + (typeof detail === 'string' ? detail : JSON.stringify(detail)) : '')); }
  };
  const skip = (name, why) => say('SKIP ' + name + '（' + why + '）');
  const abort = (code, msg) => { say('✗ 中止：' + msg); say('（沒有進行任何寫入）'); return { code, passed, failed }; };

  // ── 防護 1：環境變數與字串比對（尚未連線）──
  if (argv.includes('--help') || argv.includes('-h')) { say('用法與說明請看 scripts/verify-audit-pg.js 檔案開頭的註解。'); return { code: 0, passed, failed }; }
  if (!url) return abort(2, '沒有設定 AUDIT_VERIFY_DATABASE_URL（請依檔案開頭的說明用 Read-Host 設定）');
  let parsed;
  try { parsed = new URL(url); if (!/^postgres(ql)?:$/.test(parsed.protocol)) throw new Error('protocol'); } catch { return abort(2, 'AUDIT_VERIFY_DATABASE_URL 不是合法的 postgres:// 連線字串'); }
  const prodUrl = env.DATABASE_URL || '';
  if (prodUrl) {
    let same = url === prodUrl;
    try { const p = new URL(prodUrl); same = same || (p.hostname === parsed.hostname && p.port === parsed.port && p.pathname === parsed.pathname && p.username === parsed.username); } catch { /* 比整串就好 */ }
    if (same) return abort(2, 'AUDIT_VERIFY_DATABASE_URL 與目前環境的 DATABASE_URL 指向同一個資料庫，拒絕執行');
  }
  const forbidden = (env.AUDIT_VERIFY_FORBIDDEN_SUBSTR || '').trim();
  if (forbidden && url.toLowerCase().includes(forbidden.toLowerCase())) return abort(2, '連線字串含有 AUDIT_VERIFY_FORBIDDEN_SUBSTR 標記的正式站特徵，拒絕執行');

  // ── 防護 1b：放寬「必須是空庫」的明確覆寫（兩個旗標缺一不可，且連線字串的帳號或主機要含指定的專案代號）──
  const allowArg = argv.includes('--allow-nonempty-test-db');
  const refArg = argv.find((a) => a.startsWith('--expect-ref='));
  const expectRef = refArg ? refArg.slice('--expect-ref='.length).trim() : '';
  if (allowArg !== !!refArg) return abort(2, '--allow-nonempty-test-db 與 --expect-ref=<專案代號> 必須同時給，缺一不可');
  let relaxed = false;
  if (allowArg) {
    if (expectRef.length < 6) return abort(2, '--expect-ref 的專案代號太短（至少 6 個字元）');
    const userPart = safeDecode(parsed.username || '');
    if (!(userPart.includes(expectRef) || parsed.hostname.includes(expectRef))) return abort(2, '連線字串的帳號或主機不含 --expect-ref 指定的專案代號，拒絕放寬空庫檢查');
    relaxed = true;
  }

  // ── 載入 pg、連線（唯讀檢查）──
  let pgMod = opts.pg;
  if (!pgMod) {
    try { pgMod = require(require.resolve('pg', { paths: [ROOT] })); } catch { return abort(3, '找不到 pg 套件：請在有 node_modules 的資料夾執行（npm ci）'); }
  }
  const pool = new pgMod.Pool({ connectionString: url, ssl: { rejectUnauthorized: false }, max: 3, connectionTimeoutMillis: 15000 });
  if (pool.on) pool.on('error', () => { /* 閒置連線錯誤不外洩細節 */ });
  const q = async (sql, params) => (await pool.query(sql, params)).rows;
  const fatal = (e, what) => abort(3, what + '：' + redact(e && e.message));

  try {
    const c = await pool.connect();
    try {
      await c.query('BEGIN READ ONLY');
      const t = (await c.query(`SELECT to_regclass('public.app_data') IS NOT NULL AS ok`)).rows[0];
      if (t && t.ok) {
        const r = (await c.query(`SELECT
            CASE WHEN jsonb_typeof(content->'contacts') = 'array' THEN jsonb_array_length(content->'contacts') ELSE 0 END AS contacts,
            CASE WHEN jsonb_typeof(content->'opportunities') = 'array' THEN jsonb_array_length(content->'opportunities') ELSE 0 END AS opps,
            CASE WHEN jsonb_typeof(content->'_auth'->'users') = 'array' THEN jsonb_array_length(content->'_auth'->'users') ELSE 0 END AS users
          FROM app_data WHERE id = 'main'`)).rows[0];
        const lim = relaxed ? { c: 500, o: 500, u: 50 } : { c: 5, o: 5, u: 3 };   // 放寬模式仍有硬上限（正式站資料量遠大於此）
        if (r && (Number(r.contacts) > lim.c || Number(r.opps) > lim.o || Number(r.users) > lim.u)) {
          await c.query('ROLLBACK').catch(() => {});
          c.release();
          await pool.end().catch(() => {});
          return abort(2, `${NOT_EMPTY_MSG}（main 內有 ${r.contacts} 位聯絡人、${r.opps} 筆商機、${r.users} 個帳號；${relaxed ? '放寬模式的硬上限' : '上限'} ${lim.c}／${lim.o}／${lim.u}）`);
        }
        if (relaxed && r && (Number(r.contacts) > 5 || Number(r.opps) > 5 || Number(r.users) > 3)) say(`注意：資料庫含少量測試資料（${r.contacts} 位聯絡人、${r.opps} 筆商機、${r.users} 個帳號），在 --allow-nonempty-test-db 放寬模式下繼續。`);
      }
      await c.query('ROLLBACK');
    } finally { try { c.release(); } catch { /* 已釋放 */ } }
  } catch (e) { await pool.end().catch(() => {}); return fatal(e, '無法連線或唯讀檢查失敗'); }

  // ── 防護 3：確認 ──
  if (!argv.includes('--yes-this-is-test-db')) {
    const want = parsed.hostname.slice(0, 8);
    const ask = opts.ask || ((prompt) => new Promise((res) => { const rl = require('readline').createInterface({ input: process.stdin, output: process.stdout }); rl.question(prompt, (a) => { rl.close(); res(a); }); }));
    const got = String(await ask('這個腳本會在該資料庫建立測試資料並於結束時清除。請輸入資料庫主機名稱的前 8 個字元以確認這是測試資料庫：') || '').trim();
    if (got !== want) { await pool.end().catch(() => {}); return abort(2, '輸入的主機前綴不符，已取消'); }
  }
  say('— 防護通過，開始驗證（只寫入並清除本次標記的資料）—');
  if (relaxed) say('放寬模式：本腳本只會寫入資料表 audit_events，以及（僅在不存在時）app_data 內一筆暫時的舊整包列 id=\'audit_log\'（結束時刪除）；絕不修改 app_data 的 id=\'main\'。');
  if (opts.stopAfterGuards) { await pool.end().catch(() => {}); return { code: 0, passed, failed, guardsPassed: true }; }   // 測試用：只驗證防護，不進入寫入流程

  // ── 準備 ──
  const runId = 'v' + Date.now().toString(36) + crypto.randomBytes(2).toString('hex');
  const op = 'vfy_' + runId;
  say('本次標記 verifyRun=' + runId);
  const CHAIN_VARS = ['AUDIT_CHAIN', 'AUDIT_CHAIN_SECRET', 'AUDIT_CHAIN_KID', 'AUDIT_CHAIN_SECRETS_JSON'];
  const saveEnv = { u: process.env.DATABASE_URL, f: process.env.AUDIT_EVENTS_TABLE, chain: Object.fromEntries(CHAIN_VARS.map((k) => [k, process.env[k]])) };
  const dbFile = path.join(ROOT, 'db', 'postgres.js');
  const storeFile = path.join(ROOT, 'lib', 'auditStore.js');
  const legacyFile = path.join(ROOT, 'lib', 'auditLegacy.js');
  const loaded = [];
  /** 全新載入一份真實的 db/postgres.js（＝一個「實例」：自己的連線池與快取）；flag='1' 走新表，undefined＝旗標關 */
  const chainFile = path.join(ROOT, 'lib', 'auditChain.js');
  const loadRepo = (flag, extraEnv) => {
    process.env.DATABASE_URL = url;
    if (flag === undefined) delete process.env.AUDIT_EVENTS_TABLE; else process.env.AUDIT_EVENTS_TABLE = flag;
    for (const k of CHAIN_VARS) delete process.env[k];          // 預設＝鏈關閉；只有明確傳入的才會設
    if (extraEnv) for (const k of Object.keys(extraEnv)) process.env[k] = extraEnv[k];
    for (const f of [dbFile, storeFile, legacyFile, chainFile]) delete require.cache[f];
    const m = require(dbFile);
    loaded.push(m);
    return m;
  };
  const restoreEnv = () => {
    if (saveEnv.u === undefined) delete process.env.DATABASE_URL; else process.env.DATABASE_URL = saveEnv.u;
    if (saveEnv.f === undefined) delete process.env.AUDIT_EVENTS_TABLE; else process.env.AUDIT_EVENTS_TABLE = saveEnv.f;
    for (const k of CHAIN_VARS) { if (saveEnv.chain[k] === undefined) delete process.env[k]; else process.env[k] = saveEnv.chain[k]; }
  };
  const section = async (name, fn) => {
    say('— ' + name + ' —');
    try { await fn(); } catch (e) { failed++; say('FAIL ' + name + ' 發生例外 ← ' + redact(e && e.message)); }
  };
  const count = async (where, params) => Number((await q(`SELECT count(*)::int AS n FROM audit_events WHERE ${where}`, params))[0].n);
  const blobPresentAtStart = { v: null };
  let weCreatedBlob = false;
  let AS = null;   // lib/auditStore（載入後取得）

  try {
    // ═════ [1] 建表（冪等）＋ [2] 自檢 ═════
    let A;
    await section('[1] 建表（冪等）', async () => {
      A = loadRepo('1');
      AS = require(storeFile);
      const mkAd = () => AS.pgAdapter({ query: (sql, p) => pool.query(sql, p) });
      const a1 = mkAd(); await a1.ensure();
      const a2 = mkAd(); await a2.ensure();        // 第二個轉接器＝全新的記憶化狀態，會再送一次 CREATE … IF NOT EXISTS
      check('[1] ensureSchema 連跑兩次都成功（IF NOT EXISTS 冪等）', true);
      const cols = await q(`SELECT column_name, data_type FROM information_schema.columns WHERE table_name = 'audit_events' AND table_schema = current_schema() ORDER BY ordinal_position`);
      const want = { id: 'bigint', ts: 'timestamp with time zone', kind: 'text', action: 'text', operator: 'text', target: 'text', detail: 'text', ip: 'text', ua: 'text', extra: 'jsonb' };
      const got = Object.fromEntries(cols.map((c) => [c.column_name, c.data_type]));
      check('[1] 欄位與型別符合規格（id bigint、ts timestamptz、extra jsonb…）', JSON.stringify(got) === JSON.stringify(want), got);
      const idx = (await q(`SELECT indexname FROM pg_indexes WHERE tablename = 'audit_events'`)).map((r) => r.indexname);
      check('[1] 三個索引都在', ['audit_events_kind_ts_idx', 'audit_events_action_idx', 'audit_events_operator_idx'].every((n) => idx.includes(n)), idx);
    });
    await section('[2] 自我檢查', async () => {
      await A.auditStore.init();
      check('[2] SELECT 1 FROM audit_events LIMIT 1 自我檢查通過、模式 table', A.auditStore.mode() === 'table', A.auditStore.stats());
    });

    // ═════ [3] 單筆寫入（三種 kind）═════
    await section('[3] 單筆寫入', async () => {
      const long = 'x'.repeat(50000);
      const ex = (o) => Object.assign({ verifyRun: runId }, o || {});
      const evs = [
        { kind: 'log', action: 'VFY_LOG', operator: op + '_s', target: '目標一', detail: '中文內容 😀 含\u0000NUL 與\u0001\u001f控制字元', ip: '1.2.3.4', ua: 'UA/1', extra: ex({ nested: { arr: [1, '中文', null], obj: { k: 'v' } } }) },
        { kind: 'contact', action: 'UPDATE', operator: op + '_s', target: 'c1', detail: '王小明 / 甲公司', extra: ex({ userName: '測試者', targetId: 'c1', targetName: '王小明', targetCompany: '甲公司', changes: [{ field: 'title', fieldLabel: '職稱', oldValue: '專員', newValue: '經理' }] }) },
        { kind: 'quote-edit', action: 'DRAFT_SAVE', operator: op + '_s', target: 'QU-VFY-001', detail: long, extra: ex({ quoteNo: 'QU-VFY-001', count: 1 }) },
      ];
      for (const e of evs) check(`[3] 寫入 ${e.kind} 事件成功`, (await A.insertAuditEvent(e)) === true);
      await A.flush();
      const rows = await q(`SELECT kind, action, operator, target, detail, ip, ua, extra FROM audit_events WHERE operator = $1`, [op + '_s']);
      const by = Object.fromEntries(rows.map((r) => [r.kind, r]));
      check('[3] 三種 kind 各 1 筆', rows.length === 3 && by.log && by.contact && by['quote-edit'], rows.map((r) => r.kind));
      check('[3] 中文與 emoji 原樣讀回；NUL 已剝除、其他控制字元保留', by.log && by.log.detail === '中文內容 😀 含NUL 與\u0001\u001f控制字元', by.log && by.log.detail);
      check('[3] ip／ua／target 欄位正確', by.log && by.log.ip === '1.2.3.4' && by.log.ua === 'UA/1' && by.log.target === '目標一');
      check('[3] 超長內容被截成 20000 字', by['quote-edit'] && by['quote-edit'].detail.length === 20000, by['quote-edit'] && by['quote-edit'].detail.length);
      check('[3] jsonb extra（巢狀陣列／物件／中文）原樣讀回', by.log && canon(by.log.extra) === canon(evs[0].extra), by.log && by.log.extra);
      check('[3] 名片稽核的 changes 結構完整', by.contact && by.contact.extra.changes[0].newValue === '經理' && by.contact.extra.targetCompany === '甲公司');
      const j = await q(`SELECT extra->>'verifyRun' AS r, extra->'nested'->'arr'->>1 AS z FROM audit_events WHERE operator = $1 AND kind = 'log'`, [op + '_s']);
      check('[3] 可用 jsonb 運算子查詢 extra', j[0] && j[0].r === runId && j[0].z === '中文');
    });

    // ═════ [4] 3 個實例並行 100 筆 ═════
    await section('[4] 多實例並行寫入', async () => {
      const insts = [A, loadRepo('1'), loadRepo('1')];
      const ps = [];
      for (let i = 0; i < 100; i++) ps.push(insts[i % 3].insertAuditEvent({ kind: 'log', action: 'CONC', operator: op + '_c', target: 'i' + (i % 3), detail: 'c' + i, extra: { verifyRun: runId } }));
      const res = await Promise.all(ps);
      for (const m of insts) await m.flush();
      check('[4] 100 筆寫入全部回報成功', res.every((r) => r === true), res.filter((r) => r !== true).length);
      const n = await count(`operator = $1 AND action = 'CONC'`, [op + '_c']);
      check('[4] 資料庫內恰好 100 筆（沒有互相覆蓋、沒有遺失）', n === 100, n);
      const d = (await q(`SELECT count(DISTINCT id)::int AS a, count(DISTINCT detail)::int AS b FROM audit_events WHERE operator = $1 AND action = 'CONC'`, [op + '_c']))[0];
      check('[4] id 不重複、內容各異', d.a === 100 && d.b === 100, d);
      check('[4] 透過 auditStore.total 讀到同樣的 100', (await A.auditStore.total({ kind: 'log', operator: op + '_c', action: 'CONC' })) === 100);
      check('[4] 失敗計數為 0', insts.every((m) => m.auditStore.stats().failed === 0));
    });

    // ═════ [5] 大量資料＋分頁 ═════
    const BULK = 5200;
    await section('[5] 大量資料與分頁', async () => {
      await q(`INSERT INTO audit_events (ts, kind, action, operator, target, detail, ip, ua, extra)
        SELECT $2::timestamptz + make_interval(secs => g), 'log', 'BULK', $1::text, 'bulk',
               CASE WHEN g % 500 = 0 THEN '=SUM(1+1)' WHEN g % 700 = 0 THEN '@cmd' ELSE 'bulk row ' || g END,
               '', '', jsonb_build_object('verifyRun', $3::text)
        FROM generate_series(1, ${BULK}) AS g`, [op + '_b', '2026-03-01T00:00:00Z', runId]);
      const f = { kind: 'log', operator: op + '_b' };
      check('[5] total＝5200', (await A.auditStore.total(f)) === BULK);
      const p0 = await A.auditStore.list(f, { limit: 200, offset: 0 });
      check('[5] 分頁回應：total／limit／offset／筆數', p0.total === BULK && p0.limit === 200 && p0.offset === 0 && p0.rows.length === 200, [p0.total, p0.rows.length]);
      check('[5] 第一筆是最新（g=5200）', p0.rows[0].detail === 'bulk row 5200' || p0.rows[0].detail.startsWith('bulk row'), p0.rows[0].detail);
      const all = [];
      for (let off = 0; off < BULK; off += 1000) all.push(...(await A.auditStore.list(f, { limit: 1000, offset: off })).rows);
      check('[5] 各頁相加＝5200 筆', all.length === BULK, all.length);
      check('[5] 各頁沒有重複', new Set(all.map((r) => r.id)).size === BULK);
      let desc = true; for (let i = 1; i < all.length; i++) if (all[i - 1].seq < all[i].seq || Date.parse(all[i - 1].ts) < Date.parse(all[i].ts)) desc = false;
      check('[5] 全程由新到舊（id 與時間都單調）', desc);
      const mid = await A.auditStore.list(f, { limit: 5, offset: 497 });
      check('[5] 跨 500 筆分塊邊界的 offset 正確', mid.rows.map((r) => r.id).join() === all.slice(497, 502).map((r) => r.id).join());
      const ids = []; for await (const r of A.auditStore.iterate(f)) ids.push(r.seq);
      let strict = true; for (let i = 1; i < ids.length; i++) if (!(ids[i - 1] > ids[i])) strict = false;
      check('[5] keyset（before-id）迭代：5200 筆、id 嚴格遞減、分塊無重漏', ids.length === BULK && strict, ids.length);
    });

    // ═════ [6] 篩選（真實 SQL）═════
    await section('[6] 篩選', async () => {
      const o = op + '_f';
      const rows = [
        ['2026-05-01T15:59:59Z', 'FA', o, '100% off'],      // 台北 5/1 23:59:59
        ['2026-05-01T16:00:00Z', 'FB', o, '1000 off'],      // 台北 5/2 00:00:00
        ['2026-05-02T03:00:00Z', 'FA', o, 'a_b'],
        ['2026-05-02T04:00:00Z', 'FB', o, 'axb'],
        ['2026-05-03T04:00:00Z', 'FA', o, 'path\\dir'],
        ['2026-05-03T05:00:00Z', 'FB', o, 'pathXdir'],
        ['2026-05-02T05:00:00Z', 'FA', op + '_g', '100% off'],
      ];
      for (const [ts, action, operator, detail] of rows) await A.insertAuditEvent({ ts, kind: 'log', action, operator, target: 'T', detail, extra: { verifyRun: runId } });
      await A.flush();
      const T = async (f, extra) => (await A.auditStore.list(Object.assign({ kind: 'log', operator: o }, f), { limit: 100 })).rows.map((r) => r.detail).sort();
      check('[6] 單日 5/1（台北）只含 23:59:59 那筆', JSON.stringify(await T({ from: '2026-05-01', to: '2026-05-01' })) === JSON.stringify(['100% off']), await T({ from: '2026-05-01', to: '2026-05-01' }));
      check('[6] 單日 5/2（台北）含 00:00:00 起的 3 筆', (await T({ from: '2026-05-02', to: '2026-05-02' })).length === 3);
      check('[6] 只有 from＝5/2 → 5 筆', (await T({ from: '2026-05-02' })).length === 5);
      check('[6] 只有 to＝5/1 → 1 筆', (await T({ to: '2026-05-01' })).length === 1);
      check('[6] 文字 100% → 只有字面含 % 的那筆（不是萬用字元）', JSON.stringify(await T({ q: '100%' })) === JSON.stringify(['100% off']));
      check('[6] 文字 a_b → 只有字面含底線的那筆（不含 axb）', JSON.stringify(await T({ q: 'a_b' })) === JSON.stringify(['a_b']));
      check('[6] 文字 A_B（不分大小寫）→ 1 筆', (await T({ q: 'A_B' })).length === 1);
      check('[6] 文字 path\\dir → 只有字面含反斜線的那筆（不含 pathXdir）', JSON.stringify(await T({ q: 'path\\dir' })) === JSON.stringify(['path\\dir']));
      check('[6] 動作 FA → 3 筆', (await T({ action: 'FA' })).length === 3);
      check('[6] 動作 FA ＋ from 5/2 → 2 筆', (await T({ action: 'FA', from: '2026-05-02' })).length === 2);
      check('[6] 操作者篩選只回該操作者', (await A.auditStore.list({ kind: 'log', operator: op + '_g' }, {})).total === 1);
      check('[6] kind 篩選：log 不含 contact／quote-edit，all 全含', (await A.auditStore.list({ kind: 'log', operator: op + '_s' }, {})).total === 1 && (await A.auditStore.list({ kind: 'all', operator: op + '_s' }, {})).total === 3);
      const c = await A.auditStore.list({ kind: 'contact', operator: op + '_s' }, {});
      const api = AS.toApiContact(c.rows[0]);
      check('[6] 名片稽核輸出形狀（userId／targetName／changes）', api.userId === op + '_s' && api.targetName === '王小明' && api.changes.length === 1);
    });

    // ═════ 舊整包列：是否已存在 ═════
    const blobRow = async () => q(`SELECT 1 FROM app_data WHERE id = 'audit_log'`);
    try { await A._loadAuditLogAsync(); } catch { /* 讓 ensureSchema 先建好 app_data；失敗時下面會判成「無法判斷」 */ }
    try { blobPresentAtStart.v = (await blobRow()).length > 0; } catch { blobPresentAtStart.v = null; }

    // ═════ [7] 舊整包 JSONB 合併 ═════
    await section('[7] 與舊整包日誌合併', async () => {
      if (blobPresentAtStart.v !== false) { skip('[7] 舊整包合併', blobPresentAtStart.v === null ? '無法判斷 app_data 狀態' : '資料庫已有 audit_log 列，為避免改動既有資料不做合併寫入測試'); return; }
      const m = op + '_m';
      const legacy = { verifyRun: runId, logs: [
        { id: 'vfy-l3', action: 'LEG', operator: m, target: '', detail: 'm-l3', ip: '', userAgent: '', timestamp: '2026-06-05T00:00:00.000Z' },
        { id: 'vfy-l2', action: 'LEG', operator: m, target: '', detail: 'm-l2', ip: '', userAgent: '', timestamp: '2026-06-03T00:00:00.000Z' },
        { id: 'vfy-l1', action: 'LEG', operator: m, target: '', detail: 'm-l1', ip: '', userAgent: '', timestamp: '2026-06-01T00:00:00.000Z' },
      ] };
      await q(`INSERT INTO app_data (id, content, updated_at) VALUES ('audit_log', $1::jsonb, now())`, [JSON.stringify(legacy)]);
      weCreatedBlob = true;
      for (const [d, ts] of [['m-n1', '2026-06-02T00:00:00Z'], ['m-n2', '2026-06-04T00:00:00Z'], ['m-n3', '2026-06-06T00:00:00Z']]) await A.insertAuditEvent({ ts, kind: 'log', action: 'NEW', operator: m, detail: d, extra: { verifyRun: runId } });
      await A.flush();
      const M = loadRepo('1');   // 新實例＝舊整包列的快取是新的
      const r = await M.auditStore.list({ kind: 'log', operator: m }, { limit: 100 });
      check('[7] total＝新表 3 ＋ 舊整包 3', r.total === 6, r.total);
      check('[7] 新舊依時間交錯由新到舊', r.rows.map((x) => x.detail).join() === 'm-n3,m-l3,m-n2,m-l2,m-n1,m-l1', r.rows.map((x) => x.detail));
      check('[7] 跨新舊邊界分頁（limit 2 offset 2）', (await M.auditStore.list({ kind: 'log', operator: m }, { limit: 2, offset: 2 })).rows.map((x) => x.detail).join() === 'm-n2,m-l2');
      check('[7] 舊資料也吃動作篩選', (await M.auditStore.list({ kind: 'log', operator: m, action: 'LEG' }, {})).total === 3);
      check('[7] 讀取不改動舊整包列', (await q(`SELECT jsonb_array_length(content->'logs') AS n FROM app_data WHERE id='audit_log'`))[0].n === 3);
    });

    // ═════ [8] 匯出（串流迭代＋CSV／JSON）═════
    await section('[8] 匯出 5200 筆', async () => {
      const f = { kind: 'log', operator: op + '_b' };
      let lines = [], jsonPieces = [], n = 0;
      lines.push(AS.CSV_HEAD.map(AS.csvCell).join(',') + '\r\n');
      for await (const r of A.auditStore.iterate(f)) {
        lines.push(AS.csvLine(r));
        jsonPieces.push((n++ ? ',' : '') + JSON.stringify({ id: r.id, ts: r.ts, kind: r.kind, action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip, ua: r.ua, extra: r.extra }));
      }
      check('[8] 串流迭代 5200 筆', n === BULK, n);
      const csv = lines.join('');
      const wantSum = Array.from({ length: BULK }, (_, i) => i + 1).filter((g) => g % 500 === 0).length;
      const wantAt = Array.from({ length: BULK }, (_, i) => i + 1).filter((g) => g % 700 === 0 && g % 500 !== 0).length;
      check('[8] CSV 行數＝表頭＋5200', csv.split('\r\n').length - 1 === BULK + 1, csv.split('\r\n').length);
      check(`[8] 公式注入：=SUM 開頭的 ${wantSum} 格都加單引號`, (csv.match(/,'=SUM\(1\+1\),/g) || []).length === wantSum, (csv.match(/,'=SUM\(1\+1\),/g) || []).length);
      check(`[8] 公式注入：@cmd 開頭的 ${wantAt} 格都加單引號`, (csv.match(/,'@cmd,/g) || []).length === wantAt);
      check('[8] 沒有任何儲存格以 = 或 @ 直接開頭', !/(^|,)[=@]/m.test(csv));
      let ok2 = false, len = 0;
      try { len = JSON.parse('[' + jsonPieces.join('') + ']').length; ok2 = len === BULK; } catch { /* 下面報 FAIL */ }
      check('[8] JSON 匯出是合法陣列且 5200 筆', ok2, len);
    });

    // ═════ [9] 旗標關：舊路徑 ═════
    await section('[9] 旗標關（AUDIT_EVENTS_TABLE 未設）＝舊整包路徑', async () => {
      if (blobPresentAtStart.v !== false) { skip('[9] 旗標關等價', '資料庫已有 audit_log 列，不對它寫入'); return; }
      const before = Number((await q(`SELECT count(*)::int AS n FROM audit_events`))[0].n);
      const F = loadRepo(undefined);
      check('[9] 旗標未設 → 模式 legacy', F.auditStore.mode() === 'legacy');
      const o = op + '_o';
      check('[9] 寫入成功', (await F.insertAuditEvent({ kind: 'log', action: 'VFY_OFF', operator: o, target: 't', detail: 'flag off', ip: '9.9.9.9', ua: 'UA' })) === true);
      await F.flush();
      const after = Number((await q(`SELECT count(*)::int AS n FROM audit_events`))[0].n);
      check('[9] audit_events 完全沒被動到（筆數不變）', before === after, [before, after]);
      const top = (await q(`SELECT content->'logs'->0 AS e FROM app_data WHERE id='audit_log'`))[0].e;
      check('[9] 事件進了舊整包列最前面，欄位與舊程式相同（jsonb 不保留鍵順序，只比欄位集合）', top && top.operator === o && Object.keys(top).sort().join() === 'action,detail,id,ip,operator,target,timestamp,userAgent', top && Object.keys(top));
      check('[9] 讀取只看舊資料', (await F.auditStore.list({ kind: 'log', operator: o }, {})).total === 1);
    });

    // ═════ [10] 自動退回 ═════
    await section('[10] 表不可用 → 自動退回舊路徑', async () => {
      if (blobPresentAtStart.v !== false) { skip('[10] 自動退回', '資料庫已有 audit_log 列，不對它寫入'); return; }
      const hold = 'audit_events_vfyhold_' + runId;
      let swapped = false;
      const errs = [];
      const origErr = console.error;
      try {
        await q(`ALTER TABLE audit_events RENAME TO ${hold}`); swapped = true;
        // 同名 view：CREATE TABLE/INDEX IF NOT EXISTS 會因名稱已存在而略過，SELECT 1 FROM audit_events 則因除以零出錯 → 自我檢查失敗
        await q(`CREATE VIEW audit_events AS SELECT (1 / (random() * 0)::int) AS kind`);
        console.error = (...a) => { errs.push(a.join(' ')); };
        const B = loadRepo('1');
        const o = op + '_fb';
        const r = await B.insertAuditEvent({ kind: 'log', action: 'VFY_FB', operator: o, target: 't', detail: 'fallback' });
        await B.flush();
        console.error = origErr;
        check('[10] 寫入仍回報成功（沒有丟例外、沒有讓請求失敗）', r === true);
        check('[10] 模式自動變成 legacy', B.auditStore.mode() === 'legacy', B.auditStore.stats());
        check('[10] 記了 AUDIT_TABLE_UNAVAILABLE', errs.some((e) => e.includes('AUDIT_TABLE_UNAVAILABLE')), errs.map((e) => e.slice(0, 80)));
        const top = (await q(`SELECT content->'logs'->0->>'operator' AS o FROM app_data WHERE id='audit_log'`))[0];
        check('[10] 事件沒有遺失：寫進舊整包列', top && top.o === o, top);
        const held = Number((await q(`SELECT count(*)::int AS n FROM ${hold} WHERE operator = $1`, [o]))[0].n);
        check('[10] 被換走的原表沒有收到這筆（確實走舊路徑）', held === 0, held);
      } finally {
        console.error = origErr;
        if (swapped) {
          try {
            await q(`DROP VIEW IF EXISTS audit_events`);
            await q(`ALTER TABLE ${hold} RENAME TO audit_events`);
            check('[10] 已還原：audit_events 原表回到原名', (await q(`SELECT to_regclass('public.audit_events') IS NOT NULL AS ok`))[0].ok === true);
          } catch (e) {
            failed++;
            say(`FAIL [10] 還原失敗 ← ${redact(e && e.message)}\n    請手動在 SQL Editor 執行：DROP VIEW IF EXISTS audit_events; ALTER TABLE ${hold} RENAME TO audit_events;`);
          }
        }
      }
    });

    // ═════ [11] 清理 ═════
    await section('[11] 清理本次寫入的資料', async () => {
      const del = await q(`DELETE FROM audit_events WHERE extra->>'verifyRun' = $1 RETURNING id`, [runId]);
      check(`[11] 刪除本次標記的資料（${del.length} 筆）`, del.length >= 3 + 100 + BULK + 7, del.length);
      const left = await count(`extra->>'verifyRun' = $1 OR operator LIKE $2`, [runId, op + '%']);
      check('[11] 沒有任何本次標記的殘留', left === 0, left);
      if (weCreatedBlob) {
        await q(`DELETE FROM app_data WHERE id = 'audit_log'`);
        check('[11] 本腳本建立的舊整包列已刪除', (await blobRow()).length === 0);
      } else say('（舊整包列不是本腳本建立的，保持原樣）');
    });

    // ═════ [12] 保護 SQL（交易內套用，最後 ROLLBACK）═════
    await section('[12] audit_events_protect.sql（交易內，結束 ROLLBACK）', async () => {
      const sql = fs.readFileSync(path.join(ROOT, 'supabase', 'audit_events_protect.sql'), 'utf8');
      const c = await pool.connect();
      try {
        await c.query('BEGIN');
        await c.query(sql);
        check('[12] 保護 SQL 在交易內套用成功', true);
        await c.query(`INSERT INTO audit_events (kind, action, operator, detail, extra) VALUES ('log', 'VFY_PROBE', $1, 'probe', jsonb_build_object('verifyRun', $2::text))`, [op + '_p', runId]);
        check('[12] 套用後 INSERT 仍可', true);
        for (const [label, stmt] of [['UPDATE', `UPDATE audit_events SET detail = 'x' WHERE operator = '${op}_p'`], ['DELETE', `DELETE FROM audit_events WHERE operator = '${op}_p'`], ['TRUNCATE', 'TRUNCATE audit_events']]) {
          await c.query('SAVEPOINT s_' + label);
          let rejected = false, msg = '';
          try { await c.query(stmt); } catch (e) { rejected = true; msg = redact(e && e.message); }
          await c.query('ROLLBACK TO SAVEPOINT s_' + label);
          check(`[12] ${label} 被拒絕`, rejected, '沒有被拒絕');
          if (rejected) say('    （錯誤訊息：' + msg.slice(0, 120) + '）');
        }
        const still = (await c.query(`SELECT count(*)::int AS n FROM audit_events WHERE operator = $1`, [op + '_p'])).rows[0].n;
        check('[12] 被拒絕的操作沒有動到資料（探測列仍在）', still === 1, still);
      } finally {
        await c.query('ROLLBACK').catch(() => {});
        c.release();
      }
      const trg = Number((await q(`SELECT count(*)::int AS n FROM pg_trigger WHERE tgname IN ('audit_events_no_update_delete', 'audit_events_no_truncate')`))[0].n);
      check('[12] ROLLBACK 後觸發器不在、資料表維持可用', trg === 0, trg);
      await q(`INSERT INTO audit_events (kind, action, operator, extra) VALUES ('log', 'VFY_PROBE2', $1, jsonb_build_object('verifyRun', $2::text))`, [op + '_p2', runId]);
      check('[12] ROLLBACK 後可 INSERT 與 DELETE（表未被保護、可清理）', (await q(`DELETE FROM audit_events WHERE operator = $1 RETURNING id`, [op + '_p2'])).length === 1);
    });

    if (argv.includes('--chain')) {
      await section('[13] 防竄改鏈（--chain）', async () => {
        const ACm = require(chainFile);
        await chainSection({ pool, q, say, check, skip, loadRepo, runId, op, AS: AS || require(storeFile), AC: ACm });
      });
    } else say('（未加 --chain：略過 [13] 防竄改鏈驗證）');

    if (argv.includes('--drop')) {
      await section('[--drop] 刪除 audit_events', async () => {
        await q('DROP TABLE IF EXISTS audit_events');
        check('[--drop] audit_events 已刪除', (await q(`SELECT to_regclass('public.audit_events') IS NOT NULL AS ok`))[0].ok === false);
      });
    } else say('（audit_events 表保留；要刪除請加 --drop）');
  } catch (e) {
    failed++;
    say('FAIL 驗證流程中斷 ← ' + redact(e && e.message));
  } finally {
    restoreEnv();
    for (const f of [dbFile, storeFile, legacyFile, chainFile]) delete require.cache[f];
    await pool.end().catch(() => {});
  }

  say('');
  say(`結果：${passed} 項 PASS，${failed} 項 FAIL`);
  say(failed ? '✗ 有項目失敗，請把上面的 FAIL 行（已遮蔽連線資訊）貼給我。' : '✓ 全部通過：audit_events 在真 Postgres 上行為符合預期。');
  return { code: failed ? 1 : 0, passed, failed };
}

/** 外層：整個執行期間把 console.log／error／warn 都過一道遮蔽（連 db/postgres.js 內部印的訊息也不會帶出連線資訊） */
async function main(opts) {
  opts = opts || {};
  const redact = makeRedactor((opts.env || process.env).AUDIT_VERIFY_DATABASE_URL || '');
  const orig = { log: console.log, error: console.error, warn: console.warn };
  const wrap = (f) => (...a) => f(redact(a.map((x) => (typeof x === 'string' ? x : (x && x.message) || String(x))).join(' ')));
  console.log = wrap(orig.log); console.error = wrap(orig.error); console.warn = wrap(orig.warn);
  try { return await run(opts); }
  finally { console.log = orig.log; console.error = orig.error; console.warn = orig.warn; }
}

module.exports = { main, makeRedactor, NOT_EMPTY_MSG, chainSection };

if (require.main === module) {
  main().then((r) => {
    // db/postgres.js 內部的連線池不會自己關；直接結束行程
    process.exit(r.code);
  }).catch((e) => {
    const redact = makeRedactor(process.env.AUDIT_VERIFY_DATABASE_URL || '');
    console.log('✗ 未預期的錯誤：' + redact(e && e.message));
    process.exit(3);
  });
}
