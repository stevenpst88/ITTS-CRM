#!/usr/bin/env node
'use strict';
/**
 * 寫入競爭防護（DB_WRITE_SERIALIZE／QUOTE_NO_ATOMIC）的「真 Postgres」驗證腳本 —— 由「你本人」在自己的終端機對【拋棄式測試用】Supabase 專案執行。
 * Claude／自動化流程不得執行本腳本，也不會看到任何連線字串。
 *
 * ── 用法（Windows PowerShell；連線字串只存在這個終端機的環境變數，不會寫進任何檔案或日誌）──
 *   cd <本專案 feat/write-race-guard 工作樹>        # 需有 node_modules（沒有的話：npm ci，或建 junction 指到主專案的 node_modules）
 *   $s = Read-Host -AsSecureString "測試資料庫連線字串"
 *   $env:WRITE_RACE_VERIFY_DATABASE_URL = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s))
 *   # （沿用 verify-audit-pg.js 的環境變數 AUDIT_VERIFY_DATABASE_URL 也可以；兩個都有時以 WRITE_RACE_… 為準）
 *   # （強烈建議）再設一個「正式站專案代號片段」當保險——這是除了你自己的 --expect-ref 之外唯一獨立於你輸入的防線；連線字串含這段文字就拒絕執行：
 *   $env:WRITE_RACE_VERIFY_FORBIDDEN_SUBSTR = "<正式站專案 ref 的一部分>"
 *   node scripts/verify-write-race-pg.js --expect-ref=<測試專案代號> [--allow-nonempty-test-db]
 *   node scripts/verify-write-race-pg.js --expect-ref=<測試專案代號> --yes-this-is-test-db      # 略過「輸入主機名稱前 8 字元」的互動確認
 *   node scripts/verify-write-race-pg.js --expect-ref=<測試專案代號> --keep                      # 跑完不刪暫時 schema（除錯用）
 *   Remove-Item Env:WRITE_RACE_VERIFY_DATABASE_URL, Env:WRITE_RACE_VERIFY_FORBIDDEN_SUBSTR       # 跑完清掉環境變數
 *
 * 預期時間：約 1–3 分鐘（取決於到 Supabase 的網路延遲）。
 * 結束代碼：全部 PASS＝0；任何 FAIL＝1；防護檢查擋下或無法連線＝2／3（此時完全沒有寫入）。
 *
 * ── 安全防護（任一不通過就在「任何寫入之前」中止）──
 *   1. 必須有連線字串環境變數；不可與 DATABASE_URL 相同；不可含 …FORBIDDEN_SUBSTR。
 *   2. **一定要** --expect-ref=<專案代號>（至少 6 字元），且連線字串的帳號（Supabase pooler 為 postgres.<ref>）或主機必須含該代號。
 *      （--allow-nonempty-test-db 可給可不給：本腳本完全不讀寫 app_data 的資料，只是為了和 verify-audit-pg.js 指令一致而接受它。）
 *   3. 唯讀看一眼 app_data 的 id='main'：聯絡人 >500、商機 >500 或帳號 >50 → 中止（「這看起來像正式資料庫」）。
 *   4. 需輸入主機名稱前 8 字元，或帶 --yes-this-is-test-db。
 *   隔離：所有物件都建在一個臨時 schema「wrg_verify_<隨機>」內（含自己的 quote_no_seq／quote_no_issued／模擬 blob 表），
 *   鎖鍵用隨機值（與正式的 advisory lock 鍵不同，不會互相干擾）；結束時 DROP SCHEMA … CASCADE。
 *   **不碰 app_data、不碰 public schema 的任何表。** 所有輸出都會遮蔽連線字串／密碼／主機。本腳本不使用 data.json／auth.json。
 *
 * ── 每一組 PASS 代表什麼 ──
 *   [1] 建表     臨時 schema 內 CREATE TABLE IF NOT EXISTS ×2 + ENABLE ROW LEVEL SECURITY 可重複執行
 *   [2] 取號     A：6 個「實例」(各自連線池) 並發 200 次 INSERT … ON CONFLICT DO UPDATE … RETURNING：號碼 1..200 唯一連續、
 *                已發號紀錄 200 筆一致；blob 下界(floor)銜接 500→501；較小的下界不倒退；新月份從 1 起；序號 ≥1000 不被 lpad 截斷
 *   [3] 重現     B：不加鎖的「讀-等-改-整份寫回」× 60：資訊性回報掉了幾筆（證明問題在真 PG 上存在；掉 0 筆不算失敗）；此段不得持有任何 advisory 鎖
 *   [4] 序列化   B：同樣 60 筆但每筆包在 acquireLock 內：一筆都不掉、臨界區同時最多 1 人
 *   [5] 逾時     鎖被占住時，第二位 waitMs=500 → 約 0.5 秒後得到 LOCK_TIMEOUT（SQLSTATE 55P03）；放鎖後可取得
 *   [6] 斷線     持鎖連線被 pg_terminate_backend 終止（模擬實例當掉）→ 鎖自動交給等待者
 *   [7] 閒置     持鎖者閒置超過 idle_in_transaction_session_timeout（本測試設 1.5 秒）→ DB 自己終止它、鎖交給等待者；之後它 release 不丟例外
 *   [8] 看門狗   holdMs 到期強制釋放
 *   [10] 黑洞    持鎖者的 COMMIT 永不回應（用戶端把 COMMIT 吃掉模擬）→ release() 在用戶端逾時（本測試設 0.3 秒）內返回、連線被銷毀、伺服器端鎖已釋放
 *   [9] 清理     DROP SCHEMA 後該 schema 不存在、沒有殘留 advisory 鎖
 *   （若執行到一半被中斷：臨時 schema 名稱在開頭有印出，請手動 DROP SCHEMA "<名稱>" CASCADE。）
 *   （關閉開關時的行為不變——不建表、不拿鎖、SQL 序列與原版相同——由 scripts/check-write-guard.js 與原版差異測試驗證，與真 PG 無關。）
 */
const readline = require('readline');
const crypto = require('crypto');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const { createPgWriteGuard, DEFAULT_LOCK_KEY } = require('../lib/pgWriteGuard');

function safeDecode(s) { try { return decodeURIComponent(s); } catch { return s; } }
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
    let t = String(text == null ? '' : text).replace(/postgres(ql)?:\/\/\S+/gi, '［已遮蔽連線字串］');
    for (const s of secrets) t = t.split(s).join('［已遮蔽］');
    return t;
  };
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * opts: { env, argv, out(line), ask(prompt)->Promise<string>, pg(模組，測試注入用), scale(測試用縮小規模，預設 1) }
 * 回傳 { code, passed, failed }；不呼叫 process.exit。
 */
async function run(opts) {
  opts = opts || {};
  const env = opts.env || process.env;
  const argv = opts.argv || process.argv.slice(2);
  const scale = opts.scale || 1;
  const url = env.WRITE_RACE_VERIFY_DATABASE_URL || env.AUDIT_VERIFY_DATABASE_URL || '';
  const redact = makeRedactor(url);
  const out = opts.out || ((l) => console.log(l));
  const say = (l) => out(redact(l));
  let passed = 0, failed = 0;
  const check = (name, cond, detail) => {
    if (cond) { passed++; say('PASS ' + name); } else { failed++; say('FAIL ' + name + (detail !== undefined ? '  ← ' + (typeof detail === 'string' ? detail : JSON.stringify(detail)) : '')); }
  };
  const info = (name, text) => say('INFO ' + name + '  ' + text);
  const abort = (code, msg) => { say('✗ 中止：' + msg); say('（沒有進行任何寫入）'); return { code, passed, failed }; };

  if (argv.includes('--help') || argv.includes('-h')) { say('用法與說明請看 scripts/verify-write-race-pg.js 檔案開頭的註解。'); return { code: 0, passed, failed }; }

  // ── 防護 1：環境變數與字串比對（尚未連線）──
  if (!url) return abort(2, '沒有設定 WRITE_RACE_VERIFY_DATABASE_URL（或 AUDIT_VERIFY_DATABASE_URL）；請依檔案開頭的說明用 Read-Host 設定');
  let parsed;
  try { parsed = new URL(url); if (!/^postgres(ql)?:$/.test(parsed.protocol)) throw new Error('protocol'); } catch { return abort(2, '連線字串不是合法的 postgres:// 連線字串'); }
  const prodUrl = env.DATABASE_URL || '';
  if (prodUrl) {
    let same = url === prodUrl;
    try { const p = new URL(prodUrl); same = same || (p.hostname === parsed.hostname && p.port === parsed.port && p.pathname === parsed.pathname && p.username === parsed.username); } catch { /* 比整串就好 */ }
    if (same) return abort(2, '驗證用連線字串與目前環境的 DATABASE_URL 指向同一個資料庫，拒絕執行');
  }
  const forbidden = (env.WRITE_RACE_VERIFY_FORBIDDEN_SUBSTR || env.AUDIT_VERIFY_FORBIDDEN_SUBSTR || '').trim();
  if (forbidden && url.toLowerCase().includes(forbidden.toLowerCase())) return abort(2, '連線字串含有 …FORBIDDEN_SUBSTR 標記的正式站特徵，拒絕執行');

  // ── 防護 2：一定要 --expect-ref，且連線字串的帳號或主機含該代號 ──
  const refArg = argv.find((a) => a.startsWith('--expect-ref='));
  const expectRef = refArg ? refArg.slice('--expect-ref='.length).trim() : '';
  if (!refArg) return abort(2, '必須給 --expect-ref=<測試專案代號>（本腳本一律要求，沒有不帶的模式）');
  if (expectRef.length < 6) return abort(2, '--expect-ref 的專案代號太短（至少 6 個字元）');
  const userPart = safeDecode(parsed.username || '');
  if (!(userPart.includes(expectRef) || parsed.hostname.includes(expectRef))) return abort(2, '連線字串的帳號或主機不含 --expect-ref 指定的專案代號，拒絕執行');
  const keep = argv.includes('--keep');

  // ── 載入 pg、連線（唯讀檢查）──
  let pgMod = opts.pg;
  if (!pgMod) {
    try { pgMod = require(require.resolve('pg', { paths: [ROOT] })); } catch { return abort(3, '找不到 pg 套件：請在有 node_modules 的資料夾執行（npm ci）'); }
  }
  const mkPool = (max) => {
    const p = new pgMod.Pool({ connectionString: url, ssl: { rejectUnauthorized: false }, max: max || 3, connectionTimeoutMillis: 15000 });
    if (p.on) p.on('error', () => { /* 閒置連線錯誤不外洩細節 */ });
    return p;
  };
  const ctlPool = mkPool(3);   // 控制用（建 schema、查 pg_locks、終止連線）
  const pools = [ctlPool];
  const fatal = (e, what) => abort(3, what + '：' + redact(e && e.message));

  // ── 防護 3：唯讀確認這不像正式資料庫 ──
  try {
    const c = await ctlPool.connect();
    try {
      await c.query('BEGIN READ ONLY');
      const t = (await c.query(`SELECT to_regclass('public.app_data') IS NOT NULL AS ok`)).rows[0];
      if (t && t.ok) {
        const r = (await c.query(`SELECT content FROM app_data WHERE id = 'main'`)).rows[0];
        if (r && r.content) {
          const n = (k) => (Array.isArray(r.content[k]) ? r.content[k].length : 0);
          const users = r.content._auth && Array.isArray(r.content._auth.users) ? r.content._auth.users.length : 0;
          if (n('contacts') > 500 || n('opportunities') > 500 || users > 50) {
            await c.query('ROLLBACK'); c.release(); await ctlPool.end().catch(() => {});
            return abort(2, '這看起來像正式資料庫（聯絡人／商機／帳號數量超過測試庫上限）');
          }
        }
      }
      await c.query('ROLLBACK');
    } finally { c.release(); }
  } catch (e) { await ctlPool.end().catch(() => {}); return fatal(e, '無法連線或唯讀檢查失敗'); }

  // ── 防護 4：人工確認 ──
  if (!argv.includes('--yes-this-is-test-db')) {
    const host8 = parsed.hostname.slice(0, 8);
    const ask = opts.ask || ((q) => new Promise((resolve) => { const rl = readline.createInterface({ input: process.stdin, output: process.stdout }); rl.question(q, (a) => { rl.close(); resolve(a); }); }));
    const ans = String(await ask('這是【測試用】資料庫嗎？請輸入主機名稱的前 8 個字元確認：')).trim();
    if (ans !== host8) { await ctlPool.end().catch(() => {}); return abort(2, '確認字串不符'); }
  }

  const schema = 'wrg_verify_' + crypto.randomBytes(4).toString('hex');
  // 隨機鎖鍵（< 2^53 的正整數字串）；絕不等於正式預設鍵
  let lockKey;
  do { lockKey = String(1000000000 + crypto.randomInt(0, 2000000000)); } while (lockKey === DEFAULT_LOCK_KEY);
  const Q = (sql, params) => ctlPool.query(sql, params);
  let schemaCreated = false;
  const keyCountSql = `SELECT count(*)::int AS n FROM pg_locks WHERE locktype = 'advisory' AND granted AND objsubid = 1 AND database = (SELECT oid FROM pg_database WHERE datname = current_database()) AND ((classid::bigint << 32) | objid::bigint) = $1::bigint`;
  const heldCount = async () => (await Q(keyCountSql, [lockKey])).rows[0].n;

  try {
    const ver = (await Q('SELECT version() AS v')).rows[0].v;
    say('伺服器：' + String(ver).split(' ').slice(0, 2).join(' ') + '；port ' + (parsed.port || '5432') + (parsed.port === '6543' ? '（看起來是 transaction pooler）' : (parsed.port === '5432' && /pooler/i.test(parsed.hostname) ? '（session mode pooler：本腳本最多同時 7 個連線池×3 條連線，pooler 上限太小會出現 max clients 錯誤）' : '')) + '；臨時 schema ' + schema);

    // ═══ [1] 建表 ═══
    await Q(`CREATE SCHEMA "${schema}"`);
    schemaCreated = true;
    const mkInstance = (i) => {
      const pool = mkPool(3); pools.push(pool);
      return { pool, guard: createPgWriteGuard({ getPool: () => pool, schema, lockKey, schemaLockKey: String(Number(lockKey) + 1), idleTxMs: 25000 }) };
    };
    const insts = Array.from({ length: 6 }, (_, i) => mkInstance(i));
    await Promise.all(insts.map((x) => x.guard.ensureSchema()));
    await insts[0].guard.ensureSchema();
    const tbls = (await Q(`SELECT table_name FROM information_schema.tables WHERE table_schema = $1 ORDER BY 1`, [schema])).rows.map((r) => r.table_name);
    const rls = (await Q(`SELECT c.relname, c.relrowsecurity FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace WHERE n.nspname = $1 AND c.relkind = 'r' ORDER BY 1`, [schema])).rows;
    check('[1] 建表：兩張表建立、重複執行不出錯、且已開 RLS', tbls.join() === 'quote_no_issued,quote_no_seq' && rls.length === 2 && rls.every((r) => r.relrowsecurity === true), { tbls, rls });
    await Q(`CREATE TABLE "${schema}"."wrg_blob" (id TEXT PRIMARY KEY, content JSONB NOT NULL)`);

    // ═══ [2] 原子取號 ═══
    const N_SEQ = Math.max(20, Math.round(200 * scale));
    const t0 = Date.now();
    const got = await Promise.all(Array.from({ length: N_SEQ }, (_, i) => insts[i % insts.length].guard.nextQuoteSeq('202610', 0, 'verify' + i)));
    const serials = got.map((g) => g.serial).sort((a, b) => a - b);
    check(`[2a] 取號：${N_SEQ} 次並發（${insts.length} 個實例）號碼 1..${N_SEQ} 唯一連續`, new Set(serials).size === N_SEQ && serials.every((s, i) => s === i + 1), { first: serials.slice(0, 5), last: serials.slice(-3) });
    const issued = await insts[0].guard.listIssued('202610');
    check('[2b] 已發號紀錄筆數與取號一致、單號格式 QU-YYYYMM-NNN', issued.length === N_SEQ && got.every((g) => /^QU-202610-\d{3}$/.test(g.quoteNo)), { issued: issued.length });
    info('[2] 取號耗時', `${Date.now() - t0}ms（${N_SEQ} 次並發）`);
    const f1 = await insts[1].guard.nextQuoteSeq('202610', 500, 'floor');
    const f2 = await insts[2].guard.nextQuoteSeq('202610', 3, 'stale-floor');
    check('[2c] blob 下界 500 → 接續 501；較小的下界不倒退（→ 502）', f1.serial === 501 && f2.serial === 502, [f1.quoteNo, f2.quoteNo]);
    const m2 = await insts[3].guard.nextQuoteSeq('202611', 0, 'new-month');
    check('[2d] 新月份從 001 起', m2.quoteNo === 'QU-202611-001', m2.quoteNo);
    const b1 = await insts[0].guard.nextQuoteSeq('202612', 998, 'big');
    const b2 = await insts[0].guard.nextQuoteSeq('202612', 0, 'big');
    const b3 = await insts[0].guard.nextQuoteSeq('202612', 0, 'big');
    check('[2e] 序號 ≥1000 不被 lpad 截斷（999 → 1000 → 1001）', b1.quoteNo === 'QU-202612-999' && b2.quoteNo === 'QU-202612-1000' && b3.quoteNo === 'QU-202612-1001', [b1.quoteNo, b2.quoteNo, b3.quoteNo]);

    // ═══ [3] 無鎖重現 / [4] 加鎖 ═══
    const N_BLOB = Math.max(12, Math.round(60 * scale));
    const resetBlob = async () => { await Q(`DELETE FROM "${schema}"."wrg_blob"`); await Q(`INSERT INTO "${schema}"."wrg_blob" (id, content) VALUES ('main', '{"items":[]}'::jsonb)`); };
    const rmw = async (inst, id) => {   // 讀 → 想一下（模擬 handler 與網路）→ 整份寫回
      const r = (await inst.pool.query(`SELECT content FROM "${schema}"."wrg_blob" WHERE id = 'main'`)).rows[0].content;
      await sleep(5 + Math.floor(Math.random() * 20));
      r.items.push(id);
      await inst.pool.query(`UPDATE "${schema}"."wrg_blob" SET content = $1::jsonb WHERE id = 'main'`, [JSON.stringify(r)]);
    };
    await resetBlob();
    let lockSeen = 0;
    const mon = setInterval(() => { heldCount().then((n) => { if (n > 0) lockSeen++; }).catch(() => {}); }, 100);
    await Promise.all(Array.from({ length: N_BLOB }, (_, i) => rmw(insts[i % insts.length], 'u' + i)));
    clearInterval(mon);
    const after3 = (await Q(`SELECT content FROM "${schema}"."wrg_blob" WHERE id = 'main'`)).rows[0].content.items.length;
    info('[3] 無鎖整份覆寫（重現）', after3 < N_BLOB ? `${N_BLOB} 筆並發，最後只剩 ${after3} 筆（掉 ${N_BLOB - after3} 筆）→ 問題在真 PG 上確實存在` : `${N_BLOB} 筆都在（這次沒撞到；網路延遲很低時可能發生，不算失敗）`);
    check('[3] 無鎖路徑期間沒有任何 advisory 鎖被持有（關閉開關＝不介入）', lockSeen === 0, { lockSeen });

    await resetBlob();
    let inCrit = 0, maxCrit = 0;
    const locked = async (inst, id) => {
      const h = await inst.guard.acquireLock({ waitMs: 60000, holdMs: 0 });
      try {
        inCrit++; maxCrit = Math.max(maxCrit, inCrit);
        await rmw(inst, id);
      } finally { inCrit--; await h.release(); }
    };
    const t4 = Date.now();
    await Promise.all(Array.from({ length: N_BLOB }, (_, i) => locked(insts[i % insts.length], 'l' + i)));
    const after4 = (await Q(`SELECT content FROM "${schema}"."wrg_blob" WHERE id = 'main'`)).rows[0].content.items;
    check(`[4] 加鎖後 ${N_BLOB} 筆並發：一筆都不掉、不重複`, after4.length === N_BLOB && new Set(after4).size === N_BLOB, { len: after4.length });
    check('[4] 臨界區同時最多 1 人', maxCrit === 1, { maxCrit });
    info('[4] 序列化耗時', `${Date.now() - t4}ms（約 ${Math.round((Date.now() - t4) / N_BLOB)}ms/筆；這就是全域寫入吞吐的上限估計）`);
    check('[4] 結束後沒有殘留鎖', (await heldCount()) === 0);

    // ═══ [5] 逾時 ═══
    const holder = await insts[0].guard.acquireLock({ waitMs: 5000 });
    const t5 = Date.now();
    let e5; try { await insts[1].guard.acquireLock({ waitMs: 500 }); } catch (e) { e5 = e; }
    const el5 = Date.now() - t5;
    check('[5a] 鎖被占住：waitMs=500 → LOCK_TIMEOUT（SQLSTATE 55P03）', e5 && e5.code === 'LOCK_TIMEOUT' && e5.cause && e5.cause.code === '55P03', e5 && { code: e5.code, cause: e5.cause && e5.cause.code, msg: e5.message });
    check('[5b] 逾時時間合理（0.4～3 秒）', el5 >= 400 && el5 < 3000, el5 + 'ms');
    await holder.release();
    const after5 = await insts[1].guard.acquireLock({ waitMs: 3000 });
    check('[5c] 持有者放鎖後，別人立刻取得', !!after5);
    await after5.release();

    // ═══ [6] 持鎖連線被終止（模擬實例當掉）═══
    const victim = await insts[2].guard.acquireLock({ waitMs: 5000 });
    const pid = (await Q(`SELECT pid FROM pg_locks WHERE locktype = 'advisory' AND granted AND objsubid = 1 AND database = (SELECT oid FROM pg_database WHERE datname = current_database()) AND ((classid::bigint << 32) | objid::bigint) = $1::bigint LIMIT 1`, [lockKey])).rows[0];
    let e6 = null, got6 = null;
    const waiter = insts[3].guard.acquireLock({ waitMs: 10000 }).then((h) => { got6 = h; }, (e) => { e6 = e; });
    await sleep(300);
    const t6 = Date.now();
    if (pid) await Q('SELECT pg_terminate_backend($1::int)', [pid.pid]);
    await waiter;
    check('[6] 持鎖連線被 pg_terminate_backend 終止 → 鎖自動交給等待者（不會永久卡死）', !!pid && !!got6 && !e6 && Date.now() - t6 < 5000, { pid: !!pid, e6: e6 && e6.code, ms: Date.now() - t6 });
    if (got6) await got6.release();
    await victim.release().catch(() => {});

    // ═══ [7] 閒置交易逾時 ═══
    const idleGuard = createPgWriteGuard({ getPool: () => insts[4].pool, schema, lockKey, idleTxMs: 1500 });
    const idler = await idleGuard.acquireLock({ waitMs: 5000 });
    const t7 = Date.now();
    let h7 = null, e7 = null;
    try { h7 = await insts[5].guard.acquireLock({ waitMs: 8000 }); } catch (e) { e7 = e; }
    check('[7a] 持鎖者閒置超過 idle_in_transaction_session_timeout(1.5s) → DB 自行終止、鎖交給等待者', !!h7 && !e7 && Date.now() - t7 < 7000, { e7: e7 && (e7.code + ' ' + (e7.cause && e7.cause.message)), ms: Date.now() - t7 });
    if (h7) await h7.release();
    let e7b = null; try { await idler.release(); } catch (e) { e7b = e; }
    check('[7b] 被終止的持有者之後 release() 不丟例外', !e7b, e7b && e7b.message);

    // ═══ [10] COMMIT 黑洞 ═══
    {
      const base = insts[0].pool;
      const hangPool = {
        connect: async () => { const c = await base.connect(); const q = c.query.bind(c); c.query = (sql, ...a) => (/^\s*COMMIT\s*$/.test(String(sql)) ? new Promise(() => {}) : q(sql, ...a)); return c; },
        query: (...a) => base.query(...a),
      };
      const g10 = createPgWriteGuard({ getPool: () => hangPool, schema, lockKey, idleTxMs: 25000, commitTimeoutMs: 300 });
      const h10 = await g10.acquireLock({ waitMs: 3000 });
      const t10 = Date.now();
      await h10.release();
      const ms10 = Date.now() - t10;
      check('[10a] COMMIT 永不回應：release() 在用戶端逾時內返回（不永遠卡住）', ms10 < 1500, ms10 + 'ms');
      await sleep(800);
      check('[10b] 用戶端銷毀連線後，伺服器端鎖已釋放（沒有殘留 advisory 鎖）', (await heldCount()) === 0);
      const after10 = await insts[1].guard.acquireLock({ waitMs: 3000 }).catch((e) => e);
      check('[10c] 之後別人可取得鎖', after10 && typeof after10.release === 'function');
      if (after10 && after10.release) await after10.release();
    }

    // ═══ [8] 看門狗 ═══
    let forced = 0;
    const wdGuard = createPgWriteGuard({ getPool: () => insts[0].pool, schema, lockKey, idleTxMs: 25000, onForcedRelease: () => { forced++; } });
    const wd = await wdGuard.acquireLock({ waitMs: 3000, holdMs: 400 });
    await sleep(1200);
    const afterWd = await insts[1].guard.acquireLock({ waitMs: 3000 }).catch((e) => e);
    check('[8] holdMs=400 到期強制釋放；別人可取得', wd.forced === true && forced === 1 && afterWd && typeof afterWd.release === 'function', { forced, wdForced: wd.forced });
    if (afterWd && afterWd.release) await afterWd.release();
  } catch (e) {
    failed++;
    say('FAIL 腳本執行中發生例外  ← ' + (e && e.message));
  } finally {
    // ═══ [9] 清理 ═══
    try {
      if (schemaCreated && !keep) {
        await Q(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`);
        const gone = (await Q(`SELECT to_regnamespace($1::text) IS NULL AS gone`, [schema])).rows[0].gone;
        check('[9] 清理：臨時 schema 已刪除', gone === true);
        check('[9] 清理：沒有殘留 advisory 鎖', (await heldCount()) === 0);
      } else if (schemaCreated) {
        say(`（--keep：臨時 schema ${schema} 已保留，請自行 DROP SCHEMA "${schema}" CASCADE）`);
      }
    } catch (e) {
      failed++;
      say('FAIL 清理失敗  ← ' + (e && e.message));
      if (schemaCreated) say(`請手動執行：DROP SCHEMA IF EXISTS "${schema}" CASCADE;`);
    }
    for (const p of pools) { try { await p.end(); } catch (_) { /* ignore */ } }
  }
  say(`\n結果：${passed} 通過、${failed} 失敗`);
  return { code: failed ? 1 : 0, passed, failed };
}

module.exports = { run, makeRedactor };

if (require.main === module) {
  // 不印 e.message：連線類例外的訊息可能含主機或帳號。只印例外類型與代碼
  run().then((r) => process.exit(r.code)).catch((e) => { console.error('未預期的錯誤：' + ((e && (e.name || e.code)) || 'Error') + '（為避免外洩連線資訊不顯示細節）'); process.exit(3); });
}
