#!/usr/bin/env node
'use strict';
/**
 * 防拆單提醒（QUOTE_SPLIT_GUARD=off|log|warn）檢查：單元＋端對端＋「關閉＝與基底逐位元相同」＋變異。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 * 做法：
 *  - 單元：lib/quoteSplitGuard.js 純函式（累計、金額級距、回看天數邊界、作廢不計、不同商機不混、權限過濾、設定驗證、稽核文字）。
 *  - 端對端：在「本工作樹」建立臨時 auth.json / data.json（假資料＋_s_* 臨時帳號），自行以 PORT=QS_PORT（預設 3070）啟動 server.js，
 *    依情境用環境變數 QUOTE_SPLIT_GUARD 重啟；跑完只用 PID 關閉並刪除臨時檔。
 *  - 守門：工作樹若已有 auth.json / data.json / audit.log.json*（可能是真實資料）一律拒絕執行。禁用埠：3000 / 3020 / 3030 / 3101~3191。
 *  - 關閉＝逐位元相同：把 QS_BASE_REF（預設 feat/integrate-audit3）用 git archive 解到暫存資料夾（不是真實資料），以同一組請求打「舊程式」與「新程式（預設關閉）」，
 *    比對回應（正規化 id／時間）、儲存的報價單與稽核紀錄。
 *
 *   node scripts/check-quote-split-guard.js             # 全部
 *   QS_ONLY=U,E1 node scripts/check-quote-split-guard.js  # 只跑指定段（U 單元／S 靜態／E0 預設關閉／E1 後台切換／E2 log／E3 warn／E4 環境變數優先／E5 後台 API／E6 兩人覆核／OFF 與基底相同／M 變異）
 *   QS_SKIP_OFF=1 ...   略過與基底比對（不需要 git）；QS_SKIP_MUT=1 ... 略過變異
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const Module = require('module');
const { spawn, execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.QS_PORT || '3070', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3191)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILE_NAMES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'];
const filesIn = (dir) => FILE_NAMES.map(f => path.join(dir, f));
const guardFiles = filesIn(ROOT).filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const ONLY = process.env.QS_ONLY ? new Set(process.env.QS_ONLY.split(',').map(s => s.trim().toUpperCase())) : null;
const want = (k) => !ONLY || ONLY.has(k);

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sect = (t) => console.log('\n' + t);
const sleep = (n) => new Promise(r => setTimeout(r, n));

// ═════════════ U：單元（純函式）═════════════
const DAY = 86400000;
const NOW = Date.parse('2026-10-16T04:00:00Z');
const mkq = (id, oppId, price, o) => Object.assign({
  id, quoteNo: 'QU-T-' + id, owner: 'u1', oppId, company: '沙盒公司', status: 'draft',
  createdAt: new Date(NOW - DAY).toISOString(),
  items: [{ lid: 'l' + id, desc: '品項', unit: '式', qty: 1, unitPrice: price, cost: 0 }], discountType: 'none', discountValue: 0, approval: null,
}, o || {});
const M = 1000000;

/** 把一組斷言套到某個版本的 quoteSplitGuard 模組上（真實模組與變異版共用）。回傳這次失敗的數目 */
function unitTests(S, label) {
  let f0 = fail;
  const cfg = S.getConfig({});
  const ev = (data, q, c) => S.evaluate(data, q, { nowMs: NOW, cfg: c || cfg });
  const D = (...qs) => ({ quotations: qs });
  // 設定與預設
  eq([cfg.mode, cfg.tier1Yuan, cfg.tier2Yuan, cfg.lookbackDays], ['off', 10000000, 50000000, 90], label + '預設值：off／1,000 萬／5,000 萬／90 天（對齊簽核引擎 AMOUNT）');
  // 無商機／無金額
  let q = mkq('a', null, 6 * M); delete q.oppId;
  eq(ev(D(q), q), { applicable: false, reason: 'NO_OPP' }, label + '舊單沒有 oppId → 不評估');
  q = mkq('a', 'O1', 6 * M); q.items = [];
  eq(ev(D(q), q).reason, 'NO_AMOUNT', label + '沒有品項（金額算不出來）→ 不評估');
  // 單張
  q = mkq('a', 'O1', 6 * M);
  let r = ev(D(q), q);
  eq([r.applicable, r.suspect, r.ownCents, r.cumCents, r.siblings.length], [true, false, 6 * M * 100, 6 * M * 100, 0], label + '單張：累計＝本張，不疑似');
  // 多張累計過／未過第一級
  let a = mkq('a', 'O1', 6 * M), b = mkq('b', 'O1', 7 * M);
  r = ev(D(a, b), b);
  eq([r.suspect, r.ownLevel, r.cumLevel, r.thresholdYuan, r.cumCents], [true, 0, 1, 10000000, 13 * M * 100], label + '600 萬＋700 萬＝1,300 萬 → 疑似（0→1 級，門檻 1,000 萬）');
  r = ev(D(a, b), a);
  eq(r.suspect, true, label + '從另一張的角度評估同樣疑似（對稱）');
  a = mkq('a', 'O1', 6 * M); b = mkq('b', 'O1', 4 * M);
  r = ev(D(a, b), b);
  eq([r.suspect, r.cumLevel], [false, 0], label + '600 萬＋400 萬＝剛好 1,000 萬（嚴格大於才算）→ 不疑似');
  b = mkq('b', 'O1', 4 * M + 0.01);
  r = ev(D(a, b), b);
  eq([r.suspect, r.cumLevel], [true, 1], label + '600 萬＋400 萬 0.01 → 疑似（金額到分，不經浮點誤差）');
  // 第二級（5,000 萬）
  a = mkq('a', 'O1', 30 * M); b = mkq('b', 'O1', 25 * M);
  r = ev(D(a, b), b);
  eq([r.suspect, r.ownLevel, r.cumLevel, r.thresholdYuan], [true, 1, 2, 50000000], label + '2,500 萬＋3,000 萬＝5,500 萬 → 1→2 級，門檻 5,000 萬');
  r = ev(D(a, mkq('c2', 'O1', 20 * M), mkq('b3', 'O1', 6 * M)), mkq('b3', 'O1', 6 * M));
  eq([r.suspect, r.ownLevel, r.cumLevel], [true, 0, 2], label + '600 萬＋3,000 萬＋2,000 萬＝5,600 萬 → 0→2 級（一次跳兩級）');
  r = ev(D(a, b), a);
  eq([r.suspect, r.ownLevel, r.cumLevel], [true, 1, 2], label + '本張 3,000 萬（已需董事長）累計 5,500 萬 → 1→2 級仍疑似');
  b = mkq('b', 'O1', 10 * M);
  r = ev(D(a, b), a);
  eq([r.suspect, r.cumLevel], [false, 1], label + '3,000 萬＋1,000 萬＝4,000 萬：本張已在第一級、累計仍第一級 → 不疑似');
  a = mkq('a', 'O1', 60 * M); b = mkq('b', 'O1', 5 * M);
  r = ev(D(a, b), a);
  eq([r.suspect, r.ownLevel, r.cumLevel], [false, 2, 2], label + '本張已是第二級 → 不疑似');
  // 作廢／舊拒絕／無金額不計
  a = mkq('a', 'O1', 6 * M); b = mkq('b', 'O1', 7 * M, { invalid: { reason: 'OPP_VOID' } });
  r = ev(D(a, b), a);
  eq([r.suspect, r.cumCents, r.siblings[0].counted, r.siblings[0].why], [false, 6 * M * 100, false, 'INVALID'], label + '失效（商機作廢）的單不計入');
  b = mkq('b', 'O1', 7 * M, { status: 'rejected' });
  r = ev(D(a, b), a);
  eq([r.suspect, r.siblings[0].why], [false, 'LEGACY_REJECTED'], label + '舊狀態「已拒絕」不計入');
  b = mkq('b', 'O1', 7 * M); b.items = [];
  r = ev(D(a, b), a);
  eq([r.suspect, r.siblings[0].why], [false, 'NO_AMOUNT'], label + '沒有金額的草稿不計入（也不報錯）');
  // 回看天數邊界（90 天；剛好 90 天算在內，多 1 毫秒不算）
  a = mkq('a', 'O1', 6 * M);
  b = mkq('b', 'O1', 7 * M, { createdAt: new Date(NOW - 90 * DAY).toISOString() });
  r = ev(D(a, b), a);
  eq([r.suspect, r.siblings[0].counted], [true, true], label + '剛好 90 天前建立的算在回看範圍內');
  b = mkq('b', 'O1', 7 * M, { createdAt: new Date(NOW - 90 * DAY - 1).toISOString() });
  r = ev(D(a, b), a);
  eq([r.suspect, r.siblings[0].counted, r.siblings[0].why], [false, false, 'OUT_OF_WINDOW'], label + '早 1 毫秒 → 超出回看，不計入但仍列在清單');
  r = ev(D(a, b), a, Object.assign({}, cfg, { lookbackDays: 91 }));
  eq(r.suspect, true, label + '回看天數調成 91 → 又算進來');
  b = mkq('b', 'O1', 7 * M); delete b.createdAt;
  r = ev(D(a, b), a);
  eq(r.siblings[0].counted, true, label + '建立時間缺失 → 保守計入（漏放比多擋嚴重）');
  // 不同商機不混
  b = mkq('b', 'O2', 7 * M);
  r = ev(D(a, b), a);
  eq([r.suspect, r.siblings.length], [false, 0], label + '不同商機不混在一起');
  // 自己不重複計
  r = ev(D(a, Object.assign({}, a)), a);
  eq(r.siblings.length, 0, label + '同 id 的自己不算其他張');
  // 折扣：以折後未稅為準
  a = mkq('a', 'O1', 6 * M);
  b = mkq('b', 'O1', 8 * M, { discountType: 'percent', discountValue: 50 });   // 折後 400 萬
  r = ev(D(a, b), b);
  eq([r.ownCents, r.cumCents, r.suspect], [4 * M * 100, 10 * M * 100, false], label + '折扣後未稅 400 萬＋600 萬＝剛好 1,000 萬 → 不疑似');
  // 自訂門檻
  const c2 = Object.assign({}, cfg, { tier1Yuan: 5 * M, tier2Yuan: 8 * M });
  a = mkq('a', 'O1', 3 * M); b = mkq('b', 'O1', 3 * M);
  r = ev(D(a, b), b, c2);
  eq([r.suspect, r.thresholdYuan, r.cumLevel], [true, 5 * M, 1], label + '自訂門檻 500 萬：300 萬＋300 萬 → 疑似');
  // 大金額不溢位
  a = mkq('a', 'O1', 9e12); b = mkq('b', 'O1', 9e12);
  r = ev(D(a, b, mkq('c', 'O1', 9e12), mkq('d', 'O1', 9e12)), a);
  ok(Number.isSafeInteger(r.cumCents) && r.cumLevel === 2, label + '極大金額累計不溢位（鉗在安全整數）');
  // 排序：新的在前
  a = mkq('a', 'O1', 1 * M);
  const older = mkq('o', 'O1', 1 * M, { createdAt: new Date(NOW - 5 * DAY).toISOString() }), newer = mkq('n', 'O1', 1 * M, { createdAt: new Date(NOW - 2 * DAY).toISOString() });
  r = ev(D(a, older, newer), a);
  eq(r.siblings.map(s => s.id), ['n', 'o'], label + '其他張依建立時間新→舊排序');
  // 權限過濾
  a = mkq('a', 'O1', 6 * M); b = mkq('b', 'O1', 7 * M, { owner: 'x' }); const c = mkq('c', 'O1', 2 * M, { owner: 'y' });
  r = ev(D(a, b, c), a);
  const pub = S.toPublic(r, (x) => x.owner !== 'x', (o, s) => Object.assign(o, { ownerName: 'N:' + s.owner }));
  eq([pub.hiddenCount, pub.siblings.map(s => s.id), pub.suspect, pub.cumCents], [1, ['c'], true, 15 * M * 100], label + '看不到的其他張只計 hiddenCount，不給單號與金額；累計彙總仍給');
  ok(!JSON.stringify(pub).includes('QU-T-b') && !/"q"\s*:/.test(JSON.stringify(pub)), label + '公開形狀不含看不到的單號、也沒有 q 參照');
  eq(pub.siblings[0].ownerName, 'N:y', label + 'decorate 可補欄位');
  ok(!/cost|margin|gp|毛利|成本/i.test(JSON.stringify(pub)), label + '公開形狀沒有成本／毛利欄位');
  // 快照
  const snap = S.snapshot(r, { at: 'T', by: 'u', trigger: 'SUBMIT', mode: 'warn' });
  eq([snap.v, snap.trigger, snap.mode], [1, 'SUBMIT', 'warn'], label + '快照欄位');
  ok(snap.siblings.length === 2 && snap.cumCents === 15 * M * 100 && snap.thresholdYuan === 10 * M, label + '快照含計入的其他張與累計、門檻');
  eq(S.snapshot(ev(D(mkq('z', 'O1', 1 * M)), mkq('z', 'O1', 1 * M)), { at: 'T', by: 'u', trigger: 'CREATE', mode: 'warn' }), null, label + '不疑似 → 沒有快照');
  const sp = S.snapshotToPublic(snap, (id) => id !== 'b');
  eq([sp.hiddenCount, sp.siblings.map(s => s.id)], [1, ['c']], label + '快照依檢視者過濾');
  eq(S.snapshotToPublic(null, () => true), null, label + '沒有快照 → null');
  // 稽核文字
  const det = S.auditDetail(r, a, { trigger: 'CREATE', mode: 'log' });
  ok(det.includes('疑似拆單（建立報價單）') && det.includes('O1') && det.includes('15,000,000') && det.includes('QU-T-b') && det.includes('10,000,000'), label + '稽核 detail 含商機、累計、門檻、其他單號');
  const meta = S.parseAuditMeta(det);
  eq([meta.trigger, meta.oppId, meta.cumCents, meta.thresholdYuan, meta.siblings.length], ['CREATE', 'O1', 15 * M * 100, 10 * M, 2], label + 'meta 可還原');
  eq(S.parseAuditMeta('沒有 meta'), null, label + '無 meta → null');
  eq(S.parseAuditMeta('x｜meta={壞掉'), null, label + 'meta 壞掉 → null');
  // 設定驗證
  const good = { mode: 'warn', tier1Yuan: 8000000, tier2Yuan: '60000000', lookbackDays: '30' };
  eq(S.validateConfig(good), { ok: true, value: { mode: 'warn', tier1Yuan: 8000000, tier2Yuan: 60000000, lookbackDays: 30 } }, label + '合法設定（數字字串可）');
  [
    [{ ...good, mode: 'on' }, 'BAD_SPLIT_MODE'], [{ ...good, mode: undefined }, 'BAD_SPLIT_MODE'],
    [{ ...good, tier1Yuan: 0 }, 'BAD_SPLIT_TIER1'], [{ ...good, tier1Yuan: 'abc' }, 'BAD_SPLIT_TIER1'], [{ ...good, tier1Yuan: 1.5 }, 'BAD_SPLIT_TIER1'], [{ ...good, tier1Yuan: 1e13 }, 'BAD_SPLIT_TIER1'],
    [{ ...good, tier2Yuan: 8000000 }, 'BAD_SPLIT_TIER_ORDER'], [{ ...good, tier2Yuan: 100 }, 'BAD_SPLIT_TIER_ORDER'],
    [{ ...good, lookbackDays: 0 }, 'BAD_SPLIT_DAYS'], [{ ...good, lookbackDays: 731 }, 'BAD_SPLIT_DAYS'], [{ ...good, lookbackDays: -3 }, 'BAD_SPLIT_DAYS'], [{ ...good, lookbackDays: '9e1' }, 'BAD_SPLIT_DAYS'],
    [null, 'BAD_SPLIT_CONFIG'], ['x', 'BAD_SPLIT_CONFIG'],
  ].forEach(([b, code], i) => { const v = S.validateConfig(b); eq([v.ok, v.status, v.code], [false, 400, code], label + '非法設定 #' + i + ' → ' + code); });
  eq(S.validateConfig({ ...good, lookbackDays: 730 }).ok, true, label + '回看 730 天（上限）合法');
  eq(S.validateConfig({ ...good, lookbackDays: 1 }).ok, true, label + '回看 1 天（下限）合法');
  // 儲存值不可信
  const g1 = S.getConfig({ quoteSplitGuard: { mode: 'weird', tier1Yuan: -5, tier2Yuan: 'x', lookbackDays: 99999 } });
  eq([g1.mode, g1.tier1Yuan, g1.tier2Yuan, g1.lookbackDays], ['off', 10000000, 50000000, 90], label + '儲存值亂掉 → 全部退回預設');
  const g2 = S.getConfig({ quoteSplitGuard: { mode: 'warn', tier1Yuan: 70000000, tier2Yuan: 60000000, lookbackDays: 10 } });
  ok(g2.tier2Yuan > g2.tier1Yuan && g2.mode === 'warn' && g2.lookbackDays === 10, label + '兩級顛倒 → 第二級修正為大於第一級');
  // 環境變數
  eq(S.envMode({}), { set: false, mode: null, raw: '' }, label + '環境變數沒設');
  eq(S.envMode({ QUOTE_SPLIT_GUARD: ' WARN ' }), { set: true, mode: 'warn', raw: 'warn' }, label + '環境變數不分大小寫、去空白');
  eq(S.envMode({ QUOTE_SPLIT_GUARD: 'yes' }).mode, null, label + '環境變數值不合法 → mode null');
  eq(S.effectiveMode({ mode: 'warn' }, {}).mode, 'warn', label + '無環境變數 → 用後台設定');
  eq([S.effectiveMode({ mode: 'warn' }, { QUOTE_SPLIT_GUARD: 'log' }).mode, S.effectiveMode({ mode: 'warn' }, { QUOTE_SPLIT_GUARD: 'log' }).source], ['log', 'env'], label + '環境變數優先於後台設定');
  eq(S.effectiveMode({ mode: 'warn' }, { QUOTE_SPLIT_GUARD: 'oops' }).mode, 'off', label + '環境變數值不合法 → off（寧可關）');
  eq(S.effectiveMode({ mode: 'warn' }, { QUOTE_SPLIT_GUARD: 'off' }).mode, 'off', label + '環境變數 off 蓋過後台 warn');
  // levelOf／fmt／diff
  eq([S.levelOf(10 * M * 100, cfg), S.levelOf(10 * M * 100 + 1, cfg), S.levelOf(50 * M * 100, cfg), S.levelOf(50 * M * 100 + 1, cfg), S.levelOf(0, cfg)], [0, 1, 1, 2, 0], label + 'levelOf 邊界（嚴格大於，到分）');
  eq([S.fmtYuan(123456789), S.fmtYuan(100), S.fmtYuan(5)], ['1,234,567.89', '1', '0.05'], label + 'fmtYuan');
  eq(S.configDiff({ mode: 'off', tier1Yuan: 10 * M, tier2Yuan: 50 * M, lookbackDays: 90 }, { mode: 'warn', tier1Yuan: 8 * M, tier2Yuan: 50 * M, lookbackDays: 60 }).length, 3, label + 'configDiff 列出三項異動');
  eq(S.configDiff({ mode: 'off', tier1Yuan: 1, tier2Yuan: 2, lookbackDays: 3 }, { mode: 'off', tier1Yuan: 1, tier2Yuan: 2, lookbackDays: 3 }), [], label + 'configDiff 無異動');
  return fail - f0;
}

// ═════════════ 端對端沙盒 ═════════════
const PW = 'Qs1' + crypto.randomBytes(9).toString('hex') + '!Z';       // 只在記憶體
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false }, extra || {});
const mkOpp = (id, owner) => ({ id, title: '沙盒商機 ' + id, company: '沙盒公司甲', owner, category: 'ERP', stage: 'B', amount: 100, expectedDate: new Date().getFullYear() + '-12-31', product: 'PS', createdAt: nowIso });
function seedFiles(dir, mutate) {
  const users = [
    mkUser('_s_admin', 'admin'), mkUser('_s_admin2', 'admin'),
    mkUser('_s_m1', 'manager1', { bu: ['ERP'] }), mkUser('_s_m1b', 'manager1', { bu: ['ERP'] }),
    mkUser('_s_s1', 'user', { bu: ['ERP'], supervisor: '_s_m1' }),
    mkUser('_s_s2', 'user', { bu: ['ERP'], supervisor: '_s_m1b' }),
    mkUser('_s_gma', 'executive', {}), mkUser('_s_gmb', 'executive', {}), mkUser('_s_ch', 'executive', {}),
    mkUser('_s_sec1', 'secretary', { bu: ['ERP', 'ITS', 'MDM', 'CRM'] }), mkUser('_s_sec4', 'secretary', { bu: ['ITS'] }),
    mkUser('_s_cm', 'consult_manager_south', { bu: ['ERP'] }),
  ];
  fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [], companies: [{ id: 'co1', name: '沙盒公司甲', taxId: '12345675', industry: '製造業' }],
    opportunities: [mkOpp('o-A', '_s_s1'), mkOpp('o-B', '_s_s1'), mkOpp('o-C', '_s_s1'), mkOpp('o-S2', '_s_s2')],
    visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [],
    quoteApproval: { roster: { gm: ['_s_gma', '_s_gmb'], chairman: ['_s_ch'], boardProxy: [], costProviders: ['_s_cm'], sealManagers: [] }, productClasses: { PS: { cls: 'software', costBySales: true }, CS: { cls: 'consult', costBySales: false } } },
    approvalFlowConfig: { version: 1, buSettings: { ERP: { hasMgr2Step: false } }, exceptions: [] },
  };
  if (mutate) mutate(d);
  fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify(d, null, 2));
}
function cleanupIn(dir) { filesIn(dir).forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }
const readData = (dir) => JSON.parse(fs.readFileSync(path.join(dir || ROOT, 'data.json'), 'utf8'));
const patchData = (fn, dir) => { const d = readData(dir); fn(d); fs.writeFileSync(path.join(dir || ROOT, 'data.json'), JSON.stringify(d, null, 2)); };

let child = null, childPid = null;
async function startServer(dir, env) {
  const e = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE }, env || {});
  delete e.DATABASE_URL; delete e.MAIL_MODE;
  if (!env || !('QUOTE_SPLIT_GUARD' in env)) delete e.QUOTE_SPLIT_GUARD;
  child = spawn(process.execPath, [path.join(dir, 'server.js')], { cwd: dir, env: e, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.QS_DEBUG) process.stdout.write('[srv] ' + d); };
  child.stdout.on('data', onData); child.stderr.on('data', onData);
  for (let i = 0; i < 100; i++) {
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
let jar = {};
async function fresh(env, mutate, dir) {   // 重啟：清資料檔、重種、換環境變數
  await stopServer();
  const d = dir || ROOT;
  cleanupIn(d); seedFiles(d, mutate); jar = {};
  await startServer(d, env);
}
async function restartKeep(env, dir) { await stopServer(); jar = {}; await startServer(dir || ROOT, env); }   // 保留資料檔重啟（資料可先 patchData）

async function http(who, method, url, body) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fresh2 = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fresh2) jar[who] = fresh2.split(';')[0];
  return { status: res.status, json, text: buf.toString('utf8') };
}
async function login(who) {
  const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status !== 200) throw new Error('登入失敗 ' + who + ' ' + res.status);
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  jar[who] = sc.map(c => c.split(';')[0]).join('; ');
}
const loginAll = async (names) => { for (const n of names) await login(n); };

const mkQuote = (who, price, cost, extra) => http(who, 'POST', '/api/quotations', Object.assign({ oppId: 'o-A', company: '沙盒公司甲', products: ['PS'], items: [{ desc: '品項A', unit: '式', qty: 1, unitPrice: price, cost }] }, extra || {}));
const submitQ = (who, id) => http(who, 'POST', `/api/quotations/${id}/submit`);
const getQ = (who, id) => http(who, 'GET', `/api/quotations/${id}`);
const logs = async () => { const r = await http('_s_admin', 'GET', '/api/admin/logs'); return Array.isArray(r.json) ? r.json : []; };
const byAction = (ls, a) => ls.filter(l => l.action === a);
const stepsOf = (j) => ((j && j.approval && j.approval.steps) || []).map(s => s.tier);
const SG = '/api/admin/quote-split-guard';
const P6 = 6 * M, P7 = 7 * M;   // 價格（元）；成本設為 90%：毛利 10%（軟體列一級主管終局）→ 簽核層級只看金額（皆 ≤ 1,000 萬）

// ───────── E0：預設關閉 ─────────
async function secE0() {
  sect('[E0] 預設（環境變數沒設、後台沒設定）＝ off：完全沒有提醒');
  await fresh(undefined);
  await loginAll(['_s_admin', '_s_s1', '_s_m1']);
  let g = await http('_s_admin', 'GET', SG);
  eq([g.status, g.json.effective.mode, g.json.effective.source, g.json.config.mode], [200, 'off', 'config', 'off'], 'GET 設定：預設生效 off、來源 config');
  eq([g.json.config.tier1Yuan, g.json.config.tier2Yuan, g.json.config.lookbackDays], [10000000, 50000000, 90], '預設門檻 1,000 萬／5,000 萬／90 天');
  const q1 = await mkQuote('_s_s1', P6, P6 * 0.9);
  const q2 = await mkQuote('_s_s1', P7, P7 * 0.9);
  eq([q1.status, q2.status], [201, 201], '兩張建立成功');
  ok(!('split' in q2.json), 'off：建立回應沒有 split 欄位');
  const gq = await getQ('_s_s1', q2.json.id);
  ok(!('split' in gq.json), 'off：單筆 GET 沒有 split');
  const sb = await submitQ('_s_s1', q2.json.id);
  eq(sb.status, 200, 'off：送簽成功');
  ok(!('split' in sb.json), 'off：送簽回應沒有 split');
  eq(stepsOf(sb.json), ['mgr1'], 'off：簽核層級只看該張（700 萬 → 一級主管）');
  await sleep(300);
  const d = readData();
  ok(d.quotations.every(x => !('splitGuard' in x)), 'off：資料裡沒有任何 q.splitGuard');
  ok(!('quoteSplitGuard' in d), 'off：資料裡沒有 quoteSplitGuard 設定（沒人儲存過）');
  eq(byAction(await logs(), 'QUOTE_SPLIT_SUSPECT').length, 0, 'off：沒有 QUOTE_SPLIT_SUSPECT 稽核');
  return { q1: q1.json.id, q2: q2.json.id };
}

// ───────── E1：後台切換模式（環境變數沒設）＝ config 來源 ─────────
async function secE1() {
  sect('[E1] 後台把模式切到 warn／off（無環境變數）：立即生效、寫稽核');
  await fresh(undefined);
  await loginAll(['_s_admin', '_s_s1']);
  let r = await http('_s_admin', 'PUT', SG, { mode: 'warn', tier1Yuan: 10000000, tier2Yuan: 50000000, lookbackDays: 90 });
  eq([r.status, r.json && r.json.success], [200, true], 'PUT 模式 warn → 200');
  let g = await http('_s_admin', 'GET', SG);
  eq([g.json.effective.mode, g.json.effective.source, g.json.config.updatedBy], ['warn', 'config', '_s_admin'], '生效 warn、來源 config、updatedBy 記錄');
  const a = await mkQuote('_s_s1', P6, P6 * 0.9);
  ok(a.json && !a.json.split, 'warn：第一張（沒有其他張）沒有 split');
  const b = await mkQuote('_s_s1', P7, P7 * 0.9);
  ok(b.json && b.json.split && b.json.split.live && b.json.split.live.suspect === true, 'warn：第二張建立回應 split.live.suspect=true');
  eq([b.json.split.live.cumCents, b.json.split.live.ownCents, b.json.split.live.thresholdYuan, b.json.split.live.windowDays], [13 * M * 100, 7 * M * 100, 10000000, 90], 'split 內容：累計 1,300 萬、本張 700 萬、門檻 1,000 萬、90 天');
  eq(b.json.split.live.siblings.map(s => [s.id, s.revenueCents, s.counted]), [[a.json.id, 6 * M * 100, true]], '清單：另一張的單號 id、金額、計入');
  r = await http('_s_admin', 'PUT', SG, { mode: 'off', tier1Yuan: 10000000, tier2Yuan: 50000000, lookbackDays: 90 });
  eq(r.status, 200, 'PUT 模式 off → 200');
  const b2 = await getQ('_s_s1', b.json.id);
  ok(!('split' in b2.json), 'off 之後：單筆 GET 又沒有 split');
  const ls = await logs();
  const sv = byAction(ls, 'SAVE_QUOTE_SPLIT_GUARD');
  eq(sv.length, 2, '儲存設定寫了兩筆稽核 SAVE_QUOTE_SPLIT_GUARD');
  ok(sv.some(l => /模式 off→warn/.test(l.detail)) && sv.some(l => /模式 warn→off/.test(l.detail)), '稽核 detail 記 before→after（模式 off→warn、warn→off）');
  ok(sv.every(l => l.operator === '_s_admin' && l.target === 'quoteSplitGuard'), '稽核 operator／target 正確');
  r = await http('_s_admin', 'PUT', SG, { mode: 'warn', tier1Yuan: 8000000, tier2Yuan: 50000000, lookbackDays: 60 });
  const sv2 = byAction(await logs(), 'SAVE_QUOTE_SPLIT_GUARD');
  ok(sv2.some(l => /第一級門檻 10,000,000→8,000,000 元/.test(l.detail) && /回看天數 90→60/.test(l.detail)), '門檻與回看天數的變更也寫進稽核');
}

// ───────── E2：環境變數 log ─────────
async function secE2() {
  sect('[E2] QUOTE_SPLIT_GUARD=log：只寫稽核，回應不變');
  await fresh({ QUOTE_SPLIT_GUARD: 'log' });
  await loginAll(['_s_admin', '_s_s1']);
  const a = await mkQuote('_s_s1', P6, P6 * 0.9);
  eq(byAction(await logs(), 'QUOTE_SPLIT_SUSPECT').length, 0, 'log：第一張沒有稽核');
  const b = await mkQuote('_s_s1', P7, P7 * 0.9);
  eq(b.status, 201, 'log：建立成功');
  ok(!('split' in b.json), 'log：建立回應沒有 split（不提醒使用者）');
  let ls = byAction(await logs(), 'QUOTE_SPLIT_SUSPECT');
  eq(ls.length, 1, 'log：第二張建立 → 1 筆 QUOTE_SPLIT_SUSPECT');
  ok(ls[0].operator === '_s_s1' && ls[0].target === b.json.quoteNo, '稽核 operator＝建單人、target＝單號');
  ok(ls[0].detail.includes('o-A') && ls[0].detail.includes(a.json.quoteNo) && ls[0].detail.includes('13,000,000') && ls[0].detail.includes('10,000,000') && ls[0].detail.includes('建立報價單'), '稽核 detail：商機、其他單號、累計、門檻、時機');
  const sb = await submitQ('_s_s1', b.json.id);
  eq(sb.status, 200, 'log：送簽成功（不擋）');
  ok(!('split' in sb.json), 'log：送簽回應沒有 split');
  eq(stepsOf(sb.json), ['mgr1'], 'log：不改簽核層級（仍是一級主管）');
  ls = byAction(await logs(), 'QUOTE_SPLIT_SUSPECT');
  eq(ls.length, 2, 'log：送簽再記一筆（共 2 筆）');
  ok(/送簽/.test(ls[0].detail) || /送簽/.test(ls[1].detail), '其中一筆是送簽時機');
  await sleep(300);
  ok(readData().quotations.every(x => !('splitGuard' in x)), 'log：不在報價單上存快照');
  // 不疑似：600 + 400（剛好 1,000 萬）與 單張／不同商機
  const c = await mkQuote('_s_s1', 4 * M, 4 * M * 0.9, { oppId: 'o-B' });
  const d2 = await mkQuote('_s_s1', 6 * M, 6 * M * 0.9, { oppId: 'o-B' });
  eq(d2.status, 201, '另一個商機建兩張（共 1,000 萬）');
  eq(byAction(await logs(), 'QUOTE_SPLIT_SUSPECT').length, 2, 'log：剛好 1,000 萬、不同商機都不增加稽核');
  // 管理員查詢 API
  const rr = await http('_s_admin', 'GET', '/api/admin/quote-split-suspects');
  eq(rr.status, 200, '近期疑似拆單查詢 → 200');
  eq(rr.json.rows.length, 2, '查詢回 2 筆');
  ok(rr.json.rows.every(x => x.meta && x.meta.oppId === 'o-A' && x.meta.cumCents === 13 * M * 100 && !/cost|margin|毛利|成本/.test(JSON.stringify(x))), '查詢列含解析後 meta，且沒有成本／毛利');
  const fo = await http('_s_s1', 'GET', '/api/admin/quote-split-suspects');
  ok(fo.status === 403 || fo.status === 401, '非管理員查詢 → 拒絕（' + fo.status + '）');
}

// ───────── E3：環境變數 warn（行為、權限、快照）─────────
async function secE3() {
  sect('[E3] QUOTE_SPLIT_GUARD=warn：提醒＋稽核、送簽快照、簽核人／秘書可見、權限過濾');
  // 預先安排：q-OTHER 屬於 _s_s2（_s_s1 的主管鏈看不到），但連在 o-A（模擬商機換手），金額 300 萬
  await fresh({ QUOTE_SPLIT_GUARD: 'warn' });
  await loginAll(['_s_admin', '_s_s1', '_s_s2', '_s_m1', '_s_m1b', '_s_sec1', '_s_sec4', '_s_cm', '_s_gma']);
  const a = await mkQuote('_s_s1', P6, P6 * 0.9);
  const b = await mkQuote('_s_s1', P7, P7 * 0.9);
  ok(b.json.split && b.json.split.live.suspect, 'warn：建單回應帶疑似拆單 split.live');
  eq(b.json.split.live.hiddenCount, 0, '建單人看得到全部自己的單 → hiddenCount 0');
  ok(!/cost|margin|毛利|成本/i.test(JSON.stringify(b.json.split)), 'split 區塊沒有成本／毛利');
  const lst = await http('_s_s1', 'GET', '/api/quotations');
  ok(Array.isArray(lst.json) && lst.json.every(x => !('split' in x)), '清單 API 不帶 split');
  eq(byAction(await logs(), 'QUOTE_SPLIT_SUSPECT').length, 1, 'warn：建單時寫稽核');
  const sb = await submitQ('_s_s1', b.json.id);
  eq(sb.status, 200, 'warn：送簽成功、不被擋');
  eq(stepsOf(sb.json), ['mgr1'], 'warn：簽核層級不變（只看該張金額與毛利）');
  ok(sb.json.split && sb.json.split.atSubmit && sb.json.split.atSubmit.cumCents === 13 * M * 100, 'warn：送簽回應有 atSubmit 快照（累計 1,300 萬）');
  await sleep(300);
  let d = readData();
  const stored = d.quotations.find(x => x.id === b.json.id);
  ok(stored.splitGuard && stored.splitGuard.v === 1 && stored.splitGuard.trigger === 'SUBMIT' && stored.splitGuard.siblings.length === 1 && stored.splitGuard.siblings[0].id === a.json.id, 'q.splitGuard 快照：trigger SUBMIT、含另一張');
  ok(!('splitGuard' in d.quotations.find(x => x.id === a.json.id)), '另一張沒有快照（沒被評估為疑似）');
  eq(byAction(await logs(), 'QUOTE_SPLIT_SUSPECT').length, 2, 'warn：送簽再記一筆（建單＋送簽）');
  const notes = await http('_s_m1', 'GET', '/api/notifications');
  ok(Array.isArray(notes.json) && notes.json.some(n => /疑似拆單/.test(n.body || n.message || '')), '簽核人的待簽通知帶「疑似拆單」提示');
  // 簽核人（直屬一級主管）看得到
  let g = await getQ('_s_m1', b.json.id);
  eq(g.status, 200, '一級主管可看這張單');
  ok(g.json.split && g.json.split.atSubmit && g.json.split.atSubmit.cumCents === 13 * M * 100, '簽核人：看到送簽快照（累計 1,300 萬）');
  eq(g.json.split.atSubmit.siblings.map(s => s.id), [a.json.id], '簽核人：同商機其他報價單清單（單號 id）');
  ok(g.json.split.live.siblings[0].ownerName && g.json.split.live.siblings[0].approvalState === 'none' && g.json.split.live.siblings[0].revenueCents === 6 * M * 100, '簽核人：清單有業務名稱、狀態、金額');
  ok(!/cost|margin|毛利|成本/i.test(JSON.stringify(g.json.split)), '簽核人拿到的 split 沒有成本／毛利');
  // 秘書（任何 BU）看得到
  g = await getQ('_s_sec1', b.json.id);
  ok(g.status === 200 && g.json.split && g.json.split.live.siblings.length === 1, '秘書（全 BU）看得到 split 與清單');
  g = await getQ('_s_sec4', b.json.id);
  ok(g.status === 200 && g.json.split && g.json.split.live.siblings.length === 1 && g.json.split.live.hiddenCount === 0, '只掛 ITS 的秘書也看得到（秘書跨 BU 規則）');
  // 無權限者：看不到那張單 → 403，更看不到 split
  g = await getQ('_s_s2', b.json.id);
  eq(g.status, 403, '別的業務打不開這張單（403），不會拿到 split');
  g = await getQ('_s_m1b', b.json.id);
  eq(g.status, 403, '別組主管打不開（403）');
  // 權限過濾：把 a 改成屬於 _s_s2（他組）→ 一級主管 m1 與建單人 s1 看不到 a，只看到 hiddenCount
  await stopServer();
  patchData(x => { x.quotations.find(q => q.id === a.json.id).owner = '_s_s2'; });
  await restartKeep({ QUOTE_SPLIT_GUARD: 'warn' });
  await loginAll(['_s_admin', '_s_s1', '_s_m1', '_s_sec1', '_s_m1b']);
  g = await getQ('_s_m1', b.json.id);
  eq([g.status, g.json.split.live.siblings.length, g.json.split.live.hiddenCount], [200, 0, 1], '權限：看不到他組的其他張 → 清單 0 筆、hiddenCount 1');
  ok(!JSON.stringify(g.json).includes(a.json.quoteNo) || !JSON.stringify(g.json.split).includes(a.json.quoteNo), '權限：split 裡沒有看不到的那張的單號');
  ok(!JSON.stringify(g.json.split).includes(String(6 * M * 100)) || g.json.split.live.siblings.length === 0, '權限：split 沒有看不到那張的個別金額');
  eq(g.json.split.live.cumCents, 13 * M * 100, '累計彙總仍給（簽核人判斷所需）');
  eq([g.json.split.atSubmit.siblings.length, g.json.split.atSubmit.hiddenCount], [0, 1], '快照也依檢視者過濾');
  g = await getQ('_s_s1', b.json.id);
  eq(g.json.split.live.hiddenCount, 1, '建單人本人同樣看不到他組的單（hiddenCount 1）');
  g = await getQ('_s_sec1', b.json.id);
  eq([g.json.split.live.siblings.length, g.json.split.live.hiddenCount], [1, 0], '秘書看得到全部');
  g = await getQ('_s_admin', b.json.id);
  eq([g.json.split.live.siblings.length, g.json.split.live.hiddenCount], [1, 0], '管理員看得到全部');
  g = await getQ('_s_m1b', a.json.id);
  eq(g.status, 200, '（對照）他組主管看得到他組的那張 a');
  // 失效的單不計
  await stopServer();
  patchData(x => { x.quotations.find(q => q.id === a.json.id).invalid = { at: nowIso, by: 'sys', reason: 'OPP_VOID', oppId: 'o-A', company: 'x', prevState: 'none' }; });
  await restartKeep({ QUOTE_SPLIT_GUARD: 'warn' });
  await loginAll(['_s_admin', '_s_s1']);
  g = await getQ('_s_admin', b.json.id);
  ok(g.json.split && g.json.split.live.suspect === false && g.json.split.live.cumCents === 7 * M * 100, '失效（商機作廢）的單不計入累計 → 不再疑似（即時評估）');
  eq(g.json.split.live.siblings.map(s => [s.counted, s.why]), [[false, 'INVALID']], '清單仍列出並標「已失效」');
  ok(g.json.split.atSubmit && g.json.split.atSubmit.cumCents === 13 * M * 100, '送簽快照仍保留當時的評估（送簽時為準）');
  // 純顧問視角不給 split
  const c = await mkQuote('_s_s1', P6, 0, { oppId: 'o-C', products: ['CS'], costBy: '_s_cm' });
  const c2 = await mkQuote('_s_s1', P7, 0, { oppId: 'o-C', products: ['CS'], costBy: '_s_cm' });
  eq([c.status, c2.status], [201, 201], '顧問案建單成功（指派顧問 _s_cm）');
  await loginAll(['_s_cm']);
  g = await getQ('_s_cm', c2.json.id);
  ok(g.status === 200 && !('split' in g.json), '純顧問（填成本）視角不給 split（避免洩漏他單金額）');
  g = await getQ('_s_s1', c2.json.id);
  ok(g.json.split && g.json.split.live.suspect, '同一張單，業務本人看得到 split');
  // 不疑似但有其他張：有清單、沒有橫幅用的 suspect
  const e1 = await mkQuote('_s_s1', 4 * M, 4 * M * 0.9, { oppId: 'o-B' });
  const e2 = await mkQuote('_s_s1', 6 * M, 6 * M * 0.9, { oppId: 'o-B' });
  ok(e2.json.split && e2.json.split.live.suspect === false && e2.json.split.live.siblings.length === 1, '剛好 1,000 萬：有清單、suspect=false');
  // 撤回再送：快照重算
  const wd = await http('_s_s1', 'POST', `/api/quotations/${b.json.id}/withdraw`);
  eq(wd.status, 200, '撤回成功');
  const sb2 = await submitQ('_s_s1', b.json.id);
  eq(sb2.status, 200, '重新送簽成功');
  await sleep(300);
  const st2 = readData().quotations.find(x => x.id === b.json.id);
  ok(!st2.splitGuard, '重新送簽時 a 已失效 → 重算為不疑似 → 舊快照被移除');
  ok(sb2.json.split && !sb2.json.split.atSubmit, '重新送簽後回應沒有 atSubmit');
}

// ───────── E4：環境變數優先於後台設定 ─────────
async function secE4() {
  sect('[E4] 環境變數優先於後台設定');
  await fresh({ QUOTE_SPLIT_GUARD: 'off' });
  await loginAll(['_s_admin', '_s_s1']);
  let r = await http('_s_admin', 'PUT', SG, { mode: 'warn', tier1Yuan: 10000000, tier2Yuan: 50000000, lookbackDays: 90 });
  eq(r.status, 200, '環境變數 off 時仍可儲存後台設定（備用值）');
  let g = await http('_s_admin', 'GET', SG);
  eq([g.json.effective.mode, g.json.effective.source, g.json.effective.envValue, g.json.config.mode], ['off', 'env', 'off', 'warn'], '生效 off（來源 env），後台備用值 warn');
  await mkQuote('_s_s1', P6, P6 * 0.9);
  const b = await mkQuote('_s_s1', P7, P7 * 0.9);
  ok(!('split' in b.json), '環境變數 off 蓋過後台 warn → 沒有 split');
  eq(byAction(await logs(), 'QUOTE_SPLIT_SUSPECT').length, 0, '也沒有稽核');
  await restartKeep({ QUOTE_SPLIT_GUARD: 'bogus' });
  await loginAll(['_s_admin', '_s_s1']);
  g = await http('_s_admin', 'GET', SG);
  eq([g.json.effective.mode, g.json.effective.envInvalid], ['off', true], '環境變數值不合法 → 視為 off、標示 envInvalid');
  await restartKeep({ QUOTE_SPLIT_GUARD: 'warn' });
  await loginAll(['_s_admin', '_s_s1']);
  const c = await mkQuote('_s_s1', 3 * M, 3 * M * 0.9);
  ok(c.json.split && c.json.split.live.siblings.length === 2, '環境變數 warn：後台備用值 warn 與環境變數一致時也正常（兩張既有的算進來）');
}

// ───────── E5：後台 API 驗證與權限 ─────────
async function secE5() {
  sect('[E5] 後台 API：驗證、權限、不改設定的失敗請求');
  await fresh(undefined);
  await loginAll(['_s_admin', '_s_s1', '_s_m1', '_s_sec1']);
  const good = { mode: 'log', tier1Yuan: 9000000, tier2Yuan: 40000000, lookbackDays: 45 };
  for (const who of ['_s_s1', '_s_m1', '_s_sec1']) {
    const p = await http(who, 'PUT', SG, good);
    ok(p.status === 403 || p.status === 401, `${who} PUT 設定 → 拒絕（${p.status}）`);
    const gg = await http(who, 'GET', SG);
    ok(gg.status === 403 || gg.status === 401, `${who} GET 設定 → 拒絕（${gg.status}）`);
  }
  const un = await http(null, 'GET', SG);
  eq(un.status, 401, '未登入 → 401');
  const before = readData().quoteSplitGuard;
  ok(before === undefined, '拒絕的請求沒有寫入設定');
  for (const [body, code] of [[{ ...good, mode: 'x' }, 'BAD_SPLIT_MODE'], [{ ...good, tier2Yuan: 9000000 }, 'BAD_SPLIT_TIER_ORDER'], [{ ...good, lookbackDays: 0 }, 'BAD_SPLIT_DAYS'], [{ ...good, tier1Yuan: 'abc' }, 'BAD_SPLIT_TIER1']]) {
    const p = await http('_s_admin', 'PUT', SG, body);
    eq([p.status, p.json && p.json.code], [400, code], 'PUT 非法 → 400 ' + code);
  }
  await sleep(300);
  ok(readData().quoteSplitGuard === undefined, '驗證失敗不寫入設定');
  eq(byAction(await logs(), 'SAVE_QUOTE_SPLIT_GUARD').length, 0, '驗證失敗不寫稽核');
  const p = await http('_s_admin', 'PUT', SG, good);
  eq(p.status, 200, '合法 PUT → 200');
  await sleep(300);
  const d = readData();
  eq([d.quoteSplitGuard.mode, d.quoteSplitGuard.tier1Yuan, d.quoteSplitGuard.tier2Yuan, d.quoteSplitGuard.lookbackDays, d.quoteSplitGuard.updatedBy], ['log', 9000000, 40000000, 45, '_s_admin'], '設定已存（含 updatedBy）');
  const g = await http('_s_admin', 'GET', SG);
  eq([g.json.effective.mode, g.json.config.lookbackDays], ['log', 45], 'GET 反映新設定（log 生效）');
  const lg = byAction(await logs(), 'SAVE_QUOTE_SPLIT_GUARD');
  eq(lg.length, 1, '成功儲存寫一筆稽核');
  ok(/模式 off→log/.test(lg[0].detail) && /第一級門檻 10,000,000→9,000,000/.test(lg[0].detail) && /第二級門檻 50,000,000→40,000,000/.test(lg[0].detail) && /回看天數 90→45/.test(lg[0].detail), '稽核 detail：四項 before→after');
  // 新門檻真的被用到：自訂 tier1=900 萬，600 萬＋400 萬 → 疑似
  await mkQuote('_s_s1', 6 * M, 6 * M * 0.9, { oppId: 'o-B' });
  await mkQuote('_s_s1', 4 * M, 4 * M * 0.9, { oppId: 'o-B' });
  const ls = byAction(await logs(), 'QUOTE_SPLIT_SUSPECT');
  eq(ls.length, 1, '自訂門檻 900 萬：600 萬＋400 萬 → 疑似（log 稽核 1 筆）');
  ok(ls[0].detail.includes('9,000,000'), '稽核記的是自訂門檻 9,000,000');
  const rr = await http('_s_admin', 'GET', '/api/admin/quote-split-suspects?days=1&limit=5');
  eq([rr.status, rr.json.days, rr.json.rows.length], [200, 1, 1], '查詢 days／limit 參數有效');
  const rbad = await http('_s_admin', 'GET', '/api/admin/quote-split-suspects?days=abc&limit=-3');
  eq(rbad.status, 200, '查詢參數亂填 → 退回預設、不報錯');
  // 前端靜態：後台頁含卡片與稽核動作中文名
  const html = (await http('_s_admin', 'GET', '/admin.html')).text;
  ok(html.includes('id="sgCard"') && html.includes('QUOTE_SPLIT_SUSPECT') && html.includes('_sgSave'), '後台頁含防拆單卡片與稽核動作中文名');
}

// ───────── E6：兩人覆核 ─────────
async function secE6() {
  sect('[E6] 兩人覆核（APPROVAL_CONFIG_DUAL_REVIEW=1）：防拆單設定也走覆核');
  await fresh({ APPROVAL_CONFIG_DUAL_REVIEW: '1' });
  await loginAll(['_s_admin', '_s_admin2']);
  const body = { mode: 'warn', tier1Yuan: 9000000, tier2Yuan: 40000000, lookbackDays: 45 };
  let r = await http('_s_admin', 'PUT', SG, body);
  eq([r.status, r.json && r.json.status], [202, 'pending'], '啟用兩人覆核：PUT → 202 待覆核');
  const rid = r.json.requestId;
  let g = await http('_s_admin', 'GET', SG);
  eq([g.json.config.mode, g.json.config.lookbackDays], ['off', 90], '提案後設定沒變');
  const same = await http('_s_admin', 'PUT', SG, { mode: 'off', tier1Yuan: 10000000, tier2Yuan: 50000000, lookbackDays: 90 });
  eq([same.status, same.json && same.json.noChange], [200, true], '與現況相同 → noChange');
  const REQ = '/api/admin/approval-config-requests';
  let one = await http('_s_admin2', 'GET', REQ + '/' + rid);
  eq(one.json.kindLabel, '防拆單提醒設定', '覆核單的種類名稱');
  ok(one.json.diffSummary.some(l => /模式 off→warn/.test(l)) && one.json.diffSummary.some(l => /回看天數 90→45/.test(l)), '差異摘要列出模式與回看天數');
  r = await http('_s_admin', 'POST', REQ + '/' + rid + '/approve', {});
  eq([r.status, r.json && r.json.code], [403, 'SELF_REVIEW_FORBIDDEN'], '提案人不能核可自己的');
  r = await http('_s_admin2', 'POST', REQ + '/' + rid + '/approve', {});
  eq(r.status, 200, '另一位管理員核可 → 200');
  g = await http('_s_admin', 'GET', SG);
  eq([g.json.config.mode, g.json.config.tier1Yuan, g.json.config.lookbackDays], ['warn', 9000000, 45], '核可後設定生效');
  await sleep(300);
  const d = readData();
  ok(d.quoteSplitGuard.reviewedBy === '_s_admin2' && d.quoteSplitGuard.changeRequestId === rid, '設定記錄覆核人與申請單號');
  const lg = byAction(await logs(), 'SAVE_QUOTE_SPLIT_GUARD');
  eq(lg.length, 1, '核可套用時寫稽核 SAVE_QUOTE_SPLIT_GUARD');
}

// ───────── OFF：與基底逐位元相同 ─────────
const UUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g;
const ISO = /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z/g;
const normText = (s) => String(s).replace(UUID, '<id>').replace(ISO, '<t>').replace(/"(contentHash|hash|itemsSig|costLinesSig)":"[0-9a-f]{16,64}"/g, '"$1":"<h>"');
async function offScenario(dir) {
  const out = [];
  const rec = (name, r) => out.push({ name, status: r.status, body: normText(JSON.stringify(r.json)) });
  await loginAll(['_s_admin', '_s_s1', '_s_m1', '_s_sec1']);
  const ids = [];
  for (const [price, opp] of [[P6, 'o-A'], [P7, 'o-A'], [30 * M, 'o-A'], [P6, 'o-B']]) {
    const r = await mkQuote('_s_s1', price, price * 0.9, { oppId: opp });
    rec('create ' + price + '/' + opp, r); ids.push(r.json.id);
  }
  rec('get q2', await getQ('_s_s1', ids[1]));
  rec('submit q2', await submitQ('_s_s1', ids[1]));
  rec('get q2 by mgr', await getQ('_s_m1', ids[1]));
  rec('get q2 by sec', await getQ('_s_sec1', ids[1]));
  rec('list s1', await http('_s_s1', 'GET', '/api/quotations'));
  rec('inbox mgr', await http('_s_m1', 'GET', '/api/quotations/inbox'));
  rec('opp quotes', await http('_s_s1', 'GET', '/api/opportunities/o-A/quotes'));
  rec('submit q3 (3000萬)', await submitQ('_s_s1', ids[2]));
  rec('withdraw q3', await http('_s_s1', 'POST', `/api/quotations/${ids[2]}/withdraw`));
  await sleep(500);
  const d = readData(dir);
  const quotations = normText(JSON.stringify((d.quotations || []).map(q => { const c = JSON.parse(JSON.stringify(q)); return c; })));
  const ls = (await http('_s_admin', 'GET', '/api/admin/logs')).json || [];
  const logsN = ls.map(l => [l.action, l.operator, l.target, normText(l.detail)]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
  return { out, quotations, logs: logsN, keys: Object.keys(d).filter(k => !['contacts', 'companies', 'opportunities', 'visits', 'contracts', 'receivables', 'callins', 'keyAccounts', 'quotations', 'quoteApproval', 'approvalFlowConfig'].includes(k)).sort() };
}
function firstDiff(a, b) {
  a = String(a); b = String(b);
  let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++;
  return '位置 ' + i + ' 新：…' + a.slice(Math.max(0, i - 30), i + 60) + '… 舊：…' + b.slice(Math.max(0, i - 30), i + 60) + '…';
}
async function secOff() {
  sect('[OFF] 預設關閉：與基底（' + (process.env.QS_BASE_REF || 'feat/integrate-audit3') + '）逐位元相同');
  const ref = process.env.QS_BASE_REF || 'feat/integrate-audit3';
  const tmp = path.join(require('os').tmpdir(), 'qs_base_' + crypto.randomBytes(4).toString('hex'));
  fs.mkdirSync(tmp, { recursive: true });
  let baseRes = null;
  try {
    const tarFile = path.join(tmp, 'base.tar');
    execFileSync('git', ['-C', ROOT, 'archive', '--format=tar', '-o', tarFile, ref, '--', '.', ':(exclude,glob)*.bat', ':(exclude,glob)*.vbs']);
    const dir = path.join(tmp, 'src'); fs.mkdirSync(dir);
    execFileSync('tar', ['-xf', 'base.tar', '-C', 'src'], { cwd: tmp });
    fs.symlinkSync(path.join(ROOT, 'node_modules'), path.join(dir, 'node_modules'), 'junction');
    await fresh(undefined, undefined, dir);
    baseRes = await offScenario(dir);
    await stopServer();
    cleanupIn(dir);
  } finally {
    await stopServer();
    // 先單獨移除 node_modules 連結（只刪連結本身）；確認連結已消失才整個刪暫存資料夾，絕不讓遞迴刪除跟進到真正的 node_modules
    const jl = path.join(tmp, 'src', 'node_modules');
    let linkGone = true;
    try { if (fs.existsSync(jl)) fs.rmdirSync(jl); } catch (e) { linkGone = false; console.log('  (略過暫存資料夾清理：連結無法移除)', e.message); }
    if (linkGone && !fs.existsSync(jl)) { try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* 已刪 */ } }
  }
  await fresh(undefined);
  const mine = await offScenario(ROOT);
  eq(mine.out.length, baseRes.out.length, 'OFF：步驟數相同');
  mine.out.forEach((m, i) => {
    const same = JSON.stringify(m) === JSON.stringify(baseRes.out[i]);
    ok(same, `OFF：步驟「${m.name}」的 HTTP 狀態與回應內容與基底逐位元相同` + (same ? '' : '｜首個差異：' + firstDiff(m.body, baseRes.out[i].body)));
  });
  ok(mine.quotations === baseRes.quotations, 'OFF：儲存的報價單（含 approval／所有欄位）與基底逐位元相同' + (mine.quotations === baseRes.quotations ? '' : '｜' + firstDiff(mine.quotations, baseRes.quotations)));
  eq(mine.logs, baseRes.logs, 'OFF：稽核紀錄（action／operator／target／detail）與基底逐字相同');
  eq(mine.keys, baseRes.keys, 'OFF：資料頂層沒有多出任何鍵（沒有 quoteSplitGuard）');
  ok(mine.out.length >= 12 && mine.logs.length >= 6, 'OFF：情境有效（至少 12 個步驟、6 筆稽核）');
}

// ───────── S：靜態 ─────────
function secS() {
  sect('[S] 靜態檢查');
  const rt = fs.readFileSync(path.join(ROOT, 'lib', 'quoteRoutes.js'), 'utf8');
  const lib = fs.readFileSync(path.join(ROOT, 'lib', 'quoteSplitGuard.js'), 'utf8');
  ok(/splitAudit\(ctx, q, splitEval\(ctx, q\), 'CREATE', req\)/.test(rt), '建單路由呼叫防拆單評估（CREATE）');
  ok(/splitAudit\(ctx, q, sg, 'SUBMIT', req\)/.test(rt), '送簽路由呼叫防拆單評估（SUBMIT）');
  ok(/CR\.registerKind\('splitguard'/.test(rt), '防拆單設定已註冊兩人覆核 kind');
  ok(/QSG\.DEFAULTS|DEFAULTS = Object\.freeze\(\{ mode: 'off'/.test(lib), '預設模式 off');
  ok(!/require\('\.\/quoteSplitGuard'\)[\s\S]{0,40}QA\.(RULES_HASH|AMOUNT) *=/.test(rt), '不改簽核引擎常數');
  // 引擎檔案沒被動到（核決表／金額門檻／RULES_HASH 釘選）
  const qa = fs.readFileSync(path.join(ROOT, 'lib', 'quoteApproval.js'), 'utf8');
  ok(!/split/i.test(qa), '簽核引擎 lib/quoteApproval.js 完全沒有拆單相關程式（層級不受影響）');
  // 三個 JS 都語法正確
  for (const f of ['lib/quoteSplitGuard.js', 'lib/quoteRoutes.js', '_client/quote.js', '_client/quote-approval.js']) {
    try { new (require('vm').Script)(fs.readFileSync(path.join(ROOT, f), 'utf8').replace(/^#!.*/, ''), { filename: f }); ok(true, f); } catch (e) { ok(false, f + ' 語法：' + e.message); }
  }
}

// ───────── M：變異 ─────────
function loadMutated(find, replace) {
  const file = path.join(ROOT, 'lib', 'quoteSplitGuard.js');
  const src = fs.readFileSync(file, 'utf8');
  if (!src.includes(find)) throw new Error('變異目標不存在：' + find);
  const m = new Module(file, module);
  m.filename = file; m.paths = Module._nodeModulePaths(path.dirname(file));
  m._compile(src.replace(find, replace), file);
  return m.exports;
}
function secM() {
  sect('[M] 變異：故意改壞實作，單元測試必須抓得到');
  const muts = [
    ['疑似判斷放寬成 >=', 'const suspect = cumLevel > ownLevel;', 'const suspect = cumLevel >= ownLevel;'],
    ['失效單也計入', "if (x.invalid) s.why = 'INVALID';", "if (false) s.why = 'INVALID';"],
    ['回看邊界改成不含', 'if (Number.isFinite(t) && t < start)', 'if (Number.isFinite(t) && t <= start)'],
    ['金額門檻改成 >=', 'if (c > BigInt(cfg.tier1Yuan) * 100n) return 1;', 'if (c >= BigInt(cfg.tier1Yuan) * 100n) return 1;'],
    ['權限過濾失效', 'if (!canView(s.q)) { if (s.counted) hidden++; return; }', 'if (false) { if (s.counted) hidden++; return; }'],
    ['不同商機也混算', 'OL.linkedQuotes(data, q.oppId).forEach', '(data.quotations || []).forEach'],
    ['環境變數不優先', "if (e.set) return { mode: e.mode || 'off'", "if (false) return { mode: e.mode || 'off'"],
    ['忽略折扣（用小計）', 'const own = QA.computeFinancials(q);', 'const own = (() => { const r = QA.computeFinancials(q); return r && r.ok ? Object.assign({}, r, { revenueCents: r.subtotalCents }) : r; })();'],
  ];
  muts.forEach(([name, a, b]) => {
    let failedN = -1;
    const origLog = console.log;
    try {
      console.log = () => {};   // 變異版的失敗訊息不印
      const keep = { pass, fail, n: failures.length };
      const mod = loadMutated(a, b);
      failedN = unitTests(mod, '(變異) ');
      // 變異版的失敗不計入總失敗：還原計數
      pass = keep.pass; fail = keep.fail; failures.length = keep.n;
    } catch (e) { failedN = -2; origLog('  變異載入失敗', name, e.message); } finally { console.log = origLog; }
    ok(failedN > 0, `變異「${name}」被單元測試抓到（失敗 ${failedN} 項）`);
  });
}

// ───────── SERVE：截圖用（QS_SERVE=1）：起一台 warn 模式、已塞好疑似拆單情境的伺服器並留著（印出 PID 與各單 id；測試密碼只印到本機終端，Ctrl+C 或依 PID 關閉）─────────
async function secServe() {
  await fresh({ QUOTE_SPLIT_GUARD: 'warn' });
  await loginAll(['_s_admin', '_s_s1']);
  const a = await mkQuote('_s_s1', P6, P6 * 0.9, { projectName: '截圖案 A 期' });
  const b = await mkQuote('_s_s1', P7, P7 * 0.9, { projectName: '截圖案 B 期' });
  await submitQ('_s_s1', b.json.id);
  const c = await mkQuote('_s_s1', 5 * M, 5 * M * 0.9, { projectName: '截圖案 C 期（草稿）' });
  console.log(JSON.stringify({ pid: childPid, port: PORT, pw: process.env.QS_SHOW_PW ? PW : '(設 QS_SHOW_PW=1 才印)', ids: { a: a.json.id, b: b.json.id, c: c.json.id }, noA: a.json.quoteNo, noB: b.json.quoteNo, noC: c.json.quoteNo }));
  await new Promise(() => {});
}

// ───────── 主程式 ─────────
async function main() {
  let ranE2E = false;
  try {
    if (process.env.QS_SERVE) { await secServe(); return; }
    if (want('U')) { sect('[U] 單元：lib/quoteSplitGuard.js'); unitTests(require(path.join(ROOT, 'lib', 'quoteSplitGuard.js')), ''); }
    if (want('S')) secS();
    const steps = [['E0', secE0], ['E1', secE1], ['E2', secE2], ['E3', secE3], ['E4', secE4], ['E5', secE5], ['E6', secE6]];
    for (const [k, fn] of steps) if (want(k)) { ranE2E = true; await fn(); }
    if (want('OFF') && !process.env.QS_SKIP_OFF) { ranE2E = true; await secOff(); }
    if (want('M') && !process.env.QS_SKIP_MUT) secM();
  } catch (e) {
    fail++; failures.push('例外：' + (e && e.stack || e));
    console.log('\n例外：', e && e.stack || e);
  } finally {
    await stopServer();
    if (ranE2E || fs.existsSync(path.join(ROOT, 'auth.json'))) cleanupIn(ROOT);
  }
  console.log(`\n通過 ${pass}／失敗 ${fail}`);
  if (fail) { console.log('失敗項目：'); failures.forEach(f => console.log(' -', f)); }
  process.exit(fail ? 1 : 0);
}
main();
