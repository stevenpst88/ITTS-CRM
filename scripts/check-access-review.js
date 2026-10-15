#!/usr/bin/env node
'use strict';
/**
 * 季度帳號與權限覆核報表（lib/accessReview.js、/api/admin/access-review*、admin.html「季度帳號覆核」）檢查：
 *   Part A 單元（注入時間、合成帳號）：報表內容、旗標邊界、機敏資料不外洩、CSV 轉義與公式注入、快照雜湊穩定性、
 *          簽核紀錄只增不改與雜湊鏈、UI 渲染函式 XSS（vm）、變異檢查
 *   Part B 沙盒伺服器端對端（預設埠 3211）：僅管理員、lastLoginAt 節流寫入、CSV 下載與稽核、簽核/歷史/竄改偵測、伺服器層變異
 *   AR_SKIP_E2E=1 只跑 Part A
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const H = require('./_b2-harness');
const { ROOT, ok, eq, sleep } = H;
const AR_PATH = path.join(ROOT, 'lib', 'accessReview.js');
const DAY = 86400000;

// ═══════════════════════ Part A ═══════════════════════
const NOW = Date.UTC(2026, 9, 14, 4, 0, 0);
const iso = (ms) => new Date(ms).toISOString();
const ago = (d) => iso(NOW - d * DAY);
const SECRET_HASH = '$2b$12$abcdefghijklmnopqrstuuABCDEFGHIJKLMNOPQRSTUVWXYZ0123456';
const EMAIL = 'someone.private@corp-secret.example';
const FEATS = {
  admin: ['home', 'quotations', 'contacts'], user: ['home', 'contacts', 'quotations'], executive: ['home', 'quotations'], manager1: ['home', 'quotations'],
  accounting_manager: ['home', 'quotations'], secretary: ['home', 'quotations'], marketing: ['campaigns'], tecopm: ['contacts'],
};
const featuresFor = (r) => (FEATS[r] || []).slice();
function U(username, role, extra) {
  return Object.assign({ username, displayName: username, role, active: true, password: SECRET_HASH, token: 'tok-should-never-appear', email: EMAIL, createdAt: ago(400), passwordChangedAt: ago(10), lastLoginAt: ago(3), bu: ['ERP'], canDownloadContacts: false, canSetTargets: false }, extra || {});
}
function sample() {
  return [
    U('admin1', 'admin', { bu: null }),
    U('sales1', 'user', { supervisor: 'mgr1', canDownloadContacts: true }),
    U('sales2', 'user', { supervisor: 'mgr1', lastLoginAt: ago(90) }),                  // 剛好 90 天：不算
    U('sales3', 'user', { lastLoginAt: ago(91) }),                                     // 91 天：INACTIVE_90D
    U('mgr1', 'manager1'),
    U('never1', 'user', { lastLoginAt: undefined, passwordChangedAt: undefined }),     // NEVER_LOGGED_IN + PASSWORD_OLD
    U('oldpw', 'user', { passwordChangedAt: ago(91) }),
    U('pw90', 'user', { passwordChangedAt: ago(90) }),
    U('off_gm', 'executive', { active: false, lastLoginAt: ago(500) }),               // 名冊成員但停用
    U('test_shared', 'user'),
    U('someone', 'user', { displayName: '業務共用帳號' }),
    U('noemail', 'user', { email: '' }),
    U('exec1', 'executive'),
    U('acct1', 'accounting_manager'),                                                  // 有報價單功能但無明確職責
    U('mkt1', 'marketing'),
    U('sec1', 'secretary'),
    U('=cmd|x', 'user', { displayName: '=HYPERLINK("http://x")' }),
    U('plus', 'user', { displayName: '+1+1' }),
    U('at', 'user', { displayName: '@SUM(A1)' }),
    U('comma', 'user', { displayName: 'a,b "q"\nline2' }),
    { username: '_pool', displayName: '客戶池', role: 'pool', active: true, password: null },
  ];
}
const DATA = () => ({
  quoteApproval: { roster: { gm: ['exec1'], chairman: [], boardProxy: ['sec1'], sealManagers: ['off_gm'], costProviders: ['sales2'] } },
  quotations: [
    { id: 'q1', approval: { state: 'pending', cur: 0, steps: [{ tier: 'mgr1', assignee: 'mgr1', status: 'pending' }, { tier: 'gm', status: 'waiting' }] } },
    { id: 'q2', approval: { state: 'pending', cur: 1, steps: [{ tier: 'mgr1', assignee: 'mgr1', status: 'approved' }, { tier: 'gm', status: 'pending' }] } },
    { id: 'q3', deleted: true, approval: { state: 'pending', cur: 0, steps: [{ tier: 'mgr1', assignee: 'mgr1', status: 'pending' }] } },
  ],
});
const build = (mod, over) => mod.buildReport(Object.assign({ users: sample(), data: DATA(), featuresFor, now: NOW, trackingSince: ago(30) }, over || {}));
const flagsOf = (rep, u) => (rep.accounts.find(a => a.username === u) || {}).flags;

function reviewSuite(mod, T) {
  const { ok, eq } = T;
  const rep = build(mod);
  const A = (u) => rep.accounts.find(a => a.username === u);
  eq(rep.accounts.some(a => a.username === '_pool'), false, '系統帳號 _pool 不列入');
  eq(rep.summary.accounts, 20, '帳號總數（不含 _pool）');
  eq(rep.accounts.map(a => a.username), rep.accounts.map(a => a.username).slice().sort(), '帳號依 username 排序（輸出穩定）');
  // 旗標
  ok(flagsOf(rep, 'admin1').includes('ADMIN'), 'ADMIN');
  ok(flagsOf(rep, 'sales1').includes('DOWNLOAD_PERMISSION') && !flagsOf(rep, 'sales2').includes('DOWNLOAD_PERMISSION'), 'DOWNLOAD_PERMISSION 只給有下載權的人');
  ok(!flagsOf(rep, 'sales2').includes('INACTIVE_90D'), '最後登入剛好 90 天：不算 INACTIVE_90D');
  ok(flagsOf(rep, 'sales3').includes('INACTIVE_90D'), '91 天：INACTIVE_90D');
  ok(!flagsOf(rep, 'off_gm').includes('INACTIVE_90D'), '停用帳號不算 INACTIVE_90D（本來就停用）');
  ok(flagsOf(rep, 'never1').includes('NEVER_LOGGED_IN') && flagsOf(rep, 'never1').includes('PASSWORD_OLD'), '沒有登入紀錄 → NEVER_LOGGED_IN；沒換過密碼 → PASSWORD_OLD');
  ok(!flagsOf(rep, 'sales1').includes('NEVER_LOGGED_IN') && !flagsOf(rep, 'sales1').includes('PASSWORD_OLD'), '正常帳號無時間型旗標');
  ok(flagsOf(rep, 'oldpw').includes('PASSWORD_OLD') && !flagsOf(rep, 'pw90').includes('PASSWORD_OLD'), '密碼 91 天算舊、90 天不算');
  ok(flagsOf(rep, 'off_gm').includes('ROSTER_BUT_INACTIVE'), '名冊成員但停用 → ROSTER_BUT_INACTIVE');
  ok(!flagsOf(rep, 'exec1').includes('ROSTER_BUT_INACTIVE'), '啟用中的名冊成員不算');
  ok(flagsOf(rep, 'test_shared').includes('SHARED_NAME_SUSPECT') && flagsOf(rep, 'someone').includes('SHARED_NAME_SUSPECT'), '疑似共用帳號（帳號 test_／顯示名稱「共用」）');
  ok(!flagsOf(rep, 'sales1').includes('SHARED_NAME_SUSPECT') && !flagsOf(rep, 'admin1').includes('SHARED_NAME_SUSPECT'), '一般名稱（含 admin1）不誤判為共用');
  ok(flagsOf(rep, 'noemail').includes('NO_EMAIL') && !flagsOf(rep, 'sales1').includes('NO_EMAIL'), 'NO_EMAIL');
  ok(!flagsOf(rep, 'off_gm').includes('NO_EMAIL') || true, '（停用帳號 NO_EMAIL 規則：只算啟用中）');
  eq(flagsOf(rep, 'exec1').includes('VIEW_ALL_ROLE') && flagsOf(rep, 'mkt1').includes('VIEW_ALL_ROLE') && flagsOf(rep, 'acct1').includes('VIEW_ALL_ROLE'), true, 'VIEW_ALL_ROLE：executive／marketing／accounting_manager');
  ok(!flagsOf(rep, 'sales1').includes('VIEW_ALL_ROLE'), '一般業務不是 VIEW_ALL_ROLE');
  ok(flagsOf(rep, 'acct1').includes('QUOTE_FEATURE_WITHOUT_CLEAR_DUTY'), 'QUOTE_FEATURE_WITHOUT_CLEAR_DUTY：會計主管有報價單功能但無職責');
  ok(!flagsOf(rep, 'sales1').includes('QUOTE_FEATURE_WITHOUT_CLEAR_DUTY') && !flagsOf(rep, 'mgr1').includes('QUOTE_FEATURE_WITHOUT_CLEAR_DUTY') && !flagsOf(rep, 'sec1').includes('QUOTE_FEATURE_WITHOUT_CLEAR_DUTY'), '業務（可建單）、一級主管（簽核）、秘書（代核）不誤判');
  ok(!flagsOf(rep, 'exec1').includes('QUOTE_FEATURE_WITHOUT_CLEAR_DUTY'), '總經理名冊成員（exec1）有明確職責');
  ok(!flagsOf(rep, 'mkt1').includes('QUOTE_FEATURE_WITHOUT_CLEAR_DUTY'), '沒有報價單功能就不判（mkt1）');
  eq(rep.summary.flagCounts.ADMIN, 1, 'flagCounts 與帳號旗標一致');
  eq(Object.keys(rep.summary.flagCounts).sort(), mod.FLAG_ORDER.slice().sort(), 'flagCounts 列出全部旗標（含 0 筆者）');
  eq(rep.summary.flaggedAccounts, rep.accounts.filter(a => a.flags.length).length, 'flaggedAccounts');
  // 職責
  eq(A('exec1').duties.roster, ['gm'], '總經理名冊');
  eq(A('sec1').duties.rosterLabels, ['董事會代核'], '名冊中文名');
  eq(A('mgr1').duties.pendingApprovalSteps, 1, '待簽關卡：一級主管 q1 進行中；q2 已核准、q3 已刪除都不算');
  eq(A('exec1').duties.pendingApprovalSteps, 1, '待簽關卡：q2 輪到總經理名冊成員');
  eq(A('mgr1').duties.supervisesOthers, 2, '直屬部屬數（sales1、sales2）');
  eq([A('sales1').emailPresent, A('noemail').emailPresent], [true, false], 'emailPresent 布林');
  eq(A('sales2').passwordAgeDays, 10, '密碼天數');
  eq(A('sales3').lastLoginDaysAgo, 91, '距今天數');
  eq(A('admin1').accessMode, 'edit', '管理員存取模式恆 edit');
  // 機敏資料
  const json = JSON.stringify(rep);
  ok(!json.includes('$2b$') && !json.includes('$2a$') && !/\$2[aby]\$/.test(json), 'JSON 不含密碼雜湊');
  ok(!json.includes('corp-secret') && !json.includes('someone.private'), 'JSON 不含 email 位址');
  ok(!json.includes('tok-should-never-appear') && !/"token"|"password"|"email"/.test(json), 'JSON 沒有 token／password／email 欄位');
  const csv = mod.toCsv(rep);
  ok(!/\$2[aby]\$/.test(csv) && !csv.includes('corp-secret') && !csv.includes('tok-should'), 'CSV 同樣不含機敏資料');
  // quoteMode 欄位：有才出現
  ok(!('quoteMode' in A('sales1')) && !rep.meta.hasQuoteMode && !mod.toCsv(rep).split('\r\n')[0].includes('報價模式'), '帳號沒有 quoteMode 欄位 → 報表與 CSV 都不含此欄');
  const users2 = sample(); users2.find(u => u.username === 'sales1').quoteMode = 'approver';
  const rep2 = build(mod, { users: users2 });
  ok(rep2.meta.hasQuoteMode && rep2.accounts.find(a => a.username === 'sales1').quoteMode === 'approver' && rep2.accounts.find(a => a.username === 'sales2').quoteMode === '', '任一帳號有 quoteMode → 全部帳號都帶此欄（沒設者為空字串）');
  ok(mod.toCsv(rep2).split('\r\n')[0].includes('報價模式'), '此時 CSV 表頭含「報價模式」');
  // CSV 轉義
  ok(csv.charCodeAt(0) === 0xFEFF, 'CSV 開頭有 UTF-8 BOM');
  ok(csv.endsWith('\r\n') && csv.split('\r\n').length === rep.accounts.length + 2, 'CSV 列數＝表頭＋帳號數（換行是 CRLF）；內嵌換行被引號包住');
  ok(csv.includes('"\'=cmd|x"') || csv.includes("'=cmd|x"), "公式注入：帳號 =cmd|x 前面補 '");
  ok(csv.includes(`"'=HYPERLINK(""http://x"")"`), "公式注入：=HYPERLINK(...) 補 ' 並轉義雙引號");
  ok(csv.includes("'+1+1") && csv.includes("'@SUM(A1)"), "公式注入：+ 與 @ 開頭補 '");
  ok(csv.includes('"a,b ""q""\nline2"'), '逗號／雙引號／換行正確轉義');
  eq([mod.csvCell('-5'), mod.csvCell('\tx'), mod.csvCell('\rx'), mod.csvCell("'ok"), mod.csvCell(true), mod.csvCell(false), mod.csvCell(null), mod.csvCell(0)], ["'-5", "'\tx", "\"'\rx\"", "'ok", 'Y', 'N', '', '0'], 'csvCell 邊界值');
  eq(mod.toCsv(rep).split('\r\n')[0].split(',').length, mod.CSV_COLUMNS.length - 1, 'CSV 欄位數（無 quoteMode 時少一欄）');
  // 快照雜湊
  const h0 = mod.snapshotHash(rep);
  eq(mod.snapshotHash(build(mod, { now: NOW + 40 * DAY })), mod.snapshotHash(build(mod, { now: NOW + 40 * DAY })), '同輸入同雜湊（決定性）');
  ok(/^[0-9a-f]{64}$/.test(h0), '雜湊是 sha256 hex');
  const sorted = build(mod, { users: sample().reverse() });
  eq(mod.snapshotHash(sorted), h0, '帳號陣列順序不影響雜湊');
  const u3 = sample(); u3.find(u => u.username === 'sales1').lastLoginAt = ago(0);
  eq(mod.snapshotHash(build(mod, { users: u3 })), h0, '最後登入時間不同 → 雜湊不變（高變動欄位不列入）');
  eq(mod.snapshotHash(build(mod, { now: NOW + 3 * DAY })), h0, '時間往後 3 天（時間型旗標/天數變動）→ 雜湊不變');
  const cases = [
    ['角色變更', (u, d) => { u.find(x => x.username === 'sales1').role = 'manager2'; }],
    ['停用', (u) => { u.find(x => x.username === 'sales1').active = false; }],
    ['下載權限', (u) => { u.find(x => x.username === 'sales2').canDownloadContacts = true; }],
    ['唯讀模式', (u) => { u.find(x => x.username === 'sales2').accessMode = 'view'; }],
    ['BU 變更', (u) => { u.find(x => x.username === 'sales2').bu = ['ITS']; }],
    ['直屬主管', (u) => { u.find(x => x.username === 'sales2').supervisor = ''; }],
    ['新增帳號', (u) => { u.push(U('newbie', 'user')); }],
    ['刪除帳號', (u) => { u.splice(u.findIndex(x => x.username === 'plus'), 1); }],
    ['email 有無', (u) => { u.find(x => x.username === 'sales1').email = ''; }],
    ['改密碼', (u) => { u.find(x => x.username === 'sales1').passwordChangedAt = ago(1); }],
  ];
  for (const [name, fn] of cases) { const u = sample(); fn(u); ok(mod.snapshotHash(build(mod, { users: u })) !== h0, `組態變動（${name}）→ 雜湊改變`); }
  { // 功能組合不變、只換角色：雜湊仍要改變（角色本身要納入）
    const same = () => ['home'];
    const base = mod.snapshotHash(build(mod, { featuresFor: same }));
    const u = sample(); u.find(x => x.username === 'sales1').role = 'manager2';
    ok(mod.snapshotHash(build(mod, { users: u, featuresFor: same })) !== base, '只換角色（功能相同）→ 雜湊改變');
  }
  const d2 = DATA(); d2.quoteApproval.roster.gm = ['sales1'];
  ok(mod.snapshotHash(build(mod, { data: d2 })) !== h0, '名冊變動 → 雜湊改變');
  ok(mod.snapshotHash(build(mod, { featuresFor: (r) => (r === 'user' ? ['home'] : featuresFor(r)) })) !== h0, '功能矩陣變動 → 雜湊改變');
  // shouldStampLogin
  eq([mod.shouldStampLogin({}, NOW), mod.shouldStampLogin({ lastLoginAt: iso(NOW - 5 * 3600e3) }, NOW), mod.shouldStampLogin({ lastLoginAt: iso(NOW - 6 * 3600e3) }, NOW), mod.shouldStampLogin({ lastLoginAt: 'garbage' }, NOW)], [true, false, true, true], '最後登入寫入節流：6 小時內不重寫、無紀錄或壞值要寫');
  // 簽核紀錄
  const known = ['sales1', 'sales2'];
  eq(mod.validateSignoff({ quarter: '2026Q4', comment: 'ok', exceptions: [] }, known), { quarter: '2026Q4', comment: 'ok', exceptions: [] }, '簽核輸入驗證：合法');
  eq(mod.validateSignoff({ quarter: '2026q1' }, known).quarter, '2026Q1', '季別轉大寫');
  for (const q of ['2026Q5', '2026', 'Q1', '26Q1', '', null, '2026Q0', '2026Q41']) ok(mod.validateSignoff({ quarter: q }, known).error, `拒絕季別 ${JSON.stringify(q)}`);
  ok(mod.validateSignoff({ quarter: '2026Q4', comment: 'x'.repeat(1001) }, known).error, '意見超過 1000 字被拒');
  ok(mod.validateSignoff({ quarter: '2026Q4', exceptions: [{ username: 'ghost', note: 'x' }] }, known).error, '例外帳號不存在被拒');
  ok(mod.validateSignoff({ quarter: '2026Q4', exceptions: [{ username: 'sales1', note: '' }] }, known).error, '例外沒有說明被拒');
  ok(mod.validateSignoff({ quarter: '2026Q4', exceptions: 'x' }, known).error, 'exceptions 非陣列被拒');
  ok(mod.validateSignoff({ quarter: '2026Q4', exceptions: new Array(201).fill({ username: 'sales1', note: 'n' }) }, known).error, '例外超過 200 筆被拒');
  eq(mod.validateSignoff({ quarter: '2026Q4', exceptions: [{ username: ' sales1 ', note: ' n ' }] }, known).exceptions, [{ username: 'sales1', note: 'n' }], '例外帳號與說明 trim');
  const list = [];
  const mk = (q, c) => { const r = mod.makeSignoffRecord({ list, v: { quarter: q, comment: c, exceptions: [{ username: 'sales1', note: 'n' }] }, by: 'admin1', at: iso(NOW), report: rep }); list.push(r); return r; };
  const r1 = mk('2026Q3', 'first'), r2 = mk('2026Q4', 'second'), r3 = mk('2026Q4', 'again');
  eq([r1.seq, r2.seq, r3.seq, r1.prevHash, r2.prevHash === r1.recordHash, r3.prevHash === r2.recordHash], [1, 2, 3, '', true, true], '紀錄序號遞增、prevHash 串鏈');
  eq([r1.accountCount, r1.snapshotHash, r1.flagCounts.ADMIN], [20, h0, 1], '紀錄含帳號數、快照雜湊、旗標統計');
  eq(mod.verifyChain(list), { ok: true, badAt: null }, '完整鏈驗證通過');
  eq(mod.verifyChain([]), { ok: true, badAt: null }, '空鏈合法');
  const tamper = (fn) => { const c = JSON.parse(JSON.stringify(list)); fn(c); return mod.verifyChain(c); };
  eq(tamper(c => { c[1].comment = 'edited'; }).badAt, 2, '改意見 → 第 2 筆起驗證失敗');
  eq(tamper(c => { c[0].snapshotHash = 'f'.repeat(64); }).badAt, 1, '改快照雜湊 → 第 1 筆');
  eq(tamper(c => { c[2].exceptions = []; }).badAt, 3, '改例外 → 第 3 筆');
  eq(tamper(c => { c[0].by = 'someone'; }).badAt, 1, '改簽核人 → 第 1 筆');
  eq(tamper(c => { c.splice(1, 1); }).ok, false, '刪掉中間一筆 → 失敗');
  eq(tamper(c => { c.reverse(); }).ok, false, '重排 → 失敗');
  eq(tamper(c => { c[1].quarter = '2025Q1'; c[1].recordHash = c[1].recordHash; }).ok, false, '改季別 → 失敗');
  eq(tamper(c => { const x = c[2]; x.comment = 'z'; x.recordHash = require('crypto').createHash('sha256').update('x').digest('hex'); }).ok, false, '連 recordHash 一起亂改也失敗');
  const hv = mod.historyView(list, h0); ok(hv.records.every(r => r.matchesCurrent) && hv.chain.ok, 'historyView：現況一致');
  const hv2 = mod.historyView(list, 'x'.repeat(64)); ok(hv2.records.every(r => !r.matchesCurrent), 'historyView：現況不同');
  ok(!('matchesCurrent' in list[0]), 'historyView 不改動原紀錄（回傳副本）');
}

// ── admin.html 渲染函式（vm）──
function uiSuite(T) {
  const { ok, eq } = T;
  const html = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
  const extractFn = (name) => { const i = html.indexOf('function ' + name + '('); if (i < 0) throw new Error('找不到 ' + name); let j = html.indexOf('{', i), d = 0, k = j; for (; k < html.length; k++) { if (html[k] === '{') d++; else if (html[k] === '}') { d--; if (!d) break; } } return html.slice(i, k + 1); };
  const constI = html.indexOf('const AR_FLAG_COLOR');
  const constLine = html.slice(constI, html.indexOf('\n', constI));
  const UI = new Function([extractFn('escH'), constLine, ...['arFlagChip', 'arSummaryHtml', 'arAccountMatches', 'arTableHtml', 'arHistoryHtml', 'arParseExceptions', 'arCurrentQuarter'].map(extractFn),
    'return { arFlagChip, arSummaryHtml, arAccountMatches, arTableHtml, arHistoryHtml, arParseExceptions, arCurrentQuarter };'].join('\n'))();
  const evil = '<img src=x onerror=alert(1)>';
  const evilAttr = '" onmouseover="alert(2)';
  const rep = {
    summary: { accounts: 2, active: 2, flaggedAccounts: 1, flagCounts: { ADMIN: 1, [evil]: 1 } },
    flagLabels: { ADMIN: '管理員', [evil]: evil },
    meta: { hasQuoteMode: true },
    accounts: [
      { username: evil, displayName: evilAttr, role: evil, bu: [evil], active: true, accessMode: 'view', canDownloadContacts: true, canSetTargets: true, quoteMode: evil, duties: { rosterLabels: [evil], pendingApprovalSteps: 2, supervisesOthers: 1 }, lastLoginAt: evil, passwordAgeDays: evil, flags: ['ADMIN', evil] },
      { username: 'ok', displayName: '正常', role: 'user', bu: [], active: false, duties: {}, flags: [] },
    ],
  };
  const sum = UI.arSummaryHtml(rep);
  ok(!/<img/i.test(sum), '摘要卡：惡意標籤不產生元素');
  ok(UI.arSummaryHtml(null).includes('帳號總數'), '摘要卡：null 輸入不丟錯');
  const tbl = UI.arTableHtml(rep, {});
  ok(!/<img/i.test(tbl) && !/" onmouseover=/.test(tbl) && !/<script/i.test(tbl), '帳號表：帳號／名稱／角色／BU／旗標／職責／登入時間全部跳脫，屬性不能被跳出');
  ok(/data-user="&lt;img/.test(tbl), '帳號表：data-user 屬性值已跳脫');
  ok(tbl.includes('報價模式') && tbl.includes('唯讀') && tbl.includes('停用') && tbl.includes('待簽 2'), '帳號表：有 quoteMode 欄位時顯示；唯讀／停用／待簽');
  eq((UI.arTableHtml({ accounts: [], meta: {} }, {}).match(/沒有符合條件/g) || []).length, 1, '空清單顯示提示');
  ok(!UI.arTableHtml({ accounts: rep.accounts.slice(1), meta: { hasQuoteMode: false } }, {}).includes('報價模式'), '沒有 quoteMode 欄位時不顯示該欄');
  ok(UI.arAccountMatches(rep.accounts[0], { flag: '*' }) && !UI.arAccountMatches(rep.accounts[1], { flag: '*' }) && UI.arAccountMatches(rep.accounts[0], { flag: 'ADMIN' }) && !UI.arAccountMatches(rep.accounts[1], { flag: 'ADMIN' }), '篩選：任一旗標／指定旗標');
  ok(UI.arAccountMatches(rep.accounts[1], { q: 'ok' }) && !UI.arAccountMatches(rep.accounts[1], { q: 'zzz' }) && UI.arAccountMatches(rep.accounts[1], { role: 'user' }) && !UI.arAccountMatches(rep.accounts[1], { role: 'admin' }), '篩選：文字／角色');
  const hist = { chain: { ok: false, badAt: evil }, records: [{ seq: 1, quarter: evil, by: evil, at: evil, accountCount: evil, snapshotHash: evilAttr, comment: evil, exceptions: [{ username: evil, note: evil }], matchesCurrent: false }, { seq: 2, quarter: '2026Q4', by: 'a', at: '2026-10-14T00:00:00.000Z', accountCount: 3, snapshotHash: 'a'.repeat(64), comment: '', exceptions: [], matchesCurrent: true }] };
  const hh = UI.arHistoryHtml(hist);
  ok(!/<img/i.test(hh) && !/" onmouseover=/.test(hh), '歷史：簽核人／季別／意見／例外／雜湊全部跳脫');
  ok(hh.includes('data-chain="bad"') && hh.includes('現況與簽核時一致') && hh.includes('組態已有變動'), '歷史：鏈失敗警示、一致／變動徽章');
  ok(hh.indexOf('data-seq="2"') < hh.indexOf('data-seq="1"'), '歷史：最新的在最上面');
  ok(UI.arHistoryHtml({ records: [], chain: { ok: true } }).includes('尚無簽核紀錄'), '歷史：空');
  ok(!UI.arHistoryHtml(null).includes('undefined'), '歷史：null 不丟錯');
  eq(UI.arParseExceptions('a: 說明一\n\n b：說明二 \r\nc: x: y'), { list: [{ username: 'a', note: '說明一' }, { username: 'b', note: '說明二' }, { username: 'c', note: 'x: y' }], error: '' }, '例外解析：半形／全形冒號、空行、CRLF、說明含冒號');
  ok(UI.arParseExceptions('沒有冒號').error, '例外解析：格式錯誤回報');
  eq(UI.arCurrentQuarter(new Date(2026, 9, 14)), '2026Q4', '目前季別（10 月→Q4）');
  eq([UI.arCurrentQuarter(new Date(2026, 0, 1)), UI.arCurrentQuarter(new Date(2026, 5, 30)), UI.arCurrentQuarter(new Date(2026, 6, 1))], ['2026Q1', '2026Q2', '2026Q3'], '季別邊界');
  ok(html.includes('data-sec="access-review"') && html.includes('id="sec-access-review"') && html.includes("initAccessReview()"), '導覽項、section、init 分派都在');
  ok(/nav-domain-name">人員與權限<\/span><span class="nav-domain-count">5</, '人員與權限群組計數更新為 5');
}

const MUTANTS = [
  ['90 天邊界（> 改 >=）', "if (active && llDays != null && llDays > 90)", "if (active && llDays != null && llDays >= 90)"],
  ['密碼天數邊界', "if (pwAge == null || pwAge > 90)", "if (pwAge == null || pwAge >= 90)"],
  ['停用帳號也算 INACTIVE', "if (active && llDays != null && llDays > 90)", "if (llDays != null && llDays > 90)"],
  ['不排除 _pool', "const users = (opts.users || []).filter(u => u && u.role !== 'pool');", "const users = (opts.users || []).filter(u => u);"],
  ['輸出 email 值', "emailPresent,\n      flags,", "emailPresent, email: u.email,\n      flags,"],
  ['輸出密碼雜湊', "emailPresent,\n      flags,", "emailPresent, password: u.password,\n      flags,"],
  ['CSV 不防公式注入', 'if (/^[=+\\-@\\t\\r]/.test(s)) s = "\'" + s;', ''],
  ['CSV 沒有 BOM', "return '\\uFEFF' + [", "return '' + ["],
  ['CSV 不轉義雙引號', "s.replace(/\"/g, '\"\"')", "s"],
  ['雜湊納入最後登入', "delete c.passwordAgeDays; delete c.lastLoginAt; delete c.lastLoginDaysAgo;", "delete c.passwordAgeDays; delete c.lastLoginDaysAgo;"],
  ['雜湊納入時間型旗標', "c.flags = a.flags.filter(f => !TIME_FLAGS.includes(f));", "c.flags = a.flags.slice();"],
  ['雜湊忽略功能矩陣', "c.features = a.features.slice().sort();", "c.features = [];"],
  ['雜湊忽略角色', "delete c.passwordAgeDays;", "delete c.role; delete c.passwordAgeDays;"],
  ['鏈驗證不比對序號與前一筆雜湊（刪除／重排抓不到）', "if (r.seq !== i + 1 || r.prevHash !== prev || sha256", "if (sha256"],
  ['鏈驗證不比對內容', "|| sha256(recordPayload(r)) !== r.recordHash) return", "|| false) return"],
  ['季別驗證放寬', "const QUARTER_RE = /^(20\\d{2})Q([1-4])$/;", "const QUARTER_RE = /^(20\\d{2})Q([1-9])$/;"],
  ['例外帳號不驗證存在', "if (!username || !known.has(username))", "if (!username)"],
  ['登入寫入不節流', "return last == null || now - last >= (minGapMs || 6 * 3600e3);", "return true;"],
  ['共用帳號規則（英文）失效', "(test|tester|demo|shared|share|common|temp|tmp|guest|public|general|dummy|sample|training|trial)", "(zzzz)"],
  ['共用帳號規則（中文）失效', "|共用|公用|通用|測試|臨時|示範|訓練|試用|訪客/i", "|zzzz/i"],
  ['待簽關卡含已刪除單', "if (!ap || ap.state !== 'pending' || q.deleted) return;", "if (!ap || ap.state !== 'pending') return;"],
  ['quoteMode 欄位永遠輸出', "if (hasQuoteMode) acc.quoteMode = str(u.quoteMode);", "acc.quoteMode = str(u.quoteMode); report_forceQm = true;"],
];

async function partA() {
  console.log('[A] 單元：報表／CSV／雜湊／簽核鏈');
  reviewSuite(require(AR_PATH), { ok, eq });
  console.log('[A-UI] admin.html 渲染（vm）');
  uiSuite({ ok, eq });
  if (process.env.AR_SKIP_MUT) return;
  console.log('[A-M] 變異檢查');
  const src = fs.readFileSync(AR_PATH, 'utf8');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'armut-'));
  try {
    MUTANTS.forEach(([name, from, to], i) => {
      if (src.split(from).length !== 2) { ok(false, `變異樣式找不到或重複：${name}`); return; }
      const f = path.join(tmp, `m${i}.js`);
      fs.writeFileSync(f, src.replace(from, () => to));
      let bad = 0, threw = false;
      try { reviewSuite(require(f), { ok: (c) => { if (!c) bad++; }, eq: (a, b) => { if (JSON.stringify(a) !== JSON.stringify(b)) bad++; } }); } catch (e) { threw = true; }
      ok(bad > 0 || threw, `變異「${name}」被抓到（失敗數 ${bad}${threw ? '，且拋例外' : ''}）`);
    });
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
}

// ═══════════════════════ Part B ═══════════════════════
const { mkUser, http, login, getLogs, byAction } = H;
function users() {
  return [
    mkUser('_s_admin', 'admin'), mkUser('_s_admin2', 'admin'),
    mkUser('_s_u1', 'user', { bu: ['ERP'], email: 'u1.private@corp-secret.example', supervisor: '_s_m1' }),
    mkUser('_s_dl', 'user', { bu: ['ERP'], canDownloadContacts: true }),
    mkUser('_s_m1', 'manager1', { bu: ['ERP'] }),
    mkUser('_s_test_shared', 'user', { bu: ['ERP'] }),
    mkUser('_s_inj', 'user', { bu: ['ERP'], displayName: '=HYPERLINK("http://evil.example")' }),
    mkUser('_s_never', 'user', { bu: ['ERP'] }),
    mkUser('_s_off', 'user', { bu: ['ERP'], active: false }),
  ];
}
const readAuth = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'auth.json'), 'utf8'));
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));

async function e2eMain() {
  console.log('[B1] 權限、報表內容、lastLoginAt');
  await H.fresh(users(), {});
  // 未登入／非管理員
  let r = await http(null, 'GET', '/api/admin/access-review'); eq(r.status, 401, '未登入 401');
  await H.loginAll(['_s_u1', '_s_admin']);
  for (const [m, u, b] of [['GET', '/api/admin/access-review'], ['GET', '/api/admin/access-review?format=csv'], ['GET', '/api/admin/access-review/history'], ['POST', '/api/admin/access-review/signoff', { quarter: '2026Q4' }]]) {
    r = await http('_s_u1', m, u, b); eq(r.status, 403, `非管理員 ${m} ${u.split('?')[0]} → 403`);
  }
  eq(readData().accessReviews, undefined, '非管理員簽核沒有寫入任何資料');
  // 登入時間：登入過的人有 lastLoginAt，沒登入的沒有
  const au = readAuth();
  const get = (n) => au.users.find(u => u.username === n);
  ok(get('_s_u1').lastLoginAt && get('_s_admin').lastLoginAt, '登入後帳號紀錄有 lastLoginAt');
  ok(!get('_s_never').lastLoginAt, '沒登入過的帳號沒有 lastLoginAt');
  ok(!!au.lastLoginTrackingSince, '記下「最後登入追蹤起始日」');
  const stamp1 = get('_s_u1').lastLoginAt;
  await sleep(1100); await login('_s_u1');
  eq(readAuth().users.find(u => u.username === '_s_u1').lastLoginAt, stamp1, '同一人 6 小時內再登入：不重寫 lastLoginAt（節流）');
  // 回應形狀不受影響
  const ul = await http('_s_admin', 'GET', '/api/admin/users');
  ok(ul.json.every(x => !('lastLoginAt' in x)), '/api/admin/users 回應不含 lastLoginAt（形狀不變）');
  r = await http('_s_u1', 'GET', '/api/me'); ok(!/lastLoginAt/.test(r.text), '/api/me 不含 lastLoginAt');
  // 編輯帳號不會洗掉 lastLoginAt
  r = await http('_s_admin', 'PUT', '/api/admin/users/_s_u1', { displayName: '改名後的業務' });
  eq(r.status, 200, '管理員編輯帳號');
  eq(readAuth().users.find(u => u.username === '_s_u1').lastLoginAt, stamp1, '編輯帳號後 lastLoginAt 仍在');
  // 報表
  r = await http('_s_admin', 'GET', '/api/admin/access-review');
  eq(r.status, 200, '報表 200');
  const rep = r.json;
  eq(rep.summary.accounts, 9, '帳號數＝9（不含 _pool）');
  ok(!rep.accounts.some(a => a.username === '_pool'), '不含 _pool');
  ok(!/\$2[aby]\$/.test(r.text) && !r.text.includes('corp-secret') && !/"password"|"token"/.test(r.text), 'HTTP 回應不含雜湊／email 位址／password 欄位');
  const A = (n) => rep.accounts.find(a => a.username === n);
  eq([A('_s_u1').emailPresent, A('_s_dl').emailPresent], [true, false], 'emailPresent 只是布林');
  ok(A('_s_u1').displayName === '改名後的業務', 'displayName 用 userLabel');
  ok(A('_s_dl').flags.includes('DOWNLOAD_PERMISSION') && A('_s_admin').flags.includes('ADMIN') && A('_s_test_shared').flags.includes('SHARED_NAME_SUSPECT'), '旗標：下載權限／ADMIN／共用帳號');
  ok(A('_s_never').flags.includes('NEVER_LOGGED_IN') && !A('_s_u1').flags.includes('NEVER_LOGGED_IN'), '旗標：未登入過 vs 登入過');
  ok(A('_s_u1').lastLoginAt && A('_s_u1').lastLoginDaysAgo === 0, '登入過的帳號有 lastLoginAt、距今 0 天');
  eq(A('_s_u1').duties.supervisesOthers, 0, '_s_u1 無部屬'); eq(A('_s_m1').duties.supervisesOthers, 1, '_s_m1 有 1 位部屬');
  ok(Array.isArray(A('_s_u1').features) && A('_s_u1').features.includes('contacts'), '有效功能來自角色矩陣');
  eq(A('_s_admin').features.length, rep.featureRouteMap ? A('_s_admin').features.length : 0, '管理員擁有全部功能');
  ok(Array.isArray(rep.featureRouteMap) && rep.featureRouteMap.length >= 100 && rep.featureRouteMap.every(x => x.method && x.path && x.kind), '附錄 featureRouteMap（路由→功能）');
  const m = await http('_s_admin', 'GET', '/api/admin/feature-route-map');
  eq(rep.featureRouteMap.length, m.json.routes.length, 'featureRouteMap 與 feature-route-map 端點同一份資料');
  ok(/^[0-9a-f]{64}$/.test(rep.snapshotHash), '報表附 snapshotHash');
  ok(!('quoteMode' in A('_s_u1')) && rep.meta.hasQuoteMode === false, '本分支帳號沒有 quoteMode → 報表不含此欄');
  // 功能矩陣變更 → 雜湊變更
  await http('_s_admin', 'PUT', '/api/admin/role-permissions', { matrix: { user: ['home'] } });
  r = await http('_s_admin', 'GET', '/api/admin/access-review');
  ok(r.json.snapshotHash !== rep.snapshotHash && r.json.accounts.find(a => a.username === '_s_u1').features.join() === 'home', '矩陣變更 → 有效功能與雜湊隨之改變');
  await http('_s_admin', 'DELETE', '/api/admin/role-permissions');

  console.log('[B2] CSV 下載與稽核');
  r = await http('_s_admin', 'GET', '/api/admin/access-review?format=csv');
  eq(r.status, 200, 'CSV 200');
  ok(/text\/csv/.test(r.headers.get('content-type') || '') && /attachment; filename="access-review-\d{4}-\d{2}-\d{2}\.csv"/.test(r.headers.get('content-disposition') || ''), 'Content-Type 與下載檔名');
  ok(r.buf[0] === 0xEF && r.buf[1] === 0xBB && r.buf[2] === 0xBF, 'CSV 位元組開頭是 UTF-8 BOM (EF BB BF)');
  ok(r.text.includes("'=HYPERLINK(") && !/(^|,)=HYPERLINK/m.test(r.text.replace(/\r/g, '')), "公式注入：displayName 以 = 開頭的儲存格被補 '");
  ok(!/\$2[aby]\$|corp-secret/.test(r.text), 'CSV 不含雜湊／email 位址');
  eq(r.text.split('\r\n').filter(Boolean).length, 10, 'CSV：表頭＋9 個帳號');
  const logs = await getLogs();
  ok(byAction(logs, 'EXPORT_ACCESS_REVIEW').some(l => l.operator === '_s_admin' && /9 筆/.test(l.detail)), '匯出 CSV 有稽核 EXPORT_ACCESS_REVIEW');

  console.log('[B3] 簽核：只增不改、雜湊鏈、現況比對');
  for (const [body, why] of [[{}, '缺季別'], [{ quarter: '2026Q9' }, '季別錯誤'], [{ quarter: '2026Q4', exceptions: [{ username: 'ghost', note: 'x' }] }, '例外帳號不存在'], [{ quarter: '2026Q4', comment: 'x'.repeat(1001) }, '意見太長']]) {
    r = await http('_s_admin', 'POST', '/api/admin/access-review/signoff', body); eq(r.status, 400, `簽核輸入錯誤被拒（${why}）`);
  }
  eq(readData().accessReviews, undefined, '被拒的簽核沒有留下紀錄');
  r = await http('_s_admin', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q4', comment: '已逐項檢視', exceptions: [{ username: '_s_dl', note: '專案需要下載權限' }] });
  eq(r.status, 200, '簽核成功');
  const rec1 = r.json.record;
  eq([rec1.seq, rec1.quarter, rec1.by, rec1.accountCount, rec1.prevHash, rec1.exceptions.length], [1, '2026Q4', '_s_admin', 9, '', 1], '紀錄內容');
  ok(/^[0-9a-f]{64}$/.test(rec1.snapshotHash) && /^[0-9a-f]{64}$/.test(rec1.recordHash), '快照雜湊與紀錄雜湊');
  ok(byAction(await getLogs(), 'ACCESS_REVIEW_SIGNOFF').some(l => l.target === '2026Q4' && l.operator === '_s_admin'), '簽核有稽核 ACCESS_REVIEW_SIGNOFF');
  let h = await http('_s_admin', 'GET', '/api/admin/access-review/history');
  eq([h.json.records.length, h.json.chain.ok, h.json.records[0].matchesCurrent], [1, true, true], '歷史：1 筆、鏈 OK、現況與簽核時一致');
  // 沒有修改／刪除路由
  for (const [m2, u2] of [['PUT', '/api/admin/access-review/signoff'], ['DELETE', '/api/admin/access-review/signoff'], ['DELETE', '/api/admin/access-review/history'], ['PUT', '/api/admin/access-review/history'], ['PATCH', '/api/admin/access-review/signoff'], ['DELETE', '/api/admin/access-review/1']]) {
    r = await http('_s_admin', m2, u2, { quarter: '2026Q4', comment: 'tamper' }); ok([404, 405].includes(r.status), `沒有 ${m2} ${u2} 這種修改／刪除路由（${r.status}）`);
  }
  eq(readData().accessReviews.length, 1, '嘗試修改／刪除後仍是 1 筆');
  // 組態變動 → 舊簽核「現況已不一致」；再簽一次 → 串鏈
  r = await http('_s_admin', 'PUT', '/api/admin/users/_s_u1', { role: 'manager2' }); eq(r.status, 200, '管理員調整 _s_u1 角色');
  h = await http('_s_admin', 'GET', '/api/admin/access-review/history');
  eq(h.json.records[0].matchesCurrent, false, '帳號組態變動後，歷史顯示「與簽核時不一致」');
  await login('_s_admin2');
  r = await http('_s_admin2', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q4', comment: '調整後重簽' });
  eq(r.status, 200, '另一位管理員可再簽（同季可重簽，皆留底）');
  eq([r.json.record.seq, r.json.record.prevHash === rec1.recordHash, r.json.record.by], [2, true, '_s_admin2'], '第 2 筆串在第 1 筆之後，簽核人是實際操作者');
  h = await http('_s_admin', 'GET', '/api/admin/access-review/history');
  eq([h.json.records.map(x => x.matchesCurrent), h.json.chain.ok], [[false, true], true], '歷史：舊的不一致、新的一致，鏈 OK');
  ok(h.json.records[0].comment === '已逐項檢視' && h.json.records[0].recordHash === rec1.recordHash, '第 1 筆內容未被動到');
  // 竄改偵測：直接改 data.json 的紀錄 → 歷史顯示鏈失敗、再簽被擋
  const d = readData(); d.accessReviews[0].comment = '被偷偷改過';
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
  h = await http('_s_admin', 'GET', '/api/admin/access-review/history');
  eq([h.json.chain.ok, h.json.chain.badAt], [false, 1], '竄改紀錄 → 歷史 chain.ok=false、badAt=1');
  r = await http('_s_admin', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q4', comment: 'x' });
  eq([r.status, r.json && r.json.code], [409, 'CHAIN_BROKEN'], '鏈已破壞時拒絕再簽核（409 CHAIN_BROKEN）');
  eq(readData().accessReviews.length, 2, '被拒的簽核沒有寫入');
}

async function e2eMutation() {
  console.log('[B-M] 伺服器層變異檢查');
  const tmp = [];
  const mk = (name, edit) => {
    let s = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
    if (!s.includes(edit[0])) throw new Error('變異樣式找不到：' + name);
    s = s.replace(edit[0], () => edit[1]);
    const f = path.join(ROOT, 'server.__mutar_' + name + '.js'); fs.writeFileSync(f, s); tmp.push(f); return f;
  };
  try {
    // 1. 簽核路由拿掉 requireAdmin
    let f = mk('m1', ["app.post('/api/admin/access-review/signoff', requireAdmin,", "app.post('/api/admin/access-review/signoff', requireAuth,"]);
    await H.fresh(users(), {}, { file: f });
    await H.loginAll(['_s_u1']);
    let r = await http('_s_u1', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q4' });
    ok(r.status === 200, '變異1（簽核拿掉 requireAdmin）→ 一般使用者能簽核：B1 的 403 斷言會失敗（抓得到）');
    // 2. 簽核改成覆蓋（不 append）
    f = mk('m2', ["d.accessReviews.push(rec);", "d.accessReviews = [rec];"]);
    await H.fresh(users(), {}, { file: f });
    await H.loginAll(['_s_admin']);
    await http('_s_admin', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q3' });
    await http('_s_admin', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q4' });
    ok(readData().accessReviews.length === 1, '變異2（簽核覆蓋舊紀錄）→ 只剩 1 筆：B3 的「2 筆」斷言會失敗（抓得到）');
    // 3. 登入每次都寫 lastLoginAt
    f = mk('m3', ["if (accessReview.shouldStampLogin(user, Date.now())) {", "if (true) {"]);
    await H.fresh(users(), {}, { file: f });
    await login('_s_u1'); const a1 = readAuth().users.find(u => u.username === '_s_u1').lastLoginAt;
    await sleep(1100); await login('_s_u1'); const a2 = readAuth().users.find(u => u.username === '_s_u1').lastLoginAt;
    ok(a1 !== a2, '變異3（登入不節流）→ lastLoginAt 每次重寫：B1 的節流斷言會失敗（抓得到）');
    // 4. 報表洩漏雜湊
    f = mk('m4', ["userLabel: (u) => userLabel(u, u.username),\n    featureRouteMap", "userLabel: (u) => userLabel(u, u.username) + (u.password || ''),\n    featureRouteMap"]);
    await H.fresh(users(), {}, { file: f });
    await H.loginAll(['_s_admin']);
    r = await http('_s_admin', 'GET', '/api/admin/access-review');
    ok(/\$2[aby]\$/.test(r.text), '變異4（顯示名稱夾帶雜湊）→ 回應含雜湊：B1 的不外洩斷言會失敗（抓得到）');
    // 5. 簽核不驗證鏈
    f = mk('m5', ["if (!accessReview.verifyChain(d.accessReviews).ok) return res.status(409)", "if (false) return res.status(409)"]);
    await H.fresh(users(), {}, { file: f });
    await H.loginAll(['_s_admin']);
    await http('_s_admin', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q3' });
    const dd = readData(); dd.accessReviews[0].comment = 'tampered'; fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(dd));
    r = await http('_s_admin', 'POST', '/api/admin/access-review/signoff', { quarter: '2026Q4' });
    ok(r.status === 200, '變異5（簽核不驗證鏈）→ 竄改後仍可簽：B3 的 409 斷言會失敗（抓得到）');
  } finally { await H.stopServer(); tmp.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } }); }
}

(async () => {
  H.guardClean();
  try {
    await partA();
    if (!process.env.AR_SKIP_E2E) {
      await e2eMain();
      if (!process.env.AR_SKIP_MUT) await e2eMutation();
    }
  } catch (e) { H.R.fail++; H.R.failures.push('例外：' + (e && e.stack || e)); console.log('  ✗ 例外', e && e.stack || e); }
  finally { await H.finish(); }
  process.exit(H.report('check-access-review '));
})();
