#!/usr/bin/env node
'use strict';
/**
 * 後端功能權限驗證（lib/featureGuard.js）檢查：
 *   Part A 靜態＋單元（不開伺服器）：解析原始碼中每一條 /api 路由，必須都在對照表（新增路由沒補表 → 失敗）；
 *          角色×路由矩陣（每個功能開／關）、log 絕不阻擋、enforce 才 403、節流、計數、金標準對照、變異檢查
 *   Part B 沙盒伺服器端對端（預設埠 3211）：log／off／enforce 三種模式、全部路由掃描、13 個角色、矩陣變更即時生效、
 *          集團角色白名單、唯讀、/api/admin/feature-route-map、已註冊路由偵測、伺服器層變異
 *
 *   FG_SKIP_E2E=1 只跑 Part A；FG_SKIP_MUT=1 略過變異
 * 守門與埠限制見 scripts/_b2-harness.js。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const H = require('./_b2-harness');
const { ROOT, ok, eq, sleep } = H;
const FG_PATH = path.join(ROOT, 'lib', 'featureGuard.js');

// ───────── 原始碼路由解析（偵測「新增路由卻沒補對照表」）─────────
function walk(dir, out) {
  for (const n of fs.readdirSync(dir, { withFileTypes: true })) {
    if (n.name === 'node_modules' || n.name.startsWith('.') || n.name === 'scratchpad') continue;
    const p = path.join(dir, n.name);
    if (n.isDirectory()) walk(p, out); else if (/\.js$/.test(n.name) && !/__mut/.test(n.name)) out.push(p);
  }
  return out;
}
/** 回傳 [{ method, path, admin, file }]；admin＝該路由宣告處（到第一個 => 或 function 為止）含 requireAdmin／qAdmin */
function parseRoutes(sources) {
  const out = [];
  const re = /\b(?:app|router)\.(get|post|put|patch|delete|all)\(\s*(['"`])(\/api\/[^'"`]*)\2/g;
  for (const { file, text } of sources) {
    let m;
    while ((m = re.exec(text))) {
      const rest = text.slice(m.index + m[0].length, m.index + m[0].length + 300);
      const head = rest.split(/=>|function\b/)[0];
      out.push({ method: m[1].toUpperCase(), path: m[3], admin: /requireAdmin|qAdmin/.test(head), file });
    }
  }
  return out;
}
function loadSources(extra) {
  const files = [path.join(ROOT, 'server.js')].concat(walk(path.join(ROOT, 'lib'), []), walk(path.join(ROOT, 'api'), []), walk(path.join(ROOT, 'middleware'), []));
  const srcs = files.map(f => ({ file: path.relative(ROOT, f).replace(/\\/g, '/'), text: fs.readFileSync(f, 'utf8') }));
  return extra ? srcs.concat(extra) : srcs;
}
function featureKeysFromServer() {
  const t = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const blk = t.slice(t.indexOf('const FEATURE_REGISTRY = ['), t.indexOf('const KNOWN_ROLES'));
  return Array.from(blk.matchAll(/key:\s*'([A-Za-z]+)'/g)).map(m => m[1]);
}
function knownRolesFromServer() {
  const t = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const m = t.match(/const KNOWN_ROLES = \[([^\]]+)\]/);
  return Array.from(m[1].matchAll(/'([a-z_0-9]+)'/g)).map(x => x[1]);
}
/** 未對照的非管理員路由 */
function uncovered(mod, srcs) {
  const have = new Set(mod.ROUTES.map(e => mod.keyOf(e.method, e.path)));
  return parseRoutes(srcs).filter(r => !r.admin && !have.has(mod.keyOf(r.method, r.path))).map(r => mod.keyOf(r.method, r.path));
}

// 金標準：人工確認的「路由 → 必須能涵蓋的功能」（獨立於對照表本身，變異對照表時用來抓）
const GOLDEN = [
  ['GET', '/api/receivables', 'accountingGroup'], ['POST', '/api/receivables', 'accountingGroup'], ['GET', '/api/contracts', 'contractGroup'],
  ['GET', '/api/exec/trend', 'execDash'], ['GET', '/api/exec/bu-comparison', 'execDash'], ['GET', '/api/manager-home', 'managerHome'],
  ['GET', '/api/leads', 'leads'], ['POST', '/api/leads/x/convert', 'leads'], ['POST', '/api/campaigns', 'campaigns'],
  ['GET', '/api/yoy/report', 'yoy'], ['POST', '/api/yoy/import', 'yoy'], ['POST', '/api/transfer-contacts', 'transfer'], ['GET', '/api/contacts-by-owner', 'transfer'],
  ['GET', '/api/companies', 'companyMaster'], ['GET', '/api/companies/export', 'companyMaster'], ['GET', '/api/pipeline-report', 'pipelineReport'],
  ['GET', '/api/manager/business-type-analysis', 'bizAnalysis'], ['GET', '/api/forecast/export', 'forecast'], ['GET', '/api/manager/achievement', 'targets'],
  ['POST', '/api/visits', 'visits'], ['DELETE', '/api/visits/x', 'visits'], ['POST', '/api/contacts', 'contacts'], ['DELETE', '/api/contacts/x', 'prospects'],
  ['PUT', '/api/opportunities/x', 'pipeline'], ['DELETE', '/api/opportunities/x', 'quotations'], ['GET', '/api/lost-opportunities', 'lostOpp'], ['POST', '/api/targets', 'targets'],
];
// 金標準的反面：這些功能「不得」開通該路由
const GOLDEN_DENY = [
  ['GET', '/api/receivables', 'contacts'], ['GET', '/api/exec/trend', 'home'], ['GET', '/api/leads', 'campaigns'], ['POST', '/api/transfer-contacts', 'targets'],
  ['GET', '/api/contracts', 'accountingGroup'], ['GET', '/api/manager-home', 'execDash'], ['GET', '/api/companies', 'contacts'], ['POST', '/api/visits', 'pipeline'],
];

// ═══════════════════════ Part A ═══════════════════════
// 仿 app.use('/api/', mw) 的掛載：Express 會把 baseUrl 設成 /api、path 去掉前綴
function mkReq(method, p, user) { return { method, baseUrl: p.startsWith('/api/') ? '/api' : '', path: p.startsWith('/api/') ? p.slice(4) : p, session: { user } }; }
function mkRes() { return { code: null, body: null, status(c) { this.code = c; return this; }, json(b) { this.body = b; return this; } }; }

function featureSuite(mod, T) {
  const { ok, eq } = T;
  const keys = featureKeysFromServer();
  const srcs = loadSources();

  // ── 路由涵蓋 ──
  const missing = uncovered(mod, srcs);
  eq(missing, [], '原始碼中每一條非管理員 /api 路由都在對照表');
  const parsed = parseRoutes(srcs);
  ok(parsed.length >= 190, `路由解析數量合理（${parsed.length} 條，應 ≥ 190）`);
  const srcKeys = new Set(parsed.map(r => mod.keyOf(r.method, r.path)));
  const stale = mod.ROUTES.filter(e => !srcKeys.has(mod.keyOf(e.method, e.path))).map(e => mod.keyOf(e.method, e.path));
  eq(stale, [], '對照表沒有已不存在的過期項目（路徑打錯會在此現形）');
  eq(new Set(mod.ROUTES.map(e => mod.keyOf(e.method, e.path))).size, mod.ROUTES.length, '對照表沒有重複項目');
  // 偵測器自己要有效：偷塞一條新路由（含多行寫法）必須被抓到
  const fake = [{ file: 'fake.js', text: "app.get('/api/zz-new-thing', requireAuth, (req,res)=>{});\napp.post(\n  '/api/zz-other/:id',\n  requireAuth, (req,res)=>{});\napp.get('/api/zz-admin-ok', requireAdmin, (req,res)=>{});" }];
  eq(uncovered(mod, srcs.concat(fake)).sort(), ['GET /api/zz-new-thing', 'POST /api/zz-other/:id'], '偵測器：新增（含多行寫法）的非管理員路由會被抓出，requireAdmin 的不算');
  // 規格健全
  const kinds = new Set(['ALWAYS', 'ANY', 'SELF', 'ADMIN', 'UNCLASSIFIED']);
  ok(mod.ROUTES.every(e => Array.isArray(e.spec) ? (e.spec.length > 0 && e.spec.every(k => keys.includes(k))) : kinds.has(e.spec)), '每個 features 都是 FEATURE_REGISTRY 內的鍵、kind 都合法');
  ok(mod.ROUTES.every(e => ['GET', 'POST', 'PUT', 'PATCH', 'DELETE'].includes(e.method) && e.path.startsWith('/api/')), '方法與路徑格式正確');

  // ── 比對器 ──
  eq(mod.matchRoute('GET', '/api/companies/export').path, '/api/companies/export', '靜態段優先：/companies/export 不被 /:id 吃掉');
  eq(mod.matchRoute('GET', '/api/companies/import-template').path, '/api/companies/import-template', '靜態段優先：import-template');
  eq(mod.matchRoute('GET', '/api/quotations/inbox').path, '/api/quotations/inbox', '靜態段優先：quotations/inbox');
  eq(mod.matchRoute('GET', '/api/companies/abc').path, '/api/companies/:id', ':id 比對');
  eq(mod.matchRoute('POST', '/api/opportunities/restore/abc').path, '/api/opportunities/restore/:id', '巢狀靜態段');
  eq(mod.matchRoute('GET', '/API/Receivables/'), mod.matchRoute('GET', '/api/receivables'), '大小寫不分、結尾斜線容許（Express 預設行為）');
  eq(mod.matchRoute('HEAD', '/api/receivables'), mod.matchRoute('GET', '/api/receivables'), 'HEAD 視同 GET');
  eq(mod.matchRoute('GET', '/api/zz-nothing'), null, '不存在的路徑回 null');
  eq(mod.matchRoute('PATCH', '/api/receivables'), null, '方法不同不命中');
  eq(mod.matchRoute('GET', '/api/receivables/a/b'), null, ':id 不吃斜線');

  // ── 金標準（獨立於表格本身）──
  const only = (k) => (role) => role === 'only' ? [k] : [];
  for (const [m, p, f] of GOLDEN) {
    const g = mod.create({ mode: 'enforce', getFeatures: only(f) });
    ok(g.decide(m, p, 'only').allow === true, `金標準：${f} 可用 ${m} ${p}`);
    const g0 = mod.create({ mode: 'enforce', getFeatures: () => [] });
    ok(g0.decide(m, p, 'x').allow === false, `金標準：沒有任何功能不可用 ${m} ${p}`);
  }
  for (const [m, p, f] of GOLDEN_DENY) {
    const g = mod.create({ mode: 'enforce', getFeatures: only(f) });
    ok(g.decide(m, p, 'only').allow === false, `金標準：只有 ${f} 不可用 ${m} ${p}`);
  }

  // ── 角色×路由矩陣：每個功能單獨開、全開、全關 ──
  const matrices = [['全關', []], ['全開', keys]].concat(keys.map(k => ['只開 ' + k, [k]]));
  let mismatches = 0, cells = 0, alwaysBlocked = 0;
  for (const [, feats] of matrices) {
    const g = mod.create({ mode: 'enforce', getFeatures: () => feats });
    for (const e of mod.ROUTES) {
      const concrete = e.path.replace(/:[A-Za-z]+/g, 'x1');
      const d = g.decide(e.method, concrete, 'any');
      const kind = mod.kindOf(e);
      const expect = kind === 'FEATURES' ? e.spec.some(f => feats.includes(f)) : kind === 'ANY' ? feats.length > 0 : true;
      cells++;
      if (d.allow !== expect) mismatches++;
      if (['ALWAYS', 'SELF', 'ADMIN', 'UNCLASSIFIED'].includes(kind) && !d.allow) alwaysBlocked++;
    }
  }
  eq(mismatches, 0, `矩陣：${matrices.length} 種功能組合 × ${mod.ROUTES.length} 條路由（${cells} 格）判定與規格一致`);
  eq(alwaysBlocked, 0, 'ALWAYS／SELF／ADMIN／UNCLASSIFIED 在任何功能組合下都不被擋');
  ok(mod.create({ mode: 'enforce', getFeatures: () => [] }).decide('GET', '/api/admin/users', 'x').allow, '/api/admin/* 未列入對照表者交給 requireAdmin（守門放行）');

  // ── 模式行為（假 req/res）──
  const fakeAudit = () => { const a = []; return { a, fn: (req, action, op, target, detail) => a.push({ action, op, target, detail }) }; };
  const run = (g, method, p, user) => { const req = mkReq(method, p, user === undefined ? { username: 'u1', role: 'r' } : user); const res = mkRes(); let n = 0; g.middleware(req, res, () => n++); return { n, res }; };
  {
    const A = fakeAudit();
    const clk = { t: 1_700_000_000_000 };
    const g = mod.create({ mode: 'log', getFeatures: () => [], audit: A.fn, clock: () => clk.t });
    let r = run(g, 'GET', '/api/receivables');
    eq([r.n, r.res.code], [1, null], 'log 模式：沒權限也放行');
    eq(g.snapshot().wouldDeny, 1, 'log 模式：wouldDeny 計數 +1');
    eq(A.a.map(x => x.action), ['FEATURE_GUARD_WOULD_DENY'], 'log 模式：寫 FEATURE_GUARD_WOULD_DENY');
    for (let i = 0; i < 5; i++) run(g, 'GET', '/api/receivables');
    eq([g.snapshot().wouldDeny, A.a.length], [6, 1], '同人同路由 10 分鐘內：計數照算但稽核只記一筆（節流）');
    run(g, 'GET', '/api/contracts');
    eq(A.a.length, 2, '不同路由各記一筆');
    run(g, 'GET', '/api/receivables', { username: 'u2', role: 'r' });
    eq(A.a.length, 3, '不同人各記一筆');
    clk.t += 10 * 60000 + 1;
    run(g, 'GET', '/api/receivables');
    eq(A.a.length, 4, '10 分鐘後再記一筆');
    ok(/receivables/.test(A.a[0].detail) && A.a[0].op === 'u1' && A.a[0].target === 'GET /api/receivables', '稽核內容含路由、操作者');
    r = run(g, 'GET', '/api/me'); eq([r.n, g.snapshot().wouldDeny], [1, 9], 'ALWAYS 路由不計入');
    r = run(g, 'GET', '/api/receivables', null); eq([r.n, g.snapshot().wouldDeny], [1, 9], '未登入不介入（交給 requireAuth）');
    r = run(g, 'GET', '/index.html'); eq(r.n, 1, '非 /api/ 不介入');
    r = run(g, 'OPTIONS', '/api/receivables'); eq(r.n, 1, 'OPTIONS 不介入');
  }
  {
    const A = fakeAudit();
    const g = mod.create({ mode: 'enforce', getFeatures: () => ['contacts'], audit: A.fn });
    let r = run(g, 'GET', '/api/receivables');
    eq([r.n, r.res.code], [0, 403], 'enforce：沒權限 → 403 且不放行');
    eq(r.res.body, { error: '您的角色沒有此功能的權限', code: 'FEATURE_DENIED', feature: ['accountingGroup'] }, 'enforce：回應含 code 與所需 feature');
    eq([g.snapshot().denied, A.a[0].action], [1, 'FEATURE_GUARD_DENY'], 'enforce：denied 計數與稽核');
    run(g, 'GET', '/api/receivables'); run(g, 'GET', '/api/receivables');
    eq([g.snapshot().denied, A.a.length], [3, 1], 'enforce：稽核節流、計數不節流');
    r = run(g, 'GET', '/api/contacts'); eq([r.n, r.res.code], [1, null], 'enforce：有權限放行');
    r = run(g, 'POST', '/api/contacts/abc/../x'); // 無對照 → NOMAP 放行並記一次
    ok(r.n === 1, 'enforce：對照表沒有的路徑放行（記一次 UNCLASSIFIED）');
    ok(A.a.some(x => x.action === 'FEATURE_GUARD_UNCLASSIFIED'), '未對照路由記 FEATURE_GUARD_UNCLASSIFIED');
    const before = A.a.length; run(g, 'POST', '/api/contacts/abc/../x'); eq(A.a.length, before, 'UNCLASSIFIED 同一路由只記一次');
    r = run(g, 'DELETE', '/api/yoy/2099'); eq([r.n, A.a.filter(x => x.action === 'FEATURE_GUARD_UNCLASSIFIED' && /yoy/.test(x.target)).length], [1, 1], '對照表標 UNCLASSIFIED 的路由放行並記一次');
    run(g, 'DELETE', '/api/yoy/2098');
    eq(A.a.filter(x => x.action === 'FEATURE_GUARD_UNCLASSIFIED' && /yoy/.test(x.target)).length, 1, 'UNCLASSIFIED 項目再呼叫不重複記');
    const g2 = mod.create({ mode: 'enforce', getFeatures: () => [], audit: A.fn });
    r = run(g2, 'GET', '/api/contacts'); eq(r.res.code, 403, 'ANY：零功能的角色連預載清單也不行');
    r = run(g2, 'GET', '/api/poll-bundle'); eq(r.n, 1, 'ALWAYS：零功能的角色仍可輪詢通知');
    r = run(g2, 'GET', '/api/quotations'); eq(r.n, 1, 'SELF：報價單路由交給 quoteRoutes');
  }
  {
    const g = mod.create({ mode: 'off', getFeatures: () => [] });
    const r = run(g, 'GET', '/api/receivables'); eq([r.n, g.snapshot().wouldDeny], [1, 0], 'off：完全不介入不計數');
  }
  eq([mod.readMode({}), mod.readMode({ FEATURE_GUARD_MODE: 'ENFORCE' }), mod.readMode({ FEATURE_GUARD_MODE: 'off' }), mod.readMode({ FEATURE_GUARD_MODE: 'garbage' }), mod.readMode({ FEATURE_GUARD_MODE: ' log ' })], ['log', 'enforce', 'off', 'log', 'log'], '模式解析：預設 log、不分大小寫、無效值退回 log');
  const desc = mod.describe(keys.map(k => ({ key: k, label: 'L-' + k })));
  ok(desc.length === mod.ROUTES.length && desc.filter(x => x.kind === 'FEATURES').every(x => x.featureLabels.length === x.features.length && x.featureLabels[0].startsWith('L-')), 'describe 輸出含功能中文名');
}

const MUTANTS = [
  ['enforce 不阻擋', "return res.status(403).json({ error: '您的角色沒有此功能的權限'", "return next(); res.status(403).json({ error: '您的角色沒有此功能的權限'"],
  ['log 也阻擋', "if (mode === 'log') {", "if (false) {"],
  ['ANY 不檢查功能數', "return { allow: feats.length > 0, kind, features: ['(任一功能)'], entry };", "return { allow: true, kind, features: ['(任一功能)'], entry };"],
  ['ANY-of 變 ALL-of', "allow: entry.spec.some(f => feats.includes(f))", "allow: entry.spec.every(f => feats.includes(f))"],
  ['ALWAYS/SELF 也檢查功能', "if (kind !== 'FEATURES' && kind !== 'ANY') return { allow: true, kind, entry };", "if (kind !== 'FEATURES') return { allow: feats0(), kind, entry }; function feats0() { return (getFeatures(role) || []).length > 0; }"],
  ['稽核不節流', "if (!throttled('W|' + su.username + '|' + key, t)) {", "if (true) {"],
  ['enforce 稽核不節流', "if (!throttled('D|' + su.username + '|' + key, t)) {", "if (true) {"],
  ['未登入也檢查', "const su = req.session && req.session.user;\n    if (!su || !su.username) return next();", "const su = (req.session && req.session.user) || { username: 'anon', role: 'none' };"],
  ['路徑比對區分大小寫', "'/?$', 'i')", "'/?$')"],
  ['靜態段不優先', "(b.statics - a.statics) || (a.params - b.params)", "(a.statics - b.statics) || (b.params - a.params)"],
  ['off 仍計數', "if (mode === 'off') return next();", ""],
  ['未對照路由被擋', "if (/^\\/api\\/admin\\//i.test(path)) return { allow: true, kind: 'ADMIN', entry: null };\n      return { allow: true, kind: 'NOMAP', entry: null };", "return { allow: false, kind: 'NOMAP', entry: null };"],
  ['對照表：應收帳款改成合約功能', "add('GET', '/api/receivables', ['accountingGroup']);", "add('GET', '/api/receivables', ['contractGroup']);"],
  ['對照表：移除 Lead 管理', "add('GET', '/api/leads', ['leads']);", ""],
  ['對照表：高階儀表板放寬成 ANY', "add('GET', '/api/exec/trend', ['execDash']);", "add('GET', '/api/exec/trend', 'ANY');"],
  ['對照表：名單移轉改 ALWAYS', "add('POST', '/api/transfer-contacts', ['transfer']);", "add('POST', '/api/transfer-contacts', 'ALWAYS');"],
  ['對照表：路徑打錯', "add('GET', '/api/leads', ['leads']);", "add('GET', '/api/lead', ['leads']);"],
];

async function partA() {
  console.log('[A] 靜態＋單元');
  featureSuite(require(FG_PATH), { ok, eq });
  // 版本頁的安全開關區塊（admin.html 內嵌函式，vm 執行；值一律跳脫）
  {
    const html = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
    const extractFn = (name) => { const i = html.indexOf('function ' + name + '('); let j = html.indexOf('{', i), d = 0, k = j; for (; k < html.length; k++) { if (html[k] === '{') d++; else if (html[k] === '}') { d--; if (!d) break; } } return html.slice(i, k + 1); };
    const UI = new Function([extractFn('escH'), extractFn('verGuardHtml'), 'return { verGuardHtml };'].join('\n'))();
    eq(UI.verGuardHtml({}), '', '舊版 API 缺欄位：不丟錯、不輸出');
    const log = UI.verGuardHtml({ featureGuard: { mode: 'log', wouldDeny: 7, denied: 0 }, loginGuard: { lockout: false, idleMinutes: 0 } });
    ok(/試行中/.test(log) && /若強制執行會被擋」7 次/.test(log) && /帳號鎖定：關閉/.test(log) && /閒置逾時：關閉/.test(log), 'log 模式顯示「試行中」與計數；兩個開關關閉');
    ok(/已強制執行/.test(UI.verGuardHtml({ featureGuard: { mode: 'enforce', wouldDeny: 0, denied: 3 } })), 'enforce 顯示已強制執行');
    ok(/閒置逾時 30 分鐘/.test(UI.verGuardHtml({ loginGuard: { lockout: true, maxFails: 5, lockMinutes: 15, idleMinutes: 30 } })) , '閒置逾時啟用時顯示分鐘數');
    const evil = '<img src=x onerror=alert(1)>';
    const x = UI.verGuardHtml({ featureGuard: { mode: evil, wouldDeny: evil, denied: evil }, loginGuard: { lockout: true, maxFails: evil, lockMinutes: evil, idleMinutes: evil } });
    ok(!/<img/i.test(x) && !/onerror=alert/.test(x.replace(/&lt;[^&]*&gt;/g, '')), '惡意字串全部跳脫（不產生 <img>）');
    ok(['FEATURE_GUARD_WOULD_DENY', 'FEATURE_GUARD_DENY', 'FEATURE_GUARD_UNCLASSIFIED'].every(k => html.includes(k + ':')), '稽核動作中文標籤已加入 ACTION_LABELS');
  }
  if (process.env.FG_SKIP_MUT) return;
  console.log('[A-M] 變異檢查（每個變異都必須讓檢查失敗）');
  const src = fs.readFileSync(FG_PATH, 'utf8');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'fgmut-'));
  try {
    MUTANTS.forEach(([name, from, to], i) => {
      if (src.split(from).length !== 2) { ok(false, `變異樣式找不到或重複：${name}`); return; }
      const f = path.join(tmp, `m${i}.js`);
      fs.writeFileSync(f, src.replace(from, () => to));
      let bad = 0, threw = false;
      try { featureSuite(require(f), { ok: (c) => { if (!c) bad++; }, eq: (a, b) => { if (JSON.stringify(a) !== JSON.stringify(b)) bad++; } }); } catch (e) { threw = true; }
      ok(bad > 0 || threw, `變異「${name}」被抓到（失敗數 ${bad}${threw ? '，且拋例外' : ''}）`);
    });
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
}

// ═══════════════════════ Part B ═══════════════════════
const { mkUser, http, login, getLogs, byAction } = H;
const FG = require(FG_PATH);
const ROLES = knownRolesFromServer();
const FEATURE_KEYS = featureKeysFromServer();

// 全域限流 300/分鐘（apiLimiter，在守門之前）：自己控速，避免 429 蓋掉守門的回應
const stamps = [];
async function paced(fn) {
  for (;;) {
    const now = Date.now();
    while (stamps.length && now - stamps[0] > 60000) stamps.shift();
    if (stamps.length < 230) break;
    await sleep(stamps[0] + 60000 - now + 50);
  }
  stamps.push(Date.now());
  return fn();
}
const ph = (who, m, u, b, o) => paced(() => http(who, m, u, b, o));
const concrete = (p) => p.replace(/:year/, '2099').replace(/:[A-Za-z]+/g, 'x1');

function users() {
  const u = [mkUser('_s_admin', 'admin'), mkUser('_s_u1', 'user', { bu: ['ERP'] }), mkUser('_s_sec', 'user', { bu: ['資安'] }),
    mkUser('_s_mk', 'marketing', { bu: ['ERP'] }), mkUser('_s_view', 'marketing', { bu: ['ERP'], accessMode: 'view' })];
  for (const r of ROLES) {
    if (['admin', 'user'].includes(r)) continue;
    const ex = (r === 'tecopm' || r === 'groupsales') ? { bu: [], viewOwnerScope: '_s_u1', viewGroupId: 'g1' } : { bu: ['ERP'] };
    u.push(mkUser('_r_' + r, r, ex));
  }
  u.push(mkUser('_r_user', 'user', { bu: ['ERP'] }));
  return u;
}
const DATA = () => ({ groups: [{ id: 'g1', name: '沙盒集團', memberCompanies: [] }], contacts: [{ id: 'c1', name: '測試甲', company: '沙盒公司甲', owner: '_s_u1', bu: 'ERP', createdAt: new Date().toISOString() }] });

async function setMatrix(m) { const r = await ph('_s_admin', 'PUT', '/api/admin/role-permissions', { matrix: m }); eq(r.status, 200, '管理員更新功能矩陣'); }
const sweepRoutes = () => FG.ROUTES.filter(e => ['FEATURES', 'ANY'].includes(FG.kindOf(e)));
const BODY = (m) => (m === 'GET' ? undefined : {});

/** 用 who 掃全部 FEATURES/ANY 路由；expectedAllow(entry)→bool；回傳不一致清單 */
async function sweep(who, expectedAllow, label) {
  const bad = [];
  for (const e of sweepRoutes()) {
    const r = await ph(who, e.method, concrete(e.path), BODY(e.method));
    const denied = r.status === 403 && r.json && r.json.code === 'FEATURE_DENIED';
    const isAi = e.path.startsWith('/api/ai/');   // aiLimiter(每分鐘 5 次) 與沒有 Gemini 金鑰的 503 都發生在守門之後，不代表被擋
    if (r.status === 429 && !isAi) { bad.push(`${e.method} ${e.path} 被限流 429`); continue; }
    if (r.status >= 500 && !(isAi && r.status === 503)) bad.push(`${e.method} ${e.path} → ${r.status}`);
    if (denied === expectedAllow(e)) bad.push(`${e.method} ${e.path} 預期${expectedAllow(e) ? '放行' : '擋下'}卻${denied ? '被擋' : '放行'}（${r.status}）`);
  }
  eq(bad, [], `${label}：${sweepRoutes().length} 條路由的實際行為與對照表一致`);
}
const allowFor = (feats) => (e) => FG.kindOf(e) === 'ANY' ? feats.length > 0 : e.spec.some(f => feats.includes(f));

async function e2eLog() {
  console.log('[B1] log 模式（預設）：永不阻擋，但記錄');
  await H.fresh(users(), DATA());   // 未設 FEATURE_GUARD_MODE
  await H.loginAll(['_s_admin', '_s_mk']);
  await setMatrix({ marketing: [] });
  const bad = [];
  for (const e of sweepRoutes()) { const r = await ph('_s_mk', e.method, concrete(e.path), BODY(e.method)); if (r.json && r.json.code === 'FEATURE_DENIED') bad.push(e.method + ' ' + e.path); }
  eq(bad, [], 'log 模式：零功能角色打遍所有路由，沒有任何一條被 FEATURE_DENIED');
  const v = await ph('_s_admin', 'GET', '/api/admin/version');
  eq(v.json.featureGuard.mode, 'log', '版本頁 featureGuard.mode=log');
  ok(v.json.featureGuard.wouldDeny >= sweepRoutes().length - 5, `wouldDeny 計數（${v.json.featureGuard.wouldDeny}）≈ 被掃的路由數`);
  eq(v.json.featureGuard.denied, 0, 'denied 計數為 0');
  const logs = await getLogs();
  const wd = byAction(logs, 'FEATURE_GUARD_WOULD_DENY');
  ok(wd.length >= sweepRoutes().length - 5 && wd.every(l => l.operator === '_s_mk'), `FEATURE_GUARD_WOULD_DENY 每路由一筆（${wd.length}）`);
  const targets = wd.map(l => l.target); eq(new Set(targets).size, targets.length, '同人同路由沒有重複稽核（節流）');
  // 再打一輪，稽核筆數不增加
  const n0 = wd.length;
  await ph('_s_mk', 'GET', '/api/receivables'); await ph('_s_mk', 'GET', '/api/contracts');
  eq(byAction(await getLogs(), 'FEATURE_GUARD_WOULD_DENY').length, n0, '10 分鐘內重打不再記錄');
  // 管理員以一般使用者身分用業務 API：不受影響
  const r = await ph('_s_admin', 'GET', '/api/receivables'); ok(!(r.json && r.json.code === 'FEATURE_DENIED'), '管理員打業務 API 無事');
  eq(v.json.featureGuard.unclassified >= 0, true, 'unclassified 計數存在');
  // 路由表端點
  const m = await ph('_s_admin', 'GET', '/api/admin/feature-route-map');
  eq(m.status, 200, 'feature-route-map 可讀');
  eq(m.json.unmappedRegistered, [], '已註冊的非管理員 /api 路由（Express 路由堆疊）全都在對照表');
  eq(m.json.routes.length, FG.ROUTES.length, '輸出路由數與對照表一致');
  ok(m.json.routes.find(x => x.path === '/api/receivables' && x.method === 'GET').featureLabels[0] === '帳務管理', '輸出含功能中文名');
  const u = await ph('_s_mk', 'GET', '/api/admin/feature-route-map'); eq(u.status, 403, '非管理員不可讀路由表');
  // 堆疊與原始碼解析數量交叉驗證
  const src = parseRoutes(loadSources()).filter(r => !r.admin).length;
  ok(m.json.routes.length >= 100 && src >= m.json.routes.length - 3, `原始碼非管理員路由 ${src} 條 vs 對照表 ${m.json.routes.length} 條（對照表可多含 ADMIN／自定項目）`);
}

async function e2eOff() {
  console.log('[B2] off 模式');
  await H.fresh(users(), DATA(), { env: { FEATURE_GUARD_MODE: 'off' } });
  await H.loginAll(['_s_admin', '_s_mk']);
  await setMatrix({ marketing: [] });
  const r = await ph('_s_mk', 'GET', '/api/receivables');
  ok(!(r.json && r.json.code === 'FEATURE_DENIED'), 'off：不阻擋');
  const v = await ph('_s_admin', 'GET', '/api/admin/version');
  eq([v.json.featureGuard.mode, v.json.featureGuard.wouldDeny], ['off', 0], 'off：不計數');
  eq(byAction(await getLogs(), 'FEATURE_GUARD_WOULD_DENY').length, 0, 'off：無稽核');
}

async function e2eEnforce() {
  console.log('[B3] enforce 模式：掃描全部路由（多種功能矩陣）');
  await H.fresh(users(), DATA(), { env: { FEATURE_GUARD_MODE: 'enforce' } });
  await H.loginAll(['_s_admin', '_s_mk', '_s_view', '_s_sec']);
  // 矩陣變體：全關、只開 contacts、只開 quotations、只開 accountingGroup、全開
  const variants = [['全關', []], ['只開 contacts', ['contacts']], ['只開 quotations', ['quotations']], ['只開 accountingGroup', ['accountingGroup']], ['全開', FEATURE_KEYS]];
  for (const [label, feats] of variants) {
    await setMatrix({ marketing: feats });   // 立即生效，不需重新登入
    await sweep('_s_mk', allowFor(feats), `marketing ${label}`);
  }
  // 拒絕時的回應格式與稽核
  await setMatrix({ marketing: ['contacts'] });
  let r = await ph('_s_mk', 'GET', '/api/receivables');
  eq([r.status, r.json && r.json.code, r.json && r.json.feature], [403, 'FEATURE_DENIED', ['accountingGroup']], 'enforce：403＋code＋所需功能');
  ok(/application\/json/.test(r.headers.get('content-type') || ''), '回應為 JSON');
  await ph('_s_mk', 'GET', '/api/receivables'); await ph('_s_mk', 'GET', '/api/receivables');
  let logs = await getLogs();
  eq(byAction(logs, 'FEATURE_GUARD_DENY').filter(l => l.target === 'GET /api/receivables').length, 1, 'FEATURE_GUARD_DENY 同人同路由 10 分鐘只記一筆');
  let v = await ph('_s_admin', 'GET', '/api/admin/version');
  ok(v.json.featureGuard.denied >= sweepRoutes().length, `denied 計數 ${v.json.featureGuard.denied}`);
  // 矩陣變更立即生效（同一個 cookie）：授予後馬上可用、收回後馬上不可用
  await setMatrix({ marketing: ['contacts', 'accountingGroup'] });
  r = await ph('_s_mk', 'GET', '/api/receivables'); ok(!(r.json && r.json.code === 'FEATURE_DENIED'), '授予 accountingGroup 後立即放行');
  await setMatrix({ marketing: ['contacts'] });
  r = await ph('_s_mk', 'GET', '/api/receivables'); eq(r.json && r.json.code, 'FEATURE_DENIED', '收回後立即擋下');
  // 重設為預設值也生效
  r = await ph('_s_admin', 'DELETE', '/api/admin/role-permissions'); eq(r.status, 200, '重置矩陣為預設');
  r = await ph('_s_mk', 'GET', '/api/leads'); ok(!(r.json && r.json.code === 'FEATURE_DENIED'), '預設矩陣：marketing 有 leads → 放行');
  r = await ph('_s_mk', 'GET', '/api/receivables'); eq(r.json && r.json.code, 'FEATURE_DENIED', '預設矩陣：marketing 沒有 accountingGroup → 擋下');
  // 管理員：以一般使用者身分用業務 API 全部放行；管理 API 由 requireAdmin
  await sweep('_s_admin', () => true, 'admin（擁有全部功能）');
  r = await ph('_s_mk', 'GET', '/api/admin/users'); eq(r.status, 403, '非管理員打後台 API 仍由 requireAdmin 403');
  ok(!(r.json && r.json.code === 'FEATURE_DENIED'), '…且不是 FEATURE_DENIED');
  // 唯讀使用者：功能判斷與唯讀判斷各自獨立
  await setMatrix({ marketing: ['contacts'] });
  r = await ph('_s_view', 'GET', '/api/receivables'); eq(r.json && r.json.code, 'FEATURE_DENIED', '唯讀帳號沒有功能 → 讀取也被擋');
  r = await ph('_s_view', 'POST', '/api/contacts', { name: 'x' });
  eq(r.status, 403, '唯讀帳號寫入仍由唯讀機制擋下');
  ok(/唯讀/.test((r.json && r.json.error) || ''), '…訊息是唯讀（既有行為不變）');
  // 資安 BU 使用者與一般使用者一樣依角色判斷
  r = await ph('_s_sec', 'GET', '/api/receivables'); ok(!(r.json && r.json.code === 'FEATURE_DENIED'), '資安 BU 的 user 角色：預設有 accountingGroup → 放行');
  await setMatrix({ user: ['home'] });
  r = await ph('_s_sec', 'GET', '/api/receivables'); eq(r.json && r.json.code, 'FEATURE_DENIED', '資安 BU 使用者同樣依矩陣被擋');
  r = await ph('_s_sec', 'GET', '/api/poll-bundle'); ok(!(r.json && r.json.code === 'FEATURE_DENIED'), '通知輪詢仍可用');
  r = await ph('_s_sec', 'GET', '/api/me/permissions'); eq(r.status, 200, '/api/me/permissions 仍可用');
  await ph('_s_admin', 'DELETE', '/api/admin/role-permissions');
  // ALWAYS／SELF 類 GET 路由：任何角色（零功能）都不被守門擋
  await setMatrix({ marketing: [] });
  const bad = [];
  for (const e of FG.ROUTES.filter(x => ['ALWAYS', 'SELF', 'UNCLASSIFIED'].includes(FG.kindOf(x)) && x.method === 'GET' && !/cron|vapid|sw\.js/.test(x.path))) {
    const rr = await ph('_s_mk', 'GET', concrete(e.path));
    if (rr.json && rr.json.code === 'FEATURE_DENIED') bad.push(e.path);
  }
  eq(bad, [], '零功能角色打 ALWAYS／SELF／UNCLASSIFIED 的 GET 路由，沒有任何一條被守門擋');
  await ph('_s_admin', 'DELETE', '/api/admin/role-permissions');
  // UNCLASSIFIED 只記一次
  await ph('_s_u1', 'GET', '/api/contract-reminders'); await ph('_s_u1', 'GET', '/api/contract-reminders');
  logs = await getLogs();
  eq(byAction(logs, 'FEATURE_GUARD_UNCLASSIFIED').filter(l => /contract-reminders/.test(l.target)).length, 1, 'UNCLASSIFIED 路由只記錄一次');
}

async function e2eRoles() {
  console.log('[B4] enforce 模式：13 個角色 × 代表性路由（預設矩陣）');
  await H.fresh(users(), DATA(), { env: { FEATURE_GUARD_MODE: 'enforce' } });
  await H.loginAll(['_s_admin']);
  // 代表路由：每一種「功能組合」取一條（優先 GET）
  const seen = new Map();
  for (const e of sweepRoutes()) { const k = FG.kindOf(e) + ':' + (Array.isArray(e.spec) ? e.spec.slice().sort().join(',') : ''); if (!seen.has(k) || (e.method === 'GET' && seen.get(k).method !== 'GET')) seen.set(k, e); }
  const reps = Array.from(seen.values());
  let cells = 0, skipped = 0; const bad = [];
  const roles = ROLES.filter(r => r !== 'pool');
  for (const role of roles) {
    const who = role === 'admin' ? '_s_admin' : (role === 'user' ? '_r_user' : '_r_' + role);
    if (role !== 'admin') { const lr = await login(who); if (lr.status !== 200) { bad.push(role + ' 登入失敗 ' + lr.status); continue; } }
    const pm = await ph(who, 'GET', '/api/me/permissions');
    const feats = pm.json && pm.json.features || [];
    const g = FG.create({ mode: 'enforce', getFeatures: () => feats });
    const grouped = role === 'tecopm' || role === 'groupsales';
    for (const e of reps) {
      const r = await ph(who, e.method, concrete(e.path), BODY(e.method));
      cells++;
      const denied = r.status === 403 && r.json && r.json.code === 'FEATURE_DENIED';
      const expectAllow = g.decide(e.method, concrete(e.path), role).allow;
      if (grouped && r.status === 403 && !(r.json && r.json.code)) { skipped++; continue; }   // 集團角色路徑白名單／唯讀先擋（既有機制）
      if ((r.status === 429 || r.status >= 500) && !(e.path.startsWith('/api/ai/') && [429, 503].includes(r.status))) { bad.push(`${role} ${e.method} ${e.path} → ${r.status}`); continue; }
      if (denied === expectAllow) bad.push(`${role} ${e.method} ${e.path} 預期${expectAllow ? '放行' : '擋下'}卻${denied ? '被擋' : '放行'}`);
    }
  }
  eq(bad, [], `${roles.length} 個角色 × ${reps.length} 條代表路由（${cells} 格，集團角色白名單先擋者略過 ${skipped}）：enforce 行為與矩陣一致`);
  // 集團角色：白名單內的功能路由由功能矩陣把關；白名單機制本身不變
  await H.loginAll(['_r_tecopm', '_r_groupsales']);
  let r = await ph('_r_tecopm', 'GET', '/api/opportunities'); ok(r.status === 200, 'tecopm 讀商機（預設有 pipeline）仍 200（' + r.status + '）');
  r = await ph('_r_tecopm', 'GET', '/api/receivables'); eq([r.status, r.json && r.json.code], [403, undefined], 'tecopm 打非白名單路由：仍是原本的白名單 403（無 FEATURE_DENIED code）');
  r = await ph('_r_groupsales', 'GET', '/api/contacts'); ok(r.status === 200, 'groupsales 讀聯絡人仍 200');
  await setMatrix({ groupsales: ['forecast'] });
  r = await ph('_r_groupsales', 'GET', '/api/groups'); eq([r.status, r.json && r.json.code], [403, 'FEATURE_DENIED'], 'groupsales 被收回 contacts/prospects 後，白名單內的 /api/groups 由功能矩陣擋下');
  r = await ph('_r_groupsales', 'GET', '/api/me'); eq(r.status, 200, 'groupsales 仍可 /api/me');
}

async function e2eMutation() {
  console.log('[B-M] 伺服器層變異檢查');
  const tmp = [];
  const mk = (name, edit, libEdit) => {
    let s = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
    if (edit) { if (!s.includes(edit[0])) throw new Error('變異樣式找不到：' + name); s = s.replace(edit[0], () => edit[1]); }
    if (libEdit) {
      let l = fs.readFileSync(FG_PATH, 'utf8');
      if (!l.includes(libEdit[0])) throw new Error('變異樣式找不到：' + name);
      l = l.replace(libEdit[0], () => libEdit[1]);
      const lf = path.join(ROOT, 'lib', 'featureGuard.__mut_' + name + '.js'); fs.writeFileSync(lf, l); tmp.push(lf);
      s = s.replace("require('./lib/featureGuard')", () => "require('./lib/featureGuard.__mut_" + name + ".js')");
    }
    const f = path.join(ROOT, 'server.__mutfg_' + name + '.js'); fs.writeFileSync(f, s); tmp.push(f); return f;
  };
  try {
    // 1. 守門沒掛上
    let f = mk('m1', ["app.use('/api/', featureGuard.middleware);", '/* mut */']);
    await H.fresh(users(), DATA(), { env: { FEATURE_GUARD_MODE: 'enforce' }, file: f });
    await H.loginAll(['_s_admin', '_s_mk']); await setMatrix({ marketing: [] });
    let r = await ph('_s_mk', 'GET', '/api/receivables');
    ok(!(r.json && r.json.code === 'FEATURE_DENIED'), '變異1（守門沒掛上）→ enforce 擋不住：B3 的掃描會失敗（抓得到）');
    // 2. 用快取的舊矩陣（getFeatures 回固定值）
    f = mk('m2', ["getFeatures: (role) => getUserFeatures(role),", "getFeatures: (role) => ['home'],"]);
    await H.fresh(users(), DATA(), { env: { FEATURE_GUARD_MODE: 'enforce' }, file: f });
    await H.loginAll(['_s_admin', '_s_mk']); await setMatrix({ marketing: ['accountingGroup'] });
    r = await ph('_s_mk', 'GET', '/api/receivables');
    ok(r.json && r.json.code === 'FEATURE_DENIED', '變異2（不看真正的矩陣）→ 授予後仍被擋：B3 的「矩陣變更即時生效」會失敗（抓得到）');
    // 3. 路由表缺一條（新增路由未補表）→ 已註冊路由偵測要抓到
    f = mk('m3', null, ["add('GET', '/api/leads', ['leads']);", '']);
    await H.fresh(users(), DATA(), { env: { FEATURE_GUARD_MODE: 'log' }, file: f });
    await H.loginAll(['_s_admin']);
    r = await ph('_s_admin', 'GET', '/api/admin/feature-route-map');
    ok(r.json && r.json.unmappedRegistered.some(x => x.path === '/api/leads'), '變異3（對照表少一條）→ unmappedRegistered 列出 GET /api/leads（執行期偵測抓得到）');
    // 4. 全域 enforce 的 429 之類不影響：admin 被誤擋
    f = mk('m4', null, ["const feats = getFeatures(role) || [];", "const feats = role === 'admin' ? [] : (getFeatures(role) || []);"]);
    await H.fresh(users(), DATA(), { env: { FEATURE_GUARD_MODE: 'enforce' }, file: f });
    await H.loginAll(['_s_admin']);
    r = await ph('_s_admin', 'GET', '/api/receivables');
    ok(r.json && r.json.code === 'FEATURE_DENIED', '變異4（管理員被當成零功能）→ 管理員被擋：B3 的 admin 掃描會失敗（抓得到）');
  } finally { await H.stopServer(); tmp.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } }); }
}

(async () => {
  H.guardClean();
  try {
    await partA();
    if (!process.env.FG_SKIP_E2E) {
      const only = (process.env.FG_ONLY || '').split(',').filter(Boolean);
      const want = (k) => !only.length || only.includes(k);
      if (want('log')) await e2eLog();
      if (want('off')) await e2eOff();
      if (want('enforce')) await e2eEnforce();
      if (want('roles')) await e2eRoles();
      if (want('mut') && !process.env.FG_SKIP_MUT) await e2eMutation();
    }
  } catch (e) { H.R.fail++; H.R.failures.push('例外：' + (e && e.stack || e)); console.log('  ✗ 例外', e && e.stack || e); }
  finally { await H.finish(); }
  process.exit(H.report('check-feature-guard '));
})();
