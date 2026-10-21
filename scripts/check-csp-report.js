#!/usr/bin/env node
'use strict';
/**
 * CSP 第一期（Report-Only＋違規回報端點）檢查：
 *   Part A 單元測試（lib/cspReport.js：政策內容、清洗、去重、預算、限流、摘要）＋ 變異檢查（把規則改壞，測試必須抓得到）
 *   Part B 沙盒伺服器端對端（預設埠 3080，JSON 模式）：標頭、端點的大小／內容類型／清洗／去重／稽核寫入／未登入可用、
 *          不吃使用者 API 額度、CSP_REPORT_ONLY=off 時與基底（git archive）逐頁標頭位元級相同、on 時只多兩個標頭
 *
 *   node scripts/check-csp-report.js                 # 全部
 *   CSP_SKIP_E2E=1 node scripts/check-csp-report.js  # 只跑單元＋變異
 *   CSP_BASE_REF=<commit> ...                         # 位元級比對用的基底（預設釘死 1f1eeb0＝CSP 合併前的整併分支尖端）
 * 守門與埠限制見 scripts/_b2-harness.js（工作樹已有 auth/data/audit 檔時拒絕執行）。
 */
process.env.B2_PORT = process.env.B2_PORT || '3080';
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const os = require('os');
const path = require('path');
const Module = require('module');
const { spawn, execSync } = require('child_process');

const LIB = path.join(__dirname, '..', 'lib', 'cspReport.js');
const cspMod = require(LIB);

let upass = 0, ufail = 0; const ufailures = [];
const U = {
  ok(c, name) { if (c) upass++; else { ufail++; ufailures.push(name); } },
  eq(a, b, name) { U.ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); },
};

// ═══════════════════════ Part A：單元 ═══════════════════════
function mkClock(t0) { const c = { t: t0 || 1_700_000_000_000, now() { return c.t; }, adv(ms) { c.t += ms; } }; return c; }

function unitSuite(M, T) {
  const { ok, eq } = T;
  const mk = (env, mode) => {
    const clock = mkClock(); const logs = []; let holds = 0;
    const rep = M.create({ env: env || {}, now: clock.now, writeLog: (a, op, t, d, req) => logs.push({ a, op, t, d, req }), holdForFlush: () => { holds++; }, auditMode: () => mode || 'legacy' });
    return { rep, clock, logs, holds: () => holds };
  };
  const N = (o) => M.normOne(Object.assign({ directive: 'script-src-elem', blocked: 'inline', doc: 'https://crm.example/admin.html', src: 'https://crm.example/admin.html', line: 10, sample: 'abc', disposition: 'report' }, o || {}));

  // ── 政策內容 ──
  {
    const h = M.headersFor({});
    ok(h && h['Content-Security-Policy-Report-Only'], '預設（未設環境變數）＝on，有 Report-Only 標頭');
    const p = h['Content-Security-Policy-Report-Only'];
    const dirs = Object.fromEntries(p.split('; ').map(s => { const i = s.indexOf(' '); return i < 0 ? [s, ''] : [s.slice(0, i), s.slice(i + 1)]; }));
    eq(dirs['default-src'], "'self'", "default-src 'self'");
    eq(dirs['script-src'], "'self' https://cdn.jsdelivr.net", "script-src 只有 'self' 與 jsdelivr（沒有 unsafe-inline）");
    ok(!/unsafe-inline|unsafe-eval/.test(dirs['script-src']), 'script-src 不含 unsafe-*');
    eq(dirs['script-src-attr'], "'none'", "script-src-attr 'none'");
    eq(dirs['frame-ancestors'], "'none'", "frame-ancestors 'none'");
    eq(dirs['base-uri'], "'self'", "base-uri 'self'");
    eq(dirs['form-action'], "'self'", "form-action 'self'");
    eq(dirs['object-src'], "'none'", "object-src 'none'");
    eq(dirs['connect-src'], "'self'", "connect-src 'self'");
    eq(dirs['img-src'], "'self' data: blob:", 'img-src');
    eq(dirs['report-uri'], '/api/csp-report', 'report-uri');
    ok(!('report-to' in dirs), '預設不含 report-to（Chromium 見到 report-to 就忽略 report-uri，預設只用 report-uri）');
    ok(!('upgrade-insecure-requests' in dirs), 'Report-Only 不含 upgrade-insecure-requests（會被瀏覽器忽略並警告）');
    ok(!p.includes(','), '預設只有一條政策（沒有 style 嚴格條）');
    ok(!('Reporting-Endpoints' in h), '預設沒有 Reporting-Endpoints 標頭');
    const ht = M.headersFor({ CSP_REPORT_TO: 'on' });
    eq(ht['Reporting-Endpoints'], 'csp-endpoint="/api/csp-report"', 'CSP_REPORT_TO=on → Reporting-Endpoints');
    ok(ht['Content-Security-Policy-Report-Only'].includes('; report-uri /api/csp-report; report-to csp-endpoint'),'CSP_REPORT_TO=on → 政策同時有 report-uri 與 report-to');
    ok(!M.reportToEnabled({ CSP_REPORT_TO: 'off' }) && M.reportToEnabled({ CSP_REPORT_TO: '1' }), 'CSP_REPORT_TO 解析');
    const hs = M.headersFor({ CSP_REPORT_ONLY_STYLE: 'on' });
    const parts = hs['Content-Security-Policy-Report-Only'].split(', ');
    eq(parts.length, 2, 'CSP_REPORT_ONLY_STYLE=on 多一條政策');
    ok(/^style-src 'self'; report-uri/.test(parts[1]), '第二條只管 style-src \'self\'');
    ok(!M.styleReportEnabled({ CSP_REPORT_ONLY: 'off', CSP_REPORT_ONLY_STYLE: 'on' }), 'off 時 style 條也關');
    for (const v of ['off', 'OFF', '0', 'false', 'no']) eq(M.headersFor({ CSP_REPORT_ONLY: v }), null, `CSP_REPORT_ONLY=${v} → 不加標頭`);
    for (const v of ['on', '1', 'true', '', undefined, 'yes']) ok(!!M.headersFor({ CSP_REPORT_ONLY: v }), `CSP_REPORT_ONLY=${JSON.stringify(v)} → 加標頭`);
    // 中介層
    const mw = M.reportOnlyMiddleware({});
    const run = (p) => { const set = {}; let nx = 0; mw({ path: p }, { setHeader: (k, v) => { set[k] = v; } }, () => nx++); return { set, nx }; };
    ok(Object.keys(run('/login.html').set).length === 1 && run('/login.html').nx === 1, '/login.html 加 Report-Only 標頭並放行');
    eq(Object.keys(run('/api/me').set).length, 0, '/api/ 不加');
    eq(Object.keys(run('/q/abc').set).length, 0, '/q/ 跳板頁不加（自帶更嚴的 CSP）');
    ok(Object.keys(run('/app.js').set).length === 1, '/app.js（靜態檔）加');
    eq(M.reportOnlyMiddleware({ CSP_REPORT_ONLY: 'off' }), null, 'off 時沒有中介層');
  }

  // ── 清洗 ──
  {
    eq(M.pathOf('https://crm.example/admin.html?token=SECRET&x=1#frag'), '/admin.html', 'pathOf 去查詢字串與片段');
    eq(M.pathOf('/login.html?next=%2Fq%2Fabc'), '/login.html', '相對路徑去查詢字串');
    eq(M.pathOf('about:blank'), '(about)', '非 http(s) 只留 scheme');
    eq(M.pathOf('data:text/html,<script>alert(1)</script>'), '(data)', 'data: 只留 scheme');
    eq(M.pathOf(''), '', '空字串');
    eq(M.pathOf({}), '', '非字串');
    ok(!/[<>"' ]/.test(M.pathOf('https://x.example/a<script>alert("1")</script>b c')), 'path 內的 < > " \' 空白被替換');
    ok(M.pathOf('https://x.example/' + 'a'.repeat(5000)).length <= 120, 'path 限長 120');
    ok(!/[\u0000-\u001f]/.test(M.pathOf('/a\u0000b\nc')), 'path 去控制字元');
    eq(M.blockedOf('inline').v, 'inline', 'inline 關鍵字');
    eq(M.blockedOf('EVAL').v, 'eval', 'eval 關鍵字（大小寫）');
    eq(M.blockedOf('https://cdn.example.com/lib/x.js?key=SECRET').v, 'https://cdn.example.com', 'blocked 只留 origin（去 path 與 query）');
    eq(M.blockedOf('http://cdn.example.com:8080/a').v, 'http://cdn.example.com:8080', 'origin 含埠');
    eq(M.blockedOf('data:image/png;base64,AAAA').v, 'data', 'data: URI 只留關鍵字 data');
    eq(M.blockedOf('chrome-extension://abcdef/content.js').ext, true, '擴充套件標記 ext');
    eq(M.blockedOf('moz-extension://abc/x.js').ext, true, 'moz-extension 標記 ext');
    eq(M.blockedOf('').v, '(空)', '空 blocked');
    eq(M.blockedOf('https://evil.example\n.com/x').v === '(其他)' || /^https:\/\/[a-z0-9.-]+$/.test(M.blockedOf('https://evil.example\n.com/x').v), true, '含換行的網址不會原樣進入');
    eq(M.directiveOf("script-src-elem 'self' https://cdn"), 'script-src-elem', '取第一個 token');
    eq(M.directiveOf('<img onerror=1>'), '(未知指令)', '指令名稱白名單');
    eq(M.directiveOf(undefined), '(未知指令)', '指令缺');
    const n = N({ sample: 'SECRET_SCRIPT_BODY_'.repeat(20), line: 'abc' });
    eq(n.sampleLen, 'SECRET_SCRIPT_BODY_'.length * 20, 'sample 只留長度');
    ok(!JSON.stringify(n).includes('SECRET_SCRIPT_BODY'), '正規化結果不含 sample 內容');
    eq(n.line, 0, '行號非數字→0');
    eq(N({ line: -5 }).line, 0, '負行號→0');
    eq(N({ line: 12.9 }).line, 12, '行號取整');
    eq(N({ disposition: 'weird' }).disposition, 'report', '處置白名單');
    // 兩種格式
    const leg = M.extractReports('legacy', { 'csp-report': { 'document-uri': 'https://h/p.html?q=1', 'violated-directive': "script-src-elem 'self'", 'effective-directive': 'script-src-elem', 'blocked-uri': 'inline', 'source-file': 'https://h/p.html', 'line-number': 7, 'script-sample': 'x=1', disposition: 'report', 'original-policy': 'default-src …' } });
    eq(leg.length, 1, 'legacy 一筆'); eq([leg[0].directive, leg[0].docPath, leg[0].line, leg[0].sampleLen], ['script-src-elem', '/p.html', 7, 3], 'legacy 欄位對應');
    const api = M.extractReports('api', [{ type: 'csp-violation', url: 'https://h/a.html', body: { documentURL: 'https://h/a.html?z=1', blockedURL: 'inline', effectiveDirective: 'script-src-attr', sourceFile: 'https://h/app.js', lineNumber: 3, sample: 'doThing()', disposition: 'report' } }, { type: 'network-error', body: {} }, null, 5]);
    eq(api.length, 1, 'reports+json 只取 csp-violation'); eq([api[0].directive, api[0].docPath, api[0].srcPath], ['script-src-attr', '/a.html', '/app.js'], 'reports+json 欄位對應');
    eq(M.extractReports('legacy', [{ 'csp-report': {} }]), null, 'legacy 收到陣列→形狀錯');
    eq(M.extractReports('legacy', { x: 1 }), null, 'legacy 缺 csp-report→形狀錯');
    eq(M.extractReports('api', { type: 'csp-violation' }), null, 'api 收到物件→形狀錯');
    eq(M.extractReports('api', new Array(100).fill({ type: 'csp-violation', body: { blockedURL: 'inline' } })).length, 25, '每批最多 25 筆');
    eq(M.contentKind('application/csp-report'), 'legacy', 'CT legacy'); eq(M.contentKind('Application/CSP-Report; charset=utf-8'), 'legacy', 'CT legacy 大小寫與參數');
    eq(M.contentKind('application/reports+json'), 'api', 'CT api'); eq(M.contentKind('application/json'), '', 'application/json 不收'); eq(M.contentKind('text/plain'), '', 'text/plain 不收'); eq(M.contentKind(undefined), '', '無 CT 不收');
  }

  // ── 寫入內容與敏感資料 ──
  {
    const { rep, logs, holds } = mk();
    const r = N({ doc: 'https://crm.example/index.html?token=SECRETQ', blocked: 'https://evil.example/steal?c=SECRET2', src: 'https://crm.example/app.js?v=SECRET3', sample: 'SECRET_SAMPLE', line: 42 });
    eq(rep._process({}, r), 'written', '首筆寫入');
    eq(logs.length, 1, '一筆稽核'); eq(logs[0].a, 'CSP_VIOLATION', '動作 CSP_VIOLATION'); eq(logs[0].op, '(瀏覽器回報)', '操作者固定字串');
    eq(logs[0].t, 'script-src-elem ← https://evil.example', 'target＝指令 ← origin');
    eq(logs[0].d, '指令=script-src-elem；來源=https://evil.example；頁面=/index.html；檔案=/app.js；行=42；樣本長=13；前一窗合併=0；處置=report', 'detail 格式固定');
    ok(!/SECRET/.test(logs[0].t + logs[0].d), 'target／detail 不含查詢字串、sample 內容');
    eq(holds(), 1, '寫入前呼叫 holdForFlush（Vercel 回應前等待）');
    const s = M.summarize([{ action: 'CSP_VIOLATION', ts: '2026-01-01T00:00:00.000Z', target: logs[0].t, detail: logs[0].d }]);
    eq([s.events, s.merged, s.total, s.topDirectives[0], s.topSources[0], s.topPages[0]], [1, 0, 1, { name: 'script-src-elem', count: 1 }, { name: 'https://evil.example', count: 1 }, { name: '/index.html', count: 1 }], '摘要解析 detail');
  }

  // ── 擴充套件忽略 ──
  {
    const { rep, logs } = mk();
    eq(rep._process({}, N({ blocked: 'chrome-extension://abc/x.js' })), 'extension', '擴充套件違規忽略');
    eq(logs.length, 0, '擴充套件不寫稽核'); eq(rep.stats().ignoredExtension, 1, '計數');
  }

  // ── 去重（10 分鐘窗）──
  {
    const { rep, clock, logs } = mk();
    const a = N(), b = N({ directive: 'script-src-attr', srcPath: '' });
    eq(rep._process({}, a), 'written', '第 1 次寫入');
    for (let i = 0; i < 4; i++) { clock.adv(1000); eq(rep._process({}, N({ line: 100 + i })), 'deduped', `窗內第 ${i + 2} 次（行號不同仍同鍵）只計數`); }
    eq(logs.length, 1, '窗內只有 1 筆稽核');
    eq(rep._process({}, b), 'written', '不同指令是不同鍵'); eq(logs.length, 2, '2 筆');
    eq(rep._process({}, N({ doc: 'https://crm.example/index.html' })), 'written', '不同頁面是不同鍵'); eq(logs.length, 3, '3 筆');
    clock.adv(10 * 60 * 1000 - 5000);   // 已過 4s + 10min-5s = 剛好不到 10 分鐘窗結束？ a 的起點在 t0，目前 t0+4s+(10min-5s)=t0+10min-1s
    eq(rep._process({}, a), 'deduped', '第 10 分鐘前 1 秒仍在窗內');
    clock.adv(2000);
    eq(rep._process({}, a), 'written', '超過 10 分鐘再寫一筆');
    ok(/前一窗合併=5/.test(logs[logs.length - 1].d), `新窗註明前一窗合併次數（4+1=5）：${logs[logs.length - 1].d}`);
  }

  // ── 每小時寫入預算 ──
  {
    const { rep, clock, logs } = mk({ CSP_REPORT_MAX_WRITES_PER_HOUR: '3' });
    for (let i = 0; i < 3; i++) eq(rep._process({}, N({ blocked: `https://h${i}.example` })), 'written', `預算內第 ${i + 1} 筆`);
    eq(rep._process({}, N({ blocked: 'https://h3.example' })), 'budget', '超過預算不寫');
    eq(rep._process({}, N({ blocked: 'https://h4.example' })), 'budget', '仍不寫');
    eq(logs.filter(l => l.a === 'CSP_VIOLATION').length, 3, '只有 3 筆 CSP_VIOLATION');
    eq(logs.filter(l => l.a === 'CSP_REPORT_THROTTLED').length, 1, '節流標記每小時最多 1 筆');
    ok(/上限 3 筆/.test(logs[logs.length - 1].d), '標記說明上限');
    clock.adv(60 * 60 * 1000 + 1000);
    eq(rep._process({}, N({ blocked: 'https://h3.example' })), 'written', '一小時後預算重置');
    ok(/前一窗合併=1/.test(logs[logs.length - 1].d), '被預算吃掉的次數帶到下一個視窗');
    eq(rep.stats().budgetDropped, 2, '丟棄計數');
    // 預設值依稽核模式
    eq(mk({}, 'table').rep.stats().maxWritesPerHour, 120, '新表模式預設 120/小時'); eq(mk({}, 'legacy').rep.stats().maxWritesPerHour, 4, '舊路徑預設 4/小時（資安審查 M4：24 → 4）');
    eq(mk({ CSP_REPORT_MAX_WRITES_PER_HOUR: '0' }).rep.stats().maxWritesPerHour, 0, '可設 0（完全不寫稽核）');
    const z = mk({ CSP_REPORT_MAX_WRITES_PER_HOUR: '0' }); z.rep._process({}, N()); eq(z.logs.filter(l => l.a === 'CSP_VIOLATION').length, 0, '預算 0 → 不寫任何違規列');
  }

  // ── 資安審查 M4：來源頁面白名單、去重鍵不含 srcPath ──
  {
    const { rep, logs } = mk({ CSP_REPORT_MAX_WRITES_PER_HOUR: '100000' });
    eq(rep._process({}, N({ doc: 'https://h/evil-page-1.html' })), 'unknown-doc', '來源頁面不在已知清單 → 只計數');
    eq(rep._process({}, N({ doc: 'https://h/' + 'x'.repeat(60) + '.html' })), 'unknown-doc', '任意變造的頁面路徑 → 只計數');
    eq(rep._process({}, N({ doc: 'about:blank' })), 'unknown-doc', '非 http(s) 文件 → 只計數');
    eq(logs.length, 0, '未知頁面不寫稽核'); eq(rep.stats().unknownDoc, 3, 'unknownDoc 計數'); eq(rep.stats().dedupKeys, 0, '未知頁面不佔去重表');
    for (const p of ['/', '/index.html', '/login.html', '/admin.html', '/help.html']) eq(rep._process({}, N({ doc: 'https://crm.example' + p, blocked: 'https://k' + p.length + '.example' })), 'written', '已知頁面入帳：' + p);
    eq(rep._process({}, N({ src: 'https://crm.example/a1.js' })), 'written', '首次');
    eq(rep._process({}, N({ src: 'https://crm.example/a2.js' })), 'deduped', '只有 srcPath 不同 → 同一去重鍵（攻擊者無法用 source-file 灌出新事件）');
    // 攻擊情境：60 筆內容各異（頁面、檔案、行號皆變造）→ 舊路徑預設預算內最多 4 筆
    const a = mk({}, 'legacy');
    for (let i = 0; i < 60; i++) a.rep._process({}, N({ doc: 'https://h/p' + i + '.html', src: 'https://h/s' + i + '.js', line: i, blocked: 'https://h' + i + '.example' }));
    for (let i = 0; i < 60; i++) a.rep._process({}, N({ src: 'https://h/s' + i + '.js', line: i, blocked: 'https://h' + i + '.example' }));
    eq(a.logs.filter(l => l.a === 'CSP_VIOLATION').length, 4, '舊路徑預設：60 筆變造頁面＋60 筆已知頁面不同來源 → 稽核至多 4 筆');
  }

  // ── 鍵上限 ──
  {
    const { rep, logs } = mk({ CSP_REPORT_MAX_WRITES_PER_HOUR: '100000' });
    for (let i = 0; i < 1000; i++) rep._process({}, N({ blocked: `https://h${i}.example` }));
    eq(rep._process({}, N({ blocked: 'https://overflow.example' })), 'map-full', '去重表滿 1000 鍵 → 丟棄新鍵');
    eq(logs.length, 1000, '前 1000 鍵各一筆'); eq(rep.stats().mapFullDropped, 1, '計數');
  }

  // ── 限流 ──
  {
    const { rep, clock } = mk({ CSP_REPORT_IP_PER_MIN: '5', CSP_REPORT_GLOBAL_PER_MIN: '8' });
    const call = (ip) => { let st = 0, nx = 0; rep._limiter({ ip }, { set() {}, status(c) { st = c; return { end() {} }; } }, () => nx++); return nx ? 'pass' : st; };
    const seq = []; for (let i = 0; i < 7; i++) seq.push(call('1.1.1.1'));
    eq(seq, ['pass', 'pass', 'pass', 'pass', 'pass', 429, 429], '單一 IP 每分鐘 5 次');
    eq([call('2.2.2.2'), call('2.2.2.2'), call('2.2.2.2')], ['pass', 'pass', 'pass'], '另一個 IP 不受影響（全域 8 次內）');
    eq(call('3.3.3.3'), 429, '全域每分鐘上限（5+3=8 之後）');
    clock.adv(61000);
    eq(call('1.1.1.1'), 'pass', '一分鐘後恢復');
    eq(rep.stats().rateLimited, 3, '限流計數');
  }

  // ── 摘要 ──
  {
    const rows = [];
    const add = (dir, src, page, merged, ts) => rows.push({ action: 'CSP_VIOLATION', ts, target: `${dir} ← ${src}`, detail: `指令=${dir}；來源=${src}；頁面=${page}；行=1；樣本長=0；前一窗合併=${merged}；處置=report` });
    add('script-src-attr', 'inline', '/admin.html', 4, '2026-01-02T00:00:00.000Z'); add('script-src-attr', 'inline', '/index.html', 0, '2026-01-03T00:00:00.000Z'); add('script-src-elem', 'inline', '/admin.html', 1, '2026-01-01T00:00:00.000Z');
    add('script-src-elem', 'https://cdn.jsdelivr.net', '/admin.html', 0, '2026-01-01T00:00:00.000Z');
    rows.push({ action: 'CSP_REPORT_THROTTLED', ts: '2026-01-03T00:00:00.000Z', target: '本實例', detail: 'x' }); rows.push({ action: 'LOGIN', ts: 'x', detail: '指令=bad' });
    const s = M.summarize(rows);
    eq([s.events, s.merged, s.total, s.throttled, s.lastAt], [4, 5, 9, 1, '2026-01-03T00:00:00.000Z'], '摘要總數（不同事件＋合併次數；忽略其他動作）');
    eq(s.topDirectives, [{ name: 'script-src-attr', count: 6 }, { name: 'script-src-elem', count: 3 }], '指令排名以權重加總');
    eq(s.topSources[0], { name: 'inline', count: 8 }, '來源排名'); eq(s.topPages[0], { name: '/admin.html', count: 8 }, '頁面排名');
    const s2 = M.summarize(rows.concat(Array.from({ length: 12 }, (_, i) => ({ action: 'CSP_VIOLATION', ts: '2026-01-01T00:00:00.000Z', detail: `指令=img-src；來源=https://h${i}.example；頁面=/p.html；前一窗合併=0` }))));
    eq(s2.topSources.length, 5, '前 5 名截斷');
    eq(M.summarize([]), { events: 0, merged: 0, total: 0, throttled: 0, lastAt: '', topDirectives: [], topSources: [], topPages: [] }, '空資料');
    eq(M.summarize([{ action: 'CSP_VIOLATION', detail: '前一窗合併=999999999999' }]).merged, 1000000, '合併次數封頂（防止偽造的超大數字）');
  }
}

function loadMutant(src) {
  const m = new Module(LIB, module); m.filename = LIB; m.paths = Module._nodeModulePaths(path.dirname(LIB));
  m._compile(src, LIB); return m.exports;
}
function mutate(src, from, to, label) {
  if (!src.includes(from)) throw new Error('變異錨點不存在：' + label);
  return src.replace(from, () => to);
}

function runUnit() {
  console.log('Part A：單元測試');
  unitSuite(cspMod, U);
  console.log(`  單元：${upass} 通過，${ufail} 失敗`);
  if (ufail) console.log(ufailures.map(f => '   - ' + f).join('\n'));
  // ── 變異檢查：把規則改壞，測試必須抓得到 ──
  const src = fs.readFileSync(LIB, 'utf8');
  const muts = [
    ['pathOf 不去查詢字串', () => mutate(src, "return cleanPath(new URL(s).pathname);", "return cleanPath(new URL(s).pathname + new URL(s).search);", 'path-query')],
    ['blocked 保留完整網址', () => mutate(src, "return { v: `${sc}://${u.hostname.toLowerCase()}${u.port ? ':' + /^\\d{1,5}$/.exec(u.port)[0] : ''}`, ext: false };", "return { v: s, ext: false };", 'blocked-full')],
    ['去重失效', () => mutate(src, "if (e && t - e.t < DEDUP_WINDOW_MS) {", "if (false) {", 'dedup')],
    ['移除來源頁面白名單', () => mutate(src, "if (!KNOWN_DOC_PATHS.has(r.docPath)) {", "if (false) {", 'known-doc')],
    ['去重鍵含 srcPath', () => mutate(src, "const key = [r.directive, r.blocked, r.docPath].join('|');", "const key = [r.directive, r.blocked, r.docPath, r.srcPath].join('|');", 'dedup-src')],
    ['舊路徑預算回到 24', () => mutate(src, 'const LEGACY_MAX_WRITES = 4;', 'const LEGACY_MAX_WRITES = 24;', 'legacy-budget')],
    ['預算失效', () => mutate(src, "if (budget.used >= maxWrites()) return false;", "if (false) return false;", 'budget')],
    ['單 IP 限流失效', () => mutate(src, "if (e.n >= (parseInt(env.CSP_REPORT_IP_PER_MIN, 10) || IP_LIMIT_PER_MIN)) return false;", "", 'ip-limit')],
    ['不忽略擴充套件', () => mutate(src, "if (r.ext) { stats.ignoredExtension++; return 'extension'; }", "", 'ext')],
    ['script-src 加 unsafe-inline', () => mutate(src, `"script-src 'self' https://cdn.jsdelivr.net",`, `"script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net",`, 'policy')],
    ['sample 全文寫入', () => mutate(src, "sampleLen: Math.min(typeof o.sample === 'string' ? o.sample.length : 0, 9999),", "sampleLen: o.sample, sampleRaw: o.sample,", 'sample')],
    ['指令不走白名單', () => mutate(src, "return DIRECTIVES.has(t) ? t : '(未知指令)';", "return t || '(未知指令)';", 'directive')],
  ];
  let caught = 0;
  for (const [label, mk] of muts) {
    let mod; try { mod = loadMutant(mk()); } catch (e) { console.log('  ✗ 變異無法載入：', label, e.message); ufail++; continue; }
    let f = 0; const T = { ok(c) { if (!c) f++; }, eq(a, b) { if (JSON.stringify(a) !== JSON.stringify(b)) f++; } };
    try { unitSuite(mod, T); } catch (e) { f++; }
    if (f > 0) caught++; else { ufail++; console.log('  ✗ 變異未被抓到：', label); }
  }
  console.log(`  變異檢查：${caught}/${muts.length} 被抓到`);
}

// ═══════════════════════ Part B：沙盒端對端 ═══════════════════════
const BASE_PORT = parseInt(process.env.CSP_BASE_PORT || '3081', 10);
// 基底釘死為 commit（原本是會移動的分支 feat/integrate-audit3：分支刪除或換機器沒有該分支就失敗）。
// 1f1eeb0＝CSP 功能合併之前的整併分支尖端（不含 lib/cspReport.js）。可用 CSP_BASE_REF 覆寫。
const BASE_REF = process.env.CSP_BASE_REF || '1f1eeb097909ee70fc41310e203c1714e957af1c';

function stripVolatile(h) {
  const o = {};
  for (const [k, v] of h.entries()) {
    const kl = k.toLowerCase();
    if (['date', 'etag', 'last-modified', 'content-length', 'set-cookie', 'content-encoding', 'vary'].includes(kl)) continue;
    if (/^(x-)?ratelimit/.test(kl)) continue;
    o[kl] = v;
  }
  return o;
}

async function runE2E() {
  const H = require('./_b2-harness');
  const { ok, eq, sleep, http, mkUser, fresh, loginAll, finish, ROOT, BASE, PORT } = H;
  H.guardClean();
  console.log(`\nPart B：沙盒端對端（埠 ${PORT}，基底比對埠 ${BASE_PORT}）`);
  const raw = async (method, p, opt) => {
    const o = opt || {};
    const res = await fetch(BASE + p, { method, headers: o.headers || {}, body: o.body, redirect: 'manual' });
    const buf = Buffer.from(await res.arrayBuffer());
    return { status: res.status, headers: res.headers, text: buf.toString('utf8'), len: buf.length };
  };
  const LEGACY = { 'Content-Type': 'application/csp-report' };
  const leg = (o) => JSON.stringify({ 'csp-report': Object.assign({ 'document-uri': BASE + '/index.html', 'violated-directive': "script-src-elem 'self' https://cdn.jsdelivr.net", 'effective-directive': 'script-src-elem', 'blocked-uri': 'inline', 'source-file': BASE + '/index.html', 'line-number': 5, 'script-sample': '', disposition: 'report', 'original-policy': "default-src 'self'" }, o || {}) });
  const users = () => [mkUser('_s_admin', 'admin'), mkUser('_s_user', 'user', { bu: ['ERP'] })];
  const csp = async (who) => { const r = await http(who || '_s_admin', 'GET', '/api/admin/logs?limit=200&action=CSP_VIOLATION&paged=1'); return (r.json && r.json.rows) || []; };

  try {
    // ── 預設 on ──
    await fresh(users(), {}, { env: { CSP_REPORT_ONLY: 'on' } });
    await loginAll(['_s_admin', '_s_user']);
    const lp = await raw('GET', '/login.html');
    const ro = lp.headers.get('content-security-policy-report-only');
    ok(lp.status === 200 && ro && ro.includes("script-src 'self' https://cdn.jsdelivr.net") && ro.includes("script-src-attr 'none'") && ro.includes('report-uri /api/csp-report'), '/login.html 有 Report-Only 標頭且內容正確');
    ok(!lp.headers.get('reporting-endpoints'), '預設沒有 Reporting-Endpoints 標頭');
    const enforced = lp.headers.get('content-security-policy');
    ok(enforced && enforced.includes("script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net") && enforced.includes("script-src-attr 'unsafe-inline'") && !enforced.includes('report-uri'), '強制 CSP 照舊（仍允許 inline、沒有 report-uri）');
    const api = await raw('GET', '/api/me');
    ok(api.status === 401 && !api.headers.get('content-security-policy-report-only'), '/api/ 回應不加 Report-Only');
    const q = await raw('GET', '/q/abc');
    ok(q.headers.get('content-security-policy') && !q.headers.get('content-security-policy-report-only'), '/q/ 跳板頁維持自己的 CSP、不加 Report-Only');

    // ── 端點：未登入可用、基本寫入 ──
    const t0 = Date.now();
    const r1 = await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg({ 'document-uri': BASE + '/index.html?token=SECRETQ#h', 'blocked-uri': 'https://evil.example/steal?c=SECRET2', 'source-file': BASE + '/app.js?v=SECRET3', 'script-sample': 'SECRET_SAMPLE', 'line-number': 42 }) });
    ok(r1.status === 204 && r1.len === 0, `未登入 POST legacy 報告 → 204 空本文（得到 ${r1.status}/${r1.len}）`);
    ok(Date.now() - t0 < 3000, '回報端點回應很快');
    let rows = await csp();
    eq(rows.length, 1, '稽核有 1 筆 CSP_VIOLATION');
    ok(rows[0] && rows[0].target === 'script-src-elem ← https://evil.example' && /頁面=\/index\.html；檔案=\/app\.js；行=42；樣本長=13；前一窗合併=0/.test(rows[0].detail), `欄位清洗正確：${rows[0] && rows[0].detail}`);
    eq(rows[0] && rows[0].operator, '(瀏覽器回報)', '操作者固定字串');
    // 敏感資料掃描：JSON、CSV 匯出整份
    const dump = JSON.stringify((await http('_s_admin', 'GET', '/api/admin/logs')).json);
    const csv = (await http('_s_admin', 'GET', '/api/admin/logs/export?format=csv')).text;
    for (const s of ['SECRETQ', 'SECRET2', 'SECRET3', 'SECRET_SAMPLE', '/steal']) ok(!dump.includes(s) && !csv.includes(s), `稽核（JSON／CSV）不含敏感字串 ${s}`);
    // 去重
    for (let i = 0; i < 5; i++) await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg({ 'document-uri': BASE + '/index.html?x=' + i, 'blocked-uri': 'https://evil.example/other', 'source-file': BASE + '/app.js', 'line-number': i }) });
    rows = await csp();
    eq(rows.length, 1, '同指令＋來源＋頁面＋來源檔 5 次只記 1 筆（去重）');
    await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg({ 'effective-directive': 'script-src-attr', 'violated-directive': 'script-src-attr', 'blocked-uri': 'inline' }) });
    eq((await csp()).length, 2, '不同指令另記一筆');
    // Reporting API 格式（陣列，含一筆非 csp-violation）
    const apiBody = JSON.stringify([{ type: 'csp-violation', age: 10, url: BASE + '/admin.html', user_agent: 'x', body: { documentURL: BASE + '/admin.html?z=1', blockedURL: 'inline', effectiveDirective: 'style-src-attr', disposition: 'report', sourceFile: null, lineNumber: 9, sample: 'color:red' } }, { type: 'deprecation', body: {} }]);
    const r2 = await raw('POST', '/api/csp-report', { headers: { 'Content-Type': 'application/reports+json' }, body: apiBody });
    eq(r2.status, 204, 'reports+json → 204');
    ok((await csp()).some(x => x.target === 'style-src-attr ← inline' && /頁面=\/admin\.html/.test(x.detail)), 'reports+json 的違規入稽核');
    // 擴充套件忽略
    const before = (await csp()).length;
    await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg({ 'blocked-uri': 'chrome-extension://abc/x.js', 'effective-directive': 'script-src-elem' , 'document-uri': BASE + '/zzz.html' }) });
    eq((await csp()).length, before, '瀏覽器擴充套件造成的違規不入稽核');

    // ── 拒絕條件 ──
    const rj = async (name, status, opt, m) => { const r = await raw(m || 'POST', '/api/csp-report', opt); ok(r.status === status && r.len === 0, `${name} → ${status} 空本文（得到 ${r.status}/${r.len}）`); };
    await rj('Content-Type: application/json', 415, { headers: { 'Content-Type': 'application/json' }, body: leg() });
    await rj('Content-Type: text/plain', 415, { headers: { 'Content-Type': 'text/plain' }, body: leg() });
    await rj('沒有 Content-Type', 415, { body: Buffer.from(leg()) });
    await rj('Content-Encoding: gzip', 415, { headers: { 'Content-Type': 'application/csp-report', 'Content-Encoding': 'gzip' }, body: leg() });
    await rj('本文 40KB', 413, { headers: LEGACY, body: leg({ 'script-sample': 'x'.repeat(40000) }) });
    await rj('壞 JSON', 400, { headers: LEGACY, body: '{not json' });
    await rj('空本文', 400, { headers: LEGACY, body: '' });
    await rj('legacy 形狀錯（陣列）', 400, { headers: LEGACY, body: '[1,2]' });
    await rj('reports+json 形狀錯（物件）', 400, { headers: { 'Content-Type': 'application/reports+json' }, body: '{"type":"csp-violation"}' });
    await rj('GET', 405, {}, 'GET');
    await rj('PUT', 405, { headers: LEGACY, body: leg() }, 'PUT');
    await rj('DELETE', 405, {}, 'DELETE');
    // 邊界：剛好 30KB 的合法報告可收
    const big = await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg({ 'original-policy': 'x'.repeat(30000), 'blocked-uri': 'https://big.example/x' }) });
    eq(big.status, 204, '30KB 合法報告仍收');

    // ── 其他 API 不受影響；摘要 ──
    eq((await http('_s_admin', 'GET', '/api/me')).status, 200, '登入者 /api/me 正常');
    const sm = await http('_s_admin', 'GET', '/api/admin/csp-summary');
    ok(sm.status === 200 && sm.json && sm.json.enabled === true && sm.json.events >= 3 && Array.isArray(sm.json.topDirectives) && sm.json.memory && sm.json.memory.written >= 3, `管理員摘要正常：${JSON.stringify(sm.json).slice(0, 200)}`);
    ok(sm.json.topDirectives.every(x => x.name && x.count >= 1) && sm.json.topDirectives.length <= 5, '前 5 名結構正確');
    eq((await http('_s_user', 'GET', '/api/admin/csp-summary')).status, 403, '非管理員不能看摘要');
    eq((await http(null, 'GET', '/api/admin/csp-summary')).status, 401, '未登入不能看摘要');
    const ah = await http('_s_admin', 'GET', '/admin.html');
    ok(ah.status === 200 && /id="vCspInfo"/.test(ah.text) && /CSP_VIOLATION: 'CSP/.test(ah.text), 'admin.html 有 CSP 卡片與動作中文名');

    // ── 洪水不吃 API 額度（全域 apiLimiter 300/分）──
    await fresh(users(), {}, { env: { CSP_REPORT_ONLY: 'on', CSP_REPORT_IP_PER_MIN: '100000', CSP_REPORT_GLOBAL_PER_MIN: '100000' } });
    await loginAll(['_s_admin']);
    let n204 = 0;
    for (let i = 0; i < 340; i++) { const r = await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg({ 'line-number': i }) }); if (r.status === 204) n204++; }
    eq(n204, 340, '340 次洪水回報全部 204');
    const after = await http('_s_admin', 'GET', '/api/me');
    eq(after.status, 200, '洪水後 /api/me 仍 200（回報端點不吃 /api 全域 300/分 額度）');
    eq((await csp()).length, 1, '340 次相同回報只有 1 筆稽核');
    const mem = (await http('_s_admin', 'GET', '/api/admin/csp-summary')).json.memory;
    eq([mem.written, mem.deduped], [1, 339], '記憶體計數：寫入 1、去重 339');

    // ── 預算 ──
    await fresh(users(), {}, { env: { CSP_REPORT_ONLY: 'on', CSP_REPORT_MAX_WRITES_PER_HOUR: '3' } });
    await loginAll(['_s_admin']);
    for (let i = 0; i < 7; i++) await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg({ 'blocked-uri': `https://h${i}.example/x.js` }) });
    const ev = (await http('_s_admin', 'GET', '/api/admin/logs?limit=200&paged=1')).json.rows;
    eq(ev.filter(x => x.action === 'CSP_VIOLATION').length, 3, '預算 3 → 只寫 3 筆違規');
    eq(ev.filter(x => x.action === 'CSP_REPORT_THROTTLED').length, 1, '另有 1 筆節流標記');
    const sm2 = (await http('_s_admin', 'GET', '/api/admin/csp-summary')).json;
    ok(sm2.throttled === 1 && sm2.memory.budgetDropped === 4, `摘要反映節流：throttled=${sm2.throttled} budgetDropped=${sm2.memory.budgetDropped}`);

    // ── report-to 選項 ──
    await fresh(users(), {}, { env: { CSP_REPORT_ONLY: 'on', CSP_REPORT_TO: 'on' } });
    { const r = await raw('GET', '/login.html'); eq(r.headers.get('reporting-endpoints'), 'csp-endpoint="/api/csp-report"', 'CSP_REPORT_TO=on：有 Reporting-Endpoints'); ok(/report-to csp-endpoint/.test(r.headers.get('content-security-policy-report-only') || ''), 'CSP_REPORT_TO=on：政策有 report-to'); }

    // ── 單 IP 限流 ──
    await fresh(users(), {}, { env: { CSP_REPORT_ONLY: 'on', CSP_REPORT_IP_PER_MIN: '10' } });
    await loginAll(['_s_admin']);
    const codes = []; for (let i = 0; i < 16; i++) codes.push((await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg() })).status);
    eq(codes.filter(c => c === 429).length, 6, '單 IP 每分鐘 10 次，第 11–16 次 429');
    const lim = await raw('POST', '/api/csp-report', { headers: LEGACY, body: leg() });
    ok(lim.len === 0 && lim.headers.get('retry-after') === '60', '429 空本文並帶 Retry-After');

    // ── 與基底逐頁標頭比對（位元級）──
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'csp-base-'));
    let baseChild = null;
    try {
      const tarFile = path.join(tmp, '_base.tar');
      execSync(`git archive --format=tar -o "${tarFile}" ${BASE_REF}`, { cwd: ROOT, stdio: 'pipe' });
      execSync('tar -xf _base.tar', { cwd: tmp, stdio: 'pipe' });   // cwd＋相對路徑：GNU tar 會把 "C:\…" 當成遠端主機
      fs.unlinkSync(tarFile);
      fs.symlinkSync(path.join(ROOT, 'node_modules'), path.join(tmp, 'node_modules'), 'junction');
      const paths = ['/login.html', '/login-cinematic.js', '/login-boot.js', '/deep-link.js', '/itts-logo.svg', '/manifest.webmanifest', '/', '/index.html', '/admin.html', '/q/abc', '/api/me', '/api/nonexistent', '/api/csp-report'];
      const authed = ['/index.html', '/admin.html', '/app.js', '/style.css', '/api/me', '/api/admin/logs?limit=1'];
      const snapshot = async (port, env, label) => {
        const E = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: 'x'.repeat(48), SESSION_SECRET: 'y'.repeat(48), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${port}` }, env);
        delete E.DATABASE_URL; delete E.MAIL_MODE; delete E.CSP_REPORT_ONLY; delete E.CSP_REPORT_ONLY_STYLE; Object.assign(E, env);
        const dir = label === 'base' ? tmp : ROOT;
        // 沙盒資料：與目前工作樹相同（fresh() 已寫好 ROOT 的 auth.json／data.json）
        if (dir !== ROOT) for (const f of ['auth.json', 'data.json']) fs.copyFileSync(path.join(ROOT, f), path.join(dir, f));
        await assertPortFree(port);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
        const child = spawn(process.execPath, [path.join(dir, 'server.js')], { cwd: dir, env: E, stdio: ['ignore', 'pipe', 'pipe'] });
        if (label === 'base') baseChild = child;
        let up = false; for (let i = 0; i < 80 && !up; i++) { await sleep(250); try { up = (await fetch(`http://localhost:${port}/login.html`)).ok; } catch { /* 尚未起來 */ } }
        if (!up) throw new Error('伺服器啟動逾時 ' + label);
        const out = {};
        const get = async (p, cookie, method) => { const r = await fetch(`http://localhost:${port}${p}`, { method: method || 'GET', headers: cookie ? { Cookie: cookie } : {}, redirect: 'manual' }); await r.arrayBuffer(); return { status: r.status, h: stripVolatile(r.headers) }; };
        for (const p of paths) out['U ' + (p === '/api/csp-report' ? 'POST ' : '') + p] = await get(p, null, p === '/api/csp-report' ? 'POST' : 'GET');
        const lr = await fetch(`http://localhost:${port}/api/login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: '_s_admin', password: H.PW }) });
        const cookie = (lr.headers.getSetCookie ? lr.headers.getSetCookie() : []).map(c => c.split(';')[0]).join('; ');
        for (const p of authed) out['A ' + p] = await get(p, cookie);
        return { out, child };
      };
      // 先用目前工作樹的沙盒資料起 base；new(off)、new(on) 各起一次（皆依序，避免埠衝突）
      await H.stopServer();
      const b = await snapshot(BASE_PORT, {}, 'base');
      try { b.child.kill(); } catch { /* 已結束 */ }
      await sleep(600);
      const nOff = await snapshot(BASE_PORT, { CSP_REPORT_ONLY: 'off' }, 'new');
      try { nOff.child.kill(); } catch { /* 已結束 */ }
      await sleep(600);
      const nOn = await snapshot(BASE_PORT, { CSP_REPORT_ONLY: 'on' }, 'new');
      try { nOn.child.kill(); } catch { /* 已結束 */ }
      await sleep(600);
      let diffOff = 0; const lines = [];
      for (const k of Object.keys(b.out)) {
        const a = b.out[k], c = nOff.out[k];
        const same = a.status === c.status && JSON.stringify(a.h) === JSON.stringify(c.h);
        if (!same) { diffOff++; lines.push(`   ${k}: base ${a.status} vs off ${c.status}\n     base=${JSON.stringify(a.h)}\n     off =${JSON.stringify(c.h)}`); }
      }
      ok(diffOff === 0, `CSP_REPORT_ONLY=off：${Object.keys(b.out).length} 個回應的狀態碼與標頭和基底完全相同` + (diffOff ? '\n' + lines.join('\n') : ''));
      // on：與基底相比，只多兩個標頭，且只在非 /api/、非 /q/
      let bad = 0; const badLines = [];
      for (const k of Object.keys(b.out)) {
        if (k === 'U POST /api/csp-report') continue;   // on 時由端點處理（另行斷言）
        const a = b.out[k], c = nOn.out[k];
        const p = k.replace(/^[UA] (POST )?/, '');
        const expectAdded = !p.startsWith('/api/') && !p.startsWith('/q/');
        const added = Object.keys(c.h).filter(x => !(x in a.h)).sort();
        const removed = Object.keys(a.h).filter(x => !(x in c.h));
        const changed = Object.keys(a.h).filter(x => x in c.h && a.h[x] !== c.h[x]);
        const want = expectAdded ? ['content-security-policy-report-only'] : [];
        if (a.status !== c.status || JSON.stringify(added) !== JSON.stringify(want) || removed.length || changed.length) { bad++; badLines.push(`   ${k}: added=${added} removed=${removed} changed=${changed} status ${a.status}->${c.status}`); }
      }
      ok(bad === 0, 'CSP_REPORT_ONLY=on（預設設定）：與基底相比，HTML／靜態檔只多 Content-Security-Policy-Report-Only 一個標頭，/api/ 與 /q/ 不變' + (bad ? '\n' + badLines.join('\n') : ''));
      ok(nOn.out['U POST /api/csp-report'].status === 204 || nOn.out['U POST /api/csp-report'].status === 415, `on 時 POST /api/csp-report 由端點處理（${nOn.out['U POST /api/csp-report'].status}）`);
      eq(nOff.out['U POST /api/csp-report'].status, b.out['U POST /api/csp-report'].status, 'off 時 /api/csp-report 與基底同為未註冊路徑（登入閘門回應）');
    } finally {
      if (baseChild) { try { baseChild.kill(); } catch { /* 已結束 */ } }
      await sleep(500);
      // 只移除 junction 本身（rmdir 不遞迴）；確定連結已不在才遞迴刪暫存目錄，避免刪進主 repo 的 node_modules
      try { fs.rmdirSync(path.join(tmp, 'node_modules')); } catch { /* 沒建立 */ }
      if (!fs.existsSync(path.join(tmp, 'node_modules'))) { try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) { console.log('  (暫存基底目錄未清乾淨：' + tmp + ')'); } }
      else console.log('  (junction 未能移除，保留暫存目錄：' + tmp + ')');
    }

    // ── 盤點工具不變式：登入頁必須維持零 inline ──
    const inv = require('./csp-inventory').main();
    const login = inv.htmls.find(h => h.file === 'login.html');
    ok(login && login.inlineScripts.filter(s => s.executes).length === 0 && login.staticOn.total === 0 && login.staticStyle.total === 0 && login.inlineStyles.length === 0, 'login.html 維持零 inline script／on*／style');
  } finally {
    await finish();
  }
  const code = H.report('Part B（端對端）');
  return code;
}

(async () => {
  runUnit();
  let code = ufail ? 1 : 0;
  if (!process.env.CSP_SKIP_E2E) code = (await runE2E()) || code;
  console.log(code ? '\n有失敗項目' : '\n全部通過');
  process.exit(code);
})().catch(async (e) => { console.error('檢查腳本例外：', e); try { await require('./_b2-harness').finish(); } catch { /* 忽略 */ } process.exit(1); });
