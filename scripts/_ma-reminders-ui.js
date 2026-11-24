'use strict';
/**
 * MA／License 到期活規則提醒前端（P3b）瀏覽器檢查：headless Edge＋CDP（真實滑鼠鍵盤）。
 * 單獨執行：node scripts/_ma-reminders-ui.js            （全部區段；MAREM_UI_ONLY=off|admin|dual|bell|items|quote|look 只跑一段）
 *          MAREM_UI_MUT=1 node scripts/_ma-reminders-ui.js （UI 變異：用 CDP 把前端檔換成改壞的版本，對應區段必須變紅）
 * 埠：伺服器 B2_PORT（預設 3330，限 3330–3339）、CDP 9630（MAREM_UI_CDP，限 9630–9639）。一次只跑一個。
 *     全部假資料、臨時 _s_* 帳號、sandbox 的 data.json／auth.json（工作樹若已有這些檔案一律拒絕執行，不碰真實檔，不碰 3020）。
 * 區段：
 *   off    開關關閉：後台、使用者端首頁、鈴鐺、報價單表單與基準（git archive 3d85397）逐位元相同，沒有任何 .mar-* 元素／後台側欄項／樣式表
 *   admin  後台「提醒規則」頁：規則編輯與即時驗證（與伺服器 validateConfig 對拍）、儲存帶回規則 id、還原預設、試算器（滑桿／報價單狀態／真實納管項，對拍 simulate API）、時間軸 SVG、XSS
 *   dual   雙人覆核啟用：儲存回 202 → 顯示「已送出待另一位管理員覆核」、規則沒變、CfgReview 橫幅
 *   bell   鈴鐺 maReminders 群組：出現、展開、純文字、已讀、深連結、條件達成後下一次輪詢自動消失（送簽／標不續約）、秘書視角（無金額與合約編號）、儲存型通知類型
 *   items  到期管理頁：狀態徽章、下一個提醒、提醒時間軸、標記不續約／重新開啟、篩選「已不續約」、秘書視角、XSS
 *   quote  報價單「這是續約嗎？」：顯示、選擇、儲存、重載保留、已核准唯讀、不顯示給客戶（預覽／簽核面板）
 *   look   XSS、四種外觀對比（以實際算出的顏色驗證）、手機寬度（≤480px 不爆版）、prefers-reduced-motion、層級、截圖（scratchpad/ma4_shots/）
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');
process.env.B2_PORT = process.env.B2_PORT || '3330';
const { withBrowser } = require('./_fd-cdp');

const ROOT = path.resolve(__dirname, '..');
const J = JSON.stringify;
const CDP = parseInt(process.env.MAREM_UI_CDP || '9630', 10);
const BASE_REF = process.env.MAREM_UI_BASE || 'origin/main';   // P3a 之前（P2b 驗收後修正）：沒有任何提醒規則前端
if (!((parseInt(process.env.B2_PORT, 10) >= 3330 && parseInt(process.env.B2_PORT, 10) <= 3339) || (parseInt(process.env.B2_PORT, 10) >= 3350 && parseInt(process.env.B2_PORT, 10) <= 3359))) { console.log('拒絕：本腳本只准用埠 3330–3339 或 3350–3359'); process.exit(2); }
if (!((CDP >= 9630 && CDP <= 9639) || (CDP >= 9650 && CDP <= 9659))) { console.log('拒絕：本腳本只准用 CDP 埠 9630–9639 或 9650–9659'); process.exit(2); }

let res = [];
function t(name, ok, extra) { res.push([name, !!ok, extra == null ? '' : String(extra)]); if (!ok) console.log('  ✗ ' + name + (extra != null && extra !== '' ? '｜' + String(extra).slice(0, 260) : '')); }

// ───────── 對比度（沿祖先逐層疊合半透明底；SVG 文字用 fill）─────────
const CONTRAST_FN = `(function(el, dark){
  function parse(c){ var m=c.match(/rgba?\\(([^)]+)\\)/); if(!m) return [0,0,0,0]; var p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number); return [p[0],p[1],p[2],p.length>3?p[3]:1]; }
  function over(top,bot){ var a=top[3]; return [top[0]*a+bot[0]*(1-a), top[1]*a+bot[1]*(1-a), top[2]*a+bot[2]*(1-a), 1]; }
  function lum(c){ var f=function(v){v/=255; return v<=0.03928? v/12.92 : Math.pow((v+0.055)/1.055,2.4)}; return 0.2126*f(c[0])+0.7152*f(c[1])+0.0722*f(c[2]); }
  if(!el) return null;
  var chain=[]; for(var n=el;n&&n.nodeType===1;n=n.parentElement) chain.push(n);
  var bg=dark?[11,16,32,1]:[255,255,255,1];
  for(var i=chain.length-1;i>=0;i--){ var c=parse(getComputedStyle(chain[i]).backgroundColor); if(c[3]>0) bg=over(c,bg); }
  var cs=getComputedStyle(el); var isSvg = el instanceof SVGElement;
  var fg=parse(isSvg ? cs.fill : cs.color); fg=over(fg,bg);
  var L1=lum(fg), L2=lum(bg); var hi=Math.max(L1,L2), lo=Math.min(L1,L2);
  return Math.round(((hi+0.05)/(lo+0.05))*100)/100;
})`;
const contrastSel = (b, sel, dark) => b.ev(`${CONTRAST_FN}(document.querySelector(${J(sel)}), ${!!dark})`);
const THEMES = ['glass-dark', 'glass-light', 'plain-light', 'plain-dark'];
async function theme(b, mode) {
  await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass',${J(mode.startsWith('glass'))}); c.toggle('dark',${J(mode.endsWith('dark'))}); return true})()`);
  await b.sleep(350);
}

// ───────── 假資料 ─────────
const tp = (n) => new Date(Date.now() + (n || 0) * 86400000).toLocaleDateString('sv-SE', { timeZone: 'Asia/Taipei' });
const agoIso = (n) => new Date(Date.now() - n * 86400000).toISOString();
const NOW = new Date().toISOString();
const XSS1 = '<img src=x onerror=window.__xss=1>', XSS2 = '<script>window.__xss=2</script>';
const AMT = '987654', CNO = 'SECRET-CNO-7788';   // 金額與合約編號：秘書的鈴鐺／清單不得出現
const U = { adm: '_s_admin', adm2: '_s_admin2', own: '_s_own', oth: '_s_other', m1: '_s_m1', sec: '_s_sec', mk: '_s_mk', exe: '_s_exe' };
function mkUsers(H) {
  return [
    H.mkUser(U.adm, 'admin', { displayName: '測試管理員' }), H.mkUser(U.adm2, 'admin', { displayName: '測試管理員二' }), H.mkUser(U.exe, 'executive'),
    H.mkUser(U.m1, 'manager1', { bu: ['ERP'], displayName: '測試主管' }),
    H.mkUser(U.own, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務甲' }),
    H.mkUser(U.oth, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務乙' }),
    H.mkUser(U.mk, 'marketing', { bu: ['ERP'], displayName: '測試行銷' }), H.mkUser(U.sec, 'secretary', { bu: ['ERP'], displayName: '測試秘書' }),
  ];
}
const PROD = 'SAP PCOE License MA';   // 報價單範本用的商品（有 productClasses 設定）
const CAT = () => ({ bus: { ERP: [{ group: '假商品群組', items: [{ name: 'Svc MA', termKind: 'ma', startsTask: false }, { name: 'Svc Lic', termKind: 'license', startsTask: false }] }] } });
const MAP = 'Svc MA 一年', LICP = 'Svc Lic 永久';
const mkItem = (id, owner, company, product, dl, extra) => Object.assign({
  id, oppId: 'o_' + id, quoteId: '', owner, company, product, termKind: 'ma', from: tp(dl - 364), to: tp(dl), contractNo: '', contractId: '', status: 'tracking', source: 'request', bu: 'ERP', createdAt: agoIso(10), updatedAt: agoIso(10), history: [],
}, extra || {});
const mkOpp = (id, owner, product, stage, company, extra) => Object.assign({ id, owner, bu: 'ERP', company, contactName: '假聯絡人', product, amount: AMT, description: '', expectedDate: tp(30), stage, createdAt: NOW }, extra || {});
function mkQ(id, no, over) {
  return Object.assign({
    id, quoteNo: no, owner: U.own, oppId: 'oR1', company: 'R甲公司', contactName: '', phone: '', mobile: '', address: '', projectName: '專案', quoteDate: tp(0), status: 'draft',
    createdAt: NOW, updatedAt: NOW, validUntil: tp(20), products: [PROD], costBy: null, costFlow: { state: 'na' }, approval: null, discountType: 'none', discountValue: 0,
    items: [{ lid: 'a1', desc: '客製開發服務', unit: '式', qty: 1, unitPrice: 10000, cost: 0 }],
  }, over || {});
}
function approvedOf(QA, q) {
  return {
    state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: NOW, submittedBy: U.own, derived: null,
    steps: [{ tier: 'mgr1', label: '一級主管', status: 'approved', assignee: U.m1, by: U.m1, at: NOW, comment: '' }], cur: 1, board: null,
    history: [{ at: NOW, by: U.own, action: 'SUBMIT', comment: '', tier: '' }, { at: NOW, by: U.m1, action: 'APPROVE', comment: '', tier: 'mgr1' }],
  };
}
/** 到期管理／鈴鐺用資料（預設規則：D-90 業務＋秘書 always；D-45 一級主管 無報價；D-30 業務＋主管＋行銷 草稿；到期後每 7 天 業務＋主管） */
function seedData(opts) {
  opts = opts || {};
  const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
  const qAppr0 = mkQ('qAppr', 'QU-R-002', { oppId: 'oR3', company: 'R庚公司', renewOf: 'iG', projectName: '已核准續約單' });
  const d = {
    contacts: [], companies: [], visits: [], receivables: [], callins: [], keyAccounts: [], groups: [], contracts: [], dealTasks: [], notifications: [],
    opportunities: [
      mkOpp('oR1', U.own, PROD, 'C', 'R甲公司'), mkOpp('oR2', U.own, PROD, 'C', 'R無候選公司'), mkOpp('oR3', U.own, PROD, 'C', 'R庚公司'),
    ],
    quotations: [
      mkQ('qDraft', 'QU-R-001', { projectName: '草稿續約單', renewOf: 'iA' }),
      Object.assign({}, qAppr0, { approval: approvedOf(QA, qAppr0), status: 'approved' }),
      mkQ('qPlain', 'QU-R-003', { oppId: 'oR2', company: 'R無候選公司', projectName: '無候選' }),
    ],
    quoteApproval: { roster: { gm: [U.exe], chairman: [], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: { [PROD]: { cls: 'software', costBySales: true } } },
    maItems: [
      mkItem('iA', U.own, 'R甲公司', MAP, 40, { contractNo: CNO }),
      mkItem('iB', U.own, 'R乙公司', LICP, 20, { termKind: 'license' }),
      mkItem('iC', U.own, 'R丙公司', MAP, -20),
      mkItem('iD', U.own, 'R丁公司', MAP, 25, { status: 'not_renew', notRenewReason: '預算不足：明年再議', history: [{ at: agoIso(2), by: U.own, action: 'NOT_RENEW', note: '預算不足' }] }),
      mkItem('iE', U.own, 'R戊公司', MAP, 30, { status: 'renewed' }),
      mkItem('iF', U.own, 'R己公司', MAP, 100),
      mkItem('iG', U.own, 'R庚公司', MAP, 25),
      mkItem('iX', U.own, XSS1 + 'X公司', 'Svc MA ' + XSS2, 60, { contractNo: XSS1 }),
      mkItem('iO', U.oth, 'R他人公司', MAP, 55),
    ],
  };
  if (opts.rules) d.reminderRules = { rules: opts.rules, updatedAt: NOW, updatedBy: U.adm };
  if (opts.notifs) d.notifications = opts.notifs;
  return d;
}
const firstDiff = (a, b) => { const n = Math.min(a.length, b.length); let i = 0; while (i < n && a[i] === b[i]) i++; return i + ': ' + J(a.slice(Math.max(0, i - 40), i + 80)) + ' <> ' + J(b.slice(Math.max(0, i - 40), i + 80)); };
const REGEX_ESC = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// ───────── 瀏覽器小工具 ─────────
function browse(C, maxMs, fn) {
  return withBrowser({ cdpPort: C.cdp, maxMs }, async (b) => {
    if (C.mutFiles) { Object.keys(C.mutFiles).forEach((k) => { b.replace['/' + k] = C.mutFiles[k]; }); await b.interceptOn(); }
    const oc = b.click; b.click = (sel, off) => oc(typeof sel === 'function' ? sel.toString() : sel, off);   // 允許傳函式當選取器
    return fn(b);
  });
}
const PAGE_HELPERS = `(()=>{
  window.__reqs=[]; if(!window.__recOn){ window.__recOn=true; const of=window.fetch.bind(window); window.fetch=(u,o)=>{ if(o&&o.method&&o.method!=='GET'){ window.__reqs.push({m:o.method,u:String(u),b:o.body||''}); } return of(u,o); }; }
  window.__xss=0; return 1})()`;
const reqs = (b) => b.ev(`JSON.stringify(window.__reqs)`).then((s) => JSON.parse(s));
async function login(b, C, who, w, h) {
  const H = C.H;
  await b.viewport(w || 1440, h || 900, (w || 1440) < 700);
  await b.nav(H.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(H.PW)}})}).then(r=>r.status)`);
  t('瀏覽器登入 ' + who, st === 200, st);
}
async function openUser(b, C, who, w, h, opts) {
  opts = opts || {};
  await login(b, C, who, w, h);
  await b.ev(`sessionStorage.setItem('justLoggedIn','1'); 0`);
  if (opts.clearLs) await b.ev(`localStorage.clear(); 0`);
  await b.nav(C.H.BASE + '/index.html' + (opts.hash || ''));
  await b.waitFor(`typeof showSection==='function' && !!document.getElementById('navHome') && !!window.ITTSMaTasks && !!window.ITTSMaRem`, 25000);
  if (opts.expectOn) await b.waitFor(`ITTSMaTasks._state.itemsOn`, 15000); else await b.sleep(opts.fast ? 800 : 2600);
  await b.ev(PAGE_HELPERS);
}
async function openAdmin(b, C, who, w, h, expectOn) {
  await login(b, C, who, w, h);
  await b.nav(C.H.BASE + '/admin.html');
  await b.waitFor(`typeof initProductCatalog==='function' && !!document.querySelector('[data-sec="product-catalog"]') && !!window.MaRemAdmin`, 25000);
  if (expectOn) await b.waitFor(`!!document.getElementById('marNavItem')`, 15000); else await b.sleep(2200);
  await b.ev(PAGE_HELPERS);
}
async function gotoRules(b) {
  if ((await b.ev(`document.querySelector('.nav-domain[data-domain="quote"] .nav-domain-header').getAttribute('aria-expanded')`)) !== 'true') { await b.click('.nav-domain[data-domain="quote"] .nav-domain-header'); await b.sleep(500); }
  await b.click('#marNavItem');
  await b.waitFor(`document.querySelectorAll('#marAdminHost .mar-a-row').length>0`, 15000);
  await b.sleep(900);
}
const shotTo = async (C, b, name) => { if (C.shots) { fs.mkdirSync(C.shots, { recursive: true }); await b.shot(path.join(C.shots, name)); } };
const txt = (b, sel) => b.ev(`(()=>{const e=document.querySelector(${J(sel)}); return e?e.textContent:null})()`);
/** 真實鍵盤改下拉：focus 後用上下鍵移到目標 value（以 option 順序） */
async function pickOpt(b, sel, want) {
  await b.focus(sel);
  const opts = await b.ev(`[...document.querySelector(${J(sel)}).options].map(o=>o.value)`);
  for (let i = 0; i < opts.length + 1; i++) {
    const cur = await b.ev(`document.querySelector(${J(sel)}).value`);
    if (cur === want) return true;
    const ci = opts.indexOf(cur), wi = opts.indexOf(want);
    await b.key(wi > ci ? 'ArrowDown' : 'ArrowUp');
  }
  return (await b.ev(`document.querySelector(${J(sel)}).value`)) === want;
}
async function typeInto(b, sel, text) {
  await b.focus(sel); await b.ev(`document.querySelector(${J(sel)}).select(); 0`);
  if (text === '') { await b.key('Delete'); } else await b.type(text);
  await b.sleep(120);
}
const rowS = (i) => `#marAdminHost .mar-a-row[data-rule-index="${i}"]`;
const rulesInPage = (b) => b.ev(`JSON.stringify(MaRemAdmin._state.models.map(MaRemAdmin.toApi))`).then(JSON.parse);

// ═════════════════════════ off：開關關閉＝與基準逐位元相同 ═════════════════════════
const BASE_USER_FILES = ['app.js', 'style.css', 'index.html', 'quote.js', 'ma-tasks.js'];
const STRIP = `(h)=>h.replace(/<script[\\s\\S]*?<\\/script>/g,'').replace(/<!--[\\s\\S]*?-->/g,'').replace(/nid-[a-z0-9-]+/g,'nid-X').replace(/\\s+/g,' ')`;
async function secOff(C) {
  const { H } = C;
  await H.fresh(mkUsers(H), seedData(), { env: {} });   // 資料裡有納管項、續約報價單，但開關未設：畫面必須當它們不存在
  await H.loginAll([U.own, U.adm]);
  const snap = (useBase) => withBrowser({ cdpPort: C.cdp, maxMs: 400000 }, async (b) => {
    if (C.mutFiles && !useBase) Object.keys(C.mutFiles).forEach((k) => { b.replace['/' + k] = C.mutFiles[k]; });
    if (useBase) { BASE_USER_FILES.concat(['admin.html']).forEach((f) => { b.replace['/' + f] = fs.readFileSync(path.join(C.base, '_client', f), 'utf8'); }); b.replace['/ma-reminders.js'] = '/* 基準版本沒有此檔 */'; b.replace['/ma-reminders-admin.js'] = '/* 基準版本沒有此檔 */'; }
    if (useBase || C.mutFiles) await b.interceptOn();
    const out = { reqs: [] };
    const nrm = `(${STRIP})(document.body.innerHTML)`;
    // ── 使用者端 ──
    await b.viewport(1440, 900, false);
    await b.nav(H.BASE + '/login.html');
    await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(U.own)},password:${J(H.PW)}})}).then(r=>r.status)`);
    await b.ev(`sessionStorage.setItem('justLoggedIn','1'); 0`);
    await b.nav(H.BASE + '/index.html');
    await b.waitFor(`typeof showSection==='function' && !!document.getElementById('navHome')`, 25000);
    await b.sleep(3200);
    await b.ev(PAGE_HELPERS);
    out.rem = await b.ev(`typeof window.ITTSMaRem`);
    out.home = await b.ev(nrm);
    out.sidebar = await b.ev(`(${STRIP})(document.getElementById('sidebarNav').innerHTML)`);
    out.marNodes = await b.ev(`document.querySelectorAll('[class*="mar-"], #qRenewHost, #maRemBellHost, #navMaItems, #maItemsView').length`);
    // 鈴鐺：打開下拉
    await b.ev(`document.getElementById('notifBellBtn').click(); 0`); await b.sleep(900);
    out.bell = await b.ev(`(${STRIP})(document.getElementById('notifDropdown').innerHTML)`);
    out.bellMa = await b.ev(`document.querySelectorAll('#notifDropdown [data-marem], #notifDropdown .mar-bell').length`);
    await b.ev(`document.getElementById('notifBellBtn').click(); 0`);
    // 報價單頁與表單
    await b.ev(`showSection('quotations'); 0`);
    await b.waitFor(`typeof quoteCfg==='function' && !!quoteCfg().me && document.querySelectorAll('#quoteTbody tr').length>0`, 15000);
    await b.sleep(600);
    out.quotes = await b.ev(nrm);
    out.maTracking = await b.ev(`quoteCfg().maTracking === true`);
    await b.ev(`openQuoteModal('qDraft'); 0`);
    await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='flex' && document.querySelectorAll('#quoteItemsBody tr').length>0`, 15000);
    await b.sleep(900);
    out.formEdit = await b.ev(`(${STRIP})(document.getElementById('quoteModalOverlay').innerHTML)`);
    out.renewHost = await b.ev(`!!document.getElementById('qRenewHost')`);
    await b.ev(`document.getElementById('quoteModalOverlay').style.display='none'; 0`);
    await b.ev(`openQuoteModal(null,{oppId:'oR1'}); 0`);
    await b.waitFor(`document.querySelectorAll('#quoteItemsBody tr').length>0`, 15000);
    await b.sleep(900);
    out.formNew = await b.ev(`(${STRIP})(document.getElementById('quoteModalOverlay').innerHTML)`);
    // 儲存新單的請求不送 renewOf（直接組 payload 看 saveQuote 的內容太重，改看表單 DOM 沒有任何 renew 元素）
    out.formRenewNodes = await b.ev(`document.querySelectorAll('#quoteModalOverlay [class*="renew"], #qRenewHost').length`);
    await b.ev(`document.getElementById('quoteModalOverlay').style.display='none'; 0`);
    // ── 後台 ──
    await login(b, C, U.adm, 1440, 900);
    await b.nav(H.BASE + '/admin.html');
    await b.waitFor(`typeof initProductCatalog==='function' && !!document.querySelector('[data-sec="product-catalog"]')`, 25000);
    await b.sleep(3200);
    out.adminBody = await b.ev(nrm);
    out.adminNav = await b.ev(`(${STRIP})(document.querySelector('.admin-nav, nav, .admin-sidebar') ? document.querySelector('.admin-nav, nav, .admin-sidebar').innerHTML : '')`);
    out.adminMar = await b.ev(`document.querySelectorAll('#marNavItem, #sec-reminder-rules, #marAdminHost, #maRemAdminCss, [class*="mar-a"]').length`);
    out.adminCssLinks = await b.ev(`[...document.querySelectorAll('link[rel=stylesheet]')].map(l=>l.getAttribute('href')).join(',')`);
    out.probe = await b.ev(`fetch('/api/admin/reminder-rules').then(r=>r.status)`);
    out.errors = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
    return out;
  });
  const cur = await snap(false);
  const old = await snap(true);
  t('OFF1 開關關閉：使用者端與後台都沒有任何本功能元素（沒有 .mar-*、#qRenewHost、鈴鐺群組、提醒規則側欄項／頁面／樣式表）', cur.marNodes === 0 && cur.bellMa === 0 && cur.renewHost === false && cur.formRenewNodes === 0 && cur.adminMar === 0 && !/ma-reminders-admin\.css/.test(cur.adminCssLinks), J([cur.marNodes, cur.bellMa, cur.renewHost, cur.formRenewNodes, cur.adminMar, cur.adminCssLinks]));
  t('OFF2 開關關閉：後台提醒規則 API 回 404、報價單設定 maTracking=false', cur.probe === 404 && cur.maTracking === false, J([cur.probe, cur.maTracking]));
  const filesDiffer = BASE_USER_FILES.concat(['admin.html']).some((f) => fs.readFileSync(path.join(C.base, '_client', f), 'utf8') !== fs.readFileSync(path.join(ROOT, '_client', f), 'utf8'));
  t('OFF3 比對基準與目前前端確實不同（防止比對流於空轉）；基準沒有 ITTSMaRem、目前有', filesDiffer && cur.rem === 'object' && old.rem === 'undefined', J([filesDiffer, cur.rem, old.rem]));
  t('OFF4 開關關閉：首頁 body（去掉 script／註解）與基準逐位元相同', cur.home === old.home && cur.home.length > 3000, cur.home.length + ' vs ' + old.home.length);
  t('OFF5 開關關閉：側欄 HTML 與基準逐位元相同', cur.sidebar === old.sidebar && cur.sidebar.length > 1000, cur.sidebar.length + ' vs ' + old.sidebar.length);
  t('OFF6 開關關閉：鈴鐺下拉內容與基準逐位元相同', cur.bell === old.bell && cur.bell.length > 100, cur.bell.length + ' vs ' + old.bell.length);
  t('OFF7 開關關閉：報價單頁、編輯表單、新增表單 DOM 與基準逐位元相同', cur.quotes === old.quotes && cur.formEdit === old.formEdit && cur.formNew === old.formNew && cur.formEdit.length > 3000 && cur.formNew.length > 3000, (cur.quotes !== old.quotes ? 'quotes@' + firstDiff(cur.quotes, old.quotes) : '') + (cur.formEdit !== old.formEdit ? ' edit@' + firstDiff(cur.formEdit, old.formEdit) : '') + (cur.formNew !== old.formNew ? ' new@' + firstDiff(cur.formNew, old.formNew) : ''));
  t('OFF8 開關關閉：後台 body（去掉 script／註解）與側欄和基準逐位元相同', cur.adminBody === old.adminBody && cur.adminBody.length > 5000 && cur.adminNav === old.adminNav, cur.adminBody.length + ' vs ' + old.adminBody.length);
  t('OFF9 開關關閉：頁面沒有 JS 例外', cur.errors.length === 0 && old.errors.length === 0, J(cur.errors.concat(old.errors)).slice(0, 300));
}

module.exports = { t, browse, login, openUser, openAdmin, gotoRules, pickOpt, typeInto, rowS, rulesInPage, shotTo, txt, reqs, mkUsers, seedData, mkItem, mkOpp, mkQ, approvedOf, U, tp, agoIso, NOW, XSS1, XSS2, AMT, CNO, PROD, MAP, LICP, THEMES, theme, contrastSel, CONTRAST_FN, PAGE_HELPERS, REGEX_ESC, J, CAT };
const X = module.exports;
// 其餘區段在 _ma-reminders-ui-sections.js（共用上面的小工具）
const SECS = () => require('./_ma-reminders-ui-sections')(X);

// ═════════════════════════ 執行 ═════════════════════════
function exportBase() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'marui-base-'));
  const a = spawnSync('git', ['archive', BASE_REF], { cwd: ROOT, maxBuffer: 512 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + BASE_REF + ' 失敗：' + String(a.stderr));
  const x = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 512 * 1024 * 1024 });
  if (x.status !== 0) throw new Error('tar 失敗：' + String(x.stderr));
  return { dir, cleanup() { try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function run(C, only) {
  const want = (n) => !only || only === n;
  const S = SECS();
  if (want('off')) await secOff(C);
  if (want('admin')) await S.secAdmin(C);
  if (want('dual')) await S.secDual(C);
  if (want('bell')) await S.secBell(C);
  if (want('items')) await S.secItems(C);
  if (want('quote')) await S.secQuote(C);
  if (want('look')) await S.secLook(C);
}
function report(title) {
  const bad = res.filter((x) => !x[1]);
  console.log('[' + title + '] ' + (res.length - bad.length) + '/' + res.length + ' 通過');
  return bad.length;
}

// UI 變異：用 CDP 把前端檔換成改壞的版本（不動磁碟上的檔案），對應區段必須變紅
const UI_MUTS = require('./_ma-reminders-ui-muts');
async function groupMut() {
  console.log('[UI 變異] 用 CDP 替換前端檔為改壞的版本，對應區段的測試必須變紅');
  const H = require('./_b2-harness');
  H.guardClean();
  const exb = exportBase();
  let undetected = 0;
  try {
    const only = process.env.MAREM_UI_MUT_ONLY ? new RegExp(process.env.MAREM_UI_MUT_ONLY) : null;
    for (const m of UI_MUTS) {
      if (only && !only.test(m.name)) continue;
      let src = fs.readFileSync(path.join(ROOT, '_client', m.file), 'utf8');
      if (src.indexOf(String.fromCharCode(13)) >= 0) src = src.split(String.fromCharCode(13, 10)).join(String.fromCharCode(10));
      const edits = m.edits || [{ find: m.find, rep: m.rep }];
      let miss = false;
      edits.forEach((e) => { if (!src.includes(e.find)) miss = true; });
      if (miss) { console.log('  ✗ 變異錨點不存在：' + m.name); undetected++; continue; }
      edits.forEach((e) => { src = src.replace(e.find, () => e.rep); });
      res = [];
      const quiet = console.log; console.log = () => {};
      try { await run({ H, ROOT, base: exb.dir, shots: null, cdp: CDP, mutFiles: { [m.file]: src } }, m.section); } catch (e) { t('(ui 例外) ' + String(e && e.message).slice(0, 100), false); }
      console.log = quiet;
      const failed = res.filter((x) => !x[1]);
      const hit = failed.some((x) => m.expect.test(x[0]));
      console.log((hit ? '  ✓ ' : '  ✗ ') + m.name + '：' + failed.length + ' 項變紅' + (hit ? '' : '（預期的測試沒變紅）') + (failed.length ? '｜例：' + failed[0][0].slice(0, 60) : ''));
      if (!hit) undetected++;
    }
  } finally { await H.finish(); exb.cleanup(); }
  console.log('UI 變異測試：' + (UI_MUTS.length - undetected) + '/' + UI_MUTS.length + ' 被偵測到');
  return undetected;
}
module.exports.run = run; module.exports.setRes = (r) => { res = r; };
if (require.main === module) {
  (async () => {
    if (process.env.MAREM_UI_MUT === '1') { process.exitCode = (await groupMut()) ? 1 : 0; return; }
    const H = require('./_b2-harness');
    H.guardClean();
    const exb = exportBase();
    res = [];
    try {
      await run({ H, ROOT, base: exb.dir, shots: process.env.MAREM_UI_SHOTS || path.join(ROOT, '..', 'ma4_shots'), cdp: CDP }, process.env.MAREM_UI_ONLY || '');
    } catch (e) { t('(ui 例外) ' + (e && e.stack || e).toString().slice(0, 400), false); }
    finally { await H.finish(); exb.cleanup(); }
    const n = report('ui');
    process.exitCode = n ? 1 : 0;
  })().catch((e) => { console.error(e && e.stack || e); process.exitCode = 1; });
}
