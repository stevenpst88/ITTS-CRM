'use strict';
/**
 * MA 納管前端（P2b）瀏覽器檢查：headless Edge＋CDP（真實滑鼠鍵盤；成交的拖曳 drop 用合成事件，因為 headless 無法產生原生拖放）。
 * 單獨執行：node scripts/_ma-tracking-ui.js            （全部區段；MATRK_UI_ONLY=off|flow|tasks|items|contracts|look 只跑一段）
 *          MATRK_UI_MUT=1 node scripts/_ma-tracking-ui.js （UI 變異：用 CDP 把前端檔換成改壞的版本，對應區段必須變紅）
 * 埠：伺服器 B2_PORT（預設 3313）、CDP 9610（MATRK_UI_CDP）。一次只跑一個。全部假資料、臨時 _s_* 帳號、sandbox 的 data.json／auth.json（不碰真實檔）。
 * 區段：
 *   off      開關關閉：頁面（首頁、側欄、看板、商機表、ERP／SAP 合約頁）渲染與基準前端（git archive）逐位元相同、成交請求相同、沒有任何 .mat-* 元素
 *   flow     成交期間視窗（看板拖曳／表格／編輯表單三個入口）：預填、取消回原欄、快捷、伺服器錯誤碼 TERM_REQUIRED／TERM_INVALID 重開視窗
 *   tasks    登入彙總視窗、首頁待辦卡、任務卡（簽約／重複編號／跳過／稍後／改派）、一級主管逾期清單、鈴鐺深連結、hash 路由、鍵盤與焦點
 *   items    到期管理頁（摘要卡、篩選、秘書欄位限制、待補納管、修改期間）
 *   contracts 合約管理「Service MA」頁籤（ERP／SAP 不受影響）、合約視窗唯讀來源
 *   look     XSS、四種外觀對比、手機寬度、reduced-motion、層級、截圖（scratchpad/ma2_shots/）
 *   fix      3020 驗收後修正（scripts/_ma-tracking-ui-fix.js，FX-*）：License MA 追蹤中不算待簽約、改派候選人、簽約／刪除合約後各處刷新與任務卡重開、
 *            兩頁剩餘天數一致（台北跨日、多時區）、再成交期間預填、到期管理窄視窗橫向捲動／卡片、後台商品維護儲存中攔截離開
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');
process.env.B2_PORT = process.env.B2_PORT || '3313';
const { withBrowser } = require('./_fd-cdp');

const ROOT = path.resolve(__dirname, '..');
const J = JSON.stringify;
const CDP = parseInt(process.env.MATRK_UI_CDP || '9610', 10);
const BASE_REF = process.env.MATRK_UI_BASE || '7efdba5';   // P2b 完成的 commit（驗收後修正之前）；也可用 MATRK_UI_BASE=8e8b99c 比對沒有 MA 前端的版本

let res = [];
function t(name, ok, extra) { res.push([name, !!ok, extra == null ? '' : String(extra)]); if (!ok) console.log('  ✗ ' + name + (extra != null && extra !== '' ? '｜' + String(extra).slice(0, 220) : '')); }

// ───────── 對比度（沿祖先逐層疊合半透明底）─────────
const CONTRAST_FN = `(function(el, dark){
  function parse(c){ var m=c.match(/rgba?\\(([^)]+)\\)/); if(!m) return [0,0,0,0]; var p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number); return [p[0],p[1],p[2],p.length>3?p[3]:1]; }
  function over(top,bot){ var a=top[3]; return [top[0]*a+bot[0]*(1-a), top[1]*a+bot[1]*(1-a), top[2]*a+bot[2]*(1-a), 1]; }
  function lum(c){ var f=function(v){v/=255; return v<=0.03928? v/12.92 : Math.pow((v+0.055)/1.055,2.4)}; return 0.2126*f(c[0])+0.7152*f(c[1])+0.0722*f(c[2]); }
  if(!el) return null;
  var chain=[]; for(var n=el;n&&n.nodeType===1;n=n.parentElement) chain.push(n);
  var bg=dark?[11,16,32,1]:[255,255,255,1];
  for(var i=chain.length-1;i>=0;i--){ var c=parse(getComputedStyle(chain[i]).backgroundColor); if(c[3]>0) bg=over(c,bg); }
  var fg=parse(getComputedStyle(el).color); fg=over(fg,bg);
  var L1=lum(fg), L2=lum(bg); var hi=Math.max(L1,L2), lo=Math.min(L1,L2);
  return Math.round(((hi+0.05)/(lo+0.05))*100)/100;
})`;
const contrastSel = (b, sel, dark) => b.ev(`${CONTRAST_FN}(document.querySelector(${J(sel)}), ${!!dark})`);
const contrastFn = (b, fnSrc, dark) => b.ev(`${CONTRAST_FN}((${fnSrc})(), ${!!dark})`);
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
const U = { adm: '_s_admin', own: '_s_own', oth: '_s_other', m1: '_s_m1', sec: '_s_sec', mk: '_s_mk', off: '_s_off', exe: '_s_exe' };
function mkUsers(H) {
  return [
    H.mkUser(U.adm, 'admin', { displayName: '測試管理員' }), H.mkUser(U.exe, 'executive'),
    H.mkUser(U.m1, 'manager1', { bu: ['ERP'], displayName: '測試主管' }),
    H.mkUser(U.own, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務甲' }),
    H.mkUser(U.oth, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務乙' }),
    H.mkUser(U.off, 'user', { bu: ['ERP'], supervisor: U.m1, active: false, displayName: '離職業務' }),
    H.mkUser(U.mk, 'marketing', { bu: ['ERP'], displayName: '測試行銷' }), H.mkUser(U.sec, 'secretary', { bu: ['ERP'] }),
  ];
}
const CAT = () => ({ bus: { ERP: [{ group: '假商品群組', items: [
  { name: 'Svc MA', termKind: 'ma', startsTask: true }, { name: 'Svc Lic', termKind: 'license', startsTask: false }, { name: 'Plain' },
] }] } });
const MAP = 'Svc MA 一年', LICP = 'Svc Lic 永久', PLAINP = 'Plain 專案';
const mkOpp = (id, owner, product, stage, company, extra) => Object.assign({ id, owner, bu: 'ERP', company, contactName: '假聯絡人', product, amount: '120', description: '', expectedDate: tp(30), stage, createdAt: NOW }, extra || {});
const qi = (lid, kind, from, to) => ({ lid, desc: '品項' + lid, unit: '式', qty: 1, unitPrice: 100, cost: 0, termKind: kind, termFrom: from, termTo: to });
const mkQuote = (id, oppId, owner, company, items) => ({ id, oppId, owner, quoteNo: 'QT-' + id, company, projectName: '', items, status: 'approved', discountType: 'none', discountValue: 0,
  approval: { state: 'approved', cur: 1, steps: [{ tier: 'mgr1', label: '核決', status: 'approved', by: U.m1, at: agoIso(5) }], hash: 'x', history: [] }, createdAt: NOW, updatedAt: agoIso(5) });
const mkItem = (id, opp, owner, company, product, to, extra) => Object.assign({ id, oppId: opp, quoteId: '', owner, company, product, termKind: 'ma', from: tp(-300), to, contractNo: '', contractId: '', status: 'tracking', source: 'request', createdAt: agoIso(10), updatedAt: agoIso(10), history: [] }, extra || {});
const mkTask = (id, opp, itemIds, owner, company, product, ageDays, extra) => Object.assign({ id, type: 'ma_sign', oppId: opp, maItemId: itemIds[0], maItemIds: itemIds, owner, company, product, status: 'open',
  steps: [{ key: 'sign', status: 'pending', value: null, doneBy: '', doneAt: '', reason: '' }], createdAt: agoIso(ageDays), updatedAt: agoIso(ageDays), history: [] }, extra || {});

function seedData(opts) {
  opts = opts || {};
  const d = {
    contacts: [], companies: [], visits: [], receivables: [], callins: [], keyAccounts: [], quotations: [
      mkQuote('q1', 'oK1', U.own, 'K一公司', [qi('l1', 'ma', '2026-11-01', '2027-10-31')]),
      mkQuote('q4', 'oK4', U.own, 'K四公司', [qi('l1', 'license', '2026-11-01', '2027-10-31'), qi('l2', 'ma', '2026-11-01', '2028-10-31')]),
    ],
    productCatalog: CAT(), groups: [],
    opportunities: [
      mkOpp('oK1', U.own, MAP, 'C', 'K一公司'), mkOpp('oK2', U.own, MAP, 'B', 'K二公司'), mkOpp('oK3', U.own, PLAINP, 'C', 'K三公司'), mkOpp('oK4', U.own, LICP, 'C', 'K四公司'),
      mkOpp('oK5', U.own, MAP, 'C', 'K五公司'), mkOpp('oK6', U.own, MAP, 'B', 'K六公司'), mkOpp('oK7', U.own, MAP, 'A', 'K七公司'), mkOpp('oK8', U.own, 'Svc MA', 'A', 'K八公司', { category: '假商品群組' }),
      mkOpp('oBF', U.own, MAP, 'Won', 'BF公司', { achievedDate: tp(-3) }),
      mkOpp('oA', U.own, MAP, 'Won', 'A公司', { achievedDate: tp(-3) }), mkOpp('oB', U.own, MAP, 'Won', 'B公司', { achievedDate: tp(-20) }),
      mkOpp('oC', U.own, MAP, 'Won', 'C公司', { achievedDate: tp(-40) }), mkOpp('oD', U.own, MAP, 'Won', 'D公司', { achievedDate: tp(-40) }),
      mkOpp('oOff', U.off, MAP, 'Won', 'E離職公司', { achievedDate: tp(-30) }),
      mkOpp('oX', U.own, 'Svc MA ' + XSS2, 'Won', XSS1 + 'X公司', { achievedDate: tp(-2) }),
      mkOpp('oO', U.oth, MAP, 'Won', 'O他人公司', { achievedDate: tp(-2) }),
    ],
    maItems: [
      mkItem('iA', 'oA', U.own, 'A公司', MAP, tp(10)), mkItem('iB', 'oB', U.own, 'B公司', MAP, tp(60)),
      mkItem('iC', 'oC', U.own, 'C公司', MAP, tp(200), { contractNo: 'C-2', contractId: 'cS' }), mkItem('iD', 'oD', U.own, 'D公司', LICP, tp(-5), { termKind: 'license', contractNo: 'C-3' }),
      mkItem('iOff', 'oOff', U.off, 'E離職公司', MAP, tp(120)),
      mkItem('iX', 'oX', U.own, XSS1 + 'X公司', 'Svc MA ' + XSS2, tp(80), { contractNo: XSS1 }),
      mkItem('iO', 'oO', U.oth, 'O他人公司', MAP, tp(30)),
    ],
    dealTasks: [
      mkTask('tA', 'oA', ['iA'], U.own, 'A公司', MAP, 3), mkTask('tB', 'oB', ['iB'], U.own, 'B公司', MAP, 20),
      mkTask('tOff', 'oOff', ['iOff'], U.off, 'E離職公司', MAP, 30), mkTask('tX', 'oX', ['iX'], U.own, XSS1 + 'X公司', 'Svc MA ' + XSS2, 2),
    ],
    contracts: [
      { id: 'cDUP', owner: U.own, contractNo: 'DUP-1', company: '重複測試公司', product: 'ERP MA', startDate: '2026-01-01', endDate: tp(100), amount: '1', type: 'ERP_MA', createdAt: NOW },
      { id: 'cERP', owner: U.own, contractNo: 'ERP-1', company: 'ERP合約公司', product: 'ERP MA', startDate: '2026-01-01', endDate: tp(100), amount: '5', type: 'ERP_MA', createdAt: NOW },
      { id: 'cSAP', owner: U.own, contractNo: 'SAP-1', company: 'SAP合約公司', product: 'SAP PCE License MA', startDate: '2026-01-01', endDate: tp(100), amount: '6', type: 'SAP_MA', createdAt: NOW },
      { id: 'cS', owner: U.own, contractNo: 'C-2', company: 'C公司', product: MAP, startDate: tp(-300), endDate: tp(200), amount: '120', type: 'SERVICE_MA', oppId: 'oC', quoteId: 'q-from', maItemId: 'iC', signedDate: tp(-30), salesPerson: '測試業務甲', createdAt: NOW },
    ],
  };
  if (opts.xssContract) d.contracts.push({ id: 'cX', owner: U.own, contractNo: XSS2, company: XSS1 + '合約公司', product: 'Svc MA ' + XSS2, startDate: tp(-100), endDate: tp(50), amount: '1', type: 'SERVICE_MA', oppId: XSS1, quoteId: XSS2, maItemId: XSS1, signedDate: tp(-1), createdAt: NOW });
  return d;
}

// ───────── 瀏覽器小工具 ─────────
function browse(C, maxMs, fn) {
  return withBrowser({ cdpPort: C.cdp, maxMs }, async (b) => {
    if (C.mutFiles) { Object.keys(C.mutFiles).forEach((k) => { b.replace['/' + k] = C.mutFiles[k]; }); await b.interceptOn(); }
    const oc = b.click; b.click = (sel, off) => oc(typeof sel === 'function' ? sel.toString() : sel, off);   // 允許傳函式當選取器
    return fn(b);
  });
}
const PAGE_HELPERS = `(()=>{
  window.__top=(sel,n)=>{const o=[...document.querySelectorAll('.mat-overlay')].pop(); if(!o) return null; const l=o.querySelectorAll(sel); return l[n||0]||null};
  window.__reqs=[]; if(!window.__recOn){ window.__recOn=true; const of=window.fetch.bind(window); window.fetch=(u,o)=>{ if(o&&o.method&&o.method!=='GET'){ window.__reqs.push({m:o.method,u:String(u),b:o.body||''}); } return of(u,o); }; }
  window.__xss=0; return 1})()`;
async function openAs(b, C, who, w, h, opts) {
  opts = opts || {};
  const H = C.H;
  await b.viewport(w || 1440, h || 900, (w || 1440) < 700);
  await b.nav(H.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(H.PW)}})}).then(r=>r.status)`);
  t('瀏覽器登入 ' + who, st === 200, st);
  await b.ev(opts.noLoginFlag ? `sessionStorage.removeItem('justLoggedIn'); 0` : `sessionStorage.setItem('justLoggedIn','1'); 0`);
  if (opts.lastSection !== undefined) await b.ev(`localStorage.setItem('lastSection', ${J(opts.lastSection)}); 0`);
  await b.nav(H.BASE + (opts.hash ? '/index.html' + opts.hash : '/index.html'));
  await b.waitFor(`typeof showSection==='function' && !!document.getElementById('navHome') && !!window.ITTSMaTasks`, 25000);
  if (opts.expectOn) await b.waitFor(`ITTSMaTasks._state.itemsOn`, 15000); else await b.sleep(opts.fast ? 800 : 2600);
  await b.ev(PAGE_HELPERS);
}
const topCount = (b) => b.ev(`document.querySelectorAll('.mat-overlay').length`);
const topText = (b) => b.ev(`(()=>{const o=[...document.querySelectorAll('.mat-overlay')].pop(); return o?o.innerText:''})()`);
const btnSel = (txt) => `()=>{const o=[...document.querySelectorAll('.mat-overlay')].pop(); return o?([...o.querySelectorAll('button')].find(x=>x.textContent.trim()===${J(txt)}&&!x.disabled)||null):null}`;
const clickBtn = (b, txt) => b.click(btnSel(txt));
const hasBtn = (b, txt) => b.ev(`(${btnSel(txt)})()!==null`);
async function focusTop(b, sel, n) { return b.ev(`(()=>{const e=__top(${J(sel)},${n || 0}); if(!e) return false; e.focus(); return document.activeElement===e})()`); }
// 年份欄最多吃 6 位數字（打 YYYYMMDD 會變成 202712-03-01），所以：年(4 碼) → 右鍵移到月 → 月(2 碼，自動跳日) → 日(2 碼)
async function typeDateTop(b, n, iso) { await focusTop(b, 'input[type=date]', n); const p = iso.split('-'); await b.type(p[0]); await b.key('ArrowRight'); await b.type(p[1] + p[2]); await b.sleep(80); }
const dateVals = (b) => b.ev(`(()=>{const a=__top('input[type=date]',0), z=__top('input[type=date]',1); return [a&&a.value, z&&z.value]})()`);
const topErr = (b) => b.ev(`(()=>{const e=__top('.mat-error'); return e&&!e.hidden?e.textContent:''})()`);
const reqs = (b) => b.ev(`JSON.stringify(window.__reqs)`).then((s) => JSON.parse(s));
async function drop(b, oppId, stage) {
  await b.ev(`(()=>{ dragOppId=${J(oppId)}; document.getElementById('kanbanCards'+${J(stage)}).dispatchEvent(new Event('drop',{bubbles:true,cancelable:true})); return 0})()`);
}
const oppOf = (H, who, id) => H.http(who, 'GET', '/api/opportunities').then((r) => (Array.isArray(r.json) ? r.json : []).find((o) => o.id === id));
const shotTo = async (C, b, name) => { if (C.shots) { fs.mkdirSync(C.shots, { recursive: true }); await b.shot(path.join(C.shots, name)); } };
// 成交後等慶祝動畫與「轉為客戶」提示（3.6 秒後出現）過去，再收掉它們，避免蓋住後面的點擊
async function calm(b) {
  await b.sleep(4000);
  await b.ev(`(()=>{const e=document.getElementById('wonCelebration'); if(e) e.style.display='none'; document.querySelectorAll('.modal-overlay.open').forEach(x=>x.classList.remove('open')); return 0})()`);
}
async function toBoard(b) {
  await b.ev(`showSection('pipeline'); 0`);
  await b.waitFor(`document.querySelectorAll('#kanbanCardsC .kanban-card, #kanbanCardsC > div').length>0`, 12000);
  await b.sleep(500);
}

// ═════════════════════════ off：開關關閉＝與基準逐位元相同 ═════════════════════════
const BASE_FILES = ['app.js', 'style.css', 'index.html'];
const STRIP = `(h)=>h.replace(/<script[\\s\\S]*?<\\/script>/g,'').replace(/<!--[\\s\\S]*?-->/g,'').replace(/\\s+/g,' ')`;
async function secOff(C) {
  const { H } = C;
  await H.fresh(mkUsers(H), seedData(), { env: {} });   // 資料裡有納管項／任務卡，但開關未設：畫面必須當它們不存在
  await H.loginAll([U.own]);
  const snap = (useBase) => withBrowser({ cdpPort: C.cdp, maxMs: 300000 }, async (b) => {
    if (C.mutFiles && !useBase) Object.keys(C.mutFiles).forEach((k) => { b.replace['/' + k] = C.mutFiles[k]; });
    if (useBase) BASE_FILES.forEach((f) => { b.replace['/' + f] = fs.readFileSync(path.join(C.base, '_client', f), 'utf8'); });
    if (useBase || C.mutFiles) await b.interceptOn();
    const H2 = C.H;
    await b.viewport(1440, 900, false);
    await b.nav(H2.BASE + '/login.html');
    await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(U.own)},password:${J(H2.PW)}})}).then(r=>r.status)`);
    await b.ev(`sessionStorage.setItem('justLoggedIn','1'); 0`);
    await b.nav(H2.BASE + '/index.html');
    await b.waitFor(`typeof showSection==='function' && !!document.getElementById('navHome')`, 25000);
    await b.sleep(3200);
    await b.ev(PAGE_HELPERS);
    const out = {};
    const norm = `(${STRIP})(document.body.innerHTML)`;
    out.ma = await b.ev(`typeof window.ITTSMaTasks`);
    out.home = await b.ev(norm);
    out.sidebar = await b.ev(`(${STRIP})(document.getElementById('sidebarNav').innerHTML)`);
    out.matNodes = await b.ev(`document.querySelectorAll('[class*="mat-"], #navMaItems, #navServiceMa, #maItemsView, #serviceMaView, [data-mat-host]').length`);
    for (const s of ['pipeline', 'targets', 'erp-ma', 'sap-ma']) {
      await b.ev(`showSection(${J(s)}); 0`);
      await b.sleep(1800);
      out['sec_' + s] = await b.ev(norm);
      out['title_' + s] = await b.ev(`document.getElementById('topbarTitle').textContent`);
    }
    // 成交請求（看板 drop 與 updateOppStage）：stub 掉 PUT，記下請求
    await b.ev(`(()=>{ window.__puts=[]; const of=window.fetch.bind(window); window.fetch=(u,o)=>{ if(o&&o.method==='PUT'&&/opportunities\\//.test(String(u))){ window.__puts.push(String(u).split('/').pop()+'|'+String(o.body).replace(/"achievedDate":"[0-9-]+"/,'"achievedDate":"D"')); return Promise.resolve(new Response(JSON.stringify({id:'x'}),{status:200,headers:{'Content-Type':'application/json'}})); } return of(u,o); }; return 1})()`);
    await b.ev(`showSection('pipeline'); 0`); await b.sleep(1500);
    await drop(b, 'oK1', 'Won'); await b.sleep(900);
    await b.ev(`updateOppStage('oK2','Won'); 0`); await b.sleep(900);
    out.puts = await b.ev(`JSON.stringify(window.__puts)`);
    out.dialogs = await b.ev(`document.querySelectorAll('.mat-overlay').length`);
    out.errors = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
    return out;
  });
  const cur = await snap(false);
  const old = await snap(true);
  t('OFF1 開關關閉：前端不建立任何本功能元素（沒有 .mat-*、側欄項、view、首頁 host），也沒有成交視窗', cur.matNodes === 0 && cur.dialogs === 0, 'nodes=' + cur.matNodes + ' dialogs=' + cur.dialogs);
  // 基準早於 P2b（沒有 MA 前端）：基準沒有 ITTSMaTasks、目前有；基準已含 P2b 前端（預設 7efdba5）：兩邊都有，但檔案內容必須真的不同（防止比對流於空轉）
  const baseHasMa = old.ma === 'object';
  const filesDiffer = BASE_FILES.concat(['ma-tasks.js']).some((f) => { try { return fs.readFileSync(path.join(C.base, '_client', f), 'utf8') !== fs.readFileSync(path.join(ROOT, '_client', f), 'utf8'); } catch { return true; } });
  t('OFF2 開關關閉：比對基準與目前前端確實不同（防止比對流於空轉）', baseHasMa ? (cur.ma === 'object' && filesDiffer) : (old.ma === 'undefined' && cur.ma === 'object'), old.ma + '/' + cur.ma + '/differ=' + filesDiffer);
  t('OFF3 開關關閉：首頁 body（去掉 script／註解）與基準逐位元相同', cur.home === old.home && cur.home.length > 3000, cur.home.length + ' vs ' + old.home.length);
  t('OFF4 開關關閉：側欄 HTML 與基準逐位元相同', cur.sidebar === old.sidebar && cur.sidebar.length > 1000, cur.sidebar.length + ' vs ' + old.sidebar.length);
  for (const s of ['pipeline', 'targets', 'erp-ma', 'sap-ma']) {
    t('OFF5 開關關閉：區段 ' + s + ' 的頁面與標題和基準逐位元相同', cur['sec_' + s] === old['sec_' + s] && cur['title_' + s] === old['title_' + s] && cur['sec_' + s].length > 3000, (cur['sec_' + s].length) + ' vs ' + old['sec_' + s].length);
  }
  t('OFF6 開關關閉：看板拖到 Won 與表格 updateOppStage 送出的 PUT 與基準逐位元相同（沒有 termFrom／termTo）', cur.puts === old.puts && /Won/.test(cur.puts) && !/term/.test(cur.puts), cur.puts + ' vs ' + old.puts);
  t('OFF7 開關關閉：頁面沒有 JS 例外', cur.errors.length === 0 && old.errors.length === 0, J(cur.errors.concat(old.errors)).slice(0, 300));
}

// ═════════════════════════ flow：成交期間視窗 ═════════════════════════
async function secFlow(C) {
  const { H } = C;
  await H.fresh(mkUsers(H), seedData(), { env: { MA_TRACKING: '1' } });
  await H.loginAll([U.own, U.adm]);
  await browse(C, 500000, async (b) => {
    await openAs(b, C, U.own, 1440, 900, { expectOn: true, noLoginFlag: true });
    await toBoard(b);
    await b.ev(`window.__stubDone=0; 0`);

    // ── 入口 1：看板拖曳 ──
    await drop(b, 'oK3', 'Won');   // 一般商品：不跳視窗，直接成交
    await b.sleep(1200);
    t('FL1 看板：一般商品（Plain）拖到 Won 不跳期間視窗，直接成交', (await topCount(b)) === 0 && (await oppOf(H, U.own, 'oK3')).stage === 'Won', await topCount(b));
    await b.ev(`document.getElementById('wonCelebration') && (document.getElementById('wonCelebration').style.display='none'); 0`);
    await toBoard(b);
    await drop(b, 'oK1', 'Won');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    const txt1 = await topText(b);
    const v1 = await dateVals(b);
    t('FL2 看板：License／MA 商品拖到 Won → 跳出「成交：請填寫服務期間」視窗，說明含公司與商品', /成交：請填寫服務期間/.test(txt1) && /K一公司/.test(txt1) && /Svc MA 一年/.test(txt1), txt1.slice(0, 120));
    t('FL3 已連結核准報價單且有期間 → 預填 2026-11-01～2027-10-31，並註明來源', v1[0] === '2026-11-01' && v1[1] === '2027-10-31' && /已帶入核准報價單 QT-q1/.test(txt1), J(v1));
    t('FL4 視窗顯示「共 12 個月」', /共 12 個月/.test(txt1), txt1);
    // 取消 → 不改階段、沒有 PUT
    const before = (await reqs(b)).filter((r) => /opportunities/.test(r.u)).length;
    await clickBtn(b, '取消');
    await b.sleep(500);
    const afterCancel = await oppOf(H, U.own, 'oK1');
    t('FL5 取消：視窗關閉、伺服器階段仍是 C、沒有送出 PUT、看板卡片仍在 C 欄', (await topCount(b)) === 0 && afterCancel.stage === 'C' && (await reqs(b)).filter((r) => /opportunities/.test(r.u)).length === before
      && (await b.ev(`!!document.querySelector('#kanbanCardsC')&&document.querySelector('#kanbanCardsC').innerText.indexOf('K一公司')>=0`)), J(afterCancel.stage));
    // Esc 也算取消
    await drop(b, 'oK1', 'Won');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    await b.key('Escape'); await b.sleep(400);
    t('FL6 Esc＝取消：階段不變', (await topCount(b)) === 0 && (await oppOf(H, U.own, 'oK1')).stage === 'C');
    // 調整後確認：改迄日
    await drop(b, 'oK1', 'Won');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    await typeDateTop(b, 1, '2027-12-31');
    const v1b = await dateVals(b);
    await clickBtn(b, '確認成交');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    const celeb = await b.waitFor(`(()=>{const e=document.getElementById('wonCelebration'); return !!e && e.style.display!=='none'})()`, 5000);
    await b.sleep(300);
    const rq = (await reqs(b)).filter((r) => r.m === 'PUT' && /opportunities\/oK1/.test(r.u)).pop();
    const body1 = rq ? JSON.parse(rq.b) : {};
    t('FL7 預填後可調整：真實鍵盤改迄日 → PUT 帶 stage=Won、termFrom=2026-11-01、termTo=2027-12-31', v1b[1] === '2027-12-31' && body1.stage === 'Won' && body1.termFrom === '2026-11-01' && body1.termTo === '2027-12-31', rq && rq.b);
    const m1 = await H.http(U.own, 'GET', '/api/ma-items?filter=all');
    const it1 = (m1.json.items || []).find((x) => x.company === 'K一公司');
    // 慶祝動畫只持續約 3 秒，所以在成交後 0.5 秒內檢查（下面的 FL8 先用 HTTP 讀資料，已經夠快）
    t('FL8 成交後伺服器建立納管項（期間與視窗一致）並有任務卡', it1 && it1.from === '2026-11-01' && it1.to === '2027-12-31' && !!it1.taskId, J(it1));
    t('FL9 成交後既有慶祝動畫照常出現（wonCelebration 顯示）', celeb);
    await calm(b);

    // ── 沒有報價單：欄位為空；缺欄位按確認 → 就地紅字；快捷 ──
    await toBoard(b);
    await drop(b, 'oK2', 'Won');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    const v2 = await dateVals(b);
    t('FL10 沒有核准報價單：起訖日欄位為空', v2[0] === '' && v2[1] === '', J(v2));
    await clickBtn(b, '一年');
    t('FL11 沒有起日按「一年」：提示先填起日、不填迄日', /請先填寫起日/.test(await topErr(b)) && (await dateVals(b))[1] === '', await topErr(b));
    await clickBtn(b, '確認成交');
    t('FL12 兩欄都空按確認成交：不關視窗、就地紅字（role=alert）', (await topCount(b)) === 1 && /服務期間/.test(await topErr(b)), await topErr(b));
    await typeDateTop(b, 0, '2026-12-01');
    await clickBtn(b, '三年');
    const v2b = await dateVals(b);
    t('FL13 起日 2026-12-01 + 三年 → 迄日 2029-11-30、顯示「共 36 個月」', v2b[1] === '2029-11-30' && /共 36 個月/.test(await topText(b)), J(v2b));
    await typeDateTop(b, 1, '2026-01-01');
    await clickBtn(b, '確認成交');
    t('FL14 迄日早於起日：不送出、紅字「迄日不得早於起日」', (await topCount(b)) === 1 && /迄日不得早於起日/.test(await topErr(b)), await topErr(b));
    await clickBtn(b, '一年');
    await clickBtn(b, '確認成交');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(1200);
    t('FL15 修正後成交成功，伺服器階段 Won', (await oppOf(H, U.own, 'oK2')).stage === 'Won');
    await calm(b);

    // ── 多組期間：兩欄留空＝依報價單各組 ──
    await toBoard(b);
    await drop(b, 'oK4', 'Won');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    const txt4 = await topText(b);
    t('FL16 報價單有 2 組不同期間：說明列出兩組、欄位留空、可不填直接確認', /有 2 組不同的服務期間/.test(txt4) && (await dateVals(b)).join('') === '', txt4.slice(0, 200));
    await typeDateTop(b, 0, '2026-11-01');
    await clickBtn(b, '確認成交');
    t('FL17 多組情況只填一欄：要求補齊', (await topCount(b)) === 1 && /服務期間/.test(await topErr(b)), await topErr(b));
    await b.ev(`(()=>{const a=__top('input[type=date]',0); a.value=''; a.dispatchEvent(new Event('input',{bubbles:true})); return 0})()`);
    await clickBtn(b, '確認成交');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(1200);
    const m4 = await H.http(U.own, 'GET', '/api/ma-items?filter=all');
    t('FL18 留空確認：依報價單建立 2 個納管項（License 與 MA 各一組期間）', (m4.json.items || []).filter((x) => x.company === 'K四公司').length === 2, J((m4.json.items || []).filter((x) => x.company === 'K四公司').map((x) => x.termKind + x.to)));
    await calm(b);

    // ── 伺服器錯誤碼：前端以為不需要（目錄被 stub 成空）→ PUT 回 TERM_REQUIRED → 開同一個視窗並顯示訊息 ──
    await toBoard(b);
    await b.ev(`(()=>{ const t=window.ITTSMaTasks._state; t.catalog={bus:{}}; t.catalogAt=Date.now()+600000; return 1})()`);
    await drop(b, 'oK5', 'Won');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    const txt5 = await topText(b);
    t('FL19 伺服器回 400 TERM_REQUIRED → 開成交期間視窗並顯示伺服器訊息', /成交：請填寫服務期間/.test(txt5) && /服務期間/.test(await topErr(b)) && (await oppOf(H, U.own, 'oK5')).stage === 'C', await topErr(b));
    await clickBtn(b, '取消'); await b.sleep(400);
    t('FL20 此時取消：階段不變', (await oppOf(H, U.own, 'oK5')).stage === 'C');
    await drop(b, 'oK5', 'Won');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    await typeDateTop(b, 0, '2026-12-01'); await clickBtn(b, '一年'); await clickBtn(b, '確認成交');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(1200);
    t('FL21 補填後重送成功', (await oppOf(H, U.own, 'oK5')).stage === 'Won');
    await calm(b);

    // TERM_INVALID：把第一次 PUT 的迄日改成早於起日
    await toBoard(b);
    await b.ev(`(()=>{ const t=window.ITTSMaTasks._state; t.catalog={bus:{}}; t.catalogAt=Date.now()+600000; return 1})()`);
    await b.ev(`(()=>{ const of=window.fetch.bind(window); let n=0; window.fetch=(u,o)=>{ if(o&&o.method==='PUT'&&/opportunities\\/oK6/.test(String(u))){ const bd=JSON.parse(o.body); if(!bd.termFrom && n++===0){ bd.termFrom='2026-12-01'; bd.termTo='2026-01-01'; o=Object.assign({},o,{body:JSON.stringify(bd)}); } } return of(u,o); }; return 1})()`);
    await drop(b, 'oK6', 'Won');   // 第一次：沒有期間 → 被改成顛倒 → TERM_INVALID
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    t('FL22 伺服器回 400 TERM_INVALID → 同一個視窗，顯示伺服器訊息（不得早於起日）', /不得早於起日|不合法|有效日期/.test(await topErr(b)), await topErr(b));
    await clickBtn(b, '取消'); await b.sleep(300);
    await b.ev(`ITTSMaTasks._state.catalogAt=0; 0`);   // 還原目錄快取（前面為了走錯誤碼路徑把它 stub 成空的）

    // ── 入口 2：表格 updateOppStage ──
    await b.ev(`showSection('targets'); 0`);
    await b.waitFor(`document.querySelectorAll('.opp-stage-sel').length>0`, 12000);
    await b.sleep(400);
    await b.ev(`(()=>{ const s=document.querySelector('.opp-stage-sel[data-id="oK7"]'); s.value='Won'; s.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`);
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    t('FL23 表格：把商機階段改成 Won → 跳同一個期間視窗', /成交：請填寫服務期間/.test(await topText(b)) && /K七公司/.test(await topText(b)));
    await clickBtn(b, '取消'); await b.sleep(600);
    t('FL24 表格取消：伺服器階段不變，下拉選單還原為原階段', (await oppOf(H, U.own, 'oK7')).stage === 'A' && (await b.ev(`document.querySelector('.opp-stage-sel[data-id="oK7"]').value`)) === 'A');
    await b.ev(`(()=>{ const s=document.querySelector('.opp-stage-sel[data-id="oK7"]'); s.value='Won'; s.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`);
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    await typeDateTop(b, 0, '2027-01-01'); await clickBtn(b, '一年'); await clickBtn(b, '確認成交');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(1500);
    t('FL25 表格確認：成交，納管項期間 2027-01-01～2027-12-31', (await oppOf(H, U.own, 'oK7')).stage === 'Won' && ((await H.http(U.own, 'GET', '/api/ma-items?filter=all')).json.items || []).some((x) => x.company === 'K七公司' && x.to === '2027-12-31'));
    await calm(b);

    // ── 入口 3：商機編輯表單 ──
    await b.ev(`showSection('pipeline'); 0`); await b.sleep(800);
    await b.ev(`openOppEdit('oK8'); 0`);
    await b.waitFor(`document.getElementById('oppEditStage') && document.getElementById('oppEditStage').offsetParent!==null`, 8000);
    await b.sleep(900);   // 編輯視窗開啟後會非同步補欄位，太早改階段會被重設
    await b.ev(`(()=>{ const s=document.getElementById('oppEditStage'); s.value='Won'; s.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`);
    await b.click('#oppEditSave');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    t('FL26 編輯表單：存成 Won → 跳期間視窗', /成交：請填寫服務期間/.test(await topText(b)) && /K八公司/.test(await topText(b)));
    await clickBtn(b, '取消'); await b.sleep(500);
    t('FL27 表單取消：編輯視窗仍開著、伺服器階段不變', (await b.ev(`document.getElementById('oppEditStage').offsetParent!==null`)) && (await oppOf(H, U.own, 'oK8')).stage === 'A');
    await b.click('#oppEditSave');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    await typeDateTop(b, 0, '2027-02-01'); await clickBtn(b, '一年'); await clickBtn(b, '確認成交');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(1500);
    t('FL28 表單確認：儲存並成交，編輯視窗關閉', (await oppOf(H, U.own, 'oK8')).stage === 'Won' && (await b.ev(`document.getElementById('oppEditStage').offsetParent===null`)));
    const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
    t('FL29 流程中沒有 JS 例外', fe.length === 0, J(fe).slice(0, 300));
  });
}


// ═════════════════════════ tasks：登入彙總、首頁待辦、任務卡、鈴鐺、逾期清單 ═════════════════════════
const SKIP_REASON = '客戶改用其他方案 <img src=x onerror=window.__xss=3>';
const rowBtnSel = (company, txt) => `()=>{const r=[...document.querySelectorAll('[data-mat-host="dash"] .mat-panel .mat-row')].find(x=>x.querySelector('.mat-name')&&x.querySelector('.mat-name').textContent===${J(company)}); return r?([...r.querySelectorAll('button')].find(x=>x.textContent.trim()===${J(txt)})||null):null}`;
const sumRowBtnSel = (company, txt) => `()=>{const o=[...document.querySelectorAll('.mat-overlay')].pop(); const r=o&&[...o.querySelectorAll('.mat-row')].find(x=>x.querySelector('.mat-name')&&x.querySelector('.mat-name').textContent===${J(company)}); return r?([...r.querySelectorAll('button')].find(x=>x.textContent.trim()===${J(txt)})||null):null}`;
const panelRows = (b) => b.ev(`[...document.querySelectorAll('[data-mat-host="dash"] .mat-panel-tasks .mat-row .mat-name')].map(x=>x.textContent)`);
async function secTasks(C) {
  const { H } = C;
  const data = seedData();
  data.notifications = [{ id: 'n1', to: U.own, type: 'deal_task', title: '📝 請確認簽約：A公司', body: 'A公司 的 Svc MA 一年 已成交，請確認客戶是否簽約', refId: 'tA', read: false, createdAt: NOW }];
  await H.fresh(mkUsers(H), data, { env: { MA_TRACKING: '1' } });
  await H.loginAll([U.own, U.adm, U.m1, U.mk]);
  await browse(C, 560000, async (b) => {
    // ── 登入彙總視窗 ──
    await openAs(b, C, U.own, 1440, 900, { expectOn: true });
    const hasSum = await b.waitFor(`!!document.querySelector('.mat-dialog[role=alertdialog]')`, 12000);
    const sumTxt = hasSum ? await topText(b) : '';
    t('TK-1a 剛登入且有開啟中任務卡：自動跳彙總視窗（alertdialog），標題「你有 3 筆成交案待確認簽約」，列出公司', hasSum && /你有 3 筆成交案待確認簽約/.test(sumTxt) && /A公司/.test(sumTxt) && /B公司/.test(sumTxt), sumTxt.slice(0, 150));
    t('TK-1b 彙總視窗：逾期（超過 14 天）的 B公司 排在最前面並標「已逾 14 天」', /已逾 14 天/.test(sumTxt) && sumTxt.indexOf('B公司') < sumTxt.indexOf('A公司'), sumTxt.slice(0, 200));
    t('TK-1c 彙總視窗焦點在視窗內（第一個主要按鈕），可略過', await b.ev(`!!document.activeElement && !!document.activeElement.closest('.mat-dialog')`));
    await b.key('Escape'); await b.sleep(300);
    t('TK-1d Esc＝略過：視窗關閉', (await topCount(b)) === 0);
    // 同一次登入不重複跳：切換區段再回首頁
    await b.ev(`showSection('visits'); 0`); await b.sleep(300); await b.ev(`showDashboard(); 0`); await b.sleep(1500);
    t('TK-1e 略過後同一次登入不再自動跳', (await topCount(b)) === 0);
    // ── 首頁待辦卡 ──
    const rows = await panelRows(b);
    t('TK-2a 首頁「待辦事項：確認簽約」卡片列出 3 筆（A公司、B公司、XSS 公司；他人與離職者的不列）', rows.length === 3 && rows.includes('A公司') && rows.includes('B公司') && !rows.includes('E離職公司'), J(rows));
    t('TK-2b 公司名一律當純文字顯示（XSS 公司名原樣出現在文字中）', rows.some((x) => x.indexOf('<img src=x onerror') >= 0) && (await b.ev('window.__xss')) === 0);
    // ── 鈴鐺：deal_task 通知 → 開任務彙總 ──
    await b.click('#notifBellBtn');
    await b.waitFor(`!!document.querySelector('#notifList .notif-item[data-type="deal_task"]')`, 8000);
    t('TK-B1 鈴鐺有 deal_task 通知，圖示 📝', await b.ev(`document.querySelector('#notifList .notif-item[data-type="deal_task"] .notif-icon').textContent.trim()`) === '📝');
    await b.click('#notifList .notif-item[data-type="deal_task"]');
    const opened = await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    t('TK-B2 點 deal_task 通知 → 關掉下拉並開啟「待確認簽約」彙總視窗；通知標為已讀', opened && /待確認簽約/.test(await topText(b)) && (await H.http(U.own, 'GET', '/api/poll-bundle')).json.notifications.every((n) => n.type !== 'deal_task' || n.read), (await topText(b)).slice(0, 60));
    await b.key('Escape'); await b.sleep(300);

    // ── 任務卡 tA ──
    await b.click(rowBtnSel('A公司', '處理…'));
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    const tt = await topText(b);
    t('TK-3a 任務卡：顯示公司、品項、服務期間、負責業務', /成交後確認簽約：A公司/.test(tt) && /Svc MA 一年/.test(tt) && /\d{4}\/\d{2}\/\d{2} – \d{4}\/\d{2}\/\d{2}/.test(tt) && /測試業務甲/.test(tt), tt.slice(0, 200));
    const steps = await b.ev(`JSON.stringify([...__top('.mat-steps').children].map(li=>li.getAttribute('data-state')+':'+li.querySelector('.mat-step-l').textContent+':'+li.querySelector('.mat-step-s').textContent))`);
    t('TK-3b 狀態步驟條：成交✓ → 簽約確認(進行中) → 歸入 Service MA 管理(待辦) → 到期前提醒(待辦)，狀態用文字呈現', steps === J(['done:成交:✓ 完成', 'current:簽約確認:進行中', 'todo:歸入 Service MA 管理:待辦', 'todo:到期前提醒:待辦']), steps);
    t('TK-3c 業務本人看得到：關閉／稍後提醒／此案不簽合約…／合約已簽訂，輸入編號；看不到「改派…」', (await hasBtn(b, '稍後提醒')) && (await hasBtn(b, '此案不簽合約…')) && (await hasBtn(b, '合約已簽訂，輸入編號')) && !(await hasBtn(b, '改派…')));
    // 焦點圈住：Tab 12 次都留在視窗內
    let inside = true;
    for (let i = 0; i < 12; i++) { await b.key('Tab'); if (!(await b.ev(`!!document.activeElement && !!document.activeElement.closest('.mat-dialog')`))) inside = false; }
    t('TK-3d Tab 12 次焦點都留在對話框內（焦點圈住）', inside);
    // ── 簽約確認 ──
    await clickBtn(b, '合約已簽訂，輸入編號');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===2`, 5000);
    await clickBtn(b, '確認簽約');
    t('TK-4a 合約編號空白：就地紅字「請輸入合約編號」，不呼叫伺服器', /請輸入合約編號/.test(await topErr(b)) && !(await reqs(b)).some((r) => /\/sign/.test(r.u)), await topErr(b));
    await focusTop(b, 'input[type=text]'); await b.type('DUP-1');
    await clickBtn(b, '確認簽約');
    await b.waitFor(`(()=>{const e=__top('.mat-error'); return !!e && !e.hidden && e.textContent.length>0})()`, 5000);
    t('TK-4b 合約編號與既有合約重複：顯示後端 409 訊息「這個合約編號已經被使用，請確認編號是否正確」，視窗不關', (await topErr(b)) === '這個合約編號已經被使用，請確認編號是否正確' && (await topCount(b)) === 2, await topErr(b));
    // Esc 只關最上層（簽約視窗），任務卡視窗仍在，焦點回到開啟它的按鈕
    await b.key('Escape'); await b.sleep(300);
    t('TK-4c Esc 只關最上層：簽約視窗關閉、任務卡視窗還在；焦點回到「合約已簽訂，輸入編號」按鈕', (await topCount(b)) === 1 && (await b.ev(`document.activeElement && document.activeElement.textContent.trim()`)) === '合約已簽訂，輸入編號');
    await clickBtn(b, '合約已簽訂，輸入編號');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===2`, 5000);
    await focusTop(b, 'input[type=text]'); await b.type('SVC-NEW-1');
    await clickBtn(b, '確認簽約');
    await b.waitFor(`[...document.querySelectorAll('.mat-dlg-title')].some(x=>/已歸入合約管理/.test(x.textContent))`, 8000);
    const okTxt = await topText(b);
    t('TK-4d 簽約成功：顯示「已歸入合約管理 → Service MA」與合約編號，步驟條前三關完成', /已歸入合約管理 → Service MA/.test(okTxt) && /SVC-NEW-1/.test(okTxt) && /done|✓/.test(okTxt) && (await topCount(b)) === 1, okTxt.slice(0, 160));
    const ct = (await H.http(U.own, 'GET', '/api/contracts')).json.find((c) => c.contractNo === 'SVC-NEW-1');
    const tk = (await H.http(U.own, 'GET', '/api/my-tasks/deals')).json;
    t('TK-4e 伺服器：建立 SERVICE_MA 合約並連結商機；任務卡完成、不在待辦', ct && ct.type === 'SERVICE_MA' && ct.oppId === 'oA' && !(tk.items || []).some((x) => x.id === 'tA'), J(ct && { t: ct.type, o: ct.oppId }));
    await clickBtn(b, '關閉'); await b.sleep(500);
    t('TK-4f 完成後首頁待辦卡自動刷新（剩 2 筆，A公司不見）', J((await panelRows(b)).sort()) === J(['B公司', XSS1 + 'X公司'].sort()), J(await panelRows(b)));
    // ── 此案不簽合約（XSS 原因）──
    await b.click(rowBtnSel('B公司', '處理…'));
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    await clickBtn(b, '此案不簽合約…');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===2`, 5000);
    await clickBtn(b, '送出');
    t('TK-5a 不簽合約原因空白：紅字「請填寫…原因」，不送出', /請填寫不簽合約的原因/.test(await topErr(b)));
    await focusTop(b, 'textarea'); await b.type('不');
    await clickBtn(b, '送出');
    t('TK-5b 原因只有 1 個字：仍擋下（至少 2 個字）', /至少 2 個字/.test(await topErr(b)));
    await b.ev(`(()=>{const e=__top('textarea'); e.value=${J(SKIP_REASON)}; e.dispatchEvent(new Event('input',{bubbles:true})); return 0})()`);
    await clickBtn(b, '送出');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(900);
    const tkB = (await H.http(U.adm, 'GET', '/api/manager/deal-task-overdue')).json;
    t('TK-5c 送出後兩個視窗都關、首頁卡片少一筆；伺服器任務卡 skipped 並存下原因；XSS 原因沒有執行', !(await panelRows(b)).includes('B公司') && !(tkB.items || []).some((x) => x.id === 'tB') && (await b.ev('window.__xss')) === 0, J(await panelRows(b)));
    // ── 稍後提醒 ──
    await b.click(rowBtnSel(XSS1 + 'X公司', '處理…'));
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    await clickBtn(b, '稍後提醒');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(900);
    const sn = (await H.http(U.own, 'GET', '/api/my-tasks/deals')).json.items.find((x) => x.id === 'tX');
    t('TK-6 稍後提醒：任務卡仍開啟但 snoozed、snoozedUntil 為 3 天後；首頁卡片該列標「已延後」', sn && sn.snoozed === true && sn.snoozedUntil === tp(3) && (await b.ev(`[...document.querySelectorAll('[data-mat-host="dash"] .mat-panel-tasks .mat-chip-mute')].length`)) === 1, J(sn && sn.snoozedUntil));
    t('TK-6b 全部操作過程沒有 XSS 執行', (await b.ev('window.__xss')) === 0);

    // ── 一級主管：逾期清單（含已停用負責人）；不可改派 ──
    await openAs(b, C, U.m1, 1440, 900, { expectOn: true, noLoginFlag: true });
    await b.ev(`showSection('managerHome'); 0`);
    await b.waitFor(`!!document.querySelector('.mat-panel-overdue')`, 10000);
    const od = await b.ev(`document.querySelector('.mat-panel-overdue').innerText`);
    t('TK-7a 一級主管首頁：「成交後超過 14 天未簽約確認」清單列出團隊逾期任務（E離職公司，標「已停用」）；已跳過的 B公司 不在', /E離職公司/.test(od) && /已停用/.test(od) && !/B公司/.test(od), od.slice(0, 200));
    t('TK-7b 一級主管沒有「改派…」按鈕（僅管理員可改派）', (await b.ev(`[...document.querySelectorAll('.mat-panel-overdue button')].length`)) === 0);
    // ── 管理員：改派 ──
    await openAs(b, C, U.adm, 1440, 900, { expectOn: true, noLoginFlag: true });
    await b.ev(`showSection('managerHome'); 0`);
    await b.waitFor(`!!document.querySelector('.mat-panel-overdue')`, 10000);
    t('TK-8a 管理員的逾期清單有「改派…」按鈕', (await b.ev(`[...document.querySelectorAll('.mat-panel-overdue button')].filter(x=>x.textContent.trim()==='改派…').length`)) >= 1);
    await b.click(() => { const r = [...document.querySelectorAll('.mat-panel-overdue .mat-row')].find((x) => /E離職公司/.test(x.innerText)); return r && r.querySelector('button'); });
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    await b.waitFor(`!!__top('select') && __top('select').options.length>0`, 8000);
    const optTxt = await b.ev(`[...__top('select').options].map(o=>o.value).join(',')`);
    t('TK-8b 改派對話框：下拉只列在職業務（user）、同 BU：含 _s_other／_s_own，不含目前負責人、停用者、主管、行銷、秘書、管理員、其他 BU', optTxt.split(',').sort().join() === ['_s_other', '_s_own'].join(), optTxt);
    await b.ev(`(()=>{const s=__top('select'); s.value='_s_other'; s.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`);
    await clickBtn(b, '改派');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(600);
    const tOff = (await H.http(U.adm, 'GET', '/api/manager/deal-task-overdue')).json.items.find((x) => x.id === 'tOff');
    t('TK-8c 改派成功：伺服器任務負責人改為 _s_other，清單刷新後負責人顯示新的人', tOff && tOff.owner === '_s_other' && (await b.ev(`/測試業務乙/.test(document.querySelector('.mat-panel-overdue').innerText)`)), J(tOff && tOff.owner));

    // ── 行銷：不簽合約的通知 → 點擊開到期管理頁 ──
    await openAs(b, C, U.mk, 1440, 900, { expectOn: true, noLoginFlag: true });
    await b.click('#notifBellBtn');
    await b.waitFor(`!!document.querySelector('#notifList .notif-item[data-type="deal_task_skipped"]')`, 8000);
    t('TK-B3 行銷鈴鐺有 deal_task_skipped 通知，圖示 ⚠️，內文原因當純文字（沒有 <img> 元素、沒有執行）', (await b.ev(`document.querySelector('#notifList .notif-item[data-type="deal_task_skipped"] .notif-icon').textContent.trim()`)) === '⚠️' && (await b.ev(`document.querySelectorAll('#notifList img').length`)) === 0 && (await b.ev('window.__xss')) === 0);
    await b.click('#notifList .notif-item[data-type="deal_task_skipped"]');
    const pg = await b.waitFor(`(()=>{const v=document.getElementById('maItemsView'); return !!v && !v.hidden})()`, 8000);
    t('TK-B4 點 deal_task_skipped 通知 → 開「MA／License 到期管理」頁（標題與側欄反白）', pg && (await b.ev(`document.getElementById('topbarTitle').textContent`)) === 'MA／License 到期管理' && (await b.ev(`document.getElementById('navMaItems').classList.contains('active')`)));
    // ── hash 路由 ──
    await openAs(b, C, U.own, 1440, 900, { expectOn: true, noLoginFlag: true, hash: '#dealtasks' });
    t('TK-H1 網址 #dealtasks → 開任務彙總視窗，並清掉 hash', (await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000)) && (await b.ev(`location.hash`)) === '');
    await b.key('Escape');
    await openAs(b, C, U.own, 1440, 900, { expectOn: true, noLoginFlag: true, hash: '#maitems' });
    t('TK-H2 網址 #maitems → 開到期管理頁', await b.waitFor(`(()=>{const v=document.getElementById('maItemsView'); return !!v && !v.hidden})()`, 8000));
    // F5 還原到到期管理頁（lastSection 存了 maItems）
    await b.ev(`localStorage.setItem('lastSection','maItems'); 0`);
    await b.nav(H.BASE + '/index.html');
    t('TK-H3 F5 還原到「到期管理」：啟動後補顯示該頁（不是空白）', await b.waitFor(`(()=>{const v=document.getElementById('maItemsView'); return !!v && !v.hidden})()`, 15000));
    const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
    t('TK-9 沒有 JS 例外', fe.length === 0, J(fe).slice(0, 300));
  });
}

// ═════════════════════════ items：到期管理頁 ═════════════════════════
const countsOf = (b) => b.ev(`JSON.stringify([...document.querySelectorAll('#maItemsView .mat-card')].map(c=>c.querySelector('.mat-card-l').textContent+':'+c.querySelector('.mat-card-n').textContent))`).then(JSON.parse);
const rowCos = (b) => b.ev(`[...document.querySelectorAll('#maItemsView tbody tr')].map(r=>r.querySelector('td strong')?r.querySelector('td strong').textContent:'')`);
async function gotoItemsNav(b) {
  await b.click('#sgContractHd');   // 側欄分組要先展開
  await b.sleep(500);
  await b.click('#navMaItems');
  await b.waitFor(`(()=>{const v=document.getElementById('maItemsView'); return !!v && !v.hidden && document.querySelectorAll('#maItemsView tbody tr').length>0})()`, 10000);
  await b.sleep(300);
}
async function secItems(C) {
  const { H } = C;
  await H.fresh(mkUsers(H), seedData(), { env: { MA_TRACKING: '1' } });
  await H.loginAll([U.own, U.sec, U.mk, U.adm]);
  await browse(C, 560000, async (b) => {
    await openAs(b, C, U.own, 1440, 900, { expectOn: true, noLoginFlag: true });
    await gotoItemsNav(b);
    t('IT-1 側欄「MA／License 到期管理」只在後端啟用且 /api/ma-items 回 200 時才出現；點選後顯示頁面、標題、側欄反白', (await b.ev(`document.getElementById('topbarTitle').textContent`)) === 'MA／License 到期管理' && (await b.ev(`document.getElementById('navMaItems').classList.contains('active')`)));
    const cs = await countsOf(b);
    t('IT-2 摘要卡：納管中 5、待簽約 2、90 天內到期 3、已逾期 1（與資料一致）', J(cs) === J(['納管中:5', '待簽約:2', '90 天內到期:3', '已逾期:1']), J(cs));
    const hdr = await b.ev(`[...document.querySelectorAll('#maItemsView thead th')].map(x=>x.textContent).join('|')`);
    t('IT-3 欄位：客戶、品項、合約編號、期間、剩餘天數、狀態／下一步、負責業務、操作', hdr === '客戶|品項|合約編號|期間|剩餘天數|狀態／下一步|負責業務|操作', hdr);
    const all = await rowCos(b);
    t('IT-4 「全部」列出 5 筆；依迄日由近到遠（已逾期 D公司 在最前）', all.length === 5 && all[0] === 'D公司', J(all));
    const f = async (label) => { await b.ev(`(()=>{const x=[...document.querySelectorAll('#maItemsView .mat-fchip')].find(c=>c.textContent===${J(label)}); x.click(); return 0})()`); await b.sleep(900); return rowCos(b); };
    const fp = await f('待簽約'); const f9 = await f('90 天內到期'); const fo = await f('已逾期'); const fa = await f('全部');
    t('IT-5 篩選晶片：待簽約→2 筆(A、B)、90 天內→3 筆、已逾期→1 筆(D)、全部→5 筆；按下的晶片 aria-pressed=true', fp.length === 2 && fp.includes('A公司') && fp.includes('B公司') && f9.length === 3 && fo.length === 1 && fo[0] === 'D公司' && fa.length === 5
      && (await b.ev(`[...document.querySelectorAll('#maItemsView .mat-fchip')].find(c=>c.textContent==='全部').getAttribute('aria-pressed')`)) === 'true', J([fp, f9, fo, fa.length]));
    const dRow = await b.ev(`(()=>{const r=[...document.querySelectorAll('#maItemsView tbody tr')].find(x=>/D公司/.test(x.innerText)); return JSON.stringify({txt:r.innerText, bar:!!r.querySelector('progress.mat-bar-bad'), cls:r.className})})()`).then(JSON.parse);
    t('IT-6 剩餘天數：逾期列有紅色進度條與「已逾期 5 天」，整列標紅；合約編號與狀態顯示', /已逾期 5 天/.test(dRow.txt) && dRow.bar && /mat-tr-bad/.test(dRow.cls) && /C-3/.test(dRow.txt), dRow.txt.replace(/\s+/g, ' '));
    const aRow = await b.ev(`[...document.querySelectorAll('#maItemsView tbody tr')].find(x=>/A公司/.test(x.innerText)).innerText.replace(/\\s+/g,' ')`);
    t('IT-7 待簽約列：狀態顯示「待簽約確認」、有「處理簽約…」與「修改期間…」按鈕；進度條顏色依剩餘天數（10 天＝紅 mat-bar-bad、60 天＝黃 mat-bar-warn、200 天＝綠 mat-bar-ok）', /待簽約確認/.test(aRow) && /處理簽約…/.test(aRow) && /修改期間…/.test(aRow)
      && (await b.ev(`(()=>{const f=(n)=>[...document.querySelectorAll('#maItemsView tbody tr')].find(x=>x.innerText.indexOf(n)>=0).querySelector('progress').className; return JSON.stringify([f('A公司'),f('B公司'),f('C公司')])})()`)) === J(['mat-bar mat-bar-bad', 'mat-bar mat-bar-warn', 'mat-bar mat-bar-ok']), aRow);
    // 處理簽約 → 任務卡
    await b.click(() => [...[...document.querySelectorAll('#maItemsView tbody tr')].find((x) => /A公司/.test(x.innerText)).querySelectorAll('button')].find((x) => x.textContent === '處理簽約…'));
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    t('IT-8 「處理簽約…」開啟該筆的任務卡視窗', /成交後確認簽約：A公司/.test(await topText(b)));
    await clickBtn(b, '關閉'); await b.sleep(300);
    // 修改期間
    await b.click(() => [...[...document.querySelectorAll('#maItemsView tbody tr')].find((x) => /A公司/.test(x.innerText)).querySelectorAll('button')].find((x) => x.textContent === '修改期間…'));
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    const dv = await dateVals(b);
    t('IT-9 修改期間視窗預填目前起迄日，有「修改原因（選填）」', dv[0] && dv[1] && /修改原因/.test(await topText(b)), J(dv));
    await typeDateTop(b, 1, '2025-01-01');
    await clickBtn(b, '儲存期間');
    t('IT-10 迄日早於起日：不送出，就地紅字', (await topCount(b)) === 1 && /迄日不得早於起日/.test(await topErr(b)), await topErr(b));
    const newTo = tp(25);
    await typeDateTop(b, 1, newTo);
    await focusTop(b, 'textarea'); await b.type('客戶延後啟用');
    await clickBtn(b, '儲存期間');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(900);
    const ia = ((await H.http(U.own, 'GET', '/api/ma-items?filter=all')).json.items || []).find((x) => x.company === 'A公司');
    t('IT-11 儲存期間：伺服器 PATCH 成功、表格刷新為「剩 25 天」', ia && ia.to === newTo && (await b.ev(`/剩 25 天/.test([...document.querySelectorAll('#maItemsView tbody tr')].find(x=>/A公司/.test(x.innerText)).innerText)`)), J(ia && ia.to));
    const audit = await H.getLogs(U.adm);
    t('IT-12 修改期間留下稽核（MA_ITEM_TERM_EDIT，含原因）', audit.some((l) => l.action === 'MA_ITEM_TERM_EDIT' && /客戶延後啟用/.test(l.detail || '')), J(audit.filter((l) => /MA_/.test(l.action)).map((l) => l.action)).slice(0, 200));
    // 待補納管
    const bf = await b.ev(`(()=>{const x=document.querySelector('#maItemsView .mat-page-actions button'); return x&&!x.hidden?x.textContent:''})()`);
    t('IT-13 「待補納管（1）」按鈕出現（已成交但沒納管的 BF公司）', bf === '待補納管（1）', bf);
    await b.click('#maItemsView .mat-page-actions button');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    t('IT-14 待補納管面板列出 BF公司 與成交日', /BF公司/.test(await topText(b)) && /成交日/.test(await topText(b)));
    await clickBtn(b, '補登…');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===2`, 5000);
    await typeDateTop(b, 0, '2026-11-01'); await clickBtn(b, '一年');
    await clickBtn(b, '補登納管');
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
    await b.sleep(1200);
    const nb = ((await H.http(U.own, 'GET', '/api/ma-items?filter=all')).json.items || []).find((x) => x.company === 'BF公司');
    t('IT-15 補登成功：納管項建立（2026-11-01～2027-10-31）、面板關閉、「待補納管」按鈕消失、表格多一筆', nb && nb.from === '2026-11-01' && nb.to === '2027-10-31' && (await b.ev(`document.querySelector('#maItemsView .mat-page-actions button').hidden`)) && (await rowCos(b)).includes('BF公司'), J(nb && [nb.from, nb.to]));

    // ── 秘書：唯讀、無合約編號與金額 ──
    await openAs(b, C, U.sec, 1440, 900, { expectOn: true, noLoginFlag: true });
    await gotoItemsNav(b);
    const sh = await b.ev(`[...document.querySelectorAll('#maItemsView thead th')].map(x=>x.textContent).join('|')`);
    const sTxt = await b.ev(`document.getElementById('maItemsView').innerText`);
    const sCore = await b.ev(`['.mat-cards','.mat-filters','table'].map(q=>document.querySelector('#maItemsView '+q).innerText).join(' ')`);
    t('IT-S1 秘書欄位限制：只有 客戶|品項|期間|剩餘天數|負責業務，沒有合約編號、狀態、操作', sh === '客戶|品項|期間|剩餘天數|負責業務', sh);
    t('IT-S2 秘書畫面沒有合約編號（C-2／C-3／XSS 編號）、沒有按鈕、沒有「待補納管」、沒有「待簽約」晶片與卡片', !/C-2|C-3|合約編號|金額/.test(sCore) && (await b.ev(`document.querySelectorAll('#maItemsView tbody button').length`)) === 0 && (await b.ev(`document.querySelector('#maItemsView .mat-page-actions button').hidden`)) && !/待簽約/.test(sCore), sCore.slice(0, 200));
    t('IT-S3 秘書唯讀說明文字與摘要（清單筆數／90 天內／已逾期）', /唯讀/.test(sTxt) && /清單筆數/.test(sTxt));
    // ── 行銷：看得到全部、不能改 ──
    await openAs(b, C, U.mk, 1440, 900, { expectOn: true, noLoginFlag: true });
    await gotoItemsNav(b);
    t('IT-M1 行銷看得到所有業務的納管項，但沒有「修改期間…」「處理簽約…」按鈕', (await rowCos(b)).includes('O他人公司') && (await b.ev(`[...document.querySelectorAll('#maItemsView tbody button')].filter(x=>x.textContent.trim()!=='提醒時間軸').length`)) === 0);
    const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
    t('IT-9z 沒有 JS 例外', fe.length === 0, J(fe).slice(0, 300));
  });
}

// ═════════════════════════ contracts：合約管理 Service MA ═════════════════════════
async function secContracts(C) {
  const { H } = C;
  await H.fresh(mkUsers(H), seedData(), { env: { MA_TRACKING: '1' } });
  await H.loginAll([U.own]);
  await browse(C, 400000, async (b) => {
    await openAs(b, C, U.own, 1440, 900, { expectOn: true, noLoginFlag: true });
    await b.click('#sgContractHd'); await b.sleep(500);   // 側欄分組要先展開
    await b.click('#navContractGroup'); await b.sleep(400);
    t('CT-1 合約管理下多一個「Service MA」子項（既有 ERP MA 管理、SAP License MA 仍在）', (await b.ev(`['navErpMa','navSapMa','navServiceMa'].map(i=>document.getElementById(i)?document.getElementById(i).querySelector('.nav-label').textContent.trim():'')`)).join('|') === 'ERP MA 管理|SAP License MA|Service MA', await b.ev(`document.getElementById('subMenuContract').innerText`));
    await b.click('#navServiceMa');
    await b.waitFor(`(()=>{const v=document.getElementById('serviceMaView'); return !!v && !v.hidden && document.querySelectorAll('#serviceMaTbody tr').length>0})()`, 10000);
    await b.sleep(400);
    const nos = await b.ev(`[...document.querySelectorAll('#serviceMaTbody tr')].map(r=>r.querySelector('td strong')?r.querySelector('td strong').textContent:'')`);
    t('CT-2 Service MA 頁籤只列 type==SERVICE_MA 的合約（C-2），不含 ERP-1／DUP-1／SAP-1', J(nos) === J(['C-2']), J(nos));
    t('CT-3 標題「Service MA 合約管理」、側欄反白、說明橫幅', (await b.ev(`document.getElementById('topbarTitle').textContent`)) === 'Service MA 合約管理' && (await b.ev(`document.getElementById('navServiceMa').classList.contains('active')`)) && /來源商機與到期管理項不可在此修改/.test(await b.ev(`document.getElementById('serviceMaView').innerText`)));
    const row = await b.ev(`document.querySelector('#serviceMaTbody tr').innerText.replace(/\\s+/g,' ')`);
    t('CT-4 列內容：客戶、產品、起迄日、年費金額、業務、狀態、剩餘天數；沒有「續約」按鈕（Service MA 走到期管理）', /C公司/.test(row) && /剩餘 20[01] 天/.test(row) && /120/.test(row) && !/續約/.test(row) && /編輯/.test(row), row);
    t('CT-5 來源連結：「來源商機」「到期管理項」按鈕', (await b.ev(`[...document.querySelectorAll('#serviceMaTbody .mat-src button')].map(x=>x.textContent).join('|')`)) === '來源商機|到期管理項');
    const st = await b.ev(`document.getElementById('serviceMaStatRow').innerText.replace(/\\s+/g,' ')`);
    t('CT-6 統計卡（合約總數 1，年費合計 120K）', /合約總數 1/.test(st) && /120/.test(st), st);
    // 搜尋
    await b.focus('#serviceMaSearch'); await b.type('zzz'); await b.sleep(400);
    t('CT-7 搜尋不到時顯示 Service MA 專用的空訊息', /尚無 Service MA 合約|沒有/.test(await b.ev(`document.getElementById('serviceMaTbody').innerText`)) || (await b.ev(`document.querySelectorAll('#serviceMaTbody tr').length`)) <= 1);
    await b.ev(`(()=>{const e=document.getElementById('serviceMaSearch'); e.value=''; e.dispatchEvent(new Event('input',{bubbles:true})); return 0})()`); await b.sleep(300);
    // 合約視窗：唯讀來源
    await b.ev(`window.__reqs=[]; 0`);
    await b.click('#serviceMaTbody .ct-btn-edit');
    await b.waitFor(`document.getElementById('contractModal').classList.contains('open')`, 5000);
    const box = await b.ev(`(()=>{const x=document.getElementById('contractSourceBox'); return x?x.innerText.replace(/\\s+/g,' '):''})()`);
    t('CT-8 開啟 SERVICE_MA 合約：標題「編輯 Service MA 合約」，唯讀顯示來源商機／報價單／到期管理項／簽約日', (await b.ev(`document.getElementById('contractModalTitle').textContent`)) === '編輯 Service MA 合約' && /oC/.test(box) && /q-from/.test(box) && /iC/.test(box) && /簽約日/.test(box) && /不可修改/.test(box), box);
    t('CT-9 來源值是 <code> 純文字，不是輸入欄位（前端不能改）', (await b.ev(`document.querySelectorAll('#contractSourceBox input, #contractSourceBox textarea, #contractSourceBox select').length`)) === 0);
    await b.click('#contractSaveBtn');
    await b.waitFor(`window.__reqs.some(r=>r.m==='PUT')`, 8000);
    const put = JSON.parse((await reqs(b)).find((r) => r.m === 'PUT').b);
    t('CT-10 儲存 SERVICE_MA 合約：送出的本文 type=SERVICE_MA，且不含 oppId／quoteId／maItemId（由伺服器寫入，前端不得送出）', put.type === 'SERVICE_MA' && !('oppId' in put) && !('quoteId' in put) && !('maItemId' in put), J(Object.keys(put)));
    const after = (await H.http(U.own, 'GET', '/api/contracts')).json.find((c) => c.id === 'cS');
    t('CT-11 儲存後伺服器端的來源連結仍在（沒被洗掉）', after && after.oppId === 'oC' && after.maItemId === 'iC' && after.type === 'SERVICE_MA', J(after && [after.oppId, after.maItemId]));
    await b.sleep(500);
    // 來源商機
    await b.click('#serviceMaTbody .mat-src button');
    t('CT-12 「來源商機」開啟該筆商機的編輯視窗', await b.waitFor(`document.getElementById('oppEditId') && document.getElementById('oppEditId').value==='oC'`, 8000));
    await b.ev(`closeOppEdit(); 0`);
    // 既有頁籤不受影響
    await b.ev(`showSection('erp-ma'); 0`); await b.sleep(1200);
    const erp = await b.ev(`[...document.querySelectorAll('#contractTbody tr')].map(r=>r.querySelector('td strong')?r.querySelector('td strong').textContent:'')`);
    t('CT-13 ERP MA 頁籤只列 ERP-1、DUP-1（不含 Service MA）；且仍有「續約」按鈕、沒有來源連結', erp.sort().join('|') === 'DUP-1|ERP-1' && (await b.ev(`document.querySelectorAll('#contractTbody .ct-btn-renew').length`)) === 2 && (await b.ev(`document.querySelectorAll('#contractTbody .mat-src').length`)) === 0, J(erp));
    await b.ev(`openContractModal('cERP'); 0`); await b.sleep(300);
    t('CT-14 開 ERP 合約：標題「編輯 ERP MA 合約」、沒有來源區塊', (await b.ev(`document.getElementById('contractModalTitle').textContent`)) === '編輯 ERP MA 合約' && (await b.ev(`!document.getElementById('contractSourceBox')`)));
    await b.ev(`closeContractModal(); showSection('sap-ma'); 0`); await b.sleep(1200);
    const sap = await b.ev(`[...document.querySelectorAll('#sapContractTbody tr')].map(r=>r.querySelector('td strong')?r.querySelector('td strong').textContent:'')`);
    t('CT-15 SAP License MA 頁籤只列 SAP-1', J(sap) === J(['SAP-1']), J(sap));
    const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
    t('CT-16 沒有 JS 例外', fe.length === 0, J(fe).slice(0, 300));
  });
}


// ═════════════════════════ look：XSS、外觀、手機、reduced-motion、層級 ═════════════════════════
const MAT_NODES = `[...document.querySelectorAll('*')].filter(e=>[...e.classList].some(c=>c.indexOf('mat-')===0))`;
async function secLook(C) {
  const { H } = C;
  const data = seedData({ xssContract: true });
  await H.fresh(mkUsers(H), data, { env: { MA_TRACKING: '1' } });
  await H.loginAll([U.own]);
  await browse(C, 700000, async (b) => {
    await openAs(b, C, U.own, 1440, 900, { expectOn: true });
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 12000);
    // ── XSS：每個會顯示使用者資料的畫面都檢查 ──
    const xssState = async (label) => {
      const bad = await b.ev(`JSON.stringify({x:window.__xss, imgs:document.querySelectorAll('.mat-overlay img, [data-mat-host] img, #maItemsView img, #serviceMaView img, #contractSourceBox img, #notifList img').length, scripts:document.querySelectorAll('.mat-overlay script, [data-mat-host] script, #maItemsView script, #serviceMaView script, #contractSourceBox script').length})`).then(JSON.parse);
      t('LK-X ' + label + '：XSS 字串（<img onerror>、<script>）不執行、不產生元素', bad.x === 0 && bad.imgs === 0 && bad.scripts === 0, J(bad));
    };
    const sumTxt = await topText(b);
    t('LK-X0 登入彙總視窗以純文字顯示含 HTML 的公司名', sumTxt.indexOf('<img src=x onerror') >= 0);
    await xssState('登入彙總視窗');
    await b.key('Escape'); await b.sleep(300);
    await b.ev(`showDashboard(); 0`); await b.sleep(1200);
    await xssState('首頁待辦卡');
    await b.click(rowBtnSel(XSS1 + 'X公司', '處理…'));
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    await xssState('任務卡視窗（公司名、品項含 HTML）');
    const taskTxt = await topText(b);
    t('LK-X1 任務卡的品項（含 <script> 字串）以純文字顯示', taskTxt.indexOf('<script>window.__xss=2</script>') >= 0, taskTxt.slice(0, 150));
    await b.key('Escape'); await b.sleep(300);
    await b.ev(`showSection('maItems'); 0`);
    await b.waitFor(`document.querySelectorAll('#maItemsView tbody tr').length>0`, 8000);
    await xssState('到期管理頁（合約編號、公司、品項含 HTML）');
    t('LK-X2 到期管理頁表格以純文字顯示 HTML 合約編號', (await b.ev(`document.getElementById('maItemsView').innerText`)).indexOf('<img src=x onerror') >= 0);
    await b.ev(`showSection('service-ma'); 0`);
    await b.waitFor(`document.querySelectorAll('#serviceMaTbody tr').length>1`, 8000);
    await xssState('Service MA 頁籤');
    await b.ev(`openContractModal('cX'); 0`); await b.sleep(400);
    await xssState('合約視窗來源區塊（oppId／quoteId／maItemId 含 HTML）');
    t('LK-X3 來源區塊顯示 HTML 原字串（純文字）', (await b.ev(`document.getElementById('contractSourceBox').innerText`)).indexOf('<img src=x onerror') >= 0);
    await b.ev(`closeContractModal(); 0`);
    // ── 沒有 inline style / on* 屬性（只掃本檔建立的 .mat-* 節點；既有合約表的舊 inline style 不算）──
    const inl = await b.ev(`JSON.stringify((${MAT_NODES}).filter(e=>e.hasAttribute('style') || [...e.attributes].some(a=>/^on/i.test(a.name))).map(e=>e.className).slice(0,5))`);
    t('LK-I 本檔建立的 .mat-* 節點沒有 style 屬性、沒有 on* 屬性', inl === '[]', inl);

    // ── 層級 ──
    const z = await b.ev(`(()=>{ const mk=(c)=>{const e=document.createElement('div'); e.className=c; document.body.appendChild(e); const v=parseInt(getComputedStyle(e).zIndex,10); e.remove(); return v}; return JSON.stringify({mat:mk('mat-overlay'), cct:mk('cct-overlay'), creq:mk('creq-overlay')}) })()`).then(JSON.parse);
    t('LK-Z 層級：.mat-overlay(100000) 低於 company-card-tasks(100001) 與 company-requests(100002)，且高於一般 modal(<=10000)', z.mat < z.cct && z.cct < z.creq && z.mat > 10000, J(z));

    // ── 四種外觀的對比（>=4.5:1）與截圖 ──
    const measure = async (th, label, sels) => {
      const dark = th.endsWith('dark'); const bad = [];
      for (const [sel, nm] of sels) {
        const cr = await b.ev(`${CONTRAST_FN}(${sel.startsWith('()') ? '(' + sel + ')()' : 'document.querySelector(' + J(sel) + ')'}, ${dark})`);
        if (cr === null || cr < 4.5) bad.push(nm + '=' + cr);
      }
      t('LK-C ' + label + '（' + th + '）：文字對比 >= 4.5:1', bad.length === 0, bad.join('、'));
    };
    for (const th of THEMES) {
      await theme(b, th);
      await b.ev(`showDashboard(); 0`); await b.sleep(900);
      await measure(th, '首頁待辦卡', [['.mat-panel-title', '標題'], ['.mat-panel .mat-sub', '說明'], ['.mat-panel .mat-name', '公司'], ['.mat-panel .mat-meta', '品項期間'], ['.mat-panel .mat-count', '數量'], ['.mat-panel .mat-chip', '標籤'], ['.mat-panel .mat-chip-warn', '逾期標籤'], ['.mat-panel .mat-btn-primary', '處理按鈕']]);
      await shotTo(C, b, 'ma_home_' + th + '.png');
      await b.click(rowBtnSel('A公司', '處理…')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      await measure(th, '任務卡視窗', [['.mat-overlay .mat-dlg-title', '標題'], ['.mat-overlay .mat-lead', '說明'], ['.mat-overlay .mat-step-done .mat-step-l', '步驟(完成)'], ['.mat-overlay .mat-step-done .mat-step-s', '步驟狀態(完成)'], ['.mat-overlay .mat-step-current .mat-step-l', '步驟(進行中)'], ['.mat-overlay .mat-step-current .mat-step-s', '步驟狀態(進行中)'], ['.mat-overlay .mat-step-todo .mat-step-l', '步驟(待辦)'], ['.mat-overlay .mat-step-todo .mat-step-s', '步驟狀態(待辦)'], ['.mat-overlay .mat-dt', '欄位名'], ['.mat-overlay .mat-dd', '欄位值'], ['.mat-overlay .mat-btn-ghost', '次要按鈕'], ['.mat-overlay .mat-btn-danger-ghost', '不簽約按鈕'], ['.mat-overlay .mat-btn-primary', '主要按鈕']]);
      await shotTo(C, b, 'ma_task_' + th + '.png');
      await clickBtn(b, '合約已簽訂，輸入編號'); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===2`, 5000);
      await clickBtn(b, '確認簽約');
      await measure(th, '簽約視窗', [['()=>__top(".mat-field-label")', '欄位標籤'], ['()=>__top(".mat-input")', '輸入框'], ['()=>__top(".mat-error")', '錯誤訊息'], ['()=>__top(".mat-btn-primary")', '確認按鈕']]);
      await b.key('Escape'); await b.sleep(200); await b.key('Escape'); await b.sleep(300);
      // 成交期間視窗（含伺服器訊息與預填）
      await b.ev(`ITTSMaTasks.wonTerm({id:'oK1',company:'外觀測試公司',product:'Svc MA 一年',stage:'C'},{force:true,message:'測試用的伺服器訊息'}); 0`);
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
      await b.ev(`(()=>{const a=__top('input[type=date]',0); a.value='2026-11-01'; a.dispatchEvent(new Event('input',{bubbles:true})); const z=__top('input[type=date]',1); z.value='2027-10-31'; z.dispatchEvent(new Event('input',{bubbles:true})); return 0})()`);
      await measure(th, '成交期間視窗', [['()=>__top(".mat-lead")', '說明'], ['()=>__top(".mat-note")', '註記'], ['()=>__top(".mat-input")', '日期輸入'], ['()=>__top(".mat-span")', '共 N 個月'], ['()=>__top(".mat-quick .mat-btn")', '快捷按鈕'], ['()=>__top(".mat-error")', '錯誤訊息'], ['()=>__top(".mat-btn-primary")', '確認按鈕']]);
      await shotTo(C, b, 'ma_won_' + th + '.png');
      await b.key('Escape'); await b.sleep(300);
      await b.ev(`showSection('maItems'); 0`);
      await b.waitFor(`document.querySelectorAll('#maItemsView tbody tr').length>0`, 8000); await b.sleep(400);
      await measure(th, '到期管理頁', [['#maItemsView .mat-page-title', '標題'], ['#maItemsView .mat-sub', '說明'], ['#maItemsView .mat-card:not(.mat-card-warn):not(.mat-card-bad) .mat-card-n', '摘要卡數字'], ['#maItemsView .mat-card:not(.mat-card-warn):not(.mat-card-bad) .mat-card-l', '摘要卡標籤'], ['#maItemsView .mat-card-warn .mat-card-n', '警示卡數字'], ['#maItemsView .mat-card-warn .mat-card-l', '警示卡標籤'], ['#maItemsView .mat-card-bad .mat-card-n', '逾期卡數字'], ['#maItemsView .mat-card-bad .mat-card-l', '逾期卡標籤'], ['#maItemsView .mat-fchip:not(.mat-fchip-on)', '晶片'], ['#maItemsView .mat-fchip-on', '晶片(選取)'], ['#maItemsView thead th', '表頭'], ['#maItemsView tbody td strong', '客戶'], ['#maItemsView tbody td:nth-child(4)', '期間'], ['#maItemsView .mat-days', '剩餘天數'], ['#maItemsView .mat-days-bad', '逾期天數'], ['#maItemsView tbody .mat-chip', '類型標籤'], ['#maItemsView .mat-btn-primary', '處理簽約'], ['#maItemsView tbody .mat-btn-ghost', '修改期間'], ['#maItemsView .mat-tr-bad td strong', '逾期列客戶']]);
      await shotTo(C, b, 'ma_items_' + th + '.png');
      await b.ev(`showSection('service-ma'); 0`);
      await b.waitFor(`document.querySelectorAll('#serviceMaTbody tr').length>0`, 8000); await b.sleep(400);
      await measure(th, 'Service MA 頁籤', [['#serviceMaView .mat-notice', '說明橫幅'], ['#serviceMaTbody .mat-link', '來源連結'], ['#serviceMaTbody .mat-src-days', '來源列剩餘天數']]);
      await shotTo(C, b, 'ma_servicema_' + th + '.png');
      await b.ev(`openContractModal('cS'); 0`); await b.sleep(400);
      await measure(th, '合約視窗來源區塊', [['#contractSourceBox', '區塊'], ['#contractSourceBox .mat-srcbox-title', '標題'], ['#contractSourceBox .mat-srcbox-k', '欄位名'], ['#contractSourceBox .mat-srcbox-v', '欄位值']]);
      await shotTo(C, b, 'ma_contract_' + th + '.png');
      await b.ev(`closeContractModal(); 0`);
    }
    await theme(b, 'glass-light');

    // ── reduced-motion ──
    await b.ev(`showDashboard(); 0`); await b.sleep(600);
    await b.click(rowBtnSel('A公司', '處理…')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    const animOn = await b.ev(`getComputedStyle(__top('.mat-dialog').parentNode).animationName`);
    await b.key('Escape'); await b.sleep(200);
    await b.media([{ name: 'prefers-reduced-motion', value: 'reduce' }]);
    await b.click(rowBtnSel('A公司', '處理…')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
    const animOff = await b.ev(`getComputedStyle(__top('.mat-dialog').parentNode).animationName`);
    await b.key('Escape'); await b.sleep(200);
    await b.media([{ name: 'prefers-reduced-motion', value: 'no-preference' }]);
    t('LK-R prefers-reduced-motion: reduce 時視窗淡入動畫關閉（一般時 mat-fade、reduce 時 none）', animOn === 'mat-fade' && animOff === 'none', animOn + '/' + animOff);

    // ── 手機寬度（375×812）──
    await openAs(b, C, U.own, 375, 812, { expectOn: true });
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 12000);
    const mob = await b.ev(`(()=>{const d=__top('.mat-dialog').getBoundingClientRect(); const btns=[...__top('.mat-dlg-foot').querySelectorAll('button')].map(x=>x.getBoundingClientRect().height); return JSON.stringify({l:d.left, r:d.right, w:innerWidth, sw:document.documentElement.scrollWidth, minBtn:Math.min(...btns)})})()`).then(JSON.parse);
    t('LK-M1 手機寬度：彙總視窗不超出螢幕、頁面沒有橫向捲動、按鈕高度 >= 44px', mob.l >= -1 && mob.r <= mob.w + 1 && mob.sw <= mob.w + 1 && mob.minBtn >= 44, J(mob));
    await shotTo(C, b, 'ma_mobile_summary.png');
    await b.key('Escape'); await b.sleep(300);
    await b.ev(`showDashboard(); 0`); await b.sleep(900);
    const mh = await b.ev(`JSON.stringify({sw:document.documentElement.scrollWidth, w:innerWidth, panel:!!document.querySelector('.mat-panel-tasks')})`).then(JSON.parse);
    t('LK-M2 手機寬度：首頁待辦卡不爆版', mh.panel && mh.sw <= mh.w + 1, J(mh));
    await shotTo(C, b, 'ma_mobile_home.png');
    await b.ev(`showSection('maItems'); 0`);
    await b.waitFor(`document.querySelectorAll('#maItemsView tbody tr').length>0`, 8000); await b.sleep(400);
    const mi = await b.ev(`JSON.stringify({sw:document.documentElement.scrollWidth, w:innerWidth, head:getComputedStyle(document.querySelector('#maItemsView thead')).display, tdBlock:getComputedStyle(document.querySelector('#maItemsView tbody td')).display})`).then(JSON.parse);
    t('LK-M3 手機寬度：到期管理頁不爆版，表格改成卡片式（表頭隱藏、儲存格各占一行）', mi.sw <= mi.w + 1 && mi.head === 'none' && mi.tdBlock === 'block', J(mi));
    await shotTo(C, b, 'ma_mobile_items.png');
    await b.ev(`ITTSMaTasks.wonTerm({id:'oK1',company:'外觀測試公司',product:'Svc MA 一年',stage:'C'},{force:true}); 0`);
    await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
    const mw = await b.ev(`(()=>{const d=__top('.mat-dialog').getBoundingClientRect(); return JSON.stringify({l:d.left, r:d.right, w:innerWidth, sw:document.documentElement.scrollWidth})})()`).then(JSON.parse);
    t('LK-M4 手機寬度：成交期間視窗不超出螢幕', mw.l >= -1 && mw.r <= mw.w + 1 && mw.sw <= mw.w + 1, J(mw));
    await shotTo(C, b, 'ma_mobile_won.png');
    const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
    t('LK-E 沒有 JS 例外', fe.length === 0, J(fe).slice(0, 300));
  });
}


// ═════════════════════════ fix：3020 驗收後修正（獨立檔，共用本檔小工具）═════════════════════════
function secFix(C) {
  return require('./_ma-tracking-ui-fix')({ t, browse, openAs, gotoItemsNav, countsOf, rowCos, panelRows, topText, clickBtn, focusTop, dateVals, drop, toBoard, oppOf, shotTo, mkUsers, mkOpp, mkItem, seedData, U, tp, MAP, LICP }).secFix(C);
}

// ═════════════════════════ 執行 ═════════════════════════
function exportBase() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'matui-base-'));
  const a = spawnSync('git', ['archive', BASE_REF], { cwd: ROOT, maxBuffer: 512 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + BASE_REF + ' 失敗：' + String(a.stderr));
  const x = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 512 * 1024 * 1024 });
  if (x.status !== 0) throw new Error('tar 失敗：' + String(x.stderr));
  return { dir, cleanup() { try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function run(C, only) {
  const want = (n) => !only || only === n;
  if (want('off')) await secOff(C);
  if (want('flow')) await secFlow(C);
  if (want('tasks')) await secTasks(C);
  if (want('items')) await secItems(C);
  if (want('contracts')) await secContracts(C);
  if (want('look')) await secLook(C);
  if (want('fix')) await secFix(C);
}
function report(title) {
  const bad = res.filter((x) => !x[1]);
  console.log('[' + title + '] ' + (res.length - bad.length) + '/' + res.length + ' 通過');
  return bad.length;
}

// UI 變異：用 CDP 把前端檔換成改壞的版本（不動磁碟上的檔案），對應區段必須變紅
const UI_MUTS = [
  { name: 'UM1 el() 改用 innerHTML（XSS）', file: 'ma-tasks.js', find: 'if (text != null) e.textContent = String(text);', rep: 'if (text != null) e.innerHTML = String(text);', section: 'look', expect: /LK-X/ },
  { name: 'UM2 拿掉成交期間視窗（直接送出）', file: 'app.js', find: "if (!(body && body.stage === 'Won' && M && M.enabled())) return doPut(body);", rep: 'return doPut(body);', section: 'flow', expect: /FL2 |FL3 |FL19/ },
  { name: 'UM3 拿掉秘書欄位限制（秘書也顯示合約編號欄）', file: 'ma-tasks.js', find: "function isSecretary() { return !!(state.user && state.user.role === 'secretary'); }", rep: 'function isSecretary() { return false; }', section: 'items', expect: /IT-S/ },
  { name: 'UM4 拿掉重複編號訊息', file: 'ma-tasks.js', find: "showErr((r.data && r.data.error) || '簽約確認失敗'); no.focus(); return;", rep: "showErr('簽約確認失敗'); no.focus(); return;", section: 'tasks', expect: /TK-/ },
  { name: 'UM5 拿掉開關判斷（無論後端是否啟用都建立畫面）', file: 'ma-tasks.js', find: 'state.enabled = !!(r[0] && r[0].enabled);', rep: 'state.enabled = true;', section: 'off', expect: /OFF1|OFF3|OFF4/ },
  { name: 'UM6 取消成交視窗仍繼續送出', file: 'app.js', find: 'if (extra === null) return null;', rep: 'if (extra === null) extra = {};', section: 'flow', expect: /FL5 |FL6 / },
  { name: 'UM7 .mat-overlay 層級高於其他彈窗', file: 'style.css', find: 'z-index: 100000; display: flex; align-items: center; justify-content: center; padding: 16px; background: rgba(8, 14, 30, .62); animation: mat-fade', rep: 'z-index: 100005; display: flex; align-items: center; justify-content: center; padding: 16px; background: rgba(8, 14, 30, .62); animation: mat-fade', section: 'look', expect: /LK-Z/ },
  { name: 'UM8 Service MA 頁籤錯拿 SAP_MA 合約', file: 'app.js', find: "const list = allContracts.filter(c => c.type === 'SERVICE_MA');", rep: "const list = allContracts.filter(c => c.type === 'SAP_MA');", section: 'contracts', expect: /CT-/ },
  { name: 'UM10 無任務卡的納管項仍顯示「待簽約確認」', file: 'ma-tasks.js', find: "return '';   // 沒有簽約任務卡的納管項（例如 License MA 類）：只追蹤到期，沒有待簽約動作（徽章顯示「追蹤中」）", rep: "return '待簽約確認';", section: 'fix', expect: /FX-1b/ },
  { name: 'UM11 改派候選人改回舊的全部人員清單', file: 'ma-tasks.js', find: "'/api/deal-tasks/' + encodeURIComponent(t.id) + '/reassign-candidates'", rep: "'/api/company-responsible/candidates'", section: 'tasks', expect: /TK-8b/ },
  { name: 'UM12 簽約後不刷新 Service MA 頁', file: 'ma-tasks.js', find: 'reloadServiceMa()]).then(function () { refreshBell(); });', rep: 'Promise.resolve()]).then(function () { refreshBell(); });', section: 'fix', expect: /FX-4a/ },
  { name: 'UM13 合約頁剩餘天數改回瀏覽器本地日算法', file: 'app.js', find: 'const diff = _tpDiff !== null ? _tpDiff : Math.ceil((effectiveEnd - today) / 86400000);', rep: 'const diff = Math.ceil((effectiveEnd - today) / 86400000);', section: 'fix', expect: /FX-6a|FX-6c/ },
  { name: 'UM14 拿掉再成交期間預填', file: 'ma-tasks.js', find: '} else if (!qt && hint && !(opts.prefill && opts.prefill.termFrom)) {', rep: '} else if (false) {', section: 'fix', expect: /FX-7a|FX-7c/ },
  { name: 'UM18 期間／剩餘天數欄允許斷行（窄視窗期間被擠成多行）', file: 'style.css', find: '.mat-table td[data-label="剩餘天數"] { white-space: nowrap; }', rep: '.mat-table td[data-label="剩餘天數"] { }', section: 'fix', expect: /FX-8a/ },
  { name: 'UM16 後台商品維護儲存中仍可切換頁面', file: 'admin.html', find: "if (typeof _pcSaving !== 'undefined' && _pcSaving && _pcMa && this.dataset.sec !== 'product-catalog') {", rep: 'if (false) {', section: 'fix', expect: /FX-5a/ },
  { name: 'UM17 刪除 Service MA 合約後不刷新任務／到期管理／鈴鐺', file: 'app.js', find: 'if (M && M.enabled && M.enabled() && M.afterContractChange) M.afterContractChange();', rep: '', section: 'fix', expect: /FX-2c/ },
  { name: 'UM9 鈴鐺 deal_task 通知點了不開任務卡', file: 'app.js', find: "if (item.dataset.type === 'deal_task' && window.ITTSMaTasks) {", rep: "if (false) {", section: 'tasks', expect: /TK-B/ },
];
async function groupMut() {
  console.log('[UI 變異] 用 CDP 替換前端檔為改壞的版本，對應區段的測試必須變紅');
  const H = require('./_b2-harness');
  H.guardClean();
  const exb = exportBase();
  let undetected = 0;
  try {
    const only = process.env.MATRK_UI_MUT_ONLY ? new RegExp(process.env.MATRK_UI_MUT_ONLY) : null;
    for (const m of UI_MUTS) {
      if (only && !only.test(m.name)) continue;
      let src = fs.readFileSync(path.join(ROOT, '_client', m.file), 'utf8');
      if (src.indexOf(String.fromCharCode(13)) >= 0) src = src.split(String.fromCharCode(13, 10)).join(String.fromCharCode(10));
      const find = m.find, rep = m.rep;
      if (!src.includes(find)) { console.log('  ✗ 變異錨點不存在：' + m.name); undetected++; continue; }
      res = [];
      const quiet = console.log; console.log = () => {};
      try { await run({ H, ROOT, base: exb.dir, shots: null, cdp: CDP, mutFiles: { [m.file]: src.replace(find, () => rep) } }, m.section); } catch (e) { t('(ui 例外) ' + String(e && e.message).slice(0, 100), false); }
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
module.exports = { run, mkUsers, seedData, U, tp, UI_MUTS };
if (require.main === module) {
  (async () => {
    if (process.env.MATRK_UI_MUT === '1') { process.exitCode = (await groupMut()) ? 1 : 0; return; }
    const H = require('./_b2-harness');
    H.guardClean();
    const exb = exportBase();
    res = [];
    try {
      await run({ H, ROOT, base: exb.dir, shots: process.env.MATRK_UI_SHOTS || path.join(ROOT, '..', 'ma2_shots'), cdp: CDP }, process.env.MATRK_UI_ONLY || '');
    } catch (e) { t('(ui 例外) ' + (e && e.stack || e).toString().slice(0, 400), false); }
    finally { await H.finish(); exb.cleanup(); }
    const n = report('ui');
    process.exitCode = n ? 1 : 0;
  })().catch((e) => { console.error(e && e.stack || e); process.exitCode = 1; });
}
