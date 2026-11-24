'use strict';
/**
 * 鈴鐺通知清除與導向的前端瀏覽器檢查：headless Edge＋CDP（真實滑鼠鍵盤）。由 scripts/check-notif-actions.js 的 ui 群組呼叫
 *   NTF_UI_ONLY=list|delete|nav|dismiss|deeplink|edge|look  只跑一段
 *   NTF_MUT=1 另外跑前端變異（用 CDP 把前端檔換成改壞的版本，對應區段的測試必須變紅）
 * 埠：伺服器 NTF_PORT（3450–3458）、CDP NTF_CDP（9750–9759）。全部假資料、臨時 _s_* 帳號、sandbox 的 data.json／auth.json。
 * 區段：
 *   list     結構：每則儲存型通知有 ✕（aria-label）、合約提醒列有「開啟合約」「已處理」與說明、清除已讀鈕、XSS 不執行
 *   delete   ✕ 不觸發導向、badge 更新、清除已讀、鍵盤（Tab／Enter／焦點）、超過 20 則時補上第 21 則
 *   nav      合約三種 type 導向、已刪除友善訊息、callin／lead／receivable 導向、沒有權限的提示
 *   dismiss  「已處理」：伺服器記錄、重新整理後仍不見、改迄日後重新出現
 *   deeplink #contract:<id> 深連結（含非法格式拒絕、執行中 hashchange）
 *   edge     唯讀模式的錯誤訊息、沒有通知時與修改前相同
 *   look     四種外觀對比、手機 375／320、截圖
 */
const fs = require('fs');
const path = require('path');
const { withBrowser } = require('./_fd-cdp');

const J = JSON.stringify;
let X = null;
const t = (...a) => X.t(...a);

const REC = `(()=>{ if(window.__recOn) return; window.__recOn=true; window.__reqs=[]; window.__xss=0; const of=window.fetch.bind(window); window.fetch=(u,o)=>{ try{ window.__reqs.push({m:(o&&o.method)||'GET',u:String(u)}); }catch(e){} return of(u,o); }; })()`;
const reqs = (b) => b.ev('JSON.stringify(window.__reqs||[])').then((s) => JSON.parse(s));
const reqsMatch = async (b, m, re) => (await reqs(b)).filter((r) => r.m === m && re.test(r.u));
const clearReqs = (b) => b.ev('window.__reqs=[]; 0');

const CONTRAST_FN = `(function(el, dark){
  function parse(c){ var m=c.match(/rgba?\\(([^)]+)\\)/); if(!m) return [0,0,0,0]; var p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number); return [p[0],p[1],p[2],p.length>3?p[3]:1]; }
  function over(top,bot){ var a=top[3]; return [top[0]*a+bot[0]*(1-a), top[1]*a+bot[1]*(1-a), top[2]*a+bot[2]*(1-a), 1]; }
  function lum(c){ var f=function(v){v/=255; return v<=0.03928? v/12.92 : Math.pow((v+0.055)/1.055,2.4)}; return 0.2126*f(c[0])+0.7152*f(c[1])+0.0722*f(c[2]); }
  if(!el) return null;
  var chain=[]; for(var n=el;n&&n.nodeType===1;n=n.parentElement) chain.push(n);
  var bg=dark?[11,16,32,1]:[255,255,255,1];
  for(var i=chain.length-1;i>=0;i--){ var c=parse(getComputedStyle(chain[i]).backgroundColor); if(c[3]>0) bg=over(c,bg); }
  var cs=getComputedStyle(el);
  var fg=parse(cs.color); fg=over(fg,bg);
  var L1=lum(fg), L2=lum(bg); var hi=Math.max(L1,L2), lo=Math.min(L1,L2);
  return Math.round(((hi+0.05)/(lo+0.05))*100)/100;
})`;
const contrastSel = (b, sel, dark) => b.ev(`${CONTRAST_FN}(document.querySelector(${J(sel)}), ${!!dark})`);
const THEMES = ['glass-dark', 'glass-light', 'plain-light', 'plain-dark'];
async function theme(b, mode) {
  await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass',${J(mode.startsWith('glass'))}); c.toggle('dark',${J(mode.endsWith('dark'))}); return true})()`);
  await b.sleep(1200);   // 主題切換有顏色過渡動畫，量測太早會量到過渡中途的顏色
}

function browse(C, maxMs, fn) {
  return withBrowser({ cdpPort: C.cdp, maxMs }, async (b) => {
    if (C.mutFiles) { Object.keys(C.mutFiles).forEach((k) => { b.replace['/' + k] = C.mutFiles[k]; }); await b.interceptOn(); }
    await b.send('Page.addScriptToEvaluateOnNewDocument', { source: REC });
    return fn(b);
  });
}
async function login(b, who) {
  await b.nav(X.H.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(X.H.PW)}})}).then(r=>r.status)`);
  t('UI 登入 ' + who, st === 200, st);
}
async function openUser(b, who, o) {
  o = o || {};
  await b.viewport(o.w || 1440, o.h || 900, (o.w || 1440) < 700);
  await login(b, who);
  await b.ev('localStorage.clear(); 0');
  await b.nav(X.H.BASE + '/index.html' + (o.hash || ''));
  await b.waitFor("typeof showSection==='function' && !!document.getElementById('navHome') && !!document.getElementById('notifBellBtn')", 25000);
  await b.sleep(o.wait || 2200);
}
const dropdownOpen = (b) => b.ev("(()=>{const d=document.getElementById('notifDropdown'); return !!d && d.style.display!=='none'})()");
async function openBell(b) {
  if (!(await dropdownOpen(b))) await b.click('#notifBellBtn');
  await b.waitFor("document.querySelectorAll('#notifList .notif-item, #notifList .notif-empty').length>0", 10000);
  await b.sleep(350);
}
const toast = (b) => b.ev("document.getElementById('toast').textContent");
const sectionNow = (b) => b.ev('currentSection');
const modalState = (b) => b.ev("(()=>{const m=document.getElementById('contractModal'); return JSON.stringify({open:!!m&&m.classList.contains('open'), id:document.getElementById('contractId').value, type:document.getElementById('contractTypeHidden').value, company:document.getElementById('contractCompany').value})})()").then(JSON.parse);
async function closeModal(b) { if ((await modalState(b)).open) { await b.click('#contractCancelBtn'); await b.sleep(250); } }
const exists = (b, sel) => b.ev(`!!document.querySelector(${J(sel)})`);
const count = (b, sel) => b.ev(`document.querySelectorAll(${J(sel)}).length`);
const attr = (b, sel, a) => b.ev(`(()=>{const e=document.querySelector(${J(sel)}); return e?e.getAttribute(${J(a)}):null})()`);
const badge = (b) => b.ev("(()=>{const e=document.getElementById('notifBadge'); return e && e.style.display!=='none' ? e.textContent : ''})()");
const noErr = (b) => b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
const dataNotif = (id) => X.readData().notifications.find((n) => n.id === id);
/** 開啟鈴鐺 → 點某一列的標題（避開按鈕）→ 等導向完成 */
async function clickTitle(b, rowSel) {
  await openBell(b);
  await b.click(`${rowSel} .notif-title`);
  await b.sleep(900);
}

// ═════════════════ list ═════════════════
async function secList(C) {
  await X.fresh({});
  await browse(C, 240000, async (b) => {
    await openUser(b, X.U.u1);
    await openBell(b);
    const regular = await b.ev("[...document.querySelectorAll('#notifList .notif-item[data-type]')].map(e=>e.dataset.id).join()");
    t('UL1 u1 的 12 則儲存型通知都顯示', regular.split(',').length === 12, regular);
    t('UL2 每則儲存型通知都有 ✕', (await count(b, '#notifList .notif-item[data-type] .nt-x')) === 12);
    t('UL3 ✕ 有 aria-label「清除這則通知」且是 button', (await b.ev("[...document.querySelectorAll('#notifList .nt-x')].every(e=>e.tagName==='BUTTON' && e.getAttribute('aria-label')==='清除這則通知' && e.type==='button')")) === true);
    t('UL4 動態合約提醒列：ct-erp1／ct-sap1／ct-xss／ct-del（未啟用 MA 追蹤時 SERVICE_MA 也在動態提醒）', (await b.ev("[...document.querySelectorAll('#notifList .notif-item[data-contract]')].map(e=>e.dataset.contractId).sort().join()")) === 'ct-del,ct-erp1,ct-sap1,ct-svc1,ct-xss');
    t('UL5 每列都有「開啟合約」與「已處理」且有 aria-label', (await b.ev("[...document.querySelectorAll('#notifList .notif-item[data-contract]')].every(r=>{const o=r.querySelector('.nt-open'), d=r.querySelector('.nt-done'); return o&&d&&o.textContent==='開啟合約'&&d.textContent==='已處理'&&o.getAttribute('aria-label')&&d.getAttribute('aria-label')})")) === true);
    t('UL6 有說明「續約後在合約填寫續約日期，提醒也會自動消失」', (await b.ev("document.querySelector('#notifList .nt-hint')?.textContent||''")).includes('續約後在合約填寫續約日期，提醒也會自動消失'));
    t('UL7 頁首有「清除已讀」鈕，與「全部已讀」並列', (await b.ev("(()=>{const a=document.getElementById('notifClearRead'), r=document.getElementById('notifReadAll'); return !!a&&!!r&&a.textContent==='清除已讀'&&a.parentElement===r.parentElement})()")) === true);
    t('UL8 頁首不爆版（按鈕都在下拉寬度內）', (await b.ev("(()=>{const h=document.querySelector('.notif-header'); return h.scrollWidth<=h.clientWidth+1})()")) === true);
    // XSS
    t('UL9 XSS：通知標題／內文與合約公司名含 <img onerror>／<script> 不執行', (await b.ev('window.__xss')) === 0);
    t('UL9b XSS：鈴鐺清單裡沒有被注入的 img／script 元素', (await count(b, '#notifList img, #notifList script')) === 0);
    t('UL9c XSS：標題顯示為字面文字', (await b.ev("document.querySelector('.notif-item[data-id=\"n12\"] .notif-title').textContent")) === X.XSS1 + 'N12 惡意標題' && (await b.ev("document.querySelector('.notif-item[data-id=\"n12\"] .notif-body').textContent")) === X.XSS2);
    t('UL9d XSS：合約公司名顯示為字面文字', (await b.ev("document.querySelector('.notif-item[data-contract-id=\"ct-xss\"] .notif-title').textContent")).includes(X.XSS1 + '公司'));
    t('UL10 沒有 JS 例外', noErr(b).length === 0, J(noErr(b)).slice(0, 300));
  });
}

// ═════════════════ delete ═════════════════
async function secDelete(C) {
  await X.fresh({});
  await browse(C, 300000, async (b) => {
    // u2：1 則未讀儲存型（quote_approved，點列會導向報價單）、1 則已讀儲存型、1 則動態合約提醒（未讀）→ badge 2
    await openUser(b, X.U.u2);
    t('UD0 badge 起始＝2（1 未讀儲存型＋1 動態合約）', (await badge(b)) === '2', await badge(b));
    await openBell(b);
    const sec0 = await sectionNow(b);
    await clearReqs(b);
    await b.click('.nt-x[data-nt-id="u2a"]');
    await b.waitFor("!document.querySelector('.nt-x[data-nt-id=\"u2a\"]')", 8000);
    await b.sleep(500);
    const rq = await reqs(b);
    t('UD1 點 ✕ 只送出一個 DELETE /api/notifications/u2a', rq.filter((r) => r.m === 'DELETE').length === 1 && rq.some((r) => r.m === 'DELETE' && /\/api\/notifications\/u2a$/.test(r.u)), J(rq.filter((r) => r.m !== 'GET')));
    t('UD2 ✕ 不觸發列的導向：沒有 PUT 已讀、區段沒變、沒開報價單', !rq.some((r) => r.m === 'PUT') && (await sectionNow(b)) === sec0 && !rq.some((r) => /quotations/.test(r.u) && r.m === 'GET'), J(rq.slice(0, 8)));
    t('UD3 該列從畫面移除、下拉仍開著', !(await exists(b, '.notif-item[data-id="u2a"]')) && (await dropdownOpen(b)));
    t('UD4 badge 從 2 變 1', (await badge(b)) === '1', await badge(b));
    t('UD5 伺服器端真的刪了、別人的與自己的另一則還在', !dataNotif('u2a') && !!dataNotif('u2b') && !!dataNotif('n1'));
    // 清除已讀
    await clearReqs(b);
    await b.click('#notifClearRead');
    await b.waitFor("!document.querySelector('.notif-item[data-id=\"u2b\"]')", 8000);
    await b.sleep(400);
    t('UD6 「清除已讀」送出 DELETE /api/notifications/read', (await reqsMatch(b, 'DELETE', /\/api\/notifications\/read$/)).length === 1);
    t('UD7 已讀的那則消失、動態合約提醒仍在', !(await exists(b, '.notif-item[data-id="u2b"]')) && (await exists(b, '.notif-item[data-contract]')));
    t('UD8 顯示「已清除 1 則已讀通知」', (await toast(b)).includes('已清除 1 則已讀通知'), await toast(b));
    t('UD9 伺服器端 u2b 已刪、n 開頭（別人的已讀）不受影響', !dataNotif('u2b') && !!dataNotif('m1a'));
    await b.click('#notifClearRead');
    await b.sleep(600);
    t('UD10 再按一次：沒有可清除的 → 友善提示', (await toast(b)).includes('沒有已讀的通知可清除'), await toast(b));
    t('UD10b 沒有 JS 例外', noErr(b).length === 0, J(noErr(b)).slice(0, 300));
  });
  // 鍵盤
  await X.fresh({});
  await browse(C, 240000, async (b) => {
    await openUser(b, X.U.u1);
    await openBell(b);
    t('UD11 ✕ 可用鍵盤聚焦（Tab 序列內）', await b.focus('.nt-x[data-nt-id="n2"]'));
    await b.key('Tab');
    const ae = await b.ev("document.activeElement && document.activeElement.dataset ? (document.activeElement.dataset.ntId||document.activeElement.className) : ''");
    t('UD11b Tab 後焦點移到下一則的 ✕', ae === 'n3', ae);
    const ol = await b.ev("(()=>{const s=getComputedStyle(document.activeElement); return JSON.stringify({st:s.outlineStyle, w:parseFloat(s.outlineWidth)})})()").then(JSON.parse);
    t('UD11c 鍵盤聚焦時有可見的焦點框（outline ≥2px）', ol.st !== 'none' && ol.w >= 2, J(ol));
    await b.key('Enter');
    await b.waitFor("!document.querySelector('.notif-item[data-id=\"n3\"]')", 8000);
    await b.sleep(500);
    t('UD12 在 ✕ 上按 Enter 即清除（n3）', !(await exists(b, '.notif-item[data-id="n3"]')) && !dataNotif('n3'));
    t('UD12b 清除後焦點沒有掉到 body（交給清單）', (await b.ev("document.activeElement && document.activeElement !== document.body")) === true);
    t('UD12c 沒有觸發導向（還在儀表板）', (await sectionNow(b)) === null);
  });
  // 超過 20 則
  await X.fresh({});
  await browse(C, 240000, async (b) => {
    await openUser(b, X.U.u4);
    await openBell(b);
    t('UD13 25 則時只顯示 20 則（既有行為）', (await count(b, '.notif-item[data-type] .nt-x')) === 20 && (await b.ev("[...document.querySelectorAll('.nt-x')].pop().dataset.ntId")) === 'b20');
    await b.click('.nt-x[data-nt-id="b01"]');
    await b.waitFor("!document.querySelector('.nt-x[data-nt-id=\"b01\"]')", 8000);
    await b.sleep(700);
    t('UD14 清除一則後第 21 則補上來（仍顯示 20 則，最後一則是 b21）', (await count(b, '.notif-item[data-type] .nt-x')) === 20 && (await b.ev("[...document.querySelectorAll('.nt-x')].pop().dataset.ntId")) === 'b21');
    t('UD14b 伺服器端剩 24 則', X.readData().notifications.filter((n) => n.to === X.U.u4).length === 24);
  });
}

// ═════════════════ nav ═════════════════
async function secNav(C) {
  await X.fresh({});
  await browse(C, 360000, async (b) => {
    await openUser(b, X.U.u1);
    const openRow = async (sel) => { await clickTitle(b, sel); };
    const expectModal = async (name, section, id, type, company) => {
      const ms = await modalState(b);
      t(`UN ${name}：切到「${section}」`, (await sectionNow(b)) === section, await sectionNow(b));
      t(`UN ${name}：開啟合約視窗（${id}，${type}）`, ms.open && ms.id === id && ms.type === type && (!company || ms.company === company), J(ms));
      t(`UN ${name}：下拉已關閉`, !(await dropdownOpen(b)));
      await closeModal(b);
    };
    // 儲存型 contract_expiring_30 → ERP_MA
    await openRow('.notif-item[data-id="n4"]');
    await expectModal('儲存型 contract_expiring_30（ERP_MA）', 'erp-ma', 'ct-erp1', 'ERP_MA', 'ERP 測試客戶');
    t('UN 儲存型通知點擊後仍在清單（導向不等於刪除）、已標已讀', !!dataNotif('n4') && dataNotif('n4').read === true);
    // contract_expiring_7 → SAP_MA
    await openRow('.notif-item[data-id="n5"]');
    await expectModal('儲存型 contract_expiring_7（SAP_MA）', 'sap-ma', 'ct-sap1', 'SAP_MA', 'SAP 測試客戶');
    // SERVICE_MA：沒有 Service MA 頁，退回 erp-ma
    await openRow('.notif-item[data-id="n6"]');
    await expectModal('儲存型 contract_expiring_7（SERVICE_MA）', 'erp-ma', 'ct-svc1', 'SERVICE_MA', 'Service 測試客戶');
    // 合約已不存在
    const secBefore = await sectionNow(b);
    await openRow('.notif-item[data-id="n7"]');
    t('UN 合約已刪除：顯示友善訊息「此合約已刪除或你沒有檢視權限」', (await toast(b)).includes('此合約已刪除或你沒有檢視權限'), await toast(b));
    t('UN 合約已刪除：沒開視窗、區段沒變', !(await modalState(b)).open && (await sectionNow(b)) === secBefore);
    t('UN 合約已刪除：該通知已標已讀（伺服器端）', dataNotif('n7').read === true);
    // 動態提醒列
    await openRow('.notif-item[data-contract-id="ct-erp1"]');
    await expectModal('動態合約提醒列（ERP_MA）', 'erp-ma', 'ct-erp1', 'ERP_MA', 'ERP 測試客戶');
    await openBell(b);
    await b.click('.nt-open[data-nt-contract-id="ct-sap1"]');
    await b.sleep(900);
    await expectModal('動態「開啟合約」鈕（SAP_MA）', 'sap-ma', 'ct-sap1', 'SAP_MA', 'SAP 測試客戶');
    await openRow('.notif-item[data-contract-id="ct-del"]');
    t('UN 動態提醒指向已刪除合約：友善訊息、不開視窗', (await toast(b)).includes('此合約已刪除或你沒有檢視權限') && !(await modalState(b)).open);
    // callin / receivable / lead
    await openRow('.notif-item[data-id="n8"]');
    t('UN callin_new（Call-in 已移除）：不導向、不報錯', (await sectionNow(b)) !== 'callin');
    await openRow('.notif-item[data-id="n10"]');
    t('UN receivable_overdue → 應收帳款區段', (await sectionNow(b)) === 'receivables' && (await b.ev("getComputedStyle(document.getElementById('receivablesView')).display")) !== 'none');
    await openRow('.notif-item[data-id="n9"]');
    t('UN lead_assigned（業務沒有「潛在客戶」功能）→ 友善提示、區段不變', (await toast(b)).includes('你沒有「潛在客戶」的檢視權限') && (await sectionNow(b)) === 'receivables', await toast(b));
    await openRow('.notif-item[data-id="n11"]');
    t('UN birthday 維持現狀：不導向', (await sectionNow(b)) === 'receivables');
    t('UN 導向沒有 JS 例外', noErr(b).length === 0, J(noErr(b)).slice(0, 300));
  });
  // 行銷：lead_assigned → leads
  await X.fresh({});
  await browse(C, 240000, async (b) => {
    await openUser(b, X.U.mkt);
    await clickTitle(b, '.notif-item[data-id="mk1"]');
    t('UN lead_assigned（行銷）→ 潛在客戶區段', (await sectionNow(b)) === 'leads' && (await b.ev("getComputedStyle(document.getElementById('leadsView')).display")) !== 'none', await sectionNow(b));
  });
  // MA 關閉：SERVICE_MA 合約退回合約清單（ERP 頁）仍能開啟
  await X.fresh({});
  await browse(C, 240000, async (b) => {
    await openUser(b, X.U.u1);
    await clickTitle(b, '.notif-item[data-id="n6"]');
    const ms = await modalState(b);
    t('UN MA 關閉時 Service MA 合約仍能開（退回合約清單頁）', ms.open && ms.id === 'ct-svc1' && ms.type === 'SERVICE_MA' && (await sectionNow(b)) === 'erp-ma', J(ms) + (await sectionNow(b)));
  });
}

// ═════════════════ dismiss ═════════════════
async function secDismiss(C) {
  await X.fresh({});
  await browse(C, 300000, async (b) => {
    await openUser(b, X.U.u1);
    await openBell(b);
    const sec0 = await sectionNow(b);
    await clearReqs(b);
    await b.click('.nt-done[data-nt-contract-id="ct-sap1"]');
    await b.waitFor("!document.querySelector('.notif-item[data-contract-id=\"ct-sap1\"]')", 8000);
    await b.sleep(500);
    const rq = await reqs(b);
    t('US1 「已處理」送出 POST /api/contract-reminders/ct-sap1/dismiss（只一次）', rq.filter((r) => r.m === 'POST').length === 1 && rq.some((r) => r.m === 'POST' && /\/api\/contract-reminders\/ct-sap1\/dismiss$/.test(r.u)), J(rq.filter((r) => r.m !== 'GET')));
    t('US2 該列消失、其餘合約提醒還在', !(await exists(b, '.notif-item[data-contract-id="ct-sap1"]')) && (await exists(b, '.notif-item[data-contract-id="ct-erp1"]')));
    t('US3 按「已處理」不會同時導向／開視窗（stopPropagation）', (await sectionNow(b)) === sec0 && !(await modalState(b)).open);
    t('US4 顯示「已處理…不再提醒」訊息', (await toast(b)).includes('已處理'), await toast(b));
    const ks = Object.keys(X.readData().contractReminderDismissals || {});
    t('US5 伺服器端記錄（本人|合約|迄日）', ks.length === 1 && ks[0].startsWith(X.U.u1 + '|ct-sap1|'), ks.join());
    // 重新整理後仍不見
    await b.nav(X.H.BASE + '/index.html');
    await b.waitFor("typeof showSection==='function' && !!document.getElementById('notifBellBtn')", 25000);
    await b.sleep(2000);
    await openBell(b);
    t('US6 重新整理頁面後仍不顯示（伺服器過濾，不是 localStorage）', !(await exists(b, '.notif-item[data-contract-id="ct-sap1"]')) && (await exists(b, '.notif-item[data-contract-id="ct-erp1"]')));
    // 處理 ct-erp1，再改迄日 → 重新出現
    await b.click('.nt-done[data-nt-contract-id="ct-erp1"]');
    await b.waitFor("!document.querySelector('.notif-item[data-contract-id=\"ct-erp1\"]')", 8000);
    const newEnd = X.dstr(33);
    const st = await b.ev(`fetch('/api/contracts/ct-erp1',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({endDate:${J(newEnd)}})}).then(r=>r.status)`);
    t('US7a 前置：把 ct-erp1 的迄日改成新日期', st === 200, st);
    await b.nav(X.H.BASE + '/index.html');
    await b.waitFor("typeof showSection==='function' && !!document.getElementById('notifBellBtn')", 25000);
    await b.sleep(2000);
    await openBell(b);
    t('US7 改了迄日後該合約的提醒重新出現', await exists(b, '.notif-item[data-contract-id="ct-erp1"]'));
    // 點列＝已讀（localStorage）舊行為保留
    await b.click('.notif-item[data-contract-id="ct-xss"] .notif-title');
    await b.sleep(900);
    await closeModal(b);
    await openBell(b);
    t('US8 點列仍是「已讀」（unread 樣式移除），並沒有被「已處理」', (await b.ev("!document.querySelector('.notif-item[data-contract-id=\"ct-xss\"]').classList.contains('unread')")) === true && (await exists(b, '.notif-item[data-contract-id="ct-xss"]')));
    t('US9 沒有 JS 例外', noErr(b).length === 0, J(noErr(b)).slice(0, 300));
  });
}

// ═════════════════ deeplink ═════════════════
async function secDeeplink(C) {
  await X.fresh({});
  await browse(C, 360000, async (b) => {
    await openUser(b, X.U.u1, { hash: '#contract:ct-erp1', wait: 3000 });
    let ms = await modalState(b);
    t('UK1 #contract:<id> → 開啟該合約（ERP 頁）', ms.open && ms.id === 'ct-erp1' && (await sectionNow(b)) === 'erp-ma', J(ms));
    t('UK2 處理後 hash 立刻清掉（重新整理不會再開）', (await b.ev('location.hash')) === '');
    await closeModal(b);
    // 基準：沒有 hash 時載入會送幾次 GET /api/contracts
    await b.nav(X.H.BASE + '/index.html');
    await b.waitFor("typeof showSection==='function' && !!document.getElementById('notifBellBtn')", 25000);
    await b.sleep(2200);
    const baseN = (await reqsMatch(b, 'GET', /\/api\/contracts$/)).length;
    // 非法格式：不處理、不清 hash、不送合約請求
    const bad = ['#contract:' + 'a'.repeat(65), '#contract:ab/cd', '#contract:' + X.XSS1, '#contract:', '#contract:a b', '#contracts', '#contract:ct-erp1:cost'];
    for (const h of bad) {
      await b.nav(X.H.BASE + '/index.html' + h);
      await b.waitFor("typeof showSection==='function' && !!document.getElementById('notifBellBtn')", 25000);
      await b.sleep(2200);
      const m2 = await modalState(b);
      const n2 = (await reqsMatch(b, 'GET', /\/api\/contracts$/)).length;
      t(`UK3 非法深連結不處理：${h.slice(0, 24)}`, !m2.open && n2 === baseN && (await b.ev('window.__xss')) === 0, J(m2) + ' contractsGET=' + n2 + '/' + baseN);
    }
    // 執行中 hashchange
    await b.nav(X.H.BASE + '/index.html');
    await b.waitFor("typeof showSection==='function' && !!document.getElementById('notifBellBtn')", 25000);
    await b.sleep(2000);
    await b.ev("location.hash='#contract:ct-sap1'; 0");
    await b.waitFor("document.getElementById('contractModal').classList.contains('open')", 8000);
    ms = await modalState(b);
    t('UK4 頁面開著時改 hash（推播點擊）→ 開啟合約（SAP 頁）並清掉 hash', ms.open && ms.id === 'ct-sap1' && (await sectionNow(b)) === 'sap-ma' && (await b.ev('location.hash')) === '', J(ms));
    await closeModal(b);
    await b.ev("location.hash='#contract:ct-gone'; 0");
    await b.sleep(1200);
    t('UK5 格式正確但合約不存在 → 友善訊息、不開視窗', (await toast(b)).includes('此合約已刪除或你沒有檢視權限') && !(await modalState(b)).open);
    await b.ev("location.hash='#contract:ct-svc1'; 0");
    await b.waitFor("document.getElementById('contractModal').classList.contains('open')", 8000);
    t('UK6 Service MA 合約 → erp-ma 頁（無 Service MA 頁）', (await sectionNow(b)) === 'erp-ma' && (await modalState(b)).type === 'SERVICE_MA');
    t('UK7 沒有 JS 例外', noErr(b).length === 0, J(noErr(b)).slice(0, 300));
  });
}

// ═════════════════ edge ═════════════════
async function secEdge(C) {
  await X.fresh({});
  await browse(C, 240000, async (b) => {
    await openUser(b, X.U.vw);
    await openBell(b);
    await b.click('.nt-x[data-nt-id="vw1"]');
    await b.sleep(900);
    t('UE1 唯讀模式帳號按 ✕：顯示伺服器的唯讀訊息、該列仍在', (await toast(b)).includes('唯讀') && (await exists(b, '.notif-item[data-id="vw1"]')) && !!dataNotif('vw1'), await toast(b));
  });
  await browse(C, 240000, async (b) => {
    await openUser(b, X.U.sec);
    await openBell(b);
    t('UE2 沒有通知時與修改前相同：「暫無通知」空狀態（沒有多出任何清單元素）', (await b.ev("document.getElementById('notifList').innerHTML")) === '<div class="notif-empty">暫無通知</div>');
    t('UE3 沒有通知時 badge 不顯示', (await badge(b)) === '');
  });
}

// ═════════════════ look ═════════════════
async function secLook(C) {
  await X.fresh({});
  await browse(C, 420000, async (b) => {
    await openUser(b, X.U.u1);
    await openBell(b);
    // 未啟用 MA 追蹤時動態合約提醒多一筆，儲存型通知列被擠到捲動區下方：量測前先把 n1 捲進可視範圍
    await b.ev("(()=>{const e=document.querySelector('.notif-item[data-id=\"n1\"]'); if(e) e.scrollIntoView({block:'center'}); return true})()");
    const sels = ['.nt-x', '.nt-btn.nt-open', '.nt-btn.nt-done', '.nt-hint', '#notifClearRead', '#notifReadAll', '.notif-item[data-id="n1"] .notif-title', '.notif-item[data-id="n1"] .notif-body'];
    const worst = [];
    for (const mode of THEMES) {
      await theme(b, mode);
      for (const sel of sels) {
        const c = await contrastSel(b, sel, mode.endsWith('dark'));
        if (c === null) { worst.push(mode + ' ' + sel + '=找不到'); continue; }
        if (c < 4.5) worst.push(mode + ' ' + sel + '=' + c);
      }
      if (C.shots) { fs.mkdirSync(C.shots, { recursive: true }); await b.shot(path.join(C.shots, 'bell_desktop_' + mode + '.png')); }
    }
    t('UV1 四種外觀（玻璃深／玻璃淺／一般淺／一般深）：✕、開啟合約、已處理、說明、清除已讀、全部已讀 對比 ≥ 4.5:1', worst.length === 0, worst.join('；'));
    // 未讀列（藍底）上的 ✕
    await theme(b, 'plain-light');
    const cu = await contrastSel(b, '.notif-item.unread .nt-x', false);
    t('UV2 未讀列（藍底）上的 ✕ 對比 ≥ 4.5:1', cu !== null && cu >= 4.5, cu);
    await theme(b, 'plain-dark');
    const cd = await contrastSel(b, '.notif-item.unread .nt-x', true);
    t('UV2b 深色未讀列上的 ✕ 對比 ≥ 4.5:1', cd !== null && cd >= 4.5, cd);
    await theme(b, 'plain-light');
    t('UV3 無 inline style／on* 屬性（新增元素）', (await b.ev("[...document.querySelectorAll('#notifList .nt-x, #notifList .nt-btn, #notifList .nt-actions, #notifList .nt-hint, #notifClearRead')].every(e=>!e.hasAttribute('style') && ![...e.attributes].some(a=>/^on/i.test(a.name)))")) === true);
  });
  // 手機
  for (const [w, mode] of [[375, 'plain-light'], [375, 'glass-dark'], [320, 'plain-light'], [320, 'glass-light']]) {
    await browse(C, 240000, async (b) => {
      await openUser(b, X.U.u1, { w, h: 812 });
      await theme(b, mode);
      await openBell(b);
      const m = await b.ev(`(()=>{const d=document.getElementById('notifDropdown').getBoundingClientRect(); const xs=[...document.querySelectorAll('#notifList .nt-x')]; const bt=[...document.querySelectorAll('#notifList .nt-btn')]; const hd=document.querySelector('.notif-header'); return JSON.stringify({l:Math.round(d.left), r:Math.round(d.right), iw:innerWidth, sw:document.documentElement.scrollWidth, xh:Math.min(...xs.map(e=>e.getBoundingClientRect().height)), xw:Math.min(...xs.map(e=>e.getBoundingClientRect().width)), bh:Math.min(...bt.map(e=>e.getBoundingClientRect().height)), hdOver:hd.scrollWidth>hd.clientWidth+1, listOver:document.getElementById('notifList').scrollWidth>document.getElementById('notifList').clientWidth+1, xr:Math.max(...xs.map(e=>e.getBoundingClientRect().right))})})()`).then(JSON.parse);
      t(`UM ${w}px ${mode}：下拉完全在畫面內、整頁無橫向捲動`, m.l >= 0 && m.r <= m.iw && m.sw <= m.iw + 1, J(m));
      t(`UM ${w}px ${mode}：✕ 與按鈕夠大（觸控 ≥36px）、頁首與清單不爆版`, m.xh >= 36 && m.xw >= 36 && m.bh >= 34 && !m.hdOver && !m.listOver && m.xr <= m.r, J(m));
      if (C.shots) { fs.mkdirSync(C.shots, { recursive: true }); await b.shot(path.join(C.shots, `bell_mobile_${w}_${mode}.png`)); }
    });
  }
}

// ═════════════════ 執行 ═════════════════
const SECTIONS = { list: secList, delete: secDelete, nav: secNav, dismiss: secDismiss, deeplink: secDeeplink, edge: secEdge, look: secLook };
async function run(C, only) {
  X = C.X;
  for (const k of Object.keys(SECTIONS)) {
    if (only && only !== k) continue;
    try { await SECTIONS[k](C); } catch (e) { t('(例外) UI ' + k + '：' + String(e && e.stack || e).slice(0, 300), false); }
  }
}

// ═════════════════ 前端變異 ═════════════════
const UI_MUTS = () => [
  { name: 'UM1 ✕ 沒有 stopPropagation（會連帶觸發列的導向／已讀）', file: 'app.js', section: 'delete', expect: /^UD2/, find: "      ev.stopPropagation();\n      _notifDeleteOne(btn.dataset.ntId, btn.closest('.notif-item'));", rep: "      _notifDeleteOne(btn.dataset.ntId, btn.closest('.notif-item'));" },
  { name: 'UM2 儲存型通知不導向（拿掉 _notifGoByType 呼叫）', file: 'app.js', section: 'nav', expect: /^UN 儲存型/, find: "        _notifGoByType(item.dataset.type, item.dataset.ref);", rep: "" },
  { name: 'UM3 儲存型通知標題／內文不轉義（XSS）', file: 'app.js', section: 'list', expect: /^UL9/, find: "          <div class=\"notif-title\">${escapeHtml(n.title)}</div>\n          <div class=\"notif-body\">${escapeHtml(n.body)}</div>\n          <div class=\"notif-time\">", rep: "          <div class=\"notif-title\">${n.title}</div>\n          <div class=\"notif-body\">${n.body}</div>\n          <div class=\"notif-time\">" },
  { name: 'UM4 動態合約列點擊不導向', file: 'app.js', section: 'nav', expect: /^UN 動態合約提醒列/, find: "        if (item.dataset.contractId) openContractFromBell(item.dataset.contractId, item.dataset.contractType);", rep: "" },
  { name: 'UM5 深連結格式驗證被拿掉', file: 'app.js', section: 'deeplink', expect: /^UK3/, find: "/^#contract:([A-Za-z0-9_-]{1,64})$/.exec(h)", rep: "/^#contract:(.*)$/.exec(h)" },
  { name: 'UM6 ✕ 沒有 aria-label', file: 'app.js', section: 'list', expect: /^UL3/, find: "aria-label=\"清除這則通知\" title=\"清除這則通知\">✕", rep: "title=\"清除這則通知\">✕" },
  { name: 'UM7 ✕ 顏色太淡（對比不足）', file: 'style.css', section: 'look', expect: /^UV/, find: "background: transparent; color: #4a5260; font-size: 15px;", rep: "background: transparent; color: #c9ced6; font-size: 15px;" },
  { name: 'UM8 手機版下拉寬度規則被拿掉（320px 爆版）', file: 'style.css', section: 'look', expect: /^UM 320px/, find: "  .notif-dropdown { position: absolute; left: 8px; right: 8px; top: calc(100% + 6px); width: auto; max-width: none; }\n", rep: "" },
  { name: 'UM9 合約已刪除沒有友善訊息', file: 'app.js', section: 'nav', expect: /^UN 合約已刪除/, find: "  if (!c) { showToast('此合約已刪除或你沒有檢視權限'); return false; }", rep: "  if (!c) { return false; }" },
  { name: 'UM10 「已處理」沒有送出 POST', file: 'app.js', section: 'dismiss', expect: /^US1/, find: "    const r = await fetch(`${API}/contract-reminders/${encodeURIComponent(contractId)}/dismiss`, { method: 'POST' });", rep: "    const r = { ok: true };" },
  { name: 'UM11 「清除已讀」鈕沒有綁定', file: 'app.js', section: 'delete', expect: /^UD6/, find: "$('notifClearRead').addEventListener('click', _notifClearRead);", rep: "" },
  { name: 'UM13 沒有權限也硬切區段（lead_assigned）', file: 'app.js', section: 'nav', expect: /^UN lead_assigned（業務/, find: "  if (!_ntCan(t.feature)) { showToast(`你沒有「${t.label}」的檢視權限`); return true; }", rep: "" },
  { name: 'UM14 手機 ✕ 太小', file: 'style.css', section: 'look', expect: /^UM 375px.*夠大/, find: "  .nt-x { min-width: 40px; min-height: 40px; }", rep: "  .nt-x { min-width: 20px; min-height: 20px; }" },
  { name: 'UM15 清除後不補第 21 則（不重新取清單）', file: 'app.js', section: 'delete', expect: /^UD14/, find: "  if (itemEl && itemEl.parentNode) itemEl.remove();\n  _pollBundleEtag = null;   // server 狀態已改，強制取最新（badge、第 21 則以後的通知補上來）\n  await pollNotifications();", rep: "  if (itemEl && itemEl.parentNode) itemEl.remove();" },
  { name: 'UM16 hashchange 不監聽（推播點擊不會開合約）', file: 'app.js', section: 'deeplink', expect: /^UK4/, find: "window.addEventListener('hashchange', _handleContractDeepLink);", rep: "" },
  { name: 'UM17 ✕ 沒有鍵盤焦點框', file: 'style.css', section: 'delete', expect: /^UD11c/, find: ".nt-x:focus-visible, .nt-btn:focus-visible, .notif-read-all:focus-visible, #notifList:focus-visible { outline: 2px solid #1a56b8; outline-offset: 2px; }", rep: ".nt-x:focus-visible, .nt-btn:focus-visible { outline: none; }" },
];
async function groupMut(C) {
  X = C.X;
  console.log('[變異] 前端：用 CDP 替換前端檔為改壞的版本，對應區段的測試必須變紅');
  let undetected = 0, total = 0;
  const only = process.env.NTF_MUT_ONLY ? process.env.NTF_MUT_ONLY.split(',') : null;
  for (const m of UI_MUTS()) {
    if (only && !only.some((p) => m.name.startsWith(p))) continue;
    total++;
    let src = fs.readFileSync(path.join(X.ROOT, '_client', m.file), 'utf8');
    if (!src.includes(m.find)) { console.log('  ✗ 變異錨點不存在：' + m.name); undetected++; continue; }
    src = src.replace(m.find, () => m.rep);
    X.res = []; X.setQuiet(true);
    try { await SECTIONS[m.section]({ X, shots: null, cdp: C.cdp, mutFiles: { [m.file]: src } }); } catch (e) { t('(例外) ' + String(e && e.message).slice(0, 100), false); }
    X.setQuiet(false);
    const failed = X.res.filter((x) => !x[1]);
    const hit = failed.some((x) => m.expect.test(x[0]));
    console.log((hit ? '  ✓ ' : '  ✗ ') + m.name + '：' + failed.length + ' 項變紅' + (hit ? '' : '（預期的測試沒變紅）') + (failed.length ? '｜例：' + failed[0][0].slice(0, 60) : ''));
    if (!hit) undetected++;
  }
  console.log(`前端變異測試：${total - undetected}/${total} 被偵測到`);
  return undetected;
}
module.exports = { run, groupMut, SECTIONS };
