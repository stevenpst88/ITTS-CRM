'use strict';
/**
 * check-opp-deal-id.js 的前端（headless Edge + CDP）與變異測試部分。不單獨執行。
 * 真實滑鼠／鍵盤事件；不連外網（註冊連結只檢查屬性，並在捕獲階段 preventDefault，不會真的開新分頁）。
 */
const fs = require('fs');
const path = require('path');
const { withBrowser } = require('./_fd-cdp');

const J = JSON.stringify;
const DISPLAY = (id) => `(()=>{const e=document.getElementById(${J(id)});return !!e && !e.hidden && getComputedStyle(e).display!=='none'})()`;
const TEXT = (id) => `(document.getElementById(${J(id)})||{}).textContent||''`;

async function loginBrowser(b, C, who) {
  await b.nav(C.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(C.PW)}})}).then(r=>r.status)`);
  C.ok(st === 200, '瀏覽器登入 ' + who);
}
async function openApp(b, C, w, h) {
  if (w) await b.viewport(w, h || 900, w < 700);
  await loginBrowser(b, C, '_s_u1');
  await b.nav(C.BASE + '/');
  const ready = await b.waitFor('typeof openOppCreate==="function" && typeof loadOpportunities==="function" && typeof DEAL_ID_RULES!=="undefined" && Object.keys(DEAL_ID_RULES).length>0', 25000);
  C.ok(ready, '使用者端載入完成且商品目錄規則已就緒（DEAL_ID_RULES）');
  await b.ev(`(async()=>{await loadOpportunities(); return allOpportunities.length})()`);
  // 記錄所有非 GET 請求（驗證「前端擋住」＝根本沒送出）
  await b.ev(`(()=>{ if(window.__reqs) return; window.__reqs=[]; const of=window.fetch; window.fetch=function(u,o){ if(o&&o.method&&o.method!=='GET') window.__reqs.push({u:String(u),m:o.method,b:typeof o.body==='string'?o.body:''}); return of.apply(this,arguments); }; window.__linkClicks=[]; document.addEventListener('click',function(e){ const a=e.target.closest&&e.target.closest('a.deal-id-link'); if(a){ window.__linkClicks.push({href:a.href,target:a.target,rel:a.rel,id:a.id}); e.preventDefault(); } },true); })()`);
}
const selSet = (b, id, v) => b.ev(`(()=>{const e=document.getElementById(${J(id)}); if(!e) return false; e.value=${J(v)}; e.dispatchEvent(new Event('change',{bubbles:true})); return e.value===${J(v)}})()`);
const valSet = (b, id, v) => b.ev(`(()=>{const e=document.getElementById(${J(id)}); e.value=${J(v)}; e.dispatchEvent(new Event('input',{bubbles:true})); return true})()`);
const reqs = (b, re) => b.ev(`window.__reqs.filter(r=>${re}.test(r.u)).map(r=>({m:r.m,b:r.b}))`);
const clearReqs = (b) => b.ev(`window.__reqs.length=0; window.__linkClicks.length=0; 0`);
async function serverOpp(C, id) { return (await C.listOpps('_s_u1')).find((o) => o.id === id); }
async function theme(b, mode) {   // mode: 'glass-dark' | 'glass-light' | 'plain-light' | 'plain-dark'
  await b.ev(`(()=>{const c=document.body.classList; const g=${J(mode.startsWith('glass'))}; const d=${J(mode.endsWith('dark'))}; c.toggle('glass',g); c.toggle('dark',d); return true})()`);
  await b.sleep(250);
}
const shot = async (b, C, name) => { if (C.SHOTS) await b.shot(path.join(C.SHOTS, name)); };

// ───── 建立（商機視窗新增模式）─────
async function uiCreate(b, C) {
  console.log('  [UI] 建立商機視窗（oppEdit）');
  const { ok, eq } = C;
  await b.ev(`openOppCreate('沙盒公司A'); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  eq(await b.ev(DISPLAY('oppEditDealIdGroup')), false, 'UI1 新增視窗剛開啟：Deal ID 欄位隱藏（還沒選商品）');
  await selSet(b, 'oppEditCategory', 'ERP');
  await selSet(b, 'oppEditProduct', 'ERP 顧問導入專案 PE');
  eq(await b.ev(DISPLAY('oppEditDealIdGroup')), false, 'UI2 選「不需要」的商品：欄位維持隱藏');
  await selSet(b, 'oppEditProduct', 'SAP Public Cloud License');
  eq(await b.ev(DISPLAY('oppEditDealIdGroup')), true, 'UI3 選「需要」的商品：欄位出現');
  eq(await b.ev(DISPLAY('oppEditDealIdStar')), true, 'UI4 必填星號顯示');
  eq(await b.ev(`document.getElementById('oppEditDealId').getAttribute('aria-required')`), 'true', 'UI5 輸入框 aria-required=true');
  const lk = await b.ev(`(()=>{const a=document.getElementById('oppEditDealIdLink'); return {shown:!a.hidden && getComputedStyle(a).display!=='none', href:a.getAttribute('href'), target:a.target, rel:a.rel, text:a.textContent.trim(), inBox: !!a.closest('.deal-id-box')}})()`);
  eq([lk.shown, lk.href, lk.target, lk.inBox], [true, C.REG_URL, '_blank', true], 'UI6 「前往註冊商機」連結顯示、指向商品設定的網址、新分頁、位於輸入框容器內');
  ok(/noopener/.test(lk.rel) && /noreferrer/.test(lk.rel), 'UI7 連結 rel 含 noopener noreferrer', lk.rel);
  ok(/前往註冊商機/.test(lk.text), 'UI8 連結文字', lk.text);
  await selSet(b, 'oppEditProduct', 'SAP CCFLEX');
  eq(await b.ev(DISPLAY('oppEditDealIdLink')), false, 'UI9 沒設註冊網址的商品：不顯示連結');
  ok(/請洽管理員設定註冊網址/.test(await b.ev(TEXT('oppEditDealIdHint'))), 'UI10 沒設網址但需要 Deal ID：提示「請洽管理員設定註冊網址」');
  await selSet(b, 'oppEditProductSub', 'C&S');
  eq(await b.ev(DISPLAY('oppEditDealIdGroup')), true, 'UI11 選了明細（SAP CCFLEX C&S）仍需要 Deal ID');
  await selSet(b, 'oppEditProduct', 'SAP Public Cloud License');
  await b.ev(`(()=>{document.getElementById('oppEditAmount').value='123'; document.getElementById('oppEditExpectedDate').value='${new Date().getFullYear()}-12-30'; return 0})()`);

  // 不填就存檔 → 前端擋住（根本沒送出）
  await clearReqs(b);
  await b.click('#oppEditSave');
  await b.sleep(400);
  eq((await reqs(b, '/\\/api\\/opportunities/')).length, 0, 'UI12 沒填 Deal ID 按儲存：前端擋下，沒有送出任何請求');
  ok((await b.ev(TEXT('oppEditDealIdErr'))).length > 0, 'UI13 欄位旁顯示錯誤訊息', await b.ev(TEXT('oppEditDealIdErr')));
  eq(await b.ev(`document.getElementById('oppEditDealId').getAttribute('aria-invalid')`), 'true', 'UI14 aria-invalid=true');
  eq(await b.ev(`document.activeElement && document.activeElement.id`), 'oppEditDealId', 'UI15 焦點移到 Deal ID 欄位');
  eq(await b.ev(`document.getElementById('oppEditOverlay').classList.contains('open')`), true, 'UI16 視窗維持開啟');
  await shot(b, C, 'dealid_create_error_glassdark_1440.png');

  // 點連結：不真的開新分頁（捕獲階段 preventDefault），只檢查它被點到時的屬性
  await b.click('#oppEditDealIdLink');
  const lc = await b.ev(`window.__linkClicks`);
  eq([lc.length, lc[0] && lc[0].href, lc[0] && lc[0].target], [1, C.REG_URL, '_blank'], 'UI17 真實點擊連結：目標網址與新分頁屬性正確');
  ok(/noreferrer/.test((lc[0] || {}).rel || ''), 'UI18 點擊當下 rel 含 noreferrer');
  eq((await reqs(b, '/\\/api\\/opportunities/')).length, 0, 'UI19 點連結不會送出表單');

  // 貼上 65 字 → 即時錯誤、仍擋
  await valSet(b, 'oppEditDealId', 'z'.repeat(65));
  ok(/64/.test(await b.ev(TEXT('oppEditDealIdErr'))), 'UI20 貼上 65 字：即時顯示長度錯誤（不靜默截斷）', await b.ev(TEXT('oppEditDealIdErr')));
  await b.click('#oppEditSave'); await b.sleep(300);
  eq((await reqs(b, '/\\/api\\/opportunities/')).length, 0, 'UI21 65 字存檔被擋');
  await valSet(b, 'oppEditDealId', 'exist-1');
  ok(/已用於「沙盒公司A／SAP Public Cloud License」/.test(await b.ev(TEXT('oppEditDealIdHint'))), 'UI21b 軟性提醒：Deal ID（不分大小寫）已用於其他商機 → 提示但不擋', await b.ev(TEXT('oppEditDealIdHint')));
  eq(await b.ev(`dealIdValidate('oppEdit')`), true, 'UI21c 重複只是提醒，驗證仍通過（不硬擋）');
  await valSet(b, 'oppEditDealId', '');
  ok(!/已用於/.test(await b.ev(TEXT('oppEditDealIdHint'))), 'UI21d 改掉後重複提醒消失');

  // 真實鍵盤輸入（含前後空白）→ 錯誤消失 → 存檔
  await b.click('#oppEditDealId');
  await b.type('  DL-UI-1  ');
  eq(await b.ev(TEXT('oppEditDealIdErr')), '', 'UI22 開始輸入後錯誤訊息清除');
  await shot(b, C, 'dealid_create_filled_glassdark_1440.png');
  await b.click('#oppEditSave');
  await b.waitFor(`!document.getElementById('oppEditOverlay').classList.contains('open')`, 8000);
  const posted = await reqs(b, '/\\/api\\/opportunities$/');
  const pb = posted.find((x) => x.m === 'POST');
  eq(pb && JSON.parse(pb.b).dealId, 'DL-UI-1', 'UI23 送出的 dealId 已去前後空白');
  const made = (await C.listOpps('_s_u1')).find((o) => o.dealId === 'DL-UI-1');
  ok(!!made, 'UI24 伺服器已建立該商機並保存 dealId');

  // 再開一次：選不需要的商品 → 隱藏，且不送 dealId
  await b.ev(`openOppCreate('沙盒公司A'); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  eq(await b.ev(`document.getElementById('oppEditDealId').value`), '', 'UI25 重新開啟新增視窗：Deal ID 欄位已清空');
  await selSet(b, 'oppEditCategory', 'ERP');
  await selSet(b, 'oppEditProduct', 'SAP Public Cloud License');
  await valSet(b, 'oppEditDealId', 'TEMP-1');
  await selSet(b, 'oppEditProduct', 'ERP 顧問導入專案 PE');
  eq(await b.ev(DISPLAY('oppEditDealIdGroup')), false, 'UI26 改選不需要的商品：欄位隱藏');
  { const dv = await b.ev(`dealIdValue('oppEdit')`); ok(dv === undefined || dv === null, 'UI27 隱藏時不送值（dealIdValue 為 undefined）', dv); }
  await b.ev(`closeOppEdit(); 0`);
}

// ───── 編輯既有商機 ─────
async function uiEdit(b, C) {
  console.log('  [UI] 編輯既有商機');
  const { ok, eq } = C;
  await b.ev(`(async()=>{await loadOpportunities(); return 1})()`);
  await clearReqs(b);
  await b.ev(`openOppEdit('o-legacy'); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  eq(await b.ev(DISPLAY('oppEditDealIdGroup')), true, 'UE1 舊商機（商品需要、沒有 Deal ID）：開啟編輯就看到欄位');
  ok(/補填/.test(await b.ev(TEXT('oppEditDealIdErr'))), 'UE2 立即提示「請補填」', await b.ev(TEXT('oppEditDealIdErr')));
  await shot(b, C, 'dealid_edit_legacy_glassdark_1440.png');
  await b.click('#oppEditSave'); await b.sleep(400);
  eq((await reqs(b, '/\\/api\\/opportunities\\//')).filter((x) => x.m === 'PUT').length, 0, 'UE3 沒補填就儲存：前端擋下，沒有送 PUT');
  eq(await b.ev(`document.getElementById('oppEditOverlay').classList.contains('open')`), true, 'UE4 視窗維持開啟');
  await b.click('#oppEditDealId');
  await b.type('FIX-UI-2');
  await b.click('#oppEditSave');
  await b.waitFor(`!document.getElementById('oppEditOverlay').classList.contains('open')`, 8000);
  const put = (await reqs(b, '/\\/api\\/opportunities\\/o-legacy$/')).find((x) => x.m === 'PUT');
  eq(put && JSON.parse(put.b).dealId, 'FIX-UI-2', 'UE5 補填後送出 dealId');
  eq((await serverOpp(C, 'o-legacy') || {}).dealId, 'FIX-UI-2', 'UE6 伺服器保存');

  await clearReqs(b);
  await b.ev(`(async()=>{await loadOpportunities(); return 1})()`);
  await b.ev(`openOppEdit('o-free'); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  eq(await b.ev(DISPLAY('oppEditDealIdGroup')), false, 'UE7 不需要且沒有 Deal ID 的商機：欄位隱藏');
  await b.click('#oppEditSave');
  await b.waitFor(`!document.getElementById('oppEditOverlay').classList.contains('open')`, 8000);
  const put2 = (await reqs(b, '/\\/api\\/opportunities\\/o-free$/')).find((x) => x.m === 'PUT');
  ok(put2 && !('dealId' in JSON.parse(put2.b)), 'UE8 隱藏時儲存：PUT 不帶 dealId 鍵（不動既有資料）', put2 ? put2.b : '');

  await b.ev(`openOppEdit('o-has'); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  eq(await b.ev(`document.getElementById('oppEditDealId').value`), 'EXIST-1', 'UE9 已有 Deal ID：欄位帶出既有值');
  await selSet(b, 'oppEditProduct', 'ERP 顧問導入專案 PE');
  eq([await b.ev(DISPLAY('oppEditDealIdGroup')), await b.ev(DISPLAY('oppEditDealIdStar')), await b.ev(`document.getElementById('oppEditDealId').getAttribute('aria-required')`)], [true, false, 'false'], 'UE10 改成不需要的商品：因為已有值仍顯示，但改為選填（無星號、aria-required=false）');
  await b.ev(`closeOppEdit(); 0`);

  // 看板拖曳：只送 {stage}，不被 Deal ID 擋
  await b.ev(`(async()=>{await loadOpportunities(); renderKanban(); return 1})()`);
  await clearReqs(b);
  const dr = await b.ev(`(()=>{const card=document.querySelector('.kanban-card[data-id="o-legacy2"]'); if(!card) return 'nocard'; card.dispatchEvent(new DragEvent('dragstart',{bubbles:true})); const col=document.getElementById('kanbanCardsB'); col.dispatchEvent(new DragEvent('drop',{bubbles:true,cancelable:true})); return 'ok'})()`);
  eq(dr, 'ok', 'UE11 找得到看板卡片並觸發拖曳');
  await b.waitFor(`window.__reqs.some(r=>/o-legacy2$/.test(r.u))`, 5000);
  await b.sleep(500);
  const dput = (await reqs(b, '/\\/api\\/opportunities\\/o-legacy2$/')).find((x) => x.m === 'PUT');
  eq(dput && JSON.parse(dput.b), { stage: 'B' }, 'UE12 看板拖曳送出的 PUT 只有 {stage}');
  eq((await serverOpp(C, 'o-legacy2') || {}).stage, 'B', 'UE13 伺服器階段已更新（沒有 Deal ID 也不被擋）');
  await clearReqs(b);
  await b.ev(`updateOppStage('o-legacy2','A'); 0`); await b.sleep(800);
  eq((await serverOpp(C, 'o-legacy2') || {}).stage, 'A', 'UE14 清單下拉的階段快速變更（updateOppStage）同樣不被擋');
}

// ───── 顯示（看板卡片／清單／XSS）─────
async function uiDisplay(b, C) {
  console.log('  [UI] 顯示與 XSS');
  const { ok, eq } = C;
  await b.ev(`(async()=>{await loadOpportunities(); renderKanban(); renderOppTable(); return 1})()`);
  eq(await b.ev(`(document.querySelector('.kanban-card[data-id="o-has"] .kanban-card-dealid')||{}).textContent`), 'Deal ID：EXIST-1', 'UD1 看板卡片顯示 Deal ID');
  eq(await b.ev(`!!document.querySelector('.kanban-card[data-id="o-free"] .kanban-card-dealid')`), false, 'UD2 沒有 Deal ID 的卡片不多一行');
  const xc = await b.ev(`(()=>{const c=document.querySelector('.kanban-card[data-id="o-xss"]'); if(!c) return null; const d=c.querySelector('.kanban-card-dealid'); return {text:d&&d.textContent, imgs:c.querySelectorAll('img').length, xss: window.__xss===1}})()`);
  eq(xc && xc.text, 'Deal ID：' + C.XSS, 'UD3 XSS 酬載在看板卡片以純文字顯示（逐字相同）');
  eq(xc && [xc.imgs, xc.xss], [0, false], 'UD4 沒有建出 <img>、onerror 沒執行');
  const tr = await b.ev(`(()=>{const r=document.querySelector('#oppTableBody tr[data-id="o-xss"]'); if(!r) return null; const d=r.querySelector('.opp-dealid-line'); return {text:d&&d.textContent, imgs:r.querySelectorAll('img').length, xss: window.__xss===1}})()`);
  ok(!!tr, 'UD5 清單（表格）有該商機列');
  eq(tr && tr.text, 'Deal ID：' + C.XSS, 'UD6 XSS 酬載在清單以純文字顯示');
  eq(tr && [tr.imgs, tr.xss], [0, false], 'UD7 清單沒有建出 <img>、onerror 沒執行');
  await b.ev(`openOppEdit('o-xss'); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  eq(await b.ev(`document.getElementById('oppEditDealId').value`), C.XSS, 'UD8 編輯視窗：輸入框 value 逐字帶出酬載');
  eq(await b.ev(`document.querySelectorAll('#oppEditOverlay img').length`), 0, 'UD9 編輯視窗沒有建出 <img>');
  eq(await b.ev(`window.__xss===1`), false, 'UD10 全程 onerror 沒執行');
  await b.ev(`closeOppEdit(); 0`);
  // 被竄改成 javascript: 的註冊網址：前端不渲染成連結
  await b.ev(`(()=>{DEAL_ID_RULES['SAP Public Cloud License'].url='javascript:alert(1)'; openOppCreate('沙盒公司A'); const c=document.getElementById('oppEditCategory'); c.value='ERP'; c.dispatchEvent(new Event('change',{bubbles:true})); const p=document.getElementById('oppEditProduct'); p.value='SAP Public Cloud License'; p.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`);
  eq(await b.ev(DISPLAY('oppEditDealIdLink')), false, 'UD11 網址若是 javascript:（資料被竄改）→ 前端不顯示連結');
  eq(await b.ev(`document.getElementById('oppEditDealIdLink').getAttribute('href')`), null, 'UD12 …且沒有 href');
  await b.ev(`DEAL_ID_RULES['SAP Public Cloud License'].url=${J(C.REG_URL)}; closeOppEdit(); 0`);
}

// ───── 其他兩個表單（名片視窗商機頁籤、拜訪視窗商機頁籤）─────
async function uiOthers(b, C) {
  console.log('  [UI] 名片／拜訪視窗的商機頁籤');
  const { ok, eq } = C;
  for (const [k, cat, prod] of [['cOpp', 'cOppCategory', 'cOppProduct'], ['opp', 'oppCategory', 'oppProduct']]) {
    await selSet(b, cat, 'ERP');
    await selSet(b, prod, 'ERP 顧問導入專案 PE');
    eq(await b.ev(DISPLAY(k + 'DealIdGroup')), false, `UO1 ${k}：不需要的商品 → 隱藏`);
    await selSet(b, prod, 'SAP Public Cloud License');
    eq(await b.ev(DISPLAY(k + 'DealIdGroup')), true, `UO2 ${k}：需要的商品 → 出現`);
    eq(await b.ev(`document.getElementById('${k}DealIdLink').getAttribute('href')`), C.REG_URL, `UO3 ${k}：連結指向設定的網址`);
    eq(await b.ev(`dealIdValidate('${k}')`), false, `UO4 ${k}：沒填 → 驗證不通過`);
    ok((await b.ev(TEXT(k + 'DealIdErr'))).length > 0, `UO5 ${k}：欄位旁有錯誤訊息`);
    await valSet(b, k + 'DealId', ' OK-' + k + ' ');
    eq(await b.ev(`dealIdValidate('${k}')`), true, `UO6 ${k}：填了 → 通過`);
    eq(await b.ev(`dealIdValue('${k}')`), 'OK-' + k, `UO7 ${k}：取值已去前後空白`);
    await b.ev(`dealIdSet('${k}', ''); 0`);
  }
  // 拜訪視窗：選了商機類別但商機要 Deal ID 沒填 → 拜訪也不會被儲存（先擋）
  await b.ev(`(()=>{ openVisitModal(); return 0 })()`);
  await clearReqs(b);
  await selSet(b, 'oppCategory', 'ERP'); await selSet(b, 'oppProduct', 'SAP Public Cloud License');
  await b.ev(`(()=>{ const sel=document.getElementById('visitContactId'); document.getElementById('visitTopic').value='測試主題'; document.getElementById('oppExpectedDate').value='${new Date().getFullYear()}-12-30'; if(sel && !sel.value){ const o=document.createElement('option'); o.value='c-test'; o.textContent='x'; sel.appendChild(o); sel.value='c-test'; } return 0 })()`);
  await b.click('#visitSaveBtn'); await b.sleep(500);
  eq((await reqs(b, '/\\/api\\/(visits|opportunities)/')).length, 0, 'UO8 拜訪視窗：商機缺 Deal ID → 拜訪與商機都沒送出（不會留下半套資料）');
  await b.ev(`closeVisitModal(); 0`);
}

// ───── 版面／主題／截圖 ─────
async function uiShots(b, C) {
  console.log('  [UI] 主題與手機版面（對比檢查＋截圖）');
  const { ok, eq } = C;
  await b.ev(`openOppCreate('沙盒公司A'); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  await selSet(b, 'oppEditCategory', 'ERP'); await selSet(b, 'oppEditProduct', 'SAP Public Cloud License');
  await b.ev(`document.getElementById('oppEditDealId').focus(); dealIdValidate('oppEdit'); 0`);
  const CONTRAST = `(function(){
    const parse=c=>{const m=/rgba?\\(([^)]+)\\)/.exec(c);if(!m)return null;const p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number);return {r:p[0],g:p[1],b:p[2],a:p[3]===undefined?1:p[3]};};
    const over=(t,b)=>({r:t.r*t.a+b.r*(1-t.a),g:t.g*t.a+b.g*(1-t.a),b:t.b*t.a+b.b*(1-t.a),a:1});
    const lum=c=>{const f=v=>{v/=255;return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)};return 0.2126*f(c.r)+0.7152*f(c.g)+0.0722*f(c.b)};
    const ratio=(a,b)=>{const l1=lum(a),l2=lum(b);return (Math.max(l1,l2)+0.05)/(Math.min(l1,l2)+0.05)};
    const bgOf=el=>{const L=[];let e=el;while(e&&e.nodeType===1){const c=parse(getComputedStyle(e).backgroundColor);if(c&&c.a>0){L.push(c);if(c.a>=0.999)break;}e=e.parentElement;}
      let cur=L.length&&L[L.length-1].a>=0.999?L.pop():{r:255,g:255,b:255,a:1};while(L.length)cur=over(L.pop(),cur);return cur;};
    return ['#oppEditDealIdLink','#oppEditDealIdHint','#oppEditDealIdErr','#oppEditDealId','label[for="oppEditDealId"]'].map(sel=>{const el=document.querySelector(sel);if(!el)return [sel,0];const fg=parse(getComputedStyle(el).color);const bg=bgOf(el);const f=fg.a<1?over(fg,bg):fg;return [sel,+ratio(f,bg).toFixed(2)];});
  })()`;
  for (const m of ['glass-dark', 'glass-light', 'plain-light', 'plain-dark']) {
    await theme(b, m);
    const rs = await b.ev(CONTRAST);
    const low = rs.filter(([s, r]) => r < 4.5);
    ok(low.length === 0, `US 對比 ${m}：連結／提示／錯誤／輸入框／標籤文字對比 ≥ 4.5`, J(rs));
    console.log('    對比（' + m + '）：' + rs.map(([s, r]) => s.replace('#oppEditDealId', '').replace('label[for="oppEditDealId"]', 'label') + '=' + r).join(' '));
    await shot(b, C, `dealid_create_${m}_1440.png`);
  }
  await theme(b, 'glass-dark');
  await b.viewport(375, 812, true);
  await b.sleep(400);
  await shot(b, C, 'dealid_create_glassdark_375.png');
  const mob = await b.ev(`(()=>{const g=document.getElementById('oppEditDealIdGroup').getBoundingClientRect(); const l=document.getElementById('oppEditDealIdLink').getBoundingClientRect(); const i=document.getElementById('oppEditDealId').getBoundingClientRect(); return {gl:g.left,gr:g.right,ll:l.left,lr:l.right,lw:l.width,iw:i.width,vw:innerWidth,sw:document.documentElement.scrollWidth}})()`);
  ok(mob.gl >= 0 && mob.gr <= mob.vw + 1 && mob.lr <= mob.vw + 1 && mob.sw <= mob.vw + 1, '375 寬：欄位與連結都在畫面內、沒有橫向捲動', J(mob));
  await theme(b, 'glass-light');
  await shot(b, C, 'dealid_create_glasslight_375.png');
  await b.viewport(1440, 1000, false);
  await b.ev(`closeOppEdit(); 0`);
}

// ───── 後台：商品目錄 ─────
async function uiAdmin(b, C) {
  console.log('  [UI] 後台商品目錄');
  const { ok, eq } = C;
  await b.viewport(1440, 1100, false);
  await loginBrowser(b, C, '_s_adm');
  await b.nav(C.BASE + '/admin.html');
  await b.waitFor(`document.querySelector('[data-sec="product-catalog"]')`, 20000);
  await b.ev(`(()=>{ window.__reqs=[]; const of=window.fetch; window.fetch=function(u,o){ if(o&&o.method&&o.method!=='GET') window.__reqs.push({u:String(u),m:o.method,b:typeof o.body==='string'?o.body:''}); return of.apply(this,arguments); }; return 0})()`);
  await b.ev(`document.querySelector('[data-sec="product-catalog"]').click(); 0`);   // 側欄群組預設收合，直接觸發導覽
  ok(await b.waitFor(`document.querySelectorAll('#pcPanel [data-pcfld="req"]').length>0`, 10000), 'UA1 商品目錄載入，每個商品有「需要 Deal ID」勾選');
  const info = await b.ev(`(()=>{const q=(gi,pi)=>document.querySelector('#pcPanel [data-pcfld="req"][data-gi="'+gi+'"][data-pi="'+pi+'"]'); return {r0:q(0,0).checked, r2:q(0,2).checked, urlRows:document.querySelectorAll('#pcPanel .pc-deal-url').length}})()`);
  eq([info.r0, info.r2], [true, false], 'UA2 既有設定帶出（Public Cloud 勾選、顧問導入專案不勾）');
  eq(info.urlRows, 2, 'UA3 只有已勾選的商品才顯示網址輸入框（2 個）');
  await b.click('#pcPanel [data-pcfld="req"][data-gi="0"][data-pi="2"]');
  await b.sleep(200);
  eq(await b.ev(`!!document.querySelector('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]')`), true, 'UA4 勾選「需要 Deal ID」後出現網址輸入框');
  ok(/尚未設定/.test(await b.ev(`document.querySelector('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]').parentNode.querySelector('.pc-deal-hint').textContent`)), 'UA5 沒填網址：提示未設定');
  await b.click('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]');
  await b.type('javascript:alert(1)');
  eq(await b.ev(`document.querySelector('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]').classList.contains('pc-deal-bad')`), true, 'UA6 壞網址：即時標示（紅框）');
  ok(/http/.test(await b.ev(`document.querySelector('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]').parentNode.querySelector('.pc-deal-hint').textContent`)), 'UA7 壞網址：即時提示「必須以 http:// 或 https:// 開頭」');
  await b.ev(`window.__reqs.length=0; 0`);
  await shot(b, C, 'dealid_admin_badurl_glassdark_1440.png');
  await b.click('#pcSaveBtn'); await b.sleep(600);
  eq((await b.ev(`window.__reqs.filter(r=>/product-catalog/.test(r.u))`)).length, 0, 'UA8 壞網址：前端擋住儲存，沒有送出');
  ok(/無法儲存/.test(await b.ev(`document.getElementById('pcStatus').textContent`)), 'UA9 顯示無法儲存的原因');
  // 改成合法網址 → 儲存
  await b.ev(`(()=>{const e=document.querySelector('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]'); e.value=''; e.dispatchEvent(new Event('input',{bubbles:true})); return 0})()`);
  await b.click('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]');
  await b.type('https://deal3.example.test/new');
  eq(await b.ev(`document.querySelector('#pcPanel .pc-deal-url[data-gi="0"][data-pi="2"]').classList.contains('pc-deal-bad')`), false, 'UA10 合法網址：紅框消失');
  await shot(b, C, 'dealid_admin_ok_glassdark_1440.png');
  await b.click('#pcSaveBtn');
  ok(await b.waitFor(`/已儲存/.test(document.getElementById('pcStatus').textContent)`, 8000), 'UA11 儲存成功');
  const cat = await b.ev(`fetch('/api/product-catalog').then(r=>r.json())`);
  const it = cat.bus.ERP[0].items[2];
  eq([it.name, it.requiresDealId, it.dealRegisterUrl], ['ERP 顧問導入專案 PE', true, 'https://deal3.example.test/new'], 'UA12 伺服器保存了新設定');
  // 取消勾選 → 不再存網址
  await b.click('#pcPanel [data-pcfld="req"][data-gi="0"][data-pi="2"]'); await b.sleep(200);
  await b.click('#pcSaveBtn');
  await b.waitFor(`/已儲存/.test(document.getElementById('pcStatus').textContent)`, 8000);
  const cat2 = await b.ev(`fetch('/api/product-catalog').then(r=>r.json())`);
  const it2 = cat2.bus.ERP[0].items[2];
  eq([it2.requiresDealId === undefined, it2.dealRegisterUrl === undefined], [true, true], 'UA13 取消勾選後存檔：requiresDealId 與網址都不再保存');
  for (const m of ['glass-light', 'plain-light', 'plain-dark']) {
    await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass', ${J(m.startsWith('glass'))}); c.toggle('dark', ${J(m.endsWith('dark'))}); return 0})()`);
    await b.sleep(250);
    await shot(b, C, `dealid_admin_${m}_1440.png`);
  }
}

async function runUiOnce(C, flags, sections, mut) {
  await C.freshServer();
  for (const w of ['_s_u1', '_s_adm']) C.eq(await C.login(w), 200, '節點端登入 ' + w);
  await withBrowser({ cdpPort: C.CDP, maxMs: 420000 }, async (b) => {
    if (flags.ui) {
      if (mut) { b.replace[mut.urlTail] = mut.mutated; await b.interceptOn(); }   // 把前端檔換成改壞的版本（不動磁碟上的檔案）
      await openApp(b, C, 1440, 1000);
      for (const s of sections.ui) {
        try { await ({ create: uiCreate, edit: uiEdit, display: uiDisplay, others: uiOthers, shots: uiShots })[s](b, C); } catch (e) { C.ok(false, `UI 區段 ${s} 例外：${e.message}`); }
        if (b.errors.length) { console.log('    頁面錯誤：', b.errors.slice(0, 3).join(' | ')); C.ok(false, '頁面沒有 JS 例外', b.errors[0]); b.errors.length = 0; }
      }
    }
    if (flags.admin && sections.admin) {
      try { await uiAdmin(b, C); } catch (e) { C.ok(false, `後台區段例外：${e.message}`); }
      if (b.errors.length) { C.ok(false, '後台沒有 JS 例外', b.errors[0]); b.errors.length = 0; }
    }
  });
  console.log('  Edge 殘留程序：', withBrowser.leftover || 0);
  C.ok((withBrowser.leftover || 0) === 0, '測試結束沒有殘留的 Edge 程序');
  await C.stopServer(); C.cleanupFiles();
}

exports.run = async function (C, flags) {
  console.log('[前端] headless Edge + CDP');
  await runUiOnce(C, flags, { ui: ['create', 'edit', 'display', 'others', 'shots'], admin: true });
};
exports.runUiOnce = runUiOnce;
exports.sections = { uiCreate, uiEdit, uiDisplay, uiOthers, uiShots, uiAdmin };

exports.mutants = (C, fns) => require('./_opp-deal-id-mut.js').run(C, fns, exports.runUiOnce);
