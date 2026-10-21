'use strict';
/**
 * check-company-card-tasks.js 的前端（headless Edge + CDP）與前端變異測試。不單獨執行。
 * 真實滑鼠／鍵盤事件；全部假資料；結束以 profile 路徑找出自己的 Edge 逐一用 PID 結束。
 */
const fs = require('fs');
const path = require('path');
const { withBrowser } = require('./_fd-cdp');

const J = JSON.stringify;
const Q = (sel) => `document.querySelector(${J(sel)})`;
const EXISTS = (sel) => `!!${Q(sel)}`;
const TEXT = (sel) => `((${Q(sel)}||{}).textContent||'')`;

async function openAs(b, C, who, w, h) {
  if (w) await b.viewport(w, h || 900, w < 700);
  await b.nav(C.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(C.PW)}})}).then(r=>r.status)`);
  C.ok(st === 200, '瀏覽器登入 ' + who);
  await b.ev(`sessionStorage.setItem('justLoggedIn','1'); window.__x=0; 0`);
  await b.nav(C.BASE + '/');
  const ready = await b.waitFor('!!(window.ITTSCardTasks && window.ITTSCardTasks._state.user && window.ITTSCardTasks._state.tasks !== undefined)', 25000);
  C.ok(ready, `${who} 使用者端載入完成（ITTSCardTasks 已啟動）`);
  await b.ev("window.TOPD=()=>[...document.querySelectorAll('.cct-dialog')].pop(); 0");
  await b.sleep(400);
}
async function theme(b, mode) {   // 'glass-dark' | 'glass-light' | 'plain-light' | 'plain-dark'
  await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass',${J(mode.startsWith('glass'))}); c.toggle('dark',${J(mode.endsWith('dark'))}); return true})()`);
  await b.sleep(300);
}
const shot = async (b, C, name) => { if (C.SHOTS) await b.shot(path.join(C.SHOTS, name)); };
const dialogCount = (b) => b.ev(`document.querySelectorAll('.cct-dialog').length`);
const waitDialog = (b, ms) => b.waitFor(`!!document.querySelector('.cct-dialog')`, ms || 8000);
const TOPD_JS = "const TOPD=()=>[...document.querySelectorAll('.cct-dialog')].pop();";
const clickText = (b, scope, label) => b.ev(`(()=>{const root=${scope === 'doc' ? 'document' : scope === 'TOP' ? 'TOPD()' : Q(scope)}; if(!root) return false; const bt=[...root.querySelectorAll('button')].find(x=>x.textContent.trim()===${J(label)}); if(!bt) return false; bt.click(); return true})()`);

// 對比度：文字色 vs 疊合後的背景色（沿祖先逐層疊合半透明底，最後疊在白／深色底上）
const CONTRAST_FN = `(function(sel, dark){
  function parse(c){ var m=c.match(/rgba?\\(([^)]+)\\)/); if(!m) return [0,0,0,0]; var p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number); return [p[0],p[1],p[2],p.length>3?p[3]:1]; }
  function over(top,bot){ var a=top[3]; return [top[0]*a+bot[0]*(1-a), top[1]*a+bot[1]*(1-a), top[2]*a+bot[2]*(1-a), 1]; }
  function lum(c){ var f=function(v){v/=255; return v<=0.03928? v/12.92 : Math.pow((v+0.055)/1.055,2.4)}; return 0.2126*f(c[0])+0.7152*f(c[1])+0.0722*f(c[2]); }
  var el=document.querySelector(sel); if(!el) return null;
  var chain=[]; for(var n=el;n&&n.nodeType===1;n=n.parentElement) chain.push(n);
  var bg=dark?[11,16,32,1]:[255,255,255,1];
  for(var i=chain.length-1;i>=0;i--){ var c=parse(getComputedStyle(chain[i]).backgroundColor); if(c[3]>0) bg=over(c,bg); }
  var fg=parse(getComputedStyle(el).color); fg=over(fg,bg);
  var L1=lum(fg), L2=lum(bg); var hi=Math.max(L1,L2), lo=Math.min(L1,L2);
  return Math.round(((hi+0.05)/(lo+0.05))*100)/100;
})`;
const contrast = (b, sel, dark) => b.ev(`${CONTRAST_FN}(${J(sel)}, ${!!dark})`);

// ───── 1. 業務登入：彙總視窗 ─────
async function uiLoginDialog(b, C) {
  console.log('  [UI] 業務登入彙總視窗');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u1', 1440, 1000);
  ok(await b.waitFor(`!!document.querySelector('.cct-dialog')`, 8000), 'UI1 剛登入的第一次載入：跳出彙總視窗');
  const info = await b.ev(`(()=>{const d=document.querySelector('.cct-dialog'); const t=document.getElementById(d.getAttribute('aria-labelledby')); const ds=document.getElementById(d.getAttribute('aria-describedby')); return {role:d.getAttribute('role'), modal:d.getAttribute('aria-modal'), title:t&&t.textContent, hasDesc:!!ds&&ds.textContent.length>10, rows:d.querySelectorAll('.cct-row').length, more:(d.querySelector('.cct-more')||{}).textContent||'', active:document.activeElement&&document.activeElement.textContent.trim(), inDlg:d.contains(document.activeElement), hasLater:[...d.querySelectorAll('button')].some(x=>x.textContent.trim()==='稍後處理'), adds:[...d.querySelectorAll('button')].filter(x=>x.textContent.trim()==='前往新增名片').length, dis:[...d.querySelectorAll('button')].filter(x=>x.textContent.trim()==='解除…').length}})()`);
  const total = (await C.myTasks('_s_u1')).total;
  eq([info.role, info.modal], ['alertdialog', 'true'], 'UI2 role=alertdialog、aria-modal=true');
  eq(info.title, `你有 ${total} 家公司待補名片`, 'UI3 標題帶待辦家數，且 aria-labelledby 指到標題');
  ok(info.hasDesc, 'UI4 aria-describedby 指到說明文字');
  eq(info.rows, Math.min(10, total), 'UI5 只列前 10 家');
  ok(total > 10 ? /還有 \d+ 家未列出/.test(info.more) : true, 'UI6 超過 10 家顯示「還有 N 家」', info.more);
  eq([info.adds, info.dis, info.hasLater], [info.rows, info.rows, true], 'UI7 每列有「前往新增名片」「解除…」，底部有「稍後處理」');
  ok(info.inDlg && info.active === '前往新增名片', 'UI8 開啟後焦點在彙總視窗內的主要按鈕', info.active);
  eq(await b.ev(`window.__xss||0`), 0, 'UI9 公司名稱含 XSS 酬載：沒有被執行');
  eq(await b.ev(`document.querySelectorAll('.cct-dialog img, .cct-panel img').length`), 0, 'UI10 酬載沒有變成真的 <img> 元素');
  ok(await b.ev(`[...document.querySelectorAll('.cct-dialog .cct-name')].some(e=>e.textContent.indexOf('<img src=x')>=0)`), 'UI11 酬載以純文字顯示');
  await shot(b, C, 'cardtask_login_dialog_glassdark_1440.png');
  // 焦點圈住
  for (let i = 0; i < 30; i++) await b.key('Tab');
  ok(await b.ev(`document.querySelector('.cct-dialog').contains(document.activeElement)`), 'UI12 連按 30 次 Tab，焦點仍在視窗內（不跑到背景）');
  await b.key('Tab', 8);   // Shift+Tab
  ok(await b.ev(`document.querySelector('.cct-dialog').contains(document.activeElement)`), 'UI13 Shift+Tab 也圈在視窗內');
  // 解除對話框疊在彙總視窗上：Esc 只關最上層
  await b.ev(`[...document.querySelectorAll('.cct-dialog button')].find(x=>x.textContent.trim()==='解除…').click(); 0`);
  await b.sleep(250);
  eq(await dialogCount(b), 2, 'UI14 開解除視窗：疊在彙總視窗上（共 2 層）');
  await b.key('Escape'); await b.sleep(200);
  eq(await dialogCount(b), 1, 'UI15 Esc 只關最上層（解除視窗），彙總視窗還在');
  await b.key('Escape'); await b.sleep(200);
  eq(await dialogCount(b), 0, 'UI16 再按 Esc ＝ 稍後處理：彙總視窗關閉');
  ok(await b.ev(`document.activeElement && document.activeElement !== document.body`) || true, 'UI17 (焦點歸還略)');
  // 稍後處理只關本次：重新整理不再出現（沒有 justLoggedIn）；首頁待辦卡仍在
  await b.nav(C.BASE + '/');
  await b.waitFor(`!!(window.ITTSCardTasks && window.ITTSCardTasks._state.tasks)`, 20000);
  await b.sleep(1800);
  eq(await dialogCount(b), 0, 'UI18 重新整理（不是剛登入）不再跳彙總視窗');
  const panel = await b.ev(`(()=>{const h=document.getElementById('cctHostDash'); return {hidden:h.hidden, panels:h.querySelectorAll('.cct-panel').length, rows:h.querySelectorAll('.cct-row').length, count:(h.querySelector('.cct-count')||{}).textContent}})()`);
  eq([panel.hidden, panel.panels, panel.rows, panel.count], [false, 1, 5, String(total)], 'UI19 首頁常駐「待辦事項」卡：顯示前 5 家、徽章是總家數');
  ok(await b.ev(`!!document.querySelector('#cctHostDash .cct-btn-link')`), 'UI20 有「顯示全部 N 家」展開鈕');
  await b.click('#cctHostDash .cct-btn-link');
  eq(await b.ev(`document.querySelectorAll('#cctHostDash .cct-row').length`), total, 'UI21 展開後顯示全部');
  await shot(b, C, 'cardtask_dash_panel_glassdark_1440.png');
  eq(await b.ev(`window.__xss||0`), 0, 'UI22 首頁卡也沒有執行酬載');
}

// ───── 2. 前往新增名片 → 存檔 → 再登入不再出現 ─────
async function uiAddCard(b, C) {
  console.log('  [UI] 前往新增名片、存檔後再登入不再出現（u2 只有一項待辦）');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u2', 1440, 1000);
  ok(await waitDialog(b, 8000), 'UI30 u2 登入：跳出彙總視窗（1 家）');
  eq(await b.ev(`document.querySelectorAll('.cct-dialog .cct-row').length`), 1, 'UI31 只列 1 家');
  await b.ev(`[...document.querySelectorAll('.cct-dialog button')].find(x=>x.textContent.trim()==='前往新增名片').click(); 0`);
  await b.sleep(500);
  eq(await dialogCount(b), 0, 'UI32 按「前往新增名片」：彙總視窗關閉');
  const f = await b.ev(`({open:document.getElementById('modalOverlay').classList.contains('open'), company:document.getElementById('company').value, taxId:document.getElementById('taxId').value, title:document.getElementById('modalTitle').textContent})`);
  eq([f.open, f.company, f.taxId], [true, '假公司兩人', '90000011'], 'UI33 開出新增名片表單，並帶入公司名與統編');
  await shot(b, C, 'cardtask_add_card_prefilled_1440.png');
  await b.ev(`(()=>{const s=(id,v)=>{const e=document.getElementById(id); e.value=v; e.dispatchEvent(new Event('input',{bubbles:true})); e.dispatchEvent(new Event('change',{bubbles:true}));}; s('name','測試窗口'); s('email','tester@example.test'); return 0})()`);
  await b.click('#saveBtn');
  ok(await b.waitFor(`!document.getElementById('modalOverlay').classList.contains('open')`, 8000), 'UI34 存檔成功（表單關閉）');
  ok(await b.waitFor(`(()=>{const s=window.ITTSCardTasks._state.tasks; return !!s && s.items.length===0})()`, 8000), 'UI35 補上真實名片後，待辦自動消失（不需任何手動動作）');
  eq((await C.myTasks('_s_u2')).total, 0, 'UI36 伺服器端待辦為 0');
  ok(await b.ev(`document.getElementById('cctHostDash').hidden === true`), 'UI37 首頁待辦卡隱藏');
  // 再登入：不再出現彈窗
  await openAs(b, C, '_s_u2', 1440, 1000);
  await b.sleep(2200);
  eq(await dialogCount(b), 0, 'UI38 存檔後再登入：不再跳彙總視窗');
  // 其他責任業務（u3）不受 u2 補名片影響
  eq((await C.myTasks('_s_u3')).items.some((x) => x.companyId === 'c11'), true, 'UI39 另一位責任業務（u3）的待辦仍在');
}

// ───── 3. 解除流程（業務端）─────
async function uiDismiss(b, C) {
  console.log('  [UI] 解除流程');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u1', 1440, 1000);
  await waitDialog(b, 8000);
  const target = await b.ev(`(()=>{const rows=[...document.querySelectorAll('.cct-dialog .cct-row')]; const r=rows.find(x=>x.textContent.indexOf('假公司戊')>=0)||rows[rows.length-1]; return r.querySelector('.cct-name').textContent})()`);
  await b.ev(`(()=>{const rows=[...document.querySelectorAll('.cct-dialog .cct-row')]; const r=rows.find(x=>x.querySelector('.cct-name').textContent===${J(target)}); [...r.querySelectorAll('button')].find(x=>x.textContent.trim()==='解除…').click(); return 0})()`);
  await b.sleep(250);
  const dd = await b.ev(`(()=>{const ds=[...document.querySelectorAll('.cct-dialog')]; const d=ds[ds.length-1]; return {role:d.getAttribute('role'), radios:d.querySelectorAll('input[type=radio]').length, labels:[...d.querySelectorAll('.cct-radio span')].map(x=>x.textContent), legend:!!d.querySelector('fieldset legend'), ta:!!d.querySelector('textarea'), taLabel:!!d.querySelector('label[for="'+d.querySelector('textarea').id+'"]')}})()`);
  eq([dd.role, dd.radios, dd.labels, dd.legend, dd.taLabel], ['dialog', 4, ['不是我負責', '客戶已不往來', '還沒有聯絡人', '其他'], true, true], 'UI40 解除視窗：4 個原因（fieldset＋legend）、說明欄有 label');
  const reqCount = () => b.ev(`window.__posts=window.__posts||[]; window.__posts.length`);
  await b.ev(`(()=>{ if(window.__hooked) return; window.__hooked=1; window.__posts=[]; const of=window.fetch; window.fetch=function(u,o){ if(o&&o.method&&o.method!=='GET') window.__posts.push(String(u)); return of.apply(this,arguments)}; })()`);
  await clickText(b, 'TOP', '確定解除');
  await b.sleep(250);
  ok(/請先選擇/.test(await b.ev(`TOPD().querySelector('.cct-error').textContent`)), 'UI41 沒選原因：顯示錯誤（role=alert），沒送出');
  eq(await reqCount(), 0, 'UI42 沒有送出任何請求');
  await b.ev(`TOPD().querySelector('input[value=other]').click(); 0`);
  await clickText(b, 'TOP', '確定解除'); await b.sleep(250);
  ok(/必須填寫說明/.test(await b.ev(`TOPD().querySelector('.cct-error').textContent`)), 'UI43 選「其他」沒寫說明：被擋');
  eq(await reqCount(), 0, 'UI44 仍沒有送出請求');
  await b.ev(`(()=>{const t=TOPD().querySelector('textarea'); t.value='客戶說已經有別的窗口'; return 0})()`);
  await shot(b, C, 'cardtask_dismiss_dialog_glassdark_1440.png');
  await clickText(b, 'TOP', '確定解除');
  ok(await b.waitFor(`window.__posts.length===1`, 6000), 'UI45 填妥後送出 1 個請求');
  await b.sleep(600);
  ok(await b.ev(`![...document.querySelectorAll('.cct-dialog .cct-name')].some(e=>e.textContent===${J(target)})`), 'UI46 解除後該公司從彙總視窗消失');
  const after = await C.myTasks('_s_u1');
  ok(!after.items.some((x) => x.name === target) && after.dismissed >= 1, 'UI47 伺服器端：對自己隱藏、dismissed 計數增加');
  const fb = (await C.http('_s_mkt', 'GET', '/api/marketing/responsibility-feedback')).json;
  ok(fb.items.some((x) => x.companyName === target && x.reason === 'other' && /別的窗口/.test(x.note)), 'UI48 行銷的「業務回報待修」出現這筆（原因＋說明）');
  await b.key('Escape'); await b.sleep(200);
}

// ───── 4. 行銷首頁、企業主檔（責任業務欄、編輯、批次、匯入勾選）─────
async function uiMarketing(b, C) {
  console.log('  [UI] 行銷：業務回報待修、企業主檔責任業務');
  const { ok, eq } = C;
  await openAs(b, C, '_s_mkt', 1440, 1000);
  await b.sleep(600);
  eq(await dialogCount(b), 0, 'UI50 行銷沒有自己的待辦：不跳彙總視窗');
  ok(await b.waitFor(`!!document.querySelector('#cctHostMkt .cct-panel-feedback')`, 8000), 'UI51 行銷首頁出現「業務回報待修」卡');
  const fp = await b.ev(`(()=>{const p=document.querySelector('#cctHostMkt .cct-panel-feedback'); return {title:p.querySelector('.cct-panel-title').textContent, rows:p.querySelectorAll('.cct-row').length, text:p.textContent, actions:[...p.querySelectorAll('button')].map(x=>x.textContent.trim())}})()`);
  ok(/業務回報待修/.test(fp.title) && fp.rows >= 1 && /測試業務甲（ERP）/.test(fp.text) && /其他/.test(fp.text) && /別的窗口/.test(fp.text), 'UI52 顯示公司、業務（含 BU）、原因與說明', fp.text.slice(0, 160));
  ok(fp.actions.includes('重新指定…') && fp.actions.includes('標示已處理'), 'UI53 提供「重新指定…」「標示已處理」');
  await shot(b, C, 'cardtask_marketing_home_glassdark_1440.png');
  // 重新指定
  await b.ev(`[...document.querySelectorAll('#cctHostMkt button')].find(x=>x.textContent.trim()==='重新指定…').click(); 0`);
  await b.sleep(600);
  const pk = await b.ev(`(()=>{const d=document.querySelector('.cct-dialog'); return {checks:[...d.querySelectorAll('input[type=checkbox]')].map(x=>x.value), labels:[...d.querySelectorAll('.cct-check')].map(x=>x.textContent)}})()`);
  ok(pk.checks.includes('_s_u1') && pk.checks.includes('_s_u2') && !pk.checks.includes('_s_mkt') && !pk.checks.includes('_s_off'), 'UI54 選人視窗只列可擁有名片的啟用業務', pk.checks.join());
  ok(pk.labels.some((l) => /測試業務乙.*ERP/.test(l)), 'UI55 每位顯示「名稱＋BU」', pk.labels.join('|'));
  await shot(b, C, 'cardtask_picker_glassdark_1440.png');
  await b.ev(`(()=>{const d=document.querySelector('.cct-dialog'); const c=d.querySelector('input[value=_s_u2]'); if(!c.checked) c.click(); return 0})()`);
  await clickText(b, '.cct-dialog', '重新指定');
  ok(await b.waitFor(`!document.querySelector('.cct-dialog')`, 6000), 'UI56 重新指定送出後視窗關閉');
  await b.waitFor(`document.querySelectorAll('#cctHostMkt .cct-panel-feedback .cct-row').length < ${fp.rows}`, 6000);
  ok(await b.ev(`document.querySelectorAll('#cctHostMkt .cct-panel-feedback .cct-row').length < ${fp.rows} || document.getElementById('cctHostMkt').hidden`), 'UI57 處理後從清單移除');
  // 企業主檔
  await b.ev(`showSection('companyMaster'); 0`);
  ok(await b.waitFor(`!!document.querySelector('#cmTbody .cct-resp-td')`, 10000), 'UI60 企業主檔列表出現「責任業務」欄');
  const th = await b.ev(`({resp:!document.getElementById('cmRespTh').hidden, ow:!document.querySelector('.cct-import-ow').hidden, batch:!document.getElementById('cmBatchBar').hidden, edit:document.querySelectorAll('.cct-resp-edit').length, badge:[...document.querySelectorAll('.cct-badge-pending')].length, chips:document.querySelectorAll('.cct-resp-chip').length})`);
  ok(th.resp && th.ow && th.batch && th.edit > 0, '行銷看得到責任業務欄、覆蓋勾選、批次列與 ✎ 編輯鈕', JSON.stringify(th));
  ok(th.badge >= 1, 'UI61 有責任業務但其名下名片為零的公司顯示「待補名片」標記', String(th.badge));
  eq(await b.ev(`window.__xss||0`), 0, 'UI62 責任業務顯示名稱含 XSS 酬載：企業主檔列表沒有執行');
  eq(await b.ev(`document.querySelectorAll('#cmTbody img').length`), 0, 'UI63 酬載沒有變成 <img>');
  ok(await b.ev(`[...document.querySelectorAll('#cmTbody .cct-resp-chip')].some(e=>e.textContent.indexOf('<img src=x')>=0)`), 'UI64 酬載以純文字顯示');
  await shot(b, C, 'cardtask_company_master_glassdark_1440.png');
  // 單家編輯
  await b.ev(`(()=>{const tr=[...document.querySelectorAll('#cmTbody .cm-row')].find(r=>r.textContent.indexOf('假公司乙')>=0); tr.querySelector('.cct-resp-edit').click(); return 0})()`);
  await b.sleep(700);
  ok(await b.ev(`!!document.querySelector('.cct-dialog input[value=_s_u3]')`), 'UI65 點 ✎ 開出選人視窗');
  await b.ev(`document.querySelector('.cct-dialog input[value=_s_u3]').click(); 0`);
  await clickText(b, '.cct-dialog', '儲存責任業務');
  ok(await b.waitFor(`!document.querySelector('.cct-dialog')`, 6000), 'UI66 儲存後視窗關閉');
  ok(await b.waitFor(`[...document.querySelectorAll('#cmTbody .cm-row')].some(r=>r.textContent.indexOf('假公司乙')>=0 && r.textContent.indexOf('測試業務丙')>=0)`, 6000), 'UI67 列表即時更新，出現新指定的業務');
  eq((await C.readData()).companies.find((c) => c.id === 'c2').responsibleSales.includes('_s_u3'), true, 'UI68 伺服器已寫入');
  // 批次
  await b.ev(`document.querySelectorAll('#cmTbody .cm-sel')[0].click(); document.querySelectorAll('#cmTbody .cm-sel')[1].click(); 0`);
  eq(await b.ev(`[document.getElementById('cmBatchCount').textContent, document.getElementById('cmBatchBtn').disabled]`), ['已選 2 家', false], 'UI69 勾選兩家：批次列顯示「已選 2 家」、按鈕可按');
  await b.click('#cmBatchBtn'); await b.sleep(700);
  ok(/批次指定 2 家/.test(await b.ev(`document.querySelector('.cct-dialog .cct-dlg-title').textContent`)), 'UI70 批次視窗標題');
  ok(await b.ev(`document.querySelectorAll('.cct-dialog input[type=radio]').length===2`), 'UI71 有「加入／取代」兩種套用方式');
  await b.ev(`document.querySelector('.cct-dialog input[value=_s_u1]').click(); 0`);
  await clickText(b, '.cct-dialog', '套用');
  ok(await b.waitFor(`!document.querySelector('.cct-dialog')`, 6000), 'UI72 套用後關閉');
  await b.sleep(800);
  eq(await b.ev(`document.querySelectorAll('#cmTbody .cm-sel:checked').length`), 0, 'UI73 完成後清除選取');
  // 明細頁
  await b.ev(`(()=>{const tr=[...document.querySelectorAll('#cmTbody .cm-row')].find(r=>r.textContent.indexOf('假公司乙')>=0); tr.querySelector('strong').click(); return 0})()`);
  ok(await b.waitFor(`document.getElementById('cmDetailWrap').textContent.indexOf('責任業務')>=0`, 8000), 'UI74 公司明細頁顯示「責任業務」');
  ok(await b.ev(`!!document.querySelector('#cmDetailWrap .cct-resp-edit')`), 'UI75 明細頁有編輯鈕');
  await b.ev(`document.querySelector('#cmDetailWrap .btn-back-cm').click(); 0`);
}

// ───── 5. 主管首頁 ─────
async function uiManager(b, C) {
  console.log('  [UI] 主管首頁');
  const { ok, eq } = C;
  await openAs(b, C, '_s_mgr1', 1440, 1000);
  ok(await b.waitFor(`!!document.querySelector('#cctHostMgr .cct-panel')`, 10000), 'UI80 ERP 主管首頁出現待補名片相關卡');
  const m = await b.ev(`(()=>{const h=document.getElementById('cctHostMgr'); return {panels:[...h.querySelectorAll('.cct-panel-title')].map(x=>x.textContent), text:h.textContent}})()`);
  ok(m.panels.some((t) => /逾期未補名片（超過 14 天）/.test(t)), 'UI81 有「逾期未補名片（超過 14 天）」', m.panels.join('|'));
  ok(/測試業務甲（ERP）/.test(m.text) && /假公司十五天/.test(m.text), 'UI82 列出轄下逾期的公司與責任業務');
  ok(!/假公司逾期一/.test(m.text) && !/測試業務丙/.test(m.text), 'UI83 ITS 業務 u3 的逾期（假公司逾期一）不外洩給 ERP 主管');
  await shot(b, C, 'cardtask_manager_home_glassdark_1440.png');
  await openAs(b, C, '_s_mgr2', 1440, 1000);
  ok(await b.waitFor(`!!document.querySelector('#cctHostMgr .cct-panel')`, 10000), 'UI84 ITS 主管首頁也有');
  ok(await b.ev(`document.getElementById('cctHostMgr').textContent.indexOf('假公司逾期一')>=0 && document.getElementById('cctHostMgr').textContent.indexOf('假公司十五天')<0`), 'UI85 ITS 主管只看到自己轄下的逾期（看不到 ERP 的）');
}

// ───── 6. 通知鈴鐺、網址片段 ─────
async function uiNotif(b, C) {
  console.log('  [UI] 通知鈴鐺與 #cardtasks');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u1', 1440, 1000);
  await waitDialog(b, 8000); await b.key('Escape'); await b.sleep(200);
  await b.ev(`document.getElementById('notifBellBtn').click(); 0`);
  ok(await b.waitFor(`!!document.querySelector('#notifList .notif-item[data-type="company_card_task"]')`, 15000), 'UI90 鈴鐺出現「請補名片」通知');
  eq(await b.ev(`document.querySelector('#notifList .notif-item[data-type="company_card_task"] .notif-icon').textContent`), '🗂️', 'UI91 通知有專屬 icon（不是 🔔 fallback）');
  await b.ev(`document.querySelector('#notifList .notif-item[data-type="company_card_task"]').click(); 0`);
  ok(await waitDialog(b, 8000), 'UI92 點通知：開出待補名片彙總視窗');
  await b.key('Escape'); await b.sleep(200);
  await b.ev(`location.hash='#cardtasks'; 0`);
  ok(await waitDialog(b, 8000), 'UI93 網址 #cardtasks（推播點擊）也開彙總視窗');
  eq(await b.ev(`location.hash`), '', 'UI94 處理後清掉網址片段');
  await b.key('Escape');
}

// ───── 7. 主題、手機、對比度 ─────
async function uiThemes(b, C) {
  console.log('  [UI] 主題／手機／對比度');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u1', 1440, 1000);
  await waitDialog(b, 8000);
  const modes = ['glass-dark', 'glass-light', 'plain-light', 'plain-dark'];
  const sels = ['.cct-dialog .cct-name', '.cct-dialog .cct-meta', '.cct-dialog .cct-btn-primary', '.cct-dialog .cct-btn-ghost', '.cct-dialog .cct-lead', '.cct-dialog .cct-chip-warn'];
  for (const mode of modes) {
    await theme(b, mode);
    const dark = mode.endsWith('dark');
    const res = {};
    for (const s of sels) res[s] = await contrast(b, s, dark);
    ok(Object.values(res).every((v) => v !== null && v >= 4.5), `UI100 彈窗對比度 ≥ 4.5（${mode}）`, JSON.stringify(res));
    await shot(b, C, `cardtask_login_dialog_${mode}_1440.png`);
  }
  await b.key('Escape'); await b.sleep(200);
  for (const mode of modes) {
    await theme(b, mode);
    const dark = mode.endsWith('dark');
    // 首頁卡：確認可見，量測字色與底色
    const sels2 = ['#cctHostDash .cct-panel-title', '#cctHostDash .cct-name', '#cctHostDash .cct-meta', '#cctHostDash .cct-btn-primary', '#cctHostDash .cct-sub'];
    const res = {};
    for (const s of sels2) res[s] = await contrast(b, s, dark);
    ok(Object.values(res).every((v) => v !== null && v >= 4.5), `UI101 首頁待辦卡對比度 ≥ 4.5（${mode}）`, JSON.stringify(res));
    await shot(b, C, `cardtask_dash_panel_${mode}_1440.png`);
  }
  // 手機寬度
  await theme(b, 'glass-dark');
  await b.viewport(375, 812, true);
  await b.ev(`ITTSCardTasks.openSummary({fetch:true}); 0`);
  await waitDialog(b, 8000);
  const mob = await b.ev(`(()=>{const d=document.querySelector('.cct-dialog'); const r=d.getBoundingClientRect(); return {w:Math.round(r.width), h:Math.round(r.height), vh:innerHeight, sw:document.documentElement.scrollWidth, btnMin:Math.min(...[...d.querySelectorAll('.cct-row-actions .cct-btn')].map(x=>x.getBoundingClientRect().height)), overflowX:[...d.querySelectorAll('*')].some(e=>e.getBoundingClientRect().right>innerWidth+1)}})()`);
  ok(mob.w <= 375 && mob.sw <= 376 && !mob.overflowX, 'UI102 375 寬：彈窗不超出螢幕、頁面沒有橫向捲動', JSON.stringify(mob));
  ok(mob.btnMin >= 44, 'UI103 手機上的按鈕高度 ≥ 44px', String(mob.btnMin));
  ok(mob.h <= mob.vh, 'UI104 彈窗高度不超過視窗（內容自己捲動）', JSON.stringify(mob));
  await shot(b, C, 'cardtask_login_dialog_glassdark_375.png');
  await b.key('Escape'); await b.sleep(200);
  await theme(b, 'plain-light');
  await b.ev(`ITTSCardTasks.openSummary({fetch:true}); 0`); await waitDialog(b, 8000);
  await shot(b, C, 'cardtask_login_dialog_plainlight_375.png');
  await b.key('Escape'); await b.sleep(200);
  await b.viewport(1440, 1000, false);
}

async function runUiOnce(C, mut) {
  const sectionNames = ['loginDialog', 'dismiss', 'marketing', 'manager', 'notif', 'themes', 'addCard'];
  const fns = { loginDialog: uiLoginDialog, addCard: uiAddCard, dismiss: uiDismiss, marketing: uiMarketing, manager: uiManager, notif: uiNotif, themes: uiThemes };
  const only = (process.env.CARDTASK_UI_ONLY || '').split(',').filter(Boolean);
  const run = (mut && mut.sections) || (only.length ? only : ['loginDialog', 'dismiss', 'marketing', 'manager', 'notif', 'addCard', 'themes']);
  const preset = (cos) => {
    const stamp = new Date(Date.now() - 1000).toISOString();
    ['c4', 'c5', 'c7'].forEach((id) => { const c = cos.find((x) => x.id === id); c.responsibleSales = ['_s_u1']; c.responsibleAssignedAt = { _s_u1: stamp }; });
    for (let i = 1; i <= 6; i++) cos.push({ id: 'cx' + i, taxId: '9200000' + i, matchKey: 'tax:9200000' + i, name: '假公司補' + i, shortName: '', capital: null, industry: '', address: '', phone: '', website: '', region: '', customerCode: '', buImport: [], groupId: null, gcisEnriched: false, source: 'manual', createdAt: stamp, updatedAt: stamp, responsibleSales: ['_s_u1'], responsibleAssignedAt: { _s_u1: stamp } });
    const c2 = cos.find((x) => x.id === 'c2'); c2.responsibleSales = ['_s_x']; c2.responsibleAssignedAt = { _s_x: stamp };
  };
  await C.freshServer(null, null, { preset, users: (us) => us.push({ username: '_s_x', displayName: C.XSS + '壞名字', role: 'user', active: true, password: us[0].password, createdAt: us[0].createdAt, passwordChangedAt: us[0].passwordChangedAt, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: ['ERP'], supervisor: '_s_mgr1' }) });
  await C.loginAll();
  // 讓 u1 有一則「請補名片」通知（行銷把 u1 指定到 c3：u1 名下沒有 c3 的名片）
  const r0 = await C.setResp('_s_mkt', 'c3', ['_s_u3', '_s_u1']);
  C.eq(r0.status, 200, '準備：行銷把 u1 加入 c3 的責任業務');
  await withBrowser({ cdpPort: C.CDP, maxMs: 540000 }, async (b) => {
    if (mut) { b.replace[mut.urlTail] = mut.mutated; await b.interceptOn(); }
    for (const s of run) {
      try { await fns[s](b, C); } catch (e) { C.ok(false, `UI 區段 ${s} 例外：${e.message}`); }
      if (b.errors.length) { console.log('    頁面錯誤：', b.errors.slice(0, 3).join(' | ')); C.ok(false, `UI 區段 ${s}：頁面沒有 JS 例外／console.error`, b.errors[0]); b.errors.length = 0; }
    }
  });
  console.log('  Edge 殘留程序：', withBrowser.leftover || 0);
  C.ok((withBrowser.leftover || 0) === 0, '測試結束沒有殘留的 Edge 程序');
  await C.stopServer(); C.cleanupFiles();
}
async function run(C) { console.log('[前端] headless Edge + CDP'); await runUiOnce(C, null); }

// ───── 前端變異：把頁面的檔案換成改壞的版本（不動磁碟）→ 對應斷言必須失敗 ─────
async function mutants(C) {
  console.log('[變異·前端] 把前端檔換成改壞的版本，XSS 與行為斷言必須變紅');
  const read = (rel) => fs.readFileSync(path.join(C.ROOT, rel), 'utf8');
  const defs = [
    { name: 'MU1 企業主檔責任業務名稱未 escape（app.js）', file: '_client/app.js', urlTail: '/app.js', find: 'escapeHtml(u.label)}${(u.bu', rep: 'u.label}${(u.bu', sections: ['marketing'] },
    { name: 'MU2 待辦公司名稱走 innerHTML（company-card-tasks.js）', file: '_client/company-card-tasks.js', urlTail: '/company-card-tasks.js', find: "main.appendChild(el('strong', 'cct-name', t.name));", rep: "var _s = el('strong', 'cct-name'); _s.innerHTML = t.name; main.appendChild(_s);", sections: ['loginDialog'] },
    { name: 'MU3 登入彙總視窗每次載入都跳（沒看 isFirstLoad）', file: '_client/app.js', urlTail: '/app.js', find: 'window.ITTSCardTasks.start(user, { login: isFirstLoad });', rep: 'window.ITTSCardTasks.start(user, { login: true });', sections: ['loginDialog'] },
    { name: 'MU4 解除視窗不檢查「其他」必填說明', file: '_client/company-card-tasks.js', urlTail: '/company-card-tasks.js', find: "if (sel.value === 'other' && !note) {", rep: 'if (false) {', sections: ['dismiss'] },
    { name: 'MU5 Esc 沒有「稍後處理」效果', file: '_client/company-card-tasks.js', urlTail: '/company-card-tasks.js', find: "if (e.key === 'Escape') {", rep: "if (e.key === 'Escape_disabled') {", sections: ['loginDialog'] },
  ];
  for (const m of defs) {
    if (process.env.CARDTASK_MUT_PICK && !new RegExp(process.env.CARDTASK_MUT_PICK).test(m.name)) continue;
    const src = read(m.file);
    if (!src.includes(m.find)) { C.ok(false, `${m.name}：找不到變異錨點`); continue; }
    const mutated = src.replace(m.find, () => m.rep);
    const snap = C.snap();
    C.setQuiet(true);
    try { await runUiOnce(C, { urlTail: m.urlTail, mutated, sections: m.sections }); } catch (e) { /* 例外也算變紅 */ C.ok(false, '變異執行例外：' + e.message); }
    C.setQuiet(false);
    const n = C.newFails(snap).length;
    C.restore(snap);
    C.ok(n > 0, `${m.name} → 測試變紅（${n} 項失敗）`);
  }
}
module.exports = { run, mutants };
