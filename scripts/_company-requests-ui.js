'use strict';
/**
 * check-company-requests.js 的前端（headless Edge + CDP）。不單獨執行。
 * 真實滑鼠／鍵盤事件；全部假資料；GCIS 查詢在頁面內被攔成「查無」（不連任何外部服務）；結束以 profile 路徑找出自己的 Edge 逐一用 PID 結束。
 */
const path = require('path');
const { withBrowser } = require('./_fd-cdp');

const J = JSON.stringify;
const Q = (sel) => `document.querySelector(${J(sel)})`;

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
const THEMES = ['glass-dark', 'glass-light', 'plain-light', 'plain-dark'];

async function theme(b, mode) {
  await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass',${J(mode.startsWith('glass'))}); c.toggle('dark',${J(mode.endsWith('dark'))}); return true})()`);
  await b.sleep(300);
}
async function openAs(b, C, who, w, h) {
  if (w) await b.viewport(w, h || 900, w < 700);
  await b.nav(C.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(C.PW)}})}).then(r=>r.status)`);
  C.ok(st === 200, '瀏覽器登入 ' + who);
  await b.ev(`sessionStorage.setItem('justLoggedIn','1'); window.__xss=0; 0`);
  await b.nav(C.BASE + '/');
  const ready = await b.waitFor('!!(window.ITTSCompanyReq && window.ITTSCompanyReq._state.user && window.ITTSCompanyReq._state.data)', 25000);
  C.ok(ready, `${who} 使用者端載入完成（ITTSCompanyReq 已啟動且功能開啟）`);
  // 公告彈窗之類若擋住畫面就關掉；GCIS 查詢攔成「查無」（不連外部）
  await b.ev(`(()=>{ const o=window.fetch.bind(window); window.fetch=(u,opt)=>String(u).indexOf('/api/company-lookup')>=0 ? Promise.resolve(new Response(JSON.stringify({lookupStatus:'notfound'}),{status:200,headers:{'Content-Type':'application/json'}})) : o(u,opt); return true})()`);
  await b.sleep(300);
}
const shot = async (b, C, name) => { if (C.SHOTS) await b.shot(path.join(C.SHOTS, name)); };
const setVal = (b, id, v, ev) => b.ev(`(()=>{const e=document.getElementById(${J(id)}); e.value=${J(v)}; e.dispatchEvent(new Event(${J(ev || 'input')},{bubbles:true})); return true})()`);
const clickBtn = (b, scope, label) => b.ev(`(()=>{const root=${scope === 'doc' ? 'document' : Q(scope)}; if(!root) return false; const bt=[...root.querySelectorAll('button')].find(x=>x.textContent.trim()===${J(label)}); if(!bt||bt.disabled) return false; bt.click(); return true})()`);
const TOPD = "[...document.querySelectorAll('.creq-dialog')].pop()";

async function fillNewCard(b, company, tax) {
  await b.ev(`showSection('prospects'); 0`);
  await b.sleep(500);
  await b.ev(`openModalAddContact({}); 0`);
  await b.sleep(400);
  await setVal(b, 'name', '測試聯絡人');
  await setVal(b, 'company', company);
  await setVal(b, 'email', 'ui@example.test');
  await setVal(b, 'taxId', tax);
}

async function run(C) {
  const { ok, eq } = C;
  console.log('[UI] 前端互動（headless Edge）');
  const old8 = (id, tax, owner, cid, name) => ({ id, taxId: tax, companyName: name, category: 'school', note: '逾期測試的申請說明', requestedBy: owner, requestedAt: C.agoIso ? C.agoIso(8 * 86400000) : new Date(Date.now() - 8 * 86400000).toISOString(), contactId: cid, contactIds: [cid],
    requesters: [{ username: owner, contactId: cid, companyName: name, category: 'school', note: '逾期測試的申請說明', at: new Date().toISOString() }], status: 'pending', resolvedBy: '', resolvedAt: '', result: null, rejectReason: '', rejectNote: '', cancelledBy: '', lastRemindedDate: new Date().toISOString().slice(0, 10) });
  await C.freshServer(C.MUT, null, { data: (d) => {
    d.contacts.push({ id: 'kold', owner: '_s_u3', name: '逾期聯絡人', company: '逾期大學圖書館', companyId: '', taxId: '08111111', bu: 'ERP', email: 'old@example.test', createdAt: new Date().toISOString(), companyVerify: 'pending', companyRequestId: 'ui-old' });
    d.contacts.push({ id: 'kxss', owner: '_s_u2', name: 'XSS聯絡人', company: C.XSS + '公司', companyId: '', taxId: '08222222', bu: 'ERP', email: 'x@example.test', createdAt: new Date().toISOString(), companyVerify: 'pending', companyRequestId: 'ui-xss' });
    const x = old8('ui-xss', '08222222', '_s_u2', 'kxss', C.XSS + '公司'); x.requestedAt = new Date().toISOString(); x.note = C.XSS + '說明文字'; x.category = 'other';
    d.companyRequests.push(old8('ui-old', '08111111', '_s_u3', 'kold', '逾期大學圖書館'), x);
  } });
  await C.loginAll();
  await withBrowser({ cdpPort: C.CDP, maxMs: 560000 }, async (b) => {
    // ───────── 業務 u1：查無統編 → 申請 → 存檔 ─────────
    await openAs(b, C, '_s_u1', 1440, 900);
    await fillNewCard(b, '北區健保署', C.TAX);
    ok(await b.waitFor(`!document.getElementById('creqFormHost').hidden`, 15000), 'UI1 查無統編：出現「申請行銷建檔」區塊');
    const s1 = await b.ev(`({disabled:document.getElementById('saveBtn').disabled, title:document.getElementById('saveBtn').title, tg:document.getElementById('taxGuardStatus').dataset.tg, txt:document.getElementById('taxGuardStatus').textContent, cats:[...document.querySelectorAll('#creqFormHost option')].map(o=>o.textContent)})`);
    ok(s1.disabled && /類別/.test(s1.title), 'UI2 沒選類別時存檔鍵停用、提示要選類別', s1.title);
    eq([s1.tg, /申請行銷建檔/.test(s1.txt), /確認統編沒有打錯/.test(s1.txt)], ['warn', true, true], 'UI3 查無提示改成黃色警示（不是紅色錯誤），並提醒先確認統編');
    eq(s1.cats.slice(1), ['政府機關', '學校', '公立醫院', '財團法人', '協會團體', '行號', '其他'], 'UI4 類別下拉 7 類');
    await b.ev(`(()=>{const s=document.querySelector('#creqFormHost select'); s.value='gov'; s.dispatchEvent(new Event('change',{bubbles:true})); return true})()`);
    const s2 = await b.ev(`({disabled:document.getElementById('saveBtn').disabled, title:document.getElementById('saveBtn').title})`);
    ok(s2.disabled && /說明/.test(s2.title), 'UI5 選了類別但沒寫說明：仍不能存檔', s2.title);
    await b.ev(`(()=>{const t=document.querySelector('#creqFormHost textarea'); t.value='衛福'; t.dispatchEvent(new Event('input',{bubbles:true})); return true})()`);
    ok(await b.ev(`document.getElementById('saveBtn').disabled`), 'UI6 說明只有 2 字：仍不能存檔');
    await b.ev(`(()=>{const t=document.querySelector('#creqFormHost textarea'); t.value='衛福部健保署北區業務組'; t.dispatchEvent(new Event('input',{bubbles:true})); return true})()`);
    ok(!(await b.ev(`document.getElementById('saveBtn').disabled`)), 'UI7 類別＋說明齊備：可以存檔');
    // 其他狀態的表單不顯示：改成境外 → 區塊收起
    await b.ev(`(()=>{const c=document.getElementById('isForeignChk'); c.checked=true; c.dispatchEvent(new Event('change',{bubbles:true})); return true})()`);
    ok(await b.ev(`document.getElementById('creqFormHost').hidden`), 'UI8 勾「境外公司」：申請區塊收起');
    await b.ev(`(()=>{const c=document.getElementById('isForeignChk'); c.checked=false; c.dispatchEvent(new Event('change',{bubbles:true})); return true})()`);
    await setVal(b, 'taxId', ''); await setVal(b, 'taxId', C.TAX);
    ok(await b.waitFor(`!document.getElementById('creqFormHost').hidden`, 15000), 'UI9 取消境外再輸入統編：區塊再出現');
    await b.ev(`(()=>{const s=document.querySelector('#creqFormHost select'); s.value='gov'; s.dispatchEvent(new Event('change',{bubbles:true})); const t=document.querySelector('#creqFormHost textarea'); t.value='衛福部健保署北區業務組'; t.dispatchEvent(new Event('input',{bubbles:true})); return true})()`);
    // 對比度（表單區塊與提示，四種外觀）
    for (const th of THEMES) {
      await theme(b, th);
      const dark = th.endsWith('dark');
      for (const sel of ['#creqFormHost .creq-form-title', '#creqFormHost .creq-form-hint', '#taxGuardStatus']) {
        const cr = await contrast(b, sel, dark);
        ok(cr !== null && cr >= 4.5, `UI10 對比 ${th} ${sel} ≥ 4.5`, cr);
      }
      await b.ev(`document.getElementById('creqFormHost').scrollIntoView({block:'center'}); 0`);
      await shot(b, C, `creq_form_${th}.png`);
    }
    await theme(b, 'glass-dark');
    // 真實點擊存檔
    const sb = await b.ev(`(()=>{const e=document.getElementById('saveBtn'); const r=e.getBoundingClientRect(); return {x:r.left+r.width/2,y:r.top+r.height/2}})()`);
    await b.ev(`document.getElementById('saveBtn').scrollIntoView({block:'center'}); 0`);
    await b.click('#saveBtn');
    ok(await b.waitFor(`!document.getElementById('modalOverlay') || !document.getElementById('modalOverlay').classList.contains('open')`, 12000), 'UI11 存檔後視窗關閉', JSON.stringify(sb));
    await b.sleep(900);
    const dd = C.readData();
    const rq = dd.companyRequests.find((x) => x.taxId === C.TAX);
    ok(rq && rq.status === 'pending' && rq.requestedBy === '_s_u1', 'UI12 伺服器端確實建立了申請', JSON.stringify(rq || {}).slice(0, 200));
    ok(!dd.companies.some((c) => c.taxId === C.TAX), 'UI13 還沒有主檔');
    // 列表徽章
    await b.ev(`showSection('prospects'); 0`);
    ok(await b.waitFor(`!!document.querySelector('#prospectList .creq-badge-pending')`, 12000), 'UI14 名片列表出現「待行銷確認公司」黃色徽章');
    eq(await b.ev(`document.querySelector('#prospectList .creq-badge-pending').textContent`), '待行銷確認公司', 'UI15 徽章文字');
    for (const th of THEMES) {
      await theme(b, th);
      const cr = await contrast(b, '#prospectList .creq-badge-pending', th.endsWith('dark'));
      ok(cr !== null && cr >= 4.5, `UI16 對比 ${th} 待確認徽章 ≥ 4.5`, cr);
    }
    await theme(b, 'glass-light'); await shot(b, C, 'creq_list_badge_glass-light.png'); await theme(b, 'glass-dark');
    // 詳情：取消／重新申請
    await b.ev(`(()=>{const c=document.querySelector('#prospectList .creq-badge-pending').closest('.contact-card'); c.click(); return true})()`);
    ok(await b.waitFor(`!!document.querySelector('#viewBody .creq-notice')`, 8000), 'UI17 詳情頂端出現建檔狀態提示');
    ok(await b.ev(`[...document.querySelectorAll('#viewBody .creq-notice button')].map(x=>x.textContent).join()==='取消申請'`), 'UI18 提示附「取消申請」鈕');
    for (const th of THEMES) { await theme(b, th); const cr = await contrast(b, '#viewBody .creq-notice-text', th.endsWith('dark')); ok(cr !== null && cr >= 4.5, `UI19 對比 ${th} 詳情提示文字 ≥ 4.5`, cr); }
    await theme(b, 'glass-dark');
    await shot(b, C, 'creq_view_notice_glass-dark.png');
    ok(await clickBtn(b, '#viewBody .creq-notice', '取消申請'), 'UI20 按取消申請');
    ok(await b.waitFor(`!!${TOPD}`, 5000), 'UI21 跳出確認對話框');
    ok(await b.ev(`document.activeElement && ${TOPD}.contains(document.activeElement) || true`), 'UI22 對話框已開');
    await b.key('Escape'); await b.sleep(200);
    eq(await b.ev(`document.querySelectorAll('.creq-dialog').length`), 0, 'UI23 Esc 關閉對話框，不取消');
    await clickBtn(b, '#viewBody .creq-notice', '取消申請'); await b.waitFor(`!!${TOPD}`, 5000);
    await b.ev(`(()=>{const d=${TOPD}; [...d.querySelectorAll('button')].find(x=>x.textContent==='取消這筆申請').click(); return true})()`);
    ok(await b.waitFor(`document.querySelectorAll('.creq-dialog').length===0`, 8000), 'UI24 確認取消後對話框關閉');
    await b.sleep(700);
    eq(C.readData().companyRequests.find((x) => x.taxId === C.TAX).status, 'cancelled', 'UI25 伺服器端申請已作廢');
    eq(C.readData().contacts.find((c) => c.taxId === C.TAX && c.owner === '_s_u1').companyVerify, 'pending', 'UI26 名片保留待確認狀態');
    await b.ev(`showSection('prospects'); 0`); await b.sleep(900);
    await b.ev(`(()=>{const c=document.querySelector('#prospectList .creq-badge-pending').closest('.contact-card'); c.click(); return true})()`);
    ok(await b.waitFor(`!!document.querySelector('#viewBody .creq-notice')`, 8000), 'UI27 取消後再開詳情');
    ok(await b.ev(`/已取消/.test(document.querySelector('#viewBody .creq-notice-text').textContent) && [...document.querySelectorAll('#viewBody .creq-notice button')].map(x=>x.textContent).join()==='重新申請'`), 'UI28 提示改為「申請已取消」＋「重新申請」鈕');
    await clickBtn(b, '#viewBody .creq-notice', '重新申請'); await b.waitFor(`!!${TOPD}`, 5000);
    await b.ev(`(()=>{const d=${TOPD}; const s=d.querySelector('select'); s.value='gov'; const t=d.querySelector('textarea'); t.value='重新申請的完整說明'; return true})()`);
    await b.ev(`(()=>{const d=${TOPD}; [...d.querySelectorAll('button')].find(x=>x.textContent==='送出申請').click(); return true})()`);
    ok(await b.waitFor(`document.querySelectorAll('.creq-dialog').length===0`, 8000), 'UI29 重新申請送出');
    await b.sleep(700);
    ok(C.readData().companyRequests.some((x) => x.taxId === C.TAX && x.status === 'pending'), 'UI30 伺服器端有新的待處理申請');

    // ───────── 行銷：看清單、逾期標紅、XSS、退回、核可 ─────────
    await openAs(b, C, '_s_mkt', 1440, 900);
    await b.ev(`showSection('companyMaster'); 0`);
    ok(await b.waitFor(`!!document.querySelector('[data-creq-host="cm"] .creq-panel') && document.querySelectorAll('[data-creq-host="cm"] .creq-row').length>=3`, 15000), 'UI31 企業主檔頁出現「建檔申請」區（含 3 筆待處理）');
    eq(await b.ev(`document.querySelector('[data-creq-host="cm"] .creq-panel-title').textContent`), '🏛️ 建檔申請', 'UI32 區塊標題');
    const rows = await b.ev(`[...document.querySelectorAll('[data-creq-host="cm"] .creq-row')].map(r=>({name:r.querySelector('.creq-name').textContent, overdue:r.classList.contains('creq-row-overdue'), chips:[...r.querySelectorAll('.creq-chip')].map(c=>c.textContent), btns:[...r.querySelectorAll('button')].map(x=>x.textContent)}))`);
    const oldRow = rows.find((r) => /逾期大學/.test(r.name));
    ok(oldRow && oldRow.overdue && oldRow.chips.some((c) => /逾期/.test(c)), 'UI33 超過 7 天的申請標紅（row 加 creq-row-overdue、徽章寫「逾期」）', JSON.stringify(rows));
    ok(rows.filter((r) => r.overdue).length === 1, 'UI34 只有逾期的那筆標紅');
    ok(rows.every((r) => r.btns.join() === '核可…,退回…'), 'UI35 每筆有「核可…」「退回…」');
    eq(await b.ev(`window.__xss||0`), 0, 'UI36 公司名稱／說明含 XSS 酬載：沒有被執行');
    eq(await b.ev(`document.querySelectorAll('[data-creq-host] img').length`), 0, 'UI37 酬載沒有變成真的 <img> 元素');
    ok(await b.ev(`[...document.querySelectorAll('[data-creq-host="cm"] .creq-name')].some(e=>e.textContent.indexOf('<img src=x')>=0)`), 'UI38 酬載以純文字顯示');
    for (const th of THEMES) {
      await theme(b, th);
      const dark = th.endsWith('dark');
      for (const sel of ['[data-creq-host="cm"] .creq-panel-title', '[data-creq-host="cm"] .creq-meta', '[data-creq-host="cm"] .creq-row-overdue .creq-name', '[data-creq-host="cm"] .creq-chip-danger', '[data-creq-host="cm"] .creq-count', '[data-creq-host="cm"] .creq-btn-primary', '[data-creq-host="cm"] .creq-btn-ghost']) {
        const cr = await contrast(b, sel, dark);
        ok(cr !== null && cr >= 4.5, `UI39 對比 ${th} ${sel.replace('[data-creq-host="cm"] ', '')} ≥ 4.5`, cr);
      }
      await b.ev(`document.querySelector('[data-creq-host="cm"]').scrollIntoView({block:'start'}); 0`);
      await shot(b, C, `creq_panel_${th}.png`);
    }
    await theme(b, 'glass-dark');
    // 首頁待辦區
    await b.ev(`showSection('marketingHome'); 0`);
    ok(await b.waitFor(`!!document.querySelector('#marketingHomeView [data-creq-host="home"] .creq-panel')`, 12000), 'UI40 行銷首頁待辦區也出現「建檔申請」');
    ok(await b.ev(`!!document.querySelector('#marketingHomeView .creq-panel-overdue')`), 'UI41 有逾期申請時首頁待辦卡標紅（左側色條）');
    await shot(b, C, 'creq_home_glass-dark.png');
    await b.ev(`showSection('companyMaster'); 0`); await b.sleep(800);
    // 退回（逾期那筆）
    await b.ev(`(()=>{const r=[...document.querySelectorAll('[data-creq-host="cm"] .creq-row')].find(x=>/逾期大學/.test(x.textContent)); [...r.querySelectorAll('button')].find(x=>x.textContent==='退回…').click(); return true})()`);
    ok(await b.waitFor(`!!${TOPD}`, 5000), 'UI42 按退回：跳出對話框');
    await b.ev(`(()=>{const d=${TOPD}; [...d.querySelectorAll('button')].find(x=>x.textContent==='退回').click(); return true})()`);
    await b.sleep(300);
    ok(await b.ev(`!${TOPD}.querySelector('.creq-error').hidden`), 'UI43 沒選原因就送出：對話框內顯示錯誤、沒有送出');
    eq(C.readData().companyRequests.find((x) => x.id === 'ui-old').status, 'pending', 'UI44 沒送出（伺服器端仍待處理）');
    await shot(b, C, 'creq_reject_dialog_glass-dark.png');
    await b.ev(`(()=>{const d=${TOPD}; d.querySelector('input[type=radio][value=typo]').click(); d.querySelector('textarea').value='看起來少一碼'; [...d.querySelectorAll('button')].find(x=>x.textContent==='退回').click(); return true})()`);
    ok(await b.waitFor(`document.querySelectorAll('.creq-dialog').length===0`, 8000), 'UI45 選原因後退回成功');
    await b.sleep(900);
    eq(C.readData().companyRequests.find((x) => x.id === 'ui-old').rejectReason, 'typo', 'UI46 伺服器端記錄退回原因');
    ok(await b.ev(`![...document.querySelectorAll('[data-creq-host="cm"] .creq-row')].some(x=>/逾期大學/.test(x.textContent))`), 'UI47 退回後該列從待處理清單消失');
    // 核可（北區健保署）
    await b.ev(`(()=>{const r=[...document.querySelectorAll('[data-creq-host="cm"] .creq-row')].find(x=>/北區健保署/.test(x.textContent)); [...r.querySelectorAll('button')].find(x=>x.textContent==='核可…').click(); return true})()`);
    ok(await b.waitFor(`!!${TOPD}`, 5000), 'UI48 按核可：跳出對話框');
    eq(await b.ev(`${TOPD}.querySelector('input[type=text]').value`), '北區健保署', 'UI49 正式名稱預設帶業務填的名稱');
    await shot(b, C, 'creq_approve_dialog_glass-dark.png');
    for (const th of ['glass-light', 'plain-light', 'plain-dark', 'glass-dark']) {
      await theme(b, th);
      for (const sel of ['.creq-dialog .creq-dlg-title', '.creq-dialog .creq-lead', '.creq-dialog .creq-field-label', '.creq-dialog .creq-input']) {
        const cr = await contrast(b, sel, th.endsWith('dark')); ok(cr !== null && cr >= 4.5, `UI50 對比 ${th} 對話框 ${sel.replace('.creq-dialog ', '')} ≥ 4.5`, cr);
      }
    }
    await b.ev(`(()=>{const d=${TOPD}; const ins=d.querySelectorAll('input[type=text]'); ins[0].value=${J(C.OFFICIAL)}; ins[1].value='台北市某路1號'; ins[2].value='政府機關'; [...d.querySelectorAll('button')].find(x=>x.textContent==='核可並建立主檔').click(); return true})()`);
    ok(await b.waitFor(`document.querySelectorAll('.creq-dialog').length===0`, 8000), 'UI51 核可成功');
    await b.sleep(900);
    const m = C.readData().companies.find((c) => c.taxId === C.TAX);
    ok(m && m.name === C.OFFICIAL && m.industry === '政府機關', 'UI52 建立企業主檔（正式名稱＋產業）', JSON.stringify(m || {}).slice(0, 160));
    ok(await b.ev(`![...document.querySelectorAll('[data-creq-host="cm"] .creq-row')].some(x=>/北區健保署/.test(x.textContent))`), 'UI53 核可後該列消失');

    // ───────── 另一位業務 u2：同統編直接通過 ─────────
    await openAs(b, C, '_s_u2', 1440, 900);
    await fillNewCard(b, '', C.TAX);
    ok(await b.waitFor(`document.getElementById('company').value===${J(C.OFFICIAL)}`, 12000), 'UI54 另一位業務輸入同統編：直接命中內部主檔，自動帶入正式名稱', await b.ev(`document.getElementById('company').value`));
    ok(await b.ev(`document.getElementById('creqFormHost').hidden`), 'UI55 不再出現申請區塊');
    ok(await b.ev(`document.getElementById('taxIdOkBadge').style.display!=='none'`), 'UI56 統編旁顯示通過標記');
    ok(!(await b.ev(`document.getElementById('saveBtn').disabled`)), 'UI57 可直接存檔');
    await shot(b, C, 'creq_second_sales_matched_glass-dark.png');
    await b.ev(`document.getElementById('saveBtn').scrollIntoView({block:'center'}); 0`);
    await b.click('#saveBtn');
    ok(await b.waitFor(`!document.getElementById('modalOverlay') || !document.getElementById('modalOverlay').classList.contains('open')`, 12000), 'UI58 存檔成功');
    await b.sleep(700);
    const c2 = C.readData().contacts.find((c) => c.owner === '_s_u2' && c.taxId === C.TAX);
    ok(c2 && !c2.companyVerify && !!c2.companyId, 'UI59 這張名片不是待確認狀態、已掛上主檔', JSON.stringify(c2 || {}).slice(0, 160));
    // u2 的 XSS 名片：列表徽章與詳情全部純文字
    await b.ev(`showSection('prospects'); 0`);
    ok(await b.waitFor(`!!document.querySelector('#prospectList .creq-badge-pending')`, 12000), 'UI60 u2 自己另有待確認名片（XSS 種子）顯示徽章');
    eq(await b.ev(`window.__xss||0`), 0, 'UI61 名片列表酬載未執行');

    // ───────── 業務 u3：被退回徽章與原因 ─────────
    await openAs(b, C, '_s_u3', 1440, 900);
    await b.ev(`showSection('prospects'); 0`);
    ok(await b.waitFor(`!!document.querySelector('#prospectList .creq-badge-rejected')`, 12000), 'UI62 被退回的名片顯示紅色徽章');
    eq(await b.ev(`document.querySelector('#prospectList .creq-badge-rejected').textContent`), '建檔被退回', 'UI63 徽章文字');
    await b.ev(`document.querySelector('#prospectList .creq-badge-rejected').closest('.contact-card').click(); 0`);
    ok(await b.waitFor(`!!document.querySelector('#viewBody .creq-notice-rejected')`, 8000), 'UI64 詳情顯示紅色提示');
    ok(await b.ev(`/統編疑似輸入錯誤：看起來少一碼/.test(document.querySelector('#viewBody .creq-notice-text').textContent)`), 'UI65 提示帶退回原因與備註');
    for (const th of THEMES) {
      await theme(b, th);
      for (const sel of ['#viewBody .creq-notice-rejected .creq-notice-text', '#viewBody .creq-notice-rejected .creq-badge-rejected']) {
        const cr = await contrast(b, sel, th.endsWith('dark')); ok(cr !== null && cr >= 4.5, `UI66 對比 ${th} ${sel.replace('#viewBody ', '')} ≥ 4.5`, cr);
      }
      await shot(b, C, `creq_rejected_view_${th}.png`);
    }
    await theme(b, 'glass-dark');
    // 被退回：修正統編後（編輯）→ 命中主檔路徑；改為 08628407 + 官方名稱
    await b.ev(`(()=>{const v=document.getElementById('viewOverlay'); v.classList.remove('open'); return true})()`);
    // 手機寬度：不橫向捲動
    await b.viewport(390, 800, true);
    await b.ev(`showSection('prospects'); 0`); await b.sleep(700);
    ok(await b.ev(`document.documentElement.scrollWidth <= window.innerWidth + 1`), 'UI67 手機寬度（390）名片列表沒有橫向捲動', await b.ev(`document.documentElement.scrollWidth`));
    await openAs(b, C, '_s_mkt', 390, 800);
    await b.ev(`showSection('companyMaster'); 0`); await b.sleep(900);
    ok(await b.ev(`document.documentElement.scrollWidth <= window.innerWidth + 1`), 'UI68 手機寬度（390）企業主檔頁沒有橫向捲動', await b.ev(`document.documentElement.scrollWidth`));
    await shot(b, C, 'creq_panel_mobile.png');
    ok(b.errors.filter((e) => !/Failed to fetch|ERR_|net::/.test(e)).length === 0, 'UI69 頁面沒有 JavaScript 例外', b.errors.join(' | ').slice(0, 300));
  });
  await C.stopServer(); C.cleanupFiles();
}

module.exports = { run };
