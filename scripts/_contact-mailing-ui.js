'use strict';
/**
 * check-contact-mailing.js 的前端（headless Edge + CDP）與前端變異測試。不單獨執行。
 * 真實滑鼠點擊核取方塊與儲存鈕；全部假資料；結束以 profile 路徑找出自己的 Edge 逐一用 PID 結束。
 */
const fs = require('fs');
const path = require('path');
const { withBrowser } = require('./_fd-cdp');

const J = JSON.stringify;
const shot = async (b, C, name) => { if (C.SHOTS) await b.shot(path.join(C.SHOTS, name)); };

async function openAs(b, C, who, w, h) {
  if (w) await b.viewport(w, h || 900, w < 700);
  await b.nav(C.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(C.PW)}})}).then(r=>r.status)`);
  C.ok(st === 200, '瀏覽器登入 ' + who);
  await b.ev(`sessionStorage.setItem('justLoggedIn','1'); window.__xss=0; 0`);
  await b.nav(C.BASE + '/');
  const ready = await b.waitFor(`typeof openModal==='function' && typeof openModalAddContact==='function' && !!document.getElementById('saveBtn') && Array.isArray(allContacts) && allContacts.length>0`, 25000);
  C.ok(ready, `${who} 使用者端載入完成`);
  await b.sleep(500);
}
const SET = (id, v) => `(()=>{const e=document.getElementById(${J(id)}); e.value=${J(v)}; e.dispatchEvent(new Event('input',{bubbles:true})); e.dispatchEvent(new Event('change',{bubbles:true})); return true})()`;
const STATE = `(()=>{const m=document.getElementById('mailingAddress'),c=document.getElementById('mailingSameAsCompany'),a=document.getElementById('address');return {mail:m.value,ro:m.readOnly,chk:c.checked,addr:a.value}})()`;
const modalClosed = (b) => b.waitFor(`!document.getElementById('modalOverlay').classList.contains('open')`, 8000);
const srvContact = (b, name) => b.ev(`fetch('/api/contacts').then(r=>r.json()).then(l=>l.find(c=>c.name===${J(name)})||null)`);
const reload = (b) => b.ev(`loadContacts().then(()=>allContacts.length)`);
async function saveForm(b) { await b.click('#saveBtn'); return modalClosed(b); }

async function uiFlow(b, C) {
  console.log('  [UI] 新增／編輯／取消勾選／詳情');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u1', 1440, 1000);
  // 1. 新增（同公司既有名片帶入公司資料）
  await b.ev(`openModalAddContact({company:'假公司甲', taxId:'90000001', address:'UI路1號', mailingAddress:'', mailingSameAsCompany:false}); 0`);
  await b.sleep(500);
  let s = await b.ev(STATE);
  eq([s.chk, s.ro, s.mail, s.addr], [false, false, '', 'UI路1號'], 'UI1 新增名片：預設未勾選、郵寄地址可輸入、公司地址已帶入');
  await b.click('#mailingSameAsCompany');
  s = await b.ev(STATE);
  eq([s.chk, s.ro, s.mail], [true, true, 'UI路1號'], 'UI2 勾「同公司地址」：郵寄地址帶入公司地址並唯讀');
  await b.ev(SET('address', 'UI路1號之一'));
  s = await b.ev(STATE);
  eq(s.mail, 'UI路1號之一', 'UI3 勾選狀態下改公司地址欄（input 事件）：郵寄地址同步');
  await b.ev(SET('address', 'UI路1號'));
  await b.ev(SET('name', 'UI甲')); await b.ev(SET('email', 'ui-a@example.test'));
  await shot(b, C, 'mailaddr_new_checked_1440.png');
  ok(await saveForm(b), 'UI4 存檔成功（表單關閉）');
  let c = await srvContact(b, 'UI甲');
  eq([c && c.mailingSameAsCompany, c && c.mailingAddress], [true, 'UI路1號'], 'UI5 伺服器端：旗標 true、郵寄地址＝公司地址（前端有送出欄位）');
  // 2. 重開編輯仍勾選
  await reload(b);
  await b.ev(`openModal(allContacts.find(x=>x.name==='UI甲')); 0`); await b.sleep(400);
  s = await b.ev(STATE);
  eq([s.chk, s.ro, s.mail, s.addr], [true, true, 'UI路1號', 'UI路1號'], 'UI6 重開編輯：仍勾選、唯讀、值正確');
  // 3. 改公司地址後存檔 → 郵寄地址同步
  await b.ev(SET('address', 'UI新路2號'));
  s = await b.ev(STATE);
  eq(s.mail, 'UI新路2號', 'UI7 編輯時改公司地址：郵寄地址即時同步');
  ok(await saveForm(b), 'UI8 存檔成功');
  c = await srvContact(b, 'UI甲');
  eq([c.address, c.mailingAddress, c.mailingSameAsCompany], ['UI新路2號', 'UI新路2號', true], 'UI9 伺服器端郵寄地址跟著新公司地址');
  // 4. 取消勾選可自由輸入
  await reload(b);
  await b.ev(`openModal(allContacts.find(x=>x.name==='UI甲')); 0`); await b.sleep(400);
  await b.click('#mailingSameAsCompany');
  s = await b.ev(STATE);
  eq([s.chk, s.ro, s.mail], [false, false, 'UI新路2號'], 'UI10 取消勾選：可編輯、保留目前值');
  await b.ev(SET('mailingAddress', '自訂郵寄路3號'));
  await b.ev(SET('address', '公司又搬家路9號'));
  s = await b.ev(STATE);
  eq(s.mail, '自訂郵寄路3號', 'UI11 取消勾選後改公司地址：郵寄地址不跟著動');
  ok(await saveForm(b), 'UI12 存檔成功');
  c = await srvContact(b, 'UI甲');
  eq([c.address, c.mailingAddress, c.mailingSameAsCompany], ['公司又搬家路9號', '自訂郵寄路3號', false], 'UI13 伺服器端：自由輸入的郵寄地址保留');
  await reload(b);
  await b.ev(`openModal(allContacts.find(x=>x.name==='UI甲')); 0`); await b.sleep(400);
  s = await b.ev(STATE);
  eq([s.chk, s.ro, s.mail], [false, false, '自訂郵寄路3號'], 'UI14 重開編輯：未勾選、值正確');
  await shot(b, C, 'mailaddr_edit_unchecked_1440.png');
  await b.ev(`closeModal(); 0`);
  // 5. 詳情
  await b.ev(`openView(allContacts.find(x=>x.name==='UI甲').id); 0`); await b.sleep(300);
  const vt = await b.ev(`document.getElementById('viewBody').innerText`);
  ok(/郵寄地址/.test(vt) && vt.includes('自訂郵寄路3號'), 'UI15 詳情檢視顯示郵寄地址', vt.slice(0, 200));
  await shot(b, C, 'mailaddr_view_1440.png');
  await b.ev(`document.getElementById('viewOverlay').classList.remove('open'); 0`);
  // 6. 同公司新增聯絡人帶入（以勾選的名片為參照）
  await b.ev(`openModalAddContact({company:'假公司甲', taxId:'90000001', address:'參照路', mailingAddress:'參照路', mailingSameAsCompany:true}); 0`); await b.sleep(300);
  s = await b.ev(STATE);
  eq([s.chk, s.ro, s.mail], [true, true, '參照路'], 'UI16 openModalAddContact 帶入參照名片的郵寄地址／同公司旗標');
  await b.ev(`closeModal(); 0`);
  // 7. 輸入公司名自動複製同公司既有名片：空白才補
  await b.ev(`fetch('/api/contacts/'+allContacts.find(x=>x.company==='假公司甲').id,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({mailingAddress:'帶入用郵寄路',mailingSameAsCompany:false})}).then(r=>r.status)`);
  await reload(b);
  await b.ev(`openModal(); 0`); await b.sleep(300);
  await b.ev(SET('company', '假公司甲'));
  ok(await b.waitFor(`document.getElementById('address').value!==''`, 6000), 'UI17 輸入公司名：自動帶入既有名片資料');
  s = await b.ev(STATE);
  eq(s.mail, '帶入用郵寄路', 'UI18 郵寄地址一併帶入（欄位空白才補）');
  await b.ev(`closeModal(); 0`);
}

async function uiXss(b, C) {
  console.log('  [UI] XSS');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u1', 1440, 1000);
  const id = await b.ev(`fetch('/api/contacts').then(r=>r.json()).then(l=>l.find(c=>c.name==='名片k1').id)`);
  const st = await b.ev(`fetch('/api/contacts/${id}',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({mailingAddress:${J(C.XSS)},mailingSameAsCompany:false})}).then(r=>r.status)`);
  eq(st, 200, 'UI20 以 API 存入含酬載的郵寄地址');
  await reload(b);
  await b.ev(`openView(${J(id)}); 0`); await b.sleep(400);
  eq(await b.ev(`window.__xss||0`), 0, 'UI21 詳情檢視：酬載沒有被執行');
  eq(await b.ev(`document.querySelectorAll('#viewBody img[src="x"]').length`), 0, 'UI22 酬載沒有變成 <img> 元素');
  ok((await b.ev(`document.getElementById('viewBody').innerText`)).includes('<img src=x'), 'UI23 酬載以純文字顯示');
  await b.ev(`document.getElementById('viewOverlay').classList.remove('open'); 0`);
  await b.ev(`openModal(allContacts.find(x=>x.id===${J(id)})); 0`); await b.sleep(300);
  eq([await b.ev(`window.__xss||0`), await b.ev(`document.getElementById('mailingAddress').value`)], [0, C.XSS], 'UI24 編輯表單：酬載只是欄位文字（value）');
  await b.ev(`closeModal(); 0`);
}

async function uiLayout(b, C) {
  console.log('  [UI] 手機寬度與深色');
  const { ok } = C;
  await openAs(b, C, '_s_u1', 375, 800);
  await b.ev(`fetch('/api/contacts').then(r=>r.json()).then(l=>{const c=l.find(x=>x.name==='名片k1'); return fetch('/api/contacts/'+c.id,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({mailingAddress:'手機版很長的郵寄地址測試路一段二百三十四號十二樓之五',mailingSameAsCompany:false})})}).then(r=>r.status)`);
  await reload(b);
  const probe = `(()=>{const m=document.getElementById('mailingAddress'),a=document.getElementById('address'),g=document.getElementById('mailingAddressGroup'),mod=document.querySelector('#modalOverlay .modal');const r=m.getBoundingClientRect(),ra=a.getBoundingClientRect();const cs=getComputedStyle(m),ca=getComputedStyle(a);return {right:Math.round(r.right),vw:window.innerWidth,docW:document.documentElement.scrollWidth,modalOverflow:mod.scrollWidth-mod.clientWidth,sameBg:cs.backgroundColor===ca.backgroundColor,sameColor:cs.color===ca.color,sameBorder:cs.borderTopColor===ca.borderTopColor,sameH:Math.abs(r.height-ra.height)<1,gVis:g.getBoundingClientRect().width>0}})()`;
  for (const mode of ['plain-light', 'plain-dark', 'glass-dark', 'glass-light']) {
    await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass',${J(mode.startsWith('glass'))}); c.toggle('dark',${J(mode.endsWith('dark'))}); return true})()`);
    await b.ev(`openModal(allContacts.find(x=>x.name==='名片k1')); 0`); await b.sleep(400);
    await b.ev(`document.getElementById('mailingAddressGroup').scrollIntoView({block:'center'}); 0`); await b.sleep(200);
    const p = await b.ev(probe);
    ok(p.gVis && p.right <= p.vw && p.docW <= p.vw && p.modalOverflow <= 1, `UI30 375px ${mode}：郵寄地址欄在視窗內、沒有橫向捲動`, J(p));
    ok(p.sameBg && p.sameColor && p.sameBorder && p.sameH, `UI31 375px ${mode}：郵寄地址欄樣式與地址欄一致（深/淺色不破版）`, J(p));
    await shot(b, C, `mailaddr_375_${mode}.png`);
    await b.click('#mailingSameAsCompany');
    await b.sleep(150);
    await shot(b, C, `mailaddr_375_${mode}_checked.png`);
    await b.ev(`closeModal(); 0`);
  }
  await b.ev(`document.body.classList.remove('glass','dark'); 0`);
  await b.viewport(1440, 1000, false);
}

async function runUiOnce(C, mut) {
  const fns = { flow: uiFlow, xss: uiXss, layout: uiLayout };
  const run = (mut && mut.sections) || ['flow', 'xss', 'layout'];
  await C.freshServer(null);
  await C.loginAll();
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

async function mutants(C) {
  console.log('[變異·前端] 把前端檔換成改壞的版本，對應斷言必須變紅');
  const read = (rel) => fs.readFileSync(path.join(C.ROOT, rel), 'utf8');
  const defs = [
    { name: 'MU1 前端儲存不送郵寄地址欄位', file: '_client/app.js', urlTail: '/app.js', find: "    mailingAddress: $('mailingSameAsCompany').checked ? $('address').value.trim() : $('mailingAddress').value.trim(),\n    mailingSameAsCompany: $('mailingSameAsCompany').checked,\n", rep: '', sections: ['flow'] },
    { name: 'MU2 公司地址欄 input 不再同步郵寄地址', file: '_client/app.js', urlTail: '/app.js', find: "if (addr) addr.addEventListener('input', mailingSyncUi);", rep: '', sections: ['flow'] },
    { name: 'MU3 詳情檢視郵寄地址未 escape', file: '_client/app.js', urlTail: '/app.js', find: "${escapeHtml(c.mailingAddress) || '-'}", rep: "${c.mailingAddress || '-'}", sections: ['xss'] },
    { name: 'MU4 編輯載入不還原勾選狀態', file: '_client/app.js', urlTail: '/app.js', find: "chk.checked = !!(src && src.mailingSameAsCompany);", rep: 'chk.checked = false;', sections: ['flow'] },
  ];
  for (const m of defs) {
    if (process.env.CMAIL_MUT_PICK && !new RegExp(process.env.CMAIL_MUT_PICK).test(m.name)) continue;
    const src = read(m.file);
    if (!src.includes(m.find)) { C.ok(false, `${m.name}：找不到變異錨點`); continue; }
    const mutated = src.replace(m.find, () => m.rep);
    const snap = C.snap();
    C.setQuiet(true);
    try { await runUiOnce(C, { urlTail: m.urlTail, mutated, sections: m.sections }); } catch (e) { C.ok(false, '變異執行例外：' + e.message); }
    C.setQuiet(false);
    const fs2 = C.newFails(snap);
    C.restore(snap);
    console.log('  ' + (fs2.length > 0 ? '✓ 變紅' : '✗ 沒變紅') + '：' + m.name + '（' + fs2.length + ' 項）');
    C.ok(fs2.length > 0, `${m.name} → 測試變紅（${fs2.length} 項失敗：${fs2.slice(0, 2).join(' / ').slice(0, 80)}）`);
  }
}
module.exports = { run, mutants };
