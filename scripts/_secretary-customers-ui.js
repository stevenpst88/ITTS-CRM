'use strict';
/**
 * check-secretary-customers.js 的前端（headless Edge + CDP）與前端變異測試。不單獨執行。
 * 真實滑鼠點擊側欄／列／按鈕；全部假資料；結束以 profile 路徑找出自己的 Edge 逐一用 PID 結束。
 */
const fs = require('fs');
const path = require('path');
const { withBrowser } = require('./_fd-cdp');

const J = JSON.stringify;
const XSS = '<img src=x onerror="window.__xss=1">';
const shot = async (b, C, name) => { if (C.SHOTS) await b.shot(path.join(C.SHOTS, name)); };
const XSS_SEED = { data: (d) => { d.contacts.find((c) => c.id === 'w1').mailingAddress = XSS; return d; } };

async function openAs(b, C, who, w, h, opts) {
  opts = opts || {};
  if (w) await b.viewport(w, h || 900, w < 700);
  await b.nav(C.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(C.PW)}})}).then(r=>r.status)`);
  C.ok(st === 200, '瀏覽器登入 ' + who);
  await b.ev(`sessionStorage.setItem('justLoggedIn','1'); window.__xss=0; ${opts.last ? `localStorage.setItem('lastSection',${J(opts.last)});` : ''} 0`);
  await b.nav(C.BASE + '/');
  const ready = await b.waitFor(`typeof showSection==='function' && !!document.getElementById('dashboardView') && typeof userPermissions==='object' && !!userPermissions.features`, 15000);
  C.ok(ready, `${who} 使用者端載入完成`);
  await b.sleep(700);
}
// 側欄是分組收合的：導覽項在「銷售作業」組內，先用真實點擊展開該組（已展開就不點）再點導覽項
async function openNav(b) {
  const open = await b.ev(`document.getElementById('sgSales').classList.contains('sg-open')`);
  if (!open) { await b.click('#sgSalesHd'); await b.sleep(500); }
  await b.click('#navSecretaryCustomers');
}
const VIEW_TEXT = `(()=>{const v=document.getElementById('secretaryCustomersView');return v?v.innerText:null})()`;
const VIEW_VIS = `(()=>{const v=document.getElementById('secretaryCustomersView');return !!v && v.style.display!=='none' && v.offsetParent!==null})()`;
const ROWS = `document.querySelectorAll('#secretaryCustomersView .scu-row').length`;

async function uiFlow(b, C) {
  console.log('  [UI] 側欄／清單／搜尋／詳情／複製／F5');
  const { ok, eq } = C;
  await openAs(b, C, '_s_sec', 1440, 1000);
  ok(await b.waitFor(`!!document.getElementById('navSecretaryCustomers')`, 8000), 'UI1 秘書登入後探測 200 → 側欄出現「我的客戶（唯讀）」');
  ok(/我的客戶（唯讀）/.test(await b.ev(`document.getElementById('navSecretaryCustomers').textContent`)), 'UI1b 導覽文字');
  await openNav(b);
  ok(await b.waitFor(VIEW_VIS, 8000), 'UI2 真實點擊側欄 → 區段顯示');
  ok(await b.waitFor(`${ROWS}===4`, 8000), 'UI3 顯示 4 家公司（同 BU 已成交）', await b.ev(ROWS));
  eq(await b.ev(`document.getElementById('topbarTitle').textContent`), '我的客戶（唯讀）', 'UI3b 頁首標題');
  eq(await b.ev(`document.getElementById('dashboardView').style.display`), 'none', 'UI3c 首頁已隱藏');
  const txt = await b.ev(VIEW_TEXT);
  ok(/甲公司/.test(txt) && /乙公司/.test(txt) && /丙公司/.test(txt) && /丁公司/.test(txt), 'UI4 清單含各公司');
  ok(!/庚公司|戊公司|己公司|辛公司|壬公司|癸公司|子公司|跨公司|KA公司/.test(txt), 'UI5 清單沒有未成交／其他 BU／刪除／待補客戶');
  ok(!/(NOTE_PII_|MEMO_PII_|HOBBY_PII_|CARD_PII_|EVIL_PII_|@pii\.example\.test)/.test(txt), 'UI6 畫面文字沒有 Email／備註／名片圖等 PII 哨兵', (txt.match(/\S*_PII_\S*/) || [])[0]);
  ok(/TEL_PII_w1 \(分機 EXT_PII_w1\)/.test(txt) && /MOB_PII_w1/.test(txt) && /TEL_PII_w3 \(分機 EXT_PII_w3\)/.test(txt) && !/_PII_w3b/.test(txt), 'UI6b 列表顯示主要聯絡人的「電話 (分機 N)」與手機，不含同公司非主要者');
  ok((txt.match(/（個資已遮蔽）/g) || []).length >= 4, 'UI7 列表每列 Email 欄仍顯示「（個資已遮蔽）」', (txt.match(/（個資已遮蔽）/g) || []).length);
  ok(/<img src=x/.test(txt) && /ADDR-B路2號/.test(txt) && /ADDR-C路3號/.test(txt), 'UI8 郵寄地址欄顯示三種情境（w1 的郵寄地址在此被換成 XSS 酬載字串，以純文字顯示）');
  ok(/（公司登記地址，未另填郵寄地址）/.test(txt), 'UI9 舊資料 fallback 加註「（公司登記地址，未另填郵寄地址）」');
  ok(/（未填）/.test(txt), 'UI9b 兩種地址都沒填 → 顯示「（未填）」');
  await shot(b, C, 'scu_list_1440_light.png');
  // 搜尋
  await b.focus('#secretaryCustomersView .scu-search');
  await b.type('乙公司');
  ok(await b.waitFor(`${ROWS}===1`, 4000), 'UI10 搜尋「乙公司」→ 1 筆（真實鍵盤輸入）');
  await b.ev(`(()=>{const e=document.querySelector('#secretaryCustomersView .scu-search'); e.value='不存在的字串'; e.dispatchEvent(new Event('input',{bubbles:true})); return 0})()`);
  ok(await b.waitFor(`!!document.querySelector('#secretaryCustomersView .scu-empty')`, 3000), 'UI11 無結果 → 顯示空狀態');
  await b.ev(`(()=>{const e=document.querySelector('#secretaryCustomersView .scu-search'); e.value=''; e.dispatchEvent(new Event('input',{bubbles:true})); return 0})()`);
  ok(await b.waitFor(`${ROWS}===4`, 3000), 'UI12 清空搜尋 → 回 4 家公司');
  // 詳情（fallback 那筆）
  await b.click(`#secretaryCustomersView .scu-row[data-id="w3"]`);
  ok(await b.waitFor(`!!document.querySelector('.scu-overlay .scu-mail')`, 5000), 'UI13 點列 → 開詳情（唯讀）');
  const dt = await b.ev(`document.querySelector('.scu-overlay').innerText`);
  ok(/ADDR-C路3號/.test(dt) && /（公司登記地址，未另填郵寄地址）/.test(dt) && !/(NOTE_PII_|MEMO_PII_|HOBBY_PII_|CARD_PII_|EVIL_PII_|@pii\.example\.test)/.test(dt), 'UI14 詳情：郵寄地址高亮區含 fallback 註記、無 PII', dt.slice(0, 200));
  ok((dt.match(/（個資已遮蔽）/g) || []).length === 2 && /TEL_PII_w3 \(分機 EXT_PII_w3\)/.test(dt) && /MOB_PII_w3/.test(dt), 'UI15 詳情：電話（含分機）／手機顯示，Email／備註仍為「（個資已遮蔽）」');
  eq(await b.ev(`document.querySelectorAll('.scu-overlay input, .scu-overlay textarea, .scu-overlay select').length`), 0, 'UI16 詳情沒有任何可編輯欄位（唯讀）');
  await shot(b, C, 'scu_detail_1440_light.png');
  await b.click('.scu-overlay .scu-copy');
  ok(await b.waitFor(`/已複製/.test(document.querySelector('.scu-overlay .scu-copied').textContent)`, 4000), 'UI17 點「複製郵寄地址」→ 顯示「已複製」', await b.ev(`document.querySelector('.scu-overlay .scu-copied').textContent`));
  await b.key('Escape');
  ok(await b.waitFor(`!document.querySelector('.scu-overlay')`, 3000), 'UI18 Esc 關閉詳情');
  eq(await b.ev(`document.activeElement && document.activeElement.dataset && document.activeElement.dataset.id`), 'w3', 'UI19 關閉後焦點回到該列');
  // 沒有郵寄地址的那筆：複製鈕停用
  await b.click(`#secretaryCustomersView .scu-row[data-id="w4"]`);
  ok(await b.waitFor(`!!document.querySelector('.scu-overlay .scu-mail')`, 5000), 'UI20 開 w4 詳情');
  eq(await b.ev(`document.querySelector('.scu-overlay .scu-copy').disabled`), true, 'UI21 沒有地址 → 複製鈕停用');
  ok(/（未填）/.test(await b.ev(`document.querySelector('.scu-overlay .scu-mail').innerText`)), 'UI21b 顯示（未填）');
  await b.key('Escape');
  // 切到別的區段 → 本區段隱藏；回首頁也隱藏
  await b.click('#navQuotations');
  ok(await b.waitFor(`!(${VIEW_VIS})`, 4000), 'UI22 切到報價單 → 本區段隱藏');
  eq(await b.ev(`document.getElementById('navSecretaryCustomers').classList.contains('active')`), false, 'UI22b 導覽 active 已移除');
  await openNav(b);
  ok(await b.waitFor(VIEW_VIS, 6000), 'UI23 再點回來 → 顯示');
  await b.ev(`showDashboard(); 0`);
  ok(await b.waitFor(`!(${VIEW_VIS}) && document.getElementById('dashboardView').style.display===''`, 4000), 'UI24 回首頁 → 本區段隱藏');
  // F5 還原
  await openNav(b);
  ok(await b.waitFor(`${ROWS}===4`, 8000), 'UI25 進入本區段');
  await b.nav(C.BASE + '/');
  ok(await b.waitFor(`${VIEW_VIS} && ${ROWS}===4`, 15000), 'UI26 F5 重新整理 → 還原在本區段並載入清單（沒有 loader 未定義例外）');
  eq(await b.ev(`localStorage.getItem('lastSection')`), 'secretaryCustomers', 'UI26b lastSection 為 secretaryCustomers');
  // 稽核
  await C.login('_s_adm');
  const lg = (await C.logsOf()).filter((l) => l.action === 'SECRETARY_VIEW_CUSTOMER');
  ok(lg.length >= 2 && lg.every((l) => !/_PII_|@pii/.test(JSON.stringify(l))), 'UI27 開啟詳情有寫 SECRETARY_VIEW_CUSTOMER 稽核（不含 PII）', lg.length);
}

// 首次載入競態：secretary-customers.js 下載很慢（延遲 2.5 秒）時，入口仍必須出現（不需 F5）
async function uiLate(b, C) {
  console.log('  [UI] 腳本晚到：入口仍出現（不需 F5）');
  const { ok } = C;
  await b.interceptOn();
  b.delay['/secretary-customers.js'] = 2500;
  try {
    await openAs(b, C, '_s_sec', 1440, 900);
    ok(await b.waitFor(`!!document.getElementById('navSecretaryCustomers')`, 12000), 'UI70 secretary-customers.js 延遲 2.5 秒載入 → 首次載入側欄仍出現「我的客戶（唯讀）」（沒有 F5）');
  } finally { delete b.delay['/secretary-customers.js']; }
}

async function uiOthers(b, C) {
  console.log('  [UI] 非秘書／殘留的 lastSection 不建立元素、不出現空白頁');
  const { ok, eq } = C;
  await openAs(b, C, '_s_u1', 1440, 900, { last: 'secretaryCustomers' });
  ok(await b.waitFor(`document.getElementById('dashboardView').style.display===''`, 10000), 'UI30 業務帶著殘留 lastSection=secretaryCustomers 登入 → 退回首頁（不是空白頁）');
  eq(await b.ev(`[!!document.getElementById('navSecretaryCustomers'), !!document.getElementById('secretaryCustomersView')]`), [false, false], 'UI31 業務：側欄項與區段元素都沒有建立');
}

async function uiXss(b, C) {
  console.log('  [UI] XSS');
  const { ok, eq } = C;
  await openAs(b, C, '_s_sec', 1440, 1000);
  await openNav(b);
  ok(await b.waitFor(`${ROWS}===4`, 8000), 'UI40 載入');
  eq(await b.ev(`window.__xss||0`), 0, 'UI41 含 <img onerror> 的郵寄地址：酬載沒被執行');
  eq(await b.ev(`document.querySelectorAll('#secretaryCustomersView img').length`), 0, 'UI42 沒有變成 <img> 元素');
  ok((await b.ev(VIEW_TEXT)).includes('<img src=x'), 'UI43 以純文字顯示');
  await b.click(`#secretaryCustomersView .scu-row[data-id="w1"]`);
  ok(await b.waitFor(`!!document.querySelector('.scu-overlay .scu-mail')`, 5000), 'UI44 開詳情');
  eq([await b.ev(`window.__xss||0`), await b.ev(`document.querySelectorAll('.scu-overlay img').length`)], [0, 0], 'UI45 詳情也不執行酬載');
  await b.key('Escape');
}

async function uiLayout(b, C) {
  console.log('  [UI] 手機寬度／深色／玻璃');
  const { ok } = C;
  const probe = `(()=>{const v=document.getElementById('secretaryCustomersView'),vw=innerWidth;const rows=[...v.querySelectorAll('.scu-row')];const ov=document.querySelector('.scu-box');return {vw,docW:document.documentElement.scrollWidth,viewR:v.getBoundingClientRect().right,rowsR:Math.max(...rows.map(r=>r.getBoundingClientRect().right)),boxR:ov?ov.getBoundingClientRect().right:0,boxB:ov?ov.getBoundingClientRect().bottom:0,vh:innerHeight,fg:getComputedStyle(rows[0]).color,bg:getComputedStyle(rows[0]).backgroundColor}})()`;
  for (const [w, h] of [[375, 800], [1440, 900]]) {
    await openAs(b, C, '_s_sec', w, h);
    await b.ev(`showSection('secretaryCustomers'); 0`);
    ok(await b.waitFor(`${ROWS}===4`, 8000), `UI50 ${w}px 載入`);
    for (const mode of ['plain-light', 'plain-dark', 'glass-dark', 'glass-light']) {
      await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass',${J(mode.startsWith('glass'))}); c.toggle('dark',${J(mode.endsWith('dark'))}); return true})()`);
      await b.sleep(200);
      let p = await b.ev(probe);
      ok(p.docW <= p.vw + 1 && p.rowsR <= p.vw + 1, `UI51 ${w}px ${mode}：清單沒有橫向捲動／超出視窗`, J(p));
      if (mode.startsWith('plain')) ok(p.fg !== p.bg, `UI52 ${w}px ${mode}：文字與背景顏色不同（可讀）`, J(p));
      await shot(b, C, `scu_list_${w}_${mode}.png`);
      await b.click(`#secretaryCustomersView .scu-row[data-id="w3"]`);
      ok(await b.waitFor(`!!document.querySelector('.scu-overlay .scu-mail')`, 5000), `UI53 ${w}px ${mode}：開詳情`);
      p = await b.ev(probe);
      ok(p.boxR <= p.vw + 1 && p.boxB <= p.vh + 1, `UI54 ${w}px ${mode}：詳情視窗在畫面內`, J(p));
      await shot(b, C, `scu_detail_${w}_${mode}.png`);
      await b.key('Escape');
      await b.waitFor(`!document.querySelector('.scu-overlay')`, 2000);
    }
    await b.ev(`document.body.classList.remove('glass','dark'); 0`);
  }
  await b.viewport(1440, 1000, false);
}

async function uiOff(b, C) {
  console.log('  [UI] 開關關閉：不建立任何元素');
  const { ok, eq } = C;
  await openAs(b, C, '_s_sec', 1440, 900, { last: 'secretaryCustomers' });
  ok(await b.waitFor(`document.getElementById('dashboardView').style.display===''`, 10000), 'UI60 開關關：秘書帶殘留 lastSection 登入 → 退回首頁');
  await b.sleep(800);
  eq(await b.ev(`[!!document.getElementById('navSecretaryCustomers'), !!document.getElementById('secretaryCustomersView'), !!document.querySelector('[class*="scu-"]')]`), [false, false, false], 'UI61 開關關：側欄項、區段、任何 .scu-* 元素都沒有建立');
}

async function runOnce(C, mut, which) {
  const fns = { flow: uiFlow, others: uiOthers, xss: uiXss, layout: uiLayout, late: uiLate };
  const run = (mut && mut.sections) || which || ['flow', 'others', 'xss', 'layout', 'late', 'off'];
  const sections = [];
  const server = async (env, seed) => { await C.freshServer({ env, seed }); };
  await server({ SECRETARY_CUSTOMERS: '1' }, XSS_SEED);
  await withBrowser({ cdpPort: C.CDP, maxMs: 560000 }, async (b) => {
    if (mut) { b.replace[mut.urlTail] = mut.mutated; for (const x of mut.extra || []) b.replace[x.urlTail] = x.mutated; await b.interceptOn(); }
    const guard = async (s, fn) => {
      try { await fn(b, C); } catch (e) { C.ok(false, `UI 區段 ${s} 例外：${e.message}`); }
      if (b.errors.length) { console.log('    頁面錯誤：', b.errors.slice(0, 3).join(' | ')); C.ok(false, `UI 區段 ${s}：頁面沒有 JS 例外／console.error`, b.errors[0]); b.errors.length = 0; }
    };
    for (const s of run) {
      sections.push(s);
      if (s === 'off') continue;
      await guard(s, fns[s]);
    }
  });
  if (run.includes('off')) {
    await C.stopServer();
    await server({}, null);
    await withBrowser({ cdpPort: C.CDP, maxMs: 200000 }, async (b) => {
      if (mut) { b.replace[mut.urlTail] = mut.mutated; await b.interceptOn(); }
      try { await uiOff(b, C); } catch (e) { C.ok(false, `UI 區段 off 例外：${e.message}`); }
      if (b.errors.length) { C.ok(false, 'UI 區段 off：頁面沒有 JS 例外', b.errors[0]); b.errors.length = 0; }
    });
  }
  console.log('  Edge 殘留程序：', withBrowser.leftover || 0);
  C.ok((withBrowser.leftover || 0) === 0, '測試結束沒有殘留的 Edge 程序');
  await C.stopServer(); C.cleanupFiles();
}
async function run(C) { console.log('[前端] headless Edge + CDP'); await runOnce(C, null); }

async function mutants(C) {
  console.log('[變異·前端] 把前端檔換成改壞的版本，對應斷言必須變紅');
  const read = (rel) => fs.readFileSync(path.join(C.ROOT, rel), 'utf8');
  const defs = [
    { name: 'UM1 探測失敗也建立元素（開關關仍建側欄／區段）', file: '_client/secretary-customers.js', urlTail: '/secretary-customers.js', find: "if (!rows) return false;\n      available = true;", rep: 'rows = rows || [];\n      available = true;', sections: ['off'] },
    { name: 'UM2 切換區段不隱藏本區段（拿掉 showSection 的 setActive(false)）', file: '_client/app.js', urlTail: '/app.js', find: "  if (window.ITTSSecCust) window.ITTSSecCust.setActive(false);\n  if (section === 'secretaryCustomers')", rep: "  if (section === 'secretaryCustomers')", sections: ['flow'] },
    { name: 'UM3 fallback 不加註', file: '_client/secretary-customers.js', urlTail: '/secretary-customers.js', find: 'if (t && c.mailingFallback) wrap.appendChild', rep: 'if (false) wrap.appendChild', sections: ['flow'] },
    { name: 'UM4 地址以 innerHTML 顯示（XSS）', file: '_client/secretary-customers.js', urlTail: '/secretary-customers.js', find: "wrap.appendChild(el('span', t ? 'scu-addr-text' : 'scu-muted', t || EMPTY));", rep: "{ const sp = el('span', t ? 'scu-addr-text' : 'scu-muted'); sp.innerHTML = t || EMPTY; wrap.appendChild(sp); }", sections: ['xss'] },
    { name: 'UM5 F5 還原不處理（拿掉 _whenScriptsReady 的 show 呼叫）', file: '_client/app.js', urlTail: '/app.js', find: "if (section === 'secretaryCustomers') _whenScriptsReady(section, () => { if (window.ITTSSecCust) window.ITTSSecCust.show(); else showDashboard(); });", rep: '', sections: ['flow'] },
    { name: 'UM6 載入順序還原（腳本在 app.js 之後）＋拿掉補呼叫 → 腳本晚到時入口不出現', file: '_client/secretary-customers.js', urlTail: '/secretary-customers.js', find: "if (window._myRole === 'secretary') window.ITTSSecCust.start({ role: window._myRole });", rep: '', sections: ['late'],
      extraFn: (read) => { const NL = String.fromCharCode(10); let h = read('_client/index.html'); const a = h.indexOf('  <script src="secretary-customers.js">'), a2 = h.indexOf(NL, a) + 1; const t = h.slice(a, a2); h = h.slice(0, a) + h.slice(a2); const b = h.indexOf('  <script src="app.js">'), b2 = h.indexOf(NL, b) + 1; h = h.slice(0, b2) + t + h.slice(b2); return [{ urlTail: '/', mutated: h }]; } },
  ];
  for (const m of defs) {
    if (process.env.SCU_MUT_PICK && !new RegExp(process.env.SCU_MUT_PICK).test(m.name)) continue;
    const src = read(m.file);
    if (!src.includes(m.find)) { C.ok(false, `${m.name}：找不到變異錨點`); continue; }
    const mutated = src.replace(m.find, () => m.rep);
    const snap = C.snap();
    C.setQuiet(true);
    try { await runOnce(C, { urlTail: m.urlTail, mutated, sections: m.sections, extra: m.extraFn ? m.extraFn(read) : [] }); } catch (e) { C.ok(false, '變異執行例外：' + e.message); }
    C.setQuiet(false);
    const fs2 = C.newFails(snap);
    C.restore(snap);
    console.log('  ' + (fs2.length > 0 ? '✓ 變紅' : '✗ 沒變紅') + '：' + m.name + '（' + fs2.length + ' 項）');
    C.ok(fs2.length > 0, `${m.name} → 測試變紅（${fs2.length} 項失敗：${fs2.slice(0, 2).join(' / ').slice(0, 80)}）`);
  }
}
module.exports = { run, mutants };
