#!/usr/bin/env node
/**
 * 簽核流程設計器 M2 — 前端互動檢查（headless Edge ＋ CDP，真實滑鼠／鍵盤事件）＋ 變異測試。
 * 用法：node scripts/check-quote-flow-designer-ui.js        （全部；埠 FD_PORT 預設 3130、CDP 埠 FD_CDP 預設 9461）
 *       FD_UI_ONLY=<群組名,…>  只跑指定群組（render,drag,drawer,keyboard,save,sim,val,ov,xss,mobile,theme,motion,gate,static）
 *       FD_UI_MUT=1            另外跑變異測試（把前端檔換成改壞的版本，對應群組必須變紅）
 *       FD_SHOTS=<資料夾>      另外輸出截圖（1440 深／淺、總覽、模擬、錯誤、375 手機）
 * 沙盒：臨時 _s_* 帳號與資料檔只存在於本次執行，結束以 PID 清掉伺服器與 Edge（以臨時 profile 路徑辨識）；不連遠端資料庫。
 * 涵蓋原型 152 項互動測試中的核心項目：拖拉（卡片排序／換路線／條件卡／路線優先順序）、雙擊開抽屜、條件造句編輯器、底線即時拒絕、
 * 鍵盤操作、刪除確認、儲存／重載／樂觀並行、模擬器（高亮／比較／邊界掃描）、驗證面板定位、總覽（縮放／平移／高亮）、XSS、手機唯讀、
 * 四種外觀（淺／深／玻璃淺／玻璃深）× 六主題的對比、prefers-reduced-motion、開關關閉、非管理員、無 inline script／on*／style=。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const fx = require('./_fd-fixture');
const { withBrowser } = require('./_fd-cdp');

const ROOT = fx.ROOT;
const CDP = parseInt(process.env.FD_CDP || '9461', 10);
const ONLY = (process.env.FD_UI_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const SHOTS = process.env.FD_SHOTS || '';
fx.guard();

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = fx.sleep;

// ───────────────────────── 頁面操作小工具 ─────────────────────────
const J = async (b, expr) => JSON.parse(await b.ev(`JSON.stringify(${expr})`));
const settle = (b) => b.ev('FlowDesigner.state().queue.then(()=>true)');
const defCanon = (b) => b.ev('FlowDesigner._test.canonStr(FlowDesigner.state().def)');
const lanePath = (b, id) => J(b, `Array.from(document.querySelectorAll('#fdLanes .fd-node')).filter(e=>e.dataset.lane===${JSON.stringify(id)}).map(e=>e.dataset.tier)`);
const laneIds = (b) => J(b, `Array.from(document.querySelectorAll('#fdLanes .fd-lane')).map(e=>e.dataset.lane)`);
const text = (b, sel) => b.ev(`(()=>{const e=document.querySelector(${JSON.stringify(sel)});return e?e.textContent.replace(/\\s+/g,' ').trim():null})()`);
const exists = (b, sel) => b.ev(`!!document.querySelector(${JSON.stringify(sel)})`);
const count = (b, sel) => b.ev(`document.querySelectorAll(${JSON.stringify(sel)}).length`);
const visible = (b, sel) => b.ev(`(()=>{const e=document.querySelector(${JSON.stringify(sel)});if(!e)return false;const r=e.getBoundingClientRect();return r.width>0&&r.height>0&&getComputedStyle(e).visibility!=='hidden'&&!e.closest('[hidden]')})()`);
const bannerText = (b) => b.ev(`(()=>{const e=document.querySelector('#fdBanner');return e&&!e.hidden?e.textContent.replace(/\\s+/g,' ').trim():''})()`);
const dlgOpen = (b) => b.ev(`!!document.querySelector('#fdDlg[open]')`);
const dlgText = (b) => b.ev(`(()=>{const e=document.querySelector('#fdDlg[open]');return e?e.textContent.replace(/\\s+/g,' ').trim():''})()`);
async function dlgClick(b, v) { await b.click(`#fdDlg[open] [data-close="${v}"]`); await sleep(150); }
const lane = (id) => `#fdLanes .fd-lane[data-lane="${id}"]`;
const nodeSel = (l, i) => `#fdLanes .fd-node[data-lane="${l}"][data-idx="${i}"]`;
async function tab(b, t) { await b.click(`#fdTabBtn-${t}`); await sleep(250); }

async function login(b, who) {
  await b.nav(fx.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${JSON.stringify(who)},password:${JSON.stringify(fx.PW)}})}).then(r=>r.status)`);
  return st === 200;
}
/** 開後台、等探測完成、點側欄項、等草稿載入 */
async function openDesigner(b, w, h) {
  if (w) await b.viewport(w, h || 1000, w < 700);
  await b.nav(fx.BASE + '/admin.html');
  const probed = await b.waitFor('window.FlowDesigner && FlowDesigner.state().enabled === true', 25000);
  if (!probed) return false;
  await b.ev(`document.querySelector('.admin-nav-item[data-sec="flow-designer"]').click()`);
  return b.waitFor('FlowDesigner.state().inited && !!FlowDesigner.state().def && !!document.querySelector("#fdLanes .fd-lane")', 20000);
}

// 對比：取元素的文字色與「實際底色」（逐層疊合半透明背景；到頂仍半透明時疊在 backdrop 上）
const CONTRAST_FN = `(function(backdrop){
  const parse=c=>{const m=/rgba?\\(([^)]+)\\)/.exec(c);if(!m)return null;const p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number);return {r:p[0],g:p[1],b:p[2],a:p[3]===undefined?1:p[3]};};
  const over=(top,bot)=>({r:top.r*top.a+bot.r*(1-top.a),g:top.g*top.a+bot.g*(1-top.a),b:top.b*top.a+bot.b*(1-top.a),a:1});
  const lum=c=>{const f=v=>{v/=255;return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)};return 0.2126*f(c.r)+0.7152*f(c.g)+0.0722*f(c.b)};
  const ratio=(a,b)=>{const l1=lum(a),l2=lum(b);return (Math.max(l1,l2)+0.05)/(Math.min(l1,l2)+0.05)};
  const bgOf=el=>{const layers=[];let e=el;while(e&&e.nodeType===1){const c=parse(getComputedStyle(e).backgroundColor);if(c&&c.a>0){layers.push(c);if(c.a>=0.999)break;}e=e.parentElement;}
    let base=layers.length&&layers[layers.length-1].a>=0.999?layers.pop():parse(backdrop);let cur=base;while(layers.length){cur=over(layers.pop(),cur);}return cur;};
  return function(sel){const out=[];document.querySelectorAll(sel).forEach((el,i)=>{if(i>2)return;const r=el.getBoundingClientRect();if(!r.width||!r.height)return;const fg=parse(getComputedStyle(el).color);const bg=bgOf(el);const fgc=fg.a<1?over(fg,bg):fg;out.push({sel,ratio:+ratio(fgc,bg).toFixed(2),text:el.textContent.trim().slice(0,20)});});return out;};
})`;

// ═════════════════════════ 群組 ═════════════════════════
const G = {};

// ── render：預設流程的渲染、M3 灰色說明、沒有「生效」按鈕
G.render = async (b) => {
  console.log('  [render]');
  eq(await laneIds(b), ['r_board', 'r_l3', 'r_l2', 'r_default'], 'render 4 條路線依優先順序');
  eq(await lanePath(b, 'r_board'), ['mgr1', 'gm', 'board'], 'render 董事會路線關卡');
  eq(await lanePath(b, 'r_l3'), ['mgr1', 'gm', 'chairman'], 'render 董事長路線關卡');
  eq(await lanePath(b, 'r_default'), ['mgr1'], 'render 預設路線關卡');
  ok(await b.ev(`document.querySelector(${JSON.stringify(lane('r_default'))}).classList.contains('default')`), 'render 預設路線標記 default');
  eq(await text(b, lane('r_default') + ' .fd-cond'), '其他情況（預設）', 'render 預設路線條件文字');
  ok((await text(b, lane('r_board') + ' .fd-cond')).includes('5,000 萬'), 'render 董事會路線條件含 5,000 萬');
  ok((await text(b, lane('r_l3') + ' .fd-cond')).includes('1,000 萬') && (await text(b, lane('r_l3') + ' .fd-cond')).includes('毛利層級'), 'render 董事長路線條件含 1,000 萬與毛利層級');
  ok(await exists(b, nodeSel('r_board', 2) + '.locked'), 'render 董事會代核節點有鎖定樣式');
  eq(await count(b, '#fdPalette .fd-pal[data-pal="node"]'), 5, 'render 面板有 5 種關卡節點');
  eq(await count(b, '#fdPalette .fd-pal[data-pal="cond"]'), 7, 'render 面板有 7 種條件卡（M1 白名單變數）');
  eq(await count(b, '.fd-mod'), 2, 'render 2 個調整規則');
  ok((await text(b, '.fd-mod')).includes('ITS') && (await text(b, '.fd-mod')).includes('MDM'), 'render BU 調整規則顯示 ITS、MDM');
  ok(await b.ev(`document.querySelector('#fdPropose').disabled === true`), 'render「提交覆核（M3 開放）」是停用的灰色按鈕');
  ok(!(await b.ev(`Array.from(document.querySelectorAll('#fdApp button')).some(x=>!x.disabled && /生效|提議|核可|發布/.test(x.textContent))`)), 'render 沒有任何可按的「生效／提議／核可」按鈕（M2 草稿不能生效）');
  ok((await text(b, '#fdSavedPill')).includes('尚未儲存'), 'render 新草稿狀態「尚未儲存」');
  ok((await text(b, '#fdValPill')).includes('警告'), 'render 驗證徽章（二級主管解析警告）');
  eq(await b.ev(`FlowDesigner.state().def.origin`), 'designer', 'render 本地草稿 origin=designer');
  eq(await b.ev(`document.querySelectorAll('#sec-flow-designer button').length > 10`), true, 'render 按鈕有渲染');
};

// ── drag：真實滑鼠拖拉
G.drag = async (b) => {
  console.log('  [drag]');
  await b.viewport(1440, 2200);
  await sleep(300);
  // 面板節點拖進預設路線
  let before = await defCanon(b);
  await b.drag('.fd-pal[data-drag="pnode:gm"]', lane('r_default') + ' .fd-slot');
  await settle(b); await sleep(200);
  eq(await lanePath(b, 'r_default'), ['mgr1', 'gm'], 'drag 面板「總經理」拖進預設路線');
  ok((await defCanon(b)) !== before, 'drag 草稿內容已更新');
  ok((await text(b, '#fdSavedPill')).includes('尚未儲存'), 'drag 狀態仍為「尚未儲存」');
  // 同一個關卡不能出現兩次
  await b.drag('.fd-pal[data-drag="pnode:gm"]', lane('r_default') + ' .fd-slot');
  await settle(b); await sleep(200);
  eq(await lanePath(b, 'r_default'), ['mgr1', 'gm'], 'drag 重複拖入同一個關卡被拒（路線不變）');
  ok((await bannerText(b)).includes('不能出現兩次'), 'drag 重複關卡的拒絕訊息', await bannerText(b));
  // 卡片排序：把 gm 拖到 mgr1 前面
  await b.drag(nodeSel('r_default', 1), nodeSel('r_default', 0), { dx: -50 });
  await settle(b); await sleep(200);
  eq(await lanePath(b, 'r_default'), ['gm', 'mgr1'], 'drag 卡片排序：總經理拖到一級主管前面');
  ok((await b.ev(`FlowDesigner.state().checks.find(c=>c.n===2).status`)) === 'warn', 'drag 順序由高到低 → 驗證第 2 條警告（不擋）');
  // 底線：把董事長從董事長路線拖到總經理路線 → 即時拒絕
  const c0 = await defCanon(b);
  const rej0 = await b.ev('FlowDesigner.state().rejects');
  await b.drag(nodeSel('r_l3', 2), lane('r_l2') + ' .fd-path');
  await settle(b); await sleep(300);
  eq(await lanePath(b, 'r_l3'), ['mgr1', 'gm', 'chairman'], 'drag 拖走董事長（會違反底線）被即時拒絕：路線不變');
  eq(await defCanon(b), c0, 'drag 被拒絕後草稿內容完全不變');
  ok((await bannerText(b)).includes('保護底線') && (await bannerText(b)).includes('董事長'), 'drag 拒絕訊息說明保護底線與缺少的關卡', await bannerText(b));
  eq(await b.ev('FlowDesigner.state().rejects'), rej0 + 1, 'drag 拒絕次數 +1');
  ok((await text(b, '#fdRejPill')).includes('底線拒絕'), 'drag 狀態列出現「底線拒絕」徽章');
  // 新增路線＋用 ▲▼ 與拖曳調整路線優先順序
  await b.click('#fdAddLane'); await settle(b); await sleep(250);
  let ids = await laneIds(b);
  eq(ids.length, 5, 'drag 新增路線後有 5 條');
  eq(ids[4], 'r_default', 'drag 預設路線仍固定在最後');
  const nid = ids[3];
  await b.click(`${lane(nid)} [data-act="lane-up"]`); await settle(b); await sleep(250);
  eq((await laneIds(b))[2], nid, 'drag ▲ 把新路線上移一格（在董事長路線之後，合法）');
  const c1 = await defCanon(b);
  await b.click(`${lane(nid)} [data-act="lane-up"]`); await settle(b); await sleep(300);
  eq(await defCanon(b), c1, 'drag 再上移到董事長路線前面（MDM 大單會少董事長）被拒絕');
  ok((await bannerText(b)).includes('保護底線'), 'drag 路線上移被拒的訊息是保護底線');
  // 路線拖曳（grip）
  await b.drag(`${lane(nid)} .fd-grip`, lane('r_l2'), { dy: 40 });
  await settle(b); await sleep(250);
  const ids2 = await laneIds(b);
  ok(ids2.indexOf(nid) > ids2.indexOf('r_l2') && ids2[ids2.length - 1] === 'r_default', 'drag 拖曳路線把手：放到「總經理路線」之後', JSON.stringify(ids2));
  // 條件卡拖到條件膠囊
  const itemsBefore = await b.ev(`FlowDesigner._test.astToEdit(FlowDesigner.state().def.routes.find(r=>r.id===${JSON.stringify(nid)}).when).items.length`);
  await b.drag('.fd-pal[data-drag="pcond:amount"]', `${lane(nid)} .fd-cond`);
  await settle(b); await sleep(300);
  const itemsAfter = await b.ev(`FlowDesigner._test.astToEdit(FlowDesigner.state().def.routes.find(r=>r.id===${JSON.stringify(nid)}).when).items.length`);
  eq(itemsAfter, itemsBefore + 1, 'drag 條件卡拖到條件膠囊：路線多一個條件');
  ok(await visible(b, '#fdDrawer'), 'drag 加入條件後自動開啟右側抽屜');
  await b.click('#fdDrawer [data-act="dr-close"]');
  // 條件卡拖到預設路線：不行
  const c2 = await defCanon(b);
  await b.drag('.fd-pal[data-drag="pcond:category"]', `${lane('r_default')} .fd-cond`);
  await sleep(300);
  eq(await defCanon(b), c2, 'drag 條件卡拖到預設路線被拒（預設路線不能加條件）');
  // 點一下面板（不拖）→ 對話框選路線
  await b.click('.fd-pal[data-pal="node"][data-tier="chairman"]');
  await sleep(300);
  ok(await dlgOpen(b), 'drag 單擊面板節點開啟「加入哪條路線」對話框');
  await dlgClick(b, nid); await settle(b); await sleep(250);
  ok((await lanePath(b, nid)).includes('chairman'), 'drag 對話框選路線後節點加入該路線');
  await b.viewport(1440, 1000);
};

// ── drawer：雙擊開屬性、條件造句編輯器
G.drawer = async (b) => {
  console.log('  [drawer]');
  await b.viewport(1440, 2200); await sleep(250);
  await b.dblclick(nodeSel('r_l2', 1));
  await sleep(250);
  ok(await visible(b, '#fdDrawer') && (await text(b, '#fdDrawerH')).includes('關卡屬性'), 'drawer 雙擊關卡開啟屬性抽屜');
  await b.ev(`(()=>{const i=document.querySelector('#fdDrLabel');i.value='總經理核決';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  await b.click('#fdDrawer [data-color="amber"]');
  await b.click('#fdDrawer [data-act="dr-apply"]'); await settle(b); await sleep(300);
  eq(await b.ev(`FlowDesigner.state().def.nodes.find(n=>n.tier==='gm').label`), '總經理核決', 'drawer 套用顯示名稱');
  eq(await b.ev(`FlowDesigner.state().def.nodes.find(n=>n.tier==='gm').color`), 'amber', 'drawer 套用節點色');
  ok(!(await visible(b, '#fdDrawer')), 'drawer 套用後抽屜關閉');
  ok((await text(b, nodeSel('r_l2', 1))).includes('總經理核決'), 'drawer 畫布上所有「總經理」卡片同步改名');
  // 不合法名稱
  await b.dblclick(nodeSel('r_l2', 1)); await sleep(250);
  await b.ev(`(()=>{const i=document.querySelector('#fdDrLabel');i.value='<b>x</b>';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  const cc = await defCanon(b);
  await b.click('#fdDrawer [data-act="dr-apply"]'); await sleep(250);
  ok((await text(b, '#fdDrErr')).includes('名稱不合法'), 'drawer 名稱含 < > 被擋（抽屜內錯誤訊息）');
  eq(await defCanon(b), cc, 'drawer 不合法名稱不改草稿');
  await b.click('#fdDrawer [data-act="dr-close"]'); await sleep(200);
  // 鎖定節點
  await b.dblclick(nodeSel('r_board', 2)); await sleep(250);
  ok(await exists(b, '#fdDrawer .fd-lockbox') && await b.ev(`document.querySelector('#fdDrLabel').disabled`) && await b.ev(`document.querySelector('#fdDrawer [data-act="dr-apply"]').disabled`), 'drawer 鎖定節點：顯示鎖定說明、名稱欄與套用鈕停用');
  await b.click('#fdDrawer [data-act="dr-close"]'); await sleep(200);
  // 路線條件編輯：把 1,000 萬改成 2,000 萬 → 底線擋下；改成 800 萬 → 通過
  await b.dblclick(`${lane('r_l3')} .fd-cond`); await sleep(300);
  ok((await text(b, '#fdDrawerH')).includes('路線條件編輯器'), 'drawer 雙擊條件膠囊開啟條件編輯器');
  ok((await b.ev(`document.querySelectorAll('#fdBlocks .fd-blk').length`)) === 2 && (await text(b, '#fdDrawer .fd-joiner')).includes('條件之間'), 'drawer 「任一符合」兩個條件積木＋連接詞');
  ok((await text(b, '#fdDrSentence')).includes('如果') && (await text(b, '#fdDrSentence')).includes('1,000 萬'), 'drawer 白話預覽');
  await b.ev(`(()=>{const i=document.querySelector('#fdBlocks .fd-blk[data-i="0"] [data-f="wan"]');i.value='2000';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  ok((await text(b, '#fdDrSentence')).includes('2,000 萬'), 'drawer 白話預覽即時更新');
  await b.click('#fdDrawer [data-act="dr-apply"]'); await settle(b); await sleep(300);
  ok((await text(b, '#fdDrErr')).includes('保護底線'), 'drawer 門檻改成 2,000 萬：抽屜內顯示保護底線拒絕', await text(b, '#fdDrErr'));
  ok(await visible(b, '#fdDrawer'), 'drawer 被拒後抽屜維持開啟讓使用者修正');
  await b.ev(`(()=>{const i=document.querySelector('#fdBlocks .fd-blk[data-i="0"] [data-f="wan"]');i.value='800';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  await b.click('#fdDrawer [data-act="dr-apply"]'); await settle(b); await sleep(300);
  ok((await text(b, lane('r_l3') + ' .fd-cond')).includes('800 萬'), 'drawer 改成 800 萬（變嚴格）通過，條件膠囊更新', await text(b, lane('r_l3') + ' .fd-cond'));
  // 造句：新增條件、換變數與運算子
  await b.dblclick(`${lane('r_l3')} .fd-cond`); await sleep(300);
  await b.click('#fdDrawer [data-act="add-item"]'); await sleep(200);
  ok((await b.ev(`document.querySelectorAll('#fdBlocks .fd-blk').length`)) === 3, 'drawer 新增條件積木');
  await b.ev(`(()=>{const s=document.querySelector('#fdBlocks .fd-blk[data-i="2"] [data-f="v"]');s.value='category';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(200);
  await b.ev(`(()=>{const s=document.querySelector('#fdBlocks .fd-blk[data-i="2"] [data-f="o"]');s.value='in';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(200);
  ok((await b.ev(`document.querySelectorAll('#fdBlocks .fd-blk[data-i="2"] input[type=checkbox]').length`)) === 9, 'drawer 商品歸類「屬於」顯示 9 個核取方塊');
  await b.ev(`(()=>{const c=document.querySelector('#fdBlocks .fd-blk[data-i="2"] input[value="consult"]');c.checked=true;c.dispatchEvent(new Event('change',{bubbles:true}))})()`);
  ok((await text(b, '#fdDrSentence')).includes('顧問服務'), 'drawer 白話預覽含「顧問服務」');
  await b.ev(`(()=>{const s=document.querySelector('#fdDrMode');s.value='all';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(200);
  ok((await text(b, '#fdDrSentence')).includes('且'), 'drawer 連接詞改「且」後預覽更新');
  await b.ev(`(()=>{const s=document.querySelector('#fdDrMode');s.value='any';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(200);
  ok((await text(b, '#fdDrSentence')).includes('或'), 'drawer 連接詞改回「或」後預覽更新（維持較嚴格，不違反底線）');
  await b.click('#fdDrawer [data-act="dr-apply"]'); await settle(b); await sleep(300);
  const w3 = await b.ev(`JSON.stringify(FlowDesigner.state().def.routes.find(r=>r.id==='r_l3').when)`);
  ok(w3.startsWith('{"any":[') && w3.includes('"category","in",["consult"]'), 'drawer 套用後 AST 為 {any:[…category in […]]}', w3);
  // 空清單不合法
  await b.dblclick(`${lane('r_l3')} .fd-cond`); await sleep(300);
  await b.ev(`(()=>{const c=document.querySelector('#fdBlocks .fd-blk[data-i="2"] input[value="consult"]');c.checked=false;c.dispatchEvent(new Event('change',{bubbles:true}))})()`);
  await b.click('#fdDrawer [data-act="dr-apply"]'); await sleep(250);
  ok((await text(b, '#fdDrErr')).includes('至少選一個'), 'drawer「屬於」沒勾任何值被擋');
  await b.click('#fdDrawer [data-act="dr-close"]'); await sleep(200);
  // 路線名稱含 HTML 以純文字顯示
  await b.dblclick(`${lane('r_l2')} .fd-lname`); await sleep(300);
  await b.ev(`(()=>{const i=document.querySelector('#fdDrName');i.value='<svg onload=a()>';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  await b.click('#fdDrawer [data-act="dr-apply"]'); await settle(b); await sleep(300);
  eq(await text(b, lane('r_l2') + ' .fd-lname'), '<svg onload=a()>', 'drawer 路線名稱含 HTML 以純文字顯示');
  eq(await count(b, '#fdLanes svg, #fdLanes img'), 0, 'drawer 路線名稱沒有被當成標籤解析（無 svg／img 元素）');
  // 調整規則 BU
  await b.dblclick('.fd-mod'); await sleep(300);
  ok((await text(b, '#fdDrawerH')).includes('調整規則'), 'drawer 雙擊調整規則開啟屬性');
  await b.click('#fdDrawer input[data-f="bu"][value="CRM"]');
  await b.click('#fdDrawer [data-act="dr-apply"]'); await settle(b); await sleep(300);
  ok((await text(b, '.fd-mod')).includes('CRM'), 'drawer 調整規則勾選 CRM 後畫布顯示');
  eq(await b.ev(`JSON.stringify(FlowDesigner.state().def.modifiers[0].when.all[0].cmp[2])`), JSON.stringify(['CRM', 'ITS', 'MDM']), 'drawer BU 清單以固定排序寫回 AST');
  // 檢視 skip 調整規則為唯讀
  await b.dblclick('.fd-mod[data-mod="m_skip"]'); await sleep(250);
  ok(!(await exists(b, '#fdDrawer [data-act="dr-apply"]')) && (await text(b, '#fdDrawer')).includes('例外名單'), 'drawer skip 調整規則為唯讀，且只顯示人數、不顯示帳號');
  ok(!(await text(b, '#fdDrawer')).includes('_s_m2'), 'drawer 不洩漏例外名單帳號');
  await b.click('#fdDrawer [data-act="dr-close"]');
  await b.viewport(1440, 1000);
};

// ── keyboard
G.keyboard = async (b) => {
  console.log('  [keyboard]');
  await b.viewport(1440, 2200); await sleep(250);
  await b.focus(nodeSel('r_board', 0));
  await b.key('ArrowRight'); await settle(b); await sleep(250);
  eq(await lanePath(b, 'r_board'), ['gm', 'mgr1', 'board'], 'keyboard → 把卡片往後移');
  ok(await b.ev(`document.activeElement && document.activeElement.matches('.fd-node') && document.activeElement.dataset.lane==='r_board' && document.activeElement.dataset.idx==='1'`), 'keyboard 移動後焦點跟著卡片');
  await b.key('ArrowLeft'); await settle(b); await sleep(250);
  eq(await lanePath(b, 'r_board'), ['mgr1', 'gm', 'board'], 'keyboard ← 把卡片往前移');
  await b.key('Enter'); await sleep(250);
  ok(await visible(b, '#fdDrawer'), 'keyboard Enter 開啟屬性抽屜');
  await b.key('Escape'); await sleep(250);
  ok(!(await visible(b, '#fdDrawer')), 'keyboard Esc 關閉抽屜');
  ok(await b.ev(`document.activeElement && document.activeElement.matches('.fd-node')`), 'keyboard 抽屜關閉後焦點回到卡片');
  // Delete → 確認對話框 → 取消
  await b.drag('.fd-pal[data-drag="pnode:gm"]', lane('r_default') + ' .fd-slot'); await settle(b); await sleep(250);
  await b.focus(nodeSel('r_default', 1));
  const c0 = await defCanon(b);
  await b.key('Delete'); await sleep(300);
  ok(await dlgOpen(b) && (await dlgText(b)).includes('移除關卡'), 'keyboard Delete 先跳出確認對話框');
  await dlgClick(b, 'cancel'); await settle(b);
  eq(await defCanon(b), c0, 'keyboard 取消確認後內容不變');
  await b.focus(nodeSel('r_default', 1));
  await b.key('Delete'); await sleep(300);
  await dlgClick(b, 'ok'); await settle(b); await sleep(250);
  eq(await lanePath(b, 'r_default'), ['mgr1'], 'keyboard 確認後移除關卡');
  // 刪除會違反底線的關卡
  await b.focus(nodeSel('r_l3', 2));
  const c1 = await defCanon(b);
  await b.key('Delete'); await sleep(300); await dlgClick(b, 'ok'); await settle(b); await sleep(300);
  eq(await defCanon(b), c1, 'keyboard Delete 移除董事長（違反底線）：確認後仍被拒絕');
  ok((await bannerText(b)).includes('保護底線'), 'keyboard 被拒訊息為保護底線');
  // grip 上下鍵移動路線（先新增一條）
  await b.click('#fdAddLane'); await settle(b); await sleep(250);
  const ids = await laneIds(b); const nid = ids[3];
  await b.focus(`${lane(nid)} .fd-grip`);
  await b.key('ArrowUp'); await settle(b); await sleep(250);
  eq((await laneIds(b))[2], nid, 'keyboard 把手 ↑ 上移路線');
  ok(await b.ev(`document.activeElement && document.activeElement.classList.contains('fd-grip')`), 'keyboard 路線移動後焦點留在把手');
  // 面板 Enter 開對話框
  await b.focus('.fd-pal[data-pal="cond"][data-var="marginLevel"]');
  await b.key('Enter'); await sleep(300);
  ok(await dlgOpen(b) && (await dlgText(b)).includes('毛利層級'), 'keyboard 面板條件卡 Enter 開對話框');
  await b.key('Escape'); await sleep(250);
  ok(!(await dlgOpen(b)), 'keyboard Esc 關閉對話框');
  // 條件膠囊 Enter
  await b.focus(`${lane('r_l3')} .fd-cond`);
  await b.key('Enter'); await sleep(300);
  ok((await text(b, '#fdDrawerH')).includes('條件編輯器'), 'keyboard 條件膠囊 Enter 開條件編輯器');
  await b.key('Escape'); await sleep(200);
  // 分頁左右鍵
  await b.focus('#fdTabBtn-design');
  await b.key('ArrowRight'); await sleep(250);
  eq(await b.ev(`FlowDesigner.state().tab`), 'sim', 'keyboard 分頁 → 切到模擬器');
  ok(await b.ev(`document.activeElement.id==='fdTabBtn-sim' && document.querySelector('#fdTabBtn-sim').getAttribute('aria-selected')==='true'`), 'keyboard 分頁焦點與 aria-selected 同步');
  await b.key('ArrowLeft'); await sleep(250);
  eq(await b.ev(`FlowDesigner.state().tab`), 'design', 'keyboard 分頁 ← 回設計器');
  // aria：lane／node 有標籤
  ok(await b.ev(`Array.from(document.querySelectorAll('#fdLanes .fd-node')).every(e=>e.getAttribute('aria-label') && e.tabIndex===0)`), 'keyboard 所有關卡卡片可聚焦且有 aria-label');
  ok(await b.ev(`!!document.querySelector('#fdLive[aria-live="polite"]') && document.querySelector('#fdLive').textContent.length>0`), 'keyboard aria-live 區有播報內容');
  await b.viewport(1440, 1000);
};

// ── save：儲存／重載／樂觀並行／刪除
G.save = async (b) => {
  console.log('  [save]');
  const api = (m, u, body) => b.ev(`fetch('/api/admin/quote-flow${u}',{method:'${m}',headers:{'Content-Type':'application/json'}${body ? `,body:JSON.stringify(${JSON.stringify(body)})` : ''}}).then(async r=>({status:r.status,j:await r.json()}))`);
  await b.viewport(1440, 2200); await sleep(250);
  // 沒有名稱不能存
  await b.ev(`(()=>{const i=document.querySelector('#fdName');i.value='';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  await b.click('#fdSave'); await sleep(300);
  eq((await api('GET', '/drafts')).j.items.length, 0, 'save 沒填名稱不送出');
  // 改點東西 → 存
  await b.drag('.fd-pal[data-drag="pnode:gm"]', lane('r_default') + ' .fd-slot'); await settle(b); await sleep(250);
  await b.ev(`(()=>{const i=document.querySelector('#fdName');i.value='我的草稿';i.dispatchEvent(new Event('input',{bubbles:true}));const n=document.querySelector('#fdNote');n.value='測試備註';n.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  await b.click('#fdSave'); await sleep(900);
  const lst = (await api('GET', '/drafts')).j;
  eq(lst.items.length, 1, 'save 儲存後後端有 1 份草稿');
  ok(lst.items[0].name === '我的草稿' && lst.items[0].note === '測試備註', 'save 名稱與備註正確');
  ok((await text(b, '#fdSavedPill')).includes('已儲存'), 'save 狀態變「草稿已儲存」');
  const savedId = lst.items[0].id;
  eq(await b.ev(`FlowDesigner.state().draft.id`), savedId, 'save 本地草稿代號＝後端代號');
  // 再改 → 有未儲存的變更
  await b.drag(nodeSel('r_default', 1), nodeSel('r_default', 0), { dx: -50 }); await settle(b); await sleep(250);
  ok((await text(b, '#fdSavedPill')).includes('未儲存'), 'save 再次編輯後顯示「有未儲存的變更」');
  // 重新整理 → 載入
  const canonSaved = (await api('GET', '/drafts/' + savedId)).j.def;
  const openOk = await openDesigner(b);
  ok(openOk, 'save 重新整理頁面後設計器可再開啟');
  await tab(b, 'drafts');
  await b.waitFor(`!!document.querySelector('[data-draft-load]')`, 5000);
  ok((await text(b, '#fdDraftTbl')).includes('我的草稿'), 'save 草稿清單列出已存的草稿');
  await b.click(`[data-draft-load="${savedId}"]`); await sleep(900);
  eq(await b.ev(`FlowDesigner._test.canonStr(FlowDesigner.state().def)`), await b.ev(`FlowDesigner._test.canonStr(${JSON.stringify(canonSaved)})`), 'save 重載後內容與後端儲存的完全相同（UI→JSON→UI 不失真）');
  eq(await b.ev(`FlowDesigner.state().tab`), 'design', 'save 載入後回到設計器分頁');
  eq(await lanePath(b, 'r_default'), ['mgr1', 'gm'], 'save 重載後畫布仍是儲存時的樣子');
  eq(await b.ev(`document.querySelector('#fdName').value`), '我的草稿', 'save 名稱欄帶回');
  // 內容類錯誤不能存：把 mgr2 拖進預設路線（BU 調整又會插入一次 → 套用後重複）
  await b.drag('.fd-pal[data-drag="pnode:mgr2"]', lane('r_default') + ' .fd-slot'); await settle(b); await sleep(300);
  eq((await J(b, `FlowDesigner.state().checks.filter(c=>c.status==='error').map(c=>c.n)`)), [2], 'save 驗證第 2 條錯誤（調整規則使關卡重複）');
  await b.click('#fdSave'); await sleep(500);
  ok(await dlgOpen(b) && (await dlgText(b)).includes('還不能儲存'), 'save 有內容類錯誤：前端先擋並說明');
  await dlgClick(b, 'ok'); await sleep(200);
  eq(await b.ev(`FlowDesigner.state().tab`), 'val', 'save 擋下後帶到「驗證」分頁');
  eq((await api('GET', '/drafts/' + savedId)).j.def.routes[3].path.length, 2, 'save 後端草稿沒有被覆蓋成錯誤內容');
  // 後端也擋（直接打 API）
  const badDef = (await api('GET', '/drafts/' + savedId)).j.def; badDef.routes[3].path = ['n_mgr2', 'n_mgr1'];
  const rb = await api('PUT', '/drafts', { id: savedId, name: 'x', def: badDef });
  eq(rb.status, 400, 'save 後端獨立驗證：同樣的錯誤內容直接打 API 也被拒');
  // 還原錯誤（刪掉 mgr2）
  await tab(b, 'design');
  await b.focus(nodeSel('r_default', 2)); await b.key('Delete'); await sleep(300); await dlgClick(b, 'ok'); await settle(b); await sleep(300);
  eq((await J(b, `FlowDesigner.state().checks.filter(c=>c.status==='error').length`)), 0, 'save 移除重複關卡後驗證錯誤為 0');
  // 樂觀並行：別處先更新
  const cur = (await api('GET', '/drafts/' + savedId)).j;
  const other = JSON.parse(JSON.stringify(cur.def)); other.routes[2].name = '別處改名';
  const ro = await api('PUT', '/drafts', { id: savedId, name: cur.name, def: other, ifUpdatedAt: cur.updatedAt });
  eq(ro.status, 200, 'save（前置）另一個視窗先更新同一份草稿');
  await b.click('#fdSave'); await sleep(900);
  ok(await dlgOpen(b) && (await dlgText(b)).includes('草稿已被更新過'), 'save 過期更新：跳出「草稿已被更新過」對話框，沒有蓋掉對方', await dlgText(b));
  await dlgClick(b, 'ok'); await sleep(200);
  eq((await api('GET', '/drafts/' + savedId)).j.def.routes[2].name, '別處改名', 'save 對方的修改仍在');
  // 另存新草稿
  await b.ev(`(()=>{const i=document.querySelector('#fdName');i.value='另存的草稿';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
  await b.click('#fdSaveAs'); await sleep(900);
  eq((await api('GET', '/drafts')).j.items.length, 2, 'save 另存新草稿：共 2 份');
  // 載入別份會問是否捨棄未儲存的變更
  await b.drag('.fd-pal[data-drag="pnode:chairman"]', lane('r_default') + ' .fd-slot'); await settle(b); await sleep(250);
  await tab(b, 'drafts');
  await b.click(`[data-draft-load="${savedId}"]`); await sleep(400);
  ok(await dlgOpen(b) && (await dlgText(b)).includes('捨棄未儲存的變更'), 'save 有未儲存變更時載入別份草稿：先確認是否捨棄');
  await dlgClick(b, 'cancel'); await sleep(200);
  eq(await b.ev(`FlowDesigner.state().tab`), 'drafts', 'save 取消後留在原處');
  // 刪除草稿
  await b.click(`[data-draft-del="${savedId}"]`); await sleep(300);
  ok(await dlgOpen(b) && (await dlgText(b)).includes('永久刪除'), 'save 刪除前先確認');
  await dlgClick(b, 'ok'); await sleep(800);
  eq((await api('GET', '/drafts')).j.items.length, 1, 'save 刪除後剩 1 份');
  ok(!(await text(b, '#fdDraftTbl')).includes('我的草稿'), 'save 清單同步移除');
  // 稽核
  const logs = await b.ev(`fetch('/api/admin/logs').then(r=>r.json()).then(l=>l.filter(x=>/FLOW_DRAFT/.test(x.action)).map(x=>x.action))`);
  ok(logs.includes('SAVE_FLOW_DRAFT') && logs.includes('DELETE_FLOW_DRAFT'), 'save 稽核日誌有 SAVE／DELETE_FLOW_DRAFT');
  ok((await b.ev(`fetch('/admin.html').then(r=>r.text())`)).includes("SAVE_FLOW_DRAFT: '儲存流程草稿'"), 'save 後台稽核有中文動作名');
  // 清掉
  for (const it of (await api('GET', '/drafts')).j.items) await api('DELETE', '/drafts/' + it.id);
  await b.viewport(1440, 1000);
};

// ── sim：模擬器
G.sim = async (b) => {
  console.log('  [sim]');
  await b.viewport(1440, 2200); await sleep(250);
  await tab(b, 'sim');
  await b.click('#fdSimRun'); await sleep(900);
  ok((await text(b, '#fdSimOut h3')).includes('董事長層級'), 'sim 預設輸入（1,200 萬、層級 1、顧問、ERP）→ 走董事長路線', await text(b, '#fdSimOut h3'));
  eq(await b.ev(`document.querySelectorAll('#fdSimOut .fd-steps li').length`), 3, 'sim 3 關（一級主管、總經理、董事長）');
  ok((await text(b, '#fdSimOut .fd-why')).includes('命中') && (await text(b, '#fdSimOut .fd-why')).includes('不再檢查'), 'sim 白話說明：命中＋後面不再檢查');
  ok((await text(b, '#fdSimOut .fd-why')).includes('折後未稅金額 大於 1,000 萬（本單：1,200 萬）'), 'sim 逐條件顯示本單值');
  // 預設值帶入 6,000 萬
  await b.click('[data-preset="6000,1"]'); await sleep(900);
  ok((await text(b, '#fdSimOut .fd-steps')).includes('董事會決議'), 'sim 6,000 萬 → 董事會代核路線');
  ok((await text(b, '#fdSimOut .fd-steps')).includes('代核'), 'sim 代核關標示');
  // BU ITS：調整規則插入二級主管
  await b.click('[data-preset="800,1"]'); await sleep(600);
  await b.ev(`(()=>{const s=document.querySelector('#fdSimBu');s.value='ITS';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(900);
  eq(await b.ev(`document.querySelectorAll('#fdSimOut .fd-steps li').length`), 2, 'sim 800 萬／ITS：二級主管＋一級主管');
  ok((await text(b, '#fdSimOut .fd-steps')).includes('調整插入'), 'sim 調整規則插入的關卡有標記');
  await b.click('#fdSimNoMgr2'); await sleep(900);
  eq(await b.ev(`document.querySelectorAll('#fdSimOut .fd-steps li').length`), 1, 'sim 找不到二級主管 → 不插入（現行規則），1 關');
  ok((await text(b, '#fdSimOut')).includes('找不到二級主管'), 'sim 假想單說明含「找不到二級主管」');
  await b.click('#fdSimNoMgr2'); await sleep(500);
  // 例外名單
  ok(await visible(b, '#fdSimExcWrap'), 'sim 有例外名單時顯示「二級主管在例外名單內」選項');
  await b.click('#fdSimExc'); await sleep(900);
  eq(await b.ev(`document.querySelectorAll('#fdSimOut .fd-steps li').length`), 1, 'sim 二級主管在例外名單內 → 只剩二級主管（跳過一級主管）');
  await b.click('#fdSimExc'); await sleep(400);
  // 毛利率直接輸入與層級
  await b.ev(`(()=>{const i=document.querySelector('#fdSimPct');i.value='3';i.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(900);
  ok((await text(b, '#fdSimOut .fd-help')).includes('毛利層級 3'), 'sim 毛利率 3%（顧問）→ 層級 3', await text(b, '#fdSimOut .fd-help'));
  await b.ev(`(()=>{const i=document.querySelector('#fdSimPct');i.value='';i.dispatchEvent(new Event('change',{bubbles:true}))})()`);
  // 壞輸入
  await b.ev(`(()=>{const i=document.querySelector('#fdSimAmt');i.value='abc';})()`);
  await b.click('#fdSimRun'); await sleep(400);
  ok((await text(b, '#fdSimOut')).includes('金額要填大於 0'), 'sim 金額錯誤顯示說明');
  await b.ev(`(()=>{const i=document.querySelector('#fdSimAmt');i.value='1200';})()`);
  // OT 層級 1 不可能
  await b.ev(`(()=>{const s=document.querySelector('#fdSimCat');s.value='ot';})()`);
  await b.click('#fdSimBu'); await b.click('#fdSimRun'); await sleep(900);
  ok((await text(b, '#fdSimOut')).includes('沒有層級 1'), 'sim OT 沒有層級 1：後端回報原因');
  await b.ev(`(()=>{const s=document.querySelector('#fdSimCat');s.value='consult';})()`);
  // 比較：把總經理路線放寬
  await tab(b, 'design');
  await b.focus(nodeSel('r_l2', 1)); await b.key('Delete'); await sleep(300); await dlgClick(b, 'ok'); await settle(b); await sleep(300);
  eq(await lanePath(b, 'r_l2'), ['mgr1'], 'sim（前置）把總經理路線的總經理拿掉（放寬，但不違反底線）');
  await tab(b, 'sim');
  await b.click('[data-preset="300,2"]'); await sleep(700);
  await b.ev(`(()=>{const s=document.querySelector('#fdSimBu');s.value='ERP';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(500);
  await b.click('#fdSimCmp'); await sleep(900);
  ok((await text(b, '#fdSimOut')).includes('放寬') && (await b.ev(`document.querySelectorAll('#fdSimOut .fd-cmpgrid .fd-steps li.st-del').length`)) === 1, '比較模式：標示「放寬」並以虛線標出草稿拿掉的關卡', await text(b, '#fdSimOut'));
  ok(await b.ev(`FlowDesigner.state().checks.find(c=>c.n===9).status`) === 'warn', 'sim 驗證第 9 條放寬偵測警告');
  // 設計器畫布高亮
  await tab(b, 'design');
  ok(await exists(b, lane('r_l2') + '.hit') && (await count(b, '#fdLanes .fd-node.hit-n')) === 1 && (await text(b, '.fd-node.hit-n .fd-ord')) === '1', 'sim 設計器畫布高亮命中的路線與關卡序號');
  ok((await text(b, lane('r_l2') + ' .fd-pill.cur')).includes('本單走這條'), 'sim 命中路線有「本單走這條」徽章（不只靠顏色）');
  // 邊界掃描
  await tab(b, 'sim');
  await b.click('#fdSimScan'); await sleep(2500);
  ok((await count(b, '#fdScanTbl tbody tr')) >= 6, 'sim 邊界掃描表有列', String(await count(b, '#fdScanTbl tbody tr')));
  const tbl = await text(b, '#fdScanTbl');
  ok(tbl.includes('10,000,000.01 元') && tbl.includes('1,000 萬') && tbl.includes('5,000 萬'), 'sim 邊界掃描含 1,000 萬／5,000 萬 ±0.01');
  ok(await b.ev(`document.querySelectorAll('#fdScanTbl tr.loose').length > 0`), 'sim 邊界掃描標出「放寬」列');
  ok((await text(b, '#fdSimOut')).includes('全取樣統計'), 'sim 邊界掃描附全取樣統計');
  ok(await exists(b, '#fdScanCsv'), 'sim 有「下載 CSV」按鈕');
  eq(await b.ev(`FlowDesigner._test.csvCell('=cmd|calc')`), `"'=cmd|calc"`, 'sim CSV 欄位防公式注入');
  await b.click('#fdSimClear'); await sleep(400);
  await tab(b, 'design');
  eq(await count(b, '#fdLanes .fd-lane.hit'), 0, 'sim 清除高亮後畫布沒有命中標記');
  await b.viewport(1440, 1000);
};

// ── val：驗證面板
G.val = async (b) => {
  console.log('  [val]');
  await b.viewport(1440, 2200); await sleep(250);
  await b.drag('.fd-pal[data-drag="pnode:mgr2"]', lane('r_default') + ' .fd-slot'); await settle(b); await sleep(300);
  await tab(b, 'val');
  eq(await count(b, '#fdVal .fd-chk > li'), 11, 'val 11 條驗證');
  ok((await text(b, '#fdVal .fd-sumbar')).includes('錯誤 1'), 'val 摘要顯示錯誤 1');
  ok(await exists(b, '#fdVal .fd-chk > li.error[data-chk="2"]'), 'val 第 2 條標為錯誤');
  ok((await text(b, '#fdVal .fd-chk > li.error')).includes('✕') || (await text(b, '#fdVal .fd-chk > li.error')).includes('錯誤'), 'val 錯誤有文字＋圖示標記（不只靠顏色）');
  ok(await exists(b, '#fdCnt-val') && !(await b.ev(`document.querySelector('#fdCnt-val').hidden`)), 'val 分頁標籤上有問題數徽章');
  await b.click('#fdVal [data-chk="2"] [data-loc]'); await sleep(500);
  eq(await b.ev(`FlowDesigner.state().tab`), 'design', 'val 點「點擊定位」切到設計器');
  ok(await b.ev(`document.querySelector('#fdLanes .fd-lane[data-lane="r_default"] .fd-node.flash, #fdLanes .fd-lane.flash') !== null || document.activeElement.closest('.fd-lane')!==null`), 'val 定位後目標元素被標示或取得焦點');
  // 底線檢查的說明
  await tab(b, 'val');
  ok((await text(b, '#fdVal [data-chk="10"]')).includes('取樣'), 'val 第 10 條說明含邊界取樣組數');
  await b.viewport(1440, 1000);
};

// ── ov：總覽（唯讀自由畫布）
G.ov = async (b) => {
  console.log('  [ov]');
  await b.viewport(1440, 1300); await sleep(250);
  await tab(b, 'ov'); await sleep(1200);
  const n0 = await count(b, '.fo-node'), e0 = await count(b, '.fo-edge');
  ok(n0 >= 15 && e0 >= 15, '總覽 節點與連線有渲染', `${n0}/${e0}`);
  eq(await count(b, '.fo-node.decision'), 3, '總覽 3 個判斷菱形');
  eq(await count(b, '.fo-node.mod'), 2, '總覽 2 個調整規則備註卡');
  eq(await count(b, '#fdOutline li'), 4, '總覽 文字大綱 4 條路線');
  ok(await b.ev(`document.querySelectorAll('.fo-node').length === document.querySelectorAll('.fo-node[tabindex="0"]').length && Array.from(document.querySelectorAll('.fo-node')).every(e=>e.getAttribute('aria-label'))`), '總覽 每個節點可聚焦且有 aria-label');
  ok(!(await b.ev(`!!document.querySelector('#fdOvHost [contenteditable], #fdOvHost input, #fdOvHost textarea')`)), '總覽 唯讀：沒有任何可編輯元件');
  // 縮放／平移
  const k0 = await b.ev(`FlowDesigner.state().ovCtl.state().k`);
  await b.click('#fdOvIn'); await sleep(200);
  ok((await b.ev(`FlowDesigner.state().ovCtl.state().k`)) > k0, '總覽 ＋放大');
  await b.click('#fdOvOut'); await b.click('#fdOvOut'); await sleep(200);
  ok((await b.ev(`FlowDesigner.state().ovCtl.state().k`)) < k0, '總覽 －縮小');
  await b.click('#fdOvFit'); await sleep(200);
  ok(Math.abs((await b.ev(`FlowDesigner.state().ovCtl.state().k`)) - k0) < 0.01, '總覽 縮放至適合回到原比例');
  const r = await b.rectOf('#fdOvHost');
  const t0 = await b.ev(`FlowDesigner.state().ovCtl.state().tx`);
  await b.dragTo(r.cx, r.y + 20, r.cx + 90, r.y + 20, 6);
  ok(Math.abs((await b.ev(`FlowDesigner.state().ovCtl.state().tx`)) - t0 - 90) < 3, '總覽 拖曳平移（真實滑鼠）');
  await b.click('#fdOvFit');
  await b.ev(`document.querySelector('#fdOvHost svg').dispatchEvent(new WheelEvent('wheel',{deltaY:-120,clientX:${r.cx},clientY:${r.cy},bubbles:true,cancelable:true}))`); await sleep(150);
  ok((await b.ev(`FlowDesigner.state().ovCtl.state().k`)) > k0, '總覽 滾輪放大');
  await b.click('#fdOvFit');
  await b.focus('#fdOvHost svg');
  await b.key('+'); await sleep(150);
  ok((await b.ev(`FlowDesigner.state().ovCtl.state().k`)) > k0, '總覽 鍵盤 + 放大');
  await b.key('0'); await sleep(150);
  ok(Math.abs((await b.ev(`FlowDesigner.state().ovCtl.state().k`)) - k0) < 0.01, '總覽 鍵盤 0 縮放至適合');
  // 模擬高亮
  await tab(b, 'sim');
  await b.click('[data-preset="1200,1"]'); await sleep(500);
  await b.ev(`(()=>{const s=document.querySelector('#fdSimBu');s.value='ITS';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(900);
  await tab(b, 'ov'); await sleep(1200);
  eq(await count(b, '.fo-node.approval.hit'), 4, '總覽 模擬 1,200 萬／ITS：高亮 4 個關卡（含調整插入）');
  eq(await count(b, '.fo-node.approval.hit.fd-c-teal'), 1, '總覽 高亮的二級主管（調整插入，青色）');
  ok(await b.ev(`Array.from(document.querySelectorAll('.fo-node.approval.hit text')).some(t=>t.textContent.includes('調整插入'))`), '總覽 調整插入的關卡有文字標記');
  ok((await count(b, '.fo-edge.hit')) >= 5, '總覽 命中路徑的連線有高亮');
  ok((await count(b, '.fo-node.decision.hit')) >= 2, '總覽 走過的判斷菱形有高亮');
  ok((await text(b, '#fdOvInfo')).includes('高亮'), '總覽 說明列顯示高亮資訊');
  await b.click('#fdOvSim'); await sleep(1000);
  eq(await count(b, '.fo-node.approval.hit'), 0, '總覽 取消「套用模擬輸入高亮」後無高亮');
  await b.click('#fdOvSim'); await sleep(900);
  ok((await text(b, '#fdOutline li.hit')).includes('董事長'), '總覽 文字大綱標出命中路線');
  await b.viewport(1440, 1000);
};

// ── xss：HTML 內容一律以純文字顯示
G.xss = async (b) => {
  console.log('  [xss]');
  const api = (m, u, body) => b.ev(`fetch('/api/admin/quote-flow${u}',{method:'${m}',headers:{'Content-Type':'application/json'}${body ? `,body:JSON.stringify(${JSON.stringify(body)})` : ''}}).then(async r=>({status:r.status,j:await r.json()}))`);
  const def = (await api('GET', '/designer/default-flow')).j.def;
  def.routes[2].name = '<svg onload=x=1>';
  def.modifiers[0].name = '<svg onload=__x=2>';
  const r = await api('PUT', '/drafts', { name: '<img src=x onerror=window.__xss=3>', note: '<b onmouseover=window.__xss=4>備註</b>', def });
  eq(r.status, 200, 'xss 含 HTML 的草稿名稱／路線名稱／調整規則名稱可儲存');
  await openDesigner(b);
  await b.ev(`window.__xss=0`);
  await tab(b, 'drafts'); await sleep(600);
  eq(await count(b, '#fdDraftTbl img, #fdDraftTbl svg, #fdDraftTbl [onmouseover], #fdDraftTbl [onerror]'), 0, 'xss 草稿清單：名稱／備註以純文字顯示（沒有注入的 img／svg 元素或事件屬性）');
  ok((await text(b, '#fdDraftTbl')).includes('<img src=x onerror=window.__xss=3>'), 'xss 草稿清單顯示原始文字');
  await b.click(`[data-draft-load="${r.j.id}"]`); await sleep(1200);
  eq(await count(b, '#fdApp img, #fdApp svg:not(.fd-ovsvg)'), 0, 'xss 載入後設計器畫面沒有被注入的元素');
  eq(await text(b, lane('r_l2') + ' .fd-lname'), '<svg onload=x=1>', 'xss 畫布路線名稱為純文字');
  ok((await text(b, '.fd-mod')).includes('<svg onload=__x=2>'), 'xss 調整規則名稱為純文字');
  ok((await b.ev(`document.querySelector('#fdName').value`)) === '<img src=x onerror=window.__xss=3>', 'xss 名稱欄以 value 帶入（不會被解析）');
  await b.click(`${lane('r_l2')} [data-act="lane-del"]`); await sleep(300);
  ok((await dlgText(b)).includes('<svg onload=x=1>') && !(await b.ev(`!!document.querySelector('#fdDlg img')`)), 'xss 刪除確認對話框中的路線名稱為純文字');
  await dlgClick(b, 'cancel');
  await tab(b, 'val'); await sleep(300);
  eq(await count(b, '#fdVal img'), 0, 'xss 驗證面板（含路線名稱）無注入');
  await tab(b, 'ov'); await sleep(1200);
  eq(await count(b, '#fdOvHost img'), 0, 'xss 總覽（SVG）無 img');
  ok(await b.ev(`Array.from(document.querySelectorAll('.fo-node text')).some(t=>t.textContent.includes('<svg onload'))`), 'xss 總覽以文字節點顯示路線名稱');
  ok(await b.ev(`Array.from(document.querySelectorAll('#fdOutline li')).some(li=>li.textContent.includes('<svg onload=x=1>')) && document.querySelectorAll('#fdOutline img, #fdOutline svg').length===0`), 'xss 總覽文字大綱為純文字');
  await tab(b, 'sim'); await b.click('#fdSimRun'); await sleep(900);
  eq(await count(b, '#fdSimOut img, #fdSimOut svg'), 0, 'xss 模擬器說明（含路線名稱）無注入');
  await tab(b, 'design');
  await b.ev(`FlowDesigner._test.tryApply(d=>{d.routes[0].name='x'},'t')`); await sleep(500);
  eq(await b.ev(`[window.__xss, typeof window.x, typeof window.__x].join()`), '0,undefined,undefined', 'xss 全程沒有任何注入的腳本執行（__xss／x／__x 都沒被設定）');
  for (const it of (await api('GET', '/drafts')).j.items) await api('DELETE', '/drafts/' + it.id);
};

// ── mobile：手機唯讀
G.mobile = async (b) => {
  console.log('  [mobile]');
  const ok1 = await openDesigner(b, 375, 812);
  ok(ok1, 'mobile 375 寬可開啟設計器');
  ok(!(await visible(b, '#fdPalette')), 'mobile 節點面板隱藏（唯讀）');
  ok(!(await visible(b, '.fd-ltools')) && !(await visible(b, '.fd-grip')), 'mobile 路線工具列與拖拉把手隱藏');
  ok(await visible(b, '.fd-mobile-note'), 'mobile 顯示「手機唯讀」說明');
  eq(await b.ev(`FlowDesigner.state().def.routes.length`), 4, 'mobile 仍載入並顯示流程（唯讀）');
  ok(await b.ev(`document.querySelector('#fdSave').disabled === true && document.querySelector('#fdNew').disabled === true`), 'mobile 儲存／新增按鈕停用');
  const c0 = await defCanon(b);
  await b.dblclick(nodeSel('r_board', 0)); await sleep(300);
  ok(!(await visible(b, '#fdDrawer')), 'mobile 雙擊卡片不開編輯抽屜');
  await b.focus(nodeSel('r_l3', 2)); await b.key('Delete'); await sleep(300);
  ok(!(await dlgOpen(b)), 'mobile Delete 鍵不跳出刪除確認');
  eq(await defCanon(b), c0, 'mobile 草稿內容沒有任何變動');
  ok((await b.ev(`FlowDesigner._test.tryApply(d=>{d.routes[0].name='zzz'},'x').then(r=>r.ok)`)) === false, 'mobile 即使由程式呼叫編輯入口也被拒（canEdit=false）');
  ok(!(await b.ev(`document.documentElement.scrollWidth > window.innerWidth + 1`)), 'mobile 頁面沒有橫向捲動');
  // 模擬器可用
  await tab(b, 'sim');
  await b.click('#fdSimRun'); await sleep(900);
  ok((await count(b, '#fdSimOut .fd-steps li')) === 3, 'mobile 模擬器可用（3 關）');
  ok(!(await b.ev(`document.documentElement.scrollWidth > window.innerWidth + 1`)), 'mobile 模擬器頁面沒有橫向捲動');
  await tab(b, 'ov'); await sleep(1200);
  ok((await count(b, '.fo-node')) >= 15 && (await b.ev(`FlowDesigner.state().ovCtl.state().k`)) > 0.15, 'mobile 總覽可用');
  const rct = await b.rectOf('#fdOvHost');
  ok(rct.w <= 375 && rct.w > 250, 'mobile 總覽寬度在畫面內', JSON.stringify(rct));
  await tab(b, 'val'); await sleep(300);
  ok(!(await b.ev(`document.documentElement.scrollWidth > window.innerWidth + 1`)), 'mobile 驗證分頁沒有橫向捲動');
  await b.viewport(1440, 1000);
};

// ── theme：四種外觀 × 六主題的對比
G.theme = async (b) => {
  console.log('  [theme]');
  await b.viewport(1440, 2400); await sleep(250);
  await b.click('#fdAddLane'); await settle(b);
  await tab(b, 'sim'); await b.click('#fdSimRun'); await sleep(900);
  await tab(b, 'design'); await sleep(300);
  const SELS = ['.fd-lname', '.fd-cond', '.fd-help', '.fd-muted', '.fd-node .fd-nm', '.fd-node .fd-rl', '.fd-node .fd-stepno', '.fd-lend', '.fd-btn', '.fd-btn.primary', '.fd-pill.warn', '.fd-pill.ok', '.fd-pill.err', '.fd-pill.gray', '.fd-pill.cur', '.fd-pill.proxy', '.fd-tab[aria-selected="true"]', '.fd-tab:not([aria-selected="true"])', '.fd-pal .fd-nm', '.fd-pal .fd-ds', '.fd-mod', '.fd-title', '.fd-note', '.fd-chip.k', '.fd-slot', '.fd-pri', '.fd-ord'];
  const modes = [['淺色', 'body.classList.remove("glass","dark")', '#ffffff'], ['深色', 'body.classList.remove("glass");body.classList.add("dark")', '#0d1117']];
  const THEMES = ['ocean', 'noir', 'aurora', 'dusk', 'ember', 'graphite'];
  THEMES.forEach((t) => modes.push([`玻璃淺色-${t}`, `body.classList.add("glass");body.classList.remove("dark");body.dataset.theme="${t}"`, '#cddcff']));
  const DARKBG = { ocean: '#050914', noir: '#09090a', aurora: '#05181a', dusk: '#160a24', ember: '#1a0a07', graphite: '#0d1015' };
  THEMES.forEach((t) => modes.push([`玻璃深色-${t}`, `body.classList.add("glass","dark");body.dataset.theme="${t}"`, DARKBG[t]]));
  const results = {};
  for (const [name, js, backdrop] of modes) {
    await b.ev(`(()=>{const body=document.body;${js}})()`); await sleep(1100);
    const rows = [];
    for (const s of SELS) rows.push(...(await J(b, `(${CONTRAST_FN})(${JSON.stringify(backdrop)})(${JSON.stringify(s)})`)));
    const bad = rows.filter((r) => r.ratio < 4.5);
    results[name] = { n: rows.length, bad };
    ok(rows.length >= 25 && bad.length === 0, `theme 對比 ≥4.5:1：${name}（檢查 ${rows.length} 個元素）`, JSON.stringify(bad.slice(0, 4)));
    if (SHOTS && ['淺色', '深色'].includes(name)) { await b.viewport(1440, 1000); await b.ev('window.scrollTo(0,0)'); await sleep(200); await b.shot(path.join(SHOTS, name === '淺色' ? 'design_light_1440.png' : 'design_dark_1440.png')); await b.viewport(1440, 2400); }
  }
  // 焦點樣式可見
  await b.ev(`document.body.classList.remove("glass","dark")`); await sleep(300);
  await b.focus(nodeSel('r_l3', 0));
  await b.key('Tab'); await sleep(200);   // 鍵盤操作後才會套用 :focus-visible
  const outline = await b.ev(`(()=>{const e=document.activeElement;const s=getComputedStyle(e);return s.outlineStyle+' '+s.outlineWidth})()`);
  ok(/solid|auto/.test(outline) && !/ 0px/.test(outline), 'theme 焦點樣式清楚（outline）', outline);
  await b.ev(`document.body.classList.add("glass","dark");document.body.dataset.theme="ocean"`);
  await b.viewport(1440, 1000);
};

// ── motion：prefers-reduced-motion
G.motion = async (b) => {
  console.log('  [motion]');
  await b.media([{ name: 'prefers-reduced-motion', value: 'reduce' }]); await sleep(300);
  await tab(b, 'sim'); await b.click('#fdSimRun'); await sleep(800);
  await tab(b, 'ov'); await sleep(1200);
  const anims = await b.ev(`document.querySelector('#fdApp').getAnimations({subtree:true}).filter(a=>a.playState==='running').length`);
  eq(anims, 0, 'motion 減少動態：沒有任何執行中的動畫（含連線流動）');
  eq(await b.ev(`getComputedStyle(document.querySelector('.fo-edge.hit .fo-line')).animationName`), 'none', 'motion 減少動態：命中連線沒有流動動畫');
  await tab(b, 'design');
  eq(await b.ev(`getComputedStyle(document.querySelector('.fd-node')).transitionDuration`), '0s', 'motion 減少動態：卡片沒有過場');
  await b.media([{ name: 'prefers-reduced-motion', value: 'no-preference' }]); await sleep(300);
  await tab(b, 'ov'); await sleep(600);
  ok((await b.ev(`getComputedStyle(document.querySelector('.fo-edge.hit .fo-line')).animationName`)) === 'fo-flow', 'motion 一般模式：命中連線有流動動畫（對照組）');
  await tab(b, 'design');
};

// ── shots：截圖（FD_SHOTS 指定資料夾才執行）
G.shots = async (b) => {
  console.log('  [shots]');
  const S1 = (n) => path.join(SHOTS, n);
  // 後台版面把 body 高度固定為視窗高度：用較高的視窗並回到頂端截圖，避免卷動後底部露出頁面底色
  await b.viewport(1440, 1500); await sleep(400);
  const top = async () => { await b.ev('window.scrollTo(0,0)'); await sleep(250); };
  await b.ev(`document.body.classList.add('glass','dark');document.body.dataset.theme='ocean'`); await sleep(900);
  await top(); await b.shot(S1('design_dark_1440.png'));
  await b.ev(`document.body.classList.remove('glass','dark')`); await sleep(900);
  await top(); await b.shot(S1('design_light_1440.png'));
  await b.ev(`document.body.classList.add('glass','dark')`); await sleep(900);
  // 錯誤：把董事長拖到總經理路線 → 底線拒絕
  await b.drag(nodeSel('r_l3', 2), lane('r_l2') + ' .fd-path'); await settle(b); await sleep(500);
  ok((await bannerText(b)).includes('保護底線'), 'shots 底線拒絕訊息已出現');
  await top(); await b.shot(S1('error_floor_1440.png'));
  await b.click('#fdBannerX'); await sleep(200);
  // 抽屜：路線條件編輯器
  await b.dblclick(`${lane('r_l3')} .fd-cond`); await sleep(500);
  await top(); await b.shot(S1('drawer_1440.png'));
  await b.click('#fdDrawer [data-act="dr-close"]'); await sleep(200);
  // 模擬器（含比較與邊界掃描）
  await tab(b, 'sim');
  await b.click('[data-preset="1200,1"]'); await sleep(600);
  await b.ev(`(()=>{const s=document.querySelector('#fdSimBu');s.value='ITS';s.dispatchEvent(new Event('change',{bubbles:true}))})()`); await sleep(700);
  await b.click('#fdSimCmp'); await sleep(800);
  await b.click('#fdSimScan'); await sleep(2500);
  await top(); await b.shot(S1('sim_1440.png'));
  // 總覽（含模擬高亮）
  await tab(b, 'ov'); await sleep(1500);
  await top(); await b.shot(S1('overview_1440.png'));
  // 驗證
  await tab(b, 'val'); await sleep(300); await top(); await b.shot(S1('validation_1440.png'));
  // 手機 375
  await openDesigner(b, 375, 900); await sleep(600);
  await top(); await b.shot(S1('mobile_375_design.png'));
  await tab(b, 'sim'); await b.click('[data-preset="1200,1"]'); await sleep(900); await top(); await b.shot(S1('mobile_375_sim.png'));
  await tab(b, 'ov'); await sleep(1500); await top(); await b.shot(S1('mobile_375_overview.png'));
  await b.viewport(1440, 1000);
  ok(fs.readdirSync(SHOTS).filter((f) => /.png$/.test(f)).length >= 9, 'shots 截圖已輸出（至少 9 張）');
};

// ── static：沒有 inline script／on*／style=
G.static = async () => {
  console.log('  [static]');
  const files = ['flow-designer.js', 'flow-overview.js', 'flow-designer.css'].map((f) => [f, fs.readFileSync(path.join(ROOT, '_client', f), 'utf8')]);
  for (const [f, t] of files) {
    if (f.endsWith('.js')) {
      ok(!/\bstyle\s*=\s*["'`\\]/.test(t), `static ${f} 沒有 style= 屬性字串`);
      ok(!/\son[a-z]+\s*=\s*["'`\\]/i.test(t), `static ${f} 沒有 on* 事件屬性字串`);
      ok(!/<script/i.test(t), `static ${f} 沒有 <script> 字串`);
      ok(!/\beval\s*\(|new Function\s*\(|document\.write\s*\(|insertAdjacentHTML|setAttribute\(\s*['"]style['"]/.test(t), `static ${f} 沒有 eval／new Function／document.write／insertAdjacentHTML／setAttribute('style')`);
    }
  }
  const js = files[0][1];
  const sinks = (js.match(/\.innerHTML\s*=/g) || []).length + (js.match(/\.outerHTML\s*=/g) || []).length;
  eq(sinks, 2, 'static flow-designer.js 只有 2 個 innerHTML 出口（put() 與 drawerErr 的錯誤框，皆為已跳脫的 Raw）');
  ok(!/[.]innerHTML|[.]outerHTML|insertAdjacentHTML/.test(files[1][1]), 'static flow-overview.js 完全不使用 innerHTML（SVG 以 createElementNS＋textContent）');
  const adm = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
  const gitDiff = require('child_process').spawnSync('git', ['diff', '-U0', process.env.FD_BASE_REF || '12a763fc652d973e370bab5637b1fb1b50df7b75', '--', '_client/admin.html'], { cwd: ROOT, encoding: 'utf8' }).stdout || '';
  const added = gitDiff.split('\n').filter((l) => l.startsWith('+') && !l.startsWith('+++'));
  ok(added.length > 5, 'static admin.html 有新增行（可比對）', String(added.length));
  ok(!added.some((l) => /\sstyle\s*=|\son[a-z]+\s*=/i.test(l)), 'static admin.html 新增的行沒有 style= 或 on* 屬性', added.filter((l) => /\sstyle\s*=|\son[a-z]+\s*=/i.test(l)).join('|'));
  ok(!added.some((l) => /<script>(?!<)/.test(l)) && added.filter((l) => /<script/.test(l)).every((l) => /src="/.test(l)), 'static admin.html 只新增外部 script 標籤（沒有 inline script）');
  ok(/data-sec="flow-designer"/.test(adm) && /id="sec-flow-designer"/.test(adm) && /initApprovalFlow|window\.FlowDesigner\.init/.test(adm), 'static admin.html 有側欄項、sec-* 容器、lazy-init 分派');
  const sw = fs.readFileSync(path.join(ROOT, '_client', 'sw.js'), 'utf8');
  ok(parseInt((sw.match(/SW_VERSION = 'itts-crm-v(\d+)'/) || [])[1] || '0', 10) >= 16, 'static sw.js 版號已提升（至少 v16；之後別的功能提版不應讓此檢查失敗）');
  // 純函式：金額換算與 AST 往返
  const FDm = require(path.join(ROOT, '_client', 'flow-designer.js'));
  const T = FDm._test;
  eq([T.wanToYuan('1000'), T.wanToYuan('0.5'), T.wanToYuan('1000.000001'), T.wanToYuan('1000.0001'), T.wanToYuan('0'), T.wanToYuan('abc'), T.wanToYuan('1e3'), T.wanToYuan('-1')], ['10000000', '5000', '10000000.01', '10000001', null, null, null, null], 'static 萬元→元換算（BigInt，無浮點）');
  eq([T.yuanToWan('10000000'), T.yuanToWan('10000000.01'), T.yuanToWan('5000')], ['1000', '1000.000001', '0.5'], 'static 元→萬元換算');
  const QF = require(path.join(ROOT, 'lib', 'quoteFlow'));
  const FD = require(path.join(ROOT, 'lib', 'quoteFlowDesigner'));
  const def = JSON.parse(JSON.stringify(QF.legacyDefault({ buSettings: { ITS: { hasMgr2Step: true } }, exceptions: [] })));
  let rt = 0;
  for (const r of def.routes) { const e = T.astToEdit(r.when); const back = T.editToAst(e); if (JSON.stringify(back) !== JSON.stringify(r.when)) { rt++; console.log('  往返不一致', r.id, JSON.stringify(r.when), JSON.stringify(back)); } }
  eq(rt, 0, 'static 預設流程每條路線的條件 AST → 造句 → AST 往返完全相同（UI→JSON→UI 不失真）');
  eq(T.astToEdit({ not: { cmp: ['amount', '>', '1'] } }), { unsupported: true }, 'static 巢狀（not）條件標為超出編輯器範圍');
  eq(T.astToEdit({ all: [{ any: [{ cmp: ['amount', '>', '1'] }] }] }), { unsupported: true }, 'static 巢狀 all/any 標為超出編輯器範圍');
  // 前後端白話文字一致
  const items = [['amount', '>', '10000000'], ['amount', '<=', '9999999.99'], ['marginLevel', '==', 3], ['marginPct', '>=', '12.5'], ['category', 'in', ['consult', 'ot']], ['category', '!=', 'other'], ['classKeys', 'contains', 'software'], ['ownerBuPrimary', 'not_in', ['ERP', '資安']], ['owner', '==', 'u1'], ['category', 'in', []]];
  let tx = 0;
  for (const [v, o, val] of items) { const a = T.itemText({ v, o, val }), b2 = FD.describeCmp(v, o, val); if (a !== b2) { tx++; console.log('  文字不一致', a, '|', b2); } }
  eq(tx, 0, 'static 前端 itemText 與後端 describeCmp 的白話完全一致（10 種）');
  eq(T.esc(`<img src=x onerror="a('b')">&`), '&lt;img src=x onerror=&quot;a(&#39;b&#39;)&quot;&gt;&amp;', 'static esc 跳脫五種字元');
  eq(T.h`<b>${'<i>'}</b>${T.raw('<u>')}${['<a>', T.h`<s>${'&'}`]}`.s, '<b>&lt;i&gt;</b><u>&lt;a&gt;<s>&amp;', 'static h`` 模板：插值預設跳脫，Raw 原樣，陣列逐項處理');
  const vm = require('vm');
  new vm.Script(js);   // 語法
  ok(true, 'static flow-designer.js 可被 vm 解析');
  // 對比 token（靜態檢查 css 檔有 prefers-reduced-motion／forced-colors）
  const css = files[2][1];
  ok(/prefers-reduced-motion:\s*reduce/.test(css) && /forced-colors/.test(css) && /@media \(max-width: 640px\)/.test(css), 'static css 有 reduced-motion／forced-colors／640px 手機規則');
  ok(!/!important/.test(css.replace(/\.fd-nav-hidden[^}]*\}/, '').replace(/\.fd-edit-only[^}]*\}/g, '').replace(/body\.fd-dragging[^}]*\}/g, '').replace(/animation: none !important; transition: none !important/g, '').replace(/display: none !important/g, '').replace(/cursor: grabbing !important; user-select: none !important/g, '')), 'static css 除必要處外不使用 !important');
};

// ── gate：開關關閉／非管理員
async function gateOff() {
  console.log('  [gate：開關關閉]');
  fx.seed();
  try {
    await fx.startServer({});
    await withBrowser({ cdpPort: CDP, maxMs: 120000 }, async (b) => {
      ok(await login(b, '_s_admin'), 'gate 管理員登入（開關關閉）');
      await b.nav(fx.BASE + '/admin.html');
      await b.waitFor('typeof loadUsers==="function" && !!document.querySelector("#sec-flow-designer")', 20000);
      await sleep(2500);
      ok(await b.ev(`document.querySelector('#fdNavItem').classList.contains('fd-nav-hidden') && getComputedStyle(document.querySelector('#fdNavItem')).display==='none'`), 'gate 開關關閉：側欄項隱藏');
      eq(await b.ev(`FlowDesigner.state().enabled`), false, 'gate 開關關閉：前端不啟用');
      eq(await b.ev(`FlowDesigner.state().inited`), false, 'gate 開關關閉：沒有建立任何畫面');
      eq(await b.ev(`document.querySelector('#fdHost').children.length`), 0, 'gate 開關關閉：容器是空的');
      eq(await b.ev(`document.querySelector('.nav-domain[data-domain="quote"] .nav-domain-count').textContent.trim()`), '4', 'gate 開關關閉：領域數量徽章不變（4）');
      const st = await b.ev(`fetch('/api/admin/quote-flow/designer/status').then(r=>r.status)`);
      eq(st, 404, 'gate 開關關閉：端點 404');
    });
  } finally { await fx.stopServer(); fx.cleanup(); }
}
G.gate = async (b) => {
  console.log('  [gate：非管理員]');
  await b.nav(fx.BASE + '/login.html');
  await b.ev(`fetch('/api/logout',{method:'POST'})`);
  ok(await login(b, '_s_own'), 'gate 業務登入');
  await b.nav(fx.BASE + '/admin.html'); await sleep(2500);
  eq(await b.ev(`window.FlowDesigner ? FlowDesigner.state().enabled : false`), false, 'gate 非管理員：前端不啟用設計器');
  ok(await b.ev(`(()=>{const n=document.querySelector('#fdNavItem');return !n || n.classList.contains('fd-nav-hidden')})()`), 'gate 非管理員：看不到側欄項');
  const codes = await b.ev(`Promise.all(['/designer/status','/drafts','/designer/default-flow'].map(u=>fetch('/api/admin/quote-flow'+u).then(r=>r.status)))`);
  ok(codes.every((c) => c === 403 || c === 401), 'gate 非管理員打端點：403／401', JSON.stringify(codes));
};

// ═════════════════════════ 執行 ═════════════════════════
async function loginAndOpen(b) {
  ok(await login(b, '_s_admin'), '管理員登入');
  const o = await openDesigner(b, 1440, 1000);
  ok(o, '開啟設計器（開關 on）', JSON.stringify(b.errors.slice(0, 3)));
  return o;
}
const ORDER = ['render', 'drag', 'drawer', 'keyboard', 'save', 'sim', 'val', 'ov', 'xss', 'motion', 'theme', 'mobile'];
if (SHOTS && !ONLY.length) ORDER.push('shots');   // FD_SHOTS 與 FD_UI_ONLY=shots 都可單獨觸發
async function runGroups(names, env) {
  fx.seed();
  try {
    await fx.startServer(Object.assign({ QUOTE_FLOW_DESIGNER: '1' }, env || {}));
    await withBrowser({ cdpPort: CDP, maxMs: 560000 }, async (b) => {
      if (!(await loginAndOpen(b))) return;
      for (const g of names) {
        if (g === 'gate' || g === 'static') continue;
        // 每個群組從乾淨狀態開始：重新整理＋開設計器
        if (g !== names.filter((x) => x !== 'gate' && x !== 'static')[0]) { if (!(await openDesigner(b, 1440, 1000))) { ok(false, `${g} 開啟設計器`); continue; } }
        try { await G[g](b); } catch (e) { ok(false, `${g} 群組例外：${e.message}`); }
      }
      if (names.includes('gate')) { try { await G.gate(b); } catch (e) { ok(false, 'gate 例外：' + e.message); } }
      ok(b.errors.length === 0, '整段執行頁面沒有任何 JS 例外或 console.error', JSON.stringify(b.errors.slice(0, 3)));
    });
  } finally { await fx.stopServer(); fx.cleanup(); }
}

// ── 變異測試：把前端檔換成改壞的版本，對應群組必須變紅
const SRC = (f) => fs.readFileSync(path.join(ROOT, '_client', f), 'utf8');
const MUTANTS = [
  ['esc 不跳脫', 'flow-designer.js', "const esc = (s) => String(s == null ? '' : s).replace(/[&<>\"']/g, (c) => ESC_MAP[c]);", "const esc = (s) => String(s == null ? '' : s);", ['xss']],
  ['底線即時拒絕失效', 'flow-designer.js', "if (rc.floor && rc.floor.ok === false) {", "if (false) {", ['drag', 'drawer']],
  ['卡片拖放位置永遠算成最前面', 'flow-designer.js', "if (!before) idx++;\n    }\n    return { type: 'node'", "idx = 0;\n    }\n    return { type: 'node'", ['drag']],
  ['驗證定位失效', 'flow-designer.js', "function locate(loc) {\n    if (S.tab !== 'design') switchTab('design');", "function locate(loc) {\n    return;\n    if (S.tab !== 'design') switchTab('design');", ['val']],
  ['模擬高亮不套用', 'flow-designer.js', "const hit = !!(sim && sim.route.id === r.id);", "const hit = false;", ['sim']],
  ['手機唯讀失效', 'flow-designer.js', "const canEdit = () => S.enabled && !!S.def && !isMobile();", "const canEdit = () => S.enabled && !!S.def;", ['mobile']],
  ['總覽高亮不標記', 'flow-overview.js', "(n.hit ? ' hit' : (simOn && n.kind !== 'mod' ? ' dim' : ''))", "''", ['ov']],
  ['儲存不帶 ifUpdatedAt（樂觀並行失效）', 'flow-designer.js', "body.ifUpdatedAt = S.draft.updatedAt;", "", ['save']],
  ['刪除關卡不再確認', 'flow-designer.js', "    if (!ok) return;\n    const res = await tryApply((dd) => { laneOf(dd, l).path.splice(i, 1);", "    const res = await tryApply((dd) => { laneOf(dd, l).path.splice(i, 1);", ['keyboard']],
  ['「提交覆核」按鈕不再停用', 'flow-designer.js', 'disabled aria-disabled="true" id="fdPropose"', 'id="fdPropose"', ['render']],
  ['對比：節點文字色改成淡灰', 'flow-designer.css', ".fd-node .fd-rl { font-size: 12px; color: var(--fd-text2);", ".fd-node .fd-rl { font-size: 12px; color: #b8b8b8;", ['theme']],
  ['減少動態不生效（拿掉兩道防護）', 'flow-designer.css', ["@media (prefers-reduced-motion: reduce) { .fd-root *, .fd-root *::before, .fd-root *::after { animation: none !important; transition: none !important; } }", "@media (prefers-reduced-motion: no-preference) {"], ["", "@media all {"], ['motion']],
  ['草稿名稱改用 innerHTML 直接插入（XSS 變異）', 'flow-designer.js', "<td><b>${d.name}</b>", "<td><b>${raw(d.name)}</b>", ['xss']],
];
async function runMutants() {
  console.log('[變異測試]');
  for (const [name, file, find, rep, groups] of MUTANTS) {
    const src = SRC(file);
    let mutated = src, allOk = true;
    [].concat(find).forEach((fd, k) => { const idx = mutated.indexOf(fd); const one = idx >= 0 && mutated.indexOf(fd, idx + 1) < 0; ok(one, '變異點唯一存在：' + name, idx < 0 ? '找不到' : '出現多次'); if (!one) { allOk = false; return; } mutated = mutated.slice(0, idx) + [].concat(rep)[k] + mutated.slice(idx + fd.length); });
    if (!allOk) continue;
    const before = fail, passB = pass, fl = failures.length;
    quiet = true;
    fx.seed();
    try {
      await fx.startServer({ QUOTE_FLOW_DESIGNER: '1' });
      await withBrowser({ cdpPort: CDP, maxMs: 200000 }, async (b) => {
        b.replace[file] = mutated; await b.interceptOn();
        ok(await login(b, '_s_admin'), '(mut)登入');
        for (const g of groups) { if (!(await openDesigner(b, 1440, 1000))) { ok(false, '(mut)開啟'); continue; } try { await G[g](b); } catch (e) { ok(false, '(mut)例外 ' + e.message); } }
      });
    } finally { await fx.stopServer(); fx.cleanup(); }
    quiet = false;
    const red = fail > before;
    pass = passB; fail = before; failures.length = fl;
    ok(red, '變異必須使前端測試變紅：' + name);
  }
}

(async () => {
  try {
    if (want('static')) await G.static();
    if (want('gate')) await gateOff();
    const names = ORDER.filter(want).concat(SHOTS && ONLY.includes('shots') && !ORDER.includes('shots') ? ['shots'] : []).concat(want('gate') ? ['gate'] : []);
    if (names.length) await runGroups(names);
    if (process.env.FD_UI_MUT === '1') await runMutants();
  } catch (e) {
    fail++; failures.push('例外：' + (e && e.stack || e));
    console.log('例外：', e && e.stack || e);
  } finally {
    await fx.stopServer(); fx.cleanup();
  }
  const left = require('./_fd-cdp').withBrowser.leftover || 0;
  console.log(`\n簽核流程設計器 M2 前端檢查結果：${pass} 通過，${fail} 失敗（殘留 Edge 程序：${left}）`);
  if (fail) console.log(failures.map((f) => ' - ' + f).join('\n'));
  process.exit(fail ? 1 : 0);
})();
