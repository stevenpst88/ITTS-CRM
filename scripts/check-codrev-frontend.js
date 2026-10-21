#!/usr/bin/env node
/**
 * 程式碼審查「前端與無障礙」修正（PR-1）— 前端互動檢查（headless Edge ＋ CDP，真實滑鼠／鍵盤事件）＋ 變異測試。
 * 用法：node scripts/check-codrev-frontend.js            （全部；埠 PR1_PORT 預設 3190、CDP 埠 PR1_CDP 預設 9491）
 *       PR1_ONLY=a1,a3,…   只跑指定群組（a1 a1b a2 a3 a4 a5 a6 a7 a8 a9 static；b09 含在 a6）
 *       PR1_MUT=1          另外跑變異測試（把前端檔換成改壞的版本，對應群組必須變紅）
 *       PR1_SHOTS=<資料夾> 另外輸出截圖
 * 沙盒：臨時 _s_* 帳號與資料檔只存在於本次執行，結束以 PID 清掉伺服器與 Edge；不連遠端資料庫；不含任何帳密。
 *   A1 閒置警告「立即登出」（用 Date.now 偏移法觸發警告，真點按鈕）　A2 矮螢幕　A3 兩人覆核視窗（深色對比、焦點、Esc、陷阱、錯誤處理）
 *   A4 登入頁 #done aria-hidden　A5 季度覆核 label　A6 .fd-note 換行　A7 reduced-motion　A8 showToast　A9 名單渲染（新舊 DOM 完全相同＋計時）
 *   B-09 flow-designer 鎖定說明改中性文字
 */
'use strict';
process.env.FD_PORT = process.env.PR1_PORT || '3190';
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const fx = require('./_fd-fixture');
const { withBrowser } = require('./_fd-cdp');

const ROOT = fx.ROOT;
const CDP = parseInt(process.env.PR1_CDP || '9491', 10);
const ONLY = (process.env.PR1_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const SHOTS = process.env.PR1_SHOTS || '';
fx.guard();

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = fx.sleep;
const SRC = (f) => fs.readFileSync(path.join(ROOT, '_client', f), 'utf8');

async function login(b, who) {
  await b.nav(fx.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${JSON.stringify(who)},password:${JSON.stringify(fx.PW)}})}).then(r=>r.status)`);
  return st === 200;
}
async function openApp(b, w, h) {
  if (w) await b.viewport(w, h || 900, w < 700);
  ok(await login(b, '_s_own'), '登入使用者端');
  await b.nav(fx.BASE + '/');
  return b.waitFor('window.ITTSIdle && ITTSIdle._state().lastTouch > 0 && typeof renderContacts === "function"', 20000);
}
async function openAdmin(b, w, h) {
  if (w) await b.viewport(w, h || 1000, w < 700);
  ok(await login(b, '_s_admin'), '登入後台');
  await b.nav(fx.BASE + '/admin.html');
  return b.waitFor('window.ITTSIdle && ITTSIdle._state().lastTouch > 0', 20000);
}
/** 把 Date.now 往前撥 ms，讓閒置計時以為已閒置那麼久（不必真的等） */
/** 觸發警告：頁面自己偶爾會發 scroll（側欄群組收合動畫的 .sg-inner）而重置閒置計時，所以撥完時鐘若沒出現警告就再多撥一輪 */
async function triggerWarn(b) {
  for (let i = 1; i <= 5; i++) {
    await b.ev(`window.__skewTot = (window.__skewTot || 0) + 28.5 * 60000; 0`);
    await skewNow(b, await b.ev('window.__skewTot'));
    if (await b.waitFor(`getComputedStyle(document.getElementById('idleTimeoutOverlay')).display==='flex'`, 3500)) return true;
  }
  return false;
}
const skewNow = (b, ms) => b.ev(`(()=>{ if(!window.__rn){ window.__rn = Date.now.bind(Date); } Date.now = () => window.__rn() + ${ms}; return true })()`);
const shotTo = async (b, name) => { if (SHOTS) await b.shot(path.join(SHOTS, name)); };
const idleShown = (b) => b.ev(`(()=>{const o=document.getElementById('idleTimeoutOverlay');return !!o && getComputedStyle(o).display==='flex'})()`);
const auditHas = (needle) => fx.FILES.some((f) => { try { return fs.readFileSync(f, 'utf8').includes(needle); } catch { return false; } });

// 對比：取元素的文字色與「實際底色」（逐層疊合半透明背景）
const CONTRAST_FN = `(function(backdrop){
  const parse=c=>{const m=/rgba?\\(([^)]+)\\)/.exec(c);if(!m)return null;const p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number);return {r:p[0],g:p[1],b:p[2],a:p[3]===undefined?1:p[3]};};
  const over=(top,bot)=>({r:top.r*top.a+bot.r*(1-top.a),g:top.g*top.a+bot.g*(1-top.a),b:top.b*top.a+bot.b*(1-top.a),a:1});
  const lum=c=>{const f=v=>{v/=255;return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)};return 0.2126*f(c.r)+0.7152*f(c.g)+0.0722*f(c.b)};
  const ratio=(a,b)=>{const l1=lum(a),l2=lum(b);return (Math.max(l1,l2)+0.05)/(Math.min(l1,l2)+0.05)};
  const bgOf=el=>{const layers=[];let e=el;while(e&&e.nodeType===1){const c=parse(getComputedStyle(e).backgroundColor);if(c&&c.a>0){layers.push(c);if(c.a>=0.999)break;}e=e.parentElement;}
    let base=layers.length&&layers[layers.length-1].a>=0.999?layers.pop():parse(backdrop);let cur=base;while(layers.length){cur=over(layers.pop(),cur);}return cur;};
  return function(sel){const out=[];document.querySelectorAll(sel).forEach((el,i)=>{if(i>3)return;const r=el.getBoundingClientRect();if(!r.width||!r.height)return;const fg=parse(getComputedStyle(el).color);const bg=bgOf(el);const fgc=fg.a<1?over(fg,bg):fg;out.push({sel,ratio:+ratio(fgc,bg).toFixed(2),text:el.textContent.trim().slice(0,16)});});return out;};
})`;

const G = {};
// 公開 repo 不放人名：以拼接建立「曾經出現在說明文字裡的人名」的比對式
const NAME_RE = new RegExp(['Ste', 'ven'].join(''), 'i');

// ───────── A1＋A2：閒置警告 ─────────
G.a1 = async (b) => {
  console.log('  [a1] 閒置警告：立即登出／繼續使用／Esc／焦點');
  ok(await openApp(b, 1280, 900), 'a1 使用者端載入並完成第一次 touch');
  // 先把焦點放在一個「正在填的欄位」，之後 Esc 要還原
  await b.ev(`(()=>{const i=document.createElement('input');i.id='__prevField';i.style.cssText='position:fixed;left:8px;top:8px;z-index:1';document.body.appendChild(i);i.focus();return document.activeElement===i})()`);
  ok(await triggerWarn(b), 'a1 閒置 28.5 分鐘後出現警告視窗');
  const attrs = await b.ev(`(()=>{const m=document.querySelector('#idleTimeoutOverlay .modal');return {role:m.getAttribute('role'),modal:m.getAttribute('aria-modal'),lb:m.getAttribute('aria-labelledby'),lbText:(document.getElementById(m.getAttribute('aria-labelledby'))||{}).textContent,active:document.activeElement&&document.activeElement.id}})()`);
  eq(attrs.role, 'alertdialog', 'a1 警告視窗 role=alertdialog');
  eq(attrs.modal, 'true', 'a1 警告視窗 aria-modal=true');
  ok(/閒置/.test(attrs.lbText || ''), 'a1 aria-labelledby 指向標題', JSON.stringify(attrs));
  eq(attrs.active, 'idleStayBtn', 'a1 警告出現時焦點移到「繼續使用」');
  await shotTo(b, 'a1-idle-warning.png');
  // 視窗內捲動（wheel）不該收起視窗
  const rs = await b.rectOf('#idleStayBtn');
  await b.send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: rs.cx, y: rs.cy, deltaX: 0, deltaY: 40 });
  await sleep(250);
  ok(await idleShown(b), 'a1 在警告視窗內捲動滾輪：警告不被收掉、不續命');
  // Tab 陷阱：Tab 兩次再 Tab，焦點只在兩顆按鈕之間
  await b.key('Tab'); await b.key('Tab'); await b.key('Tab');
  ok(await b.ev(`document.getElementById('idleTimeoutOverlay').contains(document.activeElement)`), 'a1 Tab 焦點不會跑出警告視窗');
  await b.key('Tab', 8);
  ok(await b.ev(`document.getElementById('idleTimeoutOverlay').contains(document.activeElement)`), 'a1 Shift+Tab 焦點不會跑出警告視窗');
  ok(await idleShown(b), 'a1 Tab 操作不收起警告');
  // 真滑鼠按「立即登出」：pointerdown 之後視窗仍在（以前這時就被收掉了），放開後呼叫 /api/logout 並導向登入頁
  const rl = await b.rectOf('#idleLogoutBtn');
  await b.move(rl.cx, rl.cy); await b.down(rl.cx, rl.cy); await sleep(150);
  ok(await idleShown(b), 'a1 按下「立即登出」（pointerdown）後警告視窗仍在');
  await b.up(rl.cx, rl.cy);
  ok(await b.waitFor(`location.pathname==='/login.html'`, 8000), 'a1 放開後導向 /login.html');
  ok(auditHas('"LOGOUT"') || auditHas('LOGOUT'), 'a1 伺服器留下 LOGOUT 稽核（代表 /api/logout 真的被呼叫）');
  await sleep(300);
  const touch = await b.ev(`fetch('/api/session/touch',{cache:'no-store'}).then(r=>r.status)`);
  eq(touch, 401, 'a1 登出後 session 已失效（touch → 401）');

  // ── Esc＝繼續使用，且焦點還原 ──
  ok(await openApp(b, 1280, 900), 'a1 重新登入並載入');
  await b.ev(`(()=>{const i=document.createElement('input');i.id='__prevField';i.style.cssText='position:fixed;left:8px;top:8px;z-index:1';document.body.appendChild(i);i.focus();return true})()`);
  ok(await triggerWarn(b), 'a1 第二輪：警告出現');
  const t0 = await b.ev(`ITTSIdle._state().lastActivity`);
  await b.key('Escape');
  await sleep(300);
  ok(!(await idleShown(b)), 'a1 Esc 關閉警告（＝繼續使用）');
  ok((await b.ev(`ITTSIdle._state().lastActivity`)) > t0, 'a1 Esc 之後閒置計時已重置');
  eq(await b.ev(`document.activeElement && document.activeElement.id`), '__prevField', 'a1 關閉後焦點還原到原本的欄位');
  // 「繼續使用」按鈕
  ok(await triggerWarn(b), 'a1 第三輪：警告再度出現');
  await b.click('#idleStayBtn');
  await sleep(300);
  ok(!(await idleShown(b)), 'a1 按「繼續使用」關閉警告');
  ok((await b.ev(`location.pathname`)) === '/', 'a1 「繼續使用」不會登出');
};

// 後台（admin.html）沒有 #idleTimeoutOverlay，走簡易提示：同樣要「按鈕按得到、Esc 繼續使用、焦點進去」
G.a1b = async (b) => {
  console.log('  [a1b] 後台簡易閒置提示');
  ok(await openAdmin(b, 1280, 900), 'a1b 後台載入');
  const shown = async () => b.waitFor("!!document.querySelector('body > [role=alertdialog]')", 3500);
  let ok1 = false;
  for (let i = 1; i <= 5 && !ok1; i++) { await b.ev("window.__skewTot = (window.__skewTot || 0) + 28.5 * 60000; 0"); await skewNow(b, await b.ev('window.__skewTot')); ok1 = await shown(); }
  ok(ok1, 'a1b 簡易提示出現');
  const at = await b.ev("(()=>{const e=document.querySelector('body > [role=alertdialog]');return {label:e.getAttribute('aria-label'),active:document.activeElement===e.querySelector('button'),txt:e.textContent}})()");
  ok(/閒置/.test(at.label || '') && at.active, 'a1b 簡易提示有 aria-label，焦點在「繼續使用」', JSON.stringify(at));
  const rb = await b.rectOf("body > [role=alertdialog] button");
  await b.move(rb.cx, rb.cy); await b.down(rb.cx, rb.cy); await sleep(150);
  ok(await b.ev("!!document.querySelector('body > [role=alertdialog]')"), 'a1b 在提示內 pointerdown 不會先把提示收掉');
  await b.up(rb.cx, rb.cy); await sleep(300);
  ok(!(await b.ev("!!document.querySelector('body > [role=alertdialog]')")), 'a1b 放開後（click＝繼續使用）提示關閉');
  eq(await b.ev("location.pathname"), '/admin.html', 'a1b 沒有被登出');
  // Esc
  let ok2 = false;
  for (let i = 1; i <= 5 && !ok2; i++) { await b.ev("window.__skewTot = (window.__skewTot || 0) + 28.5 * 60000; 0"); await skewNow(b, await b.ev('window.__skewTot')); ok2 = await shown(); }
  ok(ok2, 'a1b 第二輪提示出現');
  await b.key('Escape'); await sleep(300);
  ok(!(await b.ev("!!document.querySelector('body > [role=alertdialog]')")), 'a1b Esc 關閉提示（＝繼續使用）');
};

G.a2 = async (b) => {
  console.log('  [a2] 矮螢幕：警告視窗按鈕完整可見');
  for (const [w, h, tag] of [[360, 568, 'iphone-se'], [667, 375, 'landscape']]) {
    ok(await openApp(b, w, h), `a2 ${tag} 載入`);
    ok(await triggerWarn(b), `a2 ${tag} 警告出現`);
    await sleep(200);
    const m = await b.ev(`(()=>{const mb=document.querySelector('#idleTimeoutOverlay .modal-body');const cs=getComputedStyle(mb);const r=document.getElementById('idleLogoutBtn').getBoundingClientRect();const r2=document.getElementById('idleStayBtn').getBoundingClientRect();return {maxH:cs.maxHeight,overflowY:cs.overflowY,sh:mb.scrollHeight,ch:mb.clientHeight,btnBottom:Math.max(r.bottom,r2.bottom),vh:innerHeight,btnTop:Math.min(r.top,r2.top)}})()`);
    ok(m.maxH === 'none', `a2 ${tag} .modal-body 不限高度`, JSON.stringify(m));
    ok(m.sh <= m.ch + 1, `a2 ${tag} 內容不需要在框內捲動（scrollHeight ${m.sh} <= clientHeight ${m.ch}）`, JSON.stringify(m));
    ok(m.btnTop >= 0 && m.btnBottom <= m.vh, `a2 ${tag} 兩顆按鈕完整落在畫面內（${Math.round(m.btnTop)}–${Math.round(m.btnBottom)} / ${m.vh}）`, JSON.stringify(m));
    await shotTo(b, `a2-idle-${tag}.png`);
  }
  await b.viewport(1280, 900, false);
};

// ───────── A3：兩人覆核視窗 ─────────
const CR_SETUP = `(()=>{
  window.__crCalls = [];
  window.__crMode = 'ok';
  const reqs = [
    { id: 'req-aaaa-1111', status: 'pending', kind: 'flow', kindLabel: '簽核流程設定', proposedBy: '_s_a1', proposedAt: new Date().toISOString(), expiresAt: new Date(Date.now()+3*864e5).toISOString(), diffSummary: ['一級主管：啟用 → 停用', 'ITS 事業部：需二級主管'], history: [{ at: new Date().toISOString(), by: '_s_a1', action: 'propose', comment: '調整流程' }], isMine: false, baseChanged: true },
    { id: 'req-bbbb-2222', status: 'pending', kind: 'roster', kindLabel: '簽核名冊', proposedBy: '_s_a2', proposedAt: new Date().toISOString(), expiresAt: new Date(Date.now()+3*864e5).toISOString(), diffSummary: ['總經理名冊新增 1 人'], history: [], isMine: false },
    { id: 'req-cccc-3333', status: 'pending', kind: 'flow', kindLabel: '簽核流程設定', proposedBy: '_s_admin', proposedAt: new Date().toISOString(), expiresAt: new Date(Date.now()+3*864e5).toISOString(), diffSummary: ['自己提的'], history: [], isMine: true },
    { id: 'req-dddd-4444', status: 'approved', kind: 'flow', kindLabel: '簽核流程設定', proposedBy: '_s_a1', proposedAt: new Date().toISOString(), decidedBy: '_s_a2', decidedAt: new Date().toISOString(), diffSummary: ['舊的差異'], history: [{ at: new Date().toISOString(), by: '_s_a2', action: 'approve', comment: '' }] },
    { id: 'req-eeee-5555', status: 'rejected', kind: 'roster', kindLabel: '簽核名冊', proposedBy: '_s_a1', proposedAt: new Date().toISOString(), decidedBy: '_s_a2', decidedAt: new Date().toISOString(), diffSummary: ['被退回的'], history: [{ at: new Date().toISOString(), by: '_s_a2', action: 'reject', comment: '不妥' }] },
  ];
  const api = async (method, path, body) => {
    window.__crCalls.push(method + ' ' + path);
    if (method === 'GET') return { ok: true, status: 200, data: { enabled: true, activeAdmins: 3, pendingCount: 3, requests: reqs, me: '_s_admin' } };
    if (window.__crMode === 'throw') throw new Error('network down');
    if (window.__crMode === 'fail') return { ok: false, status: 409, data: { error: '申請已被處理' } };
    return { ok: true, status: 200, data: {} };
  };
  let h = document.getElementById('__crHost'); if (h) h.remove();
  h = document.createElement('div'); h.id = '__crHost'; h.style.cssText = 'position:fixed;left:10px;bottom:10px;z-index:5';
  document.body.appendChild(h);
  window.__cr = CfgReview.create({ api });
  return window.__cr.refresh().then(() => { window.__cr.paint(h); return true; });
})()`;
const setMode = (b, mode) => b.ev(`(()=>{document.body.classList.toggle('glass', ${mode.glass}); document.body.classList.toggle('dark', ${mode.dark}); return document.body.className})()`);

G.a3 = async (b) => {
  console.log('  [a3] 兩人覆核視窗：深色對比／焦點／Esc／陷阱／錯誤處理');
  ok(await openAdmin(b, 1280, 1000), 'a3 後台載入');
  await b.waitFor('typeof CfgReview !== "undefined"', 10000);
  ok(await b.ev(CR_SETUP), 'a3 建立偽造 api 的 CfgReview');
  const cls0 = await b.ev('document.body.className');
  ok(/glass/.test(cls0) && /dark/.test(cls0), 'a3 後台預設就是深色玻璃（body.glass.dark）', cls0);

  // 外觀：深色玻璃／淺色玻璃／非玻璃
  const modes = [['dark-glass', { glass: true, dark: true }], ['light-glass', { glass: true, dark: false }], ['plain-light', { glass: false, dark: false }]];
  const res = {};
  for (const [tag, mode] of modes) {
    await setMode(b, mode);
    await b.ev(`window.__cr.closeModal(); document.querySelector('#__crHost [data-cr="open"]').click(); true`);
    await b.waitFor(`!!document.querySelector('.cr-modal-box .cr-card')`, 4000);
    await sleep(250);
    res[tag] = await b.ev(`(${CONTRAST_FN})('#0d1117')('.cr-title, .cr-modal-box .cr-card b, .cr-modal-box .cr-card li, .cr-modal-box summary, .cr-modal-box .cr-close')`);
    const boxBg = await b.ev(`getComputedStyle(document.querySelector('.cr-modal-box')).backgroundColor`);
    res[tag + '_bg'] = boxBg;
    res[tag + '_banner'] = await b.ev(`(${CONTRAST_FN})('#0d1117')('.cr-banner')`);
    res[tag + '_muted'] = await b.ev(`(${CONTRAST_FN})('#0d1117')('.cr-modal-box div[style*="6b7280"], .cr-modal-box [data-cr-msg]')`);
    await shotTo(b, `a3-modal-${tag}.png`);
  }
  const minOf = (arr) => arr.length ? Math.min(...arr.map((x) => x.ratio)) : 0;
  ok(res['dark-glass'].length >= 4, 'a3 深色玻璃：量到足夠的文字元素', JSON.stringify(res['dark-glass']));
  ok(minOf(res['dark-glass']) >= 4.5, `a3 深色玻璃：標題／卡片標題／差異摘要／關閉鈕對比 >= 4.5（最低 ${minOf(res['dark-glass'])}）`, JSON.stringify(res['dark-glass']));
  ok(minOf(res['dark-glass_muted']) >= 4.5 || res['dark-glass_muted'].length === 0, `a3 深色玻璃：灰色說明文字對比 >= 4.5（最低 ${minOf(res['dark-glass_muted'])}）`, JSON.stringify(res['dark-glass_muted']));
  for (const tag of ['dark-glass', 'light-glass', 'plain-light']) ok(res[tag + '_banner'].length === 1 && res[tag + '_banner'][0].ratio >= 4.5, `a3 ${tag}：「待覆核變更 N 筆」橫幅對比 >= 4.5（${JSON.stringify(res[tag + '_banner'])}）`);
  ok(res['dark-glass_bg'] !== 'rgb(255, 255, 255)', 'a3 深色玻璃：外框不再是純白底', res['dark-glass_bg']);
  eq(res['light-glass_bg'], 'rgb(255, 255, 255)', 'a3 淺色玻璃：外框仍是白底（外觀不變）');
  eq(res['plain-light_bg'], 'rgb(255, 255, 255)', 'a3 非玻璃：外框仍是白底（外觀不變）');
  ok(minOf(res['light-glass']) >= 7, `a3 淺色玻璃：文字對比 >= 7（最低 ${minOf(res['light-glass'])}）`, JSON.stringify(res['light-glass']));
  ok(minOf(res['plain-light']) >= 7, `a3 非玻璃：文字對比 >= 7（最低 ${minOf(res['plain-light'])}）`, JSON.stringify(res['plain-light']));

  // 行為（深色玻璃下測）
  await setMode(b, { glass: true, dark: true });
  await b.ev(`window.__cr.closeModal(); true`);
  await b.ev(`(()=>{const o=document.querySelector('#__crHost [data-cr="open"]'); o.focus(); return document.activeElement===o})()`);
  await b.ev(`document.querySelector('#__crHost [data-cr="open"]').click(); true`);
  await b.waitFor(`!!document.querySelector('.cr-modal-box .cr-card')`, 4000);
  await sleep(300);
  const a = await b.ev(`(()=>{const o=document.querySelector('.cr-overlay');return {role:o.getAttribute('role'),modal:o.getAttribute('aria-modal'),lb:o.getAttribute('aria-labelledby'),lbText:(document.getElementById(o.getAttribute('aria-labelledby'))||{}).textContent,inside:o.contains(document.activeElement),active:document.activeElement.tagName+':'+(document.activeElement.getAttribute('data-cr')||'')}})()`);
  eq(a.role, 'dialog', 'a3 role=dialog'); eq(a.modal, 'true', 'a3 aria-modal=true');
  ok(/覆核/.test(a.lbText || ''), 'a3 aria-labelledby 指向標題', JSON.stringify(a));
  ok(a.inside, '開啟後焦點進入彈窗內', JSON.stringify(a));
  // Tab 陷阱
  const nFocus = await b.ev(`(()=>{const o=document.querySelector('.cr-overlay');return Array.from(o.querySelectorAll('button, input, summary')).filter(e=>!e.disabled&&e.offsetParent!==null).length})()`);
  let leaked = false;
  for (let i = 0; i < nFocus + 3; i++) { await b.key('Tab'); if (!(await b.ev(`document.querySelector('.cr-overlay').contains(document.activeElement)`))) leaked = true; }
  ok(!leaked, `a3 連按 Tab ${nFocus + 3} 次焦點不會離開彈窗`);
  for (let i = 0; i < 3; i++) { await b.key('Tab', 8); if (!(await b.ev(`document.querySelector('.cr-overlay').contains(document.activeElement)`))) leaked = true; }
  ok(!leaked, 'a3 Shift+Tab 焦點不會離開彈窗');

  // 失敗處理：api 丟例外 → 按鈕不卡 disabled、顯示訊息；另一張卡的「意見」不被清掉
  await b.ev(`window.__crMode='throw'; true`);
  await b.ev(`(()=>{const i=document.querySelector('[data-cr-comment="req-bbbb-2222"]'); i.value='第二張卡的意見'; return true})()`);
  await b.click('[data-cr="approve"][data-id="req-aaaa-1111"]');
  await sleep(500);
  const afterThrow = await b.ev(`(()=>{const btn=document.querySelector('[data-cr="approve"][data-id="req-aaaa-1111"]');const m=document.querySelector('[data-cr-msg="req-aaaa-1111"]');return {disabled:btn.disabled,msg:m.textContent,comment:document.querySelector('[data-cr-comment="req-bbbb-2222"]').value}})()`);
  ok(afterThrow.disabled === false, 'a3 api 丟例外後按鈕沒有卡在 disabled', JSON.stringify(afterThrow));
  ok(/失敗|異常/.test(afterThrow.msg), 'a3 api 丟例外後顯示錯誤訊息', JSON.stringify(afterThrow));
  eq(afterThrow.comment, '第二張卡的意見', 'a3 失敗後重畫彈窗，另一張卡已輸入的意見仍保留');
  await shotTo(b, 'a3-modal-error-dark.png');
  // 失敗（回 ok:false）同樣保留
  await b.ev(`window.__crMode='fail'; true`);
  await b.click('[data-cr="reject"][data-id="req-aaaa-1111"]');   // 沒填原因 → 前端擋
  await sleep(200);
  ok(/原因/.test(await b.ev(`document.querySelector('[data-cr-msg="req-aaaa-1111"]').textContent`)), 'a3 退回沒填原因 → 前端提示');
  await b.click('[data-cr="approve"][data-id="req-aaaa-1111"]');
  await sleep(500);
  ok(/已被處理/.test(await b.ev(`document.querySelector('[data-cr-msg="req-aaaa-1111"]').textContent`)), 'a3 伺服器回 ok:false 顯示伺服器訊息');
  eq(await b.ev(`document.querySelector('[data-cr-comment="req-bbbb-2222"]').value`), '第二張卡的意見', 'a3 ok:false 後另一張卡的意見仍保留');

  // Esc 關閉＋焦點還原
  await b.key('Escape');
  await sleep(250);
  ok(!(await b.ev(`!!document.querySelector('.cr-overlay')`)), 'a3 Esc 關閉彈窗');
  eq(await b.ev(`document.activeElement.getAttribute('data-cr')`), 'open', 'a3 關閉後焦點回到「覆核」按鈕');
  // 再開再用 × 關閉也要還原；關掉後不殘留 keydown 監聽（再按 Esc 不會出錯）
  await b.ev(`document.querySelector('#__crHost [data-cr="open"]').focus(); document.querySelector('#__crHost [data-cr="open"]').click(); true`);
  await b.waitFor(`!!document.querySelector('.cr-overlay')`, 3000);
  await b.click('.cr-close');
  await sleep(200);
  ok(!(await b.ev(`!!document.querySelector('.cr-overlay')`)), 'a3 按 × 關閉');
  await b.key('Escape');
};

// ───────── A4 ─────────
G.a4 = async (b) => {
  console.log('  [a4] 登入頁 #done');
  await b.media([{ name: 'prefers-reduced-motion', value: 'reduce' }]);
  await b.nav(fx.BASE + '/login.html?motion=reduce');
  ok(await b.waitFor('!!document.getElementById("done")', 8000), 'a4 登入頁載入');
  eq(await b.ev(`document.getElementById('done').getAttribute('aria-hidden')`), 'true', 'a4 平常 #done aria-hidden=true');
  // 用真的登入表單走一遍（導向之前 #done 會短暫顯示）：以 MutationObserver 把狀態記在 sessionStorage（跨導向保留）
  await b.ev(`(()=>{const d=document.getElementById('done');const log=[];new MutationObserver(()=>{log.push({show:d.classList.contains('show'),hidden:d.getAttribute('aria-hidden')});try{sessionStorage.setItem('__doneLog',JSON.stringify(log))}catch(e){}}).observe(d,{attributes:true});
    document.getElementById('username').value='_s_own';document.getElementById('password').value=${JSON.stringify(fx.PW)};return true})()`);
  await b.click('#loginBtn');
  ok(await b.waitFor(`location.pathname==='/'`, 10000), 'a4 登入成功並導向');
  await sleep(300);
  const log = JSON.parse((await b.ev(`sessionStorage.getItem('__doneLog')||'[]'`)) || '[]');
  const shown = log.filter((x) => x.show);
  ok(shown.length > 0, 'a4 登入成功時 #done 加上 show', JSON.stringify(log));
  ok(shown.length > 0 && shown.every((x) => x.hidden === null), 'a4 顯示 show 的同時 aria-hidden 已移除', JSON.stringify(log));
  await b.media([]);
};

// ───────── A5 ─────────
G.a5 = async (b) => {
  console.log('  [a5] 季度覆核 label');
  ok(await openAdmin(b, 1280, 1000), 'a5 後台載入');
  const ids = ['arFlag', 'arRole', 'arQ', 'arQuarter', 'arComment', 'arExceptions'];
  const names = {};
  for (const id of ids) {
    const r = await b.ev(`(()=>{const e=document.getElementById(${JSON.stringify(id)});if(!e)return null;return {n:e.labels?e.labels.length:-1,t:e.labels&&e.labels[0]?e.labels[0].textContent.trim():''}})()`);
    ok(r && r.n === 1 && r.t.length > 0, `a5 ${id} 有 label（${r && r.t}）`, JSON.stringify(r));
    names[id] = r;
  }
  // 外觀不變：label 與「以前的 div 標題」在同一個容器、相同行內樣式下，計算後的樣式完全相同
  for (const [tag, mode] of [['dark-glass', { glass: true, dark: true }], ['light-glass', { glass: true, dark: false }], ['plain', { glass: false, dark: false }]]) {
    await setMode(b, mode);
    const diffs = await b.ev(`(()=>{const out=[];['arFlag','arRole','arQ','arQuarter','arComment','arExceptions'].forEach(id=>{
      const e=document.getElementById(id);const lab=e.labels[0];
      const ref=document.createElement('div');ref.setAttribute('style','font-size:12px;color:#888;margin-bottom:3px');ref.textContent=lab.textContent;lab.parentNode.insertBefore(ref,lab);
      const a=getComputedStyle(lab),r=getComputedStyle(ref);
      ['display','fontSize','fontWeight','color','marginBottom','marginTop','textTransform','letterSpacing','lineHeight','textAlign'].forEach(p=>{ if(a[p]!==r[p]) out.push(id+'.'+p+': label='+a[p]+' / div='+r[p]); });
      ref.remove();});return out})()`);
    eq(diffs, [], `a5 ${tag}：label 的計算樣式與原本的 div 標題完全相同`);
  }
  await setMode(b, { glass: true, dark: true });
};

// ───────── A6＋B-09 ─────────
G.a6 = async (b) => {
  console.log('  [a6/b09] 流程設計器：.fd-note 換行、鎖定說明');
  const open = async (w, h) => {
    await b.viewport(w, h, w < 700);
    await b.nav(fx.BASE + '/admin.html');
    if (!(await b.waitFor('window.FlowDesigner && FlowDesigner.state().enabled === true', 25000))) return false;
    await b.ev(`document.querySelector('.admin-nav-item[data-sec="flow-designer"]').click()`);
    return b.waitFor('FlowDesigner.state().inited && !!FlowDesigner.state().def && !!document.querySelector("#fdLanes .fd-lane")', 20000);
  };
  ok(await login(b, '_s_admin'), 'a6 登入');
  ok(await open(347, 800), 'a6 347px 開啟設計器');
  const m = await b.ev(`(()=>{const n=document.querySelector('.fd-note');const btn=document.getElementById('fdPropose');const r=btn.getBoundingClientRect();return {noteH:n.getBoundingClientRect().height,btnW:r.width,btnH:r.height,btnLeft:r.left,vw:innerWidth,wrap:getComputedStyle(n).flexWrap,sw:document.documentElement.scrollWidth}})()`);
  eq(m.wrap, 'wrap', 'a6 .fd-note flex-wrap=wrap');
  ok(m.btnH <= 50, `a6 347px：「提交覆核」按鈕沒有被擠成一字一行（高 ${Math.round(m.btnH)}px）`, JSON.stringify(m));
  ok(m.btnW >= 90, `a6 347px：按鈕寬度正常（${Math.round(m.btnW)}px）`, JSON.stringify(m));
  ok(m.sw <= m.vw + 1, 'a6 347px 沒有橫向捲動', JSON.stringify(m));
  await shotTo(b, 'a6-fd-note-347.png');
  ok(await open(1440, 1000), 'a6 1440px 開啟設計器');
  const d = await b.ev(`(()=>{const n=document.querySelector('.fd-note');const kids=Array.from(n.children).map(c=>Math.round(c.getBoundingClientRect().top));return {h:n.getBoundingClientRect().height,tops:kids}})()`);
  ok(d.h <= 70 && Math.max(...d.tops) - Math.min(...d.tops) <= 8, `a6 1440px：說明列仍是單行（高 ${Math.round(d.h)}px）`, JSON.stringify(d));
  await shotTo(b, 'a6-fd-note-1440.png');
  // B-09：鎖定說明不得含人名與內部決策備註
  const txt = await b.ev(`(()=>{const parts=[document.body.innerText];document.querySelectorAll('[title],[aria-label]').forEach(e=>{parts.push(e.getAttribute('title')||'');parts.push(e.getAttribute('aria-label')||'')});return parts.join('\\n')})()`);
  ok(!NAME_RE.test(txt) && !/董事長溝通|待董事長/.test(txt), 'b09 設計器畫面（含 title／aria-label）沒有人名與內部決策備註');
  ok(/簽核方式固定/.test(txt), 'b09 鎖定提示改為中性文字「簽核方式固定」');
  const titles = await b.ev(`Array.from(document.querySelectorAll('#fdLanes [title]')).map(e=>e.getAttribute('title')).filter(t=>/代核|簽核方式/.test(t))`);
  ok(titles.length > 0 && titles.every((t) => !NAME_RE.test(t) && !/溝通/.test(t)), 'b09 🔒 的 title 是中性說明', JSON.stringify(titles));
};

// ───────── A7 ─────────
G.a7 = async (b) => {
  console.log('  [a7] reduced-motion');
  ok(await openApp(b, 1280, 900), 'a7 使用者端載入');
  const probe = `(()=>{const out={};['z-badge-danger','zombie-alert-icon','bday-card-icon','ci-spinner'].forEach(c=>{const e=document.createElement('span');e.className=c;document.body.appendChild(e);out[c]=getComputedStyle(e).animationName;e.remove()});
    const m=document.createElement('button');m.className='mic-btn listening';document.body.appendChild(m);out.mic=getComputedStyle(m).animationName;m.remove();return out})()`;
  const normal = await b.ev(probe);
  ok(normal['z-badge-danger'] !== 'none' && normal['zombie-alert-icon'] !== 'none' && normal['bday-card-icon'] !== 'none' && normal.mic !== 'none', '一般模式：脈動／浮動／彈跳動畫存在', JSON.stringify(normal));
  await b.media([{ name: 'prefers-reduced-motion', value: 'reduce' }]);
  await sleep(150);
  const red = await b.ev(probe);
  eq([red['z-badge-danger'], red['zombie-alert-icon'], red['bday-card-icon'], red.mic], ['none', 'none', 'none', 'none'], 'a7 減少動態：四種無限動畫停止');
  ok(red['ci-spinner'] !== 'none', 'a7 減少動態：載入轉圈圈（狀態指示）保留', JSON.stringify(red));
  await b.media([]);
};

// ───────── A8 ─────────
G.a8 = async (b) => {
  console.log('  [a8] showToast');
  ok(await openApp(b, 1280, 900), 'a8 使用者端載入');
  await b.ev(`showToast('A', 400); true`);
  await sleep(250);
  await b.ev(`showToast('B', 1500); true`);
  await sleep(450);   // 已超過 A 的 400ms：舊寫法 A 的計時器會在這時把 B 收掉
  eq(await b.ev(`document.getElementById('toast').classList.contains('show')`), true, 'a8 連續 toast：前一個的計時器不會提早收掉後一個');
  eq(await b.ev(`document.getElementById('toast').textContent`), 'B', 'a8 顯示的是最後一個訊息');
  await sleep(1300);
  eq(await b.ev(`document.getElementById('toast').classList.contains('show')`), false, 'a8 最後一個訊息依自己的時間收掉');
};

// ───────── A9 ─────────
const SYN = `(()=>{
  let s = 12345; const rnd = () => (s = (s * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  const N_CO = 600, N_C = 1800, N_V = 5000, N_O = 2000, N_KA = 80;
  const cos = Array.from({length: N_CO}, (_, i) => '測試公司' + String(i).padStart(3, '0'));
  window.__contacts = Array.from({length: N_C}, (_, i) => ({ id: 'c' + i, name: '聯絡人' + i, company: i % 17 === 0 ? '' : cos[i % N_CO], title: i % 3 ? '經理' : '', isPrimary: i % 5 === 0, opportunityStage: i % 7 === 0 ? 'A' : (i % 11 === 0 ? 'B' : ''), industry: i % 4 ? '製造業' : '', phone: '02-0000' + i, email: 'u' + i + '@example.test', category: 'customer', createdAt: '2026-01-' + String(1 + (i % 27)).padStart(2, '0') }));
  window.__visits = Array.from({length: N_V}, (_, i) => ({ id: 'v' + i, contactId: 'c' + Math.floor(rnd() * N_C), visitDate: '2026-0' + (1 + Math.floor(rnd() * 9)) + '-' + String(1 + Math.floor(rnd() * 28)).padStart(2, '0'), topic: '議題' + (i % 50) }));
  window.__opps = Array.from({length: N_O}, (_, i) => ({ id: 'o' + i, company: cos[Math.floor(rnd() * N_CO)], stage: ['A', 'B', 'C', 'Won'][i % 4], product: '產品' + (i % 9), amount: i % 3 ? 100 * (i % 20) : 0, createdAt: '2026-0' + (1 + Math.floor(rnd() * 9)) + '-' + String(1 + Math.floor(rnd() * 28)).padStart(2, '0') }));
  // 平手情境：company5 的三位成員（c5、c605、c1205）同日拜訪，原始順序刻意是 c1205、c5、c605；company7 兩筆商機同時間（先 otie1 後 otie2）
  window.__visits.push({ id: 'vtie1', contactId: 'c1205', visitDate: '2026-12-31', topic: '平手A' }, { id: 'vtie2', contactId: 'c5', visitDate: '2026-12-31', topic: '平手B' }, { id: 'vtie3', contactId: 'c605', visitDate: '2026-12-31', topic: '平手C' });
  window.__opps.push({ id: 'otie1', company: cos[7], stage: 'B', product: '平手商機一', amount: 1, createdAt: '2026-12-31' }, { id: 'otie2', company: cos[7], stage: 'C', product: '平手商機二', amount: 2, createdAt: '2026-12-31' });
  window.__kas = Array.from({length: N_KA}, (_, i) => ({ id: 'k' + i, owner: 'someone', company: cos[(i * 7) % N_CO], note: i % 2 ? '備註' : '' }));
  window.__kas.push({ id: 'kdup', owner: 'other', company: cos[0], note: '重複' });   // 同公司兩筆：必須取第一筆
  window.__groups = [0, 1, 2, 3].map(g => ({ id: 'g' + g, name: '集團' + g, taxId: '', industry: '', memberCompanies: cos.slice(g * 10 + 1, g * 10 + 9) }));
  return { co: N_CO, c: N_C, v: N_V, o: N_O };
})()`;
const A9_RUN = `(()=>{
  allVisits = window.__visits; allOpportunities = window.__opps; allKeyAccounts = window.__kas; allGroups = window.__groups;
  currentSection = 'customers'; isSearchMode = false; contactsKaFilterMode = 'all';
  const list = document.getElementById('contactList');
  const out = {};
  const fnv = (str) => { let h = 0x811c9dc5; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193); } return (h >>> 0).toString(16) + ':' + str.length; };
  for (const mode of ['all', 'ka', 'nonka']) {
    contactsKaFilterMode = mode;
    let best = 1e9;
    for (let rep = 0; rep < 3; rep++) { const t0 = performance.now(); renderContacts(window.__contacts.map(c => Object.assign({}, c))); best = Math.min(best, performance.now() - t0); }   // 取三次最小（機器忙時較穩）
    const html = list.outerHTML; out[mode] = { ms: best, len: html.length, hash: fnv(html), nodes: list.querySelectorAll('*').length };
  }
  contactsKaFilterMode = 'all'; isSearchMode = true;
  renderContacts(window.__contacts.slice(0, 300).map(c => Object.assign({}, c)));
  { const html = list.outerHTML; out.search = { ms: 0, len: html.length, hash: fnv(html), nodes: list.querySelectorAll('*').length }; }
  isSearchMode = false;
  return out;
})()`;
async function a9Run(b, mutSrc) {
  if (mutSrc !== undefined) { b.replace['/app.js'] = mutSrc; await b.interceptOn(); }
  ok(await openApp(b, 1280, 900), 'a9 使用者端載入');
  await b.ev(SYN);
  const hasList = await b.ev(`!!document.getElementById('contactList')`);
  ok(hasList, 'a9 有 #contactList');
  const r = await b.ev(A9_RUN);
  return r;
}
G.a9 = async (b, oldSrc) => {
  console.log('  [a9] 名單渲染：新舊 DOM 相同＋計時');
  const nw = await a9Run(b);
  ok(nw.all.nodes > 20000, `a9 合成資料確實渲染出大量 DOM（${nw.all.nodes} 個節點、${nw.all.len} 字元）`);
  ok(nw.ka.len > 1000 && nw.ka.len < nw.all.len && nw.nonka.len < nw.all.len, 'a9 KA 篩選（只看 KA／排除 KA）確實生效');
  // 渲染結束後不得殘留 KA 索引；批次索引與線性掃描一致（同公司多筆取第一筆、查不到回 false）
  eq(await b.ev(`_kaBatch === null`), true, 'a9 渲染結束後 KA 批次索引已釋放（不會有過期資料）');
  const first = await b.ev(`(()=>{allKeyAccounts = window.__kas; const a = getKaRecord(allKeyAccounts[0].company); const started = _kaBatchStart(); const b2 = getKaRecord(allKeyAccounts[0].company); const h = isKeyAccount('不存在的公司'); const h2 = isKeyAccount(allKeyAccounts[0].company); _kaBatchEnd(started); return [a.id, b2.id, h, h2, _kaBatch === null]})()`);
  eq(first, ['k0', 'k0', false, true, true], 'a9 批次索引與線性掃描一致：同公司多筆取第一筆、查不到回 false、結束即釋放');
  const ties = await b.ev(`(()=>{allVisits = window.__visits; allOpportunities = window.__opps; allKeyAccounts = window.__kas; allGroups = window.__groups; contactsKaFilterMode = 'all'; isSearchMode = false; renderContacts(window.__contacts.map(c => Object.assign({}, c))); const t = document.getElementById('contactList').innerText; return [t.includes('平手A'), t.includes('平手B'), t.includes('平手C'), t.includes('平手商機一'), t.includes('平手商機二')]})()`);
  eq(ties, [true, false, false, true, false], 'a9 平手情境：同日拜訪取 allVisits 原始順序最前的一筆（平手A）、同時間商機取最前的一筆（平手商機一）');
  // 髒資料（商機沒有 createdAt、拜訪沒有 visitDate）不得丟例外；索引不會因此少渲染節點
  const dirty = await b.ev(`(()=>{ const o = window.__opps.map(x => Object.assign({}, x)); o.forEach((x, i) => { if (i % 41 === 0) delete x.createdAt; }); const v = window.__visits.map(x => Object.assign({}, x)); v.forEach((x, i) => { if (i % 37 === 0) delete x.visitDate; });
    allOpportunities = o; allVisits = v; contactsKaFilterMode = 'all'; isSearchMode = false; try { renderContacts(window.__contacts.map(c => Object.assign({}, c))); return document.querySelectorAll('#contactList *').length } catch (e) { return 'ERR ' + e.message } })()`);
  ok(typeof dirty === 'number' && dirty > 20000, 'a9 髒資料（缺 createdAt／visitDate）不丟例外、節點數正常', dirty);
  if (oldSrc) {
    // 同一份資料改用 origin/main 的 app.js 渲染，輸出必須逐字相同（用 hash 比對；兩邊各自在頁面內算）
    const old = await a9Run(b, oldSrc);
    for (const k of ['all', 'ka', 'nonka', 'search']) ok(old[k].hash === nw[k].hash, `a9 ${k}：新舊 renderContacts 輸出的 DOM 逐字相同（${old[k].len} 字元，${old[k].nodes} 節點）`, old[k].hash + ' vs ' + nw[k].hash);
    const sumOld = old.all.ms + old.ka.ms + old.nonka.ms, sumNew = nw.all.ms + nw.ka.ms + nw.nonka.ms, speed = sumOld / sumNew;
    console.log(`    計時（1800 名片／600 公司／5000 拜訪／2000 商機；三種篩選合計）：舊 ${sumOld.toFixed(0)} ms → 新 ${sumNew.toFixed(0)} ms（${speed.toFixed(2)}x）；單次 all：舊 ${old.all.ms.toFixed(0)} → 新 ${nw.all.ms.toFixed(0)} ms`);
    ok(speed >= 1.2, `a9 新寫法明顯較快（至少 1.2 倍；實測 ${speed.toFixed(2)} 倍）`, speed);
    delete b.replace['/app.js'];
  } else {
    ok(false, 'a9 缺少 origin/main 的 app.js 可供比對（需要 git）');
  }
};

// ───────── 靜態檢查 ─────────
G.static = async () => {
  console.log('  [static] 靜態檢查');
  for (const f of fs.readdirSync(path.join(ROOT, '_client')).filter((x) => x.endsWith('.js'))) {
    try { execFileSync(process.execPath, ['--check', path.join(ROOT, '_client', f)], { stdio: 'pipe' }); } catch (e) { ok(false, 'node --check ' + f, String(e.stderr).slice(0, 200)); }
  }
  ok(true, 'static 全部 _client/*.js 語法檢查');
  const sw = SRC('sw.js');
  ok(parseInt((sw.match(/SW_VERSION = 'itts-crm-v(\d+)'/) || [])[1] || '0', 10) >= 18, 'static sw.js 版號已提升（至少 v18）');
  const idle = SRC('idle-timeout.js');
  ok(/inWarn\(/.test(idle) && /Escape/.test(idle) && /prevFocus/.test(idle), 'static idle-timeout.js 含警告視窗內操作判斷／Esc／焦點還原');
  const cfg = SRC('config-review.js');
  ok(/aria-labelledby/.test(cfg) && /Escape/.test(cfg) && /try \{ r = await api/.test(cfg), 'static config-review.js 含 aria-labelledby／Esc／api try-catch');
  // 新增的程式碼不帶 inline script／on*／style=（以 git diff 相對 origin/main 檢查；沒有 git 或基準時略過）
  try {
    const diff = execFileSync('git', ['diff', 'origin/main', '--unified=0', '--', '_client'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
    const added = diff.split('\n').filter((l) => /^\+[^+]/.test(l)).map((l) => l.slice(1));
    const removed = diff.split('\n').filter((l) => /^-[^-]/.test(l)).map((l) => l.slice(1));
    const cnt = (arr, re) => arr.reduce((n, l) => n + (l.match(re) || []).length, 0);
    ok(cnt(added, /<script\b/g) <= cnt(removed, /<script\b/g), 'static 沒有新增 inline <script>');
    ok(cnt(added, /\son[a-z]+\s*=\s*["']/g) <= cnt(removed, /\son[a-z]+\s*=\s*["']/g), 'static 沒有新增 on*= 事件屬性');
    ok(cnt(added, /\sstyle\s*=\s*["']/g) <= cnt(removed, /\sstyle\s*=\s*["']/g), `static 沒有新增 style= 屬性（加 ${cnt(added, /\sstyle\s*=\s*["']/g)} / 減 ${cnt(removed, /\sstyle\s*=\s*["']/g)}）`);
  } catch (e) { console.log('    （略過 git diff 檢查：' + String(e.message).slice(0, 80) + '）'); }
};

// ═════════════ 變異測試 ═════════════
const MUTANTS = [
  // [名稱, 檔案, 找, 換成, 群組, 預期失敗名稱的正規表示式（變紅原因必須符合，避免因無關原因誤判）]
  ['A1：警告視窗內的操作仍會收起視窗（拿掉 inWarn 判斷）', 'idle-timeout.js', 'if (inWarn(ev && ev.target)) return;', '', ['a1'], /pointerdown|導向|LOGOUT|session|滾輪|Tab/],
  ['A1：Esc 不再是繼續使用', 'idle-timeout.js', "ev.key === 'Escape'", "ev.key === '__none__'", ['a1', 'a1b'], /Esc/],
  ['A3：拿掉深色玻璃視窗外框規則', 'glass.css', 'body.glass.dark .cr-modal-box {', 'body.glass.dark .cr-modal-boxX {', ['a3'], /深色玻璃/],
  ['A3：api 例外不處理（按鈕卡住）', 'config-review.js', 'catch (e) { r = null; }   // 網路例外', 'catch (e) { throw e; }   // 網路例外', ['a3'], /api 丟例外|disabled/],
  ['A3：Esc 不關閉彈窗', 'config-review.js', "if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); closeModal(); return; }", '', ['a3'], /Esc/],
  ['A3：重畫時不保留意見欄', 'config-review.js', "if (inp) inp.value = saved[id];", '', ['a3'], /意見/],
  ['A4：登入成功時不移除 aria-hidden', 'login-cinematic.js', "done.removeAttribute('aria-hidden');", '', ['a4'], /aria-hidden/],
  ['A8：toast 不清舊計時器', 'app.js', 'clearTimeout(showToast._t);', '', ['a8'], /toast/],
  ['A9：KA 索引永遠回 false', 'app.js', 'if (_kaBatch) return _kaBatch.has(companyName);', 'if (_kaBatch) return false;', ['a9'], /逐字相同|KA 篩選|批次索引/],
  ['A9：公司商機索引取錯（永遠空）', 'app.js', "(recentIdx.oppsByCompany.get(company) || [])", "([])", ['a9'], /逐字相同|平手/],
  ['A9：最近拜訪忽略原始順序（索引併入順序）', 'app.js', "mv.sort((x, y) => x[0] - y[0]);", "mv.sort((x, y) => y[0] - x[0]);", ['a9'], /逐字相同|平手/],
];
/** 等到埠真的空出來（上一支伺服器還沒完全結束時，新伺服器會綁不上埠，測試會誤打到舊的、帳號對不上而 401） */
async function waitPortFree() {
  for (let i = 0; i < 60; i++) {
    try { await fetch(fx.BASE + '/login.html', { signal: AbortSignal.timeout(1500) }); } catch { return true; }
    await sleep(500);
  }
  return false;
}
async function runMutants(oldApp) {   // MUTANTS 第 6 欄（選填）：正規表示式，變紅的原因必須含符合的失敗名稱，避免因無關原因（逾時等）誤判
  console.log('[變異測試]');
  for (const [name, file, find, rep, groups, expectRe] of MUTANTS) {
    const src = SRC(file);
    const idx = src.indexOf(find);
    const one = idx >= 0 && src.indexOf(find, idx + 1) < 0;
    ok(one, '變異點唯一存在：' + name, idx < 0 ? '找不到' : '出現多次');
    if (!one) continue;
    const mutated = src.slice(0, idx) + rep + src.slice(idx + find.length);
    const before = fail, passB = pass, fl = failures.length;
    quiet = true;
    fx.seed();
    try {
      await fx.startServer({ QUOTE_FLOW_DESIGNER: '1' });
      await withBrowser({ cdpPort: CDP, maxMs: 240000 }, async (b) => {
        b.replace[file] = mutated; await b.interceptOn();
        for (const g of groups) { try { await G[g](b, oldApp); } catch (e) { ok(false, '(mut)例外 ' + e.message); } }
      });
    } finally { await fx.stopServer(); fx.cleanup(); await waitPortFree(); }
    quiet = false;
    const newFails = failures.slice(fl);
    const red = fail > before && (!expectRe || newFails.some((n) => expectRe.test(n)));
    if (fail > before && expectRe && !red) console.log('    （變紅但原因不符預期：' + newFails.slice(0, 3).join(' | ').slice(0, 200) + '）');
    pass = passB; fail = before; failures.length = fl;
    ok(red, '變異必須使測試變紅：' + name);
  }
}

(async () => {
  let oldApp = null;
  try { oldApp = execFileSync('git', ['show', 'origin/main:_client/app.js'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }); } catch (e) { console.log('取不到 origin/main 的 app.js：', String(e.message).slice(0, 100)); }
  try {
    if (want('static')) await G.static();
    const names = ['a1', 'a1b', 'a2', 'a3', 'a4', 'a5', 'a6', 'a7', 'a8', 'a9'].filter(want);
    if (names.length) {
      fx.seed();
      await fx.startServer({ QUOTE_FLOW_DESIGNER: '1' });
      await withBrowser({ cdpPort: CDP, maxMs: 540000 }, async (b) => {
        for (const g of names) {
          try { await G[g](b, oldApp); } catch (e) { ok(false, `${g} 例外：${e.message}`); console.log(e.stack); }
          if (b.errors.length) { console.log('    頁面錯誤：', b.errors.slice(0, 3).join(' | ')); b.errors.length = 0; }
        }
      });
      console.log('  Edge 殘留程序：', withBrowser.leftover || 0);
      await fx.stopServer(); fx.cleanup(); await waitPortFree();
    }
    if (process.env.PR1_MUT === '1') await runMutants(oldApp);
  } catch (e) {
    fail++; failures.push('例外：' + (e && e.stack || e));
    console.log('例外：', e && e.stack || e);
  } finally {
    await fx.stopServer(); fx.cleanup();
  }
  console.log(`\n程式碼審查前端修正（PR-1）檢查結果：${pass} 通過，${fail} 失敗`);
  if (fail) console.log(failures.map((f) => ' - ' + f).join('\n'));
  process.exit(fail ? 1 : 0);
})();
