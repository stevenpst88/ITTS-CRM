'use strict';
/** _ma-reminders-ui.js 的測試區段（dual／bell／items／quote／look）。共用主檔的小工具。 */
const fs = require('fs');
const path = require('path');
const J = JSON.stringify;

module.exports = function makeSections2(X) {
  const { t, browse, openUser, openAdmin, gotoRules, pickOpt, typeInto, rowS, shotTo, txt, reqs, mkUsers, seedData, U, tp, XSS1, XSS2, AMT, CNO, MAP, THEMES, theme, contrastSel, PAGE_HELPERS, agoIso, NOW } = X;
  const ROOT = path.resolve(__dirname, '..');
  const slash = (s) => String(s || '').replace(/-/g, '/');
  const writeData = (fn) => { const f = path.join(ROOT, 'data.json'); const d = JSON.parse(fs.readFileSync(f, 'utf8')); fn(d); fs.writeFileSync(f, J(d, null, 2)); };
  const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
  const pollIds = (H, who) => H.http(who, 'GET', '/api/poll-bundle').then((r) => ((r.json && r.json.maReminders) || []).map((x) => x.itemId || x.id));
  const shown = (b) => b.ev(`JSON.stringify([...document.querySelectorAll('#notifList [data-marem-row]')].map(r=>r.getAttribute('data-marem-row')))`).then(JSON.parse);
  const openBell = async (b) => { await b.ev(`(()=>{const d=document.getElementById('notifDropdown'); if(d.style.display==='none') document.getElementById('notifBellBtn').click(); return 0})()`); await b.sleep(700); };
  const closeBell = async (b) => { await b.ev(`document.getElementById('notifDropdown').style.display='none'; 0`); };
  const refreshBell = async (b) => { await b.ev(`_refreshBell(); 0`); await b.sleep(1100); };
  async function toItems(b, opts) {
    await b.ev(`showSection('maItems'); 0`);
    await b.waitFor(`document.querySelectorAll('#maItemsView tbody tr').length>0`, 12000);
    if (!(opts && opts.fast)) await b.sleep(1800);   // 等「下一個提醒」逐列載入
  }
  const rowHas = (b, co) => b.ev(`(()=>{const r=[...document.querySelectorAll('#maItemsView tbody tr[data-item-id]')].find(x=>x.innerText.indexOf(${J(co)})>=0); return r?r.getAttribute('data-item-id'):null})()`);
  const rowText = (b, id) => b.ev(`(()=>{const r=document.querySelector('#maItemsView tbody tr[data-item-id="${id}"]'); return r?r.innerText.replace(/\\s+/g,' '):null})()`);
  const btnIn = (id, label) => `()=>{const r=document.querySelector('#maItemsView tbody tr[data-item-id="${id}"]'); return r?([...r.querySelectorAll('button')].find(x=>x.textContent.trim()===${J(label)})||null):null}`;
  const topBtn = (txtv) => `()=>{const o=[...document.querySelectorAll('.mat-overlay')].pop(); return o?([...o.querySelectorAll('button')].find(x=>x.textContent.trim()===${J(txtv)}&&!x.disabled)||null):null}`;
  const dlgText = (b) => b.ev(`(()=>{const o=[...document.querySelectorAll('.mat-overlay')].pop(); return o?o.innerText:''})()`);

  // ═════════════════════════ dual：雙人覆核 ═════════════════════════
  async function secDual(C) {
    const { H } = C;
    await H.fresh(mkUsers(H), seedData(), { env: { MA_TRACKING: '1', APPROVAL_CONFIG_DUAL_REVIEW: '1' } });
    await H.loginAll([U.adm, U.adm2]);
    await browse(C, 400000, async (b) => {
      await openAdmin(b, C, U.adm, 1440, 1000, true);
      await gotoRules(b);
      await b.sleep(900);
      const chip = await b.ev(`JSON.stringify({chip:document.querySelector('#marCrHost [data-cr-chip]')?document.querySelector('#marCrHost [data-cr-chip]').getAttribute('data-cr-chip'):null, save:document.querySelector('#marAdminHost .mar-a-btn-primary').textContent})`).then(JSON.parse);
      t('DU-1 雙人覆核啟用：晶片顯示「兩人覆核：啟用」、儲存鈕改成「提出變更（待覆核）」', chip.chip === 'on' && /待覆核/.test(chip.save), J(chip));
      await typeInto(b, rowS(0) + ' .mar-a-n', '75');
      await b.click('#marAdminHost .mar-a-btn-primary');
      await b.waitFor(`/待另一位管理員覆核/.test(document.querySelector('#marAdminHost .mar-a-status').textContent)`, 8000);
      const st = await txt(b, '#marAdminHost .mar-a-status');
      const put = (await reqs(b)).find((r) => r.m === 'PUT' && /reminder-rules/.test(r.u));
      const srv = (await H.http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config;
      t('DU-2 儲存回 202：顯示「已送出待另一位管理員覆核」（含單號）；伺服器目前生效的規則沒變（仍是預設、D-90）；畫面改回伺服器實際生效的內容', /已送出待另一位管理員覆核/.test(st) && /單號/.test(st) && srv.isDefault === true && srv.rules[0].whenDays === 90 && (await b.ev(`MaRemAdmin._state.models[0].nRaw`)) === '90' && !!put && JSON.parse(put.b).rules.some((r) => r.id === 'default-d90'), st + ' ' + J(srv.rules[0]));
      await b.sleep(900);
      const banner = await b.ev(`document.getElementById('marCrHost').innerText`);
      t('DU-3 CfgReview 橫幅顯示「待覆核」筆數', /待覆核/.test(banner) && /1/.test(banner), banner.slice(0, 120));
      const list = (await H.http(U.adm2, 'GET', '/api/admin/approval-config-requests')).json;
      const reqs0 = Array.isArray(list) ? list : (list.requests || []);
      const mine = reqs0.find((x) => x.kind === 'reminder_rules' && x.status === 'pending');
      t('DU-4 覆核單類型為 reminder_rules（到期提醒規則），狀態 pending', !!mine, J(reqs0.slice(0, 2)).slice(0, 200));
      const self = await H.http(U.adm, 'POST', '/api/admin/approval-config-requests/' + (mine && mine.id) + '/approve', {});
      t('DU-5 自己不能覆核自己的申請（403）', self.status === 403, self.status + ' ' + self.text.slice(0, 100));
      const ap = await H.http(U.adm2, 'POST', '/api/admin/approval-config-requests/' + (mine && mine.id) + '/approve', {});
      const srv2 = (await H.http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config;
      t('DU-6 另一位管理員核可後：規則生效（D-90 改為 75 天、規則 id 沿用 default-d90）', ap.status === 200 && srv2.rules[0].whenDays === 75 && srv2.rules[0].id === 'default-d90' && srv2.isDefault === false, ap.status + ' ' + J(srv2.rules[0]));
      await b.nav(H.BASE + '/admin.html');
      await b.waitFor(`!!document.getElementById('marNavItem')`, 15000); await b.ev(PAGE_HELPERS);
      await gotoRules(b);
      t('DU-7 重新開啟後台頁：顯示已生效的 75 天', (await b.ev(`MaRemAdmin._state.models[0].nRaw`)) === '75');
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('DU-E 沒有 JS 例外', fe.length === 0, J(fe).slice(0, 300));
    });
  }

  // ═════════════════════════ bell：鈴鐺 ═════════════════════════
  async function secBell(C) {
    const { H } = C;
    const notifs = [
      { id: 'nNR', to: U.own, type: 'ma_not_renew', title: '🚫 R丁公司 MA 不續約', body: '原因：預算不足', refId: '/index.html#maitems', read: false, createdAt: agoIso(1) },
      { id: 'nMR', to: U.own, type: 'ma_reminder', title: '📌 你有 1 項 MA／License 需要處理', body: 'R甲公司', refId: tp(0), read: false, createdAt: agoIso(0.5), maItemIds: ['iA'] },
    ];
    await H.fresh(mkUsers(H), seedData({ notifs }), { env: { MA_TRACKING: '1' } });
    await H.loginAll([U.own, U.sec, U.m1]);
    const api = await pollIds(H, U.own);
    t('BL-0 前提：業務甲的 poll-bundle maReminders ＝ 逾期 iC、iB(20)、iA(40)、iX(60)（已送簽的 iG、不續約、已續約、未達標的 iF、他人的 iO 都沒有）', J(api) === J(['iC', 'iB', 'iA', 'iX']), J(api));
    const pbFull = (await H.http(U.own, 'GET', '/api/poll-bundle')).json;
    t('BL-0b 前提：poll-bundle 的 notifications 仍含儲存型 ma_reminder（歷史），前端不應重複顯示', pbFull.notifications.some((n) => n.type === 'ma_reminder' && n.id === 'nMR') && pbFull.notifications.some((n) => n.type === 'ma_not_renew'));
    await browse(C, 700000, async (b) => {
      await openUser(b, C, U.own, 1440, 900, { expectOn: true, clearLs: true });
      await b.sleep(1500);
      // 摘要與徽章
      const sum = await b.ev(`(()=>{const s=document.querySelector('#notifList .mar-bell-sum'); return JSON.stringify({text:s?s.querySelector('.notif-title').textContent:null, unread:s?s.classList.contains('unread'):null, exp:s?s.getAttribute('aria-expanded'):null, hd:[...document.querySelectorAll('#notifList .notif-group-hd')].map(x=>x.textContent), badge:document.getElementById('notifBadge').textContent, badgeShown:document.getElementById('notifBadge').style.display!=='none'})})()`).then(JSON.parse);
      t('BL-1 鈴鐺：「你有 4 項 MA／License 需要處理」摘要（未讀、預設收合），群組標題「MA／License 到期提醒」；未讀徽章把摘要算 1 件（另有儲存型 ma_not_renew 1 件，ma_reminder 不重複顯示）', sum.text === '你有 4 項 MA／License 需要處理' && sum.unread === true && sum.exp === 'false' && sum.hd.some((h) => /MA／License 到期提醒/.test(h)) && sum.badgeShown && sum.badge === '2', J(sum));
      const dupShown = await b.ev(`[...document.querySelectorAll('#notifList .notif-item')].map(x=>x.getAttribute('data-id')||'').join(',')`);
      t('BL-2 儲存型 ma_reminder（nMR）不在鈴鐺清單；ma_not_renew（nNR）顯示成一般通知、圖示 🚫', dupShown.indexOf('nMR') < 0 && dupShown.indexOf('nNR') >= 0 && (await b.ev(`document.querySelector('#notifList .notif-item[data-id="nNR"] .notif-icon').textContent`)) === '🚫', dupShown);
      const hidden0 = await b.ev(`document.querySelector('#notifList .mar-bell-list').hidden`);
      t('BL-3 預設收合：清單列隱藏', hidden0 === true);
      await openBell(b);
      await b.click('#notifList .mar-bell-sum'); await b.sleep(500);
      const open1 = await b.ev(`(()=>{const rows=[...document.querySelectorAll('#notifList .mar-bell-row')]; return JSON.stringify({exp:document.querySelector('#notifList .mar-bell-sum').getAttribute('aria-expanded'), vis:!document.querySelector('#notifList .mar-bell-list').hidden, n:rows.length, ids:rows.map(r=>r.getAttribute('data-marem-row')), first:rows[0].innerText.replace(/\\s+/g,' '), sumUnread:document.querySelector('#notifList .mar-bell-sum').classList.contains('unread'), badge:document.getElementById('notifBadge').textContent, ls:Object.keys(localStorage).filter(k=>k.indexOf('marem_read_')===0)})})()`).then(JSON.parse);
      t('BL-4 點摘要展開：4 列依剩餘天數排序（iC 逾期 20 天最前）；每列有公司、品項、天數、所屬規則名稱；摘要變已讀、徽章剩 1；已讀存 localStorage（marem_read_ma_<id>_<日期>）', open1.exp === 'true' && open1.vis && open1.n === 4 && J(open1.ids) === J(['ma_iC', 'ma_iB', 'ma_iA', 'ma_iX']) && /R丙公司/.test(open1.first) && /已逾期 20 天/.test(open1.first) && /規則：到期後每 7 天追蹤未續約/.test(open1.first) && open1.sumUnread === false && open1.badge === '1' && open1.ls.length === 4 && open1.ls.every((k) => new RegExp('^marem_read_ma_i[A-Z]+_' + tp(0) + '$').test(k)), J(open1));
      const rowsTxt = await b.ev(`JSON.stringify([...document.querySelectorAll('#notifList .mar-bell-row')].map(r=>r.innerText.replace(/\\s+/g,' ')))`).then(JSON.parse);
      t('BL-5 各列文字：iB「R乙公司 Svc Lic 永久 License 剩 20 天到期 規則：到期前 90 天…」、iA 剩 40 天、iX XSS 公司名原樣顯示成文字', /R乙公司/.test(rowsTxt[1]) && /License/.test(rowsTxt[1]) && /剩 20 天到期/.test(rowsTxt[1]) && /剩 40 天到期/.test(rowsTxt[2]) && rowsTxt[3].indexOf(XSS1) >= 0 && rowsTxt[3].indexOf(XSS2) >= 0, J(rowsTxt));
      const noAmt = await b.ev(`(()=>{const h=document.getElementById('notifList').innerHTML; return JSON.stringify({img:document.querySelectorAll('#notifList img').length, script:document.querySelectorAll('#notifList script').length, xss:window.__xss, cno:h.indexOf(${J(CNO)})>=0, amt:h.indexOf(${J(AMT)})>=0})})()`).then(JSON.parse);
      t('BL-6 鈴鐺內沒有 <img>／<script> 元素（純文字）、XSS 沒有執行、沒有合約編號與金額', noAmt.img === 0 && noAmt.script === 0 && noAmt.xss === 0 && !noAmt.cno && !noAmt.amt, J(noAmt));
      await shotTo(C, b, 'bell_open_glass_dark.png');
      // 重新輪詢後仍已讀（今天）、展開狀態保留
      await refreshBell(b);
      const keep = await b.ev(`JSON.stringify({exp:document.querySelector('#notifList .mar-bell-sum').getAttribute('aria-expanded'), unread:document.querySelector('#notifList .mar-bell-sum').classList.contains('unread')})`).then(JSON.parse);
      t('BL-7 下一次輪詢重畫後：展開狀態保留、已讀狀態保留（沿用合約提醒的 localStorage 每日重置做法）', keep.exp === 'true' && keep.unread === false, J(keep));
      // 鍵盤：摘要 Enter 收合
      await b.focus('#notifList .mar-bell-sum'); await b.key('Enter'); await b.sleep(300);
      t('BL-8 鍵盤：摘要列 focus＋Enter 收合（aria-expanded=false）', (await b.ev(`document.querySelector('#notifList .mar-bell-sum').getAttribute('aria-expanded')`)) === 'false');
      await b.focus('#notifList .mar-bell-sum'); await b.key('Enter'); await b.sleep(300);
      // 深連結：點 iA 那列 → 開到期管理頁並標示該列
      await b.click('#notifList .mar-bell-row[data-marem-row="ma_iA"]'); await b.sleep(2000);
      const nav = await b.ev(`JSON.stringify({view:!document.getElementById('maItemsView').hidden, title:document.getElementById('topbarTitle').textContent, focus:!!document.querySelector('#maItemsView tr.mat-tr-focus[data-item-id="iA"]'), dd:document.getElementById('notifDropdown').style.display})`).then(JSON.parse);
      t('BL-9 點某一列：關閉鈴鐺、開「MA／License 到期管理」頁，並標示（外框）該納管項那一列', nav.view && nav.title === 'MA／License 到期管理' && nav.focus && nav.dd === 'none', J(nav));
      await b.ev(`showSection('home'); 0`).catch(() => {});
      // 儲存型 ma_not_renew 點擊 → 開到期管理頁
      await b.ev(`showSection('pipeline'); 0`); await b.sleep(500);
      await openBell(b);
      await b.click('#notifList .notif-item[data-id="nNR"]'); await b.sleep(1800);
      t('BL-10 點儲存型 ma_not_renew 通知：開到期管理頁（深連結 #maitems）', await b.ev(`!document.getElementById('maItemsView').hidden`));
      const putsRead = (await reqs(b)).filter((r) => r.m === 'PUT' && /notifications\/nNR\/read/.test(r.u));
      t('BL-11 儲存型通知照舊標已讀（PUT notifications/:id/read）', putsRead.length === 1, J(putsRead));
      // 全部已讀鈕也標 MA
      await b.ev(`localStorage.clear(); 0`); await refreshBell(b); await openBell(b);
      const beforeAll = await b.ev(`document.querySelector('#notifList .mar-bell-sum').classList.contains('unread')`);
      await b.click('#notifReadAll'); await b.sleep(1400);
      const afterAll = await b.ev(`document.querySelector('#notifList .mar-bell-sum').classList.contains('unread')`);
      t('BL-12 「全部標為已讀」：MA／License 群組也標已讀', beforeAll === true && afterAll === false, J([beforeAll, afterAll]));

      // ── 條件達成後下一次輪詢自動消失 ──
      await b.ev(`localStorage.clear(); 0`);
      writeData((d) => { d.quotations.find((q) => q.id === 'qDraft').approval = { state: 'pending', steps: [], cur: 0, history: [] }; });
      await refreshBell(b); await openBell(b);
      let ids = await shown(b);
      await b.ev(`document.querySelector('#notifList .mar-bell-sum').getAttribute('aria-expanded')==='true' || document.querySelector('#notifList .mar-bell-sum').click(); 0`); await b.sleep(400);
      ids = await shown(b);
      t('BL-13 續約報價單送簽（簽核中）：下一次輪詢 iA 的提醒自動消失（預設 D-90 無條件規則也停止），其餘 3 項還在，摘要變「3 項」', J(ids) === J(['ma_iC', 'ma_iB', 'ma_iX']) && (await b.ev(`document.querySelector('#notifList .mar-bell-sum .notif-title').textContent`)) === '你有 3 項 MA／License 需要處理', J(ids));
      writeData((d) => { d.quotations.find((q) => q.id === 'qDraft').invalid = { at: NOW, by: 'x', reason: '測試作廢' }; });
      await refreshBell(b);
      ids = await shown(b);
      t('BL-14 該續約報價單作廢：iA 的提醒又回來（視同沒有續約報價單）', ids.indexOf('ma_iA') >= 0 && ids.length === 4, J(ids));
      const r1 = await H.http(U.own, 'PATCH', '/api/ma-items/iB/not-renew', { reason: '測試用不續約' });
      await refreshBell(b);
      ids = await shown(b);
      t('BL-15 業務標記 iB 不續約（API）：下一次輪詢 iB 消失、摘要變 3 項', r1.status === 200 && ids.indexOf('ma_iB') < 0 && ids.length === 3, r1.status + ' ' + J(ids));
      await H.http(U.own, 'PATCH', '/api/ma-items/iC/not-renew', { reason: '測試用不續約' });
      await H.http(U.own, 'PATCH', '/api/ma-items/iX/not-renew', { reason: '測試用不續約' });
      await H.http(U.own, 'PATCH', '/api/ma-items/iA/not-renew', { reason: '測試用不續約' });
      await refreshBell(b);
      const gone = await b.ev(`JSON.stringify({sum:document.querySelectorAll('#notifList .mar-bell-sum').length, host:!!document.getElementById('maRemBellHost'), hd:[...document.querySelectorAll('#notifList .notif-group-hd')].map(x=>x.textContent), badge:document.getElementById('notifBadge').style.display})`).then(JSON.parse);
      t('BL-16 全部條件達成（都標不續約）：MA／License 群組整個消失，不留空標題', gone.sum === 0 && gone.host === false && !gone.hd.some((h) => /MA／License/.test(h)), J(gone));
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('BL-E1 沒有 JS 例外（業務）', fe.length === 0, J(fe).slice(0, 300));
    });
    // ── 秘書視角 ──
    await H.fresh(mkUsers(H), seedData({ notifs }), { env: { MA_TRACKING: '1' } });
    await H.loginAll([U.own, U.sec, U.m1, U.adm]);
    const secApi = (await H.http(U.sec, 'GET', '/api/poll-bundle')).json;
    const secRows = secApi.maReminders;
    t('BL-S0 前提：秘書的 maReminders 每列只有 11 個白名單欄位（無 itemId／規則／金額／合約編號）', secRows.length >= 3 && secRows.every((x) => Object.keys(x).sort().join() === 'body,company,days,from,id,level,ownerLabel,product,termKind,title,to,type') && !J(secRows).includes(CNO) && !J(secRows).includes(AMT), J(secRows.map((x) => x.company)));
    await browse(C, 400000, async (b) => {
      await openUser(b, C, U.sec, 1440, 900, { expectOn: true, clearLs: true });
      await b.sleep(1200);
      await openBell(b);
      await b.click('#notifList .mar-bell-sum'); await b.sleep(500);
      const sv = await b.ev(`(()=>{const rows=[...document.querySelectorAll('#notifList .mar-bell-row')]; const h=document.getElementById('notifList').innerHTML; return JSON.stringify({n:rows.length, rules:rows.map(r=>!!r.querySelector('.mar-bell-rule')&&/規則：/.test(r.querySelector('.mar-bell-rule').textContent)), owner:rows.map(r=>/業務：/.test(r.innerText)), roles:rows.filter(r=>r.querySelector('.mar-bell-roles')).length, cno:h.indexOf(${J(CNO)})>=0, amt:h.indexOf(${J(AMT)})>=0, ids:rows.map(r=>r.getAttribute('data-marem-row')), text:rows.map(r=>r.innerText.replace(/\\s+/g,' '))})})()`).then(JSON.parse);
      t('BL-S1 秘書的鈴鐺：同樣有群組與展開；每列只有公司、品項、天數與「業務：姓名」，沒有規則名稱與角色說明、沒有合約編號與金額；列代號是雜湊（ma_＋12 碼），不是納管項 id', sv.n === secRows.length && sv.rules.every((x) => x === false) && sv.owner.every((x) => x) && sv.roles === 0 && !sv.cno && !sv.amt && sv.ids.every((i) => /^ma_[0-9a-f]{12}$/.test(i)), J(sv));
      await b.click('#notifList .mar-bell-row'); await b.sleep(2000);
      const nv = await b.ev(`JSON.stringify({view:!document.getElementById('maItemsView').hidden, cols:[...document.querySelectorAll('#maItemsView thead th')].map(x=>x.textContent), page:document.getElementById('maItemsView').innerText.indexOf(${J(CNO)})>=0, hl:document.querySelectorAll('#maItemsView tr.mat-tr-focus').length})`).then(JSON.parse);
      t('BL-S2 秘書點列：開到期管理頁（5 欄唯讀，沒有合約編號欄）', nv.view && J(nv.cols) === J(['客戶', '品項', '期間', '剩餘天數', '負責業務']) && !nv.page, J(nv));
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('BL-E2 沒有 JS 例外（秘書）', fe.length === 0, J(fe).slice(0, 300));
    });
    // ── 一級主管 ──
    const m1Api = await pollIds(H, U.m1);
    t('BL-M0 一級主管的鈴鐺依規則：iB（剩 20，D-45 無報價）、iC（逾期，到期後規則）；草稿 iA 因 D-45 條件不成立而沒有', J(m1Api) === J(['iC', 'iB']), J(m1Api));
    await browse(C, 300000, async (b) => {
      await openUser(b, C, U.m1, 1440, 900, { expectOn: true, clearLs: true });
      await openBell(b);
      const mv = await b.ev(`JSON.stringify([...document.querySelectorAll('#notifList .mar-bell-row')].map(r=>r.getAttribute('data-marem-row')+':'+(r.querySelector('.mar-bell-roles')?r.querySelector('.mar-bell-roles').textContent:'')))`);
      t('BL-M1 主管的鈴鐺列有「你是：一級主管」說明', /ma_iC:你是：一級主管/.test(mv) && /ma_iB:你是：一級主管/.test(mv), mv);
    });
  }

  // ═════════════════════════ items：到期管理頁 ═════════════════════════
  async function secItems(C) {
    const { H } = C;
    await H.fresh(mkUsers(H), seedData(), { env: { MA_TRACKING: '1' } });
    await H.loginAll([U.own, U.adm, U.sec, U.m1, U.mk]);
    await browse(C, 900000, async (b) => {
      await openUser(b, C, U.own, 1440, 1000, { expectOn: true });
      await toItems(b);
      const hdr = await b.ev(`[...document.querySelectorAll('#maItemsView thead th')].map(x=>x.textContent).join('|')`);
      t('IM-1 欄位不變（客戶|品項|合約編號|期間|剩餘天數|狀態／下一步|負責業務|操作）', hdr === '客戶|品項|合約編號|期間|剩餘天數|狀態／下一步|負責業務|操作', hdr);
      const rows = await b.ev(`JSON.stringify([...document.querySelectorAll('#maItemsView tbody tr')].map(r=>[r.getAttribute('data-item-id'), r.querySelector('.mar-badge')?r.querySelector('.mar-badge').textContent:'']))`).then(JSON.parse);
      t('IM-2 全部 8 筆（自己的、不含作廢）：狀態徽章 追蹤中×6（iA、iB、iC、iF、iG、iX）／不續約×1（iD）／已續約×1（iE）', rows.length === 8 && rows.filter((r) => r[1] === '追蹤中').length === 6 && rows.filter((r) => r[1] === '不續約').length === 1 && rows.filter((r) => r[1] === '已續約').length === 1, J(rows));
      const dRow = await rowText(b, 'iD');
      t('IM-3 不續約的列：顯示徽章「不續約」與原因「預算不足：明年再議」；業務本人沒有「重新開啟」（只有管理員）', /不續約/.test(dRow) && /原因：預算不足：明年再議/.test(dRow) && !/重新開啟/.test(dRow) && /標記不續約…/.test(await rowText(b, 'iA')) && !/標記不續約…/.test(dRow), dRow);
      const eRow = await rowText(b, 'iE');
      t('IM-4 已續約的列：徽章「已續約」，沒有「標記不續約」、沒有「下一個提醒」', /已續約/.test(eRow) && !/標記不續約…/.test(eRow) && !/下一個提醒/.test(eRow), eRow);
      // 下一個提醒（與伺服器時間軸一致）
      const tlA = (await H.http(U.own, 'GET', '/api/ma-items/iA/reminders')).json;
      const act = tlA.rules.filter((r) => r.activeNow).map((r) => r.name);
      const fut = tlA.rules.filter((r) => r.enabled && r.appliesToThis && !r.reached && r.triggerDate >= tp(0)).sort((a, z) => (a.triggerDate < z.triggerDate ? -1 : 1))[0];
      const wt = (r) => (r.whenDays > 0 ? '到期前 ' + r.whenDays + ' 天' : r.whenDays === 0 ? '到期當天' : '到期後 ' + (-r.whenDays) + ' 天') + (r.repeatEveryDays ? '，每 ' + r.repeatEveryDays + ' 天重複' : '');
      const expNext = '目前提醒中：' + act.join('、') + '；下一個提醒：' + slash(fut.triggerDate) + '（' + fut.name + '，' + wt(fut) + '）';
      await b.waitFor(`/下一個提醒：|目前提醒中|沒有後續提醒|提醒已停止/.test(document.querySelector('#maItemsView tr[data-item-id="iA"] .mar-next').textContent)`, 8000);
      const nextA = await txt(b, '#maItemsView tr[data-item-id="iA"] .mar-next');
      t('IM-5 「下一個提醒」（iA 剩 40 天、草稿）：目前提醒中的規則＋下一個觸發的規則與日期，與伺服器時間軸一致', nextA === expNext && act.length === 1, nextA + ' ≠ ' + expNext);
      const tlF = (await H.http(U.own, 'GET', '/api/ma-items/iF/reminders')).json;
      const futF = tlF.rules.filter((r) => r.enabled && r.appliesToThis && !r.reached && r.triggerDate >= tp(0)).sort((x, y) => (x.triggerDate < y.triggerDate ? -1 : 1))[0];
      await b.ev(`document.querySelector('#maItemsView tr[data-item-id="iF"]').scrollIntoView({block:'center'}); 0`);
      await b.waitFor(`!/載入中/.test(document.querySelector('#maItemsView tr[data-item-id="iF"] .mar-next').textContent)`, 8000);
      const nextF = await txt(b, '#maItemsView tr[data-item-id="iF"] .mar-next');
      t('IM-5b 剩 100 天的 iF（4 條規則都還沒到）：下一個提醒是「最近」的一條（到期前 90 天，日期＝迄日−90）', nextF === '下一個提醒：' + slash(futF.triggerDate) + '（' + futF.name + '，' + wt(futF) + '）' && /到期前 90 天/.test(nextF), nextF);
      const nextC = await txt(b, '#maItemsView tr[data-item-id="iC"] .mar-next');
      const nextD = await rowText(b, 'iD');
      t('IM-6 逾期的 iC：目前提醒中（到期後規則）、沒有後續提醒或下一個；不續約的 iD 沒有「下一個提醒」區塊', /目前提醒中：到期後每 7 天/.test(nextC) && !/下一個提醒/.test(nextD), nextC);
      // 提醒時間軸展開
      await b.click(btnIn('iA', '提醒時間軸'));
      await b.waitFor(`document.querySelectorAll('#maItemsView tr.mar-tl-tr .mar-tl-li').length===4`, 8000);
      const tl = await b.ev(`JSON.stringify({exp:document.querySelector('#maItemsView tr[data-item-id="iA"] .mar-btn-tl').getAttribute('aria-expanded'), prev:document.querySelector('#maItemsView tr.mar-tl-tr').previousElementSibling.getAttribute('data-item-id'), dates:[...document.querySelectorAll('#maItemsView tr.mar-tl-tr .mar-tl-date')].map(x=>x.textContent), states:[...document.querySelectorAll('#maItemsView tr.mar-tl-tr .mar-tl-state')].map(x=>x.textContent), sub:document.querySelector('#maItemsView .mar-tl-sub').textContent, rows:document.querySelectorAll('#maItemsView tbody tr').length})`).then(JSON.parse);
      const expDates = tlA.rules.map((r) => slash(r.triggerDate)).sort();
      t('IM-7 展開提醒時間軸：緊接在該列後出現、4 條規則依觸發日排序（日期與伺服器一致）、狀態文字「目前提醒中／尚未到達」、說明規則來源與續約報價單狀態（草稿）', tl.exp === 'true' && tl.prev === 'iA' && J(tl.dates) === J(expDates) && tl.states.filter((s) => s === '目前提醒中').length === 1 && tl.states.filter((s) => s === '尚未到達').length >= 1 && /系統預設規則/.test(tl.sub) && /草稿/.test(tl.sub), J(tl));
      await shotTo(C, b, 'items_timeline.png');
      await b.click(btnIn('iA', '提醒時間軸')); await b.sleep(300);
      t('IM-8 再按一次收合：時間軸列移除、aria-expanded=false', (await b.ev(`document.querySelectorAll('#maItemsView tr.mar-tl-tr').length`)) === 0 && (await b.ev(`document.querySelector('#maItemsView tr[data-item-id="iA"] .mar-btn-tl').getAttribute('aria-expanded')`)) === 'false');
      // 篩選晶片
      await b.click(() => [...document.querySelectorAll('#maItemsView .mat-fchip')].find((c) => c.textContent === '已不續約'));
      await b.sleep(500);
      const fl = await b.ev(`JSON.stringify({rows:[...document.querySelectorAll('#maItemsView tbody tr')].map(r=>r.getAttribute('data-item-id')), pressed:[...document.querySelectorAll('#maItemsView .mat-fchip')].find(c=>c.textContent==='已不續約').getAttribute('aria-pressed'), chips:[...document.querySelectorAll('#maItemsView .mat-fchip')].map(c=>c.textContent)})`).then(JSON.parse);
      t('IM-9 篩選晶片新增「已不續約」：只列不續約的項目（iD）、按下的晶片 aria-pressed=true；其餘晶片照舊', J(fl.rows) === J(['iD']) && fl.pressed === 'true' && J(fl.chips) === J(['全部', '待簽約', '90 天內到期', '已逾期', '已不續約']), J(fl));
      await b.click(() => [...document.querySelectorAll('#maItemsView .mat-fchip')].find((c) => c.textContent === '全部')); await b.sleep(700);

      // ── 標記不續約對話框 ──
      await b.click(btnIn('iB', '標記不續約…'));
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      const dlg = await b.ev(`(()=>{const d=document.querySelector('.mat-overlay .mat-dialog'); return JSON.stringify({role:d.getAttribute('role'), modal:d.getAttribute('aria-modal'), title:d.querySelector('h2').textContent, labelled:!!d.getAttribute('aria-labelledby'), opts:[...d.querySelectorAll('select option')].map(o=>o.textContent), focus:document.activeElement.tagName})})()`).then(JSON.parse);
      t('IM-10 「標記不續約…」：對話框（role=dialog、aria-modal、標題含公司）、原因下拉（請選擇／客戶停用／轉用他廠／預算不足／併入新合約／其他）、焦點在下拉', dlg.role === 'dialog' && dlg.modal === 'true' && /R乙公司/.test(dlg.title) && dlg.labelled && J(dlg.opts) === J(['請選擇原因…', '客戶停用', '轉用他廠', '預算不足', '併入新合約', '其他']) && dlg.focus === 'SELECT', J(dlg));
      await b.key('Escape'); await b.sleep(300);
      t('IM-11 Esc 關閉對話框，焦點回到原按鈕', (await b.ev(`document.querySelectorAll('.mat-overlay').length`)) === 0 && (await b.ev(`document.activeElement && document.activeElement.textContent.trim()`)) === '標記不續約…');
      await b.click(btnIn('iB', '標記不續約…')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      await b.click(topBtn('確認不續約')); await b.sleep(200);
      t('IM-12 沒選原因就按確認：顯示「請選擇不續約的原因」、不送請求', /請選擇不續約的原因/.test(await dlgText(b)) && (await reqs(b)).filter((r) => /not-renew/.test(r.u)).length === 0);
      await pickOpt(b, '.mat-overlay select', '其他');
      await b.focus('.mat-overlay textarea'); await b.type('好'); await b.click(topBtn('確認不續約')); await b.sleep(200);
      t('IM-13 選「其他」但備註只有 1 個字：顯示「備註至少要 2 個字」、不送請求', /備註至少要 2 個字/.test(await dlgText(b)) && (await reqs(b)).filter((r) => /not-renew/.test(r.u)).length === 0);
      await b.type('客戶改用他牌');
      await b.click(topBtn('確認不續約'));
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000); await b.sleep(1500);
      const nr = (await reqs(b)).filter((r) => /not-renew/.test(r.u));
      const iB = (await H.http(U.own, 'GET', '/api/ma-items?filter=all')).json.items.find((x) => x.id === 'iB');
      t('IM-14 確認：PATCH not-renew 帶 reason「其他：好客戶改用他牌」（原因＋備註）→ 伺服器 iB＝不續約、原因同；列顯示徽章與原因、「下一個提醒」消失', nr.length === 1 && J(JSON.parse(nr[0].b)) === J({ reason: '其他：好客戶改用他牌' }) && iB.status === 'not_renew' && iB.notRenewReason === '其他：好客戶改用他牌' && /不續約/.test(await rowText(b, 'iB')) && /原因：其他：好客戶改用他牌/.test(await rowText(b, 'iB')), J(nr));
      const toastTxt = await b.ev(`(document.getElementById('toast')||{}).textContent||''`);
      const m1n = (await H.http(U.m1, 'GET', '/api/notifications')).json, mkn = (await H.http(U.mk, 'GET', '/api/notifications')).json;
      t('IM-15 通知：一級主管與行銷各收到 1 則 ma_not_renew（內文有原因、沒有金額）；提示訊息含「已通知 N 位同仁」', [m1n, mkn].every((l) => l.filter((n) => n.type === 'ma_not_renew' && /客戶改用他牌/.test(n.body)).length === 1) && !J([m1n, mkn]).includes(AMT) && /已標記不續約/.test(toastTxt) && /已通知 2 位同仁/.test(toastTxt), toastTxt);
      // 只選類別（無備註）
      await b.click(btnIn('iF', '標記不續約…')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      await pickOpt(b, '.mat-overlay select', '預算不足');
      await b.click(topBtn('確認不續約')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000); await b.sleep(1000);
      const iF = (await H.http(U.own, 'GET', '/api/ma-items?filter=all')).json.items.find((x) => x.id === 'iF');
      t('IM-16 只選類別、備註留空：reason＝「預算不足」，可以送出', iF.status === 'not_renew' && iF.notRenewReason === '預算不足', J(iF.notRenewReason));
      // XSS：原因含 HTML
      await b.click(btnIn('iG', '標記不續約…')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      await pickOpt(b, '.mat-overlay select', '其他');
      await b.focus('.mat-overlay textarea'); await b.type(XSS1.slice(0, 40));
      await b.click(topBtn('確認不續約')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000); await b.sleep(1200);
      const xr = await b.ev(`JSON.stringify({img:document.querySelectorAll('#maItemsView img').length, xss:window.__xss, txt:document.querySelector('#maItemsView tr[data-item-id="iG"] .mar-reason').textContent})`).then(JSON.parse);
      t('IM-17 不續約原因含 <img onerror>：以純文字顯示、沒有 <img> 元素、沒有執行', xr.img === 0 && xr.xss === 0 && xr.txt.indexOf('<img') >= 0, J(xr));
      // 公司名 XSS 列
      const xssRow = await b.ev(`JSON.stringify({img:document.querySelectorAll('#maItemsView tbody img').length, script:document.querySelectorAll('#maItemsView tbody script').length, xss:window.__xss})`).then(JSON.parse);
      t('IM-18 到期管理頁整頁（含 XSS 公司／品項／合約編號）沒有 <img>／<script> 元素、沒有執行', xssRow.img === 0 && xssRow.script === 0 && xssRow.xss === 0, J(xssRow));
      await shotTo(C, b, 'items_page_after_norenew.png');
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('IM-E1 沒有 JS 例外（業務）', fe.length === 0, J(fe).slice(0, 300));
    });
    // ── 管理員：重新開啟 ──
    await browse(C, 400000, async (b) => {
      await openUser(b, C, U.adm, 1440, 1000, { expectOn: true });
      await toItems(b);
      const dBtn = await b.ev(`(${btnIn('iD', '重新開啟')})()!==null`);
      t('IM-20 管理員看得到不續約列的「重新開啟」；追蹤中的列沒有', dBtn === true && (await b.ev(`(${btnIn('iA', '重新開啟')})()===null`)));
      await b.click(btnIn('iD', '重新開啟'));
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      t('IM-21 重新開啟對話框說明並顯示原因', /預算不足：明年再議/.test(await dlgText(b)));
      await b.click(topBtn('重新開啟')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000); await b.sleep(1300);
      const iD = (await H.http(U.adm, 'GET', '/api/ma-items?filter=all')).json.items.find((x) => x.id === 'iD');
      const tx = await rowText(b, 'iD');
      t('IM-22 重新開啟：PATCH reopen；伺服器 iD 回到追蹤中；列顯示「追蹤中」並重新有「下一個提醒」與「標記不續約…」', iD.status === 'tracking' && /追蹤中/.test(tx) && /標記不續約…/.test(tx) && !/重新開啟/.test(tx), tx);
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('IM-E2 沒有 JS 例外（管理員）', fe.length === 0, J(fe).slice(0, 300));
    });
    // ── 秘書 ──
    await browse(C, 300000, async (b) => {
      await openUser(b, C, U.sec, 1440, 900, { expectOn: true });
      await toItems(b, { fast: true });
      await b.sleep(1500);
      const sv = await b.ev(`JSON.stringify({cols:[...document.querySelectorAll('#maItemsView thead th')].map(x=>x.textContent), next:document.querySelectorAll('#maItemsView .mar-next, #maItemsView .mar-badge, #maItemsView .mar-btn-tl, #maItemsView .mar-tl-tr').length, btn:document.querySelectorAll('#maItemsView tbody button').length, chips:[...document.querySelectorAll('#maItemsView .mat-fchip')].map(c=>c.textContent), cno:document.getElementById('maItemsView').innerText.indexOf(${J(CNO)})>=0, amt:document.getElementById('maItemsView').innerText.indexOf(${J(AMT)})>=0})`).then(JSON.parse);
      t('IM-S1 秘書：5 欄唯讀、沒有徽章／下一個提醒／時間軸／任何按鈕、篩選晶片沒有「待簽約」「已不續約」、沒有合約編號與金額', J(sv.cols) === J(['客戶', '品項', '期間', '剩餘天數', '負責業務']) && sv.next === 0 && sv.btn === 0 && J(sv.chips) === J(['全部', '90 天內到期', '已逾期']) && !sv.cno && !sv.amt, J(sv));
    });
  }

  // ═════════════════════════ quote：續約選擇器 ═════════════════════════
  const rowSel = (n) => `#quoteItemsBody tr:nth-child(${n})`;
  const Qs = (s) => `document.querySelector(${J(s)})`;
  async function toQuotes(b) {
    await b.ev(`window.__xss=0; showSection('quotations'); 0`);
    await b.waitFor(`typeof quoteCfg==='function' && !!quoteCfg().me && document.querySelectorAll('#quoteTbody tr').length>0`, 15000);
    await b.sleep(500);
  }
  async function openEdit(b, id) {
    await b.ev(`openQuoteModal(${J(id)}); 0`);
    await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='flex' && document.querySelectorAll('#quoteItemsBody tr').length>0`, 15000);
    await b.sleep(1100);
  }
  const hostInfo = (b) => b.ev(`(()=>{const h=document.getElementById('qRenewHost'); if(!h) return JSON.stringify({exists:false}); const s=h.querySelector('select'); return JSON.stringify({exists:true, hidden:h.hidden, cls:h.className, val:s?s.value:null, opts:s?[...s.options].map(o=>[o.value,o.textContent]):[], text:h.innerText.replace(/\\s+/g,' ')})})()`).then(JSON.parse);
  async function secQuote(C) {
    const { H } = C;
    await H.fresh(mkUsers(H), seedData(), { env: { MA_TRACKING: '1' } });
    await H.loginAll([U.own, U.adm]);
    await browse(C, 900000, async (b) => {
      await openUser(b, C, U.own, 1440, 1000, { expectOn: true });
      await toQuotes(b);
      await b.ev(`openQuoteModal(null,{oppId:'oR1'}); 0`);
      await b.waitFor(`document.querySelectorAll('#quoteItemsBody tr').length>0`, 15000);
      await b.sleep(1500);
      let hi = await hostInfo(b);
      t('QR-1 新增報價單（商機＝R甲公司）：出現「這是續約嗎？」橫幅與下拉（預設「不是續約」＋續約候選 iA：品項、到期日、剩 40 天、已有續約報價單）', hi.exists && !hi.hidden && hi.val === '' && hi.opts.length === 2 && hi.opts[1][0] === 'iA' && /續約：Svc MA 一年（MA 到期 \d{4}\/\d{2}\/\d{2}，剩 40 天，已有續約報價單）/.test(hi.opts[1][1]) && /這是續約嗎？/.test(hi.text) && /不會印在客戶版報價單/.test(hi.text), J(hi));
      t('QR-2 橫幅放在「連結商機」那一步之後、不在任何 .qs-step 裡（逐步填寫時步驟收合也看得到）', await b.ev(`(()=>{const h=document.getElementById('qRenewHost'); return document.getElementById('qOppStep').nextElementSibling===h && !h.closest('.qs-step')})()`));
      await shotTo(C, b, 'quote_renew_picker.png');
      // 選擇（真實鍵盤）
      await pickOpt(b, '#qRenewHost select', 'iA');
      // 填單：品名、單價 → 走完逐步填寫 → 儲存
      await b.focus(rowSel(1) + ' .qi-desc'); await b.ev(`${Qs(rowSel(1) + ' .qi-desc')}.select(); 0`); await b.type('客製開發服務'); await b.key('Tab');
      await b.focus(rowSel(1) + ' .qi-price'); await b.ev(`${Qs(rowSel(1) + ' .qi-price')}.select(); 0`); await b.type('50000'); await b.key('Tab');
      for (let i = 0; i < 8; i++) {
        const clicked = await b.ev(`(()=>{const bt=[...document.querySelectorAll('#quoteModalOverlay button')].find(x=>x.textContent.trim().indexOf('下一步')===0 && x.offsetParent!==null && !x.disabled); if(!bt) return false; bt.click(); return true})()`);
        if (!clicked) break;
        await b.sleep(350);
      }
      const saveEnabled = await b.waitFor(`!document.getElementById('quoteModalSave').disabled`, 6000);
      if (!saveEnabled) console.log('DBG', await b.ev(`document.getElementById('quoteModalOverlay').innerText.slice(0,800)`));
      const before = (await H.http(U.own, 'GET', '/api/quotations')).json.length;
      const rb = (await reqs(b)).length;
      await b.click('#quoteModalSave');
      await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='none'`, 15000);
      await b.sleep(1000);
      const post = (await reqs(b)).slice(rb).find((r) => r.m === 'POST' && /quotations$/.test(r.u));
      const list = (await H.http(U.own, 'GET', '/api/quotations')).json;
      const mine = list.find((q) => q.company === 'R甲公司' && q.id !== 'qDraft');
      t('QR-3 儲存：POST 帶 renewOf＝iA；伺服器新單 renewOf＝iA（新增一張）', saveEnabled && !!post && JSON.parse(post.b).renewOf === 'iA' && !!mine && mine.renewOf === 'iA' && list.length === before + 1, J([post && JSON.parse(post.b).renewOf, mine && mine.renewOf]));
      // 重新載入頁面 → 編輯：選項保留
      await openUser(b, C, U.own, 1440, 1000, { expectOn: true });
      await toQuotes(b);
      await openEdit(b, mine.id);
      hi = await hostInfo(b);
      t('QR-4 重新載入頁面、開啟編輯：續約對象保留（下拉顯示 iA），沒動過所以不會送 renewOf', hi.exists && hi.val === 'iA' && (await b.ev(`_qRenew.picker.touched()`)) === false, J(hi));
      // 改成「不是續約」→ 儲存 → 清除
      await pickOpt(b, '#qRenewHost select', '');
      const rb2 = (await reqs(b)).length;
      await b.click('#quoteModalSave');
      await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='none'`, 15000);
      await b.sleep(900);
      const put2 = (await reqs(b)).slice(rb2).find((r) => r.m === 'PUT' && /quotations\//.test(r.u));
      const q2 = (await H.http(U.own, 'GET', '/api/quotations/' + mine.id)).json;
      t('QR-5 選「不是續約」再儲存：PUT 帶 renewOf＝空字串；伺服器清除', !!put2 && JSON.parse(put2.b).renewOf === '' && !q2.renewOf, J([put2 && JSON.parse(put2.b).renewOf, q2.renewOf]));
      // 草稿 qDraft（renewOf＝iA）：不動它，直接儲存其他欄位 → 不送 renewOf，伺服器保留
      await openEdit(b, 'qDraft');
      hi = await hostInfo(b);
      const rb3 = (await reqs(b)).length;
      await b.focus('#qProjectName'); await b.ev(`document.getElementById('qProjectName').select(); 0`); await b.type('改過的專案名稱');
      await b.click('#quoteModalSave');
      await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='none'`, 15000); await b.sleep(900);
      const put3 = (await reqs(b)).slice(rb3).find((r) => r.m === 'PUT');
      const qd = (await H.http(U.own, 'GET', '/api/quotations/qDraft')).json;
      t('QR-6 編輯草稿單（原本有續約對象 iA）只改專案名稱：下拉顯示 iA、儲存不送 renewOf、伺服器保留 iA', hi.val === 'iA' && !!put3 && !('renewOf' in JSON.parse(put3.b)) && qd.renewOf === 'iA' && qd.projectName === '改過的專案名稱', J([hi.val, put3 && Object.keys(JSON.parse(put3.b)).includes('renewOf'), qd.renewOf]));
      // 沒有候選的公司：橫幅隱藏
      await openEdit(b, 'qPlain');
      hi = await hostInfo(b);
      t('QR-7 沒有續約候選的公司（R無候選公司）：不顯示橫幅（host 隱藏）', hi.exists && hi.hidden === true && hi.opts.length === 0, J(hi));
      await b.ev(`document.getElementById('quoteModalOverlay').style.display='none'; 0`);
      // 已核准：唯讀
      await openEdit(b, 'qAppr');
      hi = await hostInfo(b);
      t('QR-8 已核准的報價單：唯讀顯示「續約自：品項（到期 …）」，沒有下拉；註明僅內部可見', hi.exists && !hi.hidden && /mar-renew-locked/.test(hi.cls) && hi.val === null && /續約自：Svc MA 一年（到期 \d{4}\/\d{2}\/\d{2}）/.test(hi.text) && /不會印在客戶版報價單/.test(hi.text), J(hi));
      // 已核准單改簽核涵蓋欄位（會作廢核准）→ 儲存並作廢：renewOf 保留
      await b.focus('#qProjectName'); await b.ev(`document.getElementById('qProjectName').select(); 0`); await b.type('已核准後改名');
      await b.click('#quoteModalSave');
      await b.waitFor(`!![...document.querySelectorAll('button')].find(x=>x.textContent.trim()==='儲存並作廢核准')`, 8000);
      await b.click(() => [...document.querySelectorAll('button')].find((x) => x.textContent.trim() === '儲存並作廢核准'));
      await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='none'`, 15000); await b.sleep(900);
      const qa = (await H.http(U.own, 'GET', '/api/quotations/qAppr')).json;
      t('QR-9 已核准單修改簽核涵蓋欄位並「儲存並作廢核准」（WILL_VOID 重送路徑）：續約對象 iG 保留', qa.renewOf === 'iG' && qa.projectName === '已核准後改名', J([qa.renewOf, qa.projectName, qa.approval && qa.approval.state]));
      // 不顯示給客戶：預覽與簽核面板
      await toQuotes(b);
      await b.ev(`previewQuote('qDraft'); 0`); await b.sleep(2200);
      const pv = await b.ev(`document.body.innerText`);
      await b.ev(`(()=>{document.querySelectorAll('.modal-overlay, [id$="Overlay"]').forEach(o=>{ if(o.id!=='quoteModalOverlay') o.style.display='none'; }); return 0})()`);
      t('QR-10 報價單預覽（客戶看得到的內容）不含「續約自」「續約對象」', pv.length > 500 && !/續約自|續約對象|renewOf/.test(pv), pv.slice(0, 80));
      await b.ev(`openQuoteApproval('qDraft'); 0`); await b.sleep(2200);
      const ap = await b.ev(`document.body.innerText`);
      t('QR-11 簽核面板不顯示續約對象資訊', ap.length > 500 && !/續約自|續約對象|renewOf/.test(ap), ap.slice(0, 80));
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('QR-E 沒有 JS 例外', fe.length === 0, J(fe).slice(0, 300));
    });
  }

  // ═════════════════════════ look：外觀與版面 ═════════════════════════
  const LOOK_BELL = ['.mar-bell-row', '.mar-bell-co', '.mar-bell-pd', '.mar-kind', '.mar-lv-info', '.mar-lv-warn', '.mar-lv-danger', '.mar-bell-rule', '.mar-bell-roles', '.mar-bell-openpage', '.mar-bell-sum .notif-title', '.mar-bell-sum .notif-body'];
  const LOOK_ITEMS = ['.mar-badge-tracking', '.mar-badge-not_renew', '.mar-badge-renewed', '.mar-next', '.mar-reason', '.mar-tl-name', '.mar-tl-date', '.mar-tl-when', '.mar-tl-meta', '.mar-tl-state-active', '.mar-tl-state-future', '.mar-tl-state-cond', '.mar-tl-sub', '.mat-btn.mar-btn-tl'];
  const LOOK_RENEW = ['.mar-renew-title', '.mar-renew-sub', '.mar-renew-label', '.mar-renew-select'];
  const LOOK_ADMIN = ['.mar-a-name', '.mar-a-sel', '.mar-a-btn', '.mar-a-btn-primary', '.mar-a-btn-warn', '.mar-a-btn-danger', '.mar-a-pill', '.mar-a-help', '.mar-a-title', '.mar-a-role-msg', '.mar-a-role-rules', '.mar-a-role-h', '.mar-a-status', '.mar-a-ferr', '.mar-a-error', '.mar-a-lbl', '.mar-a-daytxt', '.mar-a-chkl', '.mar-a-tlstate-on', '.mar-a-tlstate-off', '.mar-a-unit', '.mar-a-tllist li', '.mar-a-legend', 'svg .mar-a-txt', 'svg .mar-a-txt-sm', 'svg .mar-a-txt-today', '.mar-a-confirm', '.mar-a-note'];
  async function contrastAll(b, sels, theme_, label, minOk) {
    const out = {}; let worst = 99, worstSel = '';
    for (const sel of sels) {
      const c = await contrastSel(b, sel, theme_.endsWith('dark'));
      if (c === null) { out[sel] = null; continue; }
      out[sel] = c;
      if (c < worst) { worst = c; worstSel = sel; }
    }
    const seen = Object.keys(out).filter((k) => out[k] !== null).length;
    t('LK-C ' + label + '（' + theme_ + '）：全部文字對比 ≥ 4.5（量到 ' + seen + ' 種元素，最低 ' + worst + '：' + worstSel + '）', worst >= 4.5 && seen >= Math.min(minOk, sels.length - 4), J(out).slice(0, 500));
  }
  async function secLook(C) {
    const { H } = C;
    const notifs = [{ id: 'nNR', to: U.own, type: 'ma_not_renew', title: '🚫 R丁公司 MA 不續約', body: '原因：預算不足', refId: '/index.html#maitems', read: false, createdAt: agoIso(1) }];
    await H.fresh(mkUsers(H), seedData({ notifs }), { env: { MA_TRACKING: '1' } });
    await H.loginAll([U.own, U.adm, U.m1]);
    // 讓每一種顏色（info／warn／danger 徽章與狀態）都有元素可量：預先把 iG 標不續約（徽章）、管理員頁有錯誤與確認列
    await browse(C, 900000, async (b) => {
      await openUser(b, C, U.own, 1440, 1000, { expectOn: true, clearLs: true });
      await toItems(b);
      await b.click(btnIn('iA', '提醒時間軸'));
      await b.waitFor(`document.querySelectorAll('#maItemsView tr.mar-tl-tr .mar-tl-li').length===4`, 8000);
      await b.click(btnIn('iC', '提醒時間軸')); await b.waitFor(`document.querySelectorAll('#maItemsView tr.mar-tl-tr').length===2`, 8000);
      // 條件不符的狀態（cond）：iB 剩 20 天且草稿 → D-45 條件不成立
      writeData((d) => { d.quotations.push({ id: 'qLk', quoteNo: 'QU-LK', owner: U.own, oppId: 'oR2', company: 'R乙公司', status: 'draft', approval: null, items: [], renewOf: 'iB', products: [], createdAt: NOW, updatedAt: NOW }); });
      await b.click(btnIn('iB', '提醒時間軸')); await b.waitFor(`document.querySelectorAll('#maItemsView tr.mar-tl-tr').length===3`, 8000);
      await b.sleep(600);
      // 鈴鐺打開並展開
      await b.ev(`localStorage.clear(); 0`); await b.ev(`_refreshBell(); 0`); await b.sleep(1100);
      await b.ev(`(()=>{const d=document.getElementById('notifDropdown'); d.style.display=''; const s=document.querySelector('#notifList .mar-bell-sum'); if(s && s.getAttribute('aria-expanded')!=='true') s.click(); return 0})()`); await b.sleep(500);
      for (const th of THEMES) {
        await theme(b, th);
        await contrastAll(b, LOOK_BELL, th, '鈴鐺 MA／License 群組', 8);
        await contrastAll(b, LOOK_ITEMS, th, '到期管理頁（徽章、下一個提醒、時間軸）', 10);
        await shotTo(C, b, 'look_bell_items_' + th + '.png');
      }
      // 層級
      const z = await b.ev(`JSON.stringify({mar:[...document.querySelectorAll('[class*="mar-"]')].map(e=>getComputedStyle(e).zIndex).filter(v=>v!=='auto'), dd:getComputedStyle(document.getElementById('notifDropdown')).zIndex})`).then(JSON.parse);
      t('LK-Z1 .mar-* 元素沒有自己的 z-index（不會蓋過任何彈窗）；鈴鐺下拉層級 999 低於 .mat-overlay 100000', z.mar.length === 0 && z.dd === '999', J(z));
      await theme(b, 'glass-dark');
      // 對話框層級：標記不續約對話框（.mat-overlay）低於 company-card-tasks 100001
      await b.ev(`document.getElementById('notifDropdown').style.display='none'; 0`);
      await b.click(btnIn('iF', '標記不續約…')); await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      const oz = await b.ev(`getComputedStyle(document.querySelector('.mat-overlay')).zIndex`);
      t('LK-Z2 標記不續約對話框沿用 .mat-overlay（z-index 100000，低於 company-card-tasks 100001／company-requests 100002）', oz === '100000', oz);
      await shotTo(C, b, 'look_norenew_dialog.png');
      await b.key('Escape');
      // reduced motion
      await b.media([{ name: 'prefers-reduced-motion', value: 'reduce' }]); await b.sleep(300);
      const rm = await b.ev(`JSON.stringify([...document.querySelectorAll('[class*="mar-"]')].map(e=>getComputedStyle(e)).filter(c=>c.animationName!=='none' || (c.transitionDuration!=='0s' && c.transitionDuration!=='')).length)`);
      t('LK-RM prefers-reduced-motion：.mar-* 元素沒有動畫或過渡', rm === '0', rm);
      await b.media([]);
      // 報價單橫幅外觀
      await toQuotes(b);
      await b.ev(`openQuoteModal(null,{oppId:'oR1'}); 0`);
      await b.waitFor(`document.querySelectorAll('#quoteItemsBody tr').length>0`, 15000); await b.sleep(1500);
      for (const th of THEMES) { await theme(b, th); await contrastAll(b, LOOK_RENEW, th, '報價單「這是續約嗎？」', 3); }
      await theme(b, 'glass-dark');
      await shotTo(C, b, 'look_quote_renew_glass_dark.png');
      await theme(b, 'plain-light'); await shotTo(C, b, 'look_quote_renew_plain_light.png'); await theme(b, 'glass-dark');
      await b.ev(`document.getElementById('quoteModalOverlay').style.display='none'; 0`);
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('LK-E1 沒有 JS 例外（使用者端外觀）', fe.length === 0, J(fe).slice(0, 300));
    });
    // 手機寬度：使用者端
    await browse(C, 500000, async (b) => {
      await openUser(b, C, U.own, 375, 800, { expectOn: true, clearLs: true });
      await b.ev(`_refreshBell(); 0`); await b.sleep(1100);
      await b.ev(`(()=>{const d=document.getElementById('notifDropdown'); d.style.display=''; const s=document.querySelector('#notifList .mar-bell-sum'); if(s && s.getAttribute('aria-expanded')!=='true') s.click(); return 0})()`); await b.sleep(600);
      const mb = await b.ev(`(()=>{const d=document.getElementById('notifDropdown').getBoundingClientRect(); const rows=[...document.querySelectorAll('#notifList .mar-bell-row')].map(r=>{const q=r.getBoundingClientRect(); return [q.left,q.right]}); return JSON.stringify({sw:document.documentElement.scrollWidth, w:innerWidth, dl:d.left, dr:d.right, rowsOk:rows.every(x=>x[0]>=-1&&x[1]<=innerWidth+1), n:rows.length})})()`).then(JSON.parse);
      t('LK-M1 手機寬度 375：鈴鐺下拉與 MA／License 各列都在螢幕內、頁面不橫向捲動', mb.sw <= mb.w + 1 && mb.dl >= -1 && mb.dr <= mb.w + 1 && mb.rowsOk && mb.n >= 3, J(mb));
      await shotTo(C, b, 'mobile_bell.png');
      await b.ev(`document.getElementById('notifDropdown').style.display='none'; showSection('maItems'); 0`);
      await b.waitFor(`document.querySelectorAll('#maItemsView tbody tr').length>0`, 12000); await b.sleep(1600);
      await b.click(btnIn('iA', '提醒時間軸')); await b.sleep(1200);
      const mi = await b.ev(`JSON.stringify({sw:document.documentElement.scrollWidth, w:innerWidth, head:getComputedStyle(document.querySelector('#maItemsView thead')).display, tl:!!document.querySelector('#maItemsView .mar-tl-li'), tlOk:[...document.querySelectorAll('#maItemsView .mar-tl-li')].every(r=>r.getBoundingClientRect().right<=innerWidth+1), next:!!document.querySelector('#maItemsView .mar-next')})`).then(JSON.parse);
      t('LK-M2 手機寬度：到期管理頁（卡片式）含下一個提醒與時間軸不爆版', mi.sw <= mi.w + 1 && mi.head === 'none' && mi.tl && mi.tlOk && mi.next, J(mi));
      await shotTo(C, b, 'mobile_items_timeline.png');
      await b.sleep(1200);   // 列高會隨「下一個提醒」載入而變動：等穩定再點
      await b.click(btnIn('iB', '標記不續約…'));
      if (!(await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 3000))) await b.click(btnIn('iB', '標記不續約…'));
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      const md = await b.ev(`(()=>{const d=document.querySelector('.mat-overlay .mat-dialog').getBoundingClientRect(); return JSON.stringify({l:d.left,r:d.right,w:innerWidth,sw:document.documentElement.scrollWidth})})()`).then(JSON.parse);
      t('LK-M3 手機寬度：標記不續約對話框不超出螢幕', md.l >= -1 && md.r <= md.w + 1 && md.sw <= md.w + 1, J(md));
      await shotTo(C, b, 'mobile_norenew_dialog.png');
      await b.key('Escape');
      await toQuotes(b);
      await b.ev(`openQuoteModal(null,{oppId:'oR1'}); 0`);
      await b.waitFor(`document.querySelectorAll('#quoteItemsBody tr').length>0 && !!document.getElementById('qRenewHost') && !document.getElementById('qRenewHost').hidden`, 15000); await b.sleep(800);
      const mq = await b.ev(`(()=>{const h=document.getElementById('qRenewHost'); const r=h.getBoundingClientRect(); const s=h.querySelector('select').getBoundingClientRect(); return JSON.stringify({vis:!h.hidden, l:r.left, r:r.right, sr:s.right, w:innerWidth, sw:document.documentElement.scrollWidth})})()`).then(JSON.parse);
      t('LK-M4 手機寬度：報價單「這是續約嗎？」橫幅與下拉不超出螢幕', mq.vis && mq.l >= -1 && mq.r <= mq.w + 1 && mq.sr <= mq.w + 1 && mq.sw <= mq.w + 1, J(mq));
      await shotTo(C, b, 'mobile_quote_renew.png');
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('LK-E2 沒有 JS 例外（手機）', fe.length === 0, J(fe).slice(0, 300));
    });
    // 後台外觀：四種外觀對比＋手機
    await browse(C, 700000, async (b) => {
      await openAdmin(b, C, U.adm, 1440, 1000, true);
      await gotoRules(b); await b.sleep(900);
      // 製造一個錯誤列（顯示 .mar-a-ferr／.mar-a-error）與確認列
      await typeInto(b, rowS(1) + ' .mar-a-name', '');
      await b.click('#marAdminHost .mar-a-btn-warn'); await b.sleep(300);
      await b.ev(`document.querySelector('input[name="marSimQs"][value="draft"]').click(); 0`); await b.sleep(900);
      await b.ev(`(()=>{const s=document.getElementById('marSimItem'); s.value='iA'; s.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`); await b.sleep(1200);
      // 試算器需要規則合法才會有角色卡：先把名稱補回去
      await typeInto(b, rowS(1) + ' .mar-a-name', '45 天提醒'); await b.sleep(1200);
      await b.ev(`document.querySelector('#marAdminHost .mar-a-confirm').hidden=false; 0`);
      for (const th of THEMES) {
        await theme(b, th);
        await contrastAll(b, LOOK_ADMIN.filter((s) => s !== '.mar-a-ferr' && s !== '.mar-a-error' && s !== '.mar-a-note'), th, '後台提醒規則頁', 18);
        await shotTo(C, b, 'look_admin_' + th + '.png');
      }
      // 錯誤狀態的對比（ferr／error）單獨量
      await typeInto(b, rowS(1) + ' .mar-a-name', '');
      for (const th of THEMES) { await theme(b, th); await contrastAll(b, ['.mar-a-ferr', '.mar-a-error', '.mar-a-status'], th, '後台錯誤訊息', 2); }
      await typeInto(b, rowS(1) + ' .mar-a-name', '45 天提醒');
      const z = await b.ev(`JSON.stringify([...document.querySelectorAll('.mar-a, .mar-a *')].map(e=>getComputedStyle(e).zIndex).filter(v=>v!=='auto'))`);
      t('LK-Z3 後台 .mar-a-* 沒有 z-index', z === '[]', z);
      await theme(b, 'glass-dark');
      // reduced motion
      await b.media([{ name: 'prefers-reduced-motion', value: 'reduce' }]); await b.sleep(300);
      const rm = await b.ev(`JSON.stringify([...document.querySelectorAll('.mar-a, .mar-a *')].map(e=>getComputedStyle(e)).filter(c=>c.animationName!=='none' || (c.transitionDuration!=='0s' && c.transitionDuration!=='')).length)`);
      t('LK-RM2 後台 prefers-reduced-motion：沒有動畫或過渡', rm === '0', rm);
      await b.media([]);
      // 手機
      await b.viewport(375, 800, true); await b.sleep(900);
      const am = await b.ev(`JSON.stringify({sw:document.documentElement.scrollWidth, w:innerWidth, head:getComputedStyle(document.querySelector('#marAdminHost .mar-a-tbl thead')).display, wrap:document.querySelector('#marAdminHost .mar-a-tblwrap').scrollWidth<=document.querySelector('#marAdminHost .mar-a-tblwrap').clientWidth+1, svgBox:(()=>{const w=document.querySelector('#marAdminHost .mar-a-tlwrap'); return w.getBoundingClientRect().right<=innerWidth+1})(), simctl:document.querySelector('#marAdminHost .mar-a-simctl').getBoundingClientRect().right<=innerWidth+1, roles:[...document.querySelectorAll('#marAdminHost .mar-a-role')].every(r=>r.getBoundingClientRect().right<=innerWidth+1)})`).then(JSON.parse);
      t('LK-M5 手機寬度 375：後台提醒規則頁不爆版（規則表改成卡片式、時間軸在自己的捲動區內、試算控制與角色卡都在螢幕內）', am.sw <= am.w + 1 && am.head === 'none' && am.wrap && am.svgBox && am.simctl && am.roles, J(am));
      await shotTo(C, b, 'mobile_admin_rules.png');
      await b.ev(`document.getElementById('marSimDays').scrollIntoView({block:'start'}); 0`); await b.sleep(400);
      await shotTo(C, b, 'mobile_admin_simulator.png');
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('LK-E3 沒有 JS 例外（後台）', fe.length === 0, J(fe).slice(0, 300));
    });
  }

  return { secDual, secBell, secItems, secQuote, secLook };
};
