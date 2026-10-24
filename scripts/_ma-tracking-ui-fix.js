'use strict';
/**
 * MA 納管前端「3020 驗收後修正」區段（FX-*）。由 scripts/_ma-tracking-ui.js 載入（共用其小工具），不單獨執行。
 *   FX-1  沒有任務卡的 License 納管項＝追蹤中、不算待簽約（摘要卡、篩選、列內文字、沒有處理簽約鈕）
 *   FX-2  刪除 Service MA 合約 → 任務卡重開，首頁待辦／到期管理／鈴鐺刷新
 *   FX-3  改派候選人只含在職業務、同 BU（後端另驗）
 *   FX-4  簽約完成後 Service MA 頁、到期管理、首頁待辦、鈴鐺自動刷新
 *   FX-5  後台「銷售商品維護」儲存進行中離開頁面會被攔
 *   FX-6  合約頁與到期管理頁剩餘天數同一函式（台北日曆日；多個瀏覽器時區）
 *   FX-7  再成交時期間預填（上次期間；已過期則提示請確認期間）
 *   FX-8  到期管理窄視窗：≥481px 橫向捲動且期間不斷行、≤480px 卡片式
 *   FX-9  審查修正 M1：管理員在到期管理列「改負責人…」（含沒有任務卡的 License 類）；非管理員沒有這個按鈕
 *   FX-10 審查修正 L1：行銷（本來就看不到合約）的到期管理頁沒有「合約編號」欄
 */
const J = JSON.stringify;

module.exports = function makeFix(X) {
  const { t, browse, openAs, gotoItemsNav, countsOf, rowCos, panelRows, topText, clickBtn, focusTop, dateVals, drop, toBoard, oppOf, shotTo, mkUsers, mkOpp, mkItem, seedData, U, tp, MAP, LICP } = X;

  const hasBtnTxt = (b, txt) => b.ev(`(()=>{const o=[...document.querySelectorAll('.mat-overlay')].pop(); return !!o && [...o.querySelectorAll('button')].some(x=>x.textContent.trim()===${J(txt)})})()`);
  const RE_DAYS = (n) => new RegExp('(剩餘|剩) ' + n + ' 天');

  async function secFix(C) {
    const { H } = C;
    const users = mkUsers(H).concat([H.mkUser('_s_its', 'user', { bu: ['ITS'], displayName: '測試業務ITS' })]);
    const d = seedData();
    d.opportunities.push(mkOpp('oL', U.own, LICP, 'Won', 'L公司', { achievedDate: tp(-3) }));
    d.maItems.push(mkItem('iL', 'oL', U.own, 'L公司', LICP, tp(150), { termKind: 'license' }));   // License：沒有任務卡、沒有合約
    d.opportunities.push(mkOpp('oRe', U.own, MAP, 'B', 'RE公司'), mkOpp('oRx', U.own, MAP, 'B', 'RX公司'));
    d.maItems.push(mkItem('iRe', 'oRe', U.own, 'RE公司', MAP, tp(335), { from: tp(-30), status: 'void' }), mkItem('iRx', 'oRx', U.own, 'RX公司', MAP, tp(-35), { from: tp(-400), status: 'void' }));
    await H.fresh(users, d, { env: { MA_TRACKING: '1' } });
    await H.loginAll([U.own, U.adm]);

    // ── 到期管理＋Service MA ──
    await browse(C, 560000, async (b) => {
      await openAs(b, C, U.own, 1440, 900, { expectOn: true, noLoginFlag: true });
      await gotoItemsNav(b);
      const cs = await countsOf(b);
      t('FX-1a 摘要卡：License（無任務卡）算納管中但不算待簽約：納管中 6、待簽約 2（只有 A、B）', J(cs) === J(['納管中:6', '待簽約:2', '90 天內到期:3', '已逾期:1']), J(cs));
      const lRow = await b.ev(`[...document.querySelectorAll('#maItemsView tbody tr')].find(x=>/L公司/.test(x.innerText)).innerText.replace(/\\s+/g,' ')`);
      t('FX-1b L公司（License、無任務卡）狀態顯示「追蹤中」，不是「待簽約確認」，沒有「處理簽約…」', /追蹤中/.test(lRow) && !/待簽約確認/.test(lRow) && !/處理簽約/.test(lRow), lRow);
      await b.ev(`(()=>{const x=[...document.querySelectorAll('#maItemsView .mat-fchip')].find(c=>c.textContent==='待簽約'); x.click(); return 0})()`); await b.sleep(900);
      const fp = await rowCos(b);
      t('FX-1c 「待簽約」篩選：只有 A公司、B公司（L公司不在）', J(fp.slice().sort()) === J(['A公司', 'B公司']), J(fp));
      await b.ev(`(()=>{const x=[...document.querySelectorAll('#maItemsView .mat-fchip')].find(c=>c.textContent==='全部'); x.click(); return 0})()`); await b.sleep(700);
      const api = ((await H.http(U.own, 'GET', '/api/ma-items?filter=pending_sign')).json.items || []).map((x) => x.company).sort();
      t('FX-1d 後端 pending_sign 與畫面一致（A、B）', J(api) === J(['A公司', 'B公司']), J(api));

      // FX-6：同一份合約（C公司，迄日＝台北今天+200）兩頁顯示同一天數
      const itemsC = await b.ev(`[...document.querySelectorAll('#maItemsView tbody tr')].find(x=>/C公司/.test(x.innerText)).innerText.replace(/\\s+/g,' ')`);
      await b.ev(`showSection('service-ma'); 0`);
      await b.waitFor(`document.querySelectorAll('#serviceMaTbody tr').length>0 && /C-2/.test(document.getElementById('serviceMaTbody').innerText)`, 10000);
      await b.sleep(400);
      const smC = await b.ev(`document.querySelector('#serviceMaTbody tr').innerText.replace(/\\s+/g,' ')`);
      t('FX-6a Service MA 頁與到期管理頁顯示同一天數（200 天，迄日為台北今天+200）', RE_DAYS(200).test(smC) && RE_DAYS(200).test(itemsC), J([smC.slice(-60), itemsC.slice(-60)]));
      const tz = await b.ev(`JSON.stringify([ITTSMaTasks.daysLeft('2026-10-25', Date.parse('2026-10-24T16:00:00Z')), ITTSMaTasks.daysLeft('2026-10-24', Date.parse('2026-10-24T16:00:00Z')), ITTSMaTasks.daysLeft('2026-10-25', Date.parse('2026-10-24T15:59:59Z')), ITTSMaTasks.daysLeft('2026-02-30', Date.now()), ITTSMaTasks.daysLeft('', Date.now())])`).then(JSON.parse);
      t('FX-6b 共用函式（台北日曆日）：迄日當天＝0、台北已過午夜時前一天迄日＝-1、跨日前一秒仍＝1、假日期／空值＝null', J(tz) === J([0, -1, 1, null, null]), J(tz));
      for (const zone of ['Pacific/Honolulu', 'Asia/Taipei', 'Pacific/Auckland', 'America/Los_Angeles']) {
        await b.send('Emulation.setTimezoneOverride', { timezoneId: zone });
        const days = await b.ev(`JSON.stringify([contractStatus({endDate: ${J(tp(200))}, startDate: ${J(tp(-100))}}).days, contractStatus({endDate: ${J(tp(0))}}).days, contractStatus({endDate: ${J(tp(-1))}}).days, ITTSMaTasks.daysLeft(${J(tp(200))})])`).then(JSON.parse);
        t('FX-6c 瀏覽器時區 ' + zone + '：合約頁 contractStatus 的天數 = 到期管理共用函式 = 200／0（迄日當天）／逾期 1 天', J(days) === J([200, 0, 1, 200]), J(days));
      }
      await b.send('Emulation.setTimezoneOverride', { timezoneId: 'Asia/Taipei' });

      // FX-4：停在 Service MA 頁，從任務卡簽 A公司 → 列表自動多一筆；到期管理／我的任務／鈴鐺跟著刷新
      await b.ev(`window.__bell=0; const ob=window._refreshBell; window._refreshBell=function(){window.__bell++; return ob.apply(this,arguments)}; performance.clearResourceTimings(); 0`);
      const before = await b.ev(`document.querySelectorAll('#serviceMaTbody tr').length`);
      await b.ev(`ITTSMaTasks.openSummary({fetch:true}); 0`);
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
      await b.click(() => { const r = [...document.querySelectorAll('.mat-overlay .mat-row')].find((x) => /A公司/.test(x.innerText)); return r && [...r.querySelectorAll('button')].find((x) => x.textContent.trim() === '處理…'); });
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===2`, 5000);
      await clickBtn(b, '合約已簽訂，輸入編號');
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===3`, 5000);
      await focusTop(b, 'input[type=text]'); await b.type('FX-NEW-1');
      await clickBtn(b, '確認簽約');
      await b.waitFor(`[...document.querySelectorAll('.mat-dlg-title')].some(x=>/已歸入合約管理/.test(x.textContent))`, 8000);
      await b.sleep(1200);
      const smNos = await b.ev(`[...document.querySelectorAll('#serviceMaTbody tr')].map(r=>r.querySelector('td strong')?r.querySelector('td strong').textContent:'')`);
      t('FX-4a 簽約完成後 Service MA 頁自動刷新（原 ' + before + ' 筆 → 含 FX-NEW-1 共 2 筆），不必重新進入頁面', smNos.length === 2 && smNos.includes('FX-NEW-1') && smNos.includes('C-2'), J(smNos));
      const net = await b.ev(`JSON.stringify(performance.getEntriesByType('resource').map(e=>e.name.replace(location.origin,'')).filter(u=>/\\/api\\/(contracts|ma-items|my-tasks\\/deals)/.test(u)))`).then(JSON.parse);
      t('FX-4b 簽約完成後重新取得：合約列表、到期管理、我的任務（首頁待辦）', net.some((u) => /\/api\/contracts/.test(u)) && net.some((u) => /\/api\/ma-items/.test(u)) && net.some((u) => /\/api\/my-tasks\/deals/.test(u)), J(net));
      t('FX-4c 鈴鐺同步刷新（_refreshBell 被呼叫）', (await b.ev(`window.__bell`)) >= 1);
      await b.ev(`document.querySelectorAll('.mat-overlay').forEach(o=>o.remove()); 0`);

      // FX-2：從 Service MA 頁刪除剛簽的合約 → 任務卡重開；前端刷新
      await b.ev(`window.__bell=0; performance.clearResourceTimings(); 0`);
      await b.click(() => { const r = [...document.querySelectorAll('#serviceMaTbody tr')].find((x) => /FX-NEW-1/.test(x.innerText)); return r && r.querySelector('.ct-btn-delete'); });
      await b.waitFor(`document.getElementById('contractDeleteModal').classList.contains('open')`, 5000);
      await b.click('#contractDeleteConfirm');
      await b.sleep(1800);
      const tk = (await H.http(U.own, 'GET', '/api/my-tasks/deals')).json.items || [];
      const tA = tk.find((x) => x.id === 'tA');
      t('FX-2a 刪除合約後：伺服器任務卡 tA 重開（open），出現在「我的任務」', !!tA && tA.status === 'open', J(tA && tA.status));
      const itA = ((await H.http(U.own, 'GET', '/api/ma-items?filter=pending_sign')).json.items || []).find((x) => x.company === 'A公司');
      t('FX-2b 到期管理：A公司 重新是待簽約（pendingSign、任務 open、合約編號清空）', !!itA && itA.pendingSign === true && itA.taskStatus === 'open' && itA.contractNo === '', J(itA));
      const net2 = await b.ev(`JSON.stringify(performance.getEntriesByType('resource').map(e=>e.name.replace(location.origin,'')).filter(u=>/\\/api\\/(ma-items|my-tasks\\/deals)/.test(u)))`).then(JSON.parse);
      t('FX-2c 刪除 Service MA 合約後前端重新取得到期管理與我的任務；鈴鐺刷新', net2.some((u) => /\/api\/ma-items/.test(u)) && net2.some((u) => /\/api\/my-tasks\/deals/.test(u)) && (await b.ev(`window.__bell`)) >= 1, J(net2));
      await b.ev(`showDashboard(); 0`); await b.sleep(1500);
      t('FX-2d 首頁待辦卡又列出 A公司', (await panelRows(b)).includes('A公司'), J(await panelRows(b)));
      const nt = ((await H.http(U.own, 'GET', '/api/notifications')).json || []).filter((n) => n.type === 'deal_task' && n.refId === 'tA' && !n.read);
      t('FX-2e 負責業務收到重開通知（未讀、不含金額 120）', nt.length >= 1 && nt.every((n) => !/120/.test(n.body)), J(nt.map((n) => n.title)));
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('FX-E1 沒有 JS 例外（到期管理／Service MA 區段）', fe.length === 0, J(fe).slice(0, 300));
    });

    // ── 改派候選人（FX-3）──
    await browse(C, 300000, async (b) => {
      await openAs(b, C, U.adm, 1440, 900, { expectOn: true, noLoginFlag: true });
      await b.ev(`showSection('managerHome'); 0`);
      await b.waitFor(`!!document.querySelector('.mat-panel-overdue')`, 10000);
      await b.click(() => { const r = [...document.querySelectorAll('.mat-panel-overdue .mat-row')].find((x) => /E離職公司/.test(x.innerText)); return r && r.querySelector('button'); });
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      await b.waitFor(`!!__top('select') && __top('select').options.length>0`, 8000);
      const opts = await b.ev(`[...__top('select').options].map(o=>o.value).sort().join()`);
      t('FX-3a 改派下拉（負責人已停用，以商機 BU=ERP 為準）：只有 _s_other／_s_own；沒有 ITS 業務、一級主管、經營者、行銷、秘書、管理員', opts === '_s_other,_s_own', opts);
      const r400 = await H.http(U.adm, 'POST', '/api/deal-tasks/tOff/reassign', { to: '_s_its' });
      const r400b = await H.http(U.adm, 'POST', '/api/deal-tasks/tOff/reassign', { to: U.m1 });
      t('FX-3b 直接打後端改派給別的 BU 業務／主管：400（前端限制不是唯一防線）', r400.status === 400 && r400.json.code === 'ASSIGNEE_BU' && r400b.status === 400 && r400b.json.code === 'ASSIGNEE_ROLE', J([r400.status, r400.json.code, r400b.status, r400b.json.code]));
      await b.key('Escape');
    });

    // ── 管理員更換納管項負責業務（審查修正 M1）──
    await browse(C, 300000, async (b) => {
      await openAs(b, C, U.adm, 1440, 900, { expectOn: true, noLoginFlag: true });
      await gotoItemsNav(b);
      const btnOf = (company) => '()=>{ const r = [...document.querySelectorAll("#maItemsView tbody tr")].find((x) => x.querySelector("strong") && x.querySelector("strong").textContent === ' + J(company) + '); return r && [...r.querySelectorAll("button")].find((x) => x.textContent.trim() === "改負責人…"); }';
      t('FX-9a 管理員在到期管理列看得到「改負責人…」（L公司：License、沒有任務卡）', await b.ev(`!!(${btnOf('L公司')})()`));
      await b.click(btnOf('L公司'));
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 5000);
      await b.waitFor(`!!__top('select') && __top('select').options.length>0`, 8000);
      const opts9 = await b.ev(`[...__top('select').options].map(o=>o.value).sort().join()`);
      t('FX-9b 候選人只有在職、同 BU 的業務（_s_other；不含目前負責人 _s_own、ITS 業務、主管、行銷）', opts9 === '_s_other', opts9);
      t('FX-9c 對話框說明：同一商機底下的納管項與簽約任務卡一起換', /一起換成新的負責業務/.test(await topText(b)) && (await hasBtnTxt(b, '確認更換')));
      await clickBtn(b, '取消');
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 5000);
      t('FX-9d 按「取消」不送出：納管項負責人不變', ((await H.http(U.adm, 'GET', '/api/ma-items')).json.items.find((x) => x.company === 'L公司') || {}).owner === U.own);
      await b.click(btnOf('L公司'));
      await b.waitFor(`!!__top('select') && __top('select').options.length>0`, 8000);
      await clickBtn(b, '確認更換');
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===0`, 8000);
      await b.sleep(1500);
      const it9 = (await H.http(U.adm, 'GET', '/api/ma-items')).json.items.find((x) => x.company === 'L公司');
      const rowTxt9 = await b.ev(`[...document.querySelectorAll('#maItemsView tbody tr')].find(x=>/L公司/.test(x.innerText)).innerText.replace(/\\s+/g,' ')`);
      t('FX-9e 更換後：後端納管項 owner＝_s_other、畫面該列刷新顯示新負責業務', !!it9 && it9.owner === U.oth && /測試業務乙/.test(rowTxt9), J([it9 && it9.owner, rowTxt9.slice(-60)]));
      const aud9 = ((await H.http(U.adm, 'GET', '/api/admin/logs')).json || []).filter((l) => l.action === 'MA_ITEM_REASSIGN');
      t('FX-9f 稽核 MA_ITEM_REASSIGN 一筆（舊→新）', aud9.length === 1 && new RegExp(U.own + ' → ' + U.oth).test(aud9[0].detail || ''), J(aud9.map((x) => x.detail)));
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('FX-E5 沒有 JS 例外（改負責人）', fe.length === 0, J(fe).slice(0, 300));
    });
    await browse(C, 300000, async (b) => {
      await openAs(b, C, U.oth, 1440, 900, { expectOn: true, noLoginFlag: true });
      await gotoItemsNav(b);
      t('FX-9g 負責業務（非管理員）看不到「改負責人…」；剛被換過來的 L公司 出現在他的清單', (await b.ev(`[...document.querySelectorAll('#maItemsView tbody button')].filter(x=>x.textContent.trim()==='改負責人…').length`)) === 0 && (await rowCos(b)).includes('L公司'));
      await openAs(b, C, U.mk, 1440, 900, { expectOn: true, noLoginFlag: true });
      await gotoItemsNav(b);
      const hdr10 = await b.ev(`[...document.querySelectorAll('#maItemsView thead th')].map(x=>x.textContent).join('|')`);
      t('FX-10a 行銷：欄位沒有「合約編號」，其餘照舊', hdr10 === '客戶|品項|期間|剩餘天數|狀態／下一步|負責業務|操作', hdr10);
      t('FX-10b 行銷畫面沒有任何合約編號文字（C-2／C-3）；也沒有「改負責人…」', !/C-2|C-3/.test(await b.ev(`document.getElementById('maItemsView').innerText`)) && (await b.ev(`[...document.querySelectorAll('#maItemsView tbody button')].filter(x=>x.textContent.trim()==='改負責人…').length`)) === 0);
    });

    // ── 再成交期間預填（FX-7）──
    await browse(C, 300000, async (b) => {
      await openAs(b, C, U.own, 1440, 900, { expectOn: true, noLoginFlag: true });
      await toBoard(b);
      await drop(b, 'oRe', 'Won');
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
      await b.sleep(400);
      const dv = await dateVals(b); const tx = await topText(b);
      t('FX-7a 再成交（上次期間未過）：視窗預填上次成交的起迄日，並說明來源、提醒確認期間', dv[0] === tp(-30) && dv[1] === tp(335) && /上次成交的服務期間/.test(tx) && /請確認期間/.test(tx), J([dv, tx.slice(0, 160)]));
      await clickBtn(b, '取消'); await b.sleep(400);
      const st1 = await oppOf(H, U.own, 'oRe');
      t('FX-7b 取消後商機仍是原階段', st1 && st1.stage === 'B', J(st1 && st1.stage));
      await drop(b, 'oRx', 'Won');
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
      await b.sleep(400);
      const dv2 = await dateVals(b); const tx2 = await topText(b);
      t('FX-7c 上次期間已過：照樣預填，但提示期間已經過了、請確認期間', dv2[0] === tp(-400) && dv2[1] === tp(-35) && /已經過了/.test(tx2) && /請確認期間/.test(tx2), J([dv2, tx2.slice(0, 200)]));
      await clickBtn(b, '取消'); await b.sleep(300);
      await drop(b, 'oK1', 'Won');
      await b.waitFor(`document.querySelectorAll('.mat-overlay').length===1`, 8000);
      await b.sleep(400);
      const dv3 = await dateVals(b); const tx3 = await topText(b);
      t('FX-7d 有核准報價單期間的商機：仍預填報價單期間與其說明（不被上次期間取代）', dv3[0] === '2026-11-01' && dv3[1] === '2027-10-31' && /核准報價單/.test(tx3), J([dv3, tx3.slice(0, 120)]));
      await clickBtn(b, '取消'); await b.sleep(300);
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('FX-E2 沒有 JS 例外（成交視窗）', fe.length === 0, J(fe).slice(0, 300));
    });

    // ── 到期管理窄視窗（FX-8）──
    await browse(C, 400000, async (b) => {
      for (const w of [900, 800, 600, 500]) {
        await openAs(b, C, U.own, w, 800, { expectOn: true, noLoginFlag: true, fast: true });
        await b.ev(`showSection('maItems'); 0`);
        await b.waitFor(`document.querySelectorAll('#maItemsView tbody tr').length>0`, 10000);
        await b.sleep(500);
        const m = await b.ev(`(()=>{const w=document.querySelector('#maItemsView .mat-tablewrap'); const td=[...document.querySelectorAll('#maItemsView tbody td[data-label="期間"]')][0]; const rg=document.createRange(); rg.selectNodeContents(td); const tops=new Set([...rg.getClientRects()].map(x=>Math.round(x.top))); return JSON.stringify({scrollable:w.scrollWidth>w.clientWidth+1, wrapW:w.clientWidth, tblW:w.querySelector('table').getBoundingClientRect().width, page:document.documentElement.scrollWidth<=innerWidth+1, lines:tops.size, vw:innerWidth, region:w.getAttribute('role')})})()`).then(JSON.parse);
        t('FX-8a 視窗 ' + w + 'px：期間欄一行顯示（不斷行）、表格維持欄寬並可橫向捲動、整頁不橫向爆版', m.lines === 1 && m.page && (m.scrollable || m.tblW <= m.wrapW + 1) && m.tblW >= Math.min(760, m.wrapW) - 1 && m.region === 'region', J(m));
      }
      for (const w of [480, 375]) {
        await openAs(b, C, U.own, w, 800, { expectOn: true, noLoginFlag: true, fast: true });
        await b.ev(`showSection('maItems'); 0`);
        await b.waitFor(`document.querySelectorAll('#maItemsView tbody tr').length>0`, 10000);
        await b.sleep(500);
        const m = await b.ev(`(()=>{const w=document.querySelector('#maItemsView .mat-tablewrap'); const th=document.querySelector('#maItemsView thead'); const td=document.querySelector('#maItemsView tbody td[data-label="期間"]'); const tr=document.querySelector('#maItemsView tbody tr'); return JSON.stringify({head:getComputedStyle(th).display, tdd:getComputedStyle(td).display, noScroll:w.scrollWidth<=w.clientWidth+1, page:document.documentElement.scrollWidth<=innerWidth+1, inView:tr.getBoundingClientRect().right<=innerWidth+1, vw:innerWidth})})()`).then(JSON.parse);
        t('FX-8b 視窗 ' + w + 'px（手機）：改為卡片式（表頭隱藏、儲存格區塊）、無橫向捲動、整頁不爆版', m.head === 'none' && m.tdd === 'block' && m.noScroll && m.page && m.inView, J(m));
        if (w === 375) await shotTo(C, b, 'ma_fix_items_375.png');
      }
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('FX-E3 沒有 JS 例外（窄視窗）', fe.length === 0, J(fe).slice(0, 300));
    });

    // ── 後台銷售商品維護（FX-5）──
    await browse(C, 300000, async (b) => {
      await b.viewport(1440, 900, false);
      await b.nav(H.BASE + '/login.html');
      await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(U.adm)},password:${J(H.PW)}})}).then(r=>r.status)`);
      await b.nav(H.BASE + '/admin.html');
      await b.waitFor(`typeof initProductCatalog==='function' && !!document.querySelector('[data-sec="product-catalog"]')`, 20000);
      await b.ev(`document.querySelector('[data-sec="product-catalog"]').click(); 0`);
      await b.waitFor(`_pcMa===true && !!document.querySelector('#pcPanel [data-pcact]')`, 15000);
      // 讓儲存請求慢 2.2 秒，模擬儲存進行中
      await b.ev(`(()=>{ const of=window.fetch.bind(window); window.fetch=(u,o)=>{ if(o&&o.method==='POST'&&/admin\\/product-catalog/.test(String(u))&&!/preview/.test(String(u))){ return new Promise(r=>setTimeout(()=>r(of(u,o)),2200)); } return of(u,o); }; return 0})()`);
      await b.click('#pcSaveBtn');
      await b.waitFor(`document.getElementById('pcSaveBtn').textContent.indexOf('儲存中')>=0`, 5000);
      const during = await b.ev(`(()=>{ const e=new Event('beforeunload',{cancelable:true}); window.dispatchEvent(e); const p=e.defaultPrevented; document.querySelector('[data-sec="users"]').click(); return JSON.stringify({prevented:p, pcActive:document.getElementById('sec-product-catalog').classList.contains('active'), usersActive:document.getElementById('sec-users').classList.contains('active')}) })()`).then(JSON.parse);
      t('FX-5a 儲存進行中：關閉／重新整理分頁會觸發離開確認（beforeunload 被攔）；切換到其他後台頁被擋，仍停在銷售商品維護', during.prevented === true && during.pcActive === true && during.usersActive === false, J(during));
      await b.waitFor(`/已儲存/.test(document.getElementById('pcStatus').textContent)`, 10000);
      const after = await b.ev(`(()=>{ const e=new Event('beforeunload',{cancelable:true}); window.dispatchEvent(e); const p=e.defaultPrevented; document.querySelector('[data-sec="users"]').click(); return JSON.stringify({prevented:p, usersActive:document.getElementById('sec-users').classList.contains('active')}) })()`).then(JSON.parse);
      t('FX-5b 儲存完成（顯示「已儲存」）後：不再攔截離開、可以切換頁面', after.prevented === false && after.usersActive === true, J(after));
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('FX-E4 沒有 JS 例外（後台）', fe.length === 0, J(fe).slice(0, 300));
    });
  }
  return { secFix };
};
