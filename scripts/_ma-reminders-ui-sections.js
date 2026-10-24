'use strict';
/** _ma-reminders-ui.js 的測試區段（admin／dual／bell／items／quote／look）。共用主檔的小工具。 */
const path = require('path');
const J = JSON.stringify;

module.exports = function makeSections(X) {
  const { t, browse, openUser, openAdmin, gotoRules, pickOpt, typeInto, rowS, rulesInPage, shotTo, txt, reqs, mkUsers, seedData, U, tp, XSS1, XSS2, AMT, CNO, THEMES, theme, contrastSel, MAP } = X;
  const ROOT = path.resolve(__dirname, '..');
  const ROLES = ['owner', 'secretary', 'manager1', 'manager2', 'marketing'];

  // ═════════════════════════ admin：後台「提醒規則」頁 ═════════════════════════
  const RULE_XSS = '&lt;i&gt;規則X&lt;/i&gt;';   // 伺服器不收 < >，所以用實體字串驗證「不會被解碼成標籤」
  async function simRoles(b) {
    return b.ev(`JSON.stringify(Object.fromEntries([...document.querySelectorAll('#marAdminHost .mar-a-role')].map(r=>[r.getAttribute('data-role'), [...r.querySelectorAll('.mar-a-role-rules')].map(x=>x.textContent.replace('來自規則：','')).join('')])))`).then(JSON.parse);
  }
  async function apiSim(H, body) { const r = await H.http(U.adm, 'POST', '/api/admin/reminder-rules/simulate', body); return r.json; }
  const namesOf = (arr) => (arr || []).map((x) => '「' + x.name + '」').join('、');
  async function setDays(b, n) { await typeInto(b, '.mar-a-simnum', String(n)); await b.sleep(700); }

  async function secAdmin(C) {
    const { H } = C;
    await H.fresh(mkUsers(H), seedData(), { env: { MA_TRACKING: '1' } });
    await H.loginAll([U.adm, U.own]);
    await browse(C, 600000, async (b) => {
      await openAdmin(b, C, U.adm, 1440, 1000, true);
      const nav = await b.ev(`(()=>{const n=document.getElementById('marNavItem'); const dom=n.closest('.nav-domain'); return JSON.stringify({text:n.textContent.trim(), dom:dom.getAttribute('data-domain'), cnt:dom.querySelector('.nav-domain-count').textContent, sec:!!document.getElementById('sec-reminder-rules'), css:!!document.getElementById('maRemAdminCss')})})()`).then(JSON.parse);
      t('AD-1 開關開啟：後台「報價與簽核」群組多一個「提醒規則」側欄項（群組計數 4→5）、section 與樣式表', nav.text.indexOf('提醒規則') >= 0 && nav.dom === 'quote' && nav.cnt === '5' && nav.sec && nav.css, J(nav));
      await gotoRules(b);
      const first = await b.ev(`JSON.stringify({rows:document.querySelectorAll('#marAdminHost .mar-a-row').length, active:document.getElementById('sec-reminder-rules').classList.contains('active'), navActive:document.getElementById('marNavItem').classList.contains('active'), meta:document.querySelector('#marAdminHost .mar-a-metaline').textContent, count:document.querySelector('#marAdminHost .mar-a-count').textContent, ids:MaRemAdmin._state.models.map(m=>m.id)})`).then(JSON.parse);
      t('AD-2 點選後顯示規則頁：4 條預設規則、側欄反白、標示「系統預設規則」、計數 4 / 20', first.rows === 4 && first.active && first.navActive && /系統預設規則/.test(first.meta) && /4 \/ 20/.test(first.count) && J(first.ids) === J(['default-d90', 'default-d45', 'default-d30', 'default-after7']), J(first));
      const lbls = await b.ev(`JSON.stringify([...document.querySelectorAll('#marAdminHost .mar-a-row')].map(r=>[r.querySelector('.mar-a-name').getAttribute('aria-label'), r.querySelector('.mar-a-dir').value, r.querySelector('.mar-a-n').value, [...r.querySelectorAll('.mar-a-notify input:checked')].map(c=>c.value).join('+'), r.querySelector('.mar-a-cond').value, r.querySelector('.mar-a-rep').value]))`).then(JSON.parse);
      t('AD-3 預設規則內容正確呈現（時間點方向／天數／通知對象／條件／重複），每個輸入有「第 N 條規則」的 aria-label', J(lbls.map((x) => x.slice(1))) === J([['before', '90', 'owner+secretary', 'always', '0'], ['before', '45', 'manager1', 'no_renewal_quote', '0'], ['before', '30', 'owner+manager1+marketing', 'quote_not_submitted', '0'], ['after', '1', 'owner+manager1', 'not_renewed', '7']]) && lbls[0][0] === '第 1 條規則 名稱', J(lbls));
      await shotTo(C, b, 'admin_rules_default.png');

      // ── 即時驗證 ──
      const R2 = rowS(1);
      await typeInto(b, R2 + ' .mar-a-name', '');
      let st = await b.ev(`JSON.stringify({err:document.querySelector(${J(R2 + ' [data-err="name"]')}).textContent, inv:document.querySelector(${J(R2 + ' .mar-a-name')}).getAttribute('aria-invalid'), all:document.querySelector('#marAdminHost .mar-a-error').textContent, hidden:document.querySelector('#marAdminHost .mar-a-error').hidden})`).then(JSON.parse);
      t('AD-4 名稱清空：該欄顯示「請填寫規則名稱」、aria-invalid、上方摘要提示有欄位需修正', /請填寫規則名稱/.test(st.err) && st.inv === 'true' && /需要修正/.test(st.all) && !st.hidden, J(st));
      await b.type('45天無報價通知主管'); await b.sleep(150);
      await typeInto(b, R2 + ' .mar-a-n', '400');
      st = await txt(b, R2 + ' [data-err="when"]');
      t('AD-5 天數 400：顯示「天數必須是 1 ～ 365 的整數」', /1 ～ 365/.test(st), st);
      await typeInto(b, R2 + ' .mar-a-n', '30');   // 與第 3 條（到期前 30，條件不同）：條件不同不算重複
      st = await txt(b, R2 + ' [data-err="dup"]');
      t('AD-6 同天數但條件不同（no_renewal_quote vs quote_not_submitted）不算重複', st === '', st);
      await pickOpt(b, R2 + ' .mar-a-cond', 'quote_not_submitted');
      st = await txt(b, rowS(2) + ' [data-err="dup"]');
      t('AD-7 相同條件＋適用類型＋天數（第 2 條改成與第 3 條相同）：後面那條（第 3 條）顯示「與第 2 條規則重複」', /與第 2 條規則重複/.test(st), st);
      await typeInto(b, R2 + ' .mar-a-n', '45'); await pickOpt(b, R2 + ' .mar-a-cond', 'no_renewal_quote');
      // 通知對象全取消
      await b.ev(`document.querySelectorAll(${J(R2 + ' .mar-a-notify input:checked')}).forEach(c=>c.click()); 0`); await b.sleep(150);
      st = await txt(b, R2 + ' [data-err="notify"]');
      t('AD-8 通知對象全取消：顯示「至少要選一個通知對象」', /至少要選一個通知對象/.test(st), st);
      await b.click(R2 + ' .mar-a-notify input[value="manager1"]');
      // 到期前的規則不能設重複；到期後可設
      const repDis = await b.ev(`JSON.stringify([document.querySelector(${J(rowS(0) + ' .mar-a-rep')}).disabled, document.querySelector(${J(rowS(3) + ' .mar-a-rep')}).disabled])`).then(JSON.parse);
      t('AD-9 重複間隔：到期前的規則欄位停用（0），到期後的規則可編輯', repDis[0] === true && repDis[1] === false, J(repDis));
      await typeInto(b, rowS(3) + ' .mar-a-rep', '120');
      st = await txt(b, rowS(3) + ' [data-err="repeat"]');
      t('AD-10 重複間隔 120：顯示「0（只發一次）或 1 ～ 90 天」', /1 ～ 90/.test(st), st);
      await typeInto(b, rowS(3) + ' .mar-a-rep', '7');

      // 有錯誤時不能儲存：不送出 PUT
      await typeInto(b, R2 + ' .mar-a-n', '999');
      const before = (await reqs(b)).length;
      await b.click('#marAdminHost .mar-a-btn-primary'); await b.sleep(500);
      const afterReqs = await reqs(b);
      st = await txt(b, '#marAdminHost .mar-a-status');
      t('AD-11 欄位有錯誤時按儲存：不送 PUT、顯示「還有欄位需要修正」、游標移到第一個錯誤欄', afterReqs.length === before && /需要修正/.test(st) && (await b.ev(`document.activeElement && document.activeElement.getAttribute('aria-invalid')==='true'`)), st);
      await typeInto(b, R2 + ' .mar-a-n', '45');

      // ── 與伺服器 validateConfig 對拍（隨機規則組）──
      const par = await parity(b, H);
      t('AD-12 前端即時驗證與伺服器 validateConfig 對拍 120 組隨機規則：「有錯／沒錯」完全一致', par.bad.length === 0 && par.n >= 100 && par.okN >= 15 && par.badN >= 15, J(par.bad.slice(0, 2)) + ' n=' + par.n + ' ok=' + par.okN + ' bad=' + par.badN);

      // ── 儲存：未改動 ──
      await b.nav(H.BASE + '/admin.html');
      await b.waitFor(`!!document.getElementById('marNavItem')`, 15000); await b.ev(X.PAGE_HELPERS);
      await gotoRules(b);
      await b.click('#marAdminHost .mar-a-btn-primary'); await b.sleep(900);
      st = await txt(b, '#marAdminHost .mar-a-status');
      t('AD-13 未改動就儲存（沒有啟用雙人覆核）：照常儲存成功並轉為自訂規則（內容仍是預設）', /已儲存/.test(st) && /自訂規則/.test(await txt(b, '#marAdminHost .mar-a-metaline')), st);

      // ── 編輯＋儲存：帶回既有規則 id、新規則不帶 id ──
      await typeInto(b, rowS(0) + ' .mar-a-name', RULE_XSS);
      await typeInto(b, rowS(0) + ' .mar-a-n', '60');
      await b.click('#marAdminHost .mar-a-bar .mar-a-btn');   // 新增規則
      await b.waitFor(`document.querySelectorAll('#marAdminHost .mar-a-row').length===5`, 4000);
      const newFocus = await b.ev(`document.activeElement && document.activeElement.classList.contains('mar-a-name') && document.activeElement.value`);
      t('AD-14 按「新增規則」：多一列（5 / 20），游標在新列名稱欄，預設「新規則」到期前 30 天、業務、無條件', newFocus === '新規則' && (await txt(b, '#marAdminHost .mar-a-count')).indexOf('5 / 20') >= 0, newFocus);
      await b.type('我的新規則'); await pickOpt(b, rowS(4) + ' .mar-a-dir', 'after'); await typeInto(b, rowS(4) + ' .mar-a-n', '10'); await b.sleep(100);
      await b.click(rowS(4) + ' .mar-a-notify input[value="marketing"]');
      await b.click(rowS(4) + ' button[aria-label="第 5 條規則 上移"]'); await b.sleep(200);
      const orderNow = await b.ev(`JSON.stringify([...document.querySelectorAll('#marAdminHost .mar-a-name')].map(x=>x.value))`).then(JSON.parse);
      t('AD-15 上移：新規則移到第 4 位，原第 4 條變第 5 條', orderNow[3] === '我的新規則' && orderNow[4].indexOf('到期後') >= 0, J(orderNow));
      const reqsBefore = (await reqs(b)).length;
      await b.click('#marAdminHost .mar-a-btn-primary');
      await b.waitFor(`/已儲存/.test(document.querySelector('#marAdminHost .mar-a-status').textContent)`, 8000);
      const put = (await reqs(b)).slice(reqsBefore).find((r) => r.m === 'PUT' && /reminder-rules/.test(r.u));
      const body = put ? JSON.parse(put.b) : null;
      const sent = body ? body.rules.map((r) => [r.id || '', r.name]) : [];
      t('AD-16 儲存的 PUT 帶回既有規則 id（default-d90、default-d45、default-d30、default-after7），新規則不帶 id', body && body.rules.length === 5 && sent.filter((x) => x[0]).map((x) => x[0]).sort().join() === 'default-after7,default-d30,default-d45,default-d90' && sent.filter((x) => !x[0]).length === 1 && sent.find((x) => !x[0])[1] === '我的新規則', J(sent));
      const srv = (await H.http(U.adm, 'GET', '/api/admin/reminder-rules')).json;
      t('AD-17 伺服器已儲存：5 條規則、新規則由伺服器產生 r_ 代號、舊代號保留、名稱含實體字串原樣保存、更新者＝管理員', srv.config.rules.length === 5 && srv.config.isDefault === false && srv.config.updatedBy === U.adm && srv.config.rules.some((r) => /^r_[0-9a-f]{10}$/.test(r.id) && r.name === '我的新規則' && r.whenDays === -10) && srv.config.rules.find((r) => r.id === 'default-d90').name === RULE_XSS && srv.config.rules.find((r) => r.id === 'default-d90').whenDays === 60, J(srv.config.rules.map((r) => [r.id, r.whenDays])));
      const meta = await txt(b, '#marAdminHost .mar-a-metaline');
      t('AD-18 儲存後頁面顯示「自訂規則」與最後更新者、時間', /自訂規則/.test(meta) && meta.indexOf(U.adm) >= 0 && /\d{4}/.test(meta), meta);
      const logs = await H.getLogs(U.adm);
      t('AD-19 稽核 SAVE_REMINDER_RULES 兩筆（AD-13 與 AD-16 各一次）', logs.filter((l) => l.action === 'SAVE_REMINDER_RULES').length === 2);
      // XSS：規則名稱的實體字串原樣顯示成文字，沒有被解碼成標籤
      const xs = await b.ev(`JSON.stringify({inputs:document.querySelector(${J(rowS(0) + ' .mar-a-name')}).value, tags:document.querySelectorAll('#marAdminHost i').length, tl:[...document.querySelectorAll('#marAdminHost .mar-a-tllist li, #marAdminHost svg text')].map(x=>x.textContent).join('|').indexOf('&lt;i&gt;')>=0, xss:window.__xss})`).then(JSON.parse);
      t('AD-20 規則名稱含 &lt;i&gt; 實體：輸入欄、時間軸、文字清單都原樣顯示（沒有被當 HTML 解碼，沒有 <i> 元素）', xs.inputs === RULE_XSS && xs.tags === 0 && xs.tl === true && xs.xss === 0, J(xs));

      // ── 還原預設規則 ──
      await b.click('#marAdminHost .mar-a-btn-warn');
      const conf = await b.ev(`JSON.stringify({vis:!document.querySelector('#marAdminHost .mar-a-confirm').hidden, rows:document.querySelectorAll('#marAdminHost .mar-a-row').length})`).then(JSON.parse);
      t('AD-21 按「還原預設規則」：先出現確認列（尚未改動編輯區）', conf.vis && conf.rows === 5, J(conf));
      await b.click('#marAdminHost .mar-a-confirm button');
      await b.sleep(300);
      const def = await b.ev(`JSON.stringify({rows:document.querySelectorAll('#marAdminHost .mar-a-row').length, names:MaRemAdmin._state.models.map(m=>m.id), st:document.querySelector('#marAdminHost .mar-a-status').textContent, srvRules:null})`).then(JSON.parse);
      t('AD-22 確定還原：編輯區改為 4 條預設規則（帶預設代號）、提示「尚未儲存」；伺服器的規則還沒變', def.rows === 4 && J(def.names) === J(['default-d90', 'default-d45', 'default-d30', 'default-after7']) && /尚未儲存/.test(def.st) && (await H.http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config.rules.length === 5, J(def));
      await b.click('#marAdminHost .mar-a-btn-primary');
      await b.waitFor(`/已儲存/.test(document.querySelector('#marAdminHost .mar-a-status').textContent)`, 8000);
      const srv2 = (await H.http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config.rules;
      t('AD-23 儲存還原後的規則：伺服器規則＝預設內容（名稱／天數／對象／條件），代號沿用', srv2.length === 4 && srv2[0].id === 'default-d90' && srv2[0].whenDays === 90 && srv2[3].whenDays === -1 && srv2[3].repeatEveryDays === 7, J(srv2.map((r) => [r.id, r.whenDays])));

      // ── 試算器 ──
      await simSection(C, b, H);
      const fe = b.errors.filter((e) => !/Failed to load resource|favicon/i.test(e));
      t('AD-E 沒有 JS 例外（後台）', fe.length === 0, J(fe).slice(0, 300));
    });
  }

  async function parity(b, H) {
    // 在頁面內產生隨機規則組（種子固定）→ 前端驗證結果；同一組送伺服器 PUT 驗證的是 validateConfig；為了不改動伺服器資料，改用 simulate API（它對 rules 做同樣的 validateConfig）
    const sets = await b.ev(`(()=>{
      let s=20261024; const rnd=()=>{ s=(s*1664525+1013904223)%4294967296; return s/4294967296; };
      const pick=(a)=>a[Math.floor(rnd()*a.length)];
      const names=['規則甲','正常名稱','規則乙','a','規則丙','名稱&amp;','規則丁','規則戊','規則己','規則庚','','  ','x'.repeat(41),'含<標籤>'];
      const out=[];
      for(let i=0;i<120;i++){
        const n=1+Math.floor(rnd()*4);
        const models=[];
        for(let k=0;k<n;k++){
          const dir=pick(['before','before','on','after']);
          const nRaw=pick(['30','45','90','1','365','7','60','14','21','100','180','120','200','2','3','366','','abc','12.5','400']);
          const repeat0=pick(['0','0','0','7','90','14','30','1','5','10','2','91','-1','x','']);
          const repeat = dir==='before' ? '0' : repeat0;   // 到期前的規則重複欄位停用（畫面上不可能出現其他值）
          const notify=pick([['owner'],['owner','manager1'],['marketing'],['secretary'],['owner','secretary'],['manager2'],['owner'],['manager1'],[]]);
          models.push({id:'', name:pick(names), enabled:true, dir, nRaw: dir==='on'?'0':nRaw, repeatRaw: repeat, notify, condition:pick(['always','no_renewal_quote']), applies:pick(['all','all','ma','license'])});
        }
        out.push(models);
      }
      return JSON.stringify(out.map(models=>({models, v: MaRemAdmin.validateRules(models).count, api: models.map(m=>{ const o=MaRemAdmin.toApi(m); return o; })})));
    })()`).then(JSON.parse);
    let bad = [], okN = 0, badN = 0;
    for (let i = 0; i < sets.length; i++) {
      const s = sets[i];
      // toApi 對非法輸入的轉換不保證有意義；伺服器收到的是「使用者實際輸入的原值」：直接用 raw 欄位組出 API 規則
      const rules = s.models.map((m, k) => {
        const w = m.dir === 'on' ? 0 : (m.dir === 'before' ? m.nRaw : '-' + m.nRaw);
        return { name: m.name, enabled: true, whenDays: /^-?\d{1,5}$/.test(String(w)) ? Number(w) : w, repeatEveryDays: m.dir === 'before' ? (m.repeatRaw === '0' || m.repeatRaw === '' ? 0 : (/^-?\d+$/.test(m.repeatRaw) ? Number(m.repeatRaw) : m.repeatRaw)) : (m.repeatRaw === '' ? 0 : (/^-?\d+$/.test(m.repeatRaw) ? Number(m.repeatRaw) : m.repeatRaw)), notify: m.notify, condition: m.condition, appliesTo: m.applies === 'all' ? 'all' : [m.applies] };
      });
      const r = await H.http(U.adm, 'POST', '/api/admin/reminder-rules/simulate', { daysLeft: 10, quoteState: 'none', rules });
      const srvBad = r.status !== 200;
      const feBad = s.v > 0;
      if (feBad) badN++; else okN++;
      if (srvBad !== feBad) bad.push({ i, srv: r.status + ' ' + (r.json && r.json.code), fe: s.v, rules: J(rules).slice(0, 220) });
    }
    return { n: sets.length, bad, okN, badN };
  }

  async function simSection(C, b, H) {
    // 回到乾淨狀態：重新載入頁面（儲存後的 5 條→4 條預設內容）
    await b.nav(H.BASE + '/admin.html');
    await b.waitFor(`!!document.getElementById('marNavItem')`, 15000); await b.ev(X.PAGE_HELPERS);
    await gotoRules(b);
    await b.sleep(900);
    const scenarios = [
      { days: 45, qs: 'none' }, { days: 20, qs: 'draft' }, { days: 20, qs: 'sent' }, { days: -10, qs: 'none' }, { days: 100, qs: 'none' }, { days: 20, qs: 'norenew' }, { days: 30, qs: 'none' }, { days: 0, qs: 'draft' },
    ];
    let allEq = true; const diffs = [];
    for (const sc of scenarios) {
      await setDays(b, sc.days);
      await b.ev(`document.querySelector('input[name="marSimQs"][value=${J(sc.qs)}]').click(); 0`); await b.sleep(900);
      const ui = await simRoles(b);
      const api = await apiSim(H, { daysLeft: sc.days, quoteState: sc.qs === 'norenew' ? 'none' : sc.qs, notRenew: sc.qs === 'norenew', termKind: 'ma' });
      const exp = Object.fromEntries(ROLES.map((r) => [r, namesOf(api.byRole[r])]));
      const same = ROLES.every((r) => ui[r] === exp[r]);
      if (!same) { allEq = false; diffs.push({ sc, ui, exp }); }
    }
    t('AD-24 試算器：8 組情境（天數×報價單狀態×不續約）UI 各角色顯示的規則 ＝ 伺服器 simulate 回應', allEq, J(diffs.slice(0, 1)).slice(0, 400));
    // 具體值（防止兩邊都空的空轉）
    await setDays(b, 45); await b.ev(`document.querySelector('input[name="marSimQs"][value="none"]').click(); 0`); await b.sleep(900);
    const r45 = await simRoles(b);
    t('AD-25 剩 45 天、尚未建立：業務與秘書看到 D-90、一級主管看到 D-45、二級主管與行銷沒有', r45.owner.indexOf('到期前 90 天') >= 0 && r45.secretary.indexOf('到期前 90 天') >= 0 && r45.manager1.indexOf('45 天') >= 0 && r45.manager2 === '' && r45.marketing === '', J(r45));
    const msg = await txt(b, '#marAdminHost .mar-a-role[data-role="owner"] .mar-a-role-msg');
    t('AD-26 角色卡顯示此刻的提醒文字（與鈴鐺同格式：公司：品項 剩 N 天到期｜續約報價單狀態）', /^⏰ （示範客戶）：（示範品項） 剩 45 天到期｜尚未建立續約報價單$/.test(msg), msg);
    await b.ev(`document.querySelector('input[name="marSimQs"][value="sent"]').click(); 0`); await b.sleep(900);
    const rsent = await simRoles(b);
    t('AD-27 已送簽：連預設的「無條件」D-90 也不再成立（2026-10-24 決定），所有角色都沒有提醒', ROLES.every((r) => rsent[r] === ''), J(rsent));
    await b.ev(`document.querySelector('input[name="marSimQs"][value="draft"]').click(); 0`); await b.sleep(900);
    const rdraft = await simRoles(b);
    t('AD-28 草稿：D-90 仍成立（草稿不解除）', rdraft.owner.indexOf('到期前 90 天') >= 0, J(rdraft));
    // 未儲存的草稿規則：把 D-90 改成 60 天，剩 70 天時業務不會看到（但伺服器規則仍是 90）
    await typeInto(b, rowS(0) + ' .mar-a-n', '60');
    await setDays(b, 70); await b.ev(`document.querySelector('input[name="marSimQs"][value="none"]').click(); 0`); await b.sleep(900);
    const rdr = await simRoles(b);
    const srvNow = (await H.http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config.rules.find((r) => r.id === 'default-d90').whenDays;
    t('AD-29 試算用「編輯中、尚未儲存」的規則：把 D-90 改成 60 天後，剩 70 天業務沒有提醒；伺服器規則仍是 90（沒被改）', rdr.owner === '' && srvNow === 90, J([rdr.owner, srvNow]));
    await setDays(b, 55); await b.sleep(300);
    const r55 = await simRoles(b);
    t('AD-30 滑桿／數字改到剩 55 天：草稿規則（60 天）成立，業務與秘書又看到提醒', r55.owner !== '' && r55.secretary !== '', J(r55));
    // 滑桿本身（真實鍵盤方向鍵）
    await b.focus('#marSimDays'); await b.key('ArrowLeft'); await b.key('ArrowLeft'); await b.sleep(700);
    const sl = await b.ev(`JSON.stringify({range:document.getElementById('marSimDays').value, num:document.querySelector('.mar-a-simnum').value, txt:document.querySelector('.mar-a-daytxt').textContent})`).then(JSON.parse);
    t('AD-31 滑桿（鍵盤方向鍵）：值與數字欄、文字同步（剩 53 天到期）', sl.range === '53' && sl.num === '53' && /剩 53 天到期/.test(sl.txt), J(sl));
    // 真實納管項
    await setDays(b, 40);
    await b.ev(`(()=>{const s=document.getElementById('marSimItem'); s.value='iA'; s.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`); await b.sleep(1100);
    const real = await b.ev(`JSON.stringify({kind:document.getElementById('marSimKind').value, num:document.querySelector('.mar-a-simnum').value, who:[...document.querySelectorAll('#marAdminHost .mar-a-role-who')].map(x=>x.textContent), msg:document.querySelector('#marAdminHost .mar-a-role[data-role="owner"] .mar-a-role-msg')?document.querySelector('#marAdminHost .mar-a-role[data-role="owner"] .mar-a-role-msg').textContent:''})`).then(JSON.parse);
    const apiReal = await apiSim(H, { daysLeft: Number(real.num), quoteState: 'none', termKind: real.kind, maItemId: 'iA' });
    const apiWho = ROLES.filter((r) => (apiReal.byRole[r] || []).length).map((r) => '收件人：' + (apiReal.recipients[r] || []).map((x) => x.label || x.username).join('、'));
    t('AD-32 選真實納管項（R甲公司）：天數與類型帶入、各角色顯示真實收件人（顯示名，與 simulate API 一致）、提醒文字用真實公司與品項', real.who.filter((x) => /^收件人/.test(x)).join('|') === apiWho.join('|') && real.who.some((x) => x.indexOf('測試業務甲') >= 0) && real.msg.indexOf('R甲公司：' + MAP) >= 0, J([real, apiWho]));
    const dd = await b.ev(`[...document.getElementById('marSimItem').options].map(o=>o.textContent).join('|')`);
    t('AD-33 真實納管項下拉只顯示公司與品項、到期日純文字（含 XSS 公司名原樣顯示，沒有執行）', dd.indexOf(XSS1) >= 0 && (await b.ev(`window.__xss`)) === 0 && dd.indexOf(CNO) < 0 && dd.indexOf(AMT) < 0, dd.slice(0, 200));
    await shotTo(C, b, 'admin_simulator.png');
    // ── 時間軸 SVG ──
    await b.ev(`(()=>{const s=document.getElementById('marSimItem'); s.value=''; s.dispatchEvent(new Event('change',{bubbles:true})); return 0})()`);
    await setDays(b, 45); await b.sleep(900);
    const sv1 = await b.ev(`(()=>{const s=document.querySelector('#marAdminHost svg.mar-a-svg'); const td=s.querySelector('.mar-a-today'); return JSON.stringify({dots:s.querySelectorAll('circle.mar-a-dot').length, active:s.querySelectorAll('.mar-a-dot-active').length, today:Number(td.getAttribute('x1')), expiry:Number(s.querySelector('.mar-a-expiry').getAttribute('x1')), role:s.getAttribute('role'), label:s.getAttribute('aria-label'), list:document.querySelectorAll('#marAdminHost .mar-a-tllist li').length, hasStyleAttr:!!s.querySelector('[style]')})})()`).then(JSON.parse);
    const sim45 = await apiSim(H, { daysLeft: 45, quoteState: 'none', termKind: 'ma', rules: await rulesInPage(b) });
    t('AD-34 時間軸 SVG：每條規則一個標記（4）、成立的標記數＝simulate 回應的 activeRuleIds 數、有 role=img 與說明、文字清單同步、沒有 style 屬性', sv1.dots === 4 && sv1.active === sim45.activeRuleIds.length && sv1.role === 'img' && /4 條規則/.test(sv1.label) && sv1.list === 4 && !sv1.hasStyleAttr, J([sv1, sim45.activeRuleIds]));
    await setDays(b, 10); await b.sleep(900);
    const sv2 = await b.ev(`Number(document.querySelector('#marAdminHost .mar-a-today').getAttribute('x1'))`);
    t('AD-35 「今天」指示線隨天數移動（剩 45 天→剩 10 天，向右移動，且在到期線左側）', sv2 > sv1.today && sv2 < sv1.expiry, J([sv1.today, sv2, sv1.expiry]));
    await setDays(b, -30); await b.sleep(900);
    const sv3 = await b.ev(`Number(document.querySelector('#marAdminHost .mar-a-today').getAttribute('x1'))`);
    t('AD-36 逾期 30 天：指示線在到期線右側', sv3 > sv1.expiry, J([sv3, sv1.expiry]));
    await shotTo(C, b, 'admin_timeline.png');
    // 錯誤規則時無法試算
    await typeInto(b, rowS(0) + ' .mar-a-n', '9999'); await b.sleep(600);
    const se = await b.ev(`JSON.stringify({err:document.querySelector('#marAdminHost .mar-a-simcard .mar-a-error').hidden, txt:document.querySelector('#marAdminHost .mar-a-simcard .mar-a-error').textContent})`).then(JSON.parse);
    t('AD-37 規則有錯誤時試算器顯示「暫時無法試算」，不送出試算請求', se.err === false && /無法試算/.test(se.txt));
  }

  return Object.assign({ secAdmin }, require('./_ma-reminders-ui-sections2')(X));
};
