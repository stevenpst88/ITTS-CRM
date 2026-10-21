/* 企業主檔「責任業務」＋「待補名片」待辦（使用者端）
 *
 * 後端：GET /api/my-tasks/company-cards、POST …/dismiss、GET /api/marketing/responsibility-feedback、POST …/:id/resolve、
 *       GET /api/manager/company-card-overdue、GET /api/company-responsible/candidates、POST /api/companies/:id/set-responsible-sales、
 *       POST /api/companies/responsible-sales-batch。（COMPANY_CARD_TASKS=off 時後端回 enabled:false → 這裡什麼都不顯示）
 * 待辦是「推導狀態」：業務補上真實名片後，下次載入（或關閉新增名片視窗後）自動消失，不需要任何手動動作。
 *
 * 畫面：
 *   - 剛登入的第一次載入：彙總視窗（alertdialog，列前 10 家；前往新增名片／解除…／稍後處理；Esc＝稍後處理）
 *   - 首頁常駐「待辦事項」卡（業務首頁、主管首頁、行銷首頁各有掛載點 [data-cct-host]）
 *   - 行銷首頁「業務回報待修」、主管首頁「逾期未補名片」與「業務已回報」
 * 安全：全部動態內容一律 createElement＋textContent（不經 innerHTML）、不使用 inline style／on* 屬性；樣式在 style.css 的 .cct-*。
 * 對外：window.ITTSCardTasks = { start, refresh, openSummary, openFromNotification, openCompanyPicker, openBatchPicker, canAssign }
 */
(function () {
  'use strict';
  if (window.ITTSCardTasks) return;

  var DEFAULT_REASONS = [
    { key: 'not_mine', label: '不是我負責' },
    { key: 'inactive', label: '客戶已不往來' },
    { key: 'no_contact', label: '還沒有聯絡人' },
    { key: 'other', label: '其他' },
  ];
  var WRITE_ROLES = ['admin', 'marketing', 'manager1', 'manager2', 'executive'];
  var SUMMARY_MAX = 10;
  var PANEL_MAX = 5;

  var state = { user: null, tasks: null, feedback: null, overdue: null, loginShown: false, seq: 0, summary: null, expanded: false };
  var stack = [];
  var uidN = 0;

  function uid(p) { uidN++; return 'cct-' + p + '-' + uidN; }
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = String(text);
    return e;
  }
  function button(label, cls, onClick) {
    var b = el('button', cls, label);
    b.type = 'button';
    if (onClick) b.addEventListener('click', onClick);
    return b;
  }
  function toast(msg) { try { if (typeof window.showToast === 'function') window.showToast(msg); } catch (e) { /* ignore */ } }
  function jget(url) {
    return fetch(url, { credentials: 'same-origin', cache: 'no-store' }).then(function (r) { return r.ok ? r.json() : null; }).catch(function () { return null; });
  }
  function jpost(url, body) {
    return fetch(url, { method: 'POST', credentials: 'same-origin', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) })
      .then(function (r) { return r.json().catch(function () { return {}; }).then(function (d) { return { ok: r.ok, status: r.status, data: d }; }); })
      .catch(function () { return { ok: false, status: 0, data: { error: '網路錯誤，請稍後再試' } }; });
  }
  function canAssign(role) { return WRITE_ROLES.indexOf(role || (state.user && state.user.role)) >= 0; }
  function daysText(t) { return t.days > 0 ? ('已指定 ' + t.days + ' 天') : '今天指定'; }

  /* ───────── 對話框（堆疊；焦點圈住；Esc 只處理最上層）───────── */
  function onKey(e) {
    if (!stack.length) return;
    var top = stack[stack.length - 1];
    if (e.key === 'Escape') {
      e.preventDefault(); e.stopPropagation();
      if (top.onEsc) top.onEsc(); else top.close();
      return;
    }
    if (e.key === 'Tab') {
      var f = top.dlg.querySelectorAll('button:not([disabled]),input:not([disabled]),textarea:not([disabled]),select:not([disabled]),[href],[tabindex]:not([tabindex="-1"])');
      if (!f.length) return;
      var first = f[0], last = f[f.length - 1];
      if (e.shiftKey && (document.activeElement === first || !top.dlg.contains(document.activeElement))) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && (document.activeElement === last || !top.dlg.contains(document.activeElement))) { e.preventDefault(); first.focus(); }
    }
  }
  var keyBound = false;
  function openDialog(o) {
    if (!keyBound) { document.addEventListener('keydown', onKey, true); keyBound = true; }
    var prev = document.activeElement;
    var ov = el('div', 'cct-overlay');
    var dlg = el('div', 'cct-dialog' + (o.wide ? ' cct-wide' : ''));
    dlg.setAttribute('role', o.role || 'dialog');
    dlg.setAttribute('aria-modal', 'true');
    var h = el('h2', 'cct-dlg-title', o.title);
    h.id = uid('t');
    dlg.setAttribute('aria-labelledby', h.id);
    var body = el('div', 'cct-dlg-body');
    body.id = uid('d');
    dlg.setAttribute('aria-describedby', body.id);
    var foot = el('div', 'cct-dlg-foot');
    dlg.appendChild(h); dlg.appendChild(body); dlg.appendChild(foot);
    ov.appendChild(dlg);
    var entry = {
      ov: ov, dlg: dlg, body: body, foot: foot, title: h, onEsc: o.onEsc || null,
      close: function () {
        var i = stack.indexOf(entry);
        if (i >= 0) stack.splice(i, 1);
        if (ov.parentNode) ov.parentNode.removeChild(ov);
        if (o.onClose) { try { o.onClose(); } catch (e) { /* ignore */ } }
        if (prev && prev.focus && document.contains(prev)) { try { prev.focus(); } catch (e) { /* ignore */ } }
      },
    };
    stack.push(entry);
    document.body.appendChild(ov);
    return entry;
  }
  function focusFirst(entry, sel) {
    var t = (sel && entry.dlg.querySelector(sel)) || entry.dlg.querySelector('button:not([disabled]),input:not([disabled]),textarea,select');
    if (t) { try { t.focus(); } catch (e) { /* ignore */ } }
  }

  /* ───────── 前往新增名片（帶入公司資料）───────── */
  function gotoAddCard(t) {
    if (typeof window.openModalAddContact === 'function') {
      window.openModalAddContact({ company: t.name, taxId: t.taxId || '', industry: t.industry || '', address: t.address || '', website: t.website || '', isForeign: !t.taxId });
    } else {
      toast('請到「聯絡人」頁新增這家公司的名片');
    }
  }

  /* ───────── 解除對話框 ───────── */
  function openDismiss(t, onDone) {
    var reasons = (state.tasks && state.tasks.reasons && state.tasks.reasons.length) ? state.tasks.reasons : DEFAULT_REASONS;
    var d = openDialog({ title: '解除「' + t.name + '」的待補名片待辦', role: 'dialog' });
    var p = el('p', 'cct-lead', '請選擇原因。解除後這項待辦會對你隱藏，並回報給行銷與你的主管；行銷可能會重新指定責任業務。');
    d.body.appendChild(p);
    var fs = el('fieldset', 'cct-fieldset');
    fs.appendChild(el('legend', 'cct-legend', '解除原因（必選）'));
    var name = uid('reason');
    reasons.forEach(function (r) {
      var lab = el('label', 'cct-radio');
      var inp = document.createElement('input');
      inp.type = 'radio'; inp.name = name; inp.value = r.key;
      lab.appendChild(inp);
      lab.appendChild(el('span', null, r.label));
      fs.appendChild(lab);
    });
    d.body.appendChild(fs);
    var nl = el('label', 'cct-field-label', '說明（選「其他」時必填）');
    var ta = document.createElement('textarea');
    ta.className = 'cct-textarea'; ta.rows = 3; ta.maxLength = 300;
    ta.id = uid('note'); nl.setAttribute('for', ta.id);
    d.body.appendChild(nl); d.body.appendChild(ta);
    var err = el('div', 'cct-error'); err.setAttribute('role', 'alert'); err.hidden = true;
    d.body.appendChild(err);
    var okBtn = button('確定解除', 'cct-btn cct-btn-danger');
    var cancel = button('取消', 'cct-btn cct-btn-ghost', function () { d.close(); });
    d.foot.appendChild(cancel); d.foot.appendChild(okBtn);
    function showErr(m) { err.textContent = m; err.hidden = false; }
    okBtn.addEventListener('click', function () {
      var sel = d.dlg.querySelector('input[type=radio]:checked');
      if (!sel) { showErr('請先選擇解除原因'); return; }
      var note = ta.value.trim();
      if (sel.value === 'other' && !note) { showErr('選「其他」時必須填寫說明'); ta.focus(); return; }
      okBtn.disabled = true;
      jpost('/api/my-tasks/company-cards/dismiss', { companyId: t.companyId, reason: sel.value, note: note }).then(function (r) {
        if (!r.ok) { okBtn.disabled = false; showErr((r.data && r.data.error) || '解除失敗'); return; }
        d.close();
        toast('已解除，並回報給行銷與主管');
        refresh().then(function () { if (onDone) onDone(); });
      });
    });
    focusFirst(d, 'input[type=radio]');
    return d;
  }

  /* ───────── 待辦列（彙總視窗與首頁卡共用）───────── */
  function taskRow(t, afterAction) {
    var li = el('li', 'cct-row');
    var main = el('div', 'cct-row-main');
    main.appendChild(el('strong', 'cct-name', t.name));
    var meta = el('span', 'cct-meta', (t.taxId ? '統編 ' + t.taxId + '　' : '無統編　') + daysText(t));
    main.appendChild(meta);
    if (t.overdue) main.appendChild(el('span', 'cct-chip cct-chip-warn', '已逾 ' + ((state.tasks && state.tasks.overdueDays) || 14) + ' 天'));
    li.appendChild(main);
    var act = el('div', 'cct-row-actions');
    var add = button('前往新增名片', 'cct-btn cct-btn-primary cct-btn-sm', function () { if (afterAction) afterAction('add'); gotoAddCard(t); });
    add.setAttribute('aria-label', '前往新增「' + t.name + '」的名片');
    var dis = button('解除…', 'cct-btn cct-btn-ghost cct-btn-sm', function () { openDismiss(t); });
    dis.setAttribute('aria-label', '解除「' + t.name + '」的待補名片待辦');
    act.appendChild(add); act.appendChild(dis);
    li.appendChild(act);
    return li;
  }

  /* ───────── 彙總視窗 ───────── */
  function fillSummary(d) {
    var tasks = (state.tasks && state.tasks.items) || [];
    d.body.textContent = '';
    d.body.appendChild(el('p', 'cct-lead', '以下公司已指定你為責任業務，但你名下還沒有這些公司的名片。請補上名片；補上後這項待辦會自動消失。如果不是你負責，可以按「解除…」並選原因，系統會回報給行銷與主管。'));
    var ul = el('ul', 'cct-list');
    tasks.slice(0, SUMMARY_MAX).forEach(function (t) { ul.appendChild(taskRow(t, function (what) { if (what === 'add') d.close(); })); });
    d.body.appendChild(ul);
    if (tasks.length > SUMMARY_MAX) d.body.appendChild(el('p', 'cct-more', '還有 ' + (tasks.length - SUMMARY_MAX) + ' 家未列出，可在首頁「待辦事項」查看全部。'));
    d.title.textContent = '你有 ' + tasks.length + ' 家公司待補名片';
  }
  function openSummary(opts) {
    if (state.summary) return Promise.resolve();
    var go = (opts && opts.fetch) ? refresh() : Promise.resolve();
    return go.then(function () {
      var tasks = (state.tasks && state.tasks.items) || [];
      if (!tasks.length) { if (opts && opts.fetch) toast('目前沒有待補名片的待辦'); return; }
      if (state.summary) return;
      var d = openDialog({ title: '你有 ' + tasks.length + ' 家公司待補名片', role: 'alertdialog', wide: true, onClose: function () { state.summary = null; } });
      state.summary = d;
      d.onEsc = function () { d.close(); };   // Esc＝稍後處理
      fillSummary(d);
      d.foot.appendChild(button('稍後處理', 'cct-btn cct-btn-ghost cct-later', function () { d.close(); }));
      d.rerender = function () {
        var left = (state.tasks && state.tasks.items) || [];
        if (!left.length) { d.close(); return; }
        fillSummary(d);
      };
      focusFirst(d, '.cct-btn-primary');
    });
  }

  /* ───────── 首頁卡片 ───────── */
  function panel(titleText, count, cls) {
    var sec = el('section', 'cct-panel ' + (cls || ''));
    var hid = uid('h');
    sec.setAttribute('aria-labelledby', hid);
    var hd = el('div', 'cct-panel-hd');
    var h = el('h3', 'cct-panel-title', titleText);
    h.id = hid;
    hd.appendChild(h);
    if (count != null) hd.appendChild(el('span', 'cct-count', count));
    sec.appendChild(hd);
    return sec;
  }
  function myTasksPanel() {
    var tasks = (state.tasks && state.tasks.items) || [];
    if (!tasks.length) return null;
    var sec = panel('📋 待辦事項：請補名片', tasks.length, 'cct-panel-tasks');
    sec.appendChild(el('p', 'cct-sub', '你是這些公司的責任業務，但名下還沒有名片。補上名片後這項待辦會自動消失。'));
    var ul = el('ul', 'cct-list');
    var show = state.expanded ? tasks : tasks.slice(0, PANEL_MAX);
    show.forEach(function (t) { ul.appendChild(taskRow(t, null)); });
    sec.appendChild(ul);
    if (tasks.length > PANEL_MAX) {
      sec.appendChild(button(state.expanded ? '收合' : '顯示全部 ' + tasks.length + ' 家', 'cct-btn cct-btn-link', function () { state.expanded = !state.expanded; renderHosts(); }));
    }
    return sec;
  }
  function feedbackLine(f) {
    var li = el('li', 'cct-row');
    var main = el('div', 'cct-row-main');
    main.appendChild(el('strong', 'cct-name', f.companyName || ''));
    var bu = (f.bu && f.bu.length) ? '（' + f.bu.join('/') + '）' : '';
    main.appendChild(el('span', 'cct-meta', '業務 ' + (f.userLabel || f.username) + bu + ' 回報：' + (f.reasonLabel || '')));
    if (f.note) main.appendChild(el('span', 'cct-note', '說明：' + f.note));
    return { li: li, main: main };
  }
  function feedbackPanel() {
    var items = (state.feedback && state.feedback.items) || [];
    if (!items.length) return null;
    var sec = panel('🛠️ 業務回報待修', items.length, 'cct-panel-feedback');
    sec.appendChild(el('p', 'cct-sub', '業務手動解除了「請補名片」待辦。請重新指定責任業務，或標示已處理。'));
    var ul = el('ul', 'cct-list');
    items.forEach(function (f) {
      var r = feedbackLine(f);
      var act = el('div', 'cct-row-actions');
      act.appendChild(button('重新指定…', 'cct-btn cct-btn-primary cct-btn-sm', function () { openReassign(f); }));
      act.appendChild(button('標示已處理', 'cct-btn cct-btn-ghost cct-btn-sm', function () { openHandled(f); }));
      r.li.appendChild(r.main); r.li.appendChild(act);
      ul.appendChild(r.li);
    });
    sec.appendChild(ul);
    return sec;
  }
  function overduePanels() {
    var o = state.overdue;
    if (!o || !o.enabled) return [];
    var out = [];
    if (o.overdue && o.overdue.length) {
      var sec = panel('⏰ 責任業務逾期未補名片（超過 ' + (o.overdueDays || 14) + ' 天）', o.overdue.length, 'cct-panel-overdue');
      var ul = el('ul', 'cct-list');
      o.overdue.slice(0, 20).forEach(function (t) {
        var li = el('li', 'cct-row');
        var main = el('div', 'cct-row-main');
        main.appendChild(el('strong', 'cct-name', t.name));
        var bu = (t.bu && t.bu.length) ? '（' + t.bu.join('/') + '）' : '';
        main.appendChild(el('span', 'cct-meta', '責任業務 ' + (t.userLabel || t.username) + bu + '　已指定 ' + t.days + ' 天'));
        li.appendChild(main);
        ul.appendChild(li);
      });
      sec.appendChild(ul);
      if (o.overdue.length > 20) sec.appendChild(el('p', 'cct-more', '另有 ' + (o.overdue.length - 20) + ' 家未列出。'));
      out.push(sec);
    }
    if (o.dismissed && o.dismissed.length) {
      var s2 = panel('🛠️ 業務已回報（等行銷處理）', o.dismissed.length, 'cct-panel-feedback');
      var u2 = el('ul', 'cct-list');
      o.dismissed.slice(0, 20).forEach(function (f) { var r = feedbackLine(f); r.li.appendChild(r.main); u2.appendChild(r.li); });
      s2.appendChild(u2);
      out.push(s2);
    }
    return out;
  }
  function renderHosts() {
    var hosts = document.querySelectorAll('[data-cct-host]');
    for (var i = 0; i < hosts.length; i++) {
      var h = hosts[i];
      var kind = h.getAttribute('data-cct-host');
      h.textContent = '';
      var parts = [];
      if (kind === 'mkt') { parts.push(feedbackPanel()); parts.push(myTasksPanel()); }
      else if (kind === 'mgr') { overduePanels().forEach(function (p) { parts.push(p); }); parts.push(myTasksPanel()); }
      else { parts.push(myTasksPanel()); if (state.user && state.user.role === 'admin') { parts.push(feedbackPanel()); } }
      parts.forEach(function (p) { if (p) h.appendChild(p); });
      h.hidden = !h.firstChild;
    }
    if (state.summary && state.summary.rerender) state.summary.rerender();
  }

  /* ───────── 選人（責任業務）───────── */
  function candidates() {
    return jget('/api/company-responsible/candidates').then(function (d) { return (d && d.users) || []; });
  }
  // o: { title, intro, current:[…], withMode:boolean, submitLabel, onSubmit(list, mode) → Promise<{ok,data}> , onSaved(data) }
  function openPicker(o) {
    var d = openDialog({ title: o.title, role: 'dialog', wide: true });
    d.body.appendChild(el('p', 'cct-lead', o.intro || ''));
    var box = el('fieldset', 'cct-fieldset cct-picker');
    box.appendChild(el('legend', 'cct-legend', '責任業務（可多選；顯示名稱＋BU）'));
    d.body.appendChild(box);
    var loading = el('p', 'cct-sub', '載入中…');
    box.appendChild(loading);
    var modeName = uid('mode');
    if (o.withMode) {
      var mf = el('fieldset', 'cct-fieldset');
      mf.appendChild(el('legend', 'cct-legend', '套用方式'));
      [['add', '加入（保留這些公司原有的責任業務）'], ['replace', '取代（以這次選的人為準）']].forEach(function (m, i) {
        var lab = el('label', 'cct-radio');
        var inp = document.createElement('input'); inp.type = 'radio'; inp.name = modeName; inp.value = m[0]; if (i === 0) inp.checked = true;
        lab.appendChild(inp); lab.appendChild(el('span', null, m[1])); mf.appendChild(lab);
      });
      d.body.appendChild(mf);
    }
    var err = el('div', 'cct-error'); err.setAttribute('role', 'alert'); err.hidden = true;
    d.body.appendChild(err);
    var save = button(o.submitLabel || '儲存', 'cct-btn cct-btn-primary');
    save.disabled = true;
    d.foot.appendChild(button('取消', 'cct-btn cct-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(save);
    var cur = new Set(o.current || []);
    candidates().then(function (users) {
      loading.remove();
      if (!users.length) { box.appendChild(el('p', 'cct-sub', '沒有可選的業務（只有可擁有名片的啟用帳號才能被指定）。')); return; }
      users.forEach(function (u) {
        var lab = el('label', 'cct-check');
        var inp = document.createElement('input'); inp.type = 'checkbox'; inp.value = u.username; inp.checked = cur.has(u.username);
        lab.appendChild(inp);
        lab.appendChild(el('span', 'cct-check-name', u.label));
        lab.appendChild(el('span', 'cct-chip', (u.bu && u.bu.length) ? u.bu.join('/') : '無 BU'));
        box.appendChild(lab);
      });
      save.disabled = false;
      focusFirst(d, 'input[type=checkbox]');
    });
    save.addEventListener('click', function () {
      var list = [].map.call(box.querySelectorAll('input[type=checkbox]:checked'), function (i) { return i.value; });
      var mi = d.dlg.querySelector('input[name="' + modeName + '"]:checked');
      save.disabled = true;
      o.onSubmit(list, mi ? mi.value : 'replace').then(function (r) {
        if (!r.ok) { save.disabled = false; err.textContent = (r.data && r.data.error) || '儲存失敗'; err.hidden = false; return; }
        d.close();
        if (o.onSaved) o.onSaved(r.data);
      });
    });
    return d;
  }
  function openCompanyPicker(company, onSaved) {
    return openPicker({
      title: '設定「' + company.name + '」的責任業務',
      intro: '被指定的業務若名下還沒有這家公司的名片，登入時會收到「請補名片」待辦；補上名片後自動消失。取消勾選＝移除。',
      current: company.responsibleSales || [],
      submitLabel: '儲存責任業務',
      onSubmit: function (list) { return jpost('/api/companies/' + encodeURIComponent(company.id) + '/set-responsible-sales', { responsibleSales: list }); },
      onSaved: function (data) {
        toast(data && data.changed ? '✓ 責任業務已更新' + (data.notifiedUsers ? '，已通知 ' + data.notifiedUsers + ' 位業務' : '') : '責任業務沒有變更');
        if (onSaved) onSaved(data);
      },
    });
  }
  function openBatchPicker(ids, onDone) {
    return openPicker({
      title: '批次指定 ' + ids.length + ' 家公司的責任業務',
      intro: '選好業務後套用到你勾選的公司。每位業務只會收到一則彙總通知。',
      current: [],
      withMode: true,
      submitLabel: '套用',
      onSubmit: function (list, mode) { return jpost('/api/companies/responsible-sales-batch', { companyIds: ids, usernames: list, mode: mode }); },
      onSaved: function (data) {
        toast('✓ 已更新 ' + (data.updated || 0) + ' 家，未變更 ' + (data.unchanged || 0) + ' 家' + ((data.denied && data.denied.length) ? '，' + data.denied.length + ' 家無權限' : ''));
        if (onDone) onDone(data);
      },
    });
  }

  /* ───────── 行銷：處理業務回報 ───────── */
  function openReassign(f) {
    var cur = (f.responsibleSales || []).filter(function (u) { return u !== f.username; });
    return openPicker({
      title: '重新指定「' + f.companyName + '」的責任業務',
      intro: (f.userLabel || f.username) + ' 回報：' + (f.reasonLabel || '') + (f.note ? '（' + f.note + '）' : '') + '。重新指定後，被選到的業務（包含原本的人，如果你仍選他）會重新收到待辦，原本的解除會失效。',
      current: cur,
      submitLabel: '重新指定',
      onSubmit: function (list) { return jpost('/api/marketing/responsibility-feedback/' + encodeURIComponent(f.id) + '/resolve', { action: 'reassign', responsibleSales: list }); },
      onSaved: function () { toast('✓ 已重新指定責任業務'); refresh(); },
    });
  }
  function openHandled(f) {
    var d = openDialog({ title: '標示「' + f.companyName + '」已處理', role: 'dialog' });
    d.body.appendChild(el('p', 'cct-lead', '標示已處理後，這筆回報會從清單移除；該業務的待辦維持解除狀態。可附註你做了什麼（選填）。'));
    var ta = document.createElement('textarea'); ta.className = 'cct-textarea'; ta.rows = 3; ta.maxLength = 300; ta.id = uid('hn');
    var lab = el('label', 'cct-field-label', '備註（選填）'); lab.setAttribute('for', ta.id);
    d.body.appendChild(lab); d.body.appendChild(ta);
    var err = el('div', 'cct-error'); err.setAttribute('role', 'alert'); err.hidden = true; d.body.appendChild(err);
    var ok = button('標示已處理', 'cct-btn cct-btn-primary');
    d.foot.appendChild(button('取消', 'cct-btn cct-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      ok.disabled = true;
      jpost('/api/marketing/responsibility-feedback/' + encodeURIComponent(f.id) + '/resolve', { action: 'handled', note: ta.value.trim() }).then(function (r) {
        if (!r.ok) { ok.disabled = false; err.textContent = (r.data && r.data.error) || '失敗'; err.hidden = false; return; }
        d.close(); toast('✓ 已標示處理'); refresh();
      });
    });
    focusFirst(d, 'textarea');
  }

  /* ───────── 載入與啟動 ───────── */
  function refresh() {
    if (!state.user) return Promise.resolve(state);
    var role = state.user.role;
    if (role === 'tecopm') return Promise.resolve(state);   // 集團PM 唯讀、API 白名單外
    var seq = ++state.seq;
    var jobs = [
      jget('/api/my-tasks/company-cards'),
      (role === 'marketing' || role === 'admin') ? jget('/api/marketing/responsibility-feedback') : Promise.resolve(null),
      (['admin', 'executive', 'manager1', 'manager2'].indexOf(role) >= 0) ? jget('/api/manager/company-card-overdue') : Promise.resolve(null),
    ];
    return Promise.all(jobs).then(function (r) {
      if (seq !== state.seq) return state;   // 較新的請求已發出 → 丟棄過期結果
      state.tasks = (r[0] && r[0].enabled) ? r[0] : null;
      state.feedback = (r[1] && r[1].enabled) ? r[1] : null;
      state.overdue = (r[2] && r[2].enabled) ? r[2] : null;
      renderHosts();
      return state;
    });
  }
  var refreshTimer = null;
  function refreshSoon() { clearTimeout(refreshTimer); refreshTimer = setTimeout(function () { refresh(); }, 400); }

  function hasAnnouncement() { return !!document.querySelector('.ann-overlay'); }
  function whenClear(fn, tries) {
    if (!hasAnnouncement() || tries <= 0) { fn(); return; }
    setTimeout(function () { whenClear(fn, tries - 1); }, 600);
  }

  function start(user, opts) {
    state.user = user || state.user;
    if (!state.user) return Promise.resolve();
    return refresh().then(function () {
      var tasks = (state.tasks && state.tasks.items) || [];
      if (opts && opts.login && !state.loginShown && tasks.length) {
        state.loginShown = true;
        // 公告彈窗（強制閱讀）可能剛好在載入：等它關掉再顯示，兩個視窗不會疊在一起
        setTimeout(function () { whenClear(function () { openSummary(); }, 100); }, 900);
      }
      checkHash();
    });
  }
  function checkHash() {
    if (window.location.hash === '#cardtasks') {
      try { history.replaceState(null, '', window.location.pathname + window.location.search); } catch (e) { /* ignore */ }
      openSummary({ fetch: true });
    }
  }
  window.addEventListener('hashchange', function () { if (state.user) checkHash(); });

  window.ITTSCardTasks = {
    start: start, refresh: refresh, refreshSoon: refreshSoon,
    openSummary: openSummary, openFromNotification: function () { return openSummary({ fetch: true }); },
    openCompanyPicker: openCompanyPicker, openBatchPicker: openBatchPicker, canAssign: canAssign,
    _state: state,
  };
})();
