/* 統編查無 → 申請行銷建檔（使用者端）
 *
 * 後端：GET /api/company-requests、POST /api/company-requests/:id/approve|reject|cancel；名片表單帶 companyRequest{category,note} 存檔
 *       （POST/PUT /api/contacts）。COMPANY_REQUESTS=off 時後端回 404 → 這裡 enabled()=false，名片流程與沒有此功能時相同。
 *
 * 畫面：
 *   - 名片表單：統編查無時出現「申請行銷建檔」區塊（類別＋說明）；其他狀態不顯示
 *   - 名片列表／詳情：「待行銷確認公司」黃色徽章、「建檔被退回」紅色徽章＋原因；詳情有取消申請／重新申請
 *   - 行銷／管理員：企業主檔頁與首頁待辦區的「建檔申請」清單，逐筆核可／退回，逾 7 天標紅
 * 安全：全部動態內容一律 createElement＋textContent（不經 innerHTML）、不使用 inline style／on* 屬性；樣式在 style.css 的 .creq-*。
 * 對外：window.ITTSCompanyReq = { start, refresh, enabled, mountForm, showForm, formError, getRequest, resetForm,
 *                                  decorateCard, decorateView, openApply, openFromNotification }
 */
(function () {
  'use strict';
  if (window.ITTSCompanyReq) return;

  var FALLBACK_CATS = [
    { key: 'gov', label: '政府機關' }, { key: 'school', label: '學校' }, { key: 'hospital', label: '公立醫院' },
    { key: 'foundation', label: '財團法人' }, { key: 'association', label: '協會團體' }, { key: 'shop', label: '行號' }, { key: 'other', label: '其他' },
  ];
  var FALLBACK_REASONS = [
    { key: 'typo', label: '統編疑似輸入錯誤' }, { key: 'no_need', label: '不需另建檔' }, { key: 'insufficient', label: '資料不足' },
  ];
  var state = { user: null, on: false, data: null, seq: 0, timer: null };
  var stack = [];
  var uidN = 0;
  var form = { host: null, sel: null, ta: null, count: null, onChange: null, visible: false };

  function uid(p) { uidN++; return 'creq-' + p + '-' + uidN; }
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
  function jsend(method, url, body) {
    return fetch(url, { method: method, credentials: 'same-origin', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) })
      .then(function (r) { return r.json().catch(function () { return {}; }).then(function (d) { return { ok: r.ok, status: r.status, data: d }; }); })
      .catch(function () { return { ok: false, status: 0, data: { error: '網路錯誤，請稍後再試' } }; });
  }
  function cats() { return (state.data && state.data.categories && state.data.categories.length) ? state.data.categories : FALLBACK_CATS; }
  function noteMin() { return (state.data && state.data.noteMin) || 5; }
  function noteMax() { return (state.data && state.data.noteMax) || 200; }
  function enabled() { return !!state.on; }
  function canResolve() { return !!(state.data && state.data.canResolve); }
  function myName() { return window._myUsername || (state.user && state.user.username) || ''; }

  /* ───────── 對話框（堆疊；焦點圈住；Esc 只處理最上層）───────── */
  function onKey(e) {
    if (!stack.length) return;
    var top = stack[stack.length - 1];
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); top.close(); return; }
    if (e.key === 'Tab') {
      var f = top.dlg.querySelectorAll('button:not([disabled]),input:not([disabled]),textarea:not([disabled]),select:not([disabled])');
      if (!f.length) return;
      var first = f[0], last = f[f.length - 1];
      if (e.shiftKey && (document.activeElement === first || !top.dlg.contains(document.activeElement))) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && (document.activeElement === last || !top.dlg.contains(document.activeElement))) { e.preventDefault(); first.focus(); }
    }
  }
  var keyBound = false;
  function openDialog(title) {
    if (!keyBound) { document.addEventListener('keydown', onKey, true); keyBound = true; }
    var prev = document.activeElement;
    var ov = el('div', 'creq-overlay');
    var dlg = el('div', 'creq-dialog');
    dlg.setAttribute('role', 'dialog');
    dlg.setAttribute('aria-modal', 'true');
    var h = el('h2', 'creq-dlg-title', title);
    h.id = uid('t');
    dlg.setAttribute('aria-labelledby', h.id);
    var body = el('div', 'creq-dlg-body');
    var foot = el('div', 'creq-dlg-foot');
    dlg.appendChild(h); dlg.appendChild(body); dlg.appendChild(foot);
    ov.appendChild(dlg);
    var entry = {
      dlg: dlg, body: body, foot: foot,
      close: function () {
        var i = stack.indexOf(entry);
        if (i >= 0) stack.splice(i, 1);
        if (ov.parentNode) ov.parentNode.removeChild(ov);
        if (prev && prev.focus && document.contains(prev)) { try { prev.focus(); } catch (e) { /* ignore */ } }
      },
    };
    stack.push(entry);
    document.body.appendChild(ov);
    return entry;
  }
  function field(d, labelText, node) {
    var lab = el('label', 'creq-field-label', labelText);
    node.id = node.id || uid('f');
    lab.setAttribute('for', node.id);
    d.body.appendChild(lab); d.body.appendChild(node);
  }
  function errBox(d) { var e = el('div', 'creq-error'); e.setAttribute('role', 'alert'); e.hidden = true; d.body.appendChild(e); return e; }
  function showErr(box, msg) { box.textContent = msg; box.hidden = false; }

  /* ───────── 名片表單：申請區塊 ───────── */
  function buildCatSelect() {
    var sel = el('select', 'creq-select');
    var o0 = el('option', null, '請選擇申請類別…'); o0.value = '';
    sel.appendChild(o0);
    cats().forEach(function (c) { var o = el('option', null, c.label); o.value = c.key; sel.appendChild(o); });
    return sel;
  }
  function mountForm(host, onChange) {
    form.host = host; form.onChange = onChange || null;
    host.textContent = '';
    host.className = 'creq-form';
    host.hidden = true;
    host.appendChild(el('div', 'creq-form-title', '申請行銷建檔'));
    host.appendChild(el('p', 'creq-form-hint', '政府機關、學校、公立醫院、財團法人、協會、行號等不在公司登記資料中。請先確認統編沒有打錯；確定無誤，存檔時會自動送出建檔申請，由行銷核對後建立企業主檔（名片照常建立，建檔前暫時不能送簽報價單、推送 SAP）。'));
    var l1 = el('label', 'creq-field-label', '申請類別（必選）');
    form.sel = buildCatSelect(); form.sel.id = uid('cat'); l1.setAttribute('for', form.sel.id);
    var l2 = el('label', 'creq-field-label', '說明（必填，至少 ' + noteMin() + ' 字）');
    form.ta = el('textarea', 'creq-textarea'); form.ta.id = uid('note'); form.ta.rows = 2; form.ta.maxLength = noteMax();
    form.ta.placeholder = '例如：衛福部健保署北區業務組';
    l2.setAttribute('for', form.ta.id);
    form.count = el('div', 'creq-count', '');
    host.appendChild(l1); host.appendChild(form.sel); host.appendChild(l2); host.appendChild(form.ta); host.appendChild(form.count);
    var fire = function () { updateCount(); if (form.onChange) form.onChange(); };
    form.sel.addEventListener('change', fire);
    form.ta.addEventListener('input', fire);
    updateCount();
  }
  function noteLen() { return form.ta ? Array.from(form.ta.value.trim()).length : 0; }
  function updateCount() { if (form.count) form.count.textContent = noteLen() + ' / ' + noteMax(); }
  function showForm(v) { if (form.host) { form.visible = !!v; form.host.hidden = !v; } }
  function resetForm() { if (form.sel) form.sel.value = ''; if (form.ta) form.ta.value = ''; updateCount(); showForm(false); }
  function formError() {
    if (!enabled() || !form.sel) return '無法申請建檔，請聯繫行銷人員';
    if (!form.sel.value) return '請選擇建檔申請的類別';
    if (noteLen() < noteMin()) return '請填寫建檔申請說明（至少 ' + noteMin() + ' 字）';
    return '';
  }
  function getRequest() { return form.sel ? { category: form.sel.value, note: form.ta.value.trim() } : null; }

  /* ───────── 名片徽章／詳情提示 ───────── */
  function badgeOf(c) {
    if (!c || (c.companyVerify !== 'pending' && c.companyVerify !== 'rejected')) return null;
    var rej = c.companyVerify === 'rejected';
    var b = el('span', 'creq-badge ' + (rej ? 'creq-badge-rejected' : 'creq-badge-pending'), rej ? '建檔被退回' : '待行銷確認公司');
    b.title = rej ? ('建檔申請被退回：' + (c.companyVerifyReason || '') + '。請修正統編') : '統編查無公司登記，已申請行銷建檔；確認前暫時不能送簽報價單、推送 SAP';
    return b;
  }
  function decorateCard(card, c) {
    try {
      var b = badgeOf(c);
      if (!b || !card) return;
      var host = card.querySelector('.card-rel-tags') || card.querySelector('.card-rel-row') || card;
      host.appendChild(b);
    } catch (e) { /* ignore */ }
  }
  function decorateView(body, c) {
    try {
      var b = badgeOf(c);
      if (!b || !body) return;
      var rej = c.companyVerify === 'rejected';
      var box = el('div', 'creq-notice ' + (rej ? 'creq-notice-rejected' : 'creq-notice-pending'));
      box.setAttribute('role', 'status');
      box.appendChild(b);
      var msg = rej
        ? ('建檔申請被退回：' + (c.companyVerifyReason || '') + '。請修正統編（編輯名片），或按「重新申請」說明原因。退回期間不能送簽報價單、推送 SAP。')
        : (c.companyRequestId
          ? '統編查無公司登記，已送出建檔申請，等行銷確認。確認前暫時不能送簽報價單、推送 SAP；其他功能不受影響。'
          : '建檔申請已取消。可以編輯名片修正統編，或按「重新申請」再送一次。');
      box.appendChild(el('p', 'creq-notice-text', msg));
      if (enabled() && c.owner === myName()) {
        var act = el('div', 'creq-notice-actions');
        if (c.companyVerify === 'pending' && c.companyRequestId) {
          act.appendChild(button('取消申請', 'creq-btn creq-btn-ghost creq-btn-sm', function () { openCancel(c); }));
        } else {
          act.appendChild(button('重新申請', 'creq-btn creq-btn-primary creq-btn-sm', function () { openApply(c); }));
        }
        box.appendChild(act);
      }
      body.insertBefore(box, body.firstChild);
    } catch (e) { /* ignore */ }
  }

  function afterChange() {
    refresh();
    try { if (typeof window.loadContacts === 'function') { var si = document.getElementById('searchInput'); window.loadContacts(si ? si.value.trim() : ''); } } catch (e) { /* ignore */ }
    try { var ov = document.getElementById('viewOverlay'); if (ov) ov.classList.remove('open'); } catch (e) { /* ignore */ }
  }

  /* ───────── 業務：重新申請／取消 ───────── */
  function openApply(contact) {
    var d = openDialog('重新申請「' + (contact.company || '') + '」建檔');
    d.body.appendChild(el('p', 'creq-lead', '統編 ' + (contact.taxId || '') + ' 查無公司登記。請先確認統編沒有打錯；確定無誤再送出，由行銷核對後建立企業主檔。'));
    var sel = buildCatSelect();
    field(d, '申請類別（必選）', sel);
    var ta = el('textarea', 'creq-textarea'); ta.rows = 3; ta.maxLength = noteMax(); ta.placeholder = '例如：衛福部健保署北區業務組';
    field(d, '說明（必填，至少 ' + noteMin() + ' 字）', ta);
    var err = errBox(d);
    var ok = button('送出申請', 'creq-btn creq-btn-primary');
    d.foot.appendChild(button('取消', 'creq-btn creq-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      if (!sel.value) { showErr(err, '請選擇申請類別'); return; }
      if (Array.from(ta.value.trim()).length < noteMin()) { showErr(err, '說明至少 ' + noteMin() + ' 個字'); return; }
      ok.disabled = true;
      jsend('PUT', '/api/contacts/' + encodeURIComponent(contact.id), { companyRequest: { category: sel.value, note: ta.value.trim() } }).then(function (r) {
        if (!r.ok) { ok.disabled = false; showErr(err, (r.data && r.data.error) || '送出失敗'); return; }
        d.close(); toast('✓ 已送出建檔申請'); afterChange();
      });
    });
    sel.focus();
  }
  function openCancel(contact) {
    var d = openDialog('取消建檔申請');
    d.body.appendChild(el('p', 'creq-lead', '取消後名片仍保留「待行銷確認公司」狀態（暫時不能送簽報價單、推送 SAP）。你可以修正統編，或之後再按「重新申請」。確定要取消嗎？'));
    var err = errBox(d);
    var ok = button('取消這筆申請', 'creq-btn creq-btn-danger');
    d.foot.appendChild(button('先不要', 'creq-btn creq-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      ok.disabled = true;
      jsend('POST', '/api/company-requests/' + encodeURIComponent(contact.companyRequestId) + '/cancel', {}).then(function (r) {
        if (!r.ok) { ok.disabled = false; showErr(err, (r.data && r.data.error) || '取消失敗'); return; }
        d.close(); toast('✓ 已取消申請'); afterChange();
      });
    });
  }

  /* ───────── 行銷：核可／退回 ───────── */
  function openApprove(it) {
    var d = openDialog('核可建檔：' + it.companyName);
    d.body.appendChild(el('p', 'creq-lead', '統編 ' + it.taxId + '（' + it.categoryLabel + '）。核可後建立企業主檔，之後任何人輸入此統編都會直接帶入正式名稱；同統編的待確認名片會一併解除。'));
    it.requesters.forEach(function (x) {
      d.body.appendChild(el('p', 'creq-meta', '申請人 ' + x.label + '　填寫名稱「' + x.companyName + '」　說明：' + x.note));
    });
    var name = el('input', 'creq-input'); name.type = 'text'; name.maxLength = 100; name.value = it.companyName || '';
    field(d, '公司正式名稱（業務填的可能只是子單位，請改成官方名稱）', name);
    var addr = el('input', 'creq-input'); addr.type = 'text'; addr.maxLength = 200;
    field(d, '地址（選填）', addr);
    var ind = el('input', 'creq-input'); ind.type = 'text'; ind.maxLength = 50;
    var dl = document.createElement('datalist'); dl.id = uid('dl'); ind.setAttribute('list', dl.id);
    field(d, '產業（選填，只有行銷／管理員能填）', ind);
    d.body.appendChild(dl);
    jget('/api/industries').then(function (arr) {
      (Array.isArray(arr) ? arr : []).forEach(function (n) { var o = document.createElement('option'); o.value = n; dl.appendChild(o); });
    });
    var err = errBox(d);
    var ok = button('核可並建立主檔', 'creq-btn creq-btn-primary');
    d.foot.appendChild(button('取消', 'creq-btn creq-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      if (!name.value.trim()) { showErr(err, '請填寫公司正式名稱'); return; }
      ok.disabled = true;
      jsend('POST', '/api/company-requests/' + encodeURIComponent(it.id) + '/approve', { name: name.value.trim(), address: addr.value.trim(), industry: ind.value.trim() }).then(function (r) {
        if (!r.ok) { ok.disabled = false; showErr(err, (r.data && r.data.error) || '核可失敗'); return; }
        d.close(); toast('✓ 已核可並建立企業主檔'); afterResolve();
      });
    });
    name.focus();
  }
  function openReject(it) {
    var d = openDialog('退回建檔申請：' + it.companyName);
    d.body.appendChild(el('p', 'creq-lead', '統編 ' + it.taxId + '。退回後申請人的名片會顯示「建檔被退回」與原因，並收到通知；修正統編後會自動重新送出。'));
    var fs = el('fieldset', 'creq-fieldset');
    fs.appendChild(el('legend', 'creq-legend', '退回原因（必選）'));
    var rn = uid('reason');
    var reasons = (state.data && state.data.rejectReasons && state.data.rejectReasons.length) ? state.data.rejectReasons : FALLBACK_REASONS;
    reasons.forEach(function (r) {
      var lab = el('label', 'creq-radio');
      var inp = document.createElement('input'); inp.type = 'radio'; inp.name = rn; inp.value = r.key;
      lab.appendChild(inp); lab.appendChild(el('span', null, r.label)); fs.appendChild(lab);
    });
    d.body.appendChild(fs);
    var ta = el('textarea', 'creq-textarea'); ta.rows = 3; ta.maxLength = noteMax();
    field(d, '備註（選填，會顯示給申請人）', ta);
    var err = errBox(d);
    var ok = button('退回', 'creq-btn creq-btn-danger');
    d.foot.appendChild(button('取消', 'creq-btn creq-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      var chosen = d.dlg.querySelector('input[name="' + rn + '"]:checked');
      if (!chosen) { showErr(err, '請選擇退回原因'); return; }
      ok.disabled = true;
      jsend('POST', '/api/company-requests/' + encodeURIComponent(it.id) + '/reject', { reason: chosen.value, note: ta.value.trim() }).then(function (r) {
        if (!r.ok) { ok.disabled = false; showErr(err, (r.data && r.data.error) || '退回失敗'); return; }
        d.close(); toast('✓ 已退回申請'); afterResolve();
      });
    });
  }
  function afterResolve() {
    refresh();
    try { if (typeof window.loadCompanyMasterView === 'function') window.loadCompanyMasterView(); } catch (e) { /* ignore */ }
  }

  /* ───────── 行銷清單（企業主檔頁＋首頁待辦區）───────── */
  function pendingItems() { return ((state.data && state.data.items) || []).filter(function (x) { return x.status === 'pending'; }); }
  function row(it) {
    var li = el('li', 'creq-row' + (it.overdue ? ' creq-row-overdue' : ''));
    var main = el('div', 'creq-row-main');
    var top = el('div', 'creq-row-top');
    top.appendChild(el('strong', 'creq-name', it.companyName || '(未填名稱)'));
    top.appendChild(el('span', 'creq-chip', it.categoryLabel));
    top.appendChild(el('span', 'creq-chip' + (it.overdue ? ' creq-chip-danger' : ''), it.days > 0 ? ('已等 ' + it.days + ' 天' + (it.overdue ? '，逾期' : '')) : '今天送出'));
    main.appendChild(top);
    main.appendChild(el('span', 'creq-meta', '統編 ' + it.taxId));
    it.requesters.forEach(function (x) {
      main.appendChild(el('span', 'creq-meta', '申請人 ' + x.label + (x.companyName && x.companyName !== it.companyName ? '（填寫「' + x.companyName + '」）' : '') + '：' + x.note));
    });
    li.appendChild(main);
    var act = el('div', 'creq-row-actions');
    act.appendChild(button('核可…', 'creq-btn creq-btn-primary creq-btn-sm', function () { openApprove(it); }));
    act.appendChild(button('退回…', 'creq-btn creq-btn-ghost creq-btn-sm', function () { openReject(it); }));
    li.appendChild(act);
    return li;
  }
  function panel(items, showEmpty) {
    var anyOverdue = items.some(function (x) { return x.overdue; });
    var sec = el('section', 'creq-panel' + (anyOverdue ? ' creq-panel-overdue' : ''));
    var hid = uid('h');
    sec.setAttribute('aria-labelledby', hid);
    var hd = el('div', 'creq-panel-hd');
    var h = el('h3', 'creq-panel-title', '🏛️ 建檔申請');
    h.id = hid;
    hd.appendChild(h);
    hd.appendChild(el('span', 'creq-count', String(items.length)));
    sec.appendChild(hd);
    if (!items.length) {
      if (!showEmpty) return null;
      sec.appendChild(el('p', 'creq-sub', '目前沒有待處理的建檔申請。'));
      return sec;
    }
    sec.appendChild(el('p', 'creq-sub', '業務查無統編時送出的建檔申請。核可＝建立企業主檔；超過 ' + ((state.data && state.data.overdueDays) || 7) + ' 天未處理會標紅。'));
    var ul = el('ul', 'creq-list');
    items.forEach(function (it) { ul.appendChild(row(it)); });
    sec.appendChild(ul);
    return sec;
  }
  function renderHosts() {
    var hosts = document.querySelectorAll('[data-creq-host]');
    var items = canResolve() ? pendingItems() : [];
    [].forEach.call(hosts, function (h) {
      h.textContent = '';
      var show = state.on && canResolve();
      var p = show ? panel(items, h.getAttribute('data-creq-host') === 'cm') : null;
      if (p) { h.appendChild(p); h.hidden = false; } else { h.hidden = true; }
    });
  }

  /* ───────── 載入與啟動 ───────── */
  function refresh() {
    if (!state.user) return Promise.resolve();
    if (state.user.role === 'tecopm') { state.on = false; return Promise.resolve(); }
    var seq = ++state.seq;
    return jget('/api/company-requests').then(function (d) {
      if (seq !== state.seq) return;
      state.on = !!(d && d.enabled);
      state.data = state.on ? d : null;
      renderHosts();
    });
  }
  function checkHash() {
    if (window.location.hash === '#creq') {
      try { history.replaceState(null, '', window.location.pathname + window.location.search); } catch (e) { /* ignore */ }
      openFromNotification();
    }
  }
  function openFromNotification() {
    try { if (typeof window.showSection === 'function') window.showSection('companyMaster'); } catch (e) { /* ignore */ }
    return refresh().then(function () {
      var h = document.querySelector('[data-creq-host="cm"]');
      if (h && !h.hidden && h.scrollIntoView) { try { h.scrollIntoView({ block: 'start' }); } catch (e) { /* ignore */ } }
    });
  }
  function start(user) {
    state.user = user || state.user;
    if (!state.user) return Promise.resolve();
    clearInterval(state.timer);
    state.timer = setInterval(function () { if (!document.hidden && state.on && canResolve()) refresh(); }, 180000);
    return refresh().then(checkHash);
  }
  window.addEventListener('hashchange', function () { if (state.user) checkHash(); });

  window.ITTSCompanyReq = {
    start: start, refresh: refresh, enabled: enabled,
    mountForm: mountForm, showForm: showForm, formError: formError, getRequest: getRequest, resetForm: resetForm,
    decorateCard: decorateCard, decorateView: decorateView, openApply: openApply, openFromNotification: openFromNotification,
    _state: state,
  };
})();
