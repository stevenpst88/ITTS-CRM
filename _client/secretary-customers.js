/* 秘書唯讀客戶視圖（後端 SECRETARY_CUSTOMERS=1 才有；關閉時探測回 404，本檔不建立任何畫面元素）。
 * 同 BU、已成交的客戶：公司／聯絡人／職稱／BU／負責業務／郵寄地址；電話、Email、備註等個資由後端遮蔽，這裡只顯示「（個資已遮蔽）」。
 * 全部用 createElement／textContent，不使用 innerHTML。樣式 style.css 的 .scu-*。 */
(function () {
  'use strict';
  var API = '/api/secretary/customers';
  var MASK = '（個資已遮蔽）', EMPTY = '（未填）', FALLBACK_NOTE = '（公司登記地址，未另填郵寄地址）';
  var probeP = null, available = false, list = [], navEl = null, viewEl = null, listEl = null, countEl = null, searchEl = null, errEl = null, panel = null;

  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function $(id) { return document.getElementById(id); }

  function fetchList() {
    return fetch(API, { credentials: 'include' }).then(function (r) {
      if (!r.ok) return null;
      return r.json().then(function (j) { return (j && Array.isArray(j.customers)) ? j.customers : null; });
    }).catch(function () { return null; });
  }

  /** 探測：只有 200 才建立導覽與畫面；回傳 Promise<boolean>（memoize） */
  function probe() {
    if (probeP) return probeP;
    probeP = fetchList().then(function (rows) {
      if (!rows) return false;
      available = true; list = rows; build(); return true;
    });
    return probeP;
  }

  function build() {
    var anchor = $('navQuotations');
    var nav = $('navSecretaryCustomers');
    if (!nav) {
      nav = el('div', 'sidebar-nav-item'); nav.id = 'navSecretaryCustomers'; nav.setAttribute('data-section', 'secretaryCustomers');
      nav.setAttribute('role', 'button'); nav.tabIndex = 0;
      nav.appendChild(el('span', 'nav-icon', '📮'));
      nav.appendChild(el('span', 'nav-label', '我的客戶（唯讀）'));
      if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(nav, anchor.nextSibling);
      var go = function () { if (typeof showSection === 'function') showSection('secretaryCustomers'); };
      nav.addEventListener('click', go);
      nav.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } });
    }
    navEl = nav;
    if (!$('secretaryCustomersView')) {
      var host = $('quotationsView') && $('quotationsView').parentNode;
      if (!host) return;
      viewEl = el('div', 'scu-view'); viewEl.id = 'secretaryCustomersView'; viewEl.style.display = 'none';
      var head = el('div', 'scu-head');
      head.appendChild(el('h2', 'scu-title', '我的客戶（唯讀）'));
      head.appendChild(el('div', 'scu-sub', '同 BU、已成交的客戶。協助業務寄信用：可查看並複製郵寄地址；電話、Email、備註等個資已遮蔽。'));
      var bar = el('div', 'scu-bar');
      searchEl = el('input', 'scu-search'); searchEl.type = 'search'; searchEl.placeholder = '搜尋公司／聯絡人／職稱／業務／地址'; searchEl.setAttribute('aria-label', '搜尋客戶');
      searchEl.addEventListener('input', render);
      countEl = el('span', 'scu-count', '');
      bar.appendChild(searchEl); bar.appendChild(countEl);
      errEl = el('div', 'scu-err'); errEl.style.display = 'none';
      listEl = el('div', 'scu-list');
      viewEl.appendChild(head); viewEl.appendChild(bar); viewEl.appendChild(errEl); viewEl.appendChild(listEl);
      host.appendChild(viewEl);
    }
  }

  function mailText(c) { return c.mailingAddress || ''; }
  function matches(c, kw) {
    if (!kw) return true;
    return [c.company, c.name, c.title, c.ownerDisplayName, mailText(c), (c.bu || []).join(' ')].some(function (s) { return String(s || '').toLowerCase().indexOf(kw) >= 0; });
  }

  function addrNode(c, big) {
    var wrap = el('div', 'scu-addr' + (big ? ' scu-addr-big' : ''));
    var t = mailText(c);
    wrap.appendChild(el('span', t ? 'scu-addr-text' : 'scu-muted', t || EMPTY));
    if (t && c.mailingFallback) wrap.appendChild(el('span', 'scu-fallback', FALLBACK_NOTE));
    return wrap;
  }

  function cell(label, node) {
    var d = el('div', 'scu-cell'); d.appendChild(el('span', 'scu-lbl', label));
    if (typeof node === 'string') d.appendChild(el('span', 'scu-val', node)); else d.appendChild(node);
    return d;
  }

  function render() {
    if (!listEl) return;
    var kw = searchEl ? searchEl.value.trim().toLowerCase() : '';
    var rows = list.filter(function (c) { return matches(c, kw); });
    while (listEl.firstChild) listEl.removeChild(listEl.firstChild);
    countEl.textContent = '共 ' + rows.length + ' 筆' + (kw ? '（全部 ' + list.length + '）' : '');
    if (!rows.length) { listEl.appendChild(el('div', 'scu-empty', kw ? '沒有符合的客戶' : '目前沒有同 BU 的已成交客戶')); return; }
    rows.forEach(function (c) {
      var r = el('div', 'scu-row'); r.tabIndex = 0; r.setAttribute('role', 'button'); r.setAttribute('aria-label', '查看 ' + c.company + ' ' + c.name + ' 詳情');
      r.dataset.id = c.id;
      var who = el('div', 'scu-cell scu-who');
      who.appendChild(el('span', 'scu-co', c.company || EMPTY));
      who.appendChild(el('span', 'scu-nm', (c.name || EMPTY) + (c.title ? '・' + c.title : '')));
      r.appendChild(who);
      r.appendChild(cell('負責業務', (c.ownerDisplayName || EMPTY) + ((c.bu && c.bu.length) ? '（' + c.bu.join('/') + '）' : '')));
      r.appendChild(cell('電話', el('span', 'scu-muted', MASK)));
      r.appendChild(cell('Email', el('span', 'scu-muted', MASK)));
      r.appendChild(cell('郵寄地址', addrNode(c, false)));
      var open = function () { openDetail(c); };
      r.addEventListener('click', open);
      r.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); open(); } });
      listEl.appendChild(r);
    });
  }

  function closeDetail() {
    if (!panel) return;
    document.removeEventListener('keydown', panel._esc, true);
    var back = panel._back; panel.remove(); panel = null;
    if (back && back.focus) { try { back.focus(); } catch (e) { /* ignore */ } }
  }

  function copyText(text, done) {
    function legacy() {
      try {
        var ta = document.createElement('textarea'); ta.value = text; ta.setAttribute('readonly', ''); ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.select(); var ok = document.execCommand('copy'); ta.remove(); done(!!ok);
      } catch (e) { done(false); }
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { done(true); }, legacy);
    } else legacy();
  }

  function openDetail(c0) {
    closeDetail();
    var back = document.activeElement;
    var ov = el('div', 'scu-overlay'); ov.setAttribute('role', 'dialog'); ov.setAttribute('aria-modal', 'true'); ov.setAttribute('aria-label', '客戶詳情（唯讀）');
    var box = el('div', 'scu-box');
    var h = el('div', 'scu-box-h');
    h.appendChild(el('h3', 'scu-box-t', '客戶詳情（唯讀）'));
    var x = el('button', 'scu-x', '×'); x.type = 'button'; x.setAttribute('aria-label', '關閉'); x.addEventListener('click', closeDetail);
    h.appendChild(x); box.appendChild(h);
    var body = el('div', 'scu-box-b'); box.appendChild(body);
    ov.appendChild(box);
    ov.addEventListener('click', function (e) { if (e.target === ov) closeDetail(); });
    panel = ov; panel._back = back;
    panel._esc = function (e) { if (e.key === 'Escape') { e.stopPropagation(); closeDetail(); } };
    document.addEventListener('keydown', panel._esc, true);
    document.body.appendChild(ov);
    x.focus();

    function fill(c) {
      while (body.firstChild) body.removeChild(body.firstChild);
      var kv = function (k, v, cls) { var r = el('div', 'scu-kv'); r.appendChild(el('span', 'scu-k', k)); r.appendChild(el('span', 'scu-v' + (cls ? ' ' + cls : ''), v)); body.appendChild(r); };
      kv('公司', c.company || EMPTY); kv('聯絡人', c.name || EMPTY); kv('職稱', c.title || EMPTY);
      kv('BU', (c.bu && c.bu.length) ? c.bu.join(' / ') : EMPTY); kv('負責業務', c.ownerDisplayName || EMPTY);
      kv('電話', MASK, 'scu-muted'); kv('手機', MASK, 'scu-muted'); kv('Email', MASK, 'scu-muted'); kv('備註', MASK, 'scu-muted');
      var m = el('div', 'scu-mail');
      m.appendChild(el('div', 'scu-mail-h', '郵寄地址'));
      m.appendChild(addrNode(c, true));
      var row = el('div', 'scu-mail-act');
      var btn = el('button', 'scu-copy', '複製郵寄地址'); btn.type = 'button'; btn.disabled = !mailText(c);
      var msg = el('span', 'scu-copied', ''); msg.setAttribute('role', 'status'); msg.setAttribute('aria-live', 'polite');
      btn.addEventListener('click', function () { copyText(mailText(c), function (ok) { msg.textContent = ok ? '已複製' : '複製失敗，請手動選取'; }); });
      row.appendChild(btn); row.appendChild(msg); m.appendChild(row);
      body.appendChild(m);
    }
    fill(c0);   // 先用清單內容顯示，再向後端「開啟詳情」（記稽核並回最新的白名單內容）
    fetch(API + '/' + encodeURIComponent(c0.id) + '/view', { method: 'POST', credentials: 'include', headers: { 'Content-Type': 'application/json' }, body: '{}' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (j) { if (j && j.customer && panel === ov) fill(j.customer); })
      .catch(function () { /* 清單內容已顯示 */ });
  }

  function load() {
    return fetchList().then(function (rows) {
      if (!errEl) return;
      if (rows) { list = rows; errEl.style.display = 'none'; }
      else { errEl.textContent = '載入失敗，顯示的是上次載入的內容'; errEl.style.display = ''; }
      render();
    });
  }

  window.ITTSSecCust = {
    start: function (user) { if (user && user.role === 'secretary') probe(); },
    /** showSection／showDashboard 每次切換呼叫：on＝目前是否為本區段 */
    setActive: function (on) {
      if (viewEl) viewEl.style.display = on ? '' : 'none';
      if (navEl) { if (on) navEl.classList.add('active'); else navEl.classList.remove('active'); }
      if (on && window.NavGroups) { try { window.NavGroups.onActive(); } catch (e) { /* 側欄分組：展開所在組並移動膠囊 */ } }
      if (!on) closeDetail();
    },
    /** 進入本區段（含 F5 還原）；不可用（開關關閉／非秘書）時退回首頁 */
    show: function () {
      probe().then(function (ok) {
        if (!ok) { if (typeof showDashboard === 'function') showDashboard(); return; }
        if (typeof currentSection !== 'undefined' && currentSection !== 'secretaryCustomers') return;   // 等待期間已切走
        window.ITTSSecCust.setActive(true);
        var t = $('topbarTitle'); if (t) t.textContent = '我的客戶（唯讀）';
        render(); load();
      });
    },
    _debug: function () { return { available: available, count: list.length }; }
  };
})();
