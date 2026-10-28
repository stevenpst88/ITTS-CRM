/* 企業主檔「篩選 TA」（EDM 用）：公司條件 → 聯絡人名單下載
 *
 * 後端：GET /api/companies/ta/options、POST /api/companies/ta/query、POST /api/companies/ta/export（lib/companyTa.js）。
 * 範圍＝呼叫者本來看得到的公司／名片（客戶池排除）；數量人人可看，預覽 Email 與下載要有「可下載客戶名單」權限。
 * 安全：動態內容一律 createElement＋textContent、不用 inline style／on* 屬性；樣式在 style.css 的 .cta-*。
 * 對外：window.ITTSCompanyTa = { mount }   由 app.js 的 loadCompanyMasterView 在列表載入後呼叫 mount()
 */
(function () {
  'use strict';
  if (window.ITTSCompanyTa) return;

  var NONE = '__none__';
  var STATUS_LABELS = [['won', '已成交客戶'], ['open', '已開商機（未成交）'], ['noOpp', '尚無商機'], ['newImport', '新匯入（尚無名片）']];
  var host = null, panel = null, toggleBtn = null, built = false, opened = false;
  var opts = null;               // /ta/options 結果
  var F = null, O = null;        // 目前條件
  var timer = null, seq = 0;
  var els = {};                  // 結果區元素

  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = String(text); return e; }
  function btn(label, cls, onClick) { var b = el('button', cls, label); b.type = 'button'; if (onClick) b.addEventListener('click', onClick); return b; }
  function toast(msg) { try { if (typeof window.showToast === 'function') window.showToast(msg); } catch (e) { /* ignore */ } }
  var uidN = 0; function uid() { uidN++; return 'cta-' + uidN; }

  function freshF() {
    return { keyword: '', industries: [], regions: [], capMinWan: '', capMaxWan: '', capNone: '', bus: [], statuses: [], responsible: [], services: [],
      hasTaxId: '', gcis: '', createdFrom: '', createdTo: '' };
  }
  function freshO() { return { requireEmail: true, activeOnly: true, dedupeEmail: true, onePerCompany: false, titleKeyword: '' }; }

  function jfetch(method, url, body) {
    var init = { method: method, credentials: 'same-origin', cache: 'no-store' };
    if (body) { init.headers = { 'Content-Type': 'application/json' }; init.body = JSON.stringify(body); }
    return fetch(url, init);
  }

  // ── 小元件 ──
  function field(label, node) { var w = el('div', 'cta-field'); w.appendChild(el('label', 'cta-label', label)); w.appendChild(node); return w; }
  function textInput(get, set, ph, type) {
    var i = el('input', 'cta-input'); i.type = type || 'text'; if (ph) i.placeholder = ph; i.value = get();
    i.addEventListener('input', function () { set(i.value); schedule(); }); return i;
  }
  function selectInput(items, get, set) {
    var s = el('select', 'cta-input');
    items.forEach(function (it) { var o = el('option', null, it[1]); o.value = it[0]; s.appendChild(o); });
    s.value = get(); s.addEventListener('change', function () { set(s.value); schedule(); }); return s;
  }
  function checkbox(label, get, set) {
    var id = uid(), w = el('label', 'cta-check'); w.setAttribute('for', id);
    var c = el('input'); c.type = 'checkbox'; c.id = id; c.checked = !!get();
    c.addEventListener('change', function () { set(c.checked); schedule(); });
    w.appendChild(c); w.appendChild(el('span', null, label)); return w;
  }
  // 多選下拉：<details> 內有搜尋框＋勾選清單；摘要顯示已選數
  function multi(title, items, arr) {
    var d = el('details', 'cta-multi'); var sm = el('summary', 'cta-multi-sum');
    var refresh = function () { sm.textContent = arr.length ? title + '（已選 ' + arr.length + '）' : title + '（不限）'; };
    refresh(); d.appendChild(sm);
    var body = el('div', 'cta-multi-body');
    var q = el('input', 'cta-input cta-multi-q'); q.type = 'search'; q.placeholder = '搜尋…'; q.setAttribute('aria-label', title + '搜尋');
    body.appendChild(q);
    var list = el('div', 'cta-multi-list'); body.appendChild(list);
    var rows = [];
    items.forEach(function (it) {
      var id = uid(), row = el('label', 'cta-multi-row'); row.setAttribute('for', id);
      var c = el('input'); c.type = 'checkbox'; c.id = id; c.checked = arr.indexOf(it.value) >= 0;
      c.addEventListener('change', function () {
        var i = arr.indexOf(it.value);
        if (c.checked && i < 0) arr.push(it.value); if (!c.checked && i >= 0) arr.splice(i, 1);
        refresh(); schedule();
      });
      row.appendChild(c); row.appendChild(el('span', null, it.label + (it.n != null ? '（' + it.n + '）' : '')));
      list.appendChild(row); rows.push({ row: row, text: (it.label || '').toLowerCase() });
    });
    q.addEventListener('input', function () { var k = q.value.trim().toLowerCase(); rows.forEach(function (r) { r.row.hidden = !!k && r.text.indexOf(k) < 0; }); });
    d.appendChild(body); d._clear = function () { rows.forEach(function (r) { r.row.querySelector('input').checked = false; }); refresh(); };
    return d;
  }
  function group(title, children) {
    var g = el('fieldset', 'cta-group'); g.appendChild(el('legend', 'cta-legend', title));
    var row = el('div', 'cta-row'); children.forEach(function (c) { row.appendChild(c); }); g.appendChild(row); return g;
  }

  // ── 組畫面 ──
  function build() {
    panel.textContent = '';
    var multis = [];
    var track = function (m) { multis.push(m); return m; };
    var ind = [{ value: NONE, label: '（未分類）' }].concat(opts.industries.map(function (x) { return { value: x.name, label: x.name, n: x.n }; }));
    var reg = [{ value: NONE, label: '（未填區域）' }].concat(opts.regions.map(function (x) { return { value: x.name, label: x.name, n: x.n }; }));
    var people = function (arr, withNone) { return (withNone ? [{ value: NONE, label: '（未指定）' }] : []).concat(arr.map(function (p) { return { value: p.username, label: p.label }; })); };

    var capRow = el('div', 'cta-cap');
    capRow.appendChild(textInput(function () { return F.capMinWan; }, function (v) { F.capMinWan = v; }, '最低', 'number'));
    capRow.appendChild(el('span', 'cta-cap-sep', '～'));
    capRow.appendChild(textInput(function () { return F.capMaxWan; }, function (v) { F.capMaxWan = v; }, '最高', 'number'));
    capRow.appendChild(el('span', 'cta-cap-unit', '萬元'));

    panel.appendChild(group('公司條件', [
      field('公司名稱／統編', textInput(function () { return F.keyword; }, function (v) { F.keyword = v; }, '關鍵字')),
      field('產業', track(multi('產業', ind, F.industries))),
      field('區域', track(multi('區域', reg, F.regions))),
      field('資本額區間', capRow),
      field('資本額未填的公司', selectInput([['', '依區間（有填寫區間時排除）'], ['include', '一併納入'], ['only', '只要未填的']], function () { return F.capNone; }, function (v) { F.capNone = v; })),
      field('BU', track(multi('BU', opts.bus.map(function (b) { return { value: b, label: b }; }), F.bus))),
      field('客戶狀態', track(multi('客戶狀態', STATUS_LABELS.map(function (s) { return { value: s[0], label: s[1] }; }), F.statuses))),
      field('責任業務', track(multi('責任業務', people(opts.responsible, true), F.responsible))),
      field('服務業務（名片擁有者）', track(multi('服務業務', people(opts.services, false), F.services))),
      field('有無統編', selectInput([['', '不限'], ['yes', '有統編'], ['no', '無統編']], function () { return F.hasTaxId; }, function (v) { F.hasTaxId = v; })),
      field('GCIS 補全', selectInput([['', '不限'], ['yes', '已補全'], ['no', '未補全']], function () { return F.gcis; }, function (v) { F.gcis = v; })),
      field('建立時間（起）', textInput(function () { return F.createdFrom; }, function (v) { F.createdFrom = v; }, '', 'date')),
      field('建立時間（迄）', textInput(function () { return F.createdTo; }, function (v) { F.createdTo = v; }, '', 'date')),
    ]));
    panel.appendChild(group('聯絡人條件', [
      checkbox('必須有 Email（格式正確）', function () { return O.requireEmail; }, function (v) { O.requireEmail = v; }),
      checkbox('只含在職（排除已離職）', function () { return O.activeOnly; }, function (v) { O.activeOnly = v; }),
      checkbox('相同 Email 只留一筆', function () { return O.dedupeEmail; }, function (v) { O.dedupeEmail = v; }),
      checkbox('每家公司只取一位（優先主要聯絡人）', function () { return O.onePerCompany; }, function (v) { O.onePerCompany = v; }),
      field('職稱關鍵字', textInput(function () { return O.titleKeyword; }, function (v) { O.titleKeyword = v; }, '例：經理、資訊、IT')),
    ]));

    var res = el('div', 'cta-result'); els.summary = el('div', 'cta-summary', '計算中…'); els.summary.setAttribute('aria-live', 'polite'); res.appendChild(els.summary);
    els.note = el('div', 'cta-note'); res.appendChild(els.note);
    els.prev = el('div', 'cta-preview'); res.appendChild(els.prev);
    var bar = el('div', 'cta-actions');
    els.dlC = btn('⬇ 下載聯絡人名單（EDM）', 'cta-btn cta-btn-primary', function () { download('contacts'); });
    els.dlM = btn('⬇ 只下載公司清單', 'cta-btn', function () { download('companies'); });
    bar.appendChild(els.dlC); bar.appendChild(els.dlM);
    bar.appendChild(btn('清除條件', 'cta-btn cta-btn-ghost', function () { F = freshF(); O = freshO(); multis = []; build(); schedule(0); }));
    res.appendChild(bar);
    panel.appendChild(res);
    setDownloadEnabled(opts.canDownload);
  }

  function setDownloadEnabled(on, perm) {
    [els.dlC, els.dlM].forEach(function (b) { if (b) b.disabled = !on; });
    if (els.note) els.note.textContent = (perm === undefined ? on : perm) ? '' : '你的帳號沒有「可下載客戶名單」權限，只能看符合的數量；需要下載請聯繫管理員開通。';
  }

  function payload(mode) {
    var f = JSON.parse(JSON.stringify(F));
    f.capMinWan = f.capMinWan === '' ? null : Number(f.capMinWan);
    f.capMaxWan = f.capMaxWan === '' ? null : Number(f.capMaxWan);
    var p = { filters: f, opts: O }; if (mode) p.mode = mode; return p;
  }

  function schedule(ms) { clearTimeout(timer); timer = setTimeout(query, ms == null ? 400 : ms); }
  function query() {
    var my = ++seq;
    if (els.summary) els.summary.textContent = '計算中…';
    jfetch('POST', '/api/companies/ta/query', payload()).then(function (r) {
      return r.json().then(function (d) { return { ok: r.ok, d: d }; });
    }).then(function (x) {
      if (my !== seq || !els.summary) return;
      if (!x.ok) { els.summary.textContent = '⚠ ' + (x.d.error || '查詢失敗'); return; }
      var s = x.d.stats;
      els.summary.textContent = '符合 ' + s.companies + ' 家公司；聯絡人 ' + s.contacts + ' 位（符合條件且可下載）；其中有 Email 者共 ' + s.withEmail + ' 位';
      setDownloadEnabled(!!x.d.canDownload && s.companies > 0, !!x.d.canDownload);
      renderPreview(x.d.preview || []);
    }).catch(function () { if (my === seq && els.summary) els.summary.textContent = '⚠ 網路錯誤，請稍後再試'; });
  }

  function renderPreview(rows) {
    var p = els.prev; p.textContent = '';
    if (!rows.length) return;
    p.appendChild(el('div', 'cta-prev-title', '預覽（前 ' + rows.length + ' 筆）'));
    var wrap = el('div', 'cta-prev-wrap'), t = el('table', 'cta-prev-table'), hd = el('tr');
    ['公司', '姓名', '職稱', 'Email'].forEach(function (h) { hd.appendChild(el('th', null, h)); });
    t.appendChild(hd);
    rows.forEach(function (r) { var tr = el('tr'); [r.company, r.contact, r.title, r.email].forEach(function (v) { tr.appendChild(el('td', null, v)); }); t.appendChild(tr); });
    wrap.appendChild(t); p.appendChild(wrap);
  }

  function download(mode) {
    var b = mode === 'companies' ? els.dlM : els.dlC, old = b.textContent;
    b.disabled = true; b.textContent = '下載中…';
    jfetch('POST', '/api/companies/ta/export', payload(mode)).then(function (r) {
      if (!r.ok) return r.json().catch(function () { return {}; }).then(function (d) { throw new Error(d.error || '下載失敗'); });
      return r.blob();
    }).then(function (blob) {
      var d = new Date(), ds = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
      var a = document.createElement('a'); a.href = URL.createObjectURL(blob);
      a.download = (mode === 'companies' ? '企業清單_TA_' : 'EDM名單_TA_') + ds + '.xlsx';
      document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(a.href);
    }).catch(function (e) { toast('⛔ ' + e.message); }).then(function () { b.textContent = old; b.disabled = false; });
  }

  function open() {
    opened = !opened; panel.hidden = !opened; toggleBtn.setAttribute('aria-expanded', String(opened));
    toggleBtn.textContent = opened ? '🎯 收合篩選 TA' : '🎯 篩選 TA（EDM 名單）';
    if (opened && !built) {
      panel.textContent = ''; panel.appendChild(el('div', 'cta-loading', '載入選項…'));
      jfetch('GET', '/api/companies/ta/options').then(function (r) { if (!r.ok) throw new Error('載入失敗 (' + r.status + ')'); return r.json(); }).then(function (o) {
        opts = o; F = freshF(); O = freshO(); built = true; build(); query();
      }).catch(function (e) { panel.textContent = ''; panel.appendChild(el('div', 'cta-loading', '⚠ ' + e.message)); });
    }
  }

  // 列表重新載入後呼叫：換資料就重抓選項（下次展開時）
  function mount() {
    host = document.getElementById('cmTaMount');
    if (!host) return;
    built = false; opts = null;
    host.textContent = '';
    toggleBtn = btn('🎯 篩選 TA（EDM 名單）', 'cta-toggle', open);
    toggleBtn.setAttribute('aria-expanded', 'false');
    panel = el('div', 'cta-panel'); panel.hidden = true; opened = false;
    host.appendChild(toggleBtn); host.appendChild(panel);
  }

  window.ITTSCompanyTa = { mount: mount };
})();
