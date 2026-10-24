/* License／MA 成交納管：成交期間視窗、簽約確認任務卡、到期管理頁（使用者端，P2b）
 *
 * 後端（MA_TRACKING=1 才有，關閉時全部 404 → 這裡什麼都不建立、不顯示）：
 *   GET /api/my-tasks/deals、POST /api/deal-tasks/:id/sign|skip|snooze|reassign、GET /api/manager/deal-task-overdue、
 *   GET /api/ma-items?filter=、GET /api/ma-items/pending-backfill、POST /api/ma-items/backfill、PATCH /api/ma-items/:id/term。
 *   商品類型判斷用 GET /api/product-catalog（termKind；比對規則與 lib/productCatalog.js resolveTermRule 相同）。
 *
 * 畫面（全部「啟用後才由本檔動態建立」，所以開關關閉時 index.html 渲染出的 DOM 與沒有此功能時逐位元相同）：
 *   - 成交期間視窗：三個成交入口（看板拖曳／商機編輯表單／表格階段）共用 wonTerm()；取消＝不改階段
 *   - 首頁待辦區卡片（掛在既有 [data-cct-host] 旁）、登入彙總視窗（每次登入有未延後的開啟中任務跳一次，可略過）
 *   - 任務卡視窗：步驟條、合約已簽訂（編號必填）、此案不簽合約（原因必填）、稍後提醒、改派（管理員）
 *   - 一級主管：團隊「未簽超過 N 天」逾期清單
 *   - 側欄「MA／License 到期管理」（GET /api/ma-items 回 200 才出現）與合約管理下的「Service MA」子項
 * 安全：全部動態內容 createElement＋textContent（不經 innerHTML）、不使用 inline style／on* 屬性；樣式在 style.css 末尾的 .mat-*。
 * 對外：window.ITTSMaTasks = { start, enabled, wonTerm, onSection, openFromNotification, openItemsFromNotification, refresh, ... }
 */
(function () {
  'use strict';
  if (window.ITTSMaTasks) return;

  var STEPS = [
    { key: 'won', label: '成交' },
    { key: 'sign', label: '簽約確認' },
    { key: 'contract', label: '歸入 Service MA 管理' },
    { key: 'remind', label: '到期前提醒' },
  ];
  var KIND_LABEL = { ma: 'MA', license: 'License' };
  var SUMMARY_MAX = 10;
  var PANEL_MAX = 5;
  var ITEM_STATUS = { ended: '已結束', renewed: '已續約', not_renew: '不續約' };
  var FILTERS = [
    { key: 'all', label: '全部' },
    { key: 'pending_sign', label: '待簽約' },
    { key: 'd90', label: '90 天內到期' },
    { key: 'overdue', label: '已逾期' },
  ];

  var state = {
    user: null, enabled: false, tasks: null, overdue: null, seq: 0, loginShown: false, summary: null, expanded: false,
    catalog: null, catalogAt: 0, itemsOn: false, pendingSection: '', filter: 'all', itemsAll: null, pending: null,
  };
  var stack = [];
  var uidN = 0;

  /* ───────── 小工具 ───────── */
  function uid(p) { uidN++; return 'mat-' + p + '-' + uidN; }
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
  function jreq(method, url, body) {
    var opt = { method: method, credentials: 'same-origin', cache: 'no-store' };
    if (body !== undefined) { opt.headers = { 'Content-Type': 'application/json' }; opt.body = JSON.stringify(body); }
    return fetch(url, opt)
      .then(function (r) { return r.json().catch(function () { return {}; }).then(function (d) { return { ok: r.ok, status: r.status, data: d }; }); })
      .catch(function () { return { ok: false, status: 0, data: { error: '網路錯誤，請稍後再試' } }; });
  }
  function jget(url) { return jreq('GET', url).then(function (r) { return r.ok ? r.data : null; }); }
  function today() {
    try { return new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Taipei' }); } catch (e) { return new Date().toISOString().slice(0, 10); }
  }
  function isAdmin() { return !!(state.user && state.user.role === 'admin'); }
  function canAct(t) { return !!state.user && (isAdmin() || t.owner === state.user.username); }
  function slash(s) { return s ? String(s).replace(/-/g, '/') : ''; }
  function rangeText(from, to) { return from && to ? slash(from) + ' – ' + slash(to) : ''; }
  function daysText(n) {
    if (typeof n !== 'number') return '';
    if (n < 0) return '已逾期 ' + (-n) + ' 天';
    if (n === 0) return '今天到期';
    return '剩 ' + n + ' 天';
  }
  function refreshBell() {
    try { if (typeof window._refreshBell === 'function') window._refreshBell(); } catch (e) { /* ignore */ }
  }
  function termsOf(t) {
    var its = (t && t.items) || [];
    if (!its.length) return '';
    var seen = {}, out = [];
    its.forEach(function (i) {
      var s = rangeText(i.from, i.to);
      if (s && !seen[s]) { seen[s] = 1; out.push(s); }
    });
    return out.join('、');
  }

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
    var ov = el('div', 'mat-overlay');
    var dlg = el('div', 'mat-dialog' + (o.wide ? ' mat-wide' : ''));
    dlg.setAttribute('role', o.role || 'dialog');
    dlg.setAttribute('aria-modal', 'true');
    var h = el('h2', 'mat-dlg-title', o.title);
    h.id = uid('t');
    dlg.setAttribute('aria-labelledby', h.id);
    var body = el('div', 'mat-dlg-body');
    body.id = uid('d');
    dlg.setAttribute('aria-describedby', body.id);
    var foot = el('div', 'mat-dlg-foot');
    dlg.appendChild(h); dlg.appendChild(body); dlg.appendChild(foot);
    ov.appendChild(dlg);
    var entry = {
      ov: ov, dlg: dlg, body: body, foot: foot, title: h, onEsc: o.onEsc || null, closed: false,
      close: function () {
        if (entry.closed) return;
        entry.closed = true;
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
    var t = (sel && entry.dlg.querySelector(sel)) || entry.dlg.querySelector('input:not([disabled]),textarea,select,button:not([disabled])');
    if (t) { try { t.focus(); } catch (e) { /* ignore */ } }
  }
  function errBox(d) {
    var err = el('div', 'mat-error'); err.setAttribute('role', 'alert'); err.hidden = true;
    d.body.appendChild(err);
    return function (m) { if (m) { err.textContent = m; err.hidden = false; } else { err.textContent = ''; err.hidden = true; } };
  }
  function field(d, labelText, input) {
    var lab = el('label', 'mat-field-label', labelText);
    input.id = input.id || uid('f');
    lab.setAttribute('for', input.id);
    d.appendChild(lab); d.appendChild(input);
    return input;
  }

  /* ───────── 商品目錄：判斷是否為 License／MA 商品（與 lib/productCatalog.js resolveTermRule 同規則）───────── */
  function getCatalog() {
    if (state.catalog && Date.now() - state.catalogAt < 60000) return Promise.resolve(state.catalog);
    return jget('/api/product-catalog').then(function (c) {
      if (c && c.bus) { state.catalog = c; state.catalogAt = Date.now(); }
      return state.catalog;
    });
  }
  function resolveTermKind(cat, product) {
    var none = { termKind: 'none', startsTask: false };
    var p = typeof product === 'string' ? product.trim() : '';
    if (!p || !cat || !cat.bus) return none;
    var best = null;
    Object.keys(cat.bus).forEach(function (bu) {
      (Array.isArray(cat.bus[bu]) ? cat.bus[bu] : []).forEach(function (g) {
        ((g && g.items) || []).forEach(function (it) {
          if (!it || typeof it.name !== 'string' || !it.name) return;
          var exact = p === it.name;
          if (!exact && p.indexOf(it.name + ' ') !== 0) return;
          var rank = exact ? Infinity : it.name.length;
          if (!best || rank > best.rank) best = { rank: rank, items: [it] };
          else if (rank === best.rank) best.items.push(it);
        });
      });
    });
    if (!best) return none;
    var typed = best.items.filter(function (it) { return it.termKind === 'ma' || it.termKind === 'license'; });
    if (!typed.length) return none;
    var pick = typed.filter(function (it) { return it.termKind === 'ma'; })[0] || typed[0];
    return { termKind: pick.termKind, startsTask: pick.startsTask === true };
  }

  /* ───────── 服務期間視窗（成交／補登／修改期間共用）─────────
   * o: { title, lead, kind, from, to, allowBlank, blankNote, showReason, confirmLabel, message, submit(vals) → Promise<{ok, message}>, onClose }
   * vals: { from, to, reason }（allowBlank 且兩欄皆空時 from／to 為 ''）*/
  function termDialog(o) {
    var Q = window.QTERMS;
    var d = openDialog({ title: o.title, role: 'dialog', onClose: o.onClose });
    d.body.appendChild(el('p', 'mat-lead', o.lead || ''));
    if (o.note) d.body.appendChild(el('p', 'mat-note', o.note));
    var row = el('div', 'mat-daterow');
    var fromI = document.createElement('input'); fromI.type = 'date'; fromI.className = 'mat-input'; fromI.value = o.from || '';
    var toI = document.createElement('input'); toI.type = 'date'; toI.className = 'mat-input'; toI.value = o.to || '';
    var c1 = el('div', 'mat-col'), c2 = el('div', 'mat-col');
    field(c1, '服務起日', fromI); field(c2, '服務迄日', toI);
    row.appendChild(c1); row.appendChild(c2);
    d.body.appendChild(row);
    var quick = el('div', 'mat-quick');
    var spanEl = el('span', 'mat-span');
    spanEl.setAttribute('aria-live', 'polite');
    var showErr = errBox(d);
    function updSpan() {
      var t = (Q && Q.spanText) ? Q.spanText(fromI.value, toI.value) : '';
      spanEl.textContent = t;
    }
    [[1, '一年'], [3, '三年']].forEach(function (q) {
      quick.appendChild(button(q[1], 'mat-btn mat-btn-ghost mat-btn-sm', function () {
        if (!fromI.value) { showErr('請先填寫起日，再按「' + q[1] + '」快捷'); fromI.focus(); return; }
        var e = (Q && Q.quickEnd) ? Q.quickEnd(fromI.value, q[0]) : '';
        if (!e) { showErr('起日不是有效日期'); return; }
        toI.value = e; showErr(''); updSpan();
      }));
    });
    quick.appendChild(spanEl);
    d.body.insertBefore(quick, d.body.querySelector('.mat-error'));
    fromI.addEventListener('input', updSpan); toI.addEventListener('input', updSpan);
    var reasonI = null;
    if (o.showReason) {
      reasonI = document.createElement('textarea'); reasonI.className = 'mat-textarea'; reasonI.rows = 2; reasonI.maxLength = 200;
      var holder = el('div', 'mat-reason');
      field(holder, '修改原因（選填，會記在稽核紀錄與這個項目的異動紀錄）', reasonI);
      d.body.insertBefore(holder, d.body.querySelector('.mat-error'));
    }
    updSpan();
    if (o.message) showErr(o.message);
    var ok = button(o.confirmLabel || '確認', 'mat-btn mat-btn-primary');
    d.foot.appendChild(button('取消', 'mat-btn mat-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      var f = fromI.value.trim(), t = toI.value.trim();
      if (o.allowBlank && !f && !t) { /* 空白＝沿用既有來源 */ }
      else {
        var bad = Q && Q.validateTerm ? Q.validateTerm(o.kind || 'ma', f, t) : ((!f || !t) ? { reason: '請填寫服務期間的起日與迄日' } : null);
        if (bad) { showErr(bad.reason); (f ? toI : fromI).focus(); return; }
      }
      showErr('');
      ok.disabled = true;
      Promise.resolve(o.submit({ from: f, to: t, reason: reasonI ? reasonI.value.trim() : '' })).then(function (r) {
        if (r && r.ok === false) { ok.disabled = false; showErr(r.message || '失敗，請再試一次'); return; }
        d.close();
      });
    });
    focusFirst(d, 'input[type=date]');
    return d;
  }

  /* ───────── 成交入口共用：需要期間時先跳視窗 ─────────
   * 回傳 Promise：{} ＝不需要（照舊送出）；{ termFrom, termTo } ＝把這兩欄帶入 PUT；null ＝使用者取消（不改階段）。
   * opts: { product, force（伺服器已回 TERM_*，一定要顯示）, message（伺服器訊息）, prefill:{termFrom,termTo} } */
  function loadQuoteTerms(opp) {
    // 找連結到這筆商機、已核准的最近報價單的期間（只用來預填；伺服器仍是最終裁判）
    if (!opp || !opp.id) return Promise.resolve(null);
    return jget('/api/opportunities/' + encodeURIComponent(opp.id) + '/quotes').then(function (r) {
      var qs = ((r && r.quotes) || []).filter(function (q) { return q.approvalState === 'approved' && !q.invalid; });
      var pickN = qs.slice(0, 3);
      var i = 0;
      function next() {
        if (i >= pickN.length) return Promise.resolve(null);
        var q = pickN[i++];
        return jget('/api/quotations/' + encodeURIComponent(q.id)).then(function (d) {
          var groups = {}, order = [];
          ((d && d.items) || []).forEach(function (it) {
            if (it && it.termFrom && it.termTo) {
              var k = it.termFrom + '|' + it.termTo;
              if (!groups[k]) { groups[k] = { from: it.termFrom, to: it.termTo }; order.push(k); }
            }
          });
          if (!order.length) return next();
          return { quoteNo: q.quoteNo || '', groups: order.map(function (k) { return groups[k]; }) };
        });
      }
      return next();
    });
  }
  function wonTerm(opp, opts) {
    opts = opts || {};
    return new Promise(function (resolve) {
      if (!opts.force) {
        if (!state.enabled) return resolve({});
        if (opp && opp.stage === 'Won') return resolve({});
      }
      var product = opts.product ? opts.product : (opp && opp.product);
      getCatalog().then(function (cat) {
        var rule = resolveTermKind(cat, product);
        if (!opts.force && rule.termKind === 'none') return resolve({});
        loadQuoteTerms(opp).then(function (qt) {
          var from = '', to = '', note = '', multi = false;
          if (opts.prefill && opts.prefill.termFrom) { from = opts.prefill.termFrom; to = opts.prefill.termTo || ''; }
          else if (qt && qt.groups.length === 1) { from = qt.groups[0].from; to = qt.groups[0].to; }
          if (qt && qt.groups.length === 1 && !(opts.prefill && opts.prefill.termFrom)) {
            note = '已帶入核准報價單 ' + qt.quoteNo + ' 的服務期間，可以調整。';
          } else if (qt && qt.groups.length > 1) {
            multi = true;
            note = '核准報價單 ' + qt.quoteNo + ' 有 ' + qt.groups.length + ' 組不同的服務期間（' + qt.groups.map(function (g) { return rangeText(g.from, g.to); }).join('、') + '）。兩欄都留空＝依報價單各組期間納管；填寫＝全部改用這一組期間。';
          }
          var result = null;
          var who = (opp && opp.company) ? '「' + opp.company + '」' : '這筆商機';
          termDialog({
            title: '成交：請填寫服務期間',
            lead: who + '的商品「' + (product || '') + '」是 ' + (KIND_LABEL[rule.termKind] || 'License／MA') + ' 類，請填寫服務期間（起訖日）。成交後會自動納入「MA／License 到期管理」。',
            note: note,
            kind: rule.termKind === 'none' ? 'ma' : rule.termKind,
            from: from, to: to, allowBlank: multi, message: opts.message || '', confirmLabel: '確認成交',
            submit: function (v) { result = (v.from && v.to) ? { termFrom: v.from, termTo: v.to } : {}; return Promise.resolve({ ok: true }); },
            onClose: function () { resolve(result); },
          });
        });
      });
    });
  }

  /* ───────── 任務卡視窗 ───────── */
  function stepStates(t) {
    // 回傳 4 個關卡的狀態：done／current／todo／skipped
    if (t.status === 'done') return ['done', 'done', 'done', 'todo'];
    if (t.status === 'skipped') return ['done', 'skipped', 'todo', 'todo'];
    if (t.status === 'cancelled') return ['todo', 'todo', 'todo', 'todo'];
    return ['done', 'current', 'todo', 'todo'];
  }
  var STEP_MARK = { done: '✓ 完成', current: '進行中', todo: '待辦', skipped: '不簽約' };
  function stepBar(t) {
    var ol = el('ol', 'mat-steps');
    ol.setAttribute('aria-label', '處理進度');
    var st = stepStates(t);
    STEPS.forEach(function (s, i) {
      var li = el('li', 'mat-step mat-step-' + st[i]);
      li.setAttribute('data-state', st[i]);
      li.appendChild(el('span', 'mat-step-n', i + 1));
      li.appendChild(el('span', 'mat-step-l', s.label));
      li.appendChild(el('span', 'mat-step-s', STEP_MARK[st[i]]));
      ol.appendChild(li);
    });
    return ol;
  }
  function dl(pairs) {
    var d = el('dl', 'mat-dl');
    pairs.forEach(function (p) {
      if (!p[1]) return;
      d.appendChild(el('dt', 'mat-dt', p[0]));
      d.appendChild(el('dd', 'mat-dd', p[1]));
    });
    return d;
  }
  function afterAction() {
    return Promise.all([refresh(), state.itemsOn ? loadItems(true) : Promise.resolve()]).then(function () { refreshBell(); });
  }

  function openSign(t, parent) {
    var d = openDialog({ title: '合約已簽訂：輸入合約編號', role: 'dialog' });
    d.body.appendChild(el('p', 'mat-lead', '「' + t.company + '」的合約編號必須填寫；編號不可與其他合約重複。'));
    var no = document.createElement('input'); no.type = 'text'; no.className = 'mat-input'; no.maxLength = 64; no.autocomplete = 'off';
    var sd = document.createElement('input'); sd.type = 'date'; sd.className = 'mat-input'; sd.value = today();
    field(d.body, '合約編號（必填）', no);
    field(d.body, '簽約日', sd);
    var showErr = errBox(d);
    var ok = button('確認簽約', 'mat-btn mat-btn-primary');
    d.foot.appendChild(button('取消', 'mat-btn mat-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      var v = no.value.trim();
      if (!v) { showErr('請輸入合約編號'); no.focus(); return; }
      if (!sd.value) { showErr('請選擇簽約日'); sd.focus(); return; }
      showErr(''); ok.disabled = true;
      jreq('POST', '/api/deal-tasks/' + encodeURIComponent(t.id) + '/sign', { contractNo: v, signedDate: sd.value }).then(function (r) {
        if (!r.ok) { ok.disabled = false; showErr((r.data && r.data.error) || '簽約確認失敗'); no.focus(); return; }
        d.close();
        if (parent) parent.close();
        afterAction().then(function () { openSigned(t, r.data && r.data.contract); });
      });
    });
    focusFirst(d, 'input[type=text]');
  }
  function openSigned(t, c) {
    var d = openDialog({ title: '已歸入合約管理 → Service MA', role: 'dialog' });
    d.body.appendChild(el('p', 'mat-lead', '「' + t.company + '」的合約已建立' + (c && c.contractNo ? '（編號 ' + c.contractNo + '）' : '') + '，並歸入「合約管理 → Service MA」。到期日會列在「MA／License 到期管理」。'));
    d.body.appendChild(stepBar({ status: 'done' }));
    d.foot.appendChild(button('關閉', 'mat-btn mat-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(button('前往 Service MA', 'mat-btn mat-btn-primary', function () {
      d.close();
      if (typeof window.showSection === 'function') window.showSection('service-ma');
    }));
    focusFirst(d, '.mat-btn-primary');
  }
  function openSkip(t, parent) {
    var d = openDialog({ title: '此案不簽合約', role: 'dialog' });
    d.body.appendChild(el('p', 'mat-lead', '請寫明原因。系統會通知行銷與你的一級主管；這筆服務仍會繼續追蹤到期日。'));
    var ta = document.createElement('textarea'); ta.className = 'mat-textarea'; ta.rows = 3; ta.maxLength = 200;
    field(d.body, '原因（必填，至少 2 個字）', ta);
    var showErr = errBox(d);
    var ok = button('送出', 'mat-btn mat-btn-danger');
    d.foot.appendChild(button('取消', 'mat-btn mat-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    ok.addEventListener('click', function () {
      var v = ta.value.trim();
      if (v.length < 2) { showErr('請填寫不簽合約的原因（至少 2 個字）'); ta.focus(); return; }
      showErr(''); ok.disabled = true;
      jreq('POST', '/api/deal-tasks/' + encodeURIComponent(t.id) + '/skip', { reason: v }).then(function (r) {
        if (!r.ok) { ok.disabled = false; showErr((r.data && r.data.error) || '送出失敗'); return; }
        d.close();
        if (parent) parent.close();
        toast('已記錄：此案不簽合約，已通知行銷與主管');
        afterAction();
      });
    });
    focusFirst(d, 'textarea');
  }
  function snooze(t, parent) {
    jreq('POST', '/api/deal-tasks/' + encodeURIComponent(t.id) + '/snooze').then(function (r) {
      if (!r.ok) { toast((r.data && r.data.error) || '延後失敗'); return; }
      if (parent) parent.close();
      toast('好的，稍後再提醒你' + (r.data && r.data.snoozedUntil ? '（' + slash(r.data.snoozedUntil) + ' 之後）' : ''));
      afterAction();
    });
  }
  function openReassign(t, parent) {
    var d = openDialog({ title: '改派「' + t.company + '」的簽約確認', role: 'dialog' });
    d.body.appendChild(el('p', 'mat-lead', '目前負責人：' + (t.ownerLabel || t.owner) + (t.ownerActive === false ? '（已停用）' : '') + '。改派後新的負責人會收到通知。'));
    var holder = el('div', 'mat-reassign');
    d.body.appendChild(holder);
    var sel = document.createElement('select'); sel.className = 'mat-input';
    var txt = null;
    var showErr = errBox(d);
    var ok = button('改派', 'mat-btn mat-btn-primary');
    ok.disabled = true;
    d.foot.appendChild(button('取消', 'mat-btn mat-btn-ghost', function () { d.close(); }));
    d.foot.appendChild(ok);
    jget('/api/company-responsible/candidates').then(function (r) {
      var users = ((r && r.users) || []).filter(function (u) { return u.username !== t.owner; });
      if (users.length) {
        users.forEach(function (u) {
          var o = document.createElement('option'); o.value = u.username;
          o.textContent = u.label + ((u.bu && u.bu.length) ? '（' + u.bu.join('/') + '）' : '');
          sel.appendChild(o);
        });
        field(holder, '新的負責業務（在職人員）', sel);
      } else {
        txt = document.createElement('input'); txt.type = 'text'; txt.className = 'mat-input'; txt.maxLength = 64;
        field(holder, '新的負責業務帳號', txt);
      }
      ok.disabled = false;
      focusFirst(d, 'select,input');
    });
    ok.addEventListener('click', function () {
      var to = txt ? txt.value.trim() : sel.value;
      if (!to) { showErr('請選擇新的負責業務'); return; }
      showErr(''); ok.disabled = true;
      jreq('POST', '/api/deal-tasks/' + encodeURIComponent(t.id) + '/reassign', { to: to }).then(function (r) {
        if (!r.ok) { ok.disabled = false; showErr((r.data && r.data.error) || '改派失敗'); return; }
        d.close();
        if (parent) parent.close();
        toast('✓ 已改派');
        afterAction();
      });
    });
  }
  function openTask(t) {
    var d = openDialog({ title: '成交後確認簽約：' + t.company, role: 'dialog', wide: true });
    d.body.appendChild(el('p', 'mat-lead', '這筆 License／MA 商機已成交。請確認客戶是否已簽訂合約；簽訂後輸入合約編號，就會歸入「合約管理 → Service MA」。'));
    d.body.appendChild(stepBar(t));
    d.body.appendChild(dl([
      ['公司', t.company],
      ['品項', t.product],
      ['服務期間', termsOf(t)],
      ['負責業務', (t.ownerLabel || t.owner) + (t.ownerActive === false ? '（已停用）' : '')],
      ['已等待', typeof t.daysOpen === 'number' ? (t.daysOpen === 0 ? '今天成交' : t.daysOpen + ' 天') : ''],
    ]));
    if (t.snoozed && t.snoozedUntil) d.body.appendChild(el('p', 'mat-note', '你已選擇稍後提醒（' + slash(t.snoozedUntil) + ' 之後再提醒）。'));
    var open = t.status === 'open';
    if (!open) d.body.appendChild(el('p', 'mat-note', t.status === 'done' ? '這筆已完成簽約確認。' : (t.status === 'skipped' ? '這筆已標記為不簽合約。' : '這筆任務已取消。')));
    else if (!canAct(t)) d.body.appendChild(el('p', 'mat-note', '只有負責業務或管理員可以處理這張任務卡。'));
    d.foot.appendChild(button('關閉', 'mat-btn mat-btn-ghost', function () { d.close(); }));
    if (open && isAdmin()) d.foot.appendChild(button('改派…', 'mat-btn mat-btn-ghost', function () { openReassign(t, d); }));
    if (open && canAct(t)) {
      d.foot.appendChild(button('稍後提醒', 'mat-btn mat-btn-ghost', function () { snooze(t, d); }));
      d.foot.appendChild(button('此案不簽合約…', 'mat-btn mat-btn-danger-ghost', function () { openSkip(t, d); }));
      d.foot.appendChild(button('合約已簽訂，輸入編號', 'mat-btn mat-btn-primary', function () { openSign(t, d); }));
    }
    focusFirst(d, '.mat-btn-primary');
    return d;
  }

  /* ───────── 待辦列、彙總視窗、首頁卡片 ───────── */
  function taskRow(t) {
    var li = el('li', 'mat-row' + (t.snoozed ? ' mat-row-snoozed' : ''));
    var main = el('div', 'mat-row-main');
    main.appendChild(el('strong', 'mat-name', t.company));
    var meta = [t.product, termsOf(t)].filter(Boolean).join('　');
    main.appendChild(el('span', 'mat-meta', meta));
    var chips = el('span', 'mat-chips-inline');
    chips.appendChild(el('span', 'mat-chip', '待簽約確認'));
    if (typeof t.daysOpen === 'number') chips.appendChild(el('span', 'mat-chip', t.daysOpen === 0 ? '今天成交' : '成交 ' + t.daysOpen + ' 天'));
    if (t.overdue) chips.appendChild(el('span', 'mat-chip mat-chip-warn', '已逾 ' + ((state.tasks && state.tasks.overdueDays) || 14) + ' 天'));
    if (t.snoozed) chips.appendChild(el('span', 'mat-chip mat-chip-mute', '已延後'));
    main.appendChild(chips);
    li.appendChild(main);
    var act = el('div', 'mat-row-actions');
    var b = button('處理…', 'mat-btn mat-btn-primary mat-btn-sm', function () { openTask(t); });
    b.setAttribute('aria-label', '處理「' + t.company + '」的簽約確認');
    act.appendChild(b);
    li.appendChild(act);
    return li;
  }
  function panel(titleText, count, cls) {
    var sec = el('section', 'mat-panel ' + (cls || ''));
    var hid = uid('h');
    sec.setAttribute('aria-labelledby', hid);
    var hd = el('div', 'mat-panel-hd');
    var h = el('h3', 'mat-panel-title', titleText);
    h.id = hid;
    hd.appendChild(h);
    if (count != null) hd.appendChild(el('span', 'mat-count', count));
    sec.appendChild(hd);
    return sec;
  }
  function myTasksPanel() {
    var tasks = (state.tasks && state.tasks.items) || [];
    if (!tasks.length) return null;
    var sec = panel('📝 待辦事項：確認簽約', tasks.length, 'mat-panel-tasks');
    sec.appendChild(el('p', 'mat-sub', 'License／MA 商機已成交。請確認客戶是否已簽約，並輸入合約編號；完成後會歸入「合約管理 → Service MA」。'));
    var ul = el('ul', 'mat-list');
    var show = state.expanded ? tasks : tasks.slice(0, PANEL_MAX);
    show.forEach(function (t) { ul.appendChild(taskRow(t)); });
    sec.appendChild(ul);
    if (tasks.length > PANEL_MAX) {
      sec.appendChild(button(state.expanded ? '收合' : '顯示全部 ' + tasks.length + ' 筆', 'mat-btn mat-btn-link', function () { state.expanded = !state.expanded; renderHosts(); }));
    }
    return sec;
  }
  function overduePanel() {
    var o = state.overdue;
    if (!o || !o.items || !o.items.length) return null;
    var sec = panel('⏰ 成交後超過 ' + (o.overdueDays || 14) + ' 天未簽約確認', o.items.length, 'mat-panel-overdue');
    var ul = el('ul', 'mat-list');
    o.items.slice(0, 20).forEach(function (t) {
      var li = el('li', 'mat-row');
      var main = el('div', 'mat-row-main');
      main.appendChild(el('strong', 'mat-name', t.company));
      main.appendChild(el('span', 'mat-meta', '負責業務 ' + (t.ownerLabel || t.owner) + (t.ownerActive === false ? '（已停用）' : '') + '　已等待 ' + t.daysOpen + ' 天'));
      if (t.product) main.appendChild(el('span', 'mat-meta', t.product));
      li.appendChild(main);
      if (o.canReassign) {
        var act = el('div', 'mat-row-actions');
        var b = button('改派…', 'mat-btn mat-btn-ghost mat-btn-sm', function () { openReassign(t, null); });
        b.setAttribute('aria-label', '改派「' + t.company + '」的簽約確認');
        act.appendChild(b);
        li.appendChild(act);
      }
      ul.appendChild(li);
    });
    sec.appendChild(ul);
    if (o.items.length > 20) sec.appendChild(el('p', 'mat-more', '另有 ' + (o.items.length - 20) + ' 筆未列出。'));
    return sec;
  }
  function ensureHosts() {
    var hs = document.querySelectorAll('[data-cct-host]');
    for (var i = 0; i < hs.length; i++) {
      var h = hs[i];
      if (h.nextElementSibling && h.nextElementSibling.hasAttribute('data-mat-host')) continue;
      var m = el('div', 'mat-host');
      m.setAttribute('data-mat-host', h.getAttribute('data-cct-host'));
      m.hidden = true;
      h.parentNode.insertBefore(m, h.nextSibling);
    }
  }
  function renderHosts() {
    if (!state.enabled) return;
    ensureHosts();
    var hosts = document.querySelectorAll('[data-mat-host]');
    for (var i = 0; i < hosts.length; i++) {
      var h = hosts[i];
      var kind = h.getAttribute('data-mat-host');
      h.textContent = '';
      var parts = [];
      if (kind === 'mgr') parts.push(overduePanel());
      parts.push(myTasksPanel());
      parts.forEach(function (p) { if (p) h.appendChild(p); });
      h.hidden = !h.firstChild;
    }
    if (state.summary && state.summary.rerender) state.summary.rerender();
  }
  function fillSummary(d) {
    var tasks = (state.tasks && state.tasks.items) || [];
    d.body.textContent = '';
    d.body.appendChild(el('p', 'mat-lead', '以下 License／MA 商機已成交，請確認客戶是否已簽訂合約。簽訂後輸入合約編號，就會歸入「合約管理 → Service MA」。不簽合約也請說明原因。'));
    var ul = el('ul', 'mat-list');
    tasks.slice(0, SUMMARY_MAX).forEach(function (t) { ul.appendChild(taskRow(t)); });
    d.body.appendChild(ul);
    if (tasks.length > SUMMARY_MAX) d.body.appendChild(el('p', 'mat-more', '還有 ' + (tasks.length - SUMMARY_MAX) + ' 筆未列出，可在首頁「待辦事項」查看全部。'));
    d.title.textContent = '你有 ' + tasks.length + ' 筆成交案待確認簽約';
  }
  function openSummary(opts) {
    if (state.summary) return Promise.resolve();
    var go = (opts && opts.fetch) ? refresh() : Promise.resolve();
    return go.then(function () {
      var tasks = (state.tasks && state.tasks.items) || [];
      if (!tasks.length) { if (opts && opts.fetch) toast('目前沒有待確認簽約的成交案'); return; }
      if (state.summary) return;
      var d = openDialog({ title: '你有 ' + tasks.length + ' 筆成交案待確認簽約', role: 'alertdialog', wide: true, onClose: function () { state.summary = null; } });
      state.summary = d;
      d.onEsc = function () { d.close(); };   // Esc＝略過（稍後處理）
      fillSummary(d);
      d.foot.appendChild(button('略過，稍後處理', 'mat-btn mat-btn-ghost mat-later', function () { d.close(); }));
      d.rerender = function () {
        var left = (state.tasks && state.tasks.items) || [];
        if (!left.length) { d.close(); return; }
        fillSummary(d);
      };
      focusFirst(d, '.mat-btn-primary');
    });
  }

  /* ───────── 載入與啟動 ───────── */
  function refresh() {
    if (!state.user || state.user.role === 'tecopm') return Promise.resolve(state);
    var seq = ++state.seq;
    var role = state.user.role;
    var jobs = [
      jget('/api/my-tasks/deals'),
      (['admin', 'executive', 'manager1', 'manager2'].indexOf(role) >= 0) ? jget('/api/manager/deal-task-overdue') : Promise.resolve(null),
    ];
    return Promise.all(jobs).then(function (r) {
      if (seq !== state.seq) return state;   // 較新的請求已發出 → 丟棄過期結果
      state.enabled = !!(r[0] && r[0].enabled);
      state.tasks = state.enabled ? r[0] : null;
      state.overdue = (state.enabled && r[1] && r[1].enabled) ? r[1] : null;
      renderHosts();
      return state;
    });
  }
  var refreshTimer = null;
  function refreshSoon() { clearTimeout(refreshTimer); refreshTimer = setTimeout(function () { refresh(); }, 400); }

  function otherDialogOpen() { return !!document.querySelector('.ann-overlay, .cct-overlay, .creq-overlay'); }
  function whenClear(fn, tries) {
    if (!otherDialogOpen() || tries <= 0) { fn(); return; }
    setTimeout(function () { whenClear(fn, tries - 1); }, 600);
  }

  function start(user, opts) {
    state.user = user || state.user;
    if (!state.user || state.user.role === 'tecopm') return Promise.resolve();
    return refresh().then(function () {
      if (!state.enabled) { settlePending(); return; }
      return setupItems().then(function () {
        var tasks = (state.tasks && state.tasks.items) || [];
        var live = tasks.filter(function (t) { return !t.snoozed; });
        if (opts && opts.login && !state.loginShown && live.length) {
          state.loginShown = true;
          // 其他彈窗（公告、待補名片、建檔申請）可能剛好在載入：等它們關掉再顯示，視窗不會疊在一起
          setTimeout(function () { whenClear(function () { openSummary(); }, 100); }, 1800);
        }
        checkHash();
        settlePending();
      });
    });
  }
  function settlePending() {
    // F5 還原到本檔的區段（maItems／service-ma）時，啟動前只能先記下來；啟用後再補顯示，沒啟用就回首頁
    var p = state.pendingSection;
    if (!p) return;
    state.pendingSection = '';
    if (state.itemsOn && typeof window.showSection === 'function') window.showSection(p);
    else if (typeof window.showDashboard === 'function') window.showDashboard();
  }
  function checkHash() {
    var h = window.location.hash;
    if (h === '#dealtasks' || h === '#maitems') {
      try { history.replaceState(null, '', window.location.pathname + window.location.search); } catch (e) { /* ignore */ }
      if (h === '#dealtasks') openSummary({ fetch: true });
      else if (state.itemsOn && typeof window.showSection === 'function') window.showSection('maItems');
    }
  }
  window.addEventListener('hashchange', function () { if (state.user && state.enabled) checkHash(); });

  /* ───────── 側欄與區段（啟用後才動態建立）───────── */
  function navItem(id, icon, label, sub) {
    var d = el('div', 'sidebar-nav-item' + (sub ? ' sidebar-sub-item' : ''));
    d.id = id;
    d.appendChild(el('span', 'nav-icon', icon));
    d.appendChild(el('span', 'nav-label', label));
    d.setAttribute('role', 'button'); d.tabIndex = 0;
    return d;
  }
  function bindNav(n, section) {
    n.addEventListener('click', function () {
      if (typeof window.showSection === 'function') window.showSection(section);
      if (window.innerWidth <= 900) {   // 手機：點選後收起側欄（既有導覽項在載入時綁的，動態新增的要自己做）
        var sb = document.querySelector('.sidebar'); if (sb) sb.classList.remove('mobile-open');
        var ov = document.getElementById('sidebarOverlay'); if (ov) ov.classList.remove('show');
      }
    });
    n.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); n.click(); } });
  }
  function buildServiceMaView(anchor) {
    var v = el('div', 'mat-plain'); v.id = 'serviceMaView'; v.hidden = true;
    var bar = el('div', 'mat-notice');
    bar.appendChild(el('strong', null, 'Service MA 合約'));
    bar.appendChild(el('span', null, '　成交後在任務卡確認簽約而建立的合約會出現在這裡。來源商機與到期管理項不可在此修改。'));
    v.appendChild(bar);
    var stat = el('div', 'contract-stat-row'); stat.id = 'serviceMaStatRow'; v.appendChild(stat);
    var tb = el('div', 'contract-toolbar');
    var sb = el('div', 'search-box'); sb.appendChild(el('span', 'search-icon', '🔍'));
    var si = document.createElement('input'); si.type = 'text'; si.id = 'serviceMaSearch'; si.placeholder = '搜尋客戶、合約號碼...'; si.setAttribute('aria-label', '搜尋 Service MA 合約');
    sb.appendChild(si); tb.appendChild(sb);
    var sel = document.createElement('select'); sel.id = 'serviceMaStatusFilter'; sel.className = 'forecast-year-sel'; sel.setAttribute('aria-label', '狀態篩選');
    [['', '全部狀態'], ['active', '有效'], ['expiring', '即將到期'], ['expired', '逾期']].forEach(function (o) { var op = document.createElement('option'); op.value = o[0]; op.textContent = o[1]; sel.appendChild(op); });
    tb.appendChild(sel); v.appendChild(tb);
    var wrap = el('div', 'contract-table-wrap');
    var table = el('table', 'contract-table');
    var thead = document.createElement('thead'); var tr = document.createElement('tr');
    ['合約號碼', '客戶名稱', '產品／服務', '合約起日', '合約迄日', '年費金額 (NT$K)', '業務人員', '狀態', '備註', '操作'].forEach(function (h) { tr.appendChild(el('th', null, h)); });
    thead.appendChild(tr); table.appendChild(thead);
    var tbody = document.createElement('tbody'); tbody.id = 'serviceMaTbody'; table.appendChild(tbody);
    wrap.appendChild(table); v.appendChild(wrap);
    anchor.parentNode.insertBefore(v, anchor.nextSibling);
    var rerender = function () { if (typeof window.renderServiceMaList === 'function') window.renderServiceMaList(); };
    si.addEventListener('input', rerender); sel.addEventListener('change', rerender);
    return v;
  }
  function setupItems() {
    if (state.itemsOn) return Promise.resolve();
    if (state.user.role === 'tecopm') return Promise.resolve();
    return jreq('GET', '/api/ma-items?filter=all').then(function (r) {
      if (!r.ok || !r.data || r.data.enabled === false) return;
      state.itemsOn = true;
      state.itemsAll = r.data;
      var sap = document.getElementById('sapMaView');
      var inner = document.querySelector('[data-sg-inner="contract"]');
      var grp = document.getElementById('navGroupContract');
      var sub = document.getElementById('subMenuContract');
      if (inner && !document.getElementById('navMaItems')) {
        var n = navItem('navMaItems', '⏳', 'MA／License 到期管理', false);
        inner.insertBefore(n, grp || inner.firstChild);
        bindNav(n, 'maItems');
      }
      if (sub && !document.getElementById('navServiceMa')) {
        var n2 = navItem('navServiceMa', '🧾', 'Service MA', true);
        sub.appendChild(n2);
        bindNav(n2, 'service-ma');
      }
      if (sap) {
        if (!document.getElementById('serviceMaView')) buildServiceMaView(sap);
        if (!document.getElementById('maItemsView')) buildItemsView(sap);
      }
      try { if (window.NavGroups && window.NavGroups.refresh) window.NavGroups.refresh(); } catch (e) { /* ignore */ }
    });
  }

  var SECTION_TITLE = { maItems: 'MA／License 到期管理', 'service-ma': 'Service MA 合約管理' };
  // app.js 的 showSection／showDashboard 結尾呼叫；沒啟用時除了記下 F5 還原的區段外什麼都不做
  function onSection(section) {
    var mine = section === 'maItems' || section === 'service-ma';
    if (!state.itemsOn) { if (mine) state.pendingSection = section; return; }
    var iv = document.getElementById('maItemsView'), sv = document.getElementById('serviceMaView');
    if (iv) iv.hidden = section !== 'maItems';
    if (sv) sv.hidden = section !== 'service-ma';
    var n1 = document.getElementById('navMaItems'), n2 = document.getElementById('navServiceMa');
    if (n1) n1.classList.toggle('active', section === 'maItems');
    if (n2) n2.classList.toggle('active', section === 'service-ma');
    if (mine) {
      var t = document.getElementById('topbarTitle'); if (t) t.textContent = SECTION_TITLE[section];
      if (section === 'service-ma') {
        var sm = document.getElementById('subMenuContract'); if (sm) sm.classList.add('open');
        var ar = document.querySelector('#navContractGroup .nav-arrow'); if (ar) ar.classList.add('open');
        if (typeof window.loadServiceMaView === 'function') window.loadServiceMaView();
      } else {
        loadItems(false);
      }
    }
  }

  /* ───────── 到期管理頁 ───────── */
  var itemsEls = null;
  function isSecretary() { return !!(state.user && state.user.role === 'secretary'); }
  function canEditItem(it) { return !isSecretary() && !!it.id && (isAdmin() || (state.user && it.owner === state.user.username)); }
  function buildItemsView(anchor) {
    var v = el('div', 'mat-page'); v.id = 'maItemsView'; v.hidden = true;
    var head = el('div', 'mat-page-hd');
    head.appendChild(el('h2', 'mat-page-title', 'MA／License 到期管理'));
    var actions = el('div', 'mat-page-actions');
    var backfillBtn = button('待補納管', 'mat-btn mat-btn-ghost mat-btn-sm', function () { openBackfill(); });
    backfillBtn.hidden = true;
    actions.appendChild(backfillBtn);
    head.appendChild(actions);
    v.appendChild(head);
    v.appendChild(el('p', 'mat-sub', isSecretary() ? '這是同 BU 的 MA／License 到期清單（唯讀，不含合約編號與金額）。' : '成交的 License／MA 商機會自動列入這裡，追蹤簽約與到期日。'));
    var cards = el('div', 'mat-cards'); cards.setAttribute('role', 'group'); cards.setAttribute('aria-label', '摘要');
    v.appendChild(cards);
    var chips = el('div', 'mat-filters'); chips.setAttribute('role', 'group'); chips.setAttribute('aria-label', '篩選');
    v.appendChild(chips);
    var wrap = el('div', 'mat-tablewrap');
    var table = el('table', 'mat-table');
    var cap = el('caption', 'mat-sr', 'MA／License 到期管理清單'); table.appendChild(cap);
    var thead = document.createElement('thead'); var tr = document.createElement('tr');
    var cols = isSecretary() ? ['客戶', '品項', '期間', '剩餘天數', '負責業務'] : ['客戶', '品項', '合約編號', '期間', '剩餘天數', '狀態／下一步', '負責業務', '操作'];
    cols.forEach(function (h) { var th = el('th', null, h); th.scope = 'col'; tr.appendChild(th); });
    thead.appendChild(tr); table.appendChild(thead);
    var tbody = document.createElement('tbody'); table.appendChild(tbody);
    wrap.appendChild(table); v.appendChild(wrap);
    var empty = el('p', 'mat-empty'); empty.hidden = true; v.appendChild(empty);
    anchor.parentNode.insertBefore(v, anchor.nextSibling);
    itemsEls = { cards: cards, chips: chips, tbody: tbody, empty: empty, backfill: backfillBtn };
    return v;
  }
  function counts(list) {
    var c = { total: 0, pending: 0, d90: 0, overdue: 0 };
    (list || []).forEach(function (it) {
      if (it.status && it.status !== 'tracking') return;
      c.total++;
      if (it.pendingSign) c.pending++;
      if (typeof it.daysLeft === 'number') { if (it.daysLeft < 0) c.overdue++; else if (it.daysLeft <= 90) c.d90++; }
    });
    return c;
  }
  function renderCards() {
    if (!itemsEls) return;
    var c = counts(state.itemsAll && state.itemsAll.items);
    itemsEls.cards.textContent = '';
    var defs = isSecretary()
      ? [['清單筆數', c.total, ''], ['90 天內到期', c.d90, 'warn'], ['已逾期', c.overdue, 'bad']]
      : [['納管中', c.total, ''], ['待簽約', c.pending, 'warn'], ['90 天內到期', c.d90, 'warn'], ['已逾期', c.overdue, 'bad']];
    defs.forEach(function (d) {
      var card = el('div', 'mat-card' + (d[2] ? ' mat-card-' + d[2] : ''));
      card.appendChild(el('span', 'mat-card-n', d[1]));
      card.appendChild(el('span', 'mat-card-l', d[0]));
      itemsEls.cards.appendChild(card);
    });
  }
  function renderChips() {
    if (!itemsEls) return;
    itemsEls.chips.textContent = '';
    FILTERS.forEach(function (f) {
      if (isSecretary() && f.key === 'pending_sign') return;
      var b = button(f.label, 'mat-fchip' + (state.filter === f.key ? ' mat-fchip-on' : ''), function () { state.filter = f.key; renderChips(); loadItems(false, true); });
      b.setAttribute('aria-pressed', state.filter === f.key ? 'true' : 'false');
      itemsEls.chips.appendChild(b);
    });
  }
  function barFor(n) {
    var p = document.createElement('progress');
    p.className = 'mat-bar ' + (n < 0 || n <= 30 ? 'mat-bar-bad' : (n <= 90 ? 'mat-bar-warn' : 'mat-bar-ok'));
    p.max = 365; p.value = Math.max(0, Math.min(365, n));
    p.setAttribute('aria-label', daysText(n));
    return p;
  }
  function nextStep(it) {
    if (it.status && it.status !== 'tracking') return ITEM_STATUS[it.status] || it.status;
    if (it.contractNo) return '已歸入 Service MA；到期前提醒';
    if (it.taskStatus === 'skipped') return '此案不簽合約（仍追蹤到期）';
    return '待簽約確認';
  }
  function renderRows(list) {
    if (!itemsEls) return;
    var tb = itemsEls.tbody;
    tb.textContent = '';
    itemsEls.empty.hidden = list.length > 0;
    if (!list.length) itemsEls.empty.textContent = '沒有符合條件的項目。';
    list.forEach(function (it) {
      var tr = document.createElement('tr');
      if (typeof it.daysLeft === 'number' && it.daysLeft < 0) tr.className = 'mat-tr-bad';
      function td(label, node) { var c = document.createElement('td'); c.setAttribute('data-label', label); if (node) c.appendChild(typeof node === 'string' ? document.createTextNode(node) : node); tr.appendChild(c); return c; }
      td('客戶', el('strong', null, it.company || ''));
      var pc = td('品項'); pc.appendChild(el('span', null, it.product || '')); pc.appendChild(document.createTextNode(' ')); pc.appendChild(el('span', 'mat-chip', KIND_LABEL[it.termKind] || it.termKind || ''));
      if (!isSecretary()) td('合約編號', it.contractNo || '—');
      td('期間', rangeText(it.from, it.to));
      var dc = td('剩餘天數');
      if (typeof it.daysLeft === 'number') { dc.appendChild(barFor(it.daysLeft)); dc.appendChild(el('span', 'mat-days' + (it.daysLeft < 0 ? ' mat-days-bad' : ''), daysText(it.daysLeft))); }
      if (!isSecretary()) td('狀態／下一步', nextStep(it));
      td('負責業務', (it.ownerLabel || it.owner || '') + (it.ownerActive === false ? '（已停用）' : ''));
      if (!isSecretary()) {
        var ac = td('操作'); var box = el('div', 'mat-row-actions');
        if (canEditItem(it) && it.pendingSign && it.taskId && it.taskStatus === 'open') {
          var sb = button('處理簽約…', 'mat-btn mat-btn-primary mat-btn-sm', function () {
            openTask({ id: it.taskId, status: 'open', company: it.company, product: it.product, owner: it.owner, ownerLabel: it.ownerLabel, ownerActive: it.ownerActive, items: [it] });
          });
          sb.setAttribute('aria-label', '處理「' + it.company + '」的簽約確認'); box.appendChild(sb);
        }
        if (canEditItem(it) && (!it.status || it.status === 'tracking')) {
          var eb = button('修改期間…', 'mat-btn mat-btn-ghost mat-btn-sm', function () { openEditTerm(it); });
          eb.setAttribute('aria-label', '修改「' + it.company + '」的服務期間'); box.appendChild(eb);
        }
        ac.appendChild(box);
      }
      tb.appendChild(tr);
    });
  }
  var itemsSeq = 0;
  function loadItems(quiet, filterOnly) {
    if (!state.itemsOn || !itemsEls) return Promise.resolve();
    var seq = ++itemsSeq;
    var f = state.filter;
    var jobs = [filterOnly ? Promise.resolve({ ok: true, data: state.itemsAll }) : jreq('GET', '/api/ma-items?filter=all')];
    if (f !== 'all') jobs.push(jreq('GET', '/api/ma-items?filter=' + encodeURIComponent(f)));
    if (!isSecretary() && !filterOnly) jobs.push(jreq('GET', '/api/ma-items/pending-backfill'));
    return Promise.all(jobs).then(function (r) {
      if (seq !== itemsSeq) return;
      if (!r[0].ok) { if (!quiet) toast((r[0].data && r[0].data.error) || '無法載入到期管理資料'); return; }
      state.itemsAll = r[0].data;
      var rows = (f === 'all' ? r[0].data : (r[1].ok ? r[1].data : r[0].data)).items || [];
      if (f !== 'all' && !r[1].ok) rows = [];
      if (!isSecretary() && !filterOnly) {
        var pb = r[r.length - 1];
        state.pending = pb && pb.ok ? pb.data : null;
        // 補登限負責業務本人與管理員（伺服器同規則）：其他角色（主管、行銷、會計）看得到清單但不能補，所以不列
        if (state.pending && state.pending.items) state.pending = { items: state.pending.items.filter(function (p) { return isAdmin() || (state.user && p.owner === state.user.username); }) };
      }
      renderCards(); renderChips(); renderRows(rows);
      var n = state.pending && state.pending.items ? state.pending.items.length : 0;
      itemsEls.backfill.hidden = isSecretary() || n === 0;
      itemsEls.backfill.textContent = '待補納管（' + n + '）';
    });
  }

  function openEditTerm(it) {
    termDialog({
      title: '修改「' + it.company + '」的服務期間',
      lead: '品項：' + (it.product || '') + '。修改會留下稽核紀錄；已簽約的合約起迄日會一併更新。',
      kind: it.termKind || 'ma', from: it.from, to: it.to, showReason: true, confirmLabel: '儲存期間',
      submit: function (v) {
        return jreq('PATCH', '/api/ma-items/' + encodeURIComponent(it.id) + '/term', { termFrom: v.from, termTo: v.to, reason: v.reason }).then(function (r) {
          if (!r.ok) return { ok: false, message: (r.data && r.data.error) || '修改失敗' };
          toast('✓ 服務期間已更新');
          loadItems(true);
          return { ok: true };
        });
      },
    });
  }
  function openBackfill() {
    var items = (state.pending && state.pending.items) || [];
    var d = openDialog({ title: '待補納管（' + items.length + '）', role: 'dialog', wide: true });
    d.body.appendChild(el('p', 'mat-lead', '這些商機已是「成交」，但沒有納入到期管理（例如 Excel 匯入、或功能啟用前就成交的舊單）。補登時填服務期間；留空＝取已核准報價單上的期間。'));
    var ul = el('ul', 'mat-list');
    items.forEach(function (p) {
      var li = el('li', 'mat-row');
      var main = el('div', 'mat-row-main');
      main.appendChild(el('strong', 'mat-name', p.company));
      main.appendChild(el('span', 'mat-meta', [p.product, '負責業務 ' + (p.ownerLabel || p.owner), p.achievedDate ? '成交日 ' + slash(p.achievedDate) : ''].filter(Boolean).join('　')));
      li.appendChild(main);
      var act = el('div', 'mat-row-actions');
      var b = button('補登…', 'mat-btn mat-btn-primary mat-btn-sm', function () {
        termDialog({
          title: '補登「' + p.company + '」的納管', lead: '品項：' + (p.product || '') + '。', kind: p.termKind || 'ma', allowBlank: true,
          note: '兩欄留空＝取已核准報價單上的服務期間；沒有報價單期間時請直接填寫。', confirmLabel: '補登納管',
          submit: function (v) {
            var body = { oppId: p.oppId }; if (v.from && v.to) { body.termFrom = v.from; body.termTo = v.to; }
            return jreq('POST', '/api/ma-items/backfill', body).then(function (r) {
              if (!r.ok) return { ok: false, message: (r.data && r.data.error) || '補登失敗' };
              toast('✓ 已補登納管');
              d.close(); loadItems(true); refresh(); refreshBell();
              return { ok: true };
            });
          },
        });
      });
      b.setAttribute('aria-label', '補登「' + p.company + '」的納管');
      act.appendChild(b); li.appendChild(act); ul.appendChild(li);
    });
    d.body.appendChild(ul);
    d.foot.appendChild(button('關閉', 'mat-btn mat-btn-ghost', function () { d.close(); }));
    focusFirst(d, '.mat-btn-primary');
  }

  window.ITTSMaTasks = {
    start: start, refresh: refresh, refreshSoon: refreshSoon,
    enabled: function () { return state.enabled; },
    wonTerm: wonTerm, onSection: onSection,
    openSummary: openSummary,
    openFromNotification: function () { return openSummary({ fetch: true }); },
    openItemsFromNotification: function () { if (state.itemsOn && typeof window.showSection === 'function') window.showSection('maItems'); },
    itemsOn: function () { return state.itemsOn; },
    _state: state, _resolveTermKind: resolveTermKind,
  };
})();
