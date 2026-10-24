/* MA／License 到期「活規則」提醒：鈴鐺區塊、報價單「這是續約嗎？」選擇器（使用者端，P3b）
 *
 * 後端（MA_TRACKING=1 才有，關閉時 poll-bundle 沒有 maReminders 鍵、renewal-candidates 回 404 → 這裡什麼都不建立）：
 *   GET /api/poll-bundle 的 maReminders（依當下條件即時推導；條件不再成立下一次輪詢就消失，前端不必自己清）、
 *   GET /api/ma-items/renewal-candidates?company=、GET /api/ma-items?filter=all（查已指定的續約對象）。
 *
 * 對外：window.ITTSMaRem = { paintBell, isRead, markRead, markAllRead, allRead, dayText, mountRenewPicker, ... }
 *   - paintBell(host, list, opts)：把 maReminders 畫成鈴鐺裡的一個群組（摘要「你有 N 項 MA／License 需要處理」可展開，每列公司、品項、
 *     剩餘／逾期天數、所屬規則名稱；點擊開到期管理頁）。已讀狀態沿用合約提醒的 localStorage 每日重置做法（鍵 marem_read_<id>_<台北日期>）。
 *   - mountRenewPicker(host, o)：報價單表單的「這是續約嗎？」橫幅＋下拉（候選＝GET renewal-candidates；選項只有品項與到期日）。
 * 安全：全部動態內容 createElement＋textContent（不經 innerHTML）、不使用 inline style／on* 屬性；樣式在 style.css 末尾的 .mar-*。
 */
(function () {
  'use strict';
  if (window.ITTSMaRem) return;

  var KIND_LABEL = { ma: 'MA', license: 'License' };
  var ROLE_LABEL = { owner: '業務', secretary: '秘書', manager1: '一級主管', manager2: '二級主管', marketing: '行銷' };
  var LIST_MAX = 20;

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = String(text);
    return e;
  }
  function slash(s) { return s ? String(s).replace(/-/g, '/') : ''; }
  function tpToday() { return new Date(Date.now() + 8 * 3600 * 1000).toISOString().slice(0, 10); }
  function dayText(n) {
    if (typeof n !== 'number') return '';
    if (n < 0) return '已逾期 ' + (-n) + ' 天';
    if (n === 0) return '今天到期';
    return '剩 ' + n + ' 天到期';
  }

  /* ───────── 已讀（localStorage，每日重置；讀寫失敗當作未讀）───────── */
  function readKey(id) { return 'marem_read_' + id + '_' + tpToday(); }
  function isRead(id) { try { return !!localStorage.getItem(readKey(id)); } catch (e) { return false; } }
  function markRead(id) { try { localStorage.setItem(readKey(id), '1'); } catch (e) { /* ignore */ } }
  function allRead(list) { return (list || []).every(function (x) { return isRead(x.id); }); }
  function markAllRead(list) { (list || []).forEach(function (x) { markRead(x.id); }); }
  function pruneOld() {
    // 前幾天的已讀鍵沒有用了：順手清掉（失敗就算了）
    try {
      var t = tpToday(), kill = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf('marem_read_') === 0 && k.slice(-10) !== t) kill.push(k);
      }
      kill.forEach(function (k) { localStorage.removeItem(k); });
    } catch (e) { /* ignore */ }
  }

  /* ───────── 鈴鐺區塊 ───────── */
  var bellOpen = false;   // 輪詢重畫時保留展開狀態
  function sumBody(list) {
    var f = list[0];
    var s = (f.company || '') + (f.product ? '・' + f.product : '') + '　' + dayText(f.days);
    return list.length > 1 ? s + ' 等共 ' + list.length + ' 項' : s;
  }
  function ruleText(x) {
    if (x.rule && x.rule.name) return '規則：' + x.rule.name + (x.ruleCount > 1 ? '（另有 ' + (x.ruleCount - 1) + ' 條）' : '');
    return x.ownerLabel ? '業務：' + x.ownerLabel : '';
  }
  function rolesText(x) {
    if (!x.roles || !x.roles.length) return '';
    return '你是：' + x.roles.map(function (r) { return ROLE_LABEL[r] || r; }).join('、');
  }
  /**
   * host：renderNotifList 預留的空容器；list：poll-bundle 的 maReminders（已依剩餘天數排序）。
   * opts: { open(item)→關閉下拉並開到期管理頁, onRead()→重算未讀 badge }
   */
  function paintBell(host, list, opts) {
    var o = opts || {};
    host.textContent = '';
    if (!list || !list.length) return;
    pruneOld();
    var wrap = el('div', 'mar-bell');
    var unread = list.filter(function (x) { return !isRead(x.id); }).length;

    var sum = el('div', 'notif-item mar-bell-sum' + (unread ? ' unread' : ''));
    sum.setAttribute('role', 'button'); sum.tabIndex = 0;
    sum.setAttribute('aria-expanded', bellOpen ? 'true' : 'false');
    sum.setAttribute('data-marem', '1');
    sum.appendChild(el('div', 'notif-icon', '⏳'));
    var ct = el('div', 'notif-content');
    ct.appendChild(el('div', 'notif-title', '你有 ' + list.length + ' 項 MA／License 需要處理'));
    ct.appendChild(el('div', 'notif-body', sumBody(list)));
    ct.appendChild(el('div', 'notif-body mar-bell-hint', bellOpen ? '點一下收合清單' : '點一下展開清單'));
    sum.appendChild(ct);
    wrap.appendChild(sum);

    var box = el('div', 'mar-bell-list');
    box.hidden = !bellOpen;
    list.slice(0, LIST_MAX).forEach(function (x) {
      var row = el('div', 'mar-bell-row' + (isRead(x.id) ? '' : ' mar-bell-unread'));
      row.setAttribute('role', 'button'); row.tabIndex = 0;
      row.setAttribute('data-marem-row', x.id);
      var main = el('div', 'mar-bell-main');
      main.appendChild(el('strong', 'mar-bell-co', x.company || ''));
      var pd = el('span', 'mar-bell-pd', x.product || '');
      main.appendChild(pd);
      if (x.termKind && KIND_LABEL[x.termKind]) main.appendChild(el('span', 'mar-kind', KIND_LABEL[x.termKind]));
      row.appendChild(main);
      var meta = el('div', 'mar-bell-meta');
      meta.appendChild(el('span', 'mar-lv mar-lv-' + (x.level === 'danger' ? 'danger' : x.level === 'warn' ? 'warn' : 'info'), dayText(x.days)));
      var rt = ruleText(x);
      if (rt) meta.appendChild(el('span', 'mar-bell-rule', rt));
      row.appendChild(meta);
      var rl = rolesText(x);
      if (rl) row.appendChild(el('div', 'mar-bell-roles', rl));
      function go() {
        markRead(x.id);
        row.classList.remove('mar-bell-unread');
        if (!list.some(function (y) { return !isRead(y.id); })) sum.classList.remove('unread');
        if (o.onRead) o.onRead();
        if (o.open) o.open(x);
      }
      row.addEventListener('click', function (e) { e.stopPropagation(); go(); });
      row.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); go(); } });
      box.appendChild(row);
    });
    if (list.length > LIST_MAX) box.appendChild(el('div', 'mar-bell-more', '另有 ' + (list.length - LIST_MAX) + ' 項未列出，請開啟到期管理頁查看。'));
    var openBtn = el('button', 'mar-bell-openpage', '開啟 MA／License 到期管理 ›');
    openBtn.type = 'button';
    openBtn.addEventListener('click', function (e) { e.stopPropagation(); if (o.open) o.open(null); });
    box.appendChild(openBtn);
    wrap.appendChild(box);

    function toggle() {
      bellOpen = !bellOpen;
      box.hidden = !bellOpen;
      sum.setAttribute('aria-expanded', bellOpen ? 'true' : 'false');
      var hint = sum.querySelector('.mar-bell-hint'); if (hint) hint.textContent = bellOpen ? '點一下收合清單' : '點一下展開清單';
      if (bellOpen) {   // 打開清單＝看過了：整組標已讀（每日重置）
        markAllRead(list);
        sum.classList.remove('unread');
        box.querySelectorAll('.mar-bell-unread').forEach(function (r) { r.classList.remove('mar-bell-unread'); });
        if (o.onRead) o.onRead();
      }
    }
    sum.addEventListener('click', function (e) { e.stopPropagation(); toggle(); });
    sum.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); toggle(); } });
    host.appendChild(wrap);
  }

  /* ───────── 報價單「這是續約嗎？」選擇器 ─────────
   * o: {
   *   getCompany(): string   目前報價單上的客戶公司（候選以此查）
   *   value: string          表單載入時已存的 renewOf（沒有＝''）
   *   original: string       已存的客戶公司（編輯舊單）；公司改了就清掉已選的續約對象（伺服器會以公司不同擋下）
   *   locked: boolean        已核准後唯讀：只顯示「續約自：…」
   *   fetchJson(url)→Promise<object|null>
   * }
   * 回傳 { refresh(), value(), touched(), destroy() }。value() 只在使用者動過時才有意義：
   *   touched()＝false → 呼叫端不要送 renewOf（沿用伺服器的舊值，防舊分頁／未載入完成洗掉）。*/
  var pickerSeq = 0;
  function mountRenewPicker(host, o) {
    var seq = 0, destroyed = false;
    var cur = typeof o.value === 'string' ? o.value : '';
    var initial = cur;
    var touched = false;
    var cache = {};   // 公司 → candidates
    var known = null; // 已指定續約對象的資訊（品項、到期日）
    var timer = null;
    host.textContent = '';
    host.hidden = true;

    function itemLabel(it) {
      var d = it.daysLeft;
      var tail = typeof d === 'number' ? (d < 0 ? '，已逾期 ' + (-d) + ' 天' : '，剩 ' + d + ' 天') : '';
      return (it.product || '（未命名品項）') + '（' + (KIND_LABEL[it.termKind] || '') + ' 到期 ' + slash(it.to) + tail + (it.hasRenewalQuote ? '，已有續約報價單' : '') + '）';
    }
    function paintLocked(info) {
      host.textContent = '';
      host.className = 'mar-renew mar-renew-locked';
      var t = '續約自：' + (info ? (info.product || '（未命名品項）') + '（到期 ' + slash(info.to) + '）' : '（已指定續約對象）');
      host.appendChild(el('span', 'mar-renew-tag', t));
      host.appendChild(el('span', 'mar-renew-note', '　僅內部可見，不會印在客戶版報價單上。'));
      host.hidden = false;
    }
    function paint(cands) {
      host.textContent = '';
      host.className = 'mar-renew';
      var items = (cands || []).slice();
      var hasCur = !!cur && items.some(function (x) { return x.id === cur; });
      if (cur && !hasCur && known) items.unshift({ id: cur, product: known.product, termKind: known.termKind, to: known.to, daysLeft: known.daysLeft, hasRenewalQuote: false });
      if (!items.length && !cur) { host.hidden = true; return; }
      var id = 'mar-renew-sel-' + (++pickerSeq);
      var head = el('div', 'mar-renew-head');
      head.appendChild(el('strong', 'mar-renew-title', '這是續約嗎？'));
      head.appendChild(el('span', 'mar-renew-sub', '這家客戶有 ' + items.length + ' 個 MA／License 即將或已經到期。選擇後，續約報價單送簽時，系統會自動停止對該項目的「無條件」到期提醒；成交時該項目標為已續約。僅內部可見，不會印在客戶版報價單上。'));
      host.appendChild(head);
      var lab = el('label', 'mar-renew-label', '續約對象');
      lab.setAttribute('for', id);
      var sel = document.createElement('select');
      sel.id = id; sel.className = 'mar-renew-select';
      function opt(v, t) { var op = document.createElement('option'); op.value = v; op.textContent = t; sel.appendChild(op); }
      opt('', '不是續約（新購或其他）');
      var seen = {};
      items.forEach(function (it) { if (seen[it.id]) return; seen[it.id] = 1; opt(it.id, '續約：' + itemLabel(it)); });
      sel.value = cur || '';
      sel.addEventListener('change', function () {
        touched = true; cur = sel.value;
        if (o.onChange) o.onChange(cur);
      });
      host.appendChild(lab); host.appendChild(sel);
      host.hidden = false;
    }
    function lookupKnown() {
      if (!cur || known || !o.fetchJson) return Promise.resolve();
      return o.fetchJson('/api/ma-items?filter=all').then(function (r) {
        var it = r && Array.isArray(r.items) ? r.items.filter(function (x) { return x.id === cur; })[0] : null;
        known = it ? { product: it.product, termKind: it.termKind, to: it.to, daysLeft: it.daysLeft } : { product: '', termKind: '', to: '', daysLeft: null };
      });
    }
    function refresh() {
      if (destroyed) return Promise.resolve();
      var company = String((o.getCompany && o.getCompany()) || '').trim();
      // 客戶改了：之前選的續約對象不再成立 → 清掉（若舊單原本有值，會送空字串清除）
      if (cur && o.original != null && company && company !== o.original) { cur = ''; touched = initial !== ''; }
      var mine = ++seq;
      if (o.locked) {
        return lookupKnown().then(function () { if (!destroyed && mine === seq) paintLocked(cur ? known : null); });
      }
      if (!company) { paint([]); return Promise.resolve(); }
      var pre = cache[company] ? Promise.resolve(cache[company]) : (o.fetchJson ? o.fetchJson('/api/ma-items/renewal-candidates?company=' + encodeURIComponent(company)) : Promise.resolve(null)).then(function (r) {
        var list = r && Array.isArray(r.items) ? r.items : [];
        cache[company] = list; return list;
      });
      return pre.then(function (list) { return lookupKnown().then(function () { return list; }); }).then(function (list) {
        if (destroyed || mine !== seq) return;
        paint(list);
      });
    }
    function refreshSoon() { clearTimeout(timer); timer = setTimeout(refresh, 350); }
    return {
      refresh: refresh,
      refreshSoon: refreshSoon,
      value: function () { return cur; },
      touched: function () { return touched; },
      initial: function () { return initial; },
      destroy: function () { destroyed = true; clearTimeout(timer); host.textContent = ''; host.hidden = true; },
    };
  }

  window.ITTSMaRem = {
    paintBell: paintBell, isRead: isRead, markRead: markRead, markAllRead: markAllRead, allRead: allRead, dayText: dayText,
    mountRenewPicker: mountRenewPicker, _tpToday: tpToday,
  };
})();
