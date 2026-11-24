/* 後台「提醒規則」頁（MA／License 到期活規則，P3b）
 *
 * 後端（MA_TRACKING=1 才有，關閉時 GET /api/admin/reminder-rules 回 404 → 這裡什麼都不建立，後台與沒有此功能時逐位元相同）：
 *   GET／PUT /api/admin/reminder-rules（PUT 整份取代；每條規則請帶回既有 id；APPROVAL_CONFIG_DUAL_REVIEW=1 時回 202 待覆核）、
 *   POST /api/admin/reminder-rules/simulate（以目前編輯中、尚未儲存的規則試算）、GET /api/ma-items?filter=all（試算器可選真實納管項）。
 * 載入方式：admin.html 的 loadUsers 之後呼叫 MaRemAdmin.probe({ fetch: adminFetch, crApi, toast, canLeave })；
 *   探測回 200 才動態插入側欄項「提醒規則」（報價與簽核群組）、section 與樣式表（ma-reminders-admin.css）；之後比照既有後台 lazy-init。
 * 安全：全部動態內容 createElement＋textContent（不經 innerHTML）、不使用 inline style／on* 屬性；
 *   試算器的時間軸是 SVG（位置用 setAttribute 設定座標，顏色與字型全在 CSS class），樣式前綴 .mar-a-*。
 * 純函式 validateRules／toApi／fromApi／whenText／simText 掛在 MaRemAdmin 上，scripts/_ma-reminders-ui.js 以實際頁面驗證與伺服器 validateConfig 一致。
 */
(function () {
  'use strict';
  if (window.MaRemAdmin) return;

  var ROLES = ['owner', 'secretary', 'manager1', 'manager2', 'marketing'];
  var ROLE_LABEL = { owner: '業務', secretary: '秘書', manager1: '一級主管', manager2: '二級主管', marketing: '行銷' };
  var CONDITIONS = [
    ['always', '無條件（續約報價單送簽後自動停止）'],
    ['no_renewal_quote', '尚未建立續約報價單'],
    ['quote_not_submitted', '續約報價單仍是草稿（未送簽）'],
    ['not_renewed', '仍在追蹤中（未續約、未標不續約）'],
  ];
  var KIND_OPT = [['all', '全部類型'], ['license', '只有 License'], ['ma', '只有 MA']];
  var LIM = { maxRules: 20, nameMax: 40, nMax: 365, repeatMax: 90 };
  var QS_TEXT = { none: '尚未建立續約報價單', draft: '續約報價單尚未送簽', sent: '續約報價單已送簽' };
  var REASON_TEXT = {
    '': '成立', DISABLED: '規則已停用', APPLIES: '不適用此類型', TOO_OLD: '逾期過久，不再提醒', NOT_REACHED: '尚未到達這個時間點',
    EXPIRED_WINDOW: '到期前的規則，過了迄日就不再提醒', CONDITION: '條件目前不成立', NOT_TRACKING: '項目已不在追蹤中（不續約等）',
  };

  /* ───────── 小工具 ───────── */
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = String(text);
    return e;
  }
  function btn(label, cls, onClick) {
    var b = el('button', cls, label); b.type = 'button';
    if (onClick) b.addEventListener('click', onClick);
    return b;
  }
  var uidN = 0;
  function uid(p) { uidN++; return 'mara-' + p + '-' + uidN; }
  function fmtTime(iso) {
    var t = Date.parse(iso); if (!isFinite(t)) return '';
    return new Date(t).toLocaleString('zh-TW', { hour12: false });
  }
  function whenText(m) {
    var base = m.dir === 'before' ? '到期前 ' + m.nRaw + ' 天' : (m.dir === 'on' ? '到期當天' : '到期後 ' + m.nRaw + ' 天');
    return base + (m.dir !== 'before' && Number(m.repeatRaw) > 0 ? '，每 ' + Number(m.repeatRaw) + ' 天重複' : '');
  }
  function dayText(n) { return n > 0 ? '剩 ' + n + ' 天到期' : (n === 0 ? '今天到期' : '已逾期 ' + (-n) + ' 天'); }

  /* ───────── 規則模型（編輯用）↔ API ───────── */
  function fromApi(r) {
    var w = r.whenDays;
    var dir = w > 0 ? 'before' : (w === 0 ? 'on' : 'after');
    return {
      id: r.id || '', name: r.name || '', enabled: r.enabled !== false, dir: dir, nRaw: dir === 'on' ? '0' : String(Math.abs(w)),
      repeatRaw: String(r.repeatEveryDays || 0), notify: (r.notify || []).slice(), condition: r.condition || 'always',
      applies: r.appliesTo === 'all' || !r.appliesTo ? 'all' : (Array.isArray(r.appliesTo) && r.appliesTo.length === 1 ? r.appliesTo[0] : 'all'),
    };
  }
  function whenDaysOf(m) {
    if (m.dir === 'on') return 0;
    var n = /^\d{1,3}$/.test(String(m.nRaw).trim()) ? Number(m.nRaw) : NaN;
    return m.dir === 'before' ? n : (isNaN(n) ? NaN : -n);
  }
  /** 只在模型有效時才可呼叫；id 空白就不帶（伺服器產生），有就帶回（去重鍵含規則 id） */
  function toApi(m) {
    var w = whenDaysOf(m);
    var o = {
      name: String(m.name).trim(), enabled: !!m.enabled, whenDays: w,
      repeatEveryDays: w > 0 ? 0 : (/^\d{1,2}$/.test(String(m.repeatRaw).trim()) ? Number(m.repeatRaw) : 0),
      notify: ROLES.filter(function (r) { return m.notify.indexOf(r) >= 0; }), condition: m.condition,
      appliesTo: m.applies === 'all' ? 'all' : [m.applies],
    };
    if (m.id) o.id = m.id;
    return o;
  }
  /** 與伺服器 lib/maReminders.js validateConfig 同規則（即時驗證用）。回 { perRule:[{name,when,repeat,notify,dup}], count, global } */
  function validateRules(models) {
    var out = { perRule: [], count: 0, global: '' };
    if (models.length > LIM.maxRules) { out.global = '提醒規則最多 ' + LIM.maxRules + ' 條'; out.count++; }
    var dups = {};
    models.forEach(function (m) {
      var e = {};
      var name = String(m.name == null ? '' : m.name).trim();
      if (!name) e.name = '請填寫規則名稱';
      else if (Array.from(name).length > LIM.nameMax) e.name = '名稱最長 ' + LIM.nameMax + ' 個字';
      else if (/[\u0000-\u001f\u007f-\u009f<>]/.test(name)) e.name = '名稱含有不允許的字元（< >）';
      if (m.dir !== 'on') {
        var raw = String(m.nRaw == null ? '' : m.nRaw).trim();
        if (!/^\d{1,3}$/.test(raw) || Number(raw) < 1 || Number(raw) > LIM.nMax) e.when = '天數必須是 1 ～ ' + LIM.nMax + ' 的整數';
      }
      if (m.dir !== 'before') {
        var rr = String(m.repeatRaw == null ? '' : m.repeatRaw).trim();
        if (rr !== '' && (!/^\d{1,2}$/.test(rr) || Number(rr) > LIM.repeatMax)) e.repeat = '重複間隔必須是 0（只發一次）或 1 ～ ' + LIM.repeatMax + ' 天';
      }
      if (!m.notify || !m.notify.length) e.notify = '至少要選一個通知對象';
      var w = whenDaysOf(m);
      if (!e.when && !isNaN(w)) {
        var dk = m.condition + '|' + m.applies + '|' + w;
        if (dups[dk] !== undefined) { e.dup = '與第 ' + (dups[dk] + 1) + ' 條規則重複：相同條件與適用類型下，同一個時間點只能有一條'; }
        else dups[dk] = out.perRule.length;
      }
      var n = Object.keys(e).length;
      out.count += n;
      out.perRule.push(e);
    });
    return out;
  }
  /** 模擬器顯示給各角色的提醒文字（與伺服器 dynamicFor 的標題同格式） */
  function simText(sim) {
    var co = (sim.item && sim.item.company) || '（示範客戶）';
    var pd = (sim.item && sim.item.product) || '（示範品項）';
    return '⏰ ' + co + '：' + pd + ' ' + dayText(sim.daysLeft) + '｜' + (QS_TEXT[sim.quoteState] || '');
  }

  /* ───────── 狀態 ───────── */
  var S = {
    api: null, inited: false, installed: false, booting: false, models: [], defaults: [], config: null, limits: null,
    cr: null, simSeq: 0, simTimer: null, sim: null, items: null, host: null, els: {}, saving: false, dirty: false,
  };

  /* ───────── 安裝（探測成功後）───────── */
  function install(boot) {
    if (S.installed) return;
    S.installed = true;
    S.boot = boot;
    // 樣式表（動態插入：功能關閉時後台不會多載任何檔案）
    if (!document.getElementById('maRemAdminCss')) {
      var lk = document.createElement('link'); lk.rel = 'stylesheet'; lk.href = 'ma-reminders-admin.css'; lk.id = 'maRemAdminCss';
      document.head.appendChild(lk);
    }
    var dom = document.querySelector('.nav-domain[data-domain="quote"]');
    var inner = dom && dom.querySelector('.nav-domain-inner');
    var main = document.querySelector('.admin-main');
    if (!inner || !main) return;
    var item = el('div', 'admin-nav-item nav-sub');
    item.setAttribute('data-sec', 'reminder-rules'); item.id = 'marNavItem';
    item.setAttribute('role', 'button'); item.tabIndex = 0;
    item.appendChild(el('span', 'admin-nav-icon', '⏰'));
    item.appendChild(document.createTextNode(' 提醒規則'));
    var after = inner.querySelector('[data-sec="approval-flow"]');
    if (after && after.nextSibling) inner.insertBefore(item, after.nextSibling); else inner.appendChild(item);
    var cnt = dom.querySelector('.nav-domain-count');
    if (cnt && /^\d+$/.test(cnt.textContent.trim())) cnt.textContent = String(Number(cnt.textContent.trim()) + 1);

    var sec = el('div', 'admin-section'); sec.id = 'sec-reminder-rules';
    sec.appendChild(el('div', 'section-title', '⏰ 提醒規則（MA／License 到期）'));
    sec.appendChild(el('div', 'section-sub', '設定 MA／License 到期前後，系統在什麼時間點、通知誰、符合什麼條件才提醒。規則是「活的」：存檔後下一次輪詢就生效；條件達成（例如續約報價單送簽）時，鈴鐺上的提醒會自動消失。'));
    var hostEl = el('div', 'mar-a'); hostEl.id = 'marAdminHost';
    sec.appendChild(hostEl);
    var anchor = document.getElementById('sec-flow-designer') || main.querySelector('.admin-section:last-of-type');
    if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(sec, anchor.nextSibling); else main.appendChild(sec);
    S.host = hostEl;

    function activate() {
      if (typeof boot.canLeave === 'function' && boot.canLeave(item) === false) return;
      document.querySelectorAll('.admin-nav-item').forEach(function (i) { i.classList.remove('active'); });
      document.querySelectorAll('.admin-section').forEach(function (s) { s.classList.remove('active'); });
      item.classList.add('active'); sec.classList.add('active');
      if (typeof window.adminSetDomainOpen === 'function') window.adminSetDomainOpen(dom, true);
      main.scrollTop = 0;
      init();
    }
    item.addEventListener('click', activate);
    item.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); activate(); } });
  }

  /* ───────── API ───────── */
  function jsonOf(r) { return r.json().catch(function () { return {}; }); }
  function apiGetRules() { return S.boot.fetch('/api/admin/reminder-rules').then(function (r) { return jsonOf(r).then(function (d) { return { ok: r.ok, status: r.status, data: d }; }); }); }
  function apiJson(method, url, body) {
    return S.boot.fetch(url, { method: method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
      .then(function (r) { return jsonOf(r).then(function (d) { return { ok: r.ok, status: r.status, data: d }; }); });
  }

  /* ───────── 初始化頁面 ───────── */
  function init() {
    if (S.inited) { return; }
    S.inited = true;
    S.host.textContent = '';
    S.host.appendChild(el('p', 'mar-a-muted', '載入中…'));
    apiGetRules().then(function (r) {
      if (!r.ok || !r.data || !r.data.config) { S.host.textContent = ''; S.host.appendChild(el('p', 'mar-a-error', '載入失敗：' + ((r.data && r.data.error) || r.status))); S.inited = false; return; }
      S.config = r.data.config; S.defaults = r.data.defaults || []; S.limits = r.data.limits || null;
      S.models = (S.config.rules || []).map(fromApi);
      S.dirty = false;
      build();
    });
  }

  function reloadFromServer() {
    return apiGetRules().then(function (r) {
      if (!r.ok || !r.data || !r.data.config) return;
      S.config = r.data.config; S.defaults = r.data.defaults || S.defaults;
      S.models = (S.config.rules || []).map(fromApi);
      S.dirty = false;
      renderRules(); paintMeta(); runValidation(); scheduleSim(0);
    });
  }

  /* ───────── 畫面 ───────── */
  function build() {
    var host = S.host;
    host.textContent = '';
    var E = S.els = {};

    // 摘要列：更新資訊＋雙人覆核
    var meta = el('div', 'mar-a-card mar-a-meta');
    E.meta = el('div', 'mar-a-metaline');
    meta.appendChild(E.meta);
    E.crHost = el('div', 'mar-a-crhost'); E.crHost.id = 'marCrHost';
    meta.appendChild(E.crHost);
    host.appendChild(meta);

    // 規則編輯
    var card = el('section', 'mar-a-card mar-a-rules');
    card.setAttribute('aria-labelledby', 'marRulesTitle');
    var hd = el('div', 'mar-a-hd');
    var ttl = el('h3', 'mar-a-title', '規則'); ttl.id = 'marRulesTitle';
    hd.appendChild(ttl);
    E.count = el('span', 'mar-a-count', '');
    hd.appendChild(E.count);
    card.appendChild(hd);
    card.appendChild(el('p', 'mar-a-help', '時間點的「到期前 N 天」會在剩 N 天時（含之後漏跑補發）觸發一次；「到期當天／到期後」可設重複間隔。到期前的規則過了迄日就不再提醒；項目標記不續約、已續約、結束或作廢後，所有規則都會停止。'));
    E.errAll = el('div', 'mar-a-error'); E.errAll.setAttribute('role', 'alert'); E.errAll.hidden = true;
    card.appendChild(E.errAll);

    var wrap = el('div', 'mar-a-tblwrap');
    wrap.tabIndex = 0; wrap.setAttribute('role', 'region'); wrap.setAttribute('aria-label', '提醒規則編輯表（窄視窗可左右捲動）');
    var tbl = el('table', 'mar-a-tbl');
    tbl.appendChild(el('caption', 'mar-a-sr', '提醒規則'));
    var thead = document.createElement('thead'); var trh = document.createElement('tr');
    ['啟用', '規則名稱', '時間點', '通知對象', '條件', '適用類型', '重複', '操作'].forEach(function (h) { var th = el('th', null, h); th.scope = 'col'; trh.appendChild(th); });
    thead.appendChild(trh); tbl.appendChild(thead);
    E.tbody = document.createElement('tbody'); tbl.appendChild(E.tbody);
    wrap.appendChild(tbl); card.appendChild(wrap);

    var bar = el('div', 'mar-a-bar');
    E.addBtn = btn('＋ 新增規則', 'mar-a-btn', addRule);
    E.sortBtn = btn('依時間點排序', 'mar-a-btn', sortRules);
    E.defBtn = btn('還原預設規則', 'mar-a-btn mar-a-btn-warn', askDefaults);
    E.saveBtn = btn('儲存', 'mar-a-btn mar-a-btn-primary', save);
    bar.appendChild(E.addBtn); bar.appendChild(E.sortBtn); bar.appendChild(E.defBtn);
    var sp = el('span', 'mar-a-spacer'); bar.appendChild(sp);
    E.status = el('span', 'mar-a-status'); E.status.setAttribute('role', 'status'); E.status.setAttribute('aria-live', 'polite');
    bar.appendChild(E.status); bar.appendChild(E.saveBtn);
    card.appendChild(bar);
    E.confirm = el('div', 'mar-a-confirm'); E.confirm.hidden = true;
    E.confirm.setAttribute('role', 'alertdialog'); E.confirm.setAttribute('aria-label', '確認還原預設規則');
    E.confirm.appendChild(el('span', null, '編輯區將改為系統內建的 4 條預設規則（還沒儲存，可以再改；按「儲存」才會生效）。'));
    E.confirm.appendChild(btn('確定還原', 'mar-a-btn mar-a-btn-warn', function () { applyDefaults(); }));
    E.confirm.appendChild(btn('取消', 'mar-a-btn', function () { E.confirm.hidden = true; E.defBtn.focus(); }));
    card.appendChild(E.confirm);
    host.appendChild(card);

    buildSim(host);
    renderRules(); paintMeta(); runValidation();
    mountReview();
    scheduleSim(0);
    loadItems();
  }

  function paintMeta() {
    var c = S.config || {};
    var E = S.els;
    E.meta.textContent = '';
    if (c.isDefault || !c.updatedAt) {
      E.meta.appendChild(el('span', 'mar-a-pill mar-a-pill-def', '目前使用系統預設規則'));
      E.meta.appendChild(el('span', 'mar-a-muted', '　尚未有人自訂；按「儲存」後才會改用你編輯的規則。'));
    } else {
      E.meta.appendChild(el('span', 'mar-a-pill', '自訂規則'));
      E.meta.appendChild(el('span', 'mar-a-muted', '　最後更新：' + (c.updatedBy || '（未知）') + '　' + fmtTime(c.updatedAt)));
    }
  }

  /* ── 規則列 ── */
  function renderRules() {
    var E = S.els;
    E.tbody.textContent = '';
    S.models.forEach(function (m, i) { E.tbody.appendChild(ruleRow(m, i)); });
    E.addBtn.disabled = S.models.length >= LIM.maxRules;
    E.addBtn.title = S.models.length >= LIM.maxRules ? '最多 ' + LIM.maxRules + ' 條規則' : '';
    E.count.textContent = S.models.length + ' / ' + LIM.maxRules + ' 條';
    if (!S.models.length) {
      var tr = document.createElement('tr'); var td = document.createElement('td'); td.colSpan = 8; td.className = 'mar-a-empty';
      td.textContent = '目前沒有任何規則：不會發出任何提醒。可以按「新增規則」或「還原預設規則」。'; tr.appendChild(td); E.tbody.appendChild(tr);
    }
  }
  function cell(tr, label) { var td = document.createElement('td'); td.setAttribute('data-label', label); tr.appendChild(td); return td; }
  function labelled(input, text) { input.setAttribute('aria-label', text); return input; }
  function fieldErr(td, key) { var s = el('div', 'mar-a-ferr'); s.setAttribute('data-err', key); s.hidden = true; td.appendChild(s); return s; }

  function ruleRow(m, i) {
    var tr = document.createElement('tr'); tr.className = 'mar-a-row'; tr.setAttribute('data-rule-index', String(i));
    var tag = '第 ' + (i + 1) + ' 條規則';
    // 啟用
    var c0 = cell(tr, '啟用');
    var cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = !!m.enabled; cb.className = 'mar-a-chk';
    labelled(cb, tag + ' 啟用');
    cb.addEventListener('change', function () { m.enabled = cb.checked; changed(); });
    c0.appendChild(cb);
    // 名稱
    var c1 = cell(tr, '規則名稱');
    var nm = document.createElement('input'); nm.type = 'text'; nm.className = 'mar-a-inp mar-a-name'; nm.value = m.name; nm.maxLength = 80;
    labelled(nm, tag + ' 名稱'); nm.addEventListener('input', function () { m.name = nm.value; changed(); });
    c1.appendChild(nm); fieldErr(c1, 'name'); fieldErr(c1, 'dup');
    // 時間點
    var c2 = cell(tr, '時間點');
    var dir = document.createElement('select'); dir.className = 'mar-a-sel mar-a-dir';
    [['before', '到期前'], ['on', '到期當天'], ['after', '到期後']].forEach(function (o) { var op = document.createElement('option'); op.value = o[0]; op.textContent = o[1]; dir.appendChild(op); });
    dir.value = m.dir; labelled(dir, tag + ' 時間點方向');
    var nIn = document.createElement('input'); nIn.type = 'number'; nIn.className = 'mar-a-inp mar-a-n'; nIn.min = '1'; nIn.max = String(LIM.nMax); nIn.step = '1';
    nIn.value = m.dir === 'on' ? '' : m.nRaw; nIn.disabled = m.dir === 'on'; labelled(nIn, tag + ' 天數');
    var unit = el('span', 'mar-a-unit', '天');
    var rep = document.createElement('input'); // 先建立，方便 dir 變更時連動
    rep.type = 'number'; rep.className = 'mar-a-inp mar-a-rep'; rep.min = '0'; rep.max = String(LIM.repeatMax); rep.step = '1';
    rep.value = m.dir === 'before' ? '0' : m.repeatRaw; rep.disabled = m.dir === 'before'; labelled(rep, tag + ' 重複間隔（天，0＝只發一次）');
    dir.addEventListener('change', function () {
      m.dir = dir.value;
      if (m.dir === 'on') { m.nRaw = '0'; nIn.value = ''; nIn.disabled = true; }
      else { nIn.disabled = false; if (m.nRaw === '0' || m.nRaw === '') { m.nRaw = '1'; } nIn.value = m.nRaw; }
      if (m.dir === 'before') { m.repeatRaw = '0'; rep.value = '0'; rep.disabled = true; } else { rep.disabled = false; rep.value = m.repeatRaw; }
      changed();
    });
    nIn.addEventListener('input', function () { m.nRaw = nIn.value; changed(); });
    var row2 = el('div', 'mar-a-inline'); row2.appendChild(dir); row2.appendChild(nIn); row2.appendChild(unit);
    c2.appendChild(row2); fieldErr(c2, 'when');
    // 通知對象
    var c3 = cell(tr, '通知對象');
    var fs = document.createElement('fieldset'); fs.className = 'mar-a-notify';
    var lg = el('legend', 'mar-a-sr', tag + ' 通知對象'); fs.appendChild(lg);
    ROLES.forEach(function (r) {
      var lab = el('label', 'mar-a-chkl');
      var c = document.createElement('input'); c.type = 'checkbox'; c.value = r; c.checked = m.notify.indexOf(r) >= 0;
      c.addEventListener('change', function () {
        var set = ROLES.filter(function (x) { return x === r ? c.checked : m.notify.indexOf(x) >= 0; });
        m.notify = set; changed();
      });
      lab.appendChild(c); lab.appendChild(document.createTextNode(ROLE_LABEL[r]));
      fs.appendChild(lab);
    });
    c3.appendChild(fs); fieldErr(c3, 'notify');
    // 條件
    var c4 = cell(tr, '條件');
    var cond = document.createElement('select'); cond.className = 'mar-a-sel mar-a-cond';
    CONDITIONS.forEach(function (o) { var op = document.createElement('option'); op.value = o[0]; op.textContent = o[1]; cond.appendChild(op); });
    cond.value = m.condition; labelled(cond, tag + ' 條件');
    cond.addEventListener('change', function () { m.condition = cond.value; changed(); });
    c4.appendChild(cond);
    // 適用類型
    var c5 = cell(tr, '適用類型');
    var ap = document.createElement('select'); ap.className = 'mar-a-sel mar-a-applies';
    KIND_OPT.forEach(function (o) { var op = document.createElement('option'); op.value = o[0]; op.textContent = o[1]; ap.appendChild(op); });
    ap.value = m.applies; labelled(ap, tag + ' 適用類型');
    ap.addEventListener('change', function () { m.applies = ap.value; changed(); });
    c5.appendChild(ap);
    // 重複
    var c6 = cell(tr, '重複');
    var rw = el('div', 'mar-a-inline mar-a-nowrap'); rw.appendChild(el('span', 'mar-a-unit', '每')); rw.appendChild(rep); rw.appendChild(el('span', 'mar-a-unit', '天'));
    rep.addEventListener('input', function () { m.repeatRaw = rep.value; changed(); });
    c6.appendChild(rw); c6.appendChild(el('div', 'mar-a-hint', '0＝只發一次')); fieldErr(c6, 'repeat');
    // 操作
    var c7 = cell(tr, '操作');
    var ops = el('div', 'mar-a-ops');
    var up = btn('↑', 'mar-a-btn mar-a-btn-icon', function () { move(i, -1); }); up.setAttribute('aria-label', tag + ' 上移'); up.disabled = i === 0;
    var dn = btn('↓', 'mar-a-btn mar-a-btn-icon', function () { move(i, 1); }); dn.setAttribute('aria-label', tag + ' 下移'); dn.disabled = i === S.models.length - 1;
    var del = btn('刪除', 'mar-a-btn mar-a-btn-danger', function () { removeRule(i); }); del.setAttribute('aria-label', '刪除' + tag);
    ops.appendChild(up); ops.appendChild(dn); ops.appendChild(del);
    c7.appendChild(ops);
    return tr;
  }

  function changed() {
    S.dirty = true;
    setStatus('');
    runValidation();
    scheduleSim(250);
  }
  function setStatus(msg, kind) {
    var s = S.els.status; if (!s) return;
    s.textContent = msg || ''; s.className = 'mar-a-status' + (kind ? ' mar-a-status-' + kind : '');
  }
  function addRule() {
    if (S.models.length >= LIM.maxRules) return;
    S.models.push({ id: '', name: '新規則', enabled: true, dir: 'before', nRaw: '30', repeatRaw: '0', notify: ['owner'], condition: 'always', applies: 'all' });
    S.dirty = true; renderRules(); runValidation(); scheduleSim(0);
    var rows = S.els.tbody.querySelectorAll('.mar-a-row'); var last = rows[rows.length - 1];
    if (last) { var inp = last.querySelector('.mar-a-name'); if (inp) { inp.focus(); inp.select(); } }
  }
  function removeRule(i) {
    S.models.splice(i, 1); S.dirty = true; renderRules(); runValidation(); scheduleSim(0);
    var rows = S.els.tbody.querySelectorAll('.mar-a-row'); var t = rows[Math.min(i, rows.length - 1)];
    var f = t && t.querySelector('.mar-a-name'); if (f) f.focus(); else S.els.addBtn.focus();
  }
  function move(i, d) {
    var j = i + d; if (j < 0 || j >= S.models.length) return;
    var t = S.models[i]; S.models[i] = S.models[j]; S.models[j] = t;
    S.dirty = true; renderRules(); runValidation();
    var rows = S.els.tbody.querySelectorAll('.mar-a-row'); var f = rows[j] && rows[j].querySelectorAll('.mar-a-ops button')[d < 0 ? 0 : 1];
    if (f && !f.disabled) f.focus();
  }
  function sortRules() {
    S.models.sort(function (a, b) { var wa = whenDaysOf(a), wb = whenDaysOf(b); return (isNaN(wb) ? -9999 : wb) - (isNaN(wa) ? -9999 : wa); });
    S.dirty = true; renderRules(); runValidation(); scheduleSim(0);
  }
  function askDefaults() { S.els.confirm.hidden = false; var b = S.els.confirm.querySelector('button'); if (b) b.focus(); }
  function applyDefaults() {
    S.models = (S.defaults || []).map(fromApi);
    S.els.confirm.hidden = true; S.dirty = true;
    renderRules(); runValidation(); scheduleSim(0);
    setStatus('已把編輯區改為預設規則（尚未儲存）', 'info');
    S.els.saveBtn.focus();
  }

  function runValidation() {
    var v = validateRules(S.models);
    S.valid = v;
    var rows = S.els.tbody.querySelectorAll('.mar-a-row');
    rows.forEach(function (tr, i) {
      var e = v.perRule[i] || {};
      tr.querySelectorAll('.mar-a-ferr').forEach(function (s) {
        var k = s.getAttribute('data-err');
        s.textContent = e[k] || ''; s.hidden = !e[k];
      });
      var map = { name: '.mar-a-name', when: '.mar-a-n', repeat: '.mar-a-rep', notify: '.mar-a-notify' };
      Object.keys(map).forEach(function (k) {
        var t = tr.querySelector(map[k]); if (!t) return;
        var bad = !!(e[k] || (k === 'name' && e.dup));
        if (bad) t.setAttribute('aria-invalid', 'true'); else t.removeAttribute('aria-invalid');
      });
      tr.classList.toggle('mar-a-row-bad', Object.keys(e).length > 0);
    });
    var E = S.els;
    if (v.global) { E.errAll.textContent = v.global; E.errAll.hidden = false; }
    else if (v.count) { E.errAll.textContent = '有 ' + v.count + ' 個欄位需要修正（標紅的規則），修正前不能儲存，也無法試算。'; E.errAll.hidden = false; }
    else { E.errAll.textContent = ''; E.errAll.hidden = true; }
    return v;
  }

  /* ── 儲存 ── */
  function mountReview() {
    var E = S.els;
    if (!window.CfgReview || !S.boot.crApi) return;
    S.cr = window.CfgReview.create({ api: S.boot.crApi, onApplied: function () { reloadFromServer(); } });
    S.cr.paint(E.crHost);
    S.cr.refresh().then(function () { E.saveBtn.textContent = S.cr.enabled() ? S.cr.saveLabel('') : '儲存'; });
  }
  function save() {
    var E = S.els;
    if (S.saving) return;
    var v = runValidation();
    if (v.count || v.global) {
      setStatus('❌ 還有欄位需要修正', 'err');
      var bad = E.tbody.querySelector('[aria-invalid="true"]'); if (bad) bad.focus();
      return;
    }
    S.saving = true; E.saveBtn.disabled = true; setStatus('儲存中…', 'info');
    var body = { rules: S.models.map(toApi) };
    apiJson('PUT', '/api/admin/reminder-rules', body).then(function (r) {
      S.saving = false; E.saveBtn.disabled = false;
      var d = r.data || {};
      if (r.status === 202) {   // 兩人覆核啟用：只是提出，尚未生效；畫面改回伺服器目前實際生效的規則
        var msg = window.CfgReview ? window.CfgReview.proposalMessage(d) : (d.message || '');
        setStatus('📨 已送出待另一位管理員覆核。' + (msg ? ' ' + msg : ''), 'info');
        if (S.cr) S.cr.refresh();
        return reloadFromServer();
      }
      if (!r.ok) { setStatus('❌ 儲存失敗：' + (d.error || r.status) + (d.code ? '（' + d.code + '）' : ''), 'err'); return; }
      if (d.noChange) { setStatus('內容與目前生效的規則相同，沒有變更。', 'info'); return; }
      setStatus('✅ 已儲存（已寫入稽核）；下一次輪詢起生效', 'ok');
      return reloadFromServer().then(function () { setStatus('✅ 已儲存（已寫入稽核）；下一次輪詢起生效', 'ok'); });
    }).catch(function (e) {
      S.saving = false; E.saveBtn.disabled = false;
      if (e && e.message !== 'Session expired') setStatus('❌ 儲存失敗：' + (e.message || e), 'err');
    });
  }

  /* ───────── 試算器 ───────── */
  function buildSim(host) {
    var E = S.els;
    var card = el('section', 'mar-a-card mar-a-simcard');
    card.setAttribute('aria-labelledby', 'marSimTitle');
    var hd = el('div', 'mar-a-hd'); var t = el('h3', 'mar-a-title', '試算器'); t.id = 'marSimTitle'; hd.appendChild(t);
    hd.appendChild(el('span', 'mar-a-sub', '用「目前編輯中、尚未儲存」的規則試算，不會寄出任何通知。'));
    card.appendChild(hd);

    var ctl = el('div', 'mar-a-simctl');
    // 滑桿
    var g1 = el('div', 'mar-a-simgrp');
    var l1 = el('label', 'mar-a-lbl', '距離到期 N 天（負數＝已逾期）'); l1.setAttribute('for', 'marSimDays');
    E.simRange = document.createElement('input'); E.simRange.type = 'range'; E.simRange.id = 'marSimDays'; E.simRange.min = '-120'; E.simRange.max = '180'; E.simRange.step = '1'; E.simRange.value = '45';
    E.simRange.className = 'mar-a-range';
    E.simNum = document.createElement('input'); E.simNum.type = 'number'; E.simNum.min = '-400'; E.simNum.max = '400'; E.simNum.step = '1'; E.simNum.value = '45'; E.simNum.className = 'mar-a-inp mar-a-simnum';
    E.simNum.setAttribute('aria-label', '距離到期天數（數字輸入）');
    E.simDayText = el('span', 'mar-a-daytxt', dayText(45)); E.simDayText.setAttribute('aria-live', 'polite');
    E.simRange.addEventListener('input', function () { E.simNum.value = E.simRange.value; simInput(); });
    E.simNum.addEventListener('input', function () {
      var n = Number(E.simNum.value); if (!isFinite(n) || E.simNum.value === '') return;
      var c = Math.max(-400, Math.min(400, Math.round(n)));
      E.simRange.value = String(Math.max(Number(E.simRange.min), Math.min(Number(E.simRange.max), c))); simInput();
    });
    g1.appendChild(l1); g1.appendChild(E.simRange);
    var line = el('div', 'mar-a-inline'); line.appendChild(E.simNum); line.appendChild(E.simDayText);
    g1.appendChild(line);
    ctl.appendChild(g1);
    // 報價單狀態
    var g2 = document.createElement('fieldset'); g2.className = 'mar-a-simgrp mar-a-qs';
    g2.appendChild(el('legend', 'mar-a-lbl', '續約報價單狀態'));
    E.simQs = [['none', '尚未建立'], ['draft', '草稿'], ['sent', '已送簽'], ['norenew', '標記不續約']].map(function (o, i) {
      var lab = el('label', 'mar-a-radio');
      var r = document.createElement('input'); r.type = 'radio'; r.name = 'marSimQs'; r.value = o[0]; r.checked = i === 0;
      r.addEventListener('change', simInput);
      lab.appendChild(r); lab.appendChild(document.createTextNode(o[1])); g2.appendChild(lab); return r;
    });
    ctl.appendChild(g2);
    // 類型
    var g3 = el('div', 'mar-a-simgrp');
    var l3 = el('label', 'mar-a-lbl', '項目類型'); l3.setAttribute('for', 'marSimKind');
    E.simKind = document.createElement('select'); E.simKind.id = 'marSimKind'; E.simKind.className = 'mar-a-sel';
    [['ma', 'MA'], ['license', 'License']].forEach(function (o) { var op = document.createElement('option'); op.value = o[0]; op.textContent = o[1]; E.simKind.appendChild(op); });
    E.simKind.addEventListener('change', simInput);
    g3.appendChild(l3); g3.appendChild(E.simKind);
    ctl.appendChild(g3);
    // 真實納管項
    var g4 = el('div', 'mar-a-simgrp');
    var l4 = el('label', 'mar-a-lbl', '用真實納管項解析收件人（選填）'); l4.setAttribute('for', 'marSimItem');
    E.simItem = document.createElement('select'); E.simItem.id = 'marSimItem'; E.simItem.className = 'mar-a-sel';
    var op0 = document.createElement('option'); op0.value = ''; op0.textContent = '（不用，以角色示範）'; E.simItem.appendChild(op0);
    E.simItem.addEventListener('change', function () {
      var it = (S.items || []).filter(function (x) { return x.id === E.simItem.value; })[0];
      if (it) {
        if (it.termKind === 'ma' || it.termKind === 'license') E.simKind.value = it.termKind;
        if (typeof it.daysLeft === 'number') { var c = Math.max(-400, Math.min(400, it.daysLeft)); E.simNum.value = String(c); E.simRange.value = String(Math.max(-120, Math.min(180, c))); E.simDayText.textContent = dayText(c); }
      }
      simInput();
    });
    g4.appendChild(l4); g4.appendChild(E.simItem);
    ctl.appendChild(g4);
    card.appendChild(ctl);

    E.simErr = el('div', 'mar-a-error'); E.simErr.setAttribute('role', 'alert'); E.simErr.hidden = true; card.appendChild(E.simErr);

    // 時間軸
    var tlh = el('h4', 'mar-a-subtitle', '時間軸'); card.appendChild(tlh);
    E.tlWrap = el('div', 'mar-a-tlwrap');
    card.appendChild(E.tlWrap);
    E.tlList = el('ul', 'mar-a-tllist'); E.tlList.setAttribute('aria-label', '各規則在這個情境下的狀態');
    card.appendChild(E.tlList);

    // 各角色此刻會看到的提醒
    card.appendChild(el('h4', 'mar-a-subtitle', '各角色此刻會看到的提醒'));
    E.roles = el('div', 'mar-a-roles'); E.roles.setAttribute('aria-live', 'polite');
    card.appendChild(E.roles);
    host.appendChild(card);
  }
  function simInput() {
    var E = S.els;
    var c = Number(E.simNum.value); if (isFinite(c)) E.simDayText.textContent = dayText(Math.round(c));
    drawTimeline();
    scheduleSim(150);
  }
  function loadItems() {
    S.boot.fetch('/api/ma-items?filter=all').then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var items = d && Array.isArray(d.items) ? d.items.filter(function (x) { return x.id && x.status === 'tracking'; }) : [];
      items.sort(function (a, b) { return a.company < b.company ? -1 : a.company > b.company ? 1 : 0; });
      S.items = items.slice(0, 300);
      S.items.forEach(function (it) {
        var op = document.createElement('option'); op.value = it.id;
        op.textContent = it.company + '｜' + (it.product || '') + '（到期 ' + String(it.to).replace(/-/g, '/') + '）';
        S.els.simItem.appendChild(op);
      });
    }).catch(function () { /* 沒有清單就只能用角色示範 */ });
  }
  function scheduleSim(ms) {
    clearTimeout(S.simTimer);
    S.simTimer = setTimeout(runSim, ms);
  }
  function scenario() {
    var E = S.els;
    var qs = E.simQs.filter(function (r) { return r.checked; })[0].value;
    var dl = Math.round(Number(E.simNum.value));
    var o = { daysLeft: isFinite(dl) ? dl : 0, quoteState: qs === 'norenew' ? 'none' : qs, notRenew: qs === 'norenew', termKind: E.simKind.value };
    if (E.simItem.value) o.maItemId = E.simItem.value;
    return o;
  }
  function runSim() {
    var E = S.els; if (!E || !E.simRange) return;
    var v = S.valid || validateRules(S.models);
    if (v.count || v.global) { E.simErr.textContent = '規則有欄位需要修正，暫時無法試算。'; E.simErr.hidden = false; S.sim = null; drawTimeline(); paintRoles(); return; }
    E.simErr.hidden = true;
    var mine = ++S.simSeq;
    var sc = scenario();
    var body = Object.assign({}, sc, { rules: S.models.map(toApi) });
    apiJson('POST', '/api/admin/reminder-rules/simulate', body).then(function (r) {
      if (mine !== S.simSeq) return;   // 較新的請求已發出
      if (!r.ok) { E.simErr.textContent = '試算失敗：' + ((r.data && r.data.error) || r.status); E.simErr.hidden = false; S.sim = null; drawTimeline(); paintRoles(); return; }
      S.sim = r.data; S.simScenario = sc;
      drawTimeline(); paintRoles();
    }).catch(function () { /* session 過期由 adminFetch 處理 */ });
  }

  function paintRoles() {
    var E = S.els; E.roles.textContent = '';
    var sim = S.sim;
    if (!sim) { E.roles.appendChild(el('p', 'mar-a-muted', '（等待試算結果）')); return; }
    if (!sim.tracking) E.roles.appendChild(el('p', 'mar-a-note', '這個項目已標記為不續約（或不在追蹤中）：所有提醒都已停止。'));
    ROLES.forEach(function (role) {
      var hits = (sim.byRole && sim.byRole[role]) || [];
      var card = el('div', 'mar-a-role' + (hits.length ? ' mar-a-role-on' : ''));
      card.setAttribute('data-role', role);
      card.appendChild(el('h5', 'mar-a-role-h', ROLE_LABEL[role]));
      if (hits.length) {
        card.appendChild(el('p', 'mar-a-role-msg', simText(sim)));
        card.appendChild(el('p', 'mar-a-role-rules', '來自規則：' + hits.map(function (h) { return '「' + h.name + '」'; }).join('、')));
        var rec = sim.recipients && sim.recipients[role];
        if (rec && rec.length) card.appendChild(el('p', 'mar-a-role-who', '收件人：' + rec.map(function (x) { return x.label || x.username; }).join('、')));
        if (role === 'secretary') card.appendChild(el('p', 'mar-a-role-who', '秘書的提醒不含金額與合約編號。'));
      } else {
        card.appendChild(el('p', 'mar-a-role-none', '此刻不會收到提醒'));
      }
      E.roles.appendChild(card);
    });
  }

  /* ── 時間軸（SVG）── */
  var NS = 'http://www.w3.org/2000/svg';
  function svg(tag, attrs, cls, text) {
    var e = document.createElementNS(NS, tag);
    if (cls) e.setAttribute('class', cls);
    Object.keys(attrs || {}).forEach(function (k) { e.setAttribute(k, String(attrs[k])); });
    if (text != null) e.textContent = String(text);
    return e;
  }
  function drawTimeline() {
    var E = S.els; E.tlWrap.textContent = ''; E.tlList.textContent = '';
    var sc = scenario();   // 今天指示線隨滑桿立即移動；各規則的成立狀態用最近一次伺服器試算結果
    var items = [];
    S.models.forEach(function (m, i) {
      var w = whenDaysOf(m);
      if (isNaN(w)) return;
      var st = S.sim && S.sim.rules ? S.sim.rules[i] : null;
      items.push({ m: m, i: i, w: w, st: st });
    });
    var today = sc.daysLeft;
    var tMin = -120, tMax = 60;
    items.forEach(function (x) { tMin = Math.min(tMin, -x.w - 10); tMax = Math.max(tMax, -x.w + 10); });
    tMin = Math.min(tMin, -today - 10); tMax = Math.max(tMax, -today + 10);
    var W = 760, H = 210, padL = 28, padR = 28, axisY = 130;
    function X(t) { return padL + (t - tMin) / (tMax - tMin) * (W - padL - padR); }
    var s = svg('svg', { viewBox: '0 0 ' + W + ' ' + H, width: W, height: H, role: 'img', 'aria-label': '提醒時間軸：橫軸為距離到期的天數，左邊是到期前、右邊是到期後；共 ' + items.length + ' 條規則，今天距離到期 ' + today + ' 天' }, 'mar-a-svg');
    s.appendChild(svg('line', { x1: padL, y1: axisY, x2: W - padR, y2: axisY }, 'mar-a-axis'));
    var step = (tMax - tMin) > 260 ? 60 : 30;
    for (var t = Math.ceil(tMin / step) * step; t <= tMax; t += step) {
      s.appendChild(svg('line', { x1: X(t), y1: axisY - 4, x2: X(t), y2: axisY + 4 }, 'mar-a-tickmark'));
      s.appendChild(svg('text', { x: X(t), y: axisY + 20, 'text-anchor': 'middle' }, 'mar-a-txt mar-a-txt-sm', t === 0 ? '到期日' : (t < 0 ? String(-t) : '+' + t)));
    }
    s.appendChild(svg('text', { x: W - padR, y: H - 6, 'text-anchor': 'end' }, 'mar-a-txt mar-a-txt-sm', '距離到期天數（左：到期前　右：到期後）'));
    s.appendChild(svg('line', { x1: X(0), y1: 28, x2: X(0), y2: axisY + 8 }, 'mar-a-expiry'));
    // 重複刻度
    items.forEach(function (x) {
      var rp = x.w <= 0 ? Number(x.m.repeatRaw) : 0;
      if (rp > 0) for (var k = 1, tt = -x.w + rp; k <= 14 && tt <= tMax; k++, tt += rp) s.appendChild(svg('line', { x1: X(tt), y1: axisY - 6, x2: X(tt), y2: axisY + 6 }, 'mar-a-rtick' + (x.m.enabled ? '' : ' mar-a-rtick-off')));
    });
    // 規則標記（標籤分層避免重疊）
    var lanes = [];
    items.slice().sort(function (a, b) { return -a.w - -b.w; }).forEach(function (x) {
      var cx = X(-x.w);
      var label = x.m.dir === 'before' ? '到期前 ' + x.m.nRaw + ' 天' : (x.m.dir === 'on' ? '到期當天' : '到期後 ' + x.m.nRaw + ' 天');   // 圖上只標時間點；規則名稱在 <title>（滑過顯示）與下方文字清單
      var wpx = Math.max(40, Array.from(label).length * 12 + 6);
      var lane = 0;
      while (lanes[lane] !== undefined && cx - wpx / 2 < lanes[lane]) lane++;
      lanes[lane] = cx + wpx / 2;
      var cls = 'mar-a-dot ' + (!x.m.enabled ? 'mar-a-dot-off' : (x.st && x.st.active ? 'mar-a-dot-active' : 'mar-a-dot-idle'));
      var ly = axisY - 22 - lane * 18;
      s.appendChild(svg('line', { x1: cx, y1: ly + 3, x2: cx, y2: axisY - 7 }, 'mar-a-stem'));
      s.appendChild(svg('circle', { cx: cx, cy: axisY, r: 7 }, cls));
      var tx = svg('text', { x: cx, y: ly, 'text-anchor': 'middle' }, 'mar-a-txt', label);
      tx.appendChild(svg('title', {}, null, String(x.m.name)));
      s.appendChild(tx);
    });
    // 今天
    var tx = X(-today);
    s.appendChild(svg('line', { x1: tx, y1: 14, x2: tx, y2: axisY + 30 }, 'mar-a-today'));
    s.appendChild(svg('text', { x: Math.min(W - padR - 4, Math.max(padL + 4, tx)), y: 12, 'text-anchor': tx > W - 120 ? 'end' : (tx < 120 ? 'start' : 'middle') }, 'mar-a-txt mar-a-txt-today', '今天（' + dayText(today) + '）'));
    E.tlWrap.appendChild(s);
    var leg = el('div', 'mar-a-legend');
    [['mar-a-key-active', '目前成立'], ['mar-a-key-idle', '目前不成立／尚未到達'], ['mar-a-key-off', '規則已停用']].forEach(function (k) {
      var sp = el('span', 'mar-a-keyitem'); sp.appendChild(el('span', 'mar-a-key ' + k[0])); sp.appendChild(document.createTextNode(k[1])); leg.appendChild(sp);
    });
    E.tlWrap.appendChild(leg);
    // 文字版（螢幕閱讀器與不看圖的人）
    items.slice().sort(function (a, b) { return b.w - a.w; }).forEach(function (x) {
      var li = el('li', 'mar-a-tlitem');
      var state = !x.m.enabled ? '已停用' : (x.st ? (x.st.active ? '目前成立' : '目前不成立：' + (REASON_TEXT[x.st.reason] || x.st.reason)) : '等待試算');
      li.appendChild(el('strong', null, x.m.name || '（未命名）'));
      li.appendChild(document.createTextNode('　' + whenText(x.m) + '　→　'));
      li.appendChild(el('span', 'mar-a-tlstate ' + (x.st && x.st.active && x.m.enabled ? 'mar-a-tlstate-on' : 'mar-a-tlstate-off'), state));
      E.tlList.appendChild(li);
    });
  }

  /* ───────── 對外 ───────── */
  function probe(boot) {
    if (S.installed || S.booting) return Promise.resolve(S.installed);
    if (S.probed404) return Promise.resolve(false);
    S.booting = true;
    return boot.fetch('/api/admin/reminder-rules').then(function (r) {
      S.booting = false;
      if (r.status === 404) { S.probed404 = true; return false; }   // 開關關閉：確定結果，不再重打
      if (!r.ok) return false;
      install(boot);
      return true;
    }).catch(function () { S.booting = false; return false; });
  }

  window.MaRemAdmin = {
    probe: probe, validateRules: validateRules, toApi: toApi, fromApi: fromApi, whenText: whenText, simText: simText,
    _state: S, version: 'p3b',
  };
})();
