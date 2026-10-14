/* ITTS CRM — 職能專屬首頁：行銷首頁（#marketingHomeView）、顧問首頁（#consultHomeView）
 *
 * 資料一律來自既有 API（不另算商業邏輯、不編造數字）：
 *   行銷：GET /api/leads、/api/companies、/api/campaigns
 *   顧問：GET /api/quotations、/api/quotations/inbox
 * 玻璃模式（body.glass）：重用 glass-dash.js 匯出的 GlassKit（卡片舞台／補間／時間膠囊）。
 * 非玻璃（?glass=0）：同樣資料畫成一般卡片格（style.css 的 .rh-*），不依賴 glass.css。
 * 對外：window.RoleHomes.loadMarketing() / loadConsult()（由 app.js showSection 呼叫）。
 */
(function () {
  'use strict';
  if (window.RoleHomes) return;

  var KIT = function () { return window.GlassKit || null; };
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fmt(n) { return Math.round(n || 0).toLocaleString(); }
  function glassOn() { return !!(document.body && document.body.classList.contains('glass')) && !!KIT(); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function jget(url) {
    return fetch(url, { credentials: 'include' }).then(function (r) {
      if (!r.ok) { var e = new Error(url + ' → ' + r.status); e.status = r.status; throw e; }
      return r.json();
    });
  }
  function ts(s) { var t = s ? new Date(s).getTime() : NaN; return isNaN(t) ? null : t; }

  /* ───────── 共用：把「模組」掛成玻璃舞台或一般卡片格 ─────────
   * home = { view, id, mods:[{id,title,color,tag}], cols, rowH, minFocusH, extraBar:fn(barEl) }
   * 之後以 home.paint(id, {num, vis, detail, tag}) 填內容。 */
  function makeHome(cfg) {
    var home = { cfg: cfg, built: null, root: null, stage: null, cards: {}, last: {}, cap: null };
    home.mount = function () {
      var view = $(cfg.view);
      if (!view) return false;
      var wantGlass = glassOn();
      if (home.built === wantGlass && home.root && home.root.parentNode === view) return true;
      home.unmount();
      var root = document.createElement('div');
      root.id = cfg.id;
      view.appendChild(root);
      home.root = root;
      if (wantGlass) {
        var K = KIT();
        root.className = 'gd-root';
        var gooId = cfg.id + 'Goo';
        root.innerHTML = K.gooDefs(gooId) + '<div class="gd-bar"><div class="gd-bar-note">' + esc(cfg.note || '') + '</div><div class="gd-bar-r" id="' + cfg.id + 'Bar"></div></div>';
        view.classList.add('gd-on');
        home.stage = K.makeStage({
          id: cfg.id + 'Stage', mods: cfg.mods,
          cols: function (W) { return W >= 560 ? 2 : 1; },
          rowH: function (W) { return W >= 560 ? (cfg.rowH || 236) : (cfg.rowH || 236) - 10; },
          thumbCols: function (W) { return W >= 560 ? 3 : 2; },
          minFocusH: function () { return cfg.minFocusH || 640; }
        });
        root.appendChild(home.stage.el);
        if (cfg.capsule) {
          home.cap = K.makeCapsule(cfg.capsule.initial(), cfg.capsule.onSelect, gooId);
          root.querySelector('#' + cfg.id + 'Bar').appendChild(home.cap.el);
          var lbl = document.createElement('span'); lbl.className = 'gd-bar-lbl'; lbl.textContent = '資料期間';
          root.querySelector('#' + cfg.id + 'Bar').insertBefore(lbl, home.cap.el);
        }
        if (window.ResizeObserver) {
          home.ro = new ResizeObserver(function () { if (home.stage && home.stage.el.clientWidth !== home.stage.W) home.stage.layout(); });
          home.ro.observe(home.stage.el);
        }
        home.stage.layout();
      } else {
        root.className = 'rh-root';
        var bar = '<div class="rh-bar"><span class="rh-note">' + esc(cfg.note || '') + '</span><span id="' + cfg.id + 'Bar" class="rh-period"></span></div>';
        root.innerHTML = bar + '<div class="rh-grid">' + cfg.mods.map(function (m) {
          return '<div class="rh-card" data-id="' + m.id + '"><div class="rh-hd"><span class="rh-dot" style="background:' + (m.plainColor || '#1a73e8') + '"></span><span class="rh-title">' + esc(m.title) + '</span><span class="rh-tag"></span></div>' +
            '<div class="rh-num"></div><div class="rh-vis"></div><div class="rh-detail"></div></div>';
        }).join('') + '</div>';
        if (cfg.capsule) {
          var pb = root.querySelector('#' + cfg.id + 'Bar');
          [['week', '本週'], ['month', '本月'], ['quarter', '本季'], ['year', '本年']].forEach(function (p) {
            var b = document.createElement('button'); b.type = 'button'; b.className = 'rh-pbtn'; b.textContent = p[1]; b.dataset.k = p[0];
            b.addEventListener('click', function () { cfg.capsule.onSelect(p[0]); [].forEach.call(pb.children, function (x) { x.classList.toggle('on', x === b); }); });
            if (p[0] === cfg.capsule.initial()) b.classList.add('on');
            pb.appendChild(b);
          });
        }
        [].forEach.call(root.querySelectorAll('.rh-card'), function (c) { home.cards[c.dataset.id] = c; });
      }
      home.built = wantGlass;
      Object.keys(home.last).forEach(function (id) { home.paint(id, home.last[id]); });
      return true;
    };
    home.unmount = function () {
      if (home.ro) { home.ro.disconnect(); home.ro = null; }
      if (home.stage && KIT()) KIT().dropStage(home.stage);
      if (home.root && home.root.parentNode) home.root.parentNode.removeChild(home.root);
      var view = $(cfg.view); if (view) view.classList.remove('gd-on');
      home.root = null; home.stage = null; home.cap = null; home.cards = {}; home.built = null;
    };
    home.paint = function (id, o) {
      home.last[id] = o;
      if (home.built === true && home.stage) {
        home.stage.put(id, o.num || '', o.vis || '', '');
        if (o.detail != null) home.stage.detail(id).innerHTML = (o.vis && !o.noVisInDetail ? '<div class="rh-vis-d">' + o.vis.replace('preserveAspectRatio="none"', 'preserveAspectRatio="xMidYMid meet"') + '</div>' : '') + o.detail;   // 展開後摘要區會隱藏，圖表一併放進詳情
        if (o.tag != null) home.stage.tag(id, o.tag);
      } else if (home.built === false) {
        var c = home.cards[id]; if (!c) return;
        c.querySelector('.rh-num').innerHTML = o.num || '';
        c.querySelector('.rh-vis').innerHTML = o.vis || '';
        if (o.detail != null) c.querySelector('.rh-detail').innerHTML = (o.vis && !o.noVisInDetail ? '<div class="rh-vis-d">' + o.vis.replace('preserveAspectRatio="none"', 'preserveAspectRatio="xMidYMid meet"') + '</div>' : '') + o.detail;
        if (o.tag != null) c.querySelector('.rh-tag').textContent = o.tag;
      }
    };
    home.relayout = function () { if (home.stage) home.stage.layout(); };
    // 深／淺色與玻璃開關：重新掛載（資料從 last 重畫，不重打 API）
    new MutationObserver(function () {
      var v = $(cfg.view);
      if (!v || !home.root) return;
      if (home.built !== glassOn()) { home.mount(); }
    }).observe(document.body, { attributes: true, attributeFilter: ['class'] });
    return home;
  }

  function errCard(e) {
    if (e && e.status === 403) return '<div class="rh-warn">此帳號沒有讀取這項資料的權限（403）。</div>';
    return '<div class="rh-warn">載入失敗：' + esc(e && e.message) + '</div>';
  }
  function rowBtn(attrs, inner) { return '<div class="rh-row" role="button" tabindex="0" ' + attrs + '>' + inner + '</div>'; }

  /* ═════════════════ 行銷首頁 ═════════════════ */
  var PK_NOTE = { week: '本週（週一～週日，每日）', month: '本月（每日）', quarter: '本季（每週）', year: '本年（每月）' };
  var MK = { period: 'year', leads: null, companies: null, campaigns: null, err: {}, loaded: false };
  var MK_MODS = [
    { id: 'K1', title: 'Lead 累積增長曲線', color: 'var(--g-accent)', plainColor: '#1a73e8' },
    { id: 'K2', title: '本期新增 Lead／已轉商機', color: '#34d399', plainColor: '#1e8e3e' },
    { id: 'K3', title: '企業主檔產業分佈', color: '#a78bfa', plainColor: '#7c3aed' },
    { id: 'K4', title: '行銷活動', color: '#f59e0b', plainColor: '#e37400' }
  ];
  var LEAD_LABEL = { new: '新進', assigned: '已指派', contacted: '已聯繫', converted: '已轉商機', disqualified: '不合格' };
  var LEAD_COL = { new: '#1a73e8', assigned: '#e37400', contacted: '#9c27b0', converted: '#1e8e3e', disqualified: '#888' };
  var CAMP_LABEL = { planned: '計畫中', active: '進行中', completed: '已完成', cancelled: '已取消' };
  var IND_COLORS = ['#5b9dff', '#2dd4bf', '#a78bfa', '#fb7185', '#fbbf24', '#34d399', '#60a5fa', '#f472b6'];
  var NONE_COL = '#f59e0b';

  var mkHome = makeHome({
    view: 'marketingHomeView', id: 'mkRoot', mods: MK_MODS, rowH: 240, minFocusH: 620,
    note: '行銷看的是 Lead 與企業主檔，不含業績數字；點卡片放大看詳情' + '',
    capsule: { initial: function () { return MK.period; }, onSelect: function (k) { MK.period = k; mkRender(); } }
  });

  function periodRange(p) {
    var now = new Date(), y = now.getFullYear(), m = now.getMonth(), s, e;
    if (p === 'year') { s = new Date(y, 0, 1); e = new Date(y + 1, 0, 1); }
    else if (p === 'quarter') { var q = Math.floor(m / 3); s = new Date(y, q * 3, 1); e = new Date(y, q * 3 + 3, 1); }
    else if (p === 'month') { s = new Date(y, m, 1); e = new Date(y, m + 1, 1); }
    else { var dow = (now.getDay() + 6) % 7; s = new Date(y, m, now.getDate() - dow); e = new Date(s.getFullYear(), s.getMonth(), s.getDate() + 7); }
    return [s, e];
  }
  /* 時間桶：每桶 [起, 訖)，只到「現在」為止（不畫未來） */
  function buckets(p) {
    var r = periodRange(p), now = Date.now(), out = [], d = new Date(r[0]);
    while (d < r[1] && d.getTime() <= now) {
      var n;
      if (p === 'year') n = new Date(d.getFullYear(), d.getMonth() + 1, 1);
      else if (p === 'quarter') n = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 7);
      else n = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
      if (n > r[1]) n = new Date(r[1]);
      out.push({ s: new Date(d), e: n, label: p === 'year' ? (d.getMonth() + 1) + '月' : pad(d.getMonth() + 1) + '/' + pad(d.getDate()) });
      d = n;
    }
    return out;
  }

  function curveSvg(pts, labels) {
    var W = 300, H = 104, L = 6, R = 6, T = 10, B = 16;
    var n = pts.length;
    if (!n) return '<div class="rh-empty">此期間尚無資料點</div>';
    var mn = Math.min.apply(null, pts), mx = Math.max.apply(null, pts), realMax = mx;
    if (mx === mn) { mx = mn + 1; }
    var X = function (i) { return n === 1 ? W / 2 : L + (W - L - R) * i / (n - 1); };
    var Y = function (v) { return T + (H - T - B) * (1 - (v - mn) / (mx - mn)); };
    var line = pts.map(function (v, i) { return (i ? 'L' : 'M') + X(i).toFixed(1) + ' ' + Y(v).toFixed(1); }).join(' ');
    var area = line + ' L' + X(n - 1).toFixed(1) + ' ' + (H - B) + ' L' + X(0).toFixed(1) + ' ' + (H - B) + ' Z';
    var dots = pts.map(function (v, i) { return n <= 16 ? '<circle cx="' + X(i).toFixed(1) + '" cy="' + Y(v).toFixed(1) + '" r="2.4" class="rh-dotc"/>' : ''; }).join('');
    return '<svg class="rh-curve" viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="none" role="img" aria-label="Lead 累積增長曲線">' +
      '<path d="' + area + '" class="rh-area"/><path d="' + line + '" class="rh-line" fill="none"/>' + dots +
      '<text x="' + L + '" y="' + (H - 3) + '" class="rh-axis">' + esc(labels[0]) + '</text>' +
      '<text x="' + (W - R) + '" y="' + (H - 3) + '" text-anchor="end" class="rh-axis">' + esc(labels[n - 1]) + '</text>' +
      '<text x="' + (W - R) + '" y="' + (T + 2) + '" text-anchor="end" class="rh-axis">' + fmt(realMax) + '</text></svg>';
  }

  function donutSvg(segs) {
    var tt = segs.reduce(function (a, s) { return a + s.v; }, 0), off = 0, ci = '';
    if (!tt) return '<div class="rh-empty">尚無企業主檔資料</div>';
    segs.forEach(function (s) {
      var l = s.v / tt * 100;
      ci += '<circle r="15.9" cx="21" cy="21" fill="none" stroke="' + s.c + '" stroke-width="7" stroke-dasharray="' + l.toFixed(2) + ' ' + (100 - l).toFixed(2) + '" stroke-dashoffset="' + (25 - off).toFixed(2) + '"/>';
      off += l;
    });
    return '<svg class="rh-donut" viewBox="0 0 42 42" role="img" aria-label="產業分佈">' + ci + '</svg>';
  }

  function mkRender() {
    if (!mkHome.mount()) return;
    var P = MK.period, bk = buckets(P), r = periodRange(P);
    var tag = { week: '本週', month: '本月', quarter: '本季', year: '本年' }[P];

    /* K1 / K2：Lead */
    if (MK.err.leads) {
      mkHome.paint('K1', { num: '<span class="rh-n rh-mute">—</span>', vis: errCard(MK.err.leads), detail: '', tag: '' });
      mkHome.paint('K2', { num: '<span class="rh-n rh-mute">—</span>', vis: errCard(MK.err.leads), detail: '', tag: '' });
    } else if (MK.leads) {
      var L = MK.leads, noTs = 0, created = [], conv = [];
      L.forEach(function (l) {
        var c = ts(l.createdAt); if (c == null) noTs++; else created.push(c);
        if (l.status === 'converted') { var v = ts(l.convertedAt); if (v != null) conv.push(v); }
      });
      var cum = bk.map(function (b) { return created.filter(function (c) { return c < b.e.getTime(); }).length; });
      var inP = created.filter(function (c) { return c >= r[0].getTime() && c < r[1].getTime(); }).length;
      var convP = conv.filter(function (c) { return c >= r[0].getTime() && c < r[1].getTime(); }).length;
      var convAll = L.filter(function (l) { return l.status === 'converted'; }).length;
      var before = created.filter(function (c) { return c < r[0].getTime(); }).length;
      mkHome.paint('K1', {
        num: '<span class="rh-n">' + fmt(created.length) + '</span><small>累積 Lead（本期新增 ' + fmt(inP) + '）</small>',
        vis: curveSvg(cum, bk.map(function (b) { return b.label; })),
        tag: tag + '・' + PK_NOTE[P].replace(/^.*（/, '').replace(/）$/, ''),
        detail: '<div class="rh-note">曲線＝各時間點「建立時間 ≤ 該點」的 Lead 累計數（含期間開始前已有的 ' + fmt(before) + ' 筆）；資料來源：Lead 建立時間（createdAt）。' +
          (noTs ? '另有 ' + noTs + ' 筆 Lead 沒有建立時間，未納入曲線。' : '') +
          '<br>範圍：此帳號可見的 Lead（行銷帳號只看到自己建立的；管理員／主管看到全部）。</div>' +
          '<div class="rh-sub">各時間點累計</div><div class="rh-kvs">' + bk.map(function (b, i) { return '<span>' + esc(b.label) + '<b>' + fmt(cum[i]) + '</b></span>'; }).join('') + '</div>'
      });
      var st = {}; L.forEach(function (l) { st[l.status] = (st[l.status] || 0) + 1; });
      var stRows = Object.keys(st).sort(function (a, b) { return st[b] - st[a]; }).map(function (k) {
        return '<div class="rh-bar-row"><span class="rh-bl">' + esc(LEAD_LABEL[k] || k) + '</span><span class="rh-bt"><i style="width:' + (st[k] / Math.max(1, L.length) * 100) + '%;background:' + (LEAD_COL[k] || '#888') + '"></i></span><b>' + st[k] + '</b></div>';
      }).join('');
      var recent = L.slice().sort(function (a, b) { return (ts(b.createdAt) || 0) - (ts(a.createdAt) || 0); }).slice(0, 8);
      mkHome.paint('K2', {
        num: '<span class="rh-n">' + fmt(inP) + '</span><small>本期新增 Lead　已轉商機 ' + fmt(convP) + '</small>',
        vis: '<div class="rh-kvs"><span>本期新增<b>' + fmt(inP) + '</b></span><span>本期轉商機<b>' + fmt(convP) + '</b></span><span>累計已轉<b>' + fmt(convAll) + '</b></span><span>累計轉換率<b>' + (L.length ? Math.round(convAll / L.length * 100) + '%' : '—') + '</b></span></div>',
        tag: tag,
        detail: '<div class="rh-note">本期＝' + ymd(r[0]) + ' ～ ' + ymd(new Date(r[1].getTime() - 86400000)) + '；「新增」依建立時間、「轉商機」依轉換時間（convertedAt）。</div>' +
          '<div class="rh-sub">Lead 狀態分佈（全部）</div>' + (stRows || '<div class="rh-empty">尚無 Lead</div>') +
          '<div class="rh-sub">最近 Lead</div>' + (recent.length ? recent.map(function (l) {
            return rowBtn('data-go="leads"', '<span class="rh-r1">' + esc(l.company || l.contactName || '—') + '</span><span class="rh-r2">' + esc(LEAD_LABEL[l.status] || l.status || '') + '</span><span class="rh-r3">' + esc((l.createdAt || '').slice(0, 10)) + '</span>');
          }).join('') : '<div class="rh-empty">尚無 Lead</div>')
      });
    } else {
      mkHome.paint('K1', { num: '<span class="rh-n rh-mute">載入中…</span>', vis: '', detail: '', tag: '' });
      mkHome.paint('K2', { num: '<span class="rh-n rh-mute">載入中…</span>', vis: '', detail: '', tag: '' });
    }

    /* K3：企業主檔產業 */
    if (MK.err.companies) {
      mkHome.paint('K3', { num: '<span class="rh-n rh-mute">—</span>', vis: errCard(MK.err.companies), detail: '', tag: '' });
    } else if (MK.companies) {
      var cnt = {}, none = 0;
      MK.companies.forEach(function (c) { var i = (c.industry || '').trim(); if (i) cnt[i] = (cnt[i] || 0) + 1; else none++; });
      var names = Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; });
      var segs = names.slice(0, 7).map(function (k, i) { return { k: k, v: cnt[k], c: IND_COLORS[i % IND_COLORS.length] }; });
      var restV = names.slice(7).reduce(function (a, k) { return a + cnt[k]; }, 0);
      if (restV) segs.push({ k: '其他（' + (names.length - 7) + ' 類）', v: restV, c: '#94a3b8' });
      if (none) segs.push({ k: '未分類', v: none, c: NONE_COL });
      var total = MK.companies.length;
      var leg = segs.slice(0, 4).map(function (s) { return '<div class="rh-leg"><span class="rh-lgd" style="background:' + s.c + '"></span>' + esc(s.k) + '<b>' + s.v + '</b></div>'; }).join('');
      var allRows = names.map(function (k, i) {
        return '<div class="rh-bar-row"><span class="rh-bl" title="' + esc(k) + '">' + esc(k) + '</span><span class="rh-bt"><i style="width:' + (cnt[k] / Math.max(1, names.length ? cnt[names[0]] : 1) * 100) + '%;background:' + IND_COLORS[i % IND_COLORS.length] + '"></i></span><b>' + cnt[k] + '</b></div>';
      }).join('');
      mkHome.paint('K3', {
        num: '<span class="rh-n">' + fmt(total) + '</span><small>企業主檔家數　<span class="rh-none">未分類 ' + fmt(none) + '</span></small>',
        vis: '<div class="rh-donut-row">' + donutSvg(segs) + '<div class="rh-mini-leg">' + leg + '</div></div>',
        tag: '全部・不受期間影響',
        detail: '<div class="rh-none-box ' + (none ? 'on' : '') + '"><b>未分類 ' + fmt(none) + ' 家</b>' + (total ? '（' + (none > 0 && none / total < 0.01 ? '<1' : Math.round(none / total * 100)) + '%）' : '') +
          '<button type="button" class="rh-link" data-go="companyMaster-none">前往企業主檔處理未分類 →</button></div>' +
          '<div class="rh-sub">各產業家數（' + names.length + ' 類）</div>' + (allRows || '<div class="rh-empty">尚無已分類的企業</div>') +
          '<div class="rh-note">企業主檔無建立時間可靠欄位，故產業分佈不依期間切片。</div>'
      });
    } else {
      mkHome.paint('K3', { num: '<span class="rh-n rh-mute">載入中…</span>', vis: '', detail: '', tag: '' });
    }

    /* K4：行銷活動 */
    if (MK.err.campaigns) {
      mkHome.paint('K4', { num: '<span class="rh-n rh-mute">—</span>', vis: errCard(MK.err.campaigns), detail: '', tag: '' });
    } else if (MK.campaigns) {
      var C = MK.campaigns, byS = {}, lc = 0, cc = 0;
      C.forEach(function (c) { byS[c.status || 'planned'] = (byS[c.status || 'planned'] || 0) + 1; lc += c.leadCount || 0; cc += c.convertedCount || 0; });
      var top = C.slice().sort(function (a, b) { return (b.leadCount || 0) - (a.leadCount || 0); }).slice(0, 8);
      mkHome.paint('K4', {
        num: '<span class="rh-n">' + fmt(C.length) + '</span><small>活動　帶來 Lead ' + fmt(lc) + '</small>',
        vis: '<div class="rh-kvs">' + ['active', 'planned', 'completed', 'cancelled'].map(function (k) { return '<span>' + CAMP_LABEL[k] + '<b>' + (byS[k] || 0) + '</b></span>'; }).join('') + '</div>',
        tag: '全部活動',
        detail: '<div class="rh-note">累計 Lead ' + fmt(lc) + '，已轉商機 ' + fmt(cc) + '（' + (lc ? Math.round(cc / lc * 100) + '%' : '—') + '）；行銷帳號只看到自己建立的活動。</div>' +
          '<div class="rh-sub">活動（依 Lead 數）</div>' + (top.length ? top.map(function (c) {
            return rowBtn('data-go="campaigns"', '<span class="rh-r1">' + esc(c.name || '—') + '</span><span class="rh-r2">' + esc(CAMP_LABEL[c.status] || c.status || '') + '</span><span class="rh-r3">Lead ' + (c.leadCount || 0) + '・轉 ' + (c.convertedCount || 0) + '</span>');
          }).join('') : '<div class="rh-empty">尚無行銷活動</div>')
      });
    } else {
      mkHome.paint('K4', { num: '<span class="rh-n rh-mute">載入中…</span>', vis: '', detail: '', tag: '' });
    }
    mkHome.relayout();
  }

  function goNav(go) {
    if (typeof showSection !== 'function') return;
    if (go === 'companyMaster-none') { try { _cmIndustryFilter = '__none__'; } catch (e) { /* ignore */ } showSection('companyMaster'); return; }
    showSection(go);
  }
  document.addEventListener('click', function (e) {
    var g = e.target.closest && e.target.closest('[data-go]');
    if (!g || !g.closest('#marketingHomeView')) return;
    e.stopPropagation(); goNav(g.getAttribute('data-go'));
  });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    var g = e.target.closest && e.target.closest('.rh-row');
    if (g) { e.preventDefault(); g.click(); }
  });

  function loadMarketing() {
    mkRender();
    var role = window._myRole || '';
    var jobs = [
      jget('/api/leads').then(function (v) { MK.leads = v; MK.err.leads = null; }, function (e) { MK.err.leads = e; }),
      jget('/api/companies').then(function (v) { MK.companies = v; MK.err.companies = null; }, function (e) { MK.err.companies = e; }),
      jget('/api/campaigns').then(function (v) { MK.campaigns = v; MK.err.campaigns = null; }, function (e) { MK.err.campaigns = e; })
    ];
    return Promise.all(jobs).then(function () { MK.loaded = true; mkRender(); setTimeout(function () { mkHome.relayout(); }, 60); return role; });
  }

  /* ═════════════════ 顧問首頁 ═════════════════
   * 狀態模型（讀 lib/quoteRoutes.js 的結果）：
   *   待我填成本：costFlow.state==='requested' 且 costBy===我，且（inbox.costs 的規則）報價單尚未送簽/核准 → 直接用 /api/quotations/inbox 的 costs
   *   我已填成本：costFlow.state==='filled'（filledAt）
   *   顧問「看不到」approval（伺服器對顧問回 approval:null），故無法顯示送簽後的關卡、被駁回、核准完成——頁面明講，不猜。 */
  var CS = { rows: null, inbox: null, err: null, me: '' };
  var CS_MODS = [
    { id: 'S1', title: '待我填成本', color: '#f59e0b', plainColor: '#e37400' },
    { id: 'S2', title: '我已填成本', color: '#34d399', plainColor: '#1e8e3e' },
    { id: 'S3', title: '近期收件／完成', color: 'var(--g-accent)', plainColor: '#1a73e8' },
    { id: 'S4', title: '被指派的報價單總覽', color: '#a78bfa', plainColor: '#7c3aed' }
  ];
  var csHome = makeHome({
    view: 'consultHomeView', id: 'csRoot', mods: CS_MODS, rowH: 236, minFocusH: 600,
    note: '顧問在報價單流程中負責「填成本」；點卡片放大，點報價單列開啟填寫'
  });
  function daysAgo(t) { return t == null ? null : Math.max(0, Math.floor((Date.now() - t) / 86400000)); }
  function qRow(q, kind) {
    var t = kind === 'wait' ? ts(q.requestedAt || (q.costFlow && q.costFlow.requestedAt)) : ts(q.costFlow && q.costFlow.filledAt);
    var d = t == null ? '' : ymd(new Date(t));
    var age = kind === 'wait' && t != null ? '已等 ' + daysAgo(t) + ' 天' : '';
    return rowBtn('data-q="' + esc(q.id) + '"', '<span class="rh-r1">' + esc(q.quoteNo || '—') + '　' + esc(q.company || '') + '</span><span class="rh-r2">' + esc(q.projectName || '') + (q.ownerName ? '・業務 ' + esc(q.ownerName) : '') + '</span><span class="rh-r3">' + esc(age || d) + '</span>');
  }
  function csRender() {
    if (!csHome.mount()) return;
    if (CS.err) {
      CS_MODS.forEach(function (m) { csHome.paint(m.id, { num: '<span class="rh-n rh-mute">—</span>', vis: errCard(CS.err), detail: '', tag: '' }); });
      return;
    }
    if (!CS.rows) {
      CS_MODS.forEach(function (m) { csHome.paint(m.id, { num: '<span class="rh-n rh-mute">載入中…</span>', vis: '', detail: '', tag: '' }); });
      return;
    }
    var me = CS.me, mine = CS.rows.filter(function (q) { return q.costBy === me; });
    var byId = {}; mine.forEach(function (q) { byId[q.id] = q; });
    var waiting = ((CS.inbox && CS.inbox.costs) || []).map(function (c) { var q = byId[c.id] || {}; return { id: c.id, quoteNo: c.quoteNo, company: c.company, projectName: q.projectName, ownerName: q.ownerName, requestedAt: c.requestedAt }; })
      .sort(function (a, b) { return (ts(a.requestedAt) || 0) - (ts(b.requestedAt) || 0); });
    var filled = mine.filter(function (q) { return q.costFlow && q.costFlow.state === 'filled'; })
      .sort(function (a, b) { return (ts(b.costFlow.filledAt) || 0) - (ts(a.costFlow.filledAt) || 0); });
    var other = mine.length - waiting.length - filled.length;
    var oldest = waiting.length ? daysAgo(ts(waiting[0].requestedAt)) : null;
    csHome.paint('S1', {
      num: '<span class="rh-n' + (waiting.length ? ' rh-hot' : '') + '">' + waiting.length + '</span><small>張等你填成本' + (oldest != null ? '（最久 ' + oldest + ' 天）' : '') + '</small>',
      vis: waiting.length ? '<div class="rh-rows">' + waiting.slice(0, 3).map(function (q) { return qRow(q, 'wait'); }).join('') + '</div>' : '<div class="rh-empty">目前沒有待填的報價單</div>',
      tag: '即時', noVisInDetail: true,
      detail: '<div class="rh-note">條件：業務指定你為成本填寫人、尚未送簽或核准（與右上通知鈴的「成本請求」同一規則）。</div>' +
        (waiting.length ? waiting.map(function (q) { return qRow(q, 'wait'); }).join('') : '<div class="rh-empty">目前沒有待填的報價單</div>')
    });
    csHome.paint('S2', {
      num: '<span class="rh-n">' + filled.length + '</span><small>張已填成本</small>',
      vis: filled.length ? '<div class="rh-rows">' + filled.slice(0, 3).map(function (q) { return qRow(q, 'done'); }).join('') + '</div>' : '<div class="rh-empty">尚無已填成本的報價單</div>',
      tag: '成本狀態＝已填', noVisInDetail: true,
      detail: '<div class="rh-warn">送簽後的「目前關卡」與「是否被駁回／核准」：伺服器不會把簽核資料（approval）提供給顧問角色，所以這裡只能顯示「你已填成本」，無法顯示它目前簽到哪一關。若要顯示，需由系統端另外提供給顧問的唯讀欄位。</div>' +
        (filled.length ? filled.map(function (q) { return qRow(q, 'done'); }).join('') : '<div class="rh-empty">尚無已填成本的報價單</div>')
    });
    var now = new Date(), m0 = new Date(now.getFullYear(), now.getMonth(), 1).getTime(), d30 = Date.now() - 30 * 86400000;
    var recvM = mine.filter(function (q) { var t = ts(q.costFlow && q.costFlow.requestedAt); return t != null && t >= m0; }).length;
    var doneM = filled.filter(function (q) { var t = ts(q.costFlow.filledAt); return t != null && t >= m0; }).length;
    var done30 = filled.filter(function (q) { var t = ts(q.costFlow.filledAt); return t != null && t >= d30; });
    csHome.paint('S3', {
      num: '<span class="rh-n">' + doneM + '</span><small>本月完成填寫（本月收到 ' + recvM + ' 張）</small>',
      vis: '<div class="rh-kvs"><span>本月收到<b>' + recvM + '</b></span><span>本月完成<b>' + doneM + '</b></span><span>近 30 天完成<b>' + done30.length + '</b></span></div>',
      tag: '本月',
      detail: '<div class="rh-note">「收到」依成本請求時間（costFlow.requestedAt）、「完成」依成本填寫完成時間（costFlow.filledAt）。</div>' +
        '<div class="rh-sub">近 30 天完成</div>' + (done30.length ? done30.map(function (q) { return qRow(q, 'done'); }).join('') : '<div class="rh-empty">近 30 天沒有完成紀錄</div>')
    });
    csHome.paint('S4', {
      num: '<span class="rh-n">' + mine.length + '</span><small>張被指派給你（成本填寫人）</small>',
      vis: '<div class="rh-kvs"><span>待填<b>' + waiting.length + '</b></span><span>已填<b>' + filled.length + '</b></span><span>其他<b>' + other + '</b></span></div>',
      tag: '全部',
      detail: '<div class="rh-note">「其他」＝已送簽或已核准但成本請求狀態未回到待填（細節對顧問不可見）。<br>這裡<b>沒有「待我簽」區塊</b>：程式內顧問沒有簽核關卡（簽核只有 mgr2／mgr1／總經理／董事長／董事會代核），「填成本」不是簽核。</div>' +
        (mine.length ? mine.map(function (q) { var st = q.costFlow && q.costFlow.state; return qRow(q, st === 'filled' ? 'done' : 'wait'); }).join('') : '<div class="rh-empty">尚無被指派的報價單</div>')
    });
    csHome.relayout();
  }
  document.addEventListener('click', function (e) {
    var r = e.target.closest && e.target.closest('#consultHomeView [data-q]');
    if (!r) return;
    e.stopPropagation();
    var id = r.getAttribute('data-q');
    if (typeof openQuoteFromNotification === 'function') openQuoteFromNotification('quote_cost_request', id);
  });
  function loadConsult() {
    CS.me = window._myUsername || '';
    csRender();
    return Promise.all([jget('/api/quotations'), jget('/api/quotations/inbox')]).then(function (a) {
      CS.rows = a[0]; CS.inbox = a[1]; CS.err = null; csRender(); setTimeout(function () { csHome.relayout(); }, 60);
    }, function (e) { CS.err = e; csRender(); });
  }

  window.RoleHomes = { loadMarketing: loadMarketing, loadConsult: loadConsult };
})();
