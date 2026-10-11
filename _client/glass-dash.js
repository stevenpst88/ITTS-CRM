/* ITTS CRM — 玻璃首頁（業務首頁 #dashboardView、主管首頁 #managerHomeView）
 *
 * 只在 <body class="glass"> 時啟動；沒有 glass（?glass=0）時整支 no-op，舊版首頁原樣運作。
 * 設計：舊區塊（年度目標卡、季度卡、月度計畫表、漸進圖表…）仍是「唯一資料來源」，由 app.js 既有函式照舊渲染。
 *   本檔只做兩件事：
 *   1. 把舊區塊「搬進」玻璃卡片的詳情區（reparent，id 與事件都在），收合時卡片只顯示摘要；
 *   2. 摘要數字一律從舊區塊「已渲染的結果」或同一批既有函式讀取（ownScopeOpps…），不另寫商業邏輯。
 * 關閉玻璃（setGlass(false)）時 destroy()：所有搬動以 undo stack 反向還原。
 *
 * 互動：卡片滑鼠光斑＋邊緣高亮＋≤4° 傾斜；點卡片放大成詳情、其餘縮成底部縮圖列（點縮圖切換、✕／Esc 收回）；
 *   時間膠囊（本週／本月／本季／本年，metaball）重新切片「商機推進／漏斗／月度計畫」，其他卡片標示「全年／不受期間影響」。
 * 版面：卡片以絕對座標＋ transition 排版（layout()），只有展開／收合那一下會重排。
 * 對外：window.GlassDash（業務）、window.GlassMgr（主管；loadManagerHome 末端呼叫 GlassMgr.update(d)）。
 */
(function () {
  'use strict';
  if (window.__gdLoaded) return;
  window.__gdLoaded = true;

  var GAP = 12;
  var SC = { D: '#8d8d8d', C: '#1a73e8', B: '#f57c00', A: '#d32f2f' };            // 商機階段語意色（固定，不隨主題）
  var SN = { D: 'D 靜止中', C: 'C Pipeline', B: 'B Upside', A: 'A Commit' };
  var ST = ['D', 'C', 'B', 'A'];
  var PK = ['week', 'month', 'quarter', 'year'];
  var PER = { week: '本週', month: '本月', quarter: '本季', year: '本年' };
  var STAGES = [];          // 所有舞台（Esc 收合用）

  /* ───────── 小工具 ───────── */
  function $(id) { return document.getElementById(id); }
  function el(tag, cls) { var e = document.createElement(tag); if (cls) e.className = cls; return e; }
  function glassOn() { return !!(document.body && document.body.classList.contains('glass')); }
  function reduce() { try { return matchMedia('(prefers-reduced-motion: reduce)').matches; } catch (e) { return false; } }
  function fine() { try { return matchMedia('(hover: hover) and (pointer: fine)').matches; } catch (e) { return false; } }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fmt(n) { return Math.round(n || 0).toLocaleString(); }
  /* 讀舊區塊顯示出來的數字文字：'1,234 K'→1234、'(30)'→-30、'38%'→38、'—'→null */
  function pnum(t) {
    t = String(t == null ? '' : t).replace(/●/g, '').trim();
    if (!t || t === '—' || t === '--') return null;
    var neg = /^\(/.test(t) || /^-/.test(t);
    var n = parseFloat(t.replace(/[^0-9.]/g, ''));
    return isNaN(n) ? null : (neg ? -n : n);
  }
  /* _countUp 會把終值寫在 dataset.to（動畫途中 textContent 是中間值，終值才是真值） */
  function finalOf(e) {
    if (!e) return null;
    var tx = e.textContent || '';
    if (!/\d/.test(tx)) return null;
    var d = e.dataset && e.dataset.to;
    if (d != null && d !== '' && !isNaN(Number(d))) return Number(d);
    return pnum(tx);
  }
  function sum(a, from, to) { var s = 0; for (var i = from; i < to; i++) s += a[i] || 0; return s; }
  function visibleEl(e) { return !!e && e.style.display !== 'none'; }

  /* 搬動紀錄：各首頁自己一條 undo stack，destroy 時反向還原 */
  function mv(log, node, parent) {
    if (!node || !parent) return;
    log.push({ node: node, parent: node.parentNode, next: node.nextSibling });
    parent.appendChild(node);
  }
  function undoMoves(log) {
    while (log.length) {
      var m = log.pop();
      try { m.parent.insertBefore(m.node, m.next && m.next.parentNode === m.parent ? m.next : null); } catch (e) { /* ignore */ }
    }
  }
  function modalOpen() {
    var list = document.querySelectorAll('.modal-overlay');
    for (var i = 0; i < list.length; i++) { if (list[i].offsetParent !== null && getComputedStyle(list[i]).display !== 'none') return true; }
    return false;
  }

  /* ───────── 舞台（卡片 ＋ 展開 ＋ 縮圖列）───────── */
  function makeStage(cfg) {
    var st = { id: cfg.id, el: el('div', 'gd-stage'), mods: cfg.mods, focus: null, hidden: {}, cards: {}, W: 0, laid: {} };
    st.el.id = cfg.id;
    cfg.mods.forEach(function (m) {
      var c = el('div', 'gd-card');
      c.dataset.id = m.id;
      c.innerHTML = '<div class="gd-tilt"><div class="gd-in">' +
        '<div class="gd-hd"><span class="gd-dot"></span><span class="gd-title"></span><span class="gd-tag"></span>' +
        '<button type="button" class="gd-x" aria-label="收合詳情">✕</button></div>' +
        '<div class="gd-num"></div><div class="gd-vis"></div><div class="gd-slice"></div><div class="gd-detail"></div></div></div>';
      c.querySelector('.gd-dot').style.background = m.color;
      c.querySelector('.gd-title').textContent = m.title;
      st.el.appendChild(c);
      st.cards[m.id] = c;
      var tilt = c.querySelector('.gd-tilt');
      c.addEventListener('pointermove', function (e) {
        if (e.pointerType === 'touch') return;
        var r = c.getBoundingClientRect(), x = e.clientX - r.left, y = e.clientY - r.top;
        tilt.style.setProperty('--mx', x + 'px'); tilt.style.setProperty('--my', y + 'px');
        if (reduce() || !fine() || c.classList.contains('gd-main')) return;
        var ry = ((x / r.width) - .5) * 8, rx = -((y / r.height) - .5) * 8;     // ≤ ±4°
        tilt.style.transition = 'transform .12s linear';
        tilt.style.transform = 'rotateX(' + rx.toFixed(2) + 'deg) rotateY(' + ry.toFixed(2) + 'deg)';
      });
      c.addEventListener('pointerleave', function () {
        tilt.style.transition = ''; tilt.style.transform = '';
      });
      c.addEventListener('click', function (e) {
        if (e.target.closest('.gd-x')) { st.setFocus(null); return; }
        if (st.focus !== m.id) st.setFocus(m.id);
      });
      c.addEventListener('keydown', function (e) {
        if ((e.key === 'Enter' || e.key === ' ') && e.target === c && st.focus !== m.id) { e.preventDefault(); st.setFocus(m.id); }
      });
    });

    st.layout = function () {
      var W = st.el.clientWidth;
      if (!W) return;
      st.W = W;
      var vis = st.mods.filter(function (m) { return !st.hidden[m.id]; });
      var n = vis.length;
      if (st.focus && st.hidden[st.focus]) st.focus = null;
      var cols = Math.max(1, Math.min(cfg.cols(W, n), n || 1)), rowH = cfg.rowH(W);
      var rows = Math.ceil(n / cols), gridH = n ? rows * rowH + (rows - 1) * GAP : 0;
      var th = 76, tc = Math.max(1, Math.min(cfg.thumbCols(W, n), Math.max(1, n - 1)));
      var trows = Math.ceil(Math.max(0, n - 1) / tc), strip = trows * th + Math.max(0, trows - 1) * GAP;
      var tb = document.querySelector('.topbar'), avail = window.innerHeight - (tb ? tb.offsetHeight : 56) - 24;
      var Hf = Math.max(cfg.minFocusH(W), Math.min(avail, 900));
      st.el.style.height = (st.focus ? Hf : gridH) + 'px';
      st.el.classList.toggle('gd-has-focus', !!st.focus);
      st.mods.forEach(function (m) {
        var c = st.cards[m.id], r, i = vis.indexOf(m);
        if (i < 0) { c.style.display = 'none'; st.laid[m.id] = false; return; }
        if (!st.focus) {
          var w = (W - (cols - 1) * GAP) / cols;
          r = { l: (i % cols) * (w + GAP), t: Math.floor(i / cols) * (rowH + GAP), w: w, h: rowH };
          c.classList.remove('gd-main', 'gd-thumb');
        } else if (m.id === st.focus) {
          r = { l: 0, t: 0, w: W, h: Hf - GAP - strip };
          c.classList.add('gd-main'); c.classList.remove('gd-thumb');
        } else {
          var others = vis.filter(function (x) { return x.id !== st.focus; }), k = others.indexOf(m), w3 = (W - (tc - 1) * GAP) / tc;
          r = { l: (k % tc) * (w3 + GAP), t: Hf - strip + Math.floor(k / tc) * (th + GAP), w: w3, h: th };
          c.classList.add('gd-thumb'); c.classList.remove('gd-main');
        }
        var fresh = !st.laid[m.id];
        if (fresh) { c.style.transition = 'none'; }
        c.style.display = '';
        c.style.left = r.l + 'px'; c.style.top = r.t + 'px'; c.style.width = r.w + 'px'; c.style.height = r.h + 'px';
        if (fresh) { void c.offsetWidth; c.style.transition = ''; st.laid[m.id] = true; }
        var main = c.classList.contains('gd-main');
        c.setAttribute('aria-expanded', main ? 'true' : 'false');
        c.setAttribute('role', main ? 'region' : 'button');
        c.setAttribute('aria-label', m.title + (main ? '（詳情）' : '，按 Enter 展開詳情'));
        c.tabIndex = main ? -1 : 0;
        c.querySelector('.gd-x').tabIndex = main ? 0 : -1;
      });
    };
    st.setHidden = function (map) {
      var changed = false;
      st.mods.forEach(function (m) { var h = !!map[m.id]; if (!!st.hidden[m.id] !== h) { st.hidden[m.id] = h; changed = true; } });
      if (changed) st.layout();
    };
    st.setFocus = function (id) {
      var prev = st.focus;
      st.focus = id || null;
      st.layout();
      if (st.focus && st.focus !== prev) {
        var go = function () { try { st.el.scrollIntoView({ behavior: reduce() ? 'auto' : 'smooth', block: 'start' }); } catch (e) { /* ignore */ } };
        go(); setTimeout(function () { if (st.focus) go(); }, 600);      // 舞台高度有 transition，捲動範圍要等長高後才夠
      }
      if (cfg.onFocus) cfg.onFocus(st.focus);
    };
    st.put = function (id, num, vis, slice) {
      var c = st.cards[id]; if (!c) return;
      if (num != null) c.querySelector('.gd-num').innerHTML = num;
      if (vis != null) c.querySelector('.gd-vis').innerHTML = vis;
      if (slice != null) c.querySelector('.gd-slice').innerHTML = slice;
    };
    st.tag = function (id, text) {
      var c = st.cards[id]; if (!c) return;
      var t = c.querySelector('.gd-tag');
      if (t.textContent !== (text || '')) t.textContent = text || '';
      t.style.display = text ? '' : 'none';
    };
    st.detail = function (id) { return st.cards[id].querySelector('.gd-detail'); };
    STAGES.push(st);
    return st;
  }

  /* 摘要數字的補間（卡片各自一條 rAF；減少動態時直接換值） */
  function makeTweener(paint) {
    var S = {};
    return {
      S: S,
      set: function (id, to, meta, animate) {
        var s = S[id] || (S[id] = { nums: null, meta: null, raf: 0, sig: '' });
        var sig = JSON.stringify(to) + '|' + JSON.stringify(meta);
        if (sig === s.sig) return;
        s.sig = sig; s.meta = meta;
        if (s.raf) { cancelAnimationFrame(s.raf); s.raf = 0; }
        if (s.fin) { clearTimeout(s.fin); s.fin = 0; }
        var from = s.nums && s.nums.length === to.length ? s.nums : (animate ? to.map(function () { return 0; }) : null);
        if (!animate || reduce() || !from) { s.nums = to.slice(); paint(id, s); return; }
        var t0 = null, f = from.slice();
        // 保險：分頁在背景時 rAF 會暫停，補間永遠走不完 → 逾時直接落到終值
        s.fin = setTimeout(function () { s.fin = 0; if (s.raf) { cancelAnimationFrame(s.raf); s.raf = 0; } s.nums = to.slice(); paint(id, s); }, 1300);
        (function step(ts) {
          if (t0 === null) t0 = ts;
          var k = Math.min(1, (ts - t0) / 800), e = 1 - Math.pow(1 - k, 3);
          s.nums = to.map(function (v, i) { return f[i] + (v - f[i]) * e; });
          paint(id, s);
          if (k < 1) s.raf = requestAnimationFrame(step); else s.raf = 0;
        })(performance.now());
      },
      drop: function () { Object.keys(S).forEach(function (k) { if (S[k].raf) cancelAnimationFrame(S[k].raf); if (S[k].fin) clearTimeout(S[k].fin); }); S = this.S = {}; }
    };
  }

  /* 時間膠囊（metaball）：選取後其餘融合，onSelect(key) 在「選了不同期間」時呼叫 */
  function makeCapsule(initial, onSelect) {
    var tf = el('div', 'gd-tf collapsed'), goo = el('div', 'gd-goo'), W = 64, G = 8;
    tf.setAttribute('role', 'group'); tf.setAttribute('aria-label', '資料期間');
    tf.appendChild(goo);
    var sel = initial, expanded = false, busy = false, parts = [];
    PK.forEach(function (k) {
      var o = el('div', 'gd-opt'); goo.appendChild(o);
      var g = el('div', 'gd-g'); tf.appendChild(g);
      var l = el('button', 'gd-lab'); l.type = 'button'; l.textContent = PER[k]; tf.appendChild(l);
      l.addEventListener('click', function (e) { e.stopPropagation(); onTap(k); });
      parts.push({ k: k, o: o, g: g, l: l });
    });
    function pos() {
      parts.forEach(function (p, i) {
        var x = expanded ? i * (W + G) : 0;
        p.o.style.left = p.g.style.left = p.l.style.left = x + 'px';
        p.g.classList.toggle('sel', p.k === sel); p.l.classList.toggle('sel', p.k === sel);
        p.l.tabIndex = (expanded || p.k === sel) ? 0 : -1;
        p.l.setAttribute('aria-pressed', p.k === sel ? 'true' : 'false');
      });
      tf.classList.toggle('collapsed', !expanded);
    }
    function run(ms, fn) {
      busy = true; tf.classList.add('anim', 'busy'); fn();
      setTimeout(function () { tf.classList.remove('anim', 'busy', 'collapsing'); busy = false; }, reduce() ? 0 : ms);
    }
    function onTap(k) {
      if (busy) return;
      if (!expanded) { run(640, function () { expanded = true; pos(); }); return; }
      var changed = k !== sel; sel = k;
      run(560, function () { tf.classList.add('collapsing'); expanded = false; pos(); });
      if (changed) onSelect(k);
    }
    pos();
    return {
      el: tf,
      isOpen: function () { return expanded; },
      close: function () { if (expanded && !busy) run(560, function () { tf.classList.add('collapsing'); expanded = false; pos(); }); },
      value: function () { return sel; }
    };
  }

  /* ═════════════════ 業務首頁 ═════════════════ */
  var GD = { moves: [], built: false, period: 'year', stage: null, tw: null, cap: null, obs: null, viewObs: null, root: null, timer: 0 };

  var SALES_MODS = [
    { id: 'A', title: '年度業績目標', color: 'var(--g-accent)' },
    { id: 'B', title: '季度業績進度', color: '#facc15' },
    { id: 'C', title: '月度業績計畫', color: '#34d399' },
    { id: 'D', title: '商機推進進度', color: '#f57c00' },
    { id: 'E', title: '商機漏斗', color: '#d32f2f' },
    { id: 'F', title: '客戶與產業分佈', color: '#a78bfa' }
  ];

  function periodRange(p) {
    if (p === 'year') return null;                           // 本年＝與舊版相同語意（不做日期過濾）
    var now = new Date(), y = now.getFullYear(), m = now.getMonth();
    if (p === 'quarter') { var q = Math.floor(m / 3); return [new Date(y, q * 3, 1), new Date(y, q * 3 + 3, 1)]; }
    if (p === 'month') return [new Date(y, m, 1), new Date(y, m + 1, 1)];
    var dow = (now.getDay() + 6) % 7, s = new Date(y, m, now.getDate() - dow);        // 本週：週一～週日
    return [s, new Date(s.getFullYear(), s.getMonth(), s.getDate() + 7)];
  }
  function periodNote(p) {
    var r = periodRange(p), pad = function (n) { return (n < 10 ? '0' : '') + n; };
    if (!r) return '全部商機（與舊版相同）';
    var a = r[0], b = new Date(r[1].getTime() - 86400000);
    return '預計簽約日 ' + a.getFullYear() + '-' + pad(a.getMonth() + 1) + '-' + pad(a.getDate()) + ' ～ ' + b.getFullYear() + '-' + pad(b.getMonth() + 1) + '-' + pad(b.getDate());
  }
  function oppSlice(p) {
    var r = periodRange(p);
    var list = (typeof ownScopeOpps === 'function') ? ownScopeOpps() : [];       // 與舊版商機推進／漏斗同一個資料來源
    var cnt = [0, 0, 0, 0], amt = [0, 0, 0, 0];
    list.forEach(function (o) {
      var i = ST.indexOf(o.stage);
      if (i < 0) return;
      if (r) {
        if (!o.expectedDate) return;
        var d = new Date(o.expectedDate);
        if (isNaN(d) || d < r[0] || d >= r[1]) return;
      }
      cnt[i]++; amt[i] += parseFloat(o.amount) || 0;
    });
    return cnt.concat(amt);
  }

  /* —— 各卡片「摘要資料」：讀舊區塊已渲染的結果 —— */
  function calcA() {
    var amt = $('targetAmountDisplay'), ach = $('targetAchievedDisplay'), rate = $('targetRateDisplay');
    if (!amt) return null;
    var t = finalOf(amt), a = finalOf(ach), r = finalOf(rate);
    var y = new Date().getFullYear(), s0 = new Date(y, 0, 1), s1 = new Date(y + 1, 0, 1);
    var elapsed = (Date.now() - s0) / (s1 - s0) * 100;
    var has = t != null && r != null;
    return { nums: [has ? r : 0, a || 0, has ? t : 0, elapsed], meta: { has: has, msg: (amt.textContent || '').trim() || '尚未設定' } };
  }
  function calcB() {
    var cards = [].slice.call(document.querySelectorAll('#quarterGrid .q-card'));
    if (cards.length < 4) return null;
    var rate = [], tgt = [], ach = [], lab = [], fut = [], cur = -1;
    cards.slice(0, 4).forEach(function (c, i) {
      var v = c.querySelectorAll('.q-row-val');
      var t = pnum(v[0] && v[0].textContent), a = pnum(v[1] && v[1].textContent);
      tgt.push(t || 0); ach.push(a == null ? 0 : a); fut.push(a == null);
      rate.push(t > 0 && a != null ? a / t * 100 : 0);
      var lb = c.querySelector('.q-label'); lab.push(lb ? lb.textContent.trim() : 'Q' + (i + 1));
      if (c.classList.contains('q-current')) cur = i;
    });
    return { nums: rate.concat(tgt, ach), meta: { lab: lab, fut: fut, cur: cur } };
  }
  function calcC() {
    var wrap = $('monthBudgetWrap');
    var tables = [].slice.call(wrap.querySelectorAll('.mb-table'));
    if (!tables.length) return { nums: [], meta: { p: GD.period, msg: (wrap.textContent || '').replace(/\s+/g, ' ').replace('📅 月度業績計畫', '').trim() } };
    var out = { rb: [], ra: [], gb: [], ga: [] };
    for (var i = 0; i < 12; i++) { out.rb[i] = 0; out.ra[i] = 0; out.gb[i] = 0; out.ga[i] = 0; }
    var key = { '收入金額': 'rb', '收入實際': 'ra', '毛利金額': 'gb', '毛利實際': 'ga' };
    tables.forEach(function (tb) {
      [].forEach.call(tb.querySelectorAll('tbody tr'), function (tr) {
        var cells = [].slice.call(tr.children);
        var k = key[(cells[0] && cells[0].textContent || '').trim()];
        if (!k) return;
        var m = cells.slice(1).filter(function (td) { return td.classList.contains('mb-td') && !td.classList.contains('mb-q-col') && !td.classList.contains('mb-total-col'); });
        m.forEach(function (td, j) { if (j < 12) out[k][j] += pnum(td.textContent) || 0; });
      });
    });
    return { nums: out.ra.concat(out.rb, out.ga, out.gb), meta: { p: GD.period, tables: tables.length } };
  }
  function calcD() { return { nums: oppSlice(GD.period), meta: {} }; }
  function calcF() {
    var co = pnum($('statCompanies') && $('statCompanies').textContent) || 0, ct = pnum($('statContacts') && $('statContacts').textContent) || 0;
    var rows = [].slice.call(document.querySelectorAll('#industryLegend .chart-legend-row')), lab = [], col = [], cnt = [];
    rows.forEach(function (r) {
      var l = r.querySelector('.chart-legend-label'), c = r.querySelector('.chart-legend-count'), d = r.querySelector('.chart-legend-dot');
      lab.push(l ? l.textContent.trim() : ''); cnt.push(pnum(c && c.textContent) || 0); col.push(d ? d.style.background || '#888' : '#888');
    });
    return { nums: [co, ct].concat(cnt), meta: { lab: lab, col: col } };
  }

  function tagFor(id, p) {
    if (id === 'A' || id === 'B') return p === 'year' ? '全年' : '全年・不受期間影響';
    if (id === 'F') return p === 'year' ? '全部' : '不受期間影響';
    if (id === 'C') return { year: '全年', quarter: '本季', month: '本月', week: '以月計（不分週）' }[p];
    return p === 'year' ? '全部商機' : '依預計簽約日・' + PER[p];
  }

  /* —— 繪製（nums 可能是補間中的中間值）—— */
  function kv(label, val, cls) { return '<div class="gd-kv-i">' + label + '<b' + (cls ? ' class="' + cls + '"' : '') + '>' + val + '</b></div>'; }
  function bars(vals, colors, max, labels, opts) {
    opts = opts || {};
    return vals.map(function (v, i) {
      return '<div class="gd-col' + (opts.dim && opts.dim[i] ? ' gd-dim' : '') + (opts.cur === i ? ' gd-cur' : '') + '"><span>' + (opts.top ? opts.top[i] : '') + '</span><i style="height:' + Math.max(2, Math.min(100, v / max * 100)) + '%;background:' + colors[i] + '"></i><em>' + (labels ? labels[i] : '') + '</em></div>';
    }).join('');
  }

  function paintSales(id, s) {
    var st = GD.stage, n = s.nums, M = s.meta || {};
    if (id === 'A') {
      if (!M.has) {
        st.put('A', '<span class="gd-n gd-mute">' + esc(M.msg) + '</span>', '<div class="gd-kv">' + kv('已達成', fmt(n[1]) + ' K') + '</div>', '');
        return;
      }
      var r = Math.max(0, Math.min(100, n[0]));
      st.put('A', '<span class="gd-n">' + Math.round(n[0]) + '%</span><small>達成率</small>',
        '<div class="gd-pbar"><i style="width:' + r + '%"></i><u style="left:' + Math.min(100, n[3]).toFixed(1) + '%" title="今年已過 ' + Math.round(n[3]) + '%"></u></div>' +
        '<div class="gd-kv">' + kv('已達成', fmt(n[1]) + ' K', 'gd-pos') + kv('目標', fmt(n[2]) + ' K') + '</div>' +
        '<div class="gd-note">標線＝今年時間進度 ' + Math.round(n[3]) + '%</div>', '');
    } else if (id === 'B') {
      var rates = n.slice(0, 4), cur = M.cur >= 0 ? M.cur : 0;
      var mx = Math.max(100, Math.max.apply(null, rates)) * 1.1;
      var colors = rates.map(function (v) { return v >= 100 ? '#10b981' : 'var(--g-accent)'; });
      st.put('B', '<span class="gd-n">' + (M.fut[cur] ? '—' : Math.round(rates[cur]) + '%') + '</span><small>' + esc(M.lab[cur]) + ' 達成</small>',
        '<div class="gd-cols">' + bars(rates, colors, mx, M.lab, { dim: M.fut, cur: cur }) + '</div>', '');
    } else if (id === 'C') {
      if (!n.length) { st.put('C', '<span class="gd-n gd-mute">' + esc(M.msg || '尚未設定月度預算') + '</span>', '', ''); return; }
      var now = new Date(), mo = now.getMonth(), p = GD.period, a, b;
      if (p === 'year') { a = 0; b = 12; } else if (p === 'quarter') { a = Math.floor(mo / 3) * 3; b = a + 3; } else { a = mo; b = mo + 1; }
      var ra = n.slice(0, 12), rb = n.slice(12, 24), ga = n.slice(24, 36), gb = n.slice(36, 48);
      var act = sum(ra, a, b), bud = sum(rb, a, b), gma = sum(ga, a, b), gmb = sum(gb, a, b);
      var mxc = Math.max.apply(null, ra.concat(rb)) * 1.1 || 1;
      var cols = ra.map(function (v, i) {
        var bv = Math.max(0, rb[i]);
        return '<div class="gd-col2' + (i >= a && i < b ? '' : ' gd-dim') + '"><i style="height:' + Math.max(0, v) / mxc * 100 + '%"></i><u style="bottom:' + bv / mxc * 100 + '%"></u></div>';
      }).join('');
      st.put('C', '<span class="gd-n">' + fmt(act) + ' K</span><small>收入實際</small>', '<div class="gd-cols2">' + cols + '</div>',
        '<div class="gd-slice-in">' + kv('收入預算', fmt(bud) + ' K') + kv('收入實際', fmt(act) + ' K', 'gd-pos') +
        kv('收入達成', bud > 0 ? Math.round(act / bud * 100) + '%' : '—') + kv('毛利預算', fmt(gmb) + ' K') + kv('毛利實際', fmt(gma) + ' K', 'gd-pos') +
        kv('毛利達成', gmb > 0 ? Math.round(gma / gmb * 100) + '%' : '—') + '</div>' +
        '<div class="gd-note">統計範圍：' + (p === 'year' ? '全年 1–12 月' : p === 'quarter' ? (a + 1) + '–' + b + ' 月' : (a + 1) + ' 月' + (p === 'week' ? '（月度資料不分週，顯示本月）' : '')) + '（月度預算／認列，與下方完整計畫表同源）</div>');
    } else if (id === 'D') {
      var cnt = n.slice(0, 4), am = n.slice(4, 8), tot = sum(cnt, 0, 4);
      st.put('D', '<span class="gd-n">$ ' + fmt(sum(am, 0, 4)) + ' K</span><small>共 ' + Math.round(tot) + ' 筆</small>',
        '<div class="gd-cols">' + bars(cnt, ST.map(function (k) { return SC[k]; }), Math.max.apply(null, cnt.concat([1])) * 1.15, ST, { top: cnt.map(function (v) { return Math.round(v); }) }) + '</div>',
        '<div class="gd-slice-in">' + ST.map(function (k, j) { return kv(SN[k] + '（' + Math.round(cnt[j]) + ' 筆）', '$ ' + fmt(am[j]) + ' K', 'gd-st-' + k); }).join('') + '</div>' +
        '<div class="gd-note">' + esc(periodNote(GD.period)) + '</div>');
    } else if (id === 'E') {
      var c4 = n.slice(0, 4), total = sum(c4, 0, 4), y = 0, fs = '';
      if (total <= 0.5) {
        st.put('E', '<span class="gd-n gd-mute">尚無資料</span>', '', '<div class="gd-note">' + esc(periodNote(GD.period)) + '</div>');
        return;
      }
      var FW = [96, 78, 60, 44, 32], fh = 14;   // 固定漸縮外形（不依筆數畫寬度，避免資料偏斜時漏斗變形）；筆數以文字標示
      ST.forEach(function (k, j) {
        var w = FW[j], w2 = FW[j + 1];
        var x1 = 50 - w / 2, x2 = 50 + w / 2, x3 = 50 + w2 / 2, x4 = 50 - w2 / 2;
        fs += '<polygon points="' + x1.toFixed(1) + ',' + y + ' ' + x2.toFixed(1) + ',' + y + ' ' + x3.toFixed(1) + ',' + (y + fh) + ' ' + x4.toFixed(1) + ',' + (y + fh) + '" fill="' + SC[k] + '" stroke="' + SC[k] + '" stroke-width="1.6" stroke-linejoin="round" opacity=".94"/>' +
          '<text x="50" y="' + (y + fh / 2 + 2) + '" text-anchor="middle" font-size="5.4" font-weight="600" fill="#fff">' + k + '　' + Math.round(c4[j]) + '</text>';
        y += fh + 2.5;
      });
      st.put('E', '<span class="gd-n">' + Math.round(c4[3] / total * 100) + '%</span><small>A｜Commit 占比</small>',
        '<svg viewBox="0 0 100 66" preserveAspectRatio="xMidYMid meet" role="img" aria-label="商機漏斗">' + fs + '</svg>',
        '<div class="gd-slice-in">' + ST.map(function (k, j) { return kv(SN[k], Math.round(c4[j]) + ' 家', 'gd-st-' + k) + ''; }).join('') + '</div>' +
        '<div class="gd-note">各階段占比：' + ST.map(function (k, j) { return k + ' ' + Math.round(c4[j] / total * 100) + '%'; }).join('　') + '　｜　' + esc(periodNote(GD.period)) + '</div>');
    } else if (id === 'F') {
      var inds = n.slice(2), tt = sum(inds, 0, inds.length), off = 0, ci = '';
      inds.forEach(function (v, k) {
        var l = tt > 0 ? v / tt * 100 : 0;
        ci += '<circle r="15.9" cx="21" cy="21" fill="none" stroke="' + (M.col[k] || '#888') + '" stroke-width="7" stroke-dasharray="' + l.toFixed(2) + ' ' + (100 - l).toFixed(2) + '" stroke-dashoffset="' + (25 - off).toFixed(2) + '"/>';
        off += l;
      });
      var top3 = inds.map(function (v, k) { return { v: v, k: k }; }).slice(0, 3).map(function (o) {
        return '<div class="gd-leg"><span class="gd-lg-dot" style="background:' + (M.col[o.k] || '#888') + '"></span>' + esc(M.lab[o.k] || '') + '<b>' + Math.round(o.v) + '</b></div>';
      }).join('');
      st.put('F', '<span class="gd-n">' + fmt(n[0]) + '</span><small>客戶公司　聯絡人 ' + fmt(n[1]) + '</small>',
        '<div class="gd-donut-row"><svg viewBox="0 0 42 42" preserveAspectRatio="xMidYMid meet" role="img" aria-label="產業分佈">' + ci + '</svg><div class="gd-mini-leg">' + top3 + '</div></div>', '');
    }
  }

  function salesRefresh(animate) {
    if (!GD.built) return;
    var calc = { A: calcA(), B: calcB(), C: calcC(), D: calcD(), F: calcF() };
    var hidden = { B: !visibleEl($('quarterCardsWrap')), C: !visibleEl($('monthBudgetWrap')) };
    GD.stage.setHidden(hidden);
    ['A', 'B', 'C', 'D', 'F'].forEach(function (id) { if (calc[id]) GD.tw.set(id, calc[id].nums, calc[id].meta, animate); });
    GD.tw.set('E', calc.D.nums.slice(0, 4), { p: GD.period }, animate);
    SALES_MODS.forEach(function (m) { GD.stage.tag(m.id, tagFor(m.id, GD.period)); });
  }
  function scheduleRefresh() {
    if (GD.timer) clearTimeout(GD.timer);
    GD.timer = setTimeout(function () { GD.timer = 0; salesRefresh(true); GD.stage && GD.stage.layout(); }, 90);
  }

  function buildSales() {
    var view = $('dashboardView'), target = $('targetAchievementCard');
    if (GD.built || !view || !target) return;
    var root = el('div', 'gd-root'); root.id = 'gdRoot';
    root.innerHTML = '<svg width="0" height="0" style="position:absolute" aria-hidden="true" focusable="false"><defs><filter id="gdGoo" x="-20%" y="-60%" width="140%" height="220%"><feGaussianBlur in="SourceGraphic" stdDeviation="7" result="b"/><feColorMatrix in="b" mode="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 20 -9"/></filter></defs></svg>' +
      '<div class="gd-bar"><div class="gd-bar-note">點卡片放大看詳情；右側可切換資料期間</div><div class="gd-bar-r"><span class="gd-bar-lbl">資料期間</span><span id="gdTfSlot"></span></div></div>' +
      '<div class="gd-alerts" id="gdAlerts"></div>';
    view.insertBefore(root, target);
    view.classList.add('gd-on');
    GD.root = root;
    GD.cap = makeCapsule(GD.period, function (k) { GD.period = k; salesRefresh(true); });
    root.querySelector('#gdTfSlot').appendChild(GD.cap.el);
    // 提醒（殭屍商機／生日／Call-in）：保留原區塊與行為，搬到卡片區上方
    var al = root.querySelector('#gdAlerts');
    ['zombieAlertCard', 'birthdayReminderCard', 'dashCallinCard'].forEach(function (id) { mv(GD.moves, $(id), al); });
    // 舞台
    GD.stage = makeStage({
      id: 'gdStage', mods: SALES_MODS,
      cols: function (W) { return W >= 720 ? 3 : 2; },
      rowH: function (W) { return W >= 720 ? 212 : 216; },
      thumbCols: function (W) { return W >= 720 ? 5 : 3; },
      minFocusH: function (W) { return W >= 720 ? 620 : 700; }
    });
    root.appendChild(GD.stage.el);
    GD.tw = makeTweener(paintSales);
    // 詳情＝舊區塊（reparent）
    var d = GD.stage.detail;
    mv(GD.moves, $('targetAchievementCard'), d('A')); mv(GD.moves, $('buBreakdownWrap'), d('A'));
    mv(GD.moves, $('quarterCardsWrap'), d('B'));
    mv(GD.moves, $('monthBudgetWrap'), d('C'));
    mv(GD.moves, document.querySelector('#dashboardView .dash-pipeline-card'), d('D'));
    var statsD = el('div', 'stat-cards gd-stats'); d('D').appendChild(statsD);
    ['statCardA', 'statCardB', 'statCardC'].forEach(function (id) { mv(GD.moves, $(id), statsD); });
    mv(GD.moves, $('opportunityFunnel') && $('opportunityFunnel').closest('.chart-card'), d('E'));
    var statsF = el('div', 'stat-cards gd-stats'); d('F').appendChild(statsF);
    ['statCompanies', 'statContacts'].forEach(function (id) { var c = $(id) && $(id).closest('.stat-card'); mv(GD.moves, c, statsF); });
    mv(GD.moves, $('industryChart3D') && $('industryChart3D').closest('.chart-card'), d('F'));
    // 舊區塊內容變動 → 重算摘要
    GD.obs = new MutationObserver(scheduleRefresh);
    ['targetAchievementCard', 'quarterCardsWrap', 'monthBudgetWrap', 'buBreakdownWrap', 'statCompanies', 'statContacts', 'statCommit', 'statUpside', 'statPipeline',
      'dashPipelineCols', 'industryLegend', 'opportunityFunnel', 'zombieAlertCard', 'birthdayReminderCard', 'dashCallinCard'].forEach(function (id) {
      var n = $(id); if (n) GD.obs.observe(n, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['style', 'class'] });
    });
    // 首頁顯示／隱藏（showDashboard 設 display）→ 重排
    GD.viewObs = new MutationObserver(function () { if (view.style.display !== 'none') { GD.stage.layout(); scheduleRefresh(); } });
    GD.viewObs.observe(view, { attributes: true, attributeFilter: ['style'] });
    GD.built = true;
    GD.stage.layout();
    salesRefresh(true);
    GD.stage.layout();
    if (window.ResizeObserver) {
      GD.ro = new ResizeObserver(function () { if (GD.stage && GD.stage.el.clientWidth !== GD.stage.W) GD.stage.layout(); });
      GD.ro.observe(GD.stage.el);
    }
  }
  function destroySales() {
    if (!GD.built) return;
    if (GD.obs) GD.obs.disconnect();
    if (GD.viewObs) GD.viewObs.disconnect();
    if (GD.ro) GD.ro.disconnect();
    if (GD.tw) GD.tw.drop();
    if (GD.timer) clearTimeout(GD.timer);
    var i = STAGES.indexOf(GD.stage); if (i >= 0) STAGES.splice(i, 1);
    undoMoves(GD.moves);
    if (GD.root && GD.root.parentNode) GD.root.parentNode.removeChild(GD.root);
    var v = $('dashboardView'); if (v) v.classList.remove('gd-on');
    GD.built = false; GD.stage = null; GD.root = null; GD.ro = null; GD.timer = 0;
  }

  /* ═════════════════ 主管首頁 ═════════════════ */
  var GM = { built: false, stage: null, tw: null, root: null, data: null, moves: [] };
  var MGR_MODS = [
    { id: 'G1', title: '業績達成度', color: 'var(--g-accent)' },
    { id: 'G2', title: '本月可望成交', color: '#34d399' },
    { id: 'G3', title: '商機停滯時間分析（Aging）', color: '#f59e0b' },
    { id: 'G4', title: '客戶價值 TOP 10', color: '#a78bfa' }
  ];
  var BUCKET_COLORS = ['#34a853', '#7cb342', '#fbbc04', '#fb8c00', '#ea4335'];     // 與舊版 Aging 圖相同（語意：越紅越停滯）
  function barColor(p) { return p == null ? '#bbb' : p >= 90 ? '#1e8e3e' : p >= 60 ? '#e37400' : '#c5221f'; }   // 與舊版達成度燈號相同

  function mgrCalc(d) {
    var out = {}, a = d.achievement || {}, target = a.target || 0;
    var won = a.won || { achieved: a.achieved || 0, pct: a.pct }, rec = a.recognized || { achieved: 0, pct: target > 0 ? 0 : null };
    out.G1 = { nums: [won.pct == null ? 0 : won.pct, rec.pct == null ? 0 : rec.pct, won.achieved || 0, rec.achieved || 0, target], meta: { has: target > 0, wp: won.pct, rp: rec.pct } };
    var c = d.thisMonthCommit || {}, g = function (x) { return c[x] || { count: 0, amount: 0 }; };
    out.G2 = { nums: [g('confirmed').count, g('confirmed').amount, g('atRisk').count, g('atRisk').amount, g('uncertain').count, g('uncertain').amount], meta: {} };
    var ag = d.aging || { buckets: [], stages: [], data: {} }, flat = [];
    ag.stages.forEach(function (s) { ag.buckets.forEach(function (b) { flat.push((ag.data[s] || {})[b] || 0); }); });
    out.G3 = { nums: [ag.stalledCount || 0].concat(flat), meta: { stages: ag.stages, buckets: ag.buckets } };
    var list = (d.topCustomers || []).slice(0, 10), tn = [];
    list.forEach(function (x) { tn.push(x.won || 0, x.active || 0); });
    out.G4 = { nums: tn, meta: { names: list.map(function (x) { return x.company || ''; }), cnt: list.map(function (x) { return x.count || 0; }) } };
    return out;
  }
  function paintMgr(id, s) {
    var st = GM.stage, n = s.nums, M = s.meta || {};
    if (id === 'G1') {
      if (!M.has) { st.put('G1', '<span class="gd-n gd-mute">尚未設定目標</span>', '<div class="gd-note">本年成交 ' + fmt(n[2]) + ' K（無目標可比對）</div>', ''); return; }
      var row = function (lab, p, ach) {
        return '<div class="gd-hbar"><div class="gd-hbar-hd"><span>' + lab + '</span><b>' + Math.round(p) + '%</b></div><div class="gd-hbar-tr"><i style="width:' + Math.max(0, Math.min(100, p)) + '%;background:' + barColor(p) + '"></i></div><div class="gd-hbar-ft">' + fmt(ach) + ' K / 目標 ' + fmt(n[4]) + ' K</div></div>';
      };
      st.put('G1', '<span class="gd-n" style="color:' + barColor(M.wp) + '">' + Math.round(n[0]) + '%</span><small>成交達成率</small>',
        row('成交', n[0], n[2]) + (M.rp == null ? '' : row('認列', n[1], n[3])), '');
    } else if (id === 'G2') {
      var rows = [['確定可成交', '#1e8e3e', 0], ['風險案件', '#c5221f', 2], ['變數較大', '#e37400', 4]], mxa = Math.max(n[1], n[3], n[5], 1);
      st.put('G2', '<span class="gd-n">' + fmt(n[1]) + ' K</span><small>確定可成交 ' + Math.round(n[0]) + ' 案</small>',
        '<div class="gd-lines">' + rows.map(function (r) {
          return '<div class="gd-line"><span class="gd-lg-dot" style="background:' + r[1] + '"></span><span class="gd-line-l">' + r[0] + '</span><span class="gd-line-bar"><i style="width:' + (n[r[2] + 1] / mxa * 100) + '%;background:' + r[1] + '"></i></span><b>' + Math.round(n[r[2]]) + ' 案・' + fmt(n[r[2] + 1]) + 'K</b></div>';
        }).join('') + '</div>', '');
    } else if (id === 'G3') {
      var B = M.buckets || [], S2 = M.stages || [], vals = n.slice(1), mxs = 1, tots = S2.map(function (s2, i) { var t = sum(vals, i * B.length, (i + 1) * B.length); if (t > mxs) mxs = t; return t; });
      var NAMES = { A: 'A', B: 'B', C: 'C', D: 'D' };
      st.put('G3', '<span class="gd-n" style="color:' + (n[0] > 0 ? '#ef4444' : '#10b981') + '">' + Math.round(n[0]) + '</span><small>案停滯超過 60 天（C/B/A）</small>',
        '<div class="gd-lines">' + S2.map(function (s2, i) {
          var segs = B.map(function (b, j) { var v = vals[i * B.length + j] || 0; return '<i title="' + esc(b) + ' 天：' + Math.round(v) + ' 件" style="width:' + (v / mxs * 100) + '%;background:' + BUCKET_COLORS[j % 5] + '"></i>'; }).join('');
          return '<div class="gd-line"><span class="gd-line-l gd-st-' + esc(s2) + '">' + (NAMES[s2] || esc(s2)) + '</span><span class="gd-stack">' + segs + '</span><b>' + Math.round(tots[i]) + '</b></div>';
        }).join('') + '</div>', '');
    } else if (id === 'G4') {
      var k, total = sum(n, 0, n.length), top = [], mx = 1;
      for (k = 0; k < n.length; k += 2) { var t2 = n[k] + n[k + 1]; if (t2 > mx) mx = t2; }
      for (k = 0; k < Math.min(4, n.length / 2); k++) {
        var nm = M.names[k] || '', nm2 = nm.length > 10 ? nm.slice(0, 10) + '…' : nm;
        top.push('<div class="gd-line"><span class="gd-line-l gd-name" title="' + esc(nm) + '">' + esc(nm2) + '</span><span class="gd-stack"><i style="width:' + n[k * 2] / mx * 100 + '%;background:#1e8e3e"></i><i style="width:' + n[k * 2 + 1] / mx * 100 + '%;background:#1a73e8"></i></span><b>' + fmt(n[k * 2] + n[k * 2 + 1]) + 'K</b></div>');
      }
      st.put('G4', '<span class="gd-n">' + fmt(total) + ' K</span><small>TOP 10 合計（成交＋在手）</small>',
        n.length ? '<div class="gd-lines">' + top.join('') + '</div>' : '<div class="gd-note">尚無資料</div>', '');
    }
  }

  function buildMgr() {
    var view = $('managerHomeView');
    if (GM.built || !view) return;
    var gauge = $('mgrDualBars') && $('mgrDualBars').closest('.mgr-card'), commit = $('mgrCommitContent') && $('mgrCommitContent').closest('.mgr-card'),
      aging = $('mgrAgingChart') && $('mgrAgingChart').closest('.mgr-card'), topc = $('mgrTopCustChart') && $('mgrTopCustChart').closest('.mgr-card');
    if (!gauge || !commit || !aging || !topc) return;
    var root = el('div', 'gd-root'); root.id = 'gmRoot';
    var anchor = $('mgrGapBanner');
    view.insertBefore(root, anchor ? anchor.nextSibling : null);
    view.classList.add('gd-on');
    GM.root = root;
    GM.stage = makeStage({
      id: 'gmStage', mods: MGR_MODS,
      cols: function (W) { return W >= 520 ? 2 : 1; },
      rowH: function (W) { return W >= 520 ? 232 : 214; },
      thumbCols: function (W) { return W >= 520 ? 3 : 2; },
      minFocusH: function () { return 640; },
      onFocus: function () {         // Chart.js 在隱藏容器建立時尺寸為 0；詳情顯示後補一次 resize
        [380, 800].forEach(function (ms) {
          setTimeout(function () {
            ['mgrAgingChart', 'mgrTopCustChart'].forEach(function (cid) {
              var cv = $(cid); if (!cv || !window.Chart || !Chart.getChart) return;
              var ch = Chart.getChart(cv); if (ch) { try { ch.resize(); } catch (e) { /* ignore */ } }
            });
          }, ms);
        });
      }
    });
    root.appendChild(GM.stage.el);
    GM.tw = makeTweener(paintMgr);
    var d = GM.stage.detail;
    mv(GM.moves, gauge, d('G1')); mv(GM.moves, commit, d('G2')); mv(GM.moves, aging, d('G3')); mv(GM.moves, topc, d('G4'));
    GM.built = true;
    MGR_MODS.forEach(function (m) { GM.stage.put(m.id, '<span class="gd-n gd-mute">載入中…</span>', '', ''); });
    GM.stage.layout();
    if (GM.data) mgrUpdate(GM.data);
    if (window.ResizeObserver) {
      GM.ro = new ResizeObserver(function () { if (GM.stage && GM.stage.el.clientWidth !== GM.stage.W) GM.stage.layout(); });
      GM.ro.observe(GM.stage.el);
    }
    var vo = new MutationObserver(function () { if (view.style.display !== 'none' && GM.stage) GM.stage.layout(); });
    vo.observe(view, { attributes: true, attributeFilter: ['style'] });
    GM.viewObs = vo;
  }
  function mgrUpdate(d) {
    GM.data = d;
    if (!glassOn() || !GM.built) return;
    var c = mgrCalc(d), y = (typeof mgrYear !== 'undefined') ? mgrYear : '';
    ['G1', 'G2', 'G3', 'G4'].forEach(function (id) { GM.tw.set(id, c[id].nums, c[id].meta, true); });
    GM.stage.tag('G1', y ? y + ' 年' : '');
    GM.stage.tag('G2', '本月');
    GM.stage.tag('G3', '目前現況');
    GM.stage.tag('G4', '歷史成交＋在手商機');
    GM.stage.layout();
  }
  function destroyMgr() {
    if (!GM.built) return;
    if (GM.ro) GM.ro.disconnect();
    if (GM.viewObs) GM.viewObs.disconnect();
    if (GM.tw) GM.tw.drop();
    var i = STAGES.indexOf(GM.stage); if (i >= 0) STAGES.splice(i, 1);
    undoMoves(GM.moves);
    if (GM.root && GM.root.parentNode) GM.root.parentNode.removeChild(GM.root);
    var v = $('managerHomeView'); if (v) v.classList.remove('gd-on');
    GM.built = false; GM.stage = null; GM.root = null; GM.ro = null;
  }

  /* ═════════════════ 啟停＋全域事件 ═════════════════ */
  /* Chart.js 預設文字／格線色：舊版圖表在深色底上是灰字（對比不足）。只在玻璃模式覆寫，關閉時還原；已建立的圖表在下次載入時換色。 */
  var CH_ORIG = null, lastDark = null;
  function chartTheme() {
    if (!window.Chart || !Chart.defaults) return;
    if (!glassOn()) { if (CH_ORIG) { Chart.defaults.color = CH_ORIG.c; Chart.defaults.borderColor = CH_ORIG.b; CH_ORIG = null; } return; }
    if (!CH_ORIG) CH_ORIG = { c: Chart.defaults.color, b: Chart.defaults.borderColor };
    var dk = document.body.classList.contains('dark');
    Chart.defaults.color = dk ? 'rgba(255,255,255,.82)' : '#334155';
    Chart.defaults.borderColor = dk ? 'rgba(255,255,255,.14)' : 'rgba(15,27,53,.14)';
  }
  function sync() {
    chartTheme();
    if (glassOn()) {
      if (!GD.built) buildSales();
      if (!GM.built) buildMgr();
    } else if (GD.built || GM.built) {
      destroyMgr(); destroySales();
    }
  }
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape' || modalOpen()) return;
    if (GD.cap && GD.cap.isOpen()) { GD.cap.close(); return; }
    for (var i = 0; i < STAGES.length; i++) {
      var st = STAGES[i];
      if (st.focus && st.el.offsetParent !== null) { st.setFocus(null); return; }
    }
  });
  document.addEventListener('click', function (e) {
    if (GD.cap && GD.cap.isOpen() && !GD.cap.el.contains(e.target)) GD.cap.close();
  });
  var rz = 0;
  window.addEventListener('resize', function () {
    if (rz) return;
    rz = requestAnimationFrame(function () { rz = 0; STAGES.forEach(function (st) { if (st.el.offsetParent !== null) st.layout(); }); });
  });
  if (document.body) {
    new MutationObserver(function () {
      var dk = document.body.classList.contains('dark');
      sync();
      // 深／淺色切換：主管首頁的 Chart.js 圖表要重建才會換色（資料重抓，無副作用）
      if (lastDark !== null && lastDark !== dk && GM.built && glassOn() && $('managerHomeView') && $('managerHomeView').style.display !== 'none' && typeof loadManagerHome === 'function') loadManagerHome();
      lastDark = dk;
    }).observe(document.body, { attributes: true, attributeFilter: ['class'] });
    lastDark = document.body.classList.contains('dark');
  }

  window.GlassDash = {
    refresh: function () { salesRefresh(false); if (GD.stage) GD.stage.layout(); },
    setPeriod: function (p) { if (PK.indexOf(p) < 0) return; GD.period = p; salesRefresh(true); },
    state: function () { return { built: GD.built, period: GD.period, focus: GD.stage && GD.stage.focus }; },
    focus: function (id) { if (GD.stage) GD.stage.setFocus(id); }
  };
  window.GlassMgr = {
    update: mgrUpdate,
    focus: function (id) { if (GM.stage) GM.stage.setFocus(id); },
    state: function () { return { built: GM.built, focus: GM.stage && GM.stage.focus }; }
  };
  sync();
})();
