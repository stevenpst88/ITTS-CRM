/* ITTS CRM — 玻璃色彩主題選擇器（index.html / admin.html 共用）
 * 主題本體在 glass.css（body.glass[data-theme="..."]）；套用與記憶：window.setGlassTheme(key)（定義於頁面 <head>，這裡包一層同步 UI）。
 * 僅在玻璃模式（body.glass）顯示；?glass=0 時整個按鈕隱藏（見 glass.css .gt-wrap）。 */
/* ── 自動縮放（僅玻璃模式）：視窗寬於 1440 時用 CSS zoom 等比放大整個介面，文字不再過小 ──
 * 係數 = clamp(min(寬/1440, 高/700), 1, 1.5)，寬 ≤ 1024（手機／平板）恆為 1，不縮小。zoom 套在 body.glass（glass.css：zoom:var(--gz)）。
 * 使用者可在色彩主題面板用 A-／自動／A+ 微調（localStorage 'glassScale'，存步數，預設自動）；?glass=0 則整段不啟動。
 * zoom 會讓 vh/vw、getBoundingClientRect、clientX/Y、devicePixelRatio 與版面座標不一致，以下逐項補償（係數 = 1 時全部等同原生）：
 *  1) 樣式表／行內樣式中的 vh/vw 改寫成 calc(N vh / var(--gz))；2) 量測／指標座標除以係數；3) devicePixelRatio 乘係數（Chart.js 才不會糊）。
 *  帶 data-gz-off 的元素（報價 PDF 的離屏畫布）內部不 zoom，量測也不補償，輸出與 zoom=1 位元相同。 */
(function () {
  var root = document.documentElement, Z = 1, installed = false, vpSheet = null, vpTimer = 0, vpMO = null, inMO = null;
  var nat = {
    iw: Object.getOwnPropertyDescriptor(window, 'innerWidth'), ih: Object.getOwnPropertyDescriptor(window, 'innerHeight'),
    dpr: Object.getOwnPropertyDescriptor(window, 'devicePixelRatio'), gbr: Element.prototype.getBoundingClientRect,
    gcr: Element.prototype.getClientRects
  };
  var VP = /(-?\d*\.?\d+)(vh|vw|dvh|dvw|svh|svw|lvh|lvw)\b/g, VPT = /\d(vh|vw|dvh|dvw|svh|svw|lvh|lvw)\b/;
  function natW() { return nat.iw ? nat.iw.get.call(window) : window.innerWidth; }
  function natH() { return nat.ih ? nat.ih.get.call(window) : window.innerHeight; }
  function getStep() { try { var s = parseInt(localStorage.getItem('glassScale'), 10); return isFinite(s) ? Math.max(-3, Math.min(4, s)) : 0; } catch (e) { return 0; } }
  function compute() {
    if (!document.body || !document.body.classList.contains('glass')) return 1;
    var w = natW(), h = natH();
    if (w <= 1024) return 1;
    var auto = Math.min(w / 1440, h / 700);
    return Math.round(Math.max(1, Math.min(1.5, auto * (1 + 0.1 * getStep()) )) * 100) / 100;
  }
  function off(el) { return el && el.closest && el.closest('[data-gz-off]'); }
  function rect(r) { return new DOMRect(r.x / Z, r.y / Z, r.width / Z, r.height / Z); }
  function install() {
    if (installed) return; installed = true;
    function def(obj, name, get) { try { Object.defineProperty(obj, name, { configurable: true, get: get }); } catch (e) {} }
    def(window, 'innerWidth', function () { return Z === 1 ? natW() : natW() / Z; });
    def(window, 'innerHeight', function () { return Z === 1 ? natH() : natH() / Z; });
    def(window, 'devicePixelRatio', function () { var d = nat.dpr ? nat.dpr.get.call(window) : 1; return Z === 1 ? d : d * Z; });
    Element.prototype.getBoundingClientRect = function () {
      var r = nat.gbr.call(this); return (Z === 1 || off(this)) ? r : rect(r);
    };
    Element.prototype.getClientRects = function () {
      var l = nat.gcr.call(this); if (Z === 1 || off(this)) return l;
      return [].map.call(l, rect);
    };
    ['clientX', 'clientY', 'x', 'y', 'pageX', 'pageY', 'offsetX', 'offsetY'].forEach(function (k) {
      var d = Object.getOwnPropertyDescriptor(MouseEvent.prototype, k); if (!d || !d.get) return;
      def(MouseEvent.prototype, k, function () { var v = d.get.call(this); return Z === 1 ? v : v / Z; });
    });
  }
  /* vh/vw 補償：樣式表 */
  function fix(v) { return v.indexOf('--gz') >= 0 ? v : v.replace(VP, 'calc($1$2 / var(--gz,1))'); }
  function ruleText(rule) {
    if (rule.type === 1 && rule.style) {   /* CSSStyleRule */
      var out = '';
      for (var i = 0; i < rule.style.length; i++) {
        var p = rule.style[i], v = rule.style.getPropertyValue(p);
        if (VPT.test(v)) out += p + ':' + fix(v) + (rule.style.getPropertyPriority(p) ? ' !important' : '') + ';';
      }
      return out ? rule.selectorText + '{' + out + '}' : '';
    }
    if (rule.type === 4 || rule.type === 12) {   /* @media / @supports */
      var inner = ''; try { for (var j = 0; j < rule.cssRules.length; j++) inner += ruleText(rule.cssRules[j]); } catch (e) {}
      return inner ? (rule.type === 4 ? '@media ' : '@supports ') + (rule.conditionText || rule.media.mediaText) + '{' + inner + '}' : '';
    }
    return '';
  }
  function buildVp() {
    var css = '';
    [].forEach.call(document.styleSheets, function (ss) {
      if (ss === vpSheet) return;
      var rs; try { rs = ss.cssRules; } catch (e) { return; }
      for (var i = 0; i < rs.length; i++) css += ruleText(rs[i]);
    });
    if (!vpSheet) {
      var el = document.createElement('style'); el.id = 'gz-vp'; document.head.appendChild(el); vpSheet = el.sheet; vpSheet.__el = el;
    }
    vpSheet.__el.textContent = css;
  }
  function scheduleVp() { clearTimeout(vpTimer); vpTimer = setTimeout(function () { if (Z !== 1) buildVp(); }, 60); }
  /* vh/vw 補償：行內樣式 */
  function fixInline(el) {
    var st = el.style; if (!st || !VPT.test(el.getAttribute('style') || '')) return;
    for (var i = 0; i < st.length; i++) {
      var p = st[i], v = st.getPropertyValue(p);
      if (v.indexOf('--gz') < 0 && VPT.test(v)) st.setProperty(p, fix(v), st.getPropertyPriority(p));
    }
  }
  function sweep(node) {
    if (node.nodeType !== 1) return;
    fixInline(node);
    [].forEach.call(node.querySelectorAll('[style*="vh"],[style*="vw"]'), fixInline);
  }
  function startObservers() {
    if (vpMO) return;
    vpMO = new MutationObserver(function (ms) {
      for (var i = 0; i < ms.length; i++) for (var j = 0; j < ms[i].addedNodes.length; j++) {
        var n = ms[i].addedNodes[j]; if (n.nodeName === 'STYLE' || n.nodeName === 'LINK') scheduleVp();
      }
    });
    vpMO.observe(document.head, { childList: true });
    inMO = new MutationObserver(function (ms) {
      for (var i = 0; i < ms.length; i++) {
        var m = ms[i]; if (m.type === 'attributes') fixInline(m.target); else for (var j = 0; j < m.addedNodes.length; j++) sweep(m.addedNodes[j]);
      }
    });
    inMO.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['style'] });
    sweep(document.body);
  }
  var raf = 0, lastSet = 0, held = 0;
  function apply() {
    if (held) return;
    var z = compute();
    if (z === Z) return;
    if (z !== 1) { install(); if (!vpSheet) buildVp(); startObservers(); }
    Z = z;
    if (z === 1) root.style.removeProperty('--gz'); else root.style.setProperty('--gz', String(z));
    if (z !== 1 || lastSet) { lastSet = 1; try { window.dispatchEvent(new Event('resize')); } catch (e) {} }
    label();
  }
  function label() { var lab = document.getElementById('glassScaleVal'); if (lab) lab.textContent = getStep() ? Math.round(Z * 100) + '%' : '自動'; }
  var resizing = false;
  window.addEventListener('resize', function () {
    if (resizing) return;
    cancelAnimationFrame(raf); raf = requestAnimationFrame(function () { resizing = true; try { apply(); } finally { resizing = false; } });
  });
  new MutationObserver(function () { apply(); }).observe(document.body, { attributes: true, attributeFilter: ['class'] });
  window.__glassScale = {
    get: function () { return Z; },
    step: function (d) { try { if (d === 0) localStorage.removeItem('glassScale'); else localStorage.setItem('glassScale', String(Math.max(-3, Math.min(4, getStep() + d)))); } catch (e) {} Z = -1; apply(); },
    label: label,
    /* 暫停縮放（報價 PDF 擷取用）：回傳 resume()；期間 body 隱藏避免畫面跳動 */
    hold: function () {
      held++; Z = 1; root.style.setProperty('--gz', '1'); document.body.style.visibility = 'hidden';
      var done = false;
      return function () { if (done) return; done = true; held = Math.max(0, held - 1); if (!held) { document.body.style.visibility = ''; Z = -1; apply(); } };
    },
  };
  apply();
})();

(function () {
  var THEMES = [
    { k: 'ocean',    n: '深海藍', c1: '#2563eb', c2: '#d946ef' },
    { k: 'noir',     n: '黑金',   c1: '#c9a227', c2: '#e8c15b' },
    { k: 'aurora',   n: '極光',   c1: '#0fb89a', c2: '#7c3aed' },
    { k: 'dusk',     n: '暮光',   c1: '#d946ef', c2: '#f97316' },
    { k: 'ember',    n: '琥珀紅', c1: '#ef4444', c2: '#f59e0b' },
    { k: 'graphite', n: '石墨',   c1: '#64748b', c2: '#cbd5e1' }
  ];
  var wrap = document.getElementById('glassThemeWrap');
  if (!wrap) return;

  function current() {
    var t = document.body.dataset.theme;
    return THEMES.some(function (x) { return x.k === t; }) ? t : 'ocean';
  }

  wrap.innerHTML =
    '<button type="button" class="gt-btn" id="glassThemeBtn" title="色彩主題" aria-label="色彩主題" aria-haspopup="true" aria-expanded="false">\uD83C\uDFA8</button>' +
    '<div class="gt-pop" id="glassThemePop" role="menu" hidden>' +
    '<div class="gt-pop-title">色彩主題</div>' +
    '<div class="gt-grid">' + THEMES.map(function (t) {
      return '<button type="button" class="gt-item" role="menuitemradio" data-theme-key="' + t.k + '" aria-checked="false">' +
        '<span class="gt-sw" style="background:linear-gradient(135deg,' + t.c1 + ',' + t.c2 + ')"></span><span class="gt-name">' + t.n + '</span></button>';
    }).join('') + '</div>' +
    '<div class="gt-pop-title gt-size-title">畫面大小</div>' +
    '<div class="gt-size"><button type="button" class="gt-sz" data-gz="-1" aria-label="縮小">A-</button><button type="button" class="gt-sz gt-sz-auto" data-gz="0" aria-label="自動" title="依視窗大小自動"><span id="glassScaleVal">自動</span></button><button type="button" class="gt-sz" data-gz="1" aria-label="放大">A+</button></div></div>';

  var btn = document.getElementById('glassThemeBtn');
  var pop = document.getElementById('glassThemePop');

  function mark() {
    var cur = current();
    [].forEach.call(pop.querySelectorAll('.gt-item'), function (b) {
      var on = b.getAttribute('data-theme-key') === cur;
      b.classList.toggle('on', on);
      b.setAttribute('aria-checked', on ? 'true' : 'false');
    });
  }
  function open(v) {
    pop.hidden = !v;
    btn.setAttribute('aria-expanded', v ? 'true' : 'false');
    if (v) mark();
  }

  var base = window.setGlassTheme;
  window.setGlassTheme = function (k) {
    if (typeof base === 'function') base(k);
    else {
      try { localStorage.setItem('glassTheme', k); } catch (e) {}
      document.body.dataset.theme = k;
    }
    mark();
  };

  btn.addEventListener('click', function (e) { e.stopPropagation(); open(pop.hidden); });
  pop.addEventListener('click', function (e) {
    var z = e.target.closest('.gt-sz');
    if (z && window.__glassScale) { window.__glassScale.step(parseInt(z.getAttribute('data-gz'), 10)); return; }
    var b = e.target.closest('.gt-item');
    if (!b) return;
    window.setGlassTheme(b.getAttribute('data-theme-key'));
  });
  document.addEventListener('click', function (e) { if (!pop.hidden && !wrap.contains(e.target)) open(false); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !pop.hidden) { open(false); btn.focus(); } });
  mark();
  if (window.__glassScale) window.__glassScale.label();
})();
