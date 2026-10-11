/* ITTS CRM — 玻璃色彩主題選擇器（index.html / admin.html 共用）
 * 主題本體在 glass.css（body.glass[data-theme="..."]）；套用與記憶：window.setGlassTheme(key)（定義於頁面 <head>，這裡包一層同步 UI）。
 * 僅在玻璃模式（body.glass）顯示；?glass=0 時整個按鈕隱藏（見 glass.css .gt-wrap）。 */
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
    }).join('') + '</div></div>';

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
    var b = e.target.closest('.gt-item');
    if (!b) return;
    window.setGlassTheme(b.getAttribute('data-theme-key'));
  });
  document.addEventListener('click', function (e) { if (!pop.hidden && !wrap.contains(e.target)) open(false); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !pop.hidden) { open(false); btn.focus(); } });
  mark();
})();
