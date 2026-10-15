/* 登入頁（電影感開場）啟動前置：在 <head> 同步執行，先把主題／明暗套到 <html>，避免淺色使用者先閃一下深色。
   主題沿用玻璃 UI 的 localStorage 'glassTheme'（ocean/noir/aurora/dusk/ember/graphite）與 'darkMode'（'0'＝淺色，其餘＝深色；登入頁預設深色、不跟 OS）。
   ?theme= / ?mode= 只影響當次載入，不寫入儲存。所有儲存讀寫都包 try/catch，失敗不影響登入。 */
(function () {
  var THEMES = {
    ocean:{n:'深海藍',bg:'#050914',b1:'#2563eb',b2:'#d946ef',b3:'#0d9488',ac:'#7cb8ff',acl:'#1558d6',lmix:0,tint:'rgba(120,160,255,.13)',edge:'rgba(255,255,255,.20)',edge2:'rgba(255,255,255,.46)'},
    noir:{n:'黑金',bg:'#09090a',b1:'#c9a227',b2:'#7a5a12',b3:'#3d3216',ac:'#e8c15b',acl:'#6e5508',lmix:.55,tint:'rgba(232,193,91,.10)',edge:'rgba(232,193,91,.28)',edge2:'rgba(232,193,91,.65)'},
    aurora:{n:'極光',bg:'#05181a',b1:'#0fb89a',b2:'#2563eb',b3:'#7c3aed',ac:'#5eead4',acl:'#0d635c',lmix:.55,tint:'rgba(45,212,191,.12)',edge:'rgba(94,234,212,.22)',edge2:'rgba(94,234,212,.55)'},
    dusk:{n:'暮光',bg:'#160a24',b1:'#d946ef',b2:'#f97316',b3:'#6366f1',ac:'#f0abfc',acl:'#971aa3',lmix:.55,tint:'rgba(217,70,239,.12)',edge:'rgba(240,171,252,.22)',edge2:'rgba(240,171,252,.55)'},
    ember:{n:'琥珀紅',bg:'#1a0a07',b1:'#ef4444',b2:'#f59e0b',b3:'#7f1d1d',ac:'#fbbf24',acl:'#904207',lmix:.55,tint:'rgba(245,158,11,.11)',edge:'rgba(251,191,36,.24)',edge2:'rgba(251,191,36,.58)'},
    graphite:{n:'石墨',bg:'#0d1015',b1:'#64748b',b2:'#475569',b3:'#334155',ac:'#cbd5e1',acl:'#475569',lmix:.55,tint:'rgba(203,213,225,.08)',edge:'rgba(203,213,225,.18)',edge2:'rgba(203,213,225,.5)'}
  };
  var own = function (o, k) { return typeof k === 'string' && Object.prototype.hasOwnProperty.call(o, k); };
  var theme = 'ocean', mode = 'dark';
  try {
    var q = new URLSearchParams(location.search), st = null, sm = null;
    try { st = localStorage.getItem('glassTheme'); sm = localStorage.getItem('darkMode'); } catch (e) {}
    var qt = q.get('theme'), qm = q.get('mode');
    theme = own(THEMES, qt) ? qt : (own(THEMES, st) ? st : 'ocean');
    mode = (qm === 'light' || qm === 'dark') ? qm : (sm === '0' ? 'light' : 'dark');
  } catch (e) {}
  /* 把主題色變數寫到 root（CSSOM，不產生 inline style 屬性字串，CSP 友善） */
  function paint(root, th, md) {
    var T = THEMES[th], dark = md === 'dark', s = root.style;
    root.dataset.theme = th; root.dataset.mode = md;
    s.setProperty('--bg', T.bg); s.setProperty('--b1', T.b1); s.setProperty('--b2', T.b2); s.setProperty('--b3', T.b3);
    s.setProperty('--ac', dark ? T.ac : T.acl); s.setProperty('--acl', T.acl);
    s.setProperty('--btn', dark ? T.ac : T.acl); s.setProperty('--btn-tx', dark ? T.bg : '#ffffff');
    s.setProperty('--edge', dark ? T.edge : 'rgba(255,255,255,.75)'); s.setProperty('--edge2', dark ? T.edge2 : 'rgba(255,255,255,.95)');
    if (dark) s.setProperty('--card-hi', 'color-mix(in srgb,' + T.b1 + ' 20%,rgba(10,18,44,.55))'); else s.removeProperty('--card-hi');
    var meta = document.querySelector('meta[name=theme-color]'); if (meta) meta.content = dark ? T.bg : '#dfe9ff';
  }
  window.ITTSLoginBoot = { THEMES: THEMES, theme: theme, mode: mode, paint: paint };
  paint(document.documentElement, theme, mode);
})();
