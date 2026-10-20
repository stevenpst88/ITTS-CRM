/* 簽核流程總覽（唯讀自由畫布）— flow-overview.js
 * 把後端產生的「節點／連線／自動排版座標」（lib/quoteFlowDesigner.js buildOverview）畫成 SVG，提供縮放、平移、縮放至適合、鍵盤操作與模擬高亮。
 * 唯讀：沒有任何編輯能力。所有文字一律以 textContent／setAttribute 寫入（不使用 innerHTML），所以流程名稱、標籤含 HTML 也只會顯示成純文字。
 *   const ov = FlowOverview.mount(hostEl, { onFocusNode(node) })
 *   ov.setGraph(graph)   graph = { nodes[], edges[], bounds, hitRoute, outline[] }（後端回傳）
 *   ov.fit() / ov.zoomBy(1.2) / ov.reset() / ov.state()   state() → { k, tx, ty, nodes, edges }
 * 純函式（edgePath、fitText、textUnits、viewFor）掛在 FlowOverview 上，供 scripts/check-quote-flow-designer-ui.js 以 vm 測試。 */
(function (root) {
  'use strict';
  var NS = 'http://www.w3.org/2000/svg';
  var COLOR_BY_TIER = { mgr2: 'teal', mgr1: 'blue', gm: 'pink', chairman: 'amber', board: 'violet' };

  function el(name, attrs, text) {
    var e = document.createElementNS(NS, name);
    if (attrs) for (var k in attrs) if (Object.prototype.hasOwnProperty.call(attrs, k)) e.setAttribute(k, String(attrs[k]));
    if (text != null) e.textContent = text;
    return e;
  }
  /** 字寬估計：中日韓全形 1em、其他 0.56em */
  function charUnits(ch) { return ch.charCodeAt(0) > 255 ? 1 : 0.56; }
  function textUnits(s) { var w = 0; for (var i = 0; i < s.length; i++) w += charUnits(s.charAt(i)); return w; }
  /** 依最大寬度（em）截斷並加「…」 */
  function fitText(s, maxEm) {
    s = String(s == null ? '' : s);
    var w = 0, out = '';
    for (var i = 0; i < s.length; i++) {
      var c = charUnits(s.charAt(i));
      if (w + c > maxEm) return out + '…';
      w += c; out += s.charAt(i);
    }
    return out;
  }
  /** 以最大寬度（em）折行，最多 maxLines 行；超過的部分以「…」結尾 */
  function wrapText(s, maxEm, maxLines) {
    s = String(s == null ? '' : s);
    var lines = [], cur = '', w = 0;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i), c = charUnits(ch);
      if (w + c > maxEm) {
        if (lines.length + 1 >= maxLines) { return lines.concat([fitText(cur + s.slice(i), maxEm)]); }
        lines.push(cur); cur = ''; w = 0;
      }
      cur += ch; w += c;
    }
    lines.push(cur);
    return lines.slice(0, maxLines);
  }
  function byId(nodes) { var m = {}; nodes.forEach(function (n) { m[n.id] = n; }); return m; }

  /** 連線折線的點列（直角）：同列→水平；同欄→垂直；到結束節點→先右再垂直再進入；其他→先垂直再水平 */
  function edgePath(e, m, busX) {
    var a = m[e.from], b = m[e.to];
    if (!a || !b) return [];
    if (b.kind === 'end') {
      var x1 = a.x + a.w / 2, bx = b.x - b.w / 2;
      var xm = busX != null ? busX : bx - 24;
      if (Math.abs(a.y - b.y) < 1) return [[x1, a.y], [bx, b.y]];
      return [[x1, a.y], [xm, a.y], [xm, b.y], [bx, b.y]];
    }
    if (Math.abs(a.y - b.y) < 1) {
      var dir = b.x >= a.x ? 1 : -1;
      return [[a.x + dir * a.w / 2, a.y], [b.x - dir * b.w / 2, b.y]];
    }
    if (Math.abs(a.x - b.x) < 1) {
      var dy = b.y >= a.y ? 1 : -1;
      return [[a.x, a.y + dy * a.h / 2], [b.x, b.y - dy * b.h / 2]];
    }
    var dyy = b.y >= a.y ? 1 : -1;
    return [[a.x, a.y + dyy * a.h / 2], [a.x, b.y], [b.x - b.w / 2, b.y]];
  }
  function pointsToD(pts) { return pts.map(function (p, i) { return (i ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1); }).join(' '); }
  /** 讓整張圖落在可視範圍內的視圖參數 */
  function viewFor(bounds, w, h, maxK) {
    var bw = Math.max(1, bounds.x1 - bounds.x0), bh = Math.max(1, bounds.y1 - bounds.y0);
    var pad = 16;
    var k = Math.min((w - 2 * pad) / bw, (h - 2 * pad) / bh, maxK || 1.2);
    k = Math.max(0.15, k);
    return { k: k, tx: (w - bw * k) / 2 - bounds.x0 * k, ty: (h - bh * k) / 2 - bounds.y0 * k };
  }

  function mount(host, opts) {
    opts = opts || {};
    var graph = null, V = { k: 1, tx: 0, ty: 0 }, nodeEls = {}, drag = null;
    var svg = el('svg', { class: 'fd-ovsvg', role: 'group', tabindex: '0', 'aria-label': '簽核流程總覽圖（唯讀）。方向鍵平移，加號減號縮放，0 縮放至適合。' });
    var defs = el('defs');
    [['fo-ar', 'fo-mk'], ['fo-ar-hit', 'fo-mk-hit']].forEach(function (p) {
      var mk = el('marker', { id: p[0], viewBox: '0 0 10 10', refX: 9, refY: 5, markerWidth: 9, markerHeight: 9, orient: 'auto', markerUnits: 'userSpaceOnUse' });
      mk.appendChild(el('path', { d: 'M0 0L10 5L0 10z', class: p[1] }));
      defs.appendChild(mk);
    });
    svg.appendChild(defs);
    var world = el('g', { class: 'fo-world' });
    var gE = el('g', { class: 'fo-edges' }), gN = el('g', { class: 'fo-nodes' });
    world.appendChild(gE); world.appendChild(gN); svg.appendChild(world);
    host.textContent = '';
    host.appendChild(svg);

    function apply() {
      world.setAttribute('transform', 'translate(' + V.tx.toFixed(2) + ' ' + V.ty.toFixed(2) + ') scale(' + V.k.toFixed(4) + ')');
      if (opts.onView) opts.onView({ k: V.k, tx: V.tx, ty: V.ty });
    }
    function size() { var r = svg.getBoundingClientRect(); return { w: r.width || 800, h: r.height || 480 }; }
    function fit() {
      if (!graph) return;
      var s = size();
      var v = viewFor(graph.bounds, s.w, s.h, 1.2);
      V.k = v.k; V.tx = v.tx; V.ty = v.ty; apply();
    }
    function zoomAt(px, py, f) {
      var nk = Math.min(2.5, Math.max(0.15, V.k * f)), r = nk / V.k;
      V.tx = px - (px - V.tx) * r; V.ty = py - (py - V.ty) * r; V.k = nk; apply();
    }
    function zoomBy(f) { var s = size(); zoomAt(s.w / 2, s.h / 2, f); }

    function nodeShape(n, g) {
      var hw = n.w / 2, hh = n.h / 2;
      if (n.kind === 'decision') {
        g.appendChild(el('polygon', { class: 'fo-body', points: '0,' + (-hh) + ' ' + hw + ',0 0,' + hh + ' ' + (-hw) + ',0' }));
        g.appendChild(el('text', { class: 'fo-t1 fo-tc', y: -16 }, fitText(n.label, (n.w - 84) / 13.5)));
        wrapText(n.sub || '', (n.w - 96) / 11.5, 2).forEach(function (ln, i) { g.appendChild(el('text', { class: 'fo-t2 fo-tc', y: 4 + i * 14 }, ln)); });
      } else if (n.kind === 'approval') {
        g.appendChild(el('rect', { class: 'fo-body', x: -hw, y: -hh, width: n.w, height: n.h, rx: 12 }));
        g.appendChild(el('rect', { class: 'fo-stripe', x: -hw + 1.5, y: -hh + 10, width: 4, height: n.h - 20, rx: 2 }));
        g.appendChild(el('text', { class: 'fo-t1', x: -hw + 16, y: -4 }, fitText((n.locked ? '🔒 ' : '') + n.label, (n.w - 28) / 13.5)));
        g.appendChild(el('text', { class: 'fo-t2', x: -hw + 16, y: 14 }, fitText(n.sub || '', (n.w - 28) / 11.5)));
        if (n.inserted) g.appendChild(el('text', { class: 'fo-ins', x: hw - 8, y: -hh + 15, 'text-anchor': 'end' }, '＋調整插入'));
        if (n.mods && n.mods.length) g.appendChild(el('text', { class: 'fo-gear', x: hw - 8, y: hh - 8, 'text-anchor': 'end' }, '⚙'));
      } else if (n.kind === 'mod') {
        g.appendChild(el('rect', { class: 'fo-body', x: -hw, y: -hh, width: n.w, height: n.h, rx: 10 }));
        g.appendChild(el('text', { class: 'fo-t1', x: -hw + 12, y: -hh + 22 }, fitText(n.label, (n.w - 24) / 13.5)));
        wrapText(n.sub || '', (n.w - 24) / 11.5, 2).forEach(function (ln, i) { g.appendChild(el('text', { class: 'fo-t2', x: -hw + 12, y: -hh + 40 + i * 15 }, ln)); });
      } else {
        g.appendChild(el('rect', { class: 'fo-body', x: -hw, y: -hh, width: n.w, height: n.h, rx: 20 }));
        g.appendChild(el('text', { class: 'fo-t1 fo-tc', y: 5 }, n.label));
      }
    }
    function setGraph(gr, keepView) {
      graph = gr;
      gE.textContent = ''; gN.textContent = ''; nodeEls = {};
      var m = byId(gr.nodes);
      var ends = gr.nodes.filter(function (n) { return n.kind === 'end'; })[0];
      var simOn = gr.nodes.some(function (n) { return n.hit && n.kind !== 'start' && n.kind !== 'end'; }) && gr.edges.some(function (e) { return e.hit; });
      var busBase = ends ? ends.x - ends.w / 2 - 24 : null;
      gr.edges.forEach(function (e) {
        var toEnd = m[e.to] && m[e.to].kind === 'end';
        var pts = edgePath(e, m, toEnd ? busBase : null);
        if (!pts.length) return;
        var g = el('g', { class: 'fo-edge ' + e.kind + (e.hit ? ' hit' : (simOn ? ' dim' : '')), 'data-edge': e.id });
        g.appendChild(el('path', { class: 'fo-line', d: pointsToD(pts), 'marker-end': 'url(#' + (e.hit ? 'fo-ar-hit' : 'fo-ar') + ')' }));
        if (e.label) {
          var p0 = pts[0], p1 = pts[1];
          var mx = (p0[0] + p1[0]) / 2, my = (p0[1] + p1[1]) / 2;
          var w = textUnits(e.label) * 12 + 14;
          var lab = el('g', { transform: 'translate(' + mx.toFixed(1) + ' ' + my.toFixed(1) + ')' });
          lab.appendChild(el('rect', { class: 'fo-lab', x: -w / 2, y: -10, width: w, height: 20, rx: 7 }));
          lab.appendChild(el('text', { class: 'fo-labt', y: 4.5 }, e.label));
          g.appendChild(lab);
        }
        g.appendChild(el('title', null, (m[e.from] ? m[e.from].label : '') + ' → ' + (m[e.to] ? m[e.to].label : '') + (e.label ? '：' + e.label : '')));
        gE.appendChild(g);
      });
      gr.nodes.forEach(function (n) {
        var cls = 'fo-node ' + n.kind + (n.locked ? ' locked' : '') + (n.hit ? ' hit' : (simOn && n.kind !== 'mod' ? ' dim' : '')) + (n.kind === 'approval' ? ' fd-c-' + (n.color || COLOR_BY_TIER[n.tier] || 'slate') : '');
        var g = el('g', { class: cls, transform: 'translate(' + n.x + ' ' + n.y + ')', tabindex: '0', role: 'img', 'data-node': n.id, 'aria-label': n.label + (n.sub ? '，' + n.sub : '') + (n.locked ? '，鎖定的代核關' : '') + (n.hit ? '，本單走這裡' : '') });
        nodeShape(n, g);
        g.appendChild(el('title', null, n.label + (n.sub ? '：' + n.sub : '')));
        g.addEventListener('focus', function () { if (opts.onFocusNode) opts.onFocusNode(n); });
        nodeEls[n.id] = g;
        gN.appendChild(g);
      });
      if (keepView) apply(); else fit();
    }

    // 滑鼠／觸控：滾輪縮放（以游標為中心）、拖曳平移、雙指縮放
    var pts = new Map();
    svg.addEventListener('wheel', function (ev) {
      ev.preventDefault();
      var r = svg.getBoundingClientRect();
      zoomAt(ev.clientX - r.left, ev.clientY - r.top, ev.deltaY < 0 ? 1.12 : 1 / 1.12);
    }, { passive: false });
    svg.addEventListener('pointerdown', function (ev) {
      if (ev.button !== undefined && ev.button !== 0) return;
      pts.set(ev.pointerId, { x: ev.clientX, y: ev.clientY });
      try { svg.setPointerCapture(ev.pointerId); } catch (_) { /* 略 */ }
      if (pts.size === 1) drag = { type: 'pan', sx: ev.clientX, sy: ev.clientY, tx: V.tx, ty: V.ty, moved: false };
      else if (pts.size === 2) { var a = Array.from(pts.values()); drag = { type: 'pinch', d0: Math.hypot(a[0].x - a[1].x, a[0].y - a[1].y) || 1, k0: V.k }; }
    });
    svg.addEventListener('pointermove', function (ev) {
      if (!pts.has(ev.pointerId)) return;
      pts.set(ev.pointerId, { x: ev.clientX, y: ev.clientY });
      if (!drag) return;
      if (drag.type === 'pan' && pts.size === 1) {
        var dx = ev.clientX - drag.sx, dy = ev.clientY - drag.sy;
        if (!drag.moved && Math.hypot(dx, dy) < 4) return;
        drag.moved = true; svg.classList.add('panning');
        V.tx = drag.tx + dx; V.ty = drag.ty + dy; apply();
      } else if (drag.type === 'pinch' && pts.size === 2) {
        var a = Array.from(pts.values()), d = Math.hypot(a[0].x - a[1].x, a[0].y - a[1].y) || 1;
        var r = svg.getBoundingClientRect();
        zoomAt((a[0].x + a[1].x) / 2 - r.left, (a[0].y + a[1].y) / 2 - r.top, (drag.k0 * d / drag.d0) / V.k);
      }
    });
    function endPtr(ev) {
      pts.delete(ev.pointerId);
      if (!pts.size) { drag = null; svg.classList.remove('panning'); }
      else if (pts.size === 1) { var p = Array.from(pts.values())[0]; drag = { type: 'pan', sx: p.x, sy: p.y, tx: V.tx, ty: V.ty, moved: true }; }
    }
    svg.addEventListener('pointerup', endPtr);
    svg.addEventListener('pointercancel', endPtr);
    svg.addEventListener('keydown', function (ev) {
      if (ev.target !== svg) return;
      var step = 60, k = ev.key;
      if (k === 'ArrowLeft') V.tx += step; else if (k === 'ArrowRight') V.tx -= step; else if (k === 'ArrowUp') V.ty += step; else if (k === 'ArrowDown') V.ty -= step;
      else if (k === '+' || k === '=') { zoomBy(1.2); ev.preventDefault(); return; }
      else if (k === '-' || k === '_') { zoomBy(1 / 1.2); ev.preventDefault(); return; }
      else if (k === '0' || k === 'Home') { fit(); ev.preventDefault(); return; }
      else return;
      ev.preventDefault(); apply();
    });
    // 節點取得焦點時自動捲到可視範圍內
    function ensureVisible(n) {
      var s = size();
      var sx = n.x * V.k + V.tx, sy = n.y * V.k + V.ty;
      var m = 60, dx = 0, dy = 0;
      if (sx < m) dx = m - sx; else if (sx > s.w - m) dx = s.w - m - sx;
      if (sy < m) dy = m - sy; else if (sy > s.h - m) dy = s.h - m - sy;
      if (dx || dy) { V.tx += dx; V.ty += dy; apply(); }
    }
    var origFocus = opts.onFocusNode;
    opts.onFocusNode = function (n) { ensureVisible(n); if (origFocus) origFocus(n); };

    return {
      setGraph: setGraph, fit: fit, zoomBy: zoomBy,
      reset: function () { V = { k: 1, tx: 0, ty: 0 }; apply(); },
      state: function () { return { k: V.k, tx: V.tx, ty: V.ty, nodes: graph ? graph.nodes.length : 0, edges: graph ? graph.edges.length : 0 }; },
      svg: svg,
      destroy: function () { host.textContent = ''; },
    };
  }

  root.FlowOverview = { mount: mount, edgePath: edgePath, wrapText: wrapText, fitText: fitText, textUnits: textUnits, viewFor: viewFor, pointsToD: pointsToD };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.FlowOverview;
})(typeof window !== 'undefined' ? window : globalThis);
