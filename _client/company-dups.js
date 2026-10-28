/* 企業主檔「疑似重複」清單（行銷／管理員）
 *
 * 後端：GET /api/companies/duplicates、POST /api/companies/duplicates/dismiss（lib/companyDuplicates.js）。
 * 只讀偵測＋「標記為非重複」，不合併、不刪除。成員一變動（有新主檔加入）該組會自動重新出現。
 * 安全：動態內容一律 createElement＋textContent、不用 inline style／on* 屬性；樣式沿用 .cdup-*，另有 .cdup-*。
 * 對外：window.ITTSCompanyDups = { mount }   由 app.js 的 loadCompanyMasterView 在列表載入後呼叫
 */
(function () {
  'use strict';
  if (window.ITTSCompanyDups) return;

  var REASONS = {
    sameTaxId: '同一個統編有多筆主檔（資料異常）',
    taxVsNoTax: '名稱相同，但一筆有統編、一筆沒有',
    nameVariant: '都沒有統編，公司名稱寫法不同',
    diffTaxId: '名稱相同但統編不同（可能是不同法人或分公司，請人工判斷）',
  };
  var host = null, panel = null, toggleBtn = null, opened = false, showDismissed = false, data = null;

  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = String(text); return e; }
  function btn(label, cls, onClick) { var b = el('button', cls, label); b.type = 'button'; if (onClick) b.addEventListener('click', onClick); return b; }
  function toast(msg) { try { if (typeof window.showToast === 'function') window.showToast(msg); } catch (e) { /* ignore */ } }
  function fmtCap(n) { n = Number(n); if (!n) return '—'; return n >= 100000000 ? (n / 100000000).toFixed(1) + ' 億' : Math.round(n / 10000) + ' 萬'; }

  function load() {
    panel.textContent = ''; panel.appendChild(el('div', 'cdup-loading', '載入中…'));
    fetch('/api/companies/duplicates', { credentials: 'same-origin', cache: 'no-store' }).then(function (r) {
      return r.json().then(function (d) { if (!r.ok) throw new Error(d.error || '載入失敗 (' + r.status + ')'); return d; });
    }).then(function (d) { data = d; render(); })
      .catch(function (e) { panel.textContent = ''; panel.appendChild(el('div', 'cdup-loading', '⚠ ' + e.message)); });
  }

  function dismiss(g, undo) {
    fetch('/api/companies/duplicates/dismiss', { method: 'POST', credentials: 'same-origin', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ids: g.members.map(function (m) { return m.id; }), undo: !!undo }) })
      .then(function (r) { return r.json().then(function (d) { if (!r.ok) throw new Error(d.error || '操作失敗'); }); })
      .then(function () { toast(undo ? '已還原為疑似重複' : '已標記為非重複'); load(); })
      .catch(function (e) { toast('⛔ ' + e.message); });
  }

  function groupCard(g) {
    var c = el('div', 'cdup-card' + (g.level === 'high' ? ' cdup-high' : ' cdup-low'));
    var hd = el('div', 'cdup-hd');
    hd.appendChild(el('span', 'cdup-badge', g.level === 'high' ? '高信心' : '低信心'));
    hd.appendChild(el('span', 'cdup-reason', g.reasons.map(function (r) { return REASONS[r] || r; }).join('；')));
    c.appendChild(hd);
    var wrap = el('div', 'cdup-prev-wrap'), t = el('table', 'cdup-prev-table'), tr0 = el('tr');
    ['公司名稱', '統編', '產業', '區域', '資本額', '名片數', '建立日', ''].forEach(function (h) { tr0.appendChild(el('th', null, h)); });
    t.appendChild(tr0);
    g.members.forEach(function (m) {
      var tr = el('tr');
      [m.name, m.taxId || '（無）', m.industry || '—', m.region || '—', fmtCap(m.capital), m.realContactCount + (m.contactCount !== m.realContactCount ? '（含待補 ' + (m.contactCount - m.realContactCount) + '）' : ''), (m.createdAt || '').slice(0, 10)]
        .forEach(function (v) { tr.appendChild(el('td', null, v)); });
      var td = el('td'); td.appendChild(btn('查看', 'cdup-btn cdup-view', function () { if (typeof window.openCompanyMasterDetail === 'function') window.openCompanyMasterDetail(m.id); }));
      tr.appendChild(td); t.appendChild(tr);
    });
    wrap.appendChild(t); c.appendChild(wrap);
    var bar = el('div', 'cdup-actions');
    bar.appendChild(g.dismissed ? btn('還原為疑似重複', 'cdup-btn', function () { dismiss(g, true); })
      : btn('標記為非重複', 'cdup-btn', function () { dismiss(g, false); }));
    c.appendChild(bar);
    return c;
  }

  function render() {
    panel.textContent = '';
    var s = data.summary;
    panel.appendChild(el('div', 'cdup-summary', s.open ? ('共 ' + s.open + ' 組疑似重複（高信心 ' + s.high + '、低信心 ' + s.low + '）') : '目前沒有疑似重複的企業主檔 ✅'));
    panel.appendChild(el('div', 'cdup-note', '此清單只供檢視：以統編與「去掉股份有限公司等後綴」的名稱比對。沒有自動合併；確認不是同一家請按「標記為非重複」，之後有新主檔加入同組時會自動再出現。'));
    var chk = el('label', 'cdup-check'), cb = el('input'); cb.type = 'checkbox'; cb.checked = showDismissed;
    cb.addEventListener('change', function () { showDismissed = cb.checked; render(); });
    chk.appendChild(cb); chk.appendChild(el('span', null, '一併顯示已標記為非重複的（' + s.dismissed + '）'));
    panel.appendChild(chk);
    data.groups.filter(function (g) { return showDismissed || !g.dismissed; }).forEach(function (g) { panel.appendChild(groupCard(g)); });
  }

  function mount() {
    host = document.getElementById('cmDupMount');
    if (!host) return;
    host.textContent = ''; opened = false; data = null;
    toggleBtn = btn('🔍 疑似重複主檔', 'cdup-toggle', function () {
      opened = !opened; panel.hidden = !opened; toggleBtn.setAttribute('aria-expanded', String(opened));
      toggleBtn.textContent = opened ? '🔍 收合疑似重複主檔' : '🔍 疑似重複主檔';
      if (opened) load();
    });
    toggleBtn.setAttribute('aria-expanded', 'false');
    panel = el('div', 'cdup-panel'); panel.hidden = true;
    host.appendChild(toggleBtn); host.appendChild(panel);
  }

  window.ITTSCompanyDups = { mount: mount };
})();
