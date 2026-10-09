// ═════════════════════════════════════════════════
// ── 報價牌價簿 (quote-pricebook.js) ───────────────────
// 全域 QPB：管理員在後台維護的「顧問角色 × 人天牌價／人天成本」（伺服器 GET /api/quote-pricebook，只回啟用項目，含牌價與成本）。
//   QPB.load({force})            取得牌價簿（60 秒內用快取；force＝一定重抓）→ Promise<{ok, items|error}>；失敗不丟錯，已載入過的舊資料仍可用 peek() 取得
//   QPB.peek()                   目前快取的項目陣列 [{id,bu,name,price,cost}]（bu 一定有值，缺漏＝ERP）；從未載入成功＝null（成本明細編輯器用：沒有就等於沒有牌價簿，行為不變）
//   QPB.openPicker(hooks)        「從牌價簿選取」對話框：每次開啟都重新向伺服器取最新清單；勾選＋人天數 → 加入報價單品項
//                                hooks = { getCurrent():目前品項列, onApply(items, info):套用到畫面, maxRows:列數上限（預設 50）, bu:預設分頁（可省略） }
//                                對話框依 BU 分四個分頁（ERP／ITS／MDM／CRM，各顯示項目數）；跨分頁勾選的項目會保留，「加入」一次加入所有分頁勾選的項目。
//                                預設分頁＝hooks.bu（有項目才採用），否則第一個有項目的 BU。目前報價單沒有可靠的 BU 欄位（lib/quoteRoutes.js 不收也不存 bu），所以呼叫端沒傳 bu。
//   純函式（可單元測試，不碰 DOM）：parseQty／marginPct／marginText／buildItems／isBlankDefaultRow／planAppend／sanitizeList／groupByBu／pickDefaultBu
// 選取是「複製」：品項只帶 desc／unit('人天')／qty／unitPrice／cat('consult')，沒有參照也沒有出處欄位，之後牌價簿怎麼改都不影響已存的報價單；
// 伺服器不檢查報價單價格是否等於牌價（牌價簿只是建議值）。
(function (global) {
'use strict';

var MAX_ROWS = 50;            // 與 quote.js QUOTE_MAX_ROWS、伺服器 MAX_ITEMS 相同
var QTY_MAX = 100000;
var TTL_MS = 60 * 1000;
var UNIT = '人天';
var ENDPOINT = '/api/quote-pricebook';
var BUS = ['ERP', 'ITS', 'MDM', 'CRM'];   // 與 lib/quotePricebook.js BUS 同值同序

// ── 純函式 ─────────────────────────────────────────
function esc(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}
function money(n) {
  var x = Number(n);
  if (!isFinite(x)) return '';
  return x.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

/** 人天數：純數字（可小數、最多 3 位），≥1 且 ≤ QTY_MAX；其餘回 null */
function parseQty(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!/^\d+(\.\d{1,3})?$/.test(s)) return null;
  var n = Number(s);
  return n >= 1 && n <= QTY_MAX ? n : null;
}

/** 毛利率（%）＝(牌價−成本)/牌價×100；牌價 ≤ 0 或資料不是有限數字回 null */
function marginPct(price, cost) {
  var p = Number(price), c = Number(cost);
  if (!isFinite(p) || !isFinite(c) || !(p > 0)) return null;
  return (p - c) / p * 100;
}
function marginText(price, cost) {
  var m = marginPct(price, cost);
  return m === null ? '—' : m.toFixed(1) + '%';
}

/** 伺服器回傳的清單 → 乾淨的 [{id,bu,name,price,cost}]（丟掉格式不對的項目；牌價與成本要是 ≥0 的有限數字；bu 缺漏或不合法＝ERP，相容舊資料） */
function sanitizeList(raw) {
  var out = [];
  (Array.isArray(raw) ? raw : []).forEach(function (x) {
    if (!x || typeof x !== 'object') return;
    var name = typeof x.name === 'string' ? x.name.trim() : '';
    var price = Number(x.price), cost = Number(x.cost);
    if (!x.id || !name || !isFinite(price) || price < 0 || !isFinite(cost) || cost < 0) return;
    out.push({ id: String(x.id), bu: BUS.indexOf(x.bu) >= 0 ? x.bu : 'ERP', name: name, price: price, cost: cost });
  });
  return out;
}

/** 依 BU 分組（各 BU 內保持原順序）：{ERP:[],ITS:[],MDM:[],CRM:[]}；缺 bu 的算 ERP */
function groupByBu(list) {
  var g = {};
  BUS.forEach(function (b) { g[b] = []; });
  (Array.isArray(list) ? list : []).forEach(function (it) { if (it) g[BUS.indexOf(it.bu) >= 0 ? it.bu : 'ERP'].push(it); });
  return g;
}

/** 預設分頁：hint 是合法 BU 且有項目 → hint；否則第一個有項目的 BU；全空 → ERP */
function pickDefaultBu(list, hint) {
  var g = groupByBu(list);
  if (hint && g[hint] && g[hint].length) return hint;
  for (var i = 0; i < BUS.length; i++) if (g[BUS[i]].length) return BUS[i];
  return BUS[0];
}

/** 依勾選結果產生報價品項（依牌價簿順序：ERP、ITS、MDM、CRM 分頁順序，各 BU 內依清單順序；不依點選順序；找不到的 id 略過）。picks = [{id, qty}]。品項說明只有名稱（不帶 BU） */
function buildItems(list, picks) {
  var want = Object.create(null);
  (Array.isArray(picks) ? picks : []).forEach(function (p) { if (p && p.id !== undefined) want[p.id] = p.qty; });
  var out = [];
  var g = groupByBu(list), flat = [];
  BUS.forEach(function (b) { flat = flat.concat(g[b]); });
  flat.forEach(function (it) {
    if (!it || !Object.prototype.hasOwnProperty.call(want, it.id)) return;
    out.push({ desc: it.name, unit: UNIT, qty: Number(want[it.id]), unitPrice: Number(it.price), cat: 'consult' });
  });
  return out;
}

/** 「還沒動過的預設空白列」：新單一開啟就有的那一列（沒有 lid／kind、品項說明空白、單位式、數量 1、單價 0、沒有成本與分類） */
function isBlankDefaultRow(it) {
  if (!it || typeof it !== 'object') return false;
  if (it.kind || it.lid || it.needPrice) return false;
  if (String(it.desc === undefined || it.desc === null ? '' : it.desc).trim() !== '') return false;
  if ((it.unit === undefined || it.unit === null ? '式' : String(it.unit).trim() || '式') !== '式') return false;
  if (Number(it.qty === undefined ? 1 : it.qty) !== 1) return false;
  if (Number(it.unitPrice || 0) !== 0) return false;
  if (it.cost !== undefined && it.cost !== null && it.cost !== '' && Number(it.cost) !== 0) return false;
  if (it.cat) return false;
  return true;
}

/**
 * 把新品項接到目前列表後面。目前只有「一列未動過的預設空白列」時，用新品項取代它。
 * 回 { ok:true, items, replacedBlank } 或 { ok:false, total, over, capacity }（加入後超過 maxRows；capacity＝最多還能加幾項）。不改動傳入的陣列。
 */
function planAppend(current, add, maxRows) {
  var cur = Array.isArray(current) ? current : [];
  var more = Array.isArray(add) ? add : [];
  var max = maxRows > 0 ? maxRows : MAX_ROWS;
  var replaced = cur.length === 1 && isBlankDefaultRow(cur[0]);
  var base = replaced ? [] : cur.slice();
  var total = base.length + more.length;
  if (total > max) return { ok: false, total: total, over: total - max, capacity: Math.max(0, max - base.length) };
  return { ok: true, items: base.concat(more), replacedBlank: replaced };
}

// ── 載入與快取 ─────────────────────────────────────
var state = { items: null, at: 0, inflight: null };
var fetchImpl = null;
function doFetch(url) {
  var f = fetchImpl || (global.fetch && global.fetch.bind(global));
  if (!f) return Promise.reject(new Error('fetch 不可用'));
  return f(url, { credentials: 'same-origin' });
}

function load(opts) {
  var force = !!(opts && opts.force);
  if (!force && state.items && Date.now() - state.at < TTL_MS) return Promise.resolve({ ok: true, items: state.items, cached: true });
  if (!force && state.inflight) return state.inflight;
  var p = doFetch(ENDPOINT).then(function (r) {
    if (!r.ok) { var e = new Error('HTTP ' + r.status); e.status = r.status; throw e; }
    return r.json();
  }).then(function (j) {
    var items = sanitizeList(j && j.items);
    state.items = items; state.at = Date.now();
    return { ok: true, items: items };
  }, function (e) {
    return { ok: false, error: (e && e.message) || '載入失敗', status: e && e.status };
  }).then(function (res) { if (state.inflight === p) state.inflight = null; return res; });
  state.inflight = p;
  return p;
}
function peek() { return state.items; }

// ── 對話框 ─────────────────────────────────────────
var CSS = '' +
  '.qpb-hint{font-size:12.5px;color:#6b7686;margin:0 0 10px;line-height:1.6}' +
  '.qpb-wrap{max-height:52vh;overflow:auto;border:1px solid #e3e6ea;border-radius:8px}' +
  '.qpb-tbl{width:100%;border-collapse:collapse;font-size:13.5px}' +
  '.qpb-tbl th{position:sticky;top:0;background:#f6f7f9;text-align:left;font-weight:600;color:#4b5563;padding:7px 8px;font-size:12.5px;white-space:nowrap}' +
  '.qpb-tbl td{padding:6px 8px;border-top:1px solid #eef0f3;vertical-align:middle}' +
  '.qpb-tbl .r{text-align:right;white-space:nowrap}' +
  '.qpb-tbl tr.qpb-on{background:#e8f0fe}' +
  '.qpb-qty{width:76px;border:1px solid #d9dde2;border-radius:6px;padding:4px 6px;text-align:right;font-size:13px}' +
  '.qpb-qty.qpb-bad{border-color:#ea4335;background:#fde8e8}' +
  '.qpb-err{color:#c5221f;font-size:13px;min-height:18px;margin-top:8px}' +
  '.qpb-neg{color:#c5221f;font-weight:600}' +
  '.qpb-tabs{display:flex;overflow-x:auto;overflow-y:hidden;border-bottom:1px solid #e3e6ea;margin:0 0 10px}' +
  '.qpb-tab{flex:0 0 auto;border:none;background:none;cursor:pointer;padding:8px 14px;font-size:14px;font-weight:600;color:#6b7686;white-space:nowrap;border-bottom:2px solid transparent}' +
  '.qpb-tab[aria-selected="true"]{color:#1a73e8;border-bottom-color:#1a73e8;font-weight:700}' +
  '.qpb-tab .qpb-n{display:inline-block;min-width:20px;text-align:center;font-size:12px;border-radius:10px;padding:0 6px;background:#f1f3f4;margin-left:4px}' +
  '.qpb-tab[aria-selected="true"] .qpb-n{background:#e8f0fe}.qpb-tab .qpb-sel{color:#0a8a4a;font-size:12px;margin-left:4px}' +
  'body.dark .qpb-tabs{border-bottom-color:#30363d}body.dark .qpb-tab{color:#8b949e}body.dark .qpb-tab[aria-selected="true"]{color:#58a6ff;border-bottom-color:#58a6ff}' +
  'body.dark .qpb-tab .qpb-n{background:#21262d}body.dark .qpb-tab[aria-selected="true"] .qpb-n{background:#14283f}body.dark .qpb-tab .qpb-sel{color:#3fb950}' +
  'body.dark .qpb-neg{color:#ff7b72}@media(max-width:480px){.qpb-m{display:none}}' +
  '.qpb-empty{text-align:center;color:#6b7686;padding:28px 10px;font-size:14px}' +
  'body.dark .qpb-wrap{border-color:#30363d}body.dark .qpb-tbl th{background:#21262d;color:#c9d1d9}' +
  'body.dark .qpb-tbl td{border-top-color:#30363d}body.dark .qpb-tbl tr.qpb-on{background:#14283f}' +
  'body.dark .qpb-qty{background:#0d1117;color:#e6edf3;border-color:#30363d}body.dark .qpb-hint{color:#8b949e}';
function ensureStyle() {
  if (typeof document === 'undefined' || document.getElementById('qpbStyle')) return;
  var st = document.createElement('style'); st.id = 'qpbStyle'; st.textContent = CSS; document.head.appendChild(st);
}

var closeCurrent = null;

function tableHtml(list, bu) {
  if (!list.length) return '<div class="qpb-empty">' + esc(bu) + ' 尚無牌價項目' + '</div>';
  var rows = list.map(function (it) {
    var m = marginPct(it.price, it.cost);
    return '<tr data-id="' + esc(it.id) + '">' +
      '<td style="width:34px;text-align:center"><input type="checkbox" class="qpb-cb" data-id="' + esc(it.id) + '" aria-label="選取 ' + esc(it.name) + '"></td>' +
      '<td>' + esc(it.name) + '</td>' +
      '<td class="r">' + esc(money(it.price)) + '</td>' +
      '<td class="r">' + esc(money(it.cost)) + '</td>' +
      '<td class="r qpb-m' + (m !== null && m < 0 ? ' qpb-neg' : '') + '">' + esc(marginText(it.price, it.cost)) + '</td>' +
      '<td class="r"><input type="text" class="qpb-qty" data-id="' + esc(it.id) + '" value="1" inputmode="decimal" aria-label="' + esc(it.name) + ' 人天數"></td></tr>';
  }).join('');
  return '<div class="qpb-wrap"><table class="qpb-tbl"><thead><tr><th></th><th>項目</th><th class="r">牌價 / 人天</th><th class="r">成本 / 人天</th><th class="r qpb-m">毛利率</th><th class="r">人天數</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
}

/** 整個對話框內容：說明＋分頁列＋四個分頁面板（只顯示目前分頁，其餘 hidden；DOM 都在，所以跨分頁勾選與人天數都保留）＋錯誤訊息 */
function dialogHtml(groups, active) {
  var tabs = BUS.map(function (b) {
    var on = b === active;
    return '<button type="button" class="qpb-tab" role="tab" id="qpbTab-' + b + '" data-bu="' + b + '" aria-selected="' + on + '" aria-controls="qpbPanel-' + b + '" tabindex="' + (on ? 0 : -1) + '">' +
      b + '<span class="qpb-n">' + groups[b].length + '</span><span class="qpb-sel" data-sel="' + b + '"></span></button>';
  }).join('');
  var panels = BUS.map(function (b) {
    return '<div role="tabpanel" id="qpbPanel-' + b + '" aria-labelledby="qpbTab-' + b + '" data-bu="' + b + '"' + (b === active ? '' : ' hidden') + '>' + tableHtml(groups[b], b) + '</div>';
  }).join('');
  return '<p class="qpb-hint">依事業單位（BU）分頁。勾選要加入的顧問角色並填人天數（跨分頁勾選會保留，一次加入）。加入後是「複製」進報價單：之後牌價簿調整不會影響這張單，單價也可以再改。</p>' +
    '<div class="qpb-tabs" role="tablist" aria-label="事業單位（BU）">' + tabs + '</div>' + panels +
    '<div class="qpb-err" role="alert" aria-live="assertive"></div>';
}

/**
 * hooks = { getCurrent():目前品項列（含 kind 列）, onApply(items, info):用新的完整品項陣列重畫畫面, maxRows }
 * 回傳 Promise<加入的項目數>（取消＝0）。
 */
function openPicker(hooks) {
  hooks = hooks || {};
  if (typeof document === 'undefined') return Promise.resolve(0);
  if (closeCurrent) closeCurrent(0);
  ensureStyle();
  var maxRows = hooks.maxRows > 0 ? hooks.maxRows : MAX_ROWS;
  var list = [];
  var activeBu = BUS[0];
  var ov = document.createElement('div');
  ov.className = 'modal-overlay open';
  ov.id = 'qPbOverlay';
  ov.style.zIndex = '125';
  ov.innerHTML = '<div class="modal q-dlg q-dlg-wide" role="dialog" aria-modal="true" aria-labelledby="qPbTitle">' +
    '<div class="modal-header"><h2 id="qPbTitle">從牌價簿選取</h2><button type="button" class="modal-close" data-pb="cancel" aria-label="關閉">&#10005;</button></div>' +
    '<div class="modal-body q-dlg-body" id="qPbBody"><div class="qpb-empty">載入中…</div></div>' +
    '<div class="modal-footer"><span id="qPbCount" style="margin-right:auto;font-size:13px;color:#6b7686"></span>' +
    '<button type="button" class="btn btn-secondary" data-pb="cancel">取消</button>' +
    '<button type="button" class="btn btn-primary" data-pb="ok" disabled>加入已勾選的項目</button></div></div>';
  document.body.appendChild(ov);
  var body = ov.querySelector('#qPbBody'), okBtn = ov.querySelector('[data-pb="ok"]'), countEl = ov.querySelector('#qPbCount');
  var resolveFn, done = false;
  var promise = new Promise(function (res) { resolveFn = res; });
  var opener = document.activeElement;

  function finish(n) {
    if (done) return;
    done = true;
    document.removeEventListener('keydown', onKey, true);
    if (closeCurrent === finish) closeCurrent = null;
    ov.remove();
    if (opener && opener.focus) { try { opener.focus(); } catch (e) { /* 忽略 */ } }
    resolveFn(n);
  }
  closeCurrent = finish;

  function onKey(e) {
    if (e.key === 'Escape') { e.stopPropagation(); finish(0); return; }
    if (e.key !== 'Tab') return;
    var f = Array.prototype.filter.call(ov.querySelectorAll('button,input'), function (x) { return !x.disabled && x.offsetParent !== null && x.getAttribute('tabindex') !== '-1'; });
    if (!f.length) return;
    var i = f.indexOf(document.activeElement);
    if (e.shiftKey && i <= 0) { e.preventDefault(); f[f.length - 1].focus(); }
    else if (!e.shiftKey && (i === f.length - 1 || i < 0)) { e.preventDefault(); f[0].focus(); }
  }
  document.addEventListener('keydown', onKey, true);

  function setErr(msg) { var el = ov.querySelector('.qpb-err'); if (el) el.textContent = msg || ''; }
  function checked() { return Array.prototype.filter.call(ov.querySelectorAll('.qpb-cb'), function (c) { return c.checked; }); }
  function setActive(bu, focusTab) {
    if (BUS.indexOf(bu) < 0) return;
    activeBu = bu;
    Array.prototype.forEach.call(ov.querySelectorAll('.qpb-tab'), function (t) {
      var on = t.getAttribute('data-bu') === bu;
      t.setAttribute('aria-selected', on ? 'true' : 'false'); t.setAttribute('tabindex', on ? '0' : '-1');
    });
    Array.prototype.forEach.call(ov.querySelectorAll('[role="tabpanel"]'), function (p) { p.hidden = p.getAttribute('data-bu') !== bu; });
    if (focusTab) { var t = ov.querySelector('.qpb-tab[data-bu="' + bu + '"]'); if (t) t.focus(); }
  }
  function refreshCount() {
    var n = checked().length;
    countEl.textContent = n ? '已勾選 ' + n + ' 項' : '';
    okBtn.disabled = n === 0;
    BUS.forEach(function (b) {
      var el = ov.querySelector('[data-sel="' + b + '"]');
      if (!el) return;
      var c = Array.prototype.filter.call(ov.querySelectorAll('[role="tabpanel"][data-bu="' + b + '"] .qpb-cb'), function (x) { return x.checked; }).length;
      el.textContent = c ? '✓' + c : '';
    });
    Array.prototype.forEach.call(ov.querySelectorAll('tbody tr'), function (tr) {
      var cb = tr.querySelector('.qpb-cb');
      tr.classList.toggle('qpb-on', !!(cb && cb.checked));
    });
  }

  function showError(msg) {
    body.innerHTML = '<div class="qpb-empty"><div style="color:#c5221f;margin-bottom:10px">牌價簿載入失敗：' + esc(msg) + '</div><button type="button" class="btn btn-secondary btn-sm" data-pb="retry">重新載入</button></div>';
    okBtn.disabled = true; countEl.textContent = '';
  }
  function fetchAndRender() {
    body.innerHTML = '<div class="qpb-empty">載入中…</div>';
    okBtn.disabled = true;
    load({ force: true }).then(function (res) {
      if (done) return;
      if (!res.ok) { showError(res.status === 403 ? '沒有使用牌價簿的權限' : res.error); return; }
      list = res.items;
      if (!list.length) { body.innerHTML = '<div class="qpb-empty">尚無牌價項目，請管理員到後台維護</div>'; okBtn.disabled = true; return; }
      activeBu = pickDefaultBu(list, hooks.bu);
      body.innerHTML = dialogHtml(groupByBu(list), activeBu);
      refreshCount();
      var first = body.querySelector('[role="tabpanel"]:not([hidden]) .qpb-cb') || body.querySelector('.qpb-tab[aria-selected="true"]'); if (first) first.focus();
    });
  }

  function confirm() {
    setErr('');
    Array.prototype.forEach.call(ov.querySelectorAll('.qpb-qty'), function (q) { q.classList.remove('qpb-bad'); });
    var cbs = checked();
    if (!cbs.length) { setErr('請至少勾選一個項目'); return; }
    var picks = [];
    for (var i = 0; i < cbs.length; i++) {
      var id = cbs[i].getAttribute('data-id');
      var qEl = ov.querySelector('.qpb-qty[data-id="' + id.replace(/"/g, '\\"') + '"]');
      var q = parseQty(qEl ? qEl.value : '');
      if (q === null) {
        var itm = list.filter(function (x) { return x.id === id; })[0];
        if (itm) setActive(itm.bu);   // 錯誤在別的分頁 → 切過去再標示
        if (qEl) { qEl.classList.add('qpb-bad'); qEl.focus(); qEl.select && qEl.select(); }
        setErr('「' + (list.filter(function (x) { return x.id === id; })[0] || {}).name + '」的人天數需為 1 以上的數字（可含小數，最多 3 位）');
        return;
      }
      picks.push({ id: id, qty: q });
    }
    var add = buildItems(list, picks);
    var cur = typeof hooks.getCurrent === 'function' ? hooks.getCurrent() : [];
    var plan = planAppend(cur, add, maxRows);
    if (!plan.ok) {
      setErr('加入後會有 ' + plan.total + ' 列，超過上限 ' + maxRows + ' 列（含分組標題與小計列）。目前最多還能加入 ' + plan.capacity + ' 項，請減少勾選或先精簡報價單。');
      return;
    }
    if (typeof hooks.onApply === 'function') hooks.onApply(plan.items, { added: add.length, replacedBlank: plan.replacedBlank });
    finish(add.length);
  }

  ov.addEventListener('click', function (e) {
    var b = e.target.closest ? e.target.closest('[data-pb]') : null;
    if (!b) return;
    var act = b.getAttribute('data-pb');
    if (act === 'cancel') finish(0);
    else if (act === 'ok') confirm();
    else if (act === 'retry') fetchAndRender();
  });
  ov.addEventListener('click', function (e) {
    var t = e.target.closest ? e.target.closest('.qpb-tab') : null;
    if (t) setActive(t.getAttribute('data-bu'), true);
  });
  ov.addEventListener('keydown', function (e) {
    var t = e.target.closest ? e.target.closest('.qpb-tab') : null;
    if (!t) return;
    var i = BUS.indexOf(t.getAttribute('data-bu')), j = -1;
    if (e.key === 'ArrowRight') j = (i + 1) % BUS.length;
    else if (e.key === 'ArrowLeft') j = (i + BUS.length - 1) % BUS.length;
    else if (e.key === 'Home') j = 0;
    else if (e.key === 'End') j = BUS.length - 1;
    if (j < 0) return;
    e.preventDefault(); setActive(BUS[j], true);
  });
  ov.addEventListener('change', function (e) { if (e.target.classList && e.target.classList.contains('qpb-cb')) { setErr(''); refreshCount(); } });
  ov.addEventListener('input', function (e) {
    if (e.target.classList && e.target.classList.contains('qpb-qty')) { e.target.classList.remove('qpb-bad'); setErr(''); }
  });
  ov.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && e.target.classList && e.target.classList.contains('qpb-qty') && !e.isComposing) { e.preventDefault(); if (!okBtn.disabled) confirm(); }
  });

  fetchAndRender();
  return promise;
}

global.QPB = {
  MAX_ROWS: MAX_ROWS, QTY_MAX: QTY_MAX, UNIT: UNIT, BUS: BUS,
  load: load, peek: peek, openPicker: openPicker,
  parseQty: parseQty, marginPct: marginPct, marginText: marginText, sanitizeList: sanitizeList,
  groupByBu: groupByBu, pickDefaultBu: pickDefaultBu,
  buildItems: buildItems, isBlankDefaultRow: isBlankDefaultRow, planAppend: planAppend,
  // 內部（單元測試用）
  _setFetch: function (f) { fetchImpl = f; },
  _reset: function () { state.items = null; state.at = 0; state.inflight = null; },
};
})(typeof window !== 'undefined' ? window : globalThis);
