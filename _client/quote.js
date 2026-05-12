// ════════════════════════════════════════════════════════════
// ── 報價單管理 (quote.js) ────────────────────────────────────
// ════════════════════════════════════════════════════════════

let allQuotations = [];
let _pnlCosts = []; // 與 items[] 同步的成本陣列

const QUOTE_STATUS_LABEL = { draft:'草稿', sent:'已寄出', accepted:'已接受', rejected:'已拒絕' };
const QUOTE_STATUS_CLASS = { draft:'quote-st-draft', sent:'quote-st-sent', accepted:'quote-st-accepted', rejected:'quote-st-rejected' };

/**
 * 計算報價合計
 * discountType: 'none' | 'percent' | 'amount'
 * discountValue: 數字（percent 時為百分比如 90 代表九折；amount 時為直接議價金額）
 */
function quoteTotal(items, discountType, discountValue) {
  const sub = (items || []).reduce(
    (s, it) => s + (parseFloat(it.qty) || 0) * (parseFloat(it.unitPrice) || 0), 0
  );
  let discounted = sub;
  let discountAmt = 0;
  if (discountType === 'percent') {
    const pct = parseFloat(discountValue) || 100; // 百分比，如 90 = 九折
    discounted  = sub * pct / 100;
    discountAmt = sub - discounted;
  } else if (discountType === 'amount') {
    discounted  = parseFloat(discountValue) || sub;
    discountAmt = sub - discounted;
  }
  const tax   = discounted * 0.05;
  const total = discounted + tax;
  return { sub, discounted, discountAmt, tax, total };
}

function fmtMoney(n) {
  return 'NT$ ' + Math.round(n).toLocaleString();
}

// ── 報價單清單 ────────────────────────────────────────────────
async function loadQuotationsView() {
  try {
    const r = await fetch(`${API}/quotations`);
    if (!r.ok) throw new Error();
    allQuotations = await r.json();
  } catch(e) {
    allQuotations = [];
  }
  renderQuoteList();
  bindQuoteListHandlers();
}

function renderQuoteList() {
  const search   = ($('quoteSearchInput')  ? $('quoteSearchInput').value   : '').toLowerCase();
  const stFilter = ($('quoteStatusFilter') ? $('quoteStatusFilter').value  : '');
  let list = allQuotations;
  if (search)   list = list.filter(q =>
    (q.quoteNo     || '').toLowerCase().includes(search) ||
    (q.company     || '').toLowerCase().includes(search) ||
    (q.contactName || '').toLowerCase().includes(search) ||
    (q.projectName || '').toLowerCase().includes(search)
  );
  if (stFilter) list = list.filter(q => q.status === stFilter);

  const tbody = $('quoteTbody');
  if (!tbody) return;
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-msg" style="text-align:center;padding:32px">尚無報價單資料，點擊「新增報價單」開始建立</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(q => {
    const { total } = quoteTotal(q.items, q.discountType, q.discountValue);
    const stClass = QUOTE_STATUS_CLASS[q.status] || 'quote-st-draft';
    const stLabel = QUOTE_STATUS_LABEL[q.status] || q.status;
    return `<tr>
      <td><span class="quote-no">${escapeHtml(q.quoteNo || '')}</span></td>
      <td>${escapeHtml(q.company || '')}</td>
      <td>${escapeHtml(q.contactName || '')}</td>
      <td>${escapeHtml(q.projectName || '')}</td>
      <td>${escapeHtml(q.quoteDate || '')}</td>
      <td style="text-align:right;font-weight:600;font-size:13px">${fmtMoney(total)}</td>
      <td><span class="quote-status ${stClass}">${escapeHtml(stLabel)}</span></td>
      <td style="white-space:nowrap">
        <button class="btn btn-sm" onclick="openQuoteModal('${q.id}')">✏️ 編輯</button>
        <button class="btn btn-sm btn-export" onclick="exportQuote('${q.id}','${escapeHtml(q.quoteNo || '')}')">&#11015; Excel</button>
        <button class="btn btn-sm btn-soft-danger" onclick="deleteQuote('${q.id}')">🗑️</button>
      </td>
    </tr>`;
  }).join('');
}

function bindQuoteListHandlers() {
  const btn = $('addQuoteBtn');
  if (btn && !btn._qListBound) {
    btn._qListBound = true;
    btn.addEventListener('click', function() { openQuoteModal(null); });
  }
  const si = $('quoteSearchInput');
  if (si && !si._qListBound) {
    si._qListBound = true;
    si.addEventListener('input', renderQuoteList);
  }
  const sf = $('quoteStatusFilter');
  if (sf && !sf._qListBound) {
    sf._qListBound = true;
    sf.addEventListener('change', renderQuoteList);
  }
}

// ── 匯出 Excel ──────────────────────────────────────────────
async function exportQuote(id, quoteNo) {
  try {
    showToast('正在產生報價單…');
    const r = await fetch(`${API}/quotations/${id}/export`);
    if (!r.ok) {
      const e = await r.json().catch(() => ({}));
      return showToast(e.error || '匯出失敗');
    }
    const blob = await r.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `${quoteNo || 'quotation'}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showToast('報價單 Excel 已下載');
  } catch(e) {
    showToast('匯出失敗，請重試');
  }
}

// ── 刪除報價單 ──────────────────────────────────────────────
async function deleteQuote(id) {
  if (!confirm('確定要刪除此報價單嗎？此動作無法復原。')) return;
  try {
    const r = await fetch(`${API}/quotations/${id}`, { method: 'DELETE' });
    if (!r.ok) return showToast('刪除失敗');
    showToast('已刪除報價單');
    loadQuotationsView();
  } catch(e) {
    showToast('刪除失敗，請重試');
  }
}

// ════════════════════════════════════════════════════════════
// ── PNL 毛利分析 TAB ─────────────────────────────────────────
// ════════════════════════════════════════════════════════════

/** TAB 切換 */
function switchQuoteTab(tabName) {
  ['info', 'pnl'].forEach(function(t) {
    var content = $('quoteTabContent' + (t === 'info' ? 'Info' : 'Pnl'));
    var btn     = $('quoteTabBtn'     + (t === 'info' ? 'Info' : 'Pnl'));
    if (!content || !btn) return;
    content.style.display = (t === tabName) ? '' : 'none';
    btn.classList.toggle('active', t === tabName);
  });
  if (tabName === 'pnl') renderPnlTab();
}

/** 從 PNL 表格讀取目前成本值（若表格已渲染）*/
function syncPnlCostsFromTable() {
  var inputs = document.querySelectorAll('#pnlItemsBody .pnl-cost-input');
  if (!inputs.length) return;
  inputs.forEach(function(inp) {
    var idx = parseInt(inp.dataset.idx, 10);
    _pnlCosts[idx] = parseFloat(inp.value) || 0;
  });
}

/** 毛利率顏色分級 */
function marginClass(pct) {
  if (pct >= 30) return 'good';
  if (pct >= 15) return 'warn';
  return 'danger';
}

/** 渲染 PNL 表格 */
function renderPnlTab() {
  // 先把目前表格的成本值同步回 _pnlCosts（避免切換 tab 後遺失）
  syncPnlCostsFromTable();

  var items = readQuoteItems();

  // 確保 _pnlCosts 長度與 items 一致
  while (_pnlCosts.length < items.length) _pnlCosts.push(0);
  _pnlCosts.length = items.length;

  var tbody = $('pnlItemsBody');
  tbody.innerHTML = items.map(function(it, i) {
    var qty   = parseFloat(it.qty)       || 0;
    var price = parseFloat(it.unitPrice) || 0;
    var cost  = parseFloat(_pnlCosts[i]) || 0;
    var revSub  = qty * price;
    var costSub = qty * cost;
    var gp      = revSub - costSub;
    var gpPct   = revSub > 0 ? (gp / revSub * 100) : 0;
    var gpColor = gp >= 0 ? '#2e7d32' : '#c62828';
    var pctColor= gpPct >= 30 ? '#2e7d32' : gpPct >= 15 ? '#e65100' : '#c62828';
    return '<tr>' +
      '<td style="text-align:center;color:#999;font-size:12px">' + (i + 1) + '</td>' +
      '<td style="font-size:13px">' + escapeHtml(it.desc || '（未填）') + '</td>' +
      '<td style="text-align:right;font-size:13px">' + qty + ' ' + escapeHtml(it.unit || '') + '</td>' +
      '<td style="text-align:right;font-size:13px">' + fmtMoney(price) + '</td>' +
      '<td style="text-align:right;font-size:13px;font-weight:500">' + fmtMoney(revSub) + '</td>' +
      '<td style="text-align:right"><input type="number" class="pnl-cost-input" data-idx="' + i + '" value="' + cost + '" min="0" step="1" placeholder="輸入成本"></td>' +
      '<td class="pnl-cst-sub" style="text-align:right;font-size:13px">' + fmtMoney(costSub) + '</td>' +
      '<td class="pnl-gp-cell" style="text-align:right;font-size:13px;font-weight:600;color:' + gpColor + '">' + fmtMoney(gp) + '</td>' +
      '<td class="pnl-pct-cell" style="text-align:right;font-size:13px;font-weight:700;color:' + pctColor + '">' + gpPct.toFixed(1) + '%</td>' +
      '</tr>';
  }).join('');

  // 成本輸入事件：即時更新列 + 匯總
  tbody.querySelectorAll('.pnl-cost-input').forEach(function(inp) {
    inp.addEventListener('input', function() {
      var idx   = parseInt(this.dataset.idx, 10);
      var cost  = parseFloat(this.value) || 0;
      _pnlCosts[idx] = cost;

      var row   = this.closest('tr');
      var qty   = parseFloat(readQuoteItems()[idx].qty) || 0;
      var price = parseFloat(readQuoteItems()[idx].unitPrice) || 0;
      var revSub  = qty * price;
      var costSub = qty * cost;
      var gp      = revSub - costSub;
      var gpPct   = revSub > 0 ? (gp / revSub * 100) : 0;

      row.querySelector('.pnl-cst-sub').textContent   = fmtMoney(costSub);
      row.querySelector('.pnl-gp-cell').textContent   = fmtMoney(gp);
      row.querySelector('.pnl-gp-cell').style.color   = gp >= 0 ? '#2e7d32' : '#c62828';
      row.querySelector('.pnl-pct-cell').textContent  = gpPct.toFixed(1) + '%';
      row.querySelector('.pnl-pct-cell').style.color  = gpPct >= 30 ? '#2e7d32' : gpPct >= 15 ? '#e65100' : '#c62828';

      updatePnlSummary();
    });
  });

  // 折扣提示
  var discNote = $('pnlDiscountNote');
  var { discountType, discountValue } = readQuoteDiscount();
  if (discountType !== 'none' && discountValue) {
    discNote.style.display = '';
    if (discountType === 'percent') {
      discNote.innerHTML = '⚡ 已套用 <strong>' + discountValue + '%</strong> 折扣，毛利以折扣後金額為基準計算。';
    } else {
      discNote.innerHTML = '⚡ 已套用議價總額 <strong>' + fmtMoney(discountValue) + '</strong>，毛利以議價金額為基準計算。';
    }
  } else {
    discNote.style.display = 'none';
  }

  updatePnlSummary();
}

/** 更新 PNL 匯總區塊 */
function updatePnlSummary() {
  var items = readQuoteItems();
  var { discountType, discountValue } = readQuoteDiscount();
  var totals = quoteTotal(items, discountType, discountValue);
  var revenue = totals.discounted; // 優惠後未稅

  var totalCost = items.reduce(function(sum, it, i) {
    return sum + (parseFloat(it.qty) || 0) * (parseFloat(_pnlCosts[i]) || 0);
  }, 0);

  var gp      = revenue - totalCost;
  var gpPct   = revenue > 0 ? (gp / revenue * 100) : 0;
  var mc      = marginClass(gpPct);

  $('pnlRevenue').textContent     = fmtMoney(revenue);
  $('pnlCostTotal').textContent   = fmtMoney(totalCost);

  var gpEl = $('pnlGrossProfit');
  gpEl.textContent = fmtMoney(gp);
  gpEl.className   = 'pnl-sum-value pnl-gp-val ' + (gp >= 0 ? 'positive' : 'negative');

  var pctEl = $('pnlMarginPct');
  pctEl.textContent = gpPct.toFixed(1) + '%';
  pctEl.className   = 'pnl-sum-value pnl-margin-val ' + mc;
}

// ── 關閉 Modal ───────────────────────────────────────────────
function closeQuoteModal() {
  $('quoteModalOverlay').style.display = 'none';
}

// ── 開啟新增 / 編輯 Modal ────────────────────────────────────
async function openQuoteModal(idOrNull) {
  try {
    // 確保聯絡人資料已載入
    if (typeof allContacts === 'undefined' || !allContacts || allContacts.length === 0) {
      try {
        const r = await fetch(API + '/contacts');
        if (r.ok) { window.allContacts = await r.json(); }
      } catch (fetchErr) {
        console.warn('載入聯絡人失敗', fetchErr);
      }
    }

    const contacts = (typeof allContacts !== 'undefined' && allContacts) ? allContacts : [];
    const q = idOrNull ? ((allQuotations || []).find(function(x){ return x.id === idOrNull; }) || null) : null;
    const today = new Date().toISOString().slice(0, 10);

    $('quoteModalTitle').textContent = q ? ('編輯報價單 ' + (q.quoteNo || '')) : '新增報價單';
    $('quoteId').value      = q ? (q.id           || '') : '';
    $('qCompany').value     = q ? (q.company       || '') : '';
    $('qPhone').value       = q ? (q.phone         || '') : '';
    $('qMobile').value      = q ? (q.mobile        || '') : '';
    $('qAddress').value     = q ? (q.address       || '') : '';
    $('qDate').value        = q ? (q.quoteDate     || today) : today;
    $('qStatus').value      = q ? (q.status        || 'draft') : 'draft';
    $('qProjectName').value = q ? (q.projectName   || '') : '';
    $('qProjectNo').value   = q ? (q.projectNo     || '') : '';
    $('qNote').value        = q ? (q.note          || '') : '';

    // 公司 datalist
    var companySet = new Set(contacts.map(function(c){ return c.company; }).filter(Boolean));
    $('quoteCompanyList').innerHTML = Array.from(companySet).sort()
      .map(function(c){ return '<option value="' + escapeHtml(c) + '">'; }).join('');

    // 聯絡人下拉
    buildQuoteContactSelect(q ? (q.contactId || '') : '', q ? (q.company || '') : '');

    // 項目列表
    var items = (q && q.items && q.items.length)
      ? q.items
      : [{ desc: '', unit: '式', qty: 1, unitPrice: 0 }];

    // 初始化 PNL 成本陣列（從已存資料載入 cost）
    _pnlCosts = items.map(function(it) { return parseFloat(it.cost) || 0; });

    renderQuoteItems(items);

    // ── 優惠設定還原 ──
    var discType  = (q && q.discountType)  ? q.discountType  : 'none';
    var discValue = (q && q.discountValue != null && q.discountValue !== 0) ? q.discountValue : '';
    document.querySelectorAll('input[name="qDiscountType"]').forEach(function(radio) {
      radio.checked = (radio.value === discType);
    });
    $('qDiscountValue').value = discValue;
    applyDiscountMode(discType);

    // ── 事件綁定（用 overlay 旗標確保只綁一次）──
    var overlay = $('quoteModalOverlay');

    if (!overlay._qBound) {
      overlay._qBound = true;

      // 優惠模式切換
      document.querySelectorAll('input[name="qDiscountType"]').forEach(function(radio) {
        radio.addEventListener('change', function() {
          applyDiscountMode(this.value);
          updateQuoteTotals();
        });
      });
      $('qDiscountValue').addEventListener('input', updateQuoteTotals);

      // 新增項目
      $('addQuoteItemBtn').addEventListener('click', function() {
        syncPnlCostsFromTable();
        var current = readQuoteItems();
        current.push({ desc: '', unit: '式', qty: 1, unitPrice: 0 });
        _pnlCosts.push(0);
        renderQuoteItems(current);
        updateQuoteTotals();
      });

      // 公司輸入時篩聯絡人
      $('qCompany').addEventListener('input', function() {
        buildQuoteContactSelect('', this.value);
      });

      // 聯絡人選擇後自動帶入
      $('qContactId').addEventListener('change', autoFillFromContact);

      // 關閉事件
      $('quoteModalClose').addEventListener('click', closeQuoteModal);
      $('quoteModalCancel').addEventListener('click', closeQuoteModal);
    }

    // 儲存按鈕 — 每次重設 onclick 避免舊 context 殘留
    $('quoteModalSave').onclick = saveQuote;

    // 每次開啟都從第一頁開始
    switchQuoteTab('info');

    overlay.style.display = 'flex';
    updateQuoteTotals();

  } catch (err) {
    console.error('openQuoteModal 錯誤:', err);
    showToast('開啟報價單失敗：' + (err.message || err));
  }
}

// ── 優惠模式切換 UI ───────────────────────────────────────────
function applyDiscountMode(type) {
  const wrap   = $('qDiscountInputWrap');
  const prefix = $('qDiscountPrefix');
  const suffix = $('qDiscountSuffix');
  const input  = $('qDiscountValue');

  if (type === 'none') {
    wrap.style.display = 'none';
  } else {
    wrap.style.display = '';
    if (type === 'percent') {
      prefix.textContent    = '折扣百分比';
      suffix.textContent    = '（例：90 = 九折，85 = 八五折）';
      input.placeholder     = '請輸入 0–100 的數字';
      input.min = '0'; input.max = '100'; input.step = '0.1';
    } else if (type === 'amount') {
      prefix.textContent    = '議價總額（未稅）';
      suffix.textContent    = '直接輸入業務談好的未稅金額';
      input.placeholder     = '請輸入金額（NT$）';
      input.min = '0'; input.max = ''; input.step = '1';
    }
  }
}

// ── 讀取目前優惠設定 ──────────────────────────────────────────
function readQuoteDiscount() {
  const type  = (document.querySelector('input[name="qDiscountType"]:checked') || {}).value || 'none';
  const value = parseFloat($('qDiscountValue').value) || 0;
  return { discountType: type, discountValue: value };
}

// ── 聯絡人下拉選單建構 ────────────────────────────────────────
function buildQuoteContactSelect(selectedId, filterCompany) {
  const sel = $('qContactId');
  let contacts = allContacts || [];
  if (filterCompany) {
    contacts = contacts.filter(c => c.company === filterCompany);
  }
  sel.innerHTML = '<option value="">-- 選擇聯絡人（自動帶入資料）--</option>' +
    contacts.map(c =>
      '<option value="' + c.id + '"' + (c.id === selectedId ? ' selected' : '') + '>' +
      escapeHtml(c.name || '') + (c.company ? (' - ' + escapeHtml(c.company)) : '') +
      '</option>'
    ).join('');
}

// ── 聯絡人自動帶入 ────────────────────────────────────────────
function autoFillFromContact() {
  const id = $('qContactId').value;
  if (!id) return;
  const c = (allContacts || []).find(x => x.id === id);
  if (!c) return;
  if (c.company)  { $('qCompany').value  = c.company; }
  if (c.phone)    { $('qPhone').value    = c.phone;   }
  if (c.mobile)   { $('qMobile').value   = c.mobile;  }
  if (c.address)  { $('qAddress').value  = c.address; }
  buildQuoteContactSelect(id, c.company);
}

// ── 項目列表渲染 ─────────────────────────────────────────────
function renderQuoteItems(items) {
  const tbody = $('quoteItemsBody');
  tbody.innerHTML = items.map(function (it, i) {
    const qty   = parseFloat(it.qty)       || 1;
    const price = parseFloat(it.unitPrice) || 0;
    return '<tr data-idx="' + i + '">' +
      '<td style="text-align:center;color:#999;font-size:12px">' + (i + 1) + '</td>' +
      '<td><input type="text" class="qi-desc" value="' + escapeHtml(it.desc || '') + '" ' +
        'placeholder="品項說明" style="width:100%;border:1px solid #ddd;border-radius:4px;padding:5px 8px;font-size:13px;box-sizing:border-box"></td>' +
      '<td><input type="text" class="qi-unit" value="' + escapeHtml(it.unit || '式') + '" ' +
        'style="width:54px;border:1px solid #ddd;border-radius:4px;padding:5px 6px;font-size:13px;text-align:center"></td>' +
      '<td><input type="number" class="qi-qty" value="' + qty + '" min="0.001" step="1" ' +
        'style="width:64px;border:1px solid #ddd;border-radius:4px;padding:5px 6px;font-size:13px;text-align:right"></td>' +
      '<td><input type="number" class="qi-price" value="' + price + '" min="0" step="1" ' +
        'style="width:104px;border:1px solid #ddd;border-radius:4px;padding:5px 6px;font-size:13px;text-align:right"></td>' +
      '<td class="qi-subtotal" style="text-align:right;font-size:13px;padding-right:6px;white-space:nowrap">' +
        fmtMoney(qty * price) + '</td>' +
      '<td style="text-align:center"><button type="button" class="qi-remove" ' +
        'title="移除此項目" style="background:none;border:none;color:#e53935;cursor:pointer;font-size:16px;line-height:1;padding:2px 4px">&#10005;</button></td>' +
      '</tr>';
  }).join('');

  // 輸入事件：即時更新小計
  tbody.querySelectorAll('input').forEach(function (inp) {
    inp.addEventListener('input', function () {
      const row   = inp.closest('tr');
      const qty   = parseFloat(row.querySelector('.qi-qty').value)   || 0;
      const price = parseFloat(row.querySelector('.qi-price').value) || 0;
      row.querySelector('.qi-subtotal').textContent = fmtMoney(qty * price);
      updateQuoteTotals();
    });
  });

  // 移除按鈕
  tbody.querySelectorAll('.qi-remove').forEach(function (btn) {
    btn.addEventListener('click', function () {
      syncPnlCostsFromTable();
      var current = readQuoteItems();
      if (current.length <= 1) { showToast('至少需保留一個報價項目'); return; }
      var idx = parseInt(btn.closest('tr').dataset.idx);
      current.splice(idx, 1);
      _pnlCosts.splice(idx, 1);
      renderQuoteItems(current);
      updateQuoteTotals();
    });
  });
}

// ── 讀取目前項目列表 ──────────────────────────────────────────
function readQuoteItems() {
  return Array.from($('quoteItemsBody').querySelectorAll('tr')).map(function (row) {
    return {
      desc:      row.querySelector('.qi-desc').value.trim(),
      unit:      row.querySelector('.qi-unit').value.trim() || '式',
      qty:       parseFloat(row.querySelector('.qi-qty').value)   || 1,
      unitPrice: parseFloat(row.querySelector('.qi-price').value) || 0,
    };
  });
}

// ── 更新合計顯示 ─────────────────────────────────────────────
function updateQuoteTotals() {
  const items = readQuoteItems();
  const { discountType, discountValue } = readQuoteDiscount();
  const { sub, discounted, discountAmt, tax, total } = quoteTotal(items, discountType, discountValue);

  $('qSubtotal').textContent = fmtMoney(sub);
  $('qTax').textContent      = fmtMoney(tax);
  $('qTotal').textContent    = fmtMoney(total);

  // 優惠折扣列 & 優惠價列
  const hasDiscount = discountType !== 'none' && discountAmt !== 0;
  $('qDiscountRow').style.display    = hasDiscount ? '' : 'none';
  $('qDiscountedRow').style.display  = hasDiscount ? '' : 'none';

  if (hasDiscount) {
    if (discountType === 'percent') {
      $('qDiscountRowLabel').textContent = '優惠折扣（' + discountValue + '%）';
    } else {
      $('qDiscountRowLabel').textContent = '優惠折扣（議價）';
    }
    $('qDiscountAmt').textContent      = '- ' + fmtMoney(discountAmt);
    $('qDiscountedPrice').textContent  = fmtMoney(discounted);
  }
}

// ── 儲存報價單 ──────────────────────────────────────────────
async function saveQuote() {
  const company   = $('qCompany').value.trim();
  const quoteDate = $('qDate').value;
  if (!company)   { showToast('請輸入客戶公司名稱'); return; }
  if (!quoteDate) { showToast('請選擇報價日期');     return; }

  const id             = $('quoteId').value;
  const contactSel     = $('qContactId');
  const selOpt         = contactSel.selectedOptions[0];
  const autoName       = selOpt && selOpt.value
    ? selOpt.textContent.split(' - ')[0].trim()
    : '';

  const { discountType, discountValue } = readQuoteDiscount();

  // 儲存前先同步 PNL 成本（如果 PNL tab 已開過）
  syncPnlCostsFromTable();
  const items = readQuoteItems().map(function(it, i) {
    it.cost = parseFloat(_pnlCosts[i]) || 0;
    return it;
  });

  const payload = {
    contactId:     contactSel.value,
    company:       company,
    contactName:   autoName,
    phone:         $('qPhone').value.trim(),
    mobile:        $('qMobile').value.trim(),
    address:       $('qAddress').value.trim(),
    quoteDate:     quoteDate,
    status:        $('qStatus').value,
    projectName:   $('qProjectName').value.trim(),
    projectNo:     $('qProjectNo').value.trim(),
    items:         items,
    discountType:  discountType,
    discountValue: discountValue,
    note:          $('qNote').value.trim(),
  };

  try {
    const method = id ? 'PUT'  : 'POST';
    const url    = id ? (API + '/quotations/' + id) : (API + '/quotations');
    const r = await fetch(url, {
      method:  method,
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload),
    });
    if (!r.ok) {
      const e = await r.json().catch(function () { return {}; });
      showToast(e.error || '儲存失敗');
      return;
    }
    $('quoteModalOverlay').style.display = 'none';
    showToast(id ? '報價單已更新' : '報價單已建立 ✅');
    loadQuotationsView();
  } catch(e) {
    showToast('儲存失敗，請重試');
  }
}
