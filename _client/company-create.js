// 企業主檔「單筆新增」對話框（行銷／管理員）。全部 createElement＋textContent。
// 用法：ITTSCompanyCreate.open({ industries: [...], onDone: fn })
(function () {
  'use strict';
  function el(tag, cls, text) { var n = document.createElement(tag); if (cls) n.className = cls; if (text != null) n.textContent = text; return n; }

  function open(opts) {
    opts = opts || {};
    var prev = document.activeElement;
    var overlay = el('div', 'ccr-overlay');
    var box = el('div', 'ccr-box'); box.setAttribute('role', 'dialog'); box.setAttribute('aria-modal', 'true'); box.setAttribute('aria-labelledby', 'ccrTitle');
    var title = el('h3', 'ccr-title', '＋ 新增企業主檔'); title.id = 'ccrTitle';
    var form = el('form', 'ccr-form'); form.noValidate = true;
    var inputs = {};
    function field(key, label, node, hint) {
      var w = el('label', 'ccr-field'); w.appendChild(el('span', 'ccr-label', label)); w.appendChild(node);
      if (hint) w.appendChild(el('span', 'ccr-hint', hint));
      form.appendChild(w); inputs[key] = node; return node;
    }
    function text(ph, max) { var i = el('input', 'ccr-input'); i.type = 'text'; i.maxLength = max; if (ph) i.placeholder = ph; i.autocomplete = 'off'; return i; }
    field('name', '公司名稱 *', text('公司正式名稱', 120));
    var tax = field('taxId', '統一編號', text('8 碼數字；國外客戶等無統編請留空', 8), '有統編時系統以統編判斷是否重複；無統編則以公司名稱比對');
    tax.inputMode = 'numeric';
    var sel = el('select', 'ccr-input');
    sel.appendChild(new Option('— 未分類 —', ''));
    (opts.industries || []).forEach(function (n) { sel.appendChild(new Option(n, n)); });
    field('industry', '產業', sel);
    field('region', '區域', text('例：台北、新竹', 40));
    field('customerCode', '客戶代號', text('對應 ERP 客戶代號，潛客可留空', 40));
    var poolW = el('label', 'ccr-check'); var pool = el('input'); pool.type = 'checkbox'; pool.checked = true;
    poolW.appendChild(pool); poolW.appendChild(el('span', null, ' 放入客戶池（建立「待補名片」，供 Manager 分配給業務）'));
    form.appendChild(poolW);
    var msg = el('div', 'ccr-msg'); msg.setAttribute('role', 'alert'); form.appendChild(msg);
    var simBox = el('div', 'ccr-sim'); simBox.hidden = true; form.appendChild(simBox);
    var bar = el('div', 'ccr-actions');
    var cancel = el('button', 'ccr-btn', '取消'); cancel.type = 'button';
    var ok = el('button', 'ccr-btn ccr-btn-primary', '建立'); ok.type = 'submit';
    bar.appendChild(cancel); bar.appendChild(ok); form.appendChild(bar);
    box.appendChild(title); box.appendChild(form); overlay.appendChild(box); document.body.appendChild(overlay);
    inputs.name.focus();

    function close() { document.removeEventListener('keydown', onKey, true); overlay.remove(); if (prev && prev.focus) try { prev.focus(); } catch (e) { /* */ } }
    function onKey(e) { if (e.key === 'Escape') { e.stopPropagation(); close(); } }
    document.addEventListener('keydown', onKey, true);
    cancel.addEventListener('click', close);
    overlay.addEventListener('mousedown', function (e) { if (e.target === overlay) close(); });

    var force = false;
    function submit() {
      msg.textContent = '';
      var body = { name: inputs.name.value.trim(), taxId: inputs.taxId.value.trim(), industry: sel.value, region: inputs.region.value.trim(), customerCode: inputs.customerCode.value.trim(), toPool: pool.checked, force: force };
      if (!body.name) { msg.textContent = '請輸入公司名稱'; inputs.name.focus(); return; }
      if (body.taxId && !/^\d{8}$/.test(body.taxId)) { msg.textContent = '統一編號須為 8 碼數字（無統編請留空）'; inputs.taxId.focus(); return; }
      ok.disabled = true;
      fetch('/api/companies', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
        .then(function (r) { return r.json().catch(function () { return {}; }).then(function (d) { return { r: r, d: d }; }); })
        .then(function (x) {
          ok.disabled = false;
          if (x.r.ok) { if (typeof window.showToast === 'function') window.showToast('✓ 已建立企業主檔「' + x.d.name + '」'); close(); if (opts.onDone) opts.onDone(x.d); return; }
          msg.textContent = x.d.error || ('建立失敗（' + x.r.status + '）');
          if (x.d.code === 'SIMILAR') {
            simBox.hidden = false; simBox.textContent = '';
            simBox.appendChild(el('div', 'ccr-sim-t', '名稱相近的既有主檔：'));
            (x.d.similar || []).forEach(function (s) { simBox.appendChild(el('div', null, '・' + s.name + '（' + (s.taxId || '無統編') + '）')); });
            ok.textContent = '確認不是同一家，仍要建立'; force = true;
          }
        })
        .catch(function () { ok.disabled = false; msg.textContent = '網路錯誤，請稍後再試'; });
    }
    form.addEventListener('submit', function (e) { e.preventDefault(); submit(); });
    // 改了公司名稱／統編就取消「強制建立」狀態，須重新比對
    ['name', 'taxId'].forEach(function (k) { inputs[k].addEventListener('input', function () { if (force) { force = false; ok.textContent = '建立'; simBox.hidden = true; msg.textContent = ''; } }); });
  }

  window.ITTSCompanyCreate = { open: open };
})();
