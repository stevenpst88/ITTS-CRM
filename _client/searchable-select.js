/* 可搜尋下拉（ITTSSearchSelect）：把既有 <select> 升級成「點開有輸入框、即時縮小清單」的自訂面板。
 * 設計：原 <select> 保留（隱藏）當唯一資料來源——.value 與 change 事件照舊，呼叫端程式不用改。
 * 兩種模式：單選（預設，.value 與 change 事件照舊）；多選 { multiple: true }——面板內每列一個 checkbox，不勾＝全部（沒有「全部」列），
 *   勾選集合以 getSelected() 取得（原 select 設為 multiple，選項的 selected 同步，但 .value 只反映第一個，呼叫端請改用 getSelected()）；
 *   每次勾選／全選／清除即時對原 select 觸發 change（與「全部階段」多選一致）。多選鍵盤：↑↓ 移動、Enter 勾選／取消（面板保持開啟）、
 *   Space 在搜尋框為空時勾選（否則照常輸入空白）、Esc 關閉並回到按鈕、Tab 離開。
 * 純邏輯（matchItem / filterItems / sortItems / moveIndex / toggleValue / applyVisible / buttonText）不碰 DOM，Node 與瀏覽器共用，可單獨測試。
 * 全部以 createElement＋textContent 建立（不使用任何 HTML 字串 sink）。
 */
(function (root) {
  'use strict';

  function norm(s) { return String(s == null ? '' : s).normalize('NFKC').toLowerCase().trim(); }

  // item: { value, label, keywords?: string[] }；查詢字串以空白切詞，每一詞都必須命中 label／value／keywords 任一（不分大小寫）
  function matchItem(item, query) {
    const terms = norm(query).split(/\s+/).filter(Boolean);
    if (!terms.length) return true;
    const hay = [item.label, item.value].concat(item.keywords || []).map(norm);
    return terms.every(t => hay.some(h => h.indexOf(t) !== -1));
  }

  // pinned（value === ''，即「全部」）永遠保留在最上面，不受查詢影響
  function filterItems(items, query) {
    return items.filter(it => it.value === '' || matchItem(it, query));
  }

  // 依顯示名稱排序；value === '' 置頂
  function sortItems(items) {
    const coll = (typeof Intl !== 'undefined' && Intl.Collator) ? new Intl.Collator('zh-Hant', { sensitivity: 'base', numeric: true }) : null;
    const cmp = (a, b) => coll ? coll.compare(a, b) : (a < b ? -1 : a > b ? 1 : 0);
    return items.slice().sort((a, b) => {
      if (a.value === '' && b.value !== '') return -1;
      if (b.value === '' && a.value !== '') return 1;
      return cmp(String(a.label), String(b.label)) || cmp(String(a.value), String(b.value));
    });
  }

  // 鍵盤移動：回傳新的 active index（-1 表示沒有）；上下鍵循環，Home/End 到頭尾
  function moveIndex(cur, key, len) {
    if (!len) return -1;
    if (key === 'ArrowDown') return cur < 0 || cur >= len - 1 ? 0 : cur + 1;
    if (key === 'ArrowUp') return cur <= 0 ? len - 1 : cur - 1;
    if (key === 'Home') return 0;
    if (key === 'End') return len - 1;
    return cur;
  }

  // 多選純邏輯：集合以陣列表示，回傳新陣列、不改原陣列
  function toggleValue(selected, value) {
    return selected.indexOf(value) === -1 ? selected.concat([value]) : selected.filter(v => v !== value);
  }
  // 「全選／清除」只作用在目前可見（搜尋結果）的項目；mode: 'all' | 'clear'
  function applyVisible(selected, visibleValues, mode) {
    const vis = new Set(visibleValues);
    if (mode === 'clear') return selected.filter(v => !vis.has(v));
    const out = selected.slice();
    visibleValues.forEach(v => { if (out.indexOf(v) === -1) out.push(v); });
    return out;
  }
  // 按鈕文字：未選＝全部；1 人＝名字；2 人且夠短＝「A、B」；其餘「已選 N 人」（窄寬度不爆版）
  function buttonText(labels, allLabel) {
    const n = labels.length;
    if (n === 0) return allLabel;
    if (n === 1) return labels[0];
    if (n === 2 && (labels[0] + labels[1]).length <= 8) return labels[0] + '、' + labels[1];
    return '已選 ' + n + ' 人';
  }

  // 面板夾邊：panel 以 left:0 錨在按鈕左緣；回傳需要加到 left 的位移（px），使面板落在 [margin, vw-margin] 內。
  // 右緣超出 → 往左移；移完左緣仍小於 margin（面板比視窗寬）→ 以左緣為準。純函式，可單獨測。
  function clampShift(left, right, vw, margin) {
    let dx = 0;
    if (right > vw - margin) dx = (vw - margin) - right;
    if (left + dx < margin) dx = margin - left;
    return dx;
  }

  function enhance(select, opts) {
    opts = opts || {};
    const doc = select.ownerDocument;
    const placeholder = opts.placeholder || '搜尋姓名／暱稱／帳號';
    const emptyText = opts.emptyText || '找不到符合的人員';
    const multi = !!opts.multiple;
    const allLabel = opts.allLabel || '全部業務';
    let selected = [];   // 多選：已勾選的 value（依序）
    const aliases = opts.aliases || {};   // { value: [別名...] }（名稱／暱稱）
    let uid = (enhance._n = (enhance._n || 0) + 1);

    const wrap = doc.createElement('div');
    wrap.className = 'ss-wrap fc-multisel-wrap';
    const btn = doc.createElement('button');
    btn.type = 'button';
    btn.className = 'ss-btn forecast-year-sel fc-multisel-btn';
    btn.setAttribute('aria-haspopup', 'listbox');
    btn.setAttribute('aria-expanded', 'false');
    const panel = doc.createElement('div');
    panel.className = 'ss-panel fc-multisel-panel';
    panel.hidden = true;
    const input = doc.createElement('input');
    input.type = 'text';
    input.className = 'ss-input';
    input.placeholder = placeholder;
    input.autocomplete = 'off';
    input.setAttribute('role', 'combobox');
    input.setAttribute('aria-autocomplete', 'list');
    input.setAttribute('aria-controls', 'ss-list-' + uid);
    input.setAttribute('aria-label', placeholder);
    const list = doc.createElement('ul');
    list.className = 'ss-list';
    list.id = 'ss-list-' + uid;
    list.setAttribute('role', 'listbox');
    const empty = doc.createElement('div');
    empty.className = 'ss-empty';
    empty.textContent = emptyText;
    empty.hidden = true;
    let btnAll = null, btnClear = null;
    if (multi) {
      list.setAttribute('aria-multiselectable', 'true');
      const tools = doc.createElement('div');
      tools.className = 'ss-tools';
      btnAll = doc.createElement('button'); btnAll.type = 'button'; btnAll.className = 'ss-tool ss-all';
      btnClear = doc.createElement('button'); btnClear.type = 'button'; btnClear.className = 'ss-tool ss-clear';
      tools.appendChild(btnAll); tools.appendChild(btnClear);
      panel.appendChild(input); panel.appendChild(tools);
    } else panel.appendChild(input);
    panel.appendChild(list); panel.appendChild(empty);
    wrap.appendChild(btn); wrap.appendChild(panel);
    select.parentNode.insertBefore(wrap, select.nextSibling);
    select.style.display = 'none';
    if (multi) select.multiple = true;
    select.setAttribute('aria-hidden', 'true');
    select.tabIndex = -1;

    let items = [], shown = [], active = -1;

    function readItems() {
      items = sortItems(Array.from(select.options).filter(o => !multi || o.value !== '').map(o => ({ value: o.value, label: o.textContent, keywords: aliases[o.value] || [] })));
      if (multi) selected = selected.filter(v => items.some(it => it.value === v));   // 選項被移除後不留幽靈勾選
    }
    function pushSelectState() {   // 原 select 的 selected 與集合同步（相容讀 selectedOptions 的程式）
      if (!multi) return;
      Array.from(select.options).forEach(o => { o.selected = o.value !== '' && selected.indexOf(o.value) !== -1; });
    }
    function labelOf(v) { const it = items.find(x => x.value === v); return it ? it.label : v; }
    function selectedLabel() {
      const o = select.options[select.selectedIndex];
      return o ? o.textContent : '';
    }
    function syncLabel() {
      if (multi) {
        btn.textContent = buttonText(selected.map(labelOf), allLabel) + ' ▾';
        btn.classList.toggle('active', selected.length > 0);
        btn.title = selected.length > 1 ? selected.map(labelOf).join('、') : '';
        return;
      }
      btn.textContent = selectedLabel() + ' ▾';
      btn.classList.toggle('active', select.value !== '');
    }
    function setActive(i, scroll) {
      active = i;
      Array.from(list.children).forEach((li, k) => {
        const on = k === i;
        li.classList.toggle('ss-active', on);
        if (on) { input.setAttribute('aria-activedescendant', li.id); if (scroll !== false && li.scrollIntoView) li.scrollIntoView({ block: 'nearest' }); }
      });
      if (i < 0) input.removeAttribute('aria-activedescendant');
    }
    function render() {
      shown = filterItems(items, input.value);
      while (list.firstChild) list.removeChild(list.firstChild);
      shown.forEach((it, k) => {
        const li = doc.createElement('li');
        li.className = 'ss-opt';
        li.id = 'ss-opt-' + uid + '-' + k;
        li.setAttribute('role', 'option');
        const isSel = multi ? selected.indexOf(it.value) !== -1 : it.value === select.value;
        li.setAttribute('aria-selected', isSel ? 'true' : 'false');
        li.dataset.value = it.value;
        if (multi) {
          const cb = doc.createElement('input');
          cb.type = 'checkbox'; cb.className = 'ss-chk'; cb.checked = isSel; cb.tabIndex = -1;
          cb.setAttribute('aria-hidden', 'true');
          const tx = doc.createElement('span'); tx.className = 'ss-txt'; tx.textContent = it.label;
          li.appendChild(cb); li.appendChild(tx);
        } else li.textContent = it.label;
        list.appendChild(li);
      });
      // 只剩「全部」一項且有輸入文字 → 顯示無結果
      empty.hidden = multi ? !(input.value.trim() && shown.length === 0) : !(input.value.trim() && shown.length <= 1);
      if (multi) {
        const q = input.value.trim();
        btnAll.textContent = q ? '全選目前結果' : '全選';
        btnClear.textContent = q ? '清除目前結果' : '清除';
        btnAll.disabled = btnClear.disabled = shown.length === 0;
        setActive(shown.length ? Math.min(Math.max(active, 0), shown.length - 1) : -1);
        return;
      }
      const cur = shown.findIndex(it => it.value === select.value);
      // 有輸入文字時預設反白第一個「真的符合」的人（跳過置頂的全部），方便直接 Enter
      let first = cur;
      if (input.value.trim()) first = shown.length > 1 ? 1 : 0;
      if (input.value.trim() && shown.length <= 1) setActive(-1);   // 沒有符合的人：不反白任何項目，Enter 不會誤選「全部」
      else setActive(first >= 0 ? first : 0);
    }
    function isOpen() { return !panel.hidden; }
    // 開啟時（以及視窗縮放／旋轉時）依按鈕位置與目前視窗寬度重算，不快取舊位置
    function position() {
      if (!isOpen()) return;
      const view = doc.defaultView;
      panel.style.left = '';
      const r = panel.getBoundingClientRect();
      const vw = doc.documentElement.clientWidth || view.innerWidth;
      const dx = clampShift(r.left, r.right, vw, 10);
      if (dx) panel.style.left = dx + 'px';
    }
    function open() {
      if (isOpen()) return;
      readItems(); input.value = ''; active = -1;
      panel.hidden = false;
      btn.setAttribute('aria-expanded', 'true');
      render();
      position();
      input.focus();
    }
    doc.defaultView.addEventListener('resize', position);
    function close(refocus) {
      if (!isOpen()) return;
      panel.hidden = true;
      btn.setAttribute('aria-expanded', 'false');
      if (refocus) btn.focus();
    }
    function emitChange() { pushSelectState(); syncLabel(); select.dispatchEvent(new doc.defaultView.Event('change', { bubbles: true })); }
    function repaintChecks() {   // 只更新勾選狀態，不重建清單（保留捲動位置與反白）
      Array.from(list.children).forEach((li, k) => {
        const on = !!shown[k] && selected.indexOf(shown[k].value) !== -1;
        li.setAttribute('aria-selected', on ? 'true' : 'false');
        const cb = li.querySelector('.ss-chk'); if (cb) cb.checked = on;
      });
    }
    function toggleAt(i) {
      const it = shown[i]; if (!it) return;
      selected = toggleValue(selected, it.value); repaintChecks(); emitChange();
    }
    function bulk(mode) {
      const before = selected.join('\u0000');
      selected = applyVisible(selected, shown.map(it => it.value), mode);
      repaintChecks();
      if (selected.join('\u0000') !== before) emitChange();
    }
    function pick(i) {
      if (multi) { toggleAt(i); return; }
      const it = shown[i];
      if (!it) return;
      const changed = select.value !== it.value;
      select.value = it.value;
      syncLabel();
      close(true);
      if (changed) select.dispatchEvent(new doc.defaultView.Event('change', { bubbles: true }));
    }

    btn.addEventListener('click', (e) => { e.stopPropagation(); isOpen() ? close(true) : open(); });
    btn.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); open(); }
    });
    input.addEventListener('input', render);
    input.addEventListener('keydown', (e) => {
      if (e.isComposing || e.keyCode === 229) return;   // 注音／拼音選字中，不攔截
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Home' || e.key === 'End') {
        if ((e.key === 'Home' || e.key === 'End') && input.value) return;   // 輸入框內的 Home/End 用來移動游標
        e.preventDefault(); setActive(moveIndex(active, e.key, shown.length));
      } else if (e.key === 'Enter') {
        e.preventDefault(); pick(active);   // 多選：勾選／取消，面板保持開啟
      } else if (multi && e.key === ' ' && !input.value) {
        e.preventDefault(); pick(active);   // 搜尋框為空時 Space 勾選；有字時照常輸入空白（多詞搜尋）
      } else if (e.key === 'Escape') {
        e.preventDefault(); e.stopPropagation(); close(true);
      } else if (e.key === 'Tab') {
        close(false);   // 不攔截，讓焦點照常離開
      }
    });
    list.addEventListener('click', (e) => {
      const li = e.target.closest ? e.target.closest('.ss-opt') : null;
      if (!li) return;
      const k = Array.from(list.children).indexOf(li);
      if (multi) setActive(k, false);
      pick(k);
    });
    if (multi) {
      btnAll.addEventListener('click', () => { bulk('all'); input.focus(); });
      btnClear.addEventListener('click', () => { bulk('clear'); input.focus(); });
    }
    list.addEventListener('mousemove', (e) => {
      const li = e.target.closest ? e.target.closest('.ss-opt') : null;
      if (li) { const k = Array.from(list.children).indexOf(li); if (k !== active) setActive(k, false); }
    });
    doc.addEventListener('click', (e) => { if (!wrap.contains(e.target)) close(false); }, true);

    readItems();
    if (multi) selected = Array.from(select.options).filter(o => o.selected && o.value !== '').map(o => o.value);
    syncLabel();
    return {
      wrap, button: btn, panel, input, list, select,
      getSelected() { return multi ? selected.slice() : (select.value ? [select.value] : []); },
      setSelected(vals) { if (!multi) return; selected = (vals || []).filter(v => items.some(it => it.value === v)); repaintChecks(); pushSelectState(); syncLabel(); },
      refresh() { select.style.display = 'none'; readItems(); syncLabel(); if (isOpen()) render(); },
      syncLabel, open, close,
      setVisible(v) { wrap.style.display = v ? '' : 'none'; },
    };
  }

  const api = { matchItem, filterItems, sortItems, moveIndex, toggleValue, applyVisible, buttonText, clampShift, enhance };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.ITTSSearchSelect = api;
})(typeof window !== 'undefined' ? window : globalThis);
