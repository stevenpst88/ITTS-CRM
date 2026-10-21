#!/usr/bin/env node
/**
 * 客戶清單空狀態檢查：有輸入搜尋或套用 Key Account 篩選而沒有結果時，
 * 不得顯示「尚無成交客戶／尚無潛在客戶」（會誤導成資料不見了），而是顯示「找不到符合…」並提供「清除搜尋」。
 * 純 vm 檢查（不起伺服器、不寫檔）：從 _client/app.js 切出 renderContacts 區段，以最小 DOM 樁執行。
 *   node scripts/check-contacts-empty-search.js
 */
const fs = require('fs'), vm = require('vm'), path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', '_client', 'app.js'), 'utf8');
const a = src.indexOf('function renderContacts(contacts)');
const b = src.indexOf('function fillContactList');
if (a < 0 || b < a) { console.error('找不到 renderContacts 區段'); process.exit(2); }
const code = src.slice(a, b);

let pass = 0, fail = 0;
function ok(c, m) { if (c) pass++; else { fail++; console.log('  ✗ FAIL:', m); } }

function el(tag) {
  return { tag, children: [], className: '', _l: {}, type: '', _t: '', _h: '',
    set textContent(v) { this._t = v; }, get textContent() { return this._t; },
    set innerHTML(v) { this._h = v; this.children = []; }, get innerHTML() { return this._h; },
    appendChild(c) { this.children.push(c); return c; },
    addEventListener(n, f) { this._l[n] = f; } };
}
function run(opts) {
  const lists = {}; let cleared = 0;
  const inputs = { searchInput: { value: opts.term || '' }, clearSearch: { click() { cleared++; } } };
  const ctx = {
    document: { createElement: el },
    $: id => inputs[id] || lists[id] || (lists[id] = el('div')),
    currentSection: opts.section || 'contacts', contactsKaFilterMode: opts.ka || 'all', isSearchMode: !!opts.search,
    isKeyAccount: () => false, allGroups: [], fillContactList() {}, _kaBatchStart() { return false; }, _kaBatchEnd() {},
  };
  vm.createContext(ctx); vm.runInContext(code, ctx);
  ctx.renderContacts([]);
  const list = lists[opts.section === 'prospects' ? 'prospectList' : 'contactList'];
  return { list, cleared: () => cleared };
}

// 1) 搜尋無結果（客戶）
let r = run({ search: true, term: '將來' });
let box = r.list.children[0];
ok(box && box.children[1].textContent === '找不到符合「將來」的客戶', '搜尋無結果：顯示「找不到符合…的客戶」');
ok(!/尚無成交客戶/.test(r.list.innerHTML) && !/尚無成交客戶/.test(JSON.stringify(r.list.children.map(c => c.children.map(x => x.textContent)))), '搜尋無結果：不顯示「尚無成交客戶」');
const btn = box && box.children[3];
ok(btn && btn.textContent === '清除搜尋' && btn.type === 'button', '搜尋無結果：有「清除搜尋」按鈕');
if (btn) { btn._l.click(); ok(r.cleared() === 1, '按「清除搜尋」會觸發既有的清除搜尋'); }
// 2) 搜尋無結果（潛在客戶）
r = run({ search: true, term: 'abc', section: 'prospects' });
ok(r.list.children[0].children[1].textContent === '找不到符合「abc」的潛在客戶', '潛在客戶頁：顯示「找不到符合…的潛在客戶」');
// 3) Key Account 篩選無結果
r = run({ ka: 'ka' });
ok(r.list.children[0].children[1].textContent === '此篩選條件下沒有客戶', 'KA 篩選無結果：顯示「此篩選條件下沒有客戶」');
// 4) 真的沒有資料：維持原訊息
r = run({});
ok(/尚無成交客戶/.test(r.list.innerHTML), '沒搜尋也沒篩選：維持「尚無成交客戶」');
r = run({ section: 'prospects' });
ok(/尚無潛在客戶/.test(r.list.innerHTML), '潛在客戶頁沒搜尋：維持「尚無潛在客戶」');
// 5) XSS：關鍵字只以 textContent 放入
r = run({ search: true, term: '<img src=x onerror=alert(1)>' });
ok(r.list.children[0].children[1].textContent.includes('<img') && !r.list.children[0].innerHTML, 'XSS：關鍵字以純文字呈現（未使用 innerHTML）');
// 6) 靜態：該區段不得用 innerHTML 組入搜尋字
ok(!/innerHTML\s*=\s*[^;]*term/.test(code), '靜態：搜尋字不經 innerHTML');

console.log(`客戶清單空狀態檢查結果：${pass} 通過，${fail} 失敗`);
process.exit(fail ? 1 : 0);
