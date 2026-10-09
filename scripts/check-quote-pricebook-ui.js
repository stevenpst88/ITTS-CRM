#!/usr/bin/env node
/**
 * 報價牌價簿前端（_client/quote-pricebook.js 全域 QPB、_client/quote-costlines.js 的牌價簿區段、接線與跳脫）純函式／靜態檢查。
 * 用法：node scripts/check-quote-pricebook-ui.js（vm 沙盒載入，沒有 DOM、不需要伺服器）。對話框、後台表格、成本明細編輯器的 DOM 行為另以無頭瀏覽器實跑。
 *   1) QPB 純函式：parseQty、marginPct／marginText、sanitizeList、buildItems（只帶 desc／unit／qty／unitPrice／cat、依牌價簿順序）、
 *      isBlankDefaultRow、planAppend（取代單一空白預設列、50 列上限含分組標題／小計列、不改動輸入）
 *   2) QPB 載入：成功／HTTP 403／網路錯誤／壞 JSON 都不丟錯；60 秒快取、force 重抓、同時多次呼叫只送一個請求、失敗後仍保留上一份成功的資料
 *   3) QCL 牌價簿規則：pbClean、pbDescSuggest（不重複）、pbDefault（完全相符不分大小寫／去空白、成本空白或 0 才帶入、不覆蓋已填、委外廠商列不帶、非顧問區不帶、
 *      單位規則）、種子帶入（只在 pricebookSeed:true 且只帶顧問品項列的成本；沒開旗標時輸出與沒有牌價簿時逐位元相同）、載入已存的成本列（normalize）完全不動
 *   4) 靜態紀律：使用者文字進 innerHTML 前一律跳脫（QPB 對話框、後台表格）、無 eval／new Function／document.write、接線（按鈕、script 順序、?v= 版本）、
 *      既有簽核雜湊程式沒被改到、牌價簿命名空間
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const QPB_PATH = process.env.QPB_SRC || path.join(ROOT, '_client/quote-pricebook.js');       // 變異測試用：指向被破壞的副本
const QCL_PATH = process.env.QCL_SRC || path.join(ROOT, '_client/quote-costlines.js');

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const read = (p) => fs.readFileSync(p, 'utf8').replace(/\r\n/g, '\n');

const qpbSrc = read(QPB_PATH), qclSrc = read(QCL_PATH);
const ctx = { console };
vm.createContext(ctx);
vm.runInContext(qpbSrc, ctx);
vm.runInContext(qclSrc, ctx);
const P = ctx.QPB, Q = ctx.QCL;

async function run() {
  t('0. 載入只註冊 QPB／QCL（沒有 DOM 也不丟錯）', !!P && typeof P === 'object' && !!Q && typeof Q.pbDefault === 'function');

  // ═════════ 1) QPB 純函式 ═════════
  const okQ = [['1', 1], ['2.5', 2.5], [' 3 ', 3], ['100000', 100000], ['1.125', 1.125], [7, 7]];
  t('1.1 parseQty 通過：整數、小數（最多 3 位）、前後空白、上限 100000', okQ.every(([i, o]) => P.parseQty(i) === o), okQ.map(([i]) => P.parseQty(i)).join(','));
  const badQ = ['', ' ', '0', '0.5', '0.999', '-1', 'abc', '1e2', '1,000', '1.1234', '100001', null, undefined, NaN, Infinity, '.5', '5.', {}, true];
  t('1.2 parseQty 拒絕：空白、0、<1、負數、非數字、科學記號、逗號、超過 3 位小數、超過上限、NaN、Infinity', badQ.every((v) => P.parseQty(v) === null), badQ.filter((v) => P.parseQty(v) !== null).map(String).join('|'));
  t('1.3 marginPct：(牌價-成本)/牌價；牌價 0 或壞資料 → null；成本高於牌價 → 負值', P.marginPct(9000, 6000).toFixed(4) === '33.3333' && P.marginPct(0, 5) === null && P.marginPct('x', 1) === null && P.marginPct(100, 150) === -50 && P.marginPct(100, 0) === 100);
  t('1.4 marginText：一位小數＋%；無法計算顯示「—」', P.marginText(9000, 6000) === '33.3%' && P.marginText(0, 0) === '—' && P.marginText(100, 150) === '-50.0%');
  const raw = [{ id: 'a', name: ' PM ', price: 9000, cost: 6000, extra: 1 }, null, 'x', { id: 'b', name: '', price: 1, cost: 1 }, { id: 'c', name: 'N', price: -1, cost: 1 }, { id: 'd', name: 'N', price: 1, cost: NaN }, { name: 'noid', price: 1, cost: 1 }, { id: 'e', name: 'OK', price: '7000', cost: '5000' }];
  t('1.5 sanitizeList：丟掉壞項目（空名稱、負價、NaN、沒 id、非物件）；名稱去空白；只留 id／name／price／cost', J(P.sanitizeList(raw)) === J([{ id: 'a', name: 'PM', price: 9000, cost: 6000 }, { id: 'e', name: 'OK', price: 7000, cost: 5000 }]) && J(P.sanitizeList(undefined)) === '[]');
  const list = [{ id: 'a', name: 'PM 顧問經理', price: 9000, cost: 6500 }, { id: 'b', name: 'SD 顧問', price: 7000, cost: 5000 }, { id: 'c', name: 'ABAP', price: 6500, cost: 4500 }];
  const built = P.buildItems(list, [{ id: 'c', qty: 2 }, { id: 'a', qty: 3 }]);
  t('1.6 buildItems：依牌價簿順序（不是點選順序）；每項只有 desc／unit(人天)／qty／unitPrice／cat(consult)', J(built) === J([{ desc: 'PM 顧問經理', unit: '人天', qty: 3, unitPrice: 9000, cat: 'consult' }, { desc: 'ABAP', unit: '人天', qty: 2, unitPrice: 6500, cat: 'consult' }]));
  t('1.7 buildItems：沒有任何參照／出處欄位（不含 id、cost、pb*）；成本不會被帶進報價品項；找不到的 id 略過；空選取 → []', built.every((i) => !('id' in i) && !('cost' in i) && !Object.keys(i).some((k) => /^pb|prov|source/i.test(k))) && J(P.buildItems(list, [{ id: 'zz', qty: 1 }])) === '[]' && J(P.buildItems(list, [])) === '[]' && J(P.buildItems(list, null)) === '[]');
  const blank = { desc: '', unit: '式', qty: 1, unitPrice: 0, nid: 'n1', cat: '' };
  t('1.8 isBlankDefaultRow：新單預設列（含 readQuoteItems 補的 nid、cat 空字串）是；任何一欄動過、有 lid、標題／小計列、有成本、有分類都不是', P.isBlankDefaultRow(blank) && P.isBlankDefaultRow({ desc: '', unit: '式', qty: 1, unitPrice: 0 })
    && [{ desc: 'x' }, { unit: '人天' }, { qty: 2 }, { unitPrice: 1 }, { lid: 'L' }, { kind: 'title' }, { kind: 'subtotal' }, { cost: 5 }, { cat: 'consult' }, { needPrice: true }, { desc: ' x ' }].every((d) => !P.isBlankDefaultRow(Object.assign({}, blank, d))), '');
  t('1.9 isBlankDefaultRow：只有空白字元的說明仍算空白；壞輸入（null、字串）不是', P.isBlankDefaultRow(Object.assign({}, blank, { desc: '   ' })) && !P.isBlankDefaultRow(null) && !P.isBlankDefaultRow('x'));
  const add2 = built;
  let pl = P.planAppend([blank], add2, 50);
  t('1.10 planAppend：只有一列未動過的預設空白列 → 被取代（結果 2 列、replacedBlank）', pl.ok && pl.replacedBlank === true && pl.items.length === 2 && pl.items[0].desc === 'PM 顧問經理');
  const existing = [{ lid: 'L1', desc: '導入', unit: '式', qty: 1, unitPrice: 100 }];
  pl = P.planAppend(existing, add2, 50);
  t('1.11 planAppend：已有內容 → 接在後面，原列不動、replacedBlank=false', pl.ok && pl.replacedBlank === false && pl.items.length === 3 && pl.items[0] === existing[0] && pl.items[2].desc === 'ABAP');
  pl = P.planAppend([blank, Object.assign({}, blank)], add2, 50);
  t('1.12 planAppend：兩列空白列不取代（只有「單一」空白預設列才取代）', pl.ok && pl.items.length === 4 && pl.replacedBlank === false);
  pl = P.planAppend([Object.assign({}, blank, { desc: '已改過' })], add2, 50);
  t('1.13 planAppend：單一列但動過 → 不取代', pl.ok && pl.items.length === 3);
  const fortyEight = Array.from({ length: 48 }, (_, i) => ({ lid: 'x' + i, desc: 'r' + i, unit: '式', qty: 1, unitPrice: 1 }));
  t('1.14 上限：48 列＋2 項＝50 列通過；48 列＋3 項＝51 列拒絕（附 total／over／capacity），不產生半套結果', P.planAppend(fortyEight, add2, 50).ok && (() => { const r = P.planAppend(fortyEight, add2.concat([add2[0]]), 50); return !r.ok && r.total === 51 && r.over === 1 && r.capacity === 2 && !r.items; })());
  const withKinds = fortyEight.slice(0, 46).concat([{ lid: 't', kind: 'title', desc: 'Part A' }, { lid: 's', kind: 'subtotal', desc: '小計' }, { lid: 't2', kind: 'title', desc: 'B' }]);
  t('1.15 上限計入分組標題與小計列（49 列＋2 項 → 拒絕；49 列＋1 項通過）', !P.planAppend(withKinds, add2, 50).ok && P.planAppend(withKinds, [add2[0]], 50).ok);
  t('1.16 取代空白列後的上限：49 項新增到只有一列空白 → 49 列通過；51 項 → 拒絕（空白列被取代所以 capacity=50）', (() => { const many = Array.from({ length: 51 }, (_, i) => ({ desc: 'n' + i, unit: '人天', qty: 1, unitPrice: 1, cat: 'consult' })); const r = P.planAppend([blank], many, 50); return !r.ok && r.capacity === 50 && P.planAppend([blank], many.slice(0, 50), 50).ok; })());
  t('1.17 planAppend 不改動傳入陣列', (() => { const cur = [blank], snap = J(cur), addI = J(add2); P.planAppend(cur, add2, 50); return J(cur) === snap && J(add2) === addI; })());
  t('1.18 planAppend 沒給 maxRows 預設 50；空選取接在後面不報錯', !P.planAppend(fortyEight, add2.concat(add2), undefined).ok && P.planAppend(existing, [], 50).items.length === 1);

  // ═════════ 2) QPB 載入／快取 ═════════
  const mkFetch = (impl) => { const f = async (url, opts) => { f.calls.push([url, opts]); return impl(f.calls.length); }; f.calls = []; return f; };
  const okResp = (items) => ({ ok: true, status: 200, json: async () => ({ items }) });
  P._reset(); P._setFetch(mkFetch(() => okResp([{ id: 'a', name: 'PM', price: 9000, cost: 6000 }, { id: 'zz', name: '', price: 1, cost: 1 }])));
  let r1 = await P.load();
  t('2.1 成功：ok、項目經 sanitize、peek() 回同一份；請求打到 /api/quote-pricebook（同源帶 cookie）', r1.ok && r1.items.length === 1 && P.peek().length === 1 && ctx.__f === undefined);
  const f1 = mkFetch(() => okResp([])); P._setFetch(f1);
  let r2 = await P.load();
  t('2.2 60 秒內再 load 用快取（不再送請求）；force 重抓一次', r2.ok && r2.cached === true && f1.calls.length === 0 && (await P.load({ force: true })).ok && f1.calls.length === 1 && P.peek().length === 0);
  P._reset(); const slow = mkFetch(() => new Promise((rs) => setTimeout(() => rs(okResp([{ id: 'a', name: 'PM', price: 1, cost: 1 }])), 20))); P._setFetch(slow);
  const [pa, pb2] = await Promise.all([P.load(), P.load()]);
  t('2.3 同時多次 load 只送一個請求', slow.calls.length === 1 && pa.ok && pb2.ok);
  P._reset(); P._setFetch(mkFetch(() => ({ ok: false, status: 403, json: async () => ({ code: 'NO_PERMISSION' }) })));
  let r3 = await P.load();
  t('2.4 HTTP 403：回 { ok:false, status:403 }，不丟錯，peek() 仍是 null（成本明細編輯器等於沒有牌價簿）', !r3.ok && r3.status === 403 && P.peek() === null);
  P._reset(); P._setFetch(mkFetch(() => { throw new Error('網路中斷'); }));
  let r4 = await P.load();
  t('2.5 網路錯誤：{ ok:false, error } 不丟錯', !r4.ok && /網路中斷/.test(r4.error) && P.peek() === null);
  P._reset(); P._setFetch(mkFetch(() => ({ ok: true, status: 200, json: async () => { throw new SyntaxError('bad json'); } })));
  t('2.6 壞 JSON：{ ok:false } 不丟錯', !(await P.load()).ok);
  P._reset(); P._setFetch(mkFetch(() => ({ ok: true, status: 200, json: async () => ({}) })));
  const r7 = await P.load();
  t('2.7 回應沒有 items → 當空清單（ok、items=[]）', r7.ok && r7.items.length === 0);
  P._reset(); P._setFetch(mkFetch((n) => (n === 1 ? okResp([{ id: 'a', name: 'PM', price: 1, cost: 1 }]) : { ok: false, status: 500, json: async () => ({}) })));
  await P.load(); const bad = await P.load({ force: true });
  t('2.8 先成功後失敗：本次 ok:false，但 peek() 保留上一份成功的資料（編輯器不因暫時失敗而失去建議清單）', !bad.ok && P.peek() && P.peek().length === 1);
  P._reset(); P._setFetch(mkFetch(() => { throw new Error('x'); }));
  const [e1, e2] = await Promise.all([P.load(), P.load()]);
  t('2.9 失敗後 inflight 會清掉，下一次 load 可重試', !e1.ok && !e2.ok && (() => { const f = mkFetch(() => okResp([])); P._setFetch(f); return true; })() && (await P.load()).ok);
  P._reset(); P._setFetch(null);
  t('2.10 沒有 fetch 可用（沙盒）→ { ok:false } 不丟錯', !(await P.load()).ok);

  // ═════════ 3) QCL 牌價簿規則 ═════════
  const pb = Q.pbClean([{ name: 'PM 顧問經理', cost: 6500 }, { name: ' sd 顧問 ', cost: '5000' }, { name: 'Zero', cost: 0 }, { name: 'ABAP', cost: 4500 }, { name: 'abap', cost: 1 }, { name: '', cost: 5 }, { name: 'Neg', cost: -1 }, { name: 'NaN', cost: NaN }, null, 'x']);
  t('3.1 pbClean：去空白、不分大小寫去重（先到先贏）、丟掉空名稱／負成本／NaN／非物件；字串成本轉數字；成本 0 保留', J(pb) === J([{ name: 'PM 顧問經理', cost: 6500 }, { name: 'sd 顧問', cost: 5000 }, { name: 'Zero', cost: 0 }, { name: 'ABAP', cost: 4500 }]) && J(Q.pbClean(undefined)) === '[]' && J(Q.pbClean('x')) === '[]');
  const ds = Q.pbDescSuggest(pb);
  t('3.2 pbDescSuggest：牌價簿角色名＋既有 DESC_SUGGEST（PM、SD…），不分大小寫不重複；沒有牌價簿＝只有原本的清單（順序不變）', ds.indexOf('PM 顧問經理') === 0 && ds.includes('PM') && ds.includes('GUI/VAT') && new Set(ds.map((s) => s.toLowerCase())).size === ds.length && J(Q.pbDescSuggest([])) === J(['PM', 'SD', 'MM', 'PP', 'FI', 'CO', 'QM', 'BASIS', '客製', 'GUI/VAT', '電子發票']) && J(Q.pbDescSuggest(undefined)) === J(Q.pbDescSuggest([])));
  t('3.3 pbDescSuggest：牌價簿裡有和預設相同的名稱（pm／Sd）不會重複出現', (() => { const x = Q.pbDescSuggest(Q.pbClean([{ name: 'pm', cost: 1 }, { name: 'Sd', cost: 1 }])); return x.filter((s) => s.toLowerCase() === 'pm').length === 1 && x.filter((s) => s.toLowerCase() === 'sd').length === 1; })());
  const L = (o) => Object.assign({ cat: 'consult', desc: 'PM 顧問經理', vendor: '', unit: '式', unitCost: '' }, o || {});
  let d = Q.pbDefault(L(), pb);
  t('3.4 pbDefault 完全相符＋成本空白＋非委外 → 帶入牌價簿成本；單位預設「式」改「人天」', d && d.unitCost === 6500 && d.unit === '人天' && d.name === 'PM 顧問經理', J(d));
  t('3.5 不分大小寫、去頭尾空白也算相符；帶回的是牌價簿的標準名稱', (() => { const x = Q.pbDefault(L({ desc: '  pm 顧問經理 ' }), pb); const y = Q.pbDefault(L({ desc: 'SD 顧問' }), pb); return x && x.unitCost === 6500 && y && y.unitCost === 5000; })());
  t('3.6 不是完全相符（多一字、少一字、子字串、前綴）→ 不帶', ['PM 顧問', 'PM 顧問經理2', 'PM', 'xPM 顧問經理', 'PM  顧問經理', ''].every((s) => Q.pbDefault(L({ desc: s }), pb) === null));
  t('3.7 成本空白字串、空白字元、0、"0"、"0.00" 都算空 → 帶入', ['', '  ', 0, '0', '0.00', undefined, null].every((v) => { const x = Q.pbDefault(L({ unitCost: v }), pb); return x && x.unitCost === 6500; }));
  t('3.8 絕不覆蓋已填的值：成本 1、"1"、0.01、100000 → 不帶', [1, '1', 0.01, '100000', 100000].every((v) => Q.pbDefault(L({ unitCost: v }), pb) === null));
  t('3.9 輸入框裡是無效文字（"abc"、"1e"、"-"）→ 不蓋掉（交給原本的驗證擋）', ['abc', '12abc', '-'].every((v) => Q.pbDefault(L({ unitCost: v }), pb) === null));
  t('3.10 委外（委外廠商有填，含只有空白以外字元）→ 不帶；廠商只有空白視為沒有 → 帶', Q.pbDefault(L({ vendor: 'V' }), pb) === null && Q.pbDefault(L({ vendor: ' 廠商 ' }), pb) === null && Q.pbDefault(L({ vendor: '   ' }), pb) !== null);
  t('3.11 非顧問服務區（software／hw／travel／other）→ 不帶；自動列（印花稅）→ 不帶', ['software', 'hw', 'travel', 'other', '', undefined].every((c) => Q.pbDefault(L({ cat: c }), pb) === null) && Q.pbDefault(L({ auto: 'stamp' }), pb) === null);
  t('3.12 牌價簿成本為 0 的項目 → 不帶（沒有意義）；牌價簿空／壞 → 不帶；line 壞 → 不帶', Q.pbDefault(L({ desc: 'Zero' }), pb) === null && Q.pbDefault(L(), []) === null && Q.pbDefault(L(), null) === null && Q.pbDefault(null, pb) === null && Q.pbDefault(undefined, pb) === null);
  t('3.13 單位規則：空白或預設「式」→ 人天；自己填的（人月、台、人天）不動；keepUnit=true（連動列）一律維持原樣', Q.pbDefault(L({ unit: '' }), pb).unit === '人天' && Q.pbDefault(L({ unit: '人月' }), pb).unit === '人月' && Q.pbDefault(L({ unit: '台' }), pb).unit === '台' && Q.pbDefault(L({ unit: '人天' }), pb).unit === '人天' && Q.pbDefault(L({ unit: '式' }), pb, true).unit === '式' && Q.pbDefault(L({ unit: '' }), pb, true).unit === '');
  // 種子
  const items = [
    { lid: 'i1', desc: 'PM 顧問經理', unit: '人天', qty: 10, unitPrice: 9000, cat: 'consult' },
    { lid: 'i2', desc: 'sd 顧問', unit: '人天', qty: 5, unitPrice: 7000, cat: 'consult' },
    { lid: 'i3', desc: 'ABAP', unit: '人天', qty: 2, unitPrice: 6500, cat: 'consult', cost: 4000 },   // 舊式單的成本 4000 優先於牌價簿 4500
    { lid: 'i4', desc: '其他顧問', unit: '人天', qty: 1, unitPrice: 1, cat: 'consult' },
    { lid: 'i5', desc: 'PM 顧問經理', unit: '套', qty: 1, unitPrice: 1, cat: 'software' },          // 同名但是軟體分類 → 不帶
    { lid: 'i6', desc: 'Zero', unit: '人天', qty: 1, unitPrice: 1, cat: 'consult' },
  ];
  const seedPlain = Q.seedFromItems(items, { includeStamp: true });
  const seedNoFlag = Q.seedFromItems(items, { includeStamp: true, pricebook: pb });
  t('3.14 沒開 pricebookSeed → 種子與沒有牌價簿時逐位元相同（舊行為不變）', J(seedNoFlag) === J(seedPlain) && J(Q.seedFromItems(items, { includeStamp: true, pricebook: pb, pricebookSeed: false })) === J(seedPlain));
  const seedPb = Q.seedFromItems(items, { includeStamp: true, pricebook: pb, pricebookSeed: true });
  const byLid = (arr, lid) => arr.find((l) => l.forLid === lid);
  t('3.15 開 pricebookSeed → 相符的顧問品項列帶入成本（PM 6500、sd 顧問 5000）；單位維持品項的人天', byLid(seedPb, 'i1').unitCost === 6500 && byLid(seedPb, 'i2').unitCost === 5000 && byLid(seedPb, 'i1').unit === '人天');
  t('3.16 種子帶入不覆蓋舊式單已有的成本（ABAP 4000 不變）、不帶不相符（其他顧問）、不帶非顧問分類同名項（軟體）、不帶成本 0 的牌價項', byLid(seedPb, 'i3').unitCost === 4000 && !byLid(seedPb, 'i4').unitCost && !byLid(seedPb, 'i5').unitCost && !byLid(seedPb, 'i6').unitCost);
  t('3.17 種子帶入只改 unitCost：其餘欄位（說明、單位、數量、forLid、固定列）與沒帶牌價簿時完全相同', J(seedPb.map((l) => { const o = Object.assign({}, l); delete o.unitCost; return o; })) === J(seedPlain.map((l) => { const o = Object.assign({}, l); delete o.unitCost; return o; })) && seedPb.length === seedPlain.length);
  const seedLink = Q.seedFromItems(items, { mode: 'consultant', pricebook: pb, pricebookSeed: true });
  t('3.18 顧問對話框種子（連動）：成本帶入，rel 仍是 link、單位仍是品項的單位', byLid(seedLink, 'i1').unitCost === 6500 && byLid(seedLink, 'i1').rel === 'link' && byLid(seedLink, 'i1').unit === '人天');
  t('3.19 髒的牌價簿（undefined、字串、含壞項目）不影響種子', J(Q.seedFromItems(items, { pricebook: 'x', pricebookSeed: true })) === J(Q.seedFromItems(items, {})) && J(Q.seedFromItems(items, { pricebook: [null, { name: 5 }], pricebookSeed: true })) === J(Q.seedFromItems(items, {})));
  // 載入已存的成本列
  const saved = [{ cat: 'consult', desc: 'PM 顧問經理', qty: 3, unitCost: 0, unit: '式' }, { cat: 'consult', desc: 'SD 顧問', qty: 1, unitCost: 123, unit: '人天' }];
  t('3.20 載入已存的成本列（normalize／mount 的 lines 路徑）完全不動：沒有牌價簿參數，輸出與沒有牌價簿功能時一樣', J(Q.normalize(saved)) === J(Q.normalize(JSON.parse(J(saved)))) && Q.normalize(saved)[0].unitCost === 0 && Q.normalize(saved)[0].unit === '式' && Q.normalize(saved)[1].unitCost === 123);
  // 公開面
  t('3.21 QCL 匯出 pbClean／pbDescSuggest／pbDefault', typeof Q.pbClean === 'function' && typeof Q.pbDescSuggest === 'function' && typeof Q.pbDefault === 'function');

  // ═════════ 4) 靜態紀律 ═════════
  const adminSrc = read(path.join(ROOT, '_client/admin.html'));
  const a0 = adminSrc.indexOf('const PB_MAX_ITEMS'), a1 = adminSrc.indexOf('系統整合：連線設定', a0);
  const adminBlock = a0 > 0 && a1 > a0 ? adminSrc.slice(a0, a1) : '';
  t('4.1 後台牌價簿區段存在（pb* 函式）', adminBlock.length > 1000 && /function initPricebook/.test(adminBlock) && /function pbSave/.test(adminBlock));
  const nameUses = adminBlock.match(/[^\n]*\br\.name\b[^\n]*/g) || [];
  t('4.2 後台：r.name 只在「pcEsc(r.name)」「r.name.trim()」「r.name.…」或判斷式使用，沒有直接內插進 HTML', nameUses.length > 0 && nameUses.every((ln) => !/\$\{\s*r\.name\s*\}/.test(ln) && !/\+\s*r\.name\s*\+/.test(ln)) && /value="\$\{pcEsc\(r\.name\)\}"/.test(adminBlock), nameUses.filter((ln) => /\$\{\s*r\.name\s*\}|\+\s*r\.name\s*\+/.test(ln)).join('\n'));
  t('4.3 後台：price／cost／伺服器訊息／updatedBy／刪除確認的名稱，進 innerHTML 前都經 pcEsc', /value="\$\{pcEsc\(r\.price\)\}"/.test(adminBlock) && /value="\$\{pcEsc\(r\.cost\)\}"/.test(adminBlock) && /pcEsc\(d\.error \|\| /.test(adminBlock) && /pcEsc\(_pbLive\.updatedBy/.test(adminBlock) && /pcEsc\(e\.message\)/.test(adminBlock));
  const htmlAssign = (adminBlock.match(/innerHTML\s*=\s*[^;]+;/g) || []);
  t('4.4 後台：每個 innerHTML 指派裡的 ${…} 內插，要嘛是 pcEsc(…)／數字索引／固定字串，沒有未跳脫的使用者資料', htmlAssign.every((s) => !/\$\{\s*(?:r|it|d|e|_pbLive|_pbItems)\.(?:name|price|cost|error|message|updatedBy)\b/.test(s.replace(/pcEsc\([^)]*\)/g, ''))), htmlAssign.filter((s) => /\$\{\s*(?:r|it|d|e)\.(?:name|price|cost|error|message)\b/.test(s.replace(/pcEsc\([^)]*\)/g, ''))).join('\n'));
  t('4.5 後台：儲存走 PUT /admin/quote-pricebook 並帶載入時的 updatedAt；409 STALE_PRICEBOOK 有處理；有刪除確認、啟用核取、上下移', /method: 'PUT'/.test(adminBlock) && /updatedAt: _pbUpdatedAt/.test(adminBlock) && /STALE_PRICEBOOK/.test(adminBlock) && /confirm\(`刪除/.test(adminBlock) && /data-pbf="active"/.test(adminBlock) && /data-pbact="up"/.test(adminBlock) && /data-pbact="down"/.test(adminBlock));
  t('4.6 後台：側欄項 data-sec="quote-pricebook"、section id、init 分派、稽核動作中文名稱、委外說明文字都在', /data-sec="quote-pricebook"/.test(adminSrc) && /id="sec-quote-pricebook"/.test(adminSrc) && /dataset\.sec === 'quote-pricebook'\) initPricebook\(\)/.test(adminSrc) && /SAVE_QUOTE_PRICEBOOK: '/.test(adminSrc) && adminSrc.includes('委外費用依供應商報價，於填成本時輸入，不在此維護'));
  const dlgSrc = qpbSrc.slice(qpbSrc.indexOf('function tableHtml'), qpbSrc.indexOf('function openPicker'));
  t('4.7 QPB 對話框：項目名稱、id 進 HTML 前都經 esc（it.name／it.id 沒有直接相加）', /esc\(it\.name\)/.test(dlgSrc) && /esc\(it\.id\)/.test(dlgSrc) && !/\+\s*it\.name\s*\+/.test(dlgSrc) && !/\+\s*it\.id\s*\+/.test(dlgSrc) && /esc\(msg\)|esc\(res\.error\)|esc\(msg/.test(qpbSrc));
  t('4.8 QPB：沒有 eval／new Function／document.write；錯誤訊息用 textContent 顯示（setErr）', !/\beval\(|new Function|document\.write/.test(qpbSrc) && /el\.textContent = msg/.test(qpbSrc));
  t('4.9 QCL 新增的牌價簿區段沒有 eval／innerHTML 寫使用者文字（建議清單走 esc、提示用 textContent）', (() => { const seg = qclSrc.slice(qclSrc.indexOf('// ── 牌價簿 ──'), qclSrc.indexOf('global.QCL = {')); return seg.length > 500 && !/eval\(|new Function|document\.write|innerHTML/.test(seg) && /setMsg\(inst, '已依牌價簿帶入/.test(seg); })() && /esc\(x\)/.test(qclSrc.slice(qclSrc.indexOf('inst.setPricebook = function'), qclSrc.indexOf('inst.setPricebook = function') + 500)));
  const idx = read(path.join(ROOT, '_client/index.html')), srv = read(path.join(ROOT, 'server.js'));
  t('4.10 index.html：按鈕 #addQuotePbBtn 緊接在「新增項目」後；quote-pricebook.js 在 quote-costlines.js 與 quote.js 之前載入', /id="addQuoteItemBtn"[^\n]*\n\s*<button[^\n]*id="addQuotePbBtn"/.test(idx) && idx.indexOf('src="quote-pricebook.js"') > 0 && idx.indexOf('src="quote-pricebook.js"') < idx.indexOf('src="quote-costlines.js"') && idx.indexOf('src="quote-costlines.js"') < idx.indexOf('src="quote.js"'));
  t('4.11 server.js 對 quote-pricebook.js 注入 ?v= 版本（比照其他 quote*.js）', /\.replace\(\/src="quote-pricebook\\\.js"\/g,\s*`src="quote-pricebook\.js\?v=\$\{BUILD_VERSION\}"`\)/.test(srv));
  const quoteSrc = read(path.join(ROOT, '_client/quote.js'));
  const h0 = quoteSrc.indexOf("$('addQuotePbBtn').addEventListener"), h1 = quoteSrc.indexOf("$('addQuoteGroupBtn')", h0);
  const handler = h0 > 0 ? quoteSrc.slice(h0, h1) : '';
  t('4.12 quote.js 按鈕處理：QPB.openPicker → renderQuoteItems → updateQuoteTotals → QSteps.refresh；用 readQuoteItems 取得目前列、受 QUOTE_MAX_ROWS 限制；QPB 沒載入有提示', /QPB\.openPicker/.test(handler) && /getCurrent: readQuoteItems/.test(handler) && /maxRows: QUOTE_MAX_ROWS/.test(handler) && handler.indexOf('renderQuoteItems(items)') < handler.indexOf('updateQuoteTotals()') && handler.indexOf('updateQuoteTotals()') < handler.indexOf('QSteps.refresh') && /typeof QPB !== 'object'/.test(handler));
  t('4.13 quote.js：成本明細種子不帶牌價簿成本（pricebookSeed: false，避免未檢視就進簽核毛利）、掛載帶 pricebook；載入失敗（QPB 不存在或沒資料）時 pbList 是空陣列', /pricebook: pbList, pricebookSeed: false/.test(quoteSrc) && /function _qPbList\(\)[^\n]*\|\| \[\]/.test(quoteSrc));
  const qaSrc = read(path.join(ROOT, '_client/quote-approval.js'));
  t('4.14 quote-approval.js（顧問對話框）：種子與掛載都帶 pricebook 但 pricebookSeed: false；載入後 setPricebook；沒有 QPB 時退回空陣列', /pricebook: pbList, pricebookSeed: false/.test(qaSrc) && /setPricebook\(r\.items\)/.test(qaSrc) && /typeof QPB === 'object' && QPB && QPB\.peek\(\)\) \|\| \[\]/.test(qaSrc));
  // 既有簽核／雜湊／正規化程式沒被這個功能改到
  let gitOk = true, changed = '';
  try { changed = require('child_process').execFileSync('git', ['diff', '--name-only', 'HEAD', '--', 'lib/quoteApproval.js', 'lib/quoteItems.js', 'lib/quoteCostLines.js'], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim(); } catch (e) { gitOk = false; }
  if (gitOk) t('4.15 簽核雜湊／品項列／成本明細的伺服器程式（lib/quoteApproval.js、quoteItems.js、quoteCostLines.js）相對 HEAD 沒有改動', changed === '', changed);
  else console.log('SKIP 4.15（不是 git 工作樹，無法比對 HEAD）——這一項不計入');
  const routesSrc = read(path.join(ROOT, 'lib/quoteRoutes.js'));
  t('4.16 normalizeItems 沒有被改成認得牌價簿欄位（沒有 pbId／pricebook 字樣出現在 normalizeItems 區段）', (() => { const a = routesSrc.indexOf('function normalizeItems('), b = routesSrc.indexOf('畫面還沒存檔的新品項沒有 lid'); return a > 0 && b > a && !/pricebook|pbId|pb[A-Z]/i.test(routesSrc.slice(a, b)); })());
}

run().then(() => {
  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
  console.log(`\n報價牌價簿前端檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
}).catch((e) => { res.filter((r) => !r[1]).forEach((r) => console.log('FAIL ' + r[0] + (r[2] ? '  <- ' + r[2] : ''))); console.error('例外（後面的檢查沒跑到）', e.stack); process.exit(2); });
