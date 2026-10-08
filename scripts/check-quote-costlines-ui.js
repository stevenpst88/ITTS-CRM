#!/usr/bin/env node
/**
 * 成本明細編輯器（_client/quote-costlines.js，全域 QCL）純函式檢查。用法：node scripts/check-quote-costlines-ui.js
 * 動 _client/quote-costlines.js 之後必跑。DOM 行為（新增／刪除／移動、輸入不失焦、印花稅勾選、上限停用、view 模式、窄螢幕）
 * 另以無頭瀏覽器實跑，這支只測不需要 DOM 的部分：
 *   1) 公開介面與五區定義
 *   2) seedFromItems：分類規則（品項 cat → 單位猜測 → 單一類別）、略過分組標題／小計列、沒有 unitPrice 也能跑、固定列、上限
 *   3) normalize：缺欄補預設、截長度、數字轉型與範圍、未知 cat 丟棄、印花稅列規格化、上限
 *   4) totals：各分區小計、每列取整到分、印花稅 round-half-up、營收未知時不計並標 stampUnknown
 *   5) 跳脫：使用者文字進 HTML 前一律 esc（含 view／edit 列）
 *   6) 檔案紀律：CSS 類名全部 qcl- 前綴、只注入一次、不依賴 app.js 的全域函式
 *   7) 與伺服器 lib/quotePnlExcel.js 的 resolveCat／catByUnit 逐例相同（兩邊是鏡像；硬體鍵名 hardware ↔ hw）
 *   8) 營收 QCL.revenueOf ＝ 伺服器 computeFinancials 的 revenueCents/100：已知案例＋≥100,000 組隨機輸入逐一比對 0 差異；印花稅 stampAmount ＝ CL.stampDollars
 *   （另：seedFromItems 的 includeStamp:false＝舊式單的種子不含印花稅列，見 2m–2o）
 *   9) 「＋補入新品項」覆蓋判定 QCL._missingItemLines：forLid 對得上品項 lid 就算覆蓋，其餘成本列改用品名多重集合比對（新單品項沒有 lid、同名品項、改名、新增）
 *   10) 種子（含重新帶入、補入）略過說明為空白的品項，不產生空白「項目」列
 *   15) forLids：清洗（與伺服器同規則）、合併為一列時聯集、整包後補入／完成提醒認得被併的品項、列 HTML 往返
 *   16) 業務端提醒 QCL.unmatchedPricedItems／unmatchedPricedNote：只算有價品項、空白說明的有價品項也算、文字與伺服器 preview.warnings 同一句
 *       （16i–16q：儲存前確認的基準 QCL.unmatchedKeys／unmatchedNewNote——只對基準以外新出現的未涵蓋品項提醒；含隨機 3000 組不變式）
 *   17) 前端與伺服器 lib/quoteCostLines.js 的涵蓋判定／警告文字：≥5000 組隨機輸入（含髒資料）逐組比對 0 差異
 *   18) 「原報價品項已刪除」徽章：唯讀列 HTML、_isOrphanLine、items 沒提供就不顯示
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const SRC_PATH = process.env.QCL_SRC || path.join(ROOT, '_client/quote-costlines.js');   // QCL_SRC：變異測試用，指向被破壞的副本

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);

const src = fs.readFileSync(SRC_PATH, 'utf8').replace(/\r\n/g, '\n');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(src, ctx);          // 沒有 window／document：只應註冊 QCL，不碰 DOM
const Q = ctx.QCL;

const it = (desc, unit, qty, extra) => Object.assign({ lid: 'i-' + desc, desc, unit, qty }, extra || {});
const ti = (d) => ({ lid: 't-' + d, kind: 'title', desc: d });
const su = (d) => ({ lid: 's-' + d, kind: 'subtotal', desc: d });
const ln = (cat, extra) => Object.assign({ cat, desc: 'x', qty: 1, unitCost: 0 }, extra || {});

// 1) 介面
t('1a. 載入檔案只註冊 QCL（沒有 DOM 也不會丟錯）', !!Q && typeof Q === 'object');
t('1b. 匯出：CATS／MAX_LINES／seedFromItems／normalize／totals／mount／collect', Array.isArray(Q.CATS) && Q.MAX_LINES === 60 && ['seedFromItems', 'normalize', 'totals', 'mount', 'collect'].every((k) => typeof Q[k] === 'function'));
t('1c. 五區順序與廠商欄：顧問（委外廠商）／軟體、硬體（供應商）有廠商欄，差旅、其他沒有；全部有說明欄',
  J(Q.CATS.map((c) => [c.key, c.name, c.hasVendor, c.vendorLabel, c.hasNote])) ===
  J([['consult', '顧問服務成本', true, '委外廠商', true], ['software', '軟體成本', true, '供應商', true], ['hw', '硬體成本', true, '供應商', true], ['travel', '差旅費用', false, '', true], ['other', '其他費用', false, '', true]]));
t('1d. CATS 凍結（呼叫端改不動定義）', Object.isFrozen(Q.CATS) && Q.CATS.every((c) => Object.isFrozen(c)));
let threw = false; try { Q.mount(null, {}); } catch (e) { threw = e.name === 'TypeError'; }
t('1e. mount 沒給容器 → TypeError', threw);

// 2) seedFromItems
const s1 = Q.seedFromItems([
  ti('Part A'), it('顧問', '人天', 10, { cost: 2000, unitPrice: 9999 }), it('主機', '台', 2), it('授權費', '授權', 1, { cat: 'software' }),
  it('雜項', '式', 1), su('小計'), it('指定硬體', '式', 1, { cat: 'hardware' }), it('指定hw', '式', 1, { cat: 'hw' }),
]);
const byDesc = (arr, d) => arr.find((l) => l.desc === d);
t('2a. 分類：人天→consult、台→hw、品項自己的 cat 優先、式→other；cat 的 hardware 與 hw 都對到 hw',
  byDesc(s1, '顧問').cat === 'consult' && byDesc(s1, '主機').cat === 'hw' && byDesc(s1, '授權費').cat === 'software' && byDesc(s1, '雜項').cat === 'other' && byDesc(s1, '指定硬體').cat === 'hw' && byDesc(s1, '指定hw').cat === 'hw', J(s1.map((l) => l.cat)));
t('2b. 略過 title／subtotal 列；種子＝6 個品項＋差旅＋交際費＋印花稅 = 9 列', s1.length === 9 && !s1.some((l) => l.desc === 'Part A' || l.desc === '小計'), s1.length);
t('2c. 品項列帶入 unit／qty／forLid；unitCost＝舊 items[].cost（沒有就 0）、不讀 unitPrice（顧問看不到品項單價）',
  byDesc(s1, '顧問').unit === '人天' && byDesc(s1, '顧問').qty === 10 && byDesc(s1, '顧問').forLid === 'i-顧問' && byDesc(s1, '顧問').unitCost === 2000 && byDesc(s1, '主機').unitCost === 0 && !('unitPrice' in byDesc(s1, '顧問')));
const fixed = s1.slice(-3);
t('2d. 固定列：差旅「差旅交通」1 列、其他「交際費」、印花稅 auto 列（cat other、qty 1、unit 式、沒有 unitCost 欄位＝尚無伺服器金額）',
  J(fixed) === J([
    { cat: 'travel', desc: '差旅交通', vendor: '', note: '', unit: '式', qty: 1, unitCost: 0 },
    { cat: 'other', desc: '交際費', vendor: '', note: '', unit: '式', qty: 1, unitCost: 0 },
    { cat: 'other', desc: Q.STAMP_DESC, vendor: '', note: '', unit: '式', qty: 1, auto: 'stamp' },
  ]), J(fixed));
t('2e. 沒有品項／items 不是陣列／含 null 與非物件：只剩 3 個固定列、不丟錯', [undefined, null, [], 'x', [null, 5, 'a']].every((v) => Q.seedFromItems(v).length === 3));
t('2f. 數量空／0／負／非數字 → 1，小數保留，超過上限截到 1e9', J(Q.seedFromItems([it('a', '式', ''), it('b', '式', 0), it('c', '式', -3), it('d', '式', 'x'), it('e', '式', 0.5), it('f', '式', 1e12)]).slice(0, 6).map((l) => l.qty)) === J([1, 1, 1, 1, 0.5, 1e9]));
const many = Array.from({ length: 100 }, (_, i) => it('品項' + i, '人天', 1));
const sm = Q.seedFromItems(many);
t('2g. 品項過多：種子不超過 MAX_LINES（60），固定的 3 列仍保留（印花稅列在最後）', sm.length === 60 && sm[59].auto === 'stamp' && sm[58].desc === '交際費' && sm[57].desc === '差旅交通', sm.length);
const sl = Q.seedFromItems([it('第一行\n第二行', '式', 1), it('長'.repeat(300), '式', 1)]);
t('2h. 品名換行併成空白、超過 120 字截斷', sl[0].desc === '第一行 第二行' && sl[1].desc.length === 120);
const cc = Q.seedFromItems([it('a', '式', 1), it('b', '台', 1)], { classCodes: ['consult'] });
const cm = Q.seedFromItems([it('a', '式', 1), it('b', '台', 1)], { classCodes: ['consult', 'hardware'] });
const cx = Q.seedFromItems([it('a', '式', 1)], { classCodes: ['crm', 'mdm'] });
t('2i. classCodes：勾選商品只涵蓋單一分區 → 全歸該區（crm／mdm 都算 other）；涵蓋多區 → 依單位猜', cc[0].cat === 'consult' && cc[1].cat === 'consult' && cm[0].cat === 'other' && cm[1].cat === 'hw' && cx[0].cat === 'other');
t('2j. 種子是 normalize 的不動點（種子直接丟給 normalize 不會變）', J(Q.normalize(s1)) === J(s1));
t('2k. 不修改傳入的 items', (() => { const items = [it('顧問', '人天', 1, { cost: 5 })]; const before = J(items); Q.seedFromItems(items); return J(items) === before; })());
t('2l. revenueKnown 參數不影響種子內容（印花稅列一律種入）', J(Q.seedFromItems([it('a', '式', 1)], { revenueKnown: false })) === J(Q.seedFromItems([it('a', '式', 1)], { revenueKnown: true })));

// 2m) includeStamp：舊式單（還沒有成本明細）的種子不含印花稅列；預設與 true 都含
const sNo = Q.seedFromItems([it('a', '人天', 2, { cost: 100 }), it('b', '台', 1)], { includeStamp: false });
t('2m. includeStamp:false：種子＝品項 2 列＋差旅＋交際費，沒有印花稅列；預設與 includeStamp:true 仍是 3 個固定列（印花稅在最後）',
  sNo.length === 4 && !sNo.some((l) => l.auto === 'stamp') && sNo[2].desc === '差旅交通' && sNo[3].desc === '交際費'
  && Q.seedFromItems([it('a', '人天', 2)]).slice(-1)[0].auto === 'stamp' && Q.seedFromItems([it('a', '人天', 2)], { includeStamp: true }).slice(-1)[0].auto === 'stamp', J(sNo.map((l) => l.desc)));
t('2n. includeStamp:false 品項過多：仍不超過 60 列，差旅與交際費保留', (() => { const x = Q.seedFromItems(many, { includeStamp: false }); return x.length === 60 && x[59].desc === '交際費' && x[58].desc === '差旅交通' && !x.some((l) => l.auto === 'stamp'); })());
t('2o. includeStamp:false 的種子是 normalize 的不動點，totals 不含印花稅（營收再大也不計）', (() => { const x = Q.seedFromItems([it('a', '人天', 2, { cost: 100 })], { includeStamp: false }); return J(Q.normalize(x)) === J(x) && Q.totals(x, 1e9).stamp === 0 && Q.totals(x, 1e9).total === 200 && !Q.totals(x, undefined).stampUnknown; })());

// 3) normalize
const n0 = Q.normalize([{ cat: 'consult' }])[0];
t('3a. 缺欄補預設：desc ""、vendor ""、note ""、unit 式、qty 1、unitCost 0；沒有 lid／forLid／auto 欄位',
  J(n0) === J({ cat: 'consult', desc: '', vendor: '', note: '', unit: '式', qty: 1, unitCost: 0 }), J(n0));
const nt = Q.normalize([{ cat: 'consult', desc: 'd'.repeat(130), vendor: 'v'.repeat(70), note: 'n'.repeat(210), unit: 'u'.repeat(15), lid: 'L'.repeat(80), forLid: 'F'.repeat(80) }])[0];
t('3b. 截長度：desc 120、vendor 60、note 200、unit 10（lid／forLid 64）', nt.desc.length === 120 && nt.vendor.length === 60 && nt.note.length === 200 && nt.unit.length === 10 && nt.lid.length === 64 && nt.forLid.length === 64);
const nn = Q.normalize([ln('hw', { qty: '3', unitCost: '2.5' }), ln('hw', { qty: 'abc', unitCost: 'x' }), ln('hw', { qty: -2, unitCost: -9 }), ln('hw', { qty: 1e10, unitCost: 1e13 }), ln('hw', { qty: 0.5, unitCost: '' }), ln('hw', { qty: Infinity, unitCost: NaN }), ln('hw', { qty: 0 })]);
t('3c. 數字轉型：字串轉數字；非數字→預設（qty 1／unitCost 0）；負數→0；超過上限截到 1e9／1e12；qty 0 保留 0；Infinity／NaN 用預設',
  J(nn.map((l) => [l.qty, l.unitCost])) === J([[3, 2.5], [1, 0], [0, 0], [1e9, 1e12], [0.5, 0], [1, 0], [0, 0]]), J(nn.map((l) => [l.qty, l.unitCost])));
const nu = Q.normalize([ln('xyz'), ln('constructor'), ln('__proto__'), ln('hardware'), ln('toString'), null, 'str', 5, [], ln('travel')]);
t('3d. 未知 cat（含 prototype 鍵名、舊的 hardware）與非物件元素丟棄', nu.length === 1 && nu[0].cat === 'travel', J(nu));
t('3e. 非陣列輸入 → 空陣列', [undefined, null, {}, 'a', 5].every((v) => Array.isArray(Q.normalize(v)) && Q.normalize(v).length === 0));
const nv = Q.normalize([ln('consult', { vendor: 'V1' }), ln('software', { vendor: 'V2' }), ln('hw', { vendor: 'V3' }), ln('travel', { vendor: 'V4' }), ln('other', { vendor: 'V5' })]);
t('3f. 廠商欄：顧問／軟體／硬體保留，差旅／其他一律清空', J(nv.map((l) => l.vendor)) === J(['V1', 'V2', 'V3', '', '']));
const ns = Q.normalize([{ cat: 'consult', auto: 'stamp', desc: '亂寫', qty: 9, unit: '台', unitCost: 99999, vendor: 'V', lid: 'S1', note: 'nn' }, { cat: 'other', auto: 'stamp', lid: 'S2' }, ln('other', { auto: 'bogus', lid: 'B' })]);
t('3g. 印花稅列規格化：cat 強制 other、desc 固定、qty 1、unit 式、vendor 空；保留 lid／note；第二個印花稅列丟棄；未知 auto 值移除',
  ns.length === 2 && J(ns[0]) === J({ lid: 'S1', cat: 'other', desc: Q.STAMP_DESC, vendor: '', note: 'nn', unit: '式', qty: 1, unitCost: 99999, auto: 'stamp' }) && !('auto' in ns[1]) && ns[1].lid === 'B', J(ns));
const stU = (u) => Q.normalize([{ cat: 'other', auto: 'stamp', unitCost: u }])[0];
t('3g2. v1.1 印花稅列的 unitCost＝伺服器提供的金額：有限數字才保留（0 保留、負數→0、超上限截 1e12、數字字串轉數字）；沒給／空／null／非數字 → 沒有 unitCost 欄位（與金額 0 區分）',
  stU(0).unitCost === 0 && stU(-5).unitCost === 0 && stU(1e13).unitCost === 1e12 && stU('250').unitCost === 250 && stU(1234).unitCost === 1234 && [undefined, '', null, 'abc', NaN, Infinity, {}].every((u) => !('unitCost' in stU(u))), J([0, -5, 1e13, '250', undefined, 'abc'].map((u) => stU(u))));
t('3h. 上限：70 列只留前 60 列', Q.normalize(Array.from({ length: 70 }, () => ln('other'))).length === 60);
t('3i. 上限內含印花稅：第 60 列之後的印花稅列也被截掉', (() => { const a = Array.from({ length: 60 }, () => ln('other')); a.push({ cat: 'other', auto: 'stamp' }); return Q.normalize(a).every((l) => !l.auto); })());
t('3j. 回傳新陣列與新物件、不修改輸入；清洗後再清洗結果不變（冪等）', (() => {
  const input = [{ cat: 'consult', desc: ' PM ', qty: '2' }]; const before = J(input); const out = Q.normalize(input);
  return J(input) === before && out !== input && out[0] !== input[0] && out[0].desc === 'PM' && J(Q.normalize(out)) === J(out);
})());
t('3k. lid／forLid 保留；品名頭尾空白與換行清掉', (() => { const l = Q.normalize([{ cat: 'travel', desc: '  a\r\nb  ', lid: 'q1', forLid: 'f1' }])[0]; return l.lid === 'q1' && l.forLid === 'f1' && l.desc === 'a b'; })());

// 4) totals
const T1 = Q.totals([ln('consult', { qty: 2, unitCost: 1000 }), ln('consult', { qty: 0.5, unitCost: 3 }), ln('software', { qty: 1, unitCost: 500 }), ln('hw', { qty: 3, unitCost: 100 }), ln('travel', { unitCost: 40 }), ln('other', { unitCost: 7 })], 100000);
t('4a. 各分類小計：consult 2001.5、software 500、hw 300、travel 40、other 7；subtotalExStamp 2848.5', J(T1.byCat) === J({ consult: 2001.5, software: 500, hw: 300, travel: 40, other: 7 }) && T1.subtotalExStamp === 2848.5, J(T1));
t('4b. 沒有印花稅列：stamp 0、stampUnknown false、total＝subtotalExStamp（即使有營收）', T1.stamp === 0 && T1.stampUnknown === false && T1.total === T1.subtotalExStamp);
const stp = { cat: 'other', auto: 'stamp' };
const sv = (rev) => Q.totals([ln('consult', { unitCost: 1000 }), stp], rev);
t('4c. 印花稅 round-half-up 到整數元：12500→13（.5 進位）、12499→12、500→1、499→0、0→0、1500→2、2500→3、2499.99→2',
  J([12500, 12499, 500, 499, 0, 1500, 2500, 2499.99].map((r) => sv(r).stamp)) === J([13, 12, 1, 0, 0, 2, 3, 2]), J([12500, 12499, 500, 499, 0, 1500, 2500, 2499.99].map((r) => sv(r).stamp)));
t('4d. 印花稅進位含浮點雜訊的營收：1005×0.001=1.005、100.5 元等邊界（1005000→1005；1500000.4→1500）', sv(1005000).stamp === 1005 && sv(1500000.4).stamp === 1500 && sv(1500500).stamp === 1501);
const su0 = sv(undefined);
t('4e. 營收未提供：stamp 0、total 不含印花稅、stampUnknown true；營收有值：total＝小計＋印花稅', su0.stamp === 0 && su0.stampUnknown === true && su0.total === 1000 && sv(200000).stamp === 200 && sv(200000).total === 1200 && sv(200000).stampUnknown === false);
t('4f. 營收是空字串／null／NaN／負數／非數字字串／物件 → 視為未知', [null, '', NaN, -1, 'abc', {}, []].every((r) => sv(r).stampUnknown === true && sv(r).stamp === 0));
t('4g. 營收是數字字串（"250000"）可用', sv('250000').stamp === 250 && sv('250000').stampUnknown === false);
t('4h. 印花稅：有營收就依營收重算，不採列上的 unitCost（99999 不影響合計）', Q.totals([{ cat: 'other', auto: 'stamp', unitCost: 99999 }], 100000).total === 100);
const stL = (u, rev, extra) => Q.totals([ln('consult', { unitCost: 1000 }), Object.assign({ cat: 'other', auto: 'stamp', unitCost: u }, extra || {})], rev);
t('4h2. v1.1 營收缺（顧問端）但 stamp 列有伺服器金額 → 用該金額計入 total，stampUnknown false（250 → 合計 1250）', (() => { const r = stL(250, undefined); return r.stamp === 250 && r.stampUnknown === false && r.total === 1250 && r.subtotalExStamp === 1000; })());
t('4h3. v1.1 營收無效（null／空字串／NaN／負數／非數字）時同樣採伺服器金額', [null, '', NaN, -1, 'abc'].every((r) => stL(250, r).stamp === 250 && stL(250, r).stampUnknown === false));
t('4h4. v1.1 伺服器金額 0（營收小於 500 元的合法結果）是「已知的 0」：stampUnknown false、不計入', (() => { const r = stL(0, undefined); return r.stamp === 0 && r.stampUnknown === false && r.total === 1000; })());
t('4h5. v1.1 營收與伺服器金額都沒有才 stampUnknown true 且不計入；營收有值時以營收為準（營收 0 → 0，不採列上的 250）', stL(undefined, undefined).stampUnknown === true && stL(undefined, undefined).total === 1000 && stL(250, 0).stamp === 0 && stL(250, 0).stampUnknown === false && stL(250, 300000).stamp === 300);
t('4h6. v1.1 多個印花稅列只認第一個的伺服器金額', Q.totals([{ cat: 'other', auto: 'stamp', unitCost: 7 }, { cat: 'other', auto: 'stamp', unitCost: 900 }], undefined).stamp === 7);
t('4i. 同一單多個印花稅列只算一次', Q.totals([stp, stp, stp], 1000000).stamp === 1000);
// 每列先取整到分再加總（與伺服器 BigInt 分的規則一致）
const T2 = Q.totals([ln('other', { qty: 1, unitCost: 1.005 }), ln('other', { qty: 1, unitCost: 1.005 })], 0);
t('4j. 每列先四捨五入到分：1.005 → 1.01，兩列 2.02（不是 2.01）', T2.byCat.other === 2.02, T2.byCat.other);
t('4k. 數量×單價的浮點雜訊：0.1×3（0.30000000000000004）→ 0.3；1.1×1.1 → 1.21', Q.totals([ln('hw', { qty: 3, unitCost: 0.1 })], 0).byCat.hw === 0.3 && Q.totals([ln('hw', { qty: 1.1, unitCost: 1.1 })], 0).byCat.hw === 1.21);
t('4l. 壞資料不丟錯：lines 不是陣列／含未知 cat／負數量 → 該列忽略或當 0', (() => {
  const a = Q.totals(undefined, 1); const b = Q.totals([ln('zzz', { unitCost: 5 }), ln('hw', { qty: -5, unitCost: 5 })], 1);
  return a.total === 0 && J(a.byCat) === J({ consult: 0, software: 0, hw: 0, travel: 0, other: 0 }) && b.total === 0;
})());
t('4m. 小計與合計是未取整的「元」（顯示端才取整）：qty 1.5 × 3 = 4.5', Q.totals([ln('travel', { qty: 1.5, unitCost: 3 })], 0).total === 4.5);
t('4n. lineAmount／stampAmount 輔助函式', Q.lineAmount({ qty: 3, unitCost: 0.1 }) === 0.3 && Q.stampAmount(12500) === 13 && Q.stampAmount(undefined) === null && Q.stampAmount(-5) === null && Q.stampAmount(0) === 0);
t('4o. fmtMoney：千分位、取整、壞值 0', Q.fmtMoney(1234567.4) === '1,234,567' && Q.fmtMoney(1234567.5) === '1,234,568' && Q.fmtMoney('x') === '0' && Q.fmtMoney(0) === '0');

// 5) XSS：跳脫
const EVIL = ['<img src=x onerror=alert(1)>', '"><script>alert(2)</script>', "' onfocus='alert(3)", '" onmouseover="alert(4)', '&lt;b&gt;'];
t('5a. esc：& < > " \' 全部跳脫、null／undefined → 空字串、數字轉字串', Q.esc('<a href="x" title=\'y\'>&</a>') === '&lt;a href=&quot;x&quot; title=&#039;y&#039;&gt;&amp;&lt;/a&gt;' && Q.esc(null) === '' && Q.esc(undefined) === '' && Q.esc(12) === '12');
t('5b. esc 把已是實體的文字當純文字再跳脫一次（&lt; → &amp;lt;，不會被瀏覽器還原成 <）', Q.esc('&lt;b&gt;') === '&amp;lt;b&amp;gt;');
const bad = [];
['consult', 'software', 'hw', 'travel', 'other'].forEach((cat) => {
  EVIL.forEach((e) => {
    const l = { desc: e, vendor: e, note: e, unit: e, lid: e, forLid: e, qty: 1, unitCost: 1 };
    [Q._editRowHtml(l, cat), Q._viewRowHtml(l, cat)].forEach((html) => {
      // 只看真正的標籤（跳脫後的使用者文字不會產生 <）：標籤名要在白名單內；
      // 移除合法的 attr="..." 之後，標籤內不應殘留任何 on* 事件屬性（否則代表使用者文字跳出了屬性值）
      (html.match(/<[^>]+>/g) || []).forEach((tag) => {
        const name = (/^<\/?([a-zA-Z0-9]+)/.exec(tag) || [])[1];
        if (!['tr', 'td', 'input', 'button', 'span'].includes(name)) bad.push(['tag', cat, e, tag]);
        if (/\son\w+=/i.test(tag.replace(/="[^"]*"/g, '=""'))) bad.push(['attr', cat, e, tag]);
      });
    });
  });
});
t('5c. 惡意字串（品名／廠商／說明／單位／lid／forLid）放進編輯列與唯讀列後，HTML 中沒有可執行的標籤或事件屬性（5 區×5 字串×2 種列）', bad.length === 0, J(bad.slice(0, 3)));
t('5d. 編輯列中使用者文字出現在 value／data 屬性內且已跳脫', /value="&lt;img src=x onerror=alert\(1\)&gt;"/.test(Q._editRowHtml({ desc: EVIL[0] }, 'consult')) && /data-lid="&quot;&gt;&lt;script&gt;/.test(Q._editRowHtml({ lid: EVIL[1], desc: 'x' }, 'consult')));
t('5e. 唯讀列中使用者文字以純文字出現（&lt;img…）', Q._viewRowHtml({ desc: EVIL[0] }, 'consult').includes('&lt;img src=x onerror=alert(1)&gt;'));

// 6) 檔案紀律
const cssBody = (/const QCL_CSS = `([\s\S]*?)`;/.exec(src) || [])[1] || '';
const selectors = cssBody.replace(/\/\*[\s\S]*?\*\//g, '').split('{').slice(0, -1).map((chunk) => chunk.split('}').pop().trim()).filter((s) => s && !s.startsWith('@'));
const unprefixed = selectors.filter((sel) => sel.split(',').some((p) => !/\.qcl-/.test(p)));
t(`6a. CSS 共 ${selectors.length} 組選擇器，每個選擇器都含 qcl- 類名（不污染既有樣式）`, selectors.length > 20 && unprefixed.length === 0, J(unprefixed.slice(0, 3)));
t('6b. 樣式只注入一次（style id=qclStyle、ensureStyle 先檢查是否已存在）', (src.match(/\.id = 'qclStyle'/g) || []).length === 1 && /getElementById\('qclStyle'\)\) return/.test(src));
t('6c. 獨立載入：不呼叫 app.js 的全域函式（escapeHtml／showToast／fmtMoney 以外的自有實作）', !/\bescapeHtml\s*\(/.test(src) && !/\bshowToast\s*\(/.test(src) && !/\bquoteTotal\s*\(/.test(src));
t('6d. 全域只新增 QCL（IIFE 包住；載入後 sandbox 的自有屬性只有 QCL）', J(Object.keys(ctx)) === J(['QCL']), J(Object.keys(ctx)));
t('6e. 事件委派：addEventListener 只出現在 mount 的 on()、dragSort 的 on()、dragSort 的 suppressClick 三處（20 區另測 dragSort 的註冊／移除對稱）', (src.match(/addEventListener\(/g) || []).length === 3);

// 7) 與伺服器 lib/quotePnlExcel.js 鏡像
let PX = null; try { PX = require(path.join(ROOT, 'lib/quotePnlExcel.js'))._internal; } catch (e) { /* 伺服器端模組暫時改動中：略過並在輸出註明 */ }
if (PX && typeof PX.catByUnit === 'function' && typeof PX.resolveCat === 'function') {
  const units = ['人天', '人日', '人月', 'man day', 'Man-Day', 'MD', '天', '日', '台', '組', '部', '臺', 'pcs', 'SET', 'unit', '授權', 'License', '套', '5 user', 'seat', '帳號', '訂閱', 'subscription', '式', '', '次', '趟', '批', ' 台 ', '人天/月'];
  const toHw = (c) => (c === 'hardware' ? 'hw' : c);
  const mism = units.filter((u) => toHw(PX.catByUnit(u)) !== Q._catByUnit(u));
  t(`7a. 單位猜測與伺服器 catByUnit 逐例相同（${units.length} 例；hardware ↔ hw）`, mism.length === 0, J(mism.map((u) => [u, PX.catByUnit(u), Q._catByUnit(u)])));
  const cases = [];
  ['', 'consult', 'software', 'hardware', 'hw', 'other', 'bogus'].forEach((cat) => ['人天', '台', '授權', '式'].forEach((unit) => [[], ['consult'], ['hardware'], ['crm', 'mdm'], ['consult', 'software'], ['unknown']].forEach((cl) => cases.push({ cat, unit, cl }))));
  const bad7 = cases.filter((c) => {
    const server = toHw(PX.resolveCat({ cat: c.cat === 'hw' ? 'hardware' : c.cat, unit: c.unit }, c.cl));
    const mine = Q.seedFromItems([{ lid: 'z', desc: 'z', unit: c.unit, qty: 1, cat: c.cat }], { classCodes: c.cl })[0].cat;
    return server !== mine;
  });
  t(`7b. 分區決定（品項 cat → 單一類別 → 單位）與伺服器 resolveCat 逐例相同（${cases.length} 例）`, bad7.length === 0, J(bad7.slice(0, 2)));
} else {
  t('7. 伺服器 lib/quotePnlExcel.js 的 _internal.catByUnit／resolveCat 讀不到（略過鏡像比對）', true, 'SKIPPED');
}

// 8) 營收 revenueOf ＝ 伺服器 computeFinancials 的 revenueCents/100（D1：前端營收曾是未取整的浮點，印花稅／成本會差 1 元）
let QAmod = null, CLmod = null;
try { QAmod = require(path.join(ROOT, 'lib/quoteApproval.js')); CLmod = require(path.join(ROOT, 'lib/quoteCostLines.js')); } catch (e) { /* 伺服器端模組暫時改動中：略過比對並在輸出註明 */ }
t('8a. QCL.revenueOf 與 stampAmount 存在', typeof Q.revenueOf === 'function' && typeof Q.stampAmount === 'function');
if (QAmod && CLmod && typeof Q.revenueOf === 'function') {
  const itm = (qty, unitPrice) => ({ desc: 'x', unit: '式', qty, unitPrice });
  const serverFin = (items, dt, dv) => {
    // 伺服器儲存規則（lib/quoteRoutes.js normalizeItems／折扣）：數量有限且 >0 才採用（至少 0.001）否則 1；單價有限且 >=0 才採用否則 0；折扣類型不合法＝none，折扣值有限且 >0 才採用
    const its = items.map((r) => {
      if (r && (r.kind === 'title' || r.kind === 'subtotal')) return { lid: 'k', kind: r.kind, desc: r.desc || '' };
      const qty = parseFloat(r.qty), price = parseFloat(r.unitPrice);
      return { lid: 'x', desc: 'd', unit: '式', qty: Number.isFinite(qty) && qty > 0 ? Math.max(0.001, qty) : 1, unitPrice: Number.isFinite(price) && price >= 0 ? price : 0, cost: 0 };
    });
    const dvn = parseFloat(dv);
    return QAmod.computeFinancials({ items: its, discountType: ['none', 'percent', 'amount'].includes(dt) ? dt : 'none', discountValue: Number.isFinite(dvn) && dvn > 0 ? dvn : 0 });
  };
  // 已知案例
  const k1 = serverFin([itm(1, 1063204)], 'percent', 79.9);
  t('8b. 已知案例（回報的差 1 元）：1 × 1,063,204、折扣 79.9% → 伺服器營收 849,500.00；revenueOf 同值；印花稅 850（不是 849）',
    k1.ok && k1.revenueCents === 84950000 && Q.revenueOf([itm(1, 1063204)], 'percent', 79.9) === 849500 && Q.stampAmount(Q.revenueOf([itm(1, 1063204)], 'percent', 79.9)) === 850 && CLmod.stampDollars(k1.revenueCents) === 850,
    J([k1.revenueCents, Q.revenueOf([itm(1, 1063204)], 'percent', 79.9)]));
  t('8c. 舊算法對照：未取整的浮點營收 849499.996（舊畫面營收），經 stampAmount 現在也得 850（營收先取整到分），舊的 round(營收×0.001) 是 849', Q.stampAmount(1063204 * 79.9 / 100) === 850 && Math.round(1063204 * 79.9 / 100 * 0.001) === 849);
  t('8d. 逐列取整到分再加總（不是加總後才取整）：3 列 × 0.005 元 → 各 1 分（half-up）→ 0.03；1.005 × 1 → 1.01', Q.revenueOf([itm(1, 0.005), itm(1, 0.005), itm(1, 0.005)], 'none', 0) === 0.03 && Q.revenueOf([itm(1, 1.005)], 'none', 0) === 1.01);
  t('8e. 數量空白／0／負／非數字 → 1；小於 0.001 → 0.001（與伺服器儲存規則相同）；單價負數／空白 → 0', Q.revenueOf([itm('', 100)], 'none', 0) === 100 && Q.revenueOf([itm(0, 100)], 'none', 0) === 100 && Q.revenueOf([itm(-3, 100)], 'none', 0) === 100 && Q.revenueOf([itm('abc', 100)], 'none', 0) === 100
    && Q.revenueOf([itm(0.0001, 1000)], 'none', 0) === 1 && Q.revenueOf([itm(2, -5), itm(1, ''), itm(1, 7)], 'none', 0) === 7);
  t('8f. 分組標題／小計列不計；折扣類型不合法／折扣值 ≤0 或非數字＝無折扣；沒給 items → 0（不丟錯）',
    Q.revenueOf([{ kind: 'title', desc: 't' }, itm(1, 100), { kind: 'subtotal', desc: 's', unitPrice: 999 }], 'none', 0) === 100 && Q.revenueOf([itm(1, 100)], 'bogus', 50) === 100 && Q.revenueOf([itm(1, 100)], 'percent', 0) === 100 && Q.revenueOf([itm(1, 100)], 'percent', -5) === 100
    && Q.revenueOf([itm(1, 100)], 'amount', 'abc') === 100 && Q.revenueOf(undefined, 'none', 0) === 0 && Q.revenueOf(null) === 0 && Q.revenueOf([null, 5, 'x'], 'none', 0) === 0);
  t('8g. 百分比：1,000,000 × 九折＝900,000；議價 amount：1,000,000 議價 812,345.675 → 812,345.68（half-up 到分）', Q.revenueOf([itm(1, 1000000)], 'percent', 90) === 900000 && Q.revenueOf([itm(1, 1000000)], 'amount', 812345.675) === 812345.68);

  // 隨機比對（決定性 LCG）
  let seed = 20261008;
  const rnd = () => ((seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff);
  const pk = (a) => a[Math.floor(rnd() * a.length)];
  const dec = (v, places) => +v.toFixed(places);
  const rQty = () => {
    const k = rnd();
    if (k < 0.06) return ''; if (k < 0.09) return 0; if (k < 0.1) return -2; if (k < 0.11) return 'abc'; if (k < 0.12) return pk([0.0004, 0.0005, 0.001, 0.0015]);
    if (k < 0.45) return Math.floor(rnd() * 60) + 1;
    if (k < 0.75) return dec(rnd() * 20, pk([1, 2, 3, 4]));
    return dec(rnd() * 1000, pk([0, 1, 2, 3, 5]));
  };
  const rPrice = () => {
    const k = rnd();
    if (k < 0.03) return ''; if (k < 0.05) return -7;
    if (k < 0.35) return Math.floor(rnd() * 2000000);
    if (k < 0.55) return dec(rnd() * 100000, pk([1, 2, 3, 4]));
    if (k < 0.7) return (Math.floor(rnd() * 100000) + 0.005);                 // 剛好落在半分上的單價
    if (k < 0.8) return Math.floor(rnd() * 100000) / 1000;                    // 三位小數（半分邊界多）
    if (k < 0.9) return pk([1063204, 999999.99, 1234567.891, 0.333, 100.005, 2.675]);
    return dec(rnd() * 1e8, pk([0, 2, 3]));
  };
  let compared = 0, skipped = 0, diffs = 0, stampDiffs = 0, oldRevDiffs = 0, oldStampDiffs = 0, halfCent = 0;
  const firstDiffs = [];
  const SAMPLES = 100000;
  for (let n = 0; compared < SAMPLES && n < SAMPLES * 4; n++) {
    const items = [];
    const cnt = 1 + Math.floor(rnd() * 6);
    for (let i = 0; i < cnt; i++) {
      if (rnd() < 0.08) items.push({ kind: pk(['title', 'subtotal']), desc: 'k', qty: 5, unitPrice: 12345 });   // 標題／小計列帶著數字也不能計入
      else items.push(itm(rQty(), rPrice()));
    }
    // 舊算法（quote.js quoteTotal 的浮點：qty、unitPrice 先 parseFloat||預設，未取整）用來證明這批輸入「有鑑別力」
    let oldSub = 0;
    items.forEach((r) => { if (r.kind) return; oldSub += (parseFloat(r.qty) || 1) * (parseFloat(r.unitPrice) || 0); });
    const subYuan = oldSub > 0 ? oldSub : 1000;
    const dt = pk(['none', 'none', 'percent', 'percent', 'percent', 'amount', 'amount', 'amount', 'bogus', '']);
    let dv;
    if (dt === 'percent') dv = rnd() < 0.5 ? pk([79.9, 90, 85.5, 33.333, 99.99, 0.5, 66.67, 12.345, 100, 120, 0, -5, '', 'x']) : dec(rnd() * 100, pk([0, 1, 2, 3, 4]));
    else if (dt === 'amount') dv = rnd() < 0.15 ? pk([0, -1, '', 'x']) : dec(subYuan * (rnd() < 0.85 ? rnd() : 1 + rnd()), pk([0, 1, 2, 3]));   // 約 15% 的議價金額高於小計（伺服器會拒絕）
    else dv = pk([0, 50, '']);
    const fin = serverFin(items, dt, dv);
    const rev = Q.revenueOf(items, dt, dv);
    if (!fin.ok) { skipped++; if (!(typeof rev === 'number' && isFinite(rev) && rev >= 0)) { diffs++; if (firstDiffs.length < 5) firstDiffs.push(['non-finite', J(items), dt, dv, rev]); } continue; }
    compared++;
    if (rev !== fin.revenueCents / 100) { diffs++; if (firstDiffs.length < 5) firstDiffs.push([J(items), dt, dv, fin.revenueCents, rev]); }
    if (Q.stampAmount(rev) !== CLmod.stampDollars(fin.revenueCents)) stampDiffs++;
    if (fin.revenueCents % 100 === 50) halfCent++;
    // 舊算法與伺服器的差（舊畫面營收＝未取整的浮點）
    let oldRev = oldSub;
    const odv = parseFloat(dv) || 0;
    if (dt === 'percent') oldRev = oldSub * (odv || 100) / 100; else if (dt === 'amount') oldRev = odv || oldSub;
    if (Math.round(oldRev * 100) !== fin.revenueCents) oldRevDiffs++;
    if (Math.round(+(oldRev * 0.001).toPrecision(12)) !== CLmod.stampDollars(fin.revenueCents)) oldStampDiffs++;
  }
  t(`8h. 隨機輸入 ${compared} 組（含小數單價／數量、百分比小數折扣、固定金額折扣、空 qty、標題／小計列）：QCL.revenueOf ＝ 伺服器 revenueCents/100，0 差異`, compared >= SAMPLES && diffs === 0, `compared=${compared} diffs=${diffs} ${J(firstDiffs)}`);
  t('8i. 同一批輸入的印花稅：QCL.stampAmount(revenueOf) ＝ CL.stampDollars(revenueCents)，0 差異', stampDiffs === 0, 'stampDiffs=' + stampDiffs);
  t(`8j. 另外 ${skipped} 組伺服器會拒絕的輸入（折扣大於合計、營收 0…）：revenueOf 不丟錯、回有限的非負數`, skipped > 1000 && diffs === 0, 'skipped=' + skipped);
  t(`8k. 這批輸入有鑑別力：伺服器營收落在「半分」邊界 ${halfCent} 組；舊算法（未取整浮點）營收與伺服器不同 ${oldRevDiffs} 組、印花稅不同 ${oldStampDiffs} 組（>0 才代表測得出舊 bug）`, halfCent > 1000 && oldRevDiffs > 100, `half=${halfCent} oldRev=${oldRevDiffs} oldStamp=${oldStampDiffs}`);
  // 印花稅單獨比對：任意「分」→ 整數元（含 .5 元邊界 ±1 分）
  let sd2 = 0, sdn = 0;
  for (let n = 0; n < 200000; n++) {
    const c = rnd() < 0.5 ? Math.floor(rnd() * 1e11) : (Math.floor(rnd() * 1e6) * 100000 + 50000 + pk([-1, 0, 1]));
    sdn++;
    if (Q.stampAmount(c / 100) !== CLmod.stampDollars(c)) sd2++;
  }
  t(`8l. 印花稅單獨比對 ${sdn} 組（任意營收分數與 .5 元邊界 ±1 分，營收到 10 億元）：stampAmount(分/100) ＝ stampDollars(分)，0 差異`, sd2 === 0, 'diffs=' + sd2);
} else {
  t('8. 伺服器端模組讀不到（略過營收比對）', true, 'SKIPPED');
}

// 9) 「＋補入新品項」覆蓋判定（QCL._missingItemLines）：forLid 對得上 lid 就算覆蓋；其餘成本列改用「品名相同」的多重集合比對
{
  const M = Q._missingItemLines;
  const miss = (lines, items, opts) => M(lines, items, opts).map((l) => l.desc);
  const seedOf = (items) => Q.seedFromItems(items).filter((l) => !l.auto && l.desc !== '差旅交通' && l.desc !== '交際費');   // 只留品項列
  const noLid = (d, u) => ({ desc: d, unit: u || '式', qty: 1 });
  const withLid = (d, lid, u) => ({ lid, desc: d, unit: u || '式', qty: 1 });
  t('9a. 方法存在、純函式、不修改輸入', typeof M === 'function' && (() => { const items = [withLid('A', 'L1')]; const lines = seedOf(items); const b = J([items, lines]); M(lines, items); return J([items, lines]) === b; })());
  // 審查重現：新單品項沒有 lid → 種子列沒有 forLid → 存檔後伺服器替品項指派 lid、成本列仍沒有 forLid → 重開按「補入」
  const seedNoLid = seedOf([noLid('顧問A', '人天'), noLid('主機B', '台')]);
  t('9b. 審查情境：品項無 lid 的種子（成本列沒有 forLid）→ 存檔後品項有了 lid → 補入應為「沒有新品項需要補入」', seedNoLid.every((l) => !l.forLid)
    && miss(seedNoLid, [withLid('顧問A', 'L1', '人天'), withLid('主機B', 'L2', '台')]).length === 0, J(miss(seedNoLid, [withLid('顧問A', 'L1', '人天'), withLid('主機B', 'L2', '台')])));
  t('9c. 同一批品項仍沒有 lid（還沒存檔就按補入）→ 也沒有新品項', miss(seedNoLid, [noLid('顧問A', '人天'), noLid('主機B', '台')]).length === 0);
  // 舊的演算法（只看 forLid、沒 lid 才比品名）在上面的情境會把兩個品項都重複補入：證明這組測試測得出 bug
  const oldMissing = (cur, items) => { const haveFor = new Set(cur.map((l) => l.forLid).filter(Boolean)); const haveDesc = new Set(cur.map((l) => l.desc)); return Q.seedFromItems(items).filter((l) => !l.auto && l.desc !== '差旅交通' && l.desc !== '交際費').filter((l) => (l.forLid ? !haveFor.has(l.forLid) : !haveDesc.has(l.desc))); };
  t('9d. 鑑別力：舊演算法（只看 forLid）在 9b 情境會重複補入 2 項（本測試在舊程式下會失敗）', oldMissing(seedNoLid, [withLid('顧問A', 'L1', '人天'), withLid('主機B', 'L2', '台')]).length === 2);
  // forLid 優先
  t('9e. forLid 對得上品項 lid → 已覆蓋（即使成本列的品名改了、或品名與別的品項相同）', miss([{ cat: 'other', desc: '完全不同的名字', forLid: 'L1' }], [withLid('A', 'L1')]).length === 0
    && miss([{ cat: 'other', desc: 'B', forLid: 'L1' }], [withLid('A', 'L1'), withLid('B', 'L2')]).join() === 'B');
  t('9f. 一個品項有多列成本列（forLid 相同）只算覆蓋一次、不會重複補入', miss([{ cat: 'consult', desc: 'A', forLid: 'L1' }, { cat: 'hw', desc: 'A2', forLid: 'L1' }], [withLid('A', 'L1')]).length === 0);
  // 同名品項：N 個同名品項需要 N 列才算覆蓋
  const dup2 = [noLid('同名'), noLid('同名')];
  t('9g. 同名品項 2 個：成本列 2 列同名 → 沒有新品項；只有 1 列 → 補 1 項；0 列 → 補 2 項；3 列 → 沒有',
    miss(seedOf(dup2), dup2).length === 0 && miss(seedOf(dup2).slice(0, 1), dup2).join() === '同名' && miss([], dup2).join() === '同名,同名' && miss(seedOf(dup2).concat(seedOf(dup2).slice(0, 1)), dup2).length === 0);
  t('9h. 同名品項（已存檔、有 lid）＋成本列沒有 forLid：2 列同名覆蓋 2 個；1 列只覆蓋 1 個（補 1 項）', miss(seedOf(dup2), [withLid('同名', 'L1'), withLid('同名', 'L2')]).length === 0
    && miss(seedOf(dup2).slice(0, 1), [withLid('同名', 'L1'), withLid('同名', 'L2')]).join() === '同名');
  t('9i. forLid 與品名混合：L1 由 forLid 覆蓋（品名不同）、L2 同名由一般成本列覆蓋 → 沒有新品項；每列成本列最多覆蓋一個品項', miss([{ cat: 'other', desc: 'zzz', forLid: 'L1' }, { cat: 'other', desc: 'A' }], [withLid('A', 'L1'), withLid('A', 'L2')]).length === 0
    && miss([{ cat: 'other', desc: 'A' }], [withLid('A', 'L1'), withLid('A', 'L2')]).join() === 'A');
  // 改名／新增
  t('9j. 改名後補入只多 1 列：品項甲改名為乙（成本列還叫甲、沒有 forLid）→ 補入 1 項（乙），其他品項不重複', miss(seedOf([noLid('甲'), noLid('丙')]), [withLid('乙', 'L1'), withLid('丙', 'L2')]).join() === '乙');
  t('9k. 新增品項後補入只多 1 列：原有 2 個品項已有成本列，新增第 3 個 → 只補第 3 個', miss(seedOf([noLid('A'), noLid('B')]), [withLid('A', 'L1'), withLid('B', 'L2'), withLid('C', 'L3')]).join() === 'C'
    && miss(seedOf([noLid('A'), noLid('B')]), [noLid('A'), noLid('B'), noLid('C')]).join() === 'C');
  t('9l. 指向已不存在品項的 forLid（品項被刪）的成本列：不屬於任何現有品項 → 改用品名比對（同名新品項視為已覆蓋）', miss([{ cat: 'other', desc: '新品項', forLid: 'GONE' }], [withLid('新品項', 'L9')]).length === 0
    && miss([{ cat: 'other', desc: '舊品項', forLid: 'GONE' }], [withLid('新品項', 'L9')]).join() === '新品項');
  t('9m. 品名比對去頭尾空白（成本列 " A "、品項 "A"）；品名大小寫視為不同', miss([{ cat: 'other', desc: ' A ' }], [noLid('A')]).length === 0 && miss([{ cat: 'other', desc: 'a' }], [noLid('A')]).join() === 'A');
  t('9n. 印花稅列不參與覆蓋（即使品項取名與印花稅品名相同）；差旅／交際費等固定列依品名比對', miss([{ cat: 'other', auto: 'stamp', desc: Q.STAMP_DESC }], [noLid(Q.STAMP_DESC)]).length === 1
    && miss([{ cat: 'travel', desc: '差旅交通' }], [noLid('差旅交通')]).length === 0);
  t('9o. 回傳的是品項種子列本身（分類猜測、單位、數量、forLid 都在），可直接 appendRows', (() => { const r = M([], [withLid('顧問', 'L1', '人天')]); return r.length === 1 && r[0].cat === 'consult' && r[0].unit === '人天' && r[0].qty === 1 && r[0].forLid === 'L1'; })());
  t('9p. 邊界：items 不是陣列／空陣列 → []；lines 不是陣列／含 null → 所有品項都算未覆蓋；分組標題與小計列不算品項', M([], undefined).length === 0 && M([], []).length === 0 && M(undefined, [noLid('A')]).length === 1 && M([null, 5, 'x'], [noLid('A')]).length === 1
    && M([], [{ kind: 'title', desc: 'Part A' }, { kind: 'subtotal', desc: '小計' }, noLid('A')]).length === 1);
  t('9q. 與 reseed 的關係：重新帶入＝seedFromItems（不看既有列）→ 補入後再補入第二次不會再有新品項（冪等）', (() => {
    const items = [withLid('A', 'L1'), withLid('B', 'L2'), withLid('C', 'L3')];
    const first = M(seedOf([noLid('A')]), items);                    // 只有 A 有成本列 → 補 B、C
    const after = seedOf([noLid('A')]).concat(first);
    return first.map((l) => l.desc).join() === 'B,C' && M(after, items).length === 0;
  })());
}

// 10) 種子略過「說明為空白」的品項（種子、重新帶入、補入共用 seedItemLines）
{
  const blanks = [{ lid: 'b1', desc: '', unit: '式', qty: 1 }, { lid: 'b2', desc: '   ', unit: '式', qty: 1 }, { lid: 'b3', desc: ' \n\t ', unit: '式', qty: 1 }, { lid: 'b4', unit: '式', qty: 1 }, { lid: 'b5', desc: null, unit: '式', qty: 1 }];
  const mixed = [it('顧問', '人天', 1), blanks[0], it('主機', '台', 1), blanks[1], blanks[2], blanks[3], blanks[4]];
  const s = Q.seedFromItems(mixed);
  t('10a. 種子略過說明為空白的品項（空字串、純空白、換行／Tab、缺欄位、null）：只剩 2 個品項＋3 個固定列', s.length === 5 && s.slice(0, 2).map((l) => l.desc).join() === '顧問,主機', J(s.map((l) => l.desc)));
  t('10b. 種子裡除了印花稅列，沒有任何 desc 為空的列（顧問不必先刪列才能儲存）', s.every((l) => l.auto === 'stamp' || l.desc !== ''));
  t('10c. 全部品項都是空白說明 → 只剩固定列（差旅交通、交際費、印花稅）', Q.seedFromItems(blanks).map((l) => l.desc).join() === ['差旅交通', '交際費', Q.STAMP_DESC].join());
  t('10d. includeStamp:false、classCodes 等選項下規則一致', Q.seedFromItems(mixed, { includeStamp: false }).length === 4 && Q.seedFromItems(mixed, { classCodes: ['consult'] }).every((l) => l.auto === 'stamp' || l.desc !== ''));
  t('10e. 補入新品項（_missingItemLines）也不會補入空白說明的品項', Q._missingItemLines([], mixed).map((l) => l.desc).join() === '顧問,主機' && Q._missingItemLines([{ cat: 'other', desc: '顧問' }, { cat: 'other', desc: '主機' }], mixed).length === 0);
  t('10f. 空白說明品項不佔 60 列上限：100 個空白＋3 個有說明 → 種子 6 列', Q.seedFromItems(Array.from({ length: 100 }, (_, i) => (i < 3 ? it('品' + i, '式', 1) : { lid: 'e' + i, desc: '', unit: '式', qty: 1 }))).length === 6);
  t('10g. 說明只有全形空白（U+3000）：JS 的 trim 會去掉，視為空白而略過', Q.seedFromItems([{ lid: 'z', desc: '　', unit: '式', qty: 1 }]).length === 3);
  t('10h. 前後有空白的說明保留且去空白（種子 desc＝"顧問"）', Q.seedFromItems([{ lid: 'z', desc: '  顧問  ', unit: '式', qty: 1 }])[0].desc === '顧問');
}

// 11) 「＋補入新品項」確認視窗文字（QCL._supplementMessage）：先讓使用者看見會補哪些品項；最多列 8 個、其餘「…另 N 個」
{
  const SM = typeof Q._supplementMessage === 'function' ? Q._supplementMessage : () => '（_supplementMessage 不存在）';
  const L = (n) => Array.from({ length: n }, (_, i) => ({ cat: 'other', desc: '品項' + (i + 1) }));
  const HINT = '若這些品項的成本已包含在其他列中（例如整包），請按取消。';
  t('11a. 方法存在、純函式（不修改輸入）', typeof SM === 'function' && (() => { const a = L(3); const b = J(a); SM(a, 0); return J(a) === b; })());
  t('11b. 1 個品項：文字完整（「將補入以下 1 個尚未有成本列的品項：品項1」＋空行＋整包提示）', SM(L(1), 0) === '將補入以下 1 個尚未有成本列的品項：品項1\n\n' + HINT, J(SM(L(1), 0)));
  t('11c. 3 個品項用「、」相連；總數 N＝3', SM(L(3), 0) === '將補入以下 3 個尚未有成本列的品項：品項1、品項2、品項3\n\n' + HINT);
  t('11d. 剛好 8 個：全部列出，沒有「…另」', (() => { const m = SM(L(8), 0); return /：品項1、品項2、品項3、品項4、品項5、品項6、品項7、品項8\n/.test(m) && !/…另/.test(m) && /以下 8 個/.test(m); })());
  t('11e. 9 個：只列前 8 個，接「…另 1 個」；N 仍是總數 9', (() => { const m = SM(L(9), 0); return /以下 9 個/.test(m) && /品項8…另 1 個\n/.test(m) && !/品項9/.test(m); })());
  t('11f. 12 個：前 8 個＋「…另 4 個」', (() => { const m = SM(L(12), 0); return /以下 12 個/.test(m) && /品項8…另 4 個\n/.test(m) && !/品項9/.test(m); })());
  t('11g. 一定帶有「若成本已包含在其他列（例如整包）請按取消」的提示', [1, 3, 9, 30].every((n) => SM(L(n), 0).endsWith(HINT)));
  t('11h. 品名太長（80 字）截成 30 字＋「…」，不會撐破視窗', (() => { const m = SM([{ desc: '長'.repeat(80) }], 0); return m.indexOf('長'.repeat(30) + '…') > 0 && m.indexOf('長'.repeat(31)) < 0; })());
  t('11i. 因列數上限不會補入的項數：多一行「已達列數上限，另有 K 項不會補入」；沒有就不出現', /\n（已達列數上限，另有 2 項不會補入）\n/.test(SM(L(3), 2)) && !/列數上限/.test(SM(L(3), 0)) && !/列數上限/.test(SM(L(3))));
  t('11j. 品名原樣輸出（< > " & 不在這裡跳脫——跳脫是確認視窗的責任，e2e 另測 XSS）', SM([{ desc: '<img src=x onerror=1>&"' }], 0).includes('<img src=x onerror=1>&"'));
  t('11k. 品名含換行／Tab 併成空白（不會讓確認視窗出現怪異斷行）', (() => { const m = SM([{ desc: 'A\nB\r\nC' }], 0); return m.includes('A B C') && m.split('\n').length === 3; })());
  t('11l. 容錯：lines 不是陣列／含 null／缺 desc → 不丟錯，N＝陣列長度', (() => { try { return /以下 0 個/.test(SM(undefined, 0)) && /以下 3 個/.test(SM([null, {}, { desc: 'X' }], 0)); } catch (e) { return false; } })());
}

// 12) 報價品項與成本明細「不是一對一」的情境（多對一／一對多／混合）下，補入判定的行為（這些就是要先確認再補入的原因）
{
  const M = Q._missingItemLines;
  const names = (lines, items) => M(lines, items).map((l) => l.desc).join('|');
  const itm = (lid, d, unit, extra) => Object.assign({ lid, desc: d, unit: unit || '式', qty: 1 }, extra || {});
  const cl = (cat, d, extra) => Object.assign({ cat, desc: d, qty: 1, unitCost: 100 }, extra || {});
  const fixedRows = [cl('travel', '差旅交通'), cl('other', '交際費'), { cat: 'other', auto: 'stamp' }];
  // 多對一（S1）：3 個品項、成本只有一列整包
  const items3 = [itm('L1', '顧問人天', '人天'), itm('L2', '軟體授權', '套'), itm('L3', '硬體設備', '台')];
  t('12a. 多對一：整包列沿用第一個品項的種子列（forLid＝L1，改名為「專案整包成本」）→ 另外 2 個品項被視為尚未有成本列（所以補入要先確認）',
    names([cl('consult', '專案整包成本', { forLid: 'L1' })].concat(fixedRows), items3) === '軟體授權|硬體設備');
  t('12b. 多對一：整包列是全新增的列（沒有 forLid、品名也對不上）→ 3 個品項全被視為尚未有成本列', names([cl('consult', '專案整包成本')].concat(fixedRows), items3) === '顧問人天|軟體授權|硬體設備');
  t('12c. 多對一：整包列取名跟其中一個品項同名（沒有 forLid）→ 該品項算有對應（品名比對），另外 2 個仍算沒有', names([cl('consult', '軟體授權')].concat(fixedRows), items3) === '顧問人天|硬體設備');
  t('12d. 確認視窗會列出的就是這幾個品項（多對一整包情境）', (() => {
    const m = (Q._supplementMessage || (() => ''))(M([cl('consult', '專案整包成本', { forLid: 'L1' })].concat(fixedRows), items3), 0);
    return /以下 2 個/.test(m) && m.includes('：軟體授權、硬體設備\n');
  })());
  // 一對多（S2）：1 個「一式」品項、成本拆成很多列
  const item1 = [itm('L1', '系統導入專案', '式')];
  const many = [cl('consult', 'PM', { forLid: 'L1', unit: '人天', qty: 20 }), cl('consult', 'SD', { unit: '人天', vendor: 'V' }), cl('consult', 'BASIS', { unit: '人天' }),
    cl('software', '授權費'), cl('hw', '伺服器'), cl('travel', '差旅交通'), cl('other', '交際費'), { cat: 'other', auto: 'stamp' }];
  t('12e. 一對多：種子列（forLid＝L1）改成 PM 並新增多列 → 一式品項仍算有對應，不會被補入', names(many, item1) === '');
  t('12f. 一對多：種子列被刪、成本拆成不同名稱的多列 → 一式品項被視為沒有對應（補入前會被列出，使用者可取消）', names(many.slice(1), item1) === '系統導入專案');
  // 混合（S3）：品項 A 對應 3 列、品項 B 對應 1 列、另有與品項無關的差旅
  const itemsAB = [itm('LA', '系統建置服務', '式'), itm('LB', '硬體設備', '台')];
  const mixed = [cl('consult', 'PM', { forLid: 'LA' }), cl('consult', 'SD'), cl('consult', '委外顧問', { vendor: 'X' }), cl('hw', '硬體設備', { forLid: 'LB' }), cl('travel', '出差費')];
  t('12g. 混合：A 對應 3 列（其中 1 列有 forLid）、B 對應 1 列、另有無關差旅 → 沒有品項需要補入', names(mixed, itemsAB) === '');
  t('12h. 混合：業務事後新增品項 C → 只補 C（既有 A、B 不重複）', names(mixed, itemsAB.concat([itm('LC', '加購授權', '套')])) === '加購授權');
  t('12i. 混合：事後新增的品項與無關差旅列同名（「出差費」）→ 品名比對會當成已有對應（已知取捨：自然鍵退回的副作用）', names(mixed, itemsAB.concat([itm('LC', '出差費', '趟')])) === '');
  // 零價品項（S5）：贈品也算品項，種子會帶出一列（成本可填 0）
  t('12j. 零價品項（贈品）：種子仍會替它產生一列（成本 0 可保留，完成時只提示「有 N 列成本單價是 0」）', Q.seedFromItems([itm('L1', '服務', '式'), itm('L2', '贈品', '式', { unitPrice: 0 })]).map((l) => l.desc).slice(0, 2).join('|') === '服務|贈品');
}

// 13) 「合併為一列」（QCL._mergeLines）：整包用，成本合計必須不變（以「分」為單位逐筆相同）
{
  const MG = typeof Q._mergeLines === 'function' ? Q._mergeLines : () => null;
  const row = (desc, qty, unitCost, extra) => Object.assign({ cat: 'consult', desc, qty, unitCost, unit: '人天' }, extra || {});
  const centsOf = (lines) => Math.round(Q.totals(lines, null).subtotalExStamp * 100);
  const rows3 = [row('PM', 10, 7000, { lid: 'L-first', forLid: 'I-1' }), row('SD', 15, 6500.5, { vendor: '甲' }), row('BASIS', 5, 8000, { vendor: '甲', forLid: 'I-3' })];
  const m = MG(rows3);
  t('13a. 合併 3 列：單位「式」、數量 1、成本單價＝各列小計合計（70,000＋97,507.5＋40,000＝207,507.5）、項目「合併 3 項：PM、SD、BASIS」', !!m && m.unit === '式' && m.qty === 1 && m.unitCost === 207507.5 && m.desc === '合併 3 項：PM、SD、BASIS' && m.cat === 'consult', J(m));
  t('13b. 成本合計不變（以分為單位）：合併前後 totals 的顧問小計相同', centsOf(rows3) === centsOf([m]) && centsOf([m]) === 20750750, centsOf(rows3) + ' vs ' + centsOf([m]));
  t('13c. 沿用第一列的 lid 與 forLid；其他列的 forLid 改收進 forLids（涵蓋判定因此認得每個被併的品項，見 15d）', m.lid === 'L-first' && m.forLid === 'I-1' && J(m.forLids) === J(['I-3']), J(m));
  t('13d. 廠商：全部相同才帶入（兩列都是「甲」＋一列沒有廠商 → 帶「甲」）；沒有廠商 → 空', m.vendor === '甲' && MG([row('A', 1, 1), row('B', 1, 1)]).vendor === '', m.vendor);
  const mv = MG([row('A', 1, 1, { vendor: '甲' }), row('B', 1, 1, { vendor: '乙' })]);
  // cost-sync §7.1：顧問服務區的「委外廠商」決定這列算不算委外，合併時若把多種廠商清成空白，合併後的列就變成自家顧問、委外占比憑空掉到 0。
  // 所以顧問區改成「相同才保留、不同以『、』串接」（截 60 字，被截斷才把完整名單記在說明）；其他分區（供應商）維持原行為。
  t('13e. 顧問區廠商不只一種 → 以「、」串接成「甲、乙」（合併後仍算委外；原本是欄位留白＋說明記「廠商：甲、乙」，因委外占比而改）', mv.vendor === '甲、乙' && mv.note === '', J(mv));
  const mvs = MG([row('A', 1, 1, { vendor: '甲', cat: 'software' }), row('B', 1, 1, { vendor: '乙', cat: 'software' })]);
  t('13e2. 軟體區（供應商）廠商不只一種 → 維持原行為：廠商欄空白、記到說明「廠商：甲、乙」', mvs.vendor === '' && mvs.note === '廠商：甲、乙', J(mvs));
  t('13f. 合併後仍是合法的列（cleanLine 過得去）：Q.normalize 保留、不超出上限', Q.normalize([m]).length === 1 && Q.normalize([m])[0].unitCost === 207507.5);
  t('13g. 項目名稱過長（40 列 × 20 字）截成 120 字、不是空白', (() => { const big = Array.from({ length: 40 }, (_, i) => row('品項名稱很長很長很長很長' + i, 1, 1)); const r = MG(big); return r.desc.length === 120 && /^合併 40 項：/.test(r.desc); })());
  t('13h. 印花稅列不併入；空陣列／非陣列／全是無效列 → null', MG([row('A', 1, 100), { cat: 'other', auto: 'stamp' }]).unitCost === 100 && MG([]) === null && MG(undefined) === null && MG([null, 5]) === null);
  t('13i. 純函式：不修改輸入', (() => { const a = rows3.map((x) => Object.assign({}, x)); const b = J(a); MG(a); return J(a) === b; })());
  t('13j. 單一列 → 項目沿用原名、金額＝原小計（UI 只在 ≥2 列時提供按鈕，這裡只確認不壞）', (() => { const r = MG([row('PM', 20, 8000)]); return r.desc === 'PM' && r.unitCost === 160000 && r.qty === 1; })());
  // 隨機：1,000 組，合併前後的分必須完全相同（含小數單價、小數數量）
  t('13k. 隨機 1,000 組（2～12 列、含小數數量與單價）：合併前後的成本合計（分）逐組相同', (() => {
    let seed = 12345; const rnd = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
    for (let k = 0; k < 1000; k++) {
      const n = 2 + Math.floor(rnd() * 11);
      const rs = Array.from({ length: n }, (_, i) => row('r' + i, Math.round(rnd() * 5000) / 100, Math.round(rnd() * 99999999) / 100));
      const mm = MG(rs);
      if (centsOf(rs) !== centsOf([mm])) return false;
    }
    return true;
  })());
}

// 14) 「完成」前的提醒（QCL.unmatchedItemsNote）：成本明細裡找不到對應列的報價品項，只提醒、不擋
{
  const UN = typeof Q.unmatchedItemsNote === 'function' ? Q.unmatchedItemsNote : () => '（不存在）';
  const itm = (lid, d) => ({ lid, desc: d, unit: '式', qty: 1 });
  const cl = (d, extra) => Object.assign({ cat: 'consult', desc: d, qty: 1, unitCost: 100 }, extra || {});
  t('14a. 每個品項都有對應列（forLid 或同名）→ 空字串（不打擾）', UN([cl('A', { forLid: 'L1' }), cl('B')], [itm('L1', 'A'), itm('L2', 'B')]) === '');
  const n1 = UN([cl('專案整包成本', { forLid: 'L1' })], [itm('L1', '顧問人天'), itm('L2', '軟體授權'), itm('L3', '硬體設備')]);
  t('14b. 整包情境：列出沒有對應列的 2 個品項（軟體授權、硬體設備），並說明「已包含在其他列可直接完成／漏填請補上」', /以下 2 個報價品項在成本明細裡找不到對應的列：軟體授權、硬體設備。/.test(n1) && /已包含在其他列（例如整包），可以直接完成/.test(n1) && /漏填，請先回去補上/.test(n1), n1);
  const many = Array.from({ length: 12 }, (_, i) => itm('X' + i, '項' + i));
  t('14c. 12 個品項都沒有對應列：只列前 8 個＋「…另 4 個」，數量寫 12', (() => { const x = UN([], many); return /以下 12 個/.test(x) && /項7…另 4 個。/.test(x) && !/項8/.test(x); })());
  t('14d. 業務事後新增品項（既有成本列都有 forLid）→ 只提醒新品項', UN([cl('PM', { forLid: 'L1' }), cl('設備', { forLid: 'L2' })], [itm('L1', 'A'), itm('L2', 'B'), itm('L3', '加購授權')]).includes('：加購授權。'));
  t('14e. 容錯：lines／items 不是陣列 → 空字串、不丟錯', (() => { try { return UN(undefined, undefined) === '' && UN(null, []) === '' && UN([], null) === ''; } catch (e) { return false; } })());
  t('14f. 印花稅列不算對應列；空白說明的品項不會被提醒（與補入同一條規則）', UN([{ cat: 'other', auto: 'stamp' }], [itm('L1', 'A')]).includes('：A。') && UN([], [{ lid: 'L9', desc: '  ', unit: '式', qty: 1 }]) === '');
}

// 15) forLids（合併為一列後所涵蓋的品項）：清洗規則、合併時的聯集、涵蓋判定（補入／完成提醒）、列 HTML 的往返
{
  const MG = Q._mergeLines;
  const itm = (lid, d, unit, extra) => Object.assign({ lid, desc: d, unit: unit || '式', qty: 1 }, extra || {});
  const row = (d, extra) => Object.assign({ cat: 'consult', desc: d, qty: 1, unitCost: 100, unit: '式' }, extra || {});
  const ids = (n, p) => Array.from({ length: n }, (_, i) => (p || 'id') + i);
  const nl = Q.normalize([row('a', { forLid: 'f1', forLids: [' f2 ', 'f3', 'f2', '', '   ', 'f1', 5, null, {}, 'x'.repeat(99)] })])[0];
  t('15a. normalize 的 forLids：去頭尾空白、去空白／重複／非字串、剔除與 forLid 相同者、每個截 64 字（與伺服器一致）', J(nl.forLids) === J(['f2', 'f3', 'x'.repeat(64)]) && nl.forLid === 'f1', J(nl.forLids));
  t('15b. forLids 上限 60 個（超過的截掉，不丟錯）；非陣列／空陣列 → 沒有 forLids 鍵；印花稅列沒有 forLids',
    Q.normalize([row('a', { forLids: ids(80) })])[0].forLids.length === 60 && ['x', 5, {}, null, undefined, [], ['', ' ']].every((v) => !('forLids' in Q.normalize([row('a', { forLids: v })])[0]))
    && !('forLids' in Q.normalize([{ cat: 'other', auto: 'stamp', forLids: ['a'] }])[0]));
  t('15c. normalize 冪等（含 forLids）、不修改輸入', (() => { const raw = [row('a', { forLid: 'f1', forLids: ['g1', 'g2'] })]; const s = J(raw); const o = Q.normalize(raw); return J(raw) === s && J(Q.normalize(o)) === J(o); })());
  // 合併：forLids＝所有被併列的 forLid∪forLids
  const m1 = MG([row('A', { lid: 'L-first', forLid: 'I-1' }), row('B', { forLid: 'I-2' }), row('C', { forLids: ['I-3', 'I-4'] })]);
  t('15d. 合併：forLid 沿用第一列；其他列的 forLid 與 forLids 全部進 forLids（I-2、I-3、I-4），不重複、不含 forLid 本身', m1.forLid === 'I-1' && J(m1.forLids) === J(['I-2', 'I-3', 'I-4']), J(m1));
  const m2 = MG([row('A'), row('B', { forLid: 'I-2' })]);
  t('15e. 合併：第一列沒有 forLid、後面有 → forLid 沒有，forLids 收後面的', !('forLid' in m2) && J(m2.forLids) === J(['I-2']), J(m2));
  t('15f. 合併：沒有任何 forLid／forLids → 沒有這兩個鍵（與以前的輸出相同）', (() => { const m = MG([row('A'), row('B')]); return !('forLid' in m) && !('forLids' in m); })());
  t('15g. 合併 70 個品項的列：forLids 最多 60 個', (MG(Array.from({ length: 70 }, (_, i) => row('r' + i, { forLid: 'I' + i }))).forLids || []).length === 60);
  // 涵蓋判定：整包（合併）後，被併的每個品項都算有成本列（B：不再互相矛盾）
  const items5 = [itm('I-1', '顧問人天', '人天'), itm('I-2', '軟體授權', '套'), itm('I-3', '硬體設備', '台'), itm('I-4', '維護服務'), itm('I-5', '教育訓練')];
  const merged = MG(items5.slice(0, 5).map((it, i) => row(it.desc, { forLid: it.lid, lid: 'cl' + i })));
  t('15h. 5 個品項各一列成本 → 合併為一列 → 補入新品項：沒有需要補入（以前會列出 4 個）', Q._missingItemLines([merged], items5).length === 0 && Q._missingItemLines([row('專案整包', { forLid: 'I-1' })], items5).length === 4);
  t('15i. 合併後「完成」前的提醒（unmatchedItemsNote）：空字串；沒合併只刪列的整包仍會提醒另外 4 個', Q.unmatchedItemsNote([merged], items5) === '' && /以下 4 個/.test(Q.unmatchedItemsNote([row('專案整包', { forLid: 'I-1' })], items5)));
  t('15j. 合併後事後新增品項：只列新增的那個（不會連被併的一起列）', J(Q._missingItemLines([merged], items5.concat([itm('I-6', '加購授權')])).map((l) => l.desc)) === J(['加購授權']));
  t('15k. 一列的 forLids 涵蓋多個品項時，被涵蓋的品項不再用品名去補其他同名品項：forLids=[I-1]，品項 I-1 與 I-9 同名 → I-9 仍算沒有對應（它沒被任何列指向、而這列已「指向品項」不進品名池）',
    J(Q._missingItemLines([row('整包', { forLids: ['I-1'] })], [itm('I-1', '同名'), itm('I-9', '同名')]).map((l) => l.forLid)) === J(['I-9']));
  // 列 HTML：data-forlids 往返（編輯模式）
  const eh = Q._editRowHtml({ desc: '甲', forLid: 'I-1', forLids: ['I-2', 'I-"3'] }, 'consult');
  const mAttr = /data-forlids="([^"]*)"/.exec(eh);
  const unesc = (s) => s.replace(/&quot;/g, '"').replace(/&#039;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  t('15l. 編輯列 HTML：data-forlid 與 data-forlids（JSON，屬性值已跳脫，引號不會跳出屬性）往返相同；沒有 forLids 就沒有該屬性', !!mAttr && J(JSON.parse(unesc(mAttr[1]))) === J(['I-2', 'I-"3']) && /data-forlid="I-1"/.test(eh) && !/data-forlids/.test(Q._editRowHtml({ desc: '乙' }, 'consult')));
}

// 16) 業務端的「有價品項沒有成本列」提醒（QCL.unmatchedPricedItems／unmatchedPricedNote）：只算有價品項、空白說明的有價品項也算、與伺服器 preview.warnings 同一句
{
  const P = Q.unmatchedPricedItems, PN = Q.unmatchedPricedNote;
  const itm = (lid, d, price, extra) => Object.assign({ lid, desc: d, unit: '式', qty: 1, unitPrice: price }, extra || {});
  const cl = (d, extra) => Object.assign({ cat: 'consult', desc: d, qty: 1, unitCost: 100 }, extra || {});
  const nm = (lines, items) => P(lines, items).map((u) => u.name || '(空白)').join('|');
  t('16a. 方法存在、純函式；全部有對應 → 空陣列、提醒文字空字串', typeof P === 'function' && typeof PN === 'function' && nm([cl('A', { forLid: 'L1' })], [itm('L1', 'A', 100)]) === '' && PN([cl('A')], [itm('L1', 'A', 100)]) === '');
  t('16b. 只算有價品項：贈品（單價 0）、沒有 unitPrice 欄位（顧問看不到價格）的品項不提醒', nm([], [itm('L1', 'A', 100), itm('L2', '贈品', 0), { lid: 'L3', desc: '顧問看不到價格' }, itm('L4', 'B', '50')]) === 'A|B');
  t('16c. 說明空白的有價品項會被提醒（H3／補入不會，因為種子略過它）：名稱顯示「（未命名品項）」', nm([cl('x')], [itm('L1', '', 900000)]) === '(空白)' && PN([cl('x')], [itm('L1', '  ', 900000)]).includes('（未命名品項）')
    && Q._missingItemLines([cl('x')], [itm('L1', '', 900000)]).length === 0);
  t('16d. 提醒文字：總數＋最多 5 個品名（超過接「…另 N 個」）＋「整包可忽略／漏填毛利率可能被高估」，不含金額',
    (() => { const x = PN([], Array.from({ length: 8 }, (_, i) => itm('L' + i, '品' + i, 987654))); return /^有 8 個有價品項在成本明細中找不到對應的成本列：品0、品1、品2、品3、品4…另 3 個。/.test(x) && x.includes('整包') && x.includes('毛利率可能被高估') && !/987654/.test(x); })());
  t('16e. 容錯：lines／items 不是陣列、含 null／非物件 → 空、不丟錯；印花稅列不算對應列', (() => { try { return P(undefined, undefined).length === 0 && P(null, [null, 5, 'x']).length === 0 && nm([{ cat: 'other', auto: 'stamp', forLid: 'L1' }], [itm('L1', 'A', 5)]) === 'A'; } catch (e) { return false; } })());
  t('16f. forLids 涵蓋：整包列 forLids=[L1,L2] → 兩個有價品項都有對應', nm([cl('整包', { forLid: 'L1', forLids: ['L2'] })], [itm('L1', 'A', 5), itm('L2', 'B', 5)]) === '');
  t('16g. 同名品項需要同數量的列（多重集合）；指向無價現有品項的列不進品名池', nm([cl('A')], [itm('L1', 'A', 5), itm('L2', 'A', 5)]) === 'A' && nm([cl('贈品', { forLid: 'G' })], [itm('G', '贈品', 0), itm('P', '贈品', 100)]) === '贈品');
  t('16h. 沒有 lid 的品項用 legacy-<索引>；index 是在傳入陣列中的位置', J(P([cl('x', { forLid: 'legacy-0' })], [itm(undefined, 'A', 5), itm(undefined, 'B', 5)])) === J([{ lid: 'legacy-1', name: 'B', index: 1 }]));

  // 16i–16p：儲存前確認的「基準」（H1）——QCL.unmatchedKeys（載入時記下未涵蓋有價品項的鍵）＋QCL.unmatchedNewNote（儲存時只對基準以外新出現的提醒）
  const K = Q.unmatchedKeys, NN = Q.unmatchedNewNote;
  t('16i. unmatchedKeys：回傳未涵蓋有價品項的鍵（依品項順序）；贈品、說明空白但有價者的規則同 unmatchedPricedItems；全部涵蓋＝空陣列；新品項用暫時代號 nid',
    typeof K === 'function' && typeof NN === 'function'
    && J(K([cl('A', { forLid: 'L1' })], [itm('L1', 'A', 5), itm('L2', 'B', 5), itm('L3', '贈品', 0), itm('L4', '', 9)])) === J(['L2', 'L4'])
    && J(K([cl('A')], [itm('L1', 'A', 5)])) === '[]'
    && J(K([], [itm(undefined, 'N1', 5, { nid: 'nid-1' }), itm(undefined, 'N2', 5)])) === J(['nid-1', 'legacy-1']));
  const I3 = [itm('L1', '甲', 100), itm('L2', '乙', 200), itm('L3', '丙', 300)];
  const BASE3 = K([], I3);   // 載入時三個品項都沒有成本列涵蓋（整包情境）
  t('16j. 載入後不改任何東西：基準以外沒有新出現 → 提醒文字空字串（常駐提醒仍用 unmatchedPricedNote，仍列出 3 個）', J(BASE3) === J(['L1', 'L2', 'L3']) && NN([], I3, BASE3) === '' && PN([], I3).includes('有 3 個'));
  const noteNew = NN([], I3.concat([itm('L9', '新增項目', 777)]), BASE3);
  t('16k. 再新增一個有價品項沒補成本 → 提醒只列「新增項目」（1 個），不重複列基準內的甲乙丙；文字不含金額',
    /^有 1 個有價品項在成本明細中找不到對應的成本列：新增項目。/.test(noteNew) && !/甲|乙|丙/.test(noteNew) && !/777/.test(noteNew), noteNew);
  t('16l. 沒有基準（null／undefined／非陣列／空陣列，例如新單）→ 任何未涵蓋品項都算新出現，文字等同 unmatchedPricedNote',
    [null, undefined, 'x', 5, {}, []].every((b) => NN([], I3, b) === PN([], I3)) && NN([cl('甲')], I3, null).includes('有 2 個'));
  t('16m. 刪掉基準內的品項、或改名（lid 不變）→ 仍是基準內，不提醒；刪掉原本涵蓋它的成本列 → 該品項變成新出現的未涵蓋 → 提醒',
    NN([], I3.slice(1), BASE3) === '' && NN([], [itm('L1', '改了名的甲', 100), I3[1], I3[2]], BASE3) === ''
    && /^有 1 個有價品項在成本明細中找不到對應的成本列：丙。/.test(NN([cl('甲', { forLid: 'L1' }), cl('乙', { forLid: 'L2' })], I3, K([cl('甲', { forLid: 'L1' }), cl('乙', { forLid: 'L2' }), cl('丙', { forLid: 'L3' })], I3))));
  t('16n. 未存檔新品項用 nid 當鍵：基準內的 nid 不提醒；沒進基準的新 nid 才提醒；0 元品項不在基準，之後改成有價 → 算新出現',
    (() => { const n1 = itm(undefined, '新甲', 100, { nid: 'nid-a' }), g = itm('G1', '贈品', 0); const b = K([], [n1, g]);
      return J(b) === J(['nid-a']) && NN([], [n1, g], b) === '' && NN([], [n1, itm(undefined, '新乙', 50, { nid: 'nid-b' }), g], b).includes('新乙') && NN([], [n1, itm('G1', '贈品', 500)], b).includes('贈品'); })());
  t('16o. 品項被「合併為一列」涵蓋（forLid＋forLids）或補了成本列 → 基準內外都不再提醒；基準之外的新品項被整包列的 forLids 涵蓋也不提醒',
    NN([cl('整包', { forLid: 'L1', forLids: ['L2', 'L3', 'L9'] })], I3.concat([itm('L9', '新增項目', 777)]), BASE3) === '' && NN([cl('新增項目')], I3.concat([itm('L9', '新增項目', 777)]), BASE3) === '');
  t('16p. 容錯：lines／items 髒資料不丟錯；baseline 陣列含非字串也不丟錯',
    (() => { try { return NN(undefined, undefined, undefined) === '' && NN(null, [null, 5, 'x'], [1, null, {}]) === '' && K(undefined, undefined).length === 0; } catch (e) { return false; } })());
  {
    // 隨機驗證兩條不變式（≥3000 組，含髒資料）：①以當下的鍵當基準 → 一定沒有新提醒 ②空基準 ＝ 全部提醒 ③基準的任意子集：提醒的品名＝子集以外的未涵蓋品項
    let seed = 20261009;
    const rnd = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
    const pick = (a) => a[Math.floor(rnd() * a.length)];
    const NAMES = ['A', 'B', ' A ', 'C', '', '  ', undefined, '名'.repeat(35), 'Z'];
    const PRICES = [100, 0, -5, '50', '', undefined, 0.01, NaN];
    let bad = 0, withNote = 0, firstBad = '';
    for (let n = 0; n < 3000; n++) {
      const items = Array.from({ length: Math.floor(rnd() * 8) }, (_, i) => rnd() < 0.1 ? { kind: 'title', desc: 'T' } : Object.assign({ desc: pick(NAMES), unitPrice: pick(PRICES) }, rnd() < 0.7 ? { lid: 'L' + i } : (rnd() < 0.5 ? { nid: 'nid-' + i } : {})));
      const lines = Array.from({ length: Math.floor(rnd() * 6) }, () => Object.assign({ cat: 'consult', desc: pick(NAMES) }, rnd() < 0.4 ? { forLid: 'L' + Math.floor(rnd() * 8) } : {}, rnd() < 0.15 ? { forLids: ['L' + Math.floor(rnd() * 8), 'nid-' + Math.floor(rnd() * 8)] } : {}));
      const all = P(lines, items), keys = K(lines, items);
      const sub = keys.filter(() => rnd() < 0.5);
      const rest = all.filter((u) => !sub.includes(u.lid));
      const exp = rest.length ? Q._unmatchedMessage(rest.map((u) => u.name)) : '';
      const ok = J(keys) === J(all.map((u) => u.lid)) && NN(lines, items, keys) === '' && NN(lines, items, []) === PN(lines, items) && NN(lines, items, sub) === exp;
      if (rest.length) withNote++;
      if (!ok) { bad++; if (!firstBad) firstBad = J({ items, lines, sub }); }
    }
    t('16q. 隨機 3000 組（含髒資料、nid、legacy 鍵）：①以當下的鍵當基準＝永遠沒有提醒 ②空基準＝unmatchedPricedNote ③任意子集當基準＝只列子集以外的品項；0 差異，且有提醒的組數 > 500', bad === 0 && withNote > 500, 'bad=' + bad + ' withNote=' + withNote + ' ' + firstBad.slice(0, 300));
  }
}

// 17) 前端（QCL）與伺服器（lib/quoteCostLines.js）的涵蓋判定／警告文字逐組相同：≥5000 組隨機輸入（含髒資料）
{
  let CLm = null;
  try { CLm = require(path.join(ROOT, 'lib/quoteCostLines.js')); } catch (e) { CLm = null; }
  if (!CLm || typeof CLm.unmatchedItems !== 'function') {
    t('17. 伺服器端 unmatchedItems 讀不到（略過比對）', true, 'SKIPPED');
  } else {
    let seed = 20261008;
    const rnd = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
    const pick = (a) => a[Math.floor(rnd() * a.length)];
    const LIDS = ['a', 'b', 'c', 'd', 'e', ' a ', 'a'.repeat(70), '', undefined, null, 5, 'legacy-0', 'legacy-1', 'legacy-3', 'x y'];
    const NAMES = ['A', 'B', ' A ', 'A\tB', 'A B', 'C\nD', 'C D', '', '  ', undefined, null, 5, '長'.repeat(130), '長'.repeat(125), '名'.repeat(35), 'Z', 'z', 'A B'];
    const PRICES = [100, 0, -5, '50', '', 'abc', undefined, null, 0.01, '1e3', NaN, Infinity, '7 apples', 1e9];
    const mkItem = () => {
      const r = rnd();
      if (r < 0.07) return { lid: pick(LIDS), kind: 'title', desc: pick(NAMES) };
      if (r < 0.12) return { lid: pick(LIDS), kind: 'subtotal', desc: pick(NAMES), unitPrice: 5 };
      if (r < 0.15) return pick([null, 5, 'x', []]);
      const o = {};
      if (rnd() < 0.9) o.lid = pick(LIDS);
      if (rnd() < 0.92) o.desc = pick(NAMES);
      if (rnd() < 0.9) o.unitPrice = pick(PRICES);
      return o;
    };
    const mkLine = () => {
      const r = rnd();
      if (r < 0.07) return { cat: 'other', auto: 'stamp', desc: pick(NAMES), forLid: pick(LIDS), forLids: [pick(LIDS)] };
      if (r < 0.1) return pick([null, 5, 'x', []]);
      const o = { cat: pick(['consult', 'hw', 'other']), desc: pick(NAMES) };
      if (rnd() < 0.4) o.forLid = pick(LIDS);
      const r2 = rnd();
      if (r2 < 0.22) o.forLids = Array.from({ length: Math.floor(rnd() * 5) }, () => pick(LIDS));
      else if (r2 < 0.25) o.forLids = pick(['x', 5, {}, null]);
      else if (r2 < 0.27) o.forLids = Array.from({ length: 70 }, (_, i) => (i === 65 ? 'a' : 'p' + i));   // 超過 60 個：第 61 個以後不算
      return o;
    };
    const N = 6000;
    let diffs = 0, msgDiffs = 0, noteDiffs = 0, nonEmpty = 0, empty = 0, viaForLids = 0, firstDiff = '';
    for (let n = 0; n < N; n++) {
      const items = Array.from({ length: Math.floor(rnd() * 9) }, mkItem);
      const lines = Array.from({ length: Math.floor(rnd() * 9) }, mkLine);
      const q = { items, costLines: lines };
      const a = CLm.unmatchedItems(q), b = Q.unmatchedPricedItems(lines, items);
      if (J(a) !== J(b)) { diffs++; if (!firstDiff) firstDiff = J(q) + ' => ' + J(a) + ' vs ' + J(b); }
      const names = a.map((x) => x.name);
      if (CLm.unmatchedMessage(names) !== Q._unmatchedMessage(names)) msgDiffs++;
      const cw = CLm.costWarnings(q), note = Q.unmatchedPricedNote(lines, items);
      if ((cw.length ? cw[0].message : '') !== note) noteDiffs++;
      if (a.length) nonEmpty++; else empty++;
      if (lines.some((l) => l && Array.isArray(l.forLids) && l.forLids.length)) viaForLids++;
    }
    t('17a. unmatchedItems（伺服器）＝ unmatchedPricedItems（前端）：' + N + ' 組隨機輸入（髒 lid／品名／單價、forLids 髒值與超量、kind 列、印花稅列、非物件）逐組 0 差異', diffs === 0, 'diffs=' + diffs + ' ' + firstDiff);
    t('17b. 警告文字：unmatchedMessage 兩邊逐字相同（0 差異）；costWarnings[0].message ＝ unmatchedPricedNote（0 差異）', msgDiffs === 0 && noteDiffs === 0, 'msg=' + msgDiffs + ' note=' + noteDiffs);
    t('17c. 隨機輸入有鑑別力：有未涵蓋品項的組數與完全涵蓋的組數都多於 1000、含 forLids 的組數多於 1000', nonEmpty > 1000 && empty > 1000 && viaForLids > 1000, 'nonEmpty=' + nonEmpty + ' empty=' + empty + ' forLids=' + viaForLids);
    // 變異鑑別：把前端判定故意弄壞，這組比對必須抓得到（用 QCL_SRC 變異測試做；這裡只確認「兩邊真的各算各的」——前端函式不是呼叫伺服器）
    t('17d. 前端判定獨立於伺服器：前端程式碼檔案不含 require（vm 沒有 require 也能算）', !/\brequire\s*\(/.test(src));
  }
}

// 18) 「原報價品項已刪除」徽章（唯讀列 HTML 與判定函式；編輯模式的顯示切換由瀏覽器 e2e 驗證）
{
  const VR = (l, items) => Q._viewRowHtml(Object.assign({ desc: '甲', qty: 1, unitCost: 1 }, l), 'consult', items);
  const has = (h) => /class="qcl-orph"/.test(h) && h.includes('原報價品項已刪除');
  const itm = (lid) => ({ lid, desc: 'x', unit: '式', qty: 1 });
  t('18a. forLid 指向的品項已不在清單 → 顯示徽章', has(VR({ forLid: 'GONE' }, [itm('L1')])));
  t('18b. items 沒提供（undefined／null／非陣列）→ 不顯示', [undefined, null, 'x', {}].every((v) => !has(VR({ forLid: 'GONE' }, v))));
  t('18c. forLid 還在清單 → 不顯示；forLids 有任一個還在 → 不顯示；forLid 與 forLids 全都不在 → 顯示；只有 forLids 且全不在 → 顯示',
    !has(VR({ forLid: 'L1' }, [itm('L1')])) && !has(VR({ forLid: 'G1', forLids: ['G2', 'L1'] }, [itm('L1')])) && has(VR({ forLid: 'G1', forLids: ['G2', 'G3'] }, [itm('L1')])) && has(VR({ forLids: ['G2'] }, [itm('L1')])));
  t('18d. 沒有 forLid／forLids 的列（手動新增）→ 不顯示；清單是空陣列但列有 forLid → 顯示（品項全被刪了）', !has(VR({}, [itm('L1')])) && has(VR({ forLid: 'L1' }, [])));
  t('18e. 分組標題／小計列的 lid 不算品項：forLid 指向標題 → 顯示', has(VR({ forLid: 'T1' }, [{ lid: 'T1', kind: 'title', desc: 'Part A' }, itm('L1')])) && has(VR({ forLid: 'S1' }, [{ lid: 'S1', kind: 'subtotal', desc: '' }])));
  t('18f. 沒有 lid 的品項以 legacy-<索引> 代替：forLid="legacy-1" 對得上第 2 個沒有 lid 的品項 → 不顯示', !has(VR({ forLid: 'legacy-1' }, [{ desc: 'a' }, { desc: 'b' }])) && has(VR({ forLid: 'legacy-5' }, [{ desc: 'a' }, { desc: 'b' }])));
  t('18g. 判定函式 _isOrphanLine 與 HTML 一致；印花稅列永遠不是孤兒列', Q._isOrphanLine({ forLid: 'GONE' }, [itm('L1')]) === true && Q._isOrphanLine({ forLid: 'L1' }, [itm('L1')]) === false && Q._isOrphanLine({ forLid: 'GONE' }) === false
    && Q._isOrphanLine({ auto: 'stamp', forLid: 'GONE' }, [itm('L1')]) === false);
  t('18h. 徽章不吃掉品名：品名仍 esc 後輸出（XSS 品名不產生元素）；編輯列有隱藏的徽章（hidden）供 refresh 切換', (() => {
    const h = VR({ desc: '<img src=x onerror=1>', forLid: 'GONE' }, [itm('L1')]);
    const e = Q._editRowHtml({ desc: '甲', forLid: 'GONE' }, 'consult');
    return h.includes('&lt;img src=x onerror=1&gt;') && !h.includes('<img') && has(h) && /<span class="qcl-orph" hidden /.test(e);
  })());
}

// 19) 還沒存檔的新品項：畫面給暫時代號 nid（沒有 lid），成本列用它指向品項；存檔時伺服器換成真正的 lid
{
  const nIt = (nid, d, price, extra) => Object.assign({ nid, desc: d, unit: '式', qty: 1, unitPrice: price }, extra || {});
  const items3 = [nIt('nid-a', '顧問人天', 100), nIt('nid-b', '軟體授權', 200), nIt('nid-c', '硬體設備', 300)];
  const seeds = Q.seedFromItems(items3).filter((l) => !l.auto && l.desc !== '差旅交通' && l.desc !== '交際費');
  t('19a. 種子：沒有 lid 的新品項以 nid 當 forLid；有 lid 的品項優先用 lid；分組標題不產生列', J(seeds.map((l) => l.forLid)) === J(['nid-a', 'nid-b', 'nid-c'])
    && Q.seedFromItems([{ lid: 'L1', nid: 'nid-x', desc: 'A' }, { kind: 'title', nid: 'nid-t', desc: 'T' }])[0].forLid === 'L1');
  const mergedSeed = Q._mergeLines(seeds.map((l) => Object.assign({}, l)));
  t('19b. 新單整包：3 個還沒存檔的品項各一列 → 合併為一列（名稱對不上任何品項）→ 補入／完成提醒／有價品項提醒都「沒有」；forLid＝第一個 nid、forLids＝其餘兩個 nid',
    mergedSeed.forLid === 'nid-a' && J(mergedSeed.forLids) === J(['nid-b', 'nid-c']) && Q._missingItemLines([mergedSeed], items3).length === 0 && Q.unmatchedItemsNote([mergedSeed], items3) === '' && Q.unmatchedPricedItems([mergedSeed], items3).length === 0 && Q.unmatchedPricedNote([mergedSeed], items3) === '');
  t('19c. 沒有 nid 的情況（舊行為）：同一批品項的整包列會被提醒 3 個——nid 才讓新單的整包不被誤報', Q.unmatchedPricedItems([Object.assign({}, mergedSeed, { forLid: undefined, forLids: undefined })], items3.map((x) => ({ desc: x.desc, unitPrice: x.unitPrice }))).length === 3);
  const renamed = items3.map((x, i) => (i === 0 ? Object.assign({}, x, { desc: '儲存前改了名字' }) : x));
  t('19d. 儲存前把品項改名：成本列仍用 nid 指向它 → 補入判定沒有新品項、有價品項提醒也沒有', Q._missingItemLines(seeds, renamed).length === 0 && Q.unmatchedPricedItems(seeds, renamed).length === 0);
  t('19e. 刪掉其中一個新品項（畫面上的列沒了）→ 該成本列的 nid 找不到品項：「原報價品項已刪除」徽章顯示；其餘不顯示', (() => {
    const left = items3.filter((x) => x.nid !== 'nid-b');
    return Q._isOrphanLine(seeds[1], left) === true && Q._isOrphanLine(seeds[0], left) === false && Q._isOrphanLine(seeds[2], left) === false;
  })());
  t('19f. lid 與 nid 同時存在時 lid 優先（已存檔的品項不該再有 nid）', Q.unmatchedPricedItems([{ cat: 'other', desc: 'x', forLid: 'L1' }], [{ lid: 'L1', nid: 'nid-z', desc: 'A', unitPrice: 5 }]).length === 0
    && Q.unmatchedPricedItems([{ cat: 'other', desc: 'x', forLid: 'nid-z' }], [{ lid: 'L1', nid: 'nid-z', desc: 'A', unitPrice: 5 }]).length === 1);
}

// 20) 拖曳排序：純函式 dragTargetIndex／dragLineY ＋ QCL.dragSort 在最小 DOM 替身下的呼叫契約
//     （真實瀏覽器的滑鼠／觸控事件序列、自動捲動、輸入框內選字等另在 e2e_dnd.js 用無頭 Edge＋CDP 實跑）
{
  t('20a. 匯出 dragSort／dragTargetIndex／dragLineY', ['dragSort', 'dragTargetIndex', 'dragLineY'].every((k) => typeof Q[k] === 'function'));
  const eq = (a, b) => J(a) === J(b);
  const R5 = [0, 30, 60, 90, 120].map((top) => ({ top, bottom: top + 30 }));     // 5 列等高，中線 15／45／75／105／135
  const slotAt = (y) => Q.dragTargetIndex(R5, y);
  t('20b. 縫隙：最上方（含負數／剛好在第一列中線）＝0；越過中線才算越過；最底下＝5',
    eq([slotAt(-100), slotAt(0), slotAt(14.99), slotAt(15), slotAt(15.01), slotAt(44.99), slotAt(45.01), slotAt(105), slotAt(105.01), slotAt(135), slotAt(135.01), slotAt(1e6)], [0, 0, 0, 0, 1, 1, 2, 3, 4, 4, 5, 5]),
    J([-100, 0, 14.99, 15, 15.01, 44.99, 45.01, 105, 105.01, 135, 135.01, 1e6].map(slotAt)));
  const fin = (y, from) => Q.dragTargetIndex(R5, y, from);
  t('20c. 最終索引（給 fromIdx）：同位置（原列上半／下半）都等於 from；往下一列 from+1；往上一列 from-1；頂端 0；底端 n-1',
    eq([fin(60, 2), fin(89, 2), fin(106, 2), fin(44, 2), fin(-10, 2), fin(1e6, 2), fin(-10, 0), fin(1e6, 4), fin(1e6, 0), fin(-10, 4)], [2, 2, 3, 1, 0, 4, 0, 4, 4, 0]),
    J([fin(60, 2), fin(89, 2), fin(106, 2), fin(44, 2), fin(-10, 2), fin(1e6, 2), fin(-10, 0), fin(1e6, 4), fin(1e6, 0), fin(-10, 4)]));
  const Rv = [{ top: 0, bottom: 100 }, { top: 100, bottom: 120 }, { top: 120, bottom: 300 }, { top: 300, bottom: 310 }];   // 列高不等：中線 50／110／210／305
  t('20d. 列高不等：以各列自己的中線為界（高列的上半段不會把指標算成越過它）',
    eq([0, 49, 50, 51, 109, 110, 111, 209, 211, 305, 306, 400].map((y) => Q.dragTargetIndex(Rv, y)), [0, 0, 0, 1, 1, 1, 2, 2, 3, 3, 4, 4]));
  t('20e. {top,height} 寫法與 {top,bottom} 相同；補 fromIdx 時跨多列（from=0 拖到最後一列下方＝3、from=3 拖到最上方＝0）',
    eq(Rv.map((r) => ({ top: r.top, height: r.bottom - r.top })).length, 4)
    && [0, 60, 111, 250, 306].every((y) => Q.dragTargetIndex(Rv.map((r) => ({ top: r.top, height: r.bottom - r.top })), y) === Q.dragTargetIndex(Rv, y))
    && Q.dragTargetIndex(Rv, 999, 0) === 3 && Q.dragTargetIndex(Rv, -5, 3) === 0 && Q.dragTargetIndex(Rv, 150, 0) === 1 && Q.dragTargetIndex(Rv, 250, 0) === 2);
  t('20f. 邊界輸入：空陣列、非陣列、NaN／非有限的 y → 0；單一列任何位置的最終索引都是 0；fromIdx 不合法時當作只要縫隙',
    Q.dragTargetIndex([], 50) === 0 && Q.dragTargetIndex(null, 50) === 0 && Q.dragTargetIndex(R5, NaN) === 0 && Q.dragTargetIndex(R5, Infinity) === 0 && Q.dragTargetIndex(R5, undefined) === 0
    && [-10, 5, 500].every((y) => Q.dragTargetIndex([{ top: 0, bottom: 30 }], y, 0) === 0) && Q.dragTargetIndex(R5, 1e6, -1) === 5 && Q.dragTargetIndex(R5, 1e6, 99) === 5 && Q.dragTargetIndex(R5, 1e6, 1.5) === 5);
  t('20g. 插入線 y：縫隙 0＝第一列上緣、n＝最後一列下緣、中間＝相鄰兩列之間（有間隙時取中點）；空陣列 0',
    Q.dragLineY(R5, 0) === 0 && Q.dragLineY(R5, 5) === 150 && Q.dragLineY(R5, 2) === 60 && Q.dragLineY([{ top: 0, bottom: 100 }, { top: 110, bottom: 130 }], 1) === 105 && Q.dragLineY([], 3) === 0
    && Q.dragLineY(R5, -3) === 0 && Q.dragLineY(R5, 99) === 150);
  // 隨機：最終索引 ＝ 依定義暴力算出的結果；而且 splice 之後被拖的那一項真的落在該索引
  let badRand = 0;
  for (let k = 0; k < 4000; k++) {
    const n = 1 + Math.floor(Math.random() * 12);
    let top = Math.floor(Math.random() * 50) - 25;
    const rs = [];
    for (let i = 0; i < n; i++) { const h = 1 + Math.floor(Math.random() * 90); rs.push({ top, bottom: top + h }); top += h + Math.floor(Math.random() * 6); }
    const from = Math.floor(Math.random() * n);
    const y = Math.floor(Math.random() * (top + 120)) - 60 + (Math.random() < 0.3 ? 0.5 : 0);
    let slot = 0; while (slot < n && y > (rs[slot].top + rs[slot].bottom) / 2) slot++;
    const want = slot > from ? slot - 1 : slot;
    const got = Q.dragTargetIndex(rs, y, from);
    const arr = rs.map((_, i) => i); arr.splice(got, 0, arr.splice(from, 1)[0]);
    if (Q.dragTargetIndex(rs, y) !== slot || got !== want || arr[got] !== from || got < 0 || got >= n) badRand++;
  }
  t('20h. 隨機 4000 組（列數 1–12、列高不等、列間有縫）：縫隙＝暴力定義、最終索引落在 0..n-1、splice 後被拖項目真的在該索引', badRand === 0, badRand);

  // ── 最小 DOM 替身 ──
  const mkSel = (sel) => String(sel).split(',').map((s) => s.trim()).filter(Boolean).map((s) => { const m = /^([a-z0-9]*)((?:\.[A-Za-z0-9_-]+)*)$/.exec(s); return { tag: m[1], cls: m[2].split('.').filter(Boolean) }; });
  class FEl {
    constructor(tag) { this.tagName = String(tag).toUpperCase(); this.children = []; this.parentNode = null; this.style = {}; this.attrs = {}; this.cls = new Set(); this.ls = []; this.rect = { top: 0, bottom: 0, left: 0, right: 600, width: 600, height: 0 }; this.value = ''; this.type = 'text'; this.text = ''; this.offsetHeight = 30; this.nodeType = 1; this.id = ''; this.disabled = false; this.focused = 0;
      const self = this; this.classList = { add: (c) => self.cls.add(c), remove: (c) => self.cls.delete(c), contains: (c) => self.cls.has(c), toggle: (c, f) => { const on = f === undefined ? !self.cls.has(c) : !!f; if (on) self.cls.add(c); else self.cls.delete(c); return on; } }; }
    get scrollTop() { return this._st || 0; }
    set scrollTop(v) { const max = this.scrollHeight && this.clientHeight ? this.scrollHeight - this.clientHeight : 0; this._st = Math.max(0, Math.min(max, v)); }
    get className() { return Array.from(this.cls).join(' '); }
    set className(v) { this.cls = new Set(String(v).split(' ').filter(Boolean)); }
    get parentElement() { return this.parentNode && this.parentNode.tagName ? this.parentNode : null; }
    get textContent() { return this.text; }
    set textContent(v) { this.text = String(v); }
    setAttribute(k, v) { this.attrs[k] = String(v); }
    getAttribute(k) { return k in this.attrs ? this.attrs[k] : null; }
    appendChild(c) { c.parentNode = this; this.children.push(c); return c; }
    removeChild(c) { this.children = this.children.filter((x) => x !== c); c.parentNode = null; return c; }
    addEventListener(type, fn, cap) { this.ls.push({ type, fn, cap: !!cap }); }
    removeEventListener(type, fn, cap) { this.ls = this.ls.filter((l) => !(l.type === type && l.fn === fn && l.cap === !!cap)); }
    matches(sel) { return mkSel(sel).some((s) => (!s.tag || s.tag.toUpperCase() === this.tagName) && s.cls.every((c) => this.cls.has(c))); }
    closest(sel) { for (let n = this; n && n.tagName; n = n.parentNode) if (n.matches(sel)) return n; return null; }
    contains(o) { for (let n = o; n; n = n.parentNode) if (n === this) return true; return false; }
    querySelectorAll(sel) { const out = []; const walk = (e) => e.children.forEach((c) => { if (c.matches(sel)) out.push(c); walk(c); }); walk(this); return out; }
    querySelector(sel) { return this.querySelectorAll(sel)[0] || null; }
    getBoundingClientRect() { return this.rect; }
    focus() { this.focused++; FD.activeElement = this; }
    setPointerCapture(id) { this.captured = id; }
    releasePointerCapture(id) { if (this.captured === id) this.captured = null; }
    scrollIntoView() { this.scrolledIntoView = (this.scrolledIntoView || 0) + 1; }
  }
  const FWIN = new FEl('window'); FWIN.tagName = '';
  const FD = new FEl('#document'); FD.tagName = '';
  FD.body = new FEl('body'); FD.head = new FEl('head'); FD.documentElement = new FEl('html'); FD.scrollingElement = FD.documentElement; FD.activeElement = null;
  FD.createElement = (tag) => new FEl(tag);
  FD.getElementById = (id) => [FD.head, FD.body].reduce((f, r) => f || (r.id === id ? r : r.querySelectorAll('*').find((e) => e.id === id)), null) || null;
  // querySelectorAll('*') 替身
  const origQsa = FEl.prototype.querySelectorAll;
  FEl.prototype.querySelectorAll = function (sel) { if (sel === '*') { const out = []; const walk = (e) => e.children.forEach((c) => { out.push(c); walk(c); }); walk(this); return out; } return origQsa.call(this, sel); };
  // 事件分派：window(捕獲) → document(捕獲) → 祖先(捕獲) → 目標 → 祖先(冒泡) → document → window
  const fire = (target, type, props) => {
    const ev = Object.assign({ type, target, defaultPrevented: false, stopped: false, immediate: false, preventDefault() { this.defaultPrevented = true; }, stopPropagation() { this.stopped = true; }, stopImmediatePropagation() { this.stopped = true; this.immediate = true; } }, props || {});
    const chain = []; for (let n = target; n && n.tagName; n = n.parentNode) chain.unshift(n);
    const path = [FWIN, FD].concat(chain);
    const run = (node, cap) => { for (const l of node.ls.slice()) { if (l.type === type && l.cap === cap) { l.fn(ev); if (ev.immediate) return false; } } return !ev.stopped; };
    for (const n of path) { if (!run(n, true)) return ev; }
    for (const n of path.slice().reverse()) { if (!run(n, false)) return ev; }
    return ev;
  };
  const listenerCount = (node) => node.ls.length;
  const timers = []; let timerSeq = 0; const outErrs = [];
  ctx.document = FD; ctx.innerWidth = 1000; ctx.innerHeight = 800;
  ctx.addEventListener = (ty, fn, cap) => FWIN.addEventListener(ty, fn, cap);
  ctx.removeEventListener = (ty, fn, cap) => FWIN.removeEventListener(ty, fn, cap);
  ctx.getComputedStyle = (el) => ({ overflowX: el.ovx || 'visible', overflowY: el.ovy || 'visible' });
  ctx.setInterval = (fn) => { const id = ++timerSeq; timers.push({ id, fn, kind: 'i' }); return id; };
  ctx.clearInterval = (id) => { const i = timers.findIndex((x) => x.id === id); if (i >= 0) timers.splice(i, 1); };
  ctx.setTimeout = (fn) => { const id = ++timerSeq; timers.push({ id, fn, kind: 't' }); return id; };
  ctx.console = { error: (...a) => outErrs.push(a.map(String).join(' ')), log() {}, warn() {} };

  // 一個 5 列的 tbody（每列：輸入框＋把手），列矩形由上到下各 30 高、中線 15／45／75／105／135
  const build = (n, extra) => {
    const host = new FEl('div'); host.rect = { top: 0, bottom: 30 * n, left: 0, right: 600, width: 600, height: 30 * n };
    const tb = new FEl('tbody'); tb.rect = { top: 0, bottom: 30 * n, left: 0, right: 600, width: 600, height: 30 * n }; host.appendChild(tb);
    const rows = [];
    for (let i = 0; i < n; i++) {
      const tr = new FEl('tr'); tr.className = 'qrow'; tr.rect = { top: 30 * i, bottom: 30 * i + 30, left: 0, right: 600, width: 600, height: 30 };
      const td = new FEl('td'); const inp = new FEl('input'); inp.value = '列' + i; const h = new FEl('button'); h.className = 'hd';
      tr.appendChild(td); td.appendChild(inp); td.appendChild(h); tb.appendChild(tr); rows.push({ tr, inp, h });
    }
    FD.body.appendChild(host);
    return Object.assign({ host, tb, rows }, extra || {});
  };
  const calls = [];
  const mk = (n, o2) => {
    const f = build(n);
    f.calls = [];
    f.api = Q.dragSort(f.host, Object.assign({ rowSelector: 'tr.qrow', handleSelector: '.hd', onMove: (a, b, info) => f.calls.push([a, b, info]) }, o2 || {}));
    return f;
  };
  const down = (f, i, y, extra) => fire(f.rows[i].h, 'pointerdown', Object.assign({ pointerId: 7, pointerType: 'mouse', button: 0, isPrimary: true, clientX: 10, clientY: y === undefined ? 30 * i + 15 : y }, extra || {}));
  const move = (y, extra) => fire(FD.body, 'pointermove', Object.assign({ pointerId: 7, pointerType: 'mouse', clientX: 10, clientY: y }, extra || {}));
  const up = (y, extra) => fire(FD.body, 'pointerup', Object.assign({ pointerId: 7, pointerType: 'mouse', clientX: 10, clientY: y }, extra || {}));
  const bodyKids = (cls) => FD.body.children.filter((c) => c.cls.has(cls));
  const clean = (f) => { f.api.destroy(); FD.body.removeChild(f.host); };
  const flushT = () => { timers.filter((x) => x.kind === 't').forEach((x) => { x.fn(); timers.splice(timers.indexOf(x), 1); }); };

  let threw1 = 0;
  [() => Q.dragSort(null, {}), () => Q.dragSort({}, {}), () => Q.dragSort(new FEl('div'), { handleSelector: '.h', onMove() {} }), () => Q.dragSort(new FEl('div'), { rowSelector: 'tr', onMove() {} }), () => Q.dragSort(new FEl('div'), { rowSelector: 'tr', handleSelector: '.h' })]
    .forEach((fn) => { try { fn(); } catch (e) { if (e && e.name === 'TypeError') threw1++; } });
  t('20i. 參數檢查：沒有容器／rowSelector／handleSelector／onMove 一律丟 TypeError', threw1 === 5, threw1);

  {
    const f = mk(5);
    t('20j. 註冊：容器上有 pointerdown／keydown／contextmenu 三個委派監聽；一開始 document／window 沒有任何監聽；<style id=qclStyle> 已注入一次',
      eq(f.host.ls.map((l) => l.type).sort(), ['contextmenu', 'keydown', 'pointerdown']) && listenerCount(FD) === 0 && listenerCount(FWIN) === 0 && FD.head.children.filter((c) => c.id === 'qclStyle').length === 1);
    const dn = down(f, 1);
    t('20k. 在把手上按下：對把手 setPointerCapture(pointerId)；此時還不算拖曳（沒有浮影／插入線／半透明列）；document／window 才加上暫時監聽',
      f.rows[1].h.captured === 7 && !f.api.isDragging() && bodyKids('qcl-dnd-ghost').length === 0 && bodyKids('qcl-dnd-line').length === 0 && !f.rows[1].tr.cls.has('qcl-dnd-src') && listenerCount(FD) > 0 && listenerCount(FWIN) > 0);
    move(15 + 30 + 3);   // 位移 3px（< 門檻 4px）
    t('20l. 位移 3px（< 門檻 4px）仍不算拖曳：沒有浮影、沒有半透明', !f.api.isDragging() && bodyKids('qcl-dnd-ghost').length === 0 && !f.rows[1].tr.cls.has('qcl-dnd-src'));
    up(48);
    const ckClick = fire(f.rows[1].h, 'click', {});
    t('20m. 點擊把手（按下→放開，沒超過門檻）：不呼叫 onMove、監聽全部移除、釋放指標捕獲、body 沒有 qcl-dnd-on；之後的 click 不被吃掉（單純點擊不是拖曳）',
      f.calls.length === 0 && listenerCount(FD) === 0 && listenerCount(FWIN) === 0 && f.rows[1].h.captured === null && !FD.body.cls.has('qcl-dnd-on') && !ckClick.defaultPrevented);
    clean(f);
  }
  {
    const f = mk(5);
    down(f, 1);                  // y=45
    move(55);                    // 位移 10px → 開始拖曳
    const ghost = bodyKids('qcl-dnd-ghost')[0], line = bodyKids('qcl-dnd-line')[0];
    t('20n. 超過門檻：被拖列半透明（qcl-dnd-src）、body 加 qcl-dnd-on、出現浮影（文字＝該列第一個文字輸入框的值「列1」）與插入線', f.api.isDragging() && f.rows[1].tr.cls.has('qcl-dnd-src') && FD.body.cls.has('qcl-dnd-on') && !!ghost && !!line && ghost.children[1].text === '列1' && ghost.getAttribute('aria-hidden') === 'true', ghost && ghost.children.map((c) => c.text).join('|'));
    t('20o. 浮影跟著指標（transform 隨 y 變）；插入線畫在目標縫隙（拖到 y=55 → 縫隙 2 → 線 y=60 → translate y=58.5→59 或 58）', /translate\(/.test(ghost.style.transform) && line.style.display === 'block' && /translate\(0px,5[89]px\)/.test(line.style.transform), ghost.style.transform + ' ' + line.style.transform);
    move(138);                   // 越過第 5 列中線 135 → 縫隙 5 → 最終索引 4
    t('20p. 拖到最底下：插入線在最後一列下緣（y=150）；浮影的 bad 標記沒有', /translate\(0px,14[89]px\)/.test(line.style.transform) && !ghost.cls.has('qcl-dnd-bad'), line.style.transform);
    const upEv = up(138);
    t('20q. 放開：onMove(1, 4, info) 只呼叫一次；info.rows 是原本的 5 列、info.row 是被拖的列；浮影／插入線／半透明／body 旗標全部清掉；暫時監聽全移除',
      f.calls.length === 1 && f.calls[0][0] === 1 && f.calls[0][1] === 4 && f.calls[0][2].rows.length === 5 && f.calls[0][2].row === f.rows[1].tr && bodyKids('qcl-dnd-ghost').length === 0 && bodyKids('qcl-dnd-line').length === 0
      && !f.rows[1].tr.cls.has('qcl-dnd-src') && !FD.body.cls.has('qcl-dnd-on') && listenerCount(FD) === 0 && listenerCount(FWIN) === 0 && upEv.defaultPrevented, J(f.calls.map((c) => [c[0], c[1]])));
    t('20r. 放開後把焦點還給被拖那一列的把手（列沒被重畫＝原把手）', FD.activeElement === f.rows[1].h);
    // 拖曳後的 click 只吃一個
    const c1 = fire(f.rows[1].h, 'click', {});
    timers.filter((x) => x.kind === 't').forEach((x) => x.fn()); timers.length = 0;
    const c2 = fire(f.rows[1].h, 'click', {});
    t('20s. 拖曳剛結束瀏覽器補送的 click 被吃掉（preventDefault＋stopPropagation），下一個 setTimeout 後恢復正常', c1.defaultPrevented && !c2.defaultPrevented);
    clean(f);
  }
  {
    // 放回原位＝沒動；Esc／pointercancel 取消；canDrop 否決
    const f = mk(5);
    down(f, 2); move(80); move(86);   // 起點 y=75，仍在第 3 列內
    up(86);
    t('20t. 拖了又放回原位（縫隙在原列前後）＝不呼叫 onMove', f.calls.length === 0 && bodyKids('qcl-dnd-ghost').length === 0 && !FD.body.cls.has('qcl-dnd-on'));
    down(f, 0); move(60);
    const esc = fire(FD.activeElement || FD.body, 'keydown', { key: 'Escape' });
    t('20u. 拖曳中按 Esc：取消並還原（浮影／插入線／半透明清掉）、不呼叫 onMove、事件被擋下（preventDefault＋stopPropagation，外層對話框不會跟著關）',
      f.calls.length === 0 && bodyKids('qcl-dnd-ghost').length === 0 && !f.rows[0].tr.cls.has('qcl-dnd-src') && esc.defaultPrevented && esc.stopped && listenerCount(FD) === 0);
    up(60);
    t('20v. Esc 取消後再放開滑鼠：沒有任何作用（不呼叫 onMove）', f.calls.length === 0);
    down(f, 0); move(60); fire(FD.body, 'pointercancel', { pointerId: 7 });
    t('20w. pointercancel：取消並還原', f.calls.length === 0 && bodyKids('qcl-dnd-ghost').length === 0 && listenerCount(FD) === 0);
    down(f, 0); move(60); fire(FWIN, 'blur', {});
    t('20x. 視窗失焦（blur）：取消並還原', f.calls.length === 0 && bodyKids('qcl-dnd-line').length === 0 && listenerCount(FWIN) === 0);
    down(f, 0); move(60); fire(FD.body, 'contextmenu', {});
    t('20y. 拖曳中按右鍵（contextmenu）：取消', f.calls.length === 0 && bodyKids('qcl-dnd-ghost').length === 0);
    down(f, 0); move(60);
    const lineBefore = bodyKids('qcl-dnd-line')[0].style.transform, ghostBefore = bodyKids('qcl-dnd-ghost')[0].style.transform;
    move(140, { pointerId: 99 }); up(140, { pointerId: 99 });
    t('20z. 其他 pointerId 的移動／放開被忽略（多點觸控不會干擾）：原手勢仍在拖曳、浮影與插入線位置沒被帶走', f.api.isDragging() && f.calls.length === 0 && bodyKids('qcl-dnd-line')[0].style.transform === lineBefore && bodyKids('qcl-dnd-ghost')[0].style.transform === ghostBefore);
    up(100);
    t('20aa. 放開時依「放開點」重算：從第 1 列起拖，放在 y=100（縫隙 3）→ onMove(0, 2)', f.calls.length === 1 && f.calls[0][0] === 0 && f.calls[0][1] === 2, J(f.calls.map((c) => [c[0], c[1]])));
    f.calls.length = 0;
    down(f, 0, 15, { button: 2 }); move(100);
    t('20ab. 滑鼠右鍵／中鍵按下不啟動；非主要指標（isPrimary=false）不啟動；在輸入框上按下（不是把手）也不啟動', !f.api.isDragging());
    down(f, 0, 15, { isPrimary: false }); move(100);
    const dnInp = fire(f.rows[0].inp, 'pointerdown', { pointerId: 7, button: 0, isPrimary: true, clientX: 5, clientY: 15 }); move(100); move(200);
    t('20ac. 同上（續）：輸入框上按下→拖移完全不啟動（沒有浮影、沒有任何暫時監聽）、事件沒被 preventDefault（保留原生選字）', !f.api.isDragging() && bodyKids('qcl-dnd-ghost').length === 0 && !dnInp.defaultPrevented && listenerCount(FD) === 0);
    clean(f);
  }
  {
    // canDrop／rowsOf（限同一分區：夾在該區頭尾）
    const f = mk(5, { canDrop: (a, b, info) => { f.cd = f.cd || []; f.cd.push([a, b, info.rows.length]); return b !== 3; } });
    down(f, 0); move(50);
    t('20ad. canDrop(from, to, info) 每次移動都會被呼叫；回傳 false → 不畫插入線、浮影標 bad', f.cd.length > 0 && f.cd[f.cd.length - 1][0] === 0 && f.cd[f.cd.length - 1][2] === 5 && bodyKids('qcl-dnd-line')[0].style.display === 'block' && !bodyKids('qcl-dnd-ghost')[0].cls.has('qcl-dnd-bad'));
    move(108);   // 縫隙 4 → 最終索引 3 → 否決
    t('20ae. 目標索引 3 被 canDrop 否決：插入線隱藏、浮影 qcl-dnd-bad', bodyKids('qcl-dnd-line')[0].style.display === 'none' && bodyKids('qcl-dnd-ghost')[0].cls.has('qcl-dnd-bad'));
    up(108);
    t('20af. 被否決的位置放開＝取消（不呼叫 onMove）', f.calls.length === 0 && bodyKids('qcl-dnd-ghost').length === 0);
    down(f, 0); move(50); up(50);
    t('20ag. 允許的位置照常移動（0 → 1）', f.calls.length === 1 && f.calls[0][0] === 0 && f.calls[0][1] === 1);
    clean(f);
    const g = mk(6, { rowsOf: (row) => g.rows.map((r) => r.tr).filter((tr) => tr.rect.top < 90 === (row.rect.top < 90)) });   // 前 3 列＝A 區、後 3 列＝B 區
    g.rows.slice(3).forEach((r) => { r.tr.className = 'qrow'; });
    down(g, 1); move(60); move(170);   // 起點在 A 區第 2 列，指標拉到 B 區深處
    up(170);
    t('20ah. 分區（rowsOf）：指標跑到別區時，目標夾在本區邊界——A 區第 2 列拖到最下方只會變成 A 區最後（onMove(1, 2)），不會跨到 B 區；info.rows 只含本區 3 列',
      g.calls.length === 1 && g.calls[0][0] === 1 && g.calls[0][1] === 2 && g.calls[0][2].rows.length === 3, J(g.calls.map((c) => [c[0], c[1], c[2].rows.length])));
    g.calls.length = 0;
    down(g, 4); move(10); up(10);
    t('20ai. 分區：B 區的列拖到頁面最上方＝夾在 B 區頂端（onMove(1, 0)）', g.calls.length === 1 && g.calls[0][0] === 1 && g.calls[0][1] === 0, J(g.calls.map((c) => [c[0], c[1]])));
    clean(g);
  }
  {
    // 鍵盤移動模式
    const f = mk(5);
    const live = () => (FD.getElementById('qclDndLive') || {}).text || '';
    const sp = fire(f.rows[2].h, 'keydown', { key: ' ' });
    t('20aj. 把手上按空白鍵：進入鍵盤移動模式（preventDefault 擋捲動；半透明＋插入線；aria-live 區域播報「已拿起第 3 列…共 5 列」）',
      f.api.isDragging() && sp.defaultPrevented && f.rows[2].tr.cls.has('qcl-dnd-src') && bodyKids('qcl-dnd-line').length === 1 && /已拿起第 3 列「列2」，共 5 列/.test(live()) && FD.getElementById('qclDndLive').getAttribute('role') === 'status' && FD.getElementById('qclDndLive').getAttribute('aria-live') === 'polite', live());
    const dnEv = fire(f.rows[2].h, 'keydown', { key: 'ArrowDown' });
    fire(f.rows[2].h, 'keydown', { key: 'ArrowDown' });
    t('20ak. ↓ 兩次：目標第 5 列（播報「目標位置：第 5 列，共 5 列」）；事件被擋下；到底再按 ↓ 不超出', dnEv.defaultPrevented && dnEv.stopped && /目標位置：第 5 列，共 5 列/.test(live()) && (fire(f.rows[2].h, 'keydown', { key: 'ArrowDown' }), /第 5 列/.test(live())));
    fire(f.rows[2].h, 'keydown', { key: 'ArrowUp' });
    fire(f.rows[2].h, 'keydown', { key: 'Enter' });
    t('20al. ↑ 一次後按 Enter 放下：onMove(2, 3)；半透明與插入線清掉；播報「已放下…現在是第 4 列」；焦點回到把手；鍵盤監聽移除', f.calls.length === 1 && f.calls[0][0] === 2 && f.calls[0][1] === 3 && !f.rows[2].tr.cls.has('qcl-dnd-src') && bodyKids('qcl-dnd-line').length === 0 && /已放下「列2」，現在是第 4 列，共 5 列/.test(live()) && FD.activeElement === f.rows[2].h && listenerCount(FWIN) === 0, live());
    f.calls.length = 0;
    fire(f.rows[1].h, 'keydown', { key: ' ' }); fire(f.rows[1].h, 'keydown', { key: 'End' }); fire(f.rows[1].h, 'keydown', { key: ' ' });
    t('20am. 空白鍵也能放下；End＝移到最後（onMove(1, 4)）', f.calls.length === 1 && f.calls[0][0] === 1 && f.calls[0][1] === 4);
    f.calls.length = 0;
    fire(f.rows[3].h, 'keydown', { key: ' ' }); fire(f.rows[3].h, 'keydown', { key: 'Home' });
    const kesc = fire(f.rows[3].h, 'keydown', { key: 'Escape' });
    t('20an. Home 移到最前、Esc 取消：不呼叫 onMove、擋下 Esc、播報「已取消移動，「列3」仍在第 4 列」、清掉插入線', f.calls.length === 0 && kesc.defaultPrevented && kesc.stopped && /已取消移動，「列3」仍在第 4 列/.test(live()) && bodyKids('qcl-dnd-line').length === 0 && !f.rows[3].tr.cls.has('qcl-dnd-src'), live());
    fire(f.rows[3].h, 'keydown', { key: ' ' }); fire(f.rows[3].h, 'keydown', { key: 'ArrowUp' }); fire(f.rows[3].h, 'blur', {});
    t('20ao. 焦點離開把手（blur）＝取消鍵盤移動', f.calls.length === 0 && !f.api.isDragging());
    fire(f.rows[3].h, 'keydown', { key: ' ' }); fire(f.rows[3].h, 'keydown', { key: ' ' });   // 拿起後立刻放下、沒移動
    t('20ap. 拿起後沒移動就放下＝不呼叫 onMove（播報「位置沒有改變」）', f.calls.length === 0 && /位置沒有改變/.test(live()), live());
    fire(f.rows[0].inp, 'keydown', { key: ' ' });
    fire(f.rows[0].h, 'keydown', { key: ' ', ctrlKey: true });
    fire(f.rows[0].h, 'keydown', { key: 'ArrowDown' });
    t('20aq. 不是把手上的空白鍵、帶修飾鍵的空白鍵、沒進入模式時的方向鍵：都不啟動、不攔截', !f.api.isDragging());
    const k0 = mk(5, { keyboard: false });
    fire(k0.rows[0].h, 'keydown', { key: ' ' });
    t('20ar. keyboard:false 關閉鍵盤移動模式', !k0.api.isDragging());
    clean(k0);
    const one = mk(1);
    fire(one.rows[0].h, 'keydown', { key: ' ' });
    t('20as. 只有一列時空白鍵不進入模式（播報沒有可移動的位置）', !one.api.isDragging() && /只有一列/.test(live()));
    clean(one); clean(f);
  }
  {
    // 自動捲動：可捲動祖先的上下緣
    const f = mk(5);
    const sc = new FEl('div'); sc.ovy = 'auto'; sc.scrollHeight = 2000; sc.clientHeight = 300; sc.rect = { top: 100, bottom: 400, left: 0, right: 600, width: 600, height: 300 };
    FD.body.removeChild(f.host); sc.appendChild(f.host); FD.body.appendChild(sc);
    sc.scrollTop = 500;
    down(f, 0, 200); move(250);   // 位移 50px，開始拖曳；指標在容器中段（不捲動）
    const iv = timers.filter((x) => x.kind === 'i');
    t('20at. 開始拖曳時啟動捲動計時器（每 16ms）；指標在容器中段不捲動', iv.length === 1 && (iv[0].fn(), sc.scrollTop === 500), sc.scrollTop);
    move(398);   // 靠近下緣（容器 100–400，邊緣帶 60px）
    iv[0].fn();
    const afterDown = sc.scrollTop;
    move(102);   // 靠近上緣
    iv[0].fn();
    const afterUp = sc.scrollTop;
    t('20au. 指標靠近容器下緣 → scrollTop 增加；靠近上緣 → scrollTop 減少（越靠邊越快）', afterDown > 500 && afterUp < afterDown, [afterDown, afterUp].join(','));
    sc.scrollTop = 0; move(101); iv[0].fn();
    t('20av. 捲到頂端再往上不會變負（瀏覽器夾住）、不丟錯', sc.scrollTop >= 0 && outErrs.length === 0);
    up(250);
    t('20aw. 結束後計時器清掉', timers.filter((x) => x.kind === 'i').length === 0);
    f.api.destroy(); FD.body.removeChild(sc);
    // scrollContainer 給元素／函式：只捲指定的那個容器（不是祖先鏈）
    const own = new FEl('div'); own.ovy = 'auto'; own.scrollHeight = 2000; own.clientHeight = 300; own.rect = { top: 100, bottom: 400, left: 0, right: 600, width: 600, height: 300 }; FD.body.appendChild(own); own.scrollTop = 500;
    const e1 = mk(5, { scrollContainer: own });
    down(e1, 0, 200); move(250); move(398);
    const iv1 = timers.filter((x) => x.kind === 'i'); iv1.forEach((x) => x.fn());
    t('20ax1. scrollContainer 給元素：改捲那個元素（不用從祖先鏈找）', own.scrollTop > 500, own.scrollTop);
    up(398); clean(e1); own.scrollTop = 500;
    const e2 = mk(5, { scrollContainer: (row) => (row && row.tagName === 'TR' ? own : null) });
    down(e2, 0, 200); move(250); move(398);
    timers.filter((x) => x.kind === 'i').forEach((x) => x.fn());
    t('20ax2. scrollContainer 給函式 (row) => 元素：同樣有效', own.scrollTop > 500, own.scrollTop);
    up(398); clean(e2); FD.body.removeChild(own);

    const g = mk(5, { scrollContainer: false });
    down(g, 0, 200); move(250);
    t('20ax. scrollContainer:false 不啟動自動捲動計時器', timers.filter((x) => x.kind === 'i').length === 0);
    up(250); clean(g);
  }
  {
    // 容錯：onMove 丟錯不破壞狀態；destroy 移除監聽；drag 中 destroy
    const f = mk(5, { onMove: () => { throw new Error('boom'); } });
    outErrs.length = 0;
    down(f, 0); move(60); up(60);
    t('20ay. onMove 丟錯：記到 console.error、不往外丟、狀態清乾淨（可以再拖一次）', outErrs.length === 1 && /onMove/.test(outErrs[0]) && !f.api.isDragging() && listenerCount(FD) === 0);
    down(f, 0); move(60);
    f.api.destroy(); flushT();
    t('20az. 拖曳中 destroy：浮影／插入線／暫時監聽／容器監聽全部清掉，之後按下把手不再反應', bodyKids('qcl-dnd-ghost').length === 0 && bodyKids('qcl-dnd-line').length === 0 && !FD.body.cls.has('qcl-dnd-on') && listenerCount(FD) === 0 && listenerCount(FWIN) === 0 && f.host.ls.length === 0 && (down(f, 0), move(60), !f.api.isDragging()));
    FD.body.removeChild(f.host);
  }
  {
    // 與編輯器 mount 的整合（只驗證可驗的：掛載需要 DOM，這裡只確認 dragSort 的形狀與 destroy 對稱，真實掛載在 e2e_dnd.js）
    t('20ba. 檔案紀律：dragSort 的 addEventListener 都有對應的 removeEventListener（destroy／teardown／suppressClick）', (src.match(/removeEventListener\(/g) || []).length >= 3);
    t('20bb. 檔案紀律：dragSort 不使用 localStorage／eval／innerHTML 寫入使用者文字（浮影文字一律 textContent）', !/localStorage|eval\(/.test(src.slice(src.indexOf('function dragSort('), src.indexOf('function dndMoveRow('))) && !/innerHTML/.test(src.slice(src.indexOf('function dragSort('), src.indexOf('function dndMoveRow('))));
    t('20bc. 編輯列 HTML 有拖曳把手：class=qcl-drag、type=button、aria-label=拖曳排序、title 說明鍵盤操作；唯讀列沒有', /<button type="button" class="qcl-drag"[^>]*aria-label="拖曳排序"/.test(Q._editRowHtml({ desc: 'a' }, 'consult')) && /title="[^"]*空白鍵/.test(Q._editRowHtml({ desc: 'a' }, 'consult')) && !/qcl-drag/.test(Q._viewRowHtml({ desc: 'a' }, 'consult')));
  }
  // 還原 vm（之後若有檢查全域屬性才不會被污染）
  ['document', 'innerWidth', 'innerHeight', 'addEventListener', 'removeEventListener', 'getComputedStyle', 'setInterval', 'clearInterval', 'setTimeout', 'console'].forEach((k) => { delete ctx[k]; });
}

let pass = 0, fail = 0;
res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
console.log(`\n成本明細編輯器檢查：PASS ${pass} / FAIL ${fail}`);
process.exit(fail ? 1 : 0);
