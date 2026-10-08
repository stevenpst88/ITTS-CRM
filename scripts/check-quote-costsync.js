#!/usr/bin/env node
/**
 * 「顧問成本畫面連動業務報價」(cost-sync) ＋ 委外占比 ＋ 顧問姓名欄 的單元／路由層測試。用法：node scripts/check-quote-costsync.js
 * 動 lib/quoteCostLines.js（rel／consultant／applyLinks／normalizeNewItems／outsourced）、lib/quoteApproval.js（hash／簽章／ITEM_PRICE_MISSING）、
 * lib/quoteRoutes.js（PUT /costs 的寫回、cost-draft 端點、顧問可見範圍）、lib/quotePnlExcel.js（B 欄顧問姓名）之後必跑。
 *   1) costLines 新欄位 rel／consultant 的驗證、清洗、輸出
 *   2) 草稿新增的報價項目 normalizeNewItems、對應目標檢查 checkTargets
 *   3) applyLinks（連動計算）：單一連動列、多列加總、單位衝突、nid、forLids、split／none 不動、zero、0.001 下限；對獨立 oracle 的隨機比對（≥6000 組）與不變式
 *   4) 委外占比：isOutsourced／outsourcedCents／outsourcedStats（各種組合與截斷邊界）
 *   5) needPrice 送簽阻擋、結構簽章（新式不含有價位元）、hash／costLinesSig 條件式
 *   6) 路由層（以假 app＋記憶體 db 直接呼叫 handler）：顧問可見範圍矩陣、PUT /costs 草稿不動報價／done 一次寫回、單位衝突、nid 換 lid、sig 不退回、itemChanges、通知、
 *      cost-draft/summary 與 done 之後的結果一致、cost-draft/pnl-preview、權限
 *   7) 舊單位元級相容：5000 張決定性隨機舊式單＋5000 張新式單（沒有任何新欄位）對「改版前」程式的 golden 摘要
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、廠商名、真實費率）。
 */
'use strict';
const path = require('path');
const crypto = require('crypto');
const ROOT = path.join(__dirname, '..');
// 重新產生 golden 摘要時，把環境變數 COSTSYNC_LIB_ROOT 指到「改版前」的檔案樹（內含 lib/），再執行本檔：會印出摘要（不比對）
const LIBROOT = process.env.COSTSYNC_LIB_ROOT || ROOT;
const GOLDEN_MODE = !!process.env.COSTSYNC_LIB_ROOT;
const QA = require(path.join(LIBROOT, 'lib/quoteApproval.js'));
const CL = require(path.join(LIBROOT, 'lib/quoteCostLines.js'));
const registerQuoteRoutes = require(path.join(LIBROOT, 'lib/quoteRoutes.js'));
const PNL = require(path.join(LIBROOT, 'lib/quotePnlExcel.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const L = (cat, desc, qty, unitCost, extra) => Object.assign({ cat, desc, qty, unitCost }, extra || {});
let _n = 0;
const genLid = () => 'g' + (++_n);
const norm = (raw, prev, extra) => CL.normalizeCostLines(raw, Object.assign({ genLid, prev }, extra || {}));
const errCode = (r) => (r.ok ? 'OK' : r.error.code);
function lcg(seed) { let s = seed >>> 0; return () => ((s = (Math.imul(s, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff); }
const deepFreeze = (o) => { if (o && typeof o === 'object' && !Object.isFrozen(o)) { Object.freeze(o); Object.values(o).forEach(deepFreeze); } return o; };

async function run() {
  if (GOLDEN_MODE) { legacyCompat(); return; }   // 重算 golden：只跑第 7 節（改版前的 lib 沒有新函式）
  // ═════════════════ 1) rel／consultant ═════════════════
  let r;
  for (const rel of ['link', 'split', 'none']) {
    r = norm([L('consult', 'PM', 1, 1, rel === 'none' ? { rel } : { rel, forLid: 'a' })]);
    t('1.1 rel=' + rel + ' 合法並原樣輸出', r.ok && r.lines[0].rel === rel);
  }
  for (const rel of [undefined, null, '']) {
    r = norm([L('consult', 'PM', 1, 1, { rel, forLid: 'a' })]);
    t('1.2 rel=' + JSON.stringify(rel) + ' 視為沒送：輸出沒有 rel 鍵（舊行為）', r.ok && !('rel' in r.lines[0]) && r.lines[0].forLid === 'a');
  }
  for (const rel of ['LINK', 'x', 5, true, ['link'], {}, 'link ']) t('1.3 rel=' + JSON.stringify(rel) + ' → 400 BAD_COST_LINE', errCode(norm([L('consult', 'PM', 1, 1, { rel, forLid: 'a' })])) === 'BAD_COST_LINE');
  for (const rel of ['link', 'split']) {
    t('1.4 rel=' + rel + ' 沒有 forLid／forLids → 400 BAD_COST_LINE', errCode(norm([L('consult', 'PM', 1, 1, { rel })])) === 'BAD_COST_LINE' && errCode(norm([L('consult', 'PM', 1, 1, { rel, forLid: '  ', forLids: [] })])) === 'BAD_COST_LINE');
    t('1.5 rel=' + rel + ' 有 forLid 或只有 forLids 都合法', norm([L('consult', 'PM', 1, 1, { rel, forLid: 'a' })]).ok && norm([L('consult', 'PM', 1, 1, { rel, forLids: ['a', 'b'] })]).ok);
  }
  r = norm([L('travel', 'T', 1, 1, { rel: 'none', forLid: 'a', forLids: ['b', 'c'] })]).lines[0];
  t('1.6 rel=none 的 forLid／forLids 一律丟掉（沒有對應目標）', r.rel === 'none' && !('forLid' in r) && !('forLids' in r));
  r = norm([{ cat: 'other', auto: 'stamp', rel: 'link', consultant: 'X', forLid: 'a' }]);
  t('1.7 印花稅列：rel（合法值）、consultant、forLid 都丟掉，不報錯', r.ok && !('rel' in r.lines[0]) && !('consultant' in r.lines[0]) && !('forLid' in r.lines[0]) && r.lines[0].auto === 'stamp');
  t('1.8 印花稅列帶不合法 rel 仍 400（先驗證列舉）', errCode(norm([{ cat: 'other', auto: 'stamp', rel: 'bogus' }])) === 'BAD_COST_LINE');
  // consultant
  r = norm([L('consult', 'PM', 1, 1, { consultant: ' Amy ' })]).lines[0];
  t('1.9 consultant 去頭尾空白；只有顧問服務區保留', r.consultant === 'Amy');
  for (const cat of ['software', 'hw', 'travel', 'other']) t('1.10 ' + cat + ' 區的 consultant 丟掉（不報錯）', (() => { const x = norm([L(cat, 'a', 1, 1, { consultant: 'Amy' })]); return x.ok && !('consultant' in x.lines[0]); })());
  r = norm([L('consult', 'PM', 1, 1, { consultant: '   ' }), L('consult', 'SD', 1, 1)]).lines;
  t('1.11 consultant 空白／沒送 → 不輸出鍵', !('consultant' in r[0]) && !('consultant' in r[1]));
  r = norm([L('consult', 'PM', 1, 1, { consultant: 'x'.repeat(100) })]).lines[0];
  t('1.12 consultant 超過 40 字截斷', r.consultant.length === 40);
  r = norm([L('consult', 'PM', 1, 1, { consultant: 'a\r\nb\tc\u0001d' })]).lines[0];
  t('1.13 consultant 控制字元／換行換成空白', r.consultant === 'a b c d', r.consultant);
  for (const cat of ['consult', 'software', 'travel']) for (const v of [5, ['a'], {}, true]) t('1.14 ' + cat + ' consultant 型別錯 ' + JSON.stringify(v) + ' → 400（與 vendor 一致，不分分類）', errCode(norm([L(cat, 'a', 1, 1, { consultant: v })])) === 'BAD_COST_LINE');
  r = norm([L('consult', 'PM', 2, 3, { vendor: 'V', note: 'n', unit: '人天', forLid: 'a', forLids: ['b'], consultant: 'Amy', rel: 'link' })]).lines[0];
  t('1.15 輸出鍵順序固定：lid,cat,desc,vendor,note,unit,qty,unitCost,forLid,forLids,consultant,rel', eq(Object.keys(r), ['lid', 'cat', 'desc', 'vendor', 'note', 'unit', 'qty', 'unitCost', 'forLid', 'forLids', 'consultant', 'rel']), Object.keys(r).join());
  r = norm([L('consult', 'PM', 2, 3), L('software', 'S', 1, 1, { vendor: 'V' }), { cat: 'other', auto: 'stamp' }]).lines;
  t('1.16 沒用到新欄位的列：鍵集合與以前相同（沒有 consultant／rel）', r.every((l) => !('consultant' in l) && !('rel' in l)) && eq(Object.keys(r[0]), ['lid', 'cat', 'desc', 'vendor', 'note', 'unit', 'qty', 'unitCost']));
  r = norm([L('consult', 'PM', 1, 1, { rel: 'link', forLid: 'a', forLids: ['a', 'b', 'b', ' c '] })]).lines[0];
  t('1.17 link 的 forLids 去重並剔除與 forLid 相同者', eq(r.forLids, ['b', 'c']));
  // ReDoS／CPU：200 萬字元的 consultant／rel 要在 50ms 內處理完
  for (const [k, v] of [['consultant', 'x'.repeat(2e6)], ['rel', 'y'.repeat(2e6)]]) {
    const t0 = process.hrtime.bigint();
    const x = norm([L('consult', 'a', 1, 1, { [k]: v, forLid: 'a' })]);
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    t('1.18 ' + k + ' 200 萬字元：' + ms.toFixed(1) + 'ms（<50ms）；rel 不合法→400，consultant→截斷', ms < 50 && (k === 'rel' ? errCode(x) === 'BAD_COST_LINE' : x.ok && x.lines[0].consultant.length === 40), ms.toFixed(1));
  }
  // publicLines 與稽核文字
  r = CL.publicLines({ costLines: norm([L('consult', 'PM', 1, 100, { consultant: 'Amy', rel: 'link', forLid: 'a' }), L('consult', 'SD', 1, 100), L('software', 'S', 1, 1, { consultant: 'zz' }), { cat: 'other', auto: 'stamp' }]).lines }, { canSeeCost: true, canSeePrice: false, revenueCents: 1000000 });
  t('1.19 publicLines 輸出 consultant／rel（有值才有），沒有時不多鍵', r[0].consultant === 'Amy' && r[0].rel === 'link' && !('consultant' in r[1]) && !('rel' in r[1]) && !('consultant' in r[2]) && !('rel' in r[3]));
  const A1 = norm([L('consult', 'PM', 1, 100, { consultant: 'Amy' })]).lines, B1 = A1.map((l) => Object.assign({}, l, { consultant: 'ConsB', rel: 'split', forLid: 'a' }));
  const sm = CL.summarizeChanges(A1, B1, 10).join('；');
  t('1.20 summarizeChanges 記錄顧問姓名與對應方式異動，且沒有 undefined', /顧問姓名 Amy→ConsB/.test(sm) && /對應 未設定→拆項/.test(sm) && !/undefined/.test(sm), sm);
  { const o1 = norm([L('consult', 'PM', 1, 100)]).lines, o2 = o1.map((l) => Object.assign({}, l, { qty: 2 })); const txt = CL.summarizeChanges(o1, o2).join('；');
    t('1.21 沒動新欄位的列：summarizeChanges 只有數量異動，不提顧問姓名／對應（文字與以前相同）', /數量 1→2/.test(txt) && !/顧問姓名|對應/.test(txt), txt); }
  t('1.22 describeLinesFull：有顧問姓名才多「顧問 X」；沒有時與以前相同', /\(?（顧問 Amy；V）/.test(CL.describeLinesFull(norm([L('consult', 'PM', 1, 1, { consultant: 'Amy', vendor: 'V' })]).lines)) && CL.describeLinesFull(norm([L('consult', 'PM', 1, 1, { vendor: 'V' })]).lines).indexOf('顧問 ') < 0);

  // ═════════════════ 2) newItems／checkTargets ═════════════════
  const ni = (raw, ex) => CL.normalizeNewItems(raw, { existingKeys: ex });
  r = ni([{ nid: 'cn-1', desc: ' 新品項 ', unit: '', qty: '2.5' }, { nid: 'cn_2.x', desc: 'B', unit: '人天人天人天人天人天', qty: 0 }]);
  t('2.1 合法：desc 去空白、unit 空→「式」、unit 截 10、qty 數字字串轉 number、qty 0 允許（草稿）', r.ok && r.items[0].desc === '新品項' && r.items[0].unit === '式' && r.items[0].qty === 2.5 && r.items[1].unit.length === 10 && r.items[1].qty === 0 && eq(Object.keys(r.items[0]), ['nid', 'desc', 'unit', 'qty']));
  t('2.2 空陣列合法；非陣列 → 400 BAD_COST_LINE', ni([]).ok && ni([]).items.length === 0 && [undefined, null, 'x', {}, 5].every((v) => errCode({ ok: ni(v).ok, error: ni(v).error }) === 'BAD_COST_LINE'));
  t('2.3 超過 20 筆 → 400；剛好 20 筆合法', errCode({ ok: ni(Array.from({ length: 21 }, (_, i) => ({ nid: 'n' + i, desc: 'd', qty: 1 }))).ok, error: ni(Array.from({ length: 21 }, (_, i) => ({ nid: 'n' + i, desc: 'd', qty: 1 }))).error }) === 'BAD_COST_LINE' && ni(Array.from({ length: 20 }, (_, i) => ({ nid: 'n' + i, desc: 'd', qty: 1 }))).ok);
  for (const bad of ['', 'a b', 'a/b', 'a<b', 'x'.repeat(65), 5, null, undefined, '中文']) t('2.4 nid ' + String(JSON.stringify(bad)).slice(0, 20) + ' 不合法 → 400', !ni([{ nid: bad, desc: 'd', qty: 1 }]).ok);
  t('2.5 nid 重複 → 400；與既有品項 lid 相同 → 400', !ni([{ nid: 'a', desc: 'd', qty: 1 }, { nid: 'a', desc: 'e', qty: 1 }]).ok && !ni([{ nid: 'L1', desc: 'd', qty: 1 }], ['L1', 'L2']).ok && ni([{ nid: 'L3', desc: 'd', qty: 1 }], new Set(['L1'])).ok);
  for (const bad of ['', '  ', null, undefined, 5, ['x'], {}]) t('2.6 desc ' + JSON.stringify(bad) + ' → 400', !ni([{ nid: 'a', desc: bad, qty: 1 }]).ok);
  for (const bad of [-1, 1e9 + 1, 'abc', '', null, undefined, NaN, Infinity, '1e3', true]) t('2.7 qty ' + String(bad) + ' → 400', !ni([{ nid: 'a', desc: 'd', qty: bad }]).ok);
  t('2.8 desc 超過 120 字截斷；控制字元換空白', ni([{ nid: 'a', desc: 'd'.repeat(300), qty: 1 }]).items[0].desc.length === 120 && ni([{ nid: 'a', desc: 'x\ny', qty: 1 }]).items[0].desc === 'x y');
  t('2.9 列不是物件 → 400', [null, 5, 'x', [], true].every((v) => !ni([v]).ok));
  const tl = norm([L('consult', 'A', 1, 1, { rel: 'link', forLid: 'i1' }), L('consult', 'B', 1, 1, { rel: 'split', forLids: ['i2', 'n1'] }), L('travel', 'C', 1, 1, { rel: 'none' }), L('consult', 'D', 1, 1, { forLid: 'ghost' }), { cat: 'other', auto: 'stamp' }]).lines;
  t('2.10 checkTargets：link／split 的目標都在集合內 → null；rel 缺／none／印花稅列不檢查（舊行為：forLid 指向已刪品項照留）', CL.checkTargets(tl, new Set(['i1', 'i2', 'n1'])) === null);
  const bad2 = CL.checkTargets(tl, ['i1', 'i2']);
  t('2.11 checkTargets：目標不存在 → {index, message}，訊息含列序號與項目名', bad2 && bad2.index === 1 && /第 2 筆「B」/.test(bad2.message), JSON.stringify(bad2));
  t('2.12 checkTargets 對壞資料不丟例外（null 列、非陣列）', CL.checkTargets([null, 5, {}], []) === null && CL.checkTargets('x', []) === null && CL.checkTargets(undefined, undefined) === null);

  // ═════════════════ 3) applyLinks ═════════════════
  const IT = (lid, desc, qty, unit, price) => ({ lid, desc, unit: unit === undefined ? '式' : unit, qty, unitPrice: price === undefined ? 1000 : price, cost: 0 });
  const LK = (forLid, qty, unit, extra) => Object.assign({ lid: 'l' + (++_n), cat: 'consult', desc: 'x', vendor: '', note: '', unit: unit === undefined ? '式' : unit, qty, unitCost: 1, rel: 'link', forLid }, extra || {});
  let items = [IT('a', '甲', 1, '式'), { lid: 't', kind: 'title', desc: 'Part' }, IT('b', '乙', 2, '台'), { lid: 's', kind: 'subtotal', desc: '' }];
  let al = CL.applyLinks(items, [LK('a', 12, '人天')], []);
  t('3.1 單一連動列：數量與單位寫到對應品項（1 式 → 12 人天），其他品項不變；標題／小計列不在結果裡', al.items.length === 2 && al.items[0].lid === 'a' && al.items[0].qty === 12 && al.items[0].unit === '人天' && al.items[0].changed && al.items[1].lid === 'b' && !al.items[1].changed && al.items[1].qty === 2, JSON.stringify(al.items));
  t('3.2 changes：qty 與 unit 各一筆（from→to）', eq(al.changes, [{ lid: 'a', desc: '甲', field: 'qty', from: 1, to: 12 }, { lid: 'a', desc: '甲', field: 'unit', from: '式', to: '人天' }]), JSON.stringify(al.changes));
  al = CL.applyLinks(items, [LK('a', 0.1, '人天'), LK('a', 0.2, '人天'), LK('a', 0.3, '人天')], []);
  t('3.3 多列加總用十進位精確加總：0.1＋0.2＋0.3 ＝ 0.6（不是 0.6000000000000001）', al.items[0].qty === 0.6 && al.items[0].linkCount === 3, String(al.items[0].qty));
  al = CL.applyLinks(items, [LK('a', 3, '人天'), LK('a', 4, '人月')], []);
  t('3.4 同一品項的連動列單位不同 → 單位衝突：品項維持原樣、conflicts 列出單位、changes 不含該品項', al.conflicts.length === 1 && al.conflicts[0].lid === 'a' && eq(al.conflicts[0].units, ['人天', '人月']) && al.items[0].conflict && al.items[0].qty === 1 && al.items[0].unit === '式' && al.changes.length === 0, JSON.stringify(al.conflicts));
  al = CL.applyLinks(items, [LK('a', 3, ' 人天 '), LK('a', 4, '人天')], []);
  t('3.5 單位比對去頭尾空白（" 人天 " 與 "人天" 相同）', al.conflicts.length === 0 && al.items[0].qty === 7 && al.items[0].unit === '人天');
  al = CL.applyLinks(items, [LK('a', 3, ''), LK('a', 4, '式')], []);
  t('3.6 空白單位視為「式」', al.conflicts.length === 0 && al.items[0].unit === '式' && al.items[0].qty === 7);
  al = CL.applyLinks(items, [Object.assign(LK('a', 9, '人天'), { rel: 'split' }), Object.assign(LK('a', 9, '人天'), { rel: 'none', forLid: undefined }), Object.assign(LK('a', 9, '人天'), { rel: undefined }), { cat: 'other', auto: 'stamp', qty: 1, unit: '式', forLid: 'a', rel: 'link' }], []);
  t('3.7 split／none／沒有 rel 的列、印花稅列都不連動：品項完全不變', al.items.every((x) => !x.changed) && al.changes.length === 0);
  al = CL.applyLinks(items, [LK('zzz', 9, '人天')], []);
  t('3.8 連動列指向不存在的品項 → 忽略（不影響任何品項）', al.items.every((x) => !x.changed) && al.changes.length === 0);
  al = CL.applyLinks(items, [LK('t', 9, '人天'), LK('s', 9, '人天')], []);
  t('3.9 連動列指向標題／小計列 → 忽略（標題與小計不是品項）', al.items.every((x) => !x.changed));
  al = CL.applyLinks(items, [Object.assign(LK('a', 5, '人天'), { forLid: undefined, forLids: ['a', 'b'] })], []);
  t('3.10 link 列 forLids 命中多個品項 → 每個命中的品項都加總到（規格字面；前端只會產生單一目標）', al.items[0].qty === 5 && al.items[1].qty === 5 && al.items[0].unit === '人天' && al.items[1].unit === '人天');
  al = CL.applyLinks(items, [Object.assign(LK('a', 5, '人天'), { forLids: ['a', 'a'] })], []);
  t('3.11 同一列的 forLid 與 forLids 重複命中同一品項只算一次', al.items[0].qty === 5 && al.items[0].linkCount === 1);
  // 新品項
  al = CL.applyLinks(items, [LK('n1', 3, '人天'), LK('n1', 2, '人天'), LK('b', 1, '台')], [{ nid: 'n1', desc: '新增項', unit: '式', qty: 99 }, { nid: 'n2', desc: '新增二', unit: '套', qty: 4 }]);
  const n1 = al.items.find((x) => x.nid === 'n1'), n2 = al.items.find((x) => x.nid === 'n2');
  t('3.12 暫時品項被連動列命中 → 用連動結果（5 人天，不是自己填的 99 式）；沒被命中 → 保留自己的 qty／unit', n1 && n1.isNew && n1.qty === 5 && n1.unit === '人天' && n2 && n2.qty === 4 && n2.unit === '套' && !('lid' in n1), JSON.stringify([n1, n2]));
  t('3.13 暫時品項排在既有一般品項之後、每個都有一筆 changes field:"new"（from null、to 數量、帶 unit）', al.items.map((x) => x.lid || x.nid).join() === 'a,b,n1,n2' && al.changes.filter((c) => c.field === 'new').length === 2 && eq(al.changes.find((c) => c.nid === 'n2'), { nid: 'n2', desc: '新增二', field: 'new', from: null, to: 4, unit: '套' }));
  al = CL.applyLinks(items, [LK('a', 0, '人天')], [{ nid: 'n3', desc: 'z', unit: '式', qty: 0 }]);
  t('3.14 連動加總 0 → zero（品項維持原樣、不列入 changes）；沒有連動的暫時品項 qty 0 也列入 zero', al.zero.length === 2 && al.zero.some((z) => z.lid === 'a') && al.zero.some((z) => z.nid === 'n3') && al.items[0].zero && !al.changes.some((c) => c.lid === 'a'));
  al = CL.applyLinks(items, [LK('a', 0.0004, '人天')], []);
  t('3.15 連動加總 >0 但 <0.001 → 取 0.001（與業務存檔的品項數量下限相同）', al.items[0].qty === 0.001 && !al.zero.length);
  al = CL.applyLinks(items, [LK('a', 1, '式')], []);
  t('3.16 連動結果與原值相同（1 式 → 1 式）→ changed=false、沒有 changes', !al.items[0].changed && al.changes.length === 0);
  al = CL.applyLinks([IT('a', '甲', 1, '')], [LK('a', 1, '式')], []);
  t('3.17 原品項單位空白、連動列單位「式」→ 單位算變動（空白≠「式」，寫回後有明確單位）', al.items[0].unitChanged && !al.items[0].qtyChanged);
  // 不改輸入、壞資料不丟例外
  const frozenItems = deepFreeze(JSON.parse(JSON.stringify(items))), frozenLines = deepFreeze([LK('a', 3, '人天'), LK('b', 1, '台')]), frozenNew = deepFreeze([{ nid: 'n1', desc: 'd', unit: '式', qty: 1 }]);
  let threw = false;
  try { CL.applyLinks(frozenItems, frozenLines, frozenNew); CL.materializeItems(frozenItems, CL.applyLinks(frozenItems, frozenLines, frozenNew), (e) => ({ lid: e.nid })); } catch (e) { threw = true; }
  t('3.18 不改動傳入的 items／lines／newItems（全部 deep-freeze 也能跑）', !threw);
  threw = false;
  try { for (const a of [undefined, null, 'x', [null, 5, {}], [{ lid: 5 }]]) for (const b of [undefined, null, [null, 5, {}, { rel: 'link' }]]) for (const c of [undefined, [null, 5, { nid: 5 }, { nid: '' }]]) CL.applyLinks(a, b, c); } catch (e) { threw = true; }
  t('3.19 壞資料（null、非物件、缺欄位）不丟例外', !threw);
  items = [IT('a', '甲', 1, '式'), { lid: 't', kind: 'title', desc: 'P' }, IT('b', '乙', 2, '台')];
  al = CL.applyLinks(items, [LK('a', 12, '人天'), LK('n1', 3, '套')], [{ nid: 'n1', desc: '新', unit: '式', qty: 1 }]);
  const mat = CL.materializeItems(items, al, (e) => ({ lid: 'REAL', desc: e.desc, qty: e.qty, unit: e.unit, unitPrice: 0, needPrice: true }));
  t('3.20 materializeItems：改 qty／unit（其他欄位保留）、標題列原樣、暫時品項接在最後；輸入沒被改', mat.length === 4 && mat[0].qty === 12 && mat[0].unit === '人天' && mat[0].unitPrice === 1000 && mat[1].kind === 'title' && mat[2] === items[2] && mat[3].lid === 'REAL' && mat[3].qty === 3 && items[0].qty === 1);
  const al2 = CL.applyLinks(mat.filter((x) => x.lid !== 'REAL'), [LK('a', 12, '人天')], []);
  t('3.21 冪等：把結果寫回後再算一次 → 沒有任何異動', al2.changes.length === 0 && al2.items.every((x) => !x.changed));
  // 隨機比對獨立 oracle：數量一律是「最多 4 位小數」，oracle 用 ×10000 整數；單位衝突、zero、下限 0.001 都一併比
  {
    const rr = lcg(20261008);
    const pk = (a) => a[Math.floor(rr() * a.length)];
    const UNITS = ['式', '人天', '人月', '套', ' 人天 ', ''];
    let cases = 0, agree = 0, conflictCases = 0, changedCases = 0, zeroCases = 0, newCases = 0, permOk = 0;
    for (let i = 0; i < 6500; i++) {
      const its = Array.from({ length: Math.floor(rr() * 5) + 1 }, (_, k) => (rr() < 0.15 ? { lid: 'T' + k, kind: pk(['title', 'subtotal']), desc: 'x' } : IT('i' + k, 'd' + k, pk([1, 2, 0.5, 10]), pk(UNITS))));
      const gen = its.filter((x) => !x.kind);
      const nws = Array.from({ length: Math.floor(rr() * 3) }, (_, k) => ({ nid: 'n' + k, desc: 'N' + k, unit: pk(UNITS), qty: pk([0, 1, 2.5]) }));
      const targets = gen.map((x) => x.lid).concat(nws.map((x) => x.nid), ['ghost']);
      const lns = Array.from({ length: Math.floor(rr() * 7) }, () => {
        const l = LK(pk(targets), pk([0, 0.1, 0.2, 0.0001, 1, 1.25, 2.5, 7, 0.3333]), pk(UNITS));
        const k = rr();
        if (k < 0.12) l.rel = 'split'; else if (k < 0.2) { l.rel = 'none'; delete l.forLid; } else if (k < 0.24) delete l.rel; else if (k < 0.28) l.forLids = [pk(targets), pk(targets)];
        if (rr() < 0.05) { l.auto = 'stamp'; }
        return l;
      });
      // oracle
      const SCALE = 10000n;
      const toScaled = (x) => BigInt(Math.round(x * 10000));
      const expect = {};
      const entries = gen.map((x) => ({ key: x.lid, isNew: false, qty: x.qty, unit: x.unit })).concat(nws.map((x) => ({ key: x.nid, isNew: true, qty: x.qty, unit: (String(x.unit).trim() || '式') })));
      const hit = {};
      lns.forEach((l) => {
        if (l.auto === 'stamp' || l.rel !== 'link') return;
        const ks = new Set([l.forLid].concat(l.forLids || []).filter(Boolean));
        ks.forEach((k) => { (hit[k] = hit[k] || []).push(l); });
      });
      let exConf = 0, exZero = 0;
      entries.forEach((e) => {
        const ls = hit[e.key];
        let qty = e.qty, unit = e.unit, conflict = false, zero = false;
        if (ls && ls.length) {
          const us = [...new Set(ls.map((l) => String(l.unit).trim() || '式'))];
          if (us.length > 1) { conflict = true; exConf++; }
          else {
            const sum = ls.reduce((s, l) => s + toScaled(l.qty), 0n);
            unit = us[0];
            if (sum > 0n) { qty = Number(sum < toScaled(0.001) ? toScaled(0.001) : sum) / 10000; } else { qty = 0; zero = true; }
          }
        } else if (e.isNew && !(e.qty > 0)) zero = true;
        if (zero) exZero++;
        expect[e.key] = { qty, unit, conflict, zero };
      });
      const got = CL.applyLinks(its, lns, nws);
      cases++;
      let ok = got.items.length === entries.length && got.conflicts.length === exConf && got.zero.length === exZero;
      got.items.forEach((g) => {
        const k = g.lid || g.nid, ex = expect[k];
        if (!ex) { ok = false; return; }
        const exUnit = ex.conflict ? (entries.find((e) => e.key === k).unit) : ex.unit;
        const exQty = ex.conflict ? entries.find((e) => e.key === k).qty : ex.qty;
        const wantQty = ex.conflict ? exQty : (ex.zero ? 0 : ex.qty);
        if (g.conflict !== ex.conflict || g.zero !== ex.zero) ok = false;
        if (!ex.conflict && Math.abs(Number(g.qty) - wantQty) > 1e-9) ok = false;
        if (!ex.conflict && !ex.zero && String(g.unit) !== exUnit) ok = false;
      });
      // 不變式：擷取到的 changes 只含「真的有變」的非衝突非 zero 品項；套用後再算一次沒有 qty／unit 異動
      const mat2 = CL.materializeItems(its, got, (e) => ({ lid: e.nid, desc: e.desc, unit: e.unit, qty: e.qty, unitPrice: 0 }));
      const again = CL.applyLinks(mat2.filter((x) => !nws.some((n) => n.nid === x.lid)), lns, []);
      if (again.items.some((x) => x.qtyChanged || x.unitChanged) && got.conflicts.length === 0 && got.zero.length === 0) ok = false;
      // 列順序不影響結果
      const rev = CL.applyLinks(its, lns.slice().reverse(), nws);
      if (!eq(rev.items.map((x) => [x.lid || x.nid, x.qty, x.unit, x.conflict, x.zero]), got.items.map((x) => [x.lid || x.nid, x.qty, x.unit, x.conflict, x.zero]))) ok = false;
      if (ok) agree++;
      if (got.conflicts.length) conflictCases++;
      if (got.changes.some((c) => c.field === 'qty' || c.field === 'unit')) changedCases++;
      if (got.zero.length) zeroCases++;
      if (nws.length) newCases++;
      if (got.changes.every((c) => c.field === 'new' || got.items.find((x) => (x.lid || x.nid) === (c.lid || c.nid)) )) permOk++;
    }
    t('3.22 隨機 ' + cases + ' 組（含暫時品項、多目標 forLids、衝突、zero、髒單位）與獨立 oracle（×10000 整數）逐項一致，且套用後冪等、列順序不影響', agree === cases && cases >= 6000 && conflictCases > 300 && changedCases > 1500 && zeroCases > 100 && newCases > 2000 && permOk === cases, `${agree}/${cases} conflict=${conflictCases} changed=${changedCases} zero=${zeroCases} new=${newCases}`);
  }

  // 涵蓋判定與回填（rel=none 不參與；split／link 照 forLid 涵蓋；沒有 rel 的舊行為不變）
  {
    const qc = (lines) => ({ products: [], items: [IT('x1', '品項甲', 1, '式', 1000), IT('x2', '品項乙', 1, '式', 1000)], costLines: lines });
    const cv = (rel, extra) => norm([L('consult', '品項甲', 1, 1, Object.assign(rel ? { rel } : {}, extra || {}))]).lines;
    t('3.23 涵蓋判定：rel=none 的同名列「不」涵蓋同名品項（甲、乙都沒被涵蓋）', eq(CL.unmatchedItems(qc(cv('none'))).map((x) => x.name), ['品項甲', '品項乙']));
    t('3.24 涵蓋判定：沒有 rel 的同名列照舊涵蓋（品名多重集合）；link／split 以 forLid 涵蓋（即使品名不同）', eq(CL.unmatchedItems(qc(cv(''))).map((x) => x.name), ['品項乙'])
      && eq(CL.unmatchedItems(qc(norm([L('consult', '隨便', 1, 1, { rel: 'link', forLid: 'x2' }), L('consult', '另一個', 1, 1, { rel: 'split', forLids: ['x1'] })]).lines)).map((x) => x.name), []));
    const bf = CL.backfillForLids(qc([]).items, norm([L('consult', '品項甲', 1, 1, { rel: 'none' }), L('consult', '品項乙', 1, 1)]).lines);
    t('3.25 回填：rel=none 的同名列不回填 forLid；沒有 rel 的同名列照舊回填', bf[0].forLid === undefined && bf[1].forLid === 'x2');
    t('3.26 costWarnings 只列沒被涵蓋的有價品項；rel=none 的列不會讓警告消失', CL.costWarnings(qc(cv('none')))[0].count === 2 && CL.costWarnings(qc(cv('')))[0].count === 1);
  }

  // ═════════════════ 4) 委外占比 ═════════════════
  const OQ = (lines, revenue) => ({ id: 'q', items: [{ lid: 'a', desc: 'x', qty: 1, unitPrice: revenue === undefined ? 1000000 : revenue }], costLines: norm(lines).lines });
  t('4.1 isOutsourced：顧問區＋vendor 非空白才算；軟體／硬體的 vendor 是供應商不算；印花稅不算', CL.isOutsourced({ cat: 'consult', vendor: 'V' }) && !CL.isOutsourced({ cat: 'consult', vendor: '' }) && !CL.isOutsourced({ cat: 'consult', vendor: '   ' }) && !CL.isOutsourced({ cat: 'software', vendor: 'V' }) && !CL.isOutsourced({ cat: 'hw', vendor: 'V' }) && !CL.isOutsourced({ cat: 'consult', vendor: 'V', auto: 'stamp' }) && !CL.isOutsourced(null) && !CL.isOutsourced({ cat: 'consult', vendor: 5 }));
  let q = OQ([L('consult', 'PM', 10, 1000, { consultant: 'Amy' }), L('consult', 'SD', 5, 2000)]);
  let os = CL.outsourcedStats(q, 100000000);
  t('4.2 全自家（沒有委外廠商）：委外 0、占總成本 0.00%、占顧問成本 0.00%', os.outsourcedCents === 0 && os.outsourcedPctOfCost === '0.00' && os.outsourcedPctOfConsult === '0.00' && os.totalCents === 2000000 && os.consultCents === 2000000, JSON.stringify(os));
  q = OQ([L('consult', 'PM', 10, 1000, { vendor: 'V' }), L('consult', 'SD', 5, 2000, { vendor: 'W' })]);
  os = CL.outsourcedStats(q, 100000000);
  t('4.3 全委外：占顧問成本 100.00%、占總成本 100.00%（沒有其他成本時）', os.outsourcedCents === 2000000 && os.outsourcedPctOfCost === '100.00' && os.outsourcedPctOfConsult === '100.00');
  q = OQ([L('consult', 'PM', 10, 1000, { vendor: 'V' }), L('consult', 'SD', 10, 3000), L('software', 'S', 1, 10000, { vendor: 'Supplier' }), L('travel', 'T', 1, 5000), { cat: 'other', auto: 'stamp' }], 10000000);
  os = CL.outsourcedStats(q, 1000000000);
  // 營收 10,000,000 元 → 印花稅 10,000 元；總成本＝10,000＋30,000＋10,000＋5,000＋10,000＝65,000 元；顧問成本 40,000 元；委外 10,000 元
  t('4.4 混合：總成本含差旅／軟體供應商／印花稅（分母 65,000 元），委外 10,000 元 → 15.38%（向 0 截斷，不是 15.39）；占顧問成本 25.00%', os.outsourcedCents === 1000000 && os.totalCents === 6500000 && os.outsourcedPctOfCost === '15.38' && os.outsourcedPctOfConsult === '25.00', JSON.stringify(os));
  t('4.5 軟體區有 vendor 不算委外：委外只含顧問區那一列', CL.outsourcedCents(q) === 1000000);
  q = OQ([L('software', 'S', 1, 1000, { vendor: 'Supplier' }), L('travel', 'T', 1, 500)]);
  os = CL.outsourcedStats(q, 100000000);
  t('4.6 沒有顧問服務成本：占顧問成本 → null；委外 0、占總成本 0.00', os.outsourcedPctOfConsult === null && os.outsourcedPctOfCost === '0.00' && os.consultCents === 0);
  q = OQ([]);
  os = CL.outsourcedStats(q, 100000000);
  t('4.7 沒有任何成本列：總成本 0 → 占總成本 "0.00"（不是 NaN／null）、占顧問成本 null', os.outsourcedPctOfCost === '0.00' && os.outsourcedPctOfConsult === null && os.totalCents === 0 && os.outsourcedCents === 0);
  t('4.8 舊式單（沒有 costLines）：委外 0、不丟例外', CL.outsourcedCents({ items: [] }) === 0 && CL.outsourcedStats({ items: [] }, 1).outsourcedPctOfCost === '0.00');
  q = OQ([L('consult', 'A', 1, 1, { vendor: 'V' }), L('consult', 'B', 2, 1)]);
  os = CL.outsourcedStats(q, 100000000);
  t('4.9 截斷邊界：1/3＝33.33（不是 33.34 也不是 33）；2/3 的另一例 66.66', os.outsourcedPctOfConsult === '33.33' && CL.pctText(2, 3) === '66.66' && CL.pctText(1, 3) === '33.33' && CL.pctText(999, 1000) === '99.90' && CL.pctText(1, 1000) === '0.10' && CL.pctText(1, 100000) === '0.00' && CL.pctText(5, 5) === '100.00');
  t('4.10 pctText 非法輸入 → null（分母 0／負、分子負、非安全整數、NaN）', [[1, 0], [1, -5], [-1, 5], [NaN, 5], [5, NaN], [1.5, 3], [Number.MAX_SAFE_INTEGER + 2, 5], ['1', 3], [null, 3]].every(([a, b]) => CL.pctText(a, b) === null));
  // 每列取整：委外合計是「每列取整後的分」加總，不是加總後再取整
  q = OQ([L('consult', 'A', 1, 0.333, { vendor: 'V' }), L('consult', 'B', 1, 0.333, { vendor: 'V' }), L('consult', 'C', 1, 0.333, { vendor: 'V' })]);
  t('4.11 每列取整：3 列 0.333 元 → 每列 33 分、合計 99 分（不是 0.999 元取整成 100 分）', CL.outsourcedCents(q) === 99 && CL.totalsByCat(q, 0).consult === 99);
  q = OQ([L('consult', 'A', 1, 1e12, { vendor: 'V' }), L('consult', 'B', 1, 1e12, { vendor: 'V' }), L('consult', 'C', 1, 7)]);
  t('4.12 大金額（兩列各 1e12 元）：委外合計精確（2e14 分）、占顧問成本 99.99…截斷', CL.outsourcedCents(q) === 2e14 && CL.outsourcedStats(q, 1e15).outsourcedPctOfConsult === '99.99', JSON.stringify(CL.outsourcedStats(q, 1e15)));
  q = { id: 'q', costLines: [{ cat: 'consult', vendor: 'V', qty: 'bad', unitCost: 1, desc: 'x' }] };
  t('4.13 壞資料（金額算不出來）：outsourcedStats 全 0，不丟例外', CL.outsourcedStats(q, 100).outsourcedCents === 0 && CL.outsourcedStats(q, 100).outsourcedPctOfConsult === null);

  // ═════════════════ 5) needPrice／簽章／hash ═════════════════
  const qBase = (extraItems) => ({ id: 'q5', products: ['PX'], items: [{ lid: 'a', desc: '甲', unit: '式', qty: 1, unitPrice: 100000, cost: 0 }].concat(extraItems || []), costLines: norm([L('consult', 'PM', 1, 10000)]).lines });
  const classes5 = { PX: { cls: 'consult', costBySales: true } };
  let v = QA.validateForSubmit(qBase([{ lid: 'n', desc: '顧問新增', unit: '人天', qty: 3, unitPrice: 0, cost: 0, needPrice: true, addedBy: 'cons1' }]), classes5);
  t('5.1 needPrice 品項單價 0 → blocker ITEM_PRICE_MISSING（訊息列品名），不可送簽', !v.ok && v.errors.some((e) => e.code === 'ITEM_PRICE_MISSING' && /顧問新增/.test(e.message)) && v.derived === null, JSON.stringify(v.errors));
  v = QA.validateForSubmit(qBase([{ lid: 'n', desc: '顧問新增', unit: '人天', qty: 3, unitPrice: 5000, cost: 0, needPrice: true, addedBy: 'cons1' }]), classes5);
  t('5.2 單價補上（>0）即使旗標還在也不擋', v.ok && !v.errors.some((e) => e.code === 'ITEM_PRICE_MISSING'));
  v = QA.validateForSubmit(qBase([{ lid: 'n', desc: '贈品', unit: '式', qty: 1, unitPrice: 0, cost: 0 }]), classes5);
  t('5.3 沒有 needPrice 旗標的 0 元品項（贈品）不受影響', v.ok);
  v = QA.validateForSubmit(qBase(Array.from({ length: 7 }, (_, i) => ({ lid: 'n' + i, desc: '品' + i, qty: 1, unitPrice: 0, needPrice: true }))), classes5);
  t('5.4 多個待補單價：訊息最多列 5 個並寫「…另 N 項」', v.errors.some((e) => e.code === 'ITEM_PRICE_MISSING' && /…另 2 項/.test(e.message)));
  v = QA.validateForSubmit({ id: 'q', products: ['PX'], items: [{ lid: 'a', desc: 'x', qty: 1, unitPrice: 1000, cost: 500, needPrice: true }] }, classes5);
  t('5.5 needPrice 旗標對單價已 >0 的品項無影響（舊式單也一樣）', v.ok);
  // 結構簽章
  const sItems = [{ lid: 'a', desc: '甲', unit: '式', qty: 2, unitPrice: 100 }, { lid: 'b', desc: '乙', unit: '台', qty: 1, unitPrice: 0 }];
  const sOld = QA.lineStructureSig(sItems), sNew = QA.lineStructureSig(sItems, { newStyle: true });
  t('5.6 lineStructureSig：預設（舊式）每列 5 欄含有價位元，newStyle 每列 4 欄；兩者字串不同', sOld.split('\n').every((ln) => ln.split('|').length === 5) && sNew.split('\n').every((ln) => ln.split('|').length === 4) && sOld !== sNew);
  const flipPrice = sItems.map((x) => Object.assign({}, x, { unitPrice: x.unitPrice ? 0 : 50 }));
  t('5.7 newStyle：有價↔無價不改簽章；舊式會改；改說明／數量／單位／增刪兩種都會改', QA.lineStructureSig(flipPrice, { newStyle: true }) === sNew && QA.lineStructureSig(flipPrice) !== sOld
    && [(x) => { x[0].desc = 'z'; return x; }, (x) => { x[0].qty = 3; return x; }, (x) => { x[0].unit = '套'; return x; }, (x) => x.slice(1), (x) => x.concat([{ lid: 'c', desc: 'n', unit: '式', qty: 1, unitPrice: 0 }])].every((fn) => QA.lineStructureSig(fn(JSON.parse(JSON.stringify(sItems))), { newStyle: true }) !== sNew));
  t('5.8 structureSig(q)：有 costLines 用 newStyle、沒有用舊式；itemsSig 隨之', QA.structureSig({ items: sItems, costLines: [] }) === sNew && QA.structureSig({ items: sItems }) === sOld && QA.itemsSig({ id: 'z', items: sItems, costLines: [] }) !== QA.itemsSig({ id: 'z', items: sItems }));
  // hash
  const HQ_LINES = norm([L('consult', 'PM', 1, 10, { vendor: 'V' }), L('travel', 'T', 1, 5)]).lines;   // lid 只指派一次
  const hq = () => ({ id: 'h', company: 'C', projectName: 'P', products: ['PX'], items: [{ lid: 'a', desc: 'x', unit: '式', qty: 1, unitPrice: 100, cost: 0 }], costLines: JSON.parse(JSON.stringify(HQ_LINES)) });
  const h0 = QA.contentHash(hq());
  let hv = hq(); hv.costLines[0].consultant = 'Amy';
  t('5.9 contentHash：consultant 有值才進 hash（改了 hash 變），空字串／不存在 hash 不變；rel 不進 hash（forLid 也不進）', QA.contentHash(hv) !== h0 && (() => { const x = hq(); x.costLines[0].consultant = ''; return QA.contentHash(x) === h0; })() && (() => { const x = hq(); x.costLines[0].rel = 'link'; x.costLines[0].forLid = 'a'; return QA.contentHash(x) === h0; })());
  hv.costLines[0].consultant = 'ConsB';
  t('5.10 contentHash：換一個顧問姓名 hash 變；空白包圍的同名視為相同（Amy／" Amy "）', QA.contentHash(hv) !== (() => { const x = hq(); x.costLines[0].consultant = 'Amy'; return QA.contentHash(x); })() && (() => { const a = hq(), b = hq(); a.costLines[0].consultant = 'Amy'; b.costLines[0].consultant = ' Amy '; return QA.contentHash(a) === QA.contentHash(b); })());
  t('5.11 needPrice／addedBy／costDraft 不進 contentHash', (() => { const x = hq(); x.items[0].needPrice = true; x.items[0].addedBy = 'u'; x.costDraft = { newItems: [{ nid: 'n', desc: 'd', unit: '式', qty: 1 }] }; return QA.contentHash(x) === h0; })());
  const cS0 = QA.costLinesSig(hq());
  t('5.12 costLinesSig：consultant、rel、草稿新增品項都會讓簽章改變；沒用到新欄位時與以前相同（空草稿不算）', (() => { const a = hq(); a.costLines[0].consultant = 'Amy'; const b = hq(); b.costLines[0].rel = 'link'; const c = hq(); c.costDraft = { newItems: [{ nid: 'n', desc: 'd', unit: '式', qty: 1 }] }; const d = hq(); d.costDraft = { newItems: [] }; return QA.costLinesSig(a) !== cS0 && QA.costLinesSig(b) !== cS0 && QA.costLinesSig(c) !== cS0 && QA.costLinesSig(d) === cS0 && QA.costLinesSig(a) !== QA.costLinesSig(b); })());

  // ═════════════════ 6) 路由層 ═════════════════
  const USERS = [
    { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
    { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
    { username: 'own2', role: 'user', active: true, displayName: 'Owner2', supervisor: 'mgr1', bu: ['ITS'] },
    { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
    { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
    { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
    { username: 'cons1', role: 'user', active: true, displayName: 'Cons1', bu: ['ITS'] },
    { username: 'cons2', role: 'user', active: true, displayName: 'Cons2', bu: ['ITS'] },
    { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
    { username: 'x1', role: 'user', active: true, displayName: 'Other', bu: ['ERP'] },
  ];
  const CLASSES = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } };
  function mkEnv(quotations) {
    const routes = {};
    const app = {};
    ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
    const env = { data: { quotations, quoteApproval: { roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: [], costProviders: ['cons1', 'cons2'], sealManagers: [] }, productClasses: CLASSES } }, logs: [], notes: [], saves: 0, seq: 0 };
    env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
    const deps = {
      db: { load: () => env.data, save: () => { env.saves++; }, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
      requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
      writeLog: (...a) => env.logs.push(a), pushNotification: (...a) => env.notes.push(a),
      getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1' || rl === 'secretary') return ['own1', 'own2']; return [req.session.user.username]; },
      sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
      buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: PNL.buildQuotePnlExcel, QUOTE_TEMPLATE: '',
      uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
    };
    registerQuoteRoutes(app, deps);
    env.call = (user, method, p, params, body) => new Promise((resolve, reject) => {
      const h = [].concat(...env.routesOf(method, p));
      const role = (env.auth.users.find((u) => u.username === user) || {}).role;
      const req = { session: { user: { username: user, role } }, params: params || {}, query: {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
      const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
      let i = 0;
      const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
      next();
    });
    env.routesOf = (m, p) => routes[m.toUpperCase() + ' ' + p];
    return env;
  }
  const mkQuote = (over) => Object.assign({
    id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00Z', updatedAt: '2026-10-08T00:00:00Z',
    validUntil: '2026-10-30', products: ['PC'], costBy: 'cons1', approval: null,
    discountType: 'none', discountValue: 0,
    items: [
      { lid: 'i1', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 600000, cost: 0 },
      { lid: 'i2', desc: '教育訓練', unit: '式', qty: 1, unitPrice: 100000, cost: 0 },
      { lid: 'i3', desc: '軟體授權', unit: '套', qty: 2, unitPrice: 50000, cost: 0 },
    ],
    costFlow: { state: 'requested', by: 'cons1', requestedAt: '2026-10-08T00:00:00Z', filledAt: null, note: '', sig: null, consultantWrote: false },
  }, over || {});
  const costBody = (extra) => Object.assign({
    costLines: [
      L('consult', 'PM', 20, 7000, { unit: '人天', rel: 'link', forLid: 'i1', consultant: 'Amy' }),
      L('consult', 'SD', 10, 6000, { unit: '人天', rel: 'link', forLid: 'i1', vendor: 'V1' }),
      L('consult', '訓練講師', 2, 5000, { unit: '人天', rel: 'split', forLid: 'i2' }),
      L('software', '授權', 2, 30000, { unit: '套', rel: 'link', forLid: 'i3', vendor: 'Supplier' }),
      L('travel', '差旅', 1, 8000, { rel: 'none' }),
      { cat: 'other', auto: 'stamp' },
    ], done: false, costModel: 2,
  }, extra || {});
  let env = mkEnv([mkQuote()]);
  let rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, costBody());
  const qLive = () => env.data.quotations.find((x) => x.id === 'Q1');
  t('6.1 草稿儲存（done:false）→ 200；報價 items 一個欄位都沒動；成本列存下 rel／consultant', rp.s === 200 && eq(qLive().items.map((i) => [i.lid, i.qty, i.unit]), [['i1', 1, '式'], ['i2', 1, '式'], ['i3', 2, '套']]) && qLive().costLines[0].consultant === 'Amy' && qLive().costLines[0].rel === 'link' && qLive().costLines[2].rel === 'split', rp.s + ' ' + JSON.stringify(rp.j && rp.j.code));
  t('6.2 草稿期間 costFlow 仍 requested、沒有 itemChanges；回應的 costLines 帶 consultant／rel', rp.j.costFlow.state === 'requested' && !('itemChanges' in rp.j.costFlow) && rp.j.costLines[0].consultant === 'Amy' && rp.j.costLines[1].vendor === 'V1' && rp.j.costLines[3].rel === 'link');
  // 草稿新增報價項目
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, costBody({ newItems: [{ nid: 'cn-1', desc: '額外教育訓練', unit: '場', qty: 2 }], costLines: costBody().costLines.concat([L('consult', '額外講師', 4, 5000, { unit: '場', rel: 'link', forLid: 'cn-1' })]) }));
  t('6.3 草稿新增報價項目：q.costDraft 存 newItems；報價 items 仍不動；回應（canEditCost）帶 costDraft', rp.s === 200 && eq(qLive().costDraft, { newItems: [{ nid: 'cn-1', desc: '額外教育訓練', unit: '場', qty: 2 }] }) && qLive().items.length === 3 && rp.j.costDraft && rp.j.costDraft.newItems.length === 1, rp.s + JSON.stringify(rp.j && rp.j.code));
  const own = (await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' })).j;
  t('6.4 業務（擁有者）在顧問完成前看不到 costDraft 也看不到 itemChanges；items 與草稿前相同', own.costDraft === undefined && own.costFlow.itemChanges === undefined && own.items.length === 3 && own.items[0].qty === 1);
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, { newItems: undefined, costLines: costBody().costLines.concat([L('consult', '額外講師', 4, 5000, { unit: '場', rel: 'link', forLid: 'cn-1' })]), done: false, costModel: 2 });
  t('6.5 沒帶 newItems → 沿用已存草稿（不清空）', rp.s === 200 && qLive().costDraft && qLive().costDraft.newItems.length === 1);
  // 驗證錯誤
  const badCases = [
    ['6.6a rel 不合法', costBody({ costLines: [L('consult', 'PM', 1, 1, { rel: 'bogus', forLid: 'i1' })] }), 400, 'BAD_COST_LINE'],
    ['6.6b link 沒有目標', costBody({ costLines: [L('consult', 'PM', 1, 1, { rel: 'link' })] }), 400, 'BAD_COST_LINE'],
    ['6.6c split 指向不存在的品項', costBody({ costLines: [L('consult', 'PM', 1, 1, { rel: 'split', forLid: 'ghost' })] }), 400, 'BAD_COST_LINE'],
    ['6.6d link 指向不存在的暫時品項', costBody({ costLines: [L('consult', 'PM', 1, 1, { rel: 'link', forLid: 'cn-zzz' })], newItems: [] }), 400, 'BAD_COST_LINE'],
    ['6.6e newItems nid 重複', costBody({ newItems: [{ nid: 'a1', desc: 'x', qty: 1 }, { nid: 'a1', desc: 'y', qty: 1 }] }), 400, 'BAD_COST_LINE'],
    ['6.6f newItems 撞既有品項 lid', costBody({ newItems: [{ nid: 'i1', desc: 'x', qty: 1 }] }), 400, 'BAD_COST_LINE'],
    ['6.6g newItems 超過 20 筆', costBody({ newItems: Array.from({ length: 21 }, (_, i) => ({ nid: 'n' + i, desc: 'x', qty: 1 })) }), 400, 'BAD_COST_LINE'],
    ['6.6h newItems 的 qty 超出範圍', costBody({ newItems: [{ nid: 'a1', desc: 'x', qty: -1 }] }), 400, 'BAD_COST_LINE'],
    ['6.6i newItems desc 空白', costBody({ newItems: [{ nid: 'a1', desc: ' ', qty: 1 }] }), 400, 'BAD_COST_LINE'],
    ['6.6j newItems 不是陣列', costBody({ newItems: 'x' }), 400, 'BAD_COST_LINE'],
    ['6.6k consultant 型別錯', costBody({ costLines: [L('consult', 'PM', 1, 1, { consultant: 5 })] }), 400, 'BAD_COST_LINE'],
  ];
  for (const [name, body, s, code] of badCases) { const before = JSON.stringify(qLive()); const x = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, body); t(name + ' → ' + s + ' ' + code + '，且沒有半套修改', x.s === s && x.j.code === code && JSON.stringify(qLive()) === before, x.s + ' ' + (x.j && x.j.code)); }
  // done:true 一次寫回
  const doneBody = () => ({ costLines: costBody().costLines.concat([L('consult', '額外講師', 4, 5000, { unit: '場', rel: 'link', forLid: 'cn-1' })]), newItems: [{ nid: 'cn-1', desc: '額外教育訓練', unit: '場', qty: 2 }], done: true, contingencyPct: 5, costModel: 2 });
  const beforeDone = JSON.parse(JSON.stringify(qLive()));
  const draftPre = await env.call('cons1', 'POST', '/api/quotations/:id/cost-draft/summary', { id: 'Q1' }, doneBody());
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, doneBody());
  const qd = qLive();
  t('6.7 done:true → 200：連動列寫回 i1（20＋10＝30 人天）、i3（2 套，單位數量沒變）；split 的 i2 不動；暫時品項轉正式品項（新 lid、unitPrice 0、needPrice、addedBy）', rp.s === 200 && qd.items[0].qty === 30 && qd.items[0].unit === '人天' && qd.items[1].qty === 1 && qd.items[1].unit === '式' && qd.items[2].qty === 2 && qd.items[2].unit === '套'
    && qd.items.length === 4 && qd.items[3].desc === '額外教育訓練' && qd.items[3].qty === 4 && qd.items[3].unit === '場' && qd.items[3].unitPrice === 0 && qd.items[3].needPrice === true && qd.items[3].addedBy === 'cons1' && /^u\d+$/.test(qd.items[3].lid), JSON.stringify(qd.items));
  t('6.8 成本列的 nid 換成新品項的真 lid；草稿清除；單價仍是業務的（i1 600000 沒被動）', qd.costLines.find((l) => l.desc === '額外講師').forLid === qd.items[3].lid && !qd.costDraft && qd.items[0].unitPrice === 600000 && rp.j.costDraft === undefined);
  t('6.9 costFlow：state filled、sig＝寫回後的結構簽章（新式格式，不含有價位元）；itemChanges 記錄 i1 的 qty／unit 與新增品項', qd.costFlow.state === 'filled' && qd.costFlow.sig === QA.structureSig(qd) && qd.costFlow.itemChanges.length === 3 && qd.costFlow.itemChanges.some((c) => c.lid === 'i1' && c.field === 'qty' && c.from === 1 && c.to === 30) && qd.costFlow.itemChanges.some((c) => c.lid === 'i1' && c.field === 'unit' && c.from === '式' && c.to === '人天') && qd.costFlow.itemChanges.some((c) => c.field === 'new' && c.lid === qd.items[3].lid && c.to === 4 && c.unit === '場') && qd.costFlow.itemChangesBy === 'cons1' && !!qd.costFlow.itemChangesAt, JSON.stringify(qd.costFlow));
  t('6.10 回應 costFlow.itemChanges／itemChangesAt／itemChangesByName 給前端橫幅；業務端也看得到（只有品名／數量／單位）', rp.j.costFlow.itemChanges.length === 3 && rp.j.costFlow.itemChangesByName === 'Cons1' && (await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' })).j.costFlow.itemChanges.length === 3);
  const ownAfter = (await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' })).j;
  t('6.11 業務端：新品項帶 needPrice／addedBy／addedByName；送簽預覽有 ITEM_PRICE_MISSING 阻擋', ownAfter.items[3].needPrice === true && ownAfter.items[3].addedBy === 'cons1' && ownAfter.items[3].addedByName === 'Cons1' && ownAfter.preview.blockers.some((b) => b.code === 'ITEM_PRICE_MISSING'));
  t('6.12 通知業務：文字含「同步調整了報價品項」與「1 個新增品項需補單價」；沒有「可以送簽了」', env.notes.length === 1 && env.notes[0][0] === 'own1' && /同步調整了報價品項/.test(env.notes[0][3]) && /1 個新增品項需補單價/.test(env.notes[0][3]) && !/可以送簽了/.test(env.notes[0][3]), JSON.stringify(env.notes[0]));
  const alog = env.logs.filter((l) => l[0] === 'FILL_QUOTE_COST').pop();
  t('6.13 稽核 FILL_QUOTE_COST 詳述品項異動（數量 1→30、單位 式→人天、新增品項待補單價），沒有 undefined', alog && /報價品項同步（3 項）/.test(alog[3]) && /數量 1→30/.test(alog[3]) && /單位 式→人天/.test(alog[3]) && /待補單價/.test(alog[3]) && !/undefined/.test(alog[3]), alog && alog[3]);
  // 業務補單價：needPrice 解除、不退回 requested
  const putItems = (id, patch) => ({ items: qLive().items.map((i) => Object.assign({ lid: i.lid, desc: i.desc, unit: i.unit, qty: i.qty, unitPrice: i.unitPrice }, patch && patch[i.lid] ? patch[i.lid] : {})), costModel: 2, rowKinds: 1 });
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, putItems('Q1', { [qd.items[3].lid]: { unitPrice: 30000 } }));
  t('6.14 業務補單價 → 200；needPrice 清除（addedBy 保留）；costFlow 仍 filled（有價位元不再進結構簽章）；itemChanges 保留', rp.s === 200 && !('needPrice' in qLive().items[3]) && qLive().items[3].addedBy === 'cons1' && qLive().costFlow.state === 'filled' && qLive().costFlow.itemChanges.length === 3, rp.s + ' ' + JSON.stringify(rp.j && (rp.j.code || rp.j.costFlow)));
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, putItems('Q1', { i2: { unitPrice: 0 } }));
  t('6.15 業務把別的品項改成 0 元（有價↔無價）→ 仍 filled（新式成本不看有價位元）', rp.s === 200 && qLive().costFlow.state === 'filled');
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, putItems('Q1', { i2: { unitPrice: 100000, qty: 5 } }));
  t('6.16 業務改品項數量（結構改變）→ 退回 requested，仍保留 itemChanges 直到顧問下次完成', rp.s === 200 && qLive().costFlow.state === 'requested' && qLive().costFlow.itemChanges.length === 3);
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: qLive().items.map((i) => ({ lid: i.lid, desc: i.desc, unit: i.unit, qty: i.qty, unitPrice: i.lid === 'i2' ? 0 : i.unitPrice, needPrice: true, addedBy: 'someone' })), costModel: 2, rowKinds: 1 });
  t('6.16b 業務送來的 needPrice／addedBy 一律忽略（旗標只由伺服器依 lid 沿用舊值；顧問新增的品項照舊保留）', rp.s === 200 && qLive().items[1].needPrice === undefined && qLive().items[1].addedBy === undefined && qLive().items[0].needPrice === undefined && qLive().items[2].needPrice === undefined && qLive().items[3].addedBy === 'cons1', JSON.stringify(qLive().items.map((i) => [i.lid, i.needPrice, i.addedBy])));
  // 補完單價後可送簽？（預覽 blockers）
  rp = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
  t('6.17 補完單價後 ITEM_PRICE_MISSING 解除（其他阻擋如顧問重填成本仍在）', !rp.j.preview.blockers.some((b) => b.code === 'ITEM_PRICE_MISSING'));
  // 單位衝突／zero／TOO_MANY_ITEMS
  env = mkEnv([mkQuote()]);
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, costBody({ done: true, costLines: [L('consult', 'A', 3, 1000, { unit: '人天', rel: 'link', forLid: 'i1' }), L('consult', 'B', 4, 1000, { unit: '人月', rel: 'link', forLid: 'i1' })] }));
  t('6.18 done 時同一品項的連動列單位不同 → 400 LINK_UNIT_CONFLICT（帶 conflicts），報價與 costFlow 都沒動', rp.s === 400 && rp.j.code === 'LINK_UNIT_CONFLICT' && rp.j.conflicts.length === 1 && rp.j.conflicts[0].lid === 'i1' && qLive().items[0].qty === 1 && qLive().costFlow.state === 'requested');
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, costBody({ done: false, costLines: [L('consult', 'A', 3, 1000, { unit: '人天', rel: 'link', forLid: 'i1' }), L('consult', 'B', 4, 1000, { unit: '人月', rel: 'link', forLid: 'i1' })] }));
  t('6.19 單位衝突只在 done 時擋（草稿儲存可以）', rp.s === 200);
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, costBody({ done: true, costLines: [L('consult', 'A', 3, 1000, { unit: '人天' }), L('consult', 'Z', 0, 1000, { unit: '人天', rel: 'link', forLid: 'i1' })] }));
  t('6.20 done 時連動加總為 0 → 400（checkDone 先擋 qty<=0 的列 BAD_COST_LINE）', rp.s === 400 && /^(BAD_COST_LINE|LINK_QTY_ZERO)$/.test(rp.j.code), rp.j.code);
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, costBody({ done: true, newItems: [{ nid: 'zz', desc: '零數量新品項', unit: '式', qty: 0 }], costLines: [L('consult', 'A', 3, 1000, { unit: '人天' })] }));
  t('6.21 done 時暫時品項數量 0（沒有連動列補數量）→ 400 LINK_QTY_ZERO', rp.s === 400 && rp.j.code === 'LINK_QTY_ZERO' && qLive().items.length === 3, rp.j.code);
  const many = mkQuote({ items: Array.from({ length: 49 }, (_, i) => ({ lid: 'm' + i, desc: '品' + i, unit: '式', qty: 1, unitPrice: 1000, cost: 0 })) });
  env = mkEnv([many]);
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, { done: true, costModel: 2, costLines: [L('consult', 'A', 3, 1000, { unit: '人天' })], newItems: [{ nid: 'a1', desc: 'x', qty: 1 }, { nid: 'a2', desc: 'y', qty: 1 }] });
  t('6.22 done 時寫回後總列數 >50 → 400 TOO_MANY_ITEMS；剛好 50 列可以', rp.s === 400 && rp.j.code === 'TOO_MANY_ITEMS' && qLive().items.length === 49 && (await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, { done: true, costModel: 2, costLines: [L('consult', 'A', 3, 1000, { unit: '人天' })], newItems: [{ nid: 'a1', desc: 'x', qty: 1 }] })).s === 200 && qLive().items.length === 50);
  // 舊行為：沒有 rel 的明細 done → 報價不動、通知文字與以前相同
  env = mkEnv([mkQuote()]);
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, { costLines: [L('consult', 'PM', 10, 7000, { unit: '人天', forLid: 'i1' })], done: true, costModel: 2 });
  t('6.23 沒有 rel／newItems 的舊式送法（現行前端）：done → 報價完全不動、沒有 itemChanges、通知文字與以前一模一樣', rp.s === 200 && qLive().items.every((i, k) => i.qty === [1, 1, 2][k]) && !('itemChanges' in qLive().costFlow) && env.notes[0][3] === 'QU-1｜TestCo｜業務 Owner1　顧問 Cons1 已完成成本，可以送簽了。', env.notes[0] && env.notes[0][3]);
  // done 沒有任何異動時清掉上次的 itemChanges
  env = mkEnv([mkQuote({ costFlow: { state: 'requested', by: 'cons1', requestedAt: 'x', filledAt: null, note: '', sig: null, consultantWrote: true, itemChanges: [{ lid: 'i1', desc: '舊', field: 'qty', from: 1, to: 2 }], itemChangesAt: 'x', itemChangesBy: 'cons1' } })]);
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, { costLines: [L('consult', 'PM', 1, 7000, { unit: '式', rel: 'link', forLid: 'i1' })], done: true, costModel: 2 });
  t('6.24 這次完成沒有品項異動 → 清掉上次的 itemChanges（以最近一次完成為準）', rp.s === 200 && !('itemChanges' in qLive().costFlow) && !('itemChangesAt' in qLive().costFlow));
  // reconcile：業務存檔不清掉 itemChanges；需顧問↔不需顧問切換才清
  env = mkEnv([mkQuote({ costFlow: { state: 'filled', by: 'cons1', requestedAt: 'x', filledAt: 'x', note: '', sig: QA.lineStructureSig(mkQuote().items, { newStyle: true }), consultantWrote: true, itemChanges: [{ lid: 'i1', desc: '舊', field: 'qty', from: 1, to: 2 }], itemChangesAt: 'x', itemChangesBy: 'cons1' }, costLines: norm([L('consult', 'PM', 1, 1000)]).lines })]);
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { projectName: 'renamed' });
  t('6.25 業務存檔（沒改品項結構）：costFlow 維持 filled、itemChanges／At／By 都保留', rp.s === 200 && qLive().costFlow.state === 'filled' && qLive().costFlow.itemChanges.length === 1 && qLive().costFlow.itemChangesBy === 'cons1');
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { products: ['PS'], costBy: null });
  t('6.26 需顧問→不需顧問（flip）：itemChanges 與 costDraft 一併清除', rp.s === 200 && !('itemChanges' in qLive().costFlow) && !('costDraft' in qLive()), JSON.stringify(rp.j && rp.j.code));
  // cost-sync 之前存下的 filled 單（sig 含有價位元，舊格式）：上線後業務存檔不會被誤退回 requested
  const oldSig = QA.lineStructureSig(mkQuote().items);
  env = mkEnv([mkQuote({ costFlow: { state: 'filled', by: 'cons1', requestedAt: 'x', filledAt: 'x', note: '', sig: oldSig, consultantWrote: true }, costLines: norm([L('consult', 'PM', 1, 1000)]).lines })]);
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { projectName: 'renamed' });
  t('6.27 上線前存下的 filled 單（sig 是舊格式）：業務存檔、結構沒變 → 仍 filled（兩種簽章格式都認）', rp.s === 200 && qLive().costFlow.state === 'filled');
  env = mkEnv([mkQuote({ costFlow: { state: 'filled', by: 'cons1', requestedAt: 'x', filledAt: 'x', note: '', sig: oldSig, consultantWrote: true }, costLines: norm([L('consult', 'PM', 1, 1000)]).lines })]);
  rp = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, putItems('Q1', { i1: { qty: 9 } }));
  t('6.28 同上，但業務改了品項數量（結構改變）→ 退回 requested', rp.s === 200 && qLive().costFlow.state === 'requested');

  // STALE_ITEMS 檢查：新式單也認舊格式（含有價位元）的 itemsSig——cost-sync 上線前／舊式轉新式前載入的畫面不會被無謂擋下；結構真的變了仍擋
  {
    const e2 = mkEnv([mkQuote({ costLines: norm([L('consult', 'PM', 1, 1000)]).lines })]);
    const q2 = e2.data.quotations[0];
    const body = (sig) => ({ costLines: [L('consult', 'PM', 2, 1000)], done: false, costModel: 2, itemsSig: sig });
    let x = await e2.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, body(QA.itemsSig(q2)));
    t('6.28b itemsSig 新格式 → 200', x.s === 200);
    x = await e2.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, body(QA.itemsSigLegacy(q2)));
    t('6.28c itemsSig 舊格式（含有價位元）且結構沒變 → 200（新式單也認）', x.s === 200 && QA.itemsSigLegacy(q2) !== QA.itemsSig(q2));
    const staleLegacy = QA.itemsSigLegacy(q2);
    q2.items[0].qty = 9;   // 業務改了品項結構
    x = await e2.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, body(staleLegacy));
    t('6.28d 結構變了之後，拿舊的（兩種格式的）itemsSig → 409 STALE_ITEMS', x.s === 409 && x.j.code === 'STALE_ITEMS');
    x = await e2.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, body('0000000000000000'));
    t('6.28e 隨便一個簽章 → 409 STALE_ITEMS', x.s === 409 && x.j.code === 'STALE_ITEMS');
    const e3 = mkEnv([mkQuote()]);   // 舊式單（沒有 costLines）：只認舊格式
    x = await e3.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, body(QA.itemsSig(e3.data.quotations[0])));
    t('6.28f 舊式單（沒有 costLines）送自己的 itemsSig → 200（第一次存成新式）', x.s === 200);
  }

  // ── 顧問可見範圍矩陣（規格 §3.1）──
  const visFor = async (over, user) => { const e = mkEnv([mkQuote(over)]); const x = await e.call(user, 'GET', '/api/quotations/:id', { id: 'Q1' }); return x; };
  let x1 = await visFor({}, 'cons1');
  t('6.29 被指派顧問（requested）：單價、折扣、報價單預覽資料（報價期限）、preview、perm.canSeePrice 都有；approval null、沒有 contentHash', x1.s === 200 && x1.j.items[0].unitPrice === 600000 && x1.j.discountType === 'none' && x1.j.validUntil === '2026-10-30' && x1.j.perm.canSeePrice === true && !!x1.j.preview && x1.j.approval === null && x1.j.contentHash === undefined, JSON.stringify(x1.j && Object.keys(x1.j)));
  x1 = await visFor({}, 'cons1');
  t('6.30 preview 含 level／tiers／rowLabel／warnings（成本沒填完時 level null）', x1.j.preview && 'level' in x1.j.preview && 'tiers' in x1.j.preview && typeof x1.j.preview.rowLabel === 'string' && Array.isArray(x1.j.preview.warnings));
  for (const [label, over, expectOpen] of [
    ['costFlow filled', { costFlow: { state: 'filled', by: 'cons1', filledAt: 'x', sig: null } }, true],
    ['costFlow requested', {}, true],
    ['approval pending', { approval: { state: 'pending', steps: [], cur: 0 } }, false],
    ['approval approved', { approval: { state: 'approved', steps: [], cur: 0 } }, false],
    ['approval returned', { approval: { state: 'returned', steps: [], cur: 0 } }, true],
    ['costFlow needed（還沒指派流程）', { costFlow: { state: 'needed' } }, false],
    ['status 不是 draft', { status: 'closed' }, false],
  ]) {
    const x = await visFor(over, 'cons1');
    t('6.31 顧問可見範圍：' + label + ' → ' + (expectOpen ? '看得到單價與預覽' : '維持看不到（行為不變）'), x.s === 200 && (x.j.perm.canSeePrice === expectOpen) && ((x.j.items[0].unitPrice !== undefined) === expectOpen) && ((x.j.preview !== undefined) === expectOpen) && x.j.approval === null && x.j.contentHash === undefined, x.s + ' ' + JSON.stringify(x.j && x.j.perm));
  }
  const unassigned = await visFor({}, 'cons2');
  t('6.32 未被指派的顧問名單成員 → 403；無關者 → 403', unassigned.s === 403 && (await visFor({}, 'x1')).s === 403);
  for (const u of ['own1', 'admin1', 'mgr1']) { const x = await visFor({}, u); t('6.33 ' + u + ' 的可見性不變（有單價、有 preview；approval 依角色）', x.s === 200 && x.j.items[0].unitPrice === 600000 && !!x.j.preview && x.j.perm.canSeePrice === true); }
  const ownerView = await visFor({}, 'own1');
  t('6.34 業務擁有者在顧問完成前仍看不到成本明細（canSeeCost false），與以前相同', ownerView.j.perm.canSeeCost === false && ownerView.j.costLines === undefined && ownerView.j.costBreakdown === undefined);
  // 預覽相關端點
  const ex = async (p, user, over) => { const e = mkEnv([mkQuote(over)]); return e.call(user, 'GET', p, { id: 'Q1' }); };
  const il = await ex('/api/quotations/:id/issue-info', 'cons1');
  t('6.35 issue-info：被指派顧問（填成本期間）200；pending 時 403；無關者 403；業務 200', il.s === 200 && (await ex('/api/quotations/:id/issue-info', 'cons1', { approval: { state: 'pending', steps: [], cur: 0 } })).s === 403 && (await ex('/api/quotations/:id/issue-info', 'x1')).s === 403 && (await ex('/api/quotations/:id/issue-info', 'own1')).s === 200);
  const pp = await ex('/api/quotations/:id/pnl-preview', 'cons1');
  t('6.36 pnl-preview：被指派顧問 200（含 html）；pending 403；無關者 403；業務擁有者（顧問未完成）403 不變', pp.s === 200 && typeof pp.j.html === 'string' && pp.j.html.length > 1000 && (await ex('/api/quotations/:id/pnl-preview', 'cons1', { approval: { state: 'pending', steps: [], cur: 0 } })).s === 403 && (await ex('/api/quotations/:id/pnl-preview', 'x1')).s === 403 && (await ex('/api/quotations/:id/pnl-preview', 'own1')).s === 403);
  const expo = await (async () => { const e = mkEnv([mkQuote()]); return e.call('cons1', 'POST', '/api/quotations/:id/export-log', { id: 'Q1' }, { format: 'pdf' }); })();
  t('6.37 export-log（PDF 稽核）：被指派顧問 200、無關者 403', expo.s === 200 && (await (async () => { const e = mkEnv([mkQuote()]); return e.call('x1', 'POST', '/api/quotations/:id/export-log', { id: 'Q1' }, { format: 'pdf' }); })()).s === 403);

  // ── cost-draft 端點 ──
  env = mkEnv([mkQuote()]);
  const sumBody = doneBody();
  const sm1 = await env.call('cons1', 'POST', '/api/quotations/:id/cost-draft/summary', { id: 'Q1' }, sumBody);
  t('6.38 summary → 200，且完全不寫入（q 與儲存次數都不變）', sm1.s === 200 && env.saves === 0 && JSON.stringify(qLive().items) === JSON.stringify(mkQuote().items) && !qLive().costDraft && !qLive().costLines);
  t('6.39 summary 回傳形狀：items[{lid|nid,desc,unit,qty,unitPrice,changed,qtyFrom,unitFrom,needPrice}]、四張卡數字、costBreakdown、委外三個欄位、preview', Array.isArray(sm1.j.items) && sm1.j.items.length === 4 && sm1.j.items[3].nid === 'cn-1' && sm1.j.items[3].needPrice === true && sm1.j.items[3].unitPrice === 0 && sm1.j.items[0].lid === 'i1' && sm1.j.items[0].qty === 30 && sm1.j.items[0].qtyFrom === 1 && sm1.j.items[0].unit === '人天' && sm1.j.items[0].unitFrom === '式' && sm1.j.items[0].changed === true && sm1.j.items[1].changed === false
    && Number.isInteger(sm1.j.revenueCents) && Number.isInteger(sm1.j.costCents) && Number.isInteger(sm1.j.gpCents) && typeof sm1.j.marginText === 'string' && typeof sm1.j.costBreakdown.outsourced === 'number' && typeof sm1.j.outsourcedCents === 'number' && typeof sm1.j.outsourcedPctOfCost === 'string' && 'outsourcedPctOfConsult' in sm1.j && !!sm1.j.preview && 'level' in sm1.j.preview, JSON.stringify(sm1.j).slice(0, 400));
  rp = await env.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, sumBody);
  const after = rp.j, fin = QA.computeFinancials(qLive());
  const prevAfter = after.preview;
  t('6.40 summary 的數字與 done:true 之後的實際結果一致：營收、成本、毛利、毛利率、各分類成本、委外、簽核層級預覽', rp.s === 200 && sm1.j.revenueCents === fin.revenueCents && sm1.j.costCents === fin.costCents && sm1.j.gpCents === fin.gpCents && sm1.j.marginText === prevAfter.marginText
    && eq(sm1.j.costBreakdown, after.costBreakdown) && sm1.j.outsourcedCents === after.costBreakdown.outsourced && sm1.j.preview.level === prevAfter.level && eq(sm1.j.preview.tiers, prevAfter.tiers) && sm1.j.preview.rowLabel === prevAfter.rowLabel,
    JSON.stringify([sm1.j.revenueCents, fin.revenueCents, sm1.j.costCents, fin.costCents, sm1.j.marginText, prevAfter.marginText]));
  t('6.41 summary 的品項數量／單位與 done 後實際一致（i1 30 人天；新品項 4 場）', sm1.j.items[0].qty === qLive().items[0].qty && sm1.j.items[0].unit === qLive().items[0].unit && sm1.j.items[3].qty === qLive().items[3].qty && sm1.j.items[3].unit === qLive().items[3].unit);
  t('6.42 summary 的委外占比：委外＝SD 10×6000＝60,000 元；總成本＝顧問 20×7000＋10×6000＋2×5000＋(授權 60,000)＋差旅 8,000＋額外 20,000＋印花稅；占比與 outsourcedStats 一致', sm1.j.outsourcedCents === 6000000 && sm1.j.outsourcedPctOfCost === CL.outsourcedStats(qLive(), fin.revenueCents).outsourcedPctOfCost && sm1.j.outsourcedPctOfConsult === CL.outsourcedStats(qLive(), fin.revenueCents).outsourcedPctOfConsult, JSON.stringify([sm1.j.outsourcedCents, sm1.j.outsourcedPctOfCost, sm1.j.outsourcedPctOfConsult]));
  // summary 的錯誤與權限
  env = mkEnv([mkQuote()]);
  const sPost = (user, body, over) => env.call(user, 'POST', '/api/quotations/:id/cost-draft/summary', { id: 'Q1' }, body === undefined ? costBody() : body);
  t('6.43 summary 權限：業務擁有者（顧問流程中）403、無關者 403、其他顧問 403、一級主管 403、admin 403（不能填成本的人不能試算）', (await sPost('own1')).s === 403 && (await sPost('x1')).s === 403 && (await sPost('cons2')).s === 403 && (await sPost('mgr1')).s === 403 && (await sPost('admin1')).s === 403);
  t('6.44 summary 驗證：缺 costLines／rel 不合法／目標不存在／newItems 錯誤 → 400 BAD_COST_LINE', (await sPost('cons1', {})).s === 400 && (await sPost('cons1', costBody({ costLines: [L('consult', 'a', 1, 1, { rel: 'x', forLid: 'i1' })] }))).j.code === 'BAD_COST_LINE' && (await sPost('cons1', costBody({ costLines: [L('consult', 'a', 1, 1, { rel: 'link', forLid: 'ghost' })] }))).j.code === 'BAD_COST_LINE' && (await sPost('cons1', costBody({ newItems: [{ nid: 'a b', desc: 'x', qty: 1 }] }))).j.code === 'BAD_COST_LINE');
  const sConf = await sPost('cons1', costBody({ costLines: [L('consult', 'A', 3, 1000, { unit: '人天', rel: 'link', forLid: 'i1' }), L('consult', 'B', 4, 1000, { unit: '人月', rel: 'link', forLid: 'i1' })] }));
  t('6.45 summary 遇到單位衝突不報錯：conflicts 列出、該品項維持原樣（conflict:true），其他數字照算', sConf.s === 200 && sConf.j.conflicts.length === 1 && sConf.j.items[0].conflict === true && sConf.j.items[0].qty === 1 && sConf.j.items[0].unit === '式');
  const sLock = mkEnv([mkQuote({ approval: { state: 'pending', steps: [], cur: 0 } })]);
  t('6.46 單已送簽（鎖定）→ summary 409 LOCKED_PENDING', (await sLock.call('cons1', 'POST', '/api/quotations/:id/cost-draft/summary', { id: 'Q1' }, costBody())).s === 409);
  const sSales = mkEnv([mkQuote({ products: ['PS'], costBy: null, costFlow: { state: 'na' } })]);
  const sSelf = await sSales.call('own1', 'POST', '/api/quotations/:id/cost-draft/summary', { id: 'Q1' }, { costLines: [L('software', '授權', 2, 30000)] });
  t('6.47 業務自填成本（不需顧問）的擁有者也能試算（能 PUT /costs 的人）；newItems 非空 → 400（只有顧問流程能新增報價項目）', sSelf.s === 200 && sSelf.j.items.length === 3 && (await sSales.call('own1', 'POST', '/api/quotations/:id/cost-draft/summary', { id: 'Q1' }, { costLines: [], newItems: [{ nid: 'a', desc: 'x', qty: 1 }] })).s === 400);
  const envP = mkEnv([mkQuote()]);
  const pv = await envP.call('cons1', 'POST', '/api/quotations/:id/cost-draft/pnl-preview', { id: 'Q1' }, doneBody());
  t('6.48 pnl-preview（草稿）→ 200 {quoteNo,html,widthPx,heightPx}；html 含草稿的顧問姓名 Amy 與委外廠商 V1 與新品項成本；不寫入；稽核 VIEW_QUOTE_PNL 標註顧問草稿', pv.s === 200 && pv.j.quoteNo === 'QU-1' && /Amy/.test(pv.j.html) && /V1/.test(pv.j.html) && typeof pv.j.widthPx === 'number' && typeof pv.j.heightPx === 'number' && envP.saves === 0 && envP.logs.some((l) => l[0] === 'VIEW_QUOTE_PNL' && /顧問草稿/.test(l[3])), pv.s + ' ' + (pv.j && pv.j.code));
  t('6.49 pnl-preview（草稿）權限：無關者 403、業務擁有者 403、缺 costLines 400', (await envP.call('x1', 'POST', '/api/quotations/:id/cost-draft/pnl-preview', { id: 'Q1' }, doneBody())).s === 403 && (await envP.call('own1', 'POST', '/api/quotations/:id/cost-draft/pnl-preview', { id: 'Q1' }, doneBody())).s === 403 && (await envP.call('cons1', 'POST', '/api/quotations/:id/cost-draft/pnl-preview', { id: 'Q1' }, {})).s === 400);

  // ── serialize：costBreakdown.outsourced、其他檢視者 ──
  env = mkEnv([mkQuote({ costFlow: { state: 'filled', by: 'cons1', requestedAt: 'x', filledAt: 'x', note: '', sig: null, consultantWrote: true }, costLines: norm([L('consult', 'PM', 10, 7000, { consultant: 'Amy' }), L('consult', 'SD', 5, 6000, { vendor: 'V1' }), L('travel', 'T', 1, 5000)]).lines })]);
  const sv = (await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' })).j;
  t('6.50 業務（顧問完成後）：costBreakdown 多 outsourced（分）＝SD 5×6000＝30,000 元＝3,000,000 分，consult 內含它（不重複加）；costLines 帶 consultant', sv.costBreakdown.outsourced === 3000000 && sv.costBreakdown.consult === 10000000 && sv.costLines[0].consultant === 'Amy' && !('consultant' in sv.costLines[1]));
  const admin = (await env.call('admin1', 'GET', '/api/quotations/:id', { id: 'Q1' })).j;
  t('6.51 admin／一級主管同樣看得到 consultant 與 outsourced；無 canSeeCost 的人（無關者）403', admin.costLines[0].consultant === 'Amy' && admin.costBreakdown.outsourced === 3000000 && (await env.call('x1', 'GET', '/api/quotations/:id', { id: 'Q1' })).s === 403);
  // 簽核流程不洩漏給顧問：送簽後顧問回到原本的受限視角
  env = mkEnv([mkQuote({ costFlow: { state: 'filled', by: 'cons1', filledAt: 'x', sig: null }, approval: { state: 'pending', hash: 'h', steps: [{ tier: 'mgr1', label: '一級主管', status: 'pending', assignee: 'mgr1' }], cur: 0, derived: { rowKey: 'consult', rowLabel: '顧問服務', level: 1, board: false, marginText: '50.00', tiers: ['mgr1'], reasons: [], revenueCents: 1, costCents: 1, gpCents: 0 }, history: [] } })]);
  const pendingView = (await env.call('cons1', 'GET', '/api/quotations/:id', { id: 'Q1' })).j;
  t('6.52 送簽後（pending）顧問回到受限視角：沒有單價／預覽／approval，連 derived 的毛利率都沒有', pendingView.items[0].unitPrice === undefined && pendingView.preview === undefined && pendingView.approval === null && !JSON.stringify(pendingView).includes('50.00'));

  legacyCompat();
}

function legacyCompat() {
  // ═════════════════ 7) 舊單位元級相容 ═════════════════
  // 決定性隨機舊式單與新式單（沒有任何 cost-sync 新欄位）；摘要＝各項輸出串接後的 sha256，與「改版前」程式（git 的 50fe748）產生的 golden 比對。
  // 新式單的 itemsSig 依規格改成不含有價位元，所以不納入摘要（改納入 lineStructureSig 原簽章函式）。
  const rr = lcg(20261008);
  const pk = (a) => a[Math.floor(rr() * a.length)];
  const DIRTY = ['', null, undefined, 'abc', '12abc', '-5', '1e3', ' 7 ', '0x10', 'NaN', 'Infinity', 1e15, -3, 0.001, '0.1', '99999999999999999999'];
  const nn = (clean) => (rr() < 0.025 ? pk(DIRTY) : (clean !== undefined ? clean : Math.floor(rr() * 500) * 100 + (rr() < 0.2 ? Math.floor(rr() * 99) / 100 : 0)));
  const NAMES = ['品項甲', '品項乙', '品項丙', '', ' ', '<b>x</b>', '很長'.repeat(30), 'dup', 'dup'];
  const genItem = (i) => {
    const k = rr();
    if (k < 0.08) return { lid: 'L' + i, kind: 'title', desc: pk(NAMES) };
    if (k < 0.14) return { lid: 'L' + i, kind: 'subtotal', desc: pk(NAMES) };
    const it = { lid: rr() < 0.9 ? 'L' + i : undefined, desc: pk(NAMES), unit: pk(['式', '人天', '套', '台', '']), qty: nn(1 + Math.floor(rr() * 20)), unitPrice: rr() < 0.08 ? 0 : nn((1 + Math.floor(rr() * 900)) * 1000), cost: rr() < 0.2 ? 0 : nn((1 + Math.floor(rr() * 500)) * 1000) };
    if (rr() < 0.3) it.cat = pk(['consult', 'software', 'hardware', 'other', 'weird']);
    Object.keys(it).forEach((kk) => { if (it[kk] === undefined) delete it[kk]; });
    return it;
  };
  const PRODUCTS = ['PC', 'PS', 'PH', 'PX', 'PM2'];
  const cls2 = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true }, PH: { cls: 'hardware', costBySales: true }, PM2: { cls: 'consult', costBySales: true } };
  const genQ = (i, withLines) => {
    const q = { id: 'q' + i, items: Array.from({ length: Math.floor(rr() * 9) }, (_, j) => genItem(j)), products: Array.from(new Set(Array.from({ length: Math.floor(rr() * 4) }, () => pk(PRODUCTS)))), company: 'TestCo' + i, projectName: pk(['P1', ' P2 ', '']), note: pk(['', 'n1']), validUntil: pk(['', '2026-12-31']) };
    if (rr() < 0.5) { q.discountType = pk(['none', 'percent', 'amount', 'bogus']); q.discountValue = pk([0, 5, 10, 99.5, 100, 150, 5000, 250000, '12', null, -1]); }
    if (rr() < 0.4) q.costBy = 'cons1';
    if (rr() < 0.5) q.contingencyPct = pk([0, 5, 10, 15, 20, 7]);
    if (rr() < 0.6) q.costFlow = { state: pk(['na', 'needed', 'requested', 'filled']), by: 'cons1', requestedAt: 'x', filledAt: null, note: '', consultantWrote: rr() < 0.3 };
    if (withLines) {
      const lids = q.items.filter((x) => !x.kind && x.lid).map((x) => x.lid);
      q.costLines = [];
      const n = Math.floor(rr() * 10);
      for (let k = 0; k < n; k++) {
        const cat = pk(CL.CATS);
        const l = { lid: 'c' + i + '_' + k, cat, desc: pk(['PM', 'SD', '差旅', 'x', '<i>y</i>']), vendor: ['consult', 'software', 'hw'].includes(cat) && rr() < 0.4 ? pk(['V1', 'V2', ' V3 ']) : '', note: rr() < 0.3 ? 'n' : '', unit: pk(['式', '人天', '套']), qty: pk([1, 2, 0.5, 10, 3.5, 0, 12.25]), unitCost: pk([0, 100, 1234.56, 99999, 0.333, 50000]) };
        if (rr() < 0.3 && lids.length) l.forLid = pk(lids);
        if (rr() < 0.15 && lids.length > 1) l.forLids = lids.slice(0, 2);
        q.costLines.push(l);
      }
      if (rr() < 0.35) q.costLines.push({ lid: 'cs' + i, cat: 'other', desc: CL.STAMP_DESC, vendor: '', note: '', unit: '式', qty: 1, unitCost: 0, auto: 'stamp' });
    }
    return q;
  };
  const digestOf = (withLines, count) => {
    const h = crypto.createHash('sha256');
    for (let i = 0; i < count; i++) {
      const q = genQ(withLines ? 100000 + i : i, withLines);
      const parts = [QA.contentHash(q), JSON.stringify(QA.computeFinancials(q)), JSON.stringify(QA.validateForSubmit(q, cls2)), JSON.stringify(QA.buildDerived(q, cls2)), QA.lineStructureSig(q.items), QA.costLinesSig(q)];
      if (!withLines) parts.push(QA.itemsSig(q));
      else parts.push(JSON.stringify(CL.totalsByCat(q, 123456789)), JSON.stringify(CL.publicLines(q, { canSeeCost: true, canSeePrice: true, revenueCents: 5000000 })), JSON.stringify(CL.costWarnings(q)));
      h.update(parts.join('\u0001') + '\n');
    }
    return h.digest('hex');
  };
  const dOld = digestOf(false, 5000), dNew = digestOf(true, 5000);
  if (GOLDEN_MODE) { console.log('GOLDEN_OLD=' + dOld); console.log('GOLDEN_NEW=' + dNew); }
  else {
    t('7.1 舊式單 5000 張（contentHash／computeFinancials／validateForSubmit／buildDerived／lineStructureSig／costLinesSig／itemsSig）摘要＝改版前程式的 golden', dOld === GOLDEN_OLD, dOld);
    t('7.2 新式單 5000 張（沒有 rel／consultant／costDraft／needPrice；contentHash／computeFinancials／validateForSubmit／buildDerived／lineStructureSig／costLinesSig／totalsByCat／publicLines／costWarnings）摘要＝改版前程式的 golden', dNew === GOLDEN_NEW, dNew);
  }
}

// 以「改版前」的 lib（git 50fe748）、同一份產生器算出來的摘要；用 COSTSYNC_LIB_ROOT 指到改版前檔案樹即可重算
const GOLDEN_OLD = 'cc7dffed9be9927fcd6b336e2c27f8bf42f5e12bc022d7f8dd52bfeac2a98477';
const GOLDEN_NEW = '957b0cab3d7fa9429721beeb5f9d5a07958af5443ad39c4d3cbd528e5c20848a';

run().then(() => {
  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => {
    console.log((ok ? 'PASS ' : 'FAIL ') + n + (x && !ok ? '  <- ' + x : ''));
    ok ? pass++ : fail++;
  });
  console.log('\ncost-sync 單元／路由層測試：PASS ' + pass + ' / FAIL ' + fail);
  process.exit(fail ? 1 : 0);
}).catch((e) => { console.error('例外', e.stack); process.exit(2); });
