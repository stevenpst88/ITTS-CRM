#!/usr/bin/env node
/**
 * 「顧問成本畫面連動業務報價」(cost-sync) 前端純函式檢查。用法：node scripts/check-quote-costsync-ui.js
 * 動 _client/quote-costlines.js 的 cost-sync 區段（applyLinks／委外占比／卡片數字／對應下拉／合併規則）或 _client/quote-approval.js 的顧問對話框輔助函式之後必跑。
 * DOM 行為（對應下拉的連動、卡片即時更新、預覽按鈕、確認清單視窗、橫幅、375px、暗色）另以無頭瀏覽器實跑（e2e_ui_costsync.js）。
 *   1) 介面：QCL 新增的匯出、常數
 *   2) applyLinks（連動計算）與伺服器 lib/quoteCostLines.js 的同名函式逐例鏡像：已知案例＋≥6000 組隨機輸入（含髒資料）逐組比對 0 差異；materializeItems、decimalSum 同
 *   3) 百分比／毛利率文字（pctText、marginTextOf）與伺服器 CL.pctText／QA.marginText 逐例相同
 *   4) 委外判定與委外占比（isOutsourced、outsourcedStats）與伺服器 CL.outsourcedStats 逐組相同（≥5000 組，含全自家、全委外、混合、無顧問成本、含差旅／印花稅的分母）
 *   5) 卡片數字 liveSummary（連動後營收、成本、毛利、毛利率、委外占比、changes）與伺服器「草稿試算＋完成之後」的算法逐分相同（≥3000 組）
 *   6) 委外佔比卡片：顯示模型（大數字、進度條寬度、兩行小字、hover／aria 說明）、HTML、寫回（假 DOM）
 *   7) 完成前確認清單文字 syncConfirmText
 *   8) 「對應」下拉與顧問姓名欄的列 HTML（選項、selected、已刪除的原品項、datalist、委外標籤、跳脫）
 *   9) 合併為一列：rel（全連動同品項保持連動並加總、混合改拆項、全不對應）、consultant／vendor 合併規則、成本合計（分）不變
 *  10) 種子（link 選項）、rel／consultant 的清洗（與伺服器 normalizeCostLines 逐項相同）、unmatchedPricedItems 忽略 rel=none（與伺服器逐組相同）
 *  11) 顧問對話框的輔助函式（quote-approval.js：cfDraftQuote、cfNewItemsClean、cfEditorItems、cfConsultantNames、cfDraftBody、cfNewItemsProblem、cfCanSync、錯誤碼對照）
 *  12) 檔案紀律：沒有本機路徑／帳密、新增的 CSS 都是 qcl-／qap-／q- 前綴
 *  13) 舊行為不退步：與「改版前」的 quote-costlines.js（git show 50fe748，或 COSTSYNC_QCL_BASE 指到的檔案）逐組比對——沒有 rel／consultant／link 的輸入，normalize／seedFromItems／totals／合併（顧問區以外）／涵蓋判定／營收／非顧問區的列 HTML 輸出完全相同
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、廠商名、真實費率）。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const SRC_PATH = process.env.QCL_SRC || path.join(ROOT, '_client/quote-costlines.js');       // 變異測試用：指向被破壞的副本
const QAP_PATH = process.env.QAP_SRC || path.join(ROOT, '_client/quote-approval.js');
const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
const CL = require(path.join(ROOT, 'lib/quoteCostLines.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
function lcg(seed) { let s = seed >>> 0; return () => ((s = (Math.imul(s, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff); }

const src = fs.readFileSync(SRC_PATH, 'utf8').replace(/\r\n/g, '\n');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(src, ctx);
const Q = ctx.QCL;
// 跨 realm 的物件比對：一律經 JSON（鍵順序也要相同——兩邊的輸出形狀是契約）
const same = (a, b) => J(a) === J(b);

// ═════════════ 1) 介面 ═════════════
t('1a. QCL 匯出 cost-sync 函式', ['applyLinks', 'materializeItems', 'decimalSum', 'itemCents', 'isOutsourced', 'pctText', 'marginTextOf', 'costStats', 'outsourcedStats', 'outsourcedFromBreakdown', 'outsourcedCardModel', 'outsourcedCardHtml', 'paintOutsourcedCard', 'liveSummary', 'syncConfirmText', 'targetEntries', 'lineCents'].every((k) => typeof Q[k] === 'function'));
t('1b. 常數：RELS＝link／split／none、MAX_NEW_ITEMS 20、MIN_ITEM_QTY 0.001、CONSULTANT_MAX 40（與伺服器相同）',
  same(Q.RELS, ['link', 'split', 'none']) && Q.MAX_NEW_ITEMS === 20 && Q.MIN_ITEM_QTY === 0.001 && Q.CONSULTANT_MAX === 40 && CL.normalizeCostLines([{ cat: 'consult', desc: 'x', qty: 1, unitCost: 1, consultant: 'x'.repeat(60) }], { genLid: () => 'g' }).lines[0].consultant.length === 40 && same(CL.RELS, Q.RELS) && CL.MAX_NEW_ITEMS === Q.MAX_NEW_ITEMS && CL.MIN_ITEM_QTY === Q.MIN_ITEM_QTY,
  J([Q.RELS, Q.MAX_NEW_ITEMS, Q.MIN_ITEM_QTY, Q.CONSULTANT_MAX]));
t('1c. 五區：只有顧問服務區有「顧問姓名」欄（hasConsultant）；軟體／硬體的廠商欄標題維持「供應商」、顧問區是「委外廠商」',
  same(Q.CATS.map((c) => [c.key, !!c.hasConsultant, c.vendorLabel]), [['consult', true, '委外廠商'], ['software', false, '供應商'], ['hw', false, '供應商'], ['travel', false, ''], ['other', false, '']]));

// ═════════════ 2) applyLinks／materializeItems／decimalSum 與伺服器逐例鏡像 ═════════════
const it = (lid, desc, qty, unit, extra) => Object.assign({ lid, desc, qty, unit }, extra || {});
const ln = (cat, extra) => Object.assign({ cat, desc: 'x', qty: 1, unitCost: 0, unit: '式' }, extra || {});
const AL = (items, lines, news) => Q.applyLinks(items, lines, news);
const ALs = (items, lines, news) => CL.applyLinks(items, lines, news);
{
  const items = [{ lid: 'T', kind: 'title', desc: 'Part A' }, it('a', '顧問', 10, '人天'), it('b', '授權', 2, '套'), { lid: 'S', kind: 'subtotal', desc: '小計' }, it('c', '其他', 1, '式')];
  let r = AL(items, [ln('consult', { rel: 'link', forLid: 'a', qty: 8, unit: '人天' })]);
  t('2a. 單一連動列：數量 10 → 8；changes 只有 qty；沒有連動列的品項完全不變', r.items.length === 3 && r.items[0].lid === 'a' && r.items[0].qty === 8 && r.items[0].qtyChanged && !r.items[0].unitChanged && r.items[1].qty === 2 && !r.items[1].changed && same(r.changes, [{ lid: 'a', desc: '顧問', field: 'qty', from: 10, to: 8 }]), J(r.changes));
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'a', qty: 3, unit: '人月' })]);
  t('2b. 單位也改：changes 同時有 qty 與 unit（人天 → 人月）', same(r.changes, [{ lid: 'a', desc: '顧問', field: 'qty', from: 10, to: 3 }, { lid: 'a', desc: '顧問', field: 'unit', from: '人天', to: '人月' }]), J(r.changes));
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'a', qty: 0.1 }), ln('consult', { rel: 'link', forLid: 'a', qty: 0.2 }), ln('consult', { rel: 'link', forLid: 'a', qty: 0.3 })]);
  t('2c. 多列連動數量十進位精確加總：0.1＋0.2＋0.3 ＝ 0.6（不是 0.6000000000000001）；單位都「式」→ 原單位「人天」變成「式」', r.items[0].qty === 0.6 && r.items[0].linkCount === 3 && r.items[0].unit === '式', J(r.items[0]));
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'a', qty: 1, unit: '人天' }), ln('consult', { rel: 'link', forLid: 'a', qty: 1, unit: '人月' })]);
  t('2d. 單位衝突：同一品項的連動列單位不同 → conflict、該品項維持原樣、列入 conflicts（含 units）、不列入 changes', r.items[0].conflict && r.items[0].qty === 10 && r.items[0].unit === '人天' && same(r.conflicts, [{ lid: 'a', desc: '顧問', units: ['人天', '人月'] }]) && r.changes.length === 0, J(r));
  r = AL(items, [ln('consult', { rel: 'split', forLid: 'a', qty: 99 }), ln('travel', { rel: 'none', qty: 99 }), ln('consult', { forLid: 'a', qty: 99 })]);
  t('2e. split、none、沒有 rel 的列都不連動', r.changes.length === 0 && r.items.every((x) => !x.changed));
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'a', qty: 0 })]);
  t('2f. 連動加總為 0 → zero（品項維持原樣、不列 changes）', r.zero.length === 1 && r.zero[0].lid === 'a' && r.items[0].qty === 0 && r.items[0].zero && r.changes.length === 0, J(r));
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'a', qty: 0.0004 })]);
  t('2g. 連動加總 >0 但 <0.001 → 取 0.001（與業務存檔的品項數量下限相同）', r.items[0].qty === 0.001 && !r.items[0].zero, J(r.items[0]));
  r = AL(items, [{ cat: 'other', auto: 'stamp', rel: 'link', forLid: 'a', qty: 5 }]);
  t('2h. 印花稅列不參與連動', r.changes.length === 0);
  const news = [{ nid: 'nid-1', desc: '新增甲', unit: '式', qty: 2 }, { nid: 'nid-2', desc: '新增乙', unit: '人天', qty: 0 }];
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'nid-1', qty: 5, unit: '人天' })], news);
  t('2i. 顧問新增的報價項目：被連動列命中時改用連動結果（數量 2 → 5、單位 式 → 人天）；每個新增品項一筆 field:new；沒有連動且數量 0 → zero',
    r.items.length === 5 && r.items[3].nid === 'nid-1' && r.items[3].qty === 5 && r.items[3].unit === '人天' && r.items[3].isNew && r.items[4].zero && r.zero.length === 1 && r.zero[0].nid === 'nid-2'
    && r.changes.filter((c) => c.field === 'new').length === 2 && r.changes[0].field === 'new', J(r.changes));
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'zzz', qty: 5 })]);
  t('2j. 連動列指向不存在的品項 → 忽略（不影響其他品項）', r.changes.length === 0);
  r = AL([it(undefined, '無代碼', 3, '式')], [ln('consult', { rel: 'link', forLid: 'legacy-0', qty: 9 })]);
  t('2k. 沒有 lid 的品項以 legacy-<索引> 當代碼（同伺服器 serialize）', r.items[0].lid === 'legacy-0' && r.items[0].qty === 9);
  r = AL(items, [ln('consult', { rel: 'link', forLid: 'a', forLids: ['a', 'b'], qty: 4 })]);
  t('2l. 連動列同時命中多個品項（forLids）：每個命中的品項都加上該列數量（前端只會產生單一目標）；同一列對同一品項只算一次', r.items[0].qty === 4 && r.items[1].qty === 4 && r.items[0].linkCount === 1, J(r.items.map((x) => [x.lid, x.qty, x.linkCount])));
  t('2m. 輸入髒資料不丟錯：null／undefined／非陣列／內含 null', (() => { try { AL(null, undefined, 5); AL([null, 3, 'x'], [null, 7, { rel: 'link' }], [null]); return true; } catch (e) { return false; } })());
  t('2n. 純函式：不修改輸入', (() => { const a = J(items), b = J([ln('consult', { rel: 'link', forLid: 'a', qty: 3 })]); const L = [ln('consult', { rel: 'link', forLid: 'a', qty: 3 })]; AL(items, L, news); return J(items) === a && J(L) === b; })());
}
// 隨機鏡像
{
  const rnd = lcg(20261008);
  const pk = (a) => a[Math.floor(rnd() * a.length)];
  const rQty = () => pk([1, 2, 3, 0.5, 0.1, 0.2, 0.3, 1.5, 2.25, 10, 0, 0.0004, 0.0006, 1e-7, 123456.789, '4', ' 5 ', 'abc', '', undefined, null, -3, 1e9, 100]);
  const rUnit = () => pk(['式', '人天', ' 人天 ', '人月', '', undefined, '套', '台', '式']);
  const rItems = () => {
    const n = Math.floor(rnd() * 8);
    const arr = [];
    for (let i = 0; i < n; i++) {
      const k = rnd();
      if (k < 0.08) arr.push({ lid: 'T' + i, kind: 'title', desc: 'Part ' + i });
      else if (k < 0.14) arr.push({ lid: 'S' + i, kind: 'subtotal', desc: '小計' });
      else if (k < 0.2) arr.push(it(undefined, '無代碼' + i, rQty(), rUnit()));
      else if (k < 0.24) arr.push(it('dup', '重複代碼' + i, rQty(), rUnit()));
      else if (k < 0.26) arr.push(null);
      else arr.push(it('i' + i, '品項' + i, rQty(), rUnit(), { unitPrice: pk([0, 100, 5000, 12.5]) }));
    }
    return arr;
  };
  const rNews = () => {
    const n = rnd() < 0.5 ? 0 : 1 + Math.floor(rnd() * 3);
    return Array.from({ length: n }, (_, i) => (rnd() < 0.05 ? null : { nid: 'nid-' + i, desc: '新增' + i, unit: pk(['式', '人天', '', undefined]), qty: rQty() }));
  };
  const rLines = (items, news) => {
    const targets = items.map((x, i) => (x && x.lid ? x.lid : 'legacy-' + i)).concat(news.map((x) => x && x.nid)).concat(['zzz', '', undefined]);
    const n = Math.floor(rnd() * 10);
    return Array.from({ length: n }, () => {
      const k = rnd();
      if (k < 0.05) return { cat: 'other', auto: 'stamp', rel: 'link', forLid: pk(targets), qty: 1, unitCost: 5 };
      if (k < 0.07) return null;
      const l = { cat: pk(['consult', 'consult', 'software', 'hw', 'travel', 'other']), desc: 'L', qty: rQty(), unitCost: pk([0, 1000, 12.5]), unit: rUnit() };
      const rel = pk(['link', 'link', 'link', 'split', 'none', undefined]);
      if (rel) l.rel = rel;
      if (rnd() < 0.85) l.forLid = pk(targets);
      if (rnd() < 0.2) l.forLids = [pk(targets), pk(targets), pk(targets)];
      return l;
    });
  };
  let diffs = 0, firstBad = null, withChange = 0, withConf = 0, withZero = 0, withNew = 0, diffM = 0;
  for (let n = 0; n < 7000; n++) {
    const items = rItems(), news = rNews(), lines = rLines(items, news);
    const a = AL(items, lines, news), b = ALs(items, lines, news);
    if (!same(a, b)) { diffs++; if (!firstBad) firstBad = J({ items, news, lines, a, b }).slice(0, 700); }
    if (a.changes.some((c) => c.field !== 'new')) withChange++;
    if (a.conflicts.length) withConf++;
    if (a.zero.length) withZero++;
    if (a.items.some((x) => x.isNew)) withNew++;
    const mk = (e) => ({ lid: e.nid, desc: e.desc, unit: e.unit, qty: e.qty, unitPrice: 0, needPrice: true });
    if (!same(Q.materializeItems(items, a, mk), CL.materializeItems(items, b, mk))) diffM++;
  }
  t('2o. applyLinks 與伺服器 CL.applyLinks：7000 組隨機輸入（含髒資料、重複 lid、沒有 lid、null 列、新增品項）逐組輸出完全相同', diffs === 0, diffs + ' 組不同；' + (firstBad || ''));
  t('2p. 隨機輸入有涵蓋各種結果（有連動異動／單位衝突／zero／新增品項的組數都 >300，避免空轉）', withChange > 300 && withConf > 300 && withZero > 300 && withNew > 300, J([withChange, withConf, withZero, withNew]));
  t('2q. materializeItems 與伺服器 CL.materializeItems：同一批 7000 組逐組相同', diffM === 0, diffM);
  let dd = 0;
  for (let n = 0; n < 3000; n++) {
    const arr = Array.from({ length: Math.floor(rnd() * 8) }, rQty).concat(rnd() < 0.2 ? [NaN, Infinity] : []);
    if (!Object.is(Q.decimalSum(arr), CL.decimalSum(arr))) dd++;
  }
  t('2r. decimalSum 與伺服器 CL.decimalSum：3000 組隨機（含非數字、負數、NaN）逐組相同', dd === 0, dd);
  t('2s. decimalSum 已知案例：0.1＋0.2＝0.3；[]＝0；負數與非數字當 0', Q.decimalSum([0.1, 0.2]) === 0.3 && Q.decimalSum([]) === 0 && Q.decimalSum([-1, 'x', 2]) === 2 && Q.decimalSum(undefined) === 0);
}

// 單列成本（分）：卡片的成本、毛利、委外占比都由它加總，必須與伺服器 centsOf（十進位精確、half-up）逐分相同——浮點算法在數值大、小數位多時會差 1 分
{
  const rnd = lcg(2468);
  const pk = (a) => a[Math.floor(rnd() * a.length)];
  let dl = 0, first = null, big = 0;
  for (let n = 0; n < 8000; n++) {
    const qty = pk([1, 2, 0.5, 3.25, 10, 123456.789, 1e6, 999999.999, 0.001, 7.125, 1e9, 33.333]);
    const unitCost = pk([0, 1, 1000, 33.333, 12345.6789, 0.005, 99999999.9999, 1e12, 0.0001, 7654321.123456, 2.675]);
    const l = { cat: 'consult', desc: 'x', qty, unitCost, unit: '式' };
    const a = Q._lineCents(l), b = CL.lineCents(l);
    if (!Number.isFinite(b)) continue;   // 伺服器判定金額過大（超過上限，回 NaN）的列不比較
    if (qty * unitCost > 1e11) big++;
    if (a !== b) { dl++; if (!first) first = J([l, a, b]); }
  }
  t('2t. 單列成本（分）Q.lineCents 與伺服器 CL.lineCents：8000 組隨機（大數值、多位小數、half-up 邊界如 2.675）逐組相同', dl === 0 && big > 500, dl + '；' + (first || '') + ' big=' + big);
}
// ═════════════ 3) pctText／marginTextOf 與伺服器逐例相同 ═════════════
{
  const rnd = lcg(77);
  let dp = 0, dm = 0;
  const big = [0, 1, 2, 3, 7, 100, 9999, 10000, 123456789, 9007199254740990];
  for (let n = 0; n < 6000; n++) {
    const a = rnd() < 0.5 ? Math.floor(rnd() * 1e9) : big[Math.floor(rnd() * big.length)];
    const b = rnd() < 0.5 ? Math.floor(rnd() * 1e9) : big[Math.floor(rnd() * big.length)];
    if (Q.pctText(a, b) !== CL.pctText(a, b)) dp++;
    const gp = rnd() < 0.3 ? -Math.floor(rnd() * 1e9) : Math.floor(rnd() * 1e9);
    if (Q.marginTextOf(gp, b) !== QA.marginText(gp, b)) dm++;
  }
  t('3a. pctText 與伺服器 CL.pctText：6000 組逐組相同', dp === 0, dp);
  t('3b. marginTextOf 與伺服器 QA.marginText：6000 組（含負毛利）逐組相同', dm === 0, dm);
  t('3c. 截斷不是四捨五入：1/3 → 33.33、2/3 → 66.66、199/200 → 99.50、9999/10000 → 99.99；分母 0／負數／非整數 → null', Q.pctText(1, 3) === '33.33' && Q.pctText(2, 3) === '66.66' && Q.pctText(199, 200) === '99.50' && Q.pctText(9999, 10000) === '99.99' && Q.pctText(1, 0) === null && Q.pctText(-1, 5) === null && Q.pctText(1.5, 5) === null);
  t('3d. 毛利率：負毛利 -1/3 → -33.33；毛利為負但小於 0.01% 顯示 0.00（沒有負號）', Q.marginTextOf(-1, 3) === '-33.33' && Q.marginTextOf(-1, 100000) === '0.00' && Q.marginTextOf(5, 0) === null);
}

// ═════════════ 4) 委外判定與委外占比 ═════════════
const sv = (lines, revenueCents) => CL.outsourcedStats({ costLines: lines }, revenueCents);
{
  t('4a. 委外判定：顧問服務區、非印花稅、委外廠商去空白後非空；軟體／硬體的 vendor（供應商）不算；全形空白算空白以外的字元（與伺服器相同的 trim 規則）',
    Q.isOutsourced({ cat: 'consult', vendor: '甲' }) && !Q.isOutsourced({ cat: 'consult', vendor: '   ' }) && !Q.isOutsourced({ cat: 'consult' }) && !Q.isOutsourced({ cat: 'software', vendor: '甲' }) && !Q.isOutsourced({ cat: 'hw', vendor: '甲' })
    && !Q.isOutsourced({ cat: 'consult', vendor: '甲', auto: 'stamp' }) && !Q.isOutsourced(null) && !Q.isOutsourced({ cat: 'consult', vendor: 5 }));
  const mk = (cat, vendor, qty, cost, extra) => Object.assign({ cat, desc: 'x', vendor, qty, unitCost: cost, unit: '式' }, extra || {});
  const stamp = { cat: 'other', auto: 'stamp', desc: '印花稅' };
  const cases = [
    ['全自家', [mk('consult', '', 10, 1000), mk('consult', '', 5, 2000)], 1000000],
    ['全委外', [mk('consult', '甲', 10, 1000), mk('consult', '乙', 5, 2000)], 1000000],
    ['混合', [mk('consult', '甲', 10, 1000), mk('consult', '', 5, 2000), mk('software', '供應商', 1, 3000), mk('travel', '', 1, 500)], 2000000],
    ['無顧問成本', [mk('software', '供應商', 2, 1500), mk('travel', '', 1, 100)], 500000],
    ['含印花稅分母', [mk('consult', '甲', 1, 100000), stamp], 1000000000],
    ['總成本 0', [mk('consult', '甲', 1, 0)], 100000],
    ['無營收（印花稅不計）', [mk('consult', '甲', 3, 3333.33), stamp], 0],
  ];
  cases.forEach(([nm, lines, revC]) => {
    const a = Q.outsourcedStats(lines, revC / 100), b = sv(lines, revC);
    t('4b. 已知案例「' + nm + '」：與伺服器 CL.outsourcedStats 相同', same(a, b), J(a) + ' vs ' + J(b));
  });
  t('4c. 邊界：全自家 → 0.00%、顧問占比 0.00；全委外 → 100.00%（含印花稅時不到 100）；無顧問成本 → 占顧問成本 null；總成本 0 → "0.00"',
    Q.outsourcedStats(cases[0][1], 10000).outsourcedPctOfCost === '0.00' && Q.outsourcedStats(cases[0][1], 10000).outsourcedPctOfConsult === '0.00'
    && Q.outsourcedStats([mk('consult', '甲', 1, 100)], 0).outsourcedPctOfCost === '100.00' && Number(Q.outsourcedStats(cases[4][1], 10000000).outsourcedPctOfCost) < 100
    && Q.outsourcedStats(cases[3][1], 5000).outsourcedPctOfConsult === null && Q.outsourcedStats(cases[5][1], 1000).outsourcedPctOfCost === '0.00');
  const rnd = lcg(4242);
  const pk = (a) => a[Math.floor(rnd() * a.length)];
  let diffs = 0, first = null, nonZero = 0, hundred = 0, nullc = 0;
  for (let n = 0; n < 6000; n++) {
    let stampUsed = false;   // 實際資料最多一列印花稅（伺服器 normalizeCostLines 只留一列）
    const lines = Array.from({ length: Math.floor(rnd() * 9) }, () => {
      const k = rnd();
      if (k < 0.12 && !stampUsed) { stampUsed = true; return { cat: 'other', auto: 'stamp', desc: '印花稅' }; }
      return {
        cat: pk(['consult', 'consult', 'consult', 'software', 'hw', 'travel', 'other']), desc: 'L', vendor: pk(['', '', '甲', ' 乙 ', '   ']),
        qty: pk([1, 2, 0.5, 3, 10, 0, 0.333, 12.5]), unitCost: pk([0, 1, 1000, 33.333, 12345.67, 0.005, 250000]), unit: '式',
      };
    });
    const revC = pk([0, 100, 99999, 12345678, 100000000, 5000000000]);
    const a = Q.outsourcedStats(lines, revC / 100), b = sv(lines, revC);
    if (!same(a, b)) { diffs++; if (!first) first = J({ lines, revC, a, b }).slice(0, 500); }
    if (a.outsourcedCents > 0) nonZero++;
    if (a.outsourcedPctOfCost === '100.00') hundred++;
    if (a.outsourcedPctOfConsult === null) nullc++;
  }
  t('4d. outsourcedStats 與伺服器：6000 組隨機（全自家／全委外／混合／無顧問成本／含印花稅）逐組完全相同（委外成本分、總成本分、顧問成本分、兩個百分比字串）', diffs === 0, diffs + '；' + (first || ''));
  t('4e. 隨機輸入有涵蓋：有委外成本、委外占比剛好 100%、無顧問成本 的組數都 >100', nonZero > 100 && hundred > 100 && nullc > 100, J([nonZero, hundred, nullc]));
  // outsourcedFromBreakdown：核准面板直接用伺服器序列化的分類彙總（other 含印花稅）
  const fb = Q.outsourcedFromBreakdown({ consult: 1000000, software: 200000, hw: 0, travel: 50000, other: 150000, outsourced: 400000 });
  t('4f. outsourcedFromBreakdown：總成本＝五類相加（other 已含印花稅）、不含風險預留；40 萬 ÷ 140 萬 ＝ 28.57%、占顧問成本 40.00%', fb.totalCents === 1400000 && fb.outsourcedPctOfCost === '28.57' && fb.outsourcedPctOfConsult === '40.00', J(fb));
  t('4g. outsourcedFromBreakdown：缺欄位／非物件 → 0 或 null，不丟錯', Q.outsourcedFromBreakdown(null) === null && Q.outsourcedFromBreakdown({}).outsourcedPctOfCost === '0.00' && Q.outsourcedFromBreakdown({}).outsourcedPctOfConsult === null);
}

// ═════════════ 5) 卡片數字 liveSummary ＝ 伺服器草稿試算＋完成 ═════════════
{
  const rnd = lcg(909);
  const pk = (a) => a[Math.floor(rnd() * a.length)];
  let diffs = 0, first = null, okN = 0, skipped = 0, chg = 0, withOs = 0;
  let seq = 0;
  for (let n = 0; n < 3500; n++) {
    const nItems = 1 + Math.floor(rnd() * 5);
    const items = [];
    for (let i = 0; i < nItems; i++) {
      if (rnd() < 0.12) { items.push({ lid: 'T' + i, kind: 'title', desc: 'Part' }); continue; }
      items.push({ lid: 'i' + i, desc: '品項' + i, unit: pk(['式', '人天', '套', '台']), qty: pk([1, 2, 3, 0.5, 10, 2.25, 100]), unitPrice: pk([0, 1000, 5000, 12345.67, 88888, 150000]), cost: 0 });
    }
    const news = rnd() < 0.4 ? [{ nid: 'nid-a', desc: '新增', unit: pk(['式', '人天']), qty: pk([1, 2, 5, 0]) }] : [];
    const targets = items.filter((x) => !x.kind).map((x) => x.lid).concat(news.map((x) => x.nid));
    const raw = Array.from({ length: 1 + Math.floor(rnd() * 7) }, () => {
      const cat = pk(['consult', 'consult', 'software', 'hw', 'travel', 'other']);
      const l = { cat, desc: 'L' + Math.floor(rnd() * 100), unit: pk(['式', '人天', '套', '台', '人月']), qty: pk([1, 2, 3, 0.5, 0.1, 0.2, 5, 10, 2.5]), unitCost: pk([0, 500, 1000, 3333.33, 12.5, 80000]) };
      if (cat === 'consult' && rnd() < 0.5) l.vendor = pk(['甲', '乙']);
      const rel = pk(['link', 'link', 'split', 'none', undefined]);
      if (rel === 'none') l.rel = 'none';
      else if (rel) { l.rel = rel; l.forLid = pk(targets); }
      return l;
    });
    if (rnd() < 0.5) raw.push({ cat: 'other', auto: 'stamp' });
    const dtype = pk(['none', 'none', 'percent', 'amount']);
    let dval = 0;
    if (dtype === 'percent') dval = pk([50, 79.9, 90, 99.5, 85.25]);
    const nr = CL.normalizeCostLines(raw, { genLid: () => 'g' + (++seq) });
    if (!nr.ok) { skipped++; continue; }
    const q0 = { items, costLines: nr.lines, discountType: dtype === 'amount' ? 'none' : dtype, discountValue: dtype === 'amount' ? 0 : dval };
    if (dtype === 'amount') {
      const sub = QA.computeFinancials(Object.assign({}, q0)).revenueCents;
      if (!(sub > 20000)) { skipped++; continue; }
      q0.discountType = 'amount'; q0.discountValue = Math.floor(sub / 100 * 0.8);
    }
    // 伺服器：同 draftSimulation 的做法
    const newsClean = news.map((x) => ({ nid: x.nid, desc: x.desc, unit: x.unit, qty: x.qty }));
    const al = CL.applyLinks(items, nr.lines, newsClean);
    const sim = JSON.parse(JSON.stringify(q0));
    sim.items = CL.materializeItems(sim.items, al, (e) => ({ lid: e.nid, desc: '新增', unit: e.unit, qty: e.qty, unitPrice: 0, cost: 0, needPrice: true }));
    const fin = QA.computeFinancials(sim);
    // 前端：吃的是畫面上的列（QCL.normalize 後）
    const live = Q.liveSummary({ items, lines: Q.normalize(nr.lines), newItems: newsClean, discountType: q0.discountType, discountValue: q0.discountValue });
    if (!fin.ok) { skipped++; continue; }
    okN++;
    const tot = CL.totalsByCat(sim, fin.revenueCents);
    const os = CL.outsourcedStats(sim, fin.revenueCents);
    const sv2 = {
      revenueCents: fin.revenueCents, costCents: tot.total, gpCents: fin.gpCents, marginText: QA.marginText(fin.gpCents, fin.revenueCents),
      outsourcedCents: os.outsourcedCents, outsourcedPctOfCost: os.outsourcedPctOfCost, outsourcedPctOfConsult: os.outsourcedPctOfConsult,
      other: tot.other, consult: tot.consult, software: tot.software, hw: tot.hw, travel: tot.travel, changes: al.changes, conflicts: al.conflicts, zero: al.zero,
    };
    const cl2 = {
      revenueCents: live.revenueCents, costCents: live.costCents, gpCents: live.gpCents, marginText: live.marginText,
      outsourcedCents: live.outsourced.outsourcedCents, outsourcedPctOfCost: live.outsourced.outsourcedPctOfCost, outsourcedPctOfConsult: live.outsourced.outsourcedPctOfConsult,
      other: live.byCat.other + live.stampCents, consult: live.byCat.consult, software: live.byCat.software, hw: live.byCat.hw, travel: live.byCat.travel, changes: live.changes, conflicts: live.conflicts, zero: live.zero,
    };
    if (!same(sv2, cl2)) { diffs++; if (!first) first = J({ q0, news, sv2, cl2 }).slice(0, 900); }
    if (al.changes.length) chg++;
    if (os.outsourcedCents > 0) withOs++;
  }
  t('5a. liveSummary（連動後營收、成本、毛利、毛利率、五類成本、委外占比、changes／conflicts／zero）與伺服器 draftSimulation 算法逐分相同：' + okN + ' 組有效隨機輸入（含折扣、印花稅、新增報價項目、單位衝突）', diffs === 0 && okN > 2500, diffs + ' 組不同（有效 ' + okN + '）；' + (first || ''));
  t('5b. 隨機輸入有涵蓋：有連動異動、有委外成本的組數都 >500', chg > 500 && withOs > 500, J([chg, withOs, skipped]));
  // 已知案例（手算）
  const kItems = [it('a', '顧問', 10, '人天', { unitPrice: 10000 }), it('b', '授權', 1, '套', { unitPrice: 50000 })];
  const kLines = [
    { cat: 'consult', desc: '顧問甲', vendor: '外包甲', qty: 4, unitCost: 5000, unit: '人天', rel: 'link', forLid: 'a' },
    { cat: 'consult', desc: '顧問乙', consultant: '自家A', qty: 3, unitCost: 4000, unit: '人天', rel: 'link', forLid: 'a' },
    { cat: 'software', desc: '授權', qty: 1, unitCost: 30000, unit: '套', rel: 'split', forLid: 'b' },
    { cat: 'other', auto: 'stamp' },
  ];
  const k = Q.liveSummary({ items: kItems, lines: Q.normalize(kLines), newItems: [], discountType: 'none', discountValue: 0 });
  // 連動後：顧問 4＋3＝7 人天 → 營收 7×10000＋50000＝120000；成本 20000＋12000＋30000＋印花稅 120＝62120；毛利 57880；毛利率 48.23；委外 20000/62120＝32.19%、占顧問 62.50%
  t('5c. 手算案例：連動後 7 人天（業務原值 10）→ 營收 120,000、成本 62,120（含印花稅 120）、毛利 57,880、毛利率 48.23%、委外 20,000（占專案 32.19%、占顧問 62.50%）',
    k.revenueCents === 12000000 && k.costCents === 6212000 && k.gpCents === 5788000 && k.marginText === '48.23' && k.outsourced.outsourcedCents === 2000000 && k.outsourced.outsourcedPctOfCost === '32.19' && k.outsourced.outsourcedPctOfConsult === '62.50'
    && k.changes.length === 1 && k.changes[0].field === 'qty' && k.changes[0].from === 10 && k.changes[0].to === 7, J({ r: k.revenueCents, c: k.costCents, g: k.gpCents, m: k.marginText, o: k.outsourced, ch: k.changes }));
  t('5d. 沒有連動時卡片 ＝ 原報價：把兩個連動列改成拆項 → 營收回到 150,000、changes 空', (() => {
    const lines = kLines.map((l) => (l.rel === 'link' ? Object.assign({}, l, { rel: 'split' }) : l));
    const x = Q.liveSummary({ items: kItems, lines: Q.normalize(lines), newItems: [], discountType: 'none', discountValue: 0 });
    return x.revenueCents === 15000000 && x.changes.length === 0;
  })());
  t('5e. 營收 0（所有品項單價 0）→ 毛利率 null（顯示「—」），不丟錯', (() => {
    const x = Q.liveSummary({ items: [it('a', '贈品', 1, '式', { unitPrice: 0 })], lines: Q.normalize([{ cat: 'consult', desc: 'x', qty: 1, unitCost: 100 }]), newItems: [], discountType: 'none', discountValue: 0 });
    return x.revenueCents === 0 && x.marginText === null && x.gpCents === -10000;
  })());
  t('5f. 輸入髒資料不丟錯（沒給 input／items 不是陣列）', (() => { try { Q.liveSummary(); Q.liveSummary({ items: 5, lines: 'x' }); return true; } catch (e) { return false; } })());
}

// ═════════════ 6) 委外佔比卡片 ═════════════
{
  const st = (lines, rev) => Q.outsourcedStats(lines, rev);
  const m0 = Q.outsourcedCardModel(null);
  t('6a. 沒有資料的卡片：0.00%、進度條 0、「委外成本 NT$ 0」、「占顧問服務成本 —」', m0.pctLabel === '0.00%' && m0.bar === 0 && m0.line1 === '委外成本 NT$ 0' && m0.line2 === '占顧問服務成本 —', J(m0));
  const mm = Q.outsourcedCardModel(st([{ cat: 'consult', vendor: '甲', qty: 1, unitCost: 3000 }, { cat: 'consult', qty: 1, unitCost: 7000 }], 0));
  t('6b. 混合：委外 3,000 ÷ 總成本 10,000 ＝ 30.00%、進度條寬度 30、占顧問服務成本 30.00%、委外成本 NT$ 3,000', mm.pctLabel === '30.00%' && mm.bar === 30 && mm.line1 === '委外成本 NT$ 3,000' && mm.line2 === '占顧問服務成本 30.00%', J(mm));
  const m1 = Q.outsourcedCardModel(st([{ cat: 'consult', vendor: '甲', qty: 1, unitCost: 3000 }], 0));
  t('6c. 全委外：100.00%、進度條 100（不超過 100）', m1.pctLabel === '100.00%' && m1.bar === 100);
  const m2 = Q.outsourcedCardModel(st([{ cat: 'software', qty: 1, unitCost: 3000 }], 0));
  t('6d. 無顧問成本：占顧問服務成本顯示「—」、大數字 0.00%', m2.pctLabel === '0.00%' && m2.line2 === '占顧問服務成本 —');
  t('6e. 進度條寬度＝百分比數字（33.33% → 33.33），夾在 0–100；亂給的值（NaN、999、負數）不丟錯', Q.outsourcedCardModel({ outsourcedPctOfCost: '33.33', outsourcedCents: 100 }).bar === 33.33 && Q.outsourcedCardModel({ outsourcedPctOfCost: '999.00' }).bar === 100 && Q.outsourcedCardModel({ outsourcedPctOfCost: '-5' }).bar === 0 && Q.outsourcedCardModel({ outsourcedPctOfCost: 'abc' }).bar === 0);
  t('6f. hover／aria 說明包含算法（委外成本 ÷ 專案總成本、含差旅／交際費／印花稅、不含風險預留；自家顧問不算委外）', /委外成本 ÷ 專案總成本/.test(mm.title) && /含差旅／交際費／印花稅/.test(mm.title) && /不含風險預留/.test(mm.title) && /自家顧問/.test(mm.title) && mm.aria.indexOf('委外佔比 30.00%') === 0);
  const html = Q.outsourcedCardHtml(mm, 'my"id');
  t('6g. 卡片 HTML：沿用 pnl-sum-card 樣式、標題「委外佔比」、role=progressbar（aria-valuenow）、id 屬性有跳脫、兩行小字、title／aria-label 有跳脫',
    /class="pnl-sum-card qcl-os-card"/.test(html) && /pnl-sum-label">委外佔比</.test(html) && /role="progressbar"[^>]*aria-valuenow="30"/.test(html) && /style="width:30%"/.test(html) && /id="my&quot;id"/.test(html) && /委外成本 NT\$ 3,000/.test(html) && /占顧問服務成本 30\.00%/.test(html) && /tabindex="0"/.test(html), html.slice(0, 300));
  // 假 DOM：paintOutsourcedCard
  const mkEl = () => ({ textContent: '', attrs: {}, style: {}, setAttribute(k, v) { this.attrs[k] = String(v); }, getAttribute(k) { return this.attrs[k]; } });
  const card = mkEl(); card.kids = { '.qcl-os-pct': mkEl(), '.qcl-os-l1': mkEl(), '.qcl-os-l2': mkEl(), '.qcl-os-bar': mkEl(), '.qcl-os-fill': mkEl() };
  card.classList = { contains: (c) => c === 'qcl-os-card' }; card.querySelector = function (s) { return this.kids[s] || null; };
  const okPaint = Q.paintOutsourcedCard(card, mm);
  t('6h. paintOutsourcedCard 寫入：大數字、兩行小字、進度條 width 與 aria-valuenow、title／aria-label', okPaint === true && card.kids['.qcl-os-pct'].textContent === '30.00%' && card.kids['.qcl-os-l1'].textContent === '委外成本 NT$ 3,000' && card.kids['.qcl-os-l2'].textContent === '占顧問服務成本 30.00%'
    && card.kids['.qcl-os-fill'].style.width === '30%' && card.kids['.qcl-os-bar'].attrs['aria-valuenow'] === '30' && card.attrs.title === mm.title && card.attrs['aria-label'] === mm.aria);
  t('6i. paintOutsourcedCard 找不到卡片 → false，不丟錯', Q.paintOutsourcedCard(null, mm) === false && Q.paintOutsourcedCard({ querySelector: () => null, classList: { contains: () => false } }, mm) === false);
}

// ═════════════ 7) 完成前確認清單 ═════════════
{
  t('7a. 沒有異動 → 空字串', Q.syncConfirmText([]) === '' && Q.syncConfirmText(undefined) === '');
  const s1 = Q.syncConfirmText([{ lid: 'a', desc: '顧問', field: 'qty', from: 10, to: 7 }, { lid: 'a', desc: '顧問', field: 'unit', from: '人天', to: '人月' }, { nid: 'n1', desc: '新增品', field: 'new', from: null, to: 2, unit: '式' }]);
  t('7b. 清單：標題「將同步更新業務的報價」、數量 a → b、單位 x → y、新增報價項目（單價由業務補填）、補完單價前不能送簽', /^將同步更新業務的報價：/.test(s1) && /「顧問」　數量 10 → 7/.test(s1) && /「顧問」　單位 人天 → 人月/.test(s1) && /新增報價項目「新增品」　數量 2 式（單價由業務補填）/.test(s1) && /新增的 1 個品項單價是空的，業務補完單價之前不能送簽/.test(s1), s1);
  const many = Array.from({ length: 13 }, (_, i) => ({ lid: 'l' + i, desc: '品項' + i, field: 'qty', from: 1, to: 2 }));
  const s2 = Q.syncConfirmText(many);
  t('7c. 超過 10 項：只列前 10 項，其餘「…另 3 項」', (s2.match(/・/g) || []).length === 10 && /…另 3 項/.test(s2), s2);
  const s3 = Q.syncConfirmText([{ lid: 'a', desc: 'x'.repeat(80), field: 'qty', from: 1.23456789, to: 2 }, { lid: 'b', desc: '<i onerror=1>', field: 'unit', from: '', to: '式' }]);
  t('7d. 長品名截 20 字＋…；數量最多 4 位小數；單位原值空白顯示（空白）；文字一律原樣（由 qapConfirm 的 e() 跳脫，這裡不產生 HTML）', /「x{20}…」/.test(s3) && /1\.2346 → 2/.test(s3) && /單位 （空白） → 式/.test(s3) && s3.indexOf('<i onerror=1>') > 0 && !/&lt;/.test(s3), s3);
}

// ═════════════ 8) 「對應」下拉與顧問姓名欄的列 HTML ═════════════
{
  const items = [{ lid: 'a', desc: '顧問服務', unit: '人天', qty: 10 }, { lid: 'T', kind: 'title', desc: 'Part' }, { lid: 'b', desc: '<b>授權</b>', unit: '套', qty: 1 }, { nid: 'nid-1', desc: '新增品', isDraftNew: true }];
  const entries = Q.targetEntries(items);
  t('8a. targetEntries：標題／小計列不列、編號只算品項、顧問新增的標 isNew；沒有 lid 用 nid', same(entries.map((e) => [e.key, e.seq, e.isNew]), [['a', 1, false], ['b', 2, false], ['nid-1', 3, true]]));
  const html = Q._editRowHtml({ desc: '顧問甲', rel: 'link', forLid: 'b', vendor: '外包甲', consultant: '自家A' }, 'consult', { link: true, entries });
  t('8b. 顧問區列（link）：有「顧問姓名」「委外廠商」輸入、對應方式下拉三個選項（連動報價數量／拆項（報價不動）／不對應（純成本））、目標下拉選取 b、data-rel',
    /class="qcl-in qcl-consultant"[^>]*value="自家A"/.test(html) && /class="qcl-in qcl-vendor"[^>]*value="外包甲"/.test(html) && /<select class="qcl-in qcl-rel"/.test(html) && /<option value="link" selected title="連動報價數量">連動<\/option><option value="split" title="拆項（報價不動）">拆項<\/option><option value="none" title="不對應（純成本）">不對應<\/option>/.test(html)
    && /<option value="b" selected>2\. /.test(html) && /data-rel="link"/.test(html) && /list="n"/.test(html), html.slice(0, 900));
  t('8c. 目標下拉選項的品名一律跳脫（含 <b>、&）、顧問新增的品項標「＋新增」', !/<b>授權<\/b>/.test(html) && /&lt;b&gt;授權&lt;\/b&gt;/.test(html) && /＋新增/.test(html));
  const hn = Q._editRowHtml({ desc: '差旅', rel: 'none' }, 'travel', { link: true, entries });
  t('8d. 非顧問區沒有「顧問姓名」欄；不對應的列目標下拉隱藏（hidden）', !/qcl-consultant/.test(hn) && /<select class="qcl-in qcl-target" hidden/.test(hn) && /<option value="none" selected/.test(hn));
  const hs = Q._editRowHtml({ desc: '授權', vendor: '供應商甲', cat: 'software' }, 'software', { link: true, entries });
  t('8e. 軟體區：沒有顧問姓名欄、有廠商欄（供應商）；沒有 rel 的舊列（有 forLid 才視為拆項，沒有則不對應）', !/qcl-consultant/.test(hs) && /qcl-vendor/.test(hs) && /<option value="none" selected/.test(hs));
  const hl = Q._editRowHtml({ desc: '舊列', forLid: 'a' }, 'consult', { link: true, entries });
  t('8f. 舊資料（沒有 rel、有 forLid）顯示為「拆項」、沒有 data-rel 屬性（畫面不動它就不會被改成明確的 rel）', /<option value="split" selected/.test(hl) && !/data-rel=/.test(hl));
  const ho = Q._editRowHtml({ desc: '孤兒', rel: 'link', forLid: 'gone' }, 'consult', { link: true, entries });
  t('8g. 對應的品項已不存在：目標下拉多一個已選取的「（原報價品項已刪除）」選項，畫面看得出來而不是默默換成別的品項', /<option value="gone" selected>（原報價品項已刪除）<\/option>/.test(ho));
  const hp = Q._editRowHtml({ desc: '顧問甲', consultant: '自家A', vendor: '' }, 'consult');
  t('8h. 沒給 ctx（業務自填成本編輯器）：沒有「對應」欄，但顧問區仍有「顧問姓名」與「委外廠商」欄', !/qcl-rel/.test(hp) && /qcl-consultant/.test(hp) && /qcl-vendor/.test(hp));
  t('8i. 委外標籤：顧問區列首都有一個標籤，有填委外廠商才顯示（沒填的 hidden）；非顧問區沒有標籤', /<span class="qcl-ostag" title="[^"]*">委外<\/span>/.test(Q._editRowHtml({ desc: 'a', vendor: '甲' }, 'consult')) && /<span class="qcl-ostag" hidden/.test(Q._editRowHtml({ desc: 'a', vendor: '' }, 'consult')) && !/qcl-ostag/.test(Q._editRowHtml({ desc: 'a', vendor: '甲' }, 'software')));
  t('8j. 軟體／硬體的廠商欄標題維持「供應商」、顧問區是「委外廠商」，顧問姓名欄只在顧問區（欄位標籤 aria-label）', /aria-label="軟體成本：供應商"/.test(hs) && /aria-label="顧問服務成本：顧問姓名"/.test(html) && /aria-label="顧問服務成本：委外廠商"/.test(html));
  const xss = '"><img src=x onerror=window.__pwned=1>';
  const hx = Q._editRowHtml({ desc: xss, consultant: xss, vendor: xss, note: xss, rel: 'split', forLid: xss }, 'consult', { link: true, entries: [{ key: xss, seq: 1, name: xss, isNew: false }] });
  t('8k. XSS：項目、顧問姓名、廠商、說明、對應目標（代號與品名）全部跳脫，HTML 裡沒有任何未跳脫的 <img', !/<img/i.test(hx) && (hx.match(/&lt;img/g) || []).length >= 4, (hx.match(/<img/gi) || []).length);
  const vx = Q._viewRowHtml({ desc: 'a', consultant: xss, vendor: xss }, 'consult', []);
  t('8l. 唯讀列：顧問姓名欄與委外標籤；字串跳脫', !/<img/i.test(vx) && /委外/.test(vx) && /&lt;img/.test(vx));
}

// ═════════════ 9) 合併為一列 ═════════════
{
  const MG = Q._mergeLines;
  const cents = (ls) => ls.reduce((s, l) => s + Q._lineCents(l), 0);
  const row = (extra) => Object.assign({ cat: 'consult', desc: 'r', qty: 2, unitCost: 1000.5, unit: '人天' }, extra || {});
  // 沒有 rel：與改版前相同（輸出沒有 rel 鍵）
  let m = MG([row(), row({ desc: 's' })]);
  t('9a. 都沒有 rel（業務自填、舊單）→ 輸出沒有 rel 鍵、單位式、數量 1（改版前的行為）', !('rel' in m) && m.unit === '式' && m.qty === 1 && m.unitCost === 4002 && m.cat === 'consult');
  m = MG([row({ rel: 'link', forLid: 'a', qty: 2 }), row({ rel: 'link', forLid: 'a', qty: 3, unitCost: 777.77 })]);
  t('9b. 全部連動同一品項、同單位 → 保持連動、單位沿用、數量＝加總 5、成本單價＝合計 ÷ 數量，成本合計（分）不變', m.rel === 'link' && m.forLid === 'a' && m.unit === '人天' && m.qty === 5 && cents([m]) === cents([row({ qty: 2 }), row({ qty: 3, unitCost: 777.77 })]) && !('forLids' in m), J(m));
  m = MG([row({ rel: 'link', forLid: 'a' }), row({ rel: 'link', forLid: 'b' })]);
  t('9c. 連動但不同品項 → 改為拆項（報價不動），forLid／forLids 保留涵蓋判定；數量 1、單位式（舊行為）', m.rel === 'split' && m.forLid === 'a' && same(m.forLids, ['b']) && m.qty === 1 && m.unit === '式');
  m = MG([row({ rel: 'link', forLid: 'a', unit: '人天' }), row({ rel: 'link', forLid: 'a', unit: '人月' })]);
  t('9d. 連動同品項但單位不同 → 改為拆項', m.rel === 'split');
  m = MG([row({ rel: 'link', forLid: 'a' }), row({ rel: 'split', forLid: 'a' })]);
  t('9e. 連動＋拆項混合 → 拆項', m.rel === 'split');
  m = MG([row({ rel: 'none' }), row({ rel: 'none' })]);
  t('9f. 全部不對應 → 不對應（沒有 forLid／forLids）', m.rel === 'none' && !('forLid' in m) && !('forLids' in m));
  m = MG([row({ rel: 'none' }), row({ rel: 'split', forLid: 'a' })]);
  t('9g. 不對應＋拆項混合 → 拆項，而且一定有對應目標（第一列沒有 forLid 時從其他列提升一個當 forLid，否則伺服器會 400）', m.rel === 'split' && m.forLid === 'a' && !m.forLids, J(m));
  m = MG([row({ rel: 'link', forLid: 'a', qty: 0.1 }), row({ rel: 'link', forLid: 'a', qty: 0.2 })]);
  t('9h. 連動合併的數量用十進位加總：0.1＋0.2 ＝ 0.3', m.rel === 'link' && m.qty === 0.3, J(m));
  m = MG([row({ rel: 'link', forLid: 'a', qty: 0 }), row({ rel: 'link', forLid: 'a', qty: 0 })]);
  t('9i. 連動數量加總為 0（算不出單價）→ 退回拆項，不丟錯', m.rel === 'split');
  // consultant／vendor
  m = MG([row({ consultant: '甲' }), row({ consultant: '甲' })]);
  t('9j. 顧問姓名相同 → 保留', m.consultant === '甲');
  m = MG([row({ consultant: '甲' }), row({ consultant: '乙' }), row({ consultant: '甲' })]);
  t('9k. 顧問姓名不同 → 以「、」串接（不重複）', m.consultant === '甲、乙', m.consultant);
  m = MG([row({ consultant: '甲' }), row({})]);
  t('9l. 一列有姓名、一列沒有 → 保留有的', m.consultant === '甲');
  m = MG([row(), row()]);
  t('9m. 都沒有姓名 → 不輸出 consultant 鍵', !('consultant' in m));
  m = MG([row({ consultant: 'a'.repeat(30) }), row({ consultant: 'b'.repeat(30) })]);
  t('9n. 串接後超過 40 字截斷', m.consultant.length === 40);
  m = MG([row({ vendor: '甲' }), row({ vendor: '乙' })]);
  t('9o. 顧問區委外廠商不同 → 以「、」串接（合併後仍是委外）、說明不另記', m.vendor === '甲、乙' && m.note === '' && Q.isOutsourced(m));
  m = MG([row({ vendor: '甲' }), row({})]);
  t('9p. 一列委外、一列自家 → 廠商保留「甲」（整列算委外；確認視窗會提醒）', m.vendor === '甲' && Q.isOutsourced(m));
  m = MG([row({ vendor: 'v'.repeat(40) }), row({ vendor: 'w'.repeat(40) })]);
  t('9q. 廠商串接超過 60 字 → 欄位截 60 字、完整名單記在說明（不丟資訊）', m.vendor.length === 60 && /^廠商：v{40}、w{40}$/.test(m.note), J([m.vendor.length, m.note.length]));
  m = MG([row({ cat: 'software', vendor: '甲' }), row({ cat: 'software', vendor: '乙' })]);
  t('9r. 軟體區（供應商）不同 → 維持改版前行為：欄位留白、說明記「廠商：甲、乙」', m.vendor === '' && m.note === '廠商：甲、乙');
  m = MG([row({ cat: 'software', consultant: '甲' }), row({ cat: 'software', consultant: '乙' })]);
  t('9s. 非顧問區不產生 consultant', !('consultant' in m));
  // 合併前後成本合計不變（隨機）
  const rnd = lcg(31337);
  const pk = (a) => a[Math.floor(rnd() * a.length)];
  let bad = 0, relBad = 0;
  for (let n = 0; n < 2500; n++) {
    const k = 2 + Math.floor(rnd() * 5);
    const target = pk(['a', 'a', 'a', 'b']);
    const ls = Array.from({ length: k }, () => row({ desc: 'D' + Math.floor(rnd() * 9), qty: pk([1, 2, 0.5, 0.1, 3.25, 10]), unitCost: pk([0, 1000, 33.33, 1234.567, 80000]), unit: pk(['人天', '人天', '人天', '人月']), rel: pk(['link', 'link', 'split', 'none', undefined]), forLid: undefined, vendor: pk(['', '甲', '乙']), consultant: pk(['', 'A', 'B']) }));
    ls.forEach((l) => { if (l.rel === 'link' || l.rel === 'split') l.forLid = pk([target, 'a']); });
    const q = Q.normalize(ls);
    const mm = MG(q);
    if (cents(q) !== Q._lineCents(mm) && mm.rel !== 'link') bad++;
    if (mm.rel === 'link' && cents(q) !== Q._lineCents(mm)) bad++;
    if ((mm.rel === 'link' || mm.rel === 'split') && !(mm.forLid || (mm.forLids && mm.forLids.length))) relBad++;
  }
  t('9t. 隨機 2500 組（2–6 列、各種 rel／單位／廠商／姓名）：合併前後成本合計（分）逐組相同（連動合併的單價是反推的，也必須剛好相同）', bad === 0, bad);
  t('9u. 隨機輸入：合併後 rel 為連動／拆項的列一定有對應目標（否則伺服器會 400）', relBad === 0, relBad);
  const txt = Q._mergeConfirmText('consult', [row({ rel: 'link', forLid: 'a' }), row({ rel: 'link', forLid: 'a' })]);
  t('9v. 合併確認文字：說明連動維持、數量為各列加總、顧問姓名也會併', /維持「連動報價數量」/.test(txt) && /各列數量加總/.test(txt) && /顧問姓名/.test(txt), txt);
  const txt2 = Q._mergeConfirmText('consult', [row({ rel: 'link', forLid: 'a' }), row({ rel: 'link', forLid: 'b', vendor: '甲' }), row({})]);
  t('9w. 合併確認文字：連動變拆項要說明；委外與自家混合要提醒委外占比', /改為「拆項（報價不動）」/.test(txt2) && /委外也有自家顧問/.test(txt2), txt2);
}

// ═════════════ 10) 種子、清洗、涵蓋判定 ═════════════
{
  const items = [{ lid: 'T', kind: 'title', desc: 'P' }, it('a', '顧問', 10, '人天'), it('b', '授權', 1, '套', { cat: 'software' }), { nid: 'nid-9', desc: '新增', unit: '式', qty: 2, unitPrice: 0 }, { desc: '沒代碼', qty: 1, unit: '式' }];
  const plain = Q.seedFromItems(items, {});
  t('10a. 沒開 link（業務自填）：種子沒有 rel 鍵（輸出與改版前位元相同）', plain.every((l) => !('rel' in l)) && plain.length >= 3);
  const lk = Q.seedFromItems(items, { link: true });
  const byDesc = (d) => lk.find((l) => l.desc === d);
  t('10b. 開 link（顧問對話框）：品項帶出的列預設「連動」並指向該品項（含用 nid 的新增品項）；固定列（差旅、交際費）預設「不對應」；沒有 lid／nid 的舊資料品項不產生連動列',
    byDesc('顧問').rel === 'link' && byDesc('顧問').forLid === 'a' && byDesc('授權').rel === 'link' && byDesc('新增').forLid === 'nid-9' && byDesc('差旅交通').rel === 'none' && byDesc('交際費').rel === 'none' && !byDesc('沒代碼'), J(lk.map((l) => [l.desc, l.rel, l.forLid])));
  t('10c. mode:consultant 與 link:true 等價', same(Q.seedFromItems(items, { mode: 'consultant' }), lk));
  t('10d. 印花稅固定列沒有 rel', !('rel' in lk.find((l) => l.auto === 'stamp')));
  // 清洗：rel／consultant 與伺服器 normalizeCostLines 逐項相同
  const rnd = lcg(555);
  const pk = (a) => a[Math.floor(rnd() * a.length)];
  let diffs = 0, first = null, cmp = 0;
  for (let n = 0; n < 3000; n++) {
    const raw = Array.from({ length: 1 + Math.floor(rnd() * 6) }, () => {
      const cat = pk(['consult', 'consult', 'software', 'hw', 'travel', 'other']);
      const l = { cat, desc: 'L' + Math.floor(rnd() * 50), vendor: pk(['', '甲', ' 乙 ']), unit: pk(['式', '人天', ' 人月 ']), qty: pk([1, 2, 0.5, 10]), unitCost: pk([0, 100, 12.5]) };
      if (rnd() < 0.7) l.consultant = pk(['', 'Amy', ' Bob ', 'x'.repeat(60), 'a\nb', '   ']);
      const rel = pk(['link', 'split', 'none', undefined, undefined]);
      if (rel) l.rel = rel;
      if (rel === 'link' || rel === 'split') { l.forLid = pk(['a', 'b', ' c ']); if (rnd() < 0.3) l.forLids = [pk(['a', 'd', 'e']), pk(['a', 'd', 'e'])]; }
      else if (rnd() < 0.3) { l.forLid = pk(['a', 'b']); }
      return l;
    });
    let k = 0;
    const sr = CL.normalizeCostLines(raw, { genLid: () => 'g' + (++k) });
    if (!sr.ok) continue;
    cmp++;
    const cr = Q.normalize(raw);
    const proj = (l) => ({ cat: l.cat, desc: l.desc, vendor: l.vendor, consultant: l.consultant, unit: l.unit, qty: l.qty, unitCost: l.unitCost, forLid: l.forLid, forLids: l.forLids, rel: l.rel });
    if (!same(sr.lines.map(proj), cr.map(proj))) { diffs++; if (!first) { const bi = sr.lines.findIndex((l, i) => !same(proj(l), proj(cr[i] || {}))); first = J({ raw: raw[bi], s: proj(sr.lines[bi]), c: proj(cr[bi] || {}) }).slice(0, 700); } }
  }
  t('10e. rel／consultant／forLid／forLids 的清洗與伺服器 normalizeCostLines 逐列相同（' + cmp + ' 組有效隨機輸入：去空白、截斷、非顧問區丟 consultant、none 丟 forLid、forLids 剔除重複）', diffs === 0 && cmp > 2000, diffs + '；' + (first || ''));
  // 涵蓋判定：rel=none 不參與（與伺服器逐組）
  let ud = 0, ufirst = null, withNone = 0;
  for (let n = 0; n < 5000; n++) {
    const its = Array.from({ length: 1 + Math.floor(rnd() * 6) }, (_, i) => (rnd() < 0.1 ? { lid: 'T' + i, kind: 'title', desc: 'P' } : { lid: 'i' + i, desc: pk(['甲', '乙', '甲', '丙', '']), unit: '式', qty: 1, unitPrice: pk([0, 100, 5000]) }));
    const ids = its.filter((x) => !x.kind).map((x) => x.lid).concat(['zz']);
    const ls = Array.from({ length: Math.floor(rnd() * 7) }, () => {
      const l = { cat: pk(['consult', 'travel', 'other']), desc: pk(['甲', '乙', '差旅', '', '丁']), qty: 1, unitCost: 1, unit: '式' };
      const rel = pk(['link', 'split', 'none', 'none', undefined]);
      if (rel) l.rel = rel;
      if (rel === 'link' || rel === 'split') l.forLid = pk(ids); else if (rnd() < 0.3) l.forLid = pk(ids);
      if (rel === 'none') withNone++;
      return l;
    });
    const nl = CL.normalizeCostLines(ls, { genLid: () => 'g' });
    if (!nl.ok) continue;
    const a = Q.unmatchedPricedItems(nl.lines, its), b = CL.unmatchedItems({ items: its, costLines: nl.lines });
    if (!same(a, b)) { ud++; if (!ufirst) ufirst = J({ its, lines: nl.lines, a, b }).slice(0, 600); }
  }
  t('10f. unmatchedPricedItems（含 rel=none 的列不參與涵蓋判定）與伺服器 CL.unmatchedItems：5000 組隨機輸入逐組相同', ud === 0 && withNone > 3000, ud + '；' + (ufirst || '') + ' none=' + withNone);
  t('10g. 已知案例：不對應的列不涵蓋同名品項（甲 有價、一列「甲」rel=none → 仍算未涵蓋）；沒有 rel 的同名列照舊涵蓋',
    Q.unmatchedPricedItems([{ cat: 'consult', desc: '甲', rel: 'none' }], [it('i1', '甲', 1, '式', { unitPrice: 100 })]).length === 1 && Q.unmatchedPricedItems([{ cat: 'consult', desc: '甲' }], [it('i1', '甲', 1, '式', { unitPrice: 100 })]).length === 0);
}

// ═════════════ 11) 顧問對話框輔助函式（quote-approval.js） ═════════════
{
  const qapSrc = fs.readFileSync(QAP_PATH, 'utf8').replace(/\r\n/g, '\n');
  const tail = 'window.refreshQuoteInbox = refreshQuoteInbox;\n})();';
  const hasTail = qapSrc.indexOf(tail) > 0;
  t('11a. quote-approval.js 結尾是預期的 IIFE 收尾（測試掛鉤的位置）', hasTail);
  const hook = qapSrc.replace(tail, 'window.refreshQuoteInbox = refreshQuoteInbox;\nwindow.__cfT = { cfCanSync, cfNewItemsClean, cfEditorItems, cfConsultantNames, cfDraftBody, cfNewItemsProblem, cfDraftQuote, cfDraftFromQuote, costErrMsg, cfNewNid, costRiskValue };\n})();');
  const win = {};
  const dctx = {
    window: win, document: { createElement: () => ({ style: {}, setAttribute() {}, appendChild() {} }), getElementById: () => null, head: { appendChild() {} }, addEventListener() {}, removeEventListener() {} },
    escapeHtml: (v) => String(v), API: '/api', QCL: Q, console, setTimeout, clearTimeout, Date, Math, JSON, Object, Array, Map, Set, Number, String, parseFloat, isFinite,
    quoteCfg: () => ({ costProviders: [{ username: 'u1', displayName: '顧問甲' }, { username: 'u2', displayName: '顧問乙' }] }),
  };
  vm.createContext(dctx);
  let loaded = true;
  try { vm.runInContext(hook, dctx); } catch (err) { loaded = false; t('11b. quote-approval.js 在 vm 載入（有 stub）', false, String(err && err.stack || err).slice(0, 300)); }
  const T = win.__cfT;
  if (loaded && T) {
    t('11b. quote-approval.js 在 vm 載入（有 stub）並取得測試掛鉤', true);
    const q = { perm: { canEditCost: true, canSeePrice: true }, items: [{ lid: 'a', desc: '甲', unit: '式', qty: 1, unitPrice: 100 }, { lid: 'T', kind: 'title', desc: 'P' }], costByName: '我自己', costFlow: { byName: '流程人' }, costLines: [{ cat: 'consult', desc: 'x', consultant: '已填姓名', qty: 1, unitCost: 1 }] };
    t('11c. cfCanSync：可編輯成本、看得到單價、品項有 unitPrice → true；缺任何一項 → false（退回舊畫面）', T.cfCanSync({ q }) === true && T.cfCanSync({ q: Object.assign({}, q, { perm: { canEditCost: true, canSeePrice: false } }) }) === false && T.cfCanSync({ q: Object.assign({}, q, { perm: { canEditCost: false, canSeePrice: true } }) }) === false
      && T.cfCanSync({ q: Object.assign({}, q, { items: [{ lid: 'a', desc: '甲', unit: '式', qty: 1 }] }) }) === false && T.cfCanSync({ q: null }) === false && T.cfCanSync(null) === false);
    const s = { q, newItems: [{ nid: 'n1', desc: '  新增甲  ', unit: ' ', qty: '3' }, { nid: 'n2', desc: '乙', unit: '人天', qty: 'abc' }] };
    const clean = T.cfNewItemsClean(s);
    t('11d. cfNewItemsClean：品名去空白、單位空白當「式」、數量轉數字（非數字當 0）', same(clean, [{ nid: 'n1', desc: '新增甲', unit: '式', qty: 3 }, { nid: 'n2', desc: '乙', unit: '人天', qty: 0 }]), J(clean));
    const ei = T.cfEditorItems(s);
    t('11e. cfEditorItems：業務的品項＋顧問新增的（isDraftNew、needPrice、單價 0、nid 當代號）接在最後', ei.length === 4 && ei[2].nid === 'n1' && ei[2].isDraftNew === true && ei[2].needPrice === true && ei[2].unitPrice === 0 && ei[0] === q.items[0]);
    const names = T.cfConsultantNames(s);
    t('11f. cfConsultantNames：簽核設定的顧問名單顯示名＋自己（q.costByName，config 的 costProviders 不含自己）＋流程人＋已填過的姓名；去重', same(names, ['顧問甲', '顧問乙', '我自己', '流程人', '已填姓名']), J(names));
    s.risk = '10';
    const body = T.cfDraftBody(s, [{ cat: 'consult' }]);
    t('11g. cfDraftBody：costLines＋newItems（只有 nid／desc／unit／qty）＋風險預留（有選才帶）', same(body, { costLines: [{ cat: 'consult' }], newItems: [{ nid: 'n1', desc: '新增甲', unit: '式', qty: 3 }, { nid: 'n2', desc: '乙', unit: '人天', qty: 0 }], contingencyPct: 10 }), J(body));
    delete s.risk;
    t('11h. 沒選風險預留 → 不帶 contingencyPct', !('contingencyPct' in T.cfDraftBody(s, [])));
    t('11i. cfNewItemsProblem：品名空白 → 提示；數量不是有效數字 → 提示；都正常 → 空字串',
      /沒有填品名/.test(T.cfNewItemsProblem({ ov: null, newItems: [{ nid: 'a', desc: ' ', unit: '式', qty: 1 }] }, false)) && /數量不是有效的數字/.test(T.cfNewItemsProblem({ ov: null, newItems: [{ nid: 'a', desc: 'x', unit: '式', qty: 1, qtyBad: true }] }, false)) && T.cfNewItemsProblem({ ov: null, newItems: [{ nid: 'a', desc: 'x', unit: '式', qty: 1 }] }, false) === '');
    const dq = T.cfDraftQuote(q, [{ index: 0, lid: 'a', qty: 7, unit: '人天', isNew: false }, { isNew: true, nid: 'n1', desc: '新增甲', unit: '式', qty: 3 }]);
    t('11j. cfDraftQuote：連動後的數量／單位覆蓋到對應品項（標題列不動）、新增品項接在最後（單價 0、needPrice）；不改原物件', dq.items[0].qty === 7 && dq.items[0].unit === '人天' && dq.items[0].unitPrice === 100 && dq.items[1].kind === 'title' && dq.items[2].desc === '新增甲' && dq.items[2].unitPrice === 0 && dq.items[2].needPrice === true && q.items[0].qty === 1 && q.items.length === 2);
    t('11k. cfDraftFromQuote：伺服器存的草稿新增項目 → 畫面狀態', same(T.cfDraftFromQuote({ costDraft: { newItems: [{ nid: 'n1', desc: 'd', unit: '', qty: 2 }] } }), [{ nid: 'n1', desc: 'd', unit: '式', qty: 2 }]) && same(T.cfDraftFromQuote({}), []));
    const codes = ['LINK_UNIT_CONFLICT', 'LINK_QTY_ZERO', 'TOO_MANY_ITEMS', 'BAD_COST_LINE', 'STALE_ITEMS', 'STALE_COSTS', 'LOCKED_PENDING'];
    t('11l. 錯誤碼對照：新增的 LINK_UNIT_CONFLICT／LINK_QTY_ZERO／TOO_MANY_ITEMS 有人話提示（含伺服器訊息），其他沿用', codes.every((c) => T.costErrMsg({ data: { code: c, error: '伺服器訊息X' }, status: 400 }).length > 5) && /單位不一致/.test(T.costErrMsg({ data: { code: 'LINK_UNIT_CONFLICT', error: 'm' }, status: 400 })) && /大於 0/.test(T.costErrMsg({ data: { code: 'LINK_QTY_ZERO', error: 'm' }, status: 400 })));
    t('11m. cfNewNid：符合伺服器 nid 規則（A-Za-z0-9_.-、≤64、不撞 lid）且每次不同', (() => { const a = T.cfNewNid(), b = T.cfNewNid(); return a !== b && /^[A-Za-z0-9_.-]{1,64}$/.test(a) && a.indexOf('nid-') === 0; })());
  }
}

// ═════════════ 13) 舊行為不退步（對照改版前的 quote-costlines.js） ═════════════
{
  let baseSrc = null;
  try {
    if (process.env.COSTSYNC_QCL_BASE) baseSrc = fs.readFileSync(process.env.COSTSYNC_QCL_BASE, 'utf8');
    else baseSrc = require('child_process').execFileSync('git', ['show', '50fe748:_client/quote-costlines.js'], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'], maxBuffer: 1 << 26 });
  } catch (e) { baseSrc = null; }
  if (!baseSrc) {
    console.log('SKIP 13（找不到改版前的 quote-costlines.js：沒有 git 歷史 50fe748，也沒設 COSTSYNC_QCL_BASE）——這一節不計入');
  } else {
    const bctx = {}; vm.createContext(bctx); vm.runInContext(baseSrc.replace(/\r\n/g, '\n'), bctx);
    const O = bctx.QCL;
    const rnd = lcg(1357);
    const pk = (a) => a[Math.floor(rnd() * a.length)];
    const rItems = () => Array.from({ length: Math.floor(rnd() * 7) }, (_, i) => {
      const k = rnd();
      if (k < 0.1) return { lid: 'T' + i, kind: 'title', desc: 'P' };
      if (k < 0.15) return { lid: 'S' + i, kind: 'subtotal', desc: 's' };
      return Object.assign({ desc: pk(['顧問', '授權', '主機', '', '雜項', '甲']), unit: pk(['人天', '套', '台', '式', '', '次']), qty: pk([1, 2, 0.5, 10, 0, '3', 'x']), unitPrice: pk([0, 100, 5000, 12.5]) }, rnd() < 0.8 ? { lid: 'i' + i } : { nid: 'nid-' + i }, rnd() < 0.2 ? { cat: pk(['consult', 'software', 'hardware', 'other']) } : {}, rnd() < 0.2 ? { cost: pk([0, 1000]) } : {});
    });
    const rLines = (ids) => Array.from({ length: Math.floor(rnd() * 9) }, () => {
      if (rnd() < 0.08) return { cat: 'other', auto: 'stamp' };
      const l = { cat: pk(['consult', 'software', 'hw', 'travel', 'other', 'bogus']), desc: pk(['PM', '授權', '', '差旅', '甲']), vendor: pk(['', '甲', ' 乙 ']), note: pk(['', 'n']), unit: pk(['式', '人天', '']), qty: pk([1, 2, 0.5, 10, 0, '3', 'x', -1]), unitCost: pk([0, 100, 12.5, 33.333, 'y']) };
      if (rnd() < 0.5) l.forLid = pk(ids.concat(['zz']));
      if (rnd() < 0.15) l.forLids = [pk(ids.concat(['zz'])), pk(ids.concat(['zz']))];
      if (rnd() < 0.2) l.lid = 'L' + Math.floor(rnd() * 20);
      return l;
    });
    const cnt = { norm: 0, seed: 0, tot: 0, un: 0, rev: 0, html: 0, merge: 0, msg: 0, miss: 0 };
    const first = {};
    const chk = (k, a, b, ctxInfo) => { if (!same(a, b)) { cnt[k]++; if (!first[k]) first[k] = J({ ctx: ctxInfo, a, b }).slice(0, 500); } };
    const N = 6000;
    for (let n = 0; n < N; n++) {
      const items = rItems();
      const ids = items.filter((x) => !x.kind).map((x, i) => x.lid || x.nid || 'legacy-' + i);
      const lines = rLines(ids);
      const rev = pk([0, 100, 99999.5, 1188880, 5000000]);
      const so = { classCodes: pk([[], ['consult'], ['software'], ['consult', 'software']]), includeStamp: pk([true, false, undefined]) };
      chk('norm', Q.normalize(lines), O.normalize(lines), lines);
      chk('seed', Q.seedFromItems(items, so), O.seedFromItems(items, so), items);
      chk('tot', Q.totals(Q.normalize(lines), rev), O.totals(O.normalize(lines), rev), lines);
      chk('un', [Q.unmatchedPricedItems(lines, items), Q.unmatchedPricedNote(lines, items), Q.unmatchedKeys(lines, items), Q.unmatchedItemsNote(Q.normalize(lines), items, so), Q.unmatchedNewNote(lines, items, [])], [O.unmatchedPricedItems(lines, items), O.unmatchedPricedNote(lines, items), O.unmatchedKeys(lines, items), O.unmatchedItemsNote(O.normalize(lines), items, so), O.unmatchedNewNote(lines, items, [])], { items, lines });
      chk('rev', [Q.revenueOf(items, 'percent', 85), Q.revenueOf(items, 'none', 0), Q.stampAmount(rev)], [O.revenueOf(items, 'percent', 85), O.revenueOf(items, 'none', 0), O.stampAmount(rev)], items);
      chk('miss', Q._missingItemLines(Q.normalize(lines), items, so), O._missingItemLines(O.normalize(lines), items, so), { items, lines });
      const nl = Q.normalize(lines).filter((l) => l.auto !== 'stamp');
      const cat = pk(['software', 'hw', 'travel', 'other']);
      const sameCat = nl.filter((l) => l.cat === cat);
      if (sameCat.length) chk('merge', Q._mergeLines(sameCat), O._mergeLines(sameCat), sameCat);
      // 顧問區：廠商至多一種時合併結果相同（多種廠商串接是這次刻意的改變，見 13e／9o）
      const cons = nl.filter((l) => l.cat === 'consult');
      if (cons.length && new Set(cons.map((l) => l.vendor).filter(Boolean)).size <= 1) chk('merge', Q._mergeLines(cons), O._mergeLines(cons), cons);
      // 非顧問區的列 HTML（編輯／唯讀）
      if (nl[0] && nl[0].cat !== 'consult') chk('html', [Q._editRowHtml(nl[0], nl[0].cat), Q._viewRowHtml(nl[0], nl[0].cat, items)], [O._editRowHtml(nl[0], nl[0].cat), O._viewRowHtml(nl[0], nl[0].cat, items)], nl[0]);
    }
    const labels = { norm: 'normalize', seed: 'seedFromItems', tot: 'totals', un: '涵蓋判定與提醒文字', rev: '營收與印花稅', miss: '補入新品項的候選列', merge: '合併為一列（顧問區廠商至多一種／其他分區）', html: '非顧問區的編輯列與唯讀列 HTML' };
    Object.keys(labels).forEach((k) => t('13' + String.fromCharCode(97 + Object.keys(labels).indexOf(k)) + '. 舊行為不退步：' + labels[k] + '——' + N + ' 組隨機（無 rel／consultant／link）與改版前逐組相同', cnt[k] === 0, cnt[k] + ' 組不同；' + (first[k] || '')));
    // 單列成本：一般範圍內與改版前（浮點取整）相同；改成十進位精確只會在大數值多小數位時更接近伺服器
    let dl = 0;
    for (let n = 0; n < 6000; n++) { const l = { qty: pk([1, 2, 0.5, 3.25, 10, 100, 12.5]), unitCost: pk([0, 1, 1000, 33.33, 12345.67, 0.005, 250000]) }; if (Q._lineCents(l) !== Math.round(O.totals([Object.assign({ cat: 'consult', desc: 'x' }, l)], null).subtotalExStamp * 100)) dl++; }
    t('13i. 單列成本（分）在一般範圍（數量 ≤100、單價 ≤25 萬、≤2 位小數）與改版前的浮點取整相同', dl === 0, dl);
  }
}
// ═════════════ 12) 檔案紀律 ═════════════
{
  const files = ['_client/quote-costlines.js', '_client/quote-approval.js', '_client/quote.js', '_client/quote-preview.js', 'scripts/check-quote-costsync-ui.js'];
  const bad = [];
  const PATH_RE = new RegExp(['[A-Za-z]:' + String.fromCharCode(92) + 'Users', '/Us' + 'ers/[a-z]', 'One' + 'Drive', 'App' + 'Data'].join('|'), 'i');
  files.forEach((f) => {
    const s = fs.readFileSync(path.join(ROOT, f), 'utf8');
    s.split('\n').forEach((line, i) => {
      if (PATH_RE.test(line)) bad.push(f + ':' + (i + 1));
      if (/password\s*[:=]\s*['"][^'"]{4,}|api[_-]?key\s*[:=]\s*['"][^'"]{8,}|secret\s*[:=]\s*['"][^'"]{8,}/i.test(line)) bad.push(f + ':' + (i + 1) + '(secret)');
    });
  });
  t('12a. 新舊檔案都沒有本機路徑／帳密字樣', bad.length === 0, bad.join(','));
  const css = src.slice(src.indexOf('const QCL_CSS'), src.indexOf('function ensureStyle'));
  const sels = (css.match(/(^|\n)\s*([.#][A-Za-z0-9_-]+[^{\n]*)\{/g) || []);
  const offenders = sels.map((x) => x.trim()).filter((x) => x.indexOf('body.dark') !== 0 && !/\.qcl-/.test(x) && !/\.pnl-sum-grid\.qcl-g5/.test(x));
  t('12b. QCL_CSS 新增的選擇器都帶 qcl- 前綴（或 .pnl-sum-grid.qcl-g5）', offenders.length === 0, offenders.slice(0, 3).join(' | '));
  const qapAll = fs.readFileSync(QAP_PATH, 'utf8');
  const qapNew = qapAll.slice(qapAll.indexOf('function cfCanSync'), qapAll.indexOf('let _st = null;'));
  t('12c. cost-sync 新增的程式（QCL 全檔、對話框 cf* 區段）沒有 eval／new Function／document.write／outerHTML', qapNew.length > 5000 && !/\beval\(|new Function\(|document\.write\(|outerHTML/.test(src) && !/\beval\(|new Function\(|document\.write\(|outerHTML/.test(qapNew));
}

let pass = 0, fail = 0;
res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
console.log(`\ncost-sync 前端單元檢查：PASS ${pass} / FAIL ${fail}`);
process.exit(fail ? 1 : 0);
