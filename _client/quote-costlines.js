// ═════════════════════════════════════════════════
// ── 成本明細編輯器 (quote-costlines.js) ───────────────
// 全域 QCL：顧問「填寫成本」對話框與業務「毛利與簽核路徑」分頁共用同一套編輯器與資料模型（q.costLines）。
//   QCL.CATS / QCL.MAX_LINES                五個分區定義、列數上限（含所有分類，與伺服器 MAX_COST_LINES 一致）
//   QCL.seedFromItems(items, opts)          由客戶報價品項產生預設種子（品項各一列＋差旅＋交際費＋印花稅；說明空白的品項略過）
//                                           「＋補入新品項」的覆蓋判定見 missingItemLines（forLid 對得上品項 lid，其餘成本列用品名多重集合比對）；
//                                           按下後先跳確認視窗列出將補入的品項（supplementMessage），成本已含在其他列（整包）時可取消
//   QCL.unmatchedItemsNote(lines, items, opts)  「完成」前的提醒文字：成本明細裡找不到對應列的報價品項（純提醒、不擋；沒有就回空字串）
//   QCL.unmatchedPricedItems(lines, items) / QCL.unmatchedPricedNote(lines, items)
//                                           有價（unitPrice>0）品項沒有成本列涵蓋的名單／提醒文字（＝伺服器 CL.unmatchedItems／preview.warnings 那句；
//                                           業務毛利頁籤常駐提醒、儲存前確認用。涵蓋看 forLid＋forLids，其餘列以品名多重集合比對）
//   QCL.unmatchedKeys(lines, items) / QCL.unmatchedNewNote(lines, items, baseline)
//                                           儲存前確認只看「出現基準以外的未涵蓋品項」：表單載入時用 unmatchedKeys 記基準（未涵蓋有價品項的 lid／nid 陣列），
//                                           儲存時 unmatchedNewNote 只對新出現者回提醒文字；常駐提醒仍用 unmatchedPricedNote 永遠顯示全部
//   （編輯模式每區標題列有「合併為一列」：整包用，成本合計不變，先確認；見 mergeRows／mergeLines；合併後的列帶著所有被併列的 forLid∪forLids）
//   （成本列的 forLid／forLids 全都指向已不存在的品項時，列上顯示淡色徽章「原報價品項已刪除」；mount 沒給 items 就不顯示）
//   QCL.normalize(lines)                    容錯清洗（缺欄補預設、截長度、數字轉型、未知分類丟棄、印花稅列規格化）
//   QCL.totals(lines, revenue)              各分區小計、印花稅估算、合計（元、浮點，僅顯示用；金額以伺服器為準）
//   QCL.revenueOf(items, type, value)       折扣後未稅營收（元）：與伺服器 computeFinancials 的 revenueCents/100 完全一致（逐列取整到分→加總→折扣取整到分）；
//                                           畫面上所有「營收」（印花稅、毛利摘要）都用它，不要用 quote.js 的 quoteTotal 算出的 discounted（未取整的浮點）
//   QCL.mount(el, opts)                     掛載編輯器（mode: 'edit' | 'view'）→ {getLines, setLines, destroy, …}
//   QCL.collect(el)                         從 DOM 讀回 {lines, invalid, blankDesc, zeroCost, zeroQty}
//   QCL.dragSort(container, opts)           拖曳排序核心（Pointer Events 滑鼠／觸控＋鍵盤移動模式）。編輯模式每列有把手（⋮⋮），只能在同一分區內拖；
//                                           報價項目表（quote.js）也用同一個函式。純函式 QCL.dragTargetIndex(rects, y[, fromIdx])／QCL.dragLineY(rects, slot)
//                                           由各列矩形與指標座標算目標位置（見 dragSort 上方說明）。▲▼ 按鈕保留（鍵盤與備用）
//
// 設計原則：
//   · 輸入時只更新小計／合計文字，不重畫表格（輸入框不失焦）；新增／刪除／移動只動該列 DOM。
//   · 所有事件委派在掛載容器上（click／input／change／keydown 各一個），destroy 時一併移除。
//   · 使用者文字（品名／廠商／說明）一律 esc() 後才進 HTML。
//   · 印花稅＝列存在即計入（依合約金額×0.1% 四捨五入至整數元，由伺服器計算），金額一律顯示（規格 v1.1）：
//     有 revenue 就即時重算；沒有 revenue（顧問端）就用伺服器序列化給 stamp 列的 unitCost；兩者都沒有才顯示「系統依合約金額自動計算」。
//     stamp 列的 unitCost 只是「伺服器提供的顯示金額」，collect／getLines 輸出時不送（伺服器忽略並自行計算）。
//   · 樣式由模組內注入一次的 <style id="qclStyle"> 提供，類名一律 qcl- 前綴。
// 資料模型與伺服器端規則見 lib/quoteCostLines.js 檔頭（成本明細獨立於客戶報價品項；欄位存在＝新式、不存在＝舊式 items[].cost）。
// 整個檔案包在 IIFE 內，只暴露 window.QCL。
// ═════════════════════════════════════════════════
(function (global) {
'use strict';

const MAX_LINES = 60;
const MAX_FOR_LIDS = 60;   // forLids（合併為一列後所涵蓋的品項代碼）上限，與伺服器相同
const DESC_MAX = 120, VENDOR_MAX = 60, NOTE_MAX = 200, UNIT_MAX = 10, LID_MAX = 64;
const QTY_MAX = 1e9, COST_MAX = 1e12;
const DEFAULT_UNIT = '式';
const STAMP_DESC = '印花稅(合約金額×0.1%)';
const STAMP_RATE = 0.001;
const SEED_TRAVEL_DESC = '差旅交通';
const SEED_ENTERTAIN_DESC = '交際費';

// 五區（對應 PNL 的 1 顧問成本／2 軟體成本／3 硬體成本／4 差旅／5 其他費用）
const CATS = Object.freeze([
  Object.freeze({ key: 'consult',  name: '顧問服務成本', hasVendor: true,  vendorLabel: '委外廠商', vendorHint: '自有顧問免填', hasNote: true, descHint: '角色（例：PM）' }),
  Object.freeze({ key: 'software', name: '軟體成本',     hasVendor: true,  vendorLabel: '供應商',   vendorHint: '供應商',       hasNote: true, descHint: '品名' }),
  Object.freeze({ key: 'hw',       name: '硬體成本',     hasVendor: true,  vendorLabel: '供應商',   vendorHint: '供應商',       hasNote: true, descHint: '品名' }),
  Object.freeze({ key: 'travel',   name: '差旅費用',     hasVendor: false, vendorLabel: '',         vendorHint: '',             hasNote: true, descHint: '項目（例：差旅交通）' }),
  Object.freeze({ key: 'other',    name: '其他費用',     hasVendor: false, vendorLabel: '',         vendorHint: '',             hasNote: true, descHint: '項目' }),
]);
const CAT_BY_KEY = Object.create(null);
CATS.forEach((c) => { CAT_BY_KEY[c.key] = c; });

const DESC_SUGGEST = ['PM', 'SD', 'MM', 'PP', 'FI', 'CO', 'QM', 'BASIS', '客製', 'GUI/VAT', '電子發票'];
const UNIT_SUGGEST = ['人天', '人月', '式', '台', '組', '套', '授權', '次', '趟'];

// ── 工具 ───────────────────────────────────────────
/** HTML 跳脫（與 app.js 的 escapeHtml 同規則；本模組獨立載入，不依賴它） */
function esc(v) {
  if (v === null || v === undefined) return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function str(v) { return v === null || v === undefined ? '' : String(v); }

/** 文字欄位清洗：換行併成空白、去頭尾空白、截長度 */
function text(v, max) {
  const s = str(v).replace(/[\r\n]+/g, ' ').trim();
  return s.length > max ? s.slice(0, max) : s;
}

/** 數字或數字字串 → number；其餘（含空字串、null、物件）→ NaN */
function toNum(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim() !== '') return Number(v);
  return NaN;
}

/** 轉成 0..max 的有限數；缺值或非數字用 dflt */
function clampNum(v, max, dflt) {
  const n = toNum(v);
  if (!isFinite(n)) return dflt;
  return n < 0 ? 0 : (n > max ? max : n);
}

function fmtMoney(n) {
  const v = Number(n);
  return isFinite(v) ? Math.round(v).toLocaleString('en-US') : '0';
}

function fmtNum(n, maxFrac) {
  const v = Number(n);
  return isFinite(v) ? v.toLocaleString('en-US', { maximumFractionDigits: maxFrac }) : '';
}

/** 去掉浮點誤差（1.005*100 = 100.49999999999999）再四捨五入：與伺服器的精確小數「四捨五入到分」一致 */
function roundClean(x) { return Math.round(+x.toPrecision(12)); }

/** 單列金額（元，未取整，僅顯示用）＝數量×單價 */
function lineAmount(l) {
  const q = clampNum(l && l.qty, QTY_MAX, 0), c = clampNum(l && l.unitCost, COST_MAX, 0);
  return +(q * c).toPrecision(12);
}

/** 單列金額（分，整數）：與伺服器同規則——每列先取整到分再加總 */
function lineCents(l) {
  const q = clampNum(l && l.qty, QTY_MAX, 0), c = clampNum(l && l.unitCost, COST_MAX, 0);
  return roundClean(q * c * 100);
}

/**
 * 印花稅估算＝整數元 round-half-up(折扣後未稅營收×0.1%)；營收未知（空、非數字、負數）回 null。
 * 與伺服器同規則：營收先取整到「分」（QCL.revenueOf 算出來的已經是分的整數倍），再用整數運算 round-half-up(分 / 100000)，
 * 不經過浮點的 ×0.001（營收很大時 toPrecision 會吃掉小數位）。
 */
function stampAmount(revenue) {
  const r = toNum(revenue);
  if (!isFinite(r) || r < 0) return null;
  const cents = Math.round(+(r * 100).toPrecision(15));
  if (!(cents + 50000 <= 9007199254740991)) return roundClean(r * STAMP_RATE);   // 超過 2^53 分（約 90 兆元）：已超出任何實際報價，退回浮點估算
  const n = cents + 50000;
  return (n - (n % 100000)) / 100000;
}

// ── 營收（折扣後未稅，元）：與伺服器 computeFinancials 的 revenueCents/100 完全一致 ─────────────
// 伺服器規則（lib/quoteRoutes.js normalizeItems／儲存折扣 + lib/quoteApproval.js computeFinancials）：
//   1. 品項（標題／小計列不計）：數量 parseFloat 後「有限且 > 0」才採用（至少 0.001），否則當 1；單價「有限且 ≥ 0」才採用，否則 0。
//   2. 每列以十進位精確運算「取整到分（half-up）」＝ round(數量 × 單價 × 100)，再加總成小計（分）。
//   3. 折扣：折扣值 ≤ 0 或類型不是 percent／amount＝無折扣；percent → round-half-up(小計分 × 百分比 / 100)；amount → round-half-up(議價元 × 100)。
//      伺服器對範圍外的折扣（百分比 ≥ 100、議價 ≥ 小計）會拒絕儲存；這裡不拒絕，沿用同一條公式（畫面輸入到一半不要跳掉，存檔時由畫面與伺服器擋下）。
// 用 BigInt 做精確十進位運算（浮點的 1.005×100＝100.49999999999999 會差 1 分，印花稅與成本就會差 1 元）；
// 不支援 BigInt 的舊瀏覽器退回浮點估算（不丟錯）。BigInt 一律用 BigInt(n) 呼叫、不寫 123n 字面值，舊瀏覽器載入本檔才不會語法錯誤。
const HAS_BIGINT = typeof BigInt === 'function';
const POW10 = [];
function pow10(n) {
  if (!POW10.length) POW10.push(BigInt(1));
  while (POW10.length <= n) POW10.push(POW10[POW10.length - 1] * BigInt(10));
  return POW10[n];
}
/** 四捨五入（half-up），num／den 皆為非負 BigInt，den > 0 */
function divRound(num, den) { return (BigInt(2) * num + den) / (BigInt(2) * den); }

/** 有限非負 number → { n: BigInt（去掉小數點的整數）, s: 小數位數 }；不能解析（非有限、負數、指數過大）回 null。鏡像 quoteApproval.parseDec 的 number 分支 */
function decParts(x) {
  if (typeof x !== 'number' || !isFinite(x) || x < 0) return null;
  const m = /^(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(String(x));
  if (!m) return null;
  const intPart = m[1] || '', fracPart = m[2] || '';
  if (intPart === '' && fracPart === '') return null;
  const exp = m[3] ? parseInt(m[3], 10) : 0;
  if (exp > 30) return null;
  const digits = (intPart + fracPart).replace(/^0+(?=\d)/, '');
  let n = BigInt(digits === '' ? '0' : digits);
  let s = fracPart.length - exp;
  if (s < 0) { n = n * pow10(-s); s = 0; }
  if (s > 400) { n = BigInt(0); s = 0; }
  return { n, s };
}

/** 品項數量／單價套用伺服器的儲存規則（normalizeItems） */
function storedQty(v) { const q = parseFloat(v); return isFinite(q) && q > 0 ? Math.max(0.001, q) : 1; }
function storedPrice(v) { const p = parseFloat(v); return isFinite(p) && p >= 0 ? p : 0; }
function isKindRow(it) { return !!it && (it.kind === 'title' || it.kind === 'subtotal'); }

function revenueOfFloat(items, discountType, discountValue) {
  let sub = 0;
  (Array.isArray(items) ? items : []).forEach((it) => {
    if (!it || typeof it !== 'object' || isKindRow(it)) return;
    sub += roundClean(storedQty(it.qty) * storedPrice(it.unitPrice) * 100);
  });
  let rev = sub;
  const dv = parseFloat(discountValue);
  if ((discountType === 'percent' || discountType === 'amount') && isFinite(dv) && dv > 0) {
    rev = discountType === 'percent' ? roundClean(sub * dv / 100) : roundClean(dv * 100);
  }
  return rev / 100;
}

/**
 * 折扣後未稅營收（元，必為「分」的整數倍）。items：{qty, unitPrice, kind?}[]（畫面 readQuoteItems 的結果或伺服器序列化的品項）；
 * discountType：'none'|'percent'|'amount'；discountValue：百分比（90＝九折）或議價總額（元）。印花稅、成本、毛利都用這個營收。
 */
function revenueOf(items, discountType, discountValue) {
  if (!HAS_BIGINT) return revenueOfFloat(items, discountType, discountValue);
  const H = BigInt(100);
  let sub = BigInt(0);
  (Array.isArray(items) ? items : []).forEach((it) => {
    if (!it || typeof it !== 'object' || isKindRow(it)) return;
    const qty = decParts(storedQty(it.qty)), price = decParts(storedPrice(it.unitPrice));
    if (!qty || !price) return;   // 數字大到無法處理（伺服器會以 TOO_LARGE 拒絕）：該列不計
    sub += divRound(qty.n * price.n * H, pow10(qty.s + price.s));
  });
  let rev = sub;
  const dv = parseFloat(discountValue);
  if ((discountType === 'percent' || discountType === 'amount') && isFinite(dv) && dv > 0) {
    const v = decParts(dv);
    if (v) rev = discountType === 'percent' ? divRound(sub * v.n, H * pow10(v.s)) : divRound(v.n * H, pow10(v.s));
  }
  return Number(rev) / 100;
}

// ── 資料模型：清洗／種子／合計 ─────────────────────────
/**
 * 單列清洗；未知分類回 null。印花稅列依規格強制 cat/qty/unit/desc；
 * unitCost 只在來源有給有限數字時才保留（＝伺服器提供的印花稅金額，供顯示與 totals 使用；沒給就沒有這個欄位，與「金額 0」區分）。
 */
function cleanLine(raw) {
  if (!raw || typeof raw !== 'object') return null;
  const stamp = raw.auto === 'stamp';
  const cat = stamp ? 'other' : str(raw.cat);
  const def = CAT_BY_KEY[cat];
  if (!def) return null;
  const o = {};
  const lid = text(raw.lid, LID_MAX);
  if (lid) o.lid = lid;
  o.cat = cat;
  if (stamp) {
    o.desc = STAMP_DESC;
    o.vendor = '';
    o.note = text(raw.note, NOTE_MAX);
    o.unit = DEFAULT_UNIT;
    o.qty = 1;
    const given = toNum(raw.unitCost);
    if (isFinite(given)) o.unitCost = given < 0 ? 0 : (given > COST_MAX ? COST_MAX : given);
    o.auto = 'stamp';
  } else {
    o.desc = text(raw.desc, DESC_MAX);
    o.vendor = def.hasVendor ? text(raw.vendor, VENDOR_MAX) : '';   // 差旅、其他沒有廠商欄
    o.note = text(raw.note, NOTE_MAX);
    o.unit = text(raw.unit, UNIT_MAX) || DEFAULT_UNIT;
    o.qty = clampNum(raw.qty, QTY_MAX, 1);
    o.unitCost = clampNum(raw.unitCost, COST_MAX, 0);
  }
  const forLid = text(raw.forLid, LID_MAX);
  if (forLid) o.forLid = forLid;
  if (!stamp) {
    // forLids：合併為一列時帶著被併各列的 forLid∪forLids，只供涵蓋判定（與伺服器同規則：字串、去頭尾空白、截 64、去空白與重複、剔除與 forLid 相同者、上限 60）
    const fl = [];
    if (Array.isArray(raw.forLids)) {
      const seen = Object.create(null);
      for (let i = 0; i < raw.forLids.length && fl.length < MAX_FOR_LIDS; i++) {
        const e = raw.forLids[i];
        if (typeof e !== 'string') continue;
        const s = e.trim().slice(0, LID_MAX);
        if (s && s !== forLid && !seen[s]) { seen[s] = true; fl.push(s); }
      }
    }
    if (fl.length) o.forLids = fl;
  }
  return o;
}

/** 容錯清洗整個陣列：丟棄無效列、同一單只留第一個印花稅列、超過上限的列截掉；回傳新陣列 */
function normalize(lines) {
  const out = [];
  let hasStamp = false;
  (Array.isArray(lines) ? lines : []).forEach((r) => {
    if (out.length >= MAX_LINES) return;
    const l = cleanLine(r);
    if (!l) return;
    if (l.auto === 'stamp') {
      if (hasStamp) return;
      hasStamp = true;
    }
    out.push(l);
  });
  return out;
}

/** 印花稅金額來源（規格 v1.1）：有 revenue 就即時重算；否則用伺服器給 stamp 列的 unitCost；兩者都沒有回 null */
function resolveStamp(revenue, serverAmt) {
  const live = stampAmount(revenue);
  if (live !== null) return live;
  return typeof serverAmt === 'number' && isFinite(serverAmt) ? serverAmt : null;
}

/**
 * 各分區小計（元）、印花稅、合計。byCat.other 不含印花稅。
 * 印花稅：revenue 有值 → 依營收重算；否則用 stamp 列的 unitCost（伺服器提供）；兩者都沒有才不計並標 stampUnknown。
 */
function totals(lines, revenue) {
  const cents = { consult: 0, software: 0, hw: 0, travel: 0, other: 0 };
  let hasStamp = false;
  let serverAmt = null;
  (Array.isArray(lines) ? lines : []).forEach((r) => {
    const l = cleanLine(r);
    if (!l) return;
    if (l.auto === 'stamp') {
      if (!hasStamp && typeof l.unitCost === 'number') serverAmt = l.unitCost;   // 只認第一個印花稅列
      hasStamp = true;
      return;
    }
    cents[l.cat] += lineCents(l);
  });
  const amt = resolveStamp(revenue, serverAmt);
  const stamp = hasStamp && amt !== null ? amt : 0;
  const byCat = {};
  let sum = 0;
  CATS.forEach((c) => { byCat[c.key] = cents[c.key] / 100; sum += cents[c.key]; });
  return {
    byCat,
    stamp,
    stampUnknown: hasStamp && amt === null,
    subtotalExStamp: sum / 100,
    total: (sum + stamp * 100) / 100,
  };
}

// 毛利分類猜測：鏡像 lib/quotePnlExcel.js 的 catByUnit／resolveCat（鍵名硬體改用 hw）
const ITEM_CAT = Object.assign(Object.create(null), { consult: 'consult', software: 'software', hardware: 'hw', hw: 'hw', other: 'other' });
const CLASS_TO_CAT = Object.assign(Object.create(null), { consult: 'consult', software: 'software', hardware: 'hw', crm: 'other', mdm: 'other', ot: 'other', other: 'other' });

function catByUnit(unit) {
  const u = str(unit);
  if (/人天|人日|人月|man.?day|\bMD\b|^天$|^日$/i.test(u)) return 'consult';
  if (/^(台|組|部|臺|pcs|set|unit)$/i.test(u)) return 'hw';
  if (/授權|license|套|user|seat|帳號|訂閱|subscription/i.test(u)) return 'software';
  return 'other';
}

/** 品項所屬成本分區：品項自己的 cat → 勾選商品只涵蓋單一分區就全歸該區 → 依單位猜 */
function itemCat(it, classCats) {
  const own = ITEM_CAT[str(it.cat)];
  if (own) return own;
  if (classCats.length === 1) return classCats[0];
  return catByUnit(it.unit);
}

/**
 * 只含「客戶品項」那幾列的種子（略過分組標題／小計列；品項沒有 unitPrice 欄位也可以）。
 * 說明（desc）為空白（去頭尾空白後是空字串）的品項也略過：伺服器與業務表單都不要求品項說明，空白品項列很容易留在單上，
 * 但成本明細的「項目」是必填——帶成空白項目列，之後儲存／完成都會被 blankDesc 擋下，使用者只好自己刪列。
 * 種子、重新帶入、補入新品項、顧問對話框、業務毛利頁籤全部走這個函式，所以規則一致（業務端 quote.js 的 _qClItems 也是同一條規則）。
 */
function seedItemLines(items, opts) {
  const set = new Set();
  ((opts && Array.isArray(opts.classCodes)) ? opts.classCodes : []).forEach((c) => { const k = CLASS_TO_CAT[str(c)]; if (k) set.add(k); });
  const classCats = Array.from(set);
  const out = [];
  (Array.isArray(items) ? items : []).forEach((it) => {
    if (!it || typeof it !== 'object' || it.kind === 'title' || it.kind === 'subtotal') return;
    const q = toNum(it.qty);
    const line = cleanLine({
      cat: itemCat(it, classCats),
      desc: it.desc,
      unit: it.unit,
      qty: q > 0 ? q : 1,                  // 數量空／0 當 1（與伺服器、Excel 一致）
      unitCost: it.cost,                   // 舊式單的 items[].cost 帶入；成本與報價單價各自獨立（顧問看不到品項單價），所以不讀 unitPrice
      forLid: it.lid || it.nid,            // 還沒存檔的新品項沒有 lid：用畫面給的暫時代號 nid 指向它，存檔時伺服器換成真正的 lid（見 lib/quoteCostLines.js resolveItemRefs）
    });
    if (line && line.desc) out.push(line);
  });
  return out;
}

/** 固定種子列：差旅「差旅交通」1 列＋其他「交際費」＋（includeStamp）印花稅 auto 列 */
function seedFixedLines(includeStamp) {
  const out = [
    cleanLine({ cat: 'travel', desc: SEED_TRAVEL_DESC }),
    cleanLine({ cat: 'other', desc: SEED_ENTERTAIN_DESC }),
  ];
  if (includeStamp) out.push(cleanLine({ cat: 'other', auto: 'stamp' }));
  return out;
}

/**
 * 預設種子（前端產生，儲存前不進伺服器）。
 * opts.classCodes：勾選商品的類別代碼（選填，用來鏡像伺服器「單一類別就全歸該區」的規則）。
 * opts.includeStamp：預設 true＝種入印花稅 auto 列（列存在＝計入）。false＝不種入（編輯器的「印花稅」核取方塊呈未勾選）：
 *   給「已存在、還沒有成本明細的舊式單」用——舊式單沒動編輯器就存檔不會帶成本明細，伺服器的成本不含印花稅；畫面種子也不含，
 *   兩邊毛利才一致；使用者勾選印花稅（＝動了編輯器）存檔才轉成新式。
 * opts.revenueKnown：保留參數——印花稅列沒有金額欄位，金額由 mount 的 revenue 即時算或伺服器序列化時給。
 */
function seedFromItems(items, opts) {
  const fixed = seedFixedLines(!(opts && opts.includeStamp === false));
  return seedItemLines(items, opts).slice(0, MAX_LINES - fixed.length).concat(fixed);
}

// ── 樣式（注入一次）─────────────────────────────────
const QCL_CSS = `
.qcl-root, .qcl-root * { box-sizing: border-box; }
.qcl-root { --qcl-bg: #fff; --qcl-bd: #e3e6ea; --qcl-th: #f6f7f9; --qcl-tx: #222; --qcl-h: #111; --qcl-mu: #6b7684;
  --qcl-pri: #1a73e8; --qcl-pri-bg: #e8f0fe; --qcl-in-bd: #d9dde2; --qcl-in-bg: #fff;
  --qcl-bad: #ea4335; --qcl-bad-bg: #fde8e8; --qcl-warn: #8a4b00; --qcl-warn-bg: #fff4e5; --qcl-warn-bd: #f5c98b; --qcl-del: #e53935;
  font-size: 13px; color: var(--qcl-tx); text-align: left; }
body.dark .qcl-root { --qcl-bg: #161b22; --qcl-bd: #30363d; --qcl-th: #21262d; --qcl-tx: #e6edf3; --qcl-h: #e6edf3; --qcl-mu: #8b949e;
  --qcl-pri: #58a6ff; --qcl-pri-bg: #14283f; --qcl-in-bd: #30363d; --qcl-in-bg: #0d1117;
  --qcl-bad: #ff8a80; --qcl-bad-bg: #3f1717; --qcl-warn: #f0b866; --qcl-warn-bg: #3d2b0a; --qcl-warn-bd: #6b4a14; --qcl-del: #ff8a80; }
.qcl-tools { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 8px; }
.qcl-sp { flex: 1; }
.qcl-count { color: var(--qcl-mu); font-size: 12.5px; }
.qcl-hint { color: var(--qcl-mu); font-size: 12.5px; margin: 0 0 8px; line-height: 1.6; }
.qcl-limit { border-radius: 8px; padding: 7px 12px; font-size: 13px; margin-bottom: 8px; line-height: 1.6;
  background: var(--qcl-warn-bg); color: var(--qcl-warn); border: 1px solid var(--qcl-warn-bd); }
.qcl-limit[hidden] { display: none; }
.qcl-msg { min-height: 0; color: var(--qcl-pri); font-size: 12.5px; margin-bottom: 6px; }
.qcl-msg:empty { display: none; }
.qcl-btn { border: 1px solid var(--qcl-in-bd); background: var(--qcl-in-bg); color: var(--qcl-tx); border-radius: 8px; padding: 5px 12px;
  font-size: 13px; font-family: inherit; cursor: pointer; line-height: 1.4; }
.qcl-btn:hover:not(:disabled) { border-color: var(--qcl-pri); color: var(--qcl-pri); }
.qcl-btn:disabled { opacity: .5; cursor: not-allowed; }
.qcl-btn.sm { padding: 3px 10px; font-size: 12.5px; }
.qcl-sec { background: var(--qcl-bg); border: 1px solid var(--qcl-bd); border-radius: 10px; padding: 10px 12px; margin-bottom: 12px; }
.qcl-sec-h { display: flex; flex-wrap: wrap; align-items: center; gap: 4px 8px; margin-bottom: 6px; }
.qcl-sec-h h4 { margin: 0; font-size: 14px; font-weight: 700; color: var(--qcl-h); }
.qcl-sec-n { color: var(--qcl-mu); font-size: 12px; }
.qcl-sec-tot { display: none; font-weight: 700; font-size: 13px; color: var(--qcl-pri); font-variant-numeric: tabular-nums; }
.qcl-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
.qcl-table { width: 100%; min-width: 890px; table-layout: fixed; border-collapse: collapse; font-size: 13px; }
.qcl-table.nv { min-width: 790px; }
.qcl-table col.qcl-c-unit { width: 84px; }
.qcl-table col.qcl-c-qty { width: 88px; }
.qcl-table col.qcl-c-cost { width: 120px; }
.qcl-table col.qcl-c-sub { width: 112px; }
.qcl-table col.qcl-c-ctl { width: 126px; }
.qcl-table th, .qcl-table td { border-bottom: 1px solid var(--qcl-bd); padding: 5px 6px; text-align: left; vertical-align: middle; }
.qcl-table th { background: var(--qcl-th); font-weight: 600; white-space: nowrap; color: var(--qcl-tx); }
.qcl-table th.r, .qcl-table td.r { text-align: right; }
.qcl-table td.r { white-space: nowrap; font-variant-numeric: tabular-nums; }
.qcl-table td.qcl-t { white-space: pre-wrap; word-break: break-word; }
.qcl-table td.qcl-mu, .qcl-mu { color: var(--qcl-mu); }
.qcl-table th.qcl-ctl-h, .qcl-table td.qcl-ctl { text-align: center; white-space: nowrap; }
.qcl-in { width: 100%; padding: 5px 8px; border: 1px solid var(--qcl-in-bd); border-radius: 6px; font-size: 13px; font-family: inherit;
  background: var(--qcl-in-bg); color: var(--qcl-tx); outline: none; }
.qcl-in:focus { border-color: var(--qcl-pri); }
.qcl-in.qcl-bad { border-color: var(--qcl-bad); background: var(--qcl-bad-bg); }
.qcl-in.qcl-warn { border-color: var(--qcl-warn-bd); background: var(--qcl-warn-bg); }
.qcl-unit { text-align: center; }
.qcl-qty, .qcl-cost { text-align: right; }
.qcl-qty, .qcl-cost { -moz-appearance: textfield; }
.qcl-qty::-webkit-outer-spin-button, .qcl-qty::-webkit-inner-spin-button,
.qcl-cost::-webkit-outer-spin-button, .qcl-cost::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
.qcl-mv { background: none; border: none; cursor: pointer; line-height: 1; padding: 3px 4px; font-size: 13px; color: var(--qcl-mu); }
.qcl-mv:hover:not(:disabled) { color: var(--qcl-pri); }
.qcl-mv:disabled { opacity: .3; cursor: default; }
.qcl-mv.del { color: var(--qcl-del); font-size: 16px; }
/* 拖曳把手（成本明細每列、報價項目表每列共用同一組 qcl-dnd-* 拖曳樣式；把手本身：成本明細用 qcl-drag、報價項目表用 qi-drag） */
.qcl-drag { background: none; border: none; cursor: grab; touch-action: none; -webkit-user-select: none; user-select: none; -webkit-touch-callout: none;
  min-width: 22px; height: 26px; padding: 0 3px; margin-right: 2px; font-size: 14px; line-height: 1; letter-spacing: -2px; color: var(--qcl-mu); border-radius: 4px; vertical-align: middle; }
.qcl-drag:hover { color: var(--qcl-pri); background: var(--qcl-pri-bg); }
.qcl-drag:focus-visible { outline: 2px solid var(--qcl-pri); outline-offset: 1px; }
.qcl-dnd-src { opacity: .35; }
.qcl-dnd-ghost { position: fixed; z-index: 100000; left: 0; top: 0; display: flex; align-items: center; gap: 8px; padding: 6px 12px; box-sizing: border-box;
  background: #fff; color: #222; border: 1px solid #1a73e8; border-radius: 8px; box-shadow: 0 8px 24px rgba(0, 0, 0, .25); opacity: .93;
  font-size: 13px; font-family: inherit; line-height: 1.4; white-space: nowrap; pointer-events: none; }
.qcl-dnd-ghost .qcl-dnd-grip { color: #1a73e8; letter-spacing: -2px; flex: none; }
.qcl-dnd-ghost .qcl-dnd-txt { min-width: 0; overflow: hidden; text-overflow: ellipsis; }
.qcl-dnd-ghost.qcl-dnd-bad { border-color: #ea4335; opacity: .6; }
.qcl-dnd-line { position: fixed; z-index: 100000; left: 0; top: 0; width: 0; height: 3px; border-radius: 2px; background: #1a73e8; pointer-events: none; display: none; }
.qcl-dnd-line::before { content: ""; position: absolute; left: -3px; top: -2px; width: 7px; height: 7px; border-radius: 50%; background: #1a73e8; }
body.dark .qcl-dnd-ghost { background: #161b22; color: #e6edf3; border-color: #58a6ff; box-shadow: 0 8px 24px rgba(0, 0, 0, .6); }
body.dark .qcl-dnd-ghost .qcl-dnd-grip { color: #58a6ff; }
body.dark .qcl-dnd-ghost.qcl-dnd-bad { border-color: #ff8a80; }
body.dark .qcl-dnd-line, body.dark .qcl-dnd-line::before { background: #58a6ff; }
body.qcl-dnd-on, body.qcl-dnd-on * { -webkit-user-select: none !important; user-select: none !important; cursor: grabbing !important; }
.qcl-sr { position: absolute !important; width: 1px; height: 1px; margin: -1px; padding: 0; border: 0; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }
.qcl-sec { container-type: inline-size; }   /* 讓 100cqw ＝ 分區內容寬（窄螢幕空狀態文字寬度用） */
.qcl-empty td { color: var(--qcl-mu); text-align: center; padding: 10px; white-space: normal; }
.qcl-emptytxt { display: block; white-space: normal; overflow-wrap: anywhere; line-height: 1.6; }
.qcl-orph { display: block; width: fit-content; max-width: 100%; margin-top: 3px; padding: 0 7px; font-size: 11.5px; line-height: 1.6; color: var(--qcl-mu);
  background: var(--qcl-th); border: 1px solid var(--qcl-bd); border-radius: 9px; white-space: nowrap; }
.qcl-orph[hidden] { display: none; }
.qcl-t .qcl-orph { display: inline-block; margin: 0 0 0 6px; }
.qcl-stamprow td { background: var(--qcl-th); }
.qcl-stamprow label { cursor: pointer; margin: 0 6px; }
.qcl-stamprow input[type=checkbox] { margin: 0 4px 0 0; vertical-align: -2px; }
.qcl-stamp-name { font-weight: 600; }
.qcl-table tfoot td { font-weight: 700; background: var(--qcl-th); }
.qcl-grand { display: flex; flex-wrap: wrap; align-items: baseline; gap: 4px 12px; padding: 4px 4px 0; }
.qcl-grand-k { font-weight: 700; color: var(--qcl-h); }
.qcl-grand-v { font-size: 18px; font-weight: 700; color: var(--qcl-pri); font-variant-numeric: tabular-nums; }
.qcl-grand-n { color: var(--qcl-mu); font-size: 12.5px; }
@media (max-width: 640px) {
  .qcl-sec { padding: 8px; }
  /* 表格最小寬度 790/890px 要橫向捲動：空分區的提示文字改成靠左、釘在可視區左緣、寬度＝可視寬並自動換行，不必捲動就看得到 */
  .qcl-empty td { text-align: left; }
  .qcl-emptytxt { position: sticky; left: 10px; width: calc(100vw - 96px); width: calc(100cqw - 20px); }
  .qcl-sec-tot { display: inline; }   /* 表格要橫向捲動才看得到小計欄，所以窄螢幕在區標題多放一份 */
  .qcl-tools .qcl-btn { flex: 1 1 auto; }
}
`;

function ensureStyle() {
  const doc = global.document;
  if (!doc || !doc.getElementById || !doc.createElement) return;   // 沒有 DOM（單元測試的 vm）就略過
  if (doc.getElementById('qclStyle')) return;
  const st = doc.createElement('style');
  st.id = 'qclStyle';
  st.textContent = QCL_CSS;
  doc.head.appendChild(st);
}

// ── 畫面：HTML 產生 ─────────────────────────────────
let uid = 0;
const INST = new WeakMap();   // 掛載容器 → 實例

function emptyLine(cat) { return cleanLine({ cat }); }

// 「原報價品項已刪除」徽章（淡色小字）：成本列有 forLid／forLids、但指向的品項全都不在目前品項清單時顯示（只提示，不改計算；成本仍計入）。
// 編輯模式每列都先放一個隱藏的徽章，refresh 依 items 切換；唯讀模式直接依 items 產生。items 沒提供就一律不顯示。
const ORPHAN_TEXT = '原報價品項已刪除';
const ORPHAN_TITLE = '這列是由已被刪除的報價品項帶入；成本仍計入合計，請確認是否保留（不需要請按 ✕ 移除）';
const ORPHAN_BADGE_HIDDEN = '<span class="qcl-orph" hidden title="' + ORPHAN_TITLE + '">' + ORPHAN_TEXT + '</span>';
const ORPHAN_BADGE = '<span class="qcl-orph" title="' + ORPHAN_TITLE + '">' + ORPHAN_TEXT + '</span>';

/** 目前一般品項的代碼集合（items 沒提供回 null）；沒有 lid 的品項以 'legacy-<索引>' 代替（與伺服器 serialize 給的值一致） */
function itemLidSet(items) {
  return Array.isArray(items) ? new Set(generalEntries(items).map((e) => e.lid)) : null;
}
/** 這列是否「原報價品項已刪除」：有 forLid／forLids，但沒有任何一個還在目前品項清單裡；lidSet 為 null（沒提供 items）一律 false */
function isOrphanLine(l, lidSet) {
  if (!lidSet || !l || l.auto === 'stamp') return false;
  const ids = lineLidList(l);
  return ids.length > 0 && !ids.some((k) => lidSet.has(k));
}

function editRowHtml(l, def, ids) {
  const attrs = (l.lid ? ' data-lid="' + esc(l.lid) + '"' : '') + (l.forLid ? ' data-forlid="' + esc(l.forLid) + '"' : '') +
    (l.forLids && l.forLids.length ? ' data-forlids="' + esc(JSON.stringify(l.forLids)) + '"' : '');
  const lab = (label) => ' aria-label="' + esc(def.name + '：' + label) + '"';
  return '<tr class="qcl-row"' + attrs + '>' +
    '<td><input type="text" class="qcl-in qcl-desc" maxlength="' + DESC_MAX + '" value="' + esc(l.desc) + '" placeholder="' + esc(def.descHint) + '"' +
      (def.key === 'consult' ? ' list="' + ids.desc + '"' : '') + lab('項目') + '>' + ORPHAN_BADGE_HIDDEN + '</td>' +
    (def.hasVendor
      ? '<td><input type="text" class="qcl-in qcl-vendor" maxlength="' + VENDOR_MAX + '" value="' + esc(l.vendor) + '" placeholder="' + esc(def.vendorHint) + '"' + lab(def.vendorLabel) + '></td>'
      : '') +
    '<td><input type="text" class="qcl-in qcl-note" maxlength="' + NOTE_MAX + '" value="' + esc(l.note) + '" placeholder="說明（選填）"' + lab('說明') + '></td>' +
    '<td><input type="text" class="qcl-in qcl-unit" maxlength="' + UNIT_MAX + '" value="' + esc(l.unit) + '" list="' + ids.unit + '"' + lab('單位') + '></td>' +
    '<td><input type="number" class="qcl-in qcl-qty" min="0" step="any" inputmode="decimal" value="' + esc(String(l.qty)) + '"' + lab('數量') + '></td>' +
    '<td><input type="number" class="qcl-in qcl-cost" min="0" step="any" inputmode="decimal" placeholder="0" value="' + (l.unitCost ? esc(String(l.unitCost)) : '') + '"' + lab('成本單價') + '></td>' +
    '<td class="r qcl-sub">' + fmtMoney(lineAmount(l)) + '</td>' +
    '<td class="qcl-ctl">' +
      '<button type="button" class="qcl-drag" title="拖曳排序（區內）；也可用鍵盤：按空白鍵拿起，上下鍵移動，空白鍵放下，Esc 取消" aria-label="拖曳排序">&#8942;&#8942;</button>' +
      '<button type="button" class="qcl-mv" data-act="up" title="上移" aria-label="上移">&#9650;</button>' +
      '<button type="button" class="qcl-mv" data-act="down" title="下移" aria-label="下移">&#9660;</button>' +
      '<button type="button" class="qcl-mv del" data-act="del" title="移除此列" aria-label="移除此列">&#10005;</button></td></tr>';
}

/** 欄寬定義（fixed layout）：文字欄均分剩餘寬度，單位／數量／單價／小計／控制欄固定寬，各分區的數字欄上下對齊 */
function colGroup(def, withCtl) {
  return '<colgroup><col>' + (def.hasVendor ? '<col>' : '') + '<col><col class="qcl-c-unit"><col class="qcl-c-qty"><col class="qcl-c-cost"><col class="qcl-c-sub">' +
    (withCtl ? '<col class="qcl-c-ctl">' : '') + '</colgroup>';
}

function tableHead(def) {
  return '<thead><tr><th>項目</th>' + (def.hasVendor ? '<th>' + esc(def.vendorLabel) + '</th>' : '') +
    '<th>說明</th><th>單位</th><th class="r">數量</th><th class="r">成本單價</th><th class="r">小計</th><th class="qcl-ctl-h"></th></tr></thead>';
}

function editSectionHtml(def, rows, stamp, ids) {
  const cols = def.hasVendor ? 6 : 5;
  let h = '<section class="qcl-sec" data-sec="' + def.key + '">' +
    '<div class="qcl-sec-h"><h4>' + esc(def.name) + '</h4><span class="qcl-sec-n"></span><span class="qcl-sec-tot"></span><span class="qcl-sp"></span>' +
    '<button type="button" class="qcl-btn sm" data-act="merge" data-cat="' + def.key + '" title="把這一區的所有列合併成 1 列（成本合計不變），例如整包" disabled>合併為一列</button>' +
    '<button type="button" class="qcl-btn sm" data-act="add" data-cat="' + def.key + '">＋新增</button></div>' +
    '<div class="qcl-wrap"><table class="qcl-table' + (def.hasVendor ? '' : ' nv') + '">' + colGroup(def, true) + tableHead(def) +
    '<tbody data-cat="' + def.key + '">' + rows.map((l) => editRowHtml(l, def, ids)).join('') + '</tbody>';
  if (def.key === 'other') {
    h += '<tbody class="qcl-stampbody"><tr class="qcl-stamprow"' + (stamp && stamp.lid ? ' data-lid="' + esc(stamp.lid) + '"' : '') +
      (stamp && typeof stamp.unitCost === 'number' ? ' data-amt="' + esc(String(stamp.unitCost)) + '"' : '') + '>' +
      '<td colspan="' + cols + '"><span class="qcl-stamp-name">印花稅</span>' +
      '<label><input type="checkbox" class="qcl-stampchk"' + (stamp ? ' checked' : '') + '>計入（依合約金額×0.1%自動計算）</label>' +
      '<span class="qcl-mu qcl-stamphint"></span></td>' +
      '<td class="r qcl-stampamt">—</td><td></td></tr></tbody>';
  }
  return h + '<tfoot><tr><td colspan="' + cols + '">小計</td><td class="r qcl-secsub">0</td><td></td></tr></tfoot></table></div></section>';
}

function viewRowHtml(l, def, lidSet) {
  const dash = '<span class="qcl-mu">—</span>';
  return '<tr class="qcl-row">' +
    '<td class="qcl-t">' + (esc(l.desc) || dash) + (isOrphanLine(l, lidSet) ? ORPHAN_BADGE : '') + '</td>' +
    (def.hasVendor ? '<td class="qcl-t">' + (esc(l.vendor) || dash) + '</td>' : '') +
    '<td class="qcl-t">' + (esc(l.note) || dash) + '</td>' +
    '<td>' + esc(l.unit) + '</td>' +
    '<td class="r">' + esc(fmtNum(l.qty, 4)) + '</td>' +
    '<td class="r">' + esc(fmtNum(l.unitCost, 2)) + '</td>' +
    '<td class="r">' + fmtMoney(lineAmount(l)) + '</td></tr>';
}

function viewSectionHtml(def, rows, stamp, eff, lidSet) {
  const cols = def.hasVendor ? 6 : 5;
  const head = '<thead><tr><th>項目</th>' + (def.hasVendor ? '<th>' + esc(def.vendorLabel) + '</th>' : '') +
    '<th>說明</th><th>單位</th><th class="r">數量</th><th class="r">成本單價</th><th class="r">小計</th></tr></thead>';
  let body = rows.map((l) => viewRowHtml(l, def, lidSet)).join('');
  if (def.key === 'other') {
    const amt = resolveStamp(eff, stamp ? stamp.unitCost : null);
    body += '<tr class="qcl-row qcl-stamprow"><td colspan="' + cols + '"><span class="qcl-stamp-name">印花稅</span>　' +
      (stamp ? '計入（依合約金額×0.1%）' + (amt === null ? '<span class="qcl-mu">　系統依合約金額自動計算</span>' : '') : '<span class="qcl-mu">不計入</span>') +
      '</td><td class="r">' + (stamp && amt !== null ? fmtMoney(amt) : '<span class="qcl-mu">—</span>') + '</td></tr>';
  } else if (!rows.length) {
    body = '<tr class="qcl-empty"><td colspan="' + (cols + 1) + '"><span class="qcl-emptytxt">（無）</span></td></tr>';
  }
  return '<section class="qcl-sec" data-sec="' + def.key + '"><div class="qcl-sec-h"><h4>' + esc(def.name) + '</h4><span class="qcl-sec-n"></span><span class="qcl-sec-tot"></span></div>' +
    '<div class="qcl-wrap"><table class="qcl-table' + (def.hasVendor ? '' : ' nv') + '">' + colGroup(def, false) + head + '<tbody>' + body + '</tbody>' +
    '<tfoot><tr><td colspan="' + cols + '">小計</td><td class="r qcl-secsub">0</td></tr></tfoot></table></div></section>';
}

function grandHtml() {
  return '<div class="qcl-grand"><span class="qcl-grand-k">成本合計</span><span class="qcl-grand-v">0</span><span class="qcl-grand-n"></span></div>';
}

// ── 掛載：共用 ─────────────────────────────────────
/** 目前可用的營收（元）；沒給或無效回 null（canSeePrice 不再影響印花稅顯示，規格 v1.1） */
function effRevenue(inst) {
  const r = toNum(inst.revenue);
  return isFinite(r) && r >= 0 ? r : null;
}

function splitLines(lines) {
  const groups = {};
  CATS.forEach((c) => { groups[c.key] = []; });
  let stamp = null;
  lines.forEach((l) => {
    if (l.auto === 'stamp') { if (!stamp) stamp = l; } else groups[l.cat].push(l);
  });
  return { groups, stamp };
}

function stampTexts(inst, included, serverAmt) {
  const amt = resolveStamp(effRevenue(inst), serverAmt);
  return {
    amt,
    cell: included && amt !== null ? fmtMoney(amt) : '—',
    hint: amt === null ? '系統依合約金額自動計算' : '',
  };
}

/** 重算並寫入小計／合計／計數（只動文字，不碰輸入框）；回傳 totals */
function paint(inst, lines) {
  const root = inst.root;
  const t = totals(lines, effRevenue(inst));
  const counts = {};
  lines.forEach((l) => { if (l.auto !== 'stamp') counts[l.cat] = (counts[l.cat] || 0) + 1; });
  CATS.forEach((def) => {
    const sec = root.querySelector('[data-sec="' + def.key + '"]');
    if (!sec) return;
    const subText = fmtMoney(t.byCat[def.key] + (def.key === 'other' ? t.stamp : 0));
    sec.querySelector('.qcl-secsub').textContent = subText;
    sec.querySelector('.qcl-sec-tot').textContent = '小計 ' + subText;
    sec.querySelector('.qcl-sec-n').textContent = counts[def.key] ? '（' + counts[def.key] + ' 列）' : '';
  });
  const hasStamp = lines.some((l) => l.auto === 'stamp');
  root.querySelector('.qcl-grand-v').textContent = fmtMoney(t.total);
  root.querySelector('.qcl-grand-n').textContent = t.stampUnknown ? '（不含印花稅：系統依合約金額自動計算）'
    : (hasStamp ? '（含印花稅 ' + fmtMoney(t.stamp) + '）' : '');
  return t;
}

// ── 掛載：編輯模式 ──────────────────────────────────
function numField(inp, max) {
  if (!inp) return { n: 0, empty: true, bad: false };
  if (inp.validity && inp.validity.badInput) return { n: 0, empty: false, bad: true };   // type=number 的非法輸入 value 會是 ''
  const s = inp.value.trim();
  if (s === '') return { n: 0, empty: true, bad: false };
  const n = Number(s);
  if (!isFinite(n) || n < 0 || n > max) return { n: 0, empty: false, bad: true };
  return { n, empty: false, bad: false };
}

/**
 * 從 DOM 讀回目前的列。數字非法的欄位一律標紅（qcl-bad）；markBlank 為 true 時另外標示空白項目（qcl-bad）與數量 0（qcl-warn）。
 * 非法／空白的數字欄位讀成 0；呼叫端要看 invalid 決定是否擋存。
 */
function readDom(inst, markBlank) {
  const root = inst.root;
  const res = { lines: [], rows: [], invalid: 0, blankDesc: 0, zeroCost: 0, zeroQty: 0 };
  CATS.forEach((def) => {
    const tb = root.querySelector('tbody[data-cat="' + def.key + '"]');
    if (!tb) return;
    Array.prototype.forEach.call(tb.querySelectorAll('tr.qcl-row'), (tr) => {
      const q = numField(tr.querySelector('.qcl-qty'), QTY_MAX);
      const c = numField(tr.querySelector('.qcl-cost'), COST_MAX);
      const val = (sel) => { const i = tr.querySelector(sel); return i ? i.value : ''; };
      let forLids;
      const flRaw = tr.getAttribute('data-forlids');
      if (flRaw) { try { forLids = JSON.parse(flRaw); } catch (_) { forLids = undefined; } }
      const l = cleanLine({
        lid: tr.getAttribute('data-lid'), cat: def.key, desc: val('.qcl-desc'), vendor: val('.qcl-vendor'), note: val('.qcl-note'), unit: val('.qcl-unit'),
        qty: q.n, unitCost: c.n, forLid: tr.getAttribute('data-forlid'), forLids,
      });
      tr.querySelector('.qcl-qty').classList.toggle('qcl-bad', q.bad);
      tr.querySelector('.qcl-cost').classList.toggle('qcl-bad', c.bad);
      if (q.bad) res.invalid++;
      if (c.bad) res.invalid++;
      if (!l.desc) res.blankDesc++;
      if (!c.bad && l.unitCost === 0) res.zeroCost++;
      if (!q.bad && l.qty === 0) res.zeroQty++;
      if (markBlank) {
        tr.querySelector('.qcl-desc').classList.toggle('qcl-bad', !l.desc);
        tr.querySelector('.qcl-qty').classList.toggle('qcl-warn', !q.bad && l.qty === 0);
      }
      res.lines.push(l);
      res.rows.push({ tr, line: l, bad: q.bad || c.bad });
    });
  });
  // 印花稅列：勾選＝列存在。輸出的 stamp 列不帶 unitCost（伺服器忽略並自行計算）；
  // 伺服器給的金額留在 data-amt，只用於畫面與合計的估算（calcLines／serverStamp）
  const chk = root.querySelector('.qcl-stampchk');
  const srow = root.querySelector('.qcl-stamprow');
  if (chk && chk.checked) res.lines.push(cleanLine({ cat: 'other', auto: 'stamp', lid: srow.getAttribute('data-lid') }));
  const amt = srow ? srow.getAttribute('data-amt') : null;
  res.serverStamp = amt !== null && amt !== '' && isFinite(Number(amt)) ? Number(amt) : null;
  res.calcLines = res.lines.map((l) => (l.auto === 'stamp' && res.serverStamp !== null ? Object.assign({}, l, { unitCost: res.serverStamp }) : l));
  return res;
}

/** 重新計算畫面：列小計、區小計、合計、計數、列控制鈕與新增鈕的停用狀態、空區提示 */
function refresh(inst, rd) {
  const root = inst.root;
  rd = rd || readDom(inst, false);
  rd.rows.forEach((r) => { r.tr.querySelector('.qcl-sub').textContent = r.bad ? '—' : fmtMoney(lineAmount(r.line)); });
  // 「原報價品項已刪除」徽章：items 有提供才會顯示（沒提供＝null＝全部隱藏）
  const orphSet = itemLidSet(inst.items);
  rd.rows.forEach((r) => { const b = r.tr.querySelector('.qcl-orph'); if (b) b.hidden = !isOrphanLine(r.line, orphSet); });
  const t = paint(inst, rd.calcLines);

  CATS.forEach((def) => {
    const tb = root.querySelector('tbody[data-cat="' + def.key + '"]');
    const rows = tb.querySelectorAll('tr.qcl-row');
    const ph = tb.querySelector('tr.qcl-empty');
    if (!rows.length && !ph) tb.insertAdjacentHTML('beforeend', '<tr class="qcl-empty"><td colspan="' + ((def.hasVendor ? 6 : 5) + 2) + '"><span class="qcl-emptytxt">尚無項目，按「＋新增」加入</span></td></tr>');
    if (rows.length && ph) ph.remove();
    Array.prototype.forEach.call(rows, (tr, i) => {
      tr.querySelector('[data-act="up"]').disabled = i === 0;
      tr.querySelector('[data-act="down"]').disabled = i === rows.length - 1;
    });
    const mb = root.querySelector('[data-act="merge"][data-cat="' + def.key + '"]');
    if (mb) mb.disabled = rows.length < 2;   // 該區至少 2 列才能合併
  });

  const chk = root.querySelector('.qcl-stampchk');
  const st = stampTexts(inst, !!(chk && chk.checked), rd.serverStamp);
  root.querySelector('.qcl-stampamt').textContent = st.cell;
  root.querySelector('.qcl-stamphint').textContent = st.hint;

  const n = rd.lines.length;
  const atLimit = n >= inst.max;
  const hasItems = Array.isArray(inst.items);
  root.querySelector('.qcl-count').textContent = '已用 ' + n + '／' + inst.max + ' 列';
  root.querySelector('.qcl-limit').hidden = !atLimit;
  Array.prototype.forEach.call(root.querySelectorAll('[data-act="add"]'), (b) => { b.disabled = atLimit; });
  root.querySelector('[data-act="supplement"]').disabled = atLimit || !hasItems;
  root.querySelector('[data-act="reseed"]').disabled = !hasItems;
  if (chk) chk.disabled = atLimit && !chk.checked;
  return { lines: rd.lines, totals: t };
}

function setMsg(inst, m) {
  const el = inst.root && inst.root.querySelector('.qcl-msg');
  if (el) el.textContent = m || '';
}

function emitChange(inst) {
  const out = refresh(inst);
  if (typeof inst.opts.onChange === 'function') {
    try { inst.opts.onChange(out.lines, out.totals); } catch (err) { if (global.console) console.error('QCL onChange', err); }
  }
}

function renderEdit(inst, lines) {
  ensureStyle();
  const ids = { desc: 'qclDlDesc' + inst.id, unit: 'qclDlUnit' + inst.id };
  const sp = splitLines(lines);
  inst.el.innerHTML = '<div class="qcl-root" data-mode="edit">' +
    '<div class="qcl-tools">' +
      '<button type="button" class="qcl-btn" data-act="reseed" title="丟掉目前的成本明細，依客戶報價品項重新產生預設列">↻ 由報價品項重新帶入</button>' +
      '<button type="button" class="qcl-btn" data-act="supplement" title="只補入「還沒有對應成本列」的新品項，不動既有列">＋補入新品項</button>' +
      '<span class="qcl-sp"></span><span class="qcl-count"></span></div>' +
    '<p class="qcl-hint">金額單位：新台幣元（未稅）。小計＝數量×成本單價；成本明細與客戶報價品項各自獨立、不必一一對應：多個品項的成本可以併成一列（整包），一個品項也可以拆成多列。</p>' +
    '<div class="qcl-limit" hidden>已達成本明細列數上限（' + inst.max + ' 列），無法再新增；請刪除不需要的列。</div>' +
    '<div class="qcl-msg" role="status" aria-live="polite"></div>' +
    CATS.map((def) => editSectionHtml(def, sp.groups[def.key], def.key === 'other' ? sp.stamp : null, ids)).join('') +
    grandHtml() +
    '<datalist id="' + ids.desc + '">' + DESC_SUGGEST.map((s) => '<option value="' + esc(s) + '"></option>').join('') + '</datalist>' +
    '<datalist id="' + ids.unit + '">' + UNIT_SUGGEST.map((s) => '<option value="' + esc(s) + '"></option>').join('') + '</datalist>' +
    '</div>';
  inst.root = inst.el.firstElementChild;
  inst.ids = ids;
  refresh(inst);
}

function appendRows(inst, lines) {
  lines.forEach((l) => {
    const def = CAT_BY_KEY[l.cat];
    inst.root.querySelector('tbody[data-cat="' + def.key + '"]').insertAdjacentHTML('beforeend', editRowHtml(l, def, inst.ids));
  });
}

function addRow(inst, cat) {
  const def = CAT_BY_KEY[cat];
  if (!def || readDom(inst, false).lines.length >= inst.max) return;
  appendRows(inst, [emptyLine(cat)]);
  setMsg(inst, '');
  emitChange(inst);
  const rows = inst.root.querySelectorAll('tbody[data-cat="' + cat + '"] tr.qcl-row');
  rows[rows.length - 1].querySelector('.qcl-desc').focus();
}

function moveRow(inst, btn, d) {
  const tr = btn.closest('tr.qcl-row');
  const sib = d < 0 ? tr.previousElementSibling : tr.nextElementSibling;
  if (!sib || !sib.classList.contains('qcl-row')) return;
  tr.parentNode.insertBefore(d < 0 ? tr : sib, d < 0 ? sib : tr);
  emitChange(inst);
  // 重新插入 DOM 會讓焦點消失：把焦點還給同方向的按鈕（已到頭就改給另一顆）
  const again = tr.querySelector('[data-act="' + (d < 0 ? 'up' : 'down') + '"]');
  const other = tr.querySelector('[data-act="' + (d < 0 ? 'down' : 'up') + '"]');
  const fb = !again.disabled ? again : (!other.disabled ? other : null);
  if (fb) fb.focus();
}

function delRow(inst, btn) {
  const tr = btn.closest('tr.qcl-row');
  const cat = tr.parentNode.getAttribute('data-cat');
  tr.remove();
  setMsg(inst, '');
  emitChange(inst);
  const add = inst.root.querySelector('[data-act="add"][data-cat="' + cat + '"]');
  if (add && !add.disabled) add.focus();
}

/** 確認視窗：預設 window.confirm，可由 opts.confirm 換成自訂（回傳 boolean 或 Promise<boolean>） */
function ask(inst, msg) {
  const f = typeof inst.opts.confirm === 'function' ? inst.opts.confirm : (global.confirm ? global.confirm.bind(global) : () => true);
  return Promise.resolve(f(msg));
}

/**
 * 「合併為一列」：把某一區的所有列併成 1 列（整包）。成本合計不變，但各列的項目／廠商／說明／數量單價會消失，所以先確認。
 * 輸入欄位有非法數字時不合併（合計算不準）。
 */
function mergeRows(inst, cat) {
  const def = CAT_BY_KEY[cat];
  if (!def) return Promise.resolve();
  const pick = (rd) => rd.lines.filter((l) => l.auto !== 'stamp' && l.cat === cat);
  const rd = readDom(inst, false);
  if (rd.invalid > 0) { setMsg(inst, '有欄位不是有效的數字（已標紅），請先修正再合併'); return Promise.resolve(); }
  const rows = pick(rd);
  if (rows.length < 2) return Promise.resolve();
  const total = rows.reduce((s, l) => s + lineCents(l), 0) / 100;
  return ask(inst, '要把「' + def.name + '」的 ' + rows.length + ' 列合併成 1 列嗎？\n成本合計不變（' + fmtMoney(total) + ' 元），但各列的項目、廠商、說明與數量／單價會併成一列（單位「式」、數量 1），合併後可再改名。').then((ok) => {
    if (!ok || inst.dead) return;
    const rd2 = readDom(inst, false);   // 確認期間畫面可能已變動：以按下確定當下的列為準
    const rows2 = pick(rd2);
    if (rd2.invalid > 0 || rows2.length < 2) return;
    const merged = mergeLines(rows2);
    const tb = inst.root.querySelector('tbody[data-cat="' + cat + '"]');
    Array.prototype.forEach.call(tb.querySelectorAll('tr.qcl-row'), (tr) => tr.remove());
    appendRows(inst, [merged]);
    setMsg(inst, '已把 ' + rows2.length + ' 列合併成 1 列（成本合計不變）；可直接改項目名稱');
    emitChange(inst);
    const d = tb.querySelector('tr.qcl-row .qcl-desc');
    if (d) { d.focus(); if (d.select) d.select(); }
  });
}

function reseed(inst) {
  if (!Array.isArray(inst.items)) return Promise.resolve();
  return ask(inst, '要依報價品項重新帶入嗎？\n目前編輯中的成本明細（含已填的成本、廠商與新增的列）會被取代。').then((ok) => {
    if (!ok || inst.dead) return;
    renderEdit(inst, seedFromItems(inst.items, inst.opts));
    setMsg(inst, '已依報價品項重新帶入');
    emitChange(inst);
  });
}

/**
 * 「＋補入新品項」要補哪些：回傳 seedItemLines(items) 之中「還沒有對應成本列」的那幾列（純函式，可單獨測）。
 * 覆蓋判定（成本列 ↔ 客戶品項）：
 *   1. 成本列的 forLid 等於某品項的 lid → 該品項已覆蓋（以 lid 為準，改名也不受影響）。
 *   2. 其餘成本列（forLid 為空，或 forLid 不屬於目前任何品項）→ 改用「品名相同（去頭尾空白）」的多重集合比對：
 *      同名品項有 N 個就要有 N 列同名成本列才算全部覆蓋，每一列成本列最多覆蓋一個品項。
 *   為什麼不能只看 forLid：新單的品項在第一次儲存前沒有 lid（種子列沒有 forLid），伺服器存檔後替品項指派 lid 卻不回填成本列的 forLid，
 *   之後重開按「補入」，只看 forLid 會把所有品項都當成新品項再補一次。印花稅列不參與比對。
 */
function missingItemLines(curLines, items, opts) {
  const itemLines = seedItemLines(items, opts);
  const entries = itemLines.map((l) => ({ lid: lidKey(l.forLid), name: nameKey(l.desc), line: l }));
  return uncovered(curLines, entries, entries).map((e) => e.line);
}

// ── 涵蓋判定（共用核心）──────────────────────────────────────
// ⚠ 下面這段（控制字元清洗、lidKey／nameKey／lineLidList／generalEntries／uncovered／unmatchedMessage）與伺服器 lib/quoteCostLines.js 的同名函式逐字鏡像：
//   伺服器用它產生 preview.warnings／derived.costWarnings，這裡用它在畫面上即時提醒；scripts/check-quote-costlines-ui.js 以 vm 載入本檔、
//   隨機輸入逐組比對兩邊的結果（unmatchedPricedItems ＝ CL.unmatchedItems、unmatchedMessage ＝ CL.unmatchedMessage）。改這裡必須同步改伺服器。
const CTRL_RE = new RegExp('[' + String.fromCharCode(0) + '-' + String.fromCharCode(31) + String.fromCharCode(127) + String.fromCharCode(0x2028) + String.fromCharCode(0x2029) + ']+', 'g');
const WARN_NAME_MAX = 30, WARN_LIST_MAX = 5, UNNAMED_ITEM = '（未命名品項）';

/** 品名比對鍵：控制字元換空白、去頭尾空白、截 120 字；null／undefined → '' */
function nameKey(v) { return v === undefined || v === null ? '' : String(v).replace(CTRL_RE, ' ').trim().slice(0, DESC_MAX); }
/** lid 比對鍵：只認字串，去頭尾空白、截 64 字；其餘 → '' */
function lidKey(v) { return typeof v === 'string' ? v.trim().slice(0, LID_MAX) : ''; }

/** 成本列指向的品項代碼清單：forLid＋forLids（去空白；forLids 最多取前 60 個） */
function lineLidList(l) {
  const out = [];
  if (!l || typeof l !== 'object') return out;
  const a = lidKey(l.forLid);
  if (a) out.push(a);
  if (Array.isArray(l.forLids)) {
    const n = Math.min(l.forLids.length, MAX_FOR_LIDS);
    for (let i = 0; i < n; i++) { const k = lidKey(l.forLids[i]); if (k) out.push(k); }
  }
  return out;
}

/** 一般品項（略過 title／subtotal 與非物件）→ [{lid, name, priced, index}]；沒有 lid 的品項：有暫時代號 nid 用 nid，否則 'legacy-'+索引（同伺服器 serialize） */
function generalEntries(items) {
  const out = [];
  (Array.isArray(items) ? items : []).forEach((it, i) => {
    if (!it || typeof it !== 'object' || it.kind === 'title' || it.kind === 'subtotal') return;
    const p = parseFloat(it.unitPrice);
    // 畫面上還沒存檔的新品項沒有 lid，改用暫時代號 nid（伺服器的 items 不會有 nid，所以伺服器版沒有這一層；其餘逐字相同）
    out.push({ lid: lidKey(it.lid) || lidKey(it.nid) || ('legacy-' + i), name: nameKey(it.desc), priced: isFinite(p) && p > 0, index: i });
  });
  return out;
}

/**
 * cands（entries 的子集，依品項順序）裡沒有被成本列涵蓋的那幾個。
 * 成本列（略過印花稅）的 forLid／forLids 命中 entries 任一 lid → 指向品項（命中者算涵蓋）；沒命中任何品項的列 → 以品名多重集合涵蓋同名候選
 * （同名品項有 N 個要有 N 列才算全部涵蓋；空白品名不進池也不被涵蓋）。
 */
function uncovered(lines, entries, cands) {
  const lidSet = new Set(entries.map((e) => e.lid));
  const covered = new Set();
  const free = new Map();
  (Array.isArray(lines) ? lines : []).forEach((l) => {
    if (!l || typeof l !== 'object' || l.auto === 'stamp') return;
    const hit = lineLidList(l).filter((k) => lidSet.has(k));
    if (hit.length) { hit.forEach((k) => covered.add(k)); return; }
    const d = nameKey(l.desc);
    if (d) free.set(d, (free.get(d) || 0) + 1);
  });
  return cands.filter((e) => {
    if (covered.has(e.lid)) return false;
    if (e.name) {
      const n = free.get(e.name) || 0;
      if (n > 0) { free.set(e.name, n - 1); return false; }
    }
    return true;
  });
}

/** 有價（unitPrice>0）的一般品項中，沒有被任何成本列涵蓋者：[{lid, name, index}]（＝伺服器 CL.unmatchedItems；不做任何清洗，直接吃原始輸入） */
function unmatchedPricedItems(lines, items) {
  const entries = generalEntries(items);
  return uncovered(lines, entries, entries.filter((e) => e.priced)).map((e) => ({ lid: e.lid, name: e.name, index: e.index }));
}

/** 警告文字（＝伺服器 CL.unmatchedMessage）：總數＋最多 5 個品名（每個截 30 字） */
function unmatchedMessage(names) {
  const arr = Array.isArray(names) ? names : [];
  const shown = arr.slice(0, WARN_LIST_MAX).map((n) => {
    const d = nameKey(n) || UNNAMED_ITEM;
    return d.length > WARN_NAME_MAX ? d.slice(0, WARN_NAME_MAX) + '…' : d;
  });
  const rest = arr.length - shown.length;
  return '有 ' + arr.length + ' 個有價品項在成本明細中找不到對應的成本列：' + shown.join('、') + (rest > 0 ? '…另 ' + rest + ' 個' : '') +
    '。若成本已包含在其他列（例如整包）可忽略此提醒；若是漏填，毛利率可能被高估。';
}

/** 業務端常駐提醒／儲存前確認用的文字：有價品項沒有成本列涵蓋 → 與伺服器 preview.warnings 同一句；全部都有對應回 ''（純提醒、不擋） */
function unmatchedPricedNote(lines, items) {
  const un = unmatchedPricedItems(lines, items);
  return un.length ? unmatchedMessage(un.map((x) => x.name)) : '';
}

/**
 * 儲存前確認的「基準」：目前沒有成本列涵蓋的有價品項的鍵陣列（品項 lid；還沒存檔的新品項沒有 lid，鍵是畫面給的暫時代號 nid）。
 * 表單載入時記一次（編輯既有單：載入的品項＋已儲存的成本明細；新單：不記＝空基準），儲存時改用 unmatchedNewNote 只看「基準以外」新出現的未涵蓋品項。
 * 為什麼：整包成本（刪列做整包、品名對不上又沒有 forLid）是合法情境，這類單每次儲存都問一次是噪音；真正要提醒的是「這次編輯讓更多品項沒有成本列」。
 */
function unmatchedKeys(lines, items) {
  return unmatchedPricedItems(lines, items).map((x) => x.lid);
}

/**
 * 儲存前確認用的文字：只算 baseline（unmatchedKeys 記下的鍵陣列）以外、新出現的有價未涵蓋品項，文字格式同 unmatchedPricedNote（只列這幾個新品項）；
 * 沒有新出現的回 ''。baseline 不是陣列（沒有基準，例如新單）＝空基準 → 等同 unmatchedPricedNote。
 * 常駐提醒（毛利頁籤）不走這支，仍用 unmatchedPricedNote 永遠顯示全部未涵蓋品項。
 */
function unmatchedNewNote(lines, items, baseline) {
  const base = new Set(Array.isArray(baseline) ? baseline : []);
  const fresh = unmatchedPricedItems(lines, items).filter((x) => !base.has(x.lid));
  return fresh.length ? unmatchedMessage(fresh.map((x) => x.name)) : '';
}

/** 補入確認視窗最多列出幾個品名；超過的以「…另 N 個」帶過 */
const SUPPLEMENT_LIST_MAX = 8;
/** 確認視窗裡單一品名最長字數（品名最長 120 字，列 8 個會變成一整面牆） */
const SUPPLEMENT_NAME_MAX = 30;

/**
 * 「＋補入新品項」按下去之前的確認文字（純函式，可單獨測）。lines＝將補入的成本列；skipped＝因列數上限而不會補入的項數。
 * 為什麼要確認：成本明細與客戶品項不必一一對應（多個品項成本併成一列「整包」很常見），被整包涵蓋的舊品項在「覆蓋判定」裡仍算沒有對應列，
 * 直接補入會把它們又加回來；先讓使用者看見會補哪些品項，成本已含在其他列時可以取消。
 */
function supplementMessage(lines, skipped) {
  const arr = Array.isArray(lines) ? lines : [];
  const sk = Number(skipped) > 0 ? Math.floor(Number(skipped)) : 0;
  return '將補入以下 ' + arr.length + ' 個尚未有成本列的品項：' + listNames(arr) +
    (sk ? '\n（已達列數上限，另有 ' + sk + ' 項不會補入）' : '') +
    '\n\n若這些品項的成本已包含在其他列中（例如整包），請按取消。';
}

/** 品名清單文字：最多列 SUPPLEMENT_LIST_MAX 個（每個最長 SUPPLEMENT_NAME_MAX 字），超過接「…另 N 個」 */
function listNames(lines) {
  const arr = Array.isArray(lines) ? lines : [];
  const names = arr.slice(0, SUPPLEMENT_LIST_MAX).map((l) => {
    const d = text(l && l.desc, DESC_MAX);
    return d.length > SUPPLEMENT_NAME_MAX ? d.slice(0, SUPPLEMENT_NAME_MAX) + '…' : d;
  });
  const rest = arr.length - names.length;
  return names.join('、') + (rest > 0 ? '…另 ' + rest + ' 個' : '');
}

/**
 * 「完成」前的提醒（純提醒、不擋）：成本明細裡找不到對應列的報價品項。成本明細與品項不必一一對應（整包、一式拆多列），
 * 所以這裡只列出名單、由使用者自己判斷是「已包含在其他列」還是「漏填」；全部都有對應就回空字串。
 * 為什麼需要：舊版（每個品項一列成本）漏填會被「有價品項成本需 > 0」擋下；新式成本只檢查總成本 > 0，業務事後新增品項而顧問沒補成本，毛利會悄悄偏高。
 */
function unmatchedItemsNote(lines, items, opts) {
  const miss = missingItemLines(lines, items, opts);
  if (!miss.length) return '';
  return '以下 ' + miss.length + ' 個報價品項在成本明細裡找不到對應的列：' + listNames(miss) + '。\n若它們的成本已包含在其他列（例如整包），可以直接完成；若是漏填，請先回去補上。';
}

/**
 * 把同一分區的多列合併成 1 列（純函式，可單獨測）：單位「式」、數量 1、成本單價＝各列小計（取整到分）的合計，所以成本合計不變。
 * 項目＝「合併 N 項：A、B、C」（可再改名）；廠商只有 1 種就帶入，多種就改記在說明；沿用第一列的 lid 與 forLid，
 * 並把所有被併列的 forLid∪forLids 放進 forLids（涵蓋判定因此把被併的每個品項都當已有成本列）。
 * 傳入的列要同一分類、不含印花稅列；空陣列回 null。
 */
function mergeLines(rows) {
  const ls = (Array.isArray(rows) ? rows : []).map(cleanLine).filter((l) => l && l.auto !== 'stamp');
  if (!ls.length) return null;
  const first = ls[0];
  const names = ls.map((l) => l.desc).filter(Boolean);
  const vendors = Array.from(new Set(ls.map((l) => l.vendor).filter(Boolean)));
  const cents = ls.reduce((s, l) => s + lineCents(l), 0);
  const o = {
    cat: first.cat,
    desc: ls.length > 1 ? '合併 ' + ls.length + ' 項：' + names.join('、') : (names[0] || ''),
    vendor: vendors.length === 1 ? vendors[0] : '',
    note: vendors.length > 1 ? '廠商：' + vendors.join('、') : '',
    unit: DEFAULT_UNIT, qty: 1, unitCost: cents / 100,
  };
  if (first.lid) o.lid = first.lid;
  if (first.forLid) o.forLid = first.forLid;
  // 被併各列所涵蓋的品項（forLid∪forLids）全部帶著：整包後「補入新品項」與完成／儲存前的提醒仍認得這些品項已有成本列（cleanLine 會剔除與 forLid 相同者、去重、上限 60）
  const all = [];
  ls.forEach((l) => lineLidList(l).forEach((k) => { if (all.indexOf(k) < 0) all.push(k); }));
  if (all.length) o.forLids = all;
  return cleanLine(o);
}

/** 目前該補哪些：{ cand: 還沒有成本列的品項列（全部）, add: 列數上限內實際會補入的, skipped: 超過上限不補的項數 } */
function planSupplement(inst) {
  const cur = readDom(inst, false).lines;
  const cand = missingItemLines(cur, inst.items, inst.opts);
  const room = Math.max(0, inst.max - cur.length);
  const add = cand.slice(0, room);
  return { cand, add, skipped: cand.length - add.length };
}

function supplement(inst) {
  if (!Array.isArray(inst.items)) return Promise.resolve();
  const plan = planSupplement(inst);
  if (!plan.cand.length) { setMsg(inst, '沒有需要補入的新品項（每個報價品項都已有對應的成本列）'); return Promise.resolve(); }
  if (!plan.add.length) { setMsg(inst, '已達列數上限，無法補入；另有 ' + plan.skipped + ' 項尚未有對應的成本列'); return Promise.resolve(); }
  // 先讓使用者看見會補哪些品項（成本可能已含在其他列，例如整包）；取消就完全不動
  return ask(inst, supplementMessage(plan.add, plan.skipped)).then((ok) => {
    if (inst.dead) return;
    if (!ok) { setMsg(inst, '已取消，沒有補入任何列'); return; }
    const p2 = planSupplement(inst);   // 確認期間畫面可能已變動：以按下確定當下的狀態為準
    if (!p2.add.length) { setMsg(inst, '沒有需要補入的新品項（每個報價品項都已有對應的成本列）'); return; }
    appendRows(inst, p2.add);
    setMsg(inst, '已補入 ' + p2.add.length + ' 項' + (p2.skipped ? '；已達列數上限，另有 ' + p2.skipped + ' 項未補入' : ''));
    emitChange(inst);
  });
}

function onClick(inst, ev) {
  const b = ev.target.closest ? ev.target.closest('button[data-act]') : null;
  if (!b || !inst.el.contains(b) || b.disabled) return;
  const act = b.getAttribute('data-act');
  if (act === 'add') addRow(inst, b.getAttribute('data-cat'));
  else if (act === 'up') moveRow(inst, b, -1);
  else if (act === 'down') moveRow(inst, b, 1);
  else if (act === 'del') delRow(inst, b);
  else if (act === 'reseed') reseed(inst);
  else if (act === 'supplement') supplement(inst);
  else if (act === 'merge') mergeRows(inst, b.getAttribute('data-cat'));
}

function onInput(inst, ev) {
  const t = ev.target;
  if (!t || !t.classList || !t.classList.contains('qcl-in')) return;
  if (t.classList.contains('qcl-desc') && t.value.trim()) t.classList.remove('qcl-bad');
  if (t.classList.contains('qcl-qty')) t.classList.remove('qcl-warn');
  setMsg(inst, '');
  emitChange(inst);   // 只更新文字，不重畫表格（輸入框維持焦點）
}

function onChangeEv(inst, ev) {
  if (ev.target && ev.target.classList && ev.target.classList.contains('qcl-stampchk')) emitChange(inst);
}

/** Enter 不送出表單：改為移到下一個輸入框（輸入法組字中的 Enter 不攔截） */
function onKeydown(inst, ev) {
  if (ev.key !== 'Enter' || ev.isComposing || ev.keyCode === 229) return;
  const t = ev.target;
  if (!t || !t.matches || !t.matches('.qcl-in, .qcl-stampchk')) return;
  ev.preventDefault();
  const ins = Array.prototype.filter.call(inst.root.querySelectorAll('.qcl-in, .qcl-stampchk'), (x) => !x.disabled);
  const i = ins.indexOf(t);
  if (i >= 0 && i < ins.length - 1) ins[i + 1].focus();
}

// ── 掛載：唯讀模式 ──────────────────────────────────
function renderView(inst, lines) {
  ensureStyle();
  const sp = splitLines(lines);
  const lidSet = itemLidSet(inst.items);   // 沒提供 items → null → 不顯示「原報價品項已刪除」徽章
  inst.el.innerHTML = '<div class="qcl-root" data-mode="view">' +
    '<p class="qcl-hint">金額單位：新台幣元（未稅）。</p>' +
    CATS.map((def) => viewSectionHtml(def, sp.groups[def.key], def.key === 'other' ? sp.stamp : null, effRevenue(inst), lidSet)).join('') +
    grandHtml() + '</div>';
  inst.root = inst.el.firstElementChild;
  paint(inst, lines);
}

// ── 拖曳排序（滑鼠／觸控／鍵盤共用；成本明細編輯器與報價項目表都用它）──────────
// 純函式（不碰 DOM，可單元測試）。rects＝各列由上到下的矩形 {top,bottom}（或 {top,height}），y＝指標的 clientY。
/** 指標落在第幾個「縫隙」：0＝第一列之前 … n＝最後一列之後。以各列的垂直中線為界（越過該列中線才算越過該列），列高不等也成立；超出頭尾就夾在 0／n */
function dragSlot(rects, y) {
  const n = Array.isArray(rects) ? rects.length : 0;
  const yy = Number(y);
  if (!n || !isFinite(yy)) return 0;
  let slot = 0;
  for (let i = 0; i < n; i++) {
    const r = rects[i] || {};
    const top = Number(r.top);
    const bottom = r.bottom !== undefined ? Number(r.bottom) : top + Number(r.height);
    const mid = top + (bottom - top) / 2;
    if (isFinite(mid) && yy > mid) slot = i + 1; else break;
  }
  return slot;
}
/** 縫隙 → 移動後的最終索引（先抽出原列再插入，等同 arr.splice(to, 0, arr.splice(from, 1)[0])）；縫隙在原列的前後都等於沒動 */
function dragSlotToIndex(slot, from) { return slot > from ? slot - 1 : slot; }
/** dragTargetIndex(rects, y)：只給兩個參數回傳「縫隙」(0..n)；再給 fromIdx 就回傳移動後的最終索引（0..n-1），放回原位時等於 fromIdx */
function dragTargetIndex(rects, y, fromIdx) {
  const slot = dragSlot(rects, y);
  const n = Array.isArray(rects) ? rects.length : 0;
  return Number.isInteger(fromIdx) && fromIdx >= 0 && fromIdx < n ? dragSlotToIndex(slot, fromIdx) : slot;
}
/** 插入線的 y：縫隙 0＝第一列上緣、n＝最後一列下緣、其餘＝相鄰兩列之間 */
function dragLineY(rects, slot) {
  const n = Array.isArray(rects) ? rects.length : 0;
  if (!n) return 0;
  const bot = (r) => (r.bottom !== undefined ? Number(r.bottom) : Number(r.top) + Number(r.height));
  if (slot <= 0) return Number(rects[0].top);
  if (slot >= n) return bot(rects[n - 1]);
  return (bot(rects[slot - 1]) + Number(rects[slot].top)) / 2;
}

const DND_THRESHOLD = 4;      // 開始拖曳的最小位移（px）；沒超過就只是點擊
const DND_MAX_SPEED = 22;     // 自動捲動的最大速度（px／16ms）
let dndLive = null;           // 螢幕閱讀器播報用的 aria-live 區域（全頁共用一個）

function dndAnnounce(msg) {
  const doc = global.document;
  if (!doc || !doc.body || !doc.createElement) return;
  if (!dndLive || !dndLive.parentNode) {
    dndLive = doc.createElement('div');
    dndLive.id = 'qclDndLive';
    dndLive.className = 'qcl-sr';
    dndLive.setAttribute('role', 'status');
    dndLive.setAttribute('aria-live', 'polite');
    doc.body.appendChild(dndLive);
  }
  dndLive.textContent = msg;
}

/** 元素「看得見」的矩形：自己的矩形 ∩ 所有會裁切的祖先（overflow 不是 visible）∩ 視窗 */
function dndVisibleRect(el) {
  const vw = global.innerWidth > 0 ? global.innerWidth : 1e9;
  const vh = global.innerHeight > 0 ? global.innerHeight : 1e9;
  const r0 = el.getBoundingClientRect();
  let l = Math.max(0, r0.left), r = Math.min(vw, r0.right), t = Math.max(0, r0.top), b = Math.min(vh, r0.bottom);
  const doc = global.document;
  for (let n = el.parentElement; n && doc && n !== doc.body && n !== doc.documentElement; n = n.parentElement) {
    const cs = global.getComputedStyle ? global.getComputedStyle(n) : null;
    if (!cs || (cs.overflowX === 'visible' && cs.overflowY === 'visible')) continue;
    const pr = n.getBoundingClientRect();
    l = Math.max(l, pr.left); r = Math.min(r, pr.right); t = Math.max(t, pr.top); b = Math.min(b, pr.bottom);
  }
  return { left: l, right: r, top: t, bottom: b };
}

/** 自動捲動的候選容器（由內而外）：opt 給元素／函式就用它，否則往外找最近的、真的可以捲動的祖先；最後一定加上整頁 */
function dndScrollers(row, opt) {
  if (opt === false) return [];
  const doc = global.document;
  const out = [];
  if (opt) {
    const el = typeof opt === 'function' ? opt(row) : opt;
    if (el && el.nodeType) out.push(el);
  } else if (doc) {
    for (let n = row.parentElement; n && n !== doc.body && n !== doc.documentElement; n = n.parentElement) {
      const cs = global.getComputedStyle ? global.getComputedStyle(n) : null;
      const oy = cs ? cs.overflowY : '';
      if ((oy === 'auto' || oy === 'scroll' || oy === 'overlay') && n.scrollHeight > n.clientHeight + 1) out.push(n);
    }
  }
  const root = doc && (doc.scrollingElement || doc.documentElement);
  if (root && out.indexOf(root) < 0) out.push(root);
  return out;
}

/**
 * QCL.dragSort(container, opts)：讓 container 內的列可以用把手拖曳排序（Pointer Events：滑鼠、觸控、手寫筆；另有鍵盤移動模式）。
 *   opts.rowSelector      列的選擇器（必填，例如 'tr.qcl-row'）
 *   opts.handleSelector   把手的選擇器（必填，例如 '.qcl-drag'）。只有在把手上按住才會拖曳；輸入框內選字、點擊、Tab 完全不受影響
 *   opts.onMove(from, to, info)  放開後呼叫（必填）。from＝原索引、to＝移動後的最終索引（同 splice(to, 0, splice(from, 1)[0])）；放回原位不呼叫。info＝{ row, rows, handle }
 *   opts.rowsOf(row)      這一列「可互相排序」的列清單（預設：容器內所有符合 rowSelector 的列）。成本明細用它限制在同一分區：
 *                         指標跑到別區時，插入線與落點都夾在本區的頭／尾邊界（不能跨區）
 *   opts.canDrop(from, to, info)  回傳 false＝這個位置不能放（不畫插入線、放開視為取消）
 *   opts.scrollContainer  自動捲動的容器：元素、(row)=>元素，或 false（不自動捲動）；預設往外找最近的可捲動祖先，再加整頁
 *   opts.labelOf(row, idx)  浮影與螢幕閱讀器播報用的列名稱（預設取列內第一個有字的文字輸入框）
 *   opts.keyboard         false＝不提供鍵盤移動模式；opts.threshold＝開始拖曳的最小位移（預設 4px）
 * 事件全部委派在 container 上，所以列重畫（innerHTML）不必重新綁定。回傳 { destroy(), cancel(), isDragging() }。
 * 拖曳中：被拖的列半透明（qcl-dnd-src）、浮影跟著指標（qcl-dnd-ghost）、目標縫隙畫一條插入線（qcl-dnd-line）；
 * 指標靠近可捲動容器的上下緣會自動捲動；Esc／pointercancel／視窗失焦都會取消並還原。
 */
function dragSort(container, opts) {
  if (!container || typeof container.addEventListener !== 'function') throw new TypeError('QCL.dragSort: 需要容器元素');
  opts = opts || {};
  if (typeof opts.rowSelector !== 'string' || !opts.rowSelector) throw new TypeError('QCL.dragSort: 需要 rowSelector');
  if (typeof opts.handleSelector !== 'string' || !opts.handleSelector) throw new TypeError('QCL.dragSort: 需要 handleSelector');
  if (typeof opts.onMove !== 'function') throw new TypeError('QCL.dragSort: 需要 onMove(fromIdx, toIdx)');
  ensureStyle();
  const threshold = opts.threshold >= 0 ? Number(opts.threshold) : DND_THRESHOLD;
  const doc = () => global.document;

  let st = null;        // 目前的手勢：{ mode:'pointer'|'kbd', phase:'pending'|'drag'|'kbd', handle, row, rows, from, to, slot, ok, ... }
  let dead = false;
  const base = [];      // 容器層級的監聽（destroy 時移除）
  const gest = [];      // 單次手勢的監聽（結束時移除）
  const on = (list, target, type, fn, cap) => {
    if (!target || typeof target.addEventListener !== 'function') return;
    target.addEventListener(type, fn, cap);
    list.push([target, type, fn, cap]);
  };
  const off = (list) => {
    list.splice(0).forEach((x) => { try { x[0].removeEventListener(x[1], x[2], x[3]); } catch (_) { /* 已移除 */ } });
  };
  const flat = () => Array.prototype.slice.call(container.querySelectorAll(opts.rowSelector));
  const rowsFor = (row) => Array.prototype.slice.call((typeof opts.rowsOf === 'function' ? opts.rowsOf(row) : flat()) || []);
  const info = (s) => ({ row: s.row, rows: s.rows, handle: s.handle });
  const stopEv = (ev) => { ev.preventDefault(); ev.stopPropagation(); if (ev.stopImmediatePropagation) ev.stopImmediatePropagation(); };

  function labelFor(row, idx) {
    let s = '';
    if (typeof opts.labelOf === 'function') { try { s = String(opts.labelOf(row, idx) || ''); } catch (_) { s = ''; } }
    if (!s && row.querySelectorAll) {
      const ins = row.querySelectorAll('input, textarea');
      for (let i = 0; i < ins.length && !s; i++) {
        const ty = ins[i].type;
        if ((ty === 'text' || ty === 'search' || ty === 'textarea') && ins[i].value && String(ins[i].value).trim()) s = String(ins[i].value);
      }
    }
    s = s.split(/\s+/).join(' ').trim();
    if (s.length > 40) s = s.slice(0, 39) + '…';
    return s || ('第 ' + (idx + 1) + ' 列');
  }

  function makeGhost(s) {
    const d = doc();
    const g = d.createElement('div');
    g.className = 'qcl-dnd-ghost';
    g.setAttribute('aria-hidden', 'true');
    const grip = d.createElement('span');
    grip.className = 'qcl-dnd-grip';
    grip.textContent = '⋮⋮';
    const txt = d.createElement('span');
    txt.className = 'qcl-dnd-txt';
    txt.textContent = labelFor(s.row, s.from);
    g.appendChild(grip);
    g.appendChild(txt);
    s.ghostW = Math.max(120, Math.min((s.row.getBoundingClientRect().width) || 300, 340));
    g.style.width = s.ghostW + 'px';
    d.body.appendChild(g);
    s.ghostH = g.offsetHeight || 32;
    s.ghost = g;
  }
  function makeLine(s) {
    const d = doc();
    const l = d.createElement('div');
    l.className = 'qcl-dnd-line';
    l.setAttribute('aria-hidden', 'true');
    d.body.appendChild(l);
    s.line = l;
  }

  /** 插入線：畫在目標縫隙，水平範圍＝這批列「看得見」的範圍；縫隙在可視範圍外、或不能放時隱藏 */
  function paintLine(s) {
    if (!s.line) return;
    const vr = dndVisibleRect(s.rows[0].parentNode && s.rows[0].parentNode.getBoundingClientRect ? s.rows[0].parentNode : container);
    const y = dragLineY(s.rects, s.slot);
    const show = s.ok && vr.right > vr.left && y >= vr.top - 1 && y <= vr.bottom + 1;
    s.line.style.display = show ? 'block' : 'none';
    if (!show) return;
    s.line.style.width = Math.max(0, vr.right - vr.left) + 'px';
    s.line.style.transform = 'translate(' + Math.round(vr.left) + 'px,' + Math.round(y - 1.5) + 'px)';
    s.line.style.opacity = s.to === s.from ? '.4' : '1';    // 放回原位（沒動）時淡一點
  }
  function paint(s) {
    if (s.ghost) {
      const vw = global.innerWidth > 0 ? global.innerWidth : 1e9;
      const left = Math.max(4, Math.min(s.x - 18, vw - s.ghostW - 4));
      const top = s.pointerType === 'touch' ? s.y - s.ghostH - 18 : s.y - s.ghostH / 2;   // 觸控時浮影放在手指上方，不被手指遮住
      s.ghost.style.transform = 'translate(' + Math.round(left) + 'px,' + Math.round(top) + 'px)';
      s.ghost.classList.toggle('qcl-dnd-bad', !s.ok);
    }
    paintLine(s);
  }
  function judge(s) {
    s.ok = true;
    if (typeof opts.canDrop === 'function') {
      try { s.ok = opts.canDrop(s.from, s.to, info(s)) !== false; } catch (err) { s.ok = false; if (global.console) console.error('QCL.dragSort canDrop', err); }
    }
  }
  /** 依指標位置重算目標（每次移動、每次自動捲動後都要重算：捲動會讓各列的矩形跟著變） */
  function update(s) {
    s.rects = s.rows.map((r) => r.getBoundingClientRect());
    s.slot = dragSlot(s.rects, s.y);
    s.to = dragSlotToIndex(s.slot, s.from);
    judge(s);
    paint(s);
  }
  function repaint(s) {
    if (s.mode === 'pointer') { if (s.phase === 'drag') update(s); return; }
    s.rects = s.rows.map((r) => r.getBoundingClientRect());
    paintLine(s);
  }

  /** 指標靠近可捲動容器的上下緣就捲動（離邊緣越近越快）；有實際捲動回傳 true */
  function autoScroll(s) {
    const d = doc();
    const root = d && (d.scrollingElement || d.documentElement);
    const vh = global.innerHeight > 0 ? global.innerHeight : 0;
    for (let i = 0; i < s.scrollers.length; i++) {
      const el = s.scrollers[i];
      let top, bottom;
      if (el === root) { top = 0; bottom = vh; } else { const vr = dndVisibleRect(el); top = vr.top; bottom = vr.bottom; }
      const h = bottom - top;
      if (!(h > 0)) continue;
      const edge = Math.max(24, Math.min(72, h * 0.2));
      let dy = 0;
      if (s.y < top + edge) dy = -Math.min(1.5, (top + edge - s.y) / edge) * DND_MAX_SPEED;
      else if (s.y > bottom - edge) dy = Math.min(1.5, (s.y - (bottom - edge)) / edge) * DND_MAX_SPEED;
      if (!dy) continue;
      dy = dy < 0 ? Math.floor(dy) : Math.ceil(dy);
      const before = el.scrollTop;
      el.scrollTop = before + dy;
      if (el.scrollTop !== before) return true;
    }
    return false;
  }
  function tick(s) {
    if (st !== s || s.phase !== 'drag') return;
    if (autoScroll(s)) update(s);
  }

  function beginDrag(s) {
    const d = doc();
    s.phase = 'drag';
    if (d.body && d.body.classList) d.body.classList.add('qcl-dnd-on');
    s.row.classList.add('qcl-dnd-src');
    makeGhost(s);
    makeLine(s);
    s.scrollers = dndScrollers(s.row, opts.scrollContainer);
    if (s.scrollers.length && typeof global.setInterval === 'function') s.timer = global.setInterval(() => tick(s), 16);
  }

  /** 收掉所有暫時的畫面元素與監聽（不呼叫 onMove） */
  function teardown(s) {
    if (s.timer && typeof global.clearInterval === 'function') global.clearInterval(s.timer);
    s.timer = null;
    [s.ghost, s.line].forEach((e) => { if (e && e.parentNode) e.parentNode.removeChild(e); });
    s.ghost = s.line = null;
    if (s.row && s.row.classList) s.row.classList.remove('qcl-dnd-src');
    const d = doc();
    if (d && d.body && d.body.classList) d.body.classList.remove('qcl-dnd-on');
    off(gest);
    try { if (s.mode === 'pointer' && s.handle && s.handle.releasePointerCapture && s.pointerId !== undefined) s.handle.releasePointerCapture(s.pointerId); } catch (_) { /* 已釋放 */ }
  }
  /** 拖曳剛結束時瀏覽器仍會對把手送一個 click：吃掉它（只吃這一個） */
  function suppressClick() {
    if (typeof global.setTimeout !== 'function') return;
    const h = (e) => { e.stopPropagation(); e.preventDefault(); };
    container.addEventListener('click', h, true);
    global.setTimeout(() => container.removeEventListener('click', h, true), 0);
  }
  function refocus(s, offset) {
    let h = s.handle;
    if (!(h && h.isConnected !== false && container.contains(h))) {   // 列被重畫了：找回移動後那一列的把手
      const r = flat()[offset + s.to];
      h = r && r.querySelector ? r.querySelector(opts.handleSelector) : null;
    }
    if (h && typeof h.focus === 'function') { try { h.focus({ preventScroll: true }); } catch (_) { h.focus(); } }
  }
  /** 放下：呼叫 onMove（呼叫前狀態已清掉，onMove 可以放心重畫整個列表）；回傳是否真的移動了 */
  function commit(s) {
    if (!s.ok || s.to === s.from || s.row.isConnected === false) return false;
    const offset = flat().indexOf(s.rows[0]);
    try { opts.onMove(s.from, s.to, info(s)); } catch (err) { if (global.console) console.error('QCL.dragSort onMove', err); return false; }
    refocus(s, offset);
    return true;
  }
  function finish(commitIt) {
    const s = st;
    if (!s) return;
    st = null;
    teardown(s);
    if (s.phase === 'pending') return;                   // 沒超過位移門檻＝單純點擊，什麼都不做
    if (s.mode === 'pointer') suppressClick();
    const label = labelFor(s.row, s.from);
    if (commitIt && commit(s)) dndAnnounce('已放下「' + label + '」，現在是第 ' + (s.to + 1) + ' 列，共 ' + s.rows.length + ' 列。');
    else if (commitIt && s.ok) dndAnnounce('已放下「' + label + '」，位置沒有改變（第 ' + (s.from + 1) + ' 列）。');
    else dndAnnounce('已取消移動，「' + label + '」仍在第 ' + (s.from + 1) + ' 列。');
  }

  // ── 指標（滑鼠／觸控／手寫筆）──
  function onPointerDown(ev) {
    if (dead || st) return;
    if (ev.button !== undefined && ev.button !== 0) return;
    if (ev.isPrimary === false) return;
    const t = ev.target;
    const handle = t && t.closest ? t.closest(opts.handleSelector) : null;
    if (!handle || !container.contains(handle) || handle.disabled) return;
    const row = handle.closest(opts.rowSelector);
    if (!row || !container.contains(row)) return;
    const rows = rowsFor(row);
    const from = rows.indexOf(row);
    if (from < 0) return;
    const s = st = {
      mode: 'pointer', phase: 'pending', handle, row, rows, from, to: from, slot: from, ok: true,
      pointerId: ev.pointerId, pointerType: ev.pointerType || 'mouse', startX: ev.clientX, startY: ev.clientY, x: ev.clientX, y: ev.clientY,
      ghost: null, line: null, timer: null, scrollers: [], rects: null,
    };
    try { if (handle.setPointerCapture && ev.pointerId !== undefined) handle.setPointerCapture(ev.pointerId); } catch (_) { /* 不支援就靠 document 監聽 */ }
    const d = doc();
    on(gest, d, 'pointermove', onPointerMove, true);
    on(gest, d, 'pointerup', onPointerUp, true);
    on(gest, d, 'pointercancel', onPointerCancel, true);
    on(gest, d, 'contextmenu', onAbort, true);
    on(gest, d, 'selectstart', onSelectStart, true);
    on(gest, d, 'scroll', onScroll, true);
    on(gest, global, 'keydown', onGestureKey, true);
    on(gest, global, 'blur', onAbort);
    return s;
  }
  function onPointerMove(ev) {
    const s = st;
    if (!s || s.mode !== 'pointer' || ev.pointerId !== s.pointerId) return;
    s.x = ev.clientX; s.y = ev.clientY;
    if (s.phase === 'pending') {
      const dx = s.x - s.startX, dy = s.y - s.startY;
      if (dx * dx + dy * dy < threshold * threshold) return;
      beginDrag(s);
    }
    ev.preventDefault();
    update(s);
  }
  function onPointerUp(ev) {
    const s = st;
    if (!s || s.mode !== 'pointer' || ev.pointerId !== s.pointerId) return;
    s.x = ev.clientX; s.y = ev.clientY;
    if (s.phase === 'drag') { ev.preventDefault(); update(s); finish(true); } else finish(false);
  }
  function onPointerCancel(ev) {
    const s = st;
    if (!s || s.mode !== 'pointer' || ev.pointerId !== s.pointerId) return;
    finish(false);
  }
  function onSelectStart(ev) { if (st && st.mode === 'pointer' && st.phase === 'drag') ev.preventDefault(); }
  function onScroll() { if (st && st.phase !== 'pending') repaint(st); }
  function onAbort() { if (st) finish(false); }
  /** 手勢進行中的鍵盤：Esc 取消（並擋下，不讓外層的 Esc 關掉對話框）；鍵盤移動模式的方向鍵／空白鍵／Enter 也在這裡處理 */
  function onGestureKey(ev) {
    const s = st;
    if (!s) return;
    const k = ev.key;
    if (k === 'Escape' || k === 'Esc') { stopEv(ev); finish(false); return; }
    if (s.mode !== 'kbd') return;
    if (k === 'ArrowUp' || k === 'ArrowDown' || k === 'Home' || k === 'End') {
      stopEv(ev);
      const n = s.rows.length;
      const to = k === 'ArrowUp' ? s.to - 1 : k === 'ArrowDown' ? s.to + 1 : k === 'Home' ? 0 : n - 1;
      moveKbd(s, Math.max(0, Math.min(n - 1, to)));
    } else if (k === ' ' || k === 'Spacebar' || k === 'Enter') {
      stopEv(ev);
      if (!ev.repeat) finish(true);
    }
  }

  // ── 鍵盤移動模式：把手聚焦時 空白鍵拿起 → 上下鍵移動 → 空白鍵／Enter 放下，Esc 取消；全程用 aria-live 播報 ──
  function moveKbd(s, to) {
    s.to = to;
    s.slot = to > s.from ? to + 1 : to;
    const tr = s.rows[to];
    if (tr && tr.scrollIntoView && to !== s.from) tr.scrollIntoView({ block: 'nearest' });
    s.rects = s.rows.map((r) => r.getBoundingClientRect());
    judge(s);
    paintLine(s);
    dndAnnounce(s.ok ? '目標位置：第 ' + (to + 1) + ' 列，共 ' + s.rows.length + ' 列。' : '第 ' + (to + 1) + ' 列不能放在這裡。');
  }
  function onKeyDown(ev) {
    if (dead || st || opts.keyboard === false) return;
    if (ev.key !== ' ' && ev.key !== 'Spacebar') return;
    if (ev.ctrlKey || ev.metaKey || ev.altKey || ev.shiftKey) return;
    const t = ev.target;
    const handle = t && t.closest ? t.closest(opts.handleSelector) : null;
    if (!handle || !container.contains(handle)) return;
    const row = handle.closest(opts.rowSelector);
    if (!row || !container.contains(row)) return;
    ev.preventDefault();
    const rows = rowsFor(row);
    const from = rows.indexOf(row);
    if (from < 0) return;
    if (rows.length < 2) { dndAnnounce('只有一列，沒有可以移動的位置。'); return; }
    const s = st = { mode: 'kbd', phase: 'kbd', handle, row, rows, from, to: from, slot: from, ok: true, ghost: null, line: null, timer: null, rects: null };
    row.classList.add('qcl-dnd-src');
    makeLine(s);
    on(gest, global, 'keydown', onGestureKey, true);
    on(gest, global, 'blur', onAbort);
    on(gest, doc(), 'scroll', onScroll, true);
    on(gest, handle, 'blur', onAbort);
    s.rects = rows.map((r) => r.getBoundingClientRect());
    judge(s);
    paintLine(s);
    dndAnnounce('已拿起第 ' + (from + 1) + ' 列「' + labelFor(row, from) + '」，共 ' + rows.length + ' 列。按上下鍵移動，空白鍵或 Enter 放下，Esc 取消。');
  }

  on(base, container, 'pointerdown', onPointerDown);
  on(base, container, 'keydown', onKeyDown);
  on(base, container, 'contextmenu', (ev) => {
    const t = ev.target;
    if (t && t.closest && t.closest(opts.handleSelector)) ev.preventDefault();   // 觸控長按把手不要跳出右鍵選單
  });

  return {
    isDragging: () => !!st && st.phase !== 'pending',
    cancel: () => { if (st) finish(false); },
    destroy: () => {
      if (dead) return;
      if (st) { const s = st; st = null; teardown(s); }
      dead = true;
      off(base);
    },
  };
}

/** 成本明細：把同一區的某一列移到最終索引 to（只動該列 DOM，其餘列與輸入框完全不動），再走和 ▲▼ 一樣的 emitChange */
function dndMoveRow(inst, from, to, info) {
  if (inst.dead || from === to) return;
  const rows = info.rows, tr = rows[from];
  if (!tr || !tr.parentNode) return;
  const rest = rows.filter((r, i) => i !== from);
  if (!rest.length) return;
  tr.parentNode.insertBefore(tr, to < rest.length ? rest[to] : rest[rest.length - 1].nextSibling);
  setMsg(inst, '');
  emitChange(inst);
}

// ── 對外：mount / collect ───────────────────────────
function findInst(el) {
  for (let n = el; n; n = n.parentNode) if (INST.has(n)) return INST.get(n);
  return null;
}

/**
 * 掛載編輯器。opts = { lines, items, mode:'edit'|'view', revenue, maxLines, onChange(lines, totals), confirm, classCodes, canSeePrice }
 *   lines 沒給（或不是陣列）：edit 模式用 items 產生預設種子，view 模式視為沒有明細。
 *   revenue：折扣後未稅營收（元），有值就即時重算印花稅；canSeePrice 保留為可選參數，不再影響印花稅顯示（規格 v1.1）。
 * 回傳 { getLines(), setLines(lines), setItems(items), setRevenue(n), collect(), destroy() }。
 */
function mount(el, opts) {
  if (!el || !el.nodeType) throw new TypeError('QCL.mount: 需要容器元素');
  opts = opts || {};
  const old = INST.get(el);
  if (old) old.destroy();

  const mode = opts.mode === 'view' ? 'view' : 'edit';
  const inst = {
    id: ++uid, el, opts, mode, root: null, dead: false,
    items: Array.isArray(opts.items) ? opts.items : null,
    revenue: opts.revenue,
    max: Math.max(1, Math.min(MAX_LINES, parseInt(opts.maxLines, 10) || MAX_LINES)),
    lines: [],
  };
  const initial = Array.isArray(opts.lines) ? normalize(opts.lines) : (mode === 'edit' ? seedFromItems(inst.items, opts) : []);

  const handlers = [];
  if (mode === 'edit') {
    const on = (type, fn) => { const h = (ev) => fn(inst, ev); el.addEventListener(type, h); handlers.push([type, h]); };
    on('click', onClick);
    on('input', onInput);
    on('change', onChangeEv);
    on('keydown', onKeydown);
    renderEdit(inst, initial);
    // 拖曳排序：只在編輯模式；限同一分區內（rowsOf 只回傳同一個 tbody 的列），走和 ▲▼ 一樣的 emitChange
    inst.dnd = dragSort(el, {
      rowSelector: 'tr.qcl-row',
      handleSelector: '.qcl-drag',
      rowsOf: (row) => Array.prototype.filter.call(row.parentNode.children, (c) => c.classList && c.classList.contains('qcl-row')),
      onMove: (from, to, info) => dndMoveRow(inst, from, to, info),
      labelOf: (row) => { const d = row.querySelector('.qcl-desc'); return d ? d.value : ''; },
    });
  } else {
    inst.lines = initial;
    renderView(inst, initial);
  }

  inst.getLines = function () {
    if (inst.dead) return [];
    return mode === 'edit' ? readDom(inst, false).lines : outLines(inst.lines);
  };
  inst.setLines = function (lines) {
    if (inst.dead) return;
    const ls = normalize(lines);
    if (mode === 'edit') renderEdit(inst, ls); else { inst.lines = ls; renderView(inst, ls); }
  };
  inst.setItems = function (items) {
    inst.items = Array.isArray(items) ? items : null;
    if (inst.dead) return;
    if (mode === 'edit') refresh(inst); else renderView(inst, inst.lines);   // 唯讀模式的「原報價品項已刪除」徽章依 items 產生，要重畫
  };
  inst.setRevenue = function (rev) {
    inst.revenue = rev;
    if (inst.dead) return;
    if (mode === 'edit') refresh(inst); else renderView(inst, inst.lines);
  };
  inst.collect = function () { return collectInst(inst); };
  inst.destroy = function () {
    if (inst.dead) return;
    inst.dead = true;
    if (inst.dnd) inst.dnd.destroy();
    handlers.forEach((h) => el.removeEventListener(h[0], h[1]));
    el.innerHTML = '';
    if (INST.get(el) === inst) INST.delete(el);
  };
  INST.set(el, inst);
  return inst;
}

/** 輸出用的列：印花稅列不帶 unitCost（那只是伺服器給的顯示金額，伺服器會忽略並自行計算） */
function outLines(lines) {
  return lines.map((l) => {
    if (l.auto !== 'stamp') return Object.assign({}, l);
    const o = Object.assign({}, l);
    delete o.unitCost;
    return o;
  });
}

function collectInst(inst) {
  if (inst.dead || !inst.root) return { lines: [], invalid: 0, blankDesc: 0, zeroCost: 0, zeroQty: 0 };
  if (inst.mode === 'view') {
    const ls = inst.lines.filter((l) => l.auto !== 'stamp');
    return {
      lines: outLines(inst.lines), invalid: 0,
      blankDesc: ls.filter((l) => !l.desc).length, zeroCost: ls.filter((l) => l.unitCost === 0).length, zeroQty: ls.filter((l) => l.qty === 0).length,
    };
  }
  const rd = readDom(inst, true);
  refresh(inst, rd);
  return { lines: rd.lines, invalid: rd.invalid, blankDesc: rd.blankDesc, zeroCost: rd.zeroCost, zeroQty: rd.zeroQty };
}

/** 從 DOM 讀回 lines；空白項目的列保留並標紅、數字非法標紅（qcl-bad），是否擋存由呼叫端依 invalid／blankDesc 決定 */
function collect(el) {
  const inst = findInst(el);
  return inst ? collectInst(inst) : { lines: [], invalid: 0, blankDesc: 0, zeroCost: 0, zeroQty: 0 };
}

global.QCL = {
  CATS,
  MAX_LINES,
  seedFromItems,
  normalize,
  totals,
  mount,
  collect,
  dragSort,                // (container, opts) → { destroy, cancel, isDragging }：拖曳排序核心（成本明細與報價項目表共用）；見該函式上方說明
  dragTargetIndex,         // (rects, y[, fromIdx]) → 縫隙 0..n；給 fromIdx 則回傳移動後的最終索引（純函式）
  dragLineY,               // (rects, slot) → 插入線的 y（純函式）
  // 輔助（下一階段接線與測試用）
  esc,
  fmtMoney,
  stampAmount,
  revenueOf,
  lineAmount,
  unmatchedItemsNote,
  unmatchedPricedItems,    // (lines, items) → [{lid, name, index}]：有價品項沒有成本列涵蓋者（＝伺服器 CL.unmatchedItems）
  unmatchedPricedNote,     // (lines, items) → 提醒文字（＝伺服器 preview.warnings 那句）；沒有未涵蓋品項回 ''
  unmatchedKeys,           // (lines, items) → 未涵蓋有價品項的鍵陣列（lid／nid）：表單載入時記成儲存前確認的基準
  unmatchedNewNote,        // (lines, items, baseline) → 只算基準以外新出現的未涵蓋品項的提醒文字；沒有回 ''（儲存前確認用）
  STAMP_DESC,
  // 內部（單元測試用）
  _unmatchedMessage: unmatchedMessage,
  _isOrphanLine: (l, items) => isOrphanLine(l, itemLidSet(items)),
  _missingItemLines: missingItemLines,
  _supplementMessage: supplementMessage,
  _mergeLines: mergeLines,
  _catByUnit: catByUnit,
  _editRowHtml: (l, cat) => editRowHtml(cleanLine(Object.assign({}, l, { cat })), CAT_BY_KEY[cat], { desc: 'd', unit: 'u' }),
  _viewRowHtml: (l, cat, items) => viewRowHtml(cleanLine(Object.assign({}, l, { cat })), CAT_BY_KEY[cat], itemLidSet(items)),
};
})(typeof window !== 'undefined' ? window : globalThis);
