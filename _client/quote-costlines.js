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
//   QCL.mount(el, opts)                     掛載編輯器（mode: 'edit' | 'view' | 'consultant'）→ {getLines, setLines, setItems, setRevenue, setLinkInfo, setNames, peek, peekCat, collect, destroy}
//                                           peekCat(cat)＝只看某一分區的列數／小計／各項檢查（逐步填寫 quote-coststeps.js 用）
//                                           mode:'consultant'＝顧問對話框：edit 加上每列「對應」欄（連動報價數量／拆項／不對應）與孤兒對應提示；opts.consultantNames＝顧問姓名欄的建議清單
//   cost-sync（顧問成本畫面連動業務報價，與伺服器 lib/quoteCostLines.js 逐例鏡像，scripts/check-quote-costsync-ui.js 以 vm 隨機比對）：
//   QCL.applyLinks(items, lines, newItems) / materializeItems / decimalSum   連動計算：rel='link' 的列把 Σ數量／單位寫回對應的報價品項（單位不一致＝衝突、加總 0＝zero）
//   QCL.isOutsourced / outsourcedStats / outsourcedFromBreakdown              委外占比：顧問服務區有填委外廠商的列；委外 ÷ 專案總成本（含差旅／交際費／印花稅，不含風險預留）、委外 ÷ 顧問服務成本
//   QCL.liveSummary / outsourcedCardModel / outsourcedCardHtml / paintOutsourcedCard   顧問對話框頂端卡片的數字（與伺服器草稿試算逐分相同）與第五張「委外佔比」卡
//   QCL.syncConfirmText(changes)            按「完成」前的「將同步更新業務的報價」清單文字
//   牌價簿（報價牌價簿，管理員在後台維護的顧問角色人天成本；mount opts.pricebook = [{name, cost, bu?}]，沒給＝行為與以前完全相同；牌價簿依 BU 分開、同名可存在於不同 BU）：
//   QCL.pbClean / pbDescSuggest / pbDefault   清洗清單／顧問「項目」建議清單（牌價簿角色名＋DESC_SUGGEST，不分大小寫去重）／預設成本規則（純函式）。
//                                           規則：顧問服務區、品名（去頭尾空白、不分大小寫）與牌價簿完全相同、成本單價空白或 0、沒有委外廠商 → 帶入牌價簿成本（單位空白或預設「式」且非連動列才改「人天」）；
//                                           只在使用者「改完品名」（change／從建議清單選取）時觸發，絕不覆蓋已填的值，載入已存的成本列時完全不動；
//                                           opts.pricebookSeed:true 時種子（含重新帶入／補入）的顧問品項列也帶入成本（單位維持品項原樣）。inst.setPricebook(list) 可在載入完成後補設。
//                                           BU 比對（名稱只在同一個 BU 內唯一）：mount opts.pricebookBu（報價單的 BU，可省略；目前報價單沒有可靠的 BU 欄位，呼叫端沒傳）已知 → 先在該 BU 內找完全相符；
//                                           BU 不明、或該 BU 內沒有 → 只有「全部 BU 裡剛好一筆」同名才帶入；同名出現在 2 個以上 BU（有歧義）→ 不帶入。建議清單的名稱跨 BU 去重。
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
const CONSULTANT_MAX = 40;   // 顧問姓名（只有顧問服務區），與伺服器 LIMITS.consultant 相同
const QTY_MAX = 1e9, COST_MAX = 1e12;
const DEFAULT_UNIT = '式';
const STAMP_DESC = '印花稅(合約金額×0.1%)';
const STAMP_RATE = 0.001;
const SEED_TRAVEL_DESC = '差旅交通';
const SEED_ENTERTAIN_DESC = '交際費';
// 成本列與報價品項的對應方式（cost-sync，規格 §2）：link＝顧問改這列的單位／數量，按「完成」時寫回對應的報價品項；split＝對應某品項但不連動；none＝純成本
const RELS = Object.freeze(['link', 'split', 'none']);
const REL_LABELS = Object.freeze({ link: '連動報價數量', split: '拆項（報價不動）', none: '不對應（純成本）' });
const MIN_ITEM_QTY = 0.001;   // 與伺服器／業務存檔的報價品項數量下限相同
const MAX_NEW_ITEMS = 20;     // 顧問在草稿裡新增的報價項目上限（伺服器 MAX_NEW_ITEMS）

// 五區（對應 PNL 的 1 顧問成本／2 軟體成本／3 硬體成本／4 差旅／5 其他費用）
// hasConsultant：只有顧問服務區有「顧問姓名」欄（自家顧問填姓名、委外填「委外廠商」；兩欄可同時有值）
const CATS = Object.freeze([
  Object.freeze({ key: 'consult',  name: '顧問服務成本', hasVendor: true,  hasConsultant: true,  vendorLabel: '委外廠商', vendorHint: '自家顧問免填', hasNote: true, descHint: '角色（例：PM）' }),
  Object.freeze({ key: 'software', name: '軟體成本',     hasVendor: true,  hasConsultant: false, vendorLabel: '供應商',   vendorHint: '供應商',       hasNote: true, descHint: '品名' }),
  Object.freeze({ key: 'hw',       name: '硬體成本',     hasVendor: true,  hasConsultant: false, vendorLabel: '供應商',   vendorHint: '供應商',       hasNote: true, descHint: '品名' }),
  Object.freeze({ key: 'travel',   name: '差旅費用',     hasVendor: false, hasConsultant: false, vendorLabel: '',         vendorHint: '',             hasNote: true, descHint: '項目（例：差旅交通）' }),
  Object.freeze({ key: 'other',    name: '其他費用',     hasVendor: false, hasConsultant: false, vendorLabel: '',         vendorHint: '',             hasNote: true, descHint: '項目' }),
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

/**
 * 單列金額（分，整數）：與伺服器同規則——每列先取整到分再加總。
 * 有 BigInt 時用十進位精確運算（＝伺服器 centsOf：round-half-up(qty×unitCost×100)），與伺服器逐分相同（卡片上的成本／毛利／委外占比要和伺服器試算一致）；
 * 沒有 BigInt 的舊瀏覽器退回浮點估算。
 */
function lineCents(l) {
  const q = clampNum(l && l.qty, QTY_MAX, 0), c = clampNum(l && l.unitCost, COST_MAX, 0);
  if (HAS_BIGINT) {
    const a = decParts(q), b = decParts(c);
    if (a && b) return Number(divRound(a.n * b.n * BigInt(100), pow10(a.s + b.s)));
  }
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
  // rel：對應方式（只認 link／split／none；印花稅列固定不對應）。none＝純成本，沒有對應目標：forLid／forLids 一律丟掉（同伺服器）
  const rel = !stamp && typeof raw.rel === 'string' && RELS.indexOf(raw.rel) >= 0 ? raw.rel : '';
  const forLid = rel === 'none' ? '' : text(raw.forLid, LID_MAX);
  if (forLid) o.forLid = forLid;
  if (!stamp) {
    // forLids：合併為一列時帶著被併各列的 forLid∪forLids，只供涵蓋判定（與伺服器同規則：字串、去頭尾空白、截 64、去空白與重複、剔除與 forLid 相同者、上限 60）
    const fl = [];
    if (rel !== 'none' && Array.isArray(raw.forLids)) {
      const seen = Object.create(null);
      for (let i = 0; i < raw.forLids.length && fl.length < MAX_FOR_LIDS; i++) {
        const e = raw.forLids[i];
        if (typeof e !== 'string') continue;
        const s = e.trim().slice(0, LID_MAX);
        if (s && s !== forLid && !seen[s]) { seen[s] = true; fl.push(s); }
      }
    }
    if (fl.length) o.forLids = fl;
    // 顧問姓名：只有顧問服務區的列保留（其他分類丟掉）；有值才輸出，沒用到的列與改版前的輸出位元級相同
    if (cat === 'consult') { const c = text(raw.consultant, CONSULTANT_MAX); if (c) o.consultant = c; }
    if (rel) o.rel = rel;
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
  // opts.link（顧問對話框）：品項帶出的列預設「連動報價數量」（種子列的品名、單位本來就和品項相同＝規格 §1 的預設規則）；固定列（差旅、交際費）預設「不對應」。
  // 沒開 link（業務自填成本）時不寫 rel，輸出與改版前相同
  const rel = isLinkOpts(opts) ? 'link' : undefined;
  const pb = (opts && opts.pricebookSeed === true) ? cleanPricebook(opts.pricebook) : [];
  const pbBu = cleanBu(opts && opts.pricebookBu);   // 牌價簿預設成本只在呼叫端明確要求時帶入種子（舊式單的種子要和伺服器算的成本一致，不帶）
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
      rel,
    });
    if (line && pb.length) { const d = pricebookDefault(line, pb, true, pbBu); if (d) line.unitCost = d.unitCost; }   // 種子列的單位跟著品項，不改
    if (line && line.desc && !(rel && !line.forLid)) out.push(line);   // link 列一定要有對應目標：沒有 lid／nid 的品項（舊資料）不產生連動列
  });
  return out;
}

/** 顧問對話框（mode:'consultant' 或 link:true）：種子的品項列預設「連動」、固定列預設「不對應」 */
function isLinkOpts(opts) { return !!(opts && (opts.link === true || opts.mode === 'consultant')); }

/** 固定種子列：差旅「差旅交通」1 列＋其他「交際費」＋（includeStamp）印花稅 auto 列；link＝顧問對話框（固定列預設「不對應」） */
function seedFixedLines(includeStamp, link) {
  const rel = link ? 'none' : undefined;
  const out = [
    cleanLine({ cat: 'travel', desc: SEED_TRAVEL_DESC, rel }),
    cleanLine({ cat: 'other', desc: SEED_ENTERTAIN_DESC, rel }),
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
  const fixed = seedFixedLines(!(opts && opts.includeStamp === false), isLinkOpts(opts));
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
.qcl-orphbar { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 12px; border-radius: 8px; padding: 7px 12px; font-size: 13px; margin-bottom: 8px; line-height: 1.6;
  background: var(--qcl-bad-bg); color: var(--qcl-bad); border: 1px solid var(--qcl-bad); }
.qcl-orphbar[hidden] { display: none; }
.qcl-linknote .qcl-btn.sm { padding: 1px 8px; font-size: 11.5px; }
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
/* cost-sync：顧問姓名欄、委外標籤、對應欄（顧問對話框）、委外佔比卡片 */
.qcl-table.cn { min-width: 1010px; }
.qcl-table.cnv { min-width: 890px; }   /* 唯讀：沒有控制欄與對應欄，和其他唯讀表格同寬 */
.qcl-table.lk { min-width: 1120px; }
.qcl-table.nv.lk { min-width: 1020px; }
.qcl-table.cn.lk { min-width: 1260px; }
.qcl-table col.qcl-c-rel { width: 240px; }
.qcl-ostag { display: block; width: fit-content; margin: 0 0 3px; padding: 0 7px; font-size: 11px; line-height: 1.6; border-radius: 9px; font-weight: 600;
  color: #3949ab; background: #eef0fb; border: 1px solid #c5cae9; }
.qcl-ostag[hidden] { display: none; }
.qcl-t .qcl-ostag { display: inline-block; margin: 0 6px 0 0; }
.qcl-reltd .qcl-in { margin-bottom: 4px; padding: 4px 6px; font-size: 12.5px; }
.qcl-reltd .qcl-in[hidden] { display: none; }
.qcl-linknote { display: block; font-size: 11.5px; line-height: 1.5; color: var(--qcl-mu); }
.qcl-linknote:empty { display: none; }
.qcl-linknote.chg { color: var(--qcl-pri); font-weight: 600; }
.qcl-linknote.bad { color: var(--qcl-bad); font-weight: 600; }
.qcl-in.qcl-conf { border-color: var(--qcl-bad); background: var(--qcl-bad-bg); }
.pnl-sum-grid.qcl-g5 { grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; }
.qcl-os-card { text-align: center; }
.qcl-os-card:focus-visible { outline: 2px solid var(--qcl-pri, #1a73e8); outline-offset: 2px; }
.qcl-os-pct { font-variant-numeric: tabular-nums; }
.qcl-os-bar { height: 6px; border-radius: 3px; background: #dfe3ea; overflow: hidden; margin: 6px auto; max-width: 170px; }
.qcl-os-fill { display: block; height: 100%; width: 0; border-radius: 3px; background: #5c6bc0; transition: width .25s ease; }
.qcl-os-l1, .qcl-os-l2 { font-size: 11.5px; line-height: 1.55; color: #6b7684; font-variant-numeric: tabular-nums; }
.qcl-os-solo { max-width: 320px; margin: 0 0 10px; }
body.dark .qcl-os-bar { background: #30363d; }
body.dark .qcl-os-fill { background: #8c9eff; }
body.dark .qcl-os-l1, body.dark .qcl-os-l2 { color: #8b949e; }
body.dark .qcl-ostag { color: #aab4ff; background: #1c2240; border-color: #323d73; }
/* 沒資料的分區收合成只剩標題列；說明文字改成一行＋可展開 */
.qcl-sec.qcl-collapsed { padding-top: 7px; padding-bottom: 7px; }
.qcl-sec.qcl-collapsed .qcl-wrap { display: none; }
.qcl-sec.qcl-collapsed .qcl-sec-h { margin-bottom: 0; }
.qcl-sec.qcl-collapsed .qcl-sec-h h4 { font-weight: 500; color: var(--qcl-mu); }
.qcl-help { margin: 0 0 8px; font-size: 12.5px; color: var(--qcl-mu); }
.qcl-help summary { cursor: pointer; }
.qcl-help summary u { color: var(--qcl-pri); text-decoration: none; margin-left: 4px; }
.qcl-help .qcl-help-b { margin: 6px 0 0; line-height: 1.6; }
/* 對應欄：兩個下拉並排，說明文字在下一行（列高從 3 行降到 2 行） */
.qcl-reltd .qcl-in { display: inline-block; width: calc(50% - 3px); margin-right: 3px; margin-bottom: 2px; vertical-align: top; }
.qcl-reltd .qcl-in:nth-of-type(2) { margin-right: 0; }
@media (max-width: 900px) {
  .pnl-sum-grid.qcl-g5 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
}
@media (max-width: 620px) {
  .pnl-sum-grid.qcl-g5 { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
  .pnl-sum-grid.qcl-g5 > .qcl-os-card { grid-column: 1 / -1; }
}
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

// ── 對應（cost-sync）：目標下拉的選項 ─────────────────────────
/**
 * 報價品項 → 對應目標清單 [{key, seq, name, isNew}]：key＝品項 lid（還沒存檔／顧問草稿新增的品項用暫時代號 nid；舊資料沒有 lid 用 'legacy-'+索引，同伺服器），
 * seq＝畫面上的項目編號（標題／小計列不編號），isNew＝顧問在草稿裡新增的報價項目（items 裡的 isDraftNew 旗標）。
 */
function targetEntries(items) {
  const out = [];
  let seq = 0;
  (Array.isArray(items) ? items : []).forEach((it, i) => {
    if (!it || typeof it !== 'object' || isKindRow(it)) return;
    seq++;
    out.push({ key: lidKey(it.lid) || lidKey(it.nid) || ('legacy-' + i), seq, name: nameKey(it.desc), isNew: !!it.isDraftNew });
  });
  return out;
}

const TARGET_LABEL_MAX = 36;
function targetLabel(e) {
  const nm = e.name || '（未命名）';
  return e.seq + '. ' + (nm.length > TARGET_LABEL_MAX ? nm.slice(0, TARGET_LABEL_MAX) + '…' : nm) + (e.isNew ? '　＋新增' : '');
}

/** 對應目標下拉的 <option> 們；selected 不在清單裡（品項已被業務刪除）時多放一個已選取的「已刪除」選項，讓畫面看得出來而不是默默跳到別的品項 */
function targetOptionsHtml(entries, selected) {
  const sel = lidKey(selected);
  const list = Array.isArray(entries) ? entries : [];
  let h = '<option value=""' + (sel ? '' : ' selected') + '>請選擇報價品項…</option>';
  h += list.map((e) => '<option value="' + esc(e.key) + '"' + (e.key === sel ? ' selected' : '') + '>' + esc(targetLabel(e)) + '</option>').join('');
  if (sel && !list.some((e) => e.key === sel)) h += '<option value="' + esc(sel) + '" selected>（原報價品項已刪除）</option>';
  return h;
}

/** 下拉選項只顯示短字（連動／拆項／不對應），完整說明放 title（滑鼠停留可看）；確認視窗等訊息文字仍用完整的 REL_LABELS */
const REL_SHORT = Object.freeze({ link: '連動', split: '拆項', none: '不對應' });
function relOptionsHtml(rel) {
  return RELS.map((r) => '<option value="' + r + '"' + (r === rel ? ' selected' : '') + ' title="' + esc(REL_LABELS[r]) + '">' + esc(REL_SHORT[r]) + '</option>').join('');
}

/** 畫面上這列的對應方式：有 rel 用 rel；舊資料沒有 rel：有 forLid／forLids 視為拆項、沒有視為不對應（規格 §2） */
function displayRel(l) { return l.rel || (lineLidList(l).length ? 'split' : 'none'); }

// ── 列／分區 HTML ───────────────────────────────────
const OS_TAG_TITLE = '有填委外廠商的顧問列：計入委外占比';
/** 委外標籤（列首小標籤）；非委外列在編輯模式先放一個隱藏的，refresh 依廠商欄切換 */
function osTagHtml(on, hiddenSlot) {
  if (!on && !hiddenSlot) return '';
  return '<span class="qcl-ostag"' + (on ? '' : ' hidden') + ' title="' + OS_TAG_TITLE + '">委外</span>';
}

/**
 * 編輯列。ctx（選填）＝{ link: 顧問對話框（顯示「對應」欄）, entries: targetEntries(items), ids: datalist 代號 }；沒給 ctx 的輸出與改版前相同（沒有對應欄、顧問區多一個「顧問姓名」欄與委外標籤）。
 */
function editRowHtml(l, def, ids, ctx) {
  const link = !!(ctx && ctx.link);
  const attrs = (l.lid ? ' data-lid="' + esc(l.lid) + '"' : '') + (l.forLid ? ' data-forlid="' + esc(l.forLid) + '"' : '') +
    (l.forLids && l.forLids.length ? ' data-forlids="' + esc(JSON.stringify(l.forLids)) + '"' : '') + (l.rel ? ' data-rel="' + esc(l.rel) + '"' : '');
  const lab = (label) => ' aria-label="' + esc(def.name + '：' + label) + '"';
  let h = '<tr class="qcl-row"' + attrs + '>' +
    '<td>' + (def.hasConsultant ? osTagHtml(isOutsourced(l), true) : '') +
      '<input type="text" class="qcl-in qcl-desc" maxlength="' + DESC_MAX + '" value="' + esc(l.desc) + '" placeholder="' + esc(def.descHint) + '"' +
      (def.key === 'consult' ? ' list="' + ids.desc + '"' : '') + lab('項目') + '>' + ORPHAN_BADGE_HIDDEN + '</td>';
  if (def.hasConsultant) {
    h += '<td><input type="text" class="qcl-in qcl-consultant" maxlength="' + CONSULTANT_MAX + '" value="' + esc(l.consultant || '') + '" placeholder="自家顧問姓名"' +
      (ids.names ? ' list="' + ids.names + '"' : '') + lab('顧問姓名') + '></td>';
  }
  if (def.hasVendor) {
    h += '<td><input type="text" class="qcl-in qcl-vendor" maxlength="' + VENDOR_MAX + '" value="' + esc(l.vendor) + '" placeholder="' + esc(def.vendorHint) + '"' + lab(def.vendorLabel) + '></td>';
  }
  h += '<td><input type="text" class="qcl-in qcl-note" maxlength="' + NOTE_MAX + '" value="' + esc(l.note) + '" placeholder="說明（選填）"' + lab('說明') + '></td>' +
    '<td><input type="text" class="qcl-in qcl-unit" maxlength="' + UNIT_MAX + '" value="' + esc(l.unit) + '" list="' + ids.unit + '"' + lab('單位') + '></td>' +
    '<td><input type="number" class="qcl-in qcl-qty" min="0" step="any" inputmode="decimal" value="' + esc(String(l.qty)) + '"' + lab('數量') + '></td>' +
    '<td><input type="number" class="qcl-in qcl-cost" min="0" step="any" inputmode="decimal" placeholder="0" value="' + (l.unitCost ? esc(String(l.unitCost)) : '') + '"' + lab('成本單價') + '></td>' +
    '<td class="r qcl-sub">' + fmtMoney(lineAmount(l)) + '</td>';
  if (link) {
    const disp = displayRel(l);
    const tgt = l.forLid || (l.forLids && l.forLids[0]) || '';
    h += '<td class="qcl-reltd"><select class="qcl-in qcl-rel"' + lab('對應方式') + '>' + relOptionsHtml(disp) + '</select>' +
      '<select class="qcl-in qcl-target"' + (disp === 'none' ? ' hidden' : '') + lab('對應的報價品項') + '>' + targetOptionsHtml(ctx.entries, tgt) + '</select>' +
      '<span class="qcl-linknote" aria-live="polite"></span></td>';
  }
  return h + '<td class="qcl-ctl">' +
      '<button type="button" class="qcl-drag" title="拖曳排序（區內）；也可用鍵盤：按空白鍵拿起，上下鍵移動，空白鍵放下，Esc 取消" aria-label="拖曳排序">&#8942;&#8942;</button>' +
      '<button type="button" class="qcl-mv" data-act="up" title="上移" aria-label="上移">&#9650;</button>' +
      '<button type="button" class="qcl-mv" data-act="down" title="下移" aria-label="下移">&#9660;</button>' +
      '<button type="button" class="qcl-mv del" data-act="del" title="移除此列" aria-label="移除此列">&#10005;</button></td></tr>';
}

/** 到「小計」欄之前的欄數：項目［顧問姓名］［廠商］說明 單位 數量 成本單價 */
function leadCols(def) { return 1 + (def.hasConsultant ? 1 : 0) + (def.hasVendor ? 1 : 0) + 4; }

/** 欄寬定義（fixed layout）：文字欄均分剩餘寬度，單位／數量／單價／小計／對應／控制欄固定寬，各分區的數字欄上下對齊 */
function colGroup(def, withCtl, link) {
  return '<colgroup><col>' + (def.hasConsultant ? '<col>' : '') + (def.hasVendor ? '<col>' : '') + '<col><col class="qcl-c-unit"><col class="qcl-c-qty"><col class="qcl-c-cost"><col class="qcl-c-sub">' +
    (withCtl && link ? '<col class="qcl-c-rel">' : '') + (withCtl ? '<col class="qcl-c-ctl">' : '') + '</colgroup>';
}

function tableHead(def, link) {
  return '<thead><tr><th>項目</th>' + (def.hasConsultant ? '<th>顧問姓名</th>' : '') + (def.hasVendor ? '<th>' + esc(def.vendorLabel) + '</th>' : '') +
    '<th>說明</th><th>單位</th><th class="r">數量</th><th class="r">成本單價</th><th class="r">小計</th>' + (link ? '<th>對應</th>' : '') + '<th class="qcl-ctl-h"></th></tr></thead>';
}

function tableClass(def, link, view) { return 'qcl-table' + (def.hasVendor ? '' : ' nv') + (def.hasConsultant ? (view ? ' cnv' : ' cn') : '') + (link ? ' lk' : ''); }

function editSectionHtml(def, rows, stamp, ids, ctx) {
  const link = !!(ctx && ctx.link);
  const cols = leadCols(def);
  let h = '<section class="qcl-sec" data-sec="' + def.key + '">' +
    '<div class="qcl-sec-h"><h4>' + esc(def.name) + '</h4><span class="qcl-sec-n"></span><span class="qcl-sec-tot"></span><span class="qcl-sp"></span>' +
    '<button type="button" class="qcl-btn sm" data-act="merge" data-cat="' + def.key + '" title="把這一區的所有列合併成 1 列（成本合計不變），例如整包" disabled>合併為一列</button>' +
    '<button type="button" class="qcl-btn sm" data-act="add" data-cat="' + def.key + '">＋新增</button></div>' +
    '<div class="qcl-wrap"><table class="' + tableClass(def, link) + '">' + colGroup(def, true, link) + tableHead(def, link) +
    '<tbody data-cat="' + def.key + '">' + rows.map((l) => editRowHtml(l, def, ids, ctx)).join('') + '</tbody>';
  if (def.key === 'other') {
    h += '<tbody class="qcl-stampbody"><tr class="qcl-stamprow"' + (stamp && stamp.lid ? ' data-lid="' + esc(stamp.lid) + '"' : '') +
      (stamp && typeof stamp.unitCost === 'number' ? ' data-amt="' + esc(String(stamp.unitCost)) + '"' : '') + '>' +
      '<td colspan="' + cols + '"><span class="qcl-stamp-name">印花稅</span>' +
      '<label><input type="checkbox" class="qcl-stampchk"' + (stamp ? ' checked' : '') + '>計入（依合約金額×0.1%自動計算）</label>' +
      '<span class="qcl-mu qcl-stamphint"></span></td>' +
      '<td class="r qcl-stampamt">—</td>' + (link ? '<td></td>' : '') + '<td></td></tr></tbody>';
  }
  return h + '<tfoot><tr><td colspan="' + cols + '">小計</td><td class="r qcl-secsub">0</td>' + (link ? '<td></td>' : '') + '<td></td></tr></tfoot></table></div></section>';
}

function viewRowHtml(l, def, lidSet) {
  const dash = '<span class="qcl-mu">—</span>';
  return '<tr class="qcl-row">' +
    '<td class="qcl-t">' + (def.hasConsultant ? osTagHtml(isOutsourced(l), false) : '') + (esc(l.desc) || dash) + (isOrphanLine(l, lidSet) ? ORPHAN_BADGE : '') + '</td>' +
    (def.hasConsultant ? '<td class="qcl-t">' + (esc(l.consultant) || dash) + '</td>' : '') +
    (def.hasVendor ? '<td class="qcl-t">' + (esc(l.vendor) || dash) + '</td>' : '') +
    '<td class="qcl-t">' + (esc(l.note) || dash) + '</td>' +
    '<td>' + esc(l.unit) + '</td>' +
    '<td class="r">' + esc(fmtNum(l.qty, 4)) + '</td>' +
    '<td class="r">' + esc(fmtNum(l.unitCost, 2)) + '</td>' +
    '<td class="r">' + fmtMoney(lineAmount(l)) + '</td></tr>';
}

function viewSectionHtml(def, rows, stamp, eff, lidSet) {
  const cols = leadCols(def);
  const head = '<thead><tr><th>項目</th>' + (def.hasConsultant ? '<th>顧問姓名</th>' : '') + (def.hasVendor ? '<th>' + esc(def.vendorLabel) + '</th>' : '') +
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
    '<div class="qcl-wrap"><table class="' + tableClass(def, false, true) + '">' + colGroup(def, false, false) + head + '<tbody>' + body + '</tbody>' +
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
    // 沒資料的分區先收合成只剩標題列（「其他費用」區固定有印花稅勾選列，不收合）；按「＋新增」加入第一列時由這裡自動展開
    const empty = !counts[def.key] && def.key !== 'other';
    sec.classList.toggle('qcl-collapsed', empty);
    const isEdit = root.getAttribute('data-mode') === 'edit';
    sec.querySelector('.qcl-sec-n').textContent = counts[def.key] ? '（' + counts[def.key] + ' 列）' : (empty ? (isEdit ? '（尚無項目，按「＋新增」加入）' : '（無）') : '');
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
  const res = { lines: [], rows: [], invalid: 0, blankDesc: 0, zeroCost: 0, zeroQty: 0, badTarget: 0 };
  const lidSet = itemLidSet(inst.items);   // 沒給 items → null（不檢查對應目標是否還在）
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
      // 對應方式：顧問對話框用下拉的值（列上有 data-rel 才算「明確指定」：舊資料沒有 rel、使用者沒動過就維持沒有 rel，不會因為打開畫面就被改成拆項）；其他模式原樣帶著列上的 rel
      const relSel = tr.querySelector('.qcl-rel');
      const relRaw = relSel ? (tr.hasAttribute('data-rel') ? relSel.value : '') : tr.getAttribute('data-rel');
      const l = cleanLine({
        lid: tr.getAttribute('data-lid'), cat: def.key, desc: val('.qcl-desc'), vendor: val('.qcl-vendor'), consultant: val('.qcl-consultant'), note: val('.qcl-note'), unit: val('.qcl-unit'),
        qty: q.n, unitCost: c.n, forLid: tr.getAttribute('data-forlid'), forLids, rel: relRaw,
      });
      // 明確選了「連動／拆項」的列一定要有有效的報價品項（伺服器 checkTargets 同規則：對應的品項不存在 → 400）。
      // 部分目標已被刪除（合併過的列）就剔除已不存在的，剩下的當目標；全部不存在或根本沒選 → 擋（badTarget）
      let badT = false;
      if (l.rel === 'link' || l.rel === 'split') {
        const ids = lineLidList(l);
        if (!ids.length) badT = true;
        else if (lidSet) {
          const ok = ids.filter((k) => lidSet.has(k));
          if (!ok.length) badT = true;
          else if (ok.length !== ids.length) { l.forLid = ok[0]; if (ok.length > 1) l.forLids = ok.slice(1); else delete l.forLids; }
        }
      }
      const ts = tr.querySelector('.qcl-target');
      if (ts) ts.classList.toggle('qcl-bad', badT);
      if (badT) res.badTarget++;
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
      res.rows.push({ tr, line: l, bad: q.bad || c.bad, qBad: q.bad, cBad: c.bad, badTarget: badT });
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

/** 編輯器目前的對應目標清單變了（品項改名、新增、刪除）就重建每列「對應」下拉的選項（保留目前選的值）；沒變就不動 */
function syncTargetOptions(inst) {
  if (!inst.link) return;
  inst.entries = targetEntries(inst.items);
  const sig = JSON.stringify(inst.entries.map((e) => [e.key, targetLabel(e)]));
  if (sig === inst.targetSig) return;
  inst.targetSig = sig;
  Array.prototype.forEach.call(inst.root.querySelectorAll('select.qcl-target'), (sel) => {
    const tr = sel.closest('tr.qcl-row');
    const cur = tr ? (tr.getAttribute('data-forlid') || '') : sel.value;
    sel.innerHTML = targetOptionsHtml(inst.entries, cur);
  });
}

/** 連動標示（每個連動列旁的小字）：依 setLinkInfo(applyLinks 結果) 顯示「→ 報價 業務原值 → 新值」／單位衝突／數量為 0；單位衝突的列把單位欄標紅 */
function paintLinkNotes(inst, rd) {
  if (!inst.link) return;
  rd = rd || readDom(inst, false);
  const info = inst.linkInfo;
  const lidSet = itemLidSet(inst.items);
  let orph = 0;
  rd.rows.forEach((r) => {
    const l = r.line;
    const note = r.tr.querySelector('.qcl-linknote');
    const unitIn = r.tr.querySelector('.qcl-unit');
    let msg = '', cls = 'qcl-linknote', conf = false;
    // 孤兒對應：連動／拆項的列，對應的報價品項已不在（業務刪了品項、或顧問移除了自己新增的品項）。伺服器會擋（400 BAD_COST_LINE），所以在這裡提示並給「改成不對應」一鍵處理
    const tids = (l.rel === 'link' || l.rel === 'split') ? lineLidList(l) : [];
    if (tids.length && lidSet && !tids.some((k) => lidSet.has(k))) {
      orph++;
      if (note) {
        note.className = 'qcl-linknote bad';
        const k = 'orph';
        if (note.getAttribute('data-k') !== k) {
          note.setAttribute('data-k', k);
          note.innerHTML = '對應的報價品項已不存在（可能被業務刪除）。請重新選擇對應的品項，或 <button type="button" class="qcl-btn sm" data-act="relNone">改成不對應</button>';
        }
      }
      if (unitIn) unitIn.classList.remove('qcl-conf');
      return;
    }
    if (note && note.getAttribute('data-k')) note.removeAttribute('data-k');
    if (l.rel === 'link') {
      const e = info && info.get(l.forLid);
      if (e) {
        if (e.conflict) { msg = '單位衝突：同一個報價品項的連動列單位需一致（目前有 ' + (e.units || []).join('、') + '）'; cls += ' bad'; conf = true; }
        else if (e.zero) { msg = '連動後數量為 0，請填數量或改成「拆項」'; cls += ' bad'; }
        else if (e.qtyChanged || e.unitChanged) {
          msg = '→ 報價 ' + fmtNum(Number(e.qtyFrom), 4) + (e.unitFrom ? ' ' + e.unitFrom : '') + ' → ' + fmtNum(Number(e.qty), 4) + (e.unit ? ' ' + e.unit : '');
          cls += ' chg';
        } else msg = '→ 報價 ' + fmtNum(Number(e.qty), 4) + (e.unit ? ' ' + e.unit : '') + '（同業務原值）';
        if (e.linkCount > 1 && !e.conflict && !e.zero) msg += '　共 ' + e.linkCount + ' 列連動';
      }
    } else if (l.rel === 'split') {
      const n = lineLidList(l).length;
      msg = '報價不動' + (n > 1 ? '（另涵蓋 ' + (n - 1) + ' 個品項）' : '');
    }
    if (note) { note.className = cls; if (note.textContent !== msg) note.textContent = msg; }
    if (unitIn) unitIn.classList.toggle('qcl-conf', conf);
  });
  const bar = inst.root.querySelector('.qcl-orphbar');
  if (bar) {
    bar.hidden = !orph;
    const n = bar.querySelector('.qcl-orphn');
    if (n && n.textContent !== String(orph)) n.textContent = String(orph);
  }
}

/** 重新計算畫面：列小計、區小計、合計、計數、列控制鈕與新增鈕的停用狀態、空區提示 */
function refresh(inst, rd) {
  const root = inst.root;
  rd = rd || readDom(inst, false);
  rd.rows.forEach((r) => { r.tr.querySelector('.qcl-sub').textContent = r.bad ? '—' : fmtMoney(lineAmount(r.line)); });
  // 「原報價品項已刪除」徽章：items 有提供才會顯示（沒提供＝null＝全部隱藏）
  const orphSet = itemLidSet(inst.items);
  rd.rows.forEach((r) => { const b = r.tr.querySelector('.qcl-orph'); if (b) b.hidden = !isOrphanLine(r.line, orphSet); });
  // 委外標籤：顧問服務區有填委外廠商的列（廠商欄改了就即時切換）
  rd.rows.forEach((r) => { const g = r.tr.querySelector('.qcl-ostag'); if (g) g.hidden = !isOutsourced(r.line); });
  syncTargetOptions(inst);
  paintLinkNotes(inst, rd);
  const t = paint(inst, rd.calcLines);

  CATS.forEach((def) => {
    const tb = root.querySelector('tbody[data-cat="' + def.key + '"]');
    const rows = tb.querySelectorAll('tr.qcl-row');
    const ph = tb.querySelector('tr.qcl-empty');
    if (!rows.length && !ph) tb.insertAdjacentHTML('beforeend', '<tr class="qcl-empty"><td colspan="' + (leadCols(def) + 2 + (inst.link ? 1 : 0)) + '"><span class="qcl-emptytxt">尚無項目，按「＋新增」加入</span></td></tr>');
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
  return { lines: rd.lines, totals: t, badTarget: rd.badTarget };
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

/** 列 HTML 用的共同參數：link（顧問對話框的「對應」欄）、目前的對應目標清單 */
function rowCtx(inst) { return { link: inst.link, entries: inst.entries || [] }; }

function renderEdit(inst, lines) {
  ensureStyle();
  const ids = { desc: 'qclDlDesc' + inst.id, unit: 'qclDlUnit' + inst.id, names: 'qclDlNames' + inst.id };
  const sp = splitLines(lines);
  inst.entries = inst.link ? targetEntries(inst.items) : [];
  inst.targetSig = inst.link ? JSON.stringify(inst.entries.map((e) => [e.key, targetLabel(e)])) : null;
  const ctx = rowCtx(inst);
  const names = inst.names || [];
  inst.el.innerHTML = '<div class="qcl-root" data-mode="edit"' + (inst.link ? ' data-link="1"' : '') + '>' +
    '<div class="qcl-tools">' +
      '<button type="button" class="qcl-btn" data-act="reseed" title="丟掉目前的成本明細，依客戶報價品項重新產生預設列">↻ 由報價品項重新帶入</button>' +
      '<button type="button" class="qcl-btn" data-act="supplement" title="只補入「還沒有對應成本列」的新品項，不動既有列">＋補入新品項</button>' +
      '<span class="qcl-sp"></span><span class="qcl-count"></span></div>' +
    '<details class="qcl-help"><summary>金額單位：新台幣元（未稅）；小計＝數量×成本單價。<u>更多說明</u></summary>' +
      '<p class="qcl-help-b">成本明細與客戶報價品項各自獨立、不必一一對應：多個品項的成本可以併成一列（整包），一個品項也可以拆成多列。' +
      '顧問服務區：自家顧問填「顧問姓名」；委外請在「委外廠商」填廠商名稱（有填廠商的列會標示「委外」並計入委外佔比）。</p></details>' +
    '<div class="qcl-limit" hidden>已達成本明細列數上限（' + inst.max + ' 列），無法再新增；請刪除不需要的列。</div>' +
    (inst.link ? '<div class="qcl-orphbar" hidden role="alert"><span>有 <b class="qcl-orphn">0</b> 列成本明細對應的報價品項已不存在（可能被業務刪除），這種對應無法儲存，請重新選擇對應的品項或改成「不對應」。</span><button type="button" class="qcl-btn sm" data-act="relNoneAll">全部改成不對應</button></div>' : '') +
    '<div class="qcl-msg" role="status" aria-live="polite"></div>' +
    CATS.map((def) => editSectionHtml(def, sp.groups[def.key], def.key === 'other' ? sp.stamp : null, ids, ctx)).join('') +
    grandHtml() +
    '<datalist id="' + ids.desc + '">' + descSuggestFor(inst.pricebook).map((s) => '<option value="' + esc(s) + '"></option>').join('') + '</datalist>' +
    '<datalist id="' + ids.unit + '">' + UNIT_SUGGEST.map((s) => '<option value="' + esc(s) + '"></option>').join('') + '</datalist>' +
    '<datalist id="' + ids.names + '">' + names.map((s) => '<option value="' + esc(s) + '"></option>').join('') + '</datalist>' +
    '</div>';
  inst.root = inst.el.firstElementChild;
  inst.ids = ids;
  refresh(inst);
}

function appendRows(inst, lines) {
  const ctx = rowCtx(inst);
  lines.forEach((l) => {
    const def = CAT_BY_KEY[l.cat];
    inst.root.querySelector('tbody[data-cat="' + def.key + '"]').insertAdjacentHTML('beforeend', editRowHtml(l, def, inst.ids, ctx));
  });
}

function addRow(inst, cat) {
  const def = CAT_BY_KEY[cat];
  if (!def || readDom(inst, false).lines.length >= inst.max) return;
  appendRows(inst, [cleanLine({ cat, rel: inst.link ? 'none' : undefined })]);   // 顧問對話框手動新增的列預設「不對應」（要連動／拆項再自己選品項）
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
  return ask(inst, mergeConfirmText(def, rows, total)).then((ok) => {
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

/** 合併前的確認文字（純函式）：說明合併後的單位／數量、對應方式與委外判定的變化 */
function mergeConfirmText(def, rows, total) {
  const m = mergeLines(rows);
  const linked = !!m && m.rel === 'link';
  let msg = '要把「' + def.name + '」的 ' + rows.length + ' 列合併成 1 列嗎？\n成本合計不變（' + fmtMoney(total) + ' 元），但各列的項目、' + (def.hasConsultant ? '顧問姓名、' : '') + '廠商、說明與數量／單價會併成一列（' +
    (linked ? '單位「' + m.unit + '」、數量 ' + fmtNum(m.qty, 4) + '（各列數量加總）' : '單位「式」、數量 1') + '），合併後可再改名。';
  if (rows.some((l) => l.rel === 'link')) {
    msg += linked ? '\n這幾列都連動同一個報價品項：合併後維持「連動報價數量」，數量為各列加總。'
      : '\n這幾列原本有「連動報價數量」：合併後改為「拆項（報價不動）」，報價品項的數量不再由它們連動。';
  }
  if (def.hasConsultant && rows.some(isOutsourced) && rows.some((l) => !isOutsourced(l))) {
    msg += '\n注意：這幾列有委外也有自家顧問，合併後只要廠商欄有值整列都算委外，委外佔比會以合併後的列判定。';
  }
  return msg;
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
    if (l.rel === 'none') return;   // cost-sync：「不對應（純成本）」的列（差旅、交際費…）不參與涵蓋判定——既不涵蓋指定品項，也不進同名比對池（＝伺服器）
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

// ═════════════════════════════════════════════════
// ── cost-sync：連動計算、委外占比、卡片數字（純函式）────────────────────────
// ⚠ applyLinks／materializeItems／decimalSum／pctText／marginText／outsourcedStats 與伺服器 lib/quoteCostLines.js、lib/quoteApproval.js 的同名函式逐例鏡像：
//   scripts/check-quote-costsync-ui.js 以 vm 載入本檔，隨機輸入逐組比對（applyLinks ≥5000 組、委外占比／卡片數字對照伺服器試算）。改一邊必須同步改另一邊。
// ═════════════════════════════════════════════════
const unitOf = (v) => (String(v === undefined || v === null ? '' : v).trim() || DEFAULT_UNIT);

/** 精確的十進位加總（number 陣列 → number）：不經浮點累加（0.1+0.2 不會變 0.30000000000000004）；非有限數、非正數當 0 */
function decimalSum(nums) {
  const arr = Array.isArray(nums) ? nums : [];
  if (!HAS_BIGINT) return arr.reduce((s, x) => s + (typeof x === 'number' && isFinite(x) && x > 0 ? x : 0), 0);
  const parts = [];
  let maxS = 0;
  arr.forEach((x) => {
    const d = decParts(typeof x === 'number' && isFinite(x) && x > 0 ? x : 0);
    if (!d) return;
    parts.push(d);
    if (d.s > maxS) maxS = d.s;
  });
  let total = BigInt(0);
  parts.forEach((d) => { total += d.n * pow10(maxS - d.s); });
  if (maxS === 0) return Number(total);
  const s = total.toString().padStart(maxS + 1, '0');
  return Number(s.slice(0, s.length - maxS) + '.' + s.slice(s.length - maxS));
}

/**
 * 連動計算（＝伺服器 CL.applyLinks）：對每個報價品項，取所有 rel==='link' 且 forLid／forLids 命中它的成本列（印花稅列不參與）：
 *   數量 ＝ Σ 各列 qty（十進位精確加總；>0 時至少 0.001）；各列單位（去空白、空白當「式」）必須相同才算 unit，否則「單位衝突」（該品項維持原樣並列入 conflicts）。
 *   沒有連動列的品項完全不變。連動加總 ≤0 的品項列入 zero。newItems（暫時品項 [{nid,desc,unit,qty}]）接在既有一般品項之後，同樣被連動列命中時改用連動結果。
 * 回傳 { items:[{lid|nid, index, isNew, desc, unit, qty, unitFrom, qtyFrom, changed, qtyChanged, unitChanged, linkCount, conflict, zero}], changes, conflicts, zero }
 */
function applyLinks(items, lines, newItems) {
  const src = Array.isArray(items) ? items : [];
  const entries = [];
  const byKey = new Map();
  src.forEach((it, i) => {
    if (!it || typeof it !== 'object' || it.kind === 'title' || it.kind === 'subtotal') return;
    const key = lidKey(it.lid) || ('legacy-' + i);
    if (byKey.has(key)) return;   // 重複的 lid 只認第一個
    const e = { key, isNew: false, index: i, desc: String(it.desc === undefined || it.desc === null ? '' : it.desc), unit: it.unit, qty: it.qty, unitFrom: it.unit, qtyFrom: it.qty, links: [], linkCount: 0, conflict: false, zero: false };
    entries.push(e); byKey.set(key, e);
  });
  (Array.isArray(newItems) ? newItems : []).forEach((n) => {
    if (!n || typeof n !== 'object' || typeof n.nid !== 'string' || !n.nid || byKey.has(n.nid)) return;
    const e = { key: n.nid, isNew: true, index: -1, desc: String(n.desc === undefined || n.desc === null ? '' : n.desc), unit: unitOf(n.unit), qty: n.qty, unitFrom: unitOf(n.unit), qtyFrom: n.qty, links: [], linkCount: 0, conflict: false, zero: false };
    entries.push(e); byKey.set(n.nid, e);
  });
  (Array.isArray(lines) ? lines : []).forEach((l) => {
    if (!l || typeof l !== 'object' || l.auto === 'stamp' || l.rel !== 'link') return;
    const seen = new Set();
    lineLidList(l).forEach((k) => {
      if (seen.has(k)) return;
      seen.add(k);
      const e = byKey.get(k);
      if (e) e.links.push(l);
    });
  });
  const changes = [], conflicts = [], zero = [];
  const idOf = (e) => (e.isNew ? { nid: e.key } : { lid: e.key });
  const out = entries.map((e) => {
    let qtyChanged = false, unitChanged = false;
    if (e.links.length) {
      e.linkCount = e.links.length;
      const units = [];
      e.links.forEach((l) => { const u = unitOf(l.unit); if (units.indexOf(u) < 0) units.push(u); });
      if (units.length > 1) {
        e.conflict = true;
        conflicts.push(Object.assign(idOf(e), { desc: e.desc, units }));
      } else {
        const sum = decimalSum(e.links.map((l) => l.qty));
        const newQty = sum > 0 ? Math.max(MIN_ITEM_QTY, sum) : 0;
        if (!(newQty > 0)) e.zero = true;
        qtyChanged = Number(e.qtyFrom) !== newQty;
        unitChanged = String(e.unitFrom === undefined || e.unitFrom === null ? '' : e.unitFrom).trim() !== units[0];
        e.qty = newQty; e.unit = units[0];
      }
    } else if (e.isNew && !(Number(e.qty) > 0)) e.zero = true;
    if (e.zero) zero.push(Object.assign(idOf(e), { desc: e.desc }));
    return Object.assign(idOf(e), {
      index: e.index, isNew: e.isNew, desc: e.desc, unit: e.unit, qty: e.qty, unitFrom: e.unitFrom, qtyFrom: e.qtyFrom,
      changed: e.isNew || qtyChanged || unitChanged, qtyChanged, unitChanged, linkCount: e.linkCount, conflict: e.conflict, zero: e.zero,
    });
  });
  out.forEach((o) => {
    const id = o.isNew ? { nid: o.nid } : { lid: o.lid };
    if (o.isNew) { changes.push(Object.assign(id, { desc: o.desc, field: 'new', from: null, to: o.qty, unit: o.unit })); return; }
    if (o.zero || o.conflict) return;
    if (o.qtyChanged) changes.push(Object.assign({}, id, { desc: o.desc, field: 'qty', from: o.qtyFrom, to: o.qty }));
    if (o.unitChanged) changes.push(Object.assign({}, id, { desc: o.desc, field: 'unit', from: o.unitFrom === undefined || o.unitFrom === null ? '' : o.unitFrom, to: o.unit }));
  });
  return { items: out, changes, conflicts, zero };
}

/** 把 applyLinks 的結果套到報價品項陣列，回傳新陣列（不改輸入；＝伺服器 CL.materializeItems）：有 qtyChanged／unitChanged 的複製後改 qty／unit；暫時品項依序接在最後，由 makeNew(entry) 產生 */
function materializeItems(items, al, makeNew) {
  const src = Array.isArray(items) ? items : [];
  const byIndex = new Map();
  (al && al.items ? al.items : []).forEach((e) => { if (!e.isNew) byIndex.set(e.index, e); });
  const out = src.map((it, i) => {
    const e = byIndex.get(i);
    if (!e || e.conflict || e.zero || !(e.qtyChanged || e.unitChanged)) return it;
    return Object.assign({}, it, { qty: e.qty, unit: e.unit });
  });
  (al && al.items ? al.items : []).forEach((e) => { if (e.isNew) out.push(makeNew(e)); });
  return out;
}

/** 單一報價品項的金額（分）＝round-half-up(數量×單價×100)，數量／單價套用伺服器的儲存規則（同 revenueOf 的逐列取整）；標題／小計列 0 */
function itemCents(it) {
  if (!it || typeof it !== 'object' || isKindRow(it)) return 0;
  if (!HAS_BIGINT) return roundClean(storedQty(it.qty) * storedPrice(it.unitPrice) * 100);
  const q = decParts(storedQty(it.qty)), p = decParts(storedPrice(it.unitPrice));
  if (!q || !p) return 0;
  return Number(divRound(q.n * p.n * BigInt(100), pow10(q.s + p.s)));
}

// ── 委外占比 ──
/** 「委外」列＝顧問服務區（cat==='consult'）、非印花稅，且委外廠商（vendor）去頭尾空白後非空。軟體／硬體的 vendor 是「供應商」，不算委外（＝伺服器 CL.isOutsourced） */
function isOutsourced(l) {
  return !!l && typeof l === 'object' && l.auto !== 'stamp' && l.cat === 'consult' && typeof l.vendor === 'string' && l.vendor.trim() !== '';
}

/** 百分比文字 num/den×100：小數兩位、向 0 截斷（＝伺服器 CL.pctText）；num、den 要是安全整數、num>=0、den>0，否則 null */
function pctText(num, den) {
  if (!Number.isSafeInteger(num) || !Number.isSafeInteger(den) || num < 0 || den <= 0) return null;
  if (!HAS_BIGINT) {
    const h = Math.floor(num * 10000 / den);
    return Math.floor(h / 100) + '.' + (h % 100 < 10 ? '0' : '') + (h % 100);
  }
  const h = (BigInt(num) * BigInt(10000)) / BigInt(den);
  const frac = h % BigInt(100);
  return (h / BigInt(100)).toString() + '.' + (frac < BigInt(10) ? '0' : '') + frac.toString();
}

/** 毛利率文字（＝伺服器 QA.marginText）：gp／revenue 皆為「分」，兩位小數向 0 截斷；revenue<=0 或非安全整數回 null */
function marginTextOf(gpCents, revenueCents) {
  if (!Number.isSafeInteger(gpCents) || !Number.isSafeInteger(revenueCents) || revenueCents <= 0) return null;
  if (!HAS_BIGINT) {
    const a = Math.floor(Math.abs(gpCents) * 10000 / revenueCents);
    return (gpCents < 0 && a !== 0 ? '-' : '') + Math.floor(a / 100) + '.' + (a % 100 < 10 ? '0' : '') + (a % 100);
  }
  const gp = BigInt(gpCents), rev = BigInt(revenueCents);
  const neg = gp < BigInt(0);
  const hundredths = ((neg ? -gp : gp) * BigInt(10000)) / rev;
  const whole = hundredths / BigInt(100), frac = hundredths % BigInt(100);
  return (neg && hundredths !== BigInt(0) ? '-' : '') + whole.toString() + '.' + (frac < BigInt(10) ? '0' : '') + frac.toString();
}

/**
 * 成本統計（分，整數）。revenue：折扣後未稅營收（元），用來算印花稅（沒有營收才採印花稅列上伺服器給的金額，見 resolveStamp）。
 * 回傳 { byCat:{consult,software,hw,travel,other（不含印花稅）}, stampCents, totalCents（含印花稅、不含風險預留）, consultCents, outsourcedCents, stampUnknown }
 */
function costStats(lines, revenue) {
  const byCat = { consult: 0, software: 0, hw: 0, travel: 0, other: 0 };
  let hasStamp = false, serverAmt = null, outs = 0;
  (Array.isArray(lines) ? lines : []).forEach((raw) => {
    const l = cleanLine(raw);
    if (!l) return;
    if (l.auto === 'stamp') {
      if (!hasStamp && typeof l.unitCost === 'number') serverAmt = l.unitCost;
      hasStamp = true;
      return;
    }
    const c = lineCents(l);
    byCat[l.cat] += c;
    if (isOutsourced(l)) outs += c;
  });
  const amt = resolveStamp(revenue, serverAmt);
  const stampCents = hasStamp && amt !== null ? amt * 100 : 0;
  const sum = byCat.consult + byCat.software + byCat.hw + byCat.travel + byCat.other;
  return { byCat, stampCents, totalCents: sum + stampCents, consultCents: byCat.consult, outsourcedCents: outs, stampUnknown: hasStamp && amt === null };
}

/** 委外占比指標（＝伺服器 CL.outsourcedStats）：委外 ÷ 專案總成本（含差旅／交際費／印花稅，不含風險預留；總成本 0 → '0.00'）、委外 ÷ 顧問服務區成本（沒有顧問成本 → null） */
function outsourcedStats(lines, revenue) {
  const cs = costStats(lines, revenue);
  return {
    outsourcedCents: cs.outsourcedCents, totalCents: cs.totalCents, consultCents: cs.consultCents,
    outsourcedPctOfCost: cs.totalCents > 0 ? pctText(cs.outsourcedCents, cs.totalCents) : '0.00',
    outsourcedPctOfConsult: cs.consultCents > 0 ? pctText(cs.outsourcedCents, cs.consultCents) : null,
  };
}

/** 由伺服器序列化的 costBreakdown（分；consult/software/hw/travel/other(含印花稅)/outsourced）算出委外占比指標，不重算（核准面板用：數字就是伺服器的） */
function outsourcedFromBreakdown(cb) {
  if (!cb || typeof cb !== 'object') return null;
  const n = (k) => (Number.isFinite(Number(cb[k])) ? Number(cb[k]) : 0);
  const total = n('consult') + n('software') + n('hw') + n('travel') + n('other');
  const o = n('outsourced'), c = n('consult');
  return {
    outsourcedCents: o, totalCents: total, consultCents: c,
    outsourcedPctOfCost: total > 0 ? pctText(o, total) : '0.00',
    outsourcedPctOfConsult: c > 0 ? pctText(o, c) : null,
  };
}

const OS_TITLE = '委外占比＝委外成本 ÷ 專案總成本（含差旅／交際費／印花稅，不含風險預留）。委外成本＝顧問服務成本區裡有填「委外廠商」的列；自家顧問（只填顧問姓名）不算委外。';

/** 委外佔比卡片的顯示模型（純函式）：大數字、進度條寬度（0–100）、兩行小字、hover／aria 說明 */
function outsourcedCardModel(st) {
  const s = st || {};
  const pct = (typeof s.outsourcedPctOfCost === 'string' && s.outsourcedPctOfCost) ? s.outsourcedPctOfCost : '0.00';
  const pc = typeof s.outsourcedPctOfConsult === 'string' && s.outsourcedPctOfConsult ? s.outsourcedPctOfConsult : null;
  const oc = Number(s.outsourcedCents) || 0;
  const bar = Math.max(0, Math.min(100, parseFloat(pct) || 0));
  const line1 = '委外成本 NT$ ' + fmtMoney(oc / 100);
  const line2 = '占顧問服務成本 ' + (pc === null ? '—' : pc + '%');
  return {
    pct, pctLabel: pct + '%', bar, pctOfConsult: pc, outsourcedCents: oc, line1, line2,
    title: OS_TITLE,
    aria: '委外佔比 ' + pct + '%。' + line1 + '；' + line2 + '。' + OS_TITLE,
  };
}

/** 委外佔比卡片 HTML（與毛利摘要卡同樣式：pnl-sum-card）；內容由 paintOutsourcedCard 填。idAttr：選填的 id（e2e／定位用） */
function outsourcedCardHtml(model, idAttr) {
  ensureStyle();
  const m = model || outsourcedCardModel(null);
  return '<div class="pnl-sum-card qcl-os-card"' + (idAttr ? ' id="' + esc(idAttr) + '"' : '') + ' tabindex="0" role="group" title="' + esc(m.title) + '" aria-label="' + esc(m.aria) + '">' +
    '<div class="pnl-sum-label">委外佔比</div>' +
    '<div class="pnl-sum-value qcl-os-pct">' + esc(m.pctLabel) + '</div>' +
    '<div class="qcl-os-bar" role="progressbar" aria-label="委外佔比" aria-valuemin="0" aria-valuemax="100" aria-valuenow="' + esc(String(m.bar)) + '"><i class="qcl-os-fill" style="width:' + esc(String(m.bar)) + '%"></i></div>' +
    '<div class="qcl-os-l1">' + esc(m.line1) + '</div>' +
    '<div class="qcl-os-l2">' + esc(m.line2) + '</div></div>';
}

/** 把模型寫進已存在的卡片（root 內第一張 .qcl-os-card）；沒有卡片就略過 */
function paintOutsourcedCard(root, model) {
  const card = root && root.querySelector ? (root.classList && root.classList.contains('qcl-os-card') ? root : root.querySelector('.qcl-os-card')) : null;
  if (!card) return false;
  const m = model || outsourcedCardModel(null);
  const set = (sel, v) => { const el = card.querySelector(sel); if (el && el.textContent !== v) el.textContent = v; };
  set('.qcl-os-pct', m.pctLabel); set('.qcl-os-l1', m.line1); set('.qcl-os-l2', m.line2);
  const bar = card.querySelector('.qcl-os-bar'), fill = card.querySelector('.qcl-os-fill');
  if (bar) bar.setAttribute('aria-valuenow', String(m.bar));
  if (fill) fill.style.width = m.bar + '%';
  card.setAttribute('title', m.title);
  card.setAttribute('aria-label', m.aria);
  return true;
}

/**
 * 顧問對話框頂端卡片的數字（純函式）：
 *   input = { items（報價品項，含單價）, lines（成本列）, newItems（草稿新增的報價項目）, discountType, discountValue }
 *   連動後的報價（applyLinks＋materializeItems）→ 折扣後未稅營收（revenueOf）→ 成本（含印花稅）→ 毛利／毛利率／委外占比。
 *   數字與伺服器 POST /cost-draft/summary、以及按「完成」之後的實際結果一致（以分為單位逐筆相同）。
 */
function liveSummary(input) {
  const inp = input || {};
  const items = Array.isArray(inp.items) ? inp.items : [];
  const lines = Array.isArray(inp.lines) ? inp.lines : [];
  const al = applyLinks(items, lines, inp.newItems);
  const eff = materializeItems(items, al, (e) => ({ lid: e.nid, nid: e.nid, desc: e.desc, unit: e.unit, qty: e.qty, unitPrice: 0, needPrice: true, isNew: true }));
  const revenue = revenueOf(eff, inp.discountType, inp.discountValue);
  const revenueCents = Math.round(revenue * 100);
  const cs = costStats(lines, revenue);
  const gpCents = revenueCents - cs.totalCents;
  const os = outsourcedStats(lines, revenue);
  return {
    al, items: eff, revenue, revenueCents, costCents: cs.totalCents, gpCents, marginText: marginTextOf(gpCents, revenueCents),
    byCat: cs.byCat, stampCents: cs.stampCents, outsourced: os, card: outsourcedCardModel(os),
    changes: al.changes, conflicts: al.conflicts, zero: al.zero,
  };
}

/**
 * 按「完成」前的確認清單文字（純函式）：changes＝applyLinks 的 changes（qty／unit／new）。沒有異動回 ''。
 * 最多列 SYNC_LIST_MAX 項（超過接「…另 N 項」）；有新增品項時另加一句「補完單價前業務不能送簽」。
 */
const SYNC_LIST_MAX = 10;
function syncConfirmText(changes) {
  const arr = Array.isArray(changes) ? changes : [];
  if (!arr.length) return '';
  const nm = (d) => { const s = text(d, DESC_MAX); return '「' + (s.length > 20 ? s.slice(0, 20) + '…' : s) + '」'; };
  const f = (v) => fmtNum(Number(v), 4);
  const lines = arr.slice(0, SYNC_LIST_MAX).map((c) => {
    if (c.field === 'new') return '・新增報價項目' + nm(c.desc) + '　數量 ' + f(c.to) + (c.unit ? ' ' + c.unit : '') + '（單價由業務補填）';
    if (c.field === 'unit') return '・' + nm(c.desc) + '　單位 ' + (c.from === '' || c.from === null || c.from === undefined ? '（空白）' : c.from) + ' → ' + c.to;
    return '・' + nm(c.desc) + '　數量 ' + f(c.from) + ' → ' + f(c.to);
  });
  const rest = arr.length - lines.length;
  const news = arr.filter((c) => c.field === 'new').length;
  return '將同步更新業務的報價：\n' + lines.join('\n') + (rest > 0 ? '\n…另 ' + rest + ' 項' : '') +
    (news ? '\n（新增的 ' + news + ' 個品項單價是空的，業務補完單價之前不能送簽。）' : '');
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
 * 合併欄位（顧問姓名／委外廠商）：全部相同才原樣保留（含「都是空白」）；不同就把不重複的非空值以「、」串接，超過 max 字截斷（規格 §7.1）。
 * 例：['甲','甲',''] → '甲'；['甲','乙'] → '甲、乙'；['',''] → ''。
 */
function mergeField(values, max) {
  const uniq = Array.from(new Set(values));
  if (uniq.length <= 1) return uniq.length ? uniq[0] : '';
  return uniq.filter(Boolean).join('、').slice(0, max);
}

/**
 * 「全部連動、同一個品項、同單位」的合併：保持連動，數量加總（unit＝共同單位、qty＝各列數量的十進位精確加總），成本單價＝合計成本 ÷ 數量，
 * 讓成本合計（分）不變。算不出剛好的單價（數量為 0、無 BigInt、除不盡造成差一分）就回 null，由呼叫端退回「拆項」。
 */
function mergeLinked(ls, cents) {
  if (!HAS_BIGINT || ls.length < 2) return null;
  if (!ls.every((l) => l.rel === 'link' && l.forLid && !(l.forLids && l.forLids.length))) return null;
  if (!ls.every((l) => l.forLid === ls[0].forLid && l.unit === ls[0].unit)) return null;
  const qty = decimalSum(ls.map((l) => l.qty));
  const qp = decParts(qty);
  if (!(qty > 0) || !qp || qp.n === BigInt(0)) return null;
  const P = 12;   // 單價最多 12 位小數
  const x = divRound(BigInt(cents) * pow10(qp.s + P), BigInt(100) * qp.n);   // 成本單價 × 10^P（整數）
  const s = x.toString().padStart(P + 1, '0');
  const unitCost = Number(s.slice(0, s.length - P) + '.' + s.slice(s.length - P));
  if (!(unitCost >= 0) || unitCost > COST_MAX || lineCents({ qty, unitCost }) !== cents) return null;
  return { qty, unitCost };
}

/**
 * 把同一分區的多列合併成 1 列（純函式，可單獨測）：預設單位「式」、數量 1、成本單價＝各列小計（取整到分）的合計，所以成本合計不變。
 * 項目＝「合併 N 項：A、B、C」（可再改名）；顧問姓名／委外廠商依 mergeField（相同才保留，不同以「、」串接；廠商串接超過 60 字時完整名單另記在說明）；
 * 沿用第一列的 lid 與 forLid，並把所有被併列的 forLid∪forLids 放進 forLids（涵蓋判定因此把被併的每個品項都當已有成本列）。
 * 對應方式（rel，cost-sync）：所有列都沒有 rel＝維持舊行為（輸出沒有 rel）；全部「不對應」→ 不對應；全部「連動」且同一個品項、同單位 → 保持連動並把數量加總
 * （單位沿用、數量＝加總、單價＝合計 ÷ 數量，成本合計不變）；其餘（混合）→ 改為「拆項」（報價不動）。
 * 傳入的列要同一分類、不含印花稅列；空陣列回 null。
 */
function mergeLines(rows) {
  const ls = (Array.isArray(rows) ? rows : []).map(cleanLine).filter((l) => l && l.auto !== 'stamp');
  if (!ls.length) return null;
  const first = ls[0];
  const names = ls.map((l) => l.desc).filter(Boolean);
  const vendorsAll = Array.from(new Set(ls.map((l) => l.vendor).filter(Boolean)));
  const cents = ls.reduce((s, l) => s + lineCents(l), 0);
  // 顧問服務區（有「委外」判定）：不同廠商以「、」串接（合併後仍是委外，委外占比不會因合併而掉到 0）；廠商欄放得下就不另記，被截斷（超過 60 字）才把完整名單記在說明。
  // 其他分區（供應商）維持改版前的行為：只有 1 種就帶入，多種就欄位留白、記在說明
  const isConsult = first.cat === 'consult';
  const o = {
    cat: first.cat,
    desc: ls.length > 1 ? '合併 ' + ls.length + ' 項：' + names.join('、') : (names[0] || ''),
    vendor: isConsult ? mergeField(ls.map((l) => l.vendor), VENDOR_MAX) : (vendorsAll.length === 1 ? vendorsAll[0] : ''),
    note: isConsult ? (vendorsAll.join('、').length > VENDOR_MAX && vendorsAll.length > 1 ? '廠商：' + vendorsAll.join('、') : '') : (vendorsAll.length > 1 ? '廠商：' + vendorsAll.join('、') : ''),
    unit: DEFAULT_UNIT, qty: 1, unitCost: cents / 100,
  };
  if (first.cat === 'consult') { const c = mergeField(ls.map((l) => l.consultant || ''), CONSULTANT_MAX); if (c) o.consultant = c; }
  if (first.lid) o.lid = first.lid;
  if (first.forLid) o.forLid = first.forLid;
  // 被併各列所涵蓋的品項（forLid∪forLids）全部帶著：整包後「補入新品項」與完成／儲存前的提醒仍認得這些品項已有成本列（cleanLine 會剔除與 forLid 相同者、去重、上限 60）
  const all = [];
  ls.forEach((l) => lineLidList(l).forEach((k) => { if (all.indexOf(k) < 0) all.push(k); }));
  if (all.length) o.forLids = all;
  if (ls.some((l) => l.rel)) {
    const eff = ls.map((l) => l.rel || (lineLidList(l).length ? 'split' : 'none'));
    if (eff.every((r) => r === 'none') || !all.length) {
      o.rel = 'none';
    } else {
      const lk = eff.every((r) => r === 'link') ? mergeLinked(ls, cents) : null;
      if (lk) { o.rel = 'link'; o.unit = first.unit; o.qty = lk.qty; o.unitCost = lk.unitCost; delete o.forLids; }
      else {
        o.rel = 'split';
        // 第一列沒有 forLid（例如它是「不對應」）但其他列有：提升一個當 forLid，讓拆項一定有對應目標（伺服器對 link／split 沒有目標會 400）
        if (!o.forLid && all.length) { o.forLid = all[0]; if (all.length > 1) o.forLids = all.slice(1); else delete o.forLids; }
      }
    }
  }
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
  else if (act === 'relNone') relNone(inst, b.closest('tr.qcl-row'));
  else if (act === 'relNoneAll') relNone(inst, null);
}

/** 孤兒對應的一鍵處理：把該列（tr 為 null＝全部孤兒列）改成「不對應（純成本）」，清掉對應目標，成本金額不動 */
function relNone(inst, tr) {
  if (inst.dead || !inst.link) return;
  const lidSet = itemLidSet(inst.items);
  const rows = tr ? [tr] : Array.prototype.slice.call(inst.root.querySelectorAll('tr.qcl-row'));
  let n = 0;
  rows.forEach((row) => {
    const rel = row.getAttribute('data-rel');
    if (rel !== 'link' && rel !== 'split') return;
    const ids = [row.getAttribute('data-forlid') || ''];
    try { const fl = row.getAttribute('data-forlids'); if (fl) ids.push.apply(ids, JSON.parse(fl)); } catch (_) { /* 壞資料當沒有 */ }
    const live = ids.map(lidKey).filter(Boolean);
    if (!tr && lidSet && live.some((k) => lidSet.has(k))) return;   // 全部處理只動孤兒列
    row.setAttribute('data-rel', 'none');
    row.removeAttribute('data-forlid'); row.removeAttribute('data-forlids');
    const rs = row.querySelector('.qcl-rel'); if (rs) rs.value = 'none';
    const ts = row.querySelector('.qcl-target'); if (ts) { ts.hidden = true; ts.innerHTML = targetOptionsHtml(inst.entries, ''); }
    n++;
  });
  if (n) { setMsg(inst, '已把 ' + n + ' 列改成「不對應（純成本）」，成本金額不變'); emitChange(inst); }
}

function onInput(inst, ev) {
  const t = ev.target;
  if (!t || !t.classList || !t.classList.contains('qcl-in')) return;
  if (t.classList.contains('qcl-rel') || t.classList.contains('qcl-target')) return;   // 下拉由 change 事件處理（先更新列上的 data-* 再重算，避免用到舊值）
  if (t.classList.contains('qcl-desc') && t.value.trim()) t.classList.remove('qcl-bad');
  if (t.classList.contains('qcl-qty')) t.classList.remove('qcl-warn');
  setMsg(inst, '');
  emitChange(inst);   // 只更新文字，不重畫表格（輸入框維持焦點）
  // 從建議清單選取（Chrome 的 inputType 是 insertReplacementText）算「改完品名」；一般打字等 change（離開欄位）才處理，免得打到一半就被當成完整品名
  if (t.classList.contains('qcl-desc') && ev.inputType === 'insertReplacementText') applyPricebookToRow(inst, t.closest('tr.qcl-row'));
}

/** 「對應方式」下拉改了：同步列上的 data-rel／data-forlid(s)；選連動或拆項時自動預選一個目標（品名相同的品項，沒有就第一個）；選不對應時清掉目標 */
function onRelChange(inst, sel) {
  const tr = sel.closest('tr.qcl-row');
  if (!tr) return;
  const r = RELS.indexOf(sel.value) >= 0 ? sel.value : 'none';
  tr.setAttribute('data-rel', r);
  const ts = tr.querySelector('.qcl-target');
  if (r === 'none') {
    tr.removeAttribute('data-forlid'); tr.removeAttribute('data-forlids');
    if (ts) { ts.hidden = true; ts.innerHTML = targetOptionsHtml(inst.entries, ''); }
  } else {
    if (r === 'link') tr.removeAttribute('data-forlids');   // 連動只對應單一品項（合併過的多目標列選連動時只留第一個目標）
    let cur = tr.getAttribute('data-forlid') || '';
    if (!cur) {
      const d = tr.querySelector('.qcl-desc');
      const nm = nameKey(d ? d.value : '');
      const ent = (nm && inst.entries.find((e) => e.name === nm)) || inst.entries[0];
      cur = ent ? ent.key : '';
      if (cur) tr.setAttribute('data-forlid', cur);
    }
    if (ts) { ts.hidden = false; ts.innerHTML = targetOptionsHtml(inst.entries, cur); }
  }
  setMsg(inst, '');
  emitChange(inst);
}

/** 「對應的報價品項」下拉改了：同步列上的 data-forlid（連動列同時清掉多目標的 data-forlids） */
function onTargetChange(inst, sel) {
  const tr = sel.closest('tr.qcl-row');
  if (!tr) return;
  const v = lidKey(sel.value);
  if (v) tr.setAttribute('data-forlid', v); else tr.removeAttribute('data-forlid');
  if (tr.getAttribute('data-rel') === 'link') tr.removeAttribute('data-forlids');
  setMsg(inst, '');
  emitChange(inst);
}

function onChangeEv(inst, ev) {
  const t = ev.target;
  if (!t || !t.classList) return;
  if (t.classList.contains('qcl-stampchk')) emitChange(inst);
  else if (t.classList.contains('qcl-rel')) onRelChange(inst, t);
  else if (t.classList.contains('qcl-target')) onTargetChange(inst, t);
  else if (t.classList.contains('qcl-desc')) applyPricebookToRow(inst, t.closest('tr.qcl-row'));
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
    link: mode === 'edit' && isLinkOpts(opts),          // 顧問對話框：每列多「對應」欄（連動報價數量／拆項／不對應）
    names: cleanNames(opts.consultantNames),            // 「顧問姓名」欄的建議清單（datalist）
    pricebookBu: cleanBu(opts.pricebookBu),             // 報價單的 BU（可省略＝不明）：同名角色出現在多個 BU 時用來挑對的那一筆
    pricebook: cleanPricebook(opts.pricebook),          // 報價牌價簿（顧問角色人天成本）：項目欄建議清單＋改完品名自動帶入成本；沒給＝空陣列＝行為與以前相同
    linkInfo: null, entries: [], targetSig: null,
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
  /** 顧問對話框：把 applyLinks 的結果交給編輯器，在每個連動列旁顯示「→ 報價 業務原值 → 新值」、單位衝突標紅（只重畫小字，不動輸入框） */
  inst.setLinkInfo = function (al) {
    if (inst.dead || !inst.link) return;
    const m = new Map();
    const units = new Map();
    ((al && al.conflicts) || []).forEach((c) => units.set(c.lid || c.nid, c.units));
    ((al && al.items) || []).forEach((e) => { const k = e.lid || e.nid; m.set(k, Object.assign({}, e, { units: units.get(k) })); });
    inst.linkInfo = m;
    paintLinkNotes(inst);
  };
  /** 設定／更新牌價簿（載入完成後補設）：重建「項目」建議清單；不動任何已填的列 */
  inst.setPricebook = function (list, bu) {
    inst.pricebook = cleanPricebook(list);
    if (bu !== undefined) inst.pricebookBu = cleanBu(bu);
    if (inst.dead || !inst.ids) return;
    const dl = inst.el.querySelector('#' + inst.ids.desc);
    if (dl) dl.innerHTML = descSuggestFor(inst.pricebook).map((x) => '<option value="' + esc(x) + '"></option>').join('');
  };
  /** 更新「顧問姓名」建議清單 */
  inst.setNames = function (names) {
    inst.names = cleanNames(names);
    if (inst.dead || !inst.ids) return;
    const dl = inst.el.querySelector('#' + inst.ids.names);
    if (dl) dl.innerHTML = inst.names.map((x) => '<option value="' + esc(x) + '"></option>').join('');
  };
  inst.setRevenue = function (rev) {
    inst.revenue = rev;
    if (inst.dead) return;
    if (mode === 'edit') refresh(inst); else renderView(inst, inst.lines);
  };
  inst.collect = function () { return collectInst(inst); };
  /** 不標記空白項目、不動焦點地讀目前的列與檢查結果（顧問對話框送去伺服器試算前判斷「填得完不完整」）；唯讀模式回 null */
  inst.peek = function () {
    if (inst.dead || !inst.root || mode !== 'edit') return null;
    const rd = readDom(inst, false);
    return { lines: rd.lines, invalid: rd.invalid, blankDesc: rd.blankDesc, zeroQty: rd.zeroQty, badTarget: rd.badTarget };
  };
  /**
   * 逐步填寫（顧問填成本，quote-coststeps.js）用：只看某一分區（cat）的列與檢查結果，不標記空白項目、不動焦點；唯讀模式回 null。
   * 回傳 { n, cents, lines, invalid, blankDesc, zeroQty, badTarget, zeroCost, rows }：invalid／blankDesc／zeroQty／badTarget 是「列號（該區內 1 起算）」陣列，
   * zeroCost 是成本單價為 0 的列數，rows 是 [{ no, rel, ids（forLid＋forLids） }]（呼叫端用來對照單位衝突）。檢查規則與 readDom／collect 完全相同（同一個函式讀出來）。
   */
  inst.peekCat = function (cat) {
    if (inst.dead || !inst.root || mode !== 'edit' || !CAT_BY_KEY[cat]) return null;
    const rd = readDom(inst, false);
    const out = { n: 0, cents: 0, lines: [], invalid: [], blankDesc: [], zeroQty: [], badTarget: [], zeroCost: 0, rows: [] };
    rd.rows.forEach((r) => {
      const l = r.line;
      if (!l || l.auto === 'stamp' || l.cat !== cat) return;
      const no = ++out.n;
      out.lines.push(l);
      out.rows.push({ no, rel: l.rel || '', ids: lineLidList(l) });
      if (r.bad) out.invalid.push(no); else out.cents += lineCents(l);
      if (!l.desc) out.blankDesc.push(no);
      if (!r.qBad && l.qty === 0) out.zeroQty.push(no);
      if (r.badTarget) out.badTarget.push(no);
      if (!r.cBad && l.unitCost === 0) out.zeroCost++;
    });
    return out;
  };
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
  if (inst.dead || !inst.root) return { lines: [], invalid: 0, blankDesc: 0, zeroCost: 0, zeroQty: 0, badTarget: 0 };
  if (inst.mode === 'view') {
    const ls = inst.lines.filter((l) => l.auto !== 'stamp');
    return {
      lines: outLines(inst.lines), invalid: 0,
      blankDesc: ls.filter((l) => !l.desc).length, zeroCost: ls.filter((l) => l.unitCost === 0).length, zeroQty: ls.filter((l) => l.qty === 0).length, badTarget: 0,
    };
  }
  const rd = readDom(inst, true);
  refresh(inst, rd);
  return { lines: rd.lines, invalid: rd.invalid, blankDesc: rd.blankDesc, zeroCost: rd.zeroCost, zeroQty: rd.zeroQty, badTarget: rd.badTarget };
}

/** 從 DOM 讀回 lines；空白項目的列保留並標紅、數字非法標紅（qcl-bad），是否擋存由呼叫端依 invalid／blankDesc 決定 */
function collect(el) {
  const inst = findInst(el);
  return inst ? collectInst(inst) : { lines: [], invalid: 0, blankDesc: 0, zeroCost: 0, zeroQty: 0, badTarget: 0 };
}

/** 建議清單：字串、去頭尾空白、截 40 字、去空白與重複（顧問姓名 datalist 用） */
function cleanNames(list) {
  const out = [];
  (Array.isArray(list) ? list : []).forEach((x) => { const s = text(x, CONSULTANT_MAX); if (s && out.indexOf(s) < 0) out.push(s); });
  return out;
}

// ── 牌價簿 ─────────────────────────────────────────
const PB_UNIT = '人天';
const PB_BUS = ['ERP', 'ITS', 'MDM', 'CRM'];   // 與 lib/quotePricebook.js BUS 同值同序
const cleanBu = (v) => (typeof v === 'string' && PB_BUS.indexOf(v) >= 0) ? v : '';
/** 品名比對用的鍵：NFKC＋trim＋小寫（內部空白不收合，所以「PM  顧問」與「PM 顧問」不算相同——和改版前的完全相符規則一致） */
const pbKey = (v) => String(v === undefined || v === null ? '' : v).normalize('NFKC').trim().toLowerCase();
/** 清洗牌價簿清單：[{name, cost, bu}]；名稱去頭尾空白、截長度；成本要是 ≥0 的有限數字，其餘丟棄；bu 缺漏或不合法＝ERP（相容舊資料）；同一個 BU 內名稱不分大小寫去重（先到先贏），不同 BU 的同名項目都保留 */
function cleanPricebook(list) {
  const out = [], seen = Object.create(null);
  (Array.isArray(list) ? list : []).forEach((x) => {
    if (!x || typeof x !== 'object') return;
    const name = text(x.name, DESC_MAX), cost = toNum(x.cost), bu = cleanBu(x.bu) || 'ERP', k = bu + '\u0001' + pbKey(name);
    if (!name || !isFinite(cost) || cost < 0 || seen[k]) return;
    seen[k] = 1; out.push({ name, cost, bu });
  });
  return out;
}
/** 顧問「項目」欄的建議清單：牌價簿角色名在前，接著 DESC_SUGGEST；不分大小寫去重 */
function descSuggestFor(pricebook) {
  const out = [], seen = Object.create(null);
  (Array.isArray(pricebook) ? pricebook : []).concat(DESC_SUGGEST.map((n) => ({ name: n }))).forEach((p) => {
    const n = p && p.name, k = pbKey(n);
    if (n && !seen[k]) { seen[k] = 1; out.push(n); }
  });
  return out;
}
/**
 * 預設成本規則（純函式）。line＝{cat, desc, vendor, unit, unitCost}（unitCost 可以是輸入框的字串）；pricebook＝cleanPricebook 之後的清單；bu＝報價單的 BU（可省略＝不明）。
 * 條件全部成立才回傳 {unitCost, unit, name, bu}（要帶入的值），否則 null：顧問服務區、不是自動列、沒有委外廠商、成本單價空白或 0、品名（NFKC、去頭尾空白、不分大小寫）與牌價簿某項完全相同、該項成本 > 0。
 * 同名跨 BU：bu 已知 → 先看該 BU 內有沒有完全相符的；沒有（或 bu 不明）→ 全部 BU 裡「剛好一筆」同名才採用；同名在 2 個以上 BU 且無法判斷 → null（寧可不帶也不帶錯成本）。
 * unit：原本空白或是預設「式」才改「人天」，使用者自己填的單位不動；keepUnit＝true（連動列、種子列）一律維持原單位。
 */
function pricebookDefault(line, pricebook, keepUnit, bu) {
  if (!line || line.cat !== 'consult' || line.auto || !Array.isArray(pricebook) || !pricebook.length) return null;
  if (text(line.vendor, VENDOR_MAX)) return null;
  const uc = toNum(line.unitCost);
  if (isFinite(uc) && uc > 0) return null;
  if (str(line.unitCost).trim() !== '' && !isFinite(uc)) return null;   // 輸入框裡是無效文字：交給原本的驗證去擋，不蓋掉
  const d = pbKey(text(line.desc, DESC_MAX));
  if (!d) return null;
  const same = pricebook.filter((p) => pbKey(p.name) === d);
  const b = cleanBu(bu);
  let hit = b ? same.find((p) => (p.bu || 'ERP') === b) : undefined;   // BU 已知：先看該 BU 內
  if (!hit) hit = same.length === 1 ? same[0] : undefined;               // 否則只有全部 BU 裡剛好一筆才採用（0 筆＝不是牌價簿項目；2 筆以上＝有歧義）
  if (!hit || !(hit.cost > 0)) return null;
  const u = text(line.unit, UNIT_MAX);
  return { unitCost: hit.cost, unit: keepUnit ? u : ((!u || u === DEFAULT_UNIT) ? PB_UNIT : u), name: hit.name, bu: hit.bu || 'ERP' };
}
/** 使用者改完某列的品名後：符合規則就把成本（與單位）帶進該列的輸入框，並提示一句；不符合什麼都不做 */
function applyPricebookToRow(inst, tr) {
  if (!inst || inst.dead || !tr || !inst.pricebook || !inst.pricebook.length) return false;
  const tb = tr.parentNode;
  if (!tb || tb.getAttribute('data-cat') !== 'consult') return false;
  const val = (sel) => { const e = tr.querySelector(sel); return e ? e.value : ''; };
  const rel = tr.getAttribute('data-rel');
  const linked = !!rel && rel !== 'none';
  const d = pricebookDefault({ cat: 'consult', desc: val('.qcl-desc'), vendor: val('.qcl-vendor'), unit: val('.qcl-unit'), unitCost: val('.qcl-cost') }, inst.pricebook, linked, inst.pricebookBu);
  if (!d) return false;
  const cost = tr.querySelector('.qcl-cost'), unit = tr.querySelector('.qcl-unit');
  if (!cost) return false;
  cost.value = String(d.unitCost);
  if (unit && !linked && unit.value.trim() !== d.unit) unit.value = d.unit;
  emitChange(inst);
  setMsg(inst, '已依牌價簿帶入「' + d.name + '」的人天成本 ' + fmtMoney(d.unitCost) + '，可自行修改。');
  return true;
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
  // cost-sync（顧問成本畫面連動業務報價）：連動計算、委外占比、卡片數字——與伺服器 lib/quoteCostLines.js 逐例鏡像（scripts/check-quote-costsync-ui.js 以 vm 隨機比對）
  RELS, REL_LABELS, MAX_NEW_ITEMS, MIN_ITEM_QTY, CONSULTANT_MAX,
  applyLinks,              // (items, lines, newItems) → { items, changes, conflicts, zero }：連動計算（＝CL.applyLinks）
  materializeItems,        // (items, al, makeNew) → 套用連動後的報價品項陣列（＝CL.materializeItems）
  decimalSum,              // (numbers) → number：十進位精確加總
  itemCents,               // (item) → 分：單一報價品項金額（數量×單價，逐列取整）
  isOutsourced,            // (line) → boolean：顧問服務區且委外廠商非空
  pctText,                 // (num, den) → '12.34'|null：小數兩位向 0 截斷（＝CL.pctText）
  marginTextOf,            // (gpCents, revenueCents) → '31.25'|null（＝QA.marginText）
  costStats,               // (lines, revenue) → { byCat, stampCents, totalCents, consultCents, outsourcedCents, stampUnknown }（分）
  outsourcedStats,         // (lines, revenue) → { outsourcedCents, totalCents, consultCents, outsourcedPctOfCost, outsourcedPctOfConsult }（＝CL.outsourcedStats）
  outsourcedFromBreakdown, // (costBreakdown) → 同上；直接用伺服器序列化的分類彙總（核准面板用）
  outsourcedCardModel,     // (stats) → { pct, pctLabel, bar, line1, line2, title, aria }：委外佔比卡片顯示模型
  outsourcedCardHtml,      // (model[, id]) → 卡片 HTML（pnl-sum-card 樣式）
  paintOutsourcedCard,     // (root, model) → 更新卡片文字與進度條
  liveSummary,             // ({items, lines, newItems, discountType, discountValue}) → 卡片數字（營收／成本／毛利／毛利率／委外占比）與連動結果
  syncConfirmText,         // (changes) → 按「完成」前的「將同步更新業務的報價」清單文字
  targetEntries,           // (items) → [{key, seq, name, isNew}]：「對應」下拉的選項來源
  lineCents,               // (line) → 分：單列成本（數量×單價，十進位精確取整，＝伺服器 centsOf）
  // 牌價簿（純函式；見檔頭說明）
  pbClean: cleanPricebook,
  pbDescSuggest: descSuggestFor,
  pbDefault: pricebookDefault,
  // 內部（單元測試用）
  _unmatchedMessage: unmatchedMessage,
  _isOrphanLine: (l, items) => isOrphanLine(l, itemLidSet(items)),
  _missingItemLines: missingItemLines,
  _supplementMessage: supplementMessage,
  _mergeLines: mergeLines,
  _catByUnit: catByUnit,
  _editRowHtml: (l, cat, ctx) => editRowHtml(cleanLine(Object.assign({}, l, { cat })), CAT_BY_KEY[cat], { desc: 'd', unit: 'u', names: 'n' }, ctx),
  _mergeConfirmText: (cat, rows) => mergeConfirmText(CAT_BY_KEY[cat], rows, rows.reduce((t, l) => t + lineCents(l), 0) / 100),
  _lineCents: lineCents,
  _viewRowHtml: (l, cat, items) => viewRowHtml(cleanLine(Object.assign({}, l, { cat })), CAT_BY_KEY[cat], itemLidSet(items)),
};
})(typeof window !== 'undefined' ? window : globalThis);
