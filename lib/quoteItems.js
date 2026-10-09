/**
 * 報價單「品項列」的種類與小計——純函式、零相依（可被 quoteRoutes／quoteApproval／quoteExcel／quotePnlExcel 共用；
 * _client/quote-preview.js 的 _qpvSubtotals 是它的鏡像，scripts/check-quote-items.js 會逐例比對；
 * 表單的 updateQuoteSubtotalRows（DOM 版）與核准面板／董事會簽呈的逐段累加是同一規則的另外兩份，沒有自動比對，改規則時要三處一起看）。
 *
 * items 陣列裡每一列有三種 kind：
 *   'item'     一般品項（舊單沒有 kind 欄位＝item，行為與以前完全一樣）：有說明／單位／數量／單價／成本
 *   'title'    分組標題（例：Part A：系統建置）：只有 desc，不計價、不計成本
 *   'subtotal' 小計列：desc 是標籤（空白時顯示「小計」）；金額由系統算，不存、不計入總價
 * 小計金額＝「上一個標題或小計列之後」到這一列之前，所有品項（數量×單價）的合計。
 * 標題／小計列沒有數量、單價、成本、毛利分類，也不參與簽核毛利、成本填寫、毛利分析。
 * 列數上限（50）含標題與小計列。
 */
'use strict';

const NON_ITEM_KINDS = Object.freeze(['title', 'subtotal']);
const DEFAULT_SUBTOTAL_LABEL = '小計';

/** 未知或缺漏的 kind 一律當一般品項（舊單、舊前端送來的資料都沒有 kind） */
function normalizeRowKind(k) { return (k === 'title' || k === 'subtotal') ? k : 'item'; }

/** 是否為標題／小計列（只有「是物件且 kind 為 title/subtotal」才算；null、非物件交給各處原本的 BAD_ITEM 檢查） */
function isNonItemRow(it) { return !!it && typeof it === 'object' && NON_ITEM_KINDS.includes(it.kind); }
function isItemRow(it) { return !isNonItemRow(it); }
function itemRows(items) { return (Array.isArray(items) ? items : []).filter(isItemRow); }

/** 顯示用單列金額（浮點，不進位）：數量空／0→1，單價空→0；與 quoteExcel／quote-preview 一致 */
function lineAmount(it) {
  const q = parseFloat(it && it.qty) || 1;
  const p = parseFloat(it && it.unitPrice) || 0;
  return q * p;
}

/**
 * 每一列對應的小計值（與 items 等長）：小計列＝該段品項合計（浮點、不進位），其餘列＝null。
 * 段落從「最近的標題或小計列之後」開始。
 */
function subtotalValues(items) {
  const arr = Array.isArray(items) ? items : [];
  const out = new Array(arr.length).fill(null);
  let acc = 0;
  for (let i = 0; i < arr.length; i++) {
    const it = arr[i];
    if (isNonItemRow(it)) {
      if (it.kind === 'subtotal') out[i] = acc;
      acc = 0;
    } else {
      acc += lineAmount(it);
    }
  }
  return out;
}

/** 小計列的顯示標籤（空白時用預設「小計」） */
function subtotalLabel(it) {
  const s = String((it && it.desc) || '').replace(/\s+/g, ' ').trim();
  return s || DEFAULT_SUBTOTAL_LABEL;
}

// ── 品項的「說明」與「備註」（兩個選填的單行文字，會印在客戶報價單上）─────────────────────────
//   spec（說明）：品項的補充說明，≤200 字；從牌價簿選取時複製牌價簿的說明（之後與牌價簿脫鉤），可在報價單上逐單修改。
//   note（備註）：業務為這個品項寫的特殊需求備註，≤300 字。
// 只有一般品項有；分組標題／小計列沒有。**不納入核准雜湊（contentHash）也不納入 lineStructureSig／itemsSig／costLinesSig**：
// 改文字不會讓核准作廢、也不會讓顧問的成本畫面過期（錢相關欄位的改動照舊要重簽）。沒有填的品項不多存任何欄位（舊單位元級不變）。
const SPEC_MAX = 200;
const NOTE_MAX = 300;

/**
 * 清洗成「單行純文字」：非字串（含 null／數字／物件）→ ''；控制字元（含換行、tab、U+2028/2029）一律當空白，連續空白收成一個，前後去空白，
 * 再依字數（以 Unicode 字元計，不會把代理對切半）截到 max。與 desc 的差別：desc 只 trim，這裡要求單行（客戶單上一個品項只占固定的行）。
 */
function cleanItemText(v, max) {
  if (typeof v !== 'string') return '';
  const s = v.replace(/[\u0000-\u001F\u007F-\u009F\u2028\u2029]+/g, ' ').replace(/\s+/g, ' ').trim();
  const lim = max > 0 ? max : SPEC_MAX;
  const cps = Array.from(s);
  return cps.length > lim ? cps.slice(0, lim).join('').trim() : s;
}
const cleanSpec = (v) => cleanItemText(v, SPEC_MAX);
const cleanNote = (v) => cleanItemText(v, NOTE_MAX);

/** 一般品項的說明／備註（清洗後；標題、小計列、非物件一律回 ''）。輸出端（預覽、Excel、序列化）共用，避免各處自己判斷 */
function itemSpec(it) { return (it && typeof it === 'object' && !isNonItemRow(it)) ? cleanSpec(it.spec) : ''; }
function itemNote(it) { return (it && typeof it === 'object' && !isNonItemRow(it)) ? cleanNote(it.note) : ''; }

module.exports = { NON_ITEM_KINDS, DEFAULT_SUBTOTAL_LABEL, SPEC_MAX, NOTE_MAX, normalizeRowKind, isNonItemRow, isItemRow, itemRows, lineAmount, subtotalValues, subtotalLabel, cleanItemText, cleanSpec, cleanNote, itemSpec, itemNote };
