/*!
 * SAP Deal ID 格式規則 —— 前後端單一來源（無相依、不碰 DOM、不打網路）
 *   瀏覽器：<script src="deal-id-format.js"> → window.DealIdFormat
 *   Node  ：require('./_client/deal-id-format.js')（lib/productCatalog.js 的 normalizeDealId 用它）
 *
 * 規則：SAP Deal ID 必須是 9 碼半形數字（^[0-9]{9}$），以字串儲存、開頭 0 保留。
 *   輸入先做 NFKC（全形數字→半形），去掉空白與分隔符（空格、-、,）後再驗；儲存的是「正規化後的 9 碼」。
 *   空字串 / null / undefined 視為「沒填」（是否必填由商品目錄規則決定，不在這裡）。
 * 注意：識別字 dealId / requiresDealId / 錯誤碼 DEAL_ID_* 皆為相容性保留，不改名。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DealIdFormat = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  var LEN = 9;
  var MESSAGE = 'SAP Deal ID 必須是 ' + LEN + ' 碼數字';
  var RE = /^[0-9]{9}$/;
  var SEP_RE = /[\s\-,]/g;                                // 允許被忽略的分隔符
  var BAD_CTRL_RE = new RegExp('[' + '\\u0000-\\u0008\\u000b\\u000c\\u000e-\\u001f\\u007f-\\u009f' + String.fromCharCode(8232, 8233) + ']');   // non-whitespace control chars incl. U+2028/2029 => invalid

  /** 輸入即時過濾（給輸入框用）：NFKC → 只留 0-9 → 截到 9 碼。不做驗證。 */
  function sanitizeInput(v) {
    var s = String(v == null ? '' : v);
    try { s = s.normalize('NFKC'); } catch (e) { /* 舊瀏覽器 */ }
    return s.replace(/[^0-9]/g, '').slice(0, LEN);
  }

  /**
   * 正規化＋驗證。回 { ok:true, value } 或 { ok:false, code, error }。
   *   null/undefined → { ok:true, value:'' }；number 先轉字串；其他型別 → DEAL_ID_INVALID；格式不符 → DEAL_ID_FORMAT。
   */
  function normalize(v) {
    if (v === null || v === undefined) return { ok: true, value: '' };
    if (typeof v !== 'string' && typeof v !== 'number') return { ok: false, code: 'DEAL_ID_INVALID', error: 'SAP Deal ID 格式不正確（需為 ' + LEN + ' 碼數字）' };
    var s = String(v);
    if (s.length > 200) return { ok: false, code: 'DEAL_ID_FORMAT', error: MESSAGE };   // 超長直接擋，避免對巨大字串做正規化
    try { s = s.normalize('NFKC'); } catch (e) { /* ignore */ }
    if (BAD_CTRL_RE.test(s)) return { ok: false, code: 'DEAL_ID_FORMAT', error: MESSAGE };
    s = s.replace(SEP_RE, '');
    if (s === '') return { ok: true, value: '' };
    if (!RE.test(s)) return { ok: false, code: 'DEAL_ID_FORMAT', error: MESSAGE };
    return { ok: true, value: s };
  }

  return { LEN: LEN, MESSAGE: MESSAGE, RE: RE, sanitizeInput: sanitizeInput, normalize: normalize };
});
