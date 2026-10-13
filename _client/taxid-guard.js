/*!
 * 統編 ↔ 公司名稱一致性規則 —— 前後端單一來源（無相依、不碰 DOM、不打網路）
 *   瀏覽器：<script src="taxid-guard.js"> → window.TaxIdGuard
 *   Node  ：require('./_client/taxid-guard.js')（server.js 的 POST/PUT /api/contacts 用它；Vercel includeFiles 已含 _client/**）
 *
 * 規則（保守設計：寧可多請使用者按一下「使用官方名稱」，也不放行錯誤配對）
 *   1. 公司名稱「純數字」一律視為錯誤（COMPANY_IS_TAXID）——常見成因是把統編打進公司欄。
 *   2. 統編對得上官方公司名（內部主檔 / GCIS）時，輸入的公司名稱正規化後必須相同，否則 TAXID_COMPANY_MISMATCH。
 *   3. 正規化只做「不改變實體身分」的轉換：全形→半形(NFKC)、大小寫、所有空白、台/臺、
 *      (股)/（股）≡ 股份有限公司、(有)/（有）≡ 有限公司、去除 . , 。 ， 、。
 *      不會把「有限公司」與「股份有限公司」視為相同，也不會把「省略後綴」視為相同（那是不同的字串、可能是不同實體）。
 *   4. 查不到 / 查詢失敗 → 無法驗證 → 不擋（呼叫端只顯示警告）。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TaxIdGuard = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // 正規化公司名稱（僅供「是否同一個名稱」比較，不用於儲存）
  function normalizeCompanyName(s) {
    let v = String(s == null ? '' : s);
    try { v = v.normalize('NFKC'); } catch (e) { /* 舊瀏覽器：略過 */ }
    v = v.toLowerCase()
      .replace(/\s+/g, '')                 // 所有空白（含全形空白，NFKC 後已轉半形）
      .replace(/臺/g, '台')
      .replace(/\(股\)/g, '股份有限公司')   // （股）NFKC 後已是 (股)
      .replace(/\(有\)/g, '有限公司')
      .replace(/[.,。，、]/g, '');
    return v;
  }

  // 公司欄只有數字（含全形數字、空白）→ true
  function isDigitsOnly(s) {
    const v = normalizeCompanyName(s);
    return v.length > 0 && /^\d+$/.test(v);
  }

  // 統編檢查碼（財政部規則：權重 1,2,1,2,1,2,4,1；乘積十位+個位相加；和能被 5 整除，
  // 第 7 碼為 7 時和或和+1 能被 5 整除。2023 起由 10 改 5，舊制有效者在新制仍有效）
  function isValidTaxIdChecksum(tid) {
    const t = String(tid == null ? '' : tid).trim();
    if (!/^\d{8}$/.test(t)) return false;
    const w = [1, 2, 1, 2, 1, 2, 4, 1];
    let sum = 0;
    for (let i = 0; i < 8; i++) {
      const p = Number(t[i]) * w[i];
      sum += Math.floor(p / 10) + (p % 10);
    }
    if (sum % 5 === 0) return true;
    return t[6] === '7' && (sum + 1) % 5 === 0;
  }

  // 官方名稱是否「可作為比對基準」：非空、非純數字（被污染的主檔名稱 = 統編 時不可信）
  function isUsableOfficialName(name) {
    const n = String(name == null ? '' : name).trim();
    return !!n && !isDigitsOnly(n);
  }

  function sameCompany(a, b) {
    const x = normalizeCompanyName(a), y = normalizeCompanyName(b);
    return !!x && x === y;
  }

  // 組錯誤訊息（前後端共用同一句，便於對照）
  function mismatchMessage(taxId, officialName) {
    return '統編 ' + taxId + ' 對應公司為「' + officialName + '」，與輸入的公司名稱不符，請修正';
  }
  function digitsMessage() {
    return '公司名稱不可只有數字（看起來像是把統編填到公司欄了），請輸入公司名稱';
  }

  /**
   * 核心判斷（純函式）。
   * @param {object} p { company, taxId, isForeign, officialName }
   *   officialName：已查到的官方名稱（內部主檔或 GCIS）；未知/查不到傳 ''/null → 不做不符檢查
   * @returns {null | {code, error}}  null＝通過
   */
  function evaluate(p) {
    p = p || {};
    const company = String(p.company == null ? '' : p.company).trim();
    const taxId = String(p.taxId == null ? '' : p.taxId).trim();
    if (!company) return null;                       // 空值由既有「必填」規則處理
    if (isDigitsOnly(company)) {
      // 純數字：與統編相同、或長得像統編（8 碼）一律擋；短數字（例如 3 碼）後端不擋、前端見 strictDigits
      if (p.strictDigits || normalizeCompanyName(company) === taxId || /^\d{8}$/.test(normalizeCompanyName(company))) {
        return { code: 'COMPANY_IS_TAXID', error: digitsMessage() };
      }
    }
    if (p.isForeign || !/^\d{8}$/.test(taxId)) return null;
    if (!isUsableOfficialName(p.officialName)) return null;
    if (!sameCompany(company, p.officialName)) {
      return { code: 'TAXID_COMPANY_MISMATCH', error: mismatchMessage(taxId, String(p.officialName).trim()) };
    }
    return null;
  }

  // 內部企業主檔查統編（只讀 companies 陣列，不打外部）。回傳可用的官方名稱或 ''
  function masterOfficialName(companies, taxId) {
    const t = String(taxId == null ? '' : taxId).trim();
    if (!/^\d{8}$/.test(t) || !Array.isArray(companies)) return '';
    const m = companies.find(function (c) { return c && (c.matchKey === 'tax:' + t || c.taxId === t); });
    if (!m) return '';
    return isUsableOfficialName(m.name) ? String(m.name).trim() : '';
  }

  /**
   * 後端儲存路徑用：只比對內部主檔，不打外部。
   * @param {object} p { company, taxId, isForeign, companies, oldCompany, oldTaxId, isUpdate }
   *   isUpdate=true 時，company 與 taxId 都沒變動就不驗（既有舊資料的無關欄位編輯不受影響）
   */
  function checkSave(p) {
    p = p || {};
    const company = String(p.company == null ? '' : p.company).trim();
    const taxId = p.isForeign ? '' : String(p.taxId == null ? '' : p.taxId).trim();
    if (p.isUpdate) {
      const unchanged = company === String(p.oldCompany == null ? '' : p.oldCompany).trim() &&
                        taxId === String(p.oldTaxId == null ? '' : p.oldTaxId).trim();
      if (unchanged) return null;
    }
    return evaluate({
      company: company, taxId: taxId, isForeign: !!p.isForeign,
      officialName: masterOfficialName(p.companies, taxId)
    });
  }

  return {
    normalizeCompanyName: normalizeCompanyName,
    isDigitsOnly: isDigitsOnly,
    isValidTaxIdChecksum: isValidTaxIdChecksum,
    isUsableOfficialName: isUsableOfficialName,
    sameCompany: sameCompany,
    masterOfficialName: masterOfficialName,
    mismatchMessage: mismatchMessage,
    evaluate: evaluate,
    checkSave: checkSave
  };
});
