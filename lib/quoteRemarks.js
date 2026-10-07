'use strict';
/**
 * 報價單 Remarks（備註條款）的單一來源 —— Excel（lib/quoteExcel.js）、網頁預覽（_client/quote-preview.js，經 issue-info 取得）、
 * 核准雜湊（lib/quoteApproval.js）、表單驗證（lib/quoteRoutes.js）都從這裡取。
 *
 * 條款結構（Steven 2026-10-07 指示）：
 *   第 1、2、5、6 條：固定條文，不可修改（FIXED）。
 *   第 4 條：條文固定，只有日期可改（＝「報價期限」欄位）。
 *   第 3 條「付款方式」：活的。結構化「付款期程」payment＝{ net, items:[{label, pct}] }，由系統組句，不讓業務手打。
 *   第 7 條起：議價時追加的條款 extraClauses（字串陣列，一條一列，自動編號）。舊單的單行「備註」note 視為第 7 條。
 *
 * 付款期程與追加條款是「印在客戶報價單上的商業承諾」，所以納入核准內容雜湊（核准後修改要重簽）；
 * 但舊單沒有這兩個欄位時雜湊不變（只有「有值」才放進 payload）。
 * 前端 _client/quote.js 的 quotePaymentSentence() 是這裡 paymentSentence() 的鏡像（表單即時預覽用），改一邊要同步；`node scripts/check-quote-remarks.js` 會逐例比對兩邊輸出（以及範本固定條文與 FIXED 是否一致）。
 */

const FIXED = Object.freeze({
  1: '1.以上報價為東捷所提供之特惠價，基於誠信原則，本報價單相關條款合約及價格雙方均不得向第三者揭露。',
  2: '2.ABAP客製開發人天，每人天12,000元（未稅）計價',
  4: '4.本報價單於XXXX年XX月XX日前有效。',      // 日期由「報價期限」帶入（年月日不補零，例：2026年10月30日）
  5: '5.本報價單經簽名回傳視同為正式有效之合約。',
  6: '6.本報價單需加蓋東捷資訊服務(股)公司之"報價專用章"使為有效之正式報價單。',
});

/** 付款期限：0＝即付（不月結）；其餘為「月結 N 天」 */
const NET_OPTIONS = Object.freeze([0, 30, 45, 60, 90]);
const MAX_INSTALLMENTS = 8;      // 付款期數上限
const MAX_LABEL = 30;            // 每期「付款時點」文字長度上限
const MAX_EXTRA = 8;             // 追加條款條數上限（範本預建 8 列，見 LAYOUT.extraFirst/extraLast）
const MAX_CLAUSE = 200;          // 每條追加條款長度上限

/** 舊單（沒有 payment 欄位）一律印這句：簽約完成後付款總金額 100%，月結30天付款 */
const DEFAULT_PAYMENT = Object.freeze({ net: 30, items: Object.freeze([Object.freeze({ label: '簽約完成後', pct: 100 })]) });

/** 表單的付款範本（業務選一個再微調）。key 只用於表單，不存進報價單（存的是實際期程） */
const PAYMENT_PRESETS = Object.freeze([
  { key: 'once30', name: '簽約後一次付款（月結30天）', net: 30, items: [{ label: '簽約完成後', pct: 100 }] },
  { key: 'prepay', name: '簽約後即付款（硬體／小專案）', net: 0, items: [{ label: '簽約完成後', pct: 100 }] },
  { key: 'three', name: '分 3 期（小型專案）', net: 30, items: [{ label: '簽約後', pct: 30 }, { label: '系統開發完成後', pct: 40 }, { label: '上線驗收後', pct: 30 }] },
  { key: 'five', name: '分 5 期（建置專案）', net: 30, items: [{ label: '簽約後', pct: 20 }, { label: '藍圖確認完成後', pct: 20 }, { label: '系統開發完成後', pct: 20 }, { label: 'UAT 測試完成後', pct: 20 }, { label: '上線驗收後', pct: 20 }] },
]);

// 零寬字元（U+200B/200C/200D/2060/FEFF）看不見卻會讓「空白」條款通過驗證；用 fromCharCode 組成，避免原始碼夾帶不可見字元
const ZERO_WIDTH = new RegExp('[' + [0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF].map((c) => String.fromCharCode(c)).join('') + ']', 'g');
/** 只收十進位數字（不收 0x1E、1e2 這類 Number() 會吃的寫法）：數字原樣、字串必須是純十進位，其餘 NaN */
const parseNum = (v) => (typeof v === 'number' ? v : (typeof v === 'string' && /^\d+(\.\d+)?$/.test(v.trim()) ? Number(v) : NaN));
const oneLine = (s) => String(s == null ? '' : s).replace(ZERO_WIDTH, '').replace(/[\u0000-\u001F\u007F\u2028\u2029]+/g, ' ').replace(/\s+/g, ' ').trim();

/** 驗證並正規化付款期程。回傳 { value } 或 { error }。不做靜默轉換：型別不對、比例不是 100% 一律拒絕 */
function normalizePayment(raw) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return { error: '付款方式格式不正確' };
  const net = parseNum(raw.net);
  if (!NET_OPTIONS.includes(net)) return { error: `付款期限必須是 ${NET_OPTIONS.map(n => n === 0 ? '即付' : '月結' + n + '天').join('、')} 其中之一` };
  if (!Array.isArray(raw.items) || raw.items.length < 1) return { error: '付款方式至少要有一期' };
  if (raw.items.length > MAX_INSTALLMENTS) return { error: `付款期數最多 ${MAX_INSTALLMENTS} 期` };
  const items = [];
  let tenths = 0;
  for (let i = 0; i < raw.items.length; i++) {
    const it = raw.items[i];
    if (!it || typeof it !== 'object' || Array.isArray(it)) return { error: `第 ${i + 1} 期格式不正確` };
    if (typeof it.label !== 'string') return { error: `第 ${i + 1} 期的付款時點必須是文字` };
    const label = oneLine(it.label);
    if (!label) return { error: `第 ${i + 1} 期請填付款時點（例：簽約後、上線驗收後）` };
    if ([...label].length > MAX_LABEL) return { error: `第 ${i + 1} 期的付款時點最多 ${MAX_LABEL} 個字` };
    const pct = parseNum(it.pct);
    if (!Number.isFinite(pct) || pct < 0.1 || pct > 100) return { error: `第 ${i + 1} 期的比例必須介於 0.1% 到 100%` };
    const t = Math.round(pct * 10);
    if (Math.abs(pct * 10 - t) > 1e-9) return { error: `第 ${i + 1} 期的比例最多只能到小數 1 位` };
    tenths += t;
    items.push({ label, pct: t / 10 });
  }
  if (tenths !== 1000) return { error: `各期比例合計必須是 100%（目前 ${tenths / 10}%）` };
  return { value: { net, items } };
}

/**
 * 驗證並正規化追加條款（字串陣列）。空字串會被濾掉；超長／超過條數不截斷，直接拒絕。
 * opts.lenientLength：不檢查單條長度——只給「判斷表單送回來的是不是原封不動的舊單行備註」用
 * （舊欄位 note 上限 500 字，比追加條款的 200 字長；原封不動就不能因為長度被擋，否則整張單連改價格都存不了）。
 */
function normalizeClauses(raw, opts) {
  const lenient = !!(opts && opts.lenientLength);
  if (!Array.isArray(raw)) return { error: '追加條款格式不正確' };
  const out = [];
  for (let i = 0; i < raw.length; i++) {
    if (typeof raw[i] !== 'string') return { error: `第 ${i + 1} 條追加條款必須是文字` };
    const s = oneLine(raw[i]);
    if (!s) continue;
    if (!lenient && [...s].length > MAX_CLAUSE) return { error: `追加條款每條最多 ${MAX_CLAUSE} 個字（第 ${out.length + 7} 條超過）` };
    out.push(s);
  }
  if (out.length > MAX_EXTRA) return { error: `追加條款最多 ${MAX_EXTRA} 條` };
  return { value: out };
}

const samePayment = (a, b) => !!a && !!b && a.net === b.net && a.items.length === b.items.length && a.items.every((x, i) => x.label === b.items[i].label && x.pct === b.items[i].pct);
/** 是否等同「舊單預設句」（用來判斷表單送回來的值有沒有實質變更） */
const isDefaultPayment = (p) => samePayment(p, DEFAULT_PAYMENT);

const fmtPct = (n) => (Number.isInteger(n) ? String(n) : String(Math.round(n * 10) / 10));

/** 第 3 條「付款方式」整句（含編號）。payment 缺漏＝舊單預設句 */
function paymentSentence(payment) {
  // 防呆：資料若被人工改壞（items 含 null／缺欄位）就退回預設句，不讓匯出整個失敗（API 寫入時一律已正規化，正常不會發生）
  const okItems = payment && Array.isArray(payment.items) && payment.items.length && payment.items.every((it) => it && typeof it.label === 'string' && Number.isFinite(it.pct));
  const p = okItems ? payment : DEFAULT_PAYMENT;
  const net = typeof p.net === 'number' ? p.net : 30;
  if (p.items.length === 1 && p.items[0].pct === 100) {
    const label = p.items[0].label;
    return net > 0 ? `3.付款方式：${label}付款總金額 100%，月結${net}天付款。` : `3.付款方式：${label}即付款總金額 100%。`;
  }
  const parts = p.items.map((it, i) => `第${i + 1}期 ${it.label} ${fmtPct(it.pct)}%`);
  const tail = net > 0 ? `各期月結${net}天付款。` : '各期於付款時點成立後即付款。';
  return `3.付款方式：分${p.items.length}期付款（比例為占付款總金額），${parts.join('；')}；${tail}`;
}

/** 追加條款（第 7 條起）。有 extraClauses 就用它；沒有（舊單）就把單行 note 當第 7 條 */
function effectiveClauses(q) {
  if (q && Array.isArray(q.extraClauses)) return q.extraClauses.filter(s => typeof s === 'string' && s.trim()).map(oneLine);
  const note = oneLine(q && q.note);
  return note ? [note] : [];
}

/** 第 4 條：把「XXXX年XX月XX日」換成報價期限；日期不合法時維持範本佔位（呼叫端通常已保證合法） */
function validSentence(validUntilIso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(validUntilIso || ''));
  return m ? FIXED[4].replace(/X{4}年X{2}月X{2}日/, `${+m[1]}年${+m[2]}月${+m[3]}日`) : FIXED[4];
}

/** 完整 Remarks 條款（不含「Remarks ：」標題）：第 1~6 條 ＋ 追加條款（第 7 條起，自動編號） */
function composeRemarks(q, validUntilIso) {
  const extras = effectiveClauses(q).map((t, i) => `${7 + i}.${t}`);
  return [FIXED[1], FIXED[2], paymentSentence(q && q.payment), validSentence(validUntilIso), FIXED[5], FIXED[6], ...extras];
}

module.exports = {
  FIXED, NET_OPTIONS, MAX_INSTALLMENTS, MAX_LABEL, MAX_EXTRA, MAX_CLAUSE, DEFAULT_PAYMENT, PAYMENT_PRESETS,
  normalizePayment, normalizeClauses, isDefaultPayment, samePayment, paymentSentence, effectiveClauses, validSentence, composeRemarks, oneLine,
};
