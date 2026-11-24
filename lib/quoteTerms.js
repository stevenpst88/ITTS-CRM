'use strict';
/**
 * 報價單品項「服務期間」（License／MA 起訖日）— 伺服器端。純函式規則（驗證、月數、快捷、格式）的單一來源在 _client/quote-terms.js
 * （前端同一份檔，不另做鏡像，所以不會漂移；Vercel includeFiles 已含 _client/**）。本檔再加：
 *   · 功能開關 MA_TRACKING=1（預設關閉）：關閉時請求裡的 termKind／termFrom／termTo 一律忽略、回應不輸出、不驗證、沒有 item-terms 端點；
 *     已存在的欄位（功能開啟期間存下的）關閉後原樣保留、不輸出（不會因為關閉再存檔而被洗掉）。
 *   · 品項欄位解析（resolveRowTerm）：沒送欄位沿用舊值（舊分頁不會洗掉期間）。
 *   · 整批請求驗證（validateRawItems）：license／ma 起迄日必填、真實日期、迄日不早於起日，錯誤訊息指出「第幾列」（只數一般品項）。
 * 欄位 termKind／termFrom／termTo 不進 lib/quoteApproval.js 的 contentHash 與任何簽章（那些函式只取名單內的欄位，本功能完全不碰）。
 * 檢查：scripts/check-quote-terms.js
 */
const QT = require('../_client/quote-terms.js');
const QI = require('./quoteItems');

/** 功能開關：MA_TRACKING=1 才啟用（讀環境變數的當下判斷，測試可切換） */
function enabled(env) {
  const e = env || process.env;
  return String(e.MA_TRACKING == null ? '' : e.MA_TRACKING).trim() === '1';
}

const clean = (v) => (typeof v === 'string' ? v.trim() : '');

/**
 * 一列的最終期間欄位。r＝請求裡的列、old＝同 lid 的既有列（可為 null）、on＝功能是否啟用。
 * 回 { kind:'license'|'ma'|'', from, to, badKind:boolean }。
 *   · 功能關閉：完全忽略 r，沿用 old 已存的欄位（沒有就都是空）。
 *   · 欄位沒送（undefined）→ 沿用舊值（舊版畫面、舊分頁不認得這些欄位，不能洗掉）；送了就以送的為準（null／''＝清除；kind 送 ''／'none'＝改回一般）。
 *   · 類型改回一般 → 起迄日一併清掉；類型不認得 → badKind（呼叫端回 TERM_INVALID）。
 */
function resolveRowTerm(r, old, on) {
  const oldKind = old ? (QT.parseKind(old.termKind) || '') : '';
  const oldFrom = old ? clean(old.termFrom) : '';
  const oldTo = old ? clean(old.termTo) : '';
  if (!on) return { kind: oldKind, from: oldKind ? oldFrom : '', to: oldKind ? oldTo : '', badKind: false };
  const has = (k) => r && typeof r === 'object' && r[k] !== undefined;
  let kind = oldKind, badKind = false;
  if (has('termKind')) {
    const k = QT.parseKind(r.termKind);
    if (k === null) { badKind = true; kind = ''; } else kind = k;
  }
  if (!kind) return { kind: '', from: '', to: '', badKind };
  return {
    kind,
    from: has('termFrom') ? clean(r.termFrom) : oldFrom,
    to: has('termTo') ? clean(r.termTo) : oldTo,
    badKind,
  };
}

/** 依 lid 把請求的每一列對回既有列（與 quoteRoutes.normalizeItems 同一套規則：lid 重複的後面視為新列）。回 [old|null, ...]，長度＝raw 長度（超過 max 的不處理） */
function pairWithOld(raw, existing, max) {
  const byLid = new Map();
  (existing || []).forEach((it, i) => byLid.set(it && it.lid ? it.lid : ('legacy-' + i), it));
  const used = new Set();
  return (Array.isArray(raw) ? raw : []).slice(0, max || Infinity).map((r) => {
    let old = r && r.lid ? byLid.get(r.lid) : null;
    if (old && used.has(old.lid)) old = null;
    if (old && old.lid) used.add(old.lid);
    return old || null;
  });
}

const nameOf = (desc) => Array.from(String(desc == null ? '' : desc).replace(/[\u0000-\u001f\u007f-\u009f]/g, ' ').trim()).slice(0, 12).join('');
const rowMsg = (n, desc, reason) => `第 ${n} 列${nameOf(desc) ? '「' + nameOf(desc) + '」' : ''}${reason}`;

/**
 * 整批驗證（POST／PUT 用，在 normalizeItems 之前）。回 null＝通過；否則 { status:400, code, message }。
 * 功能關閉時一律回 null。標題／小計列略過；列號只數一般品項（與畫面「第 N 列」一致）。
 */
function validateRawItems(raw, existing, on, max) {
  if (!on) return null;
  const olds = pairWithOld(raw, existing, max);
  let n = 0;
  const rows = (Array.isArray(raw) ? raw : []).slice(0, max || Infinity);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (QI.normalizeRowKind(r && r.kind) !== 'item') continue;
    n++;
    const t = resolveRowTerm(r, olds[i], true);
    if (t.badKind) return { status: 400, code: 'TERM_INVALID', message: rowMsg(n, r && r.desc, '的類型不正確（只能是一般、License 或 MA）') };
    const bad = QT.validateTerm(t.kind, t.from, t.to);
    if (bad) return { status: 400, code: bad.code, message: rowMsg(n, r && r.desc, bad.code === 'TERM_REQUIRED' ? bad.reason : '：' + bad.reason) };
  }
  return null;
}

/** 已儲存的品項是否為「有期間的 License／MA 品項」 */
function hasTerm(it) { return !!it && QI.isItemRow(it) && !!QT.parseKind(it.termKind); }

/** 序列化輸出用：{ termKind, termFrom, termTo }（沒有類型回 null＝不多任何鍵） */
function termView(it) {
  if (!hasTerm(it)) return null;
  return { termKind: QT.parseKind(it.termKind), termFrom: clean(it.termFrom), termTo: clean(it.termTo) };
}

/** 送簽前檢查（buildPreview blockers）：已儲存的 License／MA 品項缺期間或期間不合法 → [{ code, message }]。功能關閉回 [] */
function blockersFor(q, on) {
  if (!on) return [];
  const out = [];
  let n = 0;
  (q && Array.isArray(q.items) ? q.items : []).forEach((it) => {
    if (!QI.isItemRow(it)) return;
    n++;
    const k = QT.parseKind(it && it.termKind);
    if (!k && k !== '') { out.push({ code: 'TERM_INVALID', message: rowMsg(n, it.desc, '的類型不正確') }); return; }
    const bad = QT.validateTerm(k, it && it.termFrom, it && it.termTo);
    if (bad) out.push({ code: bad.code, message: rowMsg(n, it.desc, bad.code === 'TERM_REQUIRED' ? bad.reason : '：' + bad.reason) });
  });
  return out;
}

const dash = (v) => (v === undefined || v === null || v === '' ? '∅' : String(v));
/** 兩列（依 lid 對應，皆為一般品項）之間期間欄位的異動：['類型 一般→License', '期間 ∅→2026/01/01 – 2026/12/31']；沒變回 [] */
function termChanges(o, n) {
  const out = [];
  const ok = QT.parseKind(o && o.termKind) || '', nk = QT.parseKind(n && n.termKind) || '';
  if (ok !== nk) out.push(`類型 ${QT.KIND_LABEL[ok]}→${QT.KIND_LABEL[nk]}`);
  const orng = ok ? (QT.formatRange(clean(o.termFrom), clean(o.termTo)) || (clean(o.termFrom) + '~' + clean(o.termTo))) : '';
  const nrng = nk ? (QT.formatRange(clean(n.termFrom), clean(n.termTo)) || (clean(n.termFrom) + '~' + clean(n.termTo))) : '';
  if (orng !== nrng) out.push(`期間 ${dash(orng)}→${dash(nrng)}`);
  return out;
}

module.exports = Object.assign({}, QT, { enabled, resolveRowTerm, pairWithOld, validateRawItems, hasTerm, termView, blockersFor, termChanges });
