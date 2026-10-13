'use strict';
/**
 * 稽核記錄的「異動前→異動後」組字工具（純函式、無 I/O）。
 *
 * 內控原則：
 *  - 只記欄位層級的 before→after 與筆數，不記密碼、雜湊、金鑰、完整 Email 位址（由呼叫端決定要記哪些欄位；
 *    這裡再加一道保險 scrub：任何看起來像 bcrypt 雜湊的字串一律遮蔽）。
 *  - detail 有長度上限（預設 1500 字），超過時從尾端截斷並明確標示「…(截斷，另有 N 項未顯示)」，不會無聲吞掉。
 *
 *  show(v)                      任意值 → 簡短可讀字串（空值＝∅、布林＝是/否、陣列用 + 串接、字串最長 40 字且去換行）
 *  fieldDiff(before, after, spec)  spec=[{key,label,fmt?}]；回傳 ["標籤 a→b", ...]（只含有變的欄位）
 *  setDiff(beforeArr, afterArr) → { add:[], del:[] }
 *  joinCapped(items, opts)      把多個片段用分隔符串起來，超過 max 時截斷並標示未顯示的數量
 *  scrub(str)                   遮蔽雜湊類字串
 *  MAX_DETAIL                   1500
 */

const MAX_DETAIL = 1500;

function show(v, maxLen) {
  const lim = maxLen || 40;
  if (v === undefined || v === null || v === '') return '∅';
  if (typeof v === 'boolean') return v ? '是' : '否';
  if (Array.isArray(v)) return v.length ? v.map((x) => show(x, lim)).join('+') : '∅';
  const t = String(v).split('').map((ch) => { const c = ch.charCodeAt(0); return (c < 32 || c === 127) ? ' ' : ch; }).join('').trim();
  if (!t) return '∅';
  return t.length > lim ? t.slice(0, lim) + '…' : t;
}

function norm(v) {
  if (v === undefined || v === null || v === '') return null;
  return v;
}

function fieldDiff(before, after, spec) {
  const out = [];
  const b = before || {}, a = after || {};
  (spec || []).forEach((f) => {
    const x = norm(b[f.key]), y = norm(a[f.key]);
    if (JSON.stringify(x) === JSON.stringify(y)) return;
    const fmt = f.fmt || show;
    out.push(`${f.label} ${fmt(x)}→${fmt(y)}`);
  });
  return out;
}

function setDiff(beforeArr, afterArr) {
  const b = Array.isArray(beforeArr) ? beforeArr : [], a = Array.isArray(afterArr) ? afterArr : [];
  return { add: a.filter((x) => !b.includes(x)), del: b.filter((x) => !a.includes(x)) };
}

function scrub(str) {
  return String(str == null ? '' : str).replace(/\$2[aby]\$[0-9]{2}\$[./A-Za-z0-9]{10,}/g, '[已遮蔽]');
}

/**
 * items：字串陣列。依序累加，總長度（含 prefix）超過 max 就停，並在尾端加「…(截斷，另有 N 項未顯示)」。
 * opts: { sep='；', max=1500, prefix='' }
 */
function joinCapped(items, opts) {
  const o = opts || {};
  const sep = o.sep == null ? '；' : o.sep;
  const max = o.max || MAX_DETAIL;
  const prefix = o.prefix || '';
  const list = (items || []).filter((x) => x !== '' && x != null).map(scrub);
  let out = prefix;
  let used = 0;
  for (let i = 0; i < list.length; i++) {
    const piece = (used ? sep : '') + list[i];
    const tail = `${sep}…(截斷，另有 ${list.length - i} 項未顯示)`;
    if (out.length + piece.length + (i < list.length - 1 ? tail.length : 0) > max && used > 0) {
      return out + `${sep}…(截斷，另有 ${list.length - i} 項未顯示)`;
    }
    if (out.length + piece.length > max && used === 0) {
      // 單一片段就超長：硬切
      return out + piece.slice(0, Math.max(0, max - out.length - 20)) + `…(截斷，另有 ${list.length - i - 1 > 0 ? list.length - i - 1 : 0} 項未顯示)`;
    }
    out += piece;
    used++;
  }
  return out;
}

module.exports = { MAX_DETAIL, show, fieldDiff, setDiff, joinCapped, scrub };
