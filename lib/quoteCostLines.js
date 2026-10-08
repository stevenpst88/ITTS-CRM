/**
 * 報價單「成本明細」(costLines) —— 純函式、零相依（quoteApproval／quoteRoutes／quotePnlExcel 共用；
 * 前端 _client/quote-costlines.js 的 QCL.totals 是它的「顯示用浮點」鏡像，伺服器永遠以這裡為準）。
 *
 * 資料模型：q.costLines（陣列）。欄位存在（即使是 []）＝新式成本；欄位不存在＝舊式（成本在 items[].cost，行為不變）。
 *   CostLine = {
 *     lid, cat('consult'|'software'|'hw'|'travel'|'other'), desc(必填≤120), vendor(≤60，只有 consult/software/hw 保留),
 *     note(≤200), unit(≤10，預設「式」), qty(0..1e9), unitCost(0..1e12，元／單位),
 *     auto?:'stamp'（僅印花稅），forLid?（由哪個客戶品項帶入，只供「涵蓋判定」，不影響計算、不進 contentHash）,
 *     forLids?（字串陣列，上限 60 個；「合併為一列」時帶著所有被併列的 forLid∪forLids，同樣只供涵蓋判定）
 *   }
 * 印花稅（auto:'stamp'）：同一單最多 1 列；伺服器強制 cat:'other'、qty:1、unit:'式'、desc:STAMP_DESC，
 * 並且「不存」client 傳來的 unitCost（存 0）；計算時 unitCost＝整數元 round-half-up(折扣後未稅營收 × 0.001)。
 * 列存在＝計入、不存在＝不計。
 *
 * ── 匯出（簽名；金額單位一律是「分」，除非另註）──────────────────────────────
 *   CATS                                   ['consult','software','hw','travel','other']（對應 PNL 的 1 顧問／2 軟體／3 硬體／4 差旅／5 其他）
 *   CAT_LABELS                             { consult:'顧問服務成本', software:'軟體成本', hw:'硬體成本', travel:'差旅費用', other:'其他費用' }
 *   MAX_COST_LINES                         60（含所有分類）
 *   STAMP_DESC                             '印花稅(合約金額×0.1%)'
 *   MAX_CENTS_NUM                          1e15（單列／合計金額上限，超過視為過大）
 *   hasCostLines(q) → boolean              Array.isArray(q.costLines)
 *   isStamp(line) → boolean                line.auto === 'stamp'
 *   normalizeCostLines(raw, opts) → {ok:true, lines} | {ok:false, error:{code,status,message}}
 *       opts: { genLid:()=>string（指派新 lid；缺省用內建隨機字串）, prev:[舊列]（客戶傳來的 lid 若在 prev 裡就保留，否則視為新列並指派新 lid） }
 *       驗證：raw 必須是陣列且 ≤60；未知 cat／欄位型別錯／qty 0..1e9／unitCost 0..1e12（非有限數、負數、'12abc' 都拒絕）／desc 空白／
 *       auto 不是 'stamp'／印花稅超過 1 列 → 400 BAD_COST_LINE；超過上限 → 400 TOO_MANY_COST_LINES；客戶傳來的 lid 重複 → 409 DUP_LID。
 *       qty／unitCost 必填：number，或純十進位字串（不含逗號、科學記號、十六進位）；缺少、null、空字串都算錯（前端空白欄位請先轉成 0）。
 *       desc 一律必填（空白即 400，不分是否按「完成」）；按「完成」另外用 checkDone 檢查 qty>0 與總成本>0。
 *       字串欄位：去頭尾空白、控制字元（含換行）換成空白、超過長度直接截斷（比照 sanitizeStr）。
 *       forLids：不是陣列、有非字串元素、超過 60 個 → 400 BAD_COST_LINE（與 forLid 型別錯的處理一致）；元素去頭尾空白、截 64 字、去空白與重複、剔除與 forLid 相同者，清完是空的就不輸出；印花稅列一律丟掉。
 *       輸出每列的鍵順序固定：lid,cat,desc,vendor,note,unit,qty,unitCost[,auto][,forLid][,forLids]；vendor 一律有（travel/other 為 ''）。
 *   checkDone(lines) → {ok:true} | {ok:false, code:'BAD_COST_LINE'|'MISSING_COST', message}
 *       顧問按「完成」的額外檢查：非印花稅列 qty>0；至少 1 列非印花稅列且其總成本>0。
 *   costLinesComplete(q) → boolean         新式成本是否完整（非印花稅列 ≥1 且總成本>0）；computeFinancials 與路由的 costComplete 共用
 *   stampDollars(revenueCents) → number    印花稅整數「元」＝round-half-up(營收(分)/100 × 0.001)；revenueCents 可為 number 或 bigint，無效或 ≤0 → 0
 *   effectiveLines(q, revenueCents) → [CostLine]   複本；印花稅列帶入 qty:1、unitCost:stampDollars(revenueCents)。沒有 costLines → []
 *   lineCentsBig(line) → bigint            單列成本（分，BigInt）＝round-half-up(qty×unitCost×100)，與 computeFinancials 對 items 的取整同一套；
 *                                          資料壞掉（非數字／負數）拋出 RangeError。印花稅列請先用 effectiveLines 帶入 unitCost
 *   lineCents(line) → number               同上轉成 number（≤ MAX_CENTS_NUM 內精確）；資料壞掉回 NaN
 *   totalsByCat(q, revenueCents) → {ok, consult, software, hw, travel, other, total, stamp, nonStamp, count, nonStampCount}
 *       各分類成本與合計（other 含印花稅），皆為 number（分）。stamp＝印花稅那列的金額，nonStamp＝total−stamp。
 *       失敗（資料壞掉／過大）→ ok:false 並帶 code('BAD_COST_LINE'|'TOO_LARGE')、index、message，數值欄位為 0。沒有 costLines → 全 0
 *   publicLines(q, {canSeeCost, canSeePrice, revenueCents}) → [CostLine] | undefined
 *       serialize 用：沒有 costLines 或 !canSeeCost → undefined；印花稅列的 unitCost 對所有 canSeeCost 的人都給（值＝stampDollars(revenueCents)）。
 *       業務規則（業主 2026-10-07 確認）：顧問與業務本來就可以互相知道成本與最終售價，印花稅金額（營收×0.1%）不需要對顧問隱藏。
 *       注意：這只限「成本明細」這個功能；顧問對品項單價／折扣／簽核資訊的可見性不在此放寬。canSeePrice 保留為參數但不影響輸出。
 *       只輸出已知欄位（lid,cat,desc,vendor,note,unit,qty,unitCost[,auto][,forLid][,forLids]）
 *   unmatchedItems(q) → [{lid, name, index}]   「有價品項沒有成本列涵蓋」名單（只提醒、不擋）：q 有 costLines 才算（舊式單回 []）；
 *       有價＝unitPrice>0 的一般品項（略過 title／subtotal）。涵蓋規則（前端 QCL.unmatchedPricedItems 逐字鏡像，scripts/check-quote-costlines-ui.js 用 vm 隨機比對）：
 *       ① 成本列（印花稅列除外）的 forLid／forLids 命中「目前任何一般品項」的 lid → 該列「指向品項」，命中的有價品項視為已涵蓋；
 *       ② 沒有命中任何品項的列（forLid／forLids 空，或全都指向已不存在的品項）→ 以去頭尾空白的品名多重集合比對，每列最多涵蓋一個同名有價品項；
 *       品名空白的有價品項只能靠 forLid／forLids 涵蓋。品項沒有 lid 時，lid 以 'legacy-<陣列索引>' 代替（與 serialize 給前端的值一致）。
 *   costWarnings(q) → [{code:'ITEMS_WITHOUT_COST_LINE', count, message}] | []   有未涵蓋品項才有一則；message 只含品名（最多 5 個、每個截 30 字）與總數，不含任何金額
 *   resolveItemRefs(lines, map) → [CostLine]   存檔時把成本列 forLid／forLids 裡畫面給新品項的暫時代號（'nid-…'）換成真正的 lid（map：Map(nid → lid)，由請求裡 items 的 nid 與
 *       正規化後 items 的順序對出來）；換不到的暫時代號丟掉；其餘原樣。必須在 backfillForLids 之前跑（先換、再用品名補剩下的）
 *   backfillForLids(items, lines) → [CostLine]   新單第一次儲存時品項才剛拿到 lid，成本列的 forLid 是空的：對「沒有 forLid／forLids 的成本列」依序以品名
 *       對應到「尚未被涵蓋」的同名一般品項並回填 forLid（只在品名完全相同才填、對不到不填、不覆蓋既有 forLid／forLids、印花稅列與 title／subtotal 不參與）；回傳新陣列，不改輸入
 *   summarizeChanges(prevLines, nextLines, max=10) → [string]   稽核／diffSummary 用的異動文字（增／刪／改，品名舊→新），不會出現 undefined；
 *                                                               超過 max 筆時最後多一句「…另 N 筆」
 *   diffCounts(prevLines, nextLines) → {added, removed, changed}   異動筆數（順序調整不計）
 *   describeLines(lines, max=10) → string  成本明細的簡短摘要：「品名」=金額元；印花稅不寫金額；超過 max 筆加「…另 N 筆」
 *   describeLinesFull(lines, {maxLines=60, nameMax=30}) → string
 *       重置前成本明細的完整留痕（需顧問↔不需顧問切換時整批明細被刪除，寫進稽核）：逐列「品名」=數量×單價=金額（廠商），
 *       列出全部（上限＝MAX_COST_LINES）；品名／廠商各截斷 nameMax 字；印花稅只寫「印花稅」
 */
'use strict';

const CATS = Object.freeze(['consult', 'software', 'hw', 'travel', 'other']);
const CAT_LABELS = Object.freeze({ consult: '顧問服務成本', software: '軟體成本', hw: '硬體成本', travel: '差旅費用', other: '其他費用' });
const CAT_SHORT = Object.freeze({ consult: '顧問', software: '軟體', hw: '硬體', travel: '差旅', other: '其他' });
const VENDOR_CATS = Object.freeze(['consult', 'software', 'hw']);   // 這幾區有「委外廠商／供應商」欄；差旅與其他區沒有
const MAX_COST_LINES = 60;
const STAMP_DESC = '印花稅(合約金額×0.1%)';
const MAX_CENTS = 1000000000000000n;       // 與 quoteApproval.MAX_CENTS 同值（單列／合計上限）
const MAX_CENTS_NUM = Number(MAX_CENTS);
const MAX_QTY = 1e9;
const MAX_UNIT_COST = 1e12;
const LIMITS = Object.freeze({ lid: 64, desc: 120, vendor: 60, note: 200, unit: 10, forLid: 64 });
const MAX_FOR_LIDS = 60;   // forLids 陣列元素上限（每個字串長度上限同 forLid）

const hasCostLines = (q) => !!q && typeof q === 'object' && Array.isArray(q.costLines);
const isStamp = (l) => !!l && typeof l === 'object' && l.auto === 'stamp';

// ───────────────────────── 十進位精算（與 quoteApproval 的 parseDec／divRound 同一套規則） ─────────────────────────
const POW10 = [1n];
function pow10(n) { while (POW10.length <= n) POW10.push(POW10[POW10.length - 1] * 10n); return POW10[n]; }
function divRound(num, den) { return (2n * num + den) / (2n * den); }   // half-up，num/den 皆非負，den>0

/** → {empty}|{bad}|{big}|{ok,n,s,neg}（鏡像 quoteApproval.parseDec） */
function parseDec(x) {
  if (x === null || x === undefined) return { empty: true };
  let str;
  if (typeof x === 'number') {
    if (!Number.isFinite(x)) return { bad: true };
    str = String(x);
  } else if (typeof x === 'string') {
    str = x.trim().replace(/,/g, '');
    if (str === '') return { empty: true };
  } else return { bad: true };
  const m = /^([+-]?)(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(str);
  if (!m) return { bad: true };
  const intPart = m[2] || '', fracPart = m[3] || '';
  if (intPart === '' && fracPart === '') return { bad: true };
  const exp = m[4] ? parseInt(m[4], 10) : 0;
  if (exp > 30) return { big: true };
  const digits = (intPart + fracPart).replace(/^0+(?=\d)/, '');
  let n = BigInt(digits === '' ? '0' : digits);
  let s = fracPart.length - exp;
  if (s < 0) { n = n * pow10(-s); s = 0; }
  if (s > 400) { n = 0n; s = 0; }
  return { ok: true, n, s, neg: m[1] === '-' && n !== 0n };
}

/** 單列成本（分，BigInt）。回 {ok,cents}|{bad}|{big}。qty／unitCost 空值視為 0（正規化後的資料一定有值） */
function centsOf(line) {
  const l = (line && typeof line === 'object') ? line : {};
  const q = parseDec(l.qty), c = parseDec(l.unitCost);
  if (q.big || c.big) return { big: true };
  if (q.bad || c.bad || q.neg || c.neg) return { bad: true };
  const qn = q.empty ? 0n : q.n, qs = q.empty ? 0 : q.s;
  const cn = c.empty ? 0n : c.n, cs = c.empty ? 0 : c.s;
  const cents = divRound(qn * cn * 100n, pow10(qs + cs));
  if (cents > MAX_CENTS) return { big: true };
  return { ok: true, cents };
}
function lineCentsBig(line) {
  const r = centsOf(line);
  if (r.ok) return r.cents;
  throw new RangeError(r.big ? '成本明細金額過大' : '成本明細資料無效');
}
function lineCents(line) {
  const r = centsOf(line);
  return r.ok ? Number(r.cents) : NaN;
}

// ───────────────────────── 印花稅 ─────────────────────────
/** 整數元＝round-half-up(營收(分)/100 × 0.001)＝round-half-up(營收分 / 100000) */
function stampDollars(revenueCents) {
  let c;
  try { c = (typeof revenueCents === 'bigint') ? revenueCents : BigInt(Math.round(Number(revenueCents))); } catch (_) { return 0; }
  if (c <= 0n) return 0;
  return Number((c + 50000n) / 100000n);
}

function effectiveLines(q, revenueCents) {
  if (!hasCostLines(q)) return [];
  const yuan = stampDollars(revenueCents);
  return q.costLines.map((l) => (isStamp(l) ? Object.assign({}, l, { qty: 1, unitCost: yuan }) : Object.assign({}, l)));
}

// ───────────────────────── 合計 ─────────────────────────
function emptyTotals() {
  return { ok: true, consult: 0, software: 0, hw: 0, travel: 0, other: 0, total: 0, stamp: 0, nonStamp: 0, count: 0, nonStampCount: 0 };
}
function totalsByCat(q, revenueCents) {
  const out = emptyTotals();
  if (!hasCostLines(q)) return out;
  const lines = effectiveLines(q, revenueCents);
  const sum = { consult: 0n, software: 0n, hw: 0n, travel: 0n, other: 0n };
  let total = 0n, stamp = 0n, nonStampCount = 0;
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    const r = centsOf(l);
    if (!r.ok) {
      const z = emptyTotals();
      z.ok = false; z.code = r.big ? 'TOO_LARGE' : 'BAD_COST_LINE'; z.index = i;
      z.message = '成本明細第 ' + (i + 1) + ' 筆' + (r.big ? '金額過大' : '資料無效');
      return z;
    }
    const cat = CATS.includes(l && l.cat) ? l.cat : 'other';   // 壞資料的分類一律當「其他」：寧可多算成本，不能漏算
    sum[cat] += r.cents; total += r.cents;
    if (isStamp(l)) stamp += r.cents; else nonStampCount++;
    if (total > MAX_CENTS) {
      const z = emptyTotals();
      z.ok = false; z.code = 'TOO_LARGE'; z.index = i; z.message = '成本明細合計過大';
      return z;
    }
  }
  CATS.forEach((k) => { out[k] = Number(sum[k]); });
  out.total = Number(total); out.stamp = Number(stamp); out.nonStamp = Number(total - stamp);
  out.count = lines.length; out.nonStampCount = nonStampCount;
  return out;
}

/** 新式成本是否完整：非印花稅列至少 1 列，且非印花稅列總成本 > 0（與營收無關，所以用 revenue=0 算即可） */
function costLinesComplete(q) {
  if (!hasCostLines(q)) return false;
  const t = totalsByCat(q, 0);
  return t.ok && t.nonStampCount >= 1 && t.nonStamp > 0;
}

// ───────────────────────── 正規化（寫入前） ─────────────────────────
function badLine(message) { return { ok: false, error: { code: 'BAD_COST_LINE', status: 400, message } }; }

// 控制字元（含換行、DEL、行／段分隔符）→ 空白；用 fromCharCode 組，避免原始碼裡出現不可見字元
const CTRL_RE = new RegExp('[' + String.fromCharCode(0) + '-' + String.fromCharCode(31) + String.fromCharCode(127) + String.fromCharCode(0x2028) + String.fromCharCode(0x2029) + ']+', 'g');
/** 單行文字欄位：非字串（含 null／undefined 以外的型別）→ null＝型別錯；字串 → 控制字元換空白、去頭尾空白、截斷 */
function cleanText(v, max) {
  if (v === undefined || v === null) return '';
  if (typeof v !== 'string') return null;
  return v.replace(CTRL_RE, ' ').trim().slice(0, max);
}

// 純十進位數字的字串形式。分支以「第一個字元是不是 '.'」互斥、且小數點之後的 \d* 只有一種切法，沒有歧義，
// 不會做二次方回溯（舊寫法 \d+\.?\d* 在「很長的數字＋一個非法字元」時是 O(n²)，2MB 的請求可以卡死事件迴圈數分鐘）。
const STRICT_NUM_RE = /^-?(?:\d+(?:\.\d*)?|\.\d+)$/;
// 數字字串長度上限（含前後空白）：超過一律視為非法數字，先擋長度再跑正規式，使用者可控的長字串永遠不會進正規式。
// 40 個字元足以容納 unitCost ≤ 1e12 加上 20 位小數；超過者不是正常輸入。
const MAX_NUM_STR_LEN = 40;

/** 嚴格數字：number（有限）或純十進位字串（不含逗號、科學記號、十六進位、尾巴雜字）。回 {v}|{err} */
function strictNum(v, label, min, max) {
  let n;
  if (typeof v === 'number') n = v;
  else if (typeof v === 'string' && v.length <= MAX_NUM_STR_LEN && STRICT_NUM_RE.test(v.trim())) n = Number(v.trim());
  else return { err: label + '必須是數字' };
  if (!Number.isFinite(n)) return { err: label + '必須是有限的數字' };
  if (n < min) return { err: label + '不可小於 ' + min };
  if (n > max) return { err: label + '不可大於 ' + max };
  return { v: n === 0 ? 0 : n };   // -0 → 0
}

let _lidSeq = 0;
function fallbackLid() { return 'cl-' + Date.now().toString(36) + '-' + (++_lidSeq).toString(36) + '-' + Math.random().toString(36).slice(2, 8); }

function normalizeCostLines(raw, opts) {
  opts = opts || {};
  const genLid = typeof opts.genLid === 'function' ? opts.genLid : fallbackLid;
  const prevLids = new Set((Array.isArray(opts.prev) ? opts.prev : []).map((l) => l && l.lid).filter((x) => typeof x === 'string' && x));
  if (!Array.isArray(raw)) return badLine('成本明細必須是陣列');
  if (raw.length > MAX_COST_LINES) return { ok: false, error: { code: 'TOO_MANY_COST_LINES', status: 400, message: '成本明細最多 ' + MAX_COST_LINES + ' 列（含所有分類），請合併相近的項目' } };
  const lines = [];
  const used = new Set();        // 已用 lid（含新指派的）
  const seenClientLid = new Set();
  let stampCount = 0;
  for (let i = 0; i < raw.length; i++) {
    const r = raw[i];
    const at = '成本明細第 ' + (i + 1) + ' 筆';
    if (!r || typeof r !== 'object' || Array.isArray(r)) return badLine(at + '資料格式不正確');
    // auto：只認 'stamp'
    let auto = '';
    if (r.auto !== undefined && r.auto !== null && r.auto !== '') {
      if (r.auto !== 'stamp') return badLine(at + '的 auto 不合法');
      auto = 'stamp';
    }
    // 分類（印花稅一律強制成 other，client 傳什麼都不看）
    const stamp = auto === 'stamp';
    if (!stamp && (typeof r.cat !== 'string' || !CATS.includes(r.cat))) return badLine(at + '的分類不合法（需為 ' + CATS.join('／') + '）');
    if (stamp) {
      stampCount++;
      if (stampCount > 1) return badLine('印花稅最多只能有 1 列');
    }
    const cat = stamp ? 'other' : r.cat;
    // 文字欄位
    const descRaw = cleanText(r.desc, LIMITS.desc);
    const vendorRaw = cleanText(r.vendor, LIMITS.vendor);
    const noteRaw = cleanText(r.note, LIMITS.note);
    const unitRaw = cleanText(r.unit, LIMITS.unit);
    const forLidRaw = (r.forLid === undefined || r.forLid === null) ? '' : (typeof r.forLid === 'string' ? r.forLid.trim().slice(0, LIMITS.forLid) : null);
    if (descRaw === null || vendorRaw === null || noteRaw === null || unitRaw === null || forLidRaw === null) return badLine(at + '的文字欄位型別不正確');
    // forLids：與 forLid 一致——型別不對（不是陣列、含非字串、超過上限）直接 400，不默默吞掉
    let forLidsRaw = [];
    if (r.forLids !== undefined && r.forLids !== null) {
      if (!Array.isArray(r.forLids)) return badLine(at + '的 forLids 必須是陣列');
      if (r.forLids.length > MAX_FOR_LIDS) return badLine(at + '的 forLids 最多 ' + MAX_FOR_LIDS + ' 個');
      const seenF = new Set();
      for (let k = 0; k < r.forLids.length; k++) {
        const e = r.forLids[k];
        if (typeof e !== 'string') return badLine(at + '的 forLids 只能放字串');
        const s = e.trim().slice(0, LIMITS.forLid);
        if (s && s !== forLidRaw && !seenF.has(s)) { seenF.add(s); forLidsRaw.push(s); }
      }
    }
    const desc = stamp ? STAMP_DESC : descRaw;
    if (!desc) return badLine(at + '的項目名稱不可空白');
    // 數字
    let qty = 1, unitCost = 0;
    if (!stamp) {
      const q = strictNum(r.qty, at + '的數量', 0, MAX_QTY);
      if (q.err) return badLine(q.err);
      const c = strictNum(r.unitCost, at + '的成本單價', 0, MAX_UNIT_COST);
      if (c.err) return badLine(c.err);
      qty = q.v; unitCost = c.v;
    }
    // lid：客戶傳來的 lid 在舊列裡才保留；同一次請求裡重複 → DUP_LID
    let lid = '';
    if (typeof r.lid === 'string' && r.lid.trim()) {
      const cl = r.lid.trim().slice(0, LIMITS.lid);
      if (seenClientLid.has(cl)) return { ok: false, error: { code: 'DUP_LID', status: 409, message: '成本明細的列代碼重複，請重新整理頁面後再編輯' } };
      seenClientLid.add(cl);
      if (prevLids.has(cl) && !used.has(cl)) lid = cl;
    }
    if (!lid) {
      for (let k = 0; k < 8 && (!lid || used.has(lid) || prevLids.has(lid) || seenClientLid.has(lid)); k++) lid = String(genLid());
      if (!lid || used.has(lid)) return badLine('無法指派成本明細的列代碼');
    }
    used.add(lid);
    const line = {
      lid, cat, desc,
      vendor: VENDOR_CATS.includes(cat) && !stamp ? vendorRaw : '',
      note: noteRaw,
      unit: stamp ? '式' : (unitRaw || '式'),
      qty, unitCost,
    };
    if (stamp) line.auto = 'stamp';
    else {
      if (forLidRaw) line.forLid = forLidRaw;
      if (forLidsRaw.length) line.forLids = forLidsRaw;
    }
    // 單列金額上限（避免存下讓 computeFinancials 一律 TOO_LARGE 的資料）
    const cr = centsOf(line);
    if (!cr.ok) return badLine(at + '的金額過大');
    lines.push(line);
  }
  // 合計上限
  const tot = totalsByCat({ costLines: lines }, 0);
  if (!tot.ok) return badLine('成本明細合計金額過大');
  return { ok: true, lines };
}

/** 顧問按「完成」時的額外檢查（需要已正規化的列） */
function checkDone(lines) {
  const arr = Array.isArray(lines) ? lines : [];
  for (let i = 0; i < arr.length; i++) {
    const l = arr[i];
    if (isStamp(l)) continue;
    if (!(Number(l.qty) > 0)) return { ok: false, code: 'BAD_COST_LINE', message: '成本明細第 ' + (i + 1) + ' 筆「' + String(l.desc || '').slice(0, 12) + '」的數量需大於 0' };
  }
  if (!costLinesComplete({ costLines: arr })) return { ok: false, code: 'MISSING_COST', message: '尚未填寫成本明細（至少需要一列非印花稅的成本，且成本需大於 0）' };
  return { ok: true };
}

// ───────────────────────── 輸出 ─────────────────────────
function pickLine(l, unitCost) {
  const o = { lid: l.lid, cat: l.cat, desc: l.desc, vendor: l.vendor || '', note: l.note || '', unit: l.unit || '式', qty: l.qty };
  if (unitCost !== undefined) o.unitCost = unitCost;
  if (l.auto === 'stamp') o.auto = 'stamp';
  else {
    if (typeof l.forLid === 'string' && l.forLid) o.forLid = l.forLid;
    if (Array.isArray(l.forLids)) {
      const f = l.forLids.filter((x) => typeof x === 'string' && x).slice(0, MAX_FOR_LIDS);
      if (f.length) o.forLids = f;
    }
  }
  return o;
}

function publicLines(q, perm) {
  if (!hasCostLines(q)) return undefined;
  perm = perm || {};
  if (!perm.canSeeCost) return undefined;
  const yuan = stampDollars(perm.revenueCents);
  return q.costLines.map((l) => {
    if (isStamp(l)) return pickLine(Object.assign({}, l, { qty: 1 }), yuan);
    return pickLine(l, l.unitCost);
  });
}

// ───────────────────────── 涵蓋判定：有價品項有沒有成本列（只提醒、不擋）＆ forLid 回填 ─────────────────────────
// ⚠ 這一段與 _client/quote-costlines.js 的 unmatchedPricedItems／backfill 判定逐字鏡像（前端畫面即時提醒用同一套規則）；
//   scripts/check-quote-costlines-ui.js 以 vm 載入前端、用隨機輸入逐組比對兩邊結果，改這裡必須同步改前端。
const WARN_NAME_MAX = 30;       // 警告文字裡單一品名最長字數
const WARN_LIST_MAX = 5;        // 警告文字最多列幾個品名
const UNNAMED_ITEM = '（未命名品項）';

/** 品名比對鍵：去控制字元（與成本列品名同一套清洗）、去頭尾空白、截 120 字；null／undefined → '' */
function nameKey(v) {
  if (v === undefined || v === null) return '';
  return String(v).replace(CTRL_RE, ' ').trim().slice(0, LIMITS.desc);
}
/** lid 比對鍵：只認字串，去頭尾空白、截 64 字；其餘 → '' */
function lidKey(v) { return typeof v === 'string' ? v.trim().slice(0, LIMITS.lid) : ''; }

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

/** 一般品項（略過 title／subtotal 與非物件）→ [{lid, name, priced, index}]；index＝在傳入陣列中的位置，沒有 lid 時 lid＝'legacy-'+index（同 serialize） */
function generalEntries(items) {
  const out = [];
  (Array.isArray(items) ? items : []).forEach((it, i) => {
    if (!it || typeof it !== 'object' || it.kind === 'title' || it.kind === 'subtotal') return;
    const p = parseFloat(it.unitPrice);
    out.push({ lid: lidKey(it.lid) || ('legacy-' + i), name: nameKey(it.desc), priced: isFinite(p) && p > 0, index: i });
  });
  return out;
}

/**
 * 核心：cands（entries 的子集，依品項順序）裡沒有被成本列涵蓋的那幾個。
 * 成本列（略過印花稅）命中 entries 任一 lid → 指向品項（命中者算涵蓋）；沒命中的列進「自由列」池，以品名多重集合涵蓋同名候選（空白品名不進池也不被涵蓋）。
 */
function uncovered(lines, entries, cands) {
  const lidSet = new Set(entries.map((e) => e.lid));
  const covered = new Set();
  const free = new Map();
  (Array.isArray(lines) ? lines : []).forEach((l) => {
    if (!l || typeof l !== 'object' || isStamp(l)) return;
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

/** 有價品項沒有被任何成本列涵蓋的名單（規則見檔頭）；舊式單（沒有 costLines 欄位）回 [] */
function unmatchedItems(q) {
  if (!hasCostLines(q)) return [];
  const entries = generalEntries(q.items);
  return uncovered(q.costLines, entries, entries.filter((e) => e.priced)).map((e) => ({ lid: e.lid, name: e.name, index: e.index }));
}

/** 警告文字（伺服器 preview.warnings／derived.costWarnings 與前端 QCL.unmatchedPricedNote 共用同一個格式）。names：未涵蓋品項的品名陣列（可含 ''） */
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

/** 建議性警告（不擋送簽）；沒有未涵蓋品項（含舊式單）回 []。只含品名與數量，不含任何金額 */
function costWarnings(q) {
  const un = unmatchedItems(q);
  if (!un.length) return [];
  return [{ code: 'ITEMS_WITHOUT_COST_LINE', count: un.length, message: unmatchedMessage(un.map((x) => x.name)) }];
}

// ── 暫時代號（nid）──
// 畫面上還沒存檔的新品項沒有 lid；編輯器的成本列（種子、合併為一列、補入新品項）改用畫面給每個新品項的暫時代號指向它（forLid／forLids 放 'nid-…'）。
// 存檔時（品項才剛拿到 lid）依請求裡 items 的順序把暫時代號換成真正的 lid；換不到的暫時代號丟掉（不留垃圾）。nid 只活在這一次請求裡，不存檔。
const TMP_REF_PREFIX = 'nid-';

/**
 * 把成本列 forLid／forLids 裡的暫時代號換成真正的品項 lid。map：Map(暫時代號 → lid)。
 * 在 map 裡的換成 lid；不在 map 裡、以 'nid-' 開頭的丟掉；其餘（真正的 lid 等）原樣保留。換完 forLids 去重、剔除與 forLid 相同者。
 * 印花稅列、沒有 forLid／forLids 的列不動（同一個物件）；有動的列回傳新物件（鍵順序維持 …unitCost,forLid,forLids）。不改輸入。
 */
function resolveItemRefs(lines, map) {
  const arr = Array.isArray(lines) ? lines : [];
  const m = map instanceof Map ? map : new Map();
  const fix = (k) => {
    const s = lidKey(k);
    if (!s) return '';
    if (m.has(s)) return lidKey(m.get(s));
    return s.indexOf(TMP_REF_PREFIX) === 0 ? '' : s;
  };
  return arr.map((l) => {
    if (!l || typeof l !== 'object' || isStamp(l)) return l;
    if (!lineLidList(l).length) return l;
    const forLid = fix(l.forLid);
    const seen = new Set();
    const forLids = [];
    (Array.isArray(l.forLids) ? l.forLids.slice(0, MAX_FOR_LIDS) : []).forEach((e) => {
      const s = fix(e);
      if (s && s !== forLid && !seen.has(s)) { seen.add(s); forLids.push(s); }
    });
    const o = Object.assign({}, l);
    delete o.forLid; delete o.forLids;
    if (forLid) o.forLid = forLid;
    if (forLids.length) o.forLids = forLids;
    return o;
  });
}

/** 回填 forLid（規則見檔頭）；回傳新陣列，未被回填的列維持原物件 */
function backfillForLids(items, lines) {
  const arr = Array.isArray(lines) ? lines : [];
  const entries = generalEntries(items);
  const lidSet = new Set(entries.map((e) => e.lid));
  const covered = new Set();
  arr.forEach((l) => {
    if (!l || typeof l !== 'object' || isStamp(l)) return;
    lineLidList(l).forEach((k) => { if (lidSet.has(k)) covered.add(k); });
  });
  return arr.map((l) => {
    if (!l || typeof l !== 'object' || isStamp(l)) return l;
    if (lineLidList(l).length) return l;          // 已有 forLid／forLids（即使指向已不存在的品項）：不覆蓋
    const d = nameKey(l.desc);
    if (!d) return l;
    const e = entries.find((x) => x.name === d && !covered.has(x.lid));
    if (!e) return l;
    covered.add(e.lid);
    return Object.assign({}, l, { forLid: e.lid });
  });
}

// ───────────────────────── 稽核文字 ─────────────────────────
const dash =(v) => (v === undefined || v === null || v === '' ? '∅' : String(v));
const labelOf = (l) => (isStamp(l) ? '「印花稅」' : '「' + String((l && l.desc) || '').slice(0, 12) + '」');
function fmtCents(c) {
  const b = BigInt(c);
  const whole = (b / 100n).toString();
  const frac = b % 100n;
  return frac === 0n ? whole : whole + '.' + (frac < 10n ? '0' : '') + frac.toString();
}

/** 異動明細：[{type:'add'|'del'|'mod'|'order', text}] */
function diffEntries(prevLines, nextLines) {
  const a = Array.isArray(prevLines) ? prevLines : [];
  const b = Array.isArray(nextLines) ? nextLines : [];
  const byLid = new Map(a.map((l) => [l && l.lid, l]));
  const seen = new Set();
  const msgs = [];
  const add = (type, text) => msgs.push({ type, text });
  b.forEach((n) => {
    seen.add(n && n.lid);
    const o = byLid.get(n && n.lid);
    if (!o) {
      add('add', '新增' + labelOf(n) + '（' + (CAT_SHORT[n.cat] || dash(n.cat)) + '）' + (isStamp(n) ? '' : '數量 ' + dash(n.qty) + '、單價 ' + dash(n.unitCost)));
      return;
    }
    const ch = [];
    if (String(o.desc || '') !== String(n.desc || '')) ch.push('品名 ' + dash(o.desc) + '→' + dash(n.desc));
    if (o.cat !== n.cat) ch.push('分類 ' + (CAT_SHORT[o.cat] || dash(o.cat)) + '→' + (CAT_SHORT[n.cat] || dash(n.cat)));
    if (String(o.vendor || '') !== String(n.vendor || '')) ch.push('廠商 ' + dash(o.vendor) + '→' + dash(n.vendor));
    if (String(o.note || '') !== String(n.note || '')) ch.push('說明');
    if (String(o.unit || '') !== String(n.unit || '')) ch.push('單位 ' + dash(o.unit) + '→' + dash(n.unit));
    if (!isStamp(n) || !isStamp(o)) {
      if (Number(o.qty) !== Number(n.qty)) ch.push('數量 ' + dash(o.qty) + '→' + dash(n.qty));
      if (Number(o.unitCost) !== Number(n.unitCost)) ch.push('單價 ' + dash(o.unitCost) + '→' + dash(n.unitCost));
    }
    if (isStamp(o) !== isStamp(n)) ch.push(isStamp(n) ? '改為印花稅' : '不再是印花稅');
    if (ch.length) add('mod', labelOf(n) + ch.join('、'));
  });
  a.forEach((o) => { if (!seen.has(o && o.lid)) add('del', '刪除' + labelOf(o) + '（' + (CAT_SHORT[o && o.cat] || dash(o && o.cat)) + '）'); });
  if (!msgs.length) {
    // 沒有任何內容異動：只可能是順序不同
    const ord = (arr) => arr.map((l) => l && l.lid).join(',');
    if (a.length === b.length && ord(a) !== ord(b)) add('order', '列順序調整');
  }
  return msgs;
}

function summarizeChanges(prevLines, nextLines, max) {
  const lim = max === undefined ? 10 : max;
  const e = diffEntries(prevLines, nextLines);
  const out = e.slice(0, lim).map((x) => x.text);
  if (e.length > lim) out.push('…另 ' + (e.length - lim) + ' 筆');
  return out;
}

/** 異動筆數：{added, removed, changed}（順序調整不計） */
function diffCounts(prevLines, nextLines) {
  const c = { added: 0, removed: 0, changed: 0 };
  diffEntries(prevLines, nextLines).forEach((x) => { if (x.type === 'add') c.added++; else if (x.type === 'del') c.removed++; else if (x.type === 'mod') c.changed++; });
  return c;
}

function describeLines(lines, max) {
  const lim = max === undefined ? 10 : max;
  const arr = Array.isArray(lines) ? lines : [];
  const parts = arr.slice(0, lim).map((l) => {
    if (isStamp(l)) return labelOf(l);
    const r = centsOf(l);
    return labelOf(l) + '=' + (r.ok ? fmtCents(r.cents) : '?');
  });
  if (arr.length > lim) parts.push('…另 ' + (arr.length - lim) + ' 筆');
  return parts.join('、');
}

/**
 * 重置前成本明細的「完整」留痕（需顧問↔不需顧問切換時整批成本明細會被刪除，這是唯一能追回的紀錄）：
 * 逐列「「品名」=數量×單價=金額（廠商）」，列出全部（上限 MAX_COST_LINES＝60 列，與資料上限一致，正常不會截斷）；
 * 品名／廠商各截斷 30 字避免單列過長；印花稅只寫「印花稅」（金額依營收計算，不存）；壞資料的金額寫「?」。
 */
function describeLinesFull(lines, opts) {
  opts = opts || {};
  const maxLines = Number.isInteger(opts.maxLines) && opts.maxLines > 0 ? opts.maxLines : MAX_COST_LINES;
  const nameMax = Number.isInteger(opts.nameMax) && opts.nameMax > 0 ? opts.nameMax : 30;
  const arr = Array.isArray(lines) ? lines : [];
  const clip = (v) => String(v === undefined || v === null ? '' : v).replace(CTRL_RE, ' ').trim().slice(0, nameMax);
  const parts = arr.slice(0, maxLines).map((l) => {
    if (isStamp(l)) return '「印花稅」';
    const r = centsOf(l);
    const vendor = clip(l && l.vendor);
    return '「' + clip(l && l.desc) + '」=' + dash(l && l.qty) + '×' + dash(l && l.unitCost) + '=' + (r.ok ? fmtCents(r.cents) : '?') + (vendor ? '（' + vendor + '）' : '');
  });
  if (arr.length > maxLines) parts.push('…另 ' + (arr.length - maxLines) + ' 筆');
  return parts.join('、');
}

module.exports = {
  CATS, CAT_LABELS, MAX_COST_LINES, STAMP_DESC, MAX_CENTS_NUM,
  hasCostLines, isStamp, normalizeCostLines, checkDone, costLinesComplete,
  stampDollars, effectiveLines, lineCentsBig, lineCents, totalsByCat, publicLines,
  summarizeChanges, diffCounts, describeLines, describeLinesFull,
  unmatchedItems, costWarnings, unmatchedMessage, backfillForLids, resolveItemRefs, TMP_REF_PREFIX,
};
