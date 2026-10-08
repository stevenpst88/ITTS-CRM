// 報價單簽核核決引擎 — 純函式（不碰資料庫 / req），只用 node 內建 crypto，方便單元測試
// 規格來源：報價單簽核實作規格 第 1.1、1.2、3 節
//
// 設計重點
//  1. 金額一律整數精算：所有金額以「分」(cents) 計；乘除用 BigInt 十進位精算（不經浮點），四捨五入為 half-up。
//  2. 毛利門檻比較：gpCents * 100 >= threshold * revenueCents（threshold 為整數），不四捨五入、不用浮點百分比。
//  3. 壓線取較低層（剛好等於門檻 → 較低層級）。
//  4. 折扣類型為 percent / amount 但值非法 → 回 BAD_DISCOUNT，不沿用舊前端 `parseFloat(v) || 100` 的怪癖。
//  5. 未歸類商品視為 other 並產生 warning。
const crypto = require('crypto');
const QI = require('./quoteItems');   // 品項列種類（item／title／subtotal）：標題與小計列不計價、不計成本
const CL = require('./quoteCostLines');   // 成本明細 costLines（新式成本；q 沒有 costLines 欄位＝舊式，行為與以前位元級相同）

const RULES_VERSION = '2026-10-06.2';

const TIERS = Object.freeze({
  mgr1: '一級主管',
  gm: '總經理',
  chairman: '董事長',
  board: '董事會決議（秘書代核）',
});

const CLASS_LABELS = Object.freeze({
  consult: '顧問服務',
  software: '軟體規劃服務',
  hardware: '硬體規劃服務',
  crm: '客服服務(CRM)',
  mdm: '帳單列印(MDM)',
  ot: 'OT 策略性專案',
  other: '其他',
});
const CLASS_KEYS = Object.freeze(Object.keys(CLASS_LABELS));

// l1：毛利 >= l1 → 一級主管即可（null＝主管不能終局）；l2：毛利 >= l2 → 總經理；再低 → 董事長
// fixedLevel：不看毛利、固定層級（other）
const ROWS = Object.freeze([
  { key: 'consult', label: '顧問服務', l1: 25, l2: 15 },
  { key: 'software', label: '軟體規劃服務', l1: 10, l2: 5 },
  { key: 'hardware', label: '硬體規劃服務', l1: 10, l2: 5 },
  { key: 'consult_software', label: '顧問服務+軟體服務', l1: 15, l2: 10 },
  { key: 'consult_hardware', label: '顧問服務+硬體服務', l1: 10, l2: 5 },
  { key: 'crm', label: '客服服務(CRM)', l1: 15, l2: 10 },
  { key: 'mdm', label: '帳單列印(MDM)', l1: 10, l2: 5 },
  { key: 'ot', label: 'OT 策略性專案', l1: null, l2: 10 },
  { key: 'other', label: '其他（表外）', l1: null, l2: null, fixedLevel: 2 },
].map(r => Object.freeze(r)));
const ROW_BY_KEY = Object.freeze(ROWS.reduce((m, r) => { m[r.key] = r; return m; }, {}));

// 金額門檻（元，嚴格「大於」）
const AMOUNT = Object.freeze({ chairman: 10000000, board: 50000000 });
const AMOUNT_CENTS = Object.freeze({ chairman: BigInt(AMOUNT.chairman) * 100n, board: BigInt(AMOUNT.board) * 100n });

// 單一金額（分）上限：1e15 分＝1e13 元（十兆），超過視為輸入錯誤，避免超出安全整數
const MAX_CENTS = 1000000000000000n;

const PATHS = Object.freeze({
  1: Object.freeze(['mgr1']),
  2: Object.freeze(['mgr1', 'gm']),
  3: Object.freeze(['mgr1', 'gm', 'chairman']),
  board: Object.freeze(['mgr1', 'gm', 'board']),
});

const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);

// ───────────────────────── 十進位精算工具 ─────────────────────────

const POW10 = [1n];
function pow10(n) {
  while (POW10.length <= n) POW10.push(POW10[POW10.length - 1] * 10n);
  return POW10[n];
}

// 解析成 { n: BigInt（非負的整數部分，含符號另列）, s: 小數位數, neg }；
// 回傳 { empty:true } / { bad:true } / { big:true } / { ok:true, n, s, neg }
function parseDec(x) {
  if (x === null || x === undefined) return { empty: true };
  let str;
  if (typeof x === 'number') {
    if (!Number.isFinite(x)) return { bad: true };
    str = String(x);
  } else if (typeof x === 'string') {
    str = x.trim().replace(/,/g, '');
    if (str === '') return { empty: true };
  } else {
    return { bad: true };
  }
  const m = /^([+-]?)(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(str);
  if (!m) return { bad: true };
  const intPart = m[2] || '';
  const fracPart = m[3] || '';
  if (intPart === '' && fracPart === '') return { bad: true };
  const exp = m[4] ? parseInt(m[4], 10) : 0;
  if (exp > 30) return { big: true };
  let digits = (intPart + fracPart).replace(/^0+(?=\d)/, '');
  let n = BigInt(digits === '' ? '0' : digits);
  let s = fracPart.length - exp;
  if (s < 0) { n = n * pow10(-s); s = 0; }
  if (s > 400) { n = 0n; s = 0; } // 小到無意義
  return { ok: true, n, s, neg: m[1] === '-' && n !== 0n };
}

// 四捨五入（half-up），num / den 皆為非負 BigInt，den > 0
function divRound(num, den) {
  return (2n * num + den) / (2n * den);
}

function toSafeNumber(b) {
  return Number(b); // 呼叫前已確認 <= MAX_CENTS（< 2^53）
}

// ───────────────────────── 商品歸類 ─────────────────────────

function normProducts(products) {
  const out = [];
  const seen = new Set();
  for (const p of (Array.isArray(products) ? products : [])) {
    if (typeof p !== 'string') continue;
    const name = p.trim();
    if (!name || seen.has(name)) continue;
    seen.add(name);
    out.push(name);
  }
  return out;
}

function classifyProducts(products, productClasses) {
  const pcs = (productClasses && typeof productClasses === 'object') ? productClasses : {};
  const names = normProducts(products);
  const keySet = new Set();
  const unclassified = [];
  const warnings = [];
  let needsConsultantCost = false;
  for (const name of names) {
    const ent = has(pcs, name) ? pcs[name] : null;
    const valid = ent && typeof ent === 'object' && typeof ent.cls === 'string' && CLASS_KEYS.indexOf(ent.cls) >= 0;
    if (!valid) {
      keySet.add('other');
      unclassified.push(name);
      warnings.push('商品『' + name + '』尚未歸類，已視為『其他』');
      needsConsultantCost = true; // 未歸類 → costBySales 視為 false → 需要顧問
      continue;
    }
    keySet.add(ent.cls);
    if (ent.costBySales !== true) needsConsultantCost = true;
  }
  const classKeys = CLASS_KEYS.filter(k => keySet.has(k)); // 以固定順序輸出，結果穩定
  return { classKeys, unclassified, needsConsultantCost, warnings };
}

function resolveRow(classKeys) {
  const set = new Set();
  let unknown = false;
  for (const k of (Array.isArray(classKeys) ? classKeys : [])) {
    if (CLASS_KEYS.indexOf(k) >= 0) set.add(k); else unknown = true;
  }
  if (set.size === 0 || unknown || set.has('other')) return ROW_BY_KEY.other;
  if (set.size === 1) return ROW_BY_KEY[Array.from(set)[0]];
  if (set.size === 2 && set.has('consult') && set.has('software')) return ROW_BY_KEY.consult_software;
  if (set.size === 2 && set.has('consult') && set.has('hardware')) return ROW_BY_KEY.consult_hardware;
  return ROW_BY_KEY.other;
}

// ───────────────────────── 財務計算 ─────────────────────────

function fail(code, message) { return { ok: false, code, message }; }

// 每列數量：空 / 0 → 1（與 lib/quoteExcel.js 匯出的 `num(it.qty) || 1` 一致）
function parseQty(x) {
  const d = parseDec(x);
  if (d.empty) return { ok: true, n: 1n, s: 0 };
  if (d.big) return { big: true };
  if (d.bad || d.neg) return { bad: true };
  if (d.n === 0n) return { ok: true, n: 1n, s: 0 };
  return d;
}

function computeFinancials(q) {
  if (!q || typeof q !== 'object') return fail('BAD_INPUT', '報價單資料無效');
  const items = Array.isArray(q.items) ? q.items : [];
  // 新式成本（q.costLines 存在）：成本取自成本明細（含差旅／交際費／印花稅），items[].cost 不再參與；沒有 costLines＝舊式，下面的流程與以前完全相同
  const newStyle = CL.hasCostLines(q);
  let subtotal = 0n;   // 分
  let cost = 0n;       // 分
  const missingCostRows = [];
  let rowNo = 0;   // 「第 N 列」只數一般品項（標題／小計列不編號），與表單、Excel 的項目編號一致
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (QI.isNonItemRow(it)) continue;   // 分組標題／小計列：不計價、不計成本
    rowNo++;
    if (!it || typeof it !== 'object') return fail('BAD_ITEM', '第 ' + rowNo + ' 列品項資料無效');
    const qty = parseQty(it.qty);
    if (qty.big) return fail('TOO_LARGE', '第 ' + rowNo + ' 列數量過大');
    if (qty.bad) return fail('BAD_ITEM', '第 ' + rowNo + ' 列數量不是有效的非負數字');
    const price = parseDec(it.unitPrice);
    if (price.big) return fail('TOO_LARGE', '第 ' + rowNo + ' 列單價過大');
    if (price.bad || price.neg) return fail('BAD_ITEM', '第 ' + rowNo + ' 列單價不是有效的非負數字');
    const c = newStyle ? { empty: true } : parseDec(it.cost);
    if (c.big) return fail('TOO_LARGE', '第 ' + rowNo + ' 列成本過大');
    if (c.bad || c.neg) return fail('BAD_ITEM', '第 ' + rowNo + ' 列成本不是有效的非負數字');

    const pn = price.empty ? 0n : price.n, ps = price.empty ? 0 : price.s;
    const cn = c.empty ? 0n : c.n, cs = c.empty ? 0 : c.s;
    const lineCents = divRound(qty.n * pn * 100n, pow10(qty.s + ps));
    const lineCost = divRound(qty.n * cn * 100n, pow10(qty.s + cs));
    if (lineCents > MAX_CENTS || lineCost > MAX_CENTS) return fail('TOO_LARGE', '第 ' + rowNo + ' 列金額過大');
    subtotal += lineCents;
    cost += lineCost;
    if (subtotal > MAX_CENTS || cost > MAX_CENTS) return fail('TOO_LARGE', '金額合計過大');
    if (!newStyle && pn > 0n && !(cn > 0n)) missingCostRows.push(rowNo); // 有價列成本需 > 0（新式成本不逐列對應品項）
  }

  // 折扣
  const dtRaw = q.discountType;
  const dt = (dtRaw === undefined || dtRaw === null || dtRaw === '') ? 'none' : dtRaw;
  let revenue = subtotal;
  let discountApplied = 'none';
  if (dt === 'percent' || dt === 'amount') {
    const v = parseDec(q.discountValue);
    if (v.big || v.bad || v.neg) return fail('BAD_DISCOUNT', '折扣值無效');
    if (!v.empty && v.n !== 0n) {
      if (dt === 'percent') {
        // 0 < v < 100（v 為實收百分比，90＝九折）
        if (!(v.n < 100n * pow10(v.s))) return fail('BAD_DISCOUNT', '折扣百分比必須大於 0 且小於 100');
        revenue = divRound(subtotal * v.n, 100n * pow10(v.s));
        discountApplied = 'percent';
      } else {
        // 議價總額（元）：0 < v < 小計（元）
        if (!(v.n * 100n < subtotal * pow10(v.s))) return fail('BAD_DISCOUNT', '議價金額必須大於 0 且小於折扣前小計');
        revenue = divRound(v.n * 100n, pow10(v.s));
        discountApplied = 'amount';
      }
    }
  } else if (dt !== 'none') {
    return fail('BAD_DISCOUNT', '折扣類型無效');
  }
  if (revenue <= 0n) return fail('ZERO_REVENUE', '折扣後未稅金額為 0，無法計算毛利');
  if (revenue > MAX_CENTS) return fail('TOO_LARGE', '金額過大');

  if (newStyle) {
    // 成本 ＝ Σ 每列取整後的分（與 items 同一套取整）＋ 印花稅（整數元＝營收×0.001 四捨五入）。
    // 簽核毛利含全部成本明細（含差旅／交際／印花稅），不含 contingency（風險預留只進 PNL 毛利）。
    const tot = CL.totalsByCat(q, toSafeNumber(revenue));
    if (!tot.ok) return fail(tot.code, tot.message);
    return {
      ok: true,
      subtotalCents: toSafeNumber(subtotal),
      revenueCents: toSafeNumber(revenue),
      costCents: tot.total,
      gpCents: toSafeNumber(revenue) - tot.total,
      discountApplied,
      costComplete: tot.nonStampCount >= 1 && tot.nonStamp > 0,   // 非印花稅列至少 1 列且總成本 > 0（與路由的 costComplete 共用 CL.costLinesComplete 的定義）
      missingCostRows: [],
      costBreakdownCents: { consult: tot.consult, software: tot.software, hw: tot.hw, travel: tot.travel, other: tot.other },
    };
  }

  return {
    ok: true,
    subtotalCents: toSafeNumber(subtotal),
    revenueCents: toSafeNumber(revenue),
    costCents: toSafeNumber(cost),
    gpCents: toSafeNumber(revenue - cost),
    discountApplied,
    costComplete: missingCostRows.length === 0,
    missingCostRows,
  };
}

// 毛利率文字：小數兩位、向 0 截斷（不四捨五入）；revenue<=0 或輸入無效 → null
function marginText(gpCents, revenueCents) {
  if (!Number.isSafeInteger(gpCents) || !Number.isSafeInteger(revenueCents) || revenueCents <= 0) return null;
  const gp = BigInt(gpCents);
  const rev = BigInt(revenueCents);
  const neg = gp < 0n;
  const abs = neg ? -gp : gp;
  const hundredths = (abs * 10000n) / rev; // 百分比 ×100，向 0 截斷
  const whole = hundredths / 100n;
  const frac = hundredths % 100n;
  const body = whole.toString() + '.' + (frac < 10n ? '0' : '') + frac.toString();
  return (neg && hundredths !== 0n ? '-' : '') + body;
}

function fmtYuan(cents) {
  const c = BigInt(cents);
  const whole = (c / 100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const frac = c % 100n;
  return frac === 0n ? whole : whole + '.' + (frac < 10n ? '0' : '') + frac.toString();
}

function requiredPath(rowKey, gpCents, revenueCents) {
  if (!Number.isSafeInteger(gpCents) || !Number.isSafeInteger(revenueCents)) throw new TypeError('gpCents / revenueCents 必須是整數');
  if (revenueCents <= 0) throw new RangeError('revenueCents 必須大於 0');
  const row = has(ROW_BY_KEY, rowKey) ? ROW_BY_KEY[rowKey] : ROW_BY_KEY.other;
  const gp100 = BigInt(gpCents) * 100n;
  const rev = BigInt(revenueCents);
  const mt = marginText(gpCents, revenueCents);
  const reasons = [];

  // 毛利層級
  let marginLevel;
  if (row.fixedLevel) {
    marginLevel = row.fixedLevel;
    reasons.push('「' + row.label + '」不在核決表內，固定需總經理核准（不看毛利，毛利率 ' + mt + '%）');
  } else if (row.l1 !== null && gp100 >= BigInt(row.l1) * rev) {
    marginLevel = 1;
    reasons.push('毛利率 ' + mt + '% 達「' + row.label + '」一級主管核決門檻 ' + row.l1 + '%');
  } else if (gp100 >= BigInt(row.l2) * rev) {
    marginLevel = 2;
    if (row.l1 !== null) reasons.push('毛利率 ' + mt + '% 低於「' + row.label + '」一級主管門檻 ' + row.l1 + '%，需總經理核准');
    else reasons.push('「' + row.label + '」主管不可終局，毛利率 ' + mt + '% 達總經理門檻 ' + row.l2 + '%，需總經理核准');
  } else {
    marginLevel = 3;
    reasons.push('毛利率 ' + mt + '% 低於「' + row.label + '」總經理門檻 ' + row.l2 + '%，需董事長核准');
  }

  // 金額層級（嚴格大於）
  let amountLevel = 1;
  let board = false;
  if (rev > AMOUNT_CENTS.chairman) {
    amountLevel = 3;
    reasons.push('折扣後未稅金額 ' + fmtYuan(revenueCents) + ' 元超過 1,000 萬，需董事長核准');
  }
  if (rev > AMOUNT_CENTS.board) {
    board = true;
    reasons.push('折扣後未稅金額 ' + fmtYuan(revenueCents) + ' 元超過 5,000 萬，需董事會決議（由管理部秘書代為核准）');
  }

  const level = Math.max(marginLevel, amountLevel);
  const tiers = (board ? PATHS.board : PATHS[level]).slice();
  return { level, board, tiers, reasons };
}

function buildDerived(q, productClasses) {
  if (!q || typeof q !== 'object') return null;
  const fin = computeFinancials(q);
  if (!fin.ok) return null;
  if (!fin.costComplete) return null; // 成本缺漏不可默默當 100% 毛利
  const cl = classifyProducts(q.products, productClasses);
  const row = resolveRow(cl.classKeys);
  const path = requiredPath(row.key, fin.gpCents, fin.revenueCents);
  // 建議性警告（不擋送簽）：新式成本、且有「有價品項沒有成本列涵蓋」才多放這個欄位，讓核准面板的簽核人看得到。
  // 舊式單與沒有未涵蓋品項的新式單完全不多任何鍵 → derived 輸出與以前位元級相同。
  const costWarnings = CL.hasCostLines(q) ? CL.costWarnings(q) : [];
  return {
    rowKey: row.key,
    rowLabel: row.label,
    level: path.level,
    board: path.board,
    marginText: marginText(fin.gpCents, fin.revenueCents),
    revenueCents: fin.revenueCents,
    costCents: fin.costCents,
    gpCents: fin.gpCents,
    tiers: path.tiers,
    reasons: path.reasons,
    warnings: cl.warnings,
    ...(costWarnings.length ? { costWarnings } : {}),
  };
}

// ───────────────────────── 雜湊 / 結構簽章 ─────────────────────────

function normText(x) {
  if (x === null || x === undefined) return '';
  return String(x).replace(/\r\n?/g, '\n').trim();
}

// 數字正規化：空 → null；有效數字 → Number；其他垃圾值 → 帶前綴字串（不同垃圾不同雜湊）
function normNum(x) {
  if (x === null || x === undefined) return null;
  if (typeof x === 'number') {
    if (!Number.isFinite(x)) return 'x:' + String(x);
    return x === 0 ? 0 : x;
  }
  if (typeof x === 'string') {
    const d = parseDec(x);
    if (d.empty) return null;
    if (d.ok) { const n = Number(x.trim().replace(/,/g, '')); return n === 0 ? 0 : n; }
    return 's:' + x.trim();
  }
  return 'x:' + String(x);
}

function canon(v) {
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  if (v && typeof v === 'object') {
    return '{' + Object.keys(v).sort().map(k => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
  }
  return JSON.stringify(v === undefined ? null : v);
}

function effectiveDiscount(q) {
  const t = (q.discountType === undefined || q.discountType === null || q.discountType === '') ? 'none' : String(q.discountType);
  if (t !== 'percent' && t !== 'amount') return { discountType: t === 'none' ? 'none' : t, discountValue: 0 };
  const v = normNum(q.discountValue);
  if (v === null || v === 0) return { discountType: 'none', discountValue: 0 }; // 值為 0/空＝無折扣
  return { discountType: t, discountValue: v };
}

function contentHash(q) {
  q = (q && typeof q === 'object') ? q : {};
  const disc = effectiveDiscount(q);
  const rawItems = Array.isArray(q.items) ? q.items : [];
  // 標題／小計列（kind）：只有「有這種列」時才在 payload 加 kind 與列順序 → 舊單（全是一般品項）雜湊與以前位元級相同
  const hasKindRows = rawItems.some(QI.isNonItemRow);
  const items = rawItems.map(it => {
    const kindRow = QI.isNonItemRow(it);
    it = (it && typeof it === 'object') ? it : {};
    return {
      lid: normText(it.lid),
      desc: normText(it.desc),
      unit: normText(it.unit),
      qty: normNum(it.qty),
      unitPrice: normNum(it.unitPrice),
      cost: normNum(it.cost),
      ...(kindRow ? { kind: it.kind } : {}),
    };
  }).map(o => ({ o, k: canon(o) }))
    .sort((a, b) => (a.o.lid < b.o.lid ? -1 : a.o.lid > b.o.lid ? 1 : (a.k < b.k ? -1 : a.k > b.k ? 1 : 0)))
    .map(x => x.o);
  const costRows = CL.hasCostLines(q) ? q.costLines.map(l => {
    l = (l && typeof l === 'object') ? l : {};
    return {
      lid: normText(l.lid), cat: normText(l.cat), desc: normText(l.desc), vendor: normText(l.vendor), note: normText(l.note),
      unit: normText(l.unit), qty: normNum(l.qty), auto: normText(l.auto),
      ...(l.auto === 'stamp' ? {} : { unitCost: normNum(l.unitCost) }),
    };
  }).map(o => ({ o, k: canon(o) }))
    .sort((a, b) => (a.o.lid < b.o.lid ? -1 : a.o.lid > b.o.lid ? 1 : (a.k < b.k ? -1 : a.k > b.k ? 1 : 0)))
    .map(x => x.o) : null;
  // v2：除價格/數量/成本/折扣/商品外，也涵蓋「會印在給客戶的報價單上」的文字欄位
  // （備註印成 Remarks 第 7 條，常放付款/保固等商業條款；聯絡人、地址、電話同理；報價期限 validUntil 印成 Remarks 第 4 條）。
  // projectNo（專案號碼）已不再印在報價單上、表單也移除了，但仍留在 payload：改雜湊的任何欄位都會讓已核准的單失效。
  // 不納入：mobile、contactId、quoteDate（不印在客戶單上；匯出日期取當天）。
  const payload = {
    v: 2,
    company: normText(q.company),
    projectName: normText(q.projectName),
    projectNo: normText(q.projectNo),
    contactName: normText(q.contactName),
    address: normText(q.address),
    phone: normText(q.phone),
    note: normText(q.note),
    // 報價期限（印在 Remarks 第 4 條，客戶單上的承諾）。沒有值就不放進 payload：功能上線前建立的舊單雜湊維持不變，已核准的不會被作廢
    ...(normText(q.validUntil) ? { validUntil: normText(q.validUntil) } : {}),
    // 付款方式（Remarks 第 3 條）與追加條款（第 7 條起）：印在客戶單上的商業承諾，核准後修改要重簽。
    // 同樣「有值才放進 payload」：舊單沒有這兩欄、雜湊維持不變。
    ...(q.payment && typeof q.payment === 'object' && Array.isArray(q.payment.items)
      ? { payment: { net: normNum(q.payment.net), items: q.payment.items.map(it => ({ label: normText(it && it.label), pct: normNum(it && it.pct) })) } } : {}),
    ...(Array.isArray(q.extraClauses) && q.extraClauses.some(s => normText(s)) ? { extraClauses: q.extraClauses.map(normText).filter(Boolean) } : {}),
    products: normProducts(q.products).sort(),
    discountType: disc.discountType,
    discountValue: disc.discountValue,
    items,
    ...(hasKindRows ? { order: rawItems.map(it => normText(it && it.lid)) } : {}),
    // 成本明細：只有「新式成本」（q.costLines 欄位存在）才放進 payload → 沒有 costLines 的舊單雜湊與以前位元級相同。
    // 印花稅列（auto:'stamp'）不放 unitCost：金額由營收算出、不是輸入值，營收變動本來就會讓雜湊變。順序無財務意義，與 items 一樣依 lid 排序。
    ...(costRows ? { costLines: costRows } : {}),
  };
  return crypto.createHash('sha256').update(canon(payload), 'utf8').digest('hex');
}

function escSig(s) {
  return normText(s).replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\n/g, '\\n');
}

// 品項結構簽章：lid|desc|qty|unit|是否有價（不含 cost、unitPrice 的實際金額）；與列順序無關。
// 「是否有價」要進簽章：贈品列（單價 0、成本免填）被改成有價列時，必須讓顧問重填成本。
function lineStructureSig(items) {
  return (Array.isArray(items) ? items : []).filter(QI.isItemRow).map(it => {
    it = (it && typeof it === 'object') ? it : {};
    const qty = normNum(it.qty);
    const price = normNum(it.unitPrice);
    const priced = (typeof price === 'number' && price > 0) ? '1' : '0';
    return [escSig(it.lid), escSig(it.desc), qty === null ? '' : String(qty), escSig(it.unit), priced].join('|');
  }).sort().join('\n');
}

// ───────────────────────── 併發保護簽章 ─────────────────────────
// 顧問對話框／業務毛利頁籤載入時拿到這兩個值、存檔時帶回，伺服器比對「目前」的值：不同＝這個分頁已經過期（別的視窗改過），
// 回 409 STALE_ITEMS／STALE_COSTS，避免過期分頁整批覆寫成本明細、或在沒看過新品項的情況下按「完成」。
// 兩者都只是 sha256 摘要（取前 16 碼，混入單據 id 與用途標籤），回應裡不含任何明文金額，只用來比對「內容有沒有變」，不是存取控制：
//   · itemsSig 的輸入是 lineStructureSig（品項 lid／說明／數量／單位，以及「各品項是否有價」一個位元；不含單價金額）。
//     對手若已知品項內容，理論上可以窮舉重算而推得這個位元；「有價／無價」不是金額，這個程度的資訊量可以接受。
//   · costLinesSig 的輸入含各列成本明細全部欄位；拿到簽章的人本來就看得到成本明細（顧問與業務可互看成本，規格 v1.1），沒有額外洩漏。
// 只由 serialize 給「看得到成本明細／可填成本」的人。

function sigDigest(tag, q, text) {
  return crypto.createHash('sha256').update(tag + '|' + String((q && q.id) || '') + '|' + text, 'utf8').digest('hex').slice(0, 16);
}

/** 品項結構簽章的摘要（＝lineStructureSig(items) 的雜湊）：品項增刪、改說明／數量／單位、有價↔無價會變；改單價金額（維持有價）、改成本不會變 */
function itemsSig(q) {
  return sigDigest('items', q, lineStructureSig(q && q.items));
}

/**
 * 目前儲存的成本明細內容摘要：沒有 costLines（舊式單）回 ''；有（含空陣列）回 16 碼摘要。
 * 涵蓋每列 lid,cat,desc,vendor,note,unit,qty,unitCost,auto,forLid 與列順序（任何欄位或順序改變都會變）。
 */
function costLinesSig(q) {
  if (!CL.hasCostLines(q)) return '';
  const rows = q.costLines.map((l) => {
    l = (l && typeof l === 'object') ? l : {};
    return [l.lid, l.cat, l.desc, l.vendor, l.note, l.unit, l.qty, l.unitCost, l.auto, l.forLid].map((v) => (v === undefined ? null : v));
  });
  return sigDigest('costLines', q, JSON.stringify(rows));
}

// ───────────────────────── 送簽檢查 ─────────────────────────

function validateForSubmit(q, productClasses) {
  const errors = [];
  if (!q || typeof q !== 'object') {
    return { ok: false, errors: [{ code: 'BAD_INPUT', message: '報價單資料無效' }], warnings: [], derived: null };
  }
  const cl = classifyProducts(q.products, productClasses);
  const warnings = cl.warnings.slice();

  if (normProducts(q.products).length < 1) errors.push({ code: 'NO_PRODUCTS', message: '請至少勾選一個商品' });
  const items = QI.itemRows(q.items);   // 只有分組標題／小計列不算有品項
  if (items.length < 1) errors.push({ code: 'NO_ITEMS', message: '請至少填寫一列品項' });

  const fin = computeFinancials(q);
  if (!fin.ok) {
    // 沒有品項時 computeFinancials 會回 ZERO_REVENUE，已有 NO_ITEMS 就不重複
    if (!(items.length < 1 && fin.code === 'ZERO_REVENUE')) errors.push({ code: fin.code, message: fin.message });
  }

  if (cl.needsConsultantCost) {
    const st = q.costFlow && q.costFlow.state;
    if (st !== 'filled') errors.push({ code: 'COST_NOT_FILLED', message: '顧問尚未填寫成本' });
  }

  if (fin.ok && !fin.costComplete) {
    const rows = fin.missingCostRows;
    // 新式成本（成本明細）不逐列對應品項：訊息不列「第 N 列」
    const e = CL.hasCostLines(q)
      ? { code: 'MISSING_COST', message: '尚未填寫成本明細', rows: [] }
      : { code: 'MISSING_COST', message: '第 ' + rows.join('、') + ' 列有單價但成本未填或為 0（贈品列請把單價設為 0）', rows: rows.slice() };
    errors.push(e);
  }

  const ok = errors.length === 0;
  return { ok, errors, warnings, derived: ok ? buildDerived(q, productClasses) : null };
}

module.exports = {
  RULES_VERSION, TIERS, CLASS_LABELS, CLASS_KEYS, ROWS, ROW_BY_KEY, AMOUNT, MAX_CENTS_NUM: Number(MAX_CENTS),
  classifyProducts, resolveRow, computeFinancials, requiredPath,
  marginText, contentHash, lineStructureSig, itemsSig, costLinesSig, validateForSubmit, buildDerived,
};
