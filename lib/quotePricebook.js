/**
 * 報價牌價簿（Pricebook）：管理員在後台維護的「顧問角色 × 人天牌價／人天成本」清單。
 *
 * 純函式模組（不碰 db／express），路由在 lib/quoteRoutes.js：
 *   GET  /api/quote-pricebook        能用報價單功能的人：只回啟用的項目（含牌價與成本）
 *   GET  /api/admin/quote-pricebook  管理員：全部項目（含停用）＋ updatedAt
 *   PUT  /api/admin/quote-pricebook  管理員：整份取代（驗證＋樂觀並行 updatedAt，不符→409 STALE_PRICEBOOK）
 *
 * 資料存在 data.pricebook = { items:[...], updatedAt, updatedBy }（獨立命名空間，不混進其他 namespace）。
 * 項目：{ id, bu, name, unit:'人天', price, cost, active }；陣列是扁平的，同一個 BU 內的順序＝陣列順序（四個 BU 分頁的順序固定 ERP、ITS、MDM、CRM）。
 *
 * BU（事業單位）：與 CRM 其他處相同的四個值與順序（lib/quoteRoutes.js ALL_BUS／server.js VALID_BUS）。每個項目有 bu；
 *  · 相容：舊資料（上線時還沒有 bu 欄位）與舊用戶端（送來的項目沒有 bu）一律視為 ERP——讀取與寫入都會補上，GET 一定回 bu；
 *    PUT 時項目沒帶 bu 但 id 對得上既有項目 → 沿用既有項目的 bu（舊的快取頁面存檔不會把別的 BU 的項目洗成 ERP），否則 ERP；
 *  · 名稱唯一性是「每個 BU 內」（nameKey 不變）；同一個角色名稱可以在不同 BU 各有一筆、費率不同；
 *  · 上限：每個 BU 60 筆（共 240 筆）。
 *
 * 設計取捨（業主決定）：
 *  · 牌價簿只是「建議值」：業務選取時把 desc／unit／qty／unitPrice 複製進報價單品項，之後與牌價簿脫鉤（沒有參照、沒有 provenance 欄位）；
 *    伺服器不強制報價單價格等於牌價；牌價簿改了也不影響任何已存的報價單。
 *  · v1 只有單位「人天」、只有分類「顧問」；委外費用依供應商報價，在填成本時手動輸入，不在這裡維護。
 *  · 沒有生效日／版本。
 */
'use strict';

const BUS = Object.freeze(['ERP', 'ITS', 'MDM', 'CRM']);   // 與 lib/quoteRoutes.js ALL_BUS 同值同序
const DEFAULT_BU = 'ERP';
const MAX_ITEMS = 60;                                   // 每個 BU 的上限
const MAX_TOTAL = MAX_ITEMS * BUS.length;               // 全部 BU 合計上限（240）
const MAX_NAME = 40;
const MAX_MONEY = 1e9;
const UNIT = '人天';
const ID_RE = /^[A-Za-z0-9_-]{1,40}$/;
const CTRL_RE = /[\u0000-\u001f\u007f]/;

const LIMITS = Object.freeze({ MAX_ITEMS, MAX_TOTAL, MAX_NAME, MAX_MONEY, UNIT, BUS });

const err = (code, message, index, field) => ({ ok: false, error: Object.assign({ code, message }, index === undefined ? {} : { index }, field ? { field } : {}) });

/** 金額：number 或「純數字字串」（不接受空字串、科學記號、負號、NaN、Infinity、布林、null）；0..MAX_MONEY；四捨五入到小數 2 位。失敗回 null */
function parseMoney(v) {
  let n;
  if (typeof v === 'number') n = v;
  else if (typeof v === 'string' && /^\s*\d+(\.\d+)?\s*$/.test(v)) n = Number(v);
  else return null;
  if (!Number.isFinite(n) || n < 0 || n > MAX_MONEY) return null;
  return Math.round(n * 100) / 100;
}

/** 項目的 BU：合法值原樣，缺漏／不合法（舊資料）一律 ERP。只給「讀取已存資料」用；PUT 的輸入驗證走 normalizePricebook 的嚴格檢查 */
const buOf = (it) => (it && typeof it.bu === 'string' && BUS.includes(it.bu)) ? it.bu : DEFAULT_BU;

const nameKey = (s) => String(s).normalize('NFKC').replace(/\s+/g, ' ').trim().toLowerCase();

/**
 * 驗證並清洗整份清單（整份取代）。
 * opts.genId()：產生新 id（必須提供；測試可注入決定性版本）。id 缺漏或格式不合 → 產生新的；已給且合法 → 保留（id 穩定）。
 * opts.existing：目前已存的項目（可選）。項目沒帶 bu 時，id 對得上既有項目就沿用它的 bu（見檔頭「相容」），否則 ERP。
 * 回傳 { ok:true, items } 或 { ok:false, error:{code,message,index?,field?} }；code 一律 BAD_PRICEBOOK。index 是整份清單裡的 0 起算位置。
 * 名稱只當純文字存放（不做 HTML 轉義／過濾；顯示端負責轉義），但拒絕控制字元。
 */
function normalizePricebook(raw, opts) {
  const genId = opts && opts.genId;
  if (typeof genId !== 'function') throw new Error('normalizePricebook: opts.genId required');
  if (!Array.isArray(raw)) return err('BAD_PRICEBOOK', '項目清單格式不正確');
  if (raw.length > MAX_TOTAL) return err('BAD_PRICEBOOK', `項目最多 ${MAX_TOTAL} 筆（目前 ${raw.length} 筆）`);
  const prevBu = new Map();
  if (opts && Array.isArray(opts.existing)) for (const e of opts.existing) if (e && typeof e.id === 'string') prevBu.set(e.id, buOf(e));
  const items = [];
  const seenNames = new Map();   // bu\u0001nameKey → 該 BU 內的第幾列（1 起算）
  const seenIds = new Set();
  const perBu = {};
  for (const b of BUS) perBu[b] = 0;
  for (let i = 0; i < raw.length; i++) {
    const r = raw[i];
    if (!r || typeof r !== 'object' || Array.isArray(r)) return err('BAD_PRICEBOOK', `第 ${i + 1} 列格式不正確`, i);
    // BU：缺漏（undefined／null／空字串）→ 沿用既有（id 對得上）或 ERP；有值就必須是四個之一
    let bu;
    if (r.bu === undefined || r.bu === null || r.bu === '') bu = (typeof r.id === 'string' && prevBu.get(r.id)) || DEFAULT_BU;
    else if (typeof r.bu === 'string' && BUS.includes(r.bu)) bu = r.bu;
    else return err('BAD_PRICEBOOK', `第 ${i + 1} 列：BU 必須是 ${BUS.join('、')} 其中之一`, i, 'bu');
    const pos = ++perBu[bu];   // 這一列在該 BU 內的第幾列
    const at = `[${bu}] 第 ${pos} 列`;
    if (pos > MAX_ITEMS) return err('BAD_PRICEBOOK', `[${bu}] 項目最多 ${MAX_ITEMS} 筆（目前超過）`);
    if (typeof r.name !== 'string') return err('BAD_PRICEBOOK', `${at}：項目名稱必須是文字`, i, 'name');
    const name = r.name.trim();
    if (!name) return err('BAD_PRICEBOOK', `${at}：項目名稱不可空白`, i, 'name');
    if (name.length > MAX_NAME) return err('BAD_PRICEBOOK', `${at}：項目名稱不可超過 ${MAX_NAME} 字`, i, 'name');
    if (CTRL_RE.test(name)) return err('BAD_PRICEBOOK', `${at}：項目名稱不可含控制字元`, i, 'name');
    const key = bu + '\u0001' + nameKey(name);
    if (seenNames.has(key)) return err('BAD_PRICEBOOK', `${at}：項目名稱「${name}」與同 BU 第 ${seenNames.get(key)} 列重複（不分大小寫）`, i, 'name');
    seenNames.set(key, pos);
    const price = parseMoney(r.price);
    if (price === null) return err('BAD_PRICEBOOK', `${at}：牌價必須是 0～${MAX_MONEY} 的數字`, i, 'price');
    const cost = parseMoney(r.cost);
    if (cost === null) return err('BAD_PRICEBOOK', `${at}：成本必須是 0～${MAX_MONEY} 的數字`, i, 'cost');
    if (r.active !== undefined && typeof r.active !== 'boolean') return err('BAD_PRICEBOOK', `${at}：啟用狀態必須是 true／false`, i, 'active');
    let id = (typeof r.id === 'string' && ID_RE.test(r.id)) ? r.id : null;
    if (id && seenIds.has(id)) return err('BAD_PRICEBOOK', `${at}：id 重複`, i, 'id');
    items.push({ id, bu, name, unit: UNIT, price, cost, active: r.active !== false });
    if (id) seenIds.add(id);
  }
  // 新 id 最後再配（避免新 id 撞上後面列已指定的 id）
  for (const it of items) {
    if (it.id) continue;
    let id, guard = 0;
    do { id = String(genId()); if (!ID_RE.test(id)) id = 'pb' + id.replace(/[^A-Za-z0-9_-]/g, '').slice(0, 30); } while ((seenIds.has(id) || !id) && ++guard < 50);
    if (!id || seenIds.has(id)) return err('BAD_PRICEBOOK', 'id 產生失敗');
    it.id = id; seenIds.add(id);
  }
  return { ok: true, items };
}

/** 給業務端：只留啟用項目、只留 picker 需要的欄位（牌價＋成本，業主決定業務看得到成本）；bu 一定有值（缺漏＝ERP） */
function publicItems(items) {
  return (Array.isArray(items) ? items : []).filter((it) => it && it.active !== false)
    .map((it) => ({ id: it.id, bu: buOf(it), name: it.name, unit: UNIT, price: it.price, cost: it.cost }));
}

/** 管理員端：全部欄位；bu 一定有值（缺漏＝ERP） */
function adminItems(items) {
  return (Array.isArray(items) ? items : []).map((it) => ({ id: it.id, bu: buOf(it), name: it.name, unit: UNIT, price: it.price, cost: it.cost, active: it.active !== false }));
}

/** 依 bu 分組（保持各 BU 內的原順序）：{ ERP:[...], ITS:[...], MDM:[...], CRM:[...] }；缺 bu 的算 ERP */
function groupByBu(items) {
  const g = {};
  for (const b of BUS) g[b] = [];
  for (const it of (Array.isArray(items) ? items : [])) if (it) g[buOf(it)].push(it);
  return g;
}

/**
 * 新舊清單差異（依 id 對應；bu 缺漏視為 ERP）：{ added, removed, renamed, changed, toggled, moved, reordered, reorderedBus }
 * 每筆都帶 bu。moved＝同一個 id 換了 BU（UI 不提供，但 API 可以）。reordered 是「任一 BU 內順序有變」，reorderedBus 列出哪些 BU。
 */
function diffPricebook(prev, next) {
  const a = Array.isArray(prev) ? prev : [], b = Array.isArray(next) ? next : [];
  const byIdA = new Map(a.map((x) => [x.id, x])), byIdB = new Map(b.map((x) => [x.id, x]));
  const d = { added: [], removed: [], renamed: [], changed: [], toggled: [], moved: [], reordered: false, reorderedBus: [] };
  for (const n of b) if (!byIdA.has(n.id)) d.added.push({ bu: buOf(n), name: n.name, price: n.price, cost: n.cost, active: n.active !== false });
  for (const o of a) if (!byIdB.has(o.id)) d.removed.push({ bu: buOf(o), name: o.name, price: o.price, cost: o.cost });
  for (const n of b) {
    const o = byIdA.get(n.id);
    if (!o) continue;
    const bu = buOf(n);
    if (buOf(o) !== bu) d.moved.push({ name: n.name, from: buOf(o), to: bu });
    if (o.name !== n.name) d.renamed.push({ bu, from: o.name, to: n.name });
    if (o.price !== n.price) d.changed.push({ bu, name: n.name, field: 'price', from: o.price, to: n.price });
    if (o.cost !== n.cost) d.changed.push({ bu, name: n.name, field: 'cost', from: o.cost, to: n.cost });
    if ((o.active !== false) !== (n.active !== false)) d.toggled.push({ bu, name: n.name, active: n.active !== false });
  }
  for (const bu of BUS) {
    const commonA = a.filter((x) => buOf(x) === bu && byIdB.has(x.id) && buOf(byIdB.get(x.id)) === bu).map((x) => x.id);
    const commonB = b.filter((x) => buOf(x) === bu && byIdA.has(x.id) && buOf(byIdA.get(x.id)) === bu).map((x) => x.id);
    if (commonA.join('\u0001') !== commonB.join('\u0001')) d.reorderedBus.push(bu);
  }
  d.reordered = d.reorderedBus.length > 0;
  return d;
}

/** 稽核用一行摘要（每個異動都標 [BU]）：新增／移除／改名／牌價成本異動 old→new／啟停用／換 BU／順序調整；沒有任何差異回「無異動」 */
function summarizeDiff(d, maxLen) {
  const parts = [];
  const money = (n) => String(n);
  const tag = (x) => `[${x.bu || DEFAULT_BU}]`;
  if (d.added.length) parts.push('新增 ' + d.added.map((x) => `${tag(x)}「${x.name}」(牌價${money(x.price)}/成本${money(x.cost)}${x.active ? '' : '/停用'})`).join('、'));
  if (d.removed.length) parts.push('移除 ' + d.removed.map((x) => `${tag(x)}「${x.name}」(牌價${money(x.price)}/成本${money(x.cost)})`).join('、'));
  if (d.renamed.length) parts.push('改名 ' + d.renamed.map((x) => `${tag(x)}「${x.from}」→「${x.to}」`).join('、'));
  const lab = { price: '牌價', cost: '成本' };
  if (d.changed.length) parts.push(d.changed.map((x) => `${tag(x)}「${x.name}」${lab[x.field]} ${money(x.from)}→${money(x.to)}`).join('、'));
  if (d.toggled.length) parts.push(d.toggled.map((x) => `${tag(x)}「${x.name}」${x.active ? '啟用' : '停用'}`).join('、'));
  if (d.moved && d.moved.length) parts.push(d.moved.map((x) => `「${x.name}」換 BU [${x.from}]→[${x.to}]`).join('、'));
  if (d.reordered) parts.push('順序調整' + ((d.reorderedBus && d.reorderedBus.length) ? ' ' + d.reorderedBus.map((b) => `[${b}]`).join('') : ''));
  let s = parts.length ? parts.join('；') : '無異動';
  const lim = maxLen || 1500;
  if (s.length > lim) s = s.slice(0, lim - 1) + '…';
  return s;
}

module.exports = { BUS, DEFAULT_BU, MAX_ITEMS, MAX_TOTAL, MAX_NAME, MAX_MONEY, UNIT, LIMITS, parseMoney, nameKey, buOf, groupByBu, normalizePricebook, publicItems, adminItems, diffPricebook, summarizeDiff };
