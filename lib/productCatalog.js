// 銷售商品目錄（Admin 可維護）— 純函式，方便單元測試
// 資料模型：{ bus: { ERP:[{group,enabled,items:[{name,enabled,subs:[]}]}], ITS:[...], MDM:[], CRM:[], 資安:[] } }

// 預設種子＝原本寫死在 _client/app.js 的 OPP_PRODUCTS + PRODUCT_SUBSELS（群組標籤保留原樣，零視覺變化）
const SEED_PRODUCT_CATALOG = {
  bus: {
    ERP: [
      { group: '一般商品', enabled: true, items: [
        { name: 'SAP Public Cloud License', enabled: true, subs: [] },
        { name: 'SAP Private Cloud License', enabled: true, subs: [] },
        { name: 'ERP 顧問導入專案 PE', enabled: true, subs: [] },
        { name: 'ERP 顧問導入專案 PCE', enabled: true, subs: [] },
        { name: 'SAP CCFLEX', enabled: true, subs: ['C&S', 'A&O'] },
        { name: '其他', enabled: true, subs: [] },
      ] },
      { group: '─── License MA ───', enabled: true, items: [
        { name: 'SAP PCOE License MA', enabled: true, subs: [] },
        { name: 'SAP PCE License MA', enabled: true, subs: [] },
        { name: 'SAP PE License MA', enabled: true, subs: [] },
      ] },
      { group: '─── Service MA ───', enabled: true, items: [
        { name: 'Service MA（Basis）', enabled: true, subs: [] },
        { name: 'Service MA（AP）', enabled: true, subs: [] },
        { name: 'Service MA（Basis & AP）', enabled: true, subs: [] },
      ] },
      { group: '─── SAP 延伸解決方案 ───', enabled: true, items: [
        { name: 'SAP CRM', enabled: true, subs: [] },
        { name: 'SAP Analytics Cloud（SAC）', enabled: true, subs: [] },
        { name: 'SAP Customer Data Platform（CDP）', enabled: true, subs: [] },
        { name: 'SAP Engagement Cloud（Emarsys）', enabled: true, subs: [] },
      ] },
    ],
    ITS: [
      { group: 'ITS 產品', enabled: true, items: [
        { name: 'MES', enabled: true, subs: [] },
        { name: 'WMS', enabled: true, subs: [] },
        { name: 'ESG', enabled: true, subs: [] },
        { name: 'DMS', enabled: true, subs: [] },
      ] },
    ],
    MDM: [],
    CRM: [],
    資安: [],
  },
};

// 由目錄衍生前端 3 個下拉需要的舊格式（只含 enabled 群組/產品）
// 回 { OPP_PRODUCTS: {BU:{group:[name]}}, PRODUCT_SUBSELS: {name:[sub]} }
function deriveDropdowns(catalog) {
  const bus = (catalog && catalog.bus) || {};
  const OPP_PRODUCTS = {};
  const PRODUCT_SUBSELS = {};
  for (const bu of Object.keys(bus)) {
    const groups = Array.isArray(bus[bu]) ? bus[bu] : [];
    const out = {};
    for (const g of groups) {
      if (!g || g.enabled === false) continue;
      const names = (g.items || []).filter(it => it && it.enabled !== false).map(it => it.name);
      if (names.length) out[g.group] = names;
      for (const it of (g.items || [])) {
        if (it && it.enabled !== false && Array.isArray(it.subs) && it.subs.length) {
          PRODUCT_SUBSELS[it.name] = it.subs.slice();
        }
      }
    }
    OPP_PRODUCTS[bu] = out;
  }
  return { OPP_PRODUCTS, PRODUCT_SUBSELS };
}

// 連動改名：把 opp.product / contract.product 中「等於舊名」或「以『舊名 』為前綴（含明細組合值）」者改為新名
// renames: [{ old, new }]；opts.dryRun=true 時只計數不寫入
// 回 { updatedOpps, updatedContracts }
function applyCatalogRenames(data, renames, opts) {
  const dryRun = !!(opts && opts.dryRun);
  const map = (Array.isArray(renames) ? renames : [])
    .filter(r => r && r.old && r.new && r.old !== r.new);
  let updatedOpps = 0, updatedContracts = 0;
  if (!map.length) return { updatedOpps, updatedContracts };

  const renameVal = (val) => {
    if (typeof val !== 'string' || !val) return null;
    for (const r of map) {
      if (val === r.old) return r.new;                                  // 完全相同
      if (val.startsWith(r.old + ' ')) return r.new + val.slice(r.old.length); // 明細組合值（保留後綴含空格）
    }
    return null;
  };

  for (const o of (data.opportunities || [])) {
    const nv = renameVal(o.product);
    if (nv !== null) { if (!dryRun) o.product = nv; updatedOpps++; }
  }
  for (const c of (data.contracts || [])) {
    const nv = renameVal(c.product);
    if (nv !== null) { if (!dryRun) c.product = nv; updatedContracts++; }
  }
  return { updatedOpps, updatedContracts };
}

// ───────────────── Deal ID（商機註冊代碼）規則 ─────────────────
// 商品目錄項目的兩個設定（皆選填，既有商品預設不需要）：
//   requiresDealId  : boolean  — 銷售此商品的商機必填 Deal ID
//   dealRegisterUrl : string   — 外部註冊網站（僅 http/https），前端用新分頁開啟
// 商機的 dealId 是純文字（不限格式）：去控制字元、去前後空白、長度 ≤ DEAL_ID_MAX。
const DEAL_ID_MAX = 64;
const DEAL_URL_MAX = 500;
// 控制字元（C0/DEL/C1）與換行分隔符 U+2028/2029
const CTRL_RE = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/g;      // 只用於 replace
const CTRL_TEST = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/;     // 只用於 test（無 g 旗標，避免 lastIndex 狀態）

/** 商機 Deal ID 正規化。回 { ok, value } 或 { ok:false, code, error }；null/undefined 視為空字串。 */
function normalizeDealId(v) {
  if (v === null || v === undefined) return { ok: true, value: '' };
  if (typeof v !== 'string' && typeof v !== 'number') return { ok: false, code: 'DEAL_ID_INVALID', error: 'Deal ID 格式不正確（需為文字）' };
  const s = String(v).replace(CTRL_RE, '').trim();
  if (s.length > DEAL_ID_MAX) return { ok: false, code: 'DEAL_ID_TOO_LONG', error: `Deal ID 長度不可超過 ${DEAL_ID_MAX} 字` };
  return { ok: true, value: s };
}

/** 註冊網址驗證：空字串合法（＝未設定）；否則必須是 http/https 的絕對網址（用 new URL 解析，存解析後的 href）。 */
function validateDealUrl(v) {
  if (v === null || v === undefined) return { ok: true, value: '' };
  if (typeof v !== 'string') return { ok: false, error: 'Deal 註冊網址格式不正確（需為文字）' };
  const s = v.trim();
  if (!s) return { ok: true, value: '' };
  if (s.length > DEAL_URL_MAX) return { ok: false, error: `Deal 註冊網址過長（上限 ${DEAL_URL_MAX} 字）` };
  if (CTRL_TEST.test(s) || /\s/.test(s)) return { ok: false, error: 'Deal 註冊網址不可含空白或控制字元' };
  if (!/^https?:\/\//i.test(s)) return { ok: false, error: 'Deal 註冊網址必須以 http:// 或 https:// 開頭' };
  let u;
  try { u = new URL(s); } catch { return { ok: false, error: 'Deal 註冊網址格式不正確' }; }
  if (u.protocol !== 'http:' && u.protocol !== 'https:') return { ok: false, error: 'Deal 註冊網址必須以 http:// 或 https:// 開頭' };
  if (!u.hostname) return { ok: false, error: 'Deal 註冊網址缺少網站名稱' };
  if (u.username || u.password) return { ok: false, error: 'Deal 註冊網址不可內含帳號密碼' };
  return { ok: true, value: u.href };
}

/**
 * 儲存商品目錄前的 Deal ID 設定正規化：只動每個商品的 requiresDealId / dealRegisterUrl，其餘欄位原樣保留。
 * 回 { bus, errors:[{ bu, group, name, error }] }；有 errors 時呼叫端應回 400、不寫入。
 * 不修改傳入物件（回傳新的 bus）。
 */
function normalizeCatalogDealSettings(bus) {
  const errors = [];
  const out = {};
  for (const bu of Object.keys(bus || {})) {
    const groups = Array.isArray(bus[bu]) ? bus[bu] : [];
    out[bu] = groups.map(g => {
      if (!g || typeof g !== 'object') return g;
      const items = Array.isArray(g.items) ? g.items.map(it => {
        if (!it || typeof it !== 'object') return it;
        const n = Object.assign({}, it);
        const req = it.requiresDealId === true || it.requiresDealId === 'true' || it.requiresDealId === 1;
        if (req) n.requiresDealId = true; else delete n.requiresDealId;
        const uv = validateDealUrl(it.dealRegisterUrl);
        if (!uv.ok) errors.push({ bu, group: g.group, name: it.name, error: uv.error });
        if (uv.ok && uv.value) n.dealRegisterUrl = uv.value; else delete n.dealRegisterUrl;
        return n;
      }) : g.items;
      return Object.assign({}, g, { items });
    });
  }
  return { bus: out, errors };
}

/**
 * 依商品文字查 Deal ID 規則。product 可為「名稱」或「名稱 + 空白 + 明細」（如 'SAP CCFLEX C&S'）。
 * 比對：先完全相符（區分大小寫、只去前後空白），否則取「名稱 + ' '」前綴最長者。商品不在目錄 → 不需要。
 * 同名商品出現在多個群組：任一個要求即視為要求；網址取第一個有設定者。
 * 回 { requires:boolean, url:string, name:string }
 */
function resolveDealIdRule(catalog, product) {
  const none = { requires: false, url: '', name: '' };
  const p = typeof product === 'string' ? product.trim() : '';
  if (!p) return none;
  const bus = (catalog && catalog.bus) || SEED_PRODUCT_CATALOG.bus;
  let best = null;   // { name, items:[] }
  for (const bu of Object.keys(bus)) {
    const groups = Array.isArray(bus[bu]) ? bus[bu] : [];
    for (const g of groups) {
      for (const it of ((g && g.items) || [])) {
        if (!it || typeof it.name !== 'string' || !it.name) continue;
        const exact = p === it.name;
        if (!exact && !p.startsWith(it.name + ' ')) continue;
        const rank = exact ? Infinity : it.name.length;
        if (!best || rank > best.rank) best = { rank, name: it.name, items: [it] };
        else if (rank === best.rank) best.items.push(it);
      }
    }
  }
  if (!best) return none;
  const reqItems = best.items.filter(it => it.requiresDealId === true);
  if (!reqItems.length) return { requires: false, url: '', name: best.name };
  const withUrl = reqItems.find(it => typeof it.dealRegisterUrl === 'string' && it.dealRegisterUrl);
  return { requires: true, url: withUrl ? withUrl.dealRegisterUrl : '', name: best.name };
}

/** 兩份目錄間 Deal ID 設定的差異（供稽核）。回 ['名稱: 需要 Deal ID 否→是', ...] */
function diffDealSettings(oldCatalog, newBus) {
  const flat = (bus) => {
    const m = new Map();
    for (const bu of Object.keys(bus || {})) {
      for (const g of (Array.isArray(bus[bu]) ? bus[bu] : [])) {
        for (const it of ((g && g.items) || [])) {
          if (it && typeof it.name === 'string') m.set(bu + '/' + it.name, { req: it.requiresDealId === true, url: it.dealRegisterUrl || '' });
        }
      }
    }
    return m;
  };
  const a = flat((oldCatalog && oldCatalog.bus) || SEED_PRODUCT_CATALOG.bus), b = flat(newBus);
  const out = [];
  for (const [k, nv] of b) {
    const ov = a.get(k) || { req: false, url: '' };
    const name = k.split('/').slice(1).join('/');
    if (ov.req !== nv.req) out.push(`${name}: 需要 Deal ID ${ov.req ? '是' : '否'}→${nv.req ? '是' : '否'}`);
    if (ov.url !== nv.url) out.push(`${name}: 註冊網址 ${ov.url ? '已設定' : '無'}→${nv.url ? '已設定' : '無'}`);
  }
  return out;
}

module.exports = {
  SEED_PRODUCT_CATALOG, deriveDropdowns, applyCatalogRenames,
  DEAL_ID_MAX, normalizeDealId, validateDealUrl, normalizeCatalogDealSettings, resolveDealIdRule, diffDealSettings,
};
