// 銷售商品目錄（Admin 可維護）— 純函式，方便單元測試
const QT = require('../_client/quote-terms.js');   // 商品類型（none／license／ma）的白名單解析；與報價單服務期間共用同一份規則
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
// opts.terms=true（MA_TRACKING=1 時由呼叫端傳入）另帶 PRODUCT_TERMS: {name:{termKind,startsTask}}（只含 termKind 不是 none 的啟用商品；同名取最後一個）。
// 沒傳 opts（預設）→ 輸出與加入商品類型之前逐位元相同。
function deriveDropdowns(catalog, opts) {
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
  if (opts && opts.terms) {
    const PRODUCT_TERMS = {};
    for (const bu of Object.keys(bus)) {
      for (const g of (Array.isArray(bus[bu]) ? bus[bu] : [])) {
        if (!g || g.enabled === false) continue;
        for (const it of (g.items || [])) {
          if (!it || it.enabled === false || !it.name) continue;
          const k = QT.parseCatalogKind(it.termKind);
          if (k && k !== 'none') PRODUCT_TERMS[it.name] = { termKind: k, startsTask: it.startsTask === true };
        }
      }
    }
    return { OPP_PRODUCTS, PRODUCT_SUBSELS, PRODUCT_TERMS };
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
// 商機的 dealId 現為 9 碼數字字串（SAP Deal ID）；舊資料可能是任意文字（grandfather，見 server.js checkOppDealId）。
const DealIdFormat = require('../_client/deal-id-format.js');
const DEAL_ID_MAX = 64;   // 舊資料上限（相容保留；新格式見 DealIdFormat）
const DEAL_URL_MAX = 500;
// 控制字元（C0/DEL/C1）與換行分隔符 U+2028/2029
const CTRL_RE = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/g;      // 只用於 replace
const CTRL_TEST = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/;     // 只用於 test（無 g 旗標，避免 lastIndex 狀態）

/** 商機 SAP Deal ID 正規化＋格式驗證（9 碼數字；單一來源 _client/deal-id-format.js，前端共用同一份）。
 *  回 { ok, value } 或 { ok:false, code, error }；null/undefined 視為空字串；number 先轉字串。
 *  code：DEAL_ID_INVALID（非字串/數字型別）／DEAL_ID_FORMAT（不是 9 碼數字）。Hotline 轉商機等其他入口請直接呼叫這個函式。 */
function normalizeDealId(v) { return DealIdFormat.normalize(v); }

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
    if (ov.req !== nv.req) out.push(`${name}: 需要 SAP Deal ID ${ov.req ? '是' : '否'}→${nv.req ? '是' : '否'}`);
    if (ov.url !== nv.url) out.push(`${name}: 註冊網址 ${ov.url ? '已設定' : '無'}→${nv.url ? '已設定' : '無'}`);
  }
  return out;
}

// ───────────────── 商品類型（License／MA 服務期間；MA_TRACKING=1 才啟用）─────────────────
// 商品目錄項目的兩個設定（皆選填，既有商品預設不需要）：
//   termKind   : 'none'|'license'|'ma'（預設 none＝不存欄位）— 報價單選到這類商品時品項要填服務期間（前端預選，業務可改）
//   startsTask : boolean（預設 false＝不存欄位；僅 termKind≠none 可為 true）— 成交後啟動任務卡。P1 只存不用（P2 才使用）
// 與 Deal ID 設定一樣「預設值不存欄位」：沒用到的目錄與加入此功能之前位元級相同。
function mapCatalogItems(bus, fn) {
  const out = {};
  for (const bu of Object.keys(bus || {})) {
    const groups = Array.isArray(bus[bu]) ? bus[bu] : [];
    out[bu] = groups.map(g => {
      if (!g || typeof g !== 'object' || !Array.isArray(g.items)) return g;
      return Object.assign({}, g, { items: g.items.map(it => (it && typeof it === 'object' ? fn(it, bu, g) : it)) });
    });
  }
  return out;
}

/** 存目錄前的商品類型正規化：回 { bus, errors:[{bu,group,name,error}] }；termKind 不認得 → errors（呼叫端回 400 TERM_KIND_INVALID）；不修改傳入物件 */
function normalizeCatalogTermSettings(bus) {
  const errors = [];
  const out = mapCatalogItems(bus, (it, bu, g) => {
    const n = Object.assign({}, it);
    const k = it.termKind === undefined ? 'none' : QT.parseCatalogKind(it.termKind);
    if (k === null) { errors.push({ bu, group: g.group, name: it.name, error: '商品類型不正確（只能是 none、license 或 ma）' }); delete n.termKind; delete n.startsTask; return n; }
    if (k !== 'none') n.termKind = k; else delete n.termKind;
    const st = k !== 'none' && (it.startsTask === true || it.startsTask === 'true' || it.startsTask === 1);
    if (st) n.startsTask = true; else delete n.startsTask;
    return n;
  });
  return { bus: out, errors };
}

/** MA_TRACKING 關閉時的輸出／儲存：把 termKind／startsTask 整個拿掉（沒有這兩欄的項目原物件不動，所以舊目錄逐位元相同） */
function stripCatalogTermSettings(bus) {
  return mapCatalogItems(bus, (it) => {
    if (it.termKind === undefined && it.startsTask === undefined) return it;
    const n = Object.assign({}, it); delete n.termKind; delete n.startsTask; return n;
  });
}

/** 兩份目錄間商品類型設定的差異（供稽核）。回 ['名稱: 類型 一般→License', '名稱: 成交後啟動任務卡 否→是', ...] */
function diffTermSettings(oldCatalog, newBus) {
  const flat = (bus) => {
    const m = new Map();
    for (const bu of Object.keys(bus || {})) {
      for (const g of (Array.isArray(bus[bu]) ? bus[bu] : [])) {
        for (const it of ((g && g.items) || [])) {
          if (it && typeof it.name === 'string') m.set(bu + '/' + it.name, { kind: QT.parseCatalogKind(it.termKind) || 'none', task: it.startsTask === true });
        }
      }
    }
    return m;
  };
  const a = flat((oldCatalog && oldCatalog.bus) || SEED_PRODUCT_CATALOG.bus), b = flat(newBus);
  const label = { none: '一般', license: 'License', ma: 'MA' };
  const out = [];
  for (const [k, nv] of b) {
    const ov = a.get(k) || { kind: 'none', task: false };
    const name = k.split('/').slice(1).join('/');
    if (ov.kind !== nv.kind) out.push(`${name}: 類型 ${label[ov.kind]}→${label[nv.kind]}`);
    if (ov.task !== nv.task) out.push(`${name}: 成交後啟動任務卡 ${ov.task ? '是' : '否'}→${nv.task ? '是' : '否'}`);
  }
  return out;
}

/**
 * 依商品文字查「商品類型」規則（P2 成交納管用）。比對邏輯與 resolveDealIdRule 完全相同（名稱完全相符，否則取「名稱＋空白」前綴最長者；同名多項時任一項有類型即採用，優先 ma）。
 * 回 { termKind:'none'|'license'|'ma', startsTask:boolean, name:string }；商品不在目錄、或沒有類型 → termKind:'none'、startsTask:false。
 */
function resolveTermRule(catalog, product) {
  const none = { termKind: 'none', startsTask: false, name: '' };
  const p = typeof product === 'string' ? product.trim() : '';
  if (!p) return none;
  const bus = (catalog && catalog.bus) || SEED_PRODUCT_CATALOG.bus;
  let best = null;
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
  const typed = best.items.filter(it => it.termKind === 'ma' || it.termKind === 'license');
  if (!typed.length) return { termKind: 'none', startsTask: false, name: best.name };
  const pick = typed.find(it => it.termKind === 'ma') || typed[0];
  return { termKind: pick.termKind, startsTask: typed.some(it => it.termKind === pick.termKind && it.startsTask === true), name: best.name };
}

/** 商品目錄 API 輸出：MA_TRACKING 開啟 → 加 maTracking:true（後台據此顯示「類型」欄）；關閉 → 去掉商品類型欄位（輸出與沒有此功能時相同） */
function catalogForOutput(catalog, on) {
  if (on) return Object.assign({}, catalog, { maTracking: true });
  const bus = catalog && catalog.bus;
  if (!bus || typeof bus !== 'object') return catalog;
  const has = Object.keys(bus).some(bu => (Array.isArray(bus[bu]) ? bus[bu] : []).some(g => g && Array.isArray(g.items) && g.items.some(it => it && (it.termKind !== undefined || it.startsTask !== undefined))));
  return has ? Object.assign({}, catalog, { bus: stripCatalogTermSettings(bus) }) : catalog;
}

module.exports = {
  SEED_PRODUCT_CATALOG, deriveDropdowns, applyCatalogRenames,
  DEAL_ID_MAX, normalizeDealId, validateDealUrl, normalizeCatalogDealSettings, resolveDealIdRule, diffDealSettings,
  normalizeCatalogTermSettings, stripCatalogTermSettings, diffTermSettings, catalogForOutput, resolveTermRule,
};
