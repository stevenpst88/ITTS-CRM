// 銷售商品目錄（Admin 可維護）— 純函式，方便單元測試
// 資料模型：{ bus: { ERP:[{group,enabled,items:[{name,enabled,subs:[]}]}], ITS:[...], MDM:[], CRM:[] } }

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

module.exports = { SEED_PRODUCT_CATALOG, deriveDropdowns, applyCatalogRenames };
