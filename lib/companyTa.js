'use strict';
// 企業主檔「篩選 TA」：公司條件（產業／區域／資本額／BU／客戶狀態／責任業務／服務業務…）→ 聯絡人名單（EDM 用）
// 純函式：不碰資料庫、不寫入；可視範圍由呼叫端（server.js）先算好再傳進來，這裡只負責「在可視範圍內篩選」。
// 語意：同一欄位內多選＝OR；不同欄位之間＝AND。

const NONE = '__none__';
const STATUSES = ['won', 'open', 'noOpp', 'newImport'];   // 已成交／已開商機（未成交）／尚無商機／新匯入（尚無真實名片）
const MAX_LIST = 200;
const MAX_STR = 60;
const EMAIL_RE = /^[^\s@,;<>()"]+@[^\s@,;<>()"]+\.[^\s@,;<>()"]+$/;

const str = v => String(v == null ? '' : v).trim();
const strList = v => (Array.isArray(v) ? v : []).map(str).filter(Boolean).map(s => s.slice(0, MAX_STR)).slice(0, MAX_LIST);
const numOrNull = v => { if (v === '' || v == null) return null; const n = Number(v); return Number.isFinite(n) && n >= 0 ? n : null; };
const tri = v => (v === 'yes' || v === 'no' ? v : '');
const dateStr = v => { const s = str(v); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; };

function normalizeFilters(raw) {
  const r = raw && typeof raw === 'object' ? raw : {};
  const cap = r.capNone === 'include' || r.capNone === 'only' ? r.capNone : '';
  return {
    keyword: str(r.keyword).slice(0, 80),
    industries: strList(r.industries),          // 可含 NONE（未分類）
    regions: strList(r.regions),                // 可含 NONE（未填）
    capMinWan: numOrNull(r.capMinWan),          // 資本額下限（萬元）
    capMaxWan: numOrNull(r.capMaxWan),
    capNone: cap,                               // ''＝有設資本額區間時排除未填；include＝連未填也納入；only＝只要未填
    bus: strList(r.bus),
    statuses: strList(r.statuses).filter(s => STATUSES.includes(s)),
    responsible: strList(r.responsible),        // 帳號；可含 NONE（未指定）
    services: strList(r.services),              // 帳號（名片擁有者）
    hasTaxId: tri(r.hasTaxId),
    gcis: tri(r.gcis),
    foreign: tri(r.foreign),
    createdFrom: dateStr(r.createdFrom),
    createdTo: dateStr(r.createdTo),
  };
}

function normalizeOpts(raw) {
  const r = raw && typeof raw === 'object' ? raw : {};
  const dflt = (v, d) => (v === undefined ? d : !!v);
  return {
    requireEmail: dflt(r.requireEmail, true),
    activeOnly: dflt(r.activeOnly, true),
    dedupeEmail: dflt(r.dedupeEmail, true),
    onePerCompany: dflt(r.onePerCompany, false),
    titleKeyword: str(r.titleKeyword).slice(0, 80),
  };
}

function summarizeFilters(f) {
  const p = [];
  if (f.keyword) p.push('關鍵字');
  if (f.industries.length) p.push('產業×' + f.industries.length);
  if (f.regions.length) p.push('區域×' + f.regions.length);
  if (f.capMinWan != null || f.capMaxWan != null) p.push('資本額 ' + (f.capMinWan != null ? f.capMinWan : '') + '~' + (f.capMaxWan != null ? f.capMaxWan : '') + ' 萬');
  if (f.capNone) p.push('資本額未填:' + f.capNone);
  if (f.bus.length) p.push('BU=' + f.bus.join('/'));
  if (f.statuses.length) p.push('狀態=' + f.statuses.join('/'));
  if (f.responsible.length) p.push('責任業務×' + f.responsible.length);
  if (f.services.length) p.push('服務業務×' + f.services.length);
  if (f.hasTaxId) p.push('統編:' + f.hasTaxId);
  if (f.gcis) p.push('GCIS:' + f.gcis);
  if (f.foreign) p.push('境外:' + f.foreign);
  if (f.createdFrom || f.createdTo) p.push('建立 ' + f.createdFrom + '~' + f.createdTo);
  return p.length ? p.join('；') : '(無公司條件)';
}
function summarizeOpts(o) {
  return ['必有Email:' + (o.requireEmail ? '是' : '否'), '僅在職:' + (o.activeOnly ? '是' : '否'), 'Email去重:' + (o.dedupeEmail ? '是' : '否'),
    '每家一位:' + (o.onePerCompany ? '是' : '否'), o.titleKeyword ? '職稱關鍵字' : ''].filter(Boolean).join('、');
}

const capitalOf = m => { const n = Number(m && m.capital); return Number.isFinite(n) && n > 0 ? n : null; };
const empStatus = c => c.employmentStatus || (c.isResigned ? 'resigned' : 'active');

// 公司層級條件（f 已 normalize）；d 為該公司的推導資料 { bu[], services[], status[] }
function companyMatches(m, d, f) {
  if (f.keyword) {
    const k = f.keyword.toLowerCase();
    if (!(String(m.name || '').toLowerCase().includes(k) || String(m.taxId || '').includes(f.keyword))) return false;
  }
  if (f.industries.length) {
    const ind = str(m.industry);
    if (!(ind ? f.industries.includes(ind) : f.industries.includes(NONE))) return false;
  }
  if (f.regions.length) {
    const rg = str(m.region);
    if (!(rg ? f.regions.includes(rg) : f.regions.includes(NONE))) return false;
  }
  const cap = capitalOf(m);
  const hasRange = f.capMinWan != null || f.capMaxWan != null;
  if (f.capNone === 'only') { if (cap != null) return false; }
  else if (hasRange || f.capNone) {
    if (cap == null) { if (f.capNone !== 'include') return false; }
    else {
      if (f.capMinWan != null && cap < f.capMinWan * 10000) return false;
      if (f.capMaxWan != null && cap > f.capMaxWan * 10000) return false;
    }
  }
  if (f.bus.length && !d.bu.some(b => f.bus.includes(b))) return false;
  if (f.statuses.length && !f.statuses.some(s => d.status.includes(s))) return false;
  if (f.responsible.length) {
    const rs = Array.isArray(m.responsibleSales) ? m.responsibleSales : [];
    if (!(rs.length ? rs.some(u => f.responsible.includes(u)) : f.responsible.includes(NONE))) return false;
  }
  if (f.services.length && !d.services.some(u => f.services.includes(u))) return false;
  if (f.hasTaxId) { const has = /^\d{8}$/.test(str(m.taxId)); if ((f.hasTaxId === 'yes') !== has) return false; }
  if (f.gcis) { if ((f.gcis === 'yes') !== !!m.gcisEnriched) return false; }
  if (f.foreign) { if ((f.foreign === 'yes') !== !!m.foreign) return false; }
  if (f.createdFrom || f.createdTo) {
    const day = str(m.createdAt).slice(0, 10);
    if (!day) return false;
    if (f.createdFrom && day < f.createdFrom) return false;
    if (f.createdTo && day > f.createdTo) return false;
  }
  return true;
}

// 主流程。入參皆為「呼叫者可視範圍」內的資料：
//   companies：可見的企業主檔；contacts：可見名片（已排除刪除、客戶池）；opps：可見商機
//   buUnion(ownerUsernames, buImport)：回傳 BU 陣列（呼叫端包 companyBuUnion）
// 回傳 { companyIds, rows:[{ company, contact }], stats:{ companies, contacts, withEmail, ... } }
function run({ companies, contacts, opps, buUnion, filters, opts }) {
  const f = normalizeFilters(filters), o = normalizeOpts(opts);
  const byCo = new Map();
  (contacts || []).forEach(c => { if (c.companyId) { if (!byCo.has(c.companyId)) byCo.set(c.companyId, []); byCo.get(c.companyId).push(c); } });
  const oppsByContact = new Map(), oppsByName = new Map();
  (opps || []).forEach(op => {
    if (op.contactId) { if (!oppsByContact.has(op.contactId)) oppsByContact.set(op.contactId, []); oppsByContact.get(op.contactId).push(op); }
    const nm = str(op.company); if (nm) { if (!oppsByName.has(nm)) oppsByName.set(nm, []); oppsByName.get(nm).push(op); }
  });

  const matched = [];
  (companies || []).forEach(m => {
    const cs = byCo.get(m.id) || [];
    const real = cs.filter(c => !c.isPlaceholder);
    const services = [...new Set(real.map(c => c.owner).filter(Boolean))];
    const myOpps = new Set();
    real.forEach(c => (oppsByContact.get(c.id) || []).forEach(op => myOpps.add(op)));
    (oppsByName.get(str(m.name)) || []).forEach(op => myOpps.add(op));
    const status = [];
    if (!real.length) status.push('newImport');
    const hasWon = [...myOpps].some(op => op.stage === 'Won');
    const hasOpen = [...myOpps].some(op => op.stage !== 'Won');
    if (hasWon) status.push('won');
    if (hasOpen) status.push('open');
    if (!myOpps.size) status.push('noOpp');
    const d = { bu: buUnion ? buUnion(services, m.buImport) : [], services, status };
    if (companyMatches(m, d, f)) matched.push({ m, d, cs: real });
  });

  const kw = o.titleKeyword.toLowerCase();
  const seenEmail = new Set();
  const rows = [];
  let withEmail = 0, candidates = 0;
  matched.forEach(({ m, cs }) => {
    let list = cs.slice();
    if (o.activeOnly) list = list.filter(c => empStatus(c) !== 'resigned');
    if (kw) list = list.filter(c => String(c.title || '').toLowerCase().includes(kw));
    candidates += list.length;
    list = list.filter(c => { const has = EMAIL_RE.test(str(c.email)); if (has) withEmail++; return o.requireEmail ? has : true; });
    if (o.onePerCompany && list.length > 1) {
      const primary = list.find(c => c.isPrimary);
      list = [primary || list.slice().sort((a, b) => String(a.createdAt || '').localeCompare(String(b.createdAt || '')))[0]];
    }
    list.forEach(c => {
      if (o.dedupeEmail) {
        const e = str(c.email).toLowerCase();
        if (e) { if (seenEmail.has(e)) return; seenEmail.add(e); }
      }
      rows.push({ company: m, contact: c });
    });
  });
  return {
    companyIds: matched.map(x => x.m.id),
    matched: matched.map(x => ({ company: x.m, bu: x.d.bu })),
    rows,
    stats: { companies: matched.length, contactsBeforeEmail: candidates, withEmail, contacts: rows.length },
    filters: f, opts: o,
  };
}

module.exports = { NONE, STATUSES, EMAIL_RE, normalizeFilters, normalizeOpts, summarizeFilters, summarizeOpts, companyMatches, run };
