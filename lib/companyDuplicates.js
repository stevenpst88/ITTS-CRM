'use strict';
// 企業主檔「疑似重複」偵測：只讀，不合併、不刪除。
// 主檔比對鍵是「統編」或「正規化公司名」，所以以下情況會產生兩筆主檔：
//   ① 一筆有統編、一筆沒統編但名稱相同（業務沒填統編建名片，行銷之後匯入帶統編）
//   ② 都沒統編，但名稱寫法不同（「甲科技」vs「甲科技股份有限公司」）
//   ③ 同一個統編卻有兩筆主檔（資料異常）
//   ④ 名稱相同（去公司後綴後）但統編不同：可能是不同法人或分公司，列為「低信心」供人工判斷
// 以並查集把「同統編」或「同寬鬆名稱」的主檔連成一組；組內 ≥2 筆才列出。

const TAX_RE = /^\d{8}$/;
const str = v => String(v == null ? '' : v).trim();

function run({ companies, contacts, loose, ignores }) {
  const list = (companies || []).filter(m => m && !m.deleted && m.id);
  const total = {}, real = {};
  (contacts || []).forEach(c => {
    if (c.deleted || !c.companyId) return;
    total[c.companyId] = (total[c.companyId] || 0) + 1;
    if (!c.isPlaceholder) real[c.companyId] = (real[c.companyId] || 0) + 1;
  });

  const parent = new Map(list.map(m => [m.id, m.id]));
  const find = x => { while (parent.get(x) !== x) { parent.set(x, parent.get(parent.get(x))); x = parent.get(x); } return x; };
  const union = (a, b) => { const ra = find(a), rb = find(b); if (ra !== rb) parent.set(ra, rb); };

  const byName = new Map(), byTax = new Map();
  list.forEach(m => {
    const k = loose ? str(loose(m.name)) : str(m.name).toLowerCase();
    if (k.length >= 2) { if (!byName.has(k)) byName.set(k, []); byName.get(k).push(m); }
    const t = str(m.taxId);
    if (TAX_RE.test(t)) { if (!byTax.has(t)) byTax.set(t, []); byTax.get(t).push(m); }
  });
  [...byName.values(), ...byTax.values()].forEach(ms => { for (let i = 1; i < ms.length; i++) union(ms[0].id, ms[i].id); });

  const comps = new Map();
  list.forEach(m => { const r = find(m.id); if (!comps.has(r)) comps.set(r, []); comps.get(r).push(m); });

  const groups = [];
  comps.forEach(ms => {
    if (ms.length < 2) return;
    const ids = ms.map(m => m.id).sort();
    const key = ids.join('|');
    if (ignores && ignores[key]) return;
    const taxes = [...new Set(ms.map(m => str(m.taxId)).filter(t => TAX_RE.test(t)))];
    const noTax = ms.filter(m => !TAX_RE.test(str(m.taxId)));
    const reasons = [];
    const sameTax = [...byTax.values()].some(g => g.length > 1 && g.every(x => ms.includes(x)));
    if (sameTax) reasons.push('sameTaxId');
    if (taxes.length >= 1 && noTax.length >= 1) reasons.push('taxVsNoTax');
    if (taxes.length === 0) reasons.push('nameVariant');
    if (taxes.length >= 2 && !sameTax) reasons.push('diffTaxId');
    const level = reasons.includes('sameTaxId') || reasons.includes('taxVsNoTax') || reasons.includes('nameVariant') ? 'high' : 'low';
    groups.push({
      key, level, reasons,
      members: ms.map(m => ({
        id: m.id, name: m.name || '', taxId: m.taxId || '', industry: m.industry || '', region: m.region || '',
        capital: m.capital == null ? null : m.capital, createdAt: m.createdAt || '', source: m.source || '',
        contactCount: total[m.id] || 0, realContactCount: real[m.id] || 0,
      })).sort((a, b) => (b.realContactCount - a.realContactCount) || (b.contactCount - a.contactCount) || a.name.localeCompare(b.name, 'zh-TW')),
    });
  });
  groups.sort((a, b) => (a.level === b.level ? 0 : a.level === 'high' ? -1 : 1)
    || (b.members.reduce((s, m) => s + m.contactCount, 0) - a.members.reduce((s, m) => s + m.contactCount, 0)));
  return { groups, summary: { total: groups.length, high: groups.filter(g => g.level === 'high').length, low: groups.filter(g => g.level === 'low').length } };
}

// 「標記為非重複」的鍵：組內主檔 id 排序後串接；組成員一變（有新主檔加入）鍵就變，自動重新出現
function keyOf(ids) { return (Array.isArray(ids) ? ids : []).map(str).filter(Boolean).sort().join('|'); }

module.exports = { run, keyOf };
