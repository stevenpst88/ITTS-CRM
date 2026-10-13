#!/usr/bin/env node
/**
 * 唯讀報表：找出「公司欄純數字」或「公司名稱與該統編在企業主檔的名稱不符」的名片。
 * 不修改任何資料；只印 id / 公司 / 統編（不傾印其他個資）。
 *
 * 用法：
 *   node scripts/report-taxid-mismatch.js <資料檔路徑>          # 例如本機 data.json 的「複本」
 *   DATA_FILE=<路徑> node scripts/report-taxid-mismatch.js [--limit 20] [--json]
 *
 * 資料檔格式：data.json（含 contacts / companies 陣列）。建議先複製一份再讀。
 * 規則與 _client/taxid-guard.js 相同（單一來源）。
 *
 * 分類與修正建議：
 *   DIGITS_COMPANY    公司欄純數字 → 若主檔有該統編的官方名稱：把公司改成官方名稱；否則用統編查 GCIS 後人工修正
 *   MASTER_POISONED   主檔名稱本身是純數字/空（被污染）→ 先到企業主檔管理修正主檔名稱，再修名片
 *   NAME_MISMATCH     名片公司 ≠ 主檔官方名稱 → 確認哪邊對：名片錯就改成主檔名稱；主檔錯就改主檔
 */
'use strict';
const fs = require('fs');
const path = require('path');
const TG = require('../_client/taxid-guard.js');

const args = process.argv.slice(2);
const _li = args.indexOf('--limit');
const file = args.find((a, i) => !a.startsWith('--') && i !== _li + 1) || process.env.DATA_FILE;
if (!file) { console.error('請指定資料檔：node scripts/report-taxid-mismatch.js <data.json 複本>（或設 DATA_FILE）'); process.exit(2); }
const limIdx = args.indexOf('--limit');
const limit = limIdx >= 0 ? Math.max(1, parseInt(args[limIdx + 1], 10) || 20) : 20;
const asJson = args.includes('--json');

const data = JSON.parse(fs.readFileSync(path.resolve(file), 'utf8'));
const contacts = (data.contacts || []).filter(c => c && !c.deleted);
const companies = data.companies || [];

const rows = [];
for (const c of contacts) {
  const company = String(c.company || '').trim();
  const taxId = c.isForeign ? '' : String(c.taxId || '').trim();
  const master = taxId ? companies.find(m => m.matchKey === 'tax:' + taxId || m.taxId === taxId) : null;
  const masterName = master ? String(master.name || '').trim() : '';
  if (TG.isDigitsOnly(company)) {
    rows.push({ kind: 'DIGITS_COMPANY', id: c.id, company, taxId, masterName: TG.isUsableOfficialName(masterName) ? masterName : '',
      suggest: TG.isUsableOfficialName(masterName) ? '公司改為主檔官方名稱' : '以統編查 GCIS 後人工修正公司名稱' });
    continue;
  }
  if (taxId && master && !TG.isUsableOfficialName(masterName)) {
    rows.push({ kind: 'MASTER_POISONED', id: c.id, company, taxId, masterName, suggest: '先修正企業主檔名稱（目前為空或純數字）' });
    continue;
  }
  if (taxId && master && !TG.sameCompany(company, masterName)) {
    rows.push({ kind: 'NAME_MISMATCH', id: c.id, company, taxId, masterName, suggest: '確認何者正確：名片錯→改為主檔名稱；主檔錯→改主檔' });
  }
}
// 另外列出「主檔本身被污染」的清單（名片數可能為 0）
const poisonedMasters = companies.filter(m => m && m.taxId && !TG.isUsableOfficialName(m.name)).map(m => ({ taxId: m.taxId, name: m.name || '' }));

const by = {};
rows.forEach(r => { (by[r.kind] = by[r.kind] || []).push(r); });
const summary = { contactsScanned: contacts.length, total: rows.length, byKind: Object.fromEntries(Object.entries(by).map(([k, v]) => [k, v.length])), poisonedMasters: poisonedMasters.length };

if (asJson) {
  console.log(JSON.stringify({ summary, samples: Object.fromEntries(Object.entries(by).map(([k, v]) => [k, v.slice(0, limit)])), poisonedMasters: poisonedMasters.slice(0, limit) }, null, 2));
} else {
  console.log('掃描名片 ' + summary.contactsScanned + ' 張；有問題 ' + summary.total + ' 張；被污染的主檔 ' + summary.poisonedMasters + ' 筆');
  for (const [k, v] of Object.entries(by)) {
    console.log('\n[' + k + '] ' + v.length + ' 張（顯示前 ' + Math.min(limit, v.length) + '）');
    v.slice(0, limit).forEach(r => console.log('  id=' + r.id + '  公司=「' + r.company + '」 統編=' + r.taxId + (r.masterName ? '  主檔=「' + r.masterName + '」' : '') + '\n    建議：' + r.suggest));
  }
  if (poisonedMasters.length) {
    console.log('\n[被污染的企業主檔] ' + poisonedMasters.length + ' 筆（顯示前 ' + Math.min(limit, poisonedMasters.length) + '）');
    poisonedMasters.slice(0, limit).forEach(m => console.log('  統編=' + m.taxId + ' 主檔名稱=「' + m.name + '」'));
  }
  console.log('\n（唯讀報表，未修改任何資料）');
}
