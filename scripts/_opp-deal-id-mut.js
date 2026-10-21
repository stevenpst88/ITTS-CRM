'use strict';
/**
 * check-opp-deal-id.js 的變異測試（DEALID_MUT=1）：把後端（server.js／lib）以「載入時改寫原始碼」的方式、前端檔以 CDP 攔截替換的方式改壞，
 * 對應的檢查項目必須變紅。磁碟上的檔案一律不動。變異點必須在原始碼中唯一存在，否則該項目直接判失敗（防止原始碼改版後變異悄悄失效）。
 * 不單獨執行。
 */
const fs = require('fs');
const path = require('path');

// [名稱, 檔案（相對專案根）, 改寫清單 [{find, rep}], 預期變紅的檢查項目名稱（正規表示式）]
const SERVER_MUTANTS = [
  ['後端強制拿掉（POST 建立）', 'server.js', [{ find: "  if (_dl.error) return res.status(_dl.status).json({ error: _dl.error, code: _dl.code });\n  const opp = {", rep: '  const opp = {' }], /^A(1|2|3|4|11|16|19)\b/],
  ['後端強制拿掉（PUT 表單編輯）', 'server.js', [{ find: "    if (_dl.error) return res.status(_dl.status).json({ error: _dl.error, code: _dl.code });\n    if (req.body.dealId !== undefined) req.body.dealId = _dl.dealId;", rep: '    if (req.body.dealId !== undefined) req.body.dealId = _dl.dealId;' }], /^B(4|8|10|11)\b/],
  ['偏更新 {stage} 被誤擋（PUT 一律檢查）', 'server.js', [{ find: '  if (req.body.product !== undefined || req.body.dealId !== undefined) {', rep: '  if (true) {' }], /^B(1|2|3|16)\b/],
  ['URL 驗證放寬（接受 javascript: 等）', 'lib/productCatalog.js', [
    { find: "if (!/^https?:\\/\\//i.test(s)) return { ok: false, error: 'Deal 註冊網址必須以 http:// 或 https:// 開頭' };", rep: '' },
    { find: "if (u.protocol !== 'http:' && u.protocol !== 'https:') return { ok: false, error: 'Deal 註冊網址必須以 http:// 或 https:// 開頭' };", rep: '' }], /^D3\b/],
  ['前後空白不去除', 'lib/productCatalog.js', [{ find: "const s = String(v).replace(CTRL_RE, '').trim();", rep: "const s = String(v).replace(CTRL_RE, '');" }], /^(A2|A5|B5)\b/],
  ['控制字元不去除', 'lib/productCatalog.js', [{ find: "const s = String(v).replace(CTRL_RE, '').trim();", rep: 'const s = String(v).trim();' }], /^(A3|A6)\b/],
  ['長度上限（64）拿掉', 'lib/productCatalog.js', [{ find: 'if (s.length > DEAL_ID_MAX) return', rep: 'if (false) return' }], /^(A8|B9)\b/],
  ['明細組合值（名稱＋空白＋明細）比對拿掉', 'lib/productCatalog.js', [{ find: "!p.startsWith(it.name + ' ')", rep: 'true' }], /^A11\b/],
  ['商品目錄設定不限管理員', 'server.js', [{ find: "app.post('/api/admin/product-catalog', requireAdmin,", rep: "app.post('/api/admin/product-catalog', requireAuth," }], /^D2\b/],
  ['商品目錄儲存時略過網址驗證', 'server.js', [{ find: '    if (norm.errors.length) {', rep: '    if (false) {' }], /^D3\b/],
  ['Deal ID 變更不寫稽核', 'server.js', [{ find: '`${data.opportunities[idx].product}${stageDetail}${dealDetail}`', rep: '`${data.opportunities[idx].product}${stageDetail}`' }], /^(B6|B18)\b/],
  ['Excel 匯入也強制（會擋匯入）', 'server.js', [{ find: '&& productCatalog.resolveDealIdRule(data.productCatalog, opp.product).requires) missingDealId++;', rep: '&& productCatalog.resolveDealIdRule(data.productCatalog, opp.product).requires) { missingDealId++; return; }' }], /^E(1|2)\b/],
  ['匯出欄位順序被改（Deal ID 插到商機ID 前面）', 'server.js', [{ find: "'商機ID', 'Deal ID'\n  ];", rep: "'Deal ID', '商機ID'\n  ];" }], /^E(4|5)\b/],
];

// [名稱, 前端檔, 要改的原始碼, 改成, 要跑的前端區段, 預期變紅的檢查項目名稱]
const UI_MUTANTS = [
  ['前端必填檢查拿掉', 'app.js', "if (rule.requires && !v) msg = '請填寫 Deal ID（此商品必填）';", "if (false) msg = '請填寫 Deal ID（此商品必填）';", ['create', 'edit'], /^(UI12|UI13|UI16|UE3)\b/],
  ['看板卡片 Deal ID 未跳脫（innerHTML）', 'app.js', "di.textContent = 'Deal ID：' + o.dealId;", "di.innerHTML = 'Deal ID：' + o.dealId;", ['display'], /^(UD3|UD4)\b/],
  ['清單 Deal ID 未跳脫', 'app.js', '<div class="opp-dealid-line" title="Deal ID：${escapeHtml(o.dealId)}">Deal ID：${escapeHtml(o.dealId)}</div>', '<div class="opp-dealid-line" title="Deal ID：x">Deal ID：${o.dealId}</div>', ['display'], /^(UD6|UD7)\b/],
  ['註冊網址不檢查協定（接受 javascript:）', 'app.js', "return (x.protocol === 'http:' || x.protocol === 'https:') ? x.href : '';", 'return x.href;', ['display'], /^(UD11|UD12)\b/],
  ['欄位一律顯示（不隨商品隱藏）', 'app.js', 'const show = rule.requires || keep;', 'const show = true;', ['create'], /^(UI1|UI2|UI26)\b/],
  ['新增時不把 dealId 放進 payload', 'app.js', "const _dlN = dealIdValue('oppEdit'); if (_dlN) payload.dealId = _dlN;", '', ['create'], /^(UI23|UI24)\b/],
  ['編輯時不把 dealId 放進 payload', 'app.js', "const _dlE = dealIdValue('oppEdit'); if (_dlE !== undefined) payload.dealId = _dlE;", '', ['edit'], /^UE5\b/],
];

exports.run = async function (C, fns, runUiOnce) {
  console.log('[變異測試]（每個變異都必須讓對應檢查變紅）');
  const ROOT = C.ROOT;
  for (const [name, file, edits, expect] of SERVER_MUTANTS) {
    const src = fs.readFileSync(path.join(ROOT, file), 'utf8');
    const uniq = edits.every((e) => src.indexOf(e.find) >= 0 && src.indexOf(e.find, src.indexOf(e.find) + 1) < 0);
    C.ok(uniq, '變異點唯一存在：' + name);
    if (!uniq) continue;
    const snap = C.snap(); C.setQuiet(true);
    let err = '';
    try { await fns.groupApi({ file, edits }); } catch (e) { err = String((e && e.message) || e).slice(0, 200); }
    const nf = C.newFails(snap); C.setQuiet(false); C.restore(snap);
    const red = nf.some((n) => expect.test(n));
    C.ok(red, '變異必須使測試變紅：' + name, err || ('實際變紅的項目：' + nf.slice(0, 4).join(' | ')));
    console.log('    ' + (red ? '紅 ✓ ' : '未變紅 ✗ ') + name + '  → ' + nf.slice(0, 3).map((x) => x.slice(0, 18)).join(' | ') + (err ? '  （例外：' + err + '）' : ''));
  }
  for (const [name, file, find, rep, secs, expect] of UI_MUTANTS) {
    const src = fs.readFileSync(path.join(ROOT, '_client', file), 'utf8');
    const i = src.indexOf(find);
    const uniq = i >= 0 && src.indexOf(find, i + 1) < 0;
    C.ok(uniq, '變異點唯一存在：' + name);
    if (!uniq) continue;
    const mutated = src.slice(0, i) + rep + src.slice(i + find.length);
    const snap = C.snap(); C.setQuiet(true);
    let err = '';
    try { await runUiOnce(C, { ui: true, admin: false }, { ui: secs, admin: false }, { urlTail: '/' + file, mutated }); } catch (e) { err = String((e && e.message) || e).slice(0, 200); }
    const nf = C.newFails(snap); C.setQuiet(false); C.restore(snap);
    const red = nf.some((n) => expect.test(n));
    C.ok(red, '變異必須使測試變紅：' + name, err || ('實際變紅的項目：' + nf.slice(0, 4).join(' | ')));
    console.log('    ' + (red ? '紅 ✓ ' : '未變紅 ✗ ') + name + '  → ' + nf.slice(0, 3).map((x) => x.slice(0, 18)).join(' | ') + (err ? '  （例外：' + err + '）' : ''));
  }
};
