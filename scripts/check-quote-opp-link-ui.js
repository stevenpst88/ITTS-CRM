#!/usr/bin/env node
/**
 * 報價單 ↔ 商機連結：前端靜態接線檢查。用法：node scripts/check-quote-opp-link-ui.js（不需要伺服器、不開瀏覽器）
 * 真實瀏覽器的步驟 0／商機卡按鈕流程由沙盒 e2e 驗（見提交說明）；這裡只守住「接線」不被改壞：
 *   · index.html：商機步驟區塊在 #qsMain 第一個、有 #qOppId／#qOppSearch／#qOppList／#qOppCreateBtn；商機視窗有 #oppEditQuoteRow／#oppEditNewQuoteBtn
 *   · quote-steps.js：DEFS 第一步是 opp、validate 讀 #qOppId、Enter 在商機搜尋框不跳步
 *   · quote.js：saveQuote 對新單擋「沒有 oppId」且只在有值時送 oppId；openQuoteModal 接收 opts.oppId 並在 QSteps.onOpen 之前初始化；列表有「未連結商機」；使用者文字都經 escapeHtml
 *   · app.js：商機建立後通知報價單視窗、openOppEdit 呼叫 qOppModalRow、卡片只在有 linkedQuoteCount 時多一行
 *   · 不碰 glass UI 檔（本分支沒有改它們）；不使用 innerHTML 寫入未跳脫的使用者文字（行為檢查：qOpp 區塊內的 innerHTML 拼接都有 escapeHtml）
 *   · quote-preview／pdf／approval 沒有被改（商機連結不影響客戶文件）
 */
'use strict';
const fs = require('fs');
const path = require('path');
const ROOT = path.join(__dirname, '..');
const rd = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
const res = [];
const t = (n, ok, ex) => res.push([n, !!ok, ex === undefined ? '' : String(ex)]);

const html = rd('_client/index.html'), steps = rd('_client/quote-steps.js'), quote = rd('_client/quote.js'), app = rd('_client/app.js');

const mainStart = html.indexOf('id="qsMain"');
const firstStep = html.slice(mainStart).match(/<div class="qs-step" data-qs="([a-z]+)"/);
t('1.1 index.html：#qsMain 的第一個步驟區塊是 opp', firstStep && firstStep[1] === 'opp', firstStep && firstStep[1]);
['qOppId', 'qOppSearch', 'qOppList', 'qOppCreateBtn', 'qOppChosen', 'qOppPicker', 'qOppAmountHint', 'qOppLockNote', 'oppEditQuoteRow', 'oppEditNewQuoteBtn', 'oppEditQuoteCount', 'oppEditQuoteList'].forEach((id) => t('1.2 index.html 有 #' + id, html.includes('id="' + id + '"')));
t('1.3 「來自商機，不可改；要換客戶請先換商機」提示文字在客戶欄位下', /id="qOppLockNote"[^>]*>來自商機，不可改；要換客戶請先換商機</.test(html));
t('1.4 步驟 id 都唯一（data-qs）', (() => { const ids = [...html.matchAll(/data-qs="([a-z]+)"/g)].map((m) => m[1]); return new Set(ids).size === ids.length; })());

const defs = steps.slice(steps.indexOf('var DEFS = ['));
t('2.1 quote-steps DEFS 第一步是 opp（pre:true、validate:vOpp）', /var DEFS = \[\s*\{ id: 'opp'[^}]*pre: true[^}]*validate: vOpp/.test(defs));
t('2.2 vOpp 讀 #qOppId', /function vOpp\(\) \{ return val\('qOppId'\)/.test(steps));
t('2.3 商機搜尋框 Enter 不跳步', /t\.id === 'qProdSearch' \|\| t\.id === 'qOppSearch'/.test(steps));
t('2.4 DEFS 其餘步驟順序不變（客戶→專案→商品→指派顧問→報價項目→優惠→付款方式→確認）', (() => { const ids = [...defs.matchAll(/\{ id: '([a-z]+)'/g)].map((m) => m[1]).slice(0, 9); return ids.join(',') === 'opp,customer,project,products,consultant,items,discount,payment,confirm'; })());

const save = quote.slice(quote.indexOf('async function saveQuote()'), quote.indexOf('async function saveQuote()') + 1500);
t('3.1 saveQuote：新單（沒有 quoteId）沒選商機 → 擋下，且在檢查公司之前', /const oppIdVal = \(\$\('qOppId'\)/.test(save) && save.indexOf('請先選擇要連結的商機') > 0 && save.indexOf('請先選擇要連結的商機') < save.indexOf('請輸入客戶公司名稱'));
t('3.2 saveQuote：只在有值時才送 oppId（舊單編輯不強迫、不送空值）', /if \(oppIdVal\) payload\.oppId = oppIdVal;/.test(quote));
t('3.3 openQuoteModal 接 opts，且 qOppInit 在 QSteps.onOpen 之前', /async function openQuoteModal\(idOrNull, opts\)/.test(quote) && quote.indexOf('await qOppInit(q, opts);') > 0 && quote.indexOf('await qOppInit(q, opts);') < quote.indexOf('QSteps.onOpen(q)'));
t('3.4 openQuoteModal 開頭保證樣式已注入（從商機卡進來時列表還沒載過）', /async function openQuoteModal[^]*?_qEnsureStyle\(\);/.test(quote.slice(quote.indexOf('async function openQuoteModal'), quote.indexOf('async function openQuoteModal') + 400)));
t('3.5 列表：舊單「未連結商機」、新單連結可點（data-qopp-open）、只負責填成本的顧問不顯示「未連結」', /未連結商機/.test(quote) && /data-qopp-open/.test(quote) && /canSeePrice === false\) \? ''/.test(quote));
t('3.6 商機可選範圍：自己的（owner===me）且階段 A／B／C／D', /o\.owner === me && QOPP_STAGES\.indexOf\(o\.stage\) >= 0/.test(quote) && /QOPP_STAGES = \['A', 'B', 'C', 'D'\]/.test(quote));
t('3.7 商機視窗的「建立報價單」鈕：需 me.canCreate、自己的商機、A／B／C／D', /me\.canCreate === true && o\.owner === me\.username && QOPP_STAGES\.indexOf\(o\.stage\) >= 0/.test(quote));
t('3.8 已核准不能換商機（前端不給更換鈕：只有 none／returned、擁有者本人、且沒失效）', quote.includes("_qOppCanChange = !!me && q.owner === me && !q.invalid && (st === 'none' || st === 'returned' || !q.approval)"));
// qOpp 區塊內所有 innerHTML 拼接的使用者欄位都要 escapeHtml
const blk = quote.slice(quote.indexOf('const QOPP_STAGES'), quote.indexOf('async function openQuoteModal'));
const raw = [...blk.matchAll(/'\s*\+\s*([a-zA-Z_.]+)\s*\+\s*'/g)].map((m) => m[1]).filter((v) => !/^(escapeHtml|rows|n|on|badge|\(|String)/.test(v))
  // 例外：prev.company／o.company 只出現在 qConfirm 的 message（qDialog 內一律 escapeHtml），total 只用在 textContent
  .filter((v) => !['prev.company', 'o.company', 'total'].includes(v));
t('3.9 qOpp 區塊：innerHTML 拼接的變數除了 on／n／badge（內部產生、已跳脫）外都經 escapeHtml', raw.length === 0, raw.join(','));
t('3.10 qOpp 區塊不使用 eval／new Function／document.write', !/\beval\(|new Function|document\.write/.test(blk));

t('4.1 app.js：建立商機成功後通知報價單視窗（onOppCreatedForQuote）', /window\.onOppCreatedForQuote/.test(app) && /await window\.onOppCreatedForQuote\(saved\)/.test(app));
t('4.2 app.js：openOppEdit 呼叫 qOppModalRow（有 typeof 守衛，載入順序不影響）', /typeof qOppModalRow === 'function'\) qOppModalRow\(o\)/.test(app));
t('4.3 app.js：商機卡只在有 linkedQuoteCount 才多一行（沒有連結的卡片外形不變）', /if \(o\.linkedQuoteCount\)/.test(app));
t('4.4 app.js：新增商機模式隱藏報價單列', /if \(isCreate\) \{ const qr = \$\('oppEditQuoteRow'\)/.test(app));

// 沒碰到的檔案（git 比對：本檢查只讀工作樹，這裡確認這些檔案不含商機連結字樣）
['_client/quote-preview.js', '_client/quote-pdf.js', '_client/quote-approval.js'].forEach((f) => { const s = fs.existsSync(path.join(ROOT, f)) ? rd(f) : ''; t('5.1 ' + f + ' 沒有商機連結字樣（客戶文件／簽核畫面不受影響）', !/qOpp|oppId|oppSnapshot/.test(s)); });
t('5.2 glass 樣式檔：main 上沒有商機連結字樣；glass 分支則 q-opp 規則只能出現在「階段 7」區塊（glass.css）', ['_client/glass.css', '_client/style-glass.css'].every((f) => !fs.existsSync(path.join(ROOT, f)) || !/q-opp|qOpp/.test(rd(f)) || (f === '_client/glass.css' && /階段 7 併入 main/.test(rd(f)))));

// 商機作廢→報價單失效、建單入口只給能建單的角色
t('6.1 index.html：「新增報價單」預設隱藏（設定載入後才顯示，避免閃現）；有 #qInvalidBanner、#oppDeleteQuoteWarn、狀態篩選 invalid', /id="addQuoteBtn" style="display:none"/.test(html) && html.includes('id="qInvalidBanner"') && html.includes('id="oppDeleteQuoteWarn"') && /<option value="invalid">/.test(html));
t('6.2 applyQuoteToolbar：依 me.canCreate 顯示「新增報價單」', /ab\.style\.display = me\.canCreate === true \? '' : 'none'/.test(quote));
t('6.3 失效報價單：狀態徽章 invalid、列淡化、動作鈕只有檢視／進度／預覽（沒有 Excel／PDF／簽核／送簽）', quote.includes("if (q.invalid) return { key: 'invalid'") && quote.includes('q-row-invalid') && (() => {
  const i = quote.indexOf("if (q.invalid) {\n    out.push(b('edit'");
  if (i < 0) return false;
  const seg = quote.slice(i, quote.indexOf('return `<div class="q-actions">', i));
  return !/b\('(export|pdf|submit|cost|pnl|approve', '✍)/.test(seg);
})());
t('6.4 失效單開啟為唯讀：qInvalidApply 停用輸入、隱藏儲存；openQuoteModal 不再因 canEdit=false 擋掉失效單', /qInvalidApply\(q\)/.test(quote) && /canEdit === false && !q\.invalid/.test(quote));
t('6.5 商機刪除確認視窗警告「將同時使 N 張報價單失效」', /將同時使 \$\{_n\} 張報價單失效/.test(app));
t('6.6 簽核歷程標籤認得 OPP_VOID_INVALIDATE／OPP_RESTORE', /OPP_VOID_INVALIDATE/.test(rd('_client/quote-approval.js')) && /OPP_RESTORE/.test(rd('_client/quote-approval.js')));

let pass = 0, fail = 0;
res.forEach(([n, ok, ex]) => { if (ok) { pass++; console.log('PASS ' + n); } else { fail++; console.log('FAIL ' + n + (ex ? '  <- ' + ex : '')); } });
console.log('\n報價單連結商機（前端接線）檢查：PASS ' + pass + ' / FAIL ' + fail);
process.exit(fail ? 1 : 0);
