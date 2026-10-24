#!/usr/bin/env node
/**
 * 報價單品項「服務期間」（License／MA 起訖日）＋商品目錄「類型」檢查（MA_TRACKING=1；P1 階段 A＝後端／目錄／共用規則）。
 * 用法：node scripts/check-quote-terms.js
 *   CTERM_ONLY=unit|api|compat|server  只跑一個群組　　CTERM_MUT=1  變異測試（改一份程式碼副本，對應測試必須變紅）
 *   CTERM_PORT（預設 3271，限 3270–3279）　CTERM_BASE_REF（舊單／關閉開關比對的基準 commit；預設 ed36964＝本功能之前的 origin/main）　CTERM_BASE_DIR（直接指定已解開的基準資料夾）
 * 不碰真實 data.json／auth.json／audit.log.json：[unit][api][compat] 用記憶體資料庫直接呼叫路由處理器；[server] 用 scripts/_b2-harness.js 在本工作樹建臨時 auth.json／data.json 與 _s_* 假帳號（結束即刪），一次只起一台。
 *   [unit]    _client/quote-terms.js（真實日期、月數、一年／三年快捷、顯示格式、時區無關）、lib/quoteTerms.js、lib/productCatalog.js 的類型正規化／稽核差異／輸出
 *   [api]     POST／PUT 驗證與沿用舊值、標題列不帶欄位、序列化、送簽 blocker、時間軸／稽核摘要、雜湊與簽章不含新欄位、PATCH item-terms（權限／不變量／稽核／通知）、關閉開關
 *   [compat]  舊單相容：與基準程式碼（git archive）逐位元比對 contentHash／儲存形狀／GET 完整回應／Excel／預覽 HTML；MA_TRACKING 關閉時帶了期間欄位的請求也與基準相同
 *   [server]  商品目錄 API（開／關）、稽核、PATCH 關閉時等同未登記、通知深連結、真實伺服器建立／改期間
 *   [browser] 階段 B 補（輸入、快捷、紅字、拖曳保留、儲存重載、預覽、PDF、四種外觀對比、手機寬度、XSS）——本階段未實作，標示「未驗證」，不偽造通過
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、真實費率）。
 */
'use strict';
const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');
const { spawnSync } = require('child_process');
const PORT_N = parseInt(process.env.CTERM_PORT || '3271', 10);
if (!(PORT_N >= 3270 && PORT_N <= 3279)) { console.log('拒絕：本檢查只用 3270–3279 埠'); process.exit(2); }
process.env.B2_PORT = String(PORT_N);
const ROOT = path.join(__dirname, '..');
const G = require('./lib-quote-notes-golden.js');
const MUT = process.env.CTERM_MUT === '1';
const ONLY = process.env.CTERM_ONLY || '';
const BASE_REF = process.env.CTERM_BASE_REF || 'ed36964';

let res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const clone = (x) => JSON.parse(JSON.stringify(x));
const sha = G.sha;

function mkQuote(over) {
  return Object.assign({
    id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00.000Z', updatedAt: '2026-10-08T00:00:00.000Z',
    validUntil: '2026-10-30', products: ['PS'], costBy: null, costFlow: { state: 'na' }, approval: null, discountType: 'none', discountValue: 0,
    items: [
      { lid: 'i1', desc: 'SAP License', unit: '套', qty: 1, unitPrice: 600000, cost: 0, termKind: 'license', termFrom: '2026-01-01', termTo: '2026-12-31' },
      { lid: 'i2', desc: 'Service MA', unit: '式', qty: 2, unitPrice: 100000, cost: 0, termKind: 'ma', termFrom: '2026-03-01', termTo: '2027-02-28' },
      { lid: 'i3', desc: '教育訓練', unit: '式', qty: 1, unitPrice: 50000, cost: 0 },
    ],
  }, over || {});
}
function approvalOf(QA, q, state) {
  return {
    state, rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: '2026-10-08T01:00:00.000Z', submittedBy: 'own1', derived: null,
    steps: [{ tier: 'mgr1', label: '一級主管', status: state === 'approved' ? 'approved' : 'pending', assignee: 'mgr1', by: state === 'approved' ? 'mgr1' : undefined, at: state === 'approved' ? '2026-10-08T02:00:00.000Z' : undefined, comment: '' }],
    cur: state === 'approved' ? 1 : 0, board: null,
    history: [{ at: '2026-10-08T01:00:00.000Z', by: 'own1', action: 'SUBMIT', comment: '', tier: '' }].concat(state === 'approved' ? [{ at: '2026-10-08T02:00:00.000Z', by: 'mgr1', action: 'APPROVE', comment: '', tier: 'mgr1' }] : []),
  };
}

// ═════════════════════════ [unit] ═════════════════════════
function groupUnit(R) {
  const QT = require(path.join(R, '_client/quote-terms.js'));
  const QTS = require(path.join(R, 'lib/quoteTerms.js'));
  const PC = require(path.join(R, 'lib/productCatalog.js'));
  // 真實日期
  t('U1 isRealDate：合法日期（含閏日 2024-02-29）通過', ['2026-01-01', '2026-12-31', '2024-02-29', '2000-02-29', '1900-01-01', '2999-12-31'].every((s) => QT.isRealDate(s)));
  t('U2 isRealDate：不存在的日期／格式錯誤／非字串全部拒絕', ['2026-02-30', '2025-02-29', '1900-02-29', '2026-13-01', '2026-00-10', '2026-01-00', '2026-04-31', '2026/01/01', '2026-1-1', '20260101', ' 2026-01-01', '2026-01-01 ', '2026-01-01T00:00:00Z', '', 'abc', '0000-01-01', '1899-12-31', '3000-01-01'].every((s) => !QT.isRealDate(s))
    && [null, undefined, 20260101, {}, [], true, new Date(2026, 0, 1)].every((v) => !QT.isRealDate(v)));
  // 類型
  t('U3 parseKind：license／ma（不分大小寫、前後空白）；空／none／null／undefined＝一般；其他＝null', QT.parseKind('license') === 'license' && QT.parseKind(' MA ') === 'ma' && QT.parseKind('LICENSE') === 'license' && QT.parseKind('') === '' && QT.parseKind('none') === '' && QT.parseKind(null) === '' && QT.parseKind(undefined) === '' && QT.parseKind('xyz') === null && QT.parseKind(5) === null && QT.parseKind({}) === null && QT.parseKind(['ma']) === null);
  t('U4 parseCatalogKind：none／license／ma；不認得 null', QT.parseCatalogKind(undefined) === 'none' && QT.parseCatalogKind('') === 'none' && QT.parseCatalogKind('ma') === 'ma' && QT.parseCatalogKind('x') === null);
  // 加減
  t('U5 addDays：跨月／跨年／閏年／負數；不合法回空字串', QT.addDays('2026-01-31', 1) === '2026-02-01' && QT.addDays('2026-12-31', 1) === '2027-01-01' && QT.addDays('2024-02-28', 1) === '2024-02-29' && QT.addDays('2025-03-01', -1) === '2025-02-28' && QT.addDays('2026-02-30', 1) === '' && QT.addDays('2026-01-01', NaN) === '' && QT.addDays('2999-12-31', 1) === '');
  t('U6 addMonths：月底取該月最後一天（1/31＋1＝2/28；閏年 2/29），負數月份，跨年', QT.addMonths('2026-01-31', 1) === '2026-02-28' && QT.addMonths('2024-01-31', 1) === '2024-02-29' && QT.addMonths('2024-02-29', 12) === '2025-02-28' && QT.addMonths('2026-03-15', -3) === '2025-12-15' && QT.addMonths('2026-11-30', 3) === '2027-02-28' && QT.addMonths('2026-02-30', 1) === '');
  // 快捷
  t('U7 一年快捷：起日＋12 個月－1 天（2026-01-01→2026-12-31；2026-10-24→2027-10-23；2024-02-29→2025-02-27）', QT.quickEnd('2026-01-01', 1) === '2026-12-31' && QT.quickEnd('2026-10-24', 1) === '2027-10-23' && QT.quickEnd('2024-02-29', 1) === '2025-02-27' && QT.quickEnd('2026-01-31') === '2027-01-30');
  t('U8 三年快捷：起日＋36 個月－1 天（2026-01-01→2028-12-31；2026-10-24→2029-10-23）；不合法起日回空字串', QT.quickEnd('2026-01-01', 3) === '2028-12-31' && QT.quickEnd('2026-10-24', 3) === '2029-10-23' && QT.quickEnd('2026-02-30', 1) === '' && QT.quickEnd('', 1) === '');
  // 月數
  const sp = (a, b) => QT.spanText(a, b);
  t('U9 共 N 個月：整月 12／24／36；跨月不整齊；單日；31 號起算', sp('2026-01-01', '2026-12-31') === '共 12 個月' && sp('2026-01-15', '2027-01-14') === '共 12 個月' && sp('2026-01-01', '2027-12-31') === '共 24 個月' && sp('2026-01-01', '2028-12-31') === '共 36 個月' && sp('2026-01-01', '2026-01-20') === '共 20 天' && sp('2026-01-01', '2026-01-01') === '共 1 天' && sp('2026-01-01', '2026-02-10') === '共 1 個月 10 天' && sp('2026-01-31', '2027-01-30') === '共 12 個月' && sp('2026-03-01', '2027-02-28') === '共 12 個月');
  t('U10 月數：迄日早於起日、不合法日期回空字串／null', sp('2026-12-31', '2026-01-01') === '' && sp('2026-02-30', '2026-03-01') === '' && QT.span('x', 'y') === null && QT.span('2026-01-01', '2026-01-01').days === 1);
  // 隨機性質：任何起日 × 一年／三年快捷，span 必為整 12／36 個月
  let rs = 20261024; const rnd = () => { rs = (rs * 1664525 + 1013904223) >>> 0; return rs / 4294967296; };
  let propOk = true, propBad = '';
  for (let i = 0; i < 3000; i++) {
    const y = 1950 + Math.floor(rnd() * 120), m = 1 + Math.floor(rnd() * 12), d = 1 + Math.floor(rnd() * 28 + (rnd() < 0.2 ? rnd() * 3 : 0));
    const f = QT.addDays(QT.addMonths(`${y}-${String(m).padStart(2, '0')}-01`, 0), d - 1);
    if (!f) continue;
    for (const yrs of [1, 3]) {
      const e = QT.quickEnd(f, yrs); const s = QT.span(f, e);
      if (!e || !s || s.months !== 12 * yrs || s.days !== 0 || !(e >= f)) { propOk = false; propBad = f + ' ' + yrs; }
    }
  }
  t('U11 隨機 3000 個起日：一年／三年快捷之後月數必為整 12／36 個月、零頭 0 天、迄日不早於起日', propOk, propBad);
  let monoOk = true;
  for (let i = 0; i < 2000; i++) {
    const a = QT.addDays('2020-01-01', Math.floor(rnd() * 3000)), b = QT.addDays(a, Math.floor(rnd() * 900));
    const s = QT.span(a, b); if (!s) { monoOk = false; continue; }
    const back = QT.addDays(QT.addMonths(a, s.months), s.days - 1);   // 起日＋整月＋零頭 －1 天 ＝ 迄日
    if (back !== b) monoOk = false;
  }
  t('U12 隨機 2000 組起迄：起日＋月數個月＋零頭天數－1 天 必等於迄日（span 與 addMonths／addDays 一致）', monoOk);
  // 格式
  t('U13 formatRange：「YYYY/MM/DD – YYYY/MM/DD」（en dash、前後空白）；servicePeriodLine 加「服務期間：」；一般類型／不合法回空字串', QT.formatRange('2026-01-01', '2026-12-31') === '2026/01/01 – 2026/12/31' && QT.servicePeriodLine('ma', '2026-01-01', '2026-12-31') === '服務期間：2026/01/01 – 2026/12/31' && QT.servicePeriodLine('', '2026-01-01', '2026-12-31') === '' && QT.servicePeriodLine('ma', '2026-02-30', '2026-12-31') === '' && QT.formatRange('x', 'y') === '');
  // 驗證
  const v = (k, a, b) => QT.validateTerm(k, a, b) || { code: null };   // null＝通過；用 {code:null} 讓變異版本不會因為讀 null 而例外
  t('U14 validateTerm：一般類型不需期間；license／ma 缺任一端→TERM_REQUIRED；不真實／格式錯／顛倒→TERM_INVALID；同一天通過；不認得類型→TERM_INVALID', v('', '', '').code === null && v('none', null, null).code === null && v('license', '', '').code === 'TERM_REQUIRED' && v('ma', '2026-01-01', '').code === 'TERM_REQUIRED' && v('ma', '', '2026-01-01').code === 'TERM_REQUIRED' && v('ma', null, undefined).code === 'TERM_REQUIRED'
    && v('ma', '2026-02-30', '2026-03-01').code === 'TERM_INVALID' && v('ma', '2026/01/01', '2026/12/31').code === 'TERM_INVALID' && v('ma', '2026-12-31', '2026-01-01').code === 'TERM_INVALID' && v('ma', '2026-01-01', '2026-01-01').code === null && v('license', '2026-01-01', '2026-12-31').code === null && v('weird', '2026-01-01', '2026-12-31').code === 'TERM_INVALID');
  // 建議
  t('U15 suggestKind：群組標題 License MA／Service MA→ma；名稱含獨立 MA 或「維護」→ma；License／「授權」→license；MA 優先；Mac／SMART 不誤判', QT.suggestKind('x', '─── License MA ───') === 'ma' && QT.suggestKind('x', '─── Service MA ───') === 'ma' && QT.suggestKind('Service MA（Basis）') === 'ma' && QT.suggestKind('年度維護') === 'ma' && QT.suggestKind('SAP Public Cloud License') === 'license' && QT.suggestKind('軟體授權') === 'license' && QT.suggestKind('SAP PCE License MA') === 'ma' && QT.suggestKind('Mac 設備') === '' && QT.suggestKind('SMART 專案') === '' && QT.suggestKind('ERP 顧問導入專案 PE') === '' && QT.suggestKind('', '') === '' && QT.suggestKind(null, null) === '');
  // 時區無關：用不同 TZ 各跑一次，輸出摘要必須相同
  const code = `const T=require(${J(path.join(R, '_client/quote-terms.js'))});const out=[];let s=7;const r=()=>{s=(s*1664525+1013904223)>>>0;return s/4294967296;};for(let i=0;i<500;i++){const a=T.addDays('2020-01-01',Math.floor(r()*4000));const b=T.addDays(a,Math.floor(r()*1000));out.push([a,b,T.quickEnd(a,1),T.quickEnd(a,3),T.spanText(a,b),T.formatRange(a,b),T.isRealDate(a)].join('|'));}out.push(T.addMonths('2024-02-29',12),T.addDays('2026-03-08',1),T.addDays('2026-11-01',1),T.addDays('2026-10-25',-1));console.log(require('crypto').createHash('sha256').update(out.join('\\n')).digest('hex')+' '+new Date(2026,6,1).getTimezoneOffset()+' '+new Date(2026,0,1).getTimezoneOffset());`;
  const tzs = ['UTC', 'Asia/Taipei', 'America/Los_Angeles', 'Pacific/Kiritimati', 'Pacific/Pago_Pago', 'Australia/Lord_Howe', 'Europe/London'];
  const outs = tzs.map((tz) => spawnSync(process.execPath, ['-e', code], { env: Object.assign({}, process.env, { TZ: tz }), encoding: 'utf8' }).stdout.trim().split(' '));
  t('U16 時區無關：7 種 TZ（UTC、台北、洛杉磯、UTC+14、UTC-11、Lord Howe 半小時夏令、倫敦）跑 500 組日期運算，摘要完全相同', new Set(outs.map((o) => o[0])).size === 1 && outs.every((o) => o[0] && o[0].length === 64), J(outs));
  t('U17 （確認 TZ 設定真的生效）各 TZ 的本機時區位移不全相同', new Set(outs.map((o) => o[1] + '/' + o[2])).size >= 4, J(outs.map((o) => o.slice(1).join('/'))));
  // lib/quoteTerms.js
  t('U18 enabled：只有 MA_TRACKING=1 為真（0、true、on、空、未設都是假）', QTS.enabled({ MA_TRACKING: '1' }) && QTS.enabled({ MA_TRACKING: ' 1 ' }) && !QTS.enabled({}) && !QTS.enabled({ MA_TRACKING: '0' }) && !QTS.enabled({ MA_TRACKING: 'true' }) && !QTS.enabled({ MA_TRACKING: 'on' }) && !QTS.enabled({ MA_TRACKING: '' }));
  const old = { lid: 'a', termKind: 'ma', termFrom: '2026-01-01', termTo: '2026-12-31' };
  const rr = (r, o, on) => QTS.resolveRowTerm(r, o, on === undefined ? true : on);
  t('U19 resolveRowTerm：沒送欄位沿用舊值；送 termTo 只改迄日；kind 送 ""／none／null 改回一般並清掉日期；日期送 ""／null＝清除（之後驗證會要求必填）', J(rr({ lid: 'a' }, old)) === J({ kind: 'ma', from: '2026-01-01', to: '2026-12-31', badKind: false }) && rr({ termTo: '2027-06-30' }, old).from === '2026-01-01' && rr({ termTo: '2027-06-30' }, old).to === '2027-06-30'
    && rr({ termKind: '' }, old).kind === '' && rr({ termKind: 'none' }, old).from === '' && rr({ termKind: null }, old).kind === '' && rr({ termFrom: '' }, old).from === '' && rr({ termTo: null }, old).to === '' && rr({ termKind: 'license' }, old).kind === 'license');
  t('U20 resolveRowTerm：新列（無舊值）只認請求；不認得的 kind→badKind；功能關閉＝完全忽略請求、沿用已存', rr({ termKind: 'ma', termFrom: '2026-01-01', termTo: '2026-01-31' }, null).kind === 'ma' && rr({}, null).kind === '' && rr({ termKind: 'zzz' }, null).badKind === true
    && J(rr({ termKind: 'license', termFrom: '2030-01-01', termTo: '2030-02-01' }, old, false)) === J({ kind: 'ma', from: '2026-01-01', to: '2026-12-31', badKind: false }) && rr({ termKind: 'ma', termFrom: '2026-01-01', termTo: '2026-02-01' }, { lid: 'x' }, false).kind === '');
  const items0 = [{ lid: 't', kind: 'title', desc: 'T', termKind: 'ma' }, { desc: 'A', termKind: 'ma' }, { desc: 'B', termKind: 'license', termFrom: '2026-01-01', termTo: '2026-12-31' }, { desc: 'C', termKind: 'ma', termFrom: '2026-12-31', termTo: '2026-01-01' }];
  const ve = QTS.validateRawItems(items0, [], true, 50);
  t('U21 validateRawItems：標題列略過、列號只數一般品項（第 1 列 A 缺期間）；功能關閉一律通過', ve && ve.code === 'TERM_REQUIRED' && /第 1 列「A」/.test(ve.message) && QTS.validateRawItems(items0, [], false, 50) === null);
  const ve2 = QTS.validateRawItems(items0.slice(2), [], true, 50);
  t('U22 validateRawItems：顛倒→TERM_INVALID 且指出第 2 列；超過 max 的列不處理', ve2 && ve2.code === 'TERM_INVALID' && /第 2 列「C」/.test(ve2.message) && QTS.validateRawItems(items0, [], true, 1) === null);
  t('U23 blockersFor：缺期間／不合法→blocker；完整→無；功能關閉→[]；標題列略過', QTS.blockersFor({ items: [{ lid: 'a', desc: 'X', termKind: 'ma' }, { lid: 'b', desc: 'Y', termKind: 'license', termFrom: '2026-01-01', termTo: '2026-02-01' }, { lid: 'c', kind: 'title', desc: 'T' }, { lid: 'd', desc: 'Z' }] }, true).map((b) => b.code).join() === 'TERM_REQUIRED' && QTS.blockersFor({ items: [{ lid: 'a', desc: 'X', termKind: 'ma' }] }, false).length === 0);
  t('U24 termChanges：類型與期間的 舊→新 文字；沒變回空；只動一般品項欄位', J(QTS.termChanges({ desc: 'a' }, { desc: 'a', termKind: 'license', termFrom: '2026-01-01', termTo: '2026-12-31' })) === J(['類型 一般→License', '期間 ∅→2026/01/01 – 2026/12/31']) && QTS.termChanges(old, Object.assign({}, old)).length === 0 && QTS.termChanges(old, Object.assign({}, old, { termTo: '2027-01-31' }))[0] === '期間 2026/01/01 – 2026/12/31→2026/01/01 – 2027/01/31');
  // productCatalog
  const seed = PC.SEED_PRODUCT_CATALOG;
  t('U25 deriveDropdowns：不傳 opts（預設）輸出與種子一致且沒有 PRODUCT_TERMS 鍵', J(Object.keys(PC.deriveDropdowns(seed))) === J(['OPP_PRODUCTS', 'PRODUCT_SUBSELS']));
  const cat = clone(seed); cat.bus.ERP[1].items[0].termKind = 'ma'; cat.bus.ERP[1].items[0].startsTask = true; cat.bus.ERP[0].items[0].termKind = 'license'; cat.bus.ERP[0].items[1].termKind = 'license'; cat.bus.ERP[0].items[1].enabled = false; cat.bus.ERP[2].enabled = false; cat.bus.ERP[2].items[0].termKind = 'ma';
  const dd = PC.deriveDropdowns(cat, { terms: true });
  t('U26 deriveDropdowns(opts.terms)：PRODUCT_TERMS 只含啟用且類型≠none 的商品；其餘輸出與不帶 opts 相同', J(dd.PRODUCT_TERMS) === J({ 'SAP Public Cloud License': { termKind: 'license', startsTask: false }, 'SAP PCOE License MA': { termKind: 'ma', startsTask: true } }) && J(dd.OPP_PRODUCTS) === J(PC.deriveDropdowns(cat).OPP_PRODUCTS) && J(dd.PRODUCT_SUBSELS) === J(PC.deriveDropdowns(cat).PRODUCT_SUBSELS) && J(PC.deriveDropdowns(cat)) === J(PC.deriveDropdowns({ bus: PC.stripCatalogTermSettings(cat.bus) })), J(dd.PRODUCT_TERMS));
  const nb = clone(seed.bus); nb.ERP[0].items[0].termKind = 'LICENSE'; nb.ERP[0].items[1].termKind = 'none'; nb.ERP[0].items[1].startsTask = true; nb.ERP[0].items[2].startsTask = true; nb.ERP[1].items[0].termKind = 'ma'; nb.ERP[1].items[0].startsTask = 'true'; nb.ERP[1].items[1].termKind = 'ma'; nb.ERP[1].items[1].startsTask = false;
  const nr = PC.normalizeCatalogTermSettings(nb);
  t('U27 normalizeCatalogTermSettings：大小寫正規化；none 不存欄位；類型 none 的 startsTask 丟掉；startsTask 預設不存；字串 "true" 視為真；不修改輸入', nr.errors.length === 0 && nr.bus.ERP[0].items[0].termKind === 'license' && !('startsTask' in nr.bus.ERP[0].items[0]) && !('termKind' in nr.bus.ERP[0].items[1]) && !('startsTask' in nr.bus.ERP[0].items[1]) && !('startsTask' in nr.bus.ERP[0].items[2])
    && nr.bus.ERP[1].items[0].startsTask === true && !('startsTask' in nr.bus.ERP[1].items[1]) && nb.ERP[0].items[0].termKind === 'LICENSE');
  const bad = clone(seed.bus); bad.ITS[0].items[0].termKind = 'xyz'; const nrb = PC.normalizeCatalogTermSettings(bad);
  t('U28 normalizeCatalogTermSettings：不認得的類型→errors（含 bu／群組／商品名）', nrb.errors.length === 1 && nrb.errors[0].name === 'MES' && nrb.errors[0].bu === 'ITS');
  const cleanN = PC.normalizeCatalogTermSettings(clone(seed.bus));
  t('U29 沒用到類型的目錄：normalize／strip 之後與原本逐位元相同（預設值不存欄位）', J(cleanN.bus) === J(seed.bus) && J(PC.stripCatalogTermSettings(seed.bus)) === J(seed.bus));
  const withT = { bus: clone(nr.bus) };
  t('U30 stripCatalogTermSettings：拿掉兩個欄位；沒有這兩欄的項目維持同一個物件（===）', (() => { const s = PC.stripCatalogTermSettings(withT.bus); return !JSON.stringify(s).includes('termKind') && !JSON.stringify(s).includes('startsTask') && s.ERP[0].items[3] === withT.bus.ERP[0].items[3]; })());
  const dft = PC.diffTermSettings({ bus: seed.bus }, nr.bus);
  t('U31 diffTermSettings：類型與啟動任務卡的 舊→新', dft.includes('SAP Public Cloud License: 類型 一般→License') && dft.includes('SAP PCOE License MA: 類型 一般→MA') && dft.includes('SAP PCOE License MA: 成交後啟動任務卡 否→是') && PC.diffTermSettings({ bus: seed.bus }, seed.bus).length === 0, J(dft));
  t('U32 catalogForOutput：開啟→多 maTracking:true（其餘不變）；關閉→去掉欄位；舊目錄關閉時回同一個物件（輸出逐位元相同）', (() => { const on = PC.catalogForOutput({ bus: nr.bus, updatedAt: 'x' }, true); const off = PC.catalogForOutput({ bus: nr.bus, updatedAt: 'x' }, false); return on.maTracking === true && J(on.bus) === J(nr.bus) && !JSON.stringify(off).includes('termKind') && !('maTracking' in off) && PC.catalogForOutput(seed, false) === seed; })());
  const dataR = { opportunities: [{ product: 'SAP PCOE License MA' }, { product: 'SAP PCOE License MA 明細' }, { product: 'x' }], contracts: [{ product: 'SAP PCOE License MA' }] };
  t('U33 applyCatalogRenames 不受影響（帶類型的商品改名連動筆數正確、dryRun 不寫入）', (() => { const d = clone(dataR); const a = PC.applyCatalogRenames(d, [{ old: 'SAP PCOE License MA', new: 'X' }], { dryRun: true }); const b = PC.applyCatalogRenames(d, [{ old: 'SAP PCOE License MA', new: 'X' }], {}); return a.updatedOpps === 2 && a.updatedContracts === 1 && b.updatedOpps === 2 && d.opportunities[1].product === 'X 明細'; })());
}

// ═════════════════════════ [api] ═════════════════════════
const postBody = (items, extra) => Object.assign({ oppId: 'O1', company: 'TestCo', projectName: 'P', products: ['PS'], validUntil: '2026-10-30', discountType: 'none', discountValue: 0, items }, extra || {});
const LIC = (over) => Object.assign({ desc: 'SAP License', unit: '套', qty: 1, unitPrice: 1000, termKind: 'license', termFrom: '2026-01-01', termTo: '2026-12-31' }, over || {});
const GEN = (over) => Object.assign({ desc: '教育訓練', unit: '式', qty: 1, unitPrice: 500 }, over || {});
const stored = (env, i) => env.data.quotations[i || 0];
const itemsOf = (q, extra) => q.items.map((x) => Object.assign({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice }, x.kind ? { kind: x.kind } : {}, extra ? extra(x) : {}));
const termKeys = (it) => ['termKind', 'termFrom', 'termTo'].filter((k) => k in it);

async function groupApi(R) {
  const QA = require(path.join(R, 'lib/quoteApproval.js'));
  const prev = process.env.MA_TRACKING;
  const setFlag = (v) => { if (v === undefined) delete process.env.MA_TRACKING; else process.env.MA_TRACKING = v; };
  try {
    setFlag('1');
    const mkq = (state, over) => { const q = mkQuote(over); if (state) q.approval = approvalOf(QA, q, state); return q; };
    // ── POST ──
    let env = G.mkEnv(R, []);
    let c = await env.call('own1', 'POST', '/api/quotations', {}, postBody([{ kind: 'title', desc: 'Part', termKind: 'ma', termFrom: '2026-01-01', termTo: '2026-02-01' }, LIC(), GEN(), { kind: 'subtotal', desc: '', termKind: 'license', termFrom: '2026-01-01', termTo: '2026-02-01' }, LIC({ desc: 'Service MA', termKind: 'MA', termFrom: ' 2026-03-01 ', termTo: '2027-02-28' })]));
    let si = c.s === 201 ? stored(env).items : [];
    t('A1 POST：License／MA 品項存 termKind／termFrom／termTo（kind 小寫化、日期去前後空白）；一般品項不存任何新欄位', c.s === 201 && si[1].termKind === 'license' && si[1].termFrom === '2026-01-01' && si[1].termTo === '2026-12-31' && si[4].termKind === 'ma' && si[4].termFrom === '2026-03-01' && termKeys(si[2]).length === 0, c.s + J(c.j && c.j.error));
    t('A2 POST：標題列／小計列即使請求帶了期間欄位也不存（欄位集合只有 lid、kind、desc）', J(Object.keys(si[0])) === J(['lid', 'kind', 'desc']) && J(Object.keys(si[3])) === J(['lid', 'kind', 'desc']));
    t('A3 POST：一般品項欄位集合與以前相同（lid、desc、unit、qty、unitPrice、cost）；License 品項多三個欄位且順序固定', J(Object.keys(si[2])) === J(['lid', 'desc', 'unit', 'qty', 'unitPrice', 'cost']) && J(Object.keys(si[1])) === J(['lid', 'desc', 'unit', 'qty', 'unitPrice', 'cost', 'termKind', 'termFrom', 'termTo']));
    const qid = stored(env).id;
    for (const who of ['own1', 'admin1', 'mgr1', 'sec1']) {
      const g = await env.call(who, 'GET', '/api/quotations/:id', { id: qid });
      t('A4 GET 序列化（' + who + '）：License／MA 品項輸出 termKind／termFrom／termTo，一般品項、標題列沒有這些鍵', g.s === 200 && g.j.items[1].termKind === 'license' && g.j.items[1].termTo === '2026-12-31' && g.j.items[4].termKind === 'ma' && termKeys(g.j.items[2]).length === 0 && termKeys(g.j.items[0]).length === 0, g.s);
    }
    // 驗證失敗
    const snap = () => J(env.data) + env.saves;
    const bads = [
      ['缺起迄日', [GEN(), LIC({ termFrom: undefined, termTo: undefined })], 'TERM_REQUIRED', /第 2 列「SAP License」/],
      ['只缺迄日', [LIC({ termTo: '' })], 'TERM_REQUIRED', /第 1 列/],
      ['只缺起日', [LIC({ termFrom: null })], 'TERM_REQUIRED', /第 1 列/],
      ['日期不是字串', [LIC({ termFrom: 20260101 })], 'TERM_REQUIRED', /第 1 列/],
      ['迄日早於起日', [LIC({ termFrom: '2026-12-31', termTo: '2026-01-01' })], 'TERM_INVALID', /第 1 列/],
      ['不存在的日期 2026-02-30', [LIC({ termTo: '2026-02-30' })], 'TERM_INVALID', /第 1 列/],
      ['閏日誤用 2025-02-29', [LIC({ termFrom: '2025-02-29' })], 'TERM_INVALID', /第 1 列/],
      ['格式 2026/01/01', [LIC({ termFrom: '2026/01/01' })], 'TERM_INVALID', /第 1 列/],
      ['文字日期', [LIC({ termTo: 'next year' })], 'TERM_INVALID', /第 1 列/],
      ['不認得的類型', [LIC({ termKind: 'xyz' })], 'TERM_INVALID', /第 1 列/],
      ['類型非字串', [LIC({ termKind: 5 })], 'TERM_INVALID', /第 1 列/],
      ['標題列在前不佔列號', [{ kind: 'title', desc: 'T' }, { kind: 'title', desc: 'T2' }, GEN(), LIC({ termTo: '' })], 'TERM_REQUIRED', /第 2 列/],
    ];
    for (const [label, items, code, re] of bads) {
      env = G.mkEnv(R, []);
      const r = await env.call('own1', 'POST', '/api/quotations', {}, postBody(items));
      t('A5 POST 驗證：' + label + ' → 400 ' + code + '，訊息指出第幾列，沒有建立報價單', r.s === 400 && r.j.code === code && re.test(r.j.error || r.j.message || '') && env.data.quotations.length === 0 && env.saves === 0, r.s + J(r.j));
    }
    env = G.mkEnv(R, []);
    c = await env.call('own1', 'POST', '/api/quotations', {}, postBody([LIC({ termKind: '', termFrom: '2026-01-01', termTo: '2026-12-31' }), LIC({ termKind: 'none' }), LIC({ termKind: null, termFrom: undefined, termTo: undefined })]));
    t('A6 POST：類型送 ""／none／null＝一般品項，不要求也不儲存期間', c.s === 201 && stored(env).items.every((x) => termKeys(x).length === 0), c.s);
    env = G.mkEnv(R, []);
    c = await env.call('own1', 'POST', '/api/quotations', {}, postBody([LIC({ termFrom: '2026-02-01', termTo: '2026-02-01' }), LIC({ desc: '閏', termFrom: '2024-02-29', termTo: '2025-02-27' })]));
    t('A7 POST：同一天（期間 1 天）與閏日合法', c.s === 201, c.s + J(c.j && c.j.error));
    // ── PUT ──
    env = G.mkEnv(R, [mkq(null)]);
    const hashBefore = QA.contentHash(stored(env));
    let u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env)) });
    t('A8 PUT：舊分頁的請求不帶任何期間欄位 → 沿用舊值（期間不會被洗掉）', u.s === 200 && stored(env).items[0].termFrom === '2026-01-01' && stored(env).items[0].termTo === '2026-12-31' && stored(env).items[1].termKind === 'ma' && termKeys(stored(env).items[2]).length === 0 && u.j.items[1].termTo === '2027-02-28', u.s + J(u.j && u.j.error));
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), (x) => (x.lid === 'i1' ? { termTo: '2027-06-30' } : {})) });
    t('A9 PUT：只送 termTo → 只改迄日，起日與類型沿用', u.s === 200 && stored(env).items[0].termFrom === '2026-01-01' && stored(env).items[0].termTo === '2027-06-30' && stored(env).items[0].termKind === 'license', u.s);
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), (x) => (x.lid === 'i2' ? { termKind: '' } : {})) });
    t('A10 PUT：類型改回一般（送 ""）→ 欄位整個移除（日期也清掉），其他列不動', u.s === 200 && termKeys(stored(env).items[1]).length === 0 && stored(env).items[0].termKind === 'license', u.s);
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), (x) => (x.lid === 'i3' ? { termKind: 'ma', termFrom: '2026-05-01', termTo: '2027-04-30' } : {})) });
    t('A11 PUT：一般品項改成 MA 並填期間 → 儲存；之後原樣送回仍沿用', u.s === 200 && stored(env).items[2].termKind === 'ma' && stored(env).items[2].termTo === '2027-04-30', u.s);
    const s0 = snap();
    for (const [label, edit, code] of [
      ['改成 License 但清掉日期', (x) => (x.lid === 'i1' ? { termFrom: '' } : {}), 'TERM_REQUIRED'],
      ['迄日顛倒', (x) => (x.lid === 'i1' ? { termFrom: '2028-01-01' } : {}), 'TERM_INVALID'],
      ['一般品項設 License 沒給日期', (x) => (x.lid === 'i2' ? { termKind: 'license' } : {}), 'TERM_REQUIRED'],
      ['不存在的日期', (x) => (x.lid === 'i1' ? { termTo: '2026-06-31' } : {}), 'TERM_INVALID'],
    ]) {
      u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), edit) });
      t('A12 PUT 驗證：' + label + ' → 400 ' + code + ' 且資料逐位元不變、沒有儲存', u.s === 400 && u.j.code === code && snap() === s0, u.s + J(u.j));
    }
    // lid 重複：後面的視為新列，需要自己的期間
    env = G.mkEnv(R, [mkq(null)]);
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env)).concat([{ lid: 'i1', desc: '重複 lid', unit: '式', qty: 1, unitPrice: 1, termKind: 'ma' }]) });
    t('A13 PUT：lid 重複的後一列視為新列，請求若標 MA 卻沒給日期 → 400 TERM_REQUIRED（驗證與儲存對舊列的對應規則一致）', u.s === 400 && u.j.code === 'TERM_REQUIRED' && /第 4 列/.test(u.j.error), u.s + J(u.j));
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env)).concat([{ lid: 'i1', desc: '重複 lid', unit: '式', qty: 1, unitPrice: 1 }]) });
    t('A14 PUT：lid 重複的後一列（沒有期間欄位）存成新的一般列，不會繼承第一列的期間', u.s === 200 && stored(env).items.length === 4 && termKeys(stored(env).items[3]).length === 0 && stored(env).items[3].lid !== 'i1' && stored(env).items[0].termKind === 'license');
    // 標題列轉換
    env = G.mkEnv(R, [mkq(null)]);
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), (x) => (x.lid === 'i1' ? { kind: 'title' } : {})), rowKinds: 1 });
    t('A15 PUT：License 品項改成分組標題 → 期間欄位一併丟掉（標題列只有 lid、kind、desc）', u.s === 200 && J(Object.keys(stored(env).items[0])) === J(['lid', 'kind', 'desc']), u.s + J(u.j && u.j.error));
    // 時間軸／稽核摘要
    env = G.mkEnv(R, [mkq(null)]);
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), (x) => (x.lid === 'i1' ? { termTo: '2027-06-30' } : {})) });
    const tl = (stored(env).timeline || []).slice(-1)[0] || {}; const lg = env.logs[env.logs.length - 1] || [];
    t('A16 PUT 改期間：時間軸 DRAFT_SAVE 摘要含「服務期間」；稽核 UPDATE_QUOTATION 含「期間 舊→新」', u.s === 200 && tl.action === 'DRAFT_SAVE' && /服務期間/.test(tl.comment) && lg[0] === 'UPDATE_QUOTATION' && /期間 2026\/01\/01 – 2026\/12\/31→2026\/01\/01 – 2027\/06\/30/.test(lg[3]), J(tl) + J(lg));
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env)) });
    t('A17 PUT 原樣再存（期間沒變）→ 稽核 UPDATE_QUOTATION 不含「期間」（時間軸同人同動作 60 秒內合併，不看）', u.s === 200 && (env.logs[env.logs.length - 1] || [])[0] === 'UPDATE_QUOTATION' && !/期間/.test((env.logs[env.logs.length - 1] || [])[3]), J(env.logs[env.logs.length - 1]));
    // 雜湊與簽章不含新欄位
    const fx = G.makeFixtures(R, 40);
    const withTerms = (q) => { const x = clone(q); x.items.forEach((it, i) => { if (!it.kind) { it.termKind = i % 2 ? 'license' : 'ma'; it.termFrom = '2026-01-0' + (1 + (i % 9)); it.termTo = '2027-01-0' + (1 + (i % 9)); } }); return x; };
    t('A18 40 張單（含標題／小計列、新舊式、核准／簽核中）每個品項加上期間：contentHash、itemsSig、costLinesSig、lineStructureSig（兩種）、structureSig 全部不變', fx.every((q) => { const w = withTerms(q); return QA.contentHash(w) === QA.contentHash(q) && QA.itemsSig(w) === QA.itemsSig(q) && QA.costLinesSig(w) === QA.costLinesSig(q) && QA.lineStructureSig(w.items) === QA.lineStructureSig(q.items) && QA.lineStructureSig(w.items, { newStyle: true }) === QA.lineStructureSig(q.items, { newStyle: true }) && QA.structureSig(w) === QA.structureSig(q); }));
    const qh = mkq(null); const qh2 = clone(qh); qh2.items[0].termFrom = '2030-05-05'; qh2.items[1].termKind = 'license'; delete qh2.items[0].termKind;
    t('A19 同一張報價單改期間／改類型／拿掉類型前後 contentHash 與所有簽章相等', QA.contentHash(qh) === QA.contentHash(qh2) && QA.itemsSig(qh) === QA.itemsSig(qh2) && QA.structureSig(qh) === QA.structureSig(qh2));
    // 已核准的單：完整表單 PUT 改期間不使核准失效
    env = G.mkEnv(R, [mkq('approved')]);
    const hAppr = stored(env).approval.hash;
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), (x) => (x.lid === 'i1' ? { termTo: '2027-06-30' } : {})) });
    t('A20 已核准 + 完整表單 PUT 只改期間 → 200（不需 confirmVoid）、核准仍有效、hash 不變、歷程無 INVALIDATE', u.s === 200 && stored(env).approval.state === 'approved' && stored(env).approval.hash === hAppr && QA.contentHash(stored(env)) === hAppr && u.j.approval.valid === true && stored(env).items[0].termTo === '2027-06-30' && !stored(env).approval.history.some((h) => h.action === 'INVALIDATE'), u.s + J(u.j && u.j.error));
    // 送簽 blocker
    env = G.mkEnv(R, [mkq(null)]);
    delete stored(env).items[0].termTo; stored(env).items[1].termFrom = '2026-13-45';
    let g = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
    const codes = (g.j.preview.blockers || []).map((b) => b.code);
    t('A21 preview.blockers：License／MA 品項缺期間→TERM_REQUIRED、日期不合法→TERM_INVALID（訊息指出第幾列）', g.s === 200 && codes.includes('TERM_REQUIRED') && codes.includes('TERM_INVALID') && g.j.preview.blockers.some((b) => /第 1 列/.test(b.message)) && g.j.preview.blockers.some((b) => /第 2 列/.test(b.message)), J(g.j.preview.blockers));
    const sBefore = snap();
    u = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' }, {});
    t('A22 送簽：期間缺漏／不合法 → 400 CANNOT_SUBMIT，blockers 帶 TERM_*，且不改變資料', u.s === 400 && u.j.code === 'CANNOT_SUBMIT' && (u.j.blockers || []).some((b) => /^TERM_/.test(b.code)) && /第 1 列/.test(u.j.error || '') && snap() === sBefore, u.s + J(u.j));
    env = G.mkEnv(R, [mkq(null)]);
    g = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
    t('A23 期間完整的單：preview.blockers 沒有任何 TERM_*', g.s === 200 && !(g.j.preview.blockers || []).some((b) => /^TERM_/.test(b.code)));
    // 設定回應
    g = await env.call('own1', 'GET', '/api/quote-approval/config', {});
    t('A24 config 回應：MA_TRACKING=1 多 maTracking:true', g.s === 200 && g.j.maTracking === true, g.s);

    // ── PATCH item-terms ──
    const P = '/api/quotations/:id/item-terms', NOTES = '/api/quotations/:id/item-notes';
    const patch = (e, who, body, id) => e.call(who, 'PATCH', P, { id: id || 'Q1' }, body);
    env = G.mkEnv(R, [mkq('approved')]);
    const before = clone(stored(env)); const hB = QA.contentHash(before), isB = QA.itemsSig(before), clB = QA.costLinesSig(before), stB = QA.structureSig(before);
    u = await patch(env, 'own1', { items: [{ lid: 'i1', termFrom: '2026-02-01', termTo: '2027-01-31' }, { lid: 'i2', termKind: 'license' }] });
    let sq = stored(env);
    t('A25 已核准 + PATCH item-terms → 200；期間寫入；核准狀態、關卡、hash 欄位完全不變；contentHash／itemsSig／costLinesSig／structureSig 不變', u.s === 200 && sq.items[0].termFrom === '2026-02-01' && sq.items[0].termTo === '2027-01-31' && sq.items[1].termKind === 'license' && sq.items[1].termTo === '2027-02-28' && sq.approval.state === 'approved' && sq.approval.hash === before.approval.hash && J(sq.approval.steps) === J(before.approval.steps)
      && QA.contentHash(sq) === hB && QA.itemsSig(sq) === isB && QA.costLinesSig(sq) === clB && QA.structureSig(sq) === stB && u.j.approval.valid === true, u.s + J(u.j && u.j.error));
    const money = (q) => J(q.items.map((x) => [x.lid, x.desc, x.unit, x.qty, x.unitPrice, x.cost]));
    t('A26 金額欄位逐位元不變、品項順序不變、status 不變、updatedAt 更新、歷程沒有 INVALIDATE', money(sq) === money(before) && sq.status === before.status && sq.updatedAt !== before.updatedAt && !sq.approval.history.some((h) => h.action === 'INVALIDATE'));
    const lgT = env.logs[env.logs.length - 1] || []; const tlT = (sq.timeline || []).slice(-1)[0] || {};
    t('A27 稽核 UPDATE_QUOTE_TERM 記下逐列 舊→新（類型、期間）且標示不影響簽核；時間軸多一筆；回應是完整序列化', lgT[0] === 'UPDATE_QUOTE_TERM' && lgT[1] === 'own1' && lgT[2] === 'QU-1' && /僅修改品項服務期間（不影響簽核）/.test(lgT[3]) && /第 1 列「SAP License」期間 2026\/01\/01 – 2026\/12\/31→2026\/02\/01 – 2027\/01\/31/.test(lgT[3]) && /第 2 列「Service MA」類型 MA→License/.test(lgT[3]) && tlT.action === 'DRAFT_SAVE' && /服務期間/.test(tlT.comment) && u.j.quoteNo === 'QU-1', J(lgT) + J(tlT));
    const nt = env.notes.filter((n) => n[1] === 'quote_term_changed');
    t('A28 通知：只發 quote_term_changed 給已簽核的一級主管（修改者本人不收）；refId＝報價單 id；內容只有單號與公司名，沒有金額／期間文字', nt.length === 1 && nt[0][0] === 'mgr1' && nt[0][4] === 'Q1' && /QU-1/.test(nt[0][3]) && /TestCo/.test(nt[0][3]) && !/NT\$|600000|100000|2026-|2027-|2026\//.test(nt[0][2] + nt[0][3]) && env.notes.every((n) => n[1] === 'quote_term_changed'), J(env.notes));
    // 權限
    for (const stt of [null, 'pending', 'approved']) {
      for (const [who, label] of [['own2', '別的業務'], ['mgr1', '主管'], ['sec1', '秘書'], ['x1', '無關者'], ['gm1', '總經理'], ['cons1', '顧問']]) {
        env = G.mkEnv(R, [mkq(stt)]);
        const sx = snap();
        const a = await patch(env, who, { items: [{ lid: 'i1', termTo: '2027-01-31' }] });
        const n = await env.call(who, 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'x' }] });
        t('A29 權限（' + (stt || '草稿') + '／' + label + '）：PATCH item-terms 被擋，且與 item-notes 同一個結果（狀態碼）；資料不變', a.s === n.s && (a.s === 403 || a.s === 404) && snap() === sx, a.s + '/' + n.s);
      }
    }
    env = G.mkEnv(R, [mkq('pending')]);
    u = await patch(env, 'admin1', { items: [{ lid: 'i1', termTo: '2027-01-31' }] });
    t('A30 管理員可改（簽核中）；通知給目前簽核人（一級主管）與業務擁有者（管理員是修改者）；簽核狀態不變', u.s === 200 && stored(env).approval.state === 'pending' && env.notes.filter((n) => n[1] === 'quote_term_changed').map((n) => n[0]).sort().join() === 'mgr1,own1', u.s + J(env.notes.map((n) => n[0])));
    env = G.mkEnv(R, [mkq(null)]);
    u = await patch(env, 'own1', { items: [{ lid: 'i1', termTo: '2027-01-31' }] });
    t('A31 草稿 + 擁有者改自己的單 → 200、有稽核與時間軸，沒有通知（沒有簽核人，修改者就是擁有者）', u.s === 200 && (env.logs.slice(-1)[0] || [])[0] === 'UPDATE_QUOTE_TERM' && env.notes.length === 0);
    // 不合格請求
    env = G.mkEnv(R, [mkq('approved', { items: [{ lid: 't1', kind: 'title', desc: 'T' }].concat(mkQuote().items) })]);
    const rejects = [
      ['多餘的頂層欄位 status', { items: [{ lid: 'i1', termTo: '2027-01-31' }], status: 'x' }, 'BAD_FIELDS'],
      ['多餘的頂層欄位 approval', { items: [{ lid: 'i1', termTo: '2027-01-31' }], approval: {} }, 'BAD_FIELDS'],
      ['品項帶 qty', { items: [{ lid: 'i1', termTo: '2027-01-31', qty: 9 }] }, 'BAD_FIELDS'], ['品項帶 unitPrice', { items: [{ lid: 'i1', termTo: '2027-01-31', unitPrice: 1 }] }, 'BAD_FIELDS'],
      ['品項帶 desc', { items: [{ lid: 'i1', termTo: '2027-01-31', desc: 'n' }] }, 'BAD_FIELDS'], ['品項帶 cost', { items: [{ lid: 'i1', termTo: '2027-01-31', cost: 1 }] }, 'BAD_FIELDS'], ['品項帶 spec', { items: [{ lid: 'i1', termTo: '2027-01-31', spec: 's' }] }, 'BAD_FIELDS'],
      ['找不到的 lid', { items: [{ lid: 'nope', termTo: '2027-01-31' }] }, 'BAD_LID'], ['沒有 lid', { items: [{ termTo: '2027-01-31' }] }, 'BAD_LID'], ['lid 不是字串', { items: [{ lid: 5, termTo: '2027-01-31' }] }, 'BAD_LID'],
      ['分組標題的 lid', { items: [{ lid: 't1', termTo: '2027-01-31' }] }, 'BAD_LID'], ['重複的 lid', { items: [{ lid: 'i1', termTo: '2027-01-31' }, { lid: 'i1', termTo: '2027-02-01' }] }, 'DUP_LID'],
      ['沒有要改的欄位', { items: [{ lid: 'i1' }] }, 'BAD_ITEM'], ['日期不是字串', { items: [{ lid: 'i1', termTo: 20270131 }] }, 'BAD_ITEM'], ['類型是物件', { items: [{ lid: 'i1', termKind: {} }] }, 'BAD_ITEM'],
      ['items 空陣列', { items: [] }, 'BAD_ITEMS'], ['items 不是陣列', { items: 'x' }, 'BAD_ITEMS'], ['沒有 items', {}, 'BAD_ITEMS'], ['品項是 null', { items: [null] }, 'BAD_ITEM'], ['品項是陣列', { items: [['x']] }, 'BAD_ITEM'],
      ['超過 50 筆', { items: Array.from({ length: 51 }, () => ({ lid: 'i1', termTo: '2027-01-31' })) }, 'TOO_MANY_ITEMS'],
      ['License 改回一般（kind ""）', { items: [{ lid: 'i1', termKind: '' }] }, 'TERM_KIND_LOCKED'], ['License 改回一般（kind none）', { items: [{ lid: 'i1', termKind: 'none' }] }, 'TERM_KIND_LOCKED'], ['License 改回一般（kind null）', { items: [{ lid: 'i2', termKind: null }] }, 'TERM_KIND_LOCKED'],
      ['不認得的類型', { items: [{ lid: 'i1', termKind: 'xyz' }] }, 'TERM_INVALID'], ['一般品項只給日期', { items: [{ lid: 'i3', termFrom: '2026-01-01', termTo: '2026-12-31' }] }, 'TERM_NOT_APPLICABLE'],
      ['一般品項設類型沒給日期', { items: [{ lid: 'i3', termKind: 'ma' }] }, 'TERM_REQUIRED'], ['迄日早於起日', { items: [{ lid: 'i1', termTo: '2025-12-31' }] }, 'TERM_INVALID'], ['不存在的日期', { items: [{ lid: 'i1', termTo: '2026-02-30' }] }, 'TERM_INVALID'],
      ['清掉日期', { items: [{ lid: 'i1', termFrom: '' }] }, 'TERM_REQUIRED'], ['第二筆壞了，第一筆好的也不能寫入', { items: [{ lid: 'i1', termTo: '2027-01-31' }, { lid: 'i2', termTo: 'bad' }] }, 'TERM_INVALID'],
    ];
    let allRej = true, untouched = true; const failedCases = []; const snapR = snap();
    for (const [label, b, code] of rejects) {
      const r = await patch(env, 'own1', b);
      if (r.s !== 400 || r.j.code !== code) { allRej = false; failedCases.push(label + ':' + r.s + '/' + (r.j && r.j.code)); }
      if (snap() !== snapR || env.logs.length || env.notes.length) { untouched = false; failedCases.push(label + ':changed'); }
    }
    t('A32 ' + rejects.length + ' 種不合格的 PATCH（多餘欄位、碰金額、壞 lid、型別錯誤、類型改回一般、日期不合法…）全部 400 且帶預期 code', allRej, failedCases.join(' | '));
    t('A33 被拒絕的 PATCH 不寫入：資料庫內容逐位元不變、沒有儲存、沒有稽核、沒有通知', untouched);
    {
      const bb = await env.call('own1', 'PATCH', P, { id: 'Q1' }, [1, 2]); const b2 = await env.call('own1', 'PATCH', P, { id: 'nope' }, { items: [{ lid: 'i1', termTo: '2027-01-31' }] });
      t('A34 body 不是物件→400 BAD_BODY；報價單不存在→404', bb.s === 400 && bb.j.code === 'BAD_BODY' && b2.s === 404);
    }
    // 無實質變更
    env = G.mkEnv(R, [mkq('approved')]); const sN = snap();
    u = await patch(env, 'own1', { items: [{ lid: 'i1', termKind: 'license', termFrom: '2026-01-01', termTo: '2026-12-31' }] });
    t('A35 沒有實質變更（送出與目前相同）→ 200，不寫入、不留稽核、不通知', u.s === 200 && snap() === sN && env.logs.length === 0 && env.notes.length === 0);
    // 一般品項補上類型、已失效的單
    env = G.mkEnv(R, [mkq('approved')]);
    u = await patch(env, 'own1', { items: [{ lid: 'i3', termKind: 'ma', termFrom: '2026-07-01', termTo: '2027-06-30' }] });
    t('A36 一般品項可補上類型＋期間（核准仍有效）；稽核含「類型 一般→MA」', u.s === 200 && stored(env).items[2].termKind === 'ma' && stored(env).approval.state === 'approved' && /類型 一般→MA/.test((env.logs.slice(-1)[0] || [])[3]), u.s + J(u.j && u.j.error));
    env = G.mkEnv(R, [mkq('approved', { invalid: { at: 'x', by: 'own1', reason: 'r' } })]);
    u = await patch(env, 'own1', { items: [{ lid: 'i1', termTo: '2027-01-31' }] });
    t('A37 已失效（商機作廢）的報價單不可改 → 409 QUOTE_INVALID', u.s === 409 && u.j.code === 'QUOTE_INVALID');
    // PUT 與 PATCH 在 licensed 單上沿用
    // ── 關閉開關 ──
    setFlag(undefined);
    env = G.mkEnv(R, [mkq(null)]);
    g = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
    t('A38 關閉：序列化不輸出任何期間欄位（即使資料裡有）；preview.blockers 無 TERM_*', g.s === 200 && g.j.items.every((x) => termKeys(x).length === 0) && !(g.j.preview.blockers || []).some((b) => /^TERM_/.test(b.code)));
    g = await env.call('own1', 'GET', '/api/quote-approval/config', {});
    t('A39 關閉：config 回應沒有 maTracking 鍵', g.s === 200 && !('maTracking' in g.j));
    env = G.mkEnv(R, [mkq(null)]);
    u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: itemsOf(stored(env), (x) => ({ termKind: 'xyz', termFrom: 'bad', termTo: 'bad' })) });
    t('A40 關閉：PUT 帶垃圾期間欄位被忽略（不驗證、200）；已存的期間原樣保留（關閉再存檔不會洗掉）', u.s === 200 && stored(env).items[0].termKind === 'license' && stored(env).items[0].termTo === '2026-12-31' && stored(env).items[1].termKind === 'ma' && termKeys(stored(env).items[2]).length === 0, u.s + J(u.j && u.j.error));
    env = G.mkEnv(R, []);
    c = await env.call('own1', 'POST', '/api/quotations', {}, postBody([LIC({ termFrom: undefined }), GEN({ termKind: 'ma' })]));
    t('A41 關閉：POST 帶期間欄位（缺日期也一樣）→ 201，品項沒有任何新欄位', c.s === 201 && stored(env).items.every((x) => termKeys(x).length === 0) && J(Object.keys(stored(env).items[0])) === J(['lid', 'desc', 'unit', 'qty', 'unitPrice', 'cost']), c.s + J(c.j && c.j.error));
    {
      const hs = env.routesOf('PATCH', P); let nextArg = 'unset'; let handled = false;
      if (hs) hs[0]({}, { status() { handled = true; return this; }, json() { handled = true; } }, (a) => { nextArg = a; });
      t('A42 關閉：PATCH item-terms 的第一個處理器直接 next(\'route\')（與未登記相同，不回應、不驗證登入）', !!hs && nextArg === 'route' && !handled);
    }
  } finally {
    setFlag(prev);
  }
}

// ═════════════════════════ [compat] 與基準（git archive）逐位元比對 ═════════════════════════
function exportBase() {
  if (process.env.CTERM_BASE_DIR) return { dir: process.env.CTERM_BASE_DIR, cleanup() {} };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cterm-base-'));
  const a = spawnSync('git', ['archive', BASE_REF], { cwd: ROOT, maxBuffer: 512 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + BASE_REF + ' 失敗：' + String(a.stderr));
  const x = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 512 * 1024 * 1024 });
  if (x.status !== 0) throw new Error('tar 失敗：' + String(x.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
const norm = (x) => String(x).replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>');
/** 對一棵程式碼樹跑一組固定請求（舊單 fixture），回傳每項輸出的摘要。withTermFields：請求裡夾帶期間欄位（關閉時必須被忽略） */
async function requestScript(R, withTermFields) {
  const fx = G.makeFixtures(R, 16);
  const env = G.mkEnv(R, clone(fx));
  const out = {};
  const ex = (o) => (withTermFields ? Object.assign({}, o, { termKind: 'ma', termFrom: '2026-01-01', termTo: '2026-12-31' }) : o);
  for (const q of fx) {
    for (const who of ['own1', 'admin1', 'mgr1', 'sec1']) { const g = await env.call(who, 'GET', '/api/quotations/:id', { id: q.id }); out['GET ' + q.id + ' ' + who] = g.s + norm(J(g.j)); }
    const g0 = await env.call('own1', 'GET', '/api/quotations/:id', { id: q.id });
    if (g0.s === 200) {
      const items = g0.j.items.map((it) => ex(Object.assign({ lid: it.lid, desc: it.desc, unit: it.unit, qty: it.qty, unitPrice: it.unitPrice }, it.kind ? { kind: it.kind } : {})));
      const u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: q.id }, { items, rowKinds: 1, costModel: 2 });
      out['PUT ' + q.id] = u.s + norm(J(u.j)) + norm(J(env.data.quotations.find((x) => x.id === q.id).items));
    }
  }
  const po = await env.call('own1', 'POST', '/api/quotations', {}, postBody([ex({ desc: 'SAP License MA', unit: '式', qty: 1, unitPrice: 10 }), ex({ desc: 'B', unit: '式', qty: 2, unitPrice: 5 })]));
  out.POST = po.s + norm(J(po.j)) + norm(J(env.data.quotations.slice(-1)[0].items));
  const cf = await env.call('own1', 'GET', '/api/quote-approval/config', {});
  out.CONFIG = cf.s + norm(J(cf.j));
  out.LOGS = norm(J(env.logs.map((l) => [l[0], l[1], l[2], l[3]])));
  out.NOTES = norm(J(env.notes));
  return out;
}
async function groupCompat(R) {
  const base = exportBase();
  const prev = process.env.MA_TRACKING;
  try {
    const QAcur = require(path.join(R, 'lib/quoteApproval.js')), QAbase = require(path.join(base.dir, 'lib/quoteApproval.js'));
    const fx = G.makeFixtures(R, 40);
    t('C1 舊單（40 張，含標題／小計列、新舊式）：contentHash／itemsSig／costLinesSig／lineStructureSig／structureSig 與基準程式碼逐位元相同', fx.every((q) => [QAcur.contentHash, QAcur.itemsSig, QAcur.costLinesSig, (x) => QAcur.lineStructureSig(x.items), QAcur.structureSig].every((f, i) => f(q) === [QAbase.contentHash, QAbase.itemsSig, QAbase.costLinesSig, (x) => QAbase.lineStructureSig(x.items), QAbase.structureSig][i](q))));
    delete process.env.MA_TRACKING;
    const gBase = await G.goldenDigests(base.dir);
    const gOff = await G.goldenDigests(R);
    t('C2 MA_TRACKING 未設：舊單的 hash／儲存形狀／GET 完整回應（含 preview、approval、perm）／客戶 Excel／客戶預覽 HTML 摘要與基準程式碼逐位元相同', J(gBase) === J(gOff), J(gBase) + ' vs ' + J(gOff));
    process.env.MA_TRACKING = '1';
    const gOn = await G.goldenDigests(R);
    t('C3 MA_TRACKING=1 但舊單沒有期間：同上五項摘要仍與基準相同（開啟功能不改變沒有期間的舊單）', J(gBase) === J(gOn), J(gBase) + ' vs ' + J(gOn));
    delete process.env.MA_TRACKING;
    const sBase = await requestScript(base.dir, true), sOff = await requestScript(R, true);
    const dk = Object.keys(sBase).filter((k) => sBase[k] !== sOff[k]);
    t('C4 MA_TRACKING 未設：夾帶 termKind／termFrom／termTo 的 PUT／POST（16 張舊單 GET×4 角色、PUT、POST、config、稽核、通知，共 ' + Object.keys(sBase).length + ' 項）輸出與基準逐位元相同', dk.length === 0 && Object.keys(sBase).length === Object.keys(sOff).length, dk.slice(0, 5).join(', '));
    if (dk.length) console.log('    差異首項：', dk[0], '\n      前：', String(sBase[dk[0]]).slice(0, 300), '\n      後：', String(sOff[dk[0]]).slice(0, 300));
    process.env.MA_TRACKING = '1';
    const sBase2 = await requestScript(base.dir, false), sOn = await requestScript(R, false);
    const dk2 = Object.keys(sBase2).filter((k) => k !== 'CONFIG' && sBase2[k] !== sOn[k]);
    t('C5 MA_TRACKING=1、請求不含期間欄位（舊畫面）：GET／PUT／POST／稽核／通知與基準逐位元相同（config 只多 maTracking 一個鍵）', dk2.length === 0 && J(Object.assign(JSON.parse(sBase2.CONFIG.slice(3)), { maTracking: true })) === J(JSON.parse(sOn.CONFIG.slice(3))), dk2.slice(0, 5).join(', '));
    if (dk2.length) console.log('    差異首項：', dk2[0], '\n      前：', String(sBase2[dk2[0]]).slice(0, 300), '\n      後：', String(sOn[dk2[0]]).slice(0, 300));
    // 商品目錄：關閉時 deriveDropdowns／輸出與基準相同
    const PCb = require(path.join(base.dir, 'lib/productCatalog.js')), PCc = require(path.join(R, 'lib/productCatalog.js'));
    const cat = clone(PCc.SEED_PRODUCT_CATALOG);
    t('C6 商品目錄：種子的 deriveDropdowns 輸出、catalogForOutput(關閉) 與基準逐位元相同', J(PCb.deriveDropdowns(cat)) === J(PCc.deriveDropdowns(cat)) && J(PCc.catalogForOutput(cat, false)) === J(cat) && PCc.catalogForOutput(cat, false) === cat);
  } finally {
    if (prev === undefined) delete process.env.MA_TRACKING; else process.env.MA_TRACKING = prev;
    base.cleanup();
  }
}

// ═════════════════════════ [server] 真實伺服器（臨時帳號、假資料；3270–3279） ═════════════════════════
async function groupServer() {
  const H = require('./_b2-harness');
  const { ok, eq } = H;
  H.guardClean();
  const dataDirExisted = fs.existsSync(path.join(ROOT, 'data')), upExisted = fs.existsSync(path.join(ROOT, 'uploads'));
  const users = [H.mkUser('_s_admin', 'admin'), H.mkUser('_s_u1', 'user', { bu: ['ERP'] }), H.mkUser('_s_u2', 'user', { bu: ['ERP'] })];
  const opp = { id: 'o-term-1', owner: '_s_u1', stage: 'C', company: '測試公司', product: 'SAP PCOE License MA', amount: '100', bu: 'ERP', createdAt: new Date().toISOString() };
  const SEED = JSON.parse(JSON.stringify(require(path.join(ROOT, 'lib/productCatalog.js')).SEED_PRODUCT_CATALOG));
  const names = ['_s_admin', '_s_u1', '_s_u2'];
  const reg = async (r) => { const logs = await H.getLogs('_s_admin'); return logs.filter((l) => l.action === r); };
  try {
    // 開啟
    await H.fresh(users, { opportunities: [opp] }, { env: { MA_TRACKING: '1' } });
    await H.loginAll(names);
    let r = await H.http('_s_admin', 'GET', '/api/product-catalog');
    ok(r.status === 200 && r.json && r.json.maTracking === true, 'S1 開啟：GET /api/product-catalog 帶 maTracking:true');
    const cat = JSON.parse(JSON.stringify(SEED)); cat.bus.ERP[1].items[0].termKind = 'ma'; cat.bus.ERP[1].items[0].startsTask = true; cat.bus.ERP[0].items[0].termKind = 'LICENSE'; cat.bus.ERP[0].items[1].startsTask = true;
    r = await H.http('_s_admin', 'POST', '/api/admin/product-catalog', { catalog: { bus: cat.bus }, renames: [] });
    ok(r.status === 200 && r.json.ok === true, 'S2 開啟：儲存帶類型的目錄 → 200', r.status + r.text.slice(0, 200));
    r = await H.http('_s_admin', 'GET', '/api/product-catalog');
    const e0 = r.json.bus.ERP[0].items[0], e1 = r.json.bus.ERP[1].items[0], e2 = r.json.bus.ERP[0].items[1];
    ok(e0.termKind === 'license' && !('startsTask' in e0) && e1.termKind === 'ma' && e1.startsTask === true && !('termKind' in e2) && !('startsTask' in e2), 'S3 開啟：讀回 termKind（大小寫正規化）、startsTask 只在類型≠一般時保留、預設值不存欄位 ' + J([e0, e1, e2]));
    let logs = await reg('SAVE_PRODUCT_CATALOG');
    ok(logs.length === 1 && /商品類型設定異動/.test(logs[0].detail) && /SAP Public Cloud License: 類型 一般→License/.test(logs[0].detail) && /SAP PCOE License MA: 類型 一般→MA/.test(logs[0].detail) && /成交後啟動任務卡 否→是/.test(logs[0].detail), 'S4 開啟：稽核 SAVE_PRODUCT_CATALOG 的 detail 含商品類型異動（舊→新） ' + (logs[0] && logs[0].detail));
    const bad = JSON.parse(JSON.stringify(cat)); bad.bus.ITS[0].items[0].termKind = 'xyz';
    r = await H.http('_s_admin', 'POST', '/api/admin/product-catalog', { catalog: { bus: bad.bus }, renames: [] });
    ok(r.status === 400 && r.json.code === 'TERM_KIND_INVALID' && /MES/.test(r.json.error), 'S5 開啟：不認得的類型 → 400 TERM_KIND_INVALID，整份不寫入 ' + r.text.slice(0, 200));
    r = await H.http('_s_admin', 'GET', '/api/product-catalog');
    ok(r.json.bus.ITS[0].items[0].termKind === undefined, 'S6 被拒絕的儲存沒有寫入');
    r = await H.http('_s_admin', 'POST', '/api/admin/product-catalog', { catalog: { bus: cat.bus }, renames: [] });
    logs = await reg('SAVE_PRODUCT_CATALOG');
    ok(r.status === 200 && logs.length === 2 && !/商品類型設定異動/.test(logs[0].detail), 'S7 開啟：同樣內容再存一次 → 稽核不再記商品類型異動 ' + (logs[0] && logs[0].detail));
    // 報價單 e2e
    const qbody = { oppId: 'o-term-1', company: '測試公司', projectName: 'P', products: ['SAP PCOE License MA'], validUntil: '2099-01-31', items: [{ desc: 'SAP PCOE License MA', unit: '式', qty: 1, unitPrice: 100, termKind: 'ma', termFrom: '2026-01-01', termTo: '2026-12-31' }] };
    r = await H.http('_s_u1', 'POST', '/api/quotations', Object.assign({}, qbody, { items: [Object.assign({}, qbody.items[0], { termTo: '' })] }));
    ok(r.status === 400 && r.json.code === 'TERM_REQUIRED', 'S8 開啟：真實伺服器 POST 缺迄日 → 400 TERM_REQUIRED ' + r.text.slice(0, 200));
    r = await H.http('_s_u1', 'POST', '/api/quotations', qbody);
    const qid = r.json && r.json.id;
    ok(r.status === 201 && r.json.items[0].termKind === 'ma' && r.json.items[0].termTo === '2026-12-31', 'S9 開啟：真實伺服器 POST 合法 → 201，回應含期間 ' + r.text.slice(0, 200));
    r = await H.http('_s_u1', 'GET', '/api/quote-approval/config');
    ok(r.status === 200 && r.json.maTracking === true, 'S10 開啟：config 有 maTracking');
    r = await H.http(null, 'PATCH', `/api/quotations/${qid}/item-terms`, { items: [{ lid: 'x', termTo: '2027-01-01' }] });
    ok(r.status === 401, 'S11 未登入 PATCH item-terms → 401 ' + r.status);
    r = await H.http('_s_u2', 'PATCH', `/api/quotations/${qid}/item-terms`, { items: [{ lid: 'x', termTo: '2027-01-01' }] });
    ok(r.status === 403 || r.status === 404, 'S12 別的業務 PATCH → 403/404 ' + r.status);
    const lid = (await H.http('_s_u1', 'GET', `/api/quotations/${qid}`)).json.items[0].lid;
    r = await H.http('_s_admin', 'PATCH', `/api/quotations/${qid}/item-terms`, { items: [{ lid, termTo: '2027-06-30' }] });
    ok(r.status === 200 && r.json.items[0].termTo === '2027-06-30', 'S13 管理員 PATCH → 200，回應是更新後的報價單 ' + r.text.slice(0, 200));
    logs = await reg('UPDATE_QUOTE_TERM');
    ok(logs.length === 1 && /期間 2026\/01\/01 – 2026\/12\/31→2026\/01\/01 – 2027\/06\/30/.test(logs[0].detail), 'S14 稽核 UPDATE_QUOTE_TERM 寫入並含舊→新 ' + (logs[0] && logs[0].detail));
    r = await H.http('_s_u1', 'GET', '/api/notifications');
    const nts = (Array.isArray(r.json) ? r.json : (r.json && r.json.notifications) || []).filter((n) => n.type === 'quote_term_changed');
    ok(nts.length === 1 && nts[0].refId === qid && /測試公司/.test(nts[0].body), 'S15 業務擁有者收到 quote_term_changed 通知（站內，refId＝報價單 id） ' + J(nts));
    const nu = spawnSync(process.execPath, ['-e', "const s=require('./server.js');console.log(s._test.notificationUrlFor('quote_term_changed','abc-123'))"], { cwd: ROOT, encoding: 'utf8', env: Object.assign({}, process.env, { DB_BACKEND: 'json', JWT_SECRET: 'x'.repeat(30), SESSION_BACKEND: 'jwt' }) });
    ok(String(nu.stdout).trim().split('\n').pop() === '/index.html#quote:abc-123', 'S16 notificationUrlFor(quote_term_changed) → 深連結 /index.html#quote:<id> ' + String(nu.stdout).slice(-120));
    // 關閉
    await H.fresh(users, { opportunities: [opp], productCatalog: { bus: cat.bus, updatedAt: 'x', updatedBy: 'x' } }, { env: {} });
    await H.loginAll(names);
    r = await H.http('_s_admin', 'GET', '/api/product-catalog');
    ok(r.status === 200 && !('maTracking' in r.json) && !JSON.stringify(r.json).includes('termKind') && !JSON.stringify(r.json).includes('startsTask'), 'S17 關閉：GET /api/product-catalog 沒有 maTracking，也沒有 termKind／startsTask（即使資料裡有）');
    r = await H.http('_s_admin', 'POST', '/api/admin/product-catalog', { catalog: { bus: cat.bus }, renames: [] });
    logs = await reg('SAVE_PRODUCT_CATALOG');
    const rawData = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
    ok(r.status === 200 && !/商品類型/.test(logs[0].detail) && !JSON.stringify(rawData.productCatalog).includes('termKind') && !JSON.stringify(rawData.productCatalog).includes('startsTask'), 'S18 關閉：儲存帶類型的目錄 → 200，類型欄位被丟掉、稽核不記 ' + (logs[0] && logs[0].detail));
    r = await H.http('_s_admin', 'POST', '/api/admin/product-catalog', { catalog: { bus: bad.bus }, renames: [] });
    ok(r.status === 200, 'S19 關閉：不認得的類型也不驗證（欄位被丟掉）→ 200');
    r = await H.http('_s_u1', 'POST', '/api/quotations', Object.assign({}, qbody, { items: [Object.assign({}, qbody.items[0], { termTo: '' })] }));
    ok(r.status === 201 && !('termKind' in r.json.items[0]), 'S20 關閉：POST 缺迄日也不擋，回應沒有期間欄位 ' + r.text.slice(0, 160));
    const rq = await H.http('_s_u1', 'GET', '/api/quote-approval/config');
    ok(rq.status === 200 && !('maTracking' in rq.json), 'S21 關閉：config 沒有 maTracking');
    const ra = await H.http('_s_admin', 'PATCH', '/api/quotations/zz/item-terms', { items: [{ lid: 'x', termTo: '2027-01-01' }] });
    const rb = await H.http('_s_admin', 'PATCH', '/api/quotations/zz/item-terms-never-registered', { items: [{ lid: 'x', termTo: '2027-01-01' }] });
    ok(ra.status === rb.status && ra.text === rb.text, 'S22 關閉：PATCH item-terms 與「從未登記的 PATCH 路徑」回應完全相同（狀態碼與內容） ' + ra.status + '/' + rb.status);
    const rc = await H.http('_s_u1', 'PATCH', `/api/quotations/${(await H.http('_s_u1', 'GET', '/api/quotations')).json[0].id}/item-terms`, { items: [{ lid: 'x', termTo: '2027-01-01' }] });
    const rd = await H.http('_s_u1', 'PATCH', '/api/quotations/zz/item-terms-never-registered', { items: [{ lid: 'x', termTo: '2027-01-01' }] });
    ok(rc.status === rd.status && rc.status !== 200, 'S23 關閉：業務對自己的單 PATCH item-terms 同樣落到未登記的處理 ' + rc.status + '/' + rd.status);
    // 預設（沒設環境變數）＝關
    r = await H.http('_s_admin', 'GET', '/api/product-catalog');
    ok(!('maTracking' in r.json), 'S24 沒設環境變數＝關');
  } finally {
    await H.finish();
    try { if (!dataDirExisted) fs.rmSync(path.join(ROOT, 'data'), { recursive: true, force: true }); } catch { /* 忽略 */ }
    try { if (!upExisted) fs.rmSync(path.join(ROOT, 'uploads'), { recursive: true, force: true }); } catch { /* 忽略 */ }
  }
  return H.R;
}

// ═════════════════════════ 變異測試 ═════════════════════════
const MUTS = [
  { name: 'M1 拿掉後端必填（license／ma 不要求起迄日）', file: '_client/quote-terms.js', find: "if (!f || !t) return { code: 'TERM_REQUIRED'", rep: "if (false) return { code: 'TERM_REQUIRED'", expect: /A5 POST 驗證：缺起迄日|A12|U14/ },
  { name: 'M2 拿掉「迄日不早於起日」', file: '_client/quote-terms.js', find: "if (t < f) return { code: 'TERM_INVALID'", rep: "if (false) return { code: 'TERM_INVALID'", expect: /A5 POST 驗證：迄日早於起日|U14/ },
  { name: 'M3 讓期間進核准雜湊（contentHash）', file: 'lib/quoteApproval.js', find: "      desc: normText(it.desc),\n      unit: normText(it.unit),\n      qty: normNum(it.qty),\n      unitPrice: normNum(it.unitPrice),\n      cost: normNum(it.cost),\n      ...(kindRow", rep: "      desc: normText(it.desc),\n      unit: normText(it.unit),\n      qty: normNum(it.qty),\n      unitPrice: normNum(it.unitPrice),\n      cost: normNum(it.cost),\n      termFrom: normText(it.termFrom),\n      ...(kindRow", expect: /A18|A19|A20|A25|C1/ },
  { name: 'M4 拿掉舊值沿用（沒送欄位就清空）', file: 'lib/quoteTerms.js', find: "from: has('termFrom') ? clean(r.termFrom) : oldFrom,", rep: "from: has('termFrom') ? clean(r.termFrom) : '',", expect: /A8 PUT|U19|A9 PUT/ },
  { name: 'M5 拿掉通知', file: 'lib/quoteRoutes.js', find: '      notifyTermChange(ctx, q, changed.length);\n', rep: '', expect: /A28|A30/ },
  { name: 'M6 拿掉 PATCH 的擁有者權限檢查', file: 'lib/quoteRoutes.js', find: "if (!(rel.isOwner || rel.isAdmin)) return fail(res, 403, 'NO_PERMISSION', '只有報價單擁有者可以修改服務期間');", rep: '', expect: /A29/ },
  { name: 'M7 拿掉稽核 UPDATE_QUOTE_TERM', file: 'lib/quoteRoutes.js', find: "writeLog('UPDATE_QUOTE_TERM'", rep: "(()=>0)('UPDATE_QUOTE_TERM'", expect: /A27/ },
  { name: 'M8 允許類型由有改回無', file: 'lib/quoteRoutes.js', find: "if (oldKind && !t.kind) throw {", rep: "if (false) throw {", expect: /A32/ },
  { name: 'M9 拿掉功能開關（永遠啟用）', file: 'lib/quoteTerms.js', find: "return String(e.MA_TRACKING == null ? '' : e.MA_TRACKING).trim() === '1';", rep: 'return true;', expect: /A38|A39|A40|A41|U18/ },
  { name: 'M10 序列化不輸出期間', file: 'lib/quoteRoutes.js', find: "if (QT.enabled()) { const tv = QT.termView(it); if (tv) Object.assign(o, tv); }", rep: '', expect: /A4 GET|A8 PUT|A25/ },
  { name: 'M11 拿掉送簽 blocker', file: 'lib/quoteRoutes.js', find: "QT.blockersFor(q, QT.enabled()).forEach(b => blockers.push(b));", rep: '', expect: /A21|A22/ },
  { name: 'M12 PATCH 寫入前不檢查不變量且改到金額欄位（把 qty 設 9）', file: 'lib/quoteRoutes.js', find: "const n = Object.assign({}, it, { termKind: t.kind, termFrom: t.from, termTo: t.to });", rep: "const n = Object.assign({}, it, { termKind: t.kind, termFrom: t.from, termTo: t.to, qty: 9 });", expect: /A25|A26/ },
  { name: 'M13 標題列也存期間欄位', file: 'lib/quoteRoutes.js', find: "return { lid: klid, kind, desc: sanitizeStr(r && r.desc, 200) };", rep: "return { lid: klid, kind, desc: sanitizeStr(r && r.desc, 200), termKind: 'ma', termFrom: '2026-01-01', termTo: '2026-01-02' };", expect: /A2 /, apiOnly: true },
  { name: 'M14 商品目錄類型不正規化（保留任意值）', file: 'lib/productCatalog.js', find: "if (k !== 'none') n.termKind = k; else delete n.termKind;", rep: "if (it.termKind !== undefined) n.termKind = it.termKind; else delete n.termKind;", expect: /U27/ },
  { name: 'M15 稽核差異不含商品類型', file: 'lib/productCatalog.js', find: "if (ov.kind !== nv.kind) out.push(", rep: "if (false) out.push(", expect: /U31/ },
];
function copyTree(dst) {
  fs.mkdirSync(dst, { recursive: true });
  const cp = (from, to) => { fs.mkdirSync(to, { recursive: true }); for (const e of fs.readdirSync(from, { withFileTypes: true })) { const f = path.join(from, e.name), g = path.join(to, e.name); if (e.isDirectory()) cp(f, g); else fs.copyFileSync(f, g); } };   // fs.cpSync 在這個環境會讓行程直接結束，改手動複製
  for (const d of ['lib', '_client', 'templates']) cp(path.join(ROOT, d), path.join(dst, d));
  for (const f of ['package.json']) { try { fs.copyFileSync(path.join(ROOT, f), path.join(dst, f)); } catch { /* 沒有就算了 */ } }
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dst, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
}
async function groupMut() {
  console.log('[變異] 對一份程式碼副本改一處，對應測試必須變紅（只跑 unit＋api；server／compat 群組由下方「真實伺服器」變異另計）');
  let undetected = 0;
  for (const m of MUTS) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cterm-mut-'));
    try {
      copyTree(dir);
      const fp = path.join(dir, m.file); let s = fs.readFileSync(fp, 'utf8').replace(/\r\n/g, '\n');
      if (!s.includes(m.find)) { console.log('  ✗ 變異錨點不存在：' + m.name); undetected++; continue; }
      fs.writeFileSync(fp, s.replace(m.find, () => m.rep));
      res = [];
      try { groupUnit(dir); } catch (e) { t('(unit 例外) ' + e.message.slice(0, 120), false); }
      try { await groupApi(dir); } catch (e) { t('(api 例外) ' + e.message.slice(0, 120), false); }
      const failed = res.filter((x) => !x[1]);
      const hit = failed.some((x) => m.expect.test(x[0]));
      console.log((hit ? '  ✓ ' : '  ✗ ') + m.name + '：' + failed.length + ' 項變紅' + (hit ? '' : '（預期的測試沒變紅）') + (failed.length ? '｜例：' + failed[0][0].slice(0, 60) : ''));
      if (!hit) undetected++;
    } finally {
      spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' });
      try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ }
    }
  }
  console.log('變異測試：' + (MUTS.length - undetected) + '/' + MUTS.length + ' 被偵測到');
  return undetected;
}

// ═════════════════════════ 主程式 ═════════════════════════
function report(title) {
  const bad = res.filter((x) => !x[1]);
  for (const [n, ok, ex] of res) if (!ok) console.log('  ✗ ' + n + (ex ? '\n      ' + ex.slice(0, 400) : ''));
  console.log(`[${title}] ${res.length - bad.length}/${res.length} 通過`);
  return bad.length;
}
(async () => {
  let fails = 0;
  if (MUT) { process.exitCode = (await groupMut()) ? 1 : 0; return; }
  const run = (g) => !ONLY || ONLY === g;
  if (run('unit')) { res = []; groupUnit(ROOT); fails += report('unit'); }
  if (run('api')) { res = []; await groupApi(ROOT); fails += report('api'); }
  if (run('compat')) { res = []; await groupCompat(ROOT); fails += report('compat'); }
  if (run('server')) { const R = await groupServer(); const n = R.fail; if (n) console.log(R.failures.map((f) => '  ✗ ' + f).join('\n')); console.log(`[server] ${R.pass}/${R.pass + R.fail} 通過`); fails += n; }
  if (!ONLY) {
    console.log('[browser] 階段 B 補（本次未驗證）：輸入／快捷鍵／紅字提示、拖曳排序保留期間、儲存重載、預覽與 PDF 的「服務期間」行、Excel 品名下灰字與列高、簽核面板唯讀顯示、四種外觀對比與手機寬度截圖、XSS（品名／期間含 <img onerror>）、renderQuoteItems 變異。');
  }
  console.log(fails ? `\n結果：失敗 ${fails} 項` : '\n結果：全部通過（瀏覽器項目除外，見上）');
  process.exitCode = fails ? 1 : 0;
})().catch((e) => { console.error(e && e.stack || e); process.exitCode = 1; });
