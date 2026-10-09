#!/usr/bin/env node
/**
 * 報價單「品項說明（spec）／備註（note）」檢查。用法：node scripts/check-quote-item-notes.js（不需要伺服器；記憶體資料庫直接呼叫路由 handler；不碰 data.json／auth.json／audit.log.json）
 * 動 lib/quoteItems.js 的 cleanItemText、lib/quoteRoutes.js 的 normalizeItems／item-notes 路由／diffSummary、lib/quoteExcel.js 的品項迴圈、_client/quote-preview.js 的品項列之後必跑。
 *   1) 文字清洗 cleanItemText／cleanSpec／cleanNote（單行純文字、字元數上限、控制字元）
 *   2) normalizeItems（經 POST／PUT）：有填才存欄位、沒送欄位沿用既有、空字串清除、標題／小計列不收、XSS 字樣原樣當純文字；GET 序列化只在有值時輸出
 *   3) 核准雜湊不涵蓋說明／備註：contentHash／itemsSig／costLinesSig／lineStructureSig 對加上說明／備註的單完全不變；動到錢的欄位（數量、單價、增刪品項、折扣…）雜湊一定變
 *   4) 簽核語意（業主決定：只改備註／說明不需重新簽核；動到錢照舊要重簽）——本功能的核心：
 *        · 已核准的單：完整表單 PUT 只改說明／備註 → 核准仍有效（hash 不變、state 不變、不用 confirmVoid），簽核歷程多一行「舊→新」、稽核紀錄有舊→新；改數量／單價／增刪品項／折扣 → 仍是 409 WILL_VOID，確認後核准作廢
 *        · 簽核中（整張鎖定）：一般 PUT 仍是 409 LOCKED_PENDING；PUT /api/quotations/:id/item-notes 只收 { items:[{lid,spec?,note?}] }，其餘欄位一律 400 且資料逐位元不變
 *        · 權限只有擁有者與管理員；沒有實質變更不寫入、不留紀錄；寫入前驗證 contentHash／itemsSig／costLinesSig 不變
 *   5) 給客戶的 Excel：說明（9pt 灰字）與備註（粗體標籤＋棕紅字）寫成 rich text 字串儲存格（= + - @ 開頭也不會變公式）、列高足夠、其餘版面逐位元不變；標題／小計列不受影響；毛利分析 Excel 不受影響
 *   6) 給客戶的網頁預覽（PDF 用同一份 HTML）：說明／備註兩行、全部跳脫；沒有時輸出逐位元不變；前端清洗函式與伺服器逐例相同
 *   7) 位元級相容（黃金值）：沒有說明／備註的舊式資料，在新程式碼上算出的 hash／簽章／存入資料庫的品項／GET 完整回應（含 preview、approval、perm）／客戶 Excel／客戶預覽 HTML，
 *      摘要必須與「功能加入之前」的程式碼（git bc921ee）算出的完全相同（黃金值產生方式見 scripts/lib-quote-notes-golden.js）
 *   8) 前端靜態紀律：表單的輸入框跳脫與長度、展開連結、讀回一律送字串；「說明／備註」小視窗只送有改的欄位；列表按鈕只在簽核中／已核准且是擁有者／管理員時出現
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、真實費率）。
 */
'use strict';
const path = require('path');
const fs = require('fs');
const ROOT = path.join(__dirname, '..');
const G = require('./lib-quote-notes-golden.js');

// ── 黃金值：在「品項說明／備註功能加入之前」的程式碼（git bc921ee）上算出（node scripts/check-quote-item-notes.js --digests <舊版樹>）──
const GOLDEN = { hash: 'e59d12b0b6b35f72', store: 'ff89e0464eb90140', serialize: 'cf3df17e6cf27944', excel: 'cfaf6b17f66be173', preview: '07fe74d247ddb9c3' };

if (process.argv[2] === '--digests') {
  G.goldenDigests(path.resolve(process.argv[3] || ROOT)).then((d) => { console.log(JSON.stringify(d)); }, (e) => { console.error(e.stack); process.exit(1); });
  return;
}

const QI = require(path.join(ROOT, 'lib/quoteItems.js'));
const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
const QE = require(path.join(ROOT, 'lib/quoteExcel.js'));
const PNL = require(path.join(ROOT, 'lib/quotePnlExcel.js'));
const XLSX = require(path.join(ROOT, 'node_modules/xlsx'));
const JSZip = require(path.join(ROOT, 'node_modules/jszip'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const clone = (x) => JSON.parse(JSON.stringify(x));
const read = (p) => fs.readFileSync(p, 'utf8').replace(/\r\n/g, '\n');
const TEMPLATE = path.join(ROOT, 'templates/quotation_template.xlsx');

const mkQuote = (over) => Object.assign({
  id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00.000Z', updatedAt: '2026-10-08T00:00:00.000Z',
  validUntil: '2026-10-30', products: ['PS'], costBy: null, costFlow: { state: 'na' }, approval: null,
  discountType: 'none', discountValue: 0,
  items: [
    { lid: 'i1', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 600000, cost: 0 },
    { lid: 'i2', desc: '教育訓練', unit: '式', qty: 2, unitPrice: 100000, cost: 0 },
    { lid: 'i3', desc: '軟體授權', unit: '套', qty: 2, unitPrice: 50000, cost: 0 },
  ],
}, over || {});
const approvalOf = (q, state, extra) => Object.assign({
  state, rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: '2026-10-08T01:00:00.000Z', submittedBy: 'own1', derived: null,
  steps: [{ tier: 'mgr1', label: '一級主管', status: state === 'approved' ? 'approved' : 'pending', assignee: 'mgr1', by: state === 'approved' ? 'mgr1' : undefined, at: state === 'approved' ? '2026-10-08T02:00:00.000Z' : undefined, comment: '' }],
  cur: state === 'approved' ? 1 : 0, board: null,
  history: [{ at: '2026-10-08T01:00:00.000Z', by: 'own1', action: 'SUBMIT', comment: '', tier: '' }].concat(state === 'approved' ? [{ at: '2026-10-08T02:00:00.000Z', by: 'mgr1', action: 'APPROVE', comment: '', tier: 'mgr1' }] : []),
}, extra || {});
const stateQuote = (state, over) => { const q = mkQuote(over); if (state) q.approval = approvalOf(q, state); return q; };
const envOf = (q) => G.mkEnv(ROOT, [q]);
const stored = (env) => env.data.quotations[0];
/** 完整表單 PUT 的 items（舊→新畫面形狀）：以目前儲存的品項為底，套用 edit(items) */
const formItems = (q, edit) => { const it = clone(q.items).map((x) => { const o = { lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, spec: x.spec || '', note: x.note || '' }; if (x.kind) { o.kind = x.kind; delete o.unit; delete o.qty; delete o.unitPrice; delete o.spec; delete o.note; } return o; }); if (edit) edit(it); return it; };
const money = (q) => J(clone(q.items).map((x) => ({ lid: x.lid, kind: x.kind, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost })));

async function run() {
  // ═════════════════ 1) 文字清洗 ═════════════════
  t('1.1 cleanItemText：非字串（null、undefined、數字、物件、陣列、布林）一律回空字串', [null, undefined, 5, {}, [], ['x'], true, NaN].every((v) => QI.cleanItemText(v, 200) === ''));
  t('1.2 單行純文字：換行（\\n、\\r\\n）、Tab、垂直 Tab、NEL、控制字元都當空白，連續空白收成一個，前後去空白', QI.cleanItemText('  a\r\n b\t\tc\u000bd\u0085e\u0001f  g  ', 200) === 'a b c d e f g', QI.cleanItemText('  a\r\n b\t\tc\u000bd\u0085e\u0001f  g  ', 200));
  t('1.3 行分隔字元 U+2028／U+2029 與不換行空白 U+00A0 也收合', QI.cleanItemText('a' + String.fromCharCode(0x2028) + 'b' + String.fromCharCode(0x2029) + 'c d', 200) === 'a b c d');
  t('1.4 字元數上限：說明 200、備註 300（以 Unicode 字元計，4-byte 字元不被切半），截斷後再去尾端空白', QI.SPEC_MAX === 200 && QI.NOTE_MAX === 300 && Array.from(QI.cleanSpec('字'.repeat(300))).length === 200 && Array.from(QI.cleanNote('字'.repeat(400))).length === 300
    && Array.from(QI.cleanSpec('\u{20BB7}'.repeat(250))).length === 200 && QI.cleanSpec('a'.repeat(199) + ' bcd') === 'a'.repeat(199));
  t('1.5 恰好上限不截斷；XSS／公式字樣原樣保留（純文字，不跳脫不過濾）', QI.cleanSpec('x'.repeat(200)).length === 200 && QI.cleanNote('y'.repeat(300)).length === 300 && QI.cleanSpec('<img src=x onerror=alert(1)>') === '<img src=x onerror=alert(1)>' && QI.cleanNote('=1+1 & "q"') === '=1+1 & "q"');
  t('1.6 itemSpec／itemNote：一般品項取清洗後的值；標題／小計列、null、非物件一律空字串', QI.itemSpec({ spec: ' a\nb ' }) === 'a b' && QI.itemNote({ note: ' n ' }) === 'n' && QI.itemSpec({ kind: 'title', spec: 'x' }) === '' && QI.itemNote({ kind: 'subtotal', note: 'x' }) === '' && QI.itemSpec(null) === '' && QI.itemNote('x') === '' && QI.itemSpec({ kind: 'weird', spec: 'x' }) === 'x');
  t('1.7 既有匯出沒被改動：isNonItemRow／isItemRow／itemRows／lineAmount／subtotalValues／subtotalLabel 仍在且語意不變', ['isNonItemRow', 'isItemRow', 'itemRows', 'lineAmount', 'subtotalValues', 'subtotalLabel', 'normalizeRowKind', 'NON_ITEM_KINDS', 'DEFAULT_SUBTOTAL_LABEL'].every((k) => k in QI) && QI.subtotalValues([{ qty: 2, unitPrice: 5 }, { kind: 'subtotal' }])[1] === 10);

  // ═════════════════ 2) normalizeItems（經 POST／PUT）與序列化 ═════════════════
  const postBody = (items, extra) => Object.assign({ company: 'C', projectName: 'P', products: ['PS'], validUntil: '2026-10-30', discountType: 'none', discountValue: 0, items }, extra || {});
  let env = G.mkEnv(ROOT, []);
  let c = await env.call('own1', 'POST', '/api/quotations', {}, postBody([
    { desc: 'A', unit: '式', qty: 1, unitPrice: 10, spec: '  第一行\n第二行  ', note: ' 備註\t文字 ' },
    { desc: 'B', unit: '式', qty: 1, unitPrice: 10 },
    { desc: 'C', unit: '式', qty: 1, unitPrice: 10, spec: '   ', note: '' },
    { desc: 'D', unit: '式', qty: 1, unitPrice: 10, spec: '字'.repeat(250), note: '字'.repeat(400) },
    { desc: 'E', unit: '式', qty: 1, unitPrice: 10, spec: '<img src=x onerror=alert(1)>', note: '=HYPERLINK("http://example.test")' },
    { desc: 'F', unit: '式', qty: 1, unitPrice: 10, spec: 123, note: { a: 1 } },
  ]));
  const si = c.s === 201 ? stored(env).items : [];
  t('2.1 POST：說明／備註清洗後存入（換行／Tab 變單一空白）；沒送、空白、空字串的品項「沒有 spec／note 欄位」（資料外形與以前相同）', c.s === 201 && si[0].spec === '第一行 第二行' && si[0].note === '備註 文字' && !('spec' in si[1]) && !('note' in si[1]) && !('spec' in si[2]) && !('note' in si[2]), J(si));
  t('2.2 超過上限的截斷成 200／300 字；XSS 與公式字樣原樣當純文字存放；非字串（數字、物件）視同空白', Array.from(si[3].spec).length === 200 && Array.from(si[3].note).length === 300 && si[4].spec === '<img src=x onerror=alert(1)>' && si[4].note === '=HYPERLINK("http://example.test")' && !('spec' in si[5]) && !('note' in si[5]));
  t('2.3 沒有說明／備註的品項欄位集合與以前相同（lid、desc、unit、qty、unitPrice、cost）', J(Object.keys(si[1])) === J(['lid', 'desc', 'unit', 'qty', 'unitPrice', 'cost']), J(Object.keys(si[1])));
  const qid = stored(env).id;
  let g = await env.call('own1', 'GET', '/api/quotations/:id', { id: qid });
  t('2.4 GET 序列化：有值才輸出 spec／note；沒有的品項沒有這兩個鍵；管理員與擁有者都看得到', g.s === 200 && g.j.items[0].spec === '第一行 第二行' && g.j.items[0].note === '備註 文字' && !('spec' in g.j.items[1]) && !('note' in g.j.items[1]) && (await env.call('admin1', 'GET', '/api/quotations/:id', { id: qid })).j.items[0].note === '備註 文字');
  // PUT 語意
  const it0 = () => stored(env).items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice }));   // 舊畫面形狀：不帶 spec／note 欄位
  let u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: qid }, { items: it0() });
  t('2.5 PUT 沒帶 spec／note 欄位（舊版畫面、舊分頁）→ 沿用既有值，不會被洗掉', u.s === 200 && stored(env).items[0].spec === '第一行 第二行' && stored(env).items[0].note === '備註 文字' && stored(env).items[4].spec === '<img src=x onerror=alert(1)>', u.s + J(u.j && u.j.error));
  let its = it0(); its[0].spec = ''; u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: qid }, { items: its });
  t('2.6 只送 spec:""→ 清除說明、備註不動；欄位被移除而不是留空字串', u.s === 200 && !('spec' in stored(env).items[0]) && stored(env).items[0].note === '備註 文字');
  its = it0(); its[0].note = null; its[1].spec = '新說明'; its[1].note = '新備註'; u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: qid }, { items: its });
  t('2.7 null＝清除；原本沒有的品項可新增說明與備註', u.s === 200 && !('note' in stored(env).items[0]) && stored(env).items[1].spec === '新說明' && stored(env).items[1].note === '新備註');
  its = it0(); its.unshift({ desc: '新增列', unit: '式', qty: 1, unitPrice: 1, spec: 's', note: 'n' }); u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: qid }, { items: its });
  t('2.8 新增的品項（沒有 lid）也可以帶說明／備註；其他品項不受影響', u.s === 200 && stored(env).items[0].desc === '新增列' && stored(env).items[0].spec === 's' && stored(env).items[0].note === 'n' && stored(env).items[2].spec === '新說明');
  // 標題／小計列
  env = G.mkEnv(ROOT, []);
  c = await env.call('own1', 'POST', '/api/quotations', {}, postBody([{ kind: 'title', desc: 'Part A', spec: 'x', note: 'y' }, { desc: 'A', unit: '式', qty: 1, unitPrice: 5, spec: 's' }, { kind: 'subtotal', desc: '', spec: 'x', note: 'y' }]));
  t('2.9 分組標題／小計列不收說明／備註（只留 lid、kind、desc）；一般品項的說明照存', c.s === 201 && J(Object.keys(stored(env).items[0])) === J(['lid', 'kind', 'desc']) && J(Object.keys(stored(env).items[2])) === J(['lid', 'kind', 'desc']) && stored(env).items[1].spec === 's');
  g = await env.call('own1', 'GET', '/api/quotations/:id', { id: stored(env).id });
  t('2.10 序列化的標題／小計列也沒有 spec／note', g.s === 200 && !('spec' in g.j.items[0]) && !('note' in g.j.items[2]) && g.j.items[1].spec === 's');
  // 一般品項 → 標題：舊的說明消失（標題沒有說明）
  its = g.j.items.map((x) => ({ lid: x.lid, kind: x.kind, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice }));
  its[1].kind = 'title'; u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: stored(env).id }, { items: its, rowKinds: 1 });
  t('2.11 品項改成分組標題 → 說明被丟掉（標題列沒有說明）', u.s === 200 && J(Object.keys(stored(env).items[1])) === J(['lid', 'kind', 'desc']));

  // ═════════════════ 3) 核准雜湊不涵蓋說明／備註；動到錢一定改雜湊 ═════════════════
  const fx = G.makeFixtures(ROOT, 40);
  const withNotes = (q) => { const x = clone(q); x.items.forEach((it, i) => { if (!it.kind) { it.spec = '說明 ' + i + ' <b>'; it.note = '備註 ' + i + ' =1+1'; } }); return x; };
  t('3.1 40 張單（含標題／小計列、新舊式）每個品項加上說明＋備註：contentHash、itemsSig、costLinesSig、lineStructureSig（兩種）、structureSig 全部不變', fx.every((q) => { const w = withNotes(q); return QA.contentHash(w) === QA.contentHash(q) && QA.itemsSig(w) === QA.itemsSig(q) && QA.costLinesSig(w) === QA.costLinesSig(q) && QA.lineStructureSig(w.items) === QA.lineStructureSig(q.items) && QA.lineStructureSig(w.items, { newStyle: true }) === QA.lineStructureSig(q.items, { newStyle: true }) && QA.structureSig(w) === QA.structureSig(q); }));
  const base = mkQuote();
  const mut = (fn) => { const x = clone(base); fn(x); return QA.contentHash(x); };
  const h0 = QA.contentHash(base);
  const moneyEdits = {
    '數量': (x) => { x.items[0].qty = 2; }, '單價': (x) => { x.items[1].unitPrice = 100001; }, '新增品項': (x) => { x.items.push({ lid: 'i9', desc: 'n', unit: '式', qty: 1, unitPrice: 1, cost: 0 }); },
    '刪除品項': (x) => { x.items.pop(); }, '品名': (x) => { x.items[0].desc = '導入顧問2'; }, '單位': (x) => { x.items[0].unit = '人天'; }, '折扣方式': (x) => { x.discountType = 'percent'; x.discountValue = 90; },
    '成本': (x) => { x.items[0].cost = 5; }, '客戶': (x) => { x.company = 'Other'; }, '付款方式': (x) => { x.payment = { net: 30, items: [{ label: 'a', pct: 100 }] }; },
  };
  for (const [k, fn] of Object.entries(moneyEdits)) t('3.2 ' + k + ' 變動 → contentHash 一定改變（仍須重新簽核）', mut(fn) !== h0);
  t('3.3 只改說明／備註（含清除）→ contentHash 不變；說明／備註 hash 不受大小寫、順序影響（不在 payload 內）', mut((x) => { x.items[0].spec = 'a'; }) === h0 && mut((x) => { x.items[1].note = 'b'; }) === h0 && mut((x) => { x.items[0].spec = 'a'; x.items[0].note = 'z'; delete x.items[0].spec; }) === h0);
  const newStyle = mkQuote({ costLines: [{ lid: 'c1', cat: 'consult', desc: 'PM', vendor: '', note: '', unit: '人天', qty: 2, unitCost: 5000, auto: '', forLid: 'i1' }] });
  const nsN = withNotes(newStyle);
  t('3.4 新式成本（有 costLines）的單加上說明／備註：contentHash、itemsSig、costLinesSig 都不變', QA.contentHash(nsN) === QA.contentHash(newStyle) && QA.itemsSig(nsN) === QA.itemsSig(newStyle) && QA.costLinesSig(nsN) === QA.costLinesSig(newStyle) && QA.itemsSigLegacy(nsN) === QA.itemsSigLegacy(newStyle));
  t('3.5 送簽檢查與簽核路徑（validateForSubmit／buildDerived）不受說明／備註影響', (() => { const a = mkQuote({ items: mkQuote().items.map((x) => Object.assign({}, x, { cost: 100 })) }); const b = withNotes(a); return J(QA.validateForSubmit(a, G.CLASSES)) === J(QA.validateForSubmit(b, G.CLASSES)) && J(QA.buildDerived(a, G.CLASSES)) === J(QA.buildDerived(b, G.CLASSES)); })());

  // ═════════════════ 4) 簽核語意（路由）═════════════════
  // 4a) 已核准：完整表單 PUT 只改說明／備註
  env = envOf(stateQuote('approved'));
  const hAppr = stored(env).approval.hash;
  let body = { items: formItems(stored(env), (a) => { a[0].spec = '導入說明'; a[1].note = '客戶要求週末施工'; }) };
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, body);
  let sq = stored(env);
  t('4.1 已核准 + PUT 只改說明／備註 → 200（不需要 confirmVoid）、核准狀態仍是 approved 且 hash 不變、回應 approval.valid=true', u.s === 200 && sq.approval.state === 'approved' && sq.approval.hash === hAppr && QA.contentHash(sq) === hAppr && u.j.approval.valid === true && sq.items[0].spec === '導入說明' && sq.items[1].note === '客戶要求週末施工', u.s + J(u.j && u.j.error));
  const last = sq.approval.history[sq.approval.history.length - 1];
  t('4.2 簽核歷程多一行 ITEM_TEXT_EDIT（誰、何時、舊→新）；沒有 INVALIDATE；原本的 SUBMIT／APPROVE 還在', last.action === 'ITEM_TEXT_EDIT' && last.by === 'own1' && /「導入顧問」說明 ∅→導入說明/.test(last.comment) && /「教育訓練」備註 ∅→客戶要求週末施工/.test(last.comment) && !sq.approval.history.some((h) => h.action === 'INVALIDATE') && sq.approval.history.slice(0, 2).map((h) => h.action).join() === 'SUBMIT,APPROVE', J(last));
  const lg = env.logs[env.logs.length - 1];
  t('4.3 稽核紀錄 UPDATE_QUOTATION 記下舊→新（說明 ∅→…、備註 ∅→…）且標示核准未作廢；通知只有一則 quote_text_edited 給已簽過的主管（沒有 quote_voided 等其他簽核通知）', lg[0] === 'UPDATE_QUOTATION' && /修改內容｜/.test(lg[3]) && !/核准已作廢/.test(lg[3]) && /說明 ∅→導入說明/.test(lg[3]) && /備註 ∅→客戶要求週末施工/.test(lg[3]) && env.notes.every((n) => n[1] === 'quote_text_edited' && n[0] === 'mgr1') && env.notes.length === 1, J(lg));
  body = { items: formItems(sq, (a) => { a[0].spec = '改過的說明'; a[1].note = ''; }) };
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, body);
  sq = stored(env);
  t('4.4 再改一次（改寫、清除）：仍有效；歷程再多一行，舊→新正確（導入說明→改過的說明；備註 …→∅）', u.s === 200 && sq.approval.state === 'approved' && sq.approval.history.filter((h) => h.action === 'ITEM_TEXT_EDIT').length === 2 && /說明 導入說明→改過的說明/.test(sq.approval.history[sq.approval.history.length - 1].comment) && /備註 客戶要求週末施工→∅/.test(sq.approval.history[sq.approval.history.length - 1].comment));
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(sq) });
  t('4.5 原樣再存一次（沒有任何文字變更）→ 不多歷程行', u.s === 200 && stored(env).approval.history.filter((h) => h.action === 'ITEM_TEXT_EDIT').length === 2);
  // 4b) 動到錢 → 仍要重簽
  const moneyCases = {
    '數量': (a) => { a[0].qty = 3; }, '單價': (a) => { a[1].unitPrice = 123456; }, '新增品項': (a) => { a.push({ desc: '追加', unit: '式', qty: 1, unitPrice: 1 }); }, '刪除品項': (a) => { a.pop(); },
    '品名': (a) => { a[0].desc = '導入顧問（改）'; }, '單位': (a) => { a[2].unit = '台'; },
  };
  for (const [k, fn] of Object.entries(moneyCases)) {
    env = envOf(stateQuote('approved'));
    const b1 = { items: formItems(stored(env), (a) => { fn(a); a[0].note = '同時改了備註'; }) };
    const s0 = J(stored(env));
    const r1 = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, b1);
    const keep = J(stored(env));
    const r2 = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, Object.assign({ confirmVoid: true }, b1));
    t('4.6 已核准 + ' + k + '（同時改備註）→ 409 WILL_VOID 且資料沒動；確認後核准作廢（state none、hash 清空、歷程 INVALIDATE）', r1.s === 409 && r1.j.code === 'WILL_VOID' && keep === s0 && keep.includes('"state":"approved"') && r2.s === 200 && stored(env).approval.state === 'none' && stored(env).approval.hash === null && stored(env).approval.history.some((h) => h.action === 'INVALIDATE'), r1.s + ' ' + r2.s);
  }
  env = envOf(stateQuote('approved'));
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(stored(env)), discountType: 'percent', discountValue: 90 });
  t('4.7 已核准 + 改折扣 → 409 WILL_VOID（折扣也是錢）', u.s === 409 && u.j.code === 'WILL_VOID');
  env = envOf(stateQuote('approved'));
  const pre = J(stored(env));
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(stored(env), (a) => { a[0].qty = 5; a[0].spec = 'x'; }) });
  t('4.8 WILL_VOID 被擋下時：連同一起送來的說明／備註也沒有寫入（整筆不動）', u.s === 409 && J(stored(env)) === pre);
  // 4c) 簽核中：整張鎖定，只有 item-notes 能動
  env = envOf(stateQuote('pending'));
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(stored(env), (a) => { a[0].note = 'x'; }) });
  t('4.9 簽核中 + 一般 PUT（即使只改備註）→ 409 LOCKED_PENDING 照舊（整張鎖定的規則沒放寬）', u.s === 409 && u.j.code === 'LOCKED_PENDING' && !('note' in stored(env).items[0]));
  const NOTES = '/api/quotations/:id/item-notes';
  env = envOf(stateQuote('pending'));
  const before = clone(stored(env));
  const hBefore = QA.contentHash(before), itemsSigB = QA.itemsSig(before), clsB = QA.costLinesSig(before);
  u = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', spec: '簽核中補的說明', note: '緊急：週五前交付' }, { lid: 'i3', note: '含運' }] });
  sq = stored(env);
  const lastP = sq.approval.history[sq.approval.history.length - 1];
  t('4.10 簽核中 + item-notes → 200；說明／備註寫入；簽核狀態、關卡、目前關卡、hash 欄位完全不變；contentHash／itemsSig／costLinesSig 不變', u.s === 200 && sq.items[0].spec === '簽核中補的說明' && sq.items[0].note === '緊急：週五前交付' && sq.items[2].note === '含運' && sq.approval.state === 'pending' && J(sq.approval.steps) === J(before.approval.steps) && sq.approval.cur === before.approval.cur && sq.approval.hash === before.approval.hash
    && QA.contentHash(sq) === hBefore && QA.itemsSig(sq) === itemsSigB && QA.costLinesSig(sq) === clsB, u.s + J(u.j && u.j.error));
  t('4.11 金額相關欄位逐位元不變（lid、品名、單位、數量、單價、成本）、品項順序與數量不變、status 不變、updatedAt 更新', money(sq) === money(before) && sq.items.length === before.items.length && sq.status === before.status && sq.updatedAt !== before.updatedAt && sq.discountType === before.discountType && sq.company === before.company);
  t('4.12 簽核歷程多一行 ITEM_TEXT_EDIT（舊→新）；稽核 UPDATE_QUOTE_ITEM_TEXT 記下舊→新；只發 quote_text_edited 通知（沒有其他簽核通知）、沒有 INVALIDATE', lastP.action === 'ITEM_TEXT_EDIT' && /「導入顧問」說明 ∅→簽核中補的說明、備註 ∅→緊急：週五前交付/.test(lastP.comment) && /「軟體授權」備註 ∅→含運/.test(lastP.comment)
    && env.logs[env.logs.length - 1][0] === 'UPDATE_QUOTE_ITEM_TEXT' && /僅修改品項說明／備註（不影響簽核）｜/.test(env.logs[env.logs.length - 1][3]) && /說明 ∅→簽核中補的說明/.test(env.logs[env.logs.length - 1][3]) && env.notes.every((n) => n[1] === 'quote_text_edited') && !sq.approval.history.some((h) => h.action === 'INVALIDATE'), J(env.logs[env.logs.length - 1]));
  t('4.13 回應是完整序列化的報價單（含新文字、approval.state 仍是 pending）', u.j.items[0].spec === '簽核中補的說明' && u.j.approval.state === 'pending' && u.j.quoteNo === 'QU-1');
  // 已核准也能用
  env = envOf(stateQuote('approved'));
  u = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i2', note: '核准後補充' }] });
  t('4.14 已核准 + item-notes → 200；核准仍有效（valid）、hash 不變；歷程與稽核各一行', u.s === 200 && u.j.approval.state === 'approved' && u.j.approval.valid === true && QA.contentHash(stored(env)) === stored(env).approval.hash && stored(env).approval.history.slice(-1)[0].action === 'ITEM_TEXT_EDIT' && env.logs.slice(-1)[0][0] === 'UPDATE_QUOTE_ITEM_TEXT');
  // 草稿／被駁回：可用，但不寫簽核歷程
  for (const st of [null, 'returned']) {
    env = envOf(st ? stateQuote(st) : stateQuote(null));
    if (st) stored(env).approval.history.push({ at: 'x', by: 'mgr1', action: 'RETURN', comment: 'r', tier: 'mgr1' });
    const nh = stored(env).approval ? stored(env).approval.history.length : 0;
    u = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'n' }] });
    t('4.15 ' + (st || '草稿') + ' + item-notes → 200 並寫稽核；不寫簽核歷程（不蓋掉歷程最後一筆）', u.s === 200 && stored(env).items[0].note === 'n' && env.logs.slice(-1)[0][0] === 'UPDATE_QUOTE_ITEM_TEXT' && (!stored(env).approval || stored(env).approval.history.length === nh));
  }
  // 4d) 只收文字：其餘一律 400，資料逐位元不變
  const rejects = [
    ['多餘的頂層欄位 status', { items: [{ lid: 'i1', note: 'x' }], status: 'accepted' }], ['多餘的頂層欄位 approval', { items: [{ lid: 'i1', note: 'x' }], approval: { state: 'approved' } }],
    ['多餘的頂層欄位 discountType', { items: [{ lid: 'i1', note: 'x' }], discountType: 'percent' }], ['多餘的頂層欄位 company', { items: [{ lid: 'i1', note: 'x' }], company: 'Z' }],
    ['品項帶 qty', { items: [{ lid: 'i1', note: 'x', qty: 9 }] }], ['品項帶 unitPrice', { items: [{ lid: 'i1', note: 'x', unitPrice: 1 }] }], ['品項帶 desc', { items: [{ lid: 'i1', note: 'x', desc: 'new' }] }],
    ['品項帶 cost', { items: [{ lid: 'i1', note: 'x', cost: 1 }] }], ['品項帶 unit', { items: [{ lid: 'i1', note: 'x', unit: '台' }] }], ['品項帶 cat', { items: [{ lid: 'i1', note: 'x', cat: 'consult' }] }], ['品項帶 kind', { items: [{ lid: 'i1', note: 'x', kind: 'title' }] }],
    ['找不到的 lid', { items: [{ lid: 'nope', note: 'x' }] }], ['沒有 lid', { items: [{ note: 'x' }] }], ['lid 不是字串', { items: [{ lid: 5, note: 'x' }] }], ['重複的 lid', { items: [{ lid: 'i1', note: 'x' }, { lid: 'i1', note: 'y' }] }],
    ['沒有要改的欄位', { items: [{ lid: 'i1' }] }], ['spec 不是字串', { items: [{ lid: 'i1', spec: 5 }] }], ['note 不是字串', { items: [{ lid: 'i1', note: { a: 1 } }] }], ['items 空陣列', { items: [] }], ['items 不是陣列', { items: 'x' }], ['沒有 items', {}],
    ['品項是 null', { items: [null] }], ['品項是陣列', { items: [['x']] }], ['超過 50 筆', { items: Array.from({ length: 51 }, () => ({ lid: 'i1', note: 'x' })) }],
  ];
  env = envOf(stateQuote('pending'));
  let allRejected = true, untouched = true; const failedCases = [];
  const snap = J(env.data) + env.saves;
  for (const [label, b] of rejects) {
    const r = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, b);
    if (r.s !== 400) { allRejected = false; failedCases.push(label + ':' + r.s); }
    if (J(env.data) + env.saves !== snap || env.logs.length) { untouched = false; failedCases.push(label + ':changed'); }
  }
  t('4.16 ' + rejects.length + ' 種不合格的請求（多餘欄位、改金額欄位、壞 lid、型別錯誤、空、超量…）全部 400', allRejected, failedCases.join(' | '));
  t('4.17 被拒絕的請求不寫入：資料庫內容逐位元不變、沒有儲存、沒有稽核、沒有歷程', untouched);
  // 標題列不能掛說明
  env = envOf(stateQuote('pending', { items: [{ lid: 't1', kind: 'title', desc: 'Part A' }].concat(mkQuote().items) }));
  u = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 't1', note: 'x' }] });
  t('4.18 對分組標題的 lid 送說明／備註 → 400 BAD_LID（標題列沒有說明／備註）', u.s === 400 && u.j.code === 'BAD_LID' && !('note' in stored(env).items[0]));
  // 4e) 權限
  env = envOf(stateQuote('pending'));
  const perms = {};
  for (const who of ['own2', 'mgr1', 'sec1', 'cons1', 'x1', 'gm1']) perms[who] = (await env.call(who, 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'x' }] })).s;
  t('4.19 只有擁有者與管理員能用：其他業務、主管、秘書、顧問、無關者、總經理一律 403（或 404）且資料不動', Object.values(perms).every((s) => s === 403 || s === 404) && !('note' in stored(env).items[0]) && env.saves === 0, J(perms));
  u = await env.call('admin1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: '管理員補' }] });
  t('4.20 管理員可用（稽核的操作者是管理員）', u.s === 200 && stored(env).items[0].note === '管理員補' && env.logs.slice(-1)[0][1] === 'admin1');
  // 4f) 沒有實質變更
  env = envOf(stateQuote('approved', { items: mkQuote().items.map((x, i) => (i === 0 ? Object.assign({}, x, { spec: 'S', note: 'N' }) : x)) }));
  const snap2 = J(env.data);
  u = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', spec: ' S ', note: 'N' }, { lid: 'i2', spec: '', note: '  ' }] });
  t('4.21 與目前內容相同（含前後空白、空→空）→ 200 但不寫入、不儲存、不留稽核／歷程、updatedAt 不變', u.s === 200 && J(env.data) === snap2 && env.saves === 0 && env.logs.length === 0);
  // 4g) 清洗、截斷、XSS
  env = envOf(stateQuote('pending'));
  u = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', spec: '字'.repeat(260) + '\n', note: '<img src=x onerror=alert(1)>\n第二行' }] });
  t('4.22 item-notes 也套用同一套清洗：截斷 200／300、換行收合；XSS 字樣原樣存成純文字', u.s === 200 && Array.from(stored(env).items[0].spec).length === 200 && stored(env).items[0].note === '<img src=x onerror=alert(1)> 第二行');
  const lc = stored(env).approval.history.slice(-1)[0].comment;
  t('4.23 歷程／稽核裡的文字被截短（每段 ≤60 字加 …）；歷程說明總長 ≤600', /…/.test(lc) && lc.length <= 600 && env.logs.slice(-1)[0][3].length < 1000, lc.length);
  // 4h) 連續多次
  env = envOf(stateQuote('approved'));
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'v1' }] });
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'v2' }] });
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: '' }] });
  t('4.24 連改三次：歷程三行（備註 ∅→v1、v1→v2、v2→∅）、核准全程有效、稽核三筆', stored(env).approval.history.filter((h) => h.action === 'ITEM_TEXT_EDIT').map((h) => h.comment.replace(/^「導入顧問」/, '')).join('|') === '備註 ∅→v1|備註 v1→v2|備註 v2→∅' && QA.contentHash(stored(env)) === stored(env).approval.hash && env.logs.filter((l) => l[0] === 'UPDATE_QUOTE_ITEM_TEXT').length === 3);
  // 4i) 顧問成本流程不受影響
  env = envOf(stateQuote(null, { products: ['PC'], costBy: 'cons1', costFlow: { state: 'filled', by: 'cons1', requestedAt: 'x', filledAt: 'x', note: '', sig: QA.lineStructureSig(mkQuote().items, { newStyle: true }), consultantWrote: true }, costLines: [{ lid: 'c1', cat: 'consult', desc: 'PM', vendor: '', note: '', unit: '人天', qty: 2, unitCost: 5000, auto: '', forLid: 'i1' }] }));
  const cfBefore = J(stored(env).costFlow), clBefore = J(stored(env).costLines), sigB = QA.itemsSig(stored(env));
  u = await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', spec: 's', note: 'n' }] });
  t('4.25 需顧問的新式單（成本已填）：改說明／備註 → 成本流程狀態、成本明細、itemsSig 全部不變（顧問不必重填、畫面不會過期）', u.s === 200 && J(stored(env).costFlow) === cfBefore && J(stored(env).costLines) === clBefore && QA.itemsSig(stored(env)) === sigB && u.j.costFlow.state === 'filled');
  env = envOf(stateQuote(null, { products: ['PC'], costBy: 'cons1', costFlow: { state: 'filled', by: 'cons1', requestedAt: 'x', filledAt: 'x', note: '', sig: QA.lineStructureSig(mkQuote().items, { newStyle: true }), consultantWrote: true }, costLines: [{ lid: 'c1', cat: 'consult', desc: 'PM', vendor: '', note: '', unit: '人天', qty: 2, unitCost: 5000, auto: '', forLid: 'i1' }] }));
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(stored(env), (a) => { a[0].spec = 'x'; a[1].note = 'y'; }), costModel: 2, products: ['PC'], costBy: 'cons1' });
  t('4.26 同上，走完整表單 PUT（只改說明／備註）：成本流程仍是 filled（不退回 requested）', u.s === 200 && stored(env).costFlow.state === 'filled' && stored(env).items[0].spec === 'x');
  // 4j) 伺服器靜態：唯讀帳號擋 PUT、路由有 qAuth
  const srv = read(path.join(ROOT, 'server.js')), qr = read(path.join(ROOT, 'lib/quoteRoutes.js'));
  t('4.27 唯讀帳號（accessMode=view）的寫入攔截涵蓋 PUT（item-notes 是 PUT，自動被擋）；路由掛 qAuth', /\['POST', 'PUT', 'DELETE', 'PATCH'\]\.includes\(req\.method\)/.test(srv) && /app\.put\('\/api\/quotations\/:id\/item-notes', qAuth,/.test(qr));
  t('4.28 item-notes 路由寫入前有「雜湊／簽章不變」防線（contentHash、itemsSig、costLinesSig 三個都比對）', /QA\.contentHash\(probe\) !== QA\.contentHash\(q\) \|\| QA\.itemsSig\(probe\) !== QA\.itemsSig\(q\) \|\| QA\.costLinesSig\(probe\) !== QA\.costLinesSig\(q\)/.test(qr));
  t('4.29 簽核歷程只在簽核中／已核准寫入（其他狀態不寫，避免蓋掉「核准作廢」標示）', /ap\.state !== 'pending' && ap\.state !== 'approved'\) \|\| !edits\.length\) return false;/.test(qr) && /pushHistory\(ap, ctx\.me, 'ITEM_TEXT_EDIT'/.test(qr));

  // ═════════════════ 5) 給客戶的 Excel ═════════════════
  const X0 = mkQuote({ items: [{ lid: 'a', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 1000, cost: 0 }, { lid: 'b', desc: '教育訓練', unit: '式', qty: 2, unitPrice: 500, cost: 0 }] });
  const xb = await G.sheetXml(ROOT, X0);
  const cellC = (xml, r) => { const m = new RegExp('<c r="C' + r + '"[^>]*?(?:/>|>[\\s\\S]*?</c>)').exec(xml); return m ? m[0] : null; };
  const rowAttrs = (xml, r) => { const m = new RegExp('<row r="' + r + '"([^>]*)>').exec(xml); return m ? m[1] : ''; };
  const rowHt = (xml, r) => { const m = /\sht="([\d.]+)"/.exec(rowAttrs(xml, r)); return m ? +m[1] : null; };
  const L = QE.LAYOUT, R1 = L.itemFirst, R2 = L.itemFirst + 1;
  t('5.1 沒有說明／備註：品名儲存格仍是純文字 inlineStr（一個 <t>、沒有 rich text），列高與以前相同', /^<c r="C17" s="\d+" t="inlineStr"><is><t xml:space="preserve">導入顧問<\/t><\/is><\/c>$/.test(cellC(xb.xml, R1)) && rowHt(xb.xml, R1) === QE._internal.itemRowHeight('導入顧問'), cellC(xb.xml, R1));
  const X1 = clone(X0); X1.items[0].spec = '含需求訪談與系統安裝'; X1.items[0].note = '客戶要求週末施工'; X1.items[1].spec = '僅說明';
  const x1 = await G.sheetXml(ROOT, X1);
  const c1 = cellC(x1.xml, R1), c2 = cellC(x1.xml, R2);
  t('5.2 有說明＋備註：rich text 四段（品名 11pt、說明 9pt 灰、「備註：」粗體 10pt 棕紅、備註內容 10pt 棕紅），行與行之間用換行；仍是 inlineStr', (c1.match(/<r>/g) || []).length === 4 && /t="inlineStr"/.test(c1) && /<sz val="11"\/>/.test(c1) && /<sz val="9"\/><color rgb="FF6B7280"\/>/.test(c1) && /<b\/><sz val="10"\/><color rgb="FF9A3412"\/>/.test(c1) && />\n含需求訪談與系統安裝</.test(c1) && />\n備註：</.test(c1) && /客戶要求週末施工/.test(c1), c1);
  t('5.3 只有說明：三段以內（品名＋說明）；只有備註：品名＋「備註：」＋內容', (c2.match(/<r>/g) || []).length === 2 && !c2.includes('備註') && await (async () => { const X2 = clone(X0); X2.items[0].note = '只有備註'; const c = cellC((await G.sheetXml(ROOT, X2)).xml, R1); return (c.match(/<r>/g) || []).length === 3 && />\n備註：</.test(c) && !/FF6B7280/.test(c); })());
  const firstSheet = (buf) => { const w = XLSX.read(buf, { type: 'buffer' }); return w.Sheets[w.SheetNames[0]]; };
  const sheet1 = firstSheet(x1.buf);
  t('5.4 用 SheetJS 讀回：C17 = 品名／說明／「備註：…」三行（字串型 s）；C18 = 品名／說明兩行', sheet1['C17'].t === 's' && sheet1['C17'].v === '導入顧問\n含需求訪談與系統安裝\n備註：客戶要求週末施工' && sheet1['C18'].v === '教育訓練\n僅說明', J(sheet1['C17']));
  // 注入
  const X3 = clone(X0); X3.items[0].desc = '=1+1'; X3.items[0].spec = '+SUM(A1)'; X3.items[0].note = '@cmd'; X3.items[1].desc = '-2+3'; X3.items[1].spec = '=HYPERLINK("http://example.test","x")'; X3.items[1].note = '<img src=x onerror=alert(1)> & "q"';
  const x3 = await G.sheetXml(ROOT, X3);
  const s3 = firstSheet(x3.buf);
  const fCountBase = (xb.xml.match(/<f>/g) || []).length;
  t('5.5 公式注入防護：品名／說明／備註以 = + - @ 開頭時仍是字串（沒有 <f>，公式數與沒有備註時相同）；SheetJS 讀回為文字', (x3.xml.match(/<f>/g) || []).length === fCountBase && !/<c r="C1[78]"[^>]*><f>/.test(x3.xml) && s3['C17'].t === 's' && s3['C17'].v === '=1+1\n+SUM(A1)\n備註：@cmd' && s3['C18'].t === 's' && s3['C18'].v.startsWith('-2+3\n=HYPERLINK('), J(s3['C17']));
  t('5.6 XML 特殊字元跳脫（<、>、&、"）：檔案裡沒有原始的 <img；讀回逐字相同', !/<img/.test(x3.xml) && x3.xml.includes('&lt;img src=x onerror=alert(1)&gt; &amp; "q"') && s3['C18'].v.endsWith('備註：<img src=x onerror=alert(1)> & "q"'));
  // 列高
  const hNoNotes = rowHt(xb.xml, R1), hBoth = rowHt(x1.xml, R1), hSpecOnly = rowHt(x1.xml, R2);
  t('5.7 列高：有說明／備註的列比沒有的高（足夠放下所有行）；只有說明的列比兩者都有的矮；公式與 itemRowHeightWithNotes 一致', hBoth > hNoNotes && hSpecOnly > hNoNotes && hSpecOnly < hBoth && hBoth === QE._internal.itemRowHeightWithNotes('導入顧問', '含需求訪談與系統安裝', '客戶要求週末施工') && hBoth >= 4 + 18 + 14 + 15, hNoNotes + ' ' + hBoth + ' ' + hSpecOnly);
  const XL = clone(X0); XL.items[0].spec = '長'.repeat(200); XL.items[0].note = '字'.repeat(300);
  const xl = await G.sheetXml(ROOT, XL);
  const hL = rowHt(xl.xml, R1);
  const needLines = Math.ceil(400 / L.specWidthUnits) + Math.ceil((600 + 6) / L.noteWidthUnits);
  t('5.8 最長的說明（200 字）＋備註（300 字）：列高隨折行增加、不超過 Excel 上限 409，且夠放（>= 品名 1 行＋說明 5 行＋備註 9 行的估算）', hL > hBoth && hL <= 409 && hL >= 18 + 14 * Math.ceil(400 / L.specWidthUnits) + 15 * Math.ceil(606 / L.noteWidthUnits), hL + ' need>=' + (18 + 14 * Math.ceil(400 / L.specWidthUnits) + 15 * Math.ceil(606 / L.noteWidthUnits)));
  const XH = clone(X0); XH.items = Array.from({ length: 50 }, (_, i) => ({ lid: 'h' + i, desc: 'Item ' + i, unit: '式', qty: 1, unitPrice: 1, cost: 0, spec: '字'.repeat(200), note: '字'.repeat(300) }));
  const xh = await G.sheetXml(ROOT, XH);
  t('5.9 50 列每列最長說明＋備註：每列高度都 ≤ 409，檔案仍可被 SheetJS 讀回（50 列都在）', Array.from({ length: 50 }, (_, i) => rowHt(xh.xml, R1 + i)).every((h) => h <= 409 && h > 21) && firstSheet(xh.buf)['C66'].v.startsWith('Item 49'));
  t('5.10 空品名＋說明：不產生開頭空行（第一段就是說明）', await (async () => { const Xe = clone(X0); Xe.items[0].desc = ''; Xe.items[0].spec = '只有說明'; const c = cellC((await G.sheetXml(ROOT, Xe)).xml, R1); return (c.match(/<r>/g) || []).length === 1 && !/>\n/.test(c); })());
  // 版面其餘不變：把 C17／C18 儲存格與這兩列的列高拿掉後，與沒有備註時逐位元相同
  const strip = (xml) => xml.replace(/<c r="C1[78]"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, '').replace(/(<row r="1[78]"[^>]*?)\sht="[^"]*"/g, '$1');
  t('5.11 其餘版面逐位元不變：拿掉品名儲存格與列高後，整份 sheet1.xml 與沒有說明／備註時完全相同（合併、隱藏列、金額欄、總額公式、Remarks 都沒動）', strip(x1.xml) === strip(xb.xml) && strip(x3.xml) === strip(xb.xml) && strip(xl.xml) === strip(xb.xml));
  const XK = clone(X0); XK.items = [{ lid: 't', kind: 'title', desc: 'Part A', spec: 'x', note: 'y' }, XK.items[0], { lid: 's', kind: 'subtotal', desc: '', spec: 'x', note: 'y' }];
  const XK0 = clone(XK); XK0.items.forEach((it) => { delete it.spec; delete it.note; });
  t('5.12 分組標題／小計列帶了 spec／note（不該有）→ 輸出與沒帶時完全相同', (await G.sheetXml(ROOT, XK)).xml === (await G.sheetXml(ROOT, XK0)).xml);
  const pnl = async (q) => { const b = await PNL.buildQuotePnlExcel(q, { classCodes: ['software'], requestedBy: 'x', issueDate: '2026-10-09', contingencyPct: 0 }); const z = await JSZip.loadAsync(b); const out = []; for (const n of Object.keys(z.files).sort()) if (/\.xml$/.test(n)) out.push(n + await z.file(n).async('string')); return out.join('\n'); };
  const XP = mkQuote({ items: mkQuote().items.map((x) => Object.assign({}, x, { cost: 100, cat: 'software' })) });
  t('5.13 毛利分析（內部）Excel 不受影響：有說明／備註的單與沒有的單輸出逐位元相同，也不會丟例外', await pnl(withNotes(XP)) === await pnl(XP));
  t('5.14 40 張固定單全加上說明／備註：給客戶的 Excel 都能產生、XML 良構（SheetJS 讀得回）、每個品項的說明與備註都在儲存格文字裡', await (async () => {
    for (const q of fx) {
      const w = withNotes(q); const r = await G.sheetXml(ROOT, w);
      const ws = firstSheet(r.buf);
      let row = L.itemFirst;
      for (const it of w.items.slice(0, 50)) { if (!it.kind) { const v = ws['C' + row].v; if (!v.includes(it.spec) || !v.includes('備註：' + it.note)) return false; } row++; }
    }
    return true;
  })());

  // ═════════════════ 6) 給客戶的預覽 HTML（PDF 用同一份）═════════════════
  const P = G.loadPreview(ROOT);
  const info = { issueDate: '2026-10-09', issuer: { name: 'S', phone: '1', ext: '2', mobile: '3' }, remarks: ['1.a'] };
  const pv0 = P.build(X0, info);
  const pv1 = P.build(X1, info);
  t('6.1 沒有說明／備註：品名欄只有品名（沒有 qpv-spec／qpv-note 標記）', !/qpv-spec|qpv-note/.test(pv0) && pv0.includes('<td class="desc">導入顧問</td>'));
  t('6.2 有說明＋備註：品名下面依序是 qpv-spec（說明）與 qpv-note（<b>備註：</b>＋內容）；只有說明的列沒有 qpv-note', pv1.includes('<td class="desc">導入顧問<div class="qpv-spec">含需求訪談與系統安裝</div><div class="qpv-note"><b>備註：</b>客戶要求週末施工</div></td>') && pv1.includes('<td class="desc">教育訓練<div class="qpv-spec">僅說明</div></td>'));
  const pv3 = P.build(X3, info);
  t('6.3 全部跳脫：<img onerror>、& 和引號都是純文字（沒有原始標籤）；公式字樣原樣顯示', !/<img src=x/.test(pv3) && pv3.includes('&lt;img src=x onerror=alert(1)&gt; &amp; &quot;q&quot;') && pv3.includes('+SUM(A1)') && pv3.includes('@cmd'));
  const X63 = clone(X0); X63.items[0].spec = '<img src=y onerror=alert(2)> & "s"'; X63.items[0].note = '<svg onload=alert(3)>';
  const pv63 = P.build(X63, info);
  t('6.3b 說明與備註各自跳脫（XSS 字樣：<img onerror>、<svg onload> 都只是文字）', !/<img src=y/.test(pv63) && !/<svg/.test(pv63) && pv63.includes('<div class="qpv-spec">&lt;img src=y onerror=alert(2)&gt; &amp; &quot;s&quot;</div>') && pv63.includes('<b>備註：</b>&lt;svg onload=alert(3)&gt;</div>'));
  t('6.4 分組標題／小計列帶 spec／note → 預覽與沒帶時逐位元相同', P.build(XK, info) === P.build(XK0, info));
  t('6.5 其餘版面不變：把 qpv-spec／qpv-note 區塊拿掉後與沒有說明／備註的預覽逐位元相同', pv1.replace(/<div class="qpv-spec">[\s\S]*?<\/div>/g, '').replace(/<div class="qpv-note">[\s\S]*?<\/div>/g, '') === P.build(Object.assign(clone(X0), { items: clone(X0.items) }), info));
  const cleanCases = ['', ' ', '  a  ', 'a\nb', 'a\r\n\tb', 'a\u0001b', 'a\u007fb\u0085c', 'x'.repeat(400), '字'.repeat(250), '\u{20BB7}'.repeat(250), 'a  b   c', 'a b', '﻿x', null, undefined, 5, {}, ['x'], '=1+1', '<b>x</b>', 'a' + String.fromCharCode(0x2028) + 'b' + String.fromCharCode(0x2029) + 'c', ' \t '];
  t('6.6 前端清洗函式 _qpvCleanText 與伺服器 cleanItemText 逐例相同（說明 200／備註 300 兩種上限）', cleanCases.every((v) => P.clean(v, 200) === QI.cleanItemText(v, 200) && P.clean(v, 300) === QI.cleanItemText(v, 300)));
  const pdf = read(path.join(ROOT, '_client/quote-pdf.js')), qpvSrc = read(path.join(ROOT, '_client/quote-preview.js'));
  t('6.7 PDF 與預覽同一份 HTML：quote-pdf.js 用 buildQuotePreviewHtml(q, info) 畫出紙張（說明／備註因此也在 PDF 裡）', /host\.innerHTML = buildQuotePreviewHtml\(q, info\)/.test(pdf));
  t('6.8 預覽樣式：說明灰色小字、備註棕紅色；與 Excel 的字色一致（#6b7280／#9a3412 ↔ FF6B7280／FF9A3412）', /\.qpv-items td\.desc \.qpv-spec \{[^}]*font-size: 0\.82em; color: #6b7280/.test(qpvSrc) && /\.qpv-items td\.desc \.qpv-note \{[^}]*font-size: 0\.9em; color: #9a3412/.test(qpvSrc) && /spec: 'FF6B7280', note: 'FF9A3412'/.test(read(path.join(ROOT, 'lib/quoteExcel.js'))));

  // ═════════════════ 7) 位元級相容（黃金值）═════════════════
  const D = await G.goldenDigests(ROOT);
  for (const k of ['hash', 'store', 'serialize', 'excel', 'preview']) {
    t('7.' + ({ hash: 1, store: 2, serialize: 3, excel: 4, preview: 5 })[k] + ' 黃金值 ' + ({ hash: 'contentHash／itemsSig／costLinesSig／lineStructureSig／structureSig', store: 'POST→PUT 後存進資料庫的品項（normalizeItems）', serialize: 'GET /api/quotations/:id 完整回應（含 preview、approval、perm；擁有者與管理員）', excel: '給客戶的 Excel（sheet1.xml）', preview: '給客戶的預覽 HTML' })[k] + '：沒有說明／備註的 40 張舊式單，新程式碼的輸出摘要與功能加入前（bc921ee）完全相同', D[k] === GOLDEN[k], D[k] + ' vs ' + GOLDEN[k]);
  }

  // ═════════════════ 8) 前端靜態紀律 ═════════════════
  const qs = read(path.join(ROOT, '_client/quote.js')), qa = read(path.join(ROOT, '_client/quote-approval.js'));
  t('8.1 表單列：說明／備註輸入框 value 經 escapeHtml、maxlength 200／300、有 aria-label 與 placeholder；兩者都空時收合成「＋ 說明／備註」連結', /class="qi-spec" value="' \+ escapeHtml\(specV\) \+ '" maxlength="200" aria-label="品項說明" placeholder="說明（選填，會印在客戶報價單）"/.test(qs) && /class="qi-note" value="' \+ escapeHtml\(noteV\) \+ '" maxlength="300" aria-label="品項備註" placeholder="備註：此項目的特殊需求（選填，會印在客戶報價單）"/.test(qs) && /notesOpen = !!\(specV \|\| noteV\)/.test(qs) && /＋ 說明／備註/.test(qs));
  const rq = qs.slice(qs.indexOf('function renderQuoteItems'), qs.indexOf('function quoteDragEnabled'));
  const kindPart = rq.slice(rq.indexOf("if (it.kind === 'title')"), rq.indexOf('seq++;'));
  t('8.2 分組標題／小計列的 HTML 沒有說明／備註輸入框', !/qi-spec|qi-note/.test(kindPart) && kindPart.length > 200);
  const rd = qs.slice(qs.indexOf('function readQuoteItems'), qs.indexOf('function updateQuoteTotals'));
  t('8.3 readQuoteItems：一般品項一律送 spec／note 字串（空字串＝清除）；標題／小計列不送', /it\.spec = specEl\.value\.trim\(\)/.test(rd) && /it\.note = noteEl\.value\.trim\(\)/.test(rd) && !/spec|note/.test(rd.slice(rd.indexOf('if (row.dataset.kind)'), rd.indexOf('const it = {'))));
  t('8.4 展開連結點擊：一次點擊展開兩個輸入框並把焦點放在說明欄', /qi-notes-toggle'\)\.forEach/.test(qs) && /sp\.focus\(\)/.test(qs) && /box\.hidden = false/.test(qs));
  t('8.5 備註不是必填：quote-steps.js 的品項檢查只看品名、數量、單價（沒有提到 spec／note）', !/qi-spec|qi-note/.test(read(path.join(ROOT, '_client/quote-steps.js'))));
  const dlg = qs.slice(qs.indexOf('async function openQuoteNotesDialog'), qs.indexOf('// ── 我的聯絡資訊'));
  t('8.6 「說明／備註」小視窗：打 PUT …/item-notes、只送有改的欄位（比對 data-orig）、內容經 escapeHtml、成功後更新列表快取；錯誤顯示在視窗內', /\/item-notes', \{ method: 'PUT'/.test(dlg) && /data-orig/.test(dlg) && /if \(sp\.value\.trim\(\) !== sp\.getAttribute\('data-orig'\)\.trim\(\)\) o\.spec/.test(dlg) && /escapeHtml\(it\.spec \|\| ''\)/.test(dlg) && /escapeHtml\(it\.note \|\| ''\)/.test(dlg) && /allQuotations\[i\] = j/.test(dlg) && /errEl\.textContent = j\.error/.test(dlg));
  t('8.7 小視窗明講「不影響簽核、客戶單上的文字立即改變、每次修改都留紀錄」', /不影響金額與簽核（不需重新簽核），但<b>客戶報價單上印出的文字會立即改變<\/b>，每次修改都會留下紀錄/.test(dlg));
  const ab = qs.slice(qs.indexOf('function _qActionButtons'), qs.indexOf('function renderQuoteList'));
  t('8.8 列表按鈕 📝 說明／備註：只有擁有者或管理員、且簽核狀態是簽核中／已核准時才出現；事件委派有 notes 分支', /\(p\.isOwner \|\| me\.isAdmin\) && a && \(a\.state === 'pending' \|\| a\.state === 'approved'\)/.test(ab) && /case 'notes':\s+p = openQuoteNotesDialog\(id\)/.test(qs));
  t('8.9 簽核歷程動作標籤有 ITEM_TEXT_EDIT；已核准提示與 WILL_VOID 對話框都註明說明／備註不會作廢核准', /ITEM_TEXT_EDIT: '修改品項說明／備註（不影響核准）'/.test(qa) && /品項的「說明」「備註」除外：改這兩欄不會作廢核准/.test(qs) && /只改品項的說明／備註不會走到這一步/.test(qs) && /（品項的說明與備註除外）/.test(qr));
  t('8.10 沒有 eval／new Function／document.write 新增；小視窗關閉時移除 keydown 監聽', !/\beval\(|new Function|document\.write/.test(dlg) && /document\.removeEventListener\('keydown', onKey, true\)/.test(dlg));

  const pl = qs.slice(qs.indexOf('const payloadItems = items.map'), qs.indexOf('const payload = {'));
  t('8.11 儲存時送出的 payload 品項（payloadItems）帶 spec／note 字串（曾經漏掉：表單看得到、存檔卻被丟掉）；標題／小計列不帶', /if \(typeof it\.spec === 'string'\) o\.spec = it\.spec;/.test(pl) && /if \(typeof it\.note === 'string'\) o\.note = it\.note;/.test(pl) && !/spec|note/.test(pl.slice(pl.indexOf('if (quoteIsKindRow(it))'), pl.indexOf('const o = {'))));
  t('8.12 窄螢幕（≤620px）：品名欄至少 220px 讓說明／備註輸入框夠寬；簽核歷程表動作欄有最小寬度、長文字 overflow-wrap:anywhere（不撐出橫向捲軸）', /@media \(max-width: 620px\) \{ #quoteItemsBody tr:not\(\[data-kind\]\) td:nth-child\(3\) \{ min-width:220px; \} \}/.test(qs) && /\.qap-hist td:nth-child\(3\) \{ min-width: 7\.5em; \}/.test(qa) && /\.qap-hist td\.desc \{ overflow-wrap: anywhere; \}/.test(qa) && /class="qap-table qap-hist"/.test(qa) && /\.qap-tdiff-row \{[^}]*overflow-wrap: anywhere/.test(qa));
  // ═════════════════ 9) 追蹤：歷程差異、q.textEdits 標記、橫幅權限、站內通知 ═════════════════
  const SECRET = 'SECRETTEXT-9f3a';
  const chain = (approvedSteps, pendingStep, extra) => {   // 自訂簽核鏈：approvedSteps＝[{tier,by}]，pendingStep＝{tier,assignee?}（有的話狀態是 pending）
    const q = mkQuote(extra);
    const steps = approvedSteps.map((x) => ({ tier: x.tier, label: x.tier, status: 'approved', by: x.by, at: '2026-10-08T02:00:00.000Z', assignee: x.assignee || x.by, comment: '' }));
    if (pendingStep) steps.push({ tier: pendingStep.tier, label: pendingStep.tier, status: 'pending', assignee: pendingStep.assignee || null, by: null, at: null, comment: '' });
    q.approval = Object.assign(approvalOf(q, pendingStep ? 'pending' : 'approved'), { steps, cur: pendingStep ? steps.length - 1 : steps.length });
    return q;
  };
  const noteTargets = (e) => e.notes.map((n) => n[0]).sort().join();
  // 9a) 結構化差異與標記
  env = envOf(stateQuote('approved'));
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', spec: '新說明', note: '長'.repeat(150) }] });
  let hq = stored(env).approval.history.slice(-1)[0];
  t('9.1 歷程項目帶結構化明細 meta.textEdits：[{lid,name,spec:[舊,新],note:[舊,新]}]，文字是完整的（備註 150 字沒有被截短），comment（簡述）才截短', hq.action === 'ITEM_TEXT_EDIT' && J(hq.meta.textEdits[0].spec) === J(['', '新說明']) && hq.meta.textEdits[0].note[1].length === 150 && hq.meta.textEdits[0].name === '導入顧問' && hq.meta.textEdits[0].lid === 'i1' && hq.meta.more === 0 && /…/.test(hq.comment), J(hq.meta).slice(0, 200));
  g = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
  const gh = g.j.approval.history.slice(-1)[0];
  t('9.2 GET 序列化：ITEM_TEXT_EDIT 的歷程多 items（舊→新）；其他動作（SUBMIT／APPROVE）的歷程欄位逐位元不變（只有 at／byName／action／comment／tier）', J(gh.items) === J(hq.meta.textEdits) && g.j.approval.history.slice(0, 2).every((h) => J(Object.keys(h)) === J(['at', 'byName', 'action', 'comment', 'tier'])));
  t('9.3 q.textEdits 摘要 {count,firstAt,lastAt,lastBy}；序列化 approval.textEdits 只有 {count,firstAt,lastAt,lastByName}（不含 username、不含文字）', J(Object.keys(stored(env).textEdits)) === J(['count', 'firstAt', 'lastAt', 'lastBy']) && stored(env).textEdits.count === 1 && J(Object.keys(g.j.approval.textEdits)) === J(['count', 'firstAt', 'lastAt', 'lastByName']) && g.j.approval.textEdits.lastByName === 'Owner1' && !J(g.j.approval.textEdits).includes('新說明'));
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i2', note: '第二次' }] });
  t('9.4 再改一次：count=2、firstAt 不變、lastAt 更新到第二筆歷程的時間', stored(env).textEdits.count === 2 && stored(env).textEdits.firstAt === hq.at && stored(env).textEdits.lastAt === stored(env).approval.history.slice(-1)[0].at);
  // 標記與雜湊
  const withTE = clone(stored(env)); const woTE = clone(stored(env)); delete woTE.textEdits;
  t('9.5 textEdits 不進任何雜湊／簽章：contentHash／itemsSig／costLinesSig／lineStructureSig／structureSig 有沒有 textEdits 都一樣；核准仍有效', QA.contentHash(withTE) === QA.contentHash(woTE) && QA.itemsSig(withTE) === QA.itemsSig(woTE) && QA.costLinesSig(withTE) === QA.costLinesSig(woTE) && QA.structureSig(withTE) === QA.structureSig(woTE) && QA.contentHash(stored(env)) === stored(env).approval.hash);
  // 狀態規則：只有簽核中／已核准才記
  const rows9 = [];
  for (const [label, q0] of [['草稿', stateQuote(null)], ['被駁回', (() => { const x = stateQuote('pending'); x.approval.state = 'returned'; return x; })()], ['作廢後（none）', (() => { const x = stateQuote(null); x.approval = { state: 'none', history: [{ at: 'x', by: 'own1', action: 'INVALIDATE', comment: '', tier: '' }], steps: [] }; return x; })()]]) {
    const e9 = envOf(q0);
    const hl = (stored(e9).approval && stored(e9).approval.history || []).length;
    const r9 = await e9.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
    const pf = await e9.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(stored(e9), (a) => { a[1].note = SECRET; }) });
    rows9.push([label, r9.s, pf.s, 'textEdits' in stored(e9), (stored(e9).approval && stored(e9).approval.history || []).length === hl, e9.notes.length]);
  }
  t('9.6 草稿／被駁回／已作廢的單改說明／備註：兩條路徑都成功寫入，但不記歷程、不設 q.textEdits、不發通知（只有稽核紀錄）', rows9.every((r) => r[1] === 200 && r[2] === 200 && r[3] === false && r[4] === true && r[5] === 0), J(rows9));
  env = envOf(stateQuote('approved'));
  await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(stored(env), (a) => { a[0].qty = 9; a[0].note = SECRET; }), confirmVoid: true });
  t('9.7 已核准但同時動到錢（確認作廢）：核准作廢（state none），不記說明／備註歷程、不設 textEdits、不發 ITEM_TEXT 通知（走 quote_voided）', stored(env).approval.state === 'none' && !('textEdits' in stored(env)) && !stored(env).approval.history.some((h) => h.action === 'ITEM_TEXT_EDIT') && env.notes.every((n) => n[1] !== 'quote_text_edited') && env.notes.some((n) => n[1] === 'quote_voided'));
  // 一般路徑（已核准、完整表單 PUT）
  env = envOf(stateQuote('approved'));
  u = await env.call('own1', 'PUT', '/api/quotations/:id', { id: 'Q1' }, { items: formItems(stored(env), (a) => { a[0].spec = '一般路徑說明'; a[2].note = '一般路徑備註'; }) });
  hq = stored(env).approval.history.slice(-1)[0];
  t('9.8 一般路徑（已核准、完整表單 PUT）也寫同樣的 ITEM_TEXT_EDIT（結構化舊→新、兩個品項）＋設 textEdits＋稽核 UPDATE_QUOTATION＋通知', u.s === 200 && hq.action === 'ITEM_TEXT_EDIT' && hq.meta.textEdits.length === 2 && J(hq.meta.textEdits[1].note) === J(['', '一般路徑備註']) && stored(env).textEdits.count === 1 && env.logs.slice(-1)[0][0] === 'UPDATE_QUOTATION' && /備註 ∅→一般路徑備註/.test(env.logs.slice(-1)[0][3]) && noteTargets(env) === 'mgr1');
  // 送簽新一輪：標記清掉、歷程保留
  env = envOf(stateQuote(null, { items: mkQuote().items.map((x) => Object.assign({}, x, { cost: 1000 })) }));
  stored(env).approval = { state: 'none', history: [{ at: '2026-10-01T00:00:00.000Z', by: 'own1', action: 'INVALIDATE', comment: '', tier: '', meta: null }, { at: '2026-10-01T00:00:01.000Z', by: 'own1', action: 'ITEM_TEXT_EDIT', comment: 'x', tier: '', meta: { textEdits: [], more: 0 } }], steps: [] };
  stored(env).textEdits = { count: 3, firstAt: 'a', lastAt: 'b', lastBy: 'own1' };
  u = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' }, {});
  t('9.9 重新送簽（新一輪簽核）：q.textEdits 清掉（簽核人這輪簽的就是目前文字）；先前的歷程（含舊的 ITEM_TEXT_EDIT）保留', u.s === 200 && !('textEdits' in stored(env)) && stored(env).approval.history.filter((h) => h.action === 'ITEM_TEXT_EDIT').length === 1 && stored(env).approval.state === 'pending', u.s + J(u.j && u.j.error));
  // 序列化：只有簽核中／已核准才輸出 textEdits；舊單沒有
  const ser = async (q0, user) => (await envOf(q0).call(user, 'GET', '/api/quotations/:id', { id: 'Q1' }));
  const mkTE = (state) => { const q0 = stateQuote(state); q0.textEdits = { count: 2, firstAt: 'a', lastAt: 'b', lastBy: 'own1' }; return q0; };
  t('9.10 橫幅資料可見性規則：簽核中／已核准有 approval.textEdits；草稿（approval=null）、被駁回、作廢後（state none）的單即使殘留 q.textEdits 也不輸出；舊單（沒有 q.textEdits）不輸出', (await ser(mkTE('approved'), 'own1')).j.approval.textEdits.count === 2 && (await ser(mkTE('pending'), 'own1')).j.approval.textEdits.count === 2
    && !('textEdits' in (await ser(mkTE('returnedX') && (() => { const x = mkTE('pending'); x.approval.state = 'returned'; return x; })(), 'own1')).j.approval) && (await ser((() => { const x = mkTE(null); return x; })(), 'own1')).j.approval === null && !('textEdits' in (await ser(stateQuote('approved'), 'own1')).j.approval));
  const seers = {};
  for (const who of ['own1', 'admin1', 'mgr1', 'gm1', 'sec1']) { const r = await ser(mkTE('approved'), who); seers[who] = r.s === 200 && r.j.approval && r.j.approval.textEdits ? r.j.approval.textEdits.count : (r.s + ':' + (r.j && r.j.approval ? 'noTE' : 'noApproval')); }
  t('9.11 能看簽核資訊的角色都看得到橫幅資料（擁有者、管理員、主管、總經理、秘書；以各自原本的可見範圍為準）', Object.values(seers).every((v) => v === 2 || /^(403|404)/.test(String(v))) && seers.own1 === 2 && seers.admin1 === 2, J(seers));
  const consQ = mkTE('approved'); consQ.products = ['PC']; consQ.costBy = 'cons1'; consQ.costFlow = { state: 'filled', by: 'cons1', requestedAt: 'x', filledAt: 'x', note: '', sig: null, consultantWrote: true };
  const rc = await ser(consQ, 'cons1'); const rx = await ser(mkTE('approved'), 'x1');
  t('9.12 不洩漏：只負責填成本的顧問（approval 一律 null）完全看不到 textEdits／歷程差異；無關的業務連單都讀不到（403/404）', rc.s === 200 && rc.j.approval === null && !J(rc.j).includes('textEdits') && [403, 404].includes(rx.s) && !J(rx.j).includes('textEdits'));
  // 9b) 通知
  const textNotes = (e) => (e.data.notifications || []).filter((n) => n.type === 'quote_text_edited');
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }, { tier: 'gm', by: 'gm1' }], null));
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', spec: SECRET, note: SECRET + '-n' }, { lid: 'i2', note: SECRET + '-2' }] });
  t('9.13 已核准、簽核鏈 mgr1→gm：收件人＝簽過的主管與總經理（mgr1、gm1），不含修改的業務自己；一人一則', noteTargets(env) === 'gm1,mgr1' && textNotes(env).length === 2);
  const nt = textNotes(env)[0];
  t('9.14 通知內容：type quote_text_edited、refId＝報價單 id、未讀；標題與內文有單號、修改人名稱、被改的品項數（2 項）；**完全沒有說明／備註的文字**（也沒有品名）', nt.refId === 'Q1' && nt.read === false && /QU-1/.test(nt.title) && /核准後說明／備註被修改/.test(nt.title) && /QU-1/.test(nt.body) && /Owner1/.test(nt.body) && /修改了 2 項品項/.test(nt.body) && !J(textNotes(env)).includes(SECRET) && !/導入顧問|教育訓練|軟體授權/.test(J(textNotes(env))), J(nt));
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }], { tier: 'gm' }));
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
  t('9.15 簽核中（mgr1 已簽、輪到 gm）：收件人＝已簽的 mgr1＋目前的簽核人（總經理名冊 gm1）；董事長 ch1 沒有被通知；標題是「簽核中…」', noteTargets(env) === 'gm1,mgr1' && textNotes(env).every((n) => /簽核中說明／備註被修改/.test(n.title)));
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }, { tier: 'gm', by: 'gm1' }], { tier: 'board' }));
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
  t('9.16 有董事會關卡：另外通知代核秘書（角色 secretary，與既有簽核通知用同一個 boardProxySet）', noteTargets(env) === 'gm1,mgr1,sec1', noteTargets(env));
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }], null));
  await env.call('admin1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
  t('9.17 管理員代改：業務擁有者 own1 與簽過的 mgr1 都收到；修改人（管理員）自己沒有；內文的修改人是管理員的顯示名稱', noteTargets(env) === 'mgr1,own1' && /Admin/.test(textNotes(env)[0].body));
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }], null)); env.auth.users.find((x) => x.username === 'mgr1').active = false;
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
  t('9.18 已停用的帳號不通知', noteTargets(env) === '' && textNotes(env).length === 0);
  // 合併
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }, { tier: 'gm', by: 'gm1' }], null));
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'a1' }] });
  const first = textNotes(env).map((n) => n.createdAt).join();
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'a2' }, { lid: 'i2', note: 'b2' }] });
  const tn = textNotes(env), mg = tn.find((n) => n.to === 'mgr1');
  t('9.19 合併：同一位修改人、同一張單、同一位收件人的「未讀」通知再改一次 → 不新增，更新品項累計（1+2＝3）、修改次數 2、時間，移到最上面；內文「累計 2 次修改」', tn.length === 2 && mg.itemCount === 3 && mg.editCount === 2 && /修改了 3 項品項/.test(mg.body) && /累計 2 次修改/.test(mg.body) && mg.read === false && env.data.notifications[0].type === 'quote_text_edited' && env.notes.length === 2, J(mg));
  tn.find((n) => n.to === 'mgr1').read = true;
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'a3' }] });
  const tn2 = textNotes(env);
  t('9.20 已讀之後再改 → 對該收件人新增一則（不改已讀的）；另一位還沒讀的（gm1）繼續合併成 3 次', tn2.length === 3 && tn2.filter((n) => n.to === 'mgr1').length === 2 && tn2.find((n) => n.to === 'gm1').editCount === 3 && tn2.filter((n) => n.to === 'mgr1' && n.read === true).length === 1);
  await env.call('admin1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: 'by admin' }] });
  const tn3 = textNotes(env);
  t('9.21 不同的修改人不合併：管理員改一次 → gm1 另外多一則（by＝admin1），不會併進業務那則', tn3.filter((n) => n.to === 'gm1').length === 2 && tn3.filter((n) => n.to === 'gm1' && n.by === 'admin1').length === 1 && tn3.filter((n) => n.to === 'gm1' && n.by === 'own1')[0].editCount === 3);
  t('9.22 合併不會漏存（db.save 有呼叫）、通知總數不超過上限邏輯（這裡只確認有 save）：env.saves ＞ 0', env.saves > 0);
  // 前端／伺服器靜態：圖示與深連結
  const appSrc = read(path.join(ROOT, '_client/app.js'));
  t('9.23 鈴鐺：quote_text_edited 有圖示（📝）；點擊走既有 quote_ 開頭的路徑（開簽核面板）；server 的 notificationUrlFor 對 quote_* 一律回 /index.html#quote:<id>', /quote_text_edited:'📝'/.test(appSrc) && /startsWith\('quote_'\) && item\.dataset\.ref\) openQuoteFromNotification/.test(appSrc) && /startsWith\('quote_'\)\) return refId \? `\/index\.html#quote:\$\{refId\}/.test(srv));
  // 前端函式（vm）：差異渲染與跳脫、橫幅文字
  const qaSrc9 = read(path.join(ROOT, '_client/quote-approval.js')), qsSrc9 = read(path.join(ROOT, '_client/quote.js'));
  const cut = (src, a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); if (i < 0 || j < 0) throw new Error('找不到 ' + a); return src.slice(i, j); };
  const vmCtx = { escapeHtml: (x) => (x === null || x === undefined ? '' : String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/\'/g, '&#039;')) };
  vmCtx.e = vmCtx.escapeHtml; vmCtx.fmtTime = (v) => 'T(' + v + ')'; vmCtx._qFmtTime = vmCtx.fmtTime;
  require('vm').createContext(vmCtx);
  require('vm').runInContext(cut(qaSrc9, 'function buildTextDiff', 'function buildHistorySection') + cut(qsSrc9, 'function quoteTextEditText', '/** 簽核階段：{ key, cls, label, title } */'), vmCtx);
  const XSSD = '<img src=x onerror="window.__pwned=1">';
  const d1 = vmCtx.buildTextDiff({ comment: 'c', items: [{ name: XSSD, spec: ['', XSSD], note: [XSSD, ''] }, { name: '品名B', note: ['舊', '新'] }], itemsMore: 3 });
  t('9.24 歷程差異渲染：每個品項一塊（品名＋說明／備註「舊 → 新」）；空白顯示「（空）」；<img onerror> 一律跳脫（輸出沒有原始標籤）；超出筆數顯示「…另 3 項」', !/<img/.test(d1) && d1.includes('&lt;img src=x onerror=&quot;window.__pwned=1&quot;&gt;') && (d1.match(/qap-tdiff-item/g) || []).length === 2 && />說明<\/span><span class="old"><span class="emp">（空）<\/span><\/span> → <span class="new">&lt;img/.test(d1) && /備註<\/span><span class="old">舊<\/span> → <span class="new">新<\/span>/.test(d1) && /<span class="new"><span class="emp">（空）<\/span><\/span>/.test(d1) && /…另 3 項/.test(d1), d1.slice(0, 400));
  t('9.25 舊歷程（沒有結構化明細）退回顯示 comment（跳脫）；沒有改到的欄位（只有備註）不出現「說明」列', vmCtx.buildTextDiff({ comment: '<b>x</b>' }) === '&lt;b&gt;x&lt;/b&gt;' && !/>說明</.test(vmCtx.buildTextDiff({ items: [{ name: 'n', note: ['a', 'b'] }] })));
  const mkA = (state, last, stepAt) => ({ approval: { state, steps: [{ status: 'approved', at: stepAt }], textEdits: { count: 2, firstAt: 'x', lastAt: last, lastByName: '王<b>' } } });
  t('9.26 橫幅文字：簽核中→「簽核中說明／備註已被修改」；已核准且最後修改晚於核准→「核准後…」；已核准但只在簽核期間改→「簽核期間…曾被修改」；沒有記錄或草稿→空字串；含最後修改日期與人', /^簽核中說明／備註已被修改（最後修改：T\(2026-10-09\)　王<b>）$/.test(vmCtx.quoteTextEditText(mkA('pending', '2026-10-09', '')))
    && /^核准後說明／備註已被修改/.test(vmCtx.quoteTextEditText(mkA('approved', '2026-10-09', '2026-10-08'))) && /^簽核期間說明／備註曾被修改/.test(vmCtx.quoteTextEditText(mkA('approved', '2026-10-07', '2026-10-08'))) && vmCtx.quoteTextEditText({ approval: { state: 'approved', steps: [] } }) === '' && vmCtx.quoteTextEditText({ approval: { state: 'none', textEdits: { count: 1 } } }) === '' && vmCtx.quoteTextEditText({}) === '');
  t('9.27 橫幅畫面接線：簽核面板有琥珀色橫幅（qap-alert warn）＋「查看差異」按鈕（data-act=textdiff，捲到歷程並閃一下）；歷程區有 id=qapHistory；橫幅文字經 e() 跳脫；列表有徽章 quoteTextEditBadge（title 跳脫）；表單簽核狀態區有「查看差異」連結開簽核面板', /qap-alert warn qap-textedit/.test(qaSrc9) && /data-act="textdiff"/.test(qaSrc9) && /id="qapHistory"/.test(qaSrc9) && /act === 'textdiff'/.test(qaSrc9) && /\$\{e\(tn\)\}/.test(qaSrc9) && /function quoteTextEditBadge/.test(qsSrc9) && /title="' \+ escapeHtml\(t\)/.test(qsSrc9) && /data-qte="\$\{escapeHtml\(q\.id \|\| ''\)\}"/.test(qsSrc9) && /qOpenApproval\(a\.getAttribute\('data-qte'\)\)/.test(qsSrc9));
  t('9.28 只有簽核資訊的接線：橫幅／徽章只靠 q.approval.textEdits（approval 為 null 的視角＝空字串，不另外打 API）', /const a = q && q\.approval, te = a && a\.textEdits;/.test(qsSrc9) && /const a = q && q\.approval, te = a && a\.textEdits;/.test(qaSrc9));

  // 9c) 代核秘書（boardProxySet）只收極簡內文（比照董事會通知：只有單號）；董事會關卡還在 waiting 時不通知
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }, { tier: 'gm', by: 'gm1' }], { tier: 'board' }));
  env.auth.users.push({ username: 'sec2', role: 'secretary', active: true, displayName: 'OtherBuSec', bu: ['ERP'] });
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
  const secN = textNotes(env).filter((n) => n.to === 'sec1' || n.to === 'sec2'), fullN = textNotes(env).filter((n) => n.to === 'mgr1' || n.to === 'gm1');
  t('9.29 董事會關卡輪到時：所有秘書（含非本部門的 sec2）收到極簡通知——內文只有單號與品項數，沒有客戶名、業務名、修改人名稱；也沒有記錄修改人帳號（沒有 by 欄位）', secN.length === 2 && secN.every((n) => n.body === 'QU-1　品項的說明／備註被修改了（共 1 項），不影響簽核。' && !/TestCo|Owner1|own1/.test(J(n)) && !('by' in n) && n.minimal === true), J(secN));
  t('9.30 已簽過的人（mgr1、gm1）仍是完整內文（客戶名、業務名、修改人）——他們本來就收得到帶客戶名的簽核通知', fullN.length === 2 && fullN.every((n) => /TestCo/.test(n.body) && /Owner1/.test(n.body) && n.by === 'own1'));
  await env.call('admin1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i2', note: 'x' }] });
  const secN2 = textNotes(env).filter((n) => n.to === 'sec1');
  t('9.31 極簡通知合併不分修改人（不記修改人）：管理員再改一次 → sec1 仍只有一則，品項累計 2；內文仍不含任何名稱', secN2.length === 1 && secN2[0].itemCount === 2 && secN2[0].editCount === 2 && !/TestCo|Owner1|Admin|own1|admin1/.test(J(secN2[0])) && secN2[0].body === 'QU-1　品項的說明／備註被修改了（共 2 項），不影響簽核。', J(secN2));
  // 重現（驗證報告 secleak）：董事會關卡還在 waiting（mgr1 還在簽）、另一個部門的秘書
  { const q0 = mkQuote(); q0.company = 'SECRET-CLIENT'; q0.approval = Object.assign(approvalOf(q0, 'pending'), { steps: [{ tier: 'mgr1', label: 'm', assignee: 'mgr1', status: 'pending' }, { tier: 'board', label: 'b', status: 'waiting' }], cur: 0 });
    const e0 = envOf(q0); e0.auth.users.push({ username: 'sec2', role: 'secretary', active: true, displayName: 'OtherBuSec', bu: ['ERP'] });
    await e0.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
    t('9.32 重現：董事會關卡還在 waiting（尚未輪到）→ 秘書（本部門 sec1、其他部門 sec2）完全沒收到通知；收到的人（目前簽核人 mgr1）才有客戶名；沒有任何通知把客戶名送給 ERP 部門的秘書', noteTargets(e0) === 'mgr1' && !textNotes(e0).some((n) => (n.to === 'sec1' || n.to === 'sec2') && /SECRET-CLIENT/.test(J(n))) && (await e0.call('sec2', 'GET', '/api/quotations/:id', { id: 'Q1' })).s !== 200); }
  env = envOf(chain([{ tier: 'mgr1', by: 'mgr1' }, { tier: 'gm', by: 'gm1' }, { tier: 'board', by: 'sec1' }], null));
  env.auth.users.push({ username: 'sec2', role: 'secretary', active: true, displayName: 'OtherBuSec', bu: ['ERP'] });
  await env.call('own1', 'PUT', NOTES, { id: 'Q1' }, { items: [{ lid: 'i1', note: SECRET }] });
  const s2n = textNotes(env).find((n) => n.to === 'sec2'), s1n = textNotes(env).find((n) => n.to === 'sec1');
  t('9.34 董事會關卡已完成（已核准）：沒簽過的秘書（sec2，非本部門）仍收極簡內文；親自代核簽過的秘書（sec1）是簽核人，收完整內文', !!s2n && s2n.minimal === true && !/TestCo|Owner1/.test(J(s2n)) && !!s1n && /TestCo/.test(s1n.body) && s1n.by === 'own1');
  // 9d) 信件：事件的品項只有 desc／unit／qty，永遠不含 spec／note
  { const QM = require(path.join(ROOT, 'lib/mail/quoteMail.js'));
    const qm = mkQuote(); qm.items.forEach((x) => { x.spec = SECRET + '-spec'; x.note = SECRET + '-note'; });
    const rows = QM.itemsOf(qm);
    t('9.33 信件事件的品項（E2）只有 desc／unit／qty；即使品項有說明／備註，事件物件也完全不含（lib/mail 沒有改動，這裡鎖住「日後有人加 note 欄位」）', rows.length === 3 && rows.every((r) => Object.keys(r).every((k) => ['desc', 'unit', 'qty'].includes(k))) && !J(rows).includes(SECRET) && !/\b(?:it|x|item|row)\.(?:spec|note)\b/.test(read(path.join(ROOT, 'lib/mail/quoteMail.js'))), J(rows).slice(0, 200));
  }

  // ═════════════════ 結果 ═════════════════
  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + String(x).slice(0, 600) : '')); ok ? pass++ : fail++; });
  console.log(`\n品項說明／備註檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
}
run().catch((e) => { res.filter((r) => !r[1]).forEach((r) => console.log('FAIL ' + r[0] + (r[2] ? '  <- ' + String(r[2]).slice(0, 400) : ''))); console.error('例外（後面的檢查沒跑到）', e.stack); process.exit(2); });
