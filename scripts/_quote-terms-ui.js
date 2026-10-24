'use strict';
/**
 * check-quote-terms.js 的 [ui] 群組：License／MA 服務期間的前端（headless Edge + CDP；真實滑鼠與鍵盤事件）。不單獨執行。
 *   · 開關關閉：品項列／預覽 HTML／簽核面板／儲存請求本文，與基準程式碼（git archive 的舊版前端檔，用 CDP 攔截替換）逐位元相同
 *   · 開關開啟：類型下拉與品名預選、起迄日輸入（真實鍵盤）、一年／三年快捷、共 N 個月、紅字提示、逐步進度、拖曳／▲▼／刪除／新增保留期間、牌價簿加入、儲存重載
 *   · 核准後「修改服務期間」小視窗（權限、錯誤碼對應、稽核、通知、時間軸）、預覽、PDF 產生、Excel 內容、簽核面板唯讀顯示
 *   · XSS（品名與髒資料的期間含 <img onerror>／<script>）、四種外觀對比量測與截圖、手機寬度
 * 全部假資料、臨時 _s_* 帳號；PDF 不真的下載（攔截 jsPDF.save）；截圖存 scratchpad/ma1_shots/。
 * UI 變異（CTERM_MUT=1 由 check-quote-terms.js 呼叫 run(C, mutId)）：用 CDP 把前端檔換成改壞的版本，對應的測試必須變紅。
 */
const fs = require('fs');
const path = require('path');
const { withBrowser } = require('./_fd-cdp');

const J = JSON.stringify;
const Q = (sel) => `document.querySelector(${J(sel)})`;

// ───────── 對比度（文字色 vs 疊合後的背景色；沿祖先逐層疊合半透明底）─────────
const CONTRAST_FN = `(function(sel, dark){
  function parse(c){ var m=c.match(/rgba?\\(([^)]+)\\)/); if(!m) return [0,0,0,0]; var p=m[1].split(/[ ,\\/]+/).filter(Boolean).map(Number); return [p[0],p[1],p[2],p.length>3?p[3]:1]; }
  function over(top,bot){ var a=top[3]; return [top[0]*a+bot[0]*(1-a), top[1]*a+bot[1]*(1-a), top[2]*a+bot[2]*(1-a), 1]; }
  function lum(c){ var f=function(v){v/=255; return v<=0.03928? v/12.92 : Math.pow((v+0.055)/1.055,2.4)}; return 0.2126*f(c[0])+0.7152*f(c[1])+0.0722*f(c[2]); }
  var el=document.querySelector(sel); if(!el) return null;
  var chain=[]; for(var n=el;n&&n.nodeType===1;n=n.parentElement) chain.push(n);
  var bg=dark?[11,16,32,1]:[255,255,255,1];
  for(var i=chain.length-1;i>=0;i--){ var c=parse(getComputedStyle(chain[i]).backgroundColor); if(c[3]>0) bg=over(c,bg); }
  var fg=parse(getComputedStyle(el).color); fg=over(fg,bg);
  var L1=lum(fg), L2=lum(bg); var hi=Math.max(L1,L2), lo=Math.min(L1,L2);
  return Math.round(((hi+0.05)/(lo+0.05))*100)/100;
})`;
const contrast = (b, sel, dark) => b.ev(`${CONTRAST_FN}(${J(sel)}, ${!!dark})`);
const THEMES = ['glass-dark', 'glass-light', 'plain-light', 'plain-dark'];
async function theme(b, mode) {
  await b.ev(`(()=>{const c=document.body.classList; c.toggle('glass',${J(mode.startsWith('glass'))}); c.toggle('dark',${J(mode.endsWith('dark'))}); return true})()`);
  await b.sleep(350);
}

// ───────── 假資料 ─────────
const T0 = '2026-10-08T00:00:00.000Z';
function mkUsers(H) {
  return [
    H.mkUser('_s_admin', 'admin'), H.mkUser('_s_own', 'user', { bu: ['ERP'], supervisor: '_s_m1' }), H.mkUser('_s_other', 'user', { bu: ['ERP'] }),
    H.mkUser('_s_m1', 'manager1', { bu: ['ERP'] }), H.mkUser('_s_gm', 'executive'), H.mkUser('_s_sec', 'secretary'),
  ];
}
const PROD = 'SAP PCOE License MA';
const OPPS = [
  { id: 'o1', owner: '_s_own', stage: 'C', company: '測試公司', product: PROD, amount: '100', bu: 'ERP', createdAt: T0 },
  { id: 'o2', owner: '_s_own', stage: 'B', company: '第二公司', product: PROD, amount: '50', bu: 'ERP', createdAt: T0 },
];
function mkQ(id, no, over) {
  return Object.assign({
    id, quoteNo: no, owner: '_s_own', oppId: 'o1', company: '測試公司', contactName: '', phone: '', mobile: '', address: '', projectName: '專案甲', quoteDate: '2026-10-08', status: 'draft',
    createdAt: T0, updatedAt: T0, validUntil: '2026-10-30', products: [PROD], costBy: null, costFlow: { state: 'na' }, approval: null, discountType: 'none', discountValue: 0, items: [],
  }, over || {});
}
function approvedOf(QA, q) {
  return {
    state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: '2026-10-08T01:00:00.000Z', submittedBy: '_s_own', derived: null,
    steps: [{ tier: 'mgr1', label: '一級主管', status: 'approved', assignee: '_s_m1', by: '_s_m1', at: '2026-10-08T02:00:00.000Z', comment: '' }], cur: 1, board: null,
    history: [{ at: '2026-10-08T01:00:00.000Z', by: '_s_own', action: 'SUBMIT', comment: '', tier: '' }, { at: '2026-10-08T02:00:00.000Z', by: '_s_m1', action: 'APPROVE', comment: '', tier: 'mgr1' }],
  };
}
const XSS1 = '<img src=x onerror=window.__xss=1>', XSS2 = '<script>window.__xss=2</script>';
/** withTerms：false＝舊單（沒有任何期間欄位）；true＝含期間。XSS 與 BIG 只在開啟情境放 */
function seedData(R, withTerms) {
  const QA = require(path.join(R, 'lib/quoteApproval.js'));
  const T = (kind, from, to) => (withTerms ? { termKind: kind, termFrom: from, termTo: to } : {});
  const qd = mkQ('QD', 'QU-T-001', {
    items: [
      Object.assign({ lid: 'd1', desc: 'SAP License 授權', unit: '套', qty: 1, unitPrice: 600000, cost: 0 }, T('license', '2026-01-01', '2026-12-31')),
      Object.assign({ lid: 'd2', desc: '年度 MA 維護', unit: '式', qty: 2, unitPrice: 100000, cost: 0, spec: '含遠端支援' }, T('ma', '2026-03-01', '2027-02-28')),
      { lid: 'd3', desc: '教育訓練', unit: '式', qty: 1, unitPrice: 50000, cost: 0 },
      { lid: 'dt', kind: 'title', desc: 'Part B' }, { lid: 'ds', kind: 'subtotal', desc: '小計' },
    ],
  });
  const qa0 = mkQ('QA', 'QU-T-002', {
    items: [
      Object.assign({ lid: 'a1', desc: PROD, unit: '式', qty: 1, unitPrice: 1000000, cost: 0 }, T('ma', '2026-01-01', '2026-12-31')),
      { lid: 'a2', desc: '顧問人天', unit: '人天', qty: 3, unitPrice: 8000, cost: 0 },
      Object.assign({ lid: 'a3', desc: 'Oracle License', unit: '套', qty: 1, unitPrice: 300000, cost: 0 }, T('license', '2026-06-01', '2027-05-31')),
      { lid: 'a4', desc: '其他維護費', unit: '式', qty: 1, unitPrice: 1000, cost: 0 },
    ],
  });
  const qa = Object.assign({}, qa0, { approval: approvedOf(QA, qa0) });
  const quotes = [qd, qa];
  if (withTerms) {
    quotes.push(mkQ('QX', 'QU-T-003', { items: [
      { lid: 'x1', desc: XSS1 + 'XSS 品名', unit: '式', qty: 1, unitPrice: 1000, cost: 0, termKind: 'ma', termFrom: XSS1, termTo: XSS2 },
      { lid: 'x2', desc: XSS2 + 'License', unit: '套', qty: 1, unitPrice: 2000, cost: 0, termKind: 'license', termFrom: '2026-01-01', termTo: '2026-12-31', spec: XSS1, note: XSS2 },
    ] }));
    const big = [];
    for (let i = 0; i < 50; i++) {
      const it = { lid: 'b' + i, desc: (i % 2 ? 'License ' : '維護 MA ') + '長品名測試項目 ' + i + ' ' + '很長的描述文字'.repeat(i % 5), unit: '式', qty: 1, unitPrice: 1000 + i, cost: 0 };
      if (i % 3 !== 2) Object.assign(it, { termKind: i % 2 ? 'license' : 'ma', termFrom: '2026-01-01', termTo: '2026-12-31' });
      if (i % 4 === 0) it.spec = '規格說明 ' + i;
      if (i % 5 === 0) it.note = '備註 ' + i;
      big.push(it);
    }
    quotes.push(mkQ('QB', 'QU-T-004', { items: big }));
  }
  return {
    opportunities: OPPS, quotations: quotes,
    quoteApproval: { roster: { gm: ['_s_gm'], chairman: [], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: { [PROD]: { cls: 'software', costBySales: true } } },
    pricebook: { items: [{ id: 'pb1', bu: 'ERP', name: '顧問人天（測試）', price: 8000, cost: 5000, active: true, spec: '' }], updatedAt: T0, updatedBy: '_s_admin' },
  };
}

// ───────── 瀏覽器小工具 ─────────
async function openAs(b, C, who, w, h) {
  const H = C.H;
  if (w) await b.viewport(w, h || 900, w < 700);
  await b.nav(H.BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(H.PW)}})}).then(r=>r.status)`);
  C.t('瀏覽器登入 ' + who, st === 200, st);
  await b.ev(`sessionStorage.setItem('justLoggedIn','1'); window.__xss=0; 0`);
  await b.nav(H.BASE + '/');
  await b.waitFor(`typeof showSection==='function' && !!document.getElementById('navQuotations')`, 25000);
  await b.sleep(1200);
}
async function toQuotes(b) {
  await b.ev(`window.__xss=0; showSection('quotations'); 0`);
  await b.waitFor(`typeof quoteCfg==='function' && !!quoteCfg().me && document.querySelectorAll('#quoteTbody tr').length>0`, 15000);
  await b.sleep(400);
}
/** 截圖存到 C.shots（scratchpad/ma1_shots）；沒指定就略過 */
async function shotTo(C, b, name) { if (C.shots) { fs.mkdirSync(C.shots, { recursive: true }); await b.shot(path.join(C.shots, name)); } }
const rowSel = (n) => `#quoteItemsBody tr:nth-child(${n})`;
async function openEdit(b, id) {
  await b.ev(`openQuoteModal(${J(id)}); 0`);
  await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='flex' && document.querySelectorAll('#quoteItemsBody tr').length>0`, 15000);
  await b.sleep(500);
}
async function closeModal(b) { await b.ev(`(()=>{const o=document.getElementById('quoteModalOverlay'); if(o) o.style.display='none'; return 0})()`); }
/** 真實鍵盤輸入日期（zh-TW 的 年/月/日 順序，已在探測確認）：focus 後輸入 YYYYMMDD */
async function typeDate(b, sel, iso) { await b.focus(sel); await b.type(iso.replace(/-/g, '')); await b.sleep(80); }
/** 真實鍵盤改下拉：focus 後用上下鍵移到目標值 */
async function pickKind(b, sel, want) {
  await b.focus(sel);
  const order = ['', 'license', 'ma'];
  for (let i = 0; i < 3; i++) {
    const cur = await b.ev(`${Q(sel)}.value`);
    if (cur === want) return true;
    await b.key(order.indexOf(want) > order.indexOf(cur) ? 'ArrowDown' : 'ArrowUp');
  }
  return (await b.ev(`${Q(sel)}.value`)) === want;
}
/** 真實鍵盤輸入品名並 Tab 離開（觸發 change） */
async function typeDesc(b, sel, text) {
  await b.focus(sel);
  await b.ev(`${Q(sel)}.select(); 0`);
  await b.type(text);
  await b.key('Tab');
  await b.sleep(120);
}
const termOf = (b, n) => b.ev(`(()=>{const r=document.querySelector(${J(rowSel(n))}); if(!r) return null; const k=r.querySelector('.qterm-kind'); return k?{kind:k.value, from:r.querySelector('.qterm-from').value, to:r.querySelector('.qterm-to').value, datesHidden:r.querySelector('.qterm-dates').hidden, err:r.querySelector('.qterm-err').hidden?'':r.querySelector('.qterm-err').textContent, span:r.querySelector('.qterm-span').textContent, hint:!r.querySelector('.qterm-hint').hidden}:{none:true}})()`);
const descOf = (b) => b.ev(`[...document.querySelectorAll('#quoteItemsBody tr')].map(r=>r.dataset.kind?('['+r.dataset.kind+']'):r.querySelector('.qi-desc').value).join('|')`);
const itemsDump = (b) => b.ev(`JSON.stringify(readQuoteItems().map(i=>({d:i.desc,k:i.termKind,f:i.termFrom,t:i.termTo,kind:i.kind})))`);


/** 開瀏覽器；C.mutFiles（UI 變異用）＝{檔名: 改壞的原始碼}，用 CDP 攔截替換，不動磁碟上的檔案 */
function browse(C, maxMs, fn) {
  return withBrowser({ cdpPort: C.cdp, maxMs }, async (b) => {
    if (C.mutFiles) { Object.keys(C.mutFiles).forEach((k) => { b.replace['/' + k] = C.mutFiles[k]; }); await b.interceptOn(); }
    return fn(b);
  });
}

// ═════════════════════════ 開關關閉：與基準逐位元相同 ═════════════════════════
const BASE_FILES = ['quote.js', 'quote-steps.js', 'quote-preview.js', 'quote-approval.js', 'quote-pricebook.js'];
async function secOff(C) {
  const { H, t } = C;
  await H.fresh(mkUsers(H), seedData(C.ROOT, true), { env: {} });   // 資料裡有期間欄位，但開關關閉：畫面與輸出必須當它們不存在
  await H.loginAll(['_s_own']);
  const snap = (useBase) => withBrowser({ cdpPort: C.cdp, maxMs: 240000 }, async (b) => {
    if (useBase) { BASE_FILES.forEach((f) => { b.replace['/' + f] = fs.readFileSync(path.join(C.base, '_client', f), 'utf8'); }); await b.interceptOn(); }
    await openAs(b, C, '_s_own', 1440, 900);
    await toQuotes(b);
    const out = {};
    out.cfg = await b.ev(`JSON.stringify({ma: quoteCfg().maTracking, hasTerms: typeof quoteTermsOn==='function' ? quoteTermsOn() : 'n/a'})`);
    await openEdit(b, 'QD');
    out.tbody = await b.ev(`document.getElementById('quoteItemsBody').innerHTML`);
    out.qtermCount = await b.ev(`document.querySelectorAll('.qterm').length`);
    // 儲存請求本文：stub 掉 PUT，記下 payload
    await b.ev(`window.__puts=[]; const of=window.fetch.bind(window); window.fetch=(u,o)=>{ if(o&&o.method==='PUT'&&String(u).endsWith('/api/quotations/QD')){ window.__puts.push(o.body); return Promise.resolve(new Response(JSON.stringify({id:'QD',items:[],approval:null,perm:{}}),{status:200,headers:{'Content-Type':'application/json'}})); } return of(u,o); }; 0`);
    await b.click('#quoteModalSave');
    await b.waitFor(`window.__puts.length>0`, 8000);
    out.put = await b.ev(`window.__puts[0]||''`);
    await closeModal(b);
    out.preview = await b.ev(`(async()=>{ const r=await fetch('/api/quotations/QD'); const q=await r.json(); return buildQuotePreviewHtml(q,{issueDate:'2026-10-24',issuer:{name:'測試',mobile:'0900',phone:'02',ext:'1'},remarks:['條款一','條款二']}); })()`);
    out.previewA = await b.ev(`(async()=>{ const r=await fetch('/api/quotations/QA'); const q=await r.json(); return buildQuotePreviewHtml(q,{issueDate:'2026-10-24',issuer:{name:'測試',mobile:'0900',phone:'02',ext:'1'},remarks:['條款一','條款二']}); })()`);
    for (const id of ['QD', 'QA']) {
      await b.ev(`openQuoteApproval(${J(id)}); 0`);
      await b.waitFor(`(()=>{const e=document.querySelector('#quoteApprovalOverlay .qap-body'); return !!e && e.innerHTML.length>200 && !!e.querySelector('.qap-table')})()`, 10000);
      out['ap' + id] = await b.ev(`document.querySelector('#quoteApprovalOverlay .qap-body').innerHTML.replace(/[0-9]{4}[-/][0-9]{2}[-/][0-9]{2}[ T][0-9]{2}:[0-9]{2}(:[0-9]{2})?/g,'<ts>')`);
      await b.key("Escape"); await b.sleep(200);
    }
    out.errors = b.errors.slice();
    return out;
  });
  const cur = await snap(false);
  const old = await snap(true);
  t('UI-OFF1 開關關閉：config 沒有 maTracking、前端判斷 quoteTermsOn()=false、品項列沒有任何 .qterm 元素', /"hasTerms":false/.test(cur.cfg) && !/"ma":true/.test(cur.cfg) && cur.qtermCount === 0, cur.cfg + ' qterm=' + cur.qtermCount);
  t('UI-OFF2 開關關閉：品項列 tbody 的 HTML 與基準前端（舊版 quote.js）逐位元相同（含標題列、小計列、說明）', cur.tbody === old.tbody && cur.tbody.length > 500, cur.tbody.length + ' vs ' + old.tbody.length);
  t('UI-OFF3 開關關閉：儲存按鈕送出的 PUT 本文與基準逐位元相同，且不含 termKind／termFrom／termTo', cur.put === old.put && cur.put.length > 100 && !/term(Kind|From|To)/.test(cur.put), cur.put.slice(0, 160));
  t('UI-OFF4 開關關閉：客戶預覽 HTML（有期間資料的單與已核准單）與基準逐位元相同，沒有「服務期間」', cur.preview === old.preview && cur.previewA === old.previewA && cur.preview.length > 500 && !/服務期間|qpv-term/.test(cur.preview + cur.previewA));
  t('UI-OFF5 開關關閉：簽核面板（草稿單、已核准單）的 HTML 與基準逐位元相同，沒有 qterm-view', cur.apQD === old.apQD && cur.apQA === old.apQA && cur.apQD.length > 300 && !/qterm/.test(cur.apQD + cur.apQA), cur.apQD.length + '/' + cur.apQA.length);
  t('UI-OFF6 開關關閉：頁面沒有 JS 例外', cur.errors.length === 0 && old.errors.length === 0, J(cur.errors.concat(old.errors)).slice(0, 300));
}

// ═════════════════════════ 開關開啟：新增報價單（逐步填寫） ═════════════════════════
async function secNew(C) {
  const { H, t } = C;
  await H.fresh(mkUsers(H), seedData(C.ROOT, true), { env: { MA_TRACKING: '1' } });
  await H.loginAll(['_s_own', '_s_admin']);
  await browse(C, 400000, async (b) => {
    await openAs(b, C, '_s_own', 1440, 900);
    await toQuotes(b);
    t('UI-N0 開關開啟：config.maTracking=true，前端 quoteTermsOn()=true', await b.ev('quoteCfg().maTracking===true && quoteTermsOn()'));
    await b.ev(`openQuoteModal(null,{oppId:'o2'}); 0`);
    await b.waitFor(`document.querySelectorAll('#quoteItemsBody tr').length>0`, 15000);
    await b.sleep(800);
    const D1 = rowSel(1) + ' .qi-desc';
    // 品名預選：含 MA／License／維護／授權時預選（只是建議）
    await typeDesc(b, D1, 'SAP PCOE License MA 年度維護');
    let r1 = await termOf(b, 1);
    t('UI-N1 新增列：品名輸入完（Tab 離開）含 License MA → 預選 MA（MA 優先於 License），日期區出現，提示「已依品名預選」', r1.kind === 'ma' && !r1.datesHidden && r1.hint, J(r1));
    t('UI-N2 預選後缺起迄日：列內立刻紅字提示「請填寫服務期間的起日與迄日」，起迄日欄位標紅（aria-invalid）', /請填寫服務期間的起日與迄日/.test(r1.err) && (await b.ev(`document.querySelectorAll('${rowSel(1)} .qterm-bad[aria-invalid="true"]').length`)) === 2, J(r1));
    await b.shot('C:/Users/steven.lee/AppData/Local/Temp/qt_new2.png');
    // 逐步進度：報價項目這一步要列出服務期間
    const need = await b.ev(`document.getElementById('quoteModalOverlay').innerText`);
    t('UI-N3 逐步進度：「還需要」清單含「第 1 列的服務期間（License／MA 需填起訖日）」', /第 1 列的服務期間（License／MA 需填起訖日）/.test(need), need.slice(need.indexOf('還需要'), need.indexOf('還需要') + 120));
    // 一年／三年快捷：沒有起日 → 游標回到起日、不填迄日
    await b.click(rowSel(1) + ' .qterm-quick');
    r1 = await termOf(b, 1);
    t('UI-N4 沒有起日按「一年」：不填迄日，游標回到起日欄', r1.to === '' && (await b.ev(`document.activeElement===${Q(rowSel(1) + ' .qterm-from')}`)), J(r1));
    await typeDate(b, rowSel(1) + ' .qterm-from', '2026-03-01');
    r1 = await termOf(b, 1);
    t('UI-N5 真實鍵盤輸入起日 → 起日 2026-03-01；仍缺迄日只標迄日欄為紅', r1.from === '2026-03-01' && /請填寫/.test(r1.err), J(r1));
    await b.click(rowSel(1) + ' .qterm-quick');
    r1 = await termOf(b, 1);
    t('UI-N6 按「一年」→ 迄日 2027-02-28、顯示「共 12 個月」、紅字消失', r1.to === '2027-02-28' && r1.span === '共 12 個月' && r1.err === '', J(r1));
    await b.click(rowSel(1) + ' .qterm-quick:nth-of-type(2)');
    r1 = await termOf(b, 1);
    t('UI-N7 按「三年」→ 迄日 2029-02-28、顯示「共 36 個月」', r1.to === '2029-02-28' && r1.span === '共 36 個月', J(r1));
    await typeDate(b, rowSel(1) + ' .qterm-to', '2026-01-01');
    r1 = await termOf(b, 1);
    t('UI-N8 迄日早於起日：紅字「迄日不得早於起日」，不顯示月數', /迄日不得早於起日/.test(r1.err) && r1.span === '', J(r1));
    await b.click(rowSel(1) + ' .qterm-quick');
    // 第二列：一般品名不預選；License 預選；手動選過就不再自動動
    await b.click('#addQuoteItemBtn');
    await b.sleep(300);
    await typeDesc(b, rowSel(2) + ' .qi-desc', '教育訓練');
    let r2 = await termOf(b, 2);
    t('UI-N9 一般品名（教育訓練）不預選，日期區隱藏', r2.kind === '' && r2.datesHidden && !r2.hint, J(r2));
    await typeDesc(b, rowSel(2) + ' .qi-desc', 'Oracle 授權');
    r2 = await termOf(b, 2);
    t('UI-N10 品名改成「Oracle 授權」→ 預選 License（尚未手動選過）', r2.kind === 'license' && r2.hint, J(r2));
    t('UI-N11 真實鍵盤改下拉為 MA（手動）→ 提示消失；再把品名改成 License → 維持 MA，不被自動覆蓋', (await pickKind(b, rowSel(2) + ' .qterm-kind', 'ma')) && !(await termOf(b, 2)).hint && (await (async () => { await typeDesc(b, rowSel(2) + ' .qi-desc', 'SAP License'); return (await termOf(b, 2)).kind === 'ma'; })()));
    await typeDate(b, rowSel(2) + ' .qterm-from', '2026-05-01'); await typeDate(b, rowSel(2) + ' .qterm-to', '2027-04-30');
    await b.ev(`${Q(rowSel(2) + ' .qi-price')}.value='1200'; ${Q(rowSel(2) + ' .qi-price')}.dispatchEvent(new Event('input',{bubbles:true})); 0`);
    // 第三列：一般
    await b.click('#addQuoteItemBtn'); await b.sleep(300);
    await typeDesc(b, rowSel(3) + ' .qi-desc', '現場安裝');
    // 排序：▲▼ 點擊
    const before = await itemsDump(b);
    await b.click(rowSel(3) + ' .qi-up'); await b.sleep(250);
    const afterUp = await itemsDump(b);
    const bA = JSON.parse(before), aA = JSON.parse(afterUp);
    t('UI-N12 ▲ 上移一般列：列順序改變，License／MA 兩列的類型與起迄日原樣跟著各自的列', aA[1].d === '現場安裝' && aA[2].d === 'SAP License' && aA[2].k === 'ma' && aA[2].f === '2026-05-01' && aA[2].t === '2027-04-30' && aA[0].k === 'ma' && aA[0].f === '2026-03-01' && aA[0].t === '2027-02-28', afterUp);
    await b.click(rowSel(2) + ' .qi-down'); await b.sleep(250);
    // 拖曳：真實滑鼠把第 1 列把手拖到第 3 列
    const d0 = JSON.parse(await itemsDump(b));
    await b.drag(rowSel(1) + ' .qi-drag', rowSel(3) + ' .qi-drag', { dy: 12 });
    await b.sleep(400);
    const d1 = JSON.parse(await itemsDump(b));
    const moved = d1.map((x) => x.d).join('|') !== d0.map((x) => x.d).join('|');
    const keep = d0.every((o) => { const n = d1.find((x) => x.d === o.d); return n && n.k === o.k && n.f === o.f && n.t === o.t; });
    t('UI-N13 真實滑鼠拖曳排序：順序改變，每一列的類型與起迄日都跟著它的列', moved && keep && d1.length === 3, J(d0) + ' → ' + J(d1));
    // 刪除列
    const nTerm = d1.findIndex((x) => x.k === 'license' || x.k === 'ma') + 1;
    const delIdx = d1.findIndex((x) => !x.k) + 1;
    await b.click(rowSel(delIdx) + ' .qi-remove'); await b.sleep(250);
    const d2 = JSON.parse(await itemsDump(b));
    t('UI-N14 刪除一般列：剩下兩個 License／MA 列的期間不變', d2.length === 2 && d2.every((x) => x.k && x.f && x.t) && J(d2) === J(d1.filter((x) => x.k)), J(d2));
    // 牌價簿加入：以函式路徑（planAppend＋renderQuoteItems＝牌價簿 onApply 的路徑）驗證既有列保留
    const pbRes = await b.ev(`(()=>{ const cur=readQuoteItems(); const add=QPB.buildItems([{id:'pb1',bu:'ERP',name:'顧問人天（測試）',price:8000,cost:5000,spec:''}],[{id:'pb1',qty:2}]); const pl=QPB.planAppend(cur,add,50); renderQuoteItems(pl.items); updateQuoteTotals(); return JSON.stringify(readQuoteItems().map(i=>[i.desc,i.termKind||'',i.termFrom||'',i.termTo||''])); })()`);
    t('UI-N15 牌價簿加入列：既有 License／MA 列的期間保留，新加的人天列是一般（沒有期間）', (() => { const a = JSON.parse(pbRes); return a.length === 3 && a[0][1] && a[1][1] && a[2][0] === '顧問人天（測試）' && a[2][1] === '' && a[0][3] && a[1][3]; })(), pbRes);
    const seed = await b.ev(`JSON.stringify(QCL.seedFromItems(_qClItems(), { classCodes: _qClClassCodes() }))`);
    t('UI-N20 成本編輯器種子（QCL.seedFromItems）不受影響：輸出不含任何期間欄位，行數＝有價品項數', typeof seed === 'string' && seed.indexOf('term') < 0 && seed.indexOf('服務期間') < 0 && JSON.parse(seed).length >= 1, String(seed).slice(0, 200));
    const blankKept = await b.ev(`(()=>{ const add=QPB.buildItems([{id:'pb1',bu:'ERP',name:'x',price:1,cost:1,spec:''}],[{id:'pb1',qty:1}]); const pl=QPB.planAppend([{desc:'',unit:'式',qty:1,unitPrice:0,termKind:'license',termFrom:'',termTo:''}],add,50); return JSON.stringify([pl.items.length, pl.replacedBlank]); })()`);
    t('UI-N21 牌價簿：只有一列「空白但已選 License／MA 類型」的預設列時不會被當成未動過的空白列取代', blankKept === '[2,false]', blankKept);
    await b.shot('C:/Users/steven.lee/AppData/Local/Temp/qt_new3.png');
    // 填單價（真實鍵盤）並走完逐步填寫後儲存
    const rowsNow = await b.ev(`document.querySelectorAll('#quoteItemsBody tr').length`);
    for (let i = 1; i <= rowsNow; i++) {
      await b.focus(rowSel(i) + ' .qi-price'); await b.ev(`${Q(rowSel(i) + ' .qi-price')}.select(); 0`); await b.type(String(1000 * i)); await b.sleep(60);
    }
    await b.key('Tab');
    const itemsBeforeSave = JSON.parse(await itemsDump(b));
    for (let i = 0; i < 8; i++) {
      const clicked = await b.ev(`(()=>{const bt=[...document.querySelectorAll('#quoteModalOverlay button')].find(x=>x.textContent.trim().indexOf('下一步')===0 && x.offsetParent!==null && !x.disabled); if(!bt) return false; bt.click(); return true})()`);
      if (!clicked) break;
      await b.sleep(350);
    }
    const saveEnabled = await b.waitFor(`!document.getElementById('quoteModalSave').disabled`, 6000);
    if (!saveEnabled) { console.log('DBG', await b.ev(`document.getElementById('quoteModalOverlay').innerText.slice(0,1500)`)); await b.shot('C:/Users/steven.lee/AppData/Local/Temp/qt_new4.png'); }
    t('UI-N16 走完逐步填寫（服務期間都填妥）→ 「儲存」可按', saveEnabled);
    await b.click('#quoteModalSave');
    await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='none'`, 15000);
    await b.sleep(800);
    const list = (await H.http('_s_own', 'GET', '/api/quotations')).json;
    const mine = list.find((q) => q.company === '第二公司');
    const sv = mine ? (await H.http('_s_own', 'GET', '/api/quotations/' + mine.id)).json : null;
    const sItems = sv ? sv.items.map((x) => [x.desc, x.termKind || '', x.termFrom || '', x.termTo || '']) : [];
    t('UI-N17 按儲存：伺服器存下兩個 License／MA 品項的類型與起迄日（與畫面一致），一般品項沒有期間欄位', !!sv && sv.items.length === 3 && itemsBeforeSave.every((x, i) => sItems[i][0] === x.d && sItems[i][1] === (x.k || '') && sItems[i][2] === (x.f || '') && sItems[i][3] === (x.t || '')) && sv.items.filter((x) => !x.termKind).every((x) => !('termFrom' in x) && !('termTo' in x)), J(sItems));
    C.newQid = mine && mine.id;
    // 重新載入頁面後再編輯：欄位原樣顯示
    await openAs(b, C, '_s_own', 1440, 900);
    await toQuotes(b);
    await openEdit(b, C.newQid);
    const re = JSON.parse(await itemsDump(b));
    t('UI-N18 重新載入頁面、開啟編輯：類型與起迄日原樣讀回', J(re.map((x) => [x.d, x.k || '', x.f || '', x.t || ''])) === J(sItems), J(re));
    // 編輯後只改迄日再存：不洗掉其他列
    await typeDate(b, rowSel(re.findIndex((x) => x.k === 'license') + 1 || 1) + ' .qterm-to', '2030-01-01');
    await b.click('#quoteModalSave');
    await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='none'`, 15000);
    await b.sleep(600);
    const sv2 = (await H.http('_s_own', 'GET', '/api/quotations/' + C.newQid)).json;
    t('UI-N19 編輯儲存只改一列迄日：該列更新、其他列的期間不變（沒有被洗掉）', sv2.items.filter((x) => x.termKind).length === 2 && sv2.items.some((x) => x.termTo === '2030-01-01') && sv2.items.filter((x) => x.termTo !== '2030-01-01' && x.termKind).every((x) => sv.items.find((o) => o.lid === x.lid).termTo === x.termTo), J(sv2.items.map((x) => [x.termKind, x.termTo])));
  });
}

// ═════════════════════════ 開關開啟：編輯既有草稿、XSS、核准後修改 ═════════════════════════
async function secEdit(C) {
  const { H, t } = C;
  await H.fresh(mkUsers(H), seedData(C.ROOT, true), { env: { MA_TRACKING: '1' } });
  await H.loginAll(['_s_own', '_s_other', '_s_admin', '_s_m1', '_s_sec']);
  await browse(C, 500000, async (b) => {
    await openAs(b, C, '_s_own', 1440, 900);
    await toQuotes(b);
    // ── 編輯 QD ──
    await openEdit(b, 'QD');
    const dump0 = JSON.parse(await itemsDump(b));
    t('UI-E1 編輯草稿：已存的類型與起迄日讀回；標題列、小計列沒有期間欄位', dump0.length === 5 && dump0[0].k === 'license' && dump0[0].f === '2026-01-01' && dump0[0].t === '2026-12-31' && dump0[1].k === 'ma' && dump0[1].t === '2027-02-28' && !dump0[2].k && !dump0[3].k && !dump0[4].k && (await b.ev(`document.querySelectorAll('#quoteItemsBody tr[data-kind] .qterm').length`)) === 0, J(dump0));
    t('UI-E2 已存列顯示「共 12 個月」（2026-01-01～2026-12-31）；2026-03-01～2027-02-28 也是 12 個月', (await termOf(b, 1)).span === '共 12 個月' && (await termOf(b, 2)).span === '共 12 個月');
    // 把 License 列改成一般 → 儲存請求帶 termKind '' 且不帶日期
    await b.ev(`window.__puts=[]; const of=window.fetch.bind(window); window.fetch=(u,o)=>{ if(o&&o.method==='PUT'){ window.__puts.push(o.body); } return of(u,o); }; 0`);
    t('UI-E3 真實鍵盤把第 1 列類型改回「一般」', await pickKind(b, rowSel(1) + ' .qterm-kind', ''));
    t('UI-E4 改回一般後日期區隱藏、沒有紅字', (await termOf(b, 1)).datesHidden && (await termOf(b, 1)).err === '');
    // 第 3 列（教育訓練）改成 License 但不填日期 → 按儲存被擋：提示第 N 列、游標到起日欄
    await pickKind(b, rowSel(3) + ' .qterm-kind', 'license');
    await b.click('#quoteModalSave');
    await b.sleep(500);
    const toast = await b.ev(`(()=>{const e=[...document.querySelectorAll('.toast,#toast,[class*=toast]')].map(x=>x.textContent).join('|'); return e})()`);
    const focusedIsFrom = await b.ev(`document.activeElement===${Q(rowSel(3) + ' .qterm-from')}`);
    t('UI-E5 License 列沒填日期按儲存：前端擋下（沒有送出 PUT）、游標在第 3 列的起日欄、toast 指出「第 3 列」', focusedIsFrom && (await b.ev('window.__puts.length')) === 0 && /第 3 列/.test(toast), toast.slice(0, 120));
    // 填妥後儲存：第 1 列清除、第 3 列新增
    await typeDate(b, rowSel(3) + ' .qterm-from', '2026-07-01'); await b.click(rowSel(3) + ' .qterm-quick');
    await b.click('#quoteModalSave');
    await b.waitFor(`document.getElementById('quoteModalOverlay').style.display==='none'`, 15000);
    await b.sleep(600);
    const put = JSON.parse(await b.ev('window.__puts[0]'));
    t('UI-E6 儲存請求：改回一般的列送 termKind:""（不帶日期）；新設 License 的列帶起迄日；沒動的 MA 列原樣帶', put.items[0].termKind === '' && !('termFrom' in put.items[0]) && put.items[1].termKind === 'ma' && put.items[1].termFrom === '2026-03-01' && put.items[2].termKind === 'license' && put.items[2].termTo === '2027-06-30' && !('termKind' in put.items[3]), J(put.items.map((x) => [x.termKind, x.termFrom, x.termTo])));
    const sv = (await H.http('_s_own', 'GET', '/api/quotations/QD')).json;
    t('UI-E7 伺服器：第 1 列欄位整個移除、第 3 列變 License 2026-07-01～2027-06-30、標題／小計列不含期間', !('termKind' in sv.items[0]) && sv.items[2].termKind === 'license' && sv.items[2].termFrom === '2026-07-01' && sv.items[2].termTo === '2027-06-30' && !('termKind' in sv.items[3]) && !('termKind' in sv.items[4]), J(sv.items.map((x) => [x.termKind, x.termFrom, x.termTo])));
    // ── XSS：髒資料（期間欄位存了 HTML）與品名含標籤 ──
    await openEdit(b, 'QX');
    const x1 = await termOf(b, 1);
    const xs = await b.ev('window.__xss');
    t('UI-E8 XSS：品名含 <img onerror>／<script> 與髒資料的期間在編輯表單只當文字：沒有執行（__xss=0）、品名輸入框保留原字串、日期欄空白並紅字提示', xs === 0 && (await b.ev(`${Q(rowSel(1) + ' .qi-desc')}.value.indexOf('<img src=x onerror')`)) === 0 && x1.from === '' && x1.to === '' && /請填寫|不是有效日期/.test(x1.err), J(x1) + ' xss=' + xs);
    t('UI-E9 XSS：表單裡沒有被解析出來的 img／script 元素', (await b.ev(`document.querySelectorAll('#quoteItemsBody img, #quoteItemsBody script').length`)) === 0);
    await closeModal(b);
    await b.ev(`previewQuote('QX'); 0`);
    await b.waitFor(`!!document.querySelector('#quotePreviewOverlay .qpv-paper')`, 10000);
    const pv = await b.ev(`document.querySelector('#quotePreviewOverlay .qpv-items').innerHTML`);
    t('UI-E10 XSS：客戶預覽——髒期間不顯示任何「服務期間」行（不合法就當沒有）；品名標籤被跳脫；沒有執行', (await b.ev('window.__xss')) === 0 && !/<img|<script/i.test(pv.replace(/&lt;[^&]*?&gt;/g, '')) && (pv.match(/qpv-term/g) || []).length === 1 && pv.indexOf('服務期間：2026/01/01 – 2026/12/31') >= 0, (pv.match(/qpv-term/g) || []).length + ' ' + pv.slice(0, 200));
    await b.ev(`closeQuotePreview(); openQuoteApproval('QX'); 0`);
    await b.waitFor(`!!document.querySelector('#quoteApprovalOverlay .qap-table')`, 10000);
    const apx = await b.ev(`document.querySelector('#quoteApprovalOverlay .qap-body').innerHTML`);
    t('UI-E11 XSS：簽核面板唯讀顯示——只有合法期間那一列有「服務期間」行，標籤全被跳脫、沒有執行', (await b.ev('window.__xss')) === 0 && (apx.match(/class="qterm-view"/g) || []).length === 1 && !/<img src=x|<script>window/.test(apx));
    await b.key('Escape'); await b.sleep(200);
    t('UI-E12 XSS 區段沒有 JS 例外', b.errors.length === 0, J(b.errors).slice(0, 300));
    // ── 權限：列表上的「📅 服務期間」鈕 ──
  });
}

// ═════════════════════════ 核准後修改服務期間 ═════════════════════════
const hasTermsBtn = (b, id) => b.ev(`!!document.querySelector('#quoteTbody [data-qact="terms"][data-id=${J(id)}]')`);
async function secApproved(C) {
  const { H, t } = C;
  await H.fresh(mkUsers(H), seedData(C.ROOT, true), { env: { MA_TRACKING: '1' } });
  await H.loginAll(['_s_own', '_s_other', '_s_admin', '_s_m1', '_s_sec']);
  const hashBefore = (await H.http('_s_own', 'GET', '/api/quotations/QA')).json.approval;
  await browse(C, 500000, async (b) => {
    // ── 別的帳號看不到按鈕 ──
    for (const who of ['_s_m1', '_s_other']) {
      await openAs(b, C, who, 1440, 900);
      await toQuotes(b);
      t('UI-A1 ' + who + '（簽核者／無關業務）的列表上沒有「📅 服務期間」鈕（擁有者與管理員才有）', !(await hasTermsBtn(b, 'QA')));
    }
    await openAs(b, C, '_s_admin', 1440, 900);
    await toQuotes(b);
    t('UI-A2 管理員的列表上有「📅 服務期間」鈕（簽核中／已核准的單，且有 License／MA 品項）', await hasTermsBtn(b, 'QA'));
    t('UI-A3 草稿單（QD）沒有「服務期間」鈕（草稿直接編輯）', !(await hasTermsBtn(b, 'QD')));
    await openAs(b, C, '_s_own', 1440, 900);
    await toQuotes(b);
    t('UI-A4 擁有者的列表上有「📅 服務期間」鈕', await hasTermsBtn(b, 'QA'));
    await b.click('#quoteTbody [data-qact="terms"][data-id="QA"]');
    t('UI-A5 按鈕開啟小視窗（role=dialog、標題含單號）', (await b.waitFor(`!!document.querySelector('#qTermsOverlay [role=dialog]')`, 8000)) && (await b.ev(`document.getElementById('qTermsTitle').textContent`)).indexOf('QU-T-002') >= 0);
    const rowsInfo = await b.ev(`JSON.stringify([...document.querySelectorAll('#qTermsOverlay .qn-row')].map(r=>({name:r.querySelector('.qn-name').textContent, kind:r.querySelector('.qterm-kind').value, opts:[...r.querySelectorAll('.qterm-kind option')].map(o=>o.value).join(','), from:r.querySelector('.qterm-from').value, to:r.querySelector('.qterm-to').value})))`);
    const ri = JSON.parse(rowsInfo);
    t('UI-A6 小視窗只列 License／MA 品項與品名像 License／MA 的品項（顧問人天不列）；已是 License／MA 的列沒有「一般」選項，其他列有', ri.length === 3 && ri[0].kind === 'ma' && ri[0].opts === 'license,ma' && ri[1].kind === 'license' && ri[2].kind === '' && ri[2].opts === ',license,ma' && !ri.some((x) => /顧問人天/.test(x.name)), rowsInfo);
    const dlgText = await b.ev(`document.getElementById('qTermsOverlay').innerText`);
    t('UI-A7 對話框內說明「客戶報價單上印出的期間會立即改變」「不需重新簽核」', /客戶報價單上印出的期間會立即改變/.test(dlgText) && /不需重新簽核/.test(dlgText));
    await shotTo(C, b, 'qterm_dialog_glass-dark.png');
    // 沒有修改 → 提示、不送出
    await b.ev(`window.__req=[]; const of=window.fetch.bind(window); window.fetch=(u,o)=>{ if(o&&o.method==='PATCH') window.__req.push(String(u)+' '+o.body); return of(u,o); }; 0`);
    await b.click('#qTermsOverlay .btn-primary');
    t('UI-A8 沒改任何東西按儲存：顯示「沒有修改任何內容」，不送 PATCH', /沒有修改任何內容/.test(await b.ev(`document.querySelector('#qTermsOverlay .qn-err').textContent`)) && (await b.ev('window.__req.length')) === 0);
    // 顛倒日期 → 前端擋
    await typeDate(b, '#qTermsOverlay .qn-row:nth-child(1) .qterm-to', '2025-01-01');
    await b.click('#qTermsOverlay .btn-primary');
    t('UI-A9 迄日早於起日按儲存：顯示「第 1 列：…迄日不得早於起日」，不送 PATCH', /第 1 列：.*迄日不得早於起日/.test(await b.ev(`document.querySelector('#qTermsOverlay .qn-err').textContent`)) && (await b.ev('window.__req.length')) === 0);
    // 伺服器錯誤碼對應中文（stub 回應）
    await typeDate(b, '#qTermsOverlay .qn-row:nth-child(1) .qterm-to', '2027-03-31');
    const codes = {};
    for (const code of ['TERM_KIND_LOCKED', 'NO_PERMISSION', 'BAD_LID', 'TERM_NOT_APPLICABLE', 'BAD_FIELDS']) {
      await b.ev(`window.__fake=${J(code)}; if(!window.__f0){ window.__f0=window.fetch; window.fetch=(u,o)=>{ if(o&&o.method==='PATCH'&&window.__fake) return Promise.resolve(new Response(JSON.stringify({error:'server msg',code:window.__fake}),{status:400,headers:{'Content-Type':'application/json'}})); return window.__f0(u,o); }; } 0`);
      await b.click('#qTermsOverlay .btn-primary'); await b.sleep(250);
      codes[code] = await b.ev(`document.querySelector('#qTermsOverlay .qn-err').textContent`);
    }
    await b.ev('window.__fake=null; 0');
    t('UI-A10 伺服器錯誤碼對應中文：TERM_KIND_LOCKED／NO_PERMISSION／BAD_LID／TERM_NOT_APPLICABLE／BAD_FIELDS 各有專屬訊息（不顯示英文代碼或伺服器原文）', Object.values(codes).every((m) => m && !/server msg|[A-Z]{4,}_/.test(m)) && new Set(Object.values(codes)).size === 5, J(codes));
    // 真的送出：改第 1 列（MA）迄日為 2027-03-31
    await b.click('#qTermsOverlay .btn-primary');
    t('UI-A11 儲存成功：小視窗關閉並自動開啟簽核進度（顯示時間軸）', await b.waitFor(`!document.getElementById('qTermsOverlay') && !!document.querySelector('#quoteApprovalOverlay .qap-body .qap-table')`, 12000));
    await b.sleep(500);
    const apHtml = await b.ev(`document.querySelector('#quoteApprovalOverlay .qap-body').innerText`);
    t('UI-A12 簽核面板品項區在品名下唯讀顯示「類型徽章＋服務期間：…」（新迄日 2027/03/31）', apHtml.indexOf('服務期間：2026/01/01 – 2027/03/31') >= 0 && apHtml.indexOf('服務期間：2026/06/01 – 2027/05/31') >= 0, apHtml.slice(Math.max(0, apHtml.indexOf('服務期間') - 40), apHtml.indexOf('服務期間') + 80));
    t('UI-A13 時間軸有「修改品項服務期間」那一筆', /修改品項服務期間/.test(apHtml), apHtml.slice(-300));
    await shotTo(C, b, 'qterm_approval_panel_glass-dark.png');
    await b.key('Escape'); await b.sleep(200);
    // 伺服器端驗證
    const after = (await H.http('_s_own', 'GET', '/api/quotations/QA')).json;
    t('UI-A14 伺服器：第 1 列迄日更新為 2027-03-31，核准仍有效（approval.valid=true、state=approved、雜湊沒變），金額沒動', after.items[0].termTo === '2027-03-31' && after.approval.state === 'approved' && after.approval.valid === true && after.approval.hash === hashBefore.hash && after.items[0].unitPrice === 1000000, J(after.approval).slice(0, 200));
    const logs = await H.getLogs('_s_admin');
    const lg = logs.filter((l) => l.action === 'UPDATE_QUOTE_TERM');
    t('UI-A15 稽核 UPDATE_QUOTE_TERM 一筆，內容含「第 1 列」與舊→新期間', lg.length === 1 && /第 1 列/.test(lg[0].detail) && lg[0].detail.indexOf('2026/01/01 – 2026/12/31→2026/01/01 – 2027/03/31') >= 0, lg[0] && lg[0].detail);
    const ntf = async (who) => { const r = await H.http(who, 'GET', '/api/notifications'); return (Array.isArray(r.json) ? r.json : (r.json && r.json.notifications) || []).filter((n) => n.type === 'quote_term_changed'); };
    const nm1 = await ntf('_s_m1'), nown = await ntf('_s_own'), nother = await ntf('_s_other');
    t('UI-A16 通知：已簽核的一級主管收到 quote_term_changed（內文只有單號與公司名、沒有金額與期間文字）；無關業務沒有', nm1.length === 1 && nm1[0].refId === 'QA' && /QU-T-002/.test(nm1[0].body || nm1[0].title) && !/1000000|2027/.test(J(nm1[0])) && nother.length === 0, J(nm1).slice(0, 300));
    t('UI-A17 通知：修改者本人（擁有者）不會收到自己觸發的通知', nown.length === 0, J(nown).slice(0, 200));
    // 把「其他維護費」（一般品項）補成 MA：小視窗第 3 列
    await toQuotes(b);
    await b.click('#quoteTbody [data-qact="terms"][data-id="QA"]');
    await b.waitFor(`!!document.querySelector('#qTermsOverlay .qn-row')`, 8000);
    await pickKind(b, '#qTermsOverlay .qn-row:nth-child(3) .qterm-kind', 'ma');
    await typeDate(b, '#qTermsOverlay .qn-row:nth-child(3) .qterm-from', '2026-09-01');
    await b.click('#qTermsOverlay .qn-row:nth-child(3) .qterm-quick');
    await b.click('#qTermsOverlay .btn-primary');
    await b.waitFor(`!document.getElementById('qTermsOverlay')`, 12000);
    await b.sleep(500);
    await b.key('Escape'); await b.sleep(200);
    const after2 = (await H.http('_s_own', 'GET', '/api/quotations/QA')).json;
    t('UI-A18 一般品項在核准後補成 MA（類型＋期間）：存入，核准仍有效', after2.items[3].termKind === 'ma' && after2.items[3].termFrom === '2026-09-01' && after2.items[3].termTo === '2027-08-31' && after2.approval.valid === true, J(after2.items[3]));
    t('UI-A19 核准後修改區段沒有 JS 例外', b.errors.length === 0, J(b.errors).slice(0, 300));
  });
}

// ═════════════════════════ 客戶文件：預覽、PDF、Excel ═════════════════════════
async function secDocs(C) {
  const { H, t } = C;
  await H.fresh(mkUsers(H), seedData(C.ROOT, true), { env: { MA_TRACKING: '1' } });
  await H.loginAll(['_s_own', '_s_admin']);
  // ── Excel（真實下載端點；伺服器用範本填，業務擁有者下載）──
  const JSZip = require(path.join(C.ROOT, 'node_modules', 'jszip'));
  const XL = require(path.join(C.ROOT, 'lib/quoteExcel.js'));
  const xlOf = async (id) => {
    const r = await H.http('_s_own', 'GET', '/api/quotations/' + id + '/export');
    if (r.status !== 200) return { status: r.status, text: r.text.slice(0, 200) };
    const zip = await JSZip.loadAsync(r.buf);
    const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
    const ss = zip.file('xl/sharedStrings.xml') ? await zip.file('xl/sharedStrings.xml').async('string') : '';
    return { status: 200, sheet, ss, all: sheet + ss };
  };
  const rowHt = (xml, r) => { const m = new RegExp('<row r="' + r + '"[^>]*>').exec(xml); const h = m && /ht="([0-9.]+)"/.exec(m[0]); return h ? parseFloat(h[1]) : null; };
  const x1 = await xlOf('QD');
  t('UI-X1 Excel（QD）下載成功，品名下多一行灰字「服務期間：2026/01/01 – 2026/12/31」「服務期間：2026/03/01 – 2027/02/28」', x1.status === 200 && x1.all.indexOf('服務期間：2026/01/01 – 2026/12/31') >= 0 && x1.all.indexOf('服務期間：2026/03/01 – 2027/02/28') >= 0, x1.status + (x1.text || ''));
  const L = XL.LAYOUT, I = XL._internal;
  const hd1 = rowHt(x1.sheet, L.itemFirst), hd3 = rowHt(x1.sheet, L.itemFirst + 2);
  t('UI-X2 Excel 列高：有期間的列依「品名＋期間行」重算（>= 單純品名的高度），沒期間的一般列維持原本公式', hd1 === I.itemRowHeightWithNotes('SAP License 授權', '', '', '服務期間：2026/01/01 – 2026/12/31') && hd1 > I.itemRowHeight('SAP License 授權') && hd3 === I.itemRowHeight('教育訓練'), 'ht1=' + hd1 + ' ht3=' + hd3);
  t('UI-X3 Excel：期間行在品名之後、說明之前（第 2 列：品名→服務期間→說明「含遠端支援」）', (() => { const a = x1.all.indexOf('年度 MA 維護'), b = x1.all.indexOf('服務期間：2026/03/01'), c = x1.all.indexOf('含遠端支援'); return a >= 0 && b > a && c > b; })());
  const xb = await xlOf('QB');
  const termCount = (xb.all.match(/服務期間：/g) || []).length;
  let maxHt = 0, rowsOk = true;
  for (let i = 0; i < 50; i++) { const h = rowHt(xb.sheet, L.itemFirst + i); if (h === null || h > 409) rowsOk = false; maxHt = Math.max(maxHt, h || 0); }
  t('UI-X4 Excel 50 列上限實測（QB：50 個品項、34 列有期間、部分有說明／備註）：產出成功、34 行服務期間都在、50 列都有列高且 <= 409', xb.status === 200 && termCount === 34 && rowsOk, 'status=' + xb.status + ' terms=' + termCount + ' maxHt=' + maxHt);
  // ── 預覽與 PDF（瀏覽器）──
  await browse(C, 500000, async (b) => {
    await openAs(b, C, '_s_own', 1440, 900);
    await toQuotes(b);
    await b.ev(`previewQuote('QD'); 0`);
    await b.waitFor(`!!document.querySelector('#quotePreviewOverlay .qpv-paper')`, 10000);
    const pv = JSON.parse(await b.ev(`JSON.stringify([...document.querySelectorAll('#quotePreviewOverlay .qpv-items tbody tr')].map(tr=>{const d=tr.querySelector('td.desc'); return d?[...d.childNodes].map(n=>n.nodeType===3?n.textContent:(n.className||n.tagName)+':'+n.textContent).join('|'):'';}).filter(Boolean))`));
    t('UI-V1 客戶預覽：License／MA 品項品名下方一行「服務期間：…」（灰字 .qpv-term），一般品項沒有', pv[0].indexOf('qpv-term:服務期間：2026/01/01 – 2026/12/31') >= 0 && pv[1].indexOf('qpv-term:服務期間：2026/03/01 – 2027/02/28') >= 0 && pv[2].indexOf('qpv-term') < 0 && pv.length >= 3, J(pv.slice(0, 3)));
    t('UI-V2 客戶預覽：期間行在品名之後、說明之前（第 2 列 qpv-term 先於 qpv-spec）', pv[1].indexOf('qpv-term') > 0 && pv[1].indexOf('qpv-term') < pv[1].indexOf('qpv-spec'), pv[1]);
    const cs = await b.ev(`(()=>{const e=document.querySelector('#quotePreviewOverlay .qpv-term'); const c=getComputedStyle(e); return {fs:parseFloat(c.fontSize), color:c.color}})()`);
    await shotTo(C, b, 'qterm_preview_glass-dark.png');
    await b.ev(`closeQuotePreview(); 0`);
    // PDF：攔截 jsPDF.save（不真的下載），記頁數與大小；大單 QB（50 列）驗證不被截斷、不報錯
    await b.ev(`window.__toasts=[]; const st=window.showToast; window.showToast=function(m){ window.__toasts.push(String(m)); return st.apply(this,arguments); }; 0`);
    await b.ev(`_qpdfEnsureLibs().then(()=>{ window.__pdf=[]; const P=window.jspdf.jsPDF.API; P.save=function(n){ window.__pdf.push({name:n,pages:this.getNumberOfPages(),bytes:this.output('arraybuffer').byteLength}); }; return 1; })`);
    for (const [id, no] of [['QD', 'QU-T-001'], ['QB', 'QU-T-004']]) {
      await b.ev(`window.__pdf.length=0; window.__pe=null; Promise.resolve(exportQuotePdf(${J(id)}, ${J(no)})).catch(e=>{window.__pe=String(e&&e.stack||e)}); 0`);
      await b.waitFor(`[...document.querySelectorAll('button')].some(x=>x.textContent.indexOf('仍要')>=0 && x.offsetParent!==null)`, 15000);   // 未核准的單下載前會先警示
      await b.ev(`[...document.querySelectorAll('button')].find(x=>x.textContent.indexOf('仍要')>=0 && x.offsetParent!==null).click(); 0`);
      const got = await b.waitFor(`window.__pdf.length>0 || window.__toasts.some(m=>/失敗|⚠/.test(m)) || !!window.__pe`, 60000);
      if (!got || await b.ev('window.__pe')) console.log('DBG', await b.ev('JSON.stringify([window.__toasts, window.__pe, typeof exportQuotePdf])'));
      const pdf = await b.ev(`JSON.stringify(window.__pdf[0]||null)`);
      const bad = await b.ev(`window.__toasts.filter(m=>/失敗|⚠/.test(m)).join('|')`);
      t('UI-P1 PDF 產生（' + id + '）：沒有錯誤訊息、攔截到 PDF（頁數與大小合理）', got && !bad && !!JSON.parse(pdf) && JSON.parse(pdf).pages >= 1 && JSON.parse(pdf).bytes > 20000, pdf + ' ' + bad);
      if (id === 'QB') t('UI-P2 50 列大單的 PDF 分多頁（>= 3 頁）', JSON.parse(pdf).pages >= 3, pdf);
    }
    // 分頁切點：不切在含「服務期間」行的列中間（切點只落在列與列之間）
    await b.ev(`previewQuote('QB'); 0`);
    await b.waitFor(`!!document.querySelector('#quotePreviewOverlay .qpv-paper')`, 10000);
    const cut = await b.ev(`(()=>{
      const paper=document.querySelector('#quotePreviewOverlay .qpv-paper'); paper.style.zoom=1;
      const pr=paper.getBoundingClientRect(); const rel=(el)=>el.getBoundingClientRect().bottom-pr.top; const total=Math.ceil(pr.height);
      const cand=[]; Array.prototype.forEach.call(paper.children,(el)=>cand.push(rel(el)));
      Array.prototype.forEach.call(paper.querySelectorAll('.qpv-items tbody tr:not(.qpv-grp)'),(el)=>{const nx=el.nextElementSibling; if(nx&&nx.classList.contains('qpv-sub')) return; cand.push(rel(el));});
      Array.prototype.forEach.call(paper.querySelectorAll('.qpv-remarks > div'),(el,i)=>{ if(i>0) cand.push(rel(el)); });
      const pages=_qpdfPlanPages(total,cand,QPV_PAPER_WIDTH/QPDF_A4_MM.w,parseFloat(getComputedStyle(paper).paddingBottom)||0);
      const rows=[...paper.querySelectorAll('.qpv-items tbody tr')].map(tr=>{const r=tr.getBoundingClientRect(); return {top:r.top-pr.top,bot:r.bottom-pr.top,term:!!tr.querySelector('.qpv-term')};});
      const cuts=pages.slice(0,-1).map(p=>p.y1); const inside=cuts.filter(y=>rows.some(r=>y>r.top+1&&y<r.bot-1));
      return JSON.stringify({pages:pages.length,total,cuts:cuts.map(Math.round),inside:inside.length,termRows:rows.filter(r=>r.term).length,maxRow:Math.round(Math.max.apply(null,rows.map(r=>r.bot-r.top)))});
    })()`);
    const cj = JSON.parse(cut);
    t('UI-P3 分頁切點：QB 的 PDF 分頁位置沒有落在任何品項列中間（含「服務期間」行的 34 列都不被截斷）', cj.pages >= 3 && cj.inside === 0 && cj.termRows === 34, cut);
    await b.ev(`closeQuotePreview(); 0`);
    t('UI-P4 預覽／PDF 區段沒有 JS 例外', b.errors.length === 0, J(b.errors).slice(0, 300));
  });
}

// ═════════════════════════ 外觀：四種外觀對比與截圖、手機寬度、後台商品類型 ═════════════════════════
async function secLook(C) {
  const { H, t } = C;
  await H.fresh(mkUsers(H), seedData(C.ROOT, true), { env: { MA_TRACKING: '1' } });
  await H.loginAll(['_s_own', '_s_admin']);
  // 各元素的對比門檻 4.5:1（一般文字）；量測前先把畫面調成「有紅字、有預選提示、有日期」的狀態
  await withBrowser({ cdpPort: C.cdp, maxMs: 560000 }, async (b) => {
    await openAs(b, C, '_s_own', 1440, 900);
    await toQuotes(b);
    await openEdit(b, 'QD');
    await b.click('#addQuoteItemBtn'); await b.sleep(300);
    await typeDesc(b, rowSel(6) + ' .qi-desc', 'Oracle 授權');   // 新列：預選 License＋提示「已依品名預選」＋缺日期紅字
    const r6 = await termOf(b, 6);
    t('UI-L0 外觀測試的準備狀態：新列預選 License、顯示提示、缺日期紅字', r6.kind === 'license' && r6.hint && /請填寫/.test(r6.err), J(r6));
    const ELS = [
      [rowSel(1) + ' .qterm-lab', '欄位標籤（類型／起／迄）'], [rowSel(1) + ' .qterm-kind', '類型下拉文字'], [rowSel(1) + ' .qterm-from', '起日輸入文字'], [rowSel(1) + ' .qterm-to', '迄日輸入文字'],
      [rowSel(1) + ' .qterm-quick', '一年／三年按鈕'], [rowSel(1) + ' .qterm-span', '「共 N 個月」'], [rowSel(6) + ' .qterm-err', '紅字提示'], [rowSel(6) + ' .qterm-hint', '預選提示'],
    ];
    for (const th of THEMES) {
      await theme(b, th);
      const dark = th.endsWith('dark');
      const bad = [];
      for (const [sel, nm] of ELS) {
        const cr = await contrast(b, sel, dark);
        if (cr === null || cr < 4.5) bad.push(nm + '=' + cr);
      }
      t('UI-L1 品項列（' + th + '）：類型／日期／快捷／月數／紅字／提示的文字對比 >= 4.5:1', bad.length === 0, bad.join('、'));
      await b.ev(`document.querySelector(${J(rowSel(1))}).scrollIntoView({block:'center'}); 0`);
      await shotTo(C, b, 'qterm_form_' + th + '.png');
    }
    await theme(b, 'glass-dark');
    await closeModal(b);
    // 簽核面板（唯讀）與預覽
    await b.ev(`openQuoteApproval('QA'); 0`);
    await b.waitFor(`!!document.querySelector('#quoteApprovalOverlay .qterm-view')`, 10000);
    for (const th of THEMES) {
      await theme(b, th);
      const dark = th.endsWith('dark');
      const c1 = await contrast(b, '#quoteApprovalOverlay .qterm-view', dark), c2 = await contrast(b, '#quoteApprovalOverlay .qterm-badge', dark);
      t('UI-L2 簽核面板（' + th + '）：唯讀「服務期間」行與類型徽章對比 >= 4.5:1', c1 >= 4.5 && c2 >= 4.5, c1 + '/' + c2);
      await shotTo(C, b, 'qterm_approval_' + th + '.png');
    }
    await b.key('Escape'); await b.sleep(200);
    await theme(b, 'glass-dark');
    await b.ev(`previewQuote('QD'); 0`);
    await b.waitFor(`!!document.querySelector('#quotePreviewOverlay .qpv-term')`, 10000);
    const cp = await b.ev(`(()=>{const e=document.querySelector('#quotePreviewOverlay .qpv-term'); const f=(s)=>{const m=s.match(/[0-9.]+/g).map(Number); return m.slice(0,3)}; const lum=(c)=>{const g=c.map(v=>{v/=255; return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)}); return 0.2126*g[0]+0.7152*g[1]+0.0722*g[2]}; const fg=f(getComputedStyle(e).color), bg=f(getComputedStyle(document.querySelector('#quotePreviewOverlay .qpv-paper')).backgroundColor); const a=lum(fg), b2=lum(bg); return Math.round((Math.max(a,b2)+0.05)/(Math.min(a,b2)+0.05)*100)/100})()`);
    t('UI-L3 客戶預覽的「服務期間」灰字（預覽紙張固定白底，不隨外觀變）對比 >= 4.5:1', cp >= 4.5, cp);
    await b.ev(`closeQuotePreview(); 0`);
    // 核准後小視窗：四種外觀
    await toQuotes(b);
    await b.click('#quoteTbody [data-qact="terms"][data-id="QA"]');
    await b.waitFor(`!!document.querySelector('#qTermsOverlay .qterm')`, 8000);
    for (const th of THEMES) {
      await theme(b, th);
      const dark = th.endsWith('dark');
      const bad = [];
      for (const sel of ['#qTermsOverlay .qn-name', '#qTermsOverlay .qterm-lab', '#qTermsOverlay .qterm-kind', '#qTermsOverlay .qterm-from', '#qTermsOverlay .qterm-quick', '#qTermsOverlay .qterm-span', '#qTermsOverlay .qterm-info,#qTermsOverlay .q-infobar']) {
        const cr = await contrast(b, sel, dark); if (cr === null || cr < 4.5) bad.push(sel.replace('#qTermsOverlay ', '') + '=' + cr);
      }
      t('UI-L4 「修改服務期間」小視窗（' + th + '）：文字對比 >= 4.5:1', bad.length === 0, bad.join('、'));
      await shotTo(C, b, 'qterm_dialog_' + th + '.png');
    }
    await b.key('Escape'); await b.sleep(200);
    // ── 手機寬度 ──
    await theme(b, 'glass-dark');
    await b.viewport(390, 844, true);
    await b.sleep(500);
    await openEdit(b, 'QD');
    await b.ev(`document.querySelector('#quoteItemsBody tr').scrollIntoView({block:'start'}); 0`);
    const mob = await b.ev(`(()=>{
      const doc=document.documentElement; const out={pageOverflow:doc.scrollWidth>window.innerWidth+1, vw:window.innerWidth, bad:[]};
      document.querySelectorAll('#quoteItemsBody tr:not([data-kind])').forEach((tr,i)=>{ const td=tr.querySelector('.qi-desc').parentNode; const tr0=td.getBoundingClientRect(); tr.querySelectorAll('.qterm *').forEach(e=>{ if(e.offsetParent===null) return; const r=e.getBoundingClientRect(); if(r.width>0 && (r.right>tr0.right+1 || r.left<tr0.left-1)) out.bad.push(i+':'+e.className+':'+Math.round(r.right-tr0.right)); }); });
      const modal=document.querySelector('#quoteModalOverlay .modal'); const mr=modal.getBoundingClientRect(); out.modalRight=Math.round(mr.right); out.modalOver=mr.right>window.innerWidth+1;
      return JSON.stringify(out); })()`);
    const mj = JSON.parse(mob);
    t('UI-L5 手機寬度（390px）：品項列的類型／日期欄不超出品名欄、整頁沒有橫向捲動（日期欄換行）', !mj.pageOverflow && mj.bad.length === 0 && !mj.modalOver, mob);
    await shotTo(C, b, 'qterm_form_mobile.png');
    await closeModal(b);
    await toQuotes(b);
    await b.click('#quoteTbody [data-qact="terms"][data-id="QA"]');
    await b.waitFor(`!!document.querySelector('#qTermsOverlay .qterm')`, 8000);
    const mob2 = await b.ev(`(()=>{ const m=document.querySelector('#qTermsOverlay .modal').getBoundingClientRect(); const bad=[...document.querySelectorAll('#qTermsOverlay .qterm *')].filter(e=>e.offsetParent!==null && e.getBoundingClientRect().right>m.right+1).map(e=>e.className); return JSON.stringify({over:document.documentElement.scrollWidth>window.innerWidth+1,bad}); })()`);
    t('UI-L6 手機寬度：「修改服務期間」小視窗沒有東西超出視窗', JSON.parse(mob2).over === false && JSON.parse(mob2).bad.length === 0, mob2);
    await shotTo(C, b, 'qterm_dialog_mobile.png');
    await b.key('Escape'); await b.sleep(200);
    await b.viewport(1440, 900, false);
    // ── 後台「銷售商品維護」的類型欄：四種外觀 ──
    await b.nav(H.BASE + '/login.html');
    await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:'_s_admin',password:${J(H.PW)}})}).then(r=>r.status)`);
    await b.nav(H.BASE + '/admin.html');
    await b.waitFor(`!!document.querySelector('[data-sec="product-catalog"]')`, 20000);
    await b.ev(`document.querySelector('[data-sec="product-catalog"]').click(); 0`);
    const okPc = await b.waitFor(`!!document.querySelector('#pcPanel .pc-term-row')`, 15000);
    t('UI-L7 後台銷售商品維護：開啟時每個商品多一列「類型＋成交後啟動任務卡」，並有「套用建議」鈕', okPc && (await b.ev(`!document.getElementById('pcSuggestBtn').hidden`)), String(okPc));
    // 套用建議 → 存檔 → 讀回
    await b.click('#pcSuggestBtn');
    await b.sleep(400);
    const nMa = await b.ev(`[...document.querySelectorAll('#pcPanel .pc-term-sel')].filter(s=>s.value==='ma').length`);
    const nLic = await b.ev(`[...document.querySelectorAll('#pcPanel .pc-term-sel')].filter(s=>s.value==='license').length`);
    t('UI-L8 按「套用建議」：依群組標題與品名預選 MA／License（只改表單），任務卡勾選全部維持未勾', nMa > 0 && nLic > 0 && (await b.ev(`[...document.querySelectorAll('#pcPanel [data-pcfld="tstart"]')].filter(c=>c.checked).length`)) === 0, nMa + '/' + nLic);
    for (const th of THEMES) {
      await theme(b, th);
      const dark = th.endsWith('dark');
      const bad = [];
      for (const sel of ['#pcPanel .pc-term-lbl', '#pcPanel .pc-term-sel', '#pcPanel .pc-term-chk', '#pcSuggestBtn']) { const cr = await contrast(b, sel, dark); if (cr === null || cr < 4.5) bad.push(sel + '=' + cr); }
      t('UI-L9 後台商品類型欄（' + th + '）：文字對比 >= 4.5:1', bad.length === 0, bad.join('、'));
      await b.ev(`document.querySelector('#pcPanel .pc-term-row').scrollIntoView({block:'center'}); 0`);
      await shotTo(C, b, 'qterm_admin_catalog_' + th + '.png');
    }
    t('UI-L10 外觀區段沒有 JS 例外', b.errors.length === 0, J(b.errors).slice(0, 300));
  });
}

// ═════════════════════════ 進入點 ═════════════════════════
async function run(C, only) {
  C.shotsOk = !!C.shots;
  const want = (n) => !only || only === n;
  if (want('off')) await secOff(C);
  if (want('new')) await secNew(C);
  if (want('edit')) await secEdit(C);
  if (want('approved')) await secApproved(C);
  if (want('docs')) await secDocs(C);
  if (want('look')) await secLook(C);
}
module.exports = { run };
