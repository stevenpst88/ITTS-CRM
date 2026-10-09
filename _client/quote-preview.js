// ═════════════════════════════════════════════════
// ── 報價單預覽 (quote-preview.js) ───────────────────────────
// 以 HTML 重現 templates/quotation_template.xlsx 的版面，讓使用者「下載前先看」。
// 這是示意圖：伺服器上沒有 Excel 可即時轉圖，實際成品以下載的 Excel 為準。
//
// 與 lib/quoteExcel.js 保持一致的地方（改其中一邊時請一起改）：
//   · 欄位對應（左框＝客戶資料；右框＝日期＋聯絡人＋手機）
//   · 優惠規則（percent 僅 0<值<100 才生效；amount 需 >0）、稅金 5%
//   · 品項下方的「說明」（灰色小字）與「備註：…」（棕紅色一行）：沒填就完全不輸出任何額外標記（沒有說明／備註的報價單，預覽與以前位元級相同）；
//     文字清洗規則與 lib/quoteItems.js cleanItemText 相同（_qpvCleanText 是鏡像，scripts/check-quote-item-notes.js 逐例比對）
//   · 公司抬頭是範本內的固定文字，範本改了這裡要同步；Remarks 條款不在這裡維護——一律用 issue-info 的 remarks（lib/quoteRemarks.js 單一來源）
// 預覽只顯示客戶看得到的內容，不含成本與毛利。
// ═════════════════════════════════════════════════

const QPV_LETTERHEAD = {
  name: '東捷資訊服務股份有限公司',
  addr: '台北市南港區三重路19-8號5樓',
  tel:  'Tel:(02) 2655-2525  fax: (02) 2655-1010',
};
const QPV_REMARKS = [
  '1.以上報價為東捷所提供之特惠價，基於誠信原則，本報價單相關條款合約及價格雙方均不得向第三者揭露。',
  '2.ABAP客製開發人天，每人天12,000元（未稅）計價',
  '3.付款方式：簽約完成後付款總金額 100%，月結30天付款。',
  '4.本報價單於XXXX年XX月XX日前有效。',
  '5.本報價單經簽名回傳視同為正式有效之合約。',
  '6.本報價單需加蓋東捷資訊服務(股)公司之"報價專用章"使為有效之正式報價單。',
];
const QPV_PAPER_WIDTH = 900;

const QPV_CSS = `
.qpv-modal { width: 980px; max-width: 96vw; }
.qpv-body { background: #e9ecf1; padding: 16px; }
.qpv-hint { font-size: 12px; color: #5f6b7a; margin-bottom: 10px; line-height: 1.6; }
.qpv-stage { overflow: auto; }
.qpv-warn { display: flex; gap: 12px; align-items: center; justify-content: space-between; background: #fff4e5; color: #8a4b00;
  border: 1px solid #f5c98b; border-radius: 8px; padding: 8px 12px; margin-bottom: 10px; font-size: 12.5px; }
.qpv-paper { width: ${QPV_PAPER_WIDTH}px; margin: 0 auto; box-sizing: border-box; background: #fff; color: #111;
  padding: 34px 38px 42px; box-shadow: 0 2px 14px rgba(0,0,0,.18);
  font-family: "Microsoft JhengHei","微軟正黑體","PingFang TC","Noto Sans TC",sans-serif; font-size: 13px; line-height: 1.55; }
.qpv-head { display: flex; justify-content: space-between; align-items: flex-start; }
.qpv-brand { display: flex; gap: 14px; align-items: flex-start; }
.qpv-brand img { height: 40px; margin-top: 4px; }
.qpv-brand .nm { font-size: 17px; font-weight: 700; }
.qpv-brand .ln { font-size: 12.5px; }
.qpv-title { text-align: center; min-width: 250px; }
.qpv-title .t { font-size: 24px; font-weight: 700; letter-spacing: 2px; }
.qpv-title .no { font-weight: 700; margin-top: 18px; font-size: 14px; }
.qpv-boxes { display: flex; justify-content: space-between; margin-top: 14px; }
.qpv-box { position: relative; border: 1.5px solid #111; padding: 14px 10px 8px; box-sizing: border-box; }
.qpv-box .lg { position: absolute; top: -11px; left: 50%; transform: translateX(-50%); background: #fff; padding: 0 10px; font-weight: 700; white-space: nowrap; }
.qpv-box.w42 { width: 42%; }
.qpv-row { display: flex; min-height: 22px; align-items: baseline; }
.qpv-l { display: inline-block; min-width: 4.2em; flex-shrink: 0; }
.qpv-v { word-break: break-all; }
.qpv-ext { margin-left: auto; margin-right: 32%; }
.qpv-cur { text-align: right; margin: 14px 4px 3px; }
.qpv-items { width: 100%; border-collapse: collapse; table-layout: fixed; }
.qpv-items th, .qpv-items td { border: 1.5px solid #111; padding: 4px 6px; vertical-align: middle; }
.qpv-items th { font-weight: 700; text-align: center; white-space: nowrap; padding-left: 2px; padding-right: 2px; }
.qpv-items td.c { text-align: center; }
.qpv-items td.r { text-align: right; }
.qpv-items td.desc { white-space: pre-wrap; word-break: break-word; }
.qpv-items td.blank { color: #333; }
.qpv-items td.desc .qpv-spec { font-size: 0.82em; color: #6b7280; line-height: 1.45; margin-top: 2px; }
.qpv-items td.desc .qpv-note { font-size: 0.9em; color: #9a3412; line-height: 1.45; margin-top: 2px; }
.qpv-items tr.qpv-grp td { background: #f2f2f2; font-weight: 700; text-align: left; padding: 5px 10px; white-space: pre-wrap; word-break: break-word; }
.qpv-items tr.qpv-sub td { font-weight: 700; }
.qpv-items tr.qpv-sub td.lbl { text-align: right; padding-right: 8px; }
.qpv-sum { display: flex; justify-content: space-between; align-items: flex-start; }
.qpv-sum .note { padding: 4px 0 0 8%; }
.qpv-sumt { width: 40%; border-collapse: collapse; table-layout: fixed; margin-top: -1.5px; }
.qpv-sumt td { border: 1.5px solid #111; padding: 2px 8px; font-weight: 700; text-align: right; }
.qpv-sumt td:first-child { width: 62%; }
.qpv-proj { width: 34%; margin-top: 34px; padding-top: 18px; padding-bottom: 14px; }
.qpv-proj .qpv-row { min-height: 34px; }
.qpv-remarks { margin-top: 18px; }
.qpv-remarks div { min-height: 22px; }
.qpv-sign { display: flex; justify-content: space-between; margin-top: 16px; }
.qpv-sign .col { width: 45%; }
.qpv-sign .sig { position: relative; height: 88px; display: flex; align-items: center; justify-content: center; }
.qpv-sign .sig .qpv-seal { position: absolute; right: 14%; top: 2px; width: 84px; height: 84px; object-fit: contain; mix-blend-mode: multiply; pointer-events: none; }
.qpv-sign .ln { border-top: 3px solid #111; padding-top: 3px; margin-top: 6px; }
.qpv-warn.qpv-unsigned { background: #fff0f0; color: #a31515; border-color: #f0b4b4; }
body.dark .qpv-body { background: #0d1117; }
body.dark .qpv-hint { color: #8b949e; }
`;

function _qpvEnsureStyle() {
  if (document.getElementById('qpvStyle')) return;
  const st = document.createElement('style');
  st.id = 'qpvStyle';
  st.textContent = QPV_CSS;
  document.head.appendChild(st);
}

/** 已核准且核准仍有效（內容雜湊吻合）才蓋章；與下載 Excel 的蓋章條件一致 */
function _qpvStamped(q) {
  const a = q && q.approval;
  return !!(a && a.state === 'approved' && a.valid === true);
}
let _qpvSealVer = Date.now();   // 每次開預覽換一個，避免瀏覽器快取到舊章

/** 預覽上方的簽核警示條：未核准／核准已失效時提醒「下載的 PDF 不會有報價專用章」（Excel 則一律不蓋章） */
function _qpvApprovalNotice(q) {
  if (_qpvStamped(q)) return '';
  const a = q && q.approval;
  const voided = a && a.state === 'approved' && a.valid === false;
  return '<div class="qpv-warn qpv-unsigned"><span>' +
    (voided ? '核准已失效（核准後報價內容被修改）：' : '主管尚未簽核完成：') +
    '下載的 PDF 不會有報價專用章（Excel 一律不蓋章）。</span></div>';
}

/**
 * 分組標題／小計列（kind，與 lib/quoteItems.js 一致；舊單沒有 kind＝一般品項）。
 * 小計值＝「上一個標題或小計列之後」的品項（數量空→1）合計，鏡像 lib/quoteItems.js 的 subtotalValues；scripts/check-quote-items.js 會逐例比對。
 */
function _qpvIsKindRow(it) { return !!it && (it.kind === 'title' || it.kind === 'subtotal'); }
function _qpvSubtotals(items) {
  const out = [];
  let acc = 0;
  (items || []).forEach((it) => {
    if (_qpvIsKindRow(it)) { out.push(it.kind === 'subtotal' ? acc : null); acc = 0; }
    else { out.push(null); acc += (parseFloat(it && it.qty) || 1) * (parseFloat(it && it.unitPrice) || 0); }
  });
  return out;
}

/** 品項的說明／備註文字（單行純文字）。鏡像 lib/quoteItems.js 的 cleanItemText；標題／小計列沒有。max：說明 200、備註 300 */
const QPV_SPEC_RE = new RegExp('[\\u0000-\\u001F\\u007F-\\u009F' + String.fromCharCode(0x2028) + String.fromCharCode(0x2029) + ']+', 'g');
function _qpvCleanText(v, max) {
  if (typeof v !== 'string') return '';
  const t = v.replace(QPV_SPEC_RE, ' ').replace(/\s+/g, ' ').trim();
  const cps = Array.from(t);
  return cps.length > max ? cps.slice(0, max).join('').trim() : t;
}
function _qpvItemTexts(it) {
  if (!it || typeof it !== 'object' || _qpvIsKindRow(it)) return { spec: '', note: '' };
  return { spec: _qpvCleanText(it.spec, 200), note: _qpvCleanText(it.note, 300) };
}

/** 金額與優惠：規則與 lib/quoteExcel.js 一致 */
function _qpvTotals(q) {
  const items = Array.isArray(q.items) ? q.items.slice(0, 50) : [];
  const sub = items.reduce((s, it) => _qpvIsKindRow(it) ? s : s + (parseFloat(it.qty) || 1) * (parseFloat(it.unitPrice) || 0), 0);
  const dv = parseFloat(q.discountValue) || 0;
  let disc = sub, note = '';
  if (q.discountType === 'percent' && dv > 0 && dv < 100) {
    disc = sub * dv / 100;
    note = `專案優惠 ${dv}%（${+(dv / 10).toFixed(1)} 折）`;
  } else if (q.discountType === 'amount' && dv > 0) {
    disc = dv;
    note = '專案議價金額';
  }
  const tax = disc * 0.05;
  return { items, sub, disc, tax, total: disc + tax, note };
}

/**
 * info＝GET /quotations/:id/issue-info 的結果 {issueDate, issuer}。
 * 日期一律用「現在送出會蓋的台灣當天」，不用報價單上存的日期；右上「廠商資料」框＝我方業務的聯絡資訊。
 */
function buildQuotePreviewHtml(q, info) {
  const e = escapeHtml;
  const n = (v) => Math.round(v || 0).toLocaleString('en-US');
  const T = _qpvTotals(q);
  const iss = (info && info.issuer) || {};
  const date = String((info && info.issueDate) || taipeiTodayClient()).replace(/-/g, '/');
  const row = (label, val) => `<div class="qpv-row"><span class="qpv-l">${label}</span><span class="qpv-v">${e(val || '')}</span></div>`;

  const subVals = _qpvSubtotals(T.items);
  let seq = 0;   // 項目編號只算一般品項（標題／小計列不編號），與 Excel 一致
  const itemRows = (T.items.length ? T.items : [null]).map((it, i) => {
    if (!it) return '<tr><td class="c">&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>';
    if (it.kind === 'title') return `<tr class="qpv-grp"><td colspan="7">${e(it.desc || '')}</td></tr>`;
    if (it.kind === 'subtotal') {
      const label = String(it.desc || '').replace(/\s+/g, ' ').trim() || '小計';
      return `<tr class="qpv-sub"><td colspan="6" class="lbl">${e(label)}</td><td class="r">${n(subVals[i])}</td></tr>`;
    }
    const qty = parseFloat(it.qty) || 1, price = parseFloat(it.unitPrice) || 0;
    const tx = _qpvItemTexts(it);
    const extra = (tx.spec ? `<div class="qpv-spec">${e(tx.spec)}</div>` : '') + (tx.note ? `<div class="qpv-note"><b>備註：</b>${e(tx.note)}</div>` : '');
    return `<tr><td class="c">${++seq}</td><td class="desc">${e(it.desc || '')}${extra}</td><td class="c">${e(String(qty))}</td>` +
           `<td class="c">${e(it.unit || '式')}</td><td class="r">${n(price)}</td><td></td><td class="r">${n(qty * price)}</td></tr>`;
  }).join('');

  // Remarks：直接用伺服器組好的條款（info.remarks，與 Excel 同一份：固定條文、付款方式、報價期限、追加條款）。
  // 沒有（舊版伺服器）才退回本檔的固定條文＋單行備註。
  let remarkLines;
  if (info && Array.isArray(info.remarks) && info.remarks.length) remarkLines = info.remarks;
  else {
    const note = String(q.note || '').replace(/\s*[\r\n]+\s*/g, ' ').trim();
    const vuM = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(q.validUntil || ''));
    // 追加條款以 serialize 回傳的 q.clauses 為準（有新格式＝clauses；舊單＝單行 note）；付款方式用 quote.js 的鏡像函式
    const extras = Array.isArray(q.clauses) ? q.clauses : (note ? [note] : []);
    const base = QPV_REMARKS.map(t => vuM ? t.replace(/X{4}年X{2}月X{2}日/, `${+vuM[1]}年${+vuM[2]}月${+vuM[3]}日`) : t);
    if (typeof quotePaymentSentence === 'function') base[2] = quotePaymentSentence(q.payment);
    remarkLines = base.concat(extras.map((t, i) => (7 + i) + '.' + t));
  }
  const remarks = remarkLines.map(t => `<div>${e(t)}</div>`).join('');

  return `<div class="qpv-paper">
    <div class="qpv-head">
      <div class="qpv-brand">
        <img src="/itts-logo.png" alt="ITTS">
        <div><div class="nm">${e(QPV_LETTERHEAD.name)}</div><div class="ln">${e(QPV_LETTERHEAD.addr)}</div><div class="ln">${e(QPV_LETTERHEAD.tel)}</div></div>
      </div>
      <div class="qpv-title"><div class="t">報價單</div><div class="no">表單編號：${e(q.quoteNo || '')}</div></div>
    </div>
    <div class="qpv-boxes">
      <div class="qpv-box w42"><span class="lg">客戶資料</span>
        ${row('公　司：', q.company)}${row('聯絡人：', q.contactName)}${row('地　址：', q.address)}${row('電　話：', q.phone)}
      </div>
      <div class="qpv-box w42"><span class="lg">廠商資料</span>
        ${row('日　期：', date)}${row('聯絡人：', iss.name)}${row('手　機：', iss.mobile)}
        <div class="qpv-row"><span class="qpv-l">電　話：</span><span class="qpv-v">${e(iss.phone || '')}</span><span class="qpv-ext">Ext.${e(iss.ext || '')}</span></div>
      </div>
    </div>
    <div class="qpv-cur">幣別：TWD</div>
    <table class="qpv-items">
      <colgroup><col style="width:4.5%"><col style="width:49%"><col style="width:5.5%"><col style="width:5.5%"><col style="width:11.5%"><col style="width:12.5%"><col style="width:11.5%"></colgroup>
      <thead><tr><th>項目</th><th>內容</th><th>數量</th><th>單位</th><th>定價(未稅)</th><th>優惠價(未稅)</th><th>小計(未稅)</th></tr></thead>
      <tbody>${itemRows}<tr><td></td><td class="blank">以下空白</td><td></td><td></td><td></td><td></td><td></td></tr></tbody>
    </table>
    <div class="qpv-sum">
      <div class="note">${e(T.note)}</div>
      <table class="qpv-sumt">
        <tr><td>專案定價（未稅）</td><td>${n(T.sub)}</td></tr>
        <tr><td>專案優惠價(未稅)</td><td>${n(T.disc)}</td></tr>
        <tr><td>稅金5%</td><td>NT$${n(T.tax)}</td></tr>
        <tr><td>專案優惠價(含稅)</td><td>NT$${n(T.total)}</td></tr>
      </table>
    </div>
    <div class="qpv-box qpv-proj"><span class="lg">專案資料</span>
      ${row('專案名稱：', q.projectName)}
    </div>
    <div class="qpv-remarks"><div>Remarks ：</div>${remarks}</div>
    <div class="qpv-sign">
      <div class="col"><div>Customer Confirme by:</div><div class="sig">&nbsp;</div><div class="ln">請簽回以確認訂單</div></div>
      <div class="col"><div>Prepared by :</div><div class="sig">${_qpvStamped(q) ? `<img class="qpv-seal" src="${API}/quotations/${encodeURIComponent(q.id)}/seal?v=${_qpvSealVer}" alt="報價專用章" onerror="this.style.display='none'">` : '&nbsp;'}</div><div class="ln">ITTS Corp.</div></div>
    </div>
  </div>`;
}

/** 紙張固定 900px 寬；視窗較窄（手機）時等比縮小，避免橫向捲動才看得到右側欄位 */
function fitQuotePreview() {
  const stage = document.getElementById('qpvStage');
  if (!stage || !stage.firstElementChild) return;
  // clientWidth 為 0（視窗還沒排版或分頁在背景）時維持 1，不能把內容縮成 0 而整張消失
  stage.firstElementChild.style.zoom = stage.clientWidth ? Math.min(1, stage.clientWidth / (parseFloat(stage.dataset.paperWidth) || QPV_PAPER_WIDTH)) : 1;
}

let _qpvCurrentId = null;   // 目前開著的預覽（儲存聯絡資訊後要刷新它）

function closeQuotePreview() {
  const ov = document.getElementById('quotePreviewOverlay');
  if (ov) ov.remove();
  _qpvCurrentId = null;
  window.removeEventListener('resize', fitQuotePreview);
}

/** 出單資訊（台灣當天日期＋這張單的業務聯絡資訊）；取不到就回 null，預覽照樣顯示、只是右框空白 */
async function _qpvLoadInfo(id) {
  try { const r = await fetch(`${API}/quotations/${id}/issue-info`); if (r.ok) return await r.json(); } catch (e) { /* 用空白聯絡資訊 */ }
  return null;
}

/** 聯絡資訊還沒維護時的提示：自己的單給「立即設定」；別人的單請對方維護 */
function _qpvNotice(info, q) {
  const approvalWarn = _qpvApprovalNotice(q);
  return approvalWarn + _qpvExpiryNotice(info, q) + _qpvIssuerNotice(info);
}

/** 報價期限已早於「今天」（出單日）→ 提醒業務更新期限；核准後改期限會使核准作廢，所以只提醒、不擋 */
function _qpvExpiryNotice(info, q) {
  const vu = q && q.validUntil;
  const today = String((info && info.issueDate) || taipeiTodayClient());
  if (!vu || vu >= today) return '';
  return '<div class="qpv-warn"><span>報價期限（' + escapeHtml(vu.replace(/-/g, '/')) + '）已早於今天，客戶收到的報價單會是已過期的。請編輯報價單更新報價期限' +
    (q.approval && q.approval.state === 'approved' ? '（已核准的單修改後需重新簽核）' : '') + '。</span></div>';
}

function _qpvIssuerNotice(info) {
  if (!info || !info.issuer || info.issuer.isSet) return '';
  return info.ownerIsMe
    ? '<div class="qpv-warn"><span>你還沒維護聯絡資訊，右上「廠商資料」框的電話與手機會是空白。</span><button class="btn btn-sm btn-primary" onclick="openMyContactModal()">立即設定</button></div>'
    : '<div class="qpv-warn"><span>這位業務還沒維護聯絡資訊，右上「廠商資料」框的電話與手機會是空白。</span></div>';
}

/** 重新載入並重畫預覽（儲存聯絡資訊後呼叫） */
async function refreshQuotePreview() {
  const stage = document.getElementById('qpvStage');
  if (!stage || !_qpvCurrentId) return;
  const q = (typeof allQuotations !== 'undefined' ? allQuotations : []).find(x => x.id === _qpvCurrentId);
  if (!q) return;
  const info = await _qpvLoadInfo(_qpvCurrentId);
  stage.innerHTML = buildQuotePreviewHtml(q, info);
  const warn = document.getElementById('qpvNotice');
  if (warn) warn.innerHTML = _qpvNotice(info, q);
  fitQuotePreview();
}

/**
 * 毛利分析預覽（內部）：伺服器把「實際要下載的毛利分析 xlsx」轉成 HTML（GET /quotations/:id/pnl-preview），
 * 所以看到的就是下載檔的內容；含成本與毛利，只有「看得到成本與價格」的人才有權限（與下載相同）。
 */
async function previewQuotePnl(id) {
  let r, j = {};
  try { r = await fetch(`${API}/quotations/${encodeURIComponent(id)}/pnl-preview`); j = await r.json().catch(() => ({})); } catch (e) { return showToast('毛利分析預覽載入失敗，請重試'); }
  if (!r.ok || !j.html) return showToast(j.error || '無法預覽毛利分析');
  showPnlPreviewModal(j, { id });
}

/**
 * 顧問「毛利預覽」（cost-sync）：用顧問畫面上目前的成本明細（含尚未儲存的調整）請伺服器產生毛利分析並轉成 HTML（POST cost-draft/pnl-preview，不寫入）。
 * body＝{ costLines, newItems?, contingencyPct? }。顯示的版面與已存檔版相同，但沒有下載鈕（草稿還不是報價內容；要下載請在完成成本後從報價單列表進行）。
 * 回傳 true＝已開啟預覽；false＝失敗（已提示）。
 */
async function previewQuotePnlDraft(id, body) {
  let r, j = {};
  try {
    r = await fetch(`${API}/quotations/${encodeURIComponent(id)}/cost-draft/pnl-preview`, {
      method: 'POST', credentials: 'same-origin', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}),
    });
    j = await r.json().catch(() => ({}));
  } catch (e) { showToast('毛利預覽載入失敗，請重試'); return false; }
  if (!r.ok || !j.html) { showToast(j.error || '無法產生毛利預覽'); return false; }
  showPnlPreviewModal(j, { id, draft: true });
  return true;
}

/** 毛利分析預覽視窗（已存檔版與顧問草稿版共用）。opts.draft＝草稿：標明「含尚未儲存的調整」，不提供下載 */
function showPnlPreviewModal(j, opts) {
  opts = opts || {};
  _qpvEnsureStyle();
  closeQuotePreview();
  const ov = document.createElement('div');
  ov.className = 'modal-overlay open';
  ov.id = 'quotePreviewOverlay';
  if (opts.draft) ov.setAttribute('data-draft', '1');
  ov.innerHTML = `
    <div class="modal qpv-modal">
      <div class="modal-header">
        <h2>毛利分析預覽（內部）${opts.draft ? '：草稿' : ''}　${escapeHtml(j.quoteNo || '')}</h2>
        <button class="modal-close" onclick="closeQuotePreview()">&#10005;</button>
      </div>
      <div class="modal-body qpv-body">
        <div class="qpv-warn qpv-unsigned"><span>內部文件：含成本與毛利率，請勿提供客戶。</span></div>
        ${opts.draft ? '<div class="qpv-warn"><span>這是依你畫面上<b>目前的調整（含尚未儲存的內容）</b>產生的草稿預覽：報價品項的數量／單位已套用「連動」的結果，新增的報價項目單價尚未填寫（由業務補填）。草稿預覽沒有下載鈕。</span></div>' : ''}
        <div class="qpv-hint">這是「毛利分析(內部)」Excel 內容的網頁預覽（和下載檔同一份資料；範本上的簽名線、選項按鈕等圖形不會顯示），實際成品以下載的 Excel 為準。</div>
        <div class="qpv-stage" id="qpvStage" data-paper-width="${Number(j.widthPx) || 950}"><div class="qpv-pnl-sheet" style="width:${Number(j.widthPx) || 950}px;margin:0 auto;background:#fff;box-shadow:0 2px 14px rgba(0,0,0,.18)">${j.html}</div></div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-secondary" onclick="closeQuotePreview()">關閉</button>
        ${opts.draft ? '' : '<button class="btn btn-export" id="qpvPnlDlBtn" type="button">&#11015; 下載毛利分析 Excel</button>'}
      </div>
    </div>`;
  document.body.appendChild(ov);
  const dl = ov.querySelector('#qpvPnlDlBtn');
  if (dl) dl.addEventListener('click', function () { exportQuote(opts.id, j.quoteNo || '', 'pnl'); });
  fitQuotePreview();
  window.addEventListener('resize', fitQuotePreview);
}

/**
 * 顧問「報價單預覽」（cost-sync）：q＝已把連動後的品項（數量／單位）與顧問新增的報價項目套進去的報價單物件（前端用 cost-draft/summary 回傳的 items 覆蓋），
 * info＝issue-info。走和一般預覽相同的渲染（buildQuotePreviewHtml），但是草稿：沒有下載鈕、沒有「主管尚未簽核」警示（顧問看不到簽核資訊），並標明含未儲存的調整。
 */
function showQuoteDraftPreview(q, info) {
  _qpvEnsureStyle();
  closeQuotePreview();
  const no = escapeHtml(q.quoteNo || '');
  const ov = document.createElement('div');
  ov.className = 'modal-overlay open';
  ov.id = 'quotePreviewOverlay';
  ov.setAttribute('data-draft', '1');
  const news = (q.items || []).filter(function (it) { return it && it.needPrice === true && !(parseFloat(it.unitPrice) > 0); }).length;
  ov.innerHTML = `
    <div class="modal qpv-modal">
      <div class="modal-header">
        <h2>報價單預覽：草稿　${no}</h2>
        <button class="modal-close" onclick="closeQuotePreview()">&#10005;</button>
      </div>
      <div class="modal-body qpv-body">
        <div id="qpvNotice"><div class="qpv-warn"><span>這是依你畫面上<b>目前的調整（含尚未儲存的內容）</b>產生的草稿預覽：品項的數量／單位已套用「連動」的結果${news ? '；新增的報價項目（' + news + ' 項）單價是 0，由業務補填' : ''}。尚未同步給業務，按「完成並通知業務」才會寫回報價。</span></div>${_qpvIssuerNotice(info)}</div>
        <div class="qpv-hint">這是示意預覽，只顯示客戶看得到的內容（不含成本與毛利）。日期為現在送出會蓋的台灣當天日期；草稿預覽沒有下載鈕。</div>
        <div class="qpv-stage" id="qpvStage">${buildQuotePreviewHtml(q, info)}</div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-secondary" onclick="closeQuotePreview()">關閉</button>
      </div>
    </div>`;
  document.body.appendChild(ov);
  fitQuotePreview();
  window.addEventListener('resize', fitQuotePreview);
}

async function previewQuote(id) {
  // 蓋章與否取決於「現在」的核准狀態，所以一律先向伺服器取最新的單；取不到才退回列表快取
  let q = null;
  try { const r = await fetch(`${API}/quotations/${encodeURIComponent(id)}`); if (r.ok) q = await r.json(); } catch (e) { /* 退回列表快取 */ }
  const list = (typeof allQuotations !== 'undefined' ? allQuotations : []);
  if (q) {
    const i = list.findIndex(x => x.id === id);
    if (i >= 0) list[i] = q;
  } else {
    q = list.find(x => x.id === id) || null;
  }
  if (!q) return showToast('找不到此報價單');
  const info = await _qpvLoadInfo(id);

  _qpvEnsureStyle();
  closeQuotePreview();
  _qpvCurrentId = id;
  _qpvSealVer = Date.now();
  const no = escapeHtml(q.quoteNo || '');
  const ov = document.createElement('div');
  ov.className = 'modal-overlay open';
  ov.id = 'quotePreviewOverlay';
  ov.innerHTML = `
    <div class="modal qpv-modal">
      <div class="modal-header">
        <h2>報價單預覽　${no}</h2>
        <button class="modal-close" onclick="closeQuotePreview()">&#10005;</button>
      </div>
      <div class="modal-body qpv-body">
        <div id="qpvNotice">${_qpvNotice(info, q)}</div>
        <div class="qpv-hint">這是示意預覽，只顯示客戶看得到的內容（不含成本與毛利）。日期為現在送出會蓋的台灣當天日期；實際成品以下載的檔案為準。報價專用章只會出現在 PDF 與此預覽，Excel 一律不蓋章。</div>
        <div class="qpv-stage" id="qpvStage">${buildQuotePreviewHtml(q, info)}</div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-secondary" onclick="closeQuotePreview()">關閉</button>
        <button class="btn btn-export" id="qpvExportBtn" type="button">&#11015; 下載 Excel</button>
        <button class="btn btn-export" id="qpvPdfBtn" type="button">&#11015; 下載 PDF</button>
      </div>
    </div>`;
  document.body.appendChild(ov);
  // 不把 id/單號拼進 inline JS（HTML 屬性的實體解碼會讓單引號跳脫失效）→ 用 listener 帶閉包值
  const exBtn = ov.querySelector('#qpvExportBtn');
  if (exBtn) exBtn.addEventListener('click', function () { exportQuote(q.id, q.quoteNo || ''); });
  const pdfBtn = ov.querySelector('#qpvPdfBtn');
  if (pdfBtn) pdfBtn.addEventListener('click', function () {
    if (pdfBtn.disabled) return;
    pdfBtn.disabled = true;   // 產生需要幾秒，避免連按產生多份
    Promise.resolve(exportQuotePdf(q.id, q.quoteNo || '')).then(function () { pdfBtn.disabled = false; }, function () { pdfBtn.disabled = false; });
  });
  fitQuotePreview();
  window.addEventListener('resize', fitQuotePreview);
}
