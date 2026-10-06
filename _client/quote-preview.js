// ═════════════════════════════════════════════════
// ── 報價單預覽 (quote-preview.js) ───────────────────────────
// 以 HTML 重現 templates/quotation_template.xlsx 的版面，讓使用者「下載前先看」。
// 這是示意圖：伺服器上沒有 Excel 可即時轉圖，實際成品以下載的 Excel 為準。
//
// 與 lib/quoteExcel.js 保持一致的地方（改其中一邊時請一起改）：
//   · 欄位對應（左框＝客戶資料；右框＝日期＋聯絡人＋手機）
//   · 優惠規則（percent 僅 0<值<100 才生效；amount 需 >0）、稅金 5%
//   · 公司抬頭與 Remarks 第 1~6 條是範本內的固定文字，範本改了這裡要同步
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

/** 預覽上方的簽核警示條：未核准／核准已失效時提醒「下載檔案不會有報價專用章」 */
function _qpvApprovalNotice(q) {
  if (_qpvStamped(q)) return '';
  const a = q && q.approval;
  const voided = a && a.state === 'approved' && a.valid === false;
  return '<div class="qpv-warn qpv-unsigned"><span>' +
    (voided ? '核准已失效（核准後報價內容被修改）：' : '主管尚未簽核完成：') +
    '下載檔案不會有報價專用章。</span></div>';
}

/** 金額與優惠：規則與 lib/quoteExcel.js 一致 */
function _qpvTotals(q) {
  const items = Array.isArray(q.items) ? q.items.slice(0, 50) : [];
  const sub = items.reduce((s, it) => s + (parseFloat(it.qty) || 1) * (parseFloat(it.unitPrice) || 0), 0);
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

  const itemRows = (T.items.length ? T.items : [null]).map((it, i) => {
    if (!it) return '<tr><td class="c">&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>';
    const qty = parseFloat(it.qty) || 1, price = parseFloat(it.unitPrice) || 0;
    return `<tr><td class="c">${i + 1}</td><td class="desc">${e(it.desc || '')}</td><td class="c">${e(String(qty))}</td>` +
           `<td class="c">${e(it.unit || '式')}</td><td class="r">${n(price)}</td><td></td><td class="r">${n(qty * price)}</td></tr>`;
  }).join('');

  const note = String(q.note || '').replace(/\s*[\r\n]+\s*/g, ' ').trim();
  // Remarks 第 4 條：把範本的「XXXX年XX月XX日」換成報價期限（與 Excel 一致，年月日不補零）
  const vuM = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(q.validUntil || ''));
  const remarkList = QPV_REMARKS.map(t => vuM ? t.replace(/X{4}年X{2}月X{2}日/, `${+vuM[1]}年${+vuM[2]}月${+vuM[3]}日`) : t);
  const remarks = remarkList.concat(note ? ['7.' + note] : []).map(t => `<div>${e(t)}</div>`).join('');

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
  stage.firstElementChild.style.zoom = Math.min(1, stage.clientWidth / QPV_PAPER_WIDTH);
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
        <div class="qpv-hint">這是示意預覽，只顯示客戶看得到的內容（不含成本與毛利）。日期為現在送出會蓋的台灣當天日期；實際成品以「下載 Excel」為準。</div>
        <div class="qpv-stage" id="qpvStage">${buildQuotePreviewHtml(q, info)}</div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-secondary" onclick="closeQuotePreview()">關閉</button>
        <button class="btn btn-export" id="qpvExportBtn" type="button">&#11015; 下載 Excel</button>
      </div>
    </div>`;
  document.body.appendChild(ov);
  // 不把 id/單號拼進 inline JS（HTML 屬性的實體解碼會讓單引號跳脫失效）→ 用 listener 帶閉包值
  const exBtn = ov.querySelector('#qpvExportBtn');
  if (exBtn) exBtn.addEventListener('click', function () { exportQuote(q.id, q.quoteNo || ''); });
  fitQuotePreview();
  window.addEventListener('resize', fitQuotePreview);
}
