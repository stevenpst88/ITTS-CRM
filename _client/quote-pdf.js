/**
 * 報價單 PDF——在瀏覽器端產生並直接下載（伺服器沒有中文字型，Vercel 上也不方便內嵌）。
 *
 * 做法：把「網頁預覽」(quote-preview.js 的 buildQuotePreviewHtml，與 Excel 同一份條款／金額／版面) 畫在畫面外的容器，
 * 用 html2canvas 轉成圖，依 A4 切頁（優先在表格列／條款之間切，不切在字中間）後用 jsPDF 組成 PDF。
 * - 兩個函式庫放在 _client/vendor/（版本寫在檔名，不走外部 CDN；首次按下 PDF 才載入）。
 * - PDF 是「圖片式」：文字不能選取、不能複製；換來的是中文用瀏覽器內建字型就能正確顯示、不必改伺服器。
 * - 報價專用章用 multiply 疊印（html2canvas 不支援 mix-blend-mode，所以章在畫完後另外疊上去）。
 * - 規則：報價專用章只出現在 PDF（與網頁預覽），Excel 一律不蓋章（伺服器 GET /quotations/:id/export 不嵌章）。
 *   數量、金額等一經調整，核准即作廢需重新簽核，所以有章的正式版只有「核准有效當下」的 PDF。
 * - 下載動作伺服器看不到，所以產生前先 POST /quotations/:id/export-log 記稽核紀錄。
 */
const QPDF_LIBS = ['/vendor/html2canvas-1.4.1.min.js', '/vendor/jspdf-4.2.1.umd.min.js'];
const QPDF_SCALE = 2;                      // 解析度倍率：2 倍在 A4 上約 220 dpi，文字清晰、單頁約 300~500 KB
// 畫布上限：iOS Safari 約 16.7M 像素、單邊約 16384，超過會「靜默」變成空白圖。超長報價單（品項多、說明長）自動降低倍率，留餘裕
const QPDF_MAX_PIXELS = 16000000;
const QPDF_MAX_SIDE = 16000;
const QPDF_A4_MM = { w: 210, h: 297 };
const QPDF_MARGIN_MM = 10;                 // 第 2 頁起的上緣、與每頁的下緣留白（第 1 頁的上緣由紙張本身的內距提供）

let _qpdfLibsPromise = null;

function _qpdfLoadScript(src) {
  return new Promise(function (resolve, reject) {
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = resolve;
    s.onerror = function () { reject(new Error('載入 PDF 元件失敗（' + src + '）')); };
    document.head.appendChild(s);
  });
}

function _qpdfEnsureLibs() {
  if (window.html2canvas && window.jspdf && window.jspdf.jsPDF) return Promise.resolve();
  if (!_qpdfLibsPromise) {
    _qpdfLibsPromise = Promise.all(QPDF_LIBS.map(_qpdfLoadScript)).then(function () { /* 載入完成 */ }, function (e) { _qpdfLibsPromise = null; throw e; });
  }
  return _qpdfLibsPromise;
}

/**
 * 依 A4 規則算每頁的切點。candidates＝可以切的位置（紙張內的 y，單位 css px）；total＝紙張總高。
 * tailSlack＝最後一頁允許多出的高度（紙張底部內距，約 42px；仍落在 10mm 下緣留白內）。
 * 回傳 [{ y0, y1, topMm }]。純函式，方便測試。
 */
function _qpdfPlanPages(total, candidates, pxPerMm, tailSlack) {
  const slack = Math.max(0, tailSlack || 0);   // 紙張底部的空白內距：剩下的只是空白就併進最後一頁，不要多出一張全白頁
  const pages = [];
  let y0 = 0;
  while (total - y0 > 0.5) {
    const topMm = pages.length ? QPDF_MARGIN_MM : 0;
    const limit = (QPDF_A4_MM.h - topMm - QPDF_MARGIN_MM) * pxPerMm;
    let y1;
    if (total - y0 <= limit + slack) {
      y1 = total;
    } else {
      y1 = y0 + limit;   // 找不到適合的切點就硬切
      const minAdvance = 30 * pxPerMm;   // 至少前進 30mm，避免產生幾乎空白的一頁
      const ok = candidates.filter(function (c) { return c > y0 + minAdvance && c <= y0 + limit; });
      if (ok.length) y1 = Math.max.apply(null, ok);
    }
    pages.push({ y0: y0, y1: y1, topMm: topMm });
    y0 = y1;
  }
  return pages;
}

/** 產生 jsPDF 文件。q＝報價單（含 perm／approval）、info＝issue-info。回傳 { doc, pageCount, stamped } */
async function _qpdfBuildDoc(q, info) {
  await _qpdfEnsureLibs();
  _qpvEnsureStyle();
  _qpvSealVer = Date.now();   // 避免瀏覽器快取到舊章
  const host = document.createElement('div');
  host.setAttribute('aria-hidden', 'true');
  host.style.cssText = 'position:fixed;left:-12000px;top:0;width:' + QPV_PAPER_WIDTH + 'px;background:#fff;pointer-events:none;';
  host.innerHTML = buildQuotePreviewHtml(q, info);
  document.body.appendChild(host);
  try {
    const paper = host.firstElementChild;
    paper.style.boxShadow = 'none';
    paper.style.margin = '0';
    // logo、報價章都是圖片，要等載入完才畫（章 404／損壞時預覽的 onerror 會把它藏起來，等待有上限）
    const imgs = Array.prototype.slice.call(paper.querySelectorAll('img'));
    await Promise.all(imgs.map(function (img) {
      if (img.complete) return null;
      return new Promise(function (res) { img.addEventListener('load', res); img.addEventListener('error', res); setTimeout(res, 8000); });
    }));
    const pr = paper.getBoundingClientRect();
    const total = Math.ceil(pr.height);
    const rel = function (el) { return el.getBoundingClientRect().bottom - pr.top; };

    // 蓋章：先記位置並隱藏，等紙張畫好再用 multiply 疊上去
    let seal = null;
    const sealImg = paper.querySelector('.qpv-seal');
    if (sealImg && sealImg.complete && sealImg.naturalWidth > 0 && sealImg.style.display !== 'none') {
      const r = sealImg.getBoundingClientRect();
      seal = { img: sealImg, x: r.left - pr.left, y: r.top - pr.top, w: r.width, h: r.height };
      sealImg.style.visibility = 'hidden';
    }

    // 可以切頁的位置：紙張的各個區塊、品項表的每一列、Remarks 的每一條（標題「Remarks ：」不單獨留在頁尾）
    const cand = [];
    Array.prototype.forEach.call(paper.children, function (el) { cand.push(rel(el)); });
    // 分組標題列之後不切（標題不能孤零零留在頁尾）；小計列的上一列之後也不切（小計不能孤立在下一頁頁首、離開它加總的品項）
    Array.prototype.forEach.call(paper.querySelectorAll('.qpv-items tbody tr:not(.qpv-grp)'), function (el) {
      const nx = el.nextElementSibling;
      if (nx && nx.classList.contains('qpv-sub')) return;
      cand.push(rel(el));
    });
    Array.prototype.forEach.call(paper.querySelectorAll('.qpv-remarks > div'), function (el, i) { if (i > 0) cand.push(rel(el)); });

    const scale = Math.max(0.5, Math.min(QPDF_SCALE, Math.sqrt(QPDF_MAX_PIXELS / (QPV_PAPER_WIDTH * total)), QPDF_MAX_SIDE / total));
    const canvas = await window.html2canvas(paper, {
      scale: scale, backgroundColor: '#ffffff', useCORS: true, logging: false,
      width: QPV_PAPER_WIDTH, windowWidth: QPV_PAPER_WIDTH,
    });
    const s = canvas.width / QPV_PAPER_WIDTH;   // 實際倍率
    if (seal) {
      const nw = seal.img.naturalWidth, nh = seal.img.naturalHeight;
      const k = Math.min(seal.w / nw, seal.h / nh);   // object-fit: contain
      const dw = nw * k, dh = nh * k;
      const c2 = canvas.getContext('2d');
      c2.save();
      c2.setTransform(1, 0, 0, 1, 0, 0);   // html2canvas 畫完後 context 仍帶著 scale/translate，不重設的話章會被畫到畫布外
      c2.globalCompositeOperation = 'multiply';
      c2.drawImage(seal.img, (seal.x + (seal.w - dw) / 2) * s, (seal.y + (seal.h - dh) / 2) * s, dw * s, dh * s);
      c2.restore();
    }

    const pages = _qpdfPlanPages(total, cand, QPV_PAPER_WIDTH / QPDF_A4_MM.w, parseFloat(getComputedStyle(paper).paddingBottom) || 0);
    const doc = new window.jspdf.jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait', compress: true });
    pages.forEach(function (pg, i) {
      if (i > 0) doc.addPage('a4', 'portrait');
      const sy = Math.round(pg.y0 * s);
      const sh = Math.max(1, Math.min(canvas.height - sy, Math.round((pg.y1 - pg.y0) * s)));
      const pc = document.createElement('canvas');
      pc.width = canvas.width;
      pc.height = sh;
      const ctx = pc.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, pc.width, pc.height);
      ctx.drawImage(canvas, 0, sy, canvas.width, sh, 0, 0, canvas.width, sh);
      doc.addImage(pc.toDataURL('image/png'), 'PNG', 0, pg.topMm, QPDF_A4_MM.w, sh / canvas.width * QPDF_A4_MM.w, undefined, 'FAST');
    });
    doc.setProperties({ title: String(q.quoteNo || 'quotation'), creator: 'ITTS-CRM' });
    return { doc: doc, pageCount: pages.length, stamped: !!seal, canvas: canvas, pages: pages };
  } finally {
    host.remove();
  }
}

/** 下載報價單 PDF（列表「⬇ PDF」與預覽視窗共用） */
async function exportQuotePdf(id, quoteNo) {
  const q = (await _qFetchQuote(id)) || _qFind(id);
  if (!q) return showToast('找不到此報價單');
  if (q.perm && q.perm.canSeePrice === false) return showToast('你沒有下載這張報價單的權限');
  // 與下載 Excel 相同：未核准／核准已失效要先警示
  if (!(await _qConfirmUnapproved(q))) return;
  showToast('正在產生 PDF…');
  try {
    const info = await _qpvLoadInfo(id);
    const built = await _qpdfBuildDoc(q, info);
    // 稽核紀錄：伺服器看不到瀏覽器端產生的檔案，所以由這裡回報。放在「產生成功之後、存檔之前」——
    // 函式庫載入失敗或畫圖失敗就不會留下「下載了 PDF」的紀錄；紀錄寫入失敗不擋下載（只在 console 留痕）。
    await fetch(API + '/quotations/' + encodeURIComponent(id) + '/export-log', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ format: 'pdf' }),
    }).then(function (r) { if (!r.ok) console.warn('[quote-pdf] export-log HTTP ' + r.status); }).catch(function (e) { console.warn('[quote-pdf] export-log 失敗', e); });
    const name = String((q.quoteNo || quoteNo || 'quotation') + (q.company ? '_' + q.company : '')).replace(/[\\/:*?"<>|\r\n]+/g, '_');
    built.doc.save(name + '.pdf');
    // 蓋章與否以「實際畫進 PDF 的結果」為準（章圖載入失敗或逾時時 PDF 沒有章，不能宣稱已蓋章）
    const approved = _qApprovalValid(q);
    showToast(approved && built.stamped ? '報價單 PDF 已下載（已蓋報價專用章）'
      : (approved
        ? '⚠ 此單已核准，但報價專用章沒有蓋上（尚未上傳或章圖載入失敗），下載的 PDF 沒有章，請聯絡管理部秘書'
        : '報價單 PDF 已下載（未蓋報價專用章）'));
  } catch (e) {
    showToast('PDF 產生失敗：' + ((e && e.message) || '請重試'));
  }
}
