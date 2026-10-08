'use strict';
/**
 * lib/mail/render.js — 事件 → 信件（主旨／HTML／純文字／meta）。純函式，無 I/O、不讀環境變數、不碰時鐘。
 *
 * 匯出與簽名：
 *   renderMail(ev, viewer, ctx)   → { subject, html, text, meta }
 *       ev      lib/mail/events.js 定義的事件（先經 validateEvent，不合法 throw MailRenderError）
 *       viewer  { username?, label, kind }   kind 必須在 visibility.KINDS 內
 *       ctx     { config, now? }             config.appBaseUrl 用來組連結；now 保留給之後的階段，本函式不使用
 *       meta    { type, quoteId, quoteNo, kind, hasAmount }   不含金額數字、email、專案／客戶名
 *   MailRenderError               extends Error，.code：BAD_EVENT / BAD_VIEWER / BAD_KIND / KIND_NOT_ALLOWED /
 *                                 BAD_CONFIG / BAD_AMOUNT / TOO_LARGE / INTERNAL（renderMail 只會 throw 這個類別）
 *   ALLOWED_KINDS                 事件類型 → 允許的收件人 kind（凍結）；不在表內的組合一律 throw KIND_NOT_ALLOWED
 *   subjectFor(ev, kind)          → string     主旨（不含金額／毛利率／客戶名／業務名；專案名稱只在 visibility.project 為 true 的 kind 才放，kind 缺或未知＝不放）
 *   buildQuoteUrl(config, quoteId, {cost}) → string   `${appBaseUrl}/q/${encodeURIComponent(id)}`，cost 加 ?cost=1
 *   formatNtd(cents)              → 'NT$ 1,234,567' / 'NT$ 1,234.50'   只用字串切割與整數，不做浮點除法
 *   formatTaipei(iso)             → 'YYYY-MM-DD HH:mm'（Asia/Taipei＝UTC+8，無日光節約，不依賴 Intl）
 *   buildModel(ev, viewer, ctx)   → templates.js 的 model（測試與預覽工具用）
 *   RESULT_INFO / TIER_TONE       結果文字與色塊／核決層級色塊的對照；tierTag(level, tierLabel) 產生色塊內的文字標籤
 *
 * 可見性：一律查 visibility.visibilityFor(viewer.kind)，在「組 model」時就把不該出現的資料丟掉
 * （templates.js 拿到的 model 根本沒有那些字串，不可能誤印）。本檔沒有任何「某角色看得到什麼」的判斷，
 * 唯一的角色表是 ALLOWED_KINDS（這個事件該不該寄給這種收件人），不是可見性。
 *
 * 事件 × 收件人 kind 矩陣（◎＝可渲染；其餘 throw KIND_NOT_ALLOWED）：
 *                  mgr1 gm chairman secretary boardProxy consultant owner
 *   E1 送簽         ◎   ◎    ◎        ◎          ◎
 *   E2 請填成本                                              ◎
 *   E3 下一關       ◎   ◎    ◎        ◎          ◎
 *   E4 結果→業務                                                     ◎
 *   E5 成本完成→業務                                                  ◎
 *   E6 撤回／作廢   ◎   ◎    ◎        ◎          ◎                    ◎   （業務本人也會收到自己單據的撤回／作廢確認）
 *
 * 決策條（只出現在 E1／E3 這種「請您簽核」的信，位於最上方）：報價金額（標示折扣後未稅）、毛利率（色塊＋文字標籤）、核決層級。
 *   E6（請勿簽核）、E4／E5（給業務）不放決策條：不需要、也不該再多放一次金額與毛利率（最小資料原則）。
 *   色塊只看 numbers.tierLevel：1＝綠、2＝琥珀、3＝紅、其他（含 null）＝灰；文字標籤一定同時存在，
 *   且由 numbers.tierLabel 組成（一級主管可核／需總經理核准／需董事長核准／需董事會決議），不由 tierLevel 寫死。
 *   numbers 為 null，或該收件人的 visibility 對應欄位為 false → 該格不輸出；三格都沒有 → 整段決策條不輸出。
 *   毛利率顯示的是 numbers.marginText（沿用簽核系統的截斷規則，不重算）；marginPct 只用來判斷是否為負。
 *
 * 其他注意：
 *   - 專案名稱、客戶、業務、關卡、原因、品項等使用者輸入，進內文前一律經 safeText（單行化、去不可見字元、截斷），
 *     進 HTML 前由 templates.js 的 escHtml 跳脫；進主旨前經 headerSafe。
 *   - 品項只取 desc／qty／unit 三個欄位（其他鍵，例如 price，不會被讀取）。
 *   - E1／E2／E6 的 actor 多半就是業務本人，所以「操作人」那一列受 visibility.owner 控管；actor 文字若等於
 *     業務名而該收件人看不到業務名，也不顯示。
 *   - 連結 href 只有 buildQuoteUrl 產生的那一個字串（appBaseUrl 開頭、quoteId 通過格式檢查）。
 */

const { headerSafe, safeText } = require('./safety');
const { validateEvent, displayMargin, isValidQuoteId, EVENT_TYPES, LIMITS } = require('./events');
const { visibilityFor, isKnownKind } = require('./visibility');
const { layoutHtml, layoutText } = require('./templates');

const MAX_MAIL_BYTES = 100 * 1024;
const MAX_ITEM_ROWS = 30;

class MailRenderError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'MailRenderError';
    this.code = code;
  }
}

const APPROVERS = Object.freeze(['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy']);
const ALLOWED_KINDS = Object.freeze({
  E1_SUBMIT: Object.freeze(APPROVERS.slice()),
  E2_COST_REQUEST: Object.freeze(['consultant']),
  E3_NEXT_STEP: Object.freeze(APPROVERS.slice()),
  E4_RESULT: Object.freeze(['owner']),
  E5_COST_DONE: Object.freeze(['owner']),
  E6_WITHDRAWN: Object.freeze(APPROVERS.concat(['owner'])),
});

// 結果通知的文字與色塊
const RESULT_INFO = Object.freeze({
  approved: Object.freeze({ label: '本關已核准', tone: 'green' }),
  final_approved: Object.freeze({ label: '已完成核准', tone: 'green' }),
  rejected: Object.freeze({ label: '已駁回', tone: 'red' }),
  returned: Object.freeze({ label: '已退回修改', tone: 'amber' }),
  withdrawn: Object.freeze({ label: '已撤回', tone: 'amber' }),
  voided: Object.freeze({ label: '核准已作廢', tone: 'red' }),
});

// 核決層級 → 色塊（以 tierLevel 為準；其他含 null 一律灰）。文字標籤另由 tierLabel 產生（見 tierTag），
// 這樣董事會關（level 3＋label「董事會」）不會被標成「需董事長核准」。
const TIER_TONE = Object.freeze({ 1: 'green', 2: 'amber', 3: 'red' });

/** 色塊內的文字標籤，一律由系統給的 tierLabel 組成：一級主管可核／需總經理核准／需董事長核准／需董事會決議 */
function tierTag(level, tierLabel) {
  if (!tierLabel) return '核決層級未定';
  if (level === 1) return tierLabel + '可核';
  return '需' + tierLabel + (tierLabel === '董事會' ? '決議' : '核准');
}

const SUBJECT_PREFIX = Object.freeze({
  E1_SUBMIT: '【簽核通知】',
  E2_COST_REQUEST: '【請填寫成本】',
  E3_NEXT_STEP: '【簽核通知】',
  E4_RESULT: '【簽核結果】',
  E5_COST_DONE: '【成本已填寫】',
  E6_WITHDRAWN: '【簽核撤回】',
});
const SUBJECT_SUFFIX = Object.freeze({ E6_WITHDRAWN: ' 請勿簽核' });
const SUBJECT_PROJECT_MAX = 40;
const SUBJECT_MAX = 120;

const BRAND = 'ITTS-CRM';
const CONFIDENTIAL = '機密，請勿轉寄';
const FOOTER_LINES = Object.freeze([
  '機密，請勿轉寄；本信由 ITTS-CRM 自動發送，請勿直接回覆。',
]);
const NOTE_SNAPSHOT = '本信為事件發生當下的通知，請以系統內的即時狀態為準。';

// ── 格式化 ──────────────────────────────────────────────────────────────────
/**
 * 金額（分，非負安全整數）→ 'NT$ 1,234,567'（分為 0 不顯示小數；非 0 顯示 2 位）。
 * 只用「十進位字串切割」：最後兩位數字是分，其餘是元。沒有任何浮點除法，所以不會有 0.1+0.2 類的誤差。
 */
function formatNtd(cents) {
  if (typeof cents !== 'number' || !Number.isSafeInteger(cents) || cents < 0) {
    throw new MailRenderError('BAD_AMOUNT', '金額必須是非負的整數（單位：分）');
  }
  const s = String(cents);                                  // 安全整數最多 16 位，不會出現科學記號
  const padded = s.length < 3 ? '000'.slice(s.length) + s : s;
  const yuan = padded.slice(0, padded.length - 2);
  const fen = padded.slice(padded.length - 2);
  const grouped = yuan.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return 'NT$ ' + grouped + (fen === '00' ? '' : '.' + fen);
}

function pad2(n) { return n < 10 ? '0' + n : String(n); }
function pad4(n) { return String(n).length >= 4 ? String(n) : '0000'.slice(String(n).length) + n; }

/** ISO 時間字串 → 台北時間 'YYYY-MM-DD HH:mm'。無法解析回 ''。 */
function formatTaipei(iso) {
  const ms = Date.parse(iso);
  if (!isFinite(ms)) return '';
  const d = new Date(ms + 8 * 3600 * 1000);
  return pad4(d.getUTCFullYear()) + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate()) + ' ' + pad2(d.getUTCHours()) + ':' + pad2(d.getUTCMinutes());
}

function formatMargin(marginText) {
  return displayMargin(marginText);                     // 補一個 %；整數超過 6 位的極端值顯示成 <-999999%（定義在 events.js）
}

function formatQty(q) {
  if (typeof q === 'number') {
    if (!(q >= 0) || !(q < 1e15)) return '-';
    if (Number.isInteger(q)) return String(q);
    return q.toFixed(4).replace(/0+$/, '').replace(/\.$/, '');
  }
  return safeText(q, 20);
}

// ── 連結 ────────────────────────────────────────────────────────────────────
const RE_BASE_URL = /^https?:\/\/[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?(?::[0-9]{1,5})?$/;

function checkedBase(config) {
  const b = config && config.appBaseUrl;
  if (typeof b !== 'string' || !RE_BASE_URL.test(b)) throw new MailRenderError('BAD_CONFIG', 'appBaseUrl 不合法，無法產生連結');
  const m = /^http:\/\/(localhost|127\.0\.0\.1)(:[0-9]{1,5})?$/.test(b);
  if (b.indexOf('http://') === 0 && !m) throw new MailRenderError('BAD_CONFIG', 'appBaseUrl 必須是 https（localhost 除外）');
  return b;
}

/** `${appBaseUrl}/q/${encodeURIComponent(quoteId)}`；cost 為真加上 ?cost=1。quoteId 格式不合法 throw。 */
function buildQuoteUrl(config, quoteId, opts) {
  const base = checkedBase(config);
  if (!isValidQuoteId(quoteId)) throw new MailRenderError('BAD_EVENT', '單據 id 格式不合法，無法產生連結');
  return base + '/q/' + encodeURIComponent(quoteId) + (opts && opts.cost ? '?cost=1' : '');
}

// ── 主旨 ────────────────────────────────────────────────────────────────────
/**
 * `【類別】單號 專案名稱`（專案名稱截 40 字、整行 headerSafe 到 120）。不含金額／毛利率／客戶名／業務名。
 * 專案名稱只在收件人 kind 的 visibility.project 為 true 時放：秘書與董事會代核人（以及缺／未知 kind）的主旨只有`【類別】單號`。
 */
function subjectLine(ev, showProject) {
  const prefix = SUBJECT_PREFIX[ev.type];
  if (!prefix) throw new MailRenderError('BAD_EVENT', '事件類型不在列舉內');
  const no = headerSafe(ev.quoteNo, 40);
  const proj = showProject ? safeText(ev.projectName, SUBJECT_PROJECT_MAX) : '';
  return headerSafe(prefix + [no, proj].filter(Boolean).join(' ') + (SUBJECT_SUFFIX[ev.type] || ''), SUBJECT_MAX);
}
function subjectFor(ev, kind) {
  return subjectLine(ev, visibilityFor(kind).project === true);
}

// ── model ───────────────────────────────────────────────────────────────────
function buildDecision(n, vis) {
  if (!n) return null;
  const cells = [];
  const tierLabel = safeText(n.tierLabel, 20);
  if (vis.amount) {
    cells.push({ kind: 'amount', caption: '報價金額', value: formatNtd(n.revenueCents), tag: '折扣後未稅', tone: null });
  }
  if (vis.margin) {
    const lv = n.tierLevel === undefined ? null : n.tierLevel;
    let tone = null;
    let tag = '';
    if (vis.tier) {
      tone = Object.prototype.hasOwnProperty.call(TIER_TONE, lv) ? TIER_TONE[lv] : 'grey';
      tag = tierTag(lv, tierLabel);
    }
    // 毛利為負（虧損單）另外成一行警示，不接在色塊標籤後面（窄格子裡會折成孤字）
    const loss = n.gpCents < 0 || n.marginPct < 0;
    cells.push({ kind: 'margin', caption: '毛利率', value: formatMargin(n.marginText), tag, warn: loss ? '毛利為負' : '', tone });
  }
  if (vis.tier && tierLabel) {
    cells.push({ kind: 'tier', caption: '核決層級', value: tierLabel, tag: '', tone: null });
  }
  return cells.length ? { cells } : null;
}

function leadFor(ev, stepLabel, isBoard) {
  const where = stepLabel ? '「' + stepLabel + '」' : '';
  switch (ev.type) {
    case 'E1_SUBMIT':
      return isBoard
        ? ['有一張報價單已送簽，目前進入' + where + '關，請依董事會決議至系統登錄。']
        : ['有一張報價單已送簽，目前輪到您處理' + where + '這一關。', '請登入系統檢視完整內容後，再進行核准或駁回。'];
    case 'E3_NEXT_STEP':
      return isBoard
        ? ['報價單上一關已核准，現在進入' + where + '關，請依董事會決議至系統登錄。']
        : ['報價單上一關已核准，現在輪到您處理' + where + '這一關。', '請登入系統檢視完整內容後，再進行核准或駁回。'];
    case 'E2_COST_REQUEST':
      return ['您被指定為這張報價單的支援顧問，請登入系統填寫各品項的成本。'];
    case 'E4_RESULT': {
      const k = ev.result && ev.result.kind;
      if (k === 'approved') return ['您的報價單已通過' + where + '這一關，將送往下一關。'];
      if (k === 'final_approved') return ['您的報價單已完成所有簽核，核准通過。'];
      if (k === 'rejected') return ['您的報價單已被駁回。'];
      return ['您的報價單已被退回，請依說明修改後重新送簽。'];
    }
    case 'E5_COST_DONE':
      return ['顧問已完成這張報價單的成本填寫，您可以回到系統繼續後續作業。'];
    case 'E6_WITHDRAWN':
      return ev.result && ev.result.kind === 'voided'
        ? ['這張報價單原先的核准已作廢（核准後內容已被修改），請勿依原核准內容辦理。']
        : ['這張報價單已撤回，簽核流程已終止；若您先前收過簽核通知，請勿再簽核。'];
    default:
      return [];
  }
}

const HEADLINE = Object.freeze({
  E1_SUBMIT: '報價單待簽核',
  E2_COST_REQUEST: '請填寫報價成本',
  E3_NEXT_STEP: '報價單待簽核',
  E4_RESULT: '報價單簽核結果',
  E5_COST_DONE: '成本已填寫完成',
  E6_WITHDRAWN: '簽核已撤回，請勿簽核',
});
const BUTTON_LABEL = Object.freeze({
  E1_SUBMIT: '前往系統簽核',
  E2_COST_REQUEST: '前往填寫成本',
  E3_NEXT_STEP: '前往系統簽核',
  E4_RESULT: '開啟報價單',
  E5_COST_DONE: '開啟報價單',
  E6_WITHDRAWN: '查看報價單',
});
const STEP_ROW_LABEL = Object.freeze({ E1_SUBMIT: '目前關卡', E3_NEXT_STEP: '目前關卡', E4_RESULT: '關卡', E6_WITHDRAWN: '原關卡' });
const SHOWS_DECISION = Object.freeze({ E1_SUBMIT: true, E3_NEXT_STEP: true });
// 這幾種事件的 actor 通常就是業務本人，所以「操作人」受 visibility.owner 控管
const ACTOR_IS_OWNER = Object.freeze({ E1_SUBMIT: true, E2_COST_REQUEST: true, E6_WITHDRAWN: true });

function buildModel(ev, viewer, ctx) {
  const vis = visibilityFor(viewer.kind);
  const type = ev.type;
  const quoteNo = safeText(ev.quoteNo, 40);
  const project = vis.project ? safeText(ev.projectName, 120) : '';      // 看不到專案名稱的 kind：在這裡就丟掉，後面（主旨、preheader、rows）都不可能誤印
  const label = safeText(viewer.label, 60);
  const stepLabel = ev.step && ev.step.label ? safeText(ev.step.label, 40) : '';
  const isBoard = !!(ev.step && ev.step.level === 'board');
  const url = buildQuoteUrl(ctx.config, ev.quoteId, { cost: type === 'E2_COST_REQUEST' });

  const company = vis.customer ? safeText(ev.company, 100) : '';
  const owner = vis.owner ? safeText(ev.ownerLabel, 60) : '';
  const ownerRaw = safeText(ev.ownerLabel, 60);

  const rows = [{ label: '報價單號', value: quoteNo }];
  if (project) rows.push({ label: '專案名稱', value: project });
  if (company) rows.push({ label: '客戶', value: company });
  if (owner) rows.push({ label: '業務', value: owner });
  if (stepLabel && STEP_ROW_LABEL[type]) rows.push({ label: STEP_ROW_LABEL[type], value: stepLabel });
  if (ev.actor && ev.actor.label) {
    const actor = safeText(ev.actor.label, 60);
    // 看不到業務名的人：E1／E2／E6 的操作人（多半就是業務）一律不顯示；操作人與業務同名時，
    // 看得到業務名的人已經有「業務」那一列，不重複列；看不到的人則不能靠這一列洩漏。
    const hide = (ACTOR_IS_OWNER[type] && !vis.owner) || (ownerRaw !== '' && actor === ownerRaw);
    if (actor && !hide) rows.push({ label: '操作人', value: actor });
  }
  const when = formatTaipei(ev.at);
  if (when) rows.push({ label: '時間', value: when });

  let result = null;
  if ((type === 'E4_RESULT' || type === 'E6_WITHDRAWN') && ev.result) {
    const info = RESULT_INFO[ev.result.kind];
    if (info) {
      const reason = safeText(ev.result.reason, 200);
      result = {
        caption: type === 'E4_RESULT' ? '簽核結果' : '狀態',
        label: info.label,
        tone: info.tone,
        reasonLabel: '原因',
        reason,
      };
    }
  }

  let items = null;
  if (vis.items && Array.isArray(ev.items) && ev.items.length > 0) {
    const shown = ev.items.slice(0, MAX_ITEM_ROWS).map((it) => [
      safeText(it.desc, 120),
      it.qty === undefined || it.qty === null ? '' : formatQty(it.qty),
      safeText(it.unit, 20),
    ]);
    const extra = ev.items.length - shown.length;
    items = { head: ['品項說明', '數量', '單位'], rows: shown, moreText: extra > 0 ? '…另 ' + extra + ' 項' : '' };
  }

  // 決策條只出現在「請您簽核」的信（E1／E3）。E6 是「請勿簽核」、E4／E5 是給業務的結果，不需要也不該再放金額與毛利率。
  const decision = SHOWS_DECISION[type] ? buildDecision(ev.numbers, vis) : null;

  const notes = [];
  if (type === 'E1_SUBMIT' || type === 'E3_NEXT_STEP') notes.push('核准或駁回需登入系統操作，本信不提供信內核准。');
  if (type === 'E2_COST_REQUEST') notes.push('請登入系統後，於報價單內填寫成本。');
  notes.push(NOTE_SNAPSHOT);

  const pre = {
    E1_SUBMIT: '待您簽核：', E3_NEXT_STEP: '待您簽核：', E2_COST_REQUEST: '請填寫成本：',
    E4_RESULT: '簽核結果：', E5_COST_DONE: '成本已填寫：', E6_WITHDRAWN: '請勿簽核：',
  }[type];

  return {
    title: subjectLine(ev, vis.project === true),
    preheader: safeText(pre + [quoteNo, project].filter(Boolean).join(' '), 90),
    brand: BRAND,
    confidential: CONFIDENTIAL,
    headline: HEADLINE[type],
    headlineTone: type === 'E6_WITHDRAWN' ? 'amber' : 'blue',
    result,
    decision,
    greeting: label ? label + '，您好：' : '您好：',
    lead: leadFor(ev, stepLabel, isBoard),
    rows,
    items,
    button: { label: BUTTON_LABEL[type], url },
    notes,
    footer: FOOTER_LINES.slice(),
  };
}

// ── 快照 ────────────────────────────────────────────────────────────────────
// 驗證與渲染都只讀「快照」：每個欄位只讀一次、只複製已知欄位（其餘鍵，例如品項的 price，在這裡就被丟掉）。
// 這樣即使輸入物件帶有 getter／Proxy（驗證時回合法值、之後改回惡意值），也不可能繞過驗證。
function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }

function snapStep(s) {
  if (!isObj(s)) return s;
  return { level: s.level, label: s.label };
}
function snapNumbers(n) {
  if (!isObj(n)) return n;
  return { revenueCents: n.revenueCents, gpCents: n.gpCents, marginText: n.marginText, marginPct: n.marginPct, tierLevel: n.tierLevel, tierLabel: n.tierLabel };
}
function snapResult(r) {
  if (!isObj(r)) return r;
  return { kind: r.kind, reason: r.reason };
}
function snapItems(items) {
  if (!Array.isArray(items)) return items;
  const out = [];
  const max = Math.min(items.length, LIMITS.items + 1);        // 多讀一筆，讓 validateEvent 照常回報「最多 200 筆」
  for (let i = 0; i < max; i++) {
    const it = items[i];
    out.push(isObj(it) ? { desc: it.desc, qty: it.qty, unit: it.unit } : it);
  }
  return out;
}
function snapshotEvent(ev) {
  if (!isObj(ev)) return ev;
  try {
    return {
      type: ev.type, quoteId: ev.quoteId, quoteNo: ev.quoteNo, projectName: ev.projectName, company: ev.company,
      ownerLabel: ev.ownerLabel, step: snapStep(ev.step), numbers: snapNumbers(ev.numbers), result: snapResult(ev.result),
      items: snapItems(ev.items), actor: isObj(ev.actor) ? { label: ev.actor.label } : ev.actor, at: ev.at, stepKey: ev.stepKey,
    };
  } catch (e) {
    throw new MailRenderError('BAD_EVENT', '事件不合法：事件內容無法讀取');
  }
}
function snapshotViewer(viewer) {
  if (!isObj(viewer)) return viewer;
  try {
    return { username: viewer.username, label: viewer.label, kind: viewer.kind };
  } catch (e) {
    throw new MailRenderError('BAD_VIEWER', 'viewer 內容無法讀取');
  }
}

// ── 入口 ────────────────────────────────────────────────────────────────────
function checkViewer(viewer) {
  if (viewer === null || typeof viewer !== 'object' || Array.isArray(viewer)) throw new MailRenderError('BAD_VIEWER', 'viewer 必須是物件');
  if (typeof viewer.kind !== 'string' || !isKnownKind(viewer.kind)) throw new MailRenderError('BAD_KIND', '收件人類型 kind 不在允許清單內');
  if (viewer.label !== undefined && viewer.label !== null && typeof viewer.label !== 'string') throw new MailRenderError('BAD_VIEWER', 'viewer.label 必須是文字');
  if (viewer.username !== undefined && viewer.username !== null && typeof viewer.username !== 'string') throw new MailRenderError('BAD_VIEWER', 'viewer.username 必須是文字');
}

/**
 * 事件 → 信件。只會 throw MailRenderError（dispatcher 會 catch 並記為失敗）。
 * 檢查順序：事件 → 收件人 → 事件×kind 矩陣 → 設定 → 組版 → 大小上限。
 */
function renderMail(rawEv, rawViewer, ctx) {
  try {
    const ev = snapshotEvent(rawEv);
    const v = validateEvent(ev);
    if (!v.ok) throw new MailRenderError('BAD_EVENT', '事件不合法：' + v.error + (v.field ? '（' + v.field + '）' : ''));
    const viewer = snapshotViewer(rawViewer);
    checkViewer(viewer);
    const allowed = ALLOWED_KINDS[ev.type];
    if (!allowed || allowed.indexOf(viewer.kind) < 0) {
      throw new MailRenderError('KIND_NOT_ALLOWED', '事件 ' + ev.type + ' 不寄給 ' + viewer.kind + ' 類型的收件人');
    }
    if (ctx === null || typeof ctx !== 'object') throw new MailRenderError('BAD_CONFIG', 'ctx.config 必填');
    const model = buildModel(ev, viewer, ctx);
    const html = layoutHtml(model);
    const text = layoutText(model);
    if (Buffer.byteLength(html, 'utf8') > MAX_MAIL_BYTES || Buffer.byteLength(text, 'utf8') > MAX_MAIL_BYTES) {
      throw new MailRenderError('TOO_LARGE', '信件超過大小上限（100KB）');
    }
    return {
      subject: model.title,
      html,
      text,
      meta: {
        type: ev.type,
        quoteId: ev.quoteId,
        quoteNo: ev.quoteNo,
        kind: viewer.kind,
        hasAmount: !!(model.decision && model.decision.cells.some((c) => c.kind === 'amount')),
      },
    };
  } catch (e) {
    if (e instanceof MailRenderError) throw e;
    throw new MailRenderError('INTERNAL', '信件渲染發生未預期的錯誤（' + (e && e.name ? e.name : 'Error') + '）');
  }
}

module.exports = {
  renderMail,
  MailRenderError,
  ALLOWED_KINDS,
  RESULT_INFO,
  TIER_TONE,
  tierTag,
  EVENT_TYPES,
  subjectFor,
  buildQuoteUrl,
  buildModel,
  formatNtd,
  formatTaipei,
};
