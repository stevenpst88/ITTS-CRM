#!/usr/bin/env node
/**
 * 報價單「簽核進度卡片」(流程設計器 M0) 單元檢查。用法：node scripts/check-quote-approval-cards.js
 * 動 _client/quote-flow-cards.js、或 quote-approval.js 裡接卡片的那幾行之後必跑。
 *   1) 介面與靜態紀律：只讀 approval.steps／cur／state／valid／board／history，不碰 derived（毛利率、說明文字、門檻、金額）
 *   2) cardsOf 狀態組合：草稿／撤回／作廢／簽核中（一關、三關、五關）／已退回／已核准全綠／核准已失效
 *   3) 代核（board）標記、代核人、決議日期文號；多位候選人；顧問視角（approval 為 null）
 *   4) 不洩漏毛利：業務視角序列化（含 derived.marginText／reasons）→ view-model 與 HTML 的 JSON 掃描不含毛利相關字串；有沒有 derived 輸出完全相同
 *   5) HTML：escape、aria（role／aria-current／aria-expanded／aria-controls）、狀態同時有文字＋符號
 *   6) 候選人展開（bind 事件委派）
 *   7) ?cards=0 後備
 *   8) CSS：每個狀態 × 淺色／深色的前景／背景對比 ≥ 4.5:1、reduced-motion、容器查詢直排／橫排
 *   9) 接線：index.html 先載入、quote-approval.js 有接線且保留舊條列式、server.js 版本參數、sw.js 版號
 * 變異測試：QFC_SRC=<被破壞的副本> node scripts/check-quote-approval-cards.js 必須變紅（見回報的變異清單）。
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名）。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const ROOT = path.join(__dirname, '..');
const SRC = process.env.QFC_SRC || path.join(ROOT, '_client/quote-flow-cards.js');
const src = fs.readFileSync(SRC, 'utf8');
const QFC = require(SRC);
const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const fmt = (v) => 'T:' + v;   // 注入時間格式，避免時區讓測試不穩

// ── 夾具：形狀與 lib/quoteRoutes.js serializeApproval 的輸出一致（業務視角：仍帶 derived.marginText／reasons，卡片必須無視它們）──
const DERIVED_LEAK = {
  rowKey: 'consult_software', rowLabel: '顧問＋軟體', level: 2, board: false,
  marginText: '毛利率 12.5%（低於 15%）', tiers: ['mgr1', 'gm'],
  reasons: ['毛利率 12.5% 低於一級主管門檻 20%，需總經理核准', '成本合計 8,750,000'],
  thresholds: { l1: 20, l2: 10, amountThresholds: { chairman: 10000000 } },
  revenueCents: 1000000000, costCents: 875000000, gpCents: 125000000,
};
const step = (tier, status, extra) => Object.assign({ tier, label: ({ mgr2: '二級主管複核', mgr1: '一級主管', gm: '總經理', chairman: '董事長', board: '董事會決議（秘書代核）' })[tier], status, assigneeName: '', byName: '', at: null, comment: '' }, extra || {});
const ap = (state, steps, extra) => Object.assign({ state, valid: state === 'approved', submittedAt: '2026-10-01T01:00:00Z', submittedByName: '業務甲', steps, cur: 0, board: null, history: [], derived: DERIVED_LEAK }, extra || {});
const S_MGR1 = (st, x) => step('mgr1', st, Object.assign({ assigneeName: '主管乙' }, x));
const S_GM = (st, x) => step('gm', st, Object.assign({ assigneeName: '總座丙、總座丁' }, x));
const S_CH = (st, x) => step('chairman', st, Object.assign({ assigneeName: '董座戊' }, x));
const S_BD = (st, x) => step('board', st, Object.assign({ assigneeName: '管理部秘書' }, x));
const C = (a) => QFC.cardsOf(a, { fmtTime: fmt });
const sts = (vm) => vm.cards.map((c) => c.status).join(',');

// ═════════════ 1) 介面與靜態紀律 ═════════════
t('1a. 匯出 cardsOf／render／bind／enabled／CSS', ['cardsOf', 'render', 'bind', 'enabled'].every((k) => typeof QFC[k] === 'function') && typeof QFC.CSS === 'string' && QFC.CSS.length > 500);
{
  const code = src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  const banned = ['derived', 'marginText', 'marginPct', 'reasons', 'thresholds', 'revenueCents', 'costCents', 'gpCents', 'rosterSnapshot'];
  const hit = banned.filter((w) => new RegExp('\\b' + w + '\\b').test(code));
  t('1b. 程式碼不引用 derived／marginText／reasons／thresholds／金額成本毛利欄位（不洩漏的結構性保證）', hit.length === 0, hit.join(','));
  t('1c. 檔案紀律：沒有本機路徑／帳密', !/[A-Za-z]:[\\/]Users|password|secret|steven/i.test(src));
}
t('1d. 每個引擎關卡 key 都有職稱與節點名（與 lib/quoteApproval.js TIERS 同步）', Object.keys(QA.TIERS).every((k) => QFC.TITLE[k] && QFC.NODE[k]), Object.keys(QA.TIERS).join(','));

// ═════════════ 2) 狀態組合 ═════════════
{
  let v = C(null);
  t('2a. 顧問視角（approval 為 null）：不報錯、沒有卡片、render 為空字串', v.cards.length === 0 && v.state === 'none' && QFC.render(v) === '');
  v = C(ap('none', [], { valid: false, derived: undefined }));
  t('2b. 草稿尚未送簽（state none、無歷程）：state=draft、無卡片、render 空', v.state === 'draft' && v.cards.length === 0 && QFC.render(v) === '');
  v = C(ap('none', [], { history: [{ action: 'SUBMIT' }, { action: 'WITHDRAW' }] }));
  t('2c. 已撤回：banner（wait）、無卡片，render 只出 banner', v.state === 'withdrawn' && v.banner && v.banner.kind === 'wait' && /已撤回/.test(v.banner.text) && /qfc-banner/.test(QFC.render(v)) && !/qfc-flow/.test(QFC.render(v)));
  v = C(ap('none', [], { history: [{ action: 'SUBMIT' }, { action: 'APPROVE' }, { action: 'INVALIDATE' }] }));
  t('2d. 核准作廢（INVALIDATE）：banner（void）', v.state === 'invalidated' && v.banner.kind === 'void' && /作廢/.test(v.banner.text));
  v = C(ap('none', [], { history: [{ action: 'INVALIDATE' }, { action: 'SUBMIT' }, { action: 'WITHDRAW' }, { action: 'COST_DONE' }] }));
  t('2e. 取最後一個簽核動作（非簽核事件如 COST_DONE 略過）：WITHDRAW 勝過更早的 INVALIDATE', v.state === 'withdrawn');

  // 簽核中：只有一關
  v = C(ap('pending', [S_MGR1('pending')]));
  t('2f. 僅一關、簽核中：1 張卡 cur；標籤「目前待 主管乙｜一級主管」；headline「目前待 主管乙 審核」', v.cards.length === 1 && sts(v) === 'cur' && v.tag.text === '目前待 主管乙｜一級主管' && v.headline === '目前待 主管乙 審核', J([v.tag, v.headline]));
  // 簽核中：三關，第 2 關
  v = C(ap('pending', [S_MGR1('approved', { byName: '主管乙', at: '2026-10-02T02:00:00Z', comment: '同意' }), S_GM('pending'), S_CH('waiting', {})], { cur: 1 }));
  t('2g. 三關到第 2 關：done,cur,wait；連線 done,wait,wait', sts(v) === 'done,cur,wait' && v.cards.map((c) => c.link).join(',') === 'done,wait,wait');
  t('2h. 已完成卡：簽核人、時間、意見摘要；目前卡與後續卡不顯示時間意見', v.cards[0].decidedBy === '主管乙' && v.cards[0].at === 'T:2026-10-02T02:00:00Z' && v.cards[0].comment === '同意' && v.cards[1].at === '' && v.cards[2].decidedBy === '');
  t('2i. 目前關卡是名冊型（總經理 2 位）：headline 用關名＋候選人數；標籤「目前待 2 位候選人｜總經理」', v.headline === '目前待 總經理（2 位候選人） 審核' && v.tag.text === '目前待 2 位候選人｜總經理', J([v.headline, v.tag]));
  // 五關
  const five = [step('mgr2', 'approved', { assigneeName: '主管己', byName: '主管己', at: '2026-10-02T02:00:00Z' }), S_MGR1('approved', { byName: '主管乙', at: '2026-10-02T03:00:00Z' }), S_GM('approved', { byName: '總座丙', at: '2026-10-02T04:00:00Z' }), S_CH('pending'), S_BD('waiting')];
  v = C(ap('pending', five, { cur: 3 }));
  t('2j. 五關到第 4 關：done,done,done,cur,wait；n 值與 STEP 序號 1..5、total=5', sts(v) === 'done,done,done,cur,wait' && v.cards.map((c) => c.n).join('') === '12345' && v.cards.every((c) => c.total === 5));
  t('2k. 五關 render：ol.n5、五個 li、恰好一個 aria-current="step"', (() => { const h = QFC.render(v); return /class="qfc-flow n5"/.test(h) && (h.match(/<li class="qfc-step/g) || []).length === 5 && (h.match(/aria-current="step"/g) || []).length === 1; })());
  v = C(ap('pending', [step('mgr2', 'pending', { assigneeName: '主管己' }), S_MGR1('waiting'), S_GM('waiting'), S_CH('waiting'), S_BD('waiting'), step('gm', 'waiting', {})], { cur: 0 }));
  t('2l. 超過五關（防呆）：class 仍是 n5，不炸', /class="qfc-flow n5"/.test(QFC.render(v)) && v.cards.length === 6);
  // cur 越界
  v = C(ap('pending', [S_MGR1('pending')], { cur: 9 }));
  t('2m. cur 越界／缺漏：夾在範圍內不報錯', v.cards.length === 1 && v.cards[0].status === 'cur');
  v = C(ap('pending', [S_MGR1('pending'), S_GM('waiting')], { cur: undefined }));
  t('2n. cur 缺漏視為 0', sts(v) === 'cur,wait');

  // 被退回
  v = C(ap('returned', [S_MGR1('approved', { byName: '主管乙', at: '2026-10-02T02:00:00Z' }), S_GM('returned', { byName: '總座丙', at: '2026-10-03T02:00:00Z', comment: '金額需再議' }), S_CH('waiting'), S_BD('waiting')], { valid: false, cur: 1 }));
  t('2o. 被退回：done,ret,skip,skip；連線 done,ret,wait,wait', sts(v) === 'done,ret,skip,skip' && v.cards.map((c) => c.link).join(',') === 'done,ret,wait,wait');
  t('2p. 退回卡：退回人、時間、意見；headline 與標籤指到該關', v.cards[1].decidedBy === '總座丙' && v.cards[1].comment === '金額需再議' && /總座丙/.test(v.headline) && /總經理/.test(v.headline) && v.tag.kind === 'ret' && /已退回｜總經理/.test(v.tag.text), J([v.headline, v.tag]));
  t('2q. 退回的 render：後續關卡「未進行」且有刪除線樣式 class（st-skip）', (() => { const h = QFC.render(v); return /st-skip/.test(h) && /未進行/.test(h) && /st-ret/.test(h) && /已退回/.test(h); })());
  v = C(ap('returned', [S_MGR1('pending')], { valid: false }));
  t('2r. 退回但資料殘留 pending：視為 ret（不顯示成目前關卡）', sts(v) === 'ret');

  // 已核准全綠
  v = C(ap('approved', [S_MGR1('approved', { byName: '主管乙', at: '2026-10-02T02:00:00Z' }), S_GM('approved', { byName: '總座丙', at: '2026-10-02T03:00:00Z' })], { valid: true }));
  t('2s. 已核准且有效：全部 done、連線全 done、headline「簽核已全部完成」、標籤已核准', sts(v) === 'done,done' && v.headline === '簽核已全部完成' && v.tag.kind === 'done' && !v.cards.some((c) => c.current) && !/aria-current/.test(QFC.render(v)));
  // 核准失效
  v = C(ap('approved', [S_MGR1('approved', { byName: '主管乙', at: '2026-10-02T02:00:00Z' }), S_GM('approved', { byName: '總座丙', at: '2026-10-02T03:00:00Z' })], { valid: false }));
  t('2t. 核准已失效（approved 但 valid=false）：整條 void 灰化＋橫幅，不能顯示全綠打勾', sts(v) === 'void,void' && v.banner.kind === 'void' && /失效/.test(v.headline) && !/st-done/.test(QFC.render(v)));
}

// ═════════════ 3) 代核標記、候選人 ═════════════
{
  let v = C(ap('pending', [S_MGR1('approved', { byName: '主管乙' }), S_GM('approved', { byName: '總座丙' }), S_BD('pending')], { cur: 2 }));
  t('3a. 董事會關（待簽）：proxy=true、顯示「管理部秘書」、不列個人候選人、不是 candidates', v.cards[2].proxy === true && v.cards[2].who === '管理部秘書' && v.cards[2].candidates === null && v.cards[0].proxy === false);
  let h = QFC.render(v);
  t('3b. 董事會卡 render 有「代核」徽章（qfc-proxy）；一般卡沒有', (h.match(/class="qfc-badge qfc-proxy"/g) || []).length === 1 && /代核/.test(h) && !/data-qfc-toggle="[^"]*c3"/.test(h));
  v = C(ap('approved', [S_GM('approved', { byName: '總座丙', at: '2026-10-02T03:00:00Z' }), S_BD('approved', { byName: '秘書甲', at: '2026-10-04T05:00:00Z' })], { valid: true, board: { resolutionDate: '2026-10-03', resolutionNo: '董字第1號', byName: '秘書甲', at: '2026-10-04T05:00:00Z' } }));
  t('3c. 董事會關已完成：代核人＝ap.board.byName、決議日期與文號、時間', v.cards[1].proxyBy === '秘書甲' && J(v.cards[1].board) === J({ resolutionDate: '2026-10-03', resolutionNo: '董字第1號' }) && v.cards[1].at === 'T:2026-10-04T05:00:00Z');
  h = QFC.render(v);
  t('3d. 代核完成的 render：「代核人：秘書甲」「決議：2026-10-03　文號 董字第1號」', /代核人：秘書甲/.test(h) && /決議：2026-10-03/.test(h) && /文號 董字第1號/.test(h));
  v = C(ap('pending', [S_BD('pending')], { cur: 0, board: null }));
  t('3e. 代核待簽時 board 為 null：不報錯、proxyBy 為空、無決議資訊', v.cards[0].proxyBy === '' && v.cards[0].board === null && !/代核人：/.test(QFC.render(v)));
  v = C(ap('pending', [S_BD('approved', { byName: '秘書乙' })], { valid: false, state: 'pending', cur: 0 }));
  t('3f. 代核人 fallback：沒有 ap.board 時用該步 byName', v.cards[0].proxyBy === '秘書乙');
  // 候選人
  v = C(ap('pending', [S_GM('pending', { assigneeName: '甲、乙、丙、丁、戊、己' })], { cur: 0 }));
  t('3g. 6 位候選人：candidates.count=6、who「6 位候選人」、名單完整', v.cards[0].candidates.count === 6 && v.cards[0].who === '6 位候選人' && v.cards[0].candidates.names.join('') === '甲乙丙丁戊己');
  h = QFC.render(v);
  t('3h. 候選人 render：按鈕 aria-expanded=false、aria-controls 指向 hidden 的 ul、6 個 li', /aria-expanded="false"/.test(h) && (() => { const m = /aria-controls="([^"]+)"/.exec(h); return m && new RegExp('<ul class="qfc-cand-list" id="' + m[1] + '" hidden>').test(h); })() && (h.match(/<ul class="qfc-cand-list"[\s\S]*?<\/ul>/)[0].match(/<li>/g) || []).length === 6);
  v = C(ap('pending', [S_GM('pending', { assigneeName: '總座丙' })], { cur: 0 }));
  t('3i. 名冊型但只有 1 人：不出現候選人展開鈕', v.cards[0].candidates === null && !/qfc-cand-btn/.test(QFC.render(v)));
  v = C(ap('pending', [S_MGR1('pending', { assigneeName: '主管乙、主管甲' })], { cur: 0 }));
  t('3j. 一級主管（非名冊型）即使名字含頓號也不當候選人（維持現況的單一指派）', v.cards[0].candidates === null);
  v = C(ap('pending', [S_GM('pending', { assigneeName: '' })], { cur: 0 }));
  t('3k. 沒有簽核人名字：不報錯，顯示「（尚無簽核人）」，headline 用關名', /尚無簽核人/.test(QFC.render(v)) && v.headline === '目前待 總經理 審核');
  v = C(ap('pending', [S_GM('pending', { assigneeName: '甲、乙、丙' })], { cur: 0 }));
  t('3l. 候選人 3 位也是候選人（≥2）', v.cards[0].candidates.count === 3);
}

// ═════════════ 4) 不洩漏毛利 ═════════════
{
  const LEAK_RE = /毛利|成本|margin|marginText|reasons|thresholds|revenue|costCents|gpCents|門檻|%|百分|12\.5|8,750,000|1000000000|875000000|125000000/i;
  const fixtures = [
    ap('pending', [S_MGR1('approved', { byName: '主管乙', at: '2026-10-02T02:00:00Z', comment: '同意' }), S_GM('pending'), S_CH('waiting')], { cur: 1 }),
    ap('returned', [S_MGR1('approved', { byName: '主管乙' }), S_GM('returned', { byName: '總座丙', comment: '再議' })], { cur: 1, valid: false }),
    ap('approved', [S_MGR1('approved', { byName: '主管乙' }), S_BD('approved', { byName: '秘書甲' })], { board: { resolutionDate: '2026-10-03', resolutionNo: 'X1', byName: '秘書甲' } }),
    ap('approved', [S_MGR1('approved', { byName: '主管乙' })], { valid: false }),
  ];
  let allClean = true, bad = '';
  fixtures.forEach((a, i) => {
    const v = C(a), h = QFC.render(v);
    if (LEAK_RE.test(J(v)) || LEAK_RE.test(h)) { allClean = false; bad += ' #' + i; }
  });
  t('4a. 業務視角序列化（含 derived.marginText／reasons／金額）：view-model 與 HTML 的 JSON 掃描不含毛利／成本／%／金額字串', allClean, bad);
  const same = fixtures.every((a) => { const b = JSON.parse(J(a)); delete b.derived; return J(C(a)) === J(C(b)) && QFC.render(C(a)).replace(/qfc\d+/g, 'q') === QFC.render(C(b)).replace(/qfc\d+/g, 'q'); });
  t('4b. 有沒有 derived（業務、管理員、簽核人視角的差別）輸出完全相同——卡片不受視角影響，也沒有任何依視角才出現的欄位', same);
  const noDerivedKeys = fixtures.every((a) => !('derived' in C(a)) && C(a).cards.every((c) => Object.keys(c).every((k) => !/margin|cost|reason|threshold|revenue|gp/i.test(k))));
  t('4c. view-model 與卡片的鍵名沒有 margin／cost／reason／threshold／revenue／gp', noDerivedKeys, Object.keys(C(fixtures[0]).cards[0]).join(','));
  // 意見是使用者輸入，不是系統洩漏：原樣（截斷）顯示
  const v = C(ap('pending', [S_MGR1('approved', { byName: '主管乙', comment: 'x'.repeat(100) }), S_GM('pending')], { cur: 1 }));
  t('4d. 意見摘要截斷在 40 字內（含…），詳細仍看簽核歷史', Array.from(v.cards[0].comment).length === 41 && v.cards[0].comment.endsWith('…'));
}

// ═════════════ 5) HTML：escape 與無障礙 ═════════════
{
  const evil = '<img src=x onerror=alert(1)>';
  const a = ap('pending', [S_MGR1('approved', { byName: evil, comment: '"><script>alert(1)</script>', at: '2026-10-02T02:00:00Z' }), S_GM('pending', { assigneeName: evil + '、乙' }), step('board', 'waiting', { label: evil, assigneeName: evil })], { cur: 1, board: { byName: evil } });
  const h = QFC.render(C(a));
  t('5a. XSS：姓名、意見、關名全部 escape（不出現原始 <img／<script，只有 &lt;）', !/<img|<script/i.test(h) && /&lt;img/.test(h) && /&lt;script/.test(h));
  t('5b. 屬性值內的雙引號被 escape（意見含 "> 不會跳出屬性）', !/"><script/.test(h));
  const h2 = QFC.render(C(ap('pending', [S_MGR1('approved', { byName: '主管乙' }), S_GM('pending'), S_CH('waiting')], { cur: 1 })));
  t('5c. 結構：ol[aria-label] 內 li；headline role=status；每張卡有 sr-only「第 n 關，共 m 關，狀態」', /<ol class="qfc-flow n3" aria-label="簽核關卡，共 3 關">/.test(h2) && /class="qfc-line" role="status"/.test(h2) && /第 2 關，共 3 關，目前關卡/.test(h2) && /第 3 關，共 3 關，後續關卡/.test(h2) && /第 1 關，共 3 關，已完成/.test(h2));
  t('5d. 狀態同時有文字＋符號（✓／●／○），符號 aria-hidden', /aria-hidden="true">✓<\/span> 已完成/.test(h2) && /aria-hidden="true">●<\/span> 目前關卡/.test(h2) && /aria-hidden="true">○<\/span> 後續關卡/.test(h2));
  t('5e. 每張卡有「流程節點：xxx核決」', /流程節點：一級主管核決/.test(h2) && /流程節點：總經理核決/.test(h2) && /流程節點：董事長核決/.test(h2));
  t('5f. 目前待句與標籤同時存在（右上角 qfc-tag）', /class="qfc-tag">目前待 /.test(h2));
  const hr = QFC.render(C(ap('returned', [S_MGR1('returned', { byName: '主管乙', comment: '退' })], { valid: false })));
  t('5g. 退回：符號 ✕ 與「已退回」文字；不是只有顏色', /aria-hidden="true">✕<\/span> 已退回/.test(hr));
  t('5h. 兩次 render 的候選人 id 不重複（同頁多張）', (() => { const x = QFC.render(C(ap('pending', [S_GM('pending', { assigneeName: '甲、乙' })], { cur: 0 }))); const y = QFC.render(C(ap('pending', [S_GM('pending', { assigneeName: '甲、乙' })], { cur: 0 }))); return /aria-controls="([^"]+)"/.exec(x)[1] !== /aria-controls="([^"]+)"/.exec(y)[1]; })());
}

// ═════════════ 6) 候選人展開（bind） ═════════════
{
  const listeners = [];
  const els = {};
  const mkEl = (attrs) => {
    const a = Object.assign({}, attrs);
    return { getAttribute: (k) => (k in a ? a[k] : null), setAttribute: (k, v) => { a[k] = String(v); }, removeAttribute: (k) => { delete a[k]; }, closest: function (sel) { return sel === '[data-qfc-toggle]' && 'data-qfc-toggle' in a ? this : null; }, _a: a };
  };
  const btn = mkEl({ 'data-qfc-toggle': 'qfcX-c1', 'aria-expanded': 'false' });
  const list = mkEl({ hidden: '', id: 'qfcX-c1' });
  const container = { addEventListener: (type, fn) => listeners.push([type, fn]), contains: () => true, querySelector: (s) => (s === '#qfcX-c1' ? list : null) };
  QFC.bind(container); QFC.bind(container);
  t('6a. bind 只綁一次（重複呼叫不重複掛監聽）', listeners.length === 1 && listeners[0][0] === 'click');
  const click = () => listeners[0][1]({ target: btn });
  click();
  t('6b. 點一下：aria-expanded=true、名單 hidden 移除', btn._a['aria-expanded'] === 'true' && !('hidden' in list._a));
  click();
  t('6c. 再點一下：收合回去', btn._a['aria-expanded'] === 'false' && 'hidden' in list._a);
  listeners[0][1]({ target: { closest: () => null } });
  t('6d. 點到其他地方不報錯', true);
}

// ═════════════ 7) ?cards=0 後備 ═════════════
t('7a. enabled：預設與 ?cards=1 為 true；?cards=0、?x=1&cards=0 為 false；?cards=00／?xcards=0 仍為 true', QFC.enabled('') === true && QFC.enabled('?cards=1') === true && QFC.enabled('?cards=0') === false && QFC.enabled('?x=1&cards=0') === false && QFC.enabled('?cards=00') === true && QFC.enabled('?xcards=0') === true);

// ═════════════ 8) CSS ═════════════
{
  const css = QFC.CSS;
  const lum = (hex) => { const n = parseInt(hex.slice(1), 16); const ch = [n >> 16, (n >> 8) & 255, n & 255].map((v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); }); return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2]; };
  const cr = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
  const modes = { light: css.split('\n').filter((l) => /^\.qfc \{ --st|^  --st/.test(l)).join(' ').split('body.dark')[0], dark: (css.match(/body\.dark \.qfc \{[^}]*\}/) || [''])[0] };
  const tok = (blk, k) => { const m = new RegExp('--st-' + k + ':(#[0-9a-fA-F]{6})').exec(blk); return m && m[1]; };
  let worst = 99, missing = [];
  ['light', 'dark'].forEach((md) => {
    ['done', 'cur', 'wait', 'ret', 'void'].forEach((s) => {
      const bg = tok(modes[md], s + '-bg'), fg = tok(modes[md], s + '-fg'), fg2 = tok(modes[md], s + '-fg2');
      if (!bg || !fg || !fg2) { missing.push(md + ':' + s); return; }
      worst = Math.min(worst, cr(fg, bg), cr(fg2, bg));
    });
    const pbg = tok(modes[md], 'proxy-bg'), pfg = tok(modes[md], 'proxy-fg');
    if (!pbg || !pfg) missing.push(md + ':proxy'); else worst = Math.min(worst, cr(pfg, pbg));
  });
  t('8a. 5 狀態＋代核 × 淺色／深色：前景（含次要文字）對背景對比全部 ≥ 4.5:1（最差值見右）', missing.length === 0 && worst >= 4.5, missing.length ? 'missing ' + missing : worst.toFixed(2));
  t('8b. 邊框色對卡片背景 ≥ 3:1（非文字 UI 元件）', (() => { let w = 99; ['light', 'dark'].forEach((md) => ['done', 'cur', 'wait', 'ret', 'void'].forEach((s) => { const b = tok(modes[md], s + '-bd'), g = tok(modes[md], s + '-bg'); if (b && g) w = Math.min(w, cr(b, g)); })); return w >= 3 || (console.log('   (info) 最低邊框對比 ' + w.toFixed(2)), false); })());
  t('8c. 所有狀態色是不透明 #rrggbb（與玻璃半透明底／六種主題無關）', !/--st-[a-z0-9-]+:\s*rgba?\(/.test(css));
  t('8d. prefers-reduced-motion：reduce 時關閉動畫與過渡；脈動只在 no-preference', /@media \(prefers-reduced-motion: reduce\)[^{]*\{[^}]*animation: none/.test(css) && /@media \(prefers-reduced-motion: no-preference\)[^{]*\{[^}]*animation: qfcpulse/.test(css));
  t('8e. 預設直排（單欄）；容器夠寬才橫排（@container n2..n5）', /\.qfc-flow \{[^}]*grid-template-columns: minmax\(0, 1fr\)/.test(css) && ['n2', 'n3', 'n4', 'n5'].every((n) => new RegExp('@container \\(min-width: \\d+px\\) \\{ \\.qfc-flow\\.' + n + ' \\{ grid-template-columns: repeat').test(css)));
  t('8f. 長字串不撐破：卡片文字 overflow-wrap:anywhere、min-width:0', /overflow-wrap: anywhere/.test(css) && /\.qfc-step \{[^}]*min-width: 0/.test(css));
  t('8g. 焦點樣式、forced-colors', /qfc-cand-btn:focus-visible/.test(css) && /forced-colors: active/.test(css));
  t('8h. CSS 全部 qfc-／st-／qfc 前綴，不污染既有 class（.qap-／.btn 等不被覆寫）', !/(^|\n)\s*\.(qap|btn|modal|card)\b/.test(css));
}

// ═════════════ 9) 接線 ═════════════
{
  const rd = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
  const html = rd('_client/index.html'), qa = rd('_client/quote-approval.js'), srv = rd('server.js'), sw = rd('_client/sw.js');
  t('9a. index.html 先載入 quote-flow-cards.js、再載入 quote-approval.js', html.indexOf('quote-flow-cards.js') > 0 && html.indexOf('quote-flow-cards.js') < html.indexOf('quote-approval.js'));
  t('9b. quote-approval.js：卡片接線（cardsOf／render／bind／CSS 注入）且保留舊條列式（qap-steps、STEP_STATUS）作為 ?cards=0 後備', /QFlowCards\.cardsOf\(q\.approval/.test(qa) && /QFlowCards\.bind\(s\.body\)/.test(qa) && /QFlowCards\.CSS/.test(qa) && /class="qap-steps"/.test(qa) && /const STEP_STATUS/.test(qa));
  t('9c. quote-approval.js 的卡片只在 enabled() 時使用（?cards=0 走舊視圖）', /if \(typeof QFlowCards !== 'undefined' && QFlowCards\.enabled\(\)\)/.test(qa));
  t('9d. server.js 為 quote-flow-cards.js 加版本參數', /quote-flow-cards/.test(srv));
  const m = /SW_VERSION = 'itts-crm-v(\d+)'/.exec(sw);
  t('9e. sw.js 版號 ≥ v15（本期 +1）', m && +m[1] >= 15, m && m[0]);
  t('9f. 流程引擎與規則釘選未動：lib/quoteApproval.js 的 RULES_HASH 仍是釘選值', /RULES_HASH = '76efc01de2f3449f0371ed3217bc1e1469ad1282bfae8066bfacef605de76c16'/.test(rd('lib/quoteApproval.js')));
}

// ── 輸出 ──
let fail = 0;
res.forEach(([n, ok, ex]) => { if (!ok) fail++; console.log((ok ? 'PASS ' : 'FAIL ') + n + (ex && (!ok || /^\d|worst/.test(ex)) ? '  [' + ex + ']' : '')); });
console.log('\n' + (res.length - fail) + '/' + res.length + ' 通過' + (fail ? '，' + fail + ' 失敗' : ''));
process.exit(fail ? 1 : 0);
