#!/usr/bin/env node
/**
 * 「顧問填成本逐步填寫」(quote-coststeps.js) 純函式檢查。用法：node scripts/check-quote-coststeps.js
 * 動 _client/quote-coststeps.js 的判定規則（哪些步驟算完成、略過、重新打開、整批操作取消確認、窄螢幕斷點）之後必跑。
 * 畫面行為（步驟一步步展開、編輯器搬到各步驟、略過按鈕、標紅、Enter、窄螢幕、暗色、XSS）另以無頭瀏覽器實跑（e2e_ui_coststeps.js）。
 * 與新增報價單逐步填寫（quote-steps.js）共用的樣式／做法（進度欄樣式、摘要卡、確認清單的版面）不在這裡重複測。
 *   1) 介面與常數：步驟順序與名稱、窄螢幕斷點與 quote-steps.js 同一個數字
 *   2) catMiss：成本分區的檢查（沒有列要按略過、各項欄位檢查的字串與列號清單）
 *   3) itemsMiss／riskMiss
 *   4) evaluate：步驟狀態（第一個沒完成的步驟、修改、佔位、檢視、變不完整就收回確認）——已知案例＋隨機輸入的不變量
 *   5) 重新打開規則：isReopen／initialFlags，加上「走一遍」的情境（草稿完整、缺風險預留、某區不完整）
 *   6) pruneSkipped：有列的分區略過旗標失效
 *   7) sigOf／bulkChanged：整批操作（重新帶入／補入新品項）取消不在畫面上的分區的確認
 *   8) catSummary／confirmRow：摘要與確認清單的列
 *   9) CSS 與檔案紀律：斷點、選擇器範圍、暗色、沒有 innerHTML／eval、沒有本機路徑
 *  10) 掛鉤：quote-approval.js／index.html／server.js 有接好（單一來源的確認視窗文字、IIFE 參數、版本注入）
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、廠商名、真實費率）。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const SRC_PATH = process.env.CST_SRC || path.join(ROOT, '_client/quote-coststeps.js');   // 變異測試用：指向被破壞的副本

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
function lcg(seed) { let s = seed >>> 0; return () => ((s = (Math.imul(s, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff); }

const src = fs.readFileSync(SRC_PATH, 'utf8').replace(/\r\n/g, '\n');
const ctx = {};
vm.createContext(ctx);
vm.runInContext(src, ctx);
const CS = ctx.CostSteps;
const C = CS.core;
const same = (a, b) => J(a) === J(b);   // 跨 realm 的物件比對：一律經 JSON

// ═════════════ 1) 介面與常數 ═════════════
t('1a. CostSteps 對外函式：allowed／attach／refresh／syncFooter／canFinish／state／core', ['allowed', 'attach', 'refresh', 'syncFooter', 'canFinish', 'state'].every((k) => typeof CS[k] === 'function') && typeof C === 'object');
t('1b. 步驟順序：確認報價品項→顧問服務成本→軟體成本→硬體成本→差旅費用→其他費用→風險預留→確認並完成（8 步）',
  same(C.STEP_IDS, ['items', 'consult', 'software', 'hw', 'travel', 'other', 'risk', 'confirm']) && same(C.STEP_IDS.map((id) => C.STEP_NAME[id]), ['確認報價品項', '顧問服務成本', '軟體成本', '硬體成本', '差旅費用', '其他費用', '風險預留', '確認並完成']));
t('1c. 五個成本分區 id 與成本明細編輯器（QCL）的分區 key 相同', same(C.CAT_IDS, ['consult', 'software', 'hw', 'travel', 'other']));
{
  const steps = fs.readFileSync(path.join(ROOT, '_client/quote-steps.js'), 'utf8');
  const m1 = steps.match(/@media \(max-width: (\d+)px\)/), m2 = C.cssText().match(/@media \(max-width: (\d+)px\)/);
  t('1d. 窄螢幕斷點與新增報價單的逐步填寫（quote-steps.js）同一個數字（639）', C.NARROW_MAX === 639 && m1 && m2 && m1[1] === m2[1] && Number(m2[1]) === C.NARROW_MAX, J([C.NARROW_MAX, m1 && m1[1], m2 && m2[1]]));
  t('1e. isNarrow：638／639 窄、640／1400 寬（≤639）', C.isNarrow(375) && C.isNarrow(638) && C.isNarrow(639) && !C.isNarrow(640) && !C.isNarrow(1400) && !C.isNarrow(NaN));
}

// ═════════════ 2) catMiss ═════════════
const sec = (o) => Object.assign({ n: 1, invalid: [], blankDesc: [], zeroQty: [], badTarget: [], conflict: [] }, o || {});
t('2a. 沒有任何列、沒略過：缺「成本明細」並寫明要按「這一類沒有成本，略過」', same(C.catMiss(sec({ n: 0 }), false), ['成本明細（沒有這一類成本請按「這一類沒有成本，略過」）']));
t('2b. 沒有任何列、按過略過：有效', same(C.catMiss(sec({ n: 0 }), true), []));
t('2c. 有列時略過旗標不再有作用（沒問題＝有效；有問題照樣缺）', same(C.catMiss(sec({ n: 2 }), true), []) && C.catMiss(sec({ n: 2, blankDesc: [2] }), true).length === 1);
t('2d. 項目名稱空白：「第 2 列的項目名稱」', same(C.catMiss(sec({ n: 3, blankDesc: [2] }), false), ['第 2 列的項目名稱']));
t('2e. 數量 0：「第 1 列的數量（需大於 0）」', same(C.catMiss(sec({ zeroQty: [1] }), false), ['第 1 列的數量（需大於 0）']));
t('2f. 數字不合法：「第 1、3 列的數量或成本單價（需為 0 以上的數字）」', same(C.catMiss(sec({ n: 3, invalid: [1, 3] }), false), ['第 1、3 列的數量或成本單價（需為 0 以上的數字）']));
t('2g. 對應目標無效：寫明連動／拆項要選報價品項或改不對應', /^第 2 列的「對應」報價品項（連動／拆項要選一個報價品項，或改成「不對應」）$/.test(C.catMiss(sec({ n: 2, badTarget: [2] }), false)[0]));
t('2h. 連動單位衝突：寫明單位需一致或改拆項', /^第 1、2 列的連動單位（同一個報價品項的連動列單位需一致，或把其中幾列改成「拆項」）$/.test(C.catMiss(sec({ n: 2, conflict: [1, 2] }), false)[0]));
t('2i. 多項問題依固定順序全部列出（數字→名稱→數量→對應→單位）', C.catMiss(sec({ n: 5, invalid: [1], blankDesc: [2], zeroQty: [3], badTarget: [4], conflict: [5] }), false).map((x) => x.replace(/^第 \d 列的/, '').slice(0, 4)).join('|') === '數量或成|項目名稱|數量（需|「對應」|連動單位');
t('2j. 列號清單最多列 6 個，超過接「…等 N 列」', C.rowList([1, 2, 3, 4, 5, 6]) === '1、2、3、4、5、6' && C.rowList([1, 2, 3, 4, 5, 6, 7, 8]) === '1、2、3、4、5、6…等 8 列' && C.rowList([]) === '' && C.rowList(null) === '');
t('2k. 輸入缺欄位（null／{}）不丟例外：沒有列', same(C.catMiss(null, false), ['成本明細（沒有這一類成本請按「這一類沒有成本，略過」）']) && same(C.catMiss({}, true), []));

// ═════════════ 3) itemsMiss／riskMiss ═════════════
t('3a. 沒有問題：有效', same(C.itemsMiss({ problem: '', zeroNew: [], total: 10 }), []) && same(C.itemsMiss(null), []));
t('3b. 新增項目欄位問題原文帶出；數量 0 的新增項目列名；總列數超過 50 要減少', same(C.itemsMiss({ problem: '新增的報價項目有一列沒有填品名，請填寫或移除該列', zeroNew: [], total: 3 }), ['新增的報價項目有一列沒有填品名，請填寫或移除該列'])
  && /^新增的報價項目數量需大於 0：「甲」、「乙」$/.test(C.itemsMiss({ zeroNew: ['甲', '乙'], total: 3 })[0]) && /超過 50 列/.test(C.itemsMiss({ total: 51 })[0]) && C.itemsMiss({ total: 50 }).length === 0);
t('3c. riskMiss：沒選（\'\'、null、undefined）缺；0 與其他百分比字串有效（「0」是有效選擇）', C.riskMiss('').length === 1 && C.riskMiss(null).length === 1 && C.riskMiss(undefined).length === 1 && C.riskMiss('0').length === 0 && C.riskMiss('10').length === 0 && /沒有風險請選 0%/.test(C.riskMiss('')[0]));

// ═════════════ 4) evaluate ═════════════
const IDS = C.STEP_IDS;
const ev = (miss, conf, editing, review) => C.evaluate(IDS, miss || {}, conf || {}, editing || null, !!review);
const stMap = (r) => r.steps.map((s) => s.st).join(',');
const allConf = {}; IDS.slice(0, 7).forEach((id) => { allConf[id] = true; });
{
  let r = ev({}, {});
  t('4a. 什麼都沒確認：步驟 1 進行中，其餘（含確認）隱藏；第一個沒完成＝0；缺 7 步', stMap(r) === 'active,future,future,future,future,future,future,future' && r.fi === 0 && r.cur === 0 && !r.allDone && r.missingN === 7, stMap(r));
  r = ev({}, { items: true, consult: true });
  t('4b. 確認了前 2 步：前 2 步 done、第 3 步進行中', stMap(r) === 'done,done,active,future,future,future,future,future' && r.fi === 2 && r.missingN === 5, stMap(r));
  r = ev({}, allConf);
  t('4c. 前 7 步都確認且有效：全部完成，確認並完成進行中', stMap(r) === 'done,done,done,done,done,done,done,active' && r.allDone && r.fi === 7 && r.cur === 7 && r.missingN === 0, stMap(r));
  r = ev({ hw: ['成本明細'] }, allConf);
  t('4d. 硬體變不完整（有 miss）：硬體的確認被收回；停在硬體，後面（差旅、其他、風險、確認）隱藏但它們仍是已確認', stMap(r) === 'done,done,done,active,future,future,future,future' && !r.confirmed.hw && r.confirmed.travel === true && r.steps[4].done === true && r.fi === 3, stMap(r));
  r = ev({}, Object.assign({}, allConf, { confirm: true }));
  t('4e. 確認步驟永遠不算 done（按不到下一步）；confirmed.confirm 被忽略', r.steps[7].done === false && r.steps[7].st === 'active');
  r = ev({}, allConf, 'hw');
  t('4f. 全部完成時修改某一步：該步 edit、確認隱藏、其餘維持 done；editing 保留', stMap(r) === 'done,done,done,edit,done,done,done,future' && r.editing === 'hw' && r.cur === 3 && r.ei === 3, stMap(r));
  r = ev({}, { items: true, consult: true }, 'items');
  t('4g. 還沒走完時修改較前面的步驟：該步 edit、目前進度（軟體）變 pending 佔位', stMap(r) === 'edit,done,pending,future,future,future,future,future' && r.cur === 0 && r.fi === 2, stMap(r));
  r = ev({}, { items: true, consult: true }, 'consult');
  r = ev({}, { items: true, consult: true }, 'software');
  t('4h. 修改的步驟就是目前進度步驟：視為一般進行中（不是 edit）', r.steps[2].st === 'active' && r.editing === 'software', stMap(r));
  r = ev({}, { items: true, consult: true }, 'travel');
  t('4i. 修改的步驟在目前進度之後（還沒到）或是確認步驟：editing 被清掉', r.editing === null && r.ei === -1 && ev({}, allConf, 'confirm').editing === null && ev({}, allConf, 'nope').editing === null);
  r = ev({}, { items: true }, null, true);
  t('4j. 沒走完時檢視確認清單（review）：確認步驟 review、目前進度步驟仍進行中；有人在修改或全部完成時 review 自動關閉', r.steps[7].st === 'review' && r.steps[1].st === 'active' && r.review === true
    && ev({}, allConf, null, true).review === false && ev({}, { items: true, consult: true }, 'items', true).review === false);
  r = ev({ consult: ['x'], risk: ['y'] }, allConf);
  t('4k. 多步不完整：停在第一個（顧問服務）；missingN 只算沒完成的（顧問＋風險 = 2；其餘仍是已確認）', r.fi === 1 && r.missingN === 2 && r.steps[6].miss[0] === 'y', J([r.fi, r.missingN]));
  r = ev({ items: ['z'] }, { items: true });
  t('4l. 步驟 1 有 miss：即使按過下一步也不算完成（confirmed 被清）', r.steps[0].done === false && !r.confirmed.items && r.steps[0].st === 'active');
  // 隨機不變量
  const rnd = lcg(20261009);
  let bad = [];
  for (let k = 0; k < 6000; k++) {
    const n = 3 + Math.floor(rnd() * 8), ids = []; for (let i = 0; i < n; i++) ids.push('s' + i);
    const miss = {}, conf = {}; ids.forEach((id) => { if (rnd() < 0.3) miss[id] = ['m']; if (rnd() < 0.6) conf[id] = true; });
    const editing = rnd() < 0.4 ? ids[Math.floor(rnd() * (n + 1))] || null : null, review = rnd() < 0.3;
    const r = C.evaluate(ids, miss, conf, editing, review);
    const last = n - 1;
    const doneAt = (i) => r.steps[i].done;
    const fail = (why) => bad.length < 5 && bad.push(why + ' ' + J({ ids: n, miss, conf, editing, review }));
    if (r.steps[last].done) fail('I1 確認步驟不可 done');
    r.steps.forEach((s, i) => { if (s.miss.length && r.confirmed[s.id]) fail('I5 有 miss 還保留確認'); if (s.done !== (i !== last && !s.miss.length && !!conf[s.id])) fail('done 判定'); });
    let fi = last; for (let i = 0; i < last; i++) if (!doneAt(i)) { fi = i; break; }
    if (fi !== r.fi) fail('I2 fi');
    if (r.allDone !== (fi === last)) fail('I3 allDone');
    for (let i = 0; i < r.fi; i++) if (!doneAt(i)) fail('I2b 第一個沒完成之前都要 done');
    const working = r.steps.filter((s) => s.st === 'active' || s.st === 'edit');
    const expectWorking = r.allDone && r.ei < 0 ? 1 : 1;   // 全部完成不修改：確認步驟 active；其餘情況：cur 那一步 active／edit
    if (working.length < expectWorking || working.length > 1) fail('I4 進行中的步驟只能有一個 ' + working.length);
    if (r.steps[r.cur].st !== 'active' && r.steps[r.cur].st !== 'edit') fail('I4b cur 必須是進行中');
    if (r.ei >= 0 && (r.ei > r.fi || r.ei === last || r.steps[r.ei].id !== r.editing)) fail('I7 editing 不合法');
    if (r.ei < 0 && r.editing !== null) fail('I7b editing 應清掉');
    if ((r.ei >= 0 || r.allDone) && r.review) fail('review 應關閉');
    r.steps.forEach((s, i) => {
      if (s.st === 'done' && !(i < r.fi && s.done)) fail('done 狀態不合法');
      if (s.st === 'pending' && !(i === r.fi && r.ei >= 0 && r.ei < r.fi)) fail('pending 不合法');
      if (s.st === 'future' && i < r.fi && i !== r.ei) fail('第一個沒完成之前不能 future');
    });
    const again = C.evaluate(ids, miss, r.confirmed, r.editing, r.review);   // 冪等：把結果餵回去不變
    if (J(again.steps) !== J(r.steps) || again.cur !== r.cur) fail('I8 冪等');
    if (r.missingN !== r.steps.filter((s, i) => i < last && !s.done).length) fail('missingN');
  }
  t('4m. 隨機 6000 組（步驟數 3–10、隨機 miss／確認／修改／檢視）的不變量：確認步驟不 done、miss 必收回確認、第一個沒完成之前都 done、只有一個進行中、editing 合法、review 規則、結果冪等', bad.length === 0, bad.join(' ## '));
}

// ═════════════ 5) 重新打開規則 ═════════════
t('5a. isReopen：costLines 是陣列（含空陣列）＝重新打開；沒有／null／非陣列＝第一次填', C.isReopen({ costLines: [] }) === true && C.isReopen({ costLines: [{ cat: 'consult' }] }) === true && C.isReopen({}) === false && C.isReopen({ costLines: null }) === false && C.isReopen(null) === false && C.isReopen({ costLines: 'x' }) === false);
{
  const f1 = C.initialFlags(IDS, false, { consult: 0, software: 0 });
  t('5b. 第一次填：沒有任何步驟預先確認、沒有略過旗標（空分區也要使用者自己按略過）', same(f1, { confirmed: {}, skipped: {} }));
  const f2 = C.initialFlags(IDS, true, { consult: 2, software: 0, hw: 0, travel: 1, other: 0 });
  t('5c. 重新打開：除「確認並完成」外全部預先確認；目前沒有列的分區（軟體、硬體、其他）標成已略過', same(Object.keys(f2.confirmed), IDS.slice(0, 7)) && same(f2.skipped, { software: true, hw: true, other: true }), J(f2));
  // 走一遍：草稿完整
  const counts = { consult: 1, software: 0, hw: 0, travel: 1, other: 0 };
  const miss = {}; C.CAT_IDS.forEach((c) => { miss[c] = C.catMiss(sec({ n: counts[c] }), !!f2.skipped[c]); }); miss.items = []; miss.risk = C.riskMiss('5');
  let r = C.evaluate(IDS, miss, f2.confirmed, null, false);
  t('5d. 重新打開、草稿完整：直接停在「確認並完成」（allDone），前 7 步 done', r.allDone && r.cur === 7 && r.steps.slice(0, 7).every((s) => s.st === 'done'), stMap(r));
  miss.risk = C.riskMiss('');
  r = C.evaluate(IDS, miss, f2.confirmed, null, false);
  t('5e. 重新打開、草稿沒有風險預留：停在「風險預留」（第一個沒通過的步驟），確認並完成還沒出現', !r.allDone && r.cur === 6 && r.steps[7].st === 'future', stMap(r));
  miss.risk = []; miss.software = C.catMiss(sec({ n: 2, blankDesc: [1] }), false);
  r = C.evaluate(IDS, miss, f2.confirmed, null, false);
  t('5f. 重新打開、草稿的軟體區有空白項目名稱：停在「軟體成本」，後面的步驟隱藏（但仍是已確認）', r.cur === 2 && r.steps[3].st === 'future' && r.steps[3].done === true, stMap(r));
  // 重新打開之後使用者把有列的分區刪光 → 要重新按略過
  const pr = C.pruneSkipped(f2.skipped, { software: 0, hw: 1, other: 0 });
  t('5g. 重新打開後的略過旗標：某區又有列了（硬體 1 列）→ 旗標失效；其他仍在', same(pr, { software: true, other: true }), J(pr));
}

// ═════════════ 6) pruneSkipped ═════════════
t('6a. pruneSkipped：有列（>0）的分區清掉旗標；沒有列的保留；空輸入不丟例外', same(C.pruneSkipped({ software: true, hw: true }, { software: 0, hw: 3 }), { software: true }) && same(C.pruneSkipped(null, null), {}) && same(C.pruneSkipped({ hw: false }, { hw: 0 }), {}));
t('6b. 刪光後再略過的完整情境：略過→有效；新增一列→旗標清掉且要檢查欄位；刪光→回到要重新按略過', (() => {
  let sk = { hw: true };
  const m0 = C.catMiss(sec({ n: 0 }), !!sk.hw);
  sk = C.pruneSkipped(sk, { hw: 1 });
  const m1 = C.catMiss(sec({ n: 1, blankDesc: [1] }), !!sk.hw);
  sk = C.pruneSkipped(sk, { hw: 0 });
  const m2 = C.catMiss(sec({ n: 0 }), !!sk.hw);
  return m0.length === 0 && m1.length === 1 && m2.length === 1;
})());

// ═════════════ 7) sigOf／bulkChanged ═════════════
{
  const L = { desc: 'a', qty: 1, unit: '式', unitCost: 5, rel: 'link', forLid: 'x', forLids: ['y'], vendor: 'v', consultant: 'c', note: 'n' };
  const base = C.sigOf([L]);
  const diffs = ['desc', 'qty', 'unit', 'unitCost', 'rel', 'forLid', 'vendor', 'consultant', 'note'].filter((k) => C.sigOf([Object.assign({}, L, { [k]: k === 'qty' || k === 'unitCost' ? 99 : 'zz' })]) === base);
  t('7a. sigOf：項目／數量／單位／成本單價／對應／目標／廠商／顧問姓名／說明任一改變，簽章都不同', diffs.length === 0, diffs.join(','));
  t('7b. sigOf：forLids 改變、列順序改變、列數改變都不同；同內容相同；無關欄位（lid）不影響；空陣列／null 不丟例外', C.sigOf([Object.assign({}, L, { forLids: ['z'] })]) !== base && C.sigOf([L, Object.assign({}, L, { desc: 'b' })]) !== C.sigOf([Object.assign({}, L, { desc: 'b' }), L])
    && C.sigOf([L, L]) !== base && C.sigOf([Object.assign({ lid: 'k' }, L)]) === base && C.sigOf([]) === '' && C.sigOf(null) === '');
  const prev = { consult: 'a', software: 'b', hw: 'c', travel: 'd', other: 'e' }, now = { consult: 'A', software: 'B', hw: 'c', travel: 'D', other: 'E' };
  const conf = { consult: true, software: true, hw: true, travel: false };
  t('7c. bulkChanged：畫面上顯示的分區（consult）自己的改動不算；其餘「先前已確認且內容變了」的才取消（software）；沒變的（hw）、沒確認的（travel、other）不動', same(C.bulkChanged(prev, now, 'consult', conf), ['software']), J(C.bulkChanged(prev, now, 'consult', conf)));
  t('7d. bulkChanged：沒有任何分區顯示（shown＝null，例如在風險預留）時，所有已確認且變了的都取消；沒有先前簽章（undefined）不算變', same(C.bulkChanged(prev, now, null, { consult: true, software: true, hw: true }), ['consult', 'software']) && same(C.bulkChanged({}, now, null, { consult: true }), []) && same(C.bulkChanged(null, now, null, null), []));
  const rnd = lcg(7);
  let bad = 0;
  for (let k = 0; k < 3000; k++) {
    const p = {}, n = {}, cf = {};
    C.CAT_IDS.forEach((c) => { p[c] = 's' + Math.floor(rnd() * 3); n[c] = rnd() < 0.5 ? p[c] : 's' + Math.floor(rnd() * 3); if (rnd() < 0.6) cf[c] = true; });
    const shown = rnd() < 0.3 ? null : C.CAT_IDS[Math.floor(rnd() * 5)];
    const got = C.bulkChanged(p, n, shown, cf);
    const exp = C.CAT_IDS.filter((c) => c !== shown && p[c] !== n[c] && cf[c]);
    if (J(got) !== J(exp)) bad++;
  }
  t('7e. bulkChanged 隨機 3000 組與直接寫法逐組相同', bad === 0, bad);
}

// ═════════════ 8) catSummary／confirmRow ═════════════
t('8a. catSummary：有列「n 列｜小計」；顧問服務有委外多「委外 …」；其他費用有印花稅多「含印花稅 …」；單位是分、四捨五入到元、千分位',
  C.catSummary('consult', { n: 2 }, 1234500, { outsourcedCents: 50000 }) === '2 列｜小計 NT$ 12,345｜委外 NT$ 500' && C.catSummary('consult', { n: 1 }, 100, { outsourcedCents: 0 }) === '1 列｜小計 NT$ 1'
  && C.catSummary('other', { n: 1 }, 220000, { stampCents: 20000 }) === '1 列｜小計 NT$ 2,200｜含印花稅 NT$ 200' && C.catSummary('travel', { n: 3 }, 0, {}) === '3 列｜小計 NT$ 0');
t('8b. catSummary 沒有列（略過）：「沒有這一類成本（已略過）」；其他費用沒有列但有印花稅：另列印花稅金額', C.catSummary('software', { n: 0 }, 0, {}) === '沒有這一類成本（已略過）' && C.catSummary('other', { n: 0 }, 0, { stampCents: 15400 }) === '沒有其他費用（已略過）｜印花稅 NT$ 154（依合約金額自動計算）'
  && C.catSummary('other', { n: 0 }, 0, { stampCents: 0 }) === '沒有這一類成本（已略過）');
{
  const S = (o) => Object.assign({ id: 'x', miss: [], valid: true, done: false, st: 'active' }, o);
  t('8c. confirmRow：已完成＝✓＋「修改」；略過＝–；缺＝✗＋原因（目前進度那一列才有「前往填寫」）；已有效但沒按下一步＝○（目前進度那一列「前往確認」）',
    same(C.confirmRow(S({ done: true }), 0, 3, '摘要', false), { row: 'ok', mark: '✓', text: '摘要', btn: '修改' }) && same(C.confirmRow(S({ done: true }), 2, 3, '沒有這一類成本（已略過）', true), { row: 'skip', mark: '–', text: '沒有這一類成本（已略過）', btn: '修改' })
    && same(C.confirmRow(S({ miss: ['A', 'B'] }), 3, 3, '', false), { row: 'bad', mark: '✗', text: '還缺：A、B', btn: '前往填寫' }) && same(C.confirmRow(S({}), 3, 3, '', false), { row: 'todo', mark: '○', text: '尚未按「下一步」確認', btn: '前往確認' }));
  t('8d. confirmRow：不是目前進度的未完成步驟沒有按鈕，並註明「前面的步驟完成後才能填」', same(C.confirmRow(S({ miss: ['A'] }), 5, 3, '', false), { row: 'bad', mark: '✗', text: '還缺：A（前面的步驟完成後才能填）', btn: '' })
    && same(C.confirmRow(S({}), 5, 3, '', false), { row: 'todo', mark: '○', text: '尚未按「下一步」確認（前面的步驟完成後才能填）', btn: '' }));
}

// ═════════════ 9) CSS 與檔案紀律 ═════════════
{
  const css = C.cssText().replace(/\/\*[\s\S]*?\*\//g, '');   // 去掉註解再解析選擇器
  const sels = [];
  css.replace(/(^|\})\s*([^{}@]+)\{/g, (m, a, sel) => { sel.split(',').forEach((x) => { const s = x.trim(); if (s && !/^(from|to|\d+%)$/.test(s)) sels.push(s); }); return m; });
  const okSel = (s) => /^(body\.dark\s+)?(\.qap-cs\b|\.qap-body\.cs-mode\b|\.qap-cf-live\.cs-sticky\b)/.test(s);
  const offenders = sels.filter((s) => !okSel(s));
  t('9a. 這個檔案的 CSS 選擇器全部限定在 .qap-cs／.qap-body.cs-mode／.qap-cf-live.cs-sticky 底下（含 body.dark 版）：不會改到新增報價單的逐步畫面或其他頁面', sels.length > 40 && offenders.length === 0, offenders.slice(0, 4).join(' | '));
  const light = sels.filter((s) => !/^body\.dark/.test(s)), dark = sels.filter((s) => /^body\.dark/.test(s));
  t('9b. 有深色模式版本：淺色的主要元件（進度欄項目、提醒框、合計）都有對應的 body.dark 規則', ['.qs-ri', '.cs-note', '.cs-tot', '.cs-sticky'].every((k) => dark.some((s) => s.indexOf(k) >= 0)) && light.length > dark.length / 2);
  t('9c. 動畫都有 prefers-reduced-motion 例外', /@keyframes csFlash/.test(css) && /prefers-reduced-motion: reduce\) \{ \.qap-cs \.qs-hint\.flash \{ animation: none/.test(css));
  const body = src.replace(/\/\/.*$/gm, '');
  t('9d. 沒有 innerHTML／insertAdjacentHTML／outerHTML／document.write／eval／new Function（使用者輸入的文字只用 textContent）', !/innerHTML|insertAdjacentHTML|outerHTML|document\.write\(|\beval\(|new Function\(/.test(body));
  t('9e. 沒有本機路徑／帳密字樣', !/[A-Za-z]:[\\/]Users|C:\\|\/Users\/|password\s*[:=]\s*['"][^'"]{4,}|api[_-]?key\s*[:=]\s*['"][^'"]{8,}/i.test(src));
  t('9f. 測試開關只在 localhost 且 window.__qNoSteps === true 時生效（與新增報價單同一條件）', /location\.hostname === 'localhost' && window\.__qNoSteps === true/.test(src));
  t('9g. 失敗保護：例外 → fail() 關掉逐步並重畫完整畫面（renderCostFill 由 api.rerender 交進來）', /function fail\(s, e\)/.test(src) && /s\.cstOff = true/.test(src) && /s\.cstApi\.rerender\(s\)/.test(src));
}

// ═════════════ 10) 掛鉤 ═════════════
{
  const qap = fs.readFileSync(path.join(ROOT, '_client/quote-approval.js'), 'utf8');
  const html = fs.readFileSync(path.join(ROOT, '_client/index.html'), 'utf8');
  const server = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  t('10a. quote-approval.js 的四個掛鉤：renderCostFill 尾端 attach、cfRecalc 尾端 refresh、setCostBusy 尾端 syncFooter、submitCostFill 的 canFinish 防線', /CostSteps\.attach\(s, CF_STEPS_API\)/.test(qap) && /CostSteps\.refresh\(s\)/.test(qap) && /CostSteps\.syncFooter\(s\)/.test(qap) && /!CostSteps\.canFinish\(s\)/.test(qap));
  t('10b. 只有「可編輯＋完整連動畫面」才分步（attach 在 sync && can 的條件裡）；舊格式／唯讀不呼叫', /if \(sync && can && typeof CostSteps !== 'undefined' && CostSteps\.allowed\(s\)\)/.test(qap));
  t('10c. 確認視窗文字只有一份：submitCostFill 與逐步畫面都走 cfDoneMessage（IIFE 內的函式用 CF_STEPS_API 交給 CostSteps）', /const msg = cfDoneMessage\(s, c, live, sync\);/.test(qap) && /doneMessage: \(s, c, live, sync\) => cfDoneMessage\(s, c, live, sync\)/.test(qap) && /riskValue: \(s\) => costRiskValue\(s\)/.test(qap) && /newItemsProblem: \(s, mark\) => cfNewItemsProblem\(s, mark\)/.test(qap) && /rerender: \(s\) => renderCostFill\(s\)/.test(qap));
  t('10d. 風險預留區塊有 id（#qapCfRisk，逐步畫面靠它把區塊搬進步驟）', /<div class="qap-sec" id="qapCfRisk">/.test(qap));
  t('10e. index.html 在 quote-approval.js 之後載入 quote-coststeps.js；server.js 為它加版本號（與其他 quote*.js 同一模式）', /<script src="quote-approval\.js"><\/script>\s*<script src="quote-coststeps\.js"><\/script>/.test(html) && /quote-coststeps\\\.js/.test(server));
  const steps = fs.readFileSync(path.join(ROOT, '_client/quote-steps.js'), 'utf8');
  t('10f. quote-steps.js 只多了共用的 ensureStyle 匯出（新增報價單的流程不變：setup 內用它注入，和以前一樣只注入一次）', /ensureStyle: ensureStyle,/.test(steps) && /function ensureStyle\(\)/.test(steps) && /ensureStyle\(\);\s*\n\s*bindOnce\(body\);/.test(steps));
}

let pass = 0, fail = 0;
res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
console.log(`\n顧問填成本逐步填寫單元檢查：PASS ${pass} / FAIL ${fail}`);
process.exit(fail ? 1 : 0);
