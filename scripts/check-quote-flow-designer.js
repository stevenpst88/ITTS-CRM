#!/usr/bin/env node
/**
 * 簽核流程設計器 M2 — 後端檢查：單元（lib/quoteFlowDesigner.js）＋ API/權限/攻擊（真實 server.js 沙盒）＋ 開關關閉對基底位元級不變 ＋ 變異測試。
 * 用法：node scripts/check-quote-flow-designer.js            （全部；預設埠 3130，FD_PORT 可改）
 *       FD_ONLY=unit|api|flagoff   只跑某一區
 *       FD_MUT=1                   另外跑變異測試（把實作改壞，測試必須變紅；較久）
 *       FD_BASE_REF=<git ref>      「基底」分支（預設 feat/quote-flow-engine），用於開關關閉的位元級比對
 * 沙盒：臨時 _s_* 帳號與資料檔只存在於本次執行，結束以 PID 關閉伺服器並刪檔；不連任何遠端資料庫；工作樹已有 auth.json／data.json／稽核檔時拒絕執行。
 *
 * 單元區
 *   A1 開關解析   A2 輸入淨化（攻擊）   A3 預設流程通過全部檢查＋結構雜湊釘選   A4 保護底線（定向案例＋隨機變異 fuzz：接受的定義用獨立暴力取樣找不到違規；被拒的有可重現的反例）
 *   A5 11 條驗證各有正反案例   A6 模擬＝QF.plan＝舊規則 oracle（requiredPath＋BU/例外）   A7 邊界掃描   A8 總覽圖不變式   A9 白話檢視
 * API 區
 *   B1 權限（未登入／各角色；11 個端點）   B2 狀態／預設流程   B3 草稿 CRUD＋樂觀並行＋上限   B4 稽核   B5 攻擊輸入   B6 底線被後端拒絕   B7 模擬＝真實送簽（舊路徑 oracle）
 *   B8 草稿不影響任何簽核（含 QUOTE_FLOW_ENGINE=on）   B9 回應不含帳號／成本   B10 XSS 內容原樣存取（前端負責跳脫）
 * 開關區  off（預設）時 10 個端點對所有身分的回應 ＝ 基底伺服器對同一路徑的回應（位元級，扣掉路徑回聲）；on 時才有功能
 */
'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.FD_PORT || '3130', 10);
if ([3000, 3020, 3030].includes(PORT)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const ONLY = process.env.FD_ONLY || '';
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
const guardFiles = FILES.filter((f) => /(auth|data|audit.log).jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }

const LIBDIR = process.env.FD_LIBDIR || path.join(ROOT, 'lib');
const FD_FILE = process.env.FD_FILE || path.join(LIBDIR, 'quoteFlowDesigner.js');
const QF = require(path.join(ROOT, 'lib', 'quoteFlow'));
const QA = require(path.join(ROOT, 'lib', 'quoteApproval'));
const FD = require(FD_FILE);
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
const failures = [];
let quiet = false;   // 變異測試內部的失敗是「預期中的紅燈」，不印出
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));
const clone = (o) => JSON.parse(JSON.stringify(o));
const canon = QF.canon;

// 固定種子的亂數（可重現）
function rng(seed) { let a = seed >>> 0; return () => { a |= 0; a = (a + 0x6D2B79F5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }

// ───────────────────────── 共用資料 ─────────────────────────
const CFG = { buSettings: { ERP: { hasMgr2Step: false }, ITS: { hasMgr2Step: true }, MDM: { hasMgr2Step: true }, CRM: { hasMgr2Step: false }, 資安: { hasMgr2Step: false } }, exceptions: [{ username: 'ux', action: 'skip_mgr1' }] };
const legacy = () => clone(QF.legacyDefault(CFG));
const ENV = (over) => Object.assign({
  roster: { gm: ['g1', 'g2'], chairman: ['c1'], boardProxy: ['b1'] }, signerState: () => 'ok', chain: { owners: 4, noMgr1: 0, noMgr2: 0 }, activeDef: QF.legacyDefault(CFG),
}, over || {});
const checkOf = (rc, n) => rc.checks.find((c) => c.n === n);
const hasCode = (rc, code) => rc.checks.some((c) => c.issues.some((i) => i.code === code));
const route = (d, id) => d.routes.find((r) => r.id === id);

// ═══════════════════════════ 單元區 ═══════════════════════════
function runUnit() {
  console.log('[A 單元]');
  // A1 開關
  ok(FD.designerEnabled('1') && FD.designerEnabled(' 1 ') && FD.designerEnabled(1), 'A1 開關：1 開啟');
  ok(!FD.designerEnabled(undefined) && !FD.designerEnabled('') && !FD.designerEnabled('0') && !FD.designerEnabled('true') && !FD.designerEnabled('on') && !FD.designerEnabled('11') && !FD.designerEnabled(null), 'A1 開關：未設／0／true／on 一律關閉（只認 1）');

  // A2 淨化
  const S = FD.sanitizeDef;
  const bads = [
    [null, 'DEF_NOT_OBJECT'], ['str', 'DEF_NOT_OBJECT'], [[], 'DEF_NOT_OBJECT'], [7, 'DEF_NOT_OBJECT'],
  ];
  bads.forEach(([v, code]) => eq(S(v).code, code, 'A2 非物件被拒：' + JSON.stringify(v)));
  const base = legacy();
  const withKey = (k, v) => Object.assign(clone(base), { [k]: v });
  eq(S(withKey('evil', 1)).code, 'DEF_UNKNOWN_KEY', 'A2 頂層未知鍵被拒');
  eq(S(JSON.parse('{"__proto__":{"x":1},"schema":1}')).code, 'DEF_UNKNOWN_KEY', 'A2 __proto__ 鍵被拒');
  const n1 = clone(base); n1.nodes[0].constructor = 'x';
  eq(S(n1).code, 'DEF_UNKNOWN_KEY', 'A2 節點 constructor 鍵被拒');
  const r1 = clone(base); r1.routes[0].extra = 1;
  eq(S(r1).code, 'DEF_UNKNOWN_KEY', 'A2 路線未知鍵被拒');
  const m1 = clone(base); m1.modifiers[0].extra = 1;
  eq(S(m1).code, 'DEF_UNKNOWN_KEY', 'A2 調整規則未知鍵被拒');
  const t1 = clone(base); t1.tables.margin[0].extra = 1;
  eq(S(t1).code, 'DEF_UNKNOWN_KEY', 'A2 決策表列未知鍵被拒');
  let deep = {}; let cur = deep; for (let i = 0; i < 6000; i++) { cur.a = {}; cur = cur.a; }
  eq(S(Object.assign(clone(base), { tables: deep })).code, 'DEF_TOO_DEEP', 'A2 6000 層深巢狀被拒（不炸堆疊）');
  const big = clone(base); big.routes[0].name = 'x'.repeat(70000);
  eq(S(big).code, 'DEF_TOO_LARGE', 'A2 超過 64KB 被拒');
  eq(S(Object.assign(clone(base), { nodes: 'x' })).code, 'DEF_BAD_SHAPE', 'A2 nodes 不是陣列被拒');
  eq(S(Object.assign(clone(base), { routes: [1] })).code, 'DEF_BAD_SHAPE', 'A2 路線項目不是物件被拒');
  const goodS = S(base);
  ok(goodS.ok && goodS.def !== base && canon(goodS.def) === canon(base), 'A2 正常定義通過且回傳新複本');
  ok(({}).polluted === undefined, 'A2 原型未被污染');

  // A3 預設流程
  const lg = legacy();
  const rcL = FD.runChecks(lg, ENV());
  ok(rcL.ok && rcL.structural && rcL.checks.length === 11, 'A3 預設流程通過全部內容類檢查');
  eq(rcL.checks.map((c) => c.status), ['pass', 'pass', 'pass', 'pass', 'pass', 'pass', 'pass', 'pass', 'pass', 'pass', 'pass'], 'A3 預設流程 11 條全部通過');
  eq(QF.defHash(QF.buildLegacyDefault({ buSettings: {}, exceptions: [] })), QF.DEFAULT_FLOW_HASH, 'A3 預設流程結構雜湊＝釘選值');
  for (const cfg of [{ buSettings: {}, exceptions: [] }, CFG, { buSettings: { ERP: { hasMgr2Step: true }, ITS: { hasMgr2Step: true }, MDM: { hasMgr2Step: true }, CRM: { hasMgr2Step: true }, 資安: { hasMgr2Step: true } }, exceptions: [{ username: 'a', action: 'skip_mgr1' }, { username: 'b', action: 'skip_mgr1' }] }]) {
    const d = clone(QF.legacyDefault(cfg));
    const rc = FD.runChecks(d, ENV({ activeDef: QF.legacyDefault(cfg) }));
    ok(rc.ok && rc.floor.ok === true && rc.loosening.length === 0, 'A3 各種 BU／例外設定的預設流程都通過底線且與自己比較無放寬');
  }

  // A4 底線：定向案例
  const viol = (mut, name, shouldViolate) => {
    const d = legacy(); mut(d);
    const rc = FD.runChecks(d, ENV());
    const c10 = checkOf(rc, 10);
    if (shouldViolate) ok(rc.structural && c10.status === 'error' && !rc.ok && rc.blocking.some((b) => b.n === 10), 'A4 違反底線被擋：' + name, JSON.stringify(c10 && c10.issues.map((i) => i.text)));
    else ok(rc.structural && c10.status === 'pass', 'A4 合法編輯不被誤擋：' + name, JSON.stringify(c10 && c10.issues.map((i) => i.text)));
    return rc;
  };
  viol((d) => { route(d, 'r_l3').path = ['n_mgr1', 'n_gm']; }, '董事長路線拿掉董事長', true);
  viol((d) => { route(d, 'r_l3').when = { any: [{ cmp: ['amount', '>', '20000000'] }, { cmp: ['marginLevel', '==', 3] }] }; }, '1,000 萬門檻改成 2,000 萬', true);
  viol((d) => { route(d, 'r_l3').when = { any: [{ cmp: ['amount', '>', '10000000.01'] }, { cmp: ['marginLevel', '==', 3] }] }; }, '門檻差 0.01 元（10,000,000.01）', true);
  viol((d) => { route(d, 'r_board').when = { cmp: ['amount', '>', '60000000'] }; }, '5,000 萬門檻改成 6,000 萬', true);
  viol((d) => { route(d, 'r_board').path = ['n_mgr1', 'n_gm', 'n_chair'.replace('chair', 'chairman')]; }, '董事會路線改成董事長（>5,000 萬缺董事會）', true);
  viol((d) => { d.routes.splice(0, 1); }, '刪掉董事會路線', true);
  viol((d) => { d.routes.splice(0, 0, { id: 'r_x', name: '剛好 2,000 萬', when: { cmp: ['amount', '==', '20000000'] }, path: ['n_mgr1'] }); }, '單點例外：金額剛好等於 2,000 萬只要一級主管', true);
  viol((d) => { d.routes.splice(0, 0, { id: 'r_x', name: 'ERP 大單', when: { all: [{ cmp: ['ownerBuPrimary', '==', 'ERP'] }, { cmp: ['amount', '>', '30000000'] }] }, path: ['n_mgr1', 'n_gm'] }); }, '特定 BU 大單少董事長', true);
  viol((d) => { d.routes.splice(0, 0, { id: 'r_x', name: '鎖定分類', when: { cmp: ['category', '==', 'consult'] }, path: ['n_mgr1'] }); }, '顧問服務全走一級主管（大單缺董事長）', true);
  viol((d) => { d.modifiers.push({ id: 'm_rmc', name: '大單移除董事長', op: 'remove', target: 'n_chairman', when: { cmp: ['amount', '>', '10000000'] } }); }, '調整規則把董事長移掉', true);
  viol((d) => { d.routes = [{ id: 'r_a', name: '只看毛利層級3', when: { cmp: ['marginLevel', '==', 3] }, path: ['n_mgr1', 'n_gm', 'n_chairman'] }, { id: 'r_b', name: '只看毛利層級2', when: { cmp: ['marginLevel', '==', 2] }, path: ['n_mgr1', 'n_gm'] }, { id: 'r_default', name: '預設', when: true, path: ['n_mgr1'] }]; }, '完全沒有金額條件（大單一律只走毛利層級）', true);
  viol((d) => { route(d, 'r_l3').path = ['n_mgr1', 'n_gm', 'n_board']; }, '董事長路線改成董事會（>1,000 萬含董事會即可）', false);
  viol((d) => { route(d, 'r_default').path = ['n_mgr1', 'n_gm']; }, '預設路線加總經理（更嚴）', false);
  viol((d) => { route(d, 'r_l3').when = { any: [{ cmp: ['amount', '>', '5000000'] }, { cmp: ['marginLevel', '==', 3] }] }; }, '門檻降到 500 萬（更嚴）', false);
  viol((d) => { route(d, 'r_l3').when = { any: [{ cmp: ['amount', '>=', '10000000'] }, { cmp: ['marginLevel', '==', 3] }] }; }, '> 改成 >=（更嚴）', false);
  viol((d) => { d.routes.splice(2, 0, { id: 'r_y', name: 'MDM 加簽', when: { cmp: ['ownerBuPrimary', '==', 'MDM'] }, path: ['n_mgr1', 'n_gm'] }); }, '新增不碰大單的 MDM 路線', false);
  // A4 隨機變異 fuzz：接受的定義用獨立暴力取樣找不到違規；被拒的反例可用 QF.evaluate 重現
  {
    const rnd = rng(20261020);
    const ri = (n) => Math.floor(rnd() * n);
    const pick = (a) => a[ri(a.length)];
    const AMTS = ['5000000', '9999999.99', '10000000', '10000000.01', '12000000', '20000000', '49999999.99', '50000000', '50000000.01', '80000000'];
    const TIERS = ['n_mgr2', 'n_mgr1', 'n_gm', 'n_chairman', 'n_board'];
    const randCond = () => {
      const atoms = [];
      const k = 1 + ri(2);
      for (let i = 0; i < k; i++) {
        const t = ri(5);
        if (t === 0) atoms.push({ cmp: ['amount', pick(['>', '>=', '<', '<=', '==']), pick(AMTS)] });
        else if (t === 1) atoms.push({ cmp: ['marginLevel', pick(['==', '>=', '>', '<']), 1 + ri(3)] });
        else if (t === 2) atoms.push({ cmp: ['category', pick(['==', '!=']), pick(QA.ROWS.map((r) => r.key))] });
        else if (t === 3) atoms.push({ cmp: ['ownerBuPrimary', 'in', [pick(FD.BUS), pick(FD.BUS)]] });
        else atoms.push({ cmp: ['marginPct', pick(['>', '>=', '<']), pick(['5', '10', '15', '25.5'])] });
      }
      return atoms.length === 1 && rnd() < 0.5 ? atoms[0] : { [pick(['all', 'any'])]: atoms };
    };
    const randPath = () => { const p = TIERS.filter(() => rnd() < 0.55); return p.length ? p : ['n_mgr1']; };
    const mutate = (d) => {
      const n = 1 + ri(3);
      for (let i = 0; i < n; i++) {
        const t = ri(10);
        const rs = d.routes;
        if (t === 0) { const r = pick(rs); const lits = []; const w = (x) => { if (x && x.cmp && x.cmp[0] === 'amount') lits.push(x); else if (x && (x.all || x.any)) (x.all || x.any).forEach(w); }; w(r.when); if (lits.length) pick(lits).cmp[2] = pick(AMTS); }
        else if (t === 1) { const r = pick(rs); const lits = []; const w = (x) => { if (x && x.cmp && x.cmp[0] === 'amount') lits.push(x); else if (x && (x.all || x.any)) (x.all || x.any).forEach(w); }; w(r.when); if (lits.length) pick(lits).cmp[1] = pick(['>', '>=', '<', '<=', '==']); }
        else if (t === 2) { const r = pick(rs); if (r.path.length > 1) r.path.splice(ri(r.path.length), 1); }
        else if (t === 3) { const r = pick(rs); const id = pick(TIERS); if (!r.path.includes(id) && r.path.length < 6) r.path.splice(ri(r.path.length + 1), 0, id); }
        else if (t === 4 && rs.length > 2) { const a = ri(rs.length - 1), b = ri(rs.length - 1); const x = rs[a]; rs[a] = rs[b]; rs[b] = x; }
        else if (t === 5 && rs.length < 8) rs.splice(ri(rs.length), 0, { id: 'r_f' + ri(1e6), name: 'F' + i, when: randCond(), path: randPath() });
        else if (t === 6 && rs.length > 2) rs.splice(ri(rs.length - 1), 1);
        else if (t === 7) { const m = d.modifiers.find((x) => x.id === 'm_mgr2'); if (m) m.when.all[0].cmp[2] = FD.BUS.filter(() => rnd() < 0.5).sort(); }
        else if (t === 8) rs[rs.length - 1].path = randPath();
        else if (t === 9) d.modifiers.push({ id: 'm_f' + ri(1e6), name: 'M', op: pick(['insert_before', 'insert_after', 'remove', 'replace']), target: pick(TIERS), node: pick(TIERS), when: randCond() });
      }
      return d;
    };
    const DEFINITELY_LITS = (d) => { const s = new Set(); const w = (x) => { if (x && x.cmp && x.cmp[0] === 'amount') s.add(Number(QF.parseMoneyCents(x.cmp[2]))); else if (x && (x.all || x.any)) (x.all || x.any).forEach(w); }; d.routes.forEach((r) => w(r.when)); d.modifiers.forEach((m) => w(m.when)); return [...s]; };
    let accepted = 0, rejected = 0, invalid = 0, bruteViolOnAccepted = 0, badCounterexample = 0;
    const N = parseInt(process.env.FD_FUZZ_N || '260', 10);
    for (let it = 0; it < N; it++) {
      const d = mutate(legacy());
      const v = QF.validateFlowDef(d);
      if (!v.ok) { invalid++; continue; }
      const an = FD.analyze(d, null);
      const fz = frozenFor(d);
      if (an.violationCount === 0) {
        accepted++;
        // 獨立暴力取樣：字面值 ±1 分 與 全域隨機
        const lits = DEFINITELY_LITS(d);
        for (let k = 0; k < 2500 && !bruteViolOnAccepted; k++) {
          const useLit = lits.length && rnd() < 0.5;
          const rev = useLit ? Math.max(1, pick(lits) + pick([-1, 0, 1])) : (rnd() < 0.5 ? Math.max(1, Math.floor(Math.pow(10, 2 + rnd() * 11))) : pick([999999999, 1000000000, 1000000001, 4999999999, 5000000000, 5000000001]) + ri(3) - 1);
          const cat = pick(QA.ROWS).key;
          const gp = Math.floor(rev * (rnd() * 1.2 - 0.1));
          const bu = pick(FD.BUS.concat(['ZZ']));
          const m2 = pick(['none', 'plain', 'ux']);
          const io = { mgr2: () => (m2 === 'none' ? null : m2 === 'plain' ? 'p' : 'ux'), mgr1: () => 'q', roster: { gm: ['g'], chairman: ['c'] }, boardProxy: () => ['b'] };
          let ev; try { ev = QF.evaluate(fz, { owner: '', ownerBu: bu, category: cat, classKeys: [], revenueCents: rev, gpCents: gp }, io); } catch (_) { continue; }
          if (FD.floorNeed(rev, ev.actualTiers)) bruteViolOnAccepted++;
        }
      } else {
        rejected++;
        an.violations.forEach((vi) => {
          // 反例必須能用 QF.evaluate 直接重現（同樣的事實 → 同樣違反底線）
          const io = { mgr2: () => (vi.mgr2 === 'none' ? null : vi.mgr2 === 'plain' ? '__sim_mgr2' : vi.mgr2), mgr1: () => '__sim_mgr1', roster: { gm: ['__sim_gm'], chairman: ['__sim_ch'] }, boardProxy: () => ['__sim_bp'] };
          let okRepro = false;
          try { const ev = QF.evaluate(fz, { owner: vi.owner, ownerBu: vi.ownerBu, category: vi.category, classKeys: vi.classKeys, revenueCents: vi.revenueCents, gpCents: vi.gpCents }, io); okRepro = !!FD.floorNeed(vi.revenueCents, ev.actualTiers) && ev.actualTiers.join() === vi.tiers.join(); } catch (_) { okRepro = false; }
          if (!okRepro) badCounterexample++;
        });
      }
    }
    console.log(`  fuzz：${N} 份變異定義 → 結構無效 ${invalid}、接受 ${accepted}、因違反底線被拒 ${rejected}`);
    ok(accepted >= 20 && rejected >= 20, 'A4 fuzz：接受與被拒的樣本都夠多（' + accepted + '／' + rejected + '）');
    eq(bruteViolOnAccepted, 0, 'A4 fuzz：被接受的定義，獨立暴力取樣（字面值±1分＋隨機）找不到任何違規');
    eq(badCounterexample, 0, 'A4 fuzz：被拒的定義，反例都能用 QF.evaluate 重現');
  }
  function frozenFor(d) { return FD.frozenCopy(d); }

  // A5 11 條驗證：各有正反案例
  {
    const rcOK = FD.runChecks(legacy(), ENV());
    // 1
    let d = legacy(); route(d, 'r_default').when = { cmp: ['amount', '>', '1'] };
    let rc = FD.runChecks(d, ENV()); ok(checkOf(rc, 1).status === 'error' || !rc.structural, 'A5-1 最後一條不是預設 → 錯誤');
    ok(hasCode(rc, 'FLOW_NO_DEFAULT'), 'A5-1 錯誤碼 FLOW_NO_DEFAULT');
    d = legacy(); route(d, 'r_default').path = []; rc = FD.runChecks(d, ENV()); ok(hasCode(rc, 'FLOW_EMPTY_PATH') && !rc.ok, 'A5-1 預設路線空 → FLOW_EMPTY_PATH 並擋');
    eq(checkOf(rcOK, 1).status, 'pass', 'A5-1 預設流程通過');
    // 2
    d = legacy(); route(d, 'r_l3').path = ['n_mgr1', 'n_gm', 'n_gm', 'n_chairman']; rc = FD.runChecks(d, ENV()); ok(hasCode(rc, 'FLOW_DUP_NODE') && !rc.ok, 'A5-2 路線內關卡重複 → FLOW_DUP_NODE');
    d = legacy(); d.limits.maxSteps = 2; rc = FD.runChecks(d, ENV()); ok(hasCode(rc, 'FLOW_TOO_LONG') && !rc.ok, 'A5-2 超過關卡上限 → FLOW_TOO_LONG');
    d = legacy(); route(d, 'r_default').path = ['n_gm', 'n_mgr1']; rc = FD.runChecks(d, ENV()); ok(checkOf(rc, 2).status === 'warn' && rc.ok, 'A5-2 順序由高到低 → 警告（不擋）');
    d = legacy(); route(d, 'r_default').path = ['n_mgr2', 'n_mgr1']; rc = FD.runChecks(d, ENV()); ok(hasCode(rc, 'FLOW_INVALID_RESULT') && !rc.ok, 'A5-2 路線本來就有二級主管，BU 調整又插入 → 送簽會被擋（FLOW_INVALID_RESULT）');
    d = legacy(); route(d, 'r_default').path = ['n_ghost']; rc = FD.runChecks(d, ENV()); ok(hasCode(rc, 'FLOW_NODE_MISSING') && !rc.ok, 'A5-2 引用不存在的節點 → FLOW_NODE_MISSING');
    // 3
    d = legacy(); d.modifiers = [d.modifiers[1], d.modifiers[0]]; rc = FD.runChecks(d, ENV()); ok(checkOf(rc, 3).status === 'error' && !rc.ok, 'A5-3 applied 引用排在後面的調整 → 第 3 條錯誤');
    eq(checkOf(rcOK, 3).status, 'pass', 'A5-3 預設流程通過');
    // 4 / 5
    rc = FD.runChecks(legacy(), ENV({ roster: { gm: [], chairman: ['c1'], boardProxy: ['b1'] } })); ok(checkOf(rc, 4).status === 'error' && hasCode(rc, 'FLOW_NO_SIGNER') && rc.ok, 'A5-4 總經理名單空 → FLOW_NO_SIGNER（環境狀態：回報但不擋草稿）');
    rc = FD.runChecks(legacy(), ENV({ signerState: (un) => (un === 'g1' || un === 'g2' ? 'readonly' : 'ok') })); ok(checkOf(rc, 5).status === 'error' && hasCode(rc, 'FLOW_ONLY_READONLY') && rc.ok, 'A5-5 總經理名單全是唯讀 → FLOW_ONLY_READONLY');
    rc = FD.runChecks(legacy(), ENV({ signerState: (un) => (un === 'c1' ? 'inactive' : 'ok') })); ok(checkOf(rc, 4).status === 'error', 'A5-4 董事長全停用 → 錯誤');
    d = legacy(); route(d, 'r_board').path = ['n_mgr1', 'n_gm', 'n_board']; rc = FD.runChecks(d, ENV({ roster: { gm: ['g1'], chairman: ['c1'], boardProxy: [] } })); ok(checkOf(rc, 4).status === 'error', 'A5-4 董事會代核名單空 → 錯誤');
    rc = FD.runChecks(legacy(), ENV({ chain: { owners: 5, noMgr1: 2, noMgr2: 0 } })); ok(checkOf(rc, 4).status === 'warn' && rc.ok, 'A5-4 有業務解析不到一級主管 → 警告');
    eq(checkOf(rcOK, 5).status, 'pass', 'A5-5 預設流程通過');
    // 6
    const badConds = [{ cmp: ['evil', '>', '1'] }, { cmp: ['amount', 'contains', '1'] }, { cmp: ['amount', '>', 100] }, { cmp: ['amount', '>', '1e9'] }, { cmp: ['marginLevel', '==', 4] }, { cmp: ['category', '==', 'nope'] }, { eval: 'process.exit()' }, { all: [] }, { cmp: ['amount', '>'] }, { cmp: ['ownerBuPrimary', 'in', 'ERP'] }];
    badConds.forEach((c, i) => { d = legacy(); route(d, 'r_l2').when = c; rc = FD.runChecks(d, ENV()); ok(!rc.ok && checkOf(rc, 6).status === 'error', 'A5-6 壞條件被擋 #' + i + ' ' + JSON.stringify(c)); });
    let nest = true; for (let i = 0; i < 12; i++) nest = { not: nest }; d = legacy(); route(d, 'r_l2').when = nest; rc = FD.runChecks(d, ENV()); ok(!rc.ok && checkOf(rc, 6).status === 'error', 'A5-6 條件巢狀過深被擋');
    d = legacy(); d.modifiers[0].when.all[0].cmp[2] = []; rc = FD.runChecks(d, ENV()); ok(checkOf(rc, 6).status === 'warn' && rc.ok, 'A5-6 調整規則沒選任何 BU → 警告');
    // 7
    d = legacy(); d.routes.splice(1, 0, { id: 'r_dup', name: '重複', when: { cmp: ['amount', '>', '50000000'] }, path: ['n_mgr1'] }); rc = FD.runChecks(d, ENV()); ok(checkOf(rc, 7).status === 'warn' && checkOf(rc, 7).issues.some((i) => i.loc && i.loc.lane === 'r_dup'), 'A5-7 永遠不會被選中的路線 → 警告並指到該路線');
    eq(checkOf(rcOK, 7).status, 'pass', 'A5-7 預設流程通過');
    // 8
    rc = FD.runChecks(legacy(), ENV({ roster: { gm: ['same'], chairman: ['same'], boardProxy: ['b1'] } })); ok(checkOf(rc, 8).status === 'warn' && hasCode(rc, 'SIGNER_OVERLAP') && rc.ok, 'A5-8 總經理與董事長名冊只有同一人 → SIGNER_OVERLAP 警告');
    eq(checkOf(rcOK, 8).status, 'pass', 'A5-8 預設流程通過');
    // 9
    d = legacy(); route(d, 'r_l2').path = ['n_mgr1']; rc = FD.runChecks(d, ENV()); ok(checkOf(rc, 9).status === 'warn' && rc.loosening.length > 0 && rc.ok, 'A5-9 總經理路線拿掉總經理 → 放寬警告');
    ok(rc.loosening[0].text.includes('原需') && rc.loosening[0].text.includes('新版只需'), 'A5-9 放寬說明文字含原需／新版只需');
    eq(checkOf(rcOK, 9).status, 'pass', 'A5-9 與自己比較不放寬');
    d = legacy(); route(d, 'r_default').path = ['n_mgr1', 'n_gm']; rc = FD.runChecks(d, ENV()); eq(checkOf(rc, 9).status, 'pass', 'A5-9 變嚴格不算放寬');
    // 10 已於 A4；11
    const lblCases = [['<svg onload=a()>', 'HTML'], ['a"b', '雙引號'], ["a'b", '單引號'], ['a&b', '&'], ['', '空'], ['x'.repeat(15), '過長'], ['a\u0007b', '控制字元']];
    lblCases.forEach(([lb, nm]) => { d = legacy(); d.nodes.find((n) => n.tier === 'gm').label = lb; rc = FD.runChecks(d, ENV()); ok(!rc.ok && checkOf(rc, 11).status === 'error', 'A5-11 不合法的節點名稱被擋：' + nm); });
    d = legacy(); d.nodes.find((n) => n.tier === 'board').label = '改名'; rc = FD.runChecks(d, ENV()); ok(!rc.ok && hasCode(rc, 'FLOW_LOCKED_EDIT'), 'A5-11 鎖定的代核節點改名被擋');
    d = legacy(); d.nodes.find((n) => n.tier === 'board').locked = false; rc = FD.runChecks(d, ENV()); ok(!rc.ok, 'A5-11 取消代核節點鎖定被擋（M1 驗證）');
    d = legacy(); d.nodes.find((n) => n.tier === 'gm').color = 'rainbow'; rc = FD.runChecks(d, ENV()); ok(!rc.ok && checkOf(rc, 11).status === 'error', 'A5-11 不在清單的節點色被擋');
    d = legacy(); d.nodes.find((n) => n.tier === 'board').color = 'teal'; rc = FD.runChecks(d, ENV()); ok(!rc.ok, 'A5-11 代核節點不可改色');
    d = legacy(); d.nodes.find((n) => n.tier === 'gm').color = 'amber'; d.nodes.find((n) => n.tier === 'gm').label = '總經理核決'; rc = FD.runChecks(d, ENV()); ok(rc.ok, 'A5-11 合法改名＋節點色通過');
    d = legacy(); route(d, 'r_l2').name = '<b>x</b> & "y"'; rc = FD.runChecks(d, ENV()); ok(rc.ok, 'A5-11 路線名稱含 HTML 字元允許（前端以純文字顯示）');
    d = legacy(); route(d, 'r_l2').name = 'x'.repeat(17); rc = FD.runChecks(d, ENV()); ok(!rc.ok, 'A5-11 路線名稱過長被擋');
    d = legacy(); route(d, 'r_l2').name = ' '; rc = FD.runChecks(d, ENV()); ok(!rc.ok, 'A5-11 路線名稱全空白被擋');
    d = legacy(); d.nodes[0].tier = 'accountant'; rc = FD.runChecks(d, ENV()); ok(!rc.ok && !rc.structural, 'A5 不可新增 tier 字串（M1 驗證）');
  }

  // A6 模擬＝QF.plan＝舊規則 oracle
  {
    const rnd = rng(77);
    const ri = (n) => Math.floor(rnd() * n);
    let n = 0, mism = 0;
    const flowCfg = CFG;
    const oracle = (cat, rev, gp, bu, noMgr2, exc) => {
      const rp = QA.requiredPath(cat, gp, rev);
      let t = rp.tiers.slice();
      const hasM2 = ['ITS', 'MDM'].includes(bu) && !noMgr2;
      if (hasM2) { t.splice(t.indexOf('mgr1'), 0, 'mgr2'); if (exc) t.splice(t.indexOf('mgr1'), 1); }
      return t;
    };
    const def = legacy();
    for (let i = 0; i < 1200; i++) {
      const cat = QA.ROWS[ri(QA.ROWS.length)].key;
      const rev = [100, 80000000, 100000000, 999999999, 1000000000, 1000000001, 4999999999, 5000000000, 5000000001, 123456789012][ri(10)] + ri(3);
      const pct = String((ri(401) - 20) / 10);   // -2.0 ~ 38.0
      const bu = FD.BUS[ri(5)];
      const noMgr2 = rnd() < 0.3, exc = rnd() < 0.3;
      const amount = FD.centsToYuanStr(rev);
      const parsed = FD.parseSimInput({ amount, marginPct: pct, category: cat, ownerBu: bu, noMgr2, mgr2Exception: exc }, flowCfg);
      if (!parsed.ok) { ok(false, 'A6 模擬輸入應合法 ' + JSON.stringify(parsed)); continue; }
      const out = FD.simulate(def, parsed, ENV(), null);
      const want = oracle(cat, parsed.input.revenueCents, parsed.input.gpCents, bu, noMgr2, exc && !noMgr2);
      n++;
      if (out.tiers.join() !== want.join()) { mism++; if (mism <= 3) console.log('  A6 不一致', amount, pct, cat, bu, noMgr2, exc, out.tiers.join(), want.join()); }
      // 與 QF.plan 直接比
      const io = { owner: '__sim_owner', costBy: null, mgr2: () => (noMgr2 ? null : (exc ? 'ux' : '__sim_mgr2')), mgr1: () => '__sim_mgr1', roster: { gm: ['g1', 'g2'], chairman: ['c1'] }, boardProxy: () => ['b1'], canApprove: () => true };
      const p = QF.plan(QF.legacyDefault(flowCfg), { owner: '__sim_owner', ownerBu: bu, category: cat, classKeys: [], revenueCents: parsed.input.revenueCents, gpCents: parsed.input.gpCents }, io);
      if (p.actualTiers.join() !== out.tiers.join()) mism++;
    }
    eq(mism, 0, `A6 模擬結果＝舊規則 oracle＝QF.plan（${n} 組隨機）`);
    // 層級輸入：選層級 N 時，引擎算出的毛利層級確實是 N
    let lvBad = 0;
    for (const row of QA.ROWS) for (const lv of [1, 2, 3]) for (const amt of ['500000', '12000000', '99999999.99']) {
      const parsed = FD.parseSimInput({ amount: amt, marginLevel: lv, category: row.key, ownerBu: 'ERP' }, flowCfg);
      if (!parsed.ok) { if (!(parsed.code === 'SIM_LEVEL_IMPOSSIBLE')) lvBad++; continue; }
      const got = QF.marginLevelOf(def.tables.margin, row.key, BigInt(parsed.input.gpCents), BigInt(parsed.input.revenueCents));
      if (!row.fixedLevel && got !== lv) lvBad++;
    }
    eq(lvBad, 0, 'A6 以「毛利層級」輸入時，引擎算出的層級與輸入一致（每個類別×層級×3 種金額）');
    const pi0 = FD.parseSimInput({ amount: '500000', marginLevel: 1, category: 'ot', ownerBu: 'ERP' }, flowCfg);
    ok(!pi0.ok && pi0.code === 'SIM_LEVEL_IMPOSSIBLE', 'A6 OT 沒有層級 1 → 明確拒絕');
    // 輸入攻擊
    const badIn = [null, 'x', [], { amount: 5 }, { amount: '-1' }, { amount: '1e3' }, { amount: '0' }, { amount: '1.234' }, { amount: '9999999999999999999' }, { amount: '1', marginLevel: 9 }, { amount: '1', marginPct: 'abc' }, { amount: '1', category: 'x' }, { amount: '1', classKeys: ['zz'] }, { amount: '1', ownerBu: '' }, { amount: '1', extra: 1 }, { amount: '1', ownerBu: 'a\u0001' }, { amount: '1', __proto__x: 1 }];
    badIn.forEach((b, i) => ok(!FD.parseSimInput(b, flowCfg).ok, 'A6 壞輸入被拒 #' + i + ' ' + JSON.stringify(b)));
    // 比較模式與放寬
    const loose = legacy(); route(loose, 'r_l2').path = ['n_mgr1'];
    const pi = FD.parseSimInput({ amount: '3000000', marginLevel: 2, category: 'consult', ownerBu: 'ERP' }, flowCfg);
    const outC = FD.simulate(loose, pi, ENV(), QF.legacyDefault(flowCfg));
    ok(outC.against && outC.against.same === false && outC.against.loose === true, 'A6 比較模式：草稿拿掉總經理 → 標示放寬');
    const outS = FD.simulate(legacy(), pi, ENV(), QF.legacyDefault(flowCfg));
    ok(outS.against.same === true && outS.against.loose === false, 'A6 比較模式：相同 → same');
    ok(outC.steps.length === 1 && outC.against.steps.length === 2 && outC.against.steps.some((s) => s.removedInDraft), 'A6 比較模式：生效版多出的關卡標示為草稿拿掉');
    // 阻擋碼
    const noGm = FD.simulate(legacy(), pi, ENV({ roster: { gm: [], chairman: ['c1'], boardProxy: ['b1'] } }), null);
    ok(noGm.blockers.some((b) => b.code === 'NO_GM'), 'A6 名冊沒有總經理 → 模擬回報 NO_GM');
    // 追蹤：命中後的路線標為不再檢查
    ok(outS.trace.filter((t) => t.kind === 'route').some((t) => t.evaluated === false) || outS.route.isDefault, 'A6 追蹤：命中後的路線標為未評估');
    const namedEnv = ENV({ roster: { gm: ['zz_gm_person'], chairman: ['zz_ch_person'], boardProxy: ['zz_bp_person'] } });
    const outN = FD.simulate(legacy(), pi, namedEnv, null);
    ok(!/zz_(gm|ch|bp)_person/.test(JSON.stringify(outN)) && !/zz_/.test(JSON.stringify(FD.sweep(legacy(), { amount: '1' }, namedEnv, null, CFG))), 'A6 模擬／掃描回應不含名冊帳號');
  }

  // A7 邊界掃描
  {
    const def = legacy();
    const sw = FD.sweep(def, { amount: '1', marginLevel: 1, category: 'consult', ownerBu: 'ERP' }, ENV(), QF.legacyDefault(CFG), CFG);
    ok(sw.ok && sw.rows.length >= 6, 'A7 掃描有列');
    const rowAt = (a) => sw.rows.find((r) => r.amount === a);
    ok(rowAt('10000000') && rowAt('10000000').tiers.join() === 'mgr1', 'A7 剛好 1,000 萬、毛利一般 → 只要一級主管（嚴格大於）');
    ok(rowAt('10000000.01') && rowAt('10000000.01').tiers.join() === 'mgr1,gm,chairman', 'A7 1,000 萬 + 0.01 → 董事長');
    ok(rowAt('50000000') && rowAt('50000000').tiers.join() === 'mgr1,gm,chairman', 'A7 剛好 5,000 萬 → 董事長（不是董事會）');
    ok(rowAt('50000000.01') && rowAt('50000000.01').tiers.join() === 'mgr1,gm,board', 'A7 5,000 萬 + 0.01 → 董事會');
    ok(sw.rows.every((r) => r.same === true && r.floorBad === false), 'A7 預設流程對自己：全部相同、無違反底線');
    eq(sw.summary.floorViolations, 0, 'A7 摘要：違反底線 0');
    const d2 = legacy(); d2.routes.splice(1, 0, { id: 'r_dead', name: '死路線', when: { cmp: ['amount', '>', '50000000'] }, path: ['n_mgr1'] });
    const sw2 = FD.sweep(d2, { amount: '1', marginLevel: 1, category: 'consult', ownerBu: 'ERP' }, ENV(), QF.legacyDefault(CFG), CFG);
    ok(sw2.summary.unreachableRoutes.some((r) => r.id === 'r_dead'), 'A7 掃描回報永遠不會生效的路線');
    const d3 = legacy(); route(d3, 'r_l2').path = ['n_mgr1'];
    const sw3 = FD.sweep(d3, { amount: '1', marginLevel: 2, category: 'consult', ownerBu: 'ERP' }, ENV(), QF.legacyDefault(CFG), CFG);
    ok(sw3.summary.loosened > 0 && sw3.summary.differing > 0 && sw3.rows.some((r) => r.loose === true), 'A7 掃描回報放寬與不同的組數');
    const sw4 = FD.sweep(legacy(), { amount: '1', marginLevel: 1, category: 'consult', ownerBu: 'ERP' }, ENV({ roster: { gm: [], chairman: [], boardProxy: [] } }), QF.legacyDefault(CFG), CFG);
    ok(sw4.summary.blockedSamples > 0 && sw4.summary.blockCodes.NO_GM > 0, 'A7 掃描回報無法解析簽核人的組數');
    ok(!FD.sweep(legacy(), { marginLevel: 9 }, ENV(), null, CFG).ok, 'A7 掃描的壞輸入被拒');
  }

  // A8 總覽圖不變式
  {
    const checkGraph = (g, name) => {
      const ids = new Set(g.nodes.map((n) => n.id));
      ok(ids.size === g.nodes.length, name + '：節點 id 不重複');
      ok(g.edges.every((e) => ids.has(e.from) && ids.has(e.to)), name + '：每條連線的兩端都是存在的節點');
      ok(g.nodes.every((n) => [n.x, n.y, n.w, n.h].every((v) => Number.isFinite(v)) && n.w > 0 && n.h > 0), name + '：座標與尺寸皆為有限數');
      let overlap = 0;
      for (let i = 0; i < g.nodes.length; i++) for (let j = i + 1; j < g.nodes.length; j++) {
        const a = g.nodes[i], b = g.nodes[j];
        if (Math.abs(a.x - b.x) * 2 < (a.w + b.w) - 1 && Math.abs(a.y - b.y) * 2 < (a.h + b.h) - 1) overlap++;
      }
      eq(overlap, 0, name + '：節點不重疊');
      ok(g.nodes.every((n) => n.x - n.w / 2 >= g.bounds.x0 && n.x + n.w / 2 <= g.bounds.x1 && n.y - n.h / 2 >= g.bounds.y0 && n.y + n.h / 2 <= g.bounds.y1), name + '：全部節點在 bounds 內');
      // 每個節點都能從 start 走到（連通），且 end 可達
      const adj = {}; g.edges.forEach((e) => { (adj[e.from] = adj[e.from] || []).push(e.to); });
      const seen = new Set(['start']); const st = ['start'];
      while (st.length) { const x = st.pop(); (adj[x] || []).forEach((y) => { if (!seen.has(y)) { seen.add(y); st.push(y); } }); }
      ok(g.nodes.filter((n) => n.kind !== 'mod').every((n) => seen.has(n.id)), name + '：除備註卡外每個節點都從「送簽」可達');
      ok(seen.has('end'), name + '：結束可達');
    };
    const def = legacy();
    const g = FD.buildOverview(def, null);
    checkGraph(g, 'A8 預設流程');
    eq(g.nodes.filter((n) => n.kind === 'decision').length, 3, 'A8 判斷菱形數＝非預設路線數');
    eq(g.nodes.filter((n) => n.kind === 'mod').length, 2, 'A8 調整規則備註卡數');
    eq(canon(FD.buildOverview(def, null)), canon(g), 'A8 排版是確定性的（同輸入同輸出）');
    // 模擬高亮：命中路線的關卡序列＝模擬結果
    for (const [amt, lv, bu] of [['800000', 1, 'ERP'], ['3000000', 2, 'ITS'], ['12000000', 1, 'MDM'], ['60000000', 1, 'ERP'], ['500000', 3, 'ITS']]) {
      const parsed = FD.parseSimInput({ amount: amt, marginLevel: lv, category: 'consult', ownerBu: bu }, CFG);
      const sim = FD.simulate(def, parsed, ENV(), null);
      const gg = FD.buildOverview(def, sim);
      checkGraph(gg, `A8 模擬 ${amt}/${lv}/${bu}`);
      const hitNodes = gg.nodes.filter((n) => n.kind === 'approval' && n.hit).sort((a, b) => a.x - b.x).map((n) => n.tier);
      eq(hitNodes, sim.tiers, `A8 高亮的關卡序列＝模擬結果（${amt}/${lv}/${bu}）`);
      const hitEdges = gg.edges.filter((e) => e.hit);
      ok(hitEdges.length >= sim.tiers.length, `A8 命中路徑有高亮連線（${amt}/${lv}/${bu}）`);
      eq(gg.nodes.filter((n) => n.kind === 'approval' && n.inserted).map((n) => n.tier), sim.steps.filter((s) => s.inserted).map((s) => s.tier), `A8 調整插入的關卡標記一致（${amt}/${lv}/${bu}）`);
    }
    // 單一預設路線
    const d1 = legacy(); d1.routes = [d1.routes[3]];
    const g1 = FD.buildOverview(d1, null);
    checkGraph(g1, 'A8 只有預設路線');
    eq(g1.nodes.filter((n) => n.kind === 'decision').length, 0, 'A8 只有預設路線時沒有判斷菱形');
    // 大流程
    const dBig = legacy(); for (let i = 0; i < 10; i++) dBig.routes.splice(1, 0, { id: 'r_b' + i, name: 'B' + i, when: { cmp: ['amount', '>', String(1000000 * (i + 1))] }, path: ['n_mgr1', 'n_gm'] });
    checkGraph(FD.buildOverview(dBig, null), 'A8 14 條路線');
  }

  // A8b 對抗性大定義：32 條路線、每條 20 個金額／毛利率條件、32 個調整規則——驗證與掃描的耗時有上限
  {
    const d = legacy();
    d.routes = [];
    for (let i = 0; i < 31; i++) d.routes.push({ id: 'r_a' + i, name: 'A' + i, when: { any: Array.from({ length: 20 }, (_, k) => (k % 2 ? { cmp: ['amount', '>', String(1000 + i * 37 + k)] } : { cmp: ['marginPct', '>=', String((i + k) % 90) + '.' + (k % 9)] })) }, path: ['n_mgr1', 'n_gm', 'n_chairman'] });
    d.routes.push({ id: 'r_default', name: '預設', when: true, path: ['n_mgr1', 'n_gm', 'n_chairman'] });
    const t0 = Date.now();
    const rc = FD.runChecks(d, ENV());
    const ms = Date.now() - t0;
    ok(rc.structural && ms < 8000, 'A8b 對抗性大定義（31 條路線×20 條件）驗證耗時有上限（' + ms + 'ms，取樣 ' + rc.samples + '，截斷 ' + rc.truncated + '）');
    ok(rc.samples <= FD.LIMITS.sampleCap, 'A8b 取樣數不超過上限');
    const t1 = Date.now();
    const sw = FD.sweep(d, { amount: '1', marginLevel: 1, category: 'consult', ownerBu: 'ERP' }, ENV(), QF.legacyDefault(CFG), CFG);
    const ms2 = Date.now() - t1;
    ok(sw.ok && ms2 < 20000, 'A8b 對抗性大定義的邊界掃描耗時有上限（' + ms2 + 'ms）');
  }

  // A9 白話檢視
  {
    const v = FD.viewOf(legacy());
    eq(v.routes.map((r) => r.isDefault), [false, false, false, true], 'A9 檢視：預設路線標記');
    ok(v.routes[0].whenText.includes('5,000 萬') && v.routes[1].whenText.includes('1,000 萬'), 'A9 檢視：條件白話含金額');
    eq(v.modifiers.map((m) => m.kind), ['bu', 'skip'], 'A9 檢視：調整規則分類（bu／skip）');
    eq(v.modifiers[0].bus, ['ITS', 'MDM'], 'A9 檢視：BU 清單');
    const d = legacy(); d.modifiers[1].op = 'replace'; d.modifiers[1].node = 'n_gm';
    eq(FD.viewOf(d).modifiers[1].kind, 'other', 'A9 檢視：不符合已知樣式的調整規則 → other（前端唯讀顯示）');
  }
}

// ═══════════════════════════ API 區 ═══════════════════════════
const PW = 'Fd1' + crypto.randomBytes(9).toString('hex') + '!Z';   // 只在記憶體
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false }, extra || {});
const USERS = () => [
  mkUser('_s_admin', 'admin'), mkUser('_s_own', 'user', { bu: ['ITS'], supervisor: '_s_m2' }), mkUser('_s_own2', 'user', { bu: ['ERP'], supervisor: '_s_m1' }),
  mkUser('_s_m2', 'manager2', { bu: ['ITS'], supervisor: '_s_m1' }), mkUser('_s_m1', 'manager1', { bu: ['ITS', 'ERP'] }),
  mkUser('_s_gm', 'executive'), mkUser('_s_ch', 'executive'), mkUser('_s_sec', 'secretary'),
];
const PRODUCTS = { PS: { cls: 'software', costBySales: true }, CS: { cls: 'consult', costBySales: true }, HW: { cls: 'hardware', costBySales: true } };
const mkQ = (id, owner, product, price, cost) => ({ id, quoteNo: 'SB-' + id, company: '沙盒公司', projectName: '沙盒專案', owner, status: 'draft', createdAt: nowIso, updatedAt: nowIso, quoteDate: nowIso.slice(0, 10), validUntil: '2099-12-31', products: [product], items: [{ lid: id + '-l1', desc: '服務', unit: '式', qty: 1, unitPrice: price, cost }], discountType: 'none', discountValue: 0, approval: null });
function seedQuotes() {
  const qs = [];
  const combos = [
    ['CS', 800000, 400000], ['CS', 800000, 680000], ['CS', 800000, 760000], ['PS', 800000, 400000], ['PS', 800000, 740000], ['PS', 800000, 780000],
    ['HW', 12000000, 6000000], ['CS', 12000000, 9000000], ['PS', 60000000, 30000000], ['CS', 60000000, 50000000], ['CS', 10000000, 5000000], ['CS', 10000000.01 > 0 ? 10000001 : 0, 5000000],
  ];
  combos.forEach(([p, price, cost], i) => {
    qs.push(mkQ('q' + (2 * i), '_s_own', p, price, cost));    // ITS：有二級主管
    qs.push(mkQ('q' + (2 * i + 1), '_s_own2', p, price, cost));   // ERP：沒有
  });
  return qs;
}
let DATADIR = ROOT;   // 比對基底時，資料檔寫在基底資料夾
function seed(extra) {
  fs.writeFileSync(path.join(DATADIR, 'auth.json'), JSON.stringify({ users: USERS() }, null, 2));
  const d = {
    contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [],
    quotations: seedQuotes(),
    quoteApproval: { roster: { gm: ['_s_gm'], chairman: ['_s_ch'], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: PRODUCTS },
    approvalFlowConfig: { version: 1, buSettings: { ITS: { hasMgr2Step: true } }, exceptions: [] },
  };
  Object.assign(d, extra || {});
  fs.writeFileSync(path.join(DATADIR, 'data.json'), JSON.stringify(d, null, 2));
}
const cleanup = () => { FILES.forEach((f) => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); if (DATADIR !== ROOT) FILES.forEach((f) => { try { fs.unlinkSync(path.join(DATADIR, path.basename(f))); } catch { /* 不存在 */ } }); };
const readData = () => JSON.parse(fs.readFileSync(path.join(DATADIR, 'data.json'), 'utf8'));

let child = null, childPid = null, jar = {};
const allPids = [];
async function startServer(env, opts) {
  opts = opts || {};
  const e = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE']) delete e[k];
  Object.assign(e, env || {});
  const cwd = opts.cwd || ROOT;
  child = spawn(process.execPath, [opts.file || path.join(cwd, 'server.js')], { cwd, env: e, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid; allPids.push(childPid);
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 100; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) { jar = {}; return; } } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-600));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-600));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 30 && child.exitCode === null; i++) await sleep(100);
  child = null;
  await sleep(300);
}
async function http(who, method, url, body, raw) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : (raw ? body : JSON.stringify(body)), redirect: 'manual' });
  const text = await res.text();
  let json = null; try { json = JSON.parse(text); } catch { /* 非 JSON */ }
  return { status: res.status, json, text, ctype: res.headers.get('content-type') || '', cache: res.headers.get('cache-control') || '', setCookie: res.headers.getSetCookie ? res.headers.getSetCookie() : [] };
}
async function login(who) {
  const r = await http(null, 'POST', '/api/login', { username: who, password: PW });
  if (r.status === 200) jar[who] = r.setCookie.map((c) => c.split(';')[0]).join('; ');
  return r.status === 200;
}
const ALLU = ['_s_admin', '_s_own', '_s_own2', '_s_m2', '_s_m1', '_s_gm', '_s_ch', '_s_sec'];
async function withServer(env, fn, opts) {
  seed(opts && opts.data);
  try {
    await startServer(env, opts);
    for (const u of ALLU) ok(await login(u), '登入 ' + u);
    await fn();
  } catch (e) { ok(false, '群組例外：' + (e && e.message), e && e.stack); }
  finally { await stopServer(); cleanup(); }
}
const DES = '/api/admin/quote-flow/designer';
const DR = '/api/admin/quote-flow/drafts';
const ENDPOINTS = [
  ['GET', DES + '/status'], ['GET', DES + '/default-flow'], ['POST', DES + '/validate'], ['POST', DES + '/simulate'], ['POST', DES + '/sweep'], ['POST', DES + '/overview'],
  ['GET', DES + '/overview'], ['GET', DR], ['GET', DR + '/fd_000000000000'], ['PUT', DR], ['DELETE', DR + '/fd_000000000000'],
];
const getDefault = async () => (await http('_s_admin', 'GET', DES + '/default-flow')).json;
const nonLeak = /_s_gm|_s_ch|_s_own|_s_m1|_s_m2|_s_sec|_s_admin/;

const G = {};
// B1 權限
G.perm = async () => {
  console.log('  [B1 權限]');
  for (const [m, u] of ENDPOINTS) {
    const anon = await http(null, m, u, m === 'GET' || m === 'DELETE' ? undefined : {});
    ok(anon.status === 401 || anon.status === 403, `B1 未登入 ${m} ${u} → 401/403（${anon.status}）`);
    for (const who of ['_s_own', '_s_m1', '_s_m2', '_s_gm', '_s_sec']) {
      const r = await http(who, m, u, m === 'GET' || m === 'DELETE' ? undefined : {});
      ok(r.status === 403, `B1 ${who} ${m} ${u} → 403（${r.status}）`);
    }
    const adm = await http('_s_admin', m, u, m === 'GET' || m === 'DELETE' ? undefined : {});
    ok(adm.status !== 401 && adm.status !== 403, `B1 管理員 ${m} ${u} 可到達處理器（${adm.status}）`);
  }
  const stat = await http('_s_admin', 'GET', DES + '/status');
  ok(stat.cache.includes('no-store'), 'B1 回應 Cache-Control: no-store');
};
// B2 狀態／預設流程
G.status = async () => {
  console.log('  [B2 狀態／預設流程]');
  const st = (await http('_s_admin', 'GET', DES + '/status')).json;
  ok(st && st.enabled === true && st.canActivate === false && /M3/.test(st.activateNote), 'B2 狀態：enabled、canActivate=false、說明 M3 開放');
  eq(st.categories.length, 9, 'B2 狀態：9 個商品歸類');
  const df = await getDefault();
  const want = QF.legacyDefault({ buSettings: { ITS: { hasMgr2Step: true } }, exceptions: [] });
  eq(canon(df.def), canon(want), 'B2 預設流程＝由目前設定轉出的 legacyDefault（逐欄位相同）');
  eq(df.defHash, QF.defHash(want), 'B2 預設流程雜湊正確');
  ok(df.checks.length === 11 && df.checks.every((c) => c.status !== 'error' || c.n === 4 || c.n === 5 || c.n === 4), 'B2 預設流程 11 條檢查無內容類錯誤');
  ok(df.view.routes.length === 4, 'B2 預設流程白話檢視');
};
// B3 草稿 CRUD
let savedId = null;
G.crud = async () => {
  console.log('  [B3 草稿 CRUD]');
  const df = await getDefault();
  const empty = await http('_s_admin', 'GET', DR);
  eq(empty.json.items, [], 'B3 一開始沒有草稿');
  const r1 = await http('_s_admin', 'PUT', DR, { name: '測試草稿一', note: '備註', def: df.def, baseVer: df.baseVer });
  ok(r1.status === 200 && r1.json.success && r1.json.created && /^fd_[0-9a-f]{12}$/.test(r1.json.id), 'B3 新增草稿 200', r1.text.slice(0, 300));
  savedId = r1.json.id;
  eq([r1.json.def.id, r1.json.def.no, r1.json.def.origin, r1.json.def.name], [savedId, 0, 'designer', '測試草稿一'], 'B3 身分由伺服器決定（id／no／origin／name）');
  const lst = await http('_s_admin', 'GET', DR);
  ok(lst.json.items.length === 1 && !('def' in lst.json.items[0]) && lst.json.items[0].routes === 4, 'B3 清單不含內容，有摘要');
  const one = await http('_s_admin', 'GET', DR + '/' + savedId);
  ok(one.status === 200 && canon(one.json.def) === canon(r1.json.def) && one.json.view.routes.length === 4, 'B3 取回單筆內容相同');
  const d = clone(one.json.def); d.routes[2].name = '總經理改名';
  await sleep(5);
  const r2 = await http('_s_admin', 'PUT', DR, { id: savedId, name: '測試草稿一', def: d, ifUpdatedAt: one.json.updatedAt });
  ok(r2.status === 200 && r2.json.created === false && r2.json.def.routes[2].name === '總經理改名' && r2.json.updatedAt !== one.json.updatedAt, 'B3 更新（帶 ifUpdatedAt）200');
  const stale = await http('_s_admin', 'PUT', DR, { id: savedId, name: '測試草稿一', def: d, ifUpdatedAt: one.json.updatedAt });
  ok(stale.status === 409 && stale.json.code === 'DRAFT_STALE', 'B3 過期的 ifUpdatedAt → 409 DRAFT_STALE', stale.text.slice(0, 200));
  const chk = await http('_s_admin', 'GET', DR + '/' + savedId);
  eq(chk.json.def.routes[2].name, '總經理改名', 'B3 409 沒有覆蓋已存內容');
  const nf = await http('_s_admin', 'PUT', DR, { id: 'fd_aaaaaaaaaaaa', name: 'x', def: d });
  ok(nf.status === 404, 'B3 更新不存在的草稿 → 404');
  const bad = [[{ name: '', def: d }, 'BAD_NAME'], [{ name: 'x'.repeat(41), def: d }, 'BAD_NAME'], [{ name: 5, def: d }, 'BAD_NAME'], [{ name: 'a\u0001', def: d }, 'BAD_NAME'], [{ name: 'ok', note: 'n'.repeat(201), def: d }, 'BAD_NOTE'], [{ name: 'ok', def: d, evil: 1 }, 'BAD_FIELD'], [{ name: 'ok', def: d, id: '../x' }, 'BAD_ID'], [{ name: 'ok', def: d, baseVer: '<x>' }, 'BAD_BASEVER']];
  for (const [b, code] of bad) { const r = await http('_s_admin', 'PUT', DR, b); ok(r.status === 400 && r.json.code === code, `B3 壞欄位 ${code}`, r.text.slice(0, 160)); }
  const del = await http('_s_admin', 'DELETE', DR + '/' + savedId);
  ok(del.status === 200 && del.json.success, 'B3 刪除 200');
  eq((await http('_s_admin', 'GET', DR + '/' + savedId)).status, 404, 'B3 刪除後取不到');
  eq((await http('_s_admin', 'DELETE', DR + '/' + savedId)).status, 404, 'B3 重複刪除 404');
  eq((await http('_s_admin', 'GET', DR + '/..%2F..%2Fetc')).status, 404, 'B3 路徑穿越形式的 id → 404');
  // 上限
  let created = 0, last = null;
  for (let i = 0; i < 52; i++) { const r = await http('_s_admin', 'PUT', DR, { name: '批量' + i, def: df.def }); if (r.status === 200) created++; else { last = r; break; } }
  ok(created === 50 && last && last.status === 409 && last.json.code === 'DRAFT_LIMIT', 'B3 草稿上限 50 份，第 51 份 → 409 DRAFT_LIMIT', created + ' ' + (last && last.text.slice(0, 100)));
  const all = (await http('_s_admin', 'GET', DR)).json.items;
  for (const it of all) await http('_s_admin', 'DELETE', DR + '/' + it.id);
  eq((await http('_s_admin', 'GET', DR)).json.items.length, 0, 'B3 清空');
  // 並發：兩個請求用同一個 ifUpdatedAt，只有一個成功
  const c1 = await http('_s_admin', 'PUT', DR, { name: '並發', def: df.def });
  const dd = clone(c1.json.def);
  const rr = await Promise.all([1, 2].map((i) => { const x = clone(dd); x.routes[0].name = '並發' + i; return http('_s_admin', 'PUT', DR, { id: c1.json.id, name: '並發', def: x, ifUpdatedAt: c1.json.updatedAt }); }));
  eq(rr.map((r) => r.status).sort(), [200, 409], 'B3 並發儲存（同一個 ifUpdatedAt）：一成功一 409');
  await http('_s_admin', 'DELETE', DR + '/' + c1.json.id);
};
// B4 稽核
G.audit = async () => {
  console.log('  [B4 稽核]');
  const df = await getDefault();
  const before = ((await http('_s_admin', 'GET', '/api/admin/logs')).json || []).length;
  const r = await http('_s_admin', 'PUT', DR, { name: '稽核測試<草稿>', def: df.def });
  const bad = clone(df.def); bad.routes[1].path = ['n_mgr1'];
  const rb = await http('_s_admin', 'PUT', DR, { name: '應被拒', def: bad });
  ok(rb.status === 400, 'B4 被拒的儲存不留 SAVE 紀錄（前置）');
  await http('_s_admin', 'GET', DR + '/' + r.json.id);
  await http('_s_admin', 'DELETE', DR + '/' + r.json.id);
  const logs = (await http('_s_admin', 'GET', '/api/admin/logs')).json || [];
  const save = logs.filter((l) => l.action === 'SAVE_FLOW_DRAFT' && l.target === '稽核測試<草稿>'), del = logs.filter((l) => l.action === 'DELETE_FLOW_DRAFT' && l.target === '稽核測試<草稿>');
  eq(save.length, 1, 'B4 成功的儲存留下 1 筆 SAVE_FLOW_DRAFT（被拒的沒有）');
  eq(del.length, 1, 'B4 刪除留下 1 筆 DELETE_FLOW_DRAFT');
  ok(save[0] && save[0].operator === '_s_admin' && del[0] && del[0].operator === '_s_admin', 'B4 稽核記錄操作者為管理員', JSON.stringify(save[0]).slice(0, 200));
  ok(save[0] && /草稿不影響任何簽核/.test(JSON.stringify(save[0])) && !/"schema"|n_mgr1|cmp/.test(JSON.stringify(save[0])), 'B4 稽核詳情是摘要，不含流程定義內容');
  ok(logs.length >= before + 2, 'B4 稽核總數增加');
};
// B5 攻擊輸入
G.attack = async () => {
  console.log('  [B5 攻擊輸入]');
  const df = await getDefault();
  const post = (u, b, raw) => http('_s_admin', 'POST', u, b, raw);
  const put = (b, raw) => http('_s_admin', 'PUT', DR, b, raw);
  const evils = [
    [null, 'null'], ['x', '字串'], [[], '陣列'], [7, '數字'],
    [Object.assign(clone(df.def), { evil: 1 }), '未知鍵'], [JSON.parse('{"__proto__":{"polluted":1},"schema":1}'), '__proto__'],
  ];
  for (const [b, nm] of evils) {
    for (const [fn, label] of [[(d) => put({ name: 'x', def: d }), 'PUT'], [(d) => post(DES + '/validate', { def: d }), 'validate'], [(d) => post(DES + '/simulate', { def: d, input: { amount: '1' } }), 'simulate'], [(d) => post(DES + '/sweep', { def: d }), 'sweep'], [(d) => post(DES + '/overview', { def: d }), 'overview']]) {
      const r = await fn(b);
      ok(r.status === 400 && r.json && typeof r.json.code === 'string', `B5 ${label} 拒絕 ${nm}（${r.status}）`, r.text.slice(0, 120));
    }
  }
  let deepStr = '{"name":"x","def":'; for (let i = 0; i < 4000; i++) deepStr += '{"a":'; deepStr += '1'; for (let i = 0; i < 4000; i++) deepStr += '}'; deepStr += '}';
  let r = await put(deepStr, true);
  ok(r.status === 400 || r.status === 413 || r.status === 500, 'B5 4000 層深 JSON 被拒或失敗但不中斷（' + r.status + '）');
  r = await put({ name: 'x', def: Object.assign(clone(df.def), { tables: JSON.parse(deepStr).def }) });
  ok(r.status === 400 && r.json.code === 'DEF_TOO_DEEP', 'B5 def 內 4000 層深 → DEF_TOO_DEEP', r.text.slice(0, 160));
  const big = clone(df.def); big.routes[0].name = 'y'.repeat(70000);
  r = await put({ name: 'x', def: big });
  ok(r.status === 400 && r.json.code === 'DEF_TOO_LARGE', 'B5 超過 64KB → DEF_TOO_LARGE', r.text.slice(0, 100));
  r = await put('{bad json', true);
  ok(r.status === 400 || r.status === 500, 'B5 壞 JSON 不當機（' + r.status + '）');
  const cond = [{ cmp: ['evil', '>', '1'] }, { cmp: ['amount', '>', 5] }, { cmp: ['amount', '>', '1e9'] }, { eval: 1 }, { all: 'x' }, { cmp: { 0: 'amount' } }, { cmp: ['__proto__', '==', 'x'] }, { cmp: ['constructor', '==', 'x'] }];
  for (const c of cond) { const d = clone(df.def); d.routes[2].when = c; const rv = await post(DES + '/validate', { def: d }); ok(rv.status === 200 && rv.json.ok === false, 'B5 壞條件被 validate 判為不可儲存 ' + JSON.stringify(c)); const rp = await put({ name: 'x', def: d }); ok(rp.status === 400, 'B5 壞條件不能存 ' + JSON.stringify(c)); }
  const ids = ['__proto__', 'constructor', 'toString', 'hasOwnProperty', 'prototype'];
  for (const id of ids) { const d = clone(df.def); d.routes[2].id = id; const rp = await put({ name: 'x', def: d }); ok(rp.status === 400, `B5 保留字路線 id「${id}」不能存`); }
  for (const id of ids) { const d = clone(df.def); d.nodes[0].id = id; d.routes.forEach((x) => { x.path = x.path.map((p) => (p === 'n_mgr2' ? id : p)); }); const rp = await put({ name: 'x', def: d }); ok(rp.status === 400, `B5 保留字節點 id「${id}」不能存`); }
  const sims = [{ amount: '1e3' }, { amount: -1 }, { amount: '1', marginLevel: '1' }, { amount: '1', ownerBu: { a: 1 } }, { amount: '1', classKeys: 'x' }, { amount: '1', noMgr2: 'yes' }, 'str', null];
  for (const s of sims) { const rs = await post(DES + '/simulate', { def: df.def, input: s }); ok(rs.status === 400, 'B5 模擬壞輸入被拒 ' + JSON.stringify(s)); }
  eq((await http('_s_admin', 'GET', DR)).json.items.length, 0, 'B5 攻擊後沒有任何草稿被存入');
  const okAfter = await put({ name: '攻擊後正常', def: df.def });
  ok(okAfter.status === 200, 'B5 攻擊後伺服器仍正常儲存');
  await http('_s_admin', 'DELETE', DR + '/' + okAfter.json.id);
  // 100 次連打 validate 不當機
  for (let i = 0; i < 20; i++) await post(DES + '/validate', { def: df.def });
  ok((await http('_s_admin', 'GET', DES + '/status')).status === 200, 'B5 連續請求後狀態端點正常');
};
// B6 底線被後端拒絕
G.floor = async () => {
  console.log('  [B6 底線（後端拒絕）]');
  const df = await getDefault();
  const muts = [
    ['董事長路線拿掉董事長', (d) => { route(d, 'r_l3').path = ['n_mgr1', 'n_gm']; }],
    ['1,000 萬門檻改 2,000 萬', (d) => { route(d, 'r_l3').when = { any: [{ cmp: ['amount', '>', '20000000'] }, { cmp: ['marginLevel', '==', 3] }] }; }],
    ['刪掉董事會路線', (d) => { d.routes.splice(0, 1); }],
    ['5,000 萬門檻改 6,000 萬', (d) => { route(d, 'r_board').when = { cmp: ['amount', '>', '60000000'] }; }],
    ['調整規則移掉董事長', (d) => { d.modifiers.push({ id: 'm_rmc', name: '大單移除董事長', op: 'remove', target: 'n_chairman', when: { cmp: ['amount', '>', '10000000'] } }); }],
  ];
  for (const [nm, f] of muts) {
    const d = clone(df.def); f(d);
    const rv = await http('_s_admin', 'POST', DES + '/validate', { def: d });
    ok(rv.status === 200 && rv.json.ok === false && rv.json.floor.ok === false && rv.json.floor.violations.length > 0, `B6 validate 回報違反底線：${nm}`, rv.text.slice(0, 160));
    const rp = await http('_s_admin', 'PUT', DR, { name: '底線測試', def: d });
    ok(rp.status === 400 && rp.json.code === 'DRAFT_INVALID' && rp.json.blocking.some((b) => b.n === 10), `B6 PUT 被拒（400 DRAFT_INVALID 第 10 條）：${nm}`, rp.text.slice(0, 200));
    ok(rp.json.blocking.find((b) => b.n === 10).issues.every((t) => /金額超過/.test(t) || /另有/.test(t)), `B6 逐條原因為白話：${nm}`);
  }
  eq((await http('_s_admin', 'GET', DR)).json.items.length, 0, 'B6 違反底線的草稿一份都沒存進去');
  // 更新既有草稿時違規也擋，且保留原內容
  const good = await http('_s_admin', 'PUT', DR, { name: '好草稿', def: df.def });
  const bad = clone(good.json.def); route(bad, 'r_l3').path = ['n_mgr1', 'n_gm'];
  const rb = await http('_s_admin', 'PUT', DR, { id: good.json.id, name: '好草稿', def: bad, ifUpdatedAt: good.json.updatedAt });
  ok(rb.status === 400, 'B6 更新成違規內容被拒');
  eq(canon((await http('_s_admin', 'GET', DR + '/' + good.json.id)).json.def.routes[1].path), canon(df.def.routes[1].path), 'B6 被拒後原草稿內容不變');
  await http('_s_admin', 'DELETE', DR + '/' + good.json.id);
  // 合法放寬可存（放寬偵測只警告）
  const lo = clone(df.def); route(lo, 'r_l2').path = ['n_mgr1'];
  const rl = await http('_s_admin', 'PUT', DR, { name: '放寬版', def: lo });
  ok(rl.status === 200 && rl.json.checks.find((c) => c.n === 9).status === 'warn', 'B6 放寬（不違反底線）可存，但第 9 條警告');
  await http('_s_admin', 'DELETE', DR + '/' + rl.json.id);
};
// B7 模擬＝真實送簽（舊路徑 oracle）
G.simreal = async () => {
  console.log('  [B7 模擬＝真實送簽]');
  const df = await getDefault();
  const data0 = readData();
  const prodCat = { PS: 'software', CS: 'consult', HW: 'hardware' };
  let cmp = 0, diff = 0;
  for (const q of data0.quotations) {
    const sub = await http(q.owner, 'POST', `/api/quotations/${q.id}/submit`, {});
    if (sub.status !== 200) { ok(false, `B7 送簽 ${q.id} 應成功`, sub.text.slice(0, 160)); continue; }
    const real = readData().quotations.find((x) => x.id === q.id).approval;
    const price = q.items[0].unitPrice, cost = q.items[0].cost;
    const marginPct = ((price - cost) / price * 100).toFixed(2);
    const bu = q.owner === '_s_own' ? 'ITS' : 'ERP';
    const rs = await http('_s_admin', 'POST', DES + '/simulate', { def: df.def, input: { amount: String(price), marginPct, category: prodCat[q.products[0]], ownerBu: bu } });
    ok(rs.status === 200, `B7 模擬 ${q.id} 200`, rs.text.slice(0, 160));
    cmp++;
    const got = (rs.json.tiers || []).join('>'), want = real.steps.map((s) => s.tier).join('>');
    if (got !== want) { diff++; console.log('  B7 不一致', q.id, price, cost, got, want); }
  }
  eq(diff, 0, `B7 模擬的關卡序列＝真實送簽（舊路徑）的 steps（${cmp} 張）`);
};
// B8 草稿不影響任何簽核（含 QUOTE_FLOW_ENGINE=on）
G.noeffect = async () => {
  console.log('  [B8 草稿不影響簽核]');
  const df = await getDefault();
  const loose = clone(df.def); route(loose, 'r_l2').path = ['n_mgr1']; route(loose, 'r_default').path = ['n_mgr1'];
  const r = await http('_s_admin', 'PUT', DR, { name: '放寬版', def: loose });
  ok(r.status === 200, 'B8 先存一份放寬草稿');
  const subL2 = await http('_s_own2', 'POST', '/api/quotations/q3/submit', {});   // ERP 業務、顧問服務 80 萬、毛利剛好 15% → 層級 2（總經理）
  ok(subL2.status === 200, 'B8 送簽成功', subL2.text.slice(0, 200));
  const a = readData().quotations;
  const tiers = (id) => a.find((q) => q.id === id).approval.steps.map((s) => s.tier).join('>');
  eq(tiers('q3'), 'mgr1>gm', 'B8 草稿把總經理路線放寬成只有一級主管，但實際送簽仍走原規則（層級 2：一級主管＞總經理）');
  ok(!('flow' in a.find((q) => q.id === 'q3').approval), 'B8 單上沒有 approval.flow（QUOTE_FLOW_ENGINE 未設）');
  ok(JSON.stringify(a).indexOf(r.json.id) < 0, 'B8 報價單資料裡沒有任何草稿代號');
  ok(readData().quoteFlowDrafts && readData().quoteFlowDrafts.items.length === 1, 'B8 草稿存在獨立命名空間 quoteFlowDrafts');
};
G.noeffectOn = async () => {
  console.log('  [B8b 草稿不影響簽核（引擎 on）]');
  const df = await getDefault();
  const loose = clone(df.def); route(loose, 'r_l2').path = ['n_mgr1'];
  await http('_s_admin', 'PUT', DR, { name: '放寬版', def: loose });
  const subL2 = await http('_s_own2', 'POST', '/api/quotations/q3/submit', {});
  ok(subL2.status === 200, 'B8b 引擎 on 送簽成功', subL2.text.slice(0, 200));
  const q = readData().quotations.find((x) => x.id === 'q3');
  eq(q.approval.steps.map((s) => s.tier).join('>'), 'mgr1>gm', 'B8b 引擎 on：仍走目前設定轉出的預設流程，不讀草稿');
  ok(q.approval.flow && q.approval.flow.origin === 'legacy', 'B8b 快照的 origin＝legacy（不是 designer）');
  ok(!JSON.stringify(q).includes('"designer"'), 'B8b 單上沒有 designer 來源的任何痕跡');
};
// B9 不洩漏帳號／成本
G.noleak = async () => {
  console.log('  [B9 不洩漏]');
  const df = await getDefault();
  const post = (u, b) => http('_s_admin', 'POST', u, b);
  const outs = [
    (await post(DES + '/validate', { def: df.def })).text,
    (await post(DES + '/simulate', { def: df.def, input: { amount: '12000000', marginLevel: 1, category: 'consult', ownerBu: 'ITS' }, compare: true })).text,
    (await post(DES + '/sweep', { def: df.def, input: { amount: '1', marginLevel: 1 } })).text,
    (await post(DES + '/overview', { def: df.def, input: { amount: '12000000', marginLevel: 1, category: 'consult', ownerBu: 'ITS' } })).text,
    (await http('_s_admin', 'GET', DES + '/status')).text,
  ];
  outs.forEach((t, i) => ok(!nonLeak.test(t), `B9 回應 #${i} 不含任何帳號名稱`));
  const r = await http('_s_admin', 'PUT', DR, { name: '洩漏測試', def: df.def });
  await http('_s_admin', 'DELETE', DR + '/' + r.json.id);
};
// B10 XSS 內容原樣存取
G.xss = async () => {
  console.log('  [B10 XSS 內容]');
  const df = await getDefault();
  const d = clone(df.def); route(d, 'r_l2').name = '<svg onload=a()>';
  d.modifiers[0].name = '<script>x()</script>';
  const r = await http('_s_admin', 'PUT', DR, { name: '<b>名稱</b>&"\'', note: '<svg onload=alert(1)>', def: d });
  ok(r.status === 200, 'B10 路線／調整規則／草稿名稱含 HTML 字元可儲存', r.text.slice(0, 160));
  const g = await http('_s_admin', 'GET', DR + '/' + r.json.id);
  ok(g.ctype.includes('application/json') && g.json.def.routes[2].name === '<svg onload=a()>' && g.json.name === '<b>名稱</b>&"\'' && g.json.note === '<svg onload=alert(1)>', 'B10 原樣存取（JSON，前端以純文字顯示）；Content-Type 為 application/json');
  const lb = clone(df.def); lb.nodes.find((n) => n.tier === 'gm').label = '<b>x</b>';
  eq((await http('_s_admin', 'PUT', DR, { name: 'x', def: lb })).status, 400, 'B10 節點名稱含 < 被擋（會流到未來的簽核卡片）');
  await http('_s_admin', 'DELETE', DR + '/' + r.json.id);
};
// B11 總覽
G.overview = async () => {
  console.log('  [B11 總覽／掃描 API]');
  const df = await getDefault();
  const o = await http('_s_admin', 'POST', DES + '/overview', { def: df.def });
  ok(o.status === 200 && o.json.graph.nodes.length >= 10 && o.json.graph.edges.length >= 10 && o.json.sim === null, 'B11 總覽（無模擬）');
  const o2 = await http('_s_admin', 'POST', DES + '/overview', { def: df.def, input: { amount: '12000000', marginLevel: 1, category: 'consult', ownerBu: 'ITS' } });
  ok(o2.status === 200 && o2.json.sim.tiers.join() === 'mgr2,mgr1,gm,chairman' && o2.json.graph.nodes.some((n) => n.hit), 'B11 總覽＋模擬高亮');
  const og = await http('_s_admin', 'GET', DES + '/overview');
  ok(og.status === 200 && og.json.source === 'default' && canon(og.json.graph) === canon(o.json.graph), 'B11 GET 總覽（預設流程）＝POST 總覽（同一份定義）');
  const dr = await http('_s_admin', 'PUT', DR, { name: '總覽草稿', def: df.def });
  const ogd = await http('_s_admin', 'GET', DES + '/overview?draftId=' + dr.json.id);
  ok(ogd.status === 200 && ogd.json.source === 'draft' && ogd.json.name === '總覽草稿' && ogd.json.graph.nodes.length === o.json.graph.nodes.length, 'B11 GET 總覽（以草稿代號載入）');
  eq((await http('_s_admin', 'GET', DES + '/overview?draftId=..%2F..')).status, 400, 'B11 GET 總覽：壞的草稿代號 400');
  eq((await http('_s_admin', 'GET', DES + '/overview?draftId=fd_aaaaaaaaaaaa')).status, 404, 'B11 GET 總覽：不存在的草稿 404');
  eq((await http('_s_admin', 'GET', DES + '/overview?draftId[]=x')).status, 400, 'B11 GET 總覽：陣列形式的 draftId 400');
  await http('_s_admin', 'DELETE', DR + '/' + dr.json.id);
  const sw = await http('_s_admin', 'POST', DES + '/sweep', { def: df.def, input: { amount: '1', marginLevel: 1, category: 'consult', ownerBu: 'ERP' } });
  ok(sw.status === 200 && sw.json.rows.length >= 6 && sw.json.summary.floorViolations === 0, 'B11 掃描 API');
  const v = await http('_s_admin', 'POST', DES + '/validate', { def: df.def });
  ok(v.status === 200 && v.json.ok && v.json.checks.length === 11 && v.json.samples > 1000, 'B11 validate API');
  const t0 = Date.now(); await http('_s_admin', 'POST', DES + '/validate', { def: df.def }); const ms = Date.now() - t0;
  ok(ms < 1500, 'B11 validate 回應時間 < 1.5 秒（' + ms + 'ms）');
};
async function runApi() {
  console.log('[B API（開關 on）]');
  // 全域 API 限流是每分鐘 300 次（行程內記憶體）：分成幾台伺服器跑，每台請求數都遠低於上限
  const on = { QUOTE_FLOW_DESIGNER: '1' };
  await withServer(on, async () => { for (const k of ['perm', 'status', 'audit']) await G[k](); });
  await withServer(on, async () => { await G.crud(); });
  await withServer(on, async () => { await G.attack(); });
  await withServer(on, async () => { for (const k of ['floor', 'noleak', 'xss', 'overview']) await G[k](); });
  await withServer(on, async () => { await G.simreal(); });
  await withServer(on, async () => { await G.noeffect(); });
  await withServer(Object.assign({ QUOTE_FLOW_ENGINE: 'on' }, on), async () => { await G.noeffectOn(); });
}

// ═══════════════════════════ 開關關閉：對基底位元級不變 ═══════════════════════════
function sh(cmd, args, opt) { return spawnSync(cmd, args, Object.assign({ encoding: 'utf8', maxBuffer: 1 << 28 }, opt || {})); }
async function runFlagOff() {
  console.log('[C 開關關閉 vs 基底]');
  const ref = process.env.FD_BASE_REF || 'feat/quote-flow-engine';
  const baseDir = path.join(os.tmpdir(), 'fd_base_' + crypto.randomBytes(4).toString('hex'));
  fs.mkdirSync(baseDir, { recursive: true });
  const ar = spawnSync('git', ['archive', '--format=tar', ref], { cwd: ROOT, maxBuffer: 1 << 29 });
  if (ar.status !== 0) { ok(false, 'C 取得基底失敗：' + String(ar.stderr).slice(0, 200)); return; }
  const tx = spawnSync('tar', ['-xf', '-'], { cwd: baseDir, input: ar.stdout, maxBuffer: 1 << 29 });
  ok(tx.status === 0, 'C 解開基底封存');
  const mk = sh('cmd', ['/c', 'mklink', '/J', path.join(baseDir, 'node_modules'), path.join(ROOT, 'node_modules')]);
  ok(fs.existsSync(path.join(baseDir, 'node_modules')), 'C 基底 node_modules junction');
  const norm = (t, p) => String(t).split(p).join('<P>').replace(/\d{4}-\d\d-\d\dT[\d:.]+Z/g, '<T>').replace(/"since":"[^"]*"/g, '"since":"<T>"').replace(/"loginAt":[0-9]+/g, '"loginAt":0').replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<U>');
  const collect = async (label) => {
    const res = [];
    for (const who of [null, '_s_admin', '_s_own', '_s_gm']) {
      for (const [m, u] of ENDPOINTS.concat([['GET', '/api/admin/quote-flow/designer-nope'], ['GET', '/api/admin/quote-flow/drafts-nope']])) {
        const r = await http(who, m, u, m === 'GET' || m === 'DELETE' ? undefined : {});
        res.push({ k: `${who} ${m} ${u}`, status: r.status, ctype: r.ctype, body: norm(r.text, u) });
      }
    }
    // 其他端點：確認沒有因為接線而改變
    for (const [m, u] of [['GET', '/api/me'], ['GET', '/api/quotations'], ['GET', '/api/admin/quote-flow/engine-status'], ['GET', '/api/admin/quote-approval/flow-config'], ['GET', '/api/admin/logs']]) {
      const r = await http('_s_admin', m, u);
      res.push({ k: `admin ${m} ${u}`, status: r.status, ctype: r.ctype, body: norm(r.text, u).replace(/"id":"[^"]*"/g, '"id":"<id>"').replace(/"time":"[^"]*"/g, '"time":"<T>"') });
    }
    return res;
  };
  let baseRes = null, mineRes = null;
  DATADIR = baseDir;
  seed();
  try {
    await startServer({}, { cwd: baseDir });
    for (const u of ALLU) await login(u);
    baseRes = await collect('base');
    // 送簽流程也比對（舊路徑）
    for (const q of ['q0', 'q5']) await http('_s_own', 'POST', `/api/quotations/${q}/submit`, {});
    baseRes.push({ k: 'submit state', body: JSON.stringify(readData().quotations.filter((x) => ['q0', 'q5'].includes(x.id)).map((x) => x.approval && x.approval.steps.map((s) => s.tier))) });
  } finally { await stopServer(); cleanup(); }
  DATADIR = ROOT;
  seed();
  try {
    await startServer({}, {});
    for (const u of ALLU) await login(u);
    mineRes = await collect('mine');
    for (const q of ['q0', 'q5']) await http('_s_own', 'POST', `/api/quotations/${q}/submit`, {});
    mineRes.push({ k: 'submit state', body: JSON.stringify(readData().quotations.filter((x) => ['q0', 'q5'].includes(x.id)).map((x) => x.approval && x.approval.steps.map((s) => s.tier))) });
  } finally { await stopServer(); cleanup(); }
  let diffs = 0;
  baseRes.forEach((b, i) => {
    const m = mineRes[i];
    const same = m && m.k === b.k && m.status === b.status && m.ctype === b.ctype && m.body === b.body;
    if (!same) { diffs++; console.log('  C 差異', b.k, b.status, m && m.status); console.log('   base:', b.body || ''); console.log('   mine:', (m && m.body) || ''); }
  });
  eq(diffs, 0, `C 開關關閉：${baseRes.length} 個回應與基底逐位元相同（含 10 個設計器端點 × 4 種身分、未註冊路徑、送簽結果）`);
  console.log('  （關閉時各身分狀態）', JSON.stringify(Object.fromEntries(['null', '_s_admin', '_s_own', '_s_gm'].map((w) => [w, [...new Set(baseRes.filter((r) => r.k.startsWith(w + ' ') && /designer|drafts/.test(r.k)).map((r) => r.status))]]))));
  const dr = baseRes.filter((r) => /designer|drafts/.test(r.k));
  ok(dr.length >= 40 && dr.filter((r) => !r.k.startsWith('null ')).every((r) => r.status === 404) && dr.filter((r) => r.k.startsWith('null ')).every((r) => r.status === 401), 'C 關閉時設計器端點：已登入者一律 404、未登入 401（與基底對未註冊路徑的行為相同）');
  sh('cmd', ['/c', 'rmdir', path.join(baseDir, 'node_modules')]);   // 只移除 junction，不遞迴
  try { fs.rmSync(baseDir, { recursive: true, force: true }); } catch { /* 略 */ }
  ok(fs.existsSync(path.join(ROOT, 'node_modules', 'bcryptjs')), 'C 清理後專案 node_modules 仍完整');
}

// ═══════════════════════════ 變異測試 ═══════════════════════════
const UNIT_MUTANTS = [
  ['底線判斷改錯：1,000 萬改成 2,000 萬才檢查', 'quoteFlowDesigner.js', "if (rev > FLOOR.chairmanCents && !(tiers", "if (rev > FLOOR.chairmanCents * 2 && !(tiers"],
  ['底線判斷改錯：5,000 萬檢查失效', 'quoteFlowDesigner.js', "if (rev > FLOOR.boardCents && !tiers.includes('board'))", "if (false && rev > FLOOR.boardCents && !tiers.includes('board'))"],
  ['底線判斷改錯：董事會代核不算滿足董事長', 'quoteFlowDesigner.js', "!(tiers.includes('chairman') || tiers.includes('board'))", "!(tiers.includes('chairman'))"],
  ['取樣漏字面值（只取 +7，不含字面值本身）', 'quoteFlowDesigner.js', "[L - 1, L, L + 1].forEach((x) => { if (x >= 1 && x <= MAXC) amt.add(x); });", "[L + 7].forEach((x) => { if (x >= 1 && x <= MAXC) amt.add(x); });"],
  ['取樣不含 1,000 萬／5,000 萬附近與高金額點（只剩字面值）', 'quoteFlowDesigner.js', ["const amt = new Set([1, 100, FLOOR.chairmanCents - 1, FLOOR.chairmanCents, FLOOR.chairmanCents + 1, FLOOR.boardCents - 1, FLOOR.boardCents, FLOOR.boardCents + 1]);", "const top = Math.min(MAXC, Math.max(sorted[sorted.length - 1] * 2, FLOOR.boardCents * 2));", "const amounts = [...new Set(sorted.concat(mids, [top, MAXC]))].sort((a, b) => a - b);"], ["const amt = new Set([1, 100]);", "const top = 100;", "const amounts = sorted.slice();"]],
  ['底線不列入擋存條件（驗證漏一條）', 'quoteFlowDesigner.js', "const BLOCKING = [1, 2, 3, 6, 10, 11];", "const BLOCKING = [1, 2, 3, 6, 11];"],
  ['標籤檢查漏掉 HTML 字元', 'quoteFlowDesigner.js', "|| /[<>\"'&]/.test(s);", "|| false;"],
  ['鎖定節點改名不再被擋', 'quoteFlowDesigner.js', "if (n.tier === 'board' && n.label !== QA.TIERS.board)", "if (false && n.tier === 'board' && n.label !== QA.TIERS.board)"],
  ['放寬偵測方向反了', 'quoteFlowDesigner.js', "if (rankOf(newT) < rankOf(oldT)) return true;", "if (rankOf(newT) > rankOf(oldT)) return true;"],
  ['模擬與引擎分歧：層級 1 的毛利少 1 分', 'quoteFlowDesigner.js', "gp = ceilDiv(revenue * BigInt(row.l1), 100n); }", "gp = ceilDiv(revenue * BigInt(row.l1), 100n) - 1n; }"],
  ['模擬與引擎分歧：回傳財務層級而非實際關卡', 'quoteFlowDesigner.js', "blockers: p.blockers, tiers: p.actualTiers.slice(),", "blockers: p.blockers, tiers: p.financialTiers.slice(),"],
  ['淨化不再擋未知鍵', 'quoteFlowDesigner.js', "return bad('DEF_UNKNOWN_KEY', '不認得的欄位：' + clipText(k, 30), at);", "return null;"],
  ['淨化深度上限失效', 'quoteFlowDesigner.js', "if (depthOf(raw, LIMITS.depth) > LIMITS.depth) return", "if (false) return"],
  ['總覽連線端點寫錯', 'quoteFlowDesigner.js', "edge('d' + i, first, '符合', 'yes', hit);", "edge('d' + i, 'zz', '符合', 'yes', hit);"],
  ['總覽模擬高亮不換成實際序列', 'quoteFlowDesigner.js', "if (i === hitIdx && sim) return sim.tiers", "if (false) return sim.tiers"],
  ['掃描邊界漏 ±0.01', 'quoteFlowDesigner.js', "[L - 1, L, L + 1].forEach((x) => { if (x >= 1 && x <= MAXC && !amounts.includes(x)) amounts.push(x); });", "[L].forEach((x) => { if (x >= 1 && x <= MAXC && !amounts.includes(x)) amounts.push(x); });"],
  ['開關解析放寬（true 也開）', 'quoteFlowDesigner.js', "=== '1'; }", "=== '1' || String(raw).trim() === 'true'; }"],
];
const API_MUTANTS = [
  ['開關失效（關閉時仍放行）', "FD.designerEnabled(process.env.QUOTE_FLOW_DESIGNER) ? next() : next('route')", "next()", ['flagoff']],
  ['權限漏檢：儲存草稿不檢查管理員', "app.put('/api/admin/quote-flow/drafts', gate, ...qAdmin,", "app.put('/api/admin/quote-flow/drafts', gate, qAdmin[1],", ['perm']],
  ['權限漏檢：模擬不檢查管理員', "app.post('/api/admin/quote-flow/designer/simulate', gate, ...qAdmin,", "app.post('/api/admin/quote-flow/designer/simulate', gate, qAdmin[1],", ['perm']],
  ['儲存不驗證（違規草稿存得進去）', "const rc = checksOr400(res, def, env, true);", "const rc = checksOr400(res, def, env, false);", ['floor', 'attack']],
  ['儲存不寫稽核', "writeLog('SAVE_FLOW_DRAFT', ctx.me,", "(() => {})('SAVE_FLOW_DRAFT', ctx.me,", ['audit']],
  ['刪除不寫稽核', "writeLog('DELETE_FLOW_DRAFT', ctx.me,", "(() => {})('DELETE_FLOW_DRAFT', ctx.me,", ['audit']],
  ['樂觀並行失效', "b.ifUpdatedAt !== existing.updatedAt)", "false)", ['crud']],
  ['草稿上限失效', "st.items.length >= FD.LIMITS.drafts)", "false)", ['crud']],
  ['身分不由伺服器決定', "def.id = id; def.no = 0; def.origin = 'designer'; def.name = name;", "", ['crud']],
  ['回應快取未關', "const noStore = (res) => res.setHeader('Cache-Control', 'no-store');", "const noStore = (res) => res;", ['perm']],
];
function runUnitUnder(file, label) {
  // 以子行程執行單元區（FD_FILE 指向變異檔），回傳失敗數
  const r = spawnSync(process.execPath, [__filename], { cwd: ROOT, env: Object.assign({}, process.env, { FD_ONLY: 'unit', FD_FILE: file, FD_FUZZ_N: '120', FD_MUT: '' }), encoding: 'utf8', maxBuffer: 1 << 26 });
  const m = /結果：(\d+) 通過，(\d+) 失敗/.exec(r.stdout || '');
  return m ? { pass: +m[1], fail: +m[2], out: r.stdout } : { pass: 0, fail: 999, out: (r.stdout || '') + (r.stderr || '') };
}
async function runMutants() {
  console.log('[D 變異測試]');
  const tmp = [];
  try {
    for (const [name, f, find, rep] of UNIT_MUTANTS) {
      let src = fs.readFileSync(path.join(ROOT, 'lib', f), 'utf8');
      const finds = [].concat(find), reps = [].concat(rep);
      let allOk = true;
      finds.forEach((fd, k) => { const idx = src.indexOf(fd); const one = idx >= 0 && src.indexOf(fd, idx + 1) < 0; ok(one, 'D 變異點唯一存在：' + name, idx < 0 ? '找不到' : '出現多次'); if (!one) { allOk = false; return; } src = src.slice(0, idx) + reps[k] + src.slice(idx + fd.length); });
      if (!allOk) continue;
      const mf = path.join(ROOT, 'lib', 'quoteFlowDesigner.__mut_' + tmp.length + '.js');
      fs.writeFileSync(mf, src); tmp.push(mf);
      const r = runUnitUnder(mf, name);
      ok(r.fail > 0, 'D 單元變異必須使測試變紅：' + name, `通過 ${r.pass} 失敗 ${r.fail}`);
    }
    for (const [name, find, rep, groups] of API_MUTANTS) {
      const rf = path.join(ROOT, 'lib', 'quoteFlowDesignerRoutes.js');
      const src = fs.readFileSync(rf, 'utf8');
      const idx = src.indexOf(find);
      ok(idx >= 0 && src.indexOf(find, idx + 1) < 0, 'D 變異點唯一存在：' + name, idx < 0 ? '找不到' : '出現多次');
      if (idx < 0) continue;
      const tag = '__mutapi' + tmp.length;
      const mRoutes = path.join(ROOT, 'lib', 'quoteFlowDesignerRoutes.' + tag + '.js');
      fs.writeFileSync(mRoutes, src.slice(0, idx) + rep + src.slice(idx + find.length)); tmp.push(mRoutes);
      const mQ = path.join(ROOT, 'lib', 'quoteRoutes.' + tag + '.js');
      fs.writeFileSync(mQ, fs.readFileSync(path.join(ROOT, 'lib', 'quoteRoutes.js'), 'utf8').replace("require('./quoteFlowDesignerRoutes')", () => "require('./quoteFlowDesignerRoutes." + tag + ".js')")); tmp.push(mQ);
      const mS = path.join(ROOT, 'server.' + tag + '.js');
      fs.writeFileSync(mS, fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8').replace("require('./lib/quoteRoutes')(app", () => "require('./lib/quoteRoutes." + tag + ".js')(app")); tmp.push(mS);
      const before = fail, passB = pass, fl = failures.length;
      quiet = true;
      if (groups.includes('flagoff')) {
        // 開關失效：關閉時（未設環境變數）管理員打 status 應是 404
        seed();
        try { await startServer({}, { file: mS }); await login('_s_admin'); const r = await http('_s_admin', 'GET', DES + '/status'); ok(r.status === 404, '（變異內）開關關閉時 status 404', r.status); } finally { await stopServer(); cleanup(); }
      } else {
        await withServer({ QUOTE_FLOW_DESIGNER: '1' }, async () => { for (const g of groups) await G[g](); }, { file: mS });
      }
      const red = fail > before;
      // 變異內部產生的通過／失敗不計入總計，只記錄「測試有沒有抓到」
      quiet = false; pass = passB; fail = before; failures.length = fl;
      ok(red, 'D API 變異必須使測試變紅：' + name);
    }
  } finally { tmp.forEach((f) => { try { fs.unlinkSync(f); } catch { /* 略 */ } }); cleanup(); }
}

(async () => {
  try {
    if (!ONLY || ONLY === 'unit') runUnit();
    if (!ONLY || ONLY === 'api') await runApi();
    if (!ONLY || ONLY === 'flagoff') await runFlagOff();
    if (process.env.FD_MUT === '1') await runMutants();
  } catch (e) {
    fail++; failures.push('例外：' + (e && e.stack || e));
    console.log('例外：', e && e.stack || e);
  } finally {
    await stopServer(); cleanup();
    for (const pid of allPids) { try { process.kill(pid); } catch { /* 已結束 */ } }
  }
  console.log(`\n簽核流程設計器 M2 後端檢查結果：${pass} 通過，${fail} 失敗`);
  if (fail) console.log(failures.map((f) => ' - ' + f).join('\n'));
  process.exit(fail ? 1 : 0);
})();
