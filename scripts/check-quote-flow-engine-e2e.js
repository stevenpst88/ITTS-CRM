#!/usr/bin/env node
/**
 * 簽核流程引擎（M1）端對端檢查：真實 server.js（JSON 模式沙盒）＋真實 HTTP 登入／送簽／核准／後台狀態端點。
 * 用法：node scripts/check-quote-flow-engine-e2e.js（預設埠 3120，QFE_PORT 可改；拒絕保留埠；工作樹已有 auth.json／data.json／稽核檔時拒絕執行）
 * 臨時 _s_* 帳號與資料檔只存在於本次執行，結束以 PID 關閉伺服器並刪檔；不連任何遠端資料庫。
 *   1) QUOTE_FLOW_ENGINE 未設／off：送簽走原路徑、資料沒有 approval.flow；狀態端點 mode=off、引擎沒跑
 *   2) shadow：送簽結果與 off 相同、資料沒有 flow；引擎有跑（runs>0）；真實稽核表近 30 天不一致／例外／熔斷＝0
 *   3) on：資料有 approval.flow 快照；任何 API 回應都看不到 flow／defHash；關卡與 off 相同
 *   4) on 送簽後把伺服器改成 off 重開（凍結快照語意）：在途單照原 steps 核准到完成，快照仍在
 *   5) 無法辨識的環境變數值（banana）＝off，狀態端點標示 envValid=false
 *   7) 模式變化稽核（QUOTE_FLOW_ENGINE_MODE）：off 不記；shadow／on 第一個請求記 off→shadow／off→on；on 改回 off 重開後記 on→off（熔斷器使用留痕）
 *   8) 商機作廢：on 模式送簽的單，連結的商機被作廢後 approval.flow 一併清掉
 *   6) 後台頁：admin.html 含「流程引擎狀態」區塊與 3 個稽核動作中文名；內嵌 script 語法正確
 */
'use strict';
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.QFE_PORT || '3120', 10);
if ([3000, 3020, 3030, 3101, 3111, 3121, 3131, 3141, 3151, 3161].includes(PORT)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
const guardFiles = FILES.filter((f) => /(auth|data|audit.log).jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
function ok(cond, name, extra) { if (cond) pass++; else { fail++; console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const sleep = (n) => new Promise((r) => setTimeout(r, n));
const PW = 'Qf1' + crypto.randomBytes(9).toString('hex') + '!Z';   // 只在記憶體
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});

function seed() {
  const users = [
    mkUser('_s_admin', 'admin'), mkUser('_s_own', 'user', { bu: ['ITS'], supervisor: '_s_m2' }), mkUser('_s_m2', 'manager2', { bu: ['ITS'], supervisor: '_s_m1' }),
    mkUser('_s_m1', 'manager1', { bu: ['ITS'] }), mkUser('_s_gm', 'executive'), mkUser('_s_ch', 'executive'),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const mkQ = (id, price, cost) => ({ id, quoteNo: 'SB-' + id, company: '沙盒公司甲', projectName: '沙盒專案', owner: '_s_own', status: 'draft', createdAt: nowIso, updatedAt: nowIso, quoteDate: nowIso.slice(0, 10), validUntil: '2099-12-31', products: ['PS'], items: [{ lid: id + '-l1', desc: '服務', unit: '式', qty: 1, unitPrice: price, cost }], discountType: 'none', discountValue: 0, approval: null });
  const d = {
    contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [],
    opportunities: [{ id: 'opp1', owner: '_s_own', stage: 'B', company: '沙盒公司甲', product: '沙盒商品', amount: 800 }],
    quotations: [mkQ('q1', 800000, 400000), mkQ('q2', 12000000, 6000000), Object.assign(mkQ('q3', 800000, 400000), { oppId: 'opp1' })],
    quoteApproval: { roster: { gm: ['_s_gm'], chairman: ['_s_ch'], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: { PS: { cls: 'software', costBySales: true } } },
    approvalFlowConfig: { version: 1, buSettings: { ITS: { hasMgr2Step: true } }, exceptions: [] },
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
const cleanup = () => FILES.forEach((f) => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));

let child = null, childPid = null, jar = {};
async function startServer(engine) {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE; delete env.VERCEL; delete env.AUDIT_EVENTS_TABLE; delete env.QUOTE_FLOW_ENGINE;
  if (engine !== undefined) env.QUOTE_FLOW_ENGINE = engine;
  await assertPortFree(PORT);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
  child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) { jar = {}; return; } } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-500));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-500));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 30 && child.exitCode === null; i++) await sleep(100);
  child = null;
  await sleep(300);
}
async function http(who, method, url, body) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const text = await res.text();
  let json = null; try { json = JSON.parse(text); } catch { /* 非 JSON */ }
  return { status: res.status, json, text, setCookie: res.headers.getSetCookie ? res.headers.getSetCookie() : [] };
}
async function login(who) {
  const r = await http(null, 'POST', '/api/login', { username: who, password: PW });
  if (r.status === 200) jar[who] = r.setCookie.map((c) => c.split(';')[0]).join('; ');
  return r.status === 200;
}
const stepTiers = (q) => ((q.approval && q.approval.steps) || []).map((s) => s.tier).join('>');
const LEAK = /defHash|"flow"|"trace"|qfe-1|routeId|factsPublic|legacyCfg/;

async function phase(name, engine, fn) {
  console.log('[' + name + ']');
  seed();
  try { await startServer(engine); for (const u of ['_s_admin', '_s_own', '_s_m2', '_s_m1', '_s_gm', '_s_ch']) ok(await login(u), name + '：登入 ' + u); await fn(); }
  finally { await stopServer(); }
}
const status = async () => (await http('_s_admin', 'GET', '/api/admin/quote-flow/engine-status'));
const submit = async (id) => http('_s_own', 'POST', `/api/quotations/${id}/submit`, {});
async function approveAll(id, expectTiers) {
  const got = [];
  for (let i = 0; i < 6; i++) {
    const q = readData().quotations.find((x) => x.id === id);
    if (!q.approval || q.approval.state !== 'pending') break;
    const step = q.approval.steps[q.approval.cur];
    const who = { mgr2: '_s_m2', mgr1: '_s_m1', gm: '_s_gm', chairman: '_s_ch' }[step.tier];
    const r = await http(who, 'POST', `/api/quotations/${id}/approve`, { hash: q.approval.hash, comment: 'ok' });
    got.push(step.tier + ':' + r.status);
    if (r.status !== 200) break;
  }
  return got;
}

(async () => {
  try {
    // 1) off（未設定）
    let offTiers = null;
    await phase('off（未設定環境變數）', undefined, async () => {
      const r = await submit('q1');
      ok(r.status === 200 && r.json && r.json.approval && r.json.approval.state === 'pending', 'off：送簽 200', r.status + ' ' + r.text.slice(0, 200));
      const q = readData().quotations.find((x) => x.id === 'q1');
      offTiers = stepTiers(q);
      ok(offTiers === 'mgr2>mgr1', 'off：q1 關卡＝mgr2>mgr1（ITS 啟用二級主管）', offTiers);
      ok(!('flow' in q.approval), 'off：資料沒有 approval.flow');
      const st = await status();
      ok(st.status === 200 && st.json.mode === 'off' && st.json.stats.runs === 0, 'off：狀態端點 mode=off、引擎沒跑', st.text.slice(0, 200));
      ok(st.json.defaultFlow.structureHash === st.json.defaultFlow.pinnedHash, 'off：預設流程結構雜湊與釘選值一致');
      ok(st.json.audit.mismatch === 0 && st.json.audit.error === 0 && st.json.audit.fallback === 0, 'off：真實稽核表近 30 天筆數＝0（可讀取）', JSON.stringify(st.json.audit));
      const anon = await http(null, 'GET', '/api/admin/quote-flow/engine-status');
      const own = await http('_s_own', 'GET', '/api/admin/quote-flow/engine-status');
      ok((anon.status === 401 || anon.status === 403) && own.status === 403, 'off：未登入／非管理員不可讀狀態端點', anon.status + '/' + own.status);
      ok(!LEAK.test(r.text), 'off：送簽回應不含 flow 相關字樣');
      const offLogs = await http('_s_admin', 'GET', '/api/admin/logs');
      ok(Array.isArray(offLogs.json) && !offLogs.json.some((l) => /^QUOTE_FLOW_/.test(l.action)), 'off：稽核日誌沒有任何 QUOTE_FLOW_* 事件（含 ENGINE_MODE）', offLogs.status);
      ok(!('quoteFlowEngineMode' in readData()), 'off：資料裡沒有 quoteFlowEngineMode（預設行為不變）');
    });
    // 2) shadow
    await phase('shadow', 'shadow', async () => {
      const r1 = await submit('q1'), r2 = await submit('q2');
      const d = readData();
      const q1 = d.quotations.find((x) => x.id === 'q1'), q2 = d.quotations.find((x) => x.id === 'q2');
      ok(r1.status === 200 && r2.status === 200, 'shadow：兩張單都送簽 200', r1.status + '/' + r2.status);
      ok(stepTiers(q1) === offTiers && stepTiers(q2) === 'mgr2>mgr1>gm>chairman', 'shadow：關卡與 off 相同（q1），q2（1,200 萬）加董事長', stepTiers(q1) + ' | ' + stepTiers(q2));
      ok(!('flow' in q1.approval) && !('flow' in q2.approval), 'shadow：資料沒有 approval.flow（不改寫報價單）');
      const st = await status();
      ok(st.json.mode === 'shadow' && st.json.stats.runs >= 2 && st.json.stats.mismatches === 0 && st.json.stats.errors === 0, 'shadow：引擎有跑且零差異零例外', JSON.stringify(st.json.stats));
      ok(st.json.audit.mismatch === 0 && st.json.audit.error === 0 && st.json.audit.fallback === 0, 'shadow：真實稽核表近 30 天不一致／例外／熔斷＝0', JSON.stringify(st.json.audit));
      ok(!LEAK.test(r1.text + r2.text), 'shadow：回應不含 flow 相關字樣');
      const logs = await http('_s_admin', 'GET', '/api/admin/logs');
      ok(Array.isArray(logs.json) && !logs.json.some((l) => /^QUOTE_FLOW_/.test(l.action) && l.action !== 'QUOTE_FLOW_ENGINE_MODE'), 'shadow：稽核日誌沒有不一致／例外／熔斷事件');
      const ml = (logs.json || []).filter((l) => l.action === 'QUOTE_FLOW_ENGINE_MODE');
      ok(ml.length === 1 && /off→shadow/.test(ml[0].detail) && ml[0].target === 'QUOTE_FLOW_ENGINE', 'shadow：第一個請求記一筆 QUOTE_FLOW_ENGINE_MODE（off→shadow），整段只有一筆', JSON.stringify(ml.map((l) => l.detail)).slice(0, 300));
      ok(readData().quoteFlowEngineMode && readData().quoteFlowEngineMode.mode === 'shadow', 'shadow：資料記下目前模式');
    });
    // 3) on，再 4) 改回 off 重開
    seed();
    console.log('[on → 改回 off]');
    let flowSnap = null, stepsAfterSubmit = null;
    try {
      await startServer('on');
      for (const u of ['_s_admin', '_s_own', '_s_m2', '_s_m1', '_s_gm', '_s_ch']) ok(await login(u), 'on：登入 ' + u);
      const r1 = await submit('q1'), r2 = await submit('q2'), r3 = await submit('q3');
      const d = readData();
      const q1 = d.quotations.find((x) => x.id === 'q1'), q2 = d.quotations.find((x) => x.id === 'q2');
      ok(r1.status === 200 && r2.status === 200 && r3.status === 200, 'on：三張單都送簽 200', r1.status + '/' + r2.status + '/' + r3.status);
      ok(!!d.quotations.find((x) => x.id === 'q3').approval.flow, 'on：q3（連結商機 opp1）送簽後有 flow 快照');
      const del = await http('_s_own', 'DELETE', '/api/opportunities/opp1', { deleteReason: '價格：沙盒測試' });
      const q3b = readData().quotations.find((x) => x.id === 'q3');
      ok(del.status === 200 && q3b.approval.state === 'none' && !('flow' in q3b.approval) && q3b.approval.derived === null && q3b.invalid && q3b.invalid.reason === 'OPP_VOID', 'on：商機作廢 → q3 簽核關閉且 approval.flow 一併清掉（state=none、derived=null、q.invalid=OPP_VOID）', del.status + ' ' + JSON.stringify(q3b.approval).slice(0, 200));
      const onLogs = await http('_s_admin', 'GET', '/api/admin/logs');
      // 稽核檔跨階段累積（只有資料檔每階段重置）：這裡只數 off→on
      const mlOn = (onLogs.json || []).filter((l) => l.action === 'QUOTE_FLOW_ENGINE_MODE' && /off→on/.test(l.detail));
      ok(mlOn.length === 1, 'on：第一個請求記一筆 QUOTE_FLOW_ENGINE_MODE（off→on），整段只有一筆', JSON.stringify(mlOn.map((l) => l.detail)).slice(0, 300));
      ok(stepTiers(q1) === offTiers && stepTiers(q2) === 'mgr2>mgr1>gm>chairman', 'on：關卡與 off／shadow 相同', stepTiers(q1) + ' | ' + stepTiers(q2));
      ok(q1.approval.flow && q1.approval.flow.engine === 'qfe-1' && q1.approval.flow.origin === 'legacy' && q1.approval.flow.tiers.join('>') === offTiers, 'on：資料有 approval.flow 快照（tiers 與 steps 相同）', JSON.stringify(q1.approval.flow).slice(0, 200));
      ok(q1.approval.history[0].meta && q1.approval.history[0].meta.flow && q1.approval.history[0].meta.flow.defHash === q1.approval.flow.defHash, 'on：SUBMIT 歷程 meta 有 flow 摘要');
      const get = await http('_s_own', 'GET', '/api/quotations/q1'), admGet = await http('_s_admin', 'GET', '/api/quotations/q1'), list = await http('_s_own', 'GET', '/api/quotations');
      ok(!LEAK.test(r1.text + r2.text + get.text + admGet.text + list.text), 'on：送簽／單筆／清單回應都不含 flow 相關字樣');
      const st = await status();
      ok(st.json.mode === 'on' && st.json.stats.runs >= 2 && st.json.stats.mismatches === 0 && st.json.stats.fallbacks === 0, 'on：狀態端點 mode=on、零差異零熔斷', JSON.stringify(st.json.stats));
      flowSnap = JSON.stringify(q1.approval.flow); stepsAfterSubmit = JSON.stringify(q1.approval.steps.map((s) => [s.tier, s.assignee]));
      await stopServer();
      // 改回 off（熔斷器）並重開：在途單照 steps 走
      await startServer('off');
      for (const u of ['_s_admin', '_s_m2', '_s_m1']) ok(await login(u), 'off（重開）：登入 ' + u);
      const got = await approveAll('q1');
      const q1b = readData().quotations.find((x) => x.id === 'q1');
      ok(got.join(',') === 'mgr2:200,mgr1:200' && q1b.approval.state === 'approved', 'off（重開）：在途單（on 時送簽）照原 steps 核准到完成', got.join(','));
      ok(JSON.stringify(q1b.approval.flow) === flowSnap && JSON.stringify(q1b.approval.steps.map((s) => [s.tier, s.assignee])) === stepsAfterSubmit, 'off（重開）：flow 快照與 steps 的關卡／指派人完全不變');
      const st2 = await status();
      ok(st2.json.mode === 'off' && st2.json.stats.runs === 0, 'off（重開）：狀態端點 mode=off、引擎沒跑');
      const offLogs2 = await http('_s_admin', 'GET', '/api/admin/logs');
      const mlOff = (offLogs2.json || []).filter((l) => l.action === 'QUOTE_FLOW_ENGINE_MODE');
      ok(mlOff.filter((l) => /on→off/.test(l.detail)).length === 1 && mlOff.filter((l) => /off→on/.test(l.detail)).length === 1, 'off（重開）：熔斷器使用留痕——改回 off 後第一個請求記一筆 on→off（連同先前的 off→on）', JSON.stringify(mlOff.map((l) => l.detail)).slice(0, 400));
      ok(readData().quoteFlowEngineMode.mode === 'off', 'off（重開）：資料記下目前模式 off');
    } finally { await stopServer(); cleanup(); }
    // 5) 亂碼環境變數
    await phase('banana', 'banana', async () => {
      const r = await submit('q1');
      const q = readData().quotations.find((x) => x.id === 'q1');
      const st = await status();
      ok(r.status === 200 && !('flow' in q.approval) && stepsOk(q), 'banana：視為 off，送簽正常且沒有 flow');
      ok(st.json.mode === 'off' && st.json.envValid === false && st.json.envRaw === 'banana' && st.json.stats.runs === 0, 'banana：狀態端點 envValid=false', JSON.stringify([st.json.mode, st.json.envValid, st.json.envRaw]));
    });
    function stepsOk(q) { return stepTiers(q) === 'mgr2>mgr1'; }
    // 6) 後台頁
    const html = fs.readFileSync(path.join(ROOT, '_client/admin.html'), 'utf8');
    ok(/id="afEngineCard"/.test(html) && /id="afEngineBody"/.test(html) && /loadFlowEngineStatus\(\);/.test(html) && /\/api\/admin\/quote-flow\/engine-status/.test(html), '後台頁：有流程引擎狀態區塊與載入呼叫');
    ok(['QUOTE_FLOW_SHADOW_MISMATCH', 'QUOTE_FLOW_SHADOW_ERROR', 'QUOTE_FLOW_ENGINE_FALLBACK'].every((a) => new RegExp(a + ": '[^']+'").test(html)), '後台頁：3 個稽核動作都有中文名（ACTION_LABELS）');
    const secStart = html.indexOf('id="afEngineCard"');
    // 區塊結束：同頁後面若有其他卡片（例如防拆單 sgCard）就以它為界，否則取原本的結尾標記
    const nextCardAt = html.indexOf('id="sgCard"', secStart), nextCard = nextCardAt > 0 ? html.lastIndexOf('<div', nextCardAt) : -1, closeTag = html.indexOf('</div>\n      </div>', secStart);
    const secEnd = (nextCard > 0 && nextCard < closeTag) ? nextCard : closeTag;
    ok(secStart > 0 && !/style=/.test(html.slice(secStart, secEnd + 20)), '後台頁：新區塊 HTML 不使用 style= 屬性');
    const fnStart = html.indexOf('async function loadFlowEngineStatus'), fnEnd = html.indexOf('function _afRender()');
    const fnSrc = html.slice(fnStart, fnEnd);
    ok(fnStart > 0 && !/style=/.test(fnSrc) && /escH\(/.test(fnSrc), '後台頁：新 JS 不產生 style=、動態文字都經 escH');
    const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g; let m, n = 0, bad = 0;
    while ((m = re.exec(html))) { n++; try { new Function(m[1]); } catch (e) { bad++; console.log('  inline script syntax:', e.message); } }
    ok(n > 0 && bad === 0, '後台頁：內嵌 script 語法（' + n + ' 段）');
  } catch (e) {
    fail++; console.log('  ✗ 例外：', e && e.stack || e);
  } finally {
    await stopServer(); cleanup();
  }
  console.log(`\n簽核流程引擎 E2E：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
})();
