#!/usr/bin/env node
/**
 * 客戶健康分數索引化（M-02）檢查。用法：node scripts/check-contacts-health-perf.js
 * 動 lib/contactHealth.js 或 server.js 的 GET /api/contacts 健康分數段落之後必跑。
 *
 * 1) 逐筆等價：REFERENCE_ORIGINAL（取自索引化之前 server.js 的逐字邏輯，只把 Date.now()／new Date() 換成固定時間）
 *    與 lib/contactHealth.js 在 2,000 筆（多組種子）與 10,000 筆資料上逐筆比對 score／icon，0 差異。
 *    資料刻意含髒值：公司名前後空白、空公司、聯絡人沒姓名、帳款沒 contactName、visitDate 缺漏／null／空字串、
 *    已刪除紀錄、owner 缺漏、剛好落在 30／60／90 天與 7 天邊界。
 * 2) 既有行為鎖定（手算期望值，原版與新版都要得到）：
 *    - 兩邊 name／contactName 都是 undefined 視為相符（沒有姓名的聯絡人被同業務的逾期帳款扣 30）
 *    - 有一筆 visitDate 缺漏 → NaN → 不扣分（與「從沒拜訪」扣 40 不同）
 *    - 公司名 trim 比對；已刪除拜訪不算；本年成交不看 deleted
 * 3) 效能：10,000 筆索引化版在時間上限內（HP_MAX_MS，預設 1500；實測約數十毫秒），2,000 筆至少比原版快 HP_MIN_SPEEDUP 倍（預設 5）。
 * 4) 接線：server.js 有 require lib/contactHealth 並在 GET /api/contacts 內使用，且舊的 O(n²) 寫法（visits90.filter）不在。
 * 5) 變異驗證：對 lib/contactHealth.js 逐一植入錯誤（見 MUTANTS），每個都必須被 1)、2) 抓到。
 *
 * 環境變數：HP_SKIP_BIG=1 略過 10,000 筆原版比對（原版要 7～10 秒）；HP_MAX_MS；HP_MIN_SPEEDUP。
 * 不啟動伺服器、不讀寫任何資料檔。
 */
const fs = require('fs');
const path = require('path');
const Module = require('module');
const ROOT = path.join(__dirname, '..');
const LIB = path.join(ROOT, 'lib/contactHealth.js');
const DAY = 86400000;
const NOW = Date.parse('2026-10-21T12:00:00.000Z');   // 固定時間：邊界資料才可重現
const TODAY = new Date(NOW).toISOString().slice(0, 10);

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);

// ── 參考實作：索引化之前 server.js GET /api/contacts 的健康分數邏輯（逐字；僅 Date.now()→NOW、new Date()→new Date(NOW)）──
function REFERENCE_ORIGINAL(contacts, data) {
  const now = NOW;
  const todayStr = new Date(NOW).toISOString().slice(0, 10);
  const yearNum = new Date(NOW).getFullYear();
  const visits90 = (data.visits || []).filter(v => !v.deleted);
  const opps = data.opportunities || [];
  const recvs = data.receivables || [];
  const contracts = data.contracts || [];

  return contacts.map(c => {
    let score = 100;
    const co = (c.company || '').trim();
    const cId = c.id;

    // 最後拜訪
    const cVisits = visits90.filter(v =>
      v.contactId === cId || (co && (v.company || '').trim() === co)
    );
    const lastVisitDays = cVisits.length
      ? Math.floor((now - new Date(cVisits.map(v => v.visitDate).sort().reverse()[0]).getTime()) / 86400000)
      : Infinity;
    if (lastVisitDays > 60) score -= 40;
    else if (lastVisitDays > 30) score -= 20;

    // 逾期帳款
    const hasOverdueRecv = recvs.some(r =>
      r.owner === c.owner &&
      ((co && (r.company || '').trim() === co) || r.contactName === c.name) &&
      r.status !== 'paid' && r.dueDate && r.dueDate < todayStr
    );
    if (hasOverdueRecv) score -= 30;

    // 合約即將到期（7天內，無 renewDate）
    const in7 = new Date(NOW); in7.setDate(in7.getDate() + 7);
    const hasExpiringContract = contracts.some(ct =>
      ct.owner === c.owner && !ct.deleted && !ct.renewDate &&
      ct.endDate && ct.endDate >= todayStr && ct.endDate <= in7.toISOString().slice(0, 10) &&
      ((co && (ct.company || '').trim() === co))
    );
    if (hasExpiringContract) score -= 20;

    // 近 90 天有新商機
    const d90 = new Date(now - 90 * 86400000).toISOString();
    const hasNewOpp = opps.some(o =>
      o.owner === c.owner && !o.deleted && o.createdAt >= d90 &&
      (o.contactId === cId || (co && (o.company || '').trim() === co))
    );
    if (hasNewOpp) score += 10;

    // 本年有成交商機
    const hasWonThis = opps.some(o =>
      o.owner === c.owner && o.stage === 'Won' &&
      new Date(o.achievedDate || o.updatedAt || o.createdAt).getFullYear() === yearNum &&
      (o.contactId === cId || (co && (o.company || '').trim() === co))
    );
    if (hasWonThis) score += 20;

    score = Math.max(0, Math.min(120, score));
    const health = score >= 70 ? '🟢' : score >= 40 ? '🟡' : '🔴';
    return { id: c.id, healthScore: score, healthIcon: health };
  });
}

// ── 載入被測模組（可載入變異版原始碼）──
function loadLibFromSource(src, file) {
  const m = new Module(file, module);
  m.filename = file; m.paths = Module._nodeModulePaths(path.dirname(file));
  m._compile(src, file);
  return m.exports;
}
function runLib(lib, contacts, data) {
  const f = lib.buildHealthScorer(data, NOW);
  return contacts.map(c => { const h = f(c); return { id: c.id, healthScore: h.score, healthIcon: h.icon }; });
}

// ── 資料產生（固定種子）──
function rng(seed) { let a = seed >>> 0; return () => { a = (a + 0x6D2B79F5) >>> 0; let x = a; x = Math.imul(x ^ (x >>> 15), x | 1); x ^= x + Math.imul(x ^ (x >>> 7), x | 61); return ((x ^ (x >>> 14)) >>> 0) / 4294967296; }; }
function gen(N, seed) {
  const R = rng(seed); const ri = (n) => Math.floor(R() * n); const pick = (a) => a[ri(a.length)];
  const ncomp = Math.max(3, Math.round(N / 3));
  const owners = Array.from({ length: 20 }, (_, i) => 'u' + i);
  const ownerOf = () => (R() < 0.01 ? undefined : pick(owners));
  const dstr = (off) => new Date(NOW + off * DAY).toISOString().slice(0, 10);
  const coName = () => { const r = R(); const base = '沙盒公司' + ri(ncomp); return r < 0.02 ? '' : r < 0.03 ? undefined : r < 0.06 ? ' ' + base + ' ' : base; };
  const contacts = Array.from({ length: N }, (_, i) => ({
    id: 'c' + i, owner: ownerOf(), name: R() < 0.03 ? undefined : '聯絡人' + ri(Math.max(2, Math.round(N / 2))), company: coName(),
  }));
  const visits = Array.from({ length: N }, () => {
    const r = R();
    const vd = r < 0.02 ? undefined : r < 0.03 ? null : r < 0.04 ? '' : dstr(-ri(121));
    const cid = R() < 0.33 ? (R() < 0.8 ? 'c' + ri(N) : pick([undefined, '', 'cX'])) : undefined;
    return { contactId: cid, company: coName(), visitDate: vd, deleted: R() < 0.05 ? true : undefined };
  });
  const opps = Array.from({ length: Math.round(N / 2) }, () => {
    const off = ri(200);
    return {
      owner: ownerOf(), contactId: R() < 0.5 ? 'c' + ri(N) : undefined, company: coName(), stage: pick(['A', 'B', 'Won', 'Won']),
      createdAt: R() < 0.01 ? undefined : new Date(NOW - (R() < 0.1 ? 90 : off) * DAY).toISOString(),
      achievedDate: R() < 0.5 ? undefined : new Date(NOW - ri(500) * DAY).toISOString(), updatedAt: R() < 0.3 ? new Date(NOW - ri(500) * DAY).toISOString() : undefined,
      deleted: R() < 0.05 ? true : undefined,
    };
  });
  const recvs = Array.from({ length: Math.round(N / 10) }, () => ({
    owner: ownerOf(), company: coName(), contactName: R() < 0.5 ? undefined : '聯絡人' + ri(Math.max(2, Math.round(N / 2))),
    status: pick(['open', 'paid', 'open']), dueDate: R() < 0.03 ? undefined : dstr(ri(61) - 30),
  }));
  const contracts = Array.from({ length: Math.round(N / 10) }, () => ({
    owner: ownerOf(), company: coName(), endDate: R() < 0.03 ? undefined : dstr(ri(25) - 5), renewDate: R() < 0.2 ? '2027-01-01' : undefined, deleted: R() < 0.05 ? true : undefined,
  }));
  return { contacts, data: { visits, opportunities: opps, receivables: recvs, contracts } };
}
function diffCount(A, B) { let d = 0; if (A.length !== B.length) return Infinity; for (let i = 0; i < A.length; i++) if (A[i].id !== B[i].id || A[i].healthScore !== B[i].healthScore || A[i].healthIcon !== B[i].healthIcon) d++; return d; }

// ── 既有行為鎖定：手算期望值 ──
function directedCases() {
  const base = { visits: [], opportunities: [], receivables: [], contracts: [] };
  const mk = (extra) => Object.assign({}, base, extra);
  const C = (o) => Object.assign({ id: 'c1', owner: 'u1', name: '甲', company: '甲公司' }, o);
  const recent = TODAY;                         // 今天拜訪 → 0 天，不扣分
  return [
    ['從沒拜訪 → 扣 40', [C()], mk({}), 60],
    ['近期拜訪 → 100', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }] }), 100],
    ['拜訪 31～60 天 → 扣 20', [C()], mk({ visits: [{ contactId: 'c1', visitDate: new Date(NOW - 45 * DAY).toISOString().slice(0, 10) }] }), 80],
    ['visitDate 缺漏 → NaN → 不扣分（鎖定既有行為）', [C()], mk({ visits: [{ contactId: 'c1' }] }), 100],
    ['同公司另一筆 visitDate 缺漏也讓整組變 NaN（排序把 undefined 放最後）', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }, { company: ' 甲公司 ' }] }), 100],
    ['已刪除拜訪不算 → 視同從沒拜訪', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent, deleted: true }] }), 60],
    ['公司名 trim 後比對拜訪', [C({ company: '  甲公司' })], mk({ visits: [{ company: '甲公司 ', visitDate: recent }] }), 100],
    ['沒有姓名的聯絡人 + 同業務沒填 contactName 的逾期帳款 → 扣 30（兩邊 undefined 相符，鎖定既有行為）', [C({ name: undefined, company: '', id: 'c1' })], mk({ visits: [{ contactId: 'c1', visitDate: recent }], receivables: [{ owner: 'u1', status: 'open', dueDate: '2026-01-01' }] }), 70],
    ['有姓名時，沒填 contactName 的帳款不相符 → 不扣', [C({ company: '' })], mk({ visits: [{ contactId: 'c1', visitDate: recent }], receivables: [{ owner: 'u1', status: 'open', dueDate: '2026-01-01' }] }), 100],
    ['逾期帳款以公司（trim）比對 → 扣 30', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], receivables: [{ owner: 'u1', company: ' 甲公司', status: 'open', dueDate: '2026-01-01' }] }), 70],
    ['已付清帳款不扣', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], receivables: [{ owner: 'u1', company: '甲公司', status: 'paid', dueDate: '2026-01-01' }] }), 100],
    ['帳款 owner 不同不扣', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], receivables: [{ owner: 'u2', company: '甲公司', status: 'open', dueDate: '2026-01-01' }] }), 100],
    ['合約 7 天內到期且無續約 → 扣 20', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], contracts: [{ owner: 'u1', company: '甲公司', endDate: new Date(NOW + 7 * DAY).toISOString().slice(0, 10) }] }), 80],
    ['合約第 8 天到期不扣', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], contracts: [{ owner: 'u1', company: '甲公司', endDate: new Date(NOW + 8 * DAY).toISOString().slice(0, 10) }] }), 100],
    ['已刪除合約不扣', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], contracts: [{ owner: 'u1', company: '甲公司', deleted: true, endDate: TODAY }] }), 100],
    ['近 90 天新商機 +10（上限 120 前）', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], opportunities: [{ owner: 'u1', contactId: 'c1', stage: 'A', createdAt: new Date(NOW - 10 * DAY).toISOString() }] }), 110],
    ['商機恰好 90 天前（>=）仍算新商機', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], opportunities: [{ owner: 'u1', contactId: 'c1', stage: 'A', createdAt: new Date(NOW - 90 * DAY).toISOString() }] }), 110],
    ['已刪除商機不算新商機', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], opportunities: [{ owner: 'u1', contactId: 'c1', stage: 'A', deleted: true, createdAt: new Date(NOW - 10 * DAY).toISOString() }] }), 100],
    ['本年成交 +20；已刪除的成交仍算（原版不看 deleted，鎖定既有行為）', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], opportunities: [{ owner: 'u1', contactId: 'c1', stage: 'Won', deleted: true, createdAt: '2020-01-01T00:00:00.000Z', achievedDate: '2026-03-01T00:00:00.000Z' }] }), 120],
    ['去年成交不加分', [C()], mk({ visits: [{ contactId: 'c1', visitDate: recent }], opportunities: [{ owner: 'u1', contactId: 'c1', stage: 'Won', createdAt: '2020-01-01T00:00:00.000Z', achievedDate: '2025-03-01T00:00:00.000Z' }] }), 100],
    ['新商機 + 成交 + 帳款逾期 + 無拜訪 → 100-40-30+10+20=60', [C()], mk({ receivables: [{ owner: 'u1', company: '甲公司', status: 'open', dueDate: '2026-01-01' }], opportunities: [{ owner: 'u1', contactId: 'c1', stage: 'Won', createdAt: new Date(NOW - 5 * DAY).toISOString(), achievedDate: new Date(NOW - 5 * DAY).toISOString() }] }), 60],
  ];
}

// ── 變異：對 lib 原始碼逐一植入錯誤，必須被抓到 ──
const MUTANTS = [
  ['公司名不 trim', "typeof s === 'string' ? s.trim() : ''", "typeof s === 'string' ? s : ''"],
  ['帳款姓名改成「兩邊 undefined 不相符」', 'addTo(recvName, r.owner, r.contactName);', 'if (r.contactName !== undefined) addTo(recvName, r.owner, r.contactName);'],
  ['拜訪日缺漏不再讓整組 NaN', 'if (d === undefined) { e.u = true; return; }', "if (d === undefined) { d = '0000'; }"],
  ['拜訪日取最小而非最大', 'if (!e.has || s >= e.s)', 'if (!e.has || s <= e.s)'],
  ['已刪除拜訪也算', 'if (v.deleted) continue;\n    keepVisit', 'keepVisit'],
  ['拜訪 60 天門檻 > 改 >=', 'lastVisitDays > 60', 'lastVisitDays >= 60'],
  ['拜訪 30 天門檻 > 改 >=', 'lastVisitDays > 30', 'lastVisitDays >= 30'],
  ['帳款少判 owner', 'addTo(recvCo, r.owner, rc)', 'addTo(recvCo, "*", rc)'],
  ['已付清帳款也扣', "r.status !== 'paid' && ", ''],
  ['合約已刪除也算', 'if (ct.deleted || ct.renewDate || !ct.endDate) continue;', 'if (ct.renewDate || !ct.endDate) continue;'],
  ['合約 7 天改 8 天', 'in7.getDate() + 7', 'in7.getDate() + 8'],
  ['新商機 90 天改 89 天', 'now - 90 * DAY', 'now - 89 * DAY'],
  ['新商機邊界 >= 改 >', 'o.createdAt >= d90', 'o.createdAt > d90'],
  ['新商機少判 owner', 'addTo(newId, o.owner, o.contactId)', 'addTo(newId, "*", o.contactId)'],
  ['成交少判年份', '.getFullYear() === yearNum', '.getFullYear() >= 1970'],
  ['成交也看 deleted', "if (o.stage === 'Won' &&", "if (!o.deleted && o.stage === 'Won' &&"],
  ['上限 120 改 100', 'Math.min(120, s)', 'Math.min(100, s)'],
  ['綠燈門檻 70 改 60', 's >= 70', 's >= 60'],
  ['黃燈門檻 40 改 50', 's >= 40', 's >= 50'],
  ['公司比對拜訪漏掉公司索引', 'const a = visitById.get(cId), b = co ? visitByCo.get(co) : undefined;', 'const a = visitById.get(cId), b = undefined;'],
];

(async () => {
  const libSrc = fs.readFileSync(LIB, 'utf8');
  const lib = loadLibFromSource(libSrc, LIB);

  // 資料集與原版結果只算一次（原版 10,000 筆要 7～10 秒）
  const sets = [];
  for (const seed of [11, 22, 33]) sets.push({ name: `2,000 筆 seed=${seed}`, N: 2000, ...gen(2000, seed) });
  if (!process.env.HP_SKIP_BIG) sets.push({ name: '10,000 筆 seed=7', N: 10000, ...gen(10000, 7) });
  let origMsAt2000 = 0, libMsAt2000 = 0;
  for (const s of sets) {
    const t0 = process.hrtime.bigint(); s.ref = REFERENCE_ORIGINAL(s.contacts, s.data); s.refMs = Number(process.hrtime.bigint() - t0) / 1e6;
    if (s.N === 2000 && !origMsAt2000) origMsAt2000 = s.refMs;
  }
  const dir = sets[0].ref.reduce((m, x) => { m[x.healthScore] = (m[x.healthScore] || 0) + 1; return m; }, {});
  t('資料集有足夠的分數分布（≥8 種不同分數，才有鑑別力）', Object.keys(dir).length >= 8, Object.keys(dir).length + ' 種');

  const suite = (L) => {
    const fails = [];
    for (const s of sets) { const d = diffCount(s.ref, runLib(L, s.contacts, s.data)); if (d) fails.push(`${s.name} 差異 ${d}`); }
    for (const [name, cs, data, expect] of directedCases()) {
      const got = runLib(L, cs, data)[0].healthScore;
      if (got !== expect) fails.push(`手算案例「${name}」得到 ${got}，預期 ${expect}`);
    }
    return fails;
  };

  // 1) 逐筆等價
  for (const s of sets) {
    const t0 = process.hrtime.bigint(); const out = runLib(lib, s.contacts, s.data); const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    if (s.N === 2000 && !libMsAt2000) libMsAt2000 = ms;
    const d = diffCount(s.ref, out);
    t(`逐筆等價：${s.name}（原版 ${s.refMs.toFixed(0)} ms、索引化 ${ms.toFixed(0)} ms）0 差異`, d === 0, d ? '差異 ' + d : '');
    if (s.N === 10000) t(`效能：10,000 筆索引化版 < ${process.env.HP_MAX_MS || 1500} ms`, ms < Number(process.env.HP_MAX_MS || 1500), ms.toFixed(0) + ' ms');
  }
  // 2) 手算案例：原版與新版都要對（先證明參考實作本身符合期望，案例才有意義）
  for (const [name, cs, data, expect] of directedCases()) {
    const r = REFERENCE_ORIGINAL(cs, data)[0].healthScore, n = runLib(lib, cs, data)[0].healthScore;
    t(`既有行為：${name}`, r === expect && n === expect, `原版 ${r}／新版 ${n}／預期 ${expect}`);
  }
  // 3) 2,000 筆加速倍數（取 3 次最快，降低抖動）
  {
    const s = sets[0]; let best = Infinity;
    for (let i = 0; i < 3; i++) { const t0 = process.hrtime.bigint(); runLib(lib, s.contacts, s.data); best = Math.min(best, Number(process.hrtime.bigint() - t0) / 1e6); }
    const need = Number(process.env.HP_MIN_SPEEDUP || 5);
    t(`效能：2,000 筆至少比原版快 ${need} 倍`, s.refMs / Math.max(best, 0.01) >= need, `原版 ${s.refMs.toFixed(0)} ms／新版 ${best.toFixed(1)} ms = ${(s.refMs / Math.max(best, 0.01)).toFixed(0)} 倍`);
  }
  // 4) 接線
  {
    const sv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8').replace(/\r\n/g, '\n');
    const a = sv.indexOf("app.get('/api/contacts', requireAuth");
    const b = sv.indexOf("app.post('/api/contacts', requireAuth");
    const seg = a >= 0 && b > a ? sv.slice(a, b) : '';
    t("server.js require('./lib/contactHealth')", /require\('\.\/lib\/contactHealth'\)/.test(sv));
    t('GET /api/contacts 使用 buildHealthScorer', /buildHealthScorer\(data\)/.test(seg));
    t('GET /api/contacts 不再有 O(n²) 寫法（visits90.filter／recvs.some／opps.some／contracts.some）', seg.length > 0 && !/visits90\.filter|recvs\.some|opps\.some|contracts\.some/.test(seg));
    t('lib/contactHealth.js 在 vercel 可被靜態追蹤（server.js 以字串字面量 require）', /require\('\.\/lib\/contactHealth'\)/.test(sv));
  }
  // 5) 變異
  let killed = 0;
  for (const [name, from, to] of MUTANTS) {
    const n = libSrc.split(from).length - 1;
    if (n !== 1) { t(`變異「${name}」錨點需恰好出現 1 次`, false, `出現 ${n} 次（lib 改過，請更新 MUTANTS）`); continue; }
    let fails;
    try { fails = suite(loadLibFromSource(libSrc.replace(from, to), LIB)); } catch (e) { fails = ['拋錯 ' + e.message]; }
    if (fails.length) killed++;
    t(`變異「${name}」被抓到`, fails.length > 0, fails.length ? fails[0] : '倖存！');
  }
  t(`變異全數被殺（${killed}/${MUTANTS.length}）`, killed === MUTANTS.length);

  const fail = res.filter(r => !r[1]);
  for (const [n, ok, ex] of res) console.log((ok ? '  ✓ ' : '  ✗ FAIL: ') + n + (ex ? '  — ' + ex : ''));
  console.log(`\n客戶健康分數索引化結果：${res.length - fail.length} 通過，${fail.length} 失敗`);
  process.exit(fail.length ? 1 : 0);
})().catch(e => { console.error('未預期錯誤', e); process.exit(1); });
