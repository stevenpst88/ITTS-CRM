'use strict';
/**
 * 簽核流程設計器 M2 — 後台 API（僅管理員）。純邏輯在 lib/quoteFlowDesigner.js，這裡只負責：開關、權限、儲存草稿、稽核、回應。
 *
 * 開關 QUOTE_FLOW_DESIGNER=0|1（預設 0）：每個路由的第一個 middleware 檢查，關閉時 next()——落到與沒有這些路由完全相同的 404 處理，
 * 連未登入者也看不出差別（回應與基底位元級相同）。開關只認環境變數，不讀資料庫。
 *
 * **草稿不能生效**：草稿存在獨立命名空間 data.quoteFlowDrafts，沒有任何程式（送簽、預覽、引擎）讀它；生效／版本治理／雙人覆核是 M3。
 *
 * 端點（全部 requireAdmin；寫入走 writeLog）
 *   GET    /api/admin/quote-flow/designer/status        開關／上限／引擎模式
 *   GET    /api/admin/quote-flow/designer/default-flow  目前設定轉出的預設流程（供畫布載入）＋白話檢視＋驗證
 *   POST   /api/admin/quote-flow/designer/validate      驗證一份定義（不儲存）：11 條檢查＋底線
 *   POST   /api/admin/quote-flow/designer/simulate      模擬一張假想單（可與目前生效版比較）
 *   POST   /api/admin/quote-flow/designer/sweep         邊界掃描
 *   POST   /api/admin/quote-flow/designer/overview      總覽圖資料（節點／連線／座標；body 帶編輯中的 def，可帶模擬輸入高亮）
 *   GET    /api/admin/quote-flow/designer/overview      總覽圖資料（唯讀載入：?draftId=<草稿代號> 取已儲存的草稿；不帶＝目前設定轉出的預設流程）
 *   GET    /api/admin/quote-flow/drafts                 草稿清單（不含內容）
 *   GET    /api/admin/quote-flow/drafts/:id             單份草稿（含內容）
 *   PUT    /api/admin/quote-flow/drafts                 新增／更新草稿（body.id 有值＝更新；ifUpdatedAt 做樂觀並行）
 *   DELETE /api/admin/quote-flow/drafts/:id             刪除草稿
 */
const QA = require('./quoteApproval');
const QF = require('./quoteFlow');
const FD = require('./quoteFlowDesigner');
const QM = require('./quoteMode');

module.exports = function registerFlowDesignerRoutes(app, h) {
  const { db, writeLog, qAdmin, fail, loadCtx, getFlowCfg, isActive, canApproveAct, findManager1, findManager2, nowIso } = h;

  /** 開關：關閉時直接放行給後面的處理（＝沒有這個路由，與基底相同的 404） */
  const gate = (req, res, next) => (FD.designerEnabled(process.env.QUOTE_FLOW_DESIGNER) ? next() : next('route'));
  const noStore = (res) => res.setHeader('Cache-Control', 'no-store');

  // ── 環境（名冊／簽核資格／主管鏈）：全部由伺服器現況算出，只用來驗證與模擬，不回傳帳號名單 ──
  function buildEnv(ctx) {
    const { users, cfg } = ctx;
    const secretaries = Object.values(users).filter((u) => u && u.role === 'secretary').map((u) => u.username);
    const roster = { gm: cfg.roster.gm.slice(), chairman: cfg.roster.chairman.slice(), boardProxy: [...new Set(cfg.roster.boardProxy.concat(secretaries))] };
    const signerState = (un) => {
      const u = users[un];
      if (!u) return 'missing';
      if (!isActive(u)) return 'inactive';
      if (canApproveAct(u)) return 'ok';
      if (u.role === 'admin' || u.accessMode === 'view') return 'unqualified';
      return 'readonly';
    };
    // 主管鏈解析報告（只有人數）：每位在職、能建報價單的業務是否解析得到一級／二級主管
    const owners = Object.values(users).filter((u) => u && isActive(u) && u.role !== 'admin' && !QM.NO_CREATE_ROLES.includes(u.role));
    let noMgr1 = 0, noMgr2 = 0;
    owners.forEach((u) => { if (!findManager1(users, u.username)) noMgr1++; if (!findManager2(users, u.username)) noMgr2++; });
    const flowCfg = getFlowCfg(ctx.data);
    return { roster, signerState, chain: { owners: owners.length, noMgr1, noMgr2 }, flowCfg, activeDef: QF.legacyDefault(flowCfg) };
  }
  const pub = (def) => JSON.parse(JSON.stringify(def));   // 回傳用的純複本（legacyDefault 是凍結物件）

  /** 淨化＋驗證請求中的 def。失敗回應 400 並回傳 null。 */
  function takeDef(res, raw) {
    const s = FD.sanitizeDef(raw);
    if (!s.ok) { fail(res, 400, s.code, s.message, { at: s.at }); return null; }
    return s.def;
  }
  function checksOr400(res, def, env, forSave) {
    const rc = FD.runChecks(def, env);
    if (!rc.structural) { fail(res, 400, 'FLOW_INVALID_DEF', '流程定義無效：' + rc.checks.flatMap((c) => c.issues).filter((i) => i.level !== 'warn').slice(0, 3).map((i) => i.text).join('；'), { checks: rc.checks }); return null; }
    if (forSave && !rc.ok) { fail(res, 400, 'DRAFT_INVALID', '草稿未通過驗證，不能儲存：' + rc.blocking.map((b) => '第 ' + b.n + ' 條（' + b.code + '）').join('、'), { blocking: rc.blocking, checks: rc.checks }); return null; }
    return rc;
  }

  // ── 狀態 ──
  app.get('/api/admin/quote-flow/designer/status', gate, ...qAdmin, (req, res) => {
    noStore(res);
    res.json({
      enabled: true, version: 1, engine: QF.ENGINE_ID, engineMode: QF.engineMode(process.env.QUOTE_FLOW_ENGINE),
      canActivate: false, activateNote: 'M3 開放：生效、版本治理與雙人覆核尚未提供，草稿不會影響任何簽核。',
      limits: { drafts: FD.LIMITS.drafts, name: FD.LIMITS.name, note: FD.LIMITS.note, routeName: FD.LIMITS.routeName, modName: FD.LIMITS.modName, nodeLabel: FD.LIMITS.nodeLabel, defBytes: FD.LIMITS.defBytes },
      colors: FD.COLORS, bus: FD.BUS,
      categories: QA.ROWS.map((r) => ({ key: r.key, label: r.label, l1: r.l1, l2: r.l2, fixedLevel: r.fixedLevel || null })),
      classKeys: QA.CLASS_KEYS.map((k) => ({ key: k, label: QA.CLASS_LABELS[k] })),
      floor: { chairman: String(QA.AMOUNT.chairman), board: String(QA.AMOUNT.board) },
    });
  });

  app.get('/api/admin/quote-flow/designer/default-flow', gate, ...qAdmin, (req, res) => {
    try {
      const ctx = loadCtx(req);
      const env = buildEnv(ctx);
      const def = pub(env.activeDef);
      const rc = FD.runChecks(def, Object.assign({}, env, { activeDef: null }));
      noStore(res);
      res.json({ def, view: FD.viewOf(def), checks: rc.checks, defHash: rc.defHash, baseVer: 'legacy-v0@' + (rc.defHash || '').slice(0, 16), skipUserCount: QF.legacyInputs(env.flowCfg).skipUsers.length });
    } catch (e) {
      console.error('[flow-designer default-flow]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '讀取預設流程失敗');
    }
  });

  app.post('/api/admin/quote-flow/designer/validate', gate, ...qAdmin, (req, res) => {
    try {
      const ctx = loadCtx(req);
      const def = takeDef(res, (req.body || {}).def);
      if (!def) return;
      const env = buildEnv(ctx);
      const rc = FD.runChecks(def, env);
      noStore(res);
      res.json(Object.assign({ view: rc.structural ? FD.viewOf(def) : null }, rc));
    } catch (e) {
      console.error('[flow-designer validate]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '驗證失敗');
    }
  });

  /** 模擬／掃描前的共同前置：淨化 def、要求結構有效、組環境 */
  function prepare(req, res) {
    const ctx = loadCtx(req);
    const def = takeDef(res, (req.body || {}).def);
    if (!def) return null;
    const v = QF.validateFlowDef(def);
    if (!v.ok) { fail(res, 400, 'FLOW_INVALID_DEF', '流程定義無效，無法模擬：' + v.errors.slice(0, 3).map((e) => e.message).join('；'), { errors: v.errors.slice(0, 10) }); return null; }
    return { ctx, def, env: buildEnv(ctx) };
  }

  app.post('/api/admin/quote-flow/designer/simulate', gate, ...qAdmin, (req, res) => {
    try {
      const p = prepare(req, res);
      if (!p) return;
      const b = req.body || {};
      const parsed = FD.parseSimInput(b.input, p.env.flowCfg);
      if (!parsed.ok) return fail(res, 400, parsed.code, parsed.message);
      const out = FD.simulate(p.def, parsed, p.env, b.compare === true ? p.env.activeDef : null);
      noStore(res);
      res.json(out);
    } catch (e) {
      console.error('[flow-designer simulate]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '模擬失敗');
    }
  });

  app.post('/api/admin/quote-flow/designer/sweep', gate, ...qAdmin, (req, res) => {
    try {
      const p = prepare(req, res);
      if (!p) return;
      const b = req.body || {};
      const base = (b.input && typeof b.input === 'object' && !Array.isArray(b.input)) ? Object.assign({}, b.input) : {};
      if (base.amount === undefined) base.amount = '1';
      const out = FD.sweep(p.def, base, p.env, p.env.activeDef, p.env.flowCfg);
      if (!out.ok) return fail(res, 400, out.code, out.message);
      noStore(res);
      res.json(out);
    } catch (e) {
      console.error('[flow-designer sweep]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '邊界掃描失敗');
    }
  });

  app.post('/api/admin/quote-flow/designer/overview', gate, ...qAdmin, (req, res) => {
    try {
      const p = prepare(req, res);
      if (!p) return;
      const b = req.body || {};
      let sim = null;
      if (b.input !== undefined && b.input !== null) {
        const parsed = FD.parseSimInput(b.input, p.env.flowCfg);
        if (!parsed.ok) return fail(res, 400, parsed.code, parsed.message);
        sim = FD.simulate(p.def, parsed, p.env, null);
      }
      noStore(res);
      res.json({ graph: FD.buildOverview(p.def, sim), sim: sim ? { route: sim.route, tiers: sim.tiers, steps: sim.steps, applied: sim.applied, blockers: sim.blockers } : null });
    } catch (e) {
      console.error('[flow-designer overview]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '產生總覽失敗');
    }
  });

  app.get('/api/admin/quote-flow/designer/overview', gate, ...qAdmin, (req, res) => {
    try {
      const ctx = loadCtx(req);
      const q = req.query.draftId;
      let def, source, name;
      if (q === undefined || q === '') { def = pub(buildEnv(ctx).activeDef); source = 'default'; name = def.name; }
      else {
        if (typeof q !== 'string' || !FD.ID_RE.test(q)) return fail(res, 400, 'BAD_ID', '草稿代號格式無效');
        const rec = (ctx.data.quoteFlowDrafts && Array.isArray(ctx.data.quoteFlowDrafts.items) ? ctx.data.quoteFlowDrafts.items : []).find((x) => x.id === q);
        if (!rec) return fail(res, 404, 'NOT_FOUND', '找不到這份草稿');
        def = pub(rec.def); source = 'draft'; name = rec.name;
      }
      const v = QF.validateFlowDef(def);
      if (!v.ok) return fail(res, 400, 'FLOW_INVALID_DEF', '流程定義無效，無法產生總覽');
      noStore(res);
      res.json({ source, name, graph: FD.buildOverview(def, null) });
    } catch (e) {
      console.error('[flow-designer overview GET]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '產生總覽失敗');
    }
  });

  // ── 草稿 CRUD（data.quoteFlowDrafts = { v:1, items:[…] }）──
  const store = (data) => {
    if (!data.quoteFlowDrafts || typeof data.quoteFlowDrafts !== 'object' || !Array.isArray(data.quoteFlowDrafts.items)) return { v: 1, items: [] };
    return data.quoteFlowDrafts;
  };
  const safeId = (s) => typeof s === 'string' && FD.ID_RE.test(s);

  app.get('/api/admin/quote-flow/drafts', gate, ...qAdmin, (req, res) => {
    const ctx = loadCtx(req);
    noStore(res);
    res.json({ items: store(ctx.data).items.map(FD.summaryOf).sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt))), limit: FD.LIMITS.drafts });
  });

  app.get('/api/admin/quote-flow/drafts/:id', gate, ...qAdmin, (req, res) => {
    const ctx = loadCtx(req);
    if (!safeId(req.params.id)) return fail(res, 404, 'NOT_FOUND', '找不到這份草稿');
    const rec = store(ctx.data).items.find((x) => x.id === req.params.id);
    if (!rec) return fail(res, 404, 'NOT_FOUND', '找不到這份草稿');
    noStore(res);
    res.json(Object.assign({}, FD.summaryOf(rec), { def: pub(rec.def), view: FD.viewOf(rec.def) }));
  });

  app.put('/api/admin/quote-flow/drafts', gate, ...qAdmin, (req, res) => {
    try {
      const ctx = loadCtx(req);
      const b = req.body || {};
      const allowed = ['id', 'name', 'note', 'def', 'baseVer', 'ifUpdatedAt'];
      for (const k of Object.keys(b)) if (!allowed.includes(k)) return fail(res, 400, 'BAD_FIELD', '不認得的欄位：' + FD.clipText(k, 30));
      const name = FD.cleanName(b.name, FD.LIMITS.name);
      if (!name) return fail(res, 400, 'BAD_NAME', `草稿名稱必填，至多 ${FD.LIMITS.name} 字，不可含控制字元`);
      let note = '';
      if (b.note !== undefined && b.note !== null && b.note !== '') {
        note = FD.cleanName(b.note, FD.LIMITS.note);
        if (!note) return fail(res, 400, 'BAD_NOTE', `備註至多 ${FD.LIMITS.note} 字，不可含控制字元`);
      }
      let baseVer = '';
      if (b.baseVer !== undefined && b.baseVer !== null && b.baseVer !== '') {
        if (typeof b.baseVer !== 'string' || !/^[A-Za-z0-9._@:-]{1,60}$/.test(b.baseVer)) return fail(res, 400, 'BAD_BASEVER', 'baseVer 格式無效');
        baseVer = b.baseVer;
      }
      if (b.id !== undefined && b.id !== null && !safeId(b.id)) return fail(res, 400, 'BAD_ID', '草稿代號格式無效');
      const def = takeDef(res, b.def);
      if (!def) return;
      const env = buildEnv(ctx);
      // 身分由伺服器決定：id＝草稿代號、no＝0（版本號是 M3 提議時才配）、origin＝designer、名稱＝草稿名稱
      const st = store(ctx.data);
      const existing = b.id ? st.items.find((x) => x.id === b.id) : null;
      if (b.id && !existing) return fail(res, 404, 'NOT_FOUND', '找不到要更新的草稿（可能已被刪除）');
      if (!existing && st.items.length >= FD.LIMITS.drafts) return fail(res, 409, 'DRAFT_LIMIT', `草稿數量已達上限（${FD.LIMITS.drafts} 份），請先刪除不用的草稿`);
      const id = existing ? existing.id : FD.newDraftId();
      def.id = id; def.no = 0; def.origin = 'designer'; def.name = name;
      const rc = checksOr400(res, def, env, true);
      if (!rc) return;
      if (existing && b.ifUpdatedAt !== undefined && b.ifUpdatedAt !== existing.updatedAt) return fail(res, 409, 'DRAFT_STALE', '這份草稿已被其他人（或其他視窗）更新過，請重新載入後再儲存', { updatedAt: existing.updatedAt, updatedBy: existing.updatedBy });
      const now = nowIso();
      const rec = existing || { id, createdAt: now, createdBy: ctx.me };
      rec.name = name; rec.note = note; rec.def = def; rec.baseVer = baseVer || rec.baseVer || ''; rec.updatedAt = now; rec.updatedBy = ctx.me;
      if (!existing) st.items.push(rec);
      ctx.data.quoteFlowDrafts = st;
      db.save(ctx.data);
      writeLog('SAVE_FLOW_DRAFT', ctx.me, FD.clipText(name, 60), `${existing ? '更新' : '新增'}流程草稿「${FD.clipText(name, 40)}」（${id}）｜路線 ${def.routes.length} 條、調整規則 ${def.modifiers.length} 個｜內容雜湊 ${(rc.defHash || '').slice(0, 12)}｜草稿不影響任何簽核（M3 才開放生效）`, req);
      noStore(res);
      res.json(Object.assign({ success: true, created: !existing }, FD.summaryOf(rec), { def: pub(rec.def), view: FD.viewOf(rec.def), checks: rc.checks }));
    } catch (e) {
      console.error('[flow-designer save]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '儲存草稿失敗');
    }
  });

  app.delete('/api/admin/quote-flow/drafts/:id', gate, ...qAdmin, (req, res) => {
    try {
      const ctx = loadCtx(req);
      if (!safeId(req.params.id)) return fail(res, 404, 'NOT_FOUND', '找不到這份草稿');
      const st = store(ctx.data);
      const i = st.items.findIndex((x) => x.id === req.params.id);
      if (i < 0) return fail(res, 404, 'NOT_FOUND', '找不到這份草稿');
      const rec = st.items[i];
      st.items.splice(i, 1);
      ctx.data.quoteFlowDrafts = st;
      db.save(ctx.data);
      writeLog('DELETE_FLOW_DRAFT', ctx.me, FD.clipText(rec.name, 60), `刪除流程草稿「${FD.clipText(rec.name, 40)}」（${rec.id}）｜原更新者 ${FD.clipText(rec.updatedBy, 40)}`, req);
      noStore(res);
      res.json({ success: true });
    } catch (e) {
      console.error('[flow-designer delete]', e && e.message);
      fail(res, 500, 'DESIGNER_FAILED', '刪除草稿失敗');
    }
  });
};
