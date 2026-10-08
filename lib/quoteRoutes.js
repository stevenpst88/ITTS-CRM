'use strict';
/**
 * 報價單路由：CRUD、成本填寫、簽核（送簽／核准／駁回／撤回／改派）、簽核設定、報價專用章、匯出。
 *
 * 規則（核決表、金額門檻、毛利率精算、內容雜湊）在 lib/quoteApproval.js；
 * 這裡負責「誰能做什麼」（權限）、狀態機、通知與稽核。規格摘要見 memory：crm_quote_approval_spec。
 *
 * 幾條不可退讓的原則（金錢控管：漏放比多擋嚴重，不確定就擋）：
 *  1. 簽核人的資格一律由伺服器判斷，前端的按鈕只是顯示；前端看到的 perm 也是伺服器算的。
 *  2. 簽核中整張單鎖定；核准後改到「雜湊涵蓋欄位」會作廢核准（需 confirmVoid）。
 *  3. 核准／駁回必須帶「當時看到的內容雜湊」，不符就 409，避免簽到一份被改過的單。
 *  4. 核准人不可是擁有者或送簽人；同一人不可簽同一張單的兩個關卡；管理員不可代簽（只能改派）。
 *  5. 先 db.save 再 pushNotification（後者內部自己 load/save，順序反了會互相覆寫）。
 */
const fs = require('fs');
const QA = require('./quoteApproval');
const quoteExcel = require('./quoteExcel');
const REM = require('./quoteRemarks');
const QI = require('./quoteItems');
const CL = require('./quoteCostLines');   // 成本明細 costLines（新式成本）：q.costLines 欄位存在＝新式；不存在＝舊式（items[].cost，行為不變）
const { sheetToHtml } = require('./xlsxSheetHtml');   // 毛利分析預覽：把實際要下載的 xlsx 轉成 HTML
const productCatalogLib = require('./productCatalog');

const MAX_ITEMS = 50;
const MAX_PRODUCTS = 30;
/** 品項的「毛利分類」（毛利分析 PNL 表把收入/成本分到這四區；空＝自動，匯出時依商品類別推測）。與 lib/quotePnlExcel.js 的 CATS 一致 */
const ITEM_CATS = ['consult', 'software', 'hardware', 'other'];
/** 風險預留（Contingency，%）：填成本的顧問主管依專案風險預估，毛利分析（內部 PNL 表）把「顧問服務成本 × 此值」另計為成本。與 lib/quotePnlExcel.js 一致 */
const CONTINGENCY_PCTS = [0, 5, 10, 15, 20];
const MAX_HISTORY = 2000;      // 簽核紀錄是唯一的內建軌跡，不能被「送簽→撤回」迴圈擠掉；上限只當記憶體保險
const SEAL_MAX_BYTES = 300 * 1024;
// 有「看成本」資格的角色（且報價單在其可視範圍內）：主管、管理類角色與秘書。業務／行銷不在內。
// 秘書的「可視範圍」依部門(BU)決定，見 inViewScope()。
const COST_VIEW_ROLES = ['manager1', 'manager2', 'executive', 'accounting_manager', 'finance_manager', 'secretary'];
// 不建立報價單的角色（規格：業務發起；業務主管、總經理、董事長、秘書不建單）。admin 保留（維運／示範用）。
const NO_CREATE_ROLES = ['executive', 'manager1', 'secretary', 'marketing', 'tecopm', 'groupsales', 'pool', 'accounting_manager', 'finance_manager', 'consult_manager_south', 'consult_manager_north'];
const ALL_BUS = ['ERP', 'ITS', 'MDM', 'CRM'];

module.exports = function registerQuoteRoutes(app, deps) {
  const {
    db, loadAuth, saveAuth, requireAuth, requireAdmin, writeLog, pushNotification, getViewableOwners,
    sanitizeStr, genQuoteNo, taipeiToday, resolveIssuer, buildQuoteWorkbook, buildQuotePnlExcel,
    QUOTE_TEMPLATE, uuidv4, normalizeBu, getUserFeatures,
  } = deps;

  // ── 小工具 ─────────────────────────────────────────────
  const nowIso = () => new Date().toISOString();
  const isActive = (u) => !!u && u.active !== false && u.role !== 'pool';
  // 能「執行簽核／代核／填成本」的人：在職、非管理員（管理員只能改派，不能代簽）、非唯讀帳號
  // （唯讀帳號的寫入會被全域 middleware 擋下；放進名冊只會卡單）。
  const canAct = (u) => isActive(u) && u.role !== 'admin' && u.accessMode !== 'view';
  const isRealDate = (s) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
    if (!m) return false;
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
  };
  const fail = (res, status, code, error, extra) => res.status(status).json(Object.assign({ error, code }, extra || {}));
  // GET 匯出的稽核紀錄：全域 flush middleware 只涵蓋 POST/PUT/DELETE/PATCH，Vercel 回應送出後實例可能立刻被凍結，
  // 所以在 res.send 前自己等寫入完成。盡力而為——失敗也不能讓使用者下載不到檔案。
  const flushAuditWrites = async () => { try { if (db && typeof db.flush === 'function') await db.flush(); } catch (_) { /* 稽核寫入失敗不影響下載 */ } };
  const num0 = (v) => { const n = parseFloat(v); return Number.isFinite(n) && n >= 0 && n <= 1e12 ? n : 0; };
  const userMap = (auth) => { const m = {}; (auth.users || []).forEach(u => { m[u.username] = u; }); return m; };
  // 暱稱（後台統一設定，用來區分同名業務）> 顯示名稱 > 帳號；與 server.js 的 userLabel 同一套規則。簽核通知、簽核面板、名冊都用它，避免同名業務被簽錯
  const dispName = (users, un) => (users[un] && (users[un].nickname || users[un].displayName || users[un].username)) || un || '';
  const strList = (x) => (Array.isArray(x) ? x.filter(s => typeof s === 'string' && s) : []);

  // ── 簽核設定（data.quoteApproval）──────────────────────────
  function getCfg(data) {
    const c = data.quoteApproval || {};
    const r = c.roster || {};
    return {
      roster: {
        gm: strList(r.gm), chairman: strList(r.chairman), boardProxy: strList(r.boardProxy),
        costProviders: strList(r.costProviders), sealManagers: strList(r.sealManagers),
      },
      productClasses: (c.productClasses && typeof c.productClasses === 'object') ? c.productClasses : {},
    };
  }
  /** 董事會代核人＝後台名冊 ∪ 所有在職 secretary 角色 */
  function boardProxySet(cfg, users) {
    const s = new Set(cfg.roster.boardProxy);
    Object.values(users).forEach(u => { if (u.role === 'secretary') s.add(u.username); });
    return new Set([...s].filter(un => canAct(users[un])));
  }
  /** 報價章管理人＝後台名冊 ∪ 所有 secretary ∪ admin */
  function sealManagerSet(cfg, users) {
    const s = new Set(cfg.roster.sealManagers);
    Object.values(users).forEach(u => { if (u.role === 'secretary' || u.role === 'admin') s.add(u.username); });
    return new Set([...s].filter(un => isActive(users[un])));
  }
  /** 業務的「一級主管」：沿 supervisor 鏈往上第一位在職 manager1。找不到就是 null（不退回 BU 推斷，避免簽給不相干的人） */
  function findManager1(users, ownerUsername) {
    const seen = new Set();
    let cur = users[ownerUsername];
    while (cur && cur.supervisor && !seen.has(cur.username)) {
      seen.add(cur.username);
      const up = users[cur.supervisor];
      if (!up) break;
      if (up.role === 'manager1' && canAct(up)) return up.username;
      cur = up;
    }
    return null;
  }

  // ── 載入脈絡（每個請求一次；由 qAuth 的 ctx middleware 建好掛在 req._qctx）──────────
  // 角色與在職狀態一律以「帳號資料現況」為準，不信任 JWT 內 8 小時前的舊值：
  // 被降級/停用的人立刻失去原有的報價單權限。
  function buildCtx(req) {
    const data = db.load();
    if (!data.quotations) data.quotations = [];
    const auth = loadAuth();
    const users = userMap(auth);
    const me = req.session.user.username;
    const live = users[me];
    if (!live || live.active === false) return { error: { status: 401, code: 'ACCOUNT_DISABLED', message: '帳號已停用或不存在，請重新登入' } };
    const role = live.role;
    const cfg = getCfg(data);
    let changed = false;
    // 舊資料沒有 lid：第一次讀到時補上並存檔（之後前端以 lid 對應成本）
    data.quotations.forEach(q => {
      (q.items || []).forEach(it => { if (!it.lid) { it.lid = uuidv4(); changed = true; } });
    });
    if (changed) db.save(data);
    // getViewableOwners 讀的是 req.session.user.role 等欄位 → 用帳號現況組一個替身 req
    const shimReq = { session: { user: Object.assign({}, req.session.user, { role, viewOwnerScope: live.viewOwnerScope, viewGroupId: live.viewGroupId }) } };
    const scope = new Set(getViewableOwners(shimReq, 'quotations'));
    // 功能權限矩陣（「報價單」功能）：沒有此功能的角色不能憑「可視範圍」瀏覽/建單；
    // 但「被指派」的人（簽核人、成本填寫人）不受影響，仍可處理自己手上的單。
    const featureOk = role === 'admin' || typeof getUserFeatures !== 'function' || getUserFeatures(role).includes('quotations');
    return { data, auth, users, cfg, scope, me, role, featureOk };
  }
  function ctxMiddleware(req, res, next) {
    let ctx;
    try { ctx = buildCtx(req); } catch (e) { console.error('[quote ctx]', e); return fail(res, 500, 'CTX_FAILED', '讀取資料失敗'); }
    if (ctx.error) return fail(res, ctx.error.status, ctx.error.code, ctx.error.message);
    req._qctx = ctx;
    next();
  }
  const loadCtx = (req) => req._qctx;
  const qAuth = [requireAuth, ctxMiddleware];
  const qAdmin = [requireAdmin, ctxMiddleware];

  const needsConsultant = (q, cfg) => QA.classifyProducts(q.products || [], cfg.productClasses).needsConsultantCost;
  /** 顧問是否「動過」這張單的成本（儲存過或完成過）。動過之後，業務不可讀取/覆寫逐列成本 */
  const consultantTouched = (q) => !!(q.costFlow && (q.costFlow.consultantWrote || q.costFlow.filledAt));

  /** 該使用者是否為這個關卡的合格簽核人（不含「不可簽自己的單」等通用規則） */
  function stepActorOk(step, me, cfg, users) {
    if (!step || !canAct(users[me])) return false;
    switch (step.tier) {
      case 'mgr1': return step.assignee === me && users[me].role === 'manager1';
      case 'gm': return cfg.roster.gm.includes(me);
      case 'chairman': return cfg.roster.chairman.includes(me);
      case 'board': return boardProxySet(cfg, users).has(me);
      default: return false;
    }
  }
  const alreadySigned = (ap, me) => (ap.steps || []).some(s => s.status === 'approved' && s.by === me);
  /** 職責分離：擁有者、送簽人、成本填寫人、已簽過其他關卡的人，都不能簽這張單 */
  const conflicted = (q, ap, me) => me === q.owner || me === ap.submittedBy
    || !!(q.costFlow && q.costFlow.by === me) || !!(q.costBy && q.costBy === me) || alreadySigned(ap, me);

  // ── 檢視者與報價單的關係／權限 ────────────────────────────
  /**
   * 這張單是否在檢視者的「可視範圍」內。
   * 一般角色沿用 getViewableOwners('quotations')；秘書改依部門(BU)：只看自己所屬 BU 的業務發起的單，
   * 掛滿四個 BU（ERP/ITS/MDM/CRM）的秘書看全部。
   */
  function inViewScope(q, ctx) {
    if (ctx.role === 'admin') return true;     // 管理員看得到全部（含帳號已被刪除的業務留下的單）
    if (!ctx.featureOk) return false;          // 角色沒有「報價單」功能 → 不能憑可視範圍瀏覽
    if (ctx.role === 'secretary') {
      const mine = normalizeBu((ctx.users[ctx.me] || {}).bu);
      if (ALL_BUS.every(b => mine.includes(b))) return true;
      const theirs = normalizeBu((ctx.users[q.owner] || {}).bu);
      return theirs.some(b => mine.includes(b));
    }
    return ctx.scope.has(q.owner);
  }

  function relationOf(q, ctx) {
    const { me, role, cfg, users } = ctx;
    const ap = q.approval;
    const rel = {
      me, role, isAdmin: role === 'admin', isOwner: q.owner === me, inScope: inViewScope(q, ctx),
      isCostBy: !!q.costBy && q.costBy === me && q.costBy !== q.owner, actTiers: [],
    };
    // 「因簽核而可檢視」只算：(a) 輪到我簽的那一關（目前待簽），(b) 我實際簽過/駁回過的關卡。
    // 不可把還在 waiting 的後段關卡也算進來：董事會關的合格人是所有秘書，會讓任何 BU 的秘書從送簽那刻就看到單。
    if (ap && Array.isArray(ap.steps) && ap.state !== 'none') {
      ap.steps.forEach((s, i) => {
        const signedByMe = !!s.by && s.by === me;
        const currentForMe = ap.state === 'pending' && i === ap.cur && s.status === 'pending' && stepActorOk(s, me, cfg, users);
        if (signedByMe || currentForMe) rel.actTiers.push(s.tier);
      });
    }
    // 董事會代核人（任何 BU 的秘書）只為了登錄決議而看內容，不因此取得逐列成本/毛利分析
    rel.costTiers = rel.actTiers.filter(t => t !== 'board');
    rel.canView = rel.inScope || rel.isCostBy || rel.actTiers.length > 0;
    // 「純顧問」視角：只因為被指派填成本才看得到這張單 → 不給單價、金額、毛利
    rel.consultantOnly = rel.isCostBy && !rel.inScope && !rel.isAdmin && rel.actTiers.length === 0;
    return rel;
  }

  function permOf(rel, q, ctx) {
    const { cfg } = ctx;
    const ap = q.approval || null;
    const st = ap ? ap.state : 'none';
    const needC = needsConsultant(q, cfg);
    const cfState = q.costFlow ? q.costFlow.state : 'na';
    const touched = consultantTouched(q);
    let canSeeCost = false;
    if (rel.isAdmin || rel.isCostBy || rel.costTiers.length) canSeeCost = true;
    // 業務（擁有者）看自己單子的成本：
    //  (a) 不需顧問且顧問從未動過 → 自己填的成本；
    //  (b) 需顧問的單 → 顧問按「完成」之後才能（唯讀）看到逐列成本、風險預留與毛利分析（Steven 2026-10-07 指示：業務也要能預覽成本）。
    //      顧問填寫中、或完成後因品項結構改動被退回 requested 時仍不顯示（半成品數字不給看）。
    // (a) 不能只看 needC：needC 由業務可改的 products 決定，改成 [] 就能讀到顧問填的成本；所以要同時確認 !touched。
    else if (rel.isOwner) canSeeCost = (!needC && !touched) || (needC && cfState === 'filled');
    else if (rel.inScope) canSeeCost = COST_VIEW_ROLES.includes(rel.role);
    let canEditCost = false;
    if (st !== 'pending' && st !== 'approved') {
      if (needC) canEditCost = rel.isCostBy && (cfState === 'requested' || cfState === 'filled');
      else canEditCost = rel.isOwner && !touched;
    }
    const perm = {
      isOwner: rel.isOwner,
      canEdit: (rel.isOwner || rel.isAdmin) && st !== 'pending',
      canEditCost, canSeeCost,
      canSeePrice: !rel.consultantOnly,
      canSubmit: rel.isOwner && (st === 'none' || st === 'returned'),
      canWithdraw: rel.isOwner && st === 'pending',
      canApprove: false, canReturn: false, canReassign: false,
      isCostProvider: rel.isCostBy,
    };
    if (st === 'pending' && ap) {
      const step = (ap.steps || [])[ap.cur];
      if (step && step.status === 'pending' && stepActorOk(step, rel.me, cfg, ctx.users) && !conflicted(q, ap, rel.me)) {
        perm.canApprove = true; perm.canReturn = true;
      }
      if (rel.isAdmin && ap.cur === 0 && step && step.tier === 'mgr1') perm.canReassign = true;
    }
    perm.canReassign = !!perm.canReassign;
    return perm;
  }

  // ── 試算與送簽前檢查 ──────────────────────────────────────
  /** 成本是否已齊：每個有價列都有成本，且（若需顧問）顧問已按「完成」 */
  function costComplete(q, cfg) {
    const items = q.items || [];
    if (!items.length) return false;
    if (CL.hasCostLines(q)) {
      // 新式成本：非印花稅的成本明細至少 1 列且總成本 > 0（與 QA.computeFinancials 的 costComplete 共用同一個定義）
      if (!CL.costLinesComplete(q)) return false;
    } else {
      const priced = items.filter(it => (parseFloat(it.unitPrice) || 0) > 0);
      if (!priced.every(it => (parseFloat(it.cost) || 0) > 0)) return false;
    }
    if (needsConsultant(q, cfg)) return !!(q.costFlow && q.costFlow.state === 'filled');
    return true;
  }

  function chainBlockers(tiers, q, ctx) {
    const { cfg, users } = ctx;
    const out = [];
    if (tiers.includes('mgr1') && !findManager1(users, q.owner)) {
      out.push({ code: 'NO_MANAGER', message: '找不到你的一級主管（直屬主管鏈上沒有在職的一級主管），請管理員先設定直屬主管' });
    }
    const activeOf = (list) => list.filter(un => canAct(users[un]));
    if (tiers.includes('gm') && !activeOf(cfg.roster.gm).length) out.push({ code: 'NO_GM', message: '尚未設定總經理名單，請管理員至「簽核設定」維護' });
    if (tiers.includes('chairman') && !activeOf(cfg.roster.chairman).length) out.push({ code: 'NO_CHAIRMAN', message: '尚未設定董事長名單，請管理員至「簽核設定」維護' });
    if (tiers.includes('board') && !boardProxySet(cfg, users).size) out.push({ code: 'NO_BOARD_PROXY', message: '找不到董事會代核人（管理部秘書），請管理員確認' });
    if (out.length) return out;
    // 每一關都必須有「不同的人」可簽：擁有者、成本填寫人不能簽自己的單；同一個人也不能簽兩關。
    // 名單重疊（例如同一人同時在總經理與董事長名單）會讓單子簽到一半無人可簽，要在送簽前就擋下。
    const excluded = new Set([q.owner]);
    if (q.costBy) excluded.add(q.costBy);
    const cands = tiers.map(t => {
      let list = [];
      if (t === 'mgr1') { const m = findManager1(users, q.owner); list = m ? [m] : []; }
      else if (t === 'gm') list = activeOf(cfg.roster.gm);
      else if (t === 'chairman') list = activeOf(cfg.roster.chairman);
      else if (t === 'board') list = [...boardProxySet(cfg, users)];
      return list.filter(un => !excluded.has(un));
    });
    const emptyIdx = cands.findIndex(c => !c.length);
    if (emptyIdx >= 0) {
      out.push({ code: 'NO_ELIGIBLE_SIGNER', message: `「${QA.TIERS[tiers[emptyIdx]]}」這一關沒有可以簽核的人（擁有者與成本填寫人本人不能簽自己的單），請管理員調整簽核名單` });
    } else if (!distinctAssignment(cands)) {
      out.push({ code: 'SIGNER_OVERLAP', message: '簽核名單有重疊：同一個人不能簽同一張單的兩個關卡，目前無法讓每一關都有不同的人簽核，請管理員調整簽核名單' });
    }
    return out;
  }
  /** 每一關各挑一位、且不重複的人是否存在（關卡最多 4 個，回溯即可） */
  function distinctAssignment(cands, i = 0, used = new Set()) {
    if (i === cands.length) return true;
    for (const p of cands[i]) {
      if (used.has(p)) continue;
      used.add(p);
      if (distinctAssignment(cands, i + 1, used)) return true;
      used.delete(p);
    }
    return false;
  }

  function buildPreview(q, ctx) {
    const { cfg } = ctx;
    const cls = QA.classifyProducts(q.products || [], cfg.productClasses);
    const blockers = [];
    const v = QA.validateForSubmit(q, cfg.productClasses);
    if (!v.ok) v.errors.forEach(e => blockers.push({ code: e.code, message: e.message }));
    const complete = costComplete(q, cfg);
    const derived = complete ? QA.buildDerived(q, cfg.productClasses) : null;
    if (derived) chainBlockers(derived.tiers, q, ctx).forEach(b => blockers.push(b));
    const row = QA.resolveRow(cls.classKeys);
    // 建議性警告（只提醒、不擋）：新式成本（成本明細）有「有價品項沒有成本列涵蓋」時加一則；文字也進 warnings，既有畫面（試算區、送簽確認、簽核面板）自動顯示。
    // 舊式單、沒有未涵蓋品項的新式單：warnings 內容與 preview 的鍵都和以前完全相同
    const costWarns = CL.hasCostLines(q) ? CL.costWarnings(q) : [];
    return {
      rowKey: derived ? derived.rowKey : row.key,
      rowLabel: derived ? derived.rowLabel : row.label,
      level: derived ? derived.level : null,
      board: derived ? derived.board : false,
      marginText: derived ? derived.marginText : null,
      tiers: derived ? derived.tiers : null,
      reasons: derived ? derived.reasons : ['成本尚未填妥，暫時無法試算毛利率與需要簽核的關卡'],
      warnings: [...new Set([].concat(cls.warnings || [], (derived && derived.warnings) || [], costWarns.map(w => w.message)))],   // 兩處都會產生同一則警告，去重
      needsConsultantCost: cls.needsConsultantCost,
      unclassified: cls.unclassified || [],
      blockers,
      ...(costWarns.length ? { costWarnings: costWarns } : {}),
    };
  }

  // ── 序列化（依檢視者過濾欄位）──────────────────────────────
  function tierPeople(tier, ctx) {
    const { cfg, users } = ctx;
    const names = (list) => list.filter(un => isActive(users[un])).map(un => dispName(users, un)).join('、');
    if (tier === 'gm') return names(cfg.roster.gm);
    if (tier === 'chairman') return names(cfg.roster.chairman);
    if (tier === 'board') return '管理部秘書';
    return '';
  }

  function serializeApproval(q, rel, perm, ctx) {
    const ap = q.approval;
    if (!ap || rel.consultantOnly) return null;
    const { users } = ctx;
    const valid = ap.state === 'approved' && QA.contentHash(q) === ap.hash;
    const v = {
      state: ap.state, valid,
      submittedAt: ap.submittedAt || null, submittedByName: dispName(users, ap.submittedBy),
      steps: (ap.steps || []).map(s => ({
        tier: s.tier, label: s.label, status: s.status,
        assigneeName: s.assignee ? dispName(users, s.assignee) : tierPeople(s.tier, ctx),
        byName: s.by ? dispName(users, s.by) : '', at: s.at || null, comment: s.comment || '',
      })),
      cur: ap.cur || 0,
      board: ap.board ? { resolutionDate: ap.board.resolutionDate, resolutionNo: ap.board.resolutionNo, byName: dispName(users, ap.board.by), at: ap.board.at } : null,
      history: (ap.history || []).map(h => ({ at: h.at, byName: dispName(users, h.by), action: h.action, comment: h.comment || '', tier: h.tier || '' })),
    };
    if (ap.derived) {
      const d = ap.derived;
      v.derived = { rowKey: d.rowKey, rowLabel: d.rowLabel, level: d.level, board: d.board, marginText: d.marginText, tiers: d.tiers, reasons: d.reasons };
      // 送簽當時凍結的建議性警告（有價品項沒有成本列；只含品名與數量，不含金額）。沒有就不多這個鍵
      if (Array.isArray(d.costWarnings) && d.costWarnings.length) v.derived.costWarnings = d.costWarnings;
      // 彙總金額（折扣後未稅／總成本／毛利）：看得到成本的人，或董事會代核人（要列印董事會簽呈）。
      // 董事會代核人只拿到彙總數字；逐列成本與毛利分析下載仍不開放。
      if (perm.canSeeCost || rel.actTiers.includes('board')) { v.derived.revenueCents = d.revenueCents; v.derived.costCents = d.costCents; v.derived.gpCents = d.gpCents; }
    }
    return v;
  }

  /** 系統是否已上傳報價專用章（PDF 下載稽核紀錄回報蓋章狀態用） */
  const sealUploaded = (ctx) => !!(ctx.data && ctx.data.quoteSeal && ctx.data.quoteSeal.base64);

  function serialize(q, ctx) {
    const { users, cfg } = ctx;
    const rel = relationOf(q, ctx);
    const perm = permOf(rel, q, ctx);
    const cf = q.costFlow || { state: needsConsultant(q, cfg) ? 'needed' : 'na' };
    const out = {
      id: q.id, quoteNo: q.quoteNo, owner: q.owner, ownerName: dispName(users, q.owner),
      company: q.company, projectName: q.projectName, projectNo: q.projectNo, quoteDate: q.quoteDate,
      status: q.status, createdAt: q.createdAt, updatedAt: q.updatedAt,
      products: q.products || [],
      costBy: q.costBy || null, costByName: q.costBy ? dispName(users, q.costBy) : '',
      costFlow: {
        state: cf.state, by: cf.by || null, byName: cf.by ? dispName(users, cf.by) : '',
        requestedAt: cf.requestedAt || null, filledAt: cf.filledAt || null, note: cf.note || '',
      },
      items: (q.items || []).map((it, i) => {
        const o = { lid: it.lid || ('legacy-' + i), desc: it.desc, unit: it.unit, qty: it.qty };
        if (it.cat) o.cat = it.cat;
        if (QI.isNonItemRow(it)) o.kind = it.kind;   // 標題／小計列（金額由前端用單價算：沒有價格權限的人看不到小計金額）
        if (perm.canSeePrice) o.unitPrice = it.unitPrice;
        if (perm.canSeeCost) o.cost = it.cost;
        return o;
      }),
      perm,
    };
    // 風險預留屬於成本資料：看得到成本的人才看得到（負責填成本的顧問、擁有者在成本已可見時、主管…）
    if (perm.canSeeCost) out.contingencyPct = typeof q.contingencyPct === 'number' ? q.contingencyPct : null;
    // 成本明細（新式成本）：只有看得到成本的人才輸出（沒有 costLines 欄位的舊單完全不多任何欄位）。
    // 印花稅金額＝營收×0.001。v1.1（業主 2026-10-07 確認：在 ITTS，顧問與業務本來就可以互相知道成本與最終售價）：
    // 所有看得到成本的人（含只負責填成本的顧問）都拿得到印花稅金額與 costBreakdown；營收可由印花稅金額約略推算，這是已接受的。
    // 範圍僅限成本明細——consultantOnly 對 unitPrice／折扣／approval／preview 的隱藏維持不變。
    if (perm.canSeeCost && CL.hasCostLines(q)) {
      const fin = QA.computeFinancials(q);
      const revenueCents = fin && fin.ok ? fin.revenueCents : 0;
      out.costLines = CL.publicLines(q, { canSeeCost: true, canSeePrice: perm.canSeePrice, revenueCents });
      const t = CL.totalsByCat(q, revenueCents);   // 單位：分；other 含印花稅
      out.costBreakdown = { consult: t.consult, software: t.software, hw: t.hw, travel: t.travel, other: t.other };
    }
    // 過期分頁保護簽章（顧問對話框／業務毛利頁籤載入時記下，存檔時帶回；伺服器比對目前的值，不同就 409 STALE_*）。
    // 只給「看得到成本」或「可填成本」的人；都是 sha256 摘要（前 16 碼），不含明文金額，只用來比對內容有沒有變（細節見 quoteApproval.js 簽章函式註解）。
    // 沒有 costLines（舊式單）時 costLinesSig 是空字串。
    if (perm.canSeeCost || perm.canEditCost) {
      out.itemsSig = QA.itemsSig(q);
      out.costLinesSig = QA.costLinesSig(q);
    }
    if (!rel.consultantOnly) {
      Object.assign(out, {
        contactId: q.contactId || '', contactName: q.contactName, phone: q.phone, mobile: q.mobile, address: q.address,
        note: q.note, discountType: q.discountType, discountValue: q.discountValue,
        // Remarks 活的部分（有效值：舊單沒存就回預設句／單行備註轉成的第 7 條，表單與送回後的「無實質變更」判斷都靠這個）
        payment: q.payment || REM.DEFAULT_PAYMENT, clauses: REM.effectiveClauses(q),
        validUntil: quoteExcel.effectiveValidUntil(q),   // 報價期限（舊單沒存就回「建立當月最後一個工作天」＝舊規則，與匯出實際印的一致）
      });
      out.approval = serializeApproval(q, rel, perm, ctx);
      out.preview = buildPreview(q, ctx);
      // 涵蓋警告（品項沒有成本列涵蓋）來自成本明細——還看不到成本的人（例如顧問填寫中的業主）不可看到它，連文字也不行
      if (!perm.canSeeCost && out.preview.costWarnings) {
        const hide = new Set(out.preview.costWarnings.map((w) => w.message));
        out.preview = Object.assign({}, out.preview, { warnings: (out.preview.warnings || []).filter((w) => !hide.has(w)) });
        delete out.preview.costWarnings;
      }
      if (perm.canApprove || perm.canReturn) out.contentHash = QA.contentHash(q);
    } else {
      out.approval = null;
    }
    return out;
  }

  // ── 品項／商品正規化 ───────────────────────────────────────
  function normProducts(raw) {
    if (!Array.isArray(raw)) return null;
    const seen = new Set(); const out = [];
    for (const p of raw) {
      const s = sanitizeStr(p, 100);
      if (s && !seen.has(s)) { seen.add(s); out.push(s); }
      if (out.length >= MAX_PRODUCTS) break;
    }
    return out;
  }
  /**
   * 依 lid 對回既有列。
   * - cost：只有 acceptCost 且前端「有送」該欄時才採用；沒送就沿用既有值；resetCosts（需顧問↔不需顧問切換）則一律從 0 起算。
   * - lid 重複：後面重複的列視為新列（新 lid、成本 0），避免多列共用同一個舊成本、且顧問只能改到最後一列。
   */
  function normalizeItems(raw, existing, acceptCost, resetCosts) {
    const byLid = new Map();
    (existing || []).forEach((it, i) => byLid.set(it.lid || ('legacy-' + i), it));
    const used = new Set();
    return raw.slice(0, MAX_ITEMS).map(r => {
      let old = r && r.lid ? byLid.get(r.lid) : null;
      if (old && used.has(old.lid)) old = null;
      // 分組標題／小計列：沒有數量、單位、單價、成本、分類（從一般品項改成標題時，舊成本隨之丟掉）。一般品項不寫 kind，維持舊資料外形
      const kind = QI.normalizeRowKind(r && r.kind);
      if (kind !== 'item') {
        const klid = old && old.lid ? old.lid : uuidv4();
        used.add(klid);
        return { lid: klid, kind, desc: sanitizeStr(r && r.desc, 200) };
      }
      const qty = parseFloat(r && r.qty);
      const price = parseFloat(r && r.unitPrice);
      let cost = (old && !resetCosts) ? num0(old.cost) : 0;
      if (acceptCost && r && r.cost !== undefined && r.cost !== null && r.cost !== '') cost = num0(r.cost);
      const lid = old && old.lid ? old.lid : uuidv4();
      used.add(lid);
      // 毛利分析（內部 PNL 表）用的分類：沒送欄位（舊版前端）→ 沿用既有值；送空字串＝改回「自動」；不認得的值一律當空。
      // 不納入核准雜湊（contentHash 只取價格/數量/成本等欄位）：改分類不會讓已核准的單失效。
      let cat = old && ITEM_CATS.includes(old.cat) ? old.cat : '';
      if (r && r.cat !== undefined) cat = ITEM_CATS.includes(r.cat) ? r.cat : '';
      return {
        lid,
        desc: sanitizeStr(r && r.desc, 200),
        unit: sanitizeStr(r && r.unit, 20),
        qty: Number.isFinite(qty) && qty > 0 ? Math.max(0.001, qty) : 1,
        unitPrice: Number.isFinite(price) && price >= 0 ? price : 0,
        cost,
        ...(cat ? { cat } : {}),
      };
    });
  }

  /**
   * 畫面還沒存檔的新品項沒有 lid，成本明細編輯器用「暫時代號 nid」指向它（成本列 forLid／forLids 放 'nid-…'）。
   * 存檔時請求的 items 每列可帶 nid；normalizeItems 的輸出與輸入逐列對齊（同一個索引），所以 nid → 該列剛拿到的 lid。
   * 回傳 Map(nid → lid)，給 CL.resolveItemRefs 換掉成本列裡的暫時代號。只用於這一次請求，nid 不存檔。
   */
  function tmpRefMap(rawItems, outItems) {
    const m = new Map();
    (Array.isArray(rawItems) ? rawItems : []).forEach((r, i) => {
      const o = Array.isArray(outItems) ? outItems[i] : null;
      if (!r || typeof r !== 'object' || typeof r.nid !== 'string' || !o || !o.lid || QI.isNonItemRow(o)) return;
      const k = r.nid.trim().slice(0, 64);
      if (k && !m.has(k)) m.set(k, o.lid);
    });
    return m;
  }

  /** 依商品/顧問/品項結構重算成本流程；回傳是否要通知顧問 */
  function reconcileCostFlow(q, prevFlow, ctx, now, flip) {
    const { cfg } = ctx;
    const prev = prevFlow || {};
    const need = needsConsultant(q, cfg);
    const out = {
      state: 'na', by: null, requestedAt: prev.requestedAt || null, filledAt: prev.filledAt || null, note: prev.note || '',
      sig: prev.sig || null, consultantWrote: !!prev.consultantWrote,
    };
    if (q.costFlow && typeof q.costFlow.note === 'string') out.note = q.costFlow.note;
    // 「需顧問」↔「不需顧問」切換：成本整批作廢重來（呼叫端已把各列成本歸零），顧問寫過的紀錄一併重置
    if (flip) { out.filledAt = null; out.sig = null; out.consultantWrote = false; }
    let notify = false;
    if (!need) {
      q.costBy = null; out.state = 'na';
    } else if (!q.costBy) {
      out.state = 'needed';
    } else {
      out.by = q.costBy;
      if (prev.by !== q.costBy) {
        out.state = 'requested'; out.requestedAt = now; out.filledAt = null; notify = true;
      } else if (prev.state === 'filled') {
        if (QA.lineStructureSig(q.items || []) !== prev.sig) { out.state = 'requested'; out.requestedAt = now; out.filledAt = null; notify = true; }
        else out.state = 'filled';
      } else if (prev.state === 'requested') {
        out.state = 'requested';
      } else {
        out.state = 'requested'; out.requestedAt = now; notify = true;
      }
    }
    q.costFlow = out;
    return { notify };
  }

  // ── 通知 ──────────────────────────────────────────────────
  function notify(ctx, actor, usernames, type, title, body, refId) {
    const seen = new Set();
    (usernames || []).forEach(un => {
      if (!un || un === actor || seen.has(un) || !isActive(ctx.users[un])) return;
      seen.add(un);
      try { pushNotification(un, type, title, body, refId); } catch (e) { console.warn('[quote notify]', e.message); }
    });
  }
  const noteLine = (q, ctx) => `${q.quoteNo}｜${q.company || ''}｜業務 ${dispName(ctx.users, q.owner)}`;
  function stepRecipients(step, ctx) {
    if (!step) return [];
    if (step.tier === 'mgr1') return [step.assignee];
    if (step.tier === 'gm') return ctx.cfg.roster.gm;
    if (step.tier === 'chairman') return ctx.cfg.roster.chairman;
    if (step.tier === 'board') return [...boardProxySet(ctx.cfg, ctx.users)];
    return [];
  }

  function findQuote(ctx, id) { return (ctx.data.quotations || []).find(x => x.id === id) || null; }
  /** meta：稽核快照（雜湊、金額、毛利率、關卡）；不會序列化給前端，作廢/撤回清掉 derived 後仍可還原當初簽了什麼 */
  function pushHistory(ap, me, action, comment, tier, meta) {
    if (!Array.isArray(ap.history)) ap.history = [];
    const h = { at: nowIso(), by: me, action, comment: comment || '', tier: tier || '' };
    if (meta) h.meta = meta;
    ap.history.push(h);
    if (ap.history.length > MAX_HISTORY) ap.history = ap.history.slice(-MAX_HISTORY);
  }
  function auditMeta(q, ap) {
    const d = (ap && ap.derived) || {};
    return {
      hash: QA.contentHash(q), rules: ap && ap.rulesVersion, marginText: d.marginText || null,
      revenueCents: d.revenueCents === undefined ? null : d.revenueCents, costCents: d.costCents === undefined ? null : d.costCents,
      tiers: d.tiers || null,
    };
  }
  /** 修改前後的差異摘要（寫進稽核紀錄；價格/數量/成本變動帶數字） */
  function diffSummary(a, b) {
    const out = [];
    const isKindRowText = (it) => QI.isNonItemRow(it);
    const same = (x, y) => String(x === undefined || x === null ? '' : x) === String(y === undefined || y === null ? '' : y);
    const LABELS = { company: '公司', contactName: '聯絡人', phone: '電話', mobile: '手機', address: '地址', projectName: '專案名稱', projectNo: '專案號碼', note: '備註', quoteDate: '報價日期', validUntil: '報價期限', discountType: '折扣方式', discountValue: '折扣值', costBy: '支援顧問' };
    Object.keys(LABELS).forEach(k => {
      if (same(a[k], b[k])) return;
      out.push(['note', 'address', 'contactName', 'phone', 'mobile', 'projectName', 'projectNo', 'company'].includes(k) ? LABELS[k] : `${LABELS[k]} ${a[k] === undefined || a[k] === null || a[k] === '' ? '∅' : a[k]}→${b[k] === undefined || b[k] === null || b[k] === '' ? '∅' : b[k]}`);
    });
    if (REM.paymentSentence(a.payment) !== REM.paymentSentence(b.payment)) out.push('付款方式');
    if (JSON.stringify(REM.effectiveClauses(a)) !== JSON.stringify(REM.effectiveClauses(b))) out.push('追加條款');
    if (!same((a.products || []).slice().sort().join(','), (b.products || []).slice().sort().join(','))) out.push(`商品 ${(a.products || []).join('/') || '∅'}→${(b.products || []).join('/') || '∅'}`);
    const oldBy = new Map((a.items || []).map(it => [it.lid, it]));
    const seen = new Set(); let added = 0; const lineChanges = [];
    (b.items || []).forEach(it => {
      seen.add(it.lid);
      const o = oldBy.get(it.lid);
      if (!o) { added++; return; }
      const ch = [];
      if (QI.isNonItemRow(o) !== QI.isNonItemRow(it) || (QI.isNonItemRow(it) && o.kind !== it.kind)) return;   // 種類改變：下面統一記「列類型」，不逐欄比數量單價成本
      if (!same(o.desc, it.desc)) ch.push(isKindRowText(it) ? '文字' : '品名');
      if (!same(o.qty, it.qty)) ch.push(`數量 ${o.qty}→${it.qty}`);
      if (!same(o.unitPrice, it.unitPrice)) ch.push(`單價 ${o.unitPrice}→${it.unitPrice}`);
      if (!same(o.cost, it.cost)) ch.push(`成本 ${o.cost}→${it.cost}`);
      if (!same(o.unit, it.unit)) ch.push('單位');
      if (!same(o.cat, it.cat)) ch.push(`毛利分類 ${o.cat || '自動'}→${it.cat || '自動'}`);
      if (ch.length) lineChanges.push(`「${String(it.desc || '').slice(0, 12)}」${ch.join('、')}`);
    });
    const removed = (a.items || []).filter(it => !seen.has(it.lid)).length;
    const kindOf = (it) => (QI.isNonItemRow(it) ? it.kind : 'item');
    const kindChanged = (b.items || []).some(it => oldBy.has(it.lid) && kindOf(oldBy.get(it.lid)) !== kindOf(it));
    if (kindChanged) out.push('列類型（品項／分組標題／小計）');
    const hasKinds = (a.items || []).concat(b.items || []).some(QI.isNonItemRow);
    if (hasKinds && !kindChanged) {
      const orderOf = (arr) => (arr || []).map(it => it.lid).filter(l => seen.has(l) && oldBy.has(l)).join(',');
      const aOrder = (a.items || []).map(it => it.lid).filter(l => seen.has(l)).join(',');
      if (aOrder !== orderOf(b.items)) out.push('列順序');
    }
    if (added) out.push(`新增 ${added} 列`);
    if (removed) out.push(`刪除 ${removed} 列`);
    lineChanges.slice(0, 10).forEach(s => out.push(s));
    if (lineChanges.length > 10) out.push(`…另 ${lineChanges.length - 10} 列有變動`);
    // 成本明細（新式成本）：筆數＋前幾筆的品名；舊式單（兩邊都沒有 costLines）完全不加任何文字
    if (CL.hasCostLines(a) || CL.hasCostLines(b)) {
      if (!CL.hasCostLines(a)) out.push('成本改用「成本明細」格式');
      else if (!CL.hasCostLines(b)) out.push('成本明細已移除');
      const cnt = CL.diffCounts(a.costLines || [], b.costLines || []);
      const det = CL.summarizeChanges(a.costLines || [], b.costLines || [], 5);
      if (det.length) out.push(`成本明細（新增 ${cnt.added}、刪除 ${cnt.removed}、修改 ${cnt.changed}）：${det.join('；')}`);
    }
    return out.join('；').slice(0, 900);
  }

  // ═════════════════════════════════════════════════════════
  // 設定／報價章
  // ═════════════════════════════════════════════════════════
  function catalogList(data) {
    const cat = data.productCatalog || productCatalogLib.SEED_PRODUCT_CATALOG;
    const out = []; const bus = (cat && cat.bus) || {};
    Object.keys(bus).forEach(bu => {
      (Array.isArray(bus[bu]) ? bus[bu] : []).forEach(g => {
        if (!g || g.enabled === false) return;
        (g.items || []).forEach(it => { if (it && it.enabled !== false && it.name) out.push({ bu, group: g.group, name: it.name }); });
      });
    });
    return out;
  }

  app.get('/api/quote-approval/config', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const { data, cfg, users, me, role } = ctx;
    const isAdmin = role === 'admin';
    const out = {
      rulesVersion: QA.RULES_VERSION,
      rows: QA.ROWS.map(r => ({ key: r.key, label: r.label, l1: r.l1 === undefined ? null : r.l1, l2: r.l2 === undefined ? null : r.l2 })),
      classLabels: QA.CLASS_LABELS,
      amount: QA.AMOUNT,
      catalog: catalogList(data),
      productClasses: cfg.productClasses,
      costProviders: cfg.roster.costProviders.filter(un => canAct(users[un]) && un !== me).map(un => ({ username: un, displayName: dispName(users, un) })),
      me: {
        username: me, isAdmin,
        isSealManager: sealManagerSet(cfg, users).has(me),
        isCostProvider: cfg.roster.costProviders.includes(me),
        isBoardProxy: boardProxySet(cfg, users).has(me),
        isGm: cfg.roster.gm.includes(me), isChairman: cfg.roster.chairman.includes(me),
      },
      hasSeal: !!(data.quoteSeal && data.quoteSeal.base64),
      // Remarks 的「活的」部分（付款方式範本、付款期限選項與上限）：表單從這裡取，不在前端重複硬編碼
      remarks: {
        paymentPresets: REM.PAYMENT_PRESETS, netOptions: REM.NET_OPTIONS, defaultPayment: REM.DEFAULT_PAYMENT,
        maxInstallments: REM.MAX_INSTALLMENTS, maxLabel: REM.MAX_LABEL, maxClauses: REM.MAX_EXTRA, maxClauseLen: REM.MAX_CLAUSE,
        fixed: REM.FIXED,
      },
    };
    if (isAdmin) {
      out.roster = cfg.roster;
      out.users = Object.values(users).filter(u => u.role !== 'pool').map(u => ({ username: u.username, displayName: dispName(users, u.username), role: u.role, active: u.active !== false }));
    }
    res.json(out);
  });

  app.put('/api/admin/quote-approval/config', qAdmin, (req, res) => {
    const ctx = loadCtx(req);
    const { data, users } = ctx;
    const b = req.body || {};
    const r = b.roster || {};
    const roster = {};
    // 簽核/代核/填成本的名單不收管理員帳號：管理員只能改派、不能代簽（否則管理員把自己加進名單就能核准任何單）
    const SIGNING_KEYS = ['gm', 'chairman', 'boardProxy', 'costProviders'];
    for (const key of ['gm', 'chairman', 'boardProxy', 'costProviders', 'sealManagers']) {
      const list = Array.isArray(r[key]) ? r[key] : [];
      if (list.length > 100) return fail(res, 400, 'BAD_ROSTER', '名單人數過多');
      const clean = [];
      for (const un of list) {
        if (typeof un !== 'string' || !users[un]) return fail(res, 400, 'BAD_ROSTER', `名單中有不存在的帳號：${String(un).slice(0, 40)}`);
        if (SIGNING_KEYS.includes(key) && (users[un].role === 'admin' || users[un].role === 'pool')) {
          return fail(res, 400, 'ADMIN_IN_ROSTER', `「${dispName(users, un)}」是系統管理員帳號，不能列入簽核／成本填寫名單（管理員只能改派，不能代簽；請改用一般角色的帳號）`);
        }
        if (!clean.includes(un)) clean.push(un);
      }
      roster[key] = clean;
    }
    const pcIn = (b.productClasses && typeof b.productClasses === 'object') ? b.productClasses : {};
    const names = Object.keys(pcIn);
    if (names.length > 1000) return fail(res, 400, 'BAD_CLASSES', '商品歸類筆數過多');
    const productClasses = {};
    for (const name of names) {
      const v = pcIn[name] || {};
      if (!name || name.length > 100) return fail(res, 400, 'BAD_CLASSES', '商品名稱不合法');
      if (!QA.CLASS_LABELS[v.cls]) return fail(res, 400, 'BAD_CLASSES', `商品「${name}」的類別不合法`);
      productClasses[name] = { cls: v.cls, costBySales: v.costBySales === true };
    }
    const prevRoster = ctx.cfg.roster;
    data.quoteApproval = { roster, productClasses, updatedAt: nowIso(), updatedBy: ctx.me };
    db.save(data);
    // 名冊異動要能追查「誰有過簽核權」：記錄每份名單新增/移除的帳號
    const KEY_LABEL = { gm: '總經理', chairman: '董事長', boardProxy: '董事會代核', costProviders: '成本填寫人', sealManagers: '章管理' };
    const rosterDiff = Object.keys(KEY_LABEL).map(k => {
      const add = roster[k].filter(u => !prevRoster[k].includes(u)), del = prevRoster[k].filter(u => !roster[k].includes(u));
      return (add.length || del.length) ? `${KEY_LABEL[k]}${add.length ? ' +' + add.join(',') : ''}${del.length ? ' -' + del.join(',') : ''}` : '';
    }).filter(Boolean).join('；');
    writeLog('UPDATE_QUOTE_APPROVAL_CONFIG', ctx.me, '簽核設定',
      `總經理${roster.gm.length}人、董事長${roster.chairman.length}人、董事會代核${roster.boardProxy.length}人、成本填寫人${roster.costProviders.length}人、章管理${roster.sealManagers.length}人；商品歸類${names.length}筆` +
      (rosterDiff ? `｜名單異動：${rosterDiff}` : ''), req);
    res.json({ success: true });
  });

  app.put('/api/quote-approval/seal', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    if (!sealManagerSet(ctx.cfg, ctx.users).has(ctx.me)) return fail(res, 403, 'NO_PERMISSION', '只有報價章管理人（管理部秘書）可以上傳報價專用章');
    const m = /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+\/=]+)$/.exec(String((req.body || {}).dataUrl || ''));
    if (!m) return fail(res, 400, 'BAD_IMAGE', '請上傳 PNG 或 JPEG 圖檔');
    const buf = Buffer.from(m[2], 'base64');
    if (buf.length > SEAL_MAX_BYTES) return fail(res, 400, 'IMAGE_TOO_LARGE', '圖檔過大（上限 300KB），請縮小後再上傳');
    // 與匯出蓋章用同一套完整驗證（含 PNG 的 IEND、JPEG 的 EOI、寬高上限）：
    // 只看檔頭的話，截斷/損壞的圖會上傳成功，之後所有核准單匯出時靜默不蓋章。
    const info = quoteExcel._internal.sniffImage(buf);
    if (!info) return fail(res, 400, 'BAD_IMAGE', '圖檔不完整或已損壞（無法作為報價專用章），請重新匯出圖檔後再上傳');
    if ((m[1] === 'image/png') !== (info.ext === 'png')) return fail(res, 400, 'BAD_IMAGE', '檔案內容與圖片格式不符');
    ctx.data.quoteSeal = { mime: m[1], base64: m[2], uploadedBy: ctx.me, uploadedAt: nowIso() };
    db.save(ctx.data);
    writeLog('UPLOAD_QUOTE_SEAL', ctx.me, '報價專用章', `${m[1]} ${buf.length} bytes`, req);
    res.json({ success: true });
  });

  app.delete('/api/quote-approval/seal', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    if (!sealManagerSet(ctx.cfg, ctx.users).has(ctx.me)) return fail(res, 403, 'NO_PERMISSION', '只有報價章管理人可以刪除報價專用章');
    delete ctx.data.quoteSeal;
    db.save(ctx.data);
    writeLog('DELETE_QUOTE_SEAL', ctx.me, '報價專用章', '刪除', req);
    res.json({ success: true });
  });

  function sendSeal(res, seal) {
    if (!seal || !seal.base64) return res.status(404).json({ error: '尚未上傳報價專用章' });
    res.setHeader('Content-Type', seal.mime);
    res.setHeader('Cache-Control', 'private, no-cache');
    res.send(Buffer.from(seal.base64, 'base64'));
  }
  // 章圖是「正式報價單」的真偽憑證：只給章管理人（設定頁預覽用）。
  // 一般使用者要看「已核准單上的章」→ 走 /api/quotations/:id/seal（需對該單有檢視權且核准有效）。
  app.get('/api/quote-approval/seal', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    if (!sealManagerSet(ctx.cfg, ctx.users).has(ctx.me)) return fail(res, 403, 'NO_PERMISSION', '只有報價章管理人可以檢視報價專用章');
    sendSeal(res, ctx.data.quoteSeal);
  });

  // ═════════════════════════════════════════════════════════
  // 清單／收件匣／單筆
  // ═════════════════════════════════════════════════════════
  app.get('/api/quotations', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const list = ctx.data.quotations
      .filter(q => relationOf(q, ctx).canView)
      .sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''))
      .map(q => serialize(q, ctx));
    res.json(list);
  });

  // 必須宣告在 /:id 之前
  app.get('/api/quotations/inbox', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const approvals = []; const costs = [];
    ctx.data.quotations.forEach(q => {
      const rel = relationOf(q, ctx);
      if (!rel.canView) return;
      const perm = permOf(rel, q, ctx);
      const ap = q.approval;
      if (perm.canApprove && ap) {
        const step = ap.steps[ap.cur];
        approvals.push({ id: q.id, quoteNo: q.quoteNo, company: q.company, tierLabel: step ? step.label : '', submittedByName: dispName(ctx.users, ap.submittedBy), submittedAt: ap.submittedAt });
      }
      if (rel.isCostBy && q.costFlow && q.costFlow.state === 'requested' && !(ap && (ap.state === 'pending' || ap.state === 'approved'))) {
        costs.push({ id: q.id, quoteNo: q.quoteNo, company: q.company, requestedAt: q.costFlow.requestedAt });
      }
    });
    res.json({ approvals, costs, counts: { approvals: approvals.length, costs: costs.length } });
  });

  app.get('/api/quotations/:id', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    if (!relationOf(q, ctx).canView) return fail(res, 403, 'NO_PERMISSION', '無權限');
    res.json(serialize(q, ctx));
  });

  // 「已核准且內容未變」的報價單上的章（預覽畫面用）；其他情況一律不給圖
  app.get('/api/quotations/:id/seal', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(q, ctx);
    if (!rel.canView || rel.consultantOnly) return fail(res, 403, 'NO_PERMISSION', '無權限');
    const ap = q.approval;
    if (!(ap && ap.state === 'approved' && QA.contentHash(q) === ap.hash)) return fail(res, 404, 'NO_SEAL', '這張報價單尚未核准（或核准已作廢），沒有報價專用章');
    sendSeal(res, ctx.data.quoteSeal);
  });

  // ═════════════════════════════════════════════════════════
  // 新增／修改／刪除
  // ═════════════════════════════════════════════════════════
  function applyCostBy(q, body, ctx) {
    if (!('costBy' in body)) return null;
    const v = body.costBy ? String(body.costBy) : null;
    if (v) {
      if (v === q.owner) return '不能選自己當支援顧問（成本必須由顧問單位獨立提供）';
      if (!(ctx.cfg.roster.costProviders.includes(v) && canAct(ctx.users[v]))) return '所選的支援顧問不在成本填寫人名單內';
    }
    q.costBy = v;
    return null;
  }
  const tooManyItems = (b) => Array.isArray(b.items) && b.items.length > MAX_ITEMS;
  /**
   * 報價期限（Remarks 第 4 條「本報價單於 X 年 X 月 X 日前有效」）由業務自行輸入：
   * 空白（''、null、未提供）→ 預設值（新規則 quoteExcel.defaultValidUntil：報價日期當月的最後一個工作天；
   * 報價日期之後（不含當天）剩餘不足 7 個工作天就順延到下個月的最後一個工作天）；必須是字串、真實存在的日期，且不早於報價日期。
   * 不做靜默轉換：陣列／物件／數字、尾巴多出來的字元（'2026-10-30xyz'）一律拒絕。回傳 {value} 或 {error}。
   */
  function resolveValidUntil(raw, quoteDate) {
    const base = quoteExcel.isRealIsoDate(quoteDate) ? quoteDate : taipeiToday();
    if (raw !== undefined && raw !== null && typeof raw !== 'string') return { error: '報價期限必須是日期字串（YYYY-MM-DD）' };
    const s = String(raw == null ? '' : raw).trim();
    const v = s || quoteExcel.defaultValidUntil(base);
    if (!quoteExcel.isRealIsoDate(v)) return { error: '報價期限必須是正確的日期（YYYY-MM-DD）' };
    if (v < base) return { error: '報價期限不能早於報價日期' };
    return { value: v };
  }

  /**
   * Remarks 的「活的」部分：付款方式（第 3 條，結構化期程）與追加條款（第 7 條起）。驗證失敗回 {code, msg}，成功回 null。
   * - 舊單（沒存 payment／extraClauses）：表單會把「等同舊內容」的值原樣送回來——這種情況不存。
   *   存了雜湊就會變，已核准的舊單會被迫作廢核准，但客戶單上印出來的條款其實完全沒變。
   * - 新格式的追加條款存在 extraClauses，單行舊欄位 note 同時清空，避免兩處並存、重複印出。
   */
  function applyRemarksFields(t, b) {
    if (b.payment !== undefined) {
      if (b.payment === null) delete t.payment;
      else {
        const r = REM.normalizePayment(b.payment);
        if (r.error) return { code: 'BAD_PAYMENT', msg: r.error };
        if (t.payment === undefined && REM.isDefaultPayment(r.value)) { /* 維持沒存（印預設句） */ }
        else t.payment = r.value;
      }
    }
    if (b.extraClauses !== undefined) {
      const legacy = REM.effectiveClauses({ note: t.note });
      // 舊單且送回來的就是原封不動的舊備註：先不檢查單條長度（舊 note 上限 500 字，追加條款上限 200 字），維持不存
      const same = t.extraClauses === undefined ? REM.normalizeClauses(b.extraClauses, { lenientLength: true }) : null;
      if (same && !same.error && JSON.stringify(same.value) === JSON.stringify(legacy)) { /* 維持舊的單行備註 */ }
      else {
        const r = REM.normalizeClauses(b.extraClauses);
        if (r.error) return { code: 'BAD_CLAUSES', msg: r.error };
        t.extraClauses = r.value; t.note = '';
      }
    }
    return null;
  }

  app.post('/api/quotations', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    if (!ctx.featureOk) return fail(res, 403, 'NO_PERMISSION', '你的角色沒有「報價單」功能權限');
    // 規格：報價單由業務發起；業務主管、總經理、董事長、秘書不建單
    if (NO_CREATE_ROLES.includes(ctx.role) || ctx.cfg.roster.gm.includes(ctx.me) || ctx.cfg.roster.chairman.includes(ctx.me)) {
      return fail(res, 403, 'NO_CREATE_ROLE', '報價單由業務發起；你的角色（主管／總經理／董事長／秘書等）不建立報價單');
    }
    const b = req.body || {};
    if (tooManyItems(b)) return fail(res, 400, 'TOO_MANY_ITEMS', `報價單品項最多 ${MAX_ITEMS} 列（範本列數上限），請拆成多張報價單`);
    const now = nowIso();
    const q = {
      id: uuidv4(), owner: ctx.me, quoteNo: '',
      contactId: sanitizeStr(b.contactId, 36), company: sanitizeStr(b.company, 100),
      contactName: sanitizeStr(b.contactName, 100), phone: sanitizeStr(b.phone, 50), mobile: sanitizeStr(b.mobile, 50),
      address: sanitizeStr(b.address, 200), quoteDate: sanitizeStr(b.quoteDate, 10) || taipeiToday(),
      projectName: sanitizeStr(b.projectName, 200), projectNo: sanitizeStr(b.projectNo, 50),
      note: sanitizeStr(b.note, 500),
      discountType: ['none', 'percent', 'amount'].includes(b.discountType) ? b.discountType : 'none',
      discountValue: Number.isFinite(parseFloat(b.discountValue)) && parseFloat(b.discountValue) > 0 ? parseFloat(b.discountValue) : 0,
      status: 'draft', products: normProducts(b.products) || [], costBy: null, approval: null,
      createdAt: now, updatedAt: now,
    };
    const err = applyCostBy(q, b, ctx);
    if (err) return fail(res, 400, 'BAD_COST_PROVIDER', err);
    const remErr = applyRemarksFields(q, b);
    if (remErr) return fail(res, 400, remErr.code, remErr.msg);
    const vuRes = resolveValidUntil(b.validUntil, q.quoteDate);
    if (vuRes.error) return fail(res, 400, 'BAD_VALID_UNTIL', vuRes.error);
    q.validUntil = vuRes.value;
    const acceptCost = !needsConsultant(q, ctx.cfg);
    // 成本明細（新式成本）：採用條件與 items[].cost 相同（不需顧問成本的單才由業務自填；需顧問的單送了也靜默忽略）。
    // 採用後 items[].cost 一律歸 0（成本只認成本明細，避免兩份成本）
    if (acceptCost && b.costLines !== undefined) {
      const cr = CL.normalizeCostLines(b.costLines, { genLid: uuidv4, prev: [] });
      if (!cr.ok) return fail(res, cr.error.status, cr.error.code, cr.error.message);
      q.costLines = cr.lines;
    }
    q.items = normalizeItems(Array.isArray(b.items) ? b.items : [], [], acceptCost && !CL.hasCostLines(q), false);
    // 新單第一次儲存：品項到這裡才拿到 lid，成本列的 forLid 還是空的 → 依品名回填（之後品項改名，「補入新品項」仍認得它已有成本列）
    // 順序：先把畫面給新品項的暫時代號（nid）換成真正的 lid，再用品名補剩下沒有任何 forLid 的列
    if (CL.hasCostLines(q)) q.costLines = CL.backfillForLids(q.items, CL.resolveItemRefs(q.costLines, tmpRefMap(b.items, q.items)));
    q.costFlow = { note: sanitizeStr(b.costNote, 500) };
    const { notify: nc } = reconcileCostFlow(q, null, ctx, now, false);
    // 所有驗證都通過、確定要存檔才給單號（Postgres 模式 db.load() 是共用快取，先給號再失敗會讓流水號永久跳號）
    q.quoteNo = genQuoteNo(ctx.data);
    ctx.data.quotations.push(q);
    db.save(ctx.data);
    writeLog('CREATE_QUOTATION', ctx.me, q.quoteNo, `${q.company}｜商品 ${q.products.length} 項`, req);
    if (nc) notify(ctx, ctx.me, [q.costBy], 'quote_cost_request', `🧾 請填寫報價成本：${q.quoteNo}`, `${noteLine(q, ctx)}　請開啟「報價單管理」填寫各品項成本。`, q.id);
    res.status(201).json(serialize(q, ctx));
  });

  app.put('/api/quotations/:id', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(q, ctx);
    if (!rel.canView) return fail(res, 403, 'NO_PERMISSION', '無權限');
    if (!(rel.isOwner || rel.isAdmin)) return fail(res, 403, 'NO_PERMISSION', '只有報價單擁有者可以修改');
    const st = q.approval ? q.approval.state : 'none';
    if (st === 'pending') return fail(res, 409, 'LOCKED_PENDING', '簽核進行中，整張報價單已鎖定；如需修改請先撤回送簽');
    const b = req.body || {};
    if (tooManyItems(b)) return fail(res, 400, 'TOO_MANY_ITEMS', `報價單品項最多 ${MAX_ITEMS} 列（範本列數上限），請拆成多張報價單`);
    const now = nowIso();

    const draft = JSON.parse(JSON.stringify(q));
    if ('company' in b) draft.company = sanitizeStr(b.company, 100) || draft.company;
    ['contactId:36', 'contactName:100', 'phone:50', 'mobile:50', 'address:200', 'projectName:200', 'projectNo:50', 'note:500'].forEach(spec => {
      const [k, n] = spec.split(':'); if (k in b) draft[k] = sanitizeStr(b[k], +n);
    });
    if ('quoteDate' in b) draft.quoteDate = sanitizeStr(b.quoteDate, 10) || draft.quoteDate;
    if ('discountType' in b && ['none', 'percent', 'amount'].includes(b.discountType)) draft.discountType = b.discountType;
    if ('discountValue' in b) { const v = parseFloat(b.discountValue); draft.discountValue = Number.isFinite(v) && v > 0 ? v : 0; }
    const prods = normProducts(b.products); if (prods) draft.products = prods;
    const err = applyCostBy(draft, b, ctx);
    if (err) return fail(res, 400, 'BAD_COST_PROVIDER', err);
    // 已改用新格式追加條款的單，舊版前端若還送單行 note，忽略它（印的是 extraClauses；note 已清空）
    if (q.extraClauses !== undefined && !('extraClauses' in b)) draft.note = q.note;
    const remErr = applyRemarksFields(draft, b);
    if (remErr) return fail(res, 400, remErr.code, remErr.msg);
    // 報價期限：前端有送就驗證後採用；沒送（舊版前端）就維持原值，但不可讓改日期後的報價日期晚於既有期限
    if ('validUntil' in b) {
      // 沒存期限的舊單收到空白（只有直接打 API 才會；表單送不出空白）：維持改版前行為，用【舊規則】推算，
      // 否則會被存成新規則的日期，舊單的列印日期就被悄悄改掉。新規則只管新單與已經存了期限的單。
      const vuBlank = b.validUntil === undefined || b.validUntil === null || (typeof b.validUntil === 'string' && !b.validUntil.trim());
      const vuBase = quoteExcel.isRealIsoDate(draft.quoteDate) ? draft.quoteDate : taipeiToday();
      const vuRes = resolveValidUntil(!draft.validUntil && vuBlank ? quoteExcel.legacyDefaultValidUntil(vuBase) : b.validUntil, draft.quoteDate);
      if (vuRes.error) return fail(res, 400, 'BAD_VALID_UNTIL', vuRes.error);
      // 功能上線前建立、沒有存報價期限的舊單：表單會把「系統推算的預設值」原樣送回來。這種情況不要存：
      // 存了雜湊就會變，已核准的舊單會被迫作廢核准，但客戶單上印出來的日期其實完全沒變（推算值固定取自建立日期）。
      // ⚠ 這裡比對的 effectiveValidUntil 刻意走【舊規則】（quoteExcel.legacyDefaultValidUntil）——舊單列印日期不能因新規則改變；
      //   不要改成新的 defaultValidUntil，否則舊單原樣送回的值對不上、會被存成實值而作廢核准。
      // 業務真的改成別的日期才會存（那是實質變更，照常需要重新簽核）。
      if (!draft.validUntil && vuRes.value === quoteExcel.effectiveValidUntil(draft)) { /* 維持「沒存」 */ }
      else draft.validUntil = vuRes.value;
    } else if (draft.validUntil && quoteExcel.isRealIsoDate(draft.quoteDate) && draft.validUntil < draft.quoteDate) {
      return fail(res, 400, 'BAD_VALID_UNTIL', '報價期限不能早於報價日期，請一併調整報價期限');
    }
    if ('costNote' in b) draft.costFlow = Object.assign({}, draft.costFlow, { note: sanitizeStr(b.costNote, 500) });

    // 「需顧問成本」是由業務可改的商品勾選決定的 → 切換（需顧問↔不需顧問）時，整批成本作廢重來：
    // 否則業務只要把商品清空就能讀到顧問填的成本；反向則能把自己寫的成本塞進顧問負責的欄位。
    const needBefore = needsConsultant(q, ctx.cfg), needAfter = needsConsultant(draft, ctx.cfg);
    const flip = needBefore !== needAfter;
    const touchedAfter = !flip && consultantTouched(draft);          // 切換後顧問的紀錄已重置
    const acceptCost = !needAfter && (rel.isOwner || rel.isAdmin) && !touchedAfter;
    // 舊版畫面（部署前就開著的分頁）不認得分組標題／小計列：送回來的列沒有 kind，伺服器會把它們存成「數量 1、單價 0 的品項」。
    // 新版畫面送 rowKinds:1 表示它看得懂；單子裡已有標題／小計列卻沒帶這個旗標 → 擋下並請使用者重新整理。
    if (Array.isArray(b.items) && (q.items || []).some(QI.isNonItemRow) && b.rowKinds !== 1) {
      return fail(res, 409, 'CLIENT_OUTDATED', '畫面版本過舊，無法保存這張含「分組標題／小計列」的報價單。請重新整理頁面後再編輯。');
    }
    // 新分頁保護（成本明細）：單上已有 costLines，卻收到沒帶 costModel:2 的 items → 是舊畫面（不認得成本明細、會把成本寫回 items[].cost），擋下
    if (Array.isArray(b.items) && CL.hasCostLines(q) && b.costModel !== 2) {
      return fail(res, 409, 'CLIENT_OUTDATED', '畫面版本過舊，無法保存這張已改用「成本明細」的報價單。請重新整理頁面後再編輯。');
    }
    // 成本明細（新式成本）：採用條件＝items[].cost 可被寫入的同一條件（acceptCost）；不符就靜默忽略（行為比照 cost）。
    // 需顧問↔不需顧問切換時整批重置：成本明細一併移除（寫入稽核），重置後才採用這次送來的明細。
    const hadLines = CL.hasCostLines(q);
    if (flip) delete draft.costLines;
    if (acceptCost && b.costLines !== undefined) {
      // 過期分頁保護：毛利頁籤載入時的成本明細摘要與目前不同（別的視窗改過）→ 409；沒帶摘要（舊分頁）不檢查
      if (typeof b.costLinesSig === 'string' && b.costLinesSig !== QA.costLinesSig(q)) {
        return fail(res, 409, 'STALE_COSTS', '成本明細已在其他視窗被更新，請關閉後重新整理');
      }
      const cr = CL.normalizeCostLines(b.costLines, { genLid: uuidv4, prev: flip ? [] : (q.costLines || []) });
      if (!cr.ok) return fail(res, cr.error.status, cr.error.code, cr.error.message);
      draft.costLines = cr.lines;
    }
    const linesAfter = CL.hasCostLines(draft);
    const switchedToLines = linesAfter && !hadLines;   // 舊式單第一次存成新式：items[].cost 全部歸 0
    // 新式成本：items[].cost 一律歸 0（不採用請求裡的 cost、也不沿用舊值）
    if (Array.isArray(b.items)) draft.items = normalizeItems(b.items, q.items, acceptCost && !linesAfter, flip || linesAfter);
    else if (flip || switchedToLines) draft.items = (draft.items || []).map(it => (QI.isNonItemRow(it) ? it : Object.assign({}, it, { cost: 0 })));
    // 這次請求寫入了成本明細：新增的品項此時才有 lid，沒有 forLid／forLids 的成本列依品名回填（只動這次寫入的明細；沒帶 costLines 的請求不碰既有明細）
    if (acceptCost && b.costLines !== undefined && CL.hasCostLines(draft)) draft.costLines = CL.backfillForLids(draft.items, CL.resolveItemRefs(draft.costLines, Array.isArray(b.items) ? tmpRefMap(b.items, draft.items) : new Map()));
    if (flip) delete draft.contingencyPct;     // 成本整批重來，顧問先前預估的風險預留也一併作廢

    // 核准後改到雜湊涵蓋欄位（含換支援顧問）→ 作廢核准（需明確確認）
    const approvedNow = st === 'approved';
    const costByChanged = (draft.costBy || null) !== (q.costBy || null);
    const willVoid = approvedNow && (QA.contentHash(draft) !== q.approval.hash || costByChanged);
    if (willVoid && b.confirmVoid !== true) {
      return fail(res, 409, 'WILL_VOID', '此報價單已核准；修改價格、數量、成本（含成本明細）、折扣、商品、分組標題與小計列（含順序）、備註與追加條款、付款方式、報價期限或客戶聯絡資料會使核准作廢並需重新簽核。確定要修改請再次確認。');
    }
    const prevFlow = q.costFlow;
    const { notify: nc } = reconcileCostFlow(draft, prevFlow, ctx, now, flip);
    let voidedSigners = [];
    if (willVoid) {
      voidedSigners = (q.approval.steps || []).filter(s => s.status === 'approved' && s.by).map(s => s.by);
      const hist = q.approval.history || [];
      const voidMeta = auditMeta(q, q.approval);
      draft.approval = { state: 'none', rulesVersion: q.approval.rulesVersion, hash: null, submittedAt: null, submittedBy: null, derived: null, steps: [], cur: 0, board: null, history: hist };
      pushHistory(draft.approval, ctx.me, 'INVALIDATE', '核准後修改內容，核准作廢', '', voidMeta);
    }
    draft.updatedAt = now;
    const diff = diffSummary(q, draft);
    // 成本重置是靜默的（前端沒有警告）→ 把重置前各列成本留在稽核紀錄，萬一是誤操作（例如舊單補勾商品）還能追回
    const itemCostList = () => QI.itemRows(q.items).map((it, i) => [i + 1, Number(it.cost) || 0]).filter(x => x[1]);   // 列號只數一般品項（標題／小計列不編號）
    const lostCosts = flip ? itemCostList() : [];
    // 成本明細的重置也要留痕：flip 時移除的整批明細（完整列出每一列，上限 60 列＝資料上限，不是前 10 筆摘要）、
    // 舊式單第一次改成新式時被歸 0 的各列成本
    const lostLines = flip && hadLines ? CL.describeLinesFull(q.costLines) : '';
    const switchCosts = (!flip && switchedToLines) ? itemCostList() : [];
    const switchNote = (!flip && switchedToLines) ? '（成本改用成本明細格式，品項成本欄位已歸 0' + (switchCosts.length ? '；原各列成本：' + switchCosts.map(x => `第${x[0]}列=${x[1]}`).join('、') : '') + '）' : '';
    Object.keys(q).forEach(k => delete q[k]);
    Object.assign(q, draft);
    db.save(ctx.data);
    const flipNote = flip ? '（需顧問成本狀態切換，成本已重置' + (lostCosts.length ? '；重置前各列成本：' + lostCosts.map(x => `第${x[0]}列=${x[1]}`).join('、') : '') + (lostLines ? '；重置前成本明細：' + lostLines : '') + '）' : '';
    writeLog('UPDATE_QUOTATION', ctx.me, q.quoteNo, (willVoid ? '修改內容（核准已作廢）' : '修改內容') + flipNote + switchNote + (diff ? '｜' + diff : ''), req);
    if (nc) notify(ctx, ctx.me, [q.costBy], 'quote_cost_request', `🧾 請填寫報價成本：${q.quoteNo}`, `${noteLine(q, ctx)}　請開啟「報價單管理」填寫各品項成本。`, q.id);
    if (willVoid) notify(ctx, ctx.me, voidedSigners.concat([q.owner]), 'quote_voided', `⚠️ 報價單核准已作廢：${q.quoteNo}`, `${noteLine(q, ctx)}　內容被修改，需重新送簽。`, q.id);
    res.json(serialize(q, ctx));
  });

  app.delete('/api/quotations/:id', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(q, ctx);
    if (!(rel.isOwner || rel.isAdmin) || !rel.canView) return fail(res, 403, 'NO_PERMISSION', '無權限');
    const ap = q.approval;
    if (ap && (ap.state !== 'none' || (ap.history || []).length)) return fail(res, 409, 'HAS_APPROVAL', '已送過簽核的報價單不可刪除（簽核紀錄需保留）');
    ctx.data.quotations = ctx.data.quotations.filter(x => x.id !== q.id);
    db.save(ctx.data);
    writeLog('DELETE_QUOTATION', ctx.me, q.quoteNo, q.company || '', req);
    res.json({ success: true });
  });

  // ═════════════════════════════════════════════════════════
  // 成本填寫
  // ═════════════════════════════════════════════════════════
  app.put('/api/quotations/:id/costs', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(q, ctx);
    if (!rel.canView) return fail(res, 403, 'NO_PERMISSION', '無權限');
    const perm = permOf(rel, q, ctx);
    if (!perm.canEditCost) {
      const st = q.approval ? q.approval.state : 'none';
      if (st === 'pending' || st === 'approved') return fail(res, 409, 'LOCKED_PENDING', '簽核中或已核准的報價單不能修改成本');
      return fail(res, 403, 'NO_PERMISSION', '你沒有填寫這張報價單成本的權限');
    }
    const b = req.body || {};
    // 過期分頁保護（新前端一律帶這兩個值；沒帶＝舊分頁／測試，不檢查，向下相容）：
    // 品項結構已被業務改過（沒看過新品項就存、尤其按「完成」會讓顧問漏填新品項的成本），或成本明細已在其他視窗被更新（整批覆寫會吃掉對方的新增／刪除）→ 409
    if (typeof b.itemsSig === 'string' && b.itemsSig !== QA.itemsSig(q)) {
      return fail(res, 409, 'STALE_ITEMS', '報價品項已被業務修改，請關閉後重新整理，確認成本明細後再送出');
    }
    if (typeof b.costLinesSig === 'string' && b.costLinesSig !== QA.costLinesSig(q)) {
      return fail(res, 409, 'STALE_COSTS', '成本明細已在其他視窗被更新，請關閉後重新整理');
    }
    const items = q.items || [];
    const hadLines = CL.hasCostLines(q);
    // 新格式（成本明細）：body 帶 costLines；或單上已是新式、body 沒帶舊格式 costs（只改 done／風險預留，明細沿用）。
    // 舊格式 {costs:[{lid,cost}]} 只在「單上還沒有 costLines」時接受：舊分頁不認得成本明細，硬寫會把成本寫進不再計算的 items[].cost
    const linesMode = b.costLines !== undefined || (hadLines && b.costs === undefined);
    if (hadLines && !linesMode) return fail(res, 409, 'CLIENT_OUTDATED', '畫面版本過舊，無法保存這張已改用「成本明細」的報價單的成本。請重新整理頁面後再填寫。');
    let newLines = null;       // 新格式：驗證並正規化後的成本明細
    if (linesMode) {
      if (b.costLines !== undefined) {
        const cr = CL.normalizeCostLines(b.costLines, { genLid: uuidv4, prev: q.costLines || [] });
        if (!cr.ok) return fail(res, cr.error.status, cr.error.code, cr.error.message);
        // 這條路徑的品項都已有 lid：暫時代號換不到（丟掉）；顧問手動新增的列（沒有 forLid）依品名回填，已有 forLid／forLids 的不動
        newLines = CL.backfillForLids(items, CL.resolveItemRefs(cr.lines, new Map()));
      } else newLines = q.costLines;
    }
    // 先全部驗證、算出新值，通過後才一次套用（驗證失敗不能留下半套修改；Postgres 模式 db.load() 是共用快取）
    const updates = new Map();
    const byLid = new Map(items.map(it => [it.lid, it]));
    if (!linesMode) {
      // 品項代碼（lid）必須唯一：重複時成本只能寫到最後一列，前面的列永遠無人能改
      if (new Set(items.map(it => it.lid)).size !== items.length) return fail(res, 409, 'DUP_LID', '品項代碼重複，請請業務重新儲存報價單後再填成本');
      const costs = Array.isArray(b.costs) ? b.costs : [];
      for (const c of costs) {
        const it = c && byLid.get(c.lid);
        if (!it || QI.isNonItemRow(it)) return fail(res, 400, 'BAD_LINE', '成本對應的品項不存在，請重新整理後再填');
        const v = parseFloat(c.cost);
        if (!Number.isFinite(v) || v < 0 || v > 1e12) return fail(res, 400, 'BAD_COST', '成本必須是 0 以上的數字');
        updates.set(it.lid, v);
      }
    }
    const now = nowIso();
    const needC = needsConsultant(q, ctx.cfg);
    const costOf = (it) => (updates.has(it.lid) ? updates.get(it.lid) : (parseFloat(it.cost) || 0));
    // 風險預留：只有「顧問填成本」的單才有意義，且只有負責填成本的人（顧問主管）能設定；沒送欄位＝不變；null 或空字串＝清除。
    // 刻意嚴格比對型別（[]、true 不可被 Number() 轉成 0 而通過）
    let riskUpdate;                       // undefined＝不變；null＝清除；number＝設定
    if (needC && b.contingencyPct !== undefined) {
      const raw = b.contingencyPct;
      if (raw === null || raw === '') riskUpdate = null;
      else if (typeof raw === 'number' || (typeof raw === 'string' && raw.trim() !== '')) {
        const n = Number(raw);
        if (!CONTINGENCY_PCTS.includes(n)) return fail(res, 400, 'BAD_CONTINGENCY', `風險預留必須是 ${CONTINGENCY_PCTS.join('、')} 其中之一（%）`);
        riskUpdate = n;
      } else return fail(res, 400, 'BAD_CONTINGENCY', `風險預留必須是 ${CONTINGENCY_PCTS.join('、')} 其中之一（%）`);
    }
    if (needC && b.done === true) {
      if (linesMode) {
        // 新式：每列 desc 非空（正規化時已擋）、非印花稅列 qty>0；至少 1 列非印花稅列且其總成本>0（單價 0 的列允許，前端會確認）
        const dc = CL.checkDone(newLines);
        if (!dc.ok) return fail(res, 400, dc.code, dc.message, dc.code === 'MISSING_COST' ? { missing: [] } : undefined);
      } else {
        const missing = items.filter(it => (parseFloat(it.unitPrice) || 0) > 0 && !(costOf(it) > 0));
        if (missing.length) return fail(res, 400, 'MISSING_COST', `還有 ${missing.length} 個品項沒有填成本（成本需大於 0）`, { missing: missing.map(it => it.lid) });
      }
    }
    const changes = [];
    let switchNote = '';
    if (linesMode) {
      CL.summarizeChanges(q.costLines || [], newLines, 1000).forEach(s => changes.push(s));
      if (!hadLines) {
        // 舊式單第一次存成新式：items[].cost 全部歸 0（避免兩份成本），歸 0 前的各列成本寫進稽核
        const lost = QI.itemRows(items).map((it, i) => [i + 1, Number(it.cost) || 0]).filter(x => x[1]);
        items.forEach(it => { if (!QI.isNonItemRow(it)) it.cost = 0; });
        switchNote = '（成本改用成本明細格式，品項成本欄位已歸 0' + (lost.length ? '；原各列成本：' + lost.map(x => `第${x[0]}列=${x[1]}`).join('、') : '') + '）';
      }
      q.costLines = newLines;
    } else {
      updates.forEach((v, lid) => {
        const it = byLid.get(lid);
        if ((parseFloat(it.cost) || 0) !== v) changes.push(`「${String(it.desc || '').slice(0, 12)}」${it.cost || 0}→${v}`);
        it.cost = v;
      });
    }
    if (riskUpdate !== undefined) {
      const before = typeof q.contingencyPct === 'number' ? q.contingencyPct : null;
      // 放最前面：稽核摘要只留前 10 筆，成本列很多時不能把風險預留的設定／變更擠掉
      if (before !== riskUpdate) changes.unshift(`風險預留 ${before === null ? '∅' : before + '%'}→${riskUpdate === null ? '∅' : riskUpdate + '%'}`);
      if (riskUpdate === null) delete q.contingencyPct; else q.contingencyPct = riskUpdate;
    }
    if (needC) q.costFlow = Object.assign({}, q.costFlow, { consultantWrote: true });
    if (needC && b.done === true) {
      q.costFlow = Object.assign({}, q.costFlow, { state: 'filled', filledAt: now, sig: QA.lineStructureSig(items) });
    } else if (needC && b.done === false && q.costFlow && q.costFlow.state === 'filled') {
      q.costFlow = Object.assign({}, q.costFlow, { state: 'requested', filledAt: null });
    }
    if (!needC && rel.isOwner && typeof b.note === 'string') q.costFlow = Object.assign({}, q.costFlow, { note: sanitizeStr(b.note, 500) });
    q.updatedAt = now;
    db.save(ctx.data);
    writeLog('FILL_QUOTE_COST', ctx.me, q.quoteNo,
      (needC ? (b.done === true ? '顧問完成成本填寫' : '顧問儲存成本') : '業務填寫成本') + switchNote + (changes.length ? `｜${changes.slice(0, 10).join('；')}${changes.length > 10 ? `…另 ${changes.length - 10} ${linesMode ? '筆' : '列'}` : ''}` : ''), req);
    if (needC && b.done === true) notify(ctx, ctx.me, [q.owner], 'quote_cost_done', `✅ 成本已填寫完成：${q.quoteNo}`, `${noteLine(q, ctx)}　顧問 ${dispName(ctx.users, ctx.me)} 已完成成本，可以送簽了。`, q.id);
    res.json(serialize(q, ctx));
  });

  // ═════════════════════════════════════════════════════════
  // 簽核：送簽／撤回／核准／駁回／改派
  // ═════════════════════════════════════════════════════════
  app.post('/api/quotations/:id/submit', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(q, ctx);
    if (!rel.isOwner) return fail(res, 403, 'NO_PERMISSION', '只有報價單擁有者可以送簽');
    const st = q.approval ? q.approval.state : 'none';
    if (st === 'pending') return fail(res, 409, 'ALREADY_PENDING', '這張報價單已在簽核中');
    if (st === 'approved') return fail(res, 409, 'ALREADY_APPROVED', '這張報價單已核准；如需修改請先編輯（會使核准作廢）');
    if (q.costBy && q.costBy === q.owner) return fail(res, 400, 'SELF_COST_PROVIDER', '支援顧問不能是報價單擁有者本人，請改選其他顧問');
    if (new Set((q.items || []).map(it => it.lid)).size !== (q.items || []).length) return fail(res, 409, 'DUP_LID', '品項代碼重複，請重新儲存報價單後再送簽');
    const pv = buildPreview(q, ctx);
    if (pv.blockers.length || !pv.tiers) {
      const blockers = pv.blockers.length ? pv.blockers : [{ code: 'COST_NOT_FILLED', message: '成本尚未填妥' }];
      return fail(res, 400, 'CANNOT_SUBMIT', '目前還不能送簽：' + blockers.map(x => x.message).join('；'), { blockers });
    }
    const derived = QA.buildDerived(q, ctx.cfg.productClasses);
    const mgr1 = findManager1(ctx.users, q.owner);
    const now = nowIso();
    const steps = derived.tiers.map((t, i) => ({
      tier: t, label: QA.TIERS[t], assignee: t === 'mgr1' ? mgr1 : null,
      status: i === 0 ? 'pending' : 'waiting', by: null, at: null, comment: '',
    }));
    const prevHistory = (q.approval && q.approval.history) || [];
    q.approval = {
      state: 'pending', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q),
      submittedAt: now, submittedBy: ctx.me, derived, steps, cur: 0, board: null, history: prevHistory,
    };
    pushHistory(q.approval, ctx.me, 'SUBMIT', '', '', auditMeta(q, q.approval));
    q.updatedAt = now;
    db.save(ctx.data);
    writeLog('SUBMIT_QUOTATION', ctx.me, q.quoteNo, `${derived.rowLabel}｜毛利率 ${derived.marginText}%｜關卡 ${derived.tiers.join('→')}`, req);
    notify(ctx, ctx.me, stepRecipients(steps[0], ctx), 'quote_submitted', `📝 報價單待簽核：${q.quoteNo}`, `${noteLine(q, ctx)}　目前關卡：${steps[0].label}`, q.id);
    res.json(serialize(q, ctx));
  });

  app.post('/api/quotations/:id/withdraw', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    if (q.owner !== ctx.me) return fail(res, 403, 'NO_PERMISSION', '只有報價單擁有者可以撤回');
    const ap = q.approval;
    if (!ap || ap.state !== 'pending') return fail(res, 409, 'NOT_PENDING', '這張報價單目前不在簽核中');
    const step = ap.steps[ap.cur];
    const recipients = stepRecipients(step, ctx).concat(ap.steps.filter(s => s.status === 'approved' && s.by).map(s => s.by));
    const wdMeta = auditMeta(q, ap);       // 清掉 derived 之前先留下撤回當下的快照
    ap.state = 'none'; ap.steps = []; ap.cur = 0; ap.hash = null; ap.derived = null; ap.board = null;
    pushHistory(ap, ctx.me, 'WITHDRAW', '', '', wdMeta);
    q.updatedAt = nowIso();
    db.save(ctx.data);
    writeLog('WITHDRAW_QUOTATION', ctx.me, q.quoteNo, '撤回送簽', req);
    notify(ctx, ctx.me, recipients, 'quote_withdrawn', `↩️ 報價單已撤回：${q.quoteNo}`, `${noteLine(q, ctx)}　業務已撤回送簽，不需再處理。`, q.id);
    res.json(serialize(q, ctx));
  });

  /** 核准／駁回共用的前置檢查；通過回傳 {q, ap, step}，否則已回應錯誤並回傳 null */
  function actionGuard(req, res, ctx) {
    const q = findQuote(ctx, req.params.id);
    if (!q) { fail(res, 404, 'NOT_FOUND', '找不到此報價單'); return null; }
    const ap = q.approval;
    if (!ap || ap.state !== 'pending') { fail(res, 409, 'NOT_PENDING', '這張報價單目前不在簽核中'); return null; }
    const step = ap.steps[ap.cur];
    if (!step || step.status !== 'pending') { fail(res, 409, 'NOT_PENDING', '簽核關卡狀態異常，請重新整理'); return null; }
    if (!stepActorOk(step, ctx.me, ctx.cfg, ctx.users)) { fail(res, 403, 'NOT_YOUR_STEP', `目前輪到「${step.label}」簽核，你不是這一關的簽核人`); return null; }
    if (ctx.me === q.owner || ctx.me === ap.submittedBy) { fail(res, 403, 'SELF_APPROVAL', '不能簽核自己發起的報價單'); return null; }
    if ((q.costFlow && q.costFlow.by === ctx.me) || (q.costBy && q.costBy === ctx.me)) { fail(res, 403, 'SELF_COST_APPROVAL', '你是這張報價單的成本填寫人，不能再簽核它'); return null; }
    if (alreadySigned(ap, ctx.me)) { fail(res, 403, 'SAME_APPROVER', '同一個人不能簽核同一張報價單的兩個關卡'); return null; }
    const cur = QA.contentHash(q);
    if (!req.body || req.body.hash !== cur || cur !== ap.hash) {
      fail(res, 409, 'CONTENT_CHANGED', '報價單內容已變動，請重新檢視後再簽核'); return null;
    }
    return { q, ap, step };
  }

  app.post('/api/quotations/:id/approve', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const g = actionGuard(req, res, ctx);
    if (!g) return;
    const { q, ap, step } = g;
    const b = req.body || {};
    const now = nowIso();
    const comment = sanitizeStr(b.comment, 500);
    let boardInfo = null;
    if (step.tier === 'board') {
      const d = sanitizeStr(b.resolutionDate, 10);
      const no = sanitizeStr(b.resolutionNo, 50);
      // 必須是真實存在的日期（2026-02-31 這種不行），且不可是未來日期（決議已經開過了才會來登錄）
      if (!isRealDate(d)) return fail(res, 400, 'BAD_RESOLUTION', '請填寫正確的董事會決議日期（YYYY-MM-DD，必須是存在的日期）');
      if (d > taipeiToday()) return fail(res, 400, 'BAD_RESOLUTION', '董事會決議日期不能是未來的日期');
      if (!no) return fail(res, 400, 'BAD_RESOLUTION', '請填寫董事會決議文號');
      boardInfo = { resolutionDate: d, resolutionNo: no, by: ctx.me, at: now };
      ap.board = boardInfo;
    }
    step.status = 'approved'; step.by = ctx.me; step.at = now; step.comment = comment;
    pushHistory(ap, ctx.me, 'APPROVE', comment, step.tier, Object.assign(auditMeta(q, ap), boardInfo ? { resolutionDate: boardInfo.resolutionDate, resolutionNo: boardInfo.resolutionNo } : {}));
    let final = false;
    if (ap.cur >= ap.steps.length - 1) {
      final = true; ap.state = 'approved'; ap.cur = ap.steps.length; ap.hash = QA.contentHash(q);
    } else {
      ap.cur += 1; ap.steps[ap.cur].status = 'pending';
    }
    q.updatedAt = now;
    db.save(ctx.data);
    writeLog('APPROVE_QUOTATION', ctx.me, q.quoteNo, `${step.label}核准${final ? '（全部完成）' : ''}${ap.board ? `｜決議 ${ap.board.resolutionNo}` : ''}`, req);
    if (final) {
      notify(ctx, ctx.me, [q.owner], 'quote_approved', `✅ 報價單已核准：${q.quoteNo}`, `${noteLine(q, ctx)}　已完成所有簽核，請下載 PDF（PDF 會蓋上報價專用章；Excel 不蓋章）。`, q.id);
    } else {
      const next = ap.steps[ap.cur];
      notify(ctx, ctx.me, [q.owner], 'quote_step_approved', `☑️ 報價單通過「${step.label}」：${q.quoteNo}`, `${noteLine(q, ctx)}　下一關：${next.label}`, q.id);
      // 董事會關會通知「所有」秘書（含非本部門）：內文只放單號，不帶客戶名與業務名
      notify(ctx, ctx.me, stepRecipients(next, ctx), next.tier === 'board' ? 'quote_board' : 'quote_submitted',
        next.tier === 'board' ? `🏛️ 報價單待董事會決議登錄：${q.quoteNo}` : `📝 報價單待簽核：${q.quoteNo}`,
        next.tier === 'board' ? `${q.quoteNo}　已完成前面各關簽核，待董事會決議後由管理部秘書代核准。` : `${noteLine(q, ctx)}　目前關卡：${next.label}`, q.id);
    }
    res.json(serialize(q, ctx));
  });

  app.post('/api/quotations/:id/return', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const comment = sanitizeStr((req.body || {}).comment, 500);
    if (!comment) return fail(res, 400, 'COMMENT_REQUIRED', '駁回時請填寫原因');
    const g = actionGuard(req, res, ctx);
    if (!g) return;
    const { q, ap, step } = g;
    const now = nowIso();
    step.status = 'returned'; step.by = ctx.me; step.at = now; step.comment = comment;
    ap.state = 'returned';
    pushHistory(ap, ctx.me, 'RETURN', comment, step.tier, auditMeta(q, ap));
    q.updatedAt = now;
    db.save(ctx.data);
    writeLog('RETURN_QUOTATION', ctx.me, q.quoteNo, `${step.label}駁回：${comment}`, req);
    notify(ctx, ctx.me, [q.owner], 'quote_returned', `↩️ 報價單被駁回：${q.quoteNo}`, `${noteLine(q, ctx)}　${step.label}駁回：${comment}`, q.id);
    res.json(serialize(q, ctx));
  });

  app.post('/api/quotations/:id/reassign', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    if (ctx.role !== 'admin') return fail(res, 403, 'NO_PERMISSION', '只有管理員可以改派');
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const ap = q.approval;
    if (!ap || ap.state !== 'pending' || ap.cur !== 0 || !ap.steps[0] || ap.steps[0].tier !== 'mgr1') {
      return fail(res, 409, 'CANNOT_REASSIGN', '只有「一級主管關尚未簽核」的報價單可以改派');
    }
    const un = String((req.body || {}).username || '');
    const u = ctx.users[un];
    if (!u || !canAct(u) || u.role !== 'manager1') return fail(res, 400, 'BAD_ASSIGNEE', '改派對象必須是在職且非唯讀的一級主管');
    if (un === q.owner || un === ap.submittedBy || un === q.costBy) return fail(res, 400, 'BAD_ASSIGNEE', '不能改派給報價單擁有者或成本填寫人');
    const from = ap.steps[0].assignee;
    ap.steps[0].assignee = un;
    pushHistory(ap, ctx.me, 'REASSIGN', `${dispName(ctx.users, from)} → ${dispName(ctx.users, un)}`, 'mgr1');
    q.updatedAt = nowIso();
    db.save(ctx.data);
    writeLog('REASSIGN_QUOTATION', ctx.me, q.quoteNo, `一級主管關改派：${from} → ${un}`, req);
    notify(ctx, ctx.me, [un], 'quote_submitted', `📝 報價單待簽核：${q.quoteNo}`, `${noteLine(q, ctx)}　管理員改派給你，目前關卡：一級主管`, q.id);
    res.json(serialize(q, ctx));
  });

  // ═════════════════════════════════════════════════════════
  // 匯出／預覽資訊
  // ═════════════════════════════════════════════════════════
  app.get('/api/quotations/:id/export', qAuth, async (req, res) => {
    const ctx = loadCtx(req);
    const live = findQuote(ctx, req.params.id);
    if (!live) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(live, ctx);
    if (!rel.canView) return fail(res, 403, 'NO_PERMISSION', '無權限');
    if (rel.consultantOnly) return fail(res, 403, 'NO_PERMISSION', '成本填寫人不能匯出報價單');
    if (!fs.existsSync(QUOTE_TEMPLATE)) return fail(res, 500, 'NO_TEMPLATE', '報價單範本不存在，請聯繫管理員');
    // 先拍一份快照：buildQuoteWorkbook 是 async（載入範本有多個 await），Postgres 模式下 live 物件是共用快取，
    // 匯出途中若別的請求就地改了內容（例如業務確認作廢後改價），「核准有效」的判斷與實際印出的內容就會不一致。
    // valid 判斷與檔案內容都只用快照。
    const q = JSON.parse(JSON.stringify(live));
    const ap = q.approval;
    const valid = !!ap && ap.state === 'approved' && QA.contentHash(q) === ap.hash;
    try {
      // 規則（Steven 2026-10-07）：Excel 一律不蓋報價專用章——章只出現在 PDF（與網頁預覽）。
      // Excel 是可編輯檔，若帶章，業務調整後轉 PDF 就成了「有章但沒有走簽核」的檔案；沒有章就一眼看得出不是正式版。
      // 數量／金額等一經調整，核准即作廢（contentHash）需重新簽核，所以有章的正式版只有「核准有效當下」的 PDF。
      const buf = await buildQuoteWorkbook(q, QUOTE_TEMPLATE, { issueDate: taipeiToday(), issuer: resolveIssuer(q.owner), approved: valid, seal: null, report: {} });
      const fname = encodeURIComponent(`${q.quoteNo}_${q.company || '報價單'}.xlsx`);
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fname}`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('X-Quote-Approval', valid ? 'approved' : 'unapproved');
      res.setHeader('X-Quote-Seal', 'none');   // 保留標頭給舊前端：Excel 一律沒有章
      writeLog('EXPORT_QUOTATION', ctx.me, q.quoteNo, valid ? '匯出 Excel（已核准；Excel 不蓋報價章，正式版請下載 PDF）' : '匯出 Excel（未核准）', req);
      await flushAuditWrites();
      res.send(buf);
    } catch (e) {
      console.error('[QuoteExport]', e.message, e.stack);
      fail(res, 500, 'EXPORT_FAILED', '報價單產生失敗，請稍後再試；若持續發生請聯絡管理員');
    }
  });

  // PDF 由瀏覽器端依「預覽畫面」產生（伺服器沒有中文字型），所以下載本身伺服器看不到；
  // 這支只負責記稽核紀錄（誰、何時、下載的是否為已核准有章的版本），並用伺服器端的狀態回報蓋章與否。
  app.post('/api/quotations/:id/export-log', qAuth, async (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(q, ctx);
    if (!rel.canView || rel.consultantOnly || !permOf(rel, q, ctx).canSeePrice) return fail(res, 403, 'NO_PERMISSION', '無權限');
    if (!req.body || req.body.format !== 'pdf') return fail(res, 400, 'BAD_FORMAT', '不支援的格式');
    const ap = q.approval;
    const valid = !!ap && ap.state === 'approved' && QA.contentHash(q) === ap.hash;
    const sealed = valid && sealUploaded(ctx);
    writeLog('EXPORT_QUOTATION_PDF', ctx.me, q.quoteNo, sealed ? '下載 PDF（已核准，含報價章）' : valid ? '下載 PDF（已核准，但尚未上傳報價專用章）' : '下載 PDF（未核准，無報價章）', req);
    await flushAuditWrites();
    res.json({ success: true, approval: valid ? 'approved' : 'unapproved', seal: sealed ? 'applied' : valid ? 'missing' : 'none' });
  });

  app.get('/api/quotations/:id/issue-info', qAuth, (req, res) => {
    const ctx = loadCtx(req);
    const q = findQuote(ctx, req.params.id);
    if (!q) return fail(res, 404, 'NOT_FOUND', '找不到此報價單');
    const rel = relationOf(q, ctx);
    if (!rel.canView || rel.consultantOnly) return fail(res, 403, 'NO_PERMISSION', '無權限');
    // remarks：與 Excel 完全同一份條款（網頁預覽直接用，不在前端另存一份條文）
    res.json({ issueDate: taipeiToday(), issuer: resolveIssuer(q.owner), ownerIsMe: q.owner === ctx.me, remarks: REM.composeRemarks(q, quoteExcel.effectiveValidUntil(q)) });
  });

  /**
   * 產生毛利分析 xlsx（下載與預覽共用，所以預覽＝下載檔的樣子）。
   * 品項沒有明確分類時，依報價單勾選商品的類別推測（單一類別就全歸該類）；申請人＝報價單擁有者（吃暱稱，內部文件）。
   * 先拍快照：產檔是 async，Postgres 模式下 q 是共用快取，途中被別的請求就地改掉會讓檔案混入新舊內容。
   */
  async function buildPnlBuffer(q, ctx) {
    const pcs = (ctx.cfg && ctx.cfg.productClasses) || {};
    const classCodes = [...new Set((q.products || []).map(p => pcs[p] && pcs[p].cls).filter(Boolean))];
    return buildQuotePnlExcel(JSON.parse(JSON.stringify(q)), { classCodes, requestedBy: dispName(ctx.users, q.owner), issueDate: taipeiToday(), contingencyPct: q.contingencyPct });
  }
  /** 毛利分析（下載／預覽）的存取檢查：必須同時看得到成本與價格（只負責填成本的顧問看不到價格，不可看）。回傳 {q} 或 null（已回應錯誤） */
  function pnlAccess(req, res, ctx) {
    const q = findQuote(ctx, req.params.id);
    if (!q) { fail(res, 404, 'NOT_FOUND', '找不到此報價單'); return null; }
    const rel = relationOf(q, ctx);
    const pm = permOf(rel, q, ctx);
    if (!rel.canView || !pm.canSeeCost || !pm.canSeePrice) { fail(res, 403, 'NO_PERMISSION', '你沒有查看這張報價單毛利分析的權限'); return null; }
    return { q };
  }

  // 毛利分析預覽（內部）：與下載同一份 xlsx 轉成 HTML；權限與下載相同，並記稽核（含成本與毛利）
  const pnlViewSeen = new Map();   // 稽核去重：'帳號|單 id' → 上次記錄時間（行程內；多實例各自去重即可）
  app.get('/api/quotations/:id/pnl-preview', qAuth, async (req, res) => {
    const ctx = loadCtx(req);
    const acc = pnlAccess(req, res, ctx);
    if (!acc) return;
    try {
      const buf = await buildPnlBuffer(acc.q, ctx);
      const view = await sheetToHtml(buf);
      const seenKey = ctx.me + '|' + acc.q.id, nowMs = Date.now();
      if (!(pnlViewSeen.get(seenKey) > nowMs - 60000)) {
        pnlViewSeen.set(seenKey, nowMs);
        if (pnlViewSeen.size > 500) for (const [k, ts] of pnlViewSeen) if (ts < nowMs - 60000) pnlViewSeen.delete(k);
        writeLog('VIEW_QUOTE_PNL', ctx.me, acc.q.quoteNo, '預覽毛利分析（含成本）', req);
        await flushAuditWrites();
      }
      res.setHeader('Cache-Control', 'no-store');
      res.json({ quoteNo: acc.q.quoteNo, html: view.html, widthPx: view.widthPx, heightPx: view.heightPx });
    } catch (e) {
      console.error('[QuotePnlPreview]', e.message, e.stack);
      fail(res, 500, 'PREVIEW_FAILED', '毛利分析預覽產生失敗，請稍後再試；若持續發生請聯絡管理員');
    }
  });

  // 毛利分析（內部）：含成本與毛利率，只有「看得到成本」的人可以下載
  app.get('/api/quotations/:id/export-pnl', qAuth, async (req, res) => {
    const ctx = loadCtx(req);
    const acc = pnlAccess(req, res, ctx);
    if (!acc) return;
    const q = acc.q;
    try {
      const buf = await buildPnlBuffer(q, ctx);
      const fname = encodeURIComponent(`${q.quoteNo}_毛利分析-內部.xlsx`);
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fname}`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      writeLog('EXPORT_QUOTE_PNL', ctx.me, q.quoteNo, '匯出毛利分析（含成本）', req);
      await flushAuditWrites();
      res.send(buf);
    } catch (e) {
      console.error('[QuotePnlExport]', e.message, e.stack);
      fail(res, 500, 'EXPORT_FAILED', '毛利分析產生失敗，請稍後再試；若持續發生請聯絡管理員');
    }
  });

  // 供測試與除錯使用的內部函式
  return { _internal: { getCfg, boardProxySet, sealManagerSet, findManager1, relationOf, permOf, buildPreview, serialize, costComplete } };
};
