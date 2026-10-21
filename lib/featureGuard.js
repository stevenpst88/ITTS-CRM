'use strict';
/**
 * 後端功能權限驗證（FEATURE_GUARD_MODE）——內控稽核第三波第二批 ①。
 *
 * 角色功能權限矩陣（server.js 的 DEFAULT_ROLE_PERMISSIONS／getRolePermissions／FEATURE_REGISTRY）原本只用來隱藏前端導覽；
 * 資料 API 多半只有 requireAuth＋getViewableOwners（可視範圍），所以「沒勾某功能」的角色仍可直接打 API。
 * 本模組以一張「路由 → 功能」宣告式對照表，在後端驗證呼叫者的角色是否擁有該路由所服務的功能之一（ANY-of：
 * 同一支 API 常被多個畫面共用，例如商機清單給看板、預測表、報價單共用）。
 *
 * 模式（環境變數 FEATURE_GUARD_MODE）：
 *   off      不介入
 *   log      預設。永不阻擋；若「在 enforce 下會被擋」就累計計數並寫稽核 FEATURE_GUARD_WOULD_DENY（每人每路由 10 分鐘最多一筆），
 *            供業務試行一段時間、確認沒有誤判後再切 enforce
 *   enforce  回 403 { code:'FEATURE_DENIED', feature:[…] }，並寫稽核 FEATURE_GUARD_DENY（同樣節流）
 * 回復：把 FEATURE_GUARD_MODE 設成 off 或 log（重新部署／重啟即生效），不需改資料。
 *
 * 路由種類（kind）：
 *   FEATURES  需擁有 features 其中之一（ANY-of）
 *   ANY       需擁有至少一個功能（全域載入用的唯讀清單，例如聯絡人／商機清單，登入後每個畫面都會預載）
 *   ALWAYS    登入後一律放行（登入／登出／我的資料／通知輪詢／公告…）
 *   SELF      由路由本身的角色／名冊／關係檢查把關（例如報價單簽核人、成本填寫人不一定有「報價單」功能；處理器內有 role 檢查的端點）
 *   ADMIN     requireAdmin 路由（本守門不重複檢查；清單供稽核參考）
 *   UNCLASSIFIED  無法有把握歸類：視同放行，但每個路由只記錄一次（FEATURE_GUARD_UNCLASSIFIED）
 * 管理員（role=admin）在 getUserFeatures 中擁有全部功能，所以以一般使用者身分用業務 API 時不受影響；管理後台 API 本來就由 requireAdmin 把關。
 *
 * 路由表必須與實際註冊的 /api 路由同步：scripts/check-feature-guard.js 會解析原始碼，新增了路由卻沒有對照表項目時直接失敗。
 */

const CARD = ['contacts', 'prospects'];                       // 名片／聯絡人卡片（新增、編輯、檢視）只從這兩個清單開啟
const OPPEDIT = ['pipeline', 'forecast', 'quotations'];       // 共用的「商機編輯」視窗只從看板、預測表、報價單開啟
const OPPCREATE = ['keyAccount', 'quotations'];

// [method(s), express 路徑, 規格, 備註, 信心]  規格：字串＝kind（ALWAYS/ANY/SELF/ADMIN/UNCLASSIFIED），陣列＝FEATURES
const R = [];
const add = (methods, p, spec, note, conf) => { methods.split(',').forEach(m => R.push({ method: m.trim().toUpperCase(), path: p, spec, note: note || '', conf: conf || 'high' })); };

// ── 帳號／通知／全域小工具：登入後一律放行 ──
add('GET', '/api/push/vapid-public-key', 'ALWAYS', '公開（requireAuth 之前）');
add('POST', '/api/login', 'ALWAYS', '登入');
add('POST', '/api/logout', 'ALWAYS');
add('GET', '/api/me', 'ALWAYS');
add('GET', '/api/session/touch', 'ALWAYS', '閒置登出的前端心跳（PR #21）；與 /api/me 同為登入後一律放行');
add('GET', '/api/me/permissions', 'ALWAYS', '前端用來決定導覽；本身回傳 features');
add('POST', '/api/push/subscribe', 'ALWAYS');
add('POST', '/api/push/unsubscribe', 'ALWAYS');
add('POST', '/api/push/test', 'ALWAYS');
add('PUT', '/api/user/password', 'ALWAYS', '改自己的密碼');
add('GET', '/api/plan', 'ALWAYS', '訂閱方案狀態（任何登入者可查）');
add('GET', '/api/notifications', 'ALWAYS', '站內通知（鈴鐺；前端已改用 poll-bundle）');
add('GET', '/api/poll-bundle', 'ALWAYS', '通知＋合約提醒＋生日提醒的合併輪詢');
add('PUT', '/api/notifications/read-all', 'ALWAYS');
add('PUT', '/api/notifications/:id/read', 'ALWAYS');
add('GET', '/api/announcements/active', 'ALWAYS', '公告彈窗');
add('POST', '/api/announcements/:id/ack', 'ALWAYS', '公告「我知道了」');
add('GET', '/api/cron/mail-outbox', 'ALWAYS', 'Vercel Cron，以 cron 密鑰驗證，無 session');
add('GET', '/api/cron/audit-anchor', 'ALWAYS', 'Vercel Cron（稽核鏈每日錨點），以 cron 密鑰驗證，無 session；AUDIT_CHAIN 未啟用時直接回 disabled');
add('POST', '/api/csp-report', 'ALWAYS', '公開（requireAuth 之前，登入頁也會回報）：CSP Report-Only 違規回報，自帶 IP 限流／大小上限／清洗（lib/cspReport.js）；CSP_REPORT_ONLY=off 時不註冊');

// ── 全域唯讀清單：登入即預載，改用「至少一個功能」 ──
add('GET', '/api/contacts', 'ANY', '登入後預載（loadContacts）；報價單精靈、商機表單共用', 'medium');
add('GET', '/api/visits', 'ANY', '隨聯絡人預載；各畫面共用', 'medium');
add('GET', '/api/opportunities', 'ANY', '隨聯絡人預載；看板、預測、報價單、儀表板共用', 'medium');
add('GET', '/api/key-accounts', 'ANY', '隨聯絡人預載（handler 內另對部分角色回空陣列）', 'medium');
add('GET', '/api/product-catalog', 'ANY', '商機商品下拉（全域預載）');
add('GET', '/api/industries', 'ANY', '產業選單（全域預載）');
add('GET', '/api/usermap', 'ANY', '業務姓名對應（多個畫面共用）');
// 儀表板（showDashboard）在任何角色重新整理或剛登入時都會載入，並不以「首頁」功能為前提
add('GET', '/api/birthday-reminders', 'ANY', '儀表板載入', 'medium');
add('GET', '/api/targets', 'ANY', '儀表板載入＋業績目標頁', 'medium');
add('GET', '/api/settings/quarter-ratios', 'ANY', '儀表板載入', 'medium');
add('GET', '/api/settings/user-quarter-ratios', 'ANY', '儀表板載入＋業績目標頁', 'medium');
add('GET', '/api/settings/opr-range', 'ANY', '儀表板載入', 'medium');
add('GET', '/api/monthly-budget', 'ANY', '儀表板載入＋業績目標／預測頁', 'medium');
add('GET', '/api/pure-supervisors', 'ANY', '儀表板載入', 'medium');
add('GET', '/api/zombie-opportunities', 'ANY', '儀表板載入＋業績目標頁', 'medium');
add('GET', '/api/users-bu', 'ANY', '儀表板載入＋業績目標／預測頁', 'medium');

// ── AI ──
add('POST', '/api/ai/ocr-card', CARD);
add('POST', '/api/ai/contact-summary', CARD);
add('POST', '/api/ai/company-insight', CARD);
add('POST', '/api/ai/visit-suggest', ['visits']);
add('POST', '/api/ai/follow-up-email', ['visits']);
add('POST', '/api/ai/opp-win-rate', OPPEDIT);

// ── 聯絡人／潛在客戶／名片 ──
add('POST', '/api/key-accounts', [...CARD, 'keyAccount'], '每張名片卡片上的星號按鈕＋重點客戶頁', 'medium');
add('DELETE', '/api/key-accounts/:id', [...CARD, 'keyAccount'], '', 'medium');
add('POST', '/api/contacts', CARD);
add('PUT', '/api/contacts/:id', [...CARD, 'pipeline', 'forecast', 'quotations', 'targets'], '含「成交商機 → 潛在客戶轉客戶」流程（從看板／商機視窗／業績目標頁觸發）', 'medium');
add('DELETE', '/api/contacts/:id', CARD);
add('GET', '/api/groups', CARD);
add('POST', '/api/groups', CARD);
add('PUT', '/api/groups/:id', CARD);
add('DELETE', '/api/groups/:id', CARD);
add('POST', '/api/upload', CARD, '名片圖片上傳');
add('GET', '/api/export', ['contacts'], '匯出聯絡人（handler 另檢查 canDownloadContacts）');
add('GET', '/api/company-master', [...CARD, 'companyMaster'], '統編欄位比對', 'medium');
add('GET', '/api/company-similar', [...CARD, 'companyMaster'], '', 'medium');
add('GET', '/api/company-lookup', [...CARD, 'companyMaster'], '統編查詢（名片表單、企業主檔）', 'medium');

// ── 拜訪 ──
add('POST', '/api/visits', ['visits'], '拜訪表單只從拜訪頁開啟（未驗證是否也從聯絡人卡開啟）', 'medium');
add('PUT', '/api/visits/:id', ['visits']);
add('DELETE', '/api/visits/:id', ['visits']);

// ── 業績目標／預算 ──
add('GET', '/api/manager/achievement', ['targets'], 'handler 另限 manager1/manager2/admin/executive/secretary');
add('PUT', '/api/manager/target/:username', ['targets']);
add('POST', '/api/targets', ['targets'], 'handler 內檢查 canSetTargets');
add('PUT', '/api/settings/user-quarter-ratios', 'SELF', 'handler 限 admin／secretary；前端只在後台頁使用');
add('PUT', '/api/monthly-budget', ['targets']);
add('PUT', '/api/monthly-budget/actuals', ['home', 'targets'], '儀表板的月預算卡片編輯鈕；handler 限 admin/secretary/executive', 'medium');

// ── 商機 ──
add('POST', '/api/opportunities', [...CARD, 'visits', ...OPPCREATE], '名片表單內建立、拜訪表單、重點客戶頁／報價單建立商機', 'medium');
add('POST', '/api/opportunities/:id/ka-confirm', OPPEDIT);
add('PUT', '/api/opportunities/:id', [...OPPEDIT, 'targets'], '看板拖曳、商機視窗、業績目標頁改階段');
add('DELETE', '/api/opportunities/:id', OPPEDIT);
add('GET', '/api/pipeline-date-changes', ['pipelineReport']);
add('GET', '/api/pipeline-report', ['pipelineReport']);
add('GET', '/api/lost-opportunities', ['lostOpp', 'targets']);
add('POST', '/api/opportunities/restore/:id', ['targets', 'lostOpp'], '前端只在業績目標頁的商機表出現還原鈕；流失商機頁一併放行', 'medium');
add('GET', '/api/forecast/export', ['forecast']);

// ── 合約／帳款 ──
add('GET', '/api/contracts', ['contractGroup']);
add('POST', '/api/contracts', ['contractGroup']);
add('PUT', '/api/contracts/:id', ['contractGroup']);
add('DELETE', '/api/contracts/:id', ['contractGroup']);
add('GET', '/api/receivables', ['accountingGroup']);
add('POST', '/api/receivables', ['accountingGroup']);
add('PUT', '/api/receivables/:id', ['accountingGroup']);
add('DELETE', '/api/receivables/:id', ['accountingGroup']);

// ── Call-in／行銷／Lead ──
add('GET', '/api/callins', ['callin']);
add('POST', '/api/callins', ['callin']);
add('PUT', '/api/callins/:id/assign', ['callin']);
add('PUT', '/api/callins/:id/respond', ['callin']);
add('GET', '/api/campaigns', ['campaigns', 'leads'], 'Lead 頁的活動下拉共用');
add('POST', '/api/campaigns', ['campaigns']);
add('PUT', '/api/campaigns/:id', ['campaigns']);
add('DELETE', '/api/campaigns/:id', ['campaigns'], '刪除鈕的前端接線未逐一確認', 'medium');
add('GET', '/api/leads', ['leads']);
add('POST', '/api/leads', ['leads']);
add('PUT', '/api/leads/:id', ['leads']);
add('POST', '/api/leads/:id/assign', ['leads']);
add('POST', '/api/leads/:id/convert', ['leads']);
add('POST', '/api/leads/:id/disqualify', ['leads']);

// ── 名單移轉 ──
add('GET', '/api/contacts-by-owner', ['transfer']);
add('POST', '/api/transfer-contacts', ['transfer']);

// ── 企業主檔 ──
add('GET', '/api/companies', ['companyMaster']);
add('GET', '/api/companies/export', ['companyMaster']);
add('POST', '/api/companies/import', ['companyMaster']);
add('GET', '/api/companies/import-template', ['companyMaster']);
add('POST', '/api/industries', ['companyMaster']);
add('DELETE', '/api/industries/:name', ['companyMaster']);
add('PUT', '/api/industries', ['companyMaster'], '前端只在後台頁使用；handler 限 admin／marketing');
add('POST', '/api/industries/cleanup', ['companyMaster'], '同上');
add('POST', '/api/companies/:id/set-industry', ['companyMaster']);
add('POST', '/api/companies/:id/set-region', ['companyMaster'], '前端只在後台頁使用；handler 限 admin／marketing');
add('POST', '/api/companies/:id/set-customer-code', ['companyMaster'], '同上');
add('POST', '/api/companies/:id/set-responsible-sales', ['companyMaster'], '企業主檔責任業務；handler 限行銷／管理員／主管且各自收斂到可視範圍；COMPANY_CARD_TASKS=off 時 404');
add('POST', '/api/companies/responsible-sales-batch', ['companyMaster'], '批次指定責任業務；同上');
add('GET', '/api/company-responsible/candidates', ['companyMaster'], '責任業務選單來源；handler 限行銷／管理員／主管');
add('DELETE', '/api/companies/:id', ['companyMaster']);
add('GET', '/api/companies/:id', ['companyMaster']);
add('GET', '/api/companies/:id/yoy', ['companyMaster', 'keyAccount', 'yoy'], 'handler 已另行檢查 yoy 功能（無權限時前端靜默忽略）');

// ── YoY ──
add('POST', '/api/yoy/import', ['yoy']);
add('GET', '/api/yoy/report', ['yoy']);
add('GET', '/api/yoy/links', ['yoy']);
add('POST', '/api/yoy/links', ['yoy']);
add('DELETE', '/api/yoy/:year', 'UNCLASSIFIED', '前端與後台都沒有呼叫者；handler 限 canImportYoy（admin／secretary）');

// ── 主管首頁／分析／高階儀表板 ──
add('GET', '/api/manager-home', ['managerHome']);
add('GET', '/api/manager/business-type-analysis', ['bizAnalysis']);
add('GET', '/api/manager/company-card-overdue', ['managerHome'], '主管首頁：責任業務逾期未補名片＋業務回報；handler 限 admin／executive／manager1／manager2 並收斂到可視範圍');

// ── 待補名片待辦（登入後每個角色都會查自己的；行銷的回報清單由 handler 限 marketing／admin）──
add('GET', '/api/my-tasks/company-cards', 'ALWAYS', '登入彙總視窗與首頁待辦卡片；只回呼叫者自己的待辦（COMPANY_CARD_TASKS=off 時回 enabled:false）');
add('POST', '/api/my-tasks/company-cards/dismiss', 'ALWAYS', '只能解除自己的待辦，原因必選');
add('GET', '/api/marketing/responsibility-feedback', 'SELF', '行銷首頁「業務回報待修」；handler 限 marketing／admin');
add('POST', '/api/marketing/responsibility-feedback/:id/resolve', 'SELF', '同上');
add('GET', '/api/exec/conversion', ['execDash']);
add('GET', '/api/exec/trend', ['execDash']);
add('GET', '/api/exec/product-analysis', ['execDash']);
add('GET', '/api/exec/bu-comparison', ['execDash']);

// ── 後台頁使用、但以 requireAuth＋handler 角色檢查把關的路由（SELF）──
add('GET', '/api/admin/pool/contacts', 'SELF', '客戶池；前端只在後台頁使用；handler 限 admin/manager1/manager2/executive');
add('POST', '/api/admin/pool/contacts/:id/delete', 'SELF', '同上');
add('POST', '/api/admin/pool/assign', 'SELF', '同上');

// ── 報價單（lib/quoteRoutes.js 自有關係／名冊／featureOk 把關；簽核人與成本填寫人不一定擁有「報價單」功能，不重複檢查）──
add('GET', '/api/me/contact', ['quotations'], '報價單廠商資料（我的聯絡資訊）');
add('PUT', '/api/me/contact', ['quotations']);
const Q = 'lib/quoteRoutes.js 自行檢查（featureOk／關係／名冊；簽核人、成本填寫人不一定有「報價單」功能）';
add('GET', '/api/quote-approval/config', 'SELF', Q);
add('GET', '/api/quote-pricebook', 'SELF', Q);
add('PUT', '/api/quote-approval/seal', 'SELF', Q);
add('DELETE', '/api/quote-approval/seal', 'SELF', Q);
add('GET', '/api/quote-approval/seal', 'SELF', Q);
add('GET', '/api/quotations', 'SELF', Q);
add('GET', '/api/quotations/inbox', 'SELF', Q);
add('GET', '/api/opportunities/:id/quotes', 'SELF', Q);
add('GET', '/api/quotations/:id', 'SELF', Q);
add('GET', '/api/quotations/:id/seal', 'SELF', Q);
add('POST', '/api/quotations', 'SELF', Q);
add('PUT', '/api/quotations/:id', 'SELF', Q);
add('PUT', '/api/quotations/:id/item-notes', 'SELF', Q);
add('DELETE', '/api/quotations/:id', 'SELF', Q);
add('PUT', '/api/quotations/:id/costs', 'SELF', Q);
add('POST', '/api/quotations/:id/cost-draft/summary', 'SELF', Q);
add('POST', '/api/quotations/:id/cost-draft/pnl-preview', 'SELF', Q);
add('POST', '/api/quotations/:id/submit', 'SELF', Q);
add('POST', '/api/quotations/:id/withdraw', 'SELF', Q);
add('POST', '/api/quotations/:id/approve', 'SELF', Q);
add('POST', '/api/quotations/:id/return', 'SELF', Q);
add('POST', '/api/quotations/:id/reassign', 'SELF', Q);
add('GET', '/api/quotations/:id/export', 'SELF', Q);
add('POST', '/api/quotations/:id/export-log', 'SELF', Q);
add('GET', '/api/quotations/:id/issue-info', 'SELF', Q);
add('GET', '/api/quotations/:id/history-check', 'SELF', Q);
add('POST', '/api/quotations/:id/issue', 'SELF', Q);
add('GET', '/api/quotations/:id/issuances', 'SELF', Q);
add('GET', '/api/quotations/:id/pnl-preview', 'SELF', Q);
add('GET', '/api/quotations/:id/export-pnl', 'SELF', Q);
add('GET', '/api/contract-reminders', 'UNCLASSIFIED', '前端沒有直接呼叫（已併入 poll-bundle，所有人的鈴鐺共用）；是否改為需要「合約管理」功能待業務決定');

// 管理員專用但路徑不在 /api/admin/ 底下
add('PUT', '/api/settings/quarter-ratios', 'ADMIN', 'requireAdmin');
add('PUT', '/api/settings/opr-range', 'ADMIN', 'requireAdmin');

const ROUTES = R;
const keyOf = (m, p) => `${m} ${p}`;

// ── 比對器：把 :param 編成一段非斜線；靜態段較多者優先（與 Express 先註冊先命中等價，因為沒有「通用路由先於特例」的情況）──
function compile(entry) {
  const segs = entry.path.split('/').filter(Boolean);
  const re = new RegExp('^/' + segs.map(s => s.startsWith(':') ? '[^/]+' : s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('/') + '/?$', 'i');
  return { entry, re, statics: segs.filter(s => !s.startsWith(':')).length, params: segs.filter(s => s.startsWith(':')).length, len: segs.length };
}
const COMPILED = ROUTES.map(compile).sort((a, b) => (b.statics - a.statics) || (a.params - b.params));
function matchRoute(method, path) {
  const m = String(method || 'GET').toUpperCase() === 'HEAD' ? 'GET' : String(method || 'GET').toUpperCase();
  for (const c of COMPILED) if (c.entry.method === m && c.re.test(path)) return c.entry;
  return null;
}

const MODES = ['off', 'log', 'enforce'];
function readMode(env) {
  const v = String((env || process.env).FEATURE_GUARD_MODE || '').trim().toLowerCase();
  return MODES.includes(v) ? v : 'log';
}
const kindOf = (e) => Array.isArray(e.spec) ? 'FEATURES' : e.spec;

/**
 * 建立守門。
 * opts: { mode, getFeatures(role)→string[], audit(req, action, operator, target, detail), clock, throttleMs, env }
 */
function create(opts) {
  opts = opts || {};
  const mode = opts.mode || readMode(opts.env);
  const clock = typeof opts.clock === 'function' ? opts.clock : Date.now;
  const throttleMs = opts.throttleMs || 10 * 60 * 1000;
  const getFeatures = opts.getFeatures || (() => []);
  const audit = opts.audit || (() => {});
  const stats = { mode, wouldDeny: 0, denied: 0, unclassified: 0 };
  const seen = new Map();            // 'action|user|routeKey' → t（節流）
  const unclassifiedSeen = new Set();

  function throttled(k, t) {
    const last = seen.get(k);
    if (last !== undefined && t - last < throttleMs) return true;
    seen.set(k, t);
    if (seen.size > 5000) for (const [kk, tt] of seen) if (t - tt >= throttleMs) seen.delete(kk);
    return false;
  }

  /** 純判斷（不寫稽核）：回傳 { allow, kind, features, entry } */
  function decide(method, path, role) {
    const entry = matchRoute(method, path);
    if (!entry) {
      if (/^\/api\/admin\//i.test(path)) return { allow: true, kind: 'ADMIN', entry: null };
      return { allow: true, kind: 'NOMAP', entry: null };
    }
    const kind = kindOf(entry);
    if (kind !== 'FEATURES' && kind !== 'ANY') return { allow: true, kind, entry };
    const feats = getFeatures(role) || [];
    if (kind === 'ANY') return { allow: feats.length > 0, kind, features: ['(任一功能)'], entry };
    return { allow: entry.spec.some(f => feats.includes(f)), kind, features: entry.spec, entry };
  }

  function middleware(req, res, next) {
    if (mode === 'off') return next();
    // app.use('/api/', …) 掛載時 req.path 會被去掉 /api 前綴，要用 baseUrl 還原完整路徑
    const p = (req.baseUrl || '') + (req.path || '');
    if (!p.startsWith('/api/') || req.method === 'OPTIONS') return next();
    const su = req.session && req.session.user;
    if (!su || !su.username) return next();
    let d;
    try { d = decide(req.method, p, su.role || 'user'); } catch (e) { console.error('[featureGuard]', e && e.message); return next(); }
    if (d.kind === 'NOMAP') {
      const key = keyOf(String(req.method).toUpperCase(), p.toLowerCase().replace(/\/[0-9a-f-]{8,}(?=\/|$)/g, '/:id').slice(0, 120));
      if (!unclassifiedSeen.has(key) && unclassifiedSeen.size < 300) { unclassifiedSeen.add(key); stats.unclassified++; try { audit(req, 'FEATURE_GUARD_UNCLASSIFIED', su.username, key, `對照表沒有此路由（${key}），已放行；請補上路由→功能對照`); } catch (e) { /* ignore */ } }
      return next();
    }
    if (d.kind === 'UNCLASSIFIED') {
      const key = keyOf(d.entry.method, d.entry.path);
      if (!unclassifiedSeen.has(key)) { unclassifiedSeen.add(key); stats.unclassified++; try { audit(req, 'FEATURE_GUARD_UNCLASSIFIED', su.username, key, `此路由尚未歸類功能（${d.entry.note || ''}），視同放行`); } catch (e) { /* ignore */ } }
      return next();
    }
    if (d.allow) return next();
    const key = keyOf(d.entry.method, d.entry.path);
    const t = clock();
    const detail = `角色 ${su.role || 'user'} 沒有「${d.features.join('／')}」任一功能，呼叫 ${req.method} ${p.slice(0, 100)}`;
    if (mode === 'log') {
      stats.wouldDeny++;
      if (!throttled('W|' + su.username + '|' + key, t)) { try { audit(req, 'FEATURE_GUARD_WOULD_DENY', su.username, key, detail + '（log 模式：未阻擋）'); } catch (e) { console.error('[featureGuard audit]', e && e.message); } }
      return next();
    }
    stats.denied++;
    if (!throttled('D|' + su.username + '|' + key, t)) { try { audit(req, 'FEATURE_GUARD_DENY', su.username, key, detail + '（已阻擋 403）'); } catch (e) { console.error('[featureGuard audit]', e && e.message); } }
    return res.status(403).json({ error: '您的角色沒有此功能的權限', code: 'FEATURE_DENIED', feature: d.features });
  }

  return { mode, stats, decide, middleware, snapshot: () => ({ mode, wouldDeny: stats.wouldDeny, denied: stats.denied, unclassified: stats.unclassified }), unclassifiedSeen };
}

/** 給稽核看的對照表（含功能中文名） */
function describe(featureRegistry) {
  const label = Object.fromEntries((featureRegistry || []).map(f => [f.key, f.label]));
  return ROUTES.map(e => ({
    method: e.method, path: e.path, kind: kindOf(e),
    features: Array.isArray(e.spec) ? e.spec : [],
    featureLabels: Array.isArray(e.spec) ? e.spec.map(k => label[k] || k) : [],
    note: e.note, confidence: e.conf,
  }));
}

/** 讀 Express 已註冊的路由（Express 4：app._router.stack）；用來找出「已註冊卻不在對照表」的 /api 路由 */
function listRegisteredRoutes(app) {
  const out = [];
  const stack = (app._router && app._router.stack) || [];
  for (const layer of stack) {
    if (!layer.route) continue;
    const paths = Array.isArray(layer.route.path) ? layer.route.path : [layer.route.path];
    const names = (layer.route.stack || []).map(h => h.name);
    const admin = names.includes('requireAdmin');
    for (const p of paths) {
      if (typeof p !== 'string' || !p.startsWith('/api/')) continue;
      for (const m of Object.keys(layer.route.methods || {})) out.push({ method: m.toUpperCase(), path: p, admin });
    }
  }
  return out;
}
function unmappedRegistered(app) {
  const have = new Set(ROUTES.map(e => keyOf(e.method, e.path)));
  return listRegisteredRoutes(app).filter(r => !r.admin && !have.has(keyOf(r.method, r.path)) && !/^\/api\/admin\//.test(r.path));
}

module.exports = { create, readMode, matchRoute, describe, listRegisteredRoutes, unmappedRegistered, ROUTES, keyOf, kindOf, MODES };
