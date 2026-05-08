// Plan-based route access control
// Returns 402 when route requires a higher subscription plan

const PLAN_MODULES = {
  starter:    ['M1', 'M2'],
  growth:     ['M1', 'M2', 'M3', 'M4'],
  pro:        ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7'],
  enterprise: ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'MULTI_BU', 'EXEC_DASH'],
};

// Route prefix → required module
const ROUTE_MODULE_MAP = [
  { prefix: '/api/opportunities',      module: 'M3' },
  { prefix: '/api/kanban',             module: 'M3' },
  { prefix: '/api/targets',            module: 'M4' },
  { prefix: '/api/manager-home',       module: 'M4' },
  { prefix: '/api/forecast',           module: 'M4' },
  { prefix: '/api/contracts',          module: 'M5' },
  { prefix: '/api/receivables',        module: 'M5' },
  { prefix: '/api/campaigns',          module: 'M6' },
  { prefix: '/api/leads',              module: 'M6' },
  { prefix: '/api/callins',            module: 'M6' },
  { prefix: '/api/exec/bu-comparison', module: 'EXEC_DASH' },
];

const MODULE_UPGRADE_HINT = {
  M3: 'Growth',
  M4: 'Growth',
  M5: 'Professional',
  M6: 'Professional',
  M7: 'Professional',
  MULTI_BU: 'Professional',
  EXEC_DASH: 'Professional',
};

function hasModule(sub, moduleName) {
  const plan = sub.plan || 'enterprise';
  const baseModules = PLAN_MODULES[plan] || PLAN_MODULES.enterprise;
  const addons = sub.addons || [];
  return baseModules.includes(moduleName) || addons.includes(moduleName);
}

function isExpired(sub) {
  if (!sub.planExpiry) return false;
  return new Date(sub.planExpiry) < new Date();
}

function isTrialActive(sub) {
  if (!sub.trialUntil) return false;
  return new Date(sub.trialUntil) >= new Date();
}

module.exports = function createPlanGuard(loadAuth) {
  return function planGuard(req, res, next) {
    const auth = loadAuth();
    const sub = auth._subscription;

    // No subscription config = legacy install, full access
    if (!sub) return next();

    // Active trial → treat as Growth plan (full access to M1-M4)
    if (isTrialActive(sub)) return next();

    // Expired plan → read-only access to contacts only
    if (isExpired(sub)) {
      const isContactRead = req.method === 'GET' && req.path.startsWith('/api/contacts');
      if (!isContactRead) {
        return res.status(402).json({
          error: '訂閱已到期，請聯繫管理員續約',
          code: 'PLAN_EXPIRED',
        });
      }
      return next();
    }

    // Check route module requirement
    const routeRule = ROUTE_MODULE_MAP.find(r => req.path.startsWith(r.prefix));
    if (!routeRule) return next();

    if (!hasModule(sub, routeRule.module)) {
      const upgradeTo = MODULE_UPGRADE_HINT[routeRule.module] || 'Professional';
      return res.status(402).json({
        error: `🔒 此功能需要 ${upgradeTo} 方案或以上，請升級訂閱`,
        code: 'PLAN_INSUFFICIENT',
        requiredModule: routeRule.module,
        upgradeTo,
        currentPlan: sub.plan,
      });
    }

    next();
  };
};

module.exports.PLAN_MODULES = PLAN_MODULES;
module.exports.MODULE_UPGRADE_HINT = MODULE_UPGRADE_HINT;
