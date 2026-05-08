// AI Credits deduction middleware
// Deducts credits before each AI API call; returns 402 if balance insufficient

const CREDIT_COSTS = {
  'ocr-card':        5,
  'visit-suggest':   3,
  'opp-win-rate':    8,
  'contact-summary': 5,
  'follow-up-email': 4,
  'company-insight': 15,
  'admin-ocr-card':  5,
};

module.exports = function createAiCreditsMiddleware(loadAuth, saveAuth) {
  return function aiCreditsGuard(featureKey) {
    return function (req, res, next) {
      const cost = CREDIT_COSTS[featureKey] || 5;
      const auth = loadAuth();
      const sub = auth._subscription;

      // No subscription config = legacy install, no credit tracking
      if (!sub) return next();

      // Enterprise with unlimited credits (aiCredits === null means unlimited)
      if (sub.plan === 'enterprise' && sub.aiCredits == null) return next();

      // AI pack not included and no credits allocated
      const balance = sub.aiCredits != null ? sub.aiCredits : 0;
      if (balance < cost) {
        return res.status(402).json({
          error: `AI Credits 不足（需要 ${cost} 點，目前剩餘 ${balance} 點），請至後台購買加值包`,
          code: 'INSUFFICIENT_CREDITS',
          required: cost,
          balance,
        });
      }

      // Deduct credits
      sub.aiCredits = balance - cost;

      // Record per-call log (keep last 500)
      if (!Array.isArray(sub._creditLog)) sub._creditLog = [];
      sub._creditLog.unshift({
        feature: featureKey,
        cost,
        balance: sub.aiCredits,
        user: req.session?.user?.username || 'unknown',
        ts: new Date().toISOString(),
      });
      if (sub._creditLog.length > 500) sub._creditLog.length = 500;

      saveAuth(auth);
      next();
    };
  };
};

module.exports.CREDIT_COSTS = CREDIT_COSTS;
