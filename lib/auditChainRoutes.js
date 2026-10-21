'use strict';
/**
 * 稽核防竄改的 HTTP 端點（演算法在 lib/auditChain.js、lib/historyChain.js）。
 *
 *   GET       /api/admin/audit-chain/head            輕量狀態（鏈是否啟用、鏈頭、上次驗證、最近錨點）——版本資訊頁用
 *   GET|POST  /api/admin/audit-chain/verify          驗證整條（或 ?from=&to= 序號區間）；POST body／query 可帶 anchor 比對（證明沒被截斷／回溯）
 *   GET       /api/admin/audit-chain/anchor          目前的簽章錨點；?download=1 → 下載 JSON 檔
 *   GET       /api/admin/quotations/chain-check      全部報價單的簽核歷程鏈檢查（分批讓出事件迴圈）
 *   GET       /api/cron/audit-anchor                 每日排程：產生錨點、存最近 90 份在主資料（data.auditAnchors）、（若有設）推送 webhook
 *                                                    Authorization: Bearer <CRON_SECRET>（沒設 CRON_SECRET 一律 401）
 * 全部管理員專用（cron 除外），不回傳金鑰；每次查驗都寫稽核事件。單張報價單的 history-check 在 lib/quoteRoutes.js（能看這張單的人可查）。
 */
const AC = require('./auditChain');
const HC = require('./historyChain');

const KEEP_ANCHORS = 90;
const IDEMPOTENT_MS = 20 * 3600 * 1000;
const MAX_ANCHOR_QUERY = 2000;

function register(app, deps) {
  const { requireAdmin, writeLog, db } = deps;
  const env = deps.env || process.env;
  const opName = (req) => (req.session && req.session.user && req.session.user.username) || 'unknown';
  const flushSafe = async () => { try { if (db && typeof db.flush === 'function') await db.flush(); } catch (e) { /* 稽核寫入失敗不影響回應 */ } };
  const getCfg = () => (db.auditStore && db.auditStore.chain) || AC.getConfig(env);
  const webhookInfo = () => {
    const u = env.AUDIT_ANCHOR_WEBHOOK_URL;
    if (!u) return { configured: false, host: '' };
    const s = AC.checkWebhookUrlShape(u);
    return { configured: true, valid: s.ok, host: s.ok ? s.host : '', reason: s.ok ? '' : s.reason };   // 只回主機名；網址路徑常帶祕密 token，不回傳
  };

  /** 鏈的資料來源與狀態；src 為 null 代表目前不是新表模式 */
  async function chainCtx() {
    const cfg = getCfg();
    let src = null, readable = false;
    try { src = await db.auditStore.chainSource(); if (src) readable = await src.chainReadable(); } catch (e) { src = null; }
    let reason = '';
    if (!src) reason = 'TABLE_MODE_OFF';
    else if (!cfg.canVerify) reason = cfg.disabledReason === 'OFF' || !cfg.disabledReason ? 'NO_KEY' : cfg.disabledReason;
    else if (!readable) reason = cfg.flag ? 'COLUMNS_NOT_READY' : 'OFF';
    return { cfg, src, readable, reason, usable: !!src && cfg.canVerify && readable };
  }
  const notEnabled = (res, c) => res.status(409).json({ error: '稽核鏈尚未啟用或目前無法驗證', code: 'CHAIN_NOT_ENABLED', enabled: false, reason: c.reason });
  const lastOf = (arr) => (Array.isArray(arr) && arr.length ? arr[arr.length - 1] : null);
  const brief = (a) => (a ? { seq: a.seq, ts: a.ts, count: a.count } : null);

  app.get('/api/admin/audit-chain/head', requireAdmin, async (req, res) => {
    res.set('Cache-Control', 'no-store');
    try {
      const c = await chainCtx();
      const data = (() => { try { return db.load() || {}; } catch (e) { return {}; } })();
      const out = {
        enabled: !!(c.cfg.enabled && c.src && c.readable),      // 新事件目前是否會上鏈
        configured: !!c.cfg.flag, canVerify: c.usable, reason: c.usable && c.cfg.enabled ? '' : c.reason,
        mode: db.auditStore.mode(), kid: c.cfg.kid, kids: c.cfg.kids, secretSource: c.cfg.secretSource,
        headSeq: 0, headHash: '', writes: null,
        lastVerify: data.auditChainLastVerify || null, lastAnchor: brief(lastOf(data.auditAnchors)), anchorsStored: Array.isArray(data.auditAnchors) ? data.auditAnchors.length : 0,
        webhook: webhookInfo(), cronSecretConfigured: typeof env.CRON_SECRET === 'string' && env.CRON_SECRET !== '',
      };
      if (c.usable) {
        const h = await c.src.chainHead();
        if (h) { out.headSeq = h.seq; out.headHash = h.row_hash; }
        const st = db.auditStore.stats();
        out.writes = st.chain || null;
        try { const u = await c.src.chainUncovered(); out.unchainedAfterStart = (u && u.unchainedAfterStart) || 0; } catch (e) { out.unchainedAfterStart = null; }   // D-08：鏈啟用後不在鏈上的事件數（跨實例、含 breakRisk 降級寫入）
      }
      res.json(out);
    } catch (e) {
      console.error('[audit-chain head]', e && e.message);
      res.status(500).json({ error: '讀取稽核鏈狀態失敗' });
    }
  });

  async function verifyHandler(req, res) {
    res.set('Cache-Control', 'no-store');
    try {
      const c = await chainCtx();
      if (!c.usable) return notEnabled(res, c);
      let anchor = null;
      const raw = (req.body && req.body.anchor) || req.query.anchor;
      if (raw) {
        try {
          if (typeof raw === 'string') { if (raw.length > MAX_ANCHOR_QUERY) throw new Error('too long'); anchor = JSON.parse(raw); }
          else anchor = raw;
          if (!anchor || typeof anchor !== 'object' || Array.isArray(anchor)) throw new Error('bad');
        } catch (e) { return res.status(400).json({ error: '錨點格式不正確（需為下載的錨點 JSON）', code: 'BAD_ANCHOR' }); }
      }
      const q = req.query || {};
      // D-14：沒有手動貼錨點時，預設比對最新一份已存錨點（偵測尾端被截／回溯）；?storedAnchor=0 可關閉（例如換金鑰後舊錨點無法驗簽）
      let anchorSource = anchor ? 'provided' : '';
      if (!anchor && String(q.storedAnchor) !== '0') {
        try { const sa = lastOf((db.load() || {}).auditAnchors); if (sa && typeof sa === 'object') { anchor = sa; anchorSource = 'stored'; } } catch (e) { /* 讀不到就不比對 */ }
      }
      const r = await AC.verifyChain(c.src, c.cfg, { from: q.from, to: q.to, anchor });
      if (anchorSource) r.anchorSource = anchorSource;
      const bk = r.firstBreak;
      writeLog('AUDIT_CHAIN_VERIFY', opName(req), 'audit_events',
        `稽核鏈驗證：${r.ok ? '通過' : '失敗'}；檢查 ${r.checked} 筆；鏈頭序號 ${r.headSeq}${r.complete ? '' : '（未驗完，下次從 ' + r.nextFrom + ' 起）'}${anchor ? '；比對' + (anchorSource === 'stored' ? '已存' : '') + '錨點 seq=' + (Number(anchor.seq) || 0) + (r.anchor && r.anchor.ok ? ' 相符' : ' 不符') : ''}${bk ? `；第一個斷點 seq=${bk.seq} ${bk.reason}` : ''}`, req);
      try {
        const data = db.load();
        data.auditChainLastVerify = { ts: r.ts, ok: r.ok, checked: r.checked, complete: r.complete, headSeq: r.headSeq, by: opName(req), firstBreak: bk ? { seq: bk.seq, reason: bk.reason } : null, anchorChecked: !!anchor, anchorSource: anchorSource || '' };
        db.save(data);
      } catch (e) { console.error('[audit-chain verify] 存上次結果失敗:', e && e.message); }
      await flushSafe();
      res.json(r);
    } catch (e) {
      console.error('[audit-chain verify]', e && e.message);
      res.status(500).json({ error: '驗證失敗（資料庫暫時無法使用？）' });
    }
  }
  app.get('/api/admin/audit-chain/verify', requireAdmin, verifyHandler);
  app.post('/api/admin/audit-chain/verify', requireAdmin, verifyHandler);

  async function currentAnchor(c) {
    const head = await c.src.chainHead();
    const count = await c.src.chainCount();
    return AC.makeAnchor(c.cfg, head, count);
  }
  app.get('/api/admin/audit-chain/anchor', requireAdmin, async (req, res) => {
    res.set('Cache-Control', 'no-store');
    try {
      const c = await chainCtx();
      if (!c.usable) return notEnabled(res, c);
      const a = await currentAnchor(c);
      writeLog('AUDIT_CHAIN_ANCHOR', opName(req), 'audit_events', `取得稽核鏈錨點：seq=${a.seq}，共 ${a.count} 筆${req.query.download === '1' ? '（下載）' : ''}`, req);
      await flushSafe();
      if (req.query.download === '1') res.set('Content-Disposition', `attachment; filename="audit-anchor-${a.seq}-${a.ts.slice(0, 10)}.json"`);
      res.json(a);
    } catch (e) {
      console.error('[audit-chain anchor]', e && e.message);
      res.status(500).json({ error: '產生錨點失敗' });
    }
  });

  app.get('/api/admin/quotations/chain-check', requireAdmin, async (req, res) => {
    res.set('Cache-Control', 'no-store');
    try {
      const r = await HC.checkAll(db.load(), AC.getConfig(env));
      writeLog('QUOTE_HISTORY_CHAIN_CHECK', opName(req), 'quotations', `報價單簽核歷程鏈檢查：${r.ok === null ? '未啟用' : r.ok ? '通過' : '失敗'}；${r.quotes} 張單、${r.entries} 筆歷程${r.firstBreak ? '；第一個異常 ' + r.firstBreak.quoteNo + ' #' + r.firstBreak.index + ' ' + r.firstBreak.reason : ''}`, req);
      await flushSafe();
      res.json(r);
    } catch (e) {
      console.error('[quotations chain-check]', e && e.message);
      res.status(500).json({ error: '檢查失敗' });
    }
  });

  // ── 每日排程：錨點 ──
  app.get('/api/cron/audit-anchor', async (req, res) => {
    res.set('Cache-Control', 'no-store');
    if (!AC.cronAuthorized(req, env)) return res.status(401).json({ error: '未授權' });
    try {
      const c = await chainCtx();
      if (!c.usable) {
        // D-07：已設定要啟用（AUDIT_CHAIN=1）卻無法使用＝錨點停了，必須讓 Vercel／監控看得到（503＋console.error）；沒設定啟用則維持 200 enabled:false
        if (c.cfg && c.cfg.flag) {
          console.error('AUDIT_CHAIN_DISABLED', '每日錨點排程無法執行：', c.reason || 'UNKNOWN');
          return res.status(503).json({ ok: false, enabled: false, configured: true, reason: c.reason, error: '稽核鏈已設定啟用但目前無法使用，錨點未產生' });
        }
        return res.json({ ok: true, enabled: false, reason: c.reason });
      }
      const a = await currentAnchor(c);
      const data = db.load();
      const list = Array.isArray(data.auditAnchors) ? data.auditAnchors : [];
      // D-14：冪等——Vercel cron 可能重複送達；最新一份存檔錨點 seq 相同且距今不到 20 小時就不再新增（避免擠掉較舊的真錨點），仍回 200
      const prevA = list.length ? list[list.length - 1] : null;
      if (prevA && Number(prevA.seq) === Number(a.seq) && Date.parse(a.ts) - Date.parse(prevA.ts) < IDEMPOTENT_MS && Date.parse(a.ts) >= Date.parse(prevA.ts)) {
        return res.json({ ok: true, enabled: true, seq: a.seq, count: a.count, stored: list.length, pushed: null, deduped: true });
      }
      list.push(a);
      data.auditAnchors = list.slice(-KEEP_ANCHORS);
      db.save(data);
      let pushed = null;
      if (env.AUDIT_ANCHOR_WEBHOOK_URL) {
        pushed = await AC.pushAnchor(env.AUDIT_ANCHOR_WEBHOOK_URL, a, deps.pushOpts);
        if (!pushed.ok) console.error('AUDIT_ANCHOR_PUSH_FAILED', pushed.reason, webhookInfo().host);   // 只記主機名，不記完整網址（路徑常帶 token）
      }
      // D-08：鎖逾時等原因讓事件「改為不上鏈寫入」時，把風險帶進排程回應與稽核（不改 ok：錨點本身是成功的）
      let risk = null;
      try {
        const st = db.auditStore.stats();
        const u = await c.src.chainUncovered();
        risk = { breakRisk: (st.chain && st.chain.breakRisk) || 0, unchainedAfterStart: (u && u.unchainedAfterStart) || 0 };
        if (risk.breakRisk > 0 || risk.unchainedAfterStart > 0) console.error(AC.CHAIN_MARK.risk, '每日錨點檢查：本實例改為不上鏈寫入 ' + risk.breakRisk + ' 筆；鏈啟用後仍有 ' + risk.unchainedAfterStart + ' 筆事件不在鏈上');
      } catch (e) { risk = null; }
      writeLog('AUDIT_CHAIN_ANCHOR_CRON', 'cron', 'audit_events', `每日稽核鏈錨點：seq=${a.seq}，共 ${a.count} 筆；已存 ${data.auditAnchors.length} 份${pushed ? '；webhook ' + (pushed.ok ? '成功' : '失敗(' + pushed.reason + ')') : ''}${risk && (risk.breakRisk || risk.unchainedAfterStart) ? `；未上鏈風險 breakRisk=${risk.breakRisk}、鏈後未上鏈 ${risk.unchainedAfterStart} 筆` : ''}`, req);
      await flushSafe();
      res.json({ ok: true, enabled: true, seq: a.seq, count: a.count, stored: data.auditAnchors.length, pushed: pushed ? { ok: pushed.ok, reason: pushed.reason || '' } : null, risk });
    } catch (e) {
      console.error('[audit-anchor cron]', e && e.message);
      res.status(500).json({ ok: false, error: '錨點排程失敗' });
    }
  });
}

module.exports = { register, KEEP_ANCHORS };
