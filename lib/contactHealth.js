'use strict';
// 客戶健康分數（rule-based）— 索引化版本。
// 原本 GET /api/contacts 對每位聯絡人各掃一遍 visits／receivables／contracts／opportunities，
// 是 O(聯絡人 × 全部紀錄)；1 萬筆名片時主管／管理員要 6～7 秒。
// 這裡改成「先掃一遍建索引（Map／Set），每位聯絡人只做 O(1) 查表」，輸出與原版逐筆相同。
//
// 保留的既有行為（scripts/check-contacts-health-perf.js 有逐項鎖定，改動前請先看該檢查）：
//   1. 帳款以 `r.contactName === c.name` 比對時是嚴格相等：兩邊都是 undefined 也算相符
//      （沒有姓名的聯絡人會被同業務、沒填 contactName 的逾期帳款扣分）。這是原版的副作用，
//      為了「0 差異」刻意保留；要修請單獨改並更新檢查。
//   2. 拜訪日取「預設字串排序的最大者」；任何一筆 visitDate 為 undefined 時排序會把它放最後、
//      reverse 後變第一筆 → new Date(undefined) 為 NaN → 不扣分（與「從沒拜訪」不同，後者扣 40）。
//   3. 公司名以 trim 後比對；空公司名不參與公司比對。
//   4. 本年成交（Won）不檢查 deleted（原版如此）；新商機與合約檢查 deleted。

const DAY = 86400000;

// 容錯 trim：原版對非字串的 company 會丟 TypeError（整個列表 500）；這裡視為不相符，不丟錯。
function T(s) { return typeof s === 'string' ? s.trim() : ''; }

function addTo(map, owner, val) {
  let set = map.get(owner);
  if (!set) { set = new Set(); map.set(owner, set); }
  set.add(val);
}
function hasIn(map, owner, val) {
  const set = map.get(owner);
  return !!set && set.has(val);
}

// 記錄一組拜訪日的「排序最大者」。u=曾出現 undefined；s=字串形（預設 sort 的比較依據）；v=原值
function keepVisit(map, key, d) {
  let e = map.get(key);
  if (!e) { e = { u: false, has: false, s: '', v: undefined }; map.set(key, e); }
  if (d === undefined) { e.u = true; return; }
  const s = String(d);
  if (!e.has || s >= e.s) { e.has = true; e.s = s; e.v = d; }   // >=：與 stable sort + reverse 取到「同值中最後一筆」一致
}

/**
 * @param {object} data   含 visits／opportunities／receivables／contracts 的資料 blob
 * @param {number} [nowMs] 測試用：固定現在時間
 * @returns {(c:object)=>{score:number, icon:string}}
 */
function buildHealthScorer(data, nowMs) {
  data = data || {};
  const nowD = nowMs === undefined ? new Date() : new Date(nowMs);
  const now = nowD.getTime();
  const todayStr = nowD.toISOString().slice(0, 10);
  const yearNum = nowD.getFullYear();
  const in7 = new Date(now); in7.setDate(in7.getDate() + 7);
  const in7Str = in7.toISOString().slice(0, 10);
  const d90 = new Date(now - 90 * DAY).toISOString();

  // 拜訪：依 contactId、依公司（trim）各存最近一次
  const visitById = new Map(), visitByCo = new Map();
  for (const v of (data.visits || [])) {
    if (v.deleted) continue;
    keepVisit(visitById, v.contactId, v.visitDate);
    const vc = T(v.company);
    if (vc) keepVisit(visitByCo, vc, v.visitDate);
  }

  // 逾期帳款：owner → 公司集合／聯絡人姓名集合（姓名含 undefined，見檔頭保留行為 1）
  const recvCo = new Map(), recvName = new Map();
  for (const r of (data.receivables || [])) {
    if (!(r.status !== 'paid' && r.dueDate && r.dueDate < todayStr)) continue;
    const rc = T(r.company); if (rc) addTo(recvCo, r.owner, rc);
    addTo(recvName, r.owner, r.contactName);
  }

  // 7 天內到期、無續約的合約：owner → 公司集合
  const ctCo = new Map();
  for (const ct of (data.contracts || [])) {
    if (ct.deleted || ct.renewDate || !ct.endDate) continue;
    if (!(ct.endDate >= todayStr && ct.endDate <= in7Str)) continue;
    const cc = T(ct.company); if (cc) addTo(ctCo, ct.owner, cc);
  }

  // 商機：近 90 天新商機（未刪除）、本年成交
  const newId = new Map(), newCo = new Map(), wonId = new Map(), wonCo = new Map();
  for (const o of (data.opportunities || [])) {
    const oc = T(o.company);
    if (!o.deleted && o.createdAt >= d90) {
      addTo(newId, o.owner, o.contactId); if (oc) addTo(newCo, o.owner, oc);
    }
    if (o.stage === 'Won' && new Date(o.achievedDate || o.updatedAt || o.createdAt).getFullYear() === yearNum) {
      addTo(wonId, o.owner, o.contactId); if (oc) addTo(wonCo, o.owner, oc);
    }
  }

  return function score(c) {
    let s = 100;
    const co = T(c.company);
    const cId = c.id;

    // 最後拜訪
    const a = visitById.get(cId), b = co ? visitByCo.get(co) : undefined;
    let lastVisitDays = Infinity;
    if (a || b) {
      if ((a && a.u) || (b && b.u)) {
        lastVisitDays = NaN;                       // 有 visitDate 缺漏 → NaN → 不扣分（保留行為 2）
      } else {
        let last;
        if (a && a.has && b && b.has) last = (a.s > b.s ? a : b).v;
        else if (a && a.has) last = a.v;
        else last = b.v;
        lastVisitDays = Math.floor((now - new Date(last).getTime()) / DAY);
      }
    }
    if (lastVisitDays > 60) s -= 40; else if (lastVisitDays > 30) s -= 20;

    // 逾期帳款（姓名為嚴格相等，見保留行為 1）
    if ((co && hasIn(recvCo, c.owner, co)) || hasIn(recvName, c.owner, c.name)) s -= 30;

    // 合約即將到期
    if (co && hasIn(ctCo, c.owner, co)) s -= 20;

    // 近 90 天新商機／本年成交
    if (hasIn(newId, c.owner, cId) || (co && hasIn(newCo, c.owner, co))) s += 10;
    if (hasIn(wonId, c.owner, cId) || (co && hasIn(wonCo, c.owner, co))) s += 20;

    s = Math.max(0, Math.min(120, s));
    return { score: s, icon: s >= 70 ? '🟢' : s >= 40 ? '🟡' : '🔴' };
  };
}

module.exports = { buildHealthScorer };
