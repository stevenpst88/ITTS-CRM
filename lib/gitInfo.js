'use strict';
// 版本資訊頁用：由 Vercel 部署時自動注入的 Git 環境變數取得「線上版本對應哪筆提交」。
// 只讀 5 個白名單變數、不回傳其他環境變數；沒有（本機）就回 null，前端顯示「本機開發」。
// 不執行 child_process（正式站不跑外部指令）。
const SAFE = /^[A-Za-z0-9._-]+$/;
const SHA = /^[0-9a-f]{7,64}$/i;

function deriveGitInfo(env) {
  env = env || {};
  const sha = String(env.VERCEL_GIT_COMMIT_SHA || '').trim();
  if (!SHA.test(sha)) return null;
  const ref = String(env.VERCEL_GIT_COMMIT_REF || '').trim().slice(0, 100);
  const message = String(env.VERCEL_GIT_COMMIT_MESSAGE || '').split(/\r?\n/)[0].trim().slice(0, 100);
  const owner = String(env.VERCEL_GIT_REPO_OWNER || '').trim();
  const slug = String(env.VERCEL_GIT_REPO_SLUG || '').trim();
  const repoUrl = (SAFE.test(owner) && SAFE.test(slug) && owner !== '.' && owner !== '..' && slug !== '.' && slug !== '..')
    ? `https://github.com/${owner}/${slug}/commit/${sha.toLowerCase()}` : null;
  return { sha7: sha.slice(0, 7).toLowerCase(), sha: sha.toLowerCase(), ref, message, repoUrl };
}

module.exports = { deriveGitInfo };
