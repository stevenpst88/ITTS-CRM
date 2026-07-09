# Business Card CRM (ITTS-CRM) — Claude 協作指引

業務名片管理 CRM：名片 OCR、拜訪、商機、合約、帳款、業績目標、統編查詢、AI 功能（Gemini）、管理員後台、SAP 整合。

## 技術架構

- **後端**：Node.js + Express，單一大檔 `server.js`（8000+ 行——別整檔讀。定位交給 Explore subagent，主對話只在拿到 file:line 後讀 <100 行的目標段落）
- **前端**：原生 HTML/JS。`_client/index.html` + `app.js`（使用者端）、`_client/admin.html`（管理員後台，JS 內嵌）
- **資料庫**：`DB_BACKEND=postgres`（Supabase `app_data` 表 JSONB 單表）或 `json`（本地 `data.json`）
- **驗證**：本地 express-session / Vercel JWT cookie（`middleware/jwtSession.js`）
- **部署**：Vercel（入口 `api/index.js`，region hnd1）。push main = 自動部署正式站
- **共用模組**：`lib/secretBox.js`（AES-256-GCM 機敏加密）、`lib/productCatalog.js`（商品目錄）、`lib/apiMonitor.js`（API 用量）

## 主資料 blob 的命名空間（`app_data.content` / `data.json`）

CRM 核心：`contacts / companies / opportunities / visits / contracts / receivables / callins / keyAccounts`
獨立命名空間（勿與核心混寫）：`yoyRevenue`（YoY 報表）、`integrations / integrationMappings / integrationLinks / integrationLogs`（SAP 整合）、`productCatalog`（商品目錄）、`_auth`（雲端帳號）

## 高風險紅線

- `data.json`、`auth.json`、`_preview_server.js`、`docs/` 皆不入 repo（gitignore 或刻意不加）。
- 機敏欄位（integration 的 password/clientSecret）永不回傳前端——只回 `hasPassword` 布林。改整合相關程式前先讀 `lib/secretBox.js` 的用法。
- commit 與 push 各自需要使用者明確要求；說 commit 就只 commit，說 push 才 push（push main = 正式部署）。
- 高風險 / 需反覆試錯的改動先在 **Demo 環境**驗證（repo `stevenpst88/ITTS-CRM-Demo` → `itts-crm-demo.vercel.app`，獨立 Supabase，壞了不影響正式）。同步方式（在 Demo repo 本地副本、用 Bash 工具執行；含 push，需使用者明說要同步 Demo）：`git fetch upstream; git merge upstream/main; git push`。

## 已踩坑的事實（改相關功能前先讀）

- **Vercel/Supabase 快取**（2026-07 驗證）：`db/postgres.js` 的 `REFRESH_TTL = 0`——每次 API 請求先做輕量 stale check（只抓 `updated_at`），DB 被其他實例改過就自動完整重抓。直接用 SQL 改 DB 只要 `updated_at` 有更新就會被抓到；僅當繞過寫入路徑、`updated_at` 未變時，才需要空 commit 強制重部署。
- **雲端/地端帳號分離**：地端 `auth.json` 的 admin 是 `Admin`（大寫）；雲端 `_auth` 的是 `admin`（小寫）。兩邊獨立不同步。
- **SAP Sales Cloud V2 API**：欄位格式、端點名、PATCH header、地址規則全部有雷。動 `push/batch` 相關程式前**必讀** `docs/SAP_V2_API_gotchas.md`；若無此檔（docs/ 不入 repo），讀 `C:/Users/steven.lee/.claude/projects/C--Users-steven-lee/memory/sap_v2_api_gotchas.md`；連這都沒有（別台電腦）→ 第一步先打 `GET /api/admin/integrations/sap-inspect` 讀 SAP 真實資料樣本，**禁止猜欄位格式**。速記三鐵則：customerRole 是字串非陣列；Contact 端點是 `contact-person-service/contactPersons`；PATCH 要 `If-Match: *` + `Content-Type: application/merge-patch+json`。
- **本機測試訣竅**：密碼雜湊用 bcryptjs；臨時帳號必須含 `passwordChangedAt` 否則被強制改密迴圈擋住；聯絡人 `bu` 傳字串（`ERP/ITS/MDM/CRM`）；CORS 白名單預設只有 localhost:3000，preview 用 3001 要加。測完還原 `auth.json`/`data.json`。
- **統編查詢**：`/api/company-lookup?taxId=` 加 `&basic=1` 只查 GCIS（~4s）；不加會跑上市櫃+財務（冷啟動 ~62s，超過 Vercel maxDuration 30s 會逾時）。
- **商品目錄**：商機/合約的 `product` 是**純文字非外鍵**。改名不自動連動歷史——要連動就走 `POST /api/admin/product-catalog` 的 renames 機制（有 preview 與確認）。

## Gemini AI

- 模型設定在 `ai/gemini.js`：全部功能統一 `gemini-flash-lite-latest`。
- 所有 Gemini 路由呼叫後執行 `apiMonitor.recordGemini(feature, usageMetadata)`；`usageMetadata` 可能 undefined，用 `?.`。
- 功能對應：admin-ocr-card / ocr-card / visit-suggest / opp-win-rate / contact-summary / follow-up-email / company-insight。

## 管理員後台（admin.html）

側欄以 `data-sec` 切換 section，重 section 採 lazy-init（首次點擊才載入，如 `initIntegration`、`initProductCatalog`）。新增後台功能照此模式：側欄項 + `sec-*` div + init 分派 + `adminFetch`（自動處理 session 過期）。

## 啟動

```bash
node server.js        # 本地 JSON 模式，port 3000（或雙擊 啟動CRM.bat）
# Supabase 模式：.env 設 DB_BACKEND=postgres + DATABASE_URL + GEMINI_API_KEY + SESSION_SECRET
```

## 稽核

所有 CRUD 走 `writeLog(action, operator, target, detail, req)`；名片欄位級異動另走 `writeContactAudit()`。新增寫入型 API 時兩者不可省。
