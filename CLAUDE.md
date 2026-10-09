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

## 開發流程（分支／PR／多個 session——2026-10 檢討後定案，所有 session 都要遵守）

- **一個主題一條分支、一份獨立工作資料夾**：`git worktree add <資料夾> -b feat/<主題> main`。**絕不讓兩個 session 共用同一份工作目錄**（曾因此被迫手動三方合併、逐段拆共用檔案的 commit）。第二個 session／agent 要動同樣的檔案，也開自己的 worktree。
- **走 PR，不直接 commit 到 main**：做完開 PR（`gh pr create`，附截圖與驗證結果），由使用者看 diff、核可後合併。一個 PR 一個主題，不要事後再拆。push main 等於正式部署，仍須使用者明確說 push（上方紅線不變）。
- **高風險改動先進 Demo 驗證**（見上方紅線）；PR 預覽部署若連到正式資料庫，不可拿來寫入測試資料。
- **測試只用隔離沙盒**：另開埠（不用 3000）、獨立資料檔複本、臨時 `_s_*` 測試帳號，測完還原；不碰真實 `data.json`／`auth.json`／`audit.log.json`，也不在另一個 session 正在用 3000 埠時啟動它。
- **commit 衛生**：永遠明確指定檔案 add，**禁止 `git add -A`／`git add .`**（工作區常有 pptx、備份 JSON、`docs/`、`_preview_server.js` 等未追蹤檔，repo 是公開的）；新增的檔案漏 add 會讓部署後整站 `MODULE_NOT_FOUND`，commit 前逐一核對。公開 repo 不得出現帳密、客戶名、人名、本機路徑；腳本一律讀環境變數。
- **驗證節奏**：調整畫面階段「小改自己做、只跑相關檢查、給截圖」；完整的獨立驗證（全新視角、自己重跑）只在準備 commit／PR 前做一次。
- **多 agent 紀律**：同一件事不要同時派兩個 agent；追加需求前先確認舊的 agent 真的停了（傳訊息會喚醒它，反而造成重複實作）；進度與決定寫進檔案（狀態檔／報告），被用量上限切斷後從檔案接續，不從頭重做。
- **分支 session（fork）**：只是同專案的另一個對話視窗，不是 Git 分支；沒有「合併」，做完就 commit 進自己的分支／PR 後關閉；後續工作回主視窗做。

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
