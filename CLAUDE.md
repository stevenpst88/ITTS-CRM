# Business Card CRM — Claude 協作指引

## 專案概述

業務名片管理 CRM 系統，功能包含：名片 OCR 辨識、拜訪記錄、商機管理、合約、帳款、業績目標、統一編號查詢、AI 功能（Gemini）、管理員後台。

## 技術架構

- **後端**：Node.js + Express，入口 `server.js`（單一大檔）
- **前端**：原生 HTML/JS，`_client/` 目錄（`index.html` 使用者端、`admin.html` 管理員後台）
- **資料庫**：雙後端切換
  - `DB_BACKEND=postgres`（Supabase）：`app_data` 表，JSONB 欄位
  - `DB_BACKEND=json`（本地）：`data.json`
- **驗證**：express-session（本地開發）/ JWT cookie（Vercel serverless，`middleware/jwtSession.js`）
- **部署**：Vercel，入口 `api/index.js`，region: hnd1（東京）

## 啟動方式

```bash
# 本地開發（本地 JSON 模式）
node server.js
# 或雙擊 啟動CRM.bat

# Supabase 模式需有 .env：
# DB_BACKEND=postgres
# DATABASE_URL=postgresql://...
# GEMINI_API_KEY=...
# SESSION_SECRET=...
```

預設 port 3000。管理員帳號：`admin` / 密碼見 `auth.json`。

## 目錄結構

```
server.js          主後端（所有路由）
_client/
  index.html       使用者前端
  admin.html       管理員後台
db/
  index.js         自動選擇後端
  json.js          本地 JSON 後端
  postgres.js      Supabase PostgreSQL 後端
lib/
  apiMonitor.js    API 使用量統計（Gemini token、統編查詢、Rate Limit）
middleware/
  jwtSession.js    Vercel 用 JWT session 替代
api/
  index.js         Vercel serverless 入口（re-export server.js）
auth.json          帳號資料（本地模式）
data.json          應用資料（本地模式）
```

## 重要設計決策

- `app_data` 表使用 JSONB 一表存所有資料（`id='main'` 主資料、`id='api-stats'` API 監控）
- 所有 Gemini 路由呼叫後執行 `apiMonitor.recordGemini(feature, usageMetadata)`
- Rate limiter 使用自訂 `handler` 以記錄觸發次數
- `writeLog(action, operator, target, detail, req)` 記錄所有 CRUD 操作到 `audit.log.json`
- `writeContactAudit()` 獨立記錄名片欄位級別的異動歷史

## Gemini 功能對應表

| 路由 | Feature Key |
|------|-------------|
| POST /api/admin/ai-ocr-card | admin-ocr-card |
| POST /api/ai/ocr-card | ocr-card |
| POST /api/ai/visit-suggest | visit-suggest |
| POST /api/ai/opp-win-rate | opp-win-rate |
| POST /api/ai/contact-summary | contact-summary |
| POST /api/ai/follow-up-email | follow-up-email |
| POST /api/ai/company-insight | company-insight |

## 管理員後台 Sections

`admin.html` 的 sidebar 使用 `data-sec` 屬性切換：

| data-sec | 說明 |
|----------|------|
| users | 帳號管理 |
| logs | 操作日誌（全 CRUD） |
| contact-audit | 名片稽核紀錄 |
| api-stats | API 使用監控（Gemini token、統編查詢、Rate Limit） |
| infra-stats | 雲端基礎設施（Supabase DB/Storage、Vercel 環境） |
| 📦 資料匯出入 | 可展開群組，含 7 個子項目 |

## 注意事項

- Vercel serverless 不能用本地 JSON 後端（`data.json`）
- `usageMetadata` 可能為 undefined，存取需用 `?.`
- `auth.json` 包含明文密碼 hash，不應上傳公開 repo（已在 .gitignore 排除）
- 統一編號查詢走 GCIS → TWSE/TPEX → DuckDuckGo 三層 fallback
