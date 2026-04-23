# ITTS CRM — 業務名片管理系統

內部使用的 CRM，整合名片管理、業務日報、商機追蹤、銷售預測、報價單（含毛利分析）、合約與應收帳款。

## 技術棧

| 層 | 本地開發 | 雲端部署 |
|---|---|---|
| Frontend | 靜態 HTML/JS/CSS | Vercel CDN |
| Backend  | Express (Node.js) | Vercel Serverless |
| 資料庫   | `data.json` 檔案  | Supabase Postgres |
| 檔案儲存 | `uploads/` 資料夾 | Supabase Storage |
| 認證     | express-session   | JWT + httpOnly cookie |

## 本地開發

```bash
npm install
npm start
# http://localhost:3000
```

`.env` 範本見 [`.env.example`](.env.example)。

## 部署

完整部署步驟請見 [`DEPLOY.md`](DEPLOY.md)。

## 架構特色

### 可插拔後端設計
所有外部依賴（DB、Storage、Session）都透過環境變數切換後端：

```bash
DB_BACKEND=json|postgres
STORAGE_BACKEND=local|supabase
SESSION_BACKEND=memory|jwt
```

本地用 `json/local/memory` 組合，部署用 `postgres/supabase/jwt` 組合，**業務邏輯零改動**。

## 主要功能

- 📋 名片 / 聯絡人管理（含公司分組、名片圖片上傳）
- 📝 業務日報 / 拜訪記錄
- 🎯 商機推進追蹤 + 銷售預測報表
- 💼 報價單管理（含 PNL 毛利分析 TAB）
- 📄 合約管理（ERP MA / SAP License MA）
- 💰 應收帳款追蹤
- 📞 Call-in Pass 分派
- 👥 多角色權限（admin / manager1 / manager2 / user / secretary）
- 📊 視覺化儀表板（Chart.js）

## License

Internal use only. All rights reserved.
