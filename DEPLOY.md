# ITTS CRM — Vercel + Supabase 部署指南

本文件記錄從本地遷移到雲端的完整步驟。

## 架構

```
Vercel（單一 serverless function）    Supabase
  ├─ public/* 靜態檔                  ├─ Postgres：app_data
  └─ api/index.js → Express app       └─ Storage：uploads bucket
```

## 部署前準備

### 1. Supabase 專案設定
1. 進入 Supabase Dashboard → 專案 `stevenpst88`
2. **SQL Editor** → 貼上 `supabase/schema.sql` → Run
   - 既有環境（先前用舊版 schema.sql 建的）另外執行 `supabase/revoke_app_data_public.sql`（見下方「資料表權限（SEC-01）」）
3. **Storage** → 確認 `uploads` bucket 存在（private）
4. 取得以下資訊（Settings → API）：
   - `Project URL`
   - `anon public` key
   - `service_role` key（機密！）
5. 取得 DB 連線字串（Settings → Database → Connection string → URI）

### 2. GitHub Repo
已設定：`https://github.com/stevenpst88/ITTS-CRM.git`

### 3. Vercel 專案
1. Vercel Dashboard → Add New → Project
2. Import `stevenpst88/ITTS-CRM`
3. Framework Preset: **Other**
4. Build Command: 留空（由 vercel.json 控制）
5. **Environment Variables**（重要！全部設好再 deploy）：

| 變數 | 值 | 說明 |
|---|---|---|
| `NODE_ENV` | `production` | |
| `SESSION_SECRET` | 64 字元 hex | `node -e "console.log(require('crypto').randomBytes(64).toString('hex'))"` |
| `JWT_SECRET` | 64 字元 hex | 同上重新產生一次 |
| `DB_BACKEND` | `postgres` | |
| `DATABASE_URL` | `postgresql://...` | Supabase DB URI |
| `SUPABASE_URL` | `https://xxx.supabase.co` | |
| `SUPABASE_ANON_KEY` | `eyJ...` | |
| `SUPABASE_SERVICE_ROLE_KEY` | `eyJ...` | |
| `STORAGE_BACKEND` | `supabase` | |
| `SUPABASE_BUCKET` | `uploads` | |

6. 點 Deploy

## 部署後驗證

### Smoke test
```bash
# 換成你的 Vercel URL
URL=https://itts-crm.vercel.app

# 1. 首頁應回傳 login.html（302 或 200）
curl -I $URL/

# 2. API 未登入應回 401
curl $URL/api/contacts

# 3. 登入（待建第一個 admin）
```

### 建立第一個帳號
首次部署後 DB 是空的。兩種方式之一：

**方式 A：本地灌 auth.json 到 DB**
```bash
# 1. 修改本地 .env 讓 DB_BACKEND=postgres
# 2. 跑 scripts/seed-admin.js（待建）
```

**方式 B：直接在 Supabase SQL Editor 執行**
```sql
UPDATE app_data SET content = jsonb_set(
  content,
  '{users}',
  '[{"username":"admin","passwordHash":"...","role":"admin"}]'::jsonb
) WHERE id='main';
```

## 本地開發

```bash
# .env 保持 DB_BACKEND=json、STORAGE_BACKEND=local、SESSION_BACKEND=memory
npm start
# http://localhost:3000
```

## 切回本地測 Postgres
```bash
# 臨時在 .env 改：
DB_BACKEND=postgres
STORAGE_BACKEND=supabase
SESSION_BACKEND=jwt
# 加上 SUPABASE_* 變數
npm start
```

## 回滾
所有改動保留完整相容：刪 `.env` 裡的 `DB_BACKEND=postgres` → 立刻回到本地 JSON 模式。

## 資料表權限（SEC-01）

**為什麼**：整個 CRM 資料（含帳號密碼雜湊、加密的 SAP 密碼）都在 `public.app_data`。Supabase 預設把 public 資料表開放給 Data API（用公開的 anon key 就能呼叫），若 RLS 關閉又沒撤權，任何人拿到 anon key 都能讀寫整張表。本應用程式只用 `DATABASE_URL` 直連 Postgres（`db/postgres.js`），前端不用 Supabase client，所以 anon／authenticated 不需要任何權限。

**做法**：執行 `supabase/revoke_app_data_public.sql`（可重複執行）：撤銷 anon／authenticated／PUBLIC 權限並開啟 RLS（不建 policy＝Data API 一律拒絕）。應用程式連線的 postgres 角色是表的 owner，不受影響。新環境用最新的 `supabase/schema.sql` 建立即已包含。先在 DEV 執行驗證，再套用正式環境。

**驗證（唯讀）**：
```sql
SELECT c.relname AS table_name, c.relrowsecurity AS rls_on,
       has_table_privilege('anon', c.oid, 'SELECT') AS anon_select,
       has_table_privilege('anon', c.oid, 'UPDATE') AS anon_update,
       has_table_privilege('authenticated', c.oid, 'SELECT') AS authed_select
  FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
 WHERE n.nspname = 'public' AND c.relkind = 'r' ORDER BY 1;
-- 預期：app_data 的 rls_on = true，anon_select／anon_update／authed_select 皆 false
```
```bash
curl "$SUPABASE_URL/rest/v1/app_data?select=id" -H "apikey: $ANON_KEY"
# 預期：401／403 或空結果，不可回出資料
```

DATABASE_URL 改用低權限角色（`itts_app`）另案處理，不在此次範圍。
