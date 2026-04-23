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
