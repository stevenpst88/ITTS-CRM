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

## 寫入競爭防護（QUOTE_NO_ATOMIC／DB_WRITE_SERIALIZE）

### 為什麼需要
Postgres 模式下整份資料是 `app_data.id='main'` 的一個 JSONB，每個請求都是「讀快取 → 在記憶體改 → 整份寫回」，**最後寫入者勝出**。
兩個 Vercel 實例（或同一實例交錯）各持有舊快照時，後寫的會整份蓋掉先寫的（靜默掉單）；報價單號也是讀快取內的計數器 +1，並發時會拿到同一號。
兩個**預設關閉**的開關（關閉時程式路徑、送出的 SQL 序列、連線池設定都與加入前相同，不建表、不多一次 DB 往返）：

| 環境變數 | 作用 |
|---|---|
| `QUOTE_NO_ATOMIC=1` | 報價單號由 DB 原子計數器發放（`quote_no_seq` / `quote_no_issued`，單一敘述 `INSERT … ON CONFLICT DO UPDATE … RETURNING`）。取號失敗＝fail-closed（503 `QUOTE_NO_UNAVAILABLE`，不退回 blob 計數器） |
| `DB_WRITE_SERIALIZE=1` | 寫入型請求以「全域寫入鎖」序列化：行程內互斥 ＋ DB advisory 鎖（`pg_advisory_xact_lock`，專用連線上的交易；交易結束或連線斷線自動釋放）。拿到鎖後**重新 `db.ready()`** 取最新資料才進 handler，存檔 flush 完才放鎖 |
| `DB_WRITE_LOCK_WAIT_MS`／`DB_WRITE_LOCK_HOLD_MS` | 等鎖上限 8000／持鎖上限 20000（最高 22000，須小於 DB 端 `idle_in_transaction_session_timeout=25s` 與 Vercel `maxDuration=30s`） |
| `DB_WRITE_LOCK_EXEMPT`／`DB_WRITE_LOCK_EXTRA_PATHS` | 追加「不拿鎖」路徑／追加「會寫入的 GET」路徑（見下方「`DB_WRITE_LOCK_EXEMPT` 的比對規則」與 `.env.example`） |
| `DB_QUERY_TIMEOUT_MS` | 僅 `DB_WRITE_SERIALIZE=1`：連線池用戶端查詢逾時（預設 15000） |

### 鎖怎麼運作（`lib/writeGuard.js`、`lib/pgWriteGuard.js`、`db/postgres.js`）
- 只有「已登入、且不會被後面的權限中介層直接拒絕」的寫入型請求（POST/PUT/PATCH/DELETE，加上 `/api/callins`、`/api/announcements/active`、`/api/admin/quote-ledger` 這幾個會寫入的 GET）才排隊拿鎖；純 GET、未登入、被拒絕的請求不拿鎖。中介層掛在強制改密碼／集團角色白名單**之後**。
  「一定會被拒絕」的預先排除（`server.js` 的 `writeGuardPreDeny`）：`/api/admin/*` 預設只有 admin 能用，非 admin 不排隊；**例外**是客戶池 `/api/admin/pool/*`（`POST …/pool/assign`、`POST …/pool/contacts/:id/delete`），handler 允許 admin／manager1／manager2／executive，登記在 `ADMIN_PATH_OPEN_ROLES`。**新增「`/api/admin/` 底下不是只有 admin 能用」的路由時，必須同步登記**，否則會被當成一定被拒絕而逃出寫入鎖（`scripts/check-write-guard.js` L6 會掃描所有 `/api/admin/*` 路由，沒登記就失敗）。
- 等鎖逾時 → 回 **503** `DB_BUSY`（`retryable:true`、`Retry-After: 2`、友善訊息「操作尚未儲存，請稍候再按一次」），handler 完全沒執行。
- 存檔失敗（**只看本請求自己發起的存檔**，以 AsyncLocalStorage 歸屬，不受其他並行請求影響）→ 回 **503** `DB_WRITE_FAILED`，並使該實例快取失效、下個請求整份重抓。同一請求內多次 `db.save` 時**只看最後一次**：每次 save 都是整份覆寫，最後一次成功＝先前（含暫時失敗那次）的變更都已落地，不誤報；最後一次失敗才回 503。快取失效發生在**放鎖之前**（放鎖後排隊的下一個寫入者隨時會進來，不能讓它讀到含失敗變更的快取）。
- DB 層保護：`COMMIT`／取鎖 `BEGIN` 都有用戶端逾時（黑洞連線也不會讓行程內互斥永遠卡住；逾時就銷毀連線，伺服器端因連線斷線釋放鎖）；`lock_timeout`／`statement_timeout`／`idle_in_transaction_session_timeout` 在鎖交易內設定；連線池另有 `query_timeout`。
- `ready()` 的重抓有世代號保護：重抓期間若有人 `save()`，抓回的（可能較舊的）內容會被丟棄，且只採用比目前快取更新的時間戳——否則無鎖請求的並行重抓會把鎖內寫入者剛存的內容蓋回舊的（審查 #1，`scripts/check-write-guard.js` 的 H1／H2 重現與守住）。
- **持鎖期間不做外部網路呼叫**：下列路由不拿鎖（精確清單見 `lib/writeGuard.js` 的 `DEFAULT_EXEMPT`）：`/api/ai/*`、`/api/admin/ai-ocr-card`、`/api/admin/integrations/test`、`/api/admin/integrations/push/batch`、`/api/admin/companies/enrich-gcis`、`/api/admin/bulk-fill-website`、`/api/company-lookup*`、`/api/push/test`、`/api/upload`、`/api/admin/mail/outbox/*`、報價單 export-log／cost-draft／牌價簿匯入預覽、登入登出。
  其中「外部呼叫之後要寫回」的路由與寫入，改成「外部呼叫（無鎖）→ **短鎖**內重讀最新資料、只套用自己改動的欄位」（`writeGuard.applyOnLatest`，會檢查存檔真的成功）：
  - `/api/ai/opp-win-rate`、`/api/ai/contact-summary` 的 AI 快取欄位：短鎖等 3 秒、持有 8 秒，單次嘗試；失敗只記 log（快取丟了下次再算即可）。
  - AI 點數扣款：**只有真的要扣點才取短鎖**（沒有訂閱設定的舊安裝、企業版無上限、點數不足回 402 都不搶鎖）；Web Push 清除失效訂閱：同樣的短鎖。
  - **SAP `push/batch` 的對照／紀錄寫回不能「失敗就算了」**（SAP 端已建立帳號，本機沒記下 `sapId`，重推會在 SAP 重複建立）：用完整等待 8 秒、持有 10 秒、**重試 3 次**；仍失敗時回應帶 `saved:false`＋`warning`，後台顯示「SAP 已推送，但本機對照寫入失敗，請勿重推，請聯絡 IT」，並把待補寫的對照／紀錄留在**行程記憶體**，同一實例的下一次 `push/batch` 開始時先併進使用的資料（避免重複建立）、結尾補寫。限制：行程記憶體隨實例回收而消失，所以 `saved:false` 的警告才是主要防線；看到這個警告時請聯絡 IT 依 SAP 端資料補對照，不要重推。（`lib/sapLinkWriteback.js`）
- multipart 上傳／Excel 匯入：路由在 multer 讀完 body 之後才由 `writeLateLock('<路由路徑>')` 取鎖（上傳時間不佔鎖）；掛上的同時把該路徑**註冊**成「multipart 延後取鎖」路由。只有已註冊的路由，全域中介層才會對 multipart 請求延後取鎖；其他路由即使偽造 `Content-Type: multipart/form-data` 也照一般請求在全域中介層取鎖。**新增 multer 路由時要接 `writeLateLock('<路徑>')` 或加進豁免**（`scripts/check-write-guard.js` N3／N4 會提醒）。
- 冷啟動的遷移／逾期提醒：短鎖（等 1.5 秒、持有 5 秒），取不到就略過、下次冷啟動再做，失敗不阻塞請求。最壞情況（依 Vercel 凍結行為推論、未實測）：冷啟動由 GET 觸發、回應後實例立刻被凍結，而背景工作正持有鎖——看門狗計時器在凍結期間不會觸發，只剩 DB 端 `idle_in_transaction_session_timeout`（約 10 秒）把鎖收回；這段期間其他實例的寫入者最多等 8 秒後回 503 `DB_BUSY`（可重試），之後自動恢復，不遺失資料。工作本身只有幾次 DB 往返（通常不到 1 秒），落在凍結窗口的機率很低。
- 觀測：鎖外的 `db.save`（豁免路由、未列入清單的 GET、鎖放掉之後的背景存檔）會 `console.warn('[db] UNLOCKED_WRITE …')`，附路由與呼叫處；上線後可據此補清單。

### 上線／回滾
1. **上線前**（一定要）：在**拋棄式測試用** Supabase 專案跑 `node scripts/verify-write-race-pg.js`（用法見檔案開頭；Claude／自動化流程不得執行），並確認 `node scripts/check-write-guard.js`、`check-quote-no-atomic.js`、`check-write-unsaved-ui.js` 全綠。
2. **先上 `QUOTE_NO_ATOMIC=1`**（只影響建單路徑、fail-closed、可隨時關），觀察：建單 503 比例、管理員「單號流水帳」API 的 `issuedUnused`／`counterHealed`。
3. **再上 `DB_WRITE_SERIALIZE=1`**。觀察：日誌的 `DB_BUSY`／`DB_WRITE_FAILED` 比例、`UNLOCKED_WRITE` 出現的路由（補進清單或確認可接受）、Supabase 連線數（連線池 max 2→3）。若走 Supabase **transaction pooler(6543)**：等鎖的請求會在 pooler 內占住一條後端連線最久 8 秒，同時寫入者數量要遠小於 pooler 的 pool size（預設約 15）。
4. **Vercel 環境變數的改動只對「新部署」生效**：改完要 Redeploy。出事時的回滾＝把變數設回 0／刪除後 Redeploy，或用 Vercel **Instant Rollback** 回到舊部署。切換期間新舊實例並存約一個部署切換窗口，窗口內有「拿鎖」與「不拿鎖」的實例同時寫入，整份覆寫保護暫時失效（與開啟前相同的風險，不會更糟）。
   - 回滾／關閉 `DB_WRITE_SERIALIZE` 沒有重號風險，可以直接做。
   - 回滾／關閉 **`QUOTE_NO_ATOMIC`**（含 Instant Rollback 到沒有原子取號的舊部署）**有重號風險，必須先照下一節「關閉 `QUOTE_NO_ATOMIC` 之前」處理**。單獨開著 `QUOTE_NO_ATOMIC`（沒開 `DB_WRITE_SERIALIZE`）的期間，不要使用 Instant Rollback。
5. 兩個開關關閉後：`quote_no_seq`／`quote_no_issued` 兩張表與 `data.quoteSeq` 都保留；連線池 max 回到 2；不建表、不拿鎖。是否會重號取決於關閉 `QUOTE_NO_ATOMIC` 前有沒有做下一節的對帳（**唯一規則在下一節**，這裡不重複）。

### 關閉 `QUOTE_NO_ATOMIC` 之前（重號風險；本節是關閉／回滾單號機制的唯一規則）
關閉後單號回到「blob 內計數器 +1」。若 `QUOTE_NO_ATOMIC=1` 期間又沒開 `DB_WRITE_SERIALIZE`，blob 可能被舊快照整份覆蓋、讓 `data.quoteSeq` 倒退，這時直接關閉（或 Instant Rollback 到舊部署）會重發已經給過的號碼。同時開著 `DB_WRITE_SERIALIZE` 時 blob 不會被舊快照覆蓋，風險低，但流程一樣建議照做。
**安全順序：**
1. 用**管理員**帳號打開一次「單號流水帳」（`GET /api/admin/quote-ledger`）。這是**目前唯一**的補救路徑：它只在 `QUOTE_NO_ATOMIC` **仍開啟**時，把 blob 的 `quoteSeq` 補到「已發號紀錄」的最大序號（只增不減，回應的 `counterHealed` 會列出補了哪些月份）。舊部署（Instant Rollback 目標）沒有這段程式，所以**一定要在切換之前**做。
2. 確認 `counterHealed` 已處理（再開一次應為空）、`issuedUnavailable` 不是 true，且最好挑沒人在建單的時段——補完到新設定上線之間若又有人用原子取號建單，新號碼仍可能不在 blob 的計數器裡。
3. 才關閉 `QUOTE_NO_ATOMIC` 並 Redeploy（或 Instant Rollback）。
**來不及做步驟 1 就回滾了**：回滾後、有人建新單之前，請 IT 以 SQL 查 `quote_no_issued` 各月最大 `serial`，核對現有報價單最大單號與 `data.quoteSeq`，必要時先處理再開放建單。
`issuedUnused` 才是判斷「發了號但單沒存成功／被覆蓋」的依據——被覆蓋而丟的單，`gapUnexplained` 會是 0；兩分鐘內剛發出的號碼標 `ISSUED_PENDING`，不計入。

### `DB_WRITE_LOCK_EXEMPT` 的比對規則（以 `lib/writeGuard.js` 的 `isExempt` 實際行為為準；`scripts/check-write-guard.js` K10 鎖住）
路徑一律先轉小寫、去尾斜線再比對。每一項：
- **不以 `/` 結尾**（例 `/api/foo`）＝該路徑本身**及其所有子路徑**（`/api/foo`、`/api/foo/1/x` 都豁免；`/api/foox` 不受影響）。
- **以 `/` 結尾**（例 `/api/foo/`）＝只比對該前綴**底下**的路徑（`/api/foo/1` 豁免；`/api/foo` 本身**不**豁免）。

所以「想豁免單一路徑、不想連子路徑一起」目前沒有對應寫法；追加豁免前先確認它底下沒有其他需要鎖的寫入路由（例如設 `/api/admin/users` 會連 `/api/admin/users/<id>` 的修改、刪除一併不拿鎖）。

### 已知限制（刻意不做）
- `enrich-gcis`、`bulk-fill-website` 兩條被豁免的路由，await 外部查詢之後仍是整份存回（鎖外讀改寫）；根治需要改成「外部呼叫 → 短鎖內逐欄位套用」，變動面大，本次不做。開啟後會有 `UNLOCKED_WRITE` 警告。
- 存檔失敗改回 503 後，該請求已在背景啟動的簽核信件**不會取消**（信件與主資料存檔無關）；使用者重試成功後可能多收一封。log 會記下啟動了幾封。
- 其他鎖外的惰性寫入：報價單舊單品項補 `lid`、`_ensurePoolUser`、管理員 QA 路由 `/api/admin/push/simulate` 的延遲通知；皆為單次性遷移或僅 QA 用，機率低。
- 看門狗到期強制放鎖（單一請求超過 `DB_WRITE_LOCK_HOLD_MS`）後，該請求之後的存檔不再受保護，但回應不會改變（log 有 error）。
- 通知「已讀」等低風險寫入若遇 503，畫面不會假裝已讀（已處理）；其餘只有 toast 提示，不自動重送。
- `updated_at` 比對為毫秒精度；開鎖後兩次提交至少相隔一次網路往返，實務上不會漏偵測。
- 真 Postgres 行為（advisory lock、`lock_timeout`、`ON CONFLICT … RETURNING`、pooler 相容性）本機只有忠實模擬的假 PG（`scripts/_fake-pg.js`）可測，**必須**跑 `verify-write-race-pg.js` 才算驗證過。
