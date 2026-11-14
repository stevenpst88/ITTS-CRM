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

CRM 核心：`contacts / companies / opportunities / visits / contracts / receivables / keyAccounts`（舊的 `callins` 為已移除的 Call-in Pass 遺留資料：仍留在主資料 blob 內但沒有任何 UI／API，程式不得假設它存在或不存在）
獨立命名空間（勿與核心混寫）：`yoyRevenue`（YoY 報表）、`integrations / integrationMappings / integrationLinks / integrationLogs`（SAP 整合）、`productCatalog`（商品目錄）、`announcements`（公告：含 readBy 已讀回條、shown 送達時間；邏輯在 lib/announcements.js）、`pricebook`（報價牌價簿：顧問人天牌價／成本）、`quoteSplitGuard`（防拆單提醒設定；報價單上另有 `q.splitGuard` 送簽快照，僅 warn 模式寫入；邏輯在 lib/quoteSplitGuard.js，環境變數 QUOTE_SPLIT_GUARD=off|log|warn 預設 off）、`quoteIssuances`／`quoteStampUses`（報價文件發出紀錄＋蓋章憑證使用紀錄：只增不改，邏輯在 lib/quoteIssuance.js；蓋章 `GET /quotations/:id/seal` 需單次使用憑證）、`_auth`（雲端帳號）

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
- **SAP Sales Cloud V2 API**：欄位格式、端點名、PATCH header、地址規則全部有雷。動 `push/batch` 相關程式前**必讀** `docs/SAP_V2_API_gotchas.md`；若無此檔（docs/ 不入 repo），讀 `~/.claude/projects/<家目錄編碼>/memory/sap_v2_api_gotchas.md`（Claude Code 的 auto-memory 資料夾）；連這都沒有（別台電腦）→ 第一步先打 `GET /api/admin/integrations/sap-inspect` 讀 SAP 真實資料樣本，**禁止猜欄位格式**。速記三鐵則：customerRole 是字串非陣列；Contact 端點是 `contact-person-service/contactPersons`；PATCH 要 `If-Match: *` + `Content-Type: application/merge-patch+json`。
- **本機測試訣竅**：密碼雜湊用 bcryptjs；臨時帳號必須含 `passwordChangedAt` 否則被強制改密迴圈擋住；聯絡人 `bu` 傳字串（`ERP/ITS/MDM/CRM/資安`）；CORS 白名單預設只有 localhost:3000，preview 用 3001 要加。測完還原 `auth.json`/`data.json`。
- **資安審查快修的環境變數開關**（2026-10，分支 fix/secrev-quick；全部有預設、不設也能跑）：`SESSION_PWD_BINDING=off` 關閉「登入憑證綁密碼版本」（預設開；舊 token 無 pwa 欄位一律放行）；`LOGIN_LOCK_SCOPE=name` 把登入鎖定鍵還原成只看帳號名稱（預設 ip＝帳號＋來源 IP）；`STAMP_TOKEN_QUERY=allow` 暫時接受網址 `?stampToken=`（預設拒絕，只收 `X-Stamp-Token` 標頭）；`QUOTE_ISSUE_REUSE_SEC`（預設 60，0＝不重用）與 `QUOTE_ISSUE_RATE_PER_MIN`（預設 20，0＝不限）控制發出紀錄的重用與每人限流。檢查腳本：`scripts/check-secrev-quick.js`。
- **程式碼審查 PR-2／PR-3／M-01 的環境變數開關**（2026-10，分支 fix/codrev-quote-audit，疊在 fix/secrev-quick 之上；全部有預設）：`SESSION_IDLE_MINUTES` 現在由 `middleware/jwtSession.js` 讀取（未設＝30 分鐘；`SESSION_IDLE_MS` 測試用優先；背景輪詢路徑與長度的單一來源在 `lib/sessionIdle.js`，`loginGuard.idleMiddleware` 只服務 `SESSION_BACKEND=memory`）；`QUOTE_CONFIG_OPEN=1` 還原 `GET /api/quote-approval/config` 只驗登入；`QUOTE_FLOW_MODE_AUDIT_NON_PROD=1`／`QUOTE_FLOW_MODE_THROTTLE_MS` 控制引擎模式紀錄（Vercel 非 production 部署預設不寫、同行程 10 分鐘節流）；`APPROVAL_CONFIG_MAX_PENDING`（預設 3，0＝不限）；`AUDIT_DDL_LOCK_TIMEOUT_MS`（預設 5000）。`/api/cron/audit-anchor` 在 `AUDIT_CHAIN=1` 卻無法使用時回 503；`/api/admin/audit-chain/verify` 預設比對最新已存錨點（`?storedAnchor=0` 關閉）。檢查腳本：`scripts/check-codrev-quote-audit.js`。
- **統編查詢**：`/api/company-lookup?taxId=` 加 `&basic=1` 只查 GCIS（~4s）；不加會跑上市櫃+財務（冷啟動 ~62s，超過 Vercel maxDuration 30s 會逾時）。
- **商品目錄**：商機/合約的 `product` 是**純文字非外鍵**。改名不自動連動歷史——要連動就走 `POST /api/admin/product-catalog` 的 renames 機制（有 preview 與確認）。
- **商機 SAP Deal ID**（2026-10；畫面一律稱「SAP Deal ID」，程式識別字 `dealId`／`requiresDealId`／`dealRegisterUrl`／錯誤碼 `DEAL_ID_REQUIRED` 不改名）：商品目錄項目可設 `requiresDealId`／`dealRegisterUrl`（只收 http/https，存檔時 `new URL` 驗證）。**格式 = 9 碼半形數字字串**（開頭 0 保留；NFKC 轉半形、去空白／`-`／`,` 後驗，存正規化值），唯一來源是 `_client/deal-id-format.js`（前端 `window.DealIdFormat`、後端經 `lib/productCatalog.js` 的 `normalizeDealId(v)` 呼叫；其他入口如 Hotline 轉商機直接用 `normalizeDealId`，不要另寫正規式）。`POST /api/opportunities` 與表單式 `PUT`（body 帶 `product` 或 `dealId`）：商品要求時缺值 → 400 `DEAL_ID_REQUIRED`；有填但不是 9 碼數字 → 400 `DEAL_ID_FORMAT`（非字串／數字型別 → `DEAL_ID_INVALID`）。**舊資料 grandfather**：請求送的 dealId 與商機現有值完全相同就放行不重驗（`server.js` `checkOppDealId` 第 4 參數），不追溯修改；看板拖曳只送 `{stage}` 不檢查；Excel 匯入不擋（只回 `missingDealId`／`invalidDealId` 統計，格式不合的值捨棄），匯入認「Deal ID」與「SAP Deal ID」兩種標題、匯出標題為「SAP Deal ID」。規則在 `lib/productCatalog.js`（`resolveDealIdRule` 以名稱完全相符或「名稱＋空白」最長前綴比對），前端 `app.js` 的 `dealId*` 函式與之同規則（輸入即時濾非數字、貼上正規化、未滿 9 碼行內錯誤）。檢查腳本：`scripts/check-opp-deal-id.js`（`DEALID_MUT=1` 跑變異）。
- **企業主檔責任業務＋待補名片待辦**（2026-10，分支 feat/company-responsible-sales）：`company.responsibleSales`（帳號陣列）／`responsibleAssignedAt`（{帳號:ISO}）是儲存欄位；企業主檔列表的「服務業務」仍是名片擁有者推導（兩者不同，別混）。待辦**不另存**，由 `lib/companyCardTasks.js` 即時推導（責任業務名下對該公司的真實名片為零＝有待辦；placeholder／已刪除不算；補名片自動消失），只儲存手動解除紀錄 `data.companyCardDismissals`（綁定當時的指定時間，行銷重新指定＝指定時間更新＝解除自動失效）。API：`/api/my-tasks/company-cards`（＋`/dismiss`）、`/api/marketing/responsibility-feedback`（＋`/:id/resolve`）、`/api/manager/company-card-overdue`、`/api/companies/:id/set-responsible-sales`、`/api/companies/responsible-sales-batch`、`/api/company-responsible/candidates`。環境變數 `COMPANY_CARD_TASKS=off` 整個關閉（輸出與沒有此功能時逐位元相同）。前端在 `_client/company-card-tasks.js`（全部 textContent、樣式在 style.css 的 `.cct-*`）。檢查腳本：`scripts/check-company-card-tasks.js`（`CARDTASK_MUT=1` 跑變異、`CARDTASK_ONLY=ui` 跑瀏覽器）。
- **商機編輯儲存會清空商品的舊 bug**（2026-10，分支 fix/opp-edit-preserve-product）：編輯視窗的商品下拉選不到原值時（類別為空——API／Excel 匯入／Lead 轉換建的舊商機；類別在目錄沒有商品清單——MDM／CRM／資安；商品已改名／停用且該類別沒有「其他」），舊版儲存會把 `product` 送成 `''`，後端 `PUT /api/opportunities/:id` 照單全收 → 商品被清空。修在前端、**後端語意刻意不改**（不要加「忽略空字串」，會改變 API 行為）：`app.js` 的 `_oppEditOrig` 記下開啟時的原類別／原商品，使用者沒動過商品欄（`productTouched`）且下拉值為空 → payload 不送 `product`（類別同理）；下拉可顯示時補「（原值）XXX」選項；Deal ID 規則經 `oppEditProductEffective()` 看原商品；改類別 ERP→ITS 仍照舊重選、改回原類別會還原原商品；使用者明確選回「-- 請選擇 --」仍會清空。檢查腳本：`scripts/check-opp-product-preserve.js`（`OPPPRES_MUT=1` 跑變異，M0＝修正前的 app.js 必須變紅）。
- **統編查無→申請行銷建檔**（2026-10，分支 feat/company-requests）：名片表單統編在內部主檔與 GCIS 都查無時，業務填類別（政府機關／學校／公立醫院／財團法人／協會團體／行號／其他）＋說明（5–200 字）存檔，`POST/PUT /api/contacts` 帶 `companyRequest{category,note}` 才會建立名片＋申請（沒帶＝與原本相同；同統編已有待處理申請就併單、記錄多位申請人）。**申請待處理期間刻意不建企業主檔**（否則會以業務填的子單位名稱先佔位），行銷核可才用 `ensureCompanyMaster` 建（產業只由核可者填，沿用 industry cascade）。資料：`data.companyRequests[]`；名片 `companyVerify`（pending／rejected，核可後清除）＋`companyRequestId`＋`companyVerifyReason`。待確認／被退回只擋「報價單送簽」（`lib/quoteRoutes.js` submit，以 q.contactId→商機聯絡人→同擁有者同公司名對應名片，找不到就不擋）與「推送 SAP」（`integrations/push/batch`），其餘照常。端點：`GET /api/company-requests`（行銷載入時順便發逾 3 天提醒，同筆每天一次；逾 7 天標紅）、`POST …/:id/approve|reject|cancel`。稽核 `COMPANY_REQUEST_CREATE／APPROVE／REJECT／CANCEL`；通知 `company_request_new`／`company_request_done`。核可後主檔已存在，業務手動輸入子單位名稱仍會被 `TaxIdGuard.checkSave` 擋（前端會自動帶官方名稱）。環境變數 `COMPANY_REQUESTS=off` 整個關閉（端點 404、輸出與沒有此功能時逐位元相同）。邏輯 `lib/companyRequests.js`、前端 `_client/company-requests.js`（全部 textContent、樣式 `.creq-*`）。檢查腳本：`scripts/check-company-requests.js`（`CREQ_MUT=1` 跑變異、`CREQ_ONLY=ui` 跑瀏覽器）。

- **SEC-02：名片圖下載授權**（2026-11，分支 fix/sec02-upload-authz；沒有環境變數開關）：`GET /uploads/:key`（server.js）原本只要登入就能下載任何 key。現在 `canViewUpload`：key 必須是 `<uuid>.<副檔名>` 格式 → 找出 `contacts[].cardImage` 引用它的聯絡人（**目前經 /uploads 提供的只有「既有的名片舊圖」，新的上傳入口已停用（見下方「名片圖不再儲存」）；沒有任何「全員共用」檔案**，日後新增別種經 /uploads 提供的檔案必須在 `canViewUpload` 明確加授權規則，否則一律 404）→ 沒人引用＝404（含上傳了但還沒存成名片、連管理員也一樣）→ `role==='admin'` 看全部「有被引用」的（含已軟刪除、客戶池）→ 其他人用 `visibleContactsOf(req,data,auth,refs)`，**與 `GET /api/contacts` 共用同一個函式**（從原本 GET /api/contacts 抽出：getViewableOwners＋filterByBu＋filterByViewGroup＋KA 跨部門共享；秘書看不到任何名片圖）。看不到一律 **404 不是 403**（不洩漏檔案存在）。未登入回 **401**（以前是 302 導向登入頁；路由因此放在 `app.use(requireAuth, express.static…)` 之前，不可挪回後面）。postgres 模式加 `dbReadyMiddleware`（/uploads 不在 /api/ 底下，不會經過全域那一個）。**副作用**：沒有名片引用的檔案（孤兒檔）在 /uploads 一律 404、無法下載。不改儲存後端。檢查腳本：`scripts/check-upload-authz.js`（單獨執行，不需瀏覽器；`SEC02_PORT` 3710–3719 預設 3710、`SEC02_ROOT` 可指向別的程式樹跑基準；113 項，含「各角色清單看得到＝下載得到」一致性比對；第 5 節因上傳入口停用，改為直接在儲存目錄放檔＋資料播種 `cardImage`（授權斷言未削弱）；第 6 節補 filterByBu／filterByViewGroup／秘書分支覆蓋（子樹內 BU 不符、集團外公司、秘書名下名片）；在修正前的程式上有 34 項變紅）。已知限制：`GET /uploads` 每次請求要 `db.load()` 一次並掃 contacts（只掃引用該 key 的少數幾筆做可視判斷）；Postgres 模式（Vercel）本機無法驗證，需在 DEV 用兩個不同業務帳號實測。

- **名片圖不再儲存於伺服器**（2026-11，分支 feat/no-card-storage，疊在 fix/sec02-upload-authz 之上；沒有環境變數開關）：決定＝名片圖不存伺服器（空間問題），名片只用 AI 辨識取欄位後即丟棄。①`POST /api/upload` 保留路由與 `requireAuth`，**不掛 multer、不解析 multipart、不寫檔**，一律回 **410**（訊息「名片圖片已不再儲存於伺服器；請使用「AI 辨識」讀取名片資料」，code `CARD_UPLOAD_DISABLED`），舊分頁／舊快取的前端送來也不會 500；`server.js` 的 multer `upload` 設定已移除（`storage` 模組仍留著供 `/uploads/:key` 讀取與刪名片時刪檔）。②`cardImage` 寫入規則：`POST /api/contacts` 一律存 `''`（忽略請求值、不報錯）；`PUT /api/contacts/:id` 請求值與現有值相同→放行不變（舊前端會原樣送回，連舊資料的不合規值都不擋）、非空且不同→忽略（維持原值、不報錯）、`''`／null→清掉（讓使用者移除舊圖，**不刪實體檔**）、沒送→不變；其他路徑（匯入、合併、客戶池指派、拜訪／Hotline 建名片等）本來就只寫 `''`，沒有任何地方會產生新的儲存或把任意值寫進 `cardImage`。③前端：名片表單移除上傳區（`cardImageInput`／`_cardPreviewBlob` 等），只留「AI 辨識」（`ocrCardBtn`／`setupOcrCard`，POST `/api/ai/ocr-card` 用 memoryStorage，從不存圖）；**既有舊圖只讀顯示**（列表縮圖、檢視視窗，以及編輯表單內一塊 `#cardImageLegacy`：有舊圖才出現，含預覽＋「移除圖片」鈕，移除後儲存 payload 帶 `cardImage:''` 即清掉；沒動就帶回原值）；`sw.js` v45。④清除既有舊圖與儲存桶孤兒檔是**另一個步驟**（需使用者確認後另行處理），本功能不刪任何既有檔案或資料；清完後 /uploads 就沒有可下載的檔案，`canViewUpload` 全部 404。⑤`scripts/check-xss-sinks.js` 的 cardImage 矩陣因此改為「任何請求值都不會被儲存」；檢查腳本 `scripts/check-no-card-storage.js`（`NOSTORE_PORT` 3720–3728 預設 3720、基準樹 `NOSTORE_BASE` 預設 `../wt_sec02`、`NOSTORE_SKIP_DOM=1` 略過瀏覽器段；含 410／不寫檔／POST／PUT 矩陣／SEC-02 授權不變／OCR 路由與基準相同／20+ 帳號 GET /api/contacts 與基準逐位元相同／headless Edge 實測）。

- **SEC-08：整合 URL 的 SSRF 防護**（2026-11，分支 fix/sec08-integration-ssrf）：整合區（`integrations/test`、`sap-inspect`、`push/batch`）原本直接 fetch 管理員設定或請求帶入的網址，被盜的管理員 session 可借伺服器出口探測內網／metadata。現在整合區 13 處外送一律走 `integFetch`（`lib/safeUrl.js` 的 `safeFetch`；稽核錨點 webhook 共用同一份 IP 判斷）：只允許 https、不可帶帳密、連接埠白名單（預設 443；SAP 租戶用別的埠才設 `INTEGRATION_URL_PORTS=443,8443`）、DNS 解析後每個位址都必須是公開位址（含 6to4／Teredo／ISATAP／NAT64 內嵌內網 IPv4）、連線釘在驗證過的 IP（防 DNS rebinding，TLS 仍依原主機名驗證）、轉址手動逐跳重檢（≤3 跳，跨來源拿掉 Authorization）、10 秒逾時、5MB 上限；`test`／`sap-inspect`／`push/batch` 先 `checkIntegrationUrl`，不合格回 400＋中文訊息（不回顯網址）。**新增整合外連一律用 `integFetch`，不要用全域 fetch**（`scripts/check-safe-url.js` 的靜態守門會掃 server.js 所有外送呼叫點，未登記的裸 fetch 會變紅）。**未涵蓋（另案）**：①測試／存檔時已存的密碼或 clientSecret 仍可能被送到「被改成別的公開網址」的 baseUrl／tokenEndpoint（存檔時也不驗網址）；②`POST /api/ai/company-insight` 的使用者自填網址仍是裸 fetch（靜態守門已登記為已知缺口）。本機沙盒 DNS 會把網域解析成保留位址，真實 https 連線需在 DEV 用真實 SAP 網址驗證。檢查腳本：`node scripts/check-safe-url.js`（`SAFEURL_ONLY=unit,static` 不開伺服器；`SAFEURL_MUT=1` 跑 21 個變異；埠 3380–3389）。
- **BUG-03：`pushNotification` 不可與「呼叫端先 load 的快照」混用**（2026-11，分支 fix/bug03-overdue-notify）：本地 JSON 後端 `db.load()` 每次回傳新物件；`checkOverdueReminders` 舊寫法先 load 快照、迴圈中每次 `pushNotification` 又自己 load＋save、最後再 save 舊快照 → 通知全被蓋掉（只剩提醒標記，隔天也不補發）。Postgres 後端 load 回同一份快取物件所以正式站沒有遺失，但寫法脆弱。修正：`pushNotification(to,type,title,body,refId,callerData)` 多選用第 6 參數——有傳＝只改該物件（不 load／不 save／不發 Web Push，回傳通知物件）；沒傳＝原行為。Web Push 抽成 `sendWebPush()`。`checkOverdueReminders` 單次 load → 全寫進同一個 data → 只 save 一次 → `db.flush` → 才發 Web Push。新寫「load → 迴圈通知 → save」時務必傳 data，或先 save 再通知。檢查腳本：`node scripts/check-overdue-notify.js`（不開伺服器；`BUG03_MUT=1` 跑變異，還原修正後必須變紅）。

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

## 升版規則（版本資訊頁）

- 後台「版本資訊」顯示 `v<package.json version> · <提交短碼>`：大版本號手動管理，提交編號由 Vercel 部署時自動注入（`VERCEL_GIT_COMMIT_SHA` 等，見 `lib/gitInfo.js`），稽核時可由短碼對應到 GitHub 某筆提交。
- **合併重要功能 PR 時手動改 `package.json` 的 `version`**（同步 `package-lock.json` 最上層兩處）：新功能升 minor、修正升 patch、不相容大改升 major。目前從 v1.8.0 起算。
