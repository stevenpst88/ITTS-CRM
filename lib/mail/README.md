# lib/mail — 簽核信件派送模組（P1–P3 核心＋伺服器端接線）

> 狀態（2026-10-08）：P1（使用者 email 純邏輯）、P2（寄信核心）、P3（連結與跳板頁純邏輯）已完成並有測試。
> **伺服器端已接線**（見 §6A）：`server.js`（單例、`res.end` 包裝、Cron、後台端點、使用者 email 欄位與批次匯入、`/q/<id>` 跳板路由、`/deep-link.js`）、
> `lib/quoteRoutes.js`（E1–E6 掛接點）、`vercel.json`（每日 Cron）。**用戶端已接線**（見 §6B：`admin.html` 的 email 欄位／寄信後台頁、`login.html` 與 `app.js` 的深層連結導回與登出清除、`index.html`、`sw.js`）。
> P4（Microsoft Graph 真傳輸）尚未實作，`createGraphTransport` 目前永遠回 `NOT_CONFIGURED`——所以任何模式都不會真的寄出。
> 本文件不含密鑰、租戶 ID、真實 email 位址或人名；範例位址一律是 `user1@example.test` 這類合成值（網域白名單的預設值 `itts.com.tw` 與 `config.js` 相同）。
> 測試（`scripts/check-mail-*.js`）的位址網域是保留網域（`*.test`、`example.*`）或預設白名單 `itts.com.tw`（整合測試沿用預設設定），**local part 一律是 `user1`、`test-user-1`、`sales1`、`gm1` 這類合成名稱，不用人名**，不可能對應真人信箱；`check-mail-core.js` 第 0b 章會掃整個 `lib/mail`（含本檔）與測試檔，抓到「人名樣式＋真實網域」的位址就失敗。掃描規則是白名單制：local part 的每一段必須是 `GENERIC` 列舉的占位詞，或符合角色樣式——單字母前綴 `m／g／u／t／c／e／p／s` 只能接數字（`m1`、`s10` 可，`ed`、`tj`、`mo`、`cy` 這類像真人縮寫的不可），`mgr／gm／sec／cons／proxy／chair／sales／prx／lead` 後面可接數字或「數字＋一個字母」（`mgr1a` 可，`mgrx` 不可）；數字不拿來洗白（`ed1` 仍被抓）。

## 1. 模組分工與資料流

```
簽核路由 handler（同步）
  ① 改單據狀態 → db.save → writeLog → notify()（站內通知，既有，完全不動）
  ② 整合膠水（新）：事件 = 單據 + approval.derived + 目前關卡；收件人 = [{username, kind}]
        │
        ▼  dispatcher.dispatch(event, recipients, {actorUsername, operatorLabel})      永不 throw
   events.validateEvent ──不合法→ 記 failed/BAD_EVENT（稽核），結束
        │ mode=off → 每位收件人記 skipped/MODE_OFF，結束（不渲染、不查帳號）
   getUsers() → recipients.resolveRecipients      排除操作者／重複／停用／無 email／格式非法／網域不在白名單
        │ 略過者記 skipped（值得注意的原因寫 QUOTE_MAIL_SKIPPED 稽核）
   outbox.enqueue（dedupeKey 冪等；重複觸發＝DUPLICATE，不重寄）
        │ 熔斷開著／超過時間預算 → 留 pending，等 drainDue
   outbox.claim（租約）→ isStillValid(job, ev)（只有回傳 true 才寄）
        → render.renderMail（依收件人 kind 過濾內容，含 visibility 表）→ transports.validateMessage
        → transports.selectTransport(config)（模式護欄）→ transport.send（總逾時 8 秒、AbortSignal）
        │ 成功 → markSent；失敗 → markFailed（退避 60/300/900 秒或 Retry-After）＋ breaker.record
        │ 最終失敗 → 寫 QUOTE_MAIL_FAILED 稽核（只有型別、帳號、錯誤碼）
        ▼
   Summary = {queued, sent, skipped:[{username,reason}], failed:[{username,code}], cancelled, errors}
             drainDue 另有（只在發生時才有這些欄位）breakerOpen:true、budgetExhausted:true

  drainDue({limit, budgetMs}) = 先清掃「租約過期且本輪次數用盡」的 sending（→ failed/LEASE_EXPIRED，列入 Summary.failed 並寫稽核）
                              → 每個工作者「預算還有、熔斷沒開」才領一筆（claimDue limit 1）→ rebuild(job) 重建事件 → 同上的檢查與寄送
                              （機會式清理、每日 Cron、後台重送都走它）。到預算或熔斷中途打開就停止領取，沒領的工作原封不動留在 pending，不扣次數
  ③ server.js 的 res.end 包裝：先 quoteMail.waitPending(req)（等這個請求累積的寄信 promise，**最多 12 秒**；每封寄信 promise 本身另有 10 秒整體時限），再 await db.flush()
```

| 檔案 | 責任 | 主要匯出 |
|---|---|---|
| `config.js` | 環境變數 → 設定（安全預設：mode=off） | `getMailConfig`、`publicConfig`、`parseMode`、`normalizeMode` |
| `safety.js` | email 驗證與網域白名單、遮罩、標頭／HTML 清理 | `normalizeEmail`、`isAllowedDomain`、`maskEmail`、`headerSafe`、`escHtml`、`safeText`、`clip` |
| `visibility.js` | 「誰看得到什麼」的唯一事實來源 | `KINDS`、`visibilityFor` |
| `events.js` | 事件定義、驗證、去重鍵 | `EVENT_TYPES`、`validateEvent`、`dedupeKey`、`isValidQuoteId` |
| `recipients.js` | username → 可寄的真實收件人 | `resolveRecipients` |
| `templates.js` / `render.js` | 事件 × 收件人 kind → 主旨／HTML／純文字 | `renderMail`、`ALLOWED_KINDS`、`formatNtd`、`MailRenderError` |
| `outbox.js` / `outboxAdapters.js` / `scrub.js` | 寄件匣：冪等入列、租約領取、退避、熔斷、清理；JSON／記憶體／Postgres 三種儲存 | `createOutbox`、`jsonFileAdapter`、`memoryAdapter`、`postgresAdapter` |
| `transports.js` | 傳輸介面、模式護欄 | `selectTransport`、`createLogTransport`、`createRedirectTransport`、`createGraphTransport`（P4 stub）、`validateMessage` |
| `dispatcher.js` | 把以上串成「永不 throw」的派送流程 | `createDispatcher` → `{dispatch, drainDue}` |
| `userEmail.js` | 使用者 email 驗證、批次匯入預覽／套用、缺 email 清單 | `validateUserEmail`、`parseBulkEmailText`、`applyBulkPlan`、`missingEmailReport`、`maskedAuditDetail` |
| `link.js`（+ `_client/deep-link.js`） | 信內連結、`/q/:id` 跳板頁、登入前暫存深層連結 | `buildQuoteLink`、`renderJumpPage`、`ITTSDeepLink.remember/consume` |
| `validity.js` | **過期判斷（純函式）**：這封待寄／待重寄的信在單據現況下還該不該寄（E1～E6 各有條件，見 §7）；E1 的 stepKey 與改派標記 | `checkValidity`、`stepKeyOf`、`stepIdxOf`、`reassignEpoch`、`e1StepKey`、`VALIDITY_CODES` |
| `quoteMail.js` | **整合膠水**：`quoteRoutes.js` 的簽核事件 → 事件＋收件人 → `dispatcher`；`isStillValid`／`rebuild`；機會式清理；寄信流程的整體時限 | `createQuoteMail` → `{bind, notifyMail, waitPending, snapApproval, pollDrain, drainDue, isStillValid, checkJob, rebuild, missingEmails, emailStatusMap, diagnostics}` |
| `routes.js` | Cron 端點、後台寄件匣／設定頁、使用者 email 批次匯入（不 require npm 套件） | `registerMailRoutes` |

事件 × 收件人 kind 矩陣（不在表內的組合 `renderMail` 會 throw `KIND_NOT_ALLOWED`，dispatcher 記為 `failed/RENDER`）：

| | mgr1 | gm | chairman | secretary | boardProxy | consultant | owner |
|---|---|---|---|---|---|---|---|
| E1 送簽 | ◎ | ◎ | ◎ | ◎ | ◎ | | |
| E2 請填成本 | | | | | | ◎ | |
| E3 下一關 | ◎ | ◎ | ◎ | ◎ | ◎ | | |
| E4 結果→業務 | | | | | | | ◎ |
| E5 成本完成→業務 | | | | | | | ◎ |
| E6 撤回／作廢 | ◎ | ◎ | ◎ | ◎ | ◎ | | ◎ |

可見性（`visibility.js`；信件內容一律查這張表，樣板內沒有任何角色判斷）：決策條（金額「折扣後未稅」、毛利率色塊、核決層級）只在 E1／E3，且只給 mgr1／gm／chairman／secretary／boardProxy；
secretary／boardProxy 的信**不帶客戶名、業務名與專案名稱**（`project` 旗標為 false：專案名稱常含客戶名；主旨、preheader、本文、純文字都只放報價單號，例如「【簽核通知】QU-2026-0123」；E1／E3／E6 對這兩種 kind 一致；業主 2026-10-08 決定）；
mgr1／gm／chairman／consultant／owner 的信照舊含專案名稱（`project: true`）；未知 kind 一律全 false。consultant 的信沒有任何金額、毛利、折扣、單價（只有品項說明／數量／單位）；owner 的信沒有金額與毛利。金額與毛利率不進主旨；連結不帶 token。
專案名稱只從事件的 `projectName` 一個欄位進信，且只在 `render.buildModel`（rows、title、preheader）與 `subjectFor(ev, kind)` 兩處讀取、都查 `visibilityFor(kind).project`；駁回原因（`result.reason`，業務可輸入的自由文字）是獨立欄位，不會被當成專案名稱的來源，而且 E4（含駁回原因）只寄給業務本人、E6 事件本身不帶 reason。

## 2. 最小用法（本機，JSON 檔案 outbox）

```js
const { getMailConfig } = require('./lib/mail/config');
const { createOutbox, jsonFileAdapter } = require('./lib/mail/outbox');
const { createDispatcher } = require('./lib/mail/dispatcher');

const config = getMailConfig();                                   // 預設 mode=off：什麼都不會寄
const outbox = createOutbox(jsonFileAdapter({ file: config.outboxFile }), { config });
const mail = createDispatcher({
  config, outbox,
  transport: undefined,                                           // P4 才有 Graph；測試時傳假傳輸 {kind, async send(msg, {signal, timeouts})}
  getUsers: () => ctx.users,                                      // 陣列（auth.json 形態）或 username 為鍵的物件（quoteRoutes 的 ctx.users）都可以
  isStillValid: (job, ev) => true,                                // 見 §6：只有回傳 true 才會寄
  rebuild: (job) => null,                                         // 見 §6：drainDue 用
  writeLog: (action, operatorLabel, quoteNo, detail) => {},       // 接 server.js 的 writeLog（多一個 req 參數，要包一層）
});

const summary = await mail.dispatch(event, [{ username: 'user1', kind: 'mgr1' }], { actorUsername: 'sales1', operatorLabel: 'Sales One' });
const drained = await mail.drainDue({ limit: 3 });                // 領取到期工作重寄；可加 budgetMs（預設 min(20 秒, 2×單封逾時)＝16 秒，到點停止領取）
```

預覽信件外觀（不連網、不寄信）：`node scripts/mail-preview.js` → `.mail-preview/gallery/index.html`（已在 `.gitignore`）。

## 3. 環境變數（全部可省略；值絕不可寫進 repo）

| 變數 | 預設 | 說明 |
|---|---|---|
| `MAIL_MODE` | `off` | `off`｜`log`｜`redirect`｜`live`。前後空白與大小寫會先正規化（`LIVE ` 等同 live）；**未設、非法、`true`／`1`／`yes`、全形字一律當 off**。只有明確的 `live` 才可能真寄。 |
| `MAIL_REDIRECT_TO` | 空 | redirect 模式的**單一**測試信箱（需通過 `normalizeEmail`，不受網域白名單限制）。`MAIL_MODE=redirect` 但缺／非法 → 有效模式降為 `log`（不會退回寄給原收件人）。Demo／本機專用，值只放該環境的環境變數。 |
| `MAIL_FROM_NAME` | `ITTS-CRM 簽核通知` | 寄件顯示名稱（去換行與 `<>"\`，≤60 字）。目前傳輸層還沒用到，P4 的 Graph 才會用。 |
| `APP_BASE_URL` | 次選 `https://`＋`VERCEL_PROJECT_PRODUCTION_URL`，最後是正式站網址 | 信內連結的網站根（只接受 origin：https；`localhost`／`127.0.0.1` 可 http；不可帶帳密／路徑／查詢）。不合法 → 預設值＋warning。 |
| `MAIL_ALLOWED_DOMAINS` | `itts.com.tw` | 收件人網域白名單（逗號分隔，**精確相等**，子網域與後綴不算）。寄送當下與帳號輸入時各驗一次。 |
| `MAIL_GRAPH_TENANT_ID`／`_CLIENT_ID`／`_CLIENT_SECRET`／`_SENDER` | 空 | P4（Graph sendMail）才使用，目前只讀取不使用；四項齊全且合法才 `configured`。`CLIENT_SECRET` 是機敏值：只放 Vercel 環境變數，`config.graph` 的這四個欄位都不可列舉，序列化／`util.inspect`／`publicConfig` 都不會輸出。 |
| `MAIL_PREVIEW_DIR` | `.mail-preview` | log 傳輸把信件寫成檔案的目錄（相對 repo 根；不可含 `..`）。**Vercel（`VERCEL` 有值）一律關閉**（唯讀檔案系統，也不把信件內容印到 console）。 |
| `MAIL_OUTBOX_FILE` | `mail-outbox.json` | 本機 JSON outbox 的檔名（`jsonFileAdapter({file: config.outboxFile})` 由接線者傳入）。**Vercel 上不可用**：JSON／記憶體 adapter 在 `VERCEL` 有值時會直接 throw `NOT_FOR_VERCEL`，正式環境一律用 `postgresAdapter`。 |
| `VERCEL` | Vercel 自動注入 | 見上面兩項。 |
| `CRON_SECRET` | 空（未設＝Cron 端點一律 401） | 保護 `/api/cron/mail-outbox`：`routes.js` 在每個請求讀取，驗 `Authorization: Bearer <值>`（`timingSafeEqual`）。Vercel 在設定了這個變數時會自動帶上此標頭（以整合當下的官方文件為準）。值只放環境變數，不寫進 repo。 |
| `DATABASE_URL` | 既有 | `postgresAdapter` 沿用 `db/postgres.js` 的連線慣例（max 2、ssl、connectionTimeoutMillis 5000）。 |

固定值（`getMailConfig` 內，未開放環境變數）：逾時 connect 3 秒／total 8 秒；退避 60／300／900 秒；熔斷「10 分鐘內累計 5 點失敗 → 開 10 分鐘」；保留 90 天。

**四種模式的行為**（整合測試 `scripts/check-mail-integration.js` 第 4 章逐項驗證）

| 模式 | 傳輸 | outbox 記錄 | 備註 |
|---|---|---|---|
| `off` | 不呼叫 | 每位收件人 `skipped/MODE_OFF` | 不渲染、不查帳號、不寫稽核。可由後台 `requeue` 在之後重送。 |
| `log` | 只寫預覽檔（本機）；Vercel 上什麼都不輸出 | `skipped/MODE_LOG` | 即使傳入真傳輸也不碰。 |
| `redirect` | 內層傳輸（真傳輸或 log）；收件人**一律**改寫成 `MAIL_REDIRECT_TO` | 內層是真傳輸 → 正常 `sent`；內層只是 log（P4 前）→ `skipped/MODE_LOG`。`toMasked` 都是預定收件人的遮罩位址 | 主旨加 `[測試轉送]`；HTML／純文字最上方加註「原收件人：s***@…」。缺目標 → `NOT_CONFIGURED`，不退回寄給原收件人。 |
| `live` | 真傳輸（P4 前＝Graph stub，回 `NOT_CONFIGURED`） | 正常 | 只有 `MAIL_MODE` 正規化後恰為 `live` 才走到這裡。 |

> P4 之前，Demo 的 `redirect`／`live` 都**不會真的寄出**（內層是 log，在 Vercel 上不輸出任何東西；live 是 stub）。
> 要看信件外觀請用本機 `node scripts/mail-preview.js`；Demo 在 P4 前只能驗證「事件有沒有觸發、outbox 有沒有紀錄」。

## 4. 寄件匣（outbox）、重試與熔斷

- 紀錄只存引用與狀態：`id, type, quoteId, quoteNo, toUser, toMasked, dedupeKey, status, attempts, attemptsInRound, nextAttemptAt, leaseUntil, lastErrorCode, lastErrorMsg(≤200,已清理), skipReason, createdAt, updatedAt, sentAt, actorLabel, meta:{level, kind}`。
  **不存**信件內容、金額、毛利、完整 email。`actorLabel` 是操作者顯示名稱（既有稽核慣例）；除此之外沒有人名。
- 狀態：`pending → sending → sent`；`failed`（最終失敗，等後台處理）、`skipped`（沒寄：MODE_OFF／MODE_LOG／NO_EMAIL…）、`cancelled`（STALE／GONE）。
- 冪等：`dedupeKey = type:quoteId:username:stepKey`。同一事件重複觸發 → `DUPLICATE`，不重寄。重新送簽、下一關是新的 `stepKey`，會再寄。管理員改派一級主管也是新的 E1：`stepKey` 帶改派時間（`<submittedAt>#0@<改派時間>`，取自簽核歷史最近一筆 `REASSIGN`），所以 A→B→A 時 A 會再收到一封；同一次改派被重複觸發仍是同一把鍵（去重）。沒有改派過的送簽維持 `<submittedAt>#0`。**只有承辦人真的改變才算新的改派**：管理員把一級主管「改派」給目前的承辦人本人（A→A，含連續按）不產生新的鍵、不多寄；`POST /reassign` 把改派前後的承辦人帳號記在該筆 `REASSIGN` 歷史的 `meta:{from,to}`，`reassignEpoch` 略過 `from===to` 的紀錄（A 還在重試中的舊 E1 因此也不會被當成過期取消，不會漏信）。沒有 `meta` 的舊紀錄無法辨識，維持原本「每次改派都是新的 E1」的行為。
- 重試：一輪 4 次嘗試（首次＋3 次重試；`retry.maxAttempts=3` 解讀為「重試次數」），退避 60→300→900 秒；Graph 429 的 `Retry-After` 優先。permanent 錯誤（AUTH／REJECTED／BAD_MESSAGE／NOT_CONFIGURED／RENDER）直接 `failed`。
  最終失敗後**不會自動再試**，由後台 `outbox.requeue(id)`（再給一輪；`attempts` 累計不歸零）。
- 租約：領取時設 60 秒租約，工作者當掉後租約到期會被接手；舊工作者的回報用 `attempts` 樂觀鎖擋掉（`LOST_LEASE`）。語意是「至少一次」：寄出成功但 `markSent` 沒寫到時（或上一個工作者其實已寄出、只是沒來得及回報），租約過期後**可能重寄一封**。
  若租約過期時本輪次數（4 次）已用盡，工作改標 `failed`／`LEASE_EXPIRED`（避免無限重寄）；這不是靜默轉換：`drainDue` 會先清掃（`outbox.expireStale()`），把它列入 `Summary.failed`（`code:'LEASE_EXPIRED'`）並寫 `QUOTE_MAIL_FAILED`（`code=LEASE_EXPIRED`），寄信紀錄頁與 `stats().failed24h` 也看得到，可用 `requeue` 再給一輪。所以**不會靜默丟失**，代價是極端 race 下的重複一封。
- 預算：`drainDue` 不會一次領走 `limit` 筆。每個工作者「預算還有（預設 16 秒）、熔斷沒開」才領**一筆**，處理完再領下一筆；到預算或熔斷在批次中途打開就停止領取（Summary 帶 `budgetExhausted:true`／`breakerOpen:true`），沒輪到的工作原封不動留在 `pending`（沒被扣次數、沒佔租約）。已開始的寄送會做完（各自受單封總逾時 8 秒約束），所以最壞耗時 ≈ 預算 + 8 秒 ≈ 24 秒，低於 `vercel.json` 的 `maxDuration` 30 秒。
- 熔斷：TIMEOUT／NETWORK／SERVER 各 1 點、THROTTLED 3 點、AUTH 直接開；窗口內累計 5 點開 10 分鐘；開著時新事件只入列（`nextAttemptAt＝熔斷結束`），`drainDue` 回 `breakerOpen:true`（開始前開著就直接返回；批次中途才打開，則寄完手上那幾封就停止領取）。結束後「半開」：下一次失敗立刻重開，成功完全重置。REJECTED／BAD_MESSAGE／NOT_CONFIGURED／RENDER 不計入。
- 稽核：只在「最終失敗」與「值得注意的略過」寫 `QUOTE_MAIL_FAILED`／`QUOTE_MAIL_SKIPPED`，`detail` 固定格式 `type=E3_NEXT_STEP to=<帳號> code=… | reason=…`，不含位址、金額、專案／客戶名。成功、MODE_OFF／MODE_LOG、重複、操作者本人都不寫。

## 5. Hobby 方案的重試設計

Vercel Hobby：沒有常駐程序；Cron **每天最多 1 次、時間精度約 ±59 分鐘**；函式最長 300 秒但本專案 `vercel.json` 設 30 秒；沒有引入 `@vercel/functions`（所以沒有 `waitUntil`，寄信一定要在請求內等完）。
（Hobby 條款僅限非商業用途——請專案負責人另行確認授權是否適用，程式不處理。）因此重試分四層，由快到慢：

| 層 | 觸發 | 頻率／上限 | 能做什麼 | 做不到什麼 |
|---|---|---|---|---|
| ① 當次嘗試 | 簽核路由內 `dispatch` | 每封 1 次，總逾時 8 秒、單次 dispatch 預算 15 秒、同時 3 封 | 正常情況下信在回應前就寄出；失敗留在 outbox | 傳輸卡住時最多讓這個請求多等 8 秒（熔斷開啟後不再嘗試，不吃逾時）；outbox／資料庫操作卡住時，整個寄信流程另有 10 秒整體時限（`notifyMs`），超時放手、業務回應照送，稽核 `QUOTE_MAIL_FAILED code=DEADLINE` |
| ② 機會式清理 | 後續請求順手 `drainDue({limit:2~3, budgetMs:3000})` | 每個實例至少間隔 60 秒一次（自己加節流） | 有流量時 60/300/900 秒的退避會被準時領取 | 沒有人用系統時什麼都不會跑 |
| ③ 每日 Cron | `vercel.json` 的 `crons`，每天 1 次打 `/api/cron/mail-outbox` | `drainDue({limit:20, budgetMs:16000})`＋`purge()` | 整晚沒流量也能在隔天早上補寄；清 90 天前的舊紀錄 | 一天只有一次嘗試；時間不精準 |
| ④ 後台重送 | 管理員在寄信紀錄頁按「立即重送」 | 人工 | 最終失敗／skipped（補了 email）的紀錄手動再給一輪 | — |

整合測試第 8 章用假時鐘模擬了這四層（含「連續 4 天每天一次 Cron 都失敗 → 最終失敗 → 後台重送」）。
需要更密的重試時：升級 Vercel 方案（Pro 才有每分鐘 Cron），或讓外部排程（例如本機 n8n）定時打同一個受保護的端點——端點是冪等的，租約保證同一時間只有一個工作者處理同一筆（語意仍是至少一次，見 §4「租約」）。

## 6. 整合階段待辦清單（原文；伺服器端項目已完成，見下面 §6A）

> 行號是 2026-10-08 工作樹的概數（另一個 session 正在改這些檔案），**以函式／路由名稱為準**。順序＝建議的實作順序。每一步完成後都先在 **Demo** 驗證（獨立 repo／Supabase／Vercel），正式站 `MAIL_MODE` 維持 `off`。

### 0. 先決條件
1. 另一個 session 未 commit 的修改（`lib/quoteRoutes.js`、`lib/quoteApproval.js`、`server.js`、`_client/app.js`、`_client/index.html`、`_client/admin.html`、`_client/quote*.js`）要先 commit 或協調，本案會疊在同一批檔案上。
2. commit 與 push 各自需要專案負責人明確指示；push main＝正式部署。
3. 三套帳號資料（本機 `auth.json`、正式 `_auth`、Demo `_auth`）各自維護 email，不會同步；目前都沒有 `email` 欄位。

### 1. P1：使用者 email（`server.js`、`admin.html`、`quote-approval.js`）
| 位置 | 改什麼 | 注意 |
|---|---|---|
| `server.js` `GET /api/admin/users`（白名單式回欄位） | 加 `email` | 只有 admin 端點可回；**不要**放進 `/api/me`、JWT payload、`/api/me/contact`（那是印在客戶報價單上的聯絡資訊） |
| `POST /api/admin/users`、`PUT /api/admin/users/:username` | 解構 `email`；`'email' in body` 才呼叫 `validateUserEmail(raw, {config, users, selfUsername})`；空字串／null＝清除；稽核 detail 追加 `maskedAuditDetail(old, new)` | `undefined` 會回 `MISSING`（刻意：路由漏帶欄位時不能無聲清掉）；`contacts[].email` 是客戶聯絡人，別混用 |
| 新增 `POST /api/admin/users/email-import/preview` 與 `/apply` | `parseBulkEmailText` → 預覽；apply 時重新解析再 `applyBulkPlan(rows, users, {config})`，逐人寫回 `saveAuth` | 上限 500 行／20 萬字元（超過整批拒絕）；**`updated[].previous`／`.email` 是完整位址，只能回給前端的預覽，不可寫稽核**，稽核用 `updated[].detail` |
| `admin.html` | 帳號 Modal 加 email 欄位；列表加「無 email」徽章；批次匯入對話框（預覽→確認）；缺 email 清單 | rows 的 `username`／`error` 可能含 `<>`，一律 `textContent`／跳脫（參考既有的後台 XSS 跳脫教訓） |
| `GET /api/quote-approval/config` 的 `users` ＋ `quote-approval.js` 名冊頁 | 每人加 `emailStatus`（`missingEmailReport({users, roster, config})` 算出的 `reason` 或 `OK`；**不回傳位址**）；名冊頁顯示徽章 | `missingEmailReport` 的角色對照：一級主管＝role `manager1`（列出全體，實際簽核人是沿 supervisor 找到的那位）、gm／chairman／boardProxy／costProviders＝名冊、秘書＝role `secretary` |

### 2. P2：派送層接線
1. **新增 `lib/quoteMail.js`（膠水）**：內容就是 `scripts/check-mail-integration.js` 開頭「整合膠水 stand-in」那一段（形狀已被整合測試驗證）。需要的函式：
   - `stepRecipients`／`findManager1`／`boardProxySet`：**直接重用 quoteRoutes.js 裡的同名函式**（不要複製第二份）；`kindOf(tier, user)`：mgr1→`mgr1`、gm→`gm`、chairman→`chairman`、board→（role secretary ? `secretary` : `boardProxy`）、業務→`owner`、顧問→`consultant`。
   - `buildEvent(q, type, …)`：見下方「事件欄位對照」。
   - `rebuild(job)`、`isStillValid(job, ev)`：見 §7。
2. **`quoteRoutes.js` 掛接點**：在每個 `notify(...)` 之後加一行 `mail.notifyMail(ctx, q, spec)`（**不要改 `notify()` 本身**，站內通知行為完全不變）。對照：

   | 現有 `notify` type（路由） | 事件 | 收件人與 kind | stepKey |
   |---|---|---|---|
   | `quote_cost_request`（新增單／修改單選顧問） | E2 | `[q.costBy]` consultant | `costFlow.requestedAt#costBy` |
   | `quote_cost_done`（顧問按完成） | E5 | `[q.owner]` owner | `costFlow.filledAt#done` |
   | `quote_submitted`（送簽；改派一級主管） | E1 | 第 0 關 `stepRecipients` mgr1 | `approval.submittedAt#0` |
   | `quote_submitted`／`quote_board`（核准後的下一關） | E3 | 下一關 `stepRecipients`（gm／chairman／board） | `approval.submittedAt#<關卡序號>` |
   | `quote_step_approved`／`quote_approved` | E4（`approved`／`final_approved`） | `[q.owner]` owner | `submittedAt#r:<kind>:<關卡序號>` |
   | `quote_returned`（駁回） | E4（`rejected`，原因放 `result.reason`） | `[q.owner]` owner | 同上 |
   | `quote_withdrawn` | E6（`withdrawn`） | 目前關卡收件人＋已簽過的人 | `submittedAt#withdrawn` |
   | `quote_voided` | E6（`voided`） | 已簽過的人＋業務 | `submittedAt#voided` |

   事件欄位對照（**金額一律是整數分**，來源是送簽時凍結的 `approval.derived`）：
   - `numbers = { revenueCents: derived.revenueCents, gpCents: derived.gpCents, marginText: derived.marginText, marginPct: Number(derived.marginText), tierLevel: derived.level, tierLabel: derived.board ? '董事會' : {1:'一級主管',2:'總經理',3:'董事長'}[derived.level] }`；`derived` 為 null（成本未完成）時 `numbers: null`。
   - **`tierLabel` 要用短標籤**。不要直接用 `QA.TIERS.board`（「董事會決議（秘書代核）」），否則色塊會變成「需董事會決議（秘書代核）核准」；`step.label` 才用 `QA.TIERS[tier]` 的完整文字。董事會關請給 `tierLevel: 3`（紅色），給 null 會是灰色。
   - `step = { level: {mgr1:1, gm:2, chairman:3, board:'board'}[tier], label: step.label }`。
   - `items`（E2）：`QI.itemRows(q.items).map(({desc, qty, unit}) => ({desc, qty, unit}))`（`lib/quoteItems.js` 的 `itemRows` 會濾掉標題列與小計列）。只放這三個欄位；就算多放價格，信件層也不會讀。
   - `ownerLabel`／`actor.label` 用 `dispName`（暱稱 > 顯示名稱 > 帳號，各截至 100 字）；`company`、`projectName` 各截至 300 字（`quoteMail.makeEvent` 的 `clip(…, 300)`，與 `events.js` 的 `LIMITS.company`／`LIMITS.projectName` 一致；單據欄位本身在 `quoteRoutes.js` 另有 100／200 字的輸入上限，是更前面的一道，不影響這裡）。
   - `at` 用事件發生時間（ISO）；retry 重建時用 `job.createdAt`，重試信才會和首次逐字相同。
   - 駁回（`/return`）用 `rejected`（「已駁回」）；`returned` 是「退回修改」，目前沒有對應的路由。
   - 傳給 `dispatch` 的 `recipients` 建議用 `notify()` 過濾後的名單（排除操作者、去重、排除非在職）；信件層會再檢查一次，所以傳原始名單也不會出事，只是停用帳號會多一筆 `QUOTE_MAIL_SKIPPED/INACTIVE` 稽核。
   - `writeLog` 要包一層：模組呼叫 `(action, operatorLabel, quoteNo, detail)`，`server.js` 的是 `writeLog(action, operator, target, detail, req)`。
3. **`server.js` 單例與回應結束包裝**
   - 建立單例：`config = getMailConfig()`；adapter＝`DB_BACKEND==='postgres' ? postgresAdapter({query}) : jsonFileAdapter({file: config.outboxFile})`；`createOutbox(adapter, {config})`；`createDispatcher({...})`。
     `query` 由接線者用 `pg.Pool` 包成 `(sql, params) => pool.query(sql, params)`（連線設定沿用 `db/postgres.js`、`lib/apiMonitor.js` 的慣例：max 2、ssl、`connectionTimeoutMillis` 5000）；第一次 outbox 操作會自動 `CREATE TABLE IF NOT EXISTS`（8 條 DDL，含 RLS）。
   - **必須 await**：在 `res.end` 包裝（先 `await db.flush()` 再送回應的那段）裡，一併等 `req._mailPending`（實作：`quoteMail.waitPending(req)`，有時限，見下方「時限」）；`notifyMail` 把 `dispatch` 的 promise 推進 `req._mailPending`。Vercel 回應送出後實例可能被凍結，Web Push 那種「發了就不管」的寫法在這裡不可用。包裝只涵蓋 POST／PUT／DELETE／PATCH，而所有會寄信的路由都是這四種。
   - **寫入順序鐵則**：先 `db.save` 再通知（`quoteRoutes.js` 檔頭註解）；`notifyMail` 排在 `notify()` 之後。寄信失敗絕不可影響送簽／核准本身（`dispatch` 永不 throw，膠水仍要 try/catch 一層）。
   - 回應會多等寄信時間：Graph 約數百毫秒（P4 前未實測）；最壞情況受 8 秒逾時＋熔斷限制。必要時把 `config.timeouts.totalMs` 降到 5 秒。
   - **時限**（FIX-3；`lib/mail/quoteMail.js` 的 `DEFAULT_LIMITS`，測試可由 `deps.limits` 覆蓋）：dispatcher 對傳輸有 8 秒總逾時、對 `getUsers`／`isStillValid`／`render` 有 5 秒輔助逾時，但對 outbox／資料庫操作沒有——所以 `notifyMail` 把整個流程（入列＋當次寄送＋路由尾端清理）包進 `deadline(…, 10 秒)`（Promise.race 的小工具，永遠 resolve、計時器一定清除、被丟下的 promise 之後 reject 也已被接住，不會變成 unhandled rejection）。超時：`console.warn`＋`diagnostics().deadlines.notify` 計數＋稽核 `QUOTE_MAIL_FAILED code=DEADLINE`（只有型別與單號），業務回應照送。被丟下的工作若已入列，之後由租約／清理接手（至少一次）；若連入列都沒完成就沒有紀錄，只留這筆稽核。`waitPending` 是外層保險（12 秒）。`drainDue`（Cron、後台重送）整體時限 = min(22 秒, 預算＋單封逾時＋1 秒)，超時回 `{timedOut:true, errors:1}`（Cron 回應標 `timedOut:true`）；Cron 的 `purge` 與 `db.flush`、`pollDrain` 之後的 `db.flush` 各 3 秒；`pollDrain` 本身仍是 4 秒。Cron 最壞 ≈ 22＋3＋3＝28 秒，低於 `maxDuration` 30 秒。
4. **機會式清理**：(a) 在會寄信的路由尾端、自己的 `dispatch` 之後再 `await mail.drainDue({limit:2, budgetMs:3000})`；(b) `GET /api/poll-bundle`（所有登入中的瀏覽器都會輪詢）加一個「每個實例至少 60 秒才跑一次」的節流後 `await Promise.race([mail.drainDue({limit:2, budgetMs:3000}), 逾時 4 秒])`——該端點若有 ETag／304 短路（盤點時有），要放在短路**之前**，且不可改變 ETag 的計算。`MAIL_MODE=off` 時 `drainDue` 立刻返回，不會碰資料庫。
   GET 沒有 flush 包裝、最壞會多等一個傳輸逾時，所以一定要節流並加逾時；被 race 丟下的工作（已領取、尚未回報）靠 60 秒租約到期後由下一個請求接手。語意是**至少一次**（見 §4「租約」）：極端 race 下（上一個工作者其實已寄出、只是沒來得及回報）可能重寄一封；不會靜默丟失——若連續被丟下到本輪 4 次用盡，會轉成 `failed`／`LEASE_EXPIRED`，並出現在下一次 `drainDue` 的 `Summary.failed` 與 `QUOTE_MAIL_FAILED` 稽核。給 `drainDue` 傳 `budgetMs`（例如 3000）可讓它在預算內自己停手，不必只靠外層逾時。
5. **Cron 端點**：`GET /api/cron/mail-outbox`，放在 `requireAuth` 之外，自己驗 `Authorization: Bearer ${CRON_SECRET}`（比對用 `crypto.timingSafeEqual`；`CRON_SECRET` 未設就一律 401）→ `await mail.drainDue({limit:20, budgetMs:16000})`（`budgetMs` 預設就是 16 秒，寫出來是為了和 `maxDuration` 一起檢查：**預算 + 單封逾時 8 秒 < `vercel.json` 的 `maxDuration` 30 秒**；到點就停止領取、剩下的留給下一次，傳輸卡住（每封都等滿 8 秒）時，兩輪（約 6 封）之後就會因預算或熔斷收手，不再像「一次領 20 筆」那樣被平台中途殺掉）→ `await outbox.purge()` → 回 `{ok, sent, failed, ...}`（不回收件人）。
   `vercel.json` 加 `"crons": [{ "path": "/api/cron/mail-outbox", "schedule": "0 0 * * *" }]`（UTC 00:00＝台北 08:00 起算一小時內）。Hobby 的排程不能比「每天一次」更頻繁（依 Vercel 官方文件，這樣的設定會讓部署失敗；整合時再對照當時的官方文件確認）。所有路徑都被 rewrite 進同一個 Express app，所以不需要新的函式檔。
6. **後台**（`admin.html` 新增「寄信」區段，照既有 `data-sec` lazy-init 模式，用 `adminFetch`）：
   - 寄信紀錄頁：`GET /api/admin/mail/outbox`（`outbox.list({status,type,quoteNo,toUser,since,limit,offset})`＋`stats()`＋`breaker.state()`）；列表以 `toMasked` 顯示，不顯示完整位址。
   - 「立即重送」：`POST /api/admin/mail/outbox/:id/requeue`（`outbox.requeue(id)`）；失敗（`failed`）與 `skipped`（補了 email 之後、MODE_LOG／MODE_OFF 之後）都可以重送。重送時會用單據現況重建內容並重新檢查，所以已過期的舊關卡信會被取消（`STALE`）而不是誤寄。
   - 設定頁（唯讀）：`publicConfig(config)`＋`warnings`；模式與密鑰只由環境變數決定，不在後台改。
   - 缺 email 清單：`missingEmailReport`。「寄測試信」與 Graph 診斷屬於 P4。
   - 所有後台寫入都要 `writeLog`（例如 `REQUEUE_MAIL`）；顯示的字串一律跳脫。

### 3. P3：連結與深層連結
| 位置 | 改什麼 | 注意 |
|---|---|---|
| `server.js` 公開路由區（`requireAuth` 與 `express.static(_client)` 之前，helmet 之後） | `GET /q/:id` → `const r = renderJumpPage({quoteId: req.params.id, cost: req.query.cost}); res.status(r.status).set(r.headers).send(r.body);` | `res.set()` 會取代 helmet 的同名標頭（只會更嚴）。不查資料庫、不洩漏單據是否存在 |
| **`server.js` 公開路由區：`/deep-link.js` 要有自己的公開靜態路由**（仿 `/itts-logo.png` 那幾行） | `app.use('/deep-link.js', express.static(path.join(__dirname, '_client', 'deep-link.js'), STATIC_NO_CACHE))` | **容易漏**：`login.html` 是公開頁，但其他 `_client/*.js` 都在 `app.use(requireAuth, express.static(_client))` 後面；沒加這行，登入頁載入 `deep-link.js` 會被導向登入頁（回 HTML，被 `nosniff` 擋掉），**靜默失效**。或者把 deep-link.js 的內容直接內嵌進 `login.html` |
| `_client/login.html` | 載入時 `ITTSDeepLink.remember(location.hash)`；登入成功導向改為 `(admin ? '/admin.html' : '/' + ITTSDeepLink.hashForRedirect())`（admin 不附片段） | `hashForRedirect` 取一次即清 |
| `_client/app.js` | (a) `initUser` 的 401 分支導向登入頁前 `remember(location.hash)`；(b) 強制改密碼流程（`mustChangePassword`）改完後續跑 `_handleQuoteDeepLink`；(c) `_handleQuoteDeepLink` 的正規式 `^#quote:([0-9a-fA-F-]{8,64})(:cost)?$` 只認 uuid，單據 id 都是 uuid 所以現在一致；要支援其他 id 就改呼叫 `ITTSDeepLink.isValidHash` | 先確認另一個 session 對 app.js 的修改已提交 |
| `_client/index.html` | 加 `<script src="deep-link.js"></script>`（已登入頁面，一般靜態路由即可） | |
| `_client/sw.js` | `isNetworkOnly` 加 `'/q/'`（避免快取跳板頁），並改 `SW_VERSION` | Cache API 不理會 `no-store` |
| 未實測 | Outlook／Teams／Safe Links 點擊後跳板頁→`/index.html` 是否確實帶上 `SameSite=Strict` cookie；未登入時 `/login.html` 是否保留 `#quote:…` 片段 | Demo 階段用真信實測 |

### 4. 環境變數與 Vercel 設定（只列名稱）
- Demo：`MAIL_MODE=redirect`、`MAIL_REDIRECT_TO`（專案負責人提供的測試信箱，**只放 Demo 的環境變數，不寫進 repo**）、`APP_BASE_URL`（Demo 網址）、`CRON_SECRET`。P4 前不會真的寄出（見 §3 的備註）。
- 正式：先 `MAIL_MODE=off` 上線程式碼；P4 完成、補齊 email、試行一位業務＋其主管一週後才改 `live`。`APP_BASE_URL`（正式站網址）、`CRON_SECRET`、`MAIL_GRAPH_*`（P4）。
- `vercel.json`：加 `crons`（見 §6-2-5）。`includeFiles` 目前不含 `lib/`，`lib` 靠 require 追蹤，模組都是相對路徑 require，不需改。
- 回滾：`MAIL_MODE=off` 立即停寄（已入列的 pending 保留，改回 live 後由 `drainDue` 繼續）；移除 `notifyMail` 呼叫後站內通知與現在完全相同。

### 5. 整合驗收（Demo）
1. `node scripts/check-mail-*.js` 全過（見 §8）。
2. Postgres 才有的項目（本機無法驗證，見 §9），**只在 Demo 的 Supabase** 做，用一支不入庫的一次性腳本（`new Pool({connectionString: process.env.DATABASE_URL, ssl: {rejectUnauthorized: false}, max: 5})` ＋ `createOutbox(postgresAdapter({query: (s, p) => pool.query(s, p)}), {})`）：
   1. `await outbox.stats()` 一次（跑 8 條 DDL 建表）→ 確認 `mail_outbox`、`mail_breaker` 存在、RLS 已啟用；對 Supabase 跑 security advisors，確認這兩張表沒有 RLS 警告。
   2. 50 個並行 `enqueue` 同一把 dedupeKey → 表內恰好 1 列。
   3. 入列 30 筆，兩個程序各 25 個並行 `claimDue({limit:1})` → 每筆 `attempts` 都是 1。
   4. `claim` 一筆後手動把 `lease_until` 弄成過去 → `claimDue` 重新領取（attempts=2）→ 用舊的 attempts=1 呼叫 `markFailed` 要得到 `LOST_LEASE`。
   5. 連續 `markFailed` 看 `next_attempt_at` 依序 +60／+300／+900 秒、第 4 次 `failed`；`requeue` 再給一輪；`breaker.record(false, 'AUTH')` 後另一個連線的 `breaker.state()` 要是 open，`record(true)` 重置。
   6. 型別往返（timestamptz、`meta` JSONB、`list` 的篩選與分頁、`purge`）；驗完 `DROP TABLE mail_outbox, mail_breaker`（僅限 Demo）。
   7. 清掃：`claim` 一筆 4 次（每次把 `lease_until` 弄成過去再 `claimDue`）後 `await outbox.expireStale()` → 要回傳那一列（`status=failed`、`last_error_code=LEASE_EXPIRED`；驗證 `UPDATE … RETURNING *` 真的回列），再呼叫一次回傳空陣列。
3. 用 Demo 帳號跑一次完整簽核（一般／需總經理／需董事長／董事會關）＋撤回＋作廢＋顧問成本，檢查寄信紀錄頁每一列、站內通知行為不變。
4. 信件外觀、深色模式、手機、Safe Links 改寫連結後的行為：P4 之後用真信在 Demo 看（目前只有 Edge 看過）。

## 6A. 伺服器端接線（2026-10-08 完成；用戶端見 §6B）

上面 §6 的清單中，**伺服器端項目已全部完成**（P1 伺服器、P2 全部、P3 伺服器）；§6 內寫 `admin.html`／`quote-approval.js`／`login.html`／`app.js`／`index.html`／`sw.js` 的項目已在用戶端階段完成（見 §6B）。
與 §6 原文不同之處（以本節為準）：

| 項目 | §6 原文 | 實際做法與理由 |
|---|---|---|
| 膠水檔名／位置 | `lib/quoteMail.js` | `lib/mail/quoteMail.js`（`lib/mail` 內的檔案只 require 相對路徑與 Node 內建模組，`check-mail-core` 第 0 章會檢查）。路由放在 `lib/mail/routes.js`，`express.static` 之類需要 npm 套件的路由留在 `server.js` |
| 重用 `stepRecipients` 等 | export 或注入 | 注入：`lib/quoteRoutes.js` 註冊路由時呼叫 `mail.bind({getCfg, stepRecipients, isActive, dispName})`，膠水不複製第二份。`findManager1`／`boardProxySet` 只在 `stepRecipients` 內部使用（一級主管用送簽時凍結的 `step.assignee`），膠水不需要直接呼叫 |
| `MAIL_MODE=off` | dispatcher 對每位收件人記 `skipped/MODE_OFF` | **整合層在 off 時根本不呼叫 dispatch**（`notifyMail`、`pollDrain`、Cron 都在最前面返回）：零渲染、零儲存體讀寫、零稽核。理由：正式站先 off 上線，不該因為「部署了程式碼」就在 Supabase 建表（Postgres 版 SQL 尚未對真資料庫驗證）。代價：off 期間發生的事件沒有軌跡、之後也不會補寄。若要保留 MODE_OFF 軌跡，拿掉 `quoteMail.enabled()` 對 off 的判斷即可（dispatcher 本身仍支援） |
| `/q/:id` | `req.params.id` | 用 `req.path.slice(3)`（未解碼）：Express 對路徑參數做 `decodeURIComponent`，壞掉的百分比編碼（`%E0%A4%A`）會在進路由前就被丟成 400 JSON，內容就和其他無效 id 的 404 頁不同 |
| 路由尾端 `drainDue` | 各路由自己呼叫 | `notifyMail` 的 promise 鏈尾端（`dispatch` 之後）呼叫 `drainDue({limit:2, budgetMs:3000})`，同一個請求內多次 `notifyMail`（核准會同時寄 E4 與 E3）共用一次清理（`req._mailDrain`） |
| `poll-bundle` | 在 handler 內 | 一個 middleware（`mailPollDrain`）放在 `requireAuth` 之後、handler 之前：handler 本體與 ETag 計算一個字都沒動。off 時是同步 `next()` |
| `requeue` | 只 requeue | requeue 之後立刻 `drainDue({limit:5, budgetMs:10000})` 一次並回傳該筆最新狀態（「立即重送」）；off 時不清理（工作留在 pending） |
| E6 重試信 | — | 撤回／作廢後 `steps` 已清空，重試信沒有「原關卡」那一列（首次寄送有）。其餘內容逐字相同（`rebuild` 用工作建立時間當 `at`） |

### 掛接點（`lib/quoteRoutes.js`：每個 `notify(...)` 後面一行 `mailNotify(req, ctx, q, spec)`；`notify()` 本身沒動）

| 路由 | 事件 | spec |
|---|---|---|
| `POST /api/quotations`、`PUT /api/quotations/:id`（選定／更換顧問、品項結構變動使成本退回） | E2 | `{type:'E2_COST_REQUEST'}`（stepKey＝`costFlow.requestedAt#costBy`） |
| `PUT /api/quotations/:id/costs`（`done:true`） | E5 | `{type:'E5_COST_DONE'}`（stepKey＝`costFlow.filledAt#done`） |
| `POST /:id/submit`、`POST /:id/reassign` | E1 | `{type:'E1_SUBMIT', idx:0}`（stepKey＝`approval.submittedAt#0`；改派過則為 `approval.submittedAt#0@<最近一次「承辦人真的改變」的改派時間>`，由 `validity.e1StepKey` 從簽核歷史算出，路由不必傳參數；唯一的配合是 `POST /reassign` 要把 `meta:{from,to}` 寫進 `REASSIGN` 歷史，A→A 才能被辨識而不重複寄） |
| `POST /:id/approve` | E4＋E3 | 最後一關：`E4 final_approved`；否則 `E4 approved`（idx＝剛簽的關）＋ `E3 {idx: ap.cur}`（下一關：總經理／董事長／董事會） |
| `POST /:id/return` | E4 | `rejected`（`reason`＝駁回原因） |
| `POST /:id/withdraw`、`PUT /:id`（核准後修改且 `confirmVoid`） | E6 | `withdrawn`／`voided`；先 `mailSnap(q, ctx)` 在清空 `steps` 之前拍下「目前關卡收件人＋已簽過的人＋原關卡＋submittedAt」 |

除了這些 `mailNotify` 行之外，`POST /:id/reassign` 的 `pushHistory(…, 'REASSIGN', …)` 多帶一個參數 `{from, to}`（改派前後承辦人帳號；`pushHistory` 本來就有 `meta` 參數，序列化給前端的歷史只取 `at／byName／action／comment／tier`，不含 `meta`），這是本模組在路由裡唯一的資料面配合：寄信模組靠它辨識「改派給目前的承辦人本人」（A→A）而不重複寄 E1。

事件由單據現況算出（`quoteMail.makeEvent`），金額一律是 `approval.derived` 的整數分；`tierLabel` 用短標籤（董事會關＝`tierLevel 3`＋「董事會」）；
收件人沿用 `notify()` 的過濾（排除操作者、去重、排除非在職），其餘（無 email、網域、格式）由 dispatcher 再檢查並記 `skipped`＋`QUOTE_MAIL_SKIPPED` 稽核。

### API 契約（前端階段依此實作；除註明外皆需管理員，非管理員 403、未登入 401；資料本身不含 HTML，顯示時一律跳脫）

| 路徑 | 請求 | 回應／錯誤 |
|---|---|---|
| `GET /api/admin/users` | — | 每筆多 `email`（無則 `''`）、`emailStatus`（`OK｜NO_EMAIL｜BAD_EMAIL｜DOMAIN_NOT_ALLOWED`） |
| `POST /api/admin/users` | 多一個可省略的 `email`（字串；省略／`null`／空字串＝不設定） | 400 `{error, code, conflictUsername?}`：`code`＝`NOT_STRING｜TOO_LONG｜CONTROL_CHAR｜NON_ASCII｜SPACE｜BAD_CHAR｜BAD_FORMAT｜BAD_LOCAL｜BAD_DOMAIN｜DOMAIN_NOT_ALLOWED｜DUPLICATE`；驗證失敗不建帳號。成功回應不變（`{success:true}`，不回位址） |
| `PUT /api/admin/users/:username` | 多一個 `email`：**沒帶＝不動**；字串＝設定（正規化為小寫）；`''`／`null`／全空白＝清除 | 同上；稽核只寫「email 未設定→已設定／已變更／已清除」 |
| `POST /api/admin/users/email-import/preview` | `{text}`（每行「帳號, Email」，分隔可為逗號／分號／Tab／空白；`#` 註解；上限 500 行／20 萬字元） | 200 `{fatal, summary:{ok,unchanged,error}, rows:[{line,username,email,status:'ok｜unchanged｜error',code?,error?,conflictUsername?}], limits}`。**唯一回傳完整位址的端點**（`rows[].email`）。超過上限 `fatal:true`、`rows` 只有一筆 `line:0` 的錯誤。`text` 不是字串 → 400 |
| `POST /api/admin/users/email-import/apply` | `{text}`（伺服器重新解析再套用；只套用 `ok` 的行） | 200 `{success, updated, unchanged, errors, skipped:[{username,code}], updatedUsers:[{username,detail}]}`（沒有位址）；`fatal`／`text` 不合法 → 400。稽核 `IMPORT_USER_EMAILS` 只有摘要與帳號名稱 |
| `GET /api/quote-approval/config` | — | 管理員的 `users[]` 每筆多 `emailStatus`（同上；不回傳位址）。非管理員沒有 `users` |
| `GET /api/admin/mail/config` | — | `{config:{mode,redirectConfigured,fromName,appBaseUrl,allowedDomains,graphConfigured,timeouts,retry,warnings}, warnings, enabled, backend, cronSecretConfigured, runtime:{enabled,available,mode,bound,poll:{runs,lastAt,last}}, missingEmails:[{username,label,roles,roleLabels,reason,readOnly}]}`。沒有 `redirectTo` 原值、密鑰、本機路徑 |
| `GET /api/admin/mail/outbox` | `?status&type&quoteNo&toUser&since(ISO)&limit(≤500)&offset` | `{available, mode, rows:[record], total, limit, offset, stats:{counts,oldestPendingAt,failed24h,total}, breaker:{open,until,failures,halfOpen}}`；`record` 見 §4（只有 `toMasked`）。非法的篩選值被忽略。寄件匣壞掉 → 500 `{error, code}`；未初始化 → 503 `MAIL_UNAVAILABLE`。只讀、沒有寫入副作用；唯一的例外是 Postgres 上第一次呼叫會 `CREATE TABLE IF NOT EXISTS` 建表（所以 `MAIL_MODE=off` 且沒有人打開後台寄信頁時，儲存體完全不會被碰） |
| `POST /api/admin/mail/outbox/:id/requeue` | — | 200 `{success, id, previousStatus, record, drain:{sent,queued,cancelled,skipped,failed,errors,failedCodes,skippedReasons}\|null}`；`record.status` 可能是 `cancelled`（`skipReason:'STALE'`：單據已走到別關／已撤回，不會誤寄）、`skipped`（log 模式）、`failed`、`sent`。400 `BAD_ID`、404 `NOT_FOUND`、409 `BAD_STATE`（只有 failed／skipped／cancelled 可重送）。稽核 `REQUEUE_MAIL` |
| `GET /api/cron/mail-outbox`（**不需登入**） | `Authorization: Bearer <CRON_SECRET>` | 401（沒有／錯／`CRON_SECRET` 未設）；200 `{ok, enabled, mode, purged, sent, queued, cancelled, skipped, failed, errors, failedCodes, skippedReasons, breakerOpen?, budgetExhausted?}`（只有計數，沒有收件人）；`MAIL_MODE=off` → `enabled:false`、全 0、不碰儲存體 |
| `GET /q/<id>`（**不需登入**） | — | 有效 id（`[A-Za-z0-9_-]{1,64}`）→ 200 跳板頁（`meta refresh` → `/index.html#quote:<id>`，`?cost=1` → `:cost`）；任何無效輸入 → 同一頁 404。不查資料庫 |
| `GET /deep-link.js`（**不需登入**） | — | `_client/deep-link.js`（登入頁需要） |

### 驗證（本機，MAIL_MODE=log／redirect／live／off，JSON 後端；任何模式都不會真寄）

`e2e_mail.js`（API 層，七個階段：main／isolation／off／redirect／live／nosecret／deadline；deadline 階段用 `hang_adapter.js` 預載模組讓儲存體卡住或變慢，驗證 FIX-3 的整體時限，見上面「時限」與 §9）；`e2e_approval`（站內通知 284 項）在 `MAIL_MODE=off` 與 `log` 各跑一次；整套 `run_all_e2e.sh`（off）回歸。數字見該次工作報告。
**未驗證**：Postgres（outbox 建表／RLS／並行領取／`RETURNING *`、`mailPgQuery` 的 pg.Pool）、Microsoft Graph、真實 Outlook／Gmail、JWT 模式（Vercel）的 cookie 內容（本機是 express-session）、
Vercel「回應送出後實例凍結」的行為（`res.end` 包裝等待寄信 promise 的機制只在本機驗證到「回應回來時信已落地」，無法重現凍結）、Vercel Cron 實際觸發與標頭、`drainDue` 在真實網路延遲下的實際耗時、`poll-bundle` 4 秒逾時分支（只驗了節流與正常路徑）。

## 6B. 用戶端接線（2026-10-08 完成）

純前端，沒有改伺服器；沒有放寬 CSP、沒有外部資源、email 不進 `/api/me` 或任何非管理員畫面。

| 檔案 | 內容 |
|---|---|
| `_client/admin.html` | 檔尾獨立的 mail `<script>`（IIFE，只透過 `window.mail*` 與主 script 互通；主 script 以 `typeof` 守衛呼叫，這段壞了帳號頁照常）。①帳號 Modal 的 Email 欄位（`mailInitEmailField`／`mailEmailPayload`／`mailShowEmailError`：新增＝有填才送、編輯＝有改才送，沒帶＝伺服器不動、清空送 `''`；即時格式／網域提示只提示、不擋儲存，伺服器訊息原樣以 `textContent` 顯示）②帳號列表 Email 欄（`mailEmailCell`：遮罩位址 `a***@網域`＋「⚠ 無 email」徽章；完整位址只在編輯 Modal）與缺 email 提示列（`mailAfterUserTable`）③批次匯入對話框（貼上→預覽→確認套用；預覽表一律 `textContent`；預覽後改文字就停用套用）④側欄「寄信」區段（`data-sec="mail"`，lazy-init）：寄信紀錄（篩選：狀態／事件／單號／收件帳號／日期；分頁；遮罩位址；`立即重送`＋確認視窗）、狀態列（模式橫幅、各狀態計數、熔斷）、設定與狀態（唯讀，不顯示轉送位址與密鑰）、缺 email 清單（每次切到該分頁都重抓）。稽核動作名稱（`IMPORT_USER_EMAILS`／`REQUEUE_MAIL`／`QUOTE_MAIL_FAILED`／`QUOTE_MAIL_SKIPPED`）補了中文。手機（≤768px）後台原本沒有 RWD，這次加了一段只在窄螢幕生效的 `@media`，桌面不變。**基本無障礙（FIX-5，只套在 mail 新增的區塊）**：寄信分頁 `role=tablist／tab／tabpanel`＋`aria-selected／aria-controls／aria-labelledby`＋roving tabindex，方向鍵（← → 首尾循環、Home／End）換分頁；批次匯入與重送確認兩個對話框 `role=dialog`＋`aria-modal`＋`aria-labelledby`，開啟時記住原焦點並聚焦第一個欄位、Tab／Shift+Tab 只在對話框內循環、Esc 關閉、關閉後焦點還給原按鈕（共用 `dlgOpen／dlgClose`）；關閉鈕有 `aria-label`；匯入文字框、Email 欄位（`aria-describedby` 指向提示與錯誤區）都有標籤；篩選列 `role=search`＋`aria-label`；寄信紀錄、缺 email 清單、匯入預覽三張表有 `aria-label`；狀態徽章本來就有文字（不只靠顏色）。既有的 `userModal` 等非 mail 區塊沒有動 |
| `_client/quote-approval.js` | 簽核名冊頁：`emailBadge`／`emailProblem` 依 `GET /api/quote-approval/config` 的 `users[].emailStatus` 在名冊（總經理／董事長／董事會代核人／成本填寫人）與秘書旁標「⚠ 無 email」等徽章，頂端彙整提醒；不顯示位址；報價專用章管理人不收簽核信，不標；舊伺服器沒有 `emailStatus` 時什麼都不顯示 |
| `_client/login.html` | 載入 `/deep-link.js`（公開路由）；載入時 `settleOnLoginPage(location.hash)`（網址有合法片段 → 記住；沒有但 app.js 為了登入而導過來（上膛）→ 保留；其他＝使用者主動打開登入頁、登出後回來 → 清掉舊暫存，**下一位登入的人不會被帶去上一位的單據**）；登入成功導向 `'/' + hashForRedirect()`（admin 導向 `/admin.html`，不附片段，但暫存仍會被清掉；取一次即清） |
| `_client/app.js` | `_rememberDeepLink(forLogin)`：`initUser` 的 401 分支導向前 `rememberForLogin`（存＋上膛，給登入頁一次性保留）、強制改密碼分支 `remember`（存、不上膛：改完密碼的 `reload` 要靠它開單）；**所有登出路徑（右上角登出、強制改密碼畫面的「改用其他帳號登入」、閒置自動登出）先 `_clearDeepLink()` 清掉暫存**；`initUser` 結尾先 `consume()` 暫存、`_handleQuoteDeepLink()` 優先用網址片段、沒有才用暫存；`_handleQuoteDeepLink(hashArg)` 的格式驗證改用 `ITTSDeepLink.isValidHash`（deep-link.js 沒載入時退回舊的 uuid 正規式）。強制改密碼完成後的 `location.reload()` 會保留網址片段，重新載入後的 `initUser` 接著開單 |
| `_client/index.html` | 在 `app.js` 之前載入 `deep-link.js` |
| `_client/sw.js` | `isNetworkOnly` 加 `/q/`（Cache API 不理會 `no-store`）；`SW_VERSION` `itts-crm-v3`→`v4`（讓舊使用者換新的 app.js） |

行為備註：
- `/q/<id>` 跳板落在 `/index.html#quote:<id>`（路徑是 `/index.html` 不是 `/`）；登入後的導向才是 `/#quote:<id>`。
- 既有行為（非本次引入，沒有改）：開別人的單，伺服器回 403「無權限」，單據不存在回 404「找不到此報價單」，兩種提示不同（單據 id 是 UUIDv4，不可枚舉）；兩種都不會開面板、不洩漏單號或專案名。
- 既有行為：任何頁面第一次安裝 Service Worker 後 `controllerchange` 會 `location.reload()` 一次（全新瀏覽器的第一次載入）；自動化測試開始前要先讓 SW 就位。
- `MAIL_MODE=off`（正式站預設）：寄信區段照常可開（空紀錄＋off 橫幅），帳號 Email 欄位與缺 email 清單照常可用（可先補 email 再上線）；開啟寄信紀錄頁不會在本機建立寄件匣檔案。Postgres 上第一次開寄信紀錄頁會觸發 `CREATE TABLE IF NOT EXISTS`（見 §6A 的 `GET /api/admin/mail/outbox`），尚未對真實資料庫驗證。
- 驗證：無頭 Edge＋CDP 真實瀏覽器腳本（log 模式 179 項、off 模式 10 項）；批次匯入、寄信紀錄等區段含惡意字串（`<img onerror>`）渲染測試，`window.__pwned` 恆為 0。**未驗證**：真實 Outlook／Teams／Safe Links 點擊後的跳板頁→登入頁→回到單據（只在 Edge 驗證）、其他瀏覽器（Safari／Firefox）、Postgres。

## 7. `rebuild` 與 `isStillValid` 的契約（drainDue 需要）

- `isStillValid(job, ev) → boolean`（可為 Promise）：**只有回傳 `true` 才寄**；`false` → 取消（`STALE`）；丟例外／逾時 → 不寄也不取消，稍後重試（`STALE_CHECK`）。
  判斷是 `lib/mail/validity.js` 的純函式 `checkValidity(job, q, {stepRecipients})` → `{valid, code}`（`quoteMail.checkJob(job)` 取得）。信件是「事件發生當下的通知」，重試可能在很久之後才發生，所以每種事件都要確認單據現況仍與事件相符；未知事件類型、壞掉的 stepKey 一律不寄。下面的 S＝stepKey 內記下的 `approval.submittedAt`（哪一次送簽）：

  | 事件 | 有效條件（全部成立才寄） | 之後哪些變化會讓它過期 |
  |---|---|---|
  | E1 送簽／改派 | `state=pending`、S＝目前的 `submittedAt`、`cur` 仍是該關、stepKey 的改派標記＝簽核歷史最近一次「承辦人真的改變」的 `REASSIGN`（沒有改派則無標記；`meta.from===meta.to` 的同人改派不算）、收件人仍在該關 | 該關已簽／撤回／駁回／重新送簽；再改派（A→B 之後 A 先前未寄出的 E1；B→A 之後 B 的 E1）；收件人被換掉。改派給同一人（A→A）不會讓 A 的 E1 過期 |
  | E3 下一關 | `state=pending`、S 相同、`cur` 仍是該關、收件人仍在該關（同級另一人先簽了就取消） | 該關已簽、撤回、重新送簽、名冊變動 |
  | E2 請填成本 | `costFlow.state=requested`、`costBy` 仍是此人、stepKey＝`requestedAt#costBy` | 顧問完成、換顧問、重新請求 |
  | E4 本關通過 | `state=pending`、S 相同、該關 `status=approved` 且 `cur` 已越過該關 | 全部簽完（已有最終核准信）、被駁回、撤回、作廢、重新送簽 |
  | E4 最終核准 | `state=approved`、S 相同、該關 `status=approved` | 核准後修改使核准作廢（`submittedAt` 變 null）、作廢後重新送簽 |
  | E4 駁回 | `state=returned`、S 相同、該關 `status=returned` | 業務重新送簽（S 變了）。駁回原因（`reason`，業務可輸入的自由文字）只在信件內容顯示，不參與有效性判斷 |
  | E5 成本完成 | `costFlow.state=filled`、`filledAt`＝stepKey 內的時間 | 顧問改回未完成、品項變動使成本退回 `requested`、顧問再完成一次（新的 `filledAt` 另有自己的去重鍵） |
  | E6 撤回 | `state=none` 且 `submittedAt`＝被撤回的那一次（撤回不清掉 `submittedAt`） | 業務重新送簽（「請勿簽核」已過期，不寄）、再撤回一次 |
  | E6 作廢 | `state=none` 且 `submittedAt` 為空（作廢把 `approval` 整個重置） | 業務重新送簽、再撤回、再核准 |

  E4／E5 另外要求收件人仍是單據的業務（`q.owner`）。已知限制：同一張單據被作廢兩次時，較早那次作廢通知的重試仍會通過（單據目前確實處於「核准已作廢」，內容為真；較晚那次另有自己的去重鍵）。`check-mail-integration.js` 裡的膠水 stand-in 對 E4～E6 仍是「單據存在即可」（它只用來驗證 dispatcher 銜接）；真正的規則由 `check-mail-glue.js` 5b／5c 對真的 `quoteMail.js`＋`validity.js` 窮舉。
  它在渲染**之前**被呼叫——所以「事件寄給不該收的角色」會先被判定過期，而不是走到 `RENDER` 失敗。
- `rebuild(job) → null | {ev, kind?}`（可為 Promise）：依 `job.quoteId`／`job.type`／`job.dedupeKey`（取第三個 `:` 之後就是 `stepKey`）讀單據現況重建事件。
  `ev.stepKey` 必須等於工作的 `stepKey`，否則 dispatcher 判定單據已進到別關而取消；單據不存在、或這一關的資料已被清空（撤回／作廢後 `steps` 為空）回 `null` → 取消（`GONE`）；`kind` 不給就沿用 `job.meta.kind`。
  重建出來的事件要用 `at: job.createdAt`，這樣重試信和首次嘗試逐字相同（整合測試有驗）。

## 8. 測試

不開伺服器、不連網、不碰 `data.json`／`auth.json`／`audit.log.json`；只寫系統暫存資料夾。

```
node scripts/check-mail-core.js         # config／safety／recipients／visibility／events／userEmail／link／deep-link 與原始碼衛生
node scripts/check-mail-render.js       # 事件×收件人矩陣、可見性、XSS、金額格式、HTML 良構性、純文字一致性
node scripts/check-mail-outbox.js       # 冪等、50 個並行 claim、退避、requeue、purge、熔斷、三種 adapter 的行為
node scripts/check-mail-dispatch.js     # 傳輸層、模式護欄窮舉、dispatcher、drainDue、永不 throw
node scripts/check-mail-integration.js  # 端到端行程內模擬（本檔 §6 的膠水 stand-in）
node scripts/check-mail-glue.js         # 真的膠水 quoteMail.js／validity.js／routes.js 的單元測試（事件建構、收件人類型、E6 快照、isStillValid／rebuild、E1～E6 過期判斷窮舉與重試路徑、改派去重、時限（永不 resolve／慢／之後才 reject 的 outbox）、poll 節流與逾時、Cron 驗證、批次匯入）
node scripts/check-mail-mutation.js     # 變異測試：在暫存副本上破壞護欄，證明測試會失敗（--list、--only M400,M401）
node scripts/mail-preview.js            # 產生 .mail-preview/gallery/ 供人工檢視
```

`check-mail-integration.js` 涵蓋：E1／E2／E3（一般、需總經理、需董事長、董事會關）／E4／E5／E6 的收件人集合與信件內文可見性；去重；缺 email／停用／網域不合法；
off／log／redirect／live 與模式護欄；失敗注入（TIMEOUT／AUTH／THROTTLED／5xx／拋例外／亂回傳／儲存體唯讀）與恢復；熔斷；`isStillValid` 取消；Hobby 四層重試；稽核與儲存內容衛生；
P1→P2 批次補 email 後補寄；信→跳板頁→深層連結；與真實 `lib/quoteApproval.js` 的 `buildDerived`／`requiredPath` 銜接（檔案可載入時才跑，否則明確標示略過）；惡意內容穿過整條管線；`dispatch`／`drainDue` 永不 throw 總帳。
`M400`–`M417` 是專門由整合測試殺死的變異。`M420`–`M451` 是獨立審查後補的變異：JSON 損壞備份的去重與份數上限（M420–M424）、Postgres 的 RLS 真的被執行／動態 SQL 與參數守衛（M425–M429）、`actorLabel` 清理（M430）、`escAttr` 引號跳脫（M431）、
`drainDue` 的時間預算／熔斷重查／limit／LEASE_EXPIRED 清掃與可見性（M432–M443）、極端虧損單毛利率（M444–M448）、深色模式邊線（M449–M451）。`M460`–`M493` 是整合膠水的變異（測試：`check-mail-glue.js`；變異工具會多帶 `lib/quoteItems.js` 進暫存副本，因為 `quoteMail.js` 相依它）。
第二輪修正之後又加：`M500`–`M514` 秘書／董事會代核人的信不放專案名稱（可見性旗標與主旨）、`M520`–`M532` 過期判斷與改派標記（E1～E6）、`M533`–`M536` 同人改派（A→A）不重複寄、`M540`–`M548` 整體時限、`M560`–`M562` 測試與文件的位址衛生、`M563`–`M566` 位址衛生掃描器收緊（`ROLE_RE` 放寬回去必須被擋）、`M570`–`M575` 登出殘留（深層連結暫存）。

## 9. 已知限制與未驗證

- **Postgres 沒有對真實資料庫驗證**（只有 SQL 靜態審查、假 query 驗參數化、記憶體模擬器跑同一組行為測試）：SQL 語法、型別推斷、鎖行為、並行建表、RLS 都要在 Demo 確認。JSON adapter 只保證同一個 Node 程序內的原子性。
- **真實 Outlook（桌面／網頁）、Gmail、Apple Mail、手機的渲染沒驗證**，只用 Edge 看過；深色模式的 Outlook 專屬選擇器未驗證。
- Graph 傳輸是 P4 stub；沒有對真實 Graph／Exchange 驗證任何東西，也沒測過 Graph 的實際延遲。
- 跳板頁→登入→深層連結的瀏覽器行為（SameSite、片段保留、sessionStorage 在隱私模式）未實測。
- `check-mail-integration.js` 的「膠水」仍是 stand-in（`quoteRoutes.js` 內部函式的等價複本）；真正的膠水是 `quoteMail.js`，它由 API 層端對端測試 `e2e_mail.js` 覆蓋（對本機伺服器、走完整條簽核）。那支測試不在 repo 內（放在工作用的暫存資料夾，依賴本機測試帳號與環境變數），重現方式見 §6A「驗證」。
- 計時類測試（逾時 100 毫秒、時間預算）在 CPU 滿載下仍通過，但屬時間敏感測試。
- `drainDue` 的預算只管「還要不要領下一筆」：已開始的寄送會做完。最壞耗時 ≈ 預算 16 秒 + 單封逾時 8 秒，**假設 `getUsers`／`isStillValid`／`render`／`rebuild` 這些輔助呼叫很快**（各自有 5 秒逾時，但全部同時變慢時會超過 30 秒）。真正接線後要用 Demo 量一次實際耗時。
- Postgres 新增的 `expireExhausted … RETURNING *`（清掃結果回傳）同樣只用假 query 驗證過，真資料庫上的行為要在 Demo 跑上面 §5-2-7。
- **時限（FIX-3）只保證「回應不被卡住」**：outbox 在入列完成前卡住（例如資料庫連線半開）時，這個事件沒有任何寄件匣紀錄，也不會被補寄，只留稽核 `QUOTE_MAIL_FAILED code=DEADLINE`（可到後台稽核紀錄查）；入列之後才卡住的由租約／清理接手。HTTP 層的時限測試靠一個不在 repo 內的測試預載模組（把 `jsonFileAdapter` 換成會卡住的版本）；Postgres 的 `query_timeout`／`statement_timeout` 沒有設，真實資料庫上的卡住行為未驗證。
- **登入頁暫存的取捨（FIX-4）**：app.js 為了登入而導向登入頁時，暫存只保證給「第一次載入的登入頁」（一次性旗標）；在登入頁按 F5 重新整理（網址上沒有片段）會失去這個暫存，登入後落在首頁而不是那張單。換來的是：使用者主動打開登入頁、登出後回到登入頁，一定不會沿用上一位的暫存。
- **無障礙（FIX-5）只驗證到屬性、無障礙樹與鍵盤行為**（無頭 Edge＋CDP）；沒有用真實螢幕閱讀器（NVDA／JAWS／VoiceOver）實測，也沒有改既有的帳號 Modal（`userModal`）等非 mail 區塊。
- JSON adapter 的「已備份過」記在 adapter 實例的記憶體裡：程序重啟後第一次讀到同樣的損壞會再備份一次（總份數仍受 3 份上限約束）。
- 深色模式的邊線修正（決策條色塊間隔改用卡片底色、品項表外框改用深色邊線）只用 Edge 驗證過；Outlook 桌面版（Word 引擎）忽略 `prefers-color-scheme`，仍是淺色版；Outlook.com 的自動反色對邊線的處理未驗證。

## 10. 開發備忘（寫這些檔案時踩過的坑）

- 本環境的寫檔工具會把四位數的 `\uXXXX`（值 ≥0xA0）轉成真正字元，U+2028／2029 落在正規式字面值裡就是語法錯誤。原始碼一律用 `\x..`、`\u{...}` 或 `String.fromCodePoint`；`check-mail-core.js` 第 0 章會掃 `lib/mail`、`scripts/check-mail-*.js`、`_client/deep-link.js` 的不可見字元。
- `fs.cpSync` 在含中文路徑的 Windows 上會讓 Node 無聲當掉（exit 127），變異工具用自己的遞迴複製。
- 用 Node `vm` 模擬瀏覽器時，要丟例外的 getter 必須掛在普通物件上（沙箱全域上的 getter 例外會被吞掉）。
- username 比對**區分大小寫**（實際帳號資料存在只差大小寫的不同帳號），任何新程式碼都不可 case-fold username；email 的唯一性才不分大小寫。
- 公開 repo：不得出現真實 email、密鑰、租戶 ID、客戶名、人名；測試信箱只放環境變數。`.mail-preview/` 與 `mail-outbox.json` 已在 `.gitignore`（暫存檔 `*.tmp` 與損壞備份 `*.bak` 被既有規則涵蓋）。
