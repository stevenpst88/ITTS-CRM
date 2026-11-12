-- ════════════════════════════════════════════════════════════
--  SEC-01：撤銷 anon / authenticated 對 public 資料表的權限，並開啟 RLS
--
--  為什麼：整個 CRM 的資料（含帳號密碼雜湊、加密的 SAP 密碼）都在 public.app_data。
--  Supabase 會把 public schema 的資料表自動開放給 Data API（PostgREST），
--  若 RLS 關閉又沒撤權，任何拿到 anon key（公開金鑰）的人都能讀寫整張表。
--  本應用程式只用 DATABASE_URL 直連 Postgres（db/postgres.js），前端完全不用
--  Supabase client，所以 anon / authenticated 不需要任何權限。
--
--  做法：REVOKE ＋ ENABLE ROW LEVEL SECURITY（刻意不建任何 policy＝Data API 一律拒絕）。
--  應用程式連線用的 postgres 角色是表的 owner / bypass RLS，不受影響。
--
--  可重複執行（idempotent）。執行方式：Supabase Dashboard → SQL Editor 貼上 → Run。
--  先在 DEV 執行並驗證，再由專案負責人決定何時套用到正式環境。
--  注意：Storage（uploads bucket）走 storage schema 與 service_role，不受此檔影響。
-- ════════════════════════════════════════════════════════════

-- 1. 主資料表
REVOKE ALL ON public.app_data FROM PUBLIC;
REVOKE ALL ON public.app_data FROM anon, authenticated;
ALTER TABLE public.app_data ENABLE ROW LEVEL SECURITY;

-- 2. 其他由應用程式建立的資料表（稽核表、信件寄送佇列／斷路器；若尚未建立就略過）
DO $$
DECLARE t text;
BEGIN
  FOREACH t IN ARRAY ARRAY['audit_events', 'mail_outbox', 'mail_breaker'] LOOP
    IF to_regclass('public.' || t) IS NOT NULL THEN
      EXECUTE format('REVOKE ALL ON public.%I FROM PUBLIC', t);
      EXECUTE format('REVOKE ALL ON public.%I FROM anon, authenticated', t);
      EXECUTE format('ALTER TABLE public.%I ENABLE ROW LEVEL SECURITY', t);
    END IF;
  END LOOP;
END $$;

-- 3. 其他 public 資料表：先用下面「驗證」第 2 段列出，逐一確認後再處理。
--    （此檔不盲目對所有表動手，避免誤傷你另外建立、確實要開放給前端的表。）

-- 4. 驗證（唯讀；預期 anon_select／anon_update 為 false、rls_on 為 true）
--   SELECT c.relname AS table_name,
--          c.relrowsecurity AS rls_on,
--          has_table_privilege('anon', c.oid, 'SELECT') AS anon_select,
--          has_table_privilege('anon', c.oid, 'UPDATE') AS anon_update,
--          has_table_privilege('authenticated', c.oid, 'SELECT') AS authed_select
--     FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
--    WHERE n.nspname = 'public' AND c.relkind = 'r'
--    ORDER BY 1;
--
--   REST 驗證（應回 401／403 或空陣列，不可回出資料）：
--   curl "$SUPABASE_URL/rest/v1/app_data?select=id" -H "apikey: $ANON_KEY"
