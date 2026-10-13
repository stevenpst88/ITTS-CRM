-- ════════════════════════════════════════════════════════════
--  稽核事件表 audit_events 的「只增不改」保護（選用；由管理者在 Supabase SQL Editor 手動執行）
--
--  背景：程式（db/postgres.js → lib/auditStore.js）只會對 audit_events 做 INSERT 與 SELECT，
--  從不 UPDATE／DELETE／TRUNCATE。本檔把這個約定「下放到資料庫」，即使應用程式有 bug 或帳密外洩，
--  也不能用應用程式使用的角色竄改／刪除稽核紀錄。
--
--  執行前提：應用程式至少啟動過一次（ensureSchema 會自動建立 audit_events）；
--  或先執行下面第 1 段（與程式內建立語法相同、可重複執行）。
--
--  重要限制（請讀）：
--   * 表的擁有者（owner）與 superuser 可以繞過 REVOKE，也可以 DROP TRIGGER。
--     Supabase 預設的 `postgres` 角色即為擁有者。要達到真正的「應用程式碰不到」，
--     建議另建一個專供應用程式連線的低權限角色（見第 4 段範例），並把 DATABASE_URL 改用它。
--     若暫時不改連線角色，本檔的 trigger 仍會攔住「一般 UPDATE／DELETE／TRUNCATE」（含 SQL Editor 誤操作），
--     只是擁有者有意識地停用 trigger 時仍可繞過——這類操作會留在 Supabase 的資料庫稽核（pgaudit）／平台日誌。
--   * 本檔不會刪除或搬動任何現有資料；舊的 app_data.id='audit_log' 整包 JSONB 保留為唯讀歷史。
--
--  回復方式：DROP TRIGGER audit_events_no_update_delete ON public.audit_events; 以及 audit_events_no_truncate。
-- ════════════════════════════════════════════════════════════

-- 1. 確保表存在（與 lib/auditStore.js 的 TABLE_SQL 相同）
CREATE TABLE IF NOT EXISTS public.audit_events (
  id        BIGSERIAL PRIMARY KEY,
  ts        TIMESTAMPTZ NOT NULL DEFAULT now(),
  kind      TEXT NOT NULL,
  action    TEXT NOT NULL DEFAULT '',
  operator  TEXT NOT NULL DEFAULT '',
  target    TEXT NOT NULL DEFAULT '',
  detail    TEXT NOT NULL DEFAULT '',
  ip        TEXT NOT NULL DEFAULT '',
  ua        TEXT NOT NULL DEFAULT '',
  extra     JSONB NOT NULL DEFAULT '{}'::jsonb
);
COMMENT ON TABLE public.audit_events IS '稽核事件（只增不改）。kind: log=一般操作稽核, contact=名片欄位級異動, quote-edit=報價單原始編輯事件。應用程式只 INSERT/SELECT。';

-- 2. 攔截 UPDATE／DELETE（任何角色，含擁有者；要繞過必須先 DROP TRIGGER）
CREATE OR REPLACE FUNCTION public.audit_events_block_mutation() RETURNS trigger
LANGUAGE plpgsql AS $$
BEGIN
  RAISE EXCEPTION 'audit_events 為只增不改的稽核表，禁止 % 操作', TG_OP
    USING ERRCODE = 'insufficient_privilege';
END;
$$;

DROP TRIGGER IF EXISTS audit_events_no_update_delete ON public.audit_events;
CREATE TRIGGER audit_events_no_update_delete
  BEFORE UPDATE OR DELETE ON public.audit_events
  FOR EACH ROW EXECUTE FUNCTION public.audit_events_block_mutation();

DROP TRIGGER IF EXISTS audit_events_no_truncate ON public.audit_events;
CREATE TRIGGER audit_events_no_truncate
  BEFORE TRUNCATE ON public.audit_events
  FOR EACH STATEMENT EXECUTE FUNCTION public.audit_events_block_mutation();

-- 3. 收回對外角色的修改權限（Supabase 的 anon／authenticated 本來就不該碰這張表）
REVOKE ALL ON public.audit_events FROM PUBLIC;
REVOKE ALL ON public.audit_events FROM anon, authenticated;
REVOKE UPDATE, DELETE, TRUNCATE ON public.audit_events FROM postgres;   -- 擁有者仍可自行 GRANT 回來；此處是「預設不給」
GRANT SELECT, INSERT ON public.audit_events TO postgres;
GRANT USAGE, SELECT ON SEQUENCE public.audit_events_id_seq TO postgres;

-- 4. （建議，選用）專供應用程式連線的低權限角色：只能 INSERT／SELECT 稽核表
--    先把 <強密碼> 換成你自己產生的隨機密碼（不要把密碼寫進程式碼或 git），再取消註解執行；
--    之後把 Vercel 環境變數 DATABASE_URL 改成使用此角色的連線字串。
--    注意：此角色也需要 app_data 等其他表的權限（見 supabase/schema.sql），請依實際需求另行 GRANT。
-- CREATE ROLE itts_app LOGIN PASSWORD '<強密碼>';
-- GRANT SELECT, INSERT ON public.audit_events TO itts_app;
-- GRANT USAGE, SELECT ON SEQUENCE public.audit_events_id_seq TO itts_app;
-- GRANT SELECT, INSERT, UPDATE ON public.app_data TO itts_app;

-- 5. 驗證（預期第一行成功、第二行出錯）
--   INSERT INTO public.audit_events (kind, action, operator, detail) VALUES ('log', 'PROTECT_TEST', 'dba', 'insert ok');
--   UPDATE public.audit_events SET detail = 'x' WHERE action = 'PROTECT_TEST';   -- 應報錯：禁止 UPDATE 操作
