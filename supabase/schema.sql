-- ════════════════════════════════════════════════════════════
--  ITTS CRM - Supabase Postgres Schema
--
--  策略：保留 JSON-blob 的存法（app_data 表）
--  業務邏輯完全不動，整個 data.json 會變成這個表的 content 欄位
--
--  執行方式：
--    1. Supabase Dashboard → SQL Editor
--    2. 貼上本檔內容並執行
-- ════════════════════════════════════════════════════════════

-- ── 主資料表（存整個 CRM JSON blob）──
CREATE TABLE IF NOT EXISTS public.app_data (
  id         TEXT PRIMARY KEY,
  content    JSONB NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- 初始化 main row（如果不存在就建空的）
INSERT INTO public.app_data (id, content)
VALUES ('main', '{"contacts":[]}'::jsonb)
ON CONFLICT (id) DO NOTHING;

-- ── 索引 ──
CREATE INDEX IF NOT EXISTS idx_app_data_updated_at
  ON public.app_data (updated_at DESC);

-- ════════════════════════════════════════════════════════════
--  RLS (Row Level Security)：關閉，因為我們用 service_role 連線
--  server.js 自己做權限控制（getViewableOwners）
-- ════════════════════════════════════════════════════════════
ALTER TABLE public.app_data DISABLE ROW LEVEL SECURITY;

-- ════════════════════════════════════════════════════════════
--  Storage Bucket：名片圖片
--  執行方式（Dashboard 手動，或下面的 SQL，2 擇 1）：
--    A) Dashboard → Storage → Create bucket "uploads"，設為 private
--    B) 執行下面 SQL（需要 service_role 權限）
-- ════════════════════════════════════════════════════════════
INSERT INTO storage.buckets (id, name, public)
VALUES ('uploads', 'uploads', false)
ON CONFLICT (id) DO NOTHING;

-- Storage policy：只允許 service_role 讀寫（預設行為，無需額外 policy）
-- 一般 anon/authenticated 不能直接存取，全部走 server.js 代理

-- ════════════════════════════════════════════════════════════
--  驗證
-- ════════════════════════════════════════════════════════════
-- SELECT id, jsonb_object_keys(content) FROM public.app_data WHERE id = 'main';
-- SELECT id, name, public FROM storage.buckets WHERE id = 'uploads';
