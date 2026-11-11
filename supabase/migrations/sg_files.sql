-- 班次空檔分析網站：共用檔案庫（由 tools\deploy.ps1 透過 Management API 執行，可重複跑）
-- 通行碼要改：update public.sg_settings set value = '新碼' where key = 'view_code';  （admin_code 同理）

create table if not exists public.sg_files (
  key         text primary key,                 -- 'schedule/1105'、'needs/1105-1106'
  kind        text not null check (kind in ('schedule', 'needs', 'matrix', 'actual', 'fleet')),
  day         text not null,                    -- MMDD（needs 可為日期範圍）
  name        text not null,                    -- 原始檔名
  size        integer not null default 0,
  mime        text,
  data        text not null,                    -- 檔案內容 base64
  summary     jsonb,                            -- 班次表：{tractors,trac,full,gaps}；清單：{days,rows}
  uploader    text,
  uploaded_at timestamptz not null default now()
);
create index if not exists sg_files_uploaded_idx on public.sg_files (uploaded_at);
-- 既有表補上 matrix（站間車程累積表）這種 kind；可重複執行
alter table public.sg_files drop constraint if exists sg_files_kind_check;
alter table public.sg_files add constraint sg_files_kind_check check (kind in ('schedule', 'needs', 'matrix', 'actual', 'fleet'));

create table if not exists public.sg_settings (
  key   text primary key,
  value text not null
);
-- 通行碼不放在這裡（repo 是公開的）：deploy.ps1 會用 tools\site.local.json 的 view／admin 覆蓋下面的佔位值
insert into public.sg_settings (key, value) values
  ('view_code', 'CHANGE-ME-VIEW'),    -- 看板通行碼（看）
  ('admin_code', 'CHANGE-ME-ADMIN'),  -- 管理碼（上傳／覆蓋／刪除）
  ('retain_days', '21')               -- 檔案保留天數
on conflict (key) do nothing;

-- 只有 service role（Edge Function）能碰；anon/authenticated 完全不給
alter table public.sg_files enable row level security;
alter table public.sg_settings enable row level security;
revoke all on public.sg_files from anon, authenticated;
revoke all on public.sg_settings from anon, authenticated;
grant all on public.sg_files to service_role;
grant all on public.sg_settings to service_role;
