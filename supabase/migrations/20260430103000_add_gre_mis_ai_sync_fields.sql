alter table public.gre_mis_needs
add column if not exists source_row_signature text,
add column if not exists last_synced_at timestamptz,
add column if not exists ai_thematic_area text,
add column if not exists ai_application_area text,
add column if not exists ai_need_kind text,
add column if not exists ai_service_kind text,
add column if not exists ai_keywords text[] not null default '{}',
add column if not exists ai_6m_signals text[] not null default '{}',
add column if not exists ai_summary text,
add column if not exists ai_engine text,
add column if not exists ai_enriched_at timestamptz,
add column if not exists ai_enrichment_status text,
add column if not exists ai_payload jsonb;

create table if not exists public.gre_mis_import_runs (
  id uuid primary key default gen_random_uuid(),
  file_name text,
  source_kind text not null default 'website_inbound_snapshot',
  imported_by_email text,
  total_rows integer not null default 0,
  inserted_count integer not null default 0,
  updated_count integer not null default 0,
  ai_updated_count integer not null default 0,
  created_at timestamptz not null default now()
);

alter table public.gre_mis_import_runs enable row level security;

drop policy if exists "admins manage import runs" on public.gre_mis_import_runs;
create policy "admins manage import runs"
on public.gre_mis_import_runs
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());
