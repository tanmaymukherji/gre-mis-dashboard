-- Migration: Add GRE solution links mapping table
-- Purpose: Map local MIS-* offering IDs to GRE numeric IDs for all linked entities
-- Created: 2026-07-18

create table if not exists public.gre_solution_links (
    local_offering_id text primary key references public.offerings(offering_id) on delete cascade,
    local_solution_id text not null,
    local_submission_id uuid,
    local_trader_id text not null,
    gre_trader_id text,
    gre_solution_id bigint,
    gre_product_id bigint,
    gre_sku_id bigint,
    gre_channel_product_id bigint,
    gre_status_summary jsonb,              -- snapshot of all status values per linked entity
    payload_hash text,                     -- sha256 of frozen payload sent
    gre_response_hash text,                -- read-back hash
    upload_state text not null default 'pending_verification',  -- 'pending_verification' | 'synced' | 'quarantined' | 'rolled_back'
    last_attempted_at timestamptz,
    verified_at timestamptz,
    acknowledged_warnings jsonb,
    uploaded_by_email text,
    manual_review_reason text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

-- Indexes for common lookups
create index if not exists idx_gre_solution_links_gre_solution_id on public.gre_solution_links (gre_solution_id);
create index if not exists idx_gre_solution_links_local_solution_id on public.gre_solution_links (local_solution_id);
create index if not exists idx_gre_solution_links_upload_state on public.gre_solution_links (upload_state);
create index if not exists idx_gre_solution_links_local_trader_id on public.gre_solution_links (local_trader_id);

-- Helper function for RLS: checks if current user has admin/moderator role (matches app.js hasAdminLikeAccess)
create or replace function public.gre_mis_rls_is_admin_like()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select (auth.jwt() -> 'app_metadata' ->> 'grameee_role') in ('admin', 'moderator')
$$;

grant execute on function public.gre_mis_rls_is_admin_like() to anon, authenticated, service_role;

-- RLS: Admins/moderators can read; Service role writes
alter table public.gre_solution_links enable row level security;

create policy "gre_solution_links_admin_read" on public.gre_solution_links
    for select
    using (public.gre_mis_rls_is_admin_like());

create policy "gre_solution_links_service_write" on public.gre_solution_links
    for all
    using (auth.role() = 'service_role')
    with check (auth.role() = 'service_role');

-- Updated_at trigger
create trigger update_gre_solution_links_updated_at
    before update on public.gre_solution_links
    for each row
    execute function public.set_updated_at();

-- Comment
comment on table public.gre_solution_links is 'Maps local MIS offering/solution IDs to GRE numeric IDs. One row per local offering. Never substitutes local IDs with GRE IDs — both coexist.';