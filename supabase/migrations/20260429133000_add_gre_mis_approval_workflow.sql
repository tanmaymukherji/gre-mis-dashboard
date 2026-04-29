alter table public.gre_mis_needs
  add column if not exists approval_status text not null default 'approved',
  add column if not exists imported_from_batch text,
  add column if not exists source_kind text not null default 'manual',
  add column if not exists last_status_change_at timestamptz;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'gre_mis_needs_approval_status_check'
  ) then
    alter table public.gre_mis_needs
      add constraint gre_mis_needs_approval_status_check
      check (approval_status in ('pending_admin', 'approved', 'rejected'));
  end if;
end
$$;

create table if not exists public.gre_mis_admin_web_sessions (
  id uuid primary key default gen_random_uuid(),
  username text not null,
  token_hash text not null unique,
  expires_at timestamptz not null,
  last_used_at timestamptz,
  created_at timestamptz not null default now()
);

create table if not exists public.gre_mis_update_requests (
  id uuid primary key default gen_random_uuid(),
  need_id text not null references public.gre_mis_needs(id) on delete cascade,
  submitted_by_curator_name text,
  submitted_by_curator_email text not null,
  proposed_status text,
  proposed_internal_status text,
  proposed_next_action text,
  proposed_curation_notes text,
  proposed_curation_call_date date,
  proposed_demand_broadcast_needed boolean,
  proposed_solutions_shared_count integer,
  proposed_invited_providers_count integer,
  review_notes text,
  approval_status text not null default 'pending',
  reviewed_by_email text,
  reviewed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'gre_mis_update_requests_approval_status_check'
  ) then
    alter table public.gre_mis_update_requests
      add constraint gre_mis_update_requests_approval_status_check
      check (approval_status in ('pending', 'approved', 'rejected'));
  end if;
end
$$;

drop trigger if exists gre_mis_update_requests_set_updated_at on public.gre_mis_update_requests;
create trigger gre_mis_update_requests_set_updated_at
before update on public.gre_mis_update_requests
for each row execute function public.set_updated_at();

alter table public.gre_mis_admin_web_sessions enable row level security;
alter table public.gre_mis_update_requests enable row level security;

drop policy if exists "public can create needs" on public.gre_mis_needs;
create policy "public can create needs"
on public.gre_mis_needs
for insert
to anon, authenticated
with check (approval_status = 'pending_admin');

drop policy if exists "admins and curators read needs" on public.gre_mis_needs;
create policy "public reads approved needs and internal users read all"
on public.gre_mis_needs
for select
to anon, authenticated
using (
  approval_status = 'approved'
  or public.is_gre_mis_admin()
  or public.is_gre_mis_curator()
);

drop policy if exists "admins and assigned curators update needs" on public.gre_mis_needs;
create policy "admins update needs directly"
on public.gre_mis_needs
for update
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

drop policy if exists "admins and curators read updates" on public.gre_mis_need_updates;
create policy "public read updates for approved needs"
on public.gre_mis_need_updates
for select
to anon, authenticated
using (
  exists (
    select 1
    from public.gre_mis_needs n
    where n.id = gre_mis_need_updates.need_id
      and (
        n.approval_status = 'approved'
        or public.is_gre_mis_admin()
        or public.is_gre_mis_curator()
      )
  )
);

drop policy if exists "admins and curators write updates" on public.gre_mis_need_updates;
create policy "internal users write updates"
on public.gre_mis_need_updates
for insert
to authenticated
with check (public.is_gre_mis_admin() or public.is_gre_mis_curator());

create policy "admins read admin web sessions"
on public.gre_mis_admin_web_sessions
for select
to authenticated
using (public.is_gre_mis_admin());

create policy "admins read update requests"
on public.gre_mis_update_requests
for select
to authenticated
using (public.is_gre_mis_admin());

create policy "admins manage update requests"
on public.gre_mis_update_requests
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

update public.gre_mis_needs
set approval_status = coalesce(approval_status, 'approved'),
    source_kind = coalesce(source_kind, 'manual')
where approval_status is null
   or source_kind is null;
