-- Verified organisation membership for My Solutions / My Help Requests.
-- This replaces free-text organisation matching as an authorization source.

create extension if not exists pgcrypto;

create or replace function public.gre_mis_normalize_org_name(p_value text)
returns text
language sql
immutable
as $$
  select lower(trim(regexp_replace(coalesce(p_value, ''), '\s+', ' ', 'g')));
$$;

create table if not exists public.gre_mis_organisations (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  normalized_name text not null,
  gre_trader_id text,
  source text not null default 'gre_mis',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint gre_mis_organisations_normalized_name_unique unique (normalized_name)
);

create unique index if not exists gre_mis_organisations_trader_unique
  on public.gre_mis_organisations (gre_trader_id)
  where gre_trader_id is not null and gre_trader_id <> '';

create table if not exists public.gre_mis_organisation_memberships (
  id uuid primary key default gen_random_uuid(),
  organisation_id uuid not null references public.gre_mis_organisations(id) on delete cascade,
  user_id uuid references public.gre_mis_users(id) on delete cascade,
  user_email text not null,
  role text not null default 'org_viewer',
  status text not null default 'pending',
  requested_by uuid references public.gre_mis_users(id) on delete set null,
  requested_at timestamptz not null default now(),
  approved_by uuid references public.gre_mis_users(id) on delete set null,
  approved_at timestamptz,
  notes text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint gre_mis_org_memberships_email_lowercase check (user_email = lower(user_email)),
  constraint gre_mis_org_memberships_role_check check (role in ('org_admin', 'org_editor', 'org_viewer')),
  constraint gre_mis_org_memberships_status_check check (status in ('pending', 'approved', 'rejected', 'revoked')),
  constraint gre_mis_org_memberships_user_or_email_unique unique (organisation_id, user_email)
);

create index if not exists gre_mis_org_memberships_user_idx
  on public.gre_mis_organisation_memberships (user_id, status);

create index if not exists gre_mis_org_memberships_email_idx
  on public.gre_mis_organisation_memberships (user_email, status);

alter table public.gre_mis_form_submissions
  add column if not exists organisation_id uuid references public.gre_mis_organisations(id) on delete set null;

alter table public.gre_mis_needs
  add column if not exists organisation_id uuid references public.gre_mis_organisations(id) on delete set null;

drop trigger if exists gre_mis_organisations_set_updated_at on public.gre_mis_organisations;
create trigger gre_mis_organisations_set_updated_at
before update on public.gre_mis_organisations
for each row execute function public.set_updated_at();

drop trigger if exists gre_mis_org_memberships_set_updated_at on public.gre_mis_organisation_memberships;
create trigger gre_mis_org_memberships_set_updated_at
before update on public.gre_mis_organisation_memberships
for each row execute function public.set_updated_at();

alter table public.gre_mis_organisations enable row level security;
alter table public.gre_mis_organisation_memberships enable row level security;

do $$ begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'public'
      and tablename = 'gre_mis_organisations'
      and policyname = 'Admins manage GRE MIS organisations'
  ) then
    create policy "Admins manage GRE MIS organisations"
      on public.gre_mis_organisations
      for all
      to authenticated
      using (public.is_gre_mis_admin())
      with check (public.is_gre_mis_admin());
  end if;
end $$;

do $$ begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'public'
      and tablename = 'gre_mis_organisation_memberships'
      and policyname = 'Admins manage GRE MIS organisation memberships'
  ) then
    create policy "Admins manage GRE MIS organisation memberships"
      on public.gre_mis_organisation_memberships
      for all
      to authenticated
      using (public.is_gre_mis_admin())
      with check (public.is_gre_mis_admin());
  end if;
end $$;

-- Backfill canonical organisations from approved local submissions and GRE trader records.
insert into public.gre_mis_organisations (name, normalized_name, gre_trader_id, source)
select
  name,
  normalized_name,
  gre_trader_id,
  'traders'
from (
  select distinct on (public.gre_mis_normalize_org_name(coalesce(t.organisation_name, t.trader_name)))
    trim(coalesce(t.organisation_name, t.trader_name)) as name,
    public.gre_mis_normalize_org_name(coalesce(t.organisation_name, t.trader_name)) as normalized_name,
    nullif(trim(t.trader_id::text), '') as gre_trader_id
  from public.traders t
  where public.gre_mis_normalize_org_name(coalesce(t.organisation_name, t.trader_name)) <> ''
  order by public.gre_mis_normalize_org_name(coalesce(t.organisation_name, t.trader_name)), t.trader_id
) t
on conflict (normalized_name) do update
set
  name = excluded.name,
  gre_trader_id = coalesce(public.gre_mis_organisations.gre_trader_id, excluded.gre_trader_id),
  source = coalesce(public.gre_mis_organisations.source, excluded.source);

insert into public.gre_mis_organisations (name, normalized_name, gre_trader_id, source)
select
  name,
  normalized_name,
  gre_trader_id,
  'form_submissions'
from (
  select distinct on (public.gre_mis_normalize_org_name(s.organization_name))
    trim(s.organization_name) as name,
    public.gre_mis_normalize_org_name(s.organization_name) as normalized_name,
    nullif(trim(s.existing_trader_id), '') as gre_trader_id
  from public.gre_mis_form_submissions s
  where public.gre_mis_normalize_org_name(s.organization_name) <> ''
  order by public.gre_mis_normalize_org_name(s.organization_name), s.created_at desc
) s
on conflict (normalized_name) do update
set
  gre_trader_id = coalesce(public.gre_mis_organisations.gre_trader_id, excluded.gre_trader_id);

update public.gre_mis_form_submissions s
set organisation_id = o.id
from public.gre_mis_organisations o
where s.organisation_id is null
  and (
    public.gre_mis_normalize_org_name(s.organization_name) = o.normalized_name
    or (nullif(trim(s.existing_trader_id), '') is not null and nullif(trim(s.existing_trader_id), '') = o.gre_trader_id)
  );

update public.gre_mis_needs n
set organisation_id = o.id
from public.gre_mis_organisations o
where n.organisation_id is null
  and public.gre_mis_normalize_org_name(n.organization_name) = o.normalized_name;

-- Existing signed-in submitters are approved as organisation admins for the
-- organisation they already submitted under. This preserves current legitimate
-- access while preventing new random users from self-claiming an organisation.
insert into public.gre_mis_organisation_memberships (
  organisation_id,
  user_id,
  user_email,
  role,
  status,
  approved_at,
  notes
)
select distinct
  s.organisation_id,
  s.submitter_user_id,
  lower(trim(s.submitter_email)),
  'org_admin',
  'approved',
  now(),
  'Backfilled from existing signed-in submission ownership.'
from public.gre_mis_form_submissions s
where s.organisation_id is not null
  and lower(trim(coalesce(s.submitter_email, ''))) <> ''
  and s.source_mode in ('signed_in', 'signed_in_edit')
on conflict (organisation_id, user_email) do update
set
  user_id = coalesce(public.gre_mis_organisation_memberships.user_id, excluded.user_id),
  role = case
    when public.gre_mis_organisation_memberships.role = 'org_admin' then 'org_admin'
    else excluded.role
  end,
  status = case
    when public.gre_mis_organisation_memberships.status = 'approved' then 'approved'
    else excluded.status
  end,
  approved_at = coalesce(public.gre_mis_organisation_memberships.approved_at, excluded.approved_at);
