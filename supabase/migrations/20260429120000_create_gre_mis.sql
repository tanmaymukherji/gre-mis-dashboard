create extension if not exists pgcrypto;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create table if not exists public.gre_mis_admins (
  email text primary key,
  display_name text,
  created_at timestamptz not null default now(),
  constraint gre_mis_admins_email_lowercase check (email = lower(email))
);

create table if not exists public.gre_mis_curators (
  id uuid primary key default gen_random_uuid(),
  display_name text not null,
  email text unique not null,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint gre_mis_curators_email_lowercase check (email = lower(email))
);

create table if not exists public.gre_mis_options (
  id uuid primary key default gen_random_uuid(),
  option_type text not null,
  label text not null,
  sort_order integer not null default 100,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint gre_mis_options_type_check check (option_type in ('status', 'internal_status', 'category', 'next_action'))
);

create table if not exists public.gre_mis_needs (
  id text primary key,
  organization_name text not null,
  website text,
  contact_person text,
  designation text,
  seeker_phone text,
  seeker_email text not null,
  requested_on timestamptz not null default now(),
  problem_statement text not null,
  state text,
  district text,
  status text not null default 'New',
  internal_status text not null default 'Need solution providers',
  curator_id uuid references public.gre_mis_curators(id) on delete set null,
  curation_call_date date,
  curation_age_days integer not null default 0,
  curation_notes text,
  curated_need text[] not null default '{}',
  demand_broadcast_needed boolean not null default false,
  solutions_shared_count integer not null default 0,
  invited_providers_count integer not null default 0,
  next_action text,
  seeker_response_status text,
  published_for_open_response boolean not null default false,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.gre_mis_need_updates (
  id uuid primary key default gen_random_uuid(),
  need_id text not null references public.gre_mis_needs(id) on delete cascade,
  update_type text not null,
  note text,
  created_by_email text,
  created_at timestamptz not null default now()
);

create table if not exists public.gre_mis_solution_providers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  email text not null,
  website text,
  organization_type text,
  states_served text[] not null default '{}',
  solution_tags text[] not null default '{}',
  notes text,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.gre_mis_solutions (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  summary text,
  categories text[] not null default '{}',
  states text[] not null default '{}',
  provider_id uuid references public.gre_mis_solution_providers(id) on delete set null,
  external_link text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.gre_mis_need_matches (
  id uuid primary key default gen_random_uuid(),
  need_id text not null references public.gre_mis_needs(id) on delete cascade,
  provider_id uuid references public.gre_mis_solution_providers(id) on delete cascade,
  solution_id uuid references public.gre_mis_solutions(id) on delete set null,
  match_score numeric(6,2),
  match_reason text,
  outreach_status text not null default 'draft',
  emailed_at timestamptz,
  emailed_by_email text,
  seeker_cc_email text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint gre_mis_need_matches_unique unique (need_id, provider_id)
);

create table if not exists public.gre_mis_email_log (
  id uuid primary key default gen_random_uuid(),
  need_id text references public.gre_mis_needs(id) on delete set null,
  provider_id uuid references public.gre_mis_solution_providers(id) on delete set null,
  recipient_email text not null,
  cc_email text,
  subject text not null,
  body_preview text,
  sent_by_email text,
  sent_at timestamptz not null default now()
);

create or replace function public.is_gre_mis_admin()
returns boolean
language sql
stable
as $$
  select exists (
    select 1
    from public.gre_mis_admins admins
    where admins.email = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;

create or replace function public.is_gre_mis_curator()
returns boolean
language sql
stable
as $$
  select exists (
    select 1
    from public.gre_mis_curators curators
    where curators.email = lower(coalesce(auth.jwt() ->> 'email', ''))
      and curators.is_active = true
  );
$$;

drop trigger if exists gre_mis_curators_set_updated_at on public.gre_mis_curators;
create trigger gre_mis_curators_set_updated_at before update on public.gre_mis_curators
for each row execute function public.set_updated_at();

drop trigger if exists gre_mis_options_set_updated_at on public.gre_mis_options;
create trigger gre_mis_options_set_updated_at before update on public.gre_mis_options
for each row execute function public.set_updated_at();

drop trigger if exists gre_mis_needs_set_updated_at on public.gre_mis_needs;
create trigger gre_mis_needs_set_updated_at before update on public.gre_mis_needs
for each row execute function public.set_updated_at();

drop trigger if exists gre_mis_solution_providers_set_updated_at on public.gre_mis_solution_providers;
create trigger gre_mis_solution_providers_set_updated_at before update on public.gre_mis_solution_providers
for each row execute function public.set_updated_at();

drop trigger if exists gre_mis_solutions_set_updated_at on public.gre_mis_solutions;
create trigger gre_mis_solutions_set_updated_at before update on public.gre_mis_solutions
for each row execute function public.set_updated_at();

drop trigger if exists gre_mis_need_matches_set_updated_at on public.gre_mis_need_matches;
create trigger gre_mis_need_matches_set_updated_at before update on public.gre_mis_need_matches
for each row execute function public.set_updated_at();

alter table public.gre_mis_admins enable row level security;
alter table public.gre_mis_curators enable row level security;
alter table public.gre_mis_options enable row level security;
alter table public.gre_mis_needs enable row level security;
alter table public.gre_mis_need_updates enable row level security;
alter table public.gre_mis_solution_providers enable row level security;
alter table public.gre_mis_solutions enable row level security;
alter table public.gre_mis_need_matches enable row level security;
alter table public.gre_mis_email_log enable row level security;

create policy "admins read gre_mis_admins"
on public.gre_mis_admins
for select
to authenticated
using (public.is_gre_mis_admin());

create policy "admins and curators read curators"
on public.gre_mis_curators
for select
to authenticated, anon
using (is_active = true or public.is_gre_mis_admin() or public.is_gre_mis_curator());

create policy "admins manage curators"
on public.gre_mis_curators
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

create policy "admins and curators read options"
on public.gre_mis_options
for select
to authenticated, anon
using (is_active = true or public.is_gre_mis_admin());

create policy "admins manage options"
on public.gre_mis_options
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

create policy "public can create needs"
on public.gre_mis_needs
for insert
to anon, authenticated
with check (status = 'New');

create policy "admins and curators read needs"
on public.gre_mis_needs
for select
to authenticated, anon
using (public.is_gre_mis_admin() or public.is_gre_mis_curator() or true);

create policy "admins and assigned curators update needs"
on public.gre_mis_needs
for update
to authenticated
using (
  public.is_gre_mis_admin()
  or exists (
    select 1
    from public.gre_mis_curators c
    where c.id = curator_id
      and c.email = lower(coalesce(auth.jwt() ->> 'email', ''))
  )
)
with check (public.is_gre_mis_admin() or public.is_gre_mis_curator());

create policy "admins and curators read updates"
on public.gre_mis_need_updates
for select
to authenticated
using (public.is_gre_mis_admin() or public.is_gre_mis_curator());

create policy "admins and curators write updates"
on public.gre_mis_need_updates
for insert
to authenticated
with check (public.is_gre_mis_admin() or public.is_gre_mis_curator());

create policy "admins and curators read providers"
on public.gre_mis_solution_providers
for select
to authenticated, anon
using (is_active = true or public.is_gre_mis_admin());

create policy "admins manage providers"
on public.gre_mis_solution_providers
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

create policy "admins and curators read solutions"
on public.gre_mis_solutions
for select
to authenticated, anon
using (public.is_gre_mis_admin() or public.is_gre_mis_curator() or true);

create policy "admins manage solutions"
on public.gre_mis_solutions
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

create policy "admins and curators read matches"
on public.gre_mis_need_matches
for select
to authenticated
using (public.is_gre_mis_admin() or public.is_gre_mis_curator());

create policy "admins and curators manage matches"
on public.gre_mis_need_matches
for all
to authenticated
using (public.is_gre_mis_admin() or public.is_gre_mis_curator())
with check (public.is_gre_mis_admin() or public.is_gre_mis_curator());

create policy "admins read email log"
on public.gre_mis_email_log
for select
to authenticated
using (public.is_gre_mis_admin());

insert into public.gre_mis_admins (email, display_name)
values ('tanmay@greenruraleconomy.in', 'Tanmay Mukherji')
on conflict (email) do update set display_name = excluded.display_name;

insert into public.gre_mis_curators (display_name, email)
values
  ('Tanmay Mukherji', 'tanmay@greenruraleconomy.in'),
  ('Phaneesh K', 'phaneesh@greenruraleconomy.in'),
  ('Swati Singh', 'swati@greenruraleconomy.in'),
  ('Shaifali Nagar', 'shaifali@greenruraleconomy.in')
on conflict (email) do update set display_name = excluded.display_name, is_active = true;

insert into public.gre_mis_options (option_type, label, sort_order)
values
  ('status', 'New', 1),
  ('status', 'Accepted', 2),
  ('status', 'In progress', 3),
  ('status', 'Closed', 4),
  ('internal_status', 'Need solution providers', 1),
  ('internal_status', 'Connection made', 2),
  ('internal_status', 'Blocked', 3),
  ('internal_status', 'Stalled', 4),
  ('category', 'Capacity building', 1),
  ('category', 'Training', 2),
  ('category', 'Business development', 3),
  ('category', 'Business consultation', 4),
  ('category', 'Business mentoring', 5),
  ('category', 'Infrastructure', 6),
  ('category', 'Branding', 7),
  ('next_action', 'Schedule curation call', 1),
  ('next_action', 'Find provider match', 2),
  ('next_action', 'Follow up with seeker', 3),
  ('next_action', 'Broadcast to ecosystem', 4),
  ('next_action', 'Escalate to admin', 5)
on conflict do nothing;

insert into public.gre_mis_solution_providers (name, email, website, organization_type, states_served, solution_tags, notes)
values
  (
    'Rural Livestock Capacity Lab',
    'connect@rurallivestocklab.org',
    'https://greenruraleconomy.in',
    'Technical Partner',
    array['Andhra Pradesh', 'Odisha'],
    array['Capacity building', 'Training', 'Livestock health'],
    'Strong fit for livestock extension, market linkages, and local training of trainers.'
  ),
  (
    'Makhana Mechanisation Network',
    'machinery@makhananetwork.in',
    'https://greenruraleconomy.in',
    'Solution Provider',
    array['Bihar'],
    array['Business development', 'Machinery', 'Vendor'],
    'Maintains directory of machinery suppliers, indicative pricing, and after-sales partners.'
  ),
  (
    'FPO Enterprise Clinic',
    'fielddesk@fpoenterprise.org',
    'https://greenruraleconomy.in',
    'Advisory',
    array['Odisha', 'Jharkhand'],
    array['Business consultation', 'Business mentoring', 'Business development'],
    'Useful for business plans, board training, and vernacular planning support.'
  )
on conflict do nothing;

insert into public.gre_mis_needs (
  id, organization_name, contact_person, seeker_email, seeker_phone, requested_on, problem_statement, state, district,
  status, internal_status, curator_id, curation_call_date, curation_age_days, curation_notes, curated_need,
  demand_broadcast_needed, solutions_shared_count, invited_providers_count, next_action
)
values
  (
    '156',
    'RURAL RECONSTRUCTION AND DEVELOPMENT SOCIETY',
    'GangiReddy Vutukuri',
    'rrds111@gmail.com',
    '9989988008',
    '2026-04-27T16:00:53Z',
    'The community lacks proper knowledge and market linkages for goat, sheep, and livestock rearing. Support is needed for capacity building, marketing skills, and animal health services.',
    'Andhra Pradesh',
    'Nellore',
    'In progress',
    'Connection made',
    (select id from public.gre_mis_curators where email = 'shaifali@greenruraleconomy.in'),
    '2026-04-27',
    2,
    'Initial curation completed. Need provider intro pack before field-level meeting.',
    array['Capacity building', 'Training'],
    false,
    1,
    1,
    'Follow up with seeker'
  ),
  (
    '155',
    'Sarva Seva Samity Sanstha (4S-India)',
    'Gaurav',
    'gaurav4sindia@gmail.com',
    '8340643639',
    '2026-04-25T11:28:04Z',
    'Need solution providers for Makhana Harvesting Machine and Makhana Popping Machine, along with vendor details and pricing.',
    'Bihar',
    'Katihar',
    'Accepted',
    'Need solution providers',
    (select id from public.gre_mis_curators where email = 'swati@greenruraleconomy.in'),
    '2026-04-27',
    2,
    'Need verified vendors, approximate pricing, and operational guidance.',
    array['Business development', 'Vendor'],
    false,
    0,
    0,
    'Find provider match'
  ),
  (
    '154',
    'Centre for Youth and Social Development',
    'Kajal Pradhan',
    'kajalpradhan@cysd.org',
    '7608009156',
    '2026-04-18T11:46:56Z',
    'Ten FPOs in tribal Odisha need training to build business development plans in Odia for their boards of directors.',
    'Odisha',
    'Khordha',
    'In progress',
    'Need solution providers',
    (select id from public.gre_mis_curators where email = 'tanmay@greenruraleconomy.in'),
    '2026-04-27',
    2,
    'Demand is strong. Could become a reusable learning product.',
    array['Business consultation', 'Business development', 'Business mentoring'],
    true,
    0,
    2,
    'Find provider match'
  )
on conflict do nothing;

insert into public.gre_mis_need_updates (need_id, update_type, note, created_by_email)
values
  ('156', 'inbound_logged', 'Need received from GRE website.', 'tanmay@greenruraleconomy.in'),
  ('156', 'connection_made', 'One livestock capacity partner identified.', 'tanmay@greenruraleconomy.in'),
  ('155', 'curation_completed', 'Need sharpened to vendor and pricing discovery.', 'tanmay@greenruraleconomy.in'),
  ('154', 'broadcast_suggested', 'Could benefit from wider provider response.', 'tanmay@greenruraleconomy.in')
on conflict do nothing;
