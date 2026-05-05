create extension if not exists pgcrypto;

create table if not exists public.gre_mis_users (
  id uuid primary key default gen_random_uuid(),
  username text not null unique,
  first_name text not null,
  full_name text not null,
  email text not null unique,
  phone text,
  password_hash text not null,
  role text not null default 'user',
  is_active boolean not null default true,
  must_change_password boolean not null default false,
  last_login_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint gre_mis_users_role_check check (role in ('admin', 'curator', 'user')),
  constraint gre_mis_users_email_lowercase check (email = lower(email)),
  constraint gre_mis_users_username_lowercase check (username = lower(username))
);

create table if not exists public.gre_mis_web_sessions (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.gre_mis_users(id) on delete cascade,
  token_hash text not null unique,
  expires_at timestamptz not null,
  created_at timestamptz not null default now(),
  last_used_at timestamptz not null default now()
);

create table if not exists public.gre_mis_password_reset_requests (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.gre_mis_users(id) on delete cascade,
  email text not null,
  code_hash text not null,
  expires_at timestamptz not null,
  used_at timestamptz,
  created_at timestamptz not null default now(),
  constraint gre_mis_password_reset_email_lowercase check (email = lower(email))
);

alter table public.gre_mis_curators
  add column if not exists first_name text,
  add column if not exists phone text,
  add column if not exists user_id uuid references public.gre_mis_users(id) on delete set null,
  add column if not exists gre_sync_status text,
  add column if not exists gre_sync_message text,
  add column if not exists gre_synced_at timestamptz;

drop trigger if exists gre_mis_users_set_updated_at on public.gre_mis_users;
create trigger gre_mis_users_set_updated_at before update on public.gre_mis_users
for each row execute function public.set_updated_at();

alter table public.gre_mis_users enable row level security;
alter table public.gre_mis_web_sessions enable row level security;
alter table public.gre_mis_password_reset_requests enable row level security;

create policy "admins manage gre mis users"
on public.gre_mis_users
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

create policy "admins manage gre mis web sessions"
on public.gre_mis_web_sessions
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

create policy "admins manage gre mis password reset requests"
on public.gre_mis_password_reset_requests
for all
to authenticated
using (public.is_gre_mis_admin())
with check (public.is_gre_mis_admin());

create or replace function public.gre_mis_user_password_matches(
  p_identifier text,
  p_password text
)
returns boolean
language sql
security definer
set search_path = public, extensions
as $$
  select exists (
    select 1
    from public.gre_mis_users
    where is_active = true
      and (
        username = lower(trim(coalesce(p_identifier, '')))
        or email = lower(trim(coalesce(p_identifier, '')))
      )
      and password_hash = crypt(coalesce(p_password, ''), password_hash)
  );
$$;

create or replace function public.gre_mis_user_set_password(
  p_user_id uuid,
  p_password text,
  p_must_change_password boolean default false
)
returns void
language sql
security definer
set search_path = public, extensions
as $$
  update public.gre_mis_users
  set
    password_hash = crypt(coalesce(p_password, ''), gen_salt('bf')),
    must_change_password = coalesce(p_must_change_password, false),
    updated_at = now()
  where id = p_user_id;
$$;

create or replace function public.gre_mis_register_user(
  p_username text,
  p_first_name text,
  p_full_name text,
  p_email text,
  p_phone text,
  p_password text
)
returns uuid
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  v_user_id uuid;
begin
  insert into public.gre_mis_users (
    username,
    first_name,
    full_name,
    email,
    phone,
    password_hash,
    role,
    is_active,
    must_change_password
  )
  values (
    lower(trim(p_username)),
    trim(p_first_name),
    trim(p_full_name),
    lower(trim(p_email)),
    nullif(trim(coalesce(p_phone, '')), ''),
    crypt(coalesce(p_password, ''), gen_salt('bf')),
    'user',
    true,
    false
  )
  returning id into v_user_id;

  return v_user_id;
end;
$$;

insert into public.gre_mis_users (
  username,
  first_name,
  full_name,
  email,
  phone,
  password_hash,
  role,
  is_active,
  must_change_password
)
values (
  'admin',
  'Admin',
  'GRE Admin',
  'tanmay@greenruraleconomy.in',
  null,
  crypt('gre1234', gen_salt('bf')),
  'admin',
  true,
  false
)
on conflict (username) do update
set
  full_name = excluded.full_name,
  email = excluded.email,
  role = 'admin',
  is_active = true;

insert into public.gre_mis_users (
  username,
  first_name,
  full_name,
  email,
  phone,
  password_hash,
  role,
  is_active,
  must_change_password
)
select
  lower(split_part(display_name, ' ', 1)) as username,
  split_part(display_name, ' ', 1) as first_name,
  display_name as full_name,
  lower(email) as email,
  phone,
  crypt('gre@1234', gen_salt('bf')),
  'curator',
  true,
  true
from public.gre_mis_curators
where is_active = true
  and lower(email) <> 'tanmay@greenruraleconomy.in'
on conflict (email) do update
set
  username = excluded.username,
  first_name = excluded.first_name,
  full_name = excluded.full_name,
  phone = excluded.phone,
  role = 'curator',
  is_active = true;

update public.gre_mis_curators cur
set
  first_name = coalesce(cur.first_name, split_part(cur.display_name, ' ', 1)),
  user_id = usr.id,
  gre_sync_status = coalesce(cur.gre_sync_status, 'pending')
from public.gre_mis_users usr
where lower(cur.email) = lower(usr.email);
