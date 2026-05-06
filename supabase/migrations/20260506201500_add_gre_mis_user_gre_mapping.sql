alter table public.gre_mis_users
  add column if not exists gre_user_id bigint,
  add column if not exists gre_login_name text,
  add column if not exists gre_sync_status text,
  add column if not exists gre_sync_message text,
  add column if not exists gre_synced_at timestamptz;

update public.gre_mis_users
set
  gre_login_name = coalesce(gre_login_name, nullif(trim(phone), ''), lower(trim(email))),
  gre_sync_status = coalesce(gre_sync_status, 'pending'),
  gre_sync_message = coalesce(gre_sync_message, 'Awaiting GRE user mapping.')
where true;
