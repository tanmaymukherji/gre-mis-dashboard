alter table public.gre_mis_users
  add column if not exists gre_pending_role text,
  add column if not exists gre_activation_mod_key text,
  add column if not exists gre_activation_requested_at timestamptz;

update public.gre_mis_users
set
  gre_pending_role = null,
  gre_activation_mod_key = null
where true;
