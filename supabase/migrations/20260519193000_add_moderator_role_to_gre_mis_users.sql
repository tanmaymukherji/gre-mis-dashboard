alter table public.gre_mis_users
  drop constraint if exists gre_mis_users_role_check;

alter table public.gre_mis_users
  add constraint gre_mis_users_role_check
  check (role in ('admin', 'moderator', 'curator', 'user'));
