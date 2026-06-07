create or replace function public.is_gre_mis_admin()
returns boolean
language sql
stable
security definer
set search_path = public
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
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.gre_mis_curators curators
    where curators.email = lower(coalesce(auth.jwt() ->> 'email', ''))
      and curators.is_active = true
  );
$$;

grant execute on function public.is_gre_mis_admin() to anon, authenticated, service_role;
grant execute on function public.is_gre_mis_curator() to anon, authenticated, service_role;
