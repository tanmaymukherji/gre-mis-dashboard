create or replace function public.gre_mis_rls_is_admin_like()
returns boolean
language sql
stable
security invoker
set search_path = public
as $$
  select (auth.jwt() -> 'app_metadata' ->> 'grameee_role') in ('admin', 'moderator')
$$;

revoke all on function public.gre_mis_rls_is_admin_like() from public, anon;
grant execute on function public.gre_mis_rls_is_admin_like() to authenticated, service_role;

drop policy if exists "gre_solution_links_admin_read" on public.gre_solution_links;
drop policy if exists "gre_solution_links_service_write" on public.gre_solution_links;
drop policy if exists "gre_solution_upload_events_admin_read" on public.gre_solution_upload_events;
drop policy if exists "gre_solution_upload_events_service_write" on public.gre_solution_upload_events;

create policy "gre_solution_links_admin_read"
on public.gre_solution_links
for select
to authenticated
using (public.gre_mis_rls_is_admin_like());

create policy "gre_solution_links_service_write"
on public.gre_solution_links
for all
to service_role
using (true)
with check (true);

create policy "gre_solution_upload_events_admin_read"
on public.gre_solution_upload_events
for select
to authenticated
using (public.gre_mis_rls_is_admin_like());

create policy "gre_solution_upload_events_service_write"
on public.gre_solution_upload_events
for insert
to service_role
with check (true);
