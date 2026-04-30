insert into public.gre_mis_curators (display_name, email, is_active)
values
  ('Prashant Mehra', 'prashant@platformcommons.com', true),
  ('Shashank Deora', 'shashank@commongroundinitiative.in', true)
on conflict (email) do update
set
  display_name = excluded.display_name,
  is_active = true,
  updated_at = now();
