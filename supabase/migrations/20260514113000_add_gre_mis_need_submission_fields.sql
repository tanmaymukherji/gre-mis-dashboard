alter table public.gre_mis_needs
  add column if not exists deployment_locations text[] not null default '{}',
  add column if not exists submitted_keywords text[] not null default '{}',
  add column if not exists submitted_thematic_area text,
  add column if not exists submitted_offering_category text,
  add column if not exists submitted_offering_type text;
