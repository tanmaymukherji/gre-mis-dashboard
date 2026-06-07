create table if not exists public.gre_mis_settings (
  key text primary key,
  value_text text,
  value_json jsonb not null default '{}'::jsonb,
  updated_by_email text,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

drop trigger if exists gre_mis_settings_set_updated_at on public.gre_mis_settings;
create trigger gre_mis_settings_set_updated_at
before update on public.gre_mis_settings
for each row execute function public.set_updated_at();

alter table public.gre_mis_settings enable row level security;

drop policy if exists "Allow read gre_mis_settings" on public.gre_mis_settings;
create policy "Allow read gre_mis_settings"
on public.gre_mis_settings
for select
to authenticated, anon
using (true);

insert into public.gre_mis_settings (key, value_text, value_json)
values
(
  'provider_intro_template',
  'Hello {{providerName}},\n\nWe are reaching out to you from GRE platform to connect you with {{seekerLabel}}, marked in copy of this mail. They have a need for {{problemStatement}} and your solution of {{viewLink}} may be of interest to them. We would suggest you to connect mutually and take this forward. Do reach out to us if you would like us to help facilitate the conversation.\n\nRegards,\nTeam GRE',
  '{}'::jsonb
),
(
  'curator_forward_template',
  'Hello {{assignedCuratorName}},\n\nFor the needs of {{seekerLabel}} for {{problemStatement}}, i feel the solution {{viewLink}} by {{providerName}} may be of interest to you. Kindly review and do the needful.\n\nRegards,\n{{actorName}}',
  '{}'::jsonb
),
(
  'inbound_auto_sync',
  null,
  '{"enabled": false}'::jsonb
)
on conflict (key) do nothing;
