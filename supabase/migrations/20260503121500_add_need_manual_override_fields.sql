alter table public.gre_mis_needs
  add column if not exists override_thematic_area text,
  add column if not exists override_application_area text,
  add column if not exists override_need_kind text,
  add column if not exists override_service_kind text,
  add column if not exists override_keywords text[] default '{}'::text[],
  add column if not exists override_6m_signals text[] default '{}'::text[],
  add column if not exists override_summary text,
  add column if not exists override_source text,
  add column if not exists override_conflict_note text,
  add column if not exists override_updated_at timestamptz,
  add column if not exists override_updated_by text;
