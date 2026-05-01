alter table public.gre_mis_needs
add column if not exists rule_thematic_hints text[] not null default '{}',
add column if not exists rule_service_hints text[] not null default '{}',
add column if not exists rule_keywords text[] not null default '{}',
add column if not exists rule_6m_signals text[] not null default '{}',
add column if not exists rule_need_kind text,
add column if not exists ai_validation_flags text[] not null default '{}',
add column if not exists ai_validation_status text,
add column if not exists ai_confidence integer not null default 0,
add column if not exists ai_prompt_version text,
add column if not exists ai_schema_version text;
