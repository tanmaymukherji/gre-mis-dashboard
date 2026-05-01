alter table public.offerings
  add column if not exists source_row_signature text,
  add column if not exists rule_thematic_hints text[] default '{}'::text[],
  add column if not exists rule_service_hints text[] default '{}'::text[],
  add column if not exists rule_keywords text[] default '{}'::text[],
  add column if not exists rule_6m_signals text[] default '{}'::text[],
  add column if not exists ai_thematic_area text,
  add column if not exists ai_application_area text,
  add column if not exists ai_offering_kind text,
  add column if not exists ai_service_kind text,
  add column if not exists ai_keywords text[] default '{}'::text[],
  add column if not exists ai_6m_signals text[] default '{}'::text[],
  add column if not exists ai_summary text,
  add column if not exists ai_engine text,
  add column if not exists ai_enriched_at timestamptz,
  add column if not exists ai_enrichment_status text,
  add column if not exists ai_payload jsonb,
  add column if not exists ai_prompt_version text,
  add column if not exists ai_schema_version text;

with base as (
  select
    offering_id,
    lower(
      concat_ws(
        ' ',
        coalesce(offering_name, ''),
        coalesce(offering_group, ''),
        coalesce(offering_type, ''),
        coalesce(offering_category, ''),
        coalesce(primary_application, ''),
        coalesce(primary_valuechain, ''),
        coalesce(array_to_string(applications, ' '), ''),
        coalesce(array_to_string(valuechains, ' '), ''),
        coalesce(array_to_string(tags, ' '), ''),
        coalesce(domain_6m, ''),
        coalesce(about_offering_text, '')
      )
    ) as haystack
  from public.offerings
),
derived as (
  select
    offering_id,
    case
      when haystack like '%service%' or haystack like '%consulting%' or haystack like '%mentoring%' or haystack like '%training%' or haystack like '%technology transfer%' then 'service'
      when haystack like '%product%' or haystack like '%machine%' or haystack like '%machinery%' or haystack like '%equipment%' or haystack like '%raw material%' then 'product'
      when haystack like '%knowledge%' or haystack like '%manual%' or haystack like '%video%' or haystack like '%sop%' or haystack like '%blog%' then 'knowledge'
      else null
    end as offering_kind,
    case
      when haystack like '%training%' or haystack like '%capacity building%' then 'training'
      when haystack like '%consulting%' or haystack like '%consultancy%' then 'consulting'
      when haystack like '%mentoring%' then 'mentoring'
      when haystack like '%technology transfer%' then 'technology transfer'
      when haystack like '%advisory%' then 'advisory'
      else null
    end as service_kind,
    array_remove(array[
      case when haystack like '%dairy%' or haystack like '%milk%' or haystack like '%cow%' or haystack like '%fodder%' then 'dairy' end,
      case when haystack like '%solar%' or haystack like '%street light%' or haystack like '%streetlight%' then 'solar' end,
      case when haystack like '%wild mango%' or haystack like '%mango%' or haystack like '%ntfp%' then 'wild mango' end,
      case when haystack like '%goat%' or haystack like '%goatery%' then 'goatery' end,
      case when haystack like '%poultry%' or haystack like '%chicken%' then 'poultry' end,
      case when haystack like '%fish%' or haystack like '%aquaculture%' then 'fisheries' end,
      case when haystack like '%makhana%' or haystack like '%fox nut%' then 'makhana' end,
      case when haystack like '%branding%' or haystack like '%logo%' or haystack like '%packaging%' then 'branding' end,
      case when haystack like '%bamboo%' then 'bamboo' end,
      case when haystack like '%moringa%' then 'moringa' end
    ], null) as themes,
    array_remove(array[
      case when haystack like '%training%' or haystack like '%capacity building%' then 'Manpower' end,
      case when haystack like '%consulting%' or haystack like '%consultancy%' or haystack like '%mentoring%' or haystack like '%technology transfer%' or haystack like '%manual%' or haystack like '%video%' or haystack like '%sop%' or haystack like '%blog%' or haystack like '%advisory%' then 'Method' end,
      case when haystack like '%machine%' or haystack like '%machinery%' or haystack like '%equipment%' or haystack like '%plant setup%' then 'Machine' end,
      case when haystack like '%raw material%' or haystack like '%material supply%' then 'Material' end,
      case when haystack like '%market%' or haystack like '%branding%' or haystack like '%packaging%' or haystack like '%marketplace%' then 'Market' end,
      case when haystack like '%financial%' or haystack like '%finance%' or haystack like '%funding%' or haystack like '%credit%' or haystack like '%loan%' then 'Money' end
    ], null) as six_m
  from base
)
update public.offerings as offerings
set
  rule_thematic_hints = case when cardinality(derived.themes) > 0 then derived.themes else coalesce(offerings.rule_thematic_hints, '{}'::text[]) end,
  rule_service_hints = case when derived.service_kind is not null then array[derived.service_kind] else coalesce(offerings.rule_service_hints, '{}'::text[]) end,
  rule_keywords = case when cardinality(derived.themes) > 0 then derived.themes else coalesce(offerings.rule_keywords, '{}'::text[]) end,
  rule_6m_signals = case when cardinality(derived.six_m) > 0 then derived.six_m else coalesce(offerings.rule_6m_signals, '{}'::text[]) end,
  ai_thematic_area = coalesce(nullif(offerings.ai_thematic_area, ''), derived.themes[1], offerings.primary_application, offerings.primary_valuechain),
  ai_application_area = coalesce(nullif(offerings.ai_application_area, ''), offerings.primary_application, offerings.primary_valuechain),
  ai_offering_kind = coalesce(nullif(offerings.ai_offering_kind, ''), derived.offering_kind, offerings.ai_offering_kind),
  ai_service_kind = coalesce(nullif(offerings.ai_service_kind, ''), derived.service_kind, offerings.ai_service_kind),
  ai_keywords = case when coalesce(cardinality(offerings.ai_keywords), 0) = 0 then coalesce(derived.themes, '{}'::text[]) else offerings.ai_keywords end,
  ai_6m_signals = case when coalesce(cardinality(offerings.ai_6m_signals), 0) = 0 then coalesce(derived.six_m, '{}'::text[]) else offerings.ai_6m_signals end,
  ai_summary = coalesce(nullif(offerings.ai_summary, ''), left(coalesce(offerings.about_offering_text, offerings.offering_name, ''), 500)),
  ai_engine = coalesce(offerings.ai_engine, 'rules_only'),
  ai_enriched_at = coalesce(offerings.ai_enriched_at, now()),
  ai_enrichment_status = coalesce(offerings.ai_enrichment_status, 'rules_only'),
  ai_prompt_version = coalesce(offerings.ai_prompt_version, '2026-05-01.gre-mis.v1'),
  ai_schema_version = coalesce(offerings.ai_schema_version, 'gre-mis-need-intelligence.v1.offering')
from derived
where offerings.offering_id = derived.offering_id;
