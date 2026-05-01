with base as (
  select
    id,
    lower(
      concat_ws(
        ' ',
        coalesce(problem_statement, ''),
        coalesce(curation_notes, ''),
        array_to_string(coalesce(curated_need, '{}'::text[]), ' ')
      )
    ) as haystack
  from public.gre_mis_needs
),
derived as (
  select
    id,
    array_remove(array[
      case when haystack like '%dairy%' or haystack like '%milk%' or haystack like '%milching%' or haystack like '%cow%' or haystack like '%livestock%' or haystack like '%fodder%' then 'dairy' end,
      case when haystack like '%solar%' or haystack like '%street light%' or haystack like '%streetlight%' then 'solar' end,
      case when haystack like '%wild mango%' or haystack like '%mango%' or haystack like '%amchur%' or haystack like '%ntfp%' or haystack like '%forest produce%' then 'wild mango' end,
      case when haystack like '%goat%' or haystack like '%goatery%' then 'goatery' end,
      case when haystack like '%poultry%' or haystack like '%chicken%' or haystack like '%broiler%' or haystack like '%layer%' then 'poultry' end,
      case when haystack like '%fishery%' or haystack like '%fisheries%' or haystack like '%aquaculture%' or haystack like '%fish farming%' then 'fisheries' end,
      case when haystack like '%makhana%' or haystack like '%fox nut%' then 'makhana' end,
      case when haystack like '%branding%' or haystack like '%logo%' or haystack like '%packaging%' or haystack like '%marketplace onboarding%' then 'branding' end,
      case when haystack like '%business plan%' or haystack like '%costing%' or haystack like '%pricing%' then 'business planning' end
    ], null) as rule_themes,
    array_remove(array[
      case when haystack like '%training%' or haystack like '%capacity building%' then 'training' end,
      case when haystack like '%consulting%' or haystack like '%consultancy%' or haystack like '%business consultation%' then 'consulting' end,
      case when haystack like '%mentoring%' or haystack like '%business mentoring%' then 'mentoring' end,
      case when haystack like '%technology transfer%' then 'technology transfer' end,
      case when haystack like '%advisory%' then 'advisory' end
    ], null) as rule_services,
    array_remove(array[
      case when haystack like '%training%' or haystack like '%capacity building%' then 'Manpower' end,
      case when haystack like '%consulting%' or haystack like '%consultancy%' or haystack like '%mentoring%' or haystack like '%technology transfer%' or haystack like '%video%' or haystack like '%manual%' or haystack like '%sop%' or haystack like '%blog%' then 'Method' end,
      case when haystack like '%machine%' or haystack like '%machinery%' or haystack like '%plant setup%' or haystack like '%equipment%' or haystack like '%street light%' then 'Machine' end,
      case when haystack like '%raw material%' or haystack like '%material supply%' then 'Material' end,
      case when haystack like '%products bought%' or haystack like '%market support%' or haystack like '%market report%' or haystack like '%branding%' or haystack like '%packaging%' or haystack like '%marketplace%' then 'Market' end,
      case when haystack like '%financial support%' or haystack like '%finance%' or haystack like '%funding%' or haystack like '%loan%' or haystack like '%credit%' then 'Money' end
    ], null) as rule_six_m,
    array_remove(array[
      case when haystack like '%dairy%' then 'dairy' end,
      case when haystack like '%milk%' then 'milk' end,
      case when haystack like '%cow%' or haystack like '%cows%' then 'cow' end,
      case when haystack like '%livestock%' then 'livestock' end,
      case when haystack like '%solar%' then 'solar' end,
      case when haystack like '%street light%' or haystack like '%street lights%' then 'street lights' end,
      case when haystack like '%wild mango%' then 'wild mango' end,
      case when haystack like '%mango%' then 'mango' end,
      case when haystack like '%goat%' or haystack like '%goatery%' then 'goatery' end,
      case when haystack like '%branding%' then 'branding' end,
      case when haystack like '%packaging%' then 'packaging' end,
      case when haystack like '%marketplace%' then 'marketplace' end,
      case when haystack like '%consulting%' or haystack like '%consultancy%' then 'consulting' end,
      case when haystack like '%mentoring%' then 'mentoring' end,
      case when haystack like '%training%' then 'training' end
    ], null) as rule_keywords,
    case
      when haystack like '%financial support%' or haystack like '%funding%' or haystack like '%loan%' or haystack like '%credit%' then 'finance'
      when haystack like '%machine%' or haystack like '%machinery%' or haystack like '%equipment%' or haystack like '%plant%' or haystack like '%street light%' then 'product'
      when haystack like '%manual%' or haystack like '%video%' or haystack like '%sop%' or haystack like '%blog%' then 'knowledge'
      when haystack like '%training%' or haystack like '%consulting%' or haystack like '%consultancy%' or haystack like '%mentoring%' or haystack like '%technology transfer%' or haystack like '%advisory%' then 'service'
      else null
    end as rule_need_kind
  from base
)
update public.gre_mis_needs as needs
set
  rule_thematic_hints = case
    when cardinality(derived.rule_themes) > 0 then derived.rule_themes
    else coalesce(needs.rule_thematic_hints, '{}'::text[])
  end,
  rule_service_hints = case
    when cardinality(derived.rule_services) > 0 then derived.rule_services
    else coalesce(needs.rule_service_hints, '{}'::text[])
  end,
  rule_keywords = case
    when cardinality(derived.rule_keywords) > 0 then derived.rule_keywords
    else coalesce(needs.rule_keywords, '{}'::text[])
  end,
  rule_6m_signals = case
    when cardinality(derived.rule_six_m) > 0 then derived.rule_six_m
    else coalesce(needs.rule_6m_signals, '{}'::text[])
  end,
  rule_need_kind = coalesce(derived.rule_need_kind, needs.rule_need_kind),
  ai_thematic_area = coalesce(nullif(needs.ai_thematic_area, ''), derived.rule_themes[1], needs.ai_thematic_area),
  ai_need_kind = coalesce(nullif(needs.ai_need_kind, ''), derived.rule_need_kind, needs.ai_need_kind),
  ai_service_kind = coalesce(nullif(needs.ai_service_kind, ''), derived.rule_services[1], needs.ai_service_kind),
  ai_keywords = case
    when coalesce(cardinality(needs.ai_keywords), 0) = 0 and cardinality(derived.rule_keywords) > 0 then derived.rule_keywords
    else coalesce(needs.ai_keywords, '{}'::text[])
  end,
  ai_6m_signals = case
    when coalesce(cardinality(needs.ai_6m_signals), 0) = 0 and cardinality(derived.rule_six_m) > 0 then derived.rule_six_m
    else coalesce(needs.ai_6m_signals, '{}'::text[])
  end,
  ai_summary = coalesce(nullif(needs.ai_summary, ''), left(coalesce(needs.problem_statement, ''), 500)),
  ai_engine = case
    when needs.ai_enrichment_status like 'error:%not configured%' or needs.ai_engine is null then 'rules_only'
    else needs.ai_engine
  end,
  ai_enriched_at = coalesce(needs.ai_enriched_at, now()),
  ai_enrichment_status = case
    when needs.ai_enrichment_status like 'error:%not configured%' or needs.ai_enrichment_status is null then 'rules_only'
    else needs.ai_enrichment_status
  end,
  ai_validation_flags = case
    when coalesce(cardinality(needs.ai_validation_flags), 0) = 0 and cardinality(derived.rule_themes) = 0 and cardinality(derived.rule_keywords) = 0 then array['needs_review']::text[]
    else coalesce(needs.ai_validation_flags, '{}'::text[])
  end,
  ai_validation_status = case
    when needs.ai_validation_status is not null then needs.ai_validation_status
    when cardinality(derived.rule_themes) > 0 or cardinality(derived.rule_keywords) > 0 then 'ready'
    else 'flagged'
  end,
  ai_confidence = case
    when needs.ai_confidence is not null and needs.ai_confidence > 0 then needs.ai_confidence
    when cardinality(derived.rule_themes) > 0 then 74
    when cardinality(derived.rule_keywords) > 0 then 61
    else 28
  end,
  ai_prompt_version = coalesce(needs.ai_prompt_version, '2026-05-01.gre-mis.v1'),
  ai_schema_version = coalesce(needs.ai_schema_version, 'gre-mis-need-intelligence.v1')
from derived
where needs.id = derived.id;
