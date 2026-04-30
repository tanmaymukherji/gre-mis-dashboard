alter table public.gre_mis_needs
add column if not exists six_m_signals text[] not null default '{}';

create or replace function public.compute_gre_mis_six_m_signals(
  curated_need text[],
  problem_statement text,
  curation_notes text
)
returns text[]
language sql
immutable
as $$
  with source_text as (
    select lower(
      coalesce(array_to_string(curated_need, ' '), '') || ' ' ||
      coalesce(problem_statement, '') || ' ' ||
      coalesce(curation_notes, '')
    ) as text_value
  ),
  matched as (
    select 'Manpower'::text as label
    from source_text
    where text_value like '%training%' or text_value like '%capacity building%'

    union

    select 'Method'::text as label
    from source_text
    where text_value like '%consulting%'
      or text_value like '%consultancy%'
      or text_value like '%business consultation%'
      or text_value like '%mentoring%'
      or text_value like '%business mentoring%'
      or text_value like '%technology transfer%'
      or text_value like '%technology%'
      or text_value like '%video%'
      or text_value like '%manual%'
      or text_value like '%sop%'
      or text_value like '%blog%'
      or text_value like '%connect collaborate%'

    union

    select 'Machine'::text as label
    from source_text
    where text_value like '%machinery%'
      or text_value like '%machine%'
      or text_value like '%plant setup%'

    union

    select 'Material'::text as label
    from source_text
    where text_value like '%raw material%'
      or text_value like '%raw materials%'
      or text_value like '%material supply%'

    union

    select 'Market'::text as label
    from source_text
    where text_value like '%products bought%'
      or text_value like '%market support%'
      or text_value like '%market reports%'
      or text_value like '%market report%'
      or text_value like '%business development%'
      or text_value like '%branding%'
      or text_value like '%packaging%'

    union

    select 'Money'::text as label
    from source_text
    where text_value like '%financial support%'
      or text_value like '%finance%'
      or text_value like '%funding%'
      or text_value like '%investment%'
      or text_value like '%investments%'
      or text_value like '%credit%'
  )
  select coalesce(array_agg(label order by label), '{}'::text[])
  from matched;
$$;

create or replace function public.set_gre_mis_need_six_m_signals()
returns trigger
language plpgsql
as $$
begin
  new.six_m_signals := public.compute_gre_mis_six_m_signals(new.curated_need, new.problem_statement, new.curation_notes);
  return new;
end;
$$;

drop trigger if exists gre_mis_needs_set_six_m_signals on public.gre_mis_needs;
create trigger gre_mis_needs_set_six_m_signals
before insert or update of curated_need, problem_statement, curation_notes
on public.gre_mis_needs
for each row execute function public.set_gre_mis_need_six_m_signals();

update public.gre_mis_needs
set six_m_signals = public.compute_gre_mis_six_m_signals(curated_need, problem_statement, curation_notes)
where true;
