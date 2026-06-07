alter table public.gre_mis_needs
  add column if not exists funding_mechanism text,
  add column if not exists seeker_provider_agreement text,
  add column if not exists solution_deployment_status text,
  add column if not exists closure_date date,
  add column if not exists feedback_about_seeker text,
  add column if not exists feedback_about_provider text;

alter table public.gre_mis_update_requests
  add column if not exists proposed_curated_need text[],
  add column if not exists proposed_funding_mechanism text,
  add column if not exists proposed_seeker_provider_agreement text,
  add column if not exists proposed_solution_deployment_status text,
  add column if not exists proposed_closure_date date,
  add column if not exists proposed_feedback_about_seeker text,
  add column if not exists proposed_feedback_about_provider text;
