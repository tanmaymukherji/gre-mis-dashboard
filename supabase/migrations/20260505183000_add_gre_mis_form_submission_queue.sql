create table if not exists public.gre_mis_form_submissions (
  id uuid primary key default gen_random_uuid(),
  submission_type text not null check (submission_type in ('need', 'solution')),
  source_mode text not null default 'shared_link' check (source_mode in ('shared_link', 'signed_in')),
  submitter_name text,
  submitter_email text,
  submitter_phone text,
  submitter_user_id uuid references public.gre_mis_users(id) on delete set null,
  organization_name text,
  existing_trader_id text,
  existing_trader_name text,
  org_exists_on_gre boolean not null default false,
  payload jsonb not null default '{}'::jsonb,
  approval_status text not null default 'pending_admin' check (approval_status in ('pending_admin', 'approved', 'rejected')),
  admin_review_notes text,
  reviewed_by_email text,
  reviewed_at timestamptz,
  synced_to_gre boolean not null default false,
  gre_sync_status text,
  gre_sync_message text,
  target_need_id text,
  target_solution_id text,
  share_context text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists gre_mis_form_submissions_pending_idx
  on public.gre_mis_form_submissions (approval_status, submission_type, created_at desc);

drop trigger if exists gre_mis_form_submissions_set_updated_at on public.gre_mis_form_submissions;
create trigger gre_mis_form_submissions_set_updated_at
before update on public.gre_mis_form_submissions
for each row execute function public.set_updated_at();
