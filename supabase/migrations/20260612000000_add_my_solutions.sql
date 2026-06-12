-- My Solutions feature: support signed_in_edit submissions, offering views, and conflict detection

-- 1. Expand source_mode check to include 'signed_in_edit'
alter table public.gre_mis_form_submissions
  drop constraint if exists gre_mis_form_submissions_source_mode_check;

alter table public.gre_mis_form_submissions
  add constraint gre_mis_form_submissions_source_mode_check
    check (source_mode in ('shared_link', 'signed_in', 'signed_in_edit'));

-- 2. Add edit lineage and conflict columns
alter table public.gre_mis_form_submissions
  add column if not exists parent_submission_id uuid
    references public.gre_mis_form_submissions(id)
    on delete set null;

alter table public.gre_mis_form_submissions
  add column if not exists local_edit_version int not null default 0;

alter table public.gre_mis_form_submissions
  add column if not exists conflict_with_gre_sync boolean not null default false;

create index if not exists gre_mis_form_submissions_parent_idx
  on public.gre_mis_form_submissions (parent_submission_id);

-- 3. Create offering views tracking table
create table if not exists public.gre_mis_offering_views (
  id uuid primary key default gen_random_uuid(),
  offering_id text not null,
  submission_id uuid references public.gre_mis_form_submissions(id) on delete cascade,
  viewer_user_id uuid references public.gre_mis_users(id) on delete set null,
  viewer_email text,
  viewer_organisation text,
  viewer_name text,
  viewed_at timestamptz not null default now()
);

create index if not exists gre_mis_offering_views_offering_idx
  on public.gre_mis_offering_views (offering_id, viewed_at desc);

create index if not exists gre_mis_offering_views_submission_idx
  on public.gre_mis_offering_views (submission_id, viewed_at desc);

-- 4. RLS: allow users to read their own form submissions
do $$ begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'public'
      and tablename = 'gre_mis_form_submissions'
      and policyname = 'Users can view own submissions'
  ) then
    create policy "Users can view own submissions"
      on public.gre_mis_form_submissions
      for select
      using (
        submitter_user_id = auth.uid()
        or submitter_email = lower(trim(concat(auth.jwt() ->> 'email')))
        or submitter_email = lower(trim(concat(auth.jwt() -> 'user_metadata' ->> 'email')))
      );
  end if;
end $$;

-- 5. RLS: allow users to view offering views for their own submissions
do $$ begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'public'
      and tablename = 'gre_mis_offering_views'
      and policyname = 'Users can view offering views for own submissions'
  ) then
    create policy "Users can view offering views for own submissions"
      on public.gre_mis_offering_views
      for select
      using (
        viewer_user_id = auth.uid()
        or viewer_email = lower(trim(concat(auth.jwt() ->> 'email')))
        or submission_id in (
          select id from public.gre_mis_form_submissions
          where submitter_user_id = auth.uid()
             or submitter_email = lower(trim(concat(auth.jwt() ->> 'email')))
        )
      );
  end if;
end $$;

-- 6. Enable RLS on the new table
alter table public.gre_mis_offering_views enable row level security;
