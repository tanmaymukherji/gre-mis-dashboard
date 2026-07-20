-- Migration: Add GRE solution upload events audit table
-- Purpose: Append-only audit log for every GRE sync attempt (no credentials stored)
-- Created: 2026-07-18

create table if not exists public.gre_solution_upload_events (
    id uuid primary key default gen_random_uuid(),
    local_offering_id text not null references public.offerings(offering_id) on delete cascade,
    action text not null,                         -- 'create_solution' | 'publish_product' | 'submit_review' | 'approve' | 'readback_verify' | 'compensating_delete' | 'dry_run'
    step text not null,                           -- e.g., 'create_solution_request', 'create_solution_response', 'publish_product_request', 'publish_product_response', 'submit_review_request', 'approve_request', 'readback_solution', 'readback_report'
    http_status int,
    gre_request_fingerprint text,                 -- hash of request body (no secrets/headers)
    gre_response_fingerprint text,                -- hash of response body
    gre_ids jsonb,                                -- {solution_id, product_id, sku_id, channel_product_id} when available
    error_sanitised text,                         -- sanitised error message (no stack traces, no credentials)
    verification_result jsonb,                    -- {stage, expected, actual, passed}
    actor_email text,
    created_at timestamptz not null default now()
);

-- Indexes
create index if not exists idx_gre_solution_upload_events_offering on public.gre_solution_upload_events (local_offering_id);
create index if not exists idx_gre_solution_upload_events_action on public.gre_solution_upload_events (action);
create index if not exists idx_gre_solution_upload_events_created on public.gre_solution_upload_events (created_at desc);

-- RLS: Admins/moderators read-only; Service role append-only
alter table public.gre_solution_upload_events enable row level security;

create policy "gre_solution_upload_events_admin_read" on public.gre_solution_upload_events
    for select
    using (public.gre_mis_rls_is_admin_like());

create policy "gre_solution_upload_events_service_write" on public.gre_solution_upload_events
    for insert
    with check (auth.role() = 'service_role');

-- No update/delete policies — append-only by design

comment on table public.gre_solution_upload_events is 'Append-only audit log for GRE solution sync attempts. Stores only request/response fingerprints, HTTP status, and sanitised errors. Never stores credentials, session IDs, or full payloads.';