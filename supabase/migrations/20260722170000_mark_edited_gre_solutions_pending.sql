-- Mark locally edited offerings as awaiting a GRE refresh.
-- The existing GRE identity is retained until its absence is verified by the
-- resync action, preventing accidental duplicate creation.

update public.gre_solution_links as link
set
  upload_state = 'update_pending',
  manual_review_reason = 'The GramEEE solution was edited after its last verified GRE sync. Delete the previous GRE record before recreating it.',
  gre_status_summary = coalesce(link.gre_status_summary, '{}'::jsonb) || jsonb_build_object(
    'localEditPending', true,
    'localEditedAt', offering.raw_payload -> 'last_manual_edit' ->> 'updated_at',
    'previousVerifiedAt', link.verified_at
  )
from public.offerings as offering
where offering.offering_id = link.local_offering_id
  and link.upload_state = 'synced'
  and jsonb_typeof(offering.raw_payload -> 'last_manual_edit') = 'object'
  and coalesce(offering.raw_payload -> 'last_manual_edit' ->> 'updated_at', '')
      ~ '^\d{4}-\d{2}-\d{2}T'
  and (
    link.verified_at is null
    or (offering.raw_payload -> 'last_manual_edit' ->> 'updated_at')::timestamptz > link.verified_at
  );

comment on column public.gre_solution_links.upload_state is
  'GRE sync state, including update_pending when a verified local offering was subsequently edited.';
