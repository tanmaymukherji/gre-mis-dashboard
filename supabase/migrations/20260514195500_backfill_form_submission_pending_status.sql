update public.gre_mis_form_submissions
set
  approval_status = 'pending_admin',
  gre_sync_status = coalesce(gre_sync_status, 'pending_admin_review'),
  gre_sync_message = coalesce(gre_sync_message, 'Submission is waiting for admin review.')
where approval_status is null
  and reviewed_at is null;
