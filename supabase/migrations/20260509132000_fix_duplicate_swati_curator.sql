update public.gre_mis_curators
set
  email = 'swati.singh@rainmatter.org',
  phone = '9586599933',
  is_active = true,
  gre_sync_status = 'synced',
  gre_sync_message = 'Curator record reconciled to GRE source-of-truth.',
  gre_synced_at = now(),
  updated_at = now()
where id = '595fb510-16f5-466b-a043-06e4d098e79b';

delete from public.gre_mis_curators
where id = 'a6f9b43c-faea-4f90-9b9f-a56bc60c9e34';
