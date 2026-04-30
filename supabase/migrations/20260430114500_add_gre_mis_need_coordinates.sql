alter table public.gre_mis_needs
add column if not exists latitude double precision,
add column if not exists longitude double precision,
add column if not exists geocoded_label text,
add column if not exists geocode_status text,
add column if not exists geocoded_at timestamptz;
