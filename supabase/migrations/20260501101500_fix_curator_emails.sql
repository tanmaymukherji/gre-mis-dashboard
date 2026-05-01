update public.gre_mis_curators
set
  email = 'agri@greenruraleconomy.in',
  updated_at = now()
where lower(display_name) = 'phaneesh k';
