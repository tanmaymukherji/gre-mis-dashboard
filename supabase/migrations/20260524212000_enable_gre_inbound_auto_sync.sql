insert into public.gre_mis_settings (key, value_text, value_json)
values ('inbound_auto_sync', null, '{"enabled": true}'::jsonb)
on conflict (key)
do update
set value_json = '{"enabled": true}'::jsonb,
    updated_at = timezone('utc', now());
