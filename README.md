# GRE MIS Dashboard

Standalone GRE operations dashboard for inbound needs, curator allocation, provider matching, broadcast decisions, and closure tracking.

## What this repository includes

- A static GitHub Pages-friendly frontend in [index.html](C:/github/gre-mis-dashboard/index.html), [styles.css](C:/github/gre-mis-dashboard/styles.css), and [app.js](C:/github/gre-mis-dashboard/app.js)
- A separate Supabase schema in [supabase/migrations/001_create_gre_mis.sql](C:/github/gre-mis-dashboard/supabase/migrations/001_create_gre_mis.sql)
- A dedicated admin edge function in [supabase/functions/gre-mis-admin/index.ts](C:/github/gre-mis-dashboard/supabase/functions/gre-mis-admin/index.ts)
- A copied GRE logo asset in [assets/gre-logo.png](C:/github/gre-mis-dashboard/assets/gre-logo.png)

## Product scope covered in this first build

- Admin dashboard and Curator workbench in the same interface
- Need status analytics, curator allocation, aging / stuck cases, and top categories
- Curator assignment from the needs list
- Solution-provider matching surface for GRE solutions and provider data
- Email trigger flow to send the problem statement to a selected solution provider, with the seeker on copy
- Add / edit missing option values for statuses, internal statuses, categories, and next actions
- Roadmap-ready placeholder area for a future seeker self-service challenge portal

## Deployment shape

- Frontend: GitHub Pages
- Backend: Independent Supabase project and dataset
- Admin approvals: same GRE admin email model, but isolated tables and functions for this MIS

## Setup

1. Copy [config.example.js](C:/github/gre-mis-dashboard/config.example.js) to `config.js`.
2. Fill in `SUPABASE_URL` and `SUPABASE_ANON_KEY`.
3. Create a brand-new Supabase project for this repository.
4. Run the SQL migration from [supabase/migrations/001_create_gre_mis.sql](C:/github/gre-mis-dashboard/supabase/migrations/001_create_gre_mis.sql).
5. Deploy the edge function from [supabase/functions/gre-mis-admin/index.ts](C:/github/gre-mis-dashboard/supabase/functions/gre-mis-admin/index.ts).
6. Add the following function secrets in Supabase:
   - `GRE_MIS_SERVICE_ROLE_KEY`
   - `GMAIL_CLIENT_ID`
   - `GMAIL_CLIENT_SECRET`
   - `GMAIL_REFRESH_TOKEN`
   - `GMAIL_SENDER_EMAIL`
7. Host the repository on GitHub Pages.

## Email flow

The provider outreach action is implemented in the edge function. It sends from the configured GRE mailbox to the selected provider and keeps the seeker in `Cc`.

## Important isolation rule

This repository is intended to be independent from existing GRE codebases and must use its own Supabase project, tables, and functions to avoid overwriting other datasets.

