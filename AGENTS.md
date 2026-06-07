# GRE MIS Dashboard — Agent Guide

## Type

Vanilla HTML/CSS/JS static site + independent Supabase backend.  
**No package.json, no build step, no npm/node, no tests, no lint, no typecheck.**  

## Key files

| File | Purpose |
|---|---|
| `index.html` | Main dashboard (overview, operations, admin, solution/need intake views) |
| `admin.html` | Standalone admin sync page (shared `app.js`) |
| `offering-detail.html` | Standalone offering detail page (self-contained `<script type="module">`) |
| `app.js` | All dashboard logic (~9K lines) |
| `styles.css` | All styles (~2K lines) |
| `config.js` | Runtime config: `SUPABASE_URL`, `SUPABASE_ANON_KEY`, `ADMIN_FUNCTION`, `MAPPLS_MAP_KEY` |
| `config.example.js` | Template — copy to `config.js` and fill in values |
| `supabase/migrations/` | 84 SQL migrations, run in date order |
| `supabase/functions/gre-mis-admin/index.ts` | Deno edge function (~9.5K lines, single `Deno.serve` handler) |
| `.github/workflows/deploy-pages.yml` | CI: copies specific files to `_site/`, deploys to GitHub Pages |

## Auth

GramEEE shared SSO via cookies (`grameee_access_token`, `grameee_user_summary`).  
Scripts loaded from `https://grameee.org/supabase-config.js` and `https://grameee.org/auth.js`.  
Roles: `admin`, `moderator`, `curator`. Admin desk visible only to admin/moderator.

## Setup for a new Supabase project

1. Copy `config.example.js` → `config.js`, fill `SUPABASE_URL` and `SUPABASE_ANON_KEY`
2. Run all `.sql` files in `supabase/migrations/` in filename order
3. Deploy edge function: `supabase functions deploy gre-mis-admin`
4. Set function secrets: `GRE_MIS_SERVICE_ROLE_KEY`, `GMAIL_CLIENT_ID`, `GMAIL_CLIENT_SECRET`, `GMAIL_REFRESH_TOKEN`, `GMAIL_SENDER_EMAIL`

## Deployment

GitHub Actions pushes to `main` → deploys to GitHub Pages.  
Deployed files: `index.html`, `admin.html`, `offering-detail.html`, `styles.css`, `app.js`, `config.js`, `assets/`.  
**Anything else must be added to `deploy-pages.yml` if it needs to reach production.**

## Conventions

- Cache busting via `?v=YYYYMMDD<tag>` query params on script/link tags
- The `app.js` global `state` object holds all UI state (see ~line 63)
- Supabase client uses anon key with RLS; edge function uses service_role key
- Edge function handles: inbound sync, AI enrichment, email sending, chatbot data refresh, user management
- Offering detail page (`offering-detail.html`) queries `offerings`, `traders`, `solutions` tables from shared GramEEE schema
- The `grameee-gre-bar.js` is an older auth bar script; current pages use `grameee.org/auth.js`
