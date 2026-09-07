-- AI-UBIKE durable shared learning pool
-- Run this once in Supabase SQL Editor.

create table if not exists public.ai_learning_records (
    record_id text primary key,
    observed_at_epoch double precision not null default 0,
    payload jsonb not null,
    updated_at timestamptz not null default now()
);

create index if not exists ai_learning_records_observed_at_idx
    on public.ai_learning_records (observed_at_epoch desc);

alter table public.ai_learning_records enable row level security;

-- Recommended runtime configuration:
-- Store the Supabase service-role key only in Streamlit Secrets, never in GitHub.
-- The service-role key bypasses RLS server-side, so no public insert/select policy
-- is required. If you intentionally use an anon key instead, create restrictive
-- RLS policies before enabling it.
--
-- Streamlit Secrets example:
-- [supabase]
-- url = "https://YOUR_PROJECT.supabase.co"
-- service_role_key = "YOUR_SERVICE_ROLE_KEY"
-- ai_table = "ai_learning_records"
