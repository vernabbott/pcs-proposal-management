-- Initial production data contract for PCS Proposal Management and PilotPoint IQ.
-- Apply only after reviewing role names and the PCS authentication model.

create extension if not exists pgcrypto;

create table if not exists public.properties (
  id uuid primary key default gen_random_uuid(),
  canonical_key text not null unique,
  normalized_address text not null,
  address text,
  city text,
  state text,
  zip_code text,
  county text,
  parcel_number text,
  latitude double precision,
  longitude double precision,
  roof_area_sqft numeric,
  roof_squares numeric,
  year_built integer,
  effective_year_built integer,
  age_estimate_year integer,
  age_estimate_years integer,
  age_estimate_source text,
  age_estimate_as_of_date date,
  data jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.roof_intelligence_jobs (
  id uuid primary key default gen_random_uuid(),
  job_type text not null check (job_type in ('individual_address', 'zip_batch')),
  requested_by uuid references auth.users(id),
  status text not null default 'queued'
    check (status in ('queued', 'running', 'completed', 'completed_with_errors', 'failed', 'cancelled')),
  stage text not null default 'queued',
  input jsonb not null default '{}'::jsonb,
  normalized_address text,
  zip_code text,
  report_limit integer,
  minimum_roof_size integer,
  minimum_age integer,
  roof_types jsonb not null default '[]'::jsonb,
  candidate_count integer not null default 0,
  completed_count integer not null default 0,
  failed_count integer not null default 0,
  skipped_count integer not null default 0,
  remaining_count integer not null default 0,
  error_code text,
  error_message text,
  retryable boolean not null default false,
  worker_version text,
  queued_at timestamptz not null default now(),
  started_at timestamptz,
  finished_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists roof_intelligence_jobs_status_idx
  on public.roof_intelligence_jobs (status, queued_at);
create index if not exists roof_intelligence_jobs_user_idx
  on public.roof_intelligence_jobs (requested_by, created_at desc);

create table if not exists public.roof_intelligence_reports (
  id uuid primary key default gen_random_uuid(),
  property_id uuid not null references public.properties(id),
  job_id uuid not null references public.roof_intelligence_jobs(id),
  requested_by uuid references auth.users(id),
  storage_bucket text not null default 'roof-intelligence-reports',
  storage_path text not null,
  pdf_size bigint,
  pdf_checksum text,
  roof_type text,
  roof_type_confidence numeric,
  condition_score numeric,
  risk_level text,
  imagery_source text,
  imagery_capture_date text,
  workflow_version text,
  result jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists roof_intelligence_reports_property_idx
  on public.roof_intelligence_reports (property_id, created_at desc);
create index if not exists roof_intelligence_reports_job_idx
  on public.roof_intelligence_reports (job_id, created_at desc);

create table if not exists public.roof_intelligence_job_items (
  id uuid primary key default gen_random_uuid(),
  job_id uuid not null references public.roof_intelligence_jobs(id) on delete cascade,
  property_id uuid references public.properties(id),
  candidate_key text,
  status text not null,
  stage text not null,
  reason_code text,
  message text,
  report_id uuid references public.roof_intelligence_reports(id),
  created_at timestamptz not null default now(),
  started_at timestamptz,
  finished_at timestamptz
);

create index if not exists roof_intelligence_job_items_job_idx
  on public.roof_intelligence_job_items (job_id, status);

create table if not exists public.notifications (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id),
  job_id uuid references public.roof_intelligence_jobs(id),
  report_id uuid references public.roof_intelligence_reports(id),
  kind text not null,
  title text not null,
  message text not null,
  is_read boolean not null default false,
  created_at timestamptz not null default now(),
  read_at timestamptz,
  unique (job_id, kind)
);

create index if not exists notifications_user_read_idx
  on public.notifications (user_id, is_read, created_at desc);

insert into storage.buckets (id, name, public)
values ('roof-intelligence-reports', 'roof-intelligence-reports', false)
on conflict (id) do update set public = excluded.public;

alter table public.properties enable row level security;
alter table public.roof_intelligence_jobs enable row level security;
alter table public.roof_intelligence_reports enable row level security;
alter table public.roof_intelligence_job_items enable row level security;
alter table public.notifications enable row level security;

-- Initial owner policies. Expand these when PCS administrative roles are defined.
create policy "Users can read their Roof Intelligence jobs"
  on public.roof_intelligence_jobs for select
  using (requested_by = auth.uid());

create policy "Users can submit Roof Intelligence jobs"
  on public.roof_intelligence_jobs for insert
  with check (requested_by = auth.uid());

create policy "Users can read properties linked to their reports"
  on public.properties for select
  using (
    exists (
      select 1 from public.roof_intelligence_reports report
      where report.property_id = id and report.requested_by = auth.uid()
    )
  );

create policy "Users can read their Roof Intelligence reports"
  on public.roof_intelligence_reports for select
  using (requested_by = auth.uid());

create policy "Users can read their Roof Intelligence job items"
  on public.roof_intelligence_job_items for select
  using (
    exists (
      select 1 from public.roof_intelligence_jobs job
      where job.id = job_id and job.requested_by = auth.uid()
    )
  );

create policy "Users can read their notifications"
  on public.notifications for select
  using (user_id = auth.uid());

create policy "Users can update their notifications"
  on public.notifications for update
  using (user_id = auth.uid())
  with check (user_id = auth.uid());

create policy "Users can read report files"
  on storage.objects for select
  using (
    bucket_id = 'roof-intelligence-reports'
    and exists (
      select 1
      from public.roof_intelligence_reports report
      where report.storage_path = name and report.requested_by = auth.uid()
    )
  );

-- PilotPoint IQ uses the Supabase service role. The service role bypasses RLS
-- to claim jobs, update canonical properties, write reports, and notify users.
