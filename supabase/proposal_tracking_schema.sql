-- Proposal tracking schema for the staged PCS spreadsheet-to-Supabase cutover.

create table if not exists public.proposal (
  id uuid primary key default gen_random_uuid(),
  customer_name text not null check (btrim(customer_name) <> ''),
  normalized_customer_name text generated always as (
    lower(regexp_replace(btrim(customer_name), '[[:space:]]+', ' ', 'g'))
  ) stored,
  project_street_address text,
  normalized_project_street_address text generated always as (
    nullif(lower(regexp_replace(btrim(project_street_address), '[[:space:]]+', ' ', 'g')), '')
  ) stored,
  project_address_line_2 text,
  project_city text,
  project_state text check (
    project_state is null or char_length(btrim(project_state)) = 2
  ),
  project_zip_code text,
  display_name text generated always as (
    btrim(customer_name) ||
    case
      when nullif(btrim(project_street_address), '') is null then ''
      else ' - ' || btrim(project_street_address)
    end
  ) stored,
  lead_source text,
  submitted_by text,
  estimated_by text,
  estimate_completed_date date,
  proposal_sent_date date,
  follow_up_date date,
  response_notes text,
  status text not null default 'draft' check (
    status in ('draft', 'sent', 'under_contract', 'dead')
  ),
  proposal_folder_name text,
  source_name text,
  source_row_number integer check (
    source_row_number is null or source_row_number >= 2
  ),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (source_name, source_row_number)
);

comment on table public.proposal is
  'Normalized proposal tracking records replacing Proposal Tracking.xlsx after a staged feature-flag cutover.';
comment on column public.proposal.display_name is
  'Customer and project street address formatted for PCS screens as Customer - Street Address.';
comment on column public.proposal.estimate_completed_date is
  'Date the estimate was completed; distinct from the date the proposal was sent.';
comment on column public.proposal.source_row_number is
  'Original spreadsheet row number used for idempotent migration and reconciliation.';

create index if not exists proposal_customer_address_idx
  on public.proposal (normalized_customer_name, normalized_project_street_address);
create index if not exists proposal_display_name_idx
  on public.proposal (display_name);
create index if not exists proposal_follow_up_queue_idx
  on public.proposal (proposal_sent_date, follow_up_date)
  where proposal_sent_date is not null and follow_up_date is null;
create index if not exists proposal_folder_name_idx
  on public.proposal (lower(btrim(proposal_folder_name)))
  where proposal_folder_name is not null;

create table if not exists public.proposal_contact (
  proposal_id uuid not null references public.proposal(id) on delete cascade,
  organization_contact_id uuid not null references public.organization_contact(id) on delete restrict,
  contact_role text not null default 'primary' check (btrim(contact_role) <> ''),
  is_primary boolean not null default false,
  created_at timestamptz not null default now(),
  primary key (proposal_id, organization_contact_id)
);

comment on table public.proposal_contact is
  'Links proposals to historical organization-contact relationships without duplicating contact names or email addresses.';

create unique index if not exists proposal_contact_one_primary_idx
  on public.proposal_contact (proposal_id)
  where is_primary;
create index if not exists proposal_contact_relationship_idx
  on public.proposal_contact (organization_contact_id);

drop trigger if exists proposal_set_updated_at on public.proposal;
create trigger proposal_set_updated_at
before update on public.proposal
for each row execute function public.set_property_management_updated_at();

alter table public.proposal enable row level security;
alter table public.proposal_contact enable row level security;

revoke all on table public.proposal from anon, authenticated;
revoke all on table public.proposal_contact from anon, authenticated;
grant select, insert, update, delete on table public.proposal to service_role;
grant select, insert, update, delete on table public.proposal_contact to service_role;
