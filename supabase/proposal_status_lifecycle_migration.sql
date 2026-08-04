-- Collapse proposal outcomes into the PCS four-state lifecycle.
-- Follow-up remains an activity date and does not change a proposal's status.

alter table public.proposal
  drop constraint if exists proposal_status_check;

update public.proposal
set status = case
  when status = 'follow_up' then 'sent'
  when status = 'won' then 'under_contract'
  when status in ('lost', 'withdrawn', 'archived') then 'dead'
  else status
end
where status in ('follow_up', 'won', 'lost', 'withdrawn', 'archived');

alter table public.proposal
  add constraint proposal_status_check
  check (status in ('draft', 'sent', 'under_contract', 'dead'));
