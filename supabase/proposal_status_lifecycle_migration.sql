-- Collapse legacy proposal outcomes into the PCS lifecycle.
-- Follow-up remains an activity date and does not change a proposal's status.

alter table public.proposal_tracking
  drop constraint if exists proposal_tracking_status_check;

update public.proposal_tracking
set status = case
  when status = 'follow_up' then 'sent'
  when status = 'won' then 'under_contract'
  when status in ('lost', 'withdrawn', 'archived') then 'dead'
  else status
end
where status in ('follow_up', 'won', 'lost', 'withdrawn', 'archived');

alter table public.proposal_tracking
  add constraint proposal_tracking_status_check
  check (status in ('draft', 'sent', 'under_contract', 'finished', 'dead'));
