# PCS Beta-to-Production Integration Inventory

Baseline comparison performed August 17, 2026:

- Production baseline: `05d7b861bad785d55660b411507e5cf841311dbb`
- Beta baseline: `eafa4b4fe11188adfb8355d68a379df64081597c`
- Common ancestor: `c1efefa41f15fc5c6edcdb70cde0a47966667e61`
- Scope: 30 files, approximately 4,157 additions and 1,000 removals

## Integration groups

| Group | Principal files | Integration treatment |
| --- | --- | --- |
| Runtime isolation | `beta_runtime.py`, `run_beta_app.py`, packaging spec | Retain beta build; add an isolated integration runtime and app on port 5052. |
| Tenant authentication | `tenant_context.py`, `tenant_settings_store.py`, tenant login template | Enable through `PCS_MULTI_TENANT_ENABLED`, independent of the beta label. |
| Proposal persistence | `proposal_tracking_store.py`, cutover flags, proposal schemas | Select explicitly with `PCS_PROPOSAL_STORAGE_MODE`: `spreadsheet`, `shadow`, or `supabase`. |
| Proposal management | proposal list/detail routes and templates | Use database-backed drafts when the storage mode is `supabase`, not merely when the app is named beta. |
| Contacts and organizations | `contact_store.py`, contact routes/templates | Retain tenant-scoped JWT access and organization/contact relationships. |
| Roof report tenancy | roof intelligence job/user keys and tenant settings | Namespace local compatibility data by tenant whenever multi-tenancy is enabled. |
| Production corrections | tracker header mapping and pricing calculations | Preserve production changes while merging the normalized beta implementation. |
| Tests | lifecycle, contact, proposal, settings, tenant suites | Run as a single release-candidate suite before production migration work. |

## Runtime contract

- Production defaults remain single-tenant and spreadsheet-backed unless explicitly configured.
- The integration build is multi-tenant, Supabase-only, isolated under Application Support, and listens on port 5052.
- The beta build remains isolated on port 5051.
- No desktop build accepts or preserves a Supabase secret/service-role key.
