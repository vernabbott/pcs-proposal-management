# PCS Proposal Beta Environment

The beta application is installed as `/Applications/PCS_Proposal_Beta.app` and
runs independently from `/Applications/PCS_Proposal.app`.

## Isolation

- Beta uses port `5051`; production continues to use `5050`.
- Beta runtime data and settings are stored in
  `~/Library/Application Support/PCS Proposal Management Beta`.
- Beta proposal folders and its Excel tracker live under that directory's
  `Workspace` folder.
- Beta loads PilotPoint from the `PilotPoint IQ Roof Intelligence Report Beta`
  project folder.
- Supabase proposal tracking is disabled by default, and beta does not inherit
  production Supabase credentials. Dedicated beta credentials can be supplied
  with `PCS_BETA_SUPABASE_URL` and `PCS_BETA_SUPABASE_PUBLISHABLE_KEY`.
- Multi-tenancy is always active in beta. It is intentionally not controlled
  by a feature flag.

## Local Supabase

The shared local beta database is managed from the PilotPoint beta repository.
It applies the Roof Intelligence, contact, organization, and proposal-tracking
migrations, then loads only a small synthetic seed. It does not copy the large
production footprint dataset, production contacts, reports, or imagery.

```sh
cd "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/PilotPoint IQ Roof Intelligence Report Beta"
colima start
npm run supabase:start
npm run supabase:status
```

If Auth reports an invalid upstream response immediately after a reset, restart
the local API gateway once with
`docker restart supabase_kong_PilotPoint_IQ_Roof_Intelligence_Report_B`.

Use `http://127.0.0.1:54321` and the local publishable key shown by the status
command in the beta Settings screen. Never place the service-role key in PCS;
that secret belongs only to the protected PilotPoint worker. The synthetic PCS
owner login is `owner@pcs-beta.test` with password `PCS-Beta-Owner-2026!`.
Proposal tracking in beta is fully cut over to the tenant-scoped Supabase
tables. Supabase supplies all proposal reads and receives all proposal writes;
the beta proposal-tracking workbook is not read or updated. The production
application rejects this loopback URL and remains isolated.

Customer, contact, proposal, job, report, revision, asset, notification, and
feedback records are tenant-owned. Large footprint and canonical-property
datasets remain shared. Local report copies are written below
`tenants/<tenant UUID>/`; protected Storage objects use
`<tenant UUID>/folders/<folder UUID>/reports/<report UUID>/revisions/...`.

When local work is finished:

```sh
npm run supabase:stop
colima stop
```

## Test and build

```sh
env PROPOSAL_TRACKING_SUPABASE_ENABLED=1 \
  PROPOSAL_TRACKING_SUPABASE_READS_ENABLED=1 \
  PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED=1 \
  PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED=0 \
  .venv/bin/python -m unittest discover

.venv/bin/python -m PyInstaller --clean --noconfirm PCS_Proposal_Beta.spec
```

The bundle must be copied out of OneDrive before final signing so Finder/File
Provider metadata can be removed. Apply an ad-hoc signature to that local copy,
verify it with `codesign --verify --deep --strict`, and then install it as
`/Applications/PCS_Proposal_Beta.app`.
