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
  with `PCS_BETA_SUPABASE_URL` and `PCS_BETA_SUPABASE_SERVICE_ROLE_KEY`.

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

Use `http://127.0.0.1:54321` and the local service-role key shown by the status
command in the beta Settings screen. Keep all proposal cutover flags off until
the database connection is verified; then enable and test shadow writes first.
The production application rejects this loopback URL and remains isolated.

When local work is finished:

```sh
npm run supabase:stop
colima stop
```

## Test and build

```sh
env PROPOSAL_TRACKING_SUPABASE_ENABLED=0 \
  PROPOSAL_TRACKING_SUPABASE_READS_ENABLED=0 \
  PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED=0 \
  PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED=0 \
  .venv/bin/python -m unittest discover

.venv/bin/python -m PyInstaller --clean --noconfirm PCS_Proposal_Beta.spec
```

The bundle must be copied out of OneDrive before final signing so Finder/File
Provider metadata can be removed. Apply an ad-hoc signature to that local copy,
verify it with `codesign --verify --deep --strict`, and then install it as
`/Applications/PCS_Proposal_Beta.app`.
