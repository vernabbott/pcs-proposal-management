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
