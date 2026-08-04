# Production Baseline

- Baseline ID: `production-2026-08-04.2`
- Component: PCS Proposal Management
- Included application source: all PCS changes through commit `d61d1a2`
- Paired PilotPoint source: all Roof Intelligence changes through commit `9d8fc00`
- Deployment: `/Applications/PCS_Proposal.app`
- PilotPoint integration: the desktop application invokes the PilotPoint worker through `ROOF_INTELLIGENCE_PROJECT_DIR`

This paired baseline is the production rollback point created before beta and
multi-tenant development begins. Production releases should be built from the
matching `production-2026-08-04.2` tag in both repositories.
