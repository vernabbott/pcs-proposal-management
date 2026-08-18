# PCS Production Rollback Runbook — Pre-Integration August 17, 2026

Last verified: August 18, 2026

Owner: Pro Coating Systems

Scope: PCS Proposal application, proposal-tracking workbook, and Supabase database

## Purpose

Use this runbook if the PCS integration release must be reversed to the production state that existed before the August 17, 2026 integration testing cutover.

The safest response is normally an **application-only rollback**. The August 17 database migration was additive and retained the legacy proposal fields needed by the pre-integration application. Revert the database only when there is evidence that its schema or data is causing the failure.

## Verified recovery assets

| Asset | Verified value |
| --- | --- |
| Pre-integration production Git tag | `rollback-pre-beta-migration-2026-08-17-pcs-production` |
| Tagged source commit | `05d7b861bad785d55660b411507e5cf841311dbb` |
| Existing production application | `/Applications/PCS_Proposal.app` |
| Production app build timestamp | August 17, 2026 at 6:54:05 AM NDT |
| Integration application | `/Applications/PCS_Proposal_Integration.app` |
| Production workbook | `/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/PCS/1 - Open Proposals/Proposal Tracking.xlsx` |
| Active production Supabase project | `Pilot Point IQ` — `sipeplmvdhfbkutzslxp` |
| Preserved pre-integration project | `Pilot Point IQ - Pre Integration 2026-08-17` — `eyxhtbdbrwmmviylkmuc` |
| Preserved database recovery point | August 17, 2026 at 12:13:18 |
| Last migration in preserved database | `20260816095102_add_finished_proposal_status` |
| First migration intentionally absent | `20260817163527_prepare_production_multitenant_cutover` |

The preserved project was verified as an independent, healthy database copy. It has no tenant tables, no integration migrations, no Auth users, no Edge Functions, no active cron/network extensions, no foreign servers, and no database subscriptions.

Supabase Storage objects are not included in a database restore. PCS proposal files, the tracking workbook, locally downloaded reports, and other OneDrive files must be protected separately.

## Recovery decision

Use the least invasive recovery level that solves the incident:

1. **Level 1 — Application rollback:** Restore the pre-integration PCS application but keep the current production database. This is the recommended first response.
2. **Level 2 — Database failover:** Run the pre-integration application against the preserved pre-integration Supabase project. This leaves the current production database untouched.
3. **Level 3 — In-place database restoration:** Replace the active production database with the pre-integration database state. Use only when the production project reference must remain unchanged and Levels 1 and 2 are insufficient.

## Before any rollback

1. Announce a maintenance window and stop users from entering proposals or changing tracking records.
2. Quit `PCS_Proposal_Integration.app` on every test workstation.
3. Stop the PilotPoint/roof-intelligence worker if it is running.
4. Close `Proposal Tracking.xlsx` in Excel and allow OneDrive to finish synchronizing.
5. Record the rollback start time and the reason for rollback.
6. Preserve the current state before changing anything:
   - Create a new Supabase backup or restore-to-new-project copy of the current production database.
   - Copy `Proposal Tracking.xlsx` to a timestamped rollback folder.
   - Preserve the current PCS Integration settings and worker configuration without placing credentials in Git or this runbook.
7. Confirm both database projects show `ACTIVE_HEALTHY` in Supabase.

Do not delete the current production project or the preserved pre-integration project during recovery.

## Level 1 — Roll back the application only

This path retains all current database data and is the fastest, lowest-risk rollback.

1. Confirm `PCS_Proposal_Integration.app` is closed.
2. Start the preserved production application:

   ```text
   /Applications/PCS_Proposal.app
   ```

3. Confirm its local settings point to the active production project:

   ```text
   https://sipeplmvdhfbkutzslxp.supabase.co
   ```

4. Confirm the production proposal-tracking configuration is:

   | Setting | Required value |
   | --- | --- |
   | Supabase master flag | Enabled |
   | Supabase reads | Disabled |
   | Supabase writes | Enabled |
   | Spreadsheet shadow writes | Enabled |

   In this configuration, the production workbook remains the source for proposal-tracking reads while changes are written to both the workbook and Supabase.

5. Do not run the integration and production applications simultaneously while performing rollback verification.
6. Run the validation checklist below.

If `/Applications/PCS_Proposal.app` is damaged, rebuild from the rollback Git tag in a separate Git worktree. Do not reset or overwrite the integration worktree. The rebuild source must resolve to commit `05d7b861bad785d55660b411507e5cf841311dbb` before packaging.

## Level 2 — Fail over to the preserved pre-integration database

Use this path when the integration database changes must be removed from service but the current production database should remain recoverable.

1. Complete all “Before any rollback” steps.
2. Keep the current production project `sipeplmvdhfbkutzslxp` online but stop all application writes to it.
3. In the preserved project `eyxhtbdbrwmmviylkmuc`, obtain new project-specific credentials from Supabase. Never reuse or copy production project credentials.
4. Configure `/Applications/PCS_Proposal.app` to use:

   ```text
   https://eyxhtbdbrwmmviylkmuc.supabase.co
   ```

5. Store the preserved project's secret credential only in the protected local configuration used by the desktop/server process. Do not expose it on a user-facing settings screen or commit it to Git.
6. Keep Supabase reads disabled and workbook reads enabled. The workbook is the source of truth for the pre-integration application.
7. Reconcile workbook changes made after the August 17 12:13 backup into the preserved database before relying on it for comparison or shadow writes. Generate and retain a mismatch report; require zero unexplained mismatches.
8. If roof intelligence needs direct database access, configure its protected worker credential for the preserved project and run county/report health checks before reopening access.
9. Run the validation checklist below.

This is a failover, not an overwrite. Returning to the newer production database remains possible by restoring the original production URL and credentials.

## Level 3 — Restore the active production database in place

Use this only when retaining the original production project URL is mandatory. It is the most disruptive option and can discard every database change made after August 17 at 12:13:18.

1. Complete all “Before any rollback” steps, including preservation of the current database.
2. Confirm the current workbook backup and current-database backup are readable.
3. If the August 17 scheduled backup is still inside Supabase retention, restore that recovery point to the active production project through **Database → Backups**.
4. If the scheduled backup has expired, use the preserved project as the source for a controlled logical export and restore. Because a project created through “Restore to a New Project” cannot currently be used as the source for another clone, this path requires a database export/import procedure.
5. Perform the restore during a maintenance window. Supabase makes the project unavailable during an in-place restoration.
6. Reapply or rotate any required custom database credentials after restoration. Physical/database backups do not preserve every project-level configuration.
7. Reconcile post-backup workbook changes into the restored database and produce a zero-unexplained-mismatch report.
8. Verify the migration boundary:

   ```text
   Present: 20260816095102_add_finished_proposal_status
   Absent:  20260817163527_prepare_production_multitenant_cutover
   Absent:  20260817172300_restrict_rls_auto_enable
   ```

9. Run the validation checklist below before reopening the application.

Do not perform a full logical restore from memory or an improvised command. Use the current Supabase backup/restore documentation and have the exact export contents reviewed before execution, especially for `auth`, `storage`, platform-managed schemas, roles, and extensions.

## Validation checklist

The rollback is complete only after all applicable checks pass:

- The intended application opens and the integration application remains closed.
- The application points to the intended Supabase project reference.
- The proposal list loads from the production workbook under Level 1 or Level 2.
- Proposal counts and lifecycle statuses match the workbook.
- A workbook-versus-Supabase comparison has zero unexplained mismatches.
- A known proposal opens with its customer, address, contact, salesperson, estimator, estimate date, sent date, follow-up date, response, and status intact.
- Proposal pricing calculations match a known pre-integration proposal.
- Saving a controlled tracking change updates the workbook and the intended database exactly once.
- The follow-up list excludes Dead records and records where follow-up is not required.
- Contact and organization lookup works for a known proposal.
- A roof-intelligence health check succeeds before report production resumes.
- Bulk-email input/output folders point to the production locations.
- No application or worker is writing to the database project that has been taken out of service.
- The rollback time, database project reference, workbook copy, comparison report, and validation results are recorded.

## Return from rollback

Do not merge data in both directions while two databases are independently accepting writes.

Before returning to the integration release:

1. Freeze writes again.
2. Identify the single source of truth used during rollback.
3. Reconcile rollback-period workbook/database changes into the intended production database.
4. Create another current-state backup.
5. Reconfigure and test the integration application against the intended production project.
6. Reopen access only after counts, field comparisons, calculations, contacts, reports, and health checks pass.

## Supabase references

- [Database backups](https://supabase.com/docs/guides/platform/backups)
- [Restore to a new project](https://supabase.com/docs/guides/platform/clone-project)
- [Migrating between Supabase projects](https://supabase.com/docs/guides/platform/migrating-within-supabase)
- [Backup and restore using the CLI](https://supabase.com/docs/guides/platform/migrating-within-supabase/backup-restore)
