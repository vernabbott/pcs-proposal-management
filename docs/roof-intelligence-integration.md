# Roof Intelligence Integration Architecture

## Document Status

- Status: Initial design
- Owning application: PCS Proposal Management
- Processing engine: PilotPoint IQ Roof Intelligence
- Shared persistence and handoff: Supabase

## Implementation Status

The PCS-side request forms, persistent job model, status API, elapsed-time display, concise failures, notifications, and canonical-property/report development adapter are implemented. Local development currently uses SQLite through the same logical data contract documented here.

The Supabase schema is defined in `supabase/roof_intelligence_schema.sql` but has not yet been applied to a Supabase project. The production Supabase adapter, PCS authentication integration, and continuously running PilotPoint IQ worker remain to be implemented. ZIP jobs can be submitted and inspected now, but report processing will begin only after the PilotPoint batch worker is connected.

## Purpose

This document defines the initial architecture for integrating PilotPoint IQ Roof Intelligence into PCS Proposal Management. The integration will use the existing Roof Intelligence option on the PCS Proposal Management landing page and will support both individual-address reports and ZIP-code report batches.

PCS Proposal Management and PilotPoint IQ will remain separate projects. PCS will own the user experience, job submission, report presentation, and notifications. PilotPoint IQ will operate as a background worker responsible for property-data collection, imagery retrieval, AI analysis, and PDF generation. Supabase will provide the shared job state, normalized records, report metadata, notifications, and PDF storage.

## Architecture Overview

```text
PCS Proposal Management
  - Roof Intelligence forms
  - Job submission
  - Progress and elapsed-time display
  - Results and notifications
              |
              v
Supabase
  - Jobs and job items
  - Canonical properties
  - Dated report records
  - Notifications
  - Private PDF storage
              |
              v
PilotPoint IQ Worker
  - County and parcel lookup
  - Building and assessor collection
  - Imagery retrieval and preparation
  - AI roof assessment
  - PDF generation
  - Temporary-file cleanup
```

The deployed PCS application will not depend on local OneDrive paths or directly import files from the PilotPoint IQ project. The two applications will communicate through a defined Supabase data contract. During local development, the PilotPoint IQ worker may run on a developer computer while reading and updating jobs in Supabase. Production will require a continuously available worker service.

## Application Responsibilities

### PCS Proposal Management

PCS Proposal Management will:

- Present the Individual Address and ZIP Code Batch options.
- Validate user input before creating a job.
- Create job records in Supabase.
- Display job status, processing stage, elapsed time, and progress.
- Present completed report summaries and secure PDF actions.
- Present concise failure explanations and retry actions.
- Display unread completion or failure notifications.
- Display recent Roof Intelligence reports and property report history.
- Enforce user authentication and report-access permissions.

### PilotPoint IQ Worker

PilotPoint IQ will:

- Claim queued jobs without allowing two workers to process the same item.
- Determine the applicable county for an address or ZIP-code property.
- Query configured county parcel, building, assessor, and imagery services.
- Normalize property and parcel information into the shared data model.
- Retrieve and prepare aerial imagery for AI processing.
- Run roof classification and condition assessment.
- Generate the final Roof Intelligence PDF.
- Upload the completed PDF to private Supabase Storage.
- Update job, job-item, property, and report records.
- Record concise warnings and failure reasons.
- Delete temporary imagery and intermediate files after completion or expiration.

### Supabase

Supabase will provide:

- PostgreSQL storage for normalized application data.
- Private object storage for final PDF reports.
- Authentication and row-level security.
- Shared job state between PCS and PilotPoint IQ.
- Realtime updates or polling support for progress and notifications.

Supabase is not expected to execute the long-running Python report-generation process.

## Roof Intelligence Page

The existing Roof Intelligence landing-page option will open a page containing two clearly separated choices:

1. Individual Address
2. ZIP Code Batch

The page should also provide access to recent jobs, unread notifications, and recently completed reports.

## Individual Address Workflow

### User Input

The user must enter a complete address containing enough information to identify the property, including:

- Street address
- City
- State
- ZIP code

The interface may use a single full-address field initially. The submitted value will be normalized before property matching.

### Processing Flow

```text
User submits full address
  -> PCS validates input
  -> PCS creates a queued job
  -> PilotPoint worker claims the job
  -> Address is normalized and geocoded
  -> County is identified
  -> County parcel, assessor, and building data is collected
  -> Aerial imagery is retrieved and prepared
  -> AI analysis is performed
  -> PDF is generated and verified
  -> PDF is uploaded to Supabase Storage
  -> Canonical property and dated report records are saved
  -> Job is completed and user is notified
  -> Temporary imagery is deleted
```

### Status and Running Clock

The job will run asynchronously. The PCS page will display a running elapsed-time clock calculated from the job's `started_at` timestamp. The database does not need to receive clock updates every second.

Recommended processing stages:

- Queued
- Locating property
- Retrieving parcel data
- Retrieving imagery
- Analyzing roof
- Creating report
- Uploading report
- Completed
- Failed

The user may navigate away or refresh the page without stopping the job. Returning to the job will reconstruct the clock and current stage from stored timestamps and status.

### Completed Result Presentation

When the job completes, the running-job panel should change into a results view on the same page. The result should include:

- Report Ready confirmation
- Property address
- Completion date and total processing time
- County and parcel number
- Roof area
- Building/roof age estimate and source
- Identified roof type and confidence
- Roof condition score and risk level
- Important imagery or data-quality limitations

Primary actions:

- View Report
- Download PDF
- Create Another Report
- View Report History
- Associate with Proposal, when implemented

`View Report` and `Download PDF` will use authenticated access or a short-lived signed URL to the private Supabase Storage object.

### Failure Presentation

A failed job must identify:

- The processing stage that failed
- A concise user-facing explanation
- Whether the failure appears retryable
- A Retry Report action
- An Edit Address action when appropriate
- A job reference identifier for support

Warnings that do not prevent report generation should produce a completed-with-limitations presentation rather than a failed job.

## ZIP Code Batch Workflow

### User Input

The ZIP Code Batch form will collect:

- ZIP code
- Report limit: maximum number of reports to process
- Minimum roof size in square feet, default `10,000`
- Minimum Building/Roof Age Estimate, optional
- Roof-type checkbox selections
- An All option that selects or clears every supported roof type

### Age Estimate Definition

County data generally provides building year or effective year built rather than a verified roof-installation date. The filter must therefore be labeled **Minimum Building/Roof Age Estimate**.

Age-source priority:

1. Verified roof installation or replacement permit date, when available
2. County effective year built
3. County original year built
4. Unknown when no usable value exists

Persist the following values:

- `age_estimate_year`
- `age_estimate_years`
- `age_estimate_source`
- `age_estimate_as_of_date`

If the user enters a minimum age and a property has no usable age data, skip that property and count it as `age_unavailable`. Do not assume the property qualifies.

### Roof-Type Filtering

Roof type may not be known until imagery has been retrieved and AI classification has run. The worker should:

1. Reuse a sufficiently current stored classification when available.
2. Otherwise perform roof classification.
3. Continue full report generation only when the classification matches the selected types.

This avoids creating unwanted PDFs, although classification may still incur processing time and AI cost. The All option bypasses the type-match exclusion.

### Batch Processing

The worker will:

- Determine which supported counties intersect the ZIP code.
- Collect eligible properties from the appropriate county services.
- Apply property, roof-size, and age filters.
- Classify roof type when required.
- Generate no more completed reports than the requested report limit.
- Track completed, failed, skipped, and remaining items.
- Continue processing other items when an individual property fails.
- Create a completion notification when the batch reaches a terminal state.

Terminal batch statuses:

- `completed`
- `completed_with_errors`
- `failed`

The results view should include counts for:

- Candidate properties
- Completed reports
- Failed properties
- Skipped properties
- Age unavailable
- Roof size below minimum
- Roof type excluded
- Parcel or imagery unavailable

## Notifications

PCS Proposal Management will display user-specific in-app notifications.

Notifications should be created when:

- An individual report completes
- An individual report fails
- A ZIP batch completes
- A ZIP batch completes with errors
- A ZIP batch fails

Each notification will have read/unread status and will link to the applicable job or results page. A user leaving the Roof Intelligence page must still be able to discover that a job completed.

## Canonical Properties and Dated Reports

Duplicate address processing must not create duplicate property records.

Property matching priority:

1. County plus canonical parcel number
2. Stable county property identifier, when available
3. Normalized address as a fallback

The `properties` table will contain the most current normalized property information. Each successful run will create a separate dated record in `roof_intelligence_reports`. This supports report history and comparison without duplicating the underlying property.

Updating a canonical property must not overwrite historical report results. A report records the assessment and relevant source metadata as they existed when the report was created.

## Proposed Supabase Data Model

### `properties`

One canonical record per property. Initial fields should include:

- `id`
- Normalized address fields
- County and state
- Parcel number and county property identifier
- Latitude and longitude
- Owner and property classification, when retained
- Building footprint or normalized footprint representation
- Roof area and roof squares
- Year built and effective year built
- Current age-estimate fields
- Created and updated timestamps

### `roof_intelligence_jobs`

One record per individual or ZIP batch submission:

- `id`
- `job_type`: `individual_address` or `zip_batch`
- Requesting user
- Status and processing stage
- Submitted input and normalized filters
- Requested report limit
- Progress counters
- `queued_at`, `started_at`, and `finished_at`
- Concise warning or error fields
- Retryable indicator
- Worker and workflow version

### `roof_intelligence_job_items`

One record for each property considered by a batch, and optionally one for an individual job:

- Parent job
- Property or candidate identifier
- Item status and stage
- Skip or failure category
- Concise explanation
- Resulting report identifier
- Started and finished timestamps

### `roof_intelligence_reports`

One dated record per completed assessment:

- `id`
- Property identifier
- Source job and job item
- Requesting user
- Report date
- Roof type and confidence
- Roof condition score and risk level
- Structured observations and summary
- Replacement-cost estimate and confidence
- Age-estimate values and source
- Imagery source, reported capture date, and limitations
- AI provider and model
- Workflow and reference-library versions
- PDF bucket and object path
- PDF size and checksum
- Created timestamp

### `roof_observations`

Optional normalized observation rows when individual findings need to be queried independently. If observations do not require relational querying, they may initially be stored as structured JSON on the report record.

### `notifications`

- `id`
- User identifier
- Notification type
- Job and report identifiers
- Short title and message
- Read/unread status
- Created and read timestamps

## Data Sources

The integration will use county-level parcel and assessor services. The earlier `colorado_parcel_data.csv` dependency was a local cache of Denver parcel data despite its statewide-sounding filename. It is not required for the modern single-address architecture.

PilotPoint owns footprint collection, Microsoft-to-county comparison, canonical remediation, and the directional 5% rule. PCS owns ordering, presentation, and review workflow only; it consumes the validation status returned by PilotPoint and does not reimplement the comparison.

Individual-address lookup:

```text
Address -> geocode -> identify county -> query county parcel/assessor service
```

ZIP lookup:

```text
ZIP boundary -> identify intersecting counties -> query each supported county
```

Previously retrieved canonical properties in Supabase may be reused subject to a defined freshness policy. County services remain the authoritative source for refreshes.

## Data Retention

### Retain Permanently

- Canonical normalized property record
- Dated structured report results
- Roof observations and limitations
- Imagery source and reported capture date
- AI model and processing-version metadata
- Final PDF in private Supabase Storage
- Job status, summary counts, and concise errors
- User notification records according to the notification-retention policy

### Delete After Successful Completion

- Raw aerial imagery
- Full-resolution source tiles
- AI roof crops
- Resized or converted imagery
- Base64 or inline API image payloads
- Temporary CSV and JSON working files
- Intermediate PDF-rendering files

Temporary files must be deleted only after:

1. AI analysis succeeds.
2. The PDF is generated and passes a basic integrity check.
3. The PDF is uploaded successfully to Supabase Storage.
4. The completed report and job records are committed successfully.

A scheduled cleanup process must delete orphaned temporary files left by failed, interrupted, or abandoned jobs.

The final PDF already embeds the aerial image used for the assessment. Separate raw imagery does not need to be retained. Manual verification may use Google Earth, recognizing that its imagery may have a different acquisition date than the image assessed in the original report.

## Storage Estimate

Existing generated PDFs average approximately 418 KB. Expected retained storage per report is:

- Final PDF in Supabase Storage: approximately 418 KB
- PostgreSQL property, report, observation, job, and index data: approximately 20-60 KB
- Planning estimate: approximately 0.5 MB per completed report

For 1,000 completed reports, plan for approximately 500 MB of combined retained data. A 750 MB to 1 GB allowance provides room for larger PDFs, indexes, and normal growth.

The PDF must be stored in Supabase Storage, not directly in a PostgreSQL binary column. PostgreSQL will store the private object path, file size, checksum, and report metadata.

## Error and Log Retention

Persist only normalized, concise error information:

- Error category
- Processing stage
- Short user-facing explanation
- Timestamp
- Retryable indicator
- Support job identifier

Do not permanently store every raw county response, AI request payload, encoded image, stack trace, or verbose worker log. Detailed worker logs may be retained temporarily for troubleshooting and then expired automatically.

Suggested error categories include:

- `invalid_input`
- `address_not_found`
- `unsupported_county`
- `parcel_not_found`
- `building_not_found`
- `county_service_unavailable`
- `imagery_unavailable`
- `ai_provider_unavailable`
- `ai_analysis_failed`
- `pdf_generation_failed`
- `storage_upload_failed`
- `internal_processing_error`

## Security

- Store PDFs in a private Supabase bucket.
- Use authenticated access or short-lived signed URLs.
- Apply row-level security to jobs, reports, and notifications.
- Permit users to access only records allowed by their PCS role.
- Keep Supabase service-role, county-service, imagery, and AI credentials in the PilotPoint worker environment.
- Never expose service-role or AI credentials in browser code.
- Record the requesting user and creation timestamps for jobs and reports.

## Reliability and Recovery

- Use atomic job claiming to prevent duplicate processing.
- Apply timeouts and bounded retries to external services.
- Use idempotent report creation so a retry does not create duplicate completed records.
- Continue ZIP batches after individual-item failures.
- Detect jobs abandoned by an unavailable worker and make them retryable.
- Verify the PDF before marking a report complete.
- Do not mark a job complete until the PDF and database records are safely stored.

## Local Development and Production

### Local Development

- PCS Proposal Management creates jobs in the configured Supabase project.
- A locally running PilotPoint IQ worker claims and processes jobs.
- The developer computer must remain running while jobs are processed.
- Local tests should include simulated county, imagery, AI, PDF, and upload failures.

### Production

- PCS Proposal Management runs as its web application service.
- PilotPoint IQ runs as a continuously available background-worker service.
- Both applications communicate through Supabase.
- Production must not reference a developer computer or OneDrive project path.
- Worker capacity and concurrency may be scaled independently from PCS web traffic.

## Implementation Sequence

1. Confirm this architecture and define the initial Supabase schema.
2. Add Supabase migrations, private Storage bucket configuration, and row-level-security policies to PCS Proposal Management.
3. Replace the obsolete synchronous single-address bridge with job creation and status retrieval.
4. Update the Roof Intelligence page with Individual Address and ZIP Code Batch forms.
5. Add the individual-job running clock, processing stages, results, and failure presentation.
6. Add recent reports, report history, and in-app notifications.
7. Update PilotPoint IQ to claim and process individual jobs.
8. Verify the individual workflow end to end, including refresh and navigation during processing.
9. Implement ZIP discovery, filtering, job items, limits, progress counts, and partial completion.
10. Add scheduled cleanup and operational monitoring.

## Initial Acceptance Criteria

The first implementation milestone is complete when a signed-in PCS user can:

1. Submit a full address from the existing Roof Intelligence page.
2. Receive a persistent asynchronous job identifier immediately.
3. See an accurate elapsed-time clock and current processing stage.
4. Refresh or leave the page without interrupting processing.
5. Return to the same running or completed job.
6. Receive either a downloadable PDF or a concise actionable failure reason.
7. Process the same address again without duplicating the canonical property.
8. See each successful run as a separate dated report.
9. Receive an unread in-app notification when processing finishes away from the page.
10. Confirm that temporary imagery is removed after the final report is safely stored.

ZIP-batch acceptance criteria will build on the same job infrastructure and add input filters, report limits, per-property status, progress counts, partial completion, and completion notifications.

## Open Implementation Decisions

The following details should be finalized during schema and worker implementation:

- Supabase project and environment separation for development and production
- Exact worker hosting platform
- Worker polling versus event-trigger mechanism
- Concurrency and per-user batch limits
- Property and stored-classification freshness periods
- Supported roof-type list and canonical values
- ZIP candidate ordering when eligible properties exceed the report limit
- Notification retention period
- Temporary diagnostic-log retention period
- Proposal-association workflow and permissions
