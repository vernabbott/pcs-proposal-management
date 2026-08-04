import os
from io import BytesIO
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
from urllib.parse import parse_qs, urlsplit

import roof_intelligence_jobs
import roof_intelligence_area_batch as area_batch
import roof_intelligence_single_address as single_address
from roof_intelligence_jobs import RoofIntelligenceJobStore, SUPPORTED_ROOF_TYPES
from roof_report_naming import roof_report_pdf_filename


class RoofIntelligenceJobStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "roof-jobs.sqlite3"
        self.store = RoofIntelligenceJobStore(self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_report_pdf_filename_always_uses_only_street_and_city(self):
        self.assertEqual(
            roof_report_pdf_filename("2261 E CORNELL AVE", "ENGLEWOOD"),
            "2261 E Cornell Ave Englewood.pdf",
        )
        self.assertEqual(
            roof_report_pdf_filename(
                "2261 E CORNELL AVE, Englewood, CO 80110",
                "ENGLEWOOD",
                revision_index=1,
            ),
            "2261 E Cornell Ave Englewood.pdf",
        )
        self.assertEqual(
            roof_report_pdf_filename(
                "2261 E CORNELL AVE STE 200",
                "ENGLEWOOD",
                revision_index=2,
            ),
            "2261 E Cornell Ave Englewood.pdf",
        )

    def test_individual_job_requires_full_address_with_zip(self):
        with self.assertRaisesRegex(ValueError, "five-digit ZIP"):
            self.store.create_individual_job("65 N Yuma St, Denver, CO")

        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")

        self.assertEqual(job["status"], "queued")
        self.assertEqual(job["job_type"], "individual_address")
        self.assertEqual(job["input"]["property_address"], "65 N Yuma St, Denver, CO 80223")

    def test_pcs_orders_enable_pilotpoint_roof_reference_workflow(self):
        reports = Mock()
        reports.load_or_create_analysis.return_value = {"roof_type": "TPO"}
        args = type(
            "Args",
            (),
            {
                "use_ai": True,
                "ai_provider": "openai",
                "allow_ai_fallback": True,
            },
        )()

        result = single_address.generate_roof_analysis(
            reports,
            {"Parcel Number": "123"},
            Path("aerial.jpg"),
            Path("analysis-cache"),
            args,
            "test-model",
        )

        self.assertEqual(result["roof_type"], "TPO")
        self.assertTrue(reports.load_or_create_analysis.call_args.kwargs["use_roof_references"])

    def test_area_job_persists_rectangle_and_applies_filters(self):
        job = self.store.create_area_job("39.75", "39.72", "-104.97", "-105.02", "100", ["All"])

        self.assertIsNone(job["report_limit"])
        self.assertEqual(job["minimum_roof_size"], 10000)
        self.assertEqual(job["input"]["minimum_roof_squares"], 100)
        self.assertIsNone(job["minimum_age"])
        self.assertNotIn("report_limit", job["input"])
        self.assertNotIn("minimum_age", job["input"])
        self.assertEqual(job["roof_types"], list(SUPPORTED_ROOF_TYPES))
        self.assertEqual(job["input"]["selection_type"], "rectangle")
        self.assertEqual(job["input"]["bounds"]["north"], 39.75)

    def test_area_job_accepts_commas_and_rounds_squares_up(self):
        job = self.store.create_area_job(
            "39.75", "39.72", "-104.97", "-105.02", "1,000.2", ["All"]
        )

        self.assertEqual(job["input"]["minimum_roof_squares"], 1001)
        self.assertEqual(job["minimum_roof_size"], 100100)

    def test_radius_area_job_derives_bounds_and_requires_distance(self):
        job = self.store.create_area_job(
            "", "", "", "", "100", ["All"],
            selection_type="radius",
            center_lat="39.7392",
            center_lng="-104.9903",
            center_address="1701 Wynkoop St, Denver, CO 80202",
            radius_miles="0.1",
        )

        self.assertEqual(job["input"]["selection_type"], "radius")
        self.assertEqual(job["input"]["center"], {"lat": 39.7392, "lng": -104.9903})
        self.assertEqual(job["input"]["center_address"], "1701 Wynkoop St, Denver, CO 80202")
        self.assertEqual(job["input"]["radius_miles"], 0.1)
        self.assertGreater(job["input"]["bounds"]["north"], 39.7392)
        self.assertLess(job["input"]["bounds"]["south"], 39.7392)

        with self.assertRaisesRegex(ValueError, "Enter radius distance in miles"):
            self.store.create_area_job(
                "", "", "", "", "100", ["All"],
                selection_type="radius",
                center_lat="39.7392",
                center_lng="-104.9903",
                center_address="1701 Wynkoop St, Denver, CO 80202",
                radius_miles="",
            )
        with self.assertRaisesRegex(ValueError, "between 0.1 and 10 miles"):
            self.store.create_area_job(
                "", "", "", "", "100", ["All"],
                selection_type="radius",
                center_lat="39.7392",
                center_lng="-104.9903",
                center_address="1701 Wynkoop St, Denver, CO 80202",
                radius_miles="10.1",
            )

    def test_area_worker_claims_radius_jobs(self):
        job = self.store.create_area_job(
            "", "", "", "", "100", ["All"],
            selection_type="radius",
            center_lat="39.7392",
            center_lng="-104.9903",
            center_address="1701 Wynkoop St, Denver, CO 80202",
            radius_miles="0.2",
        )

        self.assertTrue(self.store.has_queued_area_jobs())
        claimed = self.store.claim_next_area_job()

        self.assertEqual(claimed["id"], job["id"])
        self.assertEqual(claimed["status"], "running")
        self.assertEqual(claimed["stage"], "discovering_properties")
        self.assertIsNotNone(claimed["started_at"])

    def test_duplicate_property_keeps_one_canonical_record_and_two_reports(self):
        first = self.store.create_individual_job("65 N Yuma Street, Denver, CO 80223")
        second = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        result = {
            "address": "65 N Yuma St",
            "city": "Denver",
            "state": "CO",
            "zip": "80223",
            "county": "Denver",
            "parcel": "0508500065000",
            "roof_squares": 125,
            "report_path": "",
        }

        self.store.start_job(first["id"])
        self.store.complete_individual_job(first["id"], result)
        self.store.start_job(second["id"])
        self.store.complete_individual_job(second["id"], {**result, "roof_squares": 126})

        with self.store.connect() as connection:
            property_count = connection.execute("SELECT count(*) FROM properties").fetchone()[0]
            report_count = connection.execute("SELECT count(*) FROM roof_intelligence_reports").fetchone()[0]
            current_squares = connection.execute("SELECT roof_squares FROM properties").fetchone()[0]

        self.assertEqual(property_count, 1)
        self.assertEqual(report_count, 2)
        self.assertEqual(current_squares, 126)
        self.assertEqual(len(self.store.list_notifications()), 2)

    def test_versioned_report_preserves_initial_pdf_and_image_assets(self):
        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        self.store.start_job(job["id"])
        source_pdf = Path(self.temp_dir.name) / "generated.pdf"
        source_image = Path(self.temp_dir.name) / "aerial.jpg"
        source_analysis_image = Path(self.temp_dir.name) / "aerial-target.png"
        source_pdf.write_bytes(b"%PDF-1.4\n" + b"x" * 1200)
        source_image.write_bytes(b"image-bytes")
        source_analysis_image.write_bytes(b"analysis-image-bytes")
        report_id = "report-versioned-1"
        snapshot = {
            "snapshot_id": "revision-versioned-1",
            "report_id": report_id,
            "revision": {
                "number": 1,
                "kind": "initial",
                "created_at": "2026-07-22T12:00:00+00:00",
                "created_by": "local-user",
                "change_reason": "Fresh assessment",
            },
            "imagery": {
                "local_report_image_path": str(source_image),
                "local_analysis_image_path": str(source_analysis_image),
            },
        }
        result = {
            "report_id": report_id,
            "report_snapshot": snapshot,
            "address": "65 N Yuma St",
            "city": "Denver",
            "state": "CO",
            "zip": "80223",
            "county": "Denver",
            "parcel": "0508500065000",
            "report_path": str(source_pdf),
            "aerial_image_file": str(source_image),
            "analysis_aerial_image_file": str(source_analysis_image),
        }

        self.store.complete_individual_job(job["id"], result)

        saved_report = self.store.get_report(report_id)
        revision = self.store.get_report_revision("revision-versioned-1")
        self.assertTrue(Path(saved_report["report_path"]).is_file())
        self.assertNotEqual(Path(saved_report["report_path"]), source_pdf)
        self.assertEqual(Path(saved_report["report_path"]).name, "65 N Yuma St Denver.pdf")
        self.assertTrue(Path(revision["snapshot"]["imagery"]["local_report_image_path"]).is_file())
        self.assertTrue(Path(revision["snapshot"]["imagery"]["local_analysis_image_path"]).is_file())
        self.assertIn(
            "analysis-mask",
            Path(revision["snapshot"]["imagery"]["local_analysis_image_path"]).name,
        )
        self.assertEqual(result["aerial_image_file"], str(source_image))

    def test_manual_revision_retains_history_and_sets_future_area_override(self):
        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        self.store.start_job(job["id"])
        original_pdf = Path(self.temp_dir.name) / "original.pdf"
        original_pdf.write_bytes(b"%PDF-1.4\n" + b"a" * 1200)
        report_id = "report-history-1"
        original = {
            "schema_version": 1,
            "snapshot_id": "snapshot-history-1",
            "report_id": report_id,
            "revision": {
                "number": 1, "kind": "initial", "created_at": "2026-07-22T12:00:00+00:00",
                "created_by": "local-user", "change_reason": "Fresh assessment",
            },
            "analysis": {"roof_type": "TPO", "overall_score": 75, "risk_level": "MODERATE"},
            "calculations": {"roof_area_sqft": 10_000, "roof_squares": 100},
            "imagery": {},
        }
        self.store.complete_individual_job(job["id"], {
            "report_id": report_id, "report_snapshot": original,
            "address": "65 N Yuma St", "city": "Denver", "state": "CO", "zip": "80223",
            "county": "Denver", "parcel": "0508500065000", "report_path": str(original_pdf),
        })
        revised_pdf = Path(self.temp_dir.name) / "revised.pdf"
        revised_pdf.write_bytes(b"%PDF-1.4\n" + b"b" * 1200)
        revised = {
            **original,
            "snapshot_id": "snapshot-history-2",
            "revision": {
                "number": 2, "kind": "manual_edit", "parent_snapshot_id": "snapshot-history-1",
                "created_at": "2026-07-23T12:00:00+00:00", "created_by": "local-user",
                "change_reason": "Field measurement corrected area",
            },
            "analysis": {"roof_type": "TPO", "overall_score": 75, "risk_level": "MODERATE"},
            "calculations": {"roof_area_sqft": 12_500, "roof_squares": 125},
        }

        self.store.save_ready_report_revision(
            report_id, "snapshot-history-1", revised,
            report_path=str(revised_pdf), pdf_size=revised_pdf.stat().st_size,
            pdf_checksum="a" * 64, created_by="local-user",
            change_reason="Field measurement corrected area",
            edits={"roof_area_sqft": 12_500}, apply_square_footage_to_future=True,
        )

        revisions = self.store.list_report_revisions(report_id)
        self.assertEqual([item["revision_number"] for item in revisions], [2, 1])
        self.assertTrue(Path(revisions[1]["report_path"]).is_file())
        self.assertEqual(self.store.get_report(report_id)["report_path"], str(revised_pdf))
        override = self.store.get_active_square_footage_override(
            address="65 N Yuma St, Denver, CO 80223"
        )
        self.assertEqual(override["numeric_value"], 12_500)
        self.assertEqual(override["source_revision_id"], "snapshot-history-2")

    def test_failure_stores_concise_error_and_notification(self):
        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")

        failed = self.store.fail_job(
            job["id"],
            "imagery_unavailable",
            "County imagery service did not return a usable image.",
            retryable=True,
        )

        self.assertEqual(failed["status"], "failed")
        self.assertTrue(failed["retryable"])
        self.assertEqual(failed["error_code"], "imagery_unavailable")
        self.assertEqual(self.store.list_notifications()[0]["kind"], "job_failed")

    def test_footprint_resolution_is_audited_and_requeues_individual_job(self):
        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        details = {
            "county": "Denver",
            "parcel": "123",
            "footprint_validation": {"primary_sqft": 1000, "secondary_sqft": 1200},
        }
        self.store.fail_job(
            job["id"], "footprint_discrepancy", "Footprints differ.", error_details=details
        )

        resolved = self.store.resolve_footprint_discrepancy(
            job["id"], "county", "County geometry follows the assessor sketch."
        )

        self.assertEqual(resolved["status"], "queued")
        self.assertEqual(resolved["input"]["footprint_override"]["selected_source"], "county")
        self.assertEqual(resolved["error_details"], {})
        with self.store.connect() as connection:
            audit = connection.execute("SELECT * FROM footprint_resolutions").fetchone()
        self.assertEqual(audit["selected_source"], "county")
        self.assertEqual(audit["parcel_number"], "123")

    def test_health_history_notifies_only_on_failure_or_state_change(self):
        healthy = {
            "checked_at": "2026-07-17T10:00:00Z",
            "results": [{"county": "Denver", "status": "ok", "error": ""}],
        }
        failed = {
            "checked_at": "2026-07-17T11:00:00Z",
            "results": [{"county": "Denver", "status": "failed", "error": "imagery offline"}],
        }
        recovered = {
            "checked_at": "2026-07-17T12:00:00Z",
            "results": [{"county": "Denver", "status": "ok", "error": ""}],
        }

        self.store.record_county_health(healthy)
        self.assertEqual(self.store.list_notifications(), [])
        self.store.record_county_health(failed)
        self.store.record_county_health(failed)
        self.assertEqual(len(self.store.list_notifications()), 1)
        self.store.record_county_health(recovered)
        notifications = self.store.list_notifications()
        self.assertEqual(len(notifications), 2)
        self.assertIn("recovered", " ".join(item["title"].lower() for item in notifications))
        with self.store.connect() as connection:
            history_count = connection.execute("SELECT count(*) FROM county_health_checks").fetchone()[0]
        self.assertEqual(history_count, 4)
        self.assertEqual(len(self.store.list_county_health()), 4)
        self.assertEqual(self.store.list_county_health()[0]["result"]["county"], "Denver")
        self.assertEqual(self.store.list_county_health()[0]["status"], "healthy")

    def test_degraded_health_creates_distinct_warning_and_can_recover(self):
        degraded = {
            "checked_at": "2026-07-17T10:00:00Z",
            "results": [
                {
                    "county": "Denver",
                    "status": "degraded",
                    "error": "second address imagery unavailable",
                    "sample_count": 2,
                    "passed_count": 1,
                    "failed_count": 1,
                }
            ],
        }
        healthy = {
            "checked_at": "2026-07-17T11:00:00Z",
            "results": [{"county": "Denver", "status": "healthy", "error": ""}],
        }

        self.store.record_county_health(degraded)
        notifications = self.store.list_notifications()
        self.assertEqual(len(notifications), 1)
        self.assertIn("degraded", notifications[0]["title"].lower())

        self.store.record_county_health(healthy)
        notifications = self.store.list_notifications()
        self.assertEqual(len(notifications), 2)
        self.assertIn("recovered", " ".join(item["title"].lower() for item in notifications))

    def test_health_history_pruning_retains_recent_rows(self):
        payload = {
            "checked_at": "2020-01-01T00:00:00+00:00",
            "results": [{"county": "Denver", "status": "ok", "error": ""}],
        }
        self.store.record_county_health(payload)

        self.assertEqual(self.store.list_county_health(), [])

    def test_latest_health_display_returns_one_result_per_county(self):
        self.store.record_county_health(
            {
                "checked_at": "2026-07-17T10:00:00Z",
                "results": [
                    {"county": "Denver", "status": "failed", "error": "old failure"},
                    {"county": "Adams County", "status": "healthy", "error": ""},
                ],
            }
        )
        self.store.record_county_health(
            {
                "checked_at": "2026-07-18T10:00:00Z",
                "results": [
                    {"county": "Denver", "status": "healthy", "error": ""},
                    {"county": "Adams County", "status": "degraded", "error": "one sample failed"},
                ],
            }
        )

        latest = self.store.list_latest_county_health()

        self.assertEqual(len(latest), 2)
        self.assertEqual({item["county_key"] for item in latest}, {"DENVER", "ADAMS COUNTY"})
        self.assertTrue(all(item["checked_at"] == "2026-07-18T10:00:00Z" for item in latest))
        self.assertEqual(
            next(item for item in latest if item["county_key"] == "DENVER")["status"],
            "healthy",
        )

    def test_assessor_warnings_create_user_notification(self):
        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        self.store.start_job(job["id"])

        self.store.complete_individual_job(
            job["id"],
            {
                "address": "65 N Yuma St",
                "county": "Adams County",
                "parcel": "123",
                "report_path": "",
                "assessor_warnings": ["One assessor detail source had no matching record."],
            },
        )

        notifications = self.store.list_notifications(job_id=job["id"])
        self.assertEqual({item["kind"] for item in notifications}, {"job_completed", "assessor_warning"})
        warning = next(item for item in notifications if item["kind"] == "assessor_warning")
        self.assertIn("no matching record", warning["message"])

    def test_notifications_can_be_scoped_to_selected_job(self):
        first = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        second = self.store.create_individual_job("100 W 14th Ave, Denver, CO 80204")
        self.store.fail_job(first["id"], "first_error", "First job failed.")
        self.store.fail_job(second["id"], "second_error", "Second job failed.")

        selected = self.store.list_notifications(job_id=first["id"])

        self.assertEqual(len(selected), 1)
        self.assertEqual(selected[0]["job_id"], first["id"])
        self.assertEqual(selected[0]["message"], "First job failed.")

    def test_worker_claims_individual_jobs_in_fifo_order(self):
        first = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        second = self.store.create_individual_job("100 W 14th Ave, Denver, CO 80204")

        claimed = self.store.claim_next_individual_job()

        self.assertEqual(claimed["id"], first["id"])
        self.assertEqual(claimed["status"], "running")
        self.assertEqual(self.store.get_job(second["id"])["status"], "queued")

    def test_worker_recovers_interrupted_local_job(self):
        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        self.store.claim_next_individual_job()

        recovered = self.store.recover_interrupted_individual_jobs()

        self.assertEqual(recovered, 1)
        self.assertEqual(self.store.get_job(job["id"])["status"], "queued")

    def test_area_worker_claims_candidates_and_persists_reports(self):
        job = self.store.create_area_job(
            "39.75", "39.72", "-104.97", "-105.02", "100", ["TPO"]
        )
        claimed_job = self.store.claim_next_area_job()
        candidates = [
            {
                "candidate_key": "denver:one",
                "address": "100 Test St, Denver, CO 80202",
                "county": "Denver",
                "county_profile": "denver",
                "parcel": "one",
                "roof_area_sqft": 12000,
            },
            {
                "candidate_key": "adams:two",
                "address": "200 Test St, Thornton, CO 80229",
                "county": "Adams County",
                "county_profile": "adams",
                "parcel": "two",
                "roof_area_sqft": 15000,
            },
        ]
        self.store.prepare_area_candidates(job["id"], candidates)

        first = self.store.claim_next_area_item(job["id"])
        report_path = Path(self.temp_dir.name) / "area-report.pdf"
        report_path.write_bytes(b"%PDF-1.4 area test")
        self.store.complete_area_item(
            job["id"],
            first["id"],
            {
                **first["input"],
                "report_path": str(report_path),
                "roof_type": "TPO",
                "condition_score": 82,
            },
        )
        second = self.store.claim_next_area_item(job["id"])
        self.store.skip_area_item(job["id"], second["id"], "roof_type_excluded", "Not selected.")
        completed = self.store.finish_area_job(job["id"])

        self.assertEqual(claimed_job["status"], "running")
        self.assertEqual(completed["status"], "completed")
        self.assertEqual(completed["candidate_count"], 2)
        self.assertEqual(completed["completed_count"], 1)
        self.assertEqual(completed["skipped_count"], 1)
        self.assertEqual(len(self.store.get_reports_for_job(job["id"])), 1)

    def test_area_worker_recovers_running_item_and_cancel_updates_counts(self):
        job = self.store.create_area_job(
            "39.75", "39.72", "-104.97", "-105.02", "100", ["All"]
        )
        self.store.claim_next_area_job()
        self.store.prepare_area_candidates(
            job["id"],
            [
                {"candidate_key": "denver:one", "address": "100 Test St, Denver, CO 80202"},
                {"candidate_key": "denver:two", "address": "200 Test St, Denver, CO 80202"},
            ],
        )
        self.store.claim_next_area_item(job["id"])

        recovered = self.store.recover_interrupted_area_jobs()
        self.assertEqual(recovered, 1)
        self.assertEqual(self.store.get_job(job["id"])["status"], "queued")

        self.store.claim_next_area_job()
        cancelled = self.store.cancel_job(job["id"])
        self.assertEqual(cancelled["status"], "cancelled")
        self.assertEqual(cancelled["skipped_count"], 2)
        self.assertEqual(cancelled["remaining_count"], 0)

    def test_default_store_migrates_legacy_database_once(self):
        legacy_path = Path(self.temp_dir.name) / "legacy" / "roof-jobs.sqlite3"
        durable_path = Path(self.temp_dir.name) / "application-support" / "roof-jobs.sqlite3"
        legacy_store = RoofIntelligenceJobStore(legacy_path)
        legacy_job = legacy_store.create_individual_job("65 N Yuma St, Denver, CO 80223")

        with patch.dict(os.environ, {}, clear=False), \
             patch.object(roof_intelligence_jobs, "LEGACY_DB_PATH", legacy_path), \
             patch.object(roof_intelligence_jobs, "DEFAULT_DB_PATH", durable_path):
            os.environ.pop("ROOF_INTELLIGENCE_DB_PATH", None)
            migrated_store = RoofIntelligenceJobStore()

        self.assertTrue(durable_path.is_file())
        self.assertEqual(migrated_store.get_job(legacy_job["id"])["status"], "queued")


class AreaCandidateTests(unittest.TestCase):
    class Profile:
        key = "denver"
        display_name = "Denver"

    class Collector:
        @staticmethod
        def add_output_fields(record):
            record.setdefault("property_address", record.get("address", ""))
            record.setdefault("property_zip", record.get("zip", ""))

        @staticmethod
        def address_from_record(record):
            return record.get("address", "")

        @staticmethod
        def parcel_zip(record):
            return record.get("zip", "")

        @staticmethod
        def parcel_join_key(record):
            return record.get("parcel", "")

    def test_candidate_filters_roof_size_and_retains_age_metadata(self):
        record = {
            "address": "100 Test St",
            "property_city": "Denver",
            "property_state": "CO",
            "zip": "80202",
            "parcel": "one",
            "roof_area_est": 12500,
            "year_built": 1990,
        }

        candidate = area_batch.candidate_from_record(record, self.Profile(), self.Collector(), 10000)

        self.assertEqual(candidate["candidate_key"], "denver:one")
        self.assertEqual(candidate["address"], "100 Test St, Denver, CO 80202")
        self.assertGreaterEqual(candidate["age_estimate_years"], 20)
        self.assertIsNone(
            area_batch.candidate_from_record({**record, "roof_area_est": 9000}, self.Profile(), self.Collector(), 10000)
        )

    def test_candidate_reverse_geocodes_counties_without_situs_fields(self):
        record = {
            "parcel": "195917222900",
            "roof_area_est": 12500,
        }
        location = {
            "street": "405 Argentine St",
            "city": "Georgetown",
            "state": "CO",
            "zip": "80444",
        }
        collector = self.Collector()

        with patch.object(area_batch, "reverse_geocode_record", return_value=location) as geocode:
            candidate = area_batch.candidate_from_record(
                record, self.Profile(), collector, 10000
            )

        geocode.assert_called_once_with(record, collector)
        self.assertEqual(candidate["address"], "405 Argentine St, Georgetown, CO 80444")
        self.assertEqual(candidate["candidate_key"], "denver:195917222900")

    def test_radius_selection_uses_closed_arcgis_polygon(self):
        geometry_text, geometry_type = area_batch.spatial_query(
            {"north": 40, "south": 39, "east": -104, "west": -105},
            {
                "selection_type": "radius",
                "center": {"lat": 39.7392, "lng": -104.9903},
                "radius_miles": 0.1,
            },
        )
        geometry = __import__("json").loads(geometry_text)
        ring = geometry["rings"][0]

        self.assertEqual(geometry_type, "esriGeometryPolygon")
        self.assertEqual(geometry["spatialReference"]["wkid"], 4326)
        self.assertEqual(ring[0], ring[-1])
        self.assertEqual(len(ring), 73)

    def test_englewood_radius_only_queries_intersecting_counties(self):
        class Profile:
            def __init__(self, key):
                self.key = key

        bounds = {
            "north": 39.673783279218995,
            "south": 39.659310120781,
            "east": -105.00241246615624,
            "west": -105.02121433384374,
        }
        profiles = [Profile(key) for key in area_batch.COUNTY_WGS84_BOUNDS]

        applicable = area_batch.profiles_for_bounds(profiles, bounds)

        self.assertEqual([profile.key for profile in applicable], ["arapahoe", "denver"])

    def test_unknown_county_profile_remains_enabled(self):
        profile = type("Profile", (), {"key": "future_county"})()
        bounds = {"north": 39.7, "south": 39.6, "east": -104.9, "west": -105.0}

        self.assertTrue(area_batch.profile_intersects_bounds(profile, bounds))

    def test_discovery_defers_secondary_footprint_comparison_when_primary_matches(self):
        profile = SimpleNamespace(key="denver", display_name="Denver")
        secondary_lookup = Mock(return_value=[])
        detailed_validation = Mock()
        collector = SimpleNamespace(
            REQUEST_TIMEOUT=120,
            REQUEST_ATTEMPTS=4,
            get_parcel_bounds_in_building_crs=Mock(return_value="selected-bounds"),
            collect_buildings=Mock(return_value=[{
                "parcel": "one",
                "roof_area_est": 12_500,
                "property_address": "100 Test St",
                "property_city": "Denver",
                "property_state": "CO",
                "property_zip": "80202",
            }]),
            collect_secondary_buildings=secondary_lookup,
            combine_data=lambda buildings, _parcels: buildings,
            parcel_join_key=lambda record: record.get("parcel", ""),
            add_output_fields=lambda _record: None,
            address_from_record=lambda record: record.get("property_address", ""),
            parcel_zip=lambda record: record.get("property_zip", ""),
            validate_building_footprint_sources=detailed_validation,
        )
        county_config = SimpleNamespace(COUNTY_PROFILES={"denver": profile})
        single_address = SimpleNamespace(
            configure_collector_for_county=lambda _collector, _profile: None
        )

        with (
            patch.dict(sys.modules, {
                "collect_county_buildings_with_parcels": collector,
                "county_config": county_config,
                "roof_intelligence_single_address": single_address,
            }),
            patch.object(
                area_batch,
                "fetch_parcels_in_bounds",
                return_value=[{"parcel": "one"}],
            ),
        ):
            candidates, warnings = area_batch.discover_candidates(
                Path("."),
                {"north": 39.7, "south": 39.6, "east": -104.9, "west": -105.0},
                10_000,
                2_000,
            )

        self.assertEqual(len(candidates), 1)
        self.assertEqual(
            candidates[0]["footprint_validation"], {"status": "deferred_to_report"}
        )
        self.assertEqual(warnings, [])
        secondary_lookup.assert_not_called()
        detailed_validation.assert_not_called()

    def test_radius_arcgis_geometry_is_sent_in_post_body(self):
        class Collector:
            REQUEST_ATTEMPTS = 1
            REQUEST_TIMEOUT = 5

        params = {
            "where": "1=1",
            "geometry": '{"rings":[[[-104.9,39.7],[-104.8,39.8],[-104.9,39.7]]]}',
            "geometryType": "esriGeometryPolygon",
            "f": "json",
        }
        with patch.object(area_batch, "urlopen", return_value=BytesIO(b'{"features": []}')) as opener:
            payload = area_batch.fetch_arcgis_post(Collector(), "https://example.test/query", params)

        request = opener.call_args.args[0]
        body = parse_qs(request.data.decode("utf-8"))
        self.assertEqual(payload, {"features": []})
        self.assertEqual(request.get_method(), "POST")
        self.assertEqual(body["geometryType"], ["esriGeometryPolygon"])
        self.assertEqual(body["geometry"], [params["geometry"]])
        self.assertNotIn("geometry=", request.full_url)

    def test_exact_selected_parcel_is_loaded_by_identifier(self):
        class Collector:
            PARCELS_URL = "parcels"

            @staticmethod
            def collect_parcel_fields():
                return ["SCHEDNUM", "SITUS_ADDRESS_LINE1"]

            @staticmethod
            def fetch_page(_url, where, _offset, _fields, return_geometry=True):
                self.assertEqual(where, "SCHEDNUM = '0119107010000'")
                self.assertTrue(return_geometry)
                return {"features": [{
                    "attributes": {"SCHEDNUM": "0119107010000", "SITUS_ADDRESS_LINE1": "4300 N FOREST ST"},
                    "geometry": {"rings": []},
                }]}

            @staticmethod
            def parcel_join_key(record):
                return record.get("SCHEDNUM", "")

            @staticmethod
            def geometry_to_wkt(_geometry):
                return "POLYGON EMPTY"

        parcel = single_address.collect_live_parcel_by_id("0119107010000", Collector())

        self.assertEqual(parcel["full_parcel_number"], "0119107010000")
        self.assertEqual(parcel["parcel_geometry"], "POLYGON EMPTY")

    def test_batch_mode_does_not_pause_for_pending_footprint_review(self):
        validation = {"status": "discrepancy", "difference_pct": 20.0}

        self.assertTrue(single_address.should_pause_for_footprint_review(validation, "auto", False))
        self.assertFalse(single_address.should_pause_for_footprint_review(validation, "auto", True))


class CountyResolutionTests(unittest.TestCase):
    class Profile:
        def __init__(self, key):
            self.key = key
            self.display_name = key.title()
            self.building_url = f"{key}-buildings"
            self.parcel_url = f"{key}-parcels"
            self.imagery_sources = ()

    class Collector:
        PARCELS_URL = ""
        BUILDINGS_URL = ""
        _COLLECT_PARCEL_FIELDS = None
        _COLLECT_BUILDING_FIELDS = None

        @staticmethod
        def init_crs_transformers(_buildings, _parcels):
            return None

        @staticmethod
        def collect_parcel_fields():
            return ["SITUS_ADDRESS_LINE1", "SITUS_ZIP"]

        @staticmethod
        def parcel_zip_where(zip_codes):
            return f"SITUS_ZIP LIKE '{next(iter(zip_codes))}%'"

        @staticmethod
        def fetch_page(url, _where, _offset, _fields, return_geometry=True):
            features = []
            if url == "arapahoe-parcels":
                features = [{"attributes": {
                    "SITUS_ADDRESS_LINE1": "123 MAIN ST",
                    "SITUS_ZIP": "80012",
                    "SCHEDNUM": "A-1",
                }, "geometry": None}]
            return {"features": features}

        @staticmethod
        def geometry_to_wkt(_geometry):
            return ""

        @staticmethod
        def parcel_join_key(record):
            return record.get("SCHEDNUM", "")

        @staticmethod
        def parcel_zip(record):
            return record.get("SITUS_ZIP", "")

        @staticmethod
        def address_from_record(record):
            return record.get("SITUS_ADDRESS_LINE1", "")

    def test_zip_and_address_resolve_matching_county_profile(self):
        profiles = {
            key: self.Profile(key)
            for key in ("denver", "adams", "arapahoe", "jefferson")
        }

        with patch.object(
            single_address,
            "geocode_address_location",
            return_value={
                "longitude": -104.82,
                "latitude": 39.66,
                "county": "Arapahoe County",
                "postal_code": "80012",
            },
        ):
            profile, parcel, score, _, source = single_address.resolve_county_and_parcel(
                "123 Main St, Aurora, CO 80012",
                self.Collector,
                profiles,
            )

        self.assertEqual(profile.key, "arapahoe")
        self.assertEqual(parcel["SCHEDNUM"], "A-1")
        self.assertEqual(score, 1.0)
        self.assertIn("Arapahoe", source)

    def test_zip_validated_geocode_shortlists_target_and_boundary_counties(self):
        profiles = {
            key: self.Profile(key)
            for key in ("denver", "adams", "arapahoe", "jefferson")
        }

        with patch.object(
            single_address,
            "geocode_address_location",
            return_value={
                "longitude": -105.01,
                "latitude": 39.674,
                "county": "Arapahoe County",
                "postal_code": "80110",
            },
        ):
            shortlisted = single_address.shortlist_county_profiles(
                "1630 W Dartmouth Ave, Englewood, CO 80110",
                profiles,
            )

        self.assertEqual([profile.key for profile in shortlisted], ["arapahoe", "denver"])

    def test_mismatched_geocoder_zip_disables_shortlist(self):
        profiles = {"arapahoe": self.Profile("arapahoe")}

        with patch.object(
            single_address,
            "geocode_address_location",
            return_value={
                "longitude": -104.82,
                "latitude": 39.66,
                "county": "Arapahoe County",
                "postal_code": "80013",
            },
        ):
            shortlisted = single_address.shortlist_county_profiles(
                "123 Main St, Aurora, CO 80012",
                profiles,
            )

        self.assertEqual(shortlisted, [])

    def test_county_resolution_falls_back_after_shortlist_misses(self):
        profiles = {
            key: self.Profile(key)
            for key in ("denver", "arapahoe", "jefferson")
        }
        queried_urls = []
        original_fetch_page = self.Collector.fetch_page

        def tracked_fetch_page(url, where, offset, fields, return_geometry=True):
            queried_urls.append(url)
            return original_fetch_page(url, where, offset, fields, return_geometry)

        with (
            patch.object(
                single_address,
                "shortlist_county_profiles",
                return_value=[profiles["denver"]],
            ),
            patch.object(self.Collector, "fetch_page", side_effect=tracked_fetch_page),
        ):
            profile, parcel, *_ = single_address.resolve_county_and_parcel(
                "123 Main St, Aurora, CO 80012",
                self.Collector,
                profiles,
            )

        self.assertEqual(profile.key, "arapahoe")
        self.assertEqual(parcel["SCHEDNUM"], "A-1")
        self.assertEqual(
            list(dict.fromkeys(queried_urls)),
            ["denver-parcels", "arapahoe-parcels", "jefferson-parcels"],
        )

    def test_explicit_county_uses_live_parcel_service_only(self):
        profile = self.Profile("arapahoe")
        single_address.configure_collector_for_county(self.Collector, profile)

        parcel, score, _, source = single_address.find_parcel_live(
            "123 Main St, Aurora, CO 80012",
            self.Collector,
            profile.display_name,
        )

        self.assertEqual(parcel["SCHEDNUM"], "A-1")
        self.assertEqual(score, 1.0)
        self.assertEqual(source, "Live Arapahoe parcel service")

    def test_explicit_county_does_not_fall_back_when_live_lookup_is_empty(self):
        profile = self.Profile("denver")
        single_address.configure_collector_for_county(self.Collector, profile)

        with patch.object(single_address, "collect_live_parcels_for_address", return_value=[]):
            with self.assertRaisesRegex(RuntimeError, "No live Denver parcel match"):
                single_address.find_parcel_live(
                    "123 Main St, Denver, CO 80202",
                    self.Collector,
                    profile.display_name,
                )

    def test_zip_extraction_uses_final_code_after_five_digit_street_number(self):
        self.assertEqual(
            single_address.address_zip("12364 W Alameda Pkwy, Lakewood, CO 80228"),
            "80228",
        )
        self.assertNotEqual(
            single_address.normalize_address("12364 W Alameda Pkwy, Lakewood, CO 80228"),
            single_address.normalize_address("12850 W Alameda Pkwy"),
        )

    def test_county_query_uses_street_name_instead_of_direction(self):
        clauses = single_address.live_address_where_clauses(
            "4201 E 72nd Ave, Commerce City, CO 80022",
            self.Collector,
        )

        self.assertIn("4201%72ND%", clauses[0])

    def test_fuzzy_score_rejects_different_street_name_with_same_prefix(self):
        self.assertEqual(
            single_address.score_address("1704 HIGH ST", "1704 HIGGINS ST"),
            0.0,
        )

    @patch.object(single_address, "collect_spatial_parcel_for_address")
    @patch.object(single_address, "collect_live_parcels_for_address")
    def test_live_lookup_falls_back_to_spatial_parcel_after_bad_text_match(
        self,
        text_lookup_mock,
        spatial_lookup_mock,
    ):
        text_lookup_mock.return_value = [{
            "SITUS_ADDRESS_LINE1": "1704 HIGGINS ST",
            "SITUS_ZIP": "",
            "SCHEDNUM": "WRONG",
        }]
        spatial_lookup_mock.return_value = [{
            "attributes": {
                "SITUS_ADDRESS_LINE1": "1720 N HIGH ST",
                "SITUS_ZIP": "80218",
                "SCHEDNUM": "RIGHT",
            },
            "geometry": None,
        }]

        parcel, score, matched_address, _ = single_address.find_parcel_live(
            "1704 High St, Denver, CO 80218",
            self.Collector,
            "Denver",
        )

        self.assertEqual(parcel["SCHEDNUM"], "RIGHT")
        self.assertEqual(score, 0.9)
        self.assertEqual(matched_address, "1704 High St")

    def test_single_spatial_parcel_is_accepted_without_situs_address(self):
        parcel = {"PIN": "195917222900", "full_parcel_number": "195917222900"}

        selected, score, matched_address = single_address.find_parcel_for_address(
            "405 Argentine St, Georgetown, CO 80444",
            [parcel],
            self.Collector,
        )

        self.assertIs(selected, parcel)
        self.assertEqual(score, 0.9)
        self.assertEqual(matched_address, "405 Argentine St")


class RoofIntelligenceRouteTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import pcs_proposal_web

        cls.web = pcs_proposal_web

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "route-jobs.sqlite3"
        self.settings_path = Path(self.temp_dir.name) / "settings.json"
        self.previous_db = os.environ.get("ROOF_INTELLIGENCE_DB_PATH")
        self.previous_worker = os.environ.get("ROOF_INTELLIGENCE_LOCAL_WORKER")
        self.previous_settings = os.environ.get("PCS_SETTINGS_PATH")
        os.environ["ROOF_INTELLIGENCE_DB_PATH"] = str(self.db_path)
        os.environ["ROOF_INTELLIGENCE_LOCAL_WORKER"] = "0"
        os.environ["PCS_SETTINGS_PATH"] = str(self.settings_path)
        self.previous_health_process = self.web._county_health_process
        self.web._county_health_process = None
        self.client = self.web.app.test_client()

    def tearDown(self):
        if self.previous_db is None:
            os.environ.pop("ROOF_INTELLIGENCE_DB_PATH", None)
        else:
            os.environ["ROOF_INTELLIGENCE_DB_PATH"] = self.previous_db
        if self.previous_worker is None:
            os.environ.pop("ROOF_INTELLIGENCE_LOCAL_WORKER", None)
        else:
            os.environ["ROOF_INTELLIGENCE_LOCAL_WORKER"] = self.previous_worker
        if self.previous_settings is None:
            os.environ.pop("PCS_SETTINGS_PATH", None)
        else:
            os.environ["PCS_SETTINGS_PATH"] = self.previous_settings
        self.web._county_health_process = self.previous_health_process
        self.temp_dir.cleanup()

    def test_map_candidate_passes_exact_parcel_to_report_worker(self):
        report_path = Path(self.temp_dir.name) / "selected-parcel.pdf"
        report_path.write_bytes(b"%PDF-1.4 test")
        completed = type(
            "Completed",
            (),
            {
                "returncode": 0,
                "stdout": '{"report_path": "' + str(report_path) + '"}\n',
                "stderr": "",
            },
        )()

        with patch.object(self.web.subprocess, "run", return_value=completed) as run:
            self.web._run_local_candidate_report(
                {
                    "address": "4201 E 72nd Ave, Commerce City, CO 80022",
                    "county_profile": "adams",
                    "parcel": "0172131300018",
                }
            )

        command = run.call_args.args[0]
        self.assertIn("--county", command)
        self.assertEqual(command[command.index("--county") + 1], "adams")
        self.assertIn("--parcel-id", command)
        self.assertEqual(command[command.index("--parcel-id") + 1], "0172131300018")
        self.assertIn("--allow-pending-footprint-review", command)
        self.assertNotIn("--parcel-cache", command)

    def _create_editable_report(self, report_id="route-edit-report"):
        store = RoofIntelligenceJobStore(self.db_path)
        job = store.create_individual_job("65 N Yuma St, Denver, CO 80223")
        store.start_job(job["id"])
        pdf_path = Path(self.temp_dir.name) / f"{report_id}.pdf"
        pdf_path.write_bytes(b"%PDF-1.4\n" + b"x" * 1200)
        snapshot = {
            "schema_version": 1,
            "snapshot_id": f"{report_id}-revision-1",
            "report_id": report_id,
            "revision": {
                "number": 1, "kind": "initial", "parent_snapshot_id": None,
                "created_at": "2026-07-22T12:00:00+00:00", "created_by": "local-user",
                "change_reason": "Fresh assessment", "manual_edits": {},
            },
            "property": {"canonical_key": "DENVER:0508500065000", "address": "65 N Yuma St"},
            "report_fields": {"roof_area_sqft": 10_000, "roof_squares": 100},
            "analysis": {
                "roof_type": "TPO", "roof_system": "Single-ply membrane", "overall_score": 75,
                "condition_label": "FAIR", "risk_level": "MODERATE",
                "summary": "Fresh summary", "recommendation": "Fresh recommendation",
            },
            "imagery": {"source": "Synthetic", "limitations": []},
            "calculations": {
                "roof_area_sqft": 10_000, "roof_squares": 100,
                "roof_condition_score": 75, "condition_label": "FAIR", "risk_level": "MODERATE",
            },
            "provenance": {"manual_fields": [], "persistent_square_footage_override": False},
        }
        store.complete_individual_job(job["id"], {
            "report_id": report_id, "report_snapshot": snapshot,
            "address": "65 N Yuma St", "city": "Denver", "state": "CO", "zip": "80223",
            "county": "Denver", "parcel": "0508500065000", "report_path": str(pdf_path),
            "roof_type": "TPO", "condition_score": 75, "risk_level": "MODERATE",
        })
        return store, job, snapshot

    def test_report_review_route_is_hidden_until_both_flags_are_enabled(self):
        store, job, _ = self._create_editable_report()

        disabled = self.client.get("/roof-intelligence/reports/route-edit-report/review")
        self.assertEqual(disabled.status_code, 404)

        with patch.dict(os.environ, {
            "ROOF_INTELLIGENCE_SUPABASE_ENABLED": "1",
            "ROOF_INTELLIGENCE_REPORT_EDITING_ENABLED": "1",
        }):
            enabled = self.client.get("/roof-intelligence/reports/route-edit-report/review")
            workspace = self.client.get(f"/roof-intelligence?job_id={job['id']}")

        self.assertEqual(enabled.status_code, 200)
        self.assertIn(b"Review &amp; Edit Roof Report", enabled.data)
        self.assertIn(b"Revision 1 remains unchanged", enabled.data)
        self.assertIn(b"Roof type (roofing surface)", enabled.data)
        self.assertIn(b"Roof system (physical configuration)", enabled.data)
        self.assertIn(b"Submit this correction for review to improve future roof reports", enabled.data)
        self.assertIn(b"Review &amp; Edit", workspace.data)
        self.assertEqual(len(store.list_report_revisions("route-edit-report")), 1)

    def test_completed_report_opens_from_application_support_data_root(self):
        self._create_editable_report("application-support-report")

        with patch.object(self.web, "DEFAULT_DATA_DIR", Path(self.temp_dir.name)):
            response = self.client.get(
                "/roof-intelligence/reports/application-support-report"
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/pdf")
        response.close()

    def test_county_health_panel_has_manual_run_button(self):
        with patch.object(self.web, "_county_health_check_running", return_value=False):
            page = self.client.get("/roof-intelligence")

        self.assertEqual(page.status_code, 200)
        self.assertIn(b'id="county-health-run"', page.data)
        self.assertIn(b"/roof-intelligence/county-health/run", page.data)

    def test_manual_county_health_route_starts_background_check(self):
        with patch.object(
            self.web,
            "_start_manual_county_health_check",
            return_value=True,
        ) as start:
            response = self.client.post("/roof-intelligence/county-health/run")

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers["Location"], "/roof-intelligence")
        start.assert_called_once_with()

    def test_manual_county_health_uses_full_bounded_pilotpoint_check(self):
        project_dir = Path(self.temp_dir.name) / "PilotPoint"
        python_path = project_dir / ".venv" / "bin" / "python"
        script_path = project_dir / "county_discovery_health.py"
        python_path.parent.mkdir(parents=True)
        python_path.write_text("", encoding="utf-8")
        script_path.write_text("", encoding="utf-8")
        process = Mock()
        process.poll.return_value = None

        with (
            patch.dict(os.environ, {"ROOF_INTELLIGENCE_LOCAL_WORKER": "1"}),
            patch.object(self.web, "ROOF_INTELLIGENCE_PROJECT_DIR", str(project_dir)),
            patch.object(self.web, "DEFAULT_DATA_DIR", Path(self.temp_dir.name) / "data"),
            patch.object(self.web, "_roof_worker_readiness_error", return_value=None),
            patch.object(self.web.subprocess, "Popen", return_value=process) as popen,
        ):
            started = self.web._start_manual_county_health_check()

        self.assertTrue(started)
        command = popen.call_args.args[0]
        self.assertEqual(command[:2], [str(python_path), str(script_path)])
        self.assertIn("--all-samples", command)
        self.assertIn("--strict-discrepancies", command)
        self.assertIn("--notify-pcs", command)
        self.assertIn("--output", command)
        self.assertEqual(popen.call_args.kwargs["cwd"], str(project_dir))

    def test_manual_county_health_status_reports_running_state(self):
        with patch.object(self.web, "_county_health_check_running", return_value=True):
            response = self.client.get("/api/roof-intelligence/county-health/status")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"running": True})

    def test_edit_submission_creates_new_revision_and_recalculates(self):
        store, _, parent = self._create_editable_report("route-revision-report")

        def complete_revision(command, **kwargs):
            parent_path = Path(command[2])
            output_pdf = Path(command[4])
            output_snapshot = Path(command[5])
            submitted_parent = json.loads(parent_path.read_text(encoding="utf-8"))
            self.assertEqual(submitted_parent["snapshot_id"], parent["snapshot_id"])
            revised = json.loads(json.dumps(submitted_parent))
            revised["snapshot_id"] = "route-revision-report-revision-2"
            revised["revision"] = {
                "number": 2, "kind": "manual_edit", "parent_snapshot_id": parent["snapshot_id"],
                "created_at": "2026-07-23T12:00:00+00:00", "created_by": "local-user",
                "change_reason": "Field measurement corrected area",
                "manual_edits": {"roof_area_sqft": 12_500},
            }
            revised["report_fields"].update({"roof_area_sqft": 12_500, "roof_squares": 125})
            revised["calculations"].update({"roof_area_sqft": 12_500, "roof_squares": 125})
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            output_pdf.write_bytes(b"%PDF-1.4\n" + b"y" * 1200)
            output_snapshot.write_text(json.dumps(revised), encoding="utf-8")
            self.assertIn("--submit-for-future-processing", command)
            feedback_directory = Path(command[command.index("--feedback-directory") + 1])
            feedback_directory.mkdir(parents=True, exist_ok=True)
            feedback_path = feedback_directory / "feedback-route-1.json"
            feedback_path.write_text(
                json.dumps(
                    {
                        "schema_version": 1,
                        "feedback_id": "feedback-route-1",
                        "status": "pending_review",
                        "report_id": "route-revision-report",
                        "parent_snapshot_id": parent["snapshot_id"],
                        "revised_snapshot_id": revised["snapshot_id"],
                        "requested_by": "local-user",
                        "created_at": "2026-07-23T12:00:00+00:00",
                        "comment": "Field measurement corrected area",
                        "corrections": {"roof_area_sqft": {"before": 10000, "after": 12500}},
                        "learning_scopes": ["property_measurement"],
                        "property_identity": {"canonical_key": "DENVER:0508500065000"},
                        "imagery_identity": {"source": "Synthetic"},
                        "review": None,
                        "application": None,
                    }
                ),
                encoding="utf-8",
            )
            return type(
                "Completed",
                (),
                {
                    "returncode": 0,
                    "stdout": json.dumps({"processing_feedback_path": str(feedback_path)}),
                    "stderr": "",
                },
            )()

        with patch.dict(os.environ, {
            "ROOF_INTELLIGENCE_SUPABASE_ENABLED": "1",
            "ROOF_INTELLIGENCE_REPORT_EDITING_ENABLED": "1",
        }), patch.object(self.web, "DEFAULT_DATA_DIR", Path(self.temp_dir.name)), \
             patch.object(self.web.subprocess, "run", side_effect=complete_revision):
            response = self.client.post(
                "/roof-intelligence/reports/route-revision-report/revisions",
                data={
                    "roof_area_sqft": "12500", "roof_condition_score": "75",
                    "roof_type": "TPO", "roof_system": "Single-ply membrane",
                    "report_summary": "Fresh summary", "recommendation": "Fresh recommendation",
                    "change_reason": "Field measurement corrected area",
                    "apply_square_footage_to_future": "1",
                    "submit_for_future_processing": "1",
                },
                follow_redirects=False,
            )

        self.assertEqual(response.status_code, 302)
        revisions = store.list_report_revisions("route-revision-report")
        self.assertEqual([item["revision_number"] for item in revisions], [2, 1])
        self.assertEqual(
            Path(revisions[0]["report_path"]).name,
            "65 N Yuma St Denver.pdf",
        )
        self.assertEqual(Path(revisions[0]["report_path"]).parent.name, "revision-2")
        self.assertEqual(revisions[0]["snapshot"]["report_fields"]["roof_area_sqft"], 12_500)
        self.assertEqual(
            store.get_active_square_footage_override(address="65 N Yuma St")["numeric_value"],
            12_500,
        )
        feedback = store.list_processing_feedback(report_id="route-revision-report")
        self.assertEqual(len(feedback), 1)
        self.assertEqual(feedback[0]["status"], "pending_review")
        self.assertEqual(feedback[0]["feedback"]["comment"], "Field measurement corrected area")

    def test_canonical_review_queue_is_loaded_from_roof_project(self):
        completed = type("Completed", (), {
            "returncode": 0,
            "stdout": '[{"canonical_id": 7, "county": "Jefferson", "difference_pct": "29.000"}]',
            "stderr": "",
        })()
        with (
            patch.dict(os.environ, {"ROOF_INTELLIGENCE_LOCAL_WORKER": "1"}),
            patch.object(self.web.subprocess, "run", return_value=completed) as run,
        ):
            reviews = self.web._canonical_footprint_reviews()

        self.assertEqual(reviews[0]["canonical_id"], 7)
        self.assertIn("review_canonical_footprints.py", run.call_args.args[0][1])

    def test_canonical_review_displays_address_source_areas_and_dates(self):
        review = {
            "canonical_id": 7,
            "county": "Jefferson",
            "parcel_id": "4917105011",
            "requested_address": "12043 W Alameda Pkwy, Lakewood, CO 80228",
            "difference_pct": "29.000",
            "sources": [
                {
                    "source_type": "microsoft",
                    "footprint_sqft": "106894.42",
                    "source_updated_at": None,
                },
                {
                    "source_type": "county",
                    "footprint_sqft": "75891.80",
                    "source_updated_at": "2026-07-01T00:00:00+00:00",
                },
            ],
        }
        with patch.object(self.web, "_canonical_footprint_reviews", return_value=[review]):
            page = self.client.get("/roof-intelligence")

        self.assertEqual(page.status_code, 200)
        self.assertIn(b"12043 W Alameda Pkwy, Lakewood, CO 80228", page.data)
        self.assertIn(b"106,894 sq ft", page.data)
        self.assertIn(b"75,892 sq ft", page.data)
        self.assertIn(b"Information date: Not available", page.data)
        self.assertIn(b"Information date: 2026-07-01", page.data)
        self.assertIn(b'data-mode-button="canonical"', page.data)
        self.assertIn(b'id="canonical-review-dialog"', page.data)
        self.assertIn(b"1 footprint discrepancy requires review.", page.data)
        self.assertIn(b"review-status-required", page.data)

    def test_canonical_resolution_maps_supabase_to_microsoft_source(self):
        completed = type("Completed", (), {
            "returncode": 0, "stdout": '{"canonical_id": 7}', "stderr": "",
        })()
        with patch.object(self.web.subprocess, "run", return_value=completed) as run:
            result = self.web._resolve_canonical_footprint(7, "supabase", "Microsoft matches current imagery")

        command = run.call_args.args[0]
        self.assertEqual(result["canonical_id"], 7)
        self.assertEqual(command[command.index("--source") + 1], "microsoft")

    def test_individual_submission_redirects_to_persistent_job_and_status_api(self):
        initial_page = self.client.get("/roof-intelligence")
        self.assertEqual(initial_page.status_code, 200)
        self.assertIn(b"No footprint discrepancies to review.", initial_page.data)
        self.assertIn(b'data-mode-button="canonical"', initial_page.data)
        self.assertNotIn(b'id="canonical-review-dialog"', initial_page.data)
        self.assertIn(b"const inferredMode = null;", initial_page.data)
        self.assertIn(b'aria-label="Settings" title="Settings"', initial_page.data)
        self.assertNotIn(b'bi-gear"></i>Settings', initial_page.data)

        landing_page = self.client.get("/")
        self.assertEqual(landing_page.status_code, 200)
        self.assertIn(b'class="settings-icon-button"', landing_page.data)
        self.assertIn(b'aria-label="Settings" title="Settings"', landing_page.data)
        self.assertNotIn(b'class="work-card settings"', landing_page.data)
        self.assertNotIn(b"Open settings", landing_page.data)

        response = self.client.post(
            "/roof-intelligence/jobs/individual",
            data={"property_address": "65 N Yuma St, Denver, CO 80223"},
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        job_id = query["job_id"][0]

        status = self.client.get(f"/api/roof-intelligence/jobs/{job_id}")
        self.assertEqual(status.status_code, 200)
        payload = status.get_json()
        self.assertEqual(payload["status"], "queued")
        self.assertEqual(payload["input"]["property_address"], "65 N Yuma St, Denver, CO 80223")

        page = self.client.get(response.headers["Location"])
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Individual Address", page.data)
        self.assertIn(b"Canonical Footprint Review", page.data)
        self.assertIn(b"Map Area Batch", page.data)
        self.assertIn(b"job-clock", page.data)
        self.assertIn(b'data-finished-at=""', page.data)
        self.assertIn(b"const jobIsActive = activeStatuses.has(workspace.dataset.jobStatus)", page.data)
        self.assertIn(b"tick(jobIsActive || Number.isNaN(finishedTime) ? Date.now() : finishedTime)", page.data)
        self.assertIn(b"if (jobIsActive) window.setInterval(tick, 1000)", page.data)
        self.assertIn(b"waiting for an available PilotPoint IQ worker", page.data)
        self.assertNotIn(b"Report Request", page.data)
        self.assertIn(b'const inferredMode = "individual";', page.data)
        self.assertNotIn(b"&#39;individual&#39;", page.data)
        self.assertIn(b'aria-controls="individual-request-dialog" aria-haspopup="dialog"', page.data)
        self.assertIn(b'aria-controls="area-request-dialog" aria-haspopup="dialog"', page.data)
        self.assertIn(b'data-mode-container="individual"', page.data)
        self.assertIn(b'data-mode-container="area"', page.data)
        self.assertIn(b"grid-template-columns: repeat(2, minmax(220px, 320px)) minmax(0, 1fr)", page.data)
        self.assertIn(b'class="workflow-dialog" id="individual-request-dialog"', page.data)
        self.assertIn(b'class="workflow-dialog area-dialog" id="area-request-dialog"', page.data)
        self.assertIn(b'data-dialog-close', page.data)
        self.assertIn(b"if (dialog && !dialog.open) dialog.showModal()", page.data)
        self.assertIn(b"form.addEventListener('submit', () => form.closest('dialog')?.close())", page.data)
        self.assertIn(b"if (requestedMode && !workspace?.dataset.jobId) openModeDialog(mode)", page.data)
        self.assertIn(b'.workflow-dialog::backdrop', page.data)
        self.assertIn(b"height: clamp(390px, 55vh, 540px)", page.data)
        self.assertIn(b'id="individual-request-form"', page.data)
        self.assertIn(b'id="area-request-form"', page.data)
        self.assertIn(b'name="selection_type" value="rectangle" checked', page.data)
        self.assertIn(b'name="selection_type" value="radius"', page.data)
        self.assertIn(b'Address Radius', page.data)
        self.assertIn(b'id="radius_miles"', page.data)
        self.assertIn(b'type="range" min="0.1" max="10" step="0.1" value="1"', page.data)
        self.assertIn(b'id="radius-miles-value"', page.data)
        self.assertIn(b'0.1 mi', page.data)
        self.assertIn(b'10 mi', page.data)
        self.assertIn(b'radiusMiles.toFixed(1)', page.data)
        self.assertIn(b"setAttribute('aria-valuetext', label)", page.data)
        self.assertIn(b'id="center_lat"', page.data)
        self.assertIn(b'id="center_lng"', page.data)
        self.assertIn(b'id="center_address"', page.data)
        self.assertIn(b'new google.maps.Circle', page.data)
        self.assertIn(b'radiusMiles * 1609.344', page.data)
        self.assertIn(b'radiusMilesInput.required = radiusMode', page.data)
        self.assertIn(b'Minimum Roof Size (Squares)', page.data)
        self.assertIn(b'id="minimum_roof_squares"', page.data)
        self.assertIn(b'name="minimum_roof_squares"', page.data)
        self.assertIn(b'value="100"', page.data)
        self.assertIn(b"Math.ceil(number).toLocaleString('en-US')", page.data)
        self.assertIn(b'id="area-map"', page.data)
        self.assertNotIn(b'id="report_limit"', page.data)
        self.assertNotIn(b'id="minimum_age"', page.data)
        self.assertIn(b'id="roof-type-picker"', page.data)
        self.assertIn(b'id="roof-type-toggle"', page.data)
        self.assertIn(b'id="roof-type-menu"', page.data)
        self.assertIn(b'id="roof-type-summary"', page.data)
        self.assertIn(b"roofTypeSummary.textContent = `${selected.length} roof types selected`", page.data)
        self.assertIn(b"if (!roofTypePicker?.contains(event.target)) closeRoofTypeMenu()", page.data)
        self.assertIn(b'id="bounds_north"', page.data)
        self.assertIn(b"center: {lat: 39.7392, lng: -104.9903}", page.data)
        self.assertIn(b"new google.maps.Map", page.data)
        self.assertIn(b'id="map-address"', page.data)
        self.assertIn(b'id="map-address-suggestions"', page.data)
        self.assertIn(b'id="map-address-help"', page.data)
        self.assertIn(b'aria-controls="map-address-suggestions"', page.data)
        self.assertIn(b'id="property-address-suggestions"', page.data)
        self.assertIn(b"google.maps.importLibrary('places')", page.data)
        self.assertIn(b"AutocompleteSuggestion.fetchAutocompleteSuggestions", page.data)
        self.assertIn(b"includedRegionCodes: ['us']", page.data)
        self.assertIn(b"place.fetchFields({fields: ['formattedAddress']})", page.data)
        self.assertIn(b"place.fetchFields({fields: ['formattedAddress', 'location']})", page.data)
        self.assertIn(b"button.id = `map-address-option-${index}`", page.data)
        self.assertIn(b"useSelectedPlace(place.location", page.data)
        self.assertIn(b"event.defaultPrevented || event.key !== 'Enter'", page.data)
        self.assertIn(b"options[activeIndex >= 0 ? activeIndex : 0].click()", page.data)
        self.assertIn(b"Enter a complete address and choose Locate", page.data)
        self.assertIn(b"new Intl.DateTimeFormat(undefined", page.data)
        self.assertIn(b"timeZoneName: 'short'", page.data)
        self.assertIn(b'id="center-map-address-button"', page.data)
        self.assertIn(b"google.maps.importLibrary('geocoding')", page.data)
        self.assertIn(b"new Geocoder", page.data)
        self.assertIn(b"mapTypeId: google.maps.MapTypeId.HYBRID", page.data)
        self.assertIn(b"google.maps.MapTypeId.ROADMAP, google.maps.MapTypeId.HYBRID", page.data)
        self.assertIn(b"Map centered on ${match.formatted_address}", page.data)
        self.assertIn(b"gestureHandling: 'greedy'", page.data)
        self.assertIn(b"scrollwheel: true", page.data)
        self.assertIn(b"/api/local-settings/google-maps", page.data)
        self.assertIn(b"Google Maps is not configured", page.data)
        self.assertIn(b"scroll or pinch to zoom", page.data)
        self.assertIn(b"projection.fromContainerPixelToLatLng", page.data)
        self.assertIn(b"addEventListener('pointerdown'", page.data)
        self.assertIn(b"addEventListener('pointermove'", page.data)
        self.assertIn(b"addEventListener('pointerup'", page.data)
        self.assertIn(b"press at the first corner", page.data)
        self.assertIn(b'aria-pressed="false"', page.data)
        self.assertIn(b"No notifications for this Roof Intelligence job", page.data)

        unselected_page = self.client.get("/roof-intelligence")
        self.assertIn(b"Select a recent Roof Intelligence job", unselected_page.data)

    def test_google_maps_key_is_saved_locally_and_not_echoed_by_settings_page(self):
        key = "AIza" + ("A" * 35)
        missing = self.client.get("/api/local-settings/google-maps")
        self.assertEqual(missing.status_code, 404)

        response = self.client.post(
            "/settings",
            data={"action": "save_google_maps_key", "google_maps_api_key": key},
            follow_redirects=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Google Maps is configured", response.data)
        self.assertIn(b"ending in AAAA", response.data)
        self.assertNotIn(key.encode(), response.data)
        self.assertEqual(self.settings_path.stat().st_mode & 0o777, 0o600)

        configuration = self.client.get("/api/local-settings/google-maps")
        self.assertEqual(configuration.status_code, 200)
        self.assertEqual(configuration.get_json()["api_key"], key)
        self.assertEqual(configuration.headers["Cache-Control"], "no-store, private")

        roof_page = self.client.get("/roof-intelligence?mode=area")
        self.assertNotIn(b"Google Maps is not configured", roof_page.data)
        self.assertNotIn(key.encode(), roof_page.data)

    def test_area_submission_persists_bounds_and_filters(self):
        response = self.client.post(
            "/roof-intelligence/jobs/area",
            data={
                "bounds_north": "39.75",
                "bounds_south": "39.72",
                "bounds_east": "-104.97",
                "bounds_west": "-105.02",
                "minimum_roof_squares": "100.1",
                "roof_types": ["TPO", "Metal"],
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        payload = self.client.get(
            f"/api/roof-intelligence/jobs/{query['job_id'][0]}"
        ).get_json()
        self.assertIsNone(payload["minimum_age"])
        self.assertEqual(payload["minimum_roof_size"], 10100)
        self.assertEqual(payload["input"]["minimum_roof_squares"], 101)
        self.assertEqual(payload["roof_types"], ["TPO", "Metal"])
        self.assertIsNone(payload["report_limit"])
        self.assertEqual(payload["input"]["bounds"]["west"], -105.02)

        page = self.client.get(response.headers["Location"])
        self.assertIn(b'id="batch-report-list"', page.data)
        self.assertIn(b'id="batch-report-empty"', page.data)
        self.assertIn(b'id="batch-report-count"', page.data)
        self.assertIn(b"Completed reports will appear here as each property finishes.", page.data)
        self.assertIn(b"const displayedReportIds = new Set", page.data)
        self.assertIn(b"syncBatchReports(job.reports || [])", page.data)
        self.assertIn(b"document.createElement('li')", page.data)
        self.assertLess(
            page.data.index(b"syncBatchReports(job.reports || [])"),
            page.data.index(b"if (!activeStatuses.has(job.status)) window.location.reload()"),
        )

    def test_radius_submission_persists_center_distance_and_derived_bounds(self):
        response = self.client.post(
            "/roof-intelligence/jobs/area",
            data={
                "selection_type": "radius",
                "center_lat": "39.7392",
                "center_lng": "-104.9903",
                "center_address": "1701 Wynkoop St, Denver, CO 80202",
                "radius_miles": ".1",
                "minimum_roof_squares": "100",
                "roof_types": ["All"],
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        payload = self.client.get(
            f"/api/roof-intelligence/jobs/{query['job_id'][0]}"
        ).get_json()
        self.assertEqual(payload["input"]["selection_type"], "radius")
        self.assertEqual(payload["input"]["radius_miles"], 0.1)
        self.assertEqual(payload["input"]["center"]["lat"], 39.7392)
        self.assertGreater(payload["input"]["bounds"]["north"], 39.7392)

    def test_radius_worker_passes_circle_geometry_to_discovery_adapter(self):
        job = {
            "minimum_roof_size": 10000,
            "input": {
                "selection_type": "radius",
                "bounds": {"north": 39.75, "south": 39.72, "east": -104.97, "west": -105.02},
                "center": {"lat": 39.7392, "lng": -104.9903},
                "radius_miles": 0.1,
            },
        }
        with patch.object(self.web.subprocess, "run") as run:
            run.return_value.returncode = 0
            run.return_value.stdout = '{"candidates": [], "warnings": []}'
            run.return_value.stderr = ""
            self.web._discover_local_area_candidates(job)

        command = run.call_args.args[0]
        self.assertEqual(command[command.index("--selection-type") + 1], "radius")
        self.assertEqual(command[command.index("--center-lat") + 1], "39.7392")
        self.assertEqual(command[command.index("--center-lng") + 1], "-104.9903")
        self.assertEqual(command[command.index("--radius-miles") + 1], "0.1")

    def test_area_discovery_retries_transient_coordinate_system_failure(self):
        job = {
            "minimum_roof_size": 10000,
            "input": {
                "selection_type": "radius",
                "bounds": {"north": 39.75, "south": 39.72, "east": -104.97, "west": -105.02},
                "center": {"lat": 39.7392, "lng": -104.9903},
                "radius_miles": 0.1,
            },
        }
        failed = type("Completed", (), {
            "returncode": 1,
            "stdout": '{"error": "County parcel discovery failed: Unable to determine building/parcel coordinate systems; aborting before aerial imagery can be requested."}',
            "stderr": "",
        })()
        succeeded = type("Completed", (), {
            "returncode": 0,
            "stdout": '{"candidates": [{"candidate_key": "denver:one"}], "warnings": []}',
            "stderr": "",
        })()

        with patch.object(self.web.subprocess, "run", side_effect=[failed, succeeded]) as run, \
             patch.object(self.web.time, "sleep") as sleep:
            candidates, warnings = self.web._discover_local_area_candidates(job)

        self.assertEqual(candidates, [{"candidate_key": "denver:one"}])
        self.assertEqual(warnings, [])
        self.assertEqual(run.call_count, 2)
        sleep.assert_called_once_with(2.0)

    def test_coordinate_system_failure_is_not_reported_as_imagery_failure(self):
        code, message, retryable = self.web._roof_error_details(
            "Unable to determine building/parcel coordinate systems; "
            "aborting before aerial imagery can be requested at the wrong location."
        )

        self.assertEqual(code, "gis_service_unavailable")
        self.assertIn("County GIS services", message)
        self.assertTrue(retryable)

    def test_area_job_can_be_cancelled_from_workspace(self):
        store = RoofIntelligenceJobStore(self.db_path)
        job = store.create_area_job(
            "39.75", "39.72", "-104.97", "-105.02", "100", ["All"]
        )

        page = self.client.get(f"/roof-intelligence?job_id={job['id']}")
        self.assertIn(b"Cancel Job", page.data)
        response = self.client.post(
            f"/roof-intelligence/jobs/{job['id']}/cancel", follow_redirects=False
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(store.get_job(job["id"])["status"], "cancelled")

        completed_page = self.client.get(f"/roof-intelligence?job_id={job['id']}")
        self.assertIn(b'data-job-status="cancelled"', completed_page.data)
        self.assertIn(b'data-finished-at="20', completed_page.data)

    def test_area_worker_runs_discovery_and_report_pipeline(self):
        store = RoofIntelligenceJobStore(self.db_path)
        job = store.create_area_job(
            "39.75", "39.72", "-104.97", "-105.02", "100", ["TPO"]
        )
        store.claim_next_area_job()
        candidates = [
            {
                "candidate_key": f"denver:{parcel}",
                "address": f"{number} Test St, Denver, CO 80202",
                "county": "Denver",
                "county_profile": "denver",
                "parcel": parcel,
                "roof_area_sqft": 12000,
            }
            for parcel, number in (("one", 100), ("two", 200))
        ]

        def generate_report(candidate):
            report_path = Path(self.temp_dir.name) / f"worker-area-{candidate['parcel']}.pdf"
            report_path.write_bytes(b"%PDF-1.4 worker area test")
            return {
                **candidate,
                "report_path": str(report_path),
                "roof_type": "TPO",
                "condition_score": 80,
            }

        with patch.object(self.web, "_roof_worker_readiness_error", return_value=None), \
             patch.object(self.web, "_discover_local_area_candidates", return_value=(candidates, [])), \
             patch.object(self.web, "_run_local_candidate_report", side_effect=generate_report):
            self.web._run_local_area_roof_job(job["id"])

        completed = store.get_job(job["id"])
        self.assertEqual(completed["status"], "completed")
        self.assertEqual(completed["completed_count"], 2)
        self.assertEqual(len(store.get_reports_for_job(job["id"])), 2)

    def test_area_submission_requires_a_valid_drawn_rectangle(self):
        response = self.client.post(
            "/roof-intelligence/jobs/area",
            data={
                "bounds_north": "39.72",
                "bounds_south": "39.75",
                "bounds_east": "-104.97",
                "bounds_west": "-105.02",
                "minimum_roof_squares": "100",
            },
            follow_redirects=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"invalid north/south bounds", response.data)
        self.assertEqual(RoofIntelligenceJobStore(self.db_path).list_jobs(), [])

    def test_local_worker_accepts_supported_non_denver_address_for_county_resolution(self):
        os.environ["ROOF_INTELLIGENCE_LOCAL_WORKER"] = "1"
        with patch.object(self.web, "_roof_worker_readiness_error", return_value=None), \
             patch.object(self.web, "_ensure_roof_worker_started", return_value=True):
            response = self.client.post(
                "/roof-intelligence/jobs/individual",
                data={"property_address": "123 Main St, Aurora, CO 80012"},
                follow_redirects=False,
            )

        self.web.DESKTOP_BACKGROUND_WORK_ACTIVE.clear()
        self.assertEqual(response.status_code, 302)
        jobs = RoofIntelligenceJobStore(self.db_path).list_jobs()
        self.assertEqual(len(jobs), 1)
        self.assertEqual(jobs[0]["input"]["property_address"], "123 Main St, Aurora, CO 80012")


if __name__ == "__main__":
    unittest.main()
