import os
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from urllib.parse import parse_qs, urlsplit

import roof_intelligence_jobs
import roof_intelligence_area_batch as area_batch
import roof_intelligence_single_address as single_address
from roof_intelligence_jobs import RoofIntelligenceJobStore, SUPPORTED_ROOF_TYPES


class RoofIntelligenceJobStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "roof-jobs.sqlite3"
        self.store = RoofIntelligenceJobStore(self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_individual_job_requires_full_address_with_zip(self):
        with self.assertRaisesRegex(ValueError, "five-digit ZIP"):
            self.store.create_individual_job("65 N Yuma St, Denver, CO")

        job = self.store.create_individual_job("65 N Yuma St, Denver, CO 80223")

        self.assertEqual(job["status"], "queued")
        self.assertEqual(job["job_type"], "individual_address")
        self.assertEqual(job["input"]["property_address"], "65 N Yuma St, Denver, CO 80223")

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

    def test_health_history_pruning_retains_recent_rows(self):
        payload = {
            "checked_at": "2020-01-01T00:00:00+00:00",
            "results": [{"county": "Denver", "status": "ok", "error": ""}],
        }
        self.store.record_county_health(payload)

        self.assertEqual(self.store.list_county_health(), [])

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

    def test_microsoft_larger_footprint_is_accepted(self):
        validation = single_address.apply_directional_footprint_rule({
            "status": "discrepancy",
            "primary_sqft": 1200,
            "secondary_sqft": 1000,
            "difference_pct": 16.67,
        }, "secondary_sqft")

        self.assertEqual(validation["status"], "validated")
        self.assertEqual(validation["resolution"], "microsoft_preferred")
        self.assertEqual(validation["county_excess_pct"], 0.0)

    def test_county_up_to_five_percent_larger_is_accepted(self):
        validation = single_address.apply_directional_footprint_rule({
            "status": "validated",
            "primary_sqft": 1000,
            "secondary_sqft": 1050,
            "difference_pct": 5.0,
        }, "secondary_sqft")

        self.assertEqual(validation["status"], "validated")
        self.assertFalse(single_address.should_pause_for_footprint_review(validation, "auto", False))

    def test_county_more_than_five_percent_larger_requires_review(self):
        validation = single_address.apply_directional_footprint_rule({
            "status": "discrepancy",
            "primary_sqft": 1000,
            "secondary_sqft": 1051,
            "difference_pct": 5.1,
        }, "secondary_sqft")

        self.assertEqual(validation["status"], "discrepancy")
        self.assertEqual(validation["county_excess_pct"], 5.1)
        self.assertTrue(single_address.should_pause_for_footprint_review(validation, "auto", False))

    def test_directional_rule_applies_to_assessor_area(self):
        validation = single_address.apply_directional_footprint_rule({
            "status": "discrepancy",
            "primary_sqft": 1500,
            "assessor_sqft": 1200,
            "difference_pct": 20.0,
        }, "assessor_sqft")

        self.assertEqual(validation["status"], "validated")
        self.assertEqual(validation["resolution"], "microsoft_preferred")

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
        DENVER_BUILDINGS_URL = ""
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

        profile, parcel, score, _, source = single_address.resolve_county_and_parcel(
            "123 Main St, Aurora, CO 80012",
            self.Collector,
            profiles,
        )

        self.assertEqual(profile.key, "arapahoe")
        self.assertEqual(parcel["SCHEDNUM"], "A-1")
        self.assertEqual(score, 1.0)
        self.assertIn("Arapahoe", source)

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
