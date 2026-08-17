"""Persistent PCS-side job records for Roof Intelligence processing.

SQLite is the development adapter.  The table and field shapes intentionally
mirror the Supabase contract so the web UI does not depend on the persistence
provider used by the production worker.
"""

from __future__ import annotations

import contextlib
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shutil
import sqlite3
import sys
import uuid

from roof_report_naming import roof_report_pdf_filename


APP_DIR = Path(__file__).resolve().parent
LEGACY_DB_PATH = APP_DIR / "data" / "roof_intelligence_jobs.sqlite3"
CONFIGURED_DATA_DIR = os.environ.get("PCS_DATA_DIR", "").strip()
if CONFIGURED_DATA_DIR:
    DEFAULT_DATA_DIR = Path(CONFIGURED_DATA_DIR).expanduser()
elif sys.platform == "darwin":
    DEFAULT_DATA_DIR = Path.home() / "Library" / "Application Support" / "PCS Proposal Management"
else:
    DEFAULT_DATA_DIR = APP_DIR / "data"
DEFAULT_DB_PATH = DEFAULT_DATA_DIR / "roof_intelligence_jobs.sqlite3"
TERMINAL_STATUSES = {"completed", "completed_with_errors", "failed", "cancelled"}
ACTIVE_STATUSES = {"queued", "running"}
SUPPORTED_ROOF_TYPES = (
    "TPO",
    "PVC",
    "EPDM",
    "Modified Bitumen",
    "Ballasted",
    "Tar and Gravel",
    "Metal",
)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()


def normalize_address(value: object) -> str:
    text = re.sub(r"[^A-Z0-9]+", " ", str(value or "").upper()).strip()
    replacements = {
        "AVENUE": "AVE",
        "BOULEVARD": "BLVD",
        "COURT": "CT",
        "DRIVE": "DR",
        "LANE": "LN",
        "PARKWAY": "PKWY",
        "PLACE": "PL",
        "ROAD": "RD",
        "STREET": "ST",
        "NORTH": "N",
        "SOUTH": "S",
        "EAST": "E",
        "WEST": "W",
    }
    return " ".join(replacements.get(token, token) for token in text.split())


def normalize_county_health_status(value: object) -> str:
    status = str(value or "failed").strip().lower()
    if status == "ok":
        return "healthy"
    if status in {"healthy", "degraded", "failed"}:
        return status
    return "failed"


def validate_full_address(value: object) -> str:
    address = " ".join(str(value or "").split())
    if not address:
        raise ValueError("Enter the full property address.")
    if not re.search(r"\d", address) or not re.search(r"[A-Za-z]", address):
        raise ValueError("Enter a street number and street name.")
    if not re.search(r"\b\d{5}(?:-\d{4})?\b", address):
        raise ValueError("Include the five-digit ZIP code in the property address.")
    return address


def validate_zip(value: object) -> str:
    zip_code = str(value or "").strip()
    if not re.fullmatch(r"\d{5}", zip_code):
        raise ValueError("Enter a five-digit ZIP code.")
    return zip_code


def validate_rectangle_bounds(
    north: object,
    south: object,
    east: object,
    west: object,
) -> dict[str, float]:
    values = {}
    for label, raw_value in (("north", north), ("south", south), ("east", east), ("west", west)):
        try:
            value = float(str(raw_value).strip())
        except (TypeError, ValueError) as exc:
            raise ValueError("Draw a rectangular search area on the map before submitting.") from exc
        if not math.isfinite(value):
            raise ValueError("The selected map area contains an invalid coordinate.")
        values[label] = value
    if not (-90 <= values["south"] < values["north"] <= 90):
        raise ValueError("The selected map area has invalid north/south bounds.")
    if not (-180 <= values["west"] < values["east"] <= 180):
        raise ValueError("The selected map area has invalid east/west bounds.")
    return values


def positive_integer(value: object, label: str, *, minimum: int = 1, maximum: int | None = None) -> int:
    try:
        number = int(str(value).replace(",", "").strip())
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a whole number.") from exc
    if number < minimum or (maximum is not None and number > maximum):
        if maximum is None:
            raise ValueError(f"{label} must be at least {minimum}.")
        raise ValueError(f"{label} must be between {minimum} and {maximum}.")
    return number


def positive_number_rounded_up(value: object, label: str, *, minimum: int = 1) -> int:
    try:
        number = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a number.") from exc
    if not math.isfinite(number):
        raise ValueError(f"{label} must be a number.")
    rounded = math.ceil(number)
    if rounded < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    return rounded


def positive_decimal(value: object, label: str) -> float:
    if value is None or str(value).strip() == "":
        raise ValueError(f"Enter {label.lower()}.")
    try:
        number = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a positive number.") from exc
    if not math.isfinite(number) or number <= 0:
        raise ValueError(f"{label} must be a positive number.")
    return number


def radius_bounds(center_lat: object, center_lng: object, radius_miles: object) -> tuple[dict, dict]:
    try:
        latitude = float(center_lat)
        longitude = float(center_lng)
    except (TypeError, ValueError) as exc:
        raise ValueError("Locate an address on the map before submitting a radius search.") from exc
    if not math.isfinite(latitude) or not math.isfinite(longitude):
        raise ValueError("Locate an address on the map before submitting a radius search.")
    if not (-90 < latitude < 90 and -180 <= longitude <= 180):
        raise ValueError("The radius center is outside the supported map coordinates.")
    radius = positive_decimal(radius_miles, "Radius distance in miles")
    if not 0.1 <= radius <= 10:
        raise ValueError("Radius distance must be between 0.1 and 10 miles.")
    angular_radius = radius / 3958.7613
    latitude_delta = math.degrees(angular_radius)
    cosine_latitude = math.cos(math.radians(latitude))
    longitude_delta = math.degrees(
        math.asin(min(1.0, math.sin(angular_radius) / max(abs(cosine_latitude), 1e-12)))
    )
    bounds = validate_rectangle_bounds(
        latitude + latitude_delta,
        latitude - latitude_delta,
        longitude + longitude_delta,
        longitude - longitude_delta,
    )
    center = {"lat": latitude, "lng": longitude}
    return bounds, {"center": center, "radius_miles": radius}


def optional_nonnegative_integer(value: object, label: str) -> int | None:
    if value is None or str(value).strip() == "":
        return None
    return positive_integer(value, label, minimum=0)


def selected_roof_types(values: list[str] | tuple[str, ...] | None) -> list[str]:
    requested = [str(item).strip() for item in (values or []) if str(item).strip()]
    if not requested or "All" in requested:
        return list(SUPPORTED_ROOF_TYPES)
    invalid = sorted(set(requested) - set(SUPPORTED_ROOF_TYPES))
    if invalid:
        raise ValueError(f"Unsupported roof type: {', '.join(invalid)}")
    return [roof_type for roof_type in SUPPORTED_ROOF_TYPES if roof_type in requested]


class RoofIntelligenceJobStore:
    def __init__(self, db_path: str | os.PathLike[str] | None = None):
        configured_path = db_path or os.environ.get("ROOF_INTELLIGENCE_DB_PATH")
        self.db_path = Path(configured_path or DEFAULT_DB_PATH)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        if configured_path is None:
            self._migrate_legacy_database()
        self.initialize()

    def _migrate_legacy_database(self) -> None:
        """Copy the former in-bundle database to durable user storage once."""
        try:
            if self.db_path.exists() or not LEGACY_DB_PATH.is_file():
                return
            if self.db_path.resolve() == LEGACY_DB_PATH.resolve():
                return
            source = sqlite3.connect(LEGACY_DB_PATH, timeout=30)
            destination = sqlite3.connect(self.db_path, timeout=30)
            try:
                source.backup(destination)
            finally:
                destination.close()
                source.close()
        except (OSError, sqlite3.Error):
            # A migration failure must not prevent the app from starting. The
            # normal initializer below will create a clean durable database.
            pass

    @contextlib.contextmanager
    def connect(self):
        connection = sqlite3.connect(self.db_path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        try:
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def initialize(self) -> None:
        with self.connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS roof_intelligence_jobs (
                    id TEXT PRIMARY KEY,
                    job_type TEXT NOT NULL CHECK (job_type IN ('individual_address', 'zip_batch')),
                    user_key TEXT NOT NULL,
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    input_json TEXT NOT NULL,
                    normalized_address TEXT,
                    zip_code TEXT,
                    report_limit INTEGER,
                    minimum_roof_size INTEGER,
                    minimum_age INTEGER,
                    roof_types_json TEXT NOT NULL DEFAULT '[]',
                    candidate_count INTEGER NOT NULL DEFAULT 0,
                    completed_count INTEGER NOT NULL DEFAULT 0,
                    failed_count INTEGER NOT NULL DEFAULT 0,
                    skipped_count INTEGER NOT NULL DEFAULT 0,
                    remaining_count INTEGER NOT NULL DEFAULT 0,
                    error_code TEXT,
                    error_message TEXT,
                    error_details_json TEXT NOT NULL DEFAULT '{}',
                    retryable INTEGER NOT NULL DEFAULT 0,
                    worker_version TEXT,
                    created_at TEXT NOT NULL,
                    queued_at TEXT NOT NULL,
                    started_at TEXT,
                    finished_at TEXT,
                    updated_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_roof_jobs_user_created
                    ON roof_intelligence_jobs (user_key, created_at DESC);
                CREATE INDEX IF NOT EXISTS idx_roof_jobs_status
                    ON roof_intelligence_jobs (status, queued_at);

                CREATE TABLE IF NOT EXISTS properties (
                    id TEXT PRIMARY KEY,
                    canonical_key TEXT NOT NULL UNIQUE,
                    normalized_address TEXT NOT NULL,
                    address TEXT,
                    city TEXT,
                    state TEXT,
                    zip_code TEXT,
                    county TEXT,
                    parcel_number TEXT,
                    latitude REAL,
                    longitude REAL,
                    roof_area_sqft REAL,
                    roof_squares REAL,
                    year_built INTEGER,
                    effective_year_built INTEGER,
                    age_estimate_year INTEGER,
                    age_estimate_years INTEGER,
                    age_estimate_source TEXT,
                    age_estimate_as_of_date TEXT,
                    data_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS roof_intelligence_reports (
                    id TEXT PRIMARY KEY,
                    property_id TEXT NOT NULL REFERENCES properties(id),
                    job_id TEXT NOT NULL REFERENCES roof_intelligence_jobs(id),
                    report_path TEXT,
                    pdf_size INTEGER,
                    pdf_checksum TEXT,
                    roof_type TEXT,
                    roof_type_confidence REAL,
                    condition_score REAL,
                    risk_level TEXT,
                    imagery_source TEXT,
                    imagery_capture_date TEXT,
                    workflow_version TEXT,
                    result_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_roof_reports_job
                    ON roof_intelligence_reports (job_id, created_at DESC);

                CREATE TABLE IF NOT EXISTS roof_intelligence_report_revisions (
                    id TEXT PRIMARY KEY,
                    report_id TEXT NOT NULL REFERENCES roof_intelligence_reports(id),
                    revision_number INTEGER NOT NULL,
                    parent_revision_id TEXT REFERENCES roof_intelligence_report_revisions(id),
                    revision_kind TEXT NOT NULL,
                    generation_status TEXT NOT NULL,
                    created_by TEXT,
                    revision_note TEXT,
                    edit_patch_json TEXT NOT NULL DEFAULT '{}',
                    snapshot_json TEXT NOT NULL,
                    report_path TEXT,
                    pdf_size INTEGER,
                    pdf_checksum TEXT,
                    generated_at TEXT,
                    created_at TEXT NOT NULL,
                    UNIQUE (report_id, revision_number)
                );
                CREATE INDEX IF NOT EXISTS idx_roof_report_revisions_report
                    ON roof_intelligence_report_revisions (report_id, revision_number DESC);

                CREATE TABLE IF NOT EXISTS roof_intelligence_property_overrides (
                    id TEXT PRIMARY KEY,
                    property_id TEXT NOT NULL REFERENCES properties(id),
                    field_name TEXT NOT NULL CHECK (field_name = 'roof_area_sqft'),
                    numeric_value REAL NOT NULL,
                    source_revision_id TEXT NOT NULL REFERENCES roof_intelligence_report_revisions(id),
                    reason TEXT NOT NULL,
                    created_by TEXT,
                    effective_at TEXT NOT NULL,
                    revoked_at TEXT
                );
                CREATE UNIQUE INDEX IF NOT EXISTS idx_roof_property_override_active
                    ON roof_intelligence_property_overrides (property_id, field_name)
                    WHERE revoked_at IS NULL;

                CREATE TABLE IF NOT EXISTS roof_intelligence_processing_feedback (
                    id TEXT PRIMARY KEY,
                    report_id TEXT NOT NULL REFERENCES roof_intelligence_reports(id),
                    property_id TEXT NOT NULL REFERENCES properties(id),
                    parent_revision_id TEXT NOT NULL REFERENCES roof_intelligence_report_revisions(id),
                    completed_revision_id TEXT NOT NULL REFERENCES roof_intelligence_report_revisions(id),
                    revision_number INTEGER NOT NULL,
                    requested_by TEXT NOT NULL,
                    comment TEXT NOT NULL,
                    status TEXT NOT NULL CHECK (status IN ('pending_review', 'approved', 'rejected', 'applied')),
                    feedback_json TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_roof_processing_feedback_status
                    ON roof_intelligence_processing_feedback (status, created_at);
                CREATE INDEX IF NOT EXISTS idx_roof_processing_feedback_report
                    ON roof_intelligence_processing_feedback (report_id, created_at DESC);

                CREATE TABLE IF NOT EXISTS roof_intelligence_job_items (
                    id TEXT PRIMARY KEY,
                    job_id TEXT NOT NULL REFERENCES roof_intelligence_jobs(id),
                    property_id TEXT REFERENCES properties(id),
                    candidate_key TEXT,
                    input_json TEXT NOT NULL DEFAULT '{}',
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    reason_code TEXT,
                    message TEXT,
                    error_details_json TEXT NOT NULL DEFAULT '{}',
                    report_id TEXT REFERENCES roof_intelligence_reports(id),
                    created_at TEXT NOT NULL,
                    started_at TEXT,
                    finished_at TEXT
                );

                CREATE TABLE IF NOT EXISTS notifications (
                    id TEXT PRIMARY KEY,
                    user_key TEXT NOT NULL,
                    job_id TEXT REFERENCES roof_intelligence_jobs(id),
                    report_id TEXT REFERENCES roof_intelligence_reports(id),
                    kind TEXT NOT NULL,
                    title TEXT NOT NULL,
                    message TEXT NOT NULL,
                    is_read INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    read_at TEXT
                );

                CREATE UNIQUE INDEX IF NOT EXISTS idx_roof_notification_job_kind
                    ON notifications (job_id, kind);
                CREATE INDEX IF NOT EXISTS idx_roof_notifications_user_read
                    ON notifications (user_key, is_read, created_at DESC);

                CREATE TABLE IF NOT EXISTS footprint_resolutions (
                    id TEXT PRIMARY KEY,
                    job_id TEXT NOT NULL REFERENCES roof_intelligence_jobs(id),
                    item_id TEXT REFERENCES roof_intelligence_job_items(id),
                    user_key TEXT NOT NULL,
                    county TEXT,
                    parcel_number TEXT,
                    selected_source TEXT NOT NULL CHECK (selected_source IN ('supabase', 'county')),
                    reason TEXT NOT NULL,
                    validation_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS county_health_checks (
                    id TEXT PRIMARY KEY,
                    county_key TEXT NOT NULL,
                    status TEXT NOT NULL,
                    result_json TEXT NOT NULL,
                    checked_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_county_health_latest
                    ON county_health_checks (county_key, checked_at DESC);
                """
            )
            job_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(roof_intelligence_jobs)").fetchall()
            }
            if "error_details_json" not in job_columns:
                connection.execute(
                    "ALTER TABLE roof_intelligence_jobs ADD COLUMN error_details_json TEXT NOT NULL DEFAULT '{}'"
                )
            item_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(roof_intelligence_job_items)").fetchall()
            }
            if "input_json" not in item_columns:
                connection.execute(
                    "ALTER TABLE roof_intelligence_job_items ADD COLUMN input_json TEXT NOT NULL DEFAULT '{}'"
                )
            if "error_details_json" not in item_columns:
                connection.execute(
                    "ALTER TABLE roof_intelligence_job_items ADD COLUMN error_details_json TEXT NOT NULL DEFAULT '{}'"
                )
            report_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(roof_intelligence_reports)").fetchall()
            }
            if "last_revision_at" not in report_columns:
                connection.execute(
                    "ALTER TABLE roof_intelligence_reports ADD COLUMN last_revision_at TEXT"
                )
            if "retention_expires_at" not in report_columns:
                connection.execute(
                    "ALTER TABLE roof_intelligence_reports ADD COLUMN retention_expires_at TEXT"
                )
            override_columns = {
                row["name"]
                for row in connection.execute(
                    "PRAGMA table_info(roof_intelligence_property_overrides)"
                ).fetchall()
            }
            if "user_key" not in override_columns:
                connection.execute(
                    "ALTER TABLE roof_intelligence_property_overrides "
                    "ADD COLUMN user_key TEXT NOT NULL DEFAULT 'local-user'"
                )
            connection.execute("DROP INDEX IF EXISTS idx_roof_property_override_active")
            connection.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_roof_property_override_active "
                "ON roof_intelligence_property_overrides (user_key, property_id, field_name) "
                "WHERE revoked_at IS NULL"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_roof_job_items_job_status "
                "ON roof_intelligence_job_items (job_id, status, created_at)"
            )

    @staticmethod
    def _job_from_row(row: sqlite3.Row | None) -> dict | None:
        if row is None:
            return None
        result = dict(row)
        result["input"] = json.loads(result.pop("input_json") or "{}")
        result["roof_types"] = json.loads(result.pop("roof_types_json") or "[]")
        result["error_details"] = json.loads(result.pop("error_details_json") or "{}")
        result["retryable"] = bool(result["retryable"])
        return result

    @staticmethod
    def _notification_from_row(row: sqlite3.Row) -> dict:
        result = dict(row)
        result["is_read"] = bool(result["is_read"])
        return result

    def create_individual_job(self, address: object, user_key: str = "local-user") -> dict:
        full_address = validate_full_address(address)
        now = utc_now()
        job_id = str(uuid.uuid4())
        payload = {"property_address": full_address}
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO roof_intelligence_jobs (
                    id, job_type, user_key, status, stage, input_json,
                    normalized_address, roof_types_json, created_at, queued_at, updated_at
                ) VALUES (?, 'individual_address', ?, 'queued', 'queued', ?, ?, '[]', ?, ?, ?)
                """,
                (job_id, user_key, json.dumps(payload), normalize_address(full_address), now, now, now),
            )
        return self.get_job(job_id)

    def create_zip_job(
        self,
        zip_code: object,
        report_limit: object,
        minimum_roof_size: object = 10_000,
        minimum_age: object = None,
        roof_types: list[str] | tuple[str, ...] | None = None,
        user_key: str = "local-user",
    ) -> dict:
        clean_zip = validate_zip(zip_code)
        clean_limit = positive_integer(report_limit, "Report limit", minimum=1, maximum=1_000)
        clean_size = positive_integer(minimum_roof_size, "Minimum roof size", minimum=1)
        clean_age = optional_nonnegative_integer(minimum_age, "Minimum Building/Roof Age Estimate")
        clean_types = selected_roof_types(roof_types)
        payload = {
            "zip_code": clean_zip,
            "report_limit": clean_limit,
            "minimum_roof_size": clean_size,
            "minimum_age": clean_age,
            "roof_types": clean_types,
        }
        now = utc_now()
        job_id = str(uuid.uuid4())
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO roof_intelligence_jobs (
                    id, job_type, user_key, status, stage, input_json, zip_code,
                    report_limit, minimum_roof_size, minimum_age, roof_types_json,
                    created_at, queued_at, updated_at
                ) VALUES (?, 'zip_batch', ?, 'queued', 'queued', ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    job_id,
                    user_key,
                    json.dumps(payload),
                    clean_zip,
                    clean_limit,
                    clean_size,
                    clean_age,
                    json.dumps(clean_types),
                    now,
                    now,
                    now,
                ),
            )
        return self.get_job(job_id)

    def create_area_job(
        self,
        north: object,
        south: object,
        east: object,
        west: object,
        minimum_roof_squares: object = 100,
        roof_types: list[str] | tuple[str, ...] | None = None,
        user_key: str = "local-user",
        selection_type: object = "rectangle",
        center_lat: object = None,
        center_lng: object = None,
        center_address: object = None,
        radius_miles: object = None,
    ) -> dict:
        clean_selection = str(selection_type or "rectangle").strip().lower()
        if clean_selection == "rectangle":
            bounds = validate_rectangle_bounds(north, south, east, west)
            selection_details = {}
        elif clean_selection == "radius":
            bounds, selection_details = radius_bounds(center_lat, center_lng, radius_miles)
            clean_center_address = " ".join(str(center_address or "").split())
            if not clean_center_address:
                raise ValueError("Locate an address on the map before submitting a radius search.")
            selection_details["center_address"] = clean_center_address
        else:
            raise ValueError("Choose either Rectangle or Address Radius for the search area.")
        clean_squares = positive_number_rounded_up(
            minimum_roof_squares, "Minimum roof size in squares", minimum=1
        )
        clean_size = clean_squares * 100
        clean_types = selected_roof_types(roof_types)
        payload = {
            "selection_type": clean_selection,
            "bounds": bounds,
            **selection_details,
            "minimum_roof_squares": clean_squares,
            "roof_types": clean_types,
        }
        now = utc_now()
        job_id = str(uuid.uuid4())
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO roof_intelligence_jobs (
                    id, job_type, user_key, status, stage, input_json,
                    minimum_roof_size, roof_types_json,
                    created_at, queued_at, updated_at
                ) VALUES (?, 'zip_batch', ?, 'queued', 'queued', ?, ?, ?, ?, ?, ?)
                """,
                (
                    job_id,
                    user_key,
                    json.dumps(payload),
                    clean_size,
                    json.dumps(clean_types),
                    now,
                    now,
                    now,
                ),
            )
        return self.get_job(job_id)

    def get_job(self, job_id: str, user_key: str | None = None) -> dict | None:
        sql = "SELECT * FROM roof_intelligence_jobs WHERE id = ?"
        params: tuple[object, ...] = (job_id,)
        if user_key is not None:
            sql += " AND user_key = ?"
            params += (user_key,)
        with self.connect() as connection:
            return self._job_from_row(connection.execute(sql, params).fetchone())

    def list_jobs(self, user_key: str = "local-user", limit: int = 12) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute(
                "SELECT * FROM roof_intelligence_jobs WHERE user_key = ? ORDER BY created_at DESC LIMIT ?",
                (user_key, limit),
            ).fetchall()
        return [self._job_from_row(row) for row in rows]

    def has_queued_individual_jobs(self) -> bool:
        with self.connect() as connection:
            row = connection.execute(
                """
                SELECT 1 FROM roof_intelligence_jobs
                WHERE job_type = 'individual_address' AND status = 'queued'
                LIMIT 1
                """
            ).fetchone()
        return row is not None

    def recover_interrupted_individual_jobs(self, worker_version: str = "pcs-local-adapter") -> int:
        """Return locally claimed jobs to the queue after an interrupted app run."""
        now = utc_now()
        with self.connect() as connection:
            cursor = connection.execute(
                """
                UPDATE roof_intelligence_jobs SET
                    status = 'queued', stage = 'queued', started_at = NULL,
                    error_code = NULL, error_message = NULL, retryable = 1,
                    updated_at = ?
                WHERE job_type = 'individual_address' AND status = 'running'
                  AND worker_version = ?
                """,
                (now, worker_version),
            )
            return cursor.rowcount

    def claim_next_individual_job(self, worker_version: str = "pcs-local-adapter") -> dict | None:
        """Atomically claim the oldest individual request across app processes."""
        now = utc_now()
        claimed_id = None
        with self.connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                """
                SELECT id FROM roof_intelligence_jobs
                WHERE job_type = 'individual_address' AND status = 'queued'
                ORDER BY queued_at ASC, created_at ASC
                LIMIT 1
                """
            ).fetchone()
            if row is not None:
                claimed_id = row["id"]
                cursor = connection.execute(
                    """
                    UPDATE roof_intelligence_jobs SET
                        status = 'running', stage = 'locating_property',
                        started_at = ?, worker_version = ?, error_code = NULL,
                        error_message = NULL, retryable = 0, updated_at = ?
                    WHERE id = ? AND status = 'queued'
                    """,
                    (now, worker_version, now, claimed_id),
                )
                if cursor.rowcount != 1:
                    claimed_id = None
        return self.get_job(claimed_id) if claimed_id else None

    def has_queued_area_jobs(self) -> bool:
        with self.connect() as connection:
            row = connection.execute(
                """
                SELECT 1 FROM roof_intelligence_jobs
                WHERE job_type = 'zip_batch' AND status = 'queued'
                  AND json_extract(input_json, '$.selection_type') IN ('rectangle', 'radius')
                LIMIT 1
                """
            ).fetchone()
        return row is not None

    def recover_interrupted_area_jobs(self, worker_version: str = "pcs-local-adapter") -> int:
        now = utc_now()
        with self.connect() as connection:
            items = connection.execute(
                """
                UPDATE roof_intelligence_job_items SET
                    status = 'pending', stage = 'queued', started_at = NULL
                WHERE status = 'running' AND job_id IN (
                    SELECT id FROM roof_intelligence_jobs
                    WHERE job_type = 'zip_batch' AND status = 'running' AND worker_version = ?
                )
                """,
                (worker_version,),
            )
            connection.execute(
                """
                UPDATE roof_intelligence_jobs SET
                    status = 'queued', stage = 'queued', started_at = NULL,
                    error_code = NULL, error_message = NULL, retryable = 1,
                    updated_at = ?
                WHERE job_type = 'zip_batch' AND status = 'running' AND worker_version = ?
                """,
                (now, worker_version),
            )
            return items.rowcount

    def claim_next_area_job(self, worker_version: str = "pcs-local-adapter") -> dict | None:
        now = utc_now()
        claimed_id = None
        with self.connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                """
                SELECT id FROM roof_intelligence_jobs
                WHERE job_type = 'zip_batch' AND status = 'queued'
                  AND json_extract(input_json, '$.selection_type') IN ('rectangle', 'radius')
                ORDER BY queued_at ASC, created_at ASC
                LIMIT 1
                """
            ).fetchone()
            if row is not None:
                claimed_id = row["id"]
                cursor = connection.execute(
                    """
                    UPDATE roof_intelligence_jobs SET
                        status = 'running', stage = 'discovering_properties',
                        started_at = COALESCE(started_at, ?), worker_version = ?,
                        error_code = NULL, error_message = NULL, retryable = 0, updated_at = ?
                    WHERE id = ? AND status = 'queued'
                    """,
                    (now, worker_version, now, claimed_id),
                )
                if cursor.rowcount != 1:
                    claimed_id = None
        return self.get_job(claimed_id) if claimed_id else None

    def update_job(self, job_id: str, **fields: object) -> dict:
        allowed = {
            "status",
            "stage",
            "candidate_count",
            "completed_count",
            "failed_count",
            "skipped_count",
            "remaining_count",
            "error_code",
            "error_message",
            "error_details_json",
            "retryable",
            "worker_version",
            "started_at",
            "finished_at",
        }
        updates = {key: value for key, value in fields.items() if key in allowed}
        if not updates:
            job = self.get_job(job_id)
            if job is None:
                raise KeyError(job_id)
            return job
        updates["updated_at"] = utc_now()
        if "retryable" in updates:
            updates["retryable"] = int(bool(updates["retryable"]))
        if "error_details_json" in updates and not isinstance(updates["error_details_json"], str):
            updates["error_details_json"] = json.dumps(updates["error_details_json"], default=str)
        assignments = ", ".join(f"{key} = ?" for key in updates)
        with self.connect() as connection:
            cursor = connection.execute(
                f"UPDATE roof_intelligence_jobs SET {assignments} WHERE id = ?",
                (*updates.values(), job_id),
            )
            if cursor.rowcount != 1:
                raise KeyError(job_id)
        return self.get_job(job_id)

    @staticmethod
    def _area_item_from_row(row: sqlite3.Row | None) -> dict | None:
        if row is None:
            return None
        result = dict(row)
        result["input"] = json.loads(result.pop("input_json") or "{}")
        result["error_details"] = json.loads(result.pop("error_details_json") or "{}")
        return result

    @staticmethod
    def _refresh_area_counts(connection: sqlite3.Connection, job_id: str) -> None:
        counts = connection.execute(
            """
            SELECT COUNT(*) AS candidate_count,
                   SUM(CASE WHEN status = 'completed' THEN 1 ELSE 0 END) AS completed_count,
                   SUM(CASE WHEN status = 'failed' THEN 1 ELSE 0 END) AS failed_count,
                   SUM(CASE WHEN status = 'skipped' THEN 1 ELSE 0 END) AS skipped_count,
                   SUM(CASE WHEN status IN ('pending', 'running') THEN 1 ELSE 0 END) AS remaining_count
            FROM roof_intelligence_job_items WHERE job_id = ?
            """,
            (job_id,),
        ).fetchone()
        connection.execute(
            """
            UPDATE roof_intelligence_jobs SET
                candidate_count = ?, completed_count = ?, failed_count = ?,
                skipped_count = ?, remaining_count = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                counts["candidate_count"] or 0,
                counts["completed_count"] or 0,
                counts["failed_count"] or 0,
                counts["skipped_count"] or 0,
                counts["remaining_count"] or 0,
                utc_now(),
                job_id,
            ),
        )

    def prepare_area_candidates(self, job_id: str, candidates: list[dict]) -> list[dict]:
        now = utc_now()
        with self.connect() as connection:
            existing_keys = {
                row["candidate_key"]
                for row in connection.execute(
                    "SELECT candidate_key FROM roof_intelligence_job_items WHERE job_id = ?",
                    (job_id,),
                ).fetchall()
            }
            for candidate in candidates:
                candidate_key = str(candidate.get("candidate_key") or "").strip()
                if not candidate_key or candidate_key in existing_keys:
                    continue
                connection.execute(
                    """
                    INSERT INTO roof_intelligence_job_items (
                        id, job_id, candidate_key, input_json, status, stage, created_at
                    ) VALUES (?, ?, ?, ?, 'pending', 'queued', ?)
                    """,
                    (str(uuid.uuid4()), job_id, candidate_key, json.dumps(candidate, default=str), now),
                )
                existing_keys.add(candidate_key)
            self._refresh_area_counts(connection, job_id)
        return self.list_area_items(job_id)

    def list_area_items(self, job_id: str) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute(
                "SELECT * FROM roof_intelligence_job_items WHERE job_id = ? ORDER BY created_at, id",
                (job_id,),
            ).fetchall()
        return [self._area_item_from_row(row) for row in rows]

    def claim_next_area_item(self, job_id: str) -> dict | None:
        now = utc_now()
        claimed_id = None
        with self.connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            job = connection.execute(
                "SELECT status FROM roof_intelligence_jobs WHERE id = ?",
                (job_id,),
            ).fetchone()
            if job is None or job["status"] != "running":
                return None
            row = connection.execute(
                """
                SELECT id FROM roof_intelligence_job_items
                WHERE job_id = ? AND status = 'pending'
                ORDER BY created_at, id LIMIT 1
                """,
                (job_id,),
            ).fetchone()
            if row:
                claimed_id = row["id"]
                connection.execute(
                    """
                    UPDATE roof_intelligence_job_items SET
                        status = 'running', stage = 'processing_report', started_at = ?
                    WHERE id = ? AND status = 'pending'
                    """,
                    (now, claimed_id),
                )
                connection.execute(
                    "UPDATE roof_intelligence_jobs SET stage = 'processing_reports', updated_at = ? WHERE id = ?",
                    (now, job_id),
                )
                self._refresh_area_counts(connection, job_id)
        if not claimed_id:
            return None
        with self.connect() as connection:
            row = connection.execute(
                "SELECT * FROM roof_intelligence_job_items WHERE id = ?", (claimed_id,)
            ).fetchone()
        return self._area_item_from_row(row)

    def fail_area_item(
        self, job_id: str, item_id: str, code: str, message: str, error_details: dict | None = None
    ) -> None:
        now = utc_now()
        clean_message = " ".join(str(message or "Unable to process candidate.").split())[:500]
        with self.connect() as connection:
            connection.execute(
                """
                UPDATE roof_intelligence_job_items SET
                    status = 'failed', stage = 'failed', reason_code = ?, message = ?,
                    error_details_json = ?, finished_at = ?
                WHERE id = ? AND job_id = ?
                """,
                (code, clean_message, json.dumps(error_details or {}, default=str), now, item_id, job_id),
            )
            self._refresh_area_counts(connection, job_id)
        if code == "footprint_discrepancy":
            job = self.get_job(job_id)
            if job:
                self._create_notification(
                    job["user_key"], job_id, None,
                    f"footprint_discrepancy_{item_id}",
                    "Building footprint review required",
                    clean_message,
                )

    def skip_area_item(self, job_id: str, item_id: str, code: str, message: str) -> None:
        now = utc_now()
        with self.connect() as connection:
            connection.execute(
                """
                UPDATE roof_intelligence_job_items SET
                    status = 'skipped', stage = 'skipped', reason_code = ?, message = ?, finished_at = ?
                WHERE id = ? AND job_id = ?
                """,
                (code, " ".join(str(message).split())[:500], now, item_id, job_id),
            )
            self._refresh_area_counts(connection, job_id)

    def skip_pending_area_items(self, job_id: str, code: str, message: str) -> int:
        now = utc_now()
        with self.connect() as connection:
            cursor = connection.execute(
                """
                UPDATE roof_intelligence_job_items SET
                    status = 'skipped', stage = 'skipped', reason_code = ?, message = ?, finished_at = ?
                WHERE job_id = ? AND status = 'pending'
                """,
                (code, " ".join(str(message).split())[:500], now, job_id),
            )
            self._refresh_area_counts(connection, job_id)
            return cursor.rowcount

    def start_job(self, job_id: str, stage: str = "locating_property", worker_version: str = "pcs-local-adapter") -> dict:
        return self.update_job(
            job_id,
            status="running",
            stage=stage,
            started_at=utc_now(),
            worker_version=worker_version,
            error_code=None,
            error_message=None,
            error_details_json={},
            retryable=False,
        )

    def fail_job(
        self,
        job_id: str,
        error_code: str,
        error_message: str,
        *,
        retryable: bool = False,
        stage: str = "failed",
        error_details: dict | None = None,
    ) -> dict:
        message = " ".join(str(error_message or "Unable to complete the Roof Intelligence job.").split())[:500]
        job = self.update_job(
            job_id,
            status="failed",
            stage=stage,
            error_code=error_code,
            error_message=message,
            error_details_json=error_details or {},
            retryable=retryable,
            finished_at=utc_now(),
        )
        self._create_notification(
            job["user_key"],
            job_id,
            None,
            "job_failed",
            "Roof Intelligence job failed",
            message,
        )
        return self.get_job(job_id)

    @staticmethod
    def _canonical_property_key(result: dict, fallback_address: str) -> str:
        county = normalize_address(result.get("county"))
        parcel = normalize_address(result.get("parcel") or result.get("parcel_number"))
        if county and parcel:
            return f"{county}:{parcel}"
        return f"ADDRESS:{normalize_address(result.get('address') or fallback_address)}"

    @staticmethod
    def _retention_expiry(timestamp: str) -> str:
        value = dt.datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
        return (value + dt.timedelta(days=90)).isoformat()

    def _persist_initial_revision_assets(self, result: dict) -> None:
        """Preserve Revision 1 inputs only when a versioned snapshot exists."""
        snapshot = result.get("report_snapshot")
        if not isinstance(snapshot, dict):
            return
        report_id = str(result.get("report_id") or snapshot.get("report_id") or "").strip()
        if not report_id or str(snapshot.get("report_id") or "") != report_id:
            raise ValueError("The initial report snapshot does not match its report ID.")

        asset_dir = self.db_path.parent / "roof_intelligence_reports" / report_id
        asset_dir.mkdir(parents=True, exist_ok=True)

        def preserve(source_value: object, destination: Path) -> str:
            source = Path(str(source_value or "")).expanduser()
            if not source.is_file():
                return ""
            if source.resolve() != destination.resolve():
                temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex}.copying")
                try:
                    shutil.copy2(source, temporary)
                    os.replace(temporary, destination)
                finally:
                    if temporary.exists():
                        temporary.unlink()
            return str(destination)

        durable_pdf = preserve(
            result.get("report_path"),
            asset_dir
            / roof_report_pdf_filename(result.get("address"), result.get("city")),
        )
        if durable_pdf:
            result["report_path"] = durable_pdf

        imagery = snapshot.get("imagery")
        if isinstance(imagery, dict):
            source_image = result.get("aerial_image_file") or imagery.get("local_report_image_path")
            suffix = Path(str(source_image or "")).suffix.lower()
            if suffix not in {".jpg", ".jpeg", ".png", ".webp"}:
                suffix = ".jpg"
            durable_image = preserve(source_image, asset_dir / f"revision-1-image{suffix}")
            if durable_image:
                imagery["local_report_image_path"] = durable_image
            analysis_image = (
                result.get("analysis_aerial_image_file")
                or imagery.get("local_analysis_image_path")
            )
            analysis_suffix = Path(str(analysis_image or "")).suffix.lower()
            if analysis_suffix not in {".jpg", ".jpeg", ".png", ".webp"}:
                analysis_suffix = ".png"
            durable_analysis_image = preserve(
                analysis_image,
                asset_dir / f"revision-1-analysis-mask{analysis_suffix}",
            )
            if durable_analysis_image:
                imagery["local_analysis_image_path"] = durable_analysis_image

    @classmethod
    def _insert_initial_revision(
        cls,
        connection: sqlite3.Connection,
        report_id: str,
        result: dict,
        report_path: str,
        pdf_size: int | None,
        pdf_checksum: str | None,
        created_at: str,
    ) -> None:
        snapshot = result.get("report_snapshot")
        if not isinstance(snapshot, dict):
            return
        if str(snapshot.get("report_id") or "") != report_id:
            raise ValueError("The initial report snapshot does not match its report ID.")
        revision = snapshot.get("revision") or {}
        if revision.get("number") != 1 or revision.get("kind") != "initial":
            raise ValueError("The initial report snapshot has invalid revision metadata.")
        snapshot_id = str(snapshot.get("snapshot_id") or "").strip()
        if not snapshot_id:
            raise ValueError("The initial report snapshot is missing its snapshot ID.")
        generated_at = str(revision.get("created_at") or created_at)
        connection.execute(
            """
            INSERT OR IGNORE INTO roof_intelligence_report_revisions (
                id, report_id, revision_number, parent_revision_id, revision_kind,
                generation_status, created_by, revision_note, edit_patch_json,
                snapshot_json, report_path, pdf_size, pdf_checksum, generated_at, created_at
            ) VALUES (?, ?, 1, NULL, 'initial', 'ready', ?, ?, '{}', ?, ?, ?, ?, ?, ?)
            """,
            (
                snapshot_id,
                report_id,
                revision.get("created_by"),
                revision.get("change_reason") or "Fresh assessment",
                json.dumps(snapshot, default=str),
                report_path,
                pdf_size,
                pdf_checksum,
                generated_at,
                created_at,
            ),
        )
        connection.execute(
            """
            UPDATE roof_intelligence_reports
            SET last_revision_at = ?, retention_expires_at = ?
            WHERE id = ?
            """,
            (generated_at, cls._retention_expiry(generated_at), report_id),
        )

    def complete_individual_job(self, job_id: str, result: dict) -> dict:
        job = self.get_job(job_id)
        if job is None:
            raise KeyError(job_id)
        fallback_address = job["input"].get("property_address", "")
        self._persist_initial_revision_assets(result)
        canonical_key = self._canonical_property_key(result, fallback_address)
        now = utc_now()
        report_path = str(result.get("report_path") or "")
        pdf_size = None
        pdf_checksum = None
        if report_path and Path(report_path).is_file():
            pdf_size = Path(report_path).stat().st_size
            digest = hashlib.sha256()
            with Path(report_path).open("rb") as handle:
                for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                    digest.update(chunk)
            pdf_checksum = digest.hexdigest()

        with self.connect() as connection:
            existing = connection.execute(
                "SELECT id FROM properties WHERE canonical_key = ?", (canonical_key,)
            ).fetchone()
            property_id = existing["id"] if existing else str(uuid.uuid4())
            address = str(result.get("address") or fallback_address)
            property_values = (
                normalize_address(address),
                address,
                result.get("city"),
                result.get("state"),
                result.get("zip") or result.get("zip_code"),
                result.get("county"),
                result.get("parcel") or result.get("parcel_number"),
                result.get("latitude"),
                result.get("longitude"),
                result.get("roof_area_sqft") or result.get("building_footprint_sqft"),
                result.get("roof_squares"),
                result.get("year_built"),
                result.get("effective_year_built"),
                result.get("age_estimate_year"),
                result.get("age_estimate_years"),
                result.get("age_estimate_source"),
                result.get("age_estimate_as_of_date"),
                json.dumps(result, default=str),
                now,
            )
            if existing:
                connection.execute(
                    """
                    UPDATE properties SET
                        normalized_address = ?, address = ?, city = ?, state = ?, zip_code = ?,
                        county = ?, parcel_number = ?, latitude = ?, longitude = ?, roof_area_sqft = ?,
                        roof_squares = ?, year_built = ?, effective_year_built = ?, age_estimate_year = ?,
                        age_estimate_years = ?, age_estimate_source = ?, age_estimate_as_of_date = ?,
                        data_json = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (*property_values, property_id),
                )
            else:
                connection.execute(
                    """
                    INSERT INTO properties (
                        id, canonical_key, normalized_address, address, city, state, zip_code,
                        county, parcel_number, latitude, longitude, roof_area_sqft, roof_squares,
                        year_built, effective_year_built, age_estimate_year, age_estimate_years,
                        age_estimate_source, age_estimate_as_of_date, data_json, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (property_id, canonical_key, *property_values[:-1], now, now),
                )

            existing_report = connection.execute(
                "SELECT id FROM roof_intelligence_reports WHERE job_id = ?", (job_id,)
            ).fetchone()
            report_id = (
                existing_report["id"]
                if existing_report
                else str(result.get("report_id") or uuid.uuid4())
            )
            if existing_report:
                connection.execute(
                    """
                    UPDATE roof_intelligence_reports SET
                        property_id = ?, report_path = ?, pdf_size = ?, pdf_checksum = ?,
                        roof_type = ?, roof_type_confidence = ?, condition_score = ?, risk_level = ?,
                        imagery_source = ?, imagery_capture_date = ?, workflow_version = ?, result_json = ?
                    WHERE id = ?
                    """,
                    (
                        property_id, report_path, pdf_size, pdf_checksum,
                        result.get("roof_type"), result.get("roof_type_confidence"),
                        result.get("condition_score"), result.get("risk_level"),
                        result.get("imagery_source"), result.get("imagery_capture_date"),
                        result.get("workflow_version"), json.dumps(result, default=str), report_id,
                    ),
                )
            else:
                connection.execute(
                    """
                    INSERT INTO roof_intelligence_reports (
                        id, property_id, job_id, report_path, pdf_size, pdf_checksum,
                        roof_type, roof_type_confidence, condition_score, risk_level,
                        imagery_source, imagery_capture_date, workflow_version, result_json, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        report_id, property_id, job_id, report_path, pdf_size, pdf_checksum,
                        result.get("roof_type"), result.get("roof_type_confidence"),
                        result.get("condition_score"), result.get("risk_level"),
                        result.get("imagery_source"), result.get("imagery_capture_date"),
                        result.get("workflow_version"), json.dumps(result, default=str), now,
                    ),
                )
            self._insert_initial_revision(
                connection,
                report_id,
                result,
                report_path,
                pdf_size,
                pdf_checksum,
                now,
            )
            connection.execute(
                """
                UPDATE roof_intelligence_jobs SET
                    status = 'completed', stage = 'completed', candidate_count = 1,
                    completed_count = 1, remaining_count = 0, error_code = NULL,
                    error_message = NULL, retryable = 0, finished_at = ?, updated_at = ?
                WHERE id = ?
                """,
                (now, now, job_id),
            )

        self._create_notification(
            job["user_key"],
            job_id,
            report_id,
            "job_completed",
            "Roof Intelligence report ready",
            f"The report for {result.get('address') or fallback_address} is ready.",
        )
        assessor_warnings = [str(item).strip() for item in result.get("assessor_warnings") or [] if str(item).strip()]
        if assessor_warnings:
            self._create_notification(
                job["user_key"],
                job_id,
                report_id,
                "assessor_warning",
                "Property data notice",
                " | ".join(assessor_warnings),
            )
        return self.get_job(job_id)

    def complete_area_item(self, job_id: str, item_id: str, result: dict) -> dict:
        job = self.get_job(job_id)
        if job is None:
            raise KeyError(job_id)
        item = next((entry for entry in self.list_area_items(job_id) if entry["id"] == item_id), None)
        if item is None:
            raise KeyError(item_id)
        fallback_address = item["input"].get("address", "")
        self._persist_initial_revision_assets(result)
        canonical_key = self._canonical_property_key(result, fallback_address)
        now = utc_now()
        report_path = str(result.get("report_path") or "")
        pdf_size = None
        pdf_checksum = None
        if report_path and Path(report_path).is_file():
            pdf_size = Path(report_path).stat().st_size
            digest = hashlib.sha256()
            with Path(report_path).open("rb") as handle:
                for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                    digest.update(chunk)
            pdf_checksum = digest.hexdigest()

        with self.connect() as connection:
            existing = connection.execute(
                "SELECT id FROM properties WHERE canonical_key = ?", (canonical_key,)
            ).fetchone()
            property_id = existing["id"] if existing else str(uuid.uuid4())
            address = str(result.get("address") or fallback_address)
            candidate = item["input"]
            property_values = (
                normalize_address(address),
                address,
                result.get("city") or candidate.get("city"),
                result.get("state") or candidate.get("state"),
                result.get("zip") or result.get("zip_code") or candidate.get("zip"),
                result.get("county") or candidate.get("county"),
                result.get("parcel") or result.get("parcel_number") or candidate.get("parcel"),
                result.get("latitude"),
                result.get("longitude"),
                result.get("roof_area_sqft") or result.get("building_footprint_sqft") or candidate.get("roof_area_sqft"),
                result.get("roof_squares") or candidate.get("roof_squares"),
                result.get("year_built") or candidate.get("year_built"),
                result.get("effective_year_built") or candidate.get("effective_year_built"),
                result.get("age_estimate_year"),
                result.get("age_estimate_years") or candidate.get("age_estimate_years"),
                result.get("age_estimate_source"),
                result.get("age_estimate_as_of_date"),
                json.dumps(result, default=str),
                now,
            )
            if existing:
                connection.execute(
                    """
                    UPDATE properties SET
                        normalized_address = ?, address = ?, city = ?, state = ?, zip_code = ?,
                        county = ?, parcel_number = ?, latitude = ?, longitude = ?, roof_area_sqft = ?,
                        roof_squares = ?, year_built = ?, effective_year_built = ?, age_estimate_year = ?,
                        age_estimate_years = ?, age_estimate_source = ?, age_estimate_as_of_date = ?,
                        data_json = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (*property_values, property_id),
                )
            else:
                connection.execute(
                    """
                    INSERT INTO properties (
                        id, canonical_key, normalized_address, address, city, state, zip_code,
                        county, parcel_number, latitude, longitude, roof_area_sqft, roof_squares,
                        year_built, effective_year_built, age_estimate_year, age_estimate_years,
                        age_estimate_source, age_estimate_as_of_date, data_json, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (property_id, canonical_key, *property_values[:-1], now, now),
                )

            report_id = str(result.get("report_id") or uuid.uuid4())
            connection.execute(
                """
                INSERT INTO roof_intelligence_reports (
                    id, property_id, job_id, report_path, pdf_size, pdf_checksum,
                    roof_type, roof_type_confidence, condition_score, risk_level,
                    imagery_source, imagery_capture_date, workflow_version, result_json, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    report_id, property_id, job_id, report_path, pdf_size, pdf_checksum,
                    result.get("roof_type"), result.get("roof_type_confidence"),
                    result.get("condition_score"), result.get("risk_level"),
                    result.get("imagery_source"), result.get("imagery_capture_date"),
                    result.get("workflow_version"), json.dumps(result, default=str), now,
                ),
            )
            self._insert_initial_revision(
                connection,
                report_id,
                result,
                report_path,
                pdf_size,
                pdf_checksum,
                now,
            )
            connection.execute(
                """
                UPDATE roof_intelligence_job_items SET
                    property_id = ?, status = 'completed', stage = 'completed',
                    report_id = ?, reason_code = NULL, message = NULL, finished_at = ?
                WHERE id = ? AND job_id = ?
                """,
                (property_id, report_id, now, item_id, job_id),
            )
            self._refresh_area_counts(connection, job_id)
        assessor_warnings = [str(value).strip() for value in result.get("assessor_warnings") or [] if str(value).strip()]
        if assessor_warnings:
            self._create_notification(
                job["user_key"],
                job_id,
                report_id,
                "assessor_warning",
                "Property data notice",
                f"{address}: " + " | ".join(assessor_warnings),
            )
        return self.get_job(job_id)

    def finish_area_job(self, job_id: str) -> dict:
        job = self.get_job(job_id)
        if job is None:
            raise KeyError(job_id)
        if job["status"] == "cancelled":
            return job
        now = utc_now()
        status = "completed_with_errors" if job["failed_count"] else "completed"
        with self.connect() as connection:
            self._refresh_area_counts(connection, job_id)
            refreshed = connection.execute(
                "SELECT completed_count, failed_count FROM roof_intelligence_jobs WHERE id = ?", (job_id,)
            ).fetchone()
            status = "completed_with_errors" if refreshed["failed_count"] else "completed"
            connection.execute(
                """
                UPDATE roof_intelligence_jobs SET
                    status = ?, stage = 'completed', remaining_count = 0,
                    finished_at = ?, retryable = 0, updated_at = ? WHERE id = ?
                """,
                (status, now, now, job_id),
            )
        completed = self.get_job(job_id)
        self._create_notification(
            completed["user_key"],
            job_id,
            None,
            "batch_completed",
            "Roof Intelligence area batch finished",
            f"{completed['completed_count']} report(s) completed, "
            f"{completed['failed_count']} failed, and {completed['skipped_count']} skipped.",
        )
        return self.get_job(job_id)

    def cancel_job(self, job_id: str, user_key: str = "local-user") -> dict | None:
        now = utc_now()
        with self.connect() as connection:
            changed = connection.execute(
                """
                UPDATE roof_intelligence_jobs SET
                    status = 'cancelled', stage = 'cancelled', remaining_count = 0,
                    retryable = 0, finished_at = ?, updated_at = ?
                WHERE id = ? AND user_key = ? AND status IN ('queued', 'running')
                """,
                (now, now, job_id, user_key),
            ).rowcount
            if changed:
                connection.execute(
                    """
                    UPDATE roof_intelligence_job_items SET
                        status = 'skipped', stage = 'cancelled', reason_code = 'job_cancelled',
                        message = 'The batch was cancelled.', finished_at = ?
                    WHERE job_id = ? AND status IN ('pending', 'running')
                    """,
                    (now, job_id),
                )
                self._refresh_area_counts(connection, job_id)
                connection.execute(
                    """
                    UPDATE roof_intelligence_jobs SET
                        status = 'cancelled', stage = 'cancelled', remaining_count = 0,
                        finished_at = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (now, now, job_id),
                )
        return self.get_job(job_id, user_key)

    def get_report_for_job(self, job_id: str) -> dict | None:
        with self.connect() as connection:
            row = connection.execute(
                """
                SELECT r.*, p.address, p.city, p.state, p.zip_code, p.county,
                       p.parcel_number, p.roof_area_sqft, p.roof_squares,
                       p.age_estimate_years, p.age_estimate_source
                FROM roof_intelligence_reports r
                JOIN properties p ON p.id = r.property_id
                WHERE r.job_id = ?
                """,
                (job_id,),
            ).fetchone()
        if row is None:
            return None
        result = dict(row)
        result["result"] = json.loads(result.pop("result_json") or "{}")
        return result

    def get_reports_for_job(self, job_id: str) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute(
                """
                SELECT r.*, p.address, p.city, p.state, p.zip_code, p.county,
                       p.parcel_number, p.roof_area_sqft, p.roof_squares,
                       p.age_estimate_years, p.age_estimate_source
                FROM roof_intelligence_reports r
                JOIN properties p ON p.id = r.property_id
                WHERE r.job_id = ? ORDER BY r.created_at, r.id
                """,
                (job_id,),
            ).fetchall()
        reports = []
        for row in rows:
            result = dict(row)
            result["result"] = json.loads(result.pop("result_json") or "{}")
            reports.append(result)
        return reports

    def get_report(self, report_id: str, user_key: str | None = None) -> dict | None:
        with self.connect() as connection:
            if user_key is None:
                row = connection.execute(
                    "SELECT * FROM roof_intelligence_reports WHERE id = ?", (report_id,)
                ).fetchone()
            else:
                row = connection.execute(
                    """
                    SELECT r.* FROM roof_intelligence_reports r
                    JOIN roof_intelligence_jobs j ON j.id = r.job_id
                    WHERE r.id = ? AND j.user_key = ?
                    """,
                    (report_id, user_key),
                ).fetchone()
        if row is None:
            return None
        result = dict(row)
        result["result"] = json.loads(result.pop("result_json") or "{}")
        return result

    @staticmethod
    def _revision_from_row(row: sqlite3.Row | None) -> dict | None:
        if row is None:
            return None
        result = dict(row)
        result["edit_patch"] = json.loads(result.pop("edit_patch_json") or "{}")
        result["snapshot"] = json.loads(result.pop("snapshot_json") or "{}")
        return result

    def list_report_revisions(self, report_id: str, user_key: str | None = None) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute(
                """
                SELECT revision.* FROM roof_intelligence_report_revisions revision
                JOIN roof_intelligence_reports report ON report.id = revision.report_id
                JOIN roof_intelligence_jobs job ON job.id = report.job_id
                WHERE revision.report_id = ? AND (? IS NULL OR job.user_key = ?)
                ORDER BY revision.revision_number DESC
                """,
                (report_id, user_key, user_key),
            ).fetchall()
        return [self._revision_from_row(row) for row in rows]

    def get_report_revision(self, revision_id: str, user_key: str | None = None) -> dict | None:
        with self.connect() as connection:
            row = connection.execute(
                """
                SELECT revision.* FROM roof_intelligence_report_revisions revision
                JOIN roof_intelligence_reports report ON report.id = revision.report_id
                JOIN roof_intelligence_jobs job ON job.id = report.job_id
                WHERE revision.id = ? AND (? IS NULL OR job.user_key = ?)
                """,
                (revision_id, user_key, user_key),
            ).fetchone()
        return self._revision_from_row(row)

    def get_report_review(self, report_id: str, user_key: str | None = None) -> dict | None:
        with self.connect() as connection:
            row = connection.execute(
                """
                SELECT r.*, p.address, p.city, p.state, p.zip_code, p.county,
                       p.parcel_number, p.canonical_key
                FROM roof_intelligence_reports r
                JOIN properties p ON p.id = r.property_id
                JOIN roof_intelligence_jobs j ON j.id = r.job_id
                WHERE r.id = ? AND (? IS NULL OR j.user_key = ?)
                """,
                (report_id, user_key, user_key),
            ).fetchone()
        if row is None:
            return None
        result = dict(row)
        result["result"] = json.loads(result.pop("result_json") or "{}")
        result["revisions"] = self.list_report_revisions(report_id, user_key)
        result["latest_revision"] = result["revisions"][0] if result["revisions"] else None
        result["processing_feedback"] = self.list_processing_feedback(report_id=report_id, user_key=user_key)
        return result

    def list_processing_feedback(
        self,
        *,
        report_id: str | None = None,
        status: str | None = None,
        user_key: str | None = None,
    ) -> list[dict]:
        conditions = []
        parameters: list[object] = []
        if report_id:
            conditions.append("feedback.report_id = ?")
            parameters.append(report_id)
        if status:
            conditions.append("feedback.status = ?")
            parameters.append(status)
        if user_key:
            conditions.append("job.user_key = ?")
            parameters.append(user_key)
        where = " WHERE " + " AND ".join(conditions) if conditions else ""
        with self.connect() as connection:
            rows = connection.execute(
                "SELECT feedback.* FROM roof_intelligence_processing_feedback feedback "
                "JOIN roof_intelligence_reports report ON report.id = feedback.report_id "
                "JOIN roof_intelligence_jobs job ON job.id = report.job_id"
                + where
                + " ORDER BY feedback.created_at DESC, feedback.id",
                parameters,
            ).fetchall()
        results = []
        for row in rows:
            item = dict(row)
            item["feedback"] = json.loads(item.pop("feedback_json") or "{}")
            results.append(item)
        return results

    def get_active_square_footage_override(
        self,
        *,
        address: object = "",
        county: object = "",
        parcel_number: object = "",
        user_key: str | None = None,
    ) -> dict | None:
        county_key = normalize_address(county)
        parcel_key = normalize_address(parcel_number)
        canonical_key = f"{county_key}:{parcel_key}" if county_key and parcel_key else ""
        street_key = normalize_address(str(address or "").split(",", 1)[0])
        full_address_key = normalize_address(address)
        with self.connect() as connection:
            row = connection.execute(
                """
                SELECT o.*, p.canonical_key, p.address
                FROM roof_intelligence_property_overrides o
                JOIN properties p ON p.id = o.property_id
                WHERE o.field_name = 'roof_area_sqft' AND o.revoked_at IS NULL
                  AND (? IS NULL OR o.user_key = ?)
                  AND (
                    (? <> '' AND p.canonical_key = ?)
                    OR (? <> '' AND p.normalized_address = ?)
                    OR (? <> '' AND p.normalized_address = ?)
                  )
                ORDER BY CASE WHEN p.canonical_key = ? THEN 0 ELSE 1 END,
                         o.effective_at DESC
                LIMIT 1
                """,
                (
                    user_key, user_key,
                    canonical_key, canonical_key,
                    street_key, street_key,
                    full_address_key, full_address_key,
                    canonical_key,
                ),
            ).fetchone()
        return dict(row) if row is not None else None

    def save_ready_report_revision(
        self,
        report_id: str,
        parent_revision_id: str,
        snapshot: dict,
        *,
        report_path: str,
        pdf_size: int,
        pdf_checksum: str,
        created_by: str,
        change_reason: str,
        edits: dict,
        apply_square_footage_to_future: bool = False,
        processing_feedback: dict | None = None,
        user_key: str | None = None,
    ) -> dict:
        report = self.get_report(report_id, user_key)
        parent = self.get_report_revision(parent_revision_id, user_key)
        if report is None or parent is None or parent["report_id"] != report_id:
            raise KeyError(report_id)
        if parent["generation_status"] != "ready":
            raise ValueError("Only a ready report revision can be edited.")
        if str(snapshot.get("report_id") or "") != report_id:
            raise ValueError("The revised snapshot does not belong to this report.")
        revision = snapshot.get("revision") or {}
        next_number = max(
            (item["revision_number"] for item in self.list_report_revisions(report_id, user_key)),
            default=0,
        ) + 1
        if revision.get("number") != next_number:
            raise ValueError("The revised snapshot does not have the next revision number.")
        if revision.get("parent_snapshot_id") != parent_revision_id:
            raise ValueError("The revised snapshot parent does not match the selected revision.")
        revision_id = str(snapshot.get("snapshot_id") or "").strip()
        if not revision_id:
            raise ValueError("The revised snapshot is missing its snapshot ID.")
        generated_at = str(revision.get("created_at") or utc_now())
        analysis = snapshot.get("analysis") or {}
        calculations = snapshot.get("calculations") or {}
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO roof_intelligence_report_revisions (
                    id, report_id, revision_number, parent_revision_id, revision_kind,
                    generation_status, created_by, revision_note, edit_patch_json,
                    snapshot_json, report_path, pdf_size, pdf_checksum, generated_at, created_at
                ) VALUES (?, ?, ?, ?, 'manual_edit', 'ready', ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    revision_id,
                    report_id,
                    next_number,
                    parent_revision_id,
                    created_by,
                    change_reason,
                    json.dumps(edits, default=str),
                    json.dumps(snapshot, default=str),
                    report_path,
                    int(pdf_size),
                    pdf_checksum,
                    generated_at,
                    generated_at,
                ),
            )
            current_result = dict(report.get("result") or {})
            current_result.update(
                {
                    "report_path": report_path,
                    "report_snapshot": snapshot,
                    "roof_type": analysis.get("roof_type"),
                    "condition_score": analysis.get("overall_score"),
                    "risk_level": analysis.get("risk_level"),
                    "building_footprint_sqft": calculations.get("roof_area_sqft"),
                    "roof_squares": calculations.get("roof_squares"),
                }
            )
            connection.execute(
                """
                UPDATE roof_intelligence_reports SET
                    report_path = ?, pdf_size = ?, pdf_checksum = ?, roof_type = ?,
                    condition_score = ?, risk_level = ?, result_json = ?,
                    last_revision_at = ?, retention_expires_at = ?
                WHERE id = ?
                """,
                (
                    report_path,
                    int(pdf_size),
                    pdf_checksum,
                    analysis.get("roof_type"),
                    analysis.get("overall_score"),
                    analysis.get("risk_level"),
                    json.dumps(current_result, default=str),
                    generated_at,
                    self._retention_expiry(generated_at),
                    report_id,
                ),
            )
            if apply_square_footage_to_future:
                if "roof_area_sqft" not in edits:
                    raise ValueError("A future override requires a square-footage edit.")
                connection.execute(
                    """
                    UPDATE roof_intelligence_property_overrides
                    SET revoked_at = ?
                    WHERE user_key = ? AND property_id = ?
                      AND field_name = 'roof_area_sqft' AND revoked_at IS NULL
                    """,
                    (generated_at, user_key or created_by, report["property_id"]),
                )
                connection.execute(
                    """
                    INSERT INTO roof_intelligence_property_overrides (
                        id, user_key, property_id, field_name, numeric_value, source_revision_id,
                        reason, created_by, effective_at, revoked_at
                    ) VALUES (?, ?, ?, 'roof_area_sqft', ?, ?, ?, ?, ?, NULL)
                    """,
                    (
                        str(uuid.uuid4()),
                        user_key or created_by,
                        report["property_id"],
                        float(edits["roof_area_sqft"]),
                        revision_id,
                        change_reason,
                        created_by,
                        generated_at,
                    ),
                )
            if processing_feedback is not None:
                if str(processing_feedback.get("report_id") or "") != report_id:
                    raise ValueError("Processing feedback does not belong to this report.")
                if str(processing_feedback.get("parent_snapshot_id") or "") != parent_revision_id:
                    raise ValueError("Processing feedback has the wrong parent revision.")
                if str(processing_feedback.get("revised_snapshot_id") or "") != revision_id:
                    raise ValueError("Processing feedback has the wrong completed revision.")
                if processing_feedback.get("status") != "pending_review":
                    raise ValueError("New processing feedback must be pending review.")
                feedback_id = str(processing_feedback.get("feedback_id") or "").strip()
                if not feedback_id:
                    raise ValueError("Processing feedback is missing its ID.")
                connection.execute(
                    """
                    INSERT INTO roof_intelligence_processing_feedback (
                        id, report_id, property_id, parent_revision_id,
                        completed_revision_id, revision_number, requested_by,
                        comment, status, feedback_json, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending_review', ?, ?, ?)
                    """,
                    (
                        feedback_id,
                        report_id,
                        report["property_id"],
                        parent_revision_id,
                        revision_id,
                        next_number,
                        str(processing_feedback.get("requested_by") or created_by),
                        str(processing_feedback.get("comment") or change_reason),
                        json.dumps(processing_feedback, default=str),
                        str(processing_feedback.get("created_at") or generated_at),
                        generated_at,
                    ),
                )
        self._create_notification(
            created_by,
            report.get("job_id"),
            report_id,
            f"report_revision_{revision_id}",
            "Roof Intelligence revision ready",
            f"Revision {next_number} is ready for review.",
        )
        saved = self.get_report_revision(revision_id, user_key)
        if saved is None:
            raise RuntimeError("The report revision was not saved.")
        return saved

    def resolve_footprint_discrepancy(
        self,
        job_id: str,
        selected_source: str,
        reason: str,
        *,
        user_key: str = "local-user",
        item_id: str | None = None,
    ) -> dict:
        selected_source = str(selected_source or "").strip().lower()
        reason = " ".join(str(reason or "").split())
        if selected_source not in {"supabase", "county"}:
            raise ValueError("Select either the Supabase or county footprint.")
        if len(reason) < 10:
            raise ValueError("Enter an override reason of at least 10 characters.")
        job = self.get_job(job_id, user_key)
        if not job:
            raise KeyError(job_id)
        target = job
        if item_id:
            target = next((item for item in self.list_area_items(job_id) if item["id"] == item_id), None)
            if not target:
                raise KeyError(item_id)
        details = target.get("error_details") or {}
        if (target.get("error_code") or target.get("reason_code")) != "footprint_discrepancy":
            raise ValueError("This request is not waiting for a footprint discrepancy resolution.")
        override = {
            "selected_source": selected_source,
            "reason": reason,
            "county": details.get("county", ""),
            "parcel": details.get("parcel", ""),
        }
        now = utc_now()
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO footprint_resolutions (
                    id, job_id, item_id, user_key, county, parcel_number,
                    selected_source, reason, validation_json, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(uuid.uuid4()), job_id, item_id, user_key,
                    details.get("county"), details.get("parcel"), selected_source,
                    reason, json.dumps(details.get("footprint_validation") or {}, default=str), now,
                ),
            )
            if item_id:
                payload = dict(target.get("input") or {})
                payload["footprint_override"] = override
                connection.execute(
                    """
                    UPDATE roof_intelligence_job_items SET input_json = ?, status = 'pending',
                        stage = 'queued', reason_code = NULL, message = NULL,
                        error_details_json = '{}', started_at = NULL, finished_at = NULL
                    WHERE id = ? AND job_id = ?
                    """,
                    (json.dumps(payload), item_id, job_id),
                )
            else:
                payload = dict(job.get("input") or {})
                payload["footprint_override"] = override
                connection.execute(
                    "UPDATE roof_intelligence_jobs SET input_json = ? WHERE id = ?",
                    (json.dumps(payload), job_id),
                )
            connection.execute(
                """
                UPDATE roof_intelligence_jobs SET status = 'queued', stage = 'queued',
                    error_code = NULL, error_message = NULL, error_details_json = '{}',
                    retryable = 0, started_at = NULL, finished_at = NULL, updated_at = ?
                WHERE id = ?
                """,
                (now, job_id),
            )
            if item_id:
                self._refresh_area_counts(connection, job_id)
        self._create_notification(
            user_key, job_id, None, f"footprint_resolution_{item_id or 'job'}_{now}",
            "Footprint resolution recorded",
            f"The {selected_source} footprint was approved. The report has been queued again.",
        )
        return self.get_job(job_id, user_key)

    def record_county_health(self, payload: dict, user_key: str = "local-user") -> None:
        checked_at = str(payload.get("checked_at") or utc_now())
        notifications: list[tuple[str, str, str]] = []
        with self.connect() as connection:
            for result in payload.get("results") or []:
                county_key = normalize_address(result.get("county"))
                status = normalize_county_health_status(result.get("status"))
                normalized_result = dict(result)
                normalized_result["status"] = status
                previous = connection.execute(
                    "SELECT status, result_json FROM county_health_checks WHERE county_key = ? ORDER BY checked_at DESC LIMIT 1",
                    (county_key,),
                ).fetchone()
                connection.execute(
                    "INSERT INTO county_health_checks (id, county_key, status, result_json, checked_at) VALUES (?, ?, ?, ?, ?)",
                    (str(uuid.uuid4()), county_key, status, json.dumps(normalized_result, default=str), checked_at),
                )
                prior_result = json.loads(previous["result_json"]) if previous else {}
                previous_status = normalize_county_health_status(previous["status"]) if previous else None
                changed = (
                    (previous is None and status in {"degraded", "failed"})
                    or (
                        previous is not None
                        and (
                            previous_status != status
                            or prior_result.get("error") != normalized_result.get("error")
                        )
                    )
                )
                if changed:
                    titles = {
                        "failed": "County discovery health issue",
                        "degraded": "County discovery health degraded",
                        "healthy": "County discovery health recovered",
                    }
                    title = titles[status]
                    if normalized_result.get("error"):
                        message = normalized_result["error"]
                    elif status == "degraded":
                        message = (
                            f"{normalized_result.get('passed_count', 1)} of "
                            f"{normalized_result.get('sample_count', 2)} county samples passed."
                        )
                    else:
                        message = f"{normalized_result.get('county')} discovery checks are {status}."
                    notifications.append(
                        (
                            f"county_health_{county_key}_{checked_at}",
                            title,
                            f"{normalized_result.get('county')}: {message}",
                        )
                    )
        for kind, title, message in notifications:
            self._create_notification(user_key, None, None, kind, title, message)
        self.prune_county_health()

    def prune_county_health(self, retention_days: int = 180) -> int:
        retention_days = max(30, int(retention_days))
        cutoff = (dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=retention_days)).isoformat()
        with self.connect() as connection:
            cursor = connection.execute(
                "DELETE FROM county_health_checks WHERE checked_at < ?", (cutoff,)
            )
            return cursor.rowcount

    def list_county_health(self, limit: int = 20) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute(
                "SELECT * FROM county_health_checks ORDER BY checked_at DESC, county_key LIMIT ?",
                (max(1, min(int(limit), 200)),),
            ).fetchall()
        results = []
        for row in rows:
            item = dict(row)
            item["result"] = json.loads(item.pop("result_json") or "{}")
            results.append(item)
        return results

    def list_latest_county_health(self, limit: int = 20) -> list[dict]:
        """Return only the newest stored result for each county."""
        with self.connect() as connection:
            rows = connection.execute(
                """
                SELECT id, county_key, status, result_json, checked_at
                FROM (
                    SELECT checks.*,
                           ROW_NUMBER() OVER (
                               PARTITION BY county_key
                               ORDER BY checked_at DESC, rowid DESC
                           ) AS county_rank
                    FROM county_health_checks AS checks
                )
                WHERE county_rank = 1
                ORDER BY checked_at DESC, county_key
                LIMIT ?
                """,
                (max(1, min(int(limit), 200)),),
            ).fetchall()
        results = []
        for row in rows:
            item = dict(row)
            item["result"] = json.loads(item.pop("result_json") or "{}")
            results.append(item)
        return results

    def _create_notification(
        self,
        user_key: str,
        job_id: str | None,
        report_id: str | None,
        kind: str,
        title: str,
        message: str,
    ) -> None:
        with self.connect() as connection:
            connection.execute(
                """
                INSERT OR IGNORE INTO notifications (
                    id, user_key, job_id, report_id, kind, title, message, is_read, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 0, ?)
                """,
                (str(uuid.uuid4()), user_key, job_id, report_id, kind, title, message[:500], utc_now()),
            )

    def list_notifications(
        self,
        user_key: str = "local-user",
        limit: int = 10,
        job_id: str | None = None,
    ) -> list[dict]:
        sql = """
            SELECT * FROM notifications
            WHERE user_key = ?
        """
        params: list[object] = [user_key]
        if job_id is not None:
            sql += " AND job_id = ?"
            params.append(job_id)
        sql += " ORDER BY is_read ASC, created_at DESC LIMIT ?"
        params.append(limit)
        with self.connect() as connection:
            rows = connection.execute(sql, params).fetchall()
        return [self._notification_from_row(row) for row in rows]

    def mark_notification_read(self, notification_id: str, user_key: str = "local-user") -> None:
        with self.connect() as connection:
            connection.execute(
                """
                UPDATE notifications SET is_read = 1, read_at = ?
                WHERE id = ? AND user_key = ?
                """,
                (utc_now(), notification_id, user_key),
            )


def get_job_store() -> RoofIntelligenceJobStore:
    return RoofIntelligenceJobStore()
