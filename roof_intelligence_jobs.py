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
import os
from pathlib import Path
import re
import sqlite3
import uuid


APP_DIR = Path(__file__).resolve().parent
DEFAULT_DB_PATH = APP_DIR / "data" / "roof_intelligence_jobs.sqlite3"
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
        self.db_path = Path(db_path or os.environ.get("ROOF_INTELLIGENCE_DB_PATH") or DEFAULT_DB_PATH)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize()

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

                CREATE TABLE IF NOT EXISTS roof_intelligence_job_items (
                    id TEXT PRIMARY KEY,
                    job_id TEXT NOT NULL REFERENCES roof_intelligence_jobs(id),
                    property_id TEXT REFERENCES properties(id),
                    candidate_key TEXT,
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL,
                    reason_code TEXT,
                    message TEXT,
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
                """
            )

    @staticmethod
    def _job_from_row(row: sqlite3.Row | None) -> dict | None:
        if row is None:
            return None
        result = dict(row)
        result["input"] = json.loads(result.pop("input_json") or "{}")
        result["roof_types"] = json.loads(result.pop("roof_types_json") or "[]")
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
        assignments = ", ".join(f"{key} = ?" for key in updates)
        with self.connect() as connection:
            cursor = connection.execute(
                f"UPDATE roof_intelligence_jobs SET {assignments} WHERE id = ?",
                (*updates.values(), job_id),
            )
            if cursor.rowcount != 1:
                raise KeyError(job_id)
        return self.get_job(job_id)

    def start_job(self, job_id: str, stage: str = "locating_property", worker_version: str = "pcs-local-adapter") -> dict:
        return self.update_job(
            job_id,
            status="running",
            stage=stage,
            started_at=utc_now(),
            worker_version=worker_version,
            error_code=None,
            error_message=None,
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
    ) -> dict:
        message = " ".join(str(error_message or "Unable to complete the Roof Intelligence job.").split())[:500]
        job = self.update_job(
            job_id,
            status="failed",
            stage=stage,
            error_code=error_code,
            error_message=message,
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

    def complete_individual_job(self, job_id: str, result: dict) -> dict:
        job = self.get_job(job_id)
        if job is None:
            raise KeyError(job_id)
        fallback_address = job["input"].get("property_address", "")
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
            report_id = existing_report["id"] if existing_report else str(uuid.uuid4())
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
        return self.get_job(job_id)

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

    def get_report(self, report_id: str) -> dict | None:
        with self.connect() as connection:
            row = connection.execute(
                "SELECT * FROM roof_intelligence_reports WHERE id = ?", (report_id,)
            ).fetchone()
        if row is None:
            return None
        result = dict(row)
        result["result"] = json.loads(result.pop("result_json") or "{}")
        return result

    def _create_notification(
        self,
        user_key: str,
        job_id: str,
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

    def list_notifications(self, user_key: str = "local-user", limit: int = 10) -> list[dict]:
        with self.connect() as connection:
            rows = connection.execute(
                """
                SELECT * FROM notifications
                WHERE user_key = ?
                ORDER BY is_read ASC, created_at DESC
                LIMIT ?
                """,
                (user_key, limit),
            ).fetchall()
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
