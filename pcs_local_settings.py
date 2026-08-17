"""Local-only PCS configuration stored outside the source repository."""

from __future__ import annotations

import json
import os
from pathlib import Path
import re
import tempfile
import secrets

from roof_intelligence_jobs import DEFAULT_DATA_DIR


def settings_path() -> Path:
    configured = os.environ.get("PCS_SETTINGS_PATH")
    return Path(configured) if configured else DEFAULT_DATA_DIR / "settings.json"


def _read_settings() -> dict:
    path = settings_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def google_maps_api_key() -> str:
    environment_key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if environment_key:
        return environment_key
    return str(_read_settings().get("google_maps_api_key") or "").strip()


def validate_google_maps_api_key(value: object) -> str:
    key = str(value or "").strip()
    if not re.fullmatch(r"AIza[0-9A-Za-z_-]{30,}", key):
        raise ValueError("Enter a valid Google Maps API key beginning with AIza.")
    return key


def save_google_maps_api_key(value: object) -> None:
    key = validate_google_maps_api_key(value)
    data = _read_settings()
    data["google_maps_api_key"] = key
    _write_settings(data)


def remove_google_maps_api_key() -> None:
    data = _read_settings()
    data.pop("google_maps_api_key", None)
    _write_settings(data)


# Beta intentionally has no default hosted Supabase project. The publishable
# key below belongs only to the standard local Supabase development stack and
# is safe to embed in a client application.
DEFAULT_SUPABASE_URL = ""
DEFAULT_LOCAL_SUPABASE_URL = "http://127.0.0.1:54321"
DEFAULT_LOCAL_SUPABASE_PUBLISHABLE_KEY = "sb_publishable_ACJWlzQHlZjBrEguHvfOxg_3BJgxAaH"

_PROPOSAL_TRACKING_CUTOVER_KEYS = (
    "PROPOSAL_TRACKING_SUPABASE_ENABLED",
    "PROPOSAL_TRACKING_SUPABASE_READS_ENABLED",
    "PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED",
    "PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED",
)


def supabase_configuration() -> tuple[str, str]:
    """Return the Supabase URL and publishable key used with a user's JWT."""
    data = _read_settings()
    url = (
        os.environ.get("PCS_SUPABASE_URL", "").strip()
        or str(data.get("supabase_url") or "").strip()
        or (
            DEFAULT_LOCAL_SUPABASE_URL
            if os.environ.get("PCS_ALLOW_LOCAL_SUPABASE", "").strip().lower()
            in {"1", "true", "yes", "on"}
            else DEFAULT_SUPABASE_URL
        )
    )
    key = (
        os.environ.get("PCS_SUPABASE_PUBLISHABLE_KEY", "").strip()
        or str(data.get("supabase_publishable_key") or "").strip()
        or (
            DEFAULT_LOCAL_SUPABASE_PUBLISHABLE_KEY
            if url in {DEFAULT_LOCAL_SUPABASE_URL, "http://localhost:54321"}
            else ""
        )
    )
    return url.rstrip("/"), key


def save_supabase_configuration(url_value: object, key_value: object) -> None:
    url = str(url_value or "").strip().rstrip("/")
    key = str(key_value or "").strip()
    hosted_url = re.fullmatch(r"https://[a-z0-9-]+\.supabase\.co", url)
    local_development_url = (
        os.environ.get("PCS_ALLOW_LOCAL_SUPABASE", "").strip().lower()
        in {"1", "true", "yes", "on"}
        and re.fullmatch(r"http://(?:127\.0\.0\.1|localhost):54321", url)
    )
    if not (hosted_url or local_development_url):
        raise ValueError(
            "Enter a valid hosted Supabase URL, or the permitted local URL "
            "http://127.0.0.1:54321."
        )
    if not (key.startswith("sb_publishable_") or key.startswith("eyJ")) or len(key) < 32:
        raise ValueError("Enter a valid Supabase publishable or legacy anon key.")
    if key.startswith("sb_secret_"):
        raise ValueError("Do not store a Supabase secret key in the desktop application.")
    data = _read_settings()
    data["supabase_url"] = url
    data["supabase_publishable_key"] = key
    data.pop("supabase_service_role_key", None)
    _write_settings(data)


def remove_supabase_configuration() -> None:
    data = _read_settings()
    data.pop("supabase_url", None)
    data.pop("supabase_publishable_key", None)
    data.pop("supabase_service_role_key", None)
    _write_settings(data)


def flask_secret_key() -> str:
    """Return a durable per-installation secret for signing local sessions."""
    environment_value = os.environ.get("PCS_FLASK_SECRET_KEY", "").strip()
    if environment_value:
        return environment_value
    data = _read_settings()
    value = str(data.get("flask_secret_key") or "").strip()
    if value:
        return value
    value = secrets.token_urlsafe(48)
    data["flask_secret_key"] = value
    _write_settings(data)
    return value


def report_export_directory() -> str:
    return str(_read_settings().get("report_export_directory") or "").strip()


def save_report_export_directory(value: object) -> str:
    raw_path = str(value or "").strip()
    if not raw_path:
        raise ValueError("Choose a local folder for exported reports.")
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        raise ValueError("Enter the full path to the local report folder.")
    path.mkdir(parents=True, exist_ok=True)
    if not path.is_dir() or not os.access(path, os.W_OK):
        raise ValueError("The selected report folder is not writable.")
    data = _read_settings()
    data["report_export_directory"] = str(path)
    _write_settings(data)
    return str(path)


def proposal_tracking_cutover_environment() -> dict[str, str]:
    """Return persistent proposal cutover flags in environment-variable form."""
    configured = _read_settings().get("proposal_tracking_cutover")
    if not isinstance(configured, dict):
        return {}
    return {
        key: str(configured[key])
        for key in _PROPOSAL_TRACKING_CUTOVER_KEYS
        if key in configured
    }


def save_proposal_tracking_cutover_configuration(
    *,
    enabled: bool,
    reads_enabled: bool,
    writes_enabled: bool,
    shadow_writes_enabled: bool,
) -> None:
    """Persist proposal cutover flags without modifying Supabase credentials."""
    data = _read_settings()
    data["proposal_tracking_cutover"] = {
        "PROPOSAL_TRACKING_SUPABASE_ENABLED": bool(enabled),
        "PROPOSAL_TRACKING_SUPABASE_READS_ENABLED": bool(reads_enabled),
        "PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED": bool(writes_enabled),
        "PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED": bool(
            shadow_writes_enabled
        ),
    }
    _write_settings(data)


def _write_settings(data: dict) -> None:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, temporary_name = tempfile.mkstemp(prefix=".settings-", suffix=".json", dir=path.parent)
    try:
        with os.fdopen(file_descriptor, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2, sort_keys=True)
            handle.write("\n")
        os.chmod(temporary_name, 0o600)
        os.replace(temporary_name, path)
        os.chmod(path, 0o600)
    finally:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
