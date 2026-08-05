"""Local-only PCS configuration stored outside the source repository."""

from __future__ import annotations

import json
import os
from pathlib import Path
import re
import tempfile

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


# Beta intentionally has no default production Supabase project.
DEFAULT_SUPABASE_URL = ""

_PROPOSAL_TRACKING_CUTOVER_KEYS = (
    "PROPOSAL_TRACKING_SUPABASE_ENABLED",
    "PROPOSAL_TRACKING_SUPABASE_READS_ENABLED",
    "PROPOSAL_TRACKING_SUPABASE_WRITES_ENABLED",
    "PROPOSAL_TRACKING_SUPABASE_SHADOW_WRITES_ENABLED",
)


def supabase_configuration() -> tuple[str, str]:
    """Return the server-side Supabase URL and secret without exposing it to templates."""
    data = _read_settings()
    url = (
        os.environ.get("PCS_SUPABASE_URL", "").strip()
        or str(data.get("supabase_url") or "").strip()
        or DEFAULT_SUPABASE_URL
    )
    key = (
        os.environ.get("PCS_SUPABASE_SERVICE_ROLE_KEY", "").strip()
        or str(data.get("supabase_service_role_key") or "").strip()
    )
    return url.rstrip("/"), key


def save_supabase_configuration(url_value: object, key_value: object) -> None:
    url = str(url_value or "").strip().rstrip("/")
    key = str(key_value or "").strip()
    hosted_url = re.fullmatch(r"https://[a-z0-9-]+\.supabase\.co", url)
    local_beta_url = (
        os.environ.get("PCS_APP_ENV", "").strip().lower() == "beta"
        and re.fullmatch(r"http://(?:127\.0\.0\.1|localhost):54321", url)
    )
    if not (hosted_url or local_beta_url):
        raise ValueError(
            "Enter a valid hosted Supabase URL, or the local beta URL "
            "http://127.0.0.1:54321."
        )
    if not (key.startswith("sb_secret_") or key.startswith("eyJ")) or len(key) < 32:
        raise ValueError("Enter a valid Supabase secret or legacy service-role key.")
    data = _read_settings()
    data["supabase_url"] = url
    data["supabase_service_role_key"] = key
    _write_settings(data)


def remove_supabase_configuration() -> None:
    data = _read_settings()
    data.pop("supabase_url", None)
    data.pop("supabase_service_role_key", None)
    _write_settings(data)


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
