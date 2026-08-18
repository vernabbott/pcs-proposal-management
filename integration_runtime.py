"""Isolated runtime configuration for the PCS production integration build."""

from __future__ import annotations

import os
from pathlib import Path
from typing import MutableMapping

from pcs_runtime_config import apply_proposal_storage_mode


INTEGRATION_PILOTPOINT_PROJECT_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/"
    "PilotPoint IQ Roof Intelligence Report Beta"
)

PRODUCTION_EMAIL_TEMPLATE_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/"
    "PCS/Marketing/Email Templates"
)
PRODUCTION_EMAIL_LIST_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/"
    "PCS/Marketing/Email Lists"
)
PRODUCTION_PROPOSAL_TRACKER = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/"
    "PCS/1 - Open Proposals/Proposal Tracking.xlsx"
)

_WORKER_ENVIRONMENT_KEYS = frozenset(
    {
        "DATABASE_URL",
        "OPENAI_API_KEY",
        "SUPABASE_DB_HOST",
        "SUPABASE_DB_NAME",
        "SUPABASE_DB_PASSWORD",
        "SUPABASE_DB_PORT",
        "SUPABASE_DB_USER",
    }
)


def _load_private_worker_environment(
    path: Path,
    environ: MutableMapping[str, str],
) -> None:
    """Load approved worker-only credentials from a local private file."""
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (FileNotFoundError, OSError):
        return
    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if key not in _WORKER_ENVIRONMENT_KEYS or key in environ:
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        if value:
            environ[key] = value


def integration_environment(
    *,
    home: Path | None = None,
    pilotpoint_project_dir: Path | None = None,
) -> dict[str, str]:
    home = home or Path.home()
    data_dir = home / "Library" / "Application Support" / "PCS Proposal Integration"
    workspace = data_dir / "Workspace"
    accounts = workspace / "Accounts"
    pilotpoint_dir = pilotpoint_project_dir or INTEGRATION_PILOTPOINT_PROJECT_DIR
    values = {
        "PCS_APP_ENV": "integration",
        "PCS_APP_DISPLAY_NAME": "PCS Proposal Integration",
        "PCS_APP_STATE_DIR": "/tmp/pcs_proposal_integration_app",
        "PCS_DEFAULT_PORT": "5052",
        "PCS_SERVER_STARTUP_TIMEOUT": "90",
        "PCS_DATA_DIR": str(data_dir),
        "PCS_SETTINGS_PATH": str(data_dir / "settings.json"),
        "PCS_XLWINGS_LOG_PATH": str(data_dir / "pcs_xlwings.log"),
        "PCS_INTEGRATION_WORKER_ENV_FILE": str(data_dir / "worker.env"),
        "PCS_MULTI_TENANT_ENABLED": "1",
        "PCS_ALLOW_LOCAL_SUPABASE": "1",
        "ROOF_INTELLIGENCE_PROJECT_DIR": str(pilotpoint_dir),
        "ROOF_INTELLIGENCE_USER_KEY": "integration-local-user",
        "PCS_PROPOSAL_TEMP_DIR": str(workspace / "1. Open Proposals"),
        "PCS_CONTRACTS_DIR": str(workspace / "2. Signed Contracts"),
        "PCS_COMPLETED_DIR": str(workspace / "3. Finished Jobs"),
        "PCS_DEADFILE_DIR": str(workspace / "4. Dead Proposals"),
        "PCS_TEMPLATE_DIR": str(workspace / "Job Jacket Template"),
        "PCS_PROPOSALS_DIR": str(accounts / "PCS" / "1 - Open Proposals"),
        "PCS_DAVIDS_PROPOSALS_DIR": str(accounts / "David" / "1 - Open Proposals"),
        "PCS_LYDIAS_PROPOSALS_DIR": str(accounts / "Lydia" / "1 - Open Proposals"),
        "PCS_RANDYS_PROPOSALS_DIR": str(accounts / "Randy" / "1 - Open Proposals"),
        "PCS_EMAIL_TEMPLATE_DIR": str(PRODUCTION_EMAIL_TEMPLATE_DIR),
        "PCS_EMAIL_LIST_DIR": str(PRODUCTION_EMAIL_LIST_DIR),
        "PCS_PROPOSAL_TRACKER": str(PRODUCTION_PROPOSAL_TRACKER),
        "PCS_PROPOSALS_WEB_URL": "",
        "DAVIDS_PROPOSALS_WEB_URL": "",
        "LYDIAS_PROPOSALS_WEB_URL": "",
        "RANDYS_PROPOSALS_WEB_URL": "",
    }
    # Supabase is authoritative. The production workbook is maintained only as
    # a rollback/audit shadow and is never used for integration reads.
    apply_proposal_storage_mode(values, "supabase_shadow")
    return values


def apply_integration_environment(
    environ: MutableMapping[str, str] | None = None,
) -> dict[str, str]:
    environ = environ if environ is not None else os.environ
    pilotpoint_override = environ.get(
        "PCS_INTEGRATION_ROOF_INTELLIGENCE_PROJECT_DIR", ""
    ).strip()
    worker_environment_override = environ.get(
        "PCS_INTEGRATION_WORKER_ENV_FILE", ""
    ).strip()
    values = integration_environment(
        pilotpoint_project_dir=(
            Path(pilotpoint_override) if pilotpoint_override else None
        )
    )
    if worker_environment_override:
        values["PCS_INTEGRATION_WORKER_ENV_FILE"] = worker_environment_override
    environ.update(values)

    integration_url = environ.get("PCS_INTEGRATION_SUPABASE_URL", "").strip()
    integration_key = environ.get(
        "PCS_INTEGRATION_SUPABASE_PUBLISHABLE_KEY", ""
    ).strip()
    if integration_url and integration_key:
        environ["PCS_SUPABASE_URL"] = integration_url
        environ["PCS_SUPABASE_PUBLISHABLE_KEY"] = integration_key
    else:
        # Do not inherit another app's project. With no explicit environment
        # override, pcs_local_settings reads this integration installation's
        # own settings file and otherwise falls back to the local beta stack.
        environ.pop("PCS_SUPABASE_URL", None)
        environ.pop("PCS_SUPABASE_PUBLISHABLE_KEY", None)
    # A desktop build must never inherit a privileged server credential.
    environ.pop("PCS_SUPABASE_SERVICE_ROLE_KEY", None)
    _load_private_worker_environment(
        Path(environ["PCS_INTEGRATION_WORKER_ENV_FILE"]),
        environ,
    )
    return values
