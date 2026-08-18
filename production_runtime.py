"""Runtime configuration for the PCS production desktop application."""

from __future__ import annotations

import os
from pathlib import Path
from typing import MutableMapping

from integration_runtime import (
    PRODUCTION_EMAIL_LIST_DIR,
    PRODUCTION_EMAIL_TEMPLATE_DIR,
    PRODUCTION_PROPOSAL_TRACKER,
    _load_private_worker_environment,
)
from pcs_runtime_config import apply_proposal_storage_mode


PRODUCTION_PILOTPOINT_PROJECT_DIR = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/Visual Studio/"
    "PilotPoint IQ Roof Intelligence Report"
)
PRODUCTION_WORKSPACE_ROOT = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/"
    "Test Site"
)
PRODUCTION_ACCOUNTS_ROOT = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-ProfessionalCoatingSystems/PCS"
)


def production_environment(
    *,
    home: Path | None = None,
    pilotpoint_project_dir: Path | None = None,
) -> dict[str, str]:
    home = home or Path.home()
    data_dir = home / "Library" / "Application Support" / "PCS Proposal Management"
    pilotpoint_dir = pilotpoint_project_dir or PRODUCTION_PILOTPOINT_PROJECT_DIR
    values = {
        "PCS_APP_ENV": "production",
        "PCS_APP_DISPLAY_NAME": "PCS Proposal",
        "PCS_APP_STATE_DIR": "/tmp/pcs_proposal_app",
        "PCS_DEFAULT_PORT": "5050",
        "PCS_SERVER_STARTUP_TIMEOUT": "90",
        "PCS_DATA_DIR": str(data_dir),
        "PCS_SETTINGS_PATH": str(data_dir / "settings.json"),
        "PCS_XLWINGS_LOG_PATH": str(data_dir / "pcs_xlwings.log"),
        "PCS_PRODUCTION_WORKER_ENV_FILE": str(data_dir / "worker.env"),
        "PCS_MULTI_TENANT_ENABLED": "1",
        "ROOF_INTELLIGENCE_PROJECT_DIR": str(pilotpoint_dir),
        "ROOF_INTELLIGENCE_USER_KEY": "production-local-user",
        "PCS_PROPOSAL_TEMP_DIR": str(PRODUCTION_WORKSPACE_ROOT / "1. Open Proposals"),
        "PCS_CONTRACTS_DIR": str(PRODUCTION_WORKSPACE_ROOT / "2. Signed Contracts"),
        "PCS_COMPLETED_DIR": str(PRODUCTION_WORKSPACE_ROOT / "3. Finished Jobs"),
        "PCS_DEADFILE_DIR": str(PRODUCTION_WORKSPACE_ROOT / "4. Dead Proposals"),
        "PCS_TEMPLATE_DIR": str(PRODUCTION_WORKSPACE_ROOT / "Job Jacket Template"),
        "PCS_PROPOSALS_DIR": str(PRODUCTION_ACCOUNTS_ROOT / "1 - Open Proposals"),
        "PCS_DAVIDS_PROPOSALS_DIR": str(
            PRODUCTION_ACCOUNTS_ROOT / "David's Accounts" / "1 - Open Proposals"
        ),
        "PCS_LYDIAS_PROPOSALS_DIR": str(
            PRODUCTION_ACCOUNTS_ROOT / "Lydia's Accounts" / "1 - Open Proposals"
        ),
        "PCS_RANDYS_PROPOSALS_DIR": str(
            PRODUCTION_ACCOUNTS_ROOT / "Randy's Accounts" / "1 - Open Proposals"
        ),
        "PCS_EMAIL_TEMPLATE_DIR": str(PRODUCTION_EMAIL_TEMPLATE_DIR),
        "PCS_EMAIL_LIST_DIR": str(PRODUCTION_EMAIL_LIST_DIR),
        "PCS_PROPOSAL_TRACKER": str(PRODUCTION_PROPOSAL_TRACKER),
    }
    # Supabase is authoritative in production. Excel remains a synchronized
    # rollback/audit shadow during the cutover period.
    apply_proposal_storage_mode(values, "supabase_shadow")
    return values


def apply_production_environment(
    environ: MutableMapping[str, str] | None = None,
) -> dict[str, str]:
    environ = environ if environ is not None else os.environ
    pilotpoint_override = environ.get(
        "PCS_PRODUCTION_ROOF_INTELLIGENCE_PROJECT_DIR", ""
    ).strip()
    worker_environment_override = environ.get(
        "PCS_PRODUCTION_WORKER_ENV_FILE", ""
    ).strip()
    values = production_environment(
        pilotpoint_project_dir=(
            Path(pilotpoint_override) if pilotpoint_override else None
        )
    )
    if worker_environment_override:
        values["PCS_PRODUCTION_WORKER_ENV_FILE"] = worker_environment_override
    environ.update(values)

    production_url = environ.get("PCS_PRODUCTION_SUPABASE_URL", "").strip()
    production_key = environ.get(
        "PCS_PRODUCTION_SUPABASE_PUBLISHABLE_KEY", ""
    ).strip()
    if production_url and production_key:
        environ["PCS_SUPABASE_URL"] = production_url
        environ["PCS_SUPABASE_PUBLISHABLE_KEY"] = production_key
    else:
        # The installed app reads its publishable client configuration from
        # the protected production settings file.
        environ.pop("PCS_SUPABASE_URL", None)
        environ.pop("PCS_SUPABASE_PUBLISHABLE_KEY", None)
    environ.pop("PCS_SUPABASE_SERVICE_ROLE_KEY", None)
    _load_private_worker_environment(
        Path(environ["PCS_PRODUCTION_WORKER_ENV_FILE"]),
        environ,
    )
    return values


__all__ = [
    "apply_production_environment",
    "production_environment",
]
