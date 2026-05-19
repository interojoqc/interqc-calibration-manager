from __future__ import annotations

from cloud_integrations import cloud_status


def use_sheets_backend() -> bool:
    status = cloud_status()
    return status.google_libs and status.credentials and status.spreadsheet_id


if use_sheets_backend():
    from sheets_store import *  # noqa: F403

    BACKEND_NAME = "Google Sheets"
else:
    from store import *  # noqa: F403

    BACKEND_NAME = "Local SQLite"
