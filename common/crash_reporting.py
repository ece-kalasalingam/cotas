"""Crash reporting helpers for packaged builds."""

from __future__ import annotations

import json
import os
import platform
import traceback
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

from common.constants import APP_NAME, SYSTEM_VERSION
from common.utils import app_settings_path

CRASH_REPORT_ENDPOINT_ENV_VAR = "FOCUS_CRASH_REPORT_ENDPOINT"
CRASH_REPORTS_DIR_NAME = "crash_reports"


def capture_unhandled_exception(exc_type: object, exc_value: object, exc_traceback: object) -> Path | None:
    """Persist unhandled exceptions into a local crash spool."""
    try:
        reports_dir = _crash_reports_dir()
        reports_dir.mkdir(parents=True, exist_ok=True)
        report_path = reports_dir / f"{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}_{uuid4().hex}.json"
        payload = {
            "app_name": APP_NAME,
            "version": SYSTEM_VERSION,
            "timestamp_utc": datetime.now(timezone.utc).isoformat(),
            "platform": platform.platform(),
            "python_version": platform.python_version(),
            "exception_type": getattr(exc_type, "__name__", str(exc_type)),
            "exception_message": str(exc_value),
            "traceback": "".join(traceback.format_exception(exc_type, exc_value, exc_traceback)),
        }
        report_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return report_path
    except Exception:
        return None


def has_remote_crash_endpoint() -> bool:
    return bool(os.getenv(CRASH_REPORT_ENDPOINT_ENV_VAR, "").strip())


def _crash_reports_dir() -> Path:
    settings_path = app_settings_path(APP_NAME)
    return settings_path.parent / CRASH_REPORTS_DIR_NAME

