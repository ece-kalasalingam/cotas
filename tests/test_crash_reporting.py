from __future__ import annotations

import json
from pathlib import Path

from common import crash_reporting


def test_capture_unhandled_exception_writes_json_report(tmp_path: Path, monkeypatch) -> None:
    reports_dir = tmp_path / "crash_reports"
    monkeypatch.setattr(crash_reporting, "_crash_reports_dir", lambda: reports_dir)

    try:
        raise ValueError("boom")
    except ValueError as exc:
        report = crash_reporting.capture_unhandled_exception(type(exc), exc, exc.__traceback__)

    assert report is not None
    assert report.exists()
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["app_name"]
    assert payload["version"]
    assert payload["exception_type"] == "ValueError"
    assert payload["exception_message"] == "boom"
    assert "ValueError: boom" in payload["traceback"]


def test_capture_unhandled_exception_returns_none_on_internal_failure(monkeypatch) -> None:
    monkeypatch.setattr(crash_reporting, "_crash_reports_dir", lambda: (_ for _ in ()).throw(OSError("denied")))
    assert crash_reporting.capture_unhandled_exception(RuntimeError, RuntimeError("x"), None) is None


def test_has_remote_crash_endpoint_env_parsing(monkeypatch) -> None:
    monkeypatch.setenv(crash_reporting.CRASH_REPORT_ENDPOINT_ENV_VAR, "   ")
    assert crash_reporting.has_remote_crash_endpoint() is False

    monkeypatch.setenv(crash_reporting.CRASH_REPORT_ENDPOINT_ENV_VAR, "https://example.invalid/collect")
    assert crash_reporting.has_remote_crash_endpoint() is True


def test_crash_reports_dir_uses_settings_parent(monkeypatch) -> None:
    monkeypatch.setattr(crash_reporting, "app_settings_path", lambda _name: Path("C:/focus/settings.json"))
    out = crash_reporting._crash_reports_dir()
    assert out == Path("C:/focus") / crash_reporting.CRASH_REPORTS_DIR_NAME

