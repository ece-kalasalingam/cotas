from __future__ import annotations

import json
from pathlib import Path

from common import crash_reporting


def test_capture_unhandled_exception_writes_json_report(tmp_path: Path, monkeypatch) -> None:
    """Test capture unhandled exception writes json report.
    
    Args:
        tmp_path: Parameter value (Path).
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    reports_dir = tmp_path / "crash_reports"
    monkeypatch.setattr(crash_reporting, "_crash_reports_dir", lambda: reports_dir)

    try:
        raise ValueError("boom")
    except ValueError as exc:
        report = crash_reporting.capture_unhandled_exception(type(exc), exc, exc.__traceback__)

    if not (report is not None):
        raise AssertionError('assertion failed')
    if not (report.exists()):
        raise AssertionError('assertion failed')
    payload = json.loads(report.read_text(encoding="utf-8"))
    if not (payload["app_name"]):
        raise AssertionError('assertion failed')
    if not (payload["version"]):
        raise AssertionError('assertion failed')
    if not (payload["exception_type"] == "ValueError"):
        raise AssertionError('assertion failed')
    if not (payload["exception_message"] == "boom"):
        raise AssertionError('assertion failed')
    if "ValueError: boom" not in payload["traceback"]:
        raise AssertionError('assertion failed')


def test_capture_unhandled_exception_returns_none_on_internal_failure(monkeypatch) -> None:
    """Test capture unhandled exception returns none on internal failure.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(crash_reporting, "_crash_reports_dir", lambda: (_ for _ in ()).throw(OSError("denied")))
    if crash_reporting.capture_unhandled_exception(RuntimeError, RuntimeError("x"), None) is not None:
        raise AssertionError('assertion failed')


def test_has_remote_crash_endpoint_env_parsing(monkeypatch) -> None:
    """Test has remote crash endpoint env parsing.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setenv(crash_reporting.CRASH_REPORT_ENDPOINT_ENV_VAR, "   ")
    if crash_reporting.has_remote_crash_endpoint() is not False:
        raise AssertionError('assertion failed')

    monkeypatch.setenv(crash_reporting.CRASH_REPORT_ENDPOINT_ENV_VAR, "https://example.invalid/collect")
    if crash_reporting.has_remote_crash_endpoint() is not True:
        raise AssertionError('assertion failed')


def test_crash_reports_dir_uses_settings_parent(monkeypatch) -> None:
    """Test crash reports dir uses settings parent.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(crash_reporting, "app_settings_path", lambda _name: Path("C:/focus/settings.json"))
    out = crash_reporting._crash_reports_dir()
    if not (out == Path("C:/focus") / crash_reporting.CRASH_REPORTS_DIR_NAME):
        raise AssertionError('assertion failed')

