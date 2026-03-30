from __future__ import annotations

import logging
from datetime import datetime

from common import ui_logging


def test_format_log_line_at_blank_returns_none() -> None:
    assert ui_logging.format_log_line_at("") is None
    assert ui_logging.format_log_line_at("   ") is None


def test_format_log_line_at_preserves_existing_timestamp() -> None:
    line = "[10:20:30] already stamped"
    assert ui_logging.format_log_line_at(line) == line


def test_format_log_line_at_adds_marker_from_supplied_timestamp() -> None:
    ts = datetime(2026, 3, 16, 22, 7, 37)
    assert ui_logging.format_log_line_at("hello", timestamp=ts) == "[22:07:37] hello"


def test_build_and_parse_i18n_log_roundtrip() -> None:
    msg = ui_logging.build_i18n_log_message(
        "coordinator.status.added",
        kwargs={"count": 3},
        fallback="added",
    )
    parsed = ui_logging.parse_i18n_log_message(msg)
    assert parsed == ("coordinator.status.added", {"count": 3}, "added")


def test_parse_i18n_log_message_rejects_invalid_payloads() -> None:
    assert ui_logging.parse_i18n_log_message("plain text") is None
    assert ui_logging.parse_i18n_log_message("__I18N_LOG__:{not-json") is None
    assert ui_logging.parse_i18n_log_message("__I18N_LOG__:{}") is None


def test_resolve_i18n_log_message_resolves_prefix_payload(monkeypatch) -> None:
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: f"T({key},{kwargs.get('count')})")
    payload = ui_logging.build_i18n_log_message("coordinator.summary", kwargs={"count": 2})
    msg = f"INFO: {payload}"
    assert ui_logging.resolve_i18n_log_message(msg) == "INFO: T(coordinator.summary,2)"


def test_resolve_i18n_log_message_falls_back_on_translation_failure(monkeypatch) -> None:
    def boom(_key: str, **_kwargs: object) -> str:
        raise RuntimeError("boom")

    monkeypatch.setattr(ui_logging, "t", boom)
    payload = ui_logging.build_i18n_log_message(
        "coordinator.summary",
        kwargs={"count": 2},
        fallback="safe fallback",
    )
    assert ui_logging.resolve_i18n_log_message(payload) == "safe fallback"


def test_resolve_i18n_log_message_falls_back_when_translation_returns_raw_key(monkeypatch) -> None:
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: key)
    payload = ui_logging.build_i18n_log_message(
        "validation.batch.title_error",
        kwargs={},
        fallback="Validation Error",
    )
    assert ui_logging.resolve_i18n_log_message(payload) == "Validation Error"


def test_resolve_i18n_kwargs_handles_nested_translation(monkeypatch) -> None:
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: f"<{key}:{kwargs.get('n', '')}>")
    kwargs = {
        "title": {
            "__t_key__": "activity.log.ready",
            "kwargs": {"n": 5},
            "fallback": "fallback-title",
        }
    }
    resolved = ui_logging._resolve_i18n_kwargs(kwargs)
    assert resolved["title"] == "<activity.log.ready:5>"


def test_ui_log_handler_prefers_user_message_and_falls_back_to_level_prefix() -> None:
    seen: list[str] = []
    handler = ui_logging.UILogHandler(seen.append)

    logger = logging.getLogger("test.ui_log_handler")
    logger.handlers.clear()
    logger.propagate = False
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    logger.info("normal")
    logger.info("ignored", extra={"user_message": "custom user text"})

    assert any(line.startswith("INFO: normal") for line in seen)
    assert "custom user text" in seen
