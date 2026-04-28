from __future__ import annotations

import logging
from datetime import datetime

from common import ui_logging


def test_format_log_line_at_blank_returns_none() -> None:
    """Test format log line at blank returns none.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    if ui_logging.format_log_line_at("") is not None:
        raise AssertionError('assertion failed')
    if ui_logging.format_log_line_at("   ") is not None:
        raise AssertionError('assertion failed')


def test_format_log_line_at_preserves_existing_timestamp() -> None:
    """Test format log line at preserves existing timestamp.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    line = "[10:20:30] already stamped"
    if not (ui_logging.format_log_line_at(line) == line):
        raise AssertionError('assertion failed')


def test_format_log_line_at_adds_marker_from_supplied_timestamp() -> None:
    """Test format log line at adds marker from supplied timestamp.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    ts = datetime(2026, 3, 16, 22, 7, 37)
    if not (ui_logging.format_log_line_at("hello", timestamp=ts) == "[22:07:37] hello"):
        raise AssertionError('assertion failed')


def test_build_and_parse_i18n_log_roundtrip() -> None:
    """Test build and parse i18n log roundtrip.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    msg = ui_logging.build_i18n_log_message(
        "co_analysis.status.added",
        kwargs={"count": 3},
        fallback="added",
    )
    parsed = ui_logging.parse_i18n_log_message(msg)
    if not (parsed == ("co_analysis.status.added", {"count": 3}, "added")):
        raise AssertionError('assertion failed')


def test_parse_i18n_log_message_rejects_invalid_payloads() -> None:
    """Test parse i18n log message rejects invalid payloads.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    if ui_logging.parse_i18n_log_message("plain text") is not None:
        raise AssertionError('assertion failed')
    if ui_logging.parse_i18n_log_message("__I18N_LOG__:{not-json") is not None:
        raise AssertionError('assertion failed')
    if ui_logging.parse_i18n_log_message("__I18N_LOG__:{}") is not None:
        raise AssertionError('assertion failed')


def test_resolve_i18n_log_message_resolves_prefix_payload(monkeypatch) -> None:
    """Test resolve i18n log message resolves prefix payload.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: f"T({key},{kwargs.get('count')})")
    payload = ui_logging.build_i18n_log_message("co_analysis.summary", kwargs={"count": 2})
    msg = f"INFO: {payload}"
    if not (ui_logging.resolve_i18n_log_message(msg) == "INFO: T(co_analysis.summary,2)"):
        raise AssertionError('assertion failed')


def test_resolve_i18n_log_message_falls_back_on_translation_failure(monkeypatch) -> None:
    """Test resolve i18n log message falls back on translation failure.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def boom(_key: str, **_kwargs: object) -> str:
        """Boom.
        
        Args:
            _key: Parameter value (str).
            _kwargs: Parameter value (object).
        
        Returns:
            str: Return value.
        
        Raises:
            None.
        """
        raise RuntimeError("boom")

    monkeypatch.setattr(ui_logging, "t", boom)
    payload = ui_logging.build_i18n_log_message(
        "co_analysis.summary",
        kwargs={"count": 2},
        fallback="safe fallback",
    )
    if not (ui_logging.resolve_i18n_log_message(payload) == "safe fallback"):
        raise AssertionError('assertion failed')


def test_resolve_i18n_log_message_falls_back_when_translation_returns_raw_key(monkeypatch) -> None:
    """Test resolve i18n log message falls back when translation returns raw key.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: key)
    payload = ui_logging.build_i18n_log_message(
        "validation.batch.title_error",
        kwargs={},
        fallback="Validation Error",
    )
    if not (ui_logging.resolve_i18n_log_message(payload) == "Validation Error"):
        raise AssertionError('assertion failed')


def test_resolve_i18n_kwargs_handles_nested_translation(monkeypatch) -> None:
    """Test resolve i18n kwargs handles nested translation.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: f"<{key}:{kwargs.get('n', '')}>")
    kwargs = {
        "title": {
            "__t_key__": "activity.log.ready",
            "kwargs": {"n": 5},
            "fallback": "fallback-title",
        }
    }
    resolved = ui_logging._resolve_i18n_kwargs(kwargs)
    if not (resolved["title"] == "<activity.log.ready:5>"):
        raise AssertionError('assertion failed')


def test_ui_log_handler_prefers_user_message_and_falls_back_to_level_prefix() -> None:
    """Test ui log handler prefers user message and falls back to level prefix.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    seen: list[str] = []
    handler = ui_logging.UILogHandler(seen.append)

    logger = logging.getLogger("test.ui_log_handler")
    logger.handlers.clear()
    logger.propagate = False
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    logger.info("normal")
    logger.info("ignored", extra={"user_message": "custom user text"})

    if not (any(line.startswith("INFO: normal") for line in seen)):
        raise AssertionError('assertion failed')
    if "custom user text" not in seen:
        raise AssertionError('assertion failed')

