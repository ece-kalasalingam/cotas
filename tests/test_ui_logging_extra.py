from __future__ import annotations

import logging
from typing import Any, cast

from common import ui_logging


def test_parse_i18n_log_message_non_dict_payload_and_non_string_fallback() -> None:
    """Test parse i18n log message non dict payload and non string fallback.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    assert ui_logging.parse_i18n_log_message("__I18N_LOG__:[1,2,3]") is None
    msg = "__I18N_LOG__:{\"key\":\"k\",\"kwargs\":{},\"fallback\":123}"
    parsed = ui_logging.parse_i18n_log_message(msg)
    assert parsed == ("k", {}, None)


def test_resolve_i18n_log_message_non_string_and_invalid_embedded_payload(monkeypatch) -> None:
    """Test resolve i18n log message non string and invalid embedded payload.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: f"T:{key}")
    assert ui_logging.resolve_i18n_log_message(cast(Any, 42)) == 42
    bad = "INFO: __I18N_LOG__:{bad-json"
    assert ui_logging.resolve_i18n_log_message(bad) == bad


def test_resolve_i18n_log_message_embedded_payload_fallback_on_translation_error(monkeypatch) -> None:
    """Test resolve i18n log message embedded payload fallback on translation error.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    payload = '__I18N_LOG__:{"key":"k","kwargs":{},"fallback":"fb"}'
    monkeypatch.setattr(ui_logging, "t", lambda *_a, **_k: (_ for _ in ()).throw(RuntimeError("x")))
    assert ui_logging.resolve_i18n_log_message(f"INFO: {payload}") == "INFO: fb"


def test_resolve_i18n_log_message_embedded_payload_fallback_on_raw_key(monkeypatch) -> None:
    """Test resolve i18n log message embedded payload fallback on raw key.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    payload = '__I18N_LOG__:{"key":"k","kwargs":{},"fallback":"fb"}'
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: key)
    assert ui_logging.resolve_i18n_log_message(f"INFO: {payload}") == "INFO: fb"


def test_resolve_i18n_kwargs_nested_translation_fallback_on_error(monkeypatch) -> None:
    """Test resolve i18n kwargs nested translation fallback on error.
    
    Args:
        monkeypatch: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _boom(*_args, **_kwargs):
        """Boom.
        
        Args:
            _args: Parameter value.
            _kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        raise RuntimeError("x")

    monkeypatch.setattr(ui_logging, "t", _boom)
    resolved = ui_logging._resolve_i18n_kwargs(
        {"title": {"__t_key__": "x", "kwargs": {"n": 1}, "fallback": "fb"}}
    )
    assert resolved["title"] == "fb"


def test_ui_log_handler_emit_handles_sink_exceptions() -> None:
    """Test ui log handler emit handles sink exceptions.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    seen: list[logging.LogRecord] = []

    class _Handler(ui_logging.UILogHandler):
        def handleError(self, record: logging.LogRecord) -> None:  # noqa: N802
            """Handleerror.
            
            Args:
                record: Parameter value (logging.LogRecord).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            seen.append(record)

    handler = _Handler(lambda _m: (_ for _ in ()).throw(RuntimeError("sink-fail")))
    logger = logging.getLogger("test.ui_log_handler.error")
    logger.handlers.clear()
    logger.propagate = False
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    logger.info("boom")
    assert len(seen) == 1
