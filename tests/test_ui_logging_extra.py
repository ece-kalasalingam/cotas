from __future__ import annotations

import logging

from common import ui_logging


def test_format_log_line_wrapper_delegates() -> None:
    assert ui_logging.format_log_line("hello") is not None


def test_parse_i18n_log_message_non_dict_payload_and_non_string_fallback() -> None:
    assert ui_logging.parse_i18n_log_message("__I18N_LOG__:[1,2,3]") is None
    msg = "__I18N_LOG__:{\"key\":\"k\",\"kwargs\":{},\"fallback\":123}"
    parsed = ui_logging.parse_i18n_log_message(msg)
    assert parsed == ("k", {}, None)


def test_resolve_i18n_log_message_non_string_and_invalid_embedded_payload(monkeypatch) -> None:
    monkeypatch.setattr(ui_logging, "t", lambda key, **kwargs: f"T:{key}")
    assert ui_logging.resolve_i18n_log_message(42) == 42
    bad = "INFO: __I18N_LOG__:{bad-json"
    assert ui_logging.resolve_i18n_log_message(bad) == bad


def test_resolve_i18n_kwargs_nested_translation_fallback_on_error(monkeypatch) -> None:
    def _boom(*_args, **_kwargs):
        raise RuntimeError("x")

    monkeypatch.setattr(ui_logging, "t", _boom)
    resolved = ui_logging._resolve_i18n_kwargs(
        {"title": {"__t_key__": "x", "kwargs": {"n": 1}, "fallback": "fb"}}
    )
    assert resolved["title"] == "fb"


def test_ui_log_handler_emit_handles_sink_exceptions() -> None:
    seen: list[logging.LogRecord] = []

    class _Handler(ui_logging.UILogHandler):
        def handleError(self, record: logging.LogRecord) -> None:  # noqa: N802
            seen.append(record)

    handler = _Handler(lambda _m: (_ for _ in ()).throw(RuntimeError("sink-fail")))
    logger = logging.getLogger("test.ui_log_handler.error")
    logger.handlers.clear()
    logger.propagate = False
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    logger.info("boom")
    assert len(seen) == 1
