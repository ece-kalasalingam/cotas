from __future__ import annotations

from datetime import datetime

from modules.coordinator import messages


class _View:
    def __init__(self) -> None:
        self.lines: list[str] = []

    def appendPlainText(self, text: str) -> None:  # noqa: N802
        self.lines.append(text)

    def clear(self) -> None:
        self.lines.clear()


class _Logger:
    def __init__(self) -> None:
        self.handlers = []

    def addHandler(self, handler) -> None:  # noqa: N802
        self.handlers.append(handler)


class _Module:
    def __init__(self) -> None:
        self._logger = _Logger()
        self._ui_log_handler = None
        self._user_log_entries: list[dict[str, object]] = []
        self.user_log_view = _View()
        self.status_changed = object()


def test_publish_status_and_key_emit_and_append() -> None:
    module = _Module()
    emitted: list[str] = []
    ns = {
        "emit_user_status": lambda _sig, message, logger=None: emitted.append(message),
        "t": lambda key, **kwargs: f"T:{key}",
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"__I18N__:{key}:{fallback}",
        "parse_i18n_log_message": lambda _m: None,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: f"[{(timestamp or datetime.now()).strftime('%H:%M:%S')}] {message}",
    }

    messages.publish_status(module, "plain", ns=ns)
    messages.publish_status_key(module, "coordinator.status.added", ns=ns, count=1)

    assert emitted[0] == "plain"
    assert emitted[1].startswith("__I18N__:")
    assert len(module.user_log_view.lines) == 2


def test_setup_ui_logging_idempotent() -> None:
    module = _Module()

    class _FakeHandler:
        def __init__(self, sink) -> None:
            self.sink = sink

    ns = {
        "UILogHandler": _FakeHandler,
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"payload:{key}",
        "t": lambda key, **kwargs: f"T:{key}",
        "parse_i18n_log_message": lambda _m: None,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: message,
    }

    messages.setup_ui_logging(module, ns=ns)
    first_handler = module._ui_log_handler
    messages.setup_ui_logging(module, ns=ns)

    assert module._ui_log_handler is first_handler
    assert len(module._logger.handlers) == 1


def test_rerender_user_log_uses_fallback_when_translation_raises() -> None:
    module = _Module()
    module._user_log_entries = [
        {
            "timestamp": datetime(2026, 3, 16, 22, 7, 37),
            "message": "raw",
            "text_key": "coordinator.status.added",
            "kwargs": {"count": 1},
            "fallback": "fallback-text",
        }
    ]

    ns = {
        "t": lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("boom")),
        "format_log_line_at": lambda message, timestamp=None: f"L:{message}",
    }

    messages.rerender_user_log(module, ns=ns)

    assert module.user_log_view.lines == ["L:fallback-text"]
