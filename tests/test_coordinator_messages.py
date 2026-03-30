from __future__ import annotations

from datetime import datetime

from common import module_messages as messages
from common.exceptions import ConfigurationError


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


def _parse_test_i18n(message: str):
    if isinstance(message, str) and (message.startswith("__I18N__:") or message.startswith("payload:")):
        return ("test.key", {}, message)
    return None


def test_publish_status_plain_is_disallowed_and_key_path_still_works() -> None:
    module = _Module()
    emitted: list[str] = []
    ns = {
        "emit_user_status": lambda _sig, message, logger=None: emitted.append(message),
        "t": lambda key, **kwargs: f"T:{key}",
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"__I18N__:{key}:{fallback}",
        "parse_i18n_log_message": _parse_test_i18n,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: f"[{(timestamp or datetime.now()).strftime('%H:%M:%S')}] {message}",
    }

    try:
        messages.publish_status(module, "plain", ns=ns)
        raise AssertionError("expected ConfigurationError")
    except ConfigurationError:
        pass
    messages.publish_status_key(module, "coordinator.status.added", ns=ns, count=1)

    assert emitted[0].startswith("__I18N__:")
    assert len(module.user_log_view.lines) == 1


def test_setup_ui_logging_idempotent() -> None:
    module = _Module()

    class _FakeHandler:
        def __init__(self, sink) -> None:
            self.sink = sink

    ns = {
        "UILogHandler": _FakeHandler,
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"__I18N__:{key}",
        "t": lambda key, **kwargs: f"T:{key}",
        "parse_i18n_log_message": _parse_test_i18n,
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


def test_append_user_log_plain_is_disallowed_and_rerender_keeps_existing_entries() -> None:
    module = _Module()
    ns = {
        "parse_i18n_log_message": lambda _m: None,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda *_a, **_k: None,
    }
    try:
        messages.append_user_log(module, "plain", ns=ns)
        raise AssertionError("expected ConfigurationError")
    except ConfigurationError:
        pass
    assert module.user_log_view.lines == []

    module._user_log_entries = [{"timestamp": datetime.now(), "message": "raw"}]
    ns2 = {
        "t": lambda key, **kwargs: f"T:{key}",
        "format_log_line_at": lambda message, timestamp=None: f"L:{message}",
    }
    messages.rerender_user_log(module, ns=ns2)
    assert module.user_log_view.lines == ["L:raw"]


def test_rerender_user_log_skips_none_formatted_line_branch() -> None:
    module = _Module()
    module._user_log_entries = [
        {"timestamp": datetime.now(), "message": "skip-me"},
        {"timestamp": datetime.now(), "message": "keep-me"},
    ]
    ns = {
        "t": lambda key, **kwargs: f"T:{key}",
        "format_log_line_at": lambda message, timestamp=None: None if message == "skip-me" else f"L:{message}",
    }
    messages.rerender_user_log(module, ns=ns)
    assert module.user_log_view.lines == ["L:keep-me"]


def test_notify_message_rejects_plain_for_status_activity_and_allows_toast_only() -> None:
    module = _Module()
    captured_toasts: list[str] = []
    ns = {
        "emit_user_status": lambda _sig, message, logger=None: None,
        "t": lambda key, **kwargs: f"T:{key}",
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"__I18N__:{key}:{fallback}",
        "parse_i18n_log_message": _parse_test_i18n,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: message,
    }
    try:
        messages.notify_message(module, "plain", ns=ns, channels=("status", "activity_log"))
        raise AssertionError("expected ConfigurationError")
    except ConfigurationError:
        pass

    original_show_toast_plain = messages.show_toast_plain

    def _capture_toast(_widget, message: str, *, title: str, level: str = "info", duration_ms=None) -> None:
        del title, level, duration_ms
        captured_toasts.append(message)

    messages.show_toast_plain = _capture_toast  # type: ignore[assignment]
    try:
        messages.notify_message(module, "plain", ns=ns, channels=("toast",), toast_title="Title", toast_level="info")
    finally:
        messages.show_toast_plain = original_show_toast_plain  # type: ignore[assignment]
    assert captured_toasts == ["plain"]


def test_emit_validation_batch_feedback_emits_issues_and_summary_toast() -> None:
    module = _Module()
    emitted: list[str] = []
    toasts: list[tuple[str, str]] = []
    ns = {
        "emit_user_status": lambda _sig, message, logger=None: emitted.append(message),
        "t": lambda key, **kwargs: f"T:{key}:{kwargs}",
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"__I18N__:{key}:{kwargs}:{fallback}",
        "parse_i18n_log_message": _parse_test_i18n,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: message,
    }
    original_show_toast_plain = messages.show_toast_plain

    def _capture_toast(_widget, message: str, *, title: str, level: str = "info", duration_ms=None) -> None:
        del duration_ms
        toasts.append((title, level))
        emitted.append(message)

    messages.show_toast_plain = _capture_toast  # type: ignore[assignment]
    try:
        messages.emit_validation_batch_feedback(
            module,
            ns=ns,
            rejections=[
                {
                    "path": "a.xlsx",
                    "issue": {
                        "code": "MARKS_TEMPLATE_COHORT_MISMATCH",
                        "translation_key": "validation.marks_template.cohort_mismatch",
                        "message": "mismatch",
                        "context": {"workbook": "a.xlsx", "fields": "Course_Code"},
                    },
                }
            ],
            valid_count=1,
        )
    finally:
        messages.show_toast_plain = original_show_toast_plain  # type: ignore[assignment]

    assert any("validation.batch.details_prefix" in line for line in module.user_log_view.lines)
    assert any("validation.batch.activity_line" in str(entry.get("message", "")) for entry in module._user_log_entries)
    assert toasts and toasts[-1][1] == "warning"


def test_notify_validation_issue_handles_missing_translation_kwargs() -> None:
    module = _Module()
    emitted: list[str] = []
    ns = {
        "emit_user_status": lambda _sig, message, logger=None: emitted.append(message),
        "t": lambda key, **kwargs: ("needs {sheet}".format(**kwargs) if key == "validation.layout.sheet_missing" else f"T:{key}"),
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"__I18N__:{key}:{fallback}",
        "parse_i18n_log_message": _parse_test_i18n,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: message,
    }

    messages.notify_validation_issue(
        module,
        ns=ns,
        issue={
            "code": "COA_LAYOUT_SHEET_MISSING",
            "translation_key": "validation.layout.sheet_missing",
            "message": "Missing required layout sheet.",
            "context": {"workbook": "x.xlsx"},
        },
        file_path="x.xlsx",
        channels=("activity_log",),
    )

    assert module.user_log_view.lines


def test_emit_validation_batch_feedback_uses_localized_generic_fallback_when_reason_translation_fails() -> None:
    module = _Module()
    ns = {
        "emit_user_status": lambda _sig, message, logger=None: None,
        "t": lambda key, **kwargs: (
            (_ for _ in ()).throw(KeyError("sheet"))
            if key == "validation.layout.sheet_missing"
            else ("T:common.validation_failed_invalid_data" if key == "common.validation_failed_invalid_data" else f"T:{key}:{kwargs}")
        ),
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: f"__I18N__:{key}:{kwargs}:{fallback}",
        "parse_i18n_log_message": _parse_test_i18n,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: message,
    }
    original_show_toast_plain = messages.show_toast_plain
    messages.show_toast_plain = lambda *_a, **_k: None  # type: ignore[assignment]
    try:
        messages.emit_validation_batch_feedback(
            module,
            ns=ns,
            rejections=[
                {
                    "path": "C:/a.xlsx",
                    "issue": {
                        "code": "COA_LAYOUT_SHEET_MISSING",
                        "translation_key": "validation.layout.sheet_missing",
                        "message": "Validation failed due to invalid data.",
                        "context": {"workbook": "C:/a.xlsx"},
                    },
                }
            ],
            valid_count=0,
        )
    finally:
        messages.show_toast_plain = original_show_toast_plain  # type: ignore[assignment]
    assert any("T:common.validation_failed_invalid_data" in line for line in module.user_log_view.lines)


def test_notify_validation_issue_formats_fallback_with_sheet_alias() -> None:
    module = _Module()
    ns = {
        "emit_user_status": lambda *_a, **_k: None,
        "t": lambda key, **kwargs: (_ for _ in ()).throw(KeyError("sheet_name")),
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: (
            f"__I18N__:{key}:{kwargs}:{fallback}"
        ),
        "parse_i18n_log_message": _parse_test_i18n,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: message,
    }

    messages.notify_validation_issue(
        module,
        ns=ns,
        issue={
            "code": "COA_MARK_ENTRY_EMPTY",
            "translation_key": "validation.mark.entry_empty",
            "message": "Sheet '{sheet_name}' has an empty mark-entry cell at {cell}.",
            "context": {"sheet": "CIA 1", "cell": "D10"},
        },
        file_path="x.xlsx",
        channels=("activity_log",),
    )

    assert module.user_log_view.lines
    assert "CIA 1" in module.user_log_view.lines[-1]
    assert "{sheet_name}" not in module.user_log_view.lines[-1]
