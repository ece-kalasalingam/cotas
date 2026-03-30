from __future__ import annotations

from common.exceptions import ConfigurationError
from common.module_runtime import ModuleRuntime


class _Module:
    status_changed = object()
    _logger = object()
    _ui_log_handler = None
    _user_log_entries: list[dict[str, object]] = []
    user_log_view = object()


def _ns() -> dict[str, object]:
    return {
        "emit_user_status": lambda *_a, **_k: None,
        "t": lambda key, **kwargs: key,
        "build_i18n_log_message": lambda key, kwargs=None, fallback=None: key,
        "parse_i18n_log_message": lambda _m: None,
        "resolve_i18n_log_message": lambda m: m,
        "format_log_line_at": lambda message, timestamp=None: message,
        "UILogHandler": lambda sink: sink,
    }


def test_module_runtime_plain_status_and_append_are_disallowed() -> None:
    runtime = ModuleRuntime(
        module=_Module(),
        app_name="APP",
        logger=object(),  # type: ignore[arg-type]
        async_runner=object(),
        messages_namespace_factory=_ns,
    )
    try:
        runtime.publish_status("plain")
        raise AssertionError("expected ConfigurationError")
    except ConfigurationError:
        pass
    try:
        runtime.append_user_log("plain")
        raise AssertionError("expected ConfigurationError")
    except ConfigurationError:
        pass

