"""Shared coordinator module message and UI-log helpers."""

from __future__ import annotations

from datetime import datetime


def publish_status(module: object, message: str, *, ns: dict[str, object]) -> None:
    append_user_log(module, message, ns=ns)
    ns["emit_user_status"](module.status_changed, message, logger=module._logger)


def publish_status_key(module: object, text_key: str, *, ns: dict[str, object], **kwargs: object) -> None:
    localized = ns["t"](text_key, **kwargs)
    payload = ns["build_i18n_log_message"](text_key, kwargs=kwargs, fallback=localized)
    append_user_log(module, payload, ns=ns)
    ns["emit_user_status"](module.status_changed, payload, logger=module._logger)


def setup_ui_logging(module: object, *, ns: dict[str, object]) -> None:
    if module._ui_log_handler is not None:
        return
    module._ui_log_handler = ns["UILogHandler"](lambda message: append_user_log(module, message, ns=ns))
    module._logger.addHandler(module._ui_log_handler)
    append_user_log(
        module,
        ns["build_i18n_log_message"](
            "instructor.log.ready",
            fallback=ns["t"]("instructor.log.ready"),
        ),
        ns=ns,
    )


def append_user_log(module: object, message: str, *, ns: dict[str, object]) -> None:
    parsed = ns["parse_i18n_log_message"](message)
    localized = ns["resolve_i18n_log_message"](message)
    timestamp = datetime.now()
    if parsed is None:
        module._user_log_entries.append({"timestamp": timestamp, "message": localized})
    else:
        key, kwargs, fallback = parsed
        module._user_log_entries.append(
            {
                "timestamp": timestamp,
                "message": localized,
                "text_key": key,
                "kwargs": kwargs,
                "fallback": fallback,
            }
        )
    line = ns["format_log_line_at"](localized, timestamp=timestamp)
    if line is None:
        return
    module.user_log_view.appendPlainText(line)


def rerender_user_log(module: object, *, ns: dict[str, object]) -> None:
    module.user_log_view.clear()
    for entry in module._user_log_entries:
        timestamp = entry.get("timestamp")
        text_key = entry.get("text_key")
        fallback = entry.get("fallback")
        kwargs = entry.get("kwargs")
        message = entry.get("message")
        if isinstance(text_key, str):
            safe_kwargs = kwargs if isinstance(kwargs, dict) else {}
            try:
                resolved = ns["t"](text_key, **safe_kwargs)
            except Exception:
                resolved = fallback if isinstance(fallback, str) else str(message or "")
        else:
            resolved = str(message or "")
        ts = timestamp if isinstance(timestamp, datetime) else None
        line = ns["format_log_line_at"](resolved, timestamp=ts)
        if line is None:
            continue
        module.user_log_view.appendPlainText(line)
