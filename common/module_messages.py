"""Shared module message + UI-log orchestration helpers."""

from __future__ import annotations

from datetime import datetime
from logging import Logger
from typing import Callable, Mapping, Protocol, TypedDict, cast


class _UserLogView(Protocol):
    def appendPlainText(self, text: str) -> None:  # noqa: N802 - Qt-style name
        ...

    def clear(self) -> None:
        ...


class _UILogHandlerLogger(Protocol):
    def addHandler(self, handler: object) -> None:  # noqa: N802 - Qt-style name
        ...


class _MessageModule(Protocol):
    status_changed: object
    _logger: _UILogHandlerLogger
    _ui_log_handler: object | None
    _user_log_entries: list[dict[str, object]]
    user_log_view: _UserLogView


class _EmitUserStatus(Protocol):
    def __call__(self, signal: object, message: str, *, logger: Logger | object = ...) -> None:
        ...


class _FormatLogLine(Protocol):
    def __call__(self, message: str, *, timestamp: datetime | None = ...) -> str | None:
        ...


class MessagesNamespace(TypedDict):
    emit_user_status: _EmitUserStatus
    t: Callable[..., str]
    build_i18n_log_message: Callable[..., str]
    parse_i18n_log_message: Callable[[str], tuple[str, dict[str, object], str | None] | None]
    resolve_i18n_log_message: Callable[[str], str]
    format_log_line_at: _FormatLogLine
    UILogHandler: Callable[[Callable[[str], None]], object]


def _emit_user_status(ns: MessagesNamespace, signal: object, message: str, logger: Logger | object) -> None:
    ns["emit_user_status"](signal, message, logger=logger)


def _format_log_line(ns: MessagesNamespace, message: str, timestamp: datetime | None) -> str | None:
    return ns["format_log_line_at"](message, timestamp=timestamp)


def publish_status(module: object, message: str, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_MessageModule, module)
    typed_ns = cast(MessagesNamespace, ns)
    append_user_log(typed_module, message, ns=typed_ns)
    _emit_user_status(typed_ns, typed_module.status_changed, message, logger=typed_module._logger)


def publish_status_key(
    module: object,
    text_key: str,
    *,
    ns: Mapping[str, object],
    **kwargs: object,
) -> None:
    typed_module = cast(_MessageModule, module)
    typed_ns = cast(MessagesNamespace, ns)
    localized = typed_ns["t"](text_key, **kwargs)
    payload = typed_ns["build_i18n_log_message"](text_key, kwargs=kwargs, fallback=localized)
    append_user_log(typed_module, payload, ns=typed_ns)
    _emit_user_status(typed_ns, typed_module.status_changed, payload, logger=typed_module._logger)


def setup_ui_logging(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_MessageModule, module)
    typed_ns = cast(MessagesNamespace, ns)
    if typed_module._ui_log_handler is not None:
        return
    typed_module._ui_log_handler = typed_ns["UILogHandler"](
        lambda message: append_user_log(typed_module, message, ns=typed_ns)
    )
    typed_module._logger.addHandler(typed_module._ui_log_handler)
    append_user_log(
        typed_module,
        typed_ns["build_i18n_log_message"](
            "instructor.log.ready",
            fallback=typed_ns["t"]("instructor.log.ready"),
        ),
        ns=typed_ns,
    )


def append_user_log(module: object, message: str, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_MessageModule, module)
    typed_ns = cast(MessagesNamespace, ns)
    parsed = typed_ns["parse_i18n_log_message"](message)
    localized = typed_ns["resolve_i18n_log_message"](message)
    timestamp = datetime.now()
    if parsed is None:
        typed_module._user_log_entries.append(
            {
                "timestamp": timestamp,
                "message": localized,
                "raw_message": message,
            }
        )
    else:
        key, kwargs, fallback = parsed
        typed_module._user_log_entries.append(
            {
                "timestamp": timestamp,
                "message": localized,
                "raw_message": message,
                "text_key": key,
                "kwargs": kwargs,
                "fallback": fallback,
            }
        )
    line = _format_log_line(typed_ns, localized, timestamp=timestamp)
    if line is None:
        return
    typed_module.user_log_view.appendPlainText(line)


def rerender_user_log(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_MessageModule, module)
    typed_ns = cast(MessagesNamespace, ns)
    typed_module.user_log_view.clear()
    for entry in typed_module._user_log_entries:
        timestamp = entry.get("timestamp")
        text_key = entry.get("text_key")
        fallback = entry.get("fallback")
        kwargs = entry.get("kwargs")
        message = entry.get("message")
        raw_message = entry.get("raw_message")
        if isinstance(text_key, str):
            safe_kwargs = kwargs if isinstance(kwargs, dict) else {}
            try:
                resolved = typed_ns["t"](text_key, **safe_kwargs)
            except Exception:
                resolved = fallback if isinstance(fallback, str) else str(message or "")
        else:
            if isinstance(raw_message, str):
                resolved = typed_ns["resolve_i18n_log_message"](raw_message)
            else:
                resolved = str(message or "")
        ts = timestamp if isinstance(timestamp, datetime) else None
        line = _format_log_line(typed_ns, resolved, timestamp=ts)
        if line is None:
            continue
        typed_module.user_log_view.appendPlainText(line)
