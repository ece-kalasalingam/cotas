"""Shared module message + UI-log orchestration helpers."""

from __future__ import annotations

from datetime import datetime
from logging import Logger
from typing import Any, Callable, Iterable, Literal, Mapping, Protocol, TypedDict, cast

from PySide6.QtWidgets import QWidget

from common.toast import show_toast
from common.ui_logging import (
    UILogHandler,
    build_i18n_log_message,
    format_log_line_at,
    parse_i18n_log_message,
    resolve_i18n_log_message,
)
from common.utils import emit_user_status


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


class _FormatLogLine(Protocol):
    def __call__(self, message: str, *, timestamp: datetime | None = ...) -> str | None:
        ...


class MessagesNamespace(TypedDict):
    emit_user_status: Callable[..., None]
    t: Callable[..., str]
    build_i18n_log_message: Callable[..., str]
    parse_i18n_log_message: Callable[[str], tuple[str, dict[str, object], str | None] | None]
    resolve_i18n_log_message: Callable[[str], str]
    format_log_line_at: _FormatLogLine
    UILogHandler: Callable[[Callable[[str], None]], object]


ToastLevel = Literal["info", "success", "warning", "error"]
NotificationChannel = Literal["status", "toast", "activity_log"]


def default_messages_namespace(*, translate: Callable[..., str]) -> MessagesNamespace:
    return {
        "emit_user_status": emit_user_status,
        "t": translate,
        "build_i18n_log_message": build_i18n_log_message,
        "parse_i18n_log_message": parse_i18n_log_message,
        "resolve_i18n_log_message": resolve_i18n_log_message,
        "format_log_line_at": format_log_line_at,
        "UILogHandler": UILogHandler,
    }


def build_status_message(
    text_key: str,
    *,
    translate: Callable[..., str],
    kwargs: Mapping[str, Any] | None = None,
    fallback: str | None = None,
) -> str:
    payload_kwargs = dict(kwargs or {})
    resolved_fallback = fallback if isinstance(fallback, str) else translate(text_key, **payload_kwargs)
    return build_i18n_log_message(text_key, kwargs=payload_kwargs, fallback=resolved_fallback)


def resolve_status_message(message: str) -> str:
    return resolve_i18n_log_message(message)


def show_toast_plain(
    widget: object,
    message: str,
    *,
    title: str,
    level: ToastLevel = "info",
    duration_ms: int | None = None,
) -> None:
    show_toast(cast(QWidget | None, widget), message, title=title, level=level, duration_ms=duration_ms)


def show_toast_key(
    widget: object,
    *,
    text_key: str,
    title_key: str,
    translate: Callable[..., str],
    level: ToastLevel = "info",
    text_kwargs: Mapping[str, Any] | None = None,
    title_kwargs: Mapping[str, Any] | None = None,
    duration_ms: int | None = None,
) -> None:
    show_toast(
        cast(QWidget | None, widget),
        translate(text_key, **dict(text_kwargs or {})),
        title=translate(title_key, **dict(title_kwargs or {})),
        level=level,
        duration_ms=duration_ms,
    )


def _emit_user_status(ns: MessagesNamespace, signal: object, message: str, logger: Logger | object) -> None:
    ns["emit_user_status"](signal, message, logger=logger)


def _format_log_line(ns: MessagesNamespace, message: str, timestamp: datetime | None) -> str | None:
    return ns["format_log_line_at"](message, timestamp=timestamp)


def _normalize_channels(
    channels: Iterable[NotificationChannel] | None,
    *,
    default: tuple[NotificationChannel, ...],
) -> set[NotificationChannel]:
    default_set: set[NotificationChannel] = {channel for channel in default}
    if channels is None:
        return set(default_set)
    normalized: set[NotificationChannel] = set()
    for channel in channels:
        normalized.add(channel)
    if not normalized:
        return set(default_set)
    return normalized


def notify_message(
    module: object,
    message: str,
    *,
    ns: Mapping[str, object],
    channels: Iterable[NotificationChannel] | None = None,
    toast_title: str = "",
    toast_level: ToastLevel = "info",
    toast_duration_ms: int | None = None,
) -> None:
    typed_module = cast(_MessageModule, module)
    typed_ns = cast(MessagesNamespace, ns)
    targets = _normalize_channels(channels, default=("status", "activity_log"))
    if "activity_log" in targets:
        append_user_log(typed_module, message, ns=typed_ns)
    if "status" in targets:
        _emit_user_status(typed_ns, typed_module.status_changed, message, logger=typed_module._logger)
    if "toast" in targets:
        show_toast_plain(
            typed_module,
            typed_ns["resolve_i18n_log_message"](message),
            title=toast_title,
            level=toast_level,
            duration_ms=toast_duration_ms,
        )


def notify_message_key(
    module: object,
    text_key: str,
    *,
    ns: Mapping[str, object],
    channels: Iterable[NotificationChannel] | None = None,
    kwargs: Mapping[str, Any] | None = None,
    fallback: str | None = None,
    toast_title: str = "",
    toast_title_key: str | None = None,
    toast_title_kwargs: Mapping[str, Any] | None = None,
    toast_level: ToastLevel = "info",
    toast_duration_ms: int | None = None,
) -> None:
    typed_ns = cast(MessagesNamespace, ns)
    payload_kwargs = dict(kwargs or {})
    localized = typed_ns["t"](text_key, **payload_kwargs)
    payload = typed_ns["build_i18n_log_message"](
        text_key,
        kwargs=payload_kwargs,
        fallback=fallback if isinstance(fallback, str) else localized,
    )
    resolved_toast_title = toast_title
    if not resolved_toast_title and isinstance(toast_title_key, str):
        resolved_toast_title = typed_ns["t"](toast_title_key, **dict(toast_title_kwargs or {}))
    notify_message(
        module,
        payload,
        ns=typed_ns,
        channels=_normalize_channels(channels, default=("status", "activity_log")),
        toast_title=resolved_toast_title,
        toast_level=toast_level,
        toast_duration_ms=toast_duration_ms,
    )


def publish_status(module: object, message: str, *, ns: Mapping[str, object]) -> None:
    notify_message(module, message, ns=ns, channels=("status", "activity_log"))


def publish_status_key(
    module: object,
    text_key: str,
    *,
    ns: Mapping[str, object],
    **kwargs: object,
) -> None:
    notify_message_key(module, text_key, ns=ns, channels=("status", "activity_log"), kwargs=kwargs)


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
            "activity.log.ready",
            fallback=typed_ns["t"]("activity.log.ready"),
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
                payload = typed_ns["build_i18n_log_message"](
                    text_key,
                    kwargs=safe_kwargs,
                    fallback=fallback if isinstance(fallback, str) else None,
                )
                resolved = typed_ns["resolve_i18n_log_message"](payload)
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
