"""Shared module message + UI-log orchestration helpers."""

from __future__ import annotations

from datetime import datetime
from logging import Logger
from pathlib import Path
from typing import Any, Callable, Iterable, Literal, Mapping, Protocol, TypedDict, cast

from PySide6.QtWidgets import QWidget

from common.exceptions import ConfigurationError
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


def _context_with_aliases(context: Mapping[str, object]) -> dict[str, object]:
    payload = dict(context)
    if "sheet_name" not in payload and "sheet" in payload:
        payload["sheet_name"] = payload.get("sheet")
    if "sheet" not in payload and "sheet_name" in payload:
        payload["sheet"] = payload.get("sheet_name")
    return payload


def _format_reason_fallback(template: str, context: Mapping[str, object]) -> str:
    text = str(template or "").strip()
    if not text:
        return text
    payload = _context_with_aliases(context)
    try:
        return text.format(**payload)
    except Exception:
        return text


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


def _safe_translate(ns: MessagesNamespace, key: str, **kwargs: object) -> str:
    try:
        text = ns["t"](key, **kwargs)
    except Exception:
        return key
    return text if isinstance(text, str) and text.strip() else key


def _ensure_i18n_for_log_channels(*, message: str, ns: MessagesNamespace, channels: set[NotificationChannel]) -> None:
    if "status" not in channels and "activity_log" not in channels:
        return
    parsed = ns["parse_i18n_log_message"](message)
    if parsed is not None:
        return
    raise ConfigurationError(
        "Plain text is not allowed for status/activity_log channels. "
        "Use notify_message_key(...) or build_status_message(...)."
    )


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
    _ensure_i18n_for_log_channels(message=message, ns=typed_ns, channels=targets)
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


def notify_validation_issue(
    module: object,
    *,
    ns: Mapping[str, object],
    issue: Mapping[str, object],
    file_path: str | None = None,
    channels: Iterable[NotificationChannel] | None = ("activity_log",),
) -> None:
    typed_ns = cast(MessagesNamespace, ns)
    translation_key = str(issue.get("translation_key", "") or "").strip()
    code = str(issue.get("code", "") or "").strip()
    fallback_reason = str(issue.get("message", "") or "").strip()
    raw_context = issue.get("context")
    context: dict[str, object] = {}
    if isinstance(raw_context, Mapping):
        for key, value in raw_context.items():
            if isinstance(key, str):
                context[key] = value

    if translation_key:
        if file_path:
            reason_kwargs = _context_with_aliases(context)
            reason_text: str
            reason_payload: object
            try:
                reason_text = typed_ns["t"](translation_key, **reason_kwargs)
                reason_payload = {"__t_key__": translation_key, "__kwargs__": reason_kwargs}
            except Exception:
                reason_text = _format_reason_fallback(fallback_reason, reason_kwargs) or translation_key
                reason_payload = reason_text
            payload_kwargs: dict[str, object] = {
                "file": file_path,
                "code": code or "VALIDATION_ERROR",
                "reason": reason_payload,
            }
            try:
                fallback = typed_ns["t"](
                    "instructor.validation.file_issue_line",
                    file=file_path,
                    code=code or "VALIDATION_ERROR",
                    reason=reason_text,
                )
            except Exception:
                fallback = f"{file_path}: [{code or 'VALIDATION_ERROR'}] {reason_text}"
            notify_message(
                module,
                typed_ns["build_i18n_log_message"](
                    "instructor.validation.file_issue_line",
                    kwargs=payload_kwargs,
                    fallback=fallback,
                ),
                ns=typed_ns,
                channels=_normalize_channels(channels, default=("activity_log",)),
            )
            return
        notify_message_key(
            module,
            translation_key,
            ns=typed_ns,
            channels=channels,
            kwargs=context,
            fallback=fallback_reason or None,
        )
        return

    if file_path and fallback_reason:
        formatted_reason = _format_reason_fallback(fallback_reason, context)
        notify_message(
            module,
            typed_ns["build_i18n_log_message"](
                "instructor.validation.file_issue_line",
                kwargs={
                    "file": file_path,
                    "code": code or "VALIDATION_ERROR",
                    "reason": formatted_reason,
                },
                fallback=f"{file_path}: [{code or 'VALIDATION_ERROR'}] {formatted_reason}",
            ),
            ns=typed_ns,
            channels=_normalize_channels(channels, default=("activity_log",)),
        )
        return

    if fallback_reason:
        formatted_reason = _format_reason_fallback(fallback_reason, context)
        payload = typed_ns["build_i18n_log_message"](
            "common.validation_failed_invalid_data",
            kwargs={},
            fallback=formatted_reason,
        )
        notify_message(module, payload, ns=typed_ns, channels=channels)


def emit_validation_batch_feedback(
    module: object,
    *,
    ns: Mapping[str, object],
    rejections: Iterable[Mapping[str, object]],
    valid_count: int,
    issue_channels: Iterable[NotificationChannel] | None = ("status", "activity_log"),
    summary_channels: Iterable[NotificationChannel] | None = ("toast",),
) -> None:
    typed_ns = cast(MessagesNamespace, ns)
    rejection_items = [
        item
        for item in rejections
        if isinstance(item, Mapping) and isinstance(item.get("issue"), Mapping)
    ]
    rejected_count = len(rejection_items)
    accepted_count = max(0, int(valid_count))
    title_key = (
        "validation.batch.title_success"
        if rejected_count == 0 and accepted_count > 0
        else "validation.batch.title_error"
    )
    title_text = _safe_translate(typed_ns, title_key)

    accepted_text = (
        _safe_translate(typed_ns, "validation.batch.accepted_count", count=accepted_count)
        if accepted_count > 0
        else ""
    )
    rejected_text = (
        _safe_translate(typed_ns, "validation.batch.rejected_count", count=rejected_count)
        if rejected_count > 0
        else ""
    )

    detail_entries_text: list[str] = []
    detail_entries_marker: list[object] = []
    for item in rejection_items:
        issue_payload = cast(Mapping[str, object], item.get("issue"))
        raw_path = item.get("path")
        path_text = str(raw_path).strip() if isinstance(raw_path, str) else ""
        file_label = Path(path_text).name if path_text else ""
        code = str(issue_payload.get("code", "VALIDATION_ERROR")).strip() or "VALIDATION_ERROR"
        translation_key = str(issue_payload.get("translation_key", "")).strip()
        fallback_reason = str(issue_payload.get("message", "")).strip()
        issue_context_raw = issue_payload.get("context")
        issue_context = dict(issue_context_raw) if isinstance(issue_context_raw, Mapping) else {}
        generic_reason = _safe_translate(typed_ns, "common.validation_failed_invalid_data")
        reason_text = _format_reason_fallback(fallback_reason, issue_context)
        reason_marker: object = reason_text
        reason_context = _context_with_aliases(issue_context)
        if translation_key:
            try:
                translated = typed_ns["t"](translation_key, **reason_context)
                if isinstance(translated, str) and translated.strip() and translated != translation_key:
                    reason_text = translated
                else:
                    reason_text = generic_reason
            except Exception:
                reason_text = generic_reason
            reason_marker = {
                "__t_key__": translation_key,
                "kwargs": reason_context,
                "fallback": reason_text or generic_reason,
            }
        elif not reason_text:
            reason_text = code
            reason_marker = reason_text
        file_text = file_label or path_text or "-"
        entry = f"{file_text}: [{code}] {reason_text}"
        detail_entries_text.append(entry)
        detail_entries_marker.append(
            {
                "__t_key__": "validation.batch.detail_entry",
                "kwargs": {"file": file_text, "code": code, "reason": reason_marker},
                "fallback": entry,
            }
        )

    details_preview_marker: object = ""
    details_text = ""
    if detail_entries_text:
        max_lines = 3
        preview = "; ".join(detail_entries_text[:max_lines])
        hidden = len(detail_entries_text) - max_lines
        if hidden > 0:
            preview = (
                f"{preview}; "
                + _safe_translate(typed_ns, "validation.batch.more_suffix", count=hidden)
            )
        details_text = _safe_translate(typed_ns, "validation.batch.details_prefix", details=preview)
        if len(detail_entries_marker) == 1:
            details_preview_marker = {
                "__t_key__": "validation.batch.details_entries_1",
                "kwargs": {"entry1": detail_entries_marker[0]},
                "fallback": detail_entries_text[0],
            }
        elif len(detail_entries_marker) == 2:
            details_preview_marker = {
                "__t_key__": "validation.batch.details_entries_2",
                "kwargs": {"entry1": detail_entries_marker[0], "entry2": detail_entries_marker[1]},
                "fallback": "; ".join(detail_entries_text[:2]),
            }
        elif hidden > 0:
            details_preview_marker = {
                "__t_key__": "validation.batch.details_entries_3_more",
                "kwargs": {
                    "entry1": detail_entries_marker[0],
                    "entry2": detail_entries_marker[1],
                    "entry3": detail_entries_marker[2],
                    "more": {
                        "__t_key__": "validation.batch.more_suffix",
                        "kwargs": {"count": hidden},
                        "fallback": f"+{hidden} more",
                    },
                },
                "fallback": preview,
            }
        else:
            details_preview_marker = {
                "__t_key__": "validation.batch.details_entries_3",
                "kwargs": {
                    "entry1": detail_entries_marker[0],
                    "entry2": detail_entries_marker[1],
                    "entry3": detail_entries_marker[2],
                },
                "fallback": "; ".join(detail_entries_text[:3]),
            }

    toast_lines = [line for line in (accepted_text, rejected_text) if line]
    toast_body = "\n".join(toast_lines) if toast_lines else title_text
    toast_level: ToastLevel = "success" if rejected_count == 0 and accepted_count > 0 else "warning"
    notify_message(
        module,
        toast_body,
        ns=typed_ns,
        channels=summary_channels,
        toast_title=title_text,
        toast_level=toast_level,
    )

    accepted_segment: object = ""
    if accepted_count > 0:
        accepted_segment = {
            "__t_key__": "validation.batch.activity_segment",
            "kwargs": {
                "segment": {
                    "__t_key__": "validation.batch.accepted_count",
                    "kwargs": {"count": accepted_count},
                }
            },
        }
    rejected_segment: object = ""
    if rejected_count > 0:
        rejected_segment = {
            "__t_key__": "validation.batch.activity_segment",
            "kwargs": {
                "segment": {
                    "__t_key__": "validation.batch.rejected_count",
                    "kwargs": {"count": rejected_count},
                }
            },
        }
    details_segment: object = ""
    if details_text:
        details_segment = {
            "__t_key__": "validation.batch.activity_segment",
            "kwargs": {
                "segment": {
                    "__t_key__": "validation.batch.details_prefix",
                    "kwargs": {"details": details_preview_marker},
                    "fallback": details_text,
                }
            },
        }

    if accepted_segment or rejected_segment or details_segment:
        activity_payload = typed_ns["build_i18n_log_message"](
            "validation.batch.activity_line",
            kwargs={
                "title": {"__t_key__": title_key, "kwargs": {}},
                "accepted": accepted_segment,
                "rejected": rejected_segment,
                "details": details_segment,
            },
            fallback=(
                f"{title_text} | "
                + " | ".join(part for part in (accepted_text, rejected_text, details_text) if part)
            ),
        )
        notify_message(
            module,
            activity_payload,
            ns=typed_ns,
            channels=issue_channels,
        )


def emit_workbook_generation_feedback(
    module: object,
    *,
    ns: Mapping[str, object],
    success_count: int,
    failed_count: int,
    channels: Iterable[NotificationChannel] | None = ("status", "activity_log", "toast"),
) -> None:
    typed_ns = cast(MessagesNamespace, ns)
    success = max(0, int(success_count))
    failed = max(0, int(failed_count))
    segments: list[str] = []
    if success > 0:
        segments.append(_safe_translate(typed_ns, "workbook.generation.segment.success", count=success))
    if failed > 0:
        segments.append(_safe_translate(typed_ns, "workbook.generation.segment.failed", count=failed))
    if not segments:
        segments.append(_safe_translate(typed_ns, "workbook.generation.segment.none"))
    notify_message_key(
        module,
        "workbook.generation.summary",
        ns=typed_ns,
        channels=_normalize_channels(channels, default=("status", "activity_log", "toast")),
        kwargs={"segments": ", ".join(segments)},
        toast_title_key="instructor.msg.success_title",
        toast_level="success" if failed == 0 and success > 0 else "warning",
    )


def publish_status(module: object, message: str, *, ns: Mapping[str, object]) -> None:
    del module, message, ns
    raise ConfigurationError(
        "publish_status(message) is disallowed. Use publish_status_key(...) or notify_message_key(...)."
    )


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
    _ensure_i18n_for_log_channels(
        message=message,
        ns=typed_ns,
        channels={"activity_log"},
    )
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
