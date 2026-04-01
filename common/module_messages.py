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
        """Appendplaintext.
        
        Args:
            text: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        ...

    def clear(self) -> None:
        """Clear.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        ...


class _UILogHandlerLogger(Protocol):
    def addHandler(self, handler: object) -> None:  # noqa: N802 - Qt-style name
        """Addhandler.
        
        Args:
            handler: Parameter value (object).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        ...


class _MessageModule(Protocol):
    status_changed: object
    _logger: _UILogHandlerLogger
    _ui_log_handler: object | None
    _user_log_entries: list[dict[str, object]]
    user_log_view: _UserLogView


class _FormatLogLine(Protocol):
    def __call__(self, message: str, *, timestamp: datetime | None = ...) -> str | None:
        """Call.
        
        Args:
            message: Parameter value (str).
            timestamp: Parameter value (datetime | None).
        
        Returns:
            str | None: Return value.
        
        Raises:
            None.
        """
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
_VALIDATION_ACTIVITY_DETAIL_LIMIT = 60


def default_messages_namespace(*, translate: Callable[..., str]) -> MessagesNamespace:
    """Default messages namespace.
    
    Args:
        translate: Parameter value (Callable[..., str]).
    
    Returns:
        MessagesNamespace: Return value.
    
    Raises:
        None.
    """
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
    """Build status message.
    
    Args:
        text_key: Parameter value (str).
        translate: Parameter value (Callable[..., str]).
        kwargs: Parameter value (Mapping[str, Any] | None).
        fallback: Parameter value (str | None).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    payload_kwargs = dict(kwargs or {})
    resolved_fallback = fallback if isinstance(fallback, str) else translate(text_key, **payload_kwargs)
    return build_i18n_log_message(text_key, kwargs=payload_kwargs, fallback=resolved_fallback)


def resolve_status_message(message: str) -> str:
    """Resolve status message.
    
    Args:
        message: Parameter value (str).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    return resolve_i18n_log_message(message)


def show_toast_plain(
    widget: object,
    message: str,
    *,
    title: str,
    level: ToastLevel = "info",
    duration_ms: int | None = None,
) -> None:
    """Show toast plain.
    
    Args:
        widget: Parameter value (object).
        message: Parameter value (str).
        title: Parameter value (str).
        level: Parameter value (ToastLevel).
        duration_ms: Parameter value (int | None).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Show toast key.
    
    Args:
        widget: Parameter value (object).
        text_key: Parameter value (str).
        title_key: Parameter value (str).
        translate: Parameter value (Callable[..., str]).
        level: Parameter value (ToastLevel).
        text_kwargs: Parameter value (Mapping[str, Any] | None).
        title_kwargs: Parameter value (Mapping[str, Any] | None).
        duration_ms: Parameter value (int | None).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    show_toast(
        cast(QWidget | None, widget),
        translate(text_key, **dict(text_kwargs or {})),
        title=translate(title_key, **dict(title_kwargs or {})),
        level=level,
        duration_ms=duration_ms,
    )


def _emit_user_status(ns: MessagesNamespace, signal: object, message: str, logger: Logger | object) -> None:
    """Emit user status.
    
    Args:
        ns: Parameter value (MessagesNamespace).
        signal: Parameter value (object).
        message: Parameter value (str).
        logger: Parameter value (Logger | object).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    ns["emit_user_status"](signal, message, logger=logger)


def _format_log_line(ns: MessagesNamespace, message: str, timestamp: datetime | None) -> str | None:
    """Format log line.
    
    Args:
        ns: Parameter value (MessagesNamespace).
        message: Parameter value (str).
        timestamp: Parameter value (datetime | None).
    
    Returns:
        str | None: Return value.
    
    Raises:
        None.
    """
    return ns["format_log_line_at"](message, timestamp=timestamp)


def _context_with_aliases(context: Mapping[str, object]) -> dict[str, object]:
    """Context with aliases.
    
    Args:
        context: Parameter value (Mapping[str, object]).
    
    Returns:
        dict[str, object]: Return value.
    
    Raises:
        None.
    """
    payload = dict(context)
    if "sheet_name" not in payload and "sheet" in payload:
        payload["sheet_name"] = payload.get("sheet")
    if "sheet" not in payload and "sheet_name" in payload:
        payload["sheet"] = payload.get("sheet_name")
    return payload


def _excel_column_label(column: object) -> str:
    """Excel column label.
    
    Args:
        column: Parameter value (object).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    if isinstance(column, int):
        index = int(column)
    else:
        text = str(column or "").strip()
        if not text:
            return ""
        if text.isdigit():
            index = int(text)
        else:
            return text.upper()
    if index <= 0:
        return ""
    labels: list[str] = []
    while index > 0:
        index, rem = divmod(index - 1, 26)
        labels.append(chr(ord("A") + rem))
    return "".join(reversed(labels))


def _build_validation_location_suffix(*, path_text: str, context: Mapping[str, object]) -> str:
    """Build validation location suffix.
    
    Args:
        path_text: Parameter value (str).
        context: Parameter value (Mapping[str, object]).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    payload = _context_with_aliases(context)
    workbook_raw = str(payload.get("workbook", "") or "").strip() or path_text
    workbook_name = Path(workbook_raw).name if workbook_raw else ""
    sheet_name = str(payload.get("sheet_name", "") or "").strip()
    cell_ref = str(payload.get("cell", "") or payload.get("cell_ref", "") or "").strip()
    row_value = payload.get("row")
    row_text = str(row_value).strip() if row_value is not None else ""
    col_value = payload.get("column", payload.get("col"))
    col_label = _excel_column_label(col_value)

    if not cell_ref and row_text and col_label:
        cell_ref = f"{col_label}{row_text}"

    segments: list[str] = []
    if workbook_name:
        segments.append(f"workbook={workbook_name}")
    if sheet_name:
        segments.append(f"sheet={sheet_name}")
    if cell_ref:
        segments.append(f"cell={cell_ref}")
    if row_text:
        segments.append(f"row={row_text}")
    if col_label:
        segments.append(f"column={col_label}")
    if not segments:
        return ""
    return " | " + ", ".join(segments)


def _flatten_issue_payloads(issue_payload: Mapping[str, object], *, max_depth: int = 8) -> list[Mapping[str, object]]:
    """Flatten issue payloads.
    
    Args:
        issue_payload: Parameter value (Mapping[str, object]).
        max_depth: Parameter value (int).
    
    Returns:
        list[Mapping[str, object]]: Return value.
    
    Raises:
        None.
    """
    flattened: list[Mapping[str, object]] = []

    def _visit(payload: Mapping[str, object], depth: int) -> None:
        """Visit.
        
        Args:
            payload: Parameter value (Mapping[str, object]).
            depth: Parameter value (int).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if depth >= max_depth:
            flattened.append(payload)
            return
        context_raw = payload.get("context")
        context = dict(context_raw) if isinstance(context_raw, Mapping) else {}
        nested_raw = context.get("issues")
        nested_items = [item for item in nested_raw if isinstance(item, Mapping)] if isinstance(nested_raw, list) else []
        if not nested_items:
            flattened.append(payload)
            return
        for nested in nested_items:
            _visit(cast(Mapping[str, object], nested), depth + 1)

    _visit(issue_payload, 0)
    return flattened


def _format_reason_fallback(template: str, context: Mapping[str, object]) -> str:
    """Format reason fallback.
    
    Args:
        template: Parameter value (str).
        context: Parameter value (Mapping[str, object]).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
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
    """Normalize channels.
    
    Args:
        channels: Parameter value (Iterable[NotificationChannel] | None).
        default: Parameter value (tuple[NotificationChannel, ...]).
    
    Returns:
        set[NotificationChannel]: Return value.
    
    Raises:
        None.
    """
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
    """Safe translate.
    
    Args:
        ns: Parameter value (MessagesNamespace).
        key: Parameter value (str).
        kwargs: Parameter value (object).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    try:
        text = ns["t"](key, **kwargs)
    except Exception:
        return key
    return text if isinstance(text, str) and text.strip() else key


def _ensure_i18n_for_log_channels(*, message: str, ns: MessagesNamespace, channels: set[NotificationChannel]) -> None:
    """Ensure i18n for log channels.
    
    Args:
        message: Parameter value (str).
        ns: Parameter value (MessagesNamespace).
        channels: Parameter value (set[NotificationChannel]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Notify message.
    
    Args:
        module: Parameter value (object).
        message: Parameter value (str).
        ns: Parameter value (Mapping[str, object]).
        channels: Parameter value (Iterable[NotificationChannel] | None).
        toast_title: Parameter value (str).
        toast_level: Parameter value (ToastLevel).
        toast_duration_ms: Parameter value (int | None).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Notify message key.
    
    Args:
        module: Parameter value (object).
        text_key: Parameter value (str).
        ns: Parameter value (Mapping[str, object]).
        channels: Parameter value (Iterable[NotificationChannel] | None).
        kwargs: Parameter value (Mapping[str, Any] | None).
        fallback: Parameter value (str | None).
        toast_title: Parameter value (str).
        toast_title_key: Parameter value (str | None).
        toast_title_kwargs: Parameter value (Mapping[str, Any] | None).
        toast_level: Parameter value (ToastLevel).
        toast_duration_ms: Parameter value (int | None).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Notify validation issue.
    
    Args:
        module: Parameter value (object).
        ns: Parameter value (Mapping[str, object]).
        issue: Parameter value (Mapping[str, object]).
        file_path: Parameter value (str | None).
        channels: Parameter value (Iterable[NotificationChannel] | None).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Emit validation batch feedback.
    
    Args:
        module: Parameter value (object).
        ns: Parameter value (Mapping[str, object]).
        rejections: Parameter value (Iterable[Mapping[str, object]]).
        valid_count: Parameter value (int).
        issue_channels: Parameter value (Iterable[NotificationChannel] | None).
        summary_channels: Parameter value (Iterable[NotificationChannel] | None).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    detail_entries_kwargs: list[dict[str, object]] = []
    for item in rejection_items:
        issue_payload = cast(Mapping[str, object], item.get("issue"))
        raw_path = item.get("path")
        path_text = str(raw_path).strip() if isinstance(raw_path, str) else ""
        file_label = Path(path_text).name if path_text else ""
        issue_context_raw = issue_payload.get("context")
        issue_context = dict(issue_context_raw) if isinstance(issue_context_raw, Mapping) else {}
        candidate_issues = _flatten_issue_payloads(issue_payload)
        for candidate in candidate_issues:
            display_issue = cast(Mapping[str, object], candidate)
            code = str(display_issue.get("code", "VALIDATION_ERROR")).strip() or "VALIDATION_ERROR"
            translation_key = str(display_issue.get("translation_key", "")).strip()
            fallback_reason = str(display_issue.get("message", "")).strip()
            display_context_raw = display_issue.get("context")
            display_context = dict(issue_context)
            if isinstance(display_context_raw, Mapping):
                display_context.update(dict(display_context_raw))
            generic_reason = _safe_translate(typed_ns, "common.validation_failed_invalid_data")
            reason_text = _format_reason_fallback(fallback_reason, display_context)
            reason_marker: object = reason_text
            reason_context = _context_with_aliases(display_context)
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
            location_suffix = _build_validation_location_suffix(path_text=path_text, context=display_context)
            reason_text = f"{reason_text}{location_suffix}" if location_suffix else reason_text
            if isinstance(reason_marker, Mapping):
                marker_key = str(reason_marker.get("__t_key__", "")).strip()
                marker_kwargs = reason_marker.get("kwargs")
                marker_kwargs_dict = dict(marker_kwargs) if isinstance(marker_kwargs, Mapping) else {}
                reason_marker = {
                    "__t_key__": marker_key,
                    "kwargs": marker_kwargs_dict,
                    "fallback": reason_text,
                }
            else:
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
            detail_entries_kwargs.append({"file": file_text, "code": code, "reason": reason_marker})

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

    max_detail_lines = max(0, int(_VALIDATION_ACTIVITY_DETAIL_LIMIT))
    displayed_detail_count = 0
    for marker_kwargs, text in zip(detail_entries_kwargs, detail_entries_text):
        if displayed_detail_count >= max_detail_lines:
            break
        entry_payload = typed_ns["build_i18n_log_message"](
            "validation.batch.detail_entry",
            kwargs=marker_kwargs,
            fallback=text,
        )
        notify_message(
            module,
            entry_payload,
            ns=typed_ns,
            channels=("activity_log",),
        )
        displayed_detail_count += 1

    remaining_detail_count = len(detail_entries_text) - displayed_detail_count
    if remaining_detail_count > 0:
        notify_message_key(
            module,
            "validation.batch.more_suffix",
            ns=typed_ns,
            channels=("activity_log",),
            kwargs={"count": remaining_detail_count},
            fallback=f"+{remaining_detail_count} more",
        )


def emit_workbook_generation_feedback(
    module: object,
    *,
    ns: Mapping[str, object],
    success_count: int,
    failed_count: int,
    channels: Iterable[NotificationChannel] | None = ("status", "activity_log", "toast"),
) -> None:
    """Emit workbook generation feedback.
    
    Args:
        module: Parameter value (object).
        ns: Parameter value (Mapping[str, object]).
        success_count: Parameter value (int).
        failed_count: Parameter value (int).
        channels: Parameter value (Iterable[NotificationChannel] | None).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Publish status.
    
    Args:
        module: Parameter value (object).
        message: Parameter value (str).
        ns: Parameter value (Mapping[str, object]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Publish status key.
    
    Args:
        module: Parameter value (object).
        text_key: Parameter value (str).
        ns: Parameter value (Mapping[str, object]).
        kwargs: Parameter value (object).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    notify_message_key(module, text_key, ns=ns, channels=("status", "activity_log"), kwargs=kwargs)


def setup_ui_logging(module: object, *, ns: Mapping[str, object]) -> None:
    """Setup ui logging.
    
    Args:
        module: Parameter value (object).
        ns: Parameter value (Mapping[str, object]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Append user log.
    
    Args:
        module: Parameter value (object).
        message: Parameter value (str).
        ns: Parameter value (Mapping[str, object]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Rerender user log.
    
    Args:
        module: Parameter value (object).
        ns: Parameter value (Mapping[str, object]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
