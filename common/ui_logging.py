"""Shared UI logging helpers for module activity panels."""

from __future__ import annotations

import json
import logging
from datetime import datetime
from typing import Any

from common.i18n import t

_I18N_LOG_PREFIX = "__I18N_LOG__:"

class UILogHandler(logging.Handler):
    """Forward logger messages to a UI sink callback."""

    def __init__(self, sink) -> None:
        super().__init__(level=logging.INFO)
        self._sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        try:
            user_message = getattr(record, "user_message", None)
            if isinstance(user_message, str) and user_message.strip():
                message = user_message
            else:
                message = f"{record.levelname}: {record.getMessage()}"
            self._sink(message)
        except Exception:
            self.handleError(record)


def format_log_line(message: str) -> str | None:
    """Return a trimmed log line with HH:MM:SS prefix unless already timestamped."""
    return format_log_line_at(message)


def format_log_line_at(message: str, *, timestamp: datetime | None = None) -> str | None:
    """Return a trimmed log line with HH:MM:SS prefix unless already timestamped."""
    if not message or not message.strip():
        return None
    text = message.strip()
    if (
        len(text) >= 10
        and text[0] == "["
        and text[3] == ":"
        and text[6] == ":"
        and text[9] == "]"
    ):
        return text
    marker = (timestamp or datetime.now()).strftime("%H:%M:%S")
    return f"[{marker}] {text}"


def build_i18n_log_message(
    text_key: str,
    *,
    kwargs: dict[str, Any] | None = None,
    fallback: str | None = None,
) -> str:
    """Serialize an i18n-aware UI log payload as a string."""
    payload: dict[str, Any] = {"key": text_key, "kwargs": kwargs or {}}
    if fallback is not None:
        payload["fallback"] = fallback
    return _I18N_LOG_PREFIX + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def parse_i18n_log_message(message: str) -> tuple[str, dict[str, Any], str | None] | None:
    """Parse an i18n-aware UI log payload string."""
    if not isinstance(message, str) or not message.startswith(_I18N_LOG_PREFIX):
        return None
    try:
        payload = json.loads(message[len(_I18N_LOG_PREFIX):])
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict):
        return None
    key = payload.get("key")
    if not isinstance(key, str) or not key:
        return None
    raw_kwargs = payload.get("kwargs", {})
    kwargs = raw_kwargs if isinstance(raw_kwargs, dict) else {}
    fallback = payload.get("fallback")
    if fallback is not None and not isinstance(fallback, str):
        fallback = None
    return key, kwargs, fallback


def resolve_i18n_log_message(message: str) -> str:
    """Resolve an i18n-aware payload into localized text; return original for plain strings."""
    parsed = parse_i18n_log_message(message)
    if parsed is None:
        marker = message.find(_I18N_LOG_PREFIX) if isinstance(message, str) else -1
        if marker > 0:
            prefix = message[:marker]
            inner = parse_i18n_log_message(message[marker:])
            if inner is None:
                return message
            key, kwargs, fallback = inner
            try:
                return prefix + t(key, **_resolve_i18n_kwargs(kwargs))
            except Exception:
                return prefix + (fallback or message[marker:])
        return message
    key, kwargs, fallback = parsed
    try:
        return t(key, **_resolve_i18n_kwargs(kwargs))
    except Exception:
        return fallback or message


def _resolve_i18n_kwargs(kwargs: dict[str, Any]) -> dict[str, Any]:
    """Resolve nested translation-key markers inside kwargs."""
    resolved: dict[str, Any] = {}
    for key, value in kwargs.items():
        if isinstance(value, dict):
            nested_key = value.get("__t_key__")
            nested_kwargs = value.get("kwargs", {})
            if isinstance(nested_key, str) and nested_key:
                nested_safe_kwargs = nested_kwargs if isinstance(nested_kwargs, dict) else {}
                try:
                    resolved[key] = t(nested_key, **_resolve_i18n_kwargs(nested_safe_kwargs))
                except Exception:
                    resolved[key] = value.get("fallback", nested_key)
                continue
        resolved[key] = value
    return resolved

