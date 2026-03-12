"""Shared UI logging helpers for module activity panels."""

from __future__ import annotations

import logging
from datetime import datetime


class UILogHandler(logging.Handler):
    """Forward logger messages to a UI sink callback."""

    def __init__(self, sink) -> None:
        super().__init__(level=logging.INFO)
        self._sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        try:
            user_message = getattr(record, "user_message", None)
            message = f"{record.levelname}: {user_message or record.getMessage()}"
            self._sink(message)
        except Exception:
            self.handleError(record)


def format_log_line(message: str) -> str | None:
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
    return f"[{datetime.now().strftime('%H:%M:%S')}] {text}"
