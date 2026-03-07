"""Job context and cancellation primitives."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from threading import Event
from typing import Any, Mapping
from uuid import uuid4

from common.exceptions import JobCancelledError
from common.texts import get_language


def generate_job_id() -> str:
    return uuid4().hex[:12]


@dataclass(frozen=True)
class JobContext:
    job_id: str
    step_id: str
    language: str
    created_at_utc: str
    payload: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def create(cls, *, step_id: str, payload: Mapping[str, Any] | None = None) -> "JobContext":
        return cls(
            job_id=generate_job_id(),
            step_id=step_id,
            language=get_language(),
            created_at_utc=datetime.now(timezone.utc).isoformat(),
            payload=dict(payload or {}),
        )


class CancellationToken:
    def __init__(self) -> None:
        self._event = Event()

    def cancel(self) -> None:
        self._event.set()

    @property
    def cancelled(self) -> bool:
        return self._event.is_set()

    def raise_if_cancelled(self, *, message: str = "Job was cancelled.") -> None:
        if self.cancelled:
            raise JobCancelledError(message)
