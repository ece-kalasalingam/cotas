"""Shared workflow state models."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class BusyWorkflowState:
    busy: bool = False
    active_job_id: str | None = None

    def set_busy(self, value: bool, *, job_id: str | None = None) -> None:
        self.busy = value
        self.active_job_id = job_id if value else None
