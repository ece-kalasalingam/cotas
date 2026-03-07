"""State model for instructor workflow UI."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class InstructorWorkflowState:
    current_step: int = 1
    busy: bool = False
    active_job_id: str | None = None

    def set_busy(self, value: bool, *, job_id: str | None = None) -> None:
        self.busy = value
        self.active_job_id = job_id if value else None
