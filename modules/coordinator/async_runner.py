"""Reusable async runner utilities for coordinator workflows."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, Protocol, cast

from common.jobs import CancellationToken
from common.qt_jobs import run_in_background

_JOB_REF_KEY = "job"


class _CoordinatorRunnerTarget(Protocol):
    _cancel_token: CancellationToken | None
    _active_jobs: list[object]

    def _set_busy(self, busy: bool, *, job_id: str | None = ...) -> None:
        ...


class AsyncOperationRunner:
    """Owns async operation lifecycle for full CoordinatorModule widgets."""

    def __init__(self, target: object, *, run_async: Callable[..., object] = run_in_background) -> None:
        self._target = target
        self._run_async = run_async

    def start(
        self,
        *,
        token: CancellationToken,
        job_id: str | None,
        work: Callable[[], Any],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
        on_finally: Callable[[], None] | None = None,
    ) -> None:
        target = cast(_CoordinatorRunnerTarget, self._target)
        target._cancel_token = token
        target._set_busy(True, job_id=job_id)
        job_ref: dict[str, object] = {}

        def _finalize() -> None:
            tracked_job = job_ref.get(_JOB_REF_KEY)
            if tracked_job in target._active_jobs:
                target._active_jobs.remove(tracked_job)
            target._cancel_token = None
            target._set_busy(False)
            if on_finally is not None:
                on_finally()

        def _on_finished(result: object) -> None:
            try:
                on_success(result)
            finally:
                _finalize()

        def _on_failed(exc: Exception) -> None:
            try:
                on_failure(exc)
            finally:
                _finalize()

        job = self._run_async(work, on_finished=_on_finished, on_failed=_on_failed)
        job_ref[_JOB_REF_KEY] = job
        target._active_jobs.append(job)
