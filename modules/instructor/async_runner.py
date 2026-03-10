"""Reusable async runner utilities for instructor workflows."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from common.jobs import CancellationToken
from common.qt_jobs import run_in_background
from common.utils import emit_user_status

_ATTR_PUBLISH_STATUS = "_publish_status"
_ATTR_SET_BUSY = "_set_busy"
_ATTR_START_ASYNC_OPERATION = "_start_async_operation"
_ATTR_CANCEL_TOKEN = "_cancel_token"
_ATTR_ACTIVE_JOBS = "_active_jobs"
_ATTR_REFRESH_UI = "_refresh_ui"
_JOB_REF_KEY = "job"


def publish_status_compat(*, target: object, message: str, logger: object) -> None:
    publish = getattr(target, _ATTR_PUBLISH_STATUS, None)
    if callable(publish):
        publish(message)
        return
    emit_user_status(getattr(target, "status_changed", None), message, logger=logger)


def set_busy_compat(*, target: object, busy: bool, job_id: str | None = None) -> None:
    setter = getattr(target, _ATTR_SET_BUSY, None)
    if callable(setter):
        setter(busy, job_id=job_id)


def start_async_operation_compat(
    *,
    target: object,
    token: CancellationToken,
    job_id: str | None,
    work: Callable[[], Any],
    on_success: Callable[[object], None],
    on_failure: Callable[[Exception], None],
    run_async: Callable[..., object] = run_in_background,
) -> None:
    starter = getattr(target, _ATTR_START_ASYNC_OPERATION, None)
    if callable(starter):
        starter(
            token=token,
            job_id=job_id,
            work=work,
            on_success=on_success,
            on_failure=on_failure,
        )
        return

    setattr(target, _ATTR_CANCEL_TOKEN, token)
    set_busy_compat(target=target, busy=True, job_id=job_id)
    job_ref: dict[str, object] = {}

    def _finalize() -> None:
        active_jobs = getattr(target, _ATTR_ACTIVE_JOBS, None)
        tracked_job = job_ref.get(_JOB_REF_KEY)
        if isinstance(active_jobs, list) and tracked_job in active_jobs:
            active_jobs.remove(tracked_job)
        setattr(target, _ATTR_CANCEL_TOKEN, None)
        set_busy_compat(target=target, busy=False)
        refresh = getattr(target, _ATTR_REFRESH_UI, None)
        if callable(refresh):
            refresh()

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

    job = run_async(work, on_finished=_on_finished, on_failed=_on_failed)
    job_ref[_JOB_REF_KEY] = job
    active_jobs = getattr(target, _ATTR_ACTIVE_JOBS, None)
    if isinstance(active_jobs, list):
        active_jobs.append(job)


class AsyncOperationRunner:
    """Owns async operation lifecycle for full InstructorModule widgets."""

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
    ) -> None:
        target = self._target
        target._cancel_token = token
        target._set_busy(True, job_id=job_id)
        job_ref: dict[str, object] = {}

        def _finalize() -> None:
            tracked_job = job_ref.get(_JOB_REF_KEY)
            if tracked_job in target._active_jobs:
                target._active_jobs.remove(tracked_job)
            target._cancel_token = None
            target._set_busy(False)
            if not target._is_closing:
                target._refresh_ui()

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
