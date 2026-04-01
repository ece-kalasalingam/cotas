"""Shared async-operation lifecycle runner for UI modules."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, Protocol, cast

from common.jobs import CancellationToken
from common.qt_jobs import run_in_background

_JOB_REF_KEY = "job"


class _RunnerTarget(Protocol):
    _cancel_token: CancellationToken | None
    _active_jobs: list[object]

    def _set_busy(self, busy: bool, *, job_id: str | None = ...) -> None:
        """Set busy.
        
        Args:
            busy: Parameter value (bool).
            job_id: Parameter value (str | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        ...


class AsyncOperationRunner:
    def __init__(
        self,
        target: object,
        *,
        run_async: Callable[..., object] = run_in_background,
        refresh_ui: Callable[[], None] | None = None,
        should_refresh_ui: Callable[[], bool] | None = None,
    ) -> None:
        """Init.
        
        Args:
            target: Parameter value (object).
            run_async: Parameter value (Callable[..., object]).
            refresh_ui: Parameter value (Callable[[], None] | None).
            should_refresh_ui: Parameter value (Callable[[], bool] | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._target = target
        self._run_async = run_async
        self._refresh_ui = refresh_ui
        self._should_refresh_ui = should_refresh_ui

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
        """Start.
        
        Args:
            token: Parameter value (CancellationToken).
            job_id: Parameter value (str | None).
            work: Parameter value (Callable[[], Any]).
            on_success: Parameter value (Callable[[object], None]).
            on_failure: Parameter value (Callable[[Exception], None]).
            on_finally: Parameter value (Callable[[], None] | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        target = cast(_RunnerTarget, self._target)
        target._cancel_token = token
        target._set_busy(True, job_id=job_id)
        self._notify_global_processing_state(True)
        job_ref: dict[str, object] = {}

        def _finalize() -> None:
            """Finalize.
            
            Args:
                None.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            tracked_job = job_ref.get(_JOB_REF_KEY)
            if tracked_job in target._active_jobs:
                target._active_jobs.remove(tracked_job)
            target._cancel_token = None
            target._set_busy(False)
            self._notify_global_processing_state(False)
            if on_finally is not None:
                on_finally()
            if self._refresh_ui is not None:
                if self._should_refresh_ui is None or self._should_refresh_ui():
                    self._refresh_ui()

        def _on_finished(result: object) -> None:
            """On finished.
            
            Args:
                result: Parameter value (object).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            try:
                on_success(result)
            finally:
                _finalize()

        def _on_failed(exc: Exception) -> None:
            """On failed.
            
            Args:
                exc: Parameter value (Exception).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            try:
                on_failure(exc)
            finally:
                _finalize()

        job = self._run_async(work, on_finished=_on_finished, on_failed=_on_failed)
        job_ref[_JOB_REF_KEY] = job
        target._active_jobs.append(job)

    def _notify_global_processing_state(self, active: bool) -> None:
        """Notify global processing state.
        
        Args:
            active: Parameter value (bool).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        host_candidates: list[object] = [self._target]
        window_getter = getattr(self._target, "window", None)
        if callable(window_getter):
            try:
                window_obj = window_getter()
            except Exception:
                window_obj = None
            if window_obj is not None:
                host_candidates.append(window_obj)

        for candidate in host_candidates:
            setter = getattr(candidate, "set_global_processing_active", None)
            if callable(setter):
                try:
                    setter(active)
                except Exception:
                    continue
                break
