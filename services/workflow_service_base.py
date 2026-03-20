"""Shared workflow service telemetry and execution helpers."""

from __future__ import annotations

import os
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import TimeoutError as FuturesTimeoutError
from dataclasses import dataclass
from typing import Any, Callable, Mapping, TypeVar

from common.constants import (
    LOG_EXTRA_KEY_JOB_ID,
    LOG_EXTRA_KEY_STEP_ID,
    LOG_EXTRA_KEY_USER_MESSAGE,
    WORKFLOW_STEP_TIMEOUT_ENV_VAR,
    WORKFLOW_TIMEOUT_ERROR_TEMPLATE,
    WORKFLOW_USER_MESSAGE_CANCELLED_TEMPLATE,
    WORKFLOW_USER_MESSAGE_COMPLETED_TEMPLATE,
    WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE,
    WORKFLOW_USER_MESSAGE_STARTED_SUFFIX,
)
from common.exceptions import AppSystemError, JobCancelledError
from common.jobs import CancellationToken, JobContext

_T = TypeVar("_T")
DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS = 120
_ERROR_NONE = "NONE"
_ERROR_JOB_CANCELLED = "JOB_CANCELLED"
_ERROR_APP_SYSTEM = "APP_SYSTEM_ERROR"
_ERROR_UNEXPECTED = "UNEXPECTED_ERROR"


@dataclass(frozen=True)
class WorkflowTelemetryConfig:
    started_message: str
    completed_message: str
    cancelled_message: str
    failed_message: str
    event_step_started: str
    event_step_completed: str
    event_step_cancelled: str
    event_step_failed: str


class WorkflowMetrics:
    """In-memory workflow metrics snapshot for service observability."""

    def __init__(self) -> None:
        self._counts: Counter[str] = Counter()
        self._durations_ms: dict[str, list[int]] = defaultdict(list)

    def record(self, *, operation: str, outcome: str, duration_ms: int) -> None:
        self._counts[f"{operation}.{outcome}"] += 1
        self._durations_ms[operation].append(duration_ms)

    def snapshot(self) -> dict[str, Any]:
        return {
            "counts": dict(self._counts),
            "durations_ms": {key: list(values) for key, values in self._durations_ms.items()},
        }


class WorkflowServiceBase:
    def __init__(self, *, logger, telemetry: WorkflowTelemetryConfig) -> None:
        self._logger = logger
        self._telemetry = telemetry
        self._workflow_metrics = WorkflowMetrics()

    def create_job_context(self, *, step_id: str, payload: Mapping[str, Any] | None = None) -> JobContext:
        return JobContext.create(step_id=step_id, payload=payload)

    @staticmethod
    def _raise_if_cancelled(cancel_token: CancellationToken | None) -> None:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()

    def _execute_with_telemetry(
        self,
        *,
        context: JobContext,
        operation: str,
        cancel_token: CancellationToken | None,
        work: Callable[[CancellationToken], _T],
    ) -> _T:
        effective_cancel_token = cancel_token or CancellationToken()
        started_at = time.perf_counter()
        timeout_seconds = self._resolve_timeout_seconds()
        self._logger.info(
            self._telemetry.started_message,
            extra={
                LOG_EXTRA_KEY_USER_MESSAGE: f"{operation}{WORKFLOW_USER_MESSAGE_STARTED_SUFFIX}",
                "event": self._telemetry.event_step_started,
                "error_code": _ERROR_NONE,
                "operation": operation,
                "timeout_seconds": timeout_seconds,
                LOG_EXTRA_KEY_JOB_ID: context.job_id,
                LOG_EXTRA_KEY_STEP_ID: context.step_id,
            },
        )
        try:
            self._raise_if_cancelled(effective_cancel_token)
            result = self._run_with_timeout(
                operation=operation,
                work=lambda: work(effective_cancel_token),
                timeout_seconds=timeout_seconds,
                cancel_token=effective_cancel_token,
            )
            self._raise_if_cancelled(effective_cancel_token)
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            self._record_and_log(
                context=context,
                operation=operation,
                outcome="success",
                duration_ms=duration_ms,
                event=self._telemetry.event_step_completed,
                error_code=_ERROR_NONE,
                level="info",
                message=self._telemetry.completed_message,
                user_message_suffix=WORKFLOW_USER_MESSAGE_COMPLETED_TEMPLATE,
            )
            return result
        except JobCancelledError:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            self._record_and_log(
                context=context,
                operation=operation,
                outcome="cancelled",
                duration_ms=duration_ms,
                event=self._telemetry.event_step_cancelled,
                error_code=_ERROR_JOB_CANCELLED,
                level="info",
                message=self._telemetry.cancelled_message,
                user_message_suffix=WORKFLOW_USER_MESSAGE_CANCELLED_TEMPLATE,
            )
            raise
        except AppSystemError:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            self._record_and_log(
                context=context,
                operation=operation,
                outcome="system_error",
                duration_ms=duration_ms,
                event=self._telemetry.event_step_failed,
                error_code=_ERROR_APP_SYSTEM,
                level="error",
                message=self._telemetry.failed_message,
                user_message_suffix=WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE,
            )
            raise
        except Exception as exc:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            if self._handle_domain_exception(
                exc=exc,
                context=context,
                operation=operation,
                duration_ms=duration_ms,
            ):
                raise
            self._workflow_metrics.record(operation=operation, outcome="unexpected_error", duration_ms=duration_ms)
            self._logger.exception(
                self._telemetry.failed_message,
                exc_info=exc,
                extra=self._build_log_extra(
                    context=context,
                    operation=operation,
                    duration_ms=duration_ms,
                    event=self._telemetry.event_step_failed,
                    error_code=_ERROR_UNEXPECTED,
                    user_message_suffix=WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE,
                ),
            )
            raise

    def _handle_domain_exception(
        self,
        *,
        exc: Exception,
        context: JobContext,
        operation: str,
        duration_ms: int,
    ) -> bool:
        return False

    def _record_and_log(
        self,
        *,
        context: JobContext,
        operation: str,
        outcome: str,
        duration_ms: int,
        event: str,
        error_code: str,
        level: str,
        message: str,
        user_message_suffix: str,
    ) -> None:
        self._workflow_metrics.record(operation=operation, outcome=outcome, duration_ms=duration_ms)
        log_method = getattr(self._logger, level)
        log_method(
            message,
            extra=self._build_log_extra(
                context=context,
                operation=operation,
                duration_ms=duration_ms,
                event=event,
                error_code=error_code,
                user_message_suffix=user_message_suffix,
            ),
        )

    def _build_log_extra(
        self,
        *,
        context: JobContext,
        operation: str,
        duration_ms: int,
        event: str,
        error_code: str,
        user_message_suffix: str,
    ) -> dict[str, Any]:
        return {
            LOG_EXTRA_KEY_USER_MESSAGE: f"{operation}{user_message_suffix.format(duration_ms=duration_ms)}",
            "event": event,
            "error_code": error_code,
            "operation": operation,
            "duration_ms": duration_ms,
            "metrics": self._workflow_metrics.snapshot(),
            LOG_EXTRA_KEY_JOB_ID: context.job_id,
            LOG_EXTRA_KEY_STEP_ID: context.step_id,
        }

    @staticmethod
    def _resolve_timeout_seconds() -> int:
        raw = os.getenv(WORKFLOW_STEP_TIMEOUT_ENV_VAR, "").strip()
        if not raw:
            return DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS
        try:
            parsed = int(raw)
        except ValueError:
            return DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS
        return parsed if parsed > 0 else DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS

    @staticmethod
    def _run_with_timeout(
        *,
        operation: str,
        work: Callable[[], _T],
        timeout_seconds: int,
        cancel_token: CancellationToken | None,
    ) -> _T:
        executor = ThreadPoolExecutor(max_workers=1)
        future = executor.submit(work)
        timed_out = False
        try:
            return future.result(timeout=timeout_seconds)
        except FuturesTimeoutError as exc:
            timed_out = True
            future.cancel()
            if cancel_token is not None:
                cancel_token.cancel()
            raise AppSystemError(
                WORKFLOW_TIMEOUT_ERROR_TEMPLATE.format(
                    operation=operation,
                    timeout_seconds=timeout_seconds,
                )
            ) from exc
        finally:
            # Avoid blocking on timeout; the worker thread may continue in background.
            executor.shutdown(wait=not timed_out, cancel_futures=timed_out)
