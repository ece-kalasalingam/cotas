"""Coordinator workflow service orchestration."""

from __future__ import annotations

import logging
import os
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import TimeoutError as FuturesTimeoutError
from pathlib import Path
from typing import Any, Callable, Mapping, TypeVar

from common.constants import (
    COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT,
    COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES,
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

_logger = logging.getLogger(__name__)
_T = TypeVar("_T")
DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS = 120
_WORKFLOW_STEP_STARTED = "Coordinator workflow step started."
_WORKFLOW_STEP_COMPLETED = "Coordinator workflow step completed."
_WORKFLOW_STEP_CANCELLED = "Coordinator workflow step cancelled."
_WORKFLOW_STEP_FAILED = "Coordinator workflow step failed."
_EVENT_STEP_STARTED = "coordinator_workflow_step_started"
_EVENT_STEP_COMPLETED = "coordinator_workflow_step_completed"
_EVENT_STEP_CANCELLED = "coordinator_workflow_step_cancelled"
_EVENT_STEP_FAILED = "coordinator_workflow_step_failed"
_ERROR_NONE = "NONE"
_ERROR_JOB_CANCELLED = "JOB_CANCELLED"
_ERROR_APP_SYSTEM = "APP_SYSTEM_ERROR"
_ERROR_UNEXPECTED = "UNEXPECTED_ERROR"


class WorkflowMetrics:
    """In-memory workflow metrics snapshot for coordinator observability."""

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


_workflow_metrics = WorkflowMetrics()


class CoordinatorWorkflowService:
    def create_job_context(self, *, step_id: str, payload: Mapping[str, Any] | None = None) -> JobContext:
        return JobContext.create(step_id=step_id, payload=payload)

    def collect_files(
        self,
        dropped_files: list[str],
        *,
        existing_keys: set[str],
        existing_paths: list[str],
        analyze_dropped_files: Callable[..., dict[str, object]],
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, object]:
        return self._execute_with_telemetry(
            context=context,
            operation=COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: analyze_dropped_files(
                dropped_files,
                existing_keys=existing_keys,
                existing_paths=existing_paths,
                token=effective_cancel_token,
            ),
        )

    def calculate_attainment(
        self,
        source_paths: list[Path],
        output_path: Path,
        *,
        generate_co_attainment_workbook: Callable[..., object],
        context: JobContext,
        cancel_token: CancellationToken | None = None,
        thresholds: tuple[float, float, float] | None = None,
    ):
        return self._execute_with_telemetry(
            context=context,
            operation=COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: (
                generate_co_attainment_workbook(
                    source_paths,
                    output_path,
                    token=effective_cancel_token,
                )
                if thresholds is None
                else generate_co_attainment_workbook(
                    source_paths,
                    output_path,
                    token=effective_cancel_token,
                    thresholds=thresholds,
                )
            ),
        )

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
        _logger.info(
            _WORKFLOW_STEP_STARTED,
            extra={
                LOG_EXTRA_KEY_USER_MESSAGE: f"{operation}{WORKFLOW_USER_MESSAGE_STARTED_SUFFIX}",
                "event": _EVENT_STEP_STARTED,
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
            _workflow_metrics.record(operation=operation, outcome="success", duration_ms=duration_ms)
            _logger.info(
                _WORKFLOW_STEP_COMPLETED,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: (
                        f"{operation}{WORKFLOW_USER_MESSAGE_COMPLETED_TEMPLATE.format(duration_ms=duration_ms)}"
                    ),
                    "event": _EVENT_STEP_COMPLETED,
                    "error_code": _ERROR_NONE,
                    "operation": operation,
                    "duration_ms": duration_ms,
                    "metrics": _workflow_metrics.snapshot(),
                    LOG_EXTRA_KEY_JOB_ID: context.job_id,
                    LOG_EXTRA_KEY_STEP_ID: context.step_id,
                },
            )
            return result
        except JobCancelledError:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            _workflow_metrics.record(operation=operation, outcome="cancelled", duration_ms=duration_ms)
            _logger.info(
                _WORKFLOW_STEP_CANCELLED,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: (
                        f"{operation}{WORKFLOW_USER_MESSAGE_CANCELLED_TEMPLATE.format(duration_ms=duration_ms)}"
                    ),
                    "event": _EVENT_STEP_CANCELLED,
                    "error_code": _ERROR_JOB_CANCELLED,
                    "operation": operation,
                    "duration_ms": duration_ms,
                    "metrics": _workflow_metrics.snapshot(),
                    LOG_EXTRA_KEY_JOB_ID: context.job_id,
                    LOG_EXTRA_KEY_STEP_ID: context.step_id,
                },
            )
            raise
        except AppSystemError:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            _workflow_metrics.record(operation=operation, outcome="system_error", duration_ms=duration_ms)
            _logger.error(
                _WORKFLOW_STEP_FAILED,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: (
                        f"{operation}{WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE.format(duration_ms=duration_ms)}"
                    ),
                    "event": _EVENT_STEP_FAILED,
                    "error_code": _ERROR_APP_SYSTEM,
                    "operation": operation,
                    "duration_ms": duration_ms,
                    "metrics": _workflow_metrics.snapshot(),
                    LOG_EXTRA_KEY_JOB_ID: context.job_id,
                    LOG_EXTRA_KEY_STEP_ID: context.step_id,
                },
            )
            raise
        except Exception as exc:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            _workflow_metrics.record(operation=operation, outcome="unexpected_error", duration_ms=duration_ms)
            _logger.exception(
                _WORKFLOW_STEP_FAILED,
                exc_info=exc,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: (
                        f"{operation}{WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE.format(duration_ms=duration_ms)}"
                    ),
                    "event": _EVENT_STEP_FAILED,
                    "error_code": _ERROR_UNEXPECTED,
                    "operation": operation,
                    "duration_ms": duration_ms,
                    "metrics": _workflow_metrics.snapshot(),
                    LOG_EXTRA_KEY_JOB_ID: context.job_id,
                    LOG_EXTRA_KEY_STEP_ID: context.step_id,
                },
            )
            raise

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
            executor.shutdown(wait=not timed_out, cancel_futures=timed_out)
