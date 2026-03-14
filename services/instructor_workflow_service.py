"""Instructor workflow service orchestration."""

from __future__ import annotations

import logging
import os
import time
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
from inspect import Signature, signature
from pathlib import Path
from typing import Any, Callable, Mapping, TypeVar

from common.constants import (
    LOG_EXTRA_KEY_JOB_ID,
    LOG_EXTRA_KEY_STEP_ID,
    LOG_EXTRA_KEY_USER_MESSAGE,
    WORKFLOW_OPERATION_GENERATE_COURSE_DETAILS_TEMPLATE,
    WORKFLOW_OPERATION_GENERATE_FINAL_REPORT,
    WORKFLOW_OPERATION_GENERATE_MARKS_TEMPLATE,
    WORKFLOW_OPERATION_VALIDATE_COURSE_DETAILS_WORKBOOK,
    WORKFLOW_TIMEOUT_ERROR_TEMPLATE,
    WORKFLOW_USER_MESSAGE_CANCELLED_TEMPLATE,
    WORKFLOW_USER_MESSAGE_COMPLETED_TEMPLATE,
    WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE,
    WORKFLOW_USER_MESSAGE_STARTED_SUFFIX,
    WORKFLOW_STEP_TIMEOUT_ENV_VAR,
)
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.jobs import CancellationToken, JobContext
from domain.instructor_engine import (
    generate_course_details_template,
    generate_final_co_report,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)

_logger = logging.getLogger(__name__)
_T = TypeVar("_T")
DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS = 120
_WORKFLOW_STEP_STARTED = "Instructor workflow step started."
_WORKFLOW_STEP_COMPLETED = "Instructor workflow step completed."
_WORKFLOW_STEP_CANCELLED = "Instructor workflow step cancelled."
_WORKFLOW_STEP_FAILED = "Instructor workflow step failed."
_EVENT_STEP_STARTED = "workflow_step_started"
_EVENT_STEP_COMPLETED = "workflow_step_completed"
_EVENT_STEP_CANCELLED = "workflow_step_cancelled"
_EVENT_STEP_FAILED = "workflow_step_failed"
_ERROR_NONE = "NONE"
_ERROR_JOB_CANCELLED = "JOB_CANCELLED"
_ERROR_APP_SYSTEM = "APP_SYSTEM_ERROR"
_ERROR_UNEXPECTED = "UNEXPECTED_ERROR"
_ERROR_VALIDATION_DEFAULT = "VALIDATION_ERROR"


class WorkflowMetrics:
    """In-memory workflow metrics snapshot for operational observability."""

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


class InstructorWorkflowService:
    def create_job_context(self, *, step_id: str, payload: Mapping[str, Any] | None = None) -> JobContext:
        return JobContext.create(step_id=step_id, payload=payload)

    def generate_course_details_template(
        self,
        output_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        return self._execute_with_telemetry(
            context=context,
            operation=WORKFLOW_OPERATION_GENERATE_COURSE_DETAILS_TEMPLATE,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: self._call_with_optional_cancel_token(
                generate_course_details_template,
                output_path,
                cancel_token=effective_cancel_token,
            ),
        )

    def validate_course_details_workbook(
        self,
        workbook_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> str:
        return self._execute_with_telemetry(
            context=context,
            operation=WORKFLOW_OPERATION_VALIDATE_COURSE_DETAILS_WORKBOOK,
            cancel_token=cancel_token,
            work=lambda _effective_cancel_token: validate_course_details_workbook(workbook_path),
        )

    def generate_marks_template(
        self,
        course_details_path: str | Path,
        output_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        return self._execute_with_telemetry(
            context=context,
            operation=WORKFLOW_OPERATION_GENERATE_MARKS_TEMPLATE,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: self._call_with_optional_cancel_token(
                generate_marks_template_from_course_details,
                course_details_path,
                output_path,
                cancel_token=effective_cancel_token,
            ),
        )

    def generate_final_report(
        self,
        filled_marks_path: str | Path,
        output_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        return self._execute_with_telemetry(
            context=context,
            operation=WORKFLOW_OPERATION_GENERATE_FINAL_REPORT,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: self._call_with_optional_cancel_token(
                generate_final_co_report,
                filled_marks_path,
                output_path,
                cancel_token=effective_cancel_token,
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
        except ValidationError as exc:
            duration_ms = int((time.perf_counter() - started_at) * 1000)
            _workflow_metrics.record(operation=operation, outcome="validation_error", duration_ms=duration_ms)
            _logger.warning(
                _WORKFLOW_STEP_FAILED,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: (
                        f"{operation}{WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE.format(duration_ms=duration_ms)}"
                    ),
                    "event": _EVENT_STEP_FAILED,
                    "error_code": str(getattr(exc, "code", _ERROR_VALIDATION_DEFAULT)),
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
            # Avoid blocking on timeout; the worker thread may continue in background.
            executor.shutdown(wait=not timed_out, cancel_futures=timed_out)

    @staticmethod
    def _call_with_optional_cancel_token(fn, *args: object, cancel_token: CancellationToken | None):
        try:
            fn_signature: Signature = signature(fn)
        except (TypeError, ValueError):
            fn_signature = Signature()
        if "cancel_token" in fn_signature.parameters:
            return fn(*args, cancel_token=cancel_token)
        return fn(*args)
