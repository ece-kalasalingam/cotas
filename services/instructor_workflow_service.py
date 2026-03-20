"""Instructor workflow service orchestration."""

from __future__ import annotations

import logging
from inspect import Signature, signature
from pathlib import Path

from common.constants import (
    WORKFLOW_OPERATION_GENERATE_COURSE_DETAILS_TEMPLATE,
    WORKFLOW_OPERATION_GENERATE_FINAL_REPORT,
    WORKFLOW_OPERATION_GENERATE_MARKS_TEMPLATE,
    WORKFLOW_OPERATION_VALIDATE_COURSE_DETAILS_WORKBOOK,
    WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE,
)
from common.exceptions import ValidationError
from common.jobs import CancellationToken, JobContext
from domain.instructor_engine import (
    generate_course_details_template,
    generate_final_co_report,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)
from services import workflow_service_base as _workflow_base
from services.workflow_service_base import WorkflowServiceBase, WorkflowTelemetryConfig

WORKFLOW_STEP_TIMEOUT_ENV_VAR = _workflow_base.WORKFLOW_STEP_TIMEOUT_ENV_VAR
DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS = _workflow_base.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS

_logger = logging.getLogger(__name__)
_WORKFLOW_STEP_STARTED = "Instructor workflow step started."
_WORKFLOW_STEP_COMPLETED = "Instructor workflow step completed."
_WORKFLOW_STEP_CANCELLED = "Instructor workflow step cancelled."
_WORKFLOW_STEP_FAILED = "Instructor workflow step failed."
_EVENT_STEP_STARTED = "workflow_step_started"
_EVENT_STEP_COMPLETED = "workflow_step_completed"
_EVENT_STEP_CANCELLED = "workflow_step_cancelled"
_EVENT_STEP_FAILED = "workflow_step_failed"
_ERROR_VALIDATION_DEFAULT = "VALIDATION_ERROR"
_TELEMETRY = WorkflowTelemetryConfig(
    started_message=_WORKFLOW_STEP_STARTED,
    completed_message=_WORKFLOW_STEP_COMPLETED,
    cancelled_message=_WORKFLOW_STEP_CANCELLED,
    failed_message=_WORKFLOW_STEP_FAILED,
    event_step_started=_EVENT_STEP_STARTED,
    event_step_completed=_EVENT_STEP_COMPLETED,
    event_step_cancelled=_EVENT_STEP_CANCELLED,
    event_step_failed=_EVENT_STEP_FAILED,
)


class InstructorWorkflowService(WorkflowServiceBase):
    def __init__(self) -> None:
        super().__init__(logger=_logger, telemetry=_TELEMETRY)

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

    def _handle_domain_exception(
        self,
        *,
        exc: Exception,
        context: JobContext,
        operation: str,
        duration_ms: int,
    ) -> bool:
        if not isinstance(exc, ValidationError):
            return False
        self._workflow_metrics.record(operation=operation, outcome="validation_error", duration_ms=duration_ms)
        self._logger.warning(
            _WORKFLOW_STEP_FAILED,
            extra=self._build_log_extra(
                context=context,
                operation=operation,
                duration_ms=duration_ms,
                event=_EVENT_STEP_FAILED,
                error_code=str(getattr(exc, "code", _ERROR_VALIDATION_DEFAULT)),
                user_message_suffix=WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE,
            ),
        )
        return True

    @staticmethod
    def _call_with_optional_cancel_token(fn, *args: object, cancel_token: CancellationToken | None):
        try:
            fn_signature: Signature = signature(fn)
        except (TypeError, ValueError):
            fn_signature = Signature()
        if "cancel_token" in fn_signature.parameters:
            return fn(*args, cancel_token=cancel_token)
        return fn(*args)
