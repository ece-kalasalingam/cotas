"""Instructor workflow service orchestration."""

from __future__ import annotations

import logging
from pathlib import Path

from common.constants import (
    ID_COURSE_SETUP,
    WORKFLOW_OPERATION_GENERATE_COURSE_DETAILS_TEMPLATE,
    WORKFLOW_OPERATION_GENERATE_FINAL_REPORT,
    WORKFLOW_OPERATION_GENERATE_MARKS_TEMPLATE,
    WORKFLOW_OPERATION_VALIDATE_COURSE_DETAILS_WORKBOOK,
    WORKFLOW_USER_MESSAGE_FAILED_TEMPLATE,
)
from common.error_catalog import validation_error_from_key
from common.exceptions import ValidationError
from common.jobs import CancellationToken, JobContext
from domain.template_strategy_router import (
    generate_workbook,
    read_valid_template_id_from_system_hash_sheet,
    validate_workbook,
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
            work=lambda effective_cancel_token: self._result_to_path(
                generate_workbook(
                    template_id=ID_COURSE_SETUP,
                    output_path=output_path,
                    workbook_name=Path(output_path).name,
                    workbook_kind="course_details_template",
                    cancel_token=effective_cancel_token,
                ),
                fallback=Path(output_path),
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
            work=lambda effective_cancel_token: validate_workbook(
                workbook_path=workbook_path,
                workbook_kind="course_details",
                cancel_token=effective_cancel_token,
            ),
        )

    def generate_marks_template(
        self,
        course_details_path: str | Path,
        output_path: str | Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        template_id = self._resolve_template_id_from_workbook(course_details_path)
        return self._execute_with_telemetry(
            context=context,
            operation=WORKFLOW_OPERATION_GENERATE_MARKS_TEMPLATE,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: self._result_to_path(
                generate_workbook(
                    template_id=template_id,
                    output_path=output_path,
                    workbook_name=Path(output_path).name,
                    workbook_kind="marks_template",
                    cancel_token=effective_cancel_token,
                    context={"course_details_path": str(course_details_path)},
                ),
                fallback=Path(output_path),
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
        template_id = self._resolve_template_id_from_workbook(filled_marks_path)
        return self._execute_with_telemetry(
            context=context,
            operation=WORKFLOW_OPERATION_GENERATE_FINAL_REPORT,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: self._result_to_path(
                generate_workbook(
                    template_id=template_id,
                    output_path=output_path,
                    workbook_name=Path(output_path).name,
                    workbook_kind="final_report",
                    cancel_token=effective_cancel_token,
                    context={"filled_marks_path": str(filled_marks_path)},
                ),
                fallback=Path(output_path),
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
    def _resolve_template_id_from_workbook(workbook_path: str | Path) -> str:
        try:
            import openpyxl
        except ModuleNotFoundError as exc:
            raise validation_error_from_key(
                "validation.dependency.openpyxl_missing",
                code="OPENPYXL_MISSING",
            ) from exc
        source = Path(workbook_path)
        try:
            workbook = openpyxl.load_workbook(source, data_only=False, read_only=True)
        except Exception as exc:
            raise validation_error_from_key(
                "validation.workbook.open_failed",
                code="WORKBOOK_OPEN_FAILED",
                workbook=str(source),
            ) from exc
        try:
            return read_valid_template_id_from_system_hash_sheet(workbook)
        finally:
            workbook.close()

    @staticmethod
    def _result_to_path(result: object, *, fallback: Path) -> Path:
        if isinstance(result, Path):
            return result
        output = getattr(result, "output_path", None)
        if isinstance(output, Path):
            return output
        if isinstance(output, str) and output.strip():
            return Path(output)
        return fallback
