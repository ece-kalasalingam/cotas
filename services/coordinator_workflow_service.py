"""Coordinator workflow service orchestration."""

from __future__ import annotations

import logging
from pathlib import Path

from common.constants import (
    COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT,
    COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES,
)
from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken, JobContext
from domain.template_versions.course_setup_v1_coordinator_engine import (
    _analyze_dropped_files,
    _generate_co_attainment_workbook_course_setup_v1,
    extract_final_report_signature_from_path,
)
from services import workflow_service_base as _workflow_base
from services.workflow_service_base import WorkflowServiceBase, WorkflowTelemetryConfig

WORKFLOW_STEP_TIMEOUT_ENV_VAR = _workflow_base.WORKFLOW_STEP_TIMEOUT_ENV_VAR
DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS = _workflow_base.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS

_logger = logging.getLogger(__name__)
_WORKFLOW_STEP_STARTED = "Coordinator workflow step started."
_WORKFLOW_STEP_COMPLETED = "Coordinator workflow step completed."
_WORKFLOW_STEP_CANCELLED = "Coordinator workflow step cancelled."
_WORKFLOW_STEP_FAILED = "Coordinator workflow step failed."
_EVENT_STEP_STARTED = "coordinator_workflow_step_started"
_EVENT_STEP_COMPLETED = "coordinator_workflow_step_completed"
_EVENT_STEP_CANCELLED = "coordinator_workflow_step_cancelled"
_EVENT_STEP_FAILED = "coordinator_workflow_step_failed"

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


class CoordinatorWorkflowService(WorkflowServiceBase):
    def __init__(self) -> None:
        super().__init__(logger=_logger, telemetry=_TELEMETRY)

    def collect_files(
        self,
        dropped_files: list[str],
        *,
        existing_keys: set[str],
        existing_paths: list[str],
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, object]:
        return self._execute_with_telemetry(
            context=context,
            operation=COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: _analyze_dropped_files(
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
        context: JobContext,
        cancel_token: CancellationToken | None = None,
        thresholds: tuple[float, float, float] | None = None,
        co_attainment_percent: float | None = None,
        co_attainment_level: int | None = None,
    ):
        if not source_paths:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="COA_SOURCE_WORKBOOK_REQUIRED",
            )
        signature = extract_final_report_signature_from_path(source_paths[0])
        if signature is None:
            raise validation_error_from_key(
                "validation.workbook.open_failed",
                code="WORKBOOK_OPEN_FAILED",
                workbook=str(source_paths[0]),
            )
        return self._execute_with_telemetry(
            context=context,
            operation=COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: _generate_co_attainment_workbook_course_setup_v1(
                source_paths=source_paths,
                output_path=output_path,
                token=effective_cancel_token,
                total_outcomes=signature.total_outcomes,
                template_id=signature.template_id,
                thresholds=tuple(thresholds) if thresholds is not None else None,
                co_attainment_percent=co_attainment_percent,
                co_attainment_level=co_attainment_level,
            ),
        )
