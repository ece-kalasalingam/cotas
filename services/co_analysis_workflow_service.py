"""CO Analysis workflow service orchestration."""

from __future__ import annotations

import logging
from pathlib import Path

from common.constants import (
    CO_ANALYSIS_WORKFLOW_OPERATION_COLLECT_FILES,
    CO_ANALYSIS_WORKFLOW_OPERATION_GENERATE_WORKBOOK,
)
from common.jobs import CancellationToken, JobContext
from domain.co_analysis_engine import (
    analyze_uploaded_workbooks,
    generate_co_analysis_workbook,
)
from services.workflow_service_base import WorkflowServiceBase, WorkflowTelemetryConfig

_logger = logging.getLogger(__name__)
_WORKFLOW_STEP_STARTED = "CO Analysis workflow step started."
_WORKFLOW_STEP_COMPLETED = "CO Analysis workflow step completed."
_WORKFLOW_STEP_CANCELLED = "CO Analysis workflow step cancelled."
_WORKFLOW_STEP_FAILED = "CO Analysis workflow step failed."
_EVENT_STEP_STARTED = "co_analysis_workflow_step_started"
_EVENT_STEP_COMPLETED = "co_analysis_workflow_step_completed"
_EVENT_STEP_CANCELLED = "co_analysis_workflow_step_cancelled"
_EVENT_STEP_FAILED = "co_analysis_workflow_step_failed"
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


class CoAnalysisWorkflowService(WorkflowServiceBase):
    def __init__(self) -> None:
        super().__init__(logger=_logger, telemetry=_TELEMETRY)

    def collect_files(
        self,
        candidate_paths: list[str],
        *,
        existing_paths: list[str],
        validate_uploaded_source_workbook,
        consume_last_source_anomaly_warnings,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, object]:
        return self._execute_with_telemetry(
            context=context,
            operation=CO_ANALYSIS_WORKFLOW_OPERATION_COLLECT_FILES,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: analyze_uploaded_workbooks(
                candidate_paths,
                existing_paths=existing_paths,
                validate_uploaded_source_workbook=validate_uploaded_source_workbook,
                consume_last_source_anomaly_warnings=consume_last_source_anomaly_warnings,
                token=effective_cancel_token,
            ),
        )

    def generate_workbook(
        self,
        source_paths: list[Path],
        output_path: Path,
        *,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        return self._execute_with_telemetry(
            context=context,
            operation=CO_ANALYSIS_WORKFLOW_OPERATION_GENERATE_WORKBOOK,
            cancel_token=cancel_token,
            work=lambda effective_cancel_token: generate_co_analysis_workbook(
                source_paths,
                output_path,
                token=effective_cancel_token,
            ),
        )
