from __future__ import annotations

from pathlib import Path

from common.constants import (
    COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT,
    COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_SOURCE,
)
from common.exceptions import JobCancelledError
from common.jobs import CancellationToken, generate_job_id


def calculate_attainment_async(module: object, *, ns: dict[str, object]) -> None:
    if module.state.busy or not module._files:
        return

    t = ns["t"]
    signature = ns["_extract_final_report_signature"](module._files[0])
    default_name = ns["_build_co_attainment_default_name"](
        module._files[0],
        section=signature.section if signature is not None else "",
    )
    save_path, _ = ns["QFileDialog"].getSaveFileName(
        module,
        t("coordinator.calculate"),
        ns["resolve_dialog_start_path"](ns["APP_NAME"], default_name),
        t("instructor.dialog.filter.excel_open"),
    )
    if not save_path:
        return

    thresholds_reader = getattr(module, "get_attainment_thresholds", None)
    thresholds = thresholds_reader() if callable(thresholds_reader) else None
    if thresholds is None:
        return

    process_name = COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT
    token = CancellationToken()
    job_id = generate_job_id()
    workflow_service = getattr(module, "_workflow_service", None)
    job_context = (
        workflow_service.create_job_context(
            step_id=COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
            payload={
                WORKFLOW_PAYLOAD_KEY_SOURCE: [str(path) for path in module._files],
                WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path,
                "thresholds": list(thresholds),
            },
        )
        if workflow_service is not None
        else None
    )
    module._publish_status_key("coordinator.status.processing_started")

    def _on_finished(result: object) -> None:
        output_path = Path(save_path)
        duplicate_reg_count = 0
        duplicate_entries: tuple[tuple[str, str, str], ...] = ()
        if isinstance(result, ns["_CoAttainmentWorkbookResult"]):
            output_path = result.output_path
            duplicate_reg_count = max(0, int(result.duplicate_reg_count))
            duplicate_entries = result.duplicate_entries
        elif result:
            output_path = Path(str(result))
        if all(ns["_path_key"](path) != ns["_path_key"](output_path) for path in module._downloaded_outputs):
            module._downloaded_outputs.append(output_path)
        module._remember_dialog_dir_safe(str(output_path))
        module._publish_status_key("coordinator.status.calculate_completed")
        threshold_summary = f"thresholds=({thresholds[0]:g},{thresholds[1]:g},{thresholds[2]:g})"
        ns["log_process_message"](
            process_name,
            logger=module._logger,
            success_message=(
                f"{process_name} completed successfully. output={output_path}, "
                f"duplicates_removed={duplicate_reg_count}, {threshold_summary}"
            ),
            user_success_message=ns["build_i18n_log_message"](
                "coordinator.status.calculate_completed",
                fallback=t("coordinator.status.calculate_completed"),
            ),
            job_id=job_context.job_id if job_context else job_id,
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
        )
        ns["show_toast"](
            module,
            t("coordinator.status.calculate_completed"),
            title=t("coordinator.title"),
            level="info",
        )
        if duplicate_reg_count:
            ns["show_toast"](
                module,
                t("coordinator.regno_dedup.body", count=duplicate_reg_count),
                title=t("coordinator.regno_dedup.title"),
                level="info",
            )
            detail_lines = [
                t(
                    "coordinator.regno_dedup.log_detail",
                    reg_no=str(reg_no),
                    worksheet=str(worksheet_name),
                    workbook=str(workbook_name),
                )
                for reg_no, worksheet_name, workbook_name in duplicate_entries
            ]
            details_text = "\n".join(detail_lines) if detail_lines else t(
                "coordinator.regno_dedup.log_detail_unavailable"
            )
            module._publish_status_key(
                "coordinator.regno_dedup.log_body",
                count=duplicate_reg_count,
                details=details_text,
            )

    def _on_failed(exc: Exception) -> None:
        if isinstance(exc, JobCancelledError):
            module._publish_status_key("coordinator.status.operation_cancelled")
            module._logger.info(
                "%s cancelled by user/system request.",
                process_name,
                extra={
                    "user_message": ns["build_i18n_log_message"](
                        "coordinator.status.operation_cancelled",
                        fallback=t("coordinator.status.operation_cancelled"),
                    ),
                    "job_id": job_context.job_id if job_context else job_id,
                    "step_id": (
                        job_context.step_id
                        if job_context
                        else COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT
                    ),
                },
            )
            return
        ns["log_process_message"](
            process_name,
            logger=module._logger,
            error=exc,
            user_error_message=ns["build_i18n_log_message"](
                "coordinator.status.processing_failed",
                fallback=t("coordinator.status.processing_failed"),
            ),
            job_id=job_context.job_id if job_context else job_id,
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
        )
        ns["show_toast"](
            module,
            t("coordinator.status.processing_failed"),
            title=t("coordinator.title"),
            level="error",
        )

    module._start_async_operation(
        token=token,
        job_id=job_context.job_id if job_context else job_id,
        work=lambda: (
            workflow_service.calculate_attainment(
                list(module._files),
                Path(save_path),
                generate_co_attainment_workbook=ns["_generate_co_attainment_workbook"],
                context=job_context,
                cancel_token=token,
            )
            if workflow_service is not None and job_context is not None
            else ns["_generate_co_attainment_workbook"](
                list(module._files),
                Path(save_path),
                token=token,
            )
        ),
        on_success=_on_finished,
        on_failure=_on_failed,
        on_finally=module._drain_next_batch,
    )
