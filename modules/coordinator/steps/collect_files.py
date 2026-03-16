"""Coordinator step: validate and collect dropped final report files."""

from __future__ import annotations

from pathlib import Path

from common.constants import (
    COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES,
    COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
    WORKFLOW_PAYLOAD_KEY_PATH,
)
from common.exceptions import JobCancelledError
from common.jobs import CancellationToken, generate_job_id
from PySide6.QtCore import Qt


def process_files_async(module: object, dropped_files: list[str], *, ns: dict[str, object]) -> None:
    if not dropped_files:
        return
    if module.state.busy:
        module._pending_drop_batches.append(dropped_files)
        module._publish_status_key("coordinator.status.queued", count=len(dropped_files))
        return

    t = ns["t"]
    process_name = COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES
    token = CancellationToken()
    job_id = generate_job_id()
    existing_keys = {ns["_path_key"](path) for path in module._files}
    existing_paths = [str(path) for path in module._files]
    workflow_service = getattr(module, "_workflow_service", None)
    job_context = (
        workflow_service.create_job_context(
            step_id=COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: list(dropped_files)},
        )
        if workflow_service is not None
        else None
    )
    module._publish_status_key("coordinator.status.processing_started")

    def _on_finished(result: object) -> None:
        if not isinstance(result, dict):
            raise RuntimeError("Coordinator processing returned unexpected result type.")
        added_paths = [Path(value) for value in result.get("added", [])]
        duplicates = int(result.get("duplicates", 0))
        invalid_paths = [Path(value) for value in result.get("invalid_final_report", [])]
        ignored = int(result.get("ignored", 0))

        module._add_uploaded_paths(added_paths)

        if added_paths:
            module._publish_status_key(
                "coordinator.status.added",
                added=len(added_paths),
                total=len(module._files),
            )
        if duplicates:
            ns["show_toast"](
                module,
                t("coordinator.duplicate.body", count=duplicates),
                title=t("coordinator.duplicate.title"),
                level="info",
            )
        if invalid_paths:
            file_names = "\n".join(path.name for path in invalid_paths)
            ns["show_toast"](
                module,
                t(
                    "coordinator.invalid_final_report.body",
                    count=len(invalid_paths),
                    files=file_names,
                ),
                title=t("coordinator.invalid_final_report.title"),
                level="warning",
            )
        if ignored:
            module._publish_status_key("coordinator.status.ignored", count=ignored)

        ns["log_process_message"](
            process_name,
            logger=module._logger,
            success_message=(
                f"{process_name} completed successfully. "
                f"added={len(added_paths)}, duplicates={duplicates}, "
                f"invalid={len(invalid_paths)}, ignored={ignored}"
            ),
            user_success_message=ns["build_i18n_log_message"](
                "coordinator.status.processing_completed",
                fallback=t("coordinator.status.processing_completed"),
            ),
            job_id=job_context.job_id if job_context else job_id,
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
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
                    "step_id": job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
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
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
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
            workflow_service.collect_files(
                dropped_files,
                existing_keys=existing_keys,
                existing_paths=existing_paths,
                analyze_dropped_files=ns["_analyze_dropped_files"],
                context=job_context,
                cancel_token=token,
            )
            if workflow_service is not None and job_context is not None
            else ns["_analyze_dropped_files"](
                dropped_files,
                existing_keys=existing_keys,
                existing_paths=existing_paths,
                token=token,
            )
        ),
        on_success=_on_finished,
        on_failure=_on_failed,
        on_finally=module._drain_next_batch,
    )


def add_uploaded_paths(module: object, added_paths: list[Path], *, ns: dict[str, object]) -> None:
    for path in added_paths:
        module._files.append(path)
        item = ns["QListWidgetItem"]()
        item.setToolTip(str(path))
        item.setData(Qt.ItemDataRole.UserRole, str(path))
        module.drop_list.addItem(item)
        row_widget = module._new_file_item_widget(str(path), parent=module.drop_list)
        row_widget.removed.connect(module._remove_file_by_path)
        item.setSizeHint(row_widget.sizeHint())
        module.drop_list.setItemWidget(item, row_widget)
