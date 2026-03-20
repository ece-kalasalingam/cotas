from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Mapping, Protocol, Sequence, TypedDict, cast

from common.constants import (
    COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT,
    COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_SOURCE,
)
from common.jobs import CancellationToken, generate_job_id
from modules.coordinator.steps.shared_execution import handle_step_failure


class _ModuleState(Protocol):
    busy: bool


class _Logger(Protocol):
    def info(self, msg: str, *args: object, **kwargs: object) -> None:
        ...


class _QFileDialog(Protocol):
    def getSaveFileName(
        self,
        parent: object,
        caption: str = ...,
        dir: str = ...,
        filter: str = ...,
    ) -> tuple[str, str]:
        ...


class _CoAttainmentWorkbookResult(Protocol):
    output_path: Path
    duplicate_reg_count: int
    duplicate_entries: tuple[tuple[str, str, str], ...]
    inner_join_drop_count: int
    inner_join_drop_details: tuple[str, ...]


class _FinalReportSignature(Protocol):
    section: str


class _CoordinatorModule(Protocol):
    state: _ModuleState
    _files: list[Path]
    _downloaded_outputs: list[Path]
    _logger: _Logger

    def get_attainment_thresholds(self) -> Sequence[float] | None:
        ...

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        ...

    def _remember_dialog_dir_safe(self, path: str) -> None:
        ...

    def _start_async_operation(
        self,
        *,
        token: CancellationToken,
        job_id: str,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
        on_finally: Callable[[], None],
    ) -> None:
        ...

    def _drain_next_batch(self) -> None:
        ...


class _CalculateNamespace(TypedDict):
    t: Callable[..., str]
    QFileDialog: _QFileDialog
    APP_NAME: str
    resolve_dialog_start_path: Callable[..., str]
    _extract_final_report_signature: Callable[[Path], _FinalReportSignature | None]
    _build_co_attainment_default_name: Callable[..., str]
    _CoAttainmentWorkbookResult: type[_CoAttainmentWorkbookResult]
    _path_key: Callable[[Path], str]
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    show_toast: Callable[..., None]
    _generate_co_attainment_workbook: Callable[..., Path | _CoAttainmentWorkbookResult]
    JobCancelledError: type[Exception]


def calculate_attainment_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_CoordinatorModule, module)
    typed_ns = cast(_CalculateNamespace, ns)
    if typed_module.state.busy or not typed_module._files:
        return

    t = typed_ns["t"]
    signature = typed_ns["_extract_final_report_signature"](typed_module._files[0])
    default_name = typed_ns["_build_co_attainment_default_name"](
        typed_module._files[0],
        section=signature.section if signature is not None else "",
    )
    save_path, _ = typed_ns["QFileDialog"].getSaveFileName(
        typed_module,
        t("coordinator.calculate"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"], default_name),
        t("instructor.dialog.filter.excel_open"),
    )
    if not save_path:
        return

    thresholds_reader = cast(Any, getattr(typed_module, "get_attainment_thresholds", None))
    raw_thresholds = thresholds_reader() if callable(thresholds_reader) else None
    if raw_thresholds is None:
        return
    thresholds = cast(Sequence[float], raw_thresholds)

    process_name = COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT
    token = CancellationToken()
    job_id = generate_job_id()
    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    job_context = (
        workflow_service.create_job_context(
            step_id=COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
            payload={
                WORKFLOW_PAYLOAD_KEY_SOURCE: [str(path) for path in typed_module._files],
                WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path,
                "thresholds": list(thresholds),
            },
        )
        if workflow_service is not None
        else None
    )
    typed_module._publish_status_key("coordinator.status.processing_started")

    def _on_finished(result: object) -> None:
        output_path = Path(save_path)
        duplicate_reg_count = 0
        duplicate_entries: tuple[tuple[str, str, str], ...] = ()
        inner_join_drop_count = 0
        inner_join_drop_details: tuple[str, ...] = ()
        if isinstance(result, typed_ns["_CoAttainmentWorkbookResult"]):
            output_path = result.output_path
            duplicate_reg_count = max(0, int(result.duplicate_reg_count))
            duplicate_entries = result.duplicate_entries
            inner_join_drop_count = max(0, int(result.inner_join_drop_count))
            inner_join_drop_details = tuple(str(item) for item in result.inner_join_drop_details if str(item))
        elif result:
            output_path = Path(str(result))
        if all(
            typed_ns["_path_key"](path) != typed_ns["_path_key"](output_path)
            for path in typed_module._downloaded_outputs
        ):
            typed_module._downloaded_outputs.append(output_path)
        typed_module._remember_dialog_dir_safe(str(output_path))
        typed_module._publish_status_key("coordinator.status.calculate_completed")
        threshold_summary = f"thresholds=({thresholds[0]:g},{thresholds[1]:g},{thresholds[2]:g})"
        typed_ns["log_process_message"](
            process_name,
            logger=typed_module._logger,
            success_message=(
                f"{process_name} completed successfully. output={output_path}, "
                f"duplicates_removed={duplicate_reg_count}, "
                f"inner_join_dropped={inner_join_drop_count}, "
                f"inner_join_details={list(inner_join_drop_details)}, "
                f"{threshold_summary}"
            ),
            user_success_message=typed_ns["build_i18n_log_message"](
                "coordinator.status.calculate_completed",
                fallback=t("coordinator.status.calculate_completed"),
            ),
            job_id=job_context.job_id if job_context else job_id,
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
        )
        typed_ns["show_toast"](
            typed_module,
            t("coordinator.status.calculate_completed"),
            title=t("coordinator.title"),
            level="info",
        )
        if duplicate_reg_count:
            typed_ns["show_toast"](
                typed_module,
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
            typed_module._publish_status_key(
                "coordinator.regno_dedup.log_body",
                count=duplicate_reg_count,
                details=details_text,
            )
        if inner_join_drop_count:
            typed_ns["show_toast"](
                typed_module,
                t("coordinator.join_drop.body", count=inner_join_drop_count),
                title=t("coordinator.title"),
                level="warning",
            )

    def _on_failed(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            job_id=job_context.job_id if job_context else job_id,
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
        )

    typed_module._start_async_operation(
        token=token,
        job_id=job_context.job_id if job_context else job_id,
        work=lambda: (
            workflow_service.calculate_attainment(
                list(typed_module._files),
                Path(save_path),
                context=job_context,
                cancel_token=token,
                thresholds=thresholds,
            )
            if workflow_service is not None and job_context is not None
            else typed_ns["_generate_co_attainment_workbook"](
                list(typed_module._files),
                Path(save_path),
                token=token,
                thresholds=thresholds,
            )
        ),
        on_success=_on_finished,
        on_failure=_on_failed,
        on_finally=typed_module._drain_next_batch,
    )
