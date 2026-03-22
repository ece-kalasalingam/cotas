"""CO Analysis step: generate final workbook asynchronously."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Mapping, Protocol, TypedDict, cast

from common.constants import (
    CO_ANALYSIS_WORKFLOW_OPERATION_GENERATE_WORKBOOK,
    CO_ANALYSIS_WORKFLOW_STEP_ID_GENERATE_WORKBOOK,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_SOURCE,
)
from common.jobs import CancellationToken, JobContext
from modules.co_analysis.steps.shared_execution import handle_step_failure


class _State(Protocol):
    busy: bool


class _WorkflowService(Protocol):
    def create_job_context(self, *, step_id: str, payload: Mapping[str, object] | None = None) -> JobContext:
        ...

    def generate_workbook(
        self,
        source_paths: list[Path],
        output_path: Path,
        *,
        thresholds: tuple[float, float, float] | None = None,
        co_attainment_percent: float | None = None,
        co_attainment_level: int | None = None,
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> object:
        ...


class _Module(Protocol):
    state: _State
    _files: list[Path]
    _downloaded_outputs: list[Path]
    _logger: object
    _workflow_service: _WorkflowService

    def _read_attainment_thresholds(self) -> tuple[float, float, float]:
        ...

    def _read_co_attainment_target(self) -> tuple[float, int]:
        ...

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        ...

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        ...

    def _start_async_operation(
        self,
        *,
        token: CancellationToken,
        job_id: str | None,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
        on_finally: Callable[[], None] | None = None,
    ) -> None:
        ...


class _QFileDialog(Protocol):
    def getExistingDirectory(self, parent: object, caption: str = ..., dir: str = ...) -> str:
        ...

    def getSaveFileName(
        self,
        parent: object,
        caption: str = ...,
        dir: str = ...,
        filter: str = ...,
    ) -> tuple[str, str]:
        ...


class _Ns(TypedDict):
    QFileDialog: _QFileDialog
    APP_NAME: str
    t: Callable[..., str]
    resolve_dialog_start_path: Callable[..., str]
    canonical_path_key: Callable[[Path], str]
    _extract_course_metadata_and_students: Callable[[Path], tuple[set[str], dict[str, str]]]
    _sanitize_filename_token: Callable[[object], str]
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    show_toast: Callable[..., None]
    normalize: Callable[[object], str]
    COURSE_METADATA_COURSE_CODE_KEY: str
    COURSE_METADATA_ACADEMIC_YEAR_KEY: str
    JobCancelledError: type[Exception]


def save_workbook_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_Module, module)
    typed_ns = cast(_Ns, ns)
    if typed_module.state.busy or not typed_module._files:
        return

    process_name = CO_ANALYSIS_WORKFLOW_OPERATION_GENERATE_WORKBOOK
    output_dir = typed_ns["QFileDialog"].getExistingDirectory(
        typed_module,
        typed_ns["t"]("coordinator.calculate"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"]),
    )
    if not output_dir:
        return
    typed_module._remember_dialog_dir_safe(output_dir)

    first_metadata = typed_ns["_extract_course_metadata_and_students"](typed_module._files[0])[1]
    prefix = ""
    course_code = typed_ns["_sanitize_filename_token"](
        first_metadata.get(typed_ns["normalize"](typed_ns["COURSE_METADATA_COURSE_CODE_KEY"]), "")
    )
    academic_year = typed_ns["_sanitize_filename_token"](
        first_metadata.get(typed_ns["normalize"](typed_ns["COURSE_METADATA_ACADEMIC_YEAR_KEY"]), "")
    )
    if course_code and academic_year:
        prefix = f"{course_code}_{academic_year}_"

    output_path = Path(output_dir) / f"{prefix}CO_Analysis.xlsx"
    if output_path.exists():
        replacement_path, _ = typed_ns["QFileDialog"].getSaveFileName(
            typed_module,
            typed_ns["t"]("coordinator.calculate"),
            typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"], output_path.name),
            typed_ns["t"]("instructor.dialog.filter.excel"),
        )
        if not replacement_path:
            return
        output_path = Path(replacement_path)

    token = CancellationToken()
    thresholds = typed_module._read_attainment_thresholds()
    co_attainment_percent, co_attainment_level = typed_module._read_co_attainment_target()
    job_context = typed_module._workflow_service.create_job_context(
        step_id=CO_ANALYSIS_WORKFLOW_STEP_ID_GENERATE_WORKBOOK,
        payload={
            WORKFLOW_PAYLOAD_KEY_SOURCE: [str(path) for path in typed_module._files],
            WORKFLOW_PAYLOAD_KEY_OUTPUT: str(output_path),
            "thresholds": list(thresholds),
            "co_attainment_percent": co_attainment_percent,
            "co_attainment_level": co_attainment_level,
        },
    )
    typed_module._publish_status_key("coordinator.status.processing_started")

    def _on_success(result: object) -> None:
        result_path = Path(str(getattr(result, "output_path", result)))
        if all(
            typed_ns["canonical_path_key"](path) != typed_ns["canonical_path_key"](result_path)
            for path in typed_module._downloaded_outputs
        ):
            typed_module._downloaded_outputs.append(result_path)
        typed_module._remember_dialog_dir_safe(str(result_path))
        typed_module._publish_status_key("coordinator.status.calculate_completed")
        typed_ns["log_process_message"](
            process_name,
            logger=typed_module._logger,
            success_message=(
                "saving co analysis workbooks completed successfully. "
                f"output_dir={output_dir}, generated=1, "
                f"thresholds=({thresholds[0]:g},{thresholds[1]:g},{thresholds[2]:g}), "
                f"co_at_target=({co_attainment_percent:g},L{co_attainment_level})"
            ),
            user_success_message=typed_ns["build_i18n_log_message"](
                "coordinator.status.calculate_completed",
                fallback=typed_ns["t"]("coordinator.status.calculate_completed"),
            ),
            job_id=job_context.job_id,
            step_id=job_context.step_id,
        )
        typed_ns["show_toast"](
            typed_module,
            typed_ns["t"]("coordinator.status.calculate_completed"),
            title=typed_ns["t"]("coordinator.title"),
            level="info",
        )

    def _on_failure(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            job_id=job_context.job_id,
            step_id=job_context.step_id,
        )

    typed_module._start_async_operation(
        token=token,
        job_id=job_context.job_id,
        work=lambda: typed_module._workflow_service.generate_workbook(
            list(typed_module._files),
            Path(output_path),
            thresholds=thresholds,
            co_attainment_percent=co_attainment_percent,
            co_attainment_level=co_attainment_level,
            context=job_context,
            cancel_token=token,
        ),
        on_success=_on_success,
        on_failure=_on_failure,
    )
