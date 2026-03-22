"""CO Analysis step: validate and collect uploaded workbooks."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Mapping, Protocol, TypedDict, cast

from common.constants import (
    CO_ANALYSIS_WORKFLOW_OPERATION_COLLECT_FILES,
    CO_ANALYSIS_WORKFLOW_STEP_ID_COLLECT_FILES,
    WORKFLOW_PAYLOAD_KEY_PATH,
)
from common.jobs import CancellationToken, JobContext
from modules.co_analysis.steps.shared_execution import handle_step_failure


class _State(Protocol):
    busy: bool


class _WorkflowService(Protocol):
    def create_job_context(self, *, step_id: str, payload: Mapping[str, object] | None = None) -> JobContext:
        ...

    def collect_files(
        self,
        candidate_paths: list[str],
        *,
        existing_paths: list[str],
        validate_uploaded_source_workbook: Callable[[str | Path], None],
        context: JobContext,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, object]:
        ...


class _Module(Protocol):
    state: _State
    _files: list[Path]
    _pending_drop_batches: list[list[str]]
    _logger: object
    _workflow_service: _WorkflowService

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
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

    def _drain_next_batch(self) -> None:
        ...

    def _add_uploaded_paths(self, added_paths: list[Path]) -> None:
        ...

    def _refresh_ui(self) -> None:
        ...


class _Ns(TypedDict):
    _validate_uploaded_source_workbook: Callable[[str | Path], None]
    t: Callable[..., str]
    show_toast: Callable[..., None]
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    JobCancelledError: type[Exception]


def collect_files_async(module: object, candidate_paths: list[str], *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_Module, module)
    typed_ns = cast(_Ns, ns)
    if not candidate_paths:
        return
    if typed_module.state.busy:
        typed_module._pending_drop_batches.append(candidate_paths)
        typed_module._publish_status_key("coordinator.status.queued", count=len(candidate_paths))
        return

    process_name = CO_ANALYSIS_WORKFLOW_OPERATION_COLLECT_FILES
    token = CancellationToken()
    job_context = typed_module._workflow_service.create_job_context(
        step_id=CO_ANALYSIS_WORKFLOW_STEP_ID_COLLECT_FILES,
        payload={WORKFLOW_PAYLOAD_KEY_PATH: list(candidate_paths)},
    )
    typed_module._publish_status_key("coordinator.status.processing_started")

    def _on_success(result: object) -> None:
        if not isinstance(result, dict):
            raise RuntimeError("CO Analysis collect_files returned unexpected result.")
        added_paths = [Path(str(value)) for value in cast(list[object], result.get("added", []))]
        duplicates = int(result.get("duplicates", 0))
        invalid = int(result.get("invalid", 0))
        ignored = int(result.get("ignored", duplicates + invalid))

        typed_module._add_uploaded_paths(added_paths)
        if added_paths:
            typed_module._publish_status_key(
                "coordinator.status.added",
                added=len(added_paths),
                total=len(typed_module._files),
            )
        if duplicates or invalid:
            typed_ns["show_toast"](
                typed_module,
                typed_ns["t"](
                    "instructor.toast.step2_upload_reject_summary",
                    invalid=invalid,
                    duplicates=duplicates,
                ),
                title=typed_ns["t"]("instructor.step2.title"),
                level="warning",
            )
        if ignored:
            typed_module._publish_status_key("coordinator.status.ignored", count=ignored)
        typed_ns["log_process_message"](
            process_name,
            logger=typed_module._logger,
            success_message=(
                "collecting co analysis files completed successfully. "
                f"added={len(added_paths)}, duplicates={duplicates}, invalid={invalid}, ignored={ignored}"
            ),
            user_success_message=typed_ns["build_i18n_log_message"](
                "coordinator.status.processing_completed",
                fallback=typed_ns["t"]("coordinator.status.processing_completed"),
            ),
            job_id=job_context.job_id,
            step_id=job_context.step_id,
        )
        typed_module._refresh_ui()

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
        work=lambda: typed_module._workflow_service.collect_files(
            list(candidate_paths),
            existing_paths=[str(path) for path in typed_module._files],
            validate_uploaded_source_workbook=typed_ns["_validate_uploaded_source_workbook"],
            context=job_context,
            cancel_token=token,
        ),
        on_success=_on_success,
        on_failure=_on_failure,
        on_finally=typed_module._drain_next_batch,
    )
