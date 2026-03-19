"""Coordinator step: validate and collect dropped final report files."""

from __future__ import annotations

from pathlib import Path, PureWindowsPath
from typing import Any, Callable, Mapping, Protocol, TypedDict, cast

from PySide6.QtCore import Qt

from common.constants import (
    COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES,
    COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
    WORKFLOW_PAYLOAD_KEY_PATH,
)
from common.jobs import CancellationToken, generate_job_id
from modules.coordinator.steps.shared_execution import handle_step_failure


class _ModuleState(Protocol):
    busy: bool


class _Logger(Protocol):
    def info(self, msg: str, *args: object, **kwargs: object) -> None:
        ...


class _DropList(Protocol):
    def addItem(self, item: object) -> None:  # noqa: N802
        ...

    def setItemWidget(self, item: object, widget: object) -> None:  # noqa: N802
        ...


class _FileItemWidget(Protocol):
    class _RemovedSignal(Protocol):
        def connect(self, slot: Callable[[str], None]) -> None:
            ...

    removed: _RemovedSignal

    def sizeHint(self) -> object:  # noqa: N802
        ...


class _CoordinatorModule(Protocol):
    state: _ModuleState
    _files: list[Path]
    _pending_drop_batches: list[list[str]]
    _logger: _Logger
    drop_list: _DropList

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
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

    def _add_uploaded_paths(self, added_paths: list[Path]) -> None:
        ...

    def _new_file_item_widget(self, path_text: str, *, parent: object) -> _FileItemWidget:
        ...

    def _remove_file_by_path(self, file_path: str) -> None:
        ...


class _QListWidgetItem(Protocol):
    def setToolTip(self, text: str) -> None:  # noqa: N802
        ...

    def setData(self, role: object, value: object) -> None:  # noqa: N802
        ...

    def setSizeHint(self, hint: object) -> None:  # noqa: N802
        ...


class _AnalyzeResult(TypedDict):
    added: list[str]
    duplicates: int
    invalid_final_report: list[str]
    ignored: int


class _CollectNamespace(TypedDict):
    t: Callable[..., str]
    _path_key: Callable[[Path], str]
    show_toast: Callable[..., None]
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    _analyze_dropped_files: Callable[..., _AnalyzeResult]
    QListWidgetItem: Callable[[], _QListWidgetItem]
    JobCancelledError: type[Exception]


def process_files_async(module: object, dropped_files: list[str], *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_CoordinatorModule, module)
    typed_ns = cast(_CollectNamespace, ns)
    if not dropped_files:
        return
    if typed_module.state.busy:
        typed_module._pending_drop_batches.append(dropped_files)
        typed_module._publish_status_key("coordinator.status.queued", count=len(dropped_files))
        return

    t = typed_ns["t"]
    process_name = COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES
    token = CancellationToken()
    job_id = generate_job_id()
    existing_keys = {typed_ns["_path_key"](path) for path in typed_module._files}
    existing_paths = [str(path) for path in typed_module._files]
    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    job_context = (
        workflow_service.create_job_context(
            step_id=COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: list(dropped_files)},
        )
        if workflow_service is not None
        else None
    )
    typed_module._publish_status_key("coordinator.status.processing_started")

    def _on_finished(result: object) -> None:
        if not isinstance(result, dict):
            raise RuntimeError("Coordinator processing returned unexpected result type.")
        typed_result = cast(_AnalyzeResult, result)
        added_paths = [Path(value) for value in typed_result.get("added", [])]
        duplicates = int(typed_result.get("duplicates", 0))
        invalid_paths = [Path(value) for value in typed_result.get("invalid_final_report", [])]
        ignored = int(typed_result.get("ignored", 0))

        typed_module._add_uploaded_paths(added_paths)

        if added_paths:
            typed_module._publish_status_key(
                "coordinator.status.added",
                added=len(added_paths),
                total=len(typed_module._files),
            )
        if duplicates:
            typed_ns["show_toast"](
                typed_module,
                t("coordinator.duplicate.body", count=duplicates),
                title=t("coordinator.duplicate.title"),
                level="info",
            )
        if invalid_paths:
            file_names = "\n".join(path.name for path in invalid_paths)
            typed_ns["show_toast"](
                typed_module,
                t(
                    "coordinator.invalid_final_report.body",
                    count=len(invalid_paths),
                    files=file_names,
                ),
                title=t("coordinator.invalid_final_report.title"),
                level="warning",
            )
        if ignored:
            typed_module._publish_status_key("coordinator.status.ignored", count=ignored)

        typed_ns["log_process_message"](
            process_name,
            logger=typed_module._logger,
            success_message=(
                f"{process_name} completed successfully. "
                f"added={len(added_paths)}, duplicates={duplicates}, "
                f"invalid={len(invalid_paths)}, ignored={ignored}"
            ),
            user_success_message=typed_ns["build_i18n_log_message"](
                "coordinator.status.processing_completed",
                fallback=t("coordinator.status.processing_completed"),
            ),
            job_id=job_context.job_id if job_context else job_id,
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
        )

    def _on_failed(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            job_id=job_context.job_id if job_context else job_id,
            step_id=job_context.step_id if job_context else COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
        )

    typed_module._start_async_operation(
        token=token,
        job_id=job_context.job_id if job_context else job_id,
        work=lambda: (
            workflow_service.collect_files(
                dropped_files,
                existing_keys=existing_keys,
                existing_paths=existing_paths,
                analyze_dropped_files=typed_ns["_analyze_dropped_files"],
                context=job_context,
                cancel_token=token,
            )
            if workflow_service is not None and job_context is not None
            else typed_ns["_analyze_dropped_files"](
                dropped_files,
                existing_keys=existing_keys,
                existing_paths=existing_paths,
                token=token,
            )
        ),
        on_success=_on_finished,
        on_failure=_on_failed,
        on_finally=typed_module._drain_next_batch,
    )


def add_uploaded_paths(module: object, added_paths: list[Path], *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_CoordinatorModule, module)
    typed_ns = cast(_CollectNamespace, ns)
    for path in added_paths:
        typed_module._files.append(path)
        path_text = str(path)
        if len(path_text) >= 2 and path_text[1] == ":":
            path_text = str(PureWindowsPath(path_text))
        item = typed_ns["QListWidgetItem"]()
        item.setToolTip(path_text)
        item.setData(Qt.ItemDataRole.UserRole, path_text)
        typed_module.drop_list.addItem(item)
        row_widget = typed_module._new_file_item_widget(path_text, parent=typed_module.drop_list)
        row_widget.removed.connect(typed_module._remove_file_by_path)
        item.setSizeHint(row_widget.sizeHint())
        typed_module.drop_list.setItemWidget(item, row_widget)


