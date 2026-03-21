"""Step 2 (row 2): upload filled marks and generate final report."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Mapping, Protocol, TypedDict, cast

from common.constants import (
    PROCESS_MESSAGE_SUCCESS_SUFFIX,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_PATH,
    WORKFLOW_PAYLOAD_KEY_SOURCE,
    WORKFLOW_STEP_ID_STEP2_GENERATE_FINAL_REPORT,
    WORKFLOW_STEP_ID_STEP2_UPLOAD_FILLED_MARKS,
)
from modules.instructor.steps.shared_execution import handle_step_failure

_LOG_FINAL_REPORT_SOURCE_MISSING = "Final report generation failed: filled marks file is missing. path=%s"
class _ModuleState(Protocol):
    busy: bool


class _Logger(Protocol):
    def info(self, msg: str, *args: object, **kwargs: object) -> None:
        ...

    def warning(self, msg: str, *args: object, **kwargs: object) -> None:
        ...


class _QFileDialog(Protocol):
    def getOpenFileName(
        self,
        parent: object,
        caption: str = ...,
        dir: str = ...,
        filter: str = ...,
    ) -> tuple[str, str]:
        ...

    def getSaveFileName(
        self,
        parent: object,
        caption: str = ...,
        dir: str = ...,
        filter: str = ...,
    ) -> tuple[str, str]:
        ...

    def getExistingDirectory(  # noqa: N802 - Qt naming
        self,
        parent: object,
        caption: str = ...,
        dir: str = ...,
    ) -> str:
        ...


class _CancellationToken(Protocol):
    def raise_if_cancelled(self) -> None:
        ...


class _JobContext(Protocol):
    job_id: str | None
    step_id: str | None


class _InstructorStep2Module(Protocol):
    state: _ModuleState
    _logger: _Logger
    filled_marks_done: bool
    filled_marks_path: str | None
    filled_marks_paths: list[str]
    filled_marks_outdated: bool
    final_report_outdated: bool
    final_report_path: str | None
    final_report_paths: list[str]
    final_report_done: bool
    step1_marks_template_paths: list[str]

    def _remember_dialog_dir_safe(self, path: str) -> None:
        ...

    def _show_step_success_toast(self, step_no: int) -> None:
        ...

    def _show_validation_error_toast(self, message: str) -> None:
        ...

    def _show_system_error_toast(self, step_no: int) -> None:
        ...

    def _can_run_step(self, step_no: int) -> tuple[bool, str]:
        ...


class _StartAsyncOperation(Protocol):
    def __call__(
        self,
        module: _InstructorStep2Module,
        *,
        token: _CancellationToken,
        job_id: str | None,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
    ) -> None:
        ...


class _Step2Namespace(TypedDict):
    t: Callable[..., str]
    _localized_log_messages: Callable[[str], tuple[str, str]]
    QFileDialog: _QFileDialog
    resolve_dialog_start_path: Callable[..., str]
    APP_NAME: str
    CancellationToken: Callable[[], _CancellationToken]
    JobCancelledError: type[Exception]
    ValidationError: type[Exception]
    AppSystemError: type[Exception]
    _publish_status: Callable[..., None]
    _publish_status_key: Callable[..., None]
    _start_async_operation: _StartAsyncOperation
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    show_toast: Callable[..., None]
    _validate_uploaded_filled_marks_workbook: Callable[[str], None]
    _consume_last_filled_marks_anomaly_warnings: Callable[[], list[str]]
    _build_final_report_default_name: Callable[[str | None], str]
    _atomic_copy_file: Callable[[str, str], Path]
    _logger: _Logger

def _emit_status_key(typed_ns: _Step2Namespace, typed_module: _InstructorStep2Module, key: str, **kwargs: object) -> None:
    publish_key = cast(Callable[..., None] | None, typed_ns.get("_publish_status_key"))
    if callable(publish_key):
        publish_key(typed_module, key, **kwargs)
        return
    publish_plain = cast(Callable[..., None] | None, typed_ns.get("_publish_status"))
    if callable(publish_plain):
        publish_plain(typed_module, typed_ns["t"](key, **kwargs))


def _failure_reason(exc: Exception) -> str:
    message = str(exc).strip()
    return message or exc.__class__.__name__


def _format_failed_files_summary(failed_files: list[dict[str, str]]) -> str:
    if not failed_files:
        return ""
    return "; ".join(f"{item['source_path']} -> {item['reason']}" for item in failed_files)


def _format_warning_summary(warnings: list[str]) -> str:
    if not warnings:
        return ""
    preview = warnings[:3]
    extra_count = len(warnings) - len(preview)
    if extra_count > 0:
        preview.append(f"... +{extra_count} more")
    return " | ".join(preview)


def generate_final_reports_from_paths_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep2Module, module)
    typed_ns = cast(_Step2Namespace, ns)
    if typed_module.state.busy:
        return

    t = typed_ns["t"]
    process_name = t("instructor.log.process.generate_final_co_report")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.generate_final_co_report"
    )
    source_paths = list(getattr(typed_module, "filled_marks_paths", []) or [])
    if not source_paths and getattr(typed_module, "filled_marks_path", None):
        source_paths = [cast(str, typed_module.filled_marks_path)]
    if not source_paths:
        return

    output_dir = typed_ns["QFileDialog"].getExistingDirectory(
        typed_module,
        t("instructor.dialog.step2.generate.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"]),
    )
    if not output_dir:
        return
    typed_module._remember_dialog_dir_safe(output_dir)

    planned_pairs: list[tuple[str, str]] = []
    skipped_conflicts = 0
    used_targets: set[str] = set()
    for source_path in source_paths:
        default_name = typed_ns["_build_final_report_default_name"](source_path)
        proposed_path = str(Path(output_dir) / default_name)
        normalized_target = str(Path(proposed_path).resolve(strict=False)).lower()
        if Path(proposed_path).exists():
            replacement_path, _ = typed_ns["QFileDialog"].getSaveFileName(
                typed_module,
                t("instructor.dialog.step2.generate.title"),
                typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"], default_name),
                t("instructor.dialog.filter.excel"),
            )
            if not replacement_path:
                skipped_conflicts += 1
                continue
            proposed_path = replacement_path
            normalized_target = str(Path(proposed_path).resolve(strict=False)).lower()
        candidate = Path(proposed_path)
        suffix = 1
        while normalized_target in used_targets:
            suffix += 1
            candidate = candidate.with_name(f"{candidate.stem}_{suffix}{candidate.suffix}")
            normalized_target = str(candidate.resolve(strict=False)).lower()
        used_targets.add(normalized_target)
        planned_pairs.append((source_path, str(candidate)))

    if not planned_pairs:
        return

    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_GENERATE_FINAL_REPORT,
            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: list(source_paths), WORKFLOW_PAYLOAD_KEY_OUTPUT: output_dir},
        )
        if workflow_service is not None
        else None
    )

    def _work() -> dict[str, object]:
        generated_outputs: list[str] = []
        failed_count = 0
        failed_files: list[dict[str, str]] = []
        anomaly_warnings: list[str] = []
        for source_path, output_path in planned_pairs:
            token.raise_if_cancelled()
            try:
                typed_ns["_validate_uploaded_filled_marks_workbook"](source_path)
                anomaly_warnings.extend(
                    [f"{source_path} -> {msg}" for msg in typed_ns["_consume_last_filled_marks_anomaly_warnings"]()]
                )
                token.raise_if_cancelled()
                if workflow_service is not None and job_context is not None:
                    workflow_service.generate_final_report(
                        source_path,
                        output_path,
                        context=workflow_service.create_job_context(
                            step_id=WORKFLOW_STEP_ID_STEP2_GENERATE_FINAL_REPORT,
                            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: source_path, WORKFLOW_PAYLOAD_KEY_OUTPUT: output_path},
                        ),
                        cancel_token=token,
                    )
                else:
                    raise typed_ns["AppSystemError"](
                        "Workflow service unavailable for final report generation."
                    )
                generated_outputs.append(output_path)
            except Exception as exc:
                failed_count += 1
                failed_files.append(
                    {
                        "source_path": source_path,
                        "output_path": output_path,
                        "reason": _failure_reason(exc),
                    }
                )
        return {
            "generated_outputs": generated_outputs,
            "processed_count": len(planned_pairs),
            "total_count": len(source_paths),
            "failed_count": failed_count,
            "skipped_count": skipped_conflicts,
            "failed_files": failed_files,
            "anomaly_warnings": anomaly_warnings,
        }

    def _on_finished(result: object) -> None:
        data = cast(dict[str, object], result if isinstance(result, dict) else {})
        generated_outputs = [p for p in cast(list[str], data.get("generated_outputs", [])) if p]
        processed_count = int(cast(int, data.get("processed_count", len(planned_pairs))))
        total_count = int(cast(int, data.get("total_count", len(source_paths))))
        failed_count = int(cast(int, data.get("failed_count", 0)))
        skipped_count = int(cast(int, data.get("skipped_count", skipped_conflicts)))
        raw_failed_files = cast(list[object], data.get("failed_files", []))
        anomaly_warnings = [str(item) for item in cast(list[object], data.get("anomaly_warnings", [])) if str(item)]
        failed_files = [
            {
                "source_path": str(item.get("source_path", "")),
                "output_path": str(item.get("output_path", "")),
                "reason": str(item.get("reason", "")),
            }
            for item in raw_failed_files
            if isinstance(item, dict)
        ]
        failure_summary = _format_failed_files_summary(failed_files)
        warning_summary = _format_warning_summary(anomaly_warnings)

        typed_module.final_report_path = generated_outputs[-1] if generated_outputs else None
        typed_module.final_report_paths = list(generated_outputs)
        typed_module.final_report_done = bool(generated_outputs)
        typed_module.final_report_outdated = False
        if generated_outputs:
            _emit_status_key(typed_ns, typed_module, "instructor.status.step2_generated")
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=(
                f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX} "
                f"processed={processed_count}, total={total_count}, failed={failed_count}, skipped={skipped_count}, "
                f"failed_files={failed_files}"
            ),
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        if failure_summary:
            _emit_status_key(
                typed_ns,
                typed_module,
                "instructor.status.step2_generate_per_file_failures",
                details=failure_summary,
            )
        if warning_summary:
            _emit_status_key(
                typed_ns,
                typed_module,
                "instructor.status.step2_validation_warnings",
                details=warning_summary,
            )
            typed_ns["show_toast"](
                typed_module,
                t("instructor.toast.validation_warnings_body"),
                title=t("instructor.toast.validation_warnings_title"),
                level="warning",
            )
        typed_ns["show_toast"](
            typed_module,
            t(
                "instructor.toast.step2_generate_summary",
                processed=processed_count,
                total=total_count,
                generated=len(generated_outputs),
                failed=failed_count,
                skipped=skipped_count,
            ),
            title=t("instructor.step2.title"),
            level="success" if failed_count == 0 else "warning",
        )

    def _on_failed(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            user_error_message=user_error_message,
            step_no=2,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )

    typed_ns["_start_async_operation"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )


def upload_filled_marks_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep2Module, module)
    typed_ns = cast(_Step2Namespace, ns)
    if typed_module.state.busy:
        return

    t = typed_ns["t"]
    process_name = t("instructor.log.process.upload_filled_marks_workbook")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.upload_filled_marks_workbook"
    )
    open_path, _ = typed_ns["QFileDialog"].getOpenFileName(
        typed_module,
        t("instructor.dialog.step2.upload.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"]),
        t("instructor.dialog.filter.excel_open"),
    )
    if not open_path:
        return
    typed_module._remember_dialog_dir_safe(open_path)

    token = typed_ns["CancellationToken"]()
    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_UPLOAD_FILLED_MARKS,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: open_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(result: object) -> None:
        data = cast(dict[str, object], result if isinstance(result, dict) else {})
        anomaly_warnings = [str(item) for item in cast(list[object], data.get("anomaly_warnings", [])) if str(item)]
        replacing = typed_module.filled_marks_done
        typed_module.filled_marks_path = open_path
        typed_module.filled_marks_done = True
        typed_module.filled_marks_outdated = False

        if replacing and typed_module.filled_marks_done:
            typed_module.final_report_outdated = True
            _emit_status_key(typed_ns, typed_module, "instructor.status.step2_changed_filled")
        else:
            _emit_status_key(typed_ns, typed_module, "instructor.status.step2_uploaded_filled")
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        if anomaly_warnings:
            _emit_status_key(
                typed_ns,
                typed_module,
                "instructor.status.step2_validation_warnings",
                details=_format_warning_summary(anomaly_warnings),
            )
            typed_ns["show_toast"](
                typed_module,
                t("instructor.toast.validation_warnings_body"),
                title=t("instructor.toast.validation_warnings_title"),
                level="warning",
            )
        typed_module._show_step_success_toast(2)

    def _on_failed(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            user_error_message=user_error_message,
            step_no=2,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
            show_validation_toast=typed_module._show_validation_error_toast,
        )

    def _work() -> dict[str, object]:
        token.raise_if_cancelled()
        typed_ns["_validate_uploaded_filled_marks_workbook"](open_path)
        anomaly_warnings = typed_ns["_consume_last_filled_marks_anomaly_warnings"]()
        token.raise_if_cancelled()
        return {"ok": True, "anomaly_warnings": anomaly_warnings}

    typed_ns["_start_async_operation"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )


def generate_final_report_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep2Module, module)
    typed_ns = cast(_Step2Namespace, ns)
    if typed_module.state.busy:
        return

    t = typed_ns["t"]
    process_name = t("instructor.log.process.generate_final_co_report")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.generate_final_co_report"
    )
    if not typed_module.filled_marks_done or typed_module.filled_marks_outdated:
        return

    default_name = typed_ns["_build_final_report_default_name"](typed_module.filled_marks_path)
    save_path, _ = typed_ns["QFileDialog"].getSaveFileName(
        typed_module,
        t("instructor.dialog.step2.generate.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"], default_name),
        t("instructor.dialog.filter.excel"),
    )
    if not save_path:
        return

    if not typed_module.filled_marks_path or not Path(typed_module.filled_marks_path).exists():
        typed_ns["_logger"].warning(_LOG_FINAL_REPORT_SOURCE_MISSING, typed_module.filled_marks_path)
        return
    source_path = typed_module.filled_marks_path

    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_GENERATE_FINAL_REPORT,
            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: source_path, WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(result: object) -> None:
        data = cast(dict[str, object], result if isinstance(result, dict) else {})
        anomaly_warnings = [str(item) for item in cast(list[object], data.get("anomaly_warnings", [])) if str(item)]
        output_path_raw = data.get("output_path", save_path)
        output_path = str(output_path_raw)
        typed_module.final_report_path = output_path
        typed_module.final_report_done = True
        typed_module.final_report_outdated = False
        typed_module._remember_dialog_dir_safe(output_path)
        _emit_status_key(typed_ns, typed_module, "instructor.status.step2_generated")
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        typed_module._show_step_success_toast(2)
        if anomaly_warnings:
            _emit_status_key(
                typed_ns,
                typed_module,
                "instructor.status.step2_validation_warnings",
                details=_format_warning_summary(anomaly_warnings),
            )
            typed_ns["show_toast"](
                typed_module,
                t("instructor.toast.validation_warnings_body"),
                title=t("instructor.toast.validation_warnings_title"),
                level="warning",
            )

    def _on_failed(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            user_error_message=user_error_message,
            step_no=2,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
            show_validation_toast=typed_module._show_validation_error_toast,
        )

    def _work() -> dict[str, object]:
        token.raise_if_cancelled()
        typed_ns["_validate_uploaded_filled_marks_workbook"](source_path)
        anomaly_warnings = typed_ns["_consume_last_filled_marks_anomaly_warnings"]()
        token.raise_if_cancelled()
        if workflow_service is not None and job_context is not None:
            output_path = workflow_service.generate_final_report(
                source_path,
                save_path,
                context=job_context,
                cancel_token=token,
            )
        else:
            raise typed_ns["AppSystemError"]("Workflow service unavailable for final report generation.")
        return {"output_path": str(output_path), "anomaly_warnings": anomaly_warnings}

    typed_ns["_start_async_operation"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )







