"""Step 2: upload course details and prepare marks template."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Mapping, Protocol, TypedDict, cast

from common.constants import (
    PROCESS_MESSAGE_SUCCESS_SUFFIX,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_PATH,
    WORKFLOW_PAYLOAD_KEY_SOURCE,
    WORKFLOW_STEP_ID_STEP2_GENERATE_MARKS_TEMPLATE,
    WORKFLOW_STEP_ID_STEP2_VALIDATE_COURSE_DETAILS,
)
from modules.instructor.steps.shared_execution import handle_step_failure

_PROGRESS_UPDATE_BATCH_SIZE = 10


def _coerce_int(value: object, default: int) -> int:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        try:
            return int(value.strip())
        except ValueError:
            return default
    return default


def _failure_reason(exc: Exception) -> str:
    message = str(exc).strip()
    return message or exc.__class__.__name__


def _format_failed_files_summary(failed_files: list[dict[str, str]]) -> str:
    if not failed_files:
        return ""
    preview = failed_files[:5]
    parts = [f"{item['source_path']} -> {item['reason']}" for item in preview]
    extra_count = len(failed_files) - len(preview)
    if extra_count > 0:
        parts.append(f"... +{extra_count} more")
    return "; ".join(parts)


class _ModuleState(Protocol):
    busy: bool


class _Logger(Protocol):
    def info(self, msg: str, *args: object, **kwargs: object) -> None:
        ...


class _QFileDialog(Protocol):
    def getOpenFileNames(
        self,
        parent: object,
        caption: str = ...,
        dir: str = ...,
        filter: str = ...,
    ) -> tuple[list[str], str]:
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
    marks_template_done: bool
    step2_upload_ready: bool
    step2_course_details_path: str | None
    step1_course_details_paths: list[str]
    marks_template_path: str | None
    marks_template_paths: list[str]
    _step2_marks_default_name: str | None
    final_report_outdated: bool
    final_report_done: bool
    filled_marks_outdated: bool
    filled_marks_done: bool

    def _remember_dialog_dir_safe(self, path: str) -> None:
        ...

    def _show_step_success_toast(self, step_no: int) -> None:
        ...

    def _show_validation_error_toast(self, message: str) -> None:
        ...

    def _show_system_error_toast(self, step_no: int) -> None:
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
    _logger: _Logger
    ID_COURSE_SETUP: str
    validate_course_details_workbook: Callable[[str], str]
    _build_marks_template_default_name: Callable[[str], str]
    generate_marks_template_from_course_details: Callable[[str, str], None]

def _emit_status_key(typed_ns: _Step2Namespace, typed_module: _InstructorStep2Module, key: str, **kwargs: object) -> None:
    publish_key = cast(Callable[..., None] | None, typed_ns.get("_publish_status_key"))
    if callable(publish_key):
        publish_key(typed_module, key, **kwargs)
        return
    publish_plain = cast(Callable[..., None] | None, typed_ns.get("_publish_status"))
    if callable(publish_plain):
        publish_plain(typed_module, typed_ns["t"](key, **kwargs))


def upload_course_details_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep2Module, module)
    typed_ns = cast(_Step2Namespace, ns)
    if typed_module.state.busy:
        return

    open_paths, _ = typed_ns["QFileDialog"].getOpenFileNames(
        typed_module,
        typed_ns["t"]("instructor.dialog.step2.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"]),
        typed_ns["t"]("instructor.dialog.filter.excel_open"),
    )
    if not open_paths:
        return
    upload_course_details_from_paths_async(module, open_paths, ns=ns)


def upload_course_details_from_paths_async(
    module: object,
    open_paths: list[str],
    *,
    show_success_toast: bool = False,
    ns: Mapping[str, object],
) -> None:
    typed_module = cast(_InstructorStep2Module, module)
    typed_ns = cast(_Step2Namespace, ns)
    if typed_module.state.busy:
        return
    normalized_paths = [path for path in open_paths if path]
    if not normalized_paths:
        return

    t = typed_ns["t"]
    process_name = t("instructor.log.process.validate_course_details_workbook")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.validate_course_details_workbook"
    )
    typed_module._remember_dialog_dir_safe(normalized_paths[0])

    seen: set[str] = set()
    unique_paths: list[str] = []
    duplicate_input_count = 0
    for path in normalized_paths:
        key = str(Path(path).resolve(strict=False)).lower()
        if key in seen:
            duplicate_input_count += 1
            continue
        seen.add(key)
        unique_paths.append(path)

    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_VALIDATE_COURSE_DETAILS,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: list(unique_paths)},
        )
        if workflow_service is not None
        else None
    )

    def _work() -> dict[str, object]:
        valid_paths: list[str] = []
        invalid_paths: list[str] = []
        mismatched_template_paths: list[str] = []
        total_unique = len(unique_paths)
        processed = 0

        def _publish_progress_if_needed(*, force: bool = False) -> None:
            if not force and processed % _PROGRESS_UPDATE_BATCH_SIZE != 0:
                return
            _emit_status_key(
                typed_ns,
                typed_module,
                "instructor.status.step1_validating_progress",
                processed=processed,
                total=total_unique,
            )

        for path in unique_paths:
            token.raise_if_cancelled()
            try:
                template_id = (
                    workflow_service.validate_course_details_workbook(
                        path,
                        context=workflow_service.create_job_context(
                            step_id=WORKFLOW_STEP_ID_STEP2_VALIDATE_COURSE_DETAILS,
                            payload={WORKFLOW_PAYLOAD_KEY_PATH: path},
                        ),
                        cancel_token=token,
                    )
                    if workflow_service is not None
                    else typed_ns["validate_course_details_workbook"](path)
                )
            except typed_ns["JobCancelledError"]:
                raise
            except Exception:
                invalid_paths.append(path)
                processed += 1
                _publish_progress_if_needed()
                continue
            if template_id != typed_ns["ID_COURSE_SETUP"]:
                mismatched_template_paths.append(path)
                processed += 1
                _publish_progress_if_needed()
                continue
            valid_paths.append(path)
            processed += 1
            _publish_progress_if_needed()
        _publish_progress_if_needed(force=True)
        return {
            "valid_paths": valid_paths,
            "invalid_paths": invalid_paths,
            "mismatched_template_paths": mismatched_template_paths,
            "duplicate_input_count": duplicate_input_count,
            "total_input_count": len(normalized_paths),
        }

    def _on_finished(result: object) -> None:
        data = cast(dict[str, object], result if isinstance(result, dict) else {})
        valid_paths = [p for p in cast(list[str], data.get("valid_paths", [])) if p]
        invalid_paths = [p for p in cast(list[str], data.get("invalid_paths", [])) if p]
        mismatched_template_paths = [p for p in cast(list[str], data.get("mismatched_template_paths", [])) if p]
        duplicates = _coerce_int(data.get("duplicate_input_count"), 0)
        total = _coerce_int(data.get("total_input_count"), len(normalized_paths))
        existing_paths = list(getattr(typed_module, "step1_course_details_paths", []) or [])

        merged_paths: list[str] = []
        seen_targets: set[str] = set()
        for path in [*existing_paths, *valid_paths]:
            key = str(Path(path).resolve(strict=False)).lower()
            if key in seen_targets:
                continue
            seen_targets.add(key)
            merged_paths.append(path)

        replacing = typed_module.marks_template_done or typed_module.step2_upload_ready
        typed_module.step1_course_details_paths = list(merged_paths)
        typed_module.step2_course_details_path = merged_paths[0] if merged_paths else None
        typed_module.step2_upload_ready = bool(merged_paths)
        typed_module.marks_template_done = False
        typed_module.marks_template_path = None
        typed_module.marks_template_paths = []
        fallback_name = t("instructor.dialog.step1.prepare.default_name")
        typed_module._step2_marks_default_name = (
            typed_ns["_build_marks_template_default_name"](merged_paths[0]) if merged_paths else fallback_name
        ) or fallback_name

        display_paths_setter = getattr(typed_module, "_set_step1_course_details_files", None)
        if callable(display_paths_setter):
            display_paths_setter(merged_paths)

        if replacing:
            typed_module.final_report_outdated = typed_module.final_report_done
            typed_module.filled_marks_outdated = typed_module.filled_marks_done
            if typed_module.filled_marks_outdated or typed_module.final_report_outdated:
                _emit_status_key(typed_ns, typed_module, "instructor.status.step1_changed")
        elif merged_paths:
            _emit_status_key(typed_ns, typed_module, "instructor.status.step1_validated")

        _emit_status_key(
            typed_ns,
            typed_module,
            "instructor.status.step1_validated_progress",
            valid=len(valid_paths),
            total=total,
        )
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=(
                f"{process_name} completed successfully. "
                f"total={total}, valid={len(valid_paths)}, invalid={len(invalid_paths)}, "
                f"mismatched={len(mismatched_template_paths)}, duplicates={duplicates}"
            ),
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )

        if show_success_toast:
            typed_module._show_step_success_toast(1)
        has_validation_errors = bool(invalid_paths or mismatched_template_paths or duplicates > 0 or not valid_paths)
        if has_validation_errors:
            typed_ns["show_toast"](
                typed_module,
                t(
                    "instructor.toast.step1_validation_summary",
                    valid=len(valid_paths),
                    invalid=len(invalid_paths),
                    mismatched=len(mismatched_template_paths),
                    duplicates=duplicates,
                ),
                title=t("instructor.step1.title"),
                level="warning",
            )

    def _on_failed(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            user_error_message=user_error_message,
            step_no=1,
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

def prepare_marks_template_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep2Module, module)
    typed_ns = cast(_Step2Namespace, ns)
    if typed_module.state.busy:
        return

    t = typed_ns["t"]
    process_name = t("instructor.log.process.generate_marks_template")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.generate_marks_template"
    )
    source_paths = list(getattr(typed_module, "step1_course_details_paths", []) or [])
    if not source_paths:
        return

    planned_pairs: list[tuple[str, str]] = []
    skipped_conflicts = 0
    used_targets: set[str] = set()
    output_dir = typed_ns["QFileDialog"].getExistingDirectory(
        typed_module,
        t("instructor.dialog.step1.prepare.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"]),
    )
    if not output_dir:
        return
    typed_module._remember_dialog_dir_safe(output_dir)

    for source_path in source_paths:
        default_output_name = typed_ns["_build_marks_template_default_name"](source_path)
        proposed_path = str(Path(output_dir) / default_output_name)
        normalized_target = str(Path(proposed_path).resolve(strict=False)).lower()
        if Path(proposed_path).exists():
            replacement_path, _ = typed_ns["QFileDialog"].getSaveFileName(
                typed_module,
                t("instructor.dialog.step1.prepare.title"),
                typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"], default_output_name),
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

    output_base = (
        str(Path(planned_pairs[0][1]).parent)
        if planned_pairs
        else typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"])
    )
    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_GENERATE_MARKS_TEMPLATE,
            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: list(source_paths), WORKFLOW_PAYLOAD_KEY_OUTPUT: output_base},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(result: object) -> None:
        data = cast(dict[str, object], result if isinstance(result, dict) else {})
        if not data and planned_pairs:
            data = {
                "generated_outputs": [output for _, output in planned_pairs],
                "processed_count": len(planned_pairs),
                "total_count": len(source_paths),
                "failed_count": 0,
                "skipped_count": skipped_conflicts,
            }
        generated_outputs = [p for p in cast(list[str], data.get("generated_outputs", [])) if p]
        processed_count = _coerce_int(data.get("processed_count"), len(generated_outputs))
        total_count = _coerce_int(data.get("total_count"), len(source_paths))
        failed_count = _coerce_int(data.get("failed_count"), 0)
        skipped_count = _coerce_int(data.get("skipped_count"), skipped_conflicts)
        raw_failed_files = cast(list[object], data.get("failed_files", []))
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

        typed_module.marks_template_path = generated_outputs[-1] if generated_outputs else None
        typed_module.marks_template_paths = list(generated_outputs)
        typed_module.marks_template_done = bool(generated_outputs)
        typed_module.filled_marks_outdated = typed_module.filled_marks_done
        typed_module.final_report_outdated = typed_module.final_report_done
        if generated_outputs:
            _emit_status_key(typed_ns, typed_module, "instructor.status.step1_prepared")
        _emit_status_key(
            typed_ns,
            typed_module,
            "instructor.status.step1_prepare_progress",
            processed=processed_count,
            total=total_count,
        )
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
                "instructor.status.step1_prepare_per_file_failures",
                details=failure_summary,
            )
        typed_ns["show_toast"](
            typed_module,
            t(
                "instructor.toast.step1_prepare_summary",
                processed=processed_count,
                total=total_count,
                generated=len(generated_outputs),
                failed=failed_count,
                skipped=skipped_count,
            ),
            title=t("instructor.step1.title"),
            level="success" if failed_count == 0 else "warning",
        )

    def _on_failed(exc: Exception) -> None:
        handle_step_failure(
            exc=exc,
            ns=typed_ns,
            module=typed_module,
            process_name=process_name,
            user_error_message=user_error_message,
            step_no=1,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
            show_validation_toast=typed_module._show_validation_error_toast,
        )

    def _work() -> object:
        generated_outputs: list[str] = []
        failed_count = 0
        failed_files: list[dict[str, str]] = []
        processed_count = 0
        total_count = len(planned_pairs)
        for source_path, output_path in planned_pairs:
            token.raise_if_cancelled()
            try:
                if workflow_service is not None and job_context is not None:
                    workflow_service.generate_marks_template(
                        source_path,
                        output_path,
                        context=workflow_service.create_job_context(
                            step_id=WORKFLOW_STEP_ID_STEP2_GENERATE_MARKS_TEMPLATE,
                            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: source_path, WORKFLOW_PAYLOAD_KEY_OUTPUT: output_path},
                        ),
                        cancel_token=token,
                    )
                else:
                    typed_ns["generate_marks_template_from_course_details"](source_path, output_path)
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
            finally:
                processed_count += 1
                if processed_count % _PROGRESS_UPDATE_BATCH_SIZE == 0 or processed_count == total_count:
                    _emit_status_key(
                        typed_ns,
                        typed_module,
                        "instructor.status.step1_prepare_progress",
                        processed=processed_count,
                        total=total_count,
                    )
        return {
            "generated_outputs": generated_outputs,
            "processed_count": len(planned_pairs),
            "total_count": len(source_paths),
            "failed_count": failed_count,
            "skipped_count": skipped_conflicts,
            "failed_files": failed_files,
        }

    typed_ns["_start_async_operation"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )





