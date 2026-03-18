"""Step 3: upload filled marks and generate final report."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Mapping, Protocol, TypedDict, cast

from common.constants import (
    LOG_EXTRA_KEY_JOB_ID,
    LOG_EXTRA_KEY_STEP_ID,
    LOG_EXTRA_KEY_USER_MESSAGE,
    PROCESS_MESSAGE_CANCELLED_TEMPLATE,
    PROCESS_MESSAGE_SUCCESS_SUFFIX,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_PATH,
    WORKFLOW_PAYLOAD_KEY_SOURCE,
    WORKFLOW_STEP_ID_STEP3_GENERATE_FINAL_REPORT,
    WORKFLOW_STEP_ID_STEP3_UPLOAD_FILLED_MARKS,
)

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


class _CancellationToken(Protocol):
    def raise_if_cancelled(self) -> None:
        ...


class _JobContext(Protocol):
    job_id: str | None
    step_id: str | None


class _InstructorStep3Module(Protocol):
    state: _ModuleState
    _logger: _Logger
    filled_marks_done: bool
    filled_marks_path: str | None
    filled_marks_outdated: bool
    final_report_outdated: bool
    final_report_path: str | None
    final_report_done: bool

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
        module: _InstructorStep3Module,
        *,
        token: _CancellationToken,
        job_id: str | None,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
    ) -> None:
        ...


class _Step3Namespace(TypedDict):
    t: Callable[..., str]
    _localized_log_messages: Callable[[str], tuple[str, str]]
    QFileDialog: _QFileDialog
    resolve_dialog_start_path: Callable[..., str]
    APP_NAME: str
    CancellationToken: Callable[[], _CancellationToken]
    JobCancelledError: type[Exception]
    ValidationError: type[Exception]
    AppSystemError: type[Exception]
    _publish_status_compat: Callable[..., None]
    _start_async_operation_compat: _StartAsyncOperation
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    show_toast: Callable[..., None]
    _validate_uploaded_filled_marks_workbook: Callable[[str], None]
    _build_final_report_default_name: Callable[[str | None], str]
    _atomic_copy_file: Callable[[str, str], Path]
    _logger: _Logger


def upload_filled_marks_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep3Module, module)
    typed_ns = cast(_Step3Namespace, ns)
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
            step_id=WORKFLOW_STEP_ID_STEP3_UPLOAD_FILLED_MARKS,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: open_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        replacing = typed_module.filled_marks_done
        typed_module.filled_marks_path = open_path
        typed_module.filled_marks_done = True
        typed_module.filled_marks_outdated = False

        if replacing and typed_module.filled_marks_done:
            typed_module.final_report_outdated = True
            typed_ns["_publish_status_compat"](typed_module, t("instructor.status.step2_changed_filled"))
        else:
            typed_ns["_publish_status_compat"](typed_module, t("instructor.status.step2_uploaded_filled"))
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        typed_module._show_step_success_toast(2)

    def _on_failed(exc: Exception) -> None:
        if isinstance(exc, typed_ns["JobCancelledError"]):
            status_key = "instructor.status.operation_cancelled"
            user_message = t(status_key)
            user_message_payload = typed_ns["build_i18n_log_message"](status_key, fallback=user_message)
            typed_ns["_publish_status_compat"](typed_module, user_message)
            typed_ns["_logger"].info(
                PROCESS_MESSAGE_CANCELLED_TEMPLATE,
                process_name,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: user_message_payload,
                    LOG_EXTRA_KEY_JOB_ID: job_context.job_id if job_context else None,
                    LOG_EXTRA_KEY_STEP_ID: job_context.step_id if job_context else None,
                },
            )
            return
        if isinstance(exc, typed_ns["ValidationError"]):
            typed_ns["log_process_message"](
                process_name,
                logger=typed_ns["_logger"],
                error=exc,
                notify=lambda message, _level: typed_module._show_validation_error_toast(message),
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
        elif isinstance(exc, typed_ns["AppSystemError"]):
            typed_ns["log_process_message"](
                process_name,
                logger=typed_ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            typed_module._show_system_error_toast(2)
        else:
            typed_ns["log_process_message"](
                process_name,
                logger=typed_ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            typed_module._show_system_error_toast(2)

    def _work() -> bool:
        token.raise_if_cancelled()
        typed_ns["_validate_uploaded_filled_marks_workbook"](open_path)
        token.raise_if_cancelled()
        return True

    typed_ns["_start_async_operation_compat"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )


def generate_final_report_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep3Module, module)
    typed_ns = cast(_Step3Namespace, ns)
    if typed_module.state.busy:
        return

    t = typed_ns["t"]
    process_name = t("instructor.log.process.generate_final_co_report")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.generate_final_co_report"
    )
    can_run, reason = typed_module._can_run_step(2)
    if not can_run:
        typed_ns["show_toast"](
            typed_module,
            reason,
            title=t("instructor.msg.step_required_title"),
            level="info",
        )
        return
    if not typed_module.filled_marks_done or typed_module.filled_marks_outdated:
        typed_ns["show_toast"](
            typed_module,
            t("instructor.require.step2"),
            title=t("instructor.msg.step_required_title"),
            level="info",
        )
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
        typed_ns["show_toast"](
            typed_module,
            t("instructor.require.step2"),
            title=t("instructor.msg.step_required_title"),
            level="error",
        )
        return
    source_path = typed_module.filled_marks_path

    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP3_GENERATE_FINAL_REPORT,
            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: source_path, WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        typed_module.final_report_path = save_path
        typed_module.final_report_done = True
        typed_module.final_report_outdated = False
        typed_module._remember_dialog_dir_safe(save_path)
        typed_ns["_publish_status_compat"](typed_module, t("instructor.status.step2_generated"))
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        typed_module._show_step_success_toast(2)

    def _on_failed(exc: Exception) -> None:
        if isinstance(exc, typed_ns["JobCancelledError"]):
            status_key = "instructor.status.operation_cancelled"
            user_message = t(status_key)
            user_message_payload = typed_ns["build_i18n_log_message"](status_key, fallback=user_message)
            typed_ns["_publish_status_compat"](typed_module, user_message)
            typed_ns["_logger"].info(
                PROCESS_MESSAGE_CANCELLED_TEMPLATE,
                process_name,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: user_message_payload,
                    LOG_EXTRA_KEY_JOB_ID: job_context.job_id if job_context else None,
                    LOG_EXTRA_KEY_STEP_ID: job_context.step_id if job_context else None,
                },
            )
            return
        if isinstance(exc, typed_ns["ValidationError"]):
            typed_ns["log_process_message"](
                process_name,
                logger=typed_ns["_logger"],
                error=exc,
                notify=lambda message, _level: typed_module._show_validation_error_toast(message),
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
        else:
            typed_ns["log_process_message"](
                process_name,
                logger=typed_ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            typed_module._show_system_error_toast(2)

    def _work() -> Path:
        token.raise_if_cancelled()
        typed_ns["_validate_uploaded_filled_marks_workbook"](source_path)
        token.raise_if_cancelled()
        if workflow_service is not None and job_context is not None:
            return workflow_service.generate_final_report(
                source_path,
                save_path,
                context=job_context,
                cancel_token=token,
            )
        return typed_ns["_atomic_copy_file"](source_path, save_path)

    typed_ns["_start_async_operation_compat"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )
