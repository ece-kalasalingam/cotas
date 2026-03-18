"""Step 2: upload course details and prepare marks template."""

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
    WORKFLOW_STEP_ID_STEP2_GENERATE_MARKS_TEMPLATE,
    WORKFLOW_STEP_ID_STEP2_VALIDATE_COURSE_DETAILS,
)


class _ModuleState(Protocol):
    busy: bool


class _Logger(Protocol):
    def info(self, msg: str, *args: object, **kwargs: object) -> None:
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


class _InstructorStep2Module(Protocol):
    state: _ModuleState
    _logger: _Logger
    marks_template_done: bool
    step2_upload_ready: bool
    step2_course_details_path: str | None
    marks_template_path: str | None
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
    _publish_status_compat: Callable[..., None]
    _start_async_operation_compat: _StartAsyncOperation
    log_process_message: Callable[..., None]
    build_i18n_log_message: Callable[..., str]
    show_toast: Callable[..., None]
    _logger: _Logger
    validate_course_details_workbook: Callable[[str], None]
    _build_marks_template_default_name: Callable[[str], str]
    generate_marks_template_from_course_details: Callable[[str, str], None]


def upload_course_details_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep2Module, module)
    typed_ns = cast(_Step2Namespace, ns)
    if typed_module.state.busy:
        return

    t = typed_ns["t"]
    process_name = t("instructor.log.process.validate_course_details_workbook")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.validate_course_details_workbook"
    )
    open_path, _ = typed_ns["QFileDialog"].getOpenFileName(
        typed_module,
        t("instructor.dialog.step2.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"]),
        t("instructor.dialog.filter.excel_open"),
    )
    if not open_path:
        return
    typed_module._remember_dialog_dir_safe(open_path)

    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_VALIDATE_COURSE_DETAILS,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: open_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(result: object) -> None:
        replacing = typed_module.marks_template_done or typed_module.step2_upload_ready
        typed_module.step2_course_details_path = open_path
        typed_module.step2_upload_ready = True
        typed_module.marks_template_done = False
        typed_module.marks_template_path = None
        fallback_name = t("instructor.dialog.step1.prepare.default_name")
        default_marks_name_obj = (
            cast(dict[str, object], result).get("default_marks_name")
            if isinstance(result, dict)
            else fallback_name
        )
        typed_module._step2_marks_default_name = (
            default_marks_name_obj if isinstance(default_marks_name_obj, str) else fallback_name
        ) or fallback_name

        if replacing:
            typed_module.final_report_outdated = typed_module.final_report_done
            typed_module.filled_marks_outdated = typed_module.filled_marks_done
            if typed_module.filled_marks_outdated or typed_module.final_report_outdated:
                typed_ns["_publish_status_compat"](typed_module, t("instructor.status.step1_changed"))
        else:
            typed_ns["_publish_status_compat"](typed_module, t("instructor.status.step1_validated"))
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        typed_module._show_step_success_toast(1)

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
            typed_module._show_system_error_toast(1)
        else:
            typed_ns["log_process_message"](
                process_name,
                logger=typed_ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            typed_module._show_system_error_toast(1)

    def _work() -> dict[str, str]:
        if workflow_service is not None and job_context is not None:
            workflow_service.validate_course_details_workbook(
                open_path,
                context=job_context,
                cancel_token=token,
            )
        else:
            typed_ns["validate_course_details_workbook"](open_path)
        token.raise_if_cancelled()
        return {
            "default_marks_name": typed_ns["_build_marks_template_default_name"](open_path),
        }

    typed_ns["_start_async_operation_compat"](
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
    if not typed_module.step2_upload_ready or not typed_module.step2_course_details_path:
        typed_ns["show_toast"](
            typed_module,
            t("instructor.require.step1"),
            title=t("instructor.msg.step_required_title"),
            level="info",
        )
        return

    source_path = typed_module.step2_course_details_path
    default_name = getattr(
        typed_module,
        "_step2_marks_default_name",
        t("instructor.dialog.step1.prepare.default_name"),
    ) or t("instructor.dialog.step1.prepare.default_name")
    save_path, _ = typed_ns["QFileDialog"].getSaveFileName(
        typed_module,
        t("instructor.dialog.step1.prepare.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"], default_name),
        t("instructor.dialog.filter.excel"),
    )
    if not save_path:
        return
    typed_module._remember_dialog_dir_safe(save_path)

    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_GENERATE_MARKS_TEMPLATE,
            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: source_path, WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        typed_module.marks_template_path = save_path
        typed_module.marks_template_done = True
        typed_module.filled_marks_outdated = typed_module.filled_marks_done
        typed_module.final_report_outdated = typed_module.final_report_done
        typed_ns["_publish_status_compat"](typed_module, t("instructor.status.step1_prepared"))
        typed_ns["log_process_message"](
            process_name,
            logger=typed_ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        typed_module._show_step_success_toast(1)

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
            typed_module._show_system_error_toast(1)
        else:
            typed_ns["log_process_message"](
                process_name,
                logger=typed_ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            typed_module._show_system_error_toast(1)

    def _work() -> Path:
        if workflow_service is not None and job_context is not None:
            return workflow_service.generate_marks_template(
                source_path,
                save_path,
                context=job_context,
                cancel_token=token,
            )
        typed_ns["generate_marks_template_from_course_details"](source_path, save_path)
        return Path(save_path)

    typed_ns["_start_async_operation_compat"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )
