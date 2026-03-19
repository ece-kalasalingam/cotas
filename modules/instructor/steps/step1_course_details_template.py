"""Step 1: generate and save course details template."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Mapping, Protocol, TypedDict, cast

from common.constants import (
    PROCESS_MESSAGE_SUCCESS_SUFFIX,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_TEMPLATE_ID,
    WORKFLOW_STEP_ID_STEP1_GENERATE_COURSE_TEMPLATE,
)
from modules.instructor.steps.shared_execution import handle_step_failure


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


class _CancellationToken(Protocol):
    def raise_if_cancelled(self) -> None:
        ...


class _JobContext(Protocol):
    job_id: str | None
    step_id: str | None


class _InstructorStep1Module(Protocol):
    state: _ModuleState
    _logger: _Logger
    step1_path: str | None
    step1_done: bool

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
        module: _InstructorStep1Module,
        *,
        token: _CancellationToken,
        job_id: str | None,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
    ) -> None:
        ...


class _Step1Namespace(TypedDict):
    ID_COURSE_SETUP: str
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
    log_process_message: Callable[..., None]
    _logger: _Logger
    build_i18n_log_message: Callable[..., str]
    generate_course_details_template: Callable[..., None]
    _start_async_operation: _StartAsyncOperation


def download_course_template_async(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_InstructorStep1Module, module)
    typed_ns = cast(_Step1Namespace, ns)
    if typed_module.state.busy:
        return

    template_id = typed_ns["ID_COURSE_SETUP"]
    t = typed_ns["t"]
    process_name = t("instructor.log.process.generate_course_details_template")
    user_success_message, user_error_message = typed_ns["_localized_log_messages"](
        "instructor.log.process.generate_course_details_template"
    )

    save_path, _ = typed_ns["QFileDialog"].getSaveFileName(
        typed_module,
        t("instructor.dialog.step1.title"),
        typed_ns["resolve_dialog_start_path"](typed_ns["APP_NAME"], t("instructor.dialog.step1.default_name")),
        t("instructor.dialog.filter.excel"),
    )
    if not save_path:
        return
    typed_module._remember_dialog_dir_safe(save_path)

    workflow_service = cast(Any, getattr(typed_module, "_workflow_service", None))
    token = typed_ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP1_GENERATE_COURSE_TEMPLATE,
            payload={WORKFLOW_PAYLOAD_KEY_TEMPLATE_ID: template_id, WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        typed_module.step1_path = save_path
        typed_module.step1_done = True
        typed_ns["_publish_status"](typed_module, t("instructor.status.step1_selected"))
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

    def _work() -> Path:
        if workflow_service is not None and job_context is not None:
            return workflow_service.generate_course_details_template(
                save_path,
                context=job_context,
                cancel_token=token,
            )
        typed_ns["generate_course_details_template"](save_path, template_id=template_id)
        return Path(save_path)

    typed_ns["_start_async_operation"](
        typed_module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )
