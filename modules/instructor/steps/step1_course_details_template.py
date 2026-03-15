"""Step 1: generate and save course details template."""

from __future__ import annotations

from pathlib import Path
from common.constants import (
    LOG_EXTRA_KEY_JOB_ID,
    LOG_EXTRA_KEY_STEP_ID,
    LOG_EXTRA_KEY_USER_MESSAGE,
    PROCESS_MESSAGE_CANCELLED_TEMPLATE,
    PROCESS_MESSAGE_SUCCESS_SUFFIX,
    WORKFLOW_PAYLOAD_KEY_OUTPUT,
    WORKFLOW_PAYLOAD_KEY_TEMPLATE_ID,
    WORKFLOW_STEP_ID_STEP1_GENERATE_COURSE_TEMPLATE,
)


def download_course_template_async(module: object, *, ns: dict[str, object]) -> None:
    if module.state.busy:
        return

    template_id = ns["ID_COURSE_SETUP"]
    t = ns["t"]
    process_name = t("instructor.log.process.generate_course_details_template")
    user_success_message, user_error_message = ns["_localized_log_messages"](
        "instructor.log.process.generate_course_details_template"
    )

    save_path, _ = ns["QFileDialog"].getSaveFileName(
        module,
        t("instructor.dialog.step1.title"),
        ns["resolve_dialog_start_path"](ns["APP_NAME"], t("instructor.dialog.step1.default_name")),
        t("instructor.dialog.filter.excel"),
    )
    if not save_path:
        return
    module._remember_dialog_dir_safe(save_path)

    workflow_service = getattr(module, "_workflow_service", None)
    token = ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP1_GENERATE_COURSE_TEMPLATE,
            payload={WORKFLOW_PAYLOAD_KEY_TEMPLATE_ID: template_id, WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        module.step1_path = save_path
        module.step1_done = True
        ns["_publish_status_compat"](module, t("instructor.status.step1_selected"))
        ns["log_process_message"](
            process_name,
            logger=ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        module._show_step_success_toast(1)

    def _on_failed(exc: Exception) -> None:
        if isinstance(exc, ns["JobCancelledError"]):
            status_key = "instructor.status.operation_cancelled"
            user_message = t(status_key)
            user_message_payload = ns["build_i18n_log_message"](status_key, fallback=user_message)
            ns["_publish_status_compat"](module, user_message)
            ns["_logger"].info(
                PROCESS_MESSAGE_CANCELLED_TEMPLATE,
                process_name,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: user_message_payload,
                    LOG_EXTRA_KEY_JOB_ID: job_context.job_id if job_context else None,
                    LOG_EXTRA_KEY_STEP_ID: job_context.step_id if job_context else None,
                },
            )
            return
        if isinstance(exc, ns["ValidationError"]):
            ns["log_process_message"](
                process_name,
                logger=ns["_logger"],
                error=exc,
                notify=lambda message, _level: module._show_validation_error_toast(message),
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
        elif isinstance(exc, ns["AppSystemError"]):
            ns["log_process_message"](
                process_name,
                logger=ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            module._show_system_error_toast(1)
        else:
            ns["log_process_message"](
                process_name,
                logger=ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            module._show_system_error_toast(1)

    def _work() -> Path:
        if workflow_service is not None and job_context is not None:
            return workflow_service.generate_course_details_template(
                save_path,
                context=job_context,
                cancel_token=token,
            )
        ns["generate_course_details_template"](save_path, template_id=template_id)
        return Path(save_path)

    ns["_start_async_operation_compat"](
        module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )
