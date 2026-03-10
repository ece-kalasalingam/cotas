"""Step 2: upload course details and prepare marks template."""

from __future__ import annotations

from pathlib import Path
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


def upload_course_details_async(module: object, *, ns: dict[str, object]) -> None:
    if module.state.busy:
        return

    t = ns["t"]
    process_name = t("instructor.log.process.validate_course_details_workbook")
    user_success_message, user_error_message = ns["_localized_log_messages"](
        "instructor.log.process.validate_course_details_workbook"
    )
    open_path, _ = ns["QFileDialog"].getOpenFileName(
        module,
        t("instructor.dialog.step2.title"),
        ns["resolve_dialog_start_path"](ns["APP_NAME"]),
        t("instructor.dialog.filter.excel_open"),
    )
    if not open_path:
        return
    module._remember_dialog_dir_safe(open_path)

    workflow_service = getattr(module, "_workflow_service", None)
    token = ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_VALIDATE_COURSE_DETAILS,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: open_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(result: object) -> None:
        replacing = module.step2_done or module.step2_upload_ready
        module.step2_course_details_path = open_path
        module.step2_upload_ready = True
        module.step2_done = False
        module.step2_path = None
        module._step2_marks_default_name = (
            result.get("default_marks_name")
            if isinstance(result, dict)
            else t("instructor.dialog.step3.default_name")
        ) or t("instructor.dialog.step3.default_name")

        if replacing:
            module.step4_outdated = module.step4_done
            module.step3_outdated = module.step3_done
            if module.step3_outdated or module.step4_outdated:
                ns["_publish_status_compat"](module, t("instructor.status.step2_changed"))
        else:
            ns["_publish_status_compat"](module, t("instructor.status.step2_validated"))
        ns["log_process_message"](
            process_name,
            logger=ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        module._show_step_success_toast(2)

    def _on_failed(exc: Exception) -> None:
        if isinstance(exc, ns["JobCancelledError"]):
            user_message = t("instructor.status.operation_cancelled")
            ns["_publish_status_compat"](module, user_message)
            ns["_logger"].info(
                PROCESS_MESSAGE_CANCELLED_TEMPLATE,
                process_name,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: user_message,
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
            module._show_system_error_toast(2)
        else:
            ns["log_process_message"](
                process_name,
                logger=ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            module._show_system_error_toast(2)

    def _work() -> dict[str, str]:
        if workflow_service is not None and job_context is not None:
            workflow_service.validate_course_details_workbook(
                open_path,
                context=job_context,
                cancel_token=token,
            )
        else:
            ns["validate_course_details_workbook"](open_path)
        token.raise_if_cancelled()
        return {
            "default_marks_name": ns["_build_marks_template_default_name"](open_path),
        }

    ns["_start_async_operation_compat"](
        module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )


def prepare_marks_template_async(module: object, *, ns: dict[str, object]) -> None:
    if module.state.busy:
        return

    t = ns["t"]
    process_name = t("instructor.log.process.generate_marks_template")
    user_success_message, user_error_message = ns["_localized_log_messages"](
        "instructor.log.process.generate_marks_template"
    )
    if not module.step2_upload_ready or not module.step2_course_details_path:
        ns["show_toast"](
            module,
            t("instructor.require.step2"),
            title=t("instructor.msg.step_required_title"),
            level="info",
        )
        return

    source_path = module.step2_course_details_path
    default_name = getattr(
        module,
        "_step2_marks_default_name",
        t("instructor.dialog.step3.default_name"),
    ) or t("instructor.dialog.step3.default_name")
    save_path, _ = ns["QFileDialog"].getSaveFileName(
        module,
        t("instructor.dialog.step3.title"),
        ns["resolve_dialog_start_path"](ns["APP_NAME"], default_name),
        t("instructor.dialog.filter.excel"),
    )
    if not save_path:
        return
    module._remember_dialog_dir_safe(save_path)

    workflow_service = getattr(module, "_workflow_service", None)
    token = ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP2_GENERATE_MARKS_TEMPLATE,
            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: source_path, WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        module.step2_path = save_path
        module.step2_done = True
        module.step3_outdated = module.step3_done
        module.step4_outdated = module.step4_done
        ns["_publish_status_compat"](module, t("instructor.status.step2_uploaded"))
        ns["log_process_message"](
            process_name,
            logger=ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        module._show_step_success_toast(2)

    def _on_failed(exc: Exception) -> None:
        if isinstance(exc, ns["JobCancelledError"]):
            user_message = t("instructor.status.operation_cancelled")
            ns["_publish_status_compat"](module, user_message)
            ns["_logger"].info(
                PROCESS_MESSAGE_CANCELLED_TEMPLATE,
                process_name,
                extra={
                    LOG_EXTRA_KEY_USER_MESSAGE: user_message,
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
            module._show_system_error_toast(2)
        else:
            ns["log_process_message"](
                process_name,
                logger=ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            module._show_system_error_toast(2)

    def _work() -> Path:
        if workflow_service is not None and job_context is not None:
            return workflow_service.generate_marks_template(
                source_path,
                save_path,
                context=job_context,
                cancel_token=token,
            )
        ns["generate_marks_template_from_course_details"](source_path, save_path)
        return Path(save_path)

    ns["_start_async_operation_compat"](
        module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )
