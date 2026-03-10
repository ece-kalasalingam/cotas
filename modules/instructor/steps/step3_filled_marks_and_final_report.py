"""Step 3: upload filled marks and generate final report."""

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
    WORKFLOW_STEP_ID_STEP3_GENERATE_FINAL_REPORT,
    WORKFLOW_STEP_ID_STEP3_UPLOAD_FILLED_MARKS,
)

_LOG_STEP4_SOURCE_MISSING = "Step 4 failed: Step 3 file is missing. step3_path=%s"


def upload_filled_marks_async(module: object, *, ns: dict[str, object]) -> None:
    if module.state.busy:
        return

    t = ns["t"]
    process_name = t("instructor.log.process.upload_filled_marks_workbook")
    user_success_message, user_error_message = ns["_localized_log_messages"](
        "instructor.log.process.upload_filled_marks_workbook"
    )
    open_path, _ = ns["QFileDialog"].getOpenFileName(
        module,
        t("instructor.dialog.step3.title"),
        ns["resolve_dialog_start_path"](ns["APP_NAME"]),
        t("instructor.dialog.filter.excel_open"),
    )
    if not open_path:
        return
    module._remember_dialog_dir_safe(open_path)

    token = ns["CancellationToken"]()
    workflow_service = getattr(module, "_workflow_service", None)
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP3_UPLOAD_FILLED_MARKS,
            payload={WORKFLOW_PAYLOAD_KEY_PATH: open_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        replacing = module.step3_done
        module.step3_path = open_path
        module.step3_done = True
        module.step3_outdated = False

        if replacing and module.step3_done:
            module.step4_outdated = True
            ns["_publish_status_compat"](module, t("instructor.status.step3_changed"))
        else:
            ns["_publish_status_compat"](module, t("instructor.status.step3_uploaded"))
        ns["log_process_message"](
            process_name,
            logger=ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        module._show_step_success_toast(3)

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
            module._show_system_error_toast(3)
        else:
            ns["log_process_message"](
                process_name,
                logger=ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            module._show_system_error_toast(3)

    def _work() -> bool:
        token.raise_if_cancelled()
        ns["_validate_uploaded_filled_marks_workbook"](open_path)
        token.raise_if_cancelled()
        return True

    ns["_start_async_operation_compat"](
        module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )


def generate_final_report_async(module: object, *, ns: dict[str, object]) -> None:
    if module.state.busy:
        return

    t = ns["t"]
    process_name = t("instructor.log.process.generate_final_co_report")
    user_success_message, user_error_message = ns["_localized_log_messages"](
        "instructor.log.process.generate_final_co_report"
    )
    can_run, reason = module._can_run_step(3)
    if not can_run:
        ns["show_toast"](
            module,
            reason,
            title=t("instructor.msg.step_required_title"),
            level="info",
        )
        return
    if not module.step3_done or module.step3_outdated:
        ns["show_toast"](
            module,
            t("instructor.require.step3"),
            title=t("instructor.msg.step_required_title"),
            level="info",
        )
        return

    default_name = ns["_build_final_report_default_name"](module.step3_path)
    save_path, _ = ns["QFileDialog"].getSaveFileName(
        module,
        t("instructor.dialog.step4.title"),
        ns["resolve_dialog_start_path"](ns["APP_NAME"], default_name),
        t("instructor.dialog.filter.excel"),
    )
    if not save_path:
        return

    if not module.step3_path or not Path(module.step3_path).exists():
        ns["_logger"].warning(_LOG_STEP4_SOURCE_MISSING, module.step3_path)
        ns["show_toast"](
            module,
            t("instructor.require.step3"),
            title=t("instructor.msg.step_required_title"),
            level="error",
        )
        return
    source_path = module.step3_path

    workflow_service = getattr(module, "_workflow_service", None)
    token = ns["CancellationToken"]()
    job_context = (
        workflow_service.create_job_context(
            step_id=WORKFLOW_STEP_ID_STEP3_GENERATE_FINAL_REPORT,
            payload={WORKFLOW_PAYLOAD_KEY_SOURCE: source_path, WORKFLOW_PAYLOAD_KEY_OUTPUT: save_path},
        )
        if workflow_service is not None
        else None
    )

    def _on_finished(_result: object) -> None:
        module.step4_path = save_path
        module.step4_done = True
        module.step4_outdated = False
        module._remember_dialog_dir_safe(save_path)
        ns["_publish_status_compat"](module, t("instructor.status.step4_selected"))
        ns["log_process_message"](
            process_name,
            logger=ns["_logger"],
            success_message=f"{process_name}{PROCESS_MESSAGE_SUCCESS_SUFFIX}",
            user_success_message=user_success_message,
            job_id=job_context.job_id if job_context else None,
            step_id=job_context.step_id if job_context else None,
        )
        module._show_step_success_toast(3)

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
        else:
            ns["log_process_message"](
                process_name,
                logger=ns["_logger"],
                error=exc,
                user_error_message=user_error_message,
                job_id=job_context.job_id if job_context else None,
                step_id=job_context.step_id if job_context else None,
            )
            module._show_system_error_toast(3)

    def _work() -> Path:
        token.raise_if_cancelled()
        ns["_validate_uploaded_filled_marks_workbook"](source_path)
        token.raise_if_cancelled()
        if workflow_service is not None and job_context is not None:
            return workflow_service.generate_final_report(
                source_path,
                save_path,
                context=job_context,
                cancel_token=token,
            )
        return ns["_atomic_copy_file"](source_path, save_path)

    ns["_start_async_operation_compat"](
        module,
        token=token,
        job_id=job_context.job_id if job_context else None,
        work=_work,
        on_success=_on_finished,
        on_failure=_on_failed,
    )
