"""Shared execution helpers for instructor step actions."""

from __future__ import annotations

from typing import Any, Callable, Mapping

from common.constants import (
    LOG_EXTRA_KEY_JOB_ID,
    LOG_EXTRA_KEY_STEP_ID,
    LOG_EXTRA_KEY_USER_MESSAGE,
    PROCESS_MESSAGE_CANCELLED_TEMPLATE,
)


def handle_step_failure(
    *,
    exc: Exception,
    ns: Mapping[str, Any],
    module: object,
    process_name: str,
    user_error_message: str,
    step_no: int,
    job_id: str | None,
    step_id: str | None,
    show_validation_toast: Callable[[str], None] | None = None,
) -> bool:
    """Handle common step failure branches.

    Returns True when the exception path was fully handled by this helper.
    """

    t = ns["t"]
    if isinstance(exc, ns["JobCancelledError"]):
        status_key = "instructor.status.operation_cancelled"
        user_message = t(status_key)
        user_message_payload = ns["build_i18n_log_message"](status_key, fallback=user_message)
        publish_key = ns.get("_publish_status_key")
        if callable(publish_key):
            publish_key(module, status_key)
        else:
            ns["_publish_status"](module, user_message)
        ns["_logger"].info(
            PROCESS_MESSAGE_CANCELLED_TEMPLATE,
            process_name,
            extra={
                LOG_EXTRA_KEY_USER_MESSAGE: user_message_payload,
                LOG_EXTRA_KEY_JOB_ID: job_id,
                LOG_EXTRA_KEY_STEP_ID: step_id,
            },
        )
        return True

    if isinstance(exc, ns["ValidationError"]) and callable(show_validation_toast):
        ns["log_process_message"](
            process_name,
            logger=ns["_logger"],
            error=exc,
            notify=lambda message, _level: show_validation_toast(message),
            job_id=job_id,
            step_id=step_id,
        )
        return True

    ns["log_process_message"](
        process_name,
        logger=ns["_logger"],
        error=exc,
        user_error_message=user_error_message,
        job_id=job_id,
        step_id=step_id,
    )
    if isinstance(exc, ns["AppSystemError"]) or not isinstance(exc, ns["ValidationError"]) or (
        isinstance(exc, ns["ValidationError"]) and not callable(show_validation_toast)
    ):
        module._show_system_error_toast(step_no)  # type: ignore[attr-defined]
    return True
