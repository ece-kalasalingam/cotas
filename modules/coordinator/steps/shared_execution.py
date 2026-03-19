"""Shared execution helpers for coordinator step actions."""

from __future__ import annotations

from typing import Any, Mapping

from common.exceptions import JobCancelledError


def handle_step_failure(
    *,
    exc: Exception,
    ns: Mapping[str, Any],
    module: object,
    process_name: str,
    job_id: str,
    step_id: str,
    failed_status_key: str = "coordinator.status.processing_failed",
    cancelled_status_key: str = "coordinator.status.operation_cancelled",
) -> bool:
    t = ns["t"]
    cancelled_error = ns.get("JobCancelledError", JobCancelledError)
    if isinstance(exc, cancelled_error):
        module._publish_status_key(cancelled_status_key)  # type: ignore[attr-defined]
        module._logger.info(  # type: ignore[attr-defined]
            "%s cancelled by user/system request.",
            process_name,
            extra={
                "user_message": ns["build_i18n_log_message"](
                    cancelled_status_key,
                    fallback=t(cancelled_status_key),
                ),
                "job_id": job_id,
                "step_id": step_id,
            },
        )
        return True

    ns["log_process_message"](
        process_name,
        logger=module._logger,  # type: ignore[attr-defined]
        error=exc,
        user_error_message=ns["build_i18n_log_message"](
            failed_status_key,
            fallback=t(failed_status_key),
        ),
        job_id=job_id,
        step_id=step_id,
    )
    ns["show_toast"](
        module,
        t(failed_status_key),
        title=t("coordinator.title"),
        level="error",
    )
    return True
