"""Shared instructor module message helpers."""

from __future__ import annotations

from typing import cast

from PySide6.QtWidgets import QWidget

from common.texts import t
from common.toast import show_toast
from common.ui_logging import build_i18n_log_message


def localized_log_messages(process_key: str) -> tuple[str, str]:
    user_process_name = t(process_key)
    return (
        build_i18n_log_message(
            "instructor.log.completed_process",
            kwargs={"process": {"__t_key__": process_key}},
            fallback=t("instructor.log.completed_process", process=user_process_name),
        ),
        build_i18n_log_message(
            "instructor.log.error_while_process",
            kwargs={"process": {"__t_key__": process_key}},
            fallback=t("instructor.log.error_while_process", process=user_process_name),
        ),
    )


def show_step_success_toast(widget: object, *, step: int, title_key: str) -> None:
    show_toast(
        cast(QWidget | None, widget),
        t("instructor.msg.step_completed", step=step, title=t(title_key)),
        title=t("instructor.msg.success_title"),
        level="success",
    )


def show_validation_error_toast(widget: object, message: str) -> None:
    show_toast(
        cast(QWidget | None, widget),
        message,
        title=t("instructor.msg.validation_title"),
        level="error",
    )


def show_system_error_toast(widget: object, *, title_key: str) -> None:
    show_toast(
        cast(QWidget | None, widget),
        t("instructor.msg.failed_to_do", action=t(title_key)),
        title=t("instructor.msg.error_title"),
        level="error",
    )
