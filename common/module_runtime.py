"""Shared runtime orchestration for UI modules."""

from __future__ import annotations

from collections.abc import Callable
from logging import Logger
from collections.abc import Iterable
from typing import Mapping

from common.jobs import CancellationToken
from common.exceptions import ConfigurationError
from common.module_messages import (
    NotificationChannel,
    ToastLevel,
    emit_workbook_generation_feedback,
    emit_validation_batch_feedback,
    notify_validation_issue,
    notify_message,
    notify_message_key,
    publish_status_key,
    show_toast_key,
    show_toast_plain,
    setup_ui_logging,
)
from common.utils import remember_dialog_dir_safe


class ModuleRuntime:
    """Common module runtime helpers for status/logging/async and dialog-dir persistence."""

    def __init__(
        self,
        *,
        module: object,
        app_name: str,
        logger: Logger,
        async_runner: object,
        messages_namespace_factory: Callable[[], Mapping[str, object]],
    ) -> None:
        self._module = module
        self._app_name = app_name
        self._logger = logger
        self._async_runner = async_runner
        self._messages_namespace_factory = messages_namespace_factory

    def remember_dialog_dir_safe(self, selected_path: str) -> None:
        remember_dialog_dir_safe(
            selected_path,
            app_name=self._app_name,
            logger=self._logger,
        )

    def set_async_runner(self, async_runner: object) -> None:
        self._async_runner = async_runner

    def setup_ui_logging(self) -> None:
        setup_ui_logging(self._module, ns=self._messages_namespace_factory())

    def append_user_log(self, message: str) -> None:
        del message
        raise ConfigurationError(
            "ModuleRuntime.append_user_log(message) is disallowed. "
            "Use notify_message_key(...) / publish_status_key(...) for translatable log entries."
        )

    def publish_status(self, message: str) -> None:
        del message
        raise ConfigurationError(
            "ModuleRuntime.publish_status(message) is disallowed. "
            "Use publish_status_key(...) or notify_message_key(...)."
        )

    def publish_status_key(self, text_key: str, **kwargs: object) -> None:
        publish_status_key(self._module, text_key, ns=self._messages_namespace_factory(), **kwargs)

    def notify_message(
        self,
        message: str,
        *,
        channels: tuple[NotificationChannel, ...] = ("status", "activity_log"),
        toast_title: str = "",
        toast_level: ToastLevel = "info",
        toast_duration_ms: int | None = None,
    ) -> None:
        notify_message(
            self._module,
            message,
            ns=self._messages_namespace_factory(),
            channels=channels,
            toast_title=toast_title,
            toast_level=toast_level,
            toast_duration_ms=toast_duration_ms,
        )

    def notify_message_key(
        self,
        text_key: str,
        *,
        channels: tuple[NotificationChannel, ...] = ("status", "activity_log"),
        kwargs: Mapping[str, object] | None = None,
        fallback: str | None = None,
        toast_title: str = "",
        toast_title_key: str | None = None,
        toast_title_kwargs: Mapping[str, object] | None = None,
        toast_level: ToastLevel = "info",
        toast_duration_ms: int | None = None,
    ) -> None:
        notify_message_key(
            self._module,
            text_key,
            ns=self._messages_namespace_factory(),
            channels=channels,
            kwargs=kwargs,
            fallback=fallback,
            toast_title=toast_title,
            toast_title_key=toast_title_key,
            toast_title_kwargs=toast_title_kwargs,
            toast_level=toast_level,
            toast_duration_ms=toast_duration_ms,
        )

    def notify_validation_issue(
        self,
        issue: Mapping[str, object],
        *,
        file_path: str | None = None,
        channels: tuple[NotificationChannel, ...] = ("activity_log",),
    ) -> None:
        notify_validation_issue(
            self._module,
            ns=self._messages_namespace_factory(),
            issue=issue,
            file_path=file_path,
            channels=channels,
        )

    def emit_validation_batch_feedback(
        self,
        *,
        rejections: Iterable[Mapping[str, object]],
        valid_count: int,
        issue_channels: tuple[NotificationChannel, ...] = ("status", "activity_log"),
        summary_channels: tuple[NotificationChannel, ...] = ("toast",),
    ) -> None:
        emit_validation_batch_feedback(
            self._module,
            ns=self._messages_namespace_factory(),
            rejections=rejections,
            valid_count=valid_count,
            issue_channels=issue_channels,
            summary_channels=summary_channels,
        )

    def emit_workbook_generation_feedback(
        self,
        *,
        success_count: int,
        failed_count: int,
        channels: tuple[NotificationChannel, ...] = ("status", "activity_log", "toast"),
    ) -> None:
        emit_workbook_generation_feedback(
            self._module,
            ns=self._messages_namespace_factory(),
            success_count=success_count,
            failed_count=failed_count,
            channels=channels,
        )

    def show_toast_key(
        self,
        *,
        text_key: str,
        title_key: str,
        translate: Callable[..., str],
        level: ToastLevel = "info",
        text_kwargs: Mapping[str, object] | None = None,
        title_kwargs: Mapping[str, object] | None = None,
        duration_ms: int | None = None,
    ) -> None:
        show_toast_key(
            self._module,
            text_key=text_key,
            title_key=title_key,
            translate=translate,
            level=level,
            text_kwargs=text_kwargs,
            title_kwargs=title_kwargs,
            duration_ms=duration_ms,
        )

    def show_toast_plain(
        self,
        message: str,
        *,
        title: str,
        level: ToastLevel = "info",
        duration_ms: int | None = None,
    ) -> None:
        show_toast_plain(
            self._module,
            message,
            title=title,
            level=level,
            duration_ms=duration_ms,
        )

    def start_async_operation(
        self,
        *,
        token: CancellationToken,
        job_id: str | None,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_failure: Callable[[Exception], None],
        on_finally: Callable[[], None] | None = None,
    ) -> None:
        start = getattr(self._async_runner, "start")
        start(
            token=token,
            job_id=job_id,
            work=work,
            on_success=on_success,
            on_failure=on_failure,
            on_finally=on_finally,
        )
