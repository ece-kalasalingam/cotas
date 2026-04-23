"""Shared runtime orchestration for UI modules."""

from __future__ import annotations

from collections.abc import Callable, Iterable
from contextlib import contextmanager
from logging import Logger
from typing import Mapping

from common.exceptions import ConfigurationError
from common.jobs import CancellationToken
from common.module_messages import (
    NotificationChannel,
    ToastLevel,
    emit_validation_batch_feedback,
    emit_workbook_generation_feedback,
    notify_message,
    notify_message_key,
    notify_validation_issue,
    publish_status_key,
    setup_ui_logging,
    show_toast_key,
    show_toast_plain,
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
        """Init.
        
        Args:
            module: Parameter value (object).
            app_name: Parameter value (str).
            logger: Parameter value (Logger).
            async_runner: Parameter value (object).
            messages_namespace_factory: Parameter value (Callable[[], Mapping[str, object]]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._module = module
        self._app_name = app_name
        self._logger = logger
        self._async_runner = async_runner
        self._messages_namespace_factory = messages_namespace_factory

    def remember_dialog_dir_safe(self, selected_path: str) -> None:
        """Remember dialog dir safe.
        
        Args:
            selected_path: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        remember_dialog_dir_safe(
            selected_path,
            app_name=self._app_name,
            logger=self._logger,
        )

    def set_global_processing_active(self, active: bool) -> None:
        """Set global processing active.
        
        Args:
            active: Parameter value (bool).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        host_candidates: list[object] = [self._module]
        window_getter = getattr(self._module, "window", None)
        if callable(window_getter):
            try:
                window_obj = window_getter()
            except Exception:
                window_obj = None
            if window_obj is not None:
                host_candidates.append(window_obj)
        for candidate in host_candidates:
            setter = getattr(candidate, "set_global_processing_active", None)
            if callable(setter):
                setter(active)
                break

    @contextmanager
    def processing_indicator(self):
        """Processing indicator.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.set_global_processing_active(True)
        try:
            yield
        finally:
            self.set_global_processing_active(False)

    def set_async_runner(self, async_runner: object) -> None:
        """Set async runner.
        
        Args:
            async_runner: Parameter value (object).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._async_runner = async_runner

    def setup_ui_logging(self) -> None:
        """Setup ui logging.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        setup_ui_logging(self._module, ns=self._messages_namespace_factory())

    def append_user_log(self, message: str) -> None:
        """Append user log.
        
        Args:
            message: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        del message
        raise ConfigurationError(
            "ModuleRuntime.append_user_log(message) is disallowed. "
            "Use notify_message_key(...) / publish_status_key(...) for translatable log entries."
        )

    def publish_status(self, message: str) -> None:
        """Publish status.
        
        Args:
            message: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        del message
        raise ConfigurationError(
            "ModuleRuntime.publish_status(message) is disallowed. "
            "Use publish_status_key(...) or notify_message_key(...)."
        )

    def publish_status_key(self, text_key: str, **kwargs: object) -> None:
        """Publish status key.
        
        Args:
            text_key: Parameter value (str).
            kwargs: Parameter value (object).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Notify message.
        
        Args:
            message: Parameter value (str).
            channels: Parameter value (tuple[NotificationChannel, ...]).
            toast_title: Parameter value (str).
            toast_level: Parameter value (ToastLevel).
            toast_duration_ms: Parameter value (int | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Notify message key.
        
        Args:
            text_key: Parameter value (str).
            channels: Parameter value (tuple[NotificationChannel, ...]).
            kwargs: Parameter value (Mapping[str, object] | None).
            fallback: Parameter value (str | None).
            toast_title: Parameter value (str).
            toast_title_key: Parameter value (str | None).
            toast_title_kwargs: Parameter value (Mapping[str, object] | None).
            toast_level: Parameter value (ToastLevel).
            toast_duration_ms: Parameter value (int | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Notify validation issue.
        
        Args:
            issue: Parameter value (Mapping[str, object]).
            file_path: Parameter value (str | None).
            channels: Parameter value (tuple[NotificationChannel, ...]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Emit validation batch feedback.
        
        Args:
            rejections: Parameter value (Iterable[Mapping[str, object]]).
            valid_count: Parameter value (int).
            issue_channels: Parameter value (tuple[NotificationChannel, ...]).
            summary_channels: Parameter value (tuple[NotificationChannel, ...]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Emit workbook generation feedback.
        
        Args:
            success_count: Parameter value (int).
            failed_count: Parameter value (int).
            channels: Parameter value (tuple[NotificationChannel, ...]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Show toast key.
        
        Args:
            text_key: Parameter value (str).
            title_key: Parameter value (str).
            translate: Parameter value (Callable[..., str]).
            level: Parameter value (ToastLevel).
            text_kwargs: Parameter value (Mapping[str, object] | None).
            title_kwargs: Parameter value (Mapping[str, object] | None).
            duration_ms: Parameter value (int | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Show toast plain.
        
        Args:
            message: Parameter value (str).
            title: Parameter value (str).
            level: Parameter value (ToastLevel).
            duration_ms: Parameter value (int | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        """Start async operation.
        
        Args:
            token: Parameter value (CancellationToken).
            job_id: Parameter value (str | None).
            work: Parameter value (Callable[[], object]).
            on_success: Parameter value (Callable[[object], None]).
            on_failure: Parameter value (Callable[[Exception], None]).
            on_finally: Parameter value (Callable[[], None] | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        start = getattr(self._async_runner, "start")
        start(
            token=token,
            job_id=job_id,
            work=work,
            on_success=on_success,
            on_failure=on_failure,
            on_finally=on_finally,
        )
