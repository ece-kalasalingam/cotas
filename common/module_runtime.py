"""Shared runtime orchestration for UI modules."""

from __future__ import annotations

from collections.abc import Callable
from logging import Logger
from typing import Mapping

from common.jobs import CancellationToken
from common.module_messages import (
    append_user_log,
    publish_status,
    publish_status_key,
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
        append_user_log(self._module, message, ns=self._messages_namespace_factory())

    def publish_status(self, message: str) -> None:
        publish_status(self._module, message, ns=self._messages_namespace_factory())

    def publish_status_key(self, text_key: str, **kwargs: object) -> None:
        publish_status_key(self._module, text_key, ns=self._messages_namespace_factory(), **kwargs)

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
