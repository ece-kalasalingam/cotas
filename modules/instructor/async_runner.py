"""Reusable async runner utilities for instructor workflows."""

from __future__ import annotations

from collections.abc import Callable

from common.async_operation_runner import (
    AsyncOperationRunner as _SharedAsyncOperationRunner,
)


class AsyncOperationRunner(_SharedAsyncOperationRunner):
    """Instructor runner with guarded UI refresh semantics."""

    def __init__(self, target: object, *, run_async: Callable[..., object]) -> None:
        super().__init__(
            target,
            run_async=run_async,
            refresh_ui=lambda: target._refresh_ui(),  # type: ignore[attr-defined]
            should_refresh_ui=lambda: not bool(getattr(target, "_is_closing", False)),
        )
