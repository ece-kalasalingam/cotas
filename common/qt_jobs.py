"""Qt background job helpers for non-blocking processing."""

from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal, Slot


class JobSignals(QObject):
    finished = Signal(object)
    failed = Signal(Exception)


class FunctionJob(QRunnable):
    def __init__(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> None:
        super().__init__()
        self._fn = fn
        self._args = args
        self._kwargs = kwargs
        self.signals = JobSignals()

    @Slot()
    def run(self) -> None:
        try:
            result = self._fn(*self._args, **self._kwargs)
        except Exception as exc:
            self.signals.failed.emit(exc)
            return
        self.signals.finished.emit(result)


def run_in_background(
    fn: Callable[..., Any],
    *args: Any,
    on_finished: Callable[[Any], None] | None = None,
    on_failed: Callable[[Exception], None] | None = None,
    thread_pool: QThreadPool | None = None,
    **kwargs: Any,
) -> FunctionJob:
    pool = thread_pool or QThreadPool.globalInstance()
    job = FunctionJob(fn, *args, **kwargs)
    if on_finished is not None:
        job.signals.finished.connect(on_finished)
    if on_failed is not None:
        job.signals.failed.connect(on_failed)
    pool.start(job)
    return job
