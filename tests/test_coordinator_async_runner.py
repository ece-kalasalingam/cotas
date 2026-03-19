from __future__ import annotations

from typing import Callable, cast

from common.async_operation_runner import AsyncOperationRunner
from common.exceptions import JobCancelledError
from common.jobs import CancellationToken


class _Target:
    def __init__(self) -> None:
        self._cancel_token = None
        self._active_jobs = []
        self.busy_calls: list[tuple[bool, str | None]] = []

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.busy_calls.append((busy, job_id))


def test_async_runner_success_path_finalizes_and_calls_hooks() -> None:
    target = _Target()
    finished: list[object] = []
    finally_calls = {"count": 0}

    callbacks: dict[str, object] = {}

    def _run_async(work, *, on_finished, on_failed):
        callbacks["work"] = work
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    runner = AsyncOperationRunner(target, run_async=_run_async)
    runner.start(
        token=CancellationToken(),
        job_id="job-1",
        work=lambda: 123,
        on_success=lambda result: finished.append(result),
        on_failure=lambda exc: (_ for _ in ()).throw(exc),
        on_finally=lambda: finally_calls.__setitem__("count", finally_calls["count"] + 1),
    )

    assert len(target._active_jobs) == 1
    cast(Callable[[object], None], callbacks["on_finished"])(cast(Callable[[], object], callbacks["work"])())

    assert finished == [123]
    assert finally_calls["count"] == 1
    assert target._cancel_token is None
    assert target._active_jobs == []
    assert target.busy_calls[0] == (True, "job-1")
    assert target.busy_calls[-1] == (False, None)


def test_async_runner_failure_path_finalizes_and_calls_failure() -> None:
    target = _Target()
    failures: list[Exception] = []

    callbacks: dict[str, object] = {}

    def _run_async(_work, *, on_finished, on_failed):
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    runner = AsyncOperationRunner(target, run_async=_run_async)
    runner.start(
        token=CancellationToken(),
        job_id="job-2",
        work=lambda: 999,
        on_success=lambda _result: None,
        on_failure=lambda exc: failures.append(exc),
    )

    assert len(target._active_jobs) == 1
    cast(Callable[[Exception], None], callbacks["on_failed"])(JobCancelledError("cancelled"))

    assert len(failures) == 1
    assert isinstance(failures[0], JobCancelledError)
    assert target._cancel_token is None
    assert target._active_jobs == []
    assert target.busy_calls[0] == (True, "job-2")
    assert target.busy_calls[-1] == (False, None)
