from __future__ import annotations

from typing import Callable, cast

from common.jobs import CancellationToken
from modules.instructor import async_runner as ir


class _Signal:
    def __init__(self) -> None:
        self.values: list[str] = []

    def emit(self, value: str) -> None:
        self.values.append(value)


class _CompatTarget:
    def __init__(self) -> None:
        self.messages: list[str] = []
        self.busy_calls: list[tuple[bool, str | None]] = []
        self.status_changed = _Signal()
        self._cancel_token = None
        self._active_jobs: list[object] = []
        self.refresh_calls = 0

    def _publish_status(self, message: str) -> None:
        self.messages.append(message)

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.busy_calls.append((busy, job_id))

    def _refresh_ui(self) -> None:
        self.refresh_calls += 1


def test_publish_status_compat_prefers_publish_method() -> None:
    target = _CompatTarget()
    ir.publish_status_compat(target=target, message="hello", logger=object())
    assert target.messages == ["hello"]
    assert target.status_changed.values == []


def test_publish_status_compat_falls_back_to_signal(monkeypatch) -> None:
    class _NoPublish:
        def __init__(self) -> None:
            self.status_changed = object()

    target = _NoPublish()
    seen: list[tuple[object, str, object]] = []
    monkeypatch.setattr(ir, "emit_user_status", lambda sig, message, logger=None: seen.append((sig, message, logger)))

    ir.publish_status_compat(target=target, message="done", logger="L")
    assert seen == [(target.status_changed, "done", "L")]


def test_set_busy_compat_calls_setter_and_noops_without_setter() -> None:
    target = _CompatTarget()
    ir.set_busy_compat(target=target, busy=True, job_id="j1")
    assert target.busy_calls == [(True, "j1")]

    ir.set_busy_compat(target=object(), busy=False, job_id=None)


def test_start_async_operation_compat_delegates_when_target_has_native_starter() -> None:
    called: dict[str, object] = {}

    class _Native:
        def _start_async_operation(self, **kwargs) -> None:
            called.update(kwargs)

    target = _Native()
    token = CancellationToken()
    ir.start_async_operation_compat(
        target=target,
        token=token,
        job_id="j1",
        work=lambda: 1,
        on_success=lambda _r: None,
        on_failure=lambda _e: None,
    )

    assert called["token"] is token
    assert called["job_id"] == "j1"


def test_start_async_operation_compat_fallback_success_finalizes_and_refreshes() -> None:
    target = _CompatTarget()
    callbacks: dict[str, object] = {}
    finished: list[int] = []

    def _run_async(work, *, on_finished, on_failed):
        callbacks["work"] = work
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    token = CancellationToken()
    ir.start_async_operation_compat(
        target=target,
        token=token,
        job_id="job-1",
        work=lambda: 7,
        on_success=lambda result: finished.append(cast(int, result)),
        on_failure=lambda _exc: None,
        run_async=_run_async,
    )

    assert target._cancel_token is token
    assert target.busy_calls[0] == (True, "job-1")
    assert len(target._active_jobs) == 1

    cast(Callable[[object], None], callbacks["on_finished"])(cast(Callable[[], int], callbacks["work"])())

    assert finished == [7]
    assert target._cancel_token is None
    assert target._active_jobs == []
    assert target.busy_calls[-1] == (False, None)
    assert target.refresh_calls == 1


def test_start_async_operation_compat_fallback_failure_handles_non_list_active_jobs() -> None:
    class _Target:
        def __init__(self) -> None:
            self._cancel_token = None
            self._active_jobs = ()
            self.calls: list[tuple[bool, str | None]] = []
            self._refresh_ui = 123

        def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
            self.calls.append((busy, job_id))

    target = _Target()
    failures: list[Exception] = []
    callbacks: dict[str, object] = {}

    def _run_async(_work, *, on_finished, on_failed):
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    ir.start_async_operation_compat(
        target=target,
        token=CancellationToken(),
        job_id="job-2",
        work=lambda: 1,
        on_success=lambda _r: None,
        on_failure=lambda exc: failures.append(exc),
        run_async=_run_async,
    )

    err = RuntimeError("x")
    cast(Callable[[Exception], None], callbacks["on_failed"])(err)

    assert failures == [err]
    assert target._cancel_token is None
    assert target.calls[0] == (True, "job-2")
    assert target.calls[-1] == (False, None)


class _RunnerTarget:
    def __init__(self, *, closing: bool) -> None:
        self._cancel_token = None
        self._active_jobs: list[object] = []
        self._is_closing = closing
        self.busy_calls: list[tuple[bool, str | None]] = []
        self.refresh_calls = 0

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.busy_calls.append((busy, job_id))

    def _refresh_ui(self) -> None:
        self.refresh_calls += 1


def test_async_operation_runner_success_finalizes_even_if_success_handler_raises() -> None:
    target = _RunnerTarget(closing=False)
    callbacks: dict[str, object] = {}

    def _run_async(work, *, on_finished, on_failed):
        callbacks["work"] = work
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    runner = ir.AsyncOperationRunner(target, run_async=_run_async)
    runner.start(
        token=CancellationToken(),
        job_id="j3",
        work=lambda: 5,
        on_success=lambda _r: (_ for _ in ()).throw(RuntimeError("boom")),
        on_failure=lambda _e: None,
    )

    try:
        cast(Callable[[object], None], callbacks["on_finished"])(
            cast(Callable[[], int], callbacks["work"])()
        )
    except RuntimeError as exc:
        assert str(exc) == "boom"

    assert target._cancel_token is None
    assert target._active_jobs == []
    assert target.busy_calls[0] == (True, "j3")
    assert target.busy_calls[-1] == (False, None)
    assert target.refresh_calls == 1


def test_async_operation_runner_failure_does_not_refresh_while_closing() -> None:
    target = _RunnerTarget(closing=True)
    callbacks: dict[str, object] = {}
    failures: list[Exception] = []

    def _run_async(_work, *, on_finished, on_failed):
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    runner = ir.AsyncOperationRunner(target, run_async=_run_async)
    runner.start(
        token=CancellationToken(),
        job_id="j4",
        work=lambda: 1,
        on_success=lambda _r: None,
        on_failure=lambda exc: failures.append(exc),
    )

    err = RuntimeError("fail")
    cast(Callable[[Exception], None], callbacks["on_failed"])(err)

    assert failures == [err]
    assert target._cancel_token is None
    assert target._active_jobs == []
    assert target.refresh_calls == 0
