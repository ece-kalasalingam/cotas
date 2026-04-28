from __future__ import annotations

from typing import Callable, cast

from common.async_operation_runner import AsyncOperationRunner
from common.jobs import CancellationToken


class _RunnerTarget:
    def __init__(self, *, closing: bool) -> None:
        """Init.
        
        Args:
            closing: Parameter value (bool).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._cancel_token = None
        self._active_jobs: list[object] = []
        self._is_closing = closing
        self.busy_calls: list[tuple[bool, str | None]] = []
        self.refresh_calls = 0

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        """Set busy.
        
        Args:
            busy: Parameter value (bool).
            job_id: Parameter value (str | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.busy_calls.append((busy, job_id))

    def _refresh_ui(self) -> None:
        """Refresh ui.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.refresh_calls += 1


def test_async_operation_runner_success_finalizes_even_if_success_handler_raises() -> None:
    """Test async operation runner success finalizes even if success handler raises.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    target = _RunnerTarget(closing=False)
    callbacks: dict[str, object] = {}

    def _run_async(work, *, on_finished, on_failed):
        """Run async.
        
        Args:
            work: Parameter value.
            on_finished: Parameter value.
            on_failed: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        callbacks["work"] = work
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    runner = AsyncOperationRunner(
        target,
        run_async=_run_async,
        refresh_ui=lambda: target._refresh_ui(),
        should_refresh_ui=lambda: not target._is_closing,
    )
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
        if not (str(exc) == "boom"):
            raise AssertionError('assertion failed')

    if target._cancel_token is not None:
        raise AssertionError('assertion failed')
    if not (target._active_jobs == []):
        raise AssertionError('assertion failed')
    if not (target.busy_calls[0] == (True, "j3")):
        raise AssertionError('assertion failed')
    if not (target.busy_calls[-1] == (False, None)):
        raise AssertionError('assertion failed')
    if not (target.refresh_calls == 1):
        raise AssertionError('assertion failed')


def test_async_operation_runner_failure_does_not_refresh_while_closing() -> None:
    """Test async operation runner failure does not refresh while closing.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    target = _RunnerTarget(closing=True)
    callbacks: dict[str, object] = {}
    failures: list[Exception] = []

    def _run_async(_work, *, on_finished, on_failed):
        """Run async.
        
        Args:
            _work: Parameter value.
            on_finished: Parameter value.
            on_failed: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        callbacks["on_finished"] = on_finished
        callbacks["on_failed"] = on_failed
        return object()

    runner = AsyncOperationRunner(
        target,
        run_async=_run_async,
        refresh_ui=lambda: target._refresh_ui(),
        should_refresh_ui=lambda: not target._is_closing,
    )
    runner.start(
        token=CancellationToken(),
        job_id="j4",
        work=lambda: 1,
        on_success=lambda _r: None,
        on_failure=lambda exc: failures.append(exc),
    )

    err = RuntimeError("fail")
    cast(Callable[[Exception], None], callbacks["on_failed"])(err)

    if not (failures == [err]):
        raise AssertionError('assertion failed')
    if target._cancel_token is not None:
        raise AssertionError('assertion failed')
    if not (target._active_jobs == []):
        raise AssertionError('assertion failed')
    if not (target.refresh_calls == 0):
        raise AssertionError('assertion failed')
