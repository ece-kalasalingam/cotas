from __future__ import annotations

from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from common import qt_jobs


def test_function_job_run_emits_finished_on_success() -> None:
    """Test function job run emits finished on success.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    job = qt_jobs.FunctionJob(lambda a, b: a + b, 2, 3)
    got = {"result": None, "error": None}
    job.signals.finished.connect(lambda value: got.__setitem__("result", value))
    job.signals.failed.connect(lambda exc: got.__setitem__("error", exc))

    job.run()

    if not (got["result"] == 5):
        raise AssertionError('assertion failed')
    if got["error"] is not None:
        raise AssertionError('assertion failed')


def test_function_job_run_emits_failed_on_exception() -> None:
    """Test function job run emits failed on exception.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _boom():
        """Boom.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        raise RuntimeError("boom")

    job = qt_jobs.FunctionJob(_boom)
    got = {"result": None, "error": None}
    job.signals.finished.connect(lambda value: got.__setitem__("result", value))
    job.signals.failed.connect(lambda exc: got.__setitem__("error", exc))

    job.run()

    if got["result"] is not None:
        raise AssertionError('assertion failed')
    if not (isinstance(got["error"], RuntimeError)):
        raise AssertionError('assertion failed')


def test_run_in_background_wires_callbacks_and_starts_job() -> None:
    """Test run in background wires callbacks and starts job.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    class _Pool:
        def __init__(self) -> None:
            """Init.
            
            Args:
                None.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self.started = []

        def start(self, job):
            """Start.
            
            Args:
                job: Parameter value.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self.started.append(job)

    pool = _Pool()
    events: list[tuple[str, object]] = []

    job = qt_jobs.run_in_background(
        lambda: "ok",
        on_finished=lambda value: events.append(("finished", value)),
        on_failed=lambda exc: events.append(("failed", exc)),
        thread_pool=cast(Any, pool),
    )

    if not (pool.started == [job]):
        raise AssertionError('assertion failed')

    # Simulate thread execution completion
    job.run()
    if not (events == [("finished", "ok")]):
        raise AssertionError('assertion failed')

