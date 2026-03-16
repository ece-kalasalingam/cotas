from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from common import qt_jobs


def test_function_job_run_emits_finished_on_success() -> None:
    job = qt_jobs.FunctionJob(lambda a, b: a + b, 2, 3)
    got = {"result": None, "error": None}
    job.signals.finished.connect(lambda value: got.__setitem__("result", value))
    job.signals.failed.connect(lambda exc: got.__setitem__("error", exc))

    job.run()

    assert got["result"] == 5
    assert got["error"] is None


def test_function_job_run_emits_failed_on_exception() -> None:
    def _boom():
        raise RuntimeError("boom")

    job = qt_jobs.FunctionJob(_boom)
    got = {"result": None, "error": None}
    job.signals.finished.connect(lambda value: got.__setitem__("result", value))
    job.signals.failed.connect(lambda exc: got.__setitem__("error", exc))

    job.run()

    assert got["result"] is None
    assert isinstance(got["error"], RuntimeError)


def test_run_in_background_wires_callbacks_and_starts_job() -> None:
    class _Pool:
        def __init__(self) -> None:
            self.started = []

        def start(self, job):
            self.started.append(job)

    pool = _Pool()
    events: list[tuple[str, object]] = []

    job = qt_jobs.run_in_background(
        lambda: "ok",
        on_finished=lambda value: events.append(("finished", value)),
        on_failed=lambda exc: events.append(("failed", exc)),
        thread_pool=pool,
    )

    assert pool.started == [job]

    # Simulate thread execution completion
    job.run()
    assert events == [("finished", "ok")]
