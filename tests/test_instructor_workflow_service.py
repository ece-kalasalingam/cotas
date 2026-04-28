from __future__ import annotations

import logging
from pathlib import Path

import pytest

from common.exceptions import JobCancelledError
from common.jobs import CancellationToken
from services import instructor_workflow_service as service_mod


def test_service_logs_cancellation(
    caplog: pytest.LogCaptureFixture,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Test service logs cancellation.

    Args:
        caplog: Parameter value (pytest.LogCaptureFixture).
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).

    Returns:
        None.

    Raises:
        None.
    """
    service = service_mod.InstructorWorkflowService()
    token = CancellationToken()
    token.cancel()
    context = service.create_job_context(step_id="prepare_marks")
    source = tmp_path / "course_details.xlsx"
    output = tmp_path / "marks_template.xlsx"
    called = {"count": 0}

    monkeypatch.setattr(
        service_mod.InstructorWorkflowService,
        "_resolve_template_id_from_workbook",
        lambda _self, _path: "COURSE_SETUP_V2",
    )

    def _fake_generate_workbooks(**_kwargs):
        """Fake generate workbooks.

        Args:
            _kwargs: Parameter value.

        Returns:
            None.

        Raises:
            None.
        """
        called["count"] += 1
        return {"generated_workbook_paths": []}

    monkeypatch.setattr(service_mod, "generate_workbooks", _fake_generate_workbooks)

    caplog.set_level(logging.INFO, logger=service_mod.__name__)
    with pytest.raises(JobCancelledError):
        service.generate_marks_template(source, output, context=context, cancel_token=token)

    if not (any(record.getMessage() == "Instructor workflow step cancelled." for record in caplog.records)):
        raise AssertionError('assertion failed')
    if not (called["count"] == 0):
        raise AssertionError('assertion failed')


def test_resolve_timeout_seconds_invalid_env_uses_default(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test resolve timeout seconds invalid env uses default.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).

    Returns:
        None.

    Raises:
        None.
    """
    monkeypatch.setenv(service_mod.WORKFLOW_STEP_TIMEOUT_ENV_VAR, "not-an-int")
    if not ((
        service_mod.InstructorWorkflowService._resolve_timeout_seconds()
        == service_mod.DEFAULT_WORKFLOW_STEP_TIMEOUT_SECONDS
    )):
        raise AssertionError('assertion failed')
