from __future__ import annotations

import pytest

from common.contracts import validate_blueprint_registry_contracts
from common.exceptions import JobCancelledError, ValidationError
from common.jobs import CancellationToken, JobContext


def test_validation_error_supports_code_and_context() -> None:
    err = ValidationError("bad input", code="BAD_INPUT", context={"field": "course_code"})
    assert str(err) == "bad input"
    assert err.code == "BAD_INPUT"
    assert err.context == {"field": "course_code"}


def test_job_context_snapshots_language_and_operation() -> None:
    ctx = JobContext.create(step_id="instructor.generate_report", payload={"path": "sample.xlsx"})
    assert ctx.job_id
    assert ctx.step_id == "instructor.generate_report"
    assert ctx.payload["path"] == "sample.xlsx"
    assert isinstance(ctx.language, str)


def test_cancellation_token_raises_when_cancelled() -> None:
    token = CancellationToken()
    token.cancel()
    with pytest.raises(JobCancelledError):
        token.raise_if_cancelled()


def test_blueprint_registry_contracts_are_valid() -> None:
    validate_blueprint_registry_contracts()
