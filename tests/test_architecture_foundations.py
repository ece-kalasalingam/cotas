from __future__ import annotations

import pytest

from common.contracts import validate_blueprint_registry_contracts
from common.exceptions import JobCancelledError, ValidationError
from common.jobs import CancellationToken, JobContext


def test_validation_error_supports_code_and_context() -> None:
    """Test validation error supports code and context.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    err = ValidationError("bad input", code="BAD_INPUT", context={"field": "course_code"})
    if not (str(err) == "bad input"):
        raise AssertionError('assertion failed')
    if not (err.code == "BAD_INPUT"):
        raise AssertionError('assertion failed')
    if not (err.context == {"field": "course_code"}):
        raise AssertionError('assertion failed')


def test_job_context_snapshots_language_and_operation() -> None:
    """Test job context snapshots language and operation.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    ctx = JobContext.create(step_id="instructor.generate_report", payload={"path": "sample.xlsx"})
    if not (ctx.job_id):
        raise AssertionError('assertion failed')
    if not (ctx.step_id == "instructor.generate_report"):
        raise AssertionError('assertion failed')
    if not (ctx.payload["path"] == "sample.xlsx"):
        raise AssertionError('assertion failed')
    if not (isinstance(ctx.language, str)):
        raise AssertionError('assertion failed')


def test_cancellation_token_raises_when_cancelled() -> None:
    """Test cancellation token raises when cancelled.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    token = CancellationToken()
    token.cancel()
    with pytest.raises(JobCancelledError):
        token.raise_if_cancelled()


def test_blueprint_registry_contracts_are_valid() -> None:
    """Test blueprint registry contracts are valid.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    validate_blueprint_registry_contracts()
