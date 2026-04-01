from __future__ import annotations

from dataclasses import dataclass
from typing import Any, cast

import pytest

from common.exceptions import JobCancelledError, ValidationError
from common.jobs import CancellationToken
from domain.template_versions.course_setup_v2_impl.validation_batch_runner import (
    BatchValidationAccumulator,
    BatchValidationRunner,
    ValidationRejectionDecision,
)


def _issue_builder(code: str, context: dict[str, object], fallback_message: str) -> dict[str, object]:
    """Issue builder.
    
    Args:
        code: Parameter value (str).
        context: Parameter value (dict[str, object]).
        fallback_message: Parameter value (str).
    
    Returns:
        dict[str, object]: Return value.
    
    Raises:
        None.
    """
    return {
        "code": code,
        "category": "validation",
        "severity": "error",
        "translation_key": "validation.test",
        "message": fallback_message,
        "context": dict(context),
    }


@dataclass(frozen=True)
class _ValidationPayload:
    template_id: str


def _runner(*, duplicate_code: str = "DUP", unexpected_code: str = "UNEXPECTED") -> BatchValidationRunner[_ValidationPayload]:
    """Runner.
    
    Args:
        duplicate_code: Parameter value (str).
        unexpected_code: Parameter value (str).
    
    Returns:
        BatchValidationRunner[_ValidationPayload]: Return value.
    
    Raises:
        None.
    """
    return BatchValidationRunner[_ValidationPayload](
        issue_builder=_issue_builder,
        duplicate_path_issue_code=duplicate_code,
        unexpected_issue_code=unexpected_code,
    )


def _on_validated(acc: BatchValidationAccumulator, path: str, payload: _ValidationPayload) -> None:
    """On validated.
    
    Args:
        acc: Parameter value (BatchValidationAccumulator).
        path: Parameter value (str).
        payload: Parameter value (_ValidationPayload).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    acc.add_valid(path=path, template_id=payload.template_id)


def test_runner_handles_duplicate_paths_and_shape() -> None:
    """Test runner handles duplicate paths and shape.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    result = cast(
        dict[str, Any],
        _runner().run(
        workbook_paths=["a.xlsx", "a.xlsx", "b.xlsx"],
        validate_path=lambda path: _ValidationPayload(template_id=f"T::{path}"),
        on_validated=_on_validated,
        ),
    )

    assert sorted(result.keys()) == [
        "duplicate_paths",
        "duplicate_sections",
        "invalid_paths",
        "mismatched_paths",
        "rejections",
        "template_ids",
        "valid_paths",
    ]
    assert result["valid_paths"] == ["a.xlsx", "b.xlsx"]
    assert result["duplicate_paths"] == ["a.xlsx"]
    duplicate_rejection = cast(list[dict[str, Any]], result["rejections"])[0]
    assert duplicate_rejection["reason_kind"] == "duplicate_path"
    assert duplicate_rejection["issue"]["code"] == "DUP"


def test_runner_propagates_cancellation() -> None:
    """Test runner propagates cancellation.
    
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
        _runner().run(
            workbook_paths=["x.xlsx"],
            validate_path=lambda _path: _ValidationPayload(template_id="T"),
            on_validated=_on_validated,
            cancel_token=token,
        )


def test_runner_maps_validation_and_unexpected_rejections() -> None:
    """Test runner maps validation and unexpected rejections.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _validate(path: str) -> _ValidationPayload:
        """Validate.
        
        Args:
            path: Parameter value (str).
        
        Returns:
            _ValidationPayload: Return value.
        
        Raises:
            None.
        """
        if path == "bad.xlsx":
            raise ValidationError("Bad workbook", code="BAD_CODE", context={"workbook": path})
        if path == "boom.xlsx":
            raise RuntimeError("boom")
        return _ValidationPayload(template_id="COURSE_SETUP_V2")

    result = cast(
        dict[str, Any],
        _runner(unexpected_code="UNEXP").run(
        workbook_paths=["ok.xlsx", "bad.xlsx", "boom.xlsx"],
        validate_path=_validate,
        on_validated=_on_validated,
        classify_validation_error=lambda _path, _exc, issue: ValidationRejectionDecision(
            reason_kind="template_mismatch",
            mark_invalid=True,
            mark_mismatched=issue.get("code") == "BAD_CODE",
        ),
        ),
    )

    assert result["valid_paths"] == ["ok.xlsx"]
    assert result["invalid_paths"] == ["bad.xlsx", "boom.xlsx"]
    assert result["mismatched_paths"] == ["bad.xlsx"]
    reasons = [item["reason_kind"] for item in cast(list[dict[str, Any]], result["rejections"])]
    assert reasons == ["template_mismatch", "invalid"]
    rejections = cast(list[dict[str, Any]], result["rejections"])
    assert rejections[0]["issue"]["code"] == "BAD_CODE"
    assert rejections[1]["issue"]["code"] == "UNEXP"
