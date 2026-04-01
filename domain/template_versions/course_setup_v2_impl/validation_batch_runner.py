"""Shared batch-validation orchestration for COURSE_SETUP_V2 validators.

This module centralizes reusable orchestration only:
- path normalization + de-duplication
- cancellation-aware per-file loop
- validation/unexpected exception mapping
- standardized result payload assembly
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Generic, Protocol, TypeVar

from common.exceptions import JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.utils import canonical_path_key, dedupe_paths_by_canonical_key

T = TypeVar("T")


class IssueBuilder(Protocol):
    def __call__(
        self,
        *,
        code: str,
        context: dict[str, Any],
        fallback_message: str,
    ) -> dict[str, object]:
        """Call.
        
        Args:
            code: Parameter value (str).
            context: Parameter value (dict[str, Any]).
            fallback_message: Parameter value (str).
        
        Returns:
            dict[str, object]: Return value.
        
        Raises:
            None.
        """
        ...


@dataclass(frozen=True, slots=True)
class ValidationRejectionDecision:
    reason_kind: str = "invalid"
    mark_invalid: bool = True
    mark_mismatched: bool = False
    mark_duplicate_section: bool = False


class BatchValidationAccumulator:
    def __init__(self) -> None:
        """Init.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.valid_paths: list[str] = []
        self.invalid_paths: list[str] = []
        self.mismatched_paths: list[str] = []
        self.duplicate_paths: list[str] = []
        self.duplicate_sections: list[str] = []
        self.template_ids: dict[str, str] = {}
        self.rejections: list[dict[str, object]] = []

    def add_valid(self, *, path: str, template_id: str) -> None:
        """Add valid.
        
        Args:
            path: Parameter value (str).
            template_id: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.valid_paths.append(path)
        self.template_ids[canonical_path_key(path)] = template_id

    def add_duplicate_path_rejection(self, *, path: str, issue: dict[str, object]) -> None:
        """Add duplicate path rejection.
        
        Args:
            path: Parameter value (str).
            issue: Parameter value (dict[str, object]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.duplicate_paths.append(path)
        self.rejections.append(
            {
                "path": path,
                "reason_kind": "duplicate_path",
                "issue": issue,
            }
        )

    def add_rejection(
        self,
        *,
        path: str,
        issue: dict[str, object],
        decision: ValidationRejectionDecision,
    ) -> None:
        """Add rejection.
        
        Args:
            path: Parameter value (str).
            issue: Parameter value (dict[str, object]).
            decision: Parameter value (ValidationRejectionDecision).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if decision.mark_invalid:
            self.invalid_paths.append(path)
        if decision.mark_mismatched:
            self.mismatched_paths.append(path)
        if decision.mark_duplicate_section:
            self.duplicate_sections.append(path)
        self.rejections.append(
            {
                "path": path,
                "reason_kind": decision.reason_kind,
                "issue": issue,
            }
        )

    def to_payload(self) -> dict[str, object]:
        """To payload.
        
        Args:
            None.
        
        Returns:
            dict[str, object]: Return value.
        
        Raises:
            None.
        """
        return {
            "valid_paths": list(self.valid_paths),
            "invalid_paths": list(self.invalid_paths),
            "mismatched_paths": list(self.mismatched_paths),
            "duplicate_paths": list(self.duplicate_paths),
            "duplicate_sections": list(self.duplicate_sections),
            "template_ids": dict(self.template_ids),
            "rejections": list(self.rejections),
        }


def _dedupe_paths(workbook_paths: Sequence[str | Path]) -> tuple[list[str], list[str]]:
    """Dedupe paths.
    
    Args:
        workbook_paths: Parameter value (Sequence[str | Path]).
    
    Returns:
        tuple[list[str], list[str]]: Return value.
    
    Raises:
        None.
    """
    return dedupe_paths_by_canonical_key(workbook_paths, skip_empty=True)


class BatchValidationRunner(Generic[T]):
    def __init__(
        self,
        *,
        issue_builder: IssueBuilder,
        duplicate_path_issue_code: str,
        unexpected_issue_code: str,
    ) -> None:
        """Init.
        
        Args:
            issue_builder: Parameter value (IssueBuilder).
            duplicate_path_issue_code: Parameter value (str).
            unexpected_issue_code: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._issue_builder = issue_builder
        self._duplicate_path_issue_code = duplicate_path_issue_code
        self._unexpected_issue_code = unexpected_issue_code

    def run(
        self,
        *,
        workbook_paths: Sequence[str | Path],
        validate_path: Callable[[str], T],
        on_validated: Callable[[BatchValidationAccumulator, str, T], None],
        cancel_token: CancellationToken | None = None,
        classify_validation_error: Callable[
            [str, ValidationError, dict[str, object]], ValidationRejectionDecision
        ]
        | None = None,
    ) -> dict[str, object]:
        """Run.
        
        Args:
            workbook_paths: Parameter value (Sequence[str | Path]).
            validate_path: Parameter value (Callable[[str], T]).
            on_validated: Parameter value (Callable[[BatchValidationAccumulator, str, T], None]).
            cancel_token: Parameter value (CancellationToken | None).
            classify_validation_error: Parameter value (Callable[[str, ValidationError, dict[str, object]], ValidationRejectionDecision] | None).
        
        Returns:
            dict[str, object]: Return value.
        
        Raises:
            None.
        """
        unique_paths, duplicate_paths = _dedupe_paths(workbook_paths)
        result = BatchValidationAccumulator()

        for path in duplicate_paths:
            issue = self._issue_builder(
                code=self._duplicate_path_issue_code,
                context={"workbook": path},
                fallback_message="Duplicate file path skipped.",
            )
            result.add_duplicate_path_rejection(path=path, issue=issue)

        for path in unique_paths:
            if cancel_token is not None:
                cancel_token.raise_if_cancelled()
            try:
                validated = validate_path(path)
            except JobCancelledError:
                raise
            except ValidationError as exc:
                issue = self._issue_builder(
                    code=str(getattr(exc, "code", "VALIDATION_ERROR")),
                    context=dict(getattr(exc, "context", {}) or {}),
                    fallback_message=str(exc).strip() or "Validation failed.",
                )
                if classify_validation_error is None:
                    decision = ValidationRejectionDecision()
                else:
                    decision = classify_validation_error(path, exc, issue)
                result.add_rejection(path=path, issue=issue, decision=decision)
                continue
            except Exception as exc:
                issue = self._issue_builder(
                    code=self._unexpected_issue_code,
                    context={"workbook": path},
                    fallback_message=str(exc).strip() or "File skipped due to an unexpected validation failure.",
                )
                result.add_rejection(
                    path=path,
                    issue=issue,
                    decision=ValidationRejectionDecision(),
                )
                continue
            on_validated(result, path, validated)

        return result.to_payload()


__all__ = [
    "BatchValidationAccumulator",
    "BatchValidationRunner",
    "ValidationRejectionDecision",
]
