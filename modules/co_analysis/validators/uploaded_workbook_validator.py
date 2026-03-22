"""Validation logic for uploaded CO Analysis source workbooks."""

from __future__ import annotations

from pathlib import Path

from modules.instructor.validators.step2_filled_marks_validator import (
    validate_uploaded_filled_marks_workbook,
)


def validate_uploaded_source_workbook(workbook_path: str | Path) -> None:
    # CO Analysis uses the same signed template-version rules as Instructor Step 2.
    # Those rules are implemented via domain.template_versions.*
    validate_uploaded_filled_marks_workbook(workbook_path)
