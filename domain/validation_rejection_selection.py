"""Shared workbook-structure classification and rejection-selection helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from common.registry import (
    COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    get_sheet_name_by_key,
)
from common.runtime_dependency_guard import import_runtime_dependency
from common.utils import assert_not_symlink_path, canonical_path_key, normalize
from common.workbook_integrity.constants import SYSTEM_LAYOUT_SHEET


def classify_workbook_structure_for_validation(
    *,
    template_id: str,
    workbook_path: str | Path,
) -> str:
    """Return one of: marks_template, co_description, ambiguous, unknown."""
    openpyxl = import_runtime_dependency("openpyxl")
    source = Path(workbook_path)
    assert_not_symlink_path(source, context_key="workbook")
    try:
        workbook = openpyxl.load_workbook(source, data_only=False, read_only=True)
    except Exception:
        return "unknown"
    try:
        sheetnames = {normalize(name) for name in getattr(workbook, "sheetnames", [])}
    finally:
        workbook.close()

    metadata_sheet = normalize(get_sheet_name_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA))
    co_description_sheet = normalize(get_sheet_name_by_key(template_id, COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION))
    has_layout_sheet = normalize(SYSTEM_LAYOUT_SHEET) in sheetnames
    has_metadata_sheet = metadata_sheet in sheetnames
    has_co_description_sheet = co_description_sheet in sheetnames

    marks_candidate = has_layout_sheet
    co_description_candidate = has_metadata_sheet and has_co_description_sheet

    if marks_candidate and co_description_candidate:
        return "ambiguous"
    if marks_candidate:
        return "marks_template"
    if co_description_candidate:
        return "co_description"
    return "unknown"


def select_preferred_validation_rejection(
    *,
    template_id: str,
    workbook_path: str | Path,
    primary_kind: str,
    secondary_kind: str,
    primary_result: Mapping[str, Any] | None,
    secondary_result: Mapping[str, Any] | None,
) -> dict[str, object] | None:
    """Pick the best rejection payload for a workbook across two validation results."""
    key = canonical_path_key(str(workbook_path))

    def _rejection_for_key(result: Mapping[str, Any] | None) -> dict[str, object] | None:
        if not isinstance(result, Mapping):
            return None
        raw = result.get("rejections", [])
        if not isinstance(raw, list):
            return None
        for item in raw:
            if not isinstance(item, dict):
                continue
            item_path = str(item.get("path", "")).strip()
            if item_path and canonical_path_key(item_path) == key:
                return item
        return None

    primary_rejection = _rejection_for_key(primary_result)
    secondary_rejection = _rejection_for_key(secondary_result)
    if primary_rejection is None:
        return secondary_rejection
    if secondary_rejection is None:
        return primary_rejection

    classified_kind = classify_workbook_structure_for_validation(
        template_id=template_id,
        workbook_path=workbook_path,
    )
    if normalize(classified_kind) == normalize(primary_kind):
        return primary_rejection
    if normalize(classified_kind) == normalize(secondary_kind):
        return secondary_rejection

    primary_issue = primary_rejection.get("issue", {})
    primary_code = str(primary_issue.get("code", "")).strip() if isinstance(primary_issue, dict) else ""
    if primary_code == "COA_LAYOUT_SHEET_MISSING":
        return secondary_rejection
    return primary_rejection

