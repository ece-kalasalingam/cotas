"""Validation logic for uploaded CO Analysis source workbooks."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from common.error_catalog import validation_error_from_key
from common.exceptions import ValidationError
from common.utils import canonical_path_key
from domain.template_strategy_router import (
    get_template_strategy,
    validate_workbooks,
    read_valid_system_workbook_payload,
)

_last_validated_template_id: str = ""


def validate_uploaded_source_workbook(workbook_path: str | Path) -> None:
    global _last_validated_template_id
    try:
        import openpyxl
    except ModuleNotFoundError as exc:
        raise validation_error_from_key(
            "instructor.validation.openpyxl_missing",
            code="OPENPYXL_MISSING",
        ) from exc

    workbook_file = Path(workbook_path)
    if not workbook_file.exists():
        raise validation_error_from_key(
            "instructor.validation.workbook_not_found",
            code="WORKBOOK_NOT_FOUND",
            workbook=str(workbook_file),
        )

    try:
        workbook = openpyxl.load_workbook(workbook_file, data_only=False, read_only=True)
    except Exception as exc:
        raise validation_error_from_key(
            "instructor.validation.workbook_open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(workbook_file),
        ) from exc
    try:
        template_id, _ = _read_system_manifest_payload(workbook)
        _last_validated_template_id = template_id
    finally:
        workbook.close()

    try:
        workbook = openpyxl.load_workbook(workbook_file, data_only=False)
    except Exception as exc:
        raise validation_error_from_key(
            "instructor.validation.workbook_open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(workbook_file),
        ) from exc
    try:
        result = validate_workbooks(
            template_id=template_id,
            workbook_paths=[str(workbook_file)],
            workbook_kind="marks_template",
        )
        _raise_first_validation_issue(result=result, workbook_path=str(workbook_file))
    finally:
        workbook.close()


def consume_last_source_anomaly_warnings() -> list[str]:
    if not _last_validated_template_id:
        return []
    strategy = get_template_strategy(_last_validated_template_id)
    fn = getattr(strategy, "consume_last_marks_anomaly_warnings", None)
    if not callable(fn):
        raise validation_error_from_key(
            "validation.template.validator_missing",
            code="COA_TEMPLATE_VALIDATOR_MISSING",
            template_id=_last_validated_template_id,
            operation="consume_last_marks_anomaly_warnings",
        )
    return list(fn())


def _read_system_manifest_payload(workbook: Any) -> tuple[str, dict[str, Any]]:
    payload = read_valid_system_workbook_payload(workbook)
    return payload.template_id, payload.manifest


def _raise_first_validation_issue(*, result: dict[str, object], workbook_path: str) -> None:
    workbook_key = canonical_path_key(workbook_path)
    valid_paths_raw = result.get("valid_paths", [])
    valid_paths = [str(path) for path in valid_paths_raw] if isinstance(valid_paths_raw, list) else []
    valid_keys = {canonical_path_key(path) for path in valid_paths}
    if workbook_key in valid_keys:
        return

    rejections_raw = result.get("rejections", [])
    rejection_items = [item for item in rejections_raw if isinstance(item, dict)] if isinstance(rejections_raw, list) else []
    issue = next(
        (
            dict(item.get("issue", {}))
            for item in rejection_items
            if canonical_path_key(str(item.get("path", "")).strip()) == workbook_key and isinstance(item.get("issue"), dict)
        ),
        None,
    )
    if issue is not None:
        code = str(issue.get("code", "VALIDATION_ERROR")).strip() or "VALIDATION_ERROR"
        message = str(issue.get("message", code)).strip() or code
        context = issue.get("context", {})
        context_dict = dict(context) if isinstance(context, dict) else {}
        raise ValidationError(message, code=code, context=context_dict)
    raise validation_error_from_key(
        "instructor.validation.workbook_open_failed",
        code="WORKBOOK_OPEN_FAILED",
        workbook=workbook_path,
    )
