"""Validation logic for uploaded CO Analysis source workbooks."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from common.error_catalog import validation_error_from_key
from domain.template_strategy_router import (
    get_template_strategy,
    read_valid_system_workbook_payload,
)

_last_validated_template_id: str = ""


def validate_source_manifest_schema_by_template(
    workbook: object,
    manifest: object,
    *,
    template_id: str,
) -> None:
    strategy = get_template_strategy(template_id)
    fn = getattr(strategy, "validate_filled_marks_manifest_schema", None)
    if not callable(fn):
        raise validation_error_from_key(
            "validation.template.validator_missing",
            code="COA_TEMPLATE_VALIDATOR_MISSING",
            template_id=template_id,
            operation="validate_filled_marks_manifest_schema",
        )
    fn(workbook, manifest)


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
        template_id, manifest = _read_system_manifest_payload(workbook)
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
        validate_source_manifest_schema_by_template(workbook, manifest, template_id=template_id)
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
