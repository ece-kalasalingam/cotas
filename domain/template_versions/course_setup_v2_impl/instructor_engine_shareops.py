"""V2-local workbook sheet helpers for course template generation."""

from __future__ import annotations

from typing import Any, Sequence

from common.excel_sheet_layout import (
    build_xlsxwriter_body_format,
    build_xlsxwriter_header_format,
    protect_xlsxwriter_sheet,
)
from common.error_catalog import validation_error_from_key
from common.registry import (
    SYSTEM_HASH_HEADER_TEMPLATE_HASH,
    SYSTEM_HASH_HEADER_TEMPLATE_ID,
    SYSTEM_HASH_SHEET_NAME,
)
from common.sheet_schema import ValidationRule
from common.utils import (
    copy_system_hash_sheet as _copy_system_hash_sheet_common,
    ensure_uniform_template_id_and_copy_system_hash as _ensure_uniform_template_id_and_copy_system_hash_common,
)
from common.workbook_signing import sign_payload
from domain.template_strategy_router import read_valid_template_id_from_system_hash_sheet


def build_header_format(workbook: Any, header_style: dict[str, Any]) -> Any:
    return build_xlsxwriter_header_format(workbook, header_style)


def build_body_format(workbook: Any, body_style: dict[str, Any]) -> Any:
    return build_xlsxwriter_body_format(workbook, body_style)


def generate_worksheet(
    *,
    workbook: Any,
    sheet_name: str,
    headers: Sequence[str],
    data: Sequence[Sequence[Any]],
    header_format: Any,
    body_format: Any,
) -> Any:
    if not sheet_name or not isinstance(sheet_name, str):
        raise validation_error_from_key("instructor.validation.invalid_sheet_name")
    if not headers:
        raise validation_error_from_key("instructor.validation.headers_empty", sheet_name=sheet_name)
    if len(set(headers)) != len(headers):
        raise validation_error_from_key("instructor.validation.headers_unique", sheet_name=sheet_name)

    column_count = len(headers)
    for row_index, row in enumerate(data, start=1):
        if len(row) != column_count:
            raise validation_error_from_key(
                "instructor.validation.row_length_mismatch",
                row=row_index,
                sheet_name=sheet_name,
                expected=column_count,
                found=len(row),
            )

    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.write_row(0, 0, headers, header_format)
    for col_index, value in enumerate(headers):
        worksheet.set_column(col_index, col_index, max(12, len(str(value)) + 2))
    for row_offset, row_values in enumerate(data, start=1):
        worksheet.write_row(row_offset, 0, row_values, body_format)
    worksheet.freeze_panes(1, 0)
    return worksheet


def apply_validation(worksheet: Any, rule: ValidationRule) -> None:
    options = dict(rule.options)
    validation_type = options.pop("validate", None)
    if not validation_type:
        return
    options["validate"] = validation_type
    options.setdefault("ignore_blank", True)
    worksheet.data_validation(
        rule.first_row,
        rule.first_col,
        rule.last_row,
        rule.last_col,
        options,
    )


def protect_sheet(worksheet: Any) -> None:
    protect_xlsxwriter_sheet(worksheet)


def add_system_hash_sheet(workbook: Any, template_id: str) -> None:
    worksheet = workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
    worksheet.write_row(
        0,
        0,
        [SYSTEM_HASH_HEADER_TEMPLATE_ID, SYSTEM_HASH_HEADER_TEMPLATE_HASH],
    )
    worksheet.write_row(1, 0, [template_id, sign_payload(template_id)])
    worksheet.hide()


def copy_system_hash_sheet(source_workbook: Any, target_workbook: Any) -> None:
    _copy_system_hash_sheet_common(source_workbook, target_workbook)


def ensure_uniform_template_id_and_copy_system_hash(
    source_workbooks: list[Any],
    target_workbook: Any,
    *,
    routed_template_id: str | None = None,
    cancel_token: Any | None = None,
) -> str:
    return _ensure_uniform_template_id_and_copy_system_hash_common(
        source_workbooks,
        target_workbook,
        read_template_id=read_valid_template_id_from_system_hash_sheet,
        routed_template_id=routed_template_id,
        cancel_token=cancel_token,
    )


__all__ = [
    "add_system_hash_sheet",
    "copy_system_hash_sheet",
    "ensure_uniform_template_id_and_copy_system_hash",
    "apply_validation",
    "build_body_format",
    "build_header_format",
    "generate_worksheet",
    "protect_sheet",
]
