"""Generator for the Course Details workbook template."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any, Sequence
from uuid import uuid4

from common.constants import (
    ALLOW_FILTER,
    ALLOW_SELECT_LOCKED,
    ALLOW_SELECT_UNLOCKED,
    ALLOW_SORT,
    HEADER_PATTERNFILL_COLOR,
    ID_COURSE_SETUP,
    WORKBOOK_PASSWORD,
)
from common.exceptions import AppSystemError, ValidationError
from common.registry import BLUEPRINT_REGISTRY
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.sheet_schema import ValidationRule, WorkbookBlueprint

_logger = logging.getLogger(__name__)


def generate_course_details_template(
    output_path: str | Path, template_id: str = ID_COURSE_SETUP
) -> Path:
    """Generate and save the course details template workbook with atomic replace."""
    try:
        import xlsxwriter
    except ModuleNotFoundError as exc:
        raise ValidationError(
            "xlsxwriter is not installed. Install it to generate course templates."
        ) from exc

    blueprint = _get_blueprint(template_id)
    sample_data = SAMPLE_SETUP_DATA
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    temp_output = output.with_name(f"{output.name}.{uuid4().hex}.tmp")
    workbook = xlsxwriter.Workbook(str(temp_output), {"constant_memory": True})
    workbook_closed = False

    def _cleanup_incomplete_output() -> None:
        nonlocal workbook_closed
        if not workbook_closed:
            try:
                workbook.close()
                workbook_closed = True
            except Exception:
                pass
        if temp_output.exists():
            try:
                temp_output.unlink()
            except OSError:
                _logger.warning("Failed to cleanup temp template file: %s", temp_output)

    try:
        header_format = _build_header_format(workbook, blueprint.style_registry.get("header", {}))
        body_format = _build_body_format(workbook, blueprint.style_registry.get("body", {}))

        for sheet_schema in blueprint.sheets:
            if len(sheet_schema.header_matrix) != 1:
                raise ValidationError(
                    f"Sheet '{sheet_schema.name}' must define exactly one header row."
                )

            worksheet = generate_worksheet(
                workbook=workbook,
                sheet_name=sheet_schema.name,
                headers=sheet_schema.header_matrix[0],
                data=sample_data.get(sheet_schema.name, []),
                header_format=header_format,
                body_format=body_format,
            )

            for validation in sheet_schema.validations:
                _apply_validation(worksheet, validation)

            if sheet_schema.is_protected:
                _protect_sheet(worksheet)

        workbook.close()
        workbook_closed = True
        os.replace(temp_output, output)
    except ValidationError:
        _cleanup_incomplete_output()
        raise
    except Exception as exc:
        _cleanup_incomplete_output()
        _logger.exception(
            "Failed to generate course details template. template_id=%s output=%s",
            template_id,
            output,
        )
        raise AppSystemError(
            f"Failed to generate course details template at '{output}'."
        ) from exc

    return output


def _get_blueprint(template_id: str) -> WorkbookBlueprint:
    blueprint = BLUEPRINT_REGISTRY.get(template_id)
    if blueprint is None:
        available = ", ".join(sorted(BLUEPRINT_REGISTRY))
        raise ValidationError(
            f"Unknown workbook template '{template_id}'. Available templates: {available}."
        )
    return blueprint


def _build_header_format(workbook: Any, header_style: dict[str, Any]) -> Any:
    bg_color = str(header_style.get("bg_color", HEADER_PATTERNFILL_COLOR))
    return workbook.add_format(
        {
            "bold": bool(header_style.get("bold", True)),
            "bg_color": bg_color,
            "border": int(header_style.get("border", 1)),
            "align": str(header_style.get("align", "center")),
            "valign": str(header_style.get("valign", "vcenter")),
        }
    )


def _build_body_format(workbook: Any, body_style: dict[str, Any]) -> Any:
    return workbook.add_format(
        {
            "locked": bool(body_style.get("locked", False)),
            "border": int(body_style.get("border", 1)),
        }
    )


def generate_worksheet(
    workbook: Any,
    sheet_name: str,
    headers: Sequence[str],
    data: Sequence[Sequence[Any]],
    header_format: Any,
    body_format: Any,
) -> Any:
    """Create a worksheet with strict validation and efficient row writes."""
    if not sheet_name or not isinstance(sheet_name, str):
        raise ValidationError("Invalid sheet name.")

    if not headers:
        raise ValidationError(f"Headers cannot be empty for sheet '{sheet_name}'.")

    if len(set(headers)) != len(headers):
        raise ValidationError(f"Headers must be unique for sheet '{sheet_name}'.")

    column_count = len(headers)
    for row_index, row in enumerate(data, start=1):
        if len(row) != column_count:
            raise ValidationError(
                f"Row {row_index} length mismatch in '{sheet_name}': "
                f"expected {column_count}, got {len(row)}."
            )

    worksheet = workbook.add_worksheet(sheet_name)
    col_widths: dict[int, int] = {}
    write_row = worksheet.write_row
    write_row(0, 0, headers, header_format)
    for col_index, value in enumerate(headers):
        col_widths[col_index] = max(12, len(str(value)) + 2)

    for col_index, width in col_widths.items():
        worksheet.set_column(col_index, col_index, width)

    for row_offset, row_values in enumerate(data, start=1):
        write_row(row_offset, 0, row_values, body_format)

    worksheet.freeze_panes(1, 0)
    return worksheet


def _apply_validation(worksheet: Any, rule: ValidationRule) -> None:
    options = dict(rule.options)
    validation_type = options.pop("validate", None)
    if not validation_type:
        return

    options["validate"] = validation_type
    if "ignore_blank" not in options:
        options["ignore_blank"] = True

    worksheet.data_validation(
        rule.first_row,
        rule.first_col,
        rule.last_row,
        rule.last_col,
        options,
    )


def _protect_sheet(worksheet: Any) -> None:
    worksheet.protect(
        WORKBOOK_PASSWORD,
        {
            "sort": ALLOW_SORT,
            "autofilter": ALLOW_FILTER,
            "select_locked_cells": ALLOW_SELECT_LOCKED,
            "select_unlocked_cells": ALLOW_SELECT_UNLOCKED,
        },
    )
