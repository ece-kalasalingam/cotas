"""COURSE_SETUP_V2 CO description template generation."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from uuid import uuid4

from common.constants import ID_COURSE_SETUP, WORKBOOK_TEMP_SUFFIX
from common.error_catalog import validation_error_from_key
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.excel_sheet_layout import (
    XLSX_AUTOFIT_MAX_WIDTH,
    XLSX_AUTOFIT_MIN_WIDTH,
    XLSX_AUTOFIT_PADDING,
    XLSX_AUTOFIT_SAMPLE_ROWS,
    apply_xlsxwriter_column_widths,
    build_template_xlsxwriter_formats,
    compute_sampled_column_widths,
)
from common.i18n import t
from common.jobs import CancellationToken
from common.registry import (
    COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    get_sheet_schema_by_key,
)
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.workbook_integrity import add_system_hash_sheet
from domain.template_versions.course_setup_v2_impl import instructor_engine_sheetops as _sheetops

_logger = logging.getLogger(__name__)
_WRAP_HEADER_NAMES = {"Description", "Summary_of_Topics/Expts./Project"}


def generate_co_description_template(
    output_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    try:
        import xlsxwriter
    except ModuleNotFoundError as exc:
        raise validation_error_from_key(
            "instructor.validation.xlsxwriter_missing",
            code="XLSXWRITER_MISSING",
        ) from exc

    metadata_schema = get_sheet_schema_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
    if metadata_schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SHEET_SCHEMA_MISSING",
            sheet_key=COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
            template_id=ID_COURSE_SETUP,
        )
    schema = get_sheet_schema_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION)
    if schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SHEET_SCHEMA_MISSING",
            sheet_key=COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
            template_id=ID_COURSE_SETUP,
        )
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_output = output.with_name(f"{output.name}.{uuid4().hex}{WORKBOOK_TEMP_SUFFIX}")
    workbook = xlsxwriter.Workbook(str(temp_output), {"constant_memory": True})
    workbook_closed = False

    def _cleanup_incomplete_output() -> None:
        nonlocal workbook_closed
        if not workbook_closed:
            try:
                workbook.close()
                workbook_closed = True
            except Exception:
                _logger.debug("Suppressing workbook close error during cleanup.", exc_info=True)
        if temp_output.exists():
            try:
                temp_output.unlink()
            except OSError:
                _logger.warning("Failed to cleanup temp template file: %s", temp_output)

    try:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        format_bundle = build_template_xlsxwriter_formats(
            workbook,
            template_id=ID_COURSE_SETUP,
            include_column_wrap=True,
        )
        _sheetops.write_schema_sheet(
            workbook=workbook,
            sheet_schema=metadata_schema,
            data=SAMPLE_SETUP_DATA.get(metadata_schema.name, []),
            header_format=format_bundle["header"],
            body_format=format_bundle["body"],
            cancel_token=cancel_token,
        )
        headers = schema.header_matrix[0]
        sample_data = SAMPLE_SETUP_DATA.get(schema.name, [])
        wrap_columns = tuple(idx for idx, header in enumerate(headers) if header in _WRAP_HEADER_NAMES)
        worksheet = _sheetops.write_schema_sheet(
            workbook=workbook,
            sheet_schema=schema,
            data=sample_data,
            header_format=format_bundle["header"],
            body_format=format_bundle["body"],
            cancel_token=cancel_token,
            wrap_columns=wrap_columns,
            wrapped_body_format=format_bundle["body_wrap"],
            wrapped_column_format=format_bundle["column_wrap"],
        )
        width_sample_rows: list[list[object]] = [list(headers)]
        width_sample_rows.extend(
            [
                [row[col] if col < len(row) else "" for col in range(len(headers))]
                for row in sample_data[: max(0, XLSX_AUTOFIT_SAMPLE_ROWS - 1)]
            ]
        )
        widths = compute_sampled_column_widths(
            width_sample_rows,
            max(0, len(headers) - 1),
            min_width=XLSX_AUTOFIT_MIN_WIDTH,
            max_width=XLSX_AUTOFIT_MAX_WIDTH,
            padding=XLSX_AUTOFIT_PADDING,
        )
        apply_xlsxwriter_column_widths(
            worksheet,
            widths,
            default_width=XLSX_AUTOFIT_MIN_WIDTH,
            wrap_columns=wrap_columns,
            wrap_format=format_bundle["column_wrap"],
        )
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        add_system_hash_sheet(workbook, ID_COURSE_SETUP)

        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        workbook.close()
        workbook_closed = True
        os.replace(temp_output, output)
    except ValidationError:
        _cleanup_incomplete_output()
        raise
    except JobCancelledError:
        _cleanup_incomplete_output()
        raise
    except Exception as exc:
        _cleanup_incomplete_output()
        _logger.exception(
            "Failed to generate CO description template. template_id=%s output=%s",
            ID_COURSE_SETUP,
            output,
        )
        raise AppSystemError(
            t("co_analysis.system.co_description_template_generate_failed", output=output)
        ) from exc
    return output


__all__ = ["generate_co_description_template"]
