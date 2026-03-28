"""COURSE_SETUP_V2 CO description template generation."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from uuid import uuid4

from common.constants import ID_COURSE_SETUP, WORKBOOK_TEMP_SUFFIX
from common.error_catalog import validation_error_from_key
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.excel_sheet_layout import build_template_xlsxwriter_formats
from common.i18n import t
from common.jobs import CancellationToken
from common.registry import COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION, get_sheet_schema_by_key
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.workbook_integrity import add_system_hash_sheet
from domain.template_versions.course_setup_v2_impl import instructor_engine_sheetops as _sheetops

_logger = logging.getLogger(__name__)


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

    schema = get_sheet_schema_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION)
    if schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SHEET_SCHEMA_MISSING",
            sheet_key=COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
            template_id=ID_COURSE_SETUP,
        )
    if len(schema.header_matrix) != 1:
        raise validation_error_from_key(
            "instructor.validation.sheet_single_header_row",
            code="SHEET_HEADER_MATRIX_INVALID",
            sheet_name=schema.name,
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
        format_bundle = build_template_xlsxwriter_formats(workbook, template_id=ID_COURSE_SETUP)
        worksheet = _sheetops.generate_worksheet(
            workbook=workbook,
            sheet_name=schema.name,
            headers=schema.header_matrix[0],
            data=SAMPLE_SETUP_DATA.get(schema.name, []),
            header_format=format_bundle["header"],
            body_format=format_bundle["body"],
        )
        for validation in schema.validations:
            if cancel_token is not None:
                cancel_token.raise_if_cancelled()
            _sheetops._apply_validation(worksheet, validation)
        if schema.is_protected:
            _sheetops._protect_sheet(worksheet)

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

