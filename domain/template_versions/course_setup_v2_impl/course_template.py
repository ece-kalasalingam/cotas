"""COURSE_SETUP_V2 course-template generation."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any
from uuid import uuid4

from common.constants import WORKBOOK_TEMP_SUFFIX
from common.error_catalog import validation_error_from_key
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.registry import get_blueprint as _registry_get_blueprint
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.i18n import t
_logger = logging.getLogger(__name__)
_TEMPLATE_ID = "COURSE_SETUP_V2"
from domain.template_versions.course_setup_v2_impl import instructor_engine_shareops as _shareops

_add_system_hash_sheet = _shareops.add_system_hash_sheet
_apply_validation = _shareops.apply_validation
_build_body_format = _shareops.build_body_format
_build_header_format = _shareops.build_header_format
_protect_sheet = _shareops.protect_sheet
generate_worksheet = _shareops.generate_worksheet


def _ve(translation_key: str, *, code: str, **context: Any) -> ValidationError:
    return validation_error_from_key(translation_key, code=code, context=context)


def generate_course_details_template(
    output_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    """Generate and save COURSE_SETUP_V2 course-details template workbook."""
    try:
        import xlsxwriter
    except ModuleNotFoundError as exc:
        raise _ve(
            "instructor.validation.xlsxwriter_missing",
            code="XLSXWRITER_MISSING",
        ) from exc

    blueprint = _registry_get_blueprint(_TEMPLATE_ID)
    if blueprint is None:
        raise _ve(
            "validation.template.unknown",
            code="UNKNOWN_TEMPLATE",
            template_id=_TEMPLATE_ID,
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
        header_format = _build_header_format(workbook, blueprint.style_registry.get("header", {}))
        body_format = _build_body_format(workbook, blueprint.style_registry.get("body", {}))

        for sheet_schema in blueprint.sheets:
            if cancel_token is not None:
                cancel_token.raise_if_cancelled()
            if len(sheet_schema.header_matrix) != 1:
                raise validation_error_from_key(
                    "instructor.validation.sheet_single_header_row",
                    sheet_name=sheet_schema.name,
                )
            worksheet = generate_worksheet(
                workbook=workbook,
                sheet_name=sheet_schema.name,
                headers=sheet_schema.header_matrix[0],
                data=SAMPLE_SETUP_DATA.get(sheet_schema.name, []),
                header_format=header_format,
                body_format=body_format,
            )
            for validation in sheet_schema.validations:
                if cancel_token is not None:
                    cancel_token.raise_if_cancelled()
                _apply_validation(worksheet, validation)
            if sheet_schema.is_protected:
                _protect_sheet(worksheet)

        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        _add_system_hash_sheet(workbook, _TEMPLATE_ID)

        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        workbook.close()
        workbook_closed = True
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
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
            "Failed to generate course details template. template_id=%s output=%s",
            _TEMPLATE_ID,
            output,
        )
        raise AppSystemError(
            t("instructor.system.template_generate_failed", output=output)
        ) from exc
    return output


__all__ = ["generate_course_details_template"]

