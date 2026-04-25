"""COURSE_SETUP_V2 course-template generation."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from uuid import uuid4

from common.constants import WORKBOOK_TEMP_SUFFIX
from common.error_catalog import validation_error_from_key
from common.excel_sheet_layout import build_template_xlsxwriter_formats
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.i18n import t
from common.jobs import CancellationToken
from common.registry import COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION
from common.runtime_dependency_guard import import_runtime_dependency
from common.registry import get_blueprint as _registry_get_blueprint
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.workbook_integrity import add_system_hash_sheet
from domain.template_versions.course_setup_v2_impl import (
    instructor_engine_sheetops as _shareops,
)

_logger = logging.getLogger(__name__)
_TEMPLATE_ID = "COURSE_SETUP_V2"


def generate_course_details_template(
    output_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    """Generate and save COURSE_SETUP_V2 course-details template workbook."""
    xlsxwriter = import_runtime_dependency("xlsxwriter")

    blueprint = _registry_get_blueprint(_TEMPLATE_ID)
    if blueprint is None:
        raise validation_error_from_key(
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
        """Cleanup incomplete output.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
            template_id=_TEMPLATE_ID,
        )
        header_format = format_bundle["header"]
        body_format = format_bundle["body"]

        for sheet_schema in blueprint.sheets:
            wrap_columns: tuple[int, ...] = ()
            wrapped_body_format = None
            wrapped_column_format = None
            if sheet_schema.key == COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION:
                wrap_columns = (1, 3)
                wrapped_body_format = format_bundle.get("body_wrap", body_format)
                wrapped_column_format = format_bundle.get("column_wrap", wrapped_body_format)
            _shareops.write_schema_sheet(
                workbook=workbook,
                sheet_schema=sheet_schema,
                data=SAMPLE_SETUP_DATA.get(sheet_schema.name, []),
                header_format=header_format,
                body_format=body_format,
                cancel_token=cancel_token,
                wrap_columns=wrap_columns,
                wrapped_body_format=wrapped_body_format,
                wrapped_column_format=wrapped_column_format,
            )

        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        add_system_hash_sheet(workbook, _TEMPLATE_ID)

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
            "Failed to generate course details template. template_id=%s output=%s",
            _TEMPLATE_ID,
            output,
        )
        raise AppSystemError(
            t("instructor.system.template_generate_failed", output=output)
        ) from exc
    return output


__all__ = [
    "generate_course_details_template",
]
