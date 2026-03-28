"""COURSE_SETUP_V2 course-template generation."""

from __future__ import annotations

import logging
import os
from collections.abc import Sequence
from pathlib import Path
from uuid import uuid4

from common.constants import WORKBOOK_TEMP_SUFFIX
from common.error_catalog import validation_error_from_key
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.excel_sheet_layout import build_template_xlsxwriter_formats
from common.i18n import t
from common.jobs import CancellationToken
from common.registry import get_blueprint as _registry_get_blueprint
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.workbook_integrity import add_system_hash_sheet
from domain.template_versions.course_setup_v2_impl import instructor_engine_sheetops as _shareops
from domain.template_versions.course_setup_v2_impl.course_semantics import (
    build_course_template_filename_base,
)

_logger = logging.getLogger(__name__)
_TEMPLATE_ID = "COURSE_SETUP_V2"


def generate_course_details_template(
    output_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    """Generate and save COURSE_SETUP_V2 course-details template workbook."""
    try:
        import xlsxwriter
    except ModuleNotFoundError as exc:
        raise validation_error_from_key(
            "instructor.validation.xlsxwriter_missing",
            code="XLSXWRITER_MISSING",
        ) from exc

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
            _shareops.write_schema_sheet(
                workbook=workbook,
                sheet_schema=sheet_schema,
                data=SAMPLE_SETUP_DATA.get(sheet_schema.name, []),
                header_format=header_format,
                body_format=body_format,
                cancel_token=cancel_token,
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


def generate_course_details_templates_batch(
    *,
    workbook_paths: Sequence[str | Path],
    output_dir: str | Path,
    allow_overwrite: bool = False,
    cancel_token: CancellationToken | None = None,
) -> dict[str, object]:
    unexpected_workbook_paths = [str(raw).strip() for raw in workbook_paths if str(raw).strip()]
    if unexpected_workbook_paths:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_PATHS_NOT_APPLICABLE",
            workbook_kind="course_details_template",
            expected="empty_sequence",
        )

    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    output_name = f"{build_course_template_filename_base()}.xlsx"
    output_path = output_root / output_name

    if output_path.exists() and not allow_overwrite:
        return {
            "total": 1,
            "generated": 0,
            "failed": 1,
            "skipped": 0,
            "generated_workbook_paths": [],
            "output_urls": [],
            "results": {
                "course_template": {
                    "status": "failed",
                    "source_path": None,
                    "workbook_path": None,
                    "output": None,
                    "output_path": None,
                    "output_url": None,
                    "reason": "output_already_exists",
                    "existing_output_path": str(output_path),
                }
            },
        }

    results: dict[str, object] = {}
    try:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        generated_path = generate_course_details_template(
            output_path=output_path,
            cancel_token=cancel_token,
        )
        output_value = str(generated_path)
        results["course_template"] = {
            "status": "generated",
            "source_path": None,
            "workbook_path": output_value,
            "output": output_value,
            "output_path": output_value,
            "output_url": output_value,
            "reason": None,
        }
        generated = 1
        failed = 0
    except JobCancelledError:
        raise
    except Exception as exc:
        results["course_template"] = {
            "status": "failed",
            "source_path": None,
            "workbook_path": None,
            "output": None,
            "output_path": None,
            "output_url": None,
            "reason": str(exc),
        }
        generated = 0
        failed = 1

    generated_paths = [
        str(entry.get("workbook_path"))
        for entry in results.values()
        if isinstance(entry, dict) and str(entry.get("status")) == "generated" and entry.get("workbook_path")
    ]

    return {
        "total": 1,
        "generated": generated,
        "failed": failed,
        "skipped": 0,
        "generated_workbook_paths": generated_paths,
        "output_urls": list(generated_paths),
        "results": results,
    }


__all__ = [
    "generate_course_details_template",
    "generate_course_details_templates_batch",
]
