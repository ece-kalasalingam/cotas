"""Template-routing facade for course-template and marks-template workflows."""

from __future__ import annotations

from pathlib import Path

from common.constants import ID_COURSE_SETUP
from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from domain.template_strategy_router import (
    generate_workbook,
    read_valid_template_id_from_system_hash_sheet,
    validate_workbook,
)


def generate_course_details_template(
    output_path: str | Path,
    template_id: str = ID_COURSE_SETUP,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    result = generate_workbook(
        template_id=template_id,
        output_path=output_path,
        workbook_name=Path(output_path).name,
        workbook_kind="course_details_template",
        cancel_token=cancel_token,
    )
    return _result_to_path(result, fallback=Path(output_path))


def generate_marks_template_from_course_details(
    course_details_path: str | Path,
    output_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    source = Path(course_details_path)
    template_id = _resolve_template_id_from_workbook(source)
    result = generate_workbook(
        template_id=template_id,
        output_path=output_path,
        workbook_name=Path(output_path).name,
        workbook_kind="marks_template",
        cancel_token=cancel_token,
        context={"course_details_path": str(source)},
    )
    return _result_to_path(result, fallback=Path(output_path))


def _result_to_path(result: object, *, fallback: Path) -> Path:
    if isinstance(result, Path):
        return result
    output = getattr(result, "output_path", None)
    if isinstance(output, Path):
        return output
    if isinstance(output, str) and output.strip():
        return Path(output)
    return fallback


def _resolve_template_id_from_workbook(workbook_path: str | Path) -> str:
    try:
        import openpyxl
    except ModuleNotFoundError as exc:
        raise validation_error_from_key(
            "validation.dependency.openpyxl_missing",
            code="OPENPYXL_MISSING",
        ) from exc
    source = Path(workbook_path)
    try:
        workbook = openpyxl.load_workbook(source, data_only=False, read_only=True)
    except Exception as exc:
        raise validation_error_from_key(
            "validation.workbook.open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(source),
        ) from exc
    try:
        return read_valid_template_id_from_system_hash_sheet(workbook)
    finally:
        workbook.close()


def validate_course_details_workbook(
    workbook_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> str:
    return validate_workbook(
        workbook_path=workbook_path,
        workbook_kind="course_details",
        cancel_token=cancel_token,
    )


__all__ = [
    "generate_course_details_template",
    "generate_marks_template_from_course_details",
    "validate_course_details_workbook",
]
