"""Template-routing facade for final CO report generation."""

from __future__ import annotations

from pathlib import Path

from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from domain.template_strategy_router import (
    generate_workbook,
    read_valid_template_id_from_system_hash_sheet,
)


def generate_final_co_report(
    filled_marks_path: str | Path,
    output_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    try:
        import openpyxl
    except ModuleNotFoundError as exc:
        raise validation_error_from_key("instructor.validation.openpyxl_missing") from exc

    source = Path(filled_marks_path)
    if not source.exists():
        raise validation_error_from_key("instructor.validation.workbook_not_found", workbook=source)

    try:
        workbook = openpyxl.load_workbook(source, data_only=False, read_only=True)
    except Exception as exc:
        raise validation_error_from_key("instructor.validation.workbook_open_failed", workbook=source) from exc
    try:
        template_id = read_valid_template_id_from_system_hash_sheet(workbook)
    finally:
        workbook.close()

    result = generate_workbook(
        template_id=template_id,
        output_path=Path(output_path),
        workbook_name=Path(output_path).name,
        workbook_kind="final_report",
        cancel_token=cancel_token,
        context={"filled_marks_path": str(source)},
    )
    if isinstance(result, Path):
        return result
    output = getattr(result, "output_path", None)
    if isinstance(output, Path):
        return output
    if isinstance(output, str) and output.strip():
        return Path(output)
    return Path(output_path)
