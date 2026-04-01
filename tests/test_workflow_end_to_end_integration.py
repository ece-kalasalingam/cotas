from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import ID_COURSE_SETUP
from common.jobs import CancellationToken
from domain.template_strategy_router import generate_workbook, validate_workbooks
from domain.template_versions.course_setup_v2_impl.co_attainment import (
    extract_final_report_signature_from_path,
    generate_final_report_workbook,
)
from services.instructor_workflow_service import InstructorWorkflowService


def _set_course_section(course_details_path: Path, section: str) -> None:
    """Set course section.
    
    Args:
        course_details_path: Parameter value (Path).
        section: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(course_details_path)
    try:
        ws = wb["Course_Metadata"]
        row = 2
        while True:
            key = str(ws.cell(row=row, column=1).value or "").strip().casefold()
            value = ws.cell(row=row, column=2).value
            if not key and (value is None or str(value).strip() == ""):
                break
            if key == "section":
                ws.cell(row=row, column=2, value=section)
                break
            row += 1
        wb.save(course_details_path)
    finally:
        wb.close()


def _prefix_student_regnos(course_details_path: Path, prefix: str) -> None:
    """Prefix student regnos.
    
    Args:
        course_details_path: Parameter value (Path).
        prefix: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(course_details_path)
    try:
        ws = wb["Students"]
        row = 2
        while True:
            reg = ws.cell(row=row, column=1).value
            name = ws.cell(row=row, column=2).value
            if reg is None and name is None:
                break
            reg_text = str(reg).strip() if reg is not None else ""
            if reg_text:
                ws.cell(row=row, column=1, value=f"{prefix}{reg_text}")
            row += 1
        wb.save(course_details_path)
    finally:
        wb.close()


def _fill_marks_workbook(marks_path: Path, mark_value: float = 1.0) -> None:
    """Fill marks workbook.
    
    Args:
        marks_path: Parameter value (Path).
        mark_value: Parameter value (float).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(marks_path)
    try:
        manifest_text = wb["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest_text, str)
        manifest = json.loads(manifest_text)
        for spec in manifest.get("sheets", []):
            kind = spec.get("kind")
            if kind not in {"direct_co_wise", "direct_non_co_wise", "indirect"}:
                continue
            ws = wb[str(spec["name"])]
            header_row = int(spec["header_row"])
            header_count = len(spec["headers"])
            if kind == "indirect":
                first_data_row = header_row + 1
                mark_cols = range(4, header_count + 1)
            elif kind == "direct_non_co_wise":
                first_data_row = header_row + 3
                mark_cols = range(4, 5)
            else:
                first_data_row = header_row + 3
                mark_cols = range(4, header_count)

            row = first_data_row
            while True:
                reg_no = ws.cell(row=row, column=2).value
                student_name = ws.cell(row=row, column=3).value
                if reg_no is None and student_name is None:
                    break
                for col in mark_cols:
                    ws.cell(row=row, column=col, value=mark_value)
                row += 1
        wb.save(marks_path)
    finally:
        wb.close()


def _build_final_report(
    root: Path,
    *,
    section: str,
    reg_prefix: str = "",
) -> Path:
    """Build final report.
    
    Args:
        root: Parameter value (Path).
        section: Parameter value (str).
        reg_prefix: Parameter value (str).
    
    Returns:
        Path: Return value.
    
    Raises:
        None.
    """
    instructor = InstructorWorkflowService()
    course_details = root / f"course_details_{section}.xlsx"
    marks_template = root / f"marks_template_{section}.xlsx"
    final_report = root / f"final_report_{section}.xlsx"

    generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=course_details,
        workbook_name=course_details.name,
        workbook_kind="course_details_template",
        cancel_token=CancellationToken(),
    )
    _set_course_section(course_details, section)
    if reg_prefix:
        _prefix_student_regnos(course_details, reg_prefix)

    context_prepare = instructor.create_job_context(step_id=f"prepare_marks_{section}")
    validate_workbooks(
        template_id=ID_COURSE_SETUP,
        workbook_paths=[course_details],
        workbook_kind="course_details",
        cancel_token=CancellationToken(),
    )
    instructor.generate_marks_template(
        course_details,
        marks_template,
        context=context_prepare,
        cancel_token=CancellationToken(),
    )

    _fill_marks_workbook(marks_template, mark_value=1.0)
    generate_final_report_workbook(
        filled_marks_path=marks_template,
        output_path=final_report,
        cancel_token=CancellationToken(),
    )
    return final_report


def test_instructor_workflow_service_end_to_end_generates_signed_final_report(tmp_path: Path) -> None:
    """Test instructor workflow service end to end generates signed final report.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    final_report = _build_final_report(tmp_path, section="A")
    assert final_report.exists()
    assert extract_final_report_signature_from_path(final_report) is not None



