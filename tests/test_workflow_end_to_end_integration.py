from __future__ import annotations

import json
from pathlib import Path
from typing import Any, cast

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.jobs import CancellationToken
from domain import coordinator_engine as coordinator_processing
from services.coordinator_workflow_service import CoordinatorWorkflowService
from services.instructor_workflow_service import InstructorWorkflowService


def _set_course_section(course_details_path: Path, section: str) -> None:
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
    instructor = InstructorWorkflowService()
    course_details = root / f"course_details_{section}.xlsx"
    marks_template = root / f"marks_template_{section}.xlsx"
    final_report = root / f"final_report_{section}.xlsx"

    context_template = instructor.create_job_context(step_id=f"generate_course_template_{section}")
    instructor.generate_course_details_template(course_details, context=context_template, cancel_token=CancellationToken())
    _set_course_section(course_details, section)
    if reg_prefix:
        _prefix_student_regnos(course_details, reg_prefix)

    context_prepare = instructor.create_job_context(step_id=f"prepare_marks_{section}")
    instructor.validate_course_details_workbook(course_details, context=context_prepare, cancel_token=CancellationToken())
    instructor.generate_marks_template(
        course_details,
        marks_template,
        context=context_prepare,
        cancel_token=CancellationToken(),
    )

    _fill_marks_workbook(marks_template, mark_value=1.0)
    context_report = instructor.create_job_context(step_id=f"generate_final_report_{section}")
    instructor.generate_final_report(
        marks_template,
        final_report,
        context=context_report,
        cancel_token=CancellationToken(),
    )
    return final_report


def test_instructor_workflow_service_end_to_end_generates_signed_final_report(tmp_path: Path) -> None:
    final_report = _build_final_report(tmp_path, section="A")
    assert final_report.exists()
    assert coordinator_processing._has_valid_final_co_report(final_report) is True


def test_coordinator_workflow_service_end_to_end_collects_and_calculates(tmp_path: Path) -> None:
    report_a = _build_final_report(tmp_path, section="A", reg_prefix="A-")
    report_b = _build_final_report(tmp_path, section="B", reg_prefix="B-")

    coordinator = CoordinatorWorkflowService()
    collect_ctx = coordinator.create_job_context(step_id="collect")
    collected = coordinator.collect_files(
        [str(report_a), str(report_b)],
        existing_keys=set(),
        existing_paths=[],
        context=collect_ctx,
        cancel_token=CancellationToken(),
    )
    assert collected["added"] == [str(report_a.resolve()), str(report_b.resolve())]
    assert collected["invalid_final_report"] == []

    output = tmp_path / "co_attainment.xlsx"
    calc_ctx = coordinator.create_job_context(step_id="calculate")
    result = coordinator.calculate_attainment(
        [report_a, report_b],
        output,
        context=calc_ctx,
        cancel_token=CancellationToken(),
    )
    assert output.exists()
    assert cast(Any, result).output_path == output

    wb = openpyxl.load_workbook(output, data_only=True)
    try:
        assert "CO1" in wb.sheetnames
        ws = wb["CO1"]
        assert ws["A12"].value == "#"
        assert ws["B12"].value == "Regno"
        assert ws["C12"].value == "Student name"
        assert ws["A13"].value == 1
        assert isinstance(ws["B13"].value, str)
    finally:
        wb.close()

