from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.exceptions import JobCancelledError, ValidationError
from common.jobs import CancellationToken
from domain import instructor_template_engine as gen_mod
from domain.instructor_template_engine import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
)


def _build_course_details(tmp_path: Path) -> Path:
    source = tmp_path / "course_details.xlsx"
    generate_course_details_template(source)
    return source


def test_generate_marks_template_creates_expected_sheets(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"

    generated = generate_marks_template_from_course_details(source, output)

    assert generated == output
    assert output.exists()
    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.sheetnames[:2] == ["Course_Metadata", "Assessment_Config"]
        assert "__SYSTEM_HASH__" in workbook.sheetnames
        assert "__SYSTEM_LAYOUT__" in workbook.sheetnames
        assert "S1" in workbook.sheetnames
        assert "ESP" in workbook.sheetnames
        assert "CSURVEY" in workbook.sheetnames
        metadata_fields = [
            str(workbook["Course_Metadata"].cell(row=row, column=1).value or "")
            for row in range(2, 40)
        ]
        assert "Faculty_Name" not in metadata_fields
        assert workbook["Course_Metadata"].print_title_rows == "$1:$1"
        assert workbook["Assessment_Config"].print_title_rows == "$1:$1"
    finally:
        workbook.close()


def test_direct_co_wise_sheet_headers_formulas_and_validation(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["S1"]
        header_row = next(
            row for row in range(1, 80) if str(sheet.cell(row=row, column=1).value or "") == "#"
        )
        component_row = next(
            row
            for row in range(1, header_row)
            if str(sheet.cell(row=row, column=2).value or "") == "Component name"
        )
        assert sheet["B1"].value == "Course_Code"
        assert sheet["C1"].value == "ECE000"
        assert sheet.cell(row=component_row, column=2).value == "Component name"
        assert sheet.cell(row=component_row, column=3).value == "S1"
        assert sheet.cell(row=header_row, column=1).value == "#"
        assert sheet.cell(row=header_row, column=4).value == "Q1"
        assert sheet.cell(row=header_row, column=12).value == "Total"
        assert sheet.cell(row=header_row + 1, column=3).value == "CO"
        assert sheet.cell(row=header_row + 1, column=4).value == "CO1"
        assert sheet.cell(row=header_row + 2, column=3).value == "Max."
        assert sheet.cell(row=header_row + 2, column=4).value == 2
        assert sheet.cell(row=header_row + 3, column=1).value == 1
        assert sheet.cell(row=header_row + 3, column=2).value == "R101"
        assert sheet.cell(row=header_row + 3, column=3).value == "STUD1"
        assert sheet.cell(row=header_row + 3, column=12).value == f"=SUM(D{header_row + 3}:K{header_row + 3})"
        assert sheet.print_title_rows == f"$1:${header_row + 2}"
        active_cells = {selection.activeCell for selection in sheet.sheet_view.selection}
        assert f"D{header_row + 3}" in active_cells
        assert sheet.protection.sheet is True
        assert sheet.cell(row=header_row + 3, column=1).protection.locked is True
        assert sheet.cell(row=header_row + 3, column=4).protection.locked is False
        assert sheet.cell(row=header_row + 3, column=3).alignment.wrap_text is True
        assert sheet.page_setup.orientation == "landscape"
        assert sheet.page_setup.paperSize == 9
        assert sheet.page_margins.left == pytest.approx(0.25)
        assert sheet.page_margins.right == pytest.approx(0.25)
        assert sheet.column_dimensions["A"].width is not None
        assert sheet.column_dimensions["B"].width is not None
        assert sheet.column_dimensions["C"].width is not None
        assert sheet.column_dimensions["D"].width is not None
        assert sheet.column_dimensions["B"].width >= len("Component name") + 2
        validations = list(sheet.data_validations.dataValidation)
        assert validations
        assert any(f"D${header_row + 2}" in (validation.formula1 or "") for validation in validations)
        assert any("between 0 and 2" in (validation.error or "") for validation in validations)
    finally:
        workbook.close()


def test_direct_non_co_wise_sheet_layout_and_formulas(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["ESP"]
        header_row = next(
            row for row in range(1, 80) if str(sheet.cell(row=row, column=1).value or "") == "#"
        )
        component_row = next(
            row
            for row in range(1, header_row)
            if str(sheet.cell(row=row, column=2).value or "") == "Component name"
        )
        assert sheet["B1"].value == "Course_Code"
        assert sheet["C1"].value == "ECE000"
        assert sheet.cell(row=component_row, column=2).value == "Component name"
        assert sheet.cell(row=component_row, column=3).value == "ESP"
        assert sheet.cell(row=header_row, column=4).value == "Total"
        assert sheet.cell(row=header_row, column=5).value == "Marks for CO1"
        assert sheet.cell(row=header_row + 1, column=3).value == "CO"
        assert sheet.cell(row=header_row + 1, column=4).value in (None, "")
        assert sheet.cell(row=header_row + 1, column=5).value == "CO1"
        assert sheet.cell(row=header_row + 2, column=3).value == "Max."
        assert sheet.cell(row=header_row + 2, column=4).value == 100
        assert sheet.print_title_rows == f"$1:${header_row + 2}"
        co_marks = [sheet.cell(header_row + 2, col).value for col in range(5, sheet.max_column + 1)]
        assert co_marks
        assert sum(float(value) for value in co_marks) == pytest.approx(
            float(sheet.cell(row=header_row + 2, column=4).value), abs=0.01
        )
        assert sheet.cell(row=header_row + 3, column=1).value == 1
        active_cells = {selection.activeCell for selection in sheet.sheet_view.selection}
        assert f"D{header_row + 3}" in active_cells
        assert sheet.protection.sheet is True
        assert sheet.cell(row=header_row + 3, column=4).protection.locked is False
        assert sheet.cell(row=header_row + 3, column=3).alignment.wrap_text is True
        assert sheet.cell(row=header_row + 3, column=5).protection.locked is True
        assert "ROUND" in str(sheet.cell(row=header_row + 3, column=5).value)
        validations = list(sheet.data_validations.dataValidation)
        assert validations
        assert any(f"D${header_row + 2}" in (validation.formula1 or "") for validation in validations)
        assert any("between 0 and 100" in (validation.error or "") for validation in validations)
    finally:
        workbook.close()


def test_indirect_sheet_uses_total_outcomes_and_likert_validation(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["CSURVEY"]
        header_row = next(
            row for row in range(1, 80) if str(sheet.cell(row=row, column=1).value or "") == "#"
        )
        component_row = next(
            row
            for row in range(1, header_row)
            if str(sheet.cell(row=row, column=2).value or "") == "Component name"
        )
        assert sheet["B1"].value == "Course_Code"
        assert sheet["C1"].value == "ECE000"
        assert sheet.cell(row=component_row, column=2).value == "Component name"
        assert sheet.cell(row=component_row, column=3).value == "CSURVEY"
        assert sheet.cell(row=header_row, column=1).value == "#"
        assert sheet.cell(row=header_row, column=4).value == "CO1"
        assert sheet.cell(row=header_row, column=9).value == "CO6"
        assert sheet.cell(row=header_row + 1, column=1).value == 1
        assert sheet.print_title_rows == f"$1:${header_row}"
        active_cells = {selection.activeCell for selection in sheet.sheet_view.selection}
        assert f"D{header_row + 1}" in active_cells
        assert sheet.protection.sheet is True
        assert sheet.cell(row=header_row + 1, column=3).protection.locked is True
        assert sheet.cell(row=header_row + 1, column=4).protection.locked is False
        assert sheet.cell(row=header_row + 1, column=3).alignment.wrap_text is True
        validations = list(sheet.data_validations.dataValidation)
        assert validations
        formulas = [validation.formula1 or "" for validation in validations]
        assert any(">=1" in formula and "<=5" in formula for formula in formulas)
    finally:
        workbook.close()


def test_marks_template_initial_selection_on_setup_sheets(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook["Course_Metadata"].sheet_view.selection[0].activeCell == "A2"
        assert workbook["Assessment_Config"].sheet_view.selection[0].activeCell == "A2"
    finally:
        workbook.close()


def test_marks_template_manifest_includes_student_identity_fingerprint(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        manifest_text = workbook["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest_text, str)
        manifest = json.loads(manifest_text)
        component_specs = [
            spec
            for spec in manifest.get("sheets", [])
            if spec.get("kind") in {"direct_co_wise", "direct_non_co_wise", "indirect"}
        ]
        assert component_specs
        for spec in component_specs:
            assert isinstance(spec.get("student_count"), int)
            assert spec["student_count"] > 0
            assert isinstance(spec.get("student_identity_hash"), str)
            assert spec["student_identity_hash"]
            assert isinstance(spec.get("mark_structure"), dict)
            assert spec["mark_structure"]
    finally:
        workbook.close()


def test_generate_marks_template_honors_pre_cancel(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template_cancelled.xlsx"
    token = CancellationToken()
    token.cancel()

    with pytest.raises(JobCancelledError):
        generate_marks_template_from_course_details(source, output, cancel_token=token)


def test_generate_marks_template_rejects_unknown_template_version(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template_unknown.xlsx"
    workbook = openpyxl.load_workbook(source)
    try:
        workbook["__SYSTEM_HASH__"]["A2"] = "COURSE_SETUP_V2"
        workbook["__SYSTEM_HASH__"]["B2"] = gen_mod._compute_template_hash("COURSE_SETUP_V2")
        workbook.save(source)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="Unknown workbook template"):
        generate_marks_template_from_course_details(source, output)
