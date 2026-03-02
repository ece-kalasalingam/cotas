from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from modules.instructor.course_details_template_generator import (
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
        assert "S1" in workbook.sheetnames
        assert "ESP" in workbook.sheetnames
        assert "CSURVEY" in workbook.sheetnames
    finally:
        workbook.close()


def test_direct_co_wise_sheet_headers_formulas_and_validation(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["S1"]
        assert sheet["A1"].value == "Sl. No."
        assert sheet["D1"].value == "Q1"
        assert sheet["L1"].value == "Total"
        assert sheet["D2"].value == "CO1"
        assert sheet["D3"].value == 2
        assert sheet["A4"].value == 1
        assert sheet["B4"].value == "R101"
        assert sheet["C4"].value == "STUD1"
        assert sheet["L4"].value == "=SUM(D4:K4)"
        validations = list(sheet.data_validations.dataValidation)
        assert validations
        assert any("D$3" in (validation.formula1 or "") for validation in validations)
    finally:
        workbook.close()


def test_direct_non_co_wise_sheet_layout_and_formulas(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["ESP"]
        assert sheet["D1"].value == "Total"
        assert sheet["E1"].value == "Marks for CO1"
        assert sheet["D2"].value == "COs"
        assert sheet["E2"].value == "CO1"
        assert sheet["D3"].value == 100
        assert sheet["E3"].value == 20
        assert sheet["A4"].value == 0
        assert "/5" in str(sheet["E4"].value)
        validations = list(sheet.data_validations.dataValidation)
        assert validations
        assert any("D$3" in (validation.formula1 or "") for validation in validations)
    finally:
        workbook.close()


def test_indirect_sheet_uses_total_outcomes_and_likert_validation(tmp_path: Path) -> None:
    source = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(source, output)

    workbook = openpyxl.load_workbook(output)
    try:
        sheet = workbook["CSURVEY"]
        assert sheet["A1"].value == "Sl. No."
        assert sheet["D1"].value == "CO1"
        assert sheet["I1"].value == "CO6"
        assert sheet["A2"].value == 1
        validations = list(sheet.data_validations.dataValidation)
        assert validations
        formulas = [validation.formula1 or "" for validation in validations]
        assert any(">=1" in formula and "<=5" in formula for formula in formulas)
    finally:
        workbook.close()
