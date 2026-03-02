from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("xlsxwriter")
openpyxl = pytest.importorskip("openpyxl")

from modules.instructor.course_details_template_generator import (
    _compute_template_hash,
    generate_course_details_template,
)


def test_generated_workbook_structure_and_prefill_data(tmp_path: Path) -> None:
    output = tmp_path / "course_setup.xlsx"
    generate_course_details_template(output)

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.sheetnames == [
            "Course_Metadata",
            "Assessment_Config",
            "Question_Map",
            "Students",
            "__SYSTEM_HASH__",
        ]

        course_sheet = workbook["Course_Metadata"]
        assert course_sheet["A1"].value == "Field"
        assert course_sheet["B1"].value == "Value"
        assert course_sheet["A2"].value == "Course_Code"
        assert course_sheet["B2"].value == "ECE000"

        assessment_sheet = workbook["Assessment_Config"]
        assert assessment_sheet["A1"].value == "Component"
        assert assessment_sheet["E2"].value == "YES"

        validations = list(assessment_sheet.data_validations.dataValidation)
        assert validations
        assert "E2:E301" in str(validations[0].sqref)

        hash_sheet = workbook["__SYSTEM_HASH__"]
        assert hash_sheet.sheet_state == "hidden"
        assert hash_sheet["A1"].value == "Template_ID"
        assert hash_sheet["B1"].value == "Template_Hash"
        assert hash_sheet["A2"].value == "COURSE_SETUP_V1"
        assert hash_sheet["B2"].value == _compute_template_hash("COURSE_SETUP_V1")
    finally:
        workbook.close()


def test_generated_workbook_overwrites_existing_file_atomically(tmp_path: Path) -> None:
    output = tmp_path / "course_setup.xlsx"
    output.write_text("stale", encoding="utf-8")

    generate_course_details_template(output)

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook["Students"]["A2"].value == "R101"
        assert workbook["Students"]["B2"].value == "STUD1"
    finally:
        workbook.close()
