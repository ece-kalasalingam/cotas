from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("xlsxwriter")
openpyxl = pytest.importorskip("openpyxl")

from common.constants import ID_COURSE_SETUP
from domain.template_strategy_router import generate_workbook


def generate_course_details_template(output_path: Path) -> Path:
    """Generate course details template.
    
    Args:
        output_path: Parameter value (Path).
    
    Returns:
        Path: Return value.
    
    Raises:
        None.
    """
    result = generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=output_path,
        workbook_name=output_path.name,
        workbook_kind="course_details_template",
    )
    output = getattr(result, "output_path", None)
    if isinstance(output, str) and output.strip():
        return Path(output)
    return output_path


def generate_co_description_template(output_path: Path) -> Path:
    """Generate co description template.
    
    Args:
        output_path: Parameter value (Path).
    
    Returns:
        Path: Return value.
    
    Raises:
        None.
    """
    result = generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=output_path,
        workbook_name=output_path.name,
        workbook_kind="co_description_template",
    )
    output = getattr(result, "output_path", None)
    if isinstance(output, str) and output.strip():
        return Path(output)
    return output_path


def test_generated_co_description_template_structure_and_integrity(tmp_path: Path) -> None:
    """Test generated co description template structure and integrity.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    output = tmp_path / "co_description_template.xlsx"
    generate_co_description_template(output)

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.sheetnames == ["Course_Metadata", "CO_Description", "__SYSTEM_HASH__"]

        course_sheet = workbook["Course_Metadata"]
        assert course_sheet["A1"].value == "Field"
        assert course_sheet["B1"].value == "Value"
        assert course_sheet["A2"].value == "Course_Code"
        assert course_sheet["B2"].value == "ECE000"

        co_desc_sheet = workbook["CO_Description"]
        assert co_desc_sheet["A1"].value == "CO#"
        assert co_desc_sheet["C1"].value == "Domain_Level"
        assert co_desc_sheet["D1"].value == "Summary_of_Topics/Expts./Project"
        assert co_desc_sheet["A2"].value == 1
        assert bool(co_desc_sheet["B1"].alignment.wrap_text)
        assert bool(co_desc_sheet["D1"].alignment.wrap_text)
        assert bool(co_desc_sheet["B2"].alignment.wrap_text)
        assert bool(co_desc_sheet["D2"].alignment.wrap_text)
        co_validations = list(co_desc_sheet.data_validations.dataValidation)
        assert co_validations
        assert any("A2:A301" in str(validation.sqref) for validation in co_validations)
        assert any("C2:C301" in str(validation.sqref) for validation in co_validations)
        assert any("D2:D301" in str(validation.sqref) for validation in co_validations)

        hash_sheet = workbook["__SYSTEM_HASH__"]
        assert hash_sheet.sheet_state == "hidden"
        assert hash_sheet["A2"].value == ID_COURSE_SETUP
    finally:
        workbook.close()


def test_generated_workbook_structure_and_prefill_data(tmp_path: Path) -> None:
    """Test generated workbook structure and prefill data.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    output = tmp_path / "course_setup.xlsx"
    generate_course_details_template(output)

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook.sheetnames == [
            "Course_Metadata",
            "Assessment_Config",
            "Question_Map",
            "CO_Description",
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

        co_desc_sheet = workbook["CO_Description"]
        assert co_desc_sheet["A1"].value == "CO#"
        assert co_desc_sheet["C1"].value == "Domain_Level"
        assert co_desc_sheet["D1"].value == "Summary_of_Topics/Expts./Project"
        assert bool(co_desc_sheet["B1"].alignment.wrap_text)
        assert bool(co_desc_sheet["D1"].alignment.wrap_text)
        assert bool(co_desc_sheet["B2"].alignment.wrap_text)
        assert bool(co_desc_sheet["D2"].alignment.wrap_text)
        co_validations = list(co_desc_sheet.data_validations.dataValidation)
        assert co_validations
        assert any("A2:A301" in str(validation.sqref) for validation in co_validations)
        assert any("C2:C301" in str(validation.sqref) for validation in co_validations)
        assert any("D2:D301" in str(validation.sqref) for validation in co_validations)

        hash_sheet = workbook["__SYSTEM_HASH__"]
        assert hash_sheet.sheet_state == "hidden"
        assert hash_sheet["A1"].value == "Template_ID"
        assert hash_sheet["B1"].value == "Template_Hash"
        assert hash_sheet["A2"].value == ID_COURSE_SETUP
        assert isinstance(hash_sheet["B2"].value, str)
        assert str(hash_sheet["B2"].value).strip() != ""
    finally:
        workbook.close()


def test_generated_workbook_overwrites_existing_file_atomically(tmp_path: Path) -> None:
    """Test generated workbook overwrites existing file atomically.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    output = tmp_path / "course_setup.xlsx"
    output.write_text("stale", encoding="utf-8")

    generate_course_details_template(output)

    workbook = openpyxl.load_workbook(output)
    try:
        assert workbook["Students"]["A2"].value == "R101"
        assert workbook["Students"]["B2"].value == "STUD1"
    finally:
        workbook.close()
