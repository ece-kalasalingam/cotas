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
        if not (workbook.sheetnames == ["Course_Metadata", "CO_Description", "__SYSTEM_HASH__"]):
            raise AssertionError('assertion failed')

        course_sheet = workbook["Course_Metadata"]
        if not (course_sheet["A1"].value == "Field"):
            raise AssertionError('assertion failed')
        if not (course_sheet["B1"].value == "Value"):
            raise AssertionError('assertion failed')
        if not (course_sheet["A2"].value == "Course_Code"):
            raise AssertionError('assertion failed')
        if not (course_sheet["B2"].value == "ECE000"):
            raise AssertionError('assertion failed')

        co_desc_sheet = workbook["CO_Description"]
        if not (co_desc_sheet["A1"].value == "CO#"):
            raise AssertionError('assertion failed')
        if not (co_desc_sheet["C1"].value == "Domain_Level"):
            raise AssertionError('assertion failed')
        if not (co_desc_sheet["D1"].value == "Summary_of_Topics/Expts./Project"):
            raise AssertionError('assertion failed')
        if not (co_desc_sheet["A2"].value == 1):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["B1"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["D1"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["B2"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["D2"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        co_validations = list(co_desc_sheet.data_validations.dataValidation)
        if not (co_validations):
            raise AssertionError('assertion failed')
        if not (any("A2:A301" in str(validation.sqref) for validation in co_validations)):
            raise AssertionError('assertion failed')
        if not (any("C2:C301" in str(validation.sqref) for validation in co_validations)):
            raise AssertionError('assertion failed')
        if not (any("D2:D301" in str(validation.sqref) for validation in co_validations)):
            raise AssertionError('assertion failed')

        hash_sheet = workbook["__SYSTEM_HASH__"]
        if not (hash_sheet.sheet_state == "hidden"):
            raise AssertionError('assertion failed')
        if not (hash_sheet["A2"].value == ID_COURSE_SETUP):
            raise AssertionError('assertion failed')
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
        if not (workbook.sheetnames == [
            "Course_Metadata",
            "Assessment_Config",
            "Question_Map",
            "CO_Description",
            "Students",
            "__SYSTEM_HASH__",
        ]):
            raise AssertionError('assertion failed')

        course_sheet = workbook["Course_Metadata"]
        if not (course_sheet["A1"].value == "Field"):
            raise AssertionError('assertion failed')
        if not (course_sheet["B1"].value == "Value"):
            raise AssertionError('assertion failed')
        if not (course_sheet["A2"].value == "Course_Code"):
            raise AssertionError('assertion failed')
        if not (course_sheet["B2"].value == "ECE000"):
            raise AssertionError('assertion failed')

        assessment_sheet = workbook["Assessment_Config"]
        if not (assessment_sheet["A1"].value == "Component"):
            raise AssertionError('assertion failed')
        if not (assessment_sheet["E2"].value == "YES"):
            raise AssertionError('assertion failed')

        validations = list(assessment_sheet.data_validations.dataValidation)
        if not (validations):
            raise AssertionError('assertion failed')
        if "E2:E301" not in str(validations[0].sqref):
            raise AssertionError('assertion failed')

        co_desc_sheet = workbook["CO_Description"]
        if not (co_desc_sheet["A1"].value == "CO#"):
            raise AssertionError('assertion failed')
        if not (co_desc_sheet["C1"].value == "Domain_Level"):
            raise AssertionError('assertion failed')
        if not (co_desc_sheet["D1"].value == "Summary_of_Topics/Expts./Project"):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["B1"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["D1"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["B2"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        if not (bool(co_desc_sheet["D2"].alignment.wrap_text)):
            raise AssertionError('assertion failed')
        co_validations = list(co_desc_sheet.data_validations.dataValidation)
        if not (co_validations):
            raise AssertionError('assertion failed')
        if not (any("A2:A301" in str(validation.sqref) for validation in co_validations)):
            raise AssertionError('assertion failed')
        if not (any("C2:C301" in str(validation.sqref) for validation in co_validations)):
            raise AssertionError('assertion failed')
        if not (any("D2:D301" in str(validation.sqref) for validation in co_validations)):
            raise AssertionError('assertion failed')

        hash_sheet = workbook["__SYSTEM_HASH__"]
        if not (hash_sheet.sheet_state == "hidden"):
            raise AssertionError('assertion failed')
        if not (hash_sheet["A1"].value == "Template_ID"):
            raise AssertionError('assertion failed')
        if not (hash_sheet["B1"].value == "Template_Hash"):
            raise AssertionError('assertion failed')
        if not (hash_sheet["A2"].value == ID_COURSE_SETUP):
            raise AssertionError('assertion failed')
        if not (isinstance(hash_sheet["B2"].value, str)):
            raise AssertionError('assertion failed')
        if not (str(hash_sheet["B2"].value).strip() != ""):
            raise AssertionError('assertion failed')
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
        if not (workbook["Students"]["A2"].value == "R101"):
            raise AssertionError('assertion failed')
        if not (workbook["Students"]["B2"].value == "STUD1"):
            raise AssertionError('assertion failed')
    finally:
        workbook.close()
