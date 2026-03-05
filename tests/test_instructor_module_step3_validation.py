from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")
pytest.importorskip("PySide6")

from common.exceptions import ValidationError
from modules import instructor_module as instructor_ui
from modules.instructor.instructor_template_engine import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
)


def _build_filled_marks_template(tmp_path: Path) -> Path:
    course_details = tmp_path / "course_details.xlsx"
    marks_template = tmp_path / "marks_template.xlsx"
    generate_course_details_template(course_details)
    generate_marks_template_from_course_details(course_details, marks_template)
    return marks_template


def test_step3_validation_accepts_generated_marks_template(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_missing_layout_sheet(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        del workbook["__SYSTEM_LAYOUT__"]
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="layout sheet"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_layout_manifest_hash_tampering(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        workbook["__SYSTEM_LAYOUT__"]["B2"] = "bad-layout-hash"
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="Layout hash mismatch"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_dynamic_header_tampering(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        workbook["S1"]["A10"] = "Tampered"
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="header mismatch"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)
