from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import ID_COURSE_SETUP
from common.exceptions import ValidationError
from modules.instructor.course_details_template_generator import (
    generate_course_details_template,
    validate_course_details_workbook,
)


def _build_workbook(tmp_path: Path) -> Path:
    output = tmp_path / "course_setup.xlsx"
    generate_course_details_template(output)
    return output


def test_validate_uploaded_workbook_accepts_generated_template(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    assert validate_course_details_workbook(output) == ID_COURSE_SETUP


def test_validate_uploaded_workbook_rejects_hash_tampering(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["__SYSTEM_HASH__"]["B2"] = "bad-hash"
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="Template hash mismatch"):
        validate_course_details_workbook(output)


def test_validate_uploaded_workbook_rejects_bad_direct_weight_total(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["Assessment_Config"]["B2"] = 1
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="direct component weights must total 100"):
        validate_course_details_workbook(output)


def test_validate_uploaded_workbook_rejects_multi_co_for_co_wise_component(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["Question_Map"]["D2"] = "CO1, CO2"
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="requires exactly one CO per question"):
        validate_course_details_workbook(output)


def test_validate_uploaded_workbook_rejects_duplicate_student_reg_no(tmp_path: Path) -> None:
    output = _build_workbook(tmp_path)
    workbook = openpyxl.load_workbook(output)
    try:
        workbook["Students"]["A3"] = workbook["Students"]["A2"].value
        workbook.save(output)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="duplicate Reg_No"):
        validate_course_details_workbook(output)
