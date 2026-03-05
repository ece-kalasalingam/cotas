from __future__ import annotations

import json
from hashlib import sha256
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")
pytest.importorskip("PySide6")

from common.constants import WORKBOOK_PASSWORD
from common.exceptions import ValidationError
from common.workbook_signing import sign_payload
from modules import instructor_module as instructor_ui
from modules.instructor.instructor_template_engine import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)


def _build_course_details(tmp_path: Path) -> Path:
    output = tmp_path / "course_details.xlsx"
    generate_course_details_template(output)
    return output


def _build_marks_template(tmp_path: Path) -> Path:
    details = _build_course_details(tmp_path)
    output = tmp_path / "marks_template.xlsx"
    generate_marks_template_from_course_details(details, output)
    return output


def test_high_volume_workbook_generation_validation_and_step3_schema(tmp_path: Path) -> None:
    details_path = _build_course_details(tmp_path)
    wb = openpyxl.load_workbook(details_path)
    try:
        metadata = wb["Course_Metadata"]
        metadata["B8"] = 12

        assessment = wb["Assessment_Config"]
        for row in range(2, 400):
            for col in "ABCDE":
                assessment[f"{col}{row}"] = None
        row = 2
        for idx in range(1, 10):
            assessment[f"A{row}"] = f"D{idx:02d}"
            assessment[f"B{row}"] = 10
            assessment[f"C{row}"] = "YES"
            assessment[f"D{row}"] = "YES"
            assessment[f"E{row}"] = "YES"
            row += 1
        assessment[f"A{row}"] = "DNON"
        assessment[f"B{row}"] = 10
        assessment[f"C{row}"] = "YES"
        assessment[f"D{row}"] = "NO"
        assessment[f"E{row}"] = "YES"
        row += 1
        for idx in range(1, 6):
            assessment[f"A{row}"] = f"I{idx:02d}"
            assessment[f"B{row}"] = 20
            assessment[f"C{row}"] = "NO"
            assessment[f"D{row}"] = "YES"
            assessment[f"E{row}"] = "NO"
            row += 1

        question_map = wb["Question_Map"]
        for row in range(2, 1000):
            for col in "ABCD":
                question_map[f"{col}{row}"] = None
        row = 2
        for idx in range(1, 10):
            question_map[f"A{row}"] = f"D{idx:02d}"
            question_map[f"B{row}"] = "Q1"
            question_map[f"C{row}"] = 10
            question_map[f"D{row}"] = idx
            row += 1
        question_map[f"A{row}"] = "DNON"
        question_map[f"B{row}"] = "Q1"
        question_map[f"C{row}"] = 100
        question_map[f"D{row}"] = "1,2,3,4"

        students = wb["Students"]
        for row in range(2, 1500):
            students[f"A{row}"] = None
            students[f"B{row}"] = None
        for idx in range(1, 501):
            row = idx + 1
            students[f"A{row}"] = f"R{idx:05d}"
            students[f"B{row}"] = f"Student {idx}"
        wb.save(details_path)
    finally:
        wb.close()

    assert validate_course_details_workbook(details_path) == "COURSE_SETUP_V1"
    marks_path = tmp_path / "marks_high_volume.xlsx"
    generate_marks_template_from_course_details(details_path, marks_path)
    instructor_ui._validate_uploaded_filled_marks_workbook(marks_path)

    generated = openpyxl.load_workbook(marks_path, data_only=False)
    try:
        assert len(generated.sheetnames) >= 19
    finally:
        generated.close()


def test_validation_rejects_partial_corrupted_workbook(tmp_path: Path) -> None:
    bad = tmp_path / "corrupted.xlsx"
    bad.write_bytes(b"this-is-not-a-valid-xlsx")

    with pytest.raises(ValidationError):
        validate_course_details_workbook(bad)
    with pytest.raises(ValidationError):
        instructor_ui._validate_uploaded_filled_marks_workbook(bad)


def test_step3_rejects_malformed_layout_manifest_json(tmp_path: Path) -> None:
    marks = _build_marks_template(tmp_path)
    wb = openpyxl.load_workbook(marks)
    try:
        wb["__SYSTEM_LAYOUT__"]["A2"] = "{broken-json"
        wb["__SYSTEM_LAYOUT__"]["B2"] = sign_payload("{broken-json")
        wb.save(marks)
    finally:
        wb.close()

    with pytest.raises(ValidationError, match="Layout manifest JSON is invalid"):
        instructor_ui._validate_uploaded_filled_marks_workbook(marks)


def test_step3_rejects_formula_tampering(tmp_path: Path) -> None:
    marks = _build_marks_template(tmp_path)
    wb = openpyxl.load_workbook(marks)
    try:
        manifest_text = str(wb["__SYSTEM_LAYOUT__"]["A2"].value)
        manifest = json.loads(manifest_text)
        target_sheet = None
        target_cell = None
        for spec in manifest.get("sheets", []):
            anchors = spec.get("formula_anchors", [])
            if anchors:
                target_sheet = spec["name"]
                target_cell = anchors[0][0]
                break
        assert target_sheet is not None and target_cell is not None
        wb[target_sheet][target_cell] = "=1+1"
        wb.save(marks)
    finally:
        wb.close()

    with pytest.raises(ValidationError, match="formula"):
        instructor_ui._validate_uploaded_filled_marks_workbook(marks)


def test_backward_compat_accepts_legacy_unsigned_hash_format(tmp_path: Path) -> None:
    details = _build_course_details(tmp_path)
    wb = openpyxl.load_workbook(details)
    try:
        template_id = str(wb["__SYSTEM_HASH__"]["A2"].value)
        legacy_hash = sha256(f"{template_id}|{WORKBOOK_PASSWORD}".encode("utf-8")).hexdigest()
        wb["__SYSTEM_HASH__"]["B2"] = legacy_hash
        wb.save(details)
    finally:
        wb.close()

    assert validate_course_details_workbook(details) == "COURSE_SETUP_V1"


def test_backward_compat_accepts_legacy_layout_hash_in_step3(tmp_path: Path) -> None:
    marks = _build_marks_template(tmp_path)
    wb = openpyxl.load_workbook(marks)
    try:
        manifest_text = str(wb["__SYSTEM_LAYOUT__"]["A2"].value)
        legacy_layout_hash = sha256(f"{manifest_text}|{WORKBOOK_PASSWORD}".encode("utf-8")).hexdigest()
        wb["__SYSTEM_LAYOUT__"]["B2"] = legacy_layout_hash
        wb.save(marks)
    finally:
        wb.close()

    instructor_ui._validate_uploaded_filled_marks_workbook(marks)
