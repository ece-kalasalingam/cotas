from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")
pytest.importorskip("PySide6")

from common.exceptions import ValidationError
from modules import instructor_module as instructor_ui
from domain.instructor_template_engine import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
)


def _build_filled_marks_template(tmp_path: Path) -> Path:
    course_details = tmp_path / "course_details.xlsx"
    marks_template = tmp_path / "marks_template.xlsx"
    generate_course_details_template(course_details)
    generate_marks_template_from_course_details(course_details, marks_template)
    _fill_all_mark_entries(marks_template)
    return marks_template


def _sheet_rows_from_manifest(workbook: object, sheet_name: str) -> tuple[int, int]:
    manifest_text = workbook["__SYSTEM_LAYOUT__"]["A2"].value
    assert isinstance(manifest_text, str)
    import json

    manifest = json.loads(manifest_text)
    for spec in manifest.get("sheets", []):
        if spec.get("name") != sheet_name:
            continue
        header_row = int(spec["header_row"])
        kind = spec.get("kind")
        if kind == "indirect":
            return header_row, header_row + 1
        return header_row, header_row + 3
    raise AssertionError(f"Sheet spec not found for {sheet_name}")


def test_step3_validation_accepts_generated_marks_template(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def _fill_all_mark_entries(workbook_path: Path) -> None:
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        manifest = workbook["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest, str)
        import json

        layout = json.loads(manifest)
        for spec in layout.get("sheets", []):
            kind = spec.get("kind")
            if kind not in {"direct_co_wise", "direct_non_co_wise", "indirect"}:
                continue
            sheet = workbook[spec["name"]]
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
            student_count = 0
            row = first_data_row
            while True:
                reg_no = sheet.cell(row=row, column=2).value
                student_name = sheet.cell(row=row, column=3).value
                if reg_no is None and student_name is None:
                    break
                student_count += 1
                row += 1
            for data_row in range(first_data_row, first_data_row + student_count):
                for col in mark_cols:
                    sheet.cell(row=data_row, column=col).value = 1
        workbook.save(workbook_path)
    finally:
        workbook.close()


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
        header_row, _ = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"A{header_row}"] = "Tampered"
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="header mismatch"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_empty_mark_entry_cell(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        _, first_data_row = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"D{first_data_row}"] = None
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="empty mark-entry cell"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_student_identity_tampering(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        _, first_data_row = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"B{first_data_row}"] = "R999"
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="student Reg. No./Name rows were modified"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_mark_value_above_maximum(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        _, first_data_row = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"D{first_data_row}"] = 999
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="invalid mark value"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_non_numeric_non_a_mark_value(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        _, first_data_row = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"D{first_data_row}"] = "ABSENT"
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="invalid mark value"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_mixed_absence_and_numeric_in_row(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        _, first_data_row = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"D{first_data_row}"] = "A"
        workbook["S1"][f"E{first_data_row}"] = 1
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="mixed absence and numeric"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_excess_decimal_precision(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        _, first_data_row = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"D{first_data_row}"] = 1.234
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="too many decimal places"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_indirect_non_integer_value(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        _, first_data_row = _sheet_rows_from_manifest(workbook, "CSURVEY")
        workbook["CSURVEY"][f"D{first_data_row}"] = 2.5
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="integer Likert value"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_total_formula_tampering_on_later_rows(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        header_row, first_data_row = _sheet_rows_from_manifest(workbook, "S1")
        total_col = len(
            [
                workbook["S1"].cell(row=header_row, column=idx).value
                for idx in range(1, workbook["S1"].max_column + 1)
                if workbook["S1"].cell(row=header_row, column=idx).value is not None
            ]
        )
        workbook["S1"].cell(row=first_data_row + 1, column=total_col).value = "=1+1"
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="total formula was modified"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)


def test_step3_validation_rejects_structure_snapshot_tampering(tmp_path: Path) -> None:
    workbook_path = _build_filled_marks_template(tmp_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        header_row, _ = _sheet_rows_from_manifest(workbook, "S1")
        workbook["S1"][f"D{header_row + 2}"] = 999
        workbook.save(workbook_path)
    finally:
        workbook.close()

    with pytest.raises(ValidationError, match="invalid mark value|mark-structure cell"):
        instructor_ui._validate_uploaded_filled_marks_workbook(workbook_path)
