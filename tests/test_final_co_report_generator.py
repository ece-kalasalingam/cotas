from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import DIRECT_RATIO, INDIRECT_RATIO
from common.exceptions import ValidationError
from common.workbook_signing import sign_payload
from domain.instructor_engine import (
    generate_course_details_template,
    generate_final_co_report,
    generate_marks_template_from_course_details,
)


def _build_filled_marks_workbook(tmp_path: Path) -> Path:
    course_details = tmp_path / "course_details.xlsx"
    marks = tmp_path / "marks_template.xlsx"
    generate_course_details_template(course_details)
    generate_marks_template_from_course_details(course_details, marks)

    wb = openpyxl.load_workbook(marks)
    try:
        manifest_text = wb["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest_text, str)
        manifest = json.loads(manifest_text)
        for spec in manifest.get("sheets", []):
            kind = spec.get("kind")
            if kind not in {"direct_co_wise", "direct_non_co_wise", "indirect"}:
                continue
            ws = wb[spec["name"]]
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
                    ws.cell(row=row, column=col).value = 1
                row += 1
        wb.save(marks)
    finally:
        wb.close()
    return marks


def _find_direct_header_row(ws: object) -> int:
    row = 1
    while row <= 200:
        if ws.cell(row=row, column=1).value == "#":
            return row
        row += 1
    raise AssertionError("Could not find direct header row")


def _ratio_header(ratio: float) -> str:
    percent = ratio * 100.0
    if abs(percent - round(percent)) <= 1e-9:
        token = f"{int(round(percent))}"
    else:
        token = f"{percent:g}"
    return f"Total ({token}%)"


def test_generate_final_co_report_creates_direct_sheet_per_outcome(tmp_path: Path) -> None:
    marks = _build_filled_marks_workbook(tmp_path)
    output = tmp_path / "co_report.xlsx"

    generated = generate_final_co_report(marks, output)

    assert generated == output
    assert output.exists()
    wb = openpyxl.load_workbook(output)
    try:
        for co in range(1, 7):
            assert f"CO{co}_Direct" in wb.sheetnames
            assert f"CO{co}_Indirect" in wb.sheetnames
        ws = wb["CO1_Direct"]
        assert ws["B1"].value == "Field"
        assert ws["C1"].value == "Value"
        metadata_fields = [str(ws.cell(row=row, column=2).value or "") for row in range(2, 20)]
        metadata_values = [str(ws.cell(row=row, column=3).value or "") for row in range(2, 20)]
        assert "Faculty_Name" not in metadata_fields
        assert "Total_Outcomes" not in metadata_fields
        assert "Course Outcome" in metadata_fields
        assert "CO1 - Direct" in metadata_values
        header_row = _find_direct_header_row(ws)
        assert ws.cell(row=header_row, column=1).value == "#"
        assert ws.cell(row=header_row, column=2).value == "Reg. No."
        assert ws.cell(row=header_row, column=3).value == "Student Name"
        metadata_max_b = max(
            len(str(ws.cell(row=row, column=2).value or "").strip())
            for row in range(1, header_row)
        )
        metadata_max_c = max(
            len(str(ws.cell(row=row, column=3).value or "").strip())
            for row in range(1, header_row)
        )
        assert ws.column_dimensions["B"].width is not None
        assert ws.column_dimensions["C"].width is not None
        assert float(ws.column_dimensions["B"].width) >= metadata_max_b
        assert float(ws.column_dimensions["C"].width) >= metadata_max_c
        headers = [
            str(ws.cell(row=header_row, column=col).value or "") for col in range(1, ws.max_column + 1)
        ]
        assert any(text.startswith("S1 (") for text in headers)
        assert any(text.startswith("ESP (") for text in headers)
        assert any(text.startswith("ESE (") for text in headers)
        assert not any(text.startswith("S2 (") for text in headers)
        assert not any(text.startswith("MSP (") for text in headers)
        assert not any(text.startswith("RLP (") for text in headers)
        assert headers[-3].startswith("Total (")
        assert headers[-2] == "Total (100%)"
        assert headers[-1] == _ratio_header(DIRECT_RATIO)
        assert ws.auto_filter.ref is None
        assert ws.freeze_panes == f"D{header_row + 1}"
        first_data_row = header_row + 1
        assert ws.cell(row=first_data_row, column=1).value == 1
        assert ws.cell(row=first_data_row, column=2).value == "R101"
        assert ws.cell(row=2, column=3).alignment.wrap_text is True
        assert ws.cell(row=2, column=3).alignment.horizontal == "left"
        assert ws.cell(row=first_data_row, column=3).alignment.wrap_text is True
        assert ws.cell(row=first_data_row, column=3).alignment.horizontal == "left"
        assert ws.cell(row=first_data_row, column=4).alignment.horizontal == "center"
        assert ws.cell(row=header_row + 100, column=3).border.left.style is None
        assert str(ws.cell(row=header_row, column=1).fill.fgColor.rgb or "").upper().endswith("D9EAD3")
        assert ws["A1"].value is None
        assert ws["A1"].border.left.style is None
        assert ws["A2"].border.left.style is None
        assert ws.protection.sheet is True
        assert str(ws.page_setup.paperSize) == str(ws.PAPERSIZE_A3)
        assert ws.page_setup.orientation == ws.ORIENTATION_LANDSCAPE
        assert ws.page_setup.fitToWidth == 1
        assert ws.page_setup.fitToHeight == 0

        indirect = wb["CO1_Indirect"]
        indirect_header_row = _find_direct_header_row(indirect)
        indirect_headers = [
            str(indirect.cell(row=indirect_header_row, column=col).value or "")
            for col in range(1, indirect.max_column + 1)
        ]
        assert indirect.cell(row=2, column=3).alignment.wrap_text is True
        assert indirect.cell(row=2, column=3).alignment.horizontal == "left"
        assert indirect.cell(row=indirect_header_row + 1, column=3).alignment.wrap_text is True
        assert indirect.cell(row=indirect_header_row + 1, column=3).alignment.horizontal == "left"
        assert indirect.cell(row=indirect_header_row + 100, column=3).border.left.style is None
        assert any(text.startswith("CSURVEY (1-5)") for text in indirect_headers)
        assert any("scaled 0-4" in text for text in indirect_headers)
        assert not any("CSURVEY (100%)" in text for text in indirect_headers)
        assert indirect_headers[-2] == "Total (100%)"
        assert indirect_headers[-1] == _ratio_header(INDIRECT_RATIO)
        assert indirect.freeze_panes == f"D{indirect_header_row + 1}"
    finally:
        wb.close()


def test_generate_final_co_report_absent_student_marks_totals_na(tmp_path: Path) -> None:
    marks = _build_filled_marks_workbook(tmp_path)
    wb = openpyxl.load_workbook(marks)
    try:
        manifest_text = wb["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest_text, str)
        manifest = json.loads(manifest_text)
        s1_spec = next(spec for spec in manifest.get("sheets", []) if spec.get("name") == "S1")
        first_data_row = int(s1_spec["header_row"]) + 3
        wb["S1"][f"D{first_data_row}"] = "A"
        wb.save(marks)
    finally:
        wb.close()

    output = tmp_path / "co_report_absent.xlsx"
    generate_final_co_report(marks, output)

    wb_out = openpyxl.load_workbook(output)
    try:
        ws = wb_out["CO1_Direct"]
        header_row = _find_direct_header_row(ws)
        headers = [
            str(ws.cell(row=header_row, column=col).value or "") for col in range(1, ws.max_column + 1)
        ]
        first_data_row = header_row + 1
        s1_raw_col = next(idx for idx, text in enumerate(headers, start=1) if text.startswith("S1 ("))
        total_100_col = headers.index("Total (100%)") + 1
        total_80_col = headers.index(_ratio_header(DIRECT_RATIO)) + 1
        total_weight_col = total_100_col - 1
        assert ws.cell(row=first_data_row, column=s1_raw_col).value == "A"
        assert ws.cell(row=first_data_row, column=s1_raw_col + 1).value == "A"
        assert ws.cell(row=first_data_row, column=total_weight_col).value == "NA"
        assert ws.cell(row=first_data_row, column=total_100_col).value == "NA"
        assert ws.cell(row=first_data_row, column=total_80_col).value == "NA"
    finally:
        wb_out.close()


def test_generate_final_co_report_rejects_tampered_system_hash(tmp_path: Path) -> None:
    marks = _build_filled_marks_workbook(tmp_path)
    wb = openpyxl.load_workbook(marks)
    try:
        wb["__SYSTEM_HASH__"]["B2"] = "tampered-hash"
        wb.save(marks)
    finally:
        wb.close()

    output = tmp_path / "co_report_tampered_hash.xlsx"
    with pytest.raises(ValidationError, match="hash mismatch|hash"):
        generate_final_co_report(marks, output)


def test_generate_final_co_report_rejects_tampered_layout_hash(tmp_path: Path) -> None:
    marks = _build_filled_marks_workbook(tmp_path)
    wb = openpyxl.load_workbook(marks)
    try:
        wb["__SYSTEM_LAYOUT__"]["B2"] = "tampered-layout-hash"
        wb.save(marks)
    finally:
        wb.close()

    output = tmp_path / "co_report_tampered_layout.xlsx"
    with pytest.raises(ValidationError, match="Layout hash mismatch|layout hash"):
        generate_final_co_report(marks, output)


def test_generate_final_co_report_rejects_unsupported_template_id(tmp_path: Path) -> None:
    marks = _build_filled_marks_workbook(tmp_path)
    wb = openpyxl.load_workbook(marks)
    try:
        wb["__SYSTEM_HASH__"]["A2"] = "COURSE_SETUP_V2"
        wb["__SYSTEM_HASH__"]["B2"] = sign_payload("COURSE_SETUP_V2")
        wb.save(marks)
    finally:
        wb.close()

    output = tmp_path / "co_report_unsupported_template.xlsx"
    with pytest.raises(ValidationError, match="Unknown workbook template"):
        generate_final_co_report(marks, output)


def test_generate_final_co_report_writes_hidden_integrity_system_sheets(tmp_path: Path) -> None:
    marks = _build_filled_marks_workbook(tmp_path)
    output = tmp_path / "co_report_integrity.xlsx"
    generate_final_co_report(marks, output)

    wb = openpyxl.load_workbook(output)
    try:
        assert "__SYSTEM_HASH__" in wb.sheetnames
        assert "__REPORT_INTEGRITY__" in wb.sheetnames
        assert wb["__SYSTEM_HASH__"].sheet_state == "hidden"
        assert wb["__REPORT_INTEGRITY__"].sheet_state == "hidden"

        integrity_ws = wb["__REPORT_INTEGRITY__"]
        assert integrity_ws["A1"].value == "Report_Manifest"
        assert integrity_ws["B1"].value == "Report_Hash"
        manifest_text = integrity_ws["A2"].value
        manifest_hash = integrity_ws["B2"].value
        assert isinstance(manifest_text, str) and manifest_text
        assert isinstance(manifest_hash, str) and manifest_hash
        manifest = json.loads(manifest_text)
        assert manifest.get("schema_version") == 1
        assert isinstance(manifest.get("template_id"), str)
        assert isinstance(manifest.get("template_hash"), str)
        assert isinstance(manifest.get("sheet_order"), list)
        assert isinstance(manifest.get("sheets"), list)
    finally:
        wb.close()
