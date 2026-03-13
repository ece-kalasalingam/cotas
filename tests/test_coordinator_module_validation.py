from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from common.constants import (
    CO_REPORT_HEADER_REG_NO,
    CO_REPORT_HEADER_SERIAL,
    CO_REPORT_HEADER_STUDENT_NAME,
    CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE,
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_SECTION_KEY,
    COURSE_METADATA_SHEET,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    DIRECT_RATIO,
    INDIRECT_RATIO,
    SYSTEM_HASH_SHEET,
    SYSTEM_HASH_TEMPLATE_HASH_HEADER,
    SYSTEM_HASH_TEMPLATE_ID_HEADER,
    SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
    SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
    SYSTEM_REPORT_INTEGRITY_SHEET,
)
from common.jobs import CancellationToken
from common.workbook_signing import sign_payload
from modules import coordinator_module as coordinator


def _build_valid_final_report(
    path: Path,
    *,
    template_id: str = "COURSE_SETUP_V1",
    course_code: str = "ECE000",
    total_outcomes: int = 1,
    section: str = "A",
    direct_sheet_count: int = 1,
    indirect_sheet_count: int = 1,
) -> Path:
    wb = openpyxl.Workbook()
    try:
        first = wb.active
        first.title = COURSE_METADATA_SHEET
        first["A1"] = "Field"
        first["B1"] = "Value"
        first["A2"] = COURSE_METADATA_COURSE_CODE_KEY
        first["B2"] = course_code
        first["A3"] = COURSE_METADATA_TOTAL_OUTCOMES_KEY
        first["B3"] = total_outcomes
        first["A4"] = COURSE_METADATA_SECTION_KEY
        first["B4"] = section
        first["A5"] = "course_name"
        first["B5"] = "Signals and Systems"
        first["A6"] = COURSE_METADATA_SEMESTER_KEY
        first["B6"] = "III"
        first["A7"] = COURSE_METADATA_ACADEMIC_YEAR_KEY
        first["B7"] = "2025-26"

        for idx in range(1, direct_sheet_count + 1):
            wb.create_sheet(f"CO{idx}_Direct")
        for idx in range(1, indirect_sheet_count + 1):
            wb.create_sheet(f"CO{idx}_Indirect")

        system_hash = wb.create_sheet(SYSTEM_HASH_SHEET)
        system_hash.sheet_state = "hidden"
        system_hash["A1"] = SYSTEM_HASH_TEMPLATE_ID_HEADER
        system_hash["B1"] = SYSTEM_HASH_TEMPLATE_HASH_HEADER
        template_hash = sign_payload(template_id)
        system_hash["A2"] = template_id
        system_hash["B2"] = template_hash

        sheet_order = [COURSE_METADATA_SHEET]
        sheet_order.extend(f"CO{idx}_Direct" for idx in range(1, direct_sheet_count + 1))
        sheet_order.extend(f"CO{idx}_Indirect" for idx in range(1, indirect_sheet_count + 1))
        sheet_order.append(SYSTEM_HASH_SHEET)

        manifest = {
            "schema_version": 1,
            "template_id": template_id,
            "template_hash": template_hash,
            "sheet_order": sheet_order,
            "sheets": [
                {"name": COURSE_METADATA_SHEET, "hash": "m1"},
                *(
                    {"name": f"CO{idx}_Direct", "hash": f"d{idx}"}
                    for idx in range(1, direct_sheet_count + 1)
                ),
                *(
                    {"name": f"CO{idx}_Indirect", "hash": f"i{idx}"}
                    for idx in range(1, indirect_sheet_count + 1)
                ),
                {"name": SYSTEM_HASH_SHEET, "hash": "sys"},
            ],
        }
        manifest_text = json.dumps(manifest, sort_keys=True, separators=(",", ":"), ensure_ascii=True)
        integrity = wb.create_sheet(SYSTEM_REPORT_INTEGRITY_SHEET)
        integrity.sheet_state = "hidden"
        integrity["A1"] = SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER
        integrity["B1"] = SYSTEM_REPORT_INTEGRITY_HASH_HEADER
        integrity["A2"] = manifest_text
        integrity["B2"] = sign_payload(manifest_text)

        wb.save(path)
        return path
    finally:
        wb.close()


def test_has_valid_final_co_report_accepts_signed_final_report(tmp_path: Path) -> None:
    report = _build_valid_final_report(tmp_path / "final_co_report.xlsx")
    assert coordinator._has_valid_final_co_report(report) is True


def test_has_valid_final_co_report_rejects_missing_integrity_sheet(tmp_path: Path) -> None:
    report = _build_valid_final_report(tmp_path / "final_co_report.xlsx")
    wb = openpyxl.load_workbook(report)
    try:
        wb.remove(wb[SYSTEM_REPORT_INTEGRITY_SHEET])
        wb.save(report)
    finally:
        wb.close()

    assert coordinator._has_valid_final_co_report(report) is False


def test_analyze_dropped_files_marks_non_final_reports_invalid(tmp_path: Path) -> None:
    good = _build_valid_final_report(tmp_path / "ok.xlsx")
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not-an-excel")
    non_excel = tmp_path / "note.txt"
    non_excel.write_text("x", encoding="utf-8")

    result = coordinator._analyze_dropped_files(
        [str(good), str(bad), str(non_excel)],
        existing_keys=set(),
        existing_paths=[],
        token=CancellationToken(),
    )

    assert result["added"] == [str(good.resolve())]
    assert result["duplicates"] == 0
    assert result["invalid_final_report"] == [str(bad.resolve())]
    assert result["ignored"] == 2


def test_analyze_dropped_files_rejects_mismatched_template_id_against_base(tmp_path: Path) -> None:
    base = _build_valid_final_report(tmp_path / "base.xlsx", template_id="COURSE_SETUP_V1", section="A")
    mismatch = _build_valid_final_report(
        tmp_path / "mismatch.xlsx",
        template_id="COURSE_SETUP_V2",
        section="B",
    )

    result = coordinator._analyze_dropped_files(
        [str(mismatch)],
        existing_keys={coordinator._path_key(base)},
        existing_paths=[str(base)],
        token=CancellationToken(),
    )

    assert result["added"] == []
    assert result["invalid_final_report"] == [str(mismatch.resolve())]


def test_analyze_dropped_files_rejects_same_section_across_files(tmp_path: Path) -> None:
    base = _build_valid_final_report(tmp_path / "base.xlsx", section="A")
    same_section = _build_valid_final_report(tmp_path / "same_section.xlsx", section="A")

    result = coordinator._analyze_dropped_files(
        [str(same_section)],
        existing_keys={coordinator._path_key(base)},
        existing_paths=[str(base)],
        token=CancellationToken(),
    )

    assert result["added"] == []
    assert result["invalid_final_report"] == [str(same_section.resolve())]


def test_analyze_dropped_files_accepts_unique_sections_with_same_course_signature(tmp_path: Path) -> None:
    first = _build_valid_final_report(tmp_path / "first.xlsx", section="A")
    second = _build_valid_final_report(tmp_path / "second.xlsx", section="B")

    result = coordinator._analyze_dropped_files(
        [str(first), str(second)],
        existing_keys=set(),
        existing_paths=[],
        token=CancellationToken(),
    )

    assert result["added"] == [str(first.resolve()), str(second.resolve())]
    assert result["invalid_final_report"] == []


def test_has_valid_final_co_report_rejects_unbalanced_direct_indirect_sheet_counts(tmp_path: Path) -> None:
    report = _build_valid_final_report(tmp_path / "bad_counts.xlsx", direct_sheet_count=2, indirect_sheet_count=1)
    assert coordinator._has_valid_final_co_report(report) is False


@pytest.mark.parametrize(
    ("input_name", "section", "expected"),
    [
        ("ECE000_III_A_2025-26_COReport.xlsx", "A", "ECE000_III_2025-26_CO_Attainment.xlsx"),
        ("ECE000_III_A_2025-26_CO_Report.xlsx", "A", "ECE000_III_2025-26_CO_Attainment.xlsx"),
        ("ECE000_III_A_2025-26_CO Report.xlsx", "A", "ECE000_III_2025-26_CO_Attainment.xlsx"),
        ("ECE000_III_A_2025-26.xlsx", "A", "ECE000_III_2025-26_CO_Attainment.xlsx"),
        ("ECE000_III_B_2025-26.xlsx", "A", "ECE000_III_B_2025-26_CO_Attainment.xlsx"),
    ],
)
def test_build_co_attainment_default_name_strips_co_report_token(
    input_name: str,
    section: str,
    expected: str,
) -> None:
    assert coordinator._build_co_attainment_default_name(Path(input_name), section=section) == expected


def _ratio_header(ratio: float) -> str:
    percent = ratio * 100.0
    token = f"{int(round(percent))}" if abs(percent - round(percent)) <= 1e-9 else f"{percent:g}"
    return CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE.format(ratio=token)


def _set_co_scores(
    path: Path,
    *,
    co_index: int,
    direct_rows: list[tuple[str, str, object]],
    indirect_rows: list[tuple[str, str, object]],
) -> None:
    wb = openpyxl.load_workbook(path)
    try:
        direct = wb[f"CO{co_index}_Direct"]
        direct["A1"] = CO_REPORT_HEADER_SERIAL
        direct["B1"] = CO_REPORT_HEADER_REG_NO
        direct["C1"] = CO_REPORT_HEADER_STUDENT_NAME
        direct["D1"] = _ratio_header(DIRECT_RATIO)
        for idx, (reg_no, student_name, score) in enumerate(direct_rows, start=2):
            direct.cell(row=idx, column=1, value=idx - 1)
            direct.cell(row=idx, column=2, value=reg_no)
            direct.cell(row=idx, column=3, value=student_name)
            direct.cell(row=idx, column=4, value=score)

        indirect = wb[f"CO{co_index}_Indirect"]
        indirect["A1"] = CO_REPORT_HEADER_SERIAL
        indirect["B1"] = CO_REPORT_HEADER_REG_NO
        indirect["C1"] = CO_REPORT_HEADER_STUDENT_NAME
        indirect["D1"] = _ratio_header(INDIRECT_RATIO)
        for idx, (reg_no, student_name, score) in enumerate(indirect_rows, start=2):
            indirect.cell(row=idx, column=1, value=idx - 1)
            indirect.cell(row=idx, column=2, value=reg_no)
            indirect.cell(row=idx, column=3, value=student_name)
            indirect.cell(row=idx, column=4, value=score)
        wb.save(path)
    finally:
        wb.close()


def test_generate_co_attainment_workbook_filters_na_and_keeps_unique_registers(tmp_path: Path) -> None:
    first = _build_valid_final_report(tmp_path / "first.xlsx", section="A")
    second = _build_valid_final_report(tmp_path / "second.xlsx", section="B")

    _set_co_scores(
        first,
        co_index=1,
        direct_rows=[
            ("R001", "Student One", 80),
            ("R002", "Student Two", "NA"),
        ],
        indirect_rows=[
            ("R001", "Student One", 20),
            ("R002", "Student Two", 15),
        ],
    )
    _set_co_scores(
        second,
        co_index=1,
        direct_rows=[
            ("R001", "Student One Duplicate", 75),
            ("R003", "Student Three", 60),
            ("R004", "Student Four", 70),
        ],
        indirect_rows=[
            ("R001", "Student One Duplicate", 19),
            ("R003", "Student Three", "NA"),
            ("R004", "Student Four", 15),
        ],
    )

    out = tmp_path / "co_attainment.xlsx"
    result = coordinator._generate_co_attainment_workbook(
        [first, second],
        out,
        token=CancellationToken(),
    )

    assert result == out
    wb = openpyxl.load_workbook(out, data_only=True)
    try:
        assert "CO1" in wb.sheetnames
        ws = wb["CO1"]
        assert ws["B1"].value == "Course Code"
        assert ws["C1"].value == "ECE000"
        assert ws["B2"].value == "Course Name"
        assert ws["C2"].value == "Signals and Systems"
        assert ws["B3"].value == "Semester"
        assert ws["C3"].value == "III"
        assert ws["B4"].value == "Academic Year"
        assert ws["C4"].value == "2025-26"
        assert ws["B5"].value == "CO Number"
        assert ws["C5"].value == "CO1"

        assert ws["A7"].value == "#"
        assert ws["B7"].value == "Regno"
        assert ws["C7"].value == "Student name"
        assert ws["D7"].value == "Direct (80%)"
        assert ws["E7"].value == "Indirect (20%)"
        assert ws["F7"].value == "Total (100%)"
        assert ws.protection.sheet is True

        rows = []
        row_idx = 8
        while ws.cell(row=row_idx, column=2).value is not None:
            rows.append(
                (
                    ws.cell(row=row_idx, column=1).value,
                    ws.cell(row=row_idx, column=2).value,
                    ws.cell(row=row_idx, column=3).value,
                    ws.cell(row=row_idx, column=4).value,
                    ws.cell(row=row_idx, column=5).value,
                    ws.cell(row=row_idx, column=6).value,
                )
            )
            row_idx += 1

        assert rows == [
            (1, "R001", "Student One", 80, 20, 100),
            (2, "R002", "Student Two", "A", "A", "A"),
            (3, "R003", "Student Three", "A", "A", "A"),
            (4, "R004", "Student Four", 70, 15, 85),
        ]
    finally:
        wb.close()
