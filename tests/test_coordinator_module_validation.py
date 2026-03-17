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


def test_has_valid_final_co_report_rejects_unsupported_template_id(tmp_path: Path) -> None:
    report = _build_valid_final_report(tmp_path / "unsupported.xlsx", template_id="COURSE_SETUP_V2")
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
            ("R002", "Student Two", "NA"),
            ("R001", "Student One", 80),
        ],
        indirect_rows=[
            ("R002", "Student Two", 15),
            ("R001", "Student One", 20),
        ],
    )
    _set_co_scores(
        second,
        co_index=1,
        direct_rows=[
            ("R004", "Student Four", 70),
            ("R001", "Student One Duplicate", 75),
            ("R003", "Student Three", 60),
        ],
        indirect_rows=[
            ("R004", "Student Four", 15),
            ("R001", "Student One Duplicate", 19),
            ("R003", "Student Three", "NA"),
        ],
    )

    out = tmp_path / "co_attainment.xlsx"
    result = coordinator._generate_co_attainment_workbook(
        [first, second],
        out,
        token=CancellationToken(),
    )

    assert result.output_path == out
    assert result.duplicate_reg_count == 1
    assert result.duplicate_entries == (("R001", "CO1_Direct", "second.xlsx"),)
    wb = openpyxl.load_workbook(out, data_only=True)
    try:
        assert "CO1" in wb.sheetnames
        assert "Summary" in wb.sheetnames
        assert "Graph" in wb.sheetnames
        assert SYSTEM_HASH_SHEET in wb.sheetnames
        assert SYSTEM_REPORT_INTEGRITY_SHEET in wb.sheetnames
        assert wb[SYSTEM_HASH_SHEET].sheet_state == "hidden"
        assert wb[SYSTEM_REPORT_INTEGRITY_SHEET].sheet_state == "hidden"

        hash_sheet = wb[SYSTEM_HASH_SHEET]
        assert hash_sheet["A1"].value == SYSTEM_HASH_TEMPLATE_ID_HEADER
        assert hash_sheet["B1"].value == SYSTEM_HASH_TEMPLATE_HASH_HEADER
        template_id = str(hash_sheet["A2"].value or "")
        template_hash = str(hash_sheet["B2"].value or "")
        assert template_id
        assert template_hash == sign_payload(template_id)

        integrity_ws = wb[SYSTEM_REPORT_INTEGRITY_SHEET]
        assert integrity_ws["A1"].value == SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER
        assert integrity_ws["B1"].value == SYSTEM_REPORT_INTEGRITY_HASH_HEADER
        manifest_text = integrity_ws["A2"].value
        manifest_hash = integrity_ws["B2"].value
        assert isinstance(manifest_text, str) and manifest_text
        assert isinstance(manifest_hash, str) and manifest_hash
        assert sign_payload(manifest_text) == manifest_hash
        manifest = json.loads(manifest_text)
        assert manifest.get("schema_version") == 1
        assert manifest.get("template_id") == template_id
        assert manifest.get("template_hash") == template_hash
        assert manifest.get("sheet_order") == ["CO1", "Summary", "Graph", SYSTEM_HASH_SHEET]
        sheet_hashes = manifest.get("sheets")
        assert isinstance(sheet_hashes, list)
        assert [entry.get("name") for entry in sheet_hashes] == manifest.get("sheet_order")
        assert all(
            isinstance(entry, dict) and entry.get("hash") == sign_payload(str(entry.get("name", "")))
            for entry in sheet_hashes
        )

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
        assert ws["G7"].value == "Level"
        assert ws.column_dimensions["B"].width is not None
        assert ws.column_dimensions["C"].width is not None
        assert ws["C2"].alignment.wrap_text is True
        assert ws["C8"].alignment.wrap_text is True
        assert ws.protection.sheet is True
        assert ws.print_title_rows == "$1:$7"

        rows = []
        row_idx = 8
        while isinstance(ws.cell(row=row_idx, column=1).value, int):
            rows.append(
                (
                    ws.cell(row=row_idx, column=1).value,
                    ws.cell(row=row_idx, column=2).value,
                    ws.cell(row=row_idx, column=3).value,
                    ws.cell(row=row_idx, column=4).value,
                    ws.cell(row=row_idx, column=5).value,
                    ws.cell(row=row_idx, column=6).value,
                    ws.cell(row=row_idx, column=7).value,
                )
            )
            row_idx += 1

        assert rows == [
            (1, "R001", "Student One", 80, 20, 100, 3),
            (2, "R002", "Student Two", "A", 15, "A", "NA"),
            (3, "R003", "Student Three", 60, "A", "A", "NA"),
            (4, "R004", "Student Four", 70, 15, 85, 3),
        ]
        assert ws["B13"].value == "On Roll:"
        assert ws["C13"].value == 4
        assert ws["B14"].value == "Attended:"
        assert ws["C14"].value == 2
        assert ws["B15"].value == "Level 0:"
        assert ws["C15"].value == 0
        assert ws["B16"].value == "Level 1:"
        assert ws["C16"].value == 0
        assert ws["B17"].value == "Level 2:"
        assert ws["C17"].value == 0
        assert ws["B18"].value == "Level 3:"
        assert ws["C18"].value == 2

        summary = wb["Summary"]
        assert summary["B1"].value == "Course Code"
        assert summary["C1"].value == "ECE000"
        assert summary["B2"].value == "Course Name"
        assert summary["C2"].value == "Signals and Systems"
        assert summary["B3"].value == "Semester"
        assert summary["C3"].value == "III"
        assert summary["B4"].value == "Academic Year"
        assert summary["C4"].value == "2025-26"
        assert summary["B5"].value == "CO Number"
        assert summary["C5"].value == "All COs"
        assert summary["A7"].value == "CO"
        assert summary["B7"].value == "Level 0"
        assert summary["C7"].value == "Level 1"
        assert summary["D7"].value == "Level 2"
        assert summary["E7"].value == "Level 3"
        assert summary["F7"].value == "Attended"
        assert summary["G7"].value == "CO%"
        assert summary["A8"].value == "CO1"
        assert summary["B8"].value == 0
        assert summary["C8"].value == 0
        assert summary["D8"].value == 0
        assert summary["E8"].value == 2
        assert summary["F8"].value == 2
        assert summary["G8"].value == 100
        assert summary.print_title_rows == "$1:$7"
        graph = wb["Graph"]
        assert graph["B1"].value == "Course Code"
        assert graph["C1"].value == "ECE000"
        assert graph["B2"].value == "Course Name"
        assert graph["C2"].value == "Signals and Systems"
        assert graph["B3"].value == "Semester"
        assert graph["C3"].value == "III"
        assert graph["B4"].value == "Academic Year"
        assert graph["C4"].value == "2025-26"
        assert graph["B5"].value == "CO Number"
        assert graph["C5"].value == "All COs"
        assert graph.print_title_rows == "$1:$5"
        assert len(graph._charts) == 1
        assert graph._charts[0].series[0].dLbls is not None
    finally:
        wb.close()


def test_generate_co_attainment_workbook_level_boundaries(tmp_path: Path) -> None:
    report = _build_valid_final_report(tmp_path / "section_a.xlsx", section="A")
    _set_co_scores(
        report,
        co_index=1,
        direct_rows=[
            ("R000", "S0", 0),
            ("R001", "S1", 40),
            ("R002", "S2", 60),
            ("R003", "S3", 75),
            ("R004", "S4", 100),
            ("R005", "S5", -1),
            ("R006", "S6", 101),
            ("R007", "S7", "NA"),
        ],
        indirect_rows=[
            ("R000", "S0", 0),
            ("R001", "S1", 0),
            ("R002", "S2", 0),
            ("R003", "S3", 0),
            ("R004", "S4", 0),
            ("R005", "S5", 0),
            ("R006", "S6", 0),
            ("R007", "S7", 0),
        ],
    )

    out = tmp_path / "co_attainment_levels.xlsx"
    coordinator._generate_co_attainment_workbook([report], out, token=CancellationToken())

    wb = openpyxl.load_workbook(out, data_only=True)
    try:
        ws = wb["CO1"]
        levels: list[object] = []
        totals: list[object] = []
        row_idx = 8
        while isinstance(ws.cell(row=row_idx, column=1).value, int):
            totals.append(ws.cell(row=row_idx, column=6).value)
            levels.append(ws.cell(row=row_idx, column=7).value)
            row_idx += 1
        assert totals == [0, 40, 60, 75, 100, -1, 101, "A"]
        assert levels == [0, 1, 2, 3, 3, "NA", "NA", "NA"]
        assert ws["B17"].value == "On Roll:"
        assert ws["C17"].value == 8
        assert ws["B18"].value == "Attended:"
        assert ws["C18"].value == 7
        assert ws["B19"].value == "Level 0:"
        assert ws["C19"].value == 1
        assert ws["B20"].value == "Level 1:"
        assert ws["C20"].value == 1
        assert ws["B21"].value == "Level 2:"
        assert ws["C21"].value == 1
        assert ws["B22"].value == "Level 3:"
        assert ws["C22"].value == 2
    finally:
        wb.close()

def test_generate_co_attainment_workbook_rejects_unsupported_template_id(tmp_path: Path) -> None:
    report = _build_valid_final_report(tmp_path / "unsupported.xlsx", template_id="COURSE_SETUP_V2")
    out = tmp_path / "co_attainment.xlsx"
    with pytest.raises(
        ValueError,
        match=r"^Invalid final CO report file:",
    ):
        coordinator._generate_co_attainment_workbook([report], out, token=CancellationToken())

