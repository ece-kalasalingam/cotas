from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from common.constants import (
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SECTION_KEY,
    COURSE_METADATA_SHEET,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
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
