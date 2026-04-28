from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("xlsxwriter")
openpyxl = pytest.importorskip("openpyxl")

from common.constants import ID_COURSE_SETUP
from domain.template_strategy_router import generate_workbook
from domain.template_versions.course_setup_v2_impl.co_description_template_validator import (
    validate_co_description_workbooks,
)


def _generate_co_description_template(output_path: Path) -> Path:
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


def _fill_valid_co_rows(path: Path, *, total_outcomes: int = 6) -> None:
    workbook = openpyxl.load_workbook(path)
    try:
        sheet = workbook["CO_Description"]
        for offset in range(total_outcomes):
            row = 2 + offset
            co_index = offset + 1
            sheet.cell(row=row, column=1, value=co_index)
            sheet.cell(row=row, column=2, value=f"CO{co_index} description")
            sheet.cell(row=row, column=3, value="K2")
            sheet.cell(
                row=row,
                column=4,
                value=f"CO{co_index} summary " + ("x" * 120),
            )
        workbook.save(path)
    finally:
        workbook.close()


def _first_issue_code(result: dict[str, object]) -> str:
    rejections = result.get("rejections", [])
    if not isinstance(rejections, list) or not rejections:
        return ""
    first = rejections[0] if isinstance(rejections[0], dict) else {}
    issue = first.get("issue", {})
    if not isinstance(issue, dict):
        return ""
    return str(issue.get("code", "")).strip()


def test_validate_co_description_workbooks_accepts_valid_workbook(tmp_path: Path) -> None:
    workbook_path = _generate_co_description_template(tmp_path / "co_description_valid.xlsx")
    _fill_valid_co_rows(workbook_path)

    result = validate_co_description_workbooks(
        workbook_paths=[workbook_path],
        template_id=ID_COURSE_SETUP,
    )

    if not (result["valid_paths"] == [str(workbook_path)]):
        raise AssertionError('assertion failed')
    if not (result["rejections"] == []):
        raise AssertionError('assertion failed')


def test_validate_co_description_workbooks_rejects_missing_summary(tmp_path: Path) -> None:
    workbook_path = _generate_co_description_template(tmp_path / "co_description_missing_summary.xlsx")
    _fill_valid_co_rows(workbook_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        workbook["CO_Description"].cell(row=2, column=4, value="")
        workbook.save(workbook_path)
    finally:
        workbook.close()

    result = validate_co_description_workbooks(
        workbook_paths=[workbook_path],
        template_id=ID_COURSE_SETUP,
    )

    if not (result["valid_paths"] == []):
        raise AssertionError('assertion failed')
    if _first_issue_code(result) not in {"CELL_EMPTY_NOT_ALLOWED", "CO_DESCRIPTION_SUMMARY_REQUIRED"}:
        raise AssertionError('assertion failed')


def test_validate_co_description_workbooks_rejects_gapped_co_numbers(tmp_path: Path) -> None:
    workbook_path = _generate_co_description_template(tmp_path / "co_description_gapped.xlsx")
    _fill_valid_co_rows(workbook_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        workbook["CO_Description"].cell(row=7, column=1, value=7)
        workbook.save(workbook_path)
    finally:
        workbook.close()

    result = validate_co_description_workbooks(
        workbook_paths=[workbook_path],
        template_id=ID_COURSE_SETUP,
    )

    if not (result["valid_paths"] == []):
        raise AssertionError('assertion failed')
    if not (_first_issue_code(result) == "CO_DESCRIPTION_CO_NUMBER_SET_MISMATCH"):
        raise AssertionError('assertion failed')


def test_validate_co_description_workbooks_rejects_tampered_system_hash(tmp_path: Path) -> None:
    workbook_path = _generate_co_description_template(tmp_path / "co_description_tampered.xlsx")
    _fill_valid_co_rows(workbook_path)
    workbook = openpyxl.load_workbook(workbook_path)
    try:
        workbook["__SYSTEM_HASH__"].cell(row=2, column=1, value="COURSE_SETUP_V3")
        workbook.save(workbook_path)
    finally:
        workbook.close()

    result = validate_co_description_workbooks(
        workbook_paths=[workbook_path],
        template_id=ID_COURSE_SETUP,
    )

    if not (result["valid_paths"] == []):
        raise AssertionError('assertion failed')
    if _first_issue_code(result) not in {"COA_SYSTEM_HASH_MISMATCH", "UNKNOWN_TEMPLATE"}:
        raise AssertionError('assertion failed')
