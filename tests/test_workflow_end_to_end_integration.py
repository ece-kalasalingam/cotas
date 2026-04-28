from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import (
    CO_ANALYSIS_SHEET_FOOTER_TEXT,
    CO_REPORT_HEADER_REG_NO,
    CO_REPORT_HEADER_SERIAL,
    CO_REPORT_HEADER_STUDENT_NAME,
    ID_COURSE_SETUP,
)
from common.jobs import CancellationToken
from domain.template_strategy_router import generate_workbook, validate_workbooks
from domain.template_versions.course_setup_v2_impl.co_attainment import (
    extract_final_report_signature_from_path,
    generate_final_report_workbook,
)
from domain.template_versions.course_setup_v2_impl.co_report_sheet_generator import (
    co_direct_sheet_name,
    co_indirect_sheet_name,
    ratio_total_header,
)
from services.instructor_workflow_service import InstructorWorkflowService


def _set_course_section(course_details_path: Path, section: str) -> None:
    """Set course section.
    
    Args:
        course_details_path: Parameter value (Path).
        section: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(course_details_path)
    try:
        ws = wb["Course_Metadata"]
        row = 2
        while True:
            key = str(ws.cell(row=row, column=1).value or "").strip().casefold()
            value = ws.cell(row=row, column=2).value
            if not key and (value is None or str(value).strip() == ""):
                break
            if key == "section":
                ws.cell(row=row, column=2, value=section)
                break
            row += 1
        wb.save(course_details_path)
    finally:
        wb.close()


def _prefix_student_regnos(course_details_path: Path, prefix: str) -> None:
    """Prefix student regnos.
    
    Args:
        course_details_path: Parameter value (Path).
        prefix: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(course_details_path)
    try:
        ws = wb["Students"]
        row = 2
        while True:
            reg = ws.cell(row=row, column=1).value
            name = ws.cell(row=row, column=2).value
            if reg is None and name is None:
                break
            reg_text = str(reg).strip() if reg is not None else ""
            if reg_text:
                ws.cell(row=row, column=1, value=f"{prefix}{reg_text}")
            row += 1
        wb.save(course_details_path)
    finally:
        wb.close()


def _fill_marks_workbook(marks_path: Path, mark_value: float = 1.0) -> None:
    """Fill marks workbook.
    
    Args:
        marks_path: Parameter value (Path).
        mark_value: Parameter value (float).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(marks_path)
    try:
        manifest_text = wb["__SYSTEM_LAYOUT__"]["A2"].value
        if not (isinstance(manifest_text, str)):
            raise AssertionError('assertion failed')
        manifest = json.loads(manifest_text)
        for spec in manifest.get("sheets", []):
            kind = spec.get("kind")
            if kind not in {"direct_co_wise", "direct_non_co_wise", "indirect"}:
                continue
            ws = wb[str(spec["name"])]
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
                    ws.cell(row=row, column=col, value=mark_value)
                row += 1
        wb.save(marks_path)
    finally:
        wb.close()


def _build_final_report(
    root: Path,
    *,
    section: str,
    reg_prefix: str = "",
) -> Path:
    """Build final report.
    
    Args:
        root: Parameter value (Path).
        section: Parameter value (str).
        reg_prefix: Parameter value (str).
    
    Returns:
        Path: Return value.
    
    Raises:
        None.
    """
    instructor = InstructorWorkflowService()
    course_details = root / f"course_details_{section}.xlsx"
    marks_template = root / f"marks_template_{section}.xlsx"
    final_report = root / f"final_report_{section}.xlsx"

    generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=course_details,
        workbook_name=course_details.name,
        workbook_kind="course_details_template",
        cancel_token=CancellationToken(),
    )
    _set_course_section(course_details, section)
    if reg_prefix:
        _prefix_student_regnos(course_details, reg_prefix)

    context_prepare = instructor.create_job_context(step_id=f"prepare_marks_{section}")
    validate_workbooks(
        template_id=ID_COURSE_SETUP,
        workbook_paths=[course_details],
        workbook_kind="course_details",
        cancel_token=CancellationToken(),
    )
    instructor.generate_marks_template(
        course_details,
        marks_template,
        context=context_prepare,
        cancel_token=CancellationToken(),
    )

    _fill_marks_workbook(marks_template, mark_value=1.0)
    generate_final_report_workbook(
        filled_marks_path=marks_template,
        output_path=final_report,
        cancel_token=CancellationToken(),
    )
    return final_report


def test_instructor_workflow_service_end_to_end_generates_signed_final_report(tmp_path: Path) -> None:
    """Test instructor workflow service end to end generates signed final report.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    final_report = _build_final_report(tmp_path, section="A")
    if not (final_report.exists()):
        raise AssertionError('assertion failed')
    if not (extract_final_report_signature_from_path(final_report) is not None):
        raise AssertionError('assertion failed')


def _find_header_row(sheet, expected_headers: list[str]) -> int:  # noqa: ANN001
    """Find header row.
    
    Args:
        sheet: Parameter value.
        expected_headers: Parameter value (list[str]).
    
    Returns:
        int: Return value.
    
    Raises:
        AssertionError: if row is not found.
    """
    max_scan_rows = min(40, int(sheet.max_row or 0))
    for row_index in range(1, max_scan_rows + 1):
        values = [str(sheet.cell(row=row_index, column=col).value or "").strip() for col in range(1, len(expected_headers) + 1)]
        if values == expected_headers:
            return row_index
    raise AssertionError(f"Header row not found in sheet={sheet.title}, expected={expected_headers}")


def _collect_sheet_rows(sheet, *, header_row: int, score_col: int) -> list[tuple[str, str, object]]:  # noqa: ANN401
    """Collect sheet rows.
    
    Args:
        sheet: Parameter value.
        header_row: Parameter value (int).
        score_col: Parameter value (int).
    
    Returns:
        list[tuple[str, str, object]]: Return value.
    
    Raises:
        None.
    """
    rows: list[tuple[str, str, object]] = []
    reg_col: int | None = None
    name_col: int | None = None
    max_col = int(sheet.max_column or 0)
    for col in range(1, max_col + 1):
        token = str(sheet.cell(row=header_row, column=col).value or "").strip()
        if token == CO_REPORT_HEADER_REG_NO:
            reg_col = col
        elif token == CO_REPORT_HEADER_STUDENT_NAME:
            name_col = col
    if reg_col is None or name_col is None:
        raise AssertionError(f"Missing identity headers in sheet={sheet.title} row={header_row}")
    row_index = header_row + 1
    while row_index <= int(sheet.max_row or 0):
        reg_no = sheet.cell(row=row_index, column=reg_col).value
        student_name = sheet.cell(row=row_index, column=name_col).value
        if reg_no is None and student_name is None:
            break
        reg_text = str(reg_no).strip() if reg_no is not None else ""
        name_text = str(student_name).strip() if student_name is not None else ""
        score_value = sheet.cell(row=row_index, column=score_col).value
        rows.append((reg_text, name_text, score_value))
        row_index += 1
    return rows


def _header_values(sheet, *, header_row: int) -> list[str]:  # noqa: ANN001
    """Header values.

    Args:
        sheet: Parameter value.
        header_row: Parameter value (int).

    Returns:
        list[str]: Return value.

    Raises:
        None.
    """
    values: list[str] = []
    max_col = int(sheet.max_column or 0)
    for col in range(1, max_col + 1):
        token = str(sheet.cell(row=header_row, column=col).value or "").strip()
        if token:
            values.append(token)
    return values


def _column_for_header(sheet, *, header_row: int, header_value: str) -> int:  # noqa: ANN001
    """Column for header value in a specific row.

    Args:
        sheet: Parameter value.
        header_row: Parameter value (int).
        header_value: Parameter value (str).

    Returns:
        int: Return value.

    Raises:
        AssertionError: if header is not found.
    """
    max_col = int(sheet.max_column or 0)
    wanted = str(header_value).strip()
    for col in range(1, max_col + 1):
        token = str(sheet.cell(row=header_row, column=col).value or "").strip()
        if token == wanted:
            return col
    raise AssertionError(f"Header not found in sheet={sheet.title}, row={header_row}, header={header_value!r}")


def _normalized_score(value: object) -> float | str:
    """Normalized score.
    
    Args:
        value: Parameter value (object).
    
    Returns:
        float | str: Return value.
    
    Raises:
        None.
    """
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return str(value or "").strip()


def test_co_analysis_generation_emits_direct_indirect_and_co_triplets(tmp_path: Path) -> None:
    """Test co analysis generation emits direct indirect and co triplets.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    final_report = _build_final_report(tmp_path, section="A")
    co_analysis_path = tmp_path / "co_attainment.xlsx"
    generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=co_analysis_path,
        workbook_name=co_analysis_path.name,
        workbook_kind="co_attainment",
        cancel_token=CancellationToken(),
        context={
            "source_paths": [str(final_report)],
            "thresholds": (40.0, 60.0, 75.0),
            "co_attainment_percent": 60.0,
            "co_attainment_level": 2,
        },
    )

    workbook = openpyxl.load_workbook(co_analysis_path, data_only=False)
    source_workbook = openpyxl.load_workbook(final_report, data_only=False)
    try:
        co_titles = [name for name in workbook.sheetnames if re.fullmatch(r"CO\d+", str(name))]
        total_outcomes = len(co_titles)
        if not (total_outcomes > 0):
            raise AssertionError('assertion failed')

        expected_sheet_order: list[str] = ["Pass_Percentage", "Summary", "Graph"]
        for co_index in range(1, total_outcomes + 1):
            expected_sheet_order.extend(
                [
                    co_direct_sheet_name(co_index),
                    co_indirect_sheet_name(co_index),
                    f"CO{co_index}",
                ]
            )
        expected_sheet_order.extend(["__SYSTEM_HASH__", "__SYSTEM_LAYOUT__"])
        if not (list(workbook.sheetnames) == expected_sheet_order):
            raise AssertionError('assertion failed')
        for sheet_name in workbook.sheetnames:
            if sheet_name in {"__SYSTEM_HASH__", "__SYSTEM_LAYOUT__"}:
                continue
            sheet = workbook[sheet_name]
            values = [
                str(sheet.cell(row=row, column=1).value or "").strip()
                for row in range(1, int(sheet.max_row or 0) + 1)
            ]
            if CO_ANALYSIS_SHEET_FOOTER_TEXT not in values:
                raise AssertionError(f"Footer missing in sheet={sheet_name}")

        for co_index in range(1, total_outcomes + 1):
            direct_sheet = workbook[co_direct_sheet_name(co_index)]
            indirect_sheet = workbook[co_indirect_sheet_name(co_index)]
            co_sheet = workbook[f"CO{co_index}"]
            source_direct_sheet = source_workbook[co_direct_sheet_name(co_index)]
            source_indirect_sheet = source_workbook[co_indirect_sheet_name(co_index)]

            source_direct_header = _find_header_row(
                source_direct_sheet,
                [
                    CO_REPORT_HEADER_SERIAL,
                    CO_REPORT_HEADER_STUDENT_NAME,
                    CO_REPORT_HEADER_REG_NO,
                ],
            )
            source_direct_headers = _header_values(source_direct_sheet, header_row=source_direct_header)
            direct_header = _find_header_row(direct_sheet, source_direct_headers)
            source_indirect_header = _find_header_row(
                source_indirect_sheet,
                [
                    CO_REPORT_HEADER_SERIAL,
                    CO_REPORT_HEADER_STUDENT_NAME,
                    CO_REPORT_HEADER_REG_NO,
                ],
            )
            source_indirect_headers = _header_values(source_indirect_sheet, header_row=source_indirect_header)
            indirect_header = _find_header_row(indirect_sheet, source_indirect_headers)
            co_header = _find_header_row(
                co_sheet,
                [
                    "#",
                    "Student Name",
                    "Reg. No.",
                    "Direct (80%)",
                    "Indirect (20%)",
                ],
            )

            direct_rows = _collect_sheet_rows(
                direct_sheet,
                header_row=direct_header,
                score_col=_column_for_header(
                    direct_sheet,
                    header_row=direct_header,
                    header_value=ratio_total_header(0.8),
                ),
            )
            indirect_rows = _collect_sheet_rows(
                indirect_sheet,
                header_row=indirect_header,
                score_col=_column_for_header(
                    indirect_sheet,
                    header_row=indirect_header,
                    header_value=ratio_total_header(0.2),
                ),
            )
            co_direct_rows = _collect_sheet_rows(
                co_sheet,
                header_row=co_header,
                score_col=_column_for_header(co_sheet, header_row=co_header, header_value="Direct (80%)"),
            )
            co_indirect_rows = _collect_sheet_rows(
                co_sheet,
                header_row=co_header,
                score_col=_column_for_header(co_sheet, header_row=co_header, header_value="Indirect (20%)"),
            )

            if not ([row[:2] for row in direct_rows] == [row[:2] for row in co_direct_rows]):
                raise AssertionError('assertion failed')
            if not ([row[:2] for row in indirect_rows] == [row[:2] for row in co_indirect_rows]):
                raise AssertionError('assertion failed')
            if not ([_normalized_score(row[2]) for row in direct_rows] == [
                _normalized_score(row[2]) for row in co_direct_rows
            ]):
                raise AssertionError('assertion failed')
            if not ([_normalized_score(row[2]) for row in indirect_rows] == [
                _normalized_score(row[2]) for row in co_indirect_rows
            ]):
                raise AssertionError('assertion failed')
    finally:
        source_workbook.close()
        workbook.close()
