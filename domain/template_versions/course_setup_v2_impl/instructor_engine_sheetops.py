"""Sheet/write helpers extracted from instructor_template_engine."""

from __future__ import annotations

import re
from typing import Any, Sequence

from common.constants import (
    CO_LABEL,
    COMPONENT_NAME_LABEL,
    COURSE_METADATA_FACULTY_NAME_KEY,
    INSTRUCTOR_MAX_LABEL,
    LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
    LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
    LAYOUT_SHEET_KIND_INDIRECT,
    LAYOUT_SHEET_SPEC_KEY_ANCHORS,
    LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS,
    LAYOUT_SHEET_SPEC_KEY_HEADER_ROW,
    LAYOUT_SHEET_SPEC_KEY_HEADERS,
    LAYOUT_SHEET_SPEC_KEY_KIND,
    LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE,
    LAYOUT_SHEET_SPEC_KEY_NAME,
    LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT,
    LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH,
    LIKERT_MAX,
    LIKERT_MIN,
    MARKS_ENTRY_INDIRECT_VALIDATION_ERROR_RANGE_TEMPLATE,
    MARKS_ENTRY_ROW_HEADERS,
    MARKS_ENTRY_VALIDATION_ERROR_RANGE_TEMPLATE,
    MARKS_ENTRY_VALIDATION_ERROR_TITLE,
    MAX_EXCEL_SHEETNAME_LENGTH,
    MIN_MARK_VALUE,
)
from common.registry import (
    COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE,
    COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
    COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT,
    resolve_dynamic_sheet_headers,
)
from common.excel_sheet_layout import (
    apply_xlsxwriter_column_widths,
    apply_xlsxwriter_sheet_frame,
    apply_xlsxwriter_viewport,
    compute_sampled_column_widths,
    excel_col_name as _excel_col_name_one_based,
    protect_xlsxwriter_sheet,
    XLSX_AUTOFIT_MAX_WIDTH,
    XLSX_AUTOFIT_MIN_WIDTH,
    XLSX_AUTOFIT_PADDING,
    XLSX_AUTOFIT_SAMPLE_ROWS,
    XLSX_PAGE_MIN_MARGIN_IN,
    XLSX_PAPER_SIZE_A4,
)
from common.error_catalog import validation_error_from_key
from common.sheet_schema import SheetSchema, ValidationRule
from common.utils import (
    coerce_excel_number,
    normalize,
)
from common.workbook_integrity.workbook_signing import sign_payload

_COMPONENT_NAME_LABEL = COMPONENT_NAME_LABEL
_CO_LABEL = CO_LABEL
_MAX_LABEL = INSTRUCTOR_MAX_LABEL
_TEMPLATE_ID = "COURSE_SETUP_V2"
_FORMULA_SUM_TEMPLATE = "=SUM({start}:{end})"
_VALIDATION_KIND_CUSTOM = "custom"
_VALIDATION_KEY_KIND = "validate"
_VALIDATION_KEY_VALUE = "value"
_VALIDATION_KEY_ERROR_TITLE = "error_title"
_VALIDATION_KEY_ERROR_MESSAGE = "error_message"
_VALIDATION_KEY_IGNORE_BLANK = "ignore_blank"

def _write_two_column_copy_sheet(
    workbook: Any,
    title: str,
    header: tuple[str, str],
    rows: Sequence[Sequence[Any]],
    header_fmt: Any,
    body_fmt: Any,
) -> None:
    ws = workbook.add_worksheet(title)
    ws.write_row(0, 0, list(header), header_fmt)
    sample_rows: list[list[Any]] = [list(header)]
    sample_rows.extend([[row[0] if len(row) > 0 else "", row[1] if len(row) > 1 else ""] for row in rows])
    widths = compute_sampled_column_widths(
        sample_rows,
        1,
        min_width=XLSX_AUTOFIT_MIN_WIDTH,
        max_width=XLSX_AUTOFIT_MAX_WIDTH,
        padding=XLSX_AUTOFIT_PADDING,
    )
    apply_xlsxwriter_column_widths(
        ws,
        widths,
        default_width=XLSX_AUTOFIT_MIN_WIDTH,
    )
    for row_index, row in enumerate(rows, start=1):
        ws.write(row_index, 0, row[0] if len(row) > 0 else "", body_fmt)
        ws.write(row_index, 1, row[1] if len(row) > 1 else "", body_fmt)
    apply_xlsxwriter_sheet_frame(
        ws,
        repeat_last_row=0,
        freeze_row=1,
        freeze_col=0,
        select_row=1,
        select_col=0,
        paper_size=XLSX_PAPER_SIZE_A4,
        landscape=False,
    )


def _build_two_column_copy_sheet_spec(
    *,
    title: str,
    header: tuple[str, str],
    rows: Sequence[Sequence[Any]],
) -> dict[str, Any]:
    anchors = []
    for row_index, row in enumerate(rows, start=2):
        anchors.append([f"A{row_index}", row[0] if len(row) > 0 else ""])
        anchors.append([f"B{row_index}", row[1] if len(row) > 1 else ""])
    return {
        LAYOUT_SHEET_SPEC_KEY_NAME: title,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: list(header),
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: anchors,
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }


def _write_multi_column_copy_sheet(
    workbook: Any,
    title: str,
    header: Sequence[str],
    rows: Sequence[Sequence[Any]],
    header_fmt: Any,
    body_fmt: Any,
    num_fmt: Any,
    *,
    metadata_rows: Sequence[Sequence[Any]] | None = None,
    wrapped_body_fmt: Any | None = None,
    wrapped_column_fmt: Any | None = None,
    wrap_columns: Sequence[int] = (),
    fit_all_columns_single_page: bool = False,
    use_common_student_columns: bool = False,
    header_row_height: float | None = None,
) -> None:
    ws = workbook.add_worksheet(title)
    metadata = list(metadata_rows or [])
    header_row_index = len(metadata) + 1 if metadata else 0

    for row_index, row in enumerate(metadata):
        ws.write(row_index, 1, row[0] if len(row) > 0 else "", body_fmt)
        metadata_value_fmt = wrapped_body_fmt or body_fmt
        ws.write(row_index, 2, row[1] if len(row) > 1 else "", metadata_value_fmt)

    ws.write_row(header_row_index, 0, list(header), header_fmt)
    if header_row_height is not None:
        ws.set_row(header_row_index, header_row_height)

    first_data_row = header_row_index + 1
    for row_index, row in enumerate(rows, start=first_data_row):
        for col_index, value in enumerate(row[: len(header)]):
            if col_index in wrap_columns:
                cell_fmt = wrapped_body_fmt or body_fmt
            elif col_index == 1 and isinstance(value, (int, float)):
                cell_fmt = num_fmt
            else:
                cell_fmt = body_fmt
            ws.write(row_index, col_index, value, cell_fmt)

    width_rows: list[list[Any]] = []
    for row in metadata:
        width_rows.append(["", row[0] if len(row) > 0 else "", row[1] if len(row) > 1 else ""])
    width_rows.append(list(header))
    sampled_data_rows = rows[: min(len(rows), XLSX_AUTOFIT_SAMPLE_ROWS)]
    for row in sampled_data_rows:
        width_rows.append([row[col] if col < len(row) else "" for col in range(len(header))])
    widths = compute_sampled_column_widths(
        width_rows,
        max(0, len(header) - 1),
        min_width=XLSX_AUTOFIT_MIN_WIDTH,
        max_width=XLSX_AUTOFIT_MAX_WIDTH,
        padding=XLSX_AUTOFIT_PADDING,
    )
    effective_wrap_columns = (2,) if use_common_student_columns else tuple(wrap_columns)
    apply_xlsxwriter_column_widths(
        ws,
        widths,
        default_width=XLSX_AUTOFIT_MIN_WIDTH,
        wrap_columns=effective_wrap_columns,
        wrap_format=wrapped_column_fmt or wrapped_body_fmt or body_fmt,
    )

    effective_landscape = True if use_common_student_columns else fit_all_columns_single_page
    effective_margins = (
        (
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
        )
        if use_common_student_columns
        else None
    )
    apply_xlsxwriter_sheet_frame(
        ws,
        repeat_last_row=header_row_index,
        freeze_row=first_data_row,
        freeze_col=0,
        select_row=first_data_row,
        select_col=0,
        paper_size=XLSX_PAPER_SIZE_A4,
        landscape=effective_landscape,
        margins=effective_margins,
    )


def _build_multi_column_copy_sheet_spec(
    *,
    title: str,
    header: Sequence[str],
    rows: Sequence[Sequence[Any]],
    metadata_rows: Sequence[Sequence[Any]] | None = None,
) -> dict[str, Any]:
    anchors = []
    metadata = list(metadata_rows or [])
    for row_index, row in enumerate(metadata, start=1):
        anchors.append([f"B{row_index}", row[0] if len(row) > 0 else ""])
        anchors.append([f"C{row_index}", row[1] if len(row) > 1 else ""])

    header_row = len(metadata) + 2 if metadata else 1
    first_data_row = header_row + 1
    for row_index, row in enumerate(rows, start=first_data_row):
        for col_index, _header in enumerate(header):
            anchors.append(
                [
                    f"{_excel_col_name(col_index)}{row_index}",
                    row[col_index] if col_index < len(row) else "",
                ]
            )
    return {
        LAYOUT_SHEET_SPEC_KEY_NAME: title,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: header_row,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: list(header),
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: anchors,
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }


def _filter_marks_template_metadata_rows(
    metadata_rows: Sequence[Sequence[Any]],
) -> list[list[Any]]:
    filtered: list[list[Any]] = []
    for row in metadata_rows:
        if len(row) < 2:
            continue
        field_key = normalize(row[0])
        if field_key == normalize(COURSE_METADATA_FACULTY_NAME_KEY):
            continue
        filtered.append([row[0], row[1]])
    return filtered


def _dynamic_direct_co_wise_headers(question_count: int) -> list[str]:
    return list(
        resolve_dynamic_sheet_headers(
            _TEMPLATE_ID,
            sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE,
            context={"question_count": question_count},
        )
    )


def _dynamic_direct_non_co_wise_headers(covered_cos: list[int]) -> list[str]:
    return list(
        resolve_dynamic_sheet_headers(
            _TEMPLATE_ID,
            sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
            context={"covered_cos": covered_cos},
        )
    )


def _dynamic_indirect_headers(total_outcomes: int) -> list[str]:
    return list(
        resolve_dynamic_sheet_headers(
            _TEMPLATE_ID,
            sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT,
            context={"total_outcomes": total_outcomes},
        )
    )


def _build_direct_co_wise_sheet_spec(
    *,
    sheet_name: str,
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
    students: Sequence[tuple[str, str]],
    questions: Sequence[dict[str, Any]],
    student_identity_hash: str,
) -> dict[str, Any]:
    header_start_row = len(metadata_rows) + 2
    header_row = header_start_row + 1
    question_count = len(questions)
    total_col = 3 + question_count
    sheet_headers = _dynamic_direct_co_wise_headers(question_count)
    max_marks_values = [float(question["max_marks"]) for question in questions]
    total_header = sheet_headers[-1]

    anchors = _component_metadata_anchor_cells(metadata_rows)
    component_row = len(metadata_rows) + 1
    anchors.extend(
        [
            [f"B{component_row}", _COMPONENT_NAME_LABEL],
            [f"C{component_row}", component_name],
            [f"C{header_row + 1}", _CO_LABEL],
            [f"C{header_row + 2}", _MAX_LABEL],
            [f"{_excel_col_name(total_col)}{header_row}", total_header],
        ]
    )

    formula_anchors: list[list[str]] = []
    if students and question_count > 0:
        first_data_row = header_start_row + 3
        first_mark_col = _excel_col_name(3)
        last_mark_col = _excel_col_name(total_col - 1)
        first_row_formula = _build_total_formula_with_absent(
            first_mark_col_name=first_mark_col,
            last_mark_col_name=last_mark_col,
            row_1_based=first_data_row + 1,
        )
        formula_anchors.append([f"{_excel_col_name(total_col)}{first_data_row + 1}", first_row_formula])

    return {
        LAYOUT_SHEET_SPEC_KEY_NAME: sheet_name,
        LAYOUT_SHEET_SPEC_KEY_KIND: LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: header_row,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: sheet_headers,
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: anchors,
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: formula_anchors,
        LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT: len(students),
        LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH: student_identity_hash,
        LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE: {
            "mark_maxima": max_marks_values,
        },
    }


def _build_direct_non_co_wise_sheet_spec(
    *,
    sheet_name: str,
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
    students: Sequence[tuple[str, str]],
    questions: Sequence[dict[str, Any]],
    student_identity_hash: str,
) -> dict[str, Any]:
    header_start_row = len(metadata_rows) + 2
    header_row = header_start_row + 1
    covered_cos = sorted({co for q in questions for co in q["co_values"]})
    sheet_headers = _dynamic_direct_non_co_wise_headers(covered_cos)
    total_header = sheet_headers[len(MARKS_ENTRY_ROW_HEADERS)]
    total_max = sum(float(question["max_marks"]) for question in questions)
    max_marks_per_co = _split_equal_with_residual(total_max, max(1, len(covered_cos)))
    mark_maxima = [total_max] + [float(value) for value in max_marks_per_co]

    anchors = _component_metadata_anchor_cells(metadata_rows)
    component_row = len(metadata_rows) + 1
    anchors.extend(
        [
            [f"B{component_row}", _COMPONENT_NAME_LABEL],
            [f"C{component_row}", component_name],
            [f"C{header_row + 1}", _CO_LABEL],
            [f"C{header_row + 2}", _MAX_LABEL],
            [f"D{header_row}", total_header],
        ]
    )

    formula_anchors: list[list[str]] = []
    if students and covered_cos:
        first_data_row = header_start_row + 3
        first_row = first_data_row + 1
        divisor = len(covered_cos)
        col_name_total = _excel_col_name(3)
        first_co_col_name = _excel_col_name(4) if divisor > 1 else ""
        for idx in range(len(covered_cos)):
            co_col = 4 + idx
            if idx == len(covered_cos) - 1 and len(covered_cos) > 1:
                prev_co_col_name = _excel_col_name(co_col - 1)
                formula = _build_direct_non_co_formula(
                    total_col_name=col_name_total,
                    row_1_based=first_row,
                    divisor=divisor,
                    first_co_col_name=first_co_col_name,
                    prev_co_col_name=prev_co_col_name,
                    is_last_residual=True,
                )
            else:
                formula = _build_direct_non_co_formula(
                    total_col_name=col_name_total,
                    row_1_based=first_row,
                    divisor=divisor,
                    first_co_col_name=first_co_col_name,
                    prev_co_col_name="",
                    is_last_residual=False,
                )
            formula_anchors.append([f"{_excel_col_name(co_col)}{first_row}", formula])

    return {
        LAYOUT_SHEET_SPEC_KEY_NAME: sheet_name,
        LAYOUT_SHEET_SPEC_KEY_KIND: LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: header_row,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: sheet_headers,
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: anchors,
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: formula_anchors,
        LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT: len(students),
        LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH: student_identity_hash,
        LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE: {
            "mark_maxima": mark_maxima,
        },
    }


def _build_indirect_sheet_spec(
    *,
    sheet_name: str,
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
    students: Sequence[tuple[str, str]],
    total_outcomes: int,
    student_identity_hash: str,
) -> dict[str, Any]:
    header_start_row = len(metadata_rows) + 2
    header_row = header_start_row + 1
    headers = _dynamic_indirect_headers(total_outcomes)
    anchors = _component_metadata_anchor_cells(metadata_rows)
    component_row = len(metadata_rows) + 1
    anchors.extend(
        [
            [f"B{component_row}", _COMPONENT_NAME_LABEL],
            [f"C{component_row}", component_name],
        ]
    )
    return {
        LAYOUT_SHEET_SPEC_KEY_NAME: sheet_name,
        LAYOUT_SHEET_SPEC_KEY_KIND: LAYOUT_SHEET_KIND_INDIRECT,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: header_row,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: headers,
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: anchors,
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
        LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT: len(students),
        LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH: student_identity_hash,
        LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE: {
            "likert_range": [LIKERT_MIN, LIKERT_MAX],
        },
    }


def _write_direct_co_wise_sheet(
    workbook: Any,
    sheet_name: str,
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
    students: Sequence[tuple[str, str]],
    questions: Sequence[dict[str, Any]],
    header_fmt: Any,
    body_fmt: Any,
    wrapped_body_fmt: Any,
    wrapped_column_fmt: Any,
    num_fmt: Any,
    header_num_fmt: Any,
    unlocked_body_fmt: Any,
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    header_start_row = _write_component_course_metadata(ws, metadata_rows, component_name, body_fmt)
    question_count = len(questions)
    total_col = 3 + question_count
    sheet_headers = _dynamic_direct_co_wise_headers(question_count)
    row_header_count = len(MARKS_ENTRY_ROW_HEADERS)
    row_headers = sheet_headers[:row_header_count]
    question_headers = sheet_headers[row_header_count:-1]
    total_header = sheet_headers[-1]
    co_labels = [f"{CO_LABEL}{question['co_values'][0]}" for question in questions]
    max_marks_values = [float(question["max_marks"]) for question in questions]

    ws.write_row(header_start_row, 0, row_headers, header_fmt)
    if question_headers:
        ws.write_row(header_start_row, 3, question_headers, header_fmt)
    ws.write(header_start_row, total_col, total_header, header_fmt)

    ws.write_row(header_start_row + 1, 0, ["", "", _CO_LABEL], header_fmt)
    if co_labels:
        ws.write_row(header_start_row + 1, 3, co_labels, header_fmt)
    ws.write(header_start_row + 1, total_col, "", header_fmt)

    ws.write_row(header_start_row + 2, 0, ["", "", _MAX_LABEL], header_fmt)
    component_total = sum(max_marks_values)
    if max_marks_values:
        ws.write_row(header_start_row + 2, 3, max_marks_values, header_num_fmt)
    ws.write_number(header_start_row + 2, total_col, component_total, header_num_fmt)

    first_data_row = header_start_row + 3
    first_mark_col = _excel_col_name(3)
    last_mark_col = _excel_col_name(total_col - 1)
    blank_marks_row = [None] * question_count
    for row_offset, (reg_no, student_name) in enumerate(students, start=first_data_row):
        ws.write_number(row_offset, 0, row_offset - (first_data_row - 1), body_fmt)
        ws.write(row_offset, 1, reg_no, body_fmt)
        ws.write(row_offset, 2, student_name, wrapped_body_fmt)
        if blank_marks_row:
            ws.write_row(row_offset, 3, blank_marks_row, unlocked_body_fmt)
        ws.write_formula(
            row_offset,
            total_col,
            _build_total_formula_with_absent(
                first_mark_col_name=first_mark_col,
                last_mark_col_name=last_mark_col,
                row_1_based=row_offset + 1,
            ),
            num_fmt,
        )

    if students and question_count > 0:
        first_row = first_data_row
        last_row = first_data_row + len(students) - 1
        max_marks_row = header_start_row + 2
        for idx, max_marks_value in enumerate(max_marks_values):
            col_index = 3 + idx
            validation_formula = _build_marks_validation_formula_for_column(
                col_index=col_index,
                first_data_row=first_data_row,
                max_marks_row=max_marks_row,
            )
            ws.data_validation(
                first_row,
                col_index,
                last_row,
                col_index,
                {
                    _VALIDATION_KEY_KIND: _VALIDATION_KIND_CUSTOM,
                    _VALIDATION_KEY_VALUE: validation_formula,
                    _VALIDATION_KEY_ERROR_TITLE: MARKS_ENTRY_VALIDATION_ERROR_TITLE,
                    _VALIDATION_KEY_ERROR_MESSAGE: _build_marks_validation_error_message(max_marks_value),
                    _VALIDATION_KEY_IGNORE_BLANK: True,
                },
            )

    sample_rows: list[list[Any]] = _component_metadata_sample_rows(metadata_rows, component_name) + [
        sheet_headers,
        ["", "", _CO_LABEL] + co_labels + [""],
        ["", "", _MAX_LABEL] + max_marks_values + [component_total],
    ]
    preview_students = students[: max(0, XLSX_AUTOFIT_SAMPLE_ROWS - len(sample_rows))]
    for row_offset, (reg_no, student_name) in enumerate(preview_students, start=first_data_row):
        sample_rows.append(
            [row_offset - (first_data_row - 1), reg_no, student_name] + [""] * question_count + [""]
        )
    _set_common_student_columns(ws, total_col, sample_rows, wrapped_column_fmt)
    apply_xlsxwriter_sheet_frame(
        ws,
        repeat_last_row=header_start_row + 2,
        freeze_row=header_start_row + 3,
        freeze_col=3,
        select_row=first_data_row,
        select_col=3,
        paper_size=XLSX_PAPER_SIZE_A4,
        landscape=True,
        margins=(
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
        ),
    )


def _write_direct_non_co_wise_sheet(
    workbook: Any,
    sheet_name: str,
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
    students: Sequence[tuple[str, str]],
    questions: Sequence[dict[str, Any]],
    header_fmt: Any,
    body_fmt: Any,
    wrapped_body_fmt: Any,
    wrapped_column_fmt: Any,
    num_fmt: Any,
    header_num_fmt: Any,
    unlocked_body_fmt: Any,
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    header_start_row = _write_component_course_metadata(ws, metadata_rows, component_name, body_fmt)
    covered_cos = sorted({co for q in questions for co in q["co_values"]})
    co_count = max(1, len(covered_cos))
    total_max = sum(float(question["max_marks"]) for question in questions)
    max_marks_per_co = _split_equal_with_residual(total_max, co_count)
    sheet_headers = _dynamic_direct_non_co_wise_headers(covered_cos)
    row_header_count = len(MARKS_ENTRY_ROW_HEADERS)
    row_headers = sheet_headers[:row_header_count]
    total_header = sheet_headers[row_header_count]
    co_mark_headers = sheet_headers[row_header_count + 1 :]
    co_prefix_labels = [f"{CO_LABEL}{co}" for co in covered_cos]

    ws.write_row(header_start_row, 0, row_headers + [total_header], header_fmt)
    if co_mark_headers:
        ws.write_row(header_start_row, 4, co_mark_headers, header_fmt)

    ws.write_row(header_start_row + 1, 0, ["", "", _CO_LABEL, ""], header_fmt)
    if co_prefix_labels:
        ws.write_row(header_start_row + 1, 4, co_prefix_labels, header_fmt)

    ws.write_row(header_start_row + 2, 0, ["", "", _MAX_LABEL, ""], header_fmt)
    ws.write_number(header_start_row + 2, 3, total_max, header_num_fmt)
    if max_marks_per_co:
        ws.write_row(header_start_row + 2, 4, max_marks_per_co, header_num_fmt)

    first_data_row = header_start_row + 3
    co_total = len(covered_cos)
    col_name_total = _excel_col_name(3)
    divisor = co_total if co_total else 1
    first_co_col_name = _excel_col_name(4) if co_total > 1 else ""
    for row_offset, (reg_no, student_name) in enumerate(students, start=first_data_row):
        ws.write_number(row_offset, 0, row_offset - (first_data_row - 1), body_fmt)
        ws.write(row_offset, 1, reg_no, body_fmt)
        ws.write(row_offset, 2, student_name, wrapped_body_fmt)
        ws.write_blank(row_offset, 3, None, unlocked_body_fmt)
        if co_total > 0:
            formula_values: list[str] = []
            for idx in range(co_total):
                if idx == co_total - 1 and co_total > 1:
                    prev_co_col_name = _excel_col_name(4 + idx - 1)
                    formula_values.append(
                        _build_direct_non_co_formula(
                            total_col_name=col_name_total,
                            row_1_based=row_offset + 1,
                            divisor=divisor,
                            first_co_col_name=first_co_col_name,
                            prev_co_col_name=prev_co_col_name,
                            is_last_residual=True,
                        )
                    )
                else:
                    formula_values.append(
                        _build_direct_non_co_formula(
                            total_col_name=col_name_total,
                            row_1_based=row_offset + 1,
                            divisor=divisor,
                            first_co_col_name=first_co_col_name,
                            prev_co_col_name="",
                            is_last_residual=False,
                        )
                    )
            ws.write_row(row_offset, 4, formula_values, num_fmt)

    if students:
        first_row = first_data_row
        last_row = first_data_row + len(students) - 1
        validation_formula = _build_marks_validation_formula_for_column(
            col_index=3,
            first_data_row=first_data_row,
            max_marks_row=header_start_row + 2,
        )
        ws.data_validation(
            first_row,
            3,
            last_row,
            3,
            {
                _VALIDATION_KEY_KIND: _VALIDATION_KIND_CUSTOM,
                _VALIDATION_KEY_VALUE: validation_formula,
                _VALIDATION_KEY_ERROR_TITLE: MARKS_ENTRY_VALIDATION_ERROR_TITLE,
                _VALIDATION_KEY_ERROR_MESSAGE: _build_marks_validation_error_message(total_max),
                _VALIDATION_KEY_IGNORE_BLANK: True,
            },
        )

    sample_rows: list[list[Any]] = _component_metadata_sample_rows(metadata_rows, component_name) + [
        sheet_headers,
        ["", "", _CO_LABEL, ""] + co_prefix_labels,
        ["", "", _MAX_LABEL, total_max] + max_marks_per_co,
    ]
    preview_students = students[: max(0, XLSX_AUTOFIT_SAMPLE_ROWS - len(sample_rows))]
    for row_offset, (reg_no, student_name) in enumerate(preview_students, start=first_data_row):
        sample_rows.append([row_offset - (first_data_row - 1), reg_no, student_name, ""] + [""] * len(covered_cos))
    _set_common_student_columns(ws, 3 + len(covered_cos), sample_rows, wrapped_column_fmt)
    apply_xlsxwriter_sheet_frame(
        ws,
        repeat_last_row=header_start_row + 2,
        freeze_row=header_start_row + 3,
        freeze_col=3,
        select_row=first_data_row,
        select_col=3,
        paper_size=XLSX_PAPER_SIZE_A4,
        landscape=True,
        margins=(
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
        ),
    )


def _write_indirect_sheet(
    workbook: Any,
    sheet_name: str,
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
    students: Sequence[tuple[str, str]],
    total_outcomes: int,
    header_fmt: Any,
    body_fmt: Any,
    unlocked_body_fmt: Any,
    wrapped_body_fmt: Any,
    wrapped_column_fmt: Any,
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    header_start_row = _write_component_course_metadata(ws, metadata_rows, component_name, body_fmt)
    headers = _dynamic_indirect_headers(total_outcomes)
    ws.write_row(header_start_row, 0, headers, header_fmt)

    first_data_row = header_start_row + 1
    for row_offset, (reg_no, student_name) in enumerate(students, start=first_data_row):
        ws.write_number(row_offset, 0, row_offset - header_start_row, body_fmt)
        ws.write(row_offset, 1, reg_no, body_fmt)
        ws.write(row_offset, 2, student_name, wrapped_body_fmt)
        for col in range(3, 3 + total_outcomes):
            ws.write_blank(row_offset, col, None, unlocked_body_fmt)

    if students and total_outcomes > 0:
        first_row = first_data_row
        last_row = first_data_row + len(students) - 1
        ws.data_validation(
            first_row,
            3,
            last_row,
            2 + total_outcomes,
            {
                _VALIDATION_KEY_KIND: _VALIDATION_KIND_CUSTOM,
                _VALIDATION_KEY_VALUE: _build_indirect_validation_formula(
                    excel_data_row=first_data_row + 1,
                ),
                _VALIDATION_KEY_ERROR_TITLE: MARKS_ENTRY_VALIDATION_ERROR_TITLE,
                _VALIDATION_KEY_ERROR_MESSAGE: MARKS_ENTRY_INDIRECT_VALIDATION_ERROR_RANGE_TEMPLATE.format(
                    minimum=f"{LIKERT_MIN:g}",
                    maximum=f"{LIKERT_MAX:g}",
                ),
                _VALIDATION_KEY_IGNORE_BLANK: True,
            },
        )

    sample_rows: list[list[Any]] = _component_metadata_sample_rows(metadata_rows, component_name) + [headers]
    preview_students = students[: max(0, XLSX_AUTOFIT_SAMPLE_ROWS - len(sample_rows))]
    for row_index, (reg_no, student_name) in enumerate(preview_students, start=1):
        sample_rows.append([row_index, reg_no, student_name] + [""] * total_outcomes)
    _set_common_student_columns(ws, 2 + total_outcomes, sample_rows, wrapped_column_fmt)
    apply_xlsxwriter_sheet_frame(
        ws,
        repeat_last_row=header_start_row,
        freeze_row=header_start_row + 1,
        freeze_col=3,
        select_row=first_data_row,
        select_col=3,
        paper_size=XLSX_PAPER_SIZE_A4,
        landscape=True,
        margins=(
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
            XLSX_PAGE_MIN_MARGIN_IN,
        ),
    )


def _write_component_course_metadata(
    ws: Any,
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
    body_fmt: Any,
) -> int:
    for row_index, row in enumerate(metadata_rows):
        ws.write(row_index, 1, row[0] if len(row) > 0 else "", body_fmt)
        ws.write(row_index, 2, row[1] if len(row) > 1 else "", body_fmt)
    component_row = len(metadata_rows)
    ws.write(component_row, 1, _COMPONENT_NAME_LABEL, body_fmt)
    ws.write(component_row, 2, component_name, body_fmt)
    return len(metadata_rows) + 2


def _component_metadata_sample_rows(
    metadata_rows: Sequence[Sequence[Any]],
    component_name: str,
) -> list[list[Any]]:
    sample_rows: list[list[Any]] = []
    for row in metadata_rows:
        sample_rows.append(["", row[0] if len(row) > 0 else "", row[1] if len(row) > 1 else ""])
    sample_rows.append(["", _COMPONENT_NAME_LABEL, component_name])
    return sample_rows


def _component_metadata_anchor_cells(metadata_rows: Sequence[Sequence[Any]]) -> list[list[Any]]:
    anchors: list[list[Any]] = []
    for row_index, row in enumerate(metadata_rows, start=1):
        anchors.append([f"B{row_index}", row[0] if len(row) > 0 else ""])
        anchors.append([f"C{row_index}", row[1] if len(row) > 1 else ""])
    return anchors


def _student_identity_hash(students: Sequence[tuple[str, str]]) -> str:
    # Stable signature of ordered student identities copied from course details.
    payload = "\n".join(f"{reg_no.strip()}|{student_name.strip()}" for reg_no, student_name in students)
    return sign_payload(payload)


def _build_marks_validation_formula_for_column(
    col_index: int,
    first_data_row: int,
    max_marks_row: int,
) -> str:
    col_name = _excel_col_name(col_index)
    excel_data_row = first_data_row + 1
    excel_max_row = max_marks_row + 1
    return (
        f'=OR({col_name}{excel_data_row}="A",{col_name}{excel_data_row}="a",'
        f"AND(ISNUMBER({col_name}{excel_data_row}),{col_name}{excel_data_row}>={MIN_MARK_VALUE},"
        f"{col_name}{excel_data_row}<={col_name}${excel_max_row}))"
    )


def _build_indirect_validation_formula(*, excel_data_row: int) -> str:
    return (
        f'=OR(D{excel_data_row}="A",D{excel_data_row}="a",'
        f'AND(ISNUMBER(D{excel_data_row}),D{excel_data_row}>={MIN_MARK_VALUE},'
        f'D{excel_data_row}>={LIKERT_MIN},D{excel_data_row}<={LIKERT_MAX}))'
    )


def _build_direct_non_co_formula(
    *,
    total_col_name: str,
    row_1_based: int,
    divisor: int,
    first_co_col_name: str,
    prev_co_col_name: str,
    is_last_residual: bool,
) -> str:
    total_ref = f"${total_col_name}{row_1_based}"
    if is_last_residual:
        sum_expr = _FORMULA_SUM_TEMPLATE.format(
            start=f"{first_co_col_name}{row_1_based}",
            end=f"{prev_co_col_name}{row_1_based}",
        ).lstrip("=")
        return (
            f'=IF(OR({total_ref}="A",{total_ref}="a"),'
            f'"A",IF({total_ref}="","",{total_ref}-'
            f"{sum_expr}))"
        )
    return (
        f'=IF(OR({total_ref}="A",{total_ref}="a"),'
        f'"A",IF({total_ref}="","",ROUND({total_ref}/{divisor},2)))'
    )


def _build_total_formula_with_absent(
    *,
    first_mark_col_name: str,
    last_mark_col_name: str,
    row_1_based: int,
) -> str:
    marks_range = f"{first_mark_col_name}{row_1_based}:{last_mark_col_name}{row_1_based}"
    return (
        f'=IF(COUNTIF({marks_range},"A")+COUNTIF({marks_range},"a")>0,'
        f'"A",SUM({marks_range}))'
    )


def _build_marks_validation_error_message(max_marks_value: Any) -> str:
    coerced_max = coerce_excel_number(max_marks_value)
    if isinstance(coerced_max, bool) or not isinstance(coerced_max, (int, float)):
        max_value_text = str(max_marks_value).strip()
    else:
        max_value_text = f"{coerced_max:g}"
    return MARKS_ENTRY_VALIDATION_ERROR_RANGE_TEMPLATE.format(
        minimum=f"{MIN_MARK_VALUE:g}",
        maximum=max_value_text,
    )


def _set_common_student_columns(
    ws: Any,
    last_col: int,
    sample_rows: Sequence[Sequence[Any]],
    wrapped_c_column_format: Any,
) -> None:
    widths = compute_sampled_column_widths(
        sample_rows,
        last_col,
        min_width=XLSX_AUTOFIT_MIN_WIDTH,
        max_width=XLSX_AUTOFIT_MAX_WIDTH,
        padding=XLSX_AUTOFIT_PADDING,
    )
    apply_xlsxwriter_column_widths(
        ws,
        widths,
        default_width=XLSX_AUTOFIT_MIN_WIDTH,
        wrap_columns=(2,),
        wrap_format=wrapped_c_column_format,
    )

def _excel_col_name(col_index: int) -> str:
    return _excel_col_name_one_based(col_index + 1)


def _split_equal_with_residual(total: float, parts: int) -> list[float]:
    if parts <= 0:
        return []
    if parts == 1:
        return [round(total, 2)]
    base = round(total / parts, 2)
    values = [base] * parts
    values[-1] = round(total - sum(values[:-1]), 2)
    return values


def _safe_sheet_name(name: str, used_sheet_names: set[str]) -> str:
    token = re.sub(r"[:\\/?*\[\]]", "_", name).strip() or "Component"
    token = token[:MAX_EXCEL_SHEETNAME_LENGTH]
    base_key = normalize(token)
    if base_key not in used_sheet_names:
        used_sheet_names.add(base_key)
        return token

    counter = 2
    while True:
        suffix = f"_{counter}"
        trimmed = token[: max(1, MAX_EXCEL_SHEETNAME_LENGTH - len(suffix))]
        candidate = f"{trimmed}{suffix}"
        key = normalize(candidate)
        if key not in used_sheet_names:
            used_sheet_names.add(key)
            return candidate
        counter += 1


def generate_worksheet(
    workbook: Any,
    sheet_name: str,
    headers: Sequence[str],
    data: Sequence[Sequence[Any]],
    header_format: Any,
    body_format: Any,
    *,
    wrap_columns: Sequence[int] = (),
    wrapped_body_format: Any | None = None,
    wrapped_column_format: Any | None = None,
) -> Any:
    """Create a worksheet with strict validation and efficient row writes."""
    if not sheet_name or not isinstance(sheet_name, str):
        raise validation_error_from_key("instructor.validation.invalid_sheet_name")

    if not headers:
        raise validation_error_from_key("instructor.validation.headers_empty", sheet_name=sheet_name)

    if len(set(headers)) != len(headers):
        raise validation_error_from_key("instructor.validation.headers_unique", sheet_name=sheet_name)

    column_count = len(headers)
    for row_index, row in enumerate(data, start=1):
        if len(row) != column_count:
            raise validation_error_from_key(
                
                    "instructor.validation.row_length_mismatch",
                    row=row_index,
                    sheet_name=sheet_name,
                    expected=column_count,
                    found=len(row),
                
            )

    worksheet = workbook.add_worksheet(sheet_name)
    col_widths: dict[int, int] = {}
    write_row = worksheet.write_row
    write_cell = worksheet.write
    write_row(0, 0, headers, header_format)
    for col_index, value in enumerate(headers):
        col_widths[col_index] = max(12, len(str(value)) + 2)

    wrap_set = set(wrap_columns)
    apply_xlsxwriter_column_widths(
        worksheet,
        col_widths,
        default_width=12,
        wrap_columns=tuple(wrap_set),
        wrap_format=wrapped_column_format or wrapped_body_format,
    )

    for row_offset, row_values in enumerate(data, start=1):
        if wrap_set and wrapped_body_format is not None:
            for col_index, value in enumerate(row_values):
                write_cell(
                    row_offset,
                    col_index,
                    value,
                    wrapped_body_format if col_index in wrap_set else body_format,
                )
        else:
            write_row(row_offset, 0, row_values, body_format)

    apply_xlsxwriter_viewport(
        worksheet,
        freeze_row=1,
        freeze_col=0,
    )
    return worksheet


def _apply_validation(worksheet: Any, rule: ValidationRule) -> None:
    options = dict(rule.options)
    validation_type = options.pop("validate", None)
    if not validation_type:
        return

    options["validate"] = validation_type
    if "ignore_blank" not in options:
        options["ignore_blank"] = True

    worksheet.data_validation(
        rule.first_row,
        rule.first_col,
        rule.last_row,
        rule.last_col,
        options,
    )


def _protect_sheet(worksheet: Any) -> None:
    # Keep locked-cell selection disabled and unlocked-cell selection enabled so
    # keyboard navigation (Tab) jumps between mark-entry cells.
    protect_xlsxwriter_sheet(worksheet)


def write_schema_sheet(
    *,
    workbook: Any,
    sheet_schema: SheetSchema,
    data: Sequence[Sequence[Any]],
    header_format: Any,
    body_format: Any,
    cancel_token: Any | None = None,
    wrap_columns: Sequence[int] = (),
    wrapped_body_format: Any | None = None,
    wrapped_column_format: Any | None = None,
) -> Any:
    if cancel_token is not None:
        cancel_token.raise_if_cancelled()
    if len(sheet_schema.header_matrix) != 1:
        raise validation_error_from_key(
            "instructor.validation.sheet_single_header_row",
            code="SHEET_HEADER_MATRIX_INVALID",
            sheet_name=sheet_schema.name,
        )
    worksheet = generate_worksheet(
        workbook=workbook,
        sheet_name=sheet_schema.name,
        headers=sheet_schema.header_matrix[0],
        data=data,
        header_format=header_format,
        body_format=body_format,
        wrap_columns=wrap_columns,
        wrapped_body_format=wrapped_body_format,
        wrapped_column_format=wrapped_column_format,
    )
    for validation in sheet_schema.validations:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        _apply_validation(worksheet, validation)
    if sheet_schema.is_protected:
        _protect_sheet(worksheet)
    return worksheet


