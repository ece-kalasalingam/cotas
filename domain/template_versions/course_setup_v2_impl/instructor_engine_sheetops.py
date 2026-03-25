"""Sheet/write helpers extracted from instructor_template_engine."""

from __future__ import annotations

import json
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
    MARKS_ENTRY_CO_MARKS_LABEL_PREFIX,
    MARKS_ENTRY_INDIRECT_VALIDATION_ERROR_RANGE_TEMPLATE,
    MARKS_ENTRY_QUESTION_PREFIX,
    MARKS_ENTRY_ROW_HEADERS,
    MARKS_ENTRY_TOTAL_LABEL,
    MARKS_ENTRY_VALIDATION_ERROR_RANGE_TEMPLATE,
    MARKS_ENTRY_VALIDATION_ERROR_TITLE,
    MAX_EXCEL_SHEETNAME_LENGTH,
    MIN_MARK_VALUE,
    SYSTEM_LAYOUT_MANIFEST_HASH_HEADER,
    SYSTEM_LAYOUT_MANIFEST_HEADER,
    SYSTEM_LAYOUT_SHEET,
)
from common.registry import (
    SYSTEM_HASH_HEADER_TEMPLATE_HASH as SYSTEM_HASH_TEMPLATE_HASH_HEADER,
    SYSTEM_HASH_HEADER_TEMPLATE_ID as SYSTEM_HASH_TEMPLATE_ID_HEADER,
    SYSTEM_HASH_SHEET_NAME as SYSTEM_HASH_SHEET,
)
from common.excel_sheet_layout import (
    build_xlsxwriter_body_format,
    build_xlsxwriter_header_format,
    compute_sampled_column_widths,
    excel_col_name as _excel_col_name_one_based,
    protect_xlsxwriter_sheet,
)
from common.error_catalog import validation_error_from_key
from common.sheet_schema import ValidationRule
from common.utils import (
    coerce_excel_number,
    copy_system_hash_sheet,
    normalize,
)
from common.workbook_signing import sign_payload

_AUTO_FIT_SAMPLE_ROWS = 6
_AUTO_FIT_PADDING = 2
_AUTO_FIT_MIN_WIDTH = 8
_AUTO_FIT_MAX_WIDTH = 60
_PAGE_MIN_MARGIN_IN = 0.25
_COMPONENT_NAME_LABEL = COMPONENT_NAME_LABEL
_CO_LABEL = CO_LABEL
_MAX_LABEL = INSTRUCTOR_MAX_LABEL
_FORMULA_SUM_TEMPLATE = "=SUM({start}:{end})"
_VALIDATION_KIND_CUSTOM = "custom"
_VALIDATION_KEY_KIND = "validate"
_VALIDATION_KEY_VALUE = "value"
_VALIDATION_KEY_ERROR_TITLE = "error_title"
_VALIDATION_KEY_ERROR_MESSAGE = "error_message"
_VALIDATION_KEY_IGNORE_BLANK = "ignore_blank"

def _copy_system_hash_sheet(source_workbook: Any, target_workbook: Any) -> None:
    copy_system_hash_sheet(source_workbook, target_workbook)


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
    ws.set_column(0, 0, 24)
    ws.set_column(1, 1, 24)
    for row_index, row in enumerate(rows, start=1):
        ws.write(row_index, 0, row[0] if len(row) > 0 else "", body_fmt)
        ws.write(row_index, 1, row[1] if len(row) > 1 else "", body_fmt)
    ws.repeat_rows(0, 0)
    ws.freeze_panes(1, 0)
    ws.set_selection(1, 0, 1, 0)


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
) -> None:
    ws = workbook.add_worksheet(title)
    ws.write_row(0, 0, list(header), header_fmt)
    for col in range(len(header)):
        ws.set_column(col, col, 22)
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row[: len(header)]):
            cell_fmt = num_fmt if col_index == 1 and isinstance(value, (int, float)) else body_fmt
            ws.write(row_index, col_index, value, cell_fmt)
    ws.repeat_rows(0, 0)
    ws.freeze_panes(1, 0)
    ws.set_selection(1, 0, 1, 0)


def _build_multi_column_copy_sheet_spec(
    *,
    title: str,
    header: Sequence[str],
    rows: Sequence[Sequence[Any]],
) -> dict[str, Any]:
    anchors = []
    for row_index, row in enumerate(rows, start=2):
        for col_index, _header in enumerate(header):
            anchors.append(
                [
                    f"{_excel_col_name(col_index)}{row_index}",
                    row[col_index] if col_index < len(row) else "",
                ]
            )
    return {
        LAYOUT_SHEET_SPEC_KEY_NAME: title,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
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
    row_headers = list(MARKS_ENTRY_ROW_HEADERS)
    question_headers = [f"{MARKS_ENTRY_QUESTION_PREFIX}{idx + 1}" for idx in range(question_count)]
    max_marks_values = [float(question["max_marks"]) for question in questions]
    sheet_headers = row_headers + question_headers + [MARKS_ENTRY_TOTAL_LABEL]

    anchors = _component_metadata_anchor_cells(metadata_rows)
    component_row = len(metadata_rows) + 1
    anchors.extend(
        [
            [f"B{component_row}", _COMPONENT_NAME_LABEL],
            [f"C{component_row}", component_name],
            [f"C{header_row + 1}", _CO_LABEL],
            [f"C{header_row + 2}", _MAX_LABEL],
            [f"{_excel_col_name(total_col)}{header_row}", MARKS_ENTRY_TOTAL_LABEL],
        ]
    )

    formula_anchors: list[list[str]] = []
    if students and question_count > 0:
        first_data_row = header_start_row + 3
        first_mark_col = _excel_col_name(3)
        last_mark_col = _excel_col_name(total_col - 1)
        first_row_formula = _FORMULA_SUM_TEMPLATE.format(
            start=f"{first_mark_col}{first_data_row + 1}",
            end=f"{last_mark_col}{first_data_row + 1}",
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
    co_mark_headers = [f"{MARKS_ENTRY_CO_MARKS_LABEL_PREFIX}{co}" for co in covered_cos]
    total_max = sum(float(question["max_marks"]) for question in questions)
    max_marks_per_co = _split_equal_with_residual(total_max, max(1, len(covered_cos)))
    mark_maxima = [total_max] + [float(value) for value in max_marks_per_co]
    sheet_headers = list(MARKS_ENTRY_ROW_HEADERS) + [MARKS_ENTRY_TOTAL_LABEL] + co_mark_headers

    anchors = _component_metadata_anchor_cells(metadata_rows)
    component_row = len(metadata_rows) + 1
    anchors.extend(
        [
            [f"B{component_row}", _COMPONENT_NAME_LABEL],
            [f"C{component_row}", component_name],
            [f"C{header_row + 1}", _CO_LABEL],
            [f"C{header_row + 2}", _MAX_LABEL],
            [f"D{header_row}", MARKS_ENTRY_TOTAL_LABEL],
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
    headers = list(MARKS_ENTRY_ROW_HEADERS) + [
        f"{CO_LABEL}{i}" for i in range(1, total_outcomes + 1)
    ]
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
    row_headers = list(MARKS_ENTRY_ROW_HEADERS)
    question_headers = [f"{MARKS_ENTRY_QUESTION_PREFIX}{idx + 1}" for idx in range(question_count)]
    co_labels = [f"{CO_LABEL}{question['co_values'][0]}" for question in questions]
    max_marks_values = [float(question["max_marks"]) for question in questions]
    sheet_headers = row_headers + question_headers + [MARKS_ENTRY_TOTAL_LABEL]

    ws.write_row(header_start_row, 0, row_headers, header_fmt)
    if question_headers:
        ws.write_row(header_start_row, 3, question_headers, header_fmt)
    ws.write(header_start_row, total_col, MARKS_ENTRY_TOTAL_LABEL, header_fmt)

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
            _FORMULA_SUM_TEMPLATE.format(
                start=f"{first_mark_col}{row_offset + 1}",
                end=f"{last_mark_col}{row_offset + 1}",
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
    preview_students = students[: max(0, _AUTO_FIT_SAMPLE_ROWS - len(sample_rows))]
    for row_offset, (reg_no, student_name) in enumerate(preview_students, start=first_data_row):
        sample_rows.append(
            [row_offset - (first_data_row - 1), reg_no, student_name] + [""] * question_count + [""]
        )
    _set_common_student_columns(ws, total_col, sample_rows, wrapped_column_fmt)
    ws.repeat_rows(0, header_start_row + 2)
    ws.freeze_panes(header_start_row + 3, 3)
    ws.set_selection(first_data_row, 3, first_data_row, 3)
    _protect_sheet(ws)


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
    row_headers = list(MARKS_ENTRY_ROW_HEADERS)
    co_mark_headers = [f"{MARKS_ENTRY_CO_MARKS_LABEL_PREFIX}{co}" for co in covered_cos]
    co_prefix_labels = [f"{CO_LABEL}{co}" for co in covered_cos]
    sheet_headers = row_headers + [MARKS_ENTRY_TOTAL_LABEL] + co_mark_headers

    ws.write_row(header_start_row, 0, row_headers + [MARKS_ENTRY_TOTAL_LABEL], header_fmt)
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
    preview_students = students[: max(0, _AUTO_FIT_SAMPLE_ROWS - len(sample_rows))]
    for row_offset, (reg_no, student_name) in enumerate(preview_students, start=first_data_row):
        sample_rows.append([row_offset - (first_data_row - 1), reg_no, student_name, ""] + [""] * len(covered_cos))
    _set_common_student_columns(ws, 3 + len(covered_cos), sample_rows, wrapped_column_fmt)
    ws.repeat_rows(0, header_start_row + 2)
    ws.freeze_panes(header_start_row + 3, 3)
    ws.set_selection(first_data_row, 3, first_data_row, 3)
    _protect_sheet(ws)


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
    headers = list(MARKS_ENTRY_ROW_HEADERS) + [
        f"{CO_LABEL}{i}" for i in range(1, total_outcomes + 1)
    ]
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
    preview_students = students[: max(0, _AUTO_FIT_SAMPLE_ROWS - len(sample_rows))]
    for row_index, (reg_no, student_name) in enumerate(preview_students, start=1):
        sample_rows.append([row_index, reg_no, student_name] + [""] * total_outcomes)
    _set_common_student_columns(ws, 2 + total_outcomes, sample_rows, wrapped_column_fmt)
    ws.repeat_rows(0, header_start_row)
    ws.freeze_panes(header_start_row + 1, 3)
    ws.set_selection(first_data_row, 3, first_data_row, 3)
    _protect_sheet(ws)


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
    ws.set_paper(9)  # A4
    ws.set_landscape()
    ws.set_margins(_PAGE_MIN_MARGIN_IN, _PAGE_MIN_MARGIN_IN, _PAGE_MIN_MARGIN_IN, _PAGE_MIN_MARGIN_IN)
    ws.fit_to_pages(1, 0)

    widths = compute_sampled_column_widths(
        sample_rows,
        last_col,
        min_width=_AUTO_FIT_MIN_WIDTH,
        max_width=_AUTO_FIT_MAX_WIDTH,
        padding=_AUTO_FIT_PADDING,
    )

    for col in range(0, last_col + 1):
        width = widths.get(col, _AUTO_FIT_MIN_WIDTH)
        if col == 2:
            ws.set_column(col, col, width, wrapped_c_column_format)
        else:
            ws.set_column(col, col, width)

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


def _build_header_format(workbook: Any, header_style: dict[str, Any]) -> Any:
    return build_xlsxwriter_header_format(workbook, header_style)


def _build_body_format(workbook: Any, body_style: dict[str, Any]) -> Any:
    return build_xlsxwriter_body_format(workbook, body_style)


def generate_worksheet(
    workbook: Any,
    sheet_name: str,
    headers: Sequence[str],
    data: Sequence[Sequence[Any]],
    header_format: Any,
    body_format: Any,
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
    write_row(0, 0, headers, header_format)
    for col_index, value in enumerate(headers):
        col_widths[col_index] = max(12, len(str(value)) + 2)

    for col_index, width in col_widths.items():
        worksheet.set_column(col_index, col_index, width)

    for row_offset, row_values in enumerate(data, start=1):
        write_row(row_offset, 0, row_values, body_format)

    worksheet.freeze_panes(1, 0)
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


def _add_system_hash_sheet(workbook: Any, template_id: str) -> None:
    worksheet = workbook.add_worksheet(SYSTEM_HASH_SHEET)
    template_hash = _compute_template_hash(template_id)

    worksheet.write_row(
        0,
        0,
        [SYSTEM_HASH_TEMPLATE_ID_HEADER, SYSTEM_HASH_TEMPLATE_HASH_HEADER],
    )
    worksheet.write_row(1, 0, [template_id, template_hash])
    worksheet.hide()


def _add_system_layout_sheet(workbook: Any, layout_manifest: dict[str, Any]) -> None:
    worksheet = workbook.add_worksheet(SYSTEM_LAYOUT_SHEET)
    manifest_text = _serialize_layout_manifest(layout_manifest)
    manifest_hash = _compute_layout_manifest_hash(manifest_text)
    worksheet.write_row(
        0,
        0,
        [SYSTEM_LAYOUT_MANIFEST_HEADER, SYSTEM_LAYOUT_MANIFEST_HASH_HEADER],
    )
    worksheet.write_row(1, 0, [manifest_text, manifest_hash])
    worksheet.hide()


def _compute_template_hash(template_id: str) -> str:
    return sign_payload(template_id)


def _serialize_layout_manifest(layout_manifest: dict[str, Any]) -> str:
    return json.dumps(layout_manifest, sort_keys=True, separators=(",", ":"), ensure_ascii=True)


def _compute_layout_manifest_hash(manifest_text: str) -> str:
    return sign_payload(manifest_text)



