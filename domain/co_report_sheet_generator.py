"""Single source generator for CO direct/indirect report worksheets."""

from __future__ import annotations

from typing import Any, Protocol

from common.constants import (
    CO_REPORT_ABSENT_TOKEN,
    CO_REPORT_DIRECT_SHEET_SUFFIX,
    CO_REPORT_HEADER_REG_NO,
    CO_REPORT_HEADER_SERIAL,
    CO_REPORT_HEADER_STUDENT_NAME,
    CO_REPORT_HEADER_TOTAL,
    CO_REPORT_HEADER_TOTAL_100,
    CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE,
    CO_REPORT_INDIRECT_SHEET_SUFFIX,
    CO_REPORT_MAX_DECIMAL_PLACES,
    CO_REPORT_METADATA_OUTCOME_FIELD,
    CO_REPORT_METADATA_OUTCOME_VALUE_INDIRECT_TEMPLATE,
    CO_REPORT_METADATA_OUTCOME_VALUE_TEMPLATE,
    CO_REPORT_NOT_APPLICABLE_TOKEN,
    CO_REPORT_PERCENT_SYMBOL,
    COURSE_METADATA_FACULTY_NAME_KEY,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    DIRECT_RATIO,
    INDIRECT_RATIO,
    LIKERT_MAX,
    LIKERT_MIN,
)
from common.registry import (
    CO_REPORT_SHEET_KEY_CO_INDIRECT,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    get_sheet_headers_by_key,
    resolve_dynamic_sheet_headers,
)
from common.excel_sheet_layout import (
    apply_xlsxwriter_layout,
    set_two_column_metadata_widths,
)
from common.utils import normalize

_EXCEL_COL_FIRST_MARK = 4


def _course_metadata_headers(template_id: str) -> tuple[str, ...]:
    return get_sheet_headers_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)


class _DirectComponentComputedLike(Protocol):
    name: str
    weight: float
    max_by_co: dict[int, float]
    marks_by_co: dict[int, list[float | str]]


class _IndirectComponentComputedLike(Protocol):
    name: str
    weight: float
    marks_by_co: dict[int, list[float | str]]


def co_direct_sheet_name(co_index: int) -> str:
    return f"CO{co_index}{CO_REPORT_DIRECT_SHEET_SUFFIX}"


def co_indirect_sheet_name(co_index: int) -> str:
    return f"CO{co_index}{CO_REPORT_INDIRECT_SHEET_SUFFIX}"


def write_co_outcome_sheets(
    workbook: Any,
    *,
    template_id: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    direct_components: list[_DirectComponentComputedLike],
    indirect_components: list[_IndirectComponentComputedLike],
    formats: dict[str, Any],
) -> tuple[str, str]:
    direct_name = co_direct_sheet_name(co_index)
    indirect_name = co_indirect_sheet_name(co_index)
    _write_direct_sheet(
        workbook,
        template_id=template_id,
        sheet_name=direct_name,
        co_index=co_index,
        metadata_rows=metadata_rows,
        students=students,
        components=direct_components,
        formats=formats,
    )
    _write_indirect_sheet(
        workbook,
        template_id=template_id,
        sheet_name=indirect_name,
        co_index=co_index,
        metadata_rows=metadata_rows,
        students=students,
        components=indirect_components,
        formats=formats,
    )
    return direct_name, indirect_name


def write_co_direct_sheet(
    workbook: Any,
    *,
    template_id: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    components: list[_DirectComponentComputedLike],
    formats: dict[str, Any],
) -> str:
    direct_name = co_direct_sheet_name(co_index)
    _write_direct_sheet(
        workbook,
        template_id=template_id,
        sheet_name=direct_name,
        co_index=co_index,
        metadata_rows=metadata_rows,
        students=students,
        components=components,
        formats=formats,
    )
    return direct_name


def write_co_indirect_sheet(
    workbook: Any,
    *,
    template_id: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    components: list[_IndirectComponentComputedLike],
    formats: dict[str, Any],
) -> str:
    indirect_name = co_indirect_sheet_name(co_index)
    _write_indirect_sheet(
        workbook,
        template_id=template_id,
        sheet_name=indirect_name,
        co_index=co_index,
        metadata_rows=metadata_rows,
        students=students,
        components=components,
        formats=formats,
    )
    return indirect_name


def _write_report_metadata(
    ws: Any,
    *,
    template_id: str,
    metadata_rows: list[tuple[str, Any]],
    formats: dict[str, Any],
) -> int:
    metadata_headers = _course_metadata_headers(template_id)
    ws.write(0, 1, metadata_headers[0], formats["header"])
    ws.write(0, 2, metadata_headers[1], formats["header"])
    for idx, (field, value) in enumerate(metadata_rows, start=1):
        ws.write(idx, 1, field, formats["body"])
        ws.write(idx, 2, value, formats["body_wrap"])
    return len(metadata_rows) + 2


def _set_report_metadata_column_widths(
    ws: Any,
    *,
    template_id: str,
    metadata_rows: list[tuple[str, Any]],
) -> None:
    metadata_headers = _course_metadata_headers(template_id)
    set_two_column_metadata_widths(
        ws,
        col1_title=metadata_headers[0],
        col2_title=metadata_headers[1],
        rows=metadata_rows,
        col1_index=1,
        col2_index=2,
    )


def _apply_layout(ws: Any, *, header_row_index: int, paper_size: int, landscape: bool) -> None:
    apply_xlsxwriter_layout(
        ws,
        header_row_index=header_row_index,
        paper_size=paper_size,
        landscape=landscape,
    )


def _report_metadata_rows(
    metadata_rows: list[tuple[str, Any]],
    *,
    co_index: int,
    outcome_value_template: str,
) -> list[tuple[str, Any]]:
    filtered: list[tuple[str, Any]] = []
    drop_keys = {
        normalize(COURSE_METADATA_FACULTY_NAME_KEY),
        normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY),
    }
    for field, value in metadata_rows:
        if normalize(field) in drop_keys:
            continue
        filtered.append((field, value))
    filtered.append((CO_REPORT_METADATA_OUTCOME_FIELD, outcome_value_template.format(co=co_index)))
    return filtered


def _write_direct_sheet(
    workbook: Any,
    *,
    template_id: str,
    sheet_name: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    components: list[_DirectComponentComputedLike],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    report_metadata_rows = _report_metadata_rows(
        metadata_rows,
        co_index=co_index,
        outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_TEMPLATE,
    )
    header_row_index = _write_report_metadata(
        ws,
        template_id=template_id,
        metadata_rows=report_metadata_rows,
        formats=formats,
    ) - 1
    _set_report_metadata_column_widths(
        ws,
        template_id=template_id,
        metadata_rows=report_metadata_rows,
    )
    ws.repeat_rows(0, header_row_index)
    ws.freeze_panes(header_row_index + 1, _EXCEL_COL_FIRST_MARK - 1)
    active_components = [component for component in components if component.max_by_co.get(co_index, 0.0) > 0]
    total_weight = _round2(sum(component.weight for component in active_components))
    headers = [CO_REPORT_HEADER_SERIAL, CO_REPORT_HEADER_REG_NO, CO_REPORT_HEADER_STUDENT_NAME]
    for component in active_components:
        max_marks = _round2(component.max_by_co.get(co_index, 0.0))
        headers.append(f"{component.name} ({max_marks:g})")
        headers.append(f"{component.name} ({component.weight:g}{CO_REPORT_PERCENT_SYMBOL})")
    headers.append(f"{CO_REPORT_HEADER_TOTAL} ({total_weight:g}{CO_REPORT_PERCENT_SYMBOL})")
    headers.append(CO_REPORT_HEADER_TOTAL_100)
    headers.append(_ratio_total_header(DIRECT_RATIO))
    for col, value in enumerate(headers, start=0):
        ws.write(header_row_index, col, value, formats["header"])

    student_count = len(students)
    marks_by_component: list[list[Any]] = [
        component.marks_by_co.get(co_index, [0.0] * student_count)
        for component in active_components
    ]
    for idx, (reg_no, student_name) in enumerate(students, start=1):
        row_index = header_row_index + idx
        row_values: list[Any] = [idx, reg_no, student_name]
        absent = False
        weighted_total = 0.0
        for component, component_marks in zip(active_components, marks_by_component):
            max_marks = component.max_by_co.get(co_index, 0.0)
            raw = component_marks[idx - 1]
            if isinstance(raw, str) and _is_absent(raw):
                absent = True
                row_values.extend([CO_REPORT_ABSENT_TOKEN, CO_REPORT_ABSENT_TOKEN])
                continue
            raw_numeric = float(raw) if isinstance(raw, (int, float)) else 0.0
            weighted = _round2((raw_numeric * component.weight / max_marks) if max_marks > 0 else 0.0)
            weighted_total += weighted
            row_values.extend([_round2(raw_numeric), weighted])
        if absent:
            row_values.extend([CO_REPORT_NOT_APPLICABLE_TOKEN, CO_REPORT_NOT_APPLICABLE_TOKEN, CO_REPORT_NOT_APPLICABLE_TOKEN])
        else:
            total_100 = _round2((weighted_total * 100.0 / total_weight) if total_weight > 0 else 0.0)
            total_ratio = _round2(total_100 * DIRECT_RATIO)
            row_values.extend([_round2(weighted_total), total_100, total_ratio])

        for col, value in enumerate(row_values, start=0):
            if col == 2:
                fmt = formats["body_wrap"]
            elif col >= 3:
                fmt = formats["body_center"]
            else:
                fmt = formats["body"]
            ws.write(row_index, col, value, fmt)

    _apply_layout(ws, header_row_index=header_row_index, paper_size=8, landscape=True)


def _write_indirect_sheet(
    workbook: Any,
    *,
    template_id: str,
    sheet_name: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    components: list[_IndirectComponentComputedLike],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    report_metadata_rows = _report_metadata_rows(
        metadata_rows,
        co_index=co_index,
        outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_INDIRECT_TEMPLATE,
    )
    header_row_index = _write_report_metadata(
        ws,
        template_id=template_id,
        metadata_rows=report_metadata_rows,
        formats=formats,
    ) - 1
    _set_report_metadata_column_widths(
        ws,
        template_id=template_id,
        metadata_rows=report_metadata_rows,
    )
    ws.repeat_rows(0, header_row_index)
    ws.freeze_panes(header_row_index + 1, _EXCEL_COL_FIRST_MARK - 1)

    active_components = [
        component
        for component in components
        if any(not (isinstance(value, float) and value == 0.0) for value in component.marks_by_co.get(co_index, []))
    ]
    total_weight = _round2(sum(component.weight for component in active_components))
    scaled_max_value = max(0, LIKERT_MAX - LIKERT_MIN)
    has_single_component = len(active_components) == 1
    headers = list(
        resolve_dynamic_sheet_headers(
            template_id,
            sheet_key=CO_REPORT_SHEET_KEY_CO_INDIRECT,
            context={
                "components": [(component.name, component.weight) for component in active_components],
                "ratio": INDIRECT_RATIO,
            },
        )
    )
    for col, value in enumerate(headers, start=0):
        ws.write(header_row_index, col, value, formats["header"])

    student_count = len(students)
    marks_by_component: list[list[Any]] = [
        component.marks_by_co.get(co_index, [0.0] * student_count)
        for component in active_components
    ]
    denominator = float(max(1, scaled_max_value))
    for idx, (reg_no, student_name) in enumerate(students, start=1):
        row_index = header_row_index + idx
        row_values: list[Any] = [idx, reg_no, student_name]
        absent = False
        total_weighted = 0.0
        for component, component_marks in zip(active_components, marks_by_component):
            raw = component_marks[idx - 1]
            if isinstance(raw, str) and _is_absent(raw):
                absent = True
                row_values.extend([CO_REPORT_ABSENT_TOKEN, CO_REPORT_ABSENT_TOKEN])
                if not has_single_component:
                    row_values.append(CO_REPORT_ABSENT_TOKEN)
                continue
            raw_numeric = float(raw) if isinstance(raw, (int, float)) else 0.0
            scaled_raw = _round2(raw_numeric - LIKERT_MIN)
            scaled_raw = max(0.0, min(float(scaled_max_value), scaled_raw))
            if has_single_component:
                total_weighted = _round2((scaled_raw / denominator) * 100.0)
                row_values.extend([_round2(raw_numeric), scaled_raw])
            else:
                weighted = _round2((scaled_raw / denominator) * component.weight)
                total_weighted += weighted
                row_values.extend([_round2(raw_numeric), scaled_raw, weighted])
        if absent:
            row_values.extend([CO_REPORT_NOT_APPLICABLE_TOKEN, CO_REPORT_NOT_APPLICABLE_TOKEN])
        else:
            if has_single_component:
                total_100 = _round2(total_weighted)
            else:
                total_100 = _round2((total_weighted * 100.0 / total_weight) if total_weight > 0 else 0.0)
            total_ratio = _round2(total_100 * INDIRECT_RATIO)
            row_values.extend([total_100, total_ratio])

        for col, value in enumerate(row_values, start=0):
            if col == 2:
                fmt = formats["body_wrap"]
            elif col >= 3:
                fmt = formats["body_center"]
            else:
                fmt = formats["body"]
            ws.write(row_index, col, value, fmt)

    _apply_layout(ws, header_row_index=header_row_index, paper_size=9, landscape=False)


def _ratio_total_header(ratio: float) -> str:
    percent = ratio * 100.0
    if abs(percent - round(percent)) <= 1e-9:
        token = f"{int(round(percent))}"
    else:
        token = f"{percent:g}"
    return CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE.format(ratio=token)


def ratio_total_header(ratio: float) -> str:
    return _ratio_total_header(ratio)


def _round2(value: float) -> float:
    return round(float(value), CO_REPORT_MAX_DECIMAL_PLACES)


def _is_absent(value: Any) -> bool:
    return normalize(value) == normalize(CO_REPORT_ABSENT_TOKEN)
