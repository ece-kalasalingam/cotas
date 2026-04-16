"""Single source generator for CO direct/indirect report worksheets."""

from __future__ import annotations

from typing import Any, Protocol

from common.constants import (
    CO_REPORT_ABSENT_TOKEN,
    CO_REPORT_DIRECT_SHEET_SUFFIX,
    CO_REPORT_HEADER_REG_NO,
    CO_REPORT_HEADER_SERIAL,
    CO_REPORT_HEADER_STUDENT_NAME,
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
    DIRECT_RATIO,
    INDIRECT_RATIO,
    LIKERT_MAX,
    LIKERT_MIN,
)
from common.registry import (
    CO_REPORT_SHEET_KEY_CO_INDIRECT,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    get_sheet_headers_by_key,
    resolve_dynamic_sheet_headers,
)
from common.excel_sheet_layout import (
    apply_xlsxwriter_layout,
    set_two_column_metadata_widths,
)
from common.error_catalog import validation_error_from_key
from common.utils import normalize

_EXCEL_COL_FIRST_MARK = 4


def _course_metadata_headers(template_id: str) -> tuple[str, ...]:
    """Course metadata headers.
    
    Args:
        template_id: Parameter value (str).
    
    Returns:
        tuple[str, ...]: Return value.
    
    Raises:
        None.
    """
    return get_sheet_headers_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)


class _JoinedCoAnalysisRowLike(Protocol):
    reg_no: str
    student_name: str
    direct_score: float | str
    indirect_score: float | str


class _CoReportComponentLike(Protocol):
    name: str
    weight: float
    max_by_co: dict[int, float]
    marks_by_co: dict[int, list[Any]]


def co_direct_sheet_name(co_index: int) -> str:
    """Co direct sheet name.
    
    Args:
        co_index: Parameter value (int).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    return f"CO{co_index}{CO_REPORT_DIRECT_SHEET_SUFFIX}"


def co_indirect_sheet_name(co_index: int) -> str:
    """Co indirect sheet name.
    
    Args:
        co_index: Parameter value (int).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    return f"CO{co_index}{CO_REPORT_INDIRECT_SHEET_SUFFIX}"


def write_co_outcome_joined_total_sheets(
    workbook: Any,
    *,
    template_id: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    rows: list[_JoinedCoAnalysisRowLike],
    formats: dict[str, Any],
) -> tuple[str, str]:
    """Write co outcome joined total sheets.
    
    Args:
        workbook: Parameter value (Any).
        template_id: Parameter value (str).
        co_index: Parameter value (int).
        metadata_rows: Parameter value (list[tuple[str, Any]]).
        rows: Parameter value (list[_JoinedCoAnalysisRowLike]).
        formats: Parameter value (dict[str, Any]).
    
    Returns:
        tuple[str, str]: Return value.
    
    Raises:
        None.
    """
    resolved_rows = _resolve_joined_total_rows_for_template(template_id=template_id, rows=rows)
    direct_name = co_direct_sheet_name(co_index)
    indirect_name = co_indirect_sheet_name(co_index)
    _write_joined_total_sheet(
        workbook,
        template_id=template_id,
        sheet_name=direct_name,
        co_index=co_index,
        metadata_rows=metadata_rows,
        formats=formats,
        score_header=_ratio_total_header(DIRECT_RATIO),
        score_values=[direct_score for _reg_no, _student_name, direct_score, _indirect_score in resolved_rows],
        students=[(reg_no, student_name) for reg_no, student_name, _direct_score, _indirect_score in resolved_rows],
        outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_TEMPLATE,
    )
    _write_joined_total_sheet(
        workbook,
        template_id=template_id,
        sheet_name=indirect_name,
        co_index=co_index,
        metadata_rows=metadata_rows,
        formats=formats,
        score_header=_ratio_total_header(INDIRECT_RATIO),
        score_values=[indirect_score for _reg_no, _student_name, _direct_score, indirect_score in resolved_rows],
        students=[(reg_no, student_name) for reg_no, student_name, _direct_score, _indirect_score in resolved_rows],
        outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_INDIRECT_TEMPLATE,
    )
    return direct_name, indirect_name


def write_co_outcome_sheets(
    workbook: Any,
    *,
    template_id: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    rows: list[_JoinedCoAnalysisRowLike],
    direct_components: list[_CoReportComponentLike],
    indirect_components: list[_CoReportComponentLike],
    formats: dict[str, Any],
) -> tuple[str, str]:
    """Write co outcome sheets.

    Args:
        workbook: Parameter value (Any).
        template_id: Parameter value (str).
        co_index: Parameter value (int).
        metadata_rows: Parameter value (list[tuple[str, Any]]).
        rows: Parameter value (list[_JoinedCoAnalysisRowLike]).
        direct_components: Parameter value (list[_CoReportComponentLike]).
        indirect_components: Parameter value (list[_CoReportComponentLike]).
        formats: Parameter value (dict[str, Any]).

    Returns:
        tuple[str, str]: Return value.

    Raises:
        None.
    """
    resolved_rows = _resolve_joined_total_rows_for_template(template_id=template_id, rows=rows)
    direct_name = co_direct_sheet_name(co_index)
    indirect_name = co_indirect_sheet_name(co_index)
    students = [(reg_no, student_name) for reg_no, student_name, _direct_score, _indirect_score in resolved_rows]
    if _template_supports_component_layout(template_id):
        if direct_components:
            _write_direct_sheet(
                workbook,
                template_id=template_id,
                sheet_name=direct_name,
                co_index=co_index,
                metadata_rows=metadata_rows,
                students=students,
                components=list(direct_components),
                formats=formats,
            )
        else:
            _write_joined_total_sheet(
                workbook,
                template_id=template_id,
                sheet_name=direct_name,
                co_index=co_index,
                metadata_rows=metadata_rows,
                formats=formats,
                score_header=_ratio_total_header(DIRECT_RATIO),
                score_values=[direct_score for _reg_no, _student_name, direct_score, _indirect_score in resolved_rows],
                students=students,
                outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_TEMPLATE,
            )
        if indirect_components:
            _write_indirect_sheet(
                workbook,
                template_id=template_id,
                sheet_name=indirect_name,
                co_index=co_index,
                metadata_rows=metadata_rows,
                students=students,
                components=list(indirect_components),
                formats=formats,
            )
        else:
            _write_joined_total_sheet(
                workbook,
                template_id=template_id,
                sheet_name=indirect_name,
                co_index=co_index,
                metadata_rows=metadata_rows,
                formats=formats,
                score_header=_ratio_total_header(INDIRECT_RATIO),
                score_values=[indirect_score for _reg_no, _student_name, _direct_score, indirect_score in resolved_rows],
                students=students,
                outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_INDIRECT_TEMPLATE,
            )
        return direct_name, indirect_name
    raise validation_error_from_key(
        "validation.template.unknown",
        code="UNKNOWN_TEMPLATE",
        template_id=template_id,
    )


def _write_report_metadata(
    ws: Any,
    *,
    template_id: str,
    metadata_rows: list[tuple[str, Any]],
    formats: dict[str, Any],
) -> int:
    """Write report metadata.
    
    Args:
        ws: Parameter value (Any).
        template_id: Parameter value (str).
        metadata_rows: Parameter value (list[tuple[str, Any]]).
        formats: Parameter value (dict[str, Any]).
    
    Returns:
        int: Return value.
    
    Raises:
        None.
    """
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
    """Set report metadata column widths.
    
    Args:
        ws: Parameter value (Any).
        template_id: Parameter value (str).
        metadata_rows: Parameter value (list[tuple[str, Any]]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Apply layout.
    
    Args:
        ws: Parameter value (Any).
        header_row_index: Parameter value (int).
        paper_size: Parameter value (int).
        landscape: Parameter value (bool).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Report metadata rows.
    
    Args:
        metadata_rows: Parameter value (list[tuple[str, Any]]).
        co_index: Parameter value (int).
        outcome_value_template: Parameter value (str).
    
    Returns:
        list[tuple[str, Any]]: Return value.
    
    Raises:
        None.
    """
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
    components: list[Any],
    formats: dict[str, Any],
) -> None:
    """Write direct sheet.
    
    Args:
        workbook: Parameter value (Any).
        template_id: Parameter value (str).
        sheet_name: Parameter value (str).
        co_index: Parameter value (int).
        metadata_rows: Parameter value (list[tuple[str, Any]]).
        students: Parameter value (list[tuple[str, str]]).
        components: Parameter value (list[Any]).
        formats: Parameter value (dict[str, Any]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    headers = [CO_REPORT_HEADER_SERIAL, CO_REPORT_HEADER_REG_NO, CO_REPORT_HEADER_STUDENT_NAME]
    for component in active_components:
        max_marks = _round2(component.max_by_co.get(co_index, 0.0))
        headers.append(f"{component.name} ({max_marks:g})")
        headers.append(f"{component.name} ({component.weight:g}{CO_REPORT_PERCENT_SYMBOL})")
    headers.append(CO_REPORT_HEADER_TOTAL_100)
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
        row_total_weight = 0.0
        for component, component_marks in zip(active_components, marks_by_component):
            max_marks = component.max_by_co.get(co_index, 0.0)
            raw = component_marks[idx - 1]
            if normalize(raw) == normalize(CO_REPORT_NOT_APPLICABLE_TOKEN):
                row_values.extend([CO_REPORT_NOT_APPLICABLE_TOKEN, CO_REPORT_NOT_APPLICABLE_TOKEN])
                continue
            row_total_weight += float(component.weight)
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
            total_100 = _round2((weighted_total * 100.0 / row_total_weight) if row_total_weight > 0 else 0.0)
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


def _write_joined_total_sheet(
    workbook: Any,
    *,
    template_id: str,
    sheet_name: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    formats: dict[str, Any],
    score_header: str,
    score_values: list[float | str],
    students: list[tuple[str, str]],
    outcome_value_template: str,
) -> None:
    """Write joined total sheet.
    
    Args:
        workbook: Parameter value (Any).
        template_id: Parameter value (str).
        sheet_name: Parameter value (str).
        co_index: Parameter value (int).
        metadata_rows: Parameter value (list[tuple[str, Any]]).
        formats: Parameter value (dict[str, Any]).
        score_header: Parameter value (str).
        score_values: Parameter value (list[float | str]).
        students: Parameter value (list[tuple[str, str]]).
        outcome_value_template: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    ws = workbook.add_worksheet(sheet_name)
    report_metadata_rows = _report_metadata_rows(
        metadata_rows,
        co_index=co_index,
        outcome_value_template=outcome_value_template,
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

    headers = [CO_REPORT_HEADER_SERIAL, CO_REPORT_HEADER_REG_NO, CO_REPORT_HEADER_STUDENT_NAME, score_header]
    for col, value in enumerate(headers, start=0):
        ws.write(header_row_index, col, value, formats["header"])

    for idx, (reg_no, student_name) in enumerate(students, start=1):
        row_index = header_row_index + idx
        score_value = score_values[idx - 1] if idx - 1 < len(score_values) else 0.0
        if _is_absent(score_value) or normalize(score_value) == normalize(CO_REPORT_NOT_APPLICABLE_TOKEN):
            normalized_score: float | str = CO_REPORT_NOT_APPLICABLE_TOKEN
        elif isinstance(score_value, (int, float)):
            normalized_score = _round2(float(score_value))
        else:
            normalized_score = 0.0
        row_values: list[Any] = [idx, reg_no, student_name, normalized_score]
        for col, value in enumerate(row_values, start=0):
            if col == 2:
                fmt = formats["body_wrap"]
            elif col >= 3:
                fmt = formats["body_center"]
            else:
                fmt = formats["body"]
            ws.write(row_index, col, value, fmt)

    _apply_layout(ws, header_row_index=header_row_index, paper_size=9, landscape=True)


def _write_indirect_sheet(
    workbook: Any,
    *,
    template_id: str,
    sheet_name: str,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    components: list[Any],
    formats: dict[str, Any],
) -> None:
    """Write indirect sheet.
    
    Args:
        workbook: Parameter value (Any).
        template_id: Parameter value (str).
        sheet_name: Parameter value (str).
        co_index: Parameter value (int).
        metadata_rows: Parameter value (list[tuple[str, Any]]).
        students: Parameter value (list[tuple[str, str]]).
        components: Parameter value (list[Any]).
        formats: Parameter value (dict[str, Any]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Ratio total header.
    
    Args:
        ratio: Parameter value (float).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    percent = ratio * 100.0
    if abs(percent - round(percent)) <= 1e-9:
        token = f"{int(round(percent))}"
    else:
        token = f"{percent:g}"
    return CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE.format(ratio=token)


def ratio_total_header(ratio: float) -> str:
    """Ratio total header.
    
    Args:
        ratio: Parameter value (float).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    return _ratio_total_header(ratio)


def _round2(value: float) -> float:
    """Round2.
    
    Args:
        value: Parameter value (float).
    
    Returns:
        float: Return value.
    
    Raises:
        None.
    """
    return round(float(value), CO_REPORT_MAX_DECIMAL_PLACES)


def _is_absent(value: Any) -> bool:
    """Is absent.
    
    Args:
        value: Parameter value (Any).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
    return normalize(value) == normalize(CO_REPORT_ABSENT_TOKEN)


def _template_supports_component_layout(template_id: str) -> bool:
    """Template supports component layout.

    Args:
        template_id: Parameter value (str).

    Returns:
        bool: Return value.

    Raises:
        None.
    """
    return normalize(template_id) == normalize("COURSE_SETUP_V2")


def _resolve_joined_total_rows_for_template(
    *,
    template_id: str,
    rows: list[_JoinedCoAnalysisRowLike],
) -> list[tuple[str, str, float | str, float | str]]:
    """Resolve joined total rows for template.
    
    Args:
        template_id: Parameter value (str).
        rows: Parameter value (list[_JoinedCoAnalysisRowLike]).
    
    Returns:
        list[tuple[str, str, float | str, float | str]]: Return value.
    
    Raises:
        None.
    """
    if normalize(template_id) == normalize("COURSE_SETUP_V2"):
        return [
            (
                str(row.reg_no).strip(),
                str(row.student_name).strip(),
                row.direct_score,
                row.indirect_score,
            )
            for row in rows
        ]
    raise validation_error_from_key(
        "validation.template.unknown",
        code="UNKNOWN_TEMPLATE",
        template_id=template_id,
    )
