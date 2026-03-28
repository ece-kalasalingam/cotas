"""COURSE_SETUP_V2 filled-marks manifest validation and anomaly warning state."""

from __future__ import annotations

import logging
from typing import Any, Sequence

from common.constants import (
    LAYOUT_MANIFEST_KEY_SHEET_ORDER,
    LAYOUT_MANIFEST_KEY_SHEETS,
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
    MIN_MARK_VALUE,
)
from common.error_catalog import validation_error_from_key
from common.utils import coerce_excel_number, normalize
from common.workbook_integrity.workbook_signing import sign_payload

_logger = logging.getLogger(__name__)
_MAX_DECIMAL_PLACES = 2
_FORMULA_SUM_TEMPLATE = "=SUM({start}:{end})"
_LOG_STEP3_HIGH_ABSENCE = "Step3 anomaly: high absence ratio sheet=%s col=%s absent=%s total=%s"
_LOG_STEP3_NEAR_CONSTANT = (
    "Step3 anomaly: near-constant marks sheet=%s col=%s dominant_count=%s numeric_total=%s"
)
_last_marks_anomaly_warnings: list[str] = []
_MARK_COMPONENT_SHEET_KINDS = {
    LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
    LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
    LAYOUT_SHEET_KIND_INDIRECT,
}


def _reset_marks_anomaly_warnings() -> None:
    _last_marks_anomaly_warnings.clear()


def consume_last_marks_anomaly_warnings() -> list[str]:
    warnings = list(_last_marks_anomaly_warnings)
    _last_marks_anomaly_warnings.clear()
    return warnings


def validate_filled_marks_manifest_schema(*, workbook: Any, manifest: Any) -> None:
    _reset_marks_anomaly_warnings()
    if not isinstance(manifest, dict):
        raise validation_error_from_key("instructor.validation.step2.manifest_root_invalid")

    sheet_order = manifest.get(LAYOUT_MANIFEST_KEY_SHEET_ORDER)
    sheet_specs = manifest.get(LAYOUT_MANIFEST_KEY_SHEETS)
    if not isinstance(sheet_order, list) or not isinstance(sheet_specs, list):
        raise validation_error_from_key("instructor.validation.step2.manifest_structure_invalid")

    if list(workbook.sheetnames) != sheet_order:
        raise validation_error_from_key(
            "instructor.validation.step2.sheet_order_mismatch",
            expected=sheet_order,
            found=list(workbook.sheetnames),
        )

    has_marks_component = False
    baseline_student_hash: str | None = None
    baseline_student_sheet: str | None = None
    for spec in sheet_specs:
        if not isinstance(spec, dict):
            raise validation_error_from_key("instructor.validation.step2.manifest_sheet_spec_invalid")
        sheet_name = spec.get(LAYOUT_SHEET_SPEC_KEY_NAME)
        header_row = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADER_ROW)
        headers = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADERS)
        anchors = spec.get(LAYOUT_SHEET_SPEC_KEY_ANCHORS, [])
        formula_anchors = spec.get(LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS, [])
        if not isinstance(sheet_name, str) or sheet_name not in workbook.sheetnames:
            raise validation_error_from_key(
                "instructor.validation.step2.sheet_missing", sheet_name=sheet_name
            )
        if not isinstance(header_row, int) or header_row <= 0:
            raise validation_error_from_key(
                "instructor.validation.step2.header_row_invalid",
                sheet_name=sheet_name,
                header_row=header_row,
            )
        if not isinstance(headers, list) or not headers:
            raise validation_error_from_key(
                "instructor.validation.step2.headers_missing", sheet_name=sheet_name
            )
        if not isinstance(anchors, list):
            raise validation_error_from_key(
                "instructor.validation.step2.anchor_spec_invalid", sheet_name=sheet_name
            )
        if not isinstance(formula_anchors, list):
            raise validation_error_from_key(
                "instructor.validation.step2.formula_anchor_spec_invalid", sheet_name=sheet_name
            )

        worksheet = workbook[sheet_name]
        expected_headers = [normalize(value) for value in headers]
        actual_headers = [
            normalize(worksheet.cell(row=header_row, column=col_index + 1).value)
            for col_index in range(len(expected_headers))
        ]
        if actual_headers != expected_headers:
            raise validation_error_from_key(
                "instructor.validation.step2.header_row_mismatch",
                sheet_name=sheet_name,
                row=header_row,
                expected=headers,
            )

        for anchor in anchors:
            if not isinstance(anchor, list) or len(anchor) != 2:
                raise validation_error_from_key(
                    "instructor.validation.step2.anchor_spec_invalid", sheet_name=sheet_name
                )
            cell_ref, expected_value = anchor
            if not isinstance(cell_ref, str) or not cell_ref:
                raise validation_error_from_key(
                    "instructor.validation.step2.anchor_spec_invalid", sheet_name=sheet_name
                )
            actual_value = worksheet[cell_ref].value
            if not _filled_marks_values_match(expected_value, actual_value):
                raise validation_error_from_key(
                    "instructor.validation.step2.anchor_value_mismatch",
                    sheet_name=sheet_name,
                    cell=cell_ref,
                    expected=expected_value,
                    found=actual_value,
                )
        for formula_anchor in formula_anchors:
            if not isinstance(formula_anchor, list) or len(formula_anchor) != 2:
                raise validation_error_from_key(
                    "instructor.validation.step2.formula_anchor_spec_invalid", sheet_name=sheet_name
                )
            cell_ref, expected_formula = formula_anchor
            if not isinstance(cell_ref, str) or not isinstance(expected_formula, str):
                raise validation_error_from_key(
                    "instructor.validation.step2.formula_anchor_spec_invalid", sheet_name=sheet_name
                )
            actual_formula = worksheet[cell_ref].value
            if _normalized_formula(actual_formula) != _normalized_formula(expected_formula):
                raise validation_error_from_key(
                    "instructor.validation.step2.formula_mismatch",
                    sheet_name=sheet_name,
                    cell=cell_ref,
                )

        sheet_kind = spec.get(LAYOUT_SHEET_SPEC_KEY_KIND)
        is_mark_component = sheet_kind in _MARK_COMPONENT_SHEET_KINDS
        if is_mark_component:
            has_marks_component = True
            _validate_component_structure_snapshot(
                worksheet=worksheet,
                sheet_name=sheet_name,
                sheet_kind=sheet_kind,
                header_row=header_row,
                structure=spec.get(LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE),
                header_count=len(expected_headers),
            )
            actual_student_hash = _validate_component_student_identity(
                worksheet=worksheet,
                sheet_name=sheet_name,
                sheet_kind=sheet_kind,
                header_row=header_row,
                expected_student_count=spec.get(LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT),
                expected_student_hash=spec.get(LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH),
            )
            if baseline_student_hash is None:
                baseline_student_hash = actual_student_hash
                baseline_student_sheet = sheet_name
            elif actual_student_hash != baseline_student_hash:
                raise validation_error_from_key(
                    "instructor.validation.step2.student_identity_cross_sheet_mismatch",
                    sheet_name=sheet_name,
                    reference_sheet=baseline_student_sheet,
                )
            _validate_non_empty_marks_entries(
                worksheet=worksheet,
                sheet_name=sheet_name,
                sheet_kind=sheet_kind,
                header_count=len(expected_headers),
                header_row=header_row,
            )

    if not has_marks_component:
        raise validation_error_from_key("instructor.validation.step2.no_component_sheets")


def _filled_marks_values_match(expected_value: object, actual_value: object) -> bool:
    expected_coerced = coerce_excel_number(expected_value)
    actual_coerced = coerce_excel_number(actual_value)
    numeric_types = (int, float)
    if isinstance(expected_coerced, numeric_types) and not isinstance(expected_coerced, bool):
        if not isinstance(actual_coerced, numeric_types) or isinstance(actual_coerced, bool):
            return False
        return abs(float(expected_coerced) - float(actual_coerced)) <= 1e-9

    return normalize(expected_coerced) == normalize(actual_coerced)


def _normalized_formula(value: object) -> str:
    token = normalize(value)
    token = token.replace("$", "")
    token = token.replace(" ", "")
    return token


def _validate_component_student_identity(
    *,
    worksheet: Any,
    sheet_name: str,
    sheet_kind: Any,
    header_row: int,
    expected_student_count: Any,
    expected_student_hash: Any,
) -> str:
    if not isinstance(expected_student_count, int) or expected_student_count < 0:
        raise validation_error_from_key(
            "instructor.validation.step2.student_identity_spec_invalid", sheet_name=sheet_name
        )
    if not isinstance(expected_student_hash, str) or not expected_student_hash.strip():
        raise validation_error_from_key(
            "instructor.validation.step2.student_identity_spec_invalid", sheet_name=sheet_name
        )

    students = _extract_component_students(
        worksheet=worksheet,
        sheet_name=sheet_name,
        sheet_kind=sheet_kind,
        header_row=header_row,
    )
    if len(students) != expected_student_count:
        raise validation_error_from_key(
            "instructor.validation.step2.student_identity_mismatch",
            sheet_name=sheet_name,
        )

    actual_hash = _student_identity_hash(students)
    if actual_hash != expected_student_hash:
        raise validation_error_from_key(
            "instructor.validation.step2.student_identity_mismatch",
            sheet_name=sheet_name,
        )
    return actual_hash


def _extract_component_students(
    *,
    worksheet: Any,
    sheet_name: str,
    sheet_kind: Any,
    header_row: int,
) -> list[tuple[str, str]]:
    first_row = _marks_data_start_row(sheet_kind, header_row)
    students: list[tuple[str, str]] = []
    seen_reg_numbers: set[str] = set()
    row = first_row
    while True:
        reg_value = worksheet.cell(row=row, column=2).value
        name_value = worksheet.cell(row=row, column=3).value
        reg_no = str(reg_value).strip() if reg_value is not None else ""
        student_name = str(name_value).strip() if name_value is not None else ""
        if not reg_no and not student_name:
            break
        if not reg_no or not student_name:
            raise validation_error_from_key(
                "instructor.validation.step2.student_identity_mismatch",
                sheet_name=sheet_name,
            )
        reg_key = normalize(reg_no)
        if reg_key in seen_reg_numbers:
            raise validation_error_from_key(
                "instructor.validation.step2.student_reg_duplicate",
                sheet_name=sheet_name,
                reg_no=reg_no,
            )
        seen_reg_numbers.add(reg_key)
        students.append((reg_no, student_name))
        row += 1
    return students


def _student_identity_hash(students: Sequence[tuple[str, str]]) -> str:
    payload = "\n".join(f"{reg_no.strip()}|{student_name.strip()}" for reg_no, student_name in students)
    return sign_payload(payload)


def _validate_non_empty_marks_entries(
    *,
    worksheet: Any,
    sheet_name: str,
    sheet_kind: Any,
    header_count: int,
    header_row: int,
) -> None:
    student_count = _infer_student_count(worksheet=worksheet, sheet_kind=sheet_kind, header_row=header_row)
    if student_count <= 0:
        return

    data_start_row = _marks_data_start_row(sheet_kind, header_row)
    mark_cols = _marks_entry_columns(sheet_kind, header_count)
    absent_count_by_col: dict[int, int] = {col: 0 for col in mark_cols}
    numeric_count_by_col: dict[int, int] = {col: 0 for col in mark_cols}
    frequency_by_value_by_col: dict[int, dict[float, int]] = {col: {} for col in mark_cols}
    max_row = header_row + 2
    minimum = _mark_min_for_sheet(sheet_kind)
    maximum_by_col = {
        col: _mark_max_for_cell(worksheet, sheet_kind, max_row, col)
        for col in mark_cols
    }
    for row in range(data_start_row, data_start_row + student_count):
        has_absent = False
        has_numeric = False
        for col in mark_cols:
            cell = worksheet.cell(row=row, column=col)
            cell_value = cell.value
            token = normalize(cell_value)
            if token == "":
                raise validation_error_from_key(
                    "instructor.validation.step2.mark_entry_empty",
                    code="COA_MARK_ENTRY_EMPTY",
                    sheet_name=sheet_name,
                    cell=cell.coordinate,
                )
            if token == "a":
                has_absent = True
                absent_count_by_col[col] += 1
                continue
            has_numeric = True
            numeric_value = coerce_excel_number(cell_value)
            if isinstance(numeric_value, bool) or not isinstance(numeric_value, (int, float)):
                raise validation_error_from_key(
                    "instructor.validation.step2.mark_value_invalid",
                    code="COA_MARK_VALUE_INVALID",
                    sheet_name=sheet_name,
                    cell=cell.coordinate,
                    value=cell_value,
                    minimum=minimum,
                    maximum=maximum_by_col[col],
                )
            if not _has_allowed_decimal_precision(float(numeric_value)):
                raise validation_error_from_key(
                    "instructor.validation.step2.mark_precision_invalid",
                    code="COA_MARK_PRECISION_INVALID",
                    sheet_name=sheet_name,
                    cell=cell.coordinate,
                    value=cell_value,
                    decimals=_MAX_DECIMAL_PLACES,
                )
            if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT and not _is_integer_value(float(numeric_value)):
                raise validation_error_from_key(
                    "instructor.validation.step2.indirect_mark_must_be_integer",
                    code="COA_INDIRECT_MARK_INTEGER_REQUIRED",
                    sheet_name=sheet_name,
                    cell=cell.coordinate,
                    value=cell_value,
                )
            maximum = maximum_by_col[col]
            numeric_float = float(numeric_value)
            if numeric_float < minimum or numeric_float > maximum:
                raise validation_error_from_key(
                    "instructor.validation.step2.mark_value_invalid",
                    code="COA_MARK_VALUE_INVALID",
                    sheet_name=sheet_name,
                    cell=cell.coordinate,
                    value=cell_value,
                    minimum=minimum,
                    maximum=maximum,
                )
            numeric_count_by_col[col] += 1
            frequency_by_value = frequency_by_value_by_col[col]
            frequency_by_value[numeric_float] = frequency_by_value.get(numeric_float, 0) + 1
        _validate_absence_policy_for_row(
            sheet_name=sheet_name,
            worksheet=worksheet,
            sheet_kind=sheet_kind,
            row=row,
            mark_cols=mark_cols,
            has_absent=has_absent,
            has_numeric=has_numeric,
        )
    _log_marks_anomaly_warnings_from_stats(
        sheet_name=sheet_name,
        mark_cols=mark_cols,
        student_count=student_count,
        absent_count_by_col=absent_count_by_col,
        numeric_count_by_col=numeric_count_by_col,
        frequency_by_value_by_col=frequency_by_value_by_col,
    )
    _validate_row_total_consistency(
        worksheet=worksheet,
        sheet_name=sheet_name,
        sheet_kind=sheet_kind,
        header_count=header_count,
        header_row=header_row,
        student_count=student_count,
    )


def _marks_data_start_row(sheet_kind: Any, header_row: int) -> int:
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        return header_row + 1
    return header_row + 3


def _marks_entry_columns(sheet_kind: Any, header_count: int) -> range:
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        return range(4, 5)
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
        return range(4, header_count)
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        return range(4, header_count + 1)
    raise validation_error_from_key("instructor.validation.step2.manifest_sheet_spec_invalid")


def _mark_min_for_sheet(sheet_kind: Any) -> float:
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        return float(max(MIN_MARK_VALUE, LIKERT_MIN))
    return float(MIN_MARK_VALUE)


def _mark_max_for_cell(worksheet: Any, sheet_kind: Any, max_row: int, col: int) -> float:
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        return float(LIKERT_MAX)
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        max_value = coerce_excel_number(worksheet.cell(row=max_row, column=4).value)
    elif sheet_kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
        max_value = coerce_excel_number(worksheet.cell(row=max_row, column=col).value)
    else:
        raise validation_error_from_key("instructor.validation.step2.manifest_sheet_spec_invalid")
    if isinstance(max_value, bool) or not isinstance(max_value, (int, float)):
        raise validation_error_from_key("instructor.validation.step2.manifest_sheet_spec_invalid")
    return float(max_value)


def _infer_student_count(*, worksheet: Any, sheet_kind: Any, header_row: int) -> int:
    first_row = _marks_data_start_row(sheet_kind, header_row)
    count = 0
    row = first_row
    while True:
        reg_no = worksheet.cell(row=row, column=2).value
        student_name = worksheet.cell(row=row, column=3).value
        if normalize(reg_no) == "" and normalize(student_name) == "":
            break
        count += 1
        row += 1
    return count


def _validate_absence_policy_for_row(
    *,
    sheet_name: str,
    worksheet: Any,
    sheet_kind: Any,
    row: int,
    mark_cols: range,
    has_absent: bool,
    has_numeric: bool,
) -> None:
    if has_absent and has_numeric:
        mark_range = (
            f"{worksheet.cell(row=row, column=mark_cols.start).coordinate}:"
            f"{worksheet.cell(row=row, column=mark_cols.stop - 1).coordinate}"
        )
        raise validation_error_from_key(
            "instructor.validation.step2.absence_policy_violation",
            code="COA_ABSENCE_POLICY_VIOLATION",
            sheet_name=sheet_name,
            row=row,
            range=mark_range,
        )
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        return


def _validate_row_total_consistency(
    *,
    worksheet: Any,
    sheet_name: str,
    sheet_kind: Any,
    header_count: int,
    header_row: int,
    student_count: int,
) -> None:
    first_row = _marks_data_start_row(sheet_kind, header_row)
    last_row = first_row + student_count - 1
    if last_row < first_row:
        return

    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
        total_col = header_count
        first_mark_col = 4
        last_mark_col = header_count - 1
        for row in range(first_row, last_row + 1):
            actual = worksheet.cell(row=row, column=total_col).value
            expected = _FORMULA_SUM_TEMPLATE.format(
                start=f"{_excel_col_name(first_mark_col)}{row}",
                end=f"{_excel_col_name(last_mark_col)}{row}",
            )
            if _normalized_formula(actual) != _normalized_formula(expected):
                raise validation_error_from_key(
                    "instructor.validation.step2.total_formula_mismatch",
                    sheet_name=sheet_name,
                    cell=worksheet.cell(row=row, column=total_col).coordinate,
                )
        return

    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        for row in range(first_row, last_row + 1):
            for col in range(5, header_count + 1):
                formula = worksheet.cell(row=row, column=col).value
                if not isinstance(formula, str) or not formula.startswith("="):
                    raise validation_error_from_key(
                        "instructor.validation.step2.co_formula_mismatch",
                        sheet_name=sheet_name,
                        cell=worksheet.cell(row=row, column=col).coordinate,
                    )
        return


def _validate_component_structure_snapshot(
    *,
    worksheet: Any,
    sheet_name: str,
    sheet_kind: Any,
    header_row: int,
    structure: Any,
    header_count: int,
) -> None:
    if not isinstance(structure, dict):
        raise validation_error_from_key(
            "instructor.validation.step2.structure_snapshot_missing", sheet_name=sheet_name
        )
    max_row = header_row + 2
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
        maxima = structure.get("mark_maxima")
        if not isinstance(maxima, list):
            raise validation_error_from_key(
                "instructor.validation.step2.structure_snapshot_missing", sheet_name=sheet_name
            )
        for idx, expected in enumerate(maxima, start=4):
            actual = coerce_excel_number(worksheet.cell(row=max_row, column=idx).value)
            if not _filled_marks_values_match(expected, actual):
                raise validation_error_from_key(
                    "instructor.validation.step2.structure_snapshot_mismatch",
                    sheet_name=sheet_name,
                    cell=worksheet.cell(row=max_row, column=idx).coordinate,
                )
        return
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        maxima = structure.get("mark_maxima")
        if not isinstance(maxima, list):
            raise validation_error_from_key(
                "instructor.validation.step2.structure_snapshot_missing", sheet_name=sheet_name
            )
        for idx, expected in enumerate(maxima, start=4):
            actual = coerce_excel_number(worksheet.cell(row=max_row, column=idx).value)
            if not _filled_marks_values_match(expected, actual):
                raise validation_error_from_key(
                    "instructor.validation.step2.structure_snapshot_mismatch",
                    sheet_name=sheet_name,
                    cell=worksheet.cell(row=max_row, column=idx).coordinate,
                )
        return
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        likert_range = structure.get("likert_range")
        if likert_range != [LIKERT_MIN, LIKERT_MAX]:
            raise validation_error_from_key(
                "instructor.validation.step2.structure_snapshot_missing", sheet_name=sheet_name
            )
        return
    raise validation_error_from_key("instructor.validation.step2.manifest_sheet_spec_invalid")


def _has_allowed_decimal_precision(value: float) -> bool:
    scaled = round(value * (10**_MAX_DECIMAL_PLACES))
    return abs(value - (scaled / (10**_MAX_DECIMAL_PLACES))) <= 1e-9


def _is_integer_value(value: float) -> bool:
    return abs(value - round(value)) <= 1e-9


def _excel_col_name(col_index_1_based: int) -> str:
    index = col_index_1_based
    label = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        label = chr(65 + rem) + label
    return label


def _log_marks_anomaly_warnings_from_stats(
    *,
    sheet_name: str,
    mark_cols: range,
    student_count: int,
    absent_count_by_col: dict[int, int],
    numeric_count_by_col: dict[int, int],
    frequency_by_value_by_col: dict[int, dict[float, int]],
) -> None:
    for col in mark_cols:
        absent_count = int(absent_count_by_col.get(col, 0))
        if student_count > 0 and (absent_count / student_count) >= 0.4:
            warning_message = (
                f"High absence ratio in {sheet_name} column {_excel_col_name(col)} "
                f"({absent_count}/{student_count})."
            )
            _logger.warning(_LOG_STEP3_HIGH_ABSENCE, sheet_name, col, absent_count, student_count)
            _last_marks_anomaly_warnings.append(warning_message)

        numeric_total = int(numeric_count_by_col.get(col, 0))
        if numeric_total <= 0:
            continue
        value_frequency = frequency_by_value_by_col.get(col, {})
        dominant_count = max((int(count) for count in value_frequency.values()), default=0)
        if numeric_total >= 5 and (dominant_count / numeric_total) >= 0.9:
            warning_message = (
                f"Near-constant marks in {sheet_name} column {_excel_col_name(col)} "
                f"({dominant_count}/{numeric_total} same value)."
            )
            _logger.warning(_LOG_STEP3_NEAR_CONSTANT, sheet_name, col, dominant_count, numeric_total)
            _last_marks_anomaly_warnings.append(warning_message)


__all__ = [
    "consume_last_marks_anomaly_warnings",
    "validate_filled_marks_manifest_schema",
]
