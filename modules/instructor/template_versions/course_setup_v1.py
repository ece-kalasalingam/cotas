"""Version-specific handlers for COURSE_SETUP_V1."""

from __future__ import annotations

import re
from typing import Any, Sequence

from common.constants import (
    ASSESSMENT_CONFIG_HEADERS,
    ASSESSMENT_CONFIG_SHEET,
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SHEET,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    QUESTION_MAP_HEADERS,
    QUESTION_MAP_SHEET,
    STUDENTS_HEADERS,
    STUDENTS_SHEET,
    WEIGHT_TOTAL_EXPECTED,
    WEIGHT_TOTAL_ROUND_DIGITS,
)
from common.exceptions import ValidationError
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.texts import t
from common.utils import coerce_excel_number, normalize


def extract_marks_template_context(workbook: Any) -> dict[str, Any]:
    from modules.instructor import instructor_template_engine as generator

    return generator._extract_marks_template_context(workbook)


def write_marks_template_workbook(
    workbook: Any,
    context: dict[str, Any],
    *,
    template_id: str,
    cancel_token: Any = None,
) -> dict[str, Any]:
    from modules.instructor import instructor_template_engine as generator

    return generator._write_marks_template_workbook(
        workbook,
        context,
        template_id=template_id,
        cancel_token=cancel_token,
    )


def validate_course_details_rules(workbook: Any) -> None:
    metadata_sheet = workbook[COURSE_METADATA_SHEET]
    assessment_sheet = workbook[ASSESSMENT_CONFIG_SHEET]
    question_map_sheet = workbook[QUESTION_MAP_SHEET]
    students_sheet = workbook[STUDENTS_SHEET]

    total_outcomes = _validate_course_metadata(metadata_sheet)
    component_config = _validate_assessment_config(assessment_sheet)
    _validate_question_map(question_map_sheet, component_config, total_outcomes)
    _validate_students(students_sheet)


def validate_filled_marks_manifest_schema(workbook: Any, manifest: Any) -> None:
    if not isinstance(manifest, dict):
        raise ValidationError(t("instructor.validation.step3.manifest_root_invalid"))

    sheet_order = manifest.get("sheet_order")
    sheet_specs = manifest.get("sheets")
    if not isinstance(sheet_order, list) or not isinstance(sheet_specs, list):
        raise ValidationError(t("instructor.validation.step3.manifest_structure_invalid"))

    if list(workbook.sheetnames) != sheet_order:
        raise ValidationError(
            t(
                "instructor.validation.step3.sheet_order_mismatch",
                expected=sheet_order,
                found=list(workbook.sheetnames),
            )
        )

    has_marks_component = False
    for spec in sheet_specs:
        if not isinstance(spec, dict):
            raise ValidationError(t("instructor.validation.step3.manifest_sheet_spec_invalid"))
        sheet_name = spec.get("name")
        header_row = spec.get("header_row")
        headers = spec.get("headers")
        anchors = spec.get("anchors", [])
        formula_anchors = spec.get("formula_anchors", [])
        if not isinstance(sheet_name, str) or sheet_name not in workbook.sheetnames:
            raise ValidationError(
                t("instructor.validation.step3.sheet_missing", sheet_name=sheet_name)
            )
        if not isinstance(header_row, int) or header_row <= 0:
            raise ValidationError(
                t(
                    "instructor.validation.step3.header_row_invalid",
                    sheet_name=sheet_name,
                    header_row=header_row,
                )
            )
        if not isinstance(headers, list) or not headers:
            raise ValidationError(
                t("instructor.validation.step3.headers_missing", sheet_name=sheet_name)
            )
        if not isinstance(anchors, list):
            raise ValidationError(
                t("instructor.validation.step3.anchor_spec_invalid", sheet_name=sheet_name)
            )
        if not isinstance(formula_anchors, list):
            raise ValidationError(
                t("instructor.validation.step3.formula_anchor_spec_invalid", sheet_name=sheet_name)
            )

        worksheet = workbook[sheet_name]
        expected_headers = [normalize(value) for value in headers]
        actual_headers = [
            normalize(worksheet.cell(row=header_row, column=col_index + 1).value)
            for col_index in range(len(expected_headers))
        ]
        if actual_headers != expected_headers:
            raise ValidationError(
                t(
                    "instructor.validation.step3.header_row_mismatch",
                    sheet_name=sheet_name,
                    row=header_row,
                    expected=headers,
                )
            )

        for anchor in anchors:
            if not isinstance(anchor, list) or len(anchor) != 2:
                raise ValidationError(
                    t("instructor.validation.step3.anchor_spec_invalid", sheet_name=sheet_name)
                )
            cell_ref, expected_value = anchor
            if not isinstance(cell_ref, str) or not cell_ref:
                raise ValidationError(
                    t("instructor.validation.step3.anchor_spec_invalid", sheet_name=sheet_name)
                )
            actual_value = worksheet[cell_ref].value
            if not _filled_marks_values_match(expected_value, actual_value):
                raise ValidationError(
                    t(
                        "instructor.validation.step3.anchor_value_mismatch",
                        sheet_name=sheet_name,
                        cell=cell_ref,
                        expected=expected_value,
                        found=actual_value,
                    )
                )
        for formula_anchor in formula_anchors:
            if not isinstance(formula_anchor, list) or len(formula_anchor) != 2:
                raise ValidationError(
                    t("instructor.validation.step3.formula_anchor_spec_invalid", sheet_name=sheet_name)
                )
            cell_ref, expected_formula = formula_anchor
            if not isinstance(cell_ref, str) or not isinstance(expected_formula, str):
                raise ValidationError(
                    t("instructor.validation.step3.formula_anchor_spec_invalid", sheet_name=sheet_name)
                )
            actual_formula = worksheet[cell_ref].value
            if _normalized_formula(actual_formula) != _normalized_formula(expected_formula):
                raise ValidationError(
                    t(
                        "instructor.validation.step3.formula_mismatch",
                        sheet_name=sheet_name,
                        cell=cell_ref,
                    )
                )

        if sheet_name not in (COURSE_METADATA_SHEET, ASSESSMENT_CONFIG_SHEET):
            has_marks_component = True

    if not has_marks_component:
        raise ValidationError(t("instructor.validation.step3.no_component_sheets"))


def _iter_data_rows(worksheet: Any, expected_col_count: int) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for row in worksheet.iter_rows(min_row=2, max_col=expected_col_count, values_only=True):
        values = list(row)
        if any(normalize(value) != "" for value in values):
            rows.append(values)
    return rows


def _header_index_map(worksheet: Any, headers: Sequence[str]) -> dict[str, int]:
    index: dict[str, int] = {}
    for col_index, expected_header in enumerate(headers, start=1):
        header_value = worksheet.cell(row=1, column=col_index).value
        index[normalize(expected_header)] = col_index - 1
        if normalize(header_value) != normalize(expected_header):
            raise ValidationError(
                t(
                    "instructor.validation.unexpected_header",
                    sheet_name=worksheet.title,
                    col=col_index,
                )
            )
    return index


def _validate_course_metadata(worksheet: Any) -> int:
    expected_headers = list(COURSE_METADATA_HEADERS)
    header_map = _header_index_map(worksheet, expected_headers)
    field_header = normalize(expected_headers[0])
    value_header = normalize(expected_headers[1])
    rows = _iter_data_rows(worksheet, len(expected_headers))

    expected_field_rows = SAMPLE_SETUP_DATA.get(COURSE_METADATA_SHEET, [])
    expected_field_types: dict[str, type] = {}
    for field_name, sample_value in expected_field_rows:
        key = normalize(field_name)
        expected_field_types[key] = int if isinstance(sample_value, int) else str

    actual_values: dict[str, Any] = {}
    for row_number, row in enumerate(rows, start=2):
        field_raw = row[header_map[field_header]]
        value_raw = row[header_map[value_header]]
        field_key = normalize(field_raw)
        if not field_key:
            raise ValidationError(t("instructor.validation.course_metadata_field_empty", row=row_number))
        if field_key in actual_values:
            raise ValidationError(
                t(
                    "instructor.validation.course_metadata_duplicate_field",
                    row=row_number,
                    field=field_raw,
                )
            )
        if field_key not in expected_field_types:
            raise ValidationError(
                t(
                    "instructor.validation.course_metadata_unknown_field",
                    row=row_number,
                    field=field_raw,
                )
            )
        if normalize(value_raw) == "":
            raise ValidationError(
                t(
                    "instructor.validation.course_metadata_value_required",
                    row=row_number,
                    field=field_raw,
                )
            )
        actual_values[field_key] = coerce_excel_number(value_raw)

    missing_fields = [name for name in expected_field_types if name not in actual_values]
    if missing_fields:
        raise ValidationError(
            t(
                "instructor.validation.course_metadata_missing_fields",
                fields=", ".join(missing_fields),
            )
        )

    for field_key, expected_type in expected_field_types.items():
        value = actual_values[field_key]
        if expected_type is int:
            if isinstance(value, bool) or not isinstance(value, int):
                raise ValidationError(
                    t("instructor.validation.course_metadata_field_must_be_int", field=field_key)
                )
        else:
            if not isinstance(value, str) or normalize(value) == "":
                raise ValidationError(
                    t(
                        "instructor.validation.course_metadata_field_must_be_non_empty_str",
                        field=field_key,
                    )
                )

    total_outcomes = actual_values.get(normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY))
    if isinstance(total_outcomes, bool) or not isinstance(total_outcomes, int) or total_outcomes <= 0:
        raise ValidationError(t("instructor.validation.course_metadata_total_outcomes_invalid"))
    return total_outcomes


def _parse_yes_no(value: Any, sheet_name: str, row_number: int, field_name: str) -> bool:
    token = normalize(value)
    yes_no_tokens = {"yes", "no"}
    if token not in yes_no_tokens:
        raise ValidationError(
            t(
                "instructor.validation.yes_no_required",
                sheet_name=sheet_name,
                row=row_number,
                field=field_name,
            )
        )
    return token == "yes"


def _validate_assessment_config(worksheet: Any) -> dict[str, dict[str, Any]]:
    assessment_headers = ASSESSMENT_CONFIG_HEADERS
    expected_headers = list(assessment_headers)
    header_map = _header_index_map(worksheet, expected_headers)
    component_header = normalize(assessment_headers[0])
    weight_header = normalize(assessment_headers[1])
    cia_header = normalize(assessment_headers[2])
    co_wise_header = normalize(assessment_headers[3])
    direct_header = normalize(assessment_headers[4])
    rows = _iter_data_rows(worksheet, len(expected_headers))

    if not rows:
        raise ValidationError(t("instructor.validation.assessment_component_required_one"))

    component_config: dict[str, dict[str, Any]] = {}
    direct_weight_total = 0.0
    indirect_weight_total = 0.0
    direct_count = 0
    indirect_count = 0

    for row_number, row in enumerate(rows, start=2):
        component_raw = row[header_map[component_header]]
        component_key = normalize(component_raw)
        if not component_key:
            raise ValidationError(
                t("instructor.validation.assessment_component_required", row=row_number)
            )
        if component_key in component_config:
            raise ValidationError(
                t(
                    "instructor.validation.assessment_component_duplicate",
                    row=row_number,
                    component=component_raw,
                )
            )

        weight_value = coerce_excel_number(row[header_map[weight_header]])
        if (
            isinstance(weight_value, bool)
            or not isinstance(weight_value, (int, float))
        ):
            raise ValidationError(
                t("instructor.validation.assessment_weight_numeric", row=row_number)
            )

        is_direct = _parse_yes_no(
            row[header_map[direct_header]],
            ASSESSMENT_CONFIG_SHEET,
            row_number,
            assessment_headers[4],
        )
        co_wise_breakup = _parse_yes_no(
            row[header_map[co_wise_header]],
            ASSESSMENT_CONFIG_SHEET,
            row_number,
            assessment_headers[3],
        )
        _parse_yes_no(
            row[header_map[cia_header]],
            ASSESSMENT_CONFIG_SHEET,
            row_number,
            assessment_headers[2],
        )

        if is_direct:
            direct_weight_total += float(weight_value)
            direct_count += 1
        else:
            indirect_weight_total += float(weight_value)
            indirect_count += 1

        component_config[component_key] = {
            "display_name": str(component_raw).strip(),
            "co_wise_breakup": co_wise_breakup,
        }

    if direct_count == 0:
        raise ValidationError(t("instructor.validation.assessment_direct_missing"))
    if indirect_count == 0:
        raise ValidationError(t("instructor.validation.assessment_indirect_missing"))
    if round(direct_weight_total, WEIGHT_TOTAL_ROUND_DIGITS) != WEIGHT_TOTAL_EXPECTED:
        raise ValidationError(
            t("instructor.validation.assessment_direct_total_invalid", found=direct_weight_total)
        )
    if round(indirect_weight_total, WEIGHT_TOTAL_ROUND_DIGITS) != WEIGHT_TOTAL_EXPECTED:
        raise ValidationError(
            t("instructor.validation.assessment_indirect_total_invalid", found=indirect_weight_total)
        )

    return component_config


def _co_tokens(value: Any) -> list[int]:
    if value is None:
        return []
    if isinstance(value, bool):
        return []
    if isinstance(value, int):
        return [value]
    if isinstance(value, float):
        return [int(value)] if value.is_integer() else []

    token = str(value).strip()
    if not token:
        return []

    numbers: list[int] = []
    for item in token.split(","):
        part = item.strip()
        if not part:
            return []
        match = re.fullmatch(r"(?:co)?\s*(\d+)", part, flags=re.IGNORECASE)
        if not match:
            return []
        numbers.append(int(match.group(1)))
    return numbers


def _validate_question_map(
    worksheet: Any,
    component_config: dict[str, dict[str, Any]],
    total_outcomes: int,
) -> None:
    expected_headers = list(QUESTION_MAP_HEADERS)
    header_map = _header_index_map(worksheet, expected_headers)
    component_header = normalize(expected_headers[0])
    question_header = normalize(expected_headers[1])
    max_marks_header = normalize(expected_headers[2])
    co_header = normalize(expected_headers[3])
    rows = _iter_data_rows(worksheet, len(expected_headers))

    if not rows:
        raise ValidationError(t("instructor.validation.question_map_row_required_one"))

    seen_co_wise_questions: set[tuple[str, str]] = set()
    for row_number, row in enumerate(rows, start=2):
        component_raw = row[header_map[component_header]]
        component_key = normalize(component_raw)
        if not component_key:
            raise ValidationError(t("instructor.validation.question_component_required", row=row_number))
        if component_key not in component_config:
            raise ValidationError(
                t(
                    "instructor.validation.question_component_unknown",
                    row=row_number,
                    component=component_raw,
                )
            )

        question_raw = row[header_map[question_header]]
        question_key = normalize(question_raw)
        if not question_key:
            raise ValidationError(
                t("instructor.validation.question_label_required", row=row_number)
            )

        max_marks = coerce_excel_number(row[header_map[max_marks_header]])
        if isinstance(max_marks, bool) or not isinstance(max_marks, (int, float)):
            raise ValidationError(t("instructor.validation.question_max_marks_numeric", row=row_number))
        if float(max_marks) <= 0:
            raise ValidationError(
                t("instructor.validation.question_max_marks_positive", row=row_number)
            )

        co_values = _co_tokens(row[header_map[co_header]])
        if not co_values:
            raise ValidationError(t("instructor.validation.question_co_required", row=row_number))
        if len(set(co_values)) != len(co_values):
            raise ValidationError(t("instructor.validation.question_co_no_repeat", row=row_number))
        if any(co_number <= 0 or co_number > total_outcomes for co_number in co_values):
            raise ValidationError(
                t(
                    "instructor.validation.question_co_out_of_range",
                    row=row_number,
                    total_outcomes=total_outcomes,
                )
            )

        is_co_wise = bool(component_config[component_key]["co_wise_breakup"])
        if is_co_wise:
            if len(co_values) != 1:
                raise ValidationError(
                    t(
                        "instructor.validation.question_co_wise_requires_one",
                        row=row_number,
                        component=component_raw,
                    )
                )
            question_id = (component_key, question_key)
            if question_id in seen_co_wise_questions:
                raise ValidationError(
                    t(
                        "instructor.validation.question_duplicate_for_component",
                        row=row_number,
                        question=question_raw,
                        component=component_raw,
                    )
                )
            seen_co_wise_questions.add(question_id)


def _validate_students(worksheet: Any) -> None:
    expected_headers = list(STUDENTS_HEADERS)
    header_map = _header_index_map(worksheet, expected_headers)
    reg_no_header = normalize(expected_headers[0])
    student_name_header = normalize(expected_headers[1])
    rows = _iter_data_rows(worksheet, len(expected_headers))

    if not rows:
        raise ValidationError(t("instructor.validation.students_row_required_one"))

    seen_reg_numbers: set[str] = set()
    for row_number, row in enumerate(rows, start=2):
        reg_no_raw = row[header_map[reg_no_header]]
        student_name_raw = row[header_map[student_name_header]]

        reg_no = str(reg_no_raw).strip() if reg_no_raw is not None else ""
        student_name = str(student_name_raw).strip() if student_name_raw is not None else ""

        if not reg_no or not student_name:
            raise ValidationError(
                t("instructor.validation.students_reg_and_name_required", row=row_number)
            )

        reg_key = normalize(reg_no)
        if reg_key in seen_reg_numbers:
            raise ValidationError(
                t("instructor.validation.students_duplicate_reg_no", row=row_number, reg_no=reg_no)
            )
        seen_reg_numbers.add(reg_key)


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
