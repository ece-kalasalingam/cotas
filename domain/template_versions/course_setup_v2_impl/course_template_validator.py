"""COURSE_SETUP_V2 course-template workbook validation."""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any

from common.error_catalog import resolve_validation_issue, validation_error_from_key
from common.exceptions import JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.registry import (
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SECTION_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS,
    COURSE_SETUP_ASSESSMENT_MODE_OPTIONS,
    COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS,
    COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS,
    COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS,
    COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    COURSE_SETUP_SHEET_KEY_QUESTION_MAP,
    COURSE_SETUP_SHEET_KEY_STUDENTS,
    WEIGHT_TOTAL_EXPECTED,
    WEIGHT_TOTAL_ROUND_DIGITS,
    get_blueprint,
    get_sheet_name_by_key,
    get_sheet_schema_by_key,
)
from common.sheet_schema import SheetSchema, ValidationRule
from common.utils import canonical_path_key, coerce_excel_number, normalize
from domain.co_token_parser import parse_co_tokens
from domain.template_strategy_router import read_valid_template_id_from_system_hash_sheet

_TEMPLATE_ID = "COURSE_SETUP_V2"
_COL_COMPONENT = "component"
_COL_WEIGHT_PERCENT = "weight_percent"
_COL_QUESTION_LABEL = "question_label"
_COL_MAX_MARKS = "max_marks"
_COL_CO = "co"
_COL_BLOOM_LEVEL = "bloom_level"
_COL_REG_NO = "reg_no"
_COL_STUDENT_NAME = "student_name"


class _CourseIdentity:
    def __init__(
        self,
        *,
        template_id: str,
        course_code: str,
        semester: str,
        academic_year: str,
        total_outcomes: int,
        section: str,
    ) -> None:
        self.template_id = template_id
        self.course_code = course_code
        self.semester = semester
        self.academic_year = academic_year
        self.total_outcomes = total_outcomes
        self.section = section

    def cohort_key(self) -> tuple[str, str, str, int]:
        return (
            normalize(self.course_code),
            normalize(self.semester),
            normalize(self.academic_year),
            int(self.total_outcomes),
        )


def _issue_dict(*, code: str, context: dict[str, Any], fallback_message: str) -> dict[str, object]:
    resolved = resolve_validation_issue(code, context, fallback_message=fallback_message)
    return {
        "code": resolved.code,
        "category": resolved.category,
        "severity": resolved.severity,
        "translation_key": resolved.translation_key,
        "message": resolved.message,
        "context": dict(resolved.context),
    }


def validate_course_details_workbook(
    workbook_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> str:
    identity = _validate_course_details_workbook_impl(
        workbook_path=workbook_path,
        cancel_token=cancel_token,
    )
    return identity.template_id


def validate_course_details_workbooks(
    workbook_paths: Sequence[str | Path],
    *,
    cancel_token: CancellationToken | None = None,
) -> dict[str, object]:
    unique_paths: list[str] = []
    seen_path_keys: set[str] = set()
    duplicate_paths: list[str] = []
    for raw in workbook_paths:
        path = str(raw).strip()
        if not path:
            continue
        key = canonical_path_key(path)
        if key in seen_path_keys:
            duplicate_paths.append(path)
            continue
        seen_path_keys.add(key)
        unique_paths.append(path)

    valid_paths: list[str] = []
    invalid_paths: list[str] = []
    mismatched_paths: list[str] = []
    duplicate_sections: list[str] = []
    template_ids: dict[str, str] = {}
    baseline_cohort: tuple[str, str, str, int] | None = None
    baseline_identity: _CourseIdentity | None = None
    seen_sections: set[str] = set()
    rejections: list[dict[str, object]] = []

    for path in duplicate_paths:
        issue = _issue_dict(
            code="COURSE_DETAILS_DUPLICATE_PATH",
            context={"workbook": path},
            fallback_message="Duplicate file path skipped.",
        )
        rejections.append(
            {
                "path": path,
                "reason_kind": "duplicate_path",
                "issue": issue,
            }
        )

    for path in unique_paths:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        try:
            identity = _validate_course_details_workbook_impl(
                workbook_path=path,
                cancel_token=cancel_token,
            )
        except JobCancelledError:
            raise
        except ValidationError as exc:
            invalid_paths.append(path)
            issue = _issue_dict(
                code=str(getattr(exc, "code", "VALIDATION_ERROR")),
                context=dict(getattr(exc, "context", {}) or {}),
                fallback_message=str(exc).strip() or "Validation failed.",
            )
            rejections.append(
                {
                    "path": path,
                    "reason_kind": "invalid",
                    "issue": issue,
                }
            )
            continue
        except Exception as exc:
            invalid_paths.append(path)
            issue = _issue_dict(
                code="COURSE_DETAILS_UNEXPECTED_REJECTION",
                context={"workbook": path},
                fallback_message=str(exc).strip() or "File skipped due to an unexpected validation failure.",
            )
            rejections.append(
                {
                    "path": path,
                    "reason_kind": "invalid",
                    "issue": issue,
                }
            )
            continue
        cohort = identity.cohort_key()
        if baseline_cohort is None:
            baseline_cohort = cohort
            baseline_identity = identity
        elif cohort != baseline_cohort:
            mismatched_paths.append(path)
            mismatch_fields: list[str] = []
            if baseline_identity is not None:
                if normalize(identity.course_code) != normalize(baseline_identity.course_code):
                    mismatch_fields.append(COURSE_METADATA_COURSE_CODE_KEY)
                if normalize(identity.semester) != normalize(baseline_identity.semester):
                    mismatch_fields.append(COURSE_METADATA_SEMESTER_KEY)
                if normalize(identity.academic_year) != normalize(baseline_identity.academic_year):
                    mismatch_fields.append(COURSE_METADATA_ACADEMIC_YEAR_KEY)
                if int(identity.total_outcomes) != int(baseline_identity.total_outcomes):
                    mismatch_fields.append(COURSE_METADATA_TOTAL_OUTCOMES_KEY)
            issue = _issue_dict(
                code="COURSE_DETAILS_COHORT_MISMATCH",
                context={
                    "workbook": path,
                    "fields": ", ".join(mismatch_fields) if mismatch_fields else "cohort",
                },
                fallback_message=(
                    "File skipped because course cohort metadata does not match "
                    "(course code, semester, academic year, total outcomes must match)."
                ),
            )
            rejections.append(
                {
                    "path": path,
                    "reason_kind": "cohort_mismatch",
                    "issue": issue,
                }
            )
            continue
        section_key = normalize(identity.section)
        if section_key in seen_sections:
            duplicate_sections.append(path)
            issue = _issue_dict(
                code="COURSE_DETAILS_SECTION_DUPLICATE",
                context={
                    "workbook": path,
                    "section": identity.section,
                },
                fallback_message="Duplicate section skipped for same course cohort.",
            )
            rejections.append(
                {
                    "path": path,
                    "reason_kind": "duplicate_section",
                    "issue": issue,
                }
            )
            continue
        seen_sections.add(section_key)
        valid_paths.append(path)
        template_ids[canonical_path_key(path)] = identity.template_id

    return {
        "valid_paths": valid_paths,
        "invalid_paths": invalid_paths,
        "mismatched_paths": mismatched_paths,
        "duplicate_paths": duplicate_paths,
        "duplicate_sections": duplicate_sections,
        "template_ids": template_ids,
        "rejections": rejections,
    }


def _validate_course_details_workbook_impl(
    *,
    workbook_path: str | Path,
    cancel_token: CancellationToken | None = None,
) -> _CourseIdentity:
    try:
        import openpyxl
    except ModuleNotFoundError as exc:
        raise validation_error_from_key(
            "validation.dependency.openpyxl_missing",
            code="OPENPYXL_MISSING",
        ) from exc

    workbook_file = Path(workbook_path)
    if not workbook_file.exists():
        raise validation_error_from_key(
            "validation.workbook.not_found",
            code="WORKBOOK_NOT_FOUND",
            workbook=str(workbook_file),
        )

    try:
        workbook = openpyxl.load_workbook(workbook_file, data_only=False, read_only=False)
    except Exception as exc:
        raise validation_error_from_key(
            "validation.workbook.open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(workbook_file),
        ) from exc

    try:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        template_id = read_valid_template_id_from_system_hash_sheet(workbook)
        if normalize(template_id) != normalize(_TEMPLATE_ID):
            raise validation_error_from_key(
                "validation.template.unknown",
                code="UNKNOWN_TEMPLATE",
                template_id=template_id,
            )

        blueprint = get_blueprint(_TEMPLATE_ID)
        if blueprint is None:
            raise validation_error_from_key(
                "validation.template.unknown",
                code="UNKNOWN_TEMPLATE",
                template_id=_TEMPLATE_ID,
            )

        _validate_sheet_order(workbook, blueprint.sheets)
        _reject_any_formula_cells(workbook, blueprint.sheets, cancel_token=cancel_token)
        _validate_sheet_headers(workbook, blueprint.sheets, cancel_token=cancel_token)

        row_data_by_sheet: dict[str, list[tuple[int, list[Any]]]] = {}
        for sheet_schema in blueprint.sheets:
            if cancel_token is not None:
                cancel_token.raise_if_cancelled()
            row_data = _validated_non_empty_data_rows(workbook[sheet_schema.name], sheet_schema)
            row_data_by_sheet[sheet_schema.name] = row_data
            _validate_sheet_rules_from_schema(sheet_schema, row_data)
            _validate_percentage_columns(sheet_schema, row_data)

        identity = _validate_course_metadata_rules(row_data_by_sheet)
        component_config = _validate_assessment_config_rules(row_data_by_sheet)
        _validate_question_map_rules(row_data_by_sheet, component_config, identity.total_outcomes)
        _validate_students_rules(row_data_by_sheet)
        return identity
    except JobCancelledError:
        raise
    finally:
        workbook.close()


def _validate_sheet_order(workbook: Any, sheet_schemas: Sequence[SheetSchema]) -> None:
    expected = [sheet.name for sheet in sheet_schemas] + ["__SYSTEM_HASH__"]
    actual = list(workbook.sheetnames)
    if actual != expected:
        raise validation_error_from_key(
            "instructor.validation.workbook_sheet_mismatch",
            template_id=_TEMPLATE_ID,
            expected=expected,
            found=actual,
        )


def _validate_sheet_headers(
    workbook: Any,
    sheet_schemas: Sequence[SheetSchema],
    *,
    cancel_token: CancellationToken | None = None,
) -> None:
    for sheet_schema in sheet_schemas:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        if len(sheet_schema.header_matrix) != 1:
            raise validation_error_from_key(
                "instructor.validation.sheet_single_header_row",
                sheet_name=sheet_schema.name,
            )
        expected_headers = list(sheet_schema.header_matrix[0])
        worksheet = workbook[sheet_schema.name]
        for col_index, expected in enumerate(expected_headers, start=1):
            actual = worksheet.cell(row=1, column=col_index).value
            if normalize(actual) != normalize(expected):
                raise validation_error_from_key(
                    "instructor.validation.header_mismatch",
                    sheet_name=sheet_schema.name,
                    expected=expected_headers,
                )


def _reject_any_formula_cells(
    workbook: Any,
    sheet_schemas: Sequence[SheetSchema],
    *,
    cancel_token: CancellationToken | None = None,
) -> None:
    for sheet_schema in sheet_schemas:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()
        worksheet = workbook[sheet_schema.name]
        for row in worksheet.iter_rows(
            min_row=1,
            max_row=max(1, int(worksheet.max_row)),
            min_col=1,
            max_col=max(1, int(worksheet.max_column)),
        ):
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.strip().startswith("="):
                    raise validation_error_from_key(
                        "common.validation_failed_invalid_data",
                        code="FORMULA_NOT_ALLOWED",
                        sheet_name=sheet_schema.name,
                        cell=cell.coordinate,
                    )


def _validated_non_empty_data_rows(
    worksheet: Any,
    sheet_schema: SheetSchema,
) -> list[tuple[int, list[Any]]]:
    header_count = len(sheet_schema.header_matrix[0])
    rows: list[tuple[int, list[Any]]] = []
    for row_number in range(2, int(worksheet.max_row) + 1):
        values = [worksheet.cell(row=row_number, column=col).value for col in range(1, header_count + 1)]
        has_any = any(normalize(value) != "" for value in values)
        if not has_any:
            continue
        for value in values:
            if normalize(value) == "":
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="CELL_EMPTY_NOT_ALLOWED",
                    sheet_name=sheet_schema.name,
                    row=row_number,
                )
        rows.append((row_number, values))
    if not rows:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SHEET_DATA_REQUIRED",
            sheet_name=sheet_schema.name,
        )
    return rows


def _validate_sheet_rules_from_schema(
    sheet_schema: SheetSchema,
    rows: list[tuple[int, list[Any]]],
) -> None:
    headers = list(sheet_schema.header_matrix[0])
    for rule in sheet_schema.validations:
        _apply_schema_rule(sheet_schema.name, headers, rows, rule)


def _apply_schema_rule(
    sheet_name: str,
    headers: list[str],
    rows: list[tuple[int, list[Any]]],
    rule: ValidationRule,
) -> None:
    options = dict(rule.options)
    validation_type = normalize(options.get("validate"))
    if not validation_type:
        return

    first_col = int(rule.first_col) + 1
    last_col = int(rule.last_col) + 1
    allowed_raw = options.get("source")
    allowed_tokens = {
        normalize(value)
        for value in (allowed_raw if isinstance(allowed_raw, list) else [])
    }
    allowed_display = ", ".join(str(value) for value in (allowed_raw if isinstance(allowed_raw, list) else []))
    for row_number, values in rows:
        for col_index in range(first_col, last_col + 1):
            if col_index <= 0 or col_index > len(values):
                continue
            value = values[col_index - 1]
            field_name = headers[col_index - 1] if col_index - 1 < len(headers) else f"col_{col_index}"
            if validation_type == "list":
                if normalize(value) not in allowed_tokens:
                    raise validation_error_from_key(
                        "instructor.validation.allowed_values_required",
                        sheet_name=sheet_name,
                        row=row_number,
                        field=field_name,
                        allowed=allowed_display,
                    )
                continue
            if validation_type == "integer":
                numeric = coerce_excel_number(value)
                if isinstance(numeric, bool) or not isinstance(numeric, int):
                    raise validation_error_from_key(
                        "common.validation_failed_invalid_data",
                        code="INTEGER_VALUE_REQUIRED",
                        sheet_name=sheet_name,
                        row=row_number,
                        field=field_name,
                    )
                criteria = str(options.get("criteria") or "").strip()
                threshold = coerce_excel_number(options.get("value"))
                if criteria == ">" and isinstance(threshold, (int, float)):
                    if numeric <= float(threshold):
                        raise validation_error_from_key(
                            "common.validation_failed_invalid_data",
                            code="INTEGER_VALUE_OUT_OF_RANGE",
                            sheet_name=sheet_name,
                            row=row_number,
                            field=field_name,
                        )
                continue
            if validation_type == "length":
                criteria = str(options.get("criteria") or "").strip().lower()
                text_len = len(str(value).strip())
                minimum = coerce_excel_number(options.get("minimum"))
                maximum = coerce_excel_number(options.get("maximum"))
                if criteria == "between" and isinstance(minimum, (int, float)) and isinstance(maximum, (int, float)):
                    if not (int(minimum) <= text_len <= int(maximum)):
                        raise validation_error_from_key(
                            "common.validation_failed_invalid_data",
                            code="TEXT_LENGTH_OUT_OF_RANGE",
                            sheet_name=sheet_name,
                            row=row_number,
                            field=field_name,
                        )


def _validate_percentage_columns(
    sheet_schema: SheetSchema,
    rows: list[tuple[int, list[Any]]],
) -> None:
    headers = list(sheet_schema.header_matrix[0])
    column_keys = _column_keys(sheet_schema)
    configured_percent_keys = sheet_schema.sheet_rules.get("percentage_column_keys")
    percent_columns: list[int] = []
    if isinstance(configured_percent_keys, (list, tuple)):
        for key in configured_percent_keys:
            if not isinstance(key, str):
                continue
            index = _column_index_by_key(sheet_schema, key)
            if index is not None:
                percent_columns.append(index)
    if not percent_columns:
        percent_columns = [
            index
            for index, header in enumerate(headers)
            if "%" in str(header)
        ]
    for row_number, values in rows:
        for col_index in percent_columns:
            value = coerce_excel_number(values[col_index])
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="PERCENTAGE_NUMERIC_REQUIRED",
                    sheet_name=sheet_schema.name,
                    row=row_number,
                    field=headers[col_index],
                )
            numeric = float(value)
            if numeric < 0.0 or numeric > 100.0:
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="PERCENTAGE_RANGE_INVALID",
                    sheet_name=sheet_schema.name,
                    row=row_number,
                    field=headers[col_index],
                )


def _validate_course_metadata_rules(row_data_by_sheet: dict[str, list[tuple[int, list[Any]]]]) -> _CourseIdentity:
    metadata_sheet = get_sheet_name_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
    metadata_schema = get_sheet_schema_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
    if metadata_schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=metadata_sheet,
        )
    rows = row_data_by_sheet[metadata_sheet]
    fields: dict[str, str] = {}
    row_by_key: dict[str, int] = {}
    required_keys_raw = metadata_schema.sheet_rules.get("required_field_keys")
    required_keys = {
        normalize(key)
        for key in (required_keys_raw if isinstance(required_keys_raw, (list, tuple)) else [])
        if isinstance(key, str) and key.strip()
    }
    if not required_keys:
        required_keys = {
            normalize(COURSE_METADATA_COURSE_CODE_KEY),
            normalize(COURSE_METADATA_SEMESTER_KEY),
            normalize(COURSE_METADATA_SECTION_KEY),
            normalize(COURSE_METADATA_ACADEMIC_YEAR_KEY),
            normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY),
        }

    field_index = _required_column_index(metadata_schema, "field")
    value_index = _required_column_index(metadata_schema, "value")
    for row_number, values in rows:
        key_raw = values[field_index] if field_index < len(values) else ""
        value_raw = values[value_index] if value_index < len(values) else ""
        key = normalize(key_raw)
        if key in row_by_key:
            raise validation_error_from_key(
                "instructor.validation.course_metadata_duplicate_field",
                row=row_number,
                field=key_raw,
            )
        row_by_key[key] = row_number
        fields[key] = str(coerce_excel_number(value_raw) if value_raw is not None else "").strip()

    missing = sorted(key for key in required_keys if key not in fields or not fields[key])
    if missing:
        raise validation_error_from_key(
            "instructor.validation.course_metadata_missing_fields",
            fields=", ".join(missing),
        )
    total_outcomes_key = normalize(str(metadata_schema.sheet_rules.get("total_outcomes_key", "")))
    if not total_outcomes_key:
        total_outcomes_key = normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY)
    total_outcomes_token = fields.get(total_outcomes_key, "").strip()
    try:
        total_outcomes = int(float(total_outcomes_token))
    except (TypeError, ValueError) as exc:
        raise validation_error_from_key("instructor.validation.course_metadata_total_outcomes_invalid") from exc
    if total_outcomes <= 0:
        raise validation_error_from_key("instructor.validation.course_metadata_total_outcomes_invalid")
    return _CourseIdentity(
        template_id=_TEMPLATE_ID,
        course_code=fields.get(normalize(COURSE_METADATA_COURSE_CODE_KEY), ""),
        semester=fields.get(normalize(COURSE_METADATA_SEMESTER_KEY), ""),
        academic_year=fields.get(normalize(COURSE_METADATA_ACADEMIC_YEAR_KEY), ""),
        total_outcomes=total_outcomes,
        section=fields.get(normalize(COURSE_METADATA_SECTION_KEY), ""),
    )


def _validate_assessment_config_rules(
    row_data_by_sheet: dict[str, list[tuple[int, list[Any]]]],
) -> dict[str, dict[str, Any]]:
    assessment_sheet = get_sheet_name_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG)
    assessment_schema = get_sheet_schema_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG)
    if assessment_schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=assessment_sheet,
        )
    rows = row_data_by_sheet[assessment_sheet]
    component_index = _required_column_index(assessment_schema, _COL_COMPONENT)
    weight_index = _required_column_index(assessment_schema, _COL_WEIGHT_PERCENT)
    co_wise_index = _required_column_index(assessment_schema, "co_wise_marks_breakup")
    direct_index = _required_column_index(assessment_schema, "direct")
    assessment_type_index = _required_column_index(assessment_schema, "assessment_type")
    assessment_format_index = _required_column_index(assessment_schema, "assessment_format")
    mode_index = _required_column_index(assessment_schema, "mode")
    participation_index = _required_column_index(assessment_schema, "participation")

    direct_weight_total = 0.0
    indirect_weight_total = 0.0
    direct_count = 0
    indirect_count = 0
    seen_component_keys: set[str] = set()
    component_config: dict[str, dict[str, Any]] = {}
    allowed_assessment_type = {normalize(value) for value in COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS}
    allowed_assessment_format = {normalize(value) for value in COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS}
    allowed_mode = {normalize(value) for value in COURSE_SETUP_ASSESSMENT_MODE_OPTIONS}
    allowed_participation = {normalize(value) for value in COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS}
    for row_number, values in rows:
        component_raw = values[component_index]
        component_key = normalize(component_raw)
        if not component_key:
            raise validation_error_from_key("instructor.validation.assessment_component_required", row=row_number)
        if component_key in seen_component_keys:
            raise validation_error_from_key(
                "instructor.validation.assessment_component_duplicate",
                row=row_number,
                component=component_raw,
            )
        seen_component_keys.add(component_key)

        weight_value = coerce_excel_number(values[weight_index])
        if isinstance(weight_value, bool) or not isinstance(weight_value, (int, float)):
            raise validation_error_from_key("instructor.validation.assessment_weight_numeric", row=row_number)
        weight = float(weight_value)

        co_wise_breakup = _yes_no_value(
            values[co_wise_index],
            sheet_name=assessment_sheet,
            row_number=row_number,
            field_name=str(assessment_schema.header_matrix[0][co_wise_index]),
        )
        is_direct = _yes_no_value(
            values[direct_index],
            sheet_name=assessment_sheet,
            row_number=row_number,
            field_name=str(assessment_schema.header_matrix[0][direct_index]),
        )
        _ensure_allowed_option(
            values[assessment_type_index],
            sheet_name=assessment_sheet,
            row_number=row_number,
            field_name=str(assessment_schema.header_matrix[0][assessment_type_index]),
            allowed_tokens=allowed_assessment_type,
            allowed_display=COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS,
        )
        _ensure_allowed_option(
            values[assessment_format_index],
            sheet_name=assessment_sheet,
            row_number=row_number,
            field_name=str(assessment_schema.header_matrix[0][assessment_format_index]),
            allowed_tokens=allowed_assessment_format,
            allowed_display=COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS,
        )
        _ensure_allowed_option(
            values[mode_index],
            sheet_name=assessment_sheet,
            row_number=row_number,
            field_name=str(assessment_schema.header_matrix[0][mode_index]),
            allowed_tokens=allowed_mode,
            allowed_display=COURSE_SETUP_ASSESSMENT_MODE_OPTIONS,
        )
        _ensure_allowed_option(
            values[participation_index],
            sheet_name=assessment_sheet,
            row_number=row_number,
            field_name=str(assessment_schema.header_matrix[0][participation_index]),
            allowed_tokens=allowed_participation,
            allowed_display=COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS,
        )

        if is_direct:
            direct_weight_total += weight
            direct_count += 1
        else:
            indirect_weight_total += weight
            indirect_count += 1
        component_config[component_key] = {
            "display_name": str(component_raw).strip(),
            "co_wise_breakup": co_wise_breakup,
            "is_direct": is_direct,
        }

    if direct_count <= 0:
        raise validation_error_from_key("instructor.validation.assessment_direct_missing")
    if indirect_count <= 0:
        raise validation_error_from_key("instructor.validation.assessment_indirect_missing")
    expected_total = coerce_excel_number(assessment_schema.sheet_rules.get("weight_total_expected"))
    if isinstance(expected_total, bool) or not isinstance(expected_total, (int, float)):
        expected_total = WEIGHT_TOTAL_EXPECTED
    round_digits = coerce_excel_number(assessment_schema.sheet_rules.get("weight_total_round_digits"))
    if isinstance(round_digits, bool) or not isinstance(round_digits, int):
        round_digits = WEIGHT_TOTAL_ROUND_DIGITS
    indirect_tools_min = coerce_excel_number(assessment_schema.sheet_rules.get("indirect_tools_min"))
    if isinstance(indirect_tools_min, bool) or not isinstance(indirect_tools_min, int):
        indirect_tools_min = 1
    indirect_tools_max = coerce_excel_number(assessment_schema.sheet_rules.get("indirect_tools_max"))
    if isinstance(indirect_tools_max, bool) or not isinstance(indirect_tools_max, int):
        indirect_tools_max = 3

    if indirect_count < indirect_tools_min or indirect_count > indirect_tools_max:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="INDIRECT_TOOL_COUNT_INVALID",
            minimum=indirect_tools_min,
            maximum=indirect_tools_max,
            found=indirect_count,
        )
    if round(direct_weight_total, round_digits) != float(expected_total):
        raise validation_error_from_key(
            "instructor.validation.assessment_direct_total_invalid",
            found=direct_weight_total,
        )
    if round(indirect_weight_total, round_digits) != float(expected_total):
        raise validation_error_from_key(
            "instructor.validation.assessment_indirect_total_invalid",
            found=indirect_weight_total,
        )
    return component_config


def _validate_question_map_rules(
    row_data_by_sheet: dict[str, list[tuple[int, list[Any]]]],
    component_config: dict[str, dict[str, Any]],
    total_outcomes: int,
) -> None:
    question_sheet = get_sheet_name_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_QUESTION_MAP)
    question_schema = get_sheet_schema_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_QUESTION_MAP)
    if question_schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=question_sheet,
        )
    question_headers = list(question_schema.header_matrix[0])
    component_idx = _required_column_index(question_schema, _COL_COMPONENT)
    question_idx = _required_column_index(question_schema, _COL_QUESTION_LABEL)
    max_marks_idx = _required_column_index(question_schema, _COL_MAX_MARKS)
    co_idx = _required_column_index(question_schema, _COL_CO)
    bloom_idx = _required_column_index(question_schema, _COL_BLOOM_LEVEL)
    rows = row_data_by_sheet[question_sheet]
    allowed_bloom_levels = {normalize(value) for value in COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS}
    question_count_by_component: dict[str, int] = {}
    seen_co_wise_questions: set[tuple[str, str]] = set()

    for row_number, values in rows:
        component_raw = values[component_idx]
        question_raw = values[question_idx]
        max_marks_raw = values[max_marks_idx]
        co_raw = values[co_idx]
        bloom_raw = values[bloom_idx]

        component_key = normalize(component_raw)
        if component_key not in component_config:
            raise validation_error_from_key(
                "instructor.validation.question_component_unknown",
                row=row_number,
                component=component_raw,
            )
        question_key = normalize(question_raw)
        if not question_key:
            raise validation_error_from_key(
                "instructor.validation.question_label_required",
                row=row_number,
            )
        max_marks = coerce_excel_number(max_marks_raw)
        if isinstance(max_marks, bool) or not isinstance(max_marks, (int, float)):
            raise validation_error_from_key("instructor.validation.question_max_marks_numeric", row=row_number)
        if float(max_marks) <= 0.0:
            raise validation_error_from_key("instructor.validation.question_max_marks_positive", row=row_number)

        co_values = parse_co_tokens(co_raw, dedupe=False)
        if not co_values:
            raise validation_error_from_key("instructor.validation.question_co_required", row=row_number)
        if len(set(co_values)) != len(co_values):
            raise validation_error_from_key("instructor.validation.question_co_no_repeat", row=row_number)
        if any(value <= 0 or value > total_outcomes for value in co_values):
            raise validation_error_from_key(
                "instructor.validation.question_co_out_of_range",
                row=row_number,
                total_outcomes=total_outcomes,
            )
        if normalize(bloom_raw) not in allowed_bloom_levels:
            raise validation_error_from_key(
                "instructor.validation.allowed_values_required",
                sheet_name=question_sheet,
                row=row_number,
                field=question_headers[bloom_idx],
                allowed=", ".join(COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS),
            )

        is_co_wise = bool(component_config[component_key]["co_wise_breakup"])
        question_count_by_component[component_key] = question_count_by_component.get(component_key, 0) + 1
        if is_co_wise:
            if len(co_values) != 1:
                raise validation_error_from_key(
                    "instructor.validation.question_co_wise_requires_one",
                    row=row_number,
                    component=component_raw,
                )
            question_id = (component_key, question_key)
            if question_id in seen_co_wise_questions:
                raise validation_error_from_key(
                    "instructor.validation.question_duplicate_for_component",
                    row=row_number,
                    question=question_raw,
                    component=component_raw,
                )
            seen_co_wise_questions.add(question_id)

    for component_key, config in component_config.items():
        count = question_count_by_component.get(component_key, 0)
        if bool(config.get("is_direct")) and count <= 0:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="QUESTION_MAP_COMPONENT_MISSING",
                component=config.get("display_name", component_key),
            )


def _validate_students_rules(row_data_by_sheet: dict[str, list[tuple[int, list[Any]]]]) -> None:
    students_sheet = get_sheet_name_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_STUDENTS)
    students_schema = get_sheet_schema_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_STUDENTS)
    if students_schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=students_sheet,
        )
    reg_no_index = _required_column_index(students_schema, _COL_REG_NO)
    student_name_index = _required_column_index(students_schema, _COL_STUDENT_NAME)
    rows = row_data_by_sheet[students_sheet]
    seen_reg_numbers: set[str] = set()
    for row_number, values in rows:
        reg_no = str(values[reg_no_index]).strip() if reg_no_index < len(values) else ""
        student_name = str(values[student_name_index]).strip() if student_name_index < len(values) else ""
        if not reg_no or not student_name:
            raise validation_error_from_key(
                "instructor.validation.students_reg_and_name_required",
                row=row_number,
            )
        reg_key = normalize(reg_no)
        if reg_key in seen_reg_numbers:
            raise validation_error_from_key(
                "instructor.validation.students_duplicate_reg_no",
                row=row_number,
                reg_no=reg_no,
            )
        seen_reg_numbers.add(reg_key)


def _column_keys(sheet_schema: SheetSchema) -> tuple[str, ...]:
    raw = sheet_schema.sheet_rules.get("column_keys")
    if not isinstance(raw, (list, tuple)):
        return tuple()
    return tuple(normalize(value) for value in raw if isinstance(value, str) and normalize(value))


def _column_index_by_key(sheet_schema: SheetSchema, key: str) -> int | None:
    wanted = normalize(key)
    for index, value in enumerate(_column_keys(sheet_schema)):
        if value == wanted:
            return index
    return None


def _required_column_index(sheet_schema: SheetSchema, key: str) -> int:
    index = _column_index_by_key(sheet_schema, key)
    if index is not None:
        return index
    raise validation_error_from_key(
        "common.validation_failed_invalid_data",
        code="SCHEMA_COLUMN_KEY_MISSING",
        sheet_name=sheet_schema.name,
        column_key=key,
    )


def _yes_no_value(value: Any, *, sheet_name: str, row_number: int, field_name: str) -> bool:
    token = normalize(value)
    if token not in {"yes", "no"}:
        raise validation_error_from_key(
            "instructor.validation.yes_no_required",
            sheet_name=sheet_name,
            row=row_number,
            field=field_name,
        )
    return token == "yes"


def _ensure_allowed_option(
    value: Any,
    *,
    sheet_name: str,
    row_number: int,
    field_name: str,
    allowed_tokens: set[str],
    allowed_display: Sequence[str],
) -> None:
    if normalize(value) in allowed_tokens:
        return
    raise validation_error_from_key(
        "instructor.validation.allowed_values_required",
        sheet_name=sheet_name,
        row=row_number,
        field=field_name,
        allowed=", ".join(allowed_display),
    )


__all__ = ["validate_course_details_workbook", "validate_course_details_workbooks"]
