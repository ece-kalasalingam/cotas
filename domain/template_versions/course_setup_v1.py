"""Version-specific handlers for COURSE_SETUP_V1."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

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
    ID_COURSE_SETUP,
    LIKERT_MAX,
    LIKERT_MIN,
    MIN_MARK_VALUE,
)
from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from common.registry import (
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS,
    COURSE_SETUP_ASSESSMENT_MODE_OPTIONS,
    COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS,
    COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS,
    COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MAX_LENGTH,
    COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MIN_LENGTH,
    COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS,
    COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG,
    COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    COURSE_SETUP_SHEET_KEY_QUESTION_MAP,
    COURSE_SETUP_SHEET_KEY_STUDENTS,
    WEIGHT_TOTAL_EXPECTED,
    WEIGHT_TOTAL_ROUND_DIGITS,
    get_sheet_headers_by_key,
    get_sheet_name_by_key,
)
from common.sample_setup_data import SAMPLE_SETUP_DATA
from common.utils import coerce_excel_number, normalize
from common.workbook_integrity.workbook_signing import sign_payload
from domain.assessment_semantics import parse_assessment_components
from domain.co_token_parser import parse_co_tokens
from domain.template_strategy_router import assert_template_id_matches

_logger = logging.getLogger(__name__)
_MAX_DECIMAL_PLACES = 2
_FORMULA_SUM_TEMPLATE = "=SUM({start}:{end})"
_LOG_STEP3_HIGH_ABSENCE = "Step3 anomaly: high absence ratio sheet=%s col=%s absent=%s total=%s"
_LOG_STEP3_NEAR_CONSTANT = (
    "Step3 anomaly: near-constant marks sheet=%s col=%s dominant_count=%s numeric_total=%s"
)
_last_marks_anomaly_warnings: list[str] = []
COURSE_METADATA_SHEET = get_sheet_name_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
ASSESSMENT_CONFIG_SHEET = get_sheet_name_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG)
QUESTION_MAP_SHEET = get_sheet_name_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_QUESTION_MAP)
CO_DESCRIPTION_SHEET = get_sheet_name_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION)
STUDENTS_SHEET = get_sheet_name_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_STUDENTS)
COURSE_METADATA_HEADERS = get_sheet_headers_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
ASSESSMENT_CONFIG_HEADERS = get_sheet_headers_by_key(
    ID_COURSE_SETUP,
    COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG,
)
QUESTION_MAP_HEADERS = get_sheet_headers_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_QUESTION_MAP)
CO_DESCRIPTION_HEADERS = get_sheet_headers_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION)
STUDENTS_HEADERS = get_sheet_headers_by_key(ID_COURSE_SETUP, COURSE_SETUP_SHEET_KEY_STUDENTS)
_ASSESSMENT_TYPE_OPTION_TOKENS = {normalize(value) for value in COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS}
_ASSESSMENT_FORMAT_OPTION_TOKENS = {normalize(value) for value in COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS}
_ASSESSMENT_MODE_OPTION_TOKENS = {normalize(value) for value in COURSE_SETUP_ASSESSMENT_MODE_OPTIONS}
_ASSESSMENT_PARTICIPATION_OPTION_TOKENS = {
    normalize(value) for value in COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS
}
_QUESTION_DOMAIN_LEVEL_OPTION_TOKENS = {
    normalize(value) for value in COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS
}
_MARK_COMPONENT_SHEET_KINDS = {
    LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
    LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
    LAYOUT_SHEET_KIND_INDIRECT,
}
_SUPPORTED_OPERATIONS = frozenset(
    {
        "generate_workbook",
        "validate_filled_marks_manifest_schema",
        "consume_last_marks_anomaly_warnings",
        "generate_co_attainment",
    }
)


def _reset_marks_anomaly_warnings() -> None:
    _last_marks_anomaly_warnings.clear()


def consume_last_marks_anomaly_warnings() -> list[str]:
    warnings = list(_last_marks_anomaly_warnings)
    _last_marks_anomaly_warnings.clear()
    return warnings


@dataclass(slots=True, frozen=True)
class CourseSetupV1Strategy:
    template_id: str = "COURSE_SETUP_V1"

    def supports_operation(self, operation: str) -> bool:
        return str(operation).strip() in _SUPPORTED_OPERATIONS

    def default_workbook_name(
        self,
        *,
        workbook_kind: str,
        context: Mapping[str, Any] | None,
        fallback: str,
    ) -> str:
        del workbook_kind
        del context
        return fallback

    def generate_workbook(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        output_path: str | Path,
        workbook_name: str | None,
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> object:
        assert_template_id_matches(
            actual_template_id=template_id,
            expected_template_id=self.template_id,
        )
        resolved_workbook_name = (workbook_name or Path(output_path).name).strip()
        if not resolved_workbook_name:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_NAME_REQUIRED",
            )
        kind = normalize(workbook_kind)
        payload = dict(context or {})
        if kind == "co_attainment":
            source_paths_raw = payload.get("source_paths")
            source_paths = [Path(path) for path in source_paths_raw] if isinstance(source_paths_raw, list) else []
            if not source_paths:
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="COA_SOURCE_WORKBOOK_REQUIRED",
                )
            return self.generate_co_attainment(
                source_paths,
                Path(output_path),
                token=cancel_token or CancellationToken(),
                thresholds=payload.get("thresholds")
                if isinstance(payload.get("thresholds"), tuple)
                else None,
                co_attainment_percent=float(payload["co_attainment_percent"])
                if payload.get("co_attainment_percent") is not None
                else None,
                co_attainment_level=int(payload["co_attainment_level"])
                if payload.get("co_attainment_level") is not None
                else None,
            )
        if kind == "co_analysis":
            source_paths_raw = payload.get("source_paths")
            source_paths = [Path(path) for path in source_paths_raw] if isinstance(source_paths_raw, list) else []
            if not source_paths:
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="COA_SOURCE_WORKBOOK_REQUIRED",
                )
            from domain.co_analysis_engine import generate_co_analysis_workbook

            return generate_co_analysis_workbook(
                source_paths,
                Path(output_path),
                token=cancel_token,
                thresholds=payload.get("thresholds")
                if isinstance(payload.get("thresholds"), tuple)
                else None,
                co_attainment_percent=float(payload["co_attainment_percent"])
                if payload.get("co_attainment_percent") is not None
                else None,
                co_attainment_level=int(payload["co_attainment_level"])
                if payload.get("co_attainment_level") is not None
                else None,
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
        )

    def validate_course_details_rules(self, workbook: object, *, context: object) -> None:
        del workbook
        del context
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind="course_details_template",
            template_id=self.template_id,
        )

    def extract_marks_template_context(self, workbook: object, *, context: object) -> dict[str, Any]:
        del workbook
        del context
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind="marks_template",
            template_id=self.template_id,
        )

    def write_marks_template_workbook(
        self,
        workbook: object,
        context_data: dict[str, Any],
        *,
        context: object,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, Any]:
        del workbook
        del context_data
        del context
        del cancel_token
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind="marks_template",
            template_id=self.template_id,
        )

    def validate_filled_marks_manifest_schema(self, workbook: object, manifest: object) -> None:
        validate_filled_marks_manifest_schema(workbook, manifest)

    def consume_last_marks_anomaly_warnings(self) -> list[str]:
        return consume_last_marks_anomaly_warnings()

    def generate_final_report(
        self,
        filled_marks_path: str | Path,
        output_path: str | Path,
        *,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        del filled_marks_path
        del output_path
        del cancel_token
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind="final_report",
            template_id=self.template_id,
        )

    def generate_co_attainment(
        self,
        source_paths: list[Path],
        output_path: Path,
        *,
        token: CancellationToken,
        thresholds: tuple[float, float, float] | None = None,
        co_attainment_percent: float | None = None,
        co_attainment_level: int | None = None,
    ) -> object:
        from domain.template_versions.course_setup_v1_coordinator_engine import (
            _generate_co_attainment_workbook_course_setup_v1,
        )

        return _generate_co_attainment_workbook_course_setup_v1(
            source_paths,
            output_path,
            token=token,
            thresholds=thresholds,
            co_attainment_percent=co_attainment_percent,
            co_attainment_level=co_attainment_level,
        )

def validate_course_details_rules(workbook: Any) -> None:
    metadata_sheet = workbook[COURSE_METADATA_SHEET]
    assessment_sheet = workbook[ASSESSMENT_CONFIG_SHEET]
    question_map_sheet = workbook[QUESTION_MAP_SHEET]
    co_description_sheet = workbook[CO_DESCRIPTION_SHEET]
    students_sheet = workbook[STUDENTS_SHEET]

    total_outcomes = _validate_course_metadata(metadata_sheet)
    component_config = _validate_assessment_config(assessment_sheet)
    _validate_question_map(question_map_sheet, component_config, total_outcomes)
    _validate_co_description(co_description_sheet)
    _validate_students(students_sheet)


def validate_filled_marks_manifest_schema(workbook: Any, manifest: Any) -> None:
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
            raise validation_error_from_key(
                
                    "instructor.validation.unexpected_header",
                    sheet_name=worksheet.title,
                    col=col_index,
                
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
            raise validation_error_from_key("instructor.validation.course_metadata_field_empty", row=row_number)
        if field_key in actual_values:
            raise validation_error_from_key(
                
                    "instructor.validation.course_metadata_duplicate_field",
                    row=row_number,
                    field=field_raw,
                
            )
        if field_key not in expected_field_types:
            raise validation_error_from_key(
                
                    "instructor.validation.course_metadata_unknown_field",
                    row=row_number,
                    field=field_raw,
                
            )
        if normalize(value_raw) == "":
            raise validation_error_from_key(
                
                    "instructor.validation.course_metadata_value_required",
                    row=row_number,
                    field=field_raw,
                
            )
        actual_values[field_key] = coerce_excel_number(value_raw)

    missing_fields = [name for name in expected_field_types if name not in actual_values]
    if missing_fields:
        raise validation_error_from_key(
            
                "instructor.validation.course_metadata_missing_fields",
                fields=", ".join(missing_fields),
            
        )

    for field_key, expected_type in expected_field_types.items():
        value = actual_values[field_key]
        if expected_type is int:
            if isinstance(value, bool) or not isinstance(value, int):
                raise validation_error_from_key(
                    "instructor.validation.course_metadata_field_must_be_int", field=field_key
                )
        else:
            if not isinstance(value, str) or normalize(value) == "":
                raise validation_error_from_key(
                    
                        "instructor.validation.course_metadata_field_must_be_non_empty_str",
                        field=field_key,
                    
                )

    total_outcomes = actual_values.get(normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY))
    if isinstance(total_outcomes, bool) or not isinstance(total_outcomes, int) or total_outcomes <= 0:
        raise validation_error_from_key("instructor.validation.course_metadata_total_outcomes_invalid")
    return total_outcomes


def _parse_allowed_option(
    value: Any,
    *,
    sheet_name: str,
    row_number: int,
    field_name: str,
    allowed_tokens: set[str],
    allowed_display: Sequence[str],
) -> str:
    token = normalize(value)
    if token not in allowed_tokens:
        raise validation_error_from_key(
            
                "instructor.validation.allowed_values_required",
                sheet_name=sheet_name,
                row=row_number,
                field=field_name,
                allowed=", ".join(allowed_display),
            
        )
    return token


def _validate_assessment_config(worksheet: Any) -> dict[str, dict[str, Any]]:
    assessment_headers = ASSESSMENT_CONFIG_HEADERS
    expected_headers = list(assessment_headers)
    _header_index_map(worksheet, expected_headers)
    rows = _iter_data_rows(worksheet, len(expected_headers))
    components = parse_assessment_components(
        rows,
        template_id="COURSE_SETUP_V1",
        sheet_name=ASSESSMENT_CONFIG_SHEET,
        row_start=2,
        on_blank_component="error",
        duplicate_policy="error",
        require_non_empty=True,
        validate_allowed_options=True,
        assessment_type_allowed_tokens=_ASSESSMENT_TYPE_OPTION_TOKENS,
        assessment_type_allowed_display=COURSE_SETUP_ASSESSMENT_TYPE_OPTIONS,
        assessment_format_allowed_tokens=_ASSESSMENT_FORMAT_OPTION_TOKENS,
        assessment_format_allowed_display=COURSE_SETUP_ASSESSMENT_FORMAT_OPTIONS,
        mode_allowed_tokens=_ASSESSMENT_MODE_OPTION_TOKENS,
        mode_allowed_display=COURSE_SETUP_ASSESSMENT_MODE_OPTIONS,
        participation_allowed_tokens=_ASSESSMENT_PARTICIPATION_OPTION_TOKENS,
        participation_allowed_display=COURSE_SETUP_ASSESSMENT_PARTICIPATION_OPTIONS,
    )

    component_config: dict[str, dict[str, Any]] = {}
    direct_weight_total = 0.0
    indirect_weight_total = 0.0
    direct_count = 0
    indirect_count = 0

    for component in components:
        if component.is_direct:
            direct_weight_total += component.weight
            direct_count += 1
        else:
            indirect_weight_total += component.weight
            indirect_count += 1

        component_config[component.component_key] = {
            "display_name": component.component_name,
            "co_wise_breakup": component.co_wise_breakup,
        }

    if direct_count == 0:
        raise validation_error_from_key("instructor.validation.assessment_direct_missing")
    if indirect_count == 0:
        raise validation_error_from_key("instructor.validation.assessment_indirect_missing")
    if round(direct_weight_total, WEIGHT_TOTAL_ROUND_DIGITS) != WEIGHT_TOTAL_EXPECTED:
        raise validation_error_from_key(
            "instructor.validation.assessment_direct_total_invalid", found=direct_weight_total
        )
    if round(indirect_weight_total, WEIGHT_TOTAL_ROUND_DIGITS) != WEIGHT_TOTAL_EXPECTED:
        raise validation_error_from_key(
            "instructor.validation.assessment_indirect_total_invalid", found=indirect_weight_total
        )

    return component_config


def _co_tokens(value: Any) -> list[int]:
    return parse_co_tokens(value, dedupe=False)


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
    domain_level_header = normalize(expected_headers[4])
    rows = _iter_data_rows(worksheet, len(expected_headers))

    if not rows:
        raise validation_error_from_key("instructor.validation.question_map_row_required_one")

    seen_co_wise_questions: set[tuple[str, str]] = set()
    for row_number, row in enumerate(rows, start=2):
        component_raw = row[header_map[component_header]]
        component_key = normalize(component_raw)
        if not component_key:
            raise validation_error_from_key("instructor.validation.question_component_required", row=row_number)
        if component_key not in component_config:
            raise validation_error_from_key(
                
                    "instructor.validation.question_component_unknown",
                    row=row_number,
                    component=component_raw,
                
            )

        question_raw = row[header_map[question_header]]
        question_key = normalize(question_raw)
        if not question_key:
            raise validation_error_from_key(
                "instructor.validation.question_label_required", row=row_number
            )

        max_marks = coerce_excel_number(row[header_map[max_marks_header]])
        if isinstance(max_marks, bool) or not isinstance(max_marks, (int, float)):
            raise validation_error_from_key("instructor.validation.question_max_marks_numeric", row=row_number)
        if float(max_marks) <= 0:
            raise validation_error_from_key(
                "instructor.validation.question_max_marks_positive", row=row_number
            )

        co_values = _co_tokens(row[header_map[co_header]])
        if not co_values:
            raise validation_error_from_key("instructor.validation.question_co_required", row=row_number)
        if len(set(co_values)) != len(co_values):
            raise validation_error_from_key("instructor.validation.question_co_no_repeat", row=row_number)
        if any(co_number <= 0 or co_number > total_outcomes for co_number in co_values):
            raise validation_error_from_key(
                
                    "instructor.validation.question_co_out_of_range",
                    row=row_number,
                    total_outcomes=total_outcomes,
                
            )
        _parse_allowed_option(
            row[header_map[domain_level_header]],
            sheet_name=QUESTION_MAP_SHEET,
            row_number=row_number,
            field_name=expected_headers[4],
            allowed_tokens=_QUESTION_DOMAIN_LEVEL_OPTION_TOKENS,
            allowed_display=COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS,
        )

        is_co_wise = bool(component_config[component_key]["co_wise_breakup"])
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


def _validate_co_description(worksheet: Any) -> None:
    expected_headers = list(CO_DESCRIPTION_HEADERS)
    header_map = _header_index_map(worksheet, expected_headers)
    co_number_header = normalize(expected_headers[0])
    description_header = normalize(expected_headers[1])
    domain_level_header = normalize(expected_headers[2])
    summary_header = normalize(expected_headers[3])
    rows = _iter_data_rows(worksheet, len(expected_headers))

    if not rows:
        raise validation_error_from_key("instructor.validation.co_description_row_required_one")

    seen_co_numbers: set[int] = set()
    for row_number, row in enumerate(rows, start=2):
        co_number_raw = row[header_map[co_number_header]]
        co_number = coerce_excel_number(co_number_raw)
        if isinstance(co_number, bool) or not isinstance(co_number, int) or co_number <= 0:
            raise validation_error_from_key(
                
                    "instructor.validation.co_description_number_positive_int_required",
                    row=row_number,
                
            )
        if co_number in seen_co_numbers:
            raise validation_error_from_key(
                
                    "instructor.validation.co_description_number_duplicate",
                    row=row_number,
                    co_number=co_number,
                
            )
        seen_co_numbers.add(co_number)

        _ = row[header_map[description_header]]
        _parse_allowed_option(
            row[header_map[domain_level_header]],
            sheet_name=CO_DESCRIPTION_SHEET,
            row_number=row_number,
            field_name=expected_headers[2],
            allowed_tokens=_QUESTION_DOMAIN_LEVEL_OPTION_TOKENS,
            allowed_display=COURSE_SETUP_QUESTION_DOMAIN_LEVEL_OPTIONS,
        )
        summary_text = str(row[header_map[summary_header]]).strip() if row[header_map[summary_header]] is not None else ""
        summary_length = len(summary_text)
        if (
            summary_length < COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MIN_LENGTH
            or summary_length > COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MAX_LENGTH
        ):
            raise validation_error_from_key(
                
                    "instructor.validation.co_description_summary_length_invalid",
                    row=row_number,
                    minimum=COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MIN_LENGTH,
                    maximum=COURSE_SETUP_CO_DESCRIPTION_SUMMARY_MAX_LENGTH,
                
            )


def _validate_students(worksheet: Any) -> None:
    expected_headers = list(STUDENTS_HEADERS)
    header_map = _header_index_map(worksheet, expected_headers)
    reg_no_header = normalize(expected_headers[0])
    student_name_header = normalize(expected_headers[1])
    rows = _iter_data_rows(worksheet, len(expected_headers))

    if not rows:
        raise validation_error_from_key("instructor.validation.students_row_required_one")

    seen_reg_numbers: set[str] = set()
    for row_number, row in enumerate(rows, start=2):
        reg_no_raw = row[header_map[reg_no_header]]
        student_name_raw = row[header_map[student_name_header]]

        reg_no = str(reg_no_raw).strip() if reg_no_raw is not None else ""
        student_name = str(student_name_raw).strip() if student_name_raw is not None else ""

        if not reg_no or not student_name:
            raise validation_error_from_key(
                "instructor.validation.students_reg_and_name_required", row=row_number
            )

        reg_key = normalize(reg_no)
        if reg_key in seen_reg_numbers:
            raise validation_error_from_key(
                "instructor.validation.students_duplicate_reg_no", row=row_number, reg_no=reg_no
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
    # Column 4 is "D" and the first mark-entry column across all component sheets.
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        return range(4, 5)
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
        # Direct CO-wise sheets append "Total" as the last header; mark cells are before it.
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
        # CO split columns should remain formula-driven for every student row.
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


def _log_marks_anomaly_warnings(
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
    start_row = _marks_data_start_row(sheet_kind, header_row)
    mark_cols = _marks_entry_columns(sheet_kind, header_count)
    absent_count_by_col: dict[int, int] = {col: 0 for col in mark_cols}
    numeric_count_by_col: dict[int, int] = {col: 0 for col in mark_cols}
    frequency_by_value_by_col: dict[int, dict[float, int]] = {col: {} for col in mark_cols}
    for col in mark_cols:
        for row in range(start_row, start_row + student_count):
            cell_value = worksheet.cell(row=row, column=col).value
            token = normalize(cell_value)
            if token == "a":
                absent_count_by_col[col] += 1
                continue
            numeric = coerce_excel_number(cell_value)
            if isinstance(numeric, (int, float)) and not isinstance(numeric, bool):
                numeric_count_by_col[col] += 1
                number = float(numeric)
                frequency_by_value = frequency_by_value_by_col[col]
                frequency_by_value[number] = frequency_by_value.get(number, 0) + 1
    _log_marks_anomaly_warnings_from_stats(
        sheet_name=sheet_name,
        mark_cols=mark_cols,
        student_count=student_count,
        absent_count_by_col=absent_count_by_col,
        numeric_count_by_col=numeric_count_by_col,
        frequency_by_value_by_col=frequency_by_value_by_col,
    )


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
        absent_count = absent_count_by_col.get(col, 0)
        if absent_count / student_count >= 0.9:
            warning_message = (
                f"High absence ratio detected in {sheet_name} ({_excel_col_name(col)}): "
                f"{absent_count}/{student_count} marked absent."
            )
            _last_marks_anomaly_warnings.append(warning_message)
            _logger.warning(
                _LOG_STEP3_HIGH_ABSENCE,
                sheet_name,
                _excel_col_name(col),
                absent_count,
                student_count,
            )
        numeric_count = numeric_count_by_col.get(col, 0)
        if numeric_count:
            same = max(frequency_by_value_by_col.get(col, {0.0: 0}).values())
            if same / numeric_count >= 0.95:
                warning_message = (
                    f"Near-constant marks detected in {sheet_name} ({_excel_col_name(col)}): "
                    f"{same}/{numeric_count} values are identical."
                )
                _last_marks_anomaly_warnings.append(warning_message)
                _logger.warning(
                    _LOG_STEP3_NEAR_CONSTANT,
                    sheet_name,
                    _excel_col_name(col),
                    same,
                    numeric_count,
                )


