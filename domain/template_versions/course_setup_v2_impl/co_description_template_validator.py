"""COURSE_SETUP_V2 filled CO-description workbook validation."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from common.error_catalog import resolve_validation_issue, validation_error_from_key
from common.exceptions import ValidationError
from common.jobs import CancellationToken
from common.registry import (
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    get_sheet_name_by_key,
    get_sheet_schema_by_key,
)
from common.runtime_dependency_guard import import_runtime_dependency
from common.sheet_schema import SheetSchema
from common.utils import assert_not_symlink_path, coerce_excel_number, normalize
from domain.template_versions.course_setup_v2_impl.course_template_validator import (
    _validate_course_metadata_rules,
    _validated_non_empty_data_rows,
)
from domain.template_versions.course_setup_v2_impl.validation_batch_runner import (
    BatchValidationAccumulator,
    BatchValidationRunner,
    ValidationRejectionDecision,
)

_CO_DESCRIPTION_HEADERS = {
    "co_number": "co#",
    "description": "description",
    "summary": "summary_of_topics/expts./project",
}


@dataclass(frozen=True, slots=True)
class _CoDescriptionWorkbookIdentity:
    template_id: str
    course_code: str
    semester: str
    academic_year: str
    total_outcomes: int

    def cohort_key(self) -> tuple[str, str, str, int]:
        """Build normalized cohort key tuple for cross-workbook comparison."""
        return (
            normalize(self.course_code),
            normalize(self.semester),
            normalize(self.academic_year),
            int(self.total_outcomes),
        )


def validate_co_description_workbooks(
    workbook_paths: Sequence[str | Path],
    *,
    template_id: str,
    cancel_token: CancellationToken | None = None,
) -> dict[str, object]:
    """Validate a batch of filled CO-description workbooks for one template."""
    baseline_cohort: tuple[str, str, str, int] | None = None
    baseline_identity: _CoDescriptionWorkbookIdentity | None = None

    def _validate_path(path: str) -> _CoDescriptionWorkbookIdentity:
        return _validate_co_description_workbook_impl(
            workbook_path=path,
            expected_template_id=template_id,
            cancel_token=cancel_token,
        )

    def _on_validated(
        acc: BatchValidationAccumulator,
        path: str,
        identity: _CoDescriptionWorkbookIdentity,
    ) -> None:
        nonlocal baseline_cohort
        nonlocal baseline_identity
        cohort = identity.cohort_key()
        if baseline_cohort is None:
            baseline_cohort = cohort
            baseline_identity = identity
            acc.add_valid(path=path, template_id=identity.template_id)
            return
        if cohort != baseline_cohort:
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
                code="CO_DESCRIPTION_TEMPLATE_COHORT_MISMATCH",
                context={
                    "workbook": path,
                    "fields": ", ".join(mismatch_fields) if mismatch_fields else "cohort",
                },
                fallback_message=(
                    "File skipped because course cohort metadata does not match "
                    "(course code, semester, academic year, total outcomes must match)."
                ),
            )
            acc.add_rejection(
                path=path,
                issue=issue,
                decision=ValidationRejectionDecision(
                    reason_kind="cohort_mismatch",
                    mark_invalid=False,
                    mark_mismatched=True,
                ),
            )
            return
        acc.add_valid(path=path, template_id=identity.template_id)

    def _classify_validation_error(
        _path: str,
        _exc: ValidationError,
        issue: dict[str, object],
    ) -> ValidationRejectionDecision:
        reason_kind = (
            "template_mismatch"
            if str(issue.get("code", "")).strip() == "UNKNOWN_TEMPLATE"
            else "invalid"
        )
        return ValidationRejectionDecision(
            reason_kind=reason_kind,
            mark_invalid=True,
            mark_mismatched=reason_kind == "template_mismatch",
        )

    runner = BatchValidationRunner[_CoDescriptionWorkbookIdentity](
        issue_builder=_issue_dict,
        duplicate_path_issue_code="CO_DESCRIPTION_TEMPLATE_DUPLICATE_PATH",
        unexpected_issue_code="CO_DESCRIPTION_TEMPLATE_UNEXPECTED_REJECTION",
    )
    return runner.run(
        workbook_paths=workbook_paths,
        validate_path=_validate_path,
        on_validated=_on_validated,
        cancel_token=cancel_token,
        classify_validation_error=_classify_validation_error,
    )


def _validate_co_description_workbook_impl(
    *,
    workbook_path: str | Path,
    expected_template_id: str,
    cancel_token: CancellationToken | None = None,
) -> _CoDescriptionWorkbookIdentity:
    """Validate one filled CO-description workbook and return cohort identity."""
    from domain.template_strategy_router import (
        assert_template_id_matches,
        read_valid_template_id_from_system_hash_sheet,
    )

    openpyxl = import_runtime_dependency("openpyxl")
    workbook_file = Path(workbook_path)
    if not workbook_file.exists():
        raise validation_error_from_key(
            "validation.workbook.not_found",
            code="WORKBOOK_NOT_FOUND",
            workbook=str(workbook_file),
        )
    assert_not_symlink_path(workbook_file, context_key="workbook")

    try:
        workbook = openpyxl.load_workbook(workbook_file, data_only=False, read_only=True)
    except Exception as exc:
        raise validation_error_from_key(
            "validation.workbook.open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(workbook_file),
        ) from exc
    try:
        actual_template_id = read_valid_template_id_from_system_hash_sheet(workbook)
        assert_template_id_matches(
            actual_template_id=actual_template_id,
            expected_template_id=expected_template_id,
        )
    finally:
        workbook.close()
    if cancel_token is not None:
        cancel_token.raise_if_cancelled()

    try:
        workbook = openpyxl.load_workbook(workbook_file, data_only=False, read_only=False)
    except Exception as exc:
        raise validation_error_from_key(
            "validation.workbook.open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(workbook_file),
        ) from exc
    try:
        return _read_co_description_workbook_identity(
            workbook=workbook,
            template_id=expected_template_id,
        )
    finally:
        workbook.close()


def _read_co_description_workbook_identity(
    *,
    workbook: Any,
    template_id: str,
) -> _CoDescriptionWorkbookIdentity:
    metadata_sheet_name = get_sheet_name_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
    metadata_schema = get_sheet_schema_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
    if metadata_schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=metadata_sheet_name,
        )
    if metadata_sheet_name not in workbook.sheetnames:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SHEET_DATA_REQUIRED",
            sheet_name=metadata_sheet_name,
        )
    metadata_rows = _validated_non_empty_data_rows(workbook[metadata_sheet_name], metadata_schema)
    identity = _validate_course_metadata_rules({metadata_sheet_name: metadata_rows})

    co_description_sheet_name = get_sheet_name_by_key(template_id, COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION)
    co_description_schema = get_sheet_schema_by_key(template_id, COURSE_SETUP_SHEET_KEY_CO_DESCRIPTION)
    if co_description_schema is None:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_MISSING",
            sheet_name=co_description_sheet_name,
        )
    if co_description_sheet_name not in workbook.sheetnames:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SHEET_DATA_REQUIRED",
            sheet_name=co_description_sheet_name,
        )
    co_description_rows = _validated_non_empty_data_rows(
        workbook[co_description_sheet_name],
        co_description_schema,
    )
    _validate_co_description_rows(
        rows=co_description_rows,
        sheet_schema=co_description_schema,
        total_outcomes=int(identity.total_outcomes),
    )
    return _CoDescriptionWorkbookIdentity(
        template_id=template_id,
        course_code=identity.course_code,
        semester=identity.semester,
        academic_year=identity.academic_year,
        total_outcomes=identity.total_outcomes,
    )


def _validate_co_description_rows(
    *,
    rows: list[tuple[int, list[Any]]],
    sheet_schema: SheetSchema,
    total_outcomes: int,
) -> None:
    headers = [normalize(value) for value in sheet_schema.header_matrix[0]]
    co_index = _required_header_index(headers, _CO_DESCRIPTION_HEADERS["co_number"])
    description_index = _required_header_index(headers, _CO_DESCRIPTION_HEADERS["description"])
    summary_index = _required_header_index(headers, _CO_DESCRIPTION_HEADERS["summary"])
    summary_min_length, summary_max_length = _summary_length_bounds(
        sheet_schema=sheet_schema,
        summary_col_index=summary_index,
    )

    seen_co_numbers: set[int] = set()
    for row_number, values in rows:
        co_raw = values[co_index] if co_index < len(values) else ""
        co_number = _validated_co_number(co_raw=co_raw, row_number=row_number)
        if co_number in seen_co_numbers:
            raise validation_error_from_key(
                "instructor.validation.co_description_number_duplicate",
                code="CO_DESCRIPTION_NUMBER_DUPLICATE",
                row=row_number,
                co_number=co_number,
            )
        seen_co_numbers.add(co_number)

        description_value = str(values[description_index]).strip() if description_index < len(values) else ""
        if not description_value:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="CO_DESCRIPTION_DESCRIPTION_REQUIRED",
                row=row_number,
            )
        summary_value = str(values[summary_index]).strip() if summary_index < len(values) else ""
        if not summary_value:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="CO_DESCRIPTION_SUMMARY_REQUIRED",
                row=row_number,
            )
        if len(summary_value) < summary_min_length or len(summary_value) > summary_max_length:
            raise validation_error_from_key(
                "instructor.validation.co_description_summary_length_invalid",
                code="CO_DESCRIPTION_SUMMARY_LENGTH_INVALID",
                row=row_number,
                minimum=summary_min_length,
                maximum=summary_max_length,
            )

    if not seen_co_numbers:
        raise validation_error_from_key(
            "instructor.validation.co_description_row_required_one",
            code="CO_DESCRIPTION_ROW_REQUIRED_ONE",
        )
    if total_outcomes <= 0:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="COA_TOTAL_OUTCOMES_MISSING",
        )
    expected_numbers = set(range(1, total_outcomes + 1))
    if seen_co_numbers != expected_numbers:
        missing = sorted(expected_numbers - seen_co_numbers)
        extras = sorted(seen_co_numbers - expected_numbers)
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="CO_DESCRIPTION_CO_NUMBER_SET_MISMATCH",
            expected=f"1..{total_outcomes}",
            missing=", ".join(str(value) for value in missing) if missing else "",
            extras=", ".join(str(value) for value in extras) if extras else "",
        )


def _required_header_index(headers: list[str], expected_header: str) -> int:
    expected_token = normalize(expected_header)
    try:
        return headers.index(expected_token)
    except ValueError as exc:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="SCHEMA_COLUMN_KEY_MISSING",
            field=expected_header,
        ) from exc


def _summary_length_bounds(*, sheet_schema: SheetSchema, summary_col_index: int) -> tuple[int, int]:
    minimum = 0
    maximum = 10_000
    for rule in sheet_schema.validations:
        if int(rule.first_col) != int(summary_col_index):
            continue
        options = rule.options if isinstance(rule.options, dict) else {}
        if normalize(str(options.get("validate", ""))) != "length":
            continue
        raw_minimum = options.get("minimum")
        raw_maximum = options.get("maximum")
        minimum_value = coerce_excel_number(raw_minimum)
        maximum_value = coerce_excel_number(raw_maximum)
        if isinstance(minimum_value, (int, float)) and not isinstance(minimum_value, bool):
            minimum = int(minimum_value)
        if isinstance(maximum_value, (int, float)) and not isinstance(maximum_value, bool):
            maximum = int(maximum_value)
        break
    return minimum, maximum


def _validated_co_number(*, co_raw: object, row_number: int) -> int:
    parsed = coerce_excel_number(co_raw)
    if isinstance(parsed, bool) or not isinstance(parsed, (int, float)):
        raise validation_error_from_key(
            "instructor.validation.co_description_number_positive_int_required",
            code="CO_DESCRIPTION_NUMBER_POSITIVE_INT_REQUIRED",
            row=row_number,
        )
    parsed_float = float(parsed)
    if parsed_float <= 0.0 or int(parsed_float) != parsed_float:
        raise validation_error_from_key(
            "instructor.validation.co_description_number_positive_int_required",
            code="CO_DESCRIPTION_NUMBER_POSITIVE_INT_REQUIRED",
            row=row_number,
        )
    return int(parsed_float)


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

__all__ = ["validate_co_description_workbooks"]
