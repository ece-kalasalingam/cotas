"""COURSE_SETUP_V2 filled-marks manifest validation and anomaly warning state."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

from common.constants import (
    CO_REPORT_ABSENT_TOKEN,
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
from common.error_catalog import resolve_validation_issue, validation_error_from_key
from common.exceptions import ValidationError
from common.jobs import CancellationToken
from common.registry import (
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    get_sheet_name_by_key,
    get_sheet_schema_by_key,
)
from common.runtime_dependency_guard import import_runtime_dependency
from common.utils import assert_not_symlink_path, coerce_excel_number, normalize
from common.workbook_integrity.workbook_signing import sign_payload
from domain.template_versions.course_setup_v2_impl import (
    instructor_engine_sheetops as _sheetops,
)
from domain.template_versions.course_setup_v2_impl.course_template_validator import (
    _validate_course_metadata_rules,
    _validated_non_empty_data_rows,
)
from domain.template_versions.course_setup_v2_impl.validation_batch_runner import (
    BatchValidationAccumulator,
    BatchValidationRunner,
    ValidationRejectionDecision,
)

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


@dataclass(frozen=True, slots=True)
class _MarksWorkbookIdentity:
    template_id: str
    course_code: str
    semester: str
    academic_year: str
    total_outcomes: int
    section: str
    reg_numbers: frozenset[str]

    def cohort_key(self) -> tuple[str, str, str, int]:
        """Build normalized cohort key tuple for cross-workbook comparison.

        Returns:
            Tuple of normalized cohort identity fields.
        """
        return (
            normalize(self.course_code),
            normalize(self.semester),
            normalize(self.academic_year),
            int(self.total_outcomes),
        )


class _ValidationCollector:
    def __init__(self) -> None:
        """Initialize collector for accumulating structured validation issues."""
        self._issues: list[dict[str, object]] = []

    def add(self, exc: ValidationError) -> None:
        """Capture one validation exception as a normalized issue payload.

        Args:
            exc: Validation exception to record.
        """
        self._issues.append(
            _issue_dict(
                code=str(getattr(exc, "code", "VALIDATION_ERROR")),
                context=dict(getattr(exc, "context", {}) or {}),
                fallback_message=str(exc).strip() or "Validation failed.",
            )
        )

    def capture(self, fn, *args, **kwargs) -> Any | None:
        """Run callable and capture validation exceptions as issues.

        Args:
            fn: Callable to execute.
            *args: Positional arguments for callable.
            **kwargs: Keyword arguments for callable.

        Returns:
            Callable result, or `None` when a validation exception is captured.
        """
        try:
            return fn(*args, **kwargs)
        except ValidationError as exc:
            self.add(exc)
            return None

    def capture_ok(self, fn, *args, **kwargs) -> bool:
        """Run callable and report success while capturing validation failures.

        Args:
            fn: Callable to execute.
            *args: Positional arguments for callable.
            **kwargs: Keyword arguments for callable.

        Returns:
            True when callable completed without validation exception.
        """
        try:
            fn(*args, **kwargs)
            return True
        except ValidationError as exc:
            self.add(exc)
            return False

    def raise_if_any(self) -> None:
        """Raise single or aggregated error when issues were collected.

        Raises:
            ValidationError: If one or more issues were recorded.
        """
        if not self._issues:
            return
        if len(self._issues) == 1:
            issue = self._issues[0]
            raw_context = issue.get("context", {})
            context_map: dict[str, Any]
            if isinstance(raw_context, dict):
                context_map = {str(key): value for key, value in raw_context.items()}
            else:
                context_map = {}
            raise ValidationError(
                str(issue.get("message", "Validation failed.")),
                code=str(issue.get("code", "VALIDATION_ERROR")),
                context=context_map,
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="MARKS_TEMPLATE_VALIDATION_FAILED",
            issue_count=len(self._issues),
            issues=list(self._issues),
        )


def _reset_marks_anomaly_warnings() -> None:
    """Clear in-memory anomaly warnings captured during the latest run."""
    _last_marks_anomaly_warnings.clear()


def consume_last_marks_anomaly_warnings() -> list[str]:
    """Return and clear anomaly warnings from the latest validation run.

    Returns:
        Warning messages collected during mark validation.
    """
    warnings = list(_last_marks_anomaly_warnings)
    _last_marks_anomaly_warnings.clear()
    return warnings


def validate_filled_marks_manifest_schema(*, workbook: Any, manifest: Any) -> None:
    """Validate filled-marks workbook against signed layout manifest constraints.

    Args:
        workbook: Open workbook object containing filled marks data.
        manifest: Parsed layout manifest payload from system layout metadata.

    Returns:
        None. Validation succeeds silently.

    Raises:
        ValidationError: If manifest structure, sheet/layout expectations, anchor
            checks, formulas, identity consistency, or mark-entry rules fail.
    """
    _reset_marks_anomaly_warnings()
    if not isinstance(manifest, dict):
        raise validation_error_from_key("instructor.validation.step2.manifest_root_invalid")

    sheet_order = manifest.get(LAYOUT_MANIFEST_KEY_SHEET_ORDER)
    sheet_specs = manifest.get(LAYOUT_MANIFEST_KEY_SHEETS)
    if not isinstance(sheet_order, list) or not isinstance(sheet_specs, list):
        raise validation_error_from_key("instructor.validation.step2.manifest_structure_invalid")

    collector = _ValidationCollector()
    if list(workbook.sheetnames) != sheet_order:
        collector.add(
            validation_error_from_key(
                "instructor.validation.step2.sheet_order_mismatch",
                expected=sheet_order,
                found=list(workbook.sheetnames),
            )
        )

    has_marks_component = False
    baseline_student_hash: str | None = None
    baseline_student_sheet: str | None = None
    for spec in sheet_specs:
        if not isinstance(spec, dict):
            collector.add(validation_error_from_key("instructor.validation.step2.manifest_sheet_spec_invalid"))
            continue
        sheet_name = spec.get(LAYOUT_SHEET_SPEC_KEY_NAME)
        header_row = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADER_ROW)
        headers = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADERS)
        anchors = spec.get(LAYOUT_SHEET_SPEC_KEY_ANCHORS, [])
        formula_anchors = spec.get(LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS, [])
        if not isinstance(sheet_name, str) or sheet_name not in workbook.sheetnames:
            collector.add(
                validation_error_from_key(
                    "instructor.validation.step2.sheet_missing",
                    sheet_name=sheet_name,
                )
            )
            continue
        if not isinstance(header_row, int) or header_row <= 0:
            collector.add(
                validation_error_from_key(
                    "instructor.validation.step2.header_row_invalid",
                    sheet_name=sheet_name,
                    header_row=header_row,
                )
            )
            continue
        if not isinstance(headers, list) or not headers:
            collector.add(
                validation_error_from_key(
                    "instructor.validation.step2.headers_missing",
                    sheet_name=sheet_name,
                )
            )
            continue
        if not isinstance(anchors, list):
            collector.add(
                validation_error_from_key(
                    "instructor.validation.step2.anchor_spec_invalid",
                    sheet_name=sheet_name,
                )
            )
            continue
        if not isinstance(formula_anchors, list):
            collector.add(
                validation_error_from_key(
                    "instructor.validation.step2.formula_anchor_spec_invalid",
                    sheet_name=sheet_name,
                )
            )
            continue

        worksheet = workbook[sheet_name]
        expected_headers = [normalize(value) for value in headers]
        actual_headers = [
            normalize(worksheet.cell(row=header_row, column=col_index + 1).value)
            for col_index in range(len(expected_headers))
        ]
        if actual_headers != expected_headers:
            collector.add(
                validation_error_from_key(
                    "instructor.validation.step2.header_row_mismatch",
                    sheet_name=sheet_name,
                    row=header_row,
                    expected=headers,
                )
            )

        for anchor in anchors:
            if not isinstance(anchor, list) or len(anchor) != 2:
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.anchor_spec_invalid",
                        sheet_name=sheet_name,
                    )
                )
                continue
            cell_ref, expected_value = anchor
            if not isinstance(cell_ref, str) or not cell_ref:
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.anchor_spec_invalid",
                        sheet_name=sheet_name,
                    )
                )
                continue
            actual_value = worksheet[cell_ref].value
            if not _filled_marks_values_match(expected_value, actual_value):
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.anchor_value_mismatch",
                        sheet_name=sheet_name,
                        cell=cell_ref,
                        expected=expected_value,
                        found=actual_value,
                    )
                )
        for formula_anchor in formula_anchors:
            if not isinstance(formula_anchor, list) or len(formula_anchor) != 2:
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.formula_anchor_spec_invalid",
                        sheet_name=sheet_name,
                    )
                )
                continue
            cell_ref, expected_formula = formula_anchor
            if not isinstance(cell_ref, str) or not isinstance(expected_formula, str):
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.formula_anchor_spec_invalid",
                        sheet_name=sheet_name,
                    )
                )
                continue
            actual_formula = worksheet[cell_ref].value
            if _normalized_formula(actual_formula) != _normalized_formula(expected_formula):
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.formula_mismatch",
                        sheet_name=sheet_name,
                        cell=cell_ref,
                    )
                )

        sheet_kind = spec.get(LAYOUT_SHEET_SPEC_KEY_KIND)
        is_mark_component = sheet_kind in _MARK_COMPONENT_SHEET_KINDS
        if is_mark_component:
            has_marks_component = True
            collector.capture(
                _validate_component_structure_snapshot,
                worksheet=worksheet,
                sheet_name=sheet_name,
                sheet_kind=sheet_kind,
                header_row=header_row,
                structure=spec.get(LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE),
                header_count=len(expected_headers),
            )
            actual_student_hash = collector.capture(
                _validate_component_student_identity,
                worksheet=worksheet,
                sheet_name=sheet_name,
                sheet_kind=sheet_kind,
                header_row=header_row,
                expected_student_count=spec.get(LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT),
                expected_student_hash=spec.get(LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH),
            )
            if isinstance(actual_student_hash, str) and baseline_student_hash is None:
                baseline_student_hash = actual_student_hash
                baseline_student_sheet = sheet_name
            elif isinstance(actual_student_hash, str) and actual_student_hash != baseline_student_hash:
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.student_identity_cross_sheet_mismatch",
                        sheet_name=sheet_name,
                        reference_sheet=baseline_student_sheet,
                    )
                )
            collector.capture(
                _validate_non_empty_marks_entries,
                worksheet=worksheet,
                sheet_name=sheet_name,
                sheet_kind=sheet_kind,
                header_count=len(expected_headers),
                header_row=header_row,
            )

    if not has_marks_component:
        collector.add(validation_error_from_key("instructor.validation.step2.no_component_sheets"))
    collector.raise_if_any()


def validate_filled_marks_workbooks(
    workbook_paths: Sequence[str | Path],
    *,
    template_id: str,
    cancel_token: CancellationToken | None = None,
) -> dict[str, object]:
    """Validate a batch of filled marks workbooks for one expected template.

    In addition to per-file validation, this enforces cohort compatibility,
    duplicate section rejection, and cross-workbook register number uniqueness.

    Args:
        workbook_paths: Source workbook paths to validate.
        template_id: Expected template id for all workbook inputs.
        cancel_token: Optional cancellation token for cooperative cancellation.

    Returns:
        Batch validation summary containing valid/invalid/mismatch paths and
        structured rejection issues.

    Raises:
        JobCancelledError: If cancellation is requested during batch processing.
    """
    baseline_cohort: tuple[str, str, str, int] | None = None
    baseline_identity: _MarksWorkbookIdentity | None = None
    seen_sections: set[str] = set()
    seen_reg_numbers: set[str] = set()

    def _validate_path(path: str) -> _MarksWorkbookIdentity:
        """Validate one workbook path and return extracted identity.

        Args:
            path: Workbook path text.

        Returns:
            Parsed marks workbook identity.
        """
        return _validate_filled_marks_workbook_impl(
            workbook_path=path,
            expected_template_id=template_id,
            cancel_token=cancel_token,
        )

    def _on_validated(acc: BatchValidationAccumulator, path: str, identity: _MarksWorkbookIdentity) -> None:
        """Apply cohort/duplication policies and record batch outcome.

        Args:
            acc: Batch validation accumulator.
            path: Workbook path text.
            identity: Parsed workbook identity from validation.
        """
        nonlocal baseline_cohort
        nonlocal baseline_identity
        cohort = identity.cohort_key()
        if baseline_cohort is None:
            baseline_cohort = cohort
            baseline_identity = identity
            seen_sections.add(normalize(identity.section))
            seen_reg_numbers.update(identity.reg_numbers)
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
                code="MARKS_TEMPLATE_COHORT_MISMATCH",
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
        section_key = normalize(identity.section)
        if section_key in seen_sections:
            issue = _issue_dict(
                code="MARKS_TEMPLATE_SECTION_DUPLICATE",
                context={
                    "workbook": path,
                    "section": identity.section,
                },
                fallback_message="Duplicate section skipped for same course cohort.",
            )
            acc.add_rejection(
                path=path,
                issue=issue,
                decision=ValidationRejectionDecision(
                    reason_kind="duplicate_section",
                    mark_invalid=False,
                    mark_duplicate_section=True,
                ),
            )
            return
        duplicated_reg_numbers = sorted(
            reg_no for reg_no in identity.reg_numbers if reg_no in seen_reg_numbers
        )
        if duplicated_reg_numbers:
            issue = _issue_dict(
                code="MARKS_TEMPLATE_STUDENT_REG_DUPLICATE",
                context={
                    "workbook": path,
                    "duplicates": ", ".join(duplicated_reg_numbers[:5]),
                    "count": len(duplicated_reg_numbers),
                },
                fallback_message="Duplicate student register numbers found across workbooks.",
            )
            acc.add_rejection(
                path=path,
                issue=issue,
                decision=ValidationRejectionDecision(
                    reason_kind="duplicate_reg_no",
                    mark_invalid=True,
                ),
            )
            return
        seen_sections.add(section_key)
        seen_reg_numbers.update(identity.reg_numbers)
        acc.add_valid(path=path, template_id=identity.template_id)

    def _classify_validation_error(
        _path: str,
        _exc: ValidationError,
        issue: dict[str, object],
    ) -> ValidationRejectionDecision:
        """Classify validation issue into batch rejection decision semantics.

        Args:
            _path: Workbook path text.
            _exc: Underlying validation exception.
            issue: Resolved issue payload.

        Returns:
            Rejection decision flags used by batch accumulator.
        """
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

    runner = BatchValidationRunner[_MarksWorkbookIdentity](
        issue_builder=_issue_dict,
        duplicate_path_issue_code="MARKS_TEMPLATE_DUPLICATE_PATH",
        unexpected_issue_code="MARKS_TEMPLATE_UNEXPECTED_REJECTION",
    )
    return runner.run(
        workbook_paths=workbook_paths,
        validate_path=_validate_path,
        on_validated=_on_validated,
        cancel_token=cancel_token,
        classify_validation_error=_classify_validation_error,
    )


def _validate_filled_marks_workbook_impl(
    *,
    workbook_path: str | Path,
    expected_template_id: str,
    cancel_token: CancellationToken | None = None,
) -> _MarksWorkbookIdentity:
    """Validate one filled marks workbook and return normalized identity details.

    This performs the two-stage trust flow:
    1) Read-only open for signed template/manifest payload checks.
    2) Full open for manifest-driven sheet and mark-level validation.

    Args:
        workbook_path: Path to the workbook file.
        expected_template_id: Expected template id for the workbook.
        cancel_token: Optional cancellation token for cooperative cancellation.

    Returns:
        Workbook identity used for cohort/section/register consistency checks.

    Raises:
        ValidationError: If dependency, file-open, template/manifest integrity,
            or mark validation checks fail.
        JobCancelledError: If cancellation is requested while validating.
    """
    from domain.template_strategy_router import (
        assert_template_id_matches,
        read_valid_system_workbook_payload,
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
        payload = read_valid_system_workbook_payload(workbook)
    finally:
        workbook.close()

    collector = _ValidationCollector()
    collector.capture(
        assert_template_id_matches,
        actual_template_id=payload.template_id,
        expected_template_id=expected_template_id,
    )
    if cancel_token is not None:
        cancel_token.raise_if_cancelled()

    try:
        workbook = openpyxl.load_workbook(workbook_file, data_only=False)
    except Exception as exc:
        raise validation_error_from_key(
            "validation.workbook.open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(workbook_file),
        ) from exc
    identity: _MarksWorkbookIdentity | None = None
    try:
        manifest_ok = collector.capture_ok(
            validate_filled_marks_manifest_schema,
            workbook=workbook,
            manifest=payload.manifest,
        )
        if manifest_ok:
            captured_identity = collector.capture(
                _read_marks_workbook_identity,
                workbook=workbook,
                template_id=payload.template_id,
            )
            if isinstance(captured_identity, _MarksWorkbookIdentity):
                identity = captured_identity
    finally:
        workbook.close()
    collector.raise_if_any()
    if identity is None:
        raise validation_error_from_key("common.validation_failed_invalid_data", code="MARKS_TEMPLATE_IDENTITY_MISSING")
    return identity


def _read_marks_workbook_identity(*, workbook: Any, template_id: str) -> _MarksWorkbookIdentity:
    """Read cohort metadata identity and register numbers from workbook.

    Args:
        workbook: Open workbook containing marks and metadata sheets.
        template_id: Template id used to resolve metadata sheet schema.

    Returns:
        Parsed workbook identity for cohort-level compatibility checks.

    Raises:
        ValidationError: If metadata schema/sheet/data is missing or invalid.
    """
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
    return _MarksWorkbookIdentity(
        template_id=template_id,
        course_code=identity.course_code,
        semester=identity.semester,
        academic_year=identity.academic_year,
        total_outcomes=identity.total_outcomes,
        section=identity.section,
        reg_numbers=_extract_marks_workbook_reg_numbers(workbook=workbook),
    )


def _extract_marks_workbook_reg_numbers(*, workbook: Any) -> frozenset[str]:
    """Extract normalized register numbers from sheets with reg-no headers.

    Args:
        workbook: Open workbook to scan.

    Returns:
        Frozen set of normalized register numbers.
    """
    reg_header_tokens = {"reg no", "reg_no", "regno"}
    reg_numbers: set[str] = set()
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        max_col = int(getattr(worksheet, "max_column", 0) or 0)
        if max_col <= 0:
            continue
        header_row = 1
        reg_col: int | None = None
        for col in range(1, max_col + 1):
            header_value = normalize(worksheet.cell(row=header_row, column=col).value)
            if header_value in reg_header_tokens:
                reg_col = col
                break
        if reg_col is None:
            continue
        max_row = int(getattr(worksheet, "max_row", 0) or 0)
        for row in range(header_row + 1, max_row + 1):
            raw = worksheet.cell(row=row, column=reg_col).value
            token = normalize(raw)
            if token:
                reg_numbers.add(token)
    return frozenset(reg_numbers)


def _issue_dict(*, code: str, context: dict[str, Any], fallback_message: str) -> dict[str, object]:
    """Build normalized issue payload from issue catalog resolution.

    Args:
        code: Validation issue code.
        context: Structured context fields for issue rendering.
        fallback_message: Message used when lookup is unavailable.

    Returns:
        Issue dictionary for batch validation outputs.
    """
    resolved = resolve_validation_issue(code, context, fallback_message=fallback_message)
    return {
        "code": resolved.code,
        "category": resolved.category,
        "severity": resolved.severity,
        "translation_key": resolved.translation_key,
        "message": resolved.message,
        "context": dict(resolved.context),
    }


def _filled_marks_values_match(expected_value: object, actual_value: object) -> bool:
    """Compare expected and actual values with numeric/text normalization.

    Args:
        expected_value: Expected value from manifest/rule.
        actual_value: Actual value read from worksheet.

    Returns:
        True when values match after normalization.
    """
    expected_coerced = coerce_excel_number(expected_value)
    actual_coerced = coerce_excel_number(actual_value)
    numeric_types = (int, float)
    if isinstance(expected_coerced, numeric_types) and not isinstance(expected_coerced, bool):
        if not isinstance(actual_coerced, numeric_types) or isinstance(actual_coerced, bool):
            return False
        return abs(float(expected_coerced) - float(actual_coerced)) <= 1e-9

    return normalize(expected_coerced) == normalize(actual_coerced)


def _normalized_formula(value: object) -> str:
    """Normalize formula string for stable equality comparisons.

    Args:
        value: Formula or raw cell value.

    Returns:
        Normalized token string with spaces and absolute markers removed.
    """
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
    """Validate component-sheet student identity against manifest expectations.

    Args:
        worksheet: Component worksheet to validate.
        sheet_name: Sheet name for issue context.
        sheet_kind: Component sheet kind discriminator.
        header_row: Header row index.
        expected_student_count: Expected student count from manifest.
        expected_student_hash: Expected identity hash from manifest.

    Returns:
        Computed identity hash from worksheet rows.

    Raises:
        ValidationError: If identity row structure/count/hash is invalid.
    """
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
    """Extract ordered `(reg_no, student_name)` rows from one component sheet.

    Args:
        worksheet: Component worksheet.
        sheet_name: Sheet name for issue context.
        sheet_kind: Component sheet kind discriminator.
        header_row: Header row index.

    Returns:
        Ordered student identity tuples.

    Raises:
        ValidationError: If student identity rows are partial or duplicated.
    """
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
    """Create deterministic signature hash for ordered student tuples.

    Args:
        students: Ordered register number and name tuples.

    Returns:
        Signed hash payload for identity comparison.
    """
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
    """Validate non-empty mark entries, anomalies, and row-total formulas.

    Args:
        worksheet: Component worksheet.
        sheet_name: Sheet name for issue context.
        sheet_kind: Component sheet kind discriminator.
        header_count: Total header column count.
        header_row: Header row index.

    Returns:
        None. Validation succeeds silently.

    Raises:
        ValidationError: If marks, absence policy, or totals are invalid.
    """
    collector = _ValidationCollector()
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
            if not token:
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.mark_entry_empty",
                        code="COA_MARK_ENTRY_EMPTY",
                        sheet_name=sheet_name,
                        cell=cell.coordinate,
                    )
                )
                continue
            if token == normalize(CO_REPORT_ABSENT_TOKEN):
                has_absent = True
                absent_count_by_col[col] += 1
                continue
            has_numeric = True
            numeric_value = coerce_excel_number(cell_value)
            if isinstance(numeric_value, bool) or not isinstance(numeric_value, (int, float)):
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.mark_value_invalid",
                        code="COA_MARK_VALUE_INVALID",
                        sheet_name=sheet_name,
                        cell=cell.coordinate,
                        value=cell_value,
                        minimum=minimum,
                        maximum=maximum_by_col[col],
                    )
                )
                continue
            if not _has_allowed_decimal_precision(float(numeric_value)):
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.mark_precision_invalid",
                        code="COA_MARK_PRECISION_INVALID",
                        sheet_name=sheet_name,
                        cell=cell.coordinate,
                        value=cell_value,
                        decimals=_MAX_DECIMAL_PLACES,
                    )
                )
                continue
            if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT and not _is_integer_value(float(numeric_value)):
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.indirect_mark_must_be_integer",
                        code="COA_INDIRECT_MARK_INTEGER_REQUIRED",
                        sheet_name=sheet_name,
                        cell=cell.coordinate,
                        value=cell_value,
                    )
                )
                continue
            maximum = maximum_by_col[col]
            numeric_float = float(numeric_value)
            if numeric_float < minimum or numeric_float > maximum:
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.mark_value_invalid",
                        code="COA_MARK_VALUE_INVALID",
                        sheet_name=sheet_name,
                        cell=cell.coordinate,
                        value=cell_value,
                        minimum=minimum,
                        maximum=maximum,
                    )
                )
                continue
            numeric_count_by_col[col] += 1
            frequency_by_value = frequency_by_value_by_col[col]
            frequency_by_value[numeric_float] = frequency_by_value.get(numeric_float, 0) + 1
        try:
            _validate_absence_policy_for_row(
                sheet_name=sheet_name,
                worksheet=worksheet,
                sheet_kind=sheet_kind,
                row=row,
                mark_cols=mark_cols,
                has_absent=has_absent,
                has_numeric=has_numeric,
            )
        except ValidationError as exc:
            collector.add(exc)
    _log_marks_anomaly_warnings_from_stats(
        sheet_name=sheet_name,
        mark_cols=mark_cols,
        student_count=student_count,
        absent_count_by_col=absent_count_by_col,
        numeric_count_by_col=numeric_count_by_col,
        frequency_by_value_by_col=frequency_by_value_by_col,
    )
    try:
        _validate_row_total_consistency(
            worksheet=worksheet,
            sheet_name=sheet_name,
            sheet_kind=sheet_kind,
            header_count=header_count,
            header_row=header_row,
            student_count=student_count,
        )
    except ValidationError as exc:
        collector.add(exc)
    collector.raise_if_any()


def _marks_data_start_row(sheet_kind: Any, header_row: int) -> int:
    """Resolve first student-data row index for sheet kind.

    Args:
        sheet_kind: Component sheet kind discriminator.
        header_row: Header row index.

    Returns:
        First row index where student data begins.
    """
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        return header_row + 1
    return header_row + 3


def _marks_entry_columns(sheet_kind: Any, header_count: int) -> range:
    """Resolve mark-entry column range for component sheet kind.

    Args:
        sheet_kind: Component sheet kind discriminator.
        header_count: Total header column count.

    Returns:
        Range of columns containing mark entry cells.

    Raises:
        ValidationError: If sheet kind is unsupported.
    """
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        return range(4, 5)
    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
        return range(4, header_count)
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        return range(4, header_count + 1)
    raise validation_error_from_key("instructor.validation.step2.manifest_sheet_spec_invalid")


def _mark_min_for_sheet(sheet_kind: Any) -> float:
    """Resolve minimum allowed mark value for sheet kind.

    Args:
        sheet_kind: Component sheet kind discriminator.

    Returns:
        Minimum allowed mark value.
    """
    if sheet_kind == LAYOUT_SHEET_KIND_INDIRECT:
        return float(max(MIN_MARK_VALUE, LIKERT_MIN))
    return float(MIN_MARK_VALUE)


def _mark_max_for_cell(worksheet: Any, sheet_kind: Any, max_row: int, col: int) -> float:
    """Resolve maximum allowed mark for given cell context.

    Args:
        worksheet: Component worksheet.
        sheet_kind: Component sheet kind discriminator.
        max_row: Row index containing maxima metadata.
        col: Column index for target mark cell.

    Returns:
        Maximum allowed mark value.

    Raises:
        ValidationError: If maxima metadata is missing or invalid.
    """
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
    """Count contiguous student rows in a component sheet.

    Args:
        worksheet: Component worksheet.
        sheet_kind: Component sheet kind discriminator.
        header_row: Header row index.

    Returns:
        Number of student rows until first fully blank identity row.
    """
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
    """Validate absence-policy constraints for one student row.

    Args:
        sheet_name: Sheet name for issue context.
        worksheet: Component worksheet.
        sheet_kind: Component sheet kind discriminator.
        row: Student row index.
        mark_cols: Mark-entry column range.
        has_absent: Whether row contains absent marker.
        has_numeric: Whether row contains numeric mark.

    Returns:
        None. Validation succeeds silently.

    Raises:
        ValidationError: If absent and numeric values are mixed.
    """
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
    """Validate row-level formula consistency for totals/derived columns.

    Args:
        worksheet: Component worksheet.
        sheet_name: Sheet name for issue context.
        sheet_kind: Component sheet kind discriminator.
        header_count: Total header column count.
        header_row: Header row index.
        student_count: Number of student rows to validate.

    Returns:
        None. Validation succeeds silently.

    Raises:
        ValidationError: If expected formulas are missing or mismatched.
    """
    collector = _ValidationCollector()
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
            expected_sum = _FORMULA_SUM_TEMPLATE.format(
                start=f"{_excel_col_name(first_mark_col)}{row}",
                end=f"{_excel_col_name(last_mark_col)}{row}",
            )
            expected_with_absent = _sheetops._build_total_formula_with_absent(
                first_mark_col_name=_excel_col_name(first_mark_col),
                last_mark_col_name=_excel_col_name(last_mark_col),
                row_1_based=row,
            )
            normalized_actual = _normalized_formula(actual)
            allowed_formulas = {
                _normalized_formula(expected_sum),
                _normalized_formula(expected_with_absent),
            }
            if normalized_actual not in allowed_formulas:
                collector.add(
                    validation_error_from_key(
                        "instructor.validation.step2.total_formula_mismatch",
                        sheet_name=sheet_name,
                        cell=worksheet.cell(row=row, column=total_col).coordinate,
                    )
                )
        collector.raise_if_any()
        return

    if sheet_kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
        for row in range(first_row, last_row + 1):
            for col in range(5, header_count + 1):
                formula = worksheet.cell(row=row, column=col).value
                if not isinstance(formula, str) or not formula.startswith("="):
                    collector.add(
                        validation_error_from_key(
                            "instructor.validation.step2.co_formula_mismatch",
                            sheet_name=sheet_name,
                            cell=worksheet.cell(row=row, column=col).coordinate,
                        )
                    )
        collector.raise_if_any()
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
    """Validate component structure snapshot values from manifest payload.

    Args:
        worksheet: Component worksheet.
        sheet_name: Sheet name for issue context.
        sheet_kind: Component sheet kind discriminator.
        header_row: Header row index.
        structure: Manifest structure payload for the sheet.
        header_count: Total header column count.

    Returns:
        None. Validation succeeds silently.

    Raises:
        ValidationError: If structure snapshot does not match worksheet state.
    """
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
    """Check whether value fits configured decimal-place precision.

    Args:
        value: Numeric value to validate.

    Returns:
        True when value precision is within configured limit.
    """
    scaled = round(value * (10**_MAX_DECIMAL_PLACES))
    return abs(value - (scaled / (10**_MAX_DECIMAL_PLACES))) <= 1e-9


def _is_integer_value(value: float) -> bool:
    """Check whether numeric value is effectively an integer.

    Args:
        value: Numeric value to validate.

    Returns:
        True when value rounds to itself within tolerance.
    """
    return abs(value - round(value)) <= 1e-9


def _excel_col_name(col_index_1_based: int) -> str:
    """Convert 1-based column index to Excel letter label.

    Args:
        col_index_1_based: Column index starting from 1.

    Returns:
        Excel-style column label such as `A`, `Z`, or `AA`.
    """
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
    """Record anomaly warnings from aggregated mark-entry statistics.

    Args:
        sheet_name: Sheet name for warning context.
        mark_cols: Mark-entry columns considered.
        student_count: Number of evaluated students.
        absent_count_by_col: Absent counts by column.
        numeric_count_by_col: Numeric entry counts by column.
        frequency_by_value_by_col: Per-column value frequency mapping.

    Returns:
        None. Appends warning messages to the shared warning buffer.
    """
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
    "validate_filled_marks_workbooks",
    "validate_filled_marks_manifest_schema",
]
