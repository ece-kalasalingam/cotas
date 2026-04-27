"""Central validation issue catalog used by engines, UI logging, and toast messages."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any, Literal

from common.exceptions import ValidationError
from common.i18n import t

ValidationSeverity = Literal["info", "success", "warning", "error"]


@dataclass(frozen=True)
class ValidationIssueSpec:
    """Static catalog entry for a validation issue code."""

    code: str
    translation_key: str
    category: str
    severity: ValidationSeverity = "error"
    default_message: str = "Validation failed due to invalid data."


@dataclass(frozen=True)
class ResolvedValidationIssue:
    """Resolved validation issue ready for UI/log rendering."""

    code: str
    category: str
    severity: ValidationSeverity
    translation_key: str
    message: str
    context: dict[str, Any]


_DEFAULT_TRANSLATION_KEY = "common.validation_failed_invalid_data"
_CTX_TRANSLATION_KEY = "__translation_key"
_CTX_CATEGORY = "__category"
_CTX_SEVERITY = "__severity"
_DEFAULT_SPEC = ValidationIssueSpec(
    code="VALIDATION_ERROR",
    translation_key=_DEFAULT_TRANSLATION_KEY,
    category="validation",
)

_VALIDATION_ISSUE_CATALOG: Mapping[str, ValidationIssueSpec] = {
    "OPENPYXL_MISSING": ValidationIssueSpec(
        code="OPENPYXL_MISSING",
        translation_key="validation.dependency.openpyxl_missing",
        category="dependency",
    ),
    "XLSXWRITER_MISSING": ValidationIssueSpec(
        code="XLSXWRITER_MISSING",
        translation_key="validation.dependency.xlsxwriter_missing",
        category="dependency",
    ),
    "PYTHON_DOCX_MISSING": ValidationIssueSpec(
        code="PYTHON_DOCX_MISSING",
        translation_key="validation.dependency.python_docx_missing",
        category="dependency",
    ),
    "WORKBOOK_NOT_FOUND": ValidationIssueSpec(
        code="WORKBOOK_NOT_FOUND",
        translation_key="validation.workbook.not_found",
        category="workbook_io",
    ),
    "WORKBOOK_OPEN_FAILED": ValidationIssueSpec(
        code="WORKBOOK_OPEN_FAILED",
        translation_key="validation.workbook.open_failed",
        category="workbook_io",
    ),
    "UNKNOWN_TEMPLATE": ValidationIssueSpec(
        code="UNKNOWN_TEMPLATE",
        translation_key="validation.template.unknown",
        category="template",
    ),
    "COA_SYSTEM_SHEET_MISSING": ValidationIssueSpec(
        code="COA_SYSTEM_SHEET_MISSING",
        translation_key="validation.system.sheet_missing",
        category="system_hash",
    ),
    "COA_SYSTEM_HASH_HEADER_TEMPLATE_ID_MISSING": ValidationIssueSpec(
        code="COA_SYSTEM_HASH_HEADER_TEMPLATE_ID_MISSING",
        translation_key="validation.system_hash.header_template_id_missing",
        category="system_hash",
    ),
    "COA_SYSTEM_HASH_HEADER_TEMPLATE_HASH_MISSING": ValidationIssueSpec(
        code="COA_SYSTEM_HASH_HEADER_TEMPLATE_HASH_MISSING",
        translation_key="validation.system_hash.header_template_hash_missing",
        category="system_hash",
    ),
    "COA_SYSTEM_HASH_TEMPLATE_ID_MISSING": ValidationIssueSpec(
        code="COA_SYSTEM_HASH_TEMPLATE_ID_MISSING",
        translation_key="validation.system_hash.template_id_missing",
        category="system_hash",
    ),
    "COA_SYSTEM_HASH_MISMATCH": ValidationIssueSpec(
        code="COA_SYSTEM_HASH_MISMATCH",
        translation_key="validation.system_hash.mismatch",
        category="system_hash",
    ),
    "COA_LAYOUT_SHEET_MISSING": ValidationIssueSpec(
        code="COA_LAYOUT_SHEET_MISSING",
        translation_key="validation.layout.sheet_missing",
        category="layout_manifest",
    ),
    "COA_LAYOUT_HEADER_MISMATCH": ValidationIssueSpec(
        code="COA_LAYOUT_HEADER_MISMATCH",
        translation_key="validation.layout.header_mismatch",
        category="layout_manifest",
    ),
    "COA_LAYOUT_MANIFEST_MISSING": ValidationIssueSpec(
        code="COA_LAYOUT_MANIFEST_MISSING",
        translation_key="validation.layout.manifest_missing",
        category="layout_manifest",
    ),
    "COA_LAYOUT_HASH_MISMATCH": ValidationIssueSpec(
        code="COA_LAYOUT_HASH_MISMATCH",
        translation_key="validation.layout.hash_mismatch",
        category="layout_manifest",
    ),
    "COA_LAYOUT_MANIFEST_JSON_INVALID": ValidationIssueSpec(
        code="COA_LAYOUT_MANIFEST_JSON_INVALID",
        translation_key="validation.layout.manifest_json_invalid",
        category="layout_manifest",
    ),
    "COA_TEMPLATE_VALIDATOR_MISSING": ValidationIssueSpec(
        code="COA_TEMPLATE_VALIDATOR_MISSING",
        translation_key="validation.template.validator_missing",
        category="template",
    ),
    "COA_TEMPLATE_MIXED": ValidationIssueSpec(
        code="COA_TEMPLATE_MIXED",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="template",
        default_message="Mixed template IDs are not allowed in one CO Analysis run.",
    ),
    "COA_SOURCE_WORKBOOK_REQUIRED": ValidationIssueSpec(
        code="COA_SOURCE_WORKBOOK_REQUIRED",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="workbook_io",
        default_message="At least one source workbook is required for CO Analysis generation.",
    ),
    "COA_MARK_ENTRY_EMPTY": ValidationIssueSpec(
        code="COA_MARK_ENTRY_EMPTY",
        translation_key="validation.mark.entry_empty",
        category="marks",
    ),
    "COA_MARK_VALUE_INVALID": ValidationIssueSpec(
        code="COA_MARK_VALUE_INVALID",
        translation_key="validation.mark.value_invalid",
        category="marks",
    ),
    "COA_MARK_PRECISION_INVALID": ValidationIssueSpec(
        code="COA_MARK_PRECISION_INVALID",
        translation_key="validation.mark.precision_invalid",
        category="marks",
    ),
    "COA_INDIRECT_MARK_INTEGER_REQUIRED": ValidationIssueSpec(
        code="COA_INDIRECT_MARK_INTEGER_REQUIRED",
        translation_key="validation.mark.indirect_integer_required",
        category="marks",
    ),
    "COA_ABSENCE_POLICY_VIOLATION": ValidationIssueSpec(
        code="COA_ABSENCE_POLICY_VIOLATION",
        translation_key="validation.mark.absence_policy_violation",
        category="marks",
    ),
    "FORMULA_NOT_ALLOWED": ValidationIssueSpec(
        code="FORMULA_NOT_ALLOWED",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="Formulas are not allowed in uploaded course template cells.",
    ),
    "CELL_EMPTY_NOT_ALLOWED": ValidationIssueSpec(
        code="CELL_EMPTY_NOT_ALLOWED",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="All required cells must be non-empty in uploaded course template rows.",
    ),
    "SHEET_DATA_REQUIRED": ValidationIssueSpec(
        code="SHEET_DATA_REQUIRED",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="At least one data row is required in each course template sheet.",
    ),
    "PERCENTAGE_NUMERIC_REQUIRED": ValidationIssueSpec(
        code="PERCENTAGE_NUMERIC_REQUIRED",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="Percentage fields must be numeric.",
    ),
    "PERCENTAGE_RANGE_INVALID": ValidationIssueSpec(
        code="PERCENTAGE_RANGE_INVALID",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="Percentage fields must be between 0 and 100.",
    ),
    "INTEGER_VALUE_REQUIRED": ValidationIssueSpec(
        code="INTEGER_VALUE_REQUIRED",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="An integer value is required.",
    ),
    "INTEGER_VALUE_OUT_OF_RANGE": ValidationIssueSpec(
        code="INTEGER_VALUE_OUT_OF_RANGE",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="Integer value is outside the allowed range.",
    ),
    "TEXT_LENGTH_OUT_OF_RANGE": ValidationIssueSpec(
        code="TEXT_LENGTH_OUT_OF_RANGE",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="Text length is outside the allowed range.",
    ),
    "SCHEMA_MISSING": ValidationIssueSpec(
        code="SCHEMA_MISSING",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="configuration",
        default_message="Template schema configuration is missing.",
    ),
    "SCHEMA_COLUMN_KEY_MISSING": ValidationIssueSpec(
        code="SCHEMA_COLUMN_KEY_MISSING",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="configuration",
        default_message="Template schema column mapping is incomplete.",
    ),
    "INDIRECT_TOOL_COUNT_INVALID": ValidationIssueSpec(
        code="INDIRECT_TOOL_COUNT_INVALID",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="Indirect tool count is outside the allowed range.",
    ),
    "QUESTION_MAP_COMPONENT_MISSING": ValidationIssueSpec(
        code="QUESTION_MAP_COMPONENT_MISSING",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        default_message="A direct assessment component is missing in Question Map.",
    ),
    "COURSE_DETAILS_DUPLICATE_PATH": ValidationIssueSpec(
        code="COURSE_DETAILS_DUPLICATE_PATH",
        translation_key="validation.course_details.duplicate_path",
        category="duplicate",
        severity="warning",
        default_message="Duplicate file path skipped.",
    ),
    "COURSE_DETAILS_SECTION_DUPLICATE": ValidationIssueSpec(
        code="COURSE_DETAILS_SECTION_DUPLICATE",
        translation_key="validation.course_details.duplicate_section",
        category="duplicate",
        severity="warning",
        default_message="Duplicate section skipped for same course cohort.",
    ),
    "COURSE_DETAILS_COHORT_MISMATCH": ValidationIssueSpec(
        code="COURSE_DETAILS_COHORT_MISMATCH",
        translation_key="validation.course_details.cohort_mismatch",
        category="validation",
        severity="warning",
        default_message=(
            "File skipped because course cohort metadata does not match "
            "(course code, semester, academic year, total outcomes must match)."
        ),
    ),
    "COURSE_DETAILS_UNEXPECTED_REJECTION": ValidationIssueSpec(
        code="COURSE_DETAILS_UNEXPECTED_REJECTION",
        translation_key="validation.course_details.unexpected_rejection",
        category="validation",
        severity="warning",
        default_message="File skipped due to an unexpected validation failure.",
    ),
    "MARKS_TEMPLATE_DUPLICATE_PATH": ValidationIssueSpec(
        code="MARKS_TEMPLATE_DUPLICATE_PATH",
        translation_key="validation.marks_template.duplicate_path",
        category="duplicate",
        severity="warning",
        default_message="Duplicate file path skipped.",
    ),
    "MARKS_TEMPLATE_SECTION_DUPLICATE": ValidationIssueSpec(
        code="MARKS_TEMPLATE_SECTION_DUPLICATE",
        translation_key="validation.marks_template.duplicate_section",
        category="duplicate",
        severity="warning",
        default_message="Duplicate section skipped for same course cohort.",
    ),
    "MARKS_TEMPLATE_COHORT_MISMATCH": ValidationIssueSpec(
        code="MARKS_TEMPLATE_COHORT_MISMATCH",
        translation_key="validation.marks_template.cohort_mismatch",
        category="validation",
        severity="warning",
        default_message=(
            "File skipped because course cohort metadata does not match "
            "(course code, semester, academic year, total outcomes must match)."
        ),
    ),
    "MARKS_TEMPLATE_STUDENT_REG_DUPLICATE": ValidationIssueSpec(
        code="MARKS_TEMPLATE_STUDENT_REG_DUPLICATE",
        translation_key="validation.marks_template.duplicate_reg_no",
        category="duplicate",
        severity="warning",
        default_message="Duplicate student register number found across uploaded workbooks.",
    ),
    "MARKS_TEMPLATE_UNEXPECTED_REJECTION": ValidationIssueSpec(
        code="MARKS_TEMPLATE_UNEXPECTED_REJECTION",
        translation_key="validation.marks_template.unexpected_rejection",
        category="validation",
        severity="warning",
        default_message="File skipped due to an unexpected validation failure.",
    ),
    "CO_DESCRIPTION_TEMPLATE_DUPLICATE_PATH": ValidationIssueSpec(
        code="CO_DESCRIPTION_TEMPLATE_DUPLICATE_PATH",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="duplicate",
        severity="warning",
        default_message="Duplicate file path skipped.",
    ),
    "CO_DESCRIPTION_TEMPLATE_COHORT_MISMATCH": ValidationIssueSpec(
        code="CO_DESCRIPTION_TEMPLATE_COHORT_MISMATCH",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        severity="warning",
        default_message=(
            "File skipped because course cohort metadata does not match "
            "(course code, semester, academic year, total outcomes must match)."
        ),
    ),
    "CO_DESCRIPTION_TEMPLATE_UNEXPECTED_REJECTION": ValidationIssueSpec(
        code="CO_DESCRIPTION_TEMPLATE_UNEXPECTED_REJECTION",
        translation_key=_DEFAULT_TRANSLATION_KEY,
        category="validation",
        severity="warning",
        default_message="File skipped due to an unexpected validation failure.",
    ),
}


def _normalize_error_code(value: str) -> str:
    """Normalize error code.
    
    Args:
        value: Parameter value (str).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    return str(value or "").strip().upper()


def code_from_translation_key(translation_key: str) -> str:
    """Code from translation key.
    
    Args:
        translation_key: Parameter value (str).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    key = str(translation_key or "").strip().lower()
    if not key:
        return _DEFAULT_SPEC.code
    return key.replace(".", "_").upper()


def infer_validation_category_from_key(translation_key: str) -> str:
    """Infer validation category from key.
    
    Args:
        translation_key: Parameter value (str).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    key = str(translation_key or "").strip().lower()
    if key.startswith("validation.dependency."):
        return "dependency"
    if key.startswith("validation.workbook."):
        return "workbook_io"
    if key.startswith("validation.system_hash.") or key.startswith("validation.system."):
        return "system_hash"
    if key.startswith("validation.layout."):
        return "layout_manifest"
    if key.startswith("validation.template."):
        return "template"
    if key.startswith("validation.mark."):
        return "marks"
    return "validation"


def validation_error_from_key(
    translation_key: str,
    /,
    *,
    code: str | None = None,
    category: str | None = None,
    severity: ValidationSeverity = "error",
    context: Mapping[str, Any] | None = None,
    **context_kwargs: Any,
) -> ValidationError:
    """Validation error from key.
    
    Args:
        translation_key: Parameter value (str).
        code: Parameter value (str | None).
        category: Parameter value (str | None).
        severity: Parameter value (ValidationSeverity).
        context: Parameter value (Mapping[str, Any] | None).
        context_kwargs: Parameter value (Any).
    
    Returns:
        ValidationError: Return value.
    
    Raises:
        None.
    """
    payload = dict(context or {})
    payload.update(context_kwargs)
    payload[_CTX_TRANSLATION_KEY] = translation_key
    payload[_CTX_CATEGORY] = (category or infer_validation_category_from_key(translation_key)).strip() or "validation"
    payload[_CTX_SEVERITY] = severity
    return ValidationError(
        "",
        code=_normalize_error_code(code or code_from_translation_key(translation_key) or _DEFAULT_SPEC.code),
        context=payload,
    )


def _resolve_translated_message(
    key: str,
    *,
    context: Mapping[str, Any],
    fallback_message: str,
) -> str:
    """Resolve translated message.
    
    Args:
        key: Parameter value (str).
        context: Parameter value (Mapping[str, Any]).
        fallback_message: Parameter value (str).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    try:
        rendered = t(key, **dict(context))
        if isinstance(rendered, str) and rendered.strip():
            return rendered
    except Exception:
        return fallback_message.strip() or key
    return fallback_message.strip() or key


def resolve_validation_issue(
    code: str,
    context: Mapping[str, Any] | None = None,
    *,
    fallback_message: str = "",
) -> ResolvedValidationIssue:
    """Resolve validation issue.
    
    Args:
        code: Parameter value (str).
        context: Parameter value (Mapping[str, Any] | None).
        fallback_message: Parameter value (str).
    
    Returns:
        ResolvedValidationIssue: Return value.
    
    Raises:
        None.
    """
    normalized_code = _normalize_error_code(code) or _DEFAULT_SPEC.code
    context_payload = dict(context or {})
    translation_key = str(context_payload.get(_CTX_TRANSLATION_KEY, "") or "").strip()
    dynamic_category = str(context_payload.get(_CTX_CATEGORY, "") or "").strip()
    dynamic_severity = str(context_payload.get(_CTX_SEVERITY, "") or "").strip()
    clean_context = {
        key: value
        for key, value in context_payload.items()
        if key not in {_CTX_TRANSLATION_KEY, _CTX_CATEGORY, _CTX_SEVERITY}
    }
    spec = _VALIDATION_ISSUE_CATALOG.get(normalized_code)
    if not spec and translation_key:
        resolved_category = dynamic_category or infer_validation_category_from_key(translation_key)
        resolved_severity: ValidationSeverity = "error"
        if dynamic_severity in {"info", "success", "warning", "error"}:
            resolved_severity = dynamic_severity  # type: ignore[assignment]
        translated = _resolve_translated_message(
            translation_key,
            context=clean_context,
            fallback_message=fallback_message or _DEFAULT_SPEC.default_message,
        )
        return ResolvedValidationIssue(
            code=normalized_code,
            category=resolved_category,
            severity=resolved_severity,
            translation_key=translation_key,
            message=translated,
            context=clean_context,
        )

    if spec is None:
        default_message = (fallback_message or "").strip()
        if not default_message:
            default_message = _resolve_translated_message(
                _DEFAULT_SPEC.translation_key,
                context=clean_context,
                fallback_message=_DEFAULT_SPEC.default_message,
            )
        return ResolvedValidationIssue(
            code=normalized_code,
            category=_DEFAULT_SPEC.category,
            severity=_DEFAULT_SPEC.severity,
            translation_key=_DEFAULT_SPEC.translation_key,
            message=default_message,
            context=clean_context,
        )

    translated = _resolve_translated_message(
        spec.translation_key,
        context=clean_context,
        fallback_message=fallback_message or spec.default_message,
    )
    return ResolvedValidationIssue(
        code=normalized_code,
        category=spec.category,
        severity=spec.severity,
        translation_key=spec.translation_key,
        message=translated,
        context=clean_context,
    )
