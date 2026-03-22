"""Validation error-code to localized UI message mapping."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from common.texts import t

_VALIDATION_CODE_TO_KEY: Mapping[str, str] = {
    "OPENPYXL_MISSING": "instructor.validation.openpyxl_missing",
    "XLSXWRITER_MISSING": "instructor.validation.xlsxwriter_missing",
    "WORKBOOK_NOT_FOUND": "instructor.validation.workbook_not_found",
    "WORKBOOK_OPEN_FAILED": "instructor.validation.workbook_open_failed",
    "UNKNOWN_TEMPLATE": "instructor.validation.unknown_template",
    "COA_SYSTEM_HASH_MISMATCH": "instructor.validation.system_hash_mismatch",
    "COA_LAYOUT_MANIFEST_MISSING": "instructor.validation.step2.layout_manifest_missing",
    "COA_LAYOUT_HASH_MISMATCH": "instructor.validation.step2.layout_hash_mismatch",
    "COA_LAYOUT_MANIFEST_JSON_INVALID": "instructor.validation.step2.layout_manifest_json_invalid",
    "COA_MARK_ENTRY_EMPTY": "instructor.validation.step2.mark_entry_empty",
    "COA_MARK_VALUE_INVALID": "instructor.validation.step2.mark_value_invalid",
    "COA_MARK_PRECISION_INVALID": "instructor.validation.step2.mark_precision_invalid",
    "COA_INDIRECT_MARK_INTEGER_REQUIRED": "instructor.validation.step2.indirect_mark_must_be_integer",
    "COA_ABSENCE_POLICY_VIOLATION": "instructor.validation.step2.absence_policy_violation",
}


def resolve_validation_error_message(code: str, context: Mapping[str, Any] | None = None) -> str:
    key = _VALIDATION_CODE_TO_KEY.get(code)
    if not key:
        return code
    context_payload = dict(context or {})
    try:
        return t(key, **context_payload)
    except Exception:
        return key
