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
