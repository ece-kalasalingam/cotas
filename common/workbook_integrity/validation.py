"""System-sheet readers and integrity validators."""

from __future__ import annotations

import json
from typing import Any, Callable

from common.error_catalog import validation_error_from_key
from common.registry import (
    SYSTEM_HASH_HEADER_TEMPLATE_HASH,
    SYSTEM_HASH_HEADER_TEMPLATE_ID,
    SYSTEM_HASH_SHEET_NAME,
)
from common.workbook_integrity.constants import (
    SYSTEM_LAYOUT_MANIFEST_HASH_HEADER,
    SYSTEM_LAYOUT_MANIFEST_HEADER,
    SYSTEM_LAYOUT_SHEET,
)
from common.workbook_integrity.contracts import SystemWorkbookPayload
from common.workbook_integrity.signing import verify_payload_signature

_VERIFY_SIGNATURE = Callable[[str, str], bool]


def _normalize_text(value: Any) -> str:
    """Normalize text.
    
    Args:
        value: Parameter value (Any).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    text = str(value or "").strip()
    return " ".join(text.lower().split())


def read_template_id_from_system_hash_sheet_if_valid(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> str | None:
    """Read template id from system hash sheet if valid.
    
    Args:
        workbook: Parameter value (Any).
        verify_signature: Parameter value (_VERIFY_SIGNATURE).
    
    Returns:
        str | None: Return value.
    
    Raises:
        None.
    """
    if SYSTEM_HASH_SHEET_NAME not in getattr(workbook, "sheetnames", []):
        return None
    sheet = workbook[SYSTEM_HASH_SHEET_NAME]
    header_template_id = _normalize_text(sheet.cell(row=1, column=1).value)
    header_template_hash = _normalize_text(sheet.cell(row=1, column=2).value)
    if header_template_id != _normalize_text(SYSTEM_HASH_HEADER_TEMPLATE_ID):
        return None
    if header_template_hash != _normalize_text(SYSTEM_HASH_HEADER_TEMPLATE_HASH):
        return None
    template_id = str(sheet.cell(row=2, column=1).value or "").strip()
    template_hash = str(sheet.cell(row=2, column=2).value or "").strip()
    if not template_id or not template_hash:
        return None
    if not verify_signature(template_id, template_hash):
        return None
    return template_id


def read_valid_template_id_from_system_hash_sheet(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> str:
    """Read valid template id from system hash sheet.
    
    Args:
        workbook: Parameter value (Any).
        verify_signature: Parameter value (_VERIFY_SIGNATURE).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    if SYSTEM_HASH_SHEET_NAME not in getattr(workbook, "sheetnames", []):
        raise validation_error_from_key(
            "validation.system.sheet_missing",
            code="COA_SYSTEM_SHEET_MISSING",
            sheet_name=SYSTEM_HASH_SHEET_NAME,
        )
    sheet = workbook[SYSTEM_HASH_SHEET_NAME]
    header_template_id = _normalize_text(sheet.cell(row=1, column=1).value)
    header_template_hash = _normalize_text(sheet.cell(row=1, column=2).value)
    if header_template_id != _normalize_text(SYSTEM_HASH_HEADER_TEMPLATE_ID):
        raise validation_error_from_key(
            "validation.system_hash.header_template_id_missing",
            code="COA_SYSTEM_HASH_HEADER_TEMPLATE_ID_MISSING",
        )
    if header_template_hash != _normalize_text(SYSTEM_HASH_HEADER_TEMPLATE_HASH):
        raise validation_error_from_key(
            "validation.system_hash.header_template_hash_missing",
            code="COA_SYSTEM_HASH_HEADER_TEMPLATE_HASH_MISSING",
        )
    template_id = str(sheet.cell(row=2, column=1).value or "").strip()
    template_hash = str(sheet.cell(row=2, column=2).value or "").strip()
    if not template_id:
        raise validation_error_from_key(
            "validation.system_hash.template_id_missing",
            code="COA_SYSTEM_HASH_TEMPLATE_ID_MISSING",
        )
    if not verify_signature(template_id, template_hash):
        raise validation_error_from_key(
            "validation.system_hash.mismatch",
            code="COA_SYSTEM_HASH_MISMATCH",
            template_id=template_id,
        )
    return template_id


def read_layout_manifest_sheet_payload(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> dict[str, Any]:
    """Read layout manifest sheet payload.
    
    Args:
        workbook: Parameter value (Any).
        verify_signature: Parameter value (_VERIFY_SIGNATURE).
    
    Returns:
        dict[str, Any]: Return value.
    
    Raises:
        None.
    """
    if SYSTEM_LAYOUT_SHEET not in getattr(workbook, "sheetnames", []):
        raise validation_error_from_key(
            "validation.layout.sheet_missing",
            code="COA_LAYOUT_SHEET_MISSING",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    sheet = workbook[SYSTEM_LAYOUT_SHEET]
    manifest_header = _normalize_text(sheet.cell(row=1, column=1).value)
    manifest_hash_header = _normalize_text(sheet.cell(row=1, column=2).value)
    if manifest_header != _normalize_text(SYSTEM_LAYOUT_MANIFEST_HEADER):
        raise validation_error_from_key(
            "validation.layout.header_mismatch",
            code="COA_LAYOUT_HEADER_MISMATCH",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    if manifest_hash_header != _normalize_text(SYSTEM_LAYOUT_MANIFEST_HASH_HEADER):
        raise validation_error_from_key(
            "validation.layout.header_mismatch",
            code="COA_LAYOUT_HEADER_MISMATCH",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    manifest_text = str(sheet.cell(row=2, column=1).value or "").strip()
    manifest_hash = str(sheet.cell(row=2, column=2).value or "").strip()
    if not manifest_text or not manifest_hash:
        raise validation_error_from_key(
            "validation.layout.manifest_missing",
            code="COA_LAYOUT_MANIFEST_MISSING",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    if not verify_signature(manifest_text, manifest_hash):
        raise validation_error_from_key(
            "validation.layout.hash_mismatch",
            code="COA_LAYOUT_HASH_MISMATCH",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    try:
        manifest = json.loads(manifest_text)
    except Exception as exc:
        raise validation_error_from_key(
            "validation.layout.manifest_json_invalid",
            code="COA_LAYOUT_MANIFEST_JSON_INVALID",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        ) from exc
    if not isinstance(manifest, dict):
        raise validation_error_from_key(
            "validation.layout.manifest_json_invalid",
            code="COA_LAYOUT_MANIFEST_JSON_INVALID",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    return manifest


def read_valid_system_workbook_payload(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> SystemWorkbookPayload:
    """Read valid system workbook payload.
    
    Args:
        workbook: Parameter value (Any).
        verify_signature: Parameter value (_VERIFY_SIGNATURE).
    
    Returns:
        SystemWorkbookPayload: Return value.
    
    Raises:
        None.
    """
    template_id = read_valid_template_id_from_system_hash_sheet(
        workbook,
        verify_signature=verify_signature,
    )
    sheet = workbook[SYSTEM_HASH_SHEET_NAME]
    template_hash = str(sheet.cell(row=2, column=2).value or "").strip()
    manifest = read_layout_manifest_sheet_payload(workbook, verify_signature=verify_signature)
    return SystemWorkbookPayload(
        template_id=template_id,
        template_hash=template_hash,
        manifest=manifest,
    )


__all__ = [
    "read_layout_manifest_sheet_payload",
    "read_template_id_from_system_hash_sheet_if_valid",
    "read_valid_system_workbook_payload",
    "read_valid_template_id_from_system_hash_sheet",
]
