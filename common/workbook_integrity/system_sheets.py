"""System sheet writers/copy helpers for workbook integrity."""

from __future__ import annotations

import json
from typing import Any

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
from common.workbook_integrity.signing import sign_payload


def _serialize_layout_manifest(layout_manifest: dict[str, Any]) -> str:
    """Serialize layout manifest.
    
    Args:
        layout_manifest: Parameter value (dict[str, Any]).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    return json.dumps(layout_manifest, sort_keys=True, separators=(",", ":"), ensure_ascii=True)


def add_system_hash_sheet(workbook: Any, template_id: str) -> None:
    """Add system hash sheet.
    
    Args:
        workbook: Parameter value (Any).
        template_id: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    template_hash = sign_payload(template_id)
    worksheet = workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
    worksheet.write_row(0, 0, [SYSTEM_HASH_HEADER_TEMPLATE_ID, SYSTEM_HASH_HEADER_TEMPLATE_HASH])
    worksheet.write_row(1, 0, [template_id, template_hash])
    worksheet.hide()


def add_system_layout_sheet(workbook: Any, layout_manifest: dict[str, Any]) -> None:
    """Add system layout sheet.
    
    Args:
        workbook: Parameter value (Any).
        layout_manifest: Parameter value (dict[str, Any]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    manifest_text = _serialize_layout_manifest(layout_manifest)
    manifest_hash = sign_payload(manifest_text)
    worksheet = workbook.add_worksheet(SYSTEM_LAYOUT_SHEET)
    worksheet.write_row(
        0,
        0,
        [SYSTEM_LAYOUT_MANIFEST_HEADER, SYSTEM_LAYOUT_MANIFEST_HASH_HEADER],
    )
    worksheet.write_row(1, 0, [manifest_text, manifest_hash])
    worksheet.hide()


def copy_system_hash_sheet(source_workbook: Any, target_workbook: Any) -> None:
    """Copy system hash sheet.
    
    Args:
        source_workbook: Parameter value (Any).
        target_workbook: Parameter value (Any).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    from common.workbook_integrity.validation import (
        read_valid_template_id_from_system_hash_sheet,
    )

    template_id = read_valid_template_id_from_system_hash_sheet(source_workbook)
    source = source_workbook[SYSTEM_HASH_SHEET_NAME]
    template_hash = str(source.cell(row=2, column=2).value or "").strip()
    target = target_workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
    target.write_row(0, 0, [SYSTEM_HASH_HEADER_TEMPLATE_ID, SYSTEM_HASH_HEADER_TEMPLATE_HASH])
    target.write_row(1, 0, [template_id, template_hash])
    target.hide()


__all__ = [
    "add_system_hash_sheet",
    "add_system_layout_sheet",
    "copy_system_hash_sheet",
]
