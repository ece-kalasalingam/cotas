"""Workbook integrity package: signing, system-sheet IO, and validation."""

from common.workbook_integrity.constants import (
    SYSTEM_LAYOUT_MANIFEST_HASH_HEADER,
    SYSTEM_LAYOUT_MANIFEST_HEADER,
    SYSTEM_LAYOUT_SHEET,
)
from common.workbook_integrity.contracts import SystemWorkbookPayload
from common.workbook_integrity.signing import sign_payload, verify_payload_signature
from common.workbook_integrity.system_sheets import (
    add_system_hash_sheet,
    add_system_layout_sheet,
    copy_system_hash_sheet,
)
from common.workbook_integrity.validation import (
    read_template_id_from_system_hash_sheet_if_valid,
    read_valid_system_workbook_payload,
    read_valid_template_id_from_system_hash_sheet,
)
from common.workbook_integrity.workbook_secret import (
    WORKBOOK_SIGNATURE_VERSION,
    ensure_workbook_secret_policy,
    get_workbook_password,
)

__all__ = [
    "SystemWorkbookPayload",
    "SYSTEM_LAYOUT_SHEET",
    "SYSTEM_LAYOUT_MANIFEST_HEADER",
    "SYSTEM_LAYOUT_MANIFEST_HASH_HEADER",
    "WORKBOOK_SIGNATURE_VERSION",
    "add_system_hash_sheet",
    "add_system_layout_sheet",
    "copy_system_hash_sheet",
    "ensure_workbook_secret_policy",
    "get_workbook_password",
    "read_template_id_from_system_hash_sheet_if_valid",
    "read_valid_system_workbook_payload",
    "read_valid_template_id_from_system_hash_sheet",
    "sign_payload",
    "verify_payload_signature",
]
