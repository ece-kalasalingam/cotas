"""Template strategy router for template-aware workflow dispatch."""

from __future__ import annotations

import importlib
import json
import pkgutil
import re
from dataclasses import dataclass
from collections.abc import Sequence
from pathlib import Path
from typing import Any, Callable, Mapping, Protocol

from common.constants import (
    SYSTEM_LAYOUT_MANIFEST_HASH_HEADER,
    SYSTEM_LAYOUT_MANIFEST_HEADER,
    SYSTEM_LAYOUT_SHEET,
)
from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from common.registry import (
    SYSTEM_HASH_HEADER_TEMPLATE_HASH,
    SYSTEM_HASH_HEADER_TEMPLATE_ID,
    SYSTEM_HASH_SHEET_NAME,
)
from common.utils import normalize
from common.workbook_signing import verify_payload_signature


@dataclass(frozen=True)
class SystemWorkbookPayload:
    template_id: str
    template_hash: str
    manifest: dict[str, Any]


class _TemplateStrategy(Protocol):
    @property
    def template_id(self) -> str:
        ...

    def supports_operation(self, operation: str) -> bool:
        ...

    def generate_workbook(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        output_path: str | Path,
        workbook_name: str | None,
        cancel_token: CancellationToken | None,
        context: Mapping[str, Any] | None,
    ) -> object:
        ...

    def validate_workbook(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_path: str | Path,
        cancel_token: CancellationToken | None,
        context: Mapping[str, Any] | None,
    ) -> str:
        ...

    def validate_workbooks(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_paths: Sequence[str | Path],
        cancel_token: CancellationToken | None,
        context: Mapping[str, Any] | None,
    ) -> dict[str, object]:
        ...

    def generate_workbooks(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_paths: Sequence[str | Path],
        output_dir: str | Path,
        cancel_token: CancellationToken | None,
        context: Mapping[str, Any] | None,
    ) -> dict[str, object]:
        ...

_VERIFY_SIGNATURE = Callable[[str, str], bool]
_TOKEN_RE = re.compile(r"[A-Za-z0-9]+")


def _tokenize_template_id(template_id: str) -> list[str]:
    return [token for token in _TOKEN_RE.findall(str(template_id or "").strip()) if token]


def _strategy_names_from_template_id(template_id: str) -> tuple[str, str]:
    tokens = _tokenize_template_id(template_id)
    if not tokens:
        return "", ""
    module_name = "_".join(token.lower() for token in tokens)
    class_name = "".join(token.capitalize() for token in tokens) + "Strategy"
    return module_name, class_name


def available_template_ids() -> tuple[str, ...]:
    discovered: list[str] = []
    try:
        package = importlib.import_module("domain.template_versions")
    except Exception:
        return tuple()
    package_path = getattr(package, "__path__", None)
    if package_path is None:
        return tuple()
    for _, module_name, _is_pkg in pkgutil.iter_modules(package_path):
        if module_name.startswith("_"):
            continue
        class_name = "".join(part.capitalize() for part in module_name.split("_")) + "Strategy"
        try:
            module = importlib.import_module(f"domain.template_versions.{module_name}")
            strategy_cls = getattr(module, class_name, None)
            if strategy_cls is None:
                continue
            strategy = strategy_cls()
            template_id = str(getattr(strategy, "template_id", "")).strip()
            if template_id:
                discovered.append(template_id)
        except Exception:
            continue
    unique = sorted({item for item in discovered})
    return tuple(unique)


def assert_template_id_matches(
    *,
    actual_template_id: str,
    expected_template_id: str,
    available: str | None = None,
) -> None:
    if normalize(actual_template_id) == normalize(expected_template_id):
        return
    context: dict[str, Any] = {"template_id": actual_template_id}
    available_value = str(available or "").strip()
    if available_value:
        context["available"] = available_value
    raise validation_error_from_key(
        "validation.template.unknown",
        code="UNKNOWN_TEMPLATE",
        **context,
    )


def get_template_strategy(template_id: str) -> _TemplateStrategy:
    module_name, class_name = _strategy_names_from_template_id(template_id)
    if not module_name or not class_name:
        raise validation_error_from_key(
            "validation.template.unknown",
            code="UNKNOWN_TEMPLATE",
            template_id=template_id,
            available=", ".join(available_template_ids()),
        )
    try:
        module = importlib.import_module(f"domain.template_versions.{module_name}")
        strategy_cls = getattr(module, class_name)
        strategy = strategy_cls()
    except Exception as exc:
        raise validation_error_from_key(
            "validation.template.unknown",
            code="UNKNOWN_TEMPLATE",
            template_id=template_id,
            available=", ".join(available_template_ids()),
        ) from exc
    strategy_template_id = str(getattr(strategy, "template_id", "")).strip()
    assert_template_id_matches(
        actual_template_id=strategy_template_id,
        expected_template_id=template_id,
        available=", ".join(available_template_ids()),
    )
    return strategy


def generate_workbook(
    *,
    template_id: str,
    output_path: str | Path,
    workbook_name: str | None = None,
    workbook_kind: str = "course_details_template",
    cancel_token: CancellationToken | None = None,
    context: Mapping[str, Any] | None = None,
) -> object:
    strategy = get_template_strategy(template_id)
    return strategy.generate_workbook(
        template_id=template_id,
        workbook_kind=workbook_kind,
        output_path=output_path,
        workbook_name=workbook_name,
        cancel_token=cancel_token,
        context=context,
    )


def validate_workbook(
    *,
    workbook_path: str | Path,
    workbook_kind: str = "course_details",
    cancel_token: CancellationToken | None = None,
    context: Mapping[str, Any] | None = None,
) -> str:
    template_id = resolve_template_id_from_workbook_path(workbook_path)
    strategy = get_template_strategy(template_id)
    return strategy.validate_workbook(
        template_id=template_id,
        workbook_kind=workbook_kind,
        workbook_path=workbook_path,
        cancel_token=cancel_token,
        context=context,
    )


def validate_workbooks(
    *,
    template_id: str,
    workbook_paths: Sequence[str | Path],
    workbook_kind: str = "course_details",
    cancel_token: CancellationToken | None = None,
    context: Mapping[str, Any] | None = None,
) -> dict[str, object]:
    strategy = get_template_strategy(template_id)
    return strategy.validate_workbooks(
        template_id=template_id,
        workbook_kind=workbook_kind,
        workbook_paths=workbook_paths,
        cancel_token=cancel_token,
        context=context,
    )


def generate_workbooks(
    *,
    template_id: str,
    workbook_paths: Sequence[str | Path],
    output_dir: str | Path,
    workbook_kind: str = "marks_template",
    cancel_token: CancellationToken | None = None,
    context: Mapping[str, Any] | None = None,
) -> dict[str, object]:
    strategy = get_template_strategy(template_id)
    return strategy.generate_workbooks(
        template_id=template_id,
        workbook_kind=workbook_kind,
        workbook_paths=workbook_paths,
        output_dir=output_dir,
        cancel_token=cancel_token,
        context=context,
    )


def resolve_template_id_from_workbook_path(workbook_path: str | Path) -> str:
    try:
        import openpyxl
    except ModuleNotFoundError as exc:
        raise validation_error_from_key(
            "validation.dependency.openpyxl_missing",
            code="OPENPYXL_MISSING",
        ) from exc
    source = Path(workbook_path)
    try:
        workbook = openpyxl.load_workbook(source, data_only=False, read_only=True)
    except Exception as exc:
        raise validation_error_from_key(
            "validation.workbook.open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=str(source),
        ) from exc
    try:
        return read_valid_template_id_from_system_hash_sheet(workbook)
    finally:
        workbook.close()


def read_template_id_from_system_hash_sheet_if_valid(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> str | None:
    if SYSTEM_HASH_SHEET_NAME not in getattr(workbook, "sheetnames", []):
        return None
    sheet = workbook[SYSTEM_HASH_SHEET_NAME]
    header_template_id = normalize(sheet.cell(row=1, column=1).value)
    header_template_hash = normalize(sheet.cell(row=1, column=2).value)
    if header_template_id != normalize(SYSTEM_HASH_HEADER_TEMPLATE_ID):
        return None
    if header_template_hash != normalize(SYSTEM_HASH_HEADER_TEMPLATE_HASH):
        return None
    template_id = str(sheet.cell(row=2, column=1).value or "").strip()
    template_hash = str(sheet.cell(row=2, column=2).value or "").strip()
    if not template_id or not template_hash:
        return None
    if not verify_signature(template_id, template_hash):
        return None
    return template_id


def read_valid_template_id_from_system_hash_sheet(workbook: Any) -> str:
    if SYSTEM_HASH_SHEET_NAME not in getattr(workbook, "sheetnames", []):
        raise validation_error_from_key(
            "validation.system.sheet_missing",
            code="COA_SYSTEM_SHEET_MISSING",
            sheet_name=SYSTEM_HASH_SHEET_NAME,
        )
    sheet = workbook[SYSTEM_HASH_SHEET_NAME]
    header_template_id = normalize(sheet.cell(row=1, column=1).value)
    header_template_hash = normalize(sheet.cell(row=1, column=2).value)
    if header_template_id != normalize(SYSTEM_HASH_HEADER_TEMPLATE_ID):
        raise validation_error_from_key(
            "validation.system_hash.header_template_id_missing",
            code="COA_SYSTEM_HASH_HEADER_TEMPLATE_ID_MISSING",
        )
    if header_template_hash != normalize(SYSTEM_HASH_HEADER_TEMPLATE_HASH):
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
    if not verify_payload_signature(template_id, template_hash):
        raise validation_error_from_key(
            "validation.system_hash.mismatch",
            code="COA_SYSTEM_HASH_MISMATCH",
            template_id=template_id,
        )
    get_template_strategy(template_id)
    return template_id


def _read_manifest_sheet_payload(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE,
) -> dict[str, Any]:
    if SYSTEM_LAYOUT_SHEET not in getattr(workbook, "sheetnames", []):
        raise validation_error_from_key(
            "validation.layout.sheet_missing",
            code="COA_LAYOUT_SHEET_MISSING",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    sheet = workbook[SYSTEM_LAYOUT_SHEET]
    manifest_header = normalize(sheet.cell(row=1, column=1).value)
    manifest_hash_header = normalize(sheet.cell(row=1, column=2).value)
    if manifest_header != normalize(SYSTEM_LAYOUT_MANIFEST_HEADER):
        raise validation_error_from_key(
            "validation.layout.header_mismatch",
            code="COA_LAYOUT_HEADER_MISMATCH",
            sheet_name=SYSTEM_LAYOUT_SHEET,
        )
    if manifest_hash_header != normalize(SYSTEM_LAYOUT_MANIFEST_HASH_HEADER):
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
    template_id = read_valid_template_id_from_system_hash_sheet(workbook)
    sheet = workbook[SYSTEM_HASH_SHEET_NAME]
    template_hash = str(sheet.cell(row=2, column=2).value or "").strip()
    manifest = _read_manifest_sheet_payload(workbook, verify_signature=verify_signature)
    return SystemWorkbookPayload(
        template_id=template_id,
        template_hash=template_hash,
        manifest=manifest,
    )



__all__ = [
    "assert_template_id_matches",
    "SystemWorkbookPayload",
    "available_template_ids",
    "generate_workbook",
    "generate_workbooks",
    "validate_workbook",
    "validate_workbooks",
    "get_template_strategy",
    "read_template_id_from_system_hash_sheet_if_valid",
    "resolve_template_id_from_workbook_path",
    "read_valid_system_workbook_payload",
    "read_valid_template_id_from_system_hash_sheet",
]
