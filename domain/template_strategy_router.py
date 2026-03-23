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
    CO_REPORT_DIRECT_SHEET_SUFFIX,
    CO_REPORT_INDIRECT_SHEET_SUFFIX,
    LAYOUT_MANIFEST_KEY_SHEETS,
    SYSTEM_LAYOUT_MANIFEST_HASH_HEADER,
    SYSTEM_LAYOUT_MANIFEST_HEADER,
    SYSTEM_LAYOUT_SHEET,
    ID_COURSE_SETUP,
)
from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from common.registry import (
    COURSE_SETUP_SHEET_KEY_COURSE_METADATA,
    SYSTEM_HASH_HEADER_TEMPLATE_HASH,
    SYSTEM_HASH_HEADER_TEMPLATE_ID,
    SYSTEM_HASH_SHEET_NAME,
    get_sheet_name_by_key,
)
from common.utils import coerce_excel_number, normalize
from common.workbook_signing import verify_payload_signature


@dataclass(frozen=True)
class SystemWorkbookPayload:
    template_id: str
    template_hash: str
    manifest: dict[str, Any]


@dataclass(frozen=True)
class FinalReportWorkbookSignature:
    template_id: str
    course_code: str
    total_outcomes: int
    section: str
    direct_sheet_count: int
    indirect_sheet_count: int


class _TemplateStrategy(Protocol):
    @property
    def template_id(self) -> str:
        ...

    def supports_operation(self, operation: str) -> bool:
        ...

    def default_workbook_name(
        self,
        *,
        workbook_kind: str,
        context: Mapping[str, Any] | None,
        fallback: str,
    ) -> str:
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

    def validate_course_details_rules(self, workbook: object, *, context: object) -> None:
        ...

    def extract_marks_template_context(self, workbook: object, *, context: object) -> dict[str, Any]:
        ...

    def write_marks_template_workbook(
        self,
        workbook: object,
        context_data: dict[str, Any],
        *,
        context: object,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, Any]:
        ...

    def validate_filled_marks_manifest_schema(self, workbook: object, manifest: object) -> None:
        ...

    def consume_last_marks_anomaly_warnings(self) -> list[str]:
        ...

    def generate_final_report(
        self,
        filled_marks_path: str | Path,
        output_path: str | Path,
        *,
        cancel_token: CancellationToken | None = None,
    ) -> Path:
        ...

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
        ...

    def validate_final_report_workbook(
        self,
        workbook: Any,
        *,
        template_id: str,
        verify_signature: Any,
    ) -> Any:
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
    if normalize(strategy_template_id) != normalize(template_id):
        raise validation_error_from_key(
            "validation.template.unknown",
            code="UNKNOWN_TEMPLATE",
            template_id=template_id,
            available=", ".join(available_template_ids()),
        )
    return strategy


def _require_operation(strategy: _TemplateStrategy, operation: str) -> None:
    if strategy.supports_operation(operation):
        return
    raise validation_error_from_key(
        "validation.template.validator_missing",
        code="COA_TEMPLATE_VALIDATOR_MISSING",
        template_id=strategy.template_id,
        operation=operation,
    )


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
    _require_operation(strategy, "generate_workbook")
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
    _require_operation(strategy, "validate_workbook")
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
    _require_operation(strategy, "validate_workbooks")
    return strategy.validate_workbooks(
        template_id=template_id,
        workbook_kind=workbook_kind,
        workbook_paths=workbook_paths,
        cancel_token=cancel_token,
        context=context,
    )


def default_workbook_name(
    *,
    template_id: str,
    workbook_kind: str,
    fallback: str,
    context: Mapping[str, Any] | None = None,
) -> str:
    strategy = get_template_strategy(template_id)
    return strategy.default_workbook_name(
        workbook_kind=workbook_kind,
        context=context,
        fallback=fallback,
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


def validate_course_details_rules_by_template(workbook: Any, *, template_id: str) -> None:
    strategy = get_template_strategy(template_id)
    _require_operation(strategy, "validate_course_details_rules")
    strategy.validate_course_details_rules(workbook, context=None)


def extract_marks_template_context_by_template(workbook: Any, *, template_id: str) -> dict[str, Any]:
    strategy = get_template_strategy(template_id)
    _require_operation(strategy, "extract_marks_template_context")
    return strategy.extract_marks_template_context(workbook, context=None)


def write_marks_template_workbook_by_template(
    workbook: Any,
    context_data: dict[str, Any],
    *,
    template_id: str,
    cancel_token: CancellationToken | None = None,
) -> dict[str, Any]:
    strategy = get_template_strategy(template_id)
    _require_operation(strategy, "write_marks_template_workbook")
    return strategy.write_marks_template_workbook(
        workbook,
        context_data,
        context=None,
        cancel_token=cancel_token,
    )


def validate_filled_marks_manifest_schema_by_template(
    workbook: Any,
    manifest: object,
    *,
    template_id: str,
) -> None:
    strategy = get_template_strategy(template_id)
    _require_operation(strategy, "validate_filled_marks_manifest_schema")
    strategy.validate_filled_marks_manifest_schema(workbook, manifest)


def consume_marks_anomaly_warnings_by_template(*, template_id: str) -> list[str]:
    strategy = get_template_strategy(template_id)
    _require_operation(strategy, "consume_last_marks_anomaly_warnings")
    return strategy.consume_last_marks_anomaly_warnings()


def generate_final_report_by_template(
    *,
    template_id: str,
    filled_marks_path: str | Path,
    output_path: str | Path,
    cancel_token: CancellationToken | None = None,
) -> Path:
    result = generate_workbook(
        template_id=template_id,
        output_path=output_path,
        workbook_name=Path(output_path).name,
        workbook_kind="final_report",
        cancel_token=cancel_token,
        context={"filled_marks_path": str(filled_marks_path)},
    )
    if isinstance(result, Path):
        return result
    output = getattr(result, "output_path", None)
    if isinstance(output, Path):
        return output
    if isinstance(output, str) and output.strip():
        return Path(output)
    return Path(output_path)


def generate_co_attainment_by_template(
    *,
    template_id: str,
    source_paths: list[Path],
    output_path: Path,
    token: CancellationToken,
    thresholds: tuple[float, float, float] | None = None,
    co_attainment_percent: float | None = None,
    co_attainment_level: int | None = None,
) -> object:
    return generate_workbook(
        template_id=template_id,
        output_path=output_path,
        workbook_name=Path(output_path).name,
        workbook_kind="co_attainment",
        cancel_token=token,
        context={
            "source_paths": [str(path) for path in source_paths],
            "thresholds": tuple(thresholds) if thresholds is not None else None,
            "co_attainment_percent": co_attainment_percent,
            "co_attainment_level": co_attainment_level,
        },
    )


def _count_co_sheets_from_manifest(manifest: dict[str, Any]) -> tuple[int, int] | None:
    sheets = manifest.get(LAYOUT_MANIFEST_KEY_SHEETS, [])
    if not isinstance(sheets, list):
        return None
    direct = 0
    indirect = 0
    for entry in sheets:
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("name") or "").strip()
        if not name:
            continue
        if name.endswith(CO_REPORT_DIRECT_SHEET_SUFFIX):
            direct += 1
        elif name.endswith(CO_REPORT_INDIRECT_SHEET_SUFFIX):
            indirect += 1
    if direct <= 0 and indirect <= 0:
        return None
    return direct, indirect


def read_layout_manifest_co_sheet_counts(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> tuple[int, int] | None:
    try:
        payload = read_valid_system_workbook_payload(workbook, verify_signature=verify_signature)
    except Exception:
        return None
    return _count_co_sheets_from_manifest(payload.manifest)


def read_course_metadata_signature(
    workbook: Any,
    *,
    course_code_key: str,
    total_outcomes_key: str,
    section_key: str,
) -> tuple[str, int, str] | None:
    template_id = read_template_id_from_system_hash_sheet_if_valid(workbook)
    if not template_id:
        return None
    metadata_sheet_name = get_sheet_name_by_key(template_id, COURSE_SETUP_SHEET_KEY_COURSE_METADATA)
    if metadata_sheet_name not in getattr(workbook, "sheetnames", []):
        return None
    sheet = workbook[metadata_sheet_name]

    metadata: dict[str, str] = {}
    row = 2
    while True:
        key_raw = sheet.cell(row=row, column=1).value
        value_raw = sheet.cell(row=row, column=2).value
        if normalize(key_raw) == "" and normalize(value_raw) == "":
            break
        key = normalize(key_raw)
        if key:
            value = coerce_excel_number(value_raw)
            metadata[key] = str(value).strip() if value is not None else ""
        row += 1

    course_code = metadata.get(normalize(course_code_key), "").strip()
    section = metadata.get(normalize(section_key), "").strip()
    total_token = metadata.get(normalize(total_outcomes_key), "").strip()
    if not course_code or not section or not total_token:
        return None
    try:
        total_outcomes = int(float(total_token))
    except (TypeError, ValueError):
        return None
    if total_outcomes <= 0:
        return None
    return course_code, total_outcomes, section


def extract_final_report_signature_from_path(
    path: Path,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> FinalReportWorkbookSignature | None:
    try:
        import openpyxl
    except Exception:
        return None
    try:
        workbook = openpyxl.load_workbook(path, data_only=False, read_only=True)
    except Exception:
        return None
    try:
        template_id = read_template_id_from_system_hash_sheet_if_valid(
            workbook,
            verify_signature=verify_signature,
        )
        if not template_id:
            return None
        sheet_counts = read_layout_manifest_co_sheet_counts(workbook, verify_signature=verify_signature)
        if sheet_counts is None:
            return None
        metadata = read_course_metadata_signature(
            workbook,
            course_code_key="Course_Code",
            total_outcomes_key="Total_Outcomes",
            section_key="Section",
        )
        if metadata is None:
            return None
        course_code, total_outcomes, section = metadata
        direct_sheet_count, indirect_sheet_count = sheet_counts
        return FinalReportWorkbookSignature(
            template_id=template_id,
            course_code=course_code,
            total_outcomes=total_outcomes,
            section=section,
            direct_sheet_count=direct_sheet_count,
            indirect_sheet_count=indirect_sheet_count,
        )
    finally:
        workbook.close()


def validate_final_report_workbook_by_template(
    workbook: Any,
    *,
    template_id: str,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> FinalReportWorkbookSignature | None:
    strategy = get_template_strategy(template_id)
    _require_operation(strategy, "validate_final_report_workbook")
    signature = strategy.validate_final_report_workbook(
        workbook,
        template_id=template_id,
        verify_signature=verify_signature,
    )
    if isinstance(signature, FinalReportWorkbookSignature):
        return signature
    return None


__all__ = [
    "FinalReportWorkbookSignature",
    "SystemWorkbookPayload",
    "available_template_ids",
    "consume_marks_anomaly_warnings_by_template",
    "default_workbook_name",
    "extract_final_report_signature_from_path",
    "extract_marks_template_context_by_template",
    "generate_co_attainment_by_template",
    "generate_final_report_by_template",
    "generate_workbook",
    "validate_workbook",
    "validate_workbooks",
    "get_template_strategy",
    "read_course_metadata_signature",
    "read_layout_manifest_co_sheet_counts",
    "read_template_id_from_system_hash_sheet_if_valid",
    "resolve_template_id_from_workbook_path",
    "read_valid_system_workbook_payload",
    "read_valid_template_id_from_system_hash_sheet",
    "validate_course_details_rules_by_template",
    "validate_filled_marks_manifest_schema_by_template",
    "validate_final_report_workbook_by_template",
]
