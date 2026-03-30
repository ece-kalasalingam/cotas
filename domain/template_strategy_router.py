"""Template strategy router for template-aware workflow dispatch."""

from __future__ import annotations

import importlib
import re
from dataclasses import dataclass
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import Any, Callable, Mapping, Protocol

from common.error_catalog import validation_error_from_key
from common.jobs import CancellationToken
from common.utils import normalize
from common.workbook_integrity import (
    SystemWorkbookPayload,
    read_template_id_from_system_hash_sheet_if_valid as _read_template_id_from_system_hash_sheet_if_valid_integrity,
    read_valid_system_workbook_payload as _read_valid_system_workbook_payload_integrity,
    read_valid_template_id_from_system_hash_sheet as _read_valid_template_id_from_system_hash_sheet_integrity,
    verify_payload_signature,
)


@dataclass(frozen=True)
class WorkbookGenerationResult:
    status: str
    workbook_path: str | None
    output_path: str
    output_url: str
    reason: str | None = None


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
_ACTIVE_TEMPLATE_IDS = ("COURSE_SETUP_V2",)
_ACTIVE_TEMPLATE_IDS_NORMALIZED = frozenset(normalize(item) for item in _ACTIVE_TEMPLATE_IDS)
_SINGLE_GENERATION_KINDS = ("course_details_template", "co_description_template", "co_attainment")
_BATCH_GENERATION_KIND = "marks_template"
_BATCH_VALIDATION_KINDS = frozenset({"course_details", "marks_template"})


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
    for template_id in _ACTIVE_TEMPLATE_IDS:
        module_name, class_name = _strategy_names_from_template_id(template_id)
        if not module_name or not class_name:
            continue
        try:
            module = importlib.import_module(f"domain.template_versions.{module_name}")
            strategy_cls = getattr(module, class_name, None)
            if strategy_cls is None:
                continue
            strategy = strategy_cls()
            resolved_template_id = str(getattr(strategy, "template_id", "")).strip()
            if resolved_template_id:
                discovered.append(resolved_template_id)
        except Exception:
            continue
    return tuple(sorted({item for item in discovered}))


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
    if normalize(template_id) not in _ACTIVE_TEMPLATE_IDS_NORMALIZED:
        raise validation_error_from_key(
            "validation.template.unknown",
            code="UNKNOWN_TEMPLATE",
            template_id=template_id,
            available=", ".join(available_template_ids()),
        )
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
    _assert_workbook_kind_supported(
        workbook_kind=workbook_kind,
        expected_kinds=_SINGLE_GENERATION_KINDS,
        template_id=template_id,
    )
    strategy = get_template_strategy(template_id)
    raw = strategy.generate_workbook(
        template_id=template_id,
        workbook_kind=workbook_kind,
        output_path=output_path,
        workbook_name=workbook_name,
        cancel_token=cancel_token,
        context=context,
    )
    return _normalize_generate_workbook_result(raw=raw, fallback_output=Path(output_path))


def validate_workbooks(
    *,
    template_id: str,
    workbook_paths: Sequence[str | Path],
    workbook_kind: str = "course_details",
    cancel_token: CancellationToken | None = None,
    context: Mapping[str, Any] | None = None,
) -> dict[str, object]:
    _assert_workbook_kind_supported(
        workbook_kind=workbook_kind,
        expected_kinds=_BATCH_VALIDATION_KINDS,
        template_id=template_id,
    )
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
    _assert_workbook_kind_supported(
        workbook_kind=workbook_kind,
        expected_kinds=(_BATCH_GENERATION_KIND,),
        template_id=template_id,
    )
    strategy = get_template_strategy(template_id)
    raw = strategy.generate_workbooks(
        template_id=template_id,
        workbook_kind=workbook_kind,
        workbook_paths=workbook_paths,
        output_dir=output_dir,
        cancel_token=cancel_token,
        context=context,
    )
    return _normalize_generate_workbooks_result(raw)


def _extract_output_path_from_result(raw: object) -> str | None:
    if isinstance(raw, Path):
        return str(raw)
    if isinstance(raw, str) and raw.strip():
        return raw.strip()
    output = getattr(raw, "workbook_path", None)
    if isinstance(output, Path):
        return str(output)
    if isinstance(output, str) and output.strip():
        return output.strip()
    output = getattr(raw, "output_path", None)
    if isinstance(output, Path):
        return str(output)
    if isinstance(output, str) and output.strip():
        return output.strip()
    return None


def _assert_workbook_kind_supported(
    *,
    workbook_kind: str,
    expected_kinds: Iterable[str],
    template_id: str,
) -> None:
    normalized_kind = normalize(workbook_kind)
    normalized_expected = {normalize(item) for item in expected_kinds}
    if normalized_kind in normalized_expected:
        return
    expected_label = ", ".join(expected_kinds)
    raise validation_error_from_key(
        "common.validation_failed_invalid_data",
        code="WORKBOOK_KIND_UNSUPPORTED",
        workbook_kind=workbook_kind,
        template_id=template_id,
        expected=expected_label,
    )


def _to_int_with_default(value: object, *, default: int = 0) -> int:
    candidate = value or default
    if isinstance(candidate, bool):
        return default
    if isinstance(candidate, int):
        return candidate
    if isinstance(candidate, float):
        return int(candidate)
    if isinstance(candidate, str):
        token = candidate.strip()
        if not token:
            return default
        try:
            return int(token)
        except ValueError:
            return default
    return default


def _normalize_generate_workbook_result(
    *,
    raw: object,
    fallback_output: Path,
) -> WorkbookGenerationResult:
    output_value = _extract_output_path_from_result(raw) or str(fallback_output)
    status = str(getattr(raw, "status", "generated")).strip() or "generated"
    reason_attr = getattr(raw, "reason", None)
    reason = str(reason_attr).strip() if isinstance(reason_attr, str) and reason_attr.strip() else None
    return WorkbookGenerationResult(
        status=status,
        workbook_path=output_value if status == "generated" else None,
        output_path=output_value,
        output_url=output_value,
        reason=reason,
    )


def _normalize_generate_workbooks_result(raw: dict[str, object]) -> dict[str, object]:
    if not isinstance(raw, dict):
        return {
            "total": 0,
            "generated": 0,
            "failed": 0,
            "skipped": 0,
            "generated_workbook_paths": [],
            "output_urls": [],
            "results": {},
        }
    raw_results = raw.get("results")
    normalized_results: dict[str, object] = {}
    generated_paths: list[str] = []
    if isinstance(raw_results, dict):
        for key, item in raw_results.items():
            if not isinstance(item, dict):
                continue
            status = str(item.get("status") or "").strip() or "failed"
            output_path = str(item.get("workbook_path") or item.get("output_path") or item.get("output") or "").strip()
            reason_raw = item.get("reason")
            reason = str(reason_raw).strip() if isinstance(reason_raw, str) and reason_raw.strip() else None
            normalized_results[str(key)] = {
                "status": status,
                "source_path": item.get("source_path"),
                "workbook_path": output_path if status == "generated" and output_path else None,
                "output": output_path if output_path else None,
                "output_path": output_path if output_path else None,
                "output_url": output_path if output_path else None,
                "reason": reason,
            }
            if status == "generated" and output_path:
                generated_paths.append(output_path)
    total = _to_int_with_default(raw.get("total", len(normalized_results)))
    generated = _to_int_with_default(raw.get("generated", len(generated_paths)))
    failed = _to_int_with_default(raw.get("failed", 0))
    skipped = _to_int_with_default(raw.get("skipped", 0))
    return {
        "total": total,
        "generated": generated,
        "failed": failed,
        "skipped": skipped,
        "generated_workbook_paths": generated_paths,
        "output_urls": list(generated_paths),
        "results": normalized_results,
    }


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
    return _read_template_id_from_system_hash_sheet_if_valid_integrity(
        workbook,
        verify_signature=verify_signature,
    )


def read_valid_template_id_from_system_hash_sheet(workbook: Any) -> str:
    template_id = _read_valid_template_id_from_system_hash_sheet_integrity(workbook)
    get_template_strategy(template_id)
    return template_id


def read_valid_system_workbook_payload(
    workbook: Any,
    *,
    verify_signature: _VERIFY_SIGNATURE = verify_payload_signature,
) -> SystemWorkbookPayload:
    payload = _read_valid_system_workbook_payload_integrity(
        workbook,
        verify_signature=verify_signature,
    )
    get_template_strategy(payload.template_id)
    return payload



__all__ = [
    "assert_template_id_matches",
    "SystemWorkbookPayload",
    "available_template_ids",
    "generate_workbook",
    "generate_workbooks",
    "validate_workbooks",
    "get_template_strategy",
    "read_template_id_from_system_hash_sheet_if_valid",
    "resolve_template_id_from_workbook_path",
    "read_valid_system_workbook_payload",
    "read_valid_template_id_from_system_hash_sheet",
]
