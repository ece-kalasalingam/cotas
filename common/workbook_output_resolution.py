"""Reusable helpers for multi-workbook output collision handling."""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from common.utils import canonical_path_key, dedupe_paths_by_canonical_key


@dataclass(frozen=True, slots=True)
class OverwriteConflict:
    source_path: str
    output_path: str


@dataclass(frozen=True, slots=True)
class OverwriteResolution:
    retry_sources: list[str]
    output_path_overrides: dict[str, str]


def extract_overwrite_conflicts_from_generation_result(
    result: Mapping[str, object] | None,
) -> list[OverwriteConflict]:
    """Extract overwrite conflicts from generation result.
    
    Args:
        result: Parameter value (Mapping[str, object] | None).
    
    Returns:
        list[OverwriteConflict]: Return value.
    
    Raises:
        None.
    """
    if not isinstance(result, Mapping):
        return []
    raw_results = result.get("results", {})
    if not isinstance(raw_results, Mapping):
        return []
    conflicts: list[OverwriteConflict] = []
    for item in raw_results.values():
        if not isinstance(item, Mapping):
            continue
        if str(item.get("status") or "").strip() != "failed":
            continue
        if str(item.get("reason") or "").strip() != "output_already_exists":
            continue
        source_path = str(item.get("source_path") or "").strip()
        output_path = str(item.get("existing_output_path") or item.get("output_path") or "").strip()
        if not source_path or not output_path:
            continue
        conflicts.append(
            OverwriteConflict(
                source_path=source_path,
                output_path=output_path,
            )
        )
    return conflicts


def resolve_overwrite_conflicts(
    conflicts: Sequence[OverwriteConflict],
    *,
    per_file_native_limit: int,
    ask_overwrite_all: Callable[[list[str]], bool],
    ask_output_path: Callable[[str], str | None],
) -> OverwriteResolution:
    """Resolve overwrite conflicts.
    
    Args:
        conflicts: Parameter value (Sequence[OverwriteConflict]).
        per_file_native_limit: Parameter value (int).
        ask_overwrite_all: Parameter value (Callable[[list[str]], bool]).
        ask_output_path: Parameter value (Callable[[str], str | None]).
    
    Returns:
        OverwriteResolution: Return value.
    
    Raises:
        None.
    """
    unique_conflicts: list[OverwriteConflict] = []
    source_paths = [str(item.source_path or "").strip() for item in conflicts]
    unique_sources, _duplicates = dedupe_paths_by_canonical_key(source_paths, skip_empty=True)
    output_by_source_key: dict[str, str] = {}
    for item in conflicts:
        source_path = str(item.source_path or "").strip()
        output_path = str(item.output_path or "").strip()
        if not source_path or not output_path:
            continue
        output_by_source_key.setdefault(canonical_path_key(source_path), output_path)
    for source_path in unique_sources:
        source_key = canonical_path_key(source_path)
        output_path = output_by_source_key.get(source_key, "").strip()
        if not output_path:
            continue
        unique_conflicts.append(OverwriteConflict(source_path=source_path, output_path=output_path))
    if not unique_conflicts:
        return OverwriteResolution(retry_sources=[], output_path_overrides={})

    if len(unique_conflicts) > per_file_native_limit:
        output_paths = [item.output_path for item in unique_conflicts]
        if not ask_overwrite_all(output_paths):
            return OverwriteResolution(retry_sources=[], output_path_overrides={})
        return OverwriteResolution(
            retry_sources=[item.source_path for item in unique_conflicts],
            output_path_overrides={item.source_path: item.output_path for item in unique_conflicts},
        )

    retry_sources: list[str] = []
    output_path_overrides: dict[str, str] = {}
    for item in unique_conflicts:
        selected_output = str(ask_output_path(item.output_path) or "").strip()
        if not selected_output:
            continue
        retry_sources.append(item.source_path)
        output_path_overrides[item.source_path] = selected_output
    return OverwriteResolution(
        retry_sources=retry_sources,
        output_path_overrides=output_path_overrides,
    )


__all__ = [
    "OverwriteConflict",
    "OverwriteResolution",
    "extract_overwrite_conflicts_from_generation_result",
    "resolve_overwrite_conflicts",
]
