"""Central runtime dependency guard helpers."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from importlib import import_module
from importlib.util import find_spec
from typing import Any

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError


@dataclass(frozen=True)
class RuntimeDependencySpec:
    import_name: str
    package_name: str
    error_code: str
    translation_key: str


_RUNTIME_DEPENDENCY_SPECS: tuple[RuntimeDependencySpec, ...] = (
    RuntimeDependencySpec(
        import_name="openpyxl",
        package_name="openpyxl",
        error_code="OPENPYXL_MISSING",
        translation_key="validation.dependency.openpyxl_missing",
    ),
    RuntimeDependencySpec(
        import_name="xlsxwriter",
        package_name="xlsxwriter",
        error_code="XLSXWRITER_MISSING",
        translation_key="validation.dependency.xlsxwriter_missing",
    ),
    RuntimeDependencySpec(
        import_name="docx",
        package_name="python-docx",
        error_code="PYTHON_DOCX_MISSING",
        translation_key="validation.dependency.python_docx_missing",
    ),
)
_RUNTIME_DEPENDENCY_BY_IMPORT = {item.import_name: item for item in _RUNTIME_DEPENDENCY_SPECS}


def runtime_dependency_specs() -> tuple[RuntimeDependencySpec, ...]:
    """Runtime dependency specs."""
    return _RUNTIME_DEPENDENCY_SPECS


def runtime_dependency_spec(import_name: str) -> RuntimeDependencySpec:
    """Runtime dependency spec for one import name."""
    spec = _RUNTIME_DEPENDENCY_BY_IMPORT.get(str(import_name).strip())
    if spec is not None:
        return spec
    raise ConfigurationError(f"Unsupported runtime dependency import name: {import_name}")


def import_runtime_dependency(import_name: str) -> Any:
    """Import dependency or raise catalog-backed validation error."""
    spec = runtime_dependency_spec(import_name)
    try:
        return import_module(spec.import_name)
    except ModuleNotFoundError as exc:
        raise validation_error_from_key(
            spec.translation_key,
            code=spec.error_code,
        ) from exc


def missing_runtime_dependency_specs(
    *,
    import_names: Iterable[str] | None = None,
) -> tuple[RuntimeDependencySpec, ...]:
    """Resolve missing runtime dependency specs."""
    selected_specs = (
        [runtime_dependency_spec(name) for name in import_names]
        if import_names is not None
        else list(_RUNTIME_DEPENDENCY_SPECS)
    )
    missing: list[RuntimeDependencySpec] = []
    for spec in selected_specs:
        if find_spec(spec.import_name) is None:
            missing.append(spec)
    return tuple(missing)


def missing_runtime_dependency_packages(*, import_names: Iterable[str] | None = None) -> tuple[str, ...]:
    """Resolve missing runtime dependency package names."""
    return tuple(
        item.package_name
        for item in missing_runtime_dependency_specs(import_names=import_names)
    )

