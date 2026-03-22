"""Version-specific strategy for COURSE_SETUP_V2."""

from __future__ import annotations

import importlib.util
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Mapping, cast

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError
from common.jobs import CancellationToken
from common.utils import normalize

_SUPPORTED_OPERATIONS = frozenset({"generate_workbook"})


@dataclass(slots=True, frozen=True)
class CourseSetupV2Strategy:
    template_id: str = "COURSE_SETUP_V2"

    def supports_operation(self, operation: str) -> bool:
        return str(operation).strip() in _SUPPORTED_OPERATIONS

    def default_workbook_name(
        self,
        *,
        workbook_kind: str,
        context: Mapping[str, Any] | None,
        fallback: str,
    ) -> str:
        return fallback

    def generate_workbook(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        output_path: str | Path,
        workbook_name: str | None,
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> object:
        if normalize(template_id) != normalize(self.template_id):
            raise validation_error_from_key(
                "validation.template.unknown",
                code="UNKNOWN_TEMPLATE",
                template_id=template_id,
            )
        resolved_workbook_name = (workbook_name or Path(output_path).name).strip()
        if not resolved_workbook_name:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_NAME_REQUIRED",
            )
        if normalize(workbook_kind) != "course_details_template":
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_KIND_UNSUPPORTED",
                workbook_kind=workbook_kind,
                template_id=self.template_id,
            )
        generator = _load_course_template_generator()
        return generator(
            output_path=output_path,
            cancel_token=cancel_token,
        )


def _load_course_template_generator() -> Callable[..., Path]:
    impl_path = Path(__file__).parent / "course_setup_v2_impl" / "course_template.py"
    if not impl_path.exists():
        raise ConfigurationError(f"Missing V2 course template implementation: {impl_path}")
    module_name = "domain.template_versions._course_setup_v2_course_template"
    spec = importlib.util.spec_from_file_location(module_name, impl_path)
    if spec is None or spec.loader is None:
        raise ConfigurationError(f"Unable to load V2 course template implementation: {impl_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    fn = getattr(module, "generate_course_details_template", None)
    if not callable(fn):
        raise ConfigurationError("V2 course template implementation missing generate_course_details_template().")
    return cast(Callable[..., Path], fn)


__all__ = ["CourseSetupV2Strategy"]
