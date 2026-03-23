"""Version-specific strategy for COURSE_SETUP_V2."""

from __future__ import annotations

import importlib.util
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Mapping, cast

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError
from common.jobs import CancellationToken
from common.utils import normalize

_SUPPORTED_OPERATIONS = frozenset(
    {
        "generate_workbook",
        "validate_workbook",
        "validate_workbooks",
        "validate_course_details_rules",
        "extract_marks_template_context",
        "write_marks_template_workbook",
    }
)


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
        if normalize(workbook_kind) == "marks_template":
            return _v1_delegate().default_workbook_name(
                workbook_kind=workbook_kind,
                context=context,
                fallback=fallback,
            )
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

        kind = normalize(workbook_kind)
        if kind == "course_details_template":
            generator = _load_course_template_generator()
            return generator(
                output_path=output_path,
                cancel_token=cancel_token,
            )
        if kind == "marks_template":
            source = str((context or {}).get("course_details_path") or "").strip()
            if not source:
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="WORKBOOK_SOURCE_REQUIRED",
                    workbook_kind=workbook_kind,
                )
            generator = _load_marks_template_generator()
            return generator(
                course_details_path=source,
                output_path=output_path,
                cancel_token=cancel_token,
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
        )

    def validate_workbook(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_path: str | Path,
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> str:
        if normalize(template_id) != normalize(self.template_id):
            raise validation_error_from_key(
                "validation.template.unknown",
                code="UNKNOWN_TEMPLATE",
                template_id=template_id,
            )
        if normalize(workbook_kind) not in {"course_details", "course_details_template"}:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_KIND_UNSUPPORTED",
                workbook_kind=workbook_kind,
                template_id=self.template_id,
            )
        validator = _load_course_template_validator()
        return validator(
            workbook_path=workbook_path,
            cancel_token=cancel_token,
        )

    def validate_workbooks(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_paths: Sequence[str | Path],
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> dict[str, object]:
        if normalize(template_id) != normalize(self.template_id):
            raise validation_error_from_key(
                "validation.template.unknown",
                code="UNKNOWN_TEMPLATE",
                template_id=template_id,
            )
        if normalize(workbook_kind) not in {"course_details", "course_details_template"}:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_KIND_UNSUPPORTED",
                workbook_kind=workbook_kind,
                template_id=self.template_id,
            )
        batch_validator = _load_course_template_batch_validator()
        return batch_validator(
            workbook_paths=workbook_paths,
            cancel_token=cancel_token,
        )

    def validate_course_details_rules(self, workbook: object, *, context: object) -> None:
        _v1_delegate().validate_course_details_rules(workbook, context=context)

    def extract_marks_template_context(self, workbook: object, *, context: object) -> dict[str, Any]:
        return _v1_delegate().extract_marks_template_context(workbook, context=context)

    def write_marks_template_workbook(
        self,
        workbook: object,
        context_data: dict[str, Any],
        *,
        context: object,
        cancel_token: CancellationToken | None = None,
    ) -> dict[str, Any]:
        return _v1_delegate().write_marks_template_workbook(
            workbook,
            context_data,
            context=context,
            cancel_token=cancel_token,
        )


def _v1_delegate():
    from domain.template_versions.course_setup_v1 import CourseSetupV1Strategy

    return CourseSetupV1Strategy()


def _load_marks_template_generator() -> Callable[..., Path]:
    impl_path = Path(__file__).parent / "course_setup_v2_impl" / "marks_template.py"
    if not impl_path.exists():
        raise ConfigurationError(f"Missing V2 marks template implementation: {impl_path}")
    module_name = "domain.template_versions._course_setup_v2_marks_template"
    spec = importlib.util.spec_from_file_location(module_name, impl_path)
    if spec is None or spec.loader is None:
        raise ConfigurationError(f"Unable to load V2 marks template implementation: {impl_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    fn = getattr(module, "generate_marks_template_from_course_details", None)
    if not callable(fn):
        raise ConfigurationError(
            "V2 marks template implementation missing generate_marks_template_from_course_details()."
        )
    return cast(Callable[..., Path], fn)


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


def _load_course_template_validator() -> Callable[..., str]:
    impl_path = Path(__file__).parent / "course_setup_v2_impl" / "course_template_validator.py"
    if not impl_path.exists():
        raise ConfigurationError(f"Missing V2 course template validator: {impl_path}")
    module_name = "domain.template_versions._course_setup_v2_course_template_validator"
    spec = importlib.util.spec_from_file_location(module_name, impl_path)
    if spec is None or spec.loader is None:
        raise ConfigurationError(f"Unable to load V2 course template validator: {impl_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    fn = getattr(module, "validate_course_details_workbook", None)
    if not callable(fn):
        raise ConfigurationError("V2 course template validator missing validate_course_details_workbook().")
    return cast(Callable[..., str], fn)


def _load_course_template_batch_validator() -> Callable[..., dict[str, object]]:
    impl_path = Path(__file__).parent / "course_setup_v2_impl" / "course_template_validator.py"
    if not impl_path.exists():
        raise ConfigurationError(f"Missing V2 course template validator: {impl_path}")
    module_name = "domain.template_versions._course_setup_v2_course_template_validator"
    spec = importlib.util.spec_from_file_location(module_name, impl_path)
    if spec is None or spec.loader is None:
        raise ConfigurationError(f"Unable to load V2 course template validator: {impl_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    fn = getattr(module, "validate_course_details_workbooks", None)
    if not callable(fn):
        raise ConfigurationError("V2 course template validator missing validate_course_details_workbooks().")
    return cast(Callable[..., dict[str, object]], fn)


__all__ = ["CourseSetupV2Strategy"]
