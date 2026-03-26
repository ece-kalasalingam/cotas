"""Version-specific strategy for COURSE_SETUP_V2."""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, cast

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError
from common.jobs import CancellationToken
from common.utils import normalize

_SUPPORTED_OPERATIONS = frozenset(
    {
        "generate_workbook",
        "generate_workbooks",
        "validate_workbook",
        "validate_workbooks",
    }
)
_SUPPORTED_OPERATION_TOKENS = frozenset(normalize(value) for value in _SUPPORTED_OPERATIONS)


@dataclass(slots=True, frozen=True)
class CourseSetupV2Strategy:
    template_id: str = "COURSE_SETUP_V2"

    def supports_operation(self, operation: str) -> bool:
        return normalize(operation) in _SUPPORTED_OPERATION_TOKENS

    def default_workbook_name(
        self,
        *,
        workbook_kind: str,
        context: Mapping[str, Any] | None,
        fallback: str,
    ) -> str:
        del workbook_kind
        del context
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
        from domain.template_strategy_router import assert_template_id_matches

        assert_template_id_matches(
            actual_template_id=template_id,
            expected_template_id=self.template_id,
        )
        resolved_workbook_name = (workbook_name or Path(output_path).name).strip()
        if not resolved_workbook_name:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_NAME_REQUIRED",
            )
        kind = normalize(workbook_kind)
        if kind == "course_details_template":
            generated_path = _course_template_generator()(
                output_path=output_path,
                cancel_token=cancel_token,
            )
            output_value = str(generated_path)
            return _WorkbookGenerationResult(
                status="generated",
                workbook_path=output_value,
                output_path=output_value,
                output_url=output_value,
                reason=None,
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
        from domain.template_strategy_router import assert_template_id_matches

        assert_template_id_matches(
            actual_template_id=template_id,
            expected_template_id=self.template_id,
        )
        if normalize(workbook_kind) not in {"course_details", "course_details_template"}:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_KIND_UNSUPPORTED",
                workbook_kind=workbook_kind,
                template_id=self.template_id,
            )
        return _course_template_validator()(
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
        from domain.template_strategy_router import assert_template_id_matches

        assert_template_id_matches(
            actual_template_id=template_id,
            expected_template_id=self.template_id,
        )
        if normalize(workbook_kind) not in {"course_details", "course_details_template"}:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_KIND_UNSUPPORTED",
                workbook_kind=workbook_kind,
                template_id=self.template_id,
            )
        return _course_template_batch_validator()(
            workbook_paths=workbook_paths,
            cancel_token=cancel_token,
        )

    def generate_workbooks(
        self,
        *,
        template_id: str,
        workbook_kind: str,
        workbook_paths: Sequence[str | Path],
        output_dir: str | Path,
        cancel_token: CancellationToken | None = None,
        context: Mapping[str, Any] | None = None,
    ) -> dict[str, object]:
        from domain.template_strategy_router import assert_template_id_matches

        assert_template_id_matches(
            actual_template_id=template_id,
            expected_template_id=self.template_id,
        )
        kind = normalize(workbook_kind)
        allow_overwrite = _overwrite_existing_enabled(context)
        output_path_overrides = _output_path_overrides_from_context(context)
        if kind == "marks_template":
            return _marks_template_batch_generator()(
                workbook_paths=workbook_paths,
                output_dir=Path(output_dir),
                allow_overwrite=allow_overwrite,
                output_path_overrides=output_path_overrides,
                cancel_token=cancel_token,
            )
        if kind == "course_details_template":
            unexpected_workbook_paths = [str(raw).strip() for raw in workbook_paths if str(raw).strip()]
            if unexpected_workbook_paths:
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="WORKBOOK_PATHS_NOT_APPLICABLE",
                    workbook_kind=workbook_kind,
                    expected="empty_sequence",
                )
            return _course_template_batch_generator()(
                workbook_paths=(),
                output_dir=Path(output_dir),
                allow_overwrite=allow_overwrite,
                cancel_token=cancel_token,
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
        )


@dataclass(slots=True, frozen=True)
class _WorkbookGenerationResult:
    status: str
    workbook_path: str | None
    output_path: str
    output_url: str
    reason: str | None = None


@lru_cache(maxsize=1)
def _marks_template_batch_generator() -> Callable[..., dict[str, object]]:
    try:
        from domain.template_versions.course_setup_v2_impl import marks_template as marks_template_impl
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 marks template implementation module.") from exc
    fn = getattr(marks_template_impl, "generate_marks_templates_from_course_details_batch", None)
    if not callable(fn):
        raise ConfigurationError(
            "V2 marks template implementation missing generate_marks_templates_from_course_details_batch()."
        )
    return cast(Callable[..., dict[str, object]], fn)


@lru_cache(maxsize=1)
def _course_template_batch_generator() -> Callable[..., dict[str, object]]:
    try:
        from domain.template_versions.course_setup_v2_impl import course_template as course_template_impl
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 course template implementation module.") from exc
    fn = getattr(course_template_impl, "generate_course_details_templates_batch", None)
    if not callable(fn):
        raise ConfigurationError(
            "V2 course template implementation missing generate_course_details_templates_batch()."
        )
    return cast(Callable[..., dict[str, object]], fn)


@lru_cache(maxsize=1)
def _course_template_generator() -> Callable[..., Path]:
    try:
        from domain.template_versions.course_setup_v2_impl.course_template import (
            generate_course_details_template,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 course template implementation module.") from exc
    fn = generate_course_details_template
    if not callable(fn):
        raise ConfigurationError("V2 course template implementation missing generate_course_details_template().")
    return fn


@lru_cache(maxsize=1)
def _course_template_validator() -> Callable[..., str]:
    try:
        from domain.template_versions.course_setup_v2_impl.course_template_validator import (
            validate_course_details_workbook,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 course template validator module.") from exc
    fn = validate_course_details_workbook
    if not callable(fn):
        raise ConfigurationError("V2 course template validator missing validate_course_details_workbook().")
    return fn


@lru_cache(maxsize=1)
def _course_template_batch_validator() -> Callable[..., dict[str, object]]:
    try:
        from domain.template_versions.course_setup_v2_impl.course_template_validator import (
            validate_course_details_workbooks,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 course template validator module.") from exc
    fn = validate_course_details_workbooks
    if not callable(fn):
        raise ConfigurationError("V2 course template validator missing validate_course_details_workbooks().")
    return fn


def _overwrite_existing_enabled(context: Mapping[str, Any] | None) -> bool:
    if not isinstance(context, Mapping):
        return False
    raw = context.get("overwrite_existing", False)
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, (int, float)):
        return bool(raw)
    if isinstance(raw, str):
        token = normalize(raw)
        return token in {"1", "true", "yes", "y", "on"}
    return False


def _output_path_overrides_from_context(context: Mapping[str, Any] | None) -> dict[str, str]:
    if not isinstance(context, Mapping):
        return {}
    raw = context.get("output_path_overrides")
    if not isinstance(raw, Mapping):
        return {}
    overrides: dict[str, str] = {}
    for key, value in raw.items():
        source = str(key).strip()
        output = str(value).strip()
        if not source or not output:
            continue
        overrides[source] = output
    return overrides


__all__ = ["CourseSetupV2Strategy"]
