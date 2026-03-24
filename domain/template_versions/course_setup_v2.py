"""Version-specific strategy for COURSE_SETUP_V2."""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, cast

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError, JobCancelledError
from common.jobs import CancellationToken
from common.utils import canonical_path_key, normalize

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
            return _course_template_generator()(
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
            return _marks_template_generator()(
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
        runner = _get_workbooks_batch_runner(kind)
        if runner is None:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_KIND_UNSUPPORTED",
                workbook_kind=workbook_kind,
                template_id=self.template_id,
            )
        return runner(
            workbook_paths=workbook_paths,
            output_dir=Path(output_dir),
            cancel_token=cancel_token,
            context=context,
        )


def _run_marks_template_batch(
    *,
    workbook_paths: Sequence[str | Path],
    output_dir: Path,
    cancel_token: CancellationToken | None,
    context: Mapping[str, Any] | None,
) -> dict[str, object]:
    del context
    generator = _marks_template_generator()
    seen_keys: set[str] = set()
    seen_output_names: set[str] = set()
    results: dict[str, object] = {}
    generated = 0
    failed = 0
    skipped = 0

    for raw_path in workbook_paths:
        if cancel_token is not None:
            cancel_token.raise_if_cancelled()

        source = Path(raw_path)
        key = canonical_path_key(source)

        if key in seen_keys:
            results[key] = {"status": "skipped", "output": None, "reason": "duplicate_source"}
            skipped += 1
            continue
        seen_keys.add(key)

        output_name = source.stem + "_marks_template.xlsx"
        if output_name in seen_output_names:
            results[key] = {"status": "failed", "output": None, "reason": "output_name_collision"}
            failed += 1
            continue
        seen_output_names.add(output_name)

        output_path = output_dir / output_name
        try:
            generator(
                course_details_path=str(source),
                output_path=output_path,
                cancel_token=cancel_token,
            )
            results[key] = {"status": "generated", "output": str(output_path), "reason": None}
            generated += 1
        except JobCancelledError:
            raise
        except Exception as exc:
            results[key] = {"status": "failed", "output": None, "reason": str(exc)}
            failed += 1

    return {
        "total": len(seen_keys) + skipped,
        "generated": generated,
        "failed": failed,
        "skipped": skipped,
        "results": results,
    }


def _get_workbooks_batch_runner(
    kind: str,
) -> Callable[..., dict[str, object]] | None:
    if kind == "marks_template":
        return _run_marks_template_batch
    # future: if kind == "co_template": return _run_co_template_batch
    return None


@lru_cache(maxsize=1)
def _marks_template_generator() -> Callable[..., Path]:
    try:
        from domain.template_versions.course_setup_v2_impl import marks_template as marks_template_impl
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 marks template implementation module.") from exc
    fn = getattr(marks_template_impl, "generate_marks_template_from_course_details", None)
    if not callable(fn):
        raise ConfigurationError(
            "V2 marks template implementation missing generate_marks_template_from_course_details()."
        )
    return cast(Callable[..., Path], fn)


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


__all__ = ["CourseSetupV2Strategy"]
