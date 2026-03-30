"""Version-specific strategy for COURSE_SETUP_V2."""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, Protocol, cast

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError
from common.jobs import CancellationToken
from common.utils import normalize

_SUPPORTED_OPERATIONS = frozenset(
    {
        "generate_workbook",
        "generate_workbooks",
        "validate_workbooks",
        "consume_last_marks_anomaly_warnings",
        "generate_co_attainment",
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
        if kind == "co_description_template":
            generated_path = _co_description_template_generator()(
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
        if kind == "co_attainment":
            payload = dict(context or {})
            source_paths_raw = payload.get("source_paths")
            source_paths = [Path(path) for path in source_paths_raw] if isinstance(source_paths_raw, list) else []
            if not source_paths:
                raise validation_error_from_key(
                    "common.validation_failed_invalid_data",
                    code="COA_SOURCE_WORKBOOK_REQUIRED",
                )
            result = self.generate_co_attainment(
                source_paths,
                Path(output_path),
                token=cancel_token or CancellationToken(),
                thresholds=payload.get("thresholds")
                if isinstance(payload.get("thresholds"), tuple)
                else None,
                co_attainment_percent=float(payload["co_attainment_percent"])
                if payload.get("co_attainment_percent") is not None
                else None,
                co_attainment_level=int(payload["co_attainment_level"])
                if payload.get("co_attainment_level") is not None
                else None,
            )
            output_value = str(getattr(result, "output_path", result)).strip() or str(output_path)
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
        kind = normalize(workbook_kind)
        if kind == "course_details":
            return _course_template_batch_validator()(
                workbook_paths=workbook_paths,
                cancel_token=cancel_token,
            )
        if kind == "marks_template":
            return _marks_template_batch_validator()(
                workbook_paths=workbook_paths,
                template_id=self.template_id,
                cancel_token=cancel_token,
            )
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
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
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="WORKBOOK_KIND_UNSUPPORTED",
            workbook_kind=workbook_kind,
            template_id=self.template_id,
        )

    def consume_last_marks_anomaly_warnings(self) -> list[str]:
        return list(_marks_template_warning_consumer()())

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
        signature = _final_report_signature_reader()(source_paths[0])
        if signature is None:
            raise validation_error_from_key(
                "validation.workbook.open_failed",
                code="WORKBOOK_OPEN_FAILED",
                workbook=str(source_paths[0]),
            )
        return _co_attainment_generator()(
            source_paths=source_paths,
            output_path=output_path,
            token=token,
            total_outcomes=signature.total_outcomes,
            template_id=signature.template_id,
            thresholds=thresholds,
            co_attainment_percent=co_attainment_percent,
            co_attainment_level=co_attainment_level,
        )


@dataclass(slots=True, frozen=True)
class _WorkbookGenerationResult:
    status: str
    workbook_path: str | None
    output_path: str
    output_url: str
    reason: str | None = None


class _FinalReportSignatureLike(Protocol):
    template_id: str
    total_outcomes: int


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


@lru_cache(maxsize=1)
def _co_description_template_generator() -> Callable[..., Path]:
    try:
        from domain.template_versions.course_setup_v2_impl.co_description_template import (
            generate_co_description_template,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 CO description template implementation module.") from exc
    fn = generate_co_description_template
    if not callable(fn):
        raise ConfigurationError(
            "V2 CO description implementation missing generate_co_description_template()."
        )
    return fn


@lru_cache(maxsize=1)
def _marks_template_batch_validator() -> Callable[..., dict[str, object]]:
    try:
        from domain.template_versions.course_setup_v2_impl.marks_template_validator import (
            validate_filled_marks_workbooks,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 marks template validator module.") from exc
    fn = validate_filled_marks_workbooks
    if not callable(fn):
        raise ConfigurationError("V2 marks template validator missing validate_filled_marks_workbooks().")
    return cast(Callable[..., dict[str, object]], fn)


@lru_cache(maxsize=1)
def _marks_template_warning_consumer() -> Callable[[], list[str]]:
    try:
        from domain.template_versions.course_setup_v2_impl.marks_template_validator import (
            consume_last_marks_anomaly_warnings,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 marks template warning consumer.") from exc
    fn = consume_last_marks_anomaly_warnings
    if not callable(fn):
        raise ConfigurationError(
            "V2 marks template validator missing consume_last_marks_anomaly_warnings()."
        )
    return fn


@lru_cache(maxsize=1)
def _co_attainment_generator() -> Callable[..., object]:
    try:
        from domain.template_versions.course_setup_v1_coordinator_engine import (
            _generate_co_attainment_workbook_course_setup_v1,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import CO attainment generator.") from exc
    fn = _generate_co_attainment_workbook_course_setup_v1
    if not callable(fn):
        raise ConfigurationError(
            "CO attainment generator missing _generate_co_attainment_workbook_course_setup_v1()."
        )
    return fn


@lru_cache(maxsize=1)
def _final_report_signature_reader() -> Callable[[Path], _FinalReportSignatureLike | None]:
    try:
        from domain.template_versions.course_setup_v1_coordinator_engine import (
            extract_final_report_signature_from_path,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import final report signature reader.") from exc
    fn = extract_final_report_signature_from_path
    if not callable(fn):
        raise ConfigurationError("Missing extract_final_report_signature_from_path().")
    return cast(Callable[[Path], _FinalReportSignatureLike | None], fn)


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
