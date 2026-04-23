"""Cached V2 strategy bindings and shared context decoding helpers."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from functools import lru_cache
from pathlib import Path
from typing import Any, Protocol, cast

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError
from common.utils import normalize


class FinalReportSignatureLike(Protocol):
    template_id: str
    total_outcomes: int


@lru_cache(maxsize=1)
def marks_template_batch_generator() -> Callable[..., dict[str, object]]:
    """Marks template batch generator.
    
    Args:
        None.
    
    Returns:
        Callable[..., dict[str, object]]: Return value.
    
    Raises:
        None.
    """
    try:
        from domain.template_versions.course_setup_v2_impl import (
            marks_template as marks_template_impl,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import V2 marks template implementation module.") from exc
    fn = getattr(marks_template_impl, "generate_marks_templates_from_course_details_batch", None)
    if not callable(fn):
        raise ConfigurationError(
            "V2 marks template implementation missing generate_marks_templates_from_course_details_batch()."
        )
    return cast(Callable[..., dict[str, object]], fn)


@lru_cache(maxsize=1)
def course_template_generator() -> Callable[..., Path]:
    """Course template generator.
    
    Args:
        None.
    
    Returns:
        Callable[..., Path]: Return value.
    
    Raises:
        None.
    """
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
def course_template_batch_validator() -> Callable[..., dict[str, object]]:
    """Course template batch validator.
    
    Args:
        None.
    
    Returns:
        Callable[..., dict[str, object]]: Return value.
    
    Raises:
        None.
    """
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
def co_description_template_generator() -> Callable[..., Path]:
    """Co description template generator.
    
    Args:
        None.
    
    Returns:
        Callable[..., Path]: Return value.
    
    Raises:
        None.
    """
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
def marks_template_batch_validator() -> Callable[..., dict[str, object]]:
    """Marks template batch validator.
    
    Args:
        None.
    
    Returns:
        Callable[..., dict[str, object]]: Return value.
    
    Raises:
        None.
    """
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
def marks_template_warning_consumer() -> Callable[[], list[str]]:
    """Marks template warning consumer.
    
    Args:
        None.
    
    Returns:
        Callable[[], list[str]]: Return value.
    
    Raises:
        None.
    """
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


def consume_last_marks_anomaly_warnings() -> list[str]:
    """Consume last marks anomaly warnings.
    
    Args:
        None.
    
    Returns:
        list[str]: Return value.
    
    Raises:
        None.
    """
    return list(marks_template_warning_consumer()())


@lru_cache(maxsize=1)
def co_attainment_generator() -> Callable[..., object]:
    """Co attainment generator.
    
    Args:
        None.
    
    Returns:
        Callable[..., object]: Return value.
    
    Raises:
        None.
    """
    try:
        from domain.template_versions.course_setup_v2_impl.co_attainment import (
            generate_co_attainment_workbook,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import CO attainment generator.") from exc
    fn = generate_co_attainment_workbook
    if not callable(fn):
        raise ConfigurationError("CO attainment generator missing generate_co_attainment_workbook().")
    return fn


@lru_cache(maxsize=1)
def final_report_signature_reader() -> Callable[[Path], FinalReportSignatureLike | None]:
    """Final report signature reader.
    
    Args:
        None.
    
    Returns:
        Callable[[Path], FinalReportSignatureLike | None]: Return value.
    
    Raises:
        None.
    """
    try:
        from domain.template_versions.course_setup_v2_impl.co_attainment import (
            extract_final_report_signature_from_path,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import final report signature reader.") from exc
    fn = extract_final_report_signature_from_path
    if not callable(fn):
        raise ConfigurationError("Missing extract_final_report_signature_from_path().")
    return cast(Callable[[Path], FinalReportSignatureLike | None], fn)


@lru_cache(maxsize=1)
def total_outcomes_reader() -> Callable[[Path], int | None]:
    """Total outcomes reader.
    
    Args:
        None.
    
    Returns:
        Callable[[Path], int | None]: Return value.
    
    Raises:
        None.
    """
    try:
        from domain.template_versions.course_setup_v2_impl.co_attainment import (
            extract_total_outcomes_from_workbook_path,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import total outcomes reader.") from exc
    fn = extract_total_outcomes_from_workbook_path
    if not callable(fn):
        raise ConfigurationError("Missing extract_total_outcomes_from_workbook_path().")
    return cast(Callable[[Path], int | None], fn)


@lru_cache(maxsize=1)
def course_metadata_students_extractor() -> Callable[[Path], tuple[set[str], dict[str, str]]]:
    """Course metadata students extractor.
    
    Args:
        None.
    
    Returns:
        Callable[[Path], tuple[set[str], dict[str, str]]]: Return value.
    
    Raises:
        None.
    """
    try:
        from domain.template_versions.course_setup_v2_impl.co_attainment import (
            extract_course_metadata_and_students_from_workbook_path,
        )
    except Exception as exc:
        raise ConfigurationError("Unable to import course metadata/students extractor.") from exc
    fn = extract_course_metadata_and_students_from_workbook_path
    if not callable(fn):
        raise ConfigurationError("Missing extract_course_metadata_and_students_from_workbook_path().")
    return cast(Callable[[Path], tuple[set[str], dict[str, str]]], fn)


def co_attainment_generation_inputs(
    *,
    context: Mapping[str, Any] | None,
    output_path: str | Path,
    default_template_id: str,
) -> dict[str, object]:
    """Co attainment generation inputs.
    
    Args:
        context: Parameter value (Mapping[str, Any] | None).
        output_path: Parameter value (str | Path).
        default_template_id: Parameter value (str).
    
    Returns:
        dict[str, object]: Return value.
    
    Raises:
        None.
    """
    payload = dict(context or {})
    source_paths_raw = payload.get("source_paths")
    source_paths = [Path(path) for path in source_paths_raw] if isinstance(source_paths_raw, list) else []
    if not source_paths:
        raise validation_error_from_key(
            "common.validation_failed_invalid_data",
            code="COA_SOURCE_WORKBOOK_REQUIRED",
        )
    thresholds = payload.get("thresholds") if isinstance(payload.get("thresholds"), tuple) else None
    co_attainment_percent = (
        float(payload["co_attainment_percent"])
        if payload.get("co_attainment_percent") is not None
        else None
    )
    co_attainment_level = (
        int(payload["co_attainment_level"])
        if payload.get("co_attainment_level") is not None
        else None
    )
    signature = final_report_signature_reader()(source_paths[0])
    resolved_template_id = default_template_id
    resolved_total_outcomes: int | None = None
    if signature is None:
        resolved_total_outcomes = total_outcomes_reader()(source_paths[0])
        if resolved_total_outcomes is None:
            raise validation_error_from_key(
                "validation.workbook.open_failed",
                code="WORKBOOK_OPEN_FAILED",
                workbook=str(source_paths[0]),
            )
    else:
        resolved_template_id = signature.template_id
        resolved_total_outcomes = int(signature.total_outcomes)
    return {
        "source_paths": source_paths,
        "output_path": Path(output_path),
        "template_id": resolved_template_id,
        "total_outcomes": resolved_total_outcomes,
        "thresholds": thresholds,
        "co_attainment_percent": co_attainment_percent,
        "co_attainment_level": co_attainment_level,
    }


def overwrite_existing_enabled(context: Mapping[str, Any] | None) -> bool:
    """Overwrite existing enabled.
    
    Args:
        context: Parameter value (Mapping[str, Any] | None).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
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


def output_path_overrides_from_context(
    context: Mapping[str, Any] | None,
) -> dict[str, str]:
    """Output path overrides from context.
    
    Args:
        context: Parameter value (Mapping[str, Any] | None).
    
    Returns:
        dict[str, str]: Return value.
    
    Raises:
        None.
    """
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


__all__ = [
    "co_attainment_generation_inputs",
    "co_attainment_generator",
    "co_description_template_generator",
    "consume_last_marks_anomaly_warnings",
    "course_metadata_students_extractor",
    "course_template_batch_validator",
    "course_template_generator",
    "marks_template_batch_generator",
    "marks_template_batch_validator",
    "output_path_overrides_from_context",
    "overwrite_existing_enabled",
]
