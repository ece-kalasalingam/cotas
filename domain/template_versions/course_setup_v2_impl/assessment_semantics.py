"""COURSE_SETUP_V2 semantic reader for Assessment_Config component rows."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Literal, Sequence

from common.error_catalog import validation_error_from_key
from common.exceptions import ConfigurationError
from common.registry import (
    COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG,
    get_sheet_headers_by_key,
    get_sheet_name_by_key,
    get_sheet_schema_by_key,
)
from common.utils import coerce_excel_number, normalize
from domain.template_versions.course_setup_v2_impl.schema_columns import (
    required_column_index,
)

_TEMPLATE_ID = "COURSE_SETUP_V2"

_COL_COMPONENT = "component"
_COL_WEIGHT_PERCENT = "weight_percent"
_COL_CIA = "cia"
_COL_CO_WISE_BREAKUP = "co_wise_marks_breakup"
_COL_DIRECT = "direct"
_COL_ASSESSMENT_TYPE = "assessment_type"
_COL_ASSESSMENT_FORMAT = "assessment_format"
_COL_MODE = "mode"
_COL_PARTICIPATION = "participation"


@dataclass(slots=True, frozen=True)
class AssessmentComponent:
    row_number: int
    component_key: str
    component_name: str
    weight: float
    cia: bool
    co_wise_breakup: bool
    is_direct: bool
    assessment_type: str
    assessment_format: str
    mode: str
    participation: str


def _assessment_schema():
    """Assessment schema.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    schema = get_sheet_schema_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG)
    if schema is None:
        raise ConfigurationError(
            f"Template {_TEMPLATE_ID!r} does not define sheet key "
            f"{COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG!r}."
        )
    return schema


def _headers_from_registry() -> tuple[str, ...]:
    """Headers from registry.
    
    Args:
        None.
    
    Returns:
        tuple[str, ...]: Return value.
    
    Raises:
        None.
    """
    _assessment_schema()
    headers = get_sheet_headers_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG)
    if not headers:
        raise ConfigurationError(
            f"Assessment_Config headers are empty for template {_TEMPLATE_ID!r}."
        )
    return headers


def _parse_yes_no(value: Any, *, sheet_name: str, row_number: int, field_name: str) -> bool:
    """Parse yes no.
    
    Args:
        value: Parameter value (Any).
        sheet_name: Parameter value (str).
        row_number: Parameter value (int).
        field_name: Parameter value (str).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
    token = normalize(value)
    if token not in {"yes", "no"}:
        raise validation_error_from_key(
            "instructor.validation.yes_no_required",
            sheet_name=sheet_name,
            row=row_number,
            field=field_name,
        )
    return token == "yes"


def _parse_allowed_option(
    value: Any,
    *,
    sheet_name: str,
    row_number: int,
    field_name: str,
    allowed_tokens: set[str],
    allowed_display: Sequence[str],
) -> str:
    """Parse allowed option.
    
    Args:
        value: Parameter value (Any).
        sheet_name: Parameter value (str).
        row_number: Parameter value (int).
        field_name: Parameter value (str).
        allowed_tokens: Parameter value (set[str]).
        allowed_display: Parameter value (Sequence[str]).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    token = normalize(value)
    if token not in allowed_tokens:
        raise validation_error_from_key(
            "instructor.validation.allowed_values_required",
            sheet_name=sheet_name,
            row=row_number,
            field=field_name,
            allowed=", ".join(allowed_display),
        )
    return token


def parse_assessment_components(
    rows: Iterable[Sequence[Any]],
    *,
    sheet_name: str,
    row_start: int = 2,
    row_numbers: Sequence[int] | None = None,
    on_blank_component: Literal["error", "break", "skip"] = "error",
    duplicate_policy: Literal["error", "keep_first", "keep_all"] = "error",
    require_non_empty: bool = True,
    validate_allowed_options: bool = False,
    assessment_type_allowed_tokens: set[str] | None = None,
    assessment_type_allowed_display: Sequence[str] | None = None,
    assessment_format_allowed_tokens: set[str] | None = None,
    assessment_format_allowed_display: Sequence[str] | None = None,
    mode_allowed_tokens: set[str] | None = None,
    mode_allowed_display: Sequence[str] | None = None,
    participation_allowed_tokens: set[str] | None = None,
    participation_allowed_display: Sequence[str] | None = None,
) -> list[AssessmentComponent]:
    """Parse assessment components.
    
    Args:
        rows: Parameter value (Iterable[Sequence[Any]]).
        sheet_name: Parameter value (str).
        row_start: Parameter value (int).
        row_numbers: Parameter value (Sequence[int] | None).
        on_blank_component: Parameter value (Literal['error', 'break', 'skip']).
        duplicate_policy: Parameter value (Literal['error', 'keep_first', 'keep_all']).
        require_non_empty: Parameter value (bool).
        validate_allowed_options: Parameter value (bool).
        assessment_type_allowed_tokens: Parameter value (set[str] | None).
        assessment_type_allowed_display: Parameter value (Sequence[str] | None).
        assessment_format_allowed_tokens: Parameter value (set[str] | None).
        assessment_format_allowed_display: Parameter value (Sequence[str] | None).
        mode_allowed_tokens: Parameter value (set[str] | None).
        mode_allowed_display: Parameter value (Sequence[str] | None).
        participation_allowed_tokens: Parameter value (set[str] | None).
        participation_allowed_display: Parameter value (Sequence[str] | None).
    
    Returns:
        list[AssessmentComponent]: Return value.
    
    Raises:
        None.
    """
    expected_sheet_name = get_sheet_name_by_key(_TEMPLATE_ID, COURSE_SETUP_SHEET_KEY_ASSESSMENT_CONFIG)
    if normalize(sheet_name) != normalize(expected_sheet_name):
        raise ConfigurationError(
            f"Assessment parser expected sheet {expected_sheet_name!r} for template {_TEMPLATE_ID!r}, "
            f"but got {sheet_name!r}."
        )
    schema = _assessment_schema()
    headers = list(_headers_from_registry())
    idx = {
        _COL_COMPONENT: required_column_index(schema, _COL_COMPONENT),
        _COL_WEIGHT_PERCENT: required_column_index(schema, _COL_WEIGHT_PERCENT),
        _COL_CIA: required_column_index(schema, _COL_CIA),
        _COL_CO_WISE_BREAKUP: required_column_index(schema, _COL_CO_WISE_BREAKUP),
        _COL_DIRECT: required_column_index(schema, _COL_DIRECT),
        _COL_ASSESSMENT_TYPE: required_column_index(schema, _COL_ASSESSMENT_TYPE),
        _COL_ASSESSMENT_FORMAT: required_column_index(schema, _COL_ASSESSMENT_FORMAT),
        _COL_MODE: required_column_index(schema, _COL_MODE),
        _COL_PARTICIPATION: required_column_index(schema, _COL_PARTICIPATION),
    }
    components: list[AssessmentComponent] = []
    seen: set[str] = set()

    for offset, row in enumerate(rows):
        if row_numbers is not None and offset < len(row_numbers):
            row_number = int(row_numbers[offset])
        else:
            row_number = row_start + offset
        if len(row) <= idx[_COL_COMPONENT]:
            continue
        component_raw = row[idx[_COL_COMPONENT]]
        component_key = normalize(component_raw)
        if not component_key:
            if on_blank_component == "break":
                break
            if on_blank_component == "skip":
                continue
            raise validation_error_from_key(
                "instructor.validation.assessment_component_required",
                row=row_number,
            )

        if component_key in seen:
            if duplicate_policy == "error":
                raise validation_error_from_key(
                    "instructor.validation.assessment_component_duplicate",
                    row=row_number,
                    component=component_raw,
                )
            if duplicate_policy == "keep_first":
                continue

        weight_value = coerce_excel_number(row[idx[_COL_WEIGHT_PERCENT]])
        if isinstance(weight_value, bool) or not isinstance(weight_value, (int, float)):
            raise validation_error_from_key(
                "instructor.validation.assessment_weight_numeric",
                row=row_number,
            )

        cia = _parse_yes_no(
            row[idx[_COL_CIA]],
            sheet_name=sheet_name,
            row_number=row_number,
            field_name=headers[idx[_COL_CIA]],
        )
        co_wise_breakup = _parse_yes_no(
            row[idx[_COL_CO_WISE_BREAKUP]],
            sheet_name=sheet_name,
            row_number=row_number,
            field_name=headers[idx[_COL_CO_WISE_BREAKUP]],
        )
        is_direct = _parse_yes_no(
            row[idx[_COL_DIRECT]],
            sheet_name=sheet_name,
            row_number=row_number,
            field_name=headers[idx[_COL_DIRECT]],
        )

        if validate_allowed_options:
            if (
                assessment_type_allowed_tokens is None
                or assessment_type_allowed_display is None
                or assessment_format_allowed_tokens is None
                or assessment_format_allowed_display is None
                or mode_allowed_tokens is None
                or mode_allowed_display is None
                or participation_allowed_tokens is None
                or participation_allowed_display is None
            ):
                raise ConfigurationError(
                    "Allowed-option sets must be provided when validate_allowed_options=True."
                )
            assessment_type = _parse_allowed_option(
                row[idx[_COL_ASSESSMENT_TYPE]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_COL_ASSESSMENT_TYPE]],
                allowed_tokens=assessment_type_allowed_tokens,
                allowed_display=assessment_type_allowed_display,
            )
            assessment_format = _parse_allowed_option(
                row[idx[_COL_ASSESSMENT_FORMAT]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_COL_ASSESSMENT_FORMAT]],
                allowed_tokens=assessment_format_allowed_tokens,
                allowed_display=assessment_format_allowed_display,
            )
            mode = _parse_allowed_option(
                row[idx[_COL_MODE]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_COL_MODE]],
                allowed_tokens=mode_allowed_tokens,
                allowed_display=mode_allowed_display,
            )
            participation = _parse_allowed_option(
                row[idx[_COL_PARTICIPATION]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_COL_PARTICIPATION]],
                allowed_tokens=participation_allowed_tokens,
                allowed_display=participation_allowed_display,
            )
        else:
            assessment_type = normalize(row[idx[_COL_ASSESSMENT_TYPE]])
            assessment_format = normalize(row[idx[_COL_ASSESSMENT_FORMAT]])
            mode = normalize(row[idx[_COL_MODE]])
            participation = normalize(row[idx[_COL_PARTICIPATION]])

        seen.add(component_key)
        components.append(
            AssessmentComponent(
                row_number=row_number,
                component_key=component_key,
                component_name=str(component_raw).strip(),
                weight=float(weight_value),
                cia=cia,
                co_wise_breakup=co_wise_breakup,
                is_direct=is_direct,
                assessment_type=assessment_type,
                assessment_format=assessment_format,
                mode=mode,
                participation=participation,
            )
        )

    if require_non_empty and not components:
        raise validation_error_from_key("instructor.validation.assessment_component_required_one")
    return components


__all__ = ["AssessmentComponent", "parse_assessment_components"]
