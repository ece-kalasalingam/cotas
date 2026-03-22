"""Shared semantic reader for Assessment_Config component rows."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Literal, Sequence

from common.exceptions import ConfigurationError
from common.error_catalog import validation_error_from_key
from common.utils import coerce_excel_number, normalize


_FIELD_COMPONENT = "Component"
_FIELD_WEIGHT = "Weight (%)"
_FIELD_CIA = "CIA"
_FIELD_CO_WISE_BREAKUP = "CO_Wise_Marks_Breakup"
_FIELD_DIRECT = "Direct"
_FIELD_ASSESSMENT_TYPE = "Assessment_Type"
_FIELD_ASSESSMENT_FORMAT = "Assessment_Format"
_FIELD_MODE = "Mode"
_FIELD_PARTICIPATION = "Participation"


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


def _header_indices(headers: Sequence[str]) -> dict[str, int]:
    expected = {
        _FIELD_COMPONENT,
        _FIELD_WEIGHT,
        _FIELD_CIA,
        _FIELD_CO_WISE_BREAKUP,
        _FIELD_DIRECT,
        _FIELD_ASSESSMENT_TYPE,
        _FIELD_ASSESSMENT_FORMAT,
        _FIELD_MODE,
        _FIELD_PARTICIPATION,
    }
    mapping = {normalize(name): idx for idx, name in enumerate(headers)}
    out: dict[str, int] = {}
    for field in expected:
        key = normalize(field)
        if key not in mapping:
            raise validation_error_from_key(
                "instructor.validation.header_mismatch",
                sheet_name="Assessment_Config",
                expected=field,
            )
        out[field] = mapping[key]
    return out


def _parse_yes_no(value: Any, *, sheet_name: str, row_number: int, field_name: str) -> bool:
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
    headers: Sequence[str],
    row_start: int = 2,
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
    idx = _header_indices(headers)
    components: list[AssessmentComponent] = []
    seen: set[str] = set()

    for offset, row in enumerate(rows):
        row_number = row_start + offset
        if len(row) <= idx[_FIELD_COMPONENT]:
            continue
        component_raw = row[idx[_FIELD_COMPONENT]]
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

        weight_value = coerce_excel_number(row[idx[_FIELD_WEIGHT]])
        if isinstance(weight_value, bool) or not isinstance(weight_value, (int, float)):
            raise validation_error_from_key(
                "instructor.validation.assessment_weight_numeric",
                row=row_number,
            )

        cia = _parse_yes_no(
            row[idx[_FIELD_CIA]],
            sheet_name=sheet_name,
            row_number=row_number,
            field_name=headers[idx[_FIELD_CIA]],
        )
        co_wise_breakup = _parse_yes_no(
            row[idx[_FIELD_CO_WISE_BREAKUP]],
            sheet_name=sheet_name,
            row_number=row_number,
            field_name=headers[idx[_FIELD_CO_WISE_BREAKUP]],
        )
        is_direct = _parse_yes_no(
            row[idx[_FIELD_DIRECT]],
            sheet_name=sheet_name,
            row_number=row_number,
            field_name=headers[idx[_FIELD_DIRECT]],
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
                row[idx[_FIELD_ASSESSMENT_TYPE]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_FIELD_ASSESSMENT_TYPE]],
                allowed_tokens=assessment_type_allowed_tokens,
                allowed_display=assessment_type_allowed_display,
            )
            assessment_format = _parse_allowed_option(
                row[idx[_FIELD_ASSESSMENT_FORMAT]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_FIELD_ASSESSMENT_FORMAT]],
                allowed_tokens=assessment_format_allowed_tokens,
                allowed_display=assessment_format_allowed_display,
            )
            mode = _parse_allowed_option(
                row[idx[_FIELD_MODE]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_FIELD_MODE]],
                allowed_tokens=mode_allowed_tokens,
                allowed_display=mode_allowed_display,
            )
            participation = _parse_allowed_option(
                row[idx[_FIELD_PARTICIPATION]],
                sheet_name=sheet_name,
                row_number=row_number,
                field_name=headers[idx[_FIELD_PARTICIPATION]],
                allowed_tokens=participation_allowed_tokens,
                allowed_display=participation_allowed_display,
            )
        else:
            assessment_type = normalize(row[idx[_FIELD_ASSESSMENT_TYPE]])
            assessment_format = normalize(row[idx[_FIELD_ASSESSMENT_FORMAT]])
            mode = normalize(row[idx[_FIELD_MODE]])
            participation = normalize(row[idx[_FIELD_PARTICIPATION]])

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
