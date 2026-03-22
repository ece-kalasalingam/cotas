"""Static contract checks for workbook registry/schema."""

from __future__ import annotations

from common.constants import (
    CO_ATTAINMENT_LEVEL_DEFAULT,
    CO_ATTAINMENT_PERCENT_DEFAULT,
    DIRECT_RATIO,
    INDIRECT_RATIO,
    LEVEL_1_THRESHOLD,
    LEVEL_2_THRESHOLD,
    LEVEL_3_THRESHOLD,
    LIKERT_MAX,
    LIKERT_MIN,
)
from common.exceptions import ConfigurationError
from common.registry import BLUEPRINT_REGISTRY


def validate_blueprint_registry_contracts() -> None:
    _validate_attainment_policy_contracts()
    _validate_attainment_threshold_contracts()
    _validate_indirect_tool_policy_contracts()

    if not BLUEPRINT_REGISTRY:
        raise ConfigurationError("BLUEPRINT_REGISTRY must not be empty.")

    for type_id, blueprint in BLUEPRINT_REGISTRY.items():
        if type_id != blueprint.type_id:
            raise ConfigurationError(
                f"Blueprint key '{type_id}' does not match type_id '{blueprint.type_id}'."
            )
        if not blueprint.sheets:
            raise ConfigurationError(f"Blueprint '{type_id}' must define at least one sheet.")

        seen_sheet_names: set[str] = set()
        for sheet in blueprint.sheets:
            normalized_name = sheet.name.strip().lower()
            if not normalized_name:
                raise ConfigurationError(f"Blueprint '{type_id}' has an empty sheet name.")
            if normalized_name in seen_sheet_names:
                raise ConfigurationError(
                    f"Blueprint '{type_id}' contains duplicate sheet name '{sheet.name}'."
                )
            seen_sheet_names.add(normalized_name)

            if not sheet.header_matrix or not sheet.header_matrix[0]:
                raise ConfigurationError(
                    f"Sheet '{sheet.name}' in blueprint '{type_id}' must define headers."
                )
            headers = [str(header).strip() for header in sheet.header_matrix[0]]
            if any(not header for header in headers):
                raise ConfigurationError(
                    f"Sheet '{sheet.name}' in blueprint '{type_id}' has empty header values."
                )
            if len(set(h.lower() for h in headers)) != len(headers):
                raise ConfigurationError(
                    f"Sheet '{sheet.name}' in blueprint '{type_id}' has duplicate headers."
                )


def _validate_attainment_policy_contracts() -> None:
    if round(DIRECT_RATIO + INDIRECT_RATIO, 5) != 1.0:
        raise ConfigurationError("DIRECT_RATIO + INDIRECT_RATIO must equal 1.0")


def _validate_attainment_threshold_contracts() -> None:
    thresholds = (LEVEL_1_THRESHOLD, LEVEL_2_THRESHOLD, LEVEL_3_THRESHOLD)
    if any((not isinstance(value, (int, float))) for value in thresholds):
        raise ConfigurationError("Level thresholds must be numeric.")
    if any((value < 0.0 or value > 100.0) for value in thresholds):
        raise ConfigurationError("Level thresholds must be in the range 0 to 100.")
    if not (LEVEL_1_THRESHOLD <= LEVEL_2_THRESHOLD <= LEVEL_3_THRESHOLD):
        raise ConfigurationError("Level thresholds must be non-decreasing (L1 <= L2 <= L3).")
    if not isinstance(CO_ATTAINMENT_PERCENT_DEFAULT, (int, float)):
        raise ConfigurationError("CO_ATTAINMENT_PERCENT_DEFAULT must be numeric.")
    if not (0.0 <= float(CO_ATTAINMENT_PERCENT_DEFAULT) <= 100.0):
        raise ConfigurationError("CO_ATTAINMENT_PERCENT_DEFAULT must be in the range 0 to 100.")
    if not isinstance(CO_ATTAINMENT_LEVEL_DEFAULT, int):
        raise ConfigurationError("CO_ATTAINMENT_LEVEL_DEFAULT must be an integer level index.")
    if not (1 <= CO_ATTAINMENT_LEVEL_DEFAULT <= len(thresholds)):
        raise ConfigurationError("CO_ATTAINMENT_LEVEL_DEFAULT must map to a configured attainment level.")


def _validate_indirect_tool_policy_contracts() -> None:
    if LIKERT_MIN >= LIKERT_MAX:
        raise ConfigurationError("LIKERT_MIN must be less than LIKERT_MAX")
