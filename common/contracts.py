"""Static contract checks for workbook registry/schema."""

from __future__ import annotations

from common.constants import DIRECT_RATIO, INDIRECT_RATIO, LIKERT_MAX, LIKERT_MIN
from common.exceptions import ConfigurationError
from common.registry import BLUEPRINT_REGISTRY


def validate_blueprint_registry_contracts() -> None:
    _validate_attainment_policy_contracts()
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


def _validate_indirect_tool_policy_contracts() -> None:
    if LIKERT_MIN >= LIKERT_MAX:
        raise ConfigurationError("LIKERT_MIN must be less than LIKERT_MAX")
