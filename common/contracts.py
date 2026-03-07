"""Static contract checks for workbook registry/schema."""

from __future__ import annotations

from common.exceptions import ConfigurationError
from common.registry import BLUEPRINT_REGISTRY


def validate_blueprint_registry_contracts() -> None:
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
