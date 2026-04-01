"""Static contract checks for workbook registry/schema."""

from __future__ import annotations

from collections.abc import Mapping

from common.attainment_policy import (
    has_valid_co_attainment_percent,
    thresholds_all_numeric,
    thresholds_non_decreasing,
    thresholds_within_range_0_100,
)
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


def require_keys(namespace: Mapping[str, object], *, keys: tuple[str, ...], context: str) -> None:
    """Require keys.
    
    Args:
        namespace: Parameter value (Mapping[str, object]).
        keys: Parameter value (tuple[str, ...]).
        context: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    missing = [key for key in keys if key not in namespace]
    if missing:
        ordered = ", ".join(sorted(missing))
        raise ConfigurationError(f"{context} namespace is missing required keys: {ordered}")


def validate_blueprint_registry_contracts() -> None:
    """Validate blueprint registry contracts.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
        seen_sheet_keys: set[str] = set()
        for sheet in blueprint.sheets:
            normalized_key = sheet.key.strip().lower()
            if not normalized_key:
                raise ConfigurationError(f"Blueprint '{type_id}' has an empty sheet key.")
            if normalized_key in seen_sheet_keys:
                raise ConfigurationError(
                    f"Blueprint '{type_id}' contains duplicate sheet key '{sheet.key}'."
                )
            seen_sheet_keys.add(normalized_key)

            normalized_name = sheet.name.strip().lower()
            if not normalized_name:
                raise ConfigurationError(f"Blueprint '{type_id}' has an empty sheet name.")
            if normalized_name in seen_sheet_names:
                raise ConfigurationError(
                    f"Blueprint '{type_id}' contains duplicate sheet name '{sheet.name}'."
                )
            seen_sheet_names.add(normalized_name)

            header_kind = str(sheet.header_kind).strip().lower()
            if header_kind not in {"fixed", "dynamic"}:
                raise ConfigurationError(
                    f"Sheet '{sheet.name}' in blueprint '{type_id}' has invalid header_kind '{sheet.header_kind}'."
                )

            if header_kind == "fixed":
                if not sheet.header_matrix or not sheet.header_matrix[0]:
                    raise ConfigurationError(
                        f"Sheet '{sheet.name}' in blueprint '{type_id}' must define fixed headers."
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
            else:
                if not str(sheet.header_resolver or "").strip():
                    raise ConfigurationError(
                        f"Sheet '{sheet.name}' in blueprint '{type_id}' declares dynamic headers but no header_resolver."
                    )

        if blueprint.workbook_rules and not isinstance(blueprint.workbook_rules, dict):
            raise ConfigurationError(
                f"Blueprint '{type_id}' has invalid workbook_rules; expected a dict."
            )


def _validate_attainment_policy_contracts() -> None:
    """Validate attainment policy contracts.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    if round(DIRECT_RATIO + INDIRECT_RATIO, 5) != 1.0:
        raise ConfigurationError("DIRECT_RATIO + INDIRECT_RATIO must equal 1.0")


def _validate_attainment_threshold_contracts() -> None:
    """Validate attainment threshold contracts.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    thresholds = (LEVEL_1_THRESHOLD, LEVEL_2_THRESHOLD, LEVEL_3_THRESHOLD)
    if not thresholds_all_numeric(*thresholds):
        raise ConfigurationError("Level thresholds must be numeric.")
    if not thresholds_within_range_0_100(
        float(LEVEL_1_THRESHOLD),
        float(LEVEL_2_THRESHOLD),
        float(LEVEL_3_THRESHOLD),
    ):
        raise ConfigurationError("Level thresholds must be in the range 0 to 100.")
    if not thresholds_non_decreasing(
        float(LEVEL_1_THRESHOLD),
        float(LEVEL_2_THRESHOLD),
        float(LEVEL_3_THRESHOLD),
    ):
        raise ConfigurationError("Level thresholds must be non-decreasing (L1 <= L2 <= L3).")
    if isinstance(CO_ATTAINMENT_PERCENT_DEFAULT, bool):
        raise ConfigurationError("CO_ATTAINMENT_PERCENT_DEFAULT must be numeric.")
    if not has_valid_co_attainment_percent(CO_ATTAINMENT_PERCENT_DEFAULT):
        if not isinstance(CO_ATTAINMENT_PERCENT_DEFAULT, (int, float)):
            raise ConfigurationError("CO_ATTAINMENT_PERCENT_DEFAULT must be numeric.")
        raise ConfigurationError("CO_ATTAINMENT_PERCENT_DEFAULT must be in the range 0 to 100.")
    if not isinstance(CO_ATTAINMENT_LEVEL_DEFAULT, int):
        raise ConfigurationError("CO_ATTAINMENT_LEVEL_DEFAULT must be an integer level index.")
    if not (1 <= CO_ATTAINMENT_LEVEL_DEFAULT <= len(thresholds)):
        raise ConfigurationError("CO_ATTAINMENT_LEVEL_DEFAULT must map to a configured attainment level.")


def _validate_indirect_tool_policy_contracts() -> None:
    """Validate indirect tool policy contracts.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    if LIKERT_MIN >= LIKERT_MAX:
        raise ConfigurationError("LIKERT_MIN must be less than LIKERT_MAX")
