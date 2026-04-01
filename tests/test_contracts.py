from __future__ import annotations

import pytest

from common import contracts
from common.exceptions import ConfigurationError
from common.sheet_schema import SheetSchema, WorkbookBlueprint


def _bp(
    type_id: str = "type-a",
    *,
    sheet_key: str = "sheet_1",
    sheet_name: str = "Sheet1",
    headers: list[str] | None = None,
) -> WorkbookBlueprint:
    """Bp.
    
    Args:
        type_id: Parameter value (str).
        sheet_key: Parameter value (str).
        sheet_name: Parameter value (str).
        headers: Parameter value (list[str] | None).
    
    Returns:
        WorkbookBlueprint: Return value.
    
    Raises:
        None.
    """
    return WorkbookBlueprint(
        type_id=type_id,
        style_registry={},
        sheets=[SheetSchema(name=sheet_name, header_matrix=[headers or ["H1", "H2"]], key=sheet_key)],
    )


def test_validate_attainment_policy_contracts_requires_sum_to_one(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate attainment policy contracts requires sum to one.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "DIRECT_RATIO", 0.8)
    monkeypatch.setattr(contracts, "INDIRECT_RATIO", 0.3)
    with pytest.raises(ConfigurationError, match="must equal 1.0"):
        contracts._validate_attainment_policy_contracts()


def test_validate_indirect_tool_policy_contracts_requires_order(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate indirect tool policy contracts requires order.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "LIKERT_MIN", 5)
    monkeypatch.setattr(contracts, "LIKERT_MAX", 5)
    with pytest.raises(ConfigurationError, match="must be less"):
        contracts._validate_indirect_tool_policy_contracts()


def test_validate_blueprint_registry_contracts_rejects_empty_registry(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects empty registry.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "BLUEPRINT_REGISTRY", {})
    with pytest.raises(ConfigurationError, match="must not be empty"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_blueprint_registry_contracts_rejects_type_id_mismatch(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects type id mismatch.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "BLUEPRINT_REGISTRY", {"key-a": _bp(type_id="key-b")})
    with pytest.raises(ConfigurationError, match="does not match type_id"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_blueprint_registry_contracts_rejects_missing_sheets(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects missing sheets.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(
        contracts,
        "BLUEPRINT_REGISTRY",
        {"type-a": WorkbookBlueprint(type_id="type-a", style_registry={}, sheets=[])},
    )
    with pytest.raises(ConfigurationError, match="must define at least one sheet"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_blueprint_registry_contracts_rejects_empty_sheet_name(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects empty sheet name.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "BLUEPRINT_REGISTRY", {"type-a": _bp(sheet_name="   ")})
    with pytest.raises(ConfigurationError, match="empty sheet name"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_blueprint_registry_contracts_rejects_duplicate_sheet_names(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects duplicate sheet names.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    bp = WorkbookBlueprint(
        type_id="type-a",
        style_registry={},
        sheets=[
            SheetSchema(name="Students", header_matrix=[["A"]], key="students_1"),
            SheetSchema(name="students", header_matrix=[["B"]], key="students_2"),
        ],
    )
    monkeypatch.setattr(contracts, "BLUEPRINT_REGISTRY", {"type-a": bp})
    with pytest.raises(ConfigurationError, match="duplicate sheet name"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_blueprint_registry_contracts_rejects_missing_headers(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects missing headers.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    bp = WorkbookBlueprint(
        type_id="type-a",
        style_registry={},
        sheets=[SheetSchema(name="S", header_matrix=[], key="sheet_s")],
    )
    monkeypatch.setattr(contracts, "BLUEPRINT_REGISTRY", {"type-a": bp})
    with pytest.raises(ConfigurationError, match="must define fixed headers"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_blueprint_registry_contracts_rejects_empty_header_value(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects empty header value.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "BLUEPRINT_REGISTRY", {"type-a": _bp(headers=["A", "  "])})
    with pytest.raises(ConfigurationError, match="empty header values"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_blueprint_registry_contracts_rejects_duplicate_headers(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate blueprint registry contracts rejects duplicate headers.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "BLUEPRINT_REGISTRY", {"type-a": _bp(headers=["Name", "name"])})
    with pytest.raises(ConfigurationError, match="duplicate headers"):
        contracts.validate_blueprint_registry_contracts()


def test_validate_attainment_threshold_contracts_additional_guards(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test validate attainment threshold contracts additional guards.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    monkeypatch.setattr(contracts, "LEVEL_1_THRESHOLD", "x")
    monkeypatch.setattr(contracts, "LEVEL_2_THRESHOLD", 50)
    monkeypatch.setattr(contracts, "LEVEL_3_THRESHOLD", 75)
    with pytest.raises(ConfigurationError, match="must be numeric"):
        contracts._validate_attainment_threshold_contracts()

    monkeypatch.setattr(contracts, "LEVEL_1_THRESHOLD", -1)
    monkeypatch.setattr(contracts, "LEVEL_2_THRESHOLD", 50)
    monkeypatch.setattr(contracts, "LEVEL_3_THRESHOLD", 75)
    with pytest.raises(ConfigurationError, match="range 0 to 100"):
        contracts._validate_attainment_threshold_contracts()

    monkeypatch.setattr(contracts, "LEVEL_1_THRESHOLD", 60)
    monkeypatch.setattr(contracts, "LEVEL_2_THRESHOLD", 40)
    monkeypatch.setattr(contracts, "LEVEL_3_THRESHOLD", 80)
    with pytest.raises(ConfigurationError, match="non-decreasing"):
        contracts._validate_attainment_threshold_contracts()
