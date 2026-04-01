"""Shared attainment-threshold policy helpers (SSOT for numeric rules)."""

from __future__ import annotations

from typing import TypeGuard


NumericValue = int | float


def _is_numeric(value: object) -> TypeGuard[NumericValue]:
    """Is numeric.
    
    Args:
        value: Parameter value (object).
    
    Returns:
        TypeGuard[NumericValue]: Return value.
    
    Raises:
        None.
    """
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def thresholds_all_numeric(l1: object, l2: object, l3: object) -> bool:
    """Thresholds all numeric.
    
    Args:
        l1: Parameter value (object).
        l2: Parameter value (object).
        l3: Parameter value (object).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
    return _is_numeric(l1) and _is_numeric(l2) and _is_numeric(l3)


def thresholds_within_range_0_100(l1: float, l2: float, l3: float) -> bool:
    """Thresholds within range 0 100.
    
    Args:
        l1: Parameter value (float).
        l2: Parameter value (float).
        l3: Parameter value (float).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
    return 0.0 <= l1 <= 100.0 and 0.0 <= l2 <= 100.0 and 0.0 <= l3 <= 100.0


def thresholds_non_decreasing(l1: float, l2: float, l3: float) -> bool:
    """Thresholds non decreasing.
    
    Args:
        l1: Parameter value (float).
        l2: Parameter value (float).
        l3: Parameter value (float).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
    return l1 <= l2 <= l3


def has_valid_attainment_thresholds(l1: object, l2: object, l3: object) -> bool:
    """Has valid attainment thresholds.
    
    Args:
        l1: Parameter value (object).
        l2: Parameter value (object).
        l3: Parameter value (object).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
    if not (_is_numeric(l1) and _is_numeric(l2) and _is_numeric(l3)):
        return False
    v1 = float(l1)
    v2 = float(l2)
    v3 = float(l3)
    return thresholds_within_range_0_100(v1, v2, v3) and thresholds_non_decreasing(v1, v2, v3)


def has_valid_co_attainment_percent(value: object) -> bool:
    """Has valid co attainment percent.
    
    Args:
        value: Parameter value (object).
    
    Returns:
        bool: Return value.
    
    Raises:
        None.
    """
    if not _is_numeric(value):
        return False
    v = float(value)
    return 0.0 <= v <= 100.0


__all__ = [
    "has_valid_attainment_thresholds",
    "has_valid_co_attainment_percent",
    "thresholds_all_numeric",
    "thresholds_non_decreasing",
    "thresholds_within_range_0_100",
]
