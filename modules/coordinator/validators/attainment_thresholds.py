"""Coordinator attainment-threshold validation."""

from __future__ import annotations


def has_valid_attainment_thresholds(l1: float, l2: float, l3: float) -> bool:
    return 0.0 < l1 < l2 < l3 < 100.0

