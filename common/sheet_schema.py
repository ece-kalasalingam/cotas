from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class ValidationRule:
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    options: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class SheetSchema:
    name: str
    header_matrix: list[list[str]]
    key: str = ""
    header_kind: str = "fixed"
    header_resolver: str | None = None
    header_context: dict[str, Any] = field(default_factory=dict)
    validations: list[ValidationRule] = field(default_factory=list)
    is_protected: bool = True
    sheet_rules: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class WorkbookBlueprint:
    type_id: str
    style_registry: dict[str, dict[str, Any]]
    sheets: list[SheetSchema]
    workbook_rules: dict[str, Any] = field(default_factory=dict)
