from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


StyleDefinition = Dict[str, Any]


@dataclass(frozen=True)
class ValidationRule:
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    options: Dict[str, Any]


@dataclass(frozen=True)
class SheetSchema:
    name: str
    header_matrix: List[List[Any]]
    validations: List[ValidationRule] = field(default_factory=list)
    header_style_key: str = "header"
    data_style_key: str = "body"
    is_protected: bool = False
    freeze_panes: Optional[tuple[int, int]] = None


@dataclass(frozen=True)
class WorkbookBlueprint:
    type_id: str
    style_registry: Dict[str, StyleDefinition]
    business_rules: List[Any]
    key_map: Dict[str, str]
    sheets: List[SheetSchema]
