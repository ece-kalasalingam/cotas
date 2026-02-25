from dataclasses import dataclass, field
from typing import List, Mapping, Any, Optional, Literal, TypedDict
from core.exceptions import SystemError

class StyleDefinition(TypedDict, total=False):
    bold: bool
    bg_color: str
    font_color: str
    border: int
    align: str
    valign: str
    locked: bool
    num_format: str
    text_wrap: bool

class ValidationOptions(TypedDict, total=False):
    validate: Literal['list', 'whole', 'decimal', 'date', 'time', 'text_length', 'custom', 'any']
    criteria: str
    value: Any
    source: Any
    input_title: str
    input_message: str
    error_title: str
    error_message: str
    ignore_blank: bool
    show_input: bool
    show_error: bool
    dropdown: bool

@dataclass(frozen=True)
class ValidationRule:
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    options: ValidationOptions

@dataclass(frozen=True)
class SheetSchema:
    name: str
    header_matrix: List[List[str]]
    header_style_key: str = 'header'
    data_style_key: str = 'body'
    is_protected: bool = False
    freeze_panes: Optional[tuple] = None
    validations: List[ValidationRule] = field(default_factory=list)

@dataclass(frozen=True)
class WorkbookBlueprint:
    type_id: str
    style_registry: Mapping[str, StyleDefinition]
    sheets: List[SheetSchema]

    def validate_structure(self):
        if not self.sheets:
            raise SystemError(f"Blueprint '{self.type_id}' contains no sheets.")