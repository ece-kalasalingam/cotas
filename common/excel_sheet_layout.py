"""Shared Excel sheet layout/style helpers for openpyxl workflows."""

from __future__ import annotations

from typing import Any, Sequence

from common.constants import (
    ALLOW_FILTER,
    ALLOW_SELECT_LOCKED,
    ALLOW_SELECT_UNLOCKED,
    ALLOW_SORT,
    ID_COURSE_SETUP,
    ensure_workbook_secret_policy,
    get_workbook_password,
)
from common.registry import BLUEPRINT_REGISTRY

_STYLE_REGISTRY_HEADER = "header"
_STYLE_REGISTRY_BODY = "body"
_BORDER_THIN = "thin"
_BORDER_COLOR_BLACK = "000000"
_COLUMN_MIN_WIDTH = 8
_COLUMN_MAX_WIDTH = 60
_COLUMN_WIDTH_PADDING = 2
_HEADER_ACTIVE_COLUMN = "A"
_AUTOSIZE_SAMPLE_ROWS = 300


def style_registry_for_setup() -> tuple[dict[str, Any], dict[str, Any]]:
    blueprint = BLUEPRINT_REGISTRY.get(ID_COURSE_SETUP)
    if blueprint is None:
        return ({}, {})
    return (
        dict(blueprint.style_registry.get(_STYLE_REGISTRY_HEADER, {})),
        dict(blueprint.style_registry.get(_STYLE_REGISTRY_BODY, {})),
    )


def thin_border() -> Any:
    from openpyxl.styles import Border, Side

    thin = Side(style=_BORDER_THIN, color=_BORDER_COLOR_BLACK)
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def color_without_hash(color: str) -> str:
    return color[1:] if color.startswith("#") else color


def excel_col_name(col_index_1_based: int) -> str:
    index = col_index_1_based
    label = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        label = chr(65 + rem) + label
    return label


def autosize_columns(ws: Any, max_col: int) -> None:
    sampled_max_row = min(int(ws.max_row), _AUTOSIZE_SAMPLE_ROWS)
    if sampled_max_row <= 0:
        for col in range(1, max_col + 1):
            col_label = excel_col_name(col)
            ws.column_dimensions[col_label].width = _COLUMN_MIN_WIDTH
        return
    for col in range(1, max_col + 1):
        col_label = excel_col_name(col)
        max_len = 0
        for values in ws.iter_cols(
            min_col=col,
            max_col=col,
            min_row=1,
            max_row=sampled_max_row,
            values_only=True,
        ):
            for value in values:
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
            break
        ws.column_dimensions[col_label].width = min(
            _COLUMN_MAX_WIDTH,
            max(_COLUMN_MIN_WIDTH, max_len + _COLUMN_WIDTH_PADDING),
        )


def compute_sampled_column_widths(
    sample_rows: Sequence[Sequence[Any]],
    last_col: int,
    *,
    min_width: int = _COLUMN_MIN_WIDTH,
    max_width: int = _COLUMN_MAX_WIDTH,
    padding: int = _COLUMN_WIDTH_PADDING,
) -> dict[int, int]:
    widths: dict[int, int] = {}
    for col_index in range(last_col + 1):
        max_len = 0
        for row in sample_rows:
            if col_index >= len(row):
                continue
            value = row[col_index]
            if value is None:
                continue
            max_len = max(max_len, len(str(value).strip()))
        widths[col_index] = min(
            max_width,
            max(min_width, max_len + padding),
        )
    return widths


def set_header_selected_cell(ws: Any, header_row: int) -> None:
    active = f"{_HEADER_ACTIVE_COLUMN}{header_row}"
    ws.sheet_view.selection[0].activeCell = active
    ws.sheet_view.selection[0].sqref = active


def apply_sheet_layout_and_protection(
    *,
    ws: Any,
    header_row: int,
    header_count: int,
    paper_size: Any,
    orientation: Any,
) -> None:
    from openpyxl.worksheet.properties import PageSetupProperties

    ensure_workbook_secret_policy()
    ws.freeze_panes = ws.cell(row=header_row + 1, column=4)
    set_header_selected_cell(ws, header_row)
    autosize_columns(ws, header_count)
    ws.page_setup.paperSize = paper_size
    ws.page_setup.orientation = orientation
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.protection.sheet = True
    ws.protection.password = get_workbook_password()
    ws.protection.sort = ALLOW_SORT
    ws.protection.autoFilter = ALLOW_FILTER
    ws.protection.selectLockedCells = ALLOW_SELECT_LOCKED
    ws.protection.selectUnlockedCells = ALLOW_SELECT_UNLOCKED
