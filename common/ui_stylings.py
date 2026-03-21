"""Centralized application-level Qt stylesheets."""

from __future__ import annotations

import re

from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import QApplication

GLOBAL_QPUSHBUTTON_MIN_WIDTH = 150

QPUSHBUTTON_GLOBAL_STYLESHEET = """
QPushButton {
    padding: 6px 12px;
    min-width: %dpx;
    min-height: 30px;
    border-radius: 6px;
    border: none;
}
""".strip() % GLOBAL_QPUSHBUTTON_MIN_WIDTH

COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS = (16, 16, -16, -16)
COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS = (16, 16, -16, -8)
COORDINATOR_DROP_LIST_ITEM_SPACING = 2
COORDINATOR_FILE_ITEM_LAYOUT_MARGINS = (12, 4, 12, 4)
COORDINATOR_FILE_ITEM_LAYOUT_SPACING = 12
COORDINATOR_DROP_ZONE_LAYOUT_MARGINS = (14, 14, 14, 14)
COORDINATOR_DROP_ZONE_LAYOUT_SPACING = 0
COORDINATOR_REMOVE_BUTTON_STYLESHEET = """
QPushButton {
    background-color: transparent;
    border: none;
    padding: 0px;
    margin: 0px;
    min-width: 24px;
    min-height: 24px;
    max-width: 24px;
    max-height: 24px;
}
QPushButton:hover {
    background-color: rgba(231, 76, 60, 0.15);
    border-radius: 4px;
}
""".strip()

_MANAGED_BLOCK_RE = r"/\* COTAS:{id}:BEGIN \*/.*?/\* COTAS:{id}:END \*/"


def _managed_block(block_id: str, body: str) -> str:
    return f"/* COTAS:{block_id}:BEGIN */\n{body}\n/* COTAS:{block_id}:END */"


def _upsert_managed_block(stylesheet: str, block_id: str, body: str) -> str:
    pattern = re.compile(_MANAGED_BLOCK_RE.format(id=re.escape(block_id)), re.DOTALL)
    block = _managed_block(block_id, body.strip())
    if pattern.search(stylesheet):
        return pattern.sub(block, stylesheet)
    return f"{stylesheet}\n\n{block}".strip() if stylesheet else block


def _is_dark_theme(app: QApplication) -> bool:
    color_scheme = app.styleHints().colorScheme()
    if color_scheme == Qt.ColorScheme.Dark:
        return True
    if color_scheme == Qt.ColorScheme.Light:
        return False
    window_color = app.palette().color(QPalette.ColorRole.Window)
    return window_color.lightness() < 128


def _build_theme_adaptive_surface_styles(*, dark: bool) -> str:
    pane_bg = "#1f242c" if dark else "#FFFFFF"
    pane_border = "#d0d7e2" if dark else "#3b4352"
    selected_bg = "#374151" if dark else "#dbe4f3"
    selected_fg = "#ffffff" if dark else "#111827"
    return f"""
QFrame#coordinatorLeftCard,
QFrame#stepRail {{
    background-color: {pane_bg};
    border: 1px solid {pane_border};
    border-radius: 10px;
}}
QWidget#coordinatorLeftScrollViewport,
QWidget#instructorLeftScrollViewport,
QListWidget#stepList {{
    background-color: {pane_bg};
}}
QListWidget#stepList::item:selected {{
    background-color: {selected_bg};
    color: {selected_fg};
}}
""".strip()


def apply_global_ui_styles(app: QApplication) -> None:
    """Apply shared stylesheet rules at application scope."""
    get_stylesheet = getattr(app, "styleSheet", None)
    set_stylesheet = getattr(app, "setStyleSheet", None)
    if not callable(set_stylesheet):
        return
    current_stylesheet = ""
    if callable(get_stylesheet):
        current_value = get_stylesheet()
        current_stylesheet = current_value.strip() if isinstance(current_value, str) else ""
    merged_stylesheet = _upsert_managed_block(
        current_stylesheet,
        "pushbutton-global",
        QPUSHBUTTON_GLOBAL_STYLESHEET,
    )
    merged_stylesheet = _upsert_managed_block(
        merged_stylesheet,
        "theme-adaptive-surfaces",
        _build_theme_adaptive_surface_styles(dark=_is_dark_theme(app)),
    )
    if merged_stylesheet == current_stylesheet:
        return
    set_stylesheet(merged_stylesheet)
