"""Centralized application-level Qt stylesheets."""

from __future__ import annotations

import re

from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import QApplication

GLOBAL_QPUSHBUTTON_MIN_WIDTH = 150

_QPUSHBUTTON_GLOBAL_STYLESHEET = """
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

_MANAGED_BLOCK_TEMPLATE = "/* COTAS:{id}:BEGIN */\n{body}\n/* COTAS:{id}:END */"

INSTRUCTOR_PANEL_STYLESHEET = """
QFrame#stepRail {
    border: 1px solid palette(mid);
    border-radius: 12px;
    background-color: palette(base);
}
QFrame#activeCard {
}
QListWidget#stepList {
    outline: none;
    background-color: transparent;
}
QListWidget#stepList::item {
    padding: 8px 8px;
}
QListWidget#stepList::item:selected,
QListWidget#stepList::item:selected:!active {
    border-left: 4px solid palette(highlight);
}
QPushButton#primaryAction {
    padding: 6px 12px;
    min-width: 150px;
    min-height: 30px;
    border-radius: 6px;
}
QPushButton#primaryAction:enabled {
    background-color: palette(highlight);
    color: palette(highlighted-text);
    border: none;
    font-weight: 600;
}
QPushButton {
    padding: 6px 12px;
    min-width: 150px;
    min-height: 30px;
    border-radius: 6px;
    border: none;
}
QPushButton:disabled {
    border: 1px solid palette(mid);
}
QTabWidget#instructorInfoTabs::pane {
    border: none;
    background: palette(base);
}
QTabWidget#instructorInfoTabs QTabBar::tab:first {
    margin-left: 8px;
}
QTabWidget#instructorInfoTabs QPlainTextEdit,
QTabWidget#instructorInfoTabs QTextBrowser {
    border: 1px solid palette(mid);
    border-radius: 8px;
    background: palette(base);
    padding: 8px;
}
"""


def _upsert_managed_block(stylesheet: str, block_id: str, body: str) -> str:
    pattern = re.compile(
        rf"/\* COTAS:{re.escape(block_id)}:BEGIN \*/.*?/\* COTAS:{re.escape(block_id)}:END \*/",
        re.DOTALL,
    )
    block = _MANAGED_BLOCK_TEMPLATE.format(id=block_id, body=body.strip())
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


def _build_theme_adaptive_surface_styles() -> str:
    return f"""


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
        _QPUSHBUTTON_GLOBAL_STYLESHEET,
    )
    merged_stylesheet = _upsert_managed_block(
        merged_stylesheet,
        "theme-adaptive-surfaces",
        _build_theme_adaptive_surface_styles(),
    )
    if merged_stylesheet == current_stylesheet:
        return
    set_stylesheet(merged_stylesheet)
