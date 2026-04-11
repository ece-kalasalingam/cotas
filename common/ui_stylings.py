"""Centralized application-level Qt stylesheets."""

from __future__ import annotations

import re

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

DROP_LIST_PLACEHOLDER_TEXT_MARGINS = (16, 16, -16, -16)
DROP_LIST_PLACEHOLDER_BOTTOM_MARGINS = (16, 16, -16, -8)
DROP_LIST_ITEM_SPACING = 2
FILE_ITEM_LAYOUT_MARGINS = (12, 4, 12, 4)
FILE_ITEM_LAYOUT_SPACING = 12
DROP_ZONE_LAYOUT_MARGINS = (14, 14, 14, 14)
DROP_ZONE_LAYOUT_SPACING = 0

_MANAGED_BLOCK_TEMPLATE = "/* COTAS:{id}:BEGIN */\n{body}\n/* COTAS:{id}:END */"

INSTRUCTOR_PANEL_STYLESHEET = """
QFrame#coordinatorLeftCard,
QFrame#stepRail {
    border: 1px solid palette(mid);
    border-radius: 12px;
    background-color: palette(base);
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
QLabel#coordinatorTitle,
QLabel#coordinatorThresholdTitle,
QLabel#instructorRailTitle {
    font-weight: 700;
}
QLabel#coordinatorThresholdL1Label {
    font-weight: 700;
    color: palette(text);
}
QScrollArea#coordinatorLeftScroll,
QScrollArea#coordinatorRightScroll,
QScrollArea#instructorLeftScroll,
QScrollArea#instructorRightScroll {
    border: none;
    background: transparent;
}
QWidget#coordinatorLeftScrollViewport,
QWidget#coordinatorRightScrollViewport,
QWidget#instructorLeftScrollViewport,
QWidget#instructorRightScrollViewport {
    border: none;
    background: transparent;
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

"""

MAIN_ACTIVITYBAR_STYLESHEET = """
QToolBar#mainActivityBar {
    spacing: 0px;
    padding: 5px;
}
QToolBar#mainActivityBar QToolButton {
    min-width: 80px;
    max-width: 80px;
    padding: 5px;
    border: none;
    border-radius: 4px;
}
QToolBar#mainActivityBar QToolButton:hover {
    background-color: rgba(255, 255, 255, 0.1);
}
QToolBar#mainActivityBar QToolButton:checked {
    background-color: rgba(22, 160, 133, 0.2);
    border-bottom: 2px solid #16A085;
}
""".strip()

SHARED_INFO_PANE_STYLESHEET = """
#sharedActivityLog, #sharedGeneratedOutputs {
    border: none;
    outline: none;
}
#sharedActivityLog:focus, #sharedGeneratedOutputs:focus {
    border: none;
    outline: none;
}
#sharedActivityLog:hover, #sharedGeneratedOutputs:hover {
    border: none;
}
""".strip()

CLEAR_ALL_LINK_STYLESHEET = """
QPushButton#clearAllLink {
    background: transparent;
    border: none;
    padding: 0px;
    margin: 0px;
    min-width: 0px;
    min-height: 0px;
}
QPushButton#clearAllLink:enabled {
    text-decoration: underline;
}
""".strip()

FILE_ACTION_BUTTONS_STYLESHEET = """
QPushButton#coordinatorFileOpenButton,
QPushButton#coordinatorFolderOpenButton,
QPushButton#coordinatorFileRemoveButton {
    background-color: transparent;
    border: none;
    padding: 0px;
    margin: 0px;
    min-width: 24px;
    min-height: 24px;
    max-width: 24px;
    max-height: 24px;
}
QPushButton#coordinatorFileOpenButton:hover,
QPushButton#coordinatorFolderOpenButton:hover {
    background-color: palette(mid);
    border-radius: 4px;
}
QPushButton#coordinatorFileRemoveButton:hover {
    background-color: rgba(231, 76, 60, 0.15);
    border-radius: 4px;
}
""".strip()

ABOUT_HEADER_TEXT_STYLESHEET = """
QLabel#aboutLeftAppName {
    font-size: 24px;
    font-weight: 700;
}
QLabel#aboutLeftSubtitle {
    font-size: 15px;
}
QLabel#aboutLeftVersion {
    font-size: 9px;
}
""".strip()


def _upsert_managed_block(stylesheet: str, block_id: str, body: str) -> str:
    """Upsert managed block.
    
    Args:
        stylesheet: Parameter value (str).
        block_id: Parameter value (str).
        body: Parameter value (str).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    pattern = re.compile(
        rf"/\* COTAS:{re.escape(block_id)}:BEGIN \*/.*?/\* COTAS:{re.escape(block_id)}:END \*/",
        re.DOTALL,
    )
    block = _MANAGED_BLOCK_TEMPLATE.format(id=block_id, body=body.strip())
    if pattern.search(stylesheet):
        return pattern.sub(block, stylesheet)
    return f"{stylesheet}\n\n{block}".strip() if stylesheet else block


def apply_global_ui_styles(app: QApplication) -> None:
    """Apply managed global stylesheet blocks at application scope.

    Hybrid styling model:
    - qdarktheme owns base palette/theme state.
    - this function overlays stable app-specific UI/UX rules by managed block ids.
    - repeated calls are safe; existing blocks are replaced in-place.
    """
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
        "instructor-panel",
        INSTRUCTOR_PANEL_STYLESHEET,
    )
    merged_stylesheet = _upsert_managed_block(
        merged_stylesheet,
        "main-activitybar",
        MAIN_ACTIVITYBAR_STYLESHEET,
    )
    merged_stylesheet = _upsert_managed_block(
        merged_stylesheet,
        "shared-info-pane",
        SHARED_INFO_PANE_STYLESHEET,
    )
    merged_stylesheet = _upsert_managed_block(
        merged_stylesheet,
        "clear-all-link",
        CLEAR_ALL_LINK_STYLESHEET,
    )
    merged_stylesheet = _upsert_managed_block(
        merged_stylesheet,
        "coordinator-file-action-buttons",
        FILE_ACTION_BUTTONS_STYLESHEET,
    )
    merged_stylesheet = _upsert_managed_block(
        merged_stylesheet,
        "about-header-text",
        ABOUT_HEADER_TEXT_STYLESHEET,
    )
    if merged_stylesheet == current_stylesheet:
        return
    set_stylesheet(merged_stylesheet)

