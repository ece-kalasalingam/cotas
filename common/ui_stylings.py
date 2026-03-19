"""Centralized application-level Qt stylesheets."""

from __future__ import annotations

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

COORDINATOR_LIST_PLACEHOLDER_COLOR = "gray"
COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS = (16, 16, -16, -16)
COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS = (16, 16, -16, -8)
COORDINATOR_DROP_LIST_ITEM_SPACING = 2
COORDINATOR_FILE_ITEM_LAYOUT_MARGINS = (12, 4, 12, 4)
COORDINATOR_FILE_ITEM_LAYOUT_SPACING = 12
COORDINATOR_DROP_ZONE_LAYOUT_MARGINS = (14, 14, 14, 14)
COORDINATOR_DROP_ZONE_LAYOUT_SPACING = 0
COORDINATOR_DROPZONE_BG_ACTIVE_ALPHA = 220
COORDINATOR_DROPZONE_OUTER_RADIUS = 12
COORDINATOR_DROPZONE_INNER_RADIUS = 10
COORDINATOR_DROPZONE_OUTER_RECT_ADJUST = (1, 1, -2, -2)
COORDINATOR_DROPZONE_BORDER_WIDTH = 2
COORDINATOR_DROPZONE_BORDER_DASH_PATTERN = (4, 3)
COORDINATOR_DROPZONE_INNER_RECT_ADJUST = (6, 6, -6, -6)
COORDINATOR_DROPZONE_BORDER_INACTIVE_ALPHA = 96
COORDINATOR_DROPZONE_BORDER_ACTIVE_ALPHA = 180
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
    if QPUSHBUTTON_GLOBAL_STYLESHEET in current_stylesheet:
        return
    merged_stylesheet = (
        f"{current_stylesheet}\n\n{QPUSHBUTTON_GLOBAL_STYLESHEET}".strip()
        if current_stylesheet
        else QPUSHBUTTON_GLOBAL_STYLESHEET
    )
    set_stylesheet(merged_stylesheet)
