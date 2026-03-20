"""Shared module-shell UI engine for left pane, right pane, and footer."""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True)
class ModuleUIEngineConfig:
    left_width: int
    left_object_name: str
    right_object_name: str
    left_content_margins: tuple[int, int, int, int]
    left_layout_spacing: int
    left_scrollbar_gutter: int
    footer_height: int = 220
    show_left: bool = True
    show_right: bool = True
    show_footer: bool = True


class ModuleUIEngine:
    """Composes a consistent module shell with left, right, and footer blackbox panes."""

    def __init__(self, host: QWidget, *, config: ModuleUIEngineConfig) -> None:
        if not (config.show_left or config.show_right or config.show_footer):
            raise ValueError("At least one pane must be visible.")
        self._host = host
        self._config = config

        self.root_layout = QVBoxLayout(host)
        self.root_layout.setContentsMargins(0, 0, 0, 0)
        self.root_layout.setSpacing(10)

        self.top_row = QWidget()
        self.top_row_layout = QHBoxLayout(self.top_row)
        self.top_row_layout.setContentsMargins(0, 0, 0, 0)
        self.top_row_layout.setSpacing(0)
        self.root_layout.addWidget(self.top_row)

        self.left_widget: QWidget = QFrame()
        self.left_widget.setObjectName(config.left_object_name)

        self.left_scroll = QScrollArea()
        self.left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.left_scroll.setWidgetResizable(True)
        self.left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.left_scroll.setViewportMargins(0, 0, config.left_scrollbar_gutter, 0)
        self.left_scroll.setWidget(self.left_widget)
        self.left_scroll.setFixedWidth(config.left_width)
        self.top_row_layout.addWidget(self.left_scroll)

        self.right_widget: QWidget = QFrame()
        self.right_widget.setObjectName(config.right_object_name)
        self.top_row_layout.addWidget(self.right_widget, 1)

        self.footer_widget: QWidget = QWidget()
        self.footer_widget.setFixedHeight(config.footer_height)
        self.root_layout.addWidget(self.footer_widget, 0)

        self.left_scroll.setVisible(config.show_left)
        self.right_widget.setVisible(config.show_right)
        self.top_row.setVisible(config.show_left or config.show_right)
        self.footer_widget.setVisible(config.show_footer)

    def set_left_widget(self, widget: QWidget) -> None:
        old_widget = self.left_scroll.takeWidget()
        if old_widget is not None:
            old_widget.setParent(None)
        self.left_widget = widget
        if not self.left_widget.objectName():
            self.left_widget.setObjectName(self._config.left_object_name)
        self.left_scroll.setWidget(self.left_widget)

    def set_right_widget(self, widget: QWidget, *, stretch: int = 1) -> None:
        self.top_row_layout.replaceWidget(self.right_widget, widget)
        self.right_widget.setParent(None)
        self.right_widget = widget
        if not self.right_widget.objectName():
            self.right_widget.setObjectName(self._config.right_object_name)
        self.top_row_layout.setStretchFactor(self.right_widget, stretch)

    def set_footer_widget(self, widget: QWidget, *, stretch: int = 0) -> None:
        self.root_layout.replaceWidget(self.footer_widget, widget)
        self.footer_widget.setParent(None)
        self.footer_widget = widget
        self.root_layout.setStretchFactor(self.footer_widget, stretch)

    def set_left_visible(self, visible: bool) -> None:
        self.left_scroll.setVisible(visible)
        self._sync_top_row_visibility()

    def set_right_visible(self, visible: bool) -> None:
        self.right_widget.setVisible(visible)
        self._sync_top_row_visibility()

    def set_footer_visible(self, visible: bool) -> None:
        self.footer_widget.setVisible(visible)
        self._validate_any_pane_visible()

    def _sync_top_row_visibility(self) -> None:
        self.top_row.setVisible(self.left_scroll.isVisible() or self.right_widget.isVisible())
        self._validate_any_pane_visible()

    def _validate_any_pane_visible(self) -> None:
        if self.left_scroll.isVisible() or self.right_widget.isVisible() or self.footer_widget.isVisible():
            return
        raise ValueError("At least one pane must be visible.")
