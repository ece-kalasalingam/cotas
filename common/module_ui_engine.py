"""Shared module-shell UI engine for top region and footer region."""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtWidgets import (
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True)
class ModuleUIEngineConfig:
    top_object_name: str = "moduleTopRegion"
    footer_object_name: str = "moduleFooterRegion"
    footer_height: int = 220
    show_top: bool = True
    show_footer: bool = True


class ModuleUIEngine:
    """Composes a consistent module shell with top and footer blackbox panes."""

    def __init__(self, host: QWidget, *, config: ModuleUIEngineConfig) -> None:
        if not (config.show_top or config.show_footer):
            raise ValueError("At least one pane must be visible.")
        self._host = host
        self._config = config

        self.root_layout = QVBoxLayout(host)
        self.root_layout.setContentsMargins(0, 0, 0, 0)
        self.root_layout.setSpacing(10)

        self.top_widget: QWidget = QWidget()
        self.top_widget.setObjectName(config.top_object_name)
        self.top_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.root_layout.addWidget(self.top_widget, 1)

        self.footer_widget: QWidget = QWidget()
        self.footer_widget.setObjectName(config.footer_object_name)
        self._apply_footer_sizing(self.footer_widget)
        self.root_layout.addWidget(self.footer_widget, 0)
        self._apply_root_stretch()

        self.top_widget.setVisible(config.show_top)
        self.footer_widget.setVisible(config.show_footer)

    def set_top_widget(self, widget: QWidget, *, stretch: int = 1) -> None:
        self.root_layout.replaceWidget(self.top_widget, widget)
        self.top_widget.setParent(None)
        self.top_widget = widget
        if not self.top_widget.objectName():
            self.top_widget.setObjectName(self._config.top_object_name)
        self.top_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.root_layout.setStretchFactor(self.top_widget, stretch)
        self._apply_root_stretch(top_stretch=stretch)

    def set_footer_widget(self, widget: QWidget, *, stretch: int = 0) -> None:
        self.root_layout.replaceWidget(self.footer_widget, widget)
        self.footer_widget.setParent(None)
        self.footer_widget = widget
        self._apply_footer_sizing(self.footer_widget)
        if not self.footer_widget.objectName():
            self.footer_widget.setObjectName(self._config.footer_object_name)
        self.root_layout.setStretchFactor(self.footer_widget, stretch)
        self._apply_root_stretch(footer_stretch=stretch)

    def set_top_visible(self, visible: bool) -> None:
        self.top_widget.setVisible(visible)
        self._validate_any_pane_visible()

    def set_footer_visible(self, visible: bool) -> None:
        self.footer_widget.setVisible(visible)
        self._validate_any_pane_visible()

    def _validate_any_pane_visible(self) -> None:
        if (not self.top_widget.isHidden()) or (not self.footer_widget.isHidden()):
            return
        raise ValueError("At least one pane must be visible.")

    def _apply_footer_sizing(self, widget: QWidget) -> None:
        widget.setFixedHeight(self._config.footer_height)
        widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

    def _apply_root_stretch(self, *, top_stretch: int = 1, footer_stretch: int = 0) -> None:
        self.root_layout.setStretch(0, top_stretch)
        self.root_layout.setStretch(1, footer_stretch)
