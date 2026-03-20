"""PO Analysis placeholder module."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget

from common.constants import (
    MODULE_LEFT_PANE_CONTENT_MARGINS,
    MODULE_LEFT_PANE_LAYOUT_SPACING,
    MODULE_LEFT_PANE_SCROLLBAR_GUTTER,
    MODULE_LEFT_PANE_WIDTH_OFFSET,
)
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.output_panel import OutputPanelData
from common.texts import t
from common.ui_stylings import GLOBAL_QPUSHBUTTON_MIN_WIDTH

_LEFT_PANE_WIDTH = GLOBAL_QPUSHBUTTON_MIN_WIDTH + MODULE_LEFT_PANE_WIDTH_OFFSET


class POAnalysisModule(QWidget):
    status_changed = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                left_width=_LEFT_PANE_WIDTH,
                left_object_name="stepRail",
                right_object_name="coordinatorActiveCard",
                left_content_margins=MODULE_LEFT_PANE_CONTENT_MARGINS,
                left_layout_spacing=MODULE_LEFT_PANE_LAYOUT_SPACING,
                left_scrollbar_gutter=MODULE_LEFT_PANE_SCROLLBAR_GUTTER,
                show_left=False,
            ),
        )
        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_pane)
        self._ui_engine.set_right_widget(right_pane)
        self._label = QLabel(self)
        self._label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._label.setWordWrap(True)
        right_layout.addWidget(self._label)
        self.retranslate_ui()

    def retranslate_ui(self) -> None:
        self._label.setText(t("module.placeholder", title=t("module.po_analysis")))

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        _ = enabled

    def get_shared_outputs_data(self) -> OutputPanelData:
        return OutputPanelData(items=tuple())

