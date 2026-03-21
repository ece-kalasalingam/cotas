"""PO Analysis placeholder module."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget

from common.constants import INSTRUCTOR_INFO_TAB_FIXED_HEIGHT
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.output_panel import OutputPanelData
from common.texts import t


class POAnalysisModule(QWidget):
    status_changed = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                top_object_name="poAnalysisTopRegion",
                footer_height=INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
            ),
        )
        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_pane)
        self._ui_engine.set_top_widget(right_pane)
        self._label = QLabel(self)
        self._label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._label.setWordWrap(True)
        right_layout.addWidget(self._label)
        self.retranslate_ui()

    def retranslate_ui(self) -> None:
        self._label.setText(t("module.placeholder", title=t("module.po_analysis")))

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self._ui_engine.set_footer_visible(not enabled)

    def get_shared_outputs_data(self) -> OutputPanelData:
        return OutputPanelData(items=tuple())

