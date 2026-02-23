from PySide6.QtWidgets import (
    QFrame,
    QToolButton,
    QVBoxLayout,
    QApplication,
    QStyle
)
from PySide6.QtCore import Qt


class Sidebar(QFrame):
    def __init__(self):
        super().__init__()

        # Let theme dictate colors — do not hardcode
        self.setObjectName("sidebar")

        # Horizontal icon + text
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        style = QApplication.style()

        self.menu_co = self._make_button(
            style.standardIcon(QStyle.StandardPixmap.SP_FileIcon),
            "CO Calculation"
        )

        self.menu_reports = self._make_button(
            style.standardIcon(QStyle.StandardPixmap.SP_DesktopIcon),
            "Reports"
        )

        self.menu_settings = self._make_button(
            style.standardIcon(QStyle.StandardPixmap.SP_DriveFDIcon),
            "Settings"
        )

        # Default checked
        self.menu_co.setChecked(True)

        # Add buttons
        layout.addWidget(self.menu_co)
        layout.addWidget(self.menu_reports)
        layout.addWidget(self.menu_settings)
        layout.addStretch()

    def _make_button(self, icon, text: str):
        btn = QToolButton()
        btn.setText(text)
        btn.setIcon(icon)
        btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setCheckable(True)
        btn.setAutoExclusive(True)
        return btn