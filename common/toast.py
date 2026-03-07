from __future__ import annotations

import platform
from typing import Literal

from PySide6.QtCore import QPoint, Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication, QFrame, QGraphicsDropShadowEffect, QLabel, QVBoxLayout, QWidget

ToastLevel = Literal["info", "success", "warning", "error"]

_LEVEL_COLORS: dict[str, tuple[str, str, str]] = {
    "info": ("#E6F0FF", "#1E3A8A", "#93C5FD"),
    "success": ("#E7F8EF", "#14532D", "#86EFAC"),
    "warning": ("#FFF7E8", "#7C2D12", "#FCD34D"),
    "error": ("#FDECEC", "#7F1D1D", "#FCA5A5"),
}
_IS_WINDOWS = platform.system().lower().startswith("win")


class _ToastWidget(QFrame):
    def __init__(self, parent: QWidget | None, title: str, message: str, level: ToastLevel) -> None:
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)
        # On Windows, translucent + shadow effects can trigger
        # UpdateLayeredWindowIndirect warnings with invalid dirty regions.
        if not _IS_WINDOWS:
            self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)

        bg, fg, border = _LEVEL_COLORS[level]
        self.setStyleSheet(
            f"QFrame {{ background: {bg}; border: 1px solid {border}; border-radius: 10px; }}"
            f"QLabel {{ color: {fg}; }}"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)
        if title.strip():
            title_label = QLabel(title)
            title_label.setStyleSheet("font-weight: 600;")
            layout.addWidget(title_label)
        body = QLabel(message)
        body.setWordWrap(True)
        layout.addWidget(body)

        if not _IS_WINDOWS:
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(24)
            shadow.setOffset(0, 6)
            shadow.setColor(QColor(0, 0, 0, 60))
            self.setGraphicsEffect(shadow)


def _resolve_parent(parent: QWidget | None) -> QWidget | None:
    if parent is not None:
        return parent
    app = QApplication.instance()
    if app is None or not isinstance(app, QApplication):
        return None
    return app.activeWindow()


def show_toast(
    parent: QWidget | None,
    message: str,
    *,
    title: str = "",
    level: ToastLevel = "info",
    duration_ms: int | None = None,
) -> None:
    host = _resolve_parent(parent)
    ttl = duration_ms if duration_ms is not None else (4500 if level == "error" else 3000)
    toast = _ToastWidget(host, title=title, message=message, level=level)
    toast.adjustSize()

    margin = 16
    if host is not None:
        origin = host.mapToGlobal(QPoint(0, 0))
        x = origin.x() + host.width() - toast.width() - margin
        y = origin.y() + margin
    else:
        screen = QApplication.primaryScreen()
        if screen is None:
            x = margin
            y = margin
        else:
            geometry = screen.availableGeometry()
            x = geometry.x() + geometry.width() - toast.width() - margin
            y = geometry.y() + margin

    toast.move(max(x, margin), max(y, margin))
    toast.show()
    QTimer.singleShot(ttl, toast.close)
