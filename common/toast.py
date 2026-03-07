from __future__ import annotations

import platform
from typing import Literal

from PySide6.QtCore import QPoint, Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication, QFrame, QGraphicsDropShadowEffect, QLabel, QVBoxLayout, QWidget
from common.constants import (
    TOAST_CONTENT_MARGIN_BOTTOM,
    TOAST_CONTENT_MARGIN_LEFT,
    TOAST_CONTENT_MARGIN_RIGHT,
    TOAST_CONTENT_MARGIN_TOP,
    TOAST_CONTENT_SPACING,
    TOAST_DEFAULT_DURATION_MS,
    TOAST_ERROR_DURATION_MS,
    TOAST_MAX_WIDTH,
    TOAST_MARGIN,
    TOAST_SHADOW_ALPHA,
    TOAST_SHADOW_BLUR_RADIUS,
    TOAST_SHADOW_OFFSET_X,
    TOAST_SHADOW_OFFSET_Y,
)

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
        self.setObjectName("toastFrame")
        self._message = message
        self._title_label: QLabel | None = None
        self._body_label: QLabel | None = None
        flags = Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool
        if parent is None:
            flags |= Qt.WindowType.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)
        # On Windows, translucent + shadow effects can trigger
        # UpdateLayeredWindowIndirect warnings with invalid dirty regions.
        if not _IS_WINDOWS:
            self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)

        bg, fg, border = _LEVEL_COLORS[level]
        self.setStyleSheet(
            f"#toastFrame {{ background: {bg}; border: 1px solid {border}; border-radius: 10px; }}"
            f"#toastTitle, #toastBody {{ color: {fg}; border: none; background: transparent; margin: 0; padding: 0; }}"
            "#toastTitle { font-weight: 600; }"
        )
        self.setMaximumWidth(TOAST_MAX_WIDTH)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            TOAST_CONTENT_MARGIN_LEFT,
            TOAST_CONTENT_MARGIN_TOP,
            TOAST_CONTENT_MARGIN_RIGHT,
            TOAST_CONTENT_MARGIN_BOTTOM,
        )
        layout.setSpacing(TOAST_CONTENT_SPACING)
        if title.strip():
            title_label = QLabel(title)
            title_label.setObjectName("toastTitle")
            layout.addWidget(title_label)
            self._title_label = title_label
        body = QLabel(message)
        body.setObjectName("toastBody")
        body.setWordWrap(False)
        layout.addWidget(body)
        self._body_label = body

        if not _IS_WINDOWS:
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(TOAST_SHADOW_BLUR_RADIUS)
            shadow.setOffset(TOAST_SHADOW_OFFSET_X, TOAST_SHADOW_OFFSET_Y)
            shadow.setColor(QColor(0, 0, 0, TOAST_SHADOW_ALPHA))
            self.setGraphicsEffect(shadow)

    def fit_width(self, max_width: int) -> None:
        if self._body_label is None:
            return
        horizontal_padding = TOAST_CONTENT_MARGIN_LEFT + TOAST_CONTENT_MARGIN_RIGHT
        text_lines = self._message.splitlines() or [self._message]
        body_metrics = self._body_label.fontMetrics()
        body_text_width = max(body_metrics.horizontalAdvance(line) for line in text_lines)
        title_text_width = 0
        if self._title_label is not None:
            title_text_width = self._title_label.fontMetrics().horizontalAdvance(self._title_label.text())

        desired_content_width = max(body_text_width, title_text_width)
        desired_total_width = desired_content_width + horizontal_padding
        target_width = min(desired_total_width, max_width)
        should_wrap = desired_total_width > max_width
        self._body_label.setWordWrap(should_wrap)
        self.setFixedWidth(target_width)
        self.adjustSize()


def _resolve_parent(parent: QWidget | None) -> QWidget | None:
    if parent is not None:
        return parent.window()
    app = QApplication.instance()
    if app is None or not isinstance(app, QApplication):
        return None

    active = app.activeWindow()
    if active is not None and active.isVisible() and not active.isMinimized():
        return active

    for widget in app.topLevelWidgets():
        if widget.isVisible() and not widget.isMinimized():
            return widget
    return None


def show_toast(
    parent: QWidget | None,
    message: str,
    *,
    title: str = "",
    level: ToastLevel = "info",
    duration_ms: int | None = None,
) -> None:
    host = _resolve_parent(parent)
    if host is not None and (not host.isVisible() or host.isMinimized()):
        return
    screen = QApplication.primaryScreen()
    screen_geometry = screen.availableGeometry() if screen is not None else None
    ttl = (
        duration_ms
        if duration_ms is not None
        else (TOAST_ERROR_DURATION_MS if level == "error" else TOAST_DEFAULT_DURATION_MS)
    )
    toast = _ToastWidget(host, title=title, message=message, level=level)

    available_width = (
        host.width() - (2 * TOAST_MARGIN)
        if host is not None
        else ((screen_geometry.width() - (2 * TOAST_MARGIN)) if screen_geometry is not None else TOAST_MAX_WIDTH)
    )
    if available_width > 0:
        toast.fit_width(min(TOAST_MAX_WIDTH, available_width))
    else:
        toast.adjustSize()

    margin = TOAST_MARGIN
    if host is not None:
        origin = host.mapToGlobal(QPoint(0, 0))
        x = origin.x() + host.width() - toast.width() - margin
        y = origin.y() + margin
    else:
        if screen_geometry is None:
            x = margin
            y = margin
        else:
            x = screen_geometry.x() + screen_geometry.width() - toast.width() - margin
            y = screen_geometry.y() + margin

    toast.move(max(x, margin), max(y, margin))
    toast.show()
    QTimer.singleShot(ttl, toast.close)
