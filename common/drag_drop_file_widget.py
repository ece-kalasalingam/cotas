"""Reusable drag-and-drop file widgets."""

from __future__ import annotations

from typing import Literal

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QDragEnterEvent, QDragLeaveEvent, QDragMoveEvent, QDropEvent, QPainter, QPalette, QPen
from PySide6.QtWidgets import QFrame, QListWidget, QWidget


class DragDropFileList(QListWidget):
    files_dropped = Signal(list)
    drag_state_changed = Signal(bool)
    browse_requested = Signal()

    def __init__(
        self,
        *,
        placeholder_color: str,
        placeholder_margins: tuple[int, int, int, int],
        placeholder_bottom_margins: tuple[int, int, int, int] | None = None,
        item_spacing: int = 0,
        drop_mode: Literal["single", "multiple"] = "multiple",
    ) -> None:
        super().__init__()
        self._placeholder_text = ""
        self._placeholder_color = placeholder_color
        self._placeholder_margins = placeholder_margins
        self._placeholder_bottom_margins = placeholder_bottom_margins
        self._drop_mode: Literal["single", "multiple"] = drop_mode
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setDropIndicatorShown(False)
        self.setSpacing(item_spacing)
        self.setAlternatingRowColors(False)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

    def set_placeholder_text(self, text: str) -> None:
        self._placeholder_text = text
        self.viewport().update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if not self._placeholder_text:
            return
        painter = QPainter(self.viewport())
        painter.setPen(QColor(self._placeholder_color))
        if self.count() == 0 or self._placeholder_bottom_margins is None:
            draw_rect = self.viewport().rect().adjusted(*self._placeholder_margins)
            draw_flags = Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap
        else:
            draw_rect = self.viewport().rect().adjusted(*self._placeholder_bottom_margins)
            draw_flags = Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom | Qt.TextFlag.TextWordWrap
        painter.drawText(draw_rect, draw_flags, self._placeholder_text)
        painter.end()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            self.drag_state_changed.emit(True)
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dragLeaveEvent(self, event: QDragLeaveEvent) -> None:
        self.drag_state_changed.emit(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        dropped = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if self._drop_mode == "single" and dropped:
            dropped = dropped[:1]
        self.drag_state_changed.emit(False)
        if dropped:
            self.files_dropped.emit(dropped)
            event.acceptProposedAction()
            return
        event.ignore()

    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.browse_requested.emit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class DragDropZoneFrame(QFrame):
    def __init__(
        self,
        *,
        outer_radius: int,
        inner_radius: int,
        bg_active_alpha: int,
        outer_rect_adjust: tuple[int, int, int, int],
        inner_rect_adjust: tuple[int, int, int, int],
        border_width: int,
        border_dash_pattern: tuple[int, int],
        border_inactive_alpha: int,
        border_active_alpha: int,
        background_from_parent_window: bool = True,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._outer_radius = outer_radius
        self._inner_radius = inner_radius
        self._bg_active_alpha = bg_active_alpha
        self._outer_rect_adjust = outer_rect_adjust
        self._inner_rect_adjust = inner_rect_adjust
        self._border_width = border_width
        self._border_dash_pattern = border_dash_pattern
        self._border_inactive_alpha = border_inactive_alpha
        self._border_active_alpha = border_active_alpha
        self._background_from_parent_window = background_from_parent_window

    def set_drag_active(self, active: bool) -> None:
        self.setProperty("dragActive", active)
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        palette = self.palette()
        host = self.parentWidget()
        host_palette = host.palette() if isinstance(host, QWidget) else palette
        active = bool(self.property("dragActive"))
        if self._background_from_parent_window:
            bg_color = host_palette.color(QPalette.ColorRole.Window)
        else:
            bg_color = palette.color(QPalette.ColorRole.AlternateBase)
        if active:
            bg_color.setAlpha(self._bg_active_alpha)
        border_color = (
            palette.color(QPalette.ColorRole.Highlight)
            if active
            else palette.color(QPalette.ColorRole.WindowText)
        )
        border_color.setAlpha(
            self._border_active_alpha if active else self._border_inactive_alpha
        )

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(bg_color)
        painter.drawRoundedRect(
            self.rect().adjusted(*self._outer_rect_adjust),
            self._outer_radius,
            self._outer_radius,
        )
        pen = QPen(border_color, self._border_width, Qt.PenStyle.DashLine)
        pen.setDashPattern(list(self._border_dash_pattern))
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(
            self.rect().adjusted(*self._inner_rect_adjust),
            self._inner_radius,
            self._inner_radius,
        )
        painter.end()
