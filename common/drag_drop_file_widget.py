"""Reusable drag-and-drop file widgets."""

from __future__ import annotations

from collections.abc import Callable, Iterable
from pathlib import Path
from typing import Literal

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import (
    QColor,
    QDragEnterEvent,
    QDragLeaveEvent,
    QDragMoveEvent,
    QDropEvent,
    QPainter,
    QPalette,
    QPen,
)
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from common.removable_file_item_widget import RemovableFileItemWidget
from common.ui_stylings import (
    COORDINATOR_DROP_LIST_ITEM_SPACING,
    COORDINATOR_DROP_ZONE_LAYOUT_MARGINS,
    COORDINATOR_DROP_ZONE_LAYOUT_SPACING,
    COORDINATOR_DROPZONE_BG_ACTIVE_ALPHA,
    COORDINATOR_DROPZONE_BORDER_ACTIVE_ALPHA,
    COORDINATOR_DROPZONE_BORDER_DASH_PATTERN,
    COORDINATOR_DROPZONE_BORDER_INACTIVE_ALPHA,
    COORDINATOR_DROPZONE_BORDER_WIDTH,
    COORDINATOR_DROPZONE_INNER_RADIUS,
    COORDINATOR_DROPZONE_INNER_RECT_ADJUST,
    COORDINATOR_DROPZONE_OUTER_RADIUS,
    COORDINATOR_DROPZONE_OUTER_RECT_ADJUST,
    COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
    COORDINATOR_LIST_PLACEHOLDER_COLOR,
    COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
)


class DragDropFileList(QListWidget):
    files_dropped = Signal(list)
    drag_state_changed = Signal(bool)
    browse_requested = Signal()
    DEFAULT_PLACEHOLDER_TEXT = "Drag and Drop, or press Ctrl + O, or double-click to add files"

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
        self._placeholder_text = self.DEFAULT_PLACEHOLDER_TEXT
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


class ManagedDropFileWidget(QWidget):
    """Reusable drop widget with internal file list + remove buttons."""

    files_dropped = Signal(list)
    files_changed = Signal(list)
    files_rejected = Signal(list)
    browse_requested = Signal()
    submit_requested = Signal()

    def __init__(
        self,
        *,
        drop_mode: Literal["single", "multiple"] = "multiple",
        remove_fallback_text: str = "Remove",
        open_file_tooltip: str = "Open File",
        open_folder_tooltip: str = "Open Folder",
        remove_tooltip: str = "Remove File",
        allow_non_local_sources: bool = False,
        allowed_extensions: Iterable[str] | None = None,
        allowed_filenames: Iterable[str] | None = None,
        file_filter: Callable[[str], bool] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._drop_mode: Literal["single", "multiple"] = drop_mode
        self._remove_fallback_text = remove_fallback_text
        self._open_file_tooltip = open_file_tooltip
        self._open_folder_tooltip = open_folder_tooltip
        self._remove_tooltip = remove_tooltip
        self._allow_non_local_sources = allow_non_local_sources
        self._files: list[str] = []
        self._allowed_extensions = self._normalize_extensions(allowed_extensions)
        self._allowed_filenames = self._normalize_filenames(allowed_filenames)
        self._file_filter = file_filter
        self._summary_text_builder: Callable[[int], str] = lambda count: f"Files: {count}"
        self._submit_allowed = True

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.drop_zone = DragDropZoneFrame(
            outer_radius=COORDINATOR_DROPZONE_OUTER_RADIUS,
            inner_radius=COORDINATOR_DROPZONE_INNER_RADIUS,
            bg_active_alpha=COORDINATOR_DROPZONE_BG_ACTIVE_ALPHA,
            outer_rect_adjust=COORDINATOR_DROPZONE_OUTER_RECT_ADJUST,
            inner_rect_adjust=COORDINATOR_DROPZONE_INNER_RECT_ADJUST,
            border_width=COORDINATOR_DROPZONE_BORDER_WIDTH,
            border_dash_pattern=COORDINATOR_DROPZONE_BORDER_DASH_PATTERN,
            border_inactive_alpha=COORDINATOR_DROPZONE_BORDER_INACTIVE_ALPHA,
            border_active_alpha=COORDINATOR_DROPZONE_BORDER_ACTIVE_ALPHA,
            background_from_parent_window=True,
            parent=self,
        )
        self.drop_zone.setProperty("dragActive", False)
        zone_layout = QVBoxLayout(self.drop_zone)
        zone_layout.setContentsMargins(*COORDINATOR_DROP_ZONE_LAYOUT_MARGINS)
        zone_layout.setSpacing(COORDINATOR_DROP_ZONE_LAYOUT_SPACING)
        self.drop_list = DragDropFileList(
            placeholder_color=COORDINATOR_LIST_PLACEHOLDER_COLOR,
            placeholder_margins=COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
            placeholder_bottom_margins=COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
            item_spacing=COORDINATOR_DROP_LIST_ITEM_SPACING,
            drop_mode=drop_mode,
        )
        self.drop_list.drag_state_changed.connect(self.drop_zone.set_drag_active)
        self.drop_list.files_dropped.connect(self._on_files_dropped)
        self.drop_list.browse_requested.connect(self.browse_requested.emit)
        zone_layout.addWidget(self.drop_list)
        root.addWidget(self.drop_zone)

        controls_row = QHBoxLayout()
        controls_row.setContentsMargins(0, 0, 0, 0)
        controls_row.setSpacing(8)
        self.summary_label = QLabel()
        controls_row.addWidget(self.summary_label)

        self.clear_button = QPushButton("Clear All")
        self.clear_button.setObjectName("clearAllLink")
        self.clear_button.setAutoDefault(False)
        self.clear_button.setDefault(False)
        self.clear_button.setFlat(True)
        self.clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_button.setStyleSheet(
            """
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
            """
        )
        self.clear_button.clicked.connect(self.clear_files)
        controls_row.addWidget(self.clear_button)
        controls_row.addStretch(1)

        self.submit_button = QPushButton("Submit")
        self.submit_button.setAutoDefault(False)
        self.submit_button.setDefault(False)
        self.submit_button.clicked.connect(self.submit_requested.emit)
        controls_row.addWidget(self.submit_button)
        root.addLayout(controls_row)
        self._update_summary_label()
        self._update_clear_button_state()
        self._update_submit_button_state()

    def files(self) -> list[str]:
        return list(self._files)

    def set_summary_text_builder(self, builder: Callable[[int], str]) -> None:
        self._summary_text_builder = builder
        self._update_summary_label()

    def set_clear_button_text(self, text: str) -> None:
        self.clear_button.setText(text)

    def set_submit_button_text(self, text: str) -> None:
        self.submit_button.setText(text)

    def set_submit_allowed(self, allowed: bool) -> None:
        self._submit_allowed = bool(allowed)
        self._update_submit_button_state()

    def set_allowed_extensions(self, extensions: Iterable[str] | None) -> None:
        self._allowed_extensions = self._normalize_extensions(extensions)

    def set_allowed_filenames(self, names: Iterable[str] | None) -> None:
        self._allowed_filenames = self._normalize_filenames(names)

    def set_file_filter(self, predicate: Callable[[str], bool] | None) -> None:
        self._file_filter = predicate

    def set_allow_non_local_sources(self, value: bool) -> None:
        self._allow_non_local_sources = bool(value)

    def clear_files(self) -> None:
        if not self._files and self.drop_list.count() == 0:
            return
        self._files.clear()
        self.drop_list.clear()
        self._update_summary_label()
        self._update_clear_button_state()
        self._update_submit_button_state()
        self.files_changed.emit(self.files())

    def set_files(self, paths: list[str]) -> None:
        self._files.clear()
        self.drop_list.clear()
        added = self.add_files(paths, emit_drop=False)
        if not added:
            self._update_summary_label()
            self._update_clear_button_state()
            self._update_submit_button_state()
            self.files_changed.emit(self.files())

    def add_files(self, paths: list[str], *, emit_drop: bool = True) -> list[str]:
        normalized = [path for path in paths if path]
        if not normalized:
            return []
        accepted: list[str] = []
        rejected: list[str] = []
        for path in normalized:
            if self._accepts_path(path):
                accepted.append(path)
            else:
                rejected.append(path)
        normalized = accepted
        if not normalized:
            if rejected:
                self.files_rejected.emit(list(rejected))
            return []
        existing = set(self._files)
        deduped: list[str] = []
        for path in normalized:
            if path in existing:
                rejected.append(path)
                continue
            deduped.append(path)
            existing.add(path)
        normalized = deduped
        if rejected:
            self.files_rejected.emit(list(rejected))
        if not normalized:
            return []
        if self._drop_mode == "single":
            normalized = normalized[:1]
            self._files = []
            self.drop_list.clear()
        added: list[str] = []
        for path in normalized:
            self._files.append(path)
            added.append(path)
            self._append_row(path)
        if added:
            self._update_summary_label()
            self._update_clear_button_state()
            self._update_submit_button_state()
            if emit_drop:
                self.files_dropped.emit(list(added))
            self.files_changed.emit(self.files())
        return added

    def _append_row(self, file_path: str) -> None:
        item = QListWidgetItem()
        item.setToolTip(file_path)
        item.setData(Qt.ItemDataRole.UserRole, file_path)
        self.drop_list.addItem(item)
        row_widget = RemovableFileItemWidget(
            file_path,
            remove_fallback_text=self._remove_fallback_text,
            open_file_tooltip=self._open_file_tooltip,
            open_folder_tooltip=self._open_folder_tooltip,
            remove_tooltip=self._remove_tooltip,
            parent=self.drop_list,
        )
        row_widget.removed.connect(self._remove_path)
        item.setSizeHint(row_widget.sizeHint())
        self.drop_list.setItemWidget(item, row_widget)

    def _remove_path(self, file_path: str) -> None:
        if file_path not in self._files:
            return
        self._files = [value for value in self._files if value != file_path]
        for row in range(self.drop_list.count() - 1, -1, -1):
            item = self.drop_list.item(row)
            if item is None:
                continue
            path = item.data(Qt.ItemDataRole.UserRole)
            if isinstance(path, str) and path == file_path:
                self.drop_list.takeItem(row)
        self._update_summary_label()
        self._update_clear_button_state()
        self._update_submit_button_state()
        self.files_changed.emit(self.files())

    def setEnabled(self, enabled: bool) -> None:
        super().setEnabled(enabled)
        self._update_submit_button_state()

    def _on_files_dropped(self, paths: list[str]) -> None:
        self.add_files(paths, emit_drop=True)

    @staticmethod
    def _normalize_extensions(values: Iterable[str] | None) -> set[str] | None:
        if values is None:
            return None
        normalized: set[str] = set()
        for value in values:
            token = value.strip().lower()
            if not token:
                continue
            if not token.startswith("."):
                token = f".{token}"
            normalized.add(token)
        return normalized or None

    @staticmethod
    def _normalize_filenames(values: Iterable[str] | None) -> set[str] | None:
        if values is None:
            return None
        normalized = {value.strip().lower() for value in values if value.strip()}
        return normalized or None

    def _accepts_path(self, file_path: str) -> bool:
        if not self._allow_non_local_sources and RemovableFileItemWidget._normalize_local_path(file_path) is None:
            return False
        path = Path(file_path)
        if self._allowed_extensions is not None and path.suffix.lower() not in self._allowed_extensions:
            return False
        if self._allowed_filenames is not None and path.name.lower() not in self._allowed_filenames:
            return False
        if self._file_filter is not None and not self._file_filter(file_path):
            return False
        return True

    def _update_summary_label(self) -> None:
        count = len(self._files)
        self.summary_label.setText(self._summary_text_builder(count))
        self.summary_label.setEnabled(count > 0)

    def _update_clear_button_state(self) -> None:
        self.clear_button.setEnabled(bool(self._files))

    def _update_submit_button_state(self) -> None:
        self.submit_button.setEnabled(
            bool(self._files) and self._submit_allowed and self.isEnabled()
        )
