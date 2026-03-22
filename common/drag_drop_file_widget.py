"""Reusable drag-and-drop file widgets."""

from __future__ import annotations

from collections.abc import Callable, Iterable
from pathlib import Path
from typing import Literal

from PySide6.QtCore import QEvent, QObject, Qt, Signal
from PySide6.QtGui import (
    QDragEnterEvent,
    QDragLeaveEvent,
    QDragMoveEvent,
    QDropEvent,
    QPainter,
    QPalette,
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
    COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
    COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
)


class DragDropFileList(QListWidget):
    files_dropped = Signal(list)
    items_reordered = Signal(list)
    drag_state_changed = Signal(bool)
    browse_requested = Signal()
    DEFAULT_PLACEHOLDER_TEXT = "Drag and Drop, or press Ctrl + O, or single-click to add files"

    def __init__(
        self,
        *,
        placeholder_margins: tuple[int, int, int, int],
        placeholder_bottom_margins: tuple[int, int, int, int] | None = None,
        item_spacing: int = 0,
        drop_mode: Literal["single", "multiple"] = "multiple",
    ) -> None:
        super().__init__()
        self._placeholder_text = self.DEFAULT_PLACEHOLDER_TEXT
        self._placeholder_margins = placeholder_margins
        self._placeholder_bottom_margins = placeholder_bottom_margins
        self._drop_mode: Literal["single", "multiple"] = drop_mode
        self.setAcceptDrops(True)
        self.setDragEnabled(drop_mode == "multiple")
        if drop_mode == "multiple":
            self.setDragDropMode(QListWidget.DragDropMode.InternalMove)
            self.setDefaultDropAction(Qt.DropAction.MoveAction)
            self.setDragDropOverwriteMode(False)
        else:
            self.setDragDropMode(QListWidget.DragDropMode.NoDragDrop)
        self.setDropIndicatorShown(False)
        self.setSpacing(item_spacing)
        self.setAlternatingRowColors(False)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        # Keep list background aligned with surrounding container (avoid white Base block).
        window_color = self.palette().color(QPalette.ColorRole.Window)
        list_palette = self.palette()
        list_palette.setColor(QPalette.ColorRole.Base, window_color)
        self.setPalette(list_palette)
        self.setAutoFillBackground(True)
        viewport = self.viewport()
        viewport_palette = viewport.palette()
        viewport_palette.setColor(QPalette.ColorRole.Base, window_color)
        viewport.setPalette(viewport_palette)
        viewport.setAutoFillBackground(True)

    def set_placeholder_text(self, text: str) -> None:
        self._placeholder_text = text
        self.viewport().update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if not self._placeholder_text:
            return
        painter = QPainter(self.viewport())
        painter.setPen(self.palette().color(QPalette.ColorRole.WindowText))
        if self.count() == 0 or self._placeholder_bottom_margins is None:
            draw_rect = self.viewport().rect().adjusted(*self._placeholder_margins)
            draw_flags = Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap
        else:
            draw_rect = self.viewport().rect().adjusted(*self._placeholder_bottom_margins)
            draw_flags = Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom | Qt.TextFlag.TextWordWrap
        painter.drawText(draw_rect, draw_flags, self._placeholder_text)
        painter.end()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        event_source = getattr(event, "source", None)
        source = event_source() if callable(event_source) else None
        if self._drop_mode == "multiple" and source is self:
            event.acceptProposedAction()
            return
        if event.mimeData().hasUrls():
            self.drag_state_changed.emit(True)
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        event_source = getattr(event, "source", None)
        source = event_source() if callable(event_source) else None
        if self._drop_mode == "multiple" and source is self:
            event.acceptProposedAction()
            return
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dragLeaveEvent(self, event: QDragLeaveEvent) -> None:
        self.drag_state_changed.emit(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        event_source = getattr(event, "source", None)
        source = event_source() if callable(event_source) else None
        if self._drop_mode == "multiple" and source is self:
            super().dropEvent(event)
            event.acceptProposedAction()
            self.items_reordered.emit(self._ordered_item_paths())
            return
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

    def _ordered_item_paths(self) -> list[str]:
        ordered: list[str] = []
        for row in range(self.count()):
            item = self.item(row)
            if item is None:
                continue
            value = item.data(Qt.ItemDataRole.UserRole)
            if isinstance(value, str) and value:
                ordered.append(value)
        return ordered

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self.itemAt(event.pos()) is None:
            self.browse_requested.emit()
            event.accept()
            return
        super().mousePressEvent(event)


class DragDropZoneFrame(QFrame):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)

    def set_drag_active(self, active: bool) -> None:
        self.setProperty("dragActive", active)
        self.update()


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
        root.setSpacing(10)

        self.drop_zone = DragDropZoneFrame(parent=self)
        self.drop_zone.setObjectName("managedDropZoneFrame")
        self.drop_zone.setFrameShape(QFrame.Shape.StyledPanel)
        self.drop_zone.setFrameShadow(QFrame.Shadow.Raised)
        self.drop_zone.setMouseTracking(True)
        self.drop_zone.setCursor(Qt.CursorShape.PointingHandCursor)
        self.drop_zone.setProperty("dragActive", False)
        self.drop_zone.installEventFilter(self)

        zone_layout = QVBoxLayout(self.drop_zone)
        zone_layout.setContentsMargins(*COORDINATOR_DROP_ZONE_LAYOUT_MARGINS)
        zone_layout.setSpacing(COORDINATOR_DROP_ZONE_LAYOUT_SPACING)
        self.drop_list = DragDropFileList(
            placeholder_margins=COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
            placeholder_bottom_margins=COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
            item_spacing=COORDINATOR_DROP_LIST_ITEM_SPACING,
            drop_mode=drop_mode,
        )
        self.drop_list.drag_state_changed.connect(self.drop_zone.set_drag_active)
        self.drop_list.files_dropped.connect(self._on_files_dropped)
        self.drop_list.items_reordered.connect(self._on_items_reordered)
        self.drop_list.browse_requested.connect(self.browse_requested.emit)
        self.drop_list.setCursor(Qt.CursorShape.PointingHandCursor)
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

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        drop_zone = getattr(self, "drop_zone", None)
        if drop_zone is None:
            return super().eventFilter(watched, event)
        if watched is drop_zone:
            event_type = event.type()
            if event_type == event.Type.Enter:
                drop_zone.setProperty("hoverActive", True)
                drop_zone.update()
            elif event_type == event.Type.Leave:
                drop_zone.setProperty("hoverActive", False)
                drop_zone.update()
        return super().eventFilter(watched, event)

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

    def _on_items_reordered(self, ordered_paths: list[str]) -> None:
        if self._drop_mode != "multiple":
            return
        if self._files:
            existing = set(self._files)
            reordered = [path for path in ordered_paths if path in existing]
            if len(reordered) == len(self._files):
                self._files = reordered
        self.files_changed.emit(list(ordered_paths))

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
