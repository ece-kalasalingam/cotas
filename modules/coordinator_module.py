"""Course coordinator module for collecting Excel files via drag and drop."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Iterable

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDropEvent, QDragEnterEvent, QDragMoveEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from common.texts import t

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
REGISTER_HEADER_PATTERN = re.compile(r"\bregister\s*(?:number|no\.?)\b", re.IGNORECASE)

_logger = logging.getLogger(__name__)


def _path_key(path: Path) -> str:
    return str(path.resolve()).casefold()


def _is_supported_excel_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES


def _filter_excel_paths(paths: Iterable[str]) -> list[Path]:
    collected: list[Path] = []
    seen: set[str] = set()
    for value in paths:
        path = Path(value)
        if not _is_supported_excel_file(path):
            continue
        key = _path_key(path)
        if key in seen:
            continue
        seen.add(key)
        collected.append(path.resolve())
    return collected


def _has_valid_register_number(path: Path) -> bool:
    # openpyxl does not support legacy .xls files.
    if path.suffix.lower() == ".xls":
        return False
    try:
        from openpyxl import load_workbook
    except Exception:
        _logger.exception("openpyxl is unavailable while validating '%s'.", path)
        return False
    try:
        workbook = load_workbook(filename=path, read_only=True, data_only=True)
        try:
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for value in row:
                        if not isinstance(value, str):
                            continue
                        if REGISTER_HEADER_PATTERN.search(value.strip()):
                            return True
            return False
        finally:
            workbook.close()
    except Exception:
        _logger.exception("Failed to validate register header in '%s'.", path)
        return False


class _ExcelDropList(QListWidget):
    files_dropped = Signal(list)

    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setDropIndicatorShown(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setAlternatingRowColors(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        dropped = [url.toLocalFile() for url in urls if url.isLocalFile()]
        if dropped:
            self.files_dropped.emit(dropped)
            event.acceptProposedAction()
            return
        event.ignore()


class CoordinatorModule(QWidget):
    status_changed = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self._files: list[Path] = []
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)

        title = QLabel(t("coordinator.title"))
        title.setObjectName("coordinatorTitle")
        root.addWidget(title)

        hint = QLabel(t("coordinator.drop_hint"))
        hint.setWordWrap(True)
        root.addWidget(hint)

        self.drop_list = _ExcelDropList()
        self.drop_list.setObjectName("coordinatorDropList")
        self.drop_list.setMinimumHeight(260)
        self.drop_list.files_dropped.connect(self._on_files_dropped)

        frame = QFrame()
        frame.setObjectName("coordinatorDropFrame")
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(10, 10, 10, 10)
        frame_layout.addWidget(self.drop_list)
        root.addWidget(frame, 1)

        self.empty_hint = QLabel(t("coordinator.empty"))
        self.empty_hint.setObjectName("coordinatorEmptyHint")
        root.addWidget(self.empty_hint)

        self.remove_button = QPushButton(t("coordinator.remove_selected"))
        self.remove_button.clicked.connect(self._remove_selected)
        root.addWidget(self.remove_button)

        self.clear_button = QPushButton(t("coordinator.clear_all"))
        self.clear_button.clicked.connect(self._clear_all)
        root.addWidget(self.clear_button)

        panel_style = """
        QFrame#coordinatorDropFrame {
            border: 1px solid palette(mid);
            border-radius: 10px;
        }
        QListWidget#coordinatorDropList {
            border: 1px dashed palette(mid);
            border-radius: 10px;
            padding: 8px;
        }
        """

        self.setStyleSheet(panel_style)
        self._refresh_buttons()

    def _on_files_dropped(self, dropped_files: list[str]) -> None:
        accepted = _filter_excel_paths(dropped_files)
        existing = {_path_key(path) for path in self._files}
        added = 0
        duplicates = 0
        invalid_register: list[Path] = []
        for path in accepted:
            key = _path_key(path)
            if key in existing:
                duplicates += 1
                continue
            if not _has_valid_register_number(path):
                invalid_register.append(path)
                continue
            existing.add(key)
            self._files.append(path)
            item = QListWidgetItem(path.name)
            item.setToolTip(str(path))
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            self.drop_list.addItem(item)
            added += 1

        ignored = (len(dropped_files) - len(accepted)) + duplicates + len(invalid_register)
        self._refresh_buttons()
        if added:
            self.status_changed.emit(
                t("coordinator.status.added", added=added, total=len(self._files))
            )
        if duplicates:
            QMessageBox.information(
                self,
                t("coordinator.duplicate.title"),
                t("coordinator.duplicate.body", count=duplicates),
            )
        if invalid_register:
            file_names = "\n".join(path.name for path in invalid_register)
            QMessageBox.warning(
                self,
                t("coordinator.invalid_register.title"),
                t(
                    "coordinator.invalid_register.body",
                    count=len(invalid_register),
                    files=file_names,
                ),
            )
        if ignored:
            self.status_changed.emit(t("coordinator.status.ignored", count=ignored))

    def _remove_selected(self) -> None:
        selected = self.drop_list.selectedItems()
        if not selected:
            return

        remove_keys = {
            _path_key(Path(str(item.data(Qt.ItemDataRole.UserRole))))
            for item in selected
        }
        self._files = [path for path in self._files if _path_key(path) not in remove_keys]

        for item in selected:
            row = self.drop_list.row(item)
            self.drop_list.takeItem(row)

        self._refresh_buttons()
        self.status_changed.emit(t("coordinator.status.removed", count=len(selected)))

    def _clear_all(self) -> None:
        if not self._files:
            return
        total = len(self._files)
        self._files.clear()
        self.drop_list.clear()
        self._refresh_buttons()
        self.status_changed.emit(t("coordinator.status.cleared", count=total))

    def _refresh_buttons(self) -> None:
        has_files = bool(self._files)
        self.empty_hint.setVisible(not has_files)
        self.remove_button.setEnabled(has_files)
        self.clear_button.setEnabled(has_files)
