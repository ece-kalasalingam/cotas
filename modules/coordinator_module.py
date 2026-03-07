"""Course coordinator module for collecting Excel files via drag and drop."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
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
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from common.exceptions import JobCancelledError
from common.jobs import CancellationToken, generate_job_id
from common.qt_jobs import run_in_background
from common.texts import t
from common.toast import show_toast
from common.utils import emit_user_status, log_process_message

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
REGISTER_HEADER_PATTERN = re.compile(r"\bregister\s*(?:number|no\.?)\b", re.IGNORECASE)

_logger = logging.getLogger(__name__)


@dataclass(slots=True)
class CoordinatorWorkflowState:
    busy: bool = False
    active_job_id: str | None = None

    def set_busy(self, value: bool, *, job_id: str | None = None) -> None:
        self.busy = value
        self.active_job_id = job_id if value else None


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


def _analyze_dropped_files(
    dropped_files: list[str],
    *,
    existing_keys: set[str],
    token: CancellationToken,
) -> dict[str, object]:
    accepted = _filter_excel_paths(dropped_files)
    seen = set(existing_keys)
    added: list[str] = []
    duplicates = 0
    invalid_register: list[str] = []

    for path in accepted:
        token.raise_if_cancelled()
        key = _path_key(path)
        if key in seen:
            duplicates += 1
            continue
        if not _has_valid_register_number(path):
            invalid_register.append(str(path))
            continue
        seen.add(key)
        added.append(str(path))

    ignored = (len(dropped_files) - len(accepted)) + duplicates + len(invalid_register)
    return {
        "added": added,
        "duplicates": duplicates,
        "invalid_register": invalid_register,
        "ignored": ignored,
    }


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
        self._logger = _logger
        self.state = CoordinatorWorkflowState()
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._pending_drop_batches: list[list[str]] = []
        self._build_ui()
        self.retranslate_ui()
        self._refresh_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)

        self.title_label = QLabel()
        self.title_label.setObjectName("coordinatorTitle")
        root.addWidget(self.title_label)

        self.hint_label = QLabel()
        self.hint_label.setWordWrap(True)
        root.addWidget(self.hint_label)

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

        self.empty_hint = QLabel()
        self.empty_hint.setObjectName("coordinatorEmptyHint")
        root.addWidget(self.empty_hint)

        self.remove_button = QPushButton()
        self.remove_button.clicked.connect(self._remove_selected)
        root.addWidget(self.remove_button)

        self.clear_button = QPushButton()
        self.clear_button.clicked.connect(self._clear_all)
        root.addWidget(self.clear_button)

        self.cancel_button = QPushButton()
        self.cancel_button.clicked.connect(self._cancel_processing)
        root.addWidget(self.cancel_button)

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

    def retranslate_ui(self) -> None:
        self.title_label.setText(t("coordinator.title"))
        self.hint_label.setText(t("coordinator.drop_hint"))
        self.empty_hint.setText(t("coordinator.empty"))
        self.remove_button.setText(t("coordinator.remove_selected"))
        self.clear_button.setText(t("coordinator.clear_all"))
        self.cancel_button.setText(t("coordinator.cancel_processing"))

    def _publish_status(self, message: str) -> None:
        emit_user_status(self.status_changed, message, logger=self._logger)

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        self._refresh_ui()

    def _refresh_ui(self) -> None:
        has_files = bool(self._files)
        self.empty_hint.setVisible(not has_files)
        self.remove_button.setEnabled(has_files and not self.state.busy)
        self.clear_button.setEnabled(has_files and not self.state.busy)
        self.cancel_button.setEnabled(self.state.busy)

    def _cancel_processing(self) -> None:
        token = self._cancel_token
        if token is None:
            return
        token.cancel()
        self._publish_status(t("coordinator.status.cancelling"))

    def _drain_next_batch(self) -> None:
        if self.state.busy or not self._pending_drop_batches:
            return
        next_batch = self._pending_drop_batches.pop(0)
        self._process_files_async(next_batch)

    def _process_files_async(self, dropped_files: list[str]) -> None:
        if not dropped_files:
            return
        if self.state.busy:
            self._pending_drop_batches.append(dropped_files)
            self._publish_status(t("coordinator.status.queued", count=len(dropped_files)))
            return

        process_name = "collecting coordinator files"
        token = CancellationToken()
        job_id = generate_job_id()
        existing_keys = {_path_key(path) for path in self._files}
        self._cancel_token = token
        self._set_busy(True, job_id=job_id)
        self._publish_status(t("coordinator.status.processing_started"))

        def _finalize(job: object) -> None:
            if job in self._active_jobs:
                self._active_jobs.remove(job)
            self._cancel_token = None
            self._set_busy(False)
            self._drain_next_batch()

        def _on_finished(result: object) -> None:
            try:
                if not isinstance(result, dict):
                    raise RuntimeError("Coordinator processing returned unexpected result type.")
                added_paths = [Path(value) for value in result.get("added", [])]
                duplicates = int(result.get("duplicates", 0))
                invalid_paths = [Path(value) for value in result.get("invalid_register", [])]
                ignored = int(result.get("ignored", 0))

                for path in added_paths:
                    self._files.append(path)
                    item = QListWidgetItem(path.name)
                    item.setToolTip(str(path))
                    item.setData(Qt.ItemDataRole.UserRole, str(path))
                    self.drop_list.addItem(item)

                if added_paths:
                    self._publish_status(
                        t("coordinator.status.added", added=len(added_paths), total=len(self._files))
                    )
                if duplicates:
                    show_toast(
                        self,
                        t("coordinator.duplicate.body", count=duplicates),
                        title=t("coordinator.duplicate.title"),
                        level="info",
                    )
                if invalid_paths:
                    file_names = "\n".join(path.name for path in invalid_paths)
                    show_toast(
                        self,
                        t(
                            "coordinator.invalid_register.body",
                            count=len(invalid_paths),
                            files=file_names,
                        ),
                        title=t("coordinator.invalid_register.title"),
                        level="warning",
                    )
                if ignored:
                    self._publish_status(t("coordinator.status.ignored", count=ignored))

                log_process_message(
                    process_name,
                    logger=self._logger,
                    success_message=(
                        f"{process_name} completed successfully. "
                        f"added={len(added_paths)}, duplicates={duplicates}, "
                        f"invalid={len(invalid_paths)}, ignored={ignored}"
                    ),
                    job_id=job_id,
                    step_id="coordinator_collect_files",
                )
            finally:
                _finalize(job)

        def _on_failed(exc: Exception) -> None:
            try:
                if isinstance(exc, JobCancelledError):
                    self._publish_status(t("coordinator.status.operation_cancelled"))
                    self._logger.info(
                        "%s cancelled by user/system request.",
                        process_name,
                        extra={
                            "user_message": t("coordinator.status.operation_cancelled"),
                            "job_id": job_id,
                            "step_id": "coordinator_collect_files",
                        },
                    )
                    return
                log_process_message(
                    process_name,
                    logger=self._logger,
                    error=exc,
                    user_error_message=t("coordinator.status.processing_failed"),
                    job_id=job_id,
                    step_id="coordinator_collect_files",
                )
                show_toast(
                    self,
                    t("coordinator.status.processing_failed"),
                    title=t("coordinator.title"),
                    level="error",
                )
            finally:
                _finalize(job)

        job = run_in_background(
            _analyze_dropped_files,
            dropped_files,
            existing_keys=existing_keys,
            token=token,
            on_finished=_on_finished,
            on_failed=_on_failed,
        )
        self._active_jobs.append(job)

    def _on_files_dropped(self, dropped_files: list[str]) -> None:
        self._process_files_async(dropped_files)

    def _remove_selected(self) -> None:
        if self.state.busy:
            return
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

        self._refresh_ui()
        self._publish_status(t("coordinator.status.removed", count=len(selected)))
        log_process_message(
            "removing selected coordinator files",
            logger=self._logger,
            success_message=f"removing selected coordinator files completed successfully. removed={len(selected)}",
        )

    def _clear_all(self) -> None:
        if self.state.busy:
            return
        if not self._files:
            return
        total = len(self._files)
        self._files.clear()
        self.drop_list.clear()
        self._refresh_ui()
        self._publish_status(t("coordinator.status.cleared", count=total))
        log_process_message(
            "clearing coordinator files",
            logger=self._logger,
            success_message=f"clearing coordinator files completed successfully. removed={total}",
        )
