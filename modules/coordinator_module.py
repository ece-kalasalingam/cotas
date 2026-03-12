"""Course coordinator module for collecting Final CO report Excel files."""

from __future__ import annotations

import logging
import json
from dataclasses import dataclass
from html import escape
from pathlib import Path
from typing import Iterable

from PySide6.QtCore import Qt, QSize, QUrl, Signal
from PySide6.QtGui import (
    QColor,
    QDesktopServices,
    QDropEvent,
    QDragEnterEvent,
    QDragLeaveEvent,
    QDragMoveEvent,
    QFont,
    QKeySequence,
    QPainter,
    QShortcut,
)
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QPushButton,
    QStyle,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from common.constants import (
    APP_NAME,
    CO_REPORT_DIRECT_SHEET_SUFFIX,
    CO_REPORT_INDIRECT_SHEET_SUFFIX,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SECTION_KEY,
    COURSE_METADATA_SHEET,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    INSTRUCTOR_CARD_MARGIN,
    INSTRUCTOR_CARD_SPACING,
    INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    INSTRUCTOR_INFO_TAB_LAYOUT_MARGINS,
    INSTRUCTOR_INFO_TAB_LAYOUT_SPACING,
    SYSTEM_HASH_SHEET,
    SYSTEM_HASH_TEMPLATE_HASH_HEADER,
    SYSTEM_HASH_TEMPLATE_ID_HEADER,
    SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
    SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
    SYSTEM_REPORT_INTEGRITY_SHEET,
    UI_FONT_FAMILY,
)
from common.exceptions import JobCancelledError
from common.jobs import CancellationToken, generate_job_id
from common.qt_jobs import run_in_background
from common.texts import t
from common.toast import show_toast
from common.utils import (
    emit_user_status,
    log_process_message,
    normalize,
    remember_dialog_dir,
    remember_dialog_dir_safe,
    resolve_dialog_start_path,
)
from common.ui_logging import UILogHandler, format_log_line
from common.workbook_signing import verify_payload_signature

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}

_logger = logging.getLogger(__name__)


@dataclass(slots=True)
class CoordinatorWorkflowState:
    busy: bool = False
    active_job_id: str | None = None

    def set_busy(self, value: bool, *, job_id: str | None = None) -> None:
        self.busy = value
        self.active_job_id = job_id if value else None


@dataclass(slots=True, frozen=True)
class _FinalReportSignature:
    template_id: str
    course_code: str
    total_outcomes: int
    section: str
    direct_sheet_count: int
    indirect_sheet_count: int


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


def _has_valid_final_co_report(path: Path) -> bool:
    return _extract_final_report_signature(path) is not None


def _extract_final_report_signature(path: Path) -> _FinalReportSignature | None:
    # openpyxl does not support legacy .xls files.
    if path.suffix.lower() == ".xls":
        return None
    try:
        from openpyxl import load_workbook
    except Exception:
        _logger.exception("openpyxl is unavailable while validating '%s'.", path)
        return None
    try:
        workbook = load_workbook(filename=path, read_only=True, data_only=True)
        try:
            if SYSTEM_HASH_SHEET not in workbook.sheetnames:
                return None
            if SYSTEM_REPORT_INTEGRITY_SHEET not in workbook.sheetnames:
                return None
            if COURSE_METADATA_SHEET not in workbook.sheetnames:
                return None

            hash_sheet = workbook[SYSTEM_HASH_SHEET]
            if normalize(hash_sheet["A1"].value) != normalize(SYSTEM_HASH_TEMPLATE_ID_HEADER):
                return None
            if normalize(hash_sheet["B1"].value) != normalize(SYSTEM_HASH_TEMPLATE_HASH_HEADER):
                return None

            template_id_raw = hash_sheet["A2"].value
            template_hash_raw = hash_sheet["B2"].value
            template_id = str(template_id_raw).strip() if template_id_raw is not None else ""
            template_hash = str(template_hash_raw).strip() if template_hash_raw is not None else ""
            if not template_id or not template_hash:
                return None
            if not verify_payload_signature(template_id, template_hash):
                return None

            integrity_sheet = workbook[SYSTEM_REPORT_INTEGRITY_SHEET]
            if normalize(integrity_sheet["A1"].value) != normalize(SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER):
                return None
            if normalize(integrity_sheet["B1"].value) != normalize(SYSTEM_REPORT_INTEGRITY_HASH_HEADER):
                return None

            manifest_text_raw = integrity_sheet["A2"].value
            manifest_hash_raw = integrity_sheet["B2"].value
            manifest_text = str(manifest_text_raw).strip() if manifest_text_raw is not None else ""
            manifest_hash = str(manifest_hash_raw).strip() if manifest_hash_raw is not None else ""
            if not manifest_text or not manifest_hash:
                return None
            if not verify_payload_signature(manifest_text, manifest_hash):
                return None

            manifest = json.loads(manifest_text)
            if not isinstance(manifest, dict):
                return None
            sheet_order = manifest.get("sheet_order")
            if not isinstance(sheet_order, list):
                return None
            sheet_names = [str(name).strip() for name in sheet_order if isinstance(name, str)]
            if not sheet_names:
                return None

            direct_sheet_count = sum(
                1 for name in sheet_names if name.endswith(CO_REPORT_DIRECT_SHEET_SUFFIX)
            )
            indirect_sheet_count = sum(
                1 for name in sheet_names if name.endswith(CO_REPORT_INDIRECT_SHEET_SUFFIX)
            )
            if direct_sheet_count <= 0 or indirect_sheet_count <= 0:
                return None
            if direct_sheet_count != indirect_sheet_count:
                return None

            metadata_sheet = workbook[COURSE_METADATA_SHEET]
            metadata: dict[str, str] = {}
            row = 2
            while True:
                key = metadata_sheet.cell(row=row, column=1).value
                value = metadata_sheet.cell(row=row, column=2).value
                if normalize(key) == "" and normalize(value) == "":
                    break
                key_text = str(key).strip() if key is not None else ""
                value_text = str(value).strip() if value is not None else ""
                if key_text:
                    metadata[normalize(key_text)] = value_text
                row += 1

            course_code = metadata.get(normalize(COURSE_METADATA_COURSE_CODE_KEY), "").strip()
            total_outcomes_text = metadata.get(normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY), "").strip()
            section = metadata.get(normalize(COURSE_METADATA_SECTION_KEY), "").strip()
            if not course_code or not total_outcomes_text or not section:
                return None
            try:
                total_outcomes = int(float(total_outcomes_text))
            except (TypeError, ValueError):
                return None
            if total_outcomes <= 0:
                return None

            return _FinalReportSignature(
                template_id=template_id,
                course_code=course_code,
                total_outcomes=total_outcomes,
                section=section,
                direct_sheet_count=direct_sheet_count,
                indirect_sheet_count=indirect_sheet_count,
            )
        finally:
            workbook.close()
    except Exception:
        _logger.exception("Failed to validate final CO report workbook '%s'.", path)
        return None


def _analyze_dropped_files(
    dropped_files: list[str],
    *,
    existing_keys: set[str],
    existing_paths: list[str],
    token: CancellationToken,
) -> dict[str, object]:
    accepted = _filter_excel_paths(dropped_files)
    seen = set(existing_keys)
    added: list[str] = []
    duplicates = 0
    invalid_final_report: list[str] = []
    existing_resolved = [Path(path).resolve() for path in existing_paths if path]
    baseline_signature: _FinalReportSignature | None = None
    seen_sections: set[str] = set()

    for path in existing_resolved:
        token.raise_if_cancelled()
        signature = _extract_final_report_signature(path)
        if signature is None:
            continue
        if baseline_signature is None:
            baseline_signature = signature
        section_key = normalize(signature.section)
        if section_key:
            seen_sections.add(section_key)

    for path in accepted:
        token.raise_if_cancelled()
        key = _path_key(path)
        if key in seen:
            duplicates += 1
            continue
        signature = _extract_final_report_signature(path)
        if signature is None:
            invalid_final_report.append(str(path))
            continue
        if baseline_signature is None:
            baseline_signature = signature
        else:
            is_mismatch = (
                signature.template_id != baseline_signature.template_id
                or signature.course_code != baseline_signature.course_code
                or signature.total_outcomes != baseline_signature.total_outcomes
                or signature.direct_sheet_count != baseline_signature.direct_sheet_count
                or signature.indirect_sheet_count != baseline_signature.indirect_sheet_count
            )
            if is_mismatch:
                invalid_final_report.append(str(path))
                continue
            if normalize(signature.section) in seen_sections:
                invalid_final_report.append(str(path))
                continue
        seen_sections.add(normalize(signature.section))
        seen.add(key)
        added.append(str(path))

    ignored = (len(dropped_files) - len(accepted)) + duplicates + len(invalid_final_report)
    return {
        "added": added,
        "duplicates": duplicates,
        "invalid_final_report": invalid_final_report,
        "ignored": ignored,
    }


class _ExcelDropList(QListWidget):
    files_dropped = Signal(list)
    drag_state_changed = Signal(bool)

    def __init__(self) -> None:
        super().__init__()
        self._placeholder_text = ""
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setDropIndicatorShown(False)
        self.setSpacing(2)
        self.setAlternatingRowColors(False)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)

    def set_placeholder_text(self, text: str) -> None:
        self._placeholder_text = text
        self.viewport().update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if self.count() != 0 or not self._placeholder_text:
            return
        painter = QPainter(self.viewport())
        painter.setPen(QColor("gray"))
        painter.setFont(QFont(UI_FONT_FAMILY, 10))
        painter.drawText(
            self.viewport().rect().adjusted(16, 16, -16, -16),
            Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap,
            self._placeholder_text,
        )
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
        self.drag_state_changed.emit(False)
        if dropped:
            self.files_dropped.emit(dropped)
            event.acceptProposedAction()
            return
        event.ignore()


class _CoordinatorFileItemWidget(QWidget):
    removed = Signal(str)

    def __init__(self, file_path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.file_path = file_path

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 4, 12, 4)
        layout.setSpacing(12)

        file_name = Path(file_path).name
        name_label = QLabel(file_name)
        name_label.setFont(QFont(UI_FONT_FAMILY, 10))
        name_label.setToolTip(file_path)
        layout.addWidget(name_label, 1)

        self.remove_btn = QPushButton()
        self.remove_btn.setFixedSize(24, 24)
        self.remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
        if not icon.isNull():
            self.remove_btn.setIcon(icon)
            self.remove_btn.setIconSize(QSize(16, 16))
        else:
            self.remove_btn.setText("X")
        self.remove_btn.setStyleSheet(
            """
            QPushButton {
                background-color: transparent;
                color: #e74c3c;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(231, 76, 60, 0.15);
                border-radius: 4px;
            }
            """
        )
        self.remove_btn.clicked.connect(lambda: self.removed.emit(self.file_path))
        layout.addWidget(self.remove_btn, 0)


class CoordinatorModule(QWidget):
    status_changed = Signal(str)
    OUTPUT_LINK_OPEN_FILE_KEY = "instructor.links.open_file"
    OUTPUT_LINK_OPEN_FOLDER_KEY = "instructor.links.open_folder"
    OUTPUT_LINK_NOT_AVAILABLE_KEY = "instructor.links.not_available"
    OUTPUT_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"

    def __init__(self) -> None:
        super().__init__()
        self._files: list[Path] = []
        self._downloaded_outputs: list[Path] = []
        self._logger = _logger
        self.state = CoordinatorWorkflowState()
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._pending_drop_batches: list[list[str]] = []
        self._ui_log_handler: UILogHandler | None = None
        self._build_ui()
        self._setup_ui_logging()
        self.retranslate_ui()
        self._refresh_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(
            INSTRUCTOR_CARD_MARGIN,
            INSTRUCTOR_CARD_MARGIN,
            INSTRUCTOR_CARD_MARGIN,
            INSTRUCTOR_CARD_MARGIN,
        )
        root.setSpacing(INSTRUCTOR_CARD_SPACING)

        self.title_label = QLabel()
        self.title_label.setObjectName("coordinatorTitle")
        self.title_label.setFont(QFont(UI_FONT_FAMILY, INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE, QFont.Weight.Bold))
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.title_label)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        root.addWidget(divider)

        self.hint_label = QLabel()
        self.hint_label.setObjectName("coordinatorHint")
        self.hint_label.setFont(QFont(UI_FONT_FAMILY, 10))
        self.hint_label.setWordWrap(True)
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.hint_label)

        self.drop_list = _ExcelDropList()
        self.drop_list.setObjectName("coordinatorDropList")
        self.drop_list.setMinimumHeight(240)
        self.drop_list.files_dropped.connect(self._on_files_dropped)
        self.drop_list.drag_state_changed.connect(self._set_drop_active)

        frame = QFrame()
        frame.setObjectName("coordinatorDropFrame")
        frame_layout = QVBoxLayout(frame)
        frame_layout.setContentsMargins(14, 14, 14, 14)
        frame_layout.setSpacing(0)
        frame_layout.addWidget(self.drop_list)
        root.addWidget(frame, 1)

        button_row = QHBoxLayout()
        button_row.setContentsMargins(0, 0, 0, 0)
        button_row.setSpacing(10)

        self.add_button = QPushButton()
        self.add_button.setObjectName("primaryAction")
        self.add_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_button.setMinimumHeight(34)
        self.add_button.setMinimumWidth(130)
        self.add_button.clicked.connect(self._browse_files)
        button_row.addWidget(self.add_button)

        self.clear_button = QPushButton()
        self.clear_button.setObjectName("coordinatorClearButton")
        self.clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_button.setMinimumHeight(34)
        self.clear_button.setMinimumWidth(120)
        self.clear_button.clicked.connect(self._clear_all)
        button_row.addWidget(self.clear_button)

        button_row.addStretch(1)

        self.calculate_button = QPushButton()
        self.calculate_button.setObjectName("coordinatorCalculateButton")
        self.calculate_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.calculate_button.setMinimumHeight(34)
        self.calculate_button.setMinimumWidth(200)
        self.calculate_button.clicked.connect(self._on_calculate_clicked)
        button_row.addWidget(self.calculate_button)

        root.addLayout(button_row)

        self.summary_label = QLabel()
        self.summary_label.setObjectName("coordinatorSummary")
        self.summary_label.setFont(QFont(UI_FONT_FAMILY, 9))
        root.addWidget(self.summary_label)

        self.info_tabs = QTabWidget()
        self.info_tabs.setObjectName("instructorInfoTabs")
        self.info_tabs.setFixedHeight(INSTRUCTOR_INFO_TAB_FIXED_HEIGHT)
        self.info_tabs.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.info_tabs.tabBar().setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.info_tabs.currentChanged.connect(self._on_info_tab_changed)

        log_tab = QWidget()
        log_tab_layout = QVBoxLayout(log_tab)
        log_tab_layout.setContentsMargins(*INSTRUCTOR_INFO_TAB_LAYOUT_MARGINS)
        log_tab_layout.setSpacing(INSTRUCTOR_INFO_TAB_LAYOUT_SPACING)

        self.user_log_view = QPlainTextEdit()
        self.user_log_view.setReadOnly(True)
        self.user_log_view.setObjectName("userLogView")
        self.user_log_view.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.user_log_view.setFrameShape(QFrame.Shape.NoFrame)
        self.user_log_view.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        log_tab_layout.addWidget(self.user_log_view)

        links_tab = QWidget()
        links_tab_layout = QVBoxLayout(links_tab)
        links_tab_layout.setContentsMargins(*INSTRUCTOR_INFO_TAB_LAYOUT_MARGINS)
        links_tab_layout.setSpacing(INSTRUCTOR_INFO_TAB_LAYOUT_SPACING)

        self.generated_outputs_view = QTextBrowser()
        self.generated_outputs_view.setObjectName("generatedOutputsView")
        self.generated_outputs_view.setOpenExternalLinks(False)
        self.generated_outputs_view.setOpenLinks(False)
        self.generated_outputs_view.setFrameShape(QFrame.Shape.NoFrame)
        self.generated_outputs_view.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.generated_outputs_view.anchorClicked.connect(
            lambda url: self._on_output_link_activated(url.toString())
        )
        links_tab_layout.addWidget(self.generated_outputs_view)

        self.info_tabs.addTab(log_tab, t("instructor.log.title"))
        self.info_tabs.addTab(links_tab, t("instructor.links.title"))
        root.addWidget(self.info_tabs)

        self.shortcut_add_file = QShortcut(QKeySequence("Ctrl+O"), self)
        self.shortcut_add_file.activated.connect(self._browse_files)

        panel_style = """
        QLabel#coordinatorTitle {
        }
        QLabel#coordinatorHint {
            color: palette(mid);
        }
        QLabel#coordinatorSummary {
            color: palette(mid);
            margin-top: 0px;
        }
        QFrame#coordinatorDropFrame {
            border: 1px solid palette(mid);
            border-radius: 10px;
            background-color: palette(base);
        }
        QListWidget#coordinatorDropList {
            border: 2px dashed palette(mid);
            border-radius: 10px;
            padding: 10px;
            background-color: palette(alternate-base);
        }
        QListWidget#coordinatorDropList:hover {
            border-color: palette(highlight);
        }
        QListWidget#coordinatorDropList[dragActive="true"] {
            border-color: palette(highlight);
            background-color: rgba(22, 160, 133, 0.10);
        }
        QListWidget#coordinatorDropList::item {
            margin: 2px 0;
        }
        QPushButton {
            border-radius: 6px;
            padding: 7px 12px;
            border: 1px solid palette(mid);
        }
        QPushButton#primaryAction {
            min-width: 150px;
            min-height: 30px;
            border-radius: 6px;
            border: none;
        }
        QPushButton#primaryAction:enabled {
            background-color: palette(highlight);
            color: palette(highlighted-text);
        }
        QPushButton#coordinatorCalculateButton {
            background-color: palette(highlight);
            color: palette(highlighted-text);
            border: none;
        }
        QPushButton#coordinatorCalculateButton:hover {
            opacity: 0.92;
        }
        QPushButton#coordinatorCalculateButton:disabled {
            background-color: palette(button);
            color: palette(mid);
            border: 1px solid palette(mid);
        }
        QPushButton#coordinatorClearButton:hover {
            border-color: #c0392b;
            background-color: rgba(231, 76, 60, 0.08);
        }
        QPushButton#coordinatorClearButton:disabled {
            color: palette(mid);
        }
        QTabWidget#instructorInfoTabs::pane {
            border: none;
            background: palette(base);
        }
        QTabWidget#instructorInfoTabs QTabBar::tab:first {
            margin-left: 8px;
        }
        QTabWidget#instructorInfoTabs QPlainTextEdit,
        QTabWidget#instructorInfoTabs QTextBrowser {
            border: 1px solid palette(mid);
            border-radius: 8px;
            background: palette(base);
            padding: 8px;
        }
        """

        self.setStyleSheet(panel_style)

    def retranslate_ui(self) -> None:
        self.title_label.setText(t("coordinator.title"))
        self.hint_label.setText(t("coordinator.drop_hint"))
        self.drop_list.set_placeholder_text(t("coordinator.list_placeholder"))
        self.add_button.setText(t("coordinator.add_file"))
        self.clear_button.setText(t("coordinator.clear_all"))
        self.calculate_button.setText(t("coordinator.calculate"))
        self.info_tabs.setTabText(0, t("instructor.log.title"))
        self.info_tabs.setTabText(1, t("instructor.links.title"))
        self._refresh_output_links()
        self._refresh_summary()

    def _publish_status(self, message: str) -> None:
        self._append_user_log(message)
        emit_user_status(self.status_changed, message, logger=self._logger)

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        self._refresh_ui()

    def _refresh_ui(self) -> None:
        has_files = bool(self._files)
        self.add_button.setEnabled(not self.state.busy)
        self.clear_button.setEnabled(has_files and not self.state.busy)
        self.calculate_button.setEnabled(has_files and not self.state.busy)
        for row in range(self.drop_list.count()):
            item = self.drop_list.item(row)
            widget = self.drop_list.itemWidget(item)
            if isinstance(widget, _CoordinatorFileItemWidget):
                widget.remove_btn.setEnabled(not self.state.busy)
        self._refresh_output_links()
        self._refresh_summary()

    def _on_calculate_clicked(self) -> None:
        self._publish_status(t("coordinator.status.calculate_pending"))

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
        existing_paths = [str(path) for path in self._files]
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
                invalid_paths = [Path(value) for value in result.get("invalid_final_report", [])]
                ignored = int(result.get("ignored", 0))

                for path in added_paths:
                    self._files.append(path)
                    item = QListWidgetItem()
                    item.setToolTip(str(path))
                    item.setData(Qt.ItemDataRole.UserRole, str(path))
                    self.drop_list.addItem(item)
                    row_widget = _CoordinatorFileItemWidget(str(path), parent=self.drop_list)
                    row_widget.removed.connect(self._remove_file_by_path)
                    item.setSizeHint(row_widget.sizeHint())
                    self.drop_list.setItemWidget(item, row_widget)

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
                            "coordinator.invalid_final_report.body",
                            count=len(invalid_paths),
                            files=file_names,
                        ),
                        title=t("coordinator.invalid_final_report.title"),
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
            existing_paths=existing_paths,
            token=token,
            on_finished=_on_finished,
            on_failed=_on_failed,
        )
        self._active_jobs.append(job)

    def _on_files_dropped(self, dropped_files: list[str]) -> None:
        first_path = next((value for value in dropped_files if value), "")
        if first_path:
            self._remember_dialog_dir_safe(first_path)
        self._process_files_async(dropped_files)

    def _browse_files(self) -> None:
        if self.state.busy:
            return
        selected_files, _ = QFileDialog.getOpenFileNames(
            self,
            t("coordinator.dialog.title"),
            resolve_dialog_start_path(APP_NAME),
            t("coordinator.dialog.filter"),
        )
        if selected_files:
            self._remember_dialog_dir_safe(selected_files[0])
            self._process_files_async(selected_files)

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        try:
            remember_dialog_dir(selected_path, app_name=APP_NAME)
        except OSError:
            remember_dialog_dir_safe(
                selected_path,
                app_name=APP_NAME,
                logger=self._logger,
            )

    def _setup_ui_logging(self) -> None:
        if self._ui_log_handler is not None:
            return
        self._ui_log_handler = UILogHandler(self._append_user_log)
        self._logger.addHandler(self._ui_log_handler)
        self._append_user_log(t("instructor.log.ready"))

    def _append_user_log(self, message: str) -> None:
        line = format_log_line(message)
        if line is None:
            return
        self.user_log_view.appendPlainText(line)

    def _output_link_markup(self, label: str, path: str | None) -> str:
        if not path:
            return f"<b>{escape(label)}</b>: {t(self.OUTPUT_LINK_NOT_AVAILABLE_KEY)}"
        file_link = f'<a href="file::{path}">{t(self.OUTPUT_LINK_OPEN_FILE_KEY)}</a>'
        folder_link = f'<a href="folder::{path}">{t(self.OUTPUT_LINK_OPEN_FOLDER_KEY)}</a>'
        name = escape(Path(path).name)
        full_path = escape(str(Path(path)))
        return (
            f"<b>{escape(label)}</b>: {name}<br>"
            f"<span>{full_path}</span><br>"
            f"{file_link} | {folder_link}"
        )

    def _output_links_html(self) -> str:
        rows: list[str] = []
        for path in self._files:
            rows.append(
                f"<div style='margin-bottom:10px'>{self._output_link_markup(t('coordinator.links.uploaded_report'), str(path))}</div>"
            )
        if not rows:
            rows.append(
                f"<div style='margin-bottom:10px'>{self._output_link_markup(t('coordinator.links.uploaded_report'), None)}</div>"
            )

        if self._downloaded_outputs:
            for path in self._downloaded_outputs:
                rows.append(
                    f"<div style='margin-bottom:10px'>{self._output_link_markup(t('coordinator.links.downloaded_output'), str(path))}</div>"
                )
        else:
            rows.append(
                f"<div style='margin-bottom:10px'>{self._output_link_markup(t('coordinator.links.downloaded_output'), None)}</div>"
            )
        return "".join(rows)

    def _refresh_output_links(self) -> None:
        self.generated_outputs_view.setHtml(self._output_links_html())

    def _on_output_link_activated(self, href: str) -> None:
        mode, _, raw_path = href.partition("::")
        path = raw_path.strip()
        if not path:
            return
        target = Path(path).parent if mode == "folder" else Path(path)
        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
        if opened:
            return
        show_toast(
            self,
            t(self.OUTPUT_LINK_OPEN_FAILED_KEY),
            title=t("instructor.msg.error_title"),
            level="error",
        )

    def _clear_info_text_selection(self) -> None:
        for view in (self.user_log_view, self.generated_outputs_view):
            cursor = view.textCursor()
            if cursor.hasSelection():
                cursor.clearSelection()
                view.setTextCursor(cursor)

    def _on_info_tab_changed(self, _index: int) -> None:
        self._clear_info_text_selection()

    def _refresh_summary(self) -> None:
        self.summary_label.setText(t("coordinator.summary", count=len(self._files)))

    def _set_drop_active(self, active: bool) -> None:
        self.drop_list.setProperty("dragActive", active)
        self.drop_list.style().unpolish(self.drop_list)
        self.drop_list.style().polish(self.drop_list)
        self.drop_list.viewport().update()

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self.info_tabs.setVisible(not enabled)

    def get_shared_outputs_html(self) -> str:
        return self._output_links_html()

    def _remove_file_by_path(self, file_path: str) -> None:
        if self.state.busy:
            return
        target_key = _path_key(Path(file_path))
        before_count = len(self._files)
        self._files = [path for path in self._files if _path_key(path) != target_key]
        if len(self._files) == before_count:
            return

        for row in range(self.drop_list.count()):
            item = self.drop_list.item(row)
            path_value = str(item.data(Qt.ItemDataRole.UserRole) or "")
            if _path_key(Path(path_value)) == target_key:
                self.drop_list.takeItem(row)
                break

        self._refresh_ui()
        self._publish_status(t("coordinator.status.removed", count=1))
        log_process_message(
            "removing selected coordinator files",
            logger=self._logger,
            success_message="removing selected coordinator files completed successfully. removed=1",
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

    def closeEvent(self, event) -> None:
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._cancel_token = None
        self._active_jobs.clear()
        if self._ui_log_handler is not None:
            self._logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)
