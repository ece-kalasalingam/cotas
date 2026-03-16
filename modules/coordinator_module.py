"""Course coordinator module for collecting Final CO report Excel files."""

from __future__ import annotations

import logging
from datetime import datetime
from dataclasses import dataclass
from html import escape
from pathlib import Path
from typing import Any

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
    QPalette,
    QPainter,
    QPen,
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
    QSizePolicy,
    QStyle,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from common.constants import (
    APP_NAME,
    COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT,
    COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES,
    COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
    COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
    COORDINATOR_CONTROLS_LAYOUT_MARGINS,
    COORDINATOR_CONTROLS_LAYOUT_SPACING,
    COORDINATOR_DROPZONE_BG_ACTIVE_ALPHA,
    COORDINATOR_DROPZONE_BORDER_DASH_PATTERN,
    COORDINATOR_DROPZONE_BORDER_WIDTH,
    COORDINATOR_DROPZONE_INNER_RADIUS,
    COORDINATOR_DROPZONE_INNER_RECT_ADJUST,
    COORDINATOR_DROPZONE_OUTER_RADIUS,
    COORDINATOR_DROPZONE_OUTER_RECT_ADJUST,
    COORDINATOR_DROP_LIST_ITEM_SPACING,
    COORDINATOR_DROP_LIST_MIN_HEIGHT,
    COORDINATOR_DROP_ZONE_LAYOUT_MARGINS,
    COORDINATOR_FILE_ITEM_LAYOUT_MARGINS,
    COORDINATOR_FILE_ITEM_LAYOUT_SPACING,
    COORDINATOR_FILE_NAME_FONT_SIZE,
    COORDINATOR_HEADER_LAYOUT_MARGINS,
    COORDINATOR_HEADER_LAYOUT_SPACING,
    COORDINATOR_LIST_PLACEHOLDER_FONT_SIZE,
    COORDINATOR_LIST_PLACEHOLDER_COLOR,
    COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
    COORDINATOR_PANEL_STYLESHEET,
    COORDINATOR_REMOVE_BUTTON_ICON_SIZE,
    COORDINATOR_REMOVE_BUTTON_SIZE,
    COORDINATOR_REMOVE_BUTTON_STYLESHEET,
    COORDINATOR_ROOT_MIN_SPACING,
    COORDINATOR_SUMMARY_FONT_SIZE,
    INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE,
    INSTRUCTOR_CARD_MARGIN,
    INSTRUCTOR_CARD_SPACING,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    INSTRUCTOR_INFO_TAB_LAYOUT_MARGINS,
    INSTRUCTOR_INFO_TAB_LAYOUT_SPACING,
    OUTPUT_LINK_MODE_FILE,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX,
    OUTPUT_LINK_SEPARATOR,
    SHORTCUT_OPEN_KEY_SEQUENCE,
    SHORTCUT_SAVE_KEY_SEQUENCE,
    UI_FONT_FAMILY,
)
from common.exceptions import JobCancelledError
from common.jobs import CancellationToken, generate_job_id
from common.qt_jobs import run_in_background
from common.texts import t
from common.toast import show_toast
from common.ui_logging import (
    UILogHandler,
    build_i18n_log_message,
    format_log_line_at,
    parse_i18n_log_message,
    resolve_i18n_log_message,
)
from common.utils import (
    emit_user_status,
    log_process_message,
    remember_dialog_dir,
    remember_dialog_dir_safe,
    resolve_dialog_start_path,
)

from .coordinator_processing import (
    _CoAttainmentWorkbookResult,
    _analyze_dropped_files,
    _build_co_attainment_default_name,
    _extract_final_report_signature,
    _generate_co_attainment_workbook,
    _has_valid_final_co_report as _processing_has_valid_final_co_report,
    _path_key,
)

_logger = logging.getLogger(__name__)


@dataclass(slots=True)
class CoordinatorWorkflowState:
    busy: bool = False
    active_job_id: str | None = None

    def set_busy(self, value: bool, *, job_id: str | None = None) -> None:
        self.busy = value
        self.active_job_id = job_id if value else None


def _has_valid_final_co_report(path: Path) -> bool:
    return _processing_has_valid_final_co_report(path)


class _ExcelDropList(QListWidget):
    files_dropped = Signal(list)
    drag_state_changed = Signal(bool)
    browse_requested = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._placeholder_text = ""
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setDropIndicatorShown(False)
        self.setSpacing(COORDINATOR_DROP_LIST_ITEM_SPACING)
        self.setAlternatingRowColors(False)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

    def set_placeholder_text(self, text: str) -> None:
        self._placeholder_text = text
        self.viewport().update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if self.count() != 0 or not self._placeholder_text:
            return
        painter = QPainter(self.viewport())
        painter.setPen(QColor(COORDINATOR_LIST_PLACEHOLDER_COLOR))
        painter.setFont(QFont(UI_FONT_FAMILY, COORDINATOR_LIST_PLACEHOLDER_FONT_SIZE))
        painter.drawText(
            self.viewport().rect().adjusted(*COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS),
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

    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.browse_requested.emit()
            event.accept()
            return
        super().mouseDoubleClickEvent(event)


class _DropZoneFrame(QFrame):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("coordinatorDropZone")

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        palette = self.palette()
        active = bool(self.property("dragActive"))
        bg_color = palette.color(QPalette.ColorRole.AlternateBase)
        if active:
            bg_color.setAlpha(COORDINATOR_DROPZONE_BG_ACTIVE_ALPHA)
        border_color = palette.color(QPalette.ColorRole.Highlight) if active else palette.color(QPalette.ColorRole.Mid)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(bg_color)
        painter.drawRoundedRect(
            self.rect().adjusted(*COORDINATOR_DROPZONE_OUTER_RECT_ADJUST),
            COORDINATOR_DROPZONE_OUTER_RADIUS,
            COORDINATOR_DROPZONE_OUTER_RADIUS,
        )
        pen = QPen(border_color, COORDINATOR_DROPZONE_BORDER_WIDTH, Qt.PenStyle.DashLine)
        pen.setDashPattern(list(COORDINATOR_DROPZONE_BORDER_DASH_PATTERN))
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(
            self.rect().adjusted(*COORDINATOR_DROPZONE_INNER_RECT_ADJUST),
            COORDINATOR_DROPZONE_INNER_RADIUS,
            COORDINATOR_DROPZONE_INNER_RADIUS,
        )
        painter.end()


class _ElidedFileNameLabel(QLabel):
    def __init__(self, text: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._full_text = text
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setMinimumWidth(0)
        self._apply_elided_text()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._apply_elided_text()

    def _apply_elided_text(self) -> None:
        width = self.contentsRect().width()
        if width <= 0:
            return
        super().setText(self.fontMetrics().elidedText(self._full_text, Qt.TextElideMode.ElideMiddle, width))


class _CoordinatorFileItemWidget(QWidget):
    removed = Signal(str)

    def __init__(self, file_path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.file_path = file_path

        layout = QHBoxLayout(self)
        layout.setContentsMargins(*COORDINATOR_FILE_ITEM_LAYOUT_MARGINS)
        layout.setSpacing(COORDINATOR_FILE_ITEM_LAYOUT_SPACING)

        file_name = Path(file_path).name
        name_label = _ElidedFileNameLabel(file_name)
        name_label.setFont(QFont(UI_FONT_FAMILY, COORDINATOR_FILE_NAME_FONT_SIZE))
        name_label.setToolTip(file_path)
        layout.addWidget(name_label, 1)

        self.remove_btn = QPushButton()
        self.remove_btn.setObjectName("coordinatorFileRemoveButton")
        self.remove_btn.setFixedSize(COORDINATOR_REMOVE_BUTTON_SIZE, COORDINATOR_REMOVE_BUTTON_SIZE)
        self.remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
        if not icon.isNull():
            self.remove_btn.setIcon(icon)
            self.remove_btn.setIconSize(QSize(COORDINATOR_REMOVE_BUTTON_ICON_SIZE, COORDINATOR_REMOVE_BUTTON_ICON_SIZE))
        else:
            self.remove_btn.setText(t("coordinator.file.remove_fallback"))
        self.remove_btn.setStyleSheet(COORDINATOR_REMOVE_BUTTON_STYLESHEET)
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
        self._user_log_entries: list[dict[str, object]] = []
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
        root.setSpacing(max(COORDINATOR_ROOT_MIN_SPACING, INSTRUCTOR_CARD_SPACING))

        header_card = QFrame()
        header_card.setObjectName("coordinatorHeaderCard")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(*COORDINATOR_HEADER_LAYOUT_MARGINS)
        header_layout.setSpacing(COORDINATOR_HEADER_LAYOUT_SPACING)
        self.title_label = QLabel()
        self.title_label.setObjectName("coordinatorTitle")
        self.title_label.setFont(QFont(UI_FONT_FAMILY, INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE, QFont.Weight.Bold))
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        header_layout.addWidget(self.title_label)
        self.hint_label = QLabel()
        self.hint_label.setObjectName("coordinatorHint")
        self.hint_label.setFont(QFont(UI_FONT_FAMILY, COORDINATOR_LIST_PLACEHOLDER_FONT_SIZE))
        self.hint_label.setWordWrap(True)
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        header_layout.addWidget(self.hint_label)
        root.addWidget(header_card)

        self.drop_zone = _DropZoneFrame()
        self.drop_zone.setProperty("dragActive", False)
        zone_layout = QVBoxLayout(self.drop_zone)
        zone_layout.setContentsMargins(*COORDINATOR_DROP_ZONE_LAYOUT_MARGINS)
        zone_layout.setSpacing(0)
        self.drop_list = _ExcelDropList()
        self.drop_list.setObjectName("coordinatorDropList")
        self.drop_list.setMinimumHeight(COORDINATOR_DROP_LIST_MIN_HEIGHT)
        self.drop_list.files_dropped.connect(self._on_files_dropped)
        self.drop_list.drag_state_changed.connect(self._set_drop_active)
        self.drop_list.browse_requested.connect(self._browse_files)
        zone_layout.addWidget(self.drop_list)
        root.addWidget(self.drop_zone, 1)

        controls_row = QHBoxLayout()
        controls_row.setContentsMargins(*COORDINATOR_CONTROLS_LAYOUT_MARGINS)
        controls_row.setSpacing(COORDINATOR_CONTROLS_LAYOUT_SPACING)
        self.summary_label = QLabel()
        self.summary_label.setObjectName("coordinatorSummary")
        self.summary_label.setFont(QFont(UI_FONT_FAMILY, COORDINATOR_SUMMARY_FONT_SIZE))
        controls_row.addWidget(self.summary_label)
        controls_row.addStretch(1)
        self.clear_button = QPushButton()
        self.clear_button.setObjectName("coordinatorClearButton")
        self.clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_button.setMinimumWidth(150)
        self.clear_button.setDefault(True)
        self.clear_button.clicked.connect(self._clear_all)
        controls_row.addWidget(self.clear_button)
        self.calculate_button = QPushButton()
        self.calculate_button.setObjectName("coordinatorCalculateButton")
        self.calculate_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.calculate_button.setMinimumWidth(150)
        self.calculate_button.setAutoDefault(False)
        self.calculate_button.setDefault(False)
        self.calculate_button.clicked.connect(self._on_calculate_clicked)
        controls_row.addWidget(self.calculate_button)
        root.addLayout(controls_row)

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

        self.shortcut_add_file = QShortcut(QKeySequence(SHORTCUT_OPEN_KEY_SEQUENCE), self)
        self.shortcut_add_file.activated.connect(self._browse_files)
        self.shortcut_save_output = QShortcut(QKeySequence(SHORTCUT_SAVE_KEY_SEQUENCE), self)
        self.shortcut_save_output.activated.connect(self._on_save_shortcut_activated)

        self.setStyleSheet(COORDINATOR_PANEL_STYLESHEET)

    def retranslate_ui(self) -> None:
        self._rerender_user_log()
        self.title_label.setText(t("coordinator.title"))
        self.hint_label.setText(t("coordinator.drop_hint"))
        self.drop_list.set_placeholder_text(t("coordinator.list_placeholder"))
        self.clear_button.setText(t("coordinator.clear_all"))
        self.calculate_button.setText(t("coordinator.calculate"))
        self.info_tabs.setTabText(0, t("instructor.log.title"))
        self.info_tabs.setTabText(1, t("instructor.links.title"))
        self._refresh_output_links()
        self._refresh_summary()

    def _publish_status(self, message: str) -> None:
        self._append_user_log(message)
        emit_user_status(self.status_changed, message, logger=self._logger)

    def _publish_status_key(self, text_key: str, **kwargs: Any) -> None:
        localized = t(text_key, **kwargs)
        payload = build_i18n_log_message(text_key, kwargs=kwargs, fallback=localized)
        self._append_user_log(payload)
        emit_user_status(self.status_changed, payload, logger=self._logger)

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        self._refresh_ui()

    def _refresh_ui(self) -> None:
        has_files = bool(self._files)
        self.clear_button.setEnabled(has_files and not self.state.busy)
        self.calculate_button.setEnabled(has_files and not self.state.busy)
        self.drop_list.setEnabled(not self.state.busy)
        for row in range(self.drop_list.count()):
            item = self.drop_list.item(row)
            widget = self.drop_list.itemWidget(item)
            if isinstance(widget, _CoordinatorFileItemWidget):
                widget.remove_btn.setEnabled(not self.state.busy)
        self._refresh_output_links()
        self._refresh_summary()

    def _on_calculate_clicked(self) -> None:
        if self.state.busy or not self._files:
            return

        signature = _extract_final_report_signature(self._files[0])
        default_name = _build_co_attainment_default_name(
            self._files[0],
            section=signature.section if signature is not None else "",
        )
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            t("coordinator.calculate"),
            resolve_dialog_start_path(APP_NAME, default_name),
            t("coordinator.dialog.filter"),
        )
        if not save_path:
            return

        process_name = COORDINATOR_WORKFLOW_OPERATION_CALCULATE_ATTAINMENT
        token = CancellationToken()
        job_id = generate_job_id()
        self._cancel_token = token
        self._set_busy(True, job_id=job_id)
        self._publish_status_key("coordinator.status.processing_started")

        def _finalize(job: object) -> None:
            if job in self._active_jobs:
                self._active_jobs.remove(job)
            self._cancel_token = None
            self._set_busy(False)
            self._drain_next_batch()

        def _on_finished(result: object) -> None:
            try:
                output_path = Path(save_path)
                duplicate_reg_count = 0
                duplicate_entries: tuple[tuple[str, str, str], ...] = ()
                if isinstance(result, _CoAttainmentWorkbookResult):
                    output_path = result.output_path
                    duplicate_reg_count = max(0, int(result.duplicate_reg_count))
                    duplicate_entries = result.duplicate_entries
                elif result:
                    output_path = Path(str(result))
                if all(_path_key(path) != _path_key(output_path) for path in self._downloaded_outputs):
                    self._downloaded_outputs.append(output_path)
                self._remember_dialog_dir_safe(str(output_path))
                self._publish_status_key("coordinator.status.calculate_completed")
                log_process_message(
                    process_name,
                    logger=self._logger,
                    success_message=(
                        f"{process_name} completed successfully. output={output_path}, "
                        f"duplicates_removed={duplicate_reg_count}"
                    ),
                    user_success_message=build_i18n_log_message(
                        "coordinator.status.calculate_completed",
                        fallback=t("coordinator.status.calculate_completed"),
                    ),
                    job_id=job_id,
                    step_id=COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
                )
                show_toast(
                    self,
                    t("coordinator.status.calculate_completed"),
                    title=t("coordinator.title"),
                    level="info",
                )
                if duplicate_reg_count:
                    show_toast(
                        self,
                        t("coordinator.regno_dedup.body", count=duplicate_reg_count),
                        title=t("coordinator.regno_dedup.title"),
                        level="info",
                    )
                    detail_lines = [
                        t(
                            "coordinator.regno_dedup.log_detail",
                            reg_no=str(reg_no),
                            worksheet=str(worksheet_name),
                            workbook=str(workbook_name),
                        )
                        for reg_no, worksheet_name, workbook_name in duplicate_entries
                    ]
                    details_text = "\n".join(detail_lines) if detail_lines else t(
                        "coordinator.regno_dedup.log_detail_unavailable"
                    )
                    self._publish_status_key(
                        "coordinator.regno_dedup.log_body",
                        count=duplicate_reg_count,
                        details=details_text,
                    )
            finally:
                _finalize(job)

        def _on_failed(exc: Exception) -> None:
            try:
                if isinstance(exc, JobCancelledError):
                    self._publish_status_key("coordinator.status.operation_cancelled")
                    self._logger.info(
                        "%s cancelled by user/system request.",
                        process_name,
                        extra={
                            "user_message": build_i18n_log_message(
                                "coordinator.status.operation_cancelled",
                                fallback=t("coordinator.status.operation_cancelled"),
                            ),
                            "job_id": job_id,
                            "step_id": COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
                        },
                    )
                    return
                log_process_message(
                    process_name,
                    logger=self._logger,
                    error=exc,
                    user_error_message=build_i18n_log_message(
                        "coordinator.status.processing_failed",
                        fallback=t("coordinator.status.processing_failed"),
                    ),
                    job_id=job_id,
                    step_id=COORDINATOR_WORKFLOW_STEP_ID_CALCULATE_ATTAINMENT,
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
            _generate_co_attainment_workbook,
            list(self._files),
            Path(save_path),
            token=token,
            on_finished=_on_finished,
            on_failed=_on_failed,
        )
        self._active_jobs.append(job)

    def _on_save_shortcut_activated(self) -> None:
        if self.state.busy:
            return
        if self.calculate_button.isEnabled():
            self._on_calculate_clicked()

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
            self._publish_status_key("coordinator.status.queued", count=len(dropped_files))
            return

        process_name = COORDINATOR_WORKFLOW_OPERATION_COLLECT_FILES
        token = CancellationToken()
        job_id = generate_job_id()
        existing_keys = {_path_key(path) for path in self._files}
        existing_paths = [str(path) for path in self._files]
        self._cancel_token = token
        self._set_busy(True, job_id=job_id)
        self._publish_status_key("coordinator.status.processing_started")

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
                    self._publish_status_key(
                        "coordinator.status.added",
                        added=len(added_paths),
                        total=len(self._files),
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
                    self._publish_status_key("coordinator.status.ignored", count=ignored)

                log_process_message(
                    process_name,
                    logger=self._logger,
                    success_message=(
                        f"{process_name} completed successfully. "
                        f"added={len(added_paths)}, duplicates={duplicates}, "
                        f"invalid={len(invalid_paths)}, ignored={ignored}"
                    ),
                    user_success_message=build_i18n_log_message(
                        "coordinator.status.processing_completed",
                        fallback=t("coordinator.status.processing_completed"),
                    ),
                    job_id=job_id,
                    step_id=COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
                )
            finally:
                _finalize(job)

        def _on_failed(exc: Exception) -> None:
            try:
                if isinstance(exc, JobCancelledError):
                    self._publish_status_key("coordinator.status.operation_cancelled")
                    self._logger.info(
                        "%s cancelled by user/system request.",
                        process_name,
                        extra={
                            "user_message": build_i18n_log_message(
                                "coordinator.status.operation_cancelled",
                                fallback=t("coordinator.status.operation_cancelled"),
                            ),
                            "job_id": job_id,
                            "step_id": COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
                        },
                    )
                    return
                log_process_message(
                    process_name,
                    logger=self._logger,
                    error=exc,
                    user_error_message=build_i18n_log_message(
                        "coordinator.status.processing_failed",
                        fallback=t("coordinator.status.processing_failed"),
                    ),
                    job_id=job_id,
                    step_id=COORDINATOR_WORKFLOW_STEP_ID_COLLECT_FILES,
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
        self._append_user_log(
            build_i18n_log_message(
                "instructor.log.ready",
                fallback=t("instructor.log.ready"),
            )
        )

    def _append_user_log(self, message: str) -> None:
        parsed = parse_i18n_log_message(message)
        localized = resolve_i18n_log_message(message)
        timestamp = datetime.now()
        if parsed is None:
            self._user_log_entries.append({"timestamp": timestamp, "message": localized})
        else:
            key, kwargs, fallback = parsed
            self._user_log_entries.append(
                {
                    "timestamp": timestamp,
                    "message": localized,
                    "text_key": key,
                    "kwargs": kwargs,
                    "fallback": fallback,
                }
            )
        line = format_log_line_at(localized, timestamp=timestamp)
        if line is None:
            return
        self.user_log_view.appendPlainText(line)

    def _rerender_user_log(self) -> None:
        self.user_log_view.clear()
        for entry in self._user_log_entries:
            timestamp = entry.get("timestamp")
            text_key = entry.get("text_key")
            fallback = entry.get("fallback")
            kwargs = entry.get("kwargs")
            message = entry.get("message")
            if isinstance(text_key, str):
                safe_kwargs = kwargs if isinstance(kwargs, dict) else {}
                try:
                    resolved = t(text_key, **safe_kwargs)
                except Exception:
                    resolved = fallback if isinstance(fallback, str) else str(message or "")
            else:
                resolved = str(message or "")
            ts = timestamp if isinstance(timestamp, datetime) else None
            line = format_log_line_at(resolved, timestamp=ts)
            if line is None:
                continue
            self.user_log_view.appendPlainText(line)

    def _output_link_markup(self, label: str, path: str | None) -> str:
        if not path:
            return f"<b>{escape(label)}</b>: {t(self.OUTPUT_LINK_NOT_AVAILABLE_KEY)}"
        href_path = Path(path).as_posix()
        file_link = (
            f'<a href="{OUTPUT_LINK_MODE_FILE}{OUTPUT_LINK_SEPARATOR}{href_path}">'
            f"{t(self.OUTPUT_LINK_OPEN_FILE_KEY)}</a>"
        )
        folder_link = (
            f'<a href="{OUTPUT_LINK_MODE_FOLDER}{OUTPUT_LINK_SEPARATOR}{href_path}">'
            f"{t(self.OUTPUT_LINK_OPEN_FOLDER_KEY)}</a>"
        )
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
                f"<div style='margin-bottom:{OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX}px'>{self._output_link_markup(t('coordinator.links.uploaded_report'), str(path))}</div>"
            )
        if not rows:
            rows.append(
                f"<div style='margin-bottom:{OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX}px'>{self._output_link_markup(t('coordinator.links.uploaded_report'), None)}</div>"
            )

        if self._downloaded_outputs:
            for path in self._downloaded_outputs:
                rows.append(
                    f"<div style='margin-bottom:{OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX}px'>{self._output_link_markup(t('coordinator.links.downloaded_output'), str(path))}</div>"
                )
        else:
            rows.append(
                f"<div style='margin-bottom:{OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX}px'>{self._output_link_markup(t('coordinator.links.downloaded_output'), None)}</div>"
            )
        return "".join(rows)

    def _refresh_output_links(self) -> None:
        self.generated_outputs_view.setHtml(self._output_links_html())

    def _on_output_link_activated(self, href: str) -> None:
        mode, _, raw_path = href.partition(OUTPUT_LINK_SEPARATOR)
        path = raw_path.strip()
        if not path:
            return
        target = Path(path).parent if mode == OUTPUT_LINK_MODE_FOLDER else Path(path)
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
        self.drop_zone.setProperty("dragActive", active)
        self.drop_zone.update()

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
        self._publish_status_key("coordinator.status.removed", count=1)
        log_process_message(
            "removing selected coordinator files",
            logger=self._logger,
            success_message="removing selected coordinator files completed successfully. removed=1",
            user_success_message=build_i18n_log_message(
                "coordinator.status.removed",
                kwargs={"count": 1},
                fallback=t("coordinator.status.removed", count=1),
            ),
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
        self._publish_status_key("coordinator.status.cleared", count=total)
        log_process_message(
            "clearing coordinator files",
            logger=self._logger,
            success_message=f"clearing coordinator files completed successfully. removed={total}",
            user_success_message=build_i18n_log_message(
                "coordinator.status.cleared",
                kwargs={"count": total},
                fallback=t("coordinator.status.cleared", count=total),
            ),
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

