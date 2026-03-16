"""Course coordinator module for collecting Final CO report Excel files."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import (
    QColor,
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
    COORDINATOR_DROPZONE_INNER_RADIUS,
    COORDINATOR_DROPZONE_OUTER_RADIUS,
    COORDINATOR_LIST_PLACEHOLDER_FONT_SIZE,
    COORDINATOR_REMOVE_BUTTON_ICON_SIZE,
    COORDINATOR_REMOVE_BUTTON_SIZE,
    INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE,
    INSTRUCTOR_CARD_MARGIN,
    INSTRUCTOR_CARD_SPACING,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    INSTRUCTOR_INFO_TAB_LAYOUT_MARGINS,
    INSTRUCTOR_INFO_TAB_LAYOUT_SPACING,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
    SHORTCUT_OPEN_KEY_SEQUENCE,
    SHORTCUT_SAVE_KEY_SEQUENCE,
    UI_FONT_FAMILY,
)
from common.jobs import CancellationToken
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
from modules.coordinator.async_runner import AsyncOperationRunner
from modules.coordinator.file_actions import clear_all, remove_file_by_path
from modules.coordinator.messages import (
    append_user_log as _append_user_log_impl,
    publish_status as _publish_status_impl,
    publish_status_key as _publish_status_key_impl,
    rerender_user_log as _rerender_user_log_impl,
    setup_ui_logging as _setup_ui_logging_impl,
)
from modules.coordinator.output_links import (
    on_output_link_activated as _on_output_link_activated_impl,
    output_link_markup as _output_link_markup_impl,
    output_links_html as _output_links_html_impl,
    refresh_output_links as _refresh_output_links_impl,
)
from modules.coordinator.steps.calculate_attainment import calculate_attainment_async
from modules.coordinator.steps.collect_files import (
    add_uploaded_paths as _add_uploaded_paths_impl,
    process_files_async,
)
from services import CoordinatorWorkflowService

from .coordinator_processing import (
    _CoAttainmentWorkbookResult,
    _analyze_dropped_files,
    _build_co_attainment_default_name,
    _extract_final_report_signature,
    _generate_co_attainment_workbook,
    _has_valid_final_co_report as _processing_has_valid_final_co_report,
    _path_key,
)

# Referenced indirectly by coordinator helper modules via ns=globals().
_COORDINATOR_NS_EXPORTS = (
    QListWidgetItem,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
    show_toast,
    build_i18n_log_message,
    format_log_line_at,
    parse_i18n_log_message,
    resolve_i18n_log_message,
    emit_user_status,
    log_process_message,
    _CoAttainmentWorkbookResult,
    _analyze_dropped_files,
    _build_co_attainment_default_name,
    _extract_final_report_signature,
    _generate_co_attainment_workbook,
    _path_key,
)


COORDINATOR_LIST_PLACEHOLDER_COLOR = "gray"
COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS = (16, 16, -16, -16)
COORDINATOR_FILE_NAME_FONT_SIZE = 10
COORDINATOR_SUMMARY_FONT_SIZE = 9
COORDINATOR_ROOT_MIN_SPACING = 10
COORDINATOR_DROP_LIST_ITEM_SPACING = 2
COORDINATOR_FILE_ITEM_LAYOUT_MARGINS = (12, 4, 12, 4)
COORDINATOR_FILE_ITEM_LAYOUT_SPACING = 12
COORDINATOR_HEADER_LAYOUT_MARGINS = (16, 14, 16, 14)
COORDINATOR_HEADER_LAYOUT_SPACING = 6
COORDINATOR_DROP_ZONE_LAYOUT_MARGINS = (14, 14, 14, 14)
COORDINATOR_DROP_LIST_MIN_HEIGHT = 220
COORDINATOR_CONTROLS_LAYOUT_MARGINS = (6, 0, 6, 0)
COORDINATOR_CONTROLS_LAYOUT_SPACING = 10
COORDINATOR_DROPZONE_BG_ACTIVE_ALPHA = 220
COORDINATOR_DROPZONE_OUTER_RECT_ADJUST = (1, 1, -2, -2)
COORDINATOR_DROPZONE_BORDER_WIDTH = 2
COORDINATOR_DROPZONE_BORDER_DASH_PATTERN = (4, 3)
COORDINATOR_DROPZONE_INNER_RECT_ADJUST = (6, 6, -6, -6)
COORDINATOR_REMOVE_BUTTON_STYLESHEET = """
QPushButton {
    background-color: transparent;
    border: none;
    padding: 0px;
    margin: 0px;
    min-width: 24px;
    min-height: 24px;
    max-width: 24px;
    max-height: 24px;
}
QPushButton:hover {
    background-color: rgba(231, 76, 60, 0.15);
    border-radius: 4px;
}
"""
OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX = 10
OUTPUT_LINK_MODE_FILE = "file"
COORDINATOR_PANEL_STYLESHEET = """
QFrame#coordinatorHeaderCard { border: 1px solid palette(mid); border-radius: 12px; background-color: palette(base); }
QLabel#coordinatorTitle { letter-spacing: 0.3px; }
QLabel#coordinatorSummary { padding: 5px 10px; border: 1px solid palette(mid); border-radius: 10px; background-color: palette(alternate-base); }
QFrame#coordinatorDropZone { border: none; background: transparent; }
QListWidget#coordinatorDropList { border: none; padding: 10px; background: transparent; }
QListWidget#coordinatorDropList::item { margin: 2px 0; }
QPushButton#coordinatorClearButton, QPushButton#coordinatorCalculateButton { padding: 6px 12px; min-width: 150px; min-height: 30px; border-radius: 6px; border: none; }
QPushButton#coordinatorClearButton:disabled, QPushButton#coordinatorCalculateButton:disabled { border: 1px solid palette(mid); }
QPushButton#coordinatorCalculateButton:enabled { background-color: palette(highlight); color: palette(highlighted-text); border: none; font-weight: 600; }
QPushButton#coordinatorCalculateButton:disabled { border: 1px solid palette(mid); }
QTabWidget#instructorInfoTabs::pane { border: none; background: palette(base); }
QTabWidget#instructorInfoTabs QTabBar::tab:first { margin-left: 8px; }
QTabWidget#instructorInfoTabs QPlainTextEdit, QTabWidget#instructorInfoTabs QTextBrowser { border: 1px solid palette(mid); border-radius: 8px; background: palette(base); padding: 8px; }
"""

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
        self._workflow_service = CoordinatorWorkflowService()
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._pending_drop_batches: list[list[str]] = []
        self._ui_log_handler: UILogHandler | None = None
        self._user_log_entries: list[dict[str, object]] = []
        self._async_runner = AsyncOperationRunner(self, run_async=run_in_background)
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
        _publish_status_impl(self, message, ns=globals())

    def _publish_status_key(self, text_key: str, **kwargs: Any) -> None:
        _publish_status_key_impl(self, text_key, ns=globals(), **kwargs)

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
        calculate_attainment_async(self, ns=globals())

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
        process_files_async(self, dropped_files, ns=globals())

    def _start_async_operation(
        self,
        *,
        token: CancellationToken,
        job_id: str | None,
        work,
        on_success,
        on_failure,
        on_finally=None,
    ) -> None:
        self._async_runner.start(
            token=token,
            job_id=job_id,
            work=work,
            on_success=on_success,
            on_failure=on_failure,
            on_finally=on_finally,
        )

    def _new_file_item_widget(self, file_path: str, *, parent: QWidget | None = None) -> _CoordinatorFileItemWidget:
        return _CoordinatorFileItemWidget(file_path, parent=parent)

    def _add_uploaded_paths(self, added_paths: list[Path]) -> None:
        _add_uploaded_paths_impl(self, added_paths, ns=globals())

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
            t("instructor.dialog.filter.excel_open"),
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
        _setup_ui_logging_impl(self, ns=globals())

    def _append_user_log(self, message: str) -> None:
        _append_user_log_impl(self, message, ns=globals())

    def _rerender_user_log(self) -> None:
        _rerender_user_log_impl(self, ns=globals())

    def _output_link_markup(self, label: str, path: str | None) -> str:
        return _output_link_markup_impl(self, label, path, ns=globals())

    def _output_links_html(self) -> str:
        return _output_links_html_impl(self, ns=globals())

    def _refresh_output_links(self) -> None:
        _refresh_output_links_impl(self, ns=globals())

    def _on_output_link_activated(self, href: str) -> None:
        _on_output_link_activated_impl(self, href, ns=globals())

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
        remove_file_by_path(self, file_path, ns=globals())

    def _clear_all(self) -> None:
        clear_all(self, ns=globals())

    def closeEvent(self, event) -> None:
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._cancel_token = None
        self._active_jobs.clear()
        if self._ui_log_handler is not None:
            self._logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)




