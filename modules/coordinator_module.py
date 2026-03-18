"""Course coordinator module for collecting Final CO report Excel files."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import (
    QKeySequence,
    QShortcut,
)
from PySide6.QtWidgets import (
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from common.constants import (
    APP_NAME,
    LEVEL_1_THRESHOLD,
    LEVEL_2_THRESHOLD,
    LEVEL_3_THRESHOLD,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
)
from common.jobs import CancellationToken
from common.qt_jobs import run_in_background
from common.drag_drop_file_widget import DragDropFileList, DragDropZoneFrame
from common.removable_file_item_widget import (
    ElidedFileNameLabel as _SharedElidedFileNameLabel,
)
from common.removable_file_item_widget import (
    RemovableFileItemWidget as _SharedRemovableFileItemWidget,
)
from common.texts import t
from common.toast import show_toast
from common.ui_stylings import (
    COORDINATOR_DROP_LIST_ITEM_SPACING,
    COORDINATOR_DROP_LIST_MIN_HEIGHT,
    COORDINATOR_DROP_ZONE_LAYOUT_MARGINS,
    COORDINATOR_DROP_ZONE_LAYOUT_SPACING,
    COORDINATOR_DROPZONE_BG_ACTIVE_ALPHA,
    COORDINATOR_DROPZONE_INNER_RADIUS,
    COORDINATOR_DROPZONE_OUTER_RADIUS,
    COORDINATOR_DROPZONE_BORDER_ACTIVE_ALPHA,
    COORDINATOR_DROPZONE_BORDER_DASH_PATTERN,
    COORDINATOR_DROPZONE_BORDER_INACTIVE_ALPHA,
    COORDINATOR_DROPZONE_BORDER_WIDTH,
    COORDINATOR_DROPZONE_INNER_RECT_ADJUST,
    COORDINATOR_DROPZONE_OUTER_RECT_ADJUST,
    COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
    COORDINATOR_LIST_PLACEHOLDER_COLOR,
    COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
    COORDINATOR_SUMMARY_FONT_SIZE,
)
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
from modules.coordinator.messages import append_user_log as _append_user_log_impl
from modules.coordinator.messages import publish_status as _publish_status_impl
from modules.coordinator.messages import publish_status_key as _publish_status_key_impl
from modules.coordinator.messages import rerender_user_log as _rerender_user_log_impl
from modules.coordinator.messages import setup_ui_logging as _setup_ui_logging_impl
from modules.coordinator.output_links import (
    on_output_link_activated as _on_output_link_activated_impl,
)
from modules.coordinator.output_links import (
    output_link_markup as _output_link_markup_impl,
)
from modules.coordinator.output_links import (
    output_links_html as _output_links_html_impl,
)
from modules.coordinator.output_links import (
    refresh_output_links as _refresh_output_links_impl,
)
from modules.coordinator.steps.calculate_attainment import calculate_attainment_async
from modules.coordinator.steps.collect_files import (
    add_uploaded_paths as _add_uploaded_paths_impl,
)
from modules.coordinator.steps.collect_files import process_files_async
from services import CoordinatorWorkflowService

from .coordinator_processing import (
    _analyze_dropped_files,
    _build_co_attainment_default_name,
    _CoAttainmentWorkbookResult,
    _extract_final_report_signature,
    _generate_co_attainment_workbook,
    _path_key,
)
from .coordinator_processing import (
    _has_valid_final_co_report as _processing_has_valid_final_co_report,
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

OUTPUT_LINK_MODE_FILE = "file"

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


def _messages_namespace() -> dict[str, object]:
    return {
        "emit_user_status": emit_user_status,
        "t": t,
        "build_i18n_log_message": build_i18n_log_message,
        "parse_i18n_log_message": parse_i18n_log_message,
        "resolve_i18n_log_message": resolve_i18n_log_message,
        "format_log_line_at": format_log_line_at,
        "UILogHandler": UILogHandler,
    }


def _file_actions_namespace() -> dict[str, object]:
    return {
        "_path_key": _path_key,
        "user_role": Qt.ItemDataRole.UserRole,
        "log_process_message": log_process_message,
        "build_i18n_log_message": build_i18n_log_message,
        "t": t,
    }


def _output_links_namespace() -> dict[str, object]:
    return {
        "t": t,
        "OUTPUT_LINK_MODE_FILE": OUTPUT_LINK_MODE_FILE,
        "OUTPUT_LINK_MODE_FOLDER": OUTPUT_LINK_MODE_FOLDER,
        "OUTPUT_LINK_SEPARATOR": OUTPUT_LINK_SEPARATOR,
        "show_toast": show_toast,
    }


class _ExcelDropList(DragDropFileList):
    def __init__(self, *, drop_mode: Literal["single", "multiple"] = "multiple") -> None:
        super().__init__(
            placeholder_color=COORDINATOR_LIST_PLACEHOLDER_COLOR,
            placeholder_margins=COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
            placeholder_bottom_margins=COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
            item_spacing=COORDINATOR_DROP_LIST_ITEM_SPACING,
            drop_mode=drop_mode,
        )


class _DropZoneFrame(DragDropZoneFrame):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(
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
            parent=parent,
        )
        self.setObjectName("coordinatorDropZone")


class _ElidedFileNameLabel(_SharedElidedFileNameLabel):
    pass


class _CoordinatorFileItemWidget(_SharedRemovableFileItemWidget):
    def __init__(self, file_path: str, parent: QWidget | None = None) -> None:
        super().__init__(
            file_path,
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            parent=parent,
        )


class CoordinatorModule(QWidget):
    status_changed = Signal(str)
    OUTPUT_LINK_OPEN_FILE_KEY = "instructor.links.open_file"
    OUTPUT_LINK_OPEN_FOLDER_KEY = "instructor.links.open_folder"
    OUTPUT_LINK_NOT_AVAILABLE_KEY = "instructor.links.not_available"
    OUTPUT_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"
    _THRESHOLD_VALIDATION_KEY = "coordinator.thresholds.invalid_rule"

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
        self._threshold_violation_active = False
        self._last_widget_added_count = 0
        self._async_runner = AsyncOperationRunner(self, run_async=run_in_background)
        self._build_ui()
        self._setup_ui_logging()
        self.retranslate_ui()
        self._refresh_ui()

    def _build_ui(self) -> None:
        root = QHBoxLayout(self)

        left = QFrame()
        left.setObjectName("coordinatorLeftCard")
        left_layout = QVBoxLayout(left)
        self.title_label = QLabel()
        self.title_label.setObjectName("coordinatorTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        left_layout.addWidget(self.title_label)
        self.hint_label = QLabel()
        self.hint_label.setObjectName("coordinatorHint")
        self.hint_label.setWordWrap(True)
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        left_layout.addWidget(self.hint_label)
        thresholds_layout = QVBoxLayout()
        self.threshold_title_label = QLabel()
        self.threshold_title_label.setObjectName("coordinatorThresholdTitle")
        thresholds_layout.addWidget(self.threshold_title_label)
        self.threshold_description_label = QLabel()
        self.threshold_description_label.setWordWrap(True)
        thresholds_layout.addWidget(self.threshold_description_label)

        threshold_rows = QGridLayout()
        threshold_rows.setColumnStretch(0, 0)
        threshold_rows.setColumnStretch(1, 1)

        self.threshold_l1_label = QLabel()
        self.threshold_l1_input = QDoubleSpinBox()
        self.threshold_l1_input.setRange(0.0, 100.0)
        self.threshold_l1_input.setDecimals(2)
        self.threshold_l1_input.setSingleStep(0.5)
        self.threshold_l1_input.setValue(float(LEVEL_1_THRESHOLD))
        threshold_rows.addWidget(self.threshold_l1_label, 0, 0)
        threshold_rows.addWidget(
            self.threshold_l1_input,
            0,
            1,
            alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
        )

        self.threshold_l2_label = QLabel()
        self.threshold_l2_input = QDoubleSpinBox()
        self.threshold_l2_input.setRange(0.0, 100.0)
        self.threshold_l2_input.setDecimals(2)
        self.threshold_l2_input.setSingleStep(0.5)
        self.threshold_l2_input.setValue(float(LEVEL_2_THRESHOLD))
        threshold_rows.addWidget(self.threshold_l2_label, 1, 0)
        threshold_rows.addWidget(
            self.threshold_l2_input,
            1,
            1,
            alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
        )

        self.threshold_l3_label = QLabel()
        self.threshold_l3_input = QDoubleSpinBox()
        self.threshold_l3_input.setRange(0.0, 100.0)
        self.threshold_l3_input.setDecimals(2)
        self.threshold_l3_input.setSingleStep(0.5)
        self.threshold_l3_input.setValue(float(LEVEL_3_THRESHOLD))
        threshold_rows.addWidget(self.threshold_l3_label, 2, 0)
        threshold_rows.addWidget(
            self.threshold_l3_input,
            2,
            1,
            alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
        )
        self.threshold_l1_input.valueChanged.connect(self._on_threshold_value_changed)
        self.threshold_l2_input.valueChanged.connect(self._on_threshold_value_changed)
        self.threshold_l3_input.valueChanged.connect(self._on_threshold_value_changed)

        thresholds_layout.addLayout(threshold_rows)
        left_layout.addLayout(thresholds_layout)
        left_layout.addStretch(1)
        left_scroll = QScrollArea()
        left_scroll.setObjectName("coordinatorLeftScroll")
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        left_scroll.setWidget(left)

        right = QFrame()
        right.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right)

        self.drop_zone = _DropZoneFrame()
        self.drop_zone.setProperty("dragActive", False)
        zone_layout = QVBoxLayout(self.drop_zone)
        zone_layout.setContentsMargins(*COORDINATOR_DROP_ZONE_LAYOUT_MARGINS)
        zone_layout.setSpacing(COORDINATOR_DROP_ZONE_LAYOUT_SPACING)
        self.drop_list = _ExcelDropList()
        self.drop_list.setObjectName("coordinatorDropList")
        self.drop_list.setMinimumHeight(COORDINATOR_DROP_LIST_MIN_HEIGHT)
        self.drop_list.files_dropped.connect(self._on_files_dropped)
        self.drop_list.drag_state_changed.connect(self._set_drop_active)
        self.drop_list.browse_requested.connect(self._browse_files)
        zone_layout.addWidget(self.drop_list)
        right_layout.addWidget(self.drop_zone, 1)

        controls_row = QHBoxLayout()
        self.summary_label = QLabel()
        self.summary_label.setObjectName("coordinatorSummary")
        controls_row.addWidget(self.summary_label)
        self.last_added_label = QLabel()
        self.last_added_label.setObjectName("hintText")
        controls_row.addWidget(self.last_added_label)
        controls_row.addStretch(1)
        self.clear_button = QPushButton()
        self.clear_button.setObjectName("coordinatorClearButton")
        self.clear_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_button.setDefault(True)
        self.clear_button.clicked.connect(self._clear_all)
        controls_row.addWidget(self.clear_button)
        self.calculate_button = QPushButton()
        self.calculate_button.setObjectName("coordinatorCalculateButton")
        self.calculate_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.calculate_button.setAutoDefault(False)
        self.calculate_button.setDefault(False)
        self.calculate_button.clicked.connect(self._on_calculate_clicked)
        controls_row.addWidget(self.calculate_button)
        right_layout.addLayout(controls_row)

        self.info_tabs = QTabWidget()
        self.info_tabs.setObjectName("instructorInfoTabs")
        self.info_tabs.currentChanged.connect(self._on_info_tab_changed)

        log_tab = QWidget()
        log_tab_layout = QVBoxLayout(log_tab)

        self.user_log_view = QPlainTextEdit()
        self.user_log_view.setReadOnly(True)
        self.user_log_view.setObjectName("userLogView")
        self.user_log_view.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.user_log_view.setFrameShape(QFrame.Shape.NoFrame)
        log_tab_layout.addWidget(self.user_log_view)

        links_tab = QWidget()
        links_tab_layout = QVBoxLayout(links_tab)

        self.generated_outputs_view = QTextBrowser()
        self.generated_outputs_view.setObjectName("generatedOutputsView")
        self.generated_outputs_view.setOpenExternalLinks(False)
        self.generated_outputs_view.setOpenLinks(False)
        self.generated_outputs_view.setFrameShape(QFrame.Shape.NoFrame)
        self.generated_outputs_view.anchorClicked.connect(
            lambda url: self._on_output_link_activated(url.toString())
        )
        links_tab_layout.addWidget(self.generated_outputs_view)

        self.info_tabs.addTab(log_tab, t("instructor.log.title"))
        self.info_tabs.addTab(links_tab, t("instructor.links.title"))
        right_layout.addWidget(self.info_tabs)

        root.addWidget(left_scroll)
        root.addWidget(right, 1)

        self.shortcut_add_file = QShortcut(QKeySequence(QKeySequence.StandardKey.Open), self)
        self.shortcut_add_file.activated.connect(self._browse_files)
        self.shortcut_save_output = QShortcut(QKeySequence(QKeySequence.StandardKey.Save), self)
        self.shortcut_save_output.activated.connect(self._on_save_shortcut_activated)

    def retranslate_ui(self) -> None:
        self._rerender_user_log()
        self.title_label.setText(t("coordinator.title"))
        self.hint_label.setText(t("coordinator.drop_hint"))
        self.last_added_label.setText(t("coordinator.summary", count=self._last_widget_added_count))
        self.clear_button.setText(t("coordinator.clear_all"))
        self.calculate_button.setText(t("coordinator.calculate"))
        self.threshold_title_label.setText(t("coordinator.thresholds.title"))
        self.threshold_description_label.setText(t("coordinator.thresholds.description"))
        self.threshold_l1_label.setText(t("coordinator.thresholds.l1.label"))
        self.threshold_l2_label.setText(t("coordinator.thresholds.l2.label"))
        self.threshold_l3_label.setText(t("coordinator.thresholds.l3.label"))
        self.info_tabs.setTabText(0, t("instructor.log.title"))
        self.info_tabs.setTabText(1, t("instructor.links.title"))
        self._refresh_output_links()
        self._refresh_summary()

    def _publish_status(self, message: str) -> None:
        _publish_status_impl(self, message, ns=_messages_namespace())

    def _publish_status_key(self, text_key: str, **kwargs: Any) -> None:
        _publish_status_key_impl(self, text_key, ns=_messages_namespace(), **kwargs)

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        self._refresh_ui()

    def _refresh_ui(self) -> None:
        has_files = bool(self._files)
        self.clear_button.setEnabled(has_files and not self.state.busy)
        self.calculate_button.setEnabled(
            has_files and not self.state.busy and self._has_valid_attainment_thresholds()
        )
        self.threshold_l1_input.setEnabled(not self.state.busy)
        self.threshold_l2_input.setEnabled(not self.state.busy)
        self.threshold_l3_input.setEnabled(not self.state.busy)
        self.drop_list.setEnabled(not self.state.busy)
        for row in range(self.drop_list.count()):
            item = self.drop_list.item(row)
            widget = self.drop_list.itemWidget(item)
            if isinstance(widget, _CoordinatorFileItemWidget):
                widget.remove_btn.setEnabled(not self.state.busy)
        self._refresh_output_links()
        self._refresh_summary()

    def _read_attainment_thresholds(self) -> tuple[float, float, float]:
        return (
            float(self.threshold_l1_input.value()),
            float(self.threshold_l2_input.value()),
            float(self.threshold_l3_input.value()),
        )

    def _has_valid_attainment_thresholds(self) -> bool:
        l1, l2, l3 = self._read_attainment_thresholds()
        return 0.0 < l1 < l2 < l3 < 100.0

    def get_attainment_thresholds(self) -> tuple[float, float, float] | None:
        thresholds = self._read_attainment_thresholds()
        if not self._has_valid_attainment_thresholds():
            self._notify_threshold_violation(force=True)
            return None
        return thresholds

    def _notify_threshold_violation(self, *, force: bool) -> None:
        if self._threshold_violation_active and not force:
            return
        show_toast(
            self,
            t(self._THRESHOLD_VALIDATION_KEY),
            title=t("coordinator.title"),
            level="error",
        )
        self._publish_status_key(self._THRESHOLD_VALIDATION_KEY)
        self._threshold_violation_active = True

    def _on_threshold_value_changed(self, _value: float) -> None:
        if self._has_valid_attainment_thresholds():
            self._threshold_violation_active = False
        else:
            self._notify_threshold_violation(force=False)
        self._refresh_ui()

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
        self._last_widget_added_count = len([value for value in dropped_files if value])
        self.last_added_label.setText(t("coordinator.summary", count=self._last_widget_added_count))
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
            self._last_widget_added_count = len(selected_files)
            self.last_added_label.setText(t("coordinator.summary", count=self._last_widget_added_count))
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
        _setup_ui_logging_impl(self, ns=_messages_namespace())

    def _append_user_log(self, message: str) -> None:
        _append_user_log_impl(self, message, ns=_messages_namespace())

    def _rerender_user_log(self) -> None:
        _rerender_user_log_impl(self, ns=_messages_namespace())

    def _output_link_markup(self, label: str, path: str | None) -> str:
        return _output_link_markup_impl(self, label, path, ns=_output_links_namespace())

    def _output_links_html(self) -> str:
        return _output_links_html_impl(self, ns=_output_links_namespace())

    def _refresh_output_links(self) -> None:
        _refresh_output_links_impl(self, ns=_output_links_namespace())

    def _on_output_link_activated(self, href: str) -> None:
        _on_output_link_activated_impl(self, href, ns=_output_links_namespace())

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
        self.drop_zone.set_drag_active(active)

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self.info_tabs.setVisible(not enabled)

    def get_shared_outputs_html(self) -> str:
        return self._output_links_html()

    def _remove_file_by_path(self, file_path: str) -> None:
        remove_file_by_path(self, file_path, ns=_file_actions_namespace())

    def _clear_all(self) -> None:
        self._last_widget_added_count = 0
        self.last_added_label.setText(t("coordinator.summary", count=self._last_widget_added_count))
        clear_all(self, ns=_file_actions_namespace())

    def closeEvent(self, event) -> None:
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._cancel_token = None
        self._active_jobs.clear()
        if self._ui_log_handler is not None:
            self._logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)




