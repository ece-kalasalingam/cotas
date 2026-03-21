"""Course coordinator module for collecting Final CO report Excel files."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from PySide6.QtCore import Qt, QUrl, Signal
from PySide6.QtGui import (
    QDesktopServices,
    QKeySequence,
    QPalette,
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

from common.async_operation_runner import AsyncOperationRunner
from common.constants import (
    APP_NAME,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    LEVEL_1_THRESHOLD,
    LEVEL_2_THRESHOLD,
    LEVEL_3_THRESHOLD,
    MODULE_LEFT_PANE_CONTENT_MARGINS,
    MODULE_LEFT_PANE_LAYOUT_SPACING,
    MODULE_LEFT_PANE_WIDTH_OFFSET,
    OUTPUT_LINK_MODE_FILE,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
)
from common.drag_drop_file_widget import (
    DragDropFileList,
    DragDropZoneFrame,
    ManagedDropFileWidget,
)
from common.exceptions import JobCancelledError
from common.jobs import CancellationToken
from common.module_messages import rerender_user_log as _rerender_user_log_impl
from common.module_runtime import ModuleRuntime
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.output_panel import OutputItem, OutputPanelData, open_output_link, render_output_panel_html
from common.qt_jobs import run_in_background
from common.removable_file_item_widget import (
    ElidedFileNameLabel as _SharedElidedFileNameLabel,
)
from common.removable_file_item_widget import (
    RemovableFileItemWidget as _SharedRemovableFileItemWidget,
)
from common.texts import t
from common.toast import show_toast
from common.ui_logging import (
    UILogHandler,
    build_i18n_log_message,
    format_log_line_at,
    parse_i18n_log_message,
    resolve_i18n_log_message,
)
from common.ui_stylings import (
    COORDINATOR_DROP_LIST_ITEM_SPACING,
    COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
    COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
    GLOBAL_QPUSHBUTTON_MIN_WIDTH,
)
from common.utils import (
    emit_user_status,
    log_process_message,
    resolve_dialog_start_path,
)
from domain.coordinator_engine import (
    _analyze_dropped_files,
    _build_co_attainment_default_name,
    _CoAttainmentWorkbookResult,
    _extract_final_report_signature,
    _generate_co_attainment_workbook,
)
from domain.coordinator_engine import (
    _has_valid_final_co_report as _processing_has_valid_final_co_report,
)
from domain.coordinator_engine import (
    _path_key,
)
from modules.coordinator.contracts import require_keys
from modules.coordinator.file_actions import clear_all, remove_file_by_path
from modules.coordinator.messages import show_threshold_validation_toast
from modules.coordinator.steps.calculate_attainment import calculate_attainment_async
from modules.coordinator.steps.collect_files import (
    add_uploaded_paths as _add_uploaded_paths_impl,
)
from modules.coordinator.steps.collect_files import (
    process_files_async,
)
from modules.coordinator.workflow_controller import CoordinatorWorkflowController
from modules.coordinator.output_links import (
    output_link_markup as _legacy_output_link_markup,
)
from modules.coordinator.output_links import (
    output_links_html as _legacy_output_links_html,
)
from services import CoordinatorWorkflowService

_logger = logging.getLogger(__name__)
_QT_COMPAT_EXPORTS = (QListWidget,)
_LEFT_PANE_WIDTH = GLOBAL_QPUSHBUTTON_MIN_WIDTH + MODULE_LEFT_PANE_WIDTH_OFFSET


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


def _collect_files_namespace() -> dict[str, object]:
    return {
        "t": t,
        "_path_key": _path_key,
        "show_toast": show_toast,
        "log_process_message": log_process_message,
        "build_i18n_log_message": build_i18n_log_message,
        "_analyze_dropped_files": _analyze_dropped_files,
        "QListWidgetItem": QListWidgetItem,
        "JobCancelledError": JobCancelledError,
    }


def _calculate_attainment_namespace() -> dict[str, object]:
    return {
        "t": t,
        "QFileDialog": QFileDialog,
        "APP_NAME": APP_NAME,
        "resolve_dialog_start_path": resolve_dialog_start_path,
        "_extract_final_report_signature": _extract_final_report_signature,
        "_build_co_attainment_default_name": _build_co_attainment_default_name,
        "_CoAttainmentWorkbookResult": _CoAttainmentWorkbookResult,
        "_path_key": _path_key,
        "log_process_message": log_process_message,
        "build_i18n_log_message": build_i18n_log_message,
        "show_toast": show_toast,
        "_generate_co_attainment_workbook": _generate_co_attainment_workbook,
        "JobCancelledError": JobCancelledError,
    }


def _output_links_namespace() -> dict[str, object]:
    return {
        "t": t,
        "OUTPUT_LINK_MODE_FILE": OUTPUT_LINK_MODE_FILE,
        "OUTPUT_LINK_MODE_FOLDER": OUTPUT_LINK_MODE_FOLDER,
        "OUTPUT_LINK_SEPARATOR": OUTPUT_LINK_SEPARATOR,
        "show_toast": show_toast,
    }


def _output_link_markup_impl(module: object, label: str, path: str | None, ns: dict[str, object]) -> str:
    return _legacy_output_link_markup(module, label, path, ns=ns)


def _output_links_html_impl(module: object, ns: dict[str, object]) -> str:
    return _legacy_output_links_html(module, ns=ns)


def _refresh_output_links_impl(module: object, ns: dict[str, object]) -> None:
    typed_module = module if isinstance(module, CoordinatorModule) else None
    if typed_module is None:
        return
    typed_module.generated_outputs_view.setHtml(
        render_output_panel_html(
            typed_module.get_shared_outputs_data(),
            translate=t,
            output_link_mode_file=OUTPUT_LINK_MODE_FILE,
            output_link_mode_folder=OUTPUT_LINK_MODE_FOLDER,
            output_link_separator=OUTPUT_LINK_SEPARATOR,
        )
    )


def _on_output_link_activated_impl(module: object, href: str, ns: dict[str, object]) -> None:
    typed_module = module if isinstance(module, CoordinatorModule) else None
    if typed_module is None:
        return
    opened = open_output_link(
        href,
        output_link_mode_folder=OUTPUT_LINK_MODE_FOLDER,
        output_link_separator=OUTPUT_LINK_SEPARATOR,
        open_path=lambda target: QDesktopServices.openUrl(QUrl.fromLocalFile(str(target))),
    )
    if opened:
        return
    show_toast(
        typed_module,
        t(typed_module.OUTPUT_LINK_OPEN_FAILED_KEY),
        title=t("instructor.msg.error_title"),
        level="error",
    )


def _validate_coordinator_namespaces() -> None:
    require_keys(
        _collect_files_namespace(),
        keys=(
            "t",
            "_path_key",
            "show_toast",
            "log_process_message",
            "build_i18n_log_message",
            "_analyze_dropped_files",
            "QListWidgetItem",
            "JobCancelledError",
        ),
        context="coordinator.collect_files",
    )
    require_keys(
        _calculate_attainment_namespace(),
        keys=(
            "t",
            "QFileDialog",
            "APP_NAME",
            "resolve_dialog_start_path",
            "_extract_final_report_signature",
            "_build_co_attainment_default_name",
            "_CoAttainmentWorkbookResult",
            "_path_key",
            "log_process_message",
            "build_i18n_log_message",
            "show_toast",
            "_generate_co_attainment_workbook",
            "JobCancelledError",
        ),
        context="coordinator.calculate_attainment",
    )


class _ExcelDropList(DragDropFileList):
    def __init__(self, *, drop_mode: Literal["single", "multiple"] = "multiple") -> None:
        super().__init__(
            placeholder_margins=COORDINATOR_LIST_PLACEHOLDER_TEXT_MARGINS,
            placeholder_bottom_margins=COORDINATOR_LIST_PLACEHOLDER_BOTTOM_MARGINS,
            item_spacing=COORDINATOR_DROP_LIST_ITEM_SPACING,
            drop_mode=drop_mode,
        )


class _DropZoneFrame(DragDropZoneFrame):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent=parent)
        self.setObjectName("coordinatorDropZone")


class _ElidedFileNameLabel(_SharedElidedFileNameLabel):
    pass


class _CoordinatorFileItemWidget(_SharedRemovableFileItemWidget):
    def __init__(self, file_path: str, parent: QWidget | None = None) -> None:
        super().__init__(
            file_path,
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            open_file_tooltip=t("instructor.links.open_file"),
            open_folder_tooltip=t("instructor.links.open_folder"),
            remove_tooltip=t("coordinator.file.remove_tooltip"),
            parent=parent,
        )


class CoordinatorModule(QWidget):
    status_changed = Signal(str)
    OUTPUT_LINK_OPEN_FILE_KEY = "instructor.links.open_file"
    OUTPUT_LINK_OPEN_FOLDER_KEY = "instructor.links.open_folder"
    OUTPUT_LINK_NOT_AVAILABLE_KEY = "instructor.links.not_available"
    OUTPUT_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"
    _THRESHOLD_VALIDATION_KEY = "coordinator.thresholds.invalid_rule"

    def __init__(
        self,
        *,
        workflow_service: CoordinatorWorkflowService | None = None,
        async_runner: AsyncOperationRunner | None = None,
    ) -> None:
        super().__init__()
        _validate_coordinator_namespaces()
        self._files: list[Path] = []
        self._downloaded_outputs: list[Path] = []
        self._logger = _logger
        self.state = CoordinatorWorkflowState()
        self._workflow_service = workflow_service or CoordinatorWorkflowService()
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._pending_drop_batches: list[list[str]] = []
        self._ui_log_handler: UILogHandler | None = None
        self._user_log_entries: list[dict[str, object]] = []
        self._threshold_violation_active = False
        self._workflow_controller = CoordinatorWorkflowController(self)
        self._async_runner = async_runner or AsyncOperationRunner(self, run_async=run_in_background)
        self._runtime = ModuleRuntime(
            module=self,
            app_name=APP_NAME,
            logger=self._logger,
            async_runner=self._async_runner,
            messages_namespace_factory=_messages_namespace,
        )
        self._build_ui()
        self._setup_ui_logging()
        self.retranslate_ui()
        self._refresh_ui()

    def _build_ui(self) -> None:
        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                top_object_name="coordinatorTopRegion",
                footer_height=INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
            ),
        )
        top_pane = QWidget()
        top_layout = QHBoxLayout(top_pane)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        self._ui_engine.set_top_widget(top_pane)
        left_pane = QFrame()
        left_pane.setObjectName("coordinatorLeftCard")
        left_pane.setFrameShape(QFrame.Shape.StyledPanel)
        left_pane.setFrameShadow(QFrame.Shadow.Raised)
        left_palette = left_pane.palette()
        left_palette.setColor(
            QPalette.ColorRole.Window,
            left_palette.color(QPalette.ColorRole.Base),
        )
        left_pane.setPalette(left_palette)
        left_pane.setAutoFillBackground(True)
        left_layout = QVBoxLayout(left_pane)
        left_layout.setContentsMargins(*MODULE_LEFT_PANE_CONTENT_MARGINS)
        left_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
        left_scroll = QScrollArea()
        left_scroll.setObjectName("coordinatorLeftScroll")
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        left_scroll.viewport().setObjectName("coordinatorLeftScrollViewport")
        left_scroll.viewport().setAutoFillBackground(True)
        left_scroll_viewport_palette = left_scroll.viewport().palette()
        left_scroll_viewport_palette.setColor(
            QPalette.ColorRole.Window,
            left_scroll_viewport_palette.color(QPalette.ColorRole.Base),
        )
        left_scroll.viewport().setPalette(left_scroll_viewport_palette)
        left_scroll.setWidget(left_pane)
        left_scroll.setFixedWidth(_LEFT_PANE_WIDTH)
        top_layout.addWidget(left_scroll, 0)
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
        thresholds_layout.setContentsMargins(0, 0, 0, 0)
        thresholds_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
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
        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_pane)
        right_scroll = QScrollArea()
        right_scroll.setObjectName("coordinatorRightScroll")
        right_scroll.setFrameShape(QFrame.Shape.NoFrame)
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        right_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        right_scroll.setWidget(right_pane)
        top_layout.addWidget(right_scroll, 1)

        self.drop_widget = ManagedDropFileWidget(
            drop_mode="multiple",
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            open_file_tooltip=t("instructor.links.open_file"),
            open_folder_tooltip=t("instructor.links.open_folder"),
            remove_tooltip=t("coordinator.file.remove_tooltip"),
        )
        self.drop_widget.set_summary_text_builder(
            lambda _count: t("coordinator.summary", count=len(self._files))
        )
        self.drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.drop_widget.drop_list.setObjectName("coordinatorDropList")
        self.drop_widget.files_dropped.connect(self._on_files_dropped)
        self.drop_widget.browse_requested.connect(self._browse_files)
        self.drop_widget.clear_button.clicked.connect(self._clear_all)
        self.drop_widget.submit_requested.connect(self._on_calculate_clicked)
        self.drop_widget.set_clear_button_text(t("coordinator.clear_all"))
        right_layout.addWidget(self.drop_widget, 1)

        # Shortcut aliases for direct access in module helpers/tests.
        self.drop_zone = self.drop_widget.drop_zone
        self.drop_list = self.drop_widget.drop_list
        self.clear_button = self.drop_widget.clear_button
        self.calculate_button = self.drop_widget.submit_button
        self.calculate_button.setObjectName("coordinatorCalculateButton")
        self.calculate_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.calculate_button.setAutoDefault(False)
        self.calculate_button.setDefault(False)

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
        footer_pane = QWidget()
        footer_layout = QVBoxLayout(footer_pane)
        footer_layout.setContentsMargins(0, 0, 0, 0)
        footer_layout.setSpacing(0)
        footer_layout.addWidget(self.info_tabs)
        self._ui_engine.set_footer_widget(footer_pane)

        self.shortcut_add_file = QShortcut(QKeySequence(QKeySequence.StandardKey.Open), self)
        self.shortcut_add_file.activated.connect(self._browse_files)
        self.shortcut_save_output = QShortcut(QKeySequence(QKeySequence.StandardKey.Save), self)
        self.shortcut_save_output.activated.connect(self._on_save_shortcut_activated)

    def retranslate_ui(self) -> None:
        self._rerender_user_log()
        self.title_label.setText(t("coordinator.title"))
        self.hint_label.setText(t("coordinator.drop_hint"))
        self.drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.drop_widget.set_clear_button_text(t("coordinator.clear_all"))
        self.drop_widget.set_summary_text_builder(
            lambda _count: t("coordinator.summary", count=len(self._files))
        )
        self.drop_widget.set_submit_button_text(t("coordinator.calculate"))
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
        self._runtime.publish_status(message)

    def _publish_status_key(self, text_key: str, **kwargs: Any) -> None:
        self._runtime.publish_status_key(text_key, **kwargs)

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        self._refresh_ui()

    def _refresh_ui(self) -> None:
        has_files = bool(self._files)
        can_submit = has_files and self._has_valid_attainment_thresholds() and (not self.state.busy)
        self.drop_widget.setEnabled(not self.state.busy)
        self.drop_widget.set_submit_allowed(can_submit)
        self.threshold_l1_input.setEnabled(not self.state.busy)
        self.threshold_l2_input.setEnabled(not self.state.busy)
        self.threshold_l3_input.setEnabled(not self.state.busy)
        self.drop_list.setEnabled(not self.state.busy)
        self.clear_button.setEnabled(has_files and (not self.state.busy))
        self.calculate_button.setEnabled(can_submit)
        for row in range(self.drop_list.count()):
            item = self.drop_list.item(row)
            widget = self.drop_list.itemWidget(item)
            if isinstance(widget, _CoordinatorFileItemWidget):
                widget.remove_btn.setEnabled(not self.state.busy)
        self._refresh_output_links()
        self._refresh_summary()

    def _read_attainment_thresholds(self) -> tuple[float, float, float]:
        return self._workflow_controller.read_attainment_thresholds()

    def _has_valid_attainment_thresholds(self) -> bool:
        return self._workflow_controller.has_valid_attainment_thresholds()

    def get_attainment_thresholds(self) -> tuple[float, float, float] | None:
        thresholds = self._read_attainment_thresholds()
        if not self._has_valid_attainment_thresholds():
            self._notify_threshold_violation(force=True)
            return None
        return thresholds

    def _notify_threshold_violation(self, *, force: bool) -> None:
        self._workflow_controller.notify_threshold_violation(force=force)

    def _show_threshold_validation_toast(self) -> None:
        show_threshold_validation_toast(
            self,
            message_key=self._THRESHOLD_VALIDATION_KEY,
            title_key="coordinator.title",
            toast_fn=show_toast,
            translate=t,
        )

    def _on_threshold_value_changed(self, _value: float) -> None:
        self._workflow_controller.on_threshold_value_changed()
        self._refresh_ui()

    def _on_calculate_clicked(self) -> None:
        calculate_attainment_async(self, ns=_calculate_attainment_namespace())

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
        process_files_async(self, dropped_files, ns=_collect_files_namespace())

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
        self._runtime.set_async_runner(self._async_runner)
        self._runtime.start_async_operation(
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
        _add_uploaded_paths_impl(self, added_paths, ns=_collect_files_namespace())

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
        self._runtime.remember_dialog_dir_safe(selected_path)

    def _setup_ui_logging(self) -> None:
        self._runtime.setup_ui_logging()

    def _append_user_log(self, message: str) -> None:
        self._runtime.append_user_log(message)

    def _rerender_user_log(self) -> None:
        _rerender_user_log_impl(self, ns=_messages_namespace())

    def _output_items(self) -> tuple[OutputItem, ...]:
        items: list[OutputItem] = []
        items.extend(
            OutputItem(label_key="coordinator.links.uploaded_report", path=str(path))
            for path in self._files
        )
        items.extend(
            OutputItem(label_key="coordinator.links.downloaded_output", path=str(path))
            for path in self._downloaded_outputs
        )
        return tuple(items)

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
        self.drop_widget.set_summary_text_builder(
            lambda _count: t("coordinator.summary", count=len(self._files))
        )


    def _set_drop_active(self, active: bool) -> None:
        self.drop_zone.set_drag_active(active)

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self.info_tabs.setVisible(not enabled)
        self._ui_engine.set_footer_visible(not enabled)

    def get_shared_outputs_data(self) -> OutputPanelData:
        return OutputPanelData(items=self._output_items(), open_failed_key=self.OUTPUT_LINK_OPEN_FAILED_KEY)

    def _output_link_markup(self, label: str, path: str | None) -> str:
        return _output_link_markup_impl(self, label, path, ns=_output_links_namespace())

    def _output_links_html(self) -> str:
        return _output_links_html_impl(self, ns=_output_links_namespace())

    def get_shared_outputs_html(self) -> str:
        return self._output_links_html()

    def _remove_file_by_path(self, file_path: str) -> None:
        remove_file_by_path(self, file_path, ns=_file_actions_namespace())

    def _clear_all(self) -> None:
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




