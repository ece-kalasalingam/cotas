"""CO Analysis module."""

from __future__ import annotations

import logging
from pathlib import Path, PureWindowsPath
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QListWidgetItem,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from common.async_operation_runner import AsyncOperationRunner
from common.constants import (
    APP_NAME,
    CO_ATTAINMENT_LEVEL_DEFAULT,
    CO_ATTAINMENT_PERCENT_DEFAULT,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    LEVEL_1_THRESHOLD,
    LEVEL_2_THRESHOLD,
    LEVEL_3_THRESHOLD,
    MODULE_LEFT_PANE_CONTENT_MARGINS,
    MODULE_LEFT_PANE_LAYOUT_SPACING,
    MODULE_LEFT_PANE_WIDTH_OFFSET,
)
from common.registry import (
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
)
from common.contracts import require_keys
from common.drag_drop_file_widget import ManagedDropFileWidget
from common.exceptions import JobCancelledError
from common.jobs import CancellationToken
from common.module_messages import build_status_message as _build_status_message
from common.module_messages import default_messages_namespace as _default_messages_namespace
from common.module_messages import rerender_user_log as _rerender_user_log_impl
from common.module_messages import show_toast_plain as _show_toast_plain
from common.module_runtime import ModuleRuntime
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.output_panel import OutputItem, OutputPanelData
from common.qt_jobs import run_in_background
from common.removable_file_item_widget import (
    RemovableFileItemWidget as _SharedRemovableFileItemWidget,
)
from common.i18n import t
from common.ui_stylings import GLOBAL_QPUSHBUTTON_MIN_WIDTH
from common.utils import (
    canonical_path_key,
    log_process_message,
    normalize,
    resolve_dialog_start_path,
    sanitize_filename_token,
)
from domain import BusyWorkflowState
from domain.co_analysis_engine import (
    extract_course_metadata_and_students,
)
from modules.co_analysis.file_actions import clear_all, remove_file_by_path
from modules.co_analysis.messages import show_threshold_validation_toast
from modules.co_analysis.steps.collect_files import collect_files_async
from modules.co_analysis.steps.generate_workbook import save_workbook_async
from modules.co_analysis.validators.uploaded_workbook_validator import (
    consume_last_source_anomaly_warnings,
    validate_uploaded_source_workbook,
)
from modules.co_analysis.workflow_controller import COAnalysisWorkflowController
from services import CoAnalysisWorkflowService

_logger = logging.getLogger(__name__)
_LEFT_PANE_WIDTH = GLOBAL_QPUSHBUTTON_MIN_WIDTH + MODULE_LEFT_PANE_WIDTH_OFFSET


def _messages_namespace() -> dict[str, object]:
    return dict(_default_messages_namespace(translate=t))


def _build_i18n_message(text_key: str, *, kwargs: dict[str, object] | None = None, fallback: str | None = None) -> str:
    return _build_status_message(text_key, translate=t, kwargs=kwargs, fallback=fallback)


def _collect_files_namespace() -> dict[str, object]:
    return {
        "_validate_uploaded_source_workbook": validate_uploaded_source_workbook,
        "_consume_last_source_anomaly_warnings": consume_last_source_anomaly_warnings,
        "t": t,
        "show_toast": _show_toast_plain,
        "log_process_message": log_process_message,
        "build_i18n_log_message": _build_i18n_message,
        "JobCancelledError": JobCancelledError,
    }


def _generate_workbook_namespace() -> dict[str, object]:
    return {
        "QFileDialog": QFileDialog,
        "APP_NAME": APP_NAME,
        "t": t,
        "resolve_dialog_start_path": resolve_dialog_start_path,
        "canonical_path_key": canonical_path_key,
        "_extract_course_metadata_and_students": extract_course_metadata_and_students,
        "_sanitize_filename_token": sanitize_filename_token,
        "log_process_message": log_process_message,
        "build_i18n_log_message": _build_i18n_message,
        "show_toast": _show_toast_plain,
        "normalize": normalize,
        "COURSE_METADATA_COURSE_CODE_KEY": COURSE_METADATA_COURSE_CODE_KEY,
        "COURSE_METADATA_ACADEMIC_YEAR_KEY": COURSE_METADATA_ACADEMIC_YEAR_KEY,
        "JobCancelledError": JobCancelledError,
    }


def _file_actions_namespace() -> dict[str, object]:
    return {
        "canonical_path_key": canonical_path_key,
        "user_role": Qt.ItemDataRole.UserRole,
        "log_process_message": log_process_message,
        "build_i18n_log_message": _build_i18n_message,
        "t": t,
    }


class _LogSink:
    def appendPlainText(self, _text: str) -> None:  # noqa: N802 - Qt-style name
        return

    def clear(self) -> None:
        return


class _COAnalysisFileItemWidget(_SharedRemovableFileItemWidget):
    def __init__(self, file_path: str, parent: QWidget | None = None) -> None:
        super().__init__(
            file_path,
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            open_file_tooltip=t("outputs.open_file"),
            open_folder_tooltip=t("outputs.open_folder"),
            remove_tooltip=t("coordinator.file.remove_tooltip"),
            parent=parent,
        )


def _validate_co_analysis_namespaces() -> None:
    require_keys(
        _collect_files_namespace(),
        keys=(
            "_validate_uploaded_source_workbook",
            "_consume_last_source_anomaly_warnings",
            "t",
            "show_toast",
            "log_process_message",
            "build_i18n_log_message",
            "JobCancelledError",
        ),
        context="co_analysis.collect_files",
    )
    require_keys(
        _generate_workbook_namespace(),
        keys=(
            "QFileDialog",
            "APP_NAME",
            "t",
            "resolve_dialog_start_path",
            "canonical_path_key",
            "_extract_course_metadata_and_students",
            "_sanitize_filename_token",
            "log_process_message",
            "build_i18n_log_message",
            "show_toast",
            "normalize",
            "COURSE_METADATA_COURSE_CODE_KEY",
            "COURSE_METADATA_ACADEMIC_YEAR_KEY",
            "JobCancelledError",
        ),
        context="co_analysis.generate_workbook",
    )
    require_keys(
        _file_actions_namespace(),
        keys=(
            "canonical_path_key",
            "user_role",
            "log_process_message",
            "build_i18n_log_message",
            "t",
        ),
        context="co_analysis.file_actions",
    )


class COAnalysisModule(QWidget):
    status_changed = Signal(str)
    _THRESHOLD_VALIDATION_KEY = "coordinator.thresholds.invalid_rule"
    _CO_ATTAINMENT_TARGET_VALIDATION_KEY = "coordinator.co_attainment.invalid_percent"

    def __init__(
        self,
        *,
        workflow_service: CoAnalysisWorkflowService | None = None,
        async_runner: AsyncOperationRunner | None = None,
    ) -> None:
        super().__init__()
        _validate_co_analysis_namespaces()
        self._files: list[Path] = []
        self._downloaded_outputs: list[Path] = []
        self._logger = _logger
        self.state = BusyWorkflowState()
        self._workflow_service = workflow_service or CoAnalysisWorkflowService()
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._pending_drop_batches: list[list[str]] = []
        self._ui_log_handler: logging.Handler | None = None
        self._user_log_entries: list[dict[str, object]] = []
        self._threshold_violation_active = False
        self._workflow_controller = COAnalysisWorkflowController(self)
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
                top_object_name="coAnalysisTopRegion",
                footer_height=INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
            ),
        )
        top_pane = QWidget()
        top_layout = QHBoxLayout(top_pane)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        self._ui_engine.set_top_widget(top_pane)

        left_card = QFrame()
        left_card.setObjectName("coordinatorLeftCard")
        left_card.setFrameShape(QFrame.Shape.StyledPanel)
        left_card.setFrameShadow(QFrame.Shadow.Raised)
        left_card_layout = QVBoxLayout(left_card)
        left_card_layout.setContentsMargins(0, 0, 0, 0)
        left_card_layout.setSpacing(0)

        left_scroll = QScrollArea()
        left_scroll.setObjectName("coordinatorLeftScroll")
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        left_scroll.viewport().setObjectName("coordinatorLeftScrollViewport")
        left_content = QWidget()
        left_layout = QVBoxLayout(left_content)
        left_layout.setContentsMargins(*MODULE_LEFT_PANE_CONTENT_MARGINS)
        left_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
        left_scroll.setWidget(left_content)
        left_card_layout.addWidget(left_scroll, 1)
        left_card.setFixedWidth(_LEFT_PANE_WIDTH)
        top_layout.addWidget(left_card, 0)

        self.title_label = QLabel()
        self.title_label.setObjectName("coordinatorTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        left_layout.addWidget(self.title_label)

        self.hint_label = QLabel()
        self.hint_label.setObjectName("coordinatorHint")
        self.hint_label.setWordWrap(True)
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignJustify | Qt.AlignmentFlag.AlignTop)
        left_layout.addWidget(self.hint_label)

        thresholds_layout = QVBoxLayout()
        thresholds_layout.setContentsMargins(0, 0, 0, 0)
        thresholds_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
        self.threshold_title_label = QLabel()
        self.threshold_title_label.setObjectName("coordinatorThresholdTitle")
        thresholds_layout.addWidget(self.threshold_title_label)
        self.threshold_description_label = QLabel()
        self.threshold_description_label.setWordWrap(True)
        self.threshold_description_label.setAlignment(
            Qt.AlignmentFlag.AlignJustify | Qt.AlignmentFlag.AlignTop
        )
        thresholds_layout.addWidget(self.threshold_description_label)

        threshold_rows = QGridLayout()
        threshold_rows.setColumnStretch(0, 0)
        threshold_rows.setColumnStretch(1, 1)

        self.threshold_l1_label = QLabel()
        self.threshold_l1_label.setObjectName("coordinatorThresholdL1Label")
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
        self.threshold_l2_label.setObjectName("coordinatorThresholdInputLabel")
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
        self.threshold_l3_label.setObjectName("coordinatorThresholdInputLabel")
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
        self.threshold_l1_input.editingFinished.connect(self._on_threshold_editing_finished)
        self.threshold_l2_input.editingFinished.connect(self._on_threshold_editing_finished)
        self.threshold_l3_input.editingFinished.connect(self._on_threshold_editing_finished)

        thresholds_layout.addLayout(threshold_rows)

        self.co_attainment_description_label = QLabel()
        self.co_attainment_description_label.setWordWrap(True)
        self.co_attainment_description_label.setAlignment(
            Qt.AlignmentFlag.AlignJustify | Qt.AlignmentFlag.AlignTop
        )
        thresholds_layout.addWidget(self.co_attainment_description_label)

        co_attainment_rows = QGridLayout()
        co_attainment_rows.setColumnStretch(0, 0)
        co_attainment_rows.setColumnStretch(1, 1)

        self.co_attainment_percent_label = QLabel()
        self.co_attainment_percent_label.setObjectName("coordinatorThresholdInputLabel")
        self.co_attainment_percent_input = QDoubleSpinBox()
        self.co_attainment_percent_input.setRange(0.0, 100.0)
        self.co_attainment_percent_input.setDecimals(2)
        self.co_attainment_percent_input.setSingleStep(0.5)
        self.co_attainment_percent_input.setValue(float(CO_ATTAINMENT_PERCENT_DEFAULT))
        co_attainment_rows.addWidget(self.co_attainment_percent_label, 0, 0)
        co_attainment_rows.addWidget(
            self.co_attainment_percent_input,
            0,
            1,
            alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
        )

        self.co_attainment_level_label = QLabel()
        self.co_attainment_level_label.setObjectName("coordinatorThresholdInputLabel")
        self.co_attainment_level_input = QComboBox()
        self.co_attainment_level_input.addItem("L1", 1)
        self.co_attainment_level_input.addItem("L2", 2)
        self.co_attainment_level_input.addItem("L3", 3)
        default_level_index = max(
            0,
            min(self.co_attainment_level_input.count() - 1, CO_ATTAINMENT_LEVEL_DEFAULT - 1),
        )
        self.co_attainment_level_input.setCurrentIndex(default_level_index)
        co_attainment_rows.addWidget(self.co_attainment_level_label, 1, 0)
        co_attainment_rows.addWidget(
            self.co_attainment_level_input,
            1,
            1,
            alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
        )

        self.co_attainment_percent_input.valueChanged.connect(self._on_threshold_value_changed)
        self.co_attainment_percent_input.editingFinished.connect(self._on_threshold_editing_finished)
        self.co_attainment_level_input.currentIndexChanged.connect(
            lambda _idx: self._on_threshold_value_changed(0.0)
        )
        self.co_attainment_level_input.activated.connect(lambda _idx: self._on_threshold_editing_finished())
        thresholds_layout.addLayout(co_attainment_rows)
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
        right_scroll.viewport().setObjectName("coordinatorRightScrollViewport")
        right_scroll.setWidget(right_pane)
        top_layout.addWidget(right_scroll, 1)

        self.drop_widget = ManagedDropFileWidget(
            drop_mode="multiple",
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            open_file_tooltip=t("outputs.open_file"),
            open_folder_tooltip=t("outputs.open_folder"),
            remove_tooltip=t("coordinator.file.remove_tooltip"),
        )
        self.drop_widget.set_summary_text_builder(
            lambda _count: t("coordinator.summary", count=len(self._files))
        )
        self.drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.drop_widget.drop_list.setObjectName("coordinatorDropList")
        self.drop_widget.files_dropped.connect(self._on_files_dropped)
        self.drop_widget.drop_list.items_reordered.connect(self._on_drop_list_reordered)
        self.drop_widget.browse_requested.connect(self._browse_files)
        self.drop_widget.clear_button.clicked.connect(self._clear_all)
        self.drop_widget.submit_requested.connect(self._on_submit_requested)
        self.drop_widget.set_clear_button_text(t("coordinator.clear_all"))
        right_layout.addWidget(self.drop_widget, 1)

        self.drop_zone = self.drop_widget.drop_zone
        self.drop_list = self.drop_widget.drop_list
        self.clear_button = self.drop_widget.clear_button
        self.calculate_button = self.drop_widget.submit_button
        self.calculate_button.setObjectName("primaryAction")
        self.calculate_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.calculate_button.setAutoDefault(False)
        self.calculate_button.setDefault(False)

        self.user_log_view = _LogSink()
        self._ui_engine.set_footer_visible(False)

        self.shortcut_add_file = QShortcut(QKeySequence(QKeySequence.StandardKey.Open), self)
        self.shortcut_add_file.activated.connect(self._browse_files)

    def retranslate_ui(self) -> None:
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
        self.co_attainment_description_label.setText(t("coordinator.co_attainment.description"))
        self.co_attainment_percent_label.setText(t("coordinator.co_attainment.percent.label"))
        self.co_attainment_level_label.setText(t("coordinator.co_attainment.level.label"))
        self._refresh_summary()
        self._rerender_user_log()

    def _publish_status(self, message: str) -> None:
        self._runtime.publish_status(message)

    def _publish_status_key(self, text_key: str, **kwargs: Any) -> None:
        self._runtime.publish_status_key(text_key, **kwargs)

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        self._refresh_ui()

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

    def _drain_next_batch(self) -> None:
        if self.state.busy or not self._pending_drop_batches:
            return
        next_batch = self._pending_drop_batches.pop(0)
        self._collect_valid_files(next_batch)

    def _refresh_ui(self) -> None:
        has_files = bool(self._files)
        can_submit = (
            has_files
            and self._has_valid_attainment_thresholds()
            and self._has_valid_co_attainment_target()
            and (not self.state.busy)
        )
        self.drop_widget.setEnabled(not self.state.busy)
        self.drop_widget.set_submit_allowed(can_submit)
        self.threshold_l1_input.setEnabled(not self.state.busy)
        self.threshold_l2_input.setEnabled(not self.state.busy)
        self.threshold_l3_input.setEnabled(not self.state.busy)
        self.co_attainment_percent_input.setEnabled(not self.state.busy)
        self.co_attainment_level_input.setEnabled(not self.state.busy)
        self.drop_list.setEnabled(not self.state.busy)
        self.clear_button.setEnabled(has_files and (not self.state.busy))
        self.calculate_button.setEnabled(can_submit)
        for row in range(self.drop_list.count()):
            item = self.drop_list.item(row)
            widget = self.drop_list.itemWidget(item)
            if isinstance(widget, _COAnalysisFileItemWidget):
                widget.remove_btn.setEnabled(not self.state.busy)
        self._refresh_summary()

    def _read_attainment_thresholds(self) -> tuple[float, float, float]:
        return self._workflow_controller.read_attainment_thresholds()

    def _has_valid_attainment_thresholds(self) -> bool:
        return self._workflow_controller.has_valid_attainment_thresholds()

    def _read_co_attainment_target(self) -> tuple[float, int]:
        return self._workflow_controller.read_co_attainment_target()

    def _has_valid_co_attainment_target(self) -> bool:
        return self._workflow_controller.has_valid_co_attainment_target()

    def _notify_threshold_violation(self, *, force: bool) -> None:
        self._workflow_controller.notify_threshold_violation(force=force)

    def _show_threshold_validation_toast(self, *, message_key: str) -> None:
        show_threshold_validation_toast(
            self,
            message_key=message_key,
            title_key="coordinator.title",
            toast_fn=_show_toast_plain,
            translate=t,
        )

    def _on_threshold_value_changed(self, _value: float) -> None:
        self._workflow_controller.on_threshold_value_changed()
        self._refresh_ui()

    def _on_threshold_editing_finished(self) -> None:
        self._workflow_controller.on_threshold_editing_finished()
        self._refresh_ui()

    def _on_submit_requested(self) -> None:
        self._save_final_workbook()

    def _save_final_workbook(self) -> None:
        save_workbook_async(self, ns=_generate_workbook_namespace())

    def _on_files_dropped(self, dropped_files: list[str]) -> None:
        if dropped_files:
            self._remember_dialog_dir_safe(dropped_files[0])
        self._collect_valid_files(dropped_files)

    def _on_drop_list_reordered(self, ordered_paths: list[str]) -> None:
        key_to_path = {canonical_path_key(path): path for path in self._files}
        reordered: list[Path] = []
        for raw_path in ordered_paths:
            key = canonical_path_key(Path(raw_path))
            matched = key_to_path.get(key)
            if matched is not None:
                reordered.append(matched)
        if len(reordered) == len(self._files):
            self._files = reordered
            self._refresh_summary()

    def _browse_files(self) -> None:
        if self.state.busy:
            return
        selected_files, _ = QFileDialog.getOpenFileNames(
            self,
            t("co_analysis.dialog.select_files"),
            resolve_dialog_start_path(APP_NAME),
            t("instructor.dialog.filter.excel_open"),
        )
        if not selected_files:
            return
        self._remember_dialog_dir_safe(selected_files[0])
        self._collect_valid_files(selected_files)

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        self._runtime.remember_dialog_dir_safe(selected_path)

    def _setup_ui_logging(self) -> None:
        self._runtime.setup_ui_logging()

    def _append_user_log(self, message: str) -> None:
        self._runtime.append_user_log(message)

    def _rerender_user_log(self) -> None:
        _rerender_user_log_impl(self, ns=_messages_namespace())

    def _refresh_summary(self) -> None:
        count = len(self._files)
        self.drop_widget.set_summary_text_builder(lambda _count: t("coordinator.summary", count=count))
        # ManagedDropFileWidget enables this label using its own internal list length.
        # CO Analysis maintains file state externally, so force the visual state from module state.
        self.drop_widget.summary_label.setEnabled(count > 0)

    def _output_items(self) -> tuple[OutputItem, ...]:
        return tuple(
            OutputItem(label_key="coordinator.links.downloaded_output", path=str(path))
            for path in self._downloaded_outputs
        )

    def _collect_valid_files(self, candidate_paths: list[str]) -> None:
        collect_files_async(self, candidate_paths, ns=_collect_files_namespace())

    def _new_file_item_widget(
        self,
        file_path: str,
        *,
        parent: QWidget | None = None,
    ) -> _COAnalysisFileItemWidget:
        return _COAnalysisFileItemWidget(file_path, parent=parent)

    def _add_uploaded_paths(self, added_paths: list[Path]) -> None:
        for path in added_paths:
            self._files.append(path)
            path_text = str(path)
            if len(path_text) >= 2 and path_text[1] == ":":
                path_text = str(PureWindowsPath(path_text))
            item = QListWidgetItem()
            item.setToolTip(path_text)
            item.setData(Qt.ItemDataRole.UserRole, path_text)
            self.drop_list.addItem(item)
            row_widget = self._new_file_item_widget(path_text, parent=self.drop_list)
            row_widget.removed.connect(self._remove_file_by_path)
            item.setSizeHint(row_widget.sizeHint())
            self.drop_list.setItemWidget(item, row_widget)

    def _remove_file_by_path(self, file_path: str) -> None:
        remove_file_by_path(self, file_path, ns=_file_actions_namespace())

    def _clear_all(self) -> None:
        clear_all(self, ns=_file_actions_namespace())

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self._ui_engine.set_footer_visible(False)

    def get_shared_outputs_data(self) -> OutputPanelData:
        return OutputPanelData(items=self._output_items())

    def closeEvent(self, event) -> None:
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._cancel_token = None
        self._active_jobs.clear()
        if self._ui_log_handler is not None:
            self._logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)


