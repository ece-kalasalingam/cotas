"""CO Analysis module (UI-only scaffold)."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, cast

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from common.async_operation_runner import AsyncOperationRunner
from common.attainment_policy import (
    has_valid_attainment_thresholds as _has_valid_attainment_thresholds_policy,
)
from common.attainment_policy import (
    has_valid_co_attainment_percent as _has_valid_co_attainment_percent_policy,
)
from common.constants import (
    APP_NAME,
    CO_ANALYSIS_WORKFLOW_OPERATION_GENERATE_WORKBOOK,
    CO_ATTAINMENT_LEVEL_DEFAULT,
    CO_ATTAINMENT_PERCENT_DEFAULT,
    ID_COURSE_SETUP,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    LEVEL_1_THRESHOLD,
    LEVEL_2_THRESHOLD,
    LEVEL_3_THRESHOLD,
    MODULE_LEFT_PANE_CONTENT_MARGINS,
    MODULE_LEFT_PANE_LAYOUT_SPACING,
    MODULE_LEFT_PANE_WIDTH_OFFSET,
)
from common.drag_drop_file_widget import ManagedDropFileWidget
from common.error_catalog import resolve_validation_issue, validation_error_from_key
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.i18n import get_language, t
from common.jobs import CancellationToken
from common.module_messages import build_status_message as _build_status_message
from common.module_messages import (
    default_messages_namespace as _default_messages_namespace,
)
from common.module_messages import rerender_user_log as _rerender_user_log_impl
from common.module_runtime import ModuleRuntime
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.output_panel import OutputItem, OutputPanelData
from common.qt_jobs import run_in_background
from common.registry import (
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
)
from common.ui_stylings import GLOBAL_QPUSHBUTTON_MIN_WIDTH
from common.utils import (
    canonical_path_key,
    log_process_message,
    normalize,
    resolve_dialog_start_path,
    sanitize_filename_token,
)
from domain import BusyWorkflowState
from domain.template_strategy_router import (
    consume_marks_anomaly_warnings,
    extract_course_metadata_and_students_from_workbook_path,
    generate_workbook,
    resolve_template_id_from_workbook_path,
    validate_workbooks,
)
from services.gemini_cip_client import CipWorkerError, call_cip_worker

_LEFT_PANE_WIDTH = GLOBAL_QPUSHBUTTON_MIN_WIDTH + MODULE_LEFT_PANE_WIDTH_OFFSET
_LEFT_PANE_SCROLLBAR_GUTTER = 12
_LEFT_PANE_CONTENT_RIGHT_PADDING = 10
_LEFT_PANE_TEXT_RIGHT_SAFE_GAP = 28
_TAMIL_LANGUAGE_CODES = {"ta-in", "ta_in"}
_TAMIL_COMPACT_TEXT_STYLE = "font-size: 12px;"
_logger = logging.getLogger(__name__)


def _safe_cip_provider(payload: dict) -> str | None:
    try:
        return call_cip_worker(payload)
    except CipWorkerError:
        _logger.exception("CIP Worker call failed; Word report will be generated without CIP text.")
        return None
_DOWNLOAD_CO_DESCRIPTION_TEMPLATE_HREF = "download-co-description-template"
_WORD_FILE_FILTER = "Word Files (*.docx);;All Files (*)"


def _messages_namespace() -> dict[str, object]:
    """Messages namespace.
    
    Args:
        None.
    
    Returns:
        dict[str, object]: Return value.
    
    Raises:
        None.
    """
    return dict(_default_messages_namespace(translate=t))


def _build_i18n_message(text_key: str, *, kwargs: dict[str, object] | None = None, fallback: str | None = None) -> str:
    """Build i18n message.
    
    Args:
        text_key: Parameter value (str).
        kwargs: Parameter value (dict[str, object] | None).
        fallback: Parameter value (str | None).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    return _build_status_message(text_key, translate=t, kwargs=kwargs, fallback=fallback)


class _LogSink:
    def appendPlainText(self, _text: str) -> None:  # noqa: N802 - Qt-style name
        """Appendplaintext.
        
        Args:
            _text: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return

    def clear(self) -> None:
        """Clear.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        return


class COAnalysisModule(QWidget):
    status_changed = Signal(str)
    _THRESHOLD_VALIDATION_KEY = "co_analysis.thresholds.invalid_rule"
    _CO_ATTAINMENT_TARGET_VALIDATION_KEY = "co_analysis.co_attainment.invalid_percent"

    def __init__(self) -> None:
        """Init.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        super().__init__()
        self._files: list[Path] = []
        self._marks_files: list[Path] = []
        self._co_description_files: list[Path] = []
        self._downloaded_outputs: list[Path] = []
        self.state = BusyWorkflowState()
        self._threshold_violation_active = False
        self._logger = _logger
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._ui_log_handler: logging.Handler | None = None
        self._user_log_entries: list[dict[str, object]] = []
        self._pending_clear_count = 0
        self._syncing_validated_files = False
        self._async_runner = AsyncOperationRunner(self, run_async=run_in_background)
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
        """Build ui.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        left_content_layout = QHBoxLayout(left_content)
        left_margin, top_margin, right_margin, bottom_margin = MODULE_LEFT_PANE_CONTENT_MARGINS
        left_content_layout.setContentsMargins(
            left_margin,
            top_margin,
            right_margin + _LEFT_PANE_CONTENT_RIGHT_PADDING,
            bottom_margin,
        )
        left_content_layout.setSpacing(0)
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(0, 0, _LEFT_PANE_CONTENT_RIGHT_PADDING, 0)
        left_content_layout.addLayout(left_layout, 1)
        left_content_layout.addSpacing(_LEFT_PANE_SCROLLBAR_GUTTER)
        left_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
        left_scroll.setWidget(left_content)
        left_card_layout.addWidget(left_scroll, 1)
        left_card.setFixedWidth(_LEFT_PANE_WIDTH)
        top_layout.addWidget(left_card, 0)

        self.title_label = QLabel()
        self.title_label.setObjectName("coordinatorTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        self.hint_label = QLabel()
        self.hint_label.setObjectName("coordinatorHint")
        self.hint_label.setWordWrap(True)
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

        thresholds_layout = QVBoxLayout()
        thresholds_layout.setContentsMargins(0, 0, 0, 0)
        thresholds_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
        self.threshold_title_label = QLabel()
        self.threshold_title_label.setObjectName("coordinatorThresholdTitle")
        thresholds_layout.addWidget(self.threshold_title_label)
        self.threshold_description_label = QLabel()
        self.threshold_description_label.setWordWrap(True)
        self.threshold_description_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.threshold_description_label.setMaximumWidth(_LEFT_PANE_WIDTH - _LEFT_PANE_TEXT_RIGHT_SAFE_GAP)
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
        self.co_attainment_description_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.co_attainment_description_label.setMaximumWidth(_LEFT_PANE_WIDTH - _LEFT_PANE_TEXT_RIGHT_SAFE_GAP)
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

        self.generate_word_report_checkbox = QCheckBox()
        self.generate_word_report_checkbox.setObjectName("coAnalysisWordReportToggle")
        self.generate_word_report_checkbox.setStyleSheet(
            """
            QCheckBox#coAnalysisWordReportToggle,
            QCheckBox#coAnalysisWordReportToggle:hover,
            QCheckBox#coAnalysisWordReportToggle:focus,
            QCheckBox#coAnalysisWordReportToggle:pressed,
            QCheckBox#coAnalysisWordReportToggle:checked {
                text-decoration: none;
                border: none;
                outline: none;
            }
            """
        )
        self.generate_word_report_checkbox.setChecked(True)
        self.generate_word_report_checkbox.toggled.connect(lambda _checked: self._refresh_ui())
        thresholds_layout.addWidget(self.generate_word_report_checkbox)
        left_layout.addLayout(thresholds_layout)
        left_layout.addStretch(1)

        self.download_co_description_template_link = QLabel()
        self.download_co_description_template_link.setTextFormat(Qt.TextFormat.RichText)
        self.download_co_description_template_link.setTextInteractionFlags(
            Qt.TextInteractionFlag.LinksAccessibleByMouse | Qt.TextInteractionFlag.TextSelectableByMouse
        )
        self.download_co_description_template_link.setCursor(Qt.CursorShape.PointingHandCursor)
        self.download_co_description_template_link.setOpenExternalLinks(False)
        self.download_co_description_template_link.linkActivated.connect(
            self._on_download_co_description_template_link_activated
        )

        self.drop_widget = ManagedDropFileWidget(
            drop_mode="multiple",
            remove_fallback_text=t("co_analysis.file.remove_fallback"),
            open_file_tooltip=t("outputs.open_file"),
            open_folder_tooltip=t("outputs.open_folder"),
            remove_tooltip=t("co_analysis.file.remove_tooltip"),
        )
        self.drop_widget.set_summary_text_builder(lambda _count: t("co_analysis.summary", count=len(self._files)))
        self.drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.drop_widget.drop_list.setObjectName("coordinatorDropList")
        self.drop_widget.files_dropped.connect(self._on_files_dropped)
        self.drop_widget.files_rejected.connect(self._on_files_rejected)
        self.drop_widget.files_changed.connect(self._on_drop_widget_files_changed)
        self.drop_widget.browse_requested.connect(self._browse_files)
        self.drop_widget.clear_button.pressed.connect(self._on_clear_all_pressed)
        self.drop_widget.clear_button.clicked.connect(self._on_clear_all_clicked)
        self.drop_widget.submit_requested.connect(self._on_submit_requested)
        self.drop_widget.set_clear_button_text(t("co_analysis.clear_all"))
        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_pane)
        right_layout.addWidget(self.title_label)
        right_layout.addWidget(self.hint_label)
        right_layout.addWidget(self.download_co_description_template_link)
        right_layout.addWidget(self.drop_widget, 1)
        right_scroll = QScrollArea()
        right_scroll.setObjectName("coordinatorRightScroll")
        right_scroll.setFrameShape(QFrame.Shape.NoFrame)
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        right_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        right_scroll.viewport().setObjectName("coordinatorRightScrollViewport")
        right_scroll.setWidget(right_pane)
        top_layout.addWidget(right_scroll, 1)

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
        self.shortcut_save_output = QShortcut(QKeySequence(QKeySequence.StandardKey.Save), self)
        self.shortcut_save_output.activated.connect(self._on_submit_requested)

    def retranslate_ui(self) -> None:
        """Retranslate ui.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.title_label.setText(t("co_analysis.title"))
        self.hint_label.setText(t("co_analysis.drop_hint"))
        self.drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.drop_widget.retranslate_tooltips(
            t("outputs.open_file"),
            t("outputs.open_folder"),
            t("co_analysis.file.remove_tooltip"),
        )
        self.drop_widget.set_clear_button_text(t("co_analysis.clear_all"))
        self.drop_widget.set_summary_text_builder(lambda _count: t("co_analysis.summary", count=len(self._files)))
        self.drop_widget.set_submit_button_text(t("co_analysis.calculate"))
        self.threshold_title_label.setText(t("co_analysis.thresholds.title"))
        self.threshold_description_label.setText(t("co_analysis.thresholds.description"))
        self.threshold_l1_label.setText(t("co_analysis.thresholds.l1.label"))
        self.threshold_l2_label.setText(t("co_analysis.thresholds.l2.label"))
        self.threshold_l3_label.setText(t("co_analysis.thresholds.l3.label"))
        self.co_attainment_description_label.setText(t("co_analysis.co_attainment.description"))
        self.co_attainment_percent_label.setText(t("co_analysis.co_attainment.percent.label"))
        self.co_attainment_level_label.setText(t("co_analysis.co_attainment.level.label"))
        self.generate_word_report_checkbox.setText(t("co_analysis.generate_word_report.label"))
        self._apply_locale_text_density()
        self._set_download_co_description_template_link_enabled(not self.state.busy)
        self._refresh_summary()
        self._rerender_user_log()

    def _apply_locale_text_density(self) -> None:
        """Apply locale text density.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        is_tamil = get_language().strip().lower() in _TAMIL_LANGUAGE_CODES
        style = _TAMIL_COMPACT_TEXT_STYLE if is_tamil else ""
        self.hint_label.setStyleSheet(style)
        self.threshold_description_label.setStyleSheet(style)
        self.co_attainment_description_label.setStyleSheet(style)

    def _publish_status_key(self, text_key: str, **kwargs: Any) -> None:
        """Publish status key.
        
        Args:
            text_key: Parameter value (str).
            kwargs: Parameter value (Any).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._runtime.notify_message_key(text_key, channels=("status", "activity_log"), kwargs=kwargs)

    def _emit_uniform_validation_feedback(
        self,
        *,
        rejections: list[dict[str, object]],
        valid_count: int,
    ) -> None:
        """Emit the shared validation feedback contract (status/log + common toast summary)."""
        self._runtime.emit_validation_batch_feedback(
            rejections=rejections,
            valid_count=valid_count,
        )

    def _emit_uniform_workbook_generation_feedback(
        self,
        *,
        success_count: int,
        failed_count: int,
    ) -> None:
        """Emit the shared workbook-generation feedback contract (status/log + common toast summary)."""
        self._runtime.emit_workbook_generation_feedback(
            success_count=success_count,
            failed_count=failed_count,
        )

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        """Set busy.
        
        Args:
            busy: Parameter value (bool).
            job_id: Parameter value (str | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.state.set_busy(busy, job_id=job_id)
        host_window = self.window()
        set_switch = getattr(host_window, "set_language_switch_enabled", None)
        if callable(set_switch):
            set_switch(not busy)
        self._refresh_ui()

    def _setup_ui_logging(self) -> None:
        """Setup ui logging.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._runtime.setup_ui_logging()

    def _rerender_user_log(self) -> None:
        """Rerender user log.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        _rerender_user_log_impl(self, ns=_messages_namespace())

    def _refresh_ui(self) -> None:
        """Refresh ui.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        enabled = not self.state.busy
        has_marks_files = bool(self._marks_files)
        has_valid_co_description_selection = self._has_valid_co_description_selection()
        can_submit = (
            has_marks_files
            and has_valid_co_description_selection
            and self._has_valid_attainment_thresholds()
            and self._has_valid_co_attainment_target()
            and enabled
        )
        self.drop_widget.setEnabled(enabled)
        self._set_download_co_description_template_link_enabled(enabled)
        self.drop_widget.set_submit_allowed(can_submit)
        # Keep numeric policy controls enabled even while action flows are disabled.
        self.threshold_l1_input.setEnabled(True)
        self.threshold_l2_input.setEnabled(True)
        self.threshold_l3_input.setEnabled(True)
        self.co_attainment_percent_input.setEnabled(True)
        self.co_attainment_level_input.setEnabled(True)
        self.generate_word_report_checkbox.setEnabled(True)
        self.drop_list.setEnabled(enabled)
        self.clear_button.setEnabled(enabled and bool(self._files))
        self.calculate_button.setEnabled(can_submit)
        self._refresh_summary()

    def _set_download_co_description_template_link_enabled(self, enabled: bool) -> None:
        """Set download co description template link enabled.
        
        Args:
            enabled: Parameter value (bool).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        text = t("co_analysis.action.download_co_description_template")
        if enabled:
            self.download_co_description_template_link.setText(
                t(
                    "co_analysis.action.download_co_description_template_link_html",
                    href=_DOWNLOAD_CO_DESCRIPTION_TEMPLATE_HREF,
                    label=text,
                )
            )
        else:
            self.download_co_description_template_link.setText(text)
        self.download_co_description_template_link.setEnabled(enabled)

    def _on_download_co_description_template_link_activated(self, _href: str) -> None:
        """On download co description template link activated.
        
        Args:
            _href: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self.state.busy:
            return
        self._download_co_description_template_async()

    def _download_co_description_template_async(self) -> None:
        """Download co description template async.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self.state.busy:
            return
        start_dir = resolve_dialog_start_path(APP_NAME)
        default_path = str(Path(start_dir) / "CO_Description_Template.xlsx")
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            t("co_analysis.dialog.co_description_template.save_title"),
            default_path,
            t("instructor.dialog.filter.excel"),
        )
        if not output_path:
            return
        selected_output = Path(output_path)
        self._remember_dialog_dir_safe(str(selected_output))
        process_name = t("co_analysis.log.process.generate_co_description_template")
        token = CancellationToken()

        def _work() -> object:
            """Work.
            
            Args:
                None.
            
            Returns:
                object: Return value.
            
            Raises:
                None.
            """
            token.raise_if_cancelled()
            return generate_workbook(
                template_id=ID_COURSE_SETUP,
                output_path=selected_output,
                workbook_name=selected_output.name,
                workbook_kind="co_description_template",
                cancel_token=token,
            )

        def _on_success(result: object) -> None:
            """On success.
            
            Args:
                result: Parameter value (object).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            output_value = str(
                getattr(result, "workbook_path", None)
                or getattr(result, "output_path", None)
                or selected_output
            ).strip()
            if not output_value:
                self._handle_async_failure(
                    validation_error_from_key(
                        "common.validation_failed_invalid_data",
                        code="WORKBOOK_GENERATE_FAILED",
                        workbook_kind="co_description_template",
                    ),
                    process_name=process_name,
                )
                return
            result_path = Path(output_value)
            if all(canonical_path_key(path) != canonical_path_key(result_path) for path in self._downloaded_outputs):
                self._downloaded_outputs.append(result_path)
            self._publish_status_key("co_analysis.status.co_description_template_generated")
            log_process_message(
                process_name,
                logger=self._logger,
                success_message=f"{process_name} completed successfully.",
                user_success_message=_build_status_message(
                    "co_analysis.status.co_description_template_generated",
                    translate=t,
                    fallback=t("co_analysis.status.co_description_template_generated"),
                ),
            )
            self._emit_uniform_workbook_generation_feedback(
                success_count=1,
                failed_count=0,
            )

        def _on_failure(exc: Exception) -> None:
            """On failure.
            
            Args:
                exc: Parameter value (Exception).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self._handle_async_failure(
                exc,
                process_name=process_name,
                process_key="co_analysis.log.process.generate_co_description_template",
            )

        self._start_async_operation(
            token=token,
            job_id=None,
            work=_work,
            on_success=_on_success,
            on_failure=_on_failure,
        )

    def _read_attainment_thresholds(self) -> tuple[float, float, float]:
        """Read attainment thresholds.
        
        Args:
            None.
        
        Returns:
            tuple[float, float, float]: Return value.
        
        Raises:
            None.
        """
        return (
            float(self.threshold_l1_input.value()),
            float(self.threshold_l2_input.value()),
            float(self.threshold_l3_input.value()),
        )

    def _has_valid_attainment_thresholds(self) -> bool:
        """Has valid attainment thresholds.
        
        Args:
            None.
        
        Returns:
            bool: Return value.
        
        Raises:
            None.
        """
        l1, l2, l3 = self._read_attainment_thresholds()
        return _has_valid_attainment_thresholds_policy(l1, l2, l3)

    def _read_co_attainment_target(self) -> tuple[float, int]:
        """Read co attainment target.
        
        Args:
            None.
        
        Returns:
            tuple[float, int]: Return value.
        
        Raises:
            None.
        """
        raw_level = self.co_attainment_level_input.currentData()
        level = int(raw_level) if isinstance(raw_level, int) else 1
        return (
            float(self.co_attainment_percent_input.value()),
            level,
        )

    def _has_valid_co_attainment_target(self) -> bool:
        """Has valid co attainment target.
        
        Args:
            None.
        
        Returns:
            bool: Return value.
        
        Raises:
            None.
        """
        percent, level = self._read_co_attainment_target()
        return _has_valid_co_attainment_percent_policy(percent) and level >= 1

    def _show_threshold_validation_toast(self, *, message_key: str) -> None:
        """Show threshold validation toast.
        
        Args:
            message_key: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._runtime.notify_message_key(
            message_key,
            channels=("toast",),
            toast_title_key="co_analysis.title",
            toast_level="warning",
        )

    def _notify_threshold_violation(self, *, force: bool) -> None:
        """Notify threshold violation.
        
        Args:
            force: Parameter value (bool).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self._threshold_violation_active and not force:
            return
        if not self._has_valid_attainment_thresholds():
            key = self._THRESHOLD_VALIDATION_KEY
        else:
            key = self._CO_ATTAINMENT_TARGET_VALIDATION_KEY
        self._show_threshold_validation_toast(message_key=key)
        self._publish_status_key(key)
        self._threshold_violation_active = True

    def _on_threshold_value_changed(self, _value: float) -> None:
        """On threshold value changed.
        
        Args:
            _value: Parameter value (float).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self._has_valid_attainment_thresholds() and self._has_valid_co_attainment_target():
            self._threshold_violation_active = False
        self._refresh_ui()

    def _on_threshold_editing_finished(self) -> None:
        """On threshold editing finished.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self._has_valid_attainment_thresholds() and self._has_valid_co_attainment_target():
            self._threshold_violation_active = False
            self._refresh_ui()
            return
        self._notify_threshold_violation(force=False)
        self._refresh_ui()

    def _refresh_summary(self) -> None:
        """Refresh summary.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        count = len(self._files)
        self.drop_widget.set_summary_text_builder(lambda _count: t("co_analysis.summary", count=count))
        self.drop_widget.summary_label.setEnabled(count > 0)

    def _has_valid_co_description_selection(self) -> bool:
        """Return whether current CO-description selection satisfies submit policy."""
        if not self.generate_word_report_checkbox.isChecked():
            return True
        return len(self._co_description_files) == 1

    @staticmethod
    def _cohort_signature_from_metadata(metadata: dict[str, str]) -> tuple[str, str, str, str]:
        """Return normalized cohort signature from extracted metadata map."""
        return (
            normalize(metadata.get(normalize(COURSE_METADATA_COURSE_CODE_KEY), "")),
            normalize(metadata.get(normalize(COURSE_METADATA_SEMESTER_KEY), "")),
            normalize(metadata.get(normalize(COURSE_METADATA_ACADEMIC_YEAR_KEY), "")),
            normalize(metadata.get(normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY), "")),
        )

    @staticmethod
    def _consume_marks_anomaly_warnings(template_id: str) -> list[str]:
        """Consume marks anomaly warnings.
        
        Args:
            template_id: Parameter value (str).
        
        Returns:
            list[str]: Return value.
        
        Raises:
            None.
        """
        return consume_marks_anomaly_warnings(template_id)

    @staticmethod
    def _resolve_template_id(workbook_path: str | Path) -> str:
        """Resolve template id from workbook system-hash metadata."""
        return resolve_template_id_from_workbook_path(workbook_path)

    @classmethod
    def _resolve_single_template_id(
        cls,
        workbook_paths: list[str],
        *,
        mismatch_code: str,
    ) -> str:
        """Resolve one template id from a workbook list; fail on mixed templates."""
        template_ids = {
            cls._resolve_template_id(path)
            for path in workbook_paths
            if str(path).strip()
        }
        if not template_ids:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_SELECTION_EMPTY",
            )
        if len(template_ids) > 1:
            raise validation_error_from_key(
                "common.validation_failed_invalid_data",
                code=mismatch_code,
                count=len(template_ids),
            )
        return next(iter(template_ids))

    @staticmethod
    def _raise_first_validation_issue(*, result: dict[str, object], workbook_path: str) -> None:
        """Raise first validation issue.
        
        Args:
            result: Parameter value (dict[str, object]).
            workbook_path: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        workbook_key = canonical_path_key(workbook_path)
        valid_paths_raw = result.get("valid_paths", [])
        valid_paths = [str(path) for path in valid_paths_raw] if isinstance(valid_paths_raw, list) else []
        valid_keys = {canonical_path_key(path) for path in valid_paths}
        if workbook_key in valid_keys:
            return

        rejections_raw = result.get("rejections", [])
        rejection_items = [item for item in rejections_raw if isinstance(item, dict)] if isinstance(rejections_raw, list) else []
        issue = next(
            (
                dict(item.get("issue", {}))
                for item in rejection_items
                if canonical_path_key(str(item.get("path", "")).strip()) == workbook_key and isinstance(item.get("issue"), dict)
            ),
            None,
        )
        if issue is not None:
            code = str(issue.get("code", "VALIDATION_ERROR")).strip() or "VALIDATION_ERROR"
            message = str(issue.get("message", code)).strip() or code
            context = issue.get("context", {})
            context_dict = dict(context) if isinstance(context, dict) else {}
            raise ValidationError(message, code=code, context=context_dict)
        raise validation_error_from_key(
            "instructor.validation.workbook_open_failed",
            code="WORKBOOK_OPEN_FAILED",
            workbook=workbook_path,
        )

    def _on_drop_widget_files_changed(self, file_paths: list[str]) -> None:
        """On drop widget files changed.
        
        Args:
            file_paths: Parameter value (list[str]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self._syncing_validated_files:
            self._files = [Path(path) for path in file_paths if path]
            self._refresh_ui()
            return

        candidate_paths = [str(path).strip() for path in file_paths if str(path).strip()]
        process_name = t("co_analysis.status.processing_started")
        require_single_co_description = self.generate_word_report_checkbox.isChecked()
        token = CancellationToken()

        def _work() -> dict[str, object]:
            """Work.
            
            Args:
                None.
            
            Returns:
                dict[str, object]: Return value.
            
            Raises:
                None.
            """
            accepted_marks_paths: list[str] = []
            accepted_co_description_paths: list[str] = []
            accepted_paths: list[str] = []
            rejected_items: list[dict[str, object]] = []
            co_description_accepted_paths: list[str] = []
            evaluated_template_ids: set[str] = set()

            def _is_path_valid(*, result: dict[str, object], key: str) -> bool:
                valid_paths_raw = result.get("valid_paths", [])
                if not isinstance(valid_paths_raw, list):
                    return False
                return any(canonical_path_key(str(path)) == key for path in valid_paths_raw)

            def _rejection_for_path(*, result: dict[str, object], key: str) -> dict[str, object] | None:
                rejections_raw = result.get("rejections", [])
                if not isinstance(rejections_raw, list):
                    return None
                for item in rejections_raw:
                    if not isinstance(item, dict):
                        continue
                    item_path = str(item.get("path", "")).strip()
                    if item_path and canonical_path_key(item_path) == key:
                        return item
                return None

            for path in candidate_paths:
                key = canonical_path_key(path)
                template_id = self._resolve_template_id(path)
                evaluated_template_ids.add(template_id)
                marks_result = validate_workbooks(
                    template_id=template_id,
                    workbook_paths=[path],
                    workbook_kind="marks_template",
                )
                co_description_result = validate_workbooks(
                    template_id=template_id,
                    workbook_paths=[path],
                    workbook_kind="co_description",
                )
                is_marks_valid = _is_path_valid(result=marks_result, key=key)
                is_co_description_valid = _is_path_valid(result=co_description_result, key=key)
                if is_co_description_valid:
                    accepted_co_description_paths.append(path)
                    co_description_accepted_paths.append(path)
                    accepted_paths.append(path)
                    continue
                if is_marks_valid:
                    accepted_marks_paths.append(path)
                    accepted_paths.append(path)
                    continue
                rejection = (
                    _rejection_for_path(result=marks_result, key=key)
                    or _rejection_for_path(result=co_description_result, key=key)
                )
                if isinstance(rejection, dict):
                    rejected_items.append(rejection)
                    continue
                rejection_payload = self._build_validation_rejection_payload(
                    path_text=path,
                    exc=validation_error_from_key(
                        "validation.workbook.open_failed",
                        code="WORKBOOK_OPEN_FAILED",
                        workbook=path,
                    ),
                )
                rejected_items.append(
                    {
                        "path": path,
                        "reason_kind": "invalid",
                        "issue": rejection_payload["issue"],
                    }
                )

            if require_single_co_description and len(accepted_co_description_paths) > 1:
                accepted_paths = [
                    path
                    for path in accepted_paths
                    if canonical_path_key(path)
                    not in {canonical_path_key(item) for item in accepted_co_description_paths}
                ]
                accepted_co_description_paths = []
                for path in candidate_paths:
                    if canonical_path_key(path) not in {
                        canonical_path_key(item) for item in co_description_accepted_paths
                    }:
                        continue
                    resolved = resolve_validation_issue(
                        "CO_DESCRIPTION_SELECTION_INVALID",
                        context={
                            "workbook": path,
                            "count": len(co_description_accepted_paths),
                        },
                        fallback_message=(
                            "Exactly one CO description workbook is required when "
                            "Course Coordinator report generation is enabled."
                        ),
                    )
                    rejected_items.append(
                        {
                            "path": path,
                            "reason_kind": "invalid",
                            "issue": {
                                "code": resolved.code,
                                "category": resolved.category,
                                "severity": resolved.severity,
                                "translation_key": resolved.translation_key,
                                "message": resolved.message,
                                "context": dict(resolved.context),
                            },
                        }
                    )

            anomaly_warnings: list[str] = []
            for template_id in sorted(evaluated_template_ids):
                anomaly_warnings.extend(self._consume_marks_anomaly_warnings(template_id))
            return {
                "accepted_marks_paths": accepted_marks_paths,
                "accepted_co_description_paths": accepted_co_description_paths,
                "accepted_paths": accepted_paths,
                "rejected_items": rejected_items,
                "anomaly_warnings": anomaly_warnings,
                "original_paths": candidate_paths,
            }

        def _on_success(result: object) -> None:
            """On success.
            
            Args:
                result: Parameter value (object).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            data = result if isinstance(result, dict) else {}
            accepted_marks_paths = [
                str(path)
                for path in cast(list[object], data.get("accepted_marks_paths", []))
                if str(path).strip()
            ]
            accepted_co_description_paths = [
                str(path)
                for path in cast(list[object], data.get("accepted_co_description_paths", []))
                if str(path).strip()
            ]
            accepted_paths = [
                str(path) for path in cast(list[object], data.get("accepted_paths", [])) if str(path).strip()
            ]
            rejected_items = [
                item
                for item in cast(list[object], data.get("rejected_items", []))
                if isinstance(item, dict)
            ]
            anomaly_warnings = [str(item) for item in cast(list[object], data.get("anomaly_warnings", [])) if str(item).strip()]
            original_paths = [str(path) for path in cast(list[object], data.get("original_paths", [])) if str(path).strip()]

            self._marks_files = [Path(path) for path in accepted_marks_paths]
            self._co_description_files = [Path(path) for path in accepted_co_description_paths]

            if accepted_paths != original_paths:
                self._syncing_validated_files = True
                try:
                    self.drop_widget.set_files(accepted_paths)
                finally:
                    self._syncing_validated_files = False

            has_rejections = bool(rejected_items)
            has_accepted = bool(accepted_paths)
            if has_accepted and not has_rejections:
                self.drop_widget.set_validation_state("success")
            elif has_accepted and has_rejections:
                self.drop_widget.set_validation_state("warning")
            elif has_rejections:
                self.drop_widget.set_validation_state("error")
            else:
                self.drop_widget.set_validation_state("neutral")

            if rejected_items:
                self._emit_uniform_validation_feedback(
                    rejections=cast(list[dict[str, object]], rejected_items),
                    valid_count=len(accepted_paths),
                )
            elif accepted_paths:
                self._emit_uniform_validation_feedback(
                    rejections=[],
                    valid_count=len(accepted_paths),
                )

            if anomaly_warnings:
                self._publish_status_key("co_analysis.status.validation_warnings", count=len(anomaly_warnings))
                displayed = anomaly_warnings[:12]
                for warning in displayed:
                    self._publish_status_key("co_analysis.status.validation_warning_line", warning=warning)
                hidden_count = len(anomaly_warnings) - len(displayed)
                if hidden_count > 0:
                    self._publish_status_key("co_analysis.status.validation_warning_more", count=hidden_count)

            self._files = [Path(path) for path in accepted_paths]
            self._refresh_ui()

        def _on_failure(exc: Exception) -> None:
            """On failure.
            
            Args:
                exc: Parameter value (Exception).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self._handle_async_failure(exc, process_name=process_name)

        self._start_async_operation(
            token=token,
            job_id=None,
            work=_work,
            on_success=_on_success,
            on_failure=_on_failure,
        )

    @staticmethod
    def _build_validation_rejection_payload(*, path_text: str, exc: Exception) -> dict[str, object]:
        """Build validation rejection payload.
        
        Args:
            path_text: Parameter value (str).
            exc: Parameter value (Exception).
        
        Returns:
            dict[str, object]: Return value.
        
        Raises:
            None.
        """
        code = str(getattr(exc, "code", type(exc).__name__) or "VALIDATION_ERROR").strip() or "VALIDATION_ERROR"
        raw_context = getattr(exc, "context", {})
        context = dict(raw_context) if isinstance(raw_context, dict) else {}
        if "workbook" not in context:
            context["workbook"] = path_text
        fallback_reason = str(exc).strip()
        if fallback_reason.upper() == code.upper():
            fallback_reason = ""
        resolved = resolve_validation_issue(code, context=context, fallback_message=fallback_reason)
        issue_payload: dict[str, object] = {
            "code": resolved.code,
            "category": resolved.category,
            "severity": resolved.severity,
            "translation_key": resolved.translation_key,
            "message": resolved.message,
            "context": dict(resolved.context),
        }
        return {"path": path_text, "issue": issue_payload}

    @staticmethod
    def _compact_context_text(context: object) -> str:
        """Compact context text.
        
        Args:
            context: Parameter value (object).
        
        Returns:
            str: Return value.
        
        Raises:
            None.
        """
        if not isinstance(context, dict):
            return ""
        fields = ("sheet_name", "cell", "row", "range", "column")
        parts: list[str] = []
        for key in fields:
            value = context.get(key)
            token = str(value).strip() if value is not None else ""
            if token:
                parts.append(f"{key}={token}")
        return ", ".join(parts)

    def _on_files_dropped(self, dropped_files: list[str]) -> None:
        """On files dropped.
        
        Args:
            dropped_files: Parameter value (list[str]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        dropped_count = len([path for path in dropped_files if path])
        if dropped_count <= 0:
            return
        self.drop_widget.set_validation_state("info")
        first_path = next((value for value in dropped_files if value), "")
        if first_path:
            self._remember_dialog_dir_safe(first_path)

    def _on_files_rejected(self, rejected_files: list[str]) -> None:
        """On files rejected.
        
        Args:
            rejected_files: Parameter value (list[str]).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        rejected_count = len([path for path in rejected_files if path])
        if rejected_count <= 0:
            return
        self._publish_status_key("co_analysis.status.ignored", count=rejected_count)
        self.drop_widget.set_validation_state("warning")

    def _browse_files(self) -> None:
        """Browse files.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
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
        self.drop_widget.add_files(selected_files, emit_drop=True)

    def _on_clear_all_pressed(self) -> None:
        """On clear all pressed.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._pending_clear_count = len(self._files)

    def _on_clear_all_clicked(self) -> None:
        """On clear all clicked.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        cleared_count = self._pending_clear_count
        self._pending_clear_count = 0
        if cleared_count <= 0:
            return
        self._publish_status_key("co_analysis.status.cleared", count=cleared_count)
        self.drop_widget.set_validation_state("neutral")

    def _on_submit_requested(self) -> None:
        """On submit requested.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self.state.busy:
            return
        if not self._marks_files:
            return
        if not self._has_valid_attainment_thresholds() or not self._has_valid_co_attainment_target():
            self._notify_threshold_violation(force=False)
            return
        if not self._has_valid_co_description_selection():
            return
        self._prepare_co_analysis_async()

    def _prepare_co_analysis_async(self) -> None:
        """Prepare co analysis async.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self.state.busy or not self._marks_files:
            return

        output_dir = QFileDialog.getExistingDirectory(
            self,
            t("co_analysis.calculate"),
            resolve_dialog_start_path(APP_NAME),
        )
        if not output_dir:
            return
        self._remember_dialog_dir_safe(output_dir)

        first_metadata = extract_course_metadata_and_students_from_workbook_path(self._marks_files[0])[1]
        course_code = sanitize_filename_token(
            first_metadata.get(normalize(COURSE_METADATA_COURSE_CODE_KEY), "")
        )
        academic_year = sanitize_filename_token(
            first_metadata.get(normalize(COURSE_METADATA_ACADEMIC_YEAR_KEY), "")
        )
        prefix = f"{course_code}_{academic_year}_" if course_code and academic_year else ""
        output_path = Path(output_dir) / f"{prefix}CO_Analysis.xlsx"
        should_generate_word_report = self.generate_word_report_checkbox.isChecked()
        word_output_path = Path(output_dir) / f"{prefix}CO_Analysis_Report.docx"
        if output_path.exists():
            replacement_path, _ = QFileDialog.getSaveFileName(
                self,
                t("co_analysis.calculate"),
                resolve_dialog_start_path(APP_NAME, output_path.name),
                t("instructor.dialog.filter.excel"),
            )
            if not replacement_path:
                return
            output_path = Path(replacement_path)
            word_output_path = output_path.with_name(f"{prefix}CO_Analysis_Report.docx")
        if should_generate_word_report and word_output_path.exists():
            replacement_word_path, _ = QFileDialog.getSaveFileName(
                self,
                t("co_analysis.calculate"),
                resolve_dialog_start_path(APP_NAME, word_output_path.name),
                _WORD_FILE_FILTER,
            )
            if not replacement_word_path:
                return
            word_output_path = Path(replacement_word_path)

        token = CancellationToken()
        thresholds = self._read_attainment_thresholds()
        co_attainment_percent, co_attainment_level = self._read_co_attainment_target()
        process_name = CO_ANALYSIS_WORKFLOW_OPERATION_GENERATE_WORKBOOK
        self._publish_status_key("co_analysis.status.processing_started")

        def _work() -> object:
            """Work.
            
            Args:
                None.
            
            Returns:
                object: Return value.
            
            Raises:
                None.
            """
            token.raise_if_cancelled()
            marks_paths = [str(path) for path in self._marks_files]
            template_id = self._resolve_single_template_id(
                marks_paths,
                mismatch_code="CO_ANALYSIS_MARKS_TEMPLATE_MISMATCH",
            )
            validation_result = validate_workbooks(
                template_id=template_id,
                workbook_paths=marks_paths,
                workbook_kind="marks_template",
            )
            for workbook_path in marks_paths:
                self._raise_first_validation_issue(result=validation_result, workbook_path=workbook_path)
            _ = self._consume_marks_anomaly_warnings(template_id)

            co_description_path: str | None = None
            if should_generate_word_report:
                if len(self._co_description_files) != 1:
                    raise validation_error_from_key(
                        "common.validation_failed_invalid_data",
                        code="CO_DESCRIPTION_SELECTION_INVALID",
                        count=len(self._co_description_files),
                    )
                co_description_path = str(self._co_description_files[0])
                co_description_template_id = self._resolve_template_id(co_description_path)
                if normalize(co_description_template_id) != normalize(template_id):
                    raise validation_error_from_key(
                        "common.validation_failed_invalid_data",
                        code="CO_DESCRIPTION_MARKS_TEMPLATE_MISMATCH",
                        template_id=co_description_template_id,
                        expected_template_id=template_id,
                    )
                co_description_validation_result = validate_workbooks(
                    template_id=co_description_template_id,
                    workbook_paths=[co_description_path],
                    workbook_kind="co_description",
                )
                self._raise_first_validation_issue(
                    result=co_description_validation_result,
                    workbook_path=co_description_path,
                )

                marks_metadata = extract_course_metadata_and_students_from_workbook_path(
                    self._marks_files[0]
                )[1]
                co_description_metadata = extract_course_metadata_and_students_from_workbook_path(
                    co_description_path
                )[1]
                marks_signature = self._cohort_signature_from_metadata(marks_metadata)
                co_description_signature = self._cohort_signature_from_metadata(co_description_metadata)
                if marks_signature != co_description_signature:
                    mismatch_fields: list[str] = []
                    if marks_signature[0] != co_description_signature[0]:
                        mismatch_fields.append(COURSE_METADATA_COURSE_CODE_KEY)
                    if marks_signature[1] != co_description_signature[1]:
                        mismatch_fields.append(COURSE_METADATA_SEMESTER_KEY)
                    if marks_signature[2] != co_description_signature[2]:
                        mismatch_fields.append(COURSE_METADATA_ACADEMIC_YEAR_KEY)
                    if marks_signature[3] != co_description_signature[3]:
                        mismatch_fields.append(COURSE_METADATA_TOTAL_OUTCOMES_KEY)
                    raise validation_error_from_key(
                        "common.validation_failed_invalid_data",
                        code="CO_DESCRIPTION_MARKS_COHORT_MISMATCH",
                        fields=", ".join(mismatch_fields) if mismatch_fields else "cohort",
                    )

            return generate_workbook(
                template_id=template_id,
                output_path=output_path,
                workbook_name=output_path.name,
                workbook_kind="co_attainment",
                cancel_token=token,
                context={
                    "source_paths": [str(path) for path in self._marks_files],
                    "co_description_path": co_description_path,
                    "thresholds": tuple(thresholds),
                    "co_attainment_percent": co_attainment_percent,
                    "co_attainment_level": co_attainment_level,
                    "generate_word_report": should_generate_word_report,
                    "word_output_path": str(word_output_path),
                    "cip_text_provider": _safe_cip_provider if should_generate_word_report else None,
                },
            )

        def _on_success(result: object) -> None:
            """On success.
            
            Args:
                result: Parameter value (object).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            result_path = Path(
                str(getattr(result, "output_path", None) or getattr(result, "workbook_path", None) or output_path)
            )
            generated_excel_count = 0
            failed_word_count = 0
            if all(canonical_path_key(path) != canonical_path_key(result_path) for path in self._downloaded_outputs):
                self._downloaded_outputs.append(result_path)
            generated_excel_count = 1
            self._publish_status_key(
                "co_analysis.status.output_generated_excel",
                path=str(result_path),
            )
            word_report_path_value = str(
                getattr(result, "word_report_path", None)
                or getattr(result, "report_output_path", None)
                or ""
            ).strip()
            word_generated = False
            if word_report_path_value:
                word_report_path = Path(word_report_path_value)
                if all(canonical_path_key(path) != canonical_path_key(word_report_path) for path in self._downloaded_outputs):
                    self._downloaded_outputs.append(word_report_path)
                word_generated = True
                self._publish_status_key(
                    "co_analysis.status.output_generated_word",
                    path=str(word_report_path),
                )
                self._runtime.notify_message_key(
                    "co_analysis.status.word_report_generated",
                    channels=("status", "activity_log"),
                )
            word_report_error_key = str(getattr(result, "word_report_error_key", None) or "").strip()
            if word_report_error_key and not word_generated:
                failed_word_count = 1
                self._runtime.notify_message_key(
                    word_report_error_key,
                    channels=("status", "activity_log"),
                )
            elif should_generate_word_report and not word_generated:
                failed_word_count = 1
                self._runtime.notify_message_key(
                    "co_analysis.status.word_report_not_generated",
                    channels=("status", "activity_log"),
                )
            self._remember_dialog_dir_safe(str(result_path))
            self._publish_status_key("co_analysis.status.calculate_completed")
            log_process_message(
                process_name,
                logger=self._logger,
                success_message=(
                    "saving co analysis workbooks completed successfully. "
                    f"output_dir={output_dir}, generated=1, "
                    f"thresholds=({thresholds[0]:g},{thresholds[1]:g},{thresholds[2]:g}), "
                    f"co_at_target=({co_attainment_percent:g},L{co_attainment_level})"
                ),
                user_success_message=_build_i18n_message(
                    "co_analysis.status.calculate_completed",
                    fallback=t("co_analysis.status.calculate_completed"),
                ),
            )
            self._emit_uniform_workbook_generation_feedback(
                success_count=generated_excel_count,
                failed_count=failed_word_count,
            )

        def _on_failure(exc: Exception) -> None:
            """On failure.
            
            Args:
                exc: Parameter value (Exception).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self._handle_async_failure(exc, process_name=process_name)

        self._start_async_operation(
            token=token,
            job_id=None,
            work=_work,
            on_success=_on_success,
            on_failure=_on_failure,
        )

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        """Remember dialog dir safe.
        
        Args:
            selected_path: Parameter value (str).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._runtime.remember_dialog_dir_safe(selected_path)

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
        """Start async operation.
        
        Args:
            token: Parameter value (CancellationToken).
            job_id: Parameter value (str | None).
            work: Parameter value.
            on_success: Parameter value.
            on_failure: Parameter value.
            on_finally: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._runtime.set_async_runner(self._async_runner)
        self._runtime.start_async_operation(
            token=token,
            job_id=job_id,
            work=work,
            on_success=on_success,
            on_failure=on_failure,
            on_finally=on_finally,
        )

    def _handle_async_failure(self, exc: Exception, *, process_name: str, process_key: str | None = None) -> None:
        """Handle async failure.
        
        Args:
            exc: Parameter value (Exception).
            process_name: Parameter value (str).
            process_key: Parameter value (str | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self.drop_widget.set_validation_state("error")
        process_ref: object = process_name
        if isinstance(process_key, str) and process_key.strip():
            process_ref = {"__t_key__": process_key}
        error_payload = _build_status_message(
            "common.error_while_process",
            translate=t,
            kwargs={"process": process_ref},
            fallback=t("common.error_while_process", process=process_name),
        )
        if isinstance(exc, JobCancelledError):
            return
        if isinstance(exc, ValidationError):
            self._runtime.notify_message(
                str(exc),
                channels=("toast",),
                toast_title=t("co_analysis.title"),
                toast_level="error",
            )
            resolved = resolve_validation_issue(
                str(getattr(exc, "code", "VALIDATION_ERROR")),
                context=(getattr(exc, "context", {}) or {}),
                fallback_message=str(exc),
            )
            validation_payload = _build_status_message(
                resolved.translation_key,
                translate=t,
                kwargs=resolved.context,
                fallback=resolved.message or str(exc),
            )
            log_process_message(
                process_name,
                logger=self._logger,
                error=exc,
                user_error_message=error_payload,
                user_validation_message=validation_payload,
            )
            return
        elif isinstance(exc, AppSystemError):
            self._runtime.notify_message(
                str(exc),
                channels=("toast",),
                toast_title=t("co_analysis.title"),
                toast_level="error",
            )
            log_process_message(
                process_name,
                logger=self._logger,
                error=exc,
                user_error_message=error_payload,
            )
            return
        else:
            self._runtime.notify_message_key(
                "co_analysis.msg.failed_to_generate_co_description_template",
                channels=("toast",),
                toast_title_key="co_analysis.title",
                toast_level="error",
            )
        log_process_message(
            process_name,
            logger=self._logger,
            error=exc,
            user_error_message=error_payload,
        )

    def _output_items(self) -> tuple[OutputItem, ...]:
        """Output items.
        
        Args:
            None.
        
        Returns:
            tuple[OutputItem, ...]: Return value.
        
        Raises:
            None.
        """
        return tuple(
            OutputItem(
                label_key=(
                    "co_analysis.links.generated_word_output"
                    if str(path.suffix).lower() == ".docx"
                    else "co_analysis.links.generated_excel_output"
                ),
                path=str(path),
            )
            for path in self._downloaded_outputs
        )

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        """Set shared activity log mode.
        
        Args:
            enabled: Parameter value (bool).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        del enabled
        self._ui_engine.set_footer_visible(False)

    def get_shared_outputs_data(self) -> OutputPanelData:
        """Get shared outputs data.
        
        Args:
            None.
        
        Returns:
            OutputPanelData: Return value.
        
        Raises:
            None.
        """
        return OutputPanelData(items=self._output_items())

    def closeEvent(self, event) -> None:
        """Closeevent.
        
        Args:
            event: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        self._ui_engine.set_footer_visible(False)
        if self._ui_log_handler is not None:
            self._logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)


__all__ = ["COAnalysisModule"]
