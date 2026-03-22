"""Course Instructor CO module UI (list-based step navigation)."""

from __future__ import annotations

import importlib
import logging
from collections.abc import Callable
from pathlib import Path
from typing import cast

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from common.async_operation_runner import AsyncOperationRunner
from common.constants import (
    APP_NAME,
    ID_COURSE_SETUP,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    MODULE_LEFT_PANE_CONTENT_MARGINS,
    MODULE_LEFT_PANE_LAYOUT_SPACING,
    MODULE_LEFT_PANE_WIDTH_OFFSET,
)
from common.drag_drop_file_widget import ManagedDropFileWidget
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.module_messages import rerender_user_log as _rerender_user_log_impl
from common.module_runtime import ModuleRuntime
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.output_panel import OutputItem, OutputPanelData
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
from common.ui_stylings import (
    GLOBAL_QPUSHBUTTON_MIN_WIDTH,
)
from common.utils import (
    emit_user_status,
    log_process_message,
    remember_dialog_dir,
    resolve_dialog_start_path,
)
from domain import InstructorWorkflowState
from modules.instructor import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)
from modules.instructor.messages import (
    localized_log_messages,
    show_step_success_toast,
    show_system_error_toast,
    show_validation_error_toast,
)
from modules.instructor.steps.shared_workbook_ops import (
    atomic_copy_file as _shared_atomic_copy_file,
)
from modules.instructor.steps.shared_workbook_ops import (
    build_final_report_default_name,
    build_marks_template_default_name,
)
from modules.instructor.steps.step1_course_details_template import (
    download_course_template_async,
)
from modules.instructor.steps.step2_course_details_and_marks_template import (
    prepare_marks_template_async,
    upload_course_details_async,
    upload_course_details_from_paths_async,
)
from modules.instructor.steps.step2_filled_marks_and_final_report import (
    generate_final_report_async,
    generate_final_reports_from_paths_async,
    upload_filled_marks_async,
)
from modules.instructor.validators.step2_filled_marks_validator import (
    consume_last_filled_marks_anomaly_warnings,
    filled_marks_manifest_validators,
    validate_filled_marks_manifest_schema_by_template,
    validate_uploaded_filled_marks_workbook,
)
from modules.instructor.workflow_controller import InstructorWorkflowController
from services import InstructorWorkflowService

_logger = logging.getLogger(__name__)
shutil = importlib.import_module("shutil")
_ENABLE_SECOND_ROW_ACTIONS = True
_LEFT_PANE_WIDTH = GLOBAL_QPUSHBUTTON_MIN_WIDTH + MODULE_LEFT_PANE_WIDTH_OFFSET
_DOWNLOAD_COURSE_TEMPLATE_HREF = "download-course-template"

# Step handlers receive `ns=globals()`, so these names must stay module-visible.
_STEP_RUNTIME_GLOBALS = (
    QFileDialog,
    ID_COURSE_SETUP,
    AppSystemError,
    JobCancelledError,
    ValidationError,
    log_process_message,
    remember_dialog_dir,
    resolve_dialog_start_path,
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
    build_i18n_log_message,
)

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


def _publish_status(target: object, message: str) -> None:
    cast("InstructorModule", target)._publish_status(message)

def _publish_status_key(target: object, text_key: str, **kwargs: object) -> None:
    publish_key = getattr(target, "_publish_status_key", None)
    if callable(publish_key):
        publish_key(text_key, **kwargs)
        return
    publish_plain = getattr(target, "_publish_status", None)
    if callable(publish_plain):
        publish_plain(t(text_key, **kwargs))


def _start_async_operation(
    target: object,
    *,
    token: CancellationToken,
    job_id: str | None,
    work,
    on_success,
    on_failure,
) -> None:
    cast("InstructorModule", target)._start_async_operation(
        token=token,
        job_id=job_id,
        work=work,
        on_success=on_success,
        on_failure=on_failure,
    )


def _localized_log_messages(process_key: str) -> tuple[str, str]:
    return localized_log_messages(process_key)


def _build_marks_template_default_name(course_details_path: str | None) -> str:
    return build_marks_template_default_name(course_details_path)


def _build_final_report_default_name(filled_marks_path: str | None) -> str:
    return build_final_report_default_name(filled_marks_path)


def _atomic_copy_file(source_path: str | Path, output_path: str | Path) -> Path:
    return _shared_atomic_copy_file(source_path, output_path, logger=_logger)


def _validate_uploaded_filled_marks_workbook(workbook_path: str | Path) -> None:
    validate_uploaded_filled_marks_workbook(workbook_path)


def _consume_last_filled_marks_anomaly_warnings() -> list[str]:
    return consume_last_filled_marks_anomaly_warnings()


def _validate_filled_marks_manifest_schema_by_template(
    workbook: object,
    manifest: object,
    *,
    template_id: str,
) -> None:
    validate_filled_marks_manifest_schema_by_template(
        workbook,
        manifest,
        template_id=template_id,
    )


def _filled_marks_manifest_validators() -> dict[str, Callable[[object, object], None]]:
    return filled_marks_manifest_validators()


class _LogSink:
    def appendPlainText(self, _text: str) -> None:  # noqa: N802 - Qt-style name
        return

    def clear(self) -> None:
        return


class InstructorModule(QWidget):
    """Simple wizard-like UI for CO score workflow."""

    status_changed = Signal(str)
    WORKFLOW_STEPS = (1, 2)
    RAIL_LINKS = (
        ("instructor.links.course_details_generated", "step1_path"),
        ("instructor.links.marks_template_generated", "marks_template_path"),
        ("instructor.links.final_co_report_generated", "final_report_path"),
    )
    RAIL_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"

    STEP_TITLE_KEYS = {
        1: "instructor.step1.title",
        2: "instructor.step2.title",
    }

    STEP_DESC_KEYS = {
        1: "instructor.step1.desc",
    }

    PATH_ATTRS = {
        1: "marks_template_path",
        2: "final_report_path",
    }

    DONE_ATTRS = {
        1: "marks_template_done",
        2: "final_report_done",
    }

    OUTDATED_ATTRS = {
        2: "final_report_outdated",
    }

    ACTION_DEFAULT_KEYS = {
        1: "instructor.action.step2.default",
        2: "instructor.action.step2.generate.default",
    }

    def __init__(self):
        super().__init__()
        self.current_step = 1

        self.step1_path: str | None = None
        self.marks_template_path: str | None = None
        self.marks_template_paths: list[str] = []
        self.step2_course_details_path: str | None = None
        self.step1_course_details_paths: list[str] = []
        self.filled_marks_path: str | None = None
        self.filled_marks_paths: list[str] = []
        self.final_report_path: str | None = None
        self.final_report_paths: list[str] = []

        self.step1_done = False
        self.marks_template_done = False
        self.filled_marks_done = False
        self.final_report_done = False
        self.step2_upload_ready = False

        self.filled_marks_outdated = False  # Filled marks
        self.final_report_outdated = False  # Final report

        self.state = InstructorWorkflowState()
        self._workflow_service = InstructorWorkflowService()
        self._logger = _logger
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._is_closing = False
        self._step2_marks_default_name = t("instructor.dialog.step1.prepare.default_name")
        self._workflow_controller = InstructorWorkflowController(self)
        self._async_runner = AsyncOperationRunner(
            self,
            run_async=run_in_background,
            refresh_ui=lambda: self._refresh_ui(),
            should_refresh_ui=lambda: not self._is_closing,
        )
        self._runtime = ModuleRuntime(
            module=self,
            app_name=APP_NAME,
            logger=self._logger,
            async_runner=self._async_runner,
            messages_namespace_factory=_messages_namespace,
        )

        self._ui_log_handler: UILogHandler | None = None
        self._user_log_entries: list[dict[str, object]] = []
        self._step_items: list[QListWidgetItem] = []
        self._build_ui()
        self._setup_ui_logging()
        self._refresh_ui()

    def _build_ui(self) -> None:
        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                top_object_name="instructorTopRegion",
                footer_height=INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
            ),
        )
        top_pane = QWidget()
        top_layout = QHBoxLayout(top_pane)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        self._ui_engine.set_top_widget(top_pane)
        left_card = QFrame()
        left_card.setObjectName("stepRail")
        left_card.setFrameShape(QFrame.Shape.StyledPanel)
        left_card.setFrameShadow(QFrame.Shadow.Raised)
        left_card_layout = QVBoxLayout(left_card)
        left_card_layout.setContentsMargins(0, 0, 0, 0)
        left_card_layout.setSpacing(0)
        left_scroll = QScrollArea()
        left_scroll.setObjectName("instructorLeftScroll")
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        left_scroll.viewport().setObjectName("instructorLeftScrollViewport")
        left_content = QWidget()
        left_layout = QVBoxLayout(left_content)
        left_layout.setContentsMargins(*MODULE_LEFT_PANE_CONTENT_MARGINS)
        left_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
        left_scroll.setWidget(left_content)
        left_card_layout.addWidget(left_scroll, 1)
        left_card.setFixedWidth(_LEFT_PANE_WIDTH)
        top_layout.addWidget(left_card, 0)

        self.rail_title = QLabel(t("instructor.workflow_title"))
        self.rail_title.setObjectName("instructorRailTitle")
        left_layout.addWidget(self.rail_title)

        self.download_course_template_link = QLabel()
        self.download_course_template_link.setTextFormat(Qt.TextFormat.RichText)
        self.download_course_template_link.setTextInteractionFlags(
            Qt.TextInteractionFlag.LinksAccessibleByMouse | Qt.TextInteractionFlag.LinksAccessibleByKeyboard
        )
        self.download_course_template_link.setCursor(Qt.CursorShape.PointingHandCursor)
        self.download_course_template_link.setOpenExternalLinks(False)
        self.download_course_template_link.linkActivated.connect(self._on_download_course_template_link_activated)
        left_layout.addWidget(self.download_course_template_link)

        self.step_list = QListWidget()
        self.step_list.setObjectName("stepList")
        self.step_list.setFrameShape(QFrame.Shape.NoFrame)
        self.step_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.step_list.setCursor(Qt.CursorShape.ArrowCursor)
        self.step_list.setWordWrap(True)
        self.step_list.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.step_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        for step in self.WORKFLOW_STEPS:
            item = QListWidgetItem(self._step_list_text(step))
            self.step_list.addItem(item)
            self._step_items.append(item)
        self.step_list.currentRowChanged.connect(self._on_step_row_changed)
        left_layout.addWidget(self.step_list, 1)
        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_pane)
        right_scroll = QScrollArea()
        right_scroll.setObjectName("instructorRightScroll")
        right_scroll.setFrameShape(QFrame.Shape.NoFrame)
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        right_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        right_scroll.viewport().setObjectName("instructorRightScrollViewport")
        right_scroll.setWidget(right_pane)
        top_layout.addWidget(right_scroll, 1)

        self.active_title = QLabel()
        self.active_title.setWordWrap(True)

        self.active_desc = QLabel()
        self.active_desc.setWordWrap(True)

        self.active_note = QLabel(t("instructor.note.default"))
        self.active_note.setWordWrap(True)
        self.active_note.setObjectName("hintText")
        self.active_title.setVisible(False)
        self.active_desc.setVisible(False)
        self.active_note.setVisible(False)

        self.primary_action = QPushButton()
        self.primary_action.setObjectName("secondaryAction")
        self.primary_action.setDefault(True)
        self.primary_action.clicked.connect(self._run_current_step_action)
        self.primary_action.setVisible(False)

        # Hidden upload action used by keyboard shortcuts and test harness hooks.
        self.step1_upload_action = QPushButton()
        self.step1_upload_action.setObjectName("secondaryAction")
        self.step1_upload_action.setDefault(True)
        self.step1_upload_action.clicked.connect(self._on_step1_upload_clicked)
        self.step1_upload_action.setVisible(False)

        self.step1_drop_widget = ManagedDropFileWidget(
            drop_mode="multiple",
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            open_file_tooltip=t("instructor.links.open_file"),
            open_folder_tooltip=t("instructor.links.open_folder"),
            remove_tooltip=t("coordinator.file.remove_tooltip"),
        )
        # Shortcut aliases for direct access in module handlers/tests.
        self.step1_drop_zone = self.step1_drop_widget.drop_zone
        self.step1_drop_list = self.step1_drop_widget.drop_list
        self.step1_drop_widget.files_dropped.connect(self._on_step1_course_details_dropped)
        self.step1_drop_widget.files_rejected.connect(self._on_step1_drop_files_rejected)
        self.step1_drop_widget.browse_requested.connect(self._on_step1_drop_browse_requested)
        self.step1_drop_widget.browse_requested.connect(self._on_step1_upload_clicked)
        self.step1_drop_widget.files_changed.connect(self._on_step1_drop_files_changed)
        self.step1_drop_widget.clear_button.clicked.connect(self._on_step1_clear_all_clicked)
        self.step1_drop_widget.submit_requested.connect(self._on_step1_prepare_clicked)
        self.step1_drop_widget.set_summary_text_builder(
            lambda count: t("instructor.step1.drop.summary", count=count)
        )
        self.step1_drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.step1_prepare_action = self.step1_drop_widget.submit_button
        self.step1_prepare_action.setObjectName("primaryAction")
        self.step1_prepare_action.setCursor(Qt.CursorShape.PointingHandCursor)

        # Hidden upload action used by keyboard shortcuts and test harness hooks.
        self.step2_upload_action = QPushButton()
        self.step2_upload_action.setObjectName("secondaryAction")
        self.step2_upload_action.setDefault(True)
        self.step2_upload_action.clicked.connect(self._on_step2_upload_clicked)
        self.step2_upload_action.setVisible(False)

        self.step2_drop_widget = ManagedDropFileWidget(
            drop_mode="multiple",
            remove_fallback_text=t("coordinator.file.remove_fallback"),
            open_file_tooltip=t("instructor.links.open_file"),
            open_folder_tooltip=t("instructor.links.open_folder"),
            remove_tooltip=t("coordinator.file.remove_tooltip"),
        )
        self.step2_drop_zone = self.step2_drop_widget.drop_zone
        self.step2_drop_list = self.step2_drop_widget.drop_list
        self.step2_drop_widget.files_dropped.connect(self._on_step2_filled_marks_dropped)
        self.step2_drop_widget.files_rejected.connect(self._on_step2_drop_files_rejected)
        self.step2_drop_widget.browse_requested.connect(self._on_step2_drop_browse_requested)
        self.step2_drop_widget.browse_requested.connect(self._on_step2_upload_clicked)
        self.step2_drop_widget.files_changed.connect(self._on_step2_drop_files_changed)
        self.step2_drop_widget.clear_button.clicked.connect(self._on_step2_clear_all_clicked)
        self.step2_drop_widget.submit_requested.connect(self._on_step2_generate_clicked)
        self.step2_drop_widget.set_summary_text_builder(
            lambda count: t("instructor.step1.drop.summary", count=count)
        )
        self.step2_drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.step2_generate_action = self.step2_drop_widget.submit_button
        self.step2_generate_action.setObjectName("primaryAction")
        self.step2_generate_action.setCursor(Qt.CursorShape.PointingHandCursor)
        self.step_drop_stack = QStackedWidget()
        self.step_drop_stack.addWidget(self.step1_drop_widget)
        self.step_drop_stack.addWidget(self.step2_drop_widget)
        right_layout.addWidget(self.step_drop_stack, 1)

        self.user_log_view = _LogSink()
        self._ui_engine.set_footer_visible(False)

        self.shortcut_open_workbook = QShortcut(QKeySequence(QKeySequence.StandardKey.Open), self)
        self.shortcut_open_workbook.activated.connect(self._on_open_shortcut_activated)
        self.shortcut_save_output = QShortcut(QKeySequence(QKeySequence.StandardKey.Save), self)
        self.shortcut_save_output.activated.connect(self._on_save_shortcut_activated)

    def _output_items(self) -> tuple[OutputItem, ...]:
        allowed_attrs_by_step: dict[int, set[str]] = {
            1: {"step1_path", "marks_template_path"},
            2: {"final_report_path"},
        }
        allowed_attrs = allowed_attrs_by_step.get(
            self.current_step,
            {attr for _label_key, attr in self.RAIL_LINKS},
        )
        rows: list[OutputItem] = []
        for label_key, attr in self.RAIL_LINKS:
            if attr not in allowed_attrs:
                continue
            if attr == "marks_template_path":
                batch_paths = [path for path in self.marks_template_paths if path]
                if batch_paths:
                    rows.extend(OutputItem(label_key=label_key, path=path) for path in batch_paths)
                    continue
            if attr == "final_report_path":
                batch_paths = [path for path in self.final_report_paths if path]
                if batch_paths:
                    rows.extend(OutputItem(label_key=label_key, path=path) for path in batch_paths)
                    continue
            value = getattr(self, attr)
            if isinstance(value, str) and value:
                rows.append(OutputItem(label_key=label_key, path=value))
        return tuple(rows)

    def _step_path(self, step: int) -> str | None:
        return self._workflow_controller.step_path(step)

    def _step_done(self, step: int) -> bool:
        return self._workflow_controller.step_done(step)

    def _step_outdated(self, step: int) -> bool:
        return self._workflow_controller.step_outdated(step)

    def _step_state_text(self, step: int) -> str:
        return self._workflow_controller.step_state_text(step)

    def _step_list_text(self, step: int) -> str:
        if step == 2 and not _ENABLE_SECOND_ROW_ACTIONS:
            return ""
        return self._workflow_controller.step_list_text(step)

    def _action_text_for_step(self, step: int) -> str:
        return self._workflow_controller.action_text_for_step(step)

    def _can_run_step(self, step: int) -> tuple[bool, str]:
        return self._workflow_controller.can_run_step(step)

    def _on_step_selected(self, step: int) -> None:
        self._workflow_controller.on_step_selected(step)

    def _on_step_row_changed(self, row: int) -> None:
        if row < 0 or row >= len(self.WORKFLOW_STEPS):
            return
        self._on_step_selected(row + 1)

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self._ui_engine.set_footer_visible(False)

    def get_shared_outputs_data(self) -> OutputPanelData:
        return OutputPanelData(items=self._output_items(), open_failed_key=self.RAIL_LINK_OPEN_FAILED_KEY)

    def _refresh_ui(self) -> None:
        if self.current_step not in self.WORKFLOW_STEPS:
            self.current_step = self.WORKFLOW_STEPS[0]
            self.state.current_step = self.current_step
        self.rail_title.setText(t("instructor.workflow_title"))
        self._set_download_course_template_link_enabled(not self.state.busy)

        self.step_list.blockSignals(True)
        for index, item in enumerate(self._step_items, start=1):
            item.setText(self._step_list_text(index))
        self.step_list.setCurrentRow(self.current_step - 1)
        self.step_list.blockSignals(False)

        can_run, reason = self._can_run_step(self.current_step)
        is_step1 = self.current_step == 1
        is_step2 = self.current_step == 2
        self.active_title.clear()
        self.active_desc.clear()
        self.active_note.clear()
        self.active_title.setVisible(False)
        self.active_desc.setVisible(False)
        self.active_note.setVisible(False)

        self.primary_action.setVisible(False)
        if is_step1:
            self.step_drop_stack.setCurrentWidget(self.step1_drop_widget)
        else:
            self.step_drop_stack.setCurrentWidget(self.step2_drop_widget)
        self.step_drop_stack.setVisible(is_step1 or (is_step2 and _ENABLE_SECOND_ROW_ACTIONS))

        if is_step1:
            self.step1_upload_action.setText(t("instructor.action.step1.upload"))
            self.step1_drop_widget.set_clear_button_text(t("coordinator.clear_all"))
            self.step1_drop_widget.set_submit_button_text(t("instructor.action.step1.prepare"))
            self.step1_drop_widget.set_summary_text_builder(
                lambda count: t("instructor.step1.drop.summary", count=count)
            )
            self.step1_drop_widget.set_submit_allowed(True)
        elif is_step2:
            self.step2_upload_action.setText(t("instructor.action.step2.upload.default"))
            self.step2_drop_widget.set_submit_button_text(t("instructor.action.step2.generate.default"))
            self.step2_drop_widget.set_clear_button_text(t("coordinator.clear_all"))
            self.step2_drop_widget.set_summary_text_builder(
                lambda count: t("instructor.step1.drop.summary", count=count)
            )
            self.step2_upload_action.setEnabled(_ENABLE_SECOND_ROW_ACTIONS and can_run)
            self.step2_drop_widget.set_submit_allowed(_ENABLE_SECOND_ROW_ACTIONS and can_run)

        if reason:
            self.active_note.setText(reason)
            self.active_note.setVisible(True)
        elif is_step2 and self._step_outdated(self.current_step):
            self.active_note.setText(t("instructor.note.outdated_current"))
            self.active_note.setVisible(True)
        elif is_step2 and (self.filled_marks_outdated or self.final_report_outdated):
            self.active_note.setText(t("instructor.note.outdated_downstream"))
            self.active_note.setVisible(True)

        self.step1_upload_action.setEnabled(is_step1)
        self.step1_drop_widget.setEnabled(is_step1 and not self.state.busy)
        self.step2_upload_action.setEnabled(
            is_step2 and _ENABLE_SECOND_ROW_ACTIONS and can_run and not self.state.busy
        )
        self.step2_drop_widget.setEnabled(
            is_step2 and _ENABLE_SECOND_ROW_ACTIONS and not self.state.busy
        )

        if self.state.busy:
            self.primary_action.setEnabled(False)
            self.step1_upload_action.setEnabled(False)
            self.step1_drop_widget.setEnabled(False)
            self.step2_upload_action.setEnabled(False)
            self.step2_drop_widget.setEnabled(False)

    def retranslate_ui(self) -> None:
        self.step1_drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self.step2_drop_widget.drop_list.set_placeholder_text(t("common.dropzone.placeholder"))
        self._refresh_ui()

    def _run_current_step_action(self) -> None:
        if self.current_step == 1:
            self._upload_course_details_async()
            self._refresh_ui()
            return
        if self.current_step == 2:
            if not _ENABLE_SECOND_ROW_ACTIONS:
                self._refresh_ui()
                return
            self._generate_final_report_async()
            self._refresh_ui()
            return
        self._refresh_ui()

    def _set_download_course_template_link_enabled(self, enabled: bool) -> None:
        text = t("instructor.action.step1.default")
        if enabled:
            self.download_course_template_link.setText(
                t(
                    "instructor.action.step1.link_html",
                    href=_DOWNLOAD_COURSE_TEMPLATE_HREF,
                    text=text,
                )
            )
        else:
            self.download_course_template_link.setText(text)
        self.download_course_template_link.setEnabled(enabled)

    def _on_download_course_template_link_activated(self, _href: str) -> None:
        if self.state.busy:
            return
        self._download_course_template_async()
        self._refresh_ui()

    def _on_step1_upload_clicked(self) -> None:
        self._upload_course_details_async()
        self._refresh_ui()

    def _on_step1_drop_browse_requested(self) -> None:
        self._publish_status_key("instructor.status.step1_drop_browse_requested")

    def _on_step1_course_details_dropped(self, dropped_files: list[str]) -> None:
        dropped_count = len([path for path in dropped_files if path])
        self._publish_status_key("instructor.status.step1_drop_files_dropped", count=dropped_count)
        if self.state.busy:
            return
        selected_paths = [path for path in dropped_files if path]
        if not selected_paths:
            return
        self._upload_course_details_from_paths_async(selected_paths)
        self._refresh_ui()

    def _set_step1_drop_active(self, active: bool) -> None:
        self.step1_drop_zone.set_drag_active(active)

    def _on_step1_clear_all_clicked(self) -> None:
        if self.state.busy:
            return
        self.step1_drop_widget.clear_files()
        self._refresh_ui()

    def _set_step1_course_details_files(self, file_paths: list[str]) -> None:
        self.step1_drop_widget.set_files(file_paths)

    def _on_step1_drop_files_changed(self, files: list[str]) -> None:
        self._publish_status_key("instructor.status.step1_drop_files_changed", count=len(files))
        if self.state.busy:
            return
        current_valid = [*self.step1_course_details_paths]
        if files:
            current_keys = {str(Path(path).resolve(strict=False)).lower() for path in current_valid}
            incoming_keys = {str(Path(path).resolve(strict=False)).lower() for path in files}
            if incoming_keys.issubset(current_keys):
                key_to_existing = {
                    str(Path(path).resolve(strict=False)).lower(): path
                    for path in self.step1_course_details_paths
                }
                self.step1_course_details_paths = [
                    key_to_existing[key]
                    for key in [str(Path(path).resolve(strict=False)).lower() for path in files]
                    if key in key_to_existing
                ]
                self.step2_course_details_path = (
                    self.step1_course_details_paths[0] if self.step1_course_details_paths else None
                )
                self.step2_upload_ready = bool(self.step1_course_details_paths)
            self._refresh_ui()
            return
        if not self.step2_course_details_path:
            self._refresh_ui()
            return
        self.step1_course_details_paths = []
        self.step2_course_details_path = None
        self.step2_upload_ready = False
        self.marks_template_done = False
        self.marks_template_path = None
        self.marks_template_paths = []
        self._step2_marks_default_name = t("instructor.dialog.step1.prepare.default_name")
        self.filled_marks_outdated = self.filled_marks_done
        self.final_report_outdated = self.final_report_done
        if self.filled_marks_outdated or self.final_report_outdated:
            self._publish_status_key("instructor.status.step1_changed")
        self._refresh_ui()

    def _on_step1_drop_files_rejected(self, files: list[str]) -> None:
        rejected_count = len([path for path in files if path])
        if rejected_count <= 0:
            return
        self._publish_status_key("instructor.status.step1_drop_files_rejected", count=rejected_count)

    def _on_step1_prepare_clicked(self) -> None:
        self._prepare_marks_template_async()
        self._refresh_ui()

    def _on_step2_upload_clicked(self) -> None:
        if not _ENABLE_SECOND_ROW_ACTIONS:
            return
        self._upload_filled_marks_from_dialog_async()
        self._refresh_ui()

    def _on_step2_generate_clicked(self) -> None:
        if not _ENABLE_SECOND_ROW_ACTIONS:
            return
        self._generate_final_report_async()
        self._refresh_ui()

    def _on_step2_drop_browse_requested(self) -> None:
        self._publish_status_key("instructor.status.step2_drop_browse_requested")

    def _on_step2_filled_marks_dropped(self, dropped_files: list[str]) -> None:
        dropped_count = len([path for path in dropped_files if path])
        self._publish_status_key("instructor.status.step2_drop_files_dropped", count=dropped_count)
        if self.state.busy:
            return
        selected_paths = [path for path in dropped_files if path]
        if not selected_paths:
            return
        self._upload_filled_marks_from_paths_async(selected_paths)
        self._refresh_ui()

    def _on_step2_drop_files_rejected(self, files: list[str]) -> None:
        rejected_count = len([path for path in files if path])
        if rejected_count <= 0:
            return
        self._publish_status_key("instructor.status.step2_drop_files_rejected", count=rejected_count)

    def _on_step2_clear_all_clicked(self) -> None:
        if self.state.busy:
            return
        self.step2_drop_widget.clear_files()
        self._refresh_ui()

    def _on_step2_drop_files_changed(self, files: list[str]) -> None:
        self._publish_status_key("instructor.status.step2_drop_files_changed", count=len(files))
        if self.state.busy:
            return
        current_valid = [*self.filled_marks_paths]
        if files:
            current_keys = {str(Path(path).resolve(strict=False)).lower() for path in current_valid}
            incoming_keys = {str(Path(path).resolve(strict=False)).lower() for path in files}
            if incoming_keys.issubset(current_keys):
                key_to_existing = {
                    str(Path(path).resolve(strict=False)).lower(): path
                    for path in self.filled_marks_paths
                }
                self.filled_marks_paths = [
                    key_to_existing[key]
                    for key in [str(Path(path).resolve(strict=False)).lower() for path in files]
                    if key in key_to_existing
                ]
                self.filled_marks_path = self.filled_marks_paths[0] if self.filled_marks_paths else None
                self.filled_marks_done = bool(self.filled_marks_paths)
            self._refresh_ui()
            return
        if not self.filled_marks_paths:
            self._refresh_ui()
            return
        self.filled_marks_paths = []
        self.filled_marks_path = None
        self.filled_marks_done = False
        self.final_report_outdated = self.final_report_done
        if self.final_report_outdated:
            self._publish_status_key("instructor.status.step2_changed_filled")
        self._refresh_ui()

    def _on_open_shortcut_activated(self) -> None:
        if self.state.busy:
            return
        if self.current_step == 1 and self.step1_upload_action.isEnabled():
            self._on_step1_upload_clicked()
            return
        if _ENABLE_SECOND_ROW_ACTIONS and self.current_step == 2 and self.step2_upload_action.isEnabled():
            self._on_step2_upload_clicked()

    def _on_save_shortcut_activated(self) -> None:
        if self.state.busy:
            return
        if self.current_step == 1 and self.step1_prepare_action.isEnabled():
            self._on_step1_prepare_clicked()
            return
        if _ENABLE_SECOND_ROW_ACTIONS and self.current_step == 2:
            self._on_step2_generate_clicked()

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        self._runtime.remember_dialog_dir_safe(selected_path)

    def _setup_ui_logging(self) -> None:
        self._runtime.setup_ui_logging()

    def _append_user_log(self, message: str) -> None:
        self._runtime.append_user_log(message)

    def _publish_status(self, message: str) -> None:
        self._runtime.publish_status(message)

    def _publish_status_key(self, text_key: str, **kwargs: object) -> None:
        self._runtime.publish_status_key(text_key, **kwargs)

    def _rerender_user_log(self) -> None:
        _rerender_user_log_impl(self, ns=_messages_namespace())

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        host_window = self.window()
        set_switch = getattr(host_window, "set_language_switch_enabled", None)
        if callable(set_switch):
            set_switch(not busy)
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

    def closeEvent(self, event) -> None:
        self._is_closing = True
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._cancel_token = None
        self._active_jobs.clear()
        if self._ui_log_handler is not None:
            self._logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)

    def _prepare_marks_template_async(self) -> None:
        prepare_marks_template_async(self, ns=globals())

    def _download_course_template_async(self) -> None:
        download_course_template_async(self, ns=globals())

    def _upload_course_details_async(self) -> None:
        upload_course_details_async(self, ns=globals())

    def _upload_course_details_from_paths_async(self, open_paths: list[str]) -> None:
        upload_course_details_from_paths_async(self, open_paths, ns=globals())

    def _upload_filled_marks_async(self) -> None:
        upload_filled_marks_async(self, ns=globals())

    def _upload_filled_marks_from_dialog_async(self) -> None:
        if self.state.busy:
            return
        open_paths, _ = QFileDialog.getOpenFileNames(
            self,
            t("instructor.dialog.step2.upload.title"),
            resolve_dialog_start_path(APP_NAME),
            t("instructor.dialog.filter.excel_open"),
        )
        if not open_paths:
            return
        self._upload_filled_marks_from_paths_async(open_paths)

    def _upload_filled_marks_from_paths_async(self, open_paths: list[str]) -> None:
        if self.state.busy:
            return
        selected_paths = [path for path in open_paths if path]
        if not selected_paths:
            return
        self._remember_dialog_dir_safe(selected_paths[0])
        token = CancellationToken()

        def _work() -> dict[str, object]:
            valid: list[str] = []
            invalid: list[str] = []
            seen: set[str] = set()
            duplicates = 0
            for path in selected_paths:
                key = str(Path(path).resolve(strict=False)).lower()
                if key in seen:
                    duplicates += 1
                    continue
                seen.add(key)
                token.raise_if_cancelled()
                try:
                    _validate_uploaded_filled_marks_workbook(path)
                    valid.append(path)
                except Exception:
                    invalid.append(path)
            return {
                "valid": valid,
                "invalid": invalid,
                "duplicates": duplicates,
            }

        def _on_success(result: object) -> None:
            data = result if isinstance(result, dict) else {}
            valid = [p for p in data.get("valid", []) if isinstance(p, str) and p]
            invalid = [p for p in data.get("invalid", []) if isinstance(p, str) and p]
            duplicates = int(data.get("duplicates", 0))

            merged: list[str] = []
            seen_keys: set[str] = set()
            for value in [*self.filled_marks_paths, *valid]:
                key = str(Path(value).resolve(strict=False)).lower()
                if key in seen_keys:
                    duplicates += 1
                    continue
                seen_keys.add(key)
                merged.append(value)
            replacing = bool(self.filled_marks_paths)
            self.filled_marks_paths = merged
            self.filled_marks_path = self.filled_marks_paths[0] if self.filled_marks_paths else None
            self.filled_marks_done = bool(self.filled_marks_paths)
            self.filled_marks_outdated = False
            self.step2_drop_widget.set_files(self.filled_marks_paths)
            if replacing and self.filled_marks_done:
                self.final_report_outdated = True
                self._publish_status_key("instructor.status.step2_changed_filled")
            elif self.filled_marks_done:
                self._publish_status_key("instructor.status.step2_uploaded_filled")
            if invalid or duplicates:
                show_toast(
                    self,
                    t(
                        "instructor.toast.step2_upload_reject_summary",
                        invalid=len(invalid),
                        duplicates=duplicates,
                    ),
                    title=t("instructor.step2.title"),
                    level="warning",
                )

        def _on_failure(_exc: Exception) -> None:
            self._show_system_error_toast(2)

        self._start_async_operation(token=token, job_id=None, work=_work, on_success=_on_success, on_failure=_on_failure)

    def _generate_final_report_async(self) -> None:
        if self.filled_marks_paths:
            generate_final_reports_from_paths_async(self, ns=globals())
            return
        generate_final_report_async(self, ns=globals())

    def _show_step_success_toast(self, step: int) -> None:
        show_step_success_toast(self, step=step, title_key=self.STEP_TITLE_KEYS[step])

    def _show_validation_error_toast(self, message: str) -> None:
        show_validation_error_toast(self, message)

    def _show_system_error_toast(self, step: int) -> None:
        show_system_error_toast(self, title_key=self.STEP_TITLE_KEYS[step])
