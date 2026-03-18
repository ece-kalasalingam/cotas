"""Course Instructor CO module UI (list-based step navigation)."""

from __future__ import annotations

import importlib
import logging
import time
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QDesktopServices, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
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
    ID_COURSE_SETUP,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
)
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
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
from domain import InstructorWorkflowState
from modules.instructor import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)
from modules.instructor.async_runner import AsyncOperationRunner
from modules.instructor.async_runner import (
    publish_status_compat as _publish_status_compat_impl,
)
from modules.instructor.async_runner import set_busy_compat as _set_busy_compat_impl
from modules.instructor.async_runner import (
    start_async_operation_compat as _start_async_operation_compat_impl,
)
from modules.instructor.messages import (
    localized_log_messages,
    show_step_success_toast,
    show_system_error_toast,
    show_validation_error_toast,
)
from modules.instructor.output_links import InstructorOutputNamespace
from modules.instructor.output_links import (
    on_quick_link_activated as _on_quick_link_activated_impl,
)
from modules.instructor.output_links import quick_link_markup as _quick_link_markup_impl
from modules.instructor.output_links import quick_links_html as _quick_links_html_impl
from modules.instructor.output_links import (
    refresh_quick_links as _refresh_quick_links_impl,
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
)
from modules.instructor.steps.step2_filled_marks_and_final_report import (
    generate_final_report_async,
    upload_filled_marks_async,
)
from modules.instructor.validators.step2_filled_marks_validator import (
    filled_marks_manifest_validators,
    validate_filled_marks_manifest_schema_by_template,
    validate_uploaded_filled_marks_workbook,
)
from modules.instructor.workflow_controller import InstructorWorkflowController
from services import InstructorWorkflowService

OUTPUT_LINK_MODE_FILE = "file"
_OUTPUT_LINK_RUNTIME_GLOBALS = (
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
    QUrl,
    QDesktopServices,
    show_toast,
)

_logger = logging.getLogger(__name__)
shutil = importlib.import_module("shutil")
_INSTRUCTOR_STATUS_KEYS = (
    "instructor.status.step1_selected",
    "instructor.status.step1_validated",
    "instructor.status.step1_prepared",
    "instructor.status.step1_changed",
    "instructor.status.step2_uploaded_filled",
    "instructor.status.step2_changed_filled",
    "instructor.status.step2_generated",
    "instructor.status.operation_in_progress",
    "instructor.status.operation_cancelled",
)

# Step handlers receive `ns=globals()`, so these names must stay module-visible.
_STEP_RUNTIME_GLOBALS = (
    QFileDialog,
    ID_COURSE_SETUP,
    AppSystemError,
    JobCancelledError,
    ValidationError,
    log_process_message,
    resolve_dialog_start_path,
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
    build_i18n_log_message,
)

def _output_link_namespace() -> InstructorOutputNamespace:
    return {
        "t": t,
        "OUTPUT_LINK_MODE_FILE": OUTPUT_LINK_MODE_FILE,
        "OUTPUT_LINK_MODE_FOLDER": OUTPUT_LINK_MODE_FOLDER,
        "OUTPUT_LINK_SEPARATOR": OUTPUT_LINK_SEPARATOR,
        "url_from_local_file": QUrl.fromLocalFile,
        "open_url": QDesktopServices.openUrl,
        "show_toast": show_toast,
    }


def _publish_status_compat(target: object, message: str) -> None:
    payload = message
    supports_payload = bool(getattr(target, "_supports_i18n_status_payload", False))
    if supports_payload:
        for key in _INSTRUCTOR_STATUS_KEYS:
            if message == t(key):
                payload = build_i18n_log_message(key, fallback=message)
                break
    _publish_status_compat_impl(target=target, message=payload, logger=_logger)


def _set_busy_compat(target: object, busy: bool, *, job_id: str | None = None) -> None:
    _set_busy_compat_impl(target=target, busy=busy, job_id=job_id)


def _start_async_operation_compat(
    target: object,
    *,
    token: CancellationToken,
    job_id: str | None,
    work,
    on_success,
    on_failure,
) -> None:
    _start_async_operation_compat_impl(
        target=target,
        token=token,
        job_id=job_id,
        work=work,
        on_success=on_success,
        on_failure=on_failure,
        run_async=run_in_background,
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


class InstructorModule(QWidget):
    """Simple wizard-like UI for CO score workflow."""

    status_changed = Signal(str)
    WORKFLOW_STEPS = (1, 2)
    RAIL_LINK_TITLE_KEY = "instructor.links.title"
    RAIL_LINKS = (
        ("instructor.links.course_details_generated", "step1_path"),
        ("instructor.links.course_details_uploaded", "step2_course_details_path"),
        ("instructor.links.marks_template_generated", "marks_template_path"),
        ("instructor.links.marks_template_uploaded", "filled_marks_path"),
        ("instructor.links.final_co_report_generated", "final_report_path"),
    )
    RAIL_LINK_OPEN_FILE_KEY = "instructor.links.open_file"
    RAIL_LINK_OPEN_FOLDER_KEY = "instructor.links.open_folder"
    RAIL_LINK_NOT_AVAILABLE_KEY = "instructor.links.not_available"
    RAIL_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"

    STEP_TITLE_KEYS = {
        1: "instructor.step1.title",
        2: "instructor.step2.title",
    }

    STEP_DESC_KEYS = {
        1: "instructor.step1.desc",
        2: "instructor.step2.desc",
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
        self._supports_i18n_status_payload = True
        self.current_step = 1

        self.step1_path: str | None = None
        self.marks_template_path: str | None = None
        self.step2_course_details_path: str | None = None
        self.filled_marks_path: str | None = None
        self.final_report_path: str | None = None

        self.step1_done = False
        self.marks_template_done = False
        self.filled_marks_done = False
        self.final_report_done = False
        self.step2_upload_ready = False

        self.filled_marks_outdated = False  # Filled marks
        self.final_report_outdated = False  # Final report

        self.state = InstructorWorkflowState()
        self._workflow_service = InstructorWorkflowService()
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._is_closing = False
        self._step2_marks_default_name = t("instructor.dialog.step1.prepare.default_name")
        self._workflow_controller = InstructorWorkflowController(self)
        self._async_runner = AsyncOperationRunner(self, run_async=run_in_background)

        self._ui_log_handler: UILogHandler | None = None
        self._user_log_entries: list[dict[str, object]] = []
        self._step_items: list[QListWidgetItem] = []
        self._busy_started_at: float | None = None
        self._busy_elapsed_timer = QTimer(self)
        self._busy_elapsed_timer.setInterval(1000)
        self._busy_elapsed_timer.timeout.connect(self._update_busy_timer_label)
        self._build_ui()
        self._setup_ui_logging()
        self._refresh_ui()

    def _build_ui(self) -> None:
        root = QHBoxLayout(self)

        left = QFrame()
        left.setObjectName("stepRail")
        left_layout = QVBoxLayout(left)

        self.rail_title = QLabel(t("instructor.workflow_title"))
        left_layout.addWidget(self.rail_title)

        self.busy_timer_label = QLabel()
        self.busy_timer_label.setObjectName("hintText")
        left_layout.addWidget(self.busy_timer_label)
        self.download_course_template_button = QPushButton(t("instructor.action.step1.default"))
        self.download_course_template_button.setObjectName("secondaryAction")
        self.download_course_template_button.clicked.connect(self._on_download_course_template_clicked)
        left_layout.addWidget(self.download_course_template_button)

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
        left_scroll = QScrollArea()
        left_scroll.setObjectName("instructorLeftScroll")
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        left_scroll.setWidget(left)

        right = QFrame()
        right.setObjectName("activeCard")
        right_layout = QVBoxLayout(right)

        top_content = QWidget()
        top_layout = QVBoxLayout(top_content)

        self.active_title = QLabel()
        self.active_title.setWordWrap(True)
        top_layout.addWidget(self.active_title)

        self.active_desc = QLabel()
        self.active_desc.setWordWrap(True)
        top_layout.addWidget(self.active_desc)

        self.active_note = QLabel(t("instructor.note.default"))
        self.active_note.setWordWrap(True)
        self.active_note.setObjectName("hintText")
        top_layout.addWidget(self.active_note)

        self.primary_action = QPushButton()
        self.primary_action.setObjectName("secondaryAction")
        self.primary_action.setDefault(True)
        self.primary_action.clicked.connect(self._run_current_step_action)
        top_layout.addWidget(self.primary_action, alignment=Qt.AlignmentFlag.AlignLeft)

        self.step1_action_row = QWidget()
        step1_action_layout = QHBoxLayout(self.step1_action_row)

        self.step1_upload_action = QPushButton()
        self.step1_upload_action.setObjectName("secondaryAction")
        self.step1_upload_action.setDefault(True)
        self.step1_upload_action.clicked.connect(self._on_step1_upload_clicked)
        step1_action_layout.addWidget(self.step1_upload_action)

        self.step1_prepare_action = QPushButton()
        self.step1_prepare_action.setObjectName("primaryAction")
        self.step1_prepare_action.setAutoDefault(False)
        self.step1_prepare_action.setDefault(False)
        self.step1_prepare_action.clicked.connect(self._on_step1_prepare_clicked)
        step1_action_layout.addWidget(self.step1_prepare_action)
        step1_action_layout.addStretch(1)
        top_layout.addWidget(self.step1_action_row)

        self.step2_action_row = QWidget()
        step2_action_layout = QHBoxLayout(self.step2_action_row)

        self.step2_upload_action = QPushButton()
        self.step2_upload_action.setObjectName("secondaryAction")
        self.step2_upload_action.setDefault(True)
        self.step2_upload_action.clicked.connect(self._on_step2_upload_clicked)
        step2_action_layout.addWidget(self.step2_upload_action)

        self.step2_generate_action = QPushButton()
        self.step2_generate_action.setObjectName("primaryAction")
        self.step2_generate_action.setAutoDefault(False)
        self.step2_generate_action.setDefault(False)
        self.step2_generate_action.clicked.connect(self._on_step2_generate_clicked)
        step2_action_layout.addWidget(self.step2_generate_action)
        step2_action_layout.addStretch(1)
        top_layout.addWidget(self.step2_action_row)
        top_layout.addStretch(1)

        right_layout.addWidget(top_content, 1)

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
            lambda url: self._on_quick_link_activated(url.toString())
        )
        links_tab_layout.addWidget(self.generated_outputs_view)

        # Retained for test doubles that call _refresh_quick_links directly.
        self.quick_link_labels: dict[str, QLabel] = {}

        self.info_tabs.addTab(log_tab, t("instructor.log.title"))
        self.info_tabs.addTab(links_tab, t(self.RAIL_LINK_TITLE_KEY))
        right_layout.addWidget(self.info_tabs)

        self.shortcut_open_workbook = QShortcut(QKeySequence(QKeySequence.StandardKey.Open), self)
        self.shortcut_open_workbook.activated.connect(self._on_open_shortcut_activated)
        self.shortcut_save_output = QShortcut(QKeySequence(QKeySequence.StandardKey.Save), self)
        self.shortcut_save_output.activated.connect(self._on_save_shortcut_activated)

        root.addWidget(left_scroll)
        root.addWidget(right, 1)

    def _quick_link_items(self) -> tuple[tuple[str, str | None], ...]:
        return tuple((label_key, getattr(self, attr)) for label_key, attr in self.RAIL_LINKS)

    def _quick_link_markup(self, label_key: str, path: str | None) -> str:
        return _quick_link_markup_impl(self, label_key, path, ns=_output_link_namespace())

    def _quick_links_html(self) -> str:
        return _quick_links_html_impl(self, ns=_output_link_namespace())

    def _on_quick_link_activated(self, href: str) -> None:
        _on_quick_link_activated_impl(self, href, ns=_output_link_namespace())

    def _refresh_quick_links(self) -> None:
        _refresh_quick_links_impl(self, ns=_output_link_namespace())

    def _step_path(self, step: int) -> str | None:
        return self._workflow_controller.step_path(step)

    def _step_done(self, step: int) -> bool:
        return self._workflow_controller.step_done(step)

    def _step_outdated(self, step: int) -> bool:
        return self._workflow_controller.step_outdated(step)

    def _step_state_text(self, step: int) -> str:
        return self._workflow_controller.step_state_text(step)

    def _step_list_text(self, step: int) -> str:
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

    def _clear_info_text_selection(self) -> None:
        for view in (self.user_log_view, self.generated_outputs_view):
            cursor = view.textCursor()
            if cursor.hasSelection():
                cursor.clearSelection()
                view.setTextCursor(cursor)

    def _on_info_tab_changed(self, _index: int) -> None:
        self._clear_info_text_selection()

    def set_shared_activity_log_mode(self, enabled: bool) -> None:
        self.info_tabs.setVisible(not enabled)

    def get_shared_outputs_html(self) -> str:
        return self._quick_links_html()

    def _refresh_ui(self) -> None:
        if self.current_step not in self.WORKFLOW_STEPS:
            self.current_step = self.WORKFLOW_STEPS[0]
            self.state.current_step = self.current_step
        self.rail_title.setText(t("instructor.workflow_title"))
        self.download_course_template_button.setText(t("instructor.action.step1.default"))
        self._update_busy_timer_label()

        self.step_list.blockSignals(True)
        for index, item in enumerate(self._step_items, start=1):
            item.setText(self._step_list_text(index))
        self.step_list.setCurrentRow(self.current_step - 1)
        self.step_list.blockSignals(False)

        self.active_title.setText(t(self.STEP_TITLE_KEYS[self.current_step]))
        self.active_desc.setText(t(self.STEP_DESC_KEYS[self.current_step]))
        self.info_tabs.setTabText(0, t("instructor.log.title"))
        self.info_tabs.setTabText(1, t(self.RAIL_LINK_TITLE_KEY))
        self._refresh_quick_links()
        can_run, reason = self._can_run_step(self.current_step)
        is_step1 = self.current_step == 1
        is_step2 = self.current_step == 2

        self.primary_action.setVisible(False)
        self.step1_action_row.setVisible(is_step1)
        self.step2_action_row.setVisible(is_step2)

        if is_step1:
            self.step1_upload_action.setText(t("instructor.action.step1.upload"))
            self.step1_prepare_action.setText(t("instructor.action.step1.prepare"))
            self.step1_prepare_action.setEnabled(self.step2_upload_ready)
        elif is_step2:
            self.step2_upload_action.setText(t("instructor.action.step2.upload.default"))
            self.step2_generate_action.setText(t("instructor.action.step2.generate.default"))
            self.step2_upload_action.setEnabled(True)
            self.step2_generate_action.setEnabled(
                self.filled_marks_done and not self.filled_marks_outdated
            )

        self.step1_upload_action.setEnabled(is_step1)
        if not can_run:
            self.active_note.setText(reason)
        elif self.filled_marks_outdated or self.final_report_outdated:
            if self._step_outdated(self.current_step):
                self.active_note.setText(t("instructor.note.outdated_current"))
            else:
                self.active_note.setText(t("instructor.note.outdated_downstream"))
        else:
            self.active_note.setText(t("instructor.note.default"))

        if self.state.busy:
            self.primary_action.setEnabled(False)
            self.download_course_template_button.setEnabled(False)
            self.step1_upload_action.setEnabled(False)
            self.step1_prepare_action.setEnabled(False)
            self.step2_upload_action.setEnabled(False)
            self.step2_generate_action.setEnabled(False)
        else:
            self.download_course_template_button.setEnabled(True)

    def retranslate_ui(self) -> None:
        self._rerender_user_log()
        self._refresh_ui()
        self._clear_info_text_selection()

    def _run_current_step_action(self) -> None:
        if self.current_step == 1:
            self._upload_course_details_async()
            self._refresh_ui()
            return
        if self.current_step == 2:
            self._generate_final_report_async()
            self._refresh_ui()
            return
        self._refresh_ui()

    def _on_download_course_template_clicked(self) -> None:
        if self.state.busy:
            return
        self._download_course_template_async()
        self._refresh_ui()

    def _on_step1_upload_clicked(self) -> None:
        self._upload_course_details_async()
        self._refresh_ui()

    def _on_step1_prepare_clicked(self) -> None:
        self._prepare_marks_template_async()
        self._refresh_ui()

    def _on_step2_upload_clicked(self) -> None:
        self._upload_filled_marks_async()
        self._refresh_ui()

    def _on_step2_generate_clicked(self) -> None:
        self._generate_final_report_async()
        self._refresh_ui()

    def _on_open_shortcut_activated(self) -> None:
        if self.state.busy:
            return
        if self.current_step == 1 and self.step1_upload_action.isEnabled():
            self._on_step1_upload_clicked()
            return
        if self.current_step == 2 and self.step2_upload_action.isEnabled():
            self._on_step2_upload_clicked()

    def _on_save_shortcut_activated(self) -> None:
        if self.state.busy:
            return
        if self.current_step == 1 and self.step1_prepare_action.isEnabled():
            self._on_step1_prepare_clicked()
            return
        if self.current_step == 2 and self.step2_generate_action.isEnabled():
            self._on_step2_generate_clicked()

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        try:
            remember_dialog_dir(selected_path, app_name=APP_NAME)
        except OSError:
            remember_dialog_dir_safe(
                selected_path,
                app_name=APP_NAME,
                logger=_logger,
            )

    def _setup_ui_logging(self) -> None:
        if self._ui_log_handler is not None:
            return
        self._ui_log_handler = UILogHandler(self._append_user_log)
        _logger.addHandler(self._ui_log_handler)
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

    def _publish_status(self, message: str) -> None:
        self._append_user_log(message)
        emit_user_status(getattr(self, "status_changed", None), message, logger=_logger)

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

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:
        self.state.set_busy(busy, job_id=job_id)
        if busy:
            self._busy_started_at = time.perf_counter()
            if not self._busy_elapsed_timer.isActive():
                self._busy_elapsed_timer.start()
        else:
            if self._busy_elapsed_timer.isActive():
                self._busy_elapsed_timer.stop()
            self._busy_started_at = None
        host_window = self.window()
        set_switch = getattr(host_window, "set_language_switch_enabled", None)
        if callable(set_switch):
            set_switch(not busy)
        self._refresh_ui()

    def _update_busy_timer_label(self) -> None:
        if not self.state.busy or self._busy_started_at is None:
            self.busy_timer_label.setText(t("instructor.timer.idle"))
            return
        elapsed_seconds = max(0, int(time.perf_counter() - self._busy_started_at))
        minutes, seconds = divmod(elapsed_seconds, 60)
        elapsed_text = f"{minutes:02d}:{seconds:02d}"
        self.busy_timer_label.setText(t("instructor.timer.running", elapsed=elapsed_text))

    def _start_async_operation(
        self,
        *,
        token: CancellationToken,
        job_id: str | None,
        work,
        on_success,
        on_failure,
    ) -> None:
        self._async_runner.start(
            token=token,
            job_id=job_id,
            work=work,
            on_success=on_success,
            on_failure=on_failure,
        )

    def closeEvent(self, event) -> None:
        self._is_closing = True
        if self._busy_elapsed_timer.isActive():
            self._busy_elapsed_timer.stop()
        if self._cancel_token is not None:
            self._cancel_token.cancel()
            self._cancel_token = None
        self._active_jobs.clear()
        if self._ui_log_handler is not None:
            _logger.removeHandler(self._ui_log_handler)
            self._ui_log_handler = None
        super().closeEvent(event)

    def _prepare_marks_template_async(self) -> None:
        prepare_marks_template_async(self, ns=globals())

    def _download_course_template_async(self) -> None:
        download_course_template_async(self, ns=globals())

    def _upload_course_details_async(self) -> None:
        upload_course_details_async(self, ns=globals())

    def _upload_filled_marks_async(self) -> None:
        upload_filled_marks_async(self, ns=globals())

    def _generate_final_report_async(self) -> None:
        generate_final_report_async(self, ns=globals())

    def _show_step_success_toast(self, step: int) -> None:
        show_step_success_toast(self, step=step, title_key=self.STEP_TITLE_KEYS[step])

    def _show_validation_error_toast(self, message: str) -> None:
        show_validation_error_toast(self, message)

    def _show_system_error_toast(self, step: int) -> None:
        show_system_error_toast(self, title_key=self.STEP_TITLE_KEYS[step])

    def _download_course_template(self) -> None:
        InstructorModule._download_course_template_async(self)

    def _upload_course_details(self) -> None:
        InstructorModule._upload_course_details_async(self)

    def _prepare_marks_template(self) -> None:
        InstructorModule._prepare_marks_template_async(self)

    def _upload_filled_marks(self) -> None:
        InstructorModule._upload_filled_marks_async(self)

    def _generate_final_report(self) -> None:
        InstructorModule._generate_final_report_async(self)
