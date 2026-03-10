"""Course Instructor CO module UI (list-based step navigation)."""

from __future__ import annotations

import logging
import importlib
import time
from html import escape
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QDesktopServices, QFont
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QPushButton,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from common.constants import (
    APP_NAME,
    ID_COURSE_SETUP,
    INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE,
    INSTRUCTOR_CARD_MARGIN,
    INSTRUCTOR_CARD_SPACING,
    INSTRUCTOR_PANEL_STYLESHEET,
    INSTRUCTOR_RAIL_MAX_WIDTH,
    INSTRUCTOR_RAIL_TITLE_FONT_SIZE,
    INSTRUCTOR_STEP2_ACTION_MARGIN,
    INSTRUCTOR_STEP2_ACTION_SPACING,
    INSTRUCTOR_STEP_LIST_SPACING,
    INSTRUCTOR_TOP_LAYOUT_MARGINS,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    INSTRUCTOR_INFO_TAB_LAYOUT_MARGINS,
    INSTRUCTOR_INFO_TAB_LAYOUT_SPACING,
    UI_FONT_FAMILY,
)
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.qt_jobs import run_in_background
from common.toast import show_toast
from common.texts import t
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
from modules.instructor.async_runner import (
    AsyncOperationRunner,
    publish_status_compat as _publish_status_compat_impl,
    set_busy_compat as _set_busy_compat_impl,
    start_async_operation_compat as _start_async_operation_compat_impl,
)
from modules.instructor.messages import (
    localized_log_messages,
    show_step_success_toast,
    show_system_error_toast,
    show_validation_error_toast,
)
from modules.instructor.steps.shared_workbook_ops import (
    atomic_copy_file as _shared_atomic_copy_file,
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
from modules.instructor.steps.step3_filled_marks_and_final_report import (
    generate_final_report_async,
    upload_filled_marks_async,
)
from modules.instructor.validators.step3_filled_marks_validator import (
    filled_marks_manifest_validators,
    validate_filled_marks_manifest_schema_by_template,
    validate_uploaded_filled_marks_workbook,
)
from modules.instructor.workflow_controller import InstructorWorkflowController
from services import InstructorWorkflowService

_logger = logging.getLogger(__name__)
shutil = importlib.import_module("shutil")

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
)


class _UILogHandler(logging.Handler):
    """Forward logger messages to the in-panel user log view."""

    def __init__(self, sink):
        super().__init__(level=logging.INFO)
        self._sink = sink

    def emit(self, record: logging.LogRecord) -> None:
        try:
            user_message = getattr(record, "user_message", None)
            message = f"{record.levelname}: {user_message or record.getMessage()}"
            self._sink(message)
        except Exception:
            self.handleError(record)


def _publish_status_compat(target: object, message: str) -> None:
    _publish_status_compat_impl(target=target, message=message, logger=_logger)


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


def _filled_marks_manifest_validators() -> dict[str, object]:
    return filled_marks_manifest_validators()


class InstructorModule(QWidget):
    """Simple wizard-like UI for CO score workflow."""

    status_changed = Signal(str)
    WORKFLOW_STEPS = (1, 2, 3)
    RAIL_LINK_TITLE_KEY = "instructor.links.title"
    RAIL_LINKS = (
        ("instructor.links.course_details_generated", "step1_path"),
        ("instructor.links.course_details_uploaded", "step2_course_details_path"),
        ("instructor.links.marks_template_generated", "step2_path"),
        ("instructor.links.marks_template_uploaded", "step3_path"),
        ("instructor.links.final_co_report_generated", "step4_path"),
    )
    RAIL_LINK_OPEN_FILE_KEY = "instructor.links.open_file"
    RAIL_LINK_OPEN_FOLDER_KEY = "instructor.links.open_folder"
    RAIL_LINK_NOT_AVAILABLE_KEY = "instructor.links.not_available"
    RAIL_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"

    STEP_TITLE_KEYS = {
        1: "instructor.step1.title",
        2: "instructor.step2.title",
        3: "instructor.step3.title",
    }

    STEP_DESC_KEYS = {
        1: "instructor.step1.desc",
        2: "instructor.step2.desc",
        3: "instructor.step3.desc",
    }

    PATH_ATTRS = {
        1: "step1_path",
        2: "step2_path",
        3: "step4_path",
    }

    DONE_ATTRS = {
        1: "step1_done",
        2: "step2_done",
        3: "step4_done",
    }

    OUTDATED_ATTRS = {
        3: "step4_outdated",
    }

    ACTION_DEFAULT_KEYS = {
        1: "instructor.action.step1.default",
        2: "instructor.action.step2.default",
        3: "instructor.action.step5.default",
    }

    ACTION_REDO_KEYS = {
        1: "instructor.action.step1.redo",
        2: "instructor.action.step2.redo",
        3: "instructor.action.step5.redo",
    }

    def __init__(self):
        super().__init__()
        self.current_step = 1

        self.step1_path: str | None = None
        self.step2_path: str | None = None
        self.step2_course_details_path: str | None = None
        self.step3_path: str | None = None
        self.step4_path: str | None = None

        self.step1_done = False
        self.step2_done = False
        self.step3_done = False
        self.step4_done = False
        self.step2_upload_ready = False

        self.step3_outdated = False  # Filled marks
        self.step4_outdated = False  # Final report

        self.state = InstructorWorkflowState()
        self._workflow_service = InstructorWorkflowService()
        self._cancel_token: CancellationToken | None = None
        self._active_jobs: list[object] = []
        self._is_closing = False
        self._step2_marks_default_name = t("instructor.dialog.step3.default_name")
        self._workflow_controller = InstructorWorkflowController(self)
        self._async_runner = AsyncOperationRunner(self, run_async=run_in_background)

        self._ui_log_handler: _UILogHandler | None = None
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
        

        rail_title = QLabel(t("instructor.workflow_title"))
        rail_title.setFont(
            QFont(UI_FONT_FAMILY, INSTRUCTOR_RAIL_TITLE_FONT_SIZE, QFont.Weight.Bold)
        )
        left_layout.addWidget(rail_title)

        self.progress_label = QLabel()
        self.progress_label.setObjectName("progressText")
        left_layout.addWidget(self.progress_label)
        self.busy_timer_label = QLabel()
        self.busy_timer_label.setObjectName("hintText")
        left_layout.addWidget(self.busy_timer_label)

        self.step_list = QListWidget()
        self.step_list.setObjectName("stepList")
        self.step_list.setFrameShape(QFrame.Shape.NoFrame)
        self.step_list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.step_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.step_list.setCursor(Qt.CursorShape.ArrowCursor)
        self.step_list.setSpacing(INSTRUCTOR_STEP_LIST_SPACING)
        self.step_list.setWordWrap(True)
        self.step_list.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.step_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        for step in self.WORKFLOW_STEPS:
            item = QListWidgetItem(self._step_list_text(step))
            self.step_list.addItem(item)
            self._step_items.append(item)
        self.step_list.currentRowChanged.connect(self._on_step_row_changed)
        left_layout.addWidget(self.step_list, 1)
        left.setMaximumWidth(INSTRUCTOR_RAIL_MAX_WIDTH)

        right = QFrame()
        right.setObjectName("activeCard")
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(
            INSTRUCTOR_CARD_MARGIN,
            INSTRUCTOR_CARD_MARGIN,
            INSTRUCTOR_CARD_MARGIN,
            INSTRUCTOR_CARD_MARGIN,
        )
        right_layout.setSpacing(INSTRUCTOR_CARD_SPACING)

        top_content = QWidget()
        top_layout = QVBoxLayout(top_content)
        top_layout.setContentsMargins(*INSTRUCTOR_TOP_LAYOUT_MARGINS)
        top_layout.setSpacing(INSTRUCTOR_CARD_SPACING)

        self.active_title = QLabel()
        self.active_title.setFont(
            QFont(UI_FONT_FAMILY, INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE, QFont.Weight.Bold)
        )
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
        self.primary_action.setObjectName("primaryAction")
        self.primary_action.clicked.connect(self._run_current_step_action)
        top_layout.addWidget(self.primary_action, alignment=Qt.AlignmentFlag.AlignLeft)

        self.step2_action_row = QWidget()
        step2_action_layout = QHBoxLayout(self.step2_action_row)
        step2_action_layout.setContentsMargins(
            INSTRUCTOR_STEP2_ACTION_MARGIN,
            INSTRUCTOR_STEP2_ACTION_MARGIN,
            INSTRUCTOR_STEP2_ACTION_MARGIN,
            INSTRUCTOR_STEP2_ACTION_MARGIN,
        )
        step2_action_layout.setSpacing(INSTRUCTOR_STEP2_ACTION_SPACING)

        self.step2_upload_action = QPushButton()
        self.step2_upload_action.clicked.connect(self._on_step2_upload_clicked)
        step2_action_layout.addWidget(self.step2_upload_action)

        self.step2_prepare_action = QPushButton()
        self.step2_prepare_action.setObjectName("primaryAction")
        self.step2_prepare_action.clicked.connect(self._on_step2_prepare_clicked)
        step2_action_layout.addWidget(self.step2_prepare_action)
        step2_action_layout.addStretch(1)
        top_layout.addWidget(self.step2_action_row)

        self.step3_action_row = QWidget()
        step3_action_layout = QHBoxLayout(self.step3_action_row)
        step3_action_layout.setContentsMargins(
            INSTRUCTOR_STEP2_ACTION_MARGIN,
            INSTRUCTOR_STEP2_ACTION_MARGIN,
            INSTRUCTOR_STEP2_ACTION_MARGIN,
            INSTRUCTOR_STEP2_ACTION_MARGIN,
        )
        step3_action_layout.setSpacing(INSTRUCTOR_STEP2_ACTION_SPACING)

        self.step3_upload_action = QPushButton()
        self.step3_upload_action.clicked.connect(self._on_step3_upload_clicked)
        step3_action_layout.addWidget(self.step3_upload_action)

        self.step3_generate_action = QPushButton()
        self.step3_generate_action.setObjectName("primaryAction")
        self.step3_generate_action.clicked.connect(self._on_step3_generate_clicked)
        step3_action_layout.addWidget(self.step3_generate_action)
        step3_action_layout.addStretch(1)
        top_layout.addWidget(self.step3_action_row)
        top_layout.addStretch(1)

        right_layout.addWidget(top_content, 1)

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
            lambda url: self._on_quick_link_activated(url.toString())
        )
        links_tab_layout.addWidget(self.generated_outputs_view)

        # Retained for test doubles that call _refresh_quick_links directly.
        self.quick_link_labels: dict[str, QLabel] = {}

        self.info_tabs.addTab(log_tab, t("instructor.log.title"))
        self.info_tabs.addTab(links_tab, t(self.RAIL_LINK_TITLE_KEY))
        right_layout.addWidget(self.info_tabs)

        root.addWidget(left)
        root.addWidget(right, 1)

        self.setStyleSheet(INSTRUCTOR_PANEL_STYLESHEET)

    def _quick_link_items(self) -> tuple[tuple[str, str | None], ...]:
        return tuple((label_key, getattr(self, attr)) for label_key, attr in self.RAIL_LINKS)

    def _quick_link_markup(self, label_key: str, path: str | None) -> str:
        label = t(label_key)
        if not path:
            return f"{label}: {t(self.RAIL_LINK_NOT_AVAILABLE_KEY)}"
        file_link = f'<a href="file::{path}">{t(self.RAIL_LINK_OPEN_FILE_KEY)}</a>'
        folder_link = f'<a href="folder::{path}">{t(self.RAIL_LINK_OPEN_FOLDER_KEY)}</a>'
        name = escape(Path(path).name)
        full_path = escape(str(Path(path)))
        return (
            f"<b>{escape(label)}</b>: {name}<br>"
            f"<span>{full_path}</span><br>"
            f"{file_link} | {folder_link}"
        )

    def _quick_links_html(self) -> str:
        rows = [
            f"<div style='margin-bottom:10px'>{self._quick_link_markup(link_key, path)}</div>"
            for link_key, path in self._quick_link_items()
        ]
        return "".join(rows)

    def _on_quick_link_activated(self, href: str) -> None:
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
            t(self.RAIL_LINK_OPEN_FAILED_KEY),
            title=t("instructor.msg.error_title"),
            level="error",
        )

    def _refresh_quick_links(self) -> None:
        generated_outputs_view = getattr(self, "generated_outputs_view", None)
        if generated_outputs_view is not None:
            generated_outputs_view.setHtml(self._quick_links_html())
        for link_key, path in self._quick_link_items():
            link_label = self.quick_link_labels.get(link_key)
            if link_label is None:
                continue
            link_label.setText(self._quick_link_markup(link_key, path))

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
        if row < 0:
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

    def _refresh_ui(self) -> None:
        completed = sum(
            1
            for done, outdated in (
                (self.step1_done, False),
                (self.step2_done, False),
                (self.step4_done, self.step4_outdated),
            )
            if done and not outdated
        )
        self.progress_label.setText(
            t("instructor.progress", completed=completed, total=len(self.WORKFLOW_STEPS))
        )
        self._update_busy_timer_label()

        self.step_list.blockSignals(True)
        for index, item in enumerate(self._step_items, start=1):
            item.setText(self._step_list_text(index))
        self.step_list.setCurrentRow(self.current_step - 1)
        self.step_list.blockSignals(False)

        self.active_title.setText(
            t(
                "instructor.active_title",
                number=self.current_step,
                title=t(self.STEP_TITLE_KEYS[self.current_step]),
            )
        )
        self.active_desc.setText(t(self.STEP_DESC_KEYS[self.current_step]))
        self.info_tabs.setTabText(0, t("instructor.log.title"))
        self.info_tabs.setTabText(1, t(self.RAIL_LINK_TITLE_KEY))
        self._refresh_quick_links()
        can_run, reason = self._can_run_step(self.current_step)
        if self.current_step == 2:
            self.primary_action.setVisible(False)
            self.step2_action_row.setVisible(True)
            self.step3_action_row.setVisible(False)
            self.step2_upload_action.setText(t("instructor.action.step2.upload"))
            self.step2_prepare_action.setText(t("instructor.action.step2.prepare"))
            self.step2_prepare_action.setEnabled(self.step2_upload_ready)
            self.step2_upload_action.setDefault(not self.step2_upload_ready)
            self.step2_prepare_action.setDefault(self.step2_upload_ready)
        elif self.current_step == 3:
            self.primary_action.setVisible(False)
            self.step2_action_row.setVisible(False)
            self.step3_action_row.setVisible(True)
            self.step3_upload_action.setText(
                t("instructor.action.step4.redo")
                if self.step3_done and not self.step3_outdated
                else t("instructor.action.step4.default")
            )
            self.step3_generate_action.setText(
                t("instructor.action.step5.redo")
                if self.step4_done
                else t("instructor.action.step5.default")
            )
            self.step3_upload_action.setEnabled(True)
            self.step3_generate_action.setEnabled(
                self.step3_done and not self.step3_outdated
            )
            self.step3_upload_action.setDefault(not (self.step3_done and not self.step3_outdated))
            self.step3_generate_action.setDefault(self.step3_done and not self.step3_outdated)
        else:
            self.primary_action.setVisible(True)
            self.primary_action.setText(self._action_text_for_step(self.current_step))
            self.step2_action_row.setVisible(False)
            self.step3_action_row.setVisible(False)

        if self.current_step == 2:
            self.step2_upload_action.setEnabled(True)
        elif self.current_step == 3:
            pass
        else:
            self.primary_action.setEnabled(can_run)
        if not can_run:
            self.active_note.setText(reason)
        elif self.step3_outdated or self.step4_outdated:
            if self._step_outdated(self.current_step):
                self.active_note.setText(t("instructor.note.outdated_current"))
            else:
                self.active_note.setText(t("instructor.note.outdated_downstream"))
        else:
            self.active_note.setText(t("instructor.note.default"))

        if self.state.busy:
            self.primary_action.setEnabled(False)
            self.step2_upload_action.setEnabled(False)
            self.step2_prepare_action.setEnabled(False)
            self.step3_upload_action.setEnabled(False)
            self.step3_generate_action.setEnabled(False)

    def retranslate_ui(self) -> None:
        self._refresh_ui()
        self._clear_info_text_selection()

    def _run_current_step_action(self) -> None:
        if self.current_step == 1:
            self._download_course_template_async()
            self._refresh_ui()
            return
        if self.current_step == 2:
            self._upload_course_details_async()
            self._refresh_ui()
            return
        if self.current_step == 3:
            self._generate_final_report_async()
            self._refresh_ui()
            return
        self._refresh_ui()

    def _on_step2_upload_clicked(self) -> None:
        self._upload_course_details_async()
        self._refresh_ui()

    def _on_step2_prepare_clicked(self) -> None:
        self._prepare_marks_template_async()
        self._refresh_ui()

    def _on_step3_upload_clicked(self) -> None:
        self._upload_filled_marks_async()
        self._refresh_ui()

    def _on_step3_generate_clicked(self) -> None:
        self._generate_final_report_async()
        self._refresh_ui()

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
        self._ui_log_handler = _UILogHandler(self._append_user_log)
        _logger.addHandler(self._ui_log_handler)
        self._append_user_log(t("instructor.log.ready"))

    def _append_user_log(self, message: str) -> None:
        if not message or not message.strip():
            return
        text = message.strip()
        if not (
            len(text) >= 10
            and text[0] == "["
            and text[3] == ":"
            and text[6] == ":"
            and text[9] == "]"
        ):
            text = f"[{datetime.now().strftime('%H:%M:%S')}] {text}"
        self.user_log_view.appendPlainText(text)

    def _publish_status(self, message: str) -> None:
        self._append_user_log(message)
        emit_user_status(getattr(self, "status_changed", None), message, logger=_logger)

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
