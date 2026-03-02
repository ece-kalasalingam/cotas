"""Course Instructor CO module UI (list-based step navigation)."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
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
    UI_FONT_FAMILY,
)
from common.exceptions import AppSystemError, ValidationError
from common.toast import show_toast
from common.texts import t
from common.utils import (
    remember_dialog_dir,
    remember_dialog_dir_safe,
    resolve_dialog_start_path,
)
from modules.instructor import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)

_logger = logging.getLogger(__name__)


class InstructorModule(QWidget):
    """Simple wizard-like UI for CO score workflow."""

    status_changed = Signal(str)
    WORKFLOW_STEPS = (1, 2, 3, 4)

    STEP_TITLE_KEYS = {
        1: "instructor.step1.title",
        2: "instructor.step2.title",
        3: "instructor.step4.title",
        4: "instructor.step5.title",
    }

    STEP_DESC_KEYS = {
        1: "instructor.step1.desc",
        2: "instructor.step2.desc",
        3: "instructor.step4.desc",
        4: "instructor.step5.desc",
    }

    PATH_ATTRS = {
        1: "step1_path",
        2: "step2_path",
        3: "step3_path",
        4: "step4_path",
    }

    DONE_ATTRS = {
        1: "step1_done",
        2: "step2_done",
        3: "step3_done",
        4: "step4_done",
    }

    OUTDATED_ATTRS = {
        3: "step3_outdated",
        4: "step4_outdated",
    }

    ACTION_DEFAULT_KEYS = {
        1: "instructor.action.step1.default",
        2: "instructor.action.step2.default",
        3: "instructor.action.step4.default",
        4: "instructor.action.step5.default",
    }

    ACTION_REDO_KEYS = {
        1: "instructor.action.step1.redo",
        2: "instructor.action.step2.redo",
        3: "instructor.action.step4.redo",
        4: "instructor.action.step5.redo",
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

        self._step_items: list[QListWidgetItem] = []
        self._build_ui()
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

        self.step_list = QListWidget()
        self.step_list.setObjectName("stepList")
        self.step_list.setFrameShape(QFrame.Shape.NoFrame)
        self.step_list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.step_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.step_list.setCursor(Qt.CursorShape.ArrowCursor)
        self.step_list.setSpacing(INSTRUCTOR_STEP_LIST_SPACING)
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

        self.active_title = QLabel()
        self.active_title.setFont(
            QFont(UI_FONT_FAMILY, INSTRUCTOR_ACTIVE_TITLE_FONT_SIZE, QFont.Weight.Bold)
        )
        right_layout.addWidget(self.active_title)

        self.active_desc = QLabel()
        self.active_desc.setWordWrap(True)
        right_layout.addWidget(self.active_desc)

        self.active_file = QLabel(t("instructor.file.none"))
        self.active_file.setWordWrap(True)
        right_layout.addWidget(self.active_file)

        self.active_note = QLabel(t("instructor.note.default"))
        self.active_note.setWordWrap(True)
        self.active_note.setObjectName("hintText")
        right_layout.addWidget(self.active_note)

        self.primary_action = QPushButton()
        self.primary_action.setObjectName("primaryAction")
        self.primary_action.clicked.connect(self._run_current_step_action)
        right_layout.addWidget(self.primary_action, alignment=Qt.AlignmentFlag.AlignLeft)

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
        self.step2_upload_action.setObjectName("normalAction")
        self.step2_upload_action.clicked.connect(self._on_step2_upload_clicked)
        step2_action_layout.addWidget(self.step2_upload_action)

        self.step2_prepare_action = QPushButton()
        self.step2_prepare_action.setObjectName("primaryAction")
        self.step2_prepare_action.clicked.connect(self._on_step2_prepare_clicked)
        step2_action_layout.addWidget(self.step2_prepare_action)
        step2_action_layout.addStretch(1)
        right_layout.addWidget(self.step2_action_row)
        right_layout.addStretch(1)

        root.addWidget(left)
        root.addWidget(right, 1)

        self.setStyleSheet(INSTRUCTOR_PANEL_STYLESHEET)

    def _step_path(self, step: int) -> str | None:
        return getattr(self, self.PATH_ATTRS[step])

    def _step_done(self, step: int) -> bool:
        return bool(getattr(self, self.DONE_ATTRS[step]))

    def _step_outdated(self, step: int) -> bool:
        outdated_attr = self.OUTDATED_ATTRS.get(step)
        return bool(getattr(self, outdated_attr)) if outdated_attr else False

    def _step_state_text(self, step: int) -> str:
        done = self._step_done(step)
        outdated = self._step_outdated(step)
        if done and outdated:
            return t("instructor.badge.needs_update")
        return t("instructor.badge.done") if done else t("instructor.badge.pending")

    def _step_list_text(self, step: int) -> str:
        title = t(self.STEP_TITLE_KEYS[step])
        state = self._step_state_text(step)
        return f"{step}. {title}  {state}"

    def _file_text_for_step(self, step: int) -> str:
        if step == 2:
            if self.step2_path:
                return Path(self.step2_path).name
            if self.step2_course_details_path:
                return Path(self.step2_course_details_path).name
        path = self._step_path(step)
        return Path(path).name if path else t("instructor.file.none")

    def _action_text_for_step(self, step: int) -> str:
        return t(
            self.ACTION_REDO_KEYS[step]
            if self._step_done(step)
            else self.ACTION_DEFAULT_KEYS[step]
        )

    def _can_run_step(self, step: int) -> tuple[bool, str]:
        if step in (1, 2):
            return True, ""
        if step == 3 and not self.step2_done:
            return False, t("instructor.require.step2")
        if step == 4 and (not self.step3_done or self.step3_outdated):
            return False, t("instructor.require.step3")
        return True, ""

    def _on_step_selected(self, step: int) -> None:
        self.current_step = step
        self._refresh_ui()

    def _on_step_row_changed(self, row: int) -> None:
        if row < 0:
            return
        self._on_step_selected(row + 1)

    def _refresh_ui(self) -> None:
        completed = sum(
            1
            for done, outdated in (
                (self.step1_done, False),
                (self.step2_done, False),
                (self.step3_done, self.step3_outdated),
                (self.step4_done, self.step4_outdated),
            )
            if done and not outdated
        )
        self.progress_label.setText(
            t("instructor.progress", completed=completed, total=len(self.WORKFLOW_STEPS))
        )

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
        self.active_file.setText(self._file_text_for_step(self.current_step))
        if self.current_step == 2:
            self.primary_action.setVisible(False)
            self.step2_action_row.setVisible(True)
            self.step2_upload_action.setText(t("instructor.action.step2.upload"))
            self.step2_prepare_action.setText(t("instructor.action.step2.prepare"))
            self.step2_prepare_action.setEnabled(self.step2_upload_ready)
            self.step2_upload_action.setDefault(not self.step2_upload_ready)
            self.step2_prepare_action.setDefault(self.step2_upload_ready)
        else:
            self.primary_action.setVisible(True)
            self.primary_action.setText(self._action_text_for_step(self.current_step))
            self.step2_action_row.setVisible(False)

        can_run, reason = self._can_run_step(self.current_step)
        if self.current_step == 2:
            self.step2_upload_action.setEnabled(True)
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

    def _run_current_step_action(self) -> None:
        actions = {
            1: self._download_course_template,
            2: self._upload_course_details,
            3: self._upload_filled_marks,
            4: self._generate_final_report,
        }
        action = actions[self.current_step]
        action()
        self._refresh_ui()

    def _on_step2_upload_clicked(self) -> None:
        self._upload_course_details()
        self._refresh_ui()

    def _on_step2_prepare_clicked(self) -> None:
        self._prepare_marks_template()
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

    def _show_step_success_toast(self, step: int) -> None:
        show_toast(
            self,
            t("instructor.msg.step_completed", step=step, title=t(self.STEP_TITLE_KEYS[step])),
            title=t("instructor.msg.success_title"),
            level="success",
        )

    def _show_validation_error_toast(self, message: str) -> None:
        show_toast(
            self,
            message,
            title=t("instructor.msg.validation_title"),
            level="error",
        )

    def _show_system_error_toast(self, step: int) -> None:
        show_toast(
            self,
            t("instructor.msg.failed_to_do", action=t(self.STEP_TITLE_KEYS[step])),
            title=t("instructor.msg.error_title"),
            level="error",
        )

    def _download_course_template(self) -> None:
        template_id = ID_COURSE_SETUP

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            t("instructor.dialog.step1.title"),
            resolve_dialog_start_path(APP_NAME, t("instructor.dialog.step1.default_name")),
            t("instructor.dialog.filter.excel"),
        )
        if not save_path:
            return

        try:
            generate_course_details_template(save_path, template_id=template_id)
        except ValidationError as exc:
            _logger.warning("Template generation validation failed: %s", exc)
            self._show_validation_error_toast(str(exc))
            return
        except AppSystemError:
            _logger.exception("System failure while generating course details template.")
            self._show_system_error_toast(1)
            return
        except Exception:
            _logger.exception("Unexpected error while generating course details template.")
            self._show_system_error_toast(1)
            return

        self.step1_path = save_path
        self.step1_done = True
        self._remember_dialog_dir_safe(save_path)
        self.status_changed.emit(t("instructor.status.step1_selected"))
        self._show_step_success_toast(1)

    def _upload_course_details(self) -> None:
        open_path, _ = QFileDialog.getOpenFileName(
            self,
            t("instructor.dialog.step2.title"),
            resolve_dialog_start_path(APP_NAME),
            t("instructor.dialog.filter.excel_open"),
        )
        if not open_path:
            return

        try:
            validate_course_details_workbook(open_path)
        except ValidationError as exc:
            _logger.warning("Step 2 workbook validation failed: %s", exc)
            self._show_validation_error_toast(str(exc))
            return
        except AppSystemError:
            _logger.exception("System failure while validating Step 2 workbook.")
            self._show_system_error_toast(2)
            return
        except Exception:
            _logger.exception("Unexpected error while validating Step 2 workbook.")
            self._show_system_error_toast(2)
            return

        replacing = self.step2_done or self.step2_upload_ready
        self.step2_course_details_path = open_path
        self.step2_upload_ready = True
        self.step2_done = False
        self.step2_path = None
        self._remember_dialog_dir_safe(open_path)

        if replacing:
            self.step4_outdated = self.step4_done
            self.step3_outdated = self.step3_done
            if self.step3_outdated or self.step4_outdated:
                self.status_changed.emit(t("instructor.status.step2_changed"))
        else:
            self.status_changed.emit(t("instructor.status.step2_validated"))
        self._show_step_success_toast(2)

    def _prepare_marks_template(self) -> None:
        if not self.step2_upload_ready or not self.step2_course_details_path:
            show_toast(
                self,
                t("instructor.require.step2"),
                title=t("instructor.msg.step_required_title"),
                level="info",
            )
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            t("instructor.dialog.step3.title"),
            resolve_dialog_start_path(APP_NAME, t("instructor.dialog.step3.default_name")),
            t("instructor.dialog.filter.excel"),
        )
        if not save_path:
            return

        try:
            generate_marks_template_from_course_details(self.step2_course_details_path, save_path)
        except ValidationError as exc:
            _logger.warning("Step 2 marks template generation validation failed: %s", exc)
            self._show_validation_error_toast(str(exc))
            return
        except AppSystemError:
            _logger.exception("System failure while generating Step 2 marks template.")
            self._show_system_error_toast(2)
            return
        except Exception:
            _logger.exception("Unexpected error while generating Step 2 marks template.")
            self._show_system_error_toast(2)
            return

        self.step2_path = save_path
        self.step2_done = True
        self.step3_outdated = self.step3_done
        self.step4_outdated = self.step4_done
        self._remember_dialog_dir_safe(save_path)
        self.status_changed.emit(t("instructor.status.step2_uploaded"))
        self._show_step_success_toast(2)

    def _upload_filled_marks(self) -> None:
        open_path, _ = QFileDialog.getOpenFileName(
            self,
            t("instructor.dialog.step3.title"),
            resolve_dialog_start_path(APP_NAME),
            t("instructor.dialog.filter.excel_open"),
        )
        if not open_path:
            return

        replacing = self.step3_done
        self.step3_path = open_path
        self.step3_done = True
        self.step3_outdated = False
        self._remember_dialog_dir_safe(open_path)

        if replacing and self.step3_done:
            self.step4_outdated = True
            self.status_changed.emit(t("instructor.status.step3_changed"))
        else:
            self.status_changed.emit(t("instructor.status.step3_uploaded"))
        self._show_step_success_toast(3)

    def _generate_final_report(self) -> None:
        can_run, reason = self._can_run_step(4)
        if not can_run:
            show_toast(
                self,
                reason,
                title=t("instructor.msg.step_required_title"),
                level="info",
            )
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            t("instructor.dialog.step4.title"),
            resolve_dialog_start_path(APP_NAME, t("instructor.dialog.step4.default_name")),
            t("instructor.dialog.filter.excel"),
        )
        if not save_path:
            return

        if not self.step3_path or not Path(self.step3_path).exists():
            _logger.warning("Step 4 failed: Step 3 file is missing. step3_path=%s", self.step3_path)
            show_toast(
                self,
                t("instructor.require.step3"),
                title=t("instructor.msg.step_required_title"),
                level="error",
            )
            return

        try:
            shutil.copyfile(self.step3_path, save_path)
        except OSError:
            _logger.exception(
                "Failed to generate final report by copying filled marks. src=%s dst=%s",
                self.step3_path,
                save_path,
            )
            self._show_system_error_toast(4)
            return

        self.step4_path = save_path
        self.step4_done = True
        self.step4_outdated = False
        self._remember_dialog_dir_safe(save_path)
        self.status_changed.emit(t("instructor.status.step4_selected"))
        self._show_step_success_toast(4)
