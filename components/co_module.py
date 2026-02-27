# components/co_module.py

import os
from datetime import datetime
from typing import Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QPlainTextEdit, QGroupBox,
    QHBoxLayout, QSizePolicy, QFrame
)
from PySide6.QtCore import Signal, Qt, QThread
from PySide6.QtGui import QFont

from scripts.data_providers import UniversalDataProvider
from scripts.utils import (
    ToastNotification,
    get_run_dir,
    load_last_dir,
    save_last_dir
)
from scripts.workbook_renderer import UniversalWorkbookRenderer
from scripts.engine import UniversalEngine
from scripts import blueprints, constants
from scripts.exceptions import ValidationError, SystemError


# =========================================================
# Worker Thread
# =========================================================

class EngineWorker(QThread):
    finished = Signal(bool, str)
    status_update = Signal(str)

    def __init__(self, engine, renderer, action_type, **kwargs):
        super().__init__()
        self.engine = engine
        self.renderer = renderer
        self.action_type = action_type
        self.kwargs = kwargs

    def run(self):
        try:
            if self.action_type == "GENERATE_SETUP":
                self.status_update.emit("Generating Course Setup Template...")
                data = UniversalDataProvider.get_data_for_template(
                    constants.ID_COURSE_SETUP
                )
                context = {"type_id": constants.ID_COURSE_SETUP}

                self.renderer.render(
                    blueprints.COURSE_SETUP_BP,
                    self.kwargs["path"],
                    data,
                    fingerprint_context=context
                )

                self.finished.emit(True, "Setup Template Generated.")

            elif self.action_type == "LOAD_SETUP":
                self.status_update.emit("Validating Setup File structure...")
                success = self.engine.load_from_file(self.kwargs['path'])
                msg = "Ready" if success else "\n".join(self.engine.errors)
                self.finished.emit(success, msg)

            elif self.action_type == "LOAD_MARKS":
                self.status_update.emit("Validating Marks File with Setup context...")
                success = self.engine.load_with_external(
                    self.kwargs['path'],
                    self.kwargs['external_path']
                )
                msg = "Ready" if success else "\n".join(self.engine.errors)
                self.finished.emit(success, msg)

            elif self.action_type == "GENERATE_MARKS":
                self.status_update.emit("Preparing Marks Entry Template...")
                # TODO: Implement marks template generation
                self.finished.emit(True, "Marks Template Generated.")

        except ValidationError as e:
            self.finished.emit(False, str(e))

        except SystemError as e:
            self.finished.emit(False, f"System Error: {e}")

        except Exception as e:
            self.finished.emit(False, f"Unexpected Error: {e}")


# =========================================================
# Main UI Module
# =========================================================

class COModule(QWidget):
    status_changed = Signal(str)

    def __init__(self):
        super().__init__()

        # Initialize worker reference (CRITICAL)
        self.worker: Optional[EngineWorker] = None

        run_dir = get_run_dir()
        stored_dir = load_last_dir()

        if stored_dir and os.path.isdir(stored_dir):
            self.last_dir = os.path.normpath(stored_dir)
        else:
            self.last_dir = os.path.normpath(run_dir)

        self.setup_path: Optional[str] = None
        self.filled_path: Optional[str] = None
        self.gen_marks_path: Optional[str] = None

        self.engine = UniversalEngine(blueprints.BLUEPRINT_REGISTRY)
        self.renderer = UniversalWorkbookRenderer()

        self._build_ui()

    # =====================================================
    # UI BUILD
    # =====================================================

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)

        title = QLabel("Course Instructor CO Score Calculation")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)

        # Step 1
        self.template_group = QGroupBox("1. Download Course Details Template")
        t_layout = QHBoxLayout(self.template_group)

        self.template_label = QLabel(
            "Download the course details excel template to get started."
        )
        self.template_label.setWordWrap(True)

        self.template_button = QPushButton("Download Template")

        t_layout.addWidget(self.template_label, 1)
        t_layout.addWidget(self.template_button)
        layout.addWidget(self.template_group)

        # Step 2
        self.setup_group = QGroupBox("2. Upload Course Details")
        s_layout = QHBoxLayout(self.setup_group)

        self.setup_label = QLabel("No file selected")
        self.setup_label.setWordWrap(True)
        self.setup_button = QPushButton("Upload Course Details")

        s_layout.addWidget(self.setup_label, 1)
        s_layout.addWidget(self.setup_button)
        layout.addWidget(self.setup_group)

        # Step 3
        self.gen_group = QGroupBox("3. Generate Marks Template")
        g_layout = QHBoxLayout(self.gen_group)

        self.gen_label = QLabel("")
        self.gen_label.setWordWrap(True)
        self.gen_button = QPushButton("Generate Marks Template")

        g_layout.addWidget(self.gen_label, 1)
        g_layout.addWidget(self.gen_button)
        layout.addWidget(self.gen_group)

        # Step 4
        self.filled_group = QGroupBox("4. Upload Filled Marks")
        f_layout = QHBoxLayout(self.filled_group)

        self.filled_label = QLabel("Waiting for step 2...")
        self.filled_label.setWordWrap(True)
        self.filled_button = QPushButton("Upload Filled Marks")

        f_layout.addWidget(self.filled_label, 1)
        f_layout.addWidget(self.filled_button)
        layout.addWidget(self.filled_group)

        # Compute Button
        self.submit_button = QPushButton("Compute CO Scores")
        self.submit_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.submit_button.setEnabled(False)
        layout.addWidget(self.submit_button, alignment=Qt.AlignmentFlag.AlignCenter)

        # Log
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setVisible(False)
        self.log.setMaximumHeight(150)
        self.log.setStyleSheet("font-family: Courier New; border: none; font-size: 14px;")
        layout.addWidget(self.log)

        for b in (
            self.template_button,
            self.setup_button,
            self.gen_button,
            self.filled_button,
            self.submit_button
        ):
            b.setMinimumWidth(150)
            b.setMinimumHeight(40)
            b.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            b.setCursor(Qt.CursorShape.PointingHandCursor)

        layout.addStretch()

        self.template_button.clicked.connect(self.step_1_generate_setup)
        self.setup_button.clicked.connect(self.step_2_upload_setup)
        self.gen_button.clicked.connect(self.step_3_generate_marks)
        self.filled_button.clicked.connect(self.step_4_upload_marks)
        self.submit_button.clicked.connect(self.step_5_compute_attainment)

        self._refresh_actions()

    # =====================================================
    # Actions
    # =====================================================

    def step_1_generate_setup(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Create Course Details Template",
            os.path.join(self.last_dir, "Course_Setup_Input_Template.xlsx"),
            "Excel (*.xlsx)"
        )
        if path:
            self._run_worker("GENERATE_SETUP", path=path)

    def step_2_upload_setup(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Filled Course Details",
            self.last_dir,
            "Excel (*.xlsx)"
        )
        if path:
            self._run_worker("LOAD_SETUP", path=path)

    def step_3_generate_marks(self):
        if not self.setup_path:
            self.log_msg("Validate Setup File first.", "ERROR")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Marks Template",
            self.last_dir,
            "Excel (*.xlsx)"
        )
        if path:
            self._run_worker("GENERATE_MARKS", path=path)

    def step_4_upload_marks(self):
        if not self.setup_path:
            self.log_msg("Please load Setup File first.", "ERROR")
            return

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Filled Marks File",
            self.last_dir,
            "Excel (*.xlsx)"
        )
        if path:
            self._run_worker("LOAD_MARKS", path=path, external_path=self.setup_path)

    def step_5_compute_attainment(self):
        if not self.setup_path or not self.filled_path:
            self.log_msg("Ensure both Setup and Marks files are loaded.", "ERROR")
            return

        self.log_msg("Computing CO Totals...")

    # =====================================================
    # Thread Control
    # =====================================================
    def _run_worker(self, action, **kwargs):
        self.set_ui_enabled(False)
        self.worker = EngineWorker(self.engine, self.renderer, action, **kwargs)
        self.worker.status_update.connect(self._set_status)
        self.worker.finished.connect(lambda s, m: self.on_task_finished(s, m, action, kwargs.get('path')))
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.start()

    def on_task_finished(self, success, message, action, path):

        self.set_ui_enabled(True)

        if success:
            if action == "LOAD_SETUP":
                self.setup_path = path
                self._update_working_dir(path)
                self.log_msg(f"Uploaded details file: {path}", "SYSTEM")

            elif action == "LOAD_MARKS":
                self.filled_path = path
                self._update_working_dir(path)
                self.log_msg(f"Uploaded marks file: {path}", "SYSTEM")

            elif action == "GENERATE_SETUP":
                self._update_working_dir(path)
                ToastNotification(self.window(), "Template generated!", type="success")
                self.log_msg(f"Setup template created: {path}", "SYSTEM")

            self._update_labels()
            self._refresh_actions()
        else:
            ToastNotification(self.window(), "Failed to process action.", type="error")
            self.log_msg(message, "ERROR")

        self._set_status("Ready" if success else "Error")

    # =====================================================
    # Utilities
    # =====================================================

    def _update_working_dir(self, path: str):
        folder = os.path.dirname(path)
        self.last_dir = os.path.normpath(folder)
        save_last_dir(self.last_dir)

    def set_ui_enabled(self, enabled: bool):
        for btn in (
            self.template_button,
            self.setup_button,
            self.gen_button,
            self.filled_button,
            self.submit_button
        ):
            btn.setEnabled(enabled)

    def log_msg(self, text: str, level: str = "INFO"):
        if not self.log.isVisible():
            self.log.setVisible(True)

        ts = datetime.now().strftime("%H:%M:%S")
        self.log.appendPlainText(f"[{ts}] {level}: {text}")

    def _refresh_actions(self):
        has_setup = bool(self.setup_path)
        has_filled = bool(self.filled_path)

        self.gen_button.setEnabled(has_setup)
        self.filled_button.setEnabled(has_setup)
        self.submit_button.setEnabled(has_setup and has_filled)

    def _set_status(self, text: str):
        self.status_changed.emit(text)

    def _update_labels(self):
        self.setup_label.setText(
            self.setup_path if self.setup_path else "Waiting for course details file..."
        )
        self.filled_label.setText(
            self.filled_path if self.filled_path else "Waiting for filled marks file..."
        )
"""
    def pick_setup(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Course Details", self.last_dir, "Excel (*.xlsx)"
        )
        if not path: return

        self.setup_path = path
        self.setup_label.setText(os.path.basename(path))
        self.last_dir = os.path.dirname(path)
        self.log_msg(f"Analyzing {os.path.basename(path)}...", "SYSTEM")
        success = self.engine.load_from_file(path)
        if success:
            self.log_msg(f"Uploaded details file: {path}", "SYSTEM")
            self._update_labels()
            #self.compute_default_prefix_from_setup()
            self._set_status("Setup loaded")
            self._refresh_actions()
        else:
            self.log_msg(f"Errors loading file: {self.engine.errors}", "ERROR")
            self._set_status("Error loading setup. See log for details.")
            ToastNotification(self.window(), "Error loading setup. See log for details.", type="error")


    def pick_filled(self):
        specific_filename = f"{self._course_default_prefix}_Marks_Template.xlsx"
        filter_str = f"Specific File (*{specific_filename})"
        path, _ = QFileDialog.getOpenFileName(self, "Select Filled Marks Excel", self.last_dir, filter_str)

        if path:
            self.filled_path = path
            self.last_dir = os.path.dirname(path)
            self.log_msg(f"Uploaded filled marks file: {path}", "SYSTEM")
            self._update_labels()
            self._set_status("Marks loaded")
            self._refresh_actions()
    
    def generate_setup_template(self):
        try:
            tid = constants.ID_COURSE_SETUP
            bp = blueprints.BLUEPRINT_REGISTRY.get(tid)
            if not bp:
                raise SystemError(f"Blueprint {tid} not found in registry.")
            # 1. Get pre-filled data
            data = UniversalDataProvider.get_data_for_template("COURSE_SETUP_V1")
            # 2. Ask for Save Location
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Create Course Details Template",
                os.path.join(self.last_dir, "Course_Setup_Input_Template.xlsx"),
                "Excel (*.xlsx)"
            )
            if not path: return

            # 3. UI Update
            self.last_dir = os.path.dirname(path)
            self._set_status("Generating Template...")
            QApplication.processEvents()

            # 4. Render
            self.renderer.render(bp, path, data, fingerprint_context={"type_id": tid})

            # 5. Success
            self.log_msg(f"Setup template created: {path}", "SYSTEM")
            self._set_status("Template generated successfully")
            ToastNotification(self.window(), "Template generated successfully!", type="success")

        except Exception as e:
            self.log_msg(f"Setup Generation Error: {str(e)}", "ERROR")
            self._set_status("Error occurred.")
            ToastNotification(self.window(), "Generation Failed", type="error")

    def generate_marks_template(self):
        if not self.setup_path:
            self.log_msg("Upload course details file first.", "ERROR")
            self._set_status("Error")
            return

        try:
            # 1️⃣ DATA PREPARATION (Do this BEFORE asking where to save)
            self._set_status("Processing Data...")
            QApplication.processEvents()


        except Exception as e:
            self.log_msg(f"Marks Template Generation Error: {str(e)}", "ERROR")
            self._set_status("Error. See the log for details.")
            ToastNotification(self.window(), "Error. See the log for details.", type="error")
    
    def compute_results(self):
        if not self.setup_path or not self.filled_path:
            self.log_msg("Upload required files first.", "ERROR")
            return

        default_out = os.path.join(
            self.last_dir,
            f"{self._course_default_prefix}_CO_Results.xlsx"
        )

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Results", default_out, "Excel (*.xlsx)"
        )

        if not path:
            return

        if not self.is_file_writable(path):
            self._set_status("File is Open!")
            self.log_msg(
                f"Cannot save to {os.path.basename(path)}. Close it in Excel first.",
                "WARNING"
            )
            return

        print(f"to do: compute results and save to {path}")

        """
