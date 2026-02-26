import json
import os
import sys
from datetime import datetime
from typing import Optional
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QPlainTextEdit
    , QGroupBox, QHBoxLayout, QSizePolicy, QFrame, QApplication
)
from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QFont
from scripts.data_providers import UniversalDataProvider
from scripts.utils import ToastNotification, resource_path
from scripts.workbook_renderer import UniversalWorkbookRenderer
from scripts.engine import UniversalEngine
from scripts import blueprints
from scripts import constants


class COModule(QWidget):
    status_changed = Signal(str)
    def __init__(self):
        super().__init__()

        if getattr(sys, 'frozen', False):
            self.last_dir = os.path.dirname(sys.executable)
        else:
            self.last_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.last_dir = os.path.normpath(self.last_dir)

        self.setup_path: Optional[str] = None
        self.gen_marks_path: Optional[str] = None
        self.filled_path: Optional[str] = None
        self._course_default_prefix: str = "Course"

        self.engine = UniversalEngine(blueprints.BLUEPRINT_REGISTRY)
        self.renderer = UniversalWorkbookRenderer()

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)

        # Title Section
        title = QLabel("Course Instructor CO Score Calculation")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)

        # Step 1 Download Template
        self.template_group = QGroupBox("1. Download Course Details Template")
        template_1ay = QHBoxLayout(self.template_group)

        self.template_label = QLabel("Download the course details excel template to get started.")
        self.template_label.setWordWrap(True)
        self.template_button = QPushButton("Download Template")

        template_1ay.addWidget(self.template_label,1)
        template_1ay.addWidget(self.template_button)

        layout.addWidget(self.template_group)

        # Step 2 Upload
        self.setup_group = QGroupBox("2. Upload Course Details")
        setup_1ay = QHBoxLayout(self.setup_group)

        self.setup_label = QLabel("No file selected")
        self.setup_label.setWordWrap(True)
        self.setup_button = QPushButton("Upload Course Details")

        setup_1ay.addWidget(self.setup_label,1)
        setup_1ay.addWidget(self.setup_button)

        layout.addWidget(self.setup_group)

        # Step 3 Generate Template
        self.gen_group = QGroupBox("3. Generate Marks Template")
        gen_lay = QHBoxLayout(self.gen_group)

        self.gen_label = QLabel("")
        self.gen_label.setWordWrap(True)
        self.gen_button = QPushButton("Generate Marks Template")

        gen_lay.addWidget(self.gen_label,1)
        gen_lay.addWidget(self.gen_button)
        layout.addWidget(self.gen_group)

        # Step 4 Filled Marks
        self.filled_group = QGroupBox("4. Upload Filled Marks")
        filled_lay = QHBoxLayout(self.filled_group)

        self.filled_label = QLabel("Waiting for step 2...")
        self.filled_label.setWordWrap(True)
        self.filled_button = QPushButton("Upload Filled Marks")
        filled_lay.addWidget(self.filled_label, 1)
        filled_lay.addWidget(self.filled_button)
        layout.addWidget(self.filled_group)

        # Action Button Compute
        self.submit_button = QPushButton("Compute CO Scores")
        self.submit_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.submit_button.setEnabled(False)
        layout.addWidget(self.submit_button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # Log
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setVisible(False)
        self.log.setMaximumHeight(150)
        self.log.setStyleSheet("font-family: Consolas; border: none;")
        layout.addWidget(self.log)

        # Styling & Stretch
        for b in (self.template_button,self.setup_button, self.gen_button, self.filled_button, self.submit_button):
            b.setMinimumWidth(150)
            b.setMinimumHeight(40)
            b.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            b.setCursor(Qt.CursorShape.PointingHandCursor)

        layout.addStretch()

        # Connect signals
        
        self.template_button.clicked.connect(self.generate_setup_template)
        self.setup_button.clicked.connect(self.pick_setup)
        self.gen_button.clicked.connect(self.generate_marks_template)
        self.filled_button.clicked.connect(self.pick_filled)
        self.submit_button.clicked.connect(self.compute_results)

        self._refresh_actions()

    # -----------------------------
    # Actions
    # -----------------------------

    def log_msg(self, text: str, level: str = "INFO") -> None:
        if not self.log.isVisible(): self.log.setVisible(True)
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.appendPlainText(f"[{ts}] {level}: {text}")

    def clear_log(self) -> None:
        self.log.clear()
        self.log.setVisible(False)

    def _refresh_actions(self) -> None:
        has_setup = bool(self.setup_path)
        has_filled = bool(self.filled_path)
        
        self.gen_button.setEnabled(has_setup)
        self.filled_button.setEnabled(has_setup)
        self.submit_button.setEnabled(has_setup and has_filled)
        self.setup_label.setEnabled(has_setup)
        self.filled_label.setEnabled(has_filled)

    def _set_status(self, text: str) -> None:
        self.status_changed.emit(text)

    def _update_labels(self) -> None:
        if self.setup_path:
            self.setup_label.setText(self.setup_path)
            if self.filled_path:
                self.filled_label.setText(self.filled_path)
            else:
                self.filled_label.setText("Waiting for filled marks file...")
        else:
            self.setup_label.setText("Waiting for course details file...")
        if self.gen_marks_path:
            self.gen_label.setText(self.gen_marks_path)

    def is_file_writable(self, filepath: str) -> bool:
        if not os.path.exists(filepath):
            return True # File doesn't exist, so it's writable
        try:
            # Try to open the file in append mode to see if OS allows it
            with open(filepath, 'a'):
                pass
            return True
        except OSError:
            return False
        

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