import os
import sys
from typing import Optional
from datetime import datetime
from wsgiref.validate import validator
from PySide6.QtWidgets import (
    QApplication, QFrame, QSizePolicy, QWidget, QVBoxLayout, QLabel, QPushButton,
    QGroupBox, QHBoxLayout, QFileDialog,
    QPlainTextEdit
)
from PySide6.QtGui import QFont
from PySide6.QtCore import QTimer, Qt, Signal

from core.loader import SetupLoader
from core.generator import MarksWorkbookGenerator
from core.calculator import COCalculator
from core.setup_validator import SetupValidator
from core.setup_template import SetupTemplateGenerator


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
        self._setup_loader: Optional[SetupLoader] = None
        self._course_default_prefix: str = "Course"
        self._status_timer = QTimer(self)
        self._status_timer.setSingleShot(True)
        self._status_timer.timeout.connect(self._reset_status)

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)

        # Title Section
        title = QLabel("CO Attainment Workflow")
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
        self.setup_button = QPushButton("Browse")

        setup_1ay.addWidget(self.setup_label,1)
        setup_1ay.addWidget(self.setup_button)

        layout.addWidget(self.setup_group)

        # Step 3 Generate Template
        self.gen_group = QGroupBox("3. Generate Marks Template")
        gen_lay = QHBoxLayout(self.gen_group)

        self.gen_label = QLabel("")
        self.gen_label.setWordWrap(True)
        self.gen_button = QPushButton("Generate")

        gen_lay.addWidget(self.gen_label,1)
        gen_lay.addWidget(self.gen_button)
        layout.addWidget(self.gen_group)

        # Step 4 Filled Marks
        self.filled_group = QGroupBox("4. Upload Filled Marks")
        filled_lay = QHBoxLayout(self.filled_group)

        self.filled_label = QLabel("Waiting for step 2...")
        self.filled_label.setWordWrap(True)
        self.filled_button = QPushButton("Browse")
        filled_lay.addWidget(self.filled_label, 1)
        filled_lay.addWidget(self.filled_button)
        layout.addWidget(self.filled_group)

        # Action Button Compute
        self.submit_button = QPushButton("Compute CO Results")
        self.submit_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.submit_button.setEnabled(False)
        layout.addWidget(self.submit_button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # Log
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setVisible(False)
        self.log.setMaximumHeight(150)
        layout.addWidget(self.log)

        # Styling & Stretch
        for b in (self.setup_button, self.gen_button, self.filled_button, self.submit_button):
            b.setMinimumWidth(120)
            b.setMinimumHeight(40)
            b.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            b.setCursor(Qt.CursorShape.PointingHandCursor)

        layout.addStretch()

        # Connect signals
        
        self.template_button.clicked.connect(self.generate_setup_template)
        self.setup_button.clicked.connect(self.pick_setup)
        self.gen_button.clicked.connect(self.generate_template)
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
    
    def _reset_status(self) -> None:
        self.status_changed.emit("Idle")

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
        
    def compute_default_prefix_from_setup(self) -> None:
        try:
            if not self.setup_path: return
            self._setup_loader = SetupLoader(self.setup_path)
            meta = self._setup_loader.load_metadata()
            self._course_default_prefix = f"{meta.course_code}_{meta.section}_{meta.semester}_{meta.academic_year}"
        except:
            self._course_default_prefix = "Course"

    def pick_setup(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Course Details", self.last_dir, "Excel (*.xlsx)"
        )
        if path:
            self.setup_path = path
            self.setup_label.setText(os.path.basename(path))
            self.last_dir = os.path.dirname(path)
            self.log_msg(f"Uploaded details file: {path}", "SYSTEM")
            self._update_labels()
            self.compute_default_prefix_from_setup()
            self._set_status("Setup loaded")
            self._refresh_actions()

    def generate_template(self):
        if not self.setup_path:
            self.log_msg("Upload course details file first.", "ERROR")
            self._set_status("Error")
            return

        default_file = os.path.join(self.last_dir, f"{self._course_default_prefix}_Marks_Template.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, "Save Template", default_file, "Excel (*.xlsx)")
        
        if not path:
            return

        if path:
            if not self.is_file_writable(path):
                self._set_status("File is Open!")
                self.log_msg(f"Cannot save to {os.path.basename(path)}. Please close it in Excel first.", "WARNING")
                return
            try:
                self.last_dir = os.path.dirname(path)
                self._set_status("Generating...")
                QApplication.processEvents() # Keep UI responsive
                if self._setup_loader is None:
                    return self._set_status("Error: Setup not loaded")
                if self.setup_path is None:
                    return self._set_status("Error: Setup path missing")
                
                loader = self._setup_loader or SetupLoader(self.setup_path)
                metadata = loader.load_metadata()
                config = loader.load_config()
                students = loader.load_students()
                question_map = loader.load_question_map()

                validator = SetupValidator(
                    config=config,
                    students=students,
                    question_map=question_map,
                    only_CA=False,
                    metadata=metadata,  # Pass metadata to validator
                )
                
                validated_setup = validator.validate()
                generator = MarksWorkbookGenerator(
                    metadata=metadata,
                    validated_setup=validated_setup  # Use your actual validation logic
                )

                generator.generate(path, progress_callback=lambda m: self._set_status(m))
                self.gen_marks_path = path
                self._refresh_actions()
                self._update_labels()
                self._set_status("Marks Template Saved")
                self.log_msg(f"Marks template saved to: {self.gen_marks_path}", "SYSTEM")

            except Exception as e:
                self.log_msg(str(e), "ERROR")
                self._set_status("Generation Failed; Check the log below")

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

    def compute_results(self):
        if not self.setup_path or not self.filled_path:
            self.log_msg("Upload required files first.", "ERROR")
            return

        default_out = os.path.join(self.last_dir, f"{self._course_default_prefix}_CO_Results.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, "Save Results", default_out, "Excel (*.xlsx)")

        if not path:
            return
        if path:
                if not self.is_file_writable(path):
                    self._set_status("File is Open!")
                    self.log_msg(f"Cannot save to {os.path.basename(path)}. Please close it in Excel first.", "WARNING")
                    return
                try:
                    self.last_dir = os.path.dirname(path)
                    self._set_status("Computing...")
                    QApplication.processEvents()
                    if self.setup_path is None or self.filled_path is None:
                        return self._set_status("Error: Missing files") 
                    loader = self._setup_loader or SetupLoader(self.setup_path)
                    metadata = loader.load_metadata()
                    calc = COCalculator(metadata, self.setup_path, self.filled_path)
                    calc.compute(path)
                    
                    self._set_status("Results saved successfully")
                    self.clear_log()
                    self.log_msg(f"CO results saved to: {path}", "SYSTEM")
                    self._refresh_actions()
                    self._update_labels()

                except Exception as e:
                    self.log_msg(f"{str(e)}", "ERROR")
                    self._set_status("Computation Error. Check the log below")
    
    def generate_setup_template(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Create Course Details Template",
            "Course_Setup_Input_Template.xlsx",
            "Excel (*.xlsx)"
        )
        if not path:
            return
        self.last_dir = os.path.dirname(path)
        generator = SetupTemplateGenerator(path)
        generator.generate()

        self.log_msg(f"Setup template created: {path}", "SYSTEM")
        self._set_status("Template generated")