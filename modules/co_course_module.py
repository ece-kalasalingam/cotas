# modules/co_course_module.py

import os
from typing import List, TypedDict

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QListWidget, QMessageBox
)
from PySide6.QtCore import Qt

from core.co_results_course_engine import COResultsCourseEngine


class SectionResult(TypedDict):
    result_path: str


class COCourseModule(QWidget):

    def __init__(self):
        super().__init__()
        self.section_files: List[SectionResult] = []
        self._build_ui()

    # =========================================================
    # UI
    # =========================================================

    def _build_ui(self):

        layout = QVBoxLayout(self)

        title = QLabel("Coordinator – Multi Section CO Aggregation")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        self.add_button = QPushButton("Add Section CO Result File")
        self.compute_button = QPushButton("Compute Overall Course Result")

        layout.addWidget(self.add_button)
        layout.addWidget(self.compute_button)

        self.add_button.clicked.connect(self.add_section)
        self.compute_button.clicked.connect(self.compute)

    # =========================================================
    # Add Section Result
    # =========================================================

    def add_section(self):

        result, _ = QFileDialog.getOpenFileName(
            self,
            "Select Section CO Result File",
            "",
            "Excel (*.xlsx)"
        )

        if not result:
            return

        # prevent duplicate uploads
        for item in self.section_files:
            if item["result_path"] == result:
                QMessageBox.warning(self, "Duplicate", "This file is already added.")
                return

        self.section_files.append({
            "result_path": result
        })

        self.list_widget.addItem(os.path.basename(result))

    # =========================================================
    # Compute
    # =========================================================

    def compute(self):

        if not self.section_files:
            QMessageBox.warning(self, "Error", "No section result files added.")
            return

        output, _ = QFileDialog.getSaveFileName(
            self,
            "Save Overall Course Result",
            "Overall_CO_Result.xlsx",
            "Excel (*.xlsx)"
        )

        if not output:
            return

        try:
            engine = COResultsCourseEngine(self.section_files)
            engine.compute_and_export(output)

            QMessageBox.information(
                self,
                "Success",
                "Overall CO result generated successfully."
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))