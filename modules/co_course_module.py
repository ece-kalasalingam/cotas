import os
import sys
from datetime import datetime
from typing import List, TypedDict

from PySide6.QtWidgets import (
    QFrame, QGroupBox, QPlainTextEdit, QSizePolicy, QWidget, QVBoxLayout, 
    QHBoxLayout, QLabel, QPushButton, QFileDialog, QListWidget, 
    QListWidgetItem, QMessageBox
)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QFont, QPainter, QIcon

# Import your custom engine and utility
from core.co_results_course_engine import AttainmentPolicy, COResultsCourseEngine
from core.resources import resource_path
from modules.utils import ToastNotification

# --- Data Type ---
class SectionResult(TypedDict):
    result_path: str

# --- 1. Custom Item Widget for the List ---
class FileItemWidget(QWidget):
    """Custom row with file name and a Trash Bin button."""
    removed = Signal(str)

    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 2, 10, 2) # Reduced vertical margins

        file_name = os.path.basename(file_path)
        name_label = QLabel(file_name)
        name_label.setStyleSheet("font-size: 10pt;")

        self.remove_btn = QPushButton()
        self.remove_btn.setFixedSize(24, 24)
        self.remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        
        # Try to load trash icon
        icon_path = resource_path("assets/trash.svg")
        if os.path.exists(icon_path):
            self.remove_btn.setIcon(QIcon(icon_path))
            self.remove_btn.setIconSize(QSize(16, 16))
        else:
            self.remove_btn.setText("✕")

        self.remove_btn.setStyleSheet("""
            QPushButton { 
                background-color: transparent; 
                color: #e74c3c; 
                border: none;
            }
            QPushButton:hover { 
                background-color: rgba(231, 76, 60, 0.15); 
                border-radius: 4px; 
            }
        """)

        layout.addWidget(name_label)
        layout.addWidget(self.remove_btn)

        self.remove_btn.clicked.connect(lambda: self.removed.emit(self.file_path))

# --- 2. Custom List Widget with Drag & Drop + Placeholder ---
class FileDropListWidget(QListWidget):
    fileDropped = Signal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSpacing(2)
        # Reduced height to 200px
        self.setFixedHeight(200) 
        self.setStyleSheet("""
            QListWidget {
                border: 2px dashed #444;
                border-radius: 8px;
                outline: none;
                padding: 5px;
            }
        """)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.count() == 0:
            painter = QPainter(self.viewport())
            painter.setPen(Qt.GlobalColor.gray)
            painter.setFont(QFont("Segoe UI", 11))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, 
                             "Drag and Drop .xlsx files here")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.xlsx'):
                self.fileDropped.emit(file_path)
            else:
                ToastNotification(self.window(), "Only .xlsx files allowed!", type="error")

# --- 3. Main Module Class ---
class COCourseModule(QWidget):
    status_changed = Signal(str)

    def __init__(self):
        super().__init__()
        self.section_files: List[SectionResult] = []
        
        if getattr(sys, 'frozen', False):
            self.last_dir = os.path.dirname(sys.executable)
        else:
            self.last_dir = os.path.dirname(os.path.abspath(__file__))
        self.last_dir = os.path.normpath(self.last_dir)

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15) # Slightly tighter spacing
        layout.setContentsMargins(30, 20, 30, 20)

        # Header
        title = QLabel("Course Coordinator CO Aggregation")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)

        # File List Group
        self.files_group = QGroupBox("Input Section-wise CO Scores / Result Files")
        files_layout = QVBoxLayout(self.files_group)
        files_layout.setContentsMargins(15, 15, 15, 15)

        self.list_widget = FileDropListWidget()
        self.list_widget.fileDropped.connect(self.process_added_file)
        files_layout.addWidget(self.list_widget)

        # Button Row
        btn_layout = QHBoxLayout()
        
        self.add_button = QPushButton("Add File")
        
        self.clear_button = QPushButton("  Clear All")
        self.clear_button.setStyleSheet("""
            QPushButton:hover { 
                background-color: rgba(231, 76, 60, 0.1); 
                border-color: #e74c3c;
            }
        """)

        self.compute_button = QPushButton("Compute CO Attainment")

        for b in (self.add_button, self.clear_button, self.compute_button):
            b.setMinimumHeight(38)
            b.setCursor(Qt.CursorShape.PointingHandCursor)

        btn_layout.addWidget(self.add_button)
        btn_layout.addWidget(self.clear_button)
        btn_layout.addStretch()
        btn_layout.addWidget(self.compute_button)
        files_layout.addLayout(btn_layout)

        layout.addWidget(self.files_group)

        # Console Log
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setVisible(False)
        self.log.setMaximumHeight(150)
        self.log.setStyleSheet("font-family: Consolas; border: none;")
        layout.addWidget(self.log)

        layout.addStretch()

        # Signals
        self.add_button.clicked.connect(self.add_section_dialog)
        self.clear_button.clicked.connect(self.clear_all_files)
        self.compute_button.clicked.connect(self.compute)

        self._refresh_actions()

    def log_msg(self, text, level="INFO"):
        if not self.log.isVisible(): self.log.setVisible(True)
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.appendPlainText(f"[{ts}] {level}: {text}")

    def clear_log(self) -> None:
        self.log.clear()
        self.log.setVisible(False)

    def _set_status(self, text):
        self.status_changed.emit(text)

    def add_section_dialog(self):
        result, _ = QFileDialog.getOpenFileName(self, "Select Excel", self.last_dir, "Excel (*.xlsx)")
        if result: self.process_added_file(result)

    def _refresh_actions(self) -> None:
        has_files = len(self.section_files) > 0
        self.compute_button.setEnabled(has_files)
        self.clear_button.setEnabled(has_files)

    def process_added_file(self, file_path):
        if any(f["result_path"] == file_path for f in self.section_files):
            ToastNotification(self.window(), "File is already in the list", type="warning")
            self._set_status("File already added")
            self.log_msg(f"Attempted to add duplicate file: {file_path}", "WARNING")
            return

        self.section_files.append({"result_path": file_path})
        
        item = QListWidgetItem(self.list_widget)
        row_widget = FileItemWidget(file_path)
        row_widget.removed.connect(self.remove_file)
        item.setSizeHint(row_widget.sizeHint())
        
        self.list_widget.addItem(item)
        self.list_widget.setItemWidget(item, row_widget)
        self.log_msg(f"Added {file_path}", "SYSTEM")
        self._set_status(f"Added file: {os.path.basename(file_path)}")
        self.list_widget.update()
        self._refresh_actions()

    def remove_file(self, file_path):
        self.section_files = [f for f in self.section_files if f["result_path"] != file_path]
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            widget = self.list_widget.itemWidget(item)
            if isinstance(widget, FileItemWidget) and widget.file_path == file_path:
                self.list_widget.takeItem(i)
                break
        self.list_widget.update()
        self.log_msg(f"Removed {file_path}", "SYSTEM")
        self._set_status(f"Removed file: {os.path.basename(file_path)}")
        self._refresh_actions()
    def clear_all_files(self):
        if not self.section_files: return
        count = len(self.section_files)
        self.section_files.clear()
        self.list_widget.clear()
        ToastNotification(self.window(), f"Removed {count} files", type="info")
        self.log_msg(f"Cleared all files from the list", "SYSTEM")
        self._set_status("Cleared all files")
        self.list_widget.update()
        self._refresh_actions()

    def compute(self):
        if not self.section_files:
            ToastNotification(self.window(), "Please add files first!", type="error")
            return

        output, _ = QFileDialog.getSaveFileName(self, "Save Result", "Overall_CO_Result.xlsx", "Excel (*.xlsx)")
        if not output: return

        try:
            self._set_status("Processing...")
            policy = AttainmentPolicy(
                pass_mark=40,
                threshold_mark=60,
                high_mark=75,
                target_percent=70
            )
            engine = COResultsCourseEngine(self.section_files, policy)
            engine.compute_and_export(output)
            ToastNotification(self.window(), "Overall CO report generated!", type="success")
            self.clear_log()
            self.log_msg(f"CO Attainment calculated and result file saved to: {output}", "SYSTEM")
            self._set_status("Ready")
        except Exception as e:
            ToastNotification(self.window(), "Error computing CO results. See log for details.", type="error")
            self.log_msg(str(e), "ERROR")