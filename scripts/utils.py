from typing import Any, Tuple, List, Dict
from scripts.constants import DEFAULT_MAX_COL_WIDTH, DEFAULT_MIN_COL_WIDTH, WIDTH_PADDING
import hashlib
import json

def calculate_visual_width(value: Any) -> int:
    """
    Determines the Excel column width based on string length.
    Constrained by constants to prevent ultra-wide or invisible columns.
    """
    if value is None:
        return DEFAULT_MIN_COL_WIDTH
        
    text_len = len(str(value)) + WIDTH_PADDING
    return min(max(text_len, DEFAULT_MIN_COL_WIDTH), DEFAULT_MAX_COL_WIDTH)

def get_range_coordinates(range_str: str) -> Tuple[int, int, int, int]:
    """
    Optional helper: Converts 'A1:C10' to (0, 0, 9, 2).
    Useful if you want to allow humans to write strings in the blueprint 
    but need integers for Pylance/XlsxWriter.
    """
    import xlsxwriter.utility as xl_util
    # This splits 'A1:C10' into ['A1', 'C10']
    start, end = range_str.split(':')
    row_start, col_start = xl_util.xl_cell_to_rowcol(start)
    row_end, col_end = xl_util.xl_cell_to_rowcol(end)
    return row_start, col_start, row_end, col_end

def generate_system_fingerprint(
    setup: Any, 
    metadata: Any, 
    headers: List[str], 
    weightages: Dict[str, Any]
) -> str:
    # 1. Basic Identity
    anchor_data = {
        "course_code": getattr(metadata, "course_code", "N/A"),
        "semester": getattr(metadata, "semester", "N/A"),
        "instructor": getattr(setup, "instructor_id", "N/A"),
        "header_checksum": hashlib.md5("".join(headers).encode()).hexdigest(),
    }

    # 2. Only add weightages to the hash if they are actually present
    if weightages:
        anchor_data["weight_config"] = dict(sorted(weightages.items()))
    
    # 3. Add the secret salt
    anchor_data["secret"] = "YOUR_INTERNAL_SECRET"
    
    data_string = json.dumps(anchor_data, sort_keys=True)
    return hashlib.sha256(data_string.encode()).hexdigest()
from PySide6.QtWidgets import QLabel, QGraphicsOpacityEffect, QWidget
from PySide6.QtCore import Qt, QPropertyAnimation, QTimer, QPoint

# Helper to normalize strings: lowercase and remove all internal/external whitespace
def normalize(s) -> str:
    return "".join(str(s).split()).lower()
def safe_int(z):
    try:
        return (0, int(z))
    except:
        return (1, z)
    
import os
import sys


def resource_path(relative_path: str) -> str:
    """
    Returns absolute path to bundled resource.

    Works for:
    - Development mode
    - PyInstaller --onedir
    - PyInstaller --onefile
    """

    if getattr(sys, "frozen", False):
        # PyInstaller
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        # Development
        base_path = os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..")
        )

    return os.path.normpath(os.path.join(base_path, relative_path))
class ToastNotification(QLabel):
    def __init__(self, parent, message, type="info", duration=6000):
        super().__init__(parent)
        
        colors = {
            "success": "#2ecc71",
            "warning": "#f39c12",
            "error": "#e74c3c",
            "info": "#3498db"
        }
        
        bg_color = colors.get(type, "#333333")
        self.setText(message)
        self.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.setFixedSize(380, 55) 
        
        # Sharp corners (no border-radius)
        self.setStyleSheet(f"""
            QLabel {{
                background-color: {bg_color};
                color: white;
                border: none;
                font-size: 11pt;
                font-family: 'Segoe UI', sans-serif;
                padding-left: 15px;
                padding-right: 15px;
            }}
        """)

        self.opacity_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.opacity_effect)
        
        self.update_position()
        self.fade_in()
        QTimer.singleShot(duration, self.fade_out)

    def update_position(self):
        p = self.parent()
        if isinstance(p, QWidget):
            p_rect = p.rect()
            margin = 20
            
            # X = Total Width - Toast Width - Margin
            x = p_rect.width() - self.width() - margin
            
            # Y = Total Height - Toast Height - Margin
            y = p_rect.height() - self.height() - margin
            
            self.move(QPoint(x, y))

    def fade_in(self):
        self.anim = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.anim.setDuration(300)
        self.anim.setStartValue(0)
        self.anim.setEndValue(1)
        self.anim.start()
        self.show()

    def fade_out(self):
        self.anim = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.anim.setDuration(800)
        self.anim.setStartValue(1)
        self.anim.setEndValue(0)
        self.anim.finished.connect(self.deleteLater)
        self.anim.start()