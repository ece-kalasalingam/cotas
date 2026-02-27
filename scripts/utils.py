from typing import Any, Optional, Tuple, List, Dict
from scripts.constants import DEFAULT_MAX_COL_WIDTH, DEFAULT_MIN_COL_WIDTH, WIDTH_PADDING
import hashlib
import xlsxwriter.utility as xl_util
import sys
import os
from PySide6.QtWidgets import QLabel, QGraphicsOpacityEffect, QWidget
from PySide6.QtCore import Qt, QPropertyAnimation, QTimer, QPoint, QStandardPaths
import json

# --- 1. THE CANONICALIZATION LAYER ---

def normalize(s: Any) -> str:
    """Lowercase and remove all internal/external whitespace."""
    if s is None: return "empty"
    return "".join(str(s).split()).lower()

def canonicalize(value: Any) -> str:
    """
    Standardizes data for the Hash:
    1. Normalizes (Lower + Strip).
    2. If numeric, forces to 'f.4f' string to avoid 10 vs 10.0 confusion.
    """
    if value is None:
        return "empty"

    # Step 1: Normalize (Lower + Whitespace removal)
    val_norm = normalize(value)
    if val_norm == "" or val_norm == "none":
        return "empty"

    # Step 2: Numeric Lock (Value -> Float -> String)
    try:
        # normalize() already cleaned spaces/commas, so float() is safe
        float_val = float(val_norm)
        return f"{float_val:.4f}"
    except (ValueError, TypeError):
        # Step 3: Text Lock (Already normalized)
        return val_norm

def generate_system_fingerprint(
    type_id: str,
    metadata: Dict[str, Any],
    sheet_names: List[str],
    headers: Optional[List[str]] = None
) -> str:
    """
    Unified SHA-256 fingerprint for identity and tamper detection.
    Now correctly incorporates sheet names into the hash.
    """
    hasher = hashlib.sha256()
    
    # 1. Identity (Version Control)
    hasher.update(canonicalize(type_id).encode('utf-8'))

    # 2. Metadata (Sorted keys for determinism)
    for key in sorted(metadata.keys()):
        k_norm = normalize(key)
        v_can = canonicalize(metadata[key])
        hasher.update(f"{k_norm}:{v_can}".encode('utf-8'))

    # 3. Structural Integrity (Sheet Names)
    # We sort them so the order in the list doesn't break the hash
    for name in sorted(sheet_names):
        hasher.update(canonicalize(name).encode('utf-8'))

    # 4. Content Lock (Headers, Students, etc.)
    for item in headers or []:
        hasher.update(canonicalize(item).encode('utf-8'))

    return hasher.hexdigest()

# --- 2. PATH & CALCULATION UTILS ---

def resource_path(relative_path: str) -> str:
    """Absolute path to bundled resource (Works for Dev and PyInstaller)."""
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.normpath(os.path.join(base_path, relative_path))

def calculate_visual_width(value: Any) -> int:
    if value is None: return DEFAULT_MIN_COL_WIDTH
    text_len = len(str(value)) + WIDTH_PADDING
    return min(max(text_len, DEFAULT_MIN_COL_WIDTH), DEFAULT_MAX_COL_WIDTH)

# --- 3. GUI NOTIFICATIONS ---

def get_range_coordinates(range_str: str) -> Tuple[int, int, int, int]:
    """
    Optional helper: Converts 'A1:C10' to (0, 0, 9, 2).
    Useful if you want to allow humans to write strings in the blueprint 
    but need integers for Pylance/XlsxWriter.
    """
    # This splits 'A1:C10' into ['A1', 'C10']
    start, end = range_str.split(':')
    row_start, col_start = xl_util.xl_cell_to_rowcol(start)
    row_end, col_end = xl_util.xl_cell_to_rowcol(end)
    return row_start, col_start, row_end, col_end

def get_run_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))


def get_config_path() -> str:
    config_dir = QStandardPaths.writableLocation(
        QStandardPaths.StandardLocation.AppDataLocation
    )

    app_dir = os.path.join(config_dir, "COTAS")
    os.makedirs(app_dir, exist_ok=True)

    return os.path.join(app_dir, "config.json")


def load_last_dir() -> str | None:
    config_path = get_config_path()

    if not os.path.exists(config_path):
        return None

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        directory = data.get("last_dir")
        if directory and os.path.isdir(directory):
            return directory

    except Exception as e:
        raise SystemError(f"Failed to load config file: {e}")

    return None


def save_last_dir(directory: str) -> None:
    if not directory:
        return

    config_path = get_config_path()
    data = {}

    try:
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}

        data["last_dir"] = directory

        tmp_path = config_path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

        os.replace(tmp_path, config_path)

    except Exception as e:
        raise SystemError(f"Failed to save config file: {e}")
    
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

## --- 4. ENGINE INTERFACE ---
def safe_int(z):
    try:
        return (0, int(z))
    except:
        return (1, z)