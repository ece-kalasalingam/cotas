import hashlib
import json
import os
import sys
from pathlib import Path
from typing import Any, Iterable, Optional

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QLabel


_LAST_DIR_FILE = Path.home() / ".cotas_last_dir.json"


def resource_path(relative_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, relative_path)


def get_run_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))


def load_last_dir() -> Optional[str]:
    try:
        if not _LAST_DIR_FILE.exists():
            return None
        payload = json.loads(_LAST_DIR_FILE.read_text(encoding="utf-8"))
        return payload.get("last_dir")
    except Exception:
        return None


def save_last_dir(path: str) -> None:
    try:
        _LAST_DIR_FILE.write_text(json.dumps({"last_dir": path}), encoding="utf-8")
    except Exception:
        pass


def normalize(value: Any) -> str:
    return "" if value is None else str(value).strip().lower()


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(value))
    except Exception:
        return default


def calculate_visual_width(value: Any) -> int:
    if value is None:
        return 0
    return len(str(value))


def get_range_coordinates(first_row: int, first_col: int, last_row: int, last_col: int) -> tuple[int, int, int, int]:
    return first_row, first_col, last_row, last_col


def generate_system_fingerprint(*parts: Iterable[Any]) -> str:
    h = hashlib.sha256()
    for p in parts:
        h.update(str(p).encode("utf-8"))
    return h.hexdigest()


def generate_workbook_fingerprint(type_id: str, sheet_names: list[str]) -> str:
    h = hashlib.sha256()
    h.update(type_id.encode("utf-8"))
    for name in sorted(sheet_names):
        h.update(name.strip().lower().encode("utf-8"))
    return h.hexdigest()


class ToastNotification:
    """Lightweight in-app toast fallback (non-blocking label)."""

    def __init__(self, parent, message: str, type: str = "info", timeout_ms: int = 2200):
        if parent is None:
            return

        label = QLabel(message, parent)
        color = {
            "success": "#1b8f3a",
            "error": "#b3261e",
            "warning": "#8a6d1d",
            "info": "#2f2f2f",
        }.get(type, "#2f2f2f")

        label.setStyleSheet(
            f"background:{color}; color:white; padding:8px 12px; border-radius:6px;"
        )
        label.adjustSize()
        x = max(12, parent.width() - label.width() - 16)
        y = 12
        label.move(x, y)
        label.show()
        QTimer.singleShot(timeout_ms, label.deleteLater)
