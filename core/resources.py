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