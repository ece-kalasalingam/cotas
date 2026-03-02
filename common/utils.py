"""Shared utility helpers used across the application."""

import json
import logging
import os
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


def resource_path(relative_path: str) -> str:
    base = getattr(
        sys,
        "_MEIPASS",
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    )
    return os.path.join(base, relative_path)


def _runtime_base_dir() -> Path:
    """Return the directory where the app is currently running from."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _is_portable_forced() -> bool:
    """Allow forcing portable mode on any OS via env var."""
    return os.getenv("FOCUS_PORTABLE", "").strip().lower() in {"1", "true", "yes", "on"}


def _is_installed_exe(run_base: Path) -> bool:
    """Best-effort detection of installed exe vs portable exe."""
    if not getattr(sys, "frozen", False):
        return False
    if _is_portable_forced():
        return False

    run_base_text = str(run_base.resolve()).lower()

    if sys.platform.startswith("win"):
        install_roots = [
            os.getenv("ProgramFiles"),
            os.getenv("ProgramFiles(x86)"),
            os.path.join(os.getenv("LOCALAPPDATA", ""), "Programs"),
            os.path.join(os.getenv("LOCALAPPDATA", ""), "Microsoft", "WindowsApps"),
        ]
        normalized_roots = [
            str(Path(root)).lower() for root in install_roots if root and root.strip()
        ]
        return any(run_base_text.startswith(root) for root in normalized_roots)

    if sys.platform == "darwin":
        return (
            run_base_text.startswith("/applications/")
            or "/applications/" in run_base_text
            or ".app/contents/macos" in run_base_text
        )

    # Linux and other Unix-like platforms.
    linux_install_roots = [
        "/usr/",
        "/opt/",
        "/snap/",
        "/var/lib/flatpak/",
        "/app/",
    ]
    return any(run_base_text.startswith(root) for root in linux_install_roots)


def app_settings_path(app_name: str, file_name: str = "settings.json") -> Path:
    """Resolve settings json path based on installed/portable/dev mode."""
    run_base = _runtime_base_dir()
    if _is_installed_exe(run_base):
        if sys.platform.startswith("win"):
            app_data = Path(
                os.getenv("APPDATA", str(Path.home() / "AppData" / "Roaming"))
            )
            return app_data / app_name / file_name
        if sys.platform == "darwin":
            return Path.home() / "Library" / "Application Support" / app_name / file_name
        xdg_config_raw = os.getenv("XDG_CONFIG_HOME")
        xdg_config = Path(xdg_config_raw) if xdg_config_raw else (Path.home() / ".config")
        return xdg_config / app_name / file_name
    return run_base / file_name


def app_log_path(app_name: str, log_file_name: str = "focus.log") -> Path:
    """Return the persistent log file path for current runtime mode."""
    settings_file = app_settings_path(app_name=app_name, file_name="settings.json")
    return settings_file.parent / log_file_name


def configure_app_logging(
    app_name: str,
    *,
    log_file_name: str = "focus.log",
    level: int = logging.INFO,
) -> Path:
    """Configure app logging to file and console (console only in dev mode)."""
    if getattr(configure_app_logging, "_configured", False):
        return app_log_path(app_name=app_name, log_file_name=log_file_name)

    log_path = app_log_path(app_name=app_name, log_file_name=log_file_name)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    formatter = logging.Formatter(
        fmt="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    if not getattr(sys, "frozen", False):
        console_handler = logging.StreamHandler()
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)

    configure_app_logging._configured = True
    return log_path


def to_portable_path(path_value: str) -> str:
    """Serialize filesystem path to a JSON-safe, OS-readable form."""
    return Path(path_value).expanduser().resolve().as_posix()


def from_portable_path(path_value: str) -> str:
    """Convert serialized path back to current OS path style."""
    return os.path.normpath(path_value)


def set_last_saved_dir(
    directory: str, app_name: str, file_name: str = "settings.json"
) -> Path:
    """Create/update `last_saved_dir` in settings json and return json path."""
    settings_file = app_settings_path(app_name=app_name, file_name=file_name)
    settings_file.parent.mkdir(parents=True, exist_ok=True)

    payload: dict[str, Any] = {}
    if settings_file.exists():
        try:
            payload = json.loads(settings_file.read_text(encoding="utf-8"))
            if not isinstance(payload, dict):
                payload = {}
        except (json.JSONDecodeError, OSError):
            payload = {}

    payload["last_saved_dir"] = to_portable_path(directory)
    settings_file.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return settings_file


def get_last_saved_dir(app_name: str, file_name: str = "settings.json") -> str | None:
    """Read `last_saved_dir` from settings json if present."""
    settings_file = app_settings_path(app_name=app_name, file_name=file_name)
    if not settings_file.exists():
        return None

    try:
        payload = json.loads(settings_file.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None

    if not isinstance(payload, dict):
        return None

    raw_path = payload.get("last_saved_dir")
    if not isinstance(raw_path, str) or not raw_path.strip():
        return None
    return from_portable_path(raw_path)


def resolve_dialog_start_path(app_name: str, file_name: str = "") -> str:
    """Resolve a file dialog start path using JSON, then Downloads, then Home."""
    last_dir = get_last_saved_dir(app_name)
    if last_dir:
        candidate = Path(last_dir).expanduser()
        if candidate.exists() and candidate.is_dir():
            return str(candidate / file_name) if file_name else str(candidate)

    downloads_dir = Path.home() / "Downloads"
    if downloads_dir.exists() and downloads_dir.is_dir():
        return str(downloads_dir / file_name) if file_name else str(downloads_dir)

    return str(Path.home() / file_name) if file_name else str(Path.home())


def remember_dialog_dir(selected_path: str, app_name: str) -> None:
    """Persist selected file/folder directory as last_saved_dir if valid."""
    if not selected_path:
        return

    try:
        selected = Path(selected_path).expanduser().resolve()
    except OSError:
        return

    directory = selected if selected.is_dir() else selected.parent
    if not directory.exists() or not directory.is_dir():
        return

    set_last_saved_dir(str(directory), app_name=app_name)


def remember_dialog_dir_safe(
    selected_path: str,
    app_name: str,
    *,
    logger: logging.Logger | None = None,
) -> None:
    """Persist dialog directory and swallow filesystem failures."""
    try:
        remember_dialog_dir(selected_path, app_name=app_name)
    except OSError as exc:
        if logger is not None:
            logger.warning("Failed to persist last_saved_dir: %s", exc)


def normalize(value: Any) -> str:
    return "" if value is None else str(value).strip().lower()


def coerce_excel_number(value: Any) -> Any:
    """Convert numeric-like values (including numeric strings) into numbers.

    Rules:
    - None stays None
    - bool stays bool
    - int stays int
    - float becomes int when integral (e.g., 12.0 -> 12), else float
    - str values are stripped; numeric strings are converted to int/float
    - non-numeric strings are returned as stripped strings
    """

    if value is None:
        return None

    if isinstance(value, bool):
        return value

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        return int(value) if value.is_integer() else value

    if isinstance(value, str):
        token = value.strip()
        if token == "":
            return ""

        cleaned = token.replace(",", "")
        try:
            number = Decimal(cleaned)
        except InvalidOperation:
            return token

        if number == number.to_integral_value():
            return int(number)
        return float(number)

    return value
