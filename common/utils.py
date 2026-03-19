"""Shared utility helpers used across the application."""

import atexit
import json
import logging
import os
import shutil
import sys
import tempfile
from decimal import Decimal, InvalidOperation
from logging.handlers import RotatingFileHandler
from pathlib import Path, PureWindowsPath
from typing import Any, Callable, Literal

from common.error_catalog import resolve_validation_error_message
from common.exceptions import AppSystemError, ValidationError

SETTINGS_FILE_NAME = "settings.json"
DEFAULT_LOG_FILE_NAME = "focus.log"
SETTINGS_KEY_UI_LANGUAGE = "ui_language"
SETTINGS_KEY_LAST_SAVED_DIR = "last_saved_dir"
UI_LANGUAGE_DEFAULT = "en"
UI_LANGUAGE_AUTO_ALIASES = {"auto", "system", "os"}
PORTABLE_MODE_ENV_VAR = "FOCUS_PORTABLE"
LOG_FORMAT = "%(asctime)s %(levelname)s [%(name)s] [job=%(job_id)s step=%(step_id)s] %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
_APP_LOGGING_CONFIGURED = False
DEFAULT_LOG_MAX_BYTES = 2 * 1024 * 1024
DEFAULT_LOG_BACKUP_COUNT = 3
RUNTIME_MIN_FREE_BYTES_ENV_VAR = "FOCUS_RUNTIME_MIN_FREE_BYTES"
DEFAULT_RUNTIME_MIN_FREE_BYTES = 5 * 1024 * 1024
RuntimeStorageMode = Literal["installed", "portable", "dev"]
_RUNTIME_STORAGE_DIR_CACHE: dict[str, Path] = {}
_RUNTIME_TEMP_DIRS: set[Path] = set()
_RUNTIME_TEMP_CLEANUP_REGISTERED = False


class _SafeExtraFormatter(logging.Formatter):
    """Formatter that tolerates optional custom logging fields."""

    def format(self, record: logging.LogRecord) -> str:
        if not hasattr(record, "job_id"):
            record.job_id = "-"
        if not hasattr(record, "step_id"):
            record.step_id = "-"
        return super().format(record)


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
    return os.getenv(PORTABLE_MODE_ENV_VAR, "").strip().lower() in {"1", "true", "yes", "on"}


def _is_installed_exe(run_base: Path) -> bool:
    """Best-effort detection of installed exe vs portable exe."""
    if not getattr(sys, "frozen", False):
        return False
    if _is_portable_forced():
        return False

    if sys.platform.startswith("win"):
        # Preserve Windows semantics even when evaluated on non-Windows hosts.
        run_base_text = str(PureWindowsPath(str(run_base))).lower()
        install_roots = [
            os.getenv("ProgramFiles"),
            os.getenv("ProgramFiles(x86)"),
            os.path.join(os.getenv("LOCALAPPDATA", ""), "Programs"),
            os.path.join(os.getenv("LOCALAPPDATA", ""), "Microsoft", "WindowsApps"),
        ]
        normalized_roots = [
            str(PureWindowsPath(root)).lower().rstrip("/\\")
            for root in install_roots
            if root and root.strip()
        ]
        return any(run_base_text.startswith(root) for root in normalized_roots)

    run_base_text = str(run_base.resolve()).lower()

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


def _runtime_mode(run_base: Path | None = None) -> RuntimeStorageMode:
    base = run_base or _runtime_base_dir()
    if _is_installed_exe(base):
        return "installed"
    if getattr(sys, "frozen", False):
        return "portable"
    return "dev"


def _join_path(base: Path, child: str) -> Path:
    base_text = str(base)
    if "\\" in base_text:
        return Path(str(PureWindowsPath(base_text) / child))
    return base / child


def _installed_storage_dir(app_name: str) -> Path:
    if sys.platform.startswith("win"):
        app_data = os.getenv("APPDATA", str(Path.home() / "AppData" / "Roaming"))
        return Path(str(PureWindowsPath(app_data) / app_name))
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / app_name
    xdg_config_raw = os.getenv("XDG_CONFIG_HOME")
    xdg_config = Path(xdg_config_raw) if xdg_config_raw else (Path.home() / ".config")
    return xdg_config / app_name


def app_primary_storage_dir(app_name: str) -> Path:
    """Return preferred storage root before temp fallback checks."""
    run_base = _runtime_base_dir()
    mode = _runtime_mode(run_base)
    if mode == "installed":
        return _installed_storage_dir(app_name)
    return run_base


def _runtime_min_free_bytes() -> int:
    raw = os.getenv(RUNTIME_MIN_FREE_BYTES_ENV_VAR, "").strip()
    if not raw:
        return DEFAULT_RUNTIME_MIN_FREE_BYTES
    try:
        parsed = int(raw)
    except ValueError:
        return DEFAULT_RUNTIME_MIN_FREE_BYTES
    return max(parsed, 0)


def _directory_is_writable(path: Path) -> bool:
    probe_path = path / f".focus_write_probe_{os.getpid()}.tmp"
    try:
        path.mkdir(parents=True, exist_ok=True)
        with open(probe_path, "wb") as probe:
            probe.write(b"1")
        probe_path.unlink(missing_ok=True)
        return True
    except OSError:
        return False


def _directory_has_free_space(path: Path, *, min_free_bytes: int) -> bool:
    try:
        usage = shutil.disk_usage(str(path))
    except OSError:
        return False
    return usage.free >= min_free_bytes


def _is_storage_dir_usable(path: Path, *, min_free_bytes: int) -> bool:
    return _directory_is_writable(path) and _directory_has_free_space(path, min_free_bytes=min_free_bytes)


def _cleanup_runtime_temp_dirs() -> None:
    for path in sorted(_RUNTIME_TEMP_DIRS, key=lambda item: len(str(item)), reverse=True):
        shutil.rmtree(path, ignore_errors=True)
    _RUNTIME_TEMP_DIRS.clear()


def _register_runtime_temp_dir(path: Path) -> None:
    global _RUNTIME_TEMP_CLEANUP_REGISTERED
    _RUNTIME_TEMP_DIRS.add(path)
    if not _RUNTIME_TEMP_CLEANUP_REGISTERED:
        atexit.register(_cleanup_runtime_temp_dirs)
        _RUNTIME_TEMP_CLEANUP_REGISTERED = True


def _new_runtime_temp_dir(app_name: str) -> Path:
    temp_root = Path(tempfile.mkdtemp(prefix=f"{app_name.lower()}_runtime_"))
    _register_runtime_temp_dir(temp_root)
    return temp_root


def app_runtime_storage_dir(app_name: str) -> Path:
    """Resolve effective storage root with writeability and free-space fallback."""
    cached = _RUNTIME_STORAGE_DIR_CACHE.get(app_name)
    if cached is not None:
        if cached in _RUNTIME_TEMP_DIRS:
            return cached
        primary_now = app_primary_storage_dir(app_name)
        if str(cached) == str(primary_now):
            return cached

    primary = app_primary_storage_dir(app_name)
    if _is_storage_dir_usable(primary, min_free_bytes=_runtime_min_free_bytes()):
        _RUNTIME_STORAGE_DIR_CACHE[app_name] = primary
        return primary

    fallback = _new_runtime_temp_dir(app_name)
    _RUNTIME_STORAGE_DIR_CACHE[app_name] = fallback
    return fallback


def app_secrets_dir(app_name: str) -> Path:
    """Return folder path for persisted secret material."""
    if _runtime_mode() == "installed":
        if sys.platform.startswith("win"):
            program_data = os.getenv("PROGRAMDATA", r"C:\ProgramData")
            return Path(str(PureWindowsPath(program_data) / app_name / "secrets"))
        if sys.platform == "darwin":
            return Path("/Users/Shared") / app_name / "secrets"
        return _installed_storage_dir(app_name) / "secrets"
    return _join_path(app_runtime_storage_dir(app_name), "secrets")


def create_app_runtime_sqlite_file(app_name: str, *, prefix: str, suffix: str) -> tuple[int, str]:
    """Create a sqlite file in runtime storage; fallback to temp runtime root on errors."""
    storage_root = app_runtime_storage_dir(app_name)
    sqlite_dir = _join_path(storage_root, "sqlite")
    try:
        sqlite_dir.mkdir(parents=True, exist_ok=True)
        return tempfile.mkstemp(prefix=prefix, suffix=suffix, dir=str(sqlite_dir))
    except OSError:
        fallback_root = _new_runtime_temp_dir(app_name)
        sqlite_dir = _join_path(fallback_root, "sqlite")
        sqlite_dir.mkdir(parents=True, exist_ok=True)
        return tempfile.mkstemp(prefix=prefix, suffix=suffix, dir=str(sqlite_dir))


def _reset_runtime_storage_cache_for_tests() -> None:
    """Test helper for deterministic runtime storage path assertions."""
    _RUNTIME_STORAGE_DIR_CACHE.clear()
    _cleanup_runtime_temp_dirs()


def app_settings_path(app_name: str, file_name: str = SETTINGS_FILE_NAME) -> Path:
    """Resolve settings json path based on installed/portable/dev mode."""
    return _join_path(app_runtime_storage_dir(app_name), file_name)


def app_log_path(app_name: str, log_file_name: str = DEFAULT_LOG_FILE_NAME) -> Path:
    """Return the persistent log file path for current runtime mode."""
    settings_file = app_settings_path(app_name=app_name, file_name=SETTINGS_FILE_NAME)
    return settings_file.parent / log_file_name


def configure_app_logging(
    app_name: str,
    *,
    log_file_name: str = DEFAULT_LOG_FILE_NAME,
    level: int = logging.DEBUG,
    max_bytes: int = DEFAULT_LOG_MAX_BYTES,
    backup_count: int = DEFAULT_LOG_BACKUP_COUNT,
) -> Path:
    """Configure app logging to file only.

    User-facing messages should be shown in UI (status/toast/log panel), while
    detailed diagnostics are persisted in the log file.
    """
    global _APP_LOGGING_CONFIGURED
    if _APP_LOGGING_CONFIGURED:
        return app_log_path(app_name=app_name, log_file_name=log_file_name)

    log_path = app_log_path(app_name=app_name, log_file_name=log_file_name)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    formatter = _SafeExtraFormatter(
        fmt=LOG_FORMAT,
        datefmt=LOG_DATE_FORMAT,
    )

    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    _APP_LOGGING_CONFIGURED = True
    return log_path


def to_portable_path(path_value: str) -> str:
    """Serialize filesystem path to a JSON-safe, OS-readable form."""
    expanded = os.path.expanduser(path_value)
    normalized = os.path.normpath(expanded)
    return normalized.replace("\\", "/")


def from_portable_path(path_value: str) -> str:
    """Convert serialized path back to current OS path style."""
    return os.path.normpath(path_value)


def _read_settings_payload(settings_file: Path) -> dict[str, Any]:
    if not settings_file.exists():
        return {}
    try:
        payload = json.loads(settings_file.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _write_settings_payload(settings_file: Path, payload: dict[str, Any]) -> None:
    settings_file.parent.mkdir(parents=True, exist_ok=True)
    settings_file.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def _normalize_ui_language(value: str | None, *, default: str = UI_LANGUAGE_DEFAULT) -> str:
    normalized = (value or "").strip().replace("_", "-").lower()
    if normalized in UI_LANGUAGE_AUTO_ALIASES:
        return default
    return normalized or default


def get_ui_language_preference(app_name: str, file_name: str = SETTINGS_FILE_NAME) -> str:
    """Return persisted UI language preference with normalized default."""
    settings_file = app_settings_path(app_name=app_name, file_name=file_name)
    payload = _read_settings_payload(settings_file)
    return _normalize_ui_language(payload.get(SETTINGS_KEY_UI_LANGUAGE), default=UI_LANGUAGE_DEFAULT)


def set_ui_language_preference(
    app_name: str,
    *,
    ui_language: str,
    file_name: str = SETTINGS_FILE_NAME,
) -> Path:
    """Persist UI language preference and return settings path."""
    settings_file = app_settings_path(app_name=app_name, file_name=file_name)
    payload = _read_settings_payload(settings_file)
    payload[SETTINGS_KEY_UI_LANGUAGE] = _normalize_ui_language(
        ui_language,
        default=UI_LANGUAGE_DEFAULT,
    )
    _write_settings_payload(settings_file, payload)
    return settings_file


def set_last_saved_dir(
    directory: str, app_name: str, file_name: str = SETTINGS_FILE_NAME
) -> Path:
    """Create/update `last_saved_dir` in settings json and return json path."""
    settings_file = app_settings_path(app_name=app_name, file_name=file_name)
    payload = _read_settings_payload(settings_file)

    payload[SETTINGS_KEY_LAST_SAVED_DIR] = to_portable_path(directory)
    _write_settings_payload(settings_file, payload)
    return settings_file


def get_last_saved_dir(app_name: str, file_name: str = SETTINGS_FILE_NAME) -> str | None:
    """Read `last_saved_dir` from settings json if present."""
    settings_file = app_settings_path(app_name=app_name, file_name=file_name)
    payload = _read_settings_payload(settings_file)
    raw_path = payload.get(SETTINGS_KEY_LAST_SAVED_DIR)
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


def resolve_existing_dialog_directory(start_path: str | None) -> str:
    """Resolve nearest existing directory for folder dialogs.

    If `start_path` points to a file path (existing or not), this walks upward to
    find the first existing directory. Returns empty string when no usable parent
    exists, allowing Qt to fall back to OS defaults.
    """
    if not start_path or not str(start_path).strip():
        return ""
    try:
        candidate = Path(os.path.expanduser(start_path))
    except OSError:
        return ""

    while True:
        try:
            if candidate.exists() and candidate.is_dir():
                return str(candidate)
        except OSError:
            pass
        parent = candidate.parent
        if parent == candidate:
            return ""
        candidate = parent


def remember_dialog_dir(selected_path: str, app_name: str) -> None:
    """Persist selected file/folder directory as last_saved_dir if valid."""
    if not selected_path:
        return

    try:
        normalized_selected = os.path.normpath(os.path.expanduser(selected_path))
        selected = Path(normalized_selected)
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


def log_process_message(
    process_name: str,
    *,
    logger: logging.Logger,
    error: Exception | None = None,
    notify: Callable[[str, Literal["info", "success", "warning", "error"]], None] | None = None,
    success_message: str | None = None,
    user_success_message: str | None = None,
    user_error_message: str | None = None,
    user_validation_message: str | None = None,
    job_id: str | None = None,
    step_id: str | None = None,
) -> bool:
    """Log and publish a user-facing status update for a process.

    Behavior:
    - Success: logs info and emits success message.
    - ValidationError (data/user error): logs a generic English warning and emits detailed error text.
    - AppSystemError: logs a generic English error and emits a generic process-scoped error.
    - Other errors: logs traceback and emits a generic process-scoped error.
    """
    if error is None:
        log_message = success_message or f"{process_name} completed successfully."
        user_message = user_success_message or log_message
        logger.info(
            log_message,
            extra={
                "user_message": user_message,
                "job_id": job_id,
                "step_id": step_id,
                "error_code": "NONE",
            },
        )
        if notify is not None:
            notify(user_message, "success")
        return True

    if isinstance(error, ValidationError):
        detail = str(error).strip()
        if hasattr(error, "code"):
            code = getattr(error, "code", "")
            context = getattr(error, "context", {}) or {}
            resolved = resolve_validation_error_message(str(code), context)
            if resolved and resolved != str(code):
                detail = resolved
        if not detail:
            detail = "Validation failed due to invalid data."
        user_message = user_validation_message or detail
        logger.warning(
            "%s failed due to data error.",
            process_name,
            extra={
                "user_message": user_message,
                "job_id": job_id,
                "step_id": step_id,
                "error_code": str(getattr(error, "code", "VALIDATION_ERROR")),
            },
        )
        if notify is not None:
            notify(user_message, "error")
        return False

    user_message = user_error_message or f"Error happened while {process_name}."
    if isinstance(error, AppSystemError):
        logger.error(
            "%s failed due to a system/application error.",
            process_name,
            extra={
                "user_message": user_message,
                "job_id": job_id,
                "step_id": step_id,
                "error_code": "APP_SYSTEM_ERROR",
            },
        )
        if notify is not None:
            notify(user_message, "error")
        return False

    logger.exception(
        "%s failed due to a system/application error.",
        process_name,
        exc_info=error,
        extra={
            "user_message": user_message,
            "job_id": job_id,
            "step_id": step_id,
            "error_code": "UNEXPECTED_ERROR",
        },
    )
    if notify is not None:
        notify(user_message, "error")
    return False


def emit_user_status(
    status_signal: Any | None,
    message: str,
    *,
    logger: logging.Logger | None = None,
) -> None:
    """Emit a user-facing status message through a Qt-like signal when available."""
    if not message or not message.strip():
        return
    if status_signal is None:
        return

    emit = getattr(status_signal, "emit", None)
    if not callable(emit):
        if logger is not None:
            logger.debug("Status signal does not expose an emit() method.")
        return

    try:
        emit(message)
    except Exception:
        if logger is not None:
            logger.exception("Failed to emit user status message.")





