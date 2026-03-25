import logging
import os
import re
import sys

from PySide6.QtCore import QEvent, QLockFile, QObject, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import QApplication, QMessageBox, QSplashScreen

try:
    import qdarktheme  # type: ignore[import-not-found]
except ModuleNotFoundError:
    qdarktheme = None

from common.constants import (
    APP_NAME,
    APP_ORGANIZATION,
    MAIN_SPLASH_MS,
    QT_ADAPTIVE_STRUCTURE_SENSITIVITY,
    SINGLE_INSTANCE_ACK_PAYLOAD,
    SINGLE_INSTANCE_ACTIVATE_PAYLOAD,
    SINGLE_INSTANCE_CLIENT_ACK_TIMEOUT_MS,
    SINGLE_INSTANCE_CLIENT_CONNECT_TIMEOUT_MS,
    SINGLE_INSTANCE_CLIENT_WRITE_TIMEOUT_MS,
    SINGLE_INSTANCE_LOCK_TIMEOUT_MS,
    SINGLE_INSTANCE_SERVER_IO_TIMEOUT_MS,
    SPLASH_STATUS_COLOR,
    STARTUP_TOAST_DURATION_MS,
    STARTUP_TOAST_QUIT_DELAY_MS,
    UI_LANGUAGE,
    WIN32_SHOW_WINDOW_RESTORE,
)
from common.contracts import validate_blueprint_registry_contracts
from common.crash_reporting import (
    capture_unhandled_exception,
    has_remote_crash_endpoint,
)
from common.exceptions import ConfigurationError
from common.i18n import get_language, set_language, t
from common.toast import ToastLevel, show_toast
from common.ui_stylings import apply_global_ui_styles
from common.utils import (
    UI_LANGUAGE_AUTO_ALIASES,
    app_runtime_storage_dir,
    configure_app_logging,
    get_ui_language_preference,
    resource_path,
    set_ui_language_preference,
)
from common.workbook_integrity.workbook_secret import ensure_workbook_secret_policy
from main_window import MainWindow

_logger = logging.getLogger(__name__)
_THEME_REFRESH_DEBOUNCE_MS = 180
_theme_apply_in_progress = False
_theme_refresh_pending = False


def _build_splash_pixmap() -> QPixmap:
    pixmap = QPixmap(520, 240)
    pixmap.fill(QColor("#2957A4"))

    painter = QPainter(pixmap)
    painter.setPen(QColor("#ffffff"))
    splash_font = QFont(painter.font())
    splash_font.setPointSize(20)
    splash_font.setBold(True)
    painter.setFont(splash_font)
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, APP_NAME)
    painter.end()
    return pixmap


def _acquire_exe_single_instance_lock() -> QLockFile | None:
    # Allow multiple instances in dev/python mode.
    if not getattr(sys, "frozen", False):
        return None

    app_data = app_runtime_storage_dir(APP_NAME)
    app_data.mkdir(parents=True, exist_ok=True)

    lock_path = str(app_data / f"{APP_NAME}.lock")
    lock = QLockFile(lock_path)

    # Immediate check; if already locked, another instance is running.
    if not lock.tryLock(SINGLE_INSTANCE_LOCK_TIMEOUT_MS):
        return None

    return lock


def _activation_server_name() -> str:
    raw_name = f"{APP_ORGANIZATION}_{APP_NAME}_single_instance"
    # Keep the server name OS-safe and deterministic.
    return re.sub(r"[^A-Za-z0-9_.-]", "_", raw_name)


def _signal_existing_instance_to_activate() -> bool:
    socket = QLocalSocket()
    socket.connectToServer(_activation_server_name())
    connected = socket.waitForConnected(SINGLE_INSTANCE_CLIENT_CONNECT_TIMEOUT_MS)
    if connected:
        socket.write(SINGLE_INSTANCE_ACTIVATE_PAYLOAD)
        socket.flush()
        socket.waitForBytesWritten(SINGLE_INSTANCE_CLIENT_WRITE_TIMEOUT_MS)
        ack_ok = (
            socket.waitForReadyRead(SINGLE_INSTANCE_CLIENT_ACK_TIMEOUT_MS)
            and bytes(socket.readAll().data()).strip() == SINGLE_INSTANCE_ACK_PAYLOAD
        )
        socket.disconnectFromServer()
        socket.deleteLater()
        return ack_ok
    socket.deleteLater()
    return False


def _raise_and_activate_window(window: MainWindow) -> None:
    if window.isMinimized():
        window.showNormal()
    else:
        window.setWindowState(window.windowState() & ~Qt.WindowState.WindowMinimized)
    window.show()
    window.raise_()
    window.activateWindow()
    QApplication.alert(window)

    if sys.platform == "win32":
        try:
            import ctypes

            hwnd = int(window.winId())
            user32 = ctypes.windll.user32
            user32.ShowWindow(hwnd, WIN32_SHOW_WINDOW_RESTORE)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        except Exception:
            _logger.exception("Win32 foreground activation fallback failed.")


def _install_activation_server(window: MainWindow) -> QLocalServer:
    server = QLocalServer()
    server_name = _activation_server_name()
    QLocalServer.removeServer(server_name)
    if not server.listen(server_name):
        _logger.warning("Could not start activation server: %s", server.errorString())
        return server

    def _on_new_connection() -> None:
        while server.hasPendingConnections():
            socket = server.nextPendingConnection()
            if socket is not None:
                socket.waitForReadyRead(SINGLE_INSTANCE_SERVER_IO_TIMEOUT_MS)
                _ = socket.readAll().data()
                _raise_and_activate_window(window)
                socket.write(SINGLE_INSTANCE_ACK_PAYLOAD)
                socket.flush()
                socket.waitForBytesWritten(SINGLE_INSTANCE_SERVER_IO_TIMEOUT_MS)
                socket.disconnectFromServer()
                socket.deleteLater()

    server.newConnection.connect(_on_new_connection)
    return server


def _setup_system_theme() -> None:
    """Apply qdarktheme using OS-follow mode.

    Hybrid styling contract:
    1) qdarktheme owns base palette/theme negotiation with the OS.
    2) app-managed stylesheet blocks are layered separately by `apply_global_ui_styles`.
    """
    global _theme_apply_in_progress
    if qdarktheme is None:
        return
    if _theme_apply_in_progress:
        return
    _theme_apply_in_progress = True
    try:
        qdarktheme.setup_theme("auto")
    except (RuntimeError, AttributeError, TypeError, ValueError) as exc:
        _logger.warning("Failed to apply qdarktheme.", exc_info=exc)
    except Exception as exc:
        _logger.exception("Unexpected error while applying qdarktheme.", exc_info=exc)
    finally:
        _theme_apply_in_progress = False


def _schedule_system_theme_refresh(app: QApplication) -> None:
    """Debounced fallback refresh for missed OS theme transitions.

    This must stay debounced to avoid recursive palette/style storms.
    """
    global _theme_refresh_pending
    if qdarktheme is None or _theme_refresh_pending:
        return
    _theme_refresh_pending = True
    _logger.debug("Theme refresh scheduled (debounced).")

    def _run() -> None:
        global _theme_refresh_pending
        _theme_refresh_pending = False
        # Hybrid path: refresh qdarktheme (OS-aware) and then reapply our managed app styles.
        _setup_system_theme()
        apply_global_ui_styles(app)
        _logger.debug("Theme refresh executed (qdarktheme + managed styles).")

    QTimer.singleShot(_THEME_REFRESH_DEBOUNCE_MS, _run)


class _UiStyleRefreshFilter(QObject):
    """Event bridge that keeps managed styles in sync with runtime UI changes."""
    def __init__(self, app: QApplication) -> None:
        super().__init__()
        self._app = app

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        event_type = event.type()
        if event_type == QEvent.Type.ThemeChange:
            # Keep runtime updates lightweight: only react to concrete theme changes.
            _schedule_system_theme_refresh(self._app)
        return super().eventFilter(watched, event)


def _wire_global_style_refresh(app: QApplication) -> None:
    """Wire the hybrid theme/styling refresh hooks once per QApplication."""
    app_refresher = _UiStyleRefreshFilter(app)
    install_filter = getattr(app, "installEventFilter", None)
    if callable(install_filter):
        install_filter(app_refresher)
    setattr(app, "_ui_style_refresh_filter", app_refresher)
    _logger.debug("Hybrid UI style refresh wiring initialized.")


def _install_excepthook() -> None:
    previous_hook = sys.excepthook

    def _hook(exc_type, exc_value, exc_traceback):
        _logger.exception(
            "Unhandled exception in application.",
            exc_info=(exc_type, exc_value, exc_traceback),
        )
        report_path = capture_unhandled_exception(exc_type, exc_value, exc_traceback)
        if report_path is not None:
            _logger.error(
                "Crash report captured.",
                extra={
                    "user_message": f"Crash report saved: {report_path}",
                    "error_code": "CRASH_REPORTED",
                },
            )
            if has_remote_crash_endpoint():
                _logger.info(
                    "Crash report endpoint configured; upload pipeline can process local crash spool.",
                    extra={"error_code": "CRASH_PIPELINE_READY"},
                )
        try:
            show_toast(None, t("app.unexpected_error"), title=APP_NAME, level="error")
        except (RuntimeError, AttributeError, TypeError) as toast_exc:
            _logger.warning(
                "Unable to display unhandled-exception toast.",
                exc_info=toast_exc,
                extra={"error_code": "UNHANDLED_TOAST_FAILED"},
            )
        previous_hook(exc_type, exc_value, exc_traceback)

    sys.excepthook = _hook


def _notify_and_wait(app: QApplication, *, title: str, message: str, level: ToastLevel) -> int:
    show_toast(None, message, title=title, level=level, duration_ms=STARTUP_TOAST_DURATION_MS)
    QTimer.singleShot(STARTUP_TOAST_QUIT_DELAY_MS, app.quit)
    return app.exec()


def _show_startup_error_dialog(*, title: str, message: str) -> None:
    QMessageBox.critical(None, title, message)


def _validate_startup_workbook_password(app: QApplication) -> int | None:
    try:
        ensure_workbook_secret_policy()
        return None
    except ConfigurationError:
        if getattr(sys, "frozen", False):
            message = t("app.startup.workbook_secret_missing_frozen")
        else:
            message = t("app.startup.workbook_secret_missing_dev")
    _logger.error(
        "Startup blocked: workbook secret is unavailable.",
        extra={
            "user_message": message,
            "error_code": "MISSING_WORKBOOK_PASSWORD",
        },
    )
    _show_startup_error_dialog(title=APP_NAME, message=message)
    return 1


def main() -> int:
    os.environ["QT_ADAPTIVE_STRUCTURE_SENSITIVITY"] = QT_ADAPTIVE_STRUCTURE_SENSITIVITY
    validate_blueprint_registry_contracts()
    configure_app_logging(APP_NAME)
    startup_language = get_ui_language_preference(APP_NAME) or UI_LANGUAGE
    selected_startup_language = (
        UI_LANGUAGE if startup_language.lower() in UI_LANGUAGE_AUTO_ALIASES else startup_language
    )
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    set_language(selected_startup_language)
    app.setStyle("Fusion")
    _setup_system_theme()
    apply_global_ui_styles(app)
    _wire_global_style_refresh(app)
    app.setOrganizationName(APP_ORGANIZATION)
    app.setApplicationName(APP_NAME)
    app.setApplicationDisplayName(APP_NAME)
    _install_excepthook()
    startup_validation_error = _validate_startup_workbook_password(app)
    if startup_validation_error is not None:
        return startup_validation_error

    single_instance_lock = _acquire_exe_single_instance_lock()
    if getattr(sys, "frozen", False) and single_instance_lock is None:
        if _signal_existing_instance_to_activate():
            return 0
        return _notify_and_wait(
            app,
            title=APP_NAME,
            message=t("app.already_running"),
            level="info",
        )

    splash = QSplashScreen(
        _build_splash_pixmap(),
        Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint,
    )
    splash.showMessage(
        t("splash.starting"),
        Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter,
        QColor(SPLASH_STATUS_COLOR),
    )
    splash.show()
    app.processEvents()

    # Icon (works in onedir build)
    app.setWindowIcon(QIcon(resource_path("assets/kare-logo.ico")))
    splash.showMessage(
        t("splash.loading_main_window"),
        Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter,
        QColor(SPLASH_STATUS_COLOR),
    )
    app.processEvents()

    def _apply_language_selection(language_code: str) -> bool:
        previous = get_language()
        if language_code.lower() in UI_LANGUAGE_AUTO_ALIASES:
            set_language(UI_LANGUAGE)
        else:
            set_language(language_code)
        return get_language() != previous

    def _on_language_applied(language_code: str) -> None:
        set_ui_language_preference(APP_NAME, ui_language=language_code)
        if _apply_language_selection(language_code):
            window.apply_language_change()

    window = MainWindow(on_language_applied=_on_language_applied)
    activation_server = _install_activation_server(window)

    def _finish_startup() -> None:
        _ = activation_server
        window.show()
        splash.finish(window)

    # Keep splash visible long enough to be noticeable on fast systems.
    QTimer.singleShot(MAIN_SPLASH_MS, _finish_startup)

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())


