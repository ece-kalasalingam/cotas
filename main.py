import os
import re
import sys
import logging
from PySide6.QtCore import QLockFile, QStandardPaths, Qt, QTimer
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import QApplication, QMessageBox, QSplashScreen
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap

try:
    import qdarktheme  # type: ignore[import-not-found]
except ModuleNotFoundError:
    qdarktheme = None


from common.constants import (
    APP_NAME,
    APP_ORGANIZATION,
    MAIN_SPLASH_MS,
    SINGLE_INSTANCE_ACK_PAYLOAD,
    SINGLE_INSTANCE_ACTIVATE_PAYLOAD,
    SINGLE_INSTANCE_CLIENT_ACK_TIMEOUT_MS,
    SINGLE_INSTANCE_CLIENT_CONNECT_TIMEOUT_MS,
    SINGLE_INSTANCE_CLIENT_WRITE_TIMEOUT_MS,
    QT_ADAPTIVE_STRUCTURE_SENSITIVITY,
    SINGLE_INSTANCE_LOCK_TIMEOUT_MS,
    SINGLE_INSTANCE_SERVER_READ_TIMEOUT_MS,
    SINGLE_INSTANCE_SERVER_WRITE_TIMEOUT_MS,
    SPLASH_BG_COLOR,
    SPLASH_HEIGHT,
    SPLASH_STATUS_COLOR,
    SPLASH_TEXT_COLOR,
    SPLASH_TITLE_FONT_SIZE,
    SPLASH_WIDTH,
    STARTUP_TOAST_DURATION_MS,
    STARTUP_TOAST_QUIT_DELAY_MS,
    THEME_MODE_AUTO,
    THEME_SETUP_DEFER_MS,
    THEME_REFRESH_DEBOUNCE_MS,
    UI_FONT_FAMILY,
    UI_LANGUAGE,
    WIN32_SHOW_WINDOW_RESTORE,
    ensure_workbook_secret_policy,
)
from common.contracts import validate_blueprint_registry_contracts
from common.exceptions import ConfigurationError
from common.crash_reporting import capture_unhandled_exception, has_remote_crash_endpoint
from common.texts import get_language, set_language, t
from common.toast import ToastLevel, show_toast
from common.utils import (
    UI_LANGUAGE_AUTO_ALIASES,
    configure_app_logging,
    get_ui_language_preference,
    resource_path,
    set_ui_language_preference,
)
from main_window import MainWindow


_logger = logging.getLogger(__name__)
_theme_apply_in_progress = False
_theme_refresh_pending = False


def _build_splash_pixmap() -> QPixmap:
    pixmap = QPixmap(SPLASH_WIDTH, SPLASH_HEIGHT)
    pixmap.fill(QColor(SPLASH_BG_COLOR))

    painter = QPainter(pixmap)
    painter.setPen(QColor(SPLASH_TEXT_COLOR))
    painter.setFont(QFont(UI_FONT_FAMILY, SPLASH_TITLE_FONT_SIZE, QFont.Weight.Bold))
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, APP_NAME)
    painter.end()
    return pixmap


def _acquire_exe_single_instance_lock() -> QLockFile | None:
    # Allow multiple instances in dev/python mode.
    if not getattr(sys, "frozen", False):
        return None

    app_data = QStandardPaths.writableLocation(
        QStandardPaths.StandardLocation.AppDataLocation
    )
    os.makedirs(app_data, exist_ok=True)

    lock_path = os.path.join(app_data, f"{APP_NAME}.lock")
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
                socket.waitForReadyRead(SINGLE_INSTANCE_SERVER_READ_TIMEOUT_MS)
                _ = socket.readAll().data()
                _raise_and_activate_window(window)
                socket.write(SINGLE_INSTANCE_ACK_PAYLOAD)
                socket.flush()
                socket.waitForBytesWritten(SINGLE_INSTANCE_SERVER_WRITE_TIMEOUT_MS)
                socket.disconnectFromServer()
                socket.deleteLater()

    server.newConnection.connect(_on_new_connection)
    return server


def _setup_system_theme() -> None:
    global _theme_apply_in_progress
    if qdarktheme is None:
        return
    if _theme_apply_in_progress:
        return
    _theme_apply_in_progress = True
    try:
        qdarktheme.setup_theme(THEME_MODE_AUTO)
    except (RuntimeError, AttributeError, TypeError, ValueError) as exc:
        _logger.warning(
            "Failed to apply qdarktheme.",
            exc_info=exc,
            extra={"error_code": "THEME_APPLY_FAILED"},
        )
    except Exception as exc:
        _logger.exception(
            "Unexpected error while applying qdarktheme.",
            exc_info=exc,
            extra={"error_code": "THEME_APPLY_UNEXPECTED"},
        )
    finally:
        _theme_apply_in_progress = False


def _schedule_system_theme_refresh() -> None:
    global _theme_refresh_pending
    if qdarktheme is None or _theme_apply_in_progress or _theme_refresh_pending:
        return
    _theme_refresh_pending = True

    def _run() -> None:
        global _theme_refresh_pending
        _theme_refresh_pending = False
        _setup_system_theme()

    # Debounce palette bursts to avoid recursive signal storms.
    QTimer.singleShot(THEME_REFRESH_DEBOUNCE_MS, _run)


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
    if startup_language.lower() in UI_LANGUAGE_AUTO_ALIASES:
        set_language(UI_LANGUAGE)
    else:
        set_language(startup_language)
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
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
        # Defer theme setup until after first paint for faster perceived startup.
        QTimer.singleShot(THEME_SETUP_DEFER_MS, _setup_system_theme)

    # Keep splash visible long enough to be noticeable on fast systems.
    QTimer.singleShot(MAIN_SPLASH_MS, _finish_startup)
    if qdarktheme is not None:
        app.paletteChanged.connect(lambda: _schedule_system_theme_refresh())

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
