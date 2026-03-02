import os
import sys
import logging
from PySide6.QtCore import QLockFile, QStandardPaths, Qt, QTimer
from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap

try:
    import qdarktheme  # type: ignore[import-not-found]
except ModuleNotFoundError:
    qdarktheme = None


from common.constants import (
    APP_NAME,
    APP_ORGANIZATION,
    MAIN_SPLASH_MS,
    SPLASH_BG_COLOR,
    SPLASH_HEIGHT,
    SPLASH_STATUS_COLOR,
    SPLASH_TEXT_COLOR,
    SPLASH_TITLE_FONT_SIZE,
    SPLASH_WIDTH,
    UI_FONT_FAMILY,
    UI_LANGUAGE,
)
from common.texts import set_language, set_language_from_system, t
from common.toast import show_toast
from common.utils import configure_app_logging, resource_path
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
    if not lock.tryLock(0):
        return None

    return lock


def _setup_theme() -> None:
    global _theme_apply_in_progress
    if qdarktheme is None:
        return
    if _theme_apply_in_progress:
        return
    _theme_apply_in_progress = True
    try:
        qdarktheme.setup_theme("auto")
    except Exception:
        _logger.exception("Failed to apply qdarktheme.")
    finally:
        _theme_apply_in_progress = False


def _schedule_theme_refresh() -> None:
    global _theme_refresh_pending
    if qdarktheme is None or _theme_apply_in_progress or _theme_refresh_pending:
        return
    _theme_refresh_pending = True

    def _run() -> None:
        global _theme_refresh_pending
        _theme_refresh_pending = False
        _setup_theme()

    # Debounce palette bursts to avoid recursive signal storms.
    QTimer.singleShot(120, _run)


def _install_excepthook() -> None:
    previous_hook = sys.excepthook

    def _hook(exc_type, exc_value, exc_traceback):
        _logger.exception(
            "Unhandled exception in application.",
            exc_info=(exc_type, exc_value, exc_traceback),
        )
        try:
            show_toast(None, t("app.unexpected_error"), title=APP_NAME, level="error")
        except Exception:
            pass
        previous_hook(exc_type, exc_value, exc_traceback)

    sys.excepthook = _hook


def _notify_and_wait(app: QApplication, *, title: str, message: str, level: str) -> int:
    show_toast(None, message, title=title, level=level, duration_ms=2200)
    QTimer.singleShot(2300, app.quit)
    return app.exec()


def main() -> int:
    os.environ["QT_ADAPTIVE_STRUCTURE_SENSITIVITY"] = "1"
    configure_app_logging(APP_NAME)
    if UI_LANGUAGE.lower() in {"auto", "system", "os"}:
        set_language_from_system()
    else:
        set_language(UI_LANGUAGE)
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setOrganizationName(APP_ORGANIZATION)
    app.setApplicationName(APP_NAME)
    _install_excepthook()

    single_instance_lock = _acquire_exe_single_instance_lock()
    if getattr(sys, "frozen", False) and single_instance_lock is None:
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

    window = MainWindow()

    def _finish_startup() -> None:
        window.show()
        splash.finish(window)
        # Defer theme setup until after first paint for faster perceived startup.
        app.setStyle("Fusion")
        QTimer.singleShot(0, _setup_theme)

    # Keep splash visible long enough to be noticeable on fast systems.
    QTimer.singleShot(MAIN_SPLASH_MS, _finish_startup)
    if qdarktheme is not None:
        app.paletteChanged.connect(lambda: _schedule_theme_refresh())

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
