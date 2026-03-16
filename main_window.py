import logging
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QStackedWidget, QWidget, QHBoxLayout, QToolBar,
    QVBoxLayout, QStatusBar, QLabel, QMenu, QPushButton, QFrame, QPlainTextEdit, QTabWidget, QTextBrowser
)

from PySide6.QtGui import QAction, QActionGroup, QIcon, QDesktopServices
from PySide6.QtCore import QSize, QTimer, Qt, QUrl

# Import modules
from common.constants import (
    APP_NAME,
    INSTRUCTOR_CARD_MARGIN,
    MAIN_ACTIVITY_ICON_SIZE,
    MAIN_HIDDEN_ACTIVITY_MODULE_KEYS,
    MAIN_SHARED_TAB_LAYOUT_MARGINS,
    MAIN_SHARED_TAB_LAYOUT_SPACING,
    MAIN_WINDOW_TITLE_TEXT_KEY,
    MAIN_WINDOW_CONTENT_MARGINS,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
)
from common.texts import get_available_languages, get_language, t
from common.toast import show_toast
from common.utils import (
    get_ui_language_preference,
    resource_path,
    set_ui_language_preference,
)
from common.ui_logging import (
    build_i18n_log_message,
    format_log_line_at,
    parse_i18n_log_message,
    resolve_i18n_log_message,
)
from modules.coordinator_module import CoordinatorModule
from modules.instructor_module import InstructorModule


WINDOW_TARGET_HEIGHT_RATIO = 0.8
WINDOW_HEIGHT_CAP = 640
WINDOW_WIDTH_TO_HEIGHT_RATIO = 1.57
WINDOW_MIN_WIDTH = 1005
WINDOW_MIN_HEIGHT = 640
STATUS_FLASH_TIMEOUT_MS = 3000
MAIN_SHARED_INFO_TABS_HEIGHT = 150
MAIN_SHARED_ACTIVITY_FRAME_EXTRA_HEIGHT = 16
MAIN_SHARED_LAYOUT_MARGINS = (0, 8, 0, 8)
MAIN_SHARED_LAYOUT_SPACING = 6
MAIN_SHARED_TAB_FIRST_MARGIN_LEFT = 8
MAIN_SHARED_ACTIVITY_STYLESHEET = """
QFrame#sharedActivityFrame {
    border: none;
    background: transparent;
}
QTabWidget#sharedInfoTabs::pane {
    border: none;
    background: palette(base);
}
QTabWidget#sharedInfoTabs QTabBar::tab:first {
    margin-left: __TAB_MARGIN__px;
}
QPlainTextEdit#sharedActivityLog,
QTextBrowser#sharedGeneratedOutputs {
    border: 1px solid palette(mid);
    border-radius: 8px;
    background: palette(base);
    padding: 8px;
}
"""
MAIN_ACTIVITYBAR_STYLESHEET = """
QToolBar {
    spacing: 0px;
    padding: 5px;
}
QToolButton {
    min-width: 80px;
    max-width: 80px;
    padding: 5px;
    border: none;
    border-radius: 4px;
}
QToolButton:hover {
    background-color: rgba(255, 255, 255, 0.1);
}
QToolButton:checked {
    background-color: rgba(22, 160, 133, 0.2);
    border-bottom: 2px solid #16A085;
}
"""

_logger = logging.getLogger(__name__)


class _PlaceholderModule(QWidget):
    def __init__(self, title_key: str):
        super().__init__()
        self._title_key = title_key
        layout = QVBoxLayout(self)
        self._label = QLabel(self)
        self._label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._label.setWordWrap(True)
        layout.addWidget(self._label)
        self.retranslate_ui()

    def retranslate_ui(self) -> None:
        self._label.setText(t("module.placeholder", title=t(self._title_key)))


class POAnalysisModule(_PlaceholderModule):
    def __init__(self):
        super().__init__("module.po_analysis")


class MainWindow(QMainWindow):
    def __init__(self, on_language_applied: Callable[[str], None] | None = None):
        super().__init__()
        self._on_language_applied = on_language_applied

        # ----------------------------
        # Window Setup
        # ----------------------------
        self.setWindowTitle(t(MAIN_WINDOW_TITLE_TEXT_KEY))
        screen = QApplication.primaryScreen().geometry()
        s_height = screen.height()

        # 2. Target 80% of the screen height to be safe but look "Full"
        target_h = min(int(s_height * WINDOW_TARGET_HEIGHT_RATIO), WINDOW_HEIGHT_CAP)
        # 3. Maintain your aesthetic ratio
        target_w = int(target_h * WINDOW_WIDTH_TO_HEIGHT_RATIO)

        # 4. Apply
        self.resize(target_w, target_h)
        self.setMinimumSize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT) # Minimum for HD screens

        # ----------------------------
        # Central Container
        # ----------------------------
        central_container = QWidget()
        self.setCentralWidget(central_container)
        central_layout = QVBoxLayout(central_container)
        central_layout.setContentsMargins(*MAIN_WINDOW_CONTENT_MARGINS)

        self.work_area = QWidget()
        self.work_layout = QHBoxLayout(self.work_area)
        self.work_layout.setContentsMargins(*MAIN_WINDOW_CONTENT_MARGINS)
        self.stack = QStackedWidget()
        self.work_layout.addWidget(self.stack)
        
        # Dictionary to keep track of initialized modules
        self.modules = {}

        central_layout.addWidget(self.work_area)

        self.shared_activity_frame = QFrame()
        self.shared_activity_frame.setObjectName("sharedActivityFrame")
        shared_tabs_height = MAIN_SHARED_INFO_TABS_HEIGHT
        self.shared_activity_frame.setFixedHeight(shared_tabs_height + MAIN_SHARED_ACTIVITY_FRAME_EXTRA_HEIGHT)
        shared_layout = QVBoxLayout(self.shared_activity_frame)
        shared_layout.setContentsMargins(*MAIN_SHARED_LAYOUT_MARGINS)
        shared_layout.setSpacing(MAIN_SHARED_LAYOUT_SPACING)

        self.shared_info_tabs = QTabWidget()
        self.shared_info_tabs.setObjectName("sharedInfoTabs")
        self.shared_info_tabs.setFixedHeight(shared_tabs_height)
        self.shared_info_tabs.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.shared_info_tabs.tabBar().setFocusPolicy(Qt.FocusPolicy.NoFocus)

        shared_log_tab = QWidget()
        shared_log_layout = QVBoxLayout(shared_log_tab)
        shared_log_layout.setContentsMargins(*MAIN_SHARED_TAB_LAYOUT_MARGINS)
        shared_log_layout.setSpacing(MAIN_SHARED_TAB_LAYOUT_SPACING)
        self.shared_activity_log = QPlainTextEdit()
        self.shared_activity_log.setObjectName("sharedActivityLog")
        self.shared_activity_log.setReadOnly(True)
        self.shared_activity_log.setFrameShape(QFrame.Shape.NoFrame)
        shared_log_layout.addWidget(self.shared_activity_log)

        shared_outputs_tab = QWidget()
        shared_outputs_layout = QVBoxLayout(shared_outputs_tab)
        shared_outputs_layout.setContentsMargins(*MAIN_SHARED_TAB_LAYOUT_MARGINS)
        shared_outputs_layout.setSpacing(MAIN_SHARED_TAB_LAYOUT_SPACING)
        self.shared_generated_outputs = QTextBrowser()
        self.shared_generated_outputs.setObjectName("sharedGeneratedOutputs")
        self.shared_generated_outputs.setOpenExternalLinks(False)
        self.shared_generated_outputs.setOpenLinks(False)
        self.shared_generated_outputs.setFrameShape(QFrame.Shape.NoFrame)
        self.shared_generated_outputs.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.shared_generated_outputs.anchorClicked.connect(
            lambda url: self._on_shared_output_link_activated(url.toString())
        )
        shared_outputs_layout.addWidget(self.shared_generated_outputs)

        self.shared_info_tabs.addTab(shared_log_tab, t("instructor.log.title"))
        self.shared_info_tabs.addTab(shared_outputs_tab, t("instructor.links.title"))
        shared_layout.addWidget(self.shared_info_tabs)

        shared_row = QHBoxLayout()
        shared_row.setContentsMargins(INSTRUCTOR_CARD_MARGIN, 0, INSTRUCTOR_CARD_MARGIN, 0)
        shared_row.setSpacing(0)
        shared_row.addWidget(self.shared_activity_frame)
        central_layout.addLayout(shared_row)

        self.current_module = None
        self._shared_activity_entries: list[dict[str, object]] = []

        # ----------------------------
        # Activity Bar
        # ----------------------------
        self.activitybar = QToolBar(t("toolbar.navigation"))
        self.activitybar.setMovable(True)
        self.activitybar.setFloatable(False)
        self.activitybar.setIconSize(QSize(MAIN_ACTIVITY_ICON_SIZE, MAIN_ACTIVITY_ICON_SIZE))
        self.addToolBar(Qt.ToolBarArea.LeftToolBarArea, self.activitybar)
        self.activitybar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.activitybar.setStyleSheet(MAIN_ACTIVITYBAR_STYLESHEET)

        # SVG Icons
        self.action_co_section = QAction(
            QIcon(resource_path("assets/co_section.svg")),
            t("module.instructor"),
            self
        )
        self.action_co_section.setCheckable(True)

        self.action_co_course = QAction(
            QIcon(resource_path("assets/co_course.svg")),
            t("module.coordinator_short"),
            self
        )
        self.action_co_course.setCheckable(True)

        self.action_po = QAction(
            QIcon(resource_path("assets/po.svg")),
            t("module.po_analysis"),
            self
        )
        self.action_po.setCheckable(True)

        self.action_help = QAction(
            QIcon(resource_path("assets/help.svg")),
            t("nav.help"),
            self
        )
        self.action_help.setCheckable(True)

        self.action_about = QAction(
            QIcon(resource_path("assets/about.svg")),
            t("nav.about"),
            self
        )
        self.action_about.setCheckable(True)

        # Exclusive selection
        self.nav_group = QActionGroup(self)
        self.nav_group.setExclusive(True)

        for action in (
            self.action_co_section,
            self.action_co_course,
            self.action_po,
            self.action_help,
            self.action_about,
        ):
            self.nav_group.addAction(action)
            self.activitybar.addAction(action)

        for action in self.activitybar.actions():
            btn = self.activitybar.widgetForAction(action)
            if btn:
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.action_co_section.setChecked(True)

        self.setStyleSheet(
            MAIN_SHARED_ACTIVITY_STYLESHEET.replace(
                "__TAB_MARGIN__",
                str(MAIN_SHARED_TAB_FIRST_MARGIN_LEFT),
            )
        )

        self.language_menu = QMenu(self)
        self.language_action_group = QActionGroup(self.language_menu)
        self.language_action_group.setExclusive(True)
        self.language_menu.triggered.connect(self._on_language_menu_action)

        # ----------------------------
        # Status Bar
        # ----------------------------
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage(t("status.ready"))
        self.language_status_button = QPushButton(self)
        self.language_status_button.setFlat(True)
        self.language_status_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.language_status_button.clicked.connect(self._show_language_menu_from_statusbar)
        self.statusBar().addPermanentWidget(self.language_status_button)

        # ----------------------------
        # Connect Navigation
        # ----------------------------
        self.action_co_section.triggered.connect(
            lambda: self.load_module(InstructorModule)
        )
        self.action_co_course.triggered.connect(
            lambda: self.load_module(CoordinatorModule)
        )
        self.action_po.triggered.connect(
            lambda: self.load_module(POAnalysisModule)
        )
        self.action_help.triggered.connect(
            self._load_help_module
        )
        self.action_about.triggered.connect(
            self._load_about_module
        )

        self._refresh_language_switcher()
        self._refresh_shared_activity_texts()
        self._append_shared_activity_log(
            build_i18n_log_message(
                "instructor.log.ready",
                fallback=t("instructor.log.ready"),
            )
        )
        self._refresh_shared_outputs_html()

        # Load default module
        self.load_module(InstructorModule)

    # ----------------------------------------------------
    # Module Handling
    # ----------------------------------------------------

    def load_module(self, module_class):
        # Use class name as a unique key
        module_key = module_class.__name__
        
        if module_key not in self.modules:
            # Initialize once and add to stack
            try:
                new_module = module_class()
            except Exception as exc:
                _logger.exception("Failed to load module '%s'.", module_key)
                show_toast(
                    self,
                    t("module.load_failed_body", module=module_key, error=exc),
                    title=t("module.load_failed_title"),
                    level="error",
                )
                self.flash_status(t("module.load_failed_status", module=module_key))
                return
            self.modules[module_key] = new_module
            self.stack.addWidget(new_module)
            
            # Connect signals if they exist
            signal = getattr(new_module, "status_changed", None)
            if signal:
                signal.connect(self._on_module_status_changed)
        # Switch the visible widget
        current_module = self.modules[module_key]
        self.stack.setCurrentWidget(current_module)
        shared_enabled = module_key not in MAIN_HIDDEN_ACTIVITY_MODULE_KEYS
        self.shared_activity_frame.setVisible(shared_enabled)
        set_shared_mode = getattr(current_module, "set_shared_activity_log_mode", None)
        if callable(set_shared_mode):
            set_shared_mode(shared_enabled)
        self._refresh_shared_outputs_html()
        self.statusBar().showMessage(t("status.ready"))
    
    def flash_status(self, message: str, timeout: int = STATUS_FLASH_TIMEOUT_MS):
        self.statusBar().showMessage(message)
        QTimer.singleShot(timeout, lambda: self.statusBar().showMessage(t("status.ready")))

    def _load_help_module(self):
        # Lazy import avoids QtPdf startup cost until Help is actually opened.
        try:
            from modules.help_module import HelpModule
        except Exception as exc:
            _logger.exception("Failed to import Help module.")
            show_toast(
                self,
                t("module.load_failed_body", module=t("nav.help"), error=exc),
                title=t("module.load_failed_title"),
                level="error",
            )
            self.flash_status(t("module.load_failed_status", module=t("nav.help")))
            return

        self.load_module(HelpModule)

    def _load_about_module(self):
        try:
            from modules.about_module import AboutModule
        except Exception as exc:
            _logger.exception("Failed to import About module.")
            show_toast(
                self,
                t("module.load_failed_body", module=t("nav.about"), error=exc),
                title=t("module.load_failed_title"),
                level="error",
            )
            self.flash_status(t("module.load_failed_status", module=t("nav.about")))
            return

        self.load_module(AboutModule)

    def _refresh_language_switcher(self) -> None:
        preferred = get_ui_language_preference(APP_NAME)
        active_lang = get_language()
        language_labels = dict(get_available_languages())
        active_label = language_labels.get(active_lang, active_lang.upper())

        self.language_status_button.setText(
            t("language.switcher.button", language=active_label)
        )
        self._rebuild_language_menu(preferred)

    def _rebuild_language_menu(self, preferred: str) -> None:
        self.language_menu.clear()
        self.language_action_group = QActionGroup(self.language_menu)
        self.language_action_group.setExclusive(True)

        for code, label in get_available_languages():
            action = self.language_menu.addAction(label)
            action.setData(code)
            action.setCheckable(True)
            action.setChecked(preferred == code)
            self.language_action_group.addAction(action)

    def _show_language_menu_from_statusbar(self) -> None:
        pos = self.language_status_button.mapToGlobal(self.language_status_button.rect().topLeft())
        self.language_menu.popup(pos)

    def _on_language_menu_action(self, action: QAction) -> None:
        if not self.language_status_button.isEnabled():
            return
        language_code = action.data()
        if not isinstance(language_code, str):
            return

        set_ui_language_preference(APP_NAME, ui_language=language_code)
        self.flash_status(t("language.switcher.applied_status", language=action.text()))
        if callable(self._on_language_applied):
            self._on_language_applied(language_code)

    def apply_language_change(self) -> None:
        self.setWindowTitle(t(MAIN_WINDOW_TITLE_TEXT_KEY))
        self.activitybar.setWindowTitle(t("toolbar.navigation"))
        self.action_co_section.setText(t("module.instructor"))
        self.action_co_course.setText(t("module.coordinator_short"))
        self.action_po.setText(t("module.po_analysis"))
        self.action_help.setText(t("nav.help"))
        self.action_about.setText(t("nav.about"))
        self._refresh_language_switcher()
        self._refresh_shared_activity_texts()
        self._rerender_shared_activity_log()

        for module in self.modules.values():
            retranslate = getattr(module, "retranslate_ui", None)
            if callable(retranslate):
                retranslate()
                continue
            refresh = getattr(module, "_refresh_ui", None)
            if callable(refresh):
                refresh()

        self._refresh_shared_outputs_html()
        self.statusBar().showMessage(t("status.ready"))

    def _refresh_shared_activity_texts(self) -> None:
        self.shared_info_tabs.setTabText(0, t("instructor.log.title"))
        self.shared_info_tabs.setTabText(1, t("instructor.links.title"))

    def _on_module_status_changed(self, message: str) -> None:
        self.flash_status(resolve_i18n_log_message(message))
        self._append_shared_activity_log(message)
        self._refresh_shared_outputs_html()

    def _append_shared_activity_log(self, message: str) -> None:
        parsed = parse_i18n_log_message(message)
        localized = resolve_i18n_log_message(message)
        timestamp = datetime.now()
        if parsed is None:
            self._shared_activity_entries.append(
                {
                    "timestamp": timestamp,
                    "message": localized,
                }
            )
        else:
            key, kwargs, fallback = parsed
            self._shared_activity_entries.append(
                {
                    "timestamp": timestamp,
                    "message": localized,
                    "text_key": key,
                    "kwargs": kwargs,
                    "fallback": fallback,
                }
            )
        line = format_log_line_at(localized, timestamp=timestamp)
        if line is None:
            return
        self.shared_activity_log.appendPlainText(line)

    def _rerender_shared_activity_log(self) -> None:
        self.shared_activity_log.clear()
        for entry in self._shared_activity_entries:
            timestamp = entry.get("timestamp")
            text_key = entry.get("text_key")
            fallback = entry.get("fallback")
            kwargs = entry.get("kwargs")
            message = entry.get("message")
            if isinstance(text_key, str):
                safe_kwargs = kwargs if isinstance(kwargs, dict) else {}
                try:
                    resolved = t(text_key, **safe_kwargs)
                except Exception:
                    resolved = fallback if isinstance(fallback, str) else str(message or "")
            else:
                resolved = str(message or "")

            ts = timestamp if isinstance(timestamp, datetime) else None
            line = format_log_line_at(resolved, timestamp=ts)
            if line is None:
                continue
            self.shared_activity_log.appendPlainText(line)

    def _refresh_shared_outputs_html(self) -> None:
        widget = self.stack.currentWidget()
        if widget is None:
            self.shared_generated_outputs.setHtml("")
            return
        provider = getattr(widget, "get_shared_outputs_html", None)
        if callable(provider):
            self.shared_generated_outputs.setHtml(provider())
            return
        self.shared_generated_outputs.setHtml("")

    def _on_shared_output_link_activated(self, href: str) -> None:
        mode, _, raw_path = href.partition(OUTPUT_LINK_SEPARATOR)
        path = raw_path.strip()
        if not path:
            return
        target = Path(path).parent if mode == OUTPUT_LINK_MODE_FOLDER else Path(path)
        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
        if opened:
            return
        show_toast(
            self,
            t("instructor.links.open_failed"),
            title=t("instructor.msg.error_title"),
            level="error",
        )

    def set_language_switch_enabled(self, enabled: bool) -> None:
        self.language_status_button.setEnabled(enabled)
        self.language_menu.setEnabled(enabled)


