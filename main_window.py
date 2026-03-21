import logging
from collections.abc import Callable
from datetime import datetime
from typing import Any

from PySide6.QtCore import QSize, Qt, QTimer, QUrl
from PySide6.QtGui import QAction, QActionGroup, QDesktopServices, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QMainWindow,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QStatusBar,
    QTabWidget,
    QTextBrowser,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from common.constants import (
    APP_NAME,
    INSTRUCTOR_INFO_TAB_FIXED_HEIGHT,
    MAIN_ACTIVITY_ICON_SIZE,
    MAIN_HIDDEN_ACTIVITY_MODULE_KEYS,
    MAIN_WINDOW_TITLE_TEXT_KEY,
    OUTPUT_LINK_MODE_FILE,
    OUTPUT_LINK_MODE_FOLDER,
    OUTPUT_LINK_SEPARATOR,
)
from common.module_plugins import ModulePluginSpec
from common.output_panel import (
    OutputPanelData,
    open_output_link,
    render_output_panel_html,
)
from common.texts import get_available_languages, get_language, t
from common.toast import show_toast
from common.ui_logging import (
    build_i18n_log_message,
    format_log_line_at,
    parse_i18n_log_message,
    resolve_i18n_log_message,
)
from common.utils import (
    get_ui_language_preference,
    resource_path,
    set_ui_language_preference,
)
from modules.module_catalog import build_module_catalog

WINDOW_TARGET_HEIGHT_RATIO = 0.8
WINDOW_HEIGHT_CAP = 640
WINDOW_WIDTH_TO_HEIGHT_RATIO = 1.57
WINDOW_MIN_WIDTH = 1005
WINDOW_MIN_HEIGHT = 640
STATUS_FLASH_TIMEOUT_MS = 3000

_logger = logging.getLogger(__name__)


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

        self.work_area = QWidget()
        self.work_layout = QHBoxLayout(self.work_area)
        self.stack = QStackedWidget()
        self.work_layout.addWidget(self.stack)
        
        # Dictionary to keep track of initialized modules
        self.modules: dict[str, Any] = {}
        self._module_specs: tuple[ModulePluginSpec, ...] = build_module_catalog()
        self._module_specs_by_key = {spec.key: spec for spec in self._module_specs}
        self._module_actions_by_key: dict[str, QAction] = {}

        central_layout.addWidget(self.work_area)

        self.shared_activity_frame = QFrame()
        self.shared_activity_frame.setObjectName("sharedActivityFrame")
        self.shared_activity_frame.setFixedHeight(INSTRUCTOR_INFO_TAB_FIXED_HEIGHT)
        self.shared_activity_frame.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        shared_layout = QVBoxLayout(self.shared_activity_frame)
        shared_layout.setContentsMargins(0, 0, 0, 0)
        shared_layout.setSpacing(0)

        self.shared_info_tabs = QTabWidget()
        self.shared_info_tabs.setObjectName("sharedInfoTabs")
        self.shared_info_tabs.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.shared_info_tabs.tabBar().setFocusPolicy(Qt.FocusPolicy.NoFocus)

        shared_log_tab = QWidget()
        shared_log_layout = QVBoxLayout(shared_log_tab)
        self.shared_activity_log = QPlainTextEdit()
        self.shared_activity_log.setObjectName("sharedActivityLog")
        self.shared_activity_log.setReadOnly(True)
        self.shared_activity_log.setFrameShape(QFrame.Shape.NoFrame)
        shared_log_layout.addWidget(self.shared_activity_log)

        shared_outputs_tab = QWidget()
        shared_outputs_layout = QVBoxLayout(shared_outputs_tab)
        self.shared_generated_outputs = QTextBrowser()
        self.shared_generated_outputs.setObjectName("sharedGeneratedOutputs")
        self.shared_generated_outputs.setOpenExternalLinks(False)
        self.shared_generated_outputs.setOpenLinks(False)
        self.shared_generated_outputs.setFrameShape(QFrame.Shape.NoFrame)
        self.shared_generated_outputs.anchorClicked.connect(
            lambda url: self._on_shared_output_link_activated(url.toString())
        )
        shared_outputs_layout.addWidget(self.shared_generated_outputs)
        shared_log_layout.setContentsMargins(0, 0, 0, 0)
        shared_outputs_layout.setContentsMargins(0, 0, 0, 0)

        self.shared_info_tabs.addTab(shared_log_tab, t("instructor.log.title"))
        self.shared_info_tabs.addTab(shared_outputs_tab, t("instructor.links.title"))
        shared_layout.addWidget(self.shared_info_tabs)

        shared_row = QHBoxLayout()
        shared_row.addWidget(self.shared_activity_frame)
        central_layout.addLayout(shared_row)
        central_layout.setStretch(0, 1)
        central_layout.setStretch(1, 0)

        self.current_module: Any | None = None
        self._shared_activity_entries: list[dict[str, object]] = []

        # ----------------------------
        # Activity Bar
        # ----------------------------
        self.activitybar = QToolBar(t("toolbar.navigation"))
        self.activitybar.setObjectName("mainActivityBar")
        self.activitybar.setMovable(True)
        self.activitybar.setFloatable(False)
        self.activitybar.setIconSize(QSize(MAIN_ACTIVITY_ICON_SIZE, MAIN_ACTIVITY_ICON_SIZE))
        self.addToolBar(Qt.ToolBarArea.LeftToolBarArea, self.activitybar)
        self.activitybar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)

        # Exclusive selection
        self.nav_group = QActionGroup(self)
        self.nav_group.setExclusive(True)

        for spec in self._module_specs:
            if not spec.show_in_activity_bar:
                continue
            action = QAction(
                QIcon(resource_path(spec.icon_path)),
                t(spec.title_key),
                self,
            )
            action.setCheckable(True)
            action.triggered.connect(lambda _checked=False, key=spec.key: self._load_module_by_key(key))
            self._module_actions_by_key[spec.key] = action
            self.nav_group.addAction(action)
            self.activitybar.addAction(action)

        # Backward-compatible action aliases used by tests/internals.
        self.action_co_section = self._module_actions_by_key["instructor"]
        self.action_co_course = self._module_actions_by_key["coordinator"]
        self.action_po = self._module_actions_by_key["po_analysis"]
        self.action_help = self._module_actions_by_key["help"]
        self.action_about = self._module_actions_by_key["about"]

        for action in self.activitybar.actions():
            btn = self.activitybar.widgetForAction(action)
            if btn:
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.action_co_section.setChecked(True)

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
        self._load_module_by_key("instructor")

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
        self.current_module = current_module
        self.stack.setCurrentWidget(current_module)
        shared_enabled = module_key not in MAIN_HIDDEN_ACTIVITY_MODULE_KEYS
        self.shared_activity_frame.setVisible(shared_enabled)
        set_shared_mode = getattr(current_module, "set_shared_activity_log_mode", None)
        if callable(set_shared_mode):
            set_shared_mode(shared_enabled)
        self._refresh_shared_outputs_html()
        self.statusBar().showMessage(t("status.ready"))

    def _load_module_by_key(self, key: str) -> None:
        spec = self._module_specs_by_key.get(key)
        if spec is None:
            return
        try:
            module_class = spec.class_loader()
        except Exception as exc:
            _logger.exception("Failed to import module plugin '%s'.", key)
            module_label = t(spec.title_key)
            show_toast(
                self,
                t("module.load_failed_body", module=module_label, error=exc),
                title=t("module.load_failed_title"),
                level="error",
            )
            self.flash_status(t("module.load_failed_status", module=module_label))
            return
        self.load_module(module_class)
    
    def flash_status(self, message: str, timeout: int = STATUS_FLASH_TIMEOUT_MS):
        self.statusBar().showMessage(message)
        QTimer.singleShot(timeout, lambda: self.statusBar().showMessage(t("status.ready")))

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
        for spec in self._module_specs:
            action = self._module_actions_by_key.get(spec.key)
            if action is not None:
                action.setText(t(spec.title_key))
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
                    "raw_message": message,
                }
            )
        else:
            key, kwargs, fallback = parsed
            self._shared_activity_entries.append(
                {
                    "timestamp": timestamp,
                    "message": localized,
                    "raw_message": message,
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
            raw_message = entry.get("raw_message")
            if isinstance(text_key, str):
                safe_kwargs = kwargs if isinstance(kwargs, dict) else {}
                try:
                    resolved = t(text_key, **safe_kwargs)
                except Exception:
                    resolved = fallback if isinstance(fallback, str) else str(message or "")
            else:
                if isinstance(raw_message, str):
                    resolved = resolve_i18n_log_message(raw_message)
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
        provider = getattr(widget, "get_shared_outputs_data", None)
        if callable(provider):
            value = provider()
            payload = value if isinstance(value, OutputPanelData) else OutputPanelData(items=tuple())
            self.shared_generated_outputs.setHtml(
                render_output_panel_html(
                    payload,
                    translate=t,
                    output_link_mode_file=OUTPUT_LINK_MODE_FILE,
                    output_link_mode_folder=OUTPUT_LINK_MODE_FOLDER,
                    output_link_separator=OUTPUT_LINK_SEPARATOR,
                )
            )
            return
        self.shared_generated_outputs.setHtml("")

    def _on_shared_output_link_activated(self, href: str) -> None:
        widget = self.stack.currentWidget()
        payload_provider = getattr(widget, "get_shared_outputs_data", None) if widget is not None else None
        payload_value = payload_provider() if callable(payload_provider) else None
        payload = payload_value if isinstance(payload_value, OutputPanelData) else OutputPanelData(items=tuple())
        opened = open_output_link(
            href,
            output_link_mode_folder=OUTPUT_LINK_MODE_FOLDER,
            output_link_separator=OUTPUT_LINK_SEPARATOR,
            open_path=lambda target: QDesktopServices.openUrl(QUrl.fromLocalFile(str(target))),
        )
        if opened:
            return
        show_toast(
            self,
            t(payload.open_failed_key),
            title=t("instructor.msg.error_title"),
            level="error",
        )

    def set_language_switch_enabled(self, enabled: bool) -> None:
        self.language_status_button.setEnabled(enabled)
        self.language_menu.setEnabled(enabled)


