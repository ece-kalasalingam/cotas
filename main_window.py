import logging

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QScrollArea, QStackedWidget, QWidget, QHBoxLayout, QToolBar,
    QVBoxLayout, QStatusBar, QLabel
)

from PySide6.QtGui import QAction, QActionGroup, QIcon
from PySide6.QtCore import QSize, QTimer, Qt

# Import modules
from common.constants import (
    MAIN_ACTIVITY_ICON_SIZE,
    MAIN_ACTIVITYBAR_STYLESHEET,
    MAIN_WINDOW_TITLE,
    STATUS_FLASH_TIMEOUT_MS,
    WINDOW_HEIGHT_CAP,
    WINDOW_MIN_HEIGHT,
    WINDOW_MIN_WIDTH,
    WINDOW_TARGET_HEIGHT_RATIO,
    WINDOW_WIDTH_TO_HEIGHT_RATIO,
)
from common.texts import t
from common.toast import show_toast
from common.utils import resource_path
from modules.instructor_module import InstructorModule

_logger = logging.getLogger(__name__)


class _PlaceholderModule(QWidget):
    def __init__(self, title: str):
        super().__init__()
        layout = QVBoxLayout(self)
        label = QLabel(t("module.placeholder", title=title), self)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)


class CourseCoordinatorModule(_PlaceholderModule):
    def __init__(self):
        super().__init__(t("module.course_coordinator"))


class POAnalysisModule(_PlaceholderModule):
    def __init__(self):
        super().__init__(t("module.po_analysis"))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # ----------------------------
        # Window Setup
        # ----------------------------
        self.setWindowTitle(MAIN_WINDOW_TITLE)
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
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        central_layout = QVBoxLayout(container)
        scroll.setWidget(container)
        self.setCentralWidget(scroll)

        #central_container = QWidget()
        #self.setCentralWidget(central_container)

        #central_layout = QVBoxLayout(central_container)
        central_layout.setContentsMargins(0, 0, 0, 0)

        self.work_area = QWidget()
        self.work_layout = QHBoxLayout(self.work_area)
        self.stack = QStackedWidget()
        self.work_layout.addWidget(self.stack)
        
        # Dictionary to keep track of initialized modules
        self.modules = {}
        #self.work_layout.setContentsMargins(30, 30, 30, 30)

        central_layout.addWidget(self.work_area)

        self.current_module = None

        # ----------------------------
        # Activity Bar
        # ----------------------------
        self.activitybar = QToolBar(t("toolbar.navigation"))
        self.activitybar.setMovable(True)
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

        # ----------------------------
        # Status Bar
        # ----------------------------
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage(t("status.ready"))

        # ----------------------------
        # Connect Navigation
        # ----------------------------
        self.action_co_section.triggered.connect(
            lambda: self.load_module(InstructorModule)
        )
        self.action_co_course.triggered.connect(
            lambda: self.load_module(CourseCoordinatorModule)
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
                self.statusBar().showMessage(t("status.ready"))
                return
            self.modules[module_key] = new_module
            self.stack.addWidget(new_module)
            
            # Connect signals if they exist
            signal = getattr(new_module, "status_changed", None)
            if signal:
                signal.connect(lambda msg: self.flash_status(msg))

        # Switch the visible widget
        self.stack.setCurrentWidget(self.modules[module_key])
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
            return

        self.load_module(AboutModule)
