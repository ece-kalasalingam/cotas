from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QToolBar,
    QVBoxLayout, QStatusBar
)
from PySide6.QtGui import QAction, QActionGroup
from PySide6.QtCore import QSize, QTimer, Qt
from qdarktheme._util import get_qdarktheme_root_path
from PySide6.QtCore import QDir
from PySide6.QtGui import QIcon, QAction

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon

# Import modules
from core.resources import resource_path
from modules.co_module import COModule
from modules.co_course_module import COCourseModule
from modules.help_module import HelpModule
from modules.about_module import AboutModule


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # ----------------------------
        # Window Setup
        # ----------------------------
        self.setWindowTitle("CO Attainment")
        self.resize(1100, 700)
        self.setMinimumSize(1000, 650)

        # ----------------------------
        # Central Container
        # ----------------------------
        central_container = QWidget()
        self.setCentralWidget(central_container)

        central_layout = QVBoxLayout(central_container)
        central_layout.setContentsMargins(0, 0, 0, 0)

        self.work_area = QWidget()
        self.work_layout = QHBoxLayout(self.work_area)
        #self.work_layout.setContentsMargins(30, 30, 30, 30)

        central_layout.addWidget(self.work_area)

        self.current_module = None

        # ----------------------------
        # Left Activity Bar
        # ----------------------------
        self.activitybar = QToolBar("Navigation")
        self.activitybar.setMovable(False)
        self.activitybar.setIconSize(QSize(30, 30))
        self.addToolBar(Qt.ToolBarArea.LeftToolBarArea, self.activitybar)
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, self.activitybar)
        self.activitybar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.activitybar.setStyleSheet("QToolBar { spacing: 10px; padding: 5px; }")
        self.activitybar.setStyleSheet("""
            QToolBar {
                spacing: 0px; /* Removes gaps between buttons */
            }
            QToolButton {
                min-width: 80px;  /* Adjust this based on your preferred width */
                max-width: 80px;
                padding: 5px;
                border: none;
                border-radius: 4px;
            }
            QToolButton:hover {
                background-color: rgba(255, 255, 255, 0.1); /* Subtle hover effect */
            }
            QToolButton:checked {
                background-color: rgba(22, 160, 133, 0.2); /* Highlight using your Teal color */
                border-bottom: 2px solid #16A085; /* Visual indicator for active tab */
            }
        """)

        # SVG Icons
        self.action_co_section = QAction(
            QIcon(resource_path("assets/co_section.svg")),
            "Instructor",
            self
        )
        self.action_co_section.setCheckable(True)

        self.action_co_course = QAction(
            QIcon(resource_path("assets/co_course.svg")),
            "CCoordinator",
            self
        )
        self.action_co_course.setCheckable(True)

        self.action_po = QAction(
            QIcon(resource_path("assets/po.svg")),
            "PO Analysis",
            self
        )
        self.action_po.setCheckable(True)

        self.action_downloads = QAction(
            QIcon(resource_path("assets/download.svg")),
            "Downloads",
            self
        )
        self.action_downloads.setCheckable(True)
        self.action_help = QAction(
            QIcon(resource_path("assets/help.svg")),
            "Help",
            self
        )
        self.action_help.setCheckable(True)

        self.action_about = QAction(
            QIcon(resource_path("assets/about.svg")),
            "About",
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

        self.action_co_section.setChecked(True)

        # ----------------------------
        # Status Bar
        # ----------------------------
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

        # ----------------------------
        # Connect Navigation
        # ----------------------------
        self.action_co_section.triggered.connect(
            lambda: self.load_module(COModule)
        )
        self.action_co_course.triggered.connect(
            lambda: self.load_module(COCourseModule)
        )
        self.action_po.triggered.connect(
            lambda: self.load_module(QWidget)
        )
        self.action_help.triggered.connect(
            lambda: self.load_module(HelpModule)
        )
        self.action_about.triggered.connect(
            lambda: self.load_module(AboutModule)
        )

        # Load default module
        self.load_module(COModule)

    # ----------------------------------------------------
    # Module Handling
    # ----------------------------------------------------

    def load_module(self, module_class):
        if self.current_module:
            self.work_layout.removeWidget(self.current_module)
            self.current_module.deleteLater()
            self.current_module = None

        self.current_module = module_class()
        self.work_layout.addWidget(self.current_module)

        # Reset status to default
        self.statusBar().showMessage("Ready")
        # Optional: Connect custom signals if they exist
        signal = getattr(self.current_module, "status_changed", None)
        if signal:
            signal.connect(lambda msg: self.flash_status(msg))
    
    def flash_status(self, message: str, timeout: int = 3000):
        self.statusBar().showMessage(message)
        QTimer.singleShot(timeout, lambda: self.statusBar().showMessage("Ready"))
