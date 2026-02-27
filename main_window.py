from PySide6.QtWidgets import (
    QApplication, QMainWindow, QScrollArea, QStackedWidget, QWidget, QHBoxLayout, QToolBar,
    QVBoxLayout, QStatusBar
)

from PySide6.QtGui import QAction, QActionGroup
from PySide6.QtCore import QSize, QTimer, Qt
from qdarktheme._util import get_qdarktheme_root_path
from PySide6.QtGui import QIcon, QAction

# Import modules
from scripts.utils import resource_path
from components.co_module import COModule
from components.help_module import HelpModule
from components.about_module import AboutModule


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # ----------------------------
        # Window Setup
        # ----------------------------
        self.setWindowTitle("COTAS - CO Attainment")
        screen = QApplication.primaryScreen().geometry()
        s_height = screen.height()

        # 2. Target 80% of the screen height to be safe but look "Full"
        target_h = min(int(s_height * 0.8), 900)
        # 3. Maintain your aesthetic ratio (~1.57)
        target_w = int(target_h * 1.33)

        # 4. Apply
        self.resize(target_w, target_h)
        self.setMinimumSize(850, 640) # Minimum for HD screens

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
        self.activitybar = QToolBar("Navigation")
        self.activitybar.setMovable(False)
        self.activitybar.setIconSize(QSize(30, 30))
        self.addToolBar(Qt.ToolBarArea.LeftToolBarArea, self.activitybar)
        #self.addToolBar(Qt.ToolBarArea.TopToolBarArea, self.activitybar)
        self.activitybar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.activitybar.setStyleSheet("""
            QToolBar {
                spacing: 0px; /* Removes gaps between buttons */
                padding: 5px; 
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

        for action in self.activitybar.actions():
            btn = self.activitybar.widgetForAction(action)
            if btn:
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
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
            lambda: self.load_module(QWidget)
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
        # Use class name as a unique key
        module_key = module_class.__name__
        
        if module_key not in self.modules:
            # Initialize once and add to stack
            new_module = module_class()
            self.modules[module_key] = new_module
            self.stack.addWidget(new_module)
            
            # Connect signals if they exist
            signal = getattr(new_module, "status_changed", None)
            if signal:
                signal.connect(lambda msg: self.flash_status(msg))

        # Switch the visible widget
        self.stack.setCurrentWidget(self.modules[module_key])
        self.statusBar().showMessage("Ready")
    
    def flash_status(self, message: str, timeout: int = 3000):
        self.statusBar().showMessage(message)
        QTimer.singleShot(timeout, lambda: self.statusBar().showMessage("Ready"))
