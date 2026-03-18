# modules/about_module.py

from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from common.constants import ABOUT_ICON_SIZE, APP_NAME, SYSTEM_VERSION
from common.texts import t
from common.utils import resource_path


class AboutModule(QWidget):
    def __init__(self):
        super().__init__()

        root = QVBoxLayout(self)

        # -------------------------------------------------
        # Header Section (Icon + Title Block)
        # -------------------------------------------------
        header_layout = QHBoxLayout()

        icon_label = QLabel()
        icon_label.setFixedSize(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE)

        icon = QIcon(resource_path("assets/kare-logo.ico"))
        pix = icon.pixmap(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE)
        icon_label.setPixmap(pix)

        header_layout.addWidget(icon_label, 0, Qt.AlignmentFlag.AlignTop)

        title_layout = QVBoxLayout()

        app_name = QLabel(APP_NAME)

        self.subtitle = QLabel()

        self.version_label = QLabel()

        title_layout.addWidget(app_name)
        title_layout.addWidget(self.subtitle)
        title_layout.addWidget(self.version_label)

        header_layout.addLayout(title_layout)
        header_layout.addStretch()

        root.addLayout(header_layout)

        # Divider
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        root.addWidget(divider)

        # -------------------------------------------------
        # Body Information
        # -------------------------------------------------
        self.description = QLabel()
        self.description.setWordWrap(True)

        self.institution = QLabel()

        self.copyright_label = QLabel()

        root.addWidget(self.description)
        root.addWidget(self.institution)
        root.addWidget(self.copyright_label)

        # Push content upward
        spacer = QLabel()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        root.addWidget(spacer)

        self.retranslate_ui()

    def retranslate_ui(self) -> None:
        self.subtitle.setText(t("about.subtitle"))
        self.version_label.setText(t("about.version", version=SYSTEM_VERSION))
        self.description.setText(t("about.description", app_name=APP_NAME))
        self.institution.setText(t("about.institution"))
        self.copyright_label.setText(t("about.copyright", year=datetime.now().year))

