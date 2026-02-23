# modules/about_module.py

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QFrame, QSizePolicy
)
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt
from datetime import datetime
from core.resources import resource_path
from core.constants import SYSTEM_VERSION, REGULATION_VERSION



class AboutModule(QWidget):
    def __init__(self):
        super().__init__()

        root = QVBoxLayout(self)
        root.setContentsMargins(40, 40, 40, 40)
        root.setSpacing(18)

        # -------------------------------------------------
        # Header Section (Icon + Title Block)
        # -------------------------------------------------
        header_layout = QHBoxLayout()
        header_layout.setSpacing(20)

        icon_label = QLabel()
        icon_label.setFixedSize(72, 72)

        icon = QIcon(resource_path("assets/kare-logo.ico"))
        pix = icon.pixmap(72, 72)
        icon_label.setPixmap(pix)

        header_layout.addWidget(icon_label, 0, Qt.AlignmentFlag.AlignTop)

        title_layout = QVBoxLayout()
        title_layout.setSpacing(4)

        app_name = QLabel("COTAS")
        app_name.setStyleSheet("font-size: 24px; font-weight: 600;")

        subtitle = QLabel("Course Outcome Analysis System")
        subtitle.setStyleSheet("font-size: 15px;")

        version_label = QLabel(
            f"Version {SYSTEM_VERSION} "
        )
        version_label.setStyleSheet("font-size: 12px; color: gray;")

        title_layout.addWidget(app_name)
        title_layout.addWidget(subtitle)
        title_layout.addWidget(version_label)

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
        description = QLabel(
            "COTAS is a software tool designed for "
            "computing Course Outcome (CO) attainment and performing "
            "structured outcome analysis based on direct and indirect assessments."
        )
        description.setWordWrap(True)
        description.setStyleSheet("font-size: 12px;")

        institution = QLabel(
            "Developed at Kalasalingam Academy of Research and Education (KARE)."
        )
        institution.setStyleSheet("font-size: 12px;")

        copyright_label = QLabel(f"© {datetime.now().year} KARE. All rights reserved.")
        copyright_label.setStyleSheet("font-size: 11px; color: gray;")

        root.addWidget(description)
        root.addSpacing(8)
        root.addWidget(institution)
        root.addSpacing(4)
        root.addWidget(copyright_label)

        # Push content upward
        spacer = QLabel()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        root.addWidget(spacer)