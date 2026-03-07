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

from common.constants import (
    ABOUT_APP_NAME_STYLE,
    ABOUT_BODY_STYLE,
    ABOUT_BODY_GAP_LARGE,
    ABOUT_BODY_GAP_SMALL,
    ABOUT_COPYRIGHT_STYLE,
    ABOUT_HEADER_SPACING,
    ABOUT_META_STYLE,
    ABOUT_LAYOUT_MARGIN,
    ABOUT_LAYOUT_SPACING,
    ABOUT_SUBTITLE_STYLE,
    ABOUT_ICON_SIZE,
    ABOUT_TITLE_SPACING,
    APP_NAME,
    APP_SUBTITLE_TEXT_KEY,
    SYSTEM_VERSION,
)
from common.texts import t
from common.utils import resource_path


class AboutModule(QWidget):
    def __init__(self):
        super().__init__()

        root = QVBoxLayout(self)
        root.setContentsMargins(
            ABOUT_LAYOUT_MARGIN,
            ABOUT_LAYOUT_MARGIN,
            ABOUT_LAYOUT_MARGIN,
            ABOUT_LAYOUT_MARGIN,
        )
        root.setSpacing(ABOUT_LAYOUT_SPACING)

        # -------------------------------------------------
        # Header Section (Icon + Title Block)
        # -------------------------------------------------
        header_layout = QHBoxLayout()
        header_layout.setSpacing(ABOUT_HEADER_SPACING)

        icon_label = QLabel()
        icon_label.setFixedSize(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE)

        icon = QIcon(resource_path("assets/kare-logo.ico"))
        pix = icon.pixmap(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE)
        icon_label.setPixmap(pix)

        header_layout.addWidget(icon_label, 0, Qt.AlignmentFlag.AlignTop)

        title_layout = QVBoxLayout()
        title_layout.setSpacing(ABOUT_TITLE_SPACING)

        app_name = QLabel(APP_NAME)
        app_name.setStyleSheet(ABOUT_APP_NAME_STYLE)

        self.subtitle = QLabel()
        self.subtitle.setStyleSheet(ABOUT_SUBTITLE_STYLE)

        self.version_label = QLabel()
        self.version_label.setStyleSheet(ABOUT_META_STYLE)

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
        self.description.setStyleSheet(ABOUT_BODY_STYLE)

        self.institution = QLabel()
        self.institution.setStyleSheet(ABOUT_BODY_STYLE)

        self.copyright_label = QLabel()
        self.copyright_label.setStyleSheet(ABOUT_COPYRIGHT_STYLE)

        root.addWidget(self.description)
        root.addSpacing(ABOUT_BODY_GAP_LARGE)
        root.addWidget(self.institution)
        root.addSpacing(ABOUT_BODY_GAP_SMALL)
        root.addWidget(self.copyright_label)

        # Push content upward
        spacer = QLabel()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        root.addWidget(spacer)

        self.retranslate_ui()

    def retranslate_ui(self) -> None:
        self.subtitle.setText(t(APP_SUBTITLE_TEXT_KEY))
        self.version_label.setText(t("about.version", version=SYSTEM_VERSION))
        self.description.setText(t("about.description", app_name=APP_NAME))
        self.institution.setText(t("about.institution"))
        self.copyright_label.setText(t("about.copyright", year=datetime.now().year))
