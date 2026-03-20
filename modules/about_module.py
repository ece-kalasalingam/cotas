"""About module with engine-based left/right pane composition."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget

from common.constants import (
    ABOUT_CONTRIBUTORS_FILE,
    ABOUT_ICON_SIZE,
    APP_NAME,
    APP_REPOSITORY_URL,
    MODULE_LEFT_PANE_CONTENT_MARGINS,
    MODULE_LEFT_PANE_LAYOUT_SPACING,
    MODULE_LEFT_PANE_SCROLLBAR_GUTTER,
    MODULE_LEFT_PANE_WIDTH_OFFSET,
    SYSTEM_VERSION,
)
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.texts import t
from common.ui_stylings import GLOBAL_QPUSHBUTTON_MIN_WIDTH
from common.utils import resource_path

_LEFT_PANE_WIDTH = GLOBAL_QPUSHBUTTON_MIN_WIDTH + MODULE_LEFT_PANE_WIDTH_OFFSET


class AboutModule(QWidget):
    def __init__(self) -> None:
        super().__init__()

        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                left_width=_LEFT_PANE_WIDTH,
                left_object_name="stepRail",
                right_object_name="coordinatorActiveCard",
                left_content_margins=MODULE_LEFT_PANE_CONTENT_MARGINS,
                left_layout_spacing=MODULE_LEFT_PANE_LAYOUT_SPACING,
                left_scrollbar_gutter=MODULE_LEFT_PANE_SCROLLBAR_GUTTER,
                show_footer=False,
            ),
        )

        left_pane = QWidget()
        left_pane.setObjectName("stepRail")
        left_layout = QVBoxLayout(left_pane)
        left_layout.setContentsMargins(*MODULE_LEFT_PANE_CONTENT_MARGINS)
        left_layout.setSpacing(MODULE_LEFT_PANE_LAYOUT_SPACING)
        self._ui_engine.set_left_widget(left_pane)

        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_pane)
        right_layout.setContentsMargins(16, 12, 16, 12)
        right_layout.setSpacing(10)
        self._ui_engine.set_right_widget(right_pane)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE)
        icon = QIcon(resource_path("assets/kare-logo.ico"))
        self.icon_label.setPixmap(icon.pixmap(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE))
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

        self.left_app_name = QLabel(APP_NAME)
        self.left_app_name.setObjectName("coordinatorTitle")
        self.left_subtitle = QLabel()
        self.left_subtitle.setWordWrap(True)
        self.left_version = QLabel()
        self.left_copyright = QLabel()
        self.left_copyright.setWordWrap(True)

        left_layout.addWidget(self.icon_label)
        left_layout.addWidget(self.left_app_name)
        left_layout.addWidget(self.left_subtitle)
        left_layout.addWidget(self.left_version)
        left_layout.addWidget(self.left_copyright)
        left_layout.addStretch(1)

        self.right_description = QLabel()
        self.right_description.setWordWrap(True)
        right_layout.addWidget(self.right_description)

        divider_two = QFrame()
        divider_two.setFrameShape(QFrame.Shape.HLine)
        divider_two.setFrameShadow(QFrame.Shadow.Sunken)
        right_layout.addWidget(divider_two)

        self.contributors_title = QLabel()
        self.contributors_title.setObjectName("coordinatorTitle")
        self.contributors_body = QLabel()
        self.contributors_body.setWordWrap(True)
        self.contributors_body.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        right_layout.addWidget(self.contributors_title)
        right_layout.addWidget(self.contributors_body)

        divider_three = QFrame()
        divider_three.setFrameShape(QFrame.Shape.HLine)
        divider_three.setFrameShadow(QFrame.Shadow.Sunken)
        right_layout.addWidget(divider_three)

        self.repository_link = QLabel()
        self.repository_link.setOpenExternalLinks(True)
        self.repository_link.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        right_layout.addWidget(self.repository_link)
        right_layout.addStretch(1)

        self.retranslate_ui()

    def _contributors(self) -> list[str]:
        path = Path(resource_path(f"assets/{ABOUT_CONTRIBUTORS_FILE}"))
        if not path.exists():
            return []
        names: list[str] = []
        for line in path.read_text(encoding="utf-8").splitlines():
            value = line.strip()
            if not value or value.startswith("#"):
                continue
            names.append(value)
        return names

    def retranslate_ui(self) -> None:
        current_year = datetime.now().year
        subtitle = t("about.subtitle")
        self.left_subtitle.setText(subtitle)
        self.left_version.setText(t("about.version", version=SYSTEM_VERSION))
        self.left_copyright.setText(t("about.copyright", year=current_year))

        self.right_description.setText(t("about.description", app_name=APP_NAME))
        self.contributors_title.setText(t("about.contributors"))

        contributors = self._contributors()
        if contributors:
            self.contributors_body.setText("\n".join(f"- {name}" for name in contributors))
        else:
            self.contributors_body.setText(t("about.contributors.none"))

        link_text = t("about.repository.link_label")
        self.repository_link.setText(f'<a href="{APP_REPOSITORY_URL}">{link_text}</a>')
