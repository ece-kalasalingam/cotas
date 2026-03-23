"""About module using a single visible Module UI Engine pane."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from common.constants import (
    ABOUT_CONTRIBUTORS_FILE,
    ABOUT_ICON_SIZE,
    APP_REPOSITORY_URL,
    SYSTEM_VERSION,
)
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.i18n import t
from common.utils import resource_path


class AboutModule(QWidget):
    def __init__(self) -> None:
        super().__init__()

        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                top_object_name="coordinatorActiveCard",
                show_footer=False,
            ),
        )

        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        right_layout = QVBoxLayout(right_pane)
        right_layout.setContentsMargins(16, 12, 16, 12)
        right_layout.setSpacing(10)
        self._ui_engine.set_top_widget(right_pane)

        header_block = QWidget()
        header_layout = QHBoxLayout(header_block)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(14)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE)
        icon = QIcon(resource_path("assets/kare-logo.ico"))
        self.icon_label.setPixmap(icon.pixmap(ABOUT_ICON_SIZE, ABOUT_ICON_SIZE))
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        header_layout.addWidget(self.icon_label, 0, Qt.AlignmentFlag.AlignTop)

        header_text_block = QWidget()
        header_text_layout = QVBoxLayout(header_text_block)
        header_text_layout.setContentsMargins(0, 0, 0, 0)
        header_text_layout.setSpacing(0)

        self.left_app_name = QLabel()
        self.left_app_name.setObjectName("aboutLeftAppName")
        self.left_subtitle = QLabel()
        self.left_subtitle.setObjectName("aboutLeftSubtitle")
        self.left_subtitle.setWordWrap(True)
        self.left_version = QLabel()
        self.left_version.setObjectName("aboutLeftVersion")
        header_text_layout.addWidget(self.left_app_name)
        header_text_layout.addWidget(self.left_subtitle)
        header_text_layout.addWidget(self.left_version)
        header_layout.addWidget(header_text_block, 1)

        right_layout.addWidget(header_block)

        divider_one = QFrame()
        divider_one.setFrameShape(QFrame.Shape.HLine)
        divider_one.setFrameShadow(QFrame.Shadow.Sunken)
        right_layout.addWidget(divider_one)

        self.right_description = QLabel()
        self.right_description.setWordWrap(True)
        right_layout.addWidget(self.right_description)

        self.institution_label = QLabel()
        self.institution_label.setWordWrap(True)
        self.contributors_body = QLabel()
        self.contributors_body.setWordWrap(True)
        self.contributors_body.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        right_layout.addWidget(self.institution_label)
        right_layout.addWidget(self.contributors_body)

        divider_three = QFrame()
        divider_three.setFrameShape(QFrame.Shape.HLine)
        divider_three.setFrameShadow(QFrame.Shadow.Sunken)
        right_layout.addWidget(divider_three)

        self.meta_line = QLabel()
        self.meta_line.setWordWrap(True)
        self.meta_line.setOpenExternalLinks(True)
        self.meta_line.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        right_layout.addWidget(self.meta_line)
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
        self.left_app_name.setText(t("app.main_window_title"))
        self.left_subtitle.setText(t("about.subtitle"))
        self.left_version.setText(t("about.version", version=SYSTEM_VERSION))

        self.right_description.setText(
            t("about.description", app_name=t("app.main_window_title"))
        )
        self.institution_label.setText(t("about.institution"))

        contributors = self._contributors()
        if contributors:
            self.contributors_body.setText("\n".join(f"- {name}" for name in contributors))
        else:
            self.contributors_body.setText(t("about.contributors.none"))

        link_text = t("about.repository.link_label")
        copyright_text = t("about.copyright", year=current_year)
        self.meta_line.setText(
            t(
                "about.meta_line_html",
                copyright=copyright_text,
                url=APP_REPOSITORY_URL,
                link_label=link_text,
            )
        )

