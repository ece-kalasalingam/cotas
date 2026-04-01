"""Reusable removable file row widget for drag-and-drop lists."""

from __future__ import annotations

import re
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStyle,
    QWidget,
)

from common.constants import (
    FILE_ITEM_REMOVE_BUTTON_ICON_SIZE,
    FILE_ITEM_REMOVE_BUTTON_SIZE,
)
from common.ui_stylings import (
    FILE_ITEM_LAYOUT_MARGINS,
    FILE_ITEM_LAYOUT_SPACING,
)


class ElidedFileNameLabel(QLabel):
    """Filename label that middle-elides text when width is constrained."""

    def __init__(self, text: str, parent: QWidget | None = None) -> None:
        """Init.
        
        Args:
            text: Parameter value (str).
            parent: Parameter value (QWidget | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        super().__init__(parent)
        self._full_text = text
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setMinimumWidth(0)
        self._apply_elided_text()

    def resizeEvent(self, event) -> None:
        """Resizeevent.
        
        Args:
            event: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        super().resizeEvent(event)
        self._apply_elided_text()

    def _apply_elided_text(self) -> None:
        """Apply elided text.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        width = self.contentsRect().width()
        if width <= 0:
            return
        super().setText(
            self.fontMetrics().elidedText(
                self._full_text,
                Qt.TextElideMode.ElideMiddle,
                width,
            )
        )


class RemovableFileItemWidget(QWidget):
    """File list row with a remove (trash) button."""

    removed = Signal(str)

    def __init__(
        self,
        file_path: str,
        *,
        remove_fallback_text: str = "Remove",
        open_file_fallback_text: str = "Open file",
        open_folder_fallback_text: str = "Open folder",
        open_file_tooltip: str = "Open File",
        open_folder_tooltip: str = "Open Folder",
        remove_tooltip: str = "Remove File",
        parent: QWidget | None = None,
    ) -> None:
        """Init.
        
        Args:
            file_path: Parameter value (str).
            remove_fallback_text: Parameter value (str).
            open_file_fallback_text: Parameter value (str).
            open_folder_fallback_text: Parameter value (str).
            open_file_tooltip: Parameter value (str).
            open_folder_tooltip: Parameter value (str).
            remove_tooltip: Parameter value (str).
            parent: Parameter value (QWidget | None).
        
        Returns:
            None.
        
        Raises:
            None.
        """
        super().__init__(parent)
        self.file_path = file_path
        self._local_path = self._normalize_local_path(file_path)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(*FILE_ITEM_LAYOUT_MARGINS)
        layout.setSpacing(FILE_ITEM_LAYOUT_SPACING)

        file_name = Path(file_path).name
        name_label = ElidedFileNameLabel(file_name)
        name_label.setToolTip(file_path)
        layout.addWidget(name_label, 1)

        self.open_file_btn = QPushButton()
        self.open_file_btn.setObjectName("coordinatorFileOpenButton")
        self.open_file_btn.setFixedSize(
            FILE_ITEM_REMOVE_BUTTON_SIZE,
            FILE_ITEM_REMOVE_BUTTON_SIZE,
        )
        self.open_file_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        file_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)
        if not file_icon.isNull():
            self.open_file_btn.setIcon(file_icon)
            self.open_file_btn.setIconSize(
                QSize(
                    FILE_ITEM_REMOVE_BUTTON_ICON_SIZE,
                    FILE_ITEM_REMOVE_BUTTON_ICON_SIZE,
                )
            )
        else:
            self.open_file_btn.setText(open_file_fallback_text)
        self.open_file_btn.setEnabled(self._local_path is not None)
        self.open_file_btn.setToolTip(open_file_tooltip)
        self.open_file_btn.clicked.connect(self._open_file)
        layout.addWidget(self.open_file_btn, 0)

        self.open_folder_btn = QPushButton()
        self.open_folder_btn.setObjectName("coordinatorFolderOpenButton")
        self.open_folder_btn.setFixedSize(
            FILE_ITEM_REMOVE_BUTTON_SIZE,
            FILE_ITEM_REMOVE_BUTTON_SIZE,
        )
        self.open_folder_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        folder_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon)
        if not folder_icon.isNull():
            self.open_folder_btn.setIcon(folder_icon)
            self.open_folder_btn.setIconSize(
                QSize(
                    FILE_ITEM_REMOVE_BUTTON_ICON_SIZE,
                    FILE_ITEM_REMOVE_BUTTON_ICON_SIZE,
                )
            )
        else:
            self.open_folder_btn.setText(open_folder_fallback_text)
        self.open_folder_btn.setEnabled(self._local_path is not None)
        self.open_folder_btn.setToolTip(open_folder_tooltip)
        self.open_folder_btn.clicked.connect(self._open_folder)
        layout.addWidget(self.open_folder_btn, 0)

        self.remove_btn = QPushButton()
        self.remove_btn.setObjectName("coordinatorFileRemoveButton")
        self.remove_btn.setFixedSize(
            FILE_ITEM_REMOVE_BUTTON_SIZE,
            FILE_ITEM_REMOVE_BUTTON_SIZE,
        )
        self.remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon)
        if not icon.isNull():
            self.remove_btn.setIcon(icon)
            self.remove_btn.setIconSize(
                QSize(
                    FILE_ITEM_REMOVE_BUTTON_ICON_SIZE,
                    FILE_ITEM_REMOVE_BUTTON_ICON_SIZE,
                )
            )
        else:
            self.remove_btn.setText(remove_fallback_text)
        self.remove_btn.setToolTip(remove_tooltip)
        self.remove_btn.clicked.connect(lambda: self.removed.emit(self.file_path))
        layout.addWidget(self.remove_btn, 0)

    @staticmethod
    def _normalize_local_path(file_path: str) -> str | None:
        """Normalize local path.
        
        Args:
            file_path: Parameter value (str).
        
        Returns:
            str | None: Return value.
        
        Raises:
            None.
        """
        value = file_path.strip()
        if not value:
            return None
        if value.startswith("file://"):
            url = QUrl(value)
            local = url.toLocalFile()
            return local or None
        if "://" in value:
            return None
        if re.match(r"^[a-zA-Z]:[\\/]", value) or value.startswith("\\\\") or value.startswith("/"):
            return value
        return None

    def _open_file(self) -> None:
        """Open file.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self._local_path is None:
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(self._local_path))

    def _open_folder(self) -> None:
        """Open folder.
        
        Args:
            None.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        if self._local_path is None:
            return
        parent = Path(self._local_path).parent
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(parent)))

