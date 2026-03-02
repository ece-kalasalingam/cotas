# modules/help_module.py

import logging
import shutil
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QFileDialog,
    QMenu,
    QVBoxLayout,
    QWidget,
)

from common.constants import APP_NAME
from common.texts import t
from common.toast import show_toast
from common.utils import (
    remember_dialog_dir_safe,
    resolve_dialog_start_path,
    resource_path,
)


class HelpModule(QWidget):
    def __init__(self):
        super().__init__()

        self.pdf_path = Path(resource_path("assets/CO_Calculation_Document.pdf"))
        self._pdf_error_shown = False
        self._logger = logging.getLogger(__name__)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.pdf_view = QPdfView()
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)

        self.pdf_doc = QPdfDocument(self)
        self.pdf_view.setDocument(self.pdf_doc)
        self.pdf_doc.statusChanged.connect(self._on_pdf_status_changed)

        # Enable custom context menu
        self.pdf_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.pdf_view.customContextMenuRequested.connect(self.show_context_menu)

        layout.addWidget(self.pdf_view)
        self._load_pdf()

    # -----------------------------------------------------
    # PDF Load
    # -----------------------------------------------------

    def _load_pdf(self) -> None:
        if not self.pdf_path.exists():
            show_toast(
                self,
                t("help.doc_missing_body", path=self.pdf_path),
                title=t("help.doc_missing_title"),
                level="warning",
            )
            return

        self.pdf_doc.load(str(self.pdf_path))

    def _on_pdf_status_changed(self, status: QPdfDocument.Status) -> None:
        if status == QPdfDocument.Status.Ready:
            self._pdf_error_shown = False
            return

        if status == QPdfDocument.Status.Error and not self._pdf_error_shown:
            self._pdf_error_shown = True
            show_toast(
                self,
                t("help.doc_error_body"),
                title=t("help.doc_error_title"),
                level="warning",
            )

    # -----------------------------------------------------
    # Context Menu
    # -----------------------------------------------------

    def show_context_menu(self, position):
        menu = QMenu(self)

        download_action = menu.addAction(t("help.download_pdf"))
        open_action = menu.addAction(t("help.open_default_viewer"))

        action = menu.exec(self.pdf_view.mapToGlobal(position))

        if action == download_action:
            self.download_pdf()
        elif action == open_action:
            self.open_external()

    # -----------------------------------------------------
    # Download Logic
    # -----------------------------------------------------

    def download_pdf(self):
        if not self.pdf_path.exists():
            show_toast(
                self,
                t("help.missing_file_body"),
                title=t("help.missing_file_title"),
                level="warning",
            )
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            t("help.save_title"),
            resolve_dialog_start_path(APP_NAME, t("help.save_default_name")),
            t("help.save_filter_pdf"),
        )

        if save_path:
            try:
                shutil.copyfile(self.pdf_path, save_path)
                remember_dialog_dir_safe(
                    save_path,
                    app_name=APP_NAME,
                    logger=self._logger,
                )
            except OSError as exc:
                show_toast(
                    self,
                    t("help.save_failed_body", error=exc),
                    title=t("help.save_failed_title"),
                    level="error",
                )

    # -----------------------------------------------------
    # Open in System Viewer
    # -----------------------------------------------------

    def open_external(self):
        if not self.pdf_path.exists():
            show_toast(
                self,
                t("help.missing_file_body"),
                title=t("help.missing_file_title"),
                level="warning",
            )
            return

        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.pdf_path)))
        if not opened:
            show_toast(
                self,
                t("help.open_failed_body"),
                title=t("help.open_failed_title"),
                level="warning",
            )
