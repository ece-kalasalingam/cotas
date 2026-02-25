# modules/help_module.py

import shutil
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFileDialog, QMenu
)
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtCore import Qt

from scripts.utils import resource_path


class HelpModule(QWidget):
    def __init__(self):
        super().__init__()

        self.pdf_path = resource_path("assets/CO_Calculation_Document.pdf")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.pdf_view = QPdfView()
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)

        self.pdf_doc = QPdfDocument(self)
        self.pdf_doc.load(self.pdf_path)

        self.pdf_view.setDocument(self.pdf_doc)

        # Enable custom context menu
        self.pdf_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.pdf_view.customContextMenuRequested.connect(self.show_context_menu)

        layout.addWidget(self.pdf_view)

    # -----------------------------------------------------
    # Context Menu
    # -----------------------------------------------------

    def show_context_menu(self, position):
        menu = QMenu()

        download_action = menu.addAction("Download PDF")
        open_action = menu.addAction("Open in Default Viewer")

        action = menu.exec(self.pdf_view.mapToGlobal(position))

        if action == download_action:
            self.download_pdf()

        elif action == open_action:
            self.open_external()

    # -----------------------------------------------------
    # Download Logic
    # -----------------------------------------------------

    def download_pdf(self):
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save CO Attainment Document",
            "CO_Calculation_Document.pdf",
            "PDF Files (*.pdf)"
        )

        if save_path:
            shutil.copyfile(self.pdf_path, save_path)

    # -----------------------------------------------------
    # Open in System Viewer
    # -----------------------------------------------------

    def open_external(self):
        import os
        os.startfile(self.pdf_path)