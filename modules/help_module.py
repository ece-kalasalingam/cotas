# modules/help_module.py

import logging
import shutil
from pathlib import Path

from PySide6.QtCore import Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import QFileDialog, QMenu, QStyleFactory, QVBoxLayout, QWidget

from common.constants import (
    APP_NAME,
)
from common.module_messages import build_status_message as _build_status_message
from common.module_messages import default_messages_namespace as _default_messages_namespace
from common.module_messages import publish_status_key as _publish_status_key_impl
from common.module_messages import show_toast_key as _show_toast_key
from common.module_ui_engine import ModuleUIEngine, ModuleUIEngineConfig
from common.i18n import t
from common.utils import (
    log_process_message,
    remember_dialog_dir_safe,
    resolve_dialog_start_path,
    resource_path,
)


class _LogSink:
    def appendPlainText(self, _text: str) -> None:  # noqa: N802 - Qt-style name
        return

    def clear(self) -> None:
        return


def _messages_namespace() -> dict[str, object]:
    return dict(_default_messages_namespace(translate=t))


class HelpModule(QWidget):
    status_changed = Signal(str)

    def __init__(self):
        super().__init__()

        self.pdf_path = Path(resource_path("assets/CO_Calculation_Document.pdf"))
        self._pdf_error_shown = False
        self._logger = logging.getLogger(__name__)
        self._ui_log_handler: object | None = None
        self._user_log_entries: list[dict[str, object]] = []
        self.user_log_view = _LogSink()

        self._ui_engine = ModuleUIEngine(
            self,
            config=ModuleUIEngineConfig(
                top_object_name="coordinatorActiveCard",
                show_footer=False,
            ),
        )
        right_pane = QWidget()
        right_pane.setObjectName("coordinatorActiveCard")
        layout = QVBoxLayout(right_pane)
        layout.setContentsMargins(0, 0, 0, 0)
        self._ui_engine.set_top_widget(right_pane)

        self.pdf_view = QPdfView()
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)

        self.pdf_doc = QPdfDocument(self)
        self.pdf_view.setDocument(self.pdf_doc)
        self.pdf_doc.statusChanged.connect(self._on_pdf_status_changed)

        # Enable custom context menu
        self.pdf_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.pdf_view.customContextMenuRequested.connect(self.show_context_menu)

        layout.addWidget(self.pdf_view)
        self._load_pdf()

    def _emit_status_key(self, key: str, **kwargs: object) -> None:
        _publish_status_key_impl(self, key, ns=_messages_namespace(), **kwargs)

    # -----------------------------------------------------
    # PDF Load
    # -----------------------------------------------------

    def _load_pdf(self) -> None:
        if not self.pdf_path.exists():
            self._logger.warning("Help PDF is missing at path: %s", self.pdf_path)
            _show_toast_key(
                self,
                text_key="help.doc_missing_body",
                title_key="help.doc_missing_title",
                translate=t,
                level="warning",
                text_kwargs={"path": self.pdf_path},
            )
            self._emit_status_key("help.status.doc_missing")
            return

        self.pdf_doc.load(str(self.pdf_path))
        log_process_message(
            "loading help PDF",
            logger=self._logger,
            success_message="loading help PDF completed successfully.",
            user_success_message=_build_status_message(
                "help.status.doc_loaded",
                translate=t,
            ),
        )
        self._emit_status_key("help.status.doc_loaded")

    def _on_pdf_status_changed(self, status: QPdfDocument.Status) -> None:
        if status == QPdfDocument.Status.Ready:
            self._pdf_error_shown = False
            return

        if status == QPdfDocument.Status.Error and not self._pdf_error_shown:
            self._pdf_error_shown = True
            self._logger.warning("Help PDF failed to load. status=%s", status)
            _show_toast_key(
                self,
                text_key="help.doc_error_body",
                title_key="help.doc_error_title",
                translate=t,
                level="warning",
            )
            self._emit_status_key("help.status.doc_error")

    # -----------------------------------------------------
    # Context Menu
    # -----------------------------------------------------

    def show_context_menu(self, position):
        # Keep native OS rendering.
        menu = QMenu(self.pdf_view)
        fusion = QStyleFactory.create("Fusion")
        if fusion is not None and hasattr(menu, "setStyle"):
            menu.setStyle(fusion)

        download_action = menu.addAction(t("help.download_pdf"))
        open_action = menu.addAction(t("help.open_default_viewer"))

        action = menu.exec(self.pdf_view.viewport().mapToGlobal(position))

        if action == download_action:
            self.download_pdf()
        elif action == open_action:
            self.open_external()

    # -----------------------------------------------------
    # Download Logic
    # -----------------------------------------------------

    def download_pdf(self):
        process_name = "saving help PDF"
        if not self.pdf_path.exists():
            self._logger.warning("Help PDF save requested but source file is missing.")
            _show_toast_key(
                self,
                text_key="help.status.file_missing",
                title_key="help.missing_file_title",
                translate=t,
                level="warning",
            )
            self._emit_status_key("help.status.file_missing")
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
                log_process_message(
                    process_name,
                    logger=self._logger,
                    error=exc,
                    user_error_message=_build_status_message(
                        "help.status.save_failed",
                        translate=t,
                    ),
                )
                _show_toast_key(
                    self,
                    text_key="help.status.save_failed",
                    title_key="help.save_failed_title",
                    translate=t,
                    level="error",
                )
                self._emit_status_key("help.status.save_failed")
                return

            log_process_message(
                process_name,
                logger=self._logger,
                success_message=f"{process_name} completed successfully.",
                user_success_message=_build_status_message(
                    "help.status.save_success",
                    translate=t,
                ),
            )
            _show_toast_key(
                self,
                text_key="help.status.save_success",
                title_key="help.save_success_title",
                translate=t,
                level="success",
            )
            self._emit_status_key("help.status.save_success")

    # -----------------------------------------------------
    # Open in System Viewer
    # -----------------------------------------------------

    def open_external(self):
        process_name = "opening help PDF in default viewer"
        if not self.pdf_path.exists():
            self._logger.warning("Help PDF open requested but source file is missing.")
            _show_toast_key(
                self,
                text_key="help.status.file_missing",
                title_key="help.missing_file_title",
                translate=t,
                level="warning",
            )
            self._emit_status_key("help.status.file_missing")
            return

        opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.pdf_path)))
        if not opened:
            log_process_message(
                process_name,
                logger=self._logger,
                error=RuntimeError("Desktop service returned openUrl=False"),
                user_error_message=_build_status_message(
                    "help.status.open_failed",
                    translate=t,
                ),
            )
            _show_toast_key(
                self,
                text_key="help.open_failed_body",
                title_key="help.open_failed_title",
                translate=t,
                level="warning",
            )
            self._emit_status_key("help.status.open_failed")
            return

        log_process_message(
            process_name,
            logger=self._logger,
            success_message=f"{process_name} completed successfully.",
            user_success_message=_build_status_message(
                "help.status.open_success",
                translate=t,
            ),
        )
        _show_toast_key(
            self,
            text_key="help.open_success_body",
            title_key="help.open_success_title",
            translate=t,
            level="success",
        )
        self._emit_status_key("help.status.open_success")

