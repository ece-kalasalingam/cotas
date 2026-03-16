from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("PySide6")

from PySide6.QtPdf import QPdfDocument
from PySide6.QtWidgets import QApplication

from modules import help_module as help_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _build_widget(monkeypatch: pytest.MonkeyPatch) -> help_ui.HelpModule:
    monkeypatch.setattr(help_ui.HelpModule, "_load_pdf", lambda self: None)
    monkeypatch.setattr(help_ui, "t", lambda key, **kwargs: key)
    return help_ui.HelpModule()


def test_load_pdf_missing_file_emits_warning_status(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, qapp: QApplication
) -> None:
    toasts: list[tuple[str, str, str]] = []
    statuses: list[str] = []
    monkeypatch.setattr(help_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(
        help_ui,
        "show_toast",
        lambda _parent, message, *, title, level: toasts.append((message, title, level)),
    )
    monkeypatch.setattr(help_ui, "emit_user_status", lambda _sig, msg, logger=None: statuses.append(msg))
    monkeypatch.setattr(help_ui, "resource_path", lambda _p: str(tmp_path / "definitely_missing_help.pdf"))

    widget = help_ui.HelpModule()

    assert toasts and toasts[-1][2] == "warning"
    assert statuses, "missing-file path should emit user status"
    widget.close()


def test_pdf_status_error_is_shown_only_once_until_ready(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    widget = _build_widget(monkeypatch)
    toasts: list[str] = []
    monkeypatch.setattr(help_ui, "show_toast", lambda *_args, **_kwargs: toasts.append("toast"))
    monkeypatch.setattr(help_ui, "emit_user_status", lambda *_args, **_kwargs: None)

    widget._on_pdf_status_changed(QPdfDocument.Status.Error)
    widget._on_pdf_status_changed(QPdfDocument.Status.Error)
    assert len(toasts) == 1

    widget._on_pdf_status_changed(QPdfDocument.Status.Ready)
    widget._on_pdf_status_changed(QPdfDocument.Status.Error)
    assert len(toasts) == 2
    widget.close()


def test_download_pdf_missing_source_shows_warning(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    widget = _build_widget(monkeypatch)
    widget.pdf_path = Path("Z:/missing.pdf")
    toasts: list[tuple[str, str, str]] = []
    monkeypatch.setattr(
        help_ui,
        "show_toast",
        lambda _parent, message, *, title, level: toasts.append((message, title, level)),
    )
    monkeypatch.setattr(help_ui, "emit_user_status", lambda *_args, **_kwargs: None)

    widget.download_pdf()

    assert toasts and toasts[-1][2] == "warning"
    widget.close()


def test_open_external_paths(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, qapp: QApplication) -> None:
    widget = _build_widget(monkeypatch)
    toasts: list[str] = []
    monkeypatch.setattr(help_ui, "show_toast", lambda *_args, **kwargs: toasts.append(kwargs.get("level", "")))
    monkeypatch.setattr(help_ui, "log_process_message", lambda *_args, **_kwargs: None)

    widget.pdf_path = Path("Z:/missing.pdf")
    monkeypatch.setattr(help_ui, "emit_user_status", lambda *_args, **_kwargs: None)
    widget.open_external()
    assert toasts[-1] == "warning"

    existing = tmp_path / "help.pdf"
    existing.write_bytes(b"pdf")
    widget.pdf_path = existing
    monkeypatch.setattr(help_ui.QDesktopServices, "openUrl", lambda _url: False)
    widget.open_external()
    assert toasts[-1] == "warning"

    monkeypatch.setattr(help_ui.QDesktopServices, "openUrl", lambda _url: True)
    widget.open_external()
    assert toasts[-1] == "success"
    widget.close()


def test_load_pdf_success_logs_and_emits_status(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, qapp: QApplication
) -> None:
    pdf = tmp_path / "help.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")

    seen = {"log": 0, "statuses": []}
    monkeypatch.setattr(help_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(help_ui, "resource_path", lambda _p: str(pdf))
    monkeypatch.setattr(help_ui, "log_process_message", lambda *_args, **_kwargs: seen.__setitem__("log", seen["log"] + 1))
    monkeypatch.setattr(help_ui, "emit_user_status", lambda _sig, msg, logger=None: seen["statuses"].append(msg))

    widget = help_ui.HelpModule()

    assert seen["log"] == 1
    assert seen["statuses"]
    widget.close()


def test_show_context_menu_dispatches_actions(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    widget = _build_widget(monkeypatch)
    called = {"download": 0, "open": 0}

    monkeypatch.setattr(widget, "download_pdf", lambda: called.__setitem__("download", called["download"] + 1))
    monkeypatch.setattr(widget, "open_external", lambda: called.__setitem__("open", called["open"] + 1))

    class _FakeMenu:
        selected = "download"

        def __init__(self) -> None:
            self._actions = {}

        def setStyleSheet(self, _value: str) -> None:  # noqa: N802
            return None

        def setStyle(self, _style) -> None:  # noqa: N802
            return None

        def addAction(self, label: str):  # noqa: N802
            token = object()
            if "download" in label:
                self._actions["download"] = token
            elif "open_default_viewer" in label:
                self._actions["open"] = token
            return token

        def exec(self, _pos):
            if self.selected == "download":
                return self._actions["download"]
            if self.selected == "open":
                return self._actions["open"]
            return object()

    monkeypatch.setattr(help_ui, "QMenu", _FakeMenu)
    monkeypatch.setattr(help_ui.QStyleFactory, "create", lambda _name: None)
    monkeypatch.setattr(help_ui.QApplication, "style", lambda: object())
    monkeypatch.setattr(widget.pdf_view, "mapToGlobal", lambda pos: pos)

    _FakeMenu.selected = "download"
    widget.show_context_menu(None)
    assert called == {"download": 1, "open": 0}

    _FakeMenu.selected = "open"
    widget.show_context_menu(None)
    assert called == {"download": 1, "open": 1}

    _FakeMenu.selected = "none"
    widget.show_context_menu(None)
    assert called == {"download": 1, "open": 1}
    widget.close()
