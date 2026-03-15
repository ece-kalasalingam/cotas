from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QWidget
from PySide6.QtWidgets import QApplication

import main_window as main_window_ui
from modules import about_module as about_ui
from modules import help_module as help_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_about_module_constructs(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    monkeypatch.setattr(about_ui, "t", lambda key, **_kwargs: key)
    widget = about_ui.AboutModule()
    assert widget.layout() is not None
    assert widget.layout().count() > 0


def test_help_module_initializes(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    calls = {"load": 0}
    monkeypatch.setattr(
        help_ui.HelpModule,
        "_load_pdf",
        lambda self: calls.__setitem__("load", calls["load"] + 1),
    )

    widget = help_ui.HelpModule()
    assert calls["load"] == 1
    assert widget.pdf_view is not None


def test_main_window_reuses_module_instance_across_repeated_switches(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    original_load_module = main_window_ui.MainWindow.load_module
    monkeypatch.setattr(main_window_ui.MainWindow, "load_module", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(main_window_ui, "t", lambda key, **_kwargs: key)

    mode_calls: list[bool] = []

    class ReusedModule(QWidget):
        instances = 0

        def __init__(self) -> None:
            super().__init__()
            ReusedModule.instances += 1

        def set_shared_activity_log_mode(self, enabled: bool) -> None:
            mode_calls.append(enabled)

        def get_shared_outputs_html(self) -> str:
            return "<p>outputs</p>"

    class HelpModule(QWidget):
        def __init__(self) -> None:
            super().__init__()

        def set_shared_activity_log_mode(self, enabled: bool) -> None:
            mode_calls.append(enabled)

    window = main_window_ui.MainWindow()
    try:
        for _ in range(25):
            original_load_module(window, ReusedModule)
        assert ReusedModule.instances == 1
        assert window.stack.count() == 1
        assert window.stack.currentWidget() is window.modules["ReusedModule"]
        assert all(mode_calls)

        original_load_module(window, HelpModule)
        assert window.shared_activity_frame.isHidden() is True
        assert mode_calls[-1] is False

        original_load_module(window, ReusedModule)
        assert window.shared_activity_frame.isHidden() is False
        assert mode_calls[-1] is True
    finally:
        window.close()


def test_main_window_signal_lifecycle_does_not_duplicate_connections(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    original_load_module = main_window_ui.MainWindow.load_module
    monkeypatch.setattr(main_window_ui.MainWindow, "load_module", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(main_window_ui, "t", lambda key, **_kwargs: key)

    class SignalModule(QWidget):
        status_changed = Signal(str)

        def __init__(self) -> None:
            super().__init__()

    window = main_window_ui.MainWindow()
    captured: list[str] = []
    try:
        monkeypatch.setattr(window, "_on_module_status_changed", lambda message: captured.append(message))
        for _ in range(30):
            original_load_module(window, SignalModule)
        window.modules["SignalModule"].status_changed.emit("ready")
        qapp.processEvents()
        assert captured == ["ready"]
    finally:
        window.close()


def test_main_window_module_init_failure_does_not_mutate_stack(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    original_load_module = main_window_ui.MainWindow.load_module
    monkeypatch.setattr(main_window_ui.MainWindow, "load_module", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(main_window_ui, "t", lambda key, **_kwargs: key)

    toasts: list[str] = []
    flashes: list[str] = []
    monkeypatch.setattr(main_window_ui, "show_toast", lambda *_args, **kwargs: toasts.append(str(kwargs.get("level"))))

    class BrokenModule(QWidget):
        def __init__(self) -> None:
            raise RuntimeError("boom")

    window = main_window_ui.MainWindow()
    try:
        monkeypatch.setattr(window, "flash_status", lambda message, timeout=0: flashes.append(message))
        original_load_module(window, BrokenModule)
        assert "BrokenModule" not in window.modules
        assert window.stack.count() == 0
        assert toasts == ["error"]
        assert flashes == ["module.load_failed_status"]
    finally:
        window.close()
