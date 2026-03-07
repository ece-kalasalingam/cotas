from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

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
