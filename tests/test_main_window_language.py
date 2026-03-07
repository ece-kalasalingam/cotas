from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

import main_window as main_window_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class _RetranslateModule:
    def __init__(self) -> None:
        self.calls = 0

    def retranslate_ui(self) -> None:
        self.calls += 1


class _RefreshModule:
    def __init__(self) -> None:
        self.calls = 0

    def _refresh_ui(self) -> None:
        self.calls += 1


def test_apply_language_change_updates_ui_and_loaded_modules(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    state = {"lang": "en"}

    def fake_t(key: str, **_kwargs: object) -> str:
        return f"{state['lang']}:{key}"

    monkeypatch.setattr(main_window_ui, "t", fake_t)
    monkeypatch.setattr(main_window_ui.MainWindow, "load_module", lambda *_args, **_kwargs: None)

    window = main_window_ui.MainWindow()
    try:
        retranslate_module = _RetranslateModule()
        refresh_module = _RefreshModule()
        window.modules["r"] = retranslate_module
        window.modules["f"] = refresh_module

        state["lang"] = "ta"
        window.apply_language_change()

        assert window.windowTitle().startswith("ta:")
        assert window.action_co_section.text() == "ta:module.instructor"
        assert window.action_help.text() == "ta:nav.help"
        assert window.language_status_button.text().startswith("ta:language.switcher.button")
        assert retranslate_module.calls == 1
        assert refresh_module.calls == 1
    finally:
        window.close()


def test_language_menu_action_is_ignored_when_switch_is_disabled(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    monkeypatch.setattr(main_window_ui.MainWindow, "load_module", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(main_window_ui, "set_ui_language_preference", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(main_window_ui, "t", lambda key, **_kwargs: key)

    called = {"count": 0}
    window = main_window_ui.MainWindow(on_language_applied=lambda _code: called.__setitem__("count", 1))
    try:
        window.set_language_switch_enabled(False)
        action = window.language_menu.addAction("Tamil")
        action.setData("ta-in")

        window._on_language_menu_action(action)

        assert called["count"] == 0
    finally:
        window.close()


def test_language_menu_has_single_checked_action(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    monkeypatch.setattr(main_window_ui.MainWindow, "load_module", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(main_window_ui, "t", lambda key, **_kwargs: key)

    window = main_window_ui.MainWindow()
    try:
        window._rebuild_language_menu("auto")
        checked_auto = [
            action
            for action in window.language_menu.actions()
            if action.isCheckable() and action.isChecked()
        ]
        assert len(checked_auto) == 1
        assert checked_auto[0].data() == "auto"

        window._rebuild_language_menu("en")
        checked_en = [
            action
            for action in window.language_menu.actions()
            if action.isCheckable() and action.isChecked()
        ]
        assert len(checked_en) == 1
        assert checked_en[0].data() == "en"
    finally:
        window.close()
