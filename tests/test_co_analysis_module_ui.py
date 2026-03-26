from __future__ import annotations

from pathlib import Path
from typing import cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from modules import co_analysis_module as co_analysis_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _dispose_widget(widget: object, qapp: QApplication) -> None:
    close = getattr(widget, "close", None)
    if callable(close):
        close()
    delete_later = getattr(widget, "deleteLater", None)
    if callable(delete_later):
        delete_later()
    qapp.processEvents()


def _build_module_with_message_capture(
    monkeypatch: pytest.MonkeyPatch,
) -> tuple[co_analysis_ui.COAnalysisModule, list[str]]:
    seen_keys: list[str] = []
    monkeypatch.setattr(co_analysis_ui, "t", lambda key, **kwargs: key)

    def _capture_notify_message_key(self, text_key: str, **kwargs: object) -> None:
        del self, kwargs
        seen_keys.append(text_key)

    monkeypatch.setattr(
        co_analysis_ui.ModuleRuntime,
        "notify_message_key",
        _capture_notify_message_key,
    )
    monkeypatch.setattr(co_analysis_ui.COAnalysisModule, "_setup_ui_logging", lambda self: None)
    return co_analysis_ui.COAnalysisModule(), seen_keys


def test_drop_widget_syncs_files_and_clear(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module, _seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(["C:/a.xlsx", "C:/b.xlsx"], emit_drop=True)
    assert module._files == [Path("C:/a.xlsx"), Path("C:/b.xlsx")]
    assert module.clear_button.isEnabled() is True
    assert module.calculate_button.isEnabled() is True

    module.drop_widget.clear_button.click()
    qapp.processEvents()
    assert module._files == []
    assert module.clear_button.isEnabled() is False
    assert module.calculate_button.isEnabled() is False
    _dispose_widget(module, qapp)


def test_submit_reports_no_backend_action(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(["C:/a.xlsx"], emit_drop=True)
    module._on_submit_requested()
    assert "coordinator.status.operation_cancelled" in seen_keys
    _dispose_widget(module, qapp)


def test_submit_shows_threshold_violation_when_invalid(
    monkeypatch: pytest.MonkeyPatch,
    qapp: QApplication,
) -> None:
    module, seen_keys = _build_module_with_message_capture(monkeypatch)
    module.drop_widget.add_files(["C:/a.xlsx"], emit_drop=True)
    module.threshold_l1_input.setValue(70.0)
    module.threshold_l2_input.setValue(60.0)
    module.threshold_l3_input.setValue(90.0)
    module._on_submit_requested()
    assert "coordinator.thresholds.invalid_rule" in seen_keys
    _dispose_widget(module, qapp)
