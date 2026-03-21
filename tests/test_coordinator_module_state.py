from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from modules import coordinator_module as coordinator_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _build_module(monkeypatch: pytest.MonkeyPatch) -> coordinator_ui.CoordinatorModule:
    monkeypatch.setattr(coordinator_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(coordinator_ui.CoordinatorModule, "_setup_ui_logging", lambda self: None)
    return coordinator_ui.CoordinatorModule()


def _dispose_widget(widget: object, qapp: QApplication) -> None:
    close = getattr(widget, "close", None)
    if callable(close):
        close()
    delete_later = getattr(widget, "deleteLater", None)
    if callable(delete_later):
        delete_later()
    qapp.processEvents()


def test_save_shortcut_runs_only_when_not_busy_and_enabled(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"count": 0}
    monkeypatch.setattr(module, "_on_calculate_clicked", lambda: calls.__setitem__("count", calls["count"] + 1))

    module.state.busy = True
    module._on_save_shortcut_activated()
    assert calls["count"] == 0

    module.state.busy = False
    module.calculate_button.setEnabled(False)
    module._on_save_shortcut_activated()
    assert calls["count"] == 0

    module.calculate_button.setEnabled(True)
    module._on_save_shortcut_activated()
    assert calls["count"] == 1
    _dispose_widget(module, qapp)


def test_drain_next_batch_pops_only_when_idle(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    got: list[list[str]] = []
    monkeypatch.setattr(module, "_process_files_async", lambda dropped: got.append(dropped))

    module.state.busy = True
    module._pending_drop_batches = [["a.xlsx"], ["b.xlsx"]]
    module._drain_next_batch()
    assert got == []
    assert module._pending_drop_batches == [["a.xlsx"], ["b.xlsx"]]

    module.state.busy = False
    module._drain_next_batch()
    assert got == [["a.xlsx"]]
    assert module._pending_drop_batches == [["b.xlsx"]]
    _dispose_widget(module, qapp)


def test_browse_files_respects_busy_and_processes_selection(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)

    module.state.busy = True
    called = {"dialog": 0, "remember": 0, "process": 0}
    monkeypatch.setattr(
        coordinator_ui.QFileDialog,
        "getOpenFileNames",
        lambda *_args, **_kwargs: called.__setitem__("dialog", called["dialog"] + 1) or (["x.xlsx"], ""),
    )
    monkeypatch.setattr(module, "_remember_dialog_dir_safe", lambda *_args, **_kwargs: called.__setitem__("remember", called["remember"] + 1))
    monkeypatch.setattr(module, "_process_files_async", lambda *_args, **_kwargs: called.__setitem__("process", called["process"] + 1))

    module._browse_files()
    assert called == {"dialog": 0, "remember": 0, "process": 0}

    module.state.busy = False
    module._browse_files()
    assert called == {"dialog": 1, "remember": 1, "process": 1}
    _dispose_widget(module, qapp)


def test_remember_dialog_dir_safe_fallback(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"fallback": 0}

    class _Runtime:
        def remember_dialog_dir_safe(self, *_args, **_kwargs):
            calls["fallback"] += 1

    module._runtime = cast(Any, _Runtime())
    module._remember_dialog_dir_safe("C:/tmp/a.xlsx")
    assert calls == {"fallback": 1}
    _dispose_widget(module, qapp)


def test_close_event_cancels_token_and_removes_handler(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)

    class _Token:
        def __init__(self) -> None:
            self.cancelled = False

        def cancel(self) -> None:
            self.cancelled = True

    token = _Token()
    module._cancel_token = cast(Any, token)
    module._active_jobs = [object()]
    module._ui_log_handler = cast(Any, object())
    removed = {"count": 0}
    monkeypatch.setattr(module._logger, "removeHandler", lambda _h: removed.__setitem__("count", removed["count"] + 1))

    _dispose_widget(module, qapp)

    assert token.cancelled is True
    assert module._cancel_token is None
    assert module._active_jobs == []
    assert removed["count"] == 1
    assert module._ui_log_handler is None


def test_refresh_ui_toggles_controls_with_file_state(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    module._files = []
    module.state.busy = False
    module._refresh_ui()
    assert module.clear_button.isEnabled() is False
    assert module.calculate_button.isEnabled() is False

    module._files = [Path("C:/a.xlsx")]
    module.state.busy = False
    module._refresh_ui()
    assert module.clear_button.isEnabled() is True
    assert module.calculate_button.isEnabled() is True

    module.state.busy = True
    module._refresh_ui()
    assert module.clear_button.isEnabled() is False
    assert module.calculate_button.isEnabled() is False
    assert module.drop_list.isEnabled() is False
    _dispose_widget(module, qapp)


def test_threshold_changes_revalidate_calculate_button(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    module._files = [Path("C:/a.xlsx")]
    module.state.busy = False

    module.threshold_l1_input.setValue(40.0)
    module.threshold_l2_input.setValue(60.0)
    module.threshold_l3_input.setValue(80.0)
    module._refresh_ui()
    assert module.calculate_button.isEnabled() is True

    module.threshold_l2_input.setValue(40.0)
    assert module.calculate_button.isEnabled() is False

    module.threshold_l2_input.setValue(70.0)
    assert module.calculate_button.isEnabled() is True
    _dispose_widget(module, qapp)


def test_threshold_violation_emits_toast_and_activity_log(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    toasts: list[tuple[str, str, str]] = []
    status_keys: list[str] = []

    monkeypatch.setattr(
        coordinator_ui,
        "show_toast",
        lambda _parent, message, *, title, level: toasts.append((message, title, level)),
    )
    monkeypatch.setattr(module, "_publish_status_key", lambda key, **_kwargs: status_keys.append(key))

    module.threshold_l1_input.setValue(40.0)
    module.threshold_l2_input.setValue(60.0)
    module.threshold_l3_input.setValue(80.0)

    module.threshold_l2_input.setValue(40.0)
    module.threshold_l2_input.editingFinished.emit()
    assert toasts[-1] == (
        coordinator_ui.CoordinatorModule._THRESHOLD_VALIDATION_KEY,
        "coordinator.title",
        "error",
    )
    assert status_keys[-1] == coordinator_ui.CoordinatorModule._THRESHOLD_VALIDATION_KEY
    initial_count = len(toasts)

    module.threshold_l2_input.setValue(30.0)
    module.threshold_l2_input.editingFinished.emit()
    assert len(toasts) == initial_count

    module.threshold_l2_input.setValue(70.0)
    module.threshold_l2_input.setValue(80.0)
    module.threshold_l2_input.editingFinished.emit()
    assert len(toasts) == initial_count + 1
    assert status_keys[-1] == coordinator_ui.CoordinatorModule._THRESHOLD_VALIDATION_KEY
    _dispose_widget(module, qapp)
