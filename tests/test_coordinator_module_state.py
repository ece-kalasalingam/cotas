from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from modules import coordinator_module as coordinator_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _build_module(monkeypatch: pytest.MonkeyPatch) -> coordinator_ui.CoordinatorModule:
    monkeypatch.setattr(coordinator_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(coordinator_ui.CoordinatorModule, "_setup_ui_logging", lambda self: None)
    return coordinator_ui.CoordinatorModule()


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
    module.close()


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
    module.close()


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
    module.close()


def test_remember_dialog_dir_safe_fallback(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"primary": 0, "fallback": 0}

    def _primary(*_args, **_kwargs):
        calls["primary"] += 1
        raise OSError("nope")

    def _fallback(*_args, **_kwargs):
        calls["fallback"] += 1

    monkeypatch.setattr(coordinator_ui, "remember_dialog_dir", _primary)
    monkeypatch.setattr(coordinator_ui, "remember_dialog_dir_safe", _fallback)

    module._remember_dialog_dir_safe("C:/tmp/a.xlsx")
    assert calls == {"primary": 1, "fallback": 1}
    module.close()


def test_close_event_cancels_token_and_removes_handler(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)

    class _Token:
        def __init__(self) -> None:
            self.cancelled = False

        def cancel(self) -> None:
            self.cancelled = True

    token = _Token()
    module._cancel_token = token
    module._active_jobs = [object()]
    module._ui_log_handler = object()
    removed = {"count": 0}
    monkeypatch.setattr(module._logger, "removeHandler", lambda _h: removed.__setitem__("count", removed["count"] + 1))

    module.close()

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
    module.close()
