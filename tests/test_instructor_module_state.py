from __future__ import annotations

from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from modules import instructor_module as instructor_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _build_module(monkeypatch: pytest.MonkeyPatch) -> instructor_ui.InstructorModule:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(instructor_ui.InstructorModule, "_setup_ui_logging", lambda self: None)
    return instructor_ui.InstructorModule()


def test_on_open_shortcut_routes_by_step_and_enabled(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"step1": 0, "step2": 0}
    monkeypatch.setattr(module, "_on_step1_upload_clicked", lambda: calls.__setitem__("step1", calls["step1"] + 1))
    monkeypatch.setattr(module, "_on_step2_upload_clicked", lambda: calls.__setitem__("step2", calls["step2"] + 1))

    module.state.busy = True
    module._on_open_shortcut_activated()
    assert calls == {"step1": 0, "step2": 0}

    module.state.busy = False
    module.current_step = 1
    module.step1_upload_action.setEnabled(True)
    module._on_open_shortcut_activated()
    assert calls["step1"] == 1

    module.current_step = 2
    module.step2_upload_action.setEnabled(True)
    module._on_open_shortcut_activated()
    assert calls["step2"] == 1
    module.close()


def test_on_save_shortcut_routes_by_step_and_enabled(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"prepare": 0, "generate": 0}
    monkeypatch.setattr(module, "_on_step1_prepare_clicked", lambda: calls.__setitem__("prepare", calls["prepare"] + 1))
    monkeypatch.setattr(module, "_on_step2_generate_clicked", lambda: calls.__setitem__("generate", calls["generate"] + 1))

    module.state.busy = True
    module._on_save_shortcut_activated()
    assert calls == {"prepare": 0, "generate": 0}

    module.state.busy = False
    module.current_step = 1
    module.step1_prepare_action.setEnabled(True)
    module._on_save_shortcut_activated()
    assert calls["prepare"] == 1

    module.current_step = 2
    module.step2_generate_action.setEnabled(True)
    module._on_save_shortcut_activated()
    assert calls["generate"] == 1
    module.close()


def test_on_step_row_changed_ignores_negative_and_maps_positive(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    selected: list[int] = []
    monkeypatch.setattr(module, "_on_step_selected", lambda step: selected.append(step))

    module._on_step_row_changed(-1)
    module._on_step_row_changed(0)
    module._on_step_row_changed(1)
    module._on_step_row_changed(2)

    assert selected == [1, 2]
    module.close()


def test_run_current_step_action_dispatch(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"u2": 0, "g3": 0, "r": 0}
    monkeypatch.setattr(module, "_upload_course_details_async", lambda: calls.__setitem__("u2", calls["u2"] + 1))
    monkeypatch.setattr(module, "_generate_final_report_async", lambda: calls.__setitem__("g3", calls["g3"] + 1))
    monkeypatch.setattr(module, "_refresh_ui", lambda: calls.__setitem__("r", calls["r"] + 1))

    module.current_step = 1
    module._run_current_step_action()
    module.current_step = 2
    module._run_current_step_action()
    module.current_step = 99
    module._run_current_step_action()

    assert calls == {"u2": 1, "g3": 1, "r": 3}
    module.close()


def test_remember_dialog_dir_safe_fallback(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    calls = {"fallback": 0}

    def _fallback(*_args, **_kwargs):
        calls["fallback"] += 1

    monkeypatch.setattr(module._runtime, "remember_dialog_dir_safe", _fallback)

    module._remember_dialog_dir_safe("C:/tmp/a.xlsx")
    assert calls == {"fallback": 1}
    module.close()


def test_set_busy_toggles_host_language_switch(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
    module = _build_module(monkeypatch)
    toggles: list[bool] = []

    class _Host:
        def set_language_switch_enabled(self, enabled: bool) -> None:
            toggles.append(enabled)

    monkeypatch.setattr(module, "window", lambda: _Host())
    monkeypatch.setattr(module, "_refresh_ui", lambda: None)

    module._set_busy(True, job_id="j1")
    assert module.state.busy is True

    module._set_busy(False)
    assert module.state.busy is False
    assert toggles == [False, True]
    module.close()


def test_close_event_cleans_resources(monkeypatch: pytest.MonkeyPatch, qapp: QApplication) -> None:
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
    monkeypatch.setattr(instructor_ui._logger, "removeHandler", lambda _h: removed.__setitem__("count", removed["count"] + 1))

    module.close()

    assert module._is_closing is True
    assert token.cancelled is True
    assert module._cancel_token is None
    assert module._active_jobs == []
    assert removed["count"] == 1
    assert module._ui_log_handler is None


def test_step1_drop_signal_handlers_publish_log_and_update_count(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    monkeypatch.setattr(
        instructor_ui,
        "t",
        lambda key, **kwargs: f"{key}:{kwargs.get('count', '')}",
    )
    monkeypatch.setattr(instructor_ui.InstructorModule, "_setup_ui_logging", lambda self: None)
    module = instructor_ui.InstructorModule()

    published: list[str] = []
    uploads: list[list[str]] = []
    monkeypatch.setattr(module, "_publish_status", lambda message: published.append(message))
    monkeypatch.setattr(module, "_upload_course_details_from_paths_async", lambda paths: uploads.append(list(paths)))
    monkeypatch.setattr(module, "_refresh_ui", lambda: None)

    module._on_step1_drop_browse_requested()
    module.step1_drop_widget.add_files(["C:/course_details.xlsx"])
    module.step1_drop_widget.add_files(["https://example.com/invalid.xlsx"])

    assert uploads == [["C:/course_details.xlsx"]]
    assert module.step1_drop_widget.summary_label.text() == "instructor.step1.drop.summary:1"
    assert "instructor.status.step1_drop_browse_requested:" in published
    assert "instructor.status.step1_drop_files_dropped:1" in published
    assert "instructor.status.step1_drop_files_changed:1" in published
    assert "instructor.status.step1_drop_files_rejected:1" in published
    module.close()

