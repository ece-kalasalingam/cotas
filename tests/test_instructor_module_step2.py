from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from common.exceptions import ValidationError
from modules import instructor_module as instructor_ui

class _SignalRecorder:
    def __init__(self) -> None:
        self.messages: list[str] = []

    def emit(self, message: str) -> None:
        self.messages.append(message)

class _DummyModule:
    def __init__(self) -> None:
        self.step2_path: str | None = None
        self.step2_course_details_path: str | None = None
        self.step2_done = False
        self.step2_upload_ready = False
        self.step3_done = False
        self.step4_done = False
        self.step3_outdated = False
        self.step4_outdated = False
        self.state = type("State", (), {"busy": False})()
        self._active_jobs: list[object] = []
        self._cancel_token = None
        self._workflow_service = None
        self._step2_marks_default_name = "marks_template.xlsx"
        self.status_changed = _SignalRecorder()
        self._toasts: list[tuple[str, str]] = []

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        instructor_ui.remember_dialog_dir(selected_path, app_name=instructor_ui.APP_NAME)

    def _show_validation_error_toast(self, message: str) -> None:
        self._toasts.append(("validation", message))

    def _show_system_error_toast(self, step: int) -> None:
        self._toasts.append(("system", str(step)))

    def _show_step_success_toast(self, step: int) -> None:
        self._toasts.append(("success", str(step)))

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:  # noqa: ARG002
        if not hasattr(self, "state"):
            self.state = type("State", (), {"busy": False})()
        self.state.busy = busy

    def _refresh_ui(self) -> None:
        return

def _patch_common_dependencies(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    monkeypatch.setattr(
        instructor_ui,
        "resolve_dialog_start_path",
        lambda _app, _default=None: "D:/tmp/course_details.xlsx",
    )
    monkeypatch.setattr(instructor_ui, "run_in_background", _run_sync)
    monkeypatch.setattr(instructor_ui, "remember_dialog_dir", lambda *_args, **_kwargs: None)

def _run_sync(fn, *args, on_finished=None, on_failed=None, **kwargs):
    try:
        result = fn(*args, **kwargs)
    except Exception as exc:  # pragma: no cover - helper supports both paths
        if on_failed is not None:
            on_failed(exc)
    else:
        if on_finished is not None:
            on_finished(result)
    return object()

def test_step2_upload_cancel_keeps_state(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    validate_calls = {"count": 0}

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: validate_calls.__setitem__("count", validate_calls["count"] + 1),
    )

    instructor_ui.InstructorModule._upload_course_details(dummy)

    assert validate_calls["count"] == 0
    assert dummy.step2_upload_ready is False
    assert dummy.step2_done is False
    assert dummy.step2_path is None
    assert dummy.step2_course_details_path is None
    assert dummy.status_changed.messages == []

def test_step2_upload_validation_failure_shows_error(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(ValidationError("bad workbook")),
    )

    instructor_ui.InstructorModule._upload_course_details(dummy)

    assert dummy._toasts == [("validation", "bad workbook")]
    assert dummy.step2_upload_ready is False
    assert dummy.step2_done is False
    assert dummy.status_changed.messages == []

def test_step2_upload_success_enables_prepare(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    remembered: list[tuple[str, str]] = []

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: instructor_ui.ID_COURSE_SETUP,
    )
    monkeypatch.setattr(
        instructor_ui,
        "remember_dialog_dir",
        lambda path, app_name: remembered.append((path, app_name)),
    )

    instructor_ui.InstructorModule._upload_course_details(dummy)

    assert dummy.step2_upload_ready is True
    assert dummy.step2_course_details_path == "D:/tmp/course_details.xlsx"
    assert dummy.step2_done is False
    assert dummy.step2_path is None
    assert remembered == [("D:/tmp/course_details.xlsx", instructor_ui.APP_NAME)]
    assert dummy.status_changed.messages == ["instructor.status.step2_validated"]

def test_step2_upload_async_updates_state_on_success(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    dummy.state = type("State", (), {"busy": False})()
    dummy._active_jobs = []
    dummy._cancel_token = None
    dummy._workflow_service = None

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details_async.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: instructor_ui.ID_COURSE_SETUP,
    )

    def _run_sync(fn, *args, on_finished=None, on_failed=None, **kwargs):
        try:
            result = fn(*args, **kwargs)
        except Exception as exc:  # pragma: no cover - success path
            if on_failed is not None:
                on_failed(exc)
        else:
            if on_finished is not None:
                on_finished(result)
        return object()

    monkeypatch.setattr(instructor_ui, "run_in_background", _run_sync)

    instructor_ui.InstructorModule._upload_course_details_async(dummy)

    assert dummy.step2_upload_ready is True
    assert dummy.step2_course_details_path == "D:/tmp/course_details_async.xlsx"
    assert dummy.step2_done is False
    assert dummy.step2_path is None
    assert dummy.state.busy is False
    assert dummy.status_changed.messages == ["instructor.status.step2_validated"]


