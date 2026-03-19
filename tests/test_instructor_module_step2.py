from __future__ import annotations

from dataclasses import dataclass
from typing import Any, cast

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
    @dataclass
    class _State:
        busy: bool = False

    def __init__(self) -> None:
        self.marks_template_path: str | None = None
        self.marks_template_paths: list[str] = []
        self.step2_course_details_path: str | None = None
        self.step1_course_details_paths: list[str] = []
        self.marks_template_done = False
        self.step2_upload_ready = False
        self.filled_marks_done = False
        self.final_report_done = False
        self.filled_marks_outdated = False
        self.final_report_outdated = False
        self.state: Any = _DummyModule._State()
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
            self.state = _DummyModule._State()
        self.state.busy = busy

    def _publish_status(self, message: str) -> None:
        instructor_ui.emit_user_status(self.status_changed, message, logger=None)

    def _start_async_operation(self, *, token, job_id, work, on_success, on_failure) -> None:
        self._cancel_token = token
        self._set_busy(True, job_id=job_id)

        def _on_finished(result):
            on_success(result)
            self._cancel_token = None
            self._set_busy(False)
            self._refresh_ui()

        def _on_failed(exc):
            on_failure(exc)
            self._cancel_token = None
            self._set_busy(False)
            self._refresh_ui()

        instructor_ui.run_in_background(work, on_finished=_on_finished, on_failed=_on_failed)

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
    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda _parent, body, *, title, level: None,
    )

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
        "getOpenFileNames",
        lambda *_args, **_kwargs: ([], ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: validate_calls.__setitem__("count", validate_calls["count"] + 1),
    )

    instructor_ui.InstructorModule._upload_course_details_async(cast(Any, dummy))

    assert validate_calls["count"] == 0
    assert dummy.step2_upload_ready is False
    assert dummy.marks_template_done is False
    assert dummy.marks_template_path is None
    assert dummy.step2_course_details_path is None
    assert dummy.status_changed.messages == []

def test_step2_upload_validation_failure_shows_error(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileNames",
        lambda *_args, **_kwargs: (["D:/tmp/course_details.xlsx"], ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(ValidationError("bad workbook")),
    )

    instructor_ui.InstructorModule._upload_course_details_async(cast(Any, dummy))

    assert dummy._toasts == []
    assert dummy.step2_upload_ready is False
    assert dummy.marks_template_done is False
    assert "instructor.status.step1_validating_progress" in dummy.status_changed.messages
    assert "instructor.status.step1_validated_progress" in dummy.status_changed.messages

def test_step2_upload_success_enables_prepare(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    remembered: list[tuple[str, str]] = []

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileNames",
        lambda *_args, **_kwargs: (["D:/tmp/course_details.xlsx"], ""),
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

    instructor_ui.InstructorModule._upload_course_details_async(cast(Any, dummy))

    assert dummy.step2_upload_ready is True
    assert dummy.step2_course_details_path == "D:/tmp/course_details.xlsx"
    assert dummy.step1_course_details_paths == ["D:/tmp/course_details.xlsx"]
    assert dummy.marks_template_done is False
    assert dummy.marks_template_path is None
    assert remembered == [("D:/tmp/course_details.xlsx", instructor_ui.APP_NAME)]
    assert "instructor.status.step1_validating_progress" in dummy.status_changed.messages
    assert "instructor.status.step1_validated" in dummy.status_changed.messages
    assert "instructor.status.step1_validated_progress" in dummy.status_changed.messages

def test_step2_upload_async_updates_state_on_success(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    dummy.state = _DummyModule._State()
    dummy._active_jobs = []
    dummy._cancel_token = None
    dummy._workflow_service = None

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileNames",
        lambda *_args, **_kwargs: (["D:/tmp/course_details_async.xlsx"], ""),
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

    instructor_ui.InstructorModule._upload_course_details_async(cast(Any, dummy))

    assert dummy.step2_upload_ready is True
    assert dummy.step2_course_details_path == "D:/tmp/course_details_async.xlsx"
    assert dummy.step1_course_details_paths == ["D:/tmp/course_details_async.xlsx"]
    assert dummy.marks_template_done is False
    assert dummy.marks_template_path is None
    assert dummy.state.busy is False
    assert "instructor.status.step1_validating_progress" in dummy.status_changed.messages
    assert "instructor.status.step1_validated" in dummy.status_changed.messages
    assert "instructor.status.step1_validated_progress" in dummy.status_changed.messages


