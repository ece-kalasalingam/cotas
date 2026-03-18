from __future__ import annotations

from dataclasses import dataclass
from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from common.exceptions import JobCancelledError
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

    def __init__(self, *, filled_marks_path: str | None = None) -> None:
        self.step2_course_details_path: str | None = None
        self.step2_upload_ready = False
        self.filled_marks_path = filled_marks_path
        self.filled_marks_done = bool(filled_marks_path)
        self.state: Any = _DummyModule._State()
        self._active_jobs: list[object] = []
        self._cancel_token = None
        self._workflow_service = None
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
        self.state.busy = busy

    def _refresh_ui(self) -> None:
        return


def _run_sync(fn, *args, on_finished=None, on_failed=None, **kwargs):
    try:
        result = fn(*args, **kwargs)
    except Exception as exc:  # pragma: no cover
        if on_failed is not None:
            on_failed(exc)
    else:
        if on_finished is not None:
            on_finished(result)
    return object()


def _patch_common(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    monkeypatch.setattr(instructor_ui, "run_in_background", _run_sync)
    monkeypatch.setattr(
        instructor_ui,
        "resolve_dialog_start_path",
        lambda _app, _default=None: "D:/tmp/default.xlsx",
    )
    monkeypatch.setattr(instructor_ui, "remember_dialog_dir", lambda *_args, **_kwargs: None)


def test_step2_upload_async_cancelled_reports_status(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common(monkeypatch)
    dummy = _DummyModule()
    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "validate_course_details_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(JobCancelledError("cancelled")),
    )

    instructor_ui.InstructorModule._upload_course_details_async(cast(Any, dummy))

    assert dummy.step2_upload_ready is False
    assert dummy.step2_course_details_path is None
    assert dummy.status_changed.messages == ["instructor.status.operation_cancelled"]
    assert dummy._toasts == []


def test_step3_upload_async_cancelled_reports_status(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common(monkeypatch)
    dummy = _DummyModule()
    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/filled_marks.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "_validate_uploaded_filled_marks_workbook",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(JobCancelledError("cancelled")),
    )

    instructor_ui.InstructorModule._upload_filled_marks_async(cast(Any, dummy))

    assert dummy.filled_marks_done is False
    assert dummy.filled_marks_path is None
    assert dummy.status_changed.messages == ["instructor.status.operation_cancelled"]
    assert dummy._toasts == []
