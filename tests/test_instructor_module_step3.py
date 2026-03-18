from __future__ import annotations

from dataclasses import dataclass
from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

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
        self.marks_template_done = False
        self.filled_marks_path: str | None = None
        self.filled_marks_done = False
        self.filled_marks_outdated = True
        self.final_report_outdated = False
        self.state: Any = _DummyModule._State()
        self._active_jobs: list[object] = []
        self._cancel_token = None
        self._workflow_service = None
        self.status_changed = _SignalRecorder()
        self._toasts: list[tuple[str, str]] = []

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        instructor_ui.remember_dialog_dir(selected_path, app_name=instructor_ui.APP_NAME)

    def _show_step_success_toast(self, step: int) -> None:
        self._toasts.append(("success", str(step)))

    def _show_validation_error_toast(self, message: str) -> None:
        self._toasts.append(("validation", message))

    def _show_system_error_toast(self, step: int) -> None:
        self._toasts.append(("system", str(step)))

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:  # noqa: ARG002
        if not hasattr(self, "state"):
            self.state = _DummyModule._State()
        self.state.busy = busy

    def _refresh_ui(self) -> None:
        return


def _patch_common_dependencies(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    monkeypatch.setattr(
        instructor_ui,
        "resolve_dialog_start_path",
        lambda _app, _default=None: "D:/tmp/filled_marks.xlsx",
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


def test_step3_upload_cancel_keeps_state(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "_validate_uploaded_filled_marks_workbook",
        lambda *_args, **_kwargs: None,
    )

    instructor_ui.InstructorModule._upload_filled_marks(cast(Any, dummy))

    assert dummy.filled_marks_done is False
    assert dummy.filled_marks_path is None
    assert dummy.status_changed.messages == []


def test_step3_upload_success_allowed_without_step2(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    remembered: list[tuple[str, str]] = []

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/filled_marks.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "remember_dialog_dir",
        lambda path, app_name: remembered.append((path, app_name)),
    )
    monkeypatch.setattr(
        instructor_ui,
        "_validate_uploaded_filled_marks_workbook",
        lambda *_args, **_kwargs: None,
    )

    instructor_ui.InstructorModule._upload_filled_marks(cast(Any, dummy))

    assert dummy.marks_template_done is False
    assert dummy.filled_marks_done is True
    assert dummy.filled_marks_outdated is False
    assert dummy.filled_marks_path == "D:/tmp/filled_marks.xlsx"
    assert dummy.final_report_outdated is False
    assert remembered == [("D:/tmp/filled_marks.xlsx", instructor_ui.APP_NAME)]
    assert dummy.status_changed.messages == ["instructor.status.step2_uploaded_filled"]


def test_step3_replace_marks_flags_step4_outdated(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    dummy.filled_marks_done = True
    dummy.filled_marks_path = "D:/tmp/old_filled_marks.xlsx"

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/new_filled_marks.xlsx", ""),
    )
    monkeypatch.setattr(instructor_ui, "remember_dialog_dir", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        instructor_ui,
        "_validate_uploaded_filled_marks_workbook",
        lambda *_args, **_kwargs: None,
    )

    instructor_ui.InstructorModule._upload_filled_marks(cast(Any, dummy))

    assert dummy.filled_marks_done is True
    assert dummy.filled_marks_outdated is False
    assert dummy.filled_marks_path == "D:/tmp/new_filled_marks.xlsx"
    assert dummy.final_report_outdated is True
    assert dummy.status_changed.messages == ["instructor.status.step2_changed_filled"]


def test_step3_upload_async_updates_state_on_success(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    dummy.state = _DummyModule._State()
    dummy._active_jobs = []
    dummy._cancel_token = None
    dummy._workflow_service = None

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/filled_marks_async.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "_validate_uploaded_filled_marks_workbook",
        lambda *_args, **_kwargs: None,
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

    instructor_ui.InstructorModule._upload_filled_marks_async(cast(Any, dummy))

    assert dummy.filled_marks_done is True
    assert dummy.filled_marks_path == "D:/tmp/filled_marks_async.xlsx"
    assert dummy.filled_marks_outdated is False
    assert dummy.state.busy is False
    assert dummy.status_changed.messages == ["instructor.status.step2_uploaded_filled"]
