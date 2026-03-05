from __future__ import annotations

from pathlib import Path

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
    def __init__(self, *, step3_path: str | None = None) -> None:
        self.step1_path: str | None = None
        self.step1_done = False
        self.step2_path: str | None = None
        self.step2_course_details_path: str | None = None
        self.step2_done = False
        self.step2_upload_ready = False
        self.step3_path = step3_path
        self.step3_done = bool(step3_path)
        self.step3_outdated = False
        self.step4_path: str | None = None
        self.step4_done = False
        self.step4_outdated = False
        self._step2_marks_default_name = "marks_template.xlsx"
        self.state = type("State", (), {"busy": False})()
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

    def _can_run_step(self, _step: int) -> tuple[bool, str]:
        return True, ""

    def _set_busy(self, busy: bool, *, job_id: str | None = None) -> None:  # noqa: ARG002
        self.state.busy = busy

    def _refresh_ui(self) -> None:
        return


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


def _patch_common(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    monkeypatch.setattr(instructor_ui, "run_in_background", _run_sync)
    monkeypatch.setattr(
        instructor_ui,
        "resolve_dialog_start_path",
        lambda _app, _default=None: "D:/tmp/default.xlsx",
    )
    monkeypatch.setattr(instructor_ui, "remember_dialog_dir", lambda *_args, **_kwargs: None)


def test_step1_async_cancelled_reports_status(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common(monkeypatch)
    dummy = _DummyModule()
    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "generate_course_details_template",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(JobCancelledError("cancelled")),
    )

    instructor_ui.InstructorModule._download_course_template_async(dummy)

    assert dummy.step1_done is False
    assert dummy.step1_path is None
    assert dummy.status_changed.messages == ["instructor.status.operation_cancelled"]
    assert dummy._toasts == []


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

    instructor_ui.InstructorModule._upload_course_details_async(dummy)

    assert dummy.step2_upload_ready is False
    assert dummy.step2_course_details_path is None
    assert dummy.status_changed.messages == ["instructor.status.operation_cancelled"]
    assert dummy._toasts == []


def test_step2_prepare_async_cancelled_reports_status(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common(monkeypatch)
    dummy = _DummyModule()
    dummy.step2_upload_ready = True
    dummy.step2_course_details_path = "D:/tmp/course_details.xlsx"
    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("D:/tmp/marks_template.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "generate_marks_template_from_course_details",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(JobCancelledError("cancelled")),
    )

    instructor_ui.InstructorModule._prepare_marks_template_async(dummy)

    assert dummy.step2_done is False
    assert dummy.step2_path is None
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

    instructor_ui.InstructorModule._upload_filled_marks_async(dummy)

    assert dummy.step3_done is False
    assert dummy.step3_path is None
    assert dummy.status_changed.messages == ["instructor.status.operation_cancelled"]
    assert dummy._toasts == []


def test_step5_generate_async_cancelled_reports_status(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _patch_common(monkeypatch)
    src = tmp_path / "filled.xlsx"
    src.write_text("filled", encoding="utf-8")
    dummy = _DummyModule(step3_path=str(src))
    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: (str(tmp_path / "report.xlsx"), ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "_atomic_copy_file",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(JobCancelledError("cancelled")),
    )

    instructor_ui.InstructorModule._generate_final_report_async(dummy)

    assert dummy.step4_done is False
    assert dummy.step4_path is None
    assert dummy.status_changed.messages == ["instructor.status.operation_cancelled"]
    assert dummy._toasts == []
