from __future__ import annotations

import pytest

pytest.importorskip("PySide6")

from modules import instructor_module as instructor_ui


class _SignalRecorder:
    def __init__(self) -> None:
        self.messages: list[str] = []

    def emit(self, message: str) -> None:
        self.messages.append(message)


class _DummyModule:
    def __init__(self) -> None:
        self.step2_done = False
        self.step3_path: str | None = None
        self.step3_done = False
        self.step3_outdated = True
        self.step4_outdated = False
        self.status_changed = _SignalRecorder()
        self._toasts: list[tuple[str, str]] = []

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        instructor_ui.remember_dialog_dir(selected_path, app_name=instructor_ui.APP_NAME)

    def _show_step_success_toast(self, step: int) -> None:
        self._toasts.append(("success", str(step)))


def _patch_common_dependencies(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    monkeypatch.setattr(
        instructor_ui,
        "resolve_dialog_start_path",
        lambda _app, _default=None: "D:/tmp/filled_marks.xlsx",
    )


def test_step3_upload_cancel_keeps_state(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("", ""),
    )

    instructor_ui.InstructorModule._upload_filled_marks(dummy)

    assert dummy.step3_done is False
    assert dummy.step3_path is None
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

    instructor_ui.InstructorModule._upload_filled_marks(dummy)

    assert dummy.step2_done is False
    assert dummy.step3_done is True
    assert dummy.step3_outdated is False
    assert dummy.step3_path == "D:/tmp/filled_marks.xlsx"
    assert dummy.step4_outdated is False
    assert remembered == [("D:/tmp/filled_marks.xlsx", instructor_ui.APP_NAME)]
    assert dummy.status_changed.messages == ["instructor.status.step3_uploaded"]


def test_step3_replace_marks_flags_step4_outdated(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    dummy.step3_done = True
    dummy.step3_path = "D:/tmp/old_filled_marks.xlsx"

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getOpenFileName",
        lambda *_args, **_kwargs: ("D:/tmp/new_filled_marks.xlsx", ""),
    )
    monkeypatch.setattr(instructor_ui, "remember_dialog_dir", lambda *_args, **_kwargs: None)

    instructor_ui.InstructorModule._upload_filled_marks(dummy)

    assert dummy.step3_done is True
    assert dummy.step3_outdated is False
    assert dummy.step3_path == "D:/tmp/new_filled_marks.xlsx"
    assert dummy.step4_outdated is True
    assert dummy.status_changed.messages == ["instructor.status.step3_changed"]
