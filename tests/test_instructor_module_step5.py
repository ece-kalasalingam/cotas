from __future__ import annotations

import shutil
from pathlib import Path

import pytest

pytest.importorskip("PySide6")

from modules import instructor_module as instructor_ui


class _SignalRecorder:
    def __init__(self) -> None:
        self.messages: list[str] = []

    def emit(self, message: str) -> None:
        self.messages.append(message)


class _DummyModule:
    def __init__(self, step3_path: str | None = None) -> None:
        self.step3_path = step3_path
        self.step4_path: str | None = None
        self.step4_done = False
        self.step4_outdated = True
        self.status_changed = _SignalRecorder()
        self._can_run = (True, "")
        self._toasts: list[tuple[str, str]] = []

    def _can_run_step(self, _step: int) -> tuple[bool, str]:
        return self._can_run

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        instructor_ui.remember_dialog_dir(selected_path, app_name=instructor_ui.APP_NAME)

    def _show_system_error_toast(self, step: int) -> None:
        self._toasts.append(("system", str(step)))

    def _show_step_success_toast(self, step: int) -> None:
        self._toasts.append(("success", str(step)))


def _patch_common_dependencies(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    monkeypatch.setattr(
        instructor_ui,
        "resolve_dialog_start_path",
        lambda _app, _default=None: "D:/tmp/co_report.xlsx",
    )


def test_step5_blocked_by_dependency_shows_info(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule()
    dummy._can_run = (False, "instructor.require.step3")
    info_calls: list[tuple] = []

    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda parent, message, **kwargs: info_calls.append((parent, message, kwargs)),
    )

    instructor_ui.InstructorModule._generate_final_report(dummy)

    assert len(info_calls) == 1
    assert info_calls[0][1] == "instructor.require.step3"
    assert dummy.step4_done is False
    assert dummy.status_changed.messages == []


def test_step5_cancelled_dialog_keeps_state(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule(step3_path="D:/tmp/in.xlsx")

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("", ""),
    )

    instructor_ui.InstructorModule._generate_final_report(dummy)

    assert dummy.step4_done is False
    assert dummy.step4_path is None
    assert dummy.status_changed.messages == []


def test_step5_missing_source_file_shows_error(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_dependencies(monkeypatch)
    dummy = _DummyModule(step3_path="D:/tmp/missing.xlsx")
    toast_calls: list[tuple] = []

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("D:/tmp/out.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda parent, message, **kwargs: toast_calls.append((parent, message, kwargs)),
    )

    instructor_ui.InstructorModule._generate_final_report(dummy)

    assert len(toast_calls) == 1
    assert toast_calls[0][1] == "instructor.require.step3"
    assert dummy.step4_done is False


def test_step5_copy_failure_shows_error(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _patch_common_dependencies(monkeypatch)
    src = tmp_path / "filled.xlsx"
    src.write_text("data", encoding="utf-8")
    dst = tmp_path / "report.xlsx"
    dummy = _DummyModule(step3_path=str(src))

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: (str(dst), ""),
    )
    monkeypatch.setattr(
        instructor_ui.shutil,
        "copyfile",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(OSError("copy failed")),
    )
    instructor_ui.InstructorModule._generate_final_report(dummy)

    assert dummy._toasts == [("system", "4")]
    assert dummy.step4_done is False
    assert not dst.exists()


def test_step5_success_generates_report(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _patch_common_dependencies(monkeypatch)
    src = tmp_path / "filled.xlsx"
    dst = tmp_path / "report.xlsx"
    src.write_text("filled-marks", encoding="utf-8")
    remembered: list[tuple[str, str]] = []
    dummy = _DummyModule(step3_path=str(src))

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: (str(dst), ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "remember_dialog_dir",
        lambda path, app_name: remembered.append((path, app_name)),
    )

    instructor_ui.InstructorModule._generate_final_report(dummy)

    assert dst.exists()
    assert dst.read_text(encoding="utf-8") == "filled-marks"
    assert dummy.step4_done is True
    assert dummy.step4_outdated is False
    assert dummy.step4_path == str(dst)
    assert remembered == [(str(dst), instructor_ui.APP_NAME)]
    assert dummy.status_changed.messages == ["instructor.status.step4_selected"]
