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
        self.step1_path: str | None = None
        self.step1_done = False
        self.status_changed = _SignalRecorder()

    def _remember_dialog_dir_safe(self, selected_path: str) -> None:
        instructor_ui.remember_dialog_dir(selected_path, app_name=instructor_ui.APP_NAME)


def _patch_common_ui_dependencies(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    monkeypatch.setattr(
        instructor_ui,
        "resolve_dialog_start_path",
        lambda _app, _default=None: "D:/tmp/default.xlsx",
    )


def test_download_course_template_cancel(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_ui_dependencies(monkeypatch)
    dummy = _DummyModule()
    calls = {"generated": 0, "remembered": 0, "toast": 0}

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "generate_course_details_template",
        lambda *_args, **_kwargs: calls.__setitem__("generated", calls["generated"] + 1),
    )
    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda *_args, **_kwargs: calls.__setitem__("toast", calls["toast"] + 1),
    )
    monkeypatch.setattr(
        instructor_ui,
        "remember_dialog_dir",
        lambda *_args, **_kwargs: calls.__setitem__("remembered", calls["remembered"] + 1),
    )

    instructor_ui.InstructorModule._download_course_template(dummy)

    assert dummy.step1_done is False
    assert dummy.step1_path is None
    assert calls["generated"] == 0
    assert calls["remembered"] == 0
    assert calls["toast"] == 0
    assert dummy.status_changed.messages == []


def test_download_course_template_success(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_ui_dependencies(monkeypatch)
    dummy = _DummyModule()
    calls = {"generated": [], "remembered": []}

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "generate_course_details_template",
        lambda path, template_id: calls["generated"].append((path, template_id)),
    )
    monkeypatch.setattr(
        instructor_ui,
        "remember_dialog_dir",
        lambda path, app_name: calls["remembered"].append((path, app_name)),
    )

    instructor_ui.InstructorModule._download_course_template(dummy)

    assert dummy.step1_done is True
    assert dummy.step1_path == "D:/tmp/course_details.xlsx"
    assert calls["generated"] == [("D:/tmp/course_details.xlsx", instructor_ui.ID_COURSE_SETUP)]
    assert calls["remembered"] == [("D:/tmp/course_details.xlsx", instructor_ui.APP_NAME)]
    assert dummy.status_changed.messages == ["instructor.status.step1_selected"]


def test_download_course_template_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_ui_dependencies(monkeypatch)
    dummy = _DummyModule()
    calls = {"toast": [], "remembered": 0}

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details.xlsx", ""),
    )

    def _raise_error(*_args, **_kwargs):
        raise RuntimeError("generation failed")

    monkeypatch.setattr(instructor_ui, "generate_course_details_template", _raise_error)
    monkeypatch.setattr(
        instructor_ui,
        "remember_dialog_dir",
        lambda *_args, **_kwargs: calls.__setitem__("remembered", calls["remembered"] + 1),
    )
    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda parent, message, **kwargs: calls["toast"].append((parent, message, kwargs)),
    )

    instructor_ui.InstructorModule._download_course_template(dummy)

    assert dummy.step1_done is False
    assert dummy.step1_path is None
    assert calls["remembered"] == 0
    assert dummy.status_changed.messages == []
    assert len(calls["toast"]) == 1
    assert calls["toast"][0][1] == "app.unexpected_error"
    assert calls["toast"][0][2]["title"] == "instructor.msg.step_required_title"


def test_download_course_template_validation_error_message(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_common_ui_dependencies(monkeypatch)
    dummy = _DummyModule()
    calls = {"toast": []}

    monkeypatch.setattr(
        instructor_ui.QFileDialog,
        "getSaveFileName",
        lambda *_args, **_kwargs: ("D:/tmp/course_details.xlsx", ""),
    )
    monkeypatch.setattr(
        instructor_ui,
        "generate_course_details_template",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(ValidationError("bad template")),
    )
    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda parent, message, **kwargs: calls["toast"].append((parent, message, kwargs)),
    )

    instructor_ui.InstructorModule._download_course_template(dummy)

    assert dummy.step1_done is False
    assert dummy.step1_path is None
    assert len(calls["toast"]) == 1
    assert calls["toast"][0][1] == "bad template"


def test_download_course_template_dialog_receives_expected_defaults(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _patch_common_ui_dependencies(monkeypatch)
    dummy = _DummyModule()
    captured: dict[str, str] = {}

    def _fake_dialog(_self, title, start_path, file_filter):
        captured["title"] = title
        captured["start_path"] = start_path
        captured["filter"] = file_filter
        return ("", "")

    monkeypatch.setattr(instructor_ui.QFileDialog, "getSaveFileName", _fake_dialog)
    monkeypatch.setattr(instructor_ui, "generate_course_details_template", lambda *_args, **_kwargs: None)

    instructor_ui.InstructorModule._download_course_template(dummy)

    assert captured["title"] == "instructor.dialog.step1.title"
    assert captured["start_path"] == "D:/tmp/default.xlsx"
    assert captured["filter"] == "instructor.dialog.filter.excel"
