from __future__ import annotations

from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from modules import instructor_module as instructor_ui


class _DummyLabel:
    def __init__(self) -> None:
        self.text = ""

    def setText(self, value: str) -> None:
        self.text = value


class _DummyModule:
    def __init__(self) -> None:
        self.RAIL_LINKS = (
            ("instructor.links.course_details_generated", "step1_path"),
            ("instructor.links.course_details_uploaded", "step2_course_details_path"),
            ("instructor.links.marks_template_generated", "marks_template_path"),
            ("instructor.links.marks_template_uploaded", "filled_marks_path"),
        )
        self.RAIL_LINK_OPEN_FILE_KEY = "instructor.links.open_file"
        self.RAIL_LINK_OPEN_FOLDER_KEY = "instructor.links.open_folder"
        self.RAIL_LINK_NOT_AVAILABLE_KEY = "instructor.links.not_available"
        self.RAIL_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"
        self.step1_path: str | None = "D:/tmp/course_details_template.xlsx"
        self.step2_course_details_path: str | None = "D:/tmp/course_details_filled.xlsx"
        self.marks_template_path: str | None = "D:/tmp/marks_template.xlsx"
        self.filled_marks_path: str | None = "D:/tmp/marks_filled.xlsx"
        self.quick_link_labels = {
            "instructor.links.course_details_generated": _DummyLabel(),
            "instructor.links.course_details_uploaded": _DummyLabel(),
            "instructor.links.marks_template_generated": _DummyLabel(),
            "instructor.links.marks_template_uploaded": _DummyLabel(),
        }

    def _quick_link_items(self):
        return instructor_ui.InstructorModule._quick_link_items(cast(Any, self))

    def _quick_link_markup(self, link_key: str, path: str | None):
        return instructor_ui.InstructorModule._quick_link_markup(cast(Any, self), link_key, path)


def test_quick_link_markup_uses_filename_and_actions(monkeypatch: pytest.MonkeyPatch) -> None:
    dummy = _DummyModule()
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)

    text = instructor_ui.InstructorModule._quick_link_markup(
        cast(Any, dummy),
        "instructor.links.course_details_generated",
        "D:/tmp/course_details_template.xlsx",
    )

    assert "course_details_template.xlsx" in text
    assert "file::D:/tmp/course_details_template.xlsx" in text
    assert "folder::D:/tmp/course_details_template.xlsx" in text


def test_quick_link_markup_missing_path_shows_not_available(monkeypatch: pytest.MonkeyPatch) -> None:
    dummy = _DummyModule()
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)

    text = instructor_ui.InstructorModule._quick_link_markup(
        cast(Any, dummy),
        "instructor.links.marks_template_uploaded",
        None,
    )

    assert text == "instructor.links.marks_template_uploaded: instructor.links.not_available"


def test_refresh_quick_links_updates_all_rows(monkeypatch: pytest.MonkeyPatch) -> None:
    dummy = _DummyModule()
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)

    instructor_ui.InstructorModule._refresh_quick_links(cast(Any, dummy))

    assert "course_details_template.xlsx" in dummy.quick_link_labels[
        "instructor.links.course_details_generated"
    ].text
    assert "course_details_filled.xlsx" in dummy.quick_link_labels[
        "instructor.links.course_details_uploaded"
    ].text
    assert "marks_template.xlsx" in dummy.quick_link_labels[
        "instructor.links.marks_template_generated"
    ].text
    assert "marks_filled.xlsx" in dummy.quick_link_labels[
        "instructor.links.marks_template_uploaded"
    ].text


def test_open_file_quick_link_uses_desktop_services(monkeypatch: pytest.MonkeyPatch) -> None:
    opened_paths: list[str] = []
    toasts: list[tuple[str, str]] = []

    class _FakeQUrl:
        @staticmethod
        def fromLocalFile(path: str) -> str:
            return path

    monkeypatch.setattr(instructor_ui, "QUrl", _FakeQUrl)
    monkeypatch.setattr(
        instructor_ui.QDesktopServices,
        "openUrl",
        lambda path: opened_paths.append(path) or True,
    )
    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda _parent, message, **kwargs: toasts.append((message, kwargs.get("level", ""))),
    )
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    instructor_ui.InstructorModule._on_quick_link_activated(
        cast(Any, object()),
        "file::D:/tmp/marks_template.xlsx",
    )

    assert len(opened_paths) == 1
    assert opened_paths[0].endswith("marks_template.xlsx")
    assert toasts == []


def test_open_folder_quick_link_uses_parent_dir(monkeypatch: pytest.MonkeyPatch) -> None:
    opened_paths: list[str] = []

    class _FakeQUrl:
        @staticmethod
        def fromLocalFile(path: str) -> str:
            return path

    monkeypatch.setattr(instructor_ui, "QUrl", _FakeQUrl)
    monkeypatch.setattr(
        instructor_ui.QDesktopServices,
        "openUrl",
        lambda path: opened_paths.append(path) or True,
    )

    instructor_ui.InstructorModule._on_quick_link_activated(
        cast(Any, object()),
        "folder::D:/tmp/marks_template.xlsx",
    )

    assert len(opened_paths) == 1
    assert opened_paths[0].replace("\\", "/").endswith("/tmp")


def test_quick_link_open_failure_shows_error_toast(monkeypatch: pytest.MonkeyPatch) -> None:
    toasts: list[tuple[str, str]] = []

    class _FakeQUrl:
        @staticmethod
        def fromLocalFile(path: str) -> str:
            return path

    monkeypatch.setattr(instructor_ui, "QUrl", _FakeQUrl)
    monkeypatch.setattr(instructor_ui.QDesktopServices, "openUrl", lambda _path: False)
    monkeypatch.setattr(
        instructor_ui,
        "show_toast",
        lambda _parent, message, **kwargs: toasts.append((message, kwargs.get("title", ""))),
    )
    monkeypatch.setattr(instructor_ui, "t", lambda key, **_kwargs: key)
    instructor_ui.InstructorModule._on_quick_link_activated(
        cast(Any, _DummyModule()),
        "file::D:/tmp/missing.xlsx",
    )

    assert toasts == [("instructor.links.open_failed", "instructor.msg.error_title")]
