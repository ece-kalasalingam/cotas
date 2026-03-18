from typing import cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from common.drag_drop_file_widget import ManagedDropFileWidget


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def test_managed_drop_widget_add_set_clear_and_files(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(drop_mode="multiple")
    changed: list[list[str]] = []
    dropped: list[list[str]] = []
    widget.files_changed.connect(lambda values: changed.append(list(values)))
    widget.files_dropped.connect(lambda values: dropped.append(list(values)))

    added = widget.add_files(["C:/a.xlsx", "C:/a.xlsx", "C:/b.xlsx"])
    assert added == ["C:/a.xlsx", "C:/b.xlsx"]
    assert widget.files() == ["C:/a.xlsx", "C:/b.xlsx"]
    assert dropped[-1] == ["C:/a.xlsx", "C:/b.xlsx"]
    assert changed[-1] == ["C:/a.xlsx", "C:/b.xlsx"]

    widget.set_files(["D:/x.xlsx"])
    assert widget.files() == ["D:/x.xlsx"]
    assert changed[-1] == ["D:/x.xlsx"]

    widget.clear_files()
    assert widget.files() == []
    assert changed[-1] == []


def test_managed_drop_widget_single_mode_keeps_latest_drop_batch(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(drop_mode="single")
    widget.add_files(["C:/first.xlsx", "C:/second.xlsx"])
    assert widget.files() == ["C:/first.xlsx"]
    widget.add_files(["D:/latest.xlsx"])
    assert widget.files() == ["D:/latest.xlsx"]


def test_managed_drop_widget_extension_filter_accepts_only_allowed(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(allowed_extensions=[".xlsx", "xlsm"])
    rejected: list[list[str]] = []
    widget.files_rejected.connect(lambda values: rejected.append(list(values)))

    added = widget.add_files(["C:/a.xlsx", "C:/b.txt", "C:/c.xlsm"])
    assert added == ["C:/a.xlsx", "C:/c.xlsm"]
    assert widget.files() == ["C:/a.xlsx", "C:/c.xlsm"]
    assert rejected[-1] == ["C:/b.txt"]


def test_managed_drop_widget_filename_filter_accepts_specific_names(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(allowed_filenames=["course_details.xlsx"])
    added = widget.add_files(["C:/course_details.xlsx", "C:/other.xlsx"])
    assert added == ["C:/course_details.xlsx"]
    assert widget.files() == ["C:/course_details.xlsx"]


def test_managed_drop_widget_custom_predicate_filter(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(file_filter=lambda path: "final" in path.lower())
    added = widget.add_files(["C:/final_report.xlsx", "C:/draft_report.xlsx"])
    assert added == ["C:/final_report.xlsx"]
    assert widget.files() == ["C:/final_report.xlsx"]


def test_managed_drop_widget_filter_setters_apply_at_runtime(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget()
    widget.set_allowed_extensions(["xlsx"])
    widget.set_allowed_filenames(["course.xlsx"])
    widget.set_file_filter(lambda path: "course" in path.lower())
    added = widget.add_files(["C:/course.xlsx", "C:/course.xlsm", "C:/other.xlsx"])
    assert added == ["C:/course.xlsx"]
    assert widget.files() == ["C:/course.xlsx"]


def test_managed_drop_widget_rejects_duplicates(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(drop_mode="multiple")
    rejected: list[list[str]] = []
    widget.files_rejected.connect(lambda values: rejected.append(list(values)))

    first_added = widget.add_files(["C:/same.xlsx", "C:/other.xlsx"])
    assert first_added == ["C:/same.xlsx", "C:/other.xlsx"]
    second_added = widget.add_files(["C:/same.xlsx", "C:/other.xlsx", "C:/new.xlsx"])
    assert second_added == ["C:/new.xlsx"]
    assert widget.files() == ["C:/same.xlsx", "C:/other.xlsx", "C:/new.xlsx"]
    assert rejected[-1] == ["C:/same.xlsx", "C:/other.xlsx"]


def test_managed_drop_widget_rejects_non_local_sources_by_default(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(drop_mode="multiple")
    rejected: list[list[str]] = []
    widget.files_rejected.connect(lambda values: rejected.append(list(values)))
    added = widget.add_files(["https://example.com/a.xlsx", "C:/local.xlsx"])
    assert added == ["C:/local.xlsx"]
    assert widget.files() == ["C:/local.xlsx"]
    assert rejected[-1] == ["https://example.com/a.xlsx"]


def test_managed_drop_widget_can_allow_non_local_sources(qapp: QApplication) -> None:
    widget = ManagedDropFileWidget(drop_mode="multiple", allow_non_local_sources=True)
    added = widget.add_files(["https://example.com/a.xlsx"])
    assert added == ["https://example.com/a.xlsx"]
    assert widget.files() == ["https://example.com/a.xlsx"]
