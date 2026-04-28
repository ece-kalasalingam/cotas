from typing import Any, cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from common.drag_drop_file_widget import ManagedDropFileWidget


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    """Qapp.
    
    Args:
        None.
    
    Returns:
        QApplication: Return value.
    
    Raises:
        None.
    """
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def test_managed_drop_widget_add_set_clear_and_files(qapp: QApplication) -> None:
    """Test managed drop widget add set clear and files.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(drop_mode="multiple")
    changed: list[list[str]] = []
    dropped: list[list[str]] = []
    widget.files_changed.connect(lambda values: changed.append(list(values)))
    widget.files_dropped.connect(lambda values: dropped.append(list(values)))

    added = widget.add_files(["C:/a.xlsx", "C:/a.xlsx", "C:/b.xlsx"])
    if not (added == ["C:/a.xlsx", "C:/b.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["C:/a.xlsx", "C:/b.xlsx"]):
        raise AssertionError('assertion failed')
    if not (dropped[-1] == ["C:/a.xlsx", "C:/b.xlsx"]):
        raise AssertionError('assertion failed')
    if not (changed[-1] == ["C:/a.xlsx", "C:/b.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.summary_label.text() == "Files: 2"):
        raise AssertionError('assertion failed')

    widget.set_files(["D:/x.xlsx"])
    if not (widget.files() == ["D:/x.xlsx"]):
        raise AssertionError('assertion failed')
    if not (changed[-1] == ["D:/x.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.summary_label.text() == "Files: 1"):
        raise AssertionError('assertion failed')

    widget.clear_files()
    if not (widget.files() == []):
        raise AssertionError('assertion failed')
    if not (changed[-1] == []):
        raise AssertionError('assertion failed')
    if not (widget.summary_label.text() == "Files: 0"):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_supports_custom_summary_builder(qapp: QApplication) -> None:
    """Test managed drop widget supports custom summary builder.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(drop_mode="multiple")
    widget.set_summary_text_builder(lambda count: f"Count={count}")
    if not (widget.summary_label.text() == "Count=0"):
        raise AssertionError('assertion failed')
    widget.add_files(["C:/a.xlsx", "C:/b.xlsx"])
    if not (widget.summary_label.text() == "Count=2"):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_single_mode_keeps_latest_drop_batch(qapp: QApplication) -> None:
    """Test managed drop widget single mode keeps latest drop batch.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(drop_mode="single")
    widget.add_files(["C:/first.xlsx", "C:/second.xlsx"])
    if not (widget.files() == ["C:/first.xlsx"]):
        raise AssertionError('assertion failed')
    widget.add_files(["D:/latest.xlsx"])
    if not (widget.files() == ["D:/latest.xlsx"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_extension_filter_accepts_only_allowed(qapp: QApplication) -> None:
    """Test managed drop widget extension filter accepts only allowed.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(allowed_extensions=[".xlsx", "xlsm"])
    rejected: list[list[str]] = []
    widget.files_rejected.connect(lambda values: rejected.append(list(values)))

    added = widget.add_files(["C:/a.xlsx", "C:/b.txt", "C:/c.xlsm"])
    if not (added == ["C:/a.xlsx", "C:/c.xlsm"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["C:/a.xlsx", "C:/c.xlsm"]):
        raise AssertionError('assertion failed')
    if not (rejected[-1] == ["C:/b.txt"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_filename_filter_accepts_specific_names(qapp: QApplication) -> None:
    """Test managed drop widget filename filter accepts specific names.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(allowed_filenames=["course_details.xlsx"])
    added = widget.add_files(["C:/course_details.xlsx", "C:/other.xlsx"])
    if not (added == ["C:/course_details.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["C:/course_details.xlsx"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_custom_predicate_filter(qapp: QApplication) -> None:
    """Test managed drop widget custom predicate filter.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(file_filter=lambda path: "final" in path.lower())
    added = widget.add_files(["C:/final_report.xlsx", "C:/draft_report.xlsx"])
    if not (added == ["C:/final_report.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["C:/final_report.xlsx"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_filter_setters_apply_at_runtime(qapp: QApplication) -> None:
    """Test managed drop widget filter setters apply at runtime.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget()
    widget.set_allowed_extensions(["xlsx"])
    widget.set_allowed_filenames(["course.xlsx"])
    widget.set_file_filter(lambda path: "course" in path.lower())
    added = widget.add_files(["C:/course.xlsx", "C:/course.xlsm", "C:/other.xlsx"])
    if not (added == ["C:/course.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["C:/course.xlsx"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_rejects_duplicates(qapp: QApplication) -> None:
    """Test managed drop widget rejects duplicates.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(drop_mode="multiple")
    rejected: list[list[str]] = []
    widget.files_rejected.connect(lambda values: rejected.append(list(values)))

    first_added = widget.add_files(["C:/same.xlsx", "C:/other.xlsx"])
    if not (first_added == ["C:/same.xlsx", "C:/other.xlsx"]):
        raise AssertionError('assertion failed')
    second_added = widget.add_files(["C:/same.xlsx", "C:/other.xlsx", "C:/new.xlsx"])
    if not (second_added == ["C:/new.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["C:/same.xlsx", "C:/other.xlsx", "C:/new.xlsx"]):
        raise AssertionError('assertion failed')
    if not (rejected[-1] == ["C:/same.xlsx", "C:/other.xlsx"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_rejects_non_local_sources_by_default(qapp: QApplication) -> None:
    """Test managed drop widget rejects non local sources by default.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(drop_mode="multiple")
    rejected: list[list[str]] = []
    widget.files_rejected.connect(lambda values: rejected.append(list(values)))
    added = widget.add_files(["https://example.com/a.xlsx", "C:/local.xlsx"])
    if not (added == ["C:/local.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["C:/local.xlsx"]):
        raise AssertionError('assertion failed')
    if not (rejected[-1] == ["https://example.com/a.xlsx"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_can_allow_non_local_sources(qapp: QApplication) -> None:
    """Test managed drop widget can allow non local sources.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(drop_mode="multiple", allow_non_local_sources=True)
    added = widget.add_files(["https://example.com/a.xlsx"])
    if not (added == ["https://example.com/a.xlsx"]):
        raise AssertionError('assertion failed')
    if not (widget.files() == ["https://example.com/a.xlsx"]):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_applies_tooltips_to_row_actions(qapp: QApplication) -> None:
    """Test managed drop widget applies tooltips to row actions.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(
        drop_mode="multiple",
        open_file_tooltip="Open File",
        open_folder_tooltip="Open Folder",
        remove_tooltip="Remove File",
    )
    widget.add_files(["C:/a.xlsx"])
    item = widget.drop_list.item(0)
    row = cast(Any, widget.drop_list.itemWidget(item))
    if not (row.open_file_btn.toolTip() == "Open File"):
        raise AssertionError('assertion failed')
    if not (row.open_folder_btn.toolTip() == "Open Folder"):
        raise AssertionError('assertion failed')
    if not (row.remove_btn.toolTip() == "Remove File"):
        raise AssertionError('assertion failed')


def test_managed_drop_widget_supports_validation_state_transitions(qapp: QApplication) -> None:
    """Test managed drop widget supports validation state transitions.
    
    Args:
        qapp: Parameter value (QApplication).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    widget = ManagedDropFileWidget(drop_mode="multiple")
    if not (widget.drop_zone.property("validationState") == "neutral"):
        raise AssertionError('assertion failed')
    widget.set_validation_state("info")
    if not (widget.drop_zone.property("validationState") == "info"):
        raise AssertionError('assertion failed')
    widget.set_validation_state("success")
    if not (widget.drop_zone.property("validationState") == "success"):
        raise AssertionError('assertion failed')
    widget.set_validation_state("warning")
    if not (widget.drop_zone.property("validationState") == "warning"):
        raise AssertionError('assertion failed')
    widget.set_validation_state("error")
    if not (widget.drop_zone.property("validationState") == "error"):
        raise AssertionError('assertion failed')
    widget.clear_files()
    if not (widget.drop_zone.property("validationState") == "neutral"):
        raise AssertionError('assertion failed')
