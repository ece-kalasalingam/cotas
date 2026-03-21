from __future__ import annotations

from typing import cast

import pytest

pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication

from modules import instructor_module as instructor_ui


@pytest.fixture(scope="module")
def qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return cast(QApplication, app)


def _build_module(monkeypatch: pytest.MonkeyPatch) -> instructor_ui.InstructorModule:
    monkeypatch.setattr(instructor_ui, "t", lambda key, **kwargs: key)
    monkeypatch.setattr(instructor_ui.InstructorModule, "_setup_ui_logging", lambda self: None)
    return instructor_ui.InstructorModule()


def test_shared_outputs_data_step1_contains_template_paths(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    module = _build_module(monkeypatch)
    module.current_step = 1
    module.step1_path = "D:/tmp/course_details_template.xlsx"
    module.marks_template_paths = ["D:/tmp/marks_template_a.xlsx", "D:/tmp/marks_template_b.xlsx"]

    payload = module.get_shared_outputs_data()
    paths = [item.path for item in payload.items]

    assert "D:/tmp/course_details_template.xlsx" in paths
    assert "D:/tmp/marks_template_a.xlsx" in paths
    assert "D:/tmp/marks_template_b.xlsx" in paths
    module.close()


def test_shared_outputs_data_step2_contains_final_report_paths(
    monkeypatch: pytest.MonkeyPatch, qapp: QApplication
) -> None:
    module = _build_module(monkeypatch)
    module.current_step = 2
    module.final_report_paths = ["D:/tmp/final_a.xlsx", "D:/tmp/final_b.xlsx"]

    payload = module.get_shared_outputs_data()
    paths = [item.path for item in payload.items]

    assert paths == ["D:/tmp/final_a.xlsx", "D:/tmp/final_b.xlsx"]
    module.close()
