from __future__ import annotations

import pytest

from common.exceptions import ValidationError
from domain.template_versions.course_setup_v2 import CourseSetupV2Strategy


def test_strategy_validate_workbooks_routes_course_details(monkeypatch: pytest.MonkeyPatch) -> None:
    strategy = CourseSetupV2Strategy()

    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2._course_template_batch_validator",
        lambda: (lambda *, workbook_paths, cancel_token=None: {"valid_paths": list(workbook_paths)}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2._marks_template_batch_validator",
        lambda: (lambda *, workbook_paths, template_id, cancel_token=None: {"valid_paths": [], "template_id": template_id}),
    )

    result = strategy.validate_workbooks(
        template_id="COURSE_SETUP_V2",
        workbook_kind="course_details",
        workbook_paths=["course.xlsx"],
    )

    assert result["valid_paths"] == ["course.xlsx"]


def test_strategy_validate_workbooks_routes_marks_template(monkeypatch: pytest.MonkeyPatch) -> None:
    strategy = CourseSetupV2Strategy()

    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2._course_template_batch_validator",
        lambda: (lambda *, workbook_paths, cancel_token=None: {"valid_paths": list(workbook_paths)}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2._marks_template_batch_validator",
        lambda: (lambda *, workbook_paths, template_id, cancel_token=None: {"valid_paths": list(workbook_paths), "template_id": template_id}),
    )

    result = strategy.validate_workbooks(
        template_id="COURSE_SETUP_V2",
        workbook_kind="marks_template",
        workbook_paths=["marks.xlsx"],
    )

    assert result["valid_paths"] == ["marks.xlsx"]
    assert result["template_id"] == "COURSE_SETUP_V2"


def test_strategy_validate_workbooks_rejects_unknown_kind() -> None:
    strategy = CourseSetupV2Strategy()
    with pytest.raises(ValidationError) as excinfo:
        strategy.validate_workbooks(
            template_id="COURSE_SETUP_V2",
            workbook_kind="bad_kind",
            workbook_paths=["x.xlsx"],
        )
    assert getattr(excinfo.value, "code", None) == "WORKBOOK_KIND_UNSUPPORTED"
