from __future__ import annotations

import pytest

from common.exceptions import ValidationError
from domain.template_versions.course_setup_v2 import CourseSetupV2Strategy


def test_strategy_validate_workbooks_routes_course_details(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test strategy validate workbooks routes course details.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    strategy = CourseSetupV2Strategy()

    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.course_template_batch_validator",
        lambda: (lambda *, workbook_paths, cancel_token=None: {"valid_paths": list(workbook_paths)}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.marks_template_batch_validator",
        lambda: (lambda *, workbook_paths, template_id, cancel_token=None: {"valid_paths": [], "template_id": template_id}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.co_description_batch_validator",
        lambda: (lambda *, workbook_paths, template_id, cancel_token=None: {"valid_paths": [], "template_id": template_id}),
    )

    result = strategy.validate_workbooks(
        template_id="COURSE_SETUP_V2",
        workbook_kind="course_details",
        workbook_paths=["course.xlsx"],
    )

    if not (result["valid_paths"] == ["course.xlsx"]):
        raise AssertionError('assertion failed')


def test_strategy_validate_workbooks_routes_marks_template(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test strategy validate workbooks routes marks template.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    strategy = CourseSetupV2Strategy()

    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.course_template_batch_validator",
        lambda: (lambda *, workbook_paths, cancel_token=None: {"valid_paths": list(workbook_paths)}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.marks_template_batch_validator",
        lambda: (lambda *, workbook_paths, template_id, cancel_token=None: {"valid_paths": list(workbook_paths), "template_id": template_id}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.co_description_batch_validator",
        lambda: (lambda *, workbook_paths, template_id, cancel_token=None: {"valid_paths": [], "template_id": template_id}),
    )

    result = strategy.validate_workbooks(
        template_id="COURSE_SETUP_V2",
        workbook_kind="marks_template",
        workbook_paths=["marks.xlsx"],
    )

    if not (result["valid_paths"] == ["marks.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["template_id"] == "COURSE_SETUP_V2"):
        raise AssertionError('assertion failed')


def test_strategy_validate_workbooks_routes_co_description(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test strategy validate workbooks routes CO-description template."""
    strategy = CourseSetupV2Strategy()

    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.course_template_batch_validator",
        lambda: (lambda *, workbook_paths, cancel_token=None: {"valid_paths": []}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.marks_template_batch_validator",
        lambda: (lambda *, workbook_paths, template_id, cancel_token=None: {"valid_paths": [], "template_id": template_id}),
    )
    monkeypatch.setattr(
        "domain.template_versions.course_setup_v2_impl.strategy_bindings.co_description_batch_validator",
        lambda: (
            lambda *, workbook_paths, template_id, cancel_token=None: {
                "valid_paths": list(workbook_paths),
                "template_id": template_id,
            }
        ),
    )

    result = strategy.validate_workbooks(
        template_id="COURSE_SETUP_V2",
        workbook_kind="co_description",
        workbook_paths=["co_description.xlsx"],
    )

    if not (result["valid_paths"] == ["co_description.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["template_id"] == "COURSE_SETUP_V2"):
        raise AssertionError('assertion failed')


def test_strategy_validate_workbooks_rejects_unknown_kind() -> None:
    """Test strategy validate workbooks rejects unknown kind.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    strategy = CourseSetupV2Strategy()
    with pytest.raises(ValidationError) as excinfo:
        strategy.validate_workbooks(
            template_id="COURSE_SETUP_V2",
            workbook_kind="bad_kind",
            workbook_paths=["x.xlsx"],
        )
    if not (getattr(excinfo.value, "code", None) == "WORKBOOK_KIND_UNSUPPORTED"):
        raise AssertionError('assertion failed')
