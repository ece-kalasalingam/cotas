from __future__ import annotations

import sys
from types import SimpleNamespace
from typing import Any, cast

import pytest

from common.error_catalog import validation_error_from_key
from common.exceptions import ValidationError
from domain.template_versions.course_setup_v2_impl import (
    course_template_validator as validator,
)


def _identity(*, section: str, course_code: str = "CS101", semester: str = "V", year: str = "2026-27", outcomes: int = 3) -> object:
    """Identity.
    
    Args:
        section: Parameter value (str).
        course_code: Parameter value (str).
        semester: Parameter value (str).
        year: Parameter value (str).
        outcomes: Parameter value (int).
    
    Returns:
        object: Return value.
    
    Raises:
        None.
    """
    return validator._CourseIdentity(
        template_id="COURSE_SETUP_V2",
        course_code=course_code,
        semester=semester,
        academic_year=year,
        total_outcomes=outcomes,
        section=section,
    )


def test_course_batch_validator_accepts_mixed_cohorts_and_sections(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test course batch validator accepts mixed cohorts and sections.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _fake_validate(*, workbook_path: str, cancel_token: object | None = None) -> object:
        """Fake validate.
        
        Args:
            workbook_path: Parameter value (str).
            cancel_token: Parameter value (object | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        del cancel_token
        if workbook_path == "a.xlsx":
            return _identity(section="A")
        if workbook_path == "b.xlsx":
            return _identity(section="B")
        if workbook_path == "dup_section.xlsx":
            return _identity(section="A")
        return _identity(section="C", year="2025-26")

    monkeypatch.setattr(validator, "_validate_course_details_workbook_impl", _fake_validate)
    result = cast(
        dict[str, Any],
        validator.validate_course_details_workbooks(
            ["a.xlsx", "a.xlsx", "b.xlsx", "dup_section.xlsx", "cohort_bad.xlsx"]
        ),
    )

    assert result["valid_paths"] == ["a.xlsx", "b.xlsx", "dup_section.xlsx", "cohort_bad.xlsx"]
    assert result["duplicate_paths"] == ["a.xlsx"]
    assert result["duplicate_sections"] == []
    assert result["mismatched_paths"] == []
    assert result["invalid_paths"] == []
    reason_by_path = {
        item["path"]: item["reason_kind"]
        for item in cast(list[dict[str, Any]], result["rejections"])
    }
    assert reason_by_path["a.xlsx"] == "duplicate_path"


def test_course_batch_validator_maps_validation_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test course batch validator maps validation failure.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _fake_validate(*, workbook_path: str, cancel_token: object | None = None) -> object:
        """Fake validate.
        
        Args:
            workbook_path: Parameter value (str).
            cancel_token: Parameter value (object | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        del cancel_token
        if workbook_path == "bad.xlsx":
            raise ValidationError("Invalid workbook", code="SHEET_DATA_REQUIRED", context={"sheet_name": "Students"})
        return _identity(section="A")

    monkeypatch.setattr(validator, "_validate_course_details_workbook_impl", _fake_validate)
    result = cast(
        dict[str, Any],
        validator.validate_course_details_workbooks(["ok.xlsx", "bad.xlsx"]),
    )

    assert result["valid_paths"] == ["ok.xlsx"]
    assert result["invalid_paths"] == ["bad.xlsx"]
    rejection = next(
        item
        for item in cast(list[dict[str, Any]], result["rejections"])
        if item["path"] == "bad.xlsx"
    )
    assert rejection["reason_kind"] == "invalid"
    assert rejection["issue"]["code"] == "SHEET_DATA_REQUIRED"


@pytest.mark.parametrize(
    ("fn_name", "args"),
    [
        ("_validate_course_metadata_rules", ({},)),
        ("_validate_assessment_config_rules", ({},)),
        ("_validate_question_map_rules", ({}, {}, 3)),
        ("_validate_students_rules", ({},)),
    ],
)
def test_course_rule_validators_fail_for_missing_required_data(fn_name: str, args: tuple[object, ...]) -> None:
    """Test course rule validators fail for missing required data.
    
    Args:
        fn_name: Parameter value (str).
        args: Parameter value (tuple[object, ...]).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    fn = getattr(validator, fn_name)
    with pytest.raises(ValidationError):
        fn(*args)


def test_course_workbook_impl_raises_open_failed_for_corrupt_workbook(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    """Test course workbook impl raises open failed for corrupt workbook.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_path = tmp_path / "corrupt_course.xlsx"
    workbook_path.write_text("not-an-xlsx", encoding="utf-8")

    def _raise_open(*_args, **_kwargs):
        """Raise open.
        
        Args:
            _args: Parameter value.
            _kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        raise OSError("corrupt workbook stream")

    monkeypatch.setitem(
        sys.modules,
        "openpyxl",
        SimpleNamespace(load_workbook=_raise_open),
    )

    with pytest.raises(ValidationError) as excinfo:
        validator._validate_course_details_workbook_impl(workbook_path=workbook_path)

    assert excinfo.value.code == "WORKBOOK_OPEN_FAILED"
    assert str(excinfo.value.context.get("workbook", "")) == str(workbook_path)


def test_course_workbook_impl_rejects_symlink_path(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    workbook_path = tmp_path / "linked_course.xlsx"
    workbook_path.write_text("x", encoding="utf-8")

    monkeypatch.setitem(
        sys.modules,
        "openpyxl",
        SimpleNamespace(load_workbook=lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("unexpected"))),
    )
    monkeypatch.setattr(
        validator,
        "assert_not_symlink_path",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_SYMLINK_NOT_ALLOWED",
                workbook=str(workbook_path),
            )
        ),
    )

    with pytest.raises(ValidationError) as excinfo:
        validator._validate_course_details_workbook_impl(workbook_path=workbook_path)

    assert excinfo.value.code == "WORKBOOK_SYMLINK_NOT_ALLOWED"
