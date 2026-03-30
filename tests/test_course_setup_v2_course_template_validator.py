from __future__ import annotations

from typing import Any, cast

import pytest

from common.exceptions import ValidationError
from domain.template_versions.course_setup_v2_impl import (
    course_template_validator as validator,
)


def _identity(*, section: str, course_code: str = "CS101", semester: str = "V", year: str = "2026-27", outcomes: int = 3) -> object:
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
    def _fake_validate(*, workbook_path: str, cancel_token: object | None = None) -> object:
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
    def _fake_validate(*, workbook_path: str, cancel_token: object | None = None) -> object:
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
    fn = getattr(validator, fn_name)
    with pytest.raises(ValidationError):
        fn(*args)
