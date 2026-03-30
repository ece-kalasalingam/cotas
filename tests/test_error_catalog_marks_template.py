from __future__ import annotations

from common.error_catalog import resolve_validation_issue


def test_marks_template_catalog_entries_resolve_to_translation_keys() -> None:
    issue = resolve_validation_issue(
        "MARKS_TEMPLATE_COHORT_MISMATCH",
        context={"workbook": "a.xlsx", "fields": "Course_Code"},
    )
    assert issue.translation_key == "validation.marks_template.cohort_mismatch"
    assert issue.severity == "warning"

    issue = resolve_validation_issue(
        "MARKS_TEMPLATE_SECTION_DUPLICATE",
        context={"workbook": "b.xlsx", "section": "A"},
    )
    assert issue.translation_key == "validation.marks_template.duplicate_section"
    assert issue.severity == "warning"

    issue = resolve_validation_issue(
        "MARKS_TEMPLATE_STUDENT_REG_DUPLICATE",
        context={"workbook": "c.xlsx"},
    )
    assert issue.translation_key == "validation.marks_template.duplicate_reg_no"
    assert issue.severity == "warning"
