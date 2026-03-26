"""COURSE_SETUP_V2 business semantics shared across template workflows."""

from __future__ import annotations

from common.utils import sanitize_filename_token


def _filename_token(value: object) -> str:
    token = sanitize_filename_token(str(value or "").strip())
    return token if token else "NA"


def build_course_template_filename_base() -> str:
    return "Course_Details_Template"


def build_marks_template_filename_base_from_identity(
    *,
    academic_year: object,
    course_code: object,
    semester: object,
    section: object,
) -> str:
    ay_token = _filename_token(academic_year)
    course_token = _filename_token(course_code)
    semester_token = _filename_token(semester)
    section_token = _filename_token(section)
    return f"{ay_token}_{course_token}_{semester_token}_{section_token}_Marks"


__all__ = [
    "build_course_template_filename_base",
    "build_marks_template_filename_base_from_identity",
]
