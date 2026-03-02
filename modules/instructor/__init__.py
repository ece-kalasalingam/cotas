"""Instructor package."""

from .course_details_template_generator import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)

__all__ = [
    "generate_course_details_template",
    "generate_marks_template_from_course_details",
    "validate_course_details_workbook",
]
