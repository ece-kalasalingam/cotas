"""Instructor workbook use-cases exposed to non-UI layers."""

from __future__ import annotations

from domain.instructor_template_engine import (
    generate_course_details_template,
    generate_marks_template_from_course_details,
    validate_course_details_workbook,
)
from domain.instructor_report_engine import generate_final_co_report

__all__ = [
    "generate_course_details_template",
    "generate_marks_template_from_course_details",
    "generate_final_co_report",
    "validate_course_details_workbook",
]
