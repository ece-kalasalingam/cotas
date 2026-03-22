"""Template-version handlers and strategy implementations."""

from domain.template_versions import course_setup_v1
from domain.template_versions.course_setup_v1 import CourseSetupV1Strategy

__all__ = ["course_setup_v1", "CourseSetupV1Strategy"]
