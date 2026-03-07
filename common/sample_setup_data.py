"""Embedded sample setup data used to prefill course template workbooks."""

from __future__ import annotations

from typing import Any
from common.constants import (
    ASSESSMENT_CONFIG_SHEET,
    COURSE_METADATA_SHEET,
    QUESTION_MAP_SHEET,
    STUDENTS_SHEET,
)


SAMPLE_SETUP_DATA: dict[str, list[list[Any]]] = {
    COURSE_METADATA_SHEET: [
        ["Course_Code", "ECE000"],
        ["Course_Name", "SAMPLE COURSE"],
        ["Section", "A"],
        ["Semester", "III"],
        ["Academic_Year", "2025-26"],
        ["Faculty_Name", "ABCCE"],
        ["Total_Outcomes", 6],
    ],
    ASSESSMENT_CONFIG_SHEET: [
        ["S1", 17.5, "YES", "YES", "YES"],
        ["S2", 17.5, "YES", "YES", "YES"],
        ["MSP", 10, "YES", "YES", "YES"],
        ["RLP", 5, "YES", "YES", "YES"],
        ["ESP", 20, "NO", "NO", "YES"],
        ["ESE", 30, "NO", "NO", "YES"],
        ["CSURVEY", 100, "NO", "YES", "NO"],
    ],
    QUESTION_MAP_SHEET: [
        ["S1", 1, 2, 1],
        ["S1", 2, 2, 1],
        ["S1", 3, 2, 2],
        ["S1", 4, 2, 2],
        ["S1", 5, 2, 2],
        ["S1", 6, 16, 1],
        ["S1", 7, 8, 1],
        ["S1", 8, 16, 2],
        ["S2", 1, 2, 3],
        ["S2", 2, 2, 3],
        ["S2", 3, 2, 4],
        ["S2", 4, 2, 4],
        ["S2", 5, 2, 5],
        ["S2", 6, 16, 3],
        ["S2", 7, 8, 5],
        ["S2", 8, 16, 4],
        ["MSP", 1, 20, 3],
        ["MSP", 2, 10, 4],
        ["MSP", 3, 30, 5],
        ["MSP", 4, 20, 6],
        ["MSP", 5, 20, 6],
        ["RLP", 1, 20, 3],
        ["RLP", 2, 10, 4],
        ["RLP", 3, 30, 5],
        ["RLP", 4, 20, 6],
        ["RLP", 5, 20, 6],
        ["ESP", 1, 100, "1,2,3,4,5, 6"],
        ["ESE", 1, 100, "1,2,3,4,5, 6"],
    ],
    STUDENTS_SHEET: [
        ["R101", "STUD1"],
        ["R1032", "STUD2"],
    ],
}
