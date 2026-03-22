"""Embedded sample setup data used to prefill course template workbooks."""

from __future__ import annotations

from typing import Any

from common.constants import (
    ASSESSMENT_CONFIG_SHEET,
    CO_DESCRIPTION_SHEET,
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
        ["S1", 17.5, "YES", "YES", "YES", "FORMATIVE", "THEORY_EXAM", "WRITTEN", "INDIVIDUAL"],
        ["S2", 17.5, "YES", "YES", "YES", "FORMATIVE", "THEORY_EXAM", "WRITTEN", "INDIVIDUAL"],
        ["MSP", 10, "YES", "YES", "YES", "FORMATIVE", "PRACTICAL_EXAM", "HANDS_ON+WRITTEN+ORAL", "INDIVIDUAL"],
        ["RLP", 5, "YES", "YES", "YES", "FORMATIVE", "LAB_WORK", "HANDS_ON+WRITTEN+ORAL", "INDIVIDUAL+GROUP"],
        ["ESP", 20, "NO", "NO", "YES", "SUMMATIVE", "PRACTICAL_EXAM", "HANDS_ON+WRITTEN+ORAL", "INDIVIDUAL"],
        ["ESE", 30, "NO", "NO", "YES", "SUMMATIVE", "THEORY_EXAM", "WRITTEN", "INDIVIDUAL"],
        ["CSURVEY", 100, "NO", "YES", "NO", "SUMMATIVE", "SURVEY", "PRESENTATION+ORAL", "INDIVIDUAL+GROUP"],
    ],
    QUESTION_MAP_SHEET: [
        ["S1", 1, 2, 1,  "UNDERSTAND"],
        ["S1", 2, 2, 1,  "REMEMBER"],
        ["S1", 3, 2, 2,  "UNDERSTAND"],
        ["S1", 4, 2, 2, "UNDERSTAND"],
        ["S1", 5, 2, 2,  "UNDERSTAND"],
        ["S1", 6, 16, 1,  "APPLY"],
        ["S1", 7, 8, 1,  "APPLY"],
        ["S1", 8, 16, 2,  "ANALYZE"],
        ["S2", 1, 2, 3,  "UNDERSTAND"],
        ["S2", 2, 2, 3,  "REMEMBER"],
        ["S2", 3, 2, 4,  "UNDERSTAND"],
        ["S2", 4, 2, 4,  "UNDERSTAND"],
        ["S2", 5, 2, 5,  "UNDERSTAND"],
        ["S2", 6, 16, 3,  "APPLY"],
        ["S2", 7, 8, 5,  "APPLY"],
        ["S2", 8, 16, 4,  "ANALYZE"],
        ["MSP", 1, 20, 3,  "SKILL_DEVELOPMENT"],
        ["MSP", 2, 10, 4, "SKILL_DEVELOPMENT"],
        ["MSP", 3, 30, 5,  "SKILL_DEVELOPMENT"],
        ["MSP", 4, 20, 6,  "SKILL_DEVELOPMENT"],
        ["MSP", 5, 20, 6,  "SKILL_DEVELOPMENT"],
        ["RLP", 1, 20, 3,  "SKILL_DEVELOPMENT"],
        ["RLP", 2, 10, 4, "SKILL_DEVELOPMENT"],
        ["RLP", 3, 30, 5, "SKILL_DEVELOPMENT"],
        ["RLP", 4, 20, 6, "SKILL_DEVELOPMENT"],
        ["RLP", 5, 20, 6, "SKILL_DEVELOPMENT"],
        ["ESP", 1, 100, "1,2,3,4,5, 6",  "MULTIPLE_LEVELS"],
        ["ESE", 1, 100, "1,2,3,4,5, 6", "MULTIPLE_LEVELS"],
    ],
    CO_DESCRIPTION_SHEET: [
        [
            1,
            "Recall and explain core concepts",
            "REMEMBER",
            "Covers the fundamental concepts, terminology, and baseline definitions needed for the course, with guided examples and short recall checks.",
        ],
        [
            2,
            "Interpret and discuss theoretical foundations",
            "UNDERSTAND",
            "Focuses on explaining principles in context, interpreting relationships between ideas, and discussing how core theories map to engineering scenarios.",
        ],
        [
            3,
            "Apply concepts to practical problems",
            "APPLY",
            "Includes structured problem-solving exercises where learners select relevant methods and apply formulas or procedures to realistic tasks.",
        ],
        [
            4,
            "Analyze systems and identify patterns",
            "ANALYZE",
            "Develops analytical reasoning by breaking down systems into components, identifying constraints, and comparing alternative interpretations.",
        ],
        [
            5,
            "Evaluate alternatives with justification",
            "EVALUATE",
            "Trains students to assess competing options against criteria, justify choices with evidence, and communicate trade-offs clearly and logically.",
        ],
        [
            6,
            "Design and create solution artifacts",
            "CREATE",
            "Emphasizes project-oriented synthesis through design tasks, prototype planning, implementation details, and reflective iteration on outputs.",
        ],
    ],
    STUDENTS_SHEET: [
        ["R101", "STUD1"],
        ["R1032", "STUD2"],
    ],
}
