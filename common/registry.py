from __future__ import annotations

from common.constants import (
    ASSESSMENT_CONFIG_HEADERS,
    ASSESSMENT_CONFIG_SHEET,
    ASSESSMENT_VALIDATION_INPUT_MESSAGE,
    ASSESSMENT_VALIDATION_INPUT_TITLE,
    ASSESSMENT_VALIDATION_LAST_ROW,
    ASSESSMENT_VALIDATION_YES_NO_OPTIONS,
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SHEET,
    ID_COURSE_SETUP,
    QUESTION_MAP_HEADERS,
    QUESTION_MAP_SHEET,
    STUDENTS_HEADERS,
    STUDENTS_SHEET,
)
from common.sheet_schema import SheetSchema, ValidationRule, WorkbookBlueprint


SETUP_STYLE_REGISTRY = {
    "header": {
        "bold": True,
        "bg_color": "#D9EAD3",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    },
    "body": {
        "locked": False,
        "border": 1,
    },
}


COURSE_SETUP_BP = WorkbookBlueprint(
    type_id=ID_COURSE_SETUP,
    style_registry=SETUP_STYLE_REGISTRY,
    sheets=[
        SheetSchema(
            name=COURSE_METADATA_SHEET,
            header_matrix=[list(COURSE_METADATA_HEADERS)],
            is_protected=False,
        ),
        SheetSchema(
            name=ASSESSMENT_CONFIG_SHEET,
            header_matrix=[list(ASSESSMENT_CONFIG_HEADERS)],
            validations=[
                ValidationRule(
                    first_row=1,
                    first_col=4,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=4,
                    options={
                        "validate": "list",
                        "source": list(ASSESSMENT_VALIDATION_YES_NO_OPTIONS),
                        "input_title": ASSESSMENT_VALIDATION_INPUT_TITLE,
                        "input_message": ASSESSMENT_VALIDATION_INPUT_MESSAGE,
                    },
                )
            ],
            is_protected=False,
        ),
        SheetSchema(
            name=QUESTION_MAP_SHEET,
            header_matrix=[list(QUESTION_MAP_HEADERS)],
            is_protected=False,
        ),
        SheetSchema(
            name=STUDENTS_SHEET,
            header_matrix=[list(STUDENTS_HEADERS)],
            is_protected=False,
        ),
    ],
)


BLUEPRINT_REGISTRY = {
    ID_COURSE_SETUP: COURSE_SETUP_BP,
}
