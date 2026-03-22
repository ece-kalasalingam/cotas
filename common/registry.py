from __future__ import annotations

from common.constants import (
    ASSESSMENT_CONFIG_HEADERS,
    ASSESSMENT_FORMAT_OPTIONS,
    ASSESSMENT_CONFIG_SHEET,
    ASSESSMENT_MODE_OPTIONS,
    ASSESSMENT_PARTICIPATION_OPTIONS,
    ASSESSMENT_TYPE_OPTIONS,
    CO_DESCRIPTION_SUMMARY_MAX_LENGTH,
    CO_DESCRIPTION_SUMMARY_MIN_LENGTH,
    ASSESSMENT_VALIDATION_INPUT_MESSAGE,
    ASSESSMENT_VALIDATION_INPUT_TITLE,
    ASSESSMENT_VALIDATION_LAST_ROW,
    ASSESSMENT_VALIDATION_YES_NO_OPTIONS,
    CO_DESCRIPTION_HEADERS,
    CO_DESCRIPTION_SHEET,
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SHEET,
    ID_COURSE_SETUP,
    QUESTION_DOMAIN_LEVEL_OPTIONS,
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
                ),
                ValidationRule(
                    first_row=1,
                    first_col=5,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=5,
                    options={
                        "validate": "list",
                        "source": list(ASSESSMENT_TYPE_OPTIONS),
                        "input_title": "Assessment Type",
                        "input_message": "Select a valid Assessment_Type value",
                    },
                ),
                ValidationRule(
                    first_row=1,
                    first_col=6,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=6,
                    options={
                        "validate": "list",
                        "source": list(ASSESSMENT_FORMAT_OPTIONS),
                        "input_title": "Assessment Format",
                        "input_message": "Select a valid Assessment_Format value",
                    },
                ),
                ValidationRule(
                    first_row=1,
                    first_col=7,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=7,
                    options={
                        "validate": "list",
                        "source": list(ASSESSMENT_MODE_OPTIONS),
                        "input_title": "Mode",
                        "input_message": "Select a valid Mode value",
                    },
                ),
                ValidationRule(
                    first_row=1,
                    first_col=8,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=8,
                    options={
                        "validate": "list",
                        "source": list(ASSESSMENT_PARTICIPATION_OPTIONS),
                        "input_title": "Participation",
                        "input_message": "Select a valid Participation value",
                    },
                ),
            ],
            is_protected=False,
        ),
        SheetSchema(
            name=QUESTION_MAP_SHEET,
            header_matrix=[list(QUESTION_MAP_HEADERS)],
            validations=[
                ValidationRule(
                    first_row=1,
                    first_col=4,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=4,
                    options={
                        "validate": "list",
                        "source": list(QUESTION_DOMAIN_LEVEL_OPTIONS),
                        "input_title": "Bloom Level",
                        "input_message": "Select a valid Bloom_Level value",
                    },
                )
            ],
            is_protected=False,
        ),
        SheetSchema(
            name=CO_DESCRIPTION_SHEET,
            header_matrix=[list(CO_DESCRIPTION_HEADERS)],
            validations=[
                ValidationRule(
                    first_row=1,
                    first_col=0,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=0,
                    options={
                        "validate": "integer",
                        "criteria": ">",
                        "value": 0,
                        "input_title": "CO Number",
                        "input_message": "Enter a whole number greater than zero",
                    },
                ),
                ValidationRule(
                    first_row=1,
                    first_col=2,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=2,
                    options={
                        "validate": "list",
                        "source": list(QUESTION_DOMAIN_LEVEL_OPTIONS),
                        "input_title": "Domain Level",
                        "input_message": "Select a valid Domain_Level value",
                    },
                ),
                ValidationRule(
                    first_row=1,
                    first_col=3,
                    last_row=ASSESSMENT_VALIDATION_LAST_ROW,
                    last_col=3,
                    options={
                        "validate": "length",
                        "criteria": "between",
                        "minimum": CO_DESCRIPTION_SUMMARY_MIN_LENGTH,
                        "maximum": CO_DESCRIPTION_SUMMARY_MAX_LENGTH,
                        "input_title": "Summary Length",
                        "input_message": (
                            f"Enter text between {CO_DESCRIPTION_SUMMARY_MIN_LENGTH} and "
                            f"{CO_DESCRIPTION_SUMMARY_MAX_LENGTH} characters"
                        ),
                    },
                ),
            ],
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
