from __future__ import annotations

from common.constants import ID_COURSE_SETUP
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
            name="Course_Metadata",
            header_matrix=[["Field", "Value"]],
            is_protected=False,
        ),
        SheetSchema(
            name="Assessment_Config",
            header_matrix=[
                ["Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct"]
            ],
            validations=[
                ValidationRule(
                    first_row=1,
                    first_col=4,
                    last_row=300,
                    last_col=4,
                    options={
                        "validate": "list",
                        "source": ["YES", "NO"],
                        "input_title": "Direct",
                        "input_message": "Select YES or NO",
                    },
                )
            ],
            is_protected=False,
        ),
        SheetSchema(
            name="Question_Map",
            header_matrix=[["Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO"]],
            is_protected=False,
        ),
        SheetSchema(
            name="Students",
            header_matrix=[["Reg_No", "Student_Name"]],
            is_protected=False,
        ),
    ],
)


BLUEPRINT_REGISTRY = {
    ID_COURSE_SETUP: COURSE_SETUP_BP,
}
