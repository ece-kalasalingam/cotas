from typing import Dict

from scripts.logic_library import check_conditional_weight_sum, check_multiple_columns_empty
from scripts.rules import BusinessRule
from scripts.sheet_schema import SheetSchema, StyleDefinition, ValidationRule, WorkbookBlueprint


SETUP_STYLE_REGISTRY: Dict[str, StyleDefinition] = {
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

DIRECT_TOOLS_SUM_WEIGHT_RULE = BusinessRule(
    rule_id="direct_sum_100",
    scope="INTRA",
    logic_fn=check_conditional_weight_sum,
    versioned_params={
        "COURSE_SETUP_V1": {
            "sheet_key": "assess",
            "weight_col_key": "w",
            "direct_col_key": "direct",
            "condition_val": "YES",
            "target": 100,
        }
    },
)

INDIRECT_TOOLS_SUM_WEIGHT_RULE = BusinessRule(
    rule_id="indirect_sum_100",
    scope="INTRA",
    logic_fn=check_conditional_weight_sum,
    versioned_params={
        "COURSE_SETUP_V1": {
            "sheet_key": "assess",
            "weight_col_key": "w",
            "direct_col_key": "direct",
            "condition_val": "NO",
            "target": 100,
        }
    },
)

CHECK_SETUP_STUDENTS_EMPTY_RULE = BusinessRule(
    rule_id="check_students_empty",
    scope="INTRA",
    logic_fn=check_multiple_columns_empty,
    versioned_params={"COURSE_SETUP_V1": {"sheet_key": "students", "col_keys": ["id", "name"]}},
)

COURSE_SETUP_BP = WorkbookBlueprint(
    type_id="COURSE_SETUP_V1",
    style_registry=SETUP_STYLE_REGISTRY,
    business_rules=[
        DIRECT_TOOLS_SUM_WEIGHT_RULE,
        INDIRECT_TOOLS_SUM_WEIGHT_RULE,
        CHECK_SETUP_STUDENTS_EMPTY_RULE,
    ],
    key_map={
        "assess": "Assessment_Config",
        "assess.w": "Weight (%)",
        "assess.direct": "Direct",
        "students": "Students",
        "students.id": "Reg_No",
        "students.name": "Student_Name",
    },
    sheets=[
        SheetSchema(name="Course_Metadata", header_matrix=[["Field", "Value"]], is_protected=False),
        SheetSchema(
            name="Assessment_Config",
            header_matrix=[["Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct"]],
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
        SheetSchema(name="Question_Map", header_matrix=[["Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO"]], is_protected=False),
        SheetSchema(name="Students", header_matrix=[["Reg_No", "Student_Name"]], is_protected=False),
    ],
)

BLUEPRINT_REGISTRY = {
    "COURSE_SETUP_V1": COURSE_SETUP_BP,
    "MARKS_ENTRY_V1": None,
}
