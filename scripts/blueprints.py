from scripts.sheet_schema import StyleDefinition, WorkbookBlueprint, SheetSchema, ValidationRule
from typing import Dict
from scripts.rules import BusinessRule
from scripts.logic_library import (
    check_conditional_weight_sum,
    check_cross_workbook_sync,
    check_multiple_columns_empty
)

# Define a shared style registry for the Setup phase
SETUP_STYLE_REGISTRY: Dict[str, StyleDefinition] = {
    'header': {
        'bold': True,
        'bg_color': '#D9EAD3',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    },
    'body': {
        'locked': False,
        'border': 1
    }
}

# Define the Rule
DIRECT_TOOLS_SUM_WEIGHT_RULE = BusinessRule(
    rule_id="direct_sum_100",
    scope="INTRA",
    logic_fn=check_conditional_weight_sum,
    versioned_params={
        "COURSE_SETUP_V1": {"sheet_key": "assess", "weight_col_key": "w", "direct_col_key": "direct", "condition_val": "YES", "target": 100}
    }
)
INDIRECT_TOOLS_SUM_WEIGHT_RULE = BusinessRule(
    rule_id="indirect_sum_100",
    scope="INTRA",
    logic_fn=check_conditional_weight_sum,
    versioned_params={
        "COURSE_SETUP_V1": {"sheet_key": "assess", "weight_col_key": "w", "direct_col_key": "direct", "condition_val": "NO", "target": 100}
    }
)
CHECK_SETUP_STUDENTS_EMPTY_RULE = BusinessRule(
    rule_id="check_students_empty",
    scope="INTRA",
    logic_fn=check_multiple_columns_empty,
    versioned_params={
        "COURSE_SETUP_V1": {"sheet_key": "students", "col_keys": ["id", "name"]}
    }
)

# CROSS rule metadata scaffold for marks-template validation.
# Keep this detached until MARKS_ENTRY_V1 blueprint is implemented and wired.
MARKS_STUDENT_ID_SYNC_RULE = BusinessRule(
    rule_id="marks_student_id_sync",
    scope="CROSS",
    logic_fn=check_cross_workbook_sync,
    versioned_params={
        "MARKS_ENTRY_V1": {
            "target_type_id": "COURSE_SETUP_V1",
            "curr_sheet": "students",
            "curr_col": "id",
            "ext_sheet": "students",
            "ext_col": "id",
        }
    }
)

# Updated to match the uploaded CSV structures exactly
COURSE_SETUP_BP = WorkbookBlueprint(
    type_id="COURSE_SETUP_V1",
    style_registry=SETUP_STYLE_REGISTRY,
    business_rules=[DIRECT_TOOLS_SUM_WEIGHT_RULE, INDIRECT_TOOLS_SUM_WEIGHT_RULE, CHECK_SETUP_STUDENTS_EMPTY_RULE],
    key_map={
        "assess": "Assessment_Config",
        "assess.w": "Weight (%)",
        "students": "Students",
        "students.id": "Reg_No",
        "students.name": "Student_Name",
        "assess.direct": "Direct"
    },
    sheets=[
        # --- Sheet 1: Metadata ---
        SheetSchema(
            name="Course_Metadata",
            header_matrix=[["Field", "Value"]],
            is_protected=False,
            header_style_key='header',
            data_style_key='body'
        ),
        # --- Sheet 2: Assessment Config ---
        # Note: 'Breakup' changed to 'CO_Wise_Marks_Breakup' to match your file
        SheetSchema(
            name="Assessment_Config",
            header_matrix=[["Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct"]],
            is_protected=False,
            validations=[
                ValidationRule(
                    first_row=1, first_col=2, last_row=100, last_col=4,
                    options={
                        'validate': 'list',
                        'source': ['YES', 'NO'], # Matches the uppercase in your CSV
                        'input_title': 'Requirement',
                        'input_message': 'Select YES or NO'
                    }
                )
            ]
        ),
        # --- Sheet 3: Question Mapping ---
        # Note: 'Identifier' -> 'Q_No/Rubric_Parameter', 'Max Marks' -> 'Max_Marks', 'CO Map' -> 'CO'
        SheetSchema(
            name="Question_Map",
            header_matrix=[["Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO"]],
            validations=[
                ValidationRule(
                    first_row=1, first_col=2, last_row=500, last_col=2,
                    options={
                        'validate': 'decimal',
                        'criteria': 'greater than',
                        'value': 0,
                        'error_message': 'Max marks must be greater than 0'
                    }
                )
            ]
        ),
        # --- Sheet 4: Students (Added to match your upload) ---
        SheetSchema(
            name="Students",
            header_matrix=[["Reg_No", "Student_Name"]],
            is_protected=False,
            header_style_key='header',
            data_style_key='body'
        )
    ]
)

BLUEPRINT_REGISTRY = {"COURSE_SETUP_V1": COURSE_SETUP_BP, "MARKS_ENTRY_V1": None}
