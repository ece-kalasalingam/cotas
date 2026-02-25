from scripts.sheet_schema import StyleDefinition, WorkbookBlueprint, SheetSchema, ValidationRule
from scripts.constants import ABSENT_SYMBOL
from typing import Dict

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

# Updated to match the uploaded CSV structures exactly
COURSE_SETUP_BP = WorkbookBlueprint(
    type_id="COURSE_SETUP_V1",
    style_registry=SETUP_STYLE_REGISTRY,
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