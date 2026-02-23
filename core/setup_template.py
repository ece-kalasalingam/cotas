from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation


class SetupTemplateGenerator:
    """
    Generates a Setup Excel that is 100% compatible with:
      - SetupLoader (Course_Metadata schema)
      - SetupValidator (required columns + rules)
    """

    def __init__(self, output_path: str):
        self.output_path = output_path

    def generate(self) -> None:
        wb = Workbook()
        self._make_course_metadata(wb)
        self._make_assessment_config(wb)
        self._make_question_map(wb)
        self._make_students(wb)
        wb.save(self.output_path)

    # -------------------------
    # Course_Metadata (Loader contract)
    # -------------------------
    def _make_course_metadata(self, wb: Workbook) -> None:
        ws = wb.active
        if ws is None:
            raise ValueError("Workbook must have an active sheet")
        ws.title = "Course_Metadata"

        ws["A1"] = "Field"
        ws["B1"] = "Value"
        ws["A1"].font = Font(bold=True)
        ws["B1"].font = Font(bold=True)

        rows = [
            ("Course_Code", ""),
            ("Course_Name", ""),
            ("Section", ""),
            ("Semester", ""),
            ("Academic_Year", ""),
            ("Faculty_Name", ""),
            ("Total_Outcomes", 6),  # default; faculty edits
        ]

        r = 2
        for k, v in rows:
            ws.cell(r, 1, k)
            ws.cell(r, 2, v)
            r += 1

        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 30

    # -------------------------
    # Assessment_Config (Validator contract)
    # -------------------------
    def _make_assessment_config(self, wb: Workbook) -> None:
        ws = wb.create_sheet("Assessment_Config")

        headers = ["Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(1, c, h)
            cell.font = Font(bold=True)

        # A VALID default that passes validator immediately:
        # - at least one Direct component
        # - Direct weights sum to 100
        ws.append(["S1", 100, "yes", "yes", "yes"])

        # Yes/No validation for the 3 flag columns
        dv_yesno = DataValidation(type="list", formula1='"yes,no"', allow_blank=False)
        ws.add_data_validation(dv_yesno)
        dv_yesno.add("C2:C200")
        dv_yesno.add("D2:D200")
        dv_yesno.add("E2:E200")

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 8
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 10

    # -------------------------
    # Question_Map (Validator contract)
    # -------------------------
    def _make_question_map(self, wb: Workbook) -> None:
        ws = wb.create_sheet("Question_Map")

        headers = ["Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(1, c, h)
            cell.font = Font(bold=True)

        # Must cover CO1..CO6 (because Total_Outcomes default = 6)
        # CO_Wise_Marks_Breakup = yes => exactly one CO per row.
        ws.append(["S1", "Q1", 10, "1"])
        ws.append(["S1", "Q2", 10, "2"])
        ws.append(["S1", "Q3", 10, "3"])
        ws.append(["S1", "Q4", 10, "4"])
        ws.append(["S1", "Q5", 10, "5"])
        ws.append(["S1", "Q6", 10, "6"])

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 12

    # -------------------------
    # Students (Validator contract)
    # -------------------------
    def _make_students(self, wb: Workbook) -> None:
        ws = wb.create_sheet("Students")

        headers = ["RegNo", "Student_Name"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(1, c, h)
            cell.font = Font(bold=True)

        # IMPORTANT: no blank rows (avoid NaN -> "nan")
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 26