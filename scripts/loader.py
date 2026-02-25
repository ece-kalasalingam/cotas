import pandas as pd
from scripts.models import (
    CourseMetadata,
    AssessmentConfig,
    StudentList,
    QuestionMap,
)


class SetupLoader:
    """
    Excel setup reader.
    I/O only. No validation logic.
    """

    def __init__(self, filepath: str):
        self.filepath = filepath
        try:
            self._excel = pd.ExcelFile(filepath)
        except Exception as e:
            raise RuntimeError(f"Failed to load Excel file: {e}")

    # =====================================================
    # LOADERS
    # =====================================================

    def load_metadata(self) -> CourseMetadata:
        df = pd.read_excel(self._excel, sheet_name="Course_Metadata")

        meta = dict(zip(df["Field"], df["Value"]))

        return CourseMetadata(
            course_code=str(meta.get("Course_Code", "")).strip(),
            course_name=str(meta.get("Course_Name", "")).strip(),
            section=str(meta.get("Section", "")).strip(),
            semester=str(meta.get("Semester", "")).strip(),
            academic_year=str(meta.get("Academic_Year", "")).strip(),
            faculty_name=str(meta.get("Faculty_Name", "")).strip(),
            total_cos=int(meta.get("Total_Outcomes", 0)),
        )

    def load_config(self) -> AssessmentConfig:
        df = pd.read_excel(self._excel, sheet_name="Assessment_Config")
        return AssessmentConfig(df)

    def load_students(self) -> StudentList:
        df = pd.read_excel(self._excel, sheet_name="Students")
        return StudentList(df)

    def load_question_map(self) -> QuestionMap:
        df = pd.read_excel(self._excel, sheet_name="Question_Map")
        return QuestionMap(df)