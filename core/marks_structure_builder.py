from core.models import ValidatedSetup, CourseMetadata


class MarksTemplateStructureBuilder:
    """
    PURE structure builder.
    No openpyxl.
    Fully unit-testable.
    """

    def __init__(self, metadata: CourseMetadata, setup: ValidatedSetup):
        self.metadata = metadata
        self.setup = setup

    # --------------------------------------------------

    def build(self) -> dict:
        structure = {
            "course_info": self._build_course_info(),
            "components": self._build_components(),
            "indirect": self._build_indirect(),
        }
        return structure

    # --------------------------------------------------

    def _build_course_info(self):
        return {
            "headers": ["Field", "Value"],
            "rows": [
                [key, value]
                for key, value in vars(self.metadata).items()
            ],
        }

    # --------------------------------------------------

    def _build_components(self):
        students = [
            (s.reg_no, s.name)
            for s in self.setup.students
        ]

        components = {}

        for comp_name, comp in self.setup.components.items():

            if not comp.direct:
                continue

            question_ids = [q.identifier for q in comp.questions]

            headers = (
                ["RegNo", "Student_Name"]
                + question_ids
                + ["Total"]
            )

            co_row = (
                ["CO", ""]
                + [",".join(q.co_list) for q in comp.questions]
                + [""]
            )

            max_row = (
                ["Max", ""]
                + [q.max_marks for q in comp.questions]
                + [""]
            )

            student_rows = [
                [regno, name] + [""] * (len(question_ids) + 1)
                for regno, name in students
            ]

            components[comp_name] = {
                "headers": headers,
                "co_row": co_row,
                "max_row": max_row,
                "students": student_rows,
                "questions": comp.questions,
            }

        return components

    # --------------------------------------------------

    def _build_indirect(self):
        total_cos = self.metadata.total_cos

        headers = (
            ["RegNo", "Student_Name"]
            + [f"CO{i}" for i in range(1, total_cos + 1)]
        )

        students = [
            (s.reg_no, s.name)
            for s in self.setup.students
        ]

        indirect = {}

        for tool in self.setup.indirect_tools:
            student_rows = [
                [regno, name] + [""] * total_cos
                for regno, name in students
            ]

            indirect[tool.name] = {
                "headers": headers,
                "students": student_rows,
            }

        return indirect