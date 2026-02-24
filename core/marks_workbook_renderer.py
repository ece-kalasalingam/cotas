import hashlib
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
from core.constants import (
    WORKBOOK_PASSWORD,
    LIKERT_MIN,
    LIKERT_MAX,
    ABSENT_SYMBOL,
    DEFAULT_NUMBER_FORMAT,
)
from core.models import ValidatedSetup, CourseMetadata

class MarksTemplateWorkbookRenderer:
    def __init__(self, metadata: CourseMetadata, setup: ValidatedSetup):
        self.metadata = metadata
        self.setup = setup
        # Dictionary to track column widths per sheet
        self.column_widths = {}

    def _update_max_width(self, sheet_name, col, value):
        """Tracks the maximum character length for a column."""
        if value is None:
            return
        
        # Approximate width: length of string + a small buffer
        width = len(str(value)) + 2
        
        if sheet_name not in self.column_widths:
            self.column_widths[sheet_name] = {}
        
        self.column_widths[sheet_name][col] = max(
            self.column_widths[sheet_name].get(col, 0), width
        )

    def _apply_autofit(self, ws):
        """Applies the tracked widths to the worksheet."""
        sheet_name = ws.name
        if sheet_name in self.column_widths:
            for col, width in self.column_widths[sheet_name].items():
                # Constrain width between 10 and 50 for readability
                final_width = min(max(width, 10), 50)
                ws.set_column(col, col, final_width)

    def render(self, structure: dict, output_path: str):
        workbook = xlsxwriter.Workbook(output_path)
        
        # --- Shared Formats ---
        self.f_header = workbook.add_format({
            'bold': True, 'bg_color': '#F2F2F2', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        self.f_locked_grey = workbook.add_format({
            'bg_color': '#F2F2F2', 'border': 1, 'locked': True
        })
        self.f_unlocked = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'locked': False, 'num_format': DEFAULT_NUMBER_FORMAT
        })
        self.f_formula = workbook.add_format({
            'bg_color': '#F2F2F2', 'border': 1, 'align': 'center',
            'valign': 'vcenter', 'locked': True, 'num_format': DEFAULT_NUMBER_FORMAT
        })

        self._render_course_info(workbook, structure["course_info"])
        self._render_components(workbook, structure["components"])
        self._render_indirect(workbook, structure["indirect"])
        self._embed_system_hash(workbook)

        workbook.close()
        return output_path

    # --------------------------------------------------

    def _apply_layout(self, ws):
        ws.set_paper(9)  # A4
        ws.set_margins(0.25, 0.25, 0.5, 0.5)
        ws.set_header(f'&LCourse: {self.metadata.course_code}&R{self.metadata.academic_year}')
        ws.set_footer('&CPage &P of &N')

    def _render_course_info(self, workbook, data):
        ws = workbook.add_worksheet("Course_Info")
        self._apply_layout(ws)
        
        # Headers
        for c, h in enumerate(data["headers"]):
            ws.write(0, c, h, self.f_header)
            self._update_max_width(ws.name, c, h)

        # Rows
        for r, row_data in enumerate(data["rows"], 1):
            for c, val in enumerate(row_data):
                ws.write(r, c, val)
                self._update_max_width(ws.name, c, val)

        ws.protect(WORKBOOK_PASSWORD)
        self._apply_autofit(ws)

    def _render_components(self, workbook, components):
        for name, comp in components.items():
            ws = workbook.add_worksheet(name)
            self._apply_layout(ws)
            ws.freeze_panes(3, 2)

            # Header Rows 1-3
            for r_idx, key in enumerate(["headers", "co_row", "max_row"]):
                for c_idx, val in enumerate(comp[key]):
                    ws.write(r_idx, c_idx, val, self.f_header)
                    self._update_max_width(ws.name, c_idx, val)

            questions = comp["questions"]
            end_col = 2 + len(questions)
            last_student_row_idx = 2 + len(comp["students"])

            # Student Rows
            for r_idx, student in enumerate(comp["students"], 3):
                # RegNo & Name
                ws.write(r_idx, 0, student[0], self.f_locked_grey)
                ws.write(r_idx, 1, student[1], self.f_locked_grey)
                self._update_max_width(ws.name, 0, student[0])
                self._update_max_width(ws.name, 1, student[1])

                # Marks Input
                for c_idx in range(2, end_col):
                    val = student[c_idx] if c_idx < len(student) else None
                    ws.write(r_idx, c_idx, val, self.f_unlocked)

                # Total Formula
                col_start = xl_col_to_name(2)
                col_end = xl_col_to_name(end_col - 1)
                ws.write_formula(r_idx, end_col, f"=SUM({col_start}{r_idx+1}:{col_end}{r_idx+1})", self.f_formula)

            # Validation
            for c_idx in range(2, end_col):
                max_mark = questions[c_idx-2].max_marks
                col_let = xl_col_to_name(c_idx)
                ws.data_validation(3, c_idx, last_student_row_idx, c_idx, {
                    'validate': 'custom',
                    'value': f'=OR(AND(ISNUMBER({col_let}4),{col_let}4>=0,{col_let}4<={max_mark}),UPPER({col_let}4)="A")',
                    "input_title": "Marks entry rules",
                    "input_message": f" between 0 and {max_mark}, or 'A' for Absent.",
                    "error_title": "Invalid Marks",
                    "error_message": f"Please enter a number between 0 and {max_mark}, or 'A' for Absent."
                })

            ws.protect(WORKBOOK_PASSWORD)
            self._apply_autofit(ws)

    def _render_indirect(self, workbook, indirect):
        for tool_name, data in indirect.items():
            ws = workbook.add_worksheet(f"{tool_name}_Indirect"[:31])
            self._apply_layout(ws)
            ws.freeze_panes(1, 2)

            # Headers
            for c, h in enumerate(data["headers"]):
                ws.write(0, c, h, self.f_header)
                self._update_max_width(ws.name, c, h)

            # Rows
            for r_idx, student in enumerate(data["students"], 1):
                ws.write(r_idx, 0, student[0], self.f_locked_grey)
                ws.write(r_idx, 1, student[1], self.f_locked_grey)
                
                for c_idx in range(2, len(data["headers"])):
                    val = student[c_idx] if c_idx < len(student) else None
                    ws.write(r_idx, c_idx, val, self.f_unlocked)

            # Validation (Likert)
            for c_idx in range(2, len(data["headers"])):
                col_let = xl_col_to_name(c_idx)
                ws.data_validation(1, c_idx, len(data["students"]), c_idx, {
                    'validate': 'custom',
                    'value': f'=OR(AND(ISNUMBER({col_let}2),{col_let}2>={LIKERT_MIN},{col_let}2<={LIKERT_MAX}),UPPER({col_let}2)="{ABSENT_SYMBOL}")',
                    "input_title": "Likert Scale Entry Rules",
                    "input_message": f"Enter a number between {LIKERT_MIN} and {LIKERT_MAX}, or '{ABSENT_SYMBOL}' for Absent.",
                    "error_title": "Invalid Entry",
                    "error_message": f"Please enter a number between {LIKERT_MIN} and {LIKERT_MAX}, or '{ABSENT_SYMBOL}' for Absent."
                })

            ws.protect(WORKBOOK_PASSWORD)
            self._apply_autofit(ws)

    def _embed_system_hash(self, workbook):
        hasher = hashlib.sha256()

        # 1. Components - Sort by name
        for comp_name, comp in sorted(self.setup.components.items()):
            hasher.update(comp_name.encode())

            # 2. Questions
            for q in comp.questions:
                hasher.update(q.identifier.encode())
                
                # FORCE Consistency: convert to float then string (prevents 20 vs 20.0)
                hasher.update(str(float(q.max_marks)).encode())
                
                # Sort COs so "1,2" and "2,1" produce the same hash
                co_str = ",".join(sorted([str(x).strip() for x in q.co_list]))
                hasher.update(co_str.encode())

        # 3. Students - Sort by RegNo
        for student in sorted(self.setup.students, key=lambda x: x.reg_no):
            hasher.update(student.reg_no.encode())

        fingerprint = hasher.hexdigest()
        
        sheet = workbook.add_worksheet("__SYSTEM_HASH__")
        sheet.write(0, 0, fingerprint)
        sheet.hide()