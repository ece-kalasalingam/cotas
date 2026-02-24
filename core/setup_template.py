import xlsxwriter
import xlsxwriter.exceptions
class SetupTemplateBuilder:
    """
    Builds the initial Course Setup Excel template.
    Designed for use in a Windows Desktop Environment.
    """

    def __init__(self):
        # Define reused styles here if needed
        pass

    def build(self, save_path: str):
        """
        Creates the workbook at the specified local path.
        """
        workbook = xlsxwriter.Workbook(save_path)

        # --- Define Shared Formats ---
        # Header: Bold, Grey Background, Centered, Thin Border
        f_header = workbook.add_format({
            'bold': True,
            'bg_color': '#F2F2F2',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        f_bold = workbook.add_format({'bold': True})
        
        # Locked format for labels
        f_locked = workbook.add_format({'locked': True})

        # =====================================================
        # 1. COURSE METADATA SHEET
        # =====================================================
        ws_meta = workbook.add_worksheet("Course_Metadata")
        ws_meta.write(0, 0, "Field", f_header)
        ws_meta.write(0, 1, "Value", f_header)

        metadata_rows = [
            ("Course_Code", "ECE101"),
            ("Course_Name", "SAMPLE COURSE"),
            ("Section", "A, B"),
            ("Semester", "III"),
            ("Academic_Year", "2026-27"),
            ("Faculty_Name", "ABCECE"),
            ("Total_Outcomes", 5),
        ]

        for r_idx, (k, v) in enumerate(metadata_rows, start=1):
            ws_meta.write(r_idx, 0, k, f_bold)
            ws_meta.write(r_idx, 1, v)

        ws_meta.set_column(0, 0, 22)
        ws_meta.set_column(1, 1, 30)

        # =====================================================
        # 2. ASSESSMENT CONFIG SHEET
        # =====================================================
        ws_config = workbook.add_worksheet("Assessment_Config")
        headers = ["Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct"]
        ws_config.write_row(0, 0, headers, f_header)

        config_data = [
            ["S1", 30, "yes", "yes", "yes"],
            ["S2", 20, "yes", "yes", "yes"],
            ["ESE", 50, "no", "no", "yes"],
            ["Survey", 100, "no", "no", "no"],
        ]

        for r_idx, row in enumerate(config_data, start=1):
            ws_config.write_row(r_idx, 0, row)

        # Data Validation: Yes/No dropdown for columns C, D, and E
        ws_config.data_validation(1, 2, 200, 4, {
            'validate': 'list',
            'source': ['yes', 'no'],
            'ignore_blank': False,
            'error_message': 'Please select "yes" or "no" from the list.',
            "error_title": "Invalid Input",
            "input_title": "Select Option",
            "input_message": 'Choose "yes" or "no".'
        })
        
        ws_config.set_column(0, 4, 20)

        # =====================================================
        # 3. STUDENTS SHEET
        # =====================================================
        ws_students = workbook.add_worksheet("Students")
        ws_students.write(0, 0, "RegNo", f_header)
        ws_students.write(0, 1, "Student_Name", f_header)
        ws_students.set_column(0, 1, 25)

        # =====================================================
        # 4. QUESTION MAP SHEET
        # =====================================================
        ws_qmap = workbook.add_worksheet("Question_Map")
        q_headers = ["Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO"]
        ws_qmap.write_row(0, 0, q_headers, f_header)

        qmap_data = [
            ["S1", "Q1", 20, 1],
            ["S1", "Q2", 20, 2],
            ["S1", "Q3", 20, 3],
            ["S1", "Q4", 20, 4],
            ["S2", "Q1", 10, 4],
            ["S2", "Q2", 10, 5],
            ["ESE", "ALL QUESTIONS", 100, "1,2,3,4,5"],
        ]

        for r_idx, row in enumerate(qmap_data, start=1):
            ws_qmap.write_row(r_idx, 0, row)

        ws_qmap.set_column(0, 3, 22)

        # Close is vital in an executable to free memory and unlock the file
        try:
            workbook.close()
        except xlsxwriter.exceptions.FileCreateError:
            # Now Pylance will recognize this attribute
            raise PermissionError("The file is open in another program. Please close it.")

        return save_path