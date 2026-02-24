import xlsxwriter
import pandas as pd
from pathlib import Path
from core.constants import WORKBOOK_PASSWORD

class ReportBuilder:
    """
    Pure report writer. 
    Consumes DataFrames from Calculator.run() and applies formatting.
    """

    def __init__(self, data_package: dict):
        self.pkg = data_package

    def build(self, output_path: str) -> None:
        output_path_obj = Path(output_path)
        output_path_obj.parent.mkdir(parents=True, exist_ok=True)

        # 'nan_inf_to_errors' handles any stray numpy issues
        with xlsxwriter.Workbook(str(output_path_obj), {'nan_inf_to_errors': True}) as wb:
            
            # 1. FORMATS
            header_fmt = wb.add_format({
                'bold': True,
                'bg_color': '#D3D3D3',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True  # CRITICAL: Allows \n to work in headers
            })

            cell_fmt = wb.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # 2. SHEET GENERATION
            for sheet_key, df in self.pkg["co_sheets"].items():
                # sheet_key is "1_Direct", "2_Indirect", etc.
                ws = wb.add_worksheet(sheet_key[:31]) # Excel limit

                # 3. PAGE SETUP
                ws.set_landscape()
                ws.set_paper(8) # A3
                ws.fit_to_pages(1, 0)
                ws.freeze_panes(1, 2)
                
                # Height adjustment for multi-line headers (Max/Wt/Scaled/Total%)
                ws.set_row(0, 70) 

                # 4. WRITE HEADERS
                # We do NOT use title() or replace("_") here because 
                # Calculator has already provided the exact "Max/Wt" multi-line string.
                for col_idx, col_name in enumerate(df.columns):
                    ws.write(0, col_idx, col_name, header_fmt)

                # 5. COLUMN WIDTHS
                ws.set_column(0, 0, 14) # Reg No
                ws.set_column(1, 1, 25) # Name
                if len(df.columns) > 2:
                    ws.set_column(2, len(df.columns) - 1, 15)

                # 6. WRITE DATA
                # Calculator already provided "A", "N/A", or numeric values.
                # No need for display_df.where() logic; Calculator handled the intent.
                for row_idx, row_values in enumerate(df.values):
                    for col_idx, val in enumerate(row_values):
                        # xlsxwriter is polymorphic: ws.write handles str, float, int
                        ws.write(row_idx + 1, col_idx, val, cell_fmt)

                # 7. SECURITY
                ws.protect(WORKBOOK_PASSWORD)