import os

import xlsxwriter
import xlsxwriter.exceptions
from typing import List, Dict, Any, cast
from scripts.atomic_workbook_writer import AtomicWorkbookWriter
from scripts.sheet_schema import WorkbookBlueprint
from scripts.exceptions import ValidationError, SystemError
from scripts.constants import WORKBOOK_PASSWORD, DEFAULT_MAX_COL_WIDTH, SYSTEM_HASH_SHEET_NAME
from scripts import utils

class UniversalWorkbookRenderer:
    def __init__(self, setup: Any = None, metadata: Any = None):
        self.setup = setup
        self.metadata = metadata
        self.column_widths: Dict[str, Dict[int, int]] = {}
        self.header_accumulator: List[str] = []

    def _track_width(self, sheet_name: str, col: int, val: Any):
        """Uses utils to calculate visual width and updates tracking."""
        visual_width = utils.calculate_visual_width(val)
        self.column_widths.setdefault(sheet_name, {})[col] = max(
            self.column_widths[sheet_name].get(col, 0), visual_width
        )

    def render(self, blueprint: WorkbookBlueprint, output_path: str, data_map: Dict[str, List[List[Any]]]):
        blueprint.validate_structure()

        temp_path = AtomicWorkbookWriter.create_temp_path(output_path)
        try:
            workbook = xlsxwriter.Workbook(temp_path)
            
            # 1. Dynamic Style Generation
            formats = {
                key: workbook.add_format(props) 
                for key, props in blueprint.style_registry.items()
            }

            for sheet_schema in blueprint.sheets:
                ws = workbook.add_worksheet(sheet_schema.name)
                self.column_widths[sheet_schema.name] = {}

                # 2. Write Headers
                h_style = formats.get(sheet_schema.header_style_key)
                if not h_style:
                    raise SystemError(f"Style '{sheet_schema.header_style_key}' not found.")

                for r_idx, row_data in enumerate(sheet_schema.header_matrix):
                    ws.write_row(r_idx, 0, row_data, h_style)
                    for c_idx, val in enumerate(row_data):
                        self._track_width(sheet_schema.name, c_idx, val)
                        self.header_accumulator.append(str(val))

                # 3. Write Data
                d_style = formats.get(sheet_schema.data_style_key)
                header_height = len(sheet_schema.header_matrix)
                rows = data_map.get(sheet_schema.name, [])

                for r_idx, row_data in enumerate(rows, start=header_height):
                    for c_idx, val in enumerate(row_data):
                        ws.write(r_idx, c_idx, val, d_style)
                        self._track_width(sheet_schema.name, c_idx, val)

                # 4. Validations
                for rule in sheet_schema.validations:
                    ws.data_validation(
                        rule.first_row, rule.first_col, 
                        rule.last_row, rule.last_col, 
                        dict(rule.options)
                    )

                # 5. Sheet Config using Constants
                if sheet_schema.freeze_panes:
                    ws.freeze_panes(*sheet_schema.freeze_panes)
                
                if sheet_schema.is_protected:
                    ws.protect(WORKBOOK_PASSWORD, {'select_unlocked_cells': True})
                
                # Apply Auto-fit with Constant Limit
                for col, width in self.column_widths[sheet_schema.name].items():
                    ws.set_column(col, col, min(width, DEFAULT_MAX_COL_WIDTH))

            current_weights = cast(Dict[str, Any], data_map.get("weightages", {}))
            self._embed_system_hash(workbook, current_weights)
            AtomicWorkbookWriter.finalize(workbook, output_path)
            return output_path

        except xlsxwriter.exceptions.FileCreateError:
            raise ValidationError(f"File {output_path} is currently locked. Close it and retry.")
        except Exception as e:
            if os.path.exists(temp_path):
                try:
                    # We try to close if it's still open, but usually 
                    # just removing the file is enough.
                    os.remove(temp_path)
                except:
                    pass
            raise SystemError(f"Rendering Failed: {str(e)}")

    def _embed_system_hash(self, workbook: xlsxwriter.Workbook, weightages: Dict[str, Any]):
        """Generates a fingerprint of Metadata + Headers + Weightages."""
        try:
            hash_ws = workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
            
            # Now we use the version of the function that accepts headers/weights
            fingerprint = utils.generate_system_fingerprint(
                self.setup, 
                self.metadata, 
                self.header_accumulator,
                weightages
            )
            
            hash_ws.write(0, 0, fingerprint)
            hash_ws.hide()
        except Exception:
            raise SystemError("Integrity hash embedding failed.")