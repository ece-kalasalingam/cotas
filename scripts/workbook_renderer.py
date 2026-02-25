import os
import xlsxwriter
import xlsxwriter.exceptions
from typing import List, Dict, Any, Optional
from scripts.atomic_workbook_writer import AtomicWorkbookWriter
from scripts.sheet_schema import WorkbookBlueprint
from scripts.exceptions import ValidationError, SystemError
from scripts.constants import WORKBOOK_PASSWORD, DEFAULT_MAX_COL_WIDTH, SYSTEM_HASH_SHEET_NAME
from scripts import utils

class UniversalWorkbookRenderer:
    """
    Unit: Technical Rendering
    Universal engine that follows Blueprints to create styled Excel files.
    """
    def __init__(self):
        self.column_widths: Dict[str, Dict[int, int]] = {}
        # Used for the system hash fingerprint
        self.header_accumulator: List[str] = []

    def _track_width(self, sheet_name: str, col: int, val: Any):
        visual_width = utils.calculate_visual_width(val)
        self.column_widths.setdefault(sheet_name, {})[col] = max(
            self.column_widths[sheet_name].get(col, 0), visual_width
        )

    def render(self, 
               blueprint: WorkbookBlueprint, 
               output_path: str, 
               data_map: Dict[str, List[List[Any]]],
               fingerprint_context: Optional[Dict[str, Any]] = None) -> str:
        
        blueprint.validate_structure()
        temp_path = AtomicWorkbookWriter.create_temp_path(output_path)
        
        try:
            workbook = xlsxwriter.Workbook(temp_path)
            
            # 1. Register Styles from Blueprint Registry
            formats = {
                key: workbook.add_format(props) 
                for key, props in blueprint.style_registry.items()
            }

            for sheet_schema in blueprint.sheets:
                ws = workbook.add_worksheet(sheet_schema.name)
                self.column_widths[sheet_schema.name] = {}

                # 2. Write Headers
                h_style = formats.get(sheet_schema.header_style_key)
                for r_idx, row_data in enumerate(sheet_schema.header_matrix):
                    ws.write_row(r_idx, 0, row_data, h_style)
                    for c_idx, val in enumerate(row_data):
                        self._track_width(sheet_schema.name, c_idx, val)
                        self.header_accumulator.append(str(val))

                # 3. Write Data (Content from Data Map)
                start_row = len(sheet_schema.header_matrix)
                rows = data_map.get(sheet_schema.name, [])
                d_style = formats.get(sheet_schema.data_style_key)

                for r_idx, row_data in enumerate(rows, start=start_row):
                    for c_idx, val in enumerate(row_data):
                        ws.write(r_idx, c_idx, val, d_style)
                        self._track_width(sheet_schema.name, c_idx, val)

                # 4. Apply Validations (Filter out internal cache keys)
                for rule in sheet_schema.validations:
                    # IMPORTANT: Remove '_set_cache' because XlsxWriter doesn't allow it
                    clean_options = {k: v for k, v in rule.options.items() if not k.startswith('_')}
                    ws.data_validation(
                        rule.first_row, rule.first_col, 
                        rule.last_row, rule.last_col, 
                        clean_options
                    )

                # 5. Sheet Configuration
                if sheet_schema.freeze_panes:
                    ws.freeze_panes(*sheet_schema.freeze_panes)
                
                if sheet_schema.is_protected:
                    ws.protect(WORKBOOK_PASSWORD, {'select_unlocked_cells': True})
                
                # Auto-fit columns
                for col, width in self.column_widths[sheet_schema.name].items():
                    ws.set_column(col, col, min(width, DEFAULT_MAX_COL_WIDTH))

            # 6. Optional System Hash (Fingerprinting)
            if fingerprint_context:
                self._embed_system_hash(workbook, fingerprint_context)

            AtomicWorkbookWriter.finalize(workbook, output_path)
            return output_path

        except Exception as e:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            if isinstance(e, xlsxwriter.exceptions.FileCreateError):
                raise ValidationError(f"File {output_path} is locked. Close it and retry.")
            raise SystemError(f"Rendering Failed: {str(e)}")

    def _embed_system_hash(self, workbook: xlsxwriter.Workbook, context: Dict[str, Any]):
        """Generates a hidden integrity fingerprint."""
        try:
            hash_ws = workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
            fingerprint = utils.generate_system_fingerprint(
                setup=context.get('setup'), 
                metadata=context.get('metadata'), 
                headers=self.header_accumulator,
                weightages=context.get('weightages', {})
            )
            hash_ws.write(0, 0, fingerprint)
            hash_ws.hide()
        except Exception:
            raise SystemError("Integrity hash embedding failed.")