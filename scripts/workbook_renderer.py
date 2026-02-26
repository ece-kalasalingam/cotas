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
    A uniform engine that uses Blueprints to generate secure, styled Excel files.
    """
    def __init__(self):
        self.column_widths: Dict[str, Dict[int, int]] = {}
        self.header_accumulator: List[str] = []

    def _track_width(self, sheet_name: str, col: int, val: Any):
        """Standardizes visual width tracking for auto-fit."""
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
        self.header_accumulator = [] # Reset for this specific run
        
        temp_path = AtomicWorkbookWriter.create_temp_path(output_path)
        
        try:
            workbook = xlsxwriter.Workbook(temp_path)
            
            # 1. Uniform Style Registration
            formats = {
                key: workbook.add_format(props) 
                for key, props in blueprint.style_registry.items()
            }

            for sheet_schema in blueprint.sheets:
                ws = workbook.add_worksheet(sheet_schema.name)
                self.column_widths[sheet_schema.name] = {}

                # 2. Uniform Header Processing
                h_style = formats.get(sheet_schema.header_style_key)
                for r_idx, row_data in enumerate(sheet_schema.header_matrix):
                    ws.write_row(r_idx, 0, row_data, h_style)
                    for c_idx, val in enumerate(row_data):
                        self._track_width(sheet_schema.name, c_idx, val)
                        self.header_accumulator.append(str(val))

                # 3. Uniform Data Injection
                start_row = len(sheet_schema.header_matrix)
                rows = data_map.get(sheet_schema.name, [])
                d_style = formats.get(sheet_schema.data_style_key)

                for r_idx, row_data in enumerate(rows, start=start_row):
                    for c_idx, val in enumerate(row_data):
                        ws.write(r_idx, c_idx, val, d_style)
                        self._track_width(sheet_schema.name, c_idx, val)

                # 4. Uniform Cell Validation
                for rule in sheet_schema.validations:
                    # Clean internal keys like '_set_cache'
                    clean_options = {k: v for k, v in rule.options.items() if not k.startswith('_')}
                    ws.data_validation(
                        rule.first_row, rule.first_col, 
                        rule.last_row, rule.last_col, 
                        clean_options
                    )

                # 5. Uniform Protection & Formatting
                if sheet_schema.freeze_panes:
                    ws.freeze_panes(*sheet_schema.freeze_panes)
                
                if sheet_schema.is_protected:
                    ws.protect(WORKBOOK_PASSWORD, {'select_unlocked_cells': True})
                
                for col, width in self.column_widths[sheet_schema.name].items():
                    ws.set_column(col, col, min(width, DEFAULT_MAX_COL_WIDTH))

            # 6. Uniform Identity & Integrity Fingerprinting
            if fingerprint_context:
                # We pass 'blueprint' explicitly now to maintain uniformity
                self._embed_system_hash(workbook, fingerprint_context, blueprint)

            AtomicWorkbookWriter.finalize(workbook, output_path)
            return output_path

        except Exception as e:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            if isinstance(e, xlsxwriter.exceptions.FileCreateError):
                raise ValidationError(f"File {output_path} is locked. Close it and retry.")
            raise SystemError(f"Rendering Failed: {str(e)}")

    def _embed_system_hash(self, workbook: xlsxwriter.Workbook, context: Dict[str, Any], blueprint: WorkbookBlueprint):
        """Generates a hidden identity and integrity fingerprint."""
        try:
            hash_ws = workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
            
            # Row 0: Identity (From Blueprint or Context)
            type_id = context.get('type_id', blueprint.type_id)
            hash_ws.write(0, 0, type_id)

            # Row 1: Integrity (Includes Sheet Names + Headers)
            sheet_names = [s.name for s in blueprint.sheets]
            
            fingerprint = utils.generate_system_fingerprint(
                type_id=type_id,
                sheet_names=sheet_names,
                metadata=context.get('metadata') or {}, 
                headers=self.header_accumulator
            )
            print(f"type_id={type_id}, sheet_names={sheet_names}, metadata={context.get('metadata') or {}}, headers={self.header_accumulator}")
            hash_ws.write(1, 0, fingerprint)
            hash_ws.hide()
            
        except Exception as e:
            raise SystemError(f"Integrity hash embedding failed: {str(e)}")