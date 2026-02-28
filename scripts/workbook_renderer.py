from typing import Any, Dict

import xlsxwriter

from scripts.constants import (
    DEFAULT_MAX_COL_WIDTH,
    DEFAULT_MIN_COL_WIDTH,
    SYSTEM_HASH_SHEET_NAME,
    WIDTH_PADDING,
    WORKBOOK_PASSWORD,
)
from scripts.exceptions import ValidationError, SystemError
from scripts.sheet_schema import WorkbookBlueprint
from scripts.utils import calculate_visual_width


class UniversalWorkbookRenderer:
    def __init__(self):
        # Retained for compatibility with any external introspection.
        self.column_widths: dict[str, dict[int, int]] = {}

    @staticmethod
    def _calc_width(val: Any) -> int:
        return calculate_visual_width(val) + WIDTH_PADDING

    def render(
        self,
        blueprint: WorkbookBlueprint,
        output_path: str,
        data_map: Dict[str, list],
        fingerprint_context: Dict[str, Any] | None = None,
    ) -> str:
        if not blueprint:
            raise SystemError("Blueprint is required.")

        workbook = None
        try:
            workbook = xlsxwriter.Workbook(output_path)
            formats = {k: workbook.add_format(v) for k, v in blueprint.style_registry.items()}

            for sheet_schema in blueprint.sheets:
                ws = workbook.add_worksheet(sheet_schema.name)
                col_widths: dict[int, int] = {}
                self.column_widths[sheet_schema.name] = col_widths

                h_style = formats.get(sheet_schema.header_style_key)
                d_style = formats.get(sheet_schema.data_style_key)
                write = ws.write

                # Column-wise width scan from header rows
                for r_idx, row_data in enumerate(sheet_schema.header_matrix):
                    for c_idx, val in enumerate(row_data):
                        write(r_idx, c_idx, val, h_style)
                        w = self._calc_width(val)
                        if w > col_widths.get(c_idx, 0):
                            col_widths[c_idx] = w

                # Column-wise width scan from data rows
                start_row = len(sheet_schema.header_matrix)
                for r_off, row_data in enumerate(data_map.get(sheet_schema.name, [])):
                    for c_idx, val in enumerate(row_data):
                        write(start_row + r_off, c_idx, val, d_style)
                        w = self._calc_width(val)
                        if w > col_widths.get(c_idx, 0):
                            col_widths[c_idx] = w

                for rule in sheet_schema.validations:
                    ws.data_validation(
                        rule.first_row,
                        rule.first_col,
                        rule.last_row,
                        rule.last_col,
                        rule.options,
                    )

                if sheet_schema.freeze_panes:
                    ws.freeze_panes(*sheet_schema.freeze_panes)

                if sheet_schema.is_protected:
                    ws.protect(WORKBOOK_PASSWORD)

                # Apply width per column (independent adjustment)
                for col, width in col_widths.items():
                    ws.set_column(
                        col,
                        col,
                        min(max(width, DEFAULT_MIN_COL_WIDTH), DEFAULT_MAX_COL_WIDTH),
                    )

            self._embed_system_hash(workbook, blueprint.type_id, fingerprint_context or {})
            workbook.close()
            return output_path
        except Exception as exc:
            if workbook is not None:
                try:
                    workbook.close()
                except Exception:
                    pass
            raise ValidationError(f"Workbook render failed: {exc}") from exc

    def _embed_system_hash(self, workbook: xlsxwriter.Workbook, type_id: str, context: Dict[str, Any]):
        ws = workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
        ws.write(0, 0, type_id)
        ws.write(0, 1, str(context.get("type_id", type_id)))
        ws.hide()
