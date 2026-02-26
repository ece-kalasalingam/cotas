import openpyxl
from typing import Dict, List, Any
from scripts import utils
from scripts.constants import SYSTEM_HASH_SHEET_NAME, METADATA_SHEET_NAME

class UniversalEngine:
    def __init__(self, registry: Dict[str, Any]):
        self.registry = registry
        self.bp = None
        self.errors = []
        self.data_store: Dict[str, List[List[Any]]] = {}
        self._col_cache: Dict[str, Dict[str, int]] = {}

    def load_from_file(self, filepath: str) -> bool:
        self.errors = []
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            
            # THE GATEKEEPER
            if not self._verify_integrity(wb):
                wb.close()
                return False
            if not self.bp:
                self.errors.append("System Error: Blueprint not initialized.")
                wb.close()
                return False
            
            # Load sheets based on the BP identified during verification
            for sheet_schema in self.bp.sheets:
                if sheet_schema.name not in wb.sheetnames:
                    self.errors.append(f"Missing required sheet: {sheet_schema.name}")
                    continue
                
                ws = wb[sheet_schema.name]
                all_rows = [list(row) for row in ws.values]
                
                header_height = len(sheet_schema.header_matrix)
                blueprint_headers = sheet_schema.header_matrix[-1]
                self._col_cache[sheet_schema.name] = {
                    utils.normalize(h): idx for idx, h in enumerate(blueprint_headers)
                }
                self.data_store[sheet_schema.name] = all_rows[header_height:]

            wb.close()

            for rule in self.bp.business_rules:
                self.errors.extend(rule.run(self, ext_engine=None))

            return len(self.errors) == 0
        except Exception as e:
            self.errors.append(f"Load Error: {str(e)}")
            return False

    def _verify_integrity(self, wb: openpyxl.Workbook) -> bool:
        """Extracts and verifies the fingerprint."""
        if SYSTEM_HASH_SHEET_NAME not in wb.sheetnames:
            self.errors.append("Security Error: Invalid template (No Signature).")
            return False

        h_ws = wb[SYSTEM_HASH_SHEET_NAME]
        file_type_id = str(h_ws.cell(1, 1).value).strip()
        embedded_hash = str(h_ws.cell(2, 1).value).strip()

        # Step 1: Resolve Blueprint version
        self.bp = self.registry.get(file_type_id)
        if not self.bp:
            self.errors.append(f"Unknown Version: {file_type_id}")
            return False

        # Step 2: Recalculate Hash
        meta = self._extract_metadata(wb)
        content = self._extract_structure_stream(wb)
        
        calculated_hash = utils.generate_system_fingerprint(
            type_id=file_type_id,
            metadata=meta,
            sheet_names=wb.sheetnames, # Pass all sheet names for structural check
            content_stream=content
        )

        if calculated_hash != embedded_hash:
            self.errors.append("Security Alert: Course info or sheet headers have been altered.")
            return False
        
        return True

    def _extract_metadata(self, wb) -> Dict[str, Any]:
        meta = {}
        if METADATA_SHEET_NAME in wb.sheetnames:
            ws = wb[METADATA_SHEET_NAME]
            for row in ws.iter_rows(values_only=True):
                if row[0]: meta[str(row[0])] = row[1]
        return meta

    def _extract_structure_stream(self, wb) -> List[Any]:
        stream = []
        for name in wb.sheetnames:
            if name == SYSTEM_HASH_SHEET_NAME: continue
            ws = wb[name]
            # Capture headers to detect column tampering
            for row in ws.iter_rows(max_row=2, values_only=True):
                stream.extend([c for c in row if c is not None])
        return stream

    def get_col_idx(self, sheet_name: str, col_name: str) -> int:
        return self._col_cache.get(sheet_name, {}).get(utils.normalize(col_name), -1)