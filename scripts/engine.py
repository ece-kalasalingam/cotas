import openpyxl
from typing import Dict, List, Any, Optional

from scripts import utils
from scripts.constants import SYSTEM_HASH_SHEET_NAME, METADATA_SHEET_NAME


class UniversalEngine:
    def __init__(self, registry: Dict[str, Any]):
        # Canonicalize registry keys once
        self.registry = {
            utils.canonicalize(k): v for k, v in registry.items()
        }

        self.bp: Optional[Any] = None
        self.errors: List[str] = []
        self.data_store: Dict[str, List[List[Any]]] = {}
        self._col_cache: Dict[str, Dict[str, int]] = {}

    # =====================================================
    # PUBLIC ENTRY
    # =====================================================

    def load_from_file(self, filepath: str) -> bool:
        self.errors = []
        self.data_store = {}
        self._col_cache = {}

        try:
            wb = openpyxl.load_workbook(
                filepath,
                data_only=True,
                read_only=True
            )

            # ---------------- GATEKEEPER ----------------
            if not self._verify_integrity(wb):
                wb.close()
                return False

            assert self.bp is not None

            # ---------------- LOAD DATA ----------------
            for sheet_schema in self.bp.sheets:
                if sheet_schema.name not in wb.sheetnames:
                    self.errors.append(
                        f"Missing required sheet: {sheet_schema.name}"
                    )
                    continue

                ws = wb[sheet_schema.name]
                all_rows_iter = ws.values

                header_height = len(sheet_schema.header_matrix)

                # Column cache from blueprint last header row
                blueprint_headers = (
                    sheet_schema.header_matrix[-1]
                    if header_height > 0 else []
                )

                self._col_cache[sheet_schema.name] = {
                    utils.normalize(h): idx
                    for idx, h in enumerate(blueprint_headers)
                    if h is not None
                }

                # Skip header rows
                for _ in range(header_height):
                    next(all_rows_iter, None)

                # Store remaining rows
                self.data_store[sheet_schema.name] = [
                    list(row) for row in all_rows_iter
                ]

            wb.close()

            # ---------------- BUSINESS RULES ----------------
            for rule in self.bp.business_rules:
                self.errors.extend(rule.run(self, external_engine=None))

            return len(self.errors) == 0

        except Exception as e:
            self.errors.append(f"Load Error: {str(e)}")
            return False

    # =====================================================
    # INTEGRITY VERIFICATION
    # =====================================================

    def _verify_integrity(self, wb: openpyxl.Workbook) -> bool:
        """Performs strict structural validation against Blueprint."""

        # ---- 1. System Sheet Exists ----
        if SYSTEM_HASH_SHEET_NAME not in wb.sheetnames:
            self.errors.append("Invalid File: Not the standard template.")
            return False

        # ---- 2. Identity Check ----
        raw_type_id = wb[SYSTEM_HASH_SHEET_NAME]["A1"].value
        type_id = utils.canonicalize(raw_type_id)

        if type_id not in self.registry:
            self.errors.append(f"Unknown Template ID: {raw_type_id}")
            return False

        self.bp = self.registry[type_id]
        if not self.bp:
            self.errors.append(
                f"Registry Error: No blueprint found for ID {raw_type_id}"
            )
            return False

        # ---- 3. Unexpected Sheet Detection ----
        expected_sheets = {s.name for s in self.bp.sheets}
        actual_sheets = set(wb.sheetnames)

        allowed_extras = {SYSTEM_HASH_SHEET_NAME, METADATA_SHEET_NAME}

        unexpected = actual_sheets - expected_sheets - allowed_extras
        if unexpected:
            self.errors.append(
                f"Unexpected sheet(s) found: {', '.join(unexpected)}"
            )
            return False

        # ---- 4. Header Structural Match ----
        for sheet_bp in self.bp.sheets:
            if sheet_bp.name not in wb.sheetnames:
                self.errors.append(f"Missing sheet: {sheet_bp.name}")
                return False

            ws = wb[sheet_bp.name]
            header_height = len(sheet_bp.header_matrix)

            if header_height <= 0:
                continue

            # Determine expected width
            max_cols = max(
                (len(row) for row in sheet_bp.header_matrix if row),
                default=0
            )

            # Read header block
            actual_rows = list(ws.iter_rows(
                min_row=1,
                max_row=header_height,
                min_col=1,
                max_col=max_cols,
                values_only=True
            ))

            # Compare cell-by-cell using canonicalization
            for r_idx in range(header_height):
                expected_row = sheet_bp.header_matrix[r_idx]
                actual_row = (
                    actual_rows[r_idx]
                    if r_idx < len(actual_rows) else ()
                )

                for c_idx in range(max_cols):
                    expected_val = (
                        expected_row[c_idx]
                        if c_idx < len(expected_row) else None
                    )
                    actual_val = (
                        actual_row[c_idx]
                        if c_idx < len(actual_row) else None
                    )

                    if (
                        utils.canonicalize(expected_val)
                        != utils.canonicalize(actual_val)
                    ):
                        self.errors.append(
                            f"Header mismatch in '{sheet_bp.name}' "
                            f"at row {r_idx+1}, column {c_idx+1}"
                        )
                        return False

            # ---- Strict Extra Column Detection ----
            if ws.max_column > max_cols:
                for r in range(1, header_height + 1):
                    for c in range(max_cols + 1, ws.max_column + 1):
                        val = ws.cell(row=r, column=c).value
                        if utils.canonicalize(val) != "empty":
                            self.errors.append(
                                f"Unexpected extra header column "
                                f"in '{sheet_bp.name}' "
                                f"at row {r}, column {c}"
                            )
                            return False

        return True

    # =====================================================
    # METADATA + HEADER EXTRACTION
    # =====================================================

    def _extract_metadata(
        self, wb: openpyxl.Workbook
    ) -> Dict[str, Any]:

        meta: Dict[str, Any] = {}

        if METADATA_SHEET_NAME in wb.sheetnames:
            ws = wb[METADATA_SHEET_NAME]
            for row in ws.iter_rows(values_only=True):
                if row and row[0]:
                    meta[str(row[0])] = row[1]

        return meta

    def _get_all_headers(
        self, wb: openpyxl.Workbook
    ) -> List[Any]:
        """
        Extract headers in deterministic blueprint order.
        Canonicalized to align with validation.
        """

        headers: List[Any] = []

        if not self.bp:
            return headers

        for sheet_schema in self.bp.sheets:
            name = sheet_schema.name
            if name not in wb.sheetnames:
                continue

            ws = wb[name]
            header_height = len(sheet_schema.header_matrix)

            if header_height <= 0:
                continue

            max_cols = max(
                (len(row) for row in sheet_schema.header_matrix if row),
                default=0
            )

            for row in ws.iter_rows(
                min_row=1,
                max_row=header_height,
                min_col=1,
                max_col=max_cols,
                values_only=True,
            ):
                for cell in row:
                    headers.append(utils.canonicalize(cell))

        return headers

    # =====================================================
    # COLUMN LOOKUP
    # =====================================================

    def get_col_idx(self, sheet_name: str, col_name: str) -> int:
        return self._col_cache.get(
            sheet_name,
            {}
        ).get(utils.normalize(col_name), -1)