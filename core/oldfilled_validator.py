import pandas as pd
import numpy as np
import hashlib
from collections import Counter
from core.exceptions import ValidationError
from modules.utils import normalize, safe_int
import re


class FilledMarksValidator:
    """
    Optimized structure validator for filled Excel marks sheets.
    - O(1) sheet lookup via normalized map
    - Direct/Indirect header STRICT
    - Student list SET-based with missing/extra/duplicates reported
    - CO lists canonicalized to SORTED order (order-insensitive in sheet/setup)
    - Checksum uses sorted COs and sorted students (deterministic)
    """

    def __init__(self, validated_setup, sheets: dict):
        self.validated = validated_setup
        self.sheets = sheets

        # O(1) lookup: {normalized_name: original_key}
        self._norm_map = {normalize(k): k for k in self.sheets.keys()}
        if len(self._norm_map) != len(self.sheets):
            raise ValidationError("Workbook contains duplicate sheet names (case/space insensitive).")

    # =====================================================
    # PUBLIC ENTRY
    # =====================================================

    def validate(self):
        self._validate_sheet_names()
        self._validate_direct_structure()
        self._validate_indirect_structure()
        self._validate_checksum()

    # =====================================================
    # HELPERS
    # =====================================================

    def _get_sheet_ignore_case(self, target_name):
        target_norm = normalize(target_name)
        actual_key = self._norm_map.get(target_norm)
        if actual_key:
            return self.sheets[actual_key]
        raise ValidationError(f"Missing sheet: {target_name}")

    @staticmethod
    def _canon_cell_text(x) -> str:
        """Cell normalization for headers and identifiers (fast)."""
        if pd.isna(x):
            return ""
        if isinstance(x, (int, float, np.integer, np.floating)):
            fx = float(x)
            return str(int(fx)) if fx.is_integer() else str(fx)
        return str(x).strip()

    @staticmethod
    def _canon_co_cell_sorted(x) -> str:
        if pd.isna(x): return ""
        
        # Numeric/Float handling
        if isinstance(x, (int, float, np.integer, np.floating)):
            fx = float(x)
            return str(int(fx)) if fx.is_integer() else str(fx)

        s = str(x).strip()
        if not s: return ""

        # Delimiter normalization
        parts = re.split(r'[;,|]+', s.replace(" ", ""))
        
        cleaned = set()
        for p in parts:
            if not p: continue
            p = p.upper().replace("CO", "")
            # Handle float strings like "3.0" -> "3"
            try:
                f = float(p)
                p = str(int(f)) if f.is_integer() else str(f)
            except ValueError:
                pass
            cleaned.add(p)

        # Numerical sort (CO1, CO2, CO10) instead of Lexicographical (CO1, CO10, CO2)
        return ",".join(sorted(cleaned, key=safe_int))
    
    def _derive_sorted_co_numbers(self):
        """Derive sorted unique CO numbers from direct components/questions."""
        co_nums = sorted({
            int(str(co).replace("CO", "").strip())
            for comp in self.validated.components.values()
            if comp.direct
            for q in comp.questions
            for co in q.co_list
        })
        return co_nums

    # =====================================================
    # VALIDATION METHODS
    # =====================================================

    def _validate_sheet_names(self):
        # Expected component/direct sheets
        expected_raw = {name for name, comp in self.validated.components.items() if comp.direct}

        # Expected indirect sheets
        for tool in self.validated.indirect_tools:
            expected_raw.add(f"{tool.name}_Indirect")

        system_sheets_raw = {"__SYSTEM_HASH__", "Course_Info"}

        expected_norm = {normalize(name) for name in expected_raw}
        system_norm = {normalize(name) for name in system_sheets_raw}
        actual_norm = set(self._norm_map.keys())

        # Mandatory system sheets must exist
        for s in system_sheets_raw:
            if normalize(s) not in actual_norm:
                raise ValidationError(f"Missing mandatory sheet: {s}")

        # Component + indirect sheets must match exactly (excluding system sheets)
        actual_filtered_norm = actual_norm - system_norm
        if expected_norm != actual_filtered_norm:
            found_raw = {self._norm_map[k] for k in actual_filtered_norm}
            raise ValidationError(f"Sheet mismatch. Expected: {expected_raw}, Found: {found_raw}")

    def _validate_direct_structure(self):
        expected_regnos = [str(s.reg_no).strip().upper() for s in self.validated.students]
        expected_set = set(expected_regnos)

        for comp_name, comp in self.validated.components.items():
            if not comp.direct:
                continue

            df = self._get_sheet_ignore_case(comp_name)
            vals = df.values  # numpy backing

            expected_header = ["RegNo", "Student_Name"] + [q.identifier for q in comp.questions] + ["Total"]
            num_cols = len(expected_header)

            # Shape guards (avoid IndexError)
            if vals.shape[0] < 4:
                raise ValidationError(f"{comp_name}: Sheet has insufficient rows (need at least 4). Found {vals.shape[0]}.")
            if vals.shape[1] < num_cols:
                raise ValidationError(f"{comp_name}: Sheet has insufficient columns. Expected {num_cols}, found {vals.shape[1]}.")

            # --- 1) Strict Header Validation (row 0) ---
            actual_header = [self._canon_cell_text(x) for x in vals[0, :num_cols]]
            expected_header_c = [self._canon_cell_text(x) for x in expected_header]
            if actual_header != expected_header_c:
                raise ValidationError(
                    f"{comp_name}: Header tampered.\nExpected: {expected_header_c}\nFound:    {actual_header}"
                )

            # --- 2) CO + Max rows ---
            # rows: 1 (CO), 2 (Max). cols: question region only (exclude first 2 and last Total)
            struct_matrix = vals[1:3, 2:num_cols - 1]

            # CO row strict (but canonical-sorted)
            actual_cos = [self._canon_co_cell_sorted(x) for x in struct_matrix[0]]
            expected_cos = [self._canon_co_cell_sorted(",".join(q.co_list)) for q in comp.questions]
            if actual_cos != expected_cos:
                raise ValidationError(
                    f"{comp_name}: CO row tampered.\nExpected: {expected_cos}\nFound:    {actual_cos}"
                )

            # Max marks row strict numeric with tolerance
            try:
                actual_max = struct_matrix[1].astype(float)
            except Exception:
                raise ValidationError(f"{comp_name}: Max row contains non-numeric values.")

            expected_max = np.array([float(q.max_marks) for q in comp.questions], dtype=float)
            if actual_max.shape != expected_max.shape or not np.allclose(actual_max, expected_max, atol=1e-6):
                raise ValidationError(
                    f"{comp_name}: Max marks tampered.\nExpected: {expected_max.tolist()}\nFound:    {actual_max.tolist()}"
                )

            # --- 3) Student list: SET-based, report missing/extra/duplicates ---
            actual_regnos = [str(x).strip().upper() for x in vals[3:, 0] if str(x).strip() != ""]
            actual_set = set(actual_regnos)

            # duplicates in sheet
            counts = Counter(actual_regnos)
            dup_list = [k for k, v in counts.items() if v > 1]

            missing = sorted(expected_set - actual_set)
            extra = sorted(actual_set - expected_set)

            if missing or extra or dup_list:
                lines = [f"{comp_name}: Student list tampered."]
                if missing:
                    lines.append(f"Missing RegNo(s): {missing}")
                if extra:
                    lines.append(f"Unexpected RegNo(s): {extra}")
                if dup_list:
                    lines.append(f"Duplicate RegNo(s): {dup_list}")
                raise ValidationError("\n".join(lines))

    def _validate_indirect_structure(self):
        co_nums = self._derive_sorted_co_numbers()
        expected_cols = ["RegNo", "Student_Name"] + [f"CO{i}" for i in co_nums]
        expected_cols_c = [self._canon_cell_text(c) for c in expected_cols]

        for tool in self.validated.indirect_tools:
            sheet_name = f"{tool.name}_Indirect"
            df = self._get_sheet_ignore_case(sheet_name)

            actual_cols_c = [self._canon_cell_text(c) for c in df.columns.tolist()]

            # STRICT: must match at least the expected prefix in order; also disallow missing/renamed columns.
            if actual_cols_c != expected_cols_c:
                raise ValidationError(
                    f"{sheet_name}: Header tampered.\nExpected: {expected_cols_c}\nFound:    {actual_cols_c}"
                )

    def _validate_checksum(self):
        df_hash = self._get_sheet_ignore_case("__SYSTEM_HASH__")

        if df_hash.shape != (1, 1):
            raise ValidationError(f"__SYSTEM_HASH__ shape invalid. Expected (1,1), found {df_hash.shape}.")

        stored_hash = str(df_hash.iat[0, 0]).strip()
        if not stored_hash:
            raise ValidationError("__SYSTEM_HASH__ stored hash is blank.")

        computed = self._compute_structure_hash()
        if stored_hash != computed:
            raise ValidationError("Template structure tampered (checksum mismatch).")

    def _compute_structure_hash(self):
        """
        Deterministic hash based on:
        - component/question identity + max marks + CO mapping (COs sorted)
        - student set (sorted regnos)
        NOTE: Your template renderer must compute the same hash with the same rules.
        """
        hasher = hashlib.sha256()

        for name, comp in sorted(self.validated.components.items()):
            hasher.update(name.encode())

            for q in comp.questions:
                hasher.update(str(q.identifier).encode())
                hasher.update(f"{float(q.max_marks):.2f}".encode())

                canon = self._canon_co_cell_sorted(",".join(q.co_list))
                hasher.update(canon.encode())

        student_regnos = [str(s.reg_no).strip().upper() for s in self.validated.students]
        for regno in sorted(student_regnos):
            hasher.update(regno.encode())

        return hasher.hexdigest()