import pandas as pd
import numpy as np
import re
from dataclasses import dataclass
from collections import Counter
from core.exceptions import ValidationError
from modules.utils import normalize, safe_int

import hashlib

class HashStrategy:
    """
    Single Source of Truth for template fingerprinting.
    Used to 'lock' the template during generation and 'verify' it during upload.
    """
    @staticmethod
    def compute_structure_hash(validated_setup, canon_func) -> str:
        hasher = hashlib.sha256()

        # 1. Components & Questions (Deterministic Sort)
        for name, comp in sorted(validated_setup.components.items()):
            hasher.update(name.encode())
            for q in comp.questions:
                hasher.update(str(q.identifier).encode())
                # Use fixed-point for floats to avoid precision noise
                hasher.update(f"{float(q.max_marks):.2f}".encode())
                # Apply the strict CO canonicalizer
                hasher.update(canon_func(",".join(q.co_list)).encode())

        # 2. Students (Deterministic Sort)
        student_regnos = [str(s.reg_no).strip().upper() for s in validated_setup.students]
        for regno in sorted(student_regnos):
            hasher.update(regno.encode())

        return hasher.hexdigest()

@dataclass(frozen=True)
class TemplateLayout:
    version: str
    header_row: int
    co_row: int
    max_row: int
    data_start_row: int
    reg_col: int
    name_col: int
    marks_start_col: int

    def get_question_headers(self, vals: np.ndarray, num_qs: int) -> list:
        return vals[self.header_row, self.marks_start_col : self.marks_start_col + num_qs].tolist()

    def get_question_region(self, vals: np.ndarray, num_qs: int) -> np.ndarray:
        # Returns the CO and Max Marks slice
        return vals[self.co_row : self.max_row + 1, self.marks_start_col : self.marks_start_col + num_qs]

    def get_student_regnos(self, vals: np.ndarray) -> list:
        return vals[self.data_start_row :, self.reg_col].tolist()

LAYOUT_REGISTRY = {
    "v1": TemplateLayout("v1", 0, 1, 2, 3, 0, 1, 2),
    "v2": TemplateLayout("v2", 1, 2, 3, 4, 0, 1, 2),
}

class FilledMarksValidator:
    def __init__(self, validated_setup, sheets: dict):
        self.validated = validated_setup
        self.sheets = sheets
        self._norm_map = {normalize(k): k for k in self.sheets.keys()}
        
        # 1. Pre-flight Manifest Check (Strict)
        self.metadata = self._parse_system_metadata()
        
        # 2. Strategy Assignment
        if self.metadata['version'] not in LAYOUT_REGISTRY:
            raise ValidationError(f"Unsupported template version: {self.metadata['version']}")
        self.layout = LAYOUT_REGISTRY[self.metadata['version']]

    def _parse_system_metadata(self) -> dict:
        try:
            df = self._get_sheet_ignore_case("__SYSTEM_HASH__")
        except ValidationError:
            raise ValidationError("Mandatory sheet '__SYSTEM_HASH__' missing.")

        # Hard Fail on Malformed Manifest
        if df.shape[0] < 1 or df.shape[1] < 3:
            raise ValidationError("__SYSTEM_HASH__ is malformed. Version/Count columns missing.")

        return {
            "hash": str(df.iloc[0, 0]).strip(),
            "version": str(df.iloc[0, 1]).strip(),
            "expected_count": safe_int(df.iloc[0, 2])
        }

    def validate(self):
        self._validate_sheet_names()
        self._validate_direct_structure()
        self._validate_indirect_structure()
        self._validate_checksum()

    def _validate_direct_structure(self):
        expected_regnos = [str(s.reg_no).strip().upper() for s in self.validated.students]

        for comp_name, comp in self.validated.components.items():
            if not comp.direct: continue
            
            df = self._get_sheet_ignore_case(comp_name)
            vals = df.values
            num_qs = len(comp.questions)

            # Shape Guard
            min_rows = max(self.layout.max_row + 1, self.layout.data_start_row + 1)
            min_cols = self.layout.marks_start_col + num_qs
            if vals.shape[0] < min_rows or vals.shape[1] < min_cols:
                raise ValidationError(f"Sheet '{comp_name}' is smaller than the required layout.")

            # 1. Header Identity
            actual_header = [self._canon_cell_text(x) for x in self.layout.get_question_headers(vals, num_qs)]
            if actual_header != [q.identifier for q in comp.questions]:
                raise ValidationError(f"{comp_name}: Question ID mismatch.")

            # 2. CO & Max Integrity
            q_matrix = self.layout.get_question_region(vals, num_qs)
            
            # CO Row
            actual_cos = [self._canon_co_cell_sorted_strict(x) for x in q_matrix[0]]
            expected_cos = [self._canon_co_cell_sorted_strict(",".join(q.co_list)) for q in comp.questions]
            if actual_cos != expected_cos:
                raise ValidationError(f"{comp_name}: CO mapping tampered.")

            # Max Row
            actual_max = q_matrix[1].astype(float)
            expected_max = np.array([float(q.max_marks) for q in comp.questions])
            if not np.allclose(actual_max, expected_max, atol=1e-6):
                raise ValidationError(f"{comp_name}: Max marks tampered.")

            # 3. Student List
            actual_regnos = [str(x).strip().upper() for x in self.layout.get_student_regnos(vals) if str(x).strip()]
            self._report_student_diffs(comp_name, actual_regnos, expected_regnos)

    def _validate_indirect_structure(self):
        co_nums = self._derive_sorted_co_numbers()
        expected_cols = ["RegNo", "Student_Name"] + [f"CO{i}" for i in co_nums]
        
        for tool in self.validated.indirect_tools:
            sheet_name = f"{tool.name}_Indirect"
            df = self._get_sheet_ignore_case(sheet_name)
            if [normalize(c) for c in df.columns] != [normalize(c) for c in expected_cols]:
                raise ValidationError(f"{sheet_name}: Strict column order violation.")

    def _validate_checksum(self):
        computed = HashStrategy.compute_structure_hash(self.validated, self._canon_co_cell_sorted_strict)
        if self.metadata['hash'] != computed:
            raise ValidationError("Template structure tampered (Checksum mismatch).")

    # --- Strict Helpers ---

    @staticmethod
    def _canon_co_cell_sorted_strict(x) -> str:
        if pd.isna(x): return ""
        s = str(x).strip()
        if not s: return ""
        parts = re.split(r'[;,|]+', s.replace(" ", "").upper().replace("CO", ""))
        try:
            numeric = sorted(int(float(p)) for p in parts if p)
            return ",".join(str(n) for n in numeric)
        except (ValueError, TypeError):
            raise ValidationError(f"Invalid CO identifier: '{x}'. Must be numeric.")

    @staticmethod
    def _canon_cell_text(x) -> str:
        if pd.isna(x): return ""
        if isinstance(x, (int, float, np.integer, np.floating)):
            fx = float(x)
            return str(int(fx)) if fx.is_integer() else str(fx)
        return str(x).strip()

    def _get_sheet_ignore_case(self, name):
        key = self._norm_map.get(normalize(name))
        if not key: raise ValidationError(f"Missing sheet: {name}")
        return self.sheets[key]

    def _report_student_diffs(self, comp_name, actual, expected):
        actual_set, expected_set = set(actual), set(expected)
        missing = sorted(expected_set - actual_set)
        extra = sorted(actual_set - expected_set)
        dups = [k for k, v in Counter(actual).items() if v > 1]
        if missing or extra or dups:
            raise ValidationError(f"{comp_name}: Student list mismatch. Missing: {missing}, Extra: {extra}, Dups: {dups}")

    def _derive_sorted_co_numbers(self):
        nums = {int(str(co).replace("CO", "")) for c in self.validated.components.values() if c.direct for q in c.questions for co in q.co_list}
        return sorted(list(nums))

    def _validate_sheet_names(self):
        # Strict sheet presence check
        expected = {normalize(n) for n in self.validated.components if self.validated.components[n].direct}
        expected.update({normalize(f"{t.name}_Indirect") for t in self.validated.indirect_tools})
        actual = set(self._norm_map.keys()) - {normalize("__SYSTEM_HASH__"), normalize("Course_Info")}
        if expected != actual:
            raise ValidationError("Sheet name/count mismatch.")