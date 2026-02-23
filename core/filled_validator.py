import hashlib
import pandas as pd
from core.exceptions import ValidationError


class FilledMarksValidator:

    def __init__(self, validated_setup, filled_path: str):
        self.validated = validated_setup
        self.filled_path = filled_path
        self.excel = pd.ExcelFile(filled_path)

    # =====================================================
    # PUBLIC ENTRY
    # =====================================================

    def validate(self):

        self._validate_sheets()
        self._validate_direct_sheets()
        self._validate_indirect_sheets()
        self._validate_checksum()

    # =====================================================
    # SHEET VALIDATION
    # =====================================================

    def _validate_sheets(self):

        expected = {
            comp_name
            for comp_name, comp in self.validated.components.items()
            if comp.direct
        }

        # Add indirect sheets
        for tool in self.validated.indirect_tools:
            # Must match the exact naming logic used during sheet creation
            sheet_name = f"{tool.name}_INDIRECT"[:31]
            expected.add(sheet_name)

        # Allowed system sheets
        allowed_system_sheets = {
            "__SYSTEM_HASH__",
            "Course_Info",
        }

        actual = set(self.excel.sheet_names)

        # Remove allowed system sheets before comparison
        actual_filtered = actual - allowed_system_sheets

        if expected != actual_filtered:
            raise ValidationError(
                f"Sheet mismatch. Expected {expected}, "
                f"Found {actual_filtered}"
            )

        # Ensure system sheets exist
        if "__SYSTEM_HASH__" not in actual:
            raise ValidationError("Missing __SYSTEM_HASH__ sheet.")

        if "Course_Info" not in actual:
            raise ValidationError("Missing Course_Info sheet.")

    # =====================================================
    # DIRECT STRUCTURE VALIDATION
    # =====================================================

    def _validate_direct_sheets(self):

        expected_regnos = [s.reg_no for s in self.validated.students]

        for comp_name, comp in self.validated.components.items():

            if not comp.direct:
                continue

            df = pd.read_excel(self.excel, sheet_name=comp_name, header=None)

            # ---- Header Validation ----
            #header = df.iloc[0].tolist()

            #expected_header = ["RegNo", "Student_Name"] + \
                #[q.identifier for q in comp.questions] + ["Total"]

            #if header[:len(expected_header)] != expected_header:
                #raise ValidationError(
                    #f"{comp_name}: Header tampered."
                #)
            def _norm_cell(x) -> str:
                if pd.isna(x):
                    return ""
                # normalize numbers like 1.0 -> "1"
                if isinstance(x, (int, float)):
                    if float(x).is_integer():
                        return str(int(x))
                    return str(x)
                return str(x).strip()

            header_raw = df.iloc[0].tolist()

            expected_raw = ["RegNo", "Student_Name"] + [q.identifier for q in comp.questions] + ["Total"]

            header = [_norm_cell(x) for x in header_raw[:len(expected_raw)]]
            expected = [_norm_cell(x) for x in expected_raw]

            if header != expected:
                raise ValidationError(
                    f"{comp_name}: Header tampered.\n"
                    f"Expected: {expected}\n"
                    f"Found:    {header}"
                )


            # ---- CO Row Validation (type-robust, strict) ----
            def _canon_co_cell(x) -> str:
                """
                Canonicalize a CO mapping cell to a stable form like:
                "1" or "1,2" or "3,4"
                Accepts: 1, 1.0, "1", "CO1", "1;2", "1 | 2", "CO1, CO2", etc.
                """
                if pd.isna(x):
                    return ""

                # Numeric cell like 3.0 -> "3"
                if isinstance(x, (int, float)):
                    fx = float(x)
                    return str(int(fx)) if fx.is_integer() else str(fx)

                s = str(x).strip()
                if not s:
                    return ""

                s = s.replace(";", ",").replace("|", ",")
                parts = [p.strip() for p in s.split(",") if p.strip()]

                norm_parts = []
                for p in parts:
                    p2 = p.upper().replace(" ", "")
                    if p2.startswith("CO"):
                        p2 = p2[2:]  # drop "CO"
                    # handle "3.0" -> "3"
                    try:
                        f = float(p2)
                        p2 = str(int(f)) if f.is_integer() else str(f)
                    except Exception:
                        # if it can't be parsed, keep it (will fail if expected is numeric)
                        pass
                    norm_parts.append(p2)

                # de-dup preserve order
                seen = set()
                out = []
                for p in norm_parts:
                    if p not in seen:
                        out.append(p)
                        seen.add(p)

                return ",".join(out)


            co_row_raw = df.iloc[1].tolist()[2:-1]
            expected_cos_raw = [",".join(q.co_list) for q in comp.questions]

            co_row = [_canon_co_cell(x) for x in co_row_raw]
            expected_cos = [_canon_co_cell(x) for x in expected_cos_raw]

            if co_row != expected_cos:
                raise ValidationError(
                    f"{comp_name}: CO row tampered.\n"
                    f"Expected: {expected_cos}\n"
                    f"Found:    {co_row}"
                )


            # ---- Max Row Validation (float-robust, strict) ----
            max_row_raw = df.iloc[2].tolist()[2:-1]
            expected_max = [float(q.max_marks) for q in comp.questions]

            try:
                actual_max = [float(x) for x in max_row_raw]
            except Exception:
                raise ValidationError(f"{comp_name}: Max row tampered (non-numeric values found).")

            # Compare with rounding to avoid 2 vs 2.0 or Excel float noise
            actual_max_r = [round(x, 6) for x in actual_max]
            expected_max_r = [round(x, 6) for x in expected_max]

            if actual_max_r != expected_max_r:
                raise ValidationError(
                    f"{comp_name}: Max row tampered.\n"
                    f"Expected: {expected_max_r}\n"
                    f"Found:    {actual_max_r}"
                )


            # ---- Student Order Validation (strict, whitespace-safe) ----
            regnos = df.iloc[3:, 0].astype(str).str.strip().tolist()
            expected_regnos = [str(x).strip() for x in expected_regnos]

            if regnos != expected_regnos:
                raise ValidationError(
                    f"{comp_name}: Student list tampered or reordered."
                )


    # =====================================================
    # INDIRECT VALIDATION
    # =====================================================

    def _validate_indirect_sheets(self):

        for tool in self.validated.indirect_tools:

            #sheet_name = f"{tool}_INDIRECT"
            sheet_name = f"{tool.name}_INDIRECT"[:31]
            df = pd.read_excel(self.excel, sheet_name=sheet_name)
            total_cos = len({
                            co
                            for comp in self.validated.components.values()
                            if comp.direct
                            for q in comp.questions
                            for co in q.co_list
                        })

            expected_cols = ["RegNo", "Student_Name"] + \
                [f"CO{i}" for i in range(1, total_cos + 1)]

            if not all(col in df.columns for col in expected_cols[:2]):
                raise ValidationError(
                    f"{sheet_name}: Invalid header."
                )

    # =====================================================
    # CHECKSUM VALIDATION
    # =====================================================

    def _validate_checksum(self):

        df_hash = pd.read_excel(
            self.excel,
            sheet_name="__SYSTEM_HASH__",
            header=None
        )

        if df_hash.shape[0] < 1 or df_hash.shape[1] < 1:
            raise ValidationError("__SYSTEM_HASH__ sheet is empty or invalid.")

        stored_hash = str(df_hash.iloc[0, 0]).strip()
        if not stored_hash:
            raise ValidationError("__SYSTEM_HASH__ stored hash is blank.")

        computed_hash = self._compute_structure_hash()

        if stored_hash != computed_hash:
            raise ValidationError("Template structure tampered (checksum mismatch).")

    # =====================================================
    # HASH GENERATION
    # =====================================================

    def _compute_structure_hash(self):

        hasher = hashlib.sha256()

        for comp_name, comp in sorted(self.validated.components.items()):

            hasher.update(comp_name.encode())

            for q in comp.questions:
                hasher.update(q.identifier.encode())
                hasher.update(str(q.max_marks).encode())
                hasher.update(",".join(q.co_list).encode())

        for student in self.validated.students:
            hasher.update(student.reg_no.encode())

        return hasher.hexdigest()
