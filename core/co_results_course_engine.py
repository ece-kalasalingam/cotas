# core/co_results_course_engine.py

from __future__ import annotations

import re
from typing import Dict, List, TypedDict, Optional

import numpy as np
import pandas as pd

from core.exceptions import ValidationError
from core.constants import (
    DIRECT_RATIO,
    INDIRECT_RATIO,
    ABSENT_SYMBOL,
    PASS_MARK,
    THRESHOLD_MARK,
    HIGH_BENCHMARK_MARK,
)

class SectionResult(TypedDict):
    result_path: str


_CO_DIRECT_RE = re.compile(r"^CO(\d+)_Direct$", re.IGNORECASE)
_CO_INDIRECT_RE = re.compile(r"^CO(\d+)_Indirect$", re.IGNORECASE)
_BRACKET_RE = re.compile(r"\((\d+(?:\.\d+)?)\)")


class COResultsCourseEngine:

    def __init__(self, section_results: List[SectionResult]) -> None:
        if not section_results:
            raise ValidationError("No section result files provided.")

        self.section_results = section_results
        self.course_code: Optional[str] = None
        self.total_cos: Optional[int] = None
        self.sections: List[str] = []
        self.global_regnos: set[str] = set()
        self.co_data: Dict[int, List[pd.DataFrame]] = {}

    # =========================================================
    # PUBLIC ENTRY
    # =========================================================

    def compute_and_export(self, output_path: str) -> None:

        for item in self.section_results:
            self._process_file(item["result_path"])

        if not self.co_data:
            raise ValidationError("No valid CO data found.")

        # Final student consistency across COs
        expected_total: Optional[int] = None
        for co, frames in self.co_data.items():
            total = sum(len(f) for f in frames)
            if expected_total is None:
                expected_total = total
            elif total != expected_total:
                raise ValidationError("Total student count mismatch across COs.")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

            summary_rows: List[Dict[str, object]] = []

            for co in sorted(self.co_data.keys()):

                df = pd.concat(self.co_data[co], ignore_index=True)
                df = df.sort_values("RegNo").reset_index(drop=True)

                df["CO Attainment Level"] = self._compute_level(df["Total attainment"])

                sheet_name = f"CO{co}"

                df.to_excel(writer, sheet_name=sheet_name, index=False)

                # ---- Add Level Summary Footer ----
                counts = df["CO Attainment Level"].value_counts().to_dict()
                total = len(df)

                l0 = counts.get(0, 0)
                l1 = counts.get(1, 0)
                l2 = counts.get(2, 0)
                l3 = counts.get(3, 0)

                footer_start_row = len(df) + 2  # one empty row gap

                ws = writer.sheets[sheet_name]

                ws.cell(row=footer_start_row, column=1, value="Summary")
                ws.cell(row=footer_start_row + 1, column=1, value="Level 0")
                ws.cell(row=footer_start_row + 1, column=2, value=l0)

                ws.cell(row=footer_start_row + 2, column=1, value="Level 1")
                ws.cell(row=footer_start_row + 2, column=2, value=l1)

                ws.cell(row=footer_start_row + 3, column=1, value="Level 2")
                ws.cell(row=footer_start_row + 3, column=2, value=l2)

                ws.cell(row=footer_start_row + 4, column=1, value="Level 3")
                ws.cell(row=footer_start_row + 4, column=2, value=l3)

                ws.cell(row=footer_start_row + 5, column=1, value="Total Stud.")
                ws.cell(row=footer_start_row + 5, column=2, value=total)

                counts = df["CO Attainment Level"].value_counts().to_dict()
                total = len(df)

                l0 = counts.get(0, 0)
                l1 = counts.get(1, 0)
                l2 = counts.get(2, 0)
                l3 = counts.get(3, 0)

                co_pct = ((l2 + l3) / total) * 100 if total else 0

                summary_rows.append({
                    "CO": f"CO{co}",
                    "Level 0": l0,
                    "Level 1": l1,
                    "Level 2": l2,
                    "Level 3": l3,
                    "Total": total,
                    "CO%": round(co_pct, 2),
                    "Normalized(0-3)": round(co_pct * 3 / 100, 2),
                    "Status": "Achieved" if co_pct >= THRESHOLD_MARK else "Not Achieved",
                })

            pd.DataFrame(summary_rows).to_excel(
                writer,
                sheet_name="Overall_Summary",
                index=False
            )

    # =========================================================
    # FILE PROCESSING
    # =========================================================

    def _process_file(self, path: str) -> None:

        xl = pd.ExcelFile(path)
        sheets = xl.sheet_names

        direct_map: Dict[int, str] = {}
        indirect_map: Dict[int, str] = {}

        for sheet in sheets:
            name = str(sheet)

            m = _CO_DIRECT_RE.match(name)
            if m:
                direct_map[int(m.group(1))] = name
                continue

            m = _CO_INDIRECT_RE.match(name)
            if m:
                indirect_map[int(m.group(1))] = name

        if not direct_map:
            raise ValidationError(f"{path}: No CO sheets detected.")

        all_cos = sorted(direct_map.keys())

        for co in all_cos:
            if co not in indirect_map:
                raise ValidationError(f"{path}: CO{co} missing Indirect sheet.")

        if self.total_cos is None:
            self.total_cos = len(all_cos)
        elif len(all_cos) != self.total_cos:
            raise ValidationError(f"{path}: Total CO mismatch.")

        # Metadata extraction by key
        meta = pd.read_excel(path, sheet_name=direct_map[all_cos[0]], header=None, nrows=20)

        course_code = ""
        section = ""

        for _, row in meta.iterrows():
            key = str(row.iloc[0]).strip().lower()
            val = str(row.iloc[1]).strip()

            if key == "course code":
                course_code = val
            elif key == "section":
                section = val

        if not course_code:
            raise ValidationError(f"{path}: Missing Course Code.")
        if not section:
            raise ValidationError(f"{path}: Missing Section.")

        if self.course_code is None:
            self.course_code = course_code
        elif self.course_code != course_code:
            raise ValidationError(f"{path}: Course code mismatch.")

        if section in self.sections:
            raise ValidationError(f"{path}: Duplicate section '{section}'.")
        self.sections.append(section)

        reference_students: Optional[List[str]] = None

        for co in all_cos:

            d_df = self._read_direct(path, direct_map[co])
            i_df = self._read_indirect(path, indirect_map[co])

            if not d_df["RegNo"].equals(i_df["RegNo"]):
                raise ValidationError(f"{path}: CO{co} student mismatch between Direct and Indirect.")

            if reference_students is None:
                reference_students = d_df["RegNo"].tolist()
            elif d_df["RegNo"].tolist() != reference_students:
                raise ValidationError(f"{path}: Student list mismatch across CO sheets.")

            dup = set(d_df["RegNo"]).intersection(self.global_regnos)
            if dup:
                raise ValidationError(f"{path}: Duplicate RegNo across sections.")

            combined = pd.DataFrame({
                "RegNo": d_df["RegNo"],
                "Student Name": d_df["Student Name"],
                "Direct 100%": d_df["Direct100"],
                f"Direct {int(DIRECT_RATIO*100)}%": d_df["DirectRatio"],
                "Indirect 100%": i_df["Indirect100"],
                f"Indirect {int(INDIRECT_RATIO*100)}%": i_df["IndirectRatio"],
            })

            combined["Total attainment"] = (
                combined[f"Direct {int(DIRECT_RATIO*100)}%"] +
                combined[f"Indirect {int(INDIRECT_RATIO*100)}%"]
            ).round(2)

            self.co_data.setdefault(co, []).append(combined)

        if reference_students:
            self.global_regnos.update(reference_students)

    # =========================================================
    # DIRECT VALIDATION
    # =========================================================

    def _read_direct(self, path: str, sheet: str) -> pd.DataFrame:

        header_row = self._find_header(path, sheet)
        df = pd.read_excel(path, sheet_name=sheet, header=header_row).dropna(how="all")

        df.columns = list(map(str, df.columns))

        normalized_cols = {c.replace(" ", "").lower(): c for c in df.columns}

        if "100%" not in normalized_cols:
            raise ValidationError(f"{path}::{sheet}: Missing 100% column.")

        percent_col = normalized_cols["100%"]

        ratio_key = f"{int(DIRECT_RATIO*100)}%".replace(" ", "").lower()

        if ratio_key not in normalized_cols:
            raise ValidationError(f"{path}::{sheet}: Missing {ratio_key} column.")

        ratio_col_name = normalized_cols[ratio_key]

        reg = df["RegNo"].astype(str).str.strip()
        name = df["Student Name"].astype(str).str.strip()

        total_col = [c for c in df.columns if c.startswith("Total")]
        if not total_col:
            raise ValidationError(f"{path}::{sheet}: Missing Total column.")

        total_col = total_col[0]
        total_max = self._extract_bracket(total_col)

        numeric_total = self._validate_numeric(df[total_col], 0, total_max)

        direct100 = (numeric_total * 100 / total_max).round(2)
        stored100 = self._validate_numeric(df[percent_col], 0, 100)

        if not np.allclose(direct100, stored100, atol=0.1):
            raise ValidationError(f"{path}::{sheet}: Direct 100% mismatch.")

        ratio = (direct100 * DIRECT_RATIO).round(2)
        stored_ratio = self._validate_numeric(df[ratio_col_name], 0, 100)

        if not np.allclose(ratio, stored_ratio, atol=0.1):
            raise ValidationError(f"{path}::{sheet}: Direct ratio mismatch.")

        return pd.DataFrame({
            "RegNo": reg,
            "Student Name": name,
            "Direct100": direct100,
            "DirectRatio": ratio
        })

    # =========================================================
    # INDIRECT VALIDATION
    # =========================================================

    def _read_indirect(self, path: str, sheet: str) -> pd.DataFrame:

        header_row = self._find_header(path, sheet)
        df = pd.read_excel(path, sheet_name=sheet, header=header_row).dropna(how="all")

        df.columns = list(map(str, df.columns))

        normalized_cols = {
            c.replace(" ", "").lower(): c
            for c in df.columns
        }

        # --- 100% column ---
        if "100%" not in normalized_cols:
            raise ValidationError(f"{path}::{sheet}: Missing 100% column.")

        percent_col = normalized_cols["100%"]

        # --- Ratio column ---
        ratio_key = f"{int(INDIRECT_RATIO*100)}%".replace(" ", "").lower()

        if ratio_key not in normalized_cols:
            raise ValidationError(f"{path}::{sheet}: Missing {ratio_key} column.")

        ratio_col_name = normalized_cols[ratio_key]

        # --- Required base columns ---
        if "regno" not in normalized_cols:
            raise ValidationError(f"{path}::{sheet}: Missing RegNo column.")

        if "studentname" not in normalized_cols and "student name" not in normalized_cols:
            raise ValidationError(f"{path}::{sheet}: Missing Student Name column.")

        reg_col = normalized_cols["regno"]
        name_col = normalized_cols.get("studentname") or normalized_cols.get("student name")

        reg = df[reg_col].astype(str).str.strip()
        name = df[name_col].astype(str).str.strip()

        # --- Numeric validation ---
        indirect100 = self._validate_numeric(df[percent_col], 0, 100)
        expected_ratio = (indirect100 * INDIRECT_RATIO).round(2)
        stored_ratio = self._validate_numeric(df[ratio_col_name], 0, 100)

        if not np.allclose(expected_ratio, stored_ratio, atol=0.1):
            raise ValidationError(f"{path}::{sheet}: Indirect ratio mismatch.")

        return pd.DataFrame({
            "RegNo": reg,
            "Student Name": name,
            "Indirect100": indirect100,
            "IndirectRatio": expected_ratio
        })

    # =========================================================
    # UTILITIES
    # =========================================================

    @staticmethod
    def _find_header(path: str, sheet: str) -> int:
        preview = pd.read_excel(path, sheet_name=sheet, header=None, nrows=50)
        for i in range(len(preview)):
            row = preview.iloc[i].astype(str).str.lower().tolist()
            if "regno" in row and "student name" in row:
                return i
        raise ValidationError(f"{path}::{sheet}: Header row not found.")

    @staticmethod
    def _extract_bracket(header: str) -> float:
        m = _BRACKET_RE.search(str(header))
        if not m:
            raise ValidationError(f"Missing bracket value in '{header}'.")
        return float(m.group(1))

    @staticmethod
    def _validate_numeric(series: pd.Series, min_v: float, max_v: float) -> pd.Series:

        s = series.astype(str).str.strip()

        is_absent = s.str.lower() == ABSENT_SYMBOL.lower()
        s = s.mask(is_absent, np.nan)

        numeric = pd.to_numeric(s, errors="coerce")

        if numeric.isna().any() and not is_absent.all():
            raise ValidationError("Invalid numeric value detected.")

        if numeric.dropna().between(min_v, max_v).all() is False:
            raise ValidationError("Value out of allowed range.")

        return numeric.fillna(0.0)

    @staticmethod
    def _compute_level(series: pd.Series) -> pd.Series:
        conditions = [
            series < PASS_MARK,
            (series >= PASS_MARK) & (series < THRESHOLD_MARK),
            (series >= THRESHOLD_MARK) & (series < HIGH_BENCHMARK_MARK),
            series >= HIGH_BENCHMARK_MARK,
        ]
        return pd.Series(np.select(conditions, [0, 1, 2, 3]), index=series.index)