## core/co_results_course_engine.py

from __future__ import annotations
import re
from dataclasses import dataclass
from typing import (
    Dict, List, TypedDict, Optional, Any, 
    Tuple, Set, Mapping, Sequence, Iterable, cast
)

import numpy as np
import pandas as pd
from core.exceptions import ValidationError
from core.constants import (
   DIRECT_RATIO, INDIRECT_RATIO, ABSENT_SYMBOL
)

ABSENT_LEVEL = -1
EPSILON = 1e-9

class SectionResult(TypedDict):
    result_path: str

_CO_DIRECT_RE = re.compile(r"^CO(\d+)_Direct$", re.IGNORECASE)
_CO_INDIRECT_RE = re.compile(r"^CO(\d+)_Indirect$", re.IGNORECASE)
_BRACKET_RE = re.compile(r"\((\d+(?:\.\d+)?)\)")

@dataclass(frozen=True)
class AttainmentPolicy:
    pass_mark: float
    threshold_mark: float
    high_mark: float
    target_percent: float

    def __post_init__(self):
        params = {"pass_mark": self.pass_mark, "threshold_mark": self.threshold_mark, 
                  "high_mark": self.high_mark, "target_percent": self.target_percent}
        for name, val in params.items():
            if not isinstance(val, (int, float)) or not (0 <= val <= 100):
                raise ValidationError(f"Policy Error: {name} must be numeric (0-100).")
        if not (self.pass_mark < self.threshold_mark < self.high_mark):
            raise ValidationError("Policy Logic Error: Order must be pass < threshold < high.")

    def classify(self, series: pd.Series) -> pd.Series:
        eps = 1e-9
        p_series = series.round(4)
        conds = [
            p_series < self.pass_mark - eps,
            (p_series >= self.pass_mark - eps) & (p_series < self.threshold_mark - eps),
            (p_series >= self.threshold_mark - eps) & (p_series < self.high_mark - eps),
            p_series >= self.high_mark - eps,
        ]
        return pd.Series(np.select(conds, [0, 1, 2, 3]), index=series.index)

    def is_achieved(self, percentage: float) -> str:
        return "Achieved" if percentage >= self.target_percent else "Not Achieved"

class COResultsCourseEngine:
    def __init__(self, section_results: List[SectionResult], policy: AttainmentPolicy) -> None:
        if not section_results:
            raise ValidationError("No section result files provided.")
        
        # Ratio Sum Validation
        if not np.isclose(DIRECT_RATIO + INDIRECT_RATIO, 1.0):
            raise ValidationError(
                f"Global Ratio Configuration Mismatch: {DIRECT_RATIO} + {INDIRECT_RATIO} != 1.0"
            )

        self.section_results = section_results
        self.policy = policy
        
        self._co_data_mut: Dict[int, List[pd.DataFrame]] = {}
        self._co_student_sets_mut: Dict[int, Set[str]] = {}
        
        self.co_data: Mapping[int, Sequence[pd.DataFrame]] = {}
        self.co_student_sets: Mapping[int, frozenset[str]] = {}
        
        self.course_code: Optional[str] = None
        self.sections: List[str] = []
        self.global_student_map: Dict[str, str] = {}
        self._expected_cos: Optional[Set[int]] = None

    def compute_and_export(self, output_path: str) -> None:
        for item in self.section_results:
            self._process_file(item["result_path"])

        if not self._co_data_mut:
            raise ValidationError("No valid CO data processed.")

        self.co_data = {k: tuple(v) for k, v in self._co_data_mut.items()}
        self.co_student_sets = {k: frozenset(v) for k, v in self._co_student_sets_mut.items()}
        
        co_keys = sorted(self.co_data.keys())
        ref_co = co_keys[0]
        ref_set = self.co_student_sets[ref_co]

        if not ref_set:
            raise ValidationError(f"CO{ref_co} contains no student data.")

        for co in co_keys:
            current_set = self.co_student_sets[co]
            if current_set != ref_set:
                missing = sorted(list(ref_set - current_set))[:5]
                extra = sorted(list(current_set - ref_set))[:5]
                raise ValidationError(
                    f"Global Student ID Mismatch at CO{co}.\n"
                    f"Missing from CO{co}: {missing}\n"
                    f"Extras in CO{co}: {extra}"
                )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary_rows: List[Dict[str, Any]] = []
            for co in co_keys:
                #df = pd.concat(list(self.co_data[co]), ignore_index=True)
                df = pd.concat(self.co_data[co], ignore_index=True)
                df = df.sort_values("RegNo").reset_index(drop=True)
                df["CO Attainment Level"] = self.policy.classify(df["RawTotal"])
                df.loc[df["IsAbsent"], "CO Attainment Level"] = ABSENT_LEVEL
                
                sheet_name = f"CO{co}"
                #df.drop(columns=["RawTotal"]).to_excel(writer, sheet_name=sheet_name, index=False)
                display_df = df.copy()
                display_df["CO Attainment Level"] = display_df["CO Attainment Level"].replace(ABSENT_LEVEL, "Absent")
                display_df.drop(columns=["RawTotal", "IsAbsent"]).to_excel(writer, sheet_name=sheet_name, index=False)

                total = len(df)
                valid_df = df[df["CO Attainment Level"] != ABSENT_LEVEL]
                #counts = df["CO Attainment Level"].value_counts().to_dict()
                eligible_count = len(valid_df)
                absent_count = total - eligible_count

                counts = valid_df["CO Attainment Level"].value_counts().to_dict()

                l_stats = {f"Level {i}": counts.get(i, 0) for i in range(4)}
                self._write_footer(writer.sheets[sheet_name], {**l_stats, "Absent": absent_count}, total)
                #self._write_footer(writer.sheets[sheet_name], l_stats, total)

                co_pct = ((l_stats["Level 2"] + l_stats["Level 3"]) / eligible_count * 100) if eligible_count > 0 else 0.0
                summary_rows.append({
                    "CO": f"CO{co}", **l_stats, "Absent": absent_count, "Eligible": eligible_count,
                    "CO%": round(co_pct, 2), "Normalized(0-3)": round(co_pct * 3 / 100, 2),
                    "Status": self.policy.is_achieved(co_pct),
                })
                #summary_rows.append({
                    #"CO": f"CO{co}", **l_stats, "Total": total, "CO%": round(co_pct, 2),
                    #"Normalized(0-3)": round(co_pct * 3 / 100, 2),
                    #"Status": self.policy.is_achieved(co_pct),
                #})
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Overall_Summary", index=False)

    def _process_file(self, path: str) -> None:
        xl = pd.ExcelFile(path)
        d_map, i_map = self._map_sheets(
            cast(Sequence[str], xl.sheet_names),
            path
        )
        all_cos = sorted(d_map.keys())

        if all_cos != list(range(1, len(all_cos) + 1)):
            raise ValidationError(f"{path}: Non-continuous CO sequence: {all_cos}")

        if self._expected_cos is None:
            self._expected_cos = set(all_cos)
        elif set(all_cos) != self._expected_cos:
            raise ValidationError(f"{path}: CO count mismatch.")

        course_code, section = self._extract_metadata(xl, d_map[all_cos[0]], path)
        
        if section in self.sections:
            raise ValidationError(f"Duplicate Section '{section}' in file: {path}")
        
        if self.course_code is None:
            self.course_code = course_code
        elif self.course_code != course_code:
            raise ValidationError(f"{path}: Course Code mismatch (Found: {course_code}, Expected: {self.course_code})")
        
        self.sections.append(section)

        for co in all_cos:
            d_df = self._read_direct(xl, d_map[co], path)
            i_df = self._read_indirect(xl, i_map[co], path)
            
            d_set, i_set = set(d_df["RegNo"]), set(i_df["RegNo"])
            if d_set != i_set:
                missing = sorted(list(d_set - i_set))[:3]
                raise ValidationError(f"{path}::CO{co}: Direct/Indirect ID mismatch. Missing: {missing}")

            if len(d_df) != len(d_set):
                dupes = d_df[d_df.duplicated("RegNo")]["RegNo"].tolist()
                raise ValidationError(f"{path}::CO{co}: Duplicate IDs: {dupes}")

            current_names = dict(zip(d_df["RegNo"], d_df["Student Name"]))
            for r, n in current_names.items():
                if r in self.global_student_map and self.global_student_map[r] != n:
                    raise ValidationError(f"Name Conflict for {r}: '{n}' vs '{self.global_student_map[r]}'")
                self.global_student_map[r] = n

            is_absent = d_df["IsAbsent"] | i_df["IsAbsent"]
            raw_total = (d_df["Score"] * DIRECT_RATIO) + (i_df["Score"] * INDIRECT_RATIO)
            
            combined = pd.DataFrame({
                "RegNo": d_df["RegNo"], "Student Name": d_df["Student Name"],
                "Direct 100%": d_df["Score"].round(2),
                "Indirect 100%": i_df["Score"].round(2),
                "Total attainment": raw_total.round(2),
                "RawTotal": raw_total ,
                 "IsAbsent": is_absent 
            })
            self._co_data_mut.setdefault(co, []).append(combined)
            self._co_student_sets_mut.setdefault(co, set()).update(d_set)

    def _read_direct(self, xl: pd.ExcelFile, sheet: str, path: str) -> pd.DataFrame:
        header_row = self._find_header(xl, sheet, path)
        df = xl.parse(sheet, header=header_row).dropna(how="all")
        
        raw_cols: List[str] = [str(c) for c in df.columns]
        norm_cols: List[str] = [c.strip().replace(" ", "").lower() for c in raw_cols]
        df.columns = norm_cols

        required = {"regno", "studentname", "100%"}
        if not required.issubset(set(norm_cols)):
            raise ValidationError(f"{path}::{sheet}: Missing required columns: {required - set(norm_cols)}")

        total_key = next((c for c in norm_cols if c.startswith("total")), None)
        if not total_key:
            raise ValidationError(f"{path}::{sheet}: Missing 'Total' column.")

        idx: int = norm_cols.index(total_key)
        header_str = cast(str, raw_cols[idx]) 
        total_max = self._extract_bracket(header_str)
        
        
        num_total, abs_t = self._validate_numeric_with_mask(df[total_key], 0, total_max, f"{sheet} Total", df)
        num_100, abs_100 = self._validate_numeric_with_mask(df["100%"], 0, 100, f"{sheet} 100%", df)

        if not (abs_t == abs_100).all():
            raise ValidationError(f"{sheet}: Inconsistent Absence marking between Total and 100% columns.")
        
        if not np.allclose(num_total * 100 / total_max, num_100, atol=0.1):
            raise ValidationError(f"{sheet}: Calculation audit failed (Total/{total_max} vs 100% column).")

        return pd.DataFrame({
            "RegNo": df["regno"].astype(str).str.strip().str.upper(),
            "Student Name": df["studentname"].astype(str).str.strip(),
            "Score": num_100, "IsAbsent": abs_100
        })

    def _read_indirect(self, xl: pd.ExcelFile, sheet: str, path: str) -> pd.DataFrame:
        header_row = self._find_header(xl, sheet, path)
        df = xl.parse(sheet, header=header_row).dropna(how="all")
        norm_cols = [str(c).strip().replace(" ", "").lower() for c in df.columns]
        df.columns = norm_cols

        required = {"regno", "studentname", "100%"}
        if not required.issubset(set(norm_cols)):
            raise ValidationError(f"{path}::{sheet}: Missing columns: {required - set(norm_cols)}")

        num_100, abs_100 = self._validate_numeric_with_mask(df["100%"], 0, 100, f"{sheet} Indirect", df)

        return pd.DataFrame({
            "RegNo": df["regno"].astype(str).str.strip().str.upper(),
            "Student Name": df["studentname"].astype(str).str.strip(),
            "Score": num_100, "IsAbsent": abs_100
        })
    
    def _validate_numeric_with_mask(self, series: pd.Series, min_v: float, max_v: float, ctx: str, df: pd.DataFrame) -> Tuple[pd.Series, pd.Series]:
        s = series.astype(str).str.strip()
        absent_norm = ABSENT_SYMBOL.lower().strip()
        
        is_absent = s.str.lower().str.strip() == absent_norm
        num = pd.to_numeric(s.mask(is_absent, "0"), errors="coerce")
        
        if num.isna().any():
            bad_idx = num.index[num.isna()][0]
            raise ValidationError(f"Non-numeric '{series.iloc[bad_idx]}' in {ctx} at RegNo: {df.iloc[bad_idx].get('regno')}")
            
        if not num.between(min_v, max_v + EPSILON).all():
            out_ids = df.loc[~num.between(min_v, max_v + EPSILON), "regno"].tolist()
            raise ValidationError(f"Range Error (0-{max_v}) in {ctx} for: {out_ids[:5]}")
            
        return num, is_absent

    def _map_sheets(self, sheets: Iterable[str], path: str) -> Tuple[Dict[int, str], Dict[int, str]]:
        d_map, i_map = {}, {}
        for s in sheets:
            s_str = str(s)
            if m := _CO_DIRECT_RE.match(s_str):
                co_index: int = int(str(m.group(1)))
                d_map[co_index] = s_str
            elif m := _CO_INDIRECT_RE.match(s_str):
                co_index: int = int(str(m.group(1)))
                i_map[co_index] = s_str
        
        if not d_map: raise ValidationError(f"{path}: No CO_Direct sheets.")
        if set(d_map.keys()) != set(i_map.keys()):
            raise ValidationError(f"{path}: CO Direct/Indirect sheet mismatch")
        for co in d_map:
            if co not in i_map: raise ValidationError(f"{path}: CO{co}_Indirect missing.")
        return d_map, i_map

    def _find_header(self, xl: pd.ExcelFile, sheet: str, path: str) -> int:
        preview = xl.parse(sheet, header=None, nrows=30)

        for idx in range(len(preview)):
            row = preview.iloc[idx]
            vals = [
                str(v).lower().replace(" ", "").replace("_", "").replace("-", "")
                for v in row.values
            ]
            if "regno" in vals and "studentname" in vals:
                return idx

        raise ValidationError(f"{path}::{sheet}: Core headers not found.")

    def _extract_bracket(self, header: str) -> float:
        if m := _BRACKET_RE.search(header): return float(m.group(1))
        raise ValidationError(f"Max mark bracket missing: {header}")

    def _extract_metadata(self, xl: pd.ExcelFile, sheet: str, path: str) -> Tuple[str, str]:
        df = xl.parse(sheet, header=None, nrows=20)
        course, sec = None, None
        for _, row in df.iterrows():
            row_raw = [str(v).strip() for v in row.values]
            for i, val in enumerate(row_raw):
                k = val.lower().replace("_", "").replace("-", "").replace(" ", "").rstrip(":.")
                
                # FIXED: Constrained lookahead (i+1 to i+3) to prevent over-scanning
                if k in {"coursecode", "subjectcode"}:
                    for c in row_raw[i+1 : i+3]:
                        if c and c not in ":.-": course = c.upper(); break
                if k in {"section", "class", "sec"}:
                    for c in row_raw[i+1 : i+3]:
                        if c and c not in ":.-": sec = c.upper(); break
        
        if not course or not sec: raise ValidationError(f"{path}: Course/Section metadata missing.")
        return course, sec

    def _write_footer(self, ws: Any, stats: Dict[str, int], total: int) -> None:
        r = ws.max_row + 3
        ws.cell(row=r, column=1, value="Summary Statistics")
        for i, (k, v) in enumerate(stats.items(), 1):
            ws.cell(row=r+i, column=1, value=k); ws.cell(row=r+i, column=2, value=v)
        ws.cell(row=r+5, column=1, value="Total Registered Students"); ws.cell(row=r+5, column=2, value=total)