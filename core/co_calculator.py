import numpy as np
import pandas as pd
from core.constants import ABSENT_SYMBOL, LIKERT_MAX, LIKERT_MIN, STATUS_A, STATUS_NA, STATUS_NORMAL
from core.exceptions import ValidationError


class COCalculator:
    """
    Pure computation engine.
    No Excel I/O.
    Scalable for large student counts (3000+).
    """

    def __init__(self, validated_setup, direct_sheets: dict, indirect_sheets: dict):
        self.validated = validated_setup
        self.direct_sheets = direct_sheets
        self.indirect_sheets = indirect_sheets

        self.students_df = pd.DataFrame(
            [(s.reg_no, s.name) for s in self.validated.students],
            columns=["RegNo", "Student_Name"]
        )
        self.students_df["RegNo"] = self.students_df["RegNo"].astype(str)

    # =====================================================
    # PUBLIC ENTRY
    # =====================================================

    def run(self) -> dict:
        # 1. Fetch data fragments
        direct_frags = self._collect_direct()
        indirect_frags = self._collect_indirect()

        if direct_frags.empty and indirect_frags.empty:
            raise ValidationError("No assessment data available to process.")

        # 2. Prepare Lookups
        # Convert indirect tools tuple to O(1) map
        indirect_map = {str(t.name).strip(): t for t in self.validated.indirect_tools}
        student_map = {s.reg_no: s.name for s in self.validated.students}
        
        # Identify unique COs across both sources
        direct_cos = set(direct_frags["CO"].unique()) if not direct_frags.empty else set()
        indirect_cos = set(indirect_frags["CO"].unique()) if not indirect_frags.empty else set()
        all_cos = sorted(direct_cos | indirect_cos)

        co_sheets = {}

        # 3. Generate individual CO DataFrames
        for co_val in all_cos:
            # --- Direct Sheet for this CO ---
            df_d = direct_frags[direct_frags["CO"] == co_val]
            if not df_d.empty:
                d_names = [n for n, c in self.validated.components.items() if c.direct]
                co_sheets[f"{co_val}_Direct"] = self._build_direct_data_package(
                    df_d, student_map, d_names
                )

            # --- Indirect Sheet for this CO ---
            df_i = indirect_frags[indirect_frags["CO"] == co_val]
            if not df_i.empty:
                i_names = [t.name for t in self.validated.indirect_tools]
                co_sheets[f"{co_val}_Indirect"] = self._build_data_package(
                    df_i, student_map, i_names, indirect_map
                )

        return {"co_sheets": co_sheets}

    # =====================================================
    # BUILD THE DATA PACKAGE FOR A SINGLE CO SHEET
    # =====================================================
    def _build_direct_data_package(self, df_source, student_map, tool_names) -> pd.DataFrame:
        # This method is similar to _build_data_package but can have direct-specific logic if needed.
        # For now, it simply calls the generic builder since the structure is consistent.
        df_direct = df_source.copy()
        df_direct["RegNo"] = df_direct["RegNo"].astype(str).str.strip()
        df_direct["Component"] = df_direct["Component"].astype(str).str.strip()
        g = (df_direct.groupby(["RegNo", "Component"], as_index=False)
                    .agg(
                        Total=("total", "sum"),
                        Wtd=("Wtd", "sum"),
                         Weight=("Weight", "sum"),
                        IsAbsent=("IsAbsent", "max")
                        ))
        #lookup = g.set_index(["RegNo", "Component"])[["Wtd", "Total", "Weight", "IsAbsent"]].to_dict("index")
        lookup = g.set_index(["RegNo", "Component"])
        active_components = set(df_source["Component"].unique())
        relevant_tools = [t for t in tool_names if str(t).strip() in active_components]
        col_configs = []
        for t_name in relevant_tools:
            t_name_s = str(t_name).strip()
            direct_obj = self.validated.components.get(t_name_s)
            
            if direct_obj:
                current_co = int(df_source["CO"].iloc[0]) # The CO for this sheet
                max_m = 0.0
                if hasattr(direct_obj, "questions"):
                    for q in direct_obj.questions:
                        # 1. Parse the CO list for this specific question
                        # Handles strings "1, 2" or actual lists
                        raw_co_str = str(q.co_list).replace("(", "").replace(")", "").replace("'", "").replace('"', "")
                        q_cos = [int(v.strip()) for v in raw_co_str.split(",") if v.strip()]
                        
                        # 2. If this question belongs to the current CO, add its split share
                        if current_co in q_cos:
                            num_cos = len(q_cos)
                            max_m += q.max_marks / num_cos
                else:
                    # Fallback for simple components
                    max_m = getattr(direct_obj, "max_marks", 100)
                #max_m = sum(q.max_marks for q in direct_obj.questions) if hasattr(direct_obj, "questions") else getattr(direct_obj, "max_marks", 100)
                wt = direct_obj.weight
            else:
                continue
            col_configs.append({
                "name": t_name_s,
                "t_key": f"{t_name_s}({max_m})",
                "s_key": f"{t_name_s}(Wtd:{wt})"
            })
        rows = []
        for reg_no, name in student_map.items():
            regno_s = str(reg_no).strip()
            row = {"Reg. No.": regno_s, "Student Name": name}
            total_scaled, total_weight, co_status = 0.0, 0.0, STATUS_NORMAL 
            for cfg in col_configs:
                if (regno_s, cfg["name"]) in lookup.index:
                    data = lookup.loc[(regno_s, cfg["name"])]
                #data = lookup.get((regno_s, cfg["name"]))

                if data is None or data.empty:
                    row[cfg["t_key"]], row[cfg["s_key"]] = "N/A", "N/A"
                    co_status = STATUS_NA
                    continue

                total, wtd = float(data["Total"]), float(data["Wtd"])
                if bool(data["IsAbsent"]):
                    row[cfg["t_key"]], row[cfg["s_key"]] = ABSENT_SYMBOL, ABSENT_SYMBOL
                    if co_status == STATUS_NORMAL: co_status = STATUS_A
                else:
                    row[cfg["t_key"]] = round((total), 2) if total > 0 else 0.0
                    row[cfg["s_key"]] = round(wtd, 2)
                    total_scaled += wtd
                    total_weight += float(data["Weight"])

            # 4. Final CO Calculation
            if co_status == STATUS_NA or not col_configs:
                row["CO Total"], row[f"CO{current_co} 100%"] = "N/A", "N/A"
            elif co_status == STATUS_A:
                row["CO Total"], row[f"CO{current_co} 100%"] = ABSENT_SYMBOL, ABSENT_SYMBOL
            else:
                if total_weight == 0:
                    row["CO Total"], row[f"CO{current_co} 100%"] = "N/A", "N/A"
                else:
                    row["CO Total"] = round(total_scaled, 2)
                    row[f"CO{current_co} 100%"] = round((total_scaled / total_weight * 100.0), 2)
            rows.append(row)
        return pd.DataFrame(rows)
    def _build_data_package(self, df_source, student_map, tool_names, indirect_map) -> pd.DataFrame:
        

        # 1. Aggregation
        df_source = df_source.copy()
        df_source["RegNo"] = df_source["RegNo"].astype(str).str.strip()
        df_source["Component"] = df_source["Component"].astype(str).str.strip()

        g = (df_source.groupby(["RegNo", "Component"], as_index=False)
                    .agg(Wtd=("Wtd", "sum"),
                        Weight=("Weight", "sum"),
                        IsAbsent=("IsAbsent", "max")))
        
        lookup = g.set_index(["RegNo", "Component"])[["Wtd", "Weight", "IsAbsent"]].to_dict("index")

        # 2. Dynamic Column Filtering: Only include tools found in this specific CO fragment
        active_components = set(df_source["Component"].unique())
        relevant_tools = [t for t in tool_names if str(t).strip() in active_components]

        col_configs = []
        for t_name in relevant_tools:
            t_name_s = str(t_name).strip()
            direct_obj = self.validated.components.get(t_name_s)
            indirect_obj = indirect_map.get(t_name_s)
            
            if direct_obj:
                max_m = sum(q.max_marks for q in direct_obj.questions) if hasattr(direct_obj, "questions") else getattr(direct_obj, "max_marks", 100)
                wt = direct_obj.weight
            elif indirect_obj:
                max_m = LIKERT_MAX
                wt = indirect_obj.weight
            else:
                continue
            
            base_label = f"{t_name_s}\n(Max: {max_m}, Wt: {wt})"
            col_configs.append({
                "name": t_name_s,
                "t_key": f"{base_label}Total %",
                "s_key": f"{base_label}\nScaled"
            })

        # 3. Row Construction
        rows = []
        for reg_no, name in student_map.items():
            regno_s = str(reg_no).strip()
            row = {"Reg. No.": regno_s, "Student Name": name}
            total_scaled, total_weight, co_status = 0.0, 0.0, STATUS_NORMAL 

            for cfg in col_configs:
                data = lookup.get((regno_s, cfg["name"]))

                if data is None:
                    row[cfg["t_key"]], row[cfg["s_key"]] = "N/A", "N/A"
                    co_status = STATUS_NA
                    continue

                wtd, wt = float(data["Wtd"]), float(data["Weight"])
                if bool(data["IsAbsent"]):
                    row[cfg["t_key"]], row[cfg["s_key"]] = ABSENT_SYMBOL, ABSENT_SYMBOL
                    if co_status == STATUS_NORMAL: co_status = STATUS_A
                else:
                    row[cfg["t_key"]] = round((wtd / wt * 100.0), 2) if wt > 0 else 0.0
                    row[cfg["s_key"]] = round(wtd, 2)
                    total_scaled += wtd
                    total_weight += wt

            # 4. Final CO Calculation
            if co_status == STATUS_NA or not col_configs:
                row["CO Total Scaled"], row["CO Total 100%"] = "N/A", "N/A"
            elif co_status == STATUS_A:
                row["CO Total Scaled"], row["CO Total 100%"] = ABSENT_SYMBOL, ABSENT_SYMBOL
            else:
                if total_weight == 0:
                    row["CO Total Scaled"], row["CO Total 100%"] = "N/A", "N/A"
                else:
                    row["CO Total Scaled"] = round(total_scaled, 2)
                    row["CO Total 100%"] = round((total_scaled / total_weight * 100.0), 2)
            rows.append(row)

        return pd.DataFrame(rows)
    # =====================================================
    # DIRECT COLLECTION
    # =====================================================

    def _collect_direct(self) -> pd.DataFrame:
        fragments = []

        for comp_name, comp in self.validated.components.items():
            if not comp.direct:
                continue

            df = self.direct_sheets[comp_name]
            block = df.iloc[3:].reset_index(drop=True)

            regnos = block.iloc[:, 0].astype(str).str.strip().values
            co_accumulator = {}

            for q_idx, q in enumerate(comp.questions):
                cos = []
                for co_entry in q.co_list:
                    # This handles both cases: "1" or "1, 2"
                    split_values = str(co_entry).split(",")
                    for val in split_values:
                        clean_val = val.strip()
                        if clean_val: # Ensure it's not an empty string
                            cos.append(int(clean_val))
                num_cos = len(cos)
                divisor = num_cos if not comp.co_split else 1

                raw_series = block.iloc[:, q_idx + 2]
                raw_marks = pd.to_numeric(raw_series, errors="coerce").values
                q_is_absent = np.isnan(raw_marks)
                q_marks_adj = np.nan_to_num(raw_marks / divisor) 
                q_max_adj = q.max_marks / divisor
                
                for co in cos:
                    co_val = int(co)
                    if co_val not in co_accumulator:
                        co_accumulator[co_val] = {
                            "total_obtained": np.zeros_like(q_marks_adj),
                            "total_max": 0.0,
                            "absent_mask": np.ones_like(q_marks_adj, dtype=bool) 
                        }
                    
                    co_accumulator[co_val]["total_obtained"] += q_marks_adj
                    co_accumulator[co_val]["total_max"] += q_max_adj
                    # Logic: Absent only if ALL questions in this CO are 'A'
                    co_accumulator[co_val]["absent_mask"] &= q_is_absent

            for co_val, data in co_accumulator.items():
                # SAFE DIVISION: prevents DivideByZero if total_max is 0 [cite: 3, 13]
                performance_ratio = np.divide(
                    data["total_obtained"], 
                    data["total_max"], 
                    out=np.zeros_like(data["total_obtained"]), 
                    where=data["total_max"] > 0
                )
                wtd_score = performance_ratio * comp.weight
                fragments.append(pd.DataFrame({
                    "RegNo": regnos,
                    "CO": co_val,
                    "Component": comp_name,
                    "total": data["total_obtained"],
                    "Wtd": wtd_score,
                    "Weight": comp.weight, # w_jk [cite: 14]
                    "IsAbsent": data["absent_mask"]
                }))
        if not fragments:
            return pd.DataFrame()

        df = pd.concat(fragments, ignore_index=True)

        df["CO"] = df["CO"].astype("int8")
        df["Component"] = df["Component"].astype("category")

        #print(df.head())

        return df

    # =====================================================
    # INDIRECT COLLECTION
    # =====================================================

    def _collect_indirect(self) -> pd.DataFrame:
        fragments = []
        scale_range = LIKERT_MAX - LIKERT_MIN

        for tool in self.validated.indirect_tools:
            df = self.indirect_sheets[tool.name]

            regnos = df["RegNo"].astype(str).str.strip().values
            co_cols = [c for c in df.columns if str(c).upper().startswith("CO")]

            for col in co_cols:
                raw = pd.to_numeric(df[col], errors="coerce").values

                scaled = np.zeros_like(raw)

                if scale_range > 0:
                    scaled = (((raw - LIKERT_MIN) / scale_range * 100)
                              .clip(0) / 100) * tool.weight

                fragments.append(pd.DataFrame({
                    "RegNo": regnos,
                    "CO": int(str(col).upper().replace("CO", "")),
                    "Component": tool.name,
                    "Wtd": scaled,
                    "Weight": tool.weight,
                    "IsAbsent": np.isnan(raw)
                }))

        if not fragments:
            return pd.DataFrame()
        
        df = pd.concat(fragments, ignore_index=True)
        df["CO"] = df["CO"].astype("int8")
        df["Component"] = df["Component"].astype("category")
        return df