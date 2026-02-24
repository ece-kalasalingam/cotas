import pandas as pd
from core.models import (
    IndirectToolInfo,
    Question,
    ComponentInfo,
    Student,
    ValidatedSetup,
)
from core.exceptions import ValidationError


class SetupValidator:
    """
    Pure setup validator.
    No Excel I/O.
    Operates only on DataFrames.
    """

    def __init__(self, metadata, config, students, question_map, only_CA=False):
        self.metadata = metadata
        self.config_df = config.dataframe.copy()
        self.students_df = students.dataframe.copy()
        self.qmap_df = question_map.dataframe.copy()
        self.only_CA = only_CA

    # =====================================================
    # PUBLIC ENTRY
    # =====================================================

    def validate(self) -> ValidatedSetup:
        components_meta = self._validate_config()
        components = self._validate_question_map(components_meta)
        students = self._validate_students()
        indirect_tools = self._build_indirect_tools(components_meta)

        return ValidatedSetup(
            components=components,
            students=students,
            indirect_tools=indirect_tools,
        )

    # =====================================================
    # CONFIG VALIDATION
    # =====================================================

    def _validate_config(self):
        required = {"Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct"}
        if not required.issubset(self.config_df.columns):
            raise ValidationError(f"Assessment_Config missing columns: {required - set(self.config_df.columns)}")

        df = self.config_df.copy()
        df["Component"] = df["Component"].astype(str).str.strip()

        if df["Component"].duplicated().any():
            raise ValidationError("Duplicate components in Assessment_Config.")
        for col in ["CIA", "CO_Wise_Marks_Breakup", "Direct"]:
            self.config_df[col] = (
                self.config_df[col]
                .astype(str)
                .str.strip()
                .str.lower()
            )

            if not self.config_df[col].isin(["yes", "no"]).all():
                raise ValidationError(
                    f"{col} column must contain only 'yes' or 'no'."
                )

        if self.only_CA:
            active = self.config_df[self.config_df["CIA"] == "yes"]
            if active.empty:
                raise ValidationError(
                    "No components marked CIA = yes."
                )
        else:
            active = self.config_df.copy()

        active["Weight (%)"] = pd.to_numeric(active["Weight (%)"], errors="raise")

        if (active["Weight (%)"] < 0).any():
            raise ValidationError("Weights must be >= 0.")
        
        active["DirectFlag"] = active["Direct"].str.strip().str.lower()  == "yes"
        direct_rows = active[active["DirectFlag"]]
        if direct_rows.empty:
            raise ValidationError("At least one DIRECT component is required.")

        if not self.only_CA:
            total_weight = direct_rows["Weight (%)"].sum()
            if abs(total_weight - 100) > 0.01:
                raise ValidationError(
                    f"Total weight must equal 100. Currently: {total_weight}"
                )
        indirect_rows = active[~active["DirectFlag"]]
        if not indirect_rows.empty:
            indirect_weight = indirect_rows["Weight (%)"].sum()

            if abs(indirect_weight - 100) > 0.01:
                raise ValidationError(
                    f"Total INDIRECT tool weight must equal 100. Currently: {indirect_weight}"
                )

        active["Direct"] = active["Direct"].astype(str).str.lower().str.strip()
        active["CIA"] = active["CIA"].astype(str).str.lower().str.strip()
        active["CO_Wise_Marks_Breakup"] = active["CO_Wise_Marks_Breakup"].astype(str).str.lower().str.strip()

        components = {}

        for _, row in active.iterrows():
            components[row["Component"]] = {
                "weight": float(row["Weight (%)"]),
                "cia": row["CIA"] == "yes",
                "co_split": row["CO_Wise_Marks_Breakup"] == "yes",
                "direct": row["Direct"] == "yes",
            }

        return components

    # =====================================================
    # QUESTION MAP VALIDATION
    # =====================================================

    def _validate_question_map(self, config_meta):
        required = {"Component", "Q_No/Rubric_Parameter", "Max_Marks", "CO"}
        if not required.issubset(self.qmap_df.columns):
            raise ValidationError(f"Question_Map missing columns: {required - set(self.qmap_df.columns)}")

        total_cos = self.metadata.total_cos
        valid_cos = {str(i) for i in range(1, total_cos + 1)}
        covered_cos = set()

        final_components = {}

        for comp_name, meta in config_meta.items():
            comp_q = self.qmap_df[self.qmap_df["Component"] == comp_name]

            if meta["direct"] and comp_q.empty:
                raise ValidationError(f"{comp_name}: Direct component missing in Question_Map.")
            if not meta["direct"] and not comp_q.empty:
                raise ValidationError(f"{comp_name}: Indirect assessment must NOT appear in Question_Map.")

            if comp_q["Q_No/Rubric_Parameter"].duplicated().any():
                raise ValidationError( f"{comp_name}: Duplicate question identifiers.")

            questions = []

            for _, row in comp_q.iterrows():
                try:
                    max_marks = float(row["Max_Marks"])
                except Exception:
                    raise ValidationError(
                        f"{comp_name}: Max_Marks must be numeric."
                    )
                if max_marks <= 0:
                    raise ValidationError(f"{comp_name}: Max_Marks must be > 0.")

                co_raw = str(row["CO"]).strip()
                if not co_raw:
                    raise ValidationError(
                        f"{comp_name}: CO cannot be empty."
                    )

                co_list = tuple(
                    x.strip()
                    for x in co_raw.replace(";", ",").replace("|", ",").split(",")
                    if x.strip()
                )

                for co in co_list:
                    if co not in valid_cos:
                        raise ValidationError(f"{comp_name}: Invalid CO {co}")
                    covered_cos.add(co)
                if meta["co_split"] and len(co_list) != 1:
                    raise ValidationError(
                        f"{comp_name}: CO_Wise_Marks_Breakup = yes "
                        "requires exactly one CO per row."
                    )
                if not meta["co_split"] and len(co_list) < 1:
                    raise ValidationError(
                        f"{comp_name}: CO_Wise_Marks_Breakup = no "
                        "requires at least one CO per row."
                    )
                if not meta["co_split"] and len(comp_q) != 1:
                    raise ValidationError(
                        f"{comp_name}: CO_Wise_Marks_Breakup = no "
                        "cannot have multiple question rows."
                    )

                questions.append(
                    Question(
                        tool_name=comp_name,
                        identifier=str(row["Q_No/Rubric_Parameter"]).strip(),
                        max_marks=max_marks,
                        co_list=co_list,
                    )
                )

            final_components[comp_name] = ComponentInfo(
                weight=meta["weight"],
                cia=meta["cia"],
                co_split=meta["co_split"],
                direct=meta["direct"],
                questions=tuple(questions),
            )

        if any(meta["direct"] for meta in config_meta.values()):
            missing_cos = valid_cos - covered_cos

            if missing_cos:
                raise ValidationError(
                    f"CO(s) not covered in Question_Map: "
                    f"{sorted(missing_cos)}"
                )


        return final_components

    # =====================================================
    # STUDENTS VALIDATION
    # =====================================================

    def _validate_students(self):
        required = {"RegNo", "Student_Name"}
        if not required.issubset(self.students_df.columns):
            missing = required - set(self.students_df.columns)
            raise ValidationError(f"Students sheet missing required columns: {sorted(missing)}")

        df = self.students_df.copy()
        df["RegNo"] = df["RegNo"].astype(str).str.strip()

        if len(df) < 1:
            raise ValidationError("Students sheet must contain at least one student.")

        if df["RegNo"].duplicated().any():
            raise ValidationError("Duplicate RegNo entries.")
        if self.students_df["RegNo"].eq("").any():
            raise ValidationError("Empty RegNo found.")

        return tuple(
            Student(reg_no=row["RegNo"], name=str(row["Student_Name"]).strip())
            for _, row in df.iterrows()
        )

    # =====================================================
    # INDIRECT TOOL BUILD
    # =====================================================

    def _build_indirect_tools(self, config_meta):
        tools = []

        for name, meta in config_meta.items():
            if not meta["direct"]:
                tools.append(
                    IndirectToolInfo(name=name, weight=meta["weight"])
                )

        return tuple(tools)