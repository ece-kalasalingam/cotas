import pandas as pd
from typing import Any, Dict, Tuple
from core.models import (
    AssessmentConfig,
    CourseMetadata,
    IndirectToolInfo,
    StudentList,
    QuestionMap,
    Question,
    ComponentInfo,
    Student,
    ValidatedSetup,
)
from core.exceptions import ValidationError


class SetupValidator:

    def __init__(
        self,
        metadata: CourseMetadata,
        config: AssessmentConfig,
        students: StudentList,
        question_map: QuestionMap,
        only_CA: bool = False,
    ):
        
        self.metadata = metadata
        self.config_df = config.dataframe.copy()
        self.students_df = students.dataframe.copy()
        self.qmap_df = question_map.dataframe.copy()
        self.only_CA = only_CA

    # =====================================================
    # PUBLIC ENTRY
    # =====================================================
    def validate(self) -> ValidatedSetup:
        config_meta = self._validate_config()
        components = self._validate_question_map(config_meta)
        students = self._validate_students()
        indirect_tools = tuple(
            IndirectToolInfo(name=comp, weight=meta.get("weight", 0.0))
            for comp, meta in config_meta.items()
            if not meta.get("direct", False)
        )

        return ValidatedSetup(
            components=components,
            students=students,
            indirect_tools=indirect_tools,
        )


    # =====================================================
    # CONFIG VALIDATION
    # =====================================================
    def _validate_config(self) -> Dict[str, dict]:
        required_cols = {
            "Component",
            "Weight (%)",
            "CIA",
            "CO_Wise_Marks_Breakup",
            "Direct",
        }

        if not required_cols.issubset(self.config_df.columns):
            missing = required_cols - set(self.config_df.columns)
            raise ValidationError(
                f"Assessment_Config missing required columns: {missing}"
            )

        self.config_df["Component"] = (
            self.config_df["Component"]
            .astype(str)
            .str.strip()
        )

        if self.config_df["Component"].duplicated().any():
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

        try:
            active["Weight (%)"] = active["Weight (%)"].astype(float)
        except Exception:
            raise ValidationError("Weight (%) must be numeric.")

        if (active["Weight (%)"] < 0).any():
            raise ValidationError("All weights must be >= 0.")

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

        components: Dict[str, Dict[str, Any]] = {}

        for _, row in active.iterrows():
            comp = row["Component"]

            components[comp] = {
                "weight": float(row["Weight (%)"]),
                "cia": row["CIA"] == "yes",
                "co_split": row["CO_Wise_Marks_Breakup"] == "yes",
                "direct": row["Direct"] == "yes",
                "questions": [],
            }

        return components

    # =====================================================
    # QUESTION MAP VALIDATION + EXTRACTION
    # =====================================================
    def _validate_question_map(
        self,
        config_meta: Dict[str, dict]
    ) -> Dict[str, ComponentInfo]:

        required_cols = {
            "Component",
            "Q_No/Rubric_Parameter",
            "Max_Marks",
            "CO",
        }
        total_cos = self.metadata.total_cos
        valid_co_set = {f"{i}" for i in range(1, total_cos + 1)}
        covered_cos = set()

        if not required_cols.issubset(self.qmap_df.columns):
            missing = required_cols - set(self.qmap_df.columns)
            raise ValidationError(
                f"Question_Map missing required columns: {missing}"
            )

        self.qmap_df["Component"] = (
            self.qmap_df["Component"]
            .astype(str)
            .str.strip()
        )

        final_components: Dict[str, ComponentInfo] = {}

        for comp, meta in config_meta.items():

            comp_q = self.qmap_df[self.qmap_df["Component"] == comp]

            if meta["direct"]:
                if comp_q.empty:
                    raise ValidationError(
                        f"{comp}: Direct assessment must appear in Question_Map."
                    )
            else:
                if not comp_q.empty:
                    raise ValidationError(
                        f"{comp}: Indirect assessment must NOT appear in Question_Map."
                    )
                final_components[comp] = ComponentInfo(
                    weight=meta["weight"],
                    cia=meta["cia"],
                    co_split=meta["co_split"],
                    direct=meta["direct"],
                    questions=tuple(),
                )
                continue

            if comp_q["Q_No/Rubric_Parameter"].duplicated().any():
                raise ValidationError(
                    f"{comp}: Duplicate question identifiers."
                )

            questions = []

            for _, row in comp_q.iterrows():

                try:
                    max_marks = float(row["Max_Marks"])
                except Exception:
                    raise ValidationError(
                        f"{comp}: Max_Marks must be numeric."
                    )

                if max_marks <= 0:
                    raise ValidationError(
                        f"{comp}: Max_Marks must be > 0."
                    )

                co_raw = str(row["CO"]).strip()
                if not co_raw:
                    raise ValidationError(
                        f"{comp}: CO cannot be empty."
                    )

                co_list = tuple(
                    x.strip().upper()
                    for x in co_raw.replace(";", ",")
                    .replace("|", ",")
                    .split(",")
                    if x.strip()
                )

                # --- Validate range ---
                for co in co_list:
                    if co not in valid_co_set:
                        raise ValidationError(
                            f"{comp}: Invalid CO '{co}'. "
                            f"Must be between CO1 and CO{total_cos}."
                        )

                if meta["direct"]:
                    covered_cos.update(co_list)
                
                if meta["co_split"] and len(co_list) != 1:
                    raise ValidationError(
                        f"{comp}: CO_Wise_Marks_Breakup = yes "
                        "requires exactly one CO per row."
                    )
                if not meta["co_split"] and len(co_list) < 1:
                    raise ValidationError(
                        f"{comp}: CO_Wise_Marks_Breakup = no "
                        "requires at least one CO per row."
                    )
                if not meta["co_split"] and len(comp_q) != 1:
                    raise ValidationError(
                        f"{comp}: CO_Wise_Marks_Breakup = no "
                        "cannot have multiple question rows."
                    )

                questions.append(
                    Question(
                        identifier=str(row["Q_No/Rubric_Parameter"]).strip(),
                        max_marks=float(row["Max_Marks"]),
                        co_list=co_list,
                    )
                )

            final_components[comp] = ComponentInfo(
                weight=meta["weight"],
                cia=meta["cia"],
                co_split=meta["co_split"],
                direct=meta["direct"],
                questions=tuple(questions),
            )
        # --- Coverage Check ---
        if any(meta["direct"] for meta in config_meta.values()):
            missing_cos = valid_co_set - covered_cos

            if missing_cos:
                raise ValidationError(
                    f"CO(s) not covered in Question_Map: "
                    f"{sorted(missing_cos)}"
                )

        return final_components

    # =====================================================
    # STUDENT VALIDATION + EXTRACTION
    # =====================================================
    def _validate_students(self) -> Tuple[Student, ...]:

        required_cols = {"RegNo", "Student_Name"}

        if not required_cols.issubset(self.students_df.columns):
            missing = required_cols - set(self.students_df.columns)
            raise ValidationError(
                f"Students sheet missing required columns: {missing}"
            )

        self.students_df["RegNo"] = (
            self.students_df["RegNo"]
            .astype(str)
            .str.strip()
        )

        if self.students_df["RegNo"].duplicated().any():
            raise ValidationError("Duplicate RegNo entries found.")

        if self.students_df["RegNo"].eq("").any():
            raise ValidationError("Empty RegNo found.")

        students = tuple(
            Student(
                reg_no=row["RegNo"],
                name=str(row["Student_Name"]).strip(),
            )
            for _, row in self.students_df.iterrows()
        )

        return students
