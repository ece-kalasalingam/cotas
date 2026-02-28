from __future__ import annotations

from collections import defaultdict
from typing import Any, Dict, List

import openpyxl

from scripts.constants import ID_COURSE_SETUP, SYSTEM_HASH_SHEET_NAME


class UniversalEngine:
    def __init__(self, bp_registry: Dict[str, Any]):
        self.bp_registry = bp_registry
        self.bp = None
        self.data_store: Dict[str, List[List[Any]]] = {}
        self._headers: Dict[str, List[str]] = {}
        self.errors: List[str] = []

    def load_from_file(self, filepath: str) -> bool:
        self.errors = []
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        except Exception as exc:
            self.errors.append(f"Failed to open workbook: {exc}")
            return False

        try:
            type_id = self._detect_type_from_workbook(wb) or ID_COURSE_SETUP

            self.bp = self.bp_registry.get(type_id)
            if not self.bp:
                self.errors.append(f"Unsupported workbook type: {type_id}")
                return False

            if type_id == ID_COURSE_SETUP:
                ok = self._load_setup_workbook(wb)
                if not ok:
                    return False
                return self._run_business_rules()

            self.errors.append(f"No loader implemented for workbook type: {type_id}")
            return False
        finally:
            wb.close()

    def load_with_external(self, primary_path: str, external_path: str) -> bool:
        ext = UniversalEngine(self.bp_registry)
        if not ext.load_from_file(external_path):
            self.errors = ext.errors.copy()
            return False
        return self.load_with_external_engine(primary_path, ext)

    def load_with_external_engine(self, primary_path: str, external_engine: "UniversalEngine") -> bool:
        self.errors = []
        try:
            wb = openpyxl.load_workbook(primary_path, read_only=True, data_only=True)
        except Exception as exc:
            self.errors.append(f"Failed to read marks workbook: {exc}")
            return False

        try:
            if not external_engine.data_store:
                self.errors.append("Setup context is empty. Load setup first.")
                return False

            setup = external_engine.data_store
            expected = self._derive_expected_marks_layout(setup)

            actual_name_map = {str(name).strip().lower(): name for name in wb.sheetnames}

            missing = [name for name in expected["required"] if name.lower() not in actual_name_map]
            if missing:
                self.errors.append("Missing required marks sheet(s): " + ", ".join(missing))
                return False

            self.bp = type("DynamicBP", (), {"type_id": "MARKS_ENTRY_V1"})()

            for comp_name in expected["direct_components"]:
                ws_name = actual_name_map.get(comp_name.lower())
                if not ws_name:
                    self.errors.append(f"Missing direct sheet: {comp_name}")
                    return False

                ws = wb[ws_name]
                q_headers = expected["questions_by_component"].get(comp_name, [])
                expected_header = ["RegNo", "Student_Name"] + q_headers + ["Total"]
                actual_header = self._read_header(ws, len(expected_header))

                if actual_header != expected_header:
                    self.errors.append(f"{comp_name}: header mismatch.")
                    return False

            for tool_name in expected["indirect_tools"]:
                expected_sheet = f"{tool_name}_Indirect"
                ws_name = actual_name_map.get(expected_sheet.lower())
                if not ws_name:
                    self.errors.append(f"Missing indirect sheet: {expected_sheet}")
                    return False

                ws = wb[ws_name]
                expected_header = ["RegNo", "Student_Name"] + [f"CO{i}" for i in expected["co_numbers"]]
                actual_header = self._read_header(ws, len(expected_header))

                if actual_header != expected_header:
                    self.errors.append(f"{expected_sheet}: header mismatch.")
                    return False

            return True
        finally:
            wb.close()

    def get_col_idx(self, sheet_name: str, col_name: str) -> int:
        headers = self._headers.get(sheet_name, [])
        target = "" if col_name is None else str(col_name).strip().lower()
        for i, val in enumerate(headers):
            if str(val).strip().lower() == target:
                return i
        return -1

    def _detect_type(self, filepath: str) -> str | None:
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            try:
                return self._detect_type_from_workbook(wb)
            finally:
                wb.close()
        except Exception:
            return None

    def _detect_type_from_workbook(self, wb) -> str | None:
        if SYSTEM_HASH_SHEET_NAME in wb.sheetnames:
            ws = wb[SYSTEM_HASH_SHEET_NAME]
            val = ws.cell(row=1, column=1).value
            if val:
                return str(val).strip()
        return None

    def _load_setup_workbook(self, wb) -> bool:
        self.data_store = {}
        self._headers = {}

        for sheet_schema in self.bp.sheets:
            if sheet_schema.name not in wb.sheetnames:
                self.errors.append(f"Missing required sheet: {sheet_schema.name}")
                return False

            ws = wb[sheet_schema.name]
            expected_header = [str(v).strip() for v in sheet_schema.header_matrix[-1]]
            actual_header = self._read_header(ws, len(expected_header))

            if actual_header != expected_header:
                self.errors.append(
                    f"Header mismatch in '{sheet_schema.name}'. Expected {expected_header}, found {actual_header}"
                )
                return False

            self._headers[sheet_schema.name] = expected_header

            rows: List[List[Any]] = []
            for r in ws.iter_rows(min_row=2, max_col=len(expected_header), values_only=True):
                if all(v is None or str(v).strip() == "" for v in r):
                    continue
                rows.append(list(r))
            self.data_store[sheet_schema.name] = rows

        return True

    def _read_header(self, ws, width: int) -> List[str]:
        values = [ws.cell(row=1, column=c).value for c in range(1, width + 1)]
        return ["" if v is None else str(v).strip() for v in values]

    def _run_business_rules(self) -> bool:
        problems = []
        for rule in self.bp.business_rules:
            errs = rule.run(self)
            problems.extend([f"Rule {rule.rule_id}: {e}" for e in errs])

        if problems:
            self.errors.extend(problems)
            return False
        return True

    def _derive_expected_marks_layout(self, setup_store: Dict[str, List[List[Any]]]) -> Dict[str, Any]:
        assess_rows = setup_store.get("Assessment_Config", [])
        qmap_rows = setup_store.get("Question_Map", [])

        direct_components: List[str] = []
        indirect_tools: List[str] = []
        for row in assess_rows:
            if len(row) < 5:
                continue
            name = "" if row[0] is None else str(row[0]).strip()
            direct_flag = "" if row[4] is None else str(row[4]).strip().upper()
            if not name:
                continue
            if direct_flag == "YES":
                direct_components.append(name)
            elif direct_flag == "NO":
                indirect_tools.append(name)

        q_by_comp: Dict[str, List[str]] = defaultdict(list)
        co_numbers = set()

        for row in qmap_rows:
            if len(row) < 4:
                continue
            comp = "" if row[0] is None else str(row[0]).strip()
            qid = "" if row[1] is None else str(row[1]).strip()
            co_raw = "" if row[3] is None else str(row[3]).strip()

            if comp and qid:
                q_by_comp[comp].append(qid)

            for token in co_raw.replace("CO", "").replace(" ", "").replace(";", ",").split(","):
                if not token:
                    continue
                try:
                    co_numbers.add(int(float(token)))
                except Exception:
                    continue

        if not co_numbers:
            co_numbers = {1}

        required = list(direct_components) + [f"{name}_Indirect" for name in indirect_tools]

        return {
            "required": required,
            "direct_components": direct_components,
            "indirect_tools": indirect_tools,
            "questions_by_component": dict(q_by_comp),
            "co_numbers": sorted(co_numbers),
        }
