from __future__ import annotations

import hashlib
import re
from collections import defaultdict
from typing import Any, Dict, List, Tuple

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from scripts.constants import (
    DEFAULT_MAX_COL_WIDTH,
    DEFAULT_MIN_COL_WIDTH,
    SYSTEM_HASH_SHEET_NAME,
    WIDTH_PADDING,
)
from scripts.exceptions import ValidationError
from scripts.utils import calculate_visual_width


_INVALID_SHEET_CHARS = set("[]:*?/\\")


def _normalize_direct_flag(value: Any) -> str:
    token = "" if value is None else str(value).strip().lower()
    if token in {"yes", "y", "true", "1", "direct", "d"}:
        return "YES"
    if token in {"no", "n", "false", "0", "indirect", "i"}:
        return "NO"
    return ""


def _parse_co_numbers(raw: Any) -> List[int]:
    s = "" if raw is None else str(raw).strip()
    if not s:
        return []

    out: List[int] = []
    for token in re.split(r"[;,|]+", s.replace(" ", "").upper().replace("CO", "")):
        if not token:
            continue
        try:
            out.append(int(float(token)))
        except Exception:
            continue
    return out


def _sanitize_co_text(raw: Any) -> str:
    nums = _parse_co_numbers(raw)
    if nums:
        return ",".join(str(x) for x in nums)
    return "" if raw is None else str(raw).strip()


def _check_sheet_name(name: str) -> None:
    if not name:
        raise ValidationError("Sheet name cannot be blank.")
    if len(name) > 31:
        raise ValidationError(f"Sheet name too long for Excel (max 31): {name}")
    if any(ch in _INVALID_SHEET_CHARS for ch in name):
        raise ValidationError(f"Sheet name has invalid character(s): {name}")


def _sheet_name_for_component(component_name: Any) -> str:
    return "" if component_name is None else str(component_name).strip()

def _set_col_widths(ws, col_widths: Dict[int, int]) -> None:
    for col, width in col_widths.items():
        ws.set_column(col, col, min(max(width, DEFAULT_MIN_COL_WIDTH), DEFAULT_MAX_COL_WIDTH))


def _track_width(col_widths: Dict[int, int], col: int, val: Any) -> None:
    w = calculate_visual_width(val) + WIDTH_PADDING
    if w > col_widths.get(col, 0):
        col_widths[col] = w


def _build_metadata_rows(course_details: Dict[str, Any], setup_store: Dict[str, List[List[Any]]]) -> List[Tuple[str, Any]]:
    rows = setup_store.get("Course_Metadata", [])
    pairs: List[Tuple[str, Any]] = []

    if rows:
        for row in rows:
            if not row or len(row) < 2:
                continue
            key = "" if row[0] is None else str(row[0]).strip()
            if not key:
                continue
            pairs.append((key, row[1]))
    elif course_details:
        for k, v in course_details.items():
            pairs.append((str(k), v))

    return pairs


def _build_assessment_rows(setup_store: Dict[str, List[List[Any]]]) -> List[List[Any]]:
    rows = setup_store.get("Assessment_Config", [])
    out: List[List[Any]] = []
    for row in rows:
        if not row:
            continue
        padded = list(row[:5]) + [""] * max(0, 5 - len(row))
        if all(v is None or str(v).strip() == "" for v in padded):
            continue
        out.append(padded[:5])
    return out


def generate_marks_template_from_setup(
    setup_store: Dict[str, List[List[Any]]],
    output_path: str,
    course_details: Dict[str, Any] | None = None,
) -> str:
    if not setup_store:
        raise ValidationError("Setup data is empty. Upload and validate setup file first.")

    assess_rows = setup_store.get("Assessment_Config", [])
    qmap_rows = setup_store.get("Question_Map", [])
    students_rows = setup_store.get("Students", [])

    students: List[Tuple[str, str]] = []
    for row in students_rows:
        if not row:
            continue
        reg = "" if len(row) < 1 or row[0] is None else str(row[0]).strip()
        name = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
        if reg:
            students.append((reg, name))

    if not students:
        raise ValidationError("No students found in setup file.")

    direct_components: List[str] = []
    indirect_tools: List[str] = []
    for row in assess_rows:
        if len(row) < 5:
            continue
        comp_name = "" if row[0] is None else str(row[0]).strip()
        if not comp_name:
            continue
        flag = _normalize_direct_flag(row[4])
        if flag == "YES":
            direct_components.append(comp_name)
        elif flag == "NO":
            indirect_tools.append(comp_name)

    if not direct_components:
        raise ValidationError("No direct components found in setup file.")

    q_by_comp: Dict[str, List[Tuple[str, Any, str]]] = defaultdict(list)
    all_cos = set()

    for row in qmap_rows:
        if len(row) < 4:
            continue
        comp = "" if row[0] is None else str(row[0]).strip()
        qid = "" if row[1] is None else str(row[1]).strip()
        max_marks = row[2]
        co_text = _sanitize_co_text(row[3])
        if comp and qid:
            q_by_comp[comp].append((qid, max_marks, co_text))
        for n in _parse_co_numbers(row[3]):
            all_cos.add(n)

    if not all_cos:
        all_cos = {1}

    workbook = xlsxwriter.Workbook(output_path)

    f_header = workbook.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "align": "center"})
    f_locked = workbook.add_format({"border": 1, "bg_color": "#F2F2F2"})
    f_unlocked = workbook.add_format({"border": 1, "align": "center"})

    # Course_Info
    ws_info = workbook.add_worksheet("Course_Info")
    info_widths: Dict[int, int] = {}
    ws_info.write(0, 0, "Field", f_header)
    ws_info.write(0, 1, "Value", f_header)
    _track_width(info_widths, 0, "Field")
    _track_width(info_widths, 1, "Value")

    info_rows = _build_metadata_rows(course_details or {}, setup_store)
    for r_idx, (field, value) in enumerate(info_rows, start=1):
        ws_info.write(r_idx, 0, field)
        ws_info.write(r_idx, 1, value)
        _track_width(info_widths, 0, field)
        _track_width(info_widths, 1, value)

    _set_col_widths(ws_info, info_widths)

    # Assessment_Config (copied from setup for reference)
    ws_assess = workbook.add_worksheet("Assessment_Config")
    assess_widths: Dict[int, int] = {}
    assess_header = ["Component", "Weight (%)", "CIA", "CO_Wise_Marks_Breakup", "Direct"]

    for c_idx, val in enumerate(assess_header):
        ws_assess.write(0, c_idx, val, f_header)
        _track_width(assess_widths, c_idx, val)

    copied_assess_rows = _build_assessment_rows(setup_store)
    for r_idx, row in enumerate(copied_assess_rows, start=1):
        for c_idx, val in enumerate(row):
            ws_assess.write(r_idx, c_idx, val, f_locked)
            _track_width(assess_widths, c_idx, val)

    _set_col_widths(ws_assess, assess_widths)

    # Direct component sheets
    for comp_name in direct_components:
        sheet_name = _sheet_name_for_component(comp_name)
        _check_sheet_name(sheet_name)
        questions = q_by_comp.get(comp_name, [])
        if not questions:
            raise ValidationError(f"No question map entries found for direct component: {comp_name}")

        ws = workbook.add_worksheet(sheet_name)
        widths: Dict[int, int] = {}

        q_ids = [q[0] for q in questions]
        co_vals = [q[2] for q in questions]

        max_vals: List[float] = []
        for _, max_mark, _ in questions:
            try:
                max_vals.append(float(max_mark))
            except Exception:
                max_vals.append(0.0)

        header = ["RegNo", "Student_Name"] + q_ids + ["Total"]
        co_row = ["CO", ""] + co_vals + [""]
        max_row = ["Max", ""] + max_vals + [sum(max_vals)]

        for c_idx, val in enumerate(header):
            ws.write(0, c_idx, val, f_header)
            _track_width(widths, c_idx, val)

        for c_idx, val in enumerate(co_row):
            ws.write(1, c_idx, val, f_header)
            _track_width(widths, c_idx, val)

        for c_idx, val in enumerate(max_row):
            ws.write(2, c_idx, val, f_header)
            _track_width(widths, c_idx, val)

        end_q_col = 1 + len(q_ids) + 1
        for r_idx, (reg, stu_name) in enumerate(students, start=3):
            ws.write(r_idx, 0, reg, f_locked)
            ws.write(r_idx, 1, stu_name, f_locked)
            _track_width(widths, 0, reg)
            _track_width(widths, 1, stu_name)

            for c_idx in range(2, end_q_col):
                ws.write(r_idx, c_idx, "", f_unlocked)

            c_from = xl_col_to_name(2)
            c_to = xl_col_to_name(end_q_col - 1)
            ws.write_formula(r_idx, end_q_col, f"=SUM({c_from}{r_idx+1}:{c_to}{r_idx+1})", f_locked)
            _track_width(widths, end_q_col, "Total")

        ws.freeze_panes(3, 2)
        _set_col_widths(ws, widths)

    # Indirect sheets
    co_cols = [f"CO{n}" for n in sorted(all_cos)]
    for tool_name in indirect_tools:
        sheet_name = _sheet_name_for_component(tool_name)
        _check_sheet_name(sheet_name)

        ws = workbook.add_worksheet(sheet_name)
        widths: Dict[int, int] = {}

        header = ["RegNo", "Student_Name"] + co_cols
        for c_idx, val in enumerate(header):
            ws.write(0, c_idx, val, f_header)
            _track_width(widths, c_idx, val)

        for r_idx, (reg, stu_name) in enumerate(students, start=1):
            ws.write(r_idx, 0, reg, f_locked)
            ws.write(r_idx, 1, stu_name, f_locked)
            _track_width(widths, 0, reg)
            _track_width(widths, 1, stu_name)
            for c_idx in range(2, len(header)):
                ws.write(r_idx, c_idx, "", f_unlocked)

        ws.freeze_panes(1, 2)
        _set_col_widths(ws, widths)

    # __SYSTEM_HASH__
    hasher = hashlib.sha256()
    for name in sorted(direct_components):
        hasher.update(name.encode("utf-8"))
        for qid, max_mark, co_text in q_by_comp.get(name, []):
            hasher.update(str(qid).encode("utf-8"))
            hasher.update(str(max_mark).encode("utf-8"))
            hasher.update(str(co_text).encode("utf-8"))
    for tool in sorted(indirect_tools):
        hasher.update(tool.encode("utf-8"))
    for reg, _ in sorted(students):
        hasher.update(reg.encode("utf-8"))

    ws_hash = workbook.add_worksheet(SYSTEM_HASH_SHEET_NAME)
    ws_hash.write(0, 0, hasher.hexdigest())
    ws_hash.hide()

    workbook.close()
    return output_path

