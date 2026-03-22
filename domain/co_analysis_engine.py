"""Domain engine for CO Analysis workbook validation and generation."""

from __future__ import annotations

from pathlib import Path
from typing import Callable

from common.constants import (
    ALLOW_FILTER,
    ALLOW_SELECT_LOCKED,
    ALLOW_SELECT_UNLOCKED,
    ALLOW_SORT,
    CO_REPORT_HEADER_REG_NO,
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SECTION_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_SHEET,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    STUDENTS_HEADERS,
    STUDENTS_SHEET,
    SYSTEM_HASH_SHEET,
    SYSTEM_LAYOUT_SHEET,
)
from common.jobs import CancellationToken
from common.utils import coerce_excel_number, normalize
from common.workbook_secret import ensure_workbook_secret_policy, get_workbook_password
from domain.coordinator_engine import _path_key

_SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
_COURSE_METADATA_DUPLICATE_FIELDS = (
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_SECTION_KEY,
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
)
_COURSE_METADATA_COURSE_NAME_KEY = "course_name"
_ORDERED_METADATA_FIELDS: tuple[tuple[str, str], ...] = (
    (COURSE_METADATA_COURSE_CODE_KEY, "Course Code"),
    (_COURSE_METADATA_COURSE_NAME_KEY, "Course Name"),
    (COURSE_METADATA_TOTAL_OUTCOMES_KEY, "Total Outcomes"),
    (COURSE_METADATA_ACADEMIC_YEAR_KEY, "Academic Year"),
    (COURSE_METADATA_SEMESTER_KEY, "Semester"),
)


def analyze_uploaded_workbooks(
    candidate_paths: list[str],
    *,
    existing_paths: list[str],
    validate_uploaded_source_workbook: Callable[[str | Path], None],
    token: CancellationToken | None = None,
) -> dict[str, object]:
    existing_keys = {_path_key(Path(path)) for path in existing_paths}
    seen_metadata_signatures = {
        signature
        for signature in (extract_course_metadata_signature(Path(path)) for path in existing_paths)
        if signature is not None
    }
    added_paths: list[str] = []
    duplicates = 0
    invalid = 0
    validated_candidates: list[tuple[Path, str, tuple[str, ...] | None]] = []
    for raw_path in candidate_paths:
        if token is not None:
            token.raise_if_cancelled()
        path = Path(raw_path)
        suffix = path.suffix.lower()
        if suffix not in _SUPPORTED_EXTENSIONS or not path.exists():
            invalid += 1
            continue
        key = _path_key(path)
        try:
            validate_uploaded_source_workbook(path)
        except Exception:
            invalid += 1
            continue
        metadata_signature = extract_course_metadata_signature(path)
        validated_candidates.append((path, key, metadata_signature))

    path_counts: dict[str, int] = {}
    metadata_counts: dict[tuple[str, ...], int] = {}
    for _path, key, metadata_signature in validated_candidates:
        path_counts[key] = path_counts.get(key, 0) + 1
        if metadata_signature is not None:
            metadata_counts[metadata_signature] = metadata_counts.get(metadata_signature, 0) + 1

    for path, key, metadata_signature in validated_candidates:
        is_duplicate = False
        if key in existing_keys:
            is_duplicate = True
        elif metadata_signature is not None and metadata_signature in seen_metadata_signatures:
            is_duplicate = True
        elif path_counts.get(key, 0) > 1:
            is_duplicate = True
        elif metadata_signature is not None and metadata_counts.get(metadata_signature, 0) > 1:
            is_duplicate = True
        if is_duplicate:
            duplicates += 1
            continue
        added_paths.append(str(path))
    return {
        "added": added_paths,
        "duplicates": duplicates,
        "invalid": invalid,
        "ignored": duplicates + invalid,
    }


def extract_course_metadata_signature(path: Path) -> tuple[str, ...] | None:
    try:
        import openpyxl
    except Exception:
        return None
    try:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    try:
        if COURSE_METADATA_SHEET not in workbook.sheetnames:
            return None
        sheet = workbook[COURSE_METADATA_SHEET]
        field_header = normalize(sheet.cell(row=1, column=1).value)
        value_header = normalize(sheet.cell(row=1, column=2).value)
        if field_header != normalize(COURSE_METADATA_HEADERS[0]):
            return None
        if value_header != normalize(COURSE_METADATA_HEADERS[1]):
            return None
        metadata: dict[str, str] = {}
        row = 2
        while True:
            field_raw = sheet.cell(row=row, column=1).value
            value_raw = sheet.cell(row=row, column=2).value
            if normalize(field_raw) == "" and normalize(value_raw) == "":
                break
            field_key = normalize(field_raw)
            if field_key:
                metadata[field_key] = str(value_raw).strip() if value_raw is not None else ""
            row += 1
        signature = tuple(metadata.get(normalize(field_name), "").strip() for field_name in _COURSE_METADATA_DUPLICATE_FIELDS)
        return signature if any(signature) else None
    finally:
        workbook.close()


def extract_course_metadata_and_students(path: Path) -> tuple[set[str], dict[str, str]]:
    try:
        import openpyxl
    except Exception:
        return set(), {}
    try:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return set(), {}
    unique_students: set[str] = set()
    metadata_map: dict[str, str] = {}
    try:
        if COURSE_METADATA_SHEET in workbook.sheetnames:
            metadata_sheet = workbook[COURSE_METADATA_SHEET]
            row = 2
            while True:
                key_raw = metadata_sheet.cell(row=row, column=1).value
                value_raw = metadata_sheet.cell(row=row, column=2).value
                if normalize(key_raw) == "" and normalize(value_raw) == "":
                    break
                key_text = str(key_raw).strip() if key_raw is not None else ""
                if key_text:
                    coerced = coerce_excel_number(value_raw) if value_raw is not None else None
                    metadata_map[normalize(key_text)] = str(coerced).strip() if coerced is not None else ""
                row += 1
        if STUDENTS_SHEET in workbook.sheetnames:
            students_sheet = workbook[STUDENTS_SHEET]
            header_map: dict[str, int] = {}
            max_col = int(students_sheet.max_column)
            for col in range(1, max_col + 1):
                key = normalize(students_sheet.cell(row=1, column=col).value)
                if key and key not in header_map:
                    header_map[key] = col
            reg_col = header_map.get(normalize(STUDENTS_HEADERS[0])) or header_map.get(normalize(CO_REPORT_HEADER_REG_NO))
            if reg_col is not None:
                max_row = int(students_sheet.max_row)
                for row in range(2, max_row + 1):
                    reg_raw = students_sheet.cell(row=row, column=reg_col).value
                    coerced = coerce_excel_number(reg_raw)
                    reg_text = str(coerced).strip() if coerced is not None else ""
                    if reg_text:
                        unique_students.add(normalize(reg_text))
        if not unique_students:
            unique_students = _extract_students_from_report_sheets(workbook)
    finally:
        workbook.close()
    return unique_students, metadata_map


def generate_co_analysis_workbook(
    source_paths: list[Path],
    output_path: Path,
    *,
    token: CancellationToken | None = None,
) -> Path:
    extracted: list[tuple[set[str], dict[str, str]]] = []
    for path in source_paths:
        if token is not None:
            token.raise_if_cancelled()
        extracted.append(extract_course_metadata_and_students(path))
    return build_co_analysis_workbook(output_path, extracted)


def build_co_analysis_workbook(
    output_path: Path,
    extracted: list[tuple[set[str], dict[str, str]]],
) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    workbook = Workbook()
    course_metadata_sheet = workbook.active
    if course_metadata_sheet is None:
        raise RuntimeError("Failed to create Course Metadata worksheet.")
    course_metadata_sheet.title = COURSE_METADATA_SHEET
    course_metadata_sheet.cell(row=1, column=1, value=COURSE_METADATA_HEADERS[0])
    course_metadata_sheet.cell(row=1, column=2, value=COURSE_METADATA_HEADERS[1])
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAD3")
    header_font = Font(bold=True)
    thin = Side(border_style="thin", color="000000")
    header_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    body_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in (1, 2):
        header_cell = course_metadata_sheet.cell(row=1, column=col)
        header_cell.fill = header_fill
        header_cell.font = header_font
        header_cell.border = header_border
        header_cell.alignment = Alignment(horizontal="center", vertical="center")

    row_cursor = 2
    max_col1_len = max(len(COURSE_METADATA_HEADERS[0]), len("Students On-Roll"))
    max_col2_len = len(COURSE_METADATA_HEADERS[1])
    seen_dedup_values: dict[str, set[str]] = {}
    for unique_students, metadata_map in extracted:
        for field_key_raw, label in _ORDERED_METADATA_FIELDS:
            field_key = normalize(field_key_raw)
            field_value = metadata_map.get(field_key, "")
            if not field_value:
                continue
            value_key = normalize(field_value)
            seen_values = seen_dedup_values.setdefault(field_key, set())
            if value_key in seen_values:
                continue
            seen_values.add(value_key)
            course_metadata_sheet.cell(row=row_cursor, column=1, value=label)
            course_metadata_sheet.cell(row=row_cursor, column=2, value=field_value)
            course_metadata_sheet.cell(row=row_cursor, column=1).border = body_border
            course_metadata_sheet.cell(row=row_cursor, column=2).border = body_border
            max_col1_len = max(max_col1_len, len(str(label)))
            max_col2_len = max(max_col2_len, len(str(field_value)))
            row_cursor += 1

        section_value = metadata_map.get(normalize(COURSE_METADATA_SECTION_KEY), "")
        course_metadata_sheet.cell(row=row_cursor, column=1, value="Section")
        course_metadata_sheet.cell(row=row_cursor, column=2, value=section_value)
        course_metadata_sheet.cell(row=row_cursor, column=1).border = body_border
        course_metadata_sheet.cell(row=row_cursor, column=2).border = body_border
        max_col1_len = max(max_col1_len, len("Section"))
        max_col2_len = max(max_col2_len, len(str(section_value)))
        row_cursor += 1

        students_on_roll = len(unique_students)
        course_metadata_sheet.cell(row=row_cursor, column=1, value="Students On-Roll")
        course_metadata_sheet.cell(row=row_cursor, column=2, value=students_on_roll)
        course_metadata_sheet.cell(row=row_cursor, column=1).border = body_border
        course_metadata_sheet.cell(row=row_cursor, column=2).border = body_border
        max_col1_len = max(max_col1_len, len("Students On-Roll"))
        max_col2_len = max(max_col2_len, len(str(students_on_roll)))
        row_cursor += 2

    course_metadata_sheet.column_dimensions["A"].width = min(max(18, max_col1_len + 2), 48)
    course_metadata_sheet.column_dimensions["B"].width = min(max(20, max_col2_len + 2), 120)
    for row in range(2, max(2, row_cursor)):
        for col in (1, 2):
            cell = course_metadata_sheet.cell(row=row, column=col)
            if cell.value is None or str(cell.value).strip() == "":
                continue
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ensure_workbook_secret_policy()
    workbook.security.lockStructure = True
    course_metadata_sheet.protection.sheet = True
    course_metadata_sheet.protection.password = get_workbook_password()
    course_metadata_sheet.protection.sort = ALLOW_SORT
    course_metadata_sheet.protection.autoFilter = ALLOW_FILTER
    course_metadata_sheet.protection.selectLockedCells = ALLOW_SELECT_LOCKED
    course_metadata_sheet.protection.selectUnlockedCells = ALLOW_SELECT_UNLOCKED
    workbook.save(output_path)
    return output_path


def _extract_students_from_report_sheets(workbook: object) -> set[str]:
    unique_students: set[str] = set()
    try:
        sheets = getattr(workbook, "worksheets", [])
    except Exception:
        return unique_students
    for sheet in sheets:
        if sheet is None:
            continue
        title = str(getattr(sheet, "title", "") or "")
        if title in {SYSTEM_HASH_SHEET, SYSTEM_LAYOUT_SHEET, COURSE_METADATA_SHEET}:
            continue
        max_row = int(getattr(sheet, "max_row", 0) or 0)
        max_col = int(getattr(sheet, "max_column", 0) or 0)
        if max_row <= 0 or max_col <= 0:
            continue
        scan_rows = min(max_row, 30)
        scan_cols = min(max_col, 80)
        reg_col: int | None = None
        header_row = 0
        for row in range(1, scan_rows + 1):
            for col in range(1, scan_cols + 1):
                key = normalize(sheet.cell(row=row, column=col).value)
                if key in {normalize(CO_REPORT_HEADER_REG_NO), normalize(STUDENTS_HEADERS[0])}:
                    reg_col = col
                    header_row = row
                    break
            if reg_col is not None:
                break
        if reg_col is None:
            continue
        for row in range(header_row + 1, max_row + 1):
            reg_raw = sheet.cell(row=row, column=reg_col).value
            coerced = coerce_excel_number(reg_raw)
            reg_text = str(coerced).strip() if coerced is not None else ""
            if reg_text:
                unique_students.add(normalize(reg_text))
    return unique_students
