"""Processing helpers for the coordinator module."""

from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from common.constants import (
    CO_REPORT_ABSENT_TOKEN,
    CO_REPORT_DIRECT_SHEET_SUFFIX,
    CO_REPORT_HEADER_REG_NO,
    CO_REPORT_HEADER_SERIAL,
    CO_REPORT_HEADER_STUDENT_NAME,
    CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE,
    CO_REPORT_INDIRECT_SHEET_SUFFIX,
    CO_REPORT_MAX_DECIMAL_PLACES,
    CO_REPORT_NOT_APPLICABLE_TOKEN,
    COURSE_METADATA_ACADEMIC_YEAR_KEY,
    COURSE_METADATA_COURSE_CODE_KEY,
    COURSE_METADATA_SECTION_KEY,
    COURSE_METADATA_SEMESTER_KEY,
    COURSE_METADATA_SHEET,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    DIRECT_RATIO,
    ID_COURSE_SETUP,
    INDIRECT_RATIO,
    SYSTEM_HASH_SHEET,
    SYSTEM_HASH_TEMPLATE_HASH_HEADER,
    SYSTEM_HASH_TEMPLATE_ID_HEADER,
    SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
    SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
    SYSTEM_REPORT_INTEGRITY_SHEET,
)
from common.excel_sheet_layout import (
    apply_sheet_layout_and_protection as _apply_sheet_layout_and_protection,
    color_without_hash as _color_without_hash,
    style_registry_for_setup as _style_registry_for_setup,
    thin_border as _thin_border,
)
from common.jobs import CancellationToken
from common.utils import coerce_excel_number, normalize
from common.workbook_signing import verify_payload_signature

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}

_logger = logging.getLogger(__name__)
_COURSE_METADATA_COURSE_NAME_KEY = "course_name"
_STYLE_CACHE_ATTR = "_focus_coordinator_style_cache"
_STYLE_KEY_BORDER = "border"
_STYLE_KEY_BG_COLOR = "bg_color"
_STYLE_KEY_ALIGN = "align"
_STYLE_KEY_VALIGN = "valign"
_STYLE_KEY_BOLD = "bold"
_ALIGN_CENTER = "center"
_ALIGN_VCENTER = "vcenter"
_PATTERN_SOLID = "solid"
_CO_REPORT_NAME_TOKEN_RE = re.compile(r"(?:[_\-\s]*co[_\-\s]*report)+$", re.IGNORECASE)
_HEADER_SCAN_MAX_ROWS = 200
_NORM_SYSTEM_HASH_TEMPLATE_ID_HEADER = normalize(SYSTEM_HASH_TEMPLATE_ID_HEADER)
_NORM_SYSTEM_HASH_TEMPLATE_HASH_HEADER = normalize(SYSTEM_HASH_TEMPLATE_HASH_HEADER)
_NORM_SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER = normalize(SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER)
_NORM_SYSTEM_REPORT_INTEGRITY_HASH_HEADER = normalize(SYSTEM_REPORT_INTEGRITY_HASH_HEADER)
_NORM_COURSE_SETUP_ID = normalize(ID_COURSE_SETUP)
_NORM_COURSE_CODE_KEY = normalize(COURSE_METADATA_COURSE_CODE_KEY)
_NORM_TOTAL_OUTCOMES_KEY = normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY)
_NORM_SECTION_KEY = normalize(COURSE_METADATA_SECTION_KEY)


@dataclass(slots=True, frozen=True)
class _FinalReportSignature:
    template_id: str
    course_code: str
    total_outcomes: int
    section: str
    direct_sheet_count: int
    indirect_sheet_count: int


@dataclass(slots=True, frozen=True)
class _CoAttainmentRow:
    reg_no: str
    student_name: str
    direct_score: float | str
    indirect_score: float | str


@dataclass(slots=True)
class _CoOutputSheetState:
    sheet: Any
    header_row_index: int
    formats: dict[str, Any]
    seen_registers: set[str]
    next_row_index: int
    next_serial: int


def _path_key(path: Path) -> str:
    return str(path.resolve()).casefold()


def _build_co_attainment_default_name(source_path: Path, *, section: str = "") -> str:
    stem = source_path.stem.strip()
    cleaned = _CO_REPORT_NAME_TOKEN_RE.sub("", stem).rstrip("_- ").strip()
    section_token = section.strip()
    if section_token:
        parts = [part for part in cleaned.split("_") if normalize(part) != normalize(section_token)]
        cleaned = "_".join(parts).strip("_- ")
    base = cleaned if cleaned else stem
    return f"{base}_CO_Attainment.xlsx"


def _is_supported_excel_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES


def _filter_excel_paths(paths: Iterable[str]) -> list[Path]:
    collected: list[Path] = []
    seen: set[str] = set()
    for value in paths:
        path = Path(value)
        if not _is_supported_excel_file(path):
            continue
        key = _path_key(path)
        if key in seen:
            continue
        seen.add(key)
        collected.append(path.resolve())
    return collected


def _has_valid_final_co_report(path: Path) -> bool:
    return _extract_final_report_signature(path) is not None


def _read_template_id_from_hash_sheet(workbook: Any) -> str | None:
    if SYSTEM_HASH_SHEET not in workbook.sheetnames:
        return None
    hash_sheet = workbook[SYSTEM_HASH_SHEET]
    if normalize(hash_sheet["A1"].value) != _NORM_SYSTEM_HASH_TEMPLATE_ID_HEADER:
        return None
    if normalize(hash_sheet["B1"].value) != _NORM_SYSTEM_HASH_TEMPLATE_HASH_HEADER:
        return None
    template_id_raw = hash_sheet["A2"].value
    template_hash_raw = hash_sheet["B2"].value
    template_id = str(template_id_raw).strip() if template_id_raw is not None else ""
    template_hash = str(template_hash_raw).strip() if template_hash_raw is not None else ""
    if not template_id or not template_hash:
        return None
    if not verify_payload_signature(template_id, template_hash):
        return None
    if normalize(template_id) != _NORM_COURSE_SETUP_ID:
        return None
    return template_id


def _read_report_sheet_counts(workbook: Any) -> tuple[int, int] | None:
    if SYSTEM_REPORT_INTEGRITY_SHEET not in workbook.sheetnames:
        return None
    integrity_sheet = workbook[SYSTEM_REPORT_INTEGRITY_SHEET]
    if normalize(integrity_sheet["A1"].value) != _NORM_SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER:
        return None
    if normalize(integrity_sheet["B1"].value) != _NORM_SYSTEM_REPORT_INTEGRITY_HASH_HEADER:
        return None
    manifest_text_raw = integrity_sheet["A2"].value
    manifest_hash_raw = integrity_sheet["B2"].value
    manifest_text = str(manifest_text_raw).strip() if manifest_text_raw is not None else ""
    manifest_hash = str(manifest_hash_raw).strip() if manifest_hash_raw is not None else ""
    if not manifest_text or not manifest_hash:
        return None
    if not verify_payload_signature(manifest_text, manifest_hash):
        return None

    manifest = json.loads(manifest_text)
    if not isinstance(manifest, dict):
        return None
    sheet_order = manifest.get("sheet_order")
    if not isinstance(sheet_order, list):
        return None
    sheet_names = [str(name).strip() for name in sheet_order if isinstance(name, str)]
    if not sheet_names:
        return None
    direct_sheet_count = sum(1 for name in sheet_names if name.endswith(CO_REPORT_DIRECT_SHEET_SUFFIX))
    indirect_sheet_count = sum(1 for name in sheet_names if name.endswith(CO_REPORT_INDIRECT_SHEET_SUFFIX))
    if direct_sheet_count <= 0 or indirect_sheet_count <= 0:
        return None
    if direct_sheet_count != indirect_sheet_count:
        return None
    return direct_sheet_count, indirect_sheet_count


def _read_signature_metadata(workbook: Any) -> tuple[str, int, str] | None:
    if COURSE_METADATA_SHEET not in workbook.sheetnames:
        return None
    metadata_sheet = workbook[COURSE_METADATA_SHEET]
    course_code = ""
    total_outcomes_text = ""
    section = ""
    row = 2
    while True:
        key = metadata_sheet.cell(row=row, column=1).value
        value = metadata_sheet.cell(row=row, column=2).value
        if normalize(key) == "" and normalize(value) == "":
            break
        key_text = str(key).strip() if key is not None else ""
        value_text = str(value).strip() if value is not None else ""
        if key_text:
            normalized_key = normalize(key_text)
            if normalized_key == _NORM_COURSE_CODE_KEY:
                course_code = value_text
            elif normalized_key == _NORM_TOTAL_OUTCOMES_KEY:
                total_outcomes_text = value_text
            elif normalized_key == _NORM_SECTION_KEY:
                section = value_text
            if course_code and total_outcomes_text and section:
                break
        row += 1

    course_code = course_code.strip()
    total_outcomes_text = total_outcomes_text.strip()
    section = section.strip()
    if not course_code or not total_outcomes_text or not section:
        return None
    try:
        total_outcomes = int(float(total_outcomes_text))
    except (TypeError, ValueError):
        return None
    if total_outcomes <= 0:
        return None
    return course_code, total_outcomes, section


def _extract_final_report_signature(path: Path) -> _FinalReportSignature | None:
    if path.suffix.lower() == ".xls":
        return None
    try:
        from openpyxl import load_workbook
    except Exception:
        _logger.exception("openpyxl is unavailable while validating '%s'.", path)
        return None
    try:
        workbook = load_workbook(filename=path, read_only=True, data_only=True)
        try:
            template_id = _read_template_id_from_hash_sheet(workbook)
            if template_id is None:
                return None
            sheet_counts = _read_report_sheet_counts(workbook)
            if sheet_counts is None:
                return None
            metadata = _read_signature_metadata(workbook)
            if metadata is None:
                return None
            direct_sheet_count, indirect_sheet_count = sheet_counts
            course_code, total_outcomes, section = metadata

            return _FinalReportSignature(
                template_id=template_id,
                course_code=course_code,
                total_outcomes=total_outcomes,
                section=section,
                direct_sheet_count=direct_sheet_count,
                indirect_sheet_count=indirect_sheet_count,
            )
        finally:
            workbook.close()
    except Exception:
        _logger.exception("Failed to validate final CO report workbook '%s'.", path)
        return None


def _analyze_dropped_files(
    dropped_files: list[str],
    *,
    existing_keys: set[str],
    existing_paths: list[str],
    token: CancellationToken,
) -> dict[str, object]:
    accepted = _filter_excel_paths(dropped_files)
    seen = set(existing_keys)
    added: list[str] = []
    duplicates = 0
    invalid_final_report: list[str] = []
    existing_resolved = [Path(path).resolve() for path in existing_paths if path]
    baseline_signature: _FinalReportSignature | None = None
    seen_sections: set[str] = set()
    signature_cache: dict[str, _FinalReportSignature | None] = {}

    def _cached_signature(path: Path) -> _FinalReportSignature | None:
        key = _path_key(path)
        if key not in signature_cache:
            signature_cache[key] = _extract_final_report_signature(path)
        return signature_cache[key]

    for path in existing_resolved:
        token.raise_if_cancelled()
        signature = _cached_signature(path)
        if signature is None:
            continue
        if baseline_signature is None:
            baseline_signature = signature
        section_key = normalize(signature.section)
        if section_key:
            seen_sections.add(section_key)

    for path in accepted:
        token.raise_if_cancelled()
        key = _path_key(path)
        if key in seen:
            duplicates += 1
            continue
        signature = _cached_signature(path)
        if signature is None:
            invalid_final_report.append(str(path))
            continue
        if baseline_signature is None:
            baseline_signature = signature
        else:
            is_mismatch = (
                signature.template_id != baseline_signature.template_id
                or signature.course_code != baseline_signature.course_code
                or signature.total_outcomes != baseline_signature.total_outcomes
                or signature.direct_sheet_count != baseline_signature.direct_sheet_count
                or signature.indirect_sheet_count != baseline_signature.indirect_sheet_count
            )
            if is_mismatch or normalize(signature.section) in seen_sections:
                invalid_final_report.append(str(path))
                continue
        seen_sections.add(normalize(signature.section))
        seen.add(key)
        added.append(str(path))

    ignored = (len(dropped_files) - len(accepted)) + duplicates + len(invalid_final_report)
    return {
        "added": added,
        "duplicates": duplicates,
        "invalid_final_report": invalid_final_report,
        "ignored": ignored,
    }


def _ratio_percent_token(ratio: float) -> str:
    percent = ratio * 100.0
    if abs(percent - round(percent)) <= 1e-9:
        return f"{int(round(percent))}"
    return f"{percent:g}"


def _ratio_total_header(ratio: float) -> str:
    return CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE.format(ratio=_ratio_percent_token(ratio))


def _coerce_numeric_score(value: Any) -> float | str | None:
    if normalize(value) == normalize(CO_REPORT_NOT_APPLICABLE_TOKEN):
        return CO_REPORT_ABSENT_TOKEN
    parsed = coerce_excel_number(value)
    if isinstance(parsed, bool):
        return None
    if isinstance(parsed, (int, float)):
        return float(parsed)
    return None


def _metadata_rows_for_output(metadata: dict[str, str], co_index: int) -> list[tuple[str, str]]:
    return [
        ("Course Code", metadata.get(normalize(COURSE_METADATA_COURSE_CODE_KEY), "")),
        ("Course Name", metadata.get(normalize(_COURSE_METADATA_COURSE_NAME_KEY), "")),
        ("Semester", metadata.get(normalize(COURSE_METADATA_SEMESTER_KEY), "")),
        ("Academic Year", metadata.get(normalize(COURSE_METADATA_ACADEMIC_YEAR_KEY), "")),
        ("CO Number", f"CO{co_index}"),
    ]


def _extract_course_metadata_fields(sheet: Any) -> dict[str, str]:
    metadata: dict[str, str] = {}
    row = 2
    while True:
        key = sheet.cell(row=row, column=1).value
        value = sheet.cell(row=row, column=2).value
        if normalize(key) == "" and normalize(value) == "":
            break
        key_text = str(key).strip() if key is not None else ""
        if key_text:
            coerced = coerce_excel_number(value)
            metadata[normalize(key_text)] = str(coerced).strip() if coerced is not None else ""
        row += 1
    return metadata


def _extract_co_scores_by_reg(sheet: Any, *, ratio: float) -> dict[str, tuple[str, str, float | str]]:
    required_headers = {
        normalize(CO_REPORT_HEADER_SERIAL),
        normalize(CO_REPORT_HEADER_REG_NO),
        normalize(CO_REPORT_HEADER_STUDENT_NAME),
        normalize(_ratio_total_header(ratio)),
    }
    header_row = 0
    column_map: dict[str, int] = {}
    max_row = int(sheet.max_row)
    max_col = int(sheet.max_column)
    header_scan_limit = min(max_row, _HEADER_SCAN_MAX_ROWS)

    def _find_header_row(start_row: int, end_row: int) -> tuple[int, dict[str, int]]:
        for row_offset, values in enumerate(
            sheet.iter_rows(
                min_row=start_row,
                max_row=end_row,
                min_col=1,
                max_col=max_col,
                values_only=True,
            ),
            start=0,
        ):
            row_idx = start_row + row_offset
            row_map: dict[str, int] = {}
            missing = set(required_headers)
            for col_idx, value in enumerate(values, start=1):
                key = normalize(value)
                if key and key not in row_map:
                    row_map[key] = col_idx
                    if key in missing:
                        missing.remove(key)
                        if not missing:
                            return row_idx, row_map
            if not missing:
                return row_idx, row_map
        return 0, {}

    header_row, column_map = _find_header_row(1, header_scan_limit)
    if header_row <= 0 and header_scan_limit < max_row:
        header_row, column_map = _find_header_row(header_scan_limit + 1, max_row)
    if header_row <= 0:
        raise ValueError(f"Required headers are missing in sheet '{sheet.title}'.")

    reg_col = column_map[normalize(CO_REPORT_HEADER_REG_NO)]
    name_col = column_map[normalize(CO_REPORT_HEADER_STUDENT_NAME)]
    score_col = column_map[normalize(_ratio_total_header(ratio))]

    rows: dict[str, tuple[str, str, float | str]] = {}
    for values in sheet.iter_rows(
        min_row=header_row + 1,
        max_row=max_row,
        min_col=1,
        max_col=max_col,
        values_only=True,
    ):
        if len(values) < max(reg_col, name_col, score_col):
            continue
        reg_raw = values[reg_col - 1]
        reg_value = coerce_excel_number(reg_raw)
        reg_no = str(reg_value).strip() if reg_value is not None else ""
        if not reg_no:
            continue
        reg_key = normalize(reg_no)
        if reg_key in rows:
            continue

        score = _coerce_numeric_score(values[score_col - 1])
        if score is None:
            continue
        student_raw = values[name_col - 1]
        student_name = str(student_raw).strip() if student_raw is not None else ""
        normalized_score = round(score, CO_REPORT_MAX_DECIMAL_PLACES) if isinstance(score, (int, float)) else score
        rows[reg_key] = (reg_no, student_name, normalized_score)
    return rows


def _iter_co_rows_from_workbook(workbook: Any, *, co_index: int) -> Iterable[_CoAttainmentRow]:
    direct_name = f"CO{co_index}{CO_REPORT_DIRECT_SHEET_SUFFIX}"
    indirect_name = f"CO{co_index}{CO_REPORT_INDIRECT_SHEET_SUFFIX}"
    if direct_name not in workbook.sheetnames or indirect_name not in workbook.sheetnames:
        raise ValueError(f"Missing CO sheets for CO{co_index}.")

    direct_rows = _extract_co_scores_by_reg(workbook[direct_name], ratio=DIRECT_RATIO)
    indirect_rows = _extract_co_scores_by_reg(workbook[indirect_name], ratio=INDIRECT_RATIO)
    for reg_key, (reg_no, direct_name_value, direct_score) in direct_rows.items():
        indirect_entry = indirect_rows.get(reg_key)
        if indirect_entry is None:
            continue
        _, indirect_name_value, indirect_score = indirect_entry
        student_name = direct_name_value or indirect_name_value
        direct_is_absent = isinstance(direct_score, str) and normalize(direct_score) == normalize(CO_REPORT_ABSENT_TOKEN)
        indirect_is_absent = isinstance(indirect_score, str) and normalize(indirect_score) == normalize(CO_REPORT_ABSENT_TOKEN)
        if direct_is_absent or indirect_is_absent:
            direct_score = CO_REPORT_ABSENT_TOKEN
            indirect_score = CO_REPORT_ABSENT_TOKEN
        yield _CoAttainmentRow(
            reg_no=reg_no,
            student_name=student_name,
            direct_score=direct_score,
            indirect_score=indirect_score,
        )


def _xlsxwriter_formats(workbook: Any) -> dict[str, Any]:
    cached = getattr(workbook, _STYLE_CACHE_ATTR, None)
    if isinstance(cached, dict):
        return cached
    header_style, body_style = _style_registry_for_setup()
    header_bg = _color_without_hash(str(header_style.get(_STYLE_KEY_BG_COLOR, ""))) or "D9EAD3"
    border_enabled = int(body_style.get(_STYLE_KEY_BORDER, 1)) > 0
    header_border_enabled = int(header_style.get(_STYLE_KEY_BORDER, 1)) > 0
    border_value = 1 if border_enabled else 0
    header_border_value = 1 if header_border_enabled else 0
    formats = {
        "header": workbook.add_format(
            {
                "bold": bool(header_style.get(_STYLE_KEY_BOLD, True)),
                "border": header_border_value,
                "align": str(header_style.get(_STYLE_KEY_ALIGN, _ALIGN_CENTER)),
                "valign": str(header_style.get(_STYLE_KEY_VALIGN, _ALIGN_CENTER)).replace(_ALIGN_VCENTER, _ALIGN_CENTER),
                "text_wrap": True,
                "fg_color": header_bg,
                "pattern": 1,
            }
        ),
        "body": workbook.add_format(
            {
                "border": border_value,
                "valign": _ALIGN_CENTER,
            }
        ),
        "body_center": workbook.add_format(
            {
                "border": border_value,
                "align": _ALIGN_CENTER,
                "valign": _ALIGN_CENTER,
            }
        ),
    }
    setattr(workbook, _STYLE_CACHE_ATTR, formats)
    return formats


def _create_co_attainment_sheet(
    workbook: Any,
    *,
    co_index: int,
    metadata: dict[str, str],
    ) -> _CoOutputSheetState:
    sheet = workbook.add_worksheet(f"CO{co_index}")
    formats = _xlsxwriter_formats(workbook)
    metadata_rows = _metadata_rows_for_output(metadata, co_index)
    for row_idx, (label, value) in enumerate(metadata_rows, start=0):
        sheet.write(row_idx, 1, label, formats["body"])
        sheet.write(row_idx, 2, value, formats["body"])

    header_row_index = len(metadata_rows) + 1
    headers = [
        "#",
        "Regno",
        "Student name",
        f"Direct ({_ratio_percent_token(DIRECT_RATIO)}%)",
        f"Indirect ({_ratio_percent_token(INDIRECT_RATIO)}%)",
        "Total (100%)",
    ]
    for col_idx, header in enumerate(headers, start=0):
        sheet.write(header_row_index, col_idx, header, formats["header"])

    sheet.set_landscape()
    sheet.set_paper(9)  # A4
    sheet.fit_to_pages(1, 0)
    sheet.protect()
    sheet.set_selection(header_row_index, 0, header_row_index, 0)
    return _CoOutputSheetState(
        sheet=sheet,
        header_row_index=header_row_index,
        formats=formats,
        seen_registers=set(),
        next_row_index=header_row_index + 1,
        next_serial=1,
    )


def _append_co_attainment_row(state: _CoOutputSheetState, row: _CoAttainmentRow) -> None:
    if isinstance(row.direct_score, (int, float)) and isinstance(row.indirect_score, (int, float)):
        total: float | str = round(row.direct_score + row.indirect_score, CO_REPORT_MAX_DECIMAL_PLACES)
    else:
        total = CO_REPORT_ABSENT_TOKEN
    values = [state.next_serial, row.reg_no, row.student_name, row.direct_score, row.indirect_score, total]
    for col_idx, value in enumerate(values, start=0):
        state.sheet.write(
            state.next_row_index,
            col_idx,
            value,
            state.formats["body_center"] if col_idx >= 3 else state.formats["body"],
        )
    state.next_row_index += 1
    state.next_serial += 1


def _style_cache_for_sheet(ws: Any) -> dict[str, Any]:
    from openpyxl.styles import Alignment, Border, Font, PatternFill

    cached = getattr(ws, _STYLE_CACHE_ATTR, None)
    if isinstance(cached, dict):
        return cached
    header_style, body_style = _style_registry_for_setup()
    border_enabled = int(body_style.get(_STYLE_KEY_BORDER, 1)) > 0
    body_border = _thin_border() if border_enabled else Border()
    header_border_enabled = int(header_style.get(_STYLE_KEY_BORDER, 1)) > 0
    header_border = _thin_border() if header_border_enabled else Border()
    header_bg = _color_without_hash(str(header_style.get(_STYLE_KEY_BG_COLOR, "")))
    header_fill = PatternFill(fill_type=_PATTERN_SOLID, fgColor=header_bg) if header_bg else PatternFill(fill_type=None)
    header_alignment = Alignment(
        horizontal=str(header_style.get(_STYLE_KEY_ALIGN, _ALIGN_CENTER)),
        vertical=str(header_style.get(_STYLE_KEY_VALIGN, _ALIGN_CENTER)).replace(_ALIGN_VCENTER, _ALIGN_CENTER),
        wrap_text=True,
    )
    body_align = str(body_style.get(_STYLE_KEY_ALIGN, "")).strip()
    body_valign = str(body_style.get(_STYLE_KEY_VALIGN, _ALIGN_CENTER)).replace(_ALIGN_VCENTER, _ALIGN_CENTER)
    body_alignment = Alignment(horizontal=body_align if body_align else None, vertical=body_valign)
    body_alignment_center = Alignment(horizontal=_ALIGN_CENTER, vertical=body_valign)
    style_cache = {
        "header_border": header_border,
        "body_border": body_border,
        "header_fill": header_fill,
        "header_font": Font(bold=bool(header_style.get(_STYLE_KEY_BOLD, True))),
        "header_alignment": header_alignment,
        "body_alignment": body_alignment,
        "body_alignment_center": body_alignment_center,
    }
    setattr(ws, _STYLE_CACHE_ATTR, style_cache)
    return style_cache


def _generate_co_attainment_workbook(
    source_paths: list[Path],
    output_path: Path,
    *,
    token: CancellationToken,
) -> Path:
    if not source_paths:
        raise ValueError("No source files provided for CO attainment calculation.")

    first_signature = _extract_final_report_signature(source_paths[0])
    if first_signature is None:
        raise ValueError(f"Invalid final CO report file: {source_paths[0]}")
    if normalize(first_signature.template_id) != normalize(ID_COURSE_SETUP):
        raise ValueError(f"Unsupported template id '{first_signature.template_id}'. Expected '{ID_COURSE_SETUP}'.")
    return _generate_co_attainment_workbook_course_setup_v1(
        source_paths,
        output_path,
        token=token,
        total_outcomes=first_signature.total_outcomes,
    )


def _generate_co_attainment_workbook_course_setup_v1(
    source_paths: list[Path],
    output_path: Path,
    *,
    token: CancellationToken,
    total_outcomes: int,
) -> Path:
    try:
        import xlsxwriter
        from openpyxl import load_workbook
    except Exception as exc:  # pragma: no cover - guarded by runtime dependency availability
        raise RuntimeError("openpyxl and xlsxwriter are required for CO attainment calculation.") from exc

    if total_outcomes <= 0:
        raise ValueError("No CO outcomes available in the uploaded reports.")

    metadata: dict[str, str] = {}
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_workbook = xlsxwriter.Workbook(str(output_path), {"constant_memory": True})
    workbook_closed = False
    output_states: dict[int, _CoOutputSheetState] = {}

    try:
        for source in source_paths:
            token.raise_if_cancelled()
            workbook = load_workbook(filename=source, data_only=True, read_only=True)
            try:
                if not metadata and COURSE_METADATA_SHEET in workbook.sheetnames:
                    metadata = _extract_course_metadata_fields(workbook[COURSE_METADATA_SHEET])
                    for co_index in range(1, total_outcomes + 1):
                        output_states[co_index] = _create_co_attainment_sheet(
                            output_workbook,
                            co_index=co_index,
                            metadata=metadata,
                        )
                if not output_states:
                    for co_index in range(1, total_outcomes + 1):
                        output_states[co_index] = _create_co_attainment_sheet(
                            output_workbook,
                            co_index=co_index,
                            metadata=metadata,
                        )
                for co_index in range(1, total_outcomes + 1):
                    state = output_states[co_index]
                    for row in _iter_co_rows_from_workbook(workbook, co_index=co_index):
                        reg_key = normalize(row.reg_no)
                        if not reg_key or reg_key in state.seen_registers:
                            continue
                        state.seen_registers.add(reg_key)
                        _append_co_attainment_row(state, row)
            finally:
                workbook.close()
        output_workbook.close()
        workbook_closed = True
    finally:
        if not workbook_closed:
            try:
                output_workbook.close()
            except Exception:
                pass
    return output_path
