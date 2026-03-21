"""Final CO report generation from validated filled-marks workbooks."""

from __future__ import annotations

import json
import logging
import os
from dataclasses import dataclass
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

from common.constants import (
    ASSESSMENT_CONFIG_HEADERS,
    ASSESSMENT_CONFIG_SHEET,
    ASSESSMENT_VALIDATION_YES_NO_OPTIONS,
    CO_REPORT_ABSENT_TOKEN,
    CO_REPORT_DIRECT_SHEET_SUFFIX,
    CO_REPORT_HEADER_REG_NO,
    CO_REPORT_HEADER_SERIAL,
    CO_REPORT_HEADER_STUDENT_NAME,
    CO_REPORT_HEADER_TOTAL,
    CO_REPORT_HEADER_TOTAL_100,
    CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE,
    CO_REPORT_INDIRECT_SHEET_SUFFIX,
    CO_REPORT_MAX_DECIMAL_PLACES,
    CO_REPORT_METADATA_OUTCOME_FIELD,
    CO_REPORT_METADATA_OUTCOME_VALUE_INDIRECT_TEMPLATE,
    CO_REPORT_METADATA_OUTCOME_VALUE_TEMPLATE,
    CO_REPORT_NOT_APPLICABLE_TOKEN,
    CO_REPORT_PERCENT_SYMBOL,
    CO_REPORT_SCALED_LABEL_TEMPLATE,
    COMPONENT_NAME_LABEL,
    COURSE_METADATA_FACULTY_NAME_KEY,
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SHEET,
    COURSE_METADATA_TOTAL_OUTCOMES_KEY,
    DIRECT_RATIO,
    ID_COURSE_SETUP,
    INDIRECT_RATIO,
    LAYOUT_MANIFEST_KEY_SHEETS,
    LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
    LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
    LAYOUT_SHEET_KIND_INDIRECT,
    LAYOUT_SHEET_SPEC_KEY_ANCHORS,
    LAYOUT_SHEET_SPEC_KEY_HEADER_ROW,
    LAYOUT_SHEET_SPEC_KEY_HEADERS,
    LAYOUT_SHEET_SPEC_KEY_KIND,
    LAYOUT_SHEET_SPEC_KEY_NAME,
    LIKERT_MAX,
    LIKERT_MIN,
    SYSTEM_HASH_SHEET,
    SYSTEM_HASH_TEMPLATE_HASH_HEADER,
    SYSTEM_HASH_TEMPLATE_ID_HEADER,
    SYSTEM_LAYOUT_MANIFEST_HASH_HEADER,
    SYSTEM_LAYOUT_MANIFEST_HASH_KEY,
    SYSTEM_LAYOUT_MANIFEST_HEADER,
    SYSTEM_LAYOUT_MANIFEST_KEY,
    SYSTEM_LAYOUT_SHEET,
    WORKBOOK_TEMP_SUFFIX,
)
from common.excel_sheet_layout import color_without_hash as _color_without_hash
from common.excel_sheet_layout import (
    compute_sampled_column_widths as _compute_sampled_column_widths,
)
from common.excel_sheet_layout import (
    style_registry_for_setup as _style_registry_for_setup,
)
from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.jobs import CancellationToken
from common.texts import t
from common.utils import coerce_excel_number, normalize
from common.workbook_signing import sign_payload, verify_payload_signature

_logger = logging.getLogger(__name__)
_ANCHOR_COL_B = "B"
_ANCHOR_COL_C = "C"
_CELL_LAYOUT_MANIFEST = "A2"
_CELL_LAYOUT_MANIFEST_HASH = "B2"
_CELL_SYSTEM_TEMPLATE_ID = "A2"
_CELL_SYSTEM_TEMPLATE_HASH = "B2"
_CO_TOKEN_PREFIX = "co"
_CO_TOKEN_SEPARATOR = ","
_EXCEL_COL_REG_NO = 2
_EXCEL_COL_STUDENT_NAME = 3
_EXCEL_COL_FIRST_MARK = 4
_EXCEL_ROW_HEADER_OFFSET_DIRECT = 3
_EXCEL_ROW_HEADER_OFFSET_INDIRECT = 1
_EXCEL_ROW_CO_OFFSET = 1
_EXCEL_ROW_MAX_OFFSET = 2
_STYLE_KEY_BORDER = "border"
_STYLE_KEY_BG_COLOR = "bg_color"
_STYLE_KEY_ALIGN = "align"
_STYLE_KEY_VALIGN = "valign"
_STYLE_KEY_BOLD = "bold"
_ALIGN_CENTER = "center"
_ALIGN_VCENTER = "vcenter"
_PATTERN_SOLID = "solid"
_STYLE_CACHE_ATTR = "_focus_report_style_cache"
_INTEGRITY_SCHEMA_VERSION = 1
_LOG_FINAL_REPORT_READY = "Final CO report prepared with direct and indirect sheets."
_LOG_FINAL_REPORT_FAILED = "Final CO report generation failed unexpectedly."

@dataclass(frozen=True)
class _DirectComponent:
    name: str
    weight: float

@dataclass
class _DirectComponentComputed:
    name: str
    weight: float
    max_by_co: dict[int, float]
    marks_by_co: dict[int, list[float | str]]

@dataclass(frozen=True)
class _IndirectComponent:
    name: str
    weight: float

@dataclass
class _IndirectComponentComputed:
    name: str
    weight: float
    marks_by_co: dict[int, list[float | str]]

def generate_final_co_report(
    filled_marks_path: str | Path,
    output_path: str | Path,
    *,
    cancel_token: CancellationToken | None = None,
) -> Path:
    try:
        import openpyxl
    except ModuleNotFoundError as exc:
        raise ValidationError(t("instructor.validation.openpyxl_missing")) from exc
    try:
        import xlsxwriter
    except ModuleNotFoundError as exc:
        raise ValidationError(t("instructor.validation.openpyxl_missing")) from exc

    source = Path(filled_marks_path)
    output = Path(output_path)
    if not source.exists():
        raise ValidationError(t("instructor.validation.workbook_not_found", workbook=source))
    output.parent.mkdir(parents=True, exist_ok=True)

    try:
        _raise_if_cancelled(cancel_token)
        src_wb = openpyxl.load_workbook(source, data_only=False)
        _validate_source_workbook_integrity(src_wb)
    except JobCancelledError:
        raise
    except ValidationError:
        raise
    except Exception as exc:
        raise ValidationError(t("instructor.validation.workbook_open_failed", workbook=source)) from exc

    tmp_path: Path | None = None
    try:
        _raise_if_cancelled(cancel_token)
        metadata_rows, total_outcomes = _read_course_metadata(src_wb[COURSE_METADATA_SHEET])
        template_id_raw = src_wb[SYSTEM_HASH_SHEET][_CELL_SYSTEM_TEMPLATE_ID].value
        template_hash_raw = src_wb[SYSTEM_HASH_SHEET][_CELL_SYSTEM_TEMPLATE_HASH].value
        template_id = str(template_id_raw).strip() if template_id_raw is not None else ""
        template_hash = str(template_hash_raw).strip() if template_hash_raw is not None else ""
        direct_components = _read_direct_components(src_wb[ASSESSMENT_CONFIG_SHEET])
        indirect_components = _read_indirect_components(src_wb[ASSESSMENT_CONFIG_SHEET])
        if not direct_components and not indirect_components:
            raise ValidationError(t("instructor.validation.final_report.no_direct_components"))
        manifest = _read_layout_manifest(src_wb[SYSTEM_LAYOUT_SHEET])
        direct_sheet_specs = _component_sheet_specs_by_kind(
            manifest,
            allowed_kinds={
                LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
                LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
            },
        )
        indirect_sheet_specs = _component_sheet_specs_by_kind(
            manifest,
            allowed_kinds={LAYOUT_SHEET_KIND_INDIRECT},
        )
        students = _read_students_from_component_sheets(
            src_wb,
            direct_specs=direct_sheet_specs,
            indirect_specs=indirect_sheet_specs,
        )

        computed_components: list[_DirectComponentComputed] = []
        for component in direct_components:
            _raise_if_cancelled(cancel_token)
            computed_components.append(
                _compute_component_marks(
                    workbook=src_wb,
                    component=component,
                    students=students,
                    spec=direct_sheet_specs.get(normalize(component.name)),
                    total_outcomes=total_outcomes,
                )
            )

        computed_indirect_components: list[_IndirectComponentComputed] = []
        for component in indirect_components:
            _raise_if_cancelled(cancel_token)
            computed_indirect_components.append(
                _compute_indirect_component_marks(
                    workbook=src_wb,
                    component=component,
                    students=students,
                    spec=indirect_sheet_specs.get(normalize(component.name)),
                    total_outcomes=total_outcomes,
                )
            )

        _raise_if_cancelled(cancel_token)
        with NamedTemporaryFile(
            mode="wb",
            delete=False,
            dir=str(output.parent),
            prefix=f"{output.name}.",
            suffix=WORKBOOK_TEMP_SUFFIX,
        ) as temp_file:
            tmp_path = Path(temp_file.name)
        _write_final_report_workbook_xlsxwriter(
            xlsxwriter_module=xlsxwriter,
            output_path=tmp_path,
            metadata_rows=metadata_rows,
            total_outcomes=total_outcomes,
            students=students,
            direct_components=computed_components,
            indirect_components=computed_indirect_components,
            template_id=template_id,
            template_hash=template_hash,
        )
        _normalize_page_setup_fit(tmp_path)
        _logger.info(
            _LOG_FINAL_REPORT_READY,
            extra={
                "user_message": f"Prepared {total_outcomes} direct and {total_outcomes} indirect CO sheets.",
                "co_count": total_outcomes,
            },
        )
        _raise_if_cancelled(cancel_token)
        os.replace(tmp_path, output)
        tmp_path = None
        return output
    except ValidationError:
        raise
    except JobCancelledError:
        raise
    except Exception as exc:
        _logger.exception(_LOG_FINAL_REPORT_FAILED, exc_info=exc)
        raise AppSystemError(
            t("instructor.log.error_while_process", process=t("instructor.log.process.generate_final_co_report"))
        ) from exc
    finally:
        src_wb.close()
        if tmp_path is not None and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                _logger.warning("Failed to cleanup temp final report file: %s", tmp_path)

def _raise_if_cancelled(cancel_token: CancellationToken | None) -> None:
    if cancel_token is not None:
        cancel_token.raise_if_cancelled()

def _validate_source_workbook_integrity(workbook: Any) -> None:
    if SYSTEM_HASH_SHEET not in workbook.sheetnames:
        raise ValidationError(t("instructor.validation.system_sheet_missing", sheet=SYSTEM_HASH_SHEET))
    if SYSTEM_LAYOUT_SHEET not in workbook.sheetnames:
        raise ValidationError(t("instructor.validation.step2.layout_sheet_missing", sheet=SYSTEM_LAYOUT_SHEET))

    hash_sheet = workbook[SYSTEM_HASH_SHEET]
    template_id_raw = hash_sheet[_CELL_SYSTEM_TEMPLATE_ID].value
    template_hash_raw = hash_sheet[_CELL_SYSTEM_TEMPLATE_HASH].value
    template_id = str(template_id_raw).strip() if template_id_raw is not None else ""
    template_hash = str(template_hash_raw).strip() if template_hash_raw is not None else ""
    if not template_id or not template_hash:
        raise ValidationError(t("instructor.validation.system_hash_template_id_missing"))
    if not verify_payload_signature(template_id, template_hash):
        raise ValidationError(t("instructor.validation.system_hash_mismatch"))
    if normalize(template_id) != normalize(ID_COURSE_SETUP):
        raise ValidationError(
            t(
                "instructor.validation.unknown_template",
                template_id=template_id,
                available=ID_COURSE_SETUP,
            )
        )

    layout_sheet = workbook[SYSTEM_LAYOUT_SHEET]
    if normalize(layout_sheet["A1"].value) != normalize(SYSTEM_LAYOUT_MANIFEST_KEY):
        raise ValidationError(
            t(
                "instructor.validation.step2.layout_header_mismatch",
                column="A1",
                expected=SYSTEM_LAYOUT_MANIFEST_KEY,
            )
        )
    if normalize(layout_sheet["B1"].value) != normalize(SYSTEM_LAYOUT_MANIFEST_HASH_KEY):
        raise ValidationError(
            t(
                "instructor.validation.step2.layout_header_mismatch",
                column="B1",
                expected=SYSTEM_LAYOUT_MANIFEST_HASH_KEY,
            )
        )
    manifest_text_raw = layout_sheet[_CELL_LAYOUT_MANIFEST].value
    manifest_hash_raw = layout_sheet[_CELL_LAYOUT_MANIFEST_HASH].value
    manifest_text = str(manifest_text_raw).strip() if manifest_text_raw is not None else ""
    manifest_hash = str(manifest_hash_raw).strip() if manifest_hash_raw is not None else ""
    if not manifest_text or not manifest_hash:
        raise ValidationError(t("instructor.validation.step2.layout_manifest_missing"))
    if not verify_payload_signature(manifest_text, manifest_hash):
        raise ValidationError(t("instructor.validation.step2.layout_hash_mismatch"))

def _read_course_metadata(sheet: Any) -> tuple[list[tuple[str, Any]], int]:
    rows: list[tuple[str, Any]] = []
    total_outcomes = 0
    row = 2
    while True:
        key = sheet.cell(row=row, column=1).value
        value = sheet.cell(row=row, column=2).value
        if normalize(key) == "" and normalize(value) == "":
            break
        key_text = str(key).strip() if key is not None else ""
        rows.append((key_text, value))
        if normalize(key_text) == normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY):
            parsed = coerce_excel_number(value)
            if isinstance(parsed, (int, float)) and not isinstance(parsed, bool):
                total_outcomes = int(parsed)
        row += 1
    if total_outcomes <= 0:
        raise ValidationError(t("instructor.validation.course_metadata_total_outcomes_invalid"))
    return rows, total_outcomes

def _read_direct_components(sheet: Any) -> list[_DirectComponent]:
    expected_headers = list(ASSESSMENT_CONFIG_HEADERS)
    headers = [normalize(value) for value in expected_headers]
    header_map: dict[str, int] = {}
    for idx, expected in enumerate(headers, start=1):
        actual = normalize(sheet.cell(row=1, column=idx).value)
        if actual != expected:
            raise ValidationError(
                t("instructor.validation.header_mismatch", sheet_name=ASSESSMENT_CONFIG_SHEET, expected=expected_headers)
            )
        header_map[expected] = idx

    out: list[_DirectComponent] = []
    row = 2
    while True:
        component_name = sheet.cell(row=row, column=header_map[normalize(ASSESSMENT_CONFIG_HEADERS[0])]).value
        if normalize(component_name) == "":
            break
        is_direct = normalize(
            sheet.cell(row=row, column=header_map[normalize(ASSESSMENT_CONFIG_HEADERS[4])]).value
        ) == normalize(ASSESSMENT_VALIDATION_YES_NO_OPTIONS[0])
        if is_direct:
            weight = coerce_excel_number(
                sheet.cell(row=row, column=header_map[normalize(ASSESSMENT_CONFIG_HEADERS[1])]).value
            )
            if not isinstance(weight, (int, float)) or isinstance(weight, bool):
                raise ValidationError(
                    t("instructor.validation.assessment_weight_numeric", row=row)
                )
            out.append(
                _DirectComponent(
                    name=str(component_name).strip(),
                    weight=float(weight),
                )
            )
        row += 1
    return out

def _read_indirect_components(sheet: Any) -> list[_IndirectComponent]:
    expected_headers = list(ASSESSMENT_CONFIG_HEADERS)
    headers = [normalize(value) for value in expected_headers]
    header_map: dict[str, int] = {}
    for idx, expected in enumerate(headers, start=1):
        header_map[expected] = idx

    out: list[_IndirectComponent] = []
    row = 2
    while True:
        component_name = sheet.cell(row=row, column=header_map[normalize(ASSESSMENT_CONFIG_HEADERS[0])]).value
        if normalize(component_name) == "":
            break
        is_direct = normalize(
            sheet.cell(row=row, column=header_map[normalize(ASSESSMENT_CONFIG_HEADERS[4])]).value
        ) == normalize(ASSESSMENT_VALIDATION_YES_NO_OPTIONS[0])
        if not is_direct:
            weight = coerce_excel_number(
                sheet.cell(row=row, column=header_map[normalize(ASSESSMENT_CONFIG_HEADERS[1])]).value
            )
            if not isinstance(weight, (int, float)) or isinstance(weight, bool):
                raise ValidationError(
                    t("instructor.validation.assessment_weight_numeric", row=row)
                )
            out.append(
                _IndirectComponent(
                    name=str(component_name).strip(),
                    weight=float(weight),
                )
            )
        row += 1
    return out


def _parse_co_values(raw: Any) -> list[int]:
    token = str(raw).strip() if raw is not None else ""
    if not token:
        return []
    values: list[int] = []
    for part in token.split(_CO_TOKEN_SEPARATOR):
        p = part.strip().lower()
        if p.startswith(_CO_TOKEN_PREFIX):
            p = p[len(_CO_TOKEN_PREFIX) :].strip()
        if p.isdigit():
            values.append(int(p))
    # Preserve order and dedupe.
    seen: set[int] = set()
    out: list[int] = []
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        out.append(value)
    return out


def _read_layout_manifest(sheet: Any) -> dict[str, Any]:
    raw_manifest = sheet[_CELL_LAYOUT_MANIFEST].value
    if not isinstance(raw_manifest, str):
        raise ValidationError(t("instructor.validation.final_report.layout_manifest_invalid"))
    try:
        manifest = json.loads(raw_manifest)
    except Exception as exc:
        raise ValidationError(t("instructor.validation.final_report.layout_manifest_invalid")) from exc
    if not isinstance(manifest, dict):
        raise ValidationError(t("instructor.validation.final_report.layout_manifest_invalid"))
    return manifest


def _component_sheet_specs_by_kind(
    manifest: dict[str, Any],
    *,
    allowed_kinds: set[str],
) -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    for spec in manifest.get(LAYOUT_MANIFEST_KEY_SHEETS, []):
        if not isinstance(spec, dict):
            continue
        kind = spec.get(LAYOUT_SHEET_SPEC_KEY_KIND)
        if kind not in allowed_kinds:
            continue
        component_name = _component_name_from_spec(spec)
        if not component_name:
            continue
        out[normalize(component_name)] = spec
    return out


def _component_name_from_spec(spec: dict[str, Any]) -> str:
    anchors = spec.get(LAYOUT_SHEET_SPEC_KEY_ANCHORS, [])
    if not isinstance(anchors, list):
        return ""
    row_by_label: dict[str, int] = {}
    value_by_row: dict[int, str] = {}
    for item in anchors:
        if not isinstance(item, list) or len(item) != 2:
            continue
        cell_ref, value = item
        if not isinstance(cell_ref, str):
            continue
        row_no = _cell_row(cell_ref)
        if row_no <= 0:
            continue
        cell_col = cell_ref[:1].upper()
        text = str(value).strip() if value is not None else ""
        if cell_col == _ANCHOR_COL_B:
            row_by_label[normalize(text)] = row_no
        elif cell_col == _ANCHOR_COL_C:
            value_by_row[row_no] = text
    component_row = row_by_label.get(normalize(COMPONENT_NAME_LABEL))
    if component_row is None:
        return ""
    return value_by_row.get(component_row, "")


def _cell_row(cell_ref: str) -> int:
    digits = "".join(ch for ch in cell_ref if ch.isdigit())
    return int(digits) if digits else 0


def _read_students_from_component_sheets(
    workbook: Any,
    *,
    direct_specs: dict[str, dict[str, Any]],
    indirect_specs: dict[str, dict[str, Any]],
) -> list[tuple[str, str]]:
    specs_by_component = direct_specs or indirect_specs
    if not specs_by_component:
        raise ValidationError(t("instructor.validation.final_report.no_direct_components"))
    first_spec = next(iter(specs_by_component.values()))
    sheet_name = first_spec.get(LAYOUT_SHEET_SPEC_KEY_NAME)
    header_row = first_spec.get(LAYOUT_SHEET_SPEC_KEY_HEADER_ROW)
    if not isinstance(sheet_name, str) or not isinstance(header_row, int):
        raise ValidationError(t("instructor.validation.final_report.layout_manifest_invalid"))
    ws = workbook[sheet_name]
    first_data_row = header_row + _EXCEL_ROW_HEADER_OFFSET_DIRECT
    students: list[tuple[str, str]] = []
    row = first_data_row
    while True:
        reg = ws.cell(row=row, column=_EXCEL_COL_REG_NO).value
        name = ws.cell(row=row, column=_EXCEL_COL_STUDENT_NAME).value
        if normalize(reg) == "" and normalize(name) == "":
            break
        students.append((str(reg).strip(), str(name).strip()))
        row += 1
    return students


def _compute_component_marks(
    *,
    workbook: Any,
    component: _DirectComponent,
    students: list[tuple[str, str]],
    spec: dict[str, Any] | None,
    total_outcomes: int,
) -> _DirectComponentComputed:
    if spec is None:
        raise ValidationError(
            t("instructor.validation.final_report.direct_component_sheet_missing", component=component.name)
        )
    sheet_name = spec.get(LAYOUT_SHEET_SPEC_KEY_NAME)
    header_row = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADER_ROW)
    headers = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADERS)
    kind = spec.get(LAYOUT_SHEET_SPEC_KEY_KIND)
    if (
        not isinstance(sheet_name, str)
        or not isinstance(header_row, int)
        or not isinstance(headers, list)
        or kind not in {LAYOUT_SHEET_KIND_DIRECT_CO_WISE, LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE}
    ):
        raise ValidationError(t("instructor.validation.final_report.layout_manifest_invalid"))
    ws = workbook[sheet_name]
    first_data_row = header_row + _EXCEL_ROW_HEADER_OFFSET_DIRECT
    expected_students = len(students)

    max_by_co = {co: 0.0 for co in range(1, total_outcomes + 1)}
    marks_by_co: dict[int, list[float | str]] = {
        co: [0.0] * expected_students for co in range(1, total_outcomes + 1)
    }
    if kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
        question_cols = list(range(_EXCEL_COL_FIRST_MARK, max(_EXCEL_COL_FIRST_MARK, len(headers))))
        co_by_col: dict[int, int] = {}
        for col in question_cols:
            co_token = ws.cell(row=header_row + _EXCEL_ROW_CO_OFFSET, column=col).value
            co_values = _parse_co_values(co_token)
            if len(co_values) != 1:
                continue
            co_by_col[col] = co_values[0]
            max_marks = _to_float(ws.cell(row=header_row + _EXCEL_ROW_MAX_OFFSET, column=col).value)
            if max_marks is not None and 1 <= co_values[0] <= total_outcomes:
                max_by_co[co_values[0]] += max_marks
        if not co_by_col:
            raise ValidationError(
                t("instructor.validation.final_report.direct_component_marks_shape_invalid", component=component.name)
            )
        active_cos = [co for co in range(1, total_outcomes + 1) if max_by_co[co] > 0]
        for idx in range(expected_students):
            row = first_data_row + idx
            absent = False
            co_totals = {co: 0.0 for co in active_cos}
            for col, co_value in co_by_col.items():
                cell_value = ws.cell(row=row, column=col).value
                if _is_absent(cell_value):
                    absent = True
                    break
                numeric = _to_float(cell_value)
                if numeric is None:
                    continue
                if 1 <= co_value <= total_outcomes:
                    co_totals[co_value] += numeric
            if absent:
                for co in active_cos:
                    marks_by_co[co][idx] = CO_REPORT_ABSENT_TOKEN
            else:
                for co in active_cos:
                    marks_by_co[co][idx] = _round2(co_totals[co])
    else:
        covered_cos: list[int] = []
        for col in range(_EXCEL_COL_FIRST_MARK + 1, len(headers) + 1):
            co_token = ws.cell(row=header_row + _EXCEL_ROW_CO_OFFSET, column=col).value
            co_values = _parse_co_values(co_token)
            if len(co_values) != 1:
                continue
            co_value = co_values[0]
            if co_value < 1 or co_value > total_outcomes:
                continue
            covered_cos.append(co_value)
            max_marks = _to_float(ws.cell(row=header_row + _EXCEL_ROW_MAX_OFFSET, column=col).value)
            if max_marks is not None:
                max_by_co[co_value] = max_marks
        for idx in range(expected_students):
            row = first_data_row + idx
            total_cell_value = ws.cell(row=row, column=_EXCEL_COL_FIRST_MARK).value
            if _is_absent(total_cell_value):
                for co in covered_cos:
                    marks_by_co[co][idx] = CO_REPORT_ABSENT_TOKEN
                continue
            total_numeric = _to_float(total_cell_value)
            if total_numeric is None:
                total_numeric = 0.0
            split_values = _split_equal_with_residual(total_numeric, len(covered_cos))
            for co_idx, co in enumerate(covered_cos):
                marks_by_co[co][idx] = split_values[co_idx]

    return _DirectComponentComputed(
        name=component.name,
        weight=component.weight,
        max_by_co=max_by_co,
        marks_by_co=marks_by_co,
    )


def _compute_indirect_component_marks(
    *,
    workbook: Any,
    component: _IndirectComponent,
    students: list[tuple[str, str]],
    spec: dict[str, Any] | None,
    total_outcomes: int,
) -> _IndirectComponentComputed:
    if spec is None:
        raise ValidationError(
            t("instructor.validation.final_report.direct_component_sheet_missing", component=component.name)
        )
    sheet_name = spec.get(LAYOUT_SHEET_SPEC_KEY_NAME)
    header_row = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADER_ROW)
    headers = spec.get(LAYOUT_SHEET_SPEC_KEY_HEADERS)
    kind = spec.get(LAYOUT_SHEET_SPEC_KEY_KIND)
    if (
        not isinstance(sheet_name, str)
        or not isinstance(header_row, int)
        or not isinstance(headers, list)
        or kind != LAYOUT_SHEET_KIND_INDIRECT
    ):
        raise ValidationError(t("instructor.validation.final_report.layout_manifest_invalid"))
    ws = workbook[sheet_name]
    first_data_row = header_row + _EXCEL_ROW_HEADER_OFFSET_INDIRECT
    expected_students = len(students)

    marks_by_co: dict[int, list[float | str]] = {
        co: [0.0] * expected_students for co in range(1, total_outcomes + 1)
    }
    for idx in range(expected_students):
        row = first_data_row + idx
        for co in range(1, total_outcomes + 1):
            col = _EXCEL_COL_STUDENT_NAME + co
            cell_value = ws.cell(row=row, column=col).value
            if _is_absent(cell_value):
                marks_by_co[co][idx] = CO_REPORT_ABSENT_TOKEN
                continue
            numeric = _to_float(cell_value)
            marks_by_co[co][idx] = _round2(numeric if numeric is not None else 0.0)
    return _IndirectComponentComputed(
        name=component.name,
        weight=component.weight,
        marks_by_co=marks_by_co,
    )


def _split_equal_with_residual(total: float, count: int) -> list[float]:
    if count <= 0:
        return []
    if count == 1:
        return [_round2(total)]
    base = _round2(total / count)
    values = [base] * count
    values[-1] = _round2(total - sum(values[:-1]))
    return values


def _is_absent(value: Any) -> bool:
    token = normalize(value)
    return token == normalize(CO_REPORT_ABSENT_TOKEN)


def _to_float(value: Any) -> float | None:
    parsed = coerce_excel_number(value)
    if isinstance(parsed, bool):
        return None
    if isinstance(parsed, (int, float)):
        return float(parsed)
    return None


def _round2(value: float) -> float:
    return round(float(value), CO_REPORT_MAX_DECIMAL_PLACES)


def _xlsxwriter_formats(workbook: Any) -> dict[str, Any]:
    header_style, body_style = _style_registry_for_setup()
    border_enabled = int(body_style.get(_STYLE_KEY_BORDER, 1)) > 0
    header_border_enabled = int(header_style.get(_STYLE_KEY_BORDER, 1)) > 0
    header_bg = _color_without_hash(str(header_style.get(_STYLE_KEY_BG_COLOR, ""))) or "D9EAD3"
    return {
        "header": workbook.add_format(
            {
                "bold": bool(header_style.get(_STYLE_KEY_BOLD, True)),
                "border": 1 if header_border_enabled else 0,
                "align": str(header_style.get(_STYLE_KEY_ALIGN, _ALIGN_CENTER)),
                "valign": str(header_style.get(_STYLE_KEY_VALIGN, _ALIGN_CENTER)).replace(_ALIGN_VCENTER, _ALIGN_CENTER),
                "text_wrap": True,
                "fg_color": header_bg,
                "pattern": 1,
            }
        ),
        "body": workbook.add_format(
            {
                "border": 1 if border_enabled else 0,
                "valign": _ALIGN_VCENTER,
            }
        ),
        "body_center": workbook.add_format(
            {
                "border": 1 if border_enabled else 0,
                "align": _ALIGN_CENTER,
                "valign": _ALIGN_VCENTER,
            }
        ),
        "body_wrap": workbook.add_format(
            {
                "border": 1 if border_enabled else 0,
                "align": "left",
                "valign": _ALIGN_VCENTER,
                "text_wrap": True,
            }
        ),
    }


def _xlsxwriter_write_report_metadata(
    ws: Any,
    *,
    metadata_rows: list[tuple[str, Any]],
    formats: dict[str, Any],
) -> int:
    ws.write(0, 1, COURSE_METADATA_HEADERS[0], formats["header"])
    ws.write(0, 2, COURSE_METADATA_HEADERS[1], formats["header"])
    for idx, (field, value) in enumerate(metadata_rows, start=1):
        ws.write(idx, 1, field, formats["body"])
        ws.write(idx, 2, value, formats["body_wrap"])
    return len(metadata_rows) + 2


def _xlsxwriter_set_report_metadata_column_widths(
    ws: Any,
    *,
    metadata_rows: list[tuple[str, Any]],
    formats: dict[str, Any],
) -> None:
    sample_rows: list[list[Any]] = [["", COURSE_METADATA_HEADERS[0], COURSE_METADATA_HEADERS[1]]]
    sample_rows.extend(["", field, value] for field, value in metadata_rows)
    widths = _compute_sampled_column_widths(sample_rows, 2)
    ws.set_column(1, 1, widths.get(1, 8))
    ws.set_column(2, 2, widths.get(2, 8))


def _xlsxwriter_apply_layout(ws: Any, *, header_row_index: int, paper_size: int, landscape: bool) -> None:
    if landscape:
        ws.set_landscape()
    else:
        ws.set_portrait()
    ws.set_paper(paper_size)
    ws.fit_to_pages(1, 0)
    ws.protect()
    ws.set_selection(header_row_index, 0, header_row_index, 0)


def _xlsxwriter_write_course_metadata_sheet(
    workbook: Any,
    *,
    metadata_rows: list[tuple[str, Any]],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet(COURSE_METADATA_SHEET)
    ws.write(0, 0, COURSE_METADATA_HEADERS[0], formats["header"])
    ws.write(0, 1, COURSE_METADATA_HEADERS[1], formats["header"])
    filtered_rows = [
        (field, value)
        for field, value in metadata_rows
        if normalize(field) != normalize(COURSE_METADATA_FACULTY_NAME_KEY)
    ]
    for idx, (field, value) in enumerate(filtered_rows, start=1):
        ws.write(idx, 0, field, formats["body"])
        ws.write(idx, 1, value, formats["body"])
    _xlsxwriter_apply_layout(ws, header_row_index=0, paper_size=9, landscape=False)


def _xlsxwriter_write_direct_sheet(
    workbook: Any,
    *,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    components: list[_DirectComponentComputed],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet(f"CO{co_index}{CO_REPORT_DIRECT_SHEET_SUFFIX}")
    report_metadata_rows = _report_metadata_rows(
        metadata_rows,
        co_index=co_index,
        outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_TEMPLATE,
    )
    header_row_index = _xlsxwriter_write_report_metadata(ws, metadata_rows=report_metadata_rows, formats=formats) - 1
    _xlsxwriter_set_report_metadata_column_widths(ws, metadata_rows=report_metadata_rows, formats=formats)
    ws.repeat_rows(0, header_row_index)
    ws.freeze_panes(header_row_index + 1, _EXCEL_COL_FIRST_MARK - 1)
    active_components = [component for component in components if component.max_by_co.get(co_index, 0.0) > 0]
    total_weight = _round2(sum(component.weight for component in active_components))
    headers = [CO_REPORT_HEADER_SERIAL, CO_REPORT_HEADER_REG_NO, CO_REPORT_HEADER_STUDENT_NAME]
    for component in active_components:
        max_marks = _round2(component.max_by_co.get(co_index, 0.0))
        headers.append(f"{component.name} ({max_marks:g})")
        headers.append(f"{component.name} ({component.weight:g}{CO_REPORT_PERCENT_SYMBOL})")
    headers.append(f"{CO_REPORT_HEADER_TOTAL} ({total_weight:g}{CO_REPORT_PERCENT_SYMBOL})")
    headers.append(CO_REPORT_HEADER_TOTAL_100)
    headers.append(_ratio_total_header(DIRECT_RATIO))
    for col, value in enumerate(headers, start=0):
        ws.write(header_row_index, col, value, formats["header"])

    student_count = len(students)
    marks_by_component: list[list[Any]] = [
        component.marks_by_co.get(co_index, [0.0] * student_count)
        for component in active_components
    ]
    for idx, (reg_no, student_name) in enumerate(students, start=1):
        row_index = header_row_index + idx
        row_values: list[Any] = [idx, reg_no, student_name]
        absent = False
        weighted_total = 0.0
        for component, component_marks in zip(active_components, marks_by_component):
            max_marks = component.max_by_co.get(co_index, 0.0)
            raw = component_marks[idx - 1]
            if isinstance(raw, str) and _is_absent(raw):
                absent = True
                row_values.extend([CO_REPORT_ABSENT_TOKEN, CO_REPORT_ABSENT_TOKEN])
                continue
            raw_numeric = float(raw) if isinstance(raw, (int, float)) else 0.0
            weighted = _round2((raw_numeric * component.weight / max_marks) if max_marks > 0 else 0.0)
            weighted_total += weighted
            row_values.extend([_round2(raw_numeric), weighted])
        if absent:
            row_values.extend([CO_REPORT_NOT_APPLICABLE_TOKEN, CO_REPORT_NOT_APPLICABLE_TOKEN, CO_REPORT_NOT_APPLICABLE_TOKEN])
        else:
            total_100 = _round2((weighted_total * 100.0 / total_weight) if total_weight > 0 else 0.0)
            total_ratio = _round2(total_100 * DIRECT_RATIO)
            row_values.extend([_round2(weighted_total), total_100, total_ratio])

        for col, value in enumerate(row_values, start=0):
            if col == 2:
                fmt = formats["body_wrap"]
            elif col >= 3:
                fmt = formats["body_center"]
            else:
                fmt = formats["body"]
            ws.write(row_index, col, value, fmt)

    _xlsxwriter_apply_layout(ws, header_row_index=header_row_index, paper_size=8, landscape=True)


def _xlsxwriter_write_indirect_sheet(
    workbook: Any,
    *,
    co_index: int,
    metadata_rows: list[tuple[str, Any]],
    students: list[tuple[str, str]],
    components: list[_IndirectComponentComputed],
    formats: dict[str, Any],
) -> None:
    ws = workbook.add_worksheet(f"CO{co_index}{CO_REPORT_INDIRECT_SHEET_SUFFIX}")
    report_metadata_rows = _report_metadata_rows(
        metadata_rows,
        co_index=co_index,
        outcome_value_template=CO_REPORT_METADATA_OUTCOME_VALUE_INDIRECT_TEMPLATE,
    )
    header_row_index = _xlsxwriter_write_report_metadata(ws, metadata_rows=report_metadata_rows, formats=formats) - 1
    _xlsxwriter_set_report_metadata_column_widths(ws, metadata_rows=report_metadata_rows, formats=formats)
    ws.repeat_rows(0, header_row_index)
    ws.freeze_panes(header_row_index + 1, _EXCEL_COL_FIRST_MARK - 1)

    active_components = [
        component
        for component in components
        if any(not (isinstance(value, float) and value == 0.0) for value in component.marks_by_co.get(co_index, []))
    ]
    total_weight = _round2(sum(component.weight for component in active_components))
    headers = [CO_REPORT_HEADER_SERIAL, CO_REPORT_HEADER_REG_NO, CO_REPORT_HEADER_STUDENT_NAME]
    scaled_max_value = max(0, LIKERT_MAX - LIKERT_MIN)
    has_single_component = len(active_components) == 1
    for component in active_components:
        headers.append(f"{component.name} ({LIKERT_MIN}-{LIKERT_MAX})")
        headers.append(f"{component.name} ({CO_REPORT_SCALED_LABEL_TEMPLATE.format(max_value=scaled_max_value)})")
        if not has_single_component:
            headers.append(f"{component.name} ({component.weight:g}{CO_REPORT_PERCENT_SYMBOL})")
    headers.append(CO_REPORT_HEADER_TOTAL_100)
    headers.append(_ratio_total_header(INDIRECT_RATIO))
    for col, value in enumerate(headers, start=0):
        ws.write(header_row_index, col, value, formats["header"])

    student_count = len(students)
    marks_by_component: list[list[Any]] = [
        component.marks_by_co.get(co_index, [0.0] * student_count)
        for component in active_components
    ]
    denominator = float(max(1, scaled_max_value))
    for idx, (reg_no, student_name) in enumerate(students, start=1):
        row_index = header_row_index + idx
        row_values: list[Any] = [idx, reg_no, student_name]
        absent = False
        total_weighted = 0.0
        for component, component_marks in zip(active_components, marks_by_component):
            raw = component_marks[idx - 1]
            if isinstance(raw, str) and _is_absent(raw):
                absent = True
                row_values.extend([CO_REPORT_ABSENT_TOKEN, CO_REPORT_ABSENT_TOKEN])
                if not has_single_component:
                    row_values.append(CO_REPORT_ABSENT_TOKEN)
                continue
            raw_numeric = float(raw) if isinstance(raw, (int, float)) else 0.0
            scaled_raw = _round2(raw_numeric - LIKERT_MIN)
            scaled_raw = max(0.0, min(float(scaled_max_value), scaled_raw))
            if has_single_component:
                total_weighted = _round2((scaled_raw / denominator) * 100.0)
                row_values.extend([_round2(raw_numeric), scaled_raw])
            else:
                weighted = _round2((scaled_raw / denominator) * component.weight)
                total_weighted += weighted
                row_values.extend([_round2(raw_numeric), scaled_raw, weighted])
        if absent:
            row_values.extend([CO_REPORT_NOT_APPLICABLE_TOKEN, CO_REPORT_NOT_APPLICABLE_TOKEN])
        else:
            if has_single_component:
                total_100 = _round2(total_weighted)
            else:
                total_100 = _round2((total_weighted * 100.0 / total_weight) if total_weight > 0 else 0.0)
            total_ratio = _round2(total_100 * INDIRECT_RATIO)
            row_values.extend([total_100, total_ratio])

        for col, value in enumerate(row_values, start=0):
            if col == 2:
                fmt = formats["body_wrap"]
            elif col >= 3:
                fmt = formats["body_center"]
            else:
                fmt = formats["body"]
            ws.write(row_index, col, value, fmt)

    _xlsxwriter_apply_layout(ws, header_row_index=header_row_index, paper_size=9, landscape=False)

def _add_system_layout_sheet(workbook: Any, manifest_text: str, manifest_hash: str) -> None:
    layout_ws = workbook.add_worksheet(SYSTEM_LAYOUT_SHEET)
    layout_ws.write_row(0, 0, [SYSTEM_LAYOUT_MANIFEST_HEADER, SYSTEM_LAYOUT_MANIFEST_HASH_HEADER])
    layout_ws.write_row(1, 0, [manifest_text, manifest_hash])
    layout_ws.hide()
def _write_final_report_workbook_xlsxwriter(
    *,
    xlsxwriter_module: Any,
    output_path: Path,
    metadata_rows: list[tuple[str, Any]],
    total_outcomes: int,
    students: list[tuple[str, str]],
    direct_components: list[_DirectComponentComputed],
    indirect_components: list[_IndirectComponentComputed],
    template_id: str,
    template_hash: str,
) -> None:
    workbook = xlsxwriter_module.Workbook(str(output_path), {"constant_memory": True})
    sheet_order: list[str] = []
    try:
        formats = _xlsxwriter_formats(workbook)
        _xlsxwriter_write_course_metadata_sheet(workbook, metadata_rows=metadata_rows, formats=formats)
        sheet_order.append(COURSE_METADATA_SHEET)
        for co_index in range(1, total_outcomes + 1):
            _xlsxwriter_write_direct_sheet(
                workbook,
                co_index=co_index,
                metadata_rows=metadata_rows,
                students=students,
                components=direct_components,
                formats=formats,
            )
            sheet_order.append(f"CO{co_index}{CO_REPORT_DIRECT_SHEET_SUFFIX}")
            _xlsxwriter_write_indirect_sheet(
                workbook,
                co_index=co_index,
                metadata_rows=metadata_rows,
                students=students,
                components=indirect_components,
                formats=formats,
            )
            sheet_order.append(f"CO{co_index}{CO_REPORT_INDIRECT_SHEET_SUFFIX}")

        hash_ws = workbook.add_worksheet(SYSTEM_HASH_SHEET)
        hash_ws.write(0, 0, SYSTEM_HASH_TEMPLATE_ID_HEADER)
        hash_ws.write(0, 1, SYSTEM_HASH_TEMPLATE_HASH_HEADER)
        hash_ws.write(1, 0, template_id)
        hash_ws.write(1, 1, template_hash)
        hash_ws.hide()
        sheet_order.append(SYSTEM_HASH_SHEET)

        manifest = {
            "schema_version": _INTEGRITY_SCHEMA_VERSION,
            "template_id": template_id,
            "template_hash": template_hash,
            "sheet_order": sheet_order,
            "sheets": [{"name": name, "hash": sign_payload(name)} for name in sheet_order],
        }
        manifest_text = json.dumps(manifest, sort_keys=True, separators=(",", ":"), ensure_ascii=True)
        manifest_hash = sign_payload(manifest_text)
        _add_system_layout_sheet(workbook, manifest_text, manifest_hash)
    finally:
        workbook.close()

def _normalize_page_setup_fit(path: Path) -> None:
    try:
        import openpyxl
    except Exception:
        return
    with path.open("rb") as handle:
        wb = openpyxl.load_workbook(handle)
    try:
        for ws in wb.worksheets:
            if ws.title in {SYSTEM_HASH_SHEET, SYSTEM_LAYOUT_SHEET}:
                continue
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
        wb.save(path)
    finally:
        wb.close()

def _report_metadata_rows(
    metadata_rows: list[tuple[str, Any]],
    *,
    co_index: int,
    outcome_value_template: str,
) -> list[tuple[str, Any]]:
    filtered: list[tuple[str, Any]] = []
    drop_keys = {
        normalize(COURSE_METADATA_FACULTY_NAME_KEY),
        normalize(COURSE_METADATA_TOTAL_OUTCOMES_KEY),
    }
    for field, value in metadata_rows:
        if normalize(field) in drop_keys:
            continue
        filtered.append((field, value))
    filtered.append(
        (
            CO_REPORT_METADATA_OUTCOME_FIELD,
            outcome_value_template.format(co=co_index),
        )
    )
    return filtered
def _ratio_total_header(ratio: float) -> str:
    percent = ratio * 100.0
    if abs(percent - round(percent)) <= 1e-9:
        token = f"{int(round(percent))}"
    else:
        token = f"{percent:g}"
    return CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE.format(ratio=token)

