from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import ID_COURSE_SETUP
from common.jobs import CancellationToken
from domain.template_strategy_router import generate_workbook, generate_workbooks, validate_workbooks
from domain.template_versions.course_setup_v2_impl.co_attainment import generate_final_report_workbook
from services.instructor_workflow_service import InstructorWorkflowService


def _set_course_section(course_details_path: Path, section: str) -> None:
    """Set course section.
    
    Args:
        course_details_path: Parameter value (Path).
        section: Parameter value (str).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(course_details_path)
    try:
        ws = wb["Course_Metadata"]
        row = 2
        while True:
            key = str(ws.cell(row=row, column=1).value or "").strip().casefold()
            value = ws.cell(row=row, column=2).value
            if not key and (value is None or str(value).strip() == ""):
                break
            if key == "section":
                ws.cell(row=row, column=2, value=section)
                break
            row += 1
        wb.save(course_details_path)
    finally:
        wb.close()


def _fill_marks_workbook(marks_path: Path, mark_value: float = 1.0) -> None:
    """Fill marks workbook.
    
    Args:
        marks_path: Parameter value (Path).
        mark_value: Parameter value (float).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(marks_path)
    try:
        manifest_text = wb["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest_text, str)
        manifest = json.loads(manifest_text)
        for spec in manifest.get("sheets", []):
            kind = spec.get("kind")
            if kind not in {"direct_co_wise", "direct_non_co_wise", "indirect"}:
                continue
            ws = wb[str(spec["name"])]
            header_row = int(spec["header_row"])
            header_count = len(spec["headers"])
            if kind == "indirect":
                first_data_row = header_row + 1
                mark_cols = range(4, header_count + 1)
            elif kind == "direct_non_co_wise":
                first_data_row = header_row + 3
                mark_cols = range(4, 5)
            else:
                first_data_row = header_row + 3
                mark_cols = range(4, header_count)
            row = first_data_row
            while True:
                reg_no = ws.cell(row=row, column=2).value
                student_name = ws.cell(row=row, column=3).value
                if reg_no is None and student_name is None:
                    break
                for col in mark_cols:
                    ws.cell(row=row, column=col, value=mark_value)
                row += 1
        wb.save(marks_path)
    finally:
        wb.close()


def _safe_value(value: Any) -> Any:
    """Safe value.
    
    Args:
        value: Parameter value (Any).
    
    Returns:
        Any: Return value.
    
    Raises:
        None.
    """
    if value is None:
        return None
    if isinstance(value, (int, float, bool, str)):
        return value
    return str(value)


def _workbook_parity_fingerprint(path: Path) -> str:
    """Workbook parity fingerprint.
    
    Args:
        path: Parameter value (Path).
    
    Returns:
        str: Return value.
    
    Raises:
        None.
    """
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        payload: dict[str, Any] = {
            "sheetnames": list(wb.sheetnames),
            "workbook_security": {
                "lockStructure": bool(getattr(getattr(wb, "security", None), "lockStructure", False)),
                "workbookPassword": str(getattr(getattr(wb, "security", None), "workbookPassword", "") or ""),
            },
            "sheets": [],
        }
        for sheet in wb.worksheets:
            protection = sheet.protection
            sheet_payload: dict[str, Any] = {
                "title": sheet.title,
                "state": sheet.sheet_state,
                "max_row": int(sheet.max_row or 0),
                "max_col": int(sheet.max_column or 0),
                "protection": {
                    "sheet": bool(getattr(protection, "sheet", False)),
                    "password": str(getattr(protection, "password", "") or ""),
                    "formatCells": bool(getattr(protection, "formatCells", False)),
                    "formatRows": bool(getattr(protection, "formatRows", False)),
                    "formatColumns": bool(getattr(protection, "formatColumns", False)),
                    "insertRows": bool(getattr(protection, "insertRows", False)),
                    "insertColumns": bool(getattr(protection, "insertColumns", False)),
                    "deleteRows": bool(getattr(protection, "deleteRows", False)),
                    "deleteColumns": bool(getattr(protection, "deleteColumns", False)),
                    "sort": bool(getattr(protection, "sort", False)),
                    "autoFilter": bool(getattr(protection, "autoFilter", False)),
                    "pivotTables": bool(getattr(protection, "pivotTables", False)),
                    "selectLockedCells": bool(getattr(protection, "selectLockedCells", False)),
                    "selectUnlockedCells": bool(getattr(protection, "selectUnlockedCells", False)),
                },
                "merged_ranges": sorted(str(rng) for rng in sheet.merged_cells.ranges),
                "column_dimensions": [],
                "row_dimensions": [],
                "cells": [],
            }
            for key, dim in sorted(sheet.column_dimensions.items()):
                width = getattr(dim, "width", None)
                hidden = bool(getattr(dim, "hidden", False))
                custom = bool(getattr(dim, "customWidth", False))
                if width is None and not hidden and not custom:
                    continue
                sheet_payload["column_dimensions"].append(
                    [str(key), float(width) if isinstance(width, (int, float)) else None, hidden, custom]
                )
            for idx, dim in sorted(sheet.row_dimensions.items()):
                height = getattr(dim, "height", None)
                hidden = bool(getattr(dim, "hidden", False))
                if height is None and not hidden:
                    continue
                sheet_payload["row_dimensions"].append(
                    [int(idx), float(height) if isinstance(height, (int, float)) else None, hidden]
                )
            max_row = int(sheet.max_row or 0)
            max_col = int(sheet.max_column or 0)
            if max_row > 0 and max_col > 0:
                for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                    for cell in row:
                        if cell.value is None and int(getattr(cell, "style_id", 0) or 0) == 0:
                            continue
                        cell_prot = getattr(cell, "protection", None)
                        sheet_payload["cells"].append(
                            [
                                cell.coordinate,
                                _safe_value(cell.value),
                                str(cell.data_type),
                                str(cell.number_format),
                                int(getattr(cell, "style_id", 0) or 0),
                                bool(getattr(cell_prot, "locked", True)) if cell_prot is not None else True,
                                bool(getattr(cell_prot, "hidden", False)) if cell_prot is not None else False,
                            ]
                        )
            payload["sheets"].append(sheet_payload)
        canonical = json.dumps(payload, sort_keys=True, separators=(",", ":"), ensure_ascii=True)
        return hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    finally:
        wb.close()


def test_v2_workbook_output_parity_contract(tmp_path: Path) -> None:
    """Test v2 workbook output parity contract.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    instructor = InstructorWorkflowService()
    course_details = tmp_path / "course_details.xlsx"
    marks_template = tmp_path / "marks_template.xlsx"
    final_report = tmp_path / "final_report.xlsx"
    co_attainment = tmp_path / "co_attainment.xlsx"

    generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=course_details,
        workbook_name=course_details.name,
        workbook_kind="course_details_template",
        cancel_token=CancellationToken(),
    )
    _set_course_section(course_details, "A")

    validate_workbooks(
        template_id=ID_COURSE_SETUP,
        workbook_paths=[course_details],
        workbook_kind="course_details",
        cancel_token=CancellationToken(),
    )
    batch_result = generate_workbooks(
        template_id=ID_COURSE_SETUP,
        workbook_paths=[course_details],
        output_dir=marks_template.parent,
        workbook_kind="marks_template",
        cancel_token=CancellationToken(),
        context={
            "overwrite_existing": True,
            "output_path_overrides": {str(course_details): str(marks_template)},
        },
    )
    assert int(batch_result.get("generated", 0)) == 1
    _fill_marks_workbook(marks_template, mark_value=1.0)

    generate_final_report_workbook(
        filled_marks_path=marks_template,
        output_path=final_report,
        cancel_token=CancellationToken(),
    )
    generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=co_attainment,
        workbook_name=co_attainment.name,
        workbook_kind="co_attainment",
        cancel_token=CancellationToken(),
        context={
            "source_paths": [str(final_report)],
            "thresholds": (40.0, 60.0, 75.0),
            "co_attainment_percent": 60.0,
            "co_attainment_level": 2,
        },
    )

    marks_hash = _workbook_parity_fingerprint(marks_template)
    final_hash = _workbook_parity_fingerprint(final_report)
    coa_hash = _workbook_parity_fingerprint(co_attainment)

    # Baseline parity contract for current V2 output behavior.
    expected = {
        "marks_template": "731315de1da47b7c6d741d808f8a70537f6892eae0edf1efd122b4f6f06ea0b9",
        "final_report": "c8459723f44b8ba77bfed34190167338db2400ba04292d16fec49c7493a00d03",
        "co_attainment": "f7d5514fb022c9666a1636642780c70bc59dbe3ef2bf25c1540627c56682ad45",
    }
    actual = {
        "marks_template": marks_hash,
        "final_report": final_hash,
        "co_attainment": coa_hash,
    }
    assert actual == expected, f"Workbook parity hash mismatch:\nexpected={expected}\nactual={actual}"
