from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import (
    LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
    LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
    LAYOUT_SHEET_KIND_INDIRECT,
)
from common.registry import (
    COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE,
    COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
    COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT,
    get_dynamic_sheet_template,
    resolve_dynamic_sheet_headers,
)
from domain.template_strategy_router import generate_workbook, resolve_template_id_from_workbook_path


def generate_course_details_template(output_path: Path) -> Path:
    result = generate_workbook(
        template_id="COURSE_SETUP_V2",
        output_path=output_path,
        workbook_name=output_path.name,
        workbook_kind="course_details_template",
    )
    output = getattr(result, "output_path", None)
    if isinstance(output, str) and output.strip():
        return Path(output)
    return output_path


def generate_marks_template_from_course_details(course_details_path: Path, output_path: Path) -> Path:
    template_id = resolve_template_id_from_workbook_path(course_details_path)
    result = generate_workbook(
        template_id=template_id,
        output_path=output_path,
        workbook_name=output_path.name,
        workbook_kind="marks_template",
        context={"course_details_path": str(course_details_path)},
    )
    output = getattr(result, "output_path", None)
    if isinstance(output, str) and output.strip():
        return Path(output)
    return output_path


def test_registry_declares_v2_marks_dynamic_sheet_templates() -> None:
    direct_co = get_dynamic_sheet_template("COURSE_SETUP_V2", COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE)
    direct_non_co = get_dynamic_sheet_template(
        "COURSE_SETUP_V2",
        COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
    )
    indirect = get_dynamic_sheet_template("COURSE_SETUP_V2", COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT)

    assert direct_co["header_kind"] == "dynamic"
    assert direct_non_co["header_kind"] == "dynamic"
    assert indirect["header_kind"] == "dynamic"
    assert direct_co["header_resolver"] == "course_setup.marks_direct_co_wise_headers"
    assert direct_non_co["header_resolver"] == "course_setup.marks_direct_non_co_wise_headers"
    assert indirect["header_resolver"] == "course_setup.marks_indirect_headers"


def test_instructor_sheetops_does_not_duplicate_marks_dynamic_header_tokens() -> None:
    source = (
        Path(__file__).resolve().parents[1]
        / "domain"
        / "template_versions"
        / "course_setup_v2_impl"
        / "instructor_engine_sheetops.py"
    ).read_text(encoding="utf-8")
    assert "MARKS_ENTRY_QUESTION_PREFIX" not in source
    assert "MARKS_ENTRY_CO_MARKS_LABEL_PREFIX" not in source
    assert "resolve_dynamic_sheet_headers" in source


def test_marks_template_uses_shared_format_bundle_only() -> None:
    source = (
        Path(__file__).resolve().parents[1]
        / "domain"
        / "template_versions"
        / "course_setup_v2_impl"
        / "marks_template.py"
    ).read_text(encoding="utf-8")
    assert "build_marks_template_xlsxwriter_formats" in source
    assert ".add_format(" not in source


def test_xlsx_style_policy_constants_live_in_excel_layout_module() -> None:
    constants_source = (Path(__file__).resolve().parents[1] / "common" / "constants.py").read_text(
        encoding="utf-8"
    )
    layout_source = (
        Path(__file__).resolve().parents[1] / "common" / "excel_sheet_layout.py"
    ).read_text(encoding="utf-8")
    assert "XLSX_AUTOFIT_SAMPLE_ROWS" not in constants_source
    assert "XLSX_PAPER_SIZE_A4" not in constants_source
    assert "XLSX_AUTOFIT_SAMPLE_ROWS" in layout_source
    assert "XLSX_PAPER_SIZE_A4" in layout_source


def test_generated_marks_sheet_headers_match_registry_dynamic_resolvers(tmp_path: Path) -> None:
    course_details = tmp_path / "course_details.xlsx"
    marks_template = tmp_path / "marks_template.xlsx"
    generate_course_details_template(course_details)
    generate_marks_template_from_course_details(course_details, marks_template)

    wb = openpyxl.load_workbook(marks_template, data_only=False)
    try:
        manifest_text = wb["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest_text, str)
        manifest = json.loads(manifest_text)

        for spec in manifest.get("sheets", []):
            kind = str(spec.get("kind", ""))
            if kind not in {
                LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
                LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
                LAYOUT_SHEET_KIND_INDIRECT,
            }:
                continue
            headers = spec.get("headers", [])
            assert isinstance(headers, list)
            if kind == LAYOUT_SHEET_KIND_DIRECT_CO_WISE:
                question_count = max(1, len(headers) - 4)
                expected = resolve_dynamic_sheet_headers(
                    "COURSE_SETUP_V2",
                    sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE,
                    context={"question_count": question_count},
                )
            elif kind == LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE:
                ws = wb[str(spec["name"])]
                header_row = int(spec["header_row"])
                covered_cos: list[int] = []
                for col in range(5, len(headers) + 1):
                    token = str(ws.cell(row=header_row + 1, column=col).value or "").strip().upper()
                    if token.startswith("CO"):
                        try:
                            covered_cos.append(int(token[2:].strip()))
                        except ValueError:
                            continue
                expected = resolve_dynamic_sheet_headers(
                    "COURSE_SETUP_V2",
                    sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
                    context={"covered_cos": covered_cos},
                )
            else:
                total_outcomes = max(1, len(headers) - 3)
                expected = resolve_dynamic_sheet_headers(
                    "COURSE_SETUP_V2",
                    sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT,
                    context={"total_outcomes": total_outcomes},
                )
            assert tuple(headers) == expected
    finally:
        wb.close()
