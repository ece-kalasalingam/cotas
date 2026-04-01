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
from domain.template_strategy_router import (
    generate_workbook,
    generate_workbooks,
    resolve_template_id_from_workbook_path,
)


def generate_course_details_template(output_path: Path) -> Path:
    """Generate course details template.
    
    Args:
        output_path: Parameter value (Path).
    
    Returns:
        Path: Return value.
    
    Raises:
        None.
    """
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
    """Generate marks template from course details.
    
    Args:
        course_details_path: Parameter value (Path).
        output_path: Parameter value (Path).
    
    Returns:
        Path: Return value.
    
    Raises:
        None.
    """
    template_id = resolve_template_id_from_workbook_path(course_details_path)
    result = generate_workbooks(
        template_id=template_id,
        workbook_paths=[course_details_path],
        output_dir=output_path.parent,
        workbook_kind="marks_template",
        context={
            "overwrite_existing": True,
            "output_path_overrides": {str(course_details_path): str(output_path)},
        },
    )
    generated = result.get("generated_workbook_paths", []) if isinstance(result, dict) else []
    if isinstance(generated, list) and generated:
        output = str(generated[0]).strip()
        if output:
            return Path(output)
    return output_path


def test_registry_declares_v2_marks_dynamic_sheet_templates() -> None:
    """Test registry declares v2 marks dynamic sheet templates.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Test instructor sheetops does not duplicate marks dynamic header tokens.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Test marks template uses shared format bundle only.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Test xlsx style policy constants live in excel layout module.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
    """Test generated marks sheet headers match registry dynamic resolvers.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
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
