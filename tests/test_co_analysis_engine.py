from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")
pytest.importorskip("xlsxwriter")

from common.constants import ID_COURSE_SETUP
from common.exceptions import ValidationError
from domain.co_analysis_engine import _classify_validation_reason, generate_co_analysis_workbook
from domain.template_strategy_router import (
    generate_workbook,
    generate_workbooks,
    resolve_template_id_from_workbook_path,
)


def generate_course_details_template(output_path: Path) -> Path:
    result = generate_workbook(
        template_id=ID_COURSE_SETUP,
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


def _build_filled_marks_workbook(tmp_path: Path) -> Path:
    course_details = tmp_path / "course_details.xlsx"
    marks = tmp_path / "marks_template.xlsx"
    generate_course_details_template(course_details)
    generate_marks_template_from_course_details(course_details, marks)

    wb = openpyxl.load_workbook(marks)
    try:
        manifest_text = wb["__SYSTEM_LAYOUT__"]["A2"].value
        assert isinstance(manifest_text, str)
        manifest = json.loads(manifest_text)
        for spec in manifest.get("sheets", []):
            kind = spec.get("kind")
            if kind not in {"direct_co_wise", "direct_non_co_wise", "indirect"}:
                continue
            ws = wb[spec["name"]]
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
                    ws.cell(row=row, column=col).value = 1
                row += 1
        wb.save(marks)
    finally:
        wb.close()
    return marks


def test_generate_co_analysis_workbook_creates_direct_indirect_and_co_sheets(tmp_path: Path) -> None:
    marks = _build_filled_marks_workbook(tmp_path)
    output = tmp_path / "co_analysis.xlsx"

    generate_co_analysis_workbook(
        [marks],
        output,
        thresholds=(40.0, 60.0, 75.0),
        co_attainment_percent=80.0,
        co_attainment_level=2,
    )

    assert output.exists()
    wb = openpyxl.load_workbook(output, data_only=False)
    try:
        assert "Course_Metadata" in wb.sheetnames
        for co in range(1, 7):
            assert f"CO{co}_Direct" in wb.sheetnames
            assert f"CO{co}_Indirect" in wb.sheetnames
            assert f"CO{co}" in wb.sheetnames

        direct = wb["CO1_Direct"]
        indirect = wb["CO1_Indirect"]
        co_sheet = wb["CO1"]
        assert direct.max_row > 1
        assert indirect.max_row > 1
        assert co_sheet.max_row > 1
        assert co_sheet.cell(row=1, column=2).value == "Course Code"
        assert co_sheet.cell(row=1, column=3).value is not None
    finally:
        wb.close()


def test_build_co_analysis_workbook_routes_co_attainment_via_router(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from domain import co_analysis_engine as engine_mod

    source = tmp_path / "source.xlsx"
    wb_source = openpyxl.Workbook()
    wb_source.save(source)
    wb_source.close()

    captured: dict[str, object] = {}

    class _FakeStrategy:
        def generate_final_report(self, source_path: Path, output_path: Path, *, cancel_token=None) -> None:
            del source_path
            del cancel_token
            wb = openpyxl.Workbook()
            ws_meta = wb.active
            ws_meta.title = "Course_Metadata"
            ws_meta["A1"] = "Field"
            ws_meta["B1"] = "Value"
            ws_meta["A2"] = "Total Outcomes"
            ws_meta["B2"] = 1
            wb.create_sheet("CO1_Direct")
            wb.create_sheet("CO1_Indirect")
            wb.save(output_path)
            wb.close()

    def _fake_generate_workbook(**kwargs):
        captured.update(kwargs)
        out_path = Path(kwargs["output_path"])
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CO1"
        ws["A1"] = "Course Code"
        ws["B1"] = "CS101"
        wb.save(out_path)
        wb.close()
        return kwargs["output_path"]

    monkeypatch.setattr(engine_mod, "read_valid_template_id_from_system_hash_sheet", lambda _wb: ID_COURSE_SETUP)
    monkeypatch.setattr(engine_mod, "_course_metadata_headers", lambda _tid: ("Field", "Value"))
    monkeypatch.setattr(engine_mod, "_style_registry_for_template", lambda _tid: ({"bg_color": "#FFFFFF"}, {}))
    monkeypatch.setattr(engine_mod, "_protect_openpyxl_sheet", lambda _sheet: None)
    monkeypatch.setattr(engine_mod, "_merge_report_sheets", lambda **kwargs: None)
    monkeypatch.setattr(engine_mod, "_read_total_outcomes_from_course_metadata", lambda _wb: 1)
    monkeypatch.setattr(engine_mod, "get_template_strategy", lambda _tid: _FakeStrategy())
    monkeypatch.setattr(engine_mod, "generate_workbook", _fake_generate_workbook)

    output = tmp_path / "co_analysis.xlsx"
    result = engine_mod.build_co_analysis_workbook(
        output,
        [(set(["r1"]), {"section": "A"})],
        source_paths=[source],
        thresholds=(40.0, 60.0, 75.0),
        co_attainment_percent=80.0,
        co_attainment_level=2,
    )

    assert result == output
    assert captured["workbook_kind"] == "co_attainment"
    assert captured["template_id"] == ID_COURSE_SETUP
    context = captured["context"]
    assert isinstance(context, dict)
    assert context["co_attainment_percent"] == 80.0
    assert context["co_attainment_level"] == 2


def test_classify_validation_reason_maps_marks_cohort_mismatch_code() -> None:
    exc = ValidationError("cohort mismatch", code="MARKS_TEMPLATE_COHORT_MISMATCH")
    assert _classify_validation_reason(exc) == "cohort_mismatch"
