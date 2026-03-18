from __future__ import annotations

import builtins
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest

from common.exceptions import AppSystemError, JobCancelledError, ValidationError
from common.workbook_signing import sign_payload
from domain import instructor_report_engine as eng


@dataclass
class _Cell:
    value: Any


class _GridSheet:
    def __init__(self) -> None:
        self._cells: dict[tuple[int, int], Any] = {}
        self._ref: dict[str, Any] = {}

    def set_cell(self, row: int, col: int, value: Any) -> None:
        self._cells[(row, col)] = value

    def set_ref(self, key: str, value: Any) -> None:
        self._ref[key] = value

    def cell(self, *, row: int, column: int) -> _Cell:
        return _Cell(self._cells.get((row, column)))

    def __getitem__(self, key: str) -> _Cell:
        return _Cell(self._ref.get(key))


class _Workbook:
    def __init__(self, sheets: dict[str, Any]) -> None:
        self._sheets = sheets
        self.sheetnames = list(sheets.keys())

    def __getitem__(self, key: str) -> Any:
        return self._sheets[key]


def _valid_integrity_workbook() -> _Workbook:
    hash_sheet = _GridSheet()
    hash_sheet.set_ref("A2", eng.ID_COURSE_SETUP)
    hash_sheet.set_ref("B2", sign_payload(eng.ID_COURSE_SETUP))

    layout_manifest = '{"sheets":[]}'
    layout_sheet = _GridSheet()
    layout_sheet.set_ref("A1", eng.SYSTEM_LAYOUT_MANIFEST_KEY)
    layout_sheet.set_ref("B1", eng.SYSTEM_LAYOUT_MANIFEST_HASH_KEY)
    layout_sheet.set_ref("A2", layout_manifest)
    layout_sheet.set_ref("B2", sign_payload(layout_manifest))

    return _Workbook(
        {
            eng.SYSTEM_HASH_SHEET: hash_sheet,
            eng.SYSTEM_LAYOUT_SHEET: layout_sheet,
        }
    )


def test_generate_final_report_import_and_source_validation_branches(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    source = tmp_path / "missing.xlsx"
    output = tmp_path / "out.xlsx"
    real_import = builtins.__import__

    def _import_fail_openpyxl(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[no-untyped-def]
        if name == "openpyxl":
            raise ModuleNotFoundError("openpyxl")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _import_fail_openpyxl)
    with pytest.raises(ValidationError):
        eng.generate_final_co_report(source, output)

    def _import_fail_xlsxwriter(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[no-untyped-def]
        if name == "xlsxwriter":
            raise ModuleNotFoundError("xlsxwriter")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _import_fail_xlsxwriter)
    with pytest.raises(ValidationError):
        eng.generate_final_co_report(source, output)

    monkeypatch.setattr(builtins, "__import__", real_import)
    with pytest.raises(ValidationError):
        eng.generate_final_co_report(source, output)


def test_generate_final_report_workbook_open_failure_branch(tmp_path: Path) -> None:
    source = tmp_path / "not_a_workbook.xlsx"
    source.write_text("not-an-excel-workbook", encoding="utf-8")
    output = tmp_path / "out.xlsx"
    with pytest.raises(ValidationError):
        eng.generate_final_co_report(source, output)


def test_validate_source_workbook_integrity_branches(monkeypatch: pytest.MonkeyPatch) -> None:
    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(_Workbook({}))

    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(_Workbook({eng.SYSTEM_HASH_SHEET: _GridSheet()}))

    wb = _valid_integrity_workbook()
    eng._validate_source_workbook_integrity(wb)

    wb_empty = _valid_integrity_workbook()
    cast_hash = wb_empty[eng.SYSTEM_HASH_SHEET]
    cast_hash.set_ref("A2", "")
    cast_hash.set_ref("B2", "")
    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(wb_empty)

    monkeypatch.setattr(eng, "verify_payload_signature", lambda *_a, **_k: False)
    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(_valid_integrity_workbook())

    monkeypatch.setattr(eng, "verify_payload_signature", lambda *_a, **_k: True)
    wb_unknown = _valid_integrity_workbook()
    wb_unknown[eng.SYSTEM_HASH_SHEET].set_ref("A2", "UNKNOWN")
    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(wb_unknown)

    wb_a1_bad = _valid_integrity_workbook()
    wb_a1_bad[eng.SYSTEM_LAYOUT_SHEET].set_ref("A1", "BAD")
    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(wb_a1_bad)

    wb_b1_bad = _valid_integrity_workbook()
    wb_b1_bad[eng.SYSTEM_LAYOUT_SHEET].set_ref("B1", "BAD")
    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(wb_b1_bad)

    wb_manifest_missing = _valid_integrity_workbook()
    wb_manifest_missing[eng.SYSTEM_LAYOUT_SHEET].set_ref("A2", "")
    wb_manifest_missing[eng.SYSTEM_LAYOUT_SHEET].set_ref("B2", "")
    with pytest.raises(ValidationError):
        eng._validate_source_workbook_integrity(wb_manifest_missing)


def test_course_metadata_and_component_readers_branches() -> None:
    metadata = _GridSheet()
    metadata.set_cell(2, 1, "Course_Code")
    metadata.set_cell(2, 2, "C101")
    metadata.set_cell(3, 1, "Total_Outcomes")
    metadata.set_cell(3, 2, "bad")
    with pytest.raises(ValidationError):
        eng._read_course_metadata(metadata)

    metadata.set_cell(3, 2, "3")
    metadata.set_cell(4, 1, "")
    metadata.set_cell(4, 2, "")
    rows, total = eng._read_course_metadata(metadata)
    assert total == 3
    assert rows

    bad_header = _GridSheet()
    bad_header.set_cell(1, 1, "wrong")
    with pytest.raises(ValidationError):
        eng._read_direct_components(bad_header)

    direct = _GridSheet()
    for idx, h in enumerate(eng.ASSESSMENT_CONFIG_HEADERS, start=1):
        direct.set_cell(1, idx, h)
    direct.set_cell(2, 1, "S1")
    direct.set_cell(2, 2, True)
    direct.set_cell(2, 5, eng.ASSESSMENT_VALIDATION_YES_NO_OPTIONS[0])
    with pytest.raises(ValidationError):
        eng._read_direct_components(direct)

    indirect = _GridSheet()
    for idx, h in enumerate(eng.ASSESSMENT_CONFIG_HEADERS, start=1):
        indirect.set_cell(1, idx, h)
    indirect.set_cell(2, 1, "Survey")
    indirect.set_cell(2, 2, True)
    indirect.set_cell(2, 5, eng.ASSESSMENT_VALIDATION_YES_NO_OPTIONS[1])
    with pytest.raises(ValidationError):
        eng._read_indirect_components(indirect)


def test_layout_and_component_name_parsing_branches() -> None:
    assert eng._parse_co_values("") == []
    assert eng._parse_co_values("co1, co1, co2") == [1, 2]

    sheet = _GridSheet()
    sheet.set_ref("A2", 123)
    with pytest.raises(ValidationError):
        eng._read_layout_manifest(sheet)
    sheet.set_ref("A2", "{bad")
    with pytest.raises(ValidationError):
        eng._read_layout_manifest(sheet)
    sheet.set_ref("A2", "[]")
    with pytest.raises(ValidationError):
        eng._read_layout_manifest(sheet)

    manifest = {
        eng.LAYOUT_MANIFEST_KEY_SHEETS: [
            "not-a-dict",
            {eng.LAYOUT_SHEET_SPEC_KEY_KIND: "wrong"},
            {eng.LAYOUT_SHEET_SPEC_KEY_KIND: eng.LAYOUT_SHEET_KIND_DIRECT_CO_WISE, eng.LAYOUT_SHEET_SPEC_KEY_ANCHORS: "x"},
            {
                eng.LAYOUT_SHEET_SPEC_KEY_KIND: eng.LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
                eng.LAYOUT_SHEET_SPEC_KEY_ANCHORS: [
                    ["bad", "shape", "extra"],  # invalid shape -> continue
                    ["B0", eng.COMPONENT_NAME_LABEL],  # invalid row
                    [10, "bad"],  # invalid ref
                    ["C2", "S1"],
                    ["B2", eng.COMPONENT_NAME_LABEL],
                ],
            },
        ]
    }
    out = eng._component_sheet_specs_by_kind(
        manifest,
        allowed_kinds={eng.LAYOUT_SHEET_KIND_DIRECT_CO_WISE},
    )
    assert list(out.keys()) == [eng.normalize("S1")]
    assert eng._component_name_from_spec({eng.LAYOUT_SHEET_SPEC_KEY_ANCHORS: "bad"}) == ""
    assert eng._component_name_from_spec(
        {eng.LAYOUT_SHEET_SPEC_KEY_ANCHORS: [["C2", "S1"]]}
    ) == ""
    assert eng._cell_row("AB12") == 12


def test_students_and_component_mark_computation_branches() -> None:
    wb = _Workbook({})
    with pytest.raises(ValidationError):
        eng._read_students_from_component_sheets(wb, direct_specs={}, indirect_specs={})
    with pytest.raises(ValidationError):
        eng._read_students_from_component_sheets(
            _Workbook({"X": _GridSheet()}),
            direct_specs={"x": {eng.LAYOUT_SHEET_SPEC_KEY_NAME: "X", eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: "bad"}},
            indirect_specs={},
        )

    ws = _GridSheet()
    ws.set_cell(4, eng._EXCEL_COL_REG_NO, "R1")
    ws.set_cell(4, eng._EXCEL_COL_STUDENT_NAME, "Alice")
    ws.set_cell(5, eng._EXCEL_COL_REG_NO, "")
    ws.set_cell(5, eng._EXCEL_COL_STUDENT_NAME, "")
    students = eng._read_students_from_component_sheets(
        _Workbook({"D1": ws}),
        direct_specs={"x": {eng.LAYOUT_SHEET_SPEC_KEY_NAME: "D1", eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1}},
        indirect_specs={},
    )
    assert students == [("R1", "Alice")]

    comp = eng._DirectComponent(name="S1", weight=10.0)
    with pytest.raises(ValidationError):
        eng._compute_component_marks(workbook=_Workbook({}), component=comp, students=students, spec=None, total_outcomes=2)

    with pytest.raises(ValidationError):
        eng._compute_component_marks(
            workbook=_Workbook({"S1": ws}),
            component=comp,
            students=students,
            spec={eng.LAYOUT_SHEET_SPEC_KEY_NAME: "S1", eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1, eng.LAYOUT_SHEET_SPEC_KEY_HEADERS: [], eng.LAYOUT_SHEET_SPEC_KEY_KIND: "bad"},
            total_outcomes=2,
        )

    co_wise = _GridSheet()
    # header_row = 1 -> CO row is 2, MAX row is 3, student rows start at 4
    co_wise.set_cell(2, 4, "CO1,CO2")  # len!=1 -> skipped
    with pytest.raises(ValidationError):
        eng._compute_component_marks(
            workbook=_Workbook({"S1": co_wise}),
            component=comp,
            students=students,
            spec={
                eng.LAYOUT_SHEET_SPEC_KEY_NAME: "S1",
                eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                eng.LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "Q1", "Total"],
                eng.LAYOUT_SHEET_SPEC_KEY_KIND: eng.LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            },
            total_outcomes=2,
        )

    co_wise_non_numeric = _GridSheet()
    co_wise_non_numeric.set_cell(2, 4, "CO1")
    co_wise_non_numeric.set_cell(3, 4, 10)
    co_wise_non_numeric.set_cell(4, 4, "not-numeric")
    computed_co = eng._compute_component_marks(
        workbook=_Workbook({"S1N": co_wise_non_numeric}),
        component=comp,
        students=students,
        spec={
            eng.LAYOUT_SHEET_SPEC_KEY_NAME: "S1N",
            eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
            eng.LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "Q1", "Total"],
            eng.LAYOUT_SHEET_SPEC_KEY_KIND: eng.LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
        },
        total_outcomes=2,
    )
    assert computed_co.marks_by_co[1][0] == 0.0

    non_co = _GridSheet()
    # columns checked: 5..len(headers)
    non_co.set_cell(2, 5, "CO1,CO2")  # len!=1 -> skipped
    non_co.set_cell(2, 6, "CO9")  # out-of-range -> skipped
    non_co.set_cell(2, 7, "CO1")  # valid
    non_co.set_cell(3, 7, 10)  # max marks row
    non_co.set_cell(4, 4, eng.CO_REPORT_ABSENT_TOKEN)  # absent student
    non_co.set_cell(5, 4, "not-numeric")  # triggers total_numeric=None -> 0.0
    students2 = [("R1", "A"), ("R2", "B")]
    computed = eng._compute_component_marks(
        workbook=_Workbook({"S2": non_co}),
        component=comp,
        students=students2,
        spec={
            eng.LAYOUT_SHEET_SPEC_KEY_NAME: "S2",
            eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
            eng.LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "Total", "C1", "C2", "C3"],
            eng.LAYOUT_SHEET_SPEC_KEY_KIND: eng.LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
        },
        total_outcomes=2,
    )
    assert computed.marks_by_co[1][0] == eng.CO_REPORT_ABSENT_TOKEN

    ind_comp = eng._IndirectComponent(name="Survey", weight=100.0)
    with pytest.raises(ValidationError):
        eng._compute_indirect_component_marks(
            workbook=_Workbook({}),
            component=ind_comp,
            students=students2,
            spec=None,
            total_outcomes=2,
        )
    with pytest.raises(ValidationError):
        eng._compute_indirect_component_marks(
            workbook=_Workbook({"I1": non_co}),
            component=ind_comp,
            students=students2,
            spec={
                eng.LAYOUT_SHEET_SPEC_KEY_NAME: "I1",
                eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                eng.LAYOUT_SHEET_SPEC_KEY_HEADERS: [],
                eng.LAYOUT_SHEET_SPEC_KEY_KIND: "bad",
            },
            total_outcomes=2,
        )
    indirect = _GridSheet()
    indirect.set_cell(2, 4, eng.CO_REPORT_ABSENT_TOKEN)
    indirect.set_cell(2, 5, 3.2)
    indirect.set_cell(3, 4, 4.0)
    indirect.set_cell(3, 5, 5.0)
    ind = eng._compute_indirect_component_marks(
        workbook=_Workbook({"I2": indirect}),
        component=ind_comp,
        students=students2,
        spec={
            eng.LAYOUT_SHEET_SPEC_KEY_NAME: "I2",
            eng.LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
            eng.LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "CO1", "CO2"],
            eng.LAYOUT_SHEET_SPEC_KEY_KIND: eng.LAYOUT_SHEET_KIND_INDIRECT,
        },
        total_outcomes=2,
    )
    assert ind.marks_by_co[1][0] == eng.CO_REPORT_ABSENT_TOKEN

    assert eng._split_equal_with_residual(10.0, 0) == []
    assert eng._split_equal_with_residual(10.0, 1) == [10.0]
    assert eng._to_float(True) is None
    assert eng._to_float("x") is None


class _FakeWS:
    def __init__(self) -> None:
        self.writes: list[tuple[int, int, Any, Any]] = []
        self.freeze_calls: list[tuple[int, int]] = []

    def write(self, row: int, col: int, value: Any, fmt: Any = None) -> None:
        self.writes.append((row, col, value, fmt))

    def repeat_rows(self, *_args: Any, **_kwargs: Any) -> None:
        return None

    def freeze_panes(self, row: int, col: int) -> None:
        self.freeze_calls.append((row, col))


class _FakeWB:
    def __init__(self) -> None:
        self.ws = _FakeWS()

    def add_worksheet(self, _name: str) -> _FakeWS:
        return self.ws


def test_xlsxwriter_write_indirect_sheet_branch_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(eng, "_xlsxwriter_write_report_metadata", lambda *_a, **_k: 3)
    monkeypatch.setattr(eng, "_xlsxwriter_set_report_metadata_column_widths", lambda *_a, **_k: None)
    monkeypatch.setattr(eng, "_xlsxwriter_apply_layout", lambda *_a, **_k: None)

    wb = _FakeWB()
    formats = {"header": object(), "body": object(), "body_wrap": object(), "body_center": object()}
    comps = [
        eng._IndirectComponentComputed(name="A", weight=40.0, marks_by_co={1: [eng.CO_REPORT_ABSENT_TOKEN, 3.0]}),
        eng._IndirectComponentComputed(name="B", weight=60.0, marks_by_co={1: [2.0, 4.0]}),
    ]
    eng._xlsxwriter_write_indirect_sheet(
        wb,
        co_index=1,
        metadata_rows=[("Course_Code", "C101")],
        students=[("R1", "Alice"), ("R2", "Bob")],
        components=comps,
        formats=formats,
    )
    written_values = [value for (_r, _c, value, _f) in wb.ws.writes]
    assert any(isinstance(v, str) and "(40%)" in v for v in written_values)
    assert eng.CO_REPORT_NOT_APPLICABLE_TOKEN in written_values


def test_normalize_page_setup_fit_and_ratio_header_branches(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    real_import = builtins.__import__

    def _fail_openpyxl(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[no-untyped-def]
        if name == "openpyxl":
            raise Exception("missing")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _fail_openpyxl)
    eng._normalize_page_setup_fit(tmp_path / "no-file-needed.xlsx")
    monkeypatch.setattr(builtins, "__import__", real_import)

    assert eng._ratio_total_header(0.333) == eng.CO_REPORT_HEADER_TOTAL_RATIO_TEMPLATE.format(ratio="33.3")


def test_generate_final_report_additional_exception_and_cleanup_branches(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = tmp_path / "filled.xlsx"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "final.xlsx"

    class _WB:
        class _Sheet:
            def __getitem__(self, _key: str) -> _Cell:
                return _Cell("")

        def __getitem__(self, _key: str) -> "_WB._Sheet":
            return _WB._Sheet()

        def close(self) -> None:
            return None

    class _OpenPyxl:
        @staticmethod
        def load_workbook(*_a: Any, **_k: Any) -> _WB:
            return _WB()

    class _Xlsxwriter:
        pass

    import sys

    monkeypatch.setitem(sys.modules, "openpyxl", _OpenPyxl())
    monkeypatch.setitem(sys.modules, "xlsxwriter", _Xlsxwriter())
    monkeypatch.setattr(eng, "_validate_source_workbook_integrity", lambda _wb: None)
    monkeypatch.setattr(eng, "_read_course_metadata", lambda _sheet: ([("Course_Code", "C101")], 1))
    monkeypatch.setattr(eng, "_read_layout_manifest", lambda _sheet: {})
    monkeypatch.setattr(eng, "_component_sheet_specs_by_kind", lambda *_a, **_k: {"s1": {}})
    monkeypatch.setattr(eng, "_read_students_from_component_sheets", lambda *_a, **_k: [("R1", "A")])
    monkeypatch.setattr(
        eng,
        "_compute_component_marks",
        lambda **_k: eng._DirectComponentComputed(
            name="S1",
            weight=100.0,
            max_by_co={1: 1.0},
            marks_by_co={1: [1.0]},
        ),
    )
    monkeypatch.setattr(eng, "_normalize_page_setup_fit", lambda _p: None)
    monkeypatch.setattr(eng, "_raise_if_cancelled", lambda _t: None)

    # Second-try ValidationError branch.
    monkeypatch.setattr(eng, "_read_direct_components", lambda _s: (_ for _ in ()).throw(ValidationError("bad")))
    monkeypatch.setattr(eng, "_read_indirect_components", lambda _s: [])
    with pytest.raises(ValidationError):
        eng.generate_final_co_report(source, output)

    # Second-try JobCancelledError branch.
    monkeypatch.setattr(eng, "_read_direct_components", lambda _s: (_ for _ in ()).throw(JobCancelledError("cancel")))
    with pytest.raises(JobCancelledError):
        eng.generate_final_co_report(source, output)

    # no-direct/no-indirect branch.
    monkeypatch.setattr(eng, "_read_direct_components", lambda _s: [])
    monkeypatch.setattr(eng, "_read_indirect_components", lambda _s: [])
    with pytest.raises(ValidationError):
        eng.generate_final_co_report(source, output)

    # AppSystemError branch + cleanup warning path.
    monkeypatch.setattr(eng, "_read_direct_components", lambda _s: [eng._DirectComponent(name="S1", weight=100.0)])
    monkeypatch.setattr(eng, "_write_final_report_workbook_xlsxwriter", lambda **_k: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr(eng.Path, "unlink", lambda self: (_ for _ in ()).throw(OSError("locked")))
    warnings: list[str] = []
    monkeypatch.setattr(eng._logger, "warning", lambda msg, *_a, **_k: warnings.append(str(msg)))
    with pytest.raises(AppSystemError):
        eng.generate_final_co_report(source, output)
    assert warnings


def test_generate_final_report_initial_job_cancel_branch(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    source = tmp_path / "filled.xlsx"
    source.write_text("x", encoding="utf-8")
    output = tmp_path / "final.xlsx"

    class _OpenPyxl:
        @staticmethod
        def load_workbook(*_a: Any, **_k: Any):  # type: ignore[no-untyped-def]
            return object()

    class _Xlsxwriter:
        pass

    import sys

    monkeypatch.setitem(sys.modules, "openpyxl", _OpenPyxl())
    monkeypatch.setitem(sys.modules, "xlsxwriter", _Xlsxwriter())
    monkeypatch.setattr(eng, "_raise_if_cancelled", lambda _t: (_ for _ in ()).throw(JobCancelledError("cancel")))
    with pytest.raises(JobCancelledError):
        eng.generate_final_co_report(source, output)
