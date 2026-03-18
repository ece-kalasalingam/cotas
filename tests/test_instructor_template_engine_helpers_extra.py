from __future__ import annotations

import builtins
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest

from common.constants import (
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SHEET,
    ID_COURSE_SETUP,
    SYSTEM_HASH_SHEET,
    SYSTEM_HASH_TEMPLATE_HASH_KEY,
    SYSTEM_HASH_TEMPLATE_ID_KEY,
)
from common.exceptions import JobCancelledError, ValidationError
from common.sheet_schema import SheetSchema, WorkbookBlueprint
from common.workbook_signing import sign_payload
from domain import instructor_template_engine as eng


@dataclass
class _Cell:
    value: Any


class _Sheet:
    def __init__(self, title: str) -> None:
        self.title = title
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

    def iter_rows(self, *, min_row: int, max_col: int, values_only: bool = True):
        del values_only
        row = min_row
        while True:
            values = tuple(self._cells.get((row, col)) for col in range(1, max_col + 1))
            if all(v is None for v in values):
                break
            yield values
            row += 1


class _Workbook:
    def __init__(self, sheets: dict[str, _Sheet]) -> None:
        self._sheets = sheets
        self.sheetnames = list(sheets.keys())
        self.closed = False

    def __getitem__(self, key: str) -> _Sheet:
        return self._sheets[key]

    def close(self) -> None:
        self.closed = True


def test_get_blueprint_and_template_dispatch_missing_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    with pytest.raises(ValidationError):
        eng._get_blueprint("UNKNOWN")

    old = eng.BLUEPRINT_REGISTRY.get(ID_COURSE_SETUP)
    assert old is not None
    monkeypatch.delitem(eng.BLUEPRINT_REGISTRY, ID_COURSE_SETUP, raising=False)
    try:
        with pytest.raises(ValidationError):
            eng._get_blueprint(ID_COURSE_SETUP)
    finally:
        eng.BLUEPRINT_REGISTRY[ID_COURSE_SETUP] = old

    with pytest.raises(ValidationError):
        eng._extract_marks_template_context_by_template(object(), "UNKNOWN")
    with pytest.raises(ValidationError):
        eng._write_marks_template_workbook_by_template(object(), {}, template_id="UNKNOWN")
    with pytest.raises(ValidationError):
        eng._validate_template_specific_rules(object(), "UNKNOWN")


def test_extract_total_outcomes_students_components_questions_helpers() -> None:
    with pytest.raises(ValidationError):
        eng._extract_total_outcomes([["Total_Outcomes", True]])
    assert eng._extract_total_outcomes([["x"], ["Total_Outcomes", 4]]) == 4

    with pytest.raises(ValidationError):
        eng._extract_students([["", ""]])
    assert eng._extract_students([["R1", "Alice"]]) == [("R1", "Alice")]

    with pytest.raises(ValidationError):
        eng._extract_components([])
    comps = eng._extract_components(
        [
            ["S1", 30, "yes", "yes", "yes"],
            ["S1", 30, "yes", "yes", "yes"],  # duplicate ignored
            ["", 10, "yes", "yes", "yes"],  # ignored
        ]
    )
    assert len(comps) == 1

    q = eng._extract_questions(
        [
            ["", "Q1", 2, "CO1"],  # skipped
            ["S1", "Q1", True, "CO1"],  # skipped bool
            ["S1", "Q1", 2, ""],  # skipped empty co
            ["S1", "Q1", 2, "CO1"],
        ]
    )
    assert len(q["s1"]) == 1

    assert eng._co_tokens(None) == []
    assert eng._co_tokens(True) == []
    assert eng._co_tokens(1.5) == []
    assert eng._co_tokens("CO1, ,CO2") == []
    assert eng._co_tokens("CO1,CO2") == [1, 2]


def test_precompute_layout_manifest_missing_direct_questions_raises() -> None:
    context = {
        "students": [("R1", "A")],
        "metadata_rows": [("Course_Code", "C101"), ("Total_Outcomes", 1)],
        "assessment_rows": [],
        "questions_by_component": {},
        "total_outcomes": 1,
        "components": [
            {"display_name": "S1", "key": "s1", "is_direct": True, "co_wise_breakup": True},
        ],
    }
    with pytest.raises(ValidationError):
        eng._precompute_marks_layout_manifest(context=context)


def test_validate_course_details_workbook_openpyxl_missing_and_file_missing(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    p = tmp_path / "missing.xlsx"
    with pytest.raises(ValidationError):
        eng.validate_course_details_workbook(p)

    real_import = builtins.__import__

    def _fail_openpyxl(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[no-untyped-def]
        if name == "openpyxl":
            raise ModuleNotFoundError("openpyxl")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _fail_openpyxl)
    with pytest.raises(ValidationError):
        eng.validate_course_details_workbook(p)


def test_extract_and_validate_template_id_and_schema_branches() -> None:
    with pytest.raises(ValidationError):
        eng._extract_and_validate_template_id(_Workbook({}))

    hash_sheet = _Sheet(SYSTEM_HASH_SHEET)
    hash_sheet.set_ref("A1", "bad")
    hash_sheet.set_ref("B1", SYSTEM_HASH_TEMPLATE_HASH_KEY)
    wb = _Workbook({SYSTEM_HASH_SHEET: hash_sheet})
    with pytest.raises(ValidationError):
        eng._extract_and_validate_template_id(wb)

    hash_sheet = _Sheet(SYSTEM_HASH_SHEET)
    hash_sheet.set_ref("A1", SYSTEM_HASH_TEMPLATE_ID_KEY)
    hash_sheet.set_ref("B1", "bad")
    wb = _Workbook({SYSTEM_HASH_SHEET: hash_sheet})
    with pytest.raises(ValidationError):
        eng._extract_and_validate_template_id(wb)

    hash_sheet = _Sheet(SYSTEM_HASH_SHEET)
    hash_sheet.set_ref("A1", SYSTEM_HASH_TEMPLATE_ID_KEY)
    hash_sheet.set_ref("B1", SYSTEM_HASH_TEMPLATE_HASH_KEY)
    hash_sheet.set_ref("A2", "")
    hash_sheet.set_ref("B2", "")
    wb = _Workbook({SYSTEM_HASH_SHEET: hash_sheet})
    with pytest.raises(ValidationError):
        eng._extract_and_validate_template_id(wb)

    hash_sheet = _Sheet(SYSTEM_HASH_SHEET)
    hash_sheet.set_ref("A1", SYSTEM_HASH_TEMPLATE_ID_KEY)
    hash_sheet.set_ref("B1", SYSTEM_HASH_TEMPLATE_HASH_KEY)
    hash_sheet.set_ref("A2", ID_COURSE_SETUP)
    hash_sheet.set_ref("B2", "bad")
    wb = _Workbook({SYSTEM_HASH_SHEET: hash_sheet})
    with pytest.raises(ValidationError):
        eng._extract_and_validate_template_id(wb)

    hash_sheet.set_ref("B2", sign_payload(ID_COURSE_SETUP))
    assert eng._extract_and_validate_template_id(_Workbook({SYSTEM_HASH_SHEET: hash_sheet})) == ID_COURSE_SETUP

    blueprint = WorkbookBlueprint(
        type_id=ID_COURSE_SETUP,
        style_registry={},
        sheets=[SheetSchema(name=COURSE_METADATA_SHEET, header_matrix=[list(COURSE_METADATA_HEADERS)])],
    )
    ws = _Sheet(COURSE_METADATA_SHEET)
    ws.set_cell(1, 1, "Field")
    ws.set_cell(1, 2, "Wrong")
    wb_schema = _Workbook({COURSE_METADATA_SHEET: ws, SYSTEM_HASH_SHEET: _Sheet(SYSTEM_HASH_SHEET)})
    with pytest.raises(ValidationError):
        eng._validate_workbook_schema(wb_schema, blueprint)


def test_parse_yes_no_invalid_branch() -> None:
    with pytest.raises(ValidationError):
        eng._parse_yes_no("maybe", "Assessment_Config", 2, "Is_Direct")


class _FakeWS:
    def write_row(self, *_a: Any, **_k: Any) -> None:
        return None

    def set_column(self, *_a: Any, **_k: Any) -> None:
        return None

    def freeze_panes(self, *_a: Any, **_k: Any) -> None:
        return None

    def hide(self) -> None:
        return None

    def data_validation(self, *_a: Any, **_k: Any) -> None:
        return None


class _FakeWorkbookOK:
    def __init__(self, path: str, _opts: dict[str, Any] | None = None) -> None:
        self.path = Path(path)

    def add_format(self, fmt: dict[str, Any]) -> dict[str, Any]:
        return fmt

    def add_worksheet(self, _name: str) -> _FakeWS:
        return _FakeWS()

    def close(self) -> None:
        self.path.write_text("tmp", encoding="utf-8")


class _FakeWorkbookCloseFails(_FakeWorkbookOK):
    def close(self) -> None:
        self.path.write_text("tmp", encoding="utf-8")
        raise RuntimeError("close failed")


def _install_fake_xlsxwriter(monkeypatch: pytest.MonkeyPatch, workbook_cls: type[object]) -> None:
    class _M:
        Workbook = workbook_cls

    monkeypatch.setitem(sys.modules, "xlsxwriter", _M())


def test_generate_course_template_validation_cleanup_and_protect_branches(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    old = eng.BLUEPRINT_REGISTRY[ID_COURSE_SETUP]
    bad_blueprint = WorkbookBlueprint(
        type_id=ID_COURSE_SETUP,
        style_registry={},
        sheets=[SheetSchema(name="X", header_matrix=[["H1"], ["H2"]], is_protected=False)],
    )
    eng.BLUEPRINT_REGISTRY[ID_COURSE_SETUP] = bad_blueprint
    _install_fake_xlsxwriter(monkeypatch, _FakeWorkbookCloseFails)
    monkeypatch.setattr(eng.Path, "unlink", lambda self: (_ for _ in ()).throw(OSError("locked")))
    warnings: list[str] = []
    monkeypatch.setattr(eng._logger, "warning", lambda msg, *_a, **_k: warnings.append(str(msg)))
    try:
        with pytest.raises(ValidationError):
            eng.generate_course_details_template(tmp_path / "out.xlsx")
    finally:
        eng.BLUEPRINT_REGISTRY[ID_COURSE_SETUP] = old
    assert warnings

    # Hit the protected-sheet branch.
    protected_blueprint = WorkbookBlueprint(
        type_id=ID_COURSE_SETUP,
        style_registry={},
        sheets=[SheetSchema(name="Y", header_matrix=[["H1"]], is_protected=True)],
    )
    eng.BLUEPRINT_REGISTRY[ID_COURSE_SETUP] = protected_blueprint
    _install_fake_xlsxwriter(monkeypatch, _FakeWorkbookOK)
    called = {"protect": 0}
    monkeypatch.setattr(eng, "_protect_sheet", lambda _ws: called.__setitem__("protect", called["protect"] + 1))
    try:
        eng.generate_course_details_template(tmp_path / "out2.xlsx")
    finally:
        eng.BLUEPRINT_REGISTRY[ID_COURSE_SETUP] = old
    assert called["protect"] == 1


def test_generate_marks_template_import_and_processing_exception_branches(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = tmp_path / "source.xlsx"
    source.write_text("x", encoding="utf-8")
    out = tmp_path / "marks.xlsx"
    real_import = builtins.__import__

    def _fail_openpyxl(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[no-untyped-def]
        if name == "openpyxl":
            raise ModuleNotFoundError("openpyxl")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _fail_openpyxl)
    with pytest.raises(ValidationError):
        eng.generate_marks_template_from_course_details(source, out)

    def _fail_xlsxwriter(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[no-untyped-def]
        if name == "xlsxwriter":
            raise ModuleNotFoundError("xlsxwriter")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _fail_xlsxwriter)
    with pytest.raises(ValidationError):
        eng.generate_marks_template_from_course_details(source, out)
    monkeypatch.setattr(builtins, "__import__", real_import)

    with pytest.raises(ValidationError):
        eng.generate_marks_template_from_course_details(tmp_path / "missing.xlsx", out)

    class _OpenPyxl:
        @staticmethod
        def load_workbook(*_a: Any, **_k: Any):
            raise RuntimeError("bad workbook")

    class _Xlsxwriter:
        Workbook = _FakeWorkbookOK

    monkeypatch.setitem(sys.modules, "openpyxl", _OpenPyxl())
    monkeypatch.setitem(sys.modules, "xlsxwriter", _Xlsxwriter())
    with pytest.raises(ValidationError):
        eng.generate_marks_template_from_course_details(source, out)

    # Main processing JobCancelled and generic exception branches with cleanup warnings.
    class _WB:
        sheetnames = [SYSTEM_HASH_SHEET]

        def close(self) -> None:
            return None

    class _OpenPyxl2:
        @staticmethod
        def load_workbook(*_a: Any, **_k: Any):
            return _WB()

    class _Xlsxwriter2:
        Workbook = _FakeWorkbookCloseFails

    monkeypatch.setitem(sys.modules, "openpyxl", _OpenPyxl2())
    monkeypatch.setitem(sys.modules, "xlsxwriter", _Xlsxwriter2())
    monkeypatch.setattr(eng.Path, "unlink", lambda self: (_ for _ in ()).throw(OSError("locked")))
    monkeypatch.setattr(eng._logger, "debug", lambda *_a, **_k: None)
    monkeypatch.setattr(eng._logger, "warning", lambda *_a, **_k: None)
    monkeypatch.setattr(eng, "_extract_and_validate_template_id", lambda _wb: (_ for _ in ()).throw(JobCancelledError()))
    with pytest.raises(JobCancelledError):
        eng.generate_marks_template_from_course_details(source, out)

    monkeypatch.setattr(eng, "_extract_and_validate_template_id", lambda _wb: (_ for _ in ()).throw(RuntimeError("boom")))
    with pytest.raises(Exception):
        eng.generate_marks_template_from_course_details(source, out)


def test_extract_helpers_additional_continue_and_schema_branches() -> None:
    with pytest.raises(ValidationError):
        eng._extract_components([["too", "short"]])
    assert eng._extract_questions([["too", "short", "only"]]) == {}

    bp = WorkbookBlueprint(
        type_id=ID_COURSE_SETUP,
        style_registry={},
        sheets=[SheetSchema(name=COURSE_METADATA_SHEET, header_matrix=[list(COURSE_METADATA_HEADERS)])],
    )
    ws = _Sheet(COURSE_METADATA_SHEET)
    ws.set_cell(1, 1, "Field")
    ws.set_cell(1, 2, "Value")
    wb = _Workbook({COURSE_METADATA_SHEET: ws, SYSTEM_HASH_SHEET: _Sheet(SYSTEM_HASH_SHEET), "EXTRA": _Sheet("EXTRA")})
    with pytest.raises(ValidationError):
        eng._validate_workbook_schema(wb, bp)

    bp_bad = WorkbookBlueprint(
        type_id=ID_COURSE_SETUP,
        style_registry={},
        sheets=[SheetSchema(name=COURSE_METADATA_SHEET, header_matrix=[["A"], ["B"]])],
    )
    with pytest.raises(ValidationError):
        eng._validate_workbook_schema(_Workbook({COURSE_METADATA_SHEET: ws, SYSTEM_HASH_SHEET: _Sheet(SYSTEM_HASH_SHEET)}), bp_bad)

    assert eng._co_tokens("bad-token") == []
