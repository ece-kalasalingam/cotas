from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pytest

from common.constants import (
    ASSESSMENT_CONFIG_HEADERS,
    ASSESSMENT_CONFIG_SHEET,
    COURSE_METADATA_HEADERS,
    COURSE_METADATA_SHEET,
    LAYOUT_MANIFEST_KEY_SHEET_ORDER,
    LAYOUT_MANIFEST_KEY_SHEETS,
    LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
    LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
    LAYOUT_SHEET_KIND_INDIRECT,
    LAYOUT_SHEET_SPEC_KEY_ANCHORS,
    LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS,
    LAYOUT_SHEET_SPEC_KEY_HEADER_ROW,
    LAYOUT_SHEET_SPEC_KEY_HEADERS,
    LAYOUT_SHEET_SPEC_KEY_KIND,
    LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE,
    LAYOUT_SHEET_SPEC_KEY_NAME,
    LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT,
    LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH,
    QUESTION_MAP_HEADERS,
    QUESTION_MAP_SHEET,
    STUDENTS_HEADERS,
    STUDENTS_SHEET,
)
from common.exceptions import ValidationError
from common.sample_setup_data import SAMPLE_SETUP_DATA
from domain.template_versions import course_setup_v1 as v1


def _col_name(col: int) -> str:
    out = ""
    x = col
    while x > 0:
        x, r = divmod(x - 1, 26)
        out = chr(65 + r) + out
    return out


@dataclass
class _Cell:
    row: int
    col: int
    value: Any

    @property
    def coordinate(self) -> str:
        return f"{_col_name(self.col)}{self.row}"


class _Sheet:
    def __init__(self, title: str) -> None:
        self.title = title
        self._cells: dict[tuple[int, int], Any] = {}
        self._ref: dict[str, Any] = {}

    def set_cell(self, row: int, col: int, value: Any) -> None:
        self._cells[(row, col)] = value

    def set_ref(self, ref: str, value: Any) -> None:
        self._ref[ref] = value

    def cell(self, *, row: int, column: int) -> _Cell:
        return _Cell(row=row, col=column, value=self._cells.get((row, column)))

    def __getitem__(self, ref: str) -> _Cell:
        col = 0
        row = 0
        for ch in ref:
            if ch.isalpha():
                col = col * 26 + (ord(ch.upper()) - 64)
            elif ch.isdigit():
                row = row * 10 + int(ch)
        return _Cell(row=row or 1, col=col or 1, value=self._ref.get(ref))

    def iter_rows(self, *, min_row: int, max_col: int, values_only: bool = True):
        del values_only
        row = min_row
        while True:
            values = tuple(self._cells.get((row, c)) for c in range(1, max_col + 1))
            if all(v is None for v in values):
                break
            yield values
            row += 1


class _WB:
    def __init__(self, sheets: dict[str, _Sheet]) -> None:
        self._sheets = sheets
        self.sheetnames = list(sheets.keys())

    def __getitem__(self, key: str) -> _Sheet:
        return self._sheets[key]


def _base_headers(sheet: _Sheet, headers: list[str]) -> None:
    for i, h in enumerate(headers, start=1):
        sheet.set_cell(1, i, h)


def test_filled_manifest_schema_invalid_top_level_and_sheet_specs() -> None:
    wb = _WB({})
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, [])
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, {})

    m = {LAYOUT_MANIFEST_KEY_SHEET_ORDER: ["A"], LAYOUT_MANIFEST_KEY_SHEETS: []}
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, m)


def test_core_helper_branches_for_metadata_assessment_question_students() -> None:
    # metadata
    meta = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(meta, list(COURSE_METADATA_HEADERS))
    meta.set_cell(2, 1, "")
    meta.set_cell(2, 2, "x")
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta)

    meta = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(meta, list(COURSE_METADATA_HEADERS))
    meta.set_cell(2, 1, "Course_Code")
    meta.set_cell(2, 2, "")
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta)

    # yes/no
    with pytest.raises(ValidationError):
        v1._parse_yes_no("maybe", ASSESSMENT_CONFIG_SHEET, 2, ASSESSMENT_CONFIG_HEADERS[4])

    # assessment
    ass = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ass, list(ASSESSMENT_CONFIG_HEADERS))
    with pytest.raises(ValidationError):
        v1._validate_assessment_config(ass)

    ass.set_cell(2, 1, "S1")
    ass.set_cell(2, 2, "bad")
    ass.set_cell(2, 3, "yes")
    ass.set_cell(2, 4, "yes")
    ass.set_cell(2, 5, "yes")
    with pytest.raises(ValidationError):
        v1._validate_assessment_config(ass)

    # question map
    qm = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm, list(QUESTION_MAP_HEADERS))
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm, {"s1": {"co_wise_breakup": True}}, 2)

    qm.set_cell(2, 1, "")
    qm.set_cell(2, 2, "Q1")
    qm.set_cell(2, 3, 2)
    qm.set_cell(2, 4, "CO1")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm, {"s1": {"co_wise_breakup": True}}, 2)

    # students
    st = _Sheet(STUDENTS_SHEET)
    _base_headers(st, list(STUDENTS_HEADERS))
    with pytest.raises(ValidationError):
        v1._validate_students(st)
    st.set_cell(2, 1, "R1")
    st.set_cell(2, 2, "")
    with pytest.raises(ValidationError):
        v1._validate_students(st)


def test_tokens_and_marks_helpers_branches() -> None:
    assert v1._co_tokens(None) == []
    assert v1._co_tokens(True) == []
    assert v1._co_tokens(1.2) == []
    assert v1._co_tokens("CO1,") == []
    assert v1._co_tokens("CO1,CO2") == [1, 2]

    assert v1._marks_data_start_row(LAYOUT_SHEET_KIND_INDIRECT, 10) == 11
    assert v1._marks_data_start_row(LAYOUT_SHEET_KIND_DIRECT_CO_WISE, 10) == 13
    assert list(v1._marks_entry_columns(LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE, 8)) == [4]
    assert list(v1._marks_entry_columns(LAYOUT_SHEET_KIND_DIRECT_CO_WISE, 8)) == [4, 5, 6, 7]
    assert list(v1._marks_entry_columns(LAYOUT_SHEET_KIND_INDIRECT, 8)) == [4, 5, 6, 7, 8]
    with pytest.raises(ValidationError):
        v1._marks_entry_columns("BAD_KIND", 8)

    ws = _Sheet("S1")
    ws.set_cell(3, 4, "bad")
    with pytest.raises(ValidationError):
        v1._mark_max_for_cell(ws, "BAD_KIND", 3, 4)
    with pytest.raises(ValidationError):
        v1._mark_max_for_cell(ws, LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE, 3, 4)


def test_student_identity_and_structure_snapshot_branches() -> None:
    ws = _Sheet("Comp")
    ws.set_cell(4, 2, "R1")
    ws.set_cell(4, 3, "A")
    ws.set_cell(5, 2, "")
    ws.set_cell(5, 3, "")

    with pytest.raises(ValidationError):
        v1._validate_component_student_identity(
            worksheet=ws,
            sheet_name="Comp",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
            expected_student_count="x",
            expected_student_hash="h",
        )
    with pytest.raises(ValidationError):
        v1._validate_component_student_identity(
            worksheet=ws,
            sheet_name="Comp",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
            expected_student_count=1,
            expected_student_hash="",
        )

    students = [("R1", "A")]
    h = v1._student_identity_hash(students)
    assert (
        v1._validate_component_student_identity(
            worksheet=ws,
            sheet_name="Comp",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
            expected_student_count=1,
            expected_student_hash=h,
        )
        == h
    )

    ws_dup = _Sheet("Comp2")
    ws_dup.set_cell(4, 2, "R1")
    ws_dup.set_cell(4, 3, "A")
    ws_dup.set_cell(5, 2, "R1")
    ws_dup.set_cell(5, 3, "B")
    with pytest.raises(ValidationError):
        v1._extract_component_students(
            worksheet=ws_dup,
            sheet_name="Comp2",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
        )

    with pytest.raises(ValidationError):
        v1._validate_component_structure_snapshot(
            worksheet=ws,
            sheet_name="Comp",
            sheet_kind="BAD",
            header_row=1,
            structure={},
            header_count=4,
        )


def test_non_empty_marks_and_row_consistency_and_anomaly_logs(monkeypatch: pytest.MonkeyPatch) -> None:
    ws = _Sheet("Direct")
    # header_row=1 => data starts row4 for direct types
    ws.set_cell(4, 2, "R1")
    ws.set_cell(4, 3, "A")
    ws.set_cell(5, 2, "")
    ws.set_cell(5, 3, "")
    ws.set_cell(3, 4, 2)
    ws.set_cell(4, 4, "A")
    ws.set_cell(4, 5, "1")  # absence+numeric policy violation
    with pytest.raises(ValidationError):
        v1._validate_non_empty_marks_entries(
            worksheet=ws,
            sheet_name="Direct",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_count=6,
            header_row=1,
        )

    # Row total consistency direct_non_co_wise formula mismatch
    ws2 = _Sheet("ESP")
    ws2.set_cell(4, 2, "R1")
    ws2.set_cell(4, 3, "A")
    ws2.set_cell(5, 2, "")
    ws2.set_cell(5, 3, "")
    ws2.set_cell(4, 5, "not_formula")
    with pytest.raises(ValidationError):
        v1._validate_row_total_consistency(
            worksheet=ws2,
            sheet_name="ESP",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
            header_count=5,
            header_row=1,
            student_count=1,
        )
    # early return branch
    v1._validate_row_total_consistency(
        worksheet=ws2,
        sheet_name="ESP",
        sheet_kind=LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
        header_count=5,
        header_row=1,
        student_count=0,
    )

    warnings: list[tuple[Any, ...]] = []
    monkeypatch.setattr(v1._logger, "warning", lambda *a, **k: warnings.append(a))
    v1._log_marks_anomaly_warnings_from_stats(
        sheet_name="X",
        mark_cols=range(4, 6),
        student_count=10,
        absent_count_by_col={4: 9, 5: 0},
        numeric_count_by_col={4: 0, 5: 10},
        frequency_by_value_by_col={4: {}, 5: {1.0: 10}},
    )
    assert warnings


def test_validate_filled_marks_manifest_schema_formula_anchor_and_cross_sheet_mismatch() -> None:
    # Build two component sheets with different student hashes to trigger cross-sheet mismatch.
    cm = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(cm, list(COURSE_METADATA_HEADERS))
    ac = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ac, list(ASSESSMENT_CONFIG_HEADERS))

    s1 = _Sheet("S1")
    _base_headers(s1, ["#", "Reg", "Name", "Q1", "Total"])
    s1.set_ref("B1", "ok")
    s1.set_cell(4, 2, "R1")
    s1.set_cell(4, 3, "A")
    s1.set_cell(4, 4, 1)
    s1.set_cell(4, 5, "=SUM(D4:D4)")
    s1.set_cell(5, 2, "")
    s1.set_cell(5, 3, "")
    s1.set_cell(3, 4, 2)
    s1.set_cell(3, 5, 2)

    s2 = _Sheet("S2")
    _base_headers(s2, ["#", "Reg", "Name", "Q1", "Total"])
    s2.set_ref("B1", "ok")
    s2.set_cell(4, 2, "R2")
    s2.set_cell(4, 3, "B")
    s2.set_cell(4, 4, 1)
    s2.set_cell(4, 5, "=SUM(D4:D4)")
    s2.set_cell(5, 2, "")
    s2.set_cell(5, 3, "")
    s2.set_cell(3, 4, 2)
    s2.set_cell(3, 5, 2)

    wb = _WB(
        {
            COURSE_METADATA_SHEET: cm,
            ASSESSMENT_CONFIG_SHEET: ac,
            "S1": s1,
            "S2": s2,
        }
    )
    specs = [
        {
            LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
            LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
            LAYOUT_SHEET_SPEC_KEY_HEADERS: list(COURSE_METADATA_HEADERS),
            LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
            LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
        },
        {
            LAYOUT_SHEET_SPEC_KEY_NAME: ASSESSMENT_CONFIG_SHEET,
            LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
            LAYOUT_SHEET_SPEC_KEY_HEADERS: list(ASSESSMENT_CONFIG_HEADERS),
            LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
            LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
        },
        {
            LAYOUT_SHEET_SPEC_KEY_NAME: "S1",
            LAYOUT_SHEET_SPEC_KEY_KIND: LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
            LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "Q1", "Total"],
            LAYOUT_SHEET_SPEC_KEY_ANCHORS: [["B1", "ok"]],
            LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [["E4", "=SUM(D4:D4)"]],
            LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE: {"mark_maxima": [2, 2]},
            LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT: 1,
            LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH: v1._student_identity_hash([("R1", "A")]),
        },
        {
            LAYOUT_SHEET_SPEC_KEY_NAME: "S2",
            LAYOUT_SHEET_SPEC_KEY_KIND: LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
            LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "Q1", "Total"],
            LAYOUT_SHEET_SPEC_KEY_ANCHORS: [["B1", "ok"]],
            LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [["E4", "=SUM(D4:D4)"]],
            LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE: {"mark_maxima": [2, 2]},
            LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT: 1,
            LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH: v1._student_identity_hash([("R2", "B")]),
        },
    ]
    manifest = {
        LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
        LAYOUT_MANIFEST_KEY_SHEETS: specs,
    }
    # Baseline hash is from S1; S2 mismatch should raise cross-sheet mismatch.
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, manifest)


def test_validate_filled_manifest_schema_more_invalid_spec_shapes() -> None:
    cm = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(cm, list(COURSE_METADATA_HEADERS))
    ac = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ac, list(ASSESSMENT_CONFIG_HEADERS))
    wb = _WB({COURSE_METADATA_SHEET: cm, ASSESSMENT_CONFIG_SHEET: ac})

    def _base_manifest(spec: dict[str, Any]) -> dict[str, Any]:
        return {
            LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
            LAYOUT_MANIFEST_KEY_SHEETS: [
                {
                    LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
                    LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                    LAYOUT_SHEET_SPEC_KEY_HEADERS: list(COURSE_METADATA_HEADERS),
                    LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
                    LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
                },
                {
                    LAYOUT_SHEET_SPEC_KEY_NAME: ASSESSMENT_CONFIG_SHEET,
                    LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                    LAYOUT_SHEET_SPEC_KEY_HEADERS: list(ASSESSMENT_CONFIG_HEADERS),
                    LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
                    LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
                },
                spec,
            ],
        }

    bad_spec = {"x": "y"}
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, _base_manifest(bad_spec))

    bad_spec = {
        LAYOUT_SHEET_SPEC_KEY_NAME: "MISSING",
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#"],
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, _base_manifest(bad_spec))

    bad_spec = {
        LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 0,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#"],
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, _base_manifest(bad_spec))

    bad_spec = {
        LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: [],
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, _base_manifest(bad_spec))

    bad_spec = {
        LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: list(COURSE_METADATA_HEADERS),
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: "bad",
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, _base_manifest(bad_spec))

    bad_spec = {
        LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: list(COURSE_METADATA_HEADERS),
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: "bad",
    }
    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, _base_manifest(bad_spec))


def test_course_setup_rule_specific_error_branches() -> None:
    # _header_index_map mismatch
    ws = _Sheet("Meta")
    ws.set_cell(1, 1, "bad")
    with pytest.raises(ValidationError):
        v1._header_index_map(ws, list(COURSE_METADATA_HEADERS))

    # duplicate field / unknown field / missing fields / int/str enforcement / total outcomes invalid
    meta = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(meta, list(COURSE_METADATA_HEADERS))
    meta.set_cell(2, 1, "Course_Code")
    meta.set_cell(2, 2, "C101")
    meta.set_cell(3, 1, "Course_Code")
    meta.set_cell(3, 2, "C102")
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta)

    meta2 = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(meta2, list(COURSE_METADATA_HEADERS))
    meta2.set_cell(2, 1, "Unknown_Field")
    meta2.set_cell(2, 2, "x")
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta2)

    meta3 = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(meta3, list(COURSE_METADATA_HEADERS))
    meta3.set_cell(2, 1, "Course_Code")
    meta3.set_cell(2, 2, "C101")
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta3)

    # assessment duplicate / missing direct/indirect / total mismatches
    ass = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ass, list(ASSESSMENT_CONFIG_HEADERS))
    ass.set_cell(2, 1, "S1")
    ass.set_cell(2, 2, 100)
    ass.set_cell(2, 3, "yes")
    ass.set_cell(2, 4, "yes")
    ass.set_cell(2, 5, "yes")
    ass.set_cell(3, 1, "S1")
    ass.set_cell(3, 2, 0)
    ass.set_cell(3, 3, "yes")
    ass.set_cell(3, 4, "yes")
    ass.set_cell(3, 5, "yes")
    with pytest.raises(ValidationError):
        v1._validate_assessment_config(ass)

    ass2 = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ass2, list(ASSESSMENT_CONFIG_HEADERS))
    ass2.set_cell(2, 1, "S1")
    ass2.set_cell(2, 2, 100)
    ass2.set_cell(2, 3, "yes")
    ass2.set_cell(2, 4, "yes")
    ass2.set_cell(2, 5, "no")
    with pytest.raises(ValidationError):
        v1._validate_assessment_config(ass2)

    # question map unknown/label/max/co errors
    qm = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm, list(QUESTION_MAP_HEADERS))
    qm.set_cell(2, 1, "missing")
    qm.set_cell(2, 2, "Q1")
    qm.set_cell(2, 3, 1)
    qm.set_cell(2, 4, "CO1")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm, {"s1": {"co_wise_breakup": True}}, 2)

    qm2 = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm2, list(QUESTION_MAP_HEADERS))
    qm2.set_cell(2, 1, "s1")
    qm2.set_cell(2, 2, "")
    qm2.set_cell(2, 3, 1)
    qm2.set_cell(2, 4, "CO1")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm2, {"s1": {"co_wise_breakup": True}}, 2)

    qm3 = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm3, list(QUESTION_MAP_HEADERS))
    qm3.set_cell(2, 1, "s1")
    qm3.set_cell(2, 2, "Q1")
    qm3.set_cell(2, 3, "bad")
    qm3.set_cell(2, 4, "CO1")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm3, {"s1": {"co_wise_breakup": True}}, 2)

    qm4 = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm4, list(QUESTION_MAP_HEADERS))
    qm4.set_cell(2, 1, "s1")
    qm4.set_cell(2, 2, "Q1")
    qm4.set_cell(2, 3, 0)
    qm4.set_cell(2, 4, "CO1")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm4, {"s1": {"co_wise_breakup": True}}, 2)

    qm5 = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm5, list(QUESTION_MAP_HEADERS))
    qm5.set_cell(2, 1, "s1")
    qm5.set_cell(2, 2, "Q1")
    qm5.set_cell(2, 3, 1)
    qm5.set_cell(2, 4, "")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm5, {"s1": {"co_wise_breakup": True}}, 2)

    qm6 = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm6, list(QUESTION_MAP_HEADERS))
    qm6.set_cell(2, 1, "s1")
    qm6.set_cell(2, 2, "Q1")
    qm6.set_cell(2, 3, 1)
    qm6.set_cell(2, 4, "CO1,CO1")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm6, {"s1": {"co_wise_breakup": True}}, 2)

    qm7 = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm7, list(QUESTION_MAP_HEADERS))
    qm7.set_cell(2, 1, "s1")
    qm7.set_cell(2, 2, "Q1")
    qm7.set_cell(2, 3, 1)
    qm7.set_cell(2, 4, "CO9")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm7, {"s1": {"co_wise_breakup": True}}, 2)

    qm8 = _Sheet(QUESTION_MAP_SHEET)
    _base_headers(qm8, list(QUESTION_MAP_HEADERS))
    qm8.set_cell(2, 1, "s1")
    qm8.set_cell(2, 2, "Q1")
    qm8.set_cell(2, 3, 1)
    qm8.set_cell(2, 4, "CO1")
    qm8.set_cell(3, 1, "s1")
    qm8.set_cell(3, 2, "Q1")
    qm8.set_cell(3, 3, 1)
    qm8.set_cell(3, 4, "CO1")
    with pytest.raises(ValidationError):
        v1._validate_question_map(qm8, {"s1": {"co_wise_breakup": True}}, 2)

    # matching helper bool-return false path
    assert v1._filled_marks_values_match(1, "x") is False


def test_additional_snapshot_and_log_marks_anomaly_wrapper_paths() -> None:
    ws = _Sheet("S")
    with pytest.raises(ValidationError):
        v1._validate_component_structure_snapshot(
            worksheet=ws,
            sheet_name="S",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
            structure=None,
            header_count=5,
        )

    with pytest.raises(ValidationError):
        v1._validate_component_structure_snapshot(
            worksheet=ws,
            sheet_name="S",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
            structure={"mark_maxima": "bad"},
            header_count=5,
        )

    with pytest.raises(ValidationError):
        v1._validate_component_structure_snapshot(
            worksheet=ws,
            sheet_name="S",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
            header_row=1,
            structure={"mark_maxima": "bad"},
            header_count=5,
        )

    with pytest.raises(ValidationError):
        v1._validate_component_structure_snapshot(
            worksheet=ws,
            sheet_name="S",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_NON_CO_WISE,
            header_row=1,
            structure={"mark_maxima": [1, 2], "x": 1},
            header_count=5,
        )

    with pytest.raises(ValidationError):
        v1._validate_component_structure_snapshot(
            worksheet=ws,
            sheet_name="S",
            sheet_kind=LAYOUT_SHEET_KIND_INDIRECT,
            header_row=1,
            structure={"likert_range": [1, 6]},
            header_count=6,
        )

    # _log_marks_anomaly_warnings wrapper path with zero students -> early return.
    v1._log_marks_anomaly_warnings(
        worksheet=ws,
        sheet_name="S",
        sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
        header_count=5,
        header_row=1,
    )


def _build_metadata_sheet_with_sample_values() -> _Sheet:
    meta = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(meta, list(COURSE_METADATA_HEADERS))
    for row_idx, (field, value) in enumerate(SAMPLE_SETUP_DATA[COURSE_METADATA_SHEET], start=2):
        meta.set_cell(row_idx, 1, field)
        meta.set_cell(row_idx, 2, value)
    return meta


def test_manifest_anchor_formula_and_no_component_extra_branches() -> None:
    cm = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(cm, list(COURSE_METADATA_HEADERS))
    ac = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ac, list(ASSESSMENT_CONFIG_HEADERS))
    wb = _WB({COURSE_METADATA_SHEET: cm, ASSESSMENT_CONFIG_SHEET: ac})

    meta_base = {
        LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: list(COURSE_METADATA_HEADERS),
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }
    assess_base = {
        LAYOUT_SHEET_SPEC_KEY_NAME: ASSESSMENT_CONFIG_SHEET,
        LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
        LAYOUT_SHEET_SPEC_KEY_HEADERS: list(ASSESSMENT_CONFIG_HEADERS),
        LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
        LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
    }

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(
            wb,
            {
                LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
                LAYOUT_MANIFEST_KEY_SHEETS: [meta_base, "bad"],
            },
        )

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(
            wb,
            {
                LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
                LAYOUT_MANIFEST_KEY_SHEETS: [meta_base, assess_base],
            },
        )

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(
            wb,
            {
                LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
                LAYOUT_MANIFEST_KEY_SHEETS: [
                    {**meta_base, LAYOUT_SHEET_SPEC_KEY_ANCHORS: [["A1"]]},
                    assess_base,
                ],
            },
        )

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(
            wb,
            {
                LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
                LAYOUT_MANIFEST_KEY_SHEETS: [
                    {**meta_base, LAYOUT_SHEET_SPEC_KEY_ANCHORS: [[1, "Field"]]},
                    assess_base,
                ],
            },
        )

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(
            wb,
            {
                LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
                LAYOUT_MANIFEST_KEY_SHEETS: [
                    {**meta_base, LAYOUT_SHEET_SPEC_KEY_ANCHORS: [["A1", "WRONG"]]},
                    assess_base,
                ],
            },
        )

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(
            wb,
            {
                LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
                LAYOUT_MANIFEST_KEY_SHEETS: [
                    {**meta_base, LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [["A1"]]},
                    assess_base,
                ],
            },
        )

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(
            wb,
            {
                LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
                LAYOUT_MANIFEST_KEY_SHEETS: [
                    {**meta_base, LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [[1, "=X"]]},
                    assess_base,
                ],
            },
        )


def test_more_metadata_assessment_and_student_identity_branches() -> None:
    meta_int_invalid = _build_metadata_sheet_with_sample_values()
    meta_int_invalid.set_cell(8, 2, "six")
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta_int_invalid)

    meta_str_invalid = _build_metadata_sheet_with_sample_values()
    meta_str_invalid.set_cell(3, 2, 123)
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta_str_invalid)

    meta_total_invalid = _build_metadata_sheet_with_sample_values()
    meta_total_invalid.set_cell(8, 2, 0)
    with pytest.raises(ValidationError):
        v1._validate_course_metadata(meta_total_invalid)

    ass_missing_component = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ass_missing_component, list(ASSESSMENT_CONFIG_HEADERS))
    ass_missing_component.set_cell(2, 1, "")
    ass_missing_component.set_cell(2, 2, 100)
    ass_missing_component.set_cell(2, 3, "yes")
    ass_missing_component.set_cell(2, 4, "yes")
    ass_missing_component.set_cell(2, 5, "yes")
    with pytest.raises(ValidationError):
        v1._validate_assessment_config(ass_missing_component)

    ass_indirect_missing = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ass_indirect_missing, list(ASSESSMENT_CONFIG_HEADERS))
    ass_indirect_missing.set_cell(2, 1, "S1")
    ass_indirect_missing.set_cell(2, 2, 100)
    ass_indirect_missing.set_cell(2, 3, "yes")
    ass_indirect_missing.set_cell(2, 4, "yes")
    ass_indirect_missing.set_cell(2, 5, "yes")
    with pytest.raises(ValidationError):
        v1._validate_assessment_config(ass_indirect_missing)

    ass_indirect_total_bad = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ass_indirect_total_bad, list(ASSESSMENT_CONFIG_HEADERS))
    ass_indirect_total_bad.set_cell(2, 1, "S1")
    ass_indirect_total_bad.set_cell(2, 2, 100)
    ass_indirect_total_bad.set_cell(2, 3, "yes")
    ass_indirect_total_bad.set_cell(2, 4, "yes")
    ass_indirect_total_bad.set_cell(2, 5, "yes")
    ass_indirect_total_bad.set_cell(3, 1, "CSURVEY")
    ass_indirect_total_bad.set_cell(3, 2, 90)
    ass_indirect_total_bad.set_cell(3, 3, "no")
    ass_indirect_total_bad.set_cell(3, 4, "yes")
    ass_indirect_total_bad.set_cell(3, 5, "no")
    with pytest.raises(ValidationError):
        v1._validate_assessment_config(ass_indirect_total_bad)

    ws_students = _Sheet("Comp")
    ws_students.set_cell(4, 2, "R1")
    ws_students.set_cell(4, 3, "A")
    ws_students.set_cell(5, 2, "")
    ws_students.set_cell(5, 3, "")
    with pytest.raises(ValidationError):
        v1._validate_component_student_identity(
            worksheet=ws_students,
            sheet_name="Comp",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
            expected_student_count=2,
            expected_student_hash="hash",
        )

    ws_partial = _Sheet("CompPartial")
    ws_partial.set_cell(4, 2, "R1")
    ws_partial.set_cell(4, 3, "")
    with pytest.raises(ValidationError):
        v1._extract_component_students(
            worksheet=ws_partial,
            sheet_name="CompPartial",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_row=1,
        )

    ws_empty = _Sheet("Empty")
    v1._validate_non_empty_marks_entries(
        worksheet=ws_empty,
        sheet_name="Empty",
        sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
        header_count=5,
        header_row=1,
    )

    ws_anomaly = _Sheet("Anomaly")
    ws_anomaly.set_cell(2, 2, "R1")
    ws_anomaly.set_cell(2, 3, "A")
    ws_anomaly.set_cell(3, 2, "")
    ws_anomaly.set_cell(3, 3, "")
    ws_anomaly.set_cell(2, 4, "A")
    ws_anomaly.set_cell(2, 5, 10)
    v1._log_marks_anomaly_warnings(
        worksheet=ws_anomaly,
        sheet_name="Anomaly",
        sheet_kind=LAYOUT_SHEET_KIND_INDIRECT,
        header_count=5,
        header_row=1,
    )


def test_manifest_cross_sheet_identity_mismatch_branch(monkeypatch: pytest.MonkeyPatch) -> None:
    cm = _Sheet(COURSE_METADATA_SHEET)
    _base_headers(cm, list(COURSE_METADATA_HEADERS))
    ac = _Sheet(ASSESSMENT_CONFIG_SHEET)
    _base_headers(ac, list(ASSESSMENT_CONFIG_HEADERS))
    s1 = _Sheet("S1")
    _base_headers(s1, ["#", "Reg", "Name", "Q1", "Total"])
    s2 = _Sheet("S2")
    _base_headers(s2, ["#", "Reg", "Name", "Q1", "Total"])
    wb = _WB({COURSE_METADATA_SHEET: cm, ASSESSMENT_CONFIG_SHEET: ac, "S1": s1, "S2": s2})

    manifest = {
        LAYOUT_MANIFEST_KEY_SHEET_ORDER: wb.sheetnames,
        LAYOUT_MANIFEST_KEY_SHEETS: [
            {
                LAYOUT_SHEET_SPEC_KEY_NAME: COURSE_METADATA_SHEET,
                LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                LAYOUT_SHEET_SPEC_KEY_HEADERS: list(COURSE_METADATA_HEADERS),
                LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
                LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
            },
            {
                LAYOUT_SHEET_SPEC_KEY_NAME: ASSESSMENT_CONFIG_SHEET,
                LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                LAYOUT_SHEET_SPEC_KEY_HEADERS: list(ASSESSMENT_CONFIG_HEADERS),
                LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
                LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
            },
            {
                LAYOUT_SHEET_SPEC_KEY_NAME: "S1",
                LAYOUT_SHEET_SPEC_KEY_KIND: LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
                LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "Q1", "Total"],
                LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
                LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
                LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE: {},
                LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT: 1,
                LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH: "unused-1",
            },
            {
                LAYOUT_SHEET_SPEC_KEY_NAME: "S2",
                LAYOUT_SHEET_SPEC_KEY_KIND: LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
                LAYOUT_SHEET_SPEC_KEY_HEADER_ROW: 1,
                LAYOUT_SHEET_SPEC_KEY_HEADERS: ["#", "Reg", "Name", "Q1", "Total"],
                LAYOUT_SHEET_SPEC_KEY_ANCHORS: [],
                LAYOUT_SHEET_SPEC_KEY_FORMULA_ANCHORS: [],
                LAYOUT_SHEET_SPEC_KEY_MARK_STRUCTURE: {},
                LAYOUT_SHEET_SPEC_KEY_STUDENT_COUNT: 1,
                LAYOUT_SHEET_SPEC_KEY_STUDENT_IDENTITY_HASH: "unused-2",
            },
        ],
    }

    hashes = iter(("H1", "H2"))
    monkeypatch.setattr(v1, "_validate_component_structure_snapshot", lambda **_k: None)
    monkeypatch.setattr(v1, "_validate_non_empty_marks_entries", lambda **_k: None)
    monkeypatch.setattr(v1, "_validate_component_student_identity", lambda **_k: next(hashes))

    with pytest.raises(ValidationError):
        v1.validate_filled_marks_manifest_schema(wb, manifest)


def test_co_tokens_non_matching_token_branch() -> None:
    assert v1._co_tokens("not-a-co-token") == []
