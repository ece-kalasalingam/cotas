from typing import Any

import pytest

from common.constants import COURSE_METADATA_FACULTY_NAME_KEY
from common.exceptions import ValidationError
from domain import instructor_template_engine_sheetops as ops


class _WS:
    def __getattr__(self, _name: str):  # type: ignore[no-untyped-def]
        return lambda *_a, **_k: None


class _WB:
    def __init__(self, *, has_hash_sheet: bool = False) -> None:
        self.sheetnames = [ops.SYSTEM_HASH_SHEET] if has_hash_sheet else []
        self._src_ws = _WS()
        self._dst_ws = _WS()

    def __getitem__(self, _name: str) -> _WS:
        return self._src_ws

    def add_worksheet(self, _name: str) -> _WS:
        return self._dst_ws

    def add_format(self, payload: dict[str, Any]) -> dict[str, Any]:
        return payload


def test_sheetops_helper_uncovered_branches(monkeypatch: pytest.MonkeyPatch) -> None:
    # _copy_system_hash_sheet early return.
    ops._copy_system_hash_sheet(_WB(has_hash_sheet=False), _WB())

    # _filter_marks_template_metadata_rows short row continue.
    rows = [["only_one"], [COURSE_METADATA_FACULTY_NAME_KEY, "x"], ["Course_Code", "C101"]]
    assert ops._filter_marks_template_metadata_rows(rows) == [["Course_Code", "C101"]]

    monkeypatch.setattr(ops, "ensure_workbook_secret_policy", lambda: None)
    monkeypatch.setattr(ops, "get_workbook_password", lambda: "secret")

    wb = _WB()
    fmt = object()
    metadata_rows = [["Course_Code", "C101"]]
    students = [("R1", "Alice")]

    # Preview-student sample append branches in all three writers.
    ops._write_direct_co_wise_sheet(
        wb,
        "S1",
        metadata_rows,
        "S1",
        students,
        [{"max_marks": 2, "co_values": [1]}],
        fmt,
        fmt,
        fmt,
        fmt,
        fmt,
        fmt,
        fmt,
    )
    ops._write_direct_non_co_wise_sheet(
        wb,
        "ESP",
        metadata_rows,
        "ESP",
        students,
        [{"max_marks": 10, "co_values": [1]}],
        fmt,
        fmt,
        fmt,
        fmt,
        fmt,
        fmt,
        fmt,
    )
    ops._write_indirect_sheet(
        wb,
        "CSURVEY",
        metadata_rows,
        "CSURVEY",
        students,
        2,
        fmt,
        fmt,
        fmt,
        fmt,
        fmt,
    )

    # _build_marks_validation_error_message non-numeric max path.
    msg = ops._build_marks_validation_error_message("invalid-max")
    assert "invalid-max" in msg

    # _split_equal_with_residual edge paths.
    assert ops._split_equal_with_residual(10.0, 0) == []
    assert ops._split_equal_with_residual(10.0, 1) == [10.0]

    # _safe_sheet_name duplicate/while-loop/counter increment path.
    used = {ops.normalize("Comp"), ops.normalize("Comp_2")}
    assert ops._safe_sheet_name("Comp", used).lower() == "comp_3"

    # generate_worksheet validation branches.
    with pytest.raises(ValidationError):
        ops.generate_worksheet(_WB(), "", ["A"], [], {}, {})
    with pytest.raises(ValidationError):
        ops.generate_worksheet(_WB(), "S", [], [], {}, {})
    with pytest.raises(ValidationError):
        ops.generate_worksheet(_WB(), "S", ["A", "A"], [], {}, {})
    with pytest.raises(ValidationError):
        ops.generate_worksheet(_WB(), "S", ["A", "B"], [["x"]], {}, {})
