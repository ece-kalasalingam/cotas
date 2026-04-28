from __future__ import annotations

import sys
from dataclasses import dataclass
from types import SimpleNamespace
from typing import Any, cast

import pytest

from common.constants import LAYOUT_SHEET_KIND_DIRECT_CO_WISE
from common.error_catalog import validation_error_from_key
from common.exceptions import ValidationError
from domain import template_strategy_router
from domain.template_versions.course_setup_v2_impl import (
    marks_template_validator as v2_validator,
)


def _identity(
    *,
    section: str,
    course_code: str = "CS101",
    semester: str = "V",
    year: str = "2026-27",
    outcomes: int = 3,
    reg_numbers: frozenset[str] | None = None,
) -> object:
    """Identity.
    
    Args:
        section: Parameter value (str).
        course_code: Parameter value (str).
        semester: Parameter value (str).
        year: Parameter value (str).
        outcomes: Parameter value (int).
        reg_numbers: Parameter value (frozenset[str] | None).
    
    Returns:
        object: Return value.
    
    Raises:
        None.
    """
    resolved_reg_numbers = reg_numbers if reg_numbers is not None else frozenset({f"reg_{section.lower()}"})
    return v2_validator._MarksWorkbookIdentity(
        template_id="COURSE_SETUP_V2",
        course_code=course_code,
        semester=semester,
        academic_year=year,
        total_outcomes=outcomes,
        section=section,
        reg_numbers=resolved_reg_numbers,
    )


def test_v2_marks_validator_warning_buffer_consume_clears_state() -> None:
    """Test v2 marks validator warning buffer consume clears state.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    v2_validator._reset_marks_anomaly_warnings()
    v2_validator._log_marks_anomaly_warnings_from_stats(
        sheet_name="Sheet1",
        mark_cols=range(4, 6),
        student_count=10,
        absent_count_by_col={4: 5, 5: 0},
        numeric_count_by_col={4: 0, 5: 10},
        frequency_by_value_by_col={4: {}, 5: {1.0: 10}},
    )

    warnings = v2_validator.consume_last_marks_anomaly_warnings()
    if not (len(warnings) == 2):
        raise AssertionError('assertion failed')
    if not (any("High absence ratio" in message for message in warnings)):
        raise AssertionError('assertion failed')
    if not (any("Near-constant marks" in message for message in warnings)):
        raise AssertionError('assertion failed')
    if not (v2_validator.consume_last_marks_anomaly_warnings() == []):
        raise AssertionError('assertion failed')


def test_v2_marks_validator_resets_warning_buffer_per_run() -> None:
    """Test v2 marks validator resets warning buffer per run.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    v2_validator._last_marks_anomaly_warnings[:] = ["stale warning"]

    class _Workbook:
        sheetnames = ["OnlySheet"]

    with pytest.raises(Exception):
        v2_validator.validate_filled_marks_manifest_schema(
            workbook=_Workbook(),
            manifest={"sheet_order": [], "sheets": []},
        )

    # stale warnings should be cleared at validation start
    if not ("stale warning" not in v2_validator.consume_last_marks_anomaly_warnings()):
        raise AssertionError('assertion failed')


def test_v2_marks_batch_validator_returns_course_like_shape(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test v2 marks batch validator returns course like shape.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
        """Fake validate.
        
        Args:
            workbook_path: Parameter value (str).
            expected_template_id: Parameter value (str).
            cancel_token: Parameter value (object | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        del expected_template_id
        del cancel_token
        if "invalid" in workbook_path:
            raise ValidationError("bad marks", code="COA_MARK_ENTRY_EMPTY", context={"cell": "D10"})
        return _identity(section="A")

    monkeypatch.setattr(v2_validator, "_validate_filled_marks_workbook_impl", _fake_validate)
    result = v2_validator.validate_filled_marks_workbooks(
        ["valid.xlsx", "valid.xlsx", "invalid.xlsx"],
        template_id="COURSE_SETUP_V2",
    )

    if not (result["valid_paths"] == ["valid.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["invalid_paths"] == ["invalid.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["duplicate_paths"] == ["valid.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["mismatched_paths"] == []):
        raise AssertionError('assertion failed')
    if not (result["duplicate_sections"] == []):
        raise AssertionError('assertion failed')
    if not (isinstance(result["template_ids"], dict)):
        raise AssertionError('assertion failed')
    if not (result["rejections"]):
        raise AssertionError('assertion failed')


def test_v2_marks_batch_validator_tracks_template_mismatch(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test v2 marks batch validator tracks template mismatch.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
        """Fake validate.
        
        Args:
            workbook_path: Parameter value (str).
            expected_template_id: Parameter value (str).
            cancel_token: Parameter value (object | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        del workbook_path
        del expected_template_id
        del cancel_token
        raise ValidationError("template mismatch", code="UNKNOWN_TEMPLATE", context={"template_id": "COURSE_SETUP_X"})

    monkeypatch.setattr(v2_validator, "_validate_filled_marks_workbook_impl", _fake_validate)
    result = cast(
        dict[str, Any],
        v2_validator.validate_filled_marks_workbooks(
            ["wrong_template.xlsx"],
            template_id="COURSE_SETUP_V2",
        ),
    )

    if not (result["valid_paths"] == []):
        raise AssertionError('assertion failed')
    if not (result["invalid_paths"] == ["wrong_template.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["mismatched_paths"] == ["wrong_template.xlsx"]):
        raise AssertionError('assertion failed')
    rejection = cast(list[dict[str, Any]], result["rejections"])[0]
    if not (rejection["reason_kind"] == "template_mismatch"):
        raise AssertionError('assertion failed')


@dataclass(frozen=True)
class _Payload:
    template_id: str
    manifest: dict[str, object]


def test_marks_impl_two_stage_trust_flow_uses_read_only_then_full_workbook(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    """Test marks impl two stage trust flow uses read only then full workbook.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_path = tmp_path / "marks.xlsx"
    workbook_path.write_text("x")

    calls: list[dict[str, object]] = []
    workbooks: list[Any] = []

    class _Workbook:
        def __init__(self, mode: str) -> None:
            """Init.
            
            Args:
                mode: Parameter value (str).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self.mode = mode
            self.closed = False
            self.sheetnames = ["Sheet1"]

        def close(self) -> None:
            """Close.
            
            Args:
                None.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self.closed = True

    def _load_workbook(_path, **kwargs):
        """Load workbook.
        
        Args:
            _path: Parameter value.
            kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        calls.append(dict(kwargs))
        mode = "read_only" if kwargs.get("read_only", False) else "full"
        workbook = _Workbook(mode=mode)
        workbooks.append(workbook)
        return workbook

    openpyxl_stub = SimpleNamespace(load_workbook=_load_workbook)
    monkeypatch.setitem(sys.modules, "openpyxl", openpyxl_stub)

    monkeypatch.setattr(
        template_strategy_router,
        "read_valid_system_workbook_payload",
        lambda workbook: _Payload(template_id="COURSE_SETUP_V2", manifest={"sheet_order": [], "sheets": []}),
    )
    monkeypatch.setattr(
        template_strategy_router,
        "assert_template_id_matches",
        lambda *, actual_template_id, expected_template_id: None,
    )

    manifest_calls: list[tuple[object, object]] = []
    monkeypatch.setattr(
        v2_validator,
        "validate_filled_marks_manifest_schema",
        lambda *, workbook, manifest: manifest_calls.append((workbook, manifest)),
    )
    monkeypatch.setattr(
        v2_validator,
        "_read_marks_workbook_identity",
        lambda *, workbook, template_id: _identity(section="A"),
    )

    resolved = v2_validator._validate_filled_marks_workbook_impl(
        workbook_path=workbook_path,
        expected_template_id="COURSE_SETUP_V2",
    )

    if not (getattr(resolved, "template_id", "") == "COURSE_SETUP_V2"):
        raise AssertionError('assertion failed')
    if calls[0].get("read_only") is not True:
        raise AssertionError('assertion failed')
    if calls[1].get("read_only", False) is not False:
        raise AssertionError('assertion failed')
    if workbooks[0].closed is not True:
        raise AssertionError('assertion failed')
    if workbooks[1].closed is not True:
        raise AssertionError('assertion failed')
    if not (len(manifest_calls) == 1):
        raise AssertionError('assertion failed')
    if manifest_calls[0][0] is not workbooks[1]:
        raise AssertionError('assertion failed')


def test_marks_impl_manifest_validation_error_is_preserved(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    """Test marks impl manifest validation error is preserved.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_path = tmp_path / "marks.xlsx"
    workbook_path.write_text("x")

    class _Workbook:
        sheetnames = ["Sheet1"]

        def close(self) -> None:
            """Close.
            
            Args:
                None.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            return None

    monkeypatch.setitem(
        sys.modules,
        "openpyxl",
        SimpleNamespace(load_workbook=lambda *_args, **_kwargs: _Workbook()),
    )
    monkeypatch.setattr(
        template_strategy_router,
        "read_valid_system_workbook_payload",
        lambda _workbook: _Payload(template_id="COURSE_SETUP_V2", manifest={"sheet_order": [], "sheets": []}),
    )
    monkeypatch.setattr(
        template_strategy_router,
        "assert_template_id_matches",
        lambda *, actual_template_id, expected_template_id: None,
    )
    monkeypatch.setattr(
        v2_validator,
        "validate_filled_marks_manifest_schema",
        lambda *, workbook, manifest: (_ for _ in ()).throw(
            ValidationError("schema bad", code="COA_MARK_ENTRY_EMPTY", context={"cell": "D7"})
        ),
    )

    with pytest.raises(ValidationError) as excinfo:
        v2_validator._validate_filled_marks_workbook_impl(
            workbook_path=workbook_path,
            expected_template_id="COURSE_SETUP_V2",
        )
    if not (excinfo.value.code == "COA_MARK_ENTRY_EMPTY"):
        raise AssertionError('assertion failed')


def test_marks_impl_raises_open_failed_when_read_only_open_fails(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    """Test marks impl raises open failed when read only open fails.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_path = tmp_path / "corrupt_marks.xlsx"
    workbook_path.write_text("not-an-xlsx", encoding="utf-8")

    def _raise_open(*_args, **_kwargs):
        """Raise open.
        
        Args:
            _args: Parameter value.
            _kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        raise OSError("broken workbook zip")

    monkeypatch.setitem(
        sys.modules,
        "openpyxl",
        SimpleNamespace(load_workbook=_raise_open),
    )

    with pytest.raises(ValidationError) as excinfo:
        v2_validator._validate_filled_marks_workbook_impl(
            workbook_path=workbook_path,
            expected_template_id="COURSE_SETUP_V2",
        )

    if not (excinfo.value.code == "WORKBOOK_OPEN_FAILED"):
        raise AssertionError('assertion failed')
    if not (str(excinfo.value.context.get("workbook", "")) == str(workbook_path)):
        raise AssertionError('assertion failed')


def test_marks_impl_raises_open_failed_when_full_open_fails_after_payload_read(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    """Test marks impl raises open failed when full open fails after payload read.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    workbook_path = tmp_path / "corrupt_marks_second_pass.xlsx"
    workbook_path.write_text("not-an-xlsx", encoding="utf-8")

    class _Workbook:
        sheetnames = ["Sheet1"]

        def close(self) -> None:
            """Close.
            
            Args:
                None.
            
            Returns:
                None.
            
            Raises:
                None.
            """
            return None

    read_only_workbook = _Workbook()
    load_calls: list[dict[str, object]] = []

    def _load_workbook(_path, **kwargs):
        """Load workbook.
        
        Args:
            _path: Parameter value.
            kwargs: Parameter value.
        
        Returns:
            None.
        
        Raises:
            None.
        """
        load_calls.append(dict(kwargs))
        if kwargs.get("read_only", False):
            return read_only_workbook
        raise OSError("full-open failed for corrupt workbook")

    monkeypatch.setitem(
        sys.modules,
        "openpyxl",
        SimpleNamespace(load_workbook=_load_workbook),
    )
    monkeypatch.setattr(
        template_strategy_router,
        "read_valid_system_workbook_payload",
        lambda _workbook: _Payload(template_id="COURSE_SETUP_V2", manifest={"sheet_order": [], "sheets": []}),
    )
    monkeypatch.setattr(
        template_strategy_router,
        "assert_template_id_matches",
        lambda *, actual_template_id, expected_template_id: None,
    )

    with pytest.raises(ValidationError) as excinfo:
        v2_validator._validate_filled_marks_workbook_impl(
            workbook_path=workbook_path,
            expected_template_id="COURSE_SETUP_V2",
        )

    if load_calls[0].get("read_only") is not True:
        raise AssertionError('assertion failed')
    if load_calls[1].get("read_only", False) is not False:
        raise AssertionError('assertion failed')
    if not (excinfo.value.code == "WORKBOOK_OPEN_FAILED"):
        raise AssertionError('assertion failed')
    if not (str(excinfo.value.context.get("workbook", "")) == str(workbook_path)):
        raise AssertionError('assertion failed')


def test_marks_impl_rejects_symlink_path(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    workbook_path = tmp_path / "linked_marks.xlsx"
    workbook_path.write_text("x", encoding="utf-8")

    monkeypatch.setitem(
        sys.modules,
        "openpyxl",
        SimpleNamespace(load_workbook=lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("unexpected"))),
    )
    monkeypatch.setattr(
        v2_validator,
        "assert_not_symlink_path",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            validation_error_from_key(
                "common.validation_failed_invalid_data",
                code="WORKBOOK_SYMLINK_NOT_ALLOWED",
                workbook=str(workbook_path),
            )
        ),
    )

    with pytest.raises(ValidationError) as excinfo:
        v2_validator._validate_filled_marks_workbook_impl(
            workbook_path=workbook_path,
            expected_template_id="COURSE_SETUP_V2",
        )

    if not (excinfo.value.code == "WORKBOOK_SYMLINK_NOT_ALLOWED"):
        raise AssertionError('assertion failed')


def test_v2_marks_batch_validator_tracks_cohort_mismatch(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test v2 marks batch validator tracks cohort mismatch.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
        """Fake validate.
        
        Args:
            workbook_path: Parameter value (str).
            expected_template_id: Parameter value (str).
            cancel_token: Parameter value (object | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        del expected_template_id
        del cancel_token
        if workbook_path == "ok.xlsx":
            return _identity(section="A", year="2026-27")
        return _identity(section="B", year="2025-26")

    monkeypatch.setattr(v2_validator, "_validate_filled_marks_workbook_impl", _fake_validate)
    result = cast(
        dict[str, Any],
        v2_validator.validate_filled_marks_workbooks(
            ["ok.xlsx", "cohort_bad.xlsx"],
            template_id="COURSE_SETUP_V2",
        ),
    )

    if not (result["valid_paths"] == ["ok.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["mismatched_paths"] == ["cohort_bad.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["invalid_paths"] == []):
        raise AssertionError('assertion failed')
    rejection = next(
        item
        for item in cast(list[dict[str, Any]], result["rejections"])
        if item["path"] == "cohort_bad.xlsx"
    )
    if not (rejection["reason_kind"] == "cohort_mismatch"):
        raise AssertionError('assertion failed')
    if not (rejection["issue"]["code"] == "MARKS_TEMPLATE_COHORT_MISMATCH"):
        raise AssertionError('assertion failed')


def test_v2_marks_batch_validator_tracks_duplicate_section(monkeypatch: pytest.MonkeyPatch) -> None:
    """Test v2 marks batch validator tracks duplicate section.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
        """Fake validate.
        
        Args:
            workbook_path: Parameter value (str).
            expected_template_id: Parameter value (str).
            cancel_token: Parameter value (object | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        del expected_template_id
        del cancel_token
        if workbook_path == "first.xlsx":
            return _identity(section="A")
        return _identity(section="A")

    monkeypatch.setattr(v2_validator, "_validate_filled_marks_workbook_impl", _fake_validate)
    result = cast(
        dict[str, Any],
        v2_validator.validate_filled_marks_workbooks(
            ["first.xlsx", "dup_section.xlsx"],
            template_id="COURSE_SETUP_V2",
        ),
    )

    if not (result["valid_paths"] == ["first.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["duplicate_sections"] == ["dup_section.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["invalid_paths"] == []):
        raise AssertionError('assertion failed')
    rejection = next(
        item
        for item in cast(list[dict[str, Any]], result["rejections"])
        if item["path"] == "dup_section.xlsx"
    )
    if not (rejection["reason_kind"] == "duplicate_section"):
        raise AssertionError('assertion failed')
    if not (rejection["issue"]["code"] == "MARKS_TEMPLATE_SECTION_DUPLICATE"):
        raise AssertionError('assertion failed')


def test_v2_marks_batch_validator_rejects_cross_workbook_duplicate_reg_no(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Test v2 marks batch validator rejects cross workbook duplicate reg no.
    
    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
        """Fake validate.
        
        Args:
            workbook_path: Parameter value (str).
            expected_template_id: Parameter value (str).
            cancel_token: Parameter value (object | None).
        
        Returns:
            object: Return value.
        
        Raises:
            None.
        """
        del expected_template_id
        del cancel_token
        if workbook_path == "first.xlsx":
            return _identity(section="A", reg_numbers=frozenset({"001", "002"}))
        return _identity(section="B", reg_numbers=frozenset({"002", "003"}))

    monkeypatch.setattr(v2_validator, "_validate_filled_marks_workbook_impl", _fake_validate)
    result = cast(
        dict[str, Any],
        v2_validator.validate_filled_marks_workbooks(
            ["first.xlsx", "second.xlsx"],
            template_id="COURSE_SETUP_V2",
        ),
    )

    if not (result["valid_paths"] == ["first.xlsx"]):
        raise AssertionError('assertion failed')
    if not (result["invalid_paths"] == ["second.xlsx"]):
        raise AssertionError('assertion failed')
    rejection = next(
        item
        for item in cast(list[dict[str, Any]], result["rejections"])
        if item["path"] == "second.xlsx"
    )
    if not (rejection["reason_kind"] == "duplicate_reg_no"):
        raise AssertionError('assertion failed')
    if not (rejection["issue"]["code"] == "MARKS_TEMPLATE_STUDENT_REG_DUPLICATE"):
        raise AssertionError('assertion failed')


def test_non_empty_marks_entries_collects_multiple_row_failures() -> None:
    """Test non empty marks entries collects multiple row failures.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    class _Cell:
        def __init__(self, row: int, column: int, value: object) -> None:
            """Init.
            
            Args:
                row: Parameter value (int).
                column: Parameter value (int).
                value: Parameter value (object).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self.row = row
            self.column = column
            self.value = value
            self.coordinate = f"{v2_validator._excel_col_name(column)}{row}"

    class _Worksheet:
        def __init__(self, values: dict[tuple[int, int], object]) -> None:
            """Init.
            
            Args:
                values: Parameter value (dict[tuple[int, int], object]).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self._values = values

        def cell(self, row: int, column: int) -> _Cell:
            """Cell.
            
            Args:
                row: Parameter value (int).
                column: Parameter value (int).
            
            Returns:
                _Cell: Return value.
            
            Raises:
                None.
            """
            return _Cell(row, column, self._values.get((row, column)))

    ws = _Worksheet(
        {
            (3, 4): 10,
            (3, 5): 10,
            (3, 6): 10,
            (3, 7): 10,
            (4, 2): "REG001",
            (4, 3): "Student One",
            (4, 4): "A",
            (4, 5): 5,
            (4, 6): 3,
            (4, 7): 2,
            (4, 8): 123,
        }
    )

    with pytest.raises(ValidationError) as excinfo:
        v2_validator._validate_non_empty_marks_entries(
            worksheet=ws,
            sheet_name="S1",
            sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
            header_count=8,
            header_row=1,
        )

    if not (excinfo.value.code == "MARKS_TEMPLATE_VALIDATION_FAILED"):
        raise AssertionError('assertion failed')
    issues = cast(list[dict[str, object]], excinfo.value.context.get("issues", []))
    if not (any(str(item.get("code", "")).strip() == "COA_ABSENCE_POLICY_VIOLATION" for item in issues)):
        raise AssertionError('assertion failed')
    if not (any(
        str(item.get("code", "")).strip() == "INSTRUCTOR_VALIDATION_STEP2_TOTAL_FORMULA_MISMATCH"
        for item in issues
    )):
        raise AssertionError('assertion failed')


def test_row_total_consistency_accepts_absence_aware_total_formula() -> None:
    """Test row total consistency accepts absence aware total formula.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    class _Cell:
        def __init__(self, row: int, column: int, value: object) -> None:
            """Init.
            
            Args:
                row: Parameter value (int).
                column: Parameter value (int).
                value: Parameter value (object).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self.row = row
            self.column = column
            self.value = value
            self.coordinate = f"{v2_validator._excel_col_name(column)}{row}"

    class _Worksheet:
        def __init__(self, values: dict[tuple[int, int], object]) -> None:
            """Init.
            
            Args:
                values: Parameter value (dict[tuple[int, int], object]).
            
            Returns:
                None.
            
            Raises:
                None.
            """
            self._values = values

        def cell(self, row: int, column: int) -> _Cell:
            """Cell.
            
            Args:
                row: Parameter value (int).
                column: Parameter value (int).
            
            Returns:
                _Cell: Return value.
            
            Raises:
                None.
            """
            return _Cell(row, column, self._values.get((row, column)))

    values: dict[tuple[int, int], object] = {}
    for row in (12, 13, 14):
        values[(row, 8)] = (
            f'=IF(COUNTIF(D{row}:G{row},"A")+COUNTIF(D{row}:G{row},"a")>0,'
            f'"A",SUM(D{row}:G{row}))'
        )
    ws = _Worksheet(values)

    v2_validator._validate_row_total_consistency(
        worksheet=ws,
        sheet_name="S1",
        sheet_kind=LAYOUT_SHEET_KIND_DIRECT_CO_WISE,
        header_count=8,
        header_row=9,
        student_count=3,
    )
