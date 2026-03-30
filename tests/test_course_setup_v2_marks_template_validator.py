from __future__ import annotations

import sys
from dataclasses import dataclass
from types import SimpleNamespace
from typing import Any, cast

import pytest

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
    assert len(warnings) == 2
    assert any("High absence ratio" in message for message in warnings)
    assert any("Near-constant marks" in message for message in warnings)
    assert v2_validator.consume_last_marks_anomaly_warnings() == []


def test_v2_marks_validator_resets_warning_buffer_per_run() -> None:
    v2_validator._last_marks_anomaly_warnings[:] = ["stale warning"]

    class _Workbook:
        sheetnames = ["OnlySheet"]

    try:
        v2_validator.validate_filled_marks_manifest_schema(workbook=_Workbook(), manifest={"sheet_order": [], "sheets": []})
    except Exception:
        pass

    # stale warnings should be cleared at validation start
    assert "stale warning" not in v2_validator.consume_last_marks_anomaly_warnings()


def test_v2_marks_batch_validator_returns_course_like_shape(monkeypatch: pytest.MonkeyPatch) -> None:
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
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

    assert result["valid_paths"] == ["valid.xlsx"]
    assert result["invalid_paths"] == ["invalid.xlsx"]
    assert result["duplicate_paths"] == ["valid.xlsx"]
    assert result["mismatched_paths"] == []
    assert result["duplicate_sections"] == []
    assert isinstance(result["template_ids"], dict)
    assert result["rejections"]


def test_v2_marks_batch_validator_tracks_template_mismatch(monkeypatch: pytest.MonkeyPatch) -> None:
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
        del workbook_path
        del expected_template_id
        del cancel_token
        raise ValidationError("template mismatch", code="UNKNOWN_TEMPLATE", context={"template_id": "COURSE_SETUP_V1"})

    monkeypatch.setattr(v2_validator, "_validate_filled_marks_workbook_impl", _fake_validate)
    result = cast(
        dict[str, Any],
        v2_validator.validate_filled_marks_workbooks(
            ["wrong_template.xlsx"],
            template_id="COURSE_SETUP_V2",
        ),
    )

    assert result["valid_paths"] == []
    assert result["invalid_paths"] == ["wrong_template.xlsx"]
    assert result["mismatched_paths"] == ["wrong_template.xlsx"]
    rejection = cast(list[dict[str, Any]], result["rejections"])[0]
    assert rejection["reason_kind"] == "template_mismatch"


@dataclass(frozen=True)
class _Payload:
    template_id: str
    manifest: dict[str, object]


def test_marks_impl_two_stage_trust_flow_uses_read_only_then_full_workbook(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    workbook_path = tmp_path / "marks.xlsx"
    workbook_path.write_text("x")

    calls: list[dict[str, object]] = []
    workbooks: list[Any] = []

    class _Workbook:
        def __init__(self, mode: str) -> None:
            self.mode = mode
            self.closed = False
            self.sheetnames = ["Sheet1"]

        def close(self) -> None:
            self.closed = True

    def _load_workbook(_path, **kwargs):
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

    assert getattr(resolved, "template_id", "") == "COURSE_SETUP_V2"
    assert calls[0].get("read_only") is True
    assert calls[1].get("read_only", False) is False
    assert workbooks[0].closed is True
    assert workbooks[1].closed is True
    assert len(manifest_calls) == 1
    assert manifest_calls[0][0] is workbooks[1]


def test_marks_impl_manifest_validation_error_is_preserved(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
) -> None:
    workbook_path = tmp_path / "marks.xlsx"
    workbook_path.write_text("x")

    class _Workbook:
        sheetnames = ["Sheet1"]

        def close(self) -> None:
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
    assert excinfo.value.code == "COA_MARK_ENTRY_EMPTY"


def test_v2_marks_batch_validator_tracks_cohort_mismatch(monkeypatch: pytest.MonkeyPatch) -> None:
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
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

    assert result["valid_paths"] == ["ok.xlsx"]
    assert result["mismatched_paths"] == ["cohort_bad.xlsx"]
    assert result["invalid_paths"] == []
    rejection = next(
        item
        for item in cast(list[dict[str, Any]], result["rejections"])
        if item["path"] == "cohort_bad.xlsx"
    )
    assert rejection["reason_kind"] == "cohort_mismatch"
    assert rejection["issue"]["code"] == "MARKS_TEMPLATE_COHORT_MISMATCH"


def test_v2_marks_batch_validator_tracks_duplicate_section(monkeypatch: pytest.MonkeyPatch) -> None:
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
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

    assert result["valid_paths"] == ["first.xlsx"]
    assert result["duplicate_sections"] == ["dup_section.xlsx"]
    assert result["invalid_paths"] == []
    rejection = next(
        item
        for item in cast(list[dict[str, Any]], result["rejections"])
        if item["path"] == "dup_section.xlsx"
    )
    assert rejection["reason_kind"] == "duplicate_section"
    assert rejection["issue"]["code"] == "MARKS_TEMPLATE_SECTION_DUPLICATE"


def test_v2_marks_batch_validator_rejects_cross_workbook_duplicate_reg_no(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token: object | None = None) -> object:
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

    assert result["valid_paths"] == ["first.xlsx"]
    assert result["invalid_paths"] == ["second.xlsx"]
    rejection = next(
        item
        for item in cast(list[dict[str, Any]], result["rejections"])
        if item["path"] == "second.xlsx"
    )
    assert rejection["reason_kind"] == "duplicate_reg_no"
    assert rejection["issue"]["code"] == "MARKS_TEMPLATE_STUDENT_REG_DUPLICATE"
