from __future__ import annotations

import pytest

from domain import template_strategy_router as router


class _DummyWorkbook:
    def __init__(self, sheetnames: list[str]) -> None:
        self.sheetnames = sheetnames

    def close(self) -> None:
        return None


class _DummyOpenpyxl:
    def __init__(self, sheetnames: list[str]) -> None:
        self._sheetnames = sheetnames

    def load_workbook(self, _source: object, data_only: bool = False, read_only: bool = True) -> _DummyWorkbook:
        del data_only
        del read_only
        return _DummyWorkbook(list(self._sheetnames))


def _patch_structure_probe(
    monkeypatch: pytest.MonkeyPatch,
    *,
    sheetnames: list[str],
    metadata_sheet: str = "Course_Metadata",
    co_description_sheet: str = "CO_Description",
) -> None:
    monkeypatch.setattr(router, "assert_not_symlink_path", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(router, "import_runtime_dependency", lambda _name: _DummyOpenpyxl(sheetnames))
    monkeypatch.setattr(
        router,
        "get_sheet_name_by_key",
        lambda _template_id, sheet_key: (
            metadata_sheet
            if sheet_key == "course_metadata"
            else co_description_sheet
        ),
    )


def test_classify_workbook_structure_marks_template(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_structure_probe(monkeypatch, sheetnames=["__SYSTEM_LAYOUT__", "Other"])
    kind = router.classify_workbook_structure_for_validation(
        template_id="COURSE_SETUP_V2",
        workbook_path="marks.xlsx",
    )
    assert kind == "marks_template"


def test_classify_workbook_structure_co_description(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_structure_probe(monkeypatch, sheetnames=["Course_Metadata", "CO_Description"])
    kind = router.classify_workbook_structure_for_validation(
        template_id="COURSE_SETUP_V2",
        workbook_path="co_description.xlsx",
    )
    assert kind == "co_description"


def test_classify_workbook_structure_ambiguous(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_structure_probe(
        monkeypatch,
        sheetnames=["__SYSTEM_LAYOUT__", "Course_Metadata", "CO_Description"],
    )
    kind = router.classify_workbook_structure_for_validation(
        template_id="COURSE_SETUP_V2",
        workbook_path="mixed.xlsx",
    )
    assert kind == "ambiguous"


def test_classify_workbook_structure_unknown(monkeypatch: pytest.MonkeyPatch) -> None:
    _patch_structure_probe(monkeypatch, sheetnames=["Sheet1", "Sheet2"])
    kind = router.classify_workbook_structure_for_validation(
        template_id="COURSE_SETUP_V2",
        workbook_path="unknown.xlsx",
    )
    assert kind == "unknown"


def test_select_preferred_validation_rejection_prefers_secondary_for_co_description(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        router,
        "classify_workbook_structure_for_validation",
        lambda **_kwargs: "co_description",
    )
    marks_result = {
        "rejections": [
            {"path": "f.xlsx", "issue": {"code": "COA_LAYOUT_SHEET_MISSING", "message": "layout missing"}}
        ]
    }
    co_desc_result = {
        "rejections": [
            {"path": "f.xlsx", "issue": {"code": "CO_DESCRIPTION_SUMMARY_REQUIRED", "message": "summary required"}}
        ]
    }
    selected = router.select_preferred_validation_rejection(
        template_id="COURSE_SETUP_V2",
        workbook_path="f.xlsx",
        primary_kind="marks_template",
        secondary_kind="co_description",
        primary_result=marks_result,
        secondary_result=co_desc_result,
    )
    assert selected is not None
    assert selected["issue"]["code"] == "CO_DESCRIPTION_SUMMARY_REQUIRED"


def test_select_preferred_validation_rejection_unknown_uses_fallback_policy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        router,
        "classify_workbook_structure_for_validation",
        lambda **_kwargs: "unknown",
    )
    marks_result = {
        "rejections": [
            {"path": "f.xlsx", "issue": {"code": "COA_LAYOUT_SHEET_MISSING", "message": "layout missing"}}
        ]
    }
    co_desc_result = {
        "rejections": [
            {"path": "f.xlsx", "issue": {"code": "CO_DESCRIPTION_SUMMARY_REQUIRED", "message": "summary required"}}
        ]
    }
    selected = router.select_preferred_validation_rejection(
        template_id="COURSE_SETUP_V2",
        workbook_path="f.xlsx",
        primary_kind="marks_template",
        secondary_kind="co_description",
        primary_result=marks_result,
        secondary_result=co_desc_result,
    )
    assert selected is not None
    assert selected["issue"]["code"] == "CO_DESCRIPTION_SUMMARY_REQUIRED"

