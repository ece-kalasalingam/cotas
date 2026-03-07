from __future__ import annotations

import sys
from pathlib import Path

import pytest

from common.exceptions import AppSystemError, ValidationError
from common.sheet_schema import SheetSchema, ValidationRule, WorkbookBlueprint
from modules.instructor import course_details_template_generator as mod
from modules.instructor.course_details_template_generator import (
    generate_course_details_template,
)


class _FakeWorksheet:
    def __init__(self) -> None:
        self.validations: list[dict] = []
        self.protect_calls: list[tuple] = []
        self.writes: list[tuple[int, int, object]] = []
        self.set_columns: list[tuple[int, int, int]] = []
        self.freeze_calls: list[tuple[int, int]] = []
        self.hidden = False

    def write(self, row, col, value, _fmt=None) -> None:
        self.writes.append((row, col, value))

    def write_row(self, row, col, values, _fmt=None) -> None:
        for index, value in enumerate(values):
            self.writes.append((row, col + index, value))

    def set_column(self, first_col, last_col, width) -> None:
        self.set_columns.append((first_col, last_col, width))

    def freeze_panes(self, row, col) -> None:
        self.freeze_calls.append((row, col))

    def data_validation(self, _r1, _c1, _r2, _c2, options) -> None:
        self.validations.append(dict(options))

    def protect(self, *args, **kwargs) -> None:
        self.protect_calls.append((args, kwargs))

    def hide(self) -> None:
        self.hidden = True


class _FakeWorkbook:
    def __init__(self, path: str, _options: dict | None = None) -> None:
        self.path = Path(path)
        self.worksheets: list[_FakeWorksheet] = []
        self.sheet_names: list[str] = []
        self.formats: list[dict] = []

    def add_worksheet(self, name: str) -> _FakeWorksheet:
        ws = _FakeWorksheet()
        self.worksheets.append(ws)
        self.sheet_names.append(name)
        return ws

    def add_format(self, fmt: dict) -> dict:
        self.formats.append(dict(fmt))
        return fmt

    def close(self) -> None:
        self.path.write_text("generated", encoding="utf-8")


class _FailingWorkbook(_FakeWorkbook):
    def close(self) -> None:
        self.path.write_text("partial", encoding="utf-8")
        raise RuntimeError("close failed")


def _install_fake_xlsxwriter(monkeypatch: pytest.MonkeyPatch, workbook_cls: type) -> None:
    class _FakeModule:
        Workbook = workbook_cls

    monkeypatch.setitem(sys.modules, "xlsxwriter", _FakeModule())


def test_unknown_template_id_raises_validation_error(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _install_fake_xlsxwriter(monkeypatch, _FakeWorkbook)

    with pytest.raises(ValidationError, match="Unknown workbook template"):
        generate_course_details_template(tmp_path / "out.xlsx", template_id="UNKNOWN_TEMPLATE")


def test_atomic_replace_success(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _install_fake_xlsxwriter(monkeypatch, _FakeWorkbook)
    output = tmp_path / "course_setup.xlsx"
    output.write_text("old-content", encoding="utf-8")

    result = generate_course_details_template(output)

    assert result == output
    assert output.read_text(encoding="utf-8") == "generated"
    assert list(tmp_path.glob("course_setup.xlsx.*.tmp")) == []


def test_temp_file_is_cleaned_on_failure(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _install_fake_xlsxwriter(monkeypatch, _FailingWorkbook)
    output = tmp_path / "course_setup.xlsx"

    with pytest.raises(AppSystemError, match="Failed to generate course details template"):
        generate_course_details_template(output)

    assert not output.exists()
    assert list(tmp_path.glob("course_setup.xlsx.*.tmp")) == []


def test_missing_xlsxwriter_dependency_raises(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.delitem(sys.modules, "xlsxwriter", raising=False)

    import builtins

    real_import = builtins.__import__

    def _import(name, *args, **kwargs):
        if name == "xlsxwriter":
            raise ModuleNotFoundError("No module named 'xlsxwriter'")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _import)

    with pytest.raises(ValidationError, match="xlsxwriter is not installed"):
        generate_course_details_template(tmp_path / "out.xlsx")


def test_replace_failure_cleans_temp_and_keeps_existing_output(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    _install_fake_xlsxwriter(monkeypatch, _FakeWorkbook)
    output = tmp_path / "course_setup.xlsx"
    output.write_text("existing", encoding="utf-8")

    def _replace_fail(_src, _dst):
        raise PermissionError("target locked")

    monkeypatch.setattr(mod.os, "replace", _replace_fail)

    with pytest.raises(AppSystemError, match="Failed to generate course details template"):
        generate_course_details_template(output)

    assert output.read_text(encoding="utf-8") == "existing"
    assert list(tmp_path.glob("course_setup.xlsx.*.tmp")) == []


def test_validation_default_ignore_blank_true(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    created: dict[str, _FakeWorkbook] = {}

    class _TrackingWorkbook(_FakeWorkbook):
        def __init__(self, path: str, options: dict | None = None) -> None:
            super().__init__(path, options)
            created["wb"] = self

    _install_fake_xlsxwriter(monkeypatch, _TrackingWorkbook)
    generate_course_details_template(tmp_path / "course_setup.xlsx")

    assessment_ws = created["wb"].worksheets[1]
    assert assessment_ws.validations
    assert assessment_ws.validations[0]["ignore_blank"] is True


def test_sheet_order_and_count_match_blueprint(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    created: dict[str, _FakeWorkbook] = {}

    class _TrackingWorkbook(_FakeWorkbook):
        def __init__(self, path: str, options: dict | None = None) -> None:
            super().__init__(path, options)
            created["wb"] = self

    _install_fake_xlsxwriter(monkeypatch, _TrackingWorkbook)
    generate_course_details_template(tmp_path / "course_setup.xlsx")

    assert created["wb"].sheet_names == [
        "Course_Metadata",
        "Assessment_Config",
        "Question_Map",
        "Students",
        "__SYSTEM_HASH__",
    ]
    assert len(created["wb"].worksheets) == 5


def test_course_metadata_header_cells_written(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    created: dict[str, _FakeWorkbook] = {}

    class _TrackingWorkbook(_FakeWorkbook):
        def __init__(self, path: str, options: dict | None = None) -> None:
            super().__init__(path, options)
            created["wb"] = self

    _install_fake_xlsxwriter(monkeypatch, _TrackingWorkbook)
    generate_course_details_template(tmp_path / "course_setup.xlsx")

    first_ws = created["wb"].worksheets[0]
    assert first_ws.writes[:2] == [(0, 0, "Field"), (0, 1, "Value")]


def test_headers_set_freeze_panes_after_header_row(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    created: dict[str, _FakeWorkbook] = {}

    class _TrackingWorkbook(_FakeWorkbook):
        def __init__(self, path: str, options: dict | None = None) -> None:
            super().__init__(path, options)
            created["wb"] = self

    _install_fake_xlsxwriter(monkeypatch, _TrackingWorkbook)
    generate_course_details_template(tmp_path / "course_setup.xlsx")

    for ws in created["wb"].worksheets[:-1]:
        assert ws.freeze_calls == [(1, 0)]


def test_system_hash_sheet_contains_template_id_and_hash(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    created: dict[str, _FakeWorkbook] = {}

    class _TrackingWorkbook(_FakeWorkbook):
        def __init__(self, path: str, options: dict | None = None) -> None:
            super().__init__(path, options)
            created["wb"] = self

    _install_fake_xlsxwriter(monkeypatch, _TrackingWorkbook)
    generate_course_details_template(tmp_path / "course_setup.xlsx")

    hash_ws = created["wb"].worksheets[-1]
    assert hash_ws.writes[:4] == [
        (0, 0, "Template_ID"),
        (0, 1, "Template_Hash"),
        (1, 0, mod.ID_COURSE_SETUP),
        (1, 1, mod._compute_template_hash(mod.ID_COURSE_SETUP)),
    ]
    assert hash_ws.hidden is True


def test_apply_validation_skips_when_validate_not_set() -> None:
    ws = _FakeWorksheet()
    rule = ValidationRule(
        first_row=1,
        first_col=1,
        last_row=2,
        last_col=2,
        options={"source": ["YES", "NO"]},
    )

    mod._apply_validation(ws, rule)

    assert ws.validations == []


def test_apply_validation_respects_explicit_ignore_blank_false() -> None:
    ws = _FakeWorksheet()
    rule = ValidationRule(
        first_row=1,
        first_col=1,
        last_row=2,
        last_col=2,
        options={"validate": "list", "source": ["A", "B"], "ignore_blank": False},
    )

    mod._apply_validation(ws, rule)

    assert ws.validations[0]["ignore_blank"] is False


def test_protected_sheet_triggers_protect_call(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    created: dict[str, _FakeWorkbook] = {}

    class _TrackingWorkbook(_FakeWorkbook):
        def __init__(self, path: str, options: dict | None = None) -> None:
            super().__init__(path, options)
            created["wb"] = self

    _install_fake_xlsxwriter(monkeypatch, _TrackingWorkbook)
    test_template_id = "TEST_PROTECTED_TEMPLATE"
    mod.BLUEPRINT_REGISTRY[test_template_id] = WorkbookBlueprint(
        type_id=test_template_id,
        style_registry={},
        sheets=[SheetSchema(name="Protected", header_matrix=[["H1"]], is_protected=True)],
    )
    try:
        generate_course_details_template(tmp_path / "protected.xlsx", template_id=test_template_id)
    finally:
        del mod.BLUEPRINT_REGISTRY[test_template_id]

    ws = created["wb"].worksheets[0]
    assert len(ws.protect_calls) == 1


def test_sample_data_rows_are_written_below_header(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    created: dict[str, _FakeWorkbook] = {}

    class _TrackingWorkbook(_FakeWorkbook):
        def __init__(self, path: str, options: dict | None = None) -> None:
            super().__init__(path, options)
            created["wb"] = self

    _install_fake_xlsxwriter(monkeypatch, _TrackingWorkbook)
    generate_course_details_template(tmp_path / "course_setup.xlsx")

    first_ws = created["wb"].worksheets[0]
    assert (1, 0, "Course_Code") in first_ws.writes
    assert (1, 1, "ECE000") in first_ws.writes
