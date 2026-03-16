from __future__ import annotations

import builtins
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from modules.instructor.steps import shared_workbook_ops as ops


def test_sanitize_filename_token_strips_invalid_and_spaces() -> None:
    assert ops.sanitize_filename_token(r'  A<B>: C/ D\E|F?*  ') == 'A_B_C_D_E_F'
    assert ops.sanitize_filename_token(' ..__ ') == ''


def test_build_default_name_fallback_when_no_path(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ops, "t", lambda key, **kwargs: "fallback.xlsx")
    assert ops.build_marks_template_default_name(None) == "fallback.xlsx"


def test_build_default_name_fallback_when_openpyxl_missing(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(ops, "t", lambda key, **kwargs: "fallback.xlsx")
    original_import = builtins.__import__

    def _fake_import(name, globals=None, locals=None, fromlist=(), level=0):
        if name == "openpyxl":
            raise ModuleNotFoundError("openpyxl")
        return original_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _fake_import)
    assert ops.build_final_report_default_name(str(tmp_path / "in.xlsx")) == "fallback.xlsx"


def _write_course_metadata_workbook(path: Path, *, fields: dict[str, object], include_sheet: bool = True) -> None:
    wb = openpyxl.Workbook()
    try:
        if include_sheet:
            ws = wb.active
            ws.title = ops.COURSE_METADATA_SHEET
            ws["A1"] = "Field"
            ws["B1"] = "Value"
            row = 2
            for k, v in fields.items():
                ws.cell(row=row, column=1, value=k)
                ws.cell(row=row, column=2, value=v)
                row += 1
        else:
            wb.active.title = "Other"
        wb.save(path)
    finally:
        wb.close()


def test_build_default_name_fallback_when_metadata_missing(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ops, "t", lambda key, **kwargs: "fallback.xlsx")
    wb_path = tmp_path / "missing_meta.xlsx"
    _write_course_metadata_workbook(wb_path, fields={}, include_sheet=False)

    assert ops.build_marks_template_default_name(str(wb_path)) == "fallback.xlsx"


def test_build_default_name_from_metadata_success_and_sanitization(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ops, "t", lambda key, **kwargs: "fallback.xlsx")
    wb_path = tmp_path / "meta.xlsx"
    _write_course_metadata_workbook(
        wb_path,
        fields={
            ops.COURSE_METADATA_COURSE_CODE_KEY: "ECE<101>",
            ops.COURSE_METADATA_SEMESTER_KEY: "III",
            ops.COURSE_METADATA_SECTION_KEY: "A ",
            ops.COURSE_METADATA_ACADEMIC_YEAR_KEY: "2025/26",
        },
    )

    name = ops.build_marks_template_default_name(str(wb_path))
    assert name.endswith(ops.FILE_EXTENSION_XLSX)
    assert name.startswith("ECE_101_III_A_2025_26")


def test_build_default_name_fallback_on_incomplete_required_fields(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ops, "t", lambda key, **kwargs: "fallback.xlsx")
    wb_path = tmp_path / "incomplete.xlsx"
    _write_course_metadata_workbook(
        wb_path,
        fields={
            ops.COURSE_METADATA_COURSE_CODE_KEY: "ECE101",
            ops.COURSE_METADATA_SEMESTER_KEY: "",
            ops.COURSE_METADATA_SECTION_KEY: "A",
            ops.COURSE_METADATA_ACADEMIC_YEAR_KEY: "2025-26",
        },
    )

    assert ops.build_final_report_default_name(str(wb_path)) == "fallback.xlsx"


def test_atomic_copy_file_success(tmp_path: Path) -> None:
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "nested" / "dst.xlsx"
    src.write_text("data", encoding="utf-8")

    out = ops.atomic_copy_file(src, dst)

    assert out == dst
    assert dst.read_text(encoding="utf-8") == "data"


def test_atomic_copy_file_cleanup_warning_when_unlink_fails(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    src.write_text("data", encoding="utf-8")

    monkeypatch.setattr(ops.shutil, "copyfile", lambda *_args, **_kwargs: (_ for _ in ()).throw(OSError("copy fail")))

    original_unlink = Path.unlink

    def _bad_unlink(self, *args, **kwargs):
        raise OSError("unlink fail")

    monkeypatch.setattr(Path, "unlink", _bad_unlink)

    class _Logger:
        def __init__(self) -> None:
            self.calls: list[tuple] = []

        def warning(self, msg, *args):
            self.calls.append((msg, args))

    logger = _Logger()
    with pytest.raises(OSError, match="copy fail"):
        ops.atomic_copy_file(src, dst, logger=logger)

    assert logger.calls, "expected cleanup warning when temp file unlink fails"
    monkeypatch.setattr(Path, "unlink", original_unlink)
