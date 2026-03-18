from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any, cast

import pytest

from common.jobs import CancellationToken
from modules import coordinator_processing as cp


class _Conn:
    def __init__(self, *, fail_delete: bool = False) -> None:
        self.fail_delete = fail_delete
        self.closed = False

    def execute(self, query: str, *_args):
        if self.fail_delete and query.startswith("DELETE"):
            raise RuntimeError("db err")
        return type("_C", (), {"rowcount": 1})()

    def commit(self) -> None:
        return None

    def close(self) -> None:
        self.closed = True


class _WB:
    def __init__(self, sheetnames: list[str], sheets: dict[str, object]) -> None:
        self.sheetnames = sheetnames
        self._sheets = sheets
        self.closed = False

    def __getitem__(self, key: str):
        return self._sheets[key]

    def close(self):
        self.closed = True


class _Cell:
    def __init__(self, value):
        self.value = value


class _SheetByKey:
    def __init__(self, mapping: dict[str, object]) -> None:
        self.mapping = mapping

    def __getitem__(self, key: str):
        return _Cell(self.mapping.get(key))


class _MetaSheet:
    def __init__(self, rows: list[tuple[object, object]]) -> None:
        self.rows = rows

    def cell(self, *, row: int, column: int):
        idx = row - 2
        if idx < 0 or idx >= len(self.rows):
            return _Cell("")
        pair = self.rows[idx]
        return _Cell(pair[column - 1])


def test_register_dedup_store_memory_mode() -> None:
    store = cp._RegisterDedupStore(total_outcomes=2, use_sqlite=False)
    try:
        assert store.add_if_absent(co_index=1, reg_hash=101) is True
        assert store.add_if_absent(co_index=1, reg_hash=101) is False
        assert store.add_if_absent(co_index=2, reg_hash=101) is True
    finally:
        store.close()


def test_register_dedup_store_close_swallows_delete_and_unlink_errors(monkeypatch) -> None:
    store = cp._RegisterDedupStore(total_outcomes=1, use_sqlite=False)
    store._conn = cast(Any, _Conn(fail_delete=True))
    store._db_path = "C:/tmp/nonexistent.sqlite3"

    monkeypatch.setattr(cp._logger, "debug", lambda *_a, **_k: None)
    monkeypatch.setattr(cp.Path, "unlink", lambda self, missing_ok=False: (_ for _ in ()).throw(OSError("x")))

    store.close()
    assert store._conn is None
    assert store._db_path is None


def test_cleanup_stale_dedup_sqlite_files_ignores_missing_and_unlink_errors(monkeypatch, tmp_path: Path) -> None:
    temp_root = tmp_path / "temp"
    runtime_root = tmp_path / "runtime"
    sqlite_root = runtime_root / "sqlite"
    temp_root.mkdir(parents=True)
    sqlite_root.mkdir(parents=True)

    stale = temp_root / f"{cp._DEDUP_SQLITE_PREFIX}abc{cp._DEDUP_SQLITE_SUFFIX}"
    stale.write_text("x", encoding="utf-8")

    monkeypatch.setattr(cp.tempfile, "gettempdir", lambda: str(temp_root))
    monkeypatch.setattr(cp, "app_runtime_storage_dir", lambda _app: runtime_root)

    orig_unlink = Path.unlink

    def _unlink(self: Path, missing_ok: bool = False):
        if self == stale:
            raise OSError("locked")
        return orig_unlink(self, missing_ok=missing_ok)

    monkeypatch.setattr(cp.Path, "unlink", _unlink)
    monkeypatch.setattr(cp._logger, "debug", lambda *_a, **_k: None)

    cp._cleanup_stale_dedup_sqlite_files()


def test_filter_excel_paths_deduplicates_and_resolves(tmp_path: Path) -> None:
    good = tmp_path / "a.xlsx"
    bad = tmp_path / "b.txt"
    good.write_text("x", encoding="utf-8")
    bad.write_text("x", encoding="utf-8")

    out = cp._filter_excel_paths([str(good), str(bad), str(good)])
    assert out == [good.resolve()]


def test_read_template_id_from_hash_sheet_branches(monkeypatch) -> None:
    wb_missing = _WB([], {})
    assert cp._read_template_id_from_hash_sheet(wb_missing) is None

    wb_bad = _WB(
        [cp.SYSTEM_HASH_SHEET],
        {cp.SYSTEM_HASH_SHEET: _SheetByKey({"A1": "bad", "B1": "bad", "A2": "x", "B2": "y"})},
    )
    assert cp._read_template_id_from_hash_sheet(wb_bad) is None

    wb_sig_bad = _WB(
        [cp.SYSTEM_HASH_SHEET],
        {
            cp.SYSTEM_HASH_SHEET: _SheetByKey(
                {
                    "A1": cp.SYSTEM_HASH_TEMPLATE_ID_HEADER,
                    "B1": cp.SYSTEM_HASH_TEMPLATE_HASH_HEADER,
                    "A2": cp.ID_COURSE_SETUP,
                    "B2": "sig",
                }
            )
        },
    )
    monkeypatch.setattr(cp, "verify_payload_signature", lambda payload, sig: False)
    assert cp._read_template_id_from_hash_sheet(wb_sig_bad) is None


def test_read_report_sheet_counts_invalid_paths(monkeypatch) -> None:
    wb = _WB([], {})
    assert cp._read_report_sheet_counts(wb) is None

    wb2 = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey({"A1": "bad", "B1": "bad"})},
    )
    assert cp._read_report_sheet_counts(wb2) is None

    wb3 = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {
                    "A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
                    "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
                    "A2": "{}",
                    "B2": "sig",
                }
            )
        },
    )
    monkeypatch.setattr(cp, "verify_payload_signature", lambda payload, sig: True)
    assert cp._read_report_sheet_counts(wb3) is None


def test_read_signature_metadata_invalid_and_valid() -> None:
    wb_missing = _WB([], {})
    assert cp._read_signature_metadata(wb_missing) is None

    wb_invalid_total = _WB(
        [cp.COURSE_METADATA_SHEET],
        {
            cp.COURSE_METADATA_SHEET: _MetaSheet(
                [
                    (cp.COURSE_METADATA_COURSE_CODE_KEY, "C101"),
                    (cp.COURSE_METADATA_TOTAL_OUTCOMES_KEY, "NaN"),
                    (cp.COURSE_METADATA_SECTION_KEY, "A"),
                    ("", ""),
                ]
            )
        },
    )
    assert cp._read_signature_metadata(wb_invalid_total) is None

    wb_valid = _WB(
        [cp.COURSE_METADATA_SHEET],
        {
            cp.COURSE_METADATA_SHEET: _MetaSheet(
                [
                    (cp.COURSE_METADATA_COURSE_CODE_KEY, "C101"),
                    (cp.COURSE_METADATA_TOTAL_OUTCOMES_KEY, "6"),
                    (cp.COURSE_METADATA_SECTION_KEY, "A"),
                    ("", ""),
                ]
            )
        },
    )
    assert cp._read_signature_metadata(wb_valid) == ("C101", 6, "A")


def test_generate_co_attainment_workbook_input_guards(monkeypatch, tmp_path: Path) -> None:
    token = CancellationToken()

    with pytest.raises(ValueError, match="No source files"):
        cp._generate_co_attainment_workbook([], tmp_path / "out.xlsx", token=token)

    src = tmp_path / "a.xlsx"
    src.write_text("x", encoding="utf-8")
    monkeypatch.setattr(cp, "_extract_final_report_signature", lambda _p: None)
    with pytest.raises(ValueError, match="Invalid final CO report"):
        cp._generate_co_attainment_workbook([src], tmp_path / "out.xlsx", token=token)


class _FreezeSheet:
    def __init__(self) -> None:
        self.freeze_calls: list[tuple[int, int]] = []

    def write(self, *_args, **_kwargs) -> None:
        return None

    def write_row(self, *_args, **_kwargs) -> None:
        return None

    def set_column(self, *_args, **_kwargs) -> None:
        return None

    def set_landscape(self) -> None:
        return None

    def set_paper(self, *_args, **_kwargs) -> None:
        return None

    def fit_to_pages(self, *_args, **_kwargs) -> None:
        return None

    def repeat_rows(self, *_args, **_kwargs) -> None:
        return None

    def freeze_panes(self, row: int, col: int) -> None:
        self.freeze_calls.append((row, col))

    def protect(self) -> None:
        return None

    def set_selection(self, *_args, **_kwargs) -> None:
        return None


class _FreezeWorkbook:
    def __init__(self) -> None:
        self.sheet = _FreezeSheet()

    def add_worksheet(self, _name: str) -> _FreezeSheet:
        return self.sheet

    def add_format(self, payload: dict[str, object]) -> dict[str, object]:
        return payload


def test_create_co_attainment_sheet_freezes_headers_and_student_columns() -> None:
    workbook = _FreezeWorkbook()

    state = cp._create_co_attainment_sheet(workbook, co_index=1, metadata={})

    assert workbook.sheet.freeze_calls == [(state.header_row_index + 1, 3)]


def test_register_dedup_store_sqlite_fd_close_oserror_is_swallowed(monkeypatch, tmp_path: Path) -> None:
    sqlite_file = tmp_path / "dedup.sqlite3"
    monkeypatch.setattr(cp, "_cleanup_stale_dedup_sqlite_files", lambda: None)
    monkeypatch.setattr(cp, "create_app_runtime_sqlite_file", lambda *_a, **_k: (77, str(sqlite_file)))
    real_connect = sqlite3.connect
    monkeypatch.setattr(cp.sqlite3, "connect", lambda _p: real_connect(":memory:"))

    import os

    monkeypatch.setattr(os, "close", lambda _fd: (_ for _ in ()).throw(OSError("close failed")))
    store = cp._RegisterDedupStore(total_outcomes=1, use_sqlite=True)
    try:
        assert store.add_if_absent(co_index=1, reg_hash=1) is True
    finally:
        store.close()


def test_cleanup_stale_dedup_sqlite_files_skips_missing_root(monkeypatch, tmp_path: Path) -> None:
    temp_root = tmp_path / "temp"
    sqlite_root = tmp_path / "runtime" / "sqlite"
    temp_root.mkdir(parents=True)
    monkeypatch.setattr(cp.tempfile, "gettempdir", lambda: str(temp_root))
    monkeypatch.setattr(cp, "app_runtime_storage_dir", lambda _app: tmp_path / "runtime")
    # sqlite_root intentionally not created to exercise "if not root.exists(): continue"
    cp._cleanup_stale_dedup_sqlite_files()


def test_read_template_id_and_report_sheet_counts_more_invalid_branches(monkeypatch) -> None:
    wb_header_b1_bad = _WB(
        [cp.SYSTEM_HASH_SHEET],
        {
            cp.SYSTEM_HASH_SHEET: _SheetByKey(
                {"A1": cp.SYSTEM_HASH_TEMPLATE_ID_HEADER, "B1": "wrong", "A2": cp.ID_COURSE_SETUP, "B2": "sig"}
            )
        },
    )
    assert cp._read_template_id_from_hash_sheet(wb_header_b1_bad) is None

    wb_missing_payload = _WB(
        [cp.SYSTEM_HASH_SHEET],
        {
            cp.SYSTEM_HASH_SHEET: _SheetByKey(
                {"A1": cp.SYSTEM_HASH_TEMPLATE_ID_HEADER, "B1": cp.SYSTEM_HASH_TEMPLATE_HASH_HEADER, "A2": "", "B2": ""}
            )
        },
    )
    assert cp._read_template_id_from_hash_sheet(wb_missing_payload) is None

    wb_bad_b1 = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey({"A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER, "B1": "bad"})},
    )
    assert cp._read_report_sheet_counts(wb_bad_b1) is None

    wb_empty_manifest = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {"A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER, "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER, "A2": "", "B2": ""}
            )
        },
    )
    assert cp._read_report_sheet_counts(wb_empty_manifest) is None

    monkeypatch.setattr(cp, "verify_payload_signature", lambda *_a, **_k: False)
    wb_sig_invalid = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {"A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER, "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER, "A2": "{}", "B2": "x"}
            )
        },
    )
    assert cp._read_report_sheet_counts(wb_sig_invalid) is None

    monkeypatch.setattr(cp, "verify_payload_signature", lambda *_a, **_k: True)
    wb_manifest_not_dict = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {"A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER, "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER, "A2": "[]", "B2": "x"}
            )
        },
    )
    assert cp._read_report_sheet_counts(wb_manifest_not_dict) is None

    wb_sheet_order_not_list = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {
                    "A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
                    "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
                    "A2": '{"sheet_order":"bad"}',
                    "B2": "x",
                }
            )
        },
    )
    assert cp._read_report_sheet_counts(wb_sheet_order_not_list) is None

    wb_sheet_order_empty = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {
                    "A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
                    "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
                    "A2": '{"sheet_order":[" ", 1]}',
                    "B2": "x",
                }
            )
        },
    )
    assert cp._read_report_sheet_counts(wb_sheet_order_empty) is None

    wb_sheet_order_non_str_only = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {
                    "A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
                    "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
                    "A2": '{"sheet_order":[1,2]}',
                    "B2": "x",
                }
            )
        },
    )
    assert cp._read_report_sheet_counts(wb_sheet_order_non_str_only) is None

    wb_mismatch_counts = _WB(
        [cp.SYSTEM_REPORT_INTEGRITY_SHEET],
        {
            cp.SYSTEM_REPORT_INTEGRITY_SHEET: _SheetByKey(
                {
                    "A1": cp.SYSTEM_REPORT_INTEGRITY_MANIFEST_HEADER,
                    "B1": cp.SYSTEM_REPORT_INTEGRITY_HASH_HEADER,
                    "A2": '{"sheet_order":["CO1_Direct","CO2_Direct","CO1_Indirect"]}',
                    "B2": "x",
                }
            )
        },
    )
    assert cp._read_report_sheet_counts(wb_mismatch_counts) is None


def test_read_signature_metadata_additional_invalid_branches() -> None:
    wb_empty_immediately = _WB([cp.COURSE_METADATA_SHEET], {cp.COURSE_METADATA_SHEET: _MetaSheet([("", "")])})
    assert cp._read_signature_metadata(wb_empty_immediately) is None

    wb_missing_section = _WB(
        [cp.COURSE_METADATA_SHEET],
        {cp.COURSE_METADATA_SHEET: _MetaSheet([(cp.COURSE_METADATA_COURSE_CODE_KEY, "C101"), ("", "")])},
    )
    assert cp._read_signature_metadata(wb_missing_section) is None

    wb_zero_total = _WB(
        [cp.COURSE_METADATA_SHEET],
        {
            cp.COURSE_METADATA_SHEET: _MetaSheet(
                [
                    (cp.COURSE_METADATA_COURSE_CODE_KEY, "C101"),
                    (cp.COURSE_METADATA_TOTAL_OUTCOMES_KEY, "0"),
                    (cp.COURSE_METADATA_SECTION_KEY, "A"),
                    ("", ""),
                ]
            )
        },
    )
    assert cp._read_signature_metadata(wb_zero_total) is None


def test_extract_final_report_signature_xls_and_missing_openpyxl(monkeypatch, tmp_path: Path) -> None:
    xls = tmp_path / "a.xls"
    xls.write_text("x", encoding="utf-8")
    assert cp._extract_final_report_signature(xls) is None

    src = tmp_path / "a.xlsx"
    src.write_text("x", encoding="utf-8")

    import builtins

    real_import = builtins.__import__

    def _fake_import(name, globals=None, locals=None, fromlist=(), level=0):  # type: ignore[no-untyped-def]
        if name == "openpyxl":
            raise RuntimeError("missing")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _fake_import)
    assert cp._extract_final_report_signature(src) is None


def test_extract_final_report_signature_metadata_none_branch(monkeypatch, tmp_path: Path) -> None:
    src = tmp_path / "a.xlsx"
    src.write_text("x", encoding="utf-8")

    class _LoadWb:
        sheetnames = []

        def close(self) -> None:
            return None

    monkeypatch.setattr(cp, "_read_template_id_from_hash_sheet", lambda _w: cp.ID_COURSE_SETUP)
    monkeypatch.setattr(cp, "_read_report_sheet_counts", lambda _w: (1, 1))
    monkeypatch.setattr(cp, "_read_signature_metadata", lambda _w: None)

    import types
    import sys

    mod = types.SimpleNamespace(load_workbook=lambda **_k: _LoadWb())
    monkeypatch.setitem(sys.modules, "openpyxl", mod)
    assert cp._extract_final_report_signature(src) is None


def test_analyze_dropped_files_duplicate_and_invalid_existing(monkeypatch, tmp_path: Path) -> None:
    valid = tmp_path / "valid.xlsx"
    invalid = tmp_path / "invalid.xlsx"
    valid.write_text("x", encoding="utf-8")
    invalid.write_text("x", encoding="utf-8")

    sig = cp._FinalReportSignature(
        template_id="id",
        course_code="C101",
        total_outcomes=3,
        section="A",
        direct_sheet_count=3,
        indirect_sheet_count=3,
    )
    monkeypatch.setattr(
        cp,
        "_extract_final_report_signature",
        lambda p: None if "invalid" in str(p) else sig,
    )
    out = cp._analyze_dropped_files(
        [str(valid), str(valid), str(invalid)],
        existing_keys={cp._path_key(valid)},
        existing_paths=[str(invalid)],
        token=CancellationToken(),
    )
    assert out["duplicates"] == 1
    assert out["added"] == []
    assert len(cast(list[str], out["invalid_final_report"])) == 1


class _IterSheet:
    def __init__(self, *, title: str, max_row: int, max_column: int, rows: dict[int, tuple[Any, ...]]) -> None:
        self.title = title
        self.max_row = max_row
        self.max_column = max_column
        self._rows = rows

    def iter_rows(
        self,
        *,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
        values_only: bool = True,
    ):
        del min_col, values_only
        for row_idx in range(min_row, max_row + 1):
            row = self._rows.get(row_idx, tuple())
            if len(row) < max_col:
                row = tuple(row) + ("",) * (max_col - len(row))
            yield row[:max_col]


class _IterSheetNoPad(_IterSheet):
    def iter_rows(
        self,
        *,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
        values_only: bool = True,
    ):
        del min_col, max_col, values_only
        for row_idx in range(min_row, max_row + 1):
            yield self._rows.get(row_idx, tuple())


def test_iter_score_rows_handles_second_scan_and_row_skips() -> None:
    header = (
        cp.CO_REPORT_HEADER_SERIAL,
        cp.CO_REPORT_HEADER_REG_NO,
        cp.CO_REPORT_HEADER_STUDENT_NAME,
        cp._ratio_total_header(cp.DIRECT_RATIO),
    )
    sheet = _IterSheet(
        title="CO1_Direct",
        max_row=230,
        max_column=4,
        rows={
            220: header,
            221: ("1",),  # short row
            222: ("2", "", "Alice", "40"),  # empty reg
            223: ("3", "R1", "Alice", "40"),
            224: ("4", "R1", "Dup", "41"),  # duplicate reg
            225: ("5", "R2", "Bob", "bad"),  # invalid score
        },
    )
    rows = list(cp._iter_score_rows(sheet, ratio=cp.DIRECT_RATIO))
    assert len(rows) == 1
    assert rows[0].reg_no == "R1"


def test_iter_score_rows_short_data_row_continue_branch() -> None:
    header = (
        cp.CO_REPORT_HEADER_SERIAL,
        cp.CO_REPORT_HEADER_REG_NO,
        cp.CO_REPORT_HEADER_STUDENT_NAME,
        cp._ratio_total_header(cp.DIRECT_RATIO),
    )
    sheet = _IterSheetNoPad(
        title="CO1_Direct",
        max_row=6,
        max_column=4,
        rows={
            1: header,
            2: ("1",),  # len(values) short -> continue branch
            3: ("2", "R1", "Alice", "40"),
        },
    )
    rows = list(cp._iter_score_rows(sheet, ratio=cp.DIRECT_RATIO))
    assert len(rows) == 1 and rows[0].reg_no == "R1"


def test_iter_score_rows_missing_headers_raises() -> None:
    sheet = _IterSheet(title="NoHeaders", max_row=5, max_column=3, rows={1: ("A", "B", "C")})
    with pytest.raises(ValueError, match="Required headers are missing"):
        list(cp._iter_score_rows(sheet, ratio=cp.DIRECT_RATIO))


def test_iter_co_rows_from_workbook_missing_and_non_lockstep(monkeypatch) -> None:
    workbook_missing = type("_WB", (), {"sheetnames": [], "__getitem__": lambda self, key: None})()
    with pytest.raises(ValueError, match="Missing CO sheets"):
        list(cp._iter_co_rows_from_workbook(workbook_missing, co_index=1, workbook_name="a.xlsx"))

    direct_name = f"CO1{cp.CO_REPORT_DIRECT_SHEET_SUFFIX}"
    indirect_name = f"CO1{cp.CO_REPORT_INDIRECT_SHEET_SUFFIX}"
    workbook = type(
        "_WB2",
        (),
        {
            "sheetnames": [direct_name, indirect_name],
            "__getitem__": lambda self, key: key,
        },
    )()

    direct_rows = [
        cp._ParsedScoreRow(reg_hash=1, reg_key="r1", reg_no="R1", student_name="A", score=30.0),
        cp._ParsedScoreRow(reg_hash=2, reg_key="r2", reg_no="R2", student_name="B", score=40.0),
    ]
    indirect_rows = [
        cp._ParsedScoreRow(reg_hash=99, reg_key="r99", reg_no="R99", student_name="X", score=20.0),
        cp._ParsedScoreRow(reg_hash=2, reg_key="r2", reg_no="R2", student_name="B2", score=10.0),
    ]

    def _fake_iter_score_rows(sheet: Any, *, ratio: float):
        del ratio
        if sheet == direct_name:
            yield from direct_rows
        else:
            yield from indirect_rows

    monkeypatch.setattr(cp, "_iter_score_rows", _fake_iter_score_rows)
    out = list(cp._iter_co_rows_from_workbook(workbook, co_index=1, workbook_name="x.xlsx"))
    assert len(out) == 1
    assert out[0].reg_no == "R2"


def test_small_helper_branches_and_thresholds(monkeypatch, tmp_path: Path) -> None:
    assert cp._ratio_percent_token(0.335) == "33.5"
    assert cp._coerce_numeric_score(True) is None
    assert cp._coerce_numeric_score("not-a-number") is None
    assert cp._co_percentage(level_2=1, level_3=1, attended=0) == cp.CO_REPORT_NOT_APPLICABLE_TOKEN
    assert cp._attainment_thresholds((10, 20, 30)) == (10.0, 20.0, 30.0)

    token = CancellationToken()
    src = tmp_path / "src.xlsx"
    src.write_text("x", encoding="utf-8")
    bad_sig = cp._FinalReportSignature(
        template_id="unsupported",
        course_code="C101",
        total_outcomes=3,
        section="A",
        direct_sheet_count=3,
        indirect_sheet_count=3,
    )
    monkeypatch.setattr(cp, "_extract_final_report_signature", lambda _p: bad_sig)
    with pytest.raises(ValueError, match="Unsupported template id"):
        cp._generate_co_attainment_workbook([src], tmp_path / "out.xlsx", token=token)


def test_generate_co_attainment_workbook_course_setup_v1_remaining_branches(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    src = tmp_path / "in.xlsx"
    src.write_text("x", encoding="utf-8")

    with pytest.raises(ValueError, match="No CO outcomes"):
        cp._generate_co_attainment_workbook_course_setup_v1(
            [src],
            tmp_path / "out.xlsx",
            token=CancellationToken(),
            total_outcomes=0,
            template_id=cp.ID_COURSE_SETUP,
        )

    class _InputWbNoMeta:
        sheetnames = ["X"]

        def __getitem__(self, _k: str) -> object:
            return object()

        def close(self) -> None:
            return None

    class _OutWb:
        def __init__(self, _path: str, _opts: dict[str, object]) -> None:
            self.close_calls = 0

        def close(self) -> None:
            self.close_calls += 1

    monkeypatch.setattr(cp, "_RegisterDedupStore", lambda **_k: type("_D", (), {"add_if_absent": lambda *_a, **_b: True, "close": lambda *_a: None})())
    monkeypatch.setattr(cp, "_iter_co_rows_from_workbook", lambda *_a, **_k: [])
    monkeypatch.setattr(cp, "_create_co_attainment_sheet", lambda *_a, **_k: cp._CoOutputSheetState(sheet=object(), header_row_index=0, formats={}, next_row_index=0, next_serial=1, on_roll=0, attended=0, level_counts={0: 0, 1: 0, 2: 0, 3: 0}))
    monkeypatch.setattr(cp, "_append_co_attainment_summary", lambda *_a, **_k: None)
    monkeypatch.setattr(cp, "_create_summary_sheet", lambda *_a, **_k: (1, 1))
    monkeypatch.setattr(cp, "_create_graph_sheet", lambda *_a, **_k: None)
    monkeypatch.setattr(cp, "_write_system_integrity_sheets", lambda *_a, **_k: None)

    import sys
    import types

    monkeypatch.setitem(sys.modules, "openpyxl", types.SimpleNamespace(load_workbook=lambda **_k: _InputWbNoMeta()))
    monkeypatch.setitem(sys.modules, "xlsxwriter", types.SimpleNamespace(Workbook=_OutWb))
    result = cp._generate_co_attainment_workbook_course_setup_v1(
        [src],
        tmp_path / "out2.xlsx",
        token=CancellationToken(),
        total_outcomes=1,
        template_id=cp.ID_COURSE_SETUP,
    )
    assert result.output_path == (tmp_path / "out2.xlsx")

    # Cleanup close-exception debug branch when workbook is not marked closed.
    class _OutWbCloseFails(_OutWb):
        def close(self) -> None:
            self.close_calls += 1
            raise RuntimeError("close failed")

    monkeypatch.setitem(sys.modules, "xlsxwriter", types.SimpleNamespace(Workbook=_OutWbCloseFails))
    monkeypatch.setattr(cp, "_create_summary_sheet", lambda *_a, **_k: (_ for _ in ()).throw(RuntimeError("boom")))
    monkeypatch.setattr(cp._logger, "debug", lambda *_a, **_k: None)
    with pytest.raises(RuntimeError):
        cp._generate_co_attainment_workbook_course_setup_v1(
            [src],
            tmp_path / "out3.xlsx",
            token=CancellationToken(),
            total_outcomes=1,
            template_id=cp.ID_COURSE_SETUP,
        )
