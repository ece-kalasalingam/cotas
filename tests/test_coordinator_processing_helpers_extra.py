from __future__ import annotations

from pathlib import Path

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
    store._conn = _Conn(fail_delete=True)
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
