from __future__ import annotations

from pathlib import Path

import pytest

from common.exceptions import ValidationError
from modules.co_analysis.validators import uploaded_workbook_validator as validator_mod


class _WorkbookStub:
    def close(self) -> None:
        return


def test_uploaded_source_workbook_uses_router_marks_validation(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "marks.xlsx"
    workbook.touch()

    calls: list[dict[str, object]] = []

    monkeypatch.setattr(validator_mod, "_read_system_manifest_payload", lambda _wb: ("COURSE_SETUP_V2", {}))
    monkeypatch.setattr(
        validator_mod,
        "validate_workbooks",
        lambda **kwargs: calls.append(dict(kwargs)) or {"valid_paths": [str(workbook)], "rejections": []},
    )
    monkeypatch.setattr("openpyxl.load_workbook", lambda *args, **kwargs: _WorkbookStub())

    validator_mod.validate_uploaded_source_workbook(workbook)

    assert calls
    assert calls[0]["workbook_kind"] == "marks_template"
    assert calls[0]["workbook_paths"] == [str(workbook)]


def test_uploaded_source_workbook_raises_validation_error_from_rejection(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "marks.xlsx"
    workbook.touch()

    monkeypatch.setattr(validator_mod, "_read_system_manifest_payload", lambda _wb: ("COURSE_SETUP_V2", {}))
    monkeypatch.setattr(
        validator_mod,
        "validate_workbooks",
        lambda **kwargs: {
            "valid_paths": [],
            "rejections": [
                {
                    "path": str(workbook),
                    "issue": {
                        "code": "COA_MARK_ENTRY_EMPTY",
                        "message": "entry empty",
                        "context": {"cell": "D10"},
                    },
                }
            ],
        },
    )
    monkeypatch.setattr("openpyxl.load_workbook", lambda *args, **kwargs: _WorkbookStub())

    with pytest.raises(ValidationError) as excinfo:
        validator_mod.validate_uploaded_source_workbook(workbook)

    assert excinfo.value.code == "COA_MARK_ENTRY_EMPTY"
    assert excinfo.value.context.get("cell") == "D10"
