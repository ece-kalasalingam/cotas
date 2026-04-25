from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("xlsxwriter")
openpyxl = pytest.importorskip("openpyxl")

from common.constants import ID_COURSE_SETUP
from common.exceptions import ValidationError
from domain import template_strategy_router as router
from domain.template_strategy_router import generate_workbook


def test_router_accepts_co_description_template_kind(tmp_path: Path) -> None:
    """Test router accepts co description template kind.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    output = tmp_path / "co_description_template.xlsx"
    result = generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=output,
        workbook_name=output.name,
        workbook_kind="co_description_template",
    )
    resolved_output = str(getattr(result, "output_path", output)).strip()
    workbook = openpyxl.load_workbook(resolved_output)
    try:
        assert workbook.sheetnames == ["Course_Metadata", "CO_Description", "__SYSTEM_HASH__"]
    finally:
        workbook.close()


def test_router_rejects_unknown_single_generation_kind(tmp_path: Path) -> None:
    """Test router rejects unknown single generation kind.
    
    Args:
        tmp_path: Parameter value (Path).
    
    Returns:
        None.
    
    Raises:
        None.
    """
    output = tmp_path / "invalid.xlsx"
    with pytest.raises(ValidationError) as excinfo:
        generate_workbook(
            template_id=ID_COURSE_SETUP,
            output_path=output,
            workbook_name=output.name,
            workbook_kind="unsupported_template",
        )
    assert getattr(excinfo.value, "code", None) == "WORKBOOK_KIND_UNSUPPORTED"


def test_router_preserves_word_report_fields_on_generate_result(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Test router preserves word report fields on generate result.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).

    Returns:
        None.

    Raises:
        None.
    """

    class _FakeStrategy:
        template_id = ID_COURSE_SETUP

        def supports_operation(self, operation: str) -> bool:
            return operation == "generate_workbook"

        def generate_workbook(self, **_kwargs):
            return type(
                "_RawResult",
                (),
                {
                    "status": "generated",
                    "output_path": str(tmp_path / "co.xlsx"),
                    "word_report_path": str(tmp_path / "co.docx"),
                    "word_report_error_key": "co_analysis.status.word_report_generate_failed",
                },
            )()

    monkeypatch.setattr(router, "get_template_strategy", lambda _template_id: _FakeStrategy())
    result = generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=tmp_path / "co.xlsx",
        workbook_name="co.xlsx",
        workbook_kind="co_attainment",
    )
    assert str(getattr(result, "word_report_path", "")).endswith("co.docx")
    assert getattr(result, "word_report_error_key", "") == "co_analysis.status.word_report_generate_failed"
