from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("xlsxwriter")
openpyxl = pytest.importorskip("openpyxl")

from common.constants import ID_COURSE_SETUP
from common.exceptions import ValidationError
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
