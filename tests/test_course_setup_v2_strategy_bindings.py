from __future__ import annotations

from pathlib import Path

import pytest

from domain.template_versions.course_setup_v2_impl import strategy_bindings


def test_co_attainment_generation_inputs_parses_word_report_settings(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Test co attainment generation inputs parses word report settings.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).

    Returns:
        None.

    Raises:
        None.
    """
    monkeypatch.setattr(strategy_bindings, "final_report_signature_reader", lambda: (lambda _path: None))
    monkeypatch.setattr(strategy_bindings, "total_outcomes_reader", lambda: (lambda _path: 6))
    source_path = tmp_path / "source.xlsx"
    result = strategy_bindings.co_attainment_generation_inputs(
        context={
            "source_paths": [str(source_path)],
            "generate_word_report": True,
            "word_output_path": str(tmp_path / "co_report.docx"),
            "co_description_path": str(tmp_path / "co_description.xlsx"),
        },
        output_path=tmp_path / "co_attainment.xlsx",
        default_template_id="COURSE_SETUP_V2",
    )
    assert result["generate_word_report"] is True
    assert result["word_output_path"] == tmp_path / "co_report.docx"
    assert result["co_description_path"] == tmp_path / "co_description.xlsx"


def test_co_attainment_generation_inputs_disables_word_report_for_false_string(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Test co attainment generation inputs disables word report for false string.

    Args:
        monkeypatch: Parameter value (pytest.MonkeyPatch).
        tmp_path: Parameter value (Path).

    Returns:
        None.

    Raises:
        None.
    """
    monkeypatch.setattr(strategy_bindings, "final_report_signature_reader", lambda: (lambda _path: None))
    monkeypatch.setattr(strategy_bindings, "total_outcomes_reader", lambda: (lambda _path: 6))
    source_path = tmp_path / "source.xlsx"
    result = strategy_bindings.co_attainment_generation_inputs(
        context={
            "source_paths": [str(source_path)],
            "generate_word_report": "false",
            "co_description_path": str(tmp_path / "co_description.xlsx"),
        },
        output_path=tmp_path / "co_attainment.xlsx",
        default_template_id="COURSE_SETUP_V2",
    )
    assert result["generate_word_report"] is False
    assert result["word_output_path"] is None
    assert result["co_description_path"] == tmp_path / "co_description.xlsx"
