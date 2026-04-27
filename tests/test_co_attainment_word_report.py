from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from common.constants import ID_COURSE_SETUP
from common.exceptions import ValidationError
from common.jobs import CancellationToken
from domain.template_strategy_router import generate_workbook
from domain.template_versions.course_setup_v2_impl import co_attainment


def test_build_co_word_summary_rows_uses_existing_attainment_state() -> None:
    """Test build co word summary rows uses existing attainment state.

    Args:
        None.

    Returns:
        None.

    Raises:
        None.
    """
    pending_rows = {
        1: [
            co_attainment._CoAttainmentRow(
                reg_hash=1,
                reg_no="R1",
                student_name="Student 1",
                direct_score=64.0,
                indirect_score=16.0,
                worksheet_name="CO1",
                workbook_name="source.xlsx",
            ),
            co_attainment._CoAttainmentRow(
                reg_hash=2,
                reg_no="R2",
                student_name="Student 2",
                direct_score=48.0,
                indirect_score=12.0,
                worksheet_name="CO1",
                workbook_name="source.xlsx",
            ),
        ]
    }
    output_states = {
        1: co_attainment._CoOutputSheetState(
            sheet=None,
            header_row_index=0,
            formats={},
            next_row_index=0,
            next_serial=1,
            on_roll=2,
            attended=2,
            level_counts={0: 0, 1: 0, 2: 1, 3: 1},
        )
    }
    rows = co_attainment._build_co_word_summary_rows(
        pending_rows=pending_rows,
        output_states=output_states,
        total_outcomes=1,
        course_code="CSE101",
        co_attainment_level=2,
        co_attainment_percent=60.0,
    )
    assert rows[0]["co"] == "CSE101.1"
    assert rows[0]["direct"] == "70%"
    assert rows[0]["indirect"] == "70%"
    assert rows[0]["overall"] == "100%"
    assert rows[0]["result"] == "Attained"
    assert rows[0]["shortfall"] == "0%"
    assert rows[0]["severity"] == "On Target (Meets CO AT Target)"


def test_generate_co_attainment_word_report_writes_docx(tmp_path: Path) -> None:
    """Test generate co attainment word report writes docx.

    Args:
        tmp_path: Parameter value (Path).

    Returns:
        None.

    Raises:
        None.
    """
    pytest.importorskip("docx")
    output_path = tmp_path / "co_report.docx"
    generated = co_attainment._generate_co_attainment_word_report(
        output_path=output_path,
        metadata={
            "course_code": "CSE101",
            "semester": "5",
            "section": "A",
            "academic_year": "2025-26",
        },
        thresholds=(40.0, 60.0, 75.0),
        co_attainment_percent=60.0,
        co_attainment_level=2,
        total_outcomes=2,
        co_rows=[
            {"co": "CSE101.1", "direct": "70%", "indirect": "68%", "overall": "75%", "result": "Attained"},
            {"co": "CSE101.2", "direct": "55%", "indirect": "58%", "overall": "50%", "result": "Yet to Attain"},
        ],
        co_sentences=[
            "Understand semiconductor devices",
            "Analyze communication systems",
        ],
    )
    assert generated == output_path
    assert output_path.exists()
    with ZipFile(output_path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
    assert "KALASALINGAM ACADEMY OF RESEARCH AND EDUCATION" in document_xml
    assert "Course Coordinator Report" in document_xml
    assert "CSE101" in document_xml
    assert "2025-26" in document_xml
    assert "2 course outcomes (COs)" in document_xml or "2 COs" in document_xml
    assert "CO-wise Attainment Summary" in document_xml
    assert "Course Outcomes" in document_xml
    assert "The students will be able to:" in document_xml
    assert "CO1: Understand semiconductor devices." in document_xml
    assert "CO2: Analyze communication systems." in document_xml
    assert "Severity" not in document_xml
    assert "Identification of Shortfall COs" not in document_xml
    assert "Severity Classification" not in document_xml
    assert "Severity is classified by CO attainment shortfall percentage only" not in document_xml
    assert "Recommended Corrective Actions" not in document_xml
    assert "Continuous Improvement Action Suggestions" in document_xml


def _generate_co_description_template(path: Path) -> Path:
    result = generate_workbook(
        template_id=ID_COURSE_SETUP,
        output_path=path,
        workbook_name=path.name,
        workbook_kind="co_description_template",
    )
    output = getattr(result, "output_path", None)
    if isinstance(output, str) and output.strip():
        return Path(output)
    return path


def _fill_co_descriptions(path: Path, *, total_outcomes: int = 6) -> None:
    openpyxl = pytest.importorskip("openpyxl")
    workbook = openpyxl.load_workbook(path)
    try:
        sheet = workbook["CO_Description"]
        for index in range(1, total_outcomes + 1):
            row = index + 1
            sheet.cell(row=row, column=1, value=index)
            sheet.cell(row=row, column=2, value=f"CO{index} statement")
            sheet.cell(row=row, column=3, value="K2")
            sheet.cell(row=row, column=4, value=f"CO{index} summary " + ("x" * 120))
        workbook.save(path)
    finally:
        workbook.close()


def test_validated_co_description_sentences_reads_ordered_descriptions(tmp_path: Path) -> None:
    workbook_path = _generate_co_description_template(tmp_path / "co_description.xlsx")
    _fill_co_descriptions(workbook_path, total_outcomes=6)
    sentences = co_attainment._validated_co_description_sentences(
        co_description_path=workbook_path,
        template_id=ID_COURSE_SETUP,
        total_outcomes=6,
        token=CancellationToken(),
    )
    assert sentences == [f"CO{index} statement" for index in range(1, 7)]


def test_validated_co_description_sentences_rejects_total_outcomes_mismatch(tmp_path: Path) -> None:
    workbook_path = _generate_co_description_template(tmp_path / "co_description_mismatch.xlsx")
    _fill_co_descriptions(workbook_path, total_outcomes=6)
    with pytest.raises(ValidationError) as excinfo:
        co_attainment._validated_co_description_sentences(
            co_description_path=workbook_path,
            template_id=ID_COURSE_SETUP,
            total_outcomes=5,
            token=CancellationToken(),
        )
    assert getattr(excinfo.value, "code", "") in {
        "CO_DESCRIPTION_MARKS_COHORT_MISMATCH",
        "CO_DESCRIPTION_CO_NUMBER_SET_MISMATCH",
    }
