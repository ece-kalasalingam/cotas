from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

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
    assert "Severity" in document_xml
    assert "Identification of Shortfall COs" not in document_xml
    assert "Severity Classification" not in document_xml
    assert "Severity is classified by CO attainment shortfall percentage only" in document_xml
    assert "Recommended Corrective Actions" not in document_xml
    assert "Continuous Improvement Action Suggestions" in document_xml
