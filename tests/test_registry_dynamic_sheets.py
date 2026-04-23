from __future__ import annotations

from common.registry import (
    CO_REPORT_SHEET_KEY_CO_DIRECT,
    CO_REPORT_SHEET_KEY_CO_INDIRECT,
    COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE,
    COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
    COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT,
    resolve_dynamic_sheet_headers,
)


def test_co_direct_dynamic_headers_for_v2() -> None:
    """Test co direct dynamic headers for v2.

    Args:
        None.

    Returns:
        None.

    Raises:
        None.
    """
    context = {
        "components": [("CAT", 30.0, 25.0), ("SEE", 60.0, 75.0)],
        "ratio": 0.8,
    }
    v2_headers = resolve_dynamic_sheet_headers(
        "COURSE_SETUP_V2",
        sheet_key=CO_REPORT_SHEET_KEY_CO_DIRECT,
        context=context,
    )
    assert v2_headers[0:3] == ("#", "Reg. No.", "Student Name")
    assert v2_headers[3:7] == ("CAT (30)", "CAT (25%)", "SEE (60)", "SEE (75%)")
    assert v2_headers[-3:] == ("Total", "Total (100%)", "Total (80%)")


def test_co_indirect_dynamic_headers_for_v2() -> None:
    """Test co indirect dynamic headers for v2.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    context = {
        "components": [("Survey", 20.0), ("Alumni", 80.0)],
        "ratio": 0.2,
    }
    v2_headers = resolve_dynamic_sheet_headers(
        "COURSE_SETUP_V2",
        sheet_key=CO_REPORT_SHEET_KEY_CO_INDIRECT,
        context=context,
    )
    assert v2_headers[0:3] == ("#", "Reg. No.", "Student Name")
    assert v2_headers[-2:] == ("Total (100%)", "Total (20%)")


def test_marks_dynamic_headers_for_v2() -> None:
    """Test marks dynamic headers for v2.
    
    Args:
        None.
    
    Returns:
        None.
    
    Raises:
        None.
    """
    direct_co_headers = resolve_dynamic_sheet_headers(
        "COURSE_SETUP_V2",
        sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_CO_WISE,
        context={"question_count": 3},
    )
    direct_non_co_headers = resolve_dynamic_sheet_headers(
        "COURSE_SETUP_V2",
        sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_DIRECT_NON_CO_WISE,
        context={"covered_cos": [1, 3, 5]},
    )
    indirect_headers = resolve_dynamic_sheet_headers(
        "COURSE_SETUP_V2",
        sheet_key=COURSE_SETUP_SHEET_KEY_MARKS_INDIRECT,
        context={"total_outcomes": 4},
    )

    assert direct_co_headers == ("#", "Reg. No.", "Student Name", "Q1", "Q2", "Q3", "Total")
    assert direct_non_co_headers == (
        "#",
        "Reg. No.",
        "Student Name",
        "Total",
        "Marks for CO1",
        "Marks for CO3",
        "Marks for CO5",
    )
    assert indirect_headers == ("#", "Reg. No.", "Student Name", "CO1", "CO2", "CO3", "CO4")

