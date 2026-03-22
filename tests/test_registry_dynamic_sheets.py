from __future__ import annotations

from common.registry import (
    CO_REPORT_SHEET_KEY_CO_INDIRECT,
    resolve_dynamic_sheet_headers,
)


def test_co_indirect_dynamic_headers_match_for_v1_and_v2() -> None:
    context = {
        "components": [("Survey", 20.0), ("Alumni", 80.0)],
        "ratio": 0.2,
    }
    v1_headers = resolve_dynamic_sheet_headers(
        "COURSE_SETUP_V1",
        sheet_key=CO_REPORT_SHEET_KEY_CO_INDIRECT,
        context=context,
    )
    v2_headers = resolve_dynamic_sheet_headers(
        "COURSE_SETUP_V2",
        sheet_key=CO_REPORT_SHEET_KEY_CO_INDIRECT,
        context=context,
    )
    assert v1_headers == v2_headers
    assert v1_headers[0:3] == ("#", "Reg. No.", "Student Name")
    assert v1_headers[-2:] == ("Total (100%)", "Total (20%)")

