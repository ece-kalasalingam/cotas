from __future__ import annotations

from domain.template_versions.course_setup_v2_impl import marks_template_validator as v2_validator


def test_v2_marks_validator_warning_buffer_consume_clears_state() -> None:
    v2_validator._reset_marks_anomaly_warnings()
    v2_validator._log_marks_anomaly_warnings_from_stats(
        sheet_name="Sheet1",
        mark_cols=range(4, 6),
        student_count=10,
        absent_count_by_col={4: 5, 5: 0},
        numeric_count_by_col={4: 0, 5: 10},
        frequency_by_value_by_col={4: {}, 5: {1.0: 10}},
    )

    warnings = v2_validator.consume_last_marks_anomaly_warnings()
    assert len(warnings) == 2
    assert any("High absence ratio" in message for message in warnings)
    assert any("Near-constant marks" in message for message in warnings)
    assert v2_validator.consume_last_marks_anomaly_warnings() == []


def test_v2_marks_validator_resets_warning_buffer_per_run() -> None:
    v2_validator._last_marks_anomaly_warnings[:] = ["stale warning"]

    class _Workbook:
        sheetnames = ["OnlySheet"]

    try:
        v2_validator.validate_filled_marks_manifest_schema(workbook=_Workbook(), manifest={"sheet_order": [], "sheets": []})
    except Exception:
        pass

    # stale warnings should be cleared at validation start
    assert "stale warning" not in v2_validator.consume_last_marks_anomaly_warnings()
