from __future__ import annotations

import os
import statistics
import time
import tracemalloc
from pathlib import Path

import pytest

from common.exceptions import ValidationError
from domain.template_versions.course_setup_v2_impl import (
    marks_template_validator as marks_validator,
)


def _baseline_marks_batch(
    workbook_paths: list[str],
    *,
    template_id: str,
) -> dict[str, object]:
    unique_paths: list[str] = []
    duplicate_paths: list[str] = []
    seen: set[str] = set()
    for raw in workbook_paths:
        path = str(raw).strip()
        if not path:
            continue
        if path in seen:
            duplicate_paths.append(path)
            continue
        seen.add(path)
        unique_paths.append(path)

    valid_paths: list[str] = []
    invalid_paths: list[str] = []
    mismatched_paths: list[str] = []
    template_ids: dict[str, str] = {}
    duplicate_sections: list[str] = []
    rejections: list[dict[str, object]] = []

    for path in duplicate_paths:
        rejections.append(
            {
                "path": path,
                "reason_kind": "duplicate_path",
                "issue": {"code": "MARKS_TEMPLATE_DUPLICATE_PATH"},
            }
        )

    for path in unique_paths:
        try:
            resolved = marks_validator._validate_filled_marks_workbook_impl(
                workbook_path=path,
                expected_template_id=template_id,
                cancel_token=None,
            )
            valid_paths.append(path)
            template_ids[path] = str(getattr(resolved, "template_id", "") or "")
        except ValidationError as exc:
            invalid_paths.append(path)
            reason_kind = "template_mismatch" if getattr(exc, "code", "") == "UNKNOWN_TEMPLATE" else "invalid"
            if reason_kind == "template_mismatch":
                mismatched_paths.append(path)
            rejections.append(
                {
                    "path": path,
                    "reason_kind": reason_kind,
                    "issue": {"code": getattr(exc, "code", "VALIDATION_ERROR")},
                }
            )
        except Exception:
            invalid_paths.append(path)
            rejections.append(
                {
                    "path": path,
                    "reason_kind": "invalid",
                    "issue": {"code": "MARKS_TEMPLATE_UNEXPECTED_REJECTION"},
                }
            )

    return {
        "valid_paths": valid_paths,
        "invalid_paths": invalid_paths,
        "mismatched_paths": mismatched_paths,
        "duplicate_paths": duplicate_paths,
        "duplicate_sections": duplicate_sections,
        "template_ids": template_ids,
        "rejections": rejections,
    }


def _timed_runs(fn, *, runs: int = 7) -> tuple[float, float]:
    samples: list[float] = []
    for _ in range(runs):
        start = time.perf_counter()
        fn()
        samples.append(time.perf_counter() - start)
    median = statistics.median(samples)
    p95 = samples[min(len(samples) - 1, max(0, int(round((len(samples) - 1) * 0.95))))]
    return median, p95


@pytest.mark.skipif(
    os.environ.get("RUN_PERF_TESTS", "0") != "1",
    reason="Non-blocking perf suite. Set RUN_PERF_TESTS=1 to execute.",
)
def test_v2_marks_batch_perf_baseline_vs_refactor(monkeypatch: pytest.MonkeyPatch) -> None:
    def _fake_validate(*, workbook_path: str, expected_template_id: str, cancel_token=None) -> object:
        del expected_template_id
        del cancel_token
        idx = int(Path(workbook_path).stem.replace("wb", ""))
        if idx % 11 == 0:
            raise ValidationError("template mismatch", code="UNKNOWN_TEMPLATE", context={"workbook": workbook_path})
        if idx % 7 == 0:
            raise ValidationError("invalid data", code="COA_MARK_ENTRY_EMPTY", context={"workbook": workbook_path})
        return marks_validator._MarksWorkbookIdentity(
            template_id="COURSE_SETUP_V2",
            course_code="CS101",
            semester="V",
            academic_year="2026-27",
            total_outcomes=3,
            section=f"S{idx}",
            reg_numbers=frozenset({f"{idx:03d}"}),
        )

    monkeypatch.setattr(marks_validator, "_validate_filled_marks_workbook_impl", _fake_validate)
    workload = [f"wb{i}.xlsx" for i in range(1, 241)] + ["wb2.xlsx", "wb18.xlsx", "wb19.xlsx"]

    baseline_median, baseline_p95 = _timed_runs(
        lambda: _baseline_marks_batch(workload, template_id="COURSE_SETUP_V2")
    )
    refactor_median, refactor_p95 = _timed_runs(
        lambda: marks_validator.validate_filled_marks_workbooks(workload, template_id="COURSE_SETUP_V2")
    )

    tracemalloc.start()
    marks_validator.validate_filled_marks_workbooks(workload, template_id="COURSE_SETUP_V2")
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    assert baseline_median > 0
    assert baseline_p95 > 0
    assert refactor_median > 0
    assert refactor_p95 > 0
    assert peak > 0
    assert refactor_median <= baseline_median * 1.5, (
        f"Median regression too high: baseline={baseline_median:.6f}s, "
        f"refactor={refactor_median:.6f}s"
    )
