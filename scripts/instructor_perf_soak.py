from __future__ import annotations

import argparse
import json
import statistics
import sys
import tempfile
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from services import InstructorWorkflowService


def _time_call(fn):
    started = time.perf_counter()
    fn()
    return (time.perf_counter() - started) * 1000.0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run instructor workflow performance/soak checks."
    )
    parser.add_argument(
        "--iterations",
        type=int,
        default=5,
        help="How many full workflow iterations to run.",
    )
    parser.add_argument(
        "--enforce",
        action="store_true",
        help="Fail with non-zero exit code when thresholds are exceeded.",
    )
    parser.add_argument(
        "--max-step-ms",
        type=float,
        default=8000.0,
        help="Per-step p95 threshold in milliseconds.",
    )
    args = parser.parse_args()

    service = InstructorWorkflowService()
    timings: dict[str, list[float]] = {
        "generate_course_details_template": [],
        "validate_course_details_workbook": [],
        "generate_marks_template": [],
        "generate_final_report": [],
    }

    with tempfile.TemporaryDirectory(prefix="instructor_perf_") as temp_dir:
        root = Path(temp_dir)
        for index in range(args.iterations):
            course_details = root / f"course_details_{index}.xlsx"
            marks_template = root / f"marks_template_{index}.xlsx"
            final_report = root / f"final_report_{index}.xlsx"

            ctx1 = service.create_job_context(step_id="step1", payload={"i": index})
            timings["generate_course_details_template"].append(
                _time_call(
                    lambda: service.generate_course_details_template(
                        course_details, context=ctx1
                    )
                )
            )

            ctx2 = service.create_job_context(step_id="step2", payload={"i": index})
            timings["validate_course_details_workbook"].append(
                _time_call(
                    lambda: service.validate_course_details_workbook(
                        course_details, context=ctx2
                    )
                )
            )

            ctx3 = service.create_job_context(step_id="step2", payload={"i": index})
            timings["generate_marks_template"].append(
                _time_call(
                    lambda: service.generate_marks_template(
                        course_details, marks_template, context=ctx3
                    )
                )
            )

            ctx4 = service.create_job_context(step_id="step3", payload={"i": index})
            timings["generate_final_report"].append(
                _time_call(
                    lambda: service.generate_final_report(
                        marks_template, final_report, context=ctx4
                    )
                )
            )

    summary: dict[str, dict[str, float]] = {}
    threshold_breaches: list[str] = []
    for step_name, values in timings.items():
        ordered = sorted(values)
        p95_index = max(0, int(round(0.95 * (len(ordered) - 1))))
        p95 = ordered[p95_index]
        step_summary = {
            "min_ms": round(min(values), 2),
            "mean_ms": round(statistics.mean(values), 2),
            "p95_ms": round(p95, 2),
            "max_ms": round(max(values), 2),
        }
        summary[step_name] = step_summary
        if step_summary["p95_ms"] > args.max_step_ms:
            threshold_breaches.append(step_name)

    report = {
        "iterations": args.iterations,
        "threshold_ms": args.max_step_ms,
        "summary": summary,
        "threshold_breaches": threshold_breaches,
    }
    print(json.dumps(report, indent=2))

    if args.enforce and threshold_breaches:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
