# Quality Gate

## Purpose

Use this checklist before release promotion to verify quality, security, and runtime behavior from one place.
Keep high-level policy in `AGENTS.md`; keep executable gate commands here.

## Environment

- Conda env: `obe`
- Preferred prefix: `conda run -n obe python -m ...`

## Required Checks

1. `conda run -n obe python scripts/quality_gate.py --mode strict`
2. `conda run -n obe python -m ruff check .`
3. `conda run -n obe python -m isort --check-only --diff .`
4. `conda run -n obe python -m pyflakes .`
5. `conda run -n obe python -m pyright`
6. `conda run -n obe python scripts/check_ui_strings.py`
7. `conda run --no-capture-output -n obe python -m pytest -q`
8. `conda run -n obe python -m coverage run -m pytest -q`
9. `conda run -n obe python -m coverage report -m`
10. `conda run -n obe python -m bandit -q -c .bandit.yaml -r common modules services`
11. `conda run -n obe python -m pip_audit --cache-dir .pip_audit_cache --ignore-vuln GHSA-58qw-9mgm-455v`
12. `conda run -n obe python scripts/instructor_perf_soak.py --iterations 10 --enforce --max-step-ms 8000`

## Instructor Reliability Checks

1. Fault injection coverage:
   - `conda run -n obe python -m pytest -q tests/test_instructor_workflow_service.py tests/test_final_co_report_generator.py`
2. Cancellation coverage:
   - `conda run -n obe python -m pytest -q tests/test_instructor_module_cancellation.py tests/test_course_details_template_generator_validation.py tests/test_marks_template_generator.py`

## Performance CI Policy (Phased)

- CI runs `tests/perf` with `RUN_PERF_TESTS=1` in a separate non-blocking job (PR + nightly + manual).
- Keep perf non-blocking while thresholds are tuned and runner variance is observed.
- Promote to blocking only after stable signal:
  - no persistent false failures across at least 2 consecutive weeks
  - threshold breaches reproduce locally and correspond to real regressions
  - perf job runtime remains acceptable for CI cadence

## Release Metadata

1. Generate artifact checksum manifest:
   - `conda run -n obe python scripts/generate_artifact_manifest.py`
2. Tag release commit.
3. Attach checksum/manifest artifacts to the release tag.

## Promotion Rule

- Build in `dev`, verify/sign in `stage`, promote the same verified artifact to `prod` without rebuild.
