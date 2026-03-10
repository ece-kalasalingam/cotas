# Instructor Module Enterprise Runbook

## Scope
- Module: `modules/instructor_module.py`
- Service: `services/instructor_workflow_service.py`
- Engine facade: `domain/instructor_engine.py`
- Template/report engines: `domain/instructor_template_engine.py`, `domain/instructor_report_engine.py`

## Operational SLO Baseline
- UI responsiveness: workflow actions must run in background and must not block the Qt event loop.
- Reliability: final report write must be atomic (no partial overwrite on failure).
- Cancellation: cancellation must be honored before and during long workbook operations.

## Performance/Soak Procedure
1. Run:
   - `conda run -n obe python scripts/instructor_perf_soak.py --iterations 10`
2. Enforced gate (CI or release candidate):
   - `conda run -n obe python scripts/instructor_perf_soak.py --iterations 10 --enforce --max-step-ms 8000`
3. Review JSON output:
   - `p95_ms` by step
   - `threshold_breaches`

## Fault Injection Procedure
- File write interruption / permission simulation is covered by tests:
  - `tests/test_instructor_workflow_service.py`
  - `tests/test_instructor_module_step5.py`
- Run:
  - `conda run -n obe python -m pytest -q tests/test_instructor_workflow_service.py tests/test_instructor_module_step5.py`

## Cancellation Verification
- Regression tests:
  - `tests/test_instructor_module_cancellation.py`
  - generator cancellation in:
    - `tests/test_course_details_template_generator_validation.py`
    - `tests/test_marks_template_generator.py`
- Run:
  - `conda run -n obe python -m pytest -q tests/test_instructor_module_cancellation.py tests/test_course_details_template_generator_validation.py tests/test_marks_template_generator.py`

## Audit and Telemetry
- Service emits lifecycle logs for each workflow operation with `job_id` and `step_id`.
- Structured log extras include stable `error_code` and workflow `event`.
- In-memory metrics snapshot is attached in workflow completion/failure logs.
- Events:
  - started
  - completed (with duration)
  - cancelled (with duration)
  - failed (with duration)
- Source: `services/instructor_workflow_service.py`

## Secret Handling and Rotation
- Workbook hash secret is sourced from environment variable:
  - `FOCUS_WORKBOOK_PASSWORD`
- Optional previous secrets for rotation:
  - `FOCUS_WORKBOOK_PASSWORD_PREVIOUS` (comma-separated)
- Signature format version:
  - `FOCUS_WORKBOOK_SIGNATURE_VERSION`
- Minimum policy:
  - length >= 12
  - rotate periodically (recommended: every 90 days)
  - do not commit plaintext secrets into repo
- Rotation checklist:
  1. update deployment secret store
  2. restart app/runtime
  3. generate new templates using new secret
  4. archive old templates with version/date tags

## Release Gate
- Required:
  1. `conda run -n obe python -m pyflakes .`
  2. `conda run -n obe python -m pytest -q`
  3. `conda run -n obe python scripts/check_ui_strings.py`
  4. `conda run -n obe python -m bandit -q -r common -r modules -r services -x tests`
  5. `conda run -n obe python scripts/instructor_perf_soak.py --iterations 10 --enforce --max-step-ms 8000`

## Timeout and Crash Pipeline
- Workflow step timeout env:
  - `FOCUS_WORKFLOW_STEP_TIMEOUT_SECONDS` (default `120`)
- Crash spool for packaged builds:
  - Local JSON reports under app settings `crash_reports/`.
- Optional remote pipeline endpoint marker:
  - `FOCUS_CRASH_REPORT_ENDPOINT`
