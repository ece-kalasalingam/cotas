# Support and Incident Runbook

## Triage Inputs
1. App log file (`focus.log` in app settings directory).
2. Crash reports (`crash_reports/*.json` in app settings directory).
3. User workbook causing failure.
4. Application version and OS details.

## Severity Classification
1. `SEV-1`: outage/data loss/security impact.
2. `SEV-2`: major workflow blocked.
3. `SEV-3`: partial degradation.
4. `SEV-4`: cosmetic/non-blocking.

## Investigation Steps
1. Identify `job_id` and `step_id` from logs.
2. Check structured `error_code`.
3. Reproduce with same workbook in staging.
4. Determine whether issue is validation, system, timeout, or corruption.

## Module Routing
1. `Instructor`: template generation/validation/final report issues.
2. `Coordinator`: multi-file intake and attainment output issues.
3. `Help`: PDF load/save/open issues.
4. `About`: metadata/version rendering issues.

## Communication
1. Acknowledge incident with severity and ETA.
2. Update status every 30 minutes for `SEV-1/2`.
3. Share workaround if available.

## Resolution
1. Apply fix or rollback.
2. Verify with regression tests.
3. Publish incident summary and RCA action items.

