# Rollback Playbook

## Triggers
1. Data corruption risk.
2. Workflow failure rate spike.
3. Startup crash rate above threshold.
4. Security gate regression.

## Immediate Actions
1. Stop current rollout.
2. Revert distribution channel to last known-good signed release.
3. Notify stakeholders with incident ID.

## Technical Steps
1. Retrieve previous signed artifacts and manifest.
2. Verify checksums with `scripts/verify_artifact_manifest.py`.
3. Redeploy previous artifacts to affected environment.
4. Confirm startup, Step 1-3 workflows, and logs.

## Validation Checklist
1. App launches successfully.
2. Course template generation works.
3. Course-details validation works.
4. Filled-marks schema validation works.
5. No new unhandled crash reports.

## Post-Rollback
1. Collect logs/crash reports.
2. Open root cause analysis task.
3. Add regression tests before re-release.

