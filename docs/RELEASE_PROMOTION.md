# Release Promotion Stages

## Stages
1. `dev`
2. `stage`
3. `prod`

## Entry Criteria
1. CI green on all OS matrix jobs.
2. Security gates pass (`pip-audit`, `bandit`).
3. Quality gate pass (`ruff`, `isort`, `pyflakes`, `pyright`, UI string checks, tests, coverage).
4. Artifact checksums generated via `scripts/generate_artifact_manifest.py`.
5. Module catalog validation complete (`modules/module_catalog.py` includes expected modules and navigation labels).

## Promotion Procedure
1. Build artifacts in `dev`.
2. Sign artifacts using organization signing process.
3. Verify signed artifacts and checksum manifest in `stage`.
4. Execute UAT + smoke tests in `stage` for all modules (`Instructor`, `Coordinator`, `Help`, `About`).
5. Promote the same verified artifacts to `prod` (no rebuild).

## Promotion Controls
1. Require approver from QA.
2. Require approver from engineering lead.
3. Tag release commit and attach manifest.

