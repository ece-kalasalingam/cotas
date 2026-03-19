# Release Checklist

## Preconditions

- All CI jobs green on:
  - `windows-latest`
  - `macos-latest`
  - `ubuntu-latest`
- No dependency vulnerabilities in CI `pip-audit`.
- SAST gate is green (`bandit`).
- Workbook secret store initialization is verified on target environment.

## Build

Use commands in `installer/commands.txt`.

- Windows:
  use PyInstaller `--add-data "assets;assets"`
- macOS/Linux:
  use PyInstaller `--add-data "assets:assets"`

## Verification

- Launch app smoke test on each target OS.
- Verify workbook creation/validation flow end-to-end using current two-step instructor flow.
- Verify module navigation and lazy loading through plugin catalog (`modules/module_catalog.py`).
- Verify translations load and language switcher works.
- Verify logs are written and no unhandled exceptions on startup.
- Verify crash spool path exists for packaged builds (`crash_reports/` in app settings dir).

### Module Verification Matrix

- `Instructor`: Step 1 and Step 2 paths complete with expected outputs and cancellation behavior.
- `Coordinator`: file intake, validation, and consolidated attainment generation complete.
- `Help`: help PDF renders, exports, and opens in default viewer.
- `About`: version/build details and static assets render correctly.

## Required Quality Gates (Current)

- `conda run -n obe python scripts/quality_gate.py --mode strict`
- `conda run -n obe python -m ruff check .`
- `conda run -n obe python -m isort --check-only --diff .`
- `conda run -n obe python -m pyflakes .`
- `conda run -n obe python -m pyright`
- `conda run --no-capture-output -n obe python -m pytest -q`
- `conda run -n obe python -m coverage run -m pytest -q`
- `conda run -n obe python -m coverage report -m`
- `conda run -n obe python -m bandit -q -r common modules services -c .bandit.yaml`
- `conda run -n obe python -m pip_audit --cache-dir .pip_audit_cache`

## Platform Distribution Notes

- macOS:
  sign and notarize before distribution.
- Windows:
  sign executable/installer where possible.
- Linux:
  publish package format agreed by distribution plan.

## Artifact Integrity

1. Generate checksum manifest:
   - `python scripts/generate_artifact_manifest.py <artifact1> <artifact2> --out artifact-manifest.json`
2. Verify checksums before release:
   - `python scripts/verify_artifact_manifest.py artifact-manifest.json`
3. Attach signed artifacts and manifest to release.

## Promotion and Rollback

- Promotion stages and controls: `docs/RELEASE_PROMOTION.md`
- Rollback procedure: `docs/ROLLBACK_PLAYBOOK.md`
- Support/incident process: `docs/SUPPORT_RUNBOOK.md`

## Branch Protection (Repository Settings)

Configure in GitHub settings:
- require pull request before merge
- require status checks to pass (CI workflow)
- disallow force-push on protected branch
