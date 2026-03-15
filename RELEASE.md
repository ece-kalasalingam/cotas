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

Use commands in `commands.txt`.

- Windows:
  use PyInstaller `--add-data "assets;assets"`
- macOS/Linux:
  use PyInstaller `--add-data "assets:assets"`

## Verification

- Launch app smoke test on each target OS.
- Verify workbook creation/validation flow end-to-end.
- Verify translations load and language switcher works.
- Verify logs are written and no unhandled exceptions on startup.
- Verify crash spool path exists for packaged builds (`crash_reports/` in app settings dir).

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
