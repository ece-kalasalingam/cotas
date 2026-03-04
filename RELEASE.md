# Release Checklist

## Preconditions

- All CI jobs green on:
  - `windows-latest`
  - `macos-latest`
  - `ubuntu-latest`
- No dependency vulnerabilities in CI `pip-audit`.
- `FOCUS_WORKBOOK_PASSWORD` is provisioned in target environment.

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

## Platform Distribution Notes

- macOS:
  sign and notarize before distribution.
- Windows:
  sign executable/installer where possible.
- Linux:
  publish package format agreed by distribution plan.

## Branch Protection (Repository Settings)

Configure in GitHub settings:
- require pull request before merge
- require status checks to pass (CI workflow)
- disallow force-push on protected branch
