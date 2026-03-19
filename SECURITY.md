# Security Policy

## Supported Versions

Only the latest release branch is supported for security updates.

## Reporting a Vulnerability

Do not open public issues for security vulnerabilities.

Report privately to maintainers with:
- impact summary
- reproduction steps
- affected version/commit
- suggested remediation (if available)

Acknowledgement target: 3 business days.
Initial triage target: 7 business days.

## Security Controls in This Repository

- Required runtime secret:
  workbook secret is auto-provisioned and stored per machine/user.
- Signature versioning:
  `FOCUS_WORKBOOK_SIGNATURE_VERSION` controls active signature format version.
  Legacy unsigned-signature compatibility paths are not accepted in the current release.
- CI quality gate:
  lint, type check, UI string policy check, and tests.
- CI dependency audit:
  `pip-audit` runs on each pull request and push.
- CI SAST gate:
  `bandit` runs on each pull request and push.
- Cross-platform CI matrix:
  Windows, macOS, and Linux.

## Module Security Scope

- `Instructor`: workbook validation/signature checks and controlled output writes.
- `Coordinator`: signed workbook intake paths and aggregation integrity checks.
- `Help`: local asset/PDF handling and controlled file-save/open behavior.
- `About`: static metadata presentation (no privileged data flow).

## Operational Requirements

- Never commit secrets to git.
- Protect the host profile where app secrets are stored.
- Rotate workbook password by controlled app data reset and template regeneration when required.
