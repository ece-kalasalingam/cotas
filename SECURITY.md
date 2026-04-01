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
  - Windows: DPAPI (`CryptProtectData`) machine-scoped protection.
  - POSIX: keyring-first storage (`keyring` backend when available), with restricted-permission file fallback (`0600`) if keyring is unavailable.
- Signature versioning:
  `FOCUS_WORKBOOK_SIGNATURE_VERSION` controls active signature format version.
  Legacy unsigned-signature compatibility paths are not accepted in the current release.
- CI quality gate:
  lint, type check, UI string policy check, tests, and coverage gates.
  Executable checklist source: `docs/QUALITY_GATE.md`.
- CI dependency audit:
  `pip-audit` runs on each pull request and push.
- CI SAST gate:
  `bandit` runs on each pull request and push.
- Cross-platform CI matrix:
  Windows, macOS, and Linux.
- Module plugin architecture:
  `MainWindow` loads modules through `modules/module_catalog.py` using explicit plugin specs; this reduces ad hoc import surfaces.

## Module Security Scope

- `Instructor`: workbook validation/signature checks and controlled output writes.
- `CO Analysis`: signed source-workbook intake and routed analysis generation checks.
- `Help`: local asset/PDF handling and controlled file-save/open behavior.
- `About`: static metadata presentation (no privileged data flow).
- `PO Analysis`: placeholder module with no workbook processing or privileged data paths.

## Operational Requirements

- Never commit secrets to git.
- Protect the host profile where app secrets are stored.
- Rotate workbook password by controlled app data reset and template regeneration when required.
- POSIX fallback note: if no keyring backend is available, the fallback secret file is compatibility storage and not equivalent to hardware/OS secret vault protection.
- Release/promotion security checks must pass `AGENTS.md` release policy and the command checklist in `docs/QUALITY_GATE.md`.
