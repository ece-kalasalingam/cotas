# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

### Added
- Added test coverage to lock style/theme initialization order when both `Fusion` style and `qdarktheme` are used together.
- Added shared output-link helpers for Instructor to align with Coordinator behavior and reduce duplication.

### Changed
- Refactored Instructor flow to a two-step model and updated step usage across UI actions, workflow code, and language strings.
- Standardized native-OS-first UI behavior by reducing custom styling overrides and inline spacing customizations across modules.
- Kept Help module context menu rendering native while supporting style application order behavior safely.

### Fixed
- Resolved high-volume static typing issues across Coordinator, Instructor, common utilities, and test modules.
- Fixed Pylance/Pyright protocol and callable typing mismatches in output-link and async-runner related paths.
- Fixed test compatibility for Help context menu by guarding style application on menu test doubles.

### Removed
- Removed stale package `__all__` exports that produced unsupported-dunder-all warnings in static analysis.
- Removed dead style/constants paths left behind after native-style cleanup.

### Quality
- `pyright`: `0 errors, 0 warnings`
- `pytest -q`: `474 passed`
- `ruff check .`: `passed`
- `isort --check-only --diff .`: `passed`
- `pyflakes .`: `passed`
- `bandit -q -r . -c .bandit.yaml`: `passed`
- `pip-audit`: `No known vulnerabilities found`
