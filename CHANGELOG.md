# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

## [1.2.4] - 2026-05-06

### Added
- Added test coverage to lock style/theme initialization order when both `Fusion` style and `qdarktheme` are used together.
- Added shared output-link helpers for Instructor to align with Coordinator behavior and reduce duplication.
- Added plugin module contract (`common/module_plugins.py`) and module catalog (`modules/module_catalog.py`) for plug-and-play module scaling.
- Added dedicated `modules/po_analysis_module.py` placeholder module wired through plugin catalog.
- Added coordinator namespace contract guard (`modules/coordinator/contracts.py`) to validate step integration keys at startup.
- Added template strategy router (`domain/template_strategy_router.py`) as the central template operation dispatch path.
- Added centralized CO direct/indirect sheet writer (`domain/co_report_sheet_generator.py`) as single generation source.
- Added updated architecture/complexity documentation at `docs/COMPLEXITY_ANALYSIS.md`.

### Changed
- Refactored Instructor module from two-step flow to a single workflow (course template download + course-details validation + marks-template generation).
- Reworked Instructor i18n keys/messages to remove step-number naming in active runtime paths.
- Centralized template-aware operation routing so module/domain callsites dispatch through router + template strategy.
- Standardized native-OS-first UI behavior by reducing custom styling overrides and inline spacing customizations across modules.
- Kept Help module context menu rendering native while supporting style application order behavior safely.
- Refactored `MainWindow` to registry-driven module loading with lazy class import (no hardcoded module imports in window bootstrap path).
- Hardened Coordinator orchestration by replacing `globals()` step wiring with explicit namespace factories.
- Migrated coordinator business processing to `domain/coordinator_engine.py` and aligned module/service orchestration to the domain-layer architecture.
- Hardened instructor final-report generation to require service/domain execution (removed service-unavailable copy-bypass behavior).
- Improved coordinator and instructor diagnostics: per-file invalid reasons, full per-file batch failure details, and direct/indirect join-drop visibility.
- Consolidated UI/release/support guardrails into `AGENTS.md` and updated `README.md` and docs to match current architecture.

### Fixed
- Resolved high-volume static typing issues across Coordinator, Instructor, common utilities, and test modules.
- Fixed Pylance/Pyright protocol and callable typing mismatches in output-link and async-runner related paths.
- Fixed test compatibility for Help context menu by guarding style application on menu test doubles.
- Stabilized Qt monkeypatch teardown behavior by patching Python subclasses instead of raw `PySide6` C++ classes in coordinator UI primitive tests (prevents Linux/macOS segfault in CI).
- Tightened workflow step gating semantics (including unknown-step blocking) and updated regression coverage.
- Improved score boundary handling tolerance for attainment level classification near 0/100 floating-point edges.

### Removed
- Removed stale package `__all__` exports that produced unsupported-dunder-all warnings in static analysis.
- Removed dead style/constants paths left behind after native-style cleanup.
- Removed obsolete top-level `installerscript.ps1` duplicate script.
- Removed tracked `coverage.json` artifact file.
- Removed deprecated Instructor step artifacts (`modules/instructor/steps/step1_*`, `step2_*`, `validators/*`, `workflow_controller.py`) after single-flow migration.
- Removed compatibility-only helper paths no longer used in runtime (`main_window` action aliases, legacy error-catalog resolver).
- Removed obsolete process documents superseded by `AGENTS.md` (`docs/RELEASE_PROMOTION.md`, `docs/SUPPORT_RUNBOOK.md`, `docs/UI_LAYOUT_GUARDRAILS.md`, `docs/REPOSITORY_ACTION_PLAN.md`).

### Quality
- `pyright`: `0 errors, 0 warnings`
- `pytest -q`: `484 passed`
- `ruff check .`: `passed`
- `isort --check-only --diff .`: `passed`
- `pyflakes .`: `passed`
- `bandit -q -r . -c .bandit.yaml`: `passed`
- `pip-audit`: `No known vulnerabilities found`
