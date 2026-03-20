# Repository Action Plan (1-2 Member Team)

Date: 2026-03-20  
Repository: cotas  
Branch context: clementine (ahead of main)

## Purpose

Deliver the highest-value reliability improvements with minimal parallel work, clear ownership, and low burnout risk for a 1-2 person team.

## Team Model

- Keep only 1 active implementation track at a time.
- Every action has one DRI (owner) and optional reviewer (second member).
- Do not run more than 2 open PRs at once.

## Current Baseline (Verified)

- Test suite: 484 passed
- Static checks: ruff, isort, pyright passed
- Security checks: bandit and pip-audit passed
- Coverage (line): 95%
- Architecture boundary tests: passing
- Local environment aligned to Python 3.11 policy

## Scope Tiers

### Tier 1 (Must Do Now)

1. AP-01: Align local `obe` environment to Python 3.11.
2. AP-02: Enforce ruff/isort/pyright in CI.
3. AP-03: Add focused Step 2 workflow tests (critical paths only).

### Tier 2 (Do When Needed)

4. AP-04: Timeout/cancellation documentation and tests (only if incidents or confusion arise).
5. AP-05: Security ADR for workbook secret bootstrap/fallback (only if release/security review requires it).

### Tier 3 (Later / Opportunistic)

6. AP-06: Refactor one complexity hotspot when bandwidth permits.

## Sequenced Plan

### AP-01: Python Version Alignment (Day 1)

Objective: Match local development/runtime policy (Python 3.11) and remove version drift risk.

Tasks:
1. Update or recreate `obe` from `environment.yml`.
2. Verify interpreter reports Python 3.11.x.
3. Run strict quality gate and tests.

Suggested commands:

```powershell
conda env update -n obe -f environment.yml --prune
conda run -n obe python -V
conda run -n obe python scripts/quality_gate.py --mode strict
conda run -n obe python -m pytest -q
```

Acceptance criteria:
- `conda run -n obe python -V` reports `3.11.x`.
- Strict gate and test suite pass without code behavior changes.

Evidence required:
- Command output snippets in PR description.

### AP-02: CI Gate Parity (Day 1-2, after AP-01)

Objective: Prevent style/type regressions from reaching main.

Tasks:
1. Update `.github/workflows/ci.yml` to require:
   - `python -m ruff check .`
   - `python -m isort --check-only --diff .`
   - `python -m pyright`
2. Keep existing quality/security checks.
3. Run one full green matrix pipeline.

Acceptance criteria:
- CI fails on ruff/isort/pyright regressions.
- One successful full run across windows/macos/ubuntu.

Evidence required:
- CI run link in PR description.

### AP-03: Focused Reliability Tests (Week 1)

Objective: Improve confidence in highest-risk instructor Step 2 flows with limited test scope.

Primary target:
- `modules/instructor/steps/step2_filled_marks_and_final_report.py`

Secondary targets (only if time remains):
- `modules/instructor_module.py`
- `common/drag_drop_file_widget.py`

Tasks:
1. Add tests for Step 2 cancellation and final-report state transitions.
2. Add only the most critical drag-drop edge tests if AP-03 core scope is complete.

Acceptance criteria:
- Step 2 target coverage reaches at least 85%, or PR documents why 85% is not practical.
- Full suite remains green (no regression from 474-pass baseline behavior).

Evidence required:
- New test files/sections listed in PR.
- Coverage summary (before/after for targeted files).

## Deferred Actions (Trigger-Based)

### AP-04: Timeout/Cancellation Semantics

Trigger:
- Repeated confusion, flaky behavior, or timeout-related incidents.

Deliverable:
- Short docs note + targeted tests for cancellation propagation.

### AP-05: Security ADR

Trigger:
- Security/release audit requests formal decision record.

Deliverable:
- ADR with threat model, accepted risk, and operational checklist.

## Opportunistic Action

### AP-06: Single Hotspot Refactor

Rule:
- Refactor only one hotspot at a time, with tests before and after.

Candidate files:
- `domain/instructor_report_engine.py`
- `domain/coordinator_engine.py`
- `domain/template_versions/course_setup_v1.py`

Guardrails (mandatory):
- Do not change workbook output structure/formatting/protection behavior.
- Do not cache workbook metadata across Step 1 and Step 2 runs.

## Lightweight Tracking Template

| ID | Status | DRI | Reviewer | Start | Due | Evidence |
|---|---|---|---|---|---|---|
| AP-01 | Completed | Local | TBD | 2026-03-20 | 2026-03-20 | `conda run -n obe python -V` = `3.11.15` |
| AP-02 | Completed | Local | TBD | 2026-03-20 | 2026-03-20 | CI workflow enforces `ruff`, `isort --check-only`, `pyright`; local checks green |
| AP-03 | Completed | Local | TBD | 2026-03-20 | 2026-03-20 | `modules/instructor/steps/step2_filled_marks_and_final_report.py` coverage >= 85% with focused tests |
| AP-04 | Not Triggered | Local | TBD | 2026-03-20 |  | No repeated timeout/cancellation incidents observed |
| AP-05 | Not Triggered | Local | TBD | 2026-03-20 |  | No security/release audit request for ADR in current cycle |
| AP-06 | Completed | Local | TBD | 2026-03-20 | 2026-03-20 | Dedup store now uses practical SQLite threshold helper in `domain/coordinator_engine.py` |

## Definition of Done (Small-Team Version)

This plan is complete when:
1. AP-01 and AP-02 are merged.
2. AP-03 is merged with focused Step 2 coverage improvement and green tests.
3. Deferred items (AP-04/AP-05) are either completed or explicitly marked "not triggered".
4. CI and strict quality gates remain green.
