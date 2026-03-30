# PR Checklist

## AGENTS Guardrails

- [ ] No cross-run workbook metadata cache introduced.
- [ ] No V1 business logic changes; all validator/orchestration edits are V2-only paths.
- [ ] Template routing/branching discipline preserved (no collapse into unbranched generic logic).
- [ ] Module/UI layer remains thin and template-agnostic; dispatch stays through router entrypoints.
- [ ] Typed exceptions and error-catalog mapping behavior unchanged.

## Compatibility and Integrity

- [ ] Workbook output compatibility preserved: structure, formatting, hidden/system sheets, and protection behavior unchanged unless separately approved.
- [ ] Marks validator integrity checks are not weakened (read-only trust gate + manifest/schema validation intact).
- [ ] Course cohort mismatch and duplicate-section handling semantics unchanged.

## Validation Evidence

- [ ] Targeted validator/refactor tests passed.
- [ ] Full test suite passed in `obe`.
- [ ] Quality-gate static checks passed (see `docs/QUALITY_GATE.md`).
- [ ] Perf comparison recorded (median/p95 + peak memory) for baseline vs refactor before optimization claims.
