# Project TODO List (Machine-Friendly)

This file is human-readable. Use `TODO-LIST.json` as source-of-truth for automation.

## Current Sprint: V2 Marks Template

- [x] `MT-001` Add router batch generation entrypoint `generate_workbooks(...)`.
  - Files: `domain/template_strategy_router.py`
- [x] `MT-002` Add V2 strategy batch generation `generate_workbooks(...)` for `workbook_kind="marks_template"` with structured aggregate result.
  - Files: `domain/template_versions/course_setup_v2.py`
- [ ] `MT-003` Implement production-ready single-workbook generator in V2 marks template impl.
  - Files: `domain/template_versions/course_setup_v2_impl/marks_template.py`
- [ ] `MT-004` Add V2 batch process helper to own per-workbook iteration/orchestration.
  - Files: `domain/template_versions/course_setup_v2_impl/marks_template.py`
- [ ] `MT-005` Enforce workbook naming policy in V2 strategy/impl only (no module/router naming rules).
  - Files: `domain/template_versions/course_setup_v2.py`, `domain/template_versions/course_setup_v2_impl/marks_template.py`, `modules/instructor_module.py`
- [ ] `MT-006` Remove legacy commented code and placeholder exports in V2 marks template impl.
  - Files: `domain/template_versions/course_setup_v2_impl/marks_template.py`
- [ ] `MT-007` Refactor Instructor marks-template flow to single router batch call (remove per-file generation loop).
  - Files: `modules/instructor_module.py`
- [ ] `MT-008` Remove dead marks-template helper/import paths after batch refactor.
  - Files: `domain/template_strategy_router.py`, `domain/template_versions/course_setup_v2.py`, `modules/instructor_module.py`
- [ ] `MT-009` Update typing/contracts/exports for batch generation API + result schema.
  - Files: `domain/template_strategy_router.py`, `domain/template_versions/course_setup_v2.py`
- [ ] `MT-010` Run compile/smoke checks in `obe` env for all touched files.
  - Commands: `conda run -n obe python -m py_compile ...`
