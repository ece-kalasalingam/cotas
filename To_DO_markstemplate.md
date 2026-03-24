Here are the remaining steps for marks_template.py (V2), excluding what we already completed:

Step1: Add batch generation contract in router
File: [template_strategy_router.py](g:/JP/Git Repositories/cotas/domain/template_strategy_router.py)
Add generate_workbooks(...) entrypoint (parallel to validate_workbooks(...)) that accepts workbook list + output dir/map + context and dispatches to strategy.
Keep router thin: resolve strategy, dispatch, fail-fast only.
Step2: Add batch generation method in V2 strategy
File: [course_setup_v2.py](g:/JP/Git Repositories/cotas/domain/template_versions/course_setup_v2.py)
Add generate_workbooks(...) for workbook_kind="marks_template".
This method should own the per-workbook loop (or call a V2 process helper that owns it).
Return structured result: generated, failed, skipped, total, per-file reason/status.
Step3: Implement V2 marks-template core (single workbook)
File: [marks_template.py](g:/JP/Git Repositories/cotas/domain/template_versions/course_setup_v2_impl/marks_template.py)
Replace placeholder with production code for generate_marks_template_from_course_details(...).
Include: open source workbook, read validated template id, extract required context, generate output workbook, preserve required system sheets/signing behavior, atomic write, cancellation checks, typed validation errors.
Step 4: Add V2 marks-template batch process helper
File: [marks_template.py](g:/JP/Git Repositories/cotas/domain/template_versions/course_setup_v2_impl/marks_template.py)
Add generate_marks_templates_from_course_details_batch(...) (or similar) called by strategy batch method.
Input: list of source paths + output directory/path planner + token.
Ownership: file-name generation policy per workbook kind stays here (not module/router).
Step 5: Delete legacy commented code in V2 marks-template impl
File: [marks_template.py](g:/JP/Git Repositories/cotas/domain/template_versions/course_setup_v2_impl/marks_template.py)
Remove fully commented old block and placeholder __all__.
Keep only active, testable V2 implementation.
Step 6: Modify Instructor module to stop per-file generation loop
File: [instructor_module.py](g:/JP/Git Repositories/cotas/modules/instructor_module.py)
Keep UI duties only: collect selected files + output dir/path, then call router batch generation once.
Remove direct per-file generate_workbook(...) loop and local per-file generation orchestration.
Step 7: Update output naming flow
Files: V2 strategy/impl + instructor module
Module should pass only output directory/path intent.
V2 impl should generate final filenames for each workbook (business naming rules).
Step 8: Update protocol/types and exports
Files: router + strategy typing surfaces
Add typed contracts for batch generation return payload.
Ensure __all__ and call sites use new API only.
Step 9: Delete any now-dead marks-template utility paths
Files: router/strategy/module where old single-file glue became unused after batch refactor.
Remove dead helpers/imports to keep attack surface and complexity low.
Step 10: Verify with checks
Run compile/lint for touched files and one runtime smoke path from instructor module.
Minimum: py_compile on router, strategy, marks_template impl, instructor module.
Confirm multi-file generation returns correct aggregate status and files are written with expected names.

