# Project Agent Instructions

## Python environment

- Use the `obe` conda environment for all Python checks and test commands in this project.
- Preferred command prefix: `conda run -n obe python -m ...`

## Examples

- Lint: `conda run -n obe python -m pyflakes .`
- Tests: `conda run -n obe python -m pytest -q`
- Compile check: `conda run -n obe python -m py_compile main.py`

## Complexity Guardrails (Do Not "Optimize" These)

- Do not cache parsed workbook metadata across independent Instructor runs.
  Runs can be executed for different courses, by different users, and on different days on shared systems.
- Do not trim workbook output structure/formatting/protection.
  Current formatting, hidden/system sheets, and protection behavior are part of required output compatibility for this release.

## Compulsory Engineering Guardrail (DRY + SSOT + Reuse + Complexity)

- This is mandatory for all modules and templates: follow DRY (Don't Repeat Yourself), keep a Single Source of Truth (SSOT), and prefer reusable shared logic over duplicated logic.
- If business logic already exists in a shared/helper layer, reuse or extend that implementation; do not re-implement equivalent logic in another file/module.
- New business rules must be added to the authoritative shared location first, then consumed by callers.
- Avoid parallel implementations of the same validation/parsing/generation rules across files; keep one authoritative implementation and make other layers orchestration-thin.
- Changes must target least practical time and space complexity for the workflow scale in this project; avoid unnecessary passes, repeated parsing, and redundant allocations.
- If a tradeoff is required (readability vs micro-optimization), keep code readable while still avoiding avoidable asymptotic or repeated-work regressions.

## Compulsory Template Evolution Guardrail

- All business logic that can vary by template version must support template-specific behavioral evolution through explicit template-id branching or strategy dispatch.
- For currently supported template ids, keep behavior in clearly separated template branches (even if two branches temporarily share the same implementation).
- This branching/dispatch requirement is compulsory for shared/common code paths used by more than one template id; do not keep multi-template business logic in a single unbranched path.
- If a business rule is strictly template-specific and is used by only one template id, separate branching is not required for that rule.
- For new template versions, add a new branch/strategy path instead of editing existing template behavior in-place.
- Keep module-layer code template-agnostic; template branching must live in shared router/strategy/semantics layers.

## Compulsory V2 Migration Policy (Current Release Track)

- `COURSE_SETUP_V2` is the active target architecture for this unreleased app; all ongoing refactors must migrate business logic to V2 paths.
- Migration must be phased:
  - first refactor/route logic to V2
  - then verify runtime flows and validations on V2
  - only after that, prune dead V1 code
- During migration, do not add new business logic to V1 files unless explicitly required for a temporary bridge; prefer V2-first implementation.
- Keep pruning of V1 code intentional and evidence-based (remove only when references are eliminated and equivalent V2 behavior exists).

## Single Source Of Truth Guardrail (Sheet Configuration)

- Any sheet configuration that is not strictly template-version implementation detail must be centralized in
  `common/registry.py` + `common/sheet_schema.py`.
- Do not introduce or maintain duplicate sheet-structure truth points in module/domain/service code.
- New sheet additions, header changes, validation ranges, and structural metadata must be declared in the common schema/registry layer first.

## Single Source Of Truth Guardrail (Path Identity Helper)

- Filesystem path identity normalization must use `common/utils.py::canonical_path_key` as the single source of truth.
- Do not introduce module/domain/service-local path key helpers (for example `_path_key`) that duplicate canonicalization behavior.

## Single Source Of Truth Guardrail (CO Direct/Indirect Sheet Generation)

- `domain/template_versions/course_setup_v2_impl/co_report_sheet_generator.py` is the single source of truth for generating `COx_Direct` and `COx_Indirect` sheets.
- Do not add or keep duplicate direct/indirect sheet write logic in `instructor_report_engine.py`, `coordinator_engine.py`, `co_analysis_engine.py`, or module-layer code.
- Always use shared helpers from `co_report_sheet_generator.py`:
  - `write_co_outcome_sheets(...)` for writing both sheets
  - `co_direct_sheet_name(...)` and `co_indirect_sheet_name(...)` for sheet naming
- Template-specific behavior must be expressed through data/adapters (strategy or mapping inputs), not by cloning generator logic.
- If CO sheet layout/headers/formulas/metadata behavior changes, update `co_report_sheet_generator.py` first and keep downstream callers thin.

## Single Source Of Truth Guardrail (Excel Layout Helpers)

- Keep shared Excel layout/styling/protection/copy helpers centralized in `common/excel_sheet_layout.py`.
- Do not introduce standalone duplicate layout helper implementations in engine/module files for:
  - xlsxwriter page/layout/protection routines
  - openpyxl sheet protection and style-copy behavior
  - column width sampling helpers and column-name utilities
- If a workflow needs slightly different layout behavior, extend `common/excel_sheet_layout.py` and reuse it from callers.
- Engine-local helper wrappers are allowed only as thin delegates to shared helpers, without divergent logic.

## Workbook Secret/Protection Guardrail

- Workbook sheet protection policy must be enforced through shared helpers in `common/excel_sheet_layout.py`:
  - `protect_openpyxl_sheet(...)`
  - `protect_xlsxwriter_sheet(...)`
- Do not call `ensure_workbook_secret_policy()` or `get_workbook_password()` directly from module/domain/service workbook-generation code.
- Direct `workbook_secret` usage is allowed only in:
  - `common/excel_sheet_layout.py` (sheet protection primitives)
  - `common/workbook_signing.py` (signature primitives)
  - startup/bootstrap policy check in `main.py`
- Keep `workbook.security.lockStructure` behavior where required, but do not duplicate password/policy wiring outside shared helpers.

## Single Source Of Truth Guardrail (Xlsxwriter Format Bundles)

- Shared setup/workbook format-bundle construction must use `common/excel_sheet_layout.py::build_template_xlsxwriter_formats(..., template_id=...)`.
- Do not duplicate header/body/body_wrap/body_center/column_wrap format-bundle assembly logic in engine files.
- Engine `_xlsxwriter_formats(...)` helpers (if retained) must be thin delegates to shared helpers and local cache-attr wiring only.
- If a module needs additional reusable format variants, extend `build_template_xlsxwriter_formats(...)` (or add a shared adjacent helper) instead of cloning per-engine builders.

## Template-ID Styling Rule

- Course template generation may use current `ID_COURSE_SETUP` as the active template id.
- For uploaded/source workbooks (marks/final-report/coordinator/co-analysis), styling and format-bundle resolution must use the template id read from that workbook `SYSTEM_HASH`.
- Do not resolve uploaded-workbook styles via global/current template constants.
- Shared helpers must be template-id aware:
  - use `common/excel_sheet_layout.py::style_registry_for_template(template_id)`
  - use `common/excel_sheet_layout.py::build_template_xlsxwriter_formats(..., template_id=...)`

## Template Strategy Routing Guardrail

- Module-layer code must stay template-agnostic: collect inputs, read `SYSTEM_HASH` template id (except course-template generation), and delegate.
- Route worksheet generation/validation operations via `domain/template_strategy_router.py`; do not branch on template id inside modules.
- Route workbook-generation entrypoints through `domain/template_strategy_router.py::generate_workbook(...)`.
- `generate_workbook(...)` is the shared router entrypoint intended for workbook generation across modules; prefer extending strategy handlers over adding module-local generators.
- Router responsibility is limited to strategy resolution + operation dispatch + fail-fast contract checks.
- Template-specific orchestration must live in `domain/template_versions/<template_id>.py` strategy classes.
- Strategy classes may call shared/common helpers or template-specific helpers, but modules must not import template-version modules directly.

## Module To Workbook Flow Guardrail

- Workbook generation and workbook validation must follow this layered flow:
  - Module UI/engine layer (input collection only; no template branching)
  - `domain/template_strategy_router.py` shared entrypoint
  - Resolved template strategy class in `domain/template_versions/<template_id>.py`
  - Template-version implementation file(s) under that template version (shared helpers allowed)
- For generation, modules must call router `generate_workbook(...)` only.
- For validation, modules must call router `validate_workbook(...)` or `validate_workbooks(...)` only.
- Template id for uploaded/source workbooks must be read from `SYSTEM_HASH` and routed dynamically.
- Template id must not be hardcoded in module-layer logic for uploaded/source workbook operations.
- For multi-file workbook generation (for example marks template):
  - Module must use directory-selection mode when generating multiple outputs.
  - File naming must be resolved in template strategy/template implementation workflow code (workbook-kind aware); router must not own naming logic.
  - Module code must not define business naming rules; it may only pass user-selected output directory/path inputs.
- If a workflow kind (for example `marks_template`) is newly added, extend router/strategy/impl layers; do not bypass router from modules.

## Batch Iteration Guardrail

- For multi-workbook operations, module code must pass workbook collections to router entrypoints and stay orchestration-thin.
- Router remains template-id resolver and operation dispatcher only; router must not implement template-specific workbook iteration logic.
- Template strategy (or a template-local work-process helper called by strategy) is the authoritative location for per-workbook iteration/orchestration.
- Validation implementation files should prioritize single-workbook validation entrypoints (for example `validate_course_details_rules(workbook)`).
- Batch validation/generation helpers may call single-workbook entrypoints in a loop, but business-rule ownership remains in single-workbook validators.
- Aggregate-all-errors behavior for a workbook must be implemented in template validation code, not in module or router layers.

## Workbook Output Save/Collision Guardrail

- Single-workbook generation flows must use native save-file dialog UX (`getSaveFileName`) in module/UI layers.
- Multi-workbook generation flows must use directory-selection UX and batch generation for the first pass.
- Do not duplicate per-module collision parsing/resolution logic for multi-workbook outputs.
- Reuse `common/workbook_output_resolution.py` as the shared source of truth for:
  - extracting overwrite collisions from generation results
  - resolving collision actions (bulk overwrite vs per-file output selection)
- For multi-workbook collision retries, modules must rerun generation only for collided/selected source files and pass per-source output overrides through router context.
- Prefer `O(N + K)` collision handling (`N` total files, `K` collided files); avoid `O(2N)` pre-parse-only planning passes.

## Single Source Of Truth Guardrail (Validation Issues)

- Validation issue semantics (code -> category/severity/i18n/default) must be centralized in `common/error_catalog.py`.
- UI/toast/log rendering for validation failures must resolve through `common/error_catalog.py::resolve_validation_issue(...)` (directly or via shared utility wrappers).
- Prefer raising `ValidationError` through `common/error_catalog.py::validation_error_from_key(...)` so translation key, category, severity, and code stay centralized.
- Shared cross-module validation failures (dependency/system-hash/layout/template/workbook/mark rules) must reuse shared codes; do not create module-prefixed duplicates for the same business failure.
- Shared translation keys must use the generic `validation.*` namespace (not module-specific namespaces) when the same failure can occur in multiple modules.
- Avoid direct `raise ValidationError(t(\"...\"))` patterns in module/domain code.
- `common/utils.py::log_process_message(...)` is the shared process-status path and must remain aligned with catalog-driven validation resolution.

## Exception Contract Guardrail

- Use typed app exceptions from `common/exceptions.py` across runtime code.
- Validation/business-rule failures must raise `ValidationError` (prefer via `validation_error_from_key(...)` from `common/error_catalog.py`).
- Static config/contract failures must raise `ConfigurationError`.
- Unexpected internal/runtime failures should raise `AppSystemError`.
- Cancellation paths must use `JobCancelledError`; do not invent module-local cancellation exception classes.
- Do not raise generic `ValueError`/`RuntimeError`/`KeyError` for user-facing validation paths in module/domain/service runtime code.

## Job Contract Guardrail

- `common/jobs.py` is the shared source of truth for job metadata and cancellation primitives.
- Use `JobContext` for workflow/service execution context (`job_id`, `step_id`, language, payload).
- Use `CancellationToken` for cancellable work and call `raise_if_cancelled()` in long-running loops.
- Use `generate_job_id()` from `common/jobs.py` when explicit job IDs are needed.
- Do not introduce duplicate per-module token/context/job-id helpers.

## Module Message Guardrail

- Module/MainWindow UI messaging must use `common/module_messages.py` (directly or via `common/module_runtime.py`).
- Preferred dispatch path is unified routing via:
  - `notify_message(...)`
  - `notify_message_key(...)`
  where `channels` controls target surfaces (`status`, `toast`, `activity_log`) in any combination.
- Strict rule: any `status`/`activity_log` message must be an i18n payload (for example via
  `notify_message_key(...)` or `build_status_message(...)`); plain text is disallowed.
- `notify_message(...)` may carry plain text only for `channels=("toast",)` paths.
- `publish_status(message)` and raw `append_user_log(message)` wrappers are disallowed for module code;
  use `publish_status_key(...)` / `notify_message_key(...)` instead.
- Runtime contract is strict-fail: passing plain text into `status`/`activity_log` paths must raise
  `ConfigurationError` instead of silently accepting.
- `publish_status_key(...)` remains a valid compatibility wrapper.
- Modules that render logs must keep compatible message state fields:
  - `status_changed`
  - `_user_log_entries`
  - `_ui_log_handler`
  - `user_log_view`
- Avoid direct `emit_user_status(...)` calls from module code when the same behavior can be routed via `module_messages`.
- Keep log payload creation/rerender through `module_messages` helpers:
  - `default_messages_namespace(...)`
  - `build_status_message(...)`
  - `resolve_status_message(...)`
  - `append_user_log(...)` / `rerender_user_log(...)`
- Use centralized toast helpers from `module_messages` for UI toasts:
  - `notify_message(..., channels=(\"toast\", ...))`
  - `notify_message_key(..., channels=(\"toast\", ...))`
  - `show_toast_key(...)`
  - `show_toast_plain(...)`
- Unified feedback-toast rule (Instructor + CO Analysis):
  - Validator flows must emit summary toasts through `ModuleRuntime.emit_validation_batch_feedback(...)` only.
  - Workbook-generation flows must emit summary toasts through `ModuleRuntime.emit_workbook_generation_feedback(...)` only.
  - Do not add module-specific success/warning generation toasts that duplicate these shared feedback summaries.
- Do not import `common/ui_logging.py` directly in module/runtime UI files; `ui_logging` is an internal dependency of `module_messages`.
- Do not import `common/toast.py` directly in module/runtime UI files; route toasts through `module_messages`.
- Keep startup/crash bootstrap paths (for example `main.py`) as exceptions where direct low-level wiring may remain.
- Shared ready message key for activity log bootstrap must be generic (`activity.log.ready`), not module-specific.

## Compulsory UI Content Constraints

- Keep app title sourced from translations (`app.main_window_title`), not hardcoded constants.
- For non-English locales, title should use phonetic English transliteration in that script.
- Keep app subtitle sourced from translations (`about.subtitle`), not hardcoded constants.
- For non-English locales, subtitle should use phonetic English transliteration in that script
  (English pronunciation written in the selected language script).
- In About description (`about.description`) for non-English locales, keep key product terms
  such as app name and "Course Outcome (CO)" in transliterated form.

## Qt Translation Guardrail

- UI translation must use Qt's native translator flow via `QTranslator` with locale-based `.qm` loading.
- Runtime translation access must go through `common/i18n/__init__.py` (`t`, `set_language`, `get_language`, `get_available_languages`).
- Do not reintroduce `common/texts` catalogs or non-Qt in-memory translation registries.
- Translation source of truth must be Qt TS catalogs under `common/i18n/` (for example `obe_en_US.ts`, `obe_hi_IN.ts`, `obe_ta_IN.ts`, `obe_te_IN.ts`).
- Compiled catalogs must be generated as `.qm` files in `common/i18n/` and loaded at runtime by `QTranslator`.
- Use `scripts/build_qt_translations.py` to compile `.ts` to `.qm`; do not add alternate translation-build scripts.
- Keep translation scope limited to UI/UX strings only; do not translate workbook data, generated Excel cell values, sheet names, filenames, or filesystem paths unless explicitly required by a separate spec.

## Runtime Asset Path Guardrail

- Runtime-loaded bundled assets (for example report images used by `python-docx`) must resolve via `common/utils.py::resource_path(...)`.
- Do not read packaged assets through CWD-relative paths such as `Path("assets") / ...`; this is invalid for frozen/exe runs.
- Any asset needed by frozen runtime logic must be present in PyInstaller `datas` and resolved through `resource_path(...)` so behavior is consistent in dev and packaged modes.

## Module UI Engine Guardrail

- `common/module_ui_engine.py` must remain generic and blackbox-only.
- It should only compose panes and manage visibility/container wiring.
- Do not add module-specific styling behavior in the engine:
  no color/background changes, no theme overrides, no content formatting logic.
- Apply styling (if any) only in concrete modules, and keep it minimal.
- Preserve `ModuleUIEngine` top/footer invariants:
  footer height must be fixed at init and when replacing footer widgets;
  top must expand to fill remaining space;
  root stretch should remain top=`1`, footer=`0`;
  pane-visibility validation should rely on hidden-state semantics (`isHidden`).
- Shared activity tabs are centralized in `MainWindow` (`sharedInfoTabs`).
- `MainWindow.shared_activity_frame` is the only visible bottom activity panel; module-local visible footer panels are not allowed.
- Keep Activity Log shared/common in `MainWindow.sharedInfoTabs` across modules.
- Keep Generated Outputs module-specific through each module `get_shared_outputs_data()` payload.
- Instructor/Coordinator must not render a separate visible footer tab panel.
- In module shared-activity mode hooks, keep module footer hidden so top expands fully.
- Do not keep temporary debug border styling in engine code after debugging.
- Footer height source of truth must be `INSTRUCTOR_INFO_TAB_FIXED_HEIGHT` in `common/constants.py`.
- The global visible footer (`shared_activity_frame` in `main_window.py`) must stay fixed-height
  and use `QSizePolicy(Expanding, Fixed)` with zero inner margins/spacing.
- Do not hardcode footer heights in modules or in `MainWindow`; update only the constant.

## Footer Log I18N Guardrail

- Incident fixed: some drag/drop footer log lines were emitted as plain localized strings, so they did not
  retranslate on language switch.
- Rule: any user-facing status/log message that must remain translatable after language change must be emitted
  as an i18n payload, not as `t(...)` plain text.
- This is mandatory across all modules and step/helper submodules (not only Instructor/CO Analysis).
- In modules, prefer key-based logging:
  - use `self._runtime.notify_message_key(..., channels=(\"status\", \"activity_log\"))` for combined status+activity log emission.
  - use `self._runtime.publish_status_key(...)` (or module helper wrappers like `_publish_status_key(...)`)
  - do not use raw `notify_message(...)` for `status`/`activity_log` channels.
  - avoid `self._publish_status(t("..."))` for translatable lifecycle/status lines.
- For helpers/widgets/step modules:
  - emit i18n payloads through `module_messages`/`ModuleRuntime` abstractions.
  - do not pre-localize and store message text if the line is expected to retranslate later.
- Footer/shared activity rendering expectations:
  - shared footer log rerender happens in `main_window.py`
  - rerender path should resolve from stored i18n payload/raw message, not only from pre-localized text.
- CI guardrails:
  - AST policy tests must fail if module code uses raw `status`/`activity_log` emission
    (`notify_message(...)` with those channels, `publish_status(...)`, `append_user_log(...)`).
  - Locale coverage tests must fail on missing translation keys or key-echo translations for module-emitted
    status/activity i18n keys in enabled locales.

## Release Entry Gate

- Before any release promotion, all entry checks must pass:
  - CI green across configured OS matrix jobs.
  - Security checks pass.
  - Quality checks pass.
  - Artifact checksum manifest generated.
  - Module catalog validation complete (`modules/module_catalog.py` contains expected modules and labels).
- Run the executable gate checklist from `docs/QUALITY_GATE.md` as the source of truth for release commands.
- Dependency-audit exception tracking: re-check `GHSA-58qw-9mgm-455v` on every release; remove `pip-audit --ignore-vuln GHSA-58qw-9mgm-455v` immediately once an upstream fix is published and available.
- Promotion flow must follow immutable-artifact practice:
  - build in `dev`, verify/sign in `stage`, promote same verified artifact to `prod` without rebuild.
- Release metadata requirements:
  - tag release commit
  - attach checksum/manifest artifacts with the release tag.

## GitHub Contributor Fetch Guardrail

- `common/get_contributors.py` must use the GitHub GraphQL API (`_from_graphql_authors`) to collect all commit authors, including those credited via `Co-authored-by` trailers in squash-merge commits.
- Do not revert to REST-only fetching (`/repos/{owner}/{repo}/contributors` alone), as that endpoint only sees commit authors on the default branch and misses co-authors whose contributions were squash-merged.
- Do not re-introduce `_merged_pr_commit_logins_via_gh` or any helper that walks merged PR commit history via the `gh` CLI; that approach pulls in inactive/renamed GitHub accounts and produces incorrect output.
- The canonical source of truth for the contributors list is the GraphQL `commit.history.nodes[].authors` field, which GitHub itself uses to power the Contributors graph on the web UI.
- Token resolution must follow the priority chain: `GH_TOKEN` env var → `GITHUB_TOKEN` env var → `gh auth token` CLI fallback, so the script works in both CI and local environments without modification.
