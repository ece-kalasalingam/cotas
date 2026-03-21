# Project Agent Instructions

## Python environment

- Use the `obe` conda environment for all Python checks and test commands in this project.
- Preferred command prefix: `conda run -n obe python -m ...`

## Examples

- Lint: `conda run -n obe python -m pyflakes .`
- Tests: `conda run -n obe python -m pytest -q`
- Compile check: `conda run -n obe python -m py_compile main.py`

## Complexity Guardrails (Do Not "Optimize" These)

- Do not cache parsed workbook metadata across Instructor Step 1 and Step 2 runs.
  Step 1 and Step 2 can be executed for different courses, by different users, and on different days on shared systems.
- Do not trim workbook output structure/formatting/protection.
  Current formatting, hidden/system sheets, and protection behavior are part of required output compatibility for this release.

## Compulsory UI Content Constraints

- Keep app title sourced from translations (`app.main_window_title`), not hardcoded constants.
- For non-English locales, title should use phonetic English transliteration in that script.
- Keep app subtitle sourced from translations (`about.subtitle`), not hardcoded constants.
- For non-English locales, subtitle should use phonetic English transliteration in that script
  (English pronunciation written in the selected language script).
- In About description (`about.description`) for non-English locales, keep key product terms
  such as app name and "Course Outcome (CO)" in transliterated form.

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
- Modules with visible footer should collapse footer in shared activity mode using
  `self._ui_engine.set_footer_visible(not enabled)` so top expands fully.
- Do not keep temporary debug border styling in engine code after debugging.
- Footer height source of truth must be `INSTRUCTOR_INFO_TAB_FIXED_HEIGHT` in `common/constants.py`.
- The global visible footer (`shared_activity_frame` in `main_window.py`) must stay fixed-height
  and use `QSizePolicy(Expanding, Fixed)` with zero inner margins/spacing.
- Do not hardcode footer heights in modules or in `MainWindow`; update only the constant.
