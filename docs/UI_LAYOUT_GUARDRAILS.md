# UI Layout Guardrails (Top/Footer Engine)

This project uses:
- `ModuleUIEngine` as a two-region container for module content/layout wiring.
- `MainWindow.shared_activity_frame` (`sharedInfoTabs`) as the single visible bottom activity panel.

## Incident Summary

We hit a multi-day layout regression where:
- Footer height was not consistently reapplied when replacing footer widgets.
- Top/footer stretch behavior became inconsistent.
- Hidden module footers still occupied space while shared activity mode was active.
- Visibility validation used `isVisible()` and produced false negatives before widgets were shown.

Latest stabilization (March 21, 2026):
- The user-visible bottom panel is controlled by `main_window.py` (`shared_activity_frame`) only.
- `shared_activity_frame` is now enforced with:
  - `setFixedHeight(INSTRUCTOR_INFO_TAB_FIXED_HEIGHT)`
  - `setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)`
  - zero internal layout margins/spacing for exact visual height.
- Current standard footer height constant: `INSTRUCTOR_INFO_TAB_FIXED_HEIGHT = 200`.

## Required Invariants

Keep these invariants in `common/module_ui_engine.py`:

1. Footer is fixed height.
   - Always apply fixed footer height in both:
     - engine init
     - `set_footer_widget(...)`
2. Footer size policy is fixed vertically.
   - `QSizePolicy(Expanding, Fixed)`.
3. Top region must expand to consume remaining height.
   - `QSizePolicy(Expanding, Expanding)`.
4. Root layout stretch must be explicit and stable.
   - top stretch = `1`
   - footer stretch = `0`
5. Pane visibility validation must use hidden-state semantics.
   - Use `isHidden()` checks, not `isVisible()`, to avoid pre-show false failures.
6. Shared activity footer in `MainWindow` must also be fixed-height.
   - Keep it bound to `INSTRUCTOR_INFO_TAB_FIXED_HEIGHT`.
   - Do not apply per-module or ad-hoc footer height overrides in `main_window.py` or modules.

## Shared Panel Rule

- Keep Activity Log in `MainWindow.sharedInfoTabs` common across modules.
- Keep Generated Outputs in `MainWindow.sharedInfoTabs` module-specific by reading
  each module's `get_shared_outputs_data()`.
- Instructor/Coordinator should not render their own visible footer tabs.
- Module footers may exist structurally via `ModuleUIEngine`, but must remain hidden.

## Debugging Note

Do not keep debug border styles in production engine code. If temporary borders are needed, add and remove them in the same PR.
