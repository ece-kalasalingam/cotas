# UI Layout Guardrails (Top/Footer Engine)

This project now uses `ModuleUIEngine` as a two-region container:
- Top region (`top_widget`) for module content.
- Footer region (`footer_widget`) for module-local tabs/log/output panels.

## Incident Summary

We hit a multi-day layout regression where:
- Footer height was not consistently reapplied when replacing footer widgets.
- Top/footer stretch behavior became inconsistent.
- Hidden module footers still occupied space while shared activity mode was active.
- Visibility validation used `isVisible()` and produced false negatives before widgets were shown.

Latest stabilization (March 21, 2026):
- The user-visible bottom panel height is controlled by `main_window.py` (`shared_activity_frame`),
  not only by module-local footers.
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

## Module Usage Rule

Modules with a local footer panel must collapse engine footer in shared-activity mode:

```python
def set_shared_activity_log_mode(self, enabled: bool) -> None:
    self.info_tabs.setVisible(not enabled)
    self._ui_engine.set_footer_visible(not enabled)
```

This is required so the top region stretches fully when shared logs are shown.

## Debugging Note

Do not keep debug border styles in production engine code. If temporary borders are needed, add and remove them in the same PR.
