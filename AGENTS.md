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
