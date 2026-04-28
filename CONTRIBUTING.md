# Contributing to FOCUS

Thanks for contributing to FOCUS.
This guide keeps changes consistent, reviewable, and safe for workbook integrity.

## Project Scope

FOCUS is a desktop OBE workflow tool built with Python and PySide6.
Core guarantees include:

- workbook structure and formatting compatibility
- validation and protection integrity
- strict routing between template versions

Before changing behavior, check existing policy docs:

- `docs/PR_CHECKLIST.md`
- `docs/QUALITY_GATE.md`
- `SECURITY.md`

## Development Setup

1. Create/update the conda environment:
   - `conda env update -f environment.yml`
2. Run the app:
   - `conda run -n obe python main.py`
3. Install dev dependencies if needed:
   - `pip install -r requirements-dev.txt`

## Branch and Commit Workflow

1. Create a feature branch from `main`.
2. Keep commits focused and logically grouped.
3. Use clear commit messages that describe intent and impact.
4. Avoid mixing refactors with behavior changes unless tightly coupled.

## Coding Expectations

- Keep module/UI layers thin and template-agnostic.
- Route template-specific behavior through the strategy/router paths.
- Do not introduce cross-run workbook metadata cache.
- Preserve typed exceptions and error-catalog mappings.
- Keep workbook protection/signing paths intact unless explicitly approved.

## Testing and Quality Checks

Run at least fast checks during iteration:

- `conda run -n obe python scripts/quality_gate.py --mode fast`
- `conda run --no-capture-output -n obe python -m pytest -q`

Before opening a PR, run the strict gate:

- `conda run -n obe python scripts/quality_gate.py --mode strict`

For release-level validation, follow the full command set in:

- `docs/QUALITY_GATE.md`

## Pull Request Guidelines

Use small, reviewable PRs with:

- clear summary of what changed and why
- risk notes (especially around workbook compatibility)
- test evidence (targeted tests + full suite status)
- screenshots/log snippets for UI-visible changes when relevant

Confirm every item in:

- `docs/PR_CHECKLIST.md`

## Security and Secrets

- Never commit real secrets, generated local secrets, or private credentials.
- Use sample config files (for example `cip_config.sample.json`) for shared examples.
- Follow `SECURITY.md` for vulnerability reporting and security expectations.

## Documentation Updates

When behavior changes, update docs in the same PR:

- `README.md` for user-visible workflows
- `docs/QUALITY_GATE.md` for executable checks
- `docs/PR_CHECKLIST.md` for review/policy guardrails
