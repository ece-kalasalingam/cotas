# FOCUS

## About
FOCUS (Framework for Outcome Computation and Unification System) is a desktop tool for Outcome-Based Education (OBE) workflows.

It helps faculty and coordinators generate, validate, and process structured Excel workbooks for:
- course setup
- marks entry
- final CO report generation

Developed at Kalasalingam Academy of Research and Education (KARE).

## Modules
- `Instructor`: two-step course workflow (template generation/validation, then final report generation).
- `Coordinator`: bulk final-report intake and attainment consolidation.
- `Help`: packaged PDF guidance, export/open actions.
- `About`: version/build metadata and app identity panel.

## Instructor Workflow (Current)
- `Step 1`: Generate and validate course details, then generate marks templates.
- `Step 2`: Upload/validate filled marks and generate final CO reports.
- The instructor UI intentionally uses a two-step model only.

## Tech Stack
- Python
- PySide6
- openpyxl
- xlsxwriter
- pyqtdarktheme

## Run (Conda)
Use the project environment:

```powershell
conda run -n obe python main.py
```

## Quality Checks
```powershell
conda run -n obe python scripts/quality_gate.py --mode fast
conda run -n obe python -m ruff check .
conda run -n obe python -m isort --check-only --diff .
conda run -n obe python -m pyright
conda run -n obe python -m pyflakes .
conda run --no-capture-output -n obe python -m pytest -q
conda run -n obe python -m coverage run -m pytest -q
conda run -n obe python -m coverage report -m
conda run -n obe python -m bandit -q -r . -c .bandit.yaml
conda run -n obe python -m pip_audit --cache-dir .pip_audit_cache -r requirements.txt -r requirements-dev.txt
```

## Notes
- Workbook signing/protection uses an app-managed local secret store and signature settings.
- Runtime signing/protection paths enforce workbook secret policy.
- Legacy signature compatibility paths were removed; current version expects active versioned signatures.
