# FOCUS

## About
FOCUS (Framework for Outcome Computation and Unification System) is a desktop tool for Outcome-Based Education (OBE) workflows.

It helps faculty and coordinators generate, validate, and process structured Excel workbooks for:
- course setup
- marks entry
- final CO report generation

Developed at Kalasalingam Academy of Research and Education (KARE).

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
conda run -n obe python scripts/quality_gate.py
```

## Notes
- Workbook signing/protection uses `FOCUS_WORKBOOK_PASSWORD` and related signature settings.
- Runtime signing/protection paths enforce workbook secret policy.
