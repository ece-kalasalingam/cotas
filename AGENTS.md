# Project Agent Instructions

## Python environment

- Use the `obe` conda environment for all Python checks and test commands in this project.
- Preferred command prefix: `conda run -n obe python -m ...`

## Examples

- Lint: `conda run -n obe python -m pyflakes .`
- Tests: `conda run -n obe python -m pytest -q`
- Compile check: `conda run -n obe python -m py_compile main.py`
