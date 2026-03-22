# Time and Space Complexity Analysis - FOCUS Application

## Executive Summary

FOCUS is a desktop OBE application with these user-facing modules:
- Instructor
- Coordinator
- CO Analysis
- Help
- About

The current Instructor UX is a **single flow**:
- Download course template
- Validate uploaded course-details workbook(s)
- Generate marks-template workbook(s)

Overall complexity is dominated by workbook validation/generation and aggregation paths.

---

## 1. Architecture Coverage

### 1.1 Module Map

- Instructor UI: `modules/instructor_module.py`
  - Shared workbook naming helpers: `modules/instructor/steps/shared_workbook_ops.py`
  - Engine/domain entry points: `domain/instructor_template_engine.py`
- Coordinator UI: `modules/coordinator_module.py`
  - Step handlers: `modules/coordinator/steps/*`
  - Processing engine: `domain/coordinator_engine.py`
  - Service wrapper: `services/coordinator_workflow_service.py`
- CO Analysis UI: `modules/co_analysis_module.py`
  - Step handlers: `modules/co_analysis/steps/*`
  - Processing engine: `domain/co_analysis_engine.py`
  - Service wrapper: `services/co_analysis_workflow_service.py`
- Help UI: `modules/help_module.py`
- About UI: `modules/about_module.py`
- Plugin catalog: `modules/module_catalog.py`

### 1.2 Shared Infrastructure

- Template routing: `domain/template_strategy_router.py`
- Template strategy implementation: `domain/template_versions/course_setup_v1.py`
- Text/i18n: `common/texts/*`
- UI logging + status payloads: `common/ui_logging.py`
- Toasts: `common/toast.py`
- Workbook signing/secret policy: `common/workbook_signing.py`, `common/workbook_secret.py`
- Shared runtime helpers: `common/utils.py`, `common/module_runtime.py`

---

## 2. Instructor Module Complexity

Primary operations:
- Generate course-details template workbook
- Validate uploaded course-details workbook(s)
- Generate marks-template workbook(s)

Dominant complexity:
- Template generation: `O(R)` where `R` is emitted row count.
- Course-details validation: `O(R * C)` worst case.
- Marks-template generation: `O(Cm * S * Q)`
  - `Cm`: assessment components
  - `S`: students
  - `Q`: mark-entry columns per component

Space:
- openpyxl validation/read paths scale with workbook object graph (`~O(R * C)`).
- xlsxwriter write paths are streaming-oriented but still proportional to emitted cells.

---

## 3. Coordinator Module Complexity

Primary operations:
- Collect and validate uploaded final CO report files
- Parse and aggregate attainment data
- Generate consolidated coordinator workbook

Dominant complexity:
- File collection/validation: `O(N)` for `N` files.
- Parsing + aggregation: `O(N * S * O)`
  - `S`: students per file
  - `O`: outcomes/attainment columns
- Workbook synthesis: proportional to summary/detail rows written.

Space:
- linear memory in workbook object graphs + dedup/index maps for student/outcome keys.

---

## 4. CO Analysis, Help, and About Complexity

### 4.1 CO Analysis

Primary operations:
- Collect source files
- Validate template/signature/layout/marks constraints
- Generate analysis workbook

Dominant complexity:
- Validation and parse: `O(N * S * O)` in practical terms.
- Output generation: proportional to emitted analysis rows.

### 4.2 Help Module

- File-size-bound open/save operations: `O(F)`
- UI-level operations are effectively constant-time.

### 4.3 About Module

- Static metadata rendering: `O(1)` time, `O(1)` space (excluding packaged assets).

---

## 5. Cross-Cutting Complexity and Safety

### 5.1 I18n and Logging

- Translation lookups: average `O(1)` per key.
- Structured payload parsing/rendering: linear in message size.

### 5.2 Signature and Integrity Validation

- Signature verification is linear in payload length.
- Current release uses active signature format only.

### 5.3 Cancellation/Timeout

- Cancellation probes are constant-time checks inside long loops.
- Timeout wrappers bound runtime but do not change algorithmic order.

---

## 6. Practical Scaling Guidance

- Instructor cost grows fastest with student count x component count x mark-column width.
- Coordinator/CO Analysis cost grows with file count and per-file student/outcome volume.
- Help/About do not materially affect compute budget.

Intentional constraints:
- Do not cache workbook metadata across independent workflow runs.
- Do not reduce workbook formatting/protection/system-sheet footprint for speed-only gains.

---

## 7. Verification Status

Baseline quality checks should include:
- `pytest`
- `pyright`
- `ruff`
- `isort`
- `pyflakes`
- `bandit`
- `coverage`
- `pip-audit`

This analysis reflects the current single-flow Instructor architecture and strategy-router-based template dispatch.
