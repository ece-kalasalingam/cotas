# Time and Space Complexity Analysis - FOCUS Application

## Executive Summary

FOCUS is a desktop OBE application organized into five user-facing modules:
- Instructor
- Coordinator
- PO Analysis
- Help
- About

The current Instructor UX is intentionally **two-step**:
- Step 1: Course details + marks template preparation
- Step 2: Filled marks validation + final CO report generation

Overall complexity profile is dominated by Excel workbook generation/validation paths.

---

## 1. Architecture Coverage

### 1.1 Module Map

- Instructor UI: `modules/instructor_module.py`
  - Step logic: `modules/instructor/steps/*`
  - Validation: `modules/instructor/validators/*`
  - Services: `services/instructor_workflow_service.py`
  - Engines: `domain/instructor_template_engine.py`, `domain/instructor_report_engine.py`
- Coordinator UI: `modules/coordinator_module.py`
  - Step logic: `modules/coordinator/steps/*`
  - Runtime contracts: `modules/coordinator/contracts.py`
  - Workflow controller: `modules/coordinator/workflow_controller.py`
  - Processing: `domain/coordinator_engine.py`
  - Service: `services/coordinator_workflow_service.py`
- Help UI: `modules/help_module.py`
- About UI: `modules/about_module.py`
- PO Analysis UI: `modules/po_analysis_module.py` (placeholder plugin module)
- Plugin catalog: `modules/module_catalog.py`

### 1.2 Shared Infrastructure

- Text/i18n: `common/texts/*`
- UI logging + status payloads: `common/ui_logging.py`
- Toasts: `common/toast.py`
- Workbook signing/secret policy: `common/workbook_signing.py`, `common/workbook_secret.py`
- Runtime helpers/storage/logging: `common/utils.py`
- Module plugin contract: `common/module_plugins.py`

---

## 2. Instructor Module Complexity

### 2.1 Step 1: Course Details + Marks Template

Primary operations:
- Generate course details template
- Validate uploaded course-details workbook(s)
- Generate marks template(s)

Dominant complexity:
- Template generation: `O(R)` where `R` is row count written.
- Course-details validation: `O(R * C)` worst case (rows x effective columns).
- Marks-template generation: `O(Cm * S * Q)`
  - `Cm`: number of assessment components
  - `S`: students
  - `Q`: question/parameter columns per component

Space:
- openpyxl paths load workbook structures in memory: approx `O(R * C)`.
- xlsxwriter generation is write-optimized but still scales with emitted cells.

### 2.2 Step 2: Filled Marks + Final Report

Primary operations:
- Validate filled-marks workbook integrity + schema
- Generate final report(s)

Dominant complexity:
- Filled-marks validation: roughly `O(S * Q * Cm)` plus integrity checks.
- Final report generation/copying: `O(F)` where `F` is source/output file size.

Notes:
- Workflow service wrappers add timeout/cancellation controls; complexity overhead is constant relative to workbook processing.

---

## 3. Coordinator Module Complexity

Primary operations:
- Collect and validate uploaded final CO report files
- Parse and aggregate CO attainment data
- Generate consolidated coordinator output workbook

Dominant complexity:
- File collection/validation: `O(N)` for `N` input files.
- Workbook parsing + aggregation: approximately `O(N * S * O)`
  - `S`: students per file
  - `O`: outcomes/attainment columns processed
- Output workbook synthesis: proportional to emitted summary/detail rows.

Space:
- Parsing paths depend on openpyxl object graph + intermediate aggregation structures.
- Dedup/index structures introduce additional linear memory in unique student/outcome keys.

---

## 4. PO Analysis, Help, and About Module Complexity

### 4.1 Help Module

Primary operations:
- Load packaged PDF
- Save PDF to user-selected location
- Open PDF in default viewer

Complexity:
- Load/save/open operations are file-size bound: `O(F)`.
- UI operations are constant-time relative to workflow size.

### 4.2 About Module

Primary operations:
- Render static metadata/version content and assets.

Complexity:
- Time: `O(1)`
- Space: `O(1)` excluding static asset size already packaged.

### 4.3 PO Analysis Module (Placeholder)

Primary operations:
- Render placeholder metadata text.

Complexity:
- Time: `O(1)`
- Space: `O(1)`

---

## 5. Cross-Cutting Complexity and Safety

### 5.1 I18n and Logging

- Translation lookups are dictionary-based: `O(1)` average per key.
- Structured status/log payload parse/render is linear in message size.

### 5.2 Signature and Integrity Validation

- HMAC/signature checks are linear in payload length.
- Current version uses active versioned signature format only.

### 5.3 Cancellation/Timeout

- Cancellation checks are constant-time probes inserted through long-running loops.
- Service-level timeout wrappers do not change algorithmic order; they bound runtime.

---

## 6. Practical Scaling Guidance

- Instructor marks/report generation cost rises fastest with student count x component count x question count.
- Coordinator aggregation cost rises with number of uploaded files and per-file student/outcome volume.
- PO Analysis/Help/About do not materially affect compute budget.

Non-goals (intentional constraints):
- Do not cache Step 1 workbook metadata for Step 2 reuse across sessions. Users may run steps in different order, on different days, and for different courses on shared PCs.
- Do not reduce workbook formatting/protection/sheet footprint solely for speed. Current output structure is required by release workflow and integrity checks.

Operational guidance:
- Keep workbook templates structurally clean and bounded.
- Prefer batched processing for very large sections.
- Use perf soak + quality gates prior to release candidates.

---

## 7. Quality and Verification Status (Current)

The current baseline has passing checks for:
- `pytest -q`
- `pyright`
- `ruff`
- `isort`
- `pyflakes`
- `bandit`
- `coverage`
- `pip-audit`

This complexity analysis reflects the current two-step instructor workflow and active module architecture.
