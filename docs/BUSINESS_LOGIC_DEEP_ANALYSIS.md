# FOCUS — Deep Business Logic Analysis Report

**Generated:** 2026-03-20  
**Branch:** `clementine`  
**Test Status:** 484 passed | Quality gates green

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Architecture Overview](#2-architecture-overview)
3. [Domain Layer — Core Business Logic](#3-domain-layer--core-business-logic)
   - 3.1 [Instructor Template Engine](#31-instructor-template-engine)
   - 3.2 [Instructor Report Engine](#32-instructor-report-engine)
   - 3.3 [Template Version System](#33-template-version-system)
   - 3.4 [Workflow State Management](#34-workflow-state-management)
4. [Services Layer — Orchestration](#4-services-layer--orchestration)
5. [Modules Layer — UI & Coordination](#5-modules-layer--ui--coordination)
   - 5.1 [Instructor Module](#51-instructor-module)
   - 5.2 [Coordinator Module](#52-coordinator-module)
   - 5.3 [Supporting Modules](#53-supporting-modules)
6. [Common Layer — Shared Infrastructure](#6-common-layer--shared-infrastructure)
7. [Data Flow & Lifecycle](#7-data-flow--lifecycle)
8. [Business Rules Catalog](#8-business-rules-catalog)
9. [Attainment Computation Model](#9-attainment-computation-model)
10. [Integrity & Security Model](#10-integrity--security-model)
11. [Design Patterns & Architecture Decisions](#11-design-patterns--architecture-decisions)
12. [Test Coverage Analysis](#12-test-coverage-analysis)
13. [Known Gaps & Observations](#13-known-gaps--observations)
14. [Appendices](#14-appendices)
15. [Recommendations](#15-recommendations)

---

## 1. Executive Summary

FOCUS (Framework for Outcome Computation and Unification System) is a desktop OBE workflow tool for generating, validating, and processing Excel workbooks that track Course Outcome (CO) achievement. Its business logic spans three primary workflows:

| Workflow | Actors | Purpose |
|----------|--------|---------|
| **Instructor Step 1** | Faculty member | Generate course-details template → Validate filled details → Generate marks-entry template |
| **Instructor Step 2** | Faculty member | Validate filled marks → Generate final CO report (per-student direct/indirect achievement) |
| **Coordinator** | Program coordinator | Collect final CO reports from multiple instructors → Aggregate and compute attainment levels → Generate consolidated attainment workbook with charts |

The system enforces **workbook integrity** through HMAC-SHA256 signatures, **template versioning** through dispatch registries, and **business rule validation** through multi-pass schema + rule checking. All computations follow an **80/20 direct/indirect** contribution model with configurable attainment thresholds.

---

## 2. Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         FOCUS Application                                │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                          │
│  ┌──────────────────────────────────────────────────────────────────┐   │
│  │  MODULES LAYER (UI + Coordination)                                │   │
│  │  ┌──────────────┐ ┌──────────────┐ ┌──────┐ ┌────┐ ┌─────┐     │   │
│  │  │  Instructor   │ │ Coordinator  │ │  PO  │ │Help│ │About│     │   │
│  │  │   Module      │ │   Module     │ │ Anlys│ │    │ │     │     │   │
│  │  └──────┬───────┘ └──────┬───────┘ └──────┘ └────┘ └─────┘     │   │
│  │         │                │                                        │   │
│  │  ┌──────┴────────────────┴───────────────────────────────────┐   │   │
│  │  │  Sub-packages: steps/, validators/, output_links,          │   │   │
│  │  │  workflow_controller, messages, file_actions               │   │   │
│  │  └───────────────────────────────────────────────────────────┘   │   │
│  └──────────────────────────────┬───────────────────────────────────┘   │
│                                 │                                        │
│  ┌──────────────────────────────┴───────────────────────────────────┐   │
│  │  SERVICES LAYER (Orchestration + Telemetry)                       │   │
│  │  ┌─────────────────────────────────────────────────────────────┐    │   │
│  │  │ WorkflowServiceBase                                         │    │   │
│  │  │ - Timeout management                                        │    │   │
│  │  │ - Telemetry recording                                       │    │   │
│  │  │ - Error classification                                      │    │   │
│  │  └────────────┬──────────────────────────────┬─────────────────┘    │   │
│  │               │                              │                      │   │
│  │  ┌────────────▼────────────┐  ┌──────────────▼──────────────┐      │   │
│  │  │ InstructorWorkflowSvc   │  │ CoordinatorWorkflowSvc      │      │   │
│  │  │ - ValidationError hook  │  │ - Callback-based operations │      │   │
│  │  └─────────────────────────┘  └─────────────────────────────┘      │   │
│  └─────────────────────────────────────────────────────────────────────┘   │
│                                                                           │
│  ┌───────────────┴──────────────────────────────┴────────────────────┐  │
│  │  DOMAIN LAYER (Pure Business Logic)                                │  │
│  │  ┌──────────────────────┐  ┌────────────────────────────────┐     │  │
│  │  │ Template Engine       │  │ Report Engine                   │     │  │
│  │  │ - Template generation │  │ - Mark aggregation by CO        │     │  │
│  │  │ - Schema validation   │  │ - Weighted scoring              │     │  │
│  │  │ - Marks generation    │  │ - Direct/Indirect report sheets │     │  │
│  │  └──────────┬───────────┘  └──────────┬─────────────────────┘     │  │
│  │             │                          │                            │  │
│  │  ┌──────────┴──────────────────────────┴────────────────────┐     │  │
│  │  │ Template Versions (Strategy Pattern)                      │     │  │
│  │  │ - course_setup_v1: validators, extractors, rule checkers  │     │  │
│  │  └──────────────────────────────────────────────────────────┘     │  │
│  │  ┌──────────────────────────────┐                                  │  │
│  │  │ SheetOps (Low-level Excel)   │                                  │  │
│  │  │ - Formatting, Protection     │                                  │  │
│  │  │ - Sheet writers (3 types)    │                                  │  │
│  │  │ - Spec builders for manifest │                                  │  │
│  │  └──────────────────────────────┘                                  │  │
│  └────────────────────────────────────────────────────────────────────┘  │
│                                                                          │
│  ┌────────────────────────────────────────────────────────────────────┐  │
│  │  COMMON LAYER (Cross-Cutting Infrastructure)                        │  │
│  │  constants · contracts · exceptions · sheet_schema · registry       │  │
│  │  workbook_signing · workbook_secret · excel_sheet_layout            │  │
│  │  jobs · async_operation_runner · module_runtime · module_messages   │  │
│  │  module_plugins · error_catalog · output_markup · utils · texts     │  │
│  └────────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────┘
```

**Layer Rules (enforced by architecture tests):**

| Layer | May import from | Must not import from |
|-------|----------------|---------------------|
| Domain | Common | Services, Modules |
| Services | Domain, Common | Modules |
| Modules | Services, Domain, Common | — |
| Common | — | Domain, Services, Modules |

---

## 3. Domain Layer — Core Business Logic

### 3.1 Instructor Template Engine

**File:** `domain/instructor_template_engine.py` + `domain/instructor_template_engine_sheetops.py`

The template engine implements the **workbook generation and validation pipeline** for the instructor workflow. It has three primary use-case functions:

#### 3.1.1 Course Details Template Generation

**Function:** `generate_course_details_template(output_path, template_id, *, cancel_token) → Path`

**Algorithm:**
1. Validate `template_id` against the blueprint registry (only `COURSE_SETUP` currently supported)
2. Retrieve the `WorkbookBlueprint` from `BLUEPRINT_REGISTRY`
3. Create a temporary file with UUID suffix (atomic-write strategy)
4. For each sheet defined in the blueprint:
   - Create worksheet with validation rules (e.g., YES/NO dropdowns for Assessment Config)
   - Apply header formatting (bold, light-green background `#D9EAD3`, borders)
   - Apply body formatting (unlocked cells for data entry)
   - Protect sheet if specified in schema
5. Add hidden `__SYSTEM_HASH__` sheet with signed `template_id`
6. Atomically replace output file via `os.replace()`

**Blueprint Structure (COURSE_SETUP):**

| Sheet | Headers | Protection | Validation |
|-------|---------|------------|------------|
| Course_Metadata | Field, Value | No (user entry) | — |
| Assessment_Config | Component, Weight(%), CIA, CO_Wise_Marks_Breakup, Direct | No | YES/NO dropdown on "Direct" column |
| Question_Map | Component, Q_No/Rubric_Parameter, Max_Marks, CO | No | — |
| Students | Reg_No, Student_Name | No | — |

#### 3.1.2 Course Details Validation

**Function:** `validate_course_details_workbook(workbook_path) → template_id`

**Two-Pass Strategy:**
1. **Fast pass** (read-only load): Verify system hash sheet structure, extract and validate `template_id` signature
2. **Deep pass** (full load): Delegate to template-specific business rule validators

**Version-Specific Validation (course_setup_v1):**

| Validation | Rule | Error Behavior |
|------------|------|----------------|
| Course metadata fields | All required, non-empty, no duplicates, `total_outcomes > 0` (integer) | Fail-fast |
| Assessment config | ≥1 direct + ≥1 indirect component; direct weights sum to 100%; indirect weights sum to 100% | Fail-fast |
| Question map | Valid component references; `max_marks > 0`; valid CO references (1–total_outcomes); CO-wise breakup → exactly 1 CO per question | Fail-fast |
| Students | ≥1 student; both reg_no and name non-empty; no duplicate reg_no | Fail-fast |

#### 3.1.3 Marks Template Generation

**Function:** `generate_marks_template_from_course_details(course_details_path, output_path, *, cancel_token) → Path`

**Algorithm:**
1. Load and validate source course-details workbook (integrity check)
2. Extract template context: metadata, assessment config, questions, students
3. Pre-compute layout manifest (JSON metadata for round-trip validation)
4. Classify each component and generate the appropriate sheet type:

**Component Classification → Sheet Type:**

```
component.is_direct AND component.co_wise_breakup  → DIRECT_CO_WISE
component.is_direct AND NOT co_wise_breakup         → DIRECT_NON_CO_WISE
NOT component.is_direct                              → INDIRECT
```

**Sheet Type Details:**

| Type | Unlocked Columns | Formulas | Validation | Headers |
|------|-----------------|----------|------------|---------|
| **DIRECT_CO_WISE** | Per-question mark cells | `Total = SUM(Q1:Qn)` | `0 ≤ mark ≤ max_marks` OR `"A"` (absent) | Q1, Q2, ..., Qn, Total; CO row; Max-marks row |
| **DIRECT_NON_CO_WISE** | Total column only | CO splits: `ROUND(Total/n, 2)`, last CO = residual `Total - sum(others)` | `0 ≤ total ≤ max` OR `"A"` | Total, CO1_Marks, CO2_Marks, ...; CO row; Max row |
| **INDIRECT** | Per-CO Likert cells | None | `1 ≤ value ≤ 5` (Likert) OR `"A"` | CO1, CO2, ..., COn |

All sheets share common structure:
- Metadata rows (course info + component name) above header
- Frozen panes at header row + 1
- Sheet protection (locked except mark-entry cells)
- Student identity hash stored in manifest

**Layout Manifest:** A hidden `__SYSTEM_LAYOUT__` sheet stores a JSON object describing the complete workbook structure — sheet specs, anchors (cell references to expected values), formula expectations, student identity hash, and mark structure snapshots. This manifest enables comprehensive round-trip validation when the workbook returns filled.

### 3.2 Instructor Report Engine

**File:** `domain/instructor_report_engine.py`

The report engine converts filled marks workbooks into final CO achievement reports.

**Function:** `generate_final_co_report(filled_marks_path, output_path, *, cancel_token) → Path`

**Algorithm:**
1. Validate source workbook integrity (system hash + layout manifest signatures)
2. Read course metadata → extract `total_outcomes` count
3. Read component definitions (direct: name + weight; indirect: name + weight)
4. Parse layout manifest for sheet structure
5. **For each direct component:** Compute marks aggregated by CO
6. **For each indirect component:** Read Likert responses per CO
7. **For each CO (1 to total_outcomes):** Generate two report sheets:
   - **Direct sheet:** Raw marks → weighted marks → percentage → ratio
   - **Indirect sheet:** Likert scores → scaled scores → weighted → percentage → ratio
8. Add system integrity sheets (hash + per-sheet hashes)
9. Atomic file replacement

#### 3.2.1 Direct Report Computation

For each student and each CO:

```
Per-component weighted mark = (raw_mark × component.weight) / component.max_marks_for_co
Total weighted = Σ(weighted marks across all direct components for this CO)
Total out of 100 = (total_weighted / Σ(active component weights)) × 100
Ratio total = Total_100 × DIRECT_RATIO (0.8)
```

**Absent handling:** If any mark in a row is `"A"`, all computed values become `"N/A"`.

**Residual splitting** (non-CO-wise): Total is divided equally across covered COs, with rounding residual applied to the last CO via `split_equal_with_residual()`.

#### 3.2.2 Indirect Report Computation

For each student and each CO:

```
Scaled score = Likert_value − LIKERT_MIN (normalizes to 0–4 range)
If single component:
    Total_100 = (scaled / (LIKERT_MAX − LIKERT_MIN)) × 100
If multiple components:
    Weighted_scaled = (scaled × component.weight) / max_scaled
    Total_100 = (Σ(weighted_scaled) / Σ(weights)) × 100
Ratio total = Total_100 × INDIRECT_RATIO (0.2)
```

### 3.3 Template Version System

**File:** `domain/template_versions/course_setup_v1.py`

Implements the **Strategy pattern** — validators, extractors, and writers are registered in dispatch maps keyed by `template_id`:

```python
_template_rule_validators()      # template_id → course-details validator
_template_context_extractors()   # template_id → context extractor
_template_marks_writers()        # template_id → marks workbook writer
```

Currently only `COURSE_SETUP` (v1) is registered. This architecture supports adding future template versions (e.g., different institution formats) without modifying core engine code.

**Key v1 Validators:**

| Validator | Scope | Critical Rules |
|-----------|-------|----------------|
| `validate_course_details_rules()` | Course details workbook | All metadata present; weights sum to 100% per category; CO references valid |
| `validate_filled_marks_manifest_schema()` | Filled marks workbook | Sheet order matches manifest; anchors match expected values (ε=1e-9); formulas intact; student identity hash unchanged; marks non-empty and in-range; no mixed absent/numeric rows |

**Filled Marks Deep Validation includes:**
- Mark structure snapshot verification (max marks per column haven't changed)
- Absence policy: A row cannot mix `"A"` and numeric marks
- Anomaly warnings (non-blocking): High absence ratio (>90%), near-constant marks (>95% same value)
- Formula consistency: SUM formulas in CO-wise, CO-split formulas in non-CO-wise

### 3.4 Workflow State Management

**File:** `domain/workflow_state.py`

```python
@dataclass(slots=True)
class InstructorWorkflowState:
    current_step: int = 1
    busy: bool = False
    active_job_id: str | None = None

    def set_busy(value, *, job_id=None):
        # Atomic coupling of busy flag + job_id
```

Lightweight state container that gates UI actions during async operations.

---

## 4. Services Layer — Orchestration

**Files:** `services/instructor_workflow_service.py`, `services/coordinator_workflow_service.py`

Services wrap domain logic with cross-cutting concerns:

### 4.1 Common Service Patterns

Both services now inherit from `WorkflowServiceBase` (`services/workflow_service_base.py`) and share a single `_execute_with_telemetry()` implementation:

```
1. Resolve timeout (env var `FOCUS_WORKFLOW_STEP_TIMEOUT_SECONDS` or 120s default)
2. Log STEP_STARTED with operation name and timeout
3. Execute work via ThreadPoolExecutor with timeout
4. Record telemetry metric (operation, outcome, duration_ms)
5. Handle exceptions:
   - JobCancelledError → log cancellation, record metric
   - ValidationError → log warning with error code, re-raise
   - AppSystemError → log error, re-raise
   - Unknown → log exception, re-raise
```

**Timeout enforcement:** Uses `concurrent.futures.ThreadPoolExecutor(max_workers=1)` — if work exceeds timeout, the future is cancelled and the cancellation token is signaled.

**Cancellation token bridging (instructor only):** `_call_with_optional_cancel_token()` inspects function signatures at runtime via `inspect.signature()` to conditionally pass cancellation tokens — enabling gradual adoption of cancellation support in domain functions.

**Service-specific customization points:**
- `InstructorWorkflowService` overrides `_handle_domain_exception()` for `ValidationError` warning logs with stable error codes.
- `CoordinatorWorkflowService` reuses base error handling as-is and calls domain engine functions directly.

### 4.2 Instructor Workflow Service

| Operation | Domain Function | Returns |
|-----------|----------------|---------|
| `generate_course_details_template()` | `domain.generate_course_details_template()` | Output Path |
| `validate_course_details_workbook()` | `domain.validate_course_details_workbook()` | Template ID |
| `generate_marks_template()` | `domain.generate_marks_template_from_course_details()` | Output Path |
| `generate_final_report()` | `domain.generate_final_co_report()` | Output Path |

### 4.3 Coordinator Workflow Service

| Operation | Mechanism | Returns |
|-----------|-----------|---------|
| `collect_files()` | Calls `domain.coordinator_engine._analyze_dropped_files()` | `{added, duplicates, invalid, invalid_details, ignored}` |
| `calculate_attainment()` | Calls `domain.coordinator_engine._generate_co_attainment_workbook()` | Attainment result object |

The coordinator service now uses **direct domain imports**. Coordinator business logic resides in `domain/coordinator_engine.py`, while the module layer retains UI/thread orchestration.

---

## 5. Modules Layer — UI & Coordination

### 5.1 Instructor Module

**File:** `modules/instructor_module.py` + `modules/instructor/` package

**UI Layout:** Split-pane with fixed-width left navigation (step list) and flexible right content area (active card + drop zone + info tabs).

**State Machine:**

```
Step 1: Course Template Download
  └─ State: step1_path, step1_done

Step 2a: Course Details Upload → Marks Template Generation
  └─ State: step2_course_details_path, step2_upload_ready, marks_template_path, marks_template_done

Step 2b: Filled Marks Upload → Final Report Generation  
  └─ State: filled_marks_path, filled_marks_done, final_report_path, final_report_done, final_report_outdated
```

**Key Business Behaviors:**

1. **Step gating:** Step 2 requires `step2_upload_ready` (course details uploaded and validated)
2. **Invalidation cascade:** Changing course details invalidates marks template; changing filled marks sets `final_report_outdated`
3. **Multi-file support:** Upload dialogs accept multiple files; each validated independently
4. **Deduplication:** Files deduplicated by resolved path before validation
5. **Default naming:** Output filenames derived from workbook metadata: `CODE_SEM_SECTION_YEAR_<suffix>.xlsx`
6. **Service-backed report generation:** Final report generation requires workflow service/domain execution; unavailable-service copy fallback is disabled.

**Step Handlers (in `modules/instructor/steps/`):**

| Handler | Responsibility |
|---------|---------------|
| `step1_course_details_template.py` | Template download with save dialog |
| `step2_course_details_and_marks_template.py` | Multi-file upload, validation loop, marks generation |
| `step2_filled_marks_and_final_report.py` | Filled marks validation, iterative report generation with conflict resolution |
| `shared_workbook_ops.py` | Filename building, token sanitization, atomic copy |

### 5.2 Coordinator Module

**File:** `modules/coordinator_module.py` + `modules/coordinator/` package + `domain/coordinator_engine.py`

**UI Layout:** Left pane with threshold configuration spinboxes; right pane with drag-drop file list and calculate button.

**Attainment Threshold Configuration:**

```
Level 1 threshold: default 40.0%  (configurable via UI spinbox)
Level 2 threshold: default 60.0%
Level 3 threshold: default 75.0%
Constraint: 0 < L1 < L2 < L3 < 100
```

#### 5.2.1 File Collection Pipeline

```
User drops files
   │
   ├─ If busy → queue in _pending_drop_batches
   │
   └─ Process immediately:
       1. Filter for Excel extensions (.xlsx, .xlsm, .xls)
       2. Parallel signature validation (ThreadPoolExecutor, max 8 workers):
          - Extract template_id from SYSTEM_HASH_SHEET
          - Verify HMAC signature
          - Extract metadata: course code, total outcomes, section
       3. Establish baseline signature from first valid file
       4. Cross-file compatibility check:
          - Template ID must match
          - Course code must match
          - Total outcomes must match
          - Different sections allowed (multi-section aggregation)
       5. Deduplication against existing files
       6. Result: {added, duplicates, invalid_final_report, ignored}
       7. Show toast feedback per category
       8. Drain queued batches
```

#### 5.2.2 Attainment Calculation

**Function:** `_generate_co_attainment_workbook(source_paths, output_path, *, token, thresholds)`

**Algorithm (from `domain/coordinator_engine.py`):**

1. Extract baseline signature from first source file
2. Initialize registration deduplication store (in-memory for typical workloads; SQLite when estimated scale threshold is crossed)
3. For each CO (1 to total_outcomes):
   a. Create output sheet with metadata header
   b. For each source workbook:
       - Read direct CO sheet → extract ratio total (Total × DIRECT_RATIO)
       - Read indirect CO sheet → extract ratio total (Total × INDIRECT_RATIO)
       - Match students by registration hash (48-bit BLAKE2b)
       - Log and report dropped non-matching rows (direct-only or indirect-only)
       - For each student: `combined_score = direct_ratio_total + indirect_ratio_total`
      - Classify: `level = _score_to_attainment_level(combined_score, thresholds)`
      - Deduplicate: skip if reg_hash already seen for this CO
      - Write row: Serial, Reg No, Name, Direct%, Indirect%, Total%, Level
   c. Write summary rows: On-roll count, attended count, per-level counts
4. Create Summary sheet: CO-level aggregation with `CO% = (L2 + L3) / Attended × 100`
5. Create Graph sheet: Column chart of CO% per outcome
6. Write system integrity sheets

**Deduplication Storage (for scale):**
- Uses SQLite with aggressive pragmas (`journal_mode=OFF`, `synchronous=OFF`, `temp_store=MEMORY`)
- Secure deletion (`secure_delete=ON`) before file cleanup
- Threshold: switches from in-memory to SQLite when estimated entries exceed `_DEDUP_SQLITE_THRESHOLD_ENTRIES` (currently 10,000)

**Score Classification:**

| Score Range | Level |
|------------|-------|
| `0.0 ≤ score < L1 (40)` | Level 0 |
| `L1 ≤ score < L2 (60)` | Level 1 |
| `L2 ≤ score < L3 (75)` | Level 2 |
| `L3 ≤ score ≤ 100.0` | Level 3 |
| Outside range / absent | N/A |

**CO% Formula:** `CO% = (count_Level2 + count_Level3) / count_Attended × 100`

**Output Workbook Structure:**
- Per-CO detail sheets (paginated at 150 students per sheet)
- Summary sheet with all COs
- Graph sheet (column chart)
- Hidden system sheets (hash + integrity)

### 5.3 Supporting Modules

| Module | Description |
|--------|-------------|
| **PO Analysis** | Placeholder with centered label text; wired through plugin catalog for future implementation |
| **Help** | Embedded PDF viewer (`QPdfDocument`) with context menu: Save-as and Open-external actions |
| **About** | Static metadata: version, copyright (dynamic year), institution, description |

**Plugin Catalog (`modules/module_catalog.py`):**
All five modules registered via `build_module_catalog()` returning `tuple[ModulePluginSpec, ...]`. Uses `lazy_module_class()` for deferred imports — reduces startup time.

---

## 6. Common Layer — Shared Infrastructure

### 6.1 Policy Constants (`common/constants.py`)

| Category | Constant | Value | Purpose |
|----------|----------|-------|---------|
| Attainment | `DIRECT_RATIO` | 0.8 | Direct assessment contribution weight |
| Attainment | `INDIRECT_RATIO` | 0.2 | Indirect assessment contribution weight |
| Attainment | `LEVEL_1_THRESHOLD` | 40.0 | Minimum for Level 1 attainment |
| Attainment | `LEVEL_2_THRESHOLD` | 60.0 | Minimum for Level 2 attainment |
| Attainment | `LEVEL_3_THRESHOLD` | 75.0 | Minimum for Level 3 attainment |
| Marks | `LIKERT_MIN` / `LIKERT_MAX` | 1 / 5 | Indirect assessment scale |
| Marks | `MIN_MARK_VALUE` | 0.0 | Minimum mark in direct assessment |
| Weights | `WEIGHT_TOTAL_EXPECTED` | 100.0 | Component weights must sum to this |
| Formatting | `CO_REPORT_MAX_DECIMAL_PLACES` | 2 | Rounding precision |
| System | `SYSTEM_HASH_SHEET` | `__SYSTEM_HASH__` | Hidden integrity sheet name |
| System | `SYSTEM_LAYOUT_SHEET` | `__SYSTEM_LAYOUT__` | Hidden manifest sheet name |
| System | `SYSTEM_REPORT_INTEGRITY_SHEET` | `__REPORT_INTEGRITY__` | Hidden report integrity sheet name |

### 6.2 Contract Validation (`common/contracts.py`)

**Bootstrap invariant checks** (run at import time or startup):
- `DIRECT_RATIO + INDIRECT_RATIO ≈ 1.0` (5-decimal precision)
- `0 ≤ LEVEL_1 ≤ LEVEL_2 ≤ LEVEL_3 ≤ 100` (monotonic ordering)
- `LIKERT_MIN < LIKERT_MAX`
- Blueprint registry: non-empty, type_id matches, no duplicate sheet names, all headers non-empty

### 6.3 Exception Hierarchy

```
AppError (base)
├── ValidationError(message, code, context)  → User/data errors, recoverable
├── ConfigurationError(message)              → Policy/config errors, fail-fast
├── AppSystemError(message)                  → System/resource errors
└── JobCancelledError(message)               → User/system cancellation
```

### 6.4 Workbook Schema System

`WorkbookBlueprint → SheetSchema[] → ValidationRule[]`

Declarative schema-as-code: blueprints define sheet structure, header layout, validation rules, and protection flags. Consumed by template generators and validated by contract checks.

### 6.5 Workbook Signing (`common/workbook_signing.py`)

- **Algorithm:** HMAC-SHA256
- **Format:** `"v1:<hex_digest>"`
- **Key source:** Application-managed secret from `workbook_secret.py`
- **Verification:** Timing-safe comparison via `hmac.compare_digest()`
- **Key rotation:** Architecture supports multiple accepted secrets via `_accepted_secrets()` tuple
- **What's signed:** Template IDs, layout manifests (JSON), report integrity manifests

### 6.6 Secret Management (`common/workbook_secret.py`)

- **Windows:** DPAPI via `CryptProtectData` (machine-scoped)
- **POSIX:** Base64 obfuscation (not cryptographic)
- **Bootstrap:** XOR-obfuscated default → load from encrypted store → sanitize → cache
- **Storage:** Platform-specific paths (ProgramData, /Users/Shared, ~/.local/share)
- **Policy enforcement:** `ensure_workbook_secret_policy()` raises `ConfigurationError` if secret unavailable

### 6.7 Async Infrastructure

**`AsyncOperationRunner`:** State machine for UI async operations:
1. Set busy flag + cancellation token on target module
2. Execute work via `run_in_background` (Qt signal handler)
3. Track job in active list
4. On completion: call success/failure callback → clear busy → refresh UI

**`CancellationToken`:** Thread-safe via `threading.Event`. Checked at key domain-logic checkpoints.

**`JobContext`:** Frozen dataclass with 12-char UUID hex, step ID, language, timestamp, and optional payload. Propagated to all log records for traceability.

### 6.8 Plugin System

`ModulePluginSpec(key, title_key, icon_path, class_loader)` → `lazy_module_class()` defers import until first use. `build_module_catalog()` returns canonical tuple of all modules.

---

## 7. Data Flow & Lifecycle

### 7.1 Complete Instructor Workflow

```
┌──────────────────────────────────────────────────────────────────────┐
│ STEP 1: Course Setup                                                  │
│                                                                        │
│  generate_course_details_template()                                   │
│  ┌─────────────┐                                                      │
│  │ BLUEPRINT    │──→ xlsxwriter ──→ [Course Details Template.xlsx]    │
│  │ REGISTRY     │                   │                                  │
│  └─────────────┘                   │ Instructor fills in:             │
│                                     │  - Course metadata               │
│                                     │  - Assessment config             │
│                                     │  - Question map                  │
│                                     │  - Student roster                │
│                                     ▼                                  │
│  validate_course_details_workbook()                                   │
│  [Filled Course Details.xlsx] ──→ Fast pass (schema)                  │
│                                   ──→ Deep pass (business rules)      │
│                                   ──→ Returns template_id             │
│                                     │                                  │
│  generate_marks_template_from_course_details()                        │
│  [Filled Course Details.xlsx] ──→ Extract context                     │
│                                ──→ Classify components                 │
│                                ──→ Pre-compute manifest               │
│                                ──→ Write sheets by type               │
│                                ──→ [Marks Template.xlsx]              │
│                                     │                                  │
│                                     │ Instructor enters marks:        │
│                                     │  - Direct: numeric (0-max, A)   │
│                                     │  - Indirect: Likert (1-5, A)    │
│                                     ▼                                  │
├──────────────────────────────────────────────────────────────────────┤
│ STEP 2: Report Generation                                             │
│                                                                        │
│  validate_filled_marks_workbook()                                     │
│  [Filled Marks.xlsx] ──→ Signature check                              │
│                       ──→ Manifest schema validation                   │
│                       ──→ Mark range/type validation                   │
│                       ──→ Formula consistency check                    │
│                       ──→ Student identity hash check                  │
│                                     │                                  │
│  generate_final_co_report()                                           │
│  [Filled Marks.xlsx] ──→ Compute marks by CO                          │
│                       ──→ Weight and aggregate                         │
│                       ──→ Generate Direct report sheets                │
│                       ──→ Generate Indirect report sheets              │
│                       ──→ Sign integrity                               │
│                       ──→ [Final CO Report.xlsx]                       │
└──────────────────────────────────────────────────────────────────────┘
```

### 7.2 Complete Coordinator Workflow

```
┌──────────────────────────────────────────────────────────────────────┐
│ COORDINATOR: Attainment Consolidation                                 │
│                                                                        │
│  File Intake                                                          │
│  [Report 1.xlsx] ──┐                                                  │
│  [Report 2.xlsx] ──┤  Parallel signature validation                   │
│  [Report N.xlsx] ──┘  Baseline compatibility check                    │
│                       Section-aware grouping                           │
│                       Deduplication                                    │
│                           │                                            │
│  Attainment Calculation   │                                            │
│                           ▼                                            │
│  For each CO (1 → total_outcomes):                                    │
│    For each source workbook:                                          │
│      Read direct ratio total (student × CO)                           │
│      Read indirect ratio total (student × CO)                         │
│      Combined = direct + indirect                                     │
│      Level = classify(combined, thresholds)                           │
│      Deduplicate by registration hash                                 │
│                                                                        │
│  Aggregate:                                                           │
│    Per-CO: {On-roll, Attended, Level 0-3 counts}                     │
│    CO% = (Level2 + Level3) / Attended × 100                          │
│                                                                        │
│  Output:                                                              │
│    Per-CO sheets (detail)                                             │
│    Summary sheet (all COs)                 ──→ [CO Attainment.xlsx]   │
│    Graph sheet (column chart)                                         │
│    System integrity sheets                                            │
└──────────────────────────────────────────────────────────────────────┘
```

---

## 8. Business Rules Catalog

### 8.1 Policy Rules (System-Wide)

| ID | Rule | Enforcement | Constant |
|----|------|-------------|----------|
| POL-01 | Direct/Indirect contribution must sum to 1.0 | `contracts.py` startup check | `DIRECT_RATIO=0.8`, `INDIRECT_RATIO=0.2` |
| POL-02 | Attainment thresholds must be monotonically increasing within (0, 100) | `contracts.py` + coordinator UI validation | `L1=40.0`, `L2=60.0`, `L3=75.0` |
| POL-03 | Likert scale: min < max | `contracts.py` | `LIKERT_MIN=1`, `LIKERT_MAX=5` |
| POL-04 | Component weights (direct) must sum to 100.0% | `course_setup_v1.validate_assessment_config()` | `WEIGHT_TOTAL_EXPECTED=100.0` |
| POL-05 | Component weights (indirect) must sum to 100.0% | `course_setup_v1.validate_assessment_config()` | `WEIGHT_TOTAL_EXPECTED=100.0` |

### 8.2 Instructor Rules

| ID | Rule | Enforcement Location |
|----|------|---------------------|
| INS-01 | Template ID must be signed and verifiable | `_extract_and_validate_template_id()` via HMAC verification |
| INS-02 | At least 1 direct + 1 indirect assessment component required | `course_setup_v1._validate_assessment_config()` |
| INS-03 | CO-wise breakup questions must map to exactly 1 CO each | `course_setup_v1._validate_question_map()` |
| INS-04 | CO references must be in range [1, total_outcomes] | `course_setup_v1._validate_question_map()` |
| INS-05 | At least 1 student required | `course_setup_v1._validate_students()` |
| INS-06 | No duplicate registration numbers | `course_setup_v1._validate_students()` |
| INS-07 | Filled marks: student identity hash must match original roster | `validate_filled_marks_manifest_schema()` |
| INS-08 | Filled marks: no mixed absent/numeric in same row | `_validate_non_empty_marks_entries()` |
| INS-09 | Indirect marks must be integers in Likert range | `_validate_non_empty_marks_entries()` |
| INS-10 | All numeric marks must have ≤ 2 decimal places | `_has_allowed_decimal_precision()` |
| INS-11 | Formulas must be intact (SUM for CO-wise, split for non-CO-wise) | `_validate_row_total_consistency()` |
| INS-12 | Report output files must be written atomically | `atomic_copy_file()` via `os.replace()` |

### 8.3 Coordinator Rules

| ID | Rule | Enforcement Location |
|----|------|---------------------|
| CRD-01 | All source files must have valid HMAC signatures | `_extract_final_report_signature()` in parallel validation |
| CRD-02 | All files must share same template_id, course_code, total_outcomes | Baseline compatibility check in `_analyze_dropped_files()` |
| CRD-03 | Duplicate registrations deduplicated per CO (first occurrence wins) | `_RegisterDedupStore.add_if_absent()` |
| CRD-04 | Attainment thresholds must satisfy `0 < L1 < L2 < L3 < 100` | `has_valid_attainment_thresholds()` |
| CRD-05 | CO% = (Level2 + Level3) / Attended × 100 | Summary sheet generation |
| CRD-06 | Output workbook must include system integrity sheets | `_write_system_integrity_sheets()` |

---

## 9. Attainment Computation Model

### 9.1 Per-Student Score (Instructor Report)

```
For each Course Outcome (CO_i):

  DIRECT ASSESSMENT:
    For each direct component C_j:
      If CO-wise breakup:
        raw_mark = Σ(question marks mapped to CO_i)
        max_mark = Σ(max marks for questions mapped to CO_i)
      If non-CO-wise:
        raw_mark = total_mark / num_covered_COs  (last CO gets residual)
        max_mark = total_max / num_covered_COs    (last CO gets residual)
      
      weighted = (raw_mark × C_j.weight) / max_mark
    
    total_weighted_direct = Σ(weighted across all direct components)
    total_direct_100 = (total_weighted_direct / Σ(active weights)) × 100
    direct_ratio_total = total_direct_100 × DIRECT_RATIO  (× 0.8)

  INDIRECT ASSESSMENT:
    For each indirect component C_k:
      scaled = Likert_value − LIKERT_MIN  (0 to 4 range)
      If single component:
        total_indirect_100 = (scaled / (LIKERT_MAX − LIKERT_MIN)) × 100
      If multiple components:
        weighted_scaled = (scaled × C_k.weight) / max_scaled_value
        total_indirect_100 = (Σ(weighted_scaled) / Σ(weights)) × 100
    
    indirect_ratio_total = total_indirect_100 × INDIRECT_RATIO  (× 0.2)
```

### 9.2 Attainment Level Classification (Coordinator)

```
combined_score = direct_ratio_total + indirect_ratio_total

Level 0:  0.0 ≤ combined < 40.0   (Below threshold)
Level 1: 40.0 ≤ combined < 60.0   (Basic attainment)
Level 2: 60.0 ≤ combined < 75.0   (Moderate attainment)
Level 3: 75.0 ≤ combined ≤ 100.0  (High attainment — target)
N/A:     Student absent or score out of range

CO Attainment % = (count_Level2 + count_Level3) / count_Attended × 100
```

### 9.3 Numerical Precision

- All displayed marks rounded to 2 decimal places (`_round2()`)
- Non-CO-wise split uses residual rounding: `[3.33, 3.33, 3.34]` for 10/3
- Anchor/formula validation uses ε = 1e-9 for floating-point comparison
- Weight total comparison uses configurable rounding digits

---

## 10. Integrity & Security Model

### 10.1 Workbook Signing Chain

```
Template Generation:
  template_id ──→ HMAC-SHA256 ──→ stored in __SYSTEM_HASH__ sheet

Marks Template Generation:
  template_id hash (copied from source)
  layout_manifest (JSON) ──→ HMAC-SHA256 ──→ stored in __SYSTEM_LAYOUT__ sheet

Final Report Generation:
  template_id hash (preserved)
  report_manifest (JSON with per-sheet hashes) ──→ HMAC-SHA256 ──→ stored in __REPORT_INTEGRITY__

Coordinator Collection:
  All source files verified against signing chain before acceptance
```

### 10.2 Tamper Detection Points

| Check | Detects |
|-------|---------|
| Template ID signature | Modified template type or version |
| Layout manifest signature | Altered sheet structure, reordering, or formula changes |
| Anchor value matching | Cell content modifications in metadata rows |
| Student identity hash | Student additions, removals, or reordering |
| Formula normalization check | Tampered calculation formulas |
| Mark structure snapshot | Changed max marks or Likert ranges |

### 10.3 Sheet Protection

- All generated sheets are protected with application-managed password
- Mark-entry cells explicitly unlocked for user input
- Metadata, formulas, and header rows remain locked
- Protection settings: sort allowed, filter allowed, locked cell selection prevented

---

## 11. Design Patterns & Architecture Decisions

### 11.1 Patterns Used

| Pattern | Application | Benefit |
|---------|-------------|---------|
| **Strategy** | Template version dispatch registries | Extensible template formats without engine changes |
| **Two-Pass Validation** | Fast schema check → deep rule check | Fail fast on structural issues; defer expensive checks |
| **Atomic File Operations** | Write temp + `os.replace()` | Crash-safe workbook generation |
| **Layout Manifest** | Pre-compute then write | Round-trip validation; reproducible structure |
| **Protocol-Based DI** | Step handlers accept typed `Mapping` namespaces | Testable modules; decoupled step logic |
| **Facade** | `ModuleRuntime` wraps async + logging + status | Simplified module lifecycle |
| **Plugin Catalog** | Lazy-loaded `ModulePluginSpec` tuple | No hardcoded imports; fast startup |
| **Layered Orchestration** | Coordinator service invokes domain engine directly; module layer owns UI/threading concerns | Clear layer boundaries and better testability |
| **Contract Programming** | `contracts.py` bootstrap invariants | Fail-fast on configuration errors |
| **Hash-Based Integrity** | HMAC-SHA256 signing chain | Tamper detection without encryption |

### 11.2 Key Architecture Decisions

1. **Steps 1 and 2 are intentionally independent** — no caching of parsed workbook metadata across steps. Different courses, users, and sessions may invoke them independently.

2. **Workbook formatting/protection is part of required output** — not optional. Current formatting, hidden sheets, and protection behavior are mandatory for output compatibility.

3. **Module loading is plugin-catalog-driven** — no hardcoded module imports in `MainWindow`.

4. **Coordinator step orchestration uses explicit namespace contracts** — no `globals()` coupling.

5. **SQLite dedup for coordinator** — chosen for scale when processing many students × COs. Aggressive pragmas for performance, secure deletion for cleanup.

---

## 12. Test Coverage Analysis

### 12.1 Summary

- **474 tests passing** across 60+ test files
- **0 pyright errors**, **0 ruff/isort/pyflakes warnings**, **0 bandit findings**
- **Architecture boundary tests** enforce import rules and module size budgets

### 12.2 Coverage by Domain Area

| Area | Approx. Tests | Key Scenarios |
|------|--------------|---------------|
| Architecture boundaries | 8 | Import rules, module size limits |
| Common utilities | 50+ | Normalization, paths, settings, logging |
| Contracts/config | 10 | Ratio sums, threshold ordering, Likert, registry |
| Coordinator pipeline | 60+ | File drops, parallel validation, deduplication, attainment calculation, UI events |
| Instructor pipeline | 80+ | Template generation, validation, marks generation, report generation, cancellation |
| Template generation | 30+ | Schema stability, freeze panes, print titles, tampering detection |
| Workflow services | 15+ | Timeout enforcement, cancellation propagation, telemetry |
| Integration/E2E | 5+ | Full Step 1→2 flow, full coordinator flow, module smoke tests |
| Security/signing | 5+ | HMAC verification, version mismatch, secret rotation |

### 12.3 Test Patterns

- **Monkeypatch** for OS paths, env vars, imports
- **tmp_path** for isolated workbook creation
- **pytest.importorskip** for optional dependency gating (openpyxl, PySide6)
- **Real xlsx generation** in integration tests via openpyxl + xlsxwriter
- **Qt test doubles** with `qapp` fixture for UI module testing
- **CancellationToken** for testing early-exit paths

---

## 13. Known Gaps & Observations

### 13.1 Current Gap Status

| ID | Previous Description | Current Status |
|----|----------------------|----------------|
| GAP-01 | Step run gating was permissive | Closed — workflow controller now enforces step gating and blocks unknown steps. |
| GAP-02 | Batch report generation hid per-file details | Closed — per-file failure details are fully surfaced in logs/status output. |
| GAP-03 | Service-unavailable fallback could copy instead of generate | Closed — fallback copy path removed; service-unavailable now raises explicit system error. |

### 13.2 Analysis Observations

| Observation | Details |
|-------------|---------|
| **Single template version** | Only `COURSE_SETUP` v1 exists. The strategy dispatch infrastructure is ready but untested with multiple versions. |
| **PO Analysis placeholder** | Module is wired through the plugin catalog but contains no business logic yet. |
| **Coordinator domain engine in place** | Coordinator processing lives in `domain/coordinator_engine.py`; module steps handle UI/thread orchestration only. |
| **Anomaly warnings are non-blocking** | High absence rates and near-constant marks are logged as warnings but don't fail validation — appropriate for real-world scenarios where legitimate edge cases exist. |
| **Coordinator threshold mutability** | Thresholds are configurable via UI spinboxes at calculation time. No persistence of custom thresholds between sessions. |
| **Inner-join diagnostics are explicit** | Direct/Indirect row mismatches are logged and surfaced as dropped-row warnings during coordinator attainment generation. |

---

## 14. Appendices

### Appendix A: Sheet Name Reference

| System Sheet | Name | Visibility | Purpose |
|-------------|------|------------|---------|
| `SYSTEM_HASH_SHEET` | `__SYSTEM_HASH__` | Hidden | Template ID + HMAC signature |
| `SYSTEM_LAYOUT_SHEET` | `__SYSTEM_LAYOUT__` | Hidden | JSON layout manifest + signature |
| `SYSTEM_REPORT_INTEGRITY_SHEET` | `__REPORT_INTEGRITY__` | Hidden | Per-sheet hashes + manifest |
| `COURSE_METADATA_SHEET` | `Course_Metadata` | Visible | Course code, name, instructor, etc. |
| `ASSESSMENT_CONFIG_SHEET` | `Assessment_Config` | Visible | Component definitions + weights |
| `QUESTION_MAP_SHEET` | `Question_Map` | Visible | Question → CO mappings |
| `STUDENTS_SHEET` | `Students` | Visible | Registration roster |

### Appendix B: File → Function Quick Reference

| File | Key Functions |
|------|--------------|
| `domain/instructor_template_engine.py` | `generate_course_details_template()`, `validate_course_details_workbook()`, `generate_marks_template_from_course_details()` |
| `domain/instructor_report_engine.py` | `generate_final_co_report()` |
| `domain/template_versions/course_setup_v1.py` | `validate_course_details_rules()`, `validate_filled_marks_manifest_schema()` |
| `domain/workflow_state.py` | `InstructorWorkflowState` |
| `services/workflow_service_base.py` | `WorkflowServiceBase`, `WorkflowTelemetryConfig`, `WorkflowMetrics` |
| `services/instructor_workflow_service.py` | `InstructorWorkflowService` (4 operations + telemetry) |
| `services/coordinator_workflow_service.py` | `CoordinatorWorkflowService` (2 operations + telemetry) |
| `modules/instructor_module.py` | `InstructorModule` (Qt widget, step management) |
| `modules/coordinator_module.py` | `CoordinatorModule` (Qt widget, file collection + calculation) |
| `domain/coordinator_engine.py` | `_generate_co_attainment_workbook()`, `_analyze_dropped_files()`, `_iter_co_rows_from_workbook()` |

### Appendix C: Key Constants Quick Reference

| Constant | Value | Used In |
|----------|-------|---------|
| `DIRECT_RATIO` | 0.8 | Report engine, coordinator engine |
| `INDIRECT_RATIO` | 0.2 | Report engine, coordinator engine |
| `LEVEL_1_THRESHOLD` | 40.0 | Coordinator attainment classification |
| `LEVEL_2_THRESHOLD` | 60.0 | Coordinator attainment classification |
| `LEVEL_3_THRESHOLD` | 75.0 | Coordinator attainment classification |
| `LIKERT_MIN` / `LIKERT_MAX` | 1 / 5 | Indirect assessment validation |
| `WEIGHT_TOTAL_EXPECTED` | 100.0 | Assessment config weight validation |
| `CO_REPORT_MAX_DECIMAL_PLACES` | 2 | All numeric display rounding |
| `MAX_EXCEL_SHEETNAME_LENGTH` | 31 | Sheet name truncation |

---

## 15. Recommendations

### 15.1 High Priority

| # | Recommendation | Status | Affected Files |
|---|---------------|--------|----------------|
| 1 | Close GAP-02: surface per-file batch report errors. | **Completed** | `modules/instructor/steps/step2_filled_marks_and_final_report.py` |
| 2 | Move coordinator processing logic to domain layer. | **Completed** | `domain/coordinator_engine.py`, `modules/coordinator/steps/*` |

### 15.2 Medium Priority

| # | Recommendation | Status | Affected Files |
|---|---------------|--------|----------------|
| 3 | Strengthen POSIX secret management with keyring-first policy and documented fallback. | **Completed** | `common/workbook_secret.py`, `SECURITY.md` |

### 15.3 Low Priority

| # | Recommendation | Status | Affected Files |
|---|---------------|--------|----------------|
| 4 | Add threshold configuration metadata to coordinator output sheets. | **Completed** | `domain/coordinator_engine.py` |
| 5 | Surface anomaly warnings in instructor UI during Step 2 validation/report generation. | **Completed** | `domain/template_versions/course_setup_v1.py`, `modules/instructor/steps/step2_filled_marks_and_final_report.py`, `modules/instructor_module.py` |
| 6 | Evaluate and tune SQLite dedup threshold usage. | **Completed** | `domain/coordinator_engine.py` |

### 15.4 Current Follow-Ups

| # | Recommendation | Priority | Status |
|---|---------------|----------|--------|
| 7 | Maintain explicit diagnostics for dropped direct/indirect join rows and invalid file rejection reasons. | Medium | Completed in current branch; keep covered by regression tests. |
| 8 | Keep service-unavailable behavior strict (no generation bypass path) and monitor support incidents. | High | Completed and enforced. |

---

*End of Business Logic Deep Analysis Report*
