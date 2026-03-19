# Business Logic Traceability Matrix

Date: 2026-03-19  
Repository: cotas  
Branch context: clementine

## Purpose

This document maps core business rules to:
1. Runtime enforcement points in the codebase.
2. Regression tests that verify the rule.
3. Noted gaps where behavior is currently permissive or fallback-driven.

## Legend

- Rule IDs:
  - INSTR-* for Instructor workflow rules
  - COORD-* for Coordinator workflow rules
  - POLICY-* for global policy constants and thresholds

## Traceability Matrix

| Rule ID | Business Rule | Primary Enforcement | Test Evidence |
|---|---|---|---|
| POLICY-01 | Active template id must be COURSE_SETUP_V1 | [common/constants.py](../common/constants.py#L40), [domain/instructor_template_engine.py](../domain/instructor_template_engine.py#L709) | [tests/test_course_details_template_generator_validation.py](../tests/test_course_details_template_generator_validation.py#L37) |
| POLICY-02 | Direct and indirect contribution ratios are fixed (80/20) | [common/constants.py](../common/constants.py#L70), [modules/coordinator_processing.py](../modules/coordinator_processing.py#L724) | [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L334) |
| POLICY-03 | Coordinator attainment levels use ordered thresholds | [modules/coordinator/validators/attainment_thresholds.py](../modules/coordinator/validators/attainment_thresholds.py#L7), [modules/coordinator_processing.py](../modules/coordinator_processing.py#L947) | [tests/test_coordinator_module_state.py](../tests/test_coordinator_module_state.py#L156), [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L454) |
| INSTR-01 | Course-details workbook must contain signed system hash and valid template signature | [domain/instructor_template_engine.py](../domain/instructor_template_engine.py#L709) | [tests/test_course_details_template_generator_validation.py](../tests/test_course_details_template_generator_validation.py#L40) |
| INSTR-02 | Course-details workbook schema must match blueprint sheet/header structure | [domain/instructor_template_engine.py](../domain/instructor_template_engine.py#L779) | [tests/test_course_details_template_generator_validation.py](../tests/test_course_details_template_generator_validation.py#L63) |
| INSTR-03 | Assessment config requires at least one direct and one indirect component, each totaling 100 | [domain/template_versions/course_setup_v1.py](../domain/template_versions/course_setup_v1.py#L337), [domain/template_versions/course_setup_v1.py](../domain/template_versions/course_setup_v1.py#L417), [domain/template_versions/course_setup_v1.py](../domain/template_versions/course_setup_v1.py#L421) | [tests/test_course_details_template_generator_validation.py](../tests/test_course_details_template_generator_validation.py#L50), [tests/test_course_setup_v1_helpers_extra.py](../tests/test_course_setup_v1_helpers_extra.py#L143) |
| INSTR-04 | Question map must reference valid components, valid CO indices, and single CO for CO-wise components | [domain/template_versions/course_setup_v1.py](../domain/template_versions/course_setup_v1.py#L455), [domain/template_versions/course_setup_v1.py](../domain/template_versions/course_setup_v1.py#L515) | [tests/test_course_details_template_generator_validation.py](../tests/test_course_details_template_generator_validation.py#L63), [tests/test_course_setup_v1_helpers_extra.py](../tests/test_course_setup_v1_helpers_extra.py#L157) |
| INSTR-05 | Student list must be non-empty and Reg_No values unique | [domain/template_versions/course_setup_v1.py](../domain/template_versions/course_setup_v1.py#L533) | [tests/test_course_details_template_generator_validation.py](../tests/test_course_details_template_generator_validation.py#L76) |
| INSTR-06 | Filled-marks upload must pass signed system manifest + layout hash checks before deep validation | [modules/instructor/validators/step2_filled_marks_validator.py](../modules/instructor/validators/step2_filled_marks_validator.py#L45), [modules/instructor/validators/step2_filled_marks_validator.py](../modules/instructor/validators/step2_filled_marks_validator.py#L96) | [tests/test_instructor_module_step2_validation.py](../tests/test_instructor_module_step2_validation.py#L90), [tests/test_instructor_module_step2_validation.py](../tests/test_instructor_module_step2_validation.py#L103) |
| INSTR-07 | Filled-marks rows enforce mark-entry rules (no empties, range checks, A handling, formula integrity, structure snapshot) | [domain/template_versions/course_setup_v1.py](../domain/template_versions/course_setup_v1.py#L668) | [tests/test_instructor_module_step2_validation.py](../tests/test_instructor_module_step2_validation.py#L129), [tests/test_instructor_module_step2_validation.py](../tests/test_instructor_module_step2_validation.py#L215), [tests/test_instructor_module_step2_validation.py](../tests/test_instructor_module_step2_validation.py#L229), [tests/test_instructor_module_step2_validation.py](../tests/test_instructor_module_step2_validation.py#L250) |
| INSTR-08 | Final CO report generation must reject tampered hash/layout and unsupported template ids | [domain/instructor_report_engine.py](../domain/instructor_report_engine.py#L269) | [tests/test_final_co_report_generator.py](../tests/test_final_co_report_generator.py#L213), [tests/test_final_co_report_generator.py](../tests/test_final_co_report_generator.py#L227), [tests/test_final_co_report_generator.py](../tests/test_final_co_report_generator.py#L241) |
| INSTR-09 | Final CO report output includes hidden system integrity sheets | [domain/instructor_report_engine.py](../domain/instructor_report_engine.py#L128) | [tests/test_final_co_report_generator.py](../tests/test_final_co_report_generator.py#L256) |
| COORD-01 | Uploaded coordinator files must be valid final reports with matching baseline signature dimensions | [modules/coordinator_processing.py](../modules/coordinator_processing.py#L400), [modules/coordinator_processing.py](../modules/coordinator_processing.py#L464) | [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L151), [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L131) |
| COORD-02 | Section uniqueness is enforced across accepted final reports | [modules/coordinator_processing.py](../modules/coordinator_processing.py#L471) | [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L170), [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L185) |
| COORD-03 | Registration deduplication is enforced during attainment aggregation and duplicates are counted | [modules/coordinator_processing.py](../modules/coordinator_processing.py#L143), [modules/coordinator_processing.py](../modules/coordinator_processing.py#L1064) | [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L308), [tests/test_coordinator_calculate_attainment_step.py](../tests/test_coordinator_calculate_attainment_step.py#L191) |
| COORD-04 | Score to attainment level mapping follows boundaries and returns NA outside valid numeric domain | [modules/coordinator_processing.py](../modules/coordinator_processing.py#L947) | [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L454) |
| COORD-05 | Coordinator output workbook includes summary/graph plus hidden integrity sheets | [modules/coordinator_processing.py](../modules/coordinator_processing.py#L807), [modules/coordinator_processing.py](../modules/coordinator_processing.py#L862), [modules/coordinator_processing.py](../modules/coordinator_processing.py#L906) | [tests/test_coordinator_module_validation.py](../tests/test_coordinator_module_validation.py#L301) |

## Business Logic Observations

1. Two-pass validation (fast structural/hash pass then deep template-rule pass) is consistently used for both course-details and filled-marks flows.
2. Template-version dispatch centralizes schema/rule variance in one versioned module, which improves upgrade safety.
3. Coordinator intake enforces cross-file compatibility before expensive aggregation, reducing downstream failure risk.

## Known Gaps and Follow-up Items

| Gap ID | Observation | Current Location | Suggested Action |
|---|---|---|---|
| GAP-01 | Step run gating in controller is permissive and returns true for all steps | [modules/instructor/workflow_controller.py](../modules/instructor/workflow_controller.py#L38) | Move effective precondition logic into controller and make it authoritative |
| GAP-02 | Batch generation loops count failures but swallow per-file exception detail | [modules/instructor/steps/step2_course_details_and_marks_template.py](../modules/instructor/steps/step2_course_details_and_marks_template.py#L528), [modules/instructor/steps/step2_filled_marks_and_final_report.py](../modules/instructor/steps/step2_filled_marks_and_final_report.py#L238) | Capture and emit per-file failure reason list in result payload |
| GAP-03 | Service-unavailable fallback can copy source workbook rather than execute report generation logic | [modules/instructor/steps/step2_filled_marks_and_final_report.py](../modules/instructor/steps/step2_filled_marks_and_final_report.py#L236), [modules/instructor/steps/step2_filled_marks_and_final_report.py](../modules/instructor/steps/step2_filled_marks_and_final_report.py#L488) | Restrict fallback to explicit test mode or fail fast in production paths |

## Suggested Maintenance Process

For any new business rule:
1. Add or update constant/policy in [common/constants.py](../common/constants.py).
2. Enforce rule in domain or module step layer.
3. Add a targeted failing test in [tests](../tests).
4. Add/update this matrix row in this file.
