"""English text catalog for user-facing strings."""

TEXTS = {
    # Startup / main
    "app.already_running": "Application is already running.",
    "app.unexpected_error": "An unexpected error occurred. Please check the log file.",
    "app.startup.workbook_secret_missing_frozen": (
        "Workbook secret is unavailable. Please reinstall the software or contact support."
    ),
    "app.startup.workbook_secret_missing_dev": (
        "Workbook secret is unavailable. Delete the local workbook secret store and relaunch."
    ),
    "app.main_window_title": "FOCUS",
    "splash.starting": "Starting application...",
    "splash.loading_main_window": "Loading main window...",
    # Main window
    "toolbar.navigation": "Navigation",
    "status.ready": "Ready",
    "module.placeholder": "{title} module is not yet implemented.",
    "module.instructor": "Instructor",
    "module.coordinator_short": "Coordinator",
    "module.po_analysis": "PO Analysis",
    "module.co_analysis": "CO Analysis",
    "module.load_failed_title": "Module Load Error",
    "module.load_failed_body": "Could not load module '{module}'.\n\nDetails: {error}",
    "module.load_failed_status": "Failed to load module: {module}.",
    "nav.help": "Help",
    "nav.about": "About",
    # Settings module
    "language.switcher.button": "Language: {language}",
    "language.switcher.applied_status": "Language set to {language}.",
    # About
    "about.version": "Version {version}",
    "about.subtitle": "Framework for Outcome Computation and Unification System",
    "about.description": (
        "{app_name} is a software tool designed for computing Course Outcome (CO) "
        "attainment and performing structured outcome analysis based on direct "
        "and indirect assessments."
    ),
    "about.contributors.none": "Contributors information is not available.",
    "about.repository.link_label": "GitHub Repository",
    "about.meta_line_html": '{copyright} | <a href="{url}">{link_label}</a>',
    "about.institution": "Developed at Kalasalingam Academy of Research and Education (KARE) by the following contributors:",
    "about.copyright": "(c) {year} KARE. All rights reserved.",
    # Help module
    "help.doc_missing_title": "Help Document Missing",
    "help.doc_missing_body": "Could not find help PDF:\n{path}",
    "help.doc_error_title": "Help Document Error",
    "help.doc_error_body": "Failed to load help PDF.",
    "help.download_pdf": "Download PDF",
    "help.open_default_viewer": "Open in Default Viewer",
    "help.missing_file_title": "Missing File",
    "help.save_title": "Save CO Attainment Document",
    "help.save_default_name": "CO_Calculation_Document.pdf",
    "help.save_filter_pdf": "PDF Files (*.pdf)",
    "help.save_failed_title": "Save Failed",
    "help.save_success_title": "Saved",
    "help.open_failed_title": "Open Failed",
    "help.open_failed_body": "Could not open help PDF in the default viewer.",
    "help.open_success_title": "Opened",
    "help.open_success_body": "Help PDF opened in the default viewer.",
    "help.status.doc_loaded": "Help document loaded.",
    "help.status.doc_missing": "Help document is missing.",
    "help.status.doc_error": "Failed to load help document.",
    "help.status.file_missing": "Help PDF is not available.",
    "help.status.save_success": "Help PDF saved successfully.",
    "help.status.save_failed": "Error happened while saving help PDF.",
    "help.status.open_success": "Help PDF opened in default viewer.",
    "help.status.open_failed": "Error happened while opening help PDF in default viewer.",
    "outputs.none_generated": "No outputs generated.",
    # Instructor module
    "instructor.workflow_title": "CO Workflow",
    "instructor.step1.title": "Generate Marks Template",
    "instructor.step2.title": "Generate CO Marks Report",
    "instructor.step1.desc": "Upload the validated course-details file, then prepare and save the marks template.",
    "instructor.links.title": "Generated Outputs",
    "instructor.links.course_details_generated": "Course details template generated",
    "instructor.links.marks_template_generated": "Marks template generated",
    "instructor.links.final_co_report_generated": "Final CO report generated",
    "instructor.links.open_file": "Open file",
    "instructor.links.open_folder": "Open folder",
    "instructor.links.open_failed": "Unable to open the selected path.",
    "instructor.note.default": "You can return and replace any completed step.",
    "instructor.note.outdated_current": "This step is outdated. Run this step again.",
    "instructor.note.outdated_downstream": "An upstream change was made. Downstream steps may need regeneration.",
    "instructor.action.step1.default": "Download Course Template",
    "instructor.action.step1.link_html": '<a href="{href}">{text}</a>',
    "instructor.action.step1.upload": "Upload Course Details",
    "instructor.action.step1.prepare": "Prepare Marks Template",
    "instructor.action.step2.upload.default": "Upload Filled Marks",
    "instructor.action.step2.generate.default": "Generate Final CO Report",
    "instructor.msg.success_title": "Success",
    "instructor.msg.validation_title": "Validation Error",
    "instructor.msg.error_title": "Failed",
    "instructor.msg.step_completed": "Step {step} completed: {title}.",
    "instructor.msg.failed_to_do": "Failed to do {action}.",
    "instructor.dialog.filter.excel": "Excel Files (*.xlsx)",
    "instructor.dialog.filter.excel_open": "Excel Files (*.xlsx *.xlsm *.xls)",
    "instructor.dialog.step1.title": "Save Course Details Template",
    "instructor.dialog.step1.default_name": "course_details_template.xlsx",
    "instructor.dialog.step2.title": "Select Course Details File",
    "instructor.dialog.step1.prepare.title": "Save Marks Template",
    "instructor.dialog.step1.prepare.default_name": "marks_template.xlsx",
    "instructor.dialog.step2.upload.title": "Select Filled Marks File",
    "instructor.dialog.step2.generate.title": "Save Final CO Report",
    "instructor.status.step1_selected": "Template download path selected.",
    "instructor.status.step1_validated": "Course details uploaded and validated. You can now prepare marks template.",
    "instructor.status.step1_prepared": "Marks template generated.",
    "instructor.status.step1_changed": "Course details changed. Redo downstream steps.",
    "instructor.status.step2_uploaded_filled": "Filled marks uploaded.",
    "instructor.status.step2_changed_filled": "Filled marks changed. Regenerate final report.",
    "instructor.status.step2_generated": "Final CO report path selected.",
    "instructor.status.operation_cancelled": "Operation cancelled.",
    "instructor.status.step1_drop_browse_requested": "Course details browse requested from drag-and-drop widget.",
    "instructor.status.step1_drop_files_dropped": "{count} file(s) dropped into course details widget.",
    "instructor.status.step1_drop_files_changed": "Course details widget now has {count} file(s).",
    "instructor.status.step1_drop_files_rejected": "{count} dropped file(s) were rejected by course details widget.",
    "instructor.status.step2_drop_browse_requested": "Filled marks browse requested from drag-and-drop widget.",
    "instructor.status.step2_drop_files_dropped": "{count} file(s) dropped into filled marks widget.",
    "instructor.status.step2_drop_files_changed": "Filled marks widget now has {count} file(s).",
    "instructor.status.step2_drop_files_rejected": "{count} dropped file(s) were rejected by filled marks widget.",
    "instructor.status.step2_validation_warnings": "Validation anomaly warnings: {details}",
    "instructor.status.step2_generate_per_file_failures": "Final report per-file failures: {details}",
    "instructor.status.step1_validating_progress": "Validating course template files: {processed}/{total}.",
    "instructor.status.step1_validated_progress": "Validated valid files: {valid}/{total}.",
    "instructor.status.step1_prepare_progress": "Processed marks templates: {processed}/{total}.",
    "instructor.status.step1_prepare_per_file_failures": "Marks-template per-file failures: {details}",
    "instructor.step1.drop.summary": "Files: {count}",
    "common.dropzone.placeholder": "Drag and Drop, or press Ctrl + O, or single-click to add files",
    "common.validation_failed_invalid_data": "Validation failed due to invalid data.",
    "common.error_while_process": "Error happened while {process}.",
    "instructor.toast.step1_validation_summary": (
        "Validation complete: {valid} valid, {invalid} invalid, {mismatched} wrong template, {duplicates} duplicate input."
    ),
    "instructor.toast.step1_prepare_summary": (
        "Marks template generation complete: processed {processed}/{total}, generated {generated}, failed {failed}, skipped {skipped}."
    ),
    "instructor.toast.step2_generate_summary": (
        "CO report generation complete: processed {processed}/{total}, generated {generated}, failed {failed}, skipped {skipped}."
    ),
    "instructor.toast.step2_upload_reject_summary": (
        "Some files were not accepted. Invalid={invalid}, duplicates={duplicates}."
    ),
    "instructor.toast.validation_warnings_title": "Validation Warnings",
    "instructor.toast.validation_warnings_body": (
        "Validation completed with anomaly warnings. Check activity log details."
    ),
    "instructor.log.title": "Activity Log",
    "instructor.log.ready": "Activity log initialized.",
    "instructor.log.completed_process": "{process} completed successfully.",
    "instructor.log.error_while_process": "Error happened while {process}.",
    "instructor.log.process.generate_course_details_template": "generating course details template",
    "instructor.log.process.validate_course_details_workbook": "validating uploaded course details workbook",
    "instructor.log.process.generate_marks_template": "generating marks template",
    "instructor.log.process.upload_filled_marks_workbook": "uploading filled marks workbook",
    "instructor.log.process.generate_final_co_report": "generating final CO report",
    "instructor.validation.xlsxwriter_missing": "xlsxwriter is not installed. Install it to generate course templates.",
    "instructor.validation.sheet_single_header_row": "Sheet '{sheet_name}' must define exactly one header row.",
    "instructor.system.template_generate_failed": "Failed to generate course details template at '{output}'.",
    "instructor.validation.unknown_template": "Unknown workbook template '{template_id}'. Available templates: {available}.",
    "instructor.validation.invalid_sheet_name": "Invalid sheet name.",
    "instructor.validation.headers_empty": "Headers cannot be empty for sheet '{sheet_name}'.",
    "instructor.validation.headers_unique": "Headers must be unique for sheet '{sheet_name}'.",
    "instructor.validation.row_length_mismatch": "Row {row} length mismatch in '{sheet_name}': expected {expected}, got {found}.",
    "instructor.validation.openpyxl_missing": "openpyxl is not installed. Install it to validate uploaded course details.",
    "instructor.validation.workbook_not_found": "Course details workbook not found: {workbook}",
    "instructor.validation.workbook_open_failed": "Unable to open course details workbook '{workbook}'.",
    "instructor.validation.system_sheet_missing": "Missing required system sheet '{sheet}' in uploaded workbook.",
    "instructor.validation.system_hash_missing_template_id_header": "Invalid system hash sheet format: missing Template_ID header.",
    "instructor.validation.system_hash_missing_template_hash_header": "Invalid system hash sheet format: missing Template_Hash header.",
    "instructor.validation.system_hash_template_id_missing": "Template_ID is missing in system hash sheet.",
    "instructor.validation.system_hash_mismatch": "Template hash mismatch. Please use a valid generated template.",
    "instructor.validation.workbook_sheet_mismatch": "Workbook sheets do not match template '{template_id}'. Expected: {expected}. Found: {found}.",
    "instructor.validation.header_mismatch": "Header mismatch in sheet '{sheet_name}'. Expected: {expected}.",
    "instructor.validation.unexpected_header": "Unexpected header in sheet '{sheet_name}' at column {col}.",
    "instructor.validation.validator_missing": "No validator implemented for template '{template_id}'.",
    "instructor.validation.course_metadata_field_empty": "Course_Metadata row {row}: Field cannot be empty.",
    "instructor.validation.course_metadata_duplicate_field": "Course_Metadata row {row}: duplicate field '{field}'.",
    "instructor.validation.course_metadata_unknown_field": "Course_Metadata row {row}: unknown field '{field}'.",
    "instructor.validation.course_metadata_value_required": "Course_Metadata row {row}: Value is required for '{field}'.",
    "instructor.validation.course_metadata_missing_fields": "Course_Metadata is missing required fields: {fields}.",
    "instructor.validation.course_metadata_field_must_be_int": "Course_Metadata field '{field}' must be an integer.",
    "instructor.validation.course_metadata_field_must_be_non_empty_str": "Course_Metadata field '{field}' must be a non-empty string.",
    "instructor.validation.course_metadata_total_outcomes_invalid": "Course_Metadata field 'Total_Outcomes' must be an integer > 0.",
    "instructor.validation.yes_no_required": "{sheet_name} row {row}: '{field}' must be YES or NO.",
    "instructor.validation.assessment_component_required_one": "Assessment_Config must contain at least one component.",
    "instructor.validation.assessment_component_required": "Assessment_Config row {row}: Component is required.",
    "instructor.validation.assessment_component_duplicate": "Assessment_Config row {row}: duplicate component '{component}'.",
    "instructor.validation.assessment_weight_numeric": "Assessment_Config row {row}: Weight (%) must be numeric.",
    "instructor.validation.assessment_direct_missing": "Assessment_Config must have at least one direct component.",
    "instructor.validation.assessment_indirect_missing": "Assessment_Config must have at least one indirect component.",
    "instructor.validation.assessment_direct_total_invalid": "Assessment_Config direct component weights must total 100. Found: {found}.",
    "instructor.validation.assessment_indirect_total_invalid": "Assessment_Config indirect component weights must total 100. Found: {found}.",
    "instructor.validation.question_map_row_required_one": "Question_Map must contain at least one question row.",
    "instructor.validation.question_component_required": "Question_Map row {row}: Component is required.",
    "instructor.validation.question_component_unknown": "Question_Map row {row}: unknown component '{component}'.",
    "instructor.validation.question_label_required": "Question_Map row {row}: Q_No/Rubric_Parameter is required.",
    "instructor.validation.question_max_marks_numeric": "Question_Map row {row}: Max_Marks must be numeric.",
    "instructor.validation.question_max_marks_positive": "Question_Map row {row}: Max_Marks must be greater than zero.",
    "instructor.validation.question_co_required": "Question_Map row {row}: CO must contain at least one value.",
    "instructor.validation.question_co_no_repeat": "Question_Map row {row}: CO values cannot repeat.",
    "instructor.validation.question_co_out_of_range": "Question_Map row {row}: CO value out of range 1..{total_outcomes}.",
    "instructor.validation.question_co_wise_requires_one": "Question_Map row {row}: component '{component}' requires exactly one CO per question.",
    "instructor.validation.question_duplicate_for_component": "Question_Map row {row}: duplicate question '{question}' for component '{component}'.",
    "instructor.validation.students_row_required_one": "Students sheet must contain at least one student row.",
    "instructor.validation.students_reg_and_name_required": "Students row {row}: both Reg_No and Student_Name are required.",
    "instructor.validation.students_duplicate_reg_no": "Students row {row}: duplicate Reg_No '{reg_no}'.",
    "instructor.validation.step2.layout_sheet_missing": "Missing required layout sheet '{sheet}' in filled marks workbook.",
    "instructor.validation.step2.layout_header_mismatch": "Invalid layout sheet header at {column}. Expected '{expected}'.",
    "instructor.validation.step2.layout_manifest_missing": "Layout manifest or layout hash is missing in system layout sheet.",
    "instructor.validation.step2.layout_hash_mismatch": "Layout hash mismatch. Please use an untampered generated marks template.",
    "instructor.validation.step2.layout_manifest_json_invalid": "Layout manifest JSON is invalid.",
    "instructor.validation.step2.template_validator_missing": "No Step 2 validator available for template '{template_id}'.",
    "instructor.validation.step2.manifest_root_invalid": "Layout manifest root must be an object.",
    "instructor.validation.step2.manifest_structure_invalid": "Layout manifest must contain both 'sheet_order' and 'sheets'.",
    "instructor.validation.step2.sheet_order_mismatch": "Filled marks sheets mismatch. Expected order: {expected}. Found: {found}.",
    "instructor.validation.step2.manifest_sheet_spec_invalid": "Layout manifest contains an invalid sheet specification.",
    "instructor.validation.step2.sheet_missing": "Expected sheet '{sheet_name}' is missing in workbook.",
    "instructor.validation.step2.header_row_invalid": "Sheet '{sheet_name}' has invalid header row metadata: {header_row}.",
    "instructor.validation.step2.headers_missing": "Sheet '{sheet_name}' is missing expected header definitions in layout manifest.",
    "instructor.validation.step2.anchor_spec_invalid": "Sheet '{sheet_name}' has invalid anchor specification in layout manifest.",
    "instructor.validation.step2.formula_anchor_spec_invalid": "Sheet '{sheet_name}' has invalid formula anchor specification in layout manifest.",
    "instructor.validation.step2.header_row_mismatch": "Sheet '{sheet_name}' header mismatch at row {row}. Expected: {expected}.",
    "instructor.validation.step2.anchor_value_mismatch": "Sheet '{sheet_name}' cell {cell} was modified. Expected '{expected}', found '{found}'.",
    "instructor.validation.step2.formula_mismatch": "Sheet '{sheet_name}' formula at {cell} was modified.",
    "instructor.validation.step2.no_component_sheets": "No component mark-entry sheets found in filled marks workbook.",
    "instructor.validation.step2.mark_entry_empty": "Sheet '{sheet_name}' has an empty mark-entry cell at {cell}. Fill all marks before upload.",
    "instructor.validation.step2.mark_value_invalid": "Sheet '{sheet_name}' cell {cell} has invalid mark value '{value}'. Allowed: A/a or numeric between {minimum} and {maximum}.",
    "instructor.validation.step2.mark_precision_invalid": "Sheet '{sheet_name}' cell {cell} has too many decimal places in '{value}'. Maximum allowed is {decimals}.",
    "instructor.validation.step2.indirect_mark_must_be_integer": "Sheet '{sheet_name}' cell {cell} has invalid indirect mark '{value}'. Use an integer Likert value.",
    "instructor.validation.step2.absence_policy_violation": "Sheet '{sheet_name}' row {row} has mixed absence and numeric entries in {range}. Enter either all A/a or all numeric marks.",
    "instructor.validation.step2.total_formula_mismatch": "Sheet '{sheet_name}' total formula was modified at {cell}.",
    "instructor.validation.step2.co_formula_mismatch": "Sheet '{sheet_name}' CO split formula was modified at {cell}.",
    "instructor.validation.step2.structure_snapshot_missing": "Sheet '{sheet_name}' is missing mark-structure metadata in layout manifest.",
    "instructor.validation.step2.structure_snapshot_mismatch": "Sheet '{sheet_name}' mark-structure cell {cell} was modified.",
    "instructor.validation.step2.student_identity_spec_invalid": "Sheet '{sheet_name}' has invalid student identity metadata in layout manifest.",
    "instructor.validation.step2.student_identity_mismatch": "Sheet '{sheet_name}' student Reg. No./Name rows were modified.",
    "instructor.validation.step2.student_reg_duplicate": "Sheet '{sheet_name}' has duplicate student Reg. No. '{reg_no}'.",
    "instructor.validation.step2.student_identity_cross_sheet_mismatch": "Student Reg. No./Name rows in sheet '{sheet_name}' do not match sheet '{reference_sheet}'.",
    "instructor.validation.final_report.layout_manifest_invalid": "Final report generation failed: layout manifest is invalid.",
    "instructor.validation.final_report.direct_component_sheet_missing": "Final report generation failed: missing direct component sheet for '{component}'.",
    "instructor.validation.final_report.direct_component_marks_shape_invalid": "Final report generation failed: invalid marks shape for component '{component}'.",
    "instructor.validation.final_report.no_direct_components": "Final report generation failed: no direct components found.",
    # Instructor template headers / labels
    # Coordinator module
    "coordinator.title": "Course Attainment",
    "coordinator.drop_hint": "Upload Final CO report workbooks generated from the Instructor module.",
    "coordinator.calculate": "Calculate CO Attainment",
    "coordinator.thresholds.title": "CO Attainment Thresholds",
    "coordinator.thresholds.description": (
        "- L1 Threshold - pass mark of the course or course average for the 3 batches offered in the previous regulation.\n"
        "- L2 Threshold - First Class.\n"
        "- L3 Threshold - Distinction Class."
    ),
    "coordinator.thresholds.l1.label": "L1 Threshold:",
    "coordinator.thresholds.l2.label": "L2 Threshold:",
    "coordinator.thresholds.l3.label": "L3 Threshold:",
    "coordinator.thresholds.invalid_rule": "Invalid thresholds: 0 < L1 < L2 < L3 < 100 is required.",
    "coordinator.co_attainment.description": "Attained if the following % of students are at level fixed and above.",
    "coordinator.co_attainment.percent.label": "CO AT%:",
    "coordinator.co_attainment.level.label": "CO AT Level:",
    "coordinator.co_attainment.invalid_percent": "Invalid CO AT%: value must be between 0 and 100.",
    "coordinator.dialog.title": "Select Excel Files",
    "coordinator.links.downloaded_output": "Downloaded Output",
    "coordinator.summary": "Files: {count}",
    "coordinator.file.remove_fallback": "Remove",
    "coordinator.file.remove_tooltip": "Remove File",
    "coordinator.clear_all": "Clear All",
    "coordinator.status.added": "{added} file(s) added. Total: {total}.",
    "coordinator.duplicate.title": "File Already Exists",
    "coordinator.duplicate.body": "{count} duplicate file(s) were skipped because they already exist in the list.",
    "coordinator.invalid_final_report.title": "Invalid Final CO Report",
    "coordinator.invalid_final_report.body": (
        "Only Final CO report workbooks generated from Instructor Step 2 are allowed.\n\n{files}"
    ),
    "coordinator.invalid_final_report.details_prefix": "Details:",
    "coordinator.invalid_final_report.detail_line": "{file}: {reason}",
    "coordinator.status.ignored": (
        "{count} file(s) were ignored (unsupported type, missing, duplicate, "
        "or invalid final report workbook)."
    ),
    "coordinator.status.removed": "{count} file(s) removed.",
    "coordinator.status.cleared": "{count} file(s) cleared.",
    "coordinator.status.processing_started": "Processing selected files...",
    "coordinator.status.queued": "{count} file(s) queued while current processing is running.",
    "coordinator.status.operation_cancelled": "Coordinator processing cancelled.",
    "coordinator.status.processing_failed": "Coordinator processing failed.",
    "coordinator.status.processing_completed": "Coordinator processing completed.",
    "coordinator.status.calculate_completed": "CO attainment calculation completed and output workbook generated.",
    "coordinator.regno_dedup.title": "Duplicate Register Numbers Removed",
    "coordinator.regno_dedup.body": (
        "{count} duplicate register number entr(ies) were removed while generating CO attainment."
    ),
    "coordinator.regno_dedup.log_body": (
        "{count} duplicate register number entr(ies) were removed:\n{details}"
    ),
    "coordinator.regno_dedup.log_detail": (
        "Reg No: {reg_no} | Worksheet: {worksheet} | Workbook: {workbook}"
    ),
    "coordinator.regno_dedup.log_detail_unavailable": "Duplicate entry details are unavailable.",
    "coordinator.join_drop.body": (
        "Some rows were ignored because they existed only in Direct or only in Indirect sheets. "
        "Dropped rows: {count}. See activity log for details."
    ),
    "co_analysis.status.duplicate_reg_numbers": "{count} file(s) rejected due to duplicate register numbers.",
    "co_analysis.status.co_count_mismatch": "{count} file(s) rejected due to Total Outcomes mismatch.",
    "co_analysis.status.validation_warnings": "{count} validation warning(s) reported.",
    "co_analysis.status.rejected_code_with_context": "Rejected {file}: {code} ({context})",
    "co_analysis.validation.co_count_mismatch_body": (
        "{count} file(s) were rejected because Total Outcomes differs from the selected batch."
    ),
    "co_analysis.validation.anomaly_warnings_body": (
        "{count} validation anomaly warning(s) were reported in uploaded files."
    ),
    "co_analysis.validation.rejection_breakdown_title": "CO Analysis File Rejections",
    "co_analysis.validation.rejection_breakdown_body": (
        "Rejected files by reason:\n"
        "Unsupported/Missing: {unsupported_or_missing}\n"
        "Invalid Workbook: {invalid_workbook}\n"
        "System Hash/Identity Failures: {invalid_hash}\n"
        "Marks Unfilled: {marks_unfilled}\n"
        "Layout/Manifest Failures: {layout_manifest}\n"
        "Template Mismatch: {template_mismatch}\n"
        "Invalid Mark Value/Format: {mark_value}\n"
        "Other Validation Failures: {invalid_other}\n"
        "Duplicates: {duplicates}\n"
        "Duplicate Register Numbers: {duplicate_reg}\n"
        "Total Outcomes Mismatch: {co_count_mismatch}"
    ),
    "co_analysis.status.ignored_summary": (
        "{count} file(s) ignored. "
        "{part1}{sep1_2}{part2}{sep2_3}{part3}{sep3_4}{part4}{sep4_5}{part5}{sep5_6}{part6}"
        "{sep6_7}{part7}{sep7_8}{part8}{sep8_9}{part9}{sep9_10}{part10}{sep10_11}{part11}{sep11_12}{part12}"
    ),
    "co_analysis.status.ignored_total": "{count} file(s) ignored.",
    "co_analysis.status.ignored_reason.unsupported_or_missing": "Unsupported/Missing={count}",
    "co_analysis.status.ignored_reason.invalid_workbook": "Invalid Workbook={count}",
    "co_analysis.status.ignored_reason.invalid_hash": "Hash={count}",
    "co_analysis.status.ignored_reason.marks_unfilled": "Unfilled={count}",
    "co_analysis.status.ignored_reason.layout_manifest": "Layout/Manifest={count}",
    "co_analysis.status.ignored_reason.template_mismatch": "Template Mismatch={count}",
    "co_analysis.status.ignored_reason.mark_value": "Mark Value={count}",
    "co_analysis.status.ignored_reason.invalid_other": "Other Validation={count}",
    "co_analysis.status.ignored_reason.duplicates": "Duplicates={count}",
    "co_analysis.status.ignored_reason.duplicate_reg": "Duplicate Reg No={count}",
    "co_analysis.status.ignored_reason.co_count_mismatch": "Total Outcomes Mismatch={count}",
}



